"""Chart of accounts tree (schema.ACCOUNT_TREE, modAccounts) and manual journal entries
(ManualEntries, modManualEntry, frmManualEntry):
  * the seeded tree: five classes, every account under a main account of its own type, at most
    five levels, the automatic entries post only to sub-accounts that are system accounts;
  * the tree is rebuilt on the SQLite mirror with the same steps (tools/sim.py) and the trial
    balance by levels adds up: level-1 totals balance, each main account = its sub-accounts;
  * the fixture manual entry becomes a journal entry; changed -> same number; deleted -> removed;
  * the VBA keeps the rules (balanced, sub-accounts only, permission, transaction) and the screen
    controls used by the code exist."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day
import forms as F
import queries as Q
from schema import ACCOUNT_TREE, table
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}
TREE = {r[0]: r for r in ACCOUNT_TREE}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


def rule_accounts():
    used = set()
    for q in Q.JOURNAL_SOURCE_QUERIES:
        used |= {int(n) for n in re.findall(r"\b(\d{4})\b", Q.query(q).sql)}
    return used - {2000}


class SeededTreeTests(unittest.TestCase):

    def test_five_classes(self):
        roots = [r for r in ACCOUNT_TREE if r[3] is None]
        self.assertEqual([(r[0], r[2]) for r in roots],
                         [(1, "ASSET"), (2, "LIABILITY"), (3, "EQUITY"), (4, "REVENUE"), (5, "EXPENSE")])

    def test_every_account_under_a_main_account_of_its_type(self):
        for code, name, kind, parent, posting, system in ACCOUNT_TREE:
            if parent is None:
                continue
            self.assertIn(parent, TREE, code)
            self.assertFalse(TREE[parent][4], f"{code}: parent {parent} must be a main account")
            self.assertEqual(TREE[parent][2], kind, code)

    def test_levels(self):
        def level(code):
            return 1 if TREE[code][3] is None else 1 + level(TREE[code][3])
        self.assertLessEqual(max(level(c) for c in TREE), 5)
        self.assertEqual({c for c in TREE if level(c) == 2}, {11, 12, 21, 22, 31, 32, 41, 42, 51, 52, 53})

    def test_automatic_entries_post_to_system_sub_accounts(self):
        for code in rule_accounts():
            self.assertIn(code, TREE, code)
            self.assertTrue(TREE[code][4], f"{code} must accept entries")
            self.assertTrue(TREE[code][5], f"{code} must be a system account")
        self.assertFalse(TREE[1100][4])
        self.assertFalse(TREE[5300][4])
        self.assertIn("110000 + b.CashBoxID, b.BoxName, 'ASSET', 1100, True, True", read("modJournal"))
        self.assertIn("530000 + t.ExpenseTypeID, t.ExpenseTypeName, 'EXPENSE', 5300, True, True", read("modJournal"))

    def test_older_back_ends_are_upgraded(self):
        body = proc(read("modBuildSchema"), "UpgradeAccountTree")
        for code, name, kind, parent, posting, system in ACCOUNT_TREE:
            if parent is not None:
                self.assertIn(f"SET [ParentCode] = {parent} WHERE [AccountCode] = {code} AND [ParentCode] Is Null", body)
        self.assertIn("[IsPosting] = False WHERE [AccountCode] IN (1, 11, 1100,", body)
        self.assertIn("BETWEEN 530001 AND 539999", body)
        build = proc(read("modBuildSchema"), "BuildSchema")
        self.assertLess(build.index("UpgradeAccountTree"), build.index('Application.Run "RebuildAccountTree"'))
        self.assertTrue(table("Accounts").seed_missing)


class TreeOnMirrorTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.s = Store(self.db.con, now=day(0, 9))
        self.s.sync_journal()
        self.db.set_period(30)

    def one(self, sql, *args):
        return self.s.one(sql, *args)

    def test_every_account_placed(self):
        self.assertEqual(self.s.rebuild_account_tree(), [])
        self.assertEqual(self.one("SELECT COUNT(*) FROM Accounts WHERE AccountLevel IS NULL OR TreeKey IS NULL"), 0)
        box = 110000 + self.db.ids["BOXC"]
        self.assertEqual(self.one("SELECT AccountLevel FROM Accounts WHERE AccountCode = ?", box), 4)
        self.assertEqual(self.one("SELECT Level2Code FROM Accounts WHERE AccountCode = ?", box), 11)
        self.assertEqual(self.one("SELECT TreeKey FROM Accounts WHERE AccountCode = 1300"),
                         f"{1:010d}{11:010d}{1300:010d}")
        keys = [r[0] for r in self.db.con.execute("SELECT TreeKey FROM AccountTreeQuery")]
        self.assertEqual(keys, sorted(keys))

    def test_loops_and_depth_are_reported(self):
        c = self.db.con
        c.execute("UPDATE Accounts SET ParentCode = 1300 WHERE AccountCode = 11")    # 11 -> 1300 -> 11
        self.assertIn(11, self.s.rebuild_account_tree())

    def test_trial_balance_by_levels(self):
        scalar = self.db.scalar
        self.assertEqual(round(scalar("SELECT Sum(ClosingBalance) FROM TrialBalanceTreeQuery WHERE AccountLevel = 1"), 4), 0)
        self.assertEqual(round(scalar("SELECT Sum(PeriodDebit) - Sum(PeriodCredit) FROM TrialBalanceTreeQuery "
                                      "WHERE AccountLevel = 1"), 4), 0)
        posting = "SELECT Sum(t.ClosingBalance) FROM TrialBalanceQuery AS t INNER JOIN Accounts AS a ON " \
                  "t.AccountCode = a.AccountCode WHERE "
        for code, where in [(1, "a.AccountType = 'ASSET'"), (21, "a.Level2Code = 21"), (1100, "a.ParentCode = 1100"),
                            (5, "a.AccountType = 'EXPENSE'")]:
            self.assertAlmostEqual(scalar(f"SELECT ClosingBalance FROM TrialBalanceTreeQuery WHERE AccountCode = {code}"),
                                   scalar(posting + where), places=4, msg=code)
        self.assertEqual(scalar("SELECT ClosingBalance FROM TrialBalanceTreeQuery WHERE AccountCode = 1300"), 164)
        self.assertEqual(scalar("SELECT COUNT(*) FROM TrialBalanceTreeQuery WHERE AccountCode = 1250"), 0,
                         "accounts without entries are not listed")

    def test_manual_entry_in_the_journal(self):
        ids = self.db.ids
        e = self.one("SELECT EntryID FROM JournalEntries WHERE SourceType = 'MANUAL' AND SourceID = ?", ids["MJ1"])
        self.assertIsNotNone(e)
        self.assertEqual(self.one("SELECT TotalDebit FROM JournalEntries WHERE EntryID = ?", e), 3000)
        self.assertEqual(self.one("SELECT COUNT(*) FROM JournalLines WHERE EntryID = ?", e), 3)
        self.assertEqual(self.one("SELECT Credit FROM JournalLines WHERE EntryID = ? AND AccountCode = 2310", e), 2500)
        self.assertIn("TEST رواتب الشهر المستحقة", self.one("SELECT Description FROM JournalEntries WHERE EntryID = ?", e))
        number = self.one("SELECT EntryNumber FROM JournalEntries WHERE EntryID = ?", e)
        c = self.db.con
        c.execute("UPDATE ManualEntryLines SET Debit = 3500 WHERE ManualEntryID = ? AND LineNumber = 1", (ids["MJ1"],))
        c.execute("UPDATE ManualEntryLines SET Credit = 3000 WHERE ManualEntryID = ? AND LineNumber = 2", (ids["MJ1"],))
        self.assertEqual(self.s.sync_journal(), (0, 1, 0))
        self.assertEqual(self.one("SELECT EntryNumber FROM JournalEntries WHERE EntryID = ?", e), number)
        self.assertEqual(self.one("SELECT TotalDebit FROM JournalEntries WHERE EntryID = ?", e), 3500)
        c.execute("DELETE FROM ManualEntries WHERE ManualEntryID = ?", (ids["MJ1"],))
        self.assertEqual(self.one("SELECT COUNT(*) FROM ManualEntryLines WHERE ManualEntryID = ?", ids["MJ1"]), 0)
        self.assertEqual(self.s.sync_journal(), (0, 0, 1))

    def test_line_rule(self):
        t = table("ManualEntryLines")
        self.assertEqual(t.rule, "([Debit]=0 Or [Credit]=0) And [Debit]+[Credit]>0")


class ManualEntryCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modManualEntry")

    def test_rules(self):
        body = proc(self.text, "ManualEntryProblem")
        for part in ("If n < 2 Then", "ElseIf debit <> credit Then", "AND IsPosting = True ", "AND IsActive = True",
                     "Nz(rs!Debit, 0) > 0 And Nz(rs!Credit, 0) > 0"):
            self.assertIn(part, body)
        post = proc(self.text, "PostManualEntry")
        self.assertLess(post.index('HasPermission("MANUAL_ENTRY")'), post.index("ws.BeginTrans"))
        self.assertLess(post.index("ManualEntryProblem(EntryDate, Description)"), post.index("ws.BeginTrans"))
        self.assertIn('NextNumber("MANUAL_ENTRY")', post)
        self.assertIn("ws.Rollback", post)
        save = proc(self.text, "SaveManualEntry")
        self.assertIn('CanScreenAction(frm.Name, IIf(id = 0, "ADD", "EDIT"))', save)
        self.assertLess(save.index("PostManualEntry("), save.index("SyncJournal()"))
        self.assertIn('CanScreenAction(frm.Name, "DELETE")', proc(self.text, "DeleteManualEntry"))
        self.assertIn('HasPermission("MANUAL_ENTRY")', proc(self.text, "RemoveManualEntry"))

    def test_reverse_swaps_debit_and_credit(self):
        body = proc(self.text, "ReverseManualEntry")
        # the target columns are swapped: (AccountCode, Credit, Debit, ...) SELECT AccountCode, <debit>, <credit>
        self.assertIn('"INSERT INTO tmpManualLines (AccountCode, Credit, Debit, LineText, LineCenter) SELECT AccountCode, " & _\n'
                      '                      LINE_AMOUNTS & ", LineText, CostCenterID FROM ManualEntryLines', body)
        self.assertIn("frm!txtEntryID.Value = Null", body)

    def test_opened_from_the_journal(self):
        self.assertIn('Case "MANUAL":           OpenScreen "frmManualEntry", 0, id',
                      proc(read("modJournal"), "OpenJournalSource"))
        self.assertIn(("MANUAL", "قيد يدوي", 14), table("JournalSourceTypes").seed_rows)

    def test_screen_controls_exist(self):
        names = {c.name for c in MODELS["frmManualEntry"].controls}
        for pname in re.findall(r"^(?:Public|Private) (?:Sub|Function) (\w+)\(ByVal frm As Access\.Form", self.text, re.M):
            for ctl in set(re.findall(r"frm!(\w+)", proc(self.text, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        lines = {c.name for c in MODELS["frmManualLines"].controls}
        for pname in ("ManualLineChanged", "ManualRemoveLine"):
            for ctl in set(re.findall(r"sf!(\w+)", proc(self.text, pname))):
                self.assertIn(ctl, lines, f"{pname}: {ctl}")
        tmp = re.search(r'CREATE TABLE tmpManualLines \((.*?)\)", dbFailOnError', read("modPOS"), re.S).group(1)
        for col in ("LineNo", "AccountCode", "Debit", "Credit", "LineText"):
            self.assertIn(col, tmp)

    def test_permission_and_screens(self):
        self.assertIn(("MANUAL_ENTRY", "MJ-", 1, 6, "القيود اليدوية"), table("Sequences").seed_rows)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmManualEntry"], "MANUAL_ENTRY")
        self.assertIn('GrantNewPermission "MANUAL_ENTRY", "1,2"', read("modBuildSchema"))
        keys = {r.key for r in F.REPORTS}
        self.assertLessEqual({"TRIAL_BALANCE_TREE", "ACCOUNT_TREE"}, keys)
        self.assertIn("Case \"Accounts\"\n            If Not ValidateAccount(frm) Then Exit Function", read("modForms"))

    def test_in_access_tests(self):
        body = proc(read("modJournal"), "TestJournal")
        for part in ("قيد يدوي غير متوازن يُرفض", "قيد يدوي على حساب رئيسي يُرفض", "تعديل القيد اليدوي يحدّث قيده بنفس الرقم",
                     "حذف القيد اليدوي يحذف قيده", "RebuildAccountTree()"):
            self.assertIn(part, body)


class AccountRuleTests(unittest.TestCase):

    def test_validate_account(self):
        body = proc(read("modAccounts"), "ValidateAccount")
        for part in ("لا يتغير رقم حساب محفوظ", "اختر الحساب الرئيسي", "الحساب الرئيسي يجب أن يكون حسابًا تجميعيًا",
                     "لا يتبع الحساب حسابًا من الحسابات التابعة له", "MAX_ACCOUNT_LEVELS",
                     "هذا حساب أساسي تستخدمه القيود الآلية", "AccountHasEntries(code)", "AccountHasChildren(code)"):
            self.assertIn(part, body)
        self.assertIn("AccountDeleteProblem(CLng(id))", read("modForms"))

    def test_rebuild_matches_the_replay(self):
        body = proc(read("modAccounts"), "RebuildAccountTree")
        self.assertIn('Format$(chain(n - k + 1), "0000000000")', body)
        self.assertIn("If code <> 0 And Not parents.Exists(code) Then code = 0", body)
        self.assertIn('rs("Level" & k & "Code") = chain(n - k + 1)', body)


class Static_modAccounts(VbaModuleChecks, unittest.TestCase):
    module_name = "modAccounts"
    vba = read("modAccounts")


class Static_modManualEntry(VbaModuleChecks, unittest.TestCase):
    module_name = "modManualEntry"
    vba = read("modManualEntry")


if __name__ == "__main__":
    unittest.main()
