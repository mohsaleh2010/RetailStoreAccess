"""Journal entries (modJournal, frmJournal / frmJournalEntry, qryJournal*):
  * the fixture business is synchronised on the SQLite mirror with the same steps as
    SyncJournal (tools/sim.py): one balanced entry per operation, linked to its origin,
    numbered in date order; the trial balance agrees with the customer, supplier,
    stock and cash balances; a changed operation keeps its entry number, a deleted
    one loses its entry, a second sync changes nothing;
  * the VBA is pinned to those steps; every kind of operation can be opened;
  * schema upgrade safety, screens, report centre and navigation wiring."""
import os
import re
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools"))
sys.path.insert(0, HERE)

import forms as F
import queries as Q
from access_sqlite import AccessOnSqlite, day
from helpers import VbaModuleChecks
from schema import table
from sim import Store


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


FORMS = {m.name: m for m in F.all_forms()}


class JournalSyncTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.s = Store(self.db.con, now=day(0, 9))
        self.first = self.s.sync_journal()

    def one(self, sql, *args):
        return self.s.one(sql, *args)

    def balance(self, code):
        return round(self.one("SELECT Sum(Debit) - Sum(Credit) FROM JournalLines WHERE AccountCode = ?", code) or 0, 2)

    def test_one_entry_per_operation(self):
        ids = self.db.ids
        for stype, key in [("SALE", "INV0"), ("SALE", "INV1"), ("SALE", "INV2"), ("SALES_RETURN", "CRN1"),
                           ("PURCHASE", "PUR1"), ("PURCHASE_RETURN", "PRT1"), ("CUSTOMER_PAYMENT", "RCV1"),
                           ("SUPPLIER_PAYMENT", "PAY1"), ("EXPENSE", "EXP1"), ("EXPENSE", "EXP2"),
                           ("CASH_VOUCHER", "V1"), ("CASH_VOUCHER", "V3"), ("STOCK_MOVE", "T1"),
                           ("STOCK_MOVE", "T10"), ("BOX_OPENING", "BOXM"), ("CUSTOMER_OPENING", "C2")]:
            self.assertEqual(self.one("SELECT COUNT(*) FROM JournalEntries WHERE SourceType = ? AND SourceID = ?",
                                      stype, ids[key]), 1, (stype, key))
        self.assertEqual(self.one("SELECT COUNT(*) FROM JournalEntries WHERE SourceType = 'EXPENSE' AND SourceID = ?",
                                  ids["EXPV"]), 0, "booked by its cash voucher")
        self.assertEqual(self.first[0], self.one("SELECT COUNT(*) FROM JournalEntries"))

    def test_entries_are_balanced_and_complete(self):
        self.assertEqual(self.one("SELECT COUNT(*) FROM JournalEntries WHERE abs(TotalDebit - TotalCredit) > 0.001"), 0)
        self.assertEqual(self.one("SELECT COUNT(*) FROM JournalEntries AS e WHERE abs(e.TotalDebit - "
                                  "(SELECT Sum(Debit) FROM JournalLines AS l WHERE l.EntryID = e.EntryID)) > 0.001 "
                                  "OR e.LineCount <> (SELECT COUNT(*) FROM JournalLines AS l WHERE l.EntryID = e.EntryID)"),
                         0)
        self.assertEqual(self.one("SELECT COUNT(*) FROM JournalLines AS l LEFT JOIN Accounts AS a "
                                  "ON l.AccountCode = a.AccountCode WHERE a.AccountCode Is Null"), 0)

    def test_accounts_agree_with_the_balances(self):
        ids = self.db.ids
        self.assertEqual(self.balance(1300), 164)                 # CustomerBalanceQuery
        self.assertEqual(self.balance(2100), -2855)               # SupplierBalanceQuery (credit)
        self.assertEqual(self.balance(1400), 7040)                # stock at cost
        self.assertEqual(self.balance(110000 + ids["BOXC"]), 770)
        self.assertEqual(self.balance(110000 + ids["BOXM"]), 7700)
        self.assertEqual(self.one("SELECT AccountName FROM Accounts WHERE AccountCode = ?", 110000 + ids["BOXC"]),
                         "TEST صندوق كاشير")

    def test_trial_balance(self):
        db = self.db
        db.set_period(30)
        self.assertEqual(round(db.scalar("SELECT Sum(ClosingBalance) FROM TrialBalanceQuery"), 4), 0)
        self.assertEqual(round(db.scalar("SELECT Sum(PeriodDebit) - Sum(PeriodCredit) FROM TrialBalanceQuery"), 4), 0)
        self.assertEqual(db.scalar("SELECT ClosingBalance FROM TrialBalanceQuery WHERE AccountCode = 1300"), 164)
        self.assertEqual(db.scalar("SELECT OpeningBalance FROM TrialBalanceQuery WHERE AccountCode = 1300"), 50)
        self.assertEqual(db.scalar("SELECT ClosingBalance FROM TrialBalanceQuery WHERE AccountCode = 4100"), -1500)
        self.assertGreater(db.scalar("SELECT COUNT(*) FROM JournalLinesQuery"), 0)

    def test_numbers_follow_the_dates(self):
        dates = [r[0] for r in self.db.con.execute("SELECT EntryDate FROM JournalEntries ORDER BY EntryNumber")]
        self.assertEqual(dates, sorted(dates))
        self.assertEqual(self.one("SELECT Min(EntryNumber) FROM JournalEntries"), "JV-000001")

    def test_second_sync_changes_nothing(self):
        self.assertEqual(self.s.sync_journal(), (0, 0, 0))

    def test_changed_new_and_deleted_operations(self):
        ids, c = self.db.ids, self.db.con
        number = self.one("SELECT EntryNumber FROM JournalEntries WHERE SourceType = 'EXPENSE' AND SourceID = ?",
                          ids["EXP2"])
        c.execute("UPDATE Expenses SET ExpenseTypeID = 6 WHERE ExpenseID = ?", (ids["EXP2"],))     # same amount
        self.assertEqual(self.s.sync_journal(), (0, 1, 0), "another account is a change too")
        self.assertEqual(self.one("SELECT EntryNumber FROM JournalEntries WHERE SourceType = 'EXPENSE' "
                                  "AND SourceID = ?", ids["EXP2"]), number)
        self.assertEqual(self.balance(530006), 1000)
        c.execute("DELETE FROM Expenses WHERE ExpenseID = ?", (ids["EXP2"],))
        new = self.s.expense(2, 10, 0, "TEST جديد")
        self.assertEqual(self.s.sync_journal(), (1, 0, 1))
        self.assertEqual(self.one("SELECT COUNT(*) FROM JournalEntries WHERE SourceType = 'EXPENSE' AND SourceID = ?",
                                  ids["EXP2"]), 0)
        self.assertEqual(self.one("SELECT EntryNumber FROM JournalEntries WHERE SourceType = 'EXPENSE' "
                                  "AND SourceID = ?", new), self.one("SELECT Max(EntryNumber) FROM JournalEntries"))


class JournalCodeTests(unittest.TestCase):

    def test_sources_and_steps_match_the_replay(self):
        text = read("modJournal")
        listed = re.search(r'SOURCE_QUERIES As String = ((?:"[^"]*"(?: & _\s+)?)+)', text).group(1)
        self.assertEqual("".join(re.findall(r'"([^"]*)"', listed)).split(","), Q.JOURNAL_SOURCE_QUERIES)
        sync = proc(text, "SyncJournal")
        for part in ["Sum(AccountCode * (Debit + Debit + Credit)) + Sum(CostCenter * (Debit + Debit + Credit) * 7) AS Sig", "Max(Party) AS FirstText",
                     "GROUP BY SourceType, SourceID", 'e!EntryNumber = "~"', "WHERE e.LineCount < 0",
                     "UPDATE JournalEntries SET LineCount = -LineCount WHERE LineCount < 0",
                     "EntryNumber Like '~*'", "ORDER BY EntryDate, EntryID", 'NextNumber("JOURNAL")',
                     "If Not seen.Exists(k) Then", "ws.BeginTrans", "ws.Rollback"]:
            self.assertIn(part, sync)
        self.assertIn("e!LineCount = -rs!LineTotal", proc(text, "FillHeader"))
        ensure = proc(text, "EnsureAccounts")
        self.assertIn("110000 + b.CashBoxID, b.BoxName, 'ASSET', 1100", ensure)
        self.assertIn("530000 + t.ExpenseTypeID, t.ExpenseTypeName, 'EXPENSE', 5300", ensure)

    def test_every_kind_of_operation_can_be_opened(self):
        body = proc(read("modJournal"), "OpenJournalSource")
        for kind, *_ in table("JournalSourceTypes").seed_rows:
            self.assertIn(f'Case "{kind}":', body)
            self.assertIn(f"'{kind}'", " ".join(q.sql for q in Q.QUERIES if q.name in Q.JOURNAL_SOURCE_QUERIES))

    def test_in_access_test_and_demo(self):
        self.assertIn('"TestJournal"', read("modTestAll"))
        demo = read("modDemoData")
        self.assertIn('Check SyncJournal(), "قيود اليومية"', demo)
        self.assertIn("FROM JournalEntries WHERE TotalDebit <> TotalCredit", demo)


class JournalSchemaTests(unittest.TestCase):

    def test_upgrade_adds_accounts_kinds_and_permission(self):
        self.assertTrue(table("Accounts").seed_missing and table("JournalSourceTypes").seed_missing)
        schema = read("modBuildSchema")
        self.assertIn('GrantNewPermission "JOURNAL", "1,2"', schema)
        self.assertIn("SeedRow \"[SequenceName] = 'JOURNAL'\"", schema)
        codes = {r[0] for r in table("Accounts").seed_rows}
        used = set()
        for q in Q.JOURNAL_SOURCE_QUERIES:
            used |= {int(n) for n in re.findall(r"\b(\d{4})\b", Q.query(q).sql)}
        self.assertLessEqual(used - {2000}, codes, "every fixed account used by the rules is seeded")

    def test_entry_is_linked_once_to_its_origin(self):
        t = table("JournalEntries")
        self.assertIn(["SourceType", "SourceID"], [ix.fields for ix in t.indexes if ix.unique])
        self.assertEqual(t.rule, "[TotalDebit]=[TotalCredit]")


class JournalScreenTests(unittest.TestCase):

    def test_controls_used_by_the_code_exist(self):
        text = read("modJournal")
        for pname in re.findall(r"^Public Sub (\w+)\(ByVal frm As Access\.Form", text, re.M):
            form = "frmJournalEntry" if pname == "JournalEntryLoad" else "frmJournal"
            names = {c.name for c in FORMS[form].controls}
            for ctl in set(re.findall(r"frm!(\w+)", proc(text, pname))):
                self.assertIn(ctl, names, f"{form}: {pname} uses {ctl}")

    def test_navigation_permissions_and_reports(self):
        nav = {n.key: n.target for n in F.NAV_ITEMS}             # the journal is a tile of the accounting hub
        self.assertEqual(nav["Accounting"], "frmAccounting")
        import forms_accounting as A
        self.assertIn("frmJournal", {t[3] for t in A.HUB_TILES})
        for form in ("frmJournal", "frmJournalEntry", "frmAccounts"):
            self.assertEqual(F.SCREEN_PERMISSIONS[form], "JOURNAL")
        keys = {r.key: r for r in F.REPORTS}
        self.assertIn("J", keys["JOURNAL"].needs)
        self.assertIn("J", keys["TRIAL_BALANCE"].needs)
        self.assertIn('If HasNeed(needs, "J") Then', read("modScreens"))


class Static_modJournal(VbaModuleChecks, unittest.TestCase):
    module_name = "modJournal"
    vba = read("modJournal")


if __name__ == "__main__":
    unittest.main()
