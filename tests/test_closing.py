"""Period and fiscal year closing (modClosing, frmPeriodClosing, Settings.ClosedThrough,
FiscalYearClosings, journal source YEAR_CLOSE) on the SQLite mirror:
  * the journal never changes in a closed period: a changed, deleted or new operation dated
    there keeps / gets no entry and is reported; operations after it still update;
  * the year closing (same steps as CloseFiscalYear, tools/sim.py): revenue and expense back to
    zero, net profit in retained earnings 3300, the income statement still shows the year's
    profit, the balance sheet stays balanced with the profit in retained earnings;
  * the VBA keeps the rules: dates checked on every screen that can write in the past, years in
    order, only after their end, reopening the last one with a reason; the screen and permission."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day
import forms as F
import queries as Q
from schema import table
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


class ClosedPeriodJournalTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.s = Store(self.db.con, now=day(0, 9))
        self.s.sync_journal()
        self.c = self.db.con
        self.c.execute("UPDATE Settings SET ClosedThrough = ? WHERE SettingID = 1", (day(20),))
        self.ids = self.db.ids

    def entry(self, kind, key):
        return self.s.one("SELECT TotalDebit FROM JournalEntries WHERE SourceType = ? AND SourceID = ?",
                          kind, self.ids[key])

    def test_changed_operation_in_closed_period_keeps_its_entry(self):
        self.c.execute("UPDATE Expenses SET Amount = 999, TotalAmount = 999, Tax = 0 WHERE ExpenseID = ?",
                       (self.ids["EXP2"],))                        # 40 days ago: closed
        self.assertEqual(self.s.sync_journal(), (0, 0, 0))
        self.assertIn(("EXPENSE", self.ids["EXP2"]), self.s.locked)
        self.assertEqual(self.entry("EXPENSE", "EXP2"), 1000)

    def test_deleted_operation_in_closed_period_keeps_its_entry(self):
        self.c.execute("DELETE FROM Expenses WHERE ExpenseID = ?", (self.ids["EXP2"],))
        self.assertEqual(self.s.sync_journal(), (0, 0, 0))
        self.assertEqual(self.entry("EXPENSE", "EXP2"), 1000)

    def test_open_period_still_updates(self):
        self.c.execute("UPDATE Expenses SET Amount = 100, Tax = 15, TotalAmount = 115 WHERE ExpenseID = ?",
                       (self.ids["EXP1"],))                        # 6 days ago: open
        self.assertEqual(self.s.sync_journal(), (0, 1, 0))
        self.assertEqual(self.s.locked, [])
        self.assertEqual(self.entry("EXPENSE", "EXP1"), 115)

    def test_new_operation_in_closed_period_gets_no_entry(self):
        mid = self.s.insert("ManualEntries", EntryNumber="TEST-MJ-OLD", EntryDate=day(25), Description="TEST قديم",
                            TotalAmount=10, EmployeeID=1)
        self.s.insert("ManualEntryLines", ManualEntryID=mid, LineNumber=1, AccountCode=5900, Debit=10, Credit=0)
        self.s.insert("ManualEntryLines", ManualEntryID=mid, LineNumber=2, AccountCode=3100, Debit=0, Credit=10)
        self.assertEqual(self.s.sync_journal(), (0, 0, 0))
        self.assertIn(("MANUAL", mid), self.s.locked)


class YearClosingTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.s = Store(self.db.con, now=day(0, 9))
        self.s.sync_journal()
        self.db.params["PeriodStart"] = "2000-01-01 00:00:00"
        self.db.params["PeriodEnd"] = "2027-01-01 00:00:00"
        self.db.params["CompareStart"] = "1999-01-01 00:00:00"
        self.db.params["CompareEnd"] = "1999-02-01 00:00:00"
        self.profit_before = self.db.scalar("SELECT CurrentValue FROM IncomeStatementQuery WHERE Block = 60")
        self.cid, self.profit = self.s.close_year(2026)

    def test_closing_entry(self):
        e = self.s.one("SELECT EntryID FROM JournalEntries WHERE SourceType = 'YEAR_CLOSE' AND SourceID = ?", self.cid)
        self.assertIsNotNone(e)
        self.assertAlmostEqual(float(self.profit), self.profit_before, places=4)
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0)
        self.assertEqual(self.s.one("SELECT EntryDate FROM JournalEntries WHERE EntryID = ?", e)[:10], "2026-12-31")

    def test_revenue_and_expenses_back_to_zero(self):
        left = self.db.scalar("SELECT COUNT(*) FROM TrialBalanceQuery AS t INNER JOIN Accounts AS a ON t.AccountCode = "
                              "a.AccountCode WHERE a.AccountType IN ('REVENUE', 'EXPENSE') AND "
                              "abs(t.ClosingBalance) > 0.0001")
        self.assertEqual(left, 0)
        self.assertAlmostEqual(self.db.scalar("SELECT ClosingBalance FROM TrialBalanceQuery WHERE AccountCode = 3300"),
                               -self.profit_before, places=4)

    def test_income_statement_still_shows_the_year(self):
        self.assertAlmostEqual(self.db.scalar("SELECT CurrentValue FROM IncomeStatementQuery WHERE Block = 60"),
                               self.profit_before, places=4)

    def test_balance_sheet_after_closing(self):
        assets = self.db.scalar("SELECT CurrentValue FROM BalanceSheetQuery WHERE ClassNo = 1 AND RowKind = 'T'")
        other = self.db.scalar("SELECT CurrentValue FROM BalanceSheetQuery WHERE ClassNo = 4")
        self.assertAlmostEqual(assets, other, places=4)
        self.assertAlmostEqual(self.db.scalar("SELECT CurrentValue FROM BalanceSheetQuery WHERE AccountKey = 'Z'"), 0,
                               places=4)
        self.assertAlmostEqual(self.db.scalar("SELECT CurrentValue FROM BalanceSheetQuery WHERE LineAccount = 3300"),
                               self.profit_before, places=4)

    def test_the_closing_survives_its_own_lock(self):
        self.assertEqual(self.s.one("SELECT ClosedThrough FROM Settings")[:10], "2026-12-31")
        self.assertEqual(self.s.sync_journal(), (0, 0, 0))
        self.db.con.execute("DELETE FROM FiscalYearClosings WHERE YearClosingID = ?", (self.cid,))   # reopen
        self.assertEqual(self.s.sync_journal(), (0, 0, 1), "the closing entry goes although the period is closed")


class ClosingCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modClosing")

    def test_every_place_that_writes_in_the_past_checks_the_date(self):
        forms_text = read("modForms")
        self.assertIn("closedMsg = ClosedRecordProblem(frm)", proc(forms_text, "FormBeforeUpdate"))
        self.assertIn("ClosedRecordProblem(frm, True)", proc(forms_text, "DeleteRecord"))
        body = proc(self.text, "ClosedRecordProblem")
        for part in ('Case "Expenses"', 'Case "Customers", "Suppliers"', 'Case "CashBoxes"', "frm!ExpenseDate.OldValue",
                     "frm!CreatedAt.Value", "frm!OpeningDate.OldValue"):
            self.assertIn(part, body)
        manual = read("modManualEntry")
        self.assertIn("ClosedPeriodProblem(EntryDate)", proc(manual, "PostManualEntry"))
        self.assertIn("ClosedPeriodProblem(DbValue(\"SELECT EntryDate FROM ManualEntries", proc(manual, "RemoveManualEntry"))
        sync = proc(read("modJournal"), "SyncJournal")
        self.assertEqual(sync.count("InClosedPeriod("), 4)
        self.assertIn('SourceType = "YEAR_CLOSE"', proc(read("modJournal"), "InClosedPeriod"))

    def test_rules(self):
        close = proc(self.text, "ClosePeriod")
        self.assertIn("If Through >= Date Then", close)
        self.assertIn("JournalReady()", close)
        reopen = proc(self.text, "ReopenPeriod")
        self.assertIn("أعد فتح السنة أولًا", reopen)
        self.assertIn("اكتب سبب إعادة الفتح", reopen)
        problem = proc(self.text, "YearCloseProblem")
        self.assertIn("FiscalYear >= Year(Date)", problem)
        self.assertIn("السنوات تُقفل بالترتيب", problem)
        year = proc(self.text, "CloseFiscalYear")
        for part in ("RETAINED_EARNINGS", "SetClosedThrough yearEnd", "ws.BeginTrans", "SyncJournal()",
                     "HAVING Sum(l.Debit) <> Sum(l.Credit)"):
            self.assertIn(part, year)
        self.assertIn("تُعاد فتح آخر سنة مقفلة فقط", proc(self.text, "ReopenFiscalYear"))
        self.assertIn("e.SourceType <> 'YEAR_CLOSE'", Q.query("qryIncomeMoves").sql)
        self.assertIn("qryIncomeMoves AS c", Q.query("qryIncomeAccounts").sql)

    def test_schema_permission_and_screen(self):
        self.assertIn(("YEAR_CLOSE", "قيد إقفال السنة", 15), table("JournalSourceTypes").seed_rows)
        self.assertIn("ClosedThrough", [f.name for f in table("Settings").fields])
        grants = [r for r in table("RolePermissions").seed_rows if r[1] == "PERIOD_CLOSE"]
        self.assertEqual(grants, [(1, "PERIOD_CLOSE")], "administrators only")
        self.assertEqual(F.SCREEN_PERMISSIONS["frmPeriodClosing"], "PERIOD_CLOSE")
        names = {c.name for c in MODELS["frmPeriodClosing"].controls}
        for pname in re.findall(r"^(?:Public|Private) (?:Sub|Function) (\w+)\(ByVal frm As Access\.Form", self.text, re.M):
            if pname == "ClosedRecordProblem":           # the data screens, not this one
                continue
            for ctl in set(re.findall(r"frm!(\w+)", proc(self.text, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        self.assertIn([True], [[t[5] for t in __import__("schema").ACCOUNT_TREE if t[0] == 3300]])

    def test_in_access_tests(self):
        body = proc(read("modJournal"), "TestJournal")
        for part in ("إقفال الفترة يمنع القيد اليدوي بتاريخ مقفل", "إعادة فتح الفترة"):
            self.assertIn(part, body)


class Static_modClosing(VbaModuleChecks, unittest.TestCase):
    module_name = "modClosing"
    vba = read("modClosing")


if __name__ == "__main__":
    unittest.main()
