"""The budget (modBudget, Budgets, BudgetLines, BudgetVsActualQuery) on the SQLite mirror:
  * the budget of the months of the period (whole months, the year of the period start);
  * the actual from the journal: the account and its sub-accounts, in the line's cost centre when it has one;
  * variance, percentage and the favourable / unfavourable note for revenue and expenses;
  * the VBA keeps the rules; the screen, the report and the permission."""
import datetime as dt
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day, TODAY
import forms as F
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


class BudgetTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.c = self.db.con
        self.s = Store(self.c, now=day(0, 9))
        self.s.sync_journal()
        self.year = TODAY.year
        self.bid = self.s.insert("Budgets", BudgetYear=self.year, BudgetName="TEST", EmployeeID=1)
        months = {f"M{k}": 100 * k for k in range(1, 13)}
        self.sales = self.s.insert("BudgetLines", BudgetID=self.bid, AccountCode=4100, **months)
        self.opex = self.s.insert("BudgetLines", BudgetID=self.bid, AccountCode=52, **{f"M{k}": 10 for k in range(1, 13)})
        # the whole year
        self.db.params["PeriodStart"] = f"{self.year}-01-01 00:00:00"
        self.db.params["PeriodEnd"] = f"{self.year + 1}-01-01 00:00:00"

    def row(self, line):
        return self.c.execute("SELECT BudgetAmount, ActualAmount, Variance, VariancePct, VarianceNote FROM "
                              "BudgetVsActualQuery WHERE BudgetLineID = ?", (line,)).fetchone()

    def actual(self, accounts_sql, sign):
        return self.s.one(f"SELECT Coalesce(Sum({sign}), 0) FROM JournalLines AS l INNER JOIN JournalEntries AS e ON "
                          f"l.EntryID = e.EntryID WHERE l.AccountCode IN ({accounts_sql}) AND e.EntryDate >= ? "
                          f"AND e.EntryDate < ?", self.db.params["PeriodStart"], self.db.params["PeriodEnd"])

    def test_year(self):
        budget, actual, variance, pct, note = self.row(self.sales)
        self.assertEqual(budget, 7800)                                   # 100 + 200 + ... + 1200
        self.assertAlmostEqual(actual, self.actual("4100", "l.Credit - l.Debit"), places=4)
        self.assertAlmostEqual(variance, actual - budget, places=4)
        self.assertAlmostEqual(pct, (actual - budget) / budget, places=6)
        self.assertEqual(note, "ملائم" if actual > budget else "غير ملائم")

    def test_main_account_takes_its_sub_accounts(self):
        budget, actual, *_ = self.row(self.opex)
        self.assertEqual(budget, 120)
        expected = self.actual("SELECT AccountCode FROM Accounts WHERE Level2Code = 52", "l.Debit - l.Credit")
        self.assertAlmostEqual(actual, expected, places=4)
        self.assertGreater(actual, 0)
        self.assertEqual(self.row(self.opex)[4], "غير ملائم" if actual > budget else "ملائم")   # expenses: more is bad

    def test_whole_months_of_the_period(self):
        self.db.params["PeriodStart"] = f"{self.year}-03-15 00:00:00"   # March to May
        self.db.params["PeriodEnd"] = f"{self.year}-05-02 00:00:00"
        self.assertEqual(self.row(self.sales)[0], 300 + 400 + 500)

    def test_cost_centre(self):
        center = self.s.insert("CostCenters", CenterCode="R", CenterName="TEST", IsDefault=1, IsActive=1)
        self.c.execute("UPDATE BudgetLines SET CostCenterID = ? WHERE BudgetLineID = ?", (center, self.sales))
        self.assertEqual(self.row(self.sales)[1] or 0, 0)               # nothing on that centre yet
        self.c.execute("UPDATE SalesInvoices SET CostCenterID = ?", (center,))
        self.s.sync_journal()
        self.assertAlmostEqual(self.row(self.sales)[1], self.actual("4100", "l.Credit - l.Debit"), places=4)

    def test_other_year(self):
        self.db.params["PeriodStart"] = f"{self.year - 1}-01-01 00:00:00"
        self.db.params["PeriodEnd"] = f"{self.year}-01-01 00:00:00"
        self.assertEqual(self.c.execute("SELECT COUNT(*) FROM BudgetVsActualQuery").fetchone()[0], 0)


class BudgetCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modBudget")

    def test_rules(self):
        self.assertIn("AccountType IN ('REVENUE', 'EXPENSE')", proc(self.text, "BudgetAccountProblem"))
        self.assertIn("موجود في الموازنة لنفس المركز", proc(self.text, "BudgetLineProblem"))
        self.assertIn("توجد موازنة لسنة", proc(self.text, "CreateBudget"))
        self.assertIn("rs!M12 = Annual - 11 * part", proc(self.text, "SpreadAnnual"))
        self.assertIn("e.SourceType <> 'YEAR_CLOSE'", proc(self.text, "ActualOfMonth"))
        self.assertIn("المقارنة داخل سنة واحدة", proc(self.text, "ComparePeriodOK"))
        self.assertIn('"TestBudget"', read("modTestAll"))

    def test_screen_and_report(self):
        self.assertEqual(F.SCREEN_PERMISSIONS["frmBudget"], "BUDGET")
        names = {c.name for c in MODELS["frmBudget"].controls}
        for pname in re.findall(r"^Public Sub (\w+)\(ByVal frm As Access\.Form", self.text, re.M):
            for ctl in set(re.findall(r"frm!(\w+)", proc(self.text, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        lines = {c.source for c in MODELS["frmBudgetLines"].controls}
        self.assertTrue({f"M{k}" for k in range(1, 13)} | {"AccountCode", "CostCenterID", "BudgetID"} <= lines)
        self.assertIn("BUDGET_VS_ACTUAL", {r.key for r in F.REPORTS})


class Static_modBudget(VbaModuleChecks, unittest.TestCase):
    module_name = "modBudget"
    vba = read("modBudget")


if __name__ == "__main__":
    unittest.main()
