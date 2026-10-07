"""Cost centres (modCostCenters, CostCenters, JournalLines.CostCenterID, CostCenterProfitQuery,
CostCenterAccountsQuery) on the SQLite mirror:
  * entries of documents without a centre keep their signature (no rebuild, nothing in closed periods);
  * a centre on a document moves every line of its entry there; payroll lines go to each employee's centre;
    depreciation to the asset's centre; manual entries per line;
  * the centres together give the same profit as the income statement;
  * the VBA sets the centre where documents are made; the screens and reports."""
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


class CostCenterTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.c = self.db.con
        self.s = Store(self.c, now=day(0, 9))
        self.s.sync_journal()
        self.retail = self.s.insert("CostCenters", CenterCode="R", CenterName="TEST التجزئة", IsDefault=1, IsActive=1)
        self.cafe = self.s.insert("CostCenters", CenterCode="C", CenterName="TEST المقهى", IsDefault=0, IsActive=1)
        self.db.params.update({"PeriodStart": "2000-01-01 00:00:00", "PeriodEnd": "2100-01-01 00:00:00",
                               "CompareStart": "1999-01-01 00:00:00", "CompareEnd": "1999-02-01 00:00:00"})
        self.inv = self.db.ids["INV2"]

    def lines_of(self, kind, sid):
        return self.c.execute("SELECT l.CostCenterID FROM JournalLines AS l INNER JOIN JournalEntries AS e ON "
                              "l.EntryID = e.EntryID WHERE e.SourceType = ? AND e.SourceID = ?", (kind, sid)).fetchall()

    def test_entries_without_centre_are_not_rebuilt(self):
        self.assertEqual(self.s.sync_journal(), (0, 0, 0))
        self.assertEqual(self.c.execute("SELECT COUNT(*) FROM JournalLines WHERE CostCenterID Is Not Null").fetchone()[0], 0)

    def test_a_centre_moves_the_whole_entry(self):
        self.c.execute("UPDATE SalesInvoices SET CostCenterID = ? WHERE SalesInvoiceID = ?", (self.cafe, self.inv))
        self.assertEqual(self.s.sync_journal(), (0, 1, 0))
        self.assertEqual({r[0] for r in self.lines_of("SALE", self.inv)}, {self.cafe})

    def test_centres_add_up_to_the_income_statement(self):
        self.c.execute("UPDATE SalesInvoices SET CostCenterID = ? WHERE SalesInvoiceID = ?", (self.cafe, self.inv))
        self.c.execute("UPDATE Expenses SET CostCenterID = ?", (self.retail,))
        self.s.sync_journal()
        rows = self.c.execute("SELECT CenterKey, Revenue, CostOfSales, Expenses, NetProfit FROM CostCenterProfitQuery"
                              ).fetchall()
        total = sum(r[4] for r in rows)
        self.assertAlmostEqual(total, self.db.scalar("SELECT CurrentValue FROM IncomeStatementQuery WHERE Block = 60"),
                               places=4)
        by = {r[0]: r for r in rows}
        self.assertGreater(by[self.cafe][1], 0)                       # the café's sale
        self.assertGreater(by[self.retail][3], 0)                     # the retail expenses
        for key, revenue, cost, expenses, net in rows:
            self.assertAlmostEqual(net, revenue - cost - expenses, places=4)
        detail = self.c.execute("SELECT Sum(CenterAmount) FROM CostCenterAccountsQuery WHERE CenterKey = ?",
                                (self.cafe,)).fetchone()[0]
        self.assertIsNotNone(detail)

    def test_payroll_by_employee_centre(self):
        second = self.s.insert("Employees", EmployeeName="TEST موظف المقهى", Username="test_cafe", RoleID=3, IsActive=1)
        for emp, center in ((1, self.retail), (second, self.cafe)):
            self.c.execute("UPDATE Employees SET OnPayroll = 1, IsActive = 1, BasicSalary = 3000, HousingAllowance = 500, "
                           "IsSaudi = 0, CostCenterID = ? WHERE EmployeeID = ?", (center, emp))
        run = self.s.create_payroll(day(1)[:8] + "28")
        self.c.execute("UPDATE PayrollLines SET CostCenterID = (SELECT CostCenterID FROM Employees AS e WHERE "
                       "e.EmployeeID = PayrollLines.EmployeeID) WHERE PayrollRunID = ?", (run,))      # AddLines
        self.c.execute("UPDATE PayrollRuns SET Status = 'POSTED' WHERE PayrollRunID = ?", (run,))
        self.s.sync_journal()
        rows = self.c.execute("SELECT l.CostCenterID, Sum(l.Debit) FROM JournalLines AS l INNER JOIN JournalEntries AS e "
                              "ON l.EntryID = e.EntryID WHERE e.SourceType = 'PAYROLL' AND l.AccountCode = 5500 "
                              "GROUP BY l.CostCenterID", ).fetchall()
        self.assertEqual(dict(rows), {self.retail: 3500, self.cafe: 3500})
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0)

    def test_depreciation_and_manual_lines(self):
        aid = self.s.insert("FixedAssets", AssetCode="TEST-FA", AssetName="TEST", AssetAccount=1710,
                            CostCenterID=self.cafe, PurchaseDate=day(60), Cost=1200, InputVAT=0, SalvageValue=0,
                            UsefulLifeMonths=12, DepStartDate=day(60), SourceType="OPENING", OpeningAccumDep=0,
                            Status="ACTIVE", DisposalProceeds=0, DisposalAccumDep=0, EmployeeID=1)
        run, _ = self.s.depreciate(day(1)[:10])
        self.assertEqual({r[0] for r in self.lines_of("DEPRECIATION", run)}, {self.cafe})
        self.assertEqual({r[0] for r in self.lines_of("ASSET", aid)}, {self.cafe})
        mid = self.s.insert("ManualEntries", EntryNumber="TEST-MJ", EntryDate=day(1), Description="TEST", TotalAmount=10,
                            EmployeeID=1)
        self.s.insert("ManualEntryLines", ManualEntryID=mid, LineNumber=1, AccountCode=5900, Debit=10, Credit=0,
                      CostCenterID=self.retail)
        self.s.insert("ManualEntryLines", ManualEntryID=mid, LineNumber=2, AccountCode=3100, Debit=0, Credit=10)
        self.s.sync_journal()
        self.assertEqual({r[0] for r in self.lines_of("MANUAL", mid)}, {None, self.retail})


class CostCenterCodeTests(unittest.TestCase):

    def test_documents_get_their_centre(self):
        self.assertIn("rs!CostCenterID = CostCenterFor()", proc(read("modSales"), "PostSaleFromCart"))
        self.assertIn('rs!CostCenterID = DbValue("SELECT CostCenterID FROM SalesInvoices', proc(read("modSales"),
                                                                                            "PostSalesReturn"))
        self.assertEqual(read("modCash").count("rs!CostCenterID = CostCenterFor()"), 2)
        self.assertIn("frm!CostCenterID.Value = CostCenterFor()", read("modForms"))
        self.assertIn("l!CostCenterID = e!CostCenterID", read("modPayroll"))
        self.assertIn("h!CostCenterID = rs!LineCenter", read("modManualEntry"))
        self.assertIn("UPDATE FixedAssets SET CostCenterID", read("modAssets"))
        self.assertIn("ADD COLUMN LineCenter", read("modPOS"))
        self.assertIn("IIf(q.CostCenter = 0, Null, q.CostCenter)", proc(read("modJournal"), "SyncJournal"))
        self.assertIn('"TestCostCenters"', read("modTestAll"))
        for q in Q.JOURNAL_SOURCE_QUERIES:
            self.assertIn("AS CostCenter", Q.query(q).sql, q)

    def test_screens_and_reports(self):
        self.assertEqual(F.SCREEN_PERMISSIONS["frmCostCenters"], "JOURNAL")
        for form, ctl in [("frmExpenses", "CostCenterID"), ("frmEmployeePay", "CostCenterID"),
                          ("frmAssets", "cboCenter"), ("frmManualLines", "LineCenter")]:
            names = {c.name for c in MODELS[form].controls} | {c.source for c in MODELS[form].controls}
            self.assertIn(ctl, names, form)
        keys = {r.key for r in F.REPORTS}
        self.assertTrue({"COST_CENTER_PROFIT", "COST_CENTER_ACCOUNTS"} <= keys)
        for t in ("SalesInvoices", "SalesReturns", "Expenses", "CashVouchers", "FixedAssets", "PayrollLines",
                  "ManualEntryLines", "JournalLines", "Employees"):
            self.assertIn("CostCenterID", [f.name for f in table(t).fields], t)


class Static_modCostCenters(VbaModuleChecks, unittest.TestCase):
    module_name = "modCostCenters"
    vba = read("modCostCenters")


if __name__ == "__main__":
    unittest.main()
