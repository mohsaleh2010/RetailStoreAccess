"""Sales reps (modSalesReps, SalesReps, SalesRepTargets, CommissionRuns, CommissionLines): the rep of a
document, the performance and commission queries and the commission entry run on the SQLite mirror, the
commission amount runs in LibreOffice, the posting code and the screens are checked."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day
import forms as F
import queries as Q
import vba_harness as H
from relations import index_load, ACCESS_MAX_INDEXES
from schema import ACCOUNT_TREE, SCREEN_LIST, table
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}
REP_DOCS = ("Customers", "SalesInvoices", "SalesReturns", "CustomerPayments", "CashVouchers")


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


def commission(base, rate, adjustment=0):
    """Mirror of modSalesReps.CalcCommission."""
    c = round(base * rate + 1e-9, 2) + adjustment
    return max(c, 0)


class SalesRepSchemaTests(unittest.TestCase):

    def test_tables_and_document_fields(self):
        for t in REP_DOCS:
            fld = next(f for f in table(t).fields if f.name == "SalesRepID")
            self.assertEqual(fld.fk, "SalesReps.SalesRepID", t)
            self.assertFalse(fld.required, t)                         # BuildSchema upgrades an existing back-end
        accounts = {a[0]: a for a in ACCOUNT_TREE}
        self.assertEqual(accounts[2330][2], "LIABILITY")
        self.assertEqual(accounts[5530][2], "EXPENSE")
        self.assertIn("COMMISSION", {r[0] for r in table("JournalSourceTypes").seed_rows})
        self.assertIn("COMMISSION", next(f.rule for f in table("CashVouchers").fields if f.name == "Category"))
        self.assertLessEqual({"SALES_REP", "COMMISSION_RUN"}, {r[0] for r in table("Sequences").seed_rows})
        for t in ("SalesReps", "SalesRepTargets", "CommissionRuns", "CommissionLines", "Customers", "CashVouchers"):
            self.assertLessEqual(index_load(t), ACCESS_MAX_INDEXES - 4, t)
        screens = {s[0]: s for s in SCREEN_LIST}
        for name in ("frmSalesReps", "frmRepTargets", "frmCommissions"):
            self.assertEqual(screens[name][3], "SALES_REPS")
            self.assertEqual(F.SCREEN_PERMISSIONS[name], "SALES_REPS")


class SalesRepQueryTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.c = self.db.con
        self.s = Store(self.c, now=day(0, 9))
        self.s.insert("SalesReps", RepCode="REP-001", RepName="TEST مندوب", CommissionRate=0.02,
                      CommissionBase="SALES", IsActive=True)
        self.rep = self.s.one("SELECT SalesRepID FROM SalesReps WHERE RepCode = 'REP-001'")
        ids = self.db.ids
        self.c.execute("UPDATE SalesInvoices SET SalesRepID = ? WHERE SalesInvoiceID IN (?, ?)",
                       (self.rep, ids["INV1"], ids["INV2"]))
        self.c.execute("UPDATE CustomerPayments SET SalesRepID = ? WHERE PaymentID = ?", (self.rep, ids["RCV1"]))
        self.c.execute("UPDATE Customers SET SalesRepID = ? WHERE CustomerID = ?", (self.rep, ids["C2"]))

    def row(self, sql, *args):
        cur = self.c.execute(sql, args)
        r = cur.fetchone()
        return dict(zip([d[0] for d in cur.description], r)) if r else None

    def test_performance_and_target(self):
        ids = self.db.ids
        sales = self.s.one("SELECT Sum(TaxableAmount) FROM SalesInvoices WHERE SalesInvoiceID IN (?, ?)",
                           ids["INV1"], ids["INV2"])
        paid = self.s.one("SELECT Sum(PaidAmount) FROM SalesInvoices WHERE SalesInvoiceID IN (?, ?)",
                          ids["INV1"], ids["INV2"])
        when = day(0)[:7]
        self.s.insert("SalesRepTargets", SalesRepID=self.rep, TargetYear=int(when[:4]), TargetMonth=int(when[5:7]),
                      TargetAmount=1000)
        self.db.set_period(31)
        r = self.row("SELECT * FROM RepPerformanceQuery WHERE SalesRepID = ?", self.rep)
        self.assertAlmostEqual(r["NetSales"], sales, places=2)
        self.assertAlmostEqual(r["Collections"], paid + 200, places=2)
        self.assertAlmostEqual(r["Target"], 1000, places=2)
        self.assertAlmostEqual(r["Achievement"], sales / 1000, places=4)
        self.assertAlmostEqual(r["Commission"], commission(sales, 0.02), places=2)
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM RepCustomersQuery WHERE SalesRepID = ?", self.rep), 1)

    def test_commission_entry_and_payment(self):
        self.s.insert("CostCenters", CenterCode="C1", CenterName="TEST فرع", IsActive=True)
        center = self.s.one("SELECT CostCenterID FROM CostCenters WHERE CenterCode = 'C1'")
        self.s.insert("CommissionRuns", RunNumber="COM-00001", RunMonth=day(0)[:10], Status="POSTED", TotalAmount=150,
                      EmployeeID=1)
        run = self.s.one("SELECT CommissionRunID FROM CommissionRuns")
        self.s.insert("CommissionLines", CommissionRunID=run, SalesRepID=self.rep, RepName="TEST مندوب",
                      CostCenterID=center, NetSales=5000, Collections=0, CommissionBase="SALES", BaseAmount=5000,
                      CommissionRate=0.03, Adjustment=0, Commission=150)
        self.s.sync_journal()
        entry = self.row("SELECT * FROM JournalEntries WHERE SourceType = 'COMMISSION' AND SourceID = ?", run)
        self.assertEqual((entry["TotalDebit"], entry["TotalCredit"]), (150, 150))
        lines = self.c.execute("SELECT AccountCode, Debit, Credit, CostCenterID FROM JournalLines WHERE EntryID = ? "
                               "ORDER BY AccountCode", (entry["EntryID"],)).fetchall()
        self.assertEqual([tuple(l) for l in lines], [(2330, 0, 150, center), (5530, 150, 0, center)])
        bal = self.row("SELECT * FROM RepCommissionBalanceQuery WHERE SalesRepID = ?", self.rep)
        self.assertEqual((bal["Posted"], bal["Paid"], bal["Payable"]), (150, 0, 150))
        # the payment: a cash voucher of category COMMISSION debits 2330
        box = self.s.one("SELECT Min(CashBoxID) FROM CashBoxes")
        self.s.insert("CashVouchers", VoucherNumber="TEST-COM", VoucherDate=day(0, 10), VoucherType="OUT",
                      CashBoxID=box, Category="COMMISSION", Amount=100, Description="TEST", EmployeeID=1,
                      SalesRepID=self.rep)
        self.s.sync_journal()
        self.assertEqual(self.row("SELECT * FROM RepCommissionBalanceQuery WHERE SalesRepID = ?", self.rep)["Payable"], 50)
        voucher = self.s.one("SELECT CashVoucherID FROM CashVouchers WHERE VoucherNumber = 'TEST-COM'")
        debit = self.s.one("SELECT l.AccountCode FROM JournalLines AS l INNER JOIN JournalEntries AS e ON "
                           "l.EntryID = e.EntryID WHERE e.SourceType = 'CASH_VOUCHER' AND e.SourceID = ? AND l.Debit > 0",
                           voucher)
        self.assertEqual(debit, 2330)
        # a draft has no entry
        self.c.execute("UPDATE CommissionRuns SET Status = 'DRAFT'")
        self.s.sync_journal()
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM JournalEntries WHERE SourceType = 'COMMISSION'"), 0)


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class CommissionRuntimeTests(unittest.TestCase):

    def test_calc_commission(self):
        h = H.Harness()
        try:
            wrap = ("Option Explicit\n"
                    "Public Function WCalc(ByVal b As Double, ByVal r As Double, ByVal a As Double) As Double\n"
                    "    WCalc = CDbl(CalcCommission(CCur(b), r, CCur(a)))\nEnd Function\n")
            h.load({"modSalesReps": H.read_module("modSalesReps"), "modWrap": wrap,
                    "modRound": "Option Explicit\n" + H.HARNESS_ROUND})
            for base, rate, adj in [(1000, 0.025, 0), (1000, 0.025, -40), (-500, 0.1, 0), (1234.5, 0.015, 10),
                                    (333.33, 0.03, 0)]:
                self.assertAlmostEqual(h.call("modWrap", "WCalc", float(base), float(rate), float(adj)),
                                       commission(base, rate, adj), places=2, msg=(base, rate, adj))
        finally:
            h.close()


class SalesRepCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modSalesReps")

    def test_documents_take_the_rep(self):
        self.assertIn("rs!SalesRepID = SalesRepFor(CustomerID)", proc(read("modSales"), "PostSaleFromCart"))
        self.assertIn("rs!SalesRepID = SalesRepFor(CustomerID)", proc(read("modSales"), "PostCustomerPayment"))
        self.assertIn("SELECT SalesRepID FROM SalesInvoices", proc(read("modSales"), "PostSalesReturn"))
        rep_for = proc(self.text, "SalesRepFor")
        self.assertIn("s.IsActive = True", rep_for)
        self.assertIn("EmployeeID = \" & CurrentUserID()", rep_for)

    def test_posting_rules(self):
        post = proc(self.text, "PostCommissionRun")
        self.assertIn("ClosedPeriodProblem(rm)", post)
        self.assertIn("SyncJournal", post)
        self.assertIn("(CommissionRate < 0 OR CommissionRate >= 1) AND", post)
        unpost = proc(self.text, "UnpostCommissionRun")
        self.assertIn("RepPayable(rs!SalesRepID) < rs!Commission", unpost)
        delete = proc(self.text, "DeleteCommissionRun")
        self.assertLess(delete.index("AuditSnapshot("), delete.index("DELETE FROM CommissionRuns"))
        self.assertIn('AuditDeleted "COMMISSION_DELETE"', delete)
        self.assertIn("qryJournalCommission", read("modJournal"))
        self.assertIn("qryJournalCommission", Q.JOURNAL_SOURCE_QUERIES)

    def test_cash_voucher(self):
        cash = read("modCash")
        self.assertIn("COMMISSION;صرف عمولة مندوب", cash)
        self.assertIn("CommissionVoucherProblem(kind", proc(cash, "SaveCashVoucher"))
        self.assertIn("UPDATE CashVouchers SET SalesRepID", proc(cash, "SaveCashVoucher"))
        self.assertIn('Case "COMMISSION"', proc(cash, "VoucherCategoryChanged"))

    def test_screens(self):
        names = {c.name for c in MODELS["frmCommissions"].controls}
        for ctl in re.findall(r"frm!(\w+)", self.text):
            if ctl not in ("SalesRepID", "lblRepInfo"):
                self.assertIn(ctl, names, ctl)
        reps = {c.name for c in MODELS["frmSalesReps"].controls}
        self.assertLessEqual({"lblRepInfo", "CommissionRate", "CommissionBase"}, reps)
        self.assertEqual(table("SalesReps").pk, ["SalesRepID"])          # frm!SalesRepID: a field of the record
        lines = {c.name for c in MODELS["frmCommissionLines"].controls}
        for ctl in re.findall(r"lines!(\w+)", self.text):
            self.assertIn(ctl, lines, ctl)
        self.assertIn("SalesRepID", {c.name for c in MODELS["frmCustomers"].controls})
        self.assertIn('Case "SalesReps": SalesRepCurrent frm', read("modForms"))
        self.assertLessEqual({"REP_PERFORMANCE", "REP_CUSTOMERS", "REP_COMMISSION_BALANCE"}, {r.key for r in F.REPORTS})
        self.assertIn('"TestSalesReps"', read("modTestAll"))


class Static_modSalesReps(VbaModuleChecks, unittest.TestCase):
    module_name = "modSalesReps"
    vba = read("modSalesReps")


if __name__ == "__main__":
    unittest.main()
