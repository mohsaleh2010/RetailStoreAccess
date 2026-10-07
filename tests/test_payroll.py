"""Monthly payroll (modPayroll, PayrollRuns, PayrollLines, frmPayroll, frmEmployeePay, journal sources PAYROLL /
PAYROLL_PAYMENT, employee advances) on the SQLite mirror with tools/sim.py (payroll_line = CalcPayLine,
create_payroll = CreatePayroll):
  * GOSI: Saudi employee and employer shares on basic + housing up to the maximum wage, non-Saudi employer share;
  * the draft takes the advance installment, never more than the advances or the salary;
  * the posted payroll entry (5500, 5510, 5520 / 2320, 1600, 4200, 2310) and the payment from a cash box;
  * the VBA keeps the rules, the screens, the permission, the report."""
import os
import re
import unittest
from decimal import Decimal as D

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day
import forms as F
from schema import ACCOUNT_TREE, table
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


class PayrollTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.c = self.db.con
        self.s = Store(self.c, now=day(0, 9))
        self.s.sync_journal()
        self.box = self.db.ids["BOXM"]
        self.c.execute("UPDATE Employees SET OnPayroll = 0")
        self.emp = 1
        self.c.execute("UPDATE Employees SET OnPayroll = 1, IsActive = 1, BasicSalary = 4000, HousingAllowance = 1000, "
                       "TransportAllowance = 400, OtherAllowance = 100, IsSaudi = 1, AdvanceInstallment = 500 "
                       "WHERE EmployeeID = 1")
        self.month = day(1)[:8] + "28"                 # any month end works for the mirror

    def balance(self, account):
        return round(self.s.one("SELECT Coalesce(Sum(Debit) - Sum(Credit), 0) FROM JournalLines WHERE AccountCode = ?",
                                account), 4)

    def advance(self, amount, kind="OUT"):
        self.s.insert("CashVouchers", VoucherNumber=f"TEST-ADV-{amount}-{kind}", VoucherDate=day(20), VoucherType=kind,
                      CashBoxID=self.box, Category="ADVANCE", Amount=amount, EmployeeID=1, AdvanceEmployeeID=self.emp)
        self.s.sync_journal()

    def test_gosi(self):
        self.assertEqual(self.s.payroll_line(5000, 1250, 500, saudi=True),
                         (D("6140.62"), D(6250), D("609.38"), D("734.38")))
        net, wage, ee, er = self.s.payroll_line(50000, 0, 0)
        self.assertEqual((wage, ee, er), (D(45000), D(0), D(900)))
        self.assertEqual(net, D(50000))

    def test_advances(self):
        self.advance(1200)
        self.advance(100, "IN")                        # paid back in cash
        self.assertEqual(self.s.advance_balance(self.emp), D(1100))
        self.assertAlmostEqual(self.balance(1600), self.s.one("SELECT Sum(Balance) FROM AdvanceBalanceQuery"), places=4)
        run = self.s.create_payroll(self.month)
        self.assertEqual(self.s.one("SELECT AdvanceDeduction FROM PayrollLines WHERE PayrollRunID = ?", run), 500)
        self.c.execute("UPDATE PayrollRuns SET Status = 'POSTED' WHERE PayrollRunID = ?", (run,))
        self.s.sync_journal()
        self.assertEqual(self.s.advance_balance(self.emp), D(600))
        self.assertAlmostEqual(self.balance(1600), self.s.one("SELECT Sum(Balance) FROM AdvanceBalanceQuery"), places=4)

    def test_posted_payroll_entry_and_payment(self):
        run = self.s.create_payroll(self.month)
        self.c.execute("UPDATE PayrollLines SET Overtime = 200, OtherDeduction = 50 WHERE PayrollRunID = ?", (run,))
        net, wage, ee, er = self.s.payroll_line(4000, 1000, 500, overtime=200, penalties=50, saudi=True)
        self.c.execute("UPDATE PayrollLines SET GosiEmployee = ?, GosiEmployer = ?, NetPay = ? WHERE PayrollRunID = ?",
                       (float(ee), float(er), float(net), run))
        before = {a: self.balance(a) for a in (5500, 5510, 5520, 2320, 4200, 2310)}
        self.c.execute("UPDATE PayrollRuns SET Status = 'POSTED' WHERE PayrollRunID = ?", (run,))
        self.s.sync_journal()
        self.assertEqual(self.balance(5500) - before[5500], 5000)
        self.assertEqual(self.balance(5510) - before[5510], 700)
        self.assertAlmostEqual(self.balance(5520) - before[5520], float(er), places=4)
        self.assertAlmostEqual(before[2320] - self.balance(2320), float(ee + er), places=4)
        self.assertEqual(before[4200] - self.balance(4200), 50)
        self.assertAlmostEqual(before[2310] - self.balance(2310), float(net), places=4)
        box_before = self.s.one("SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = ?", self.box)
        self.c.execute("UPDATE PayrollRuns SET PaidDate = ?, PaidFrom = 'CASHBOX', CashBoxID = ?, PaidAmount = ? "
                       "WHERE PayrollRunID = ?", (day(0), self.box, float(net), run))
        self.s.sync_journal()
        self.assertAlmostEqual(self.balance(2310), before[2310], places=4)
        self.assertAlmostEqual(self.s.one("SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = ?", self.box),
                               box_before - float(net), places=4)
        self.assertAlmostEqual(self.balance(110000 + self.box),
                               self.s.one("SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = ?", self.box), places=4)
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0)
        self.db.params["PayrollRunID"] = run
        self.assertEqual(self.c.execute("SELECT COUNT(*) FROM PayrollSheetQuery").fetchone()[0], 1)

    def test_draft_makes_no_entry(self):
        self.s.create_payroll(self.month)
        self.s.sync_journal()
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM JournalEntries WHERE SourceType LIKE 'PAYROLL%'"), 0)


class PayrollCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modPayroll")

    def test_rules(self):
        problem = proc(self.text, "PayrollProblem")
        for part in ("خصم الغياب أكبر من الأساسي والسكن", "خصم السلفة أكبر من رصيد سلفه", "الخصومات أكبر من الراتب"):
            self.assertIn(part, problem)
        self.assertIn("ClosedPeriodProblem(pm)", proc(self.text, "PostPayroll"))
        self.assertIn("يوجد مسير مرحَّل بعده", proc(self.text, "UnpostPayroll"))
        pay = proc(self.text, "PayPayroll")
        for part in ("CashBoxBalance(", "تاريخ الصرف قبل شهر المسير", "ClosedPeriodProblem(PaidDate)"):
            self.assertIn(part, pay)
        self.assertIn("SourceType = 'PAYROLL_PAYMENT'", proc(self.text, "UndoPayrollPayment"))
        calc = proc(self.text, "CalcPayLine")
        for part in ('SettingValue("GosiMaxWage")', 'SettingValue("GosiEmployeeRate")', 'SettingValue("GosiNonSaudiRate")'):
            self.assertIn(part, calc)
        cash = read("modCash")
        self.assertIn("AdvanceEmployeeProblem(kind,", cash)
        self.assertIn("UPDATE CashVouchers SET AdvanceEmployeeID", cash)
        self.assertIn('"ADVANCE;سداد سلفة موظف"', cash)
        self.assertIn("qryJournalPayroll", read("modJournal"))
        self.assertIn('"TestPayroll"', read("modTestAll"))

    def test_schema_screens(self):
        tree = {r[0]: r for r in ACCOUNT_TREE}
        for code in (2310, 2320, 5500, 5510, 5520):
            self.assertTrue(tree[code][5], code)
        self.assertIn("AdvanceEmployeeID", [f.name for f in table("CashVouchers").fields])
        for screen in ("frmPayroll", "frmEmployeePay"):
            self.assertEqual(F.SCREEN_PERMISSIONS[screen], "PAYROLL")
        names = {c.name for c in MODELS["frmPayroll"].controls}
        for pname in re.findall(r"^Public Sub (\w+)\(ByVal frm As Access\.Form", self.text, re.M):
            for ctl in set(re.findall(r"frm!(\w+)", proc(self.text, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        # a bound form reaches every field of its record source (Me!Field), with or without a control
        lines = {c.source for c in MODELS["frmPayrollLines"].controls} | {f.name for f in table("PayrollLines").fields}
        for ctl in set(re.findall(r"lines!(\w+)", proc(self.text, "PayrollLineChanged"))):
            self.assertIn(ctl, lines)
        self.assertIn("cboEmployee", {c.name for c in MODELS["frmCashVoucher"].controls})


class Static_modPayroll(VbaModuleChecks, unittest.TestCase):
    module_name = "modPayroll"
    vba = read("modPayroll")


if __name__ == "__main__":
    unittest.main()
