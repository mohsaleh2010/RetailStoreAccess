"""Receivables / payables aging and payment links (modAging, frmAging, frmAllocation, rptAging,
CustomerAllocations / SupplierAllocations, DueDate, PaymentTermsDays, Settings.CreditBlockDays) on the
SQLite mirror, with the engine of tools/sim.py (Store.aging = ComputeAging):
  * the open documents of every party add up to its balance (customers and suppliers);
  * unlinked payments pay the oldest due document first; a linked payment pays its invoice;
    a return pays its own invoice; the due date comes from DueDate or the terms of the party;
  * the free amounts of payments and invoices (the queries behind frmAllocation);
  * the VBA keeps the rules (links, the credit block, due dates on new credit invoices) and the
    screens, report and permissions."""
import datetime as dt
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day, TODAY
import forms as F
import queries as Q
import reports as RP
import vba_harness as H
from schema import table, SCREEN_LIST
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


class AgingTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.c = self.db.con
        self.s = Store(self.c, now=day(0, 9))
        self.ids = self.db.ids
        self.today = TODAY.isoformat()
        self.cust = self.ids["C2"]

    def balances(self, kind):
        q = "CustomerBalanceQuery" if kind == "C" else "SupplierBalanceQuery"
        key = "CustomerID" if kind == "C" else "SupplierID"
        return {p: round(b, 4) for p, b in self.c.execute(f"SELECT {key}, Balance FROM {q}") if round(b, 4)}

    def aging_totals(self, kind):
        rows, credits = self.s.aging(kind, self.today)
        out = {}
        for party, *_rest, amount in rows:
            out[party] = out.get(party, 0) + float(amount)
        for party, amount in credits.items():
            out[party] = out.get(party, 0) - float(amount)
        return {p: round(v, 4) for p, v in out.items() if round(v, 4)}

    def open_of(self, inv):
        rows, _ = self.s.aging("C", self.today)
        return {r[2]: float(r[6]) for r in rows if r[1] == "INVOICE"}.get(inv, 0.0)

    def credit_invoice(self, amount, days_ago, due=None):
        no = self.s.next_number("SALES_INVOICE")
        inv = self.s.insert("SalesInvoices", InvoiceNumber=no, InvoiceDate=day(days_ago, 10), CustomerID=self.cust,
                            EmployeeID=1, PaymentType="CREDIT", PaymentMethodID=1, SubTotal=amount, Discount=0,
                            TaxableAmount=amount, Tax=0, TotalAmount=amount, PaidAmount=0, RemainingAmount=amount,
                            AmountTendered=0, ChangeDue=0, DueDate=due)
        self.s.adjust("Customers", "CustomerID", self.cust, amount)
        return inv

    def test_open_documents_add_up_to_the_balances(self):
        for kind in "CS":
            self.assertEqual(self.aging_totals(kind), self.balances(kind), kind)

    def test_oldest_due_paid_first(self):
        old = self.ids["INV2"]
        new = self.credit_invoice(300, 1)
        self.s.customer_payment(self.cust, 100)
        self.assertAlmostEqual(self.open_of(old), 64, places=4)          # 164 - 100
        self.assertAlmostEqual(self.open_of(new), 300, places=4)
        self.assertEqual(self.aging_totals("C"), self.balances("C"))

    def test_linked_payment_pays_its_invoice(self):
        old = self.ids["INV2"]
        new = self.credit_invoice(300, 1)
        self.s.customer_payment(self.cust, 100)
        pay = self.s.one("SELECT Max(PaymentID) FROM CustomerPayments")
        self.s.insert("CustomerAllocations", PaymentID=pay, SalesInvoiceID=new, Amount=100, EmployeeID=1)
        self.assertAlmostEqual(self.open_of(old), 164, places=4)
        self.assertAlmostEqual(self.open_of(new), 200, places=4)
        self.assertEqual(self.aging_totals("C"), self.balances("C"))
        free = self.c.execute("SELECT Allocated, Free FROM qryCustomerPaymentFree WHERE PaymentID = ?", (pay,)).fetchone()
        self.assertEqual(free, (100, 0))
        inv_free = self.c.execute("SELECT Allocated, Free FROM qryCustomerInvoiceFree WHERE InvoiceID = ?",
                                  (new,)).fetchone()
        self.assertEqual(inv_free, (100, 200))

    def test_payment_made_for_an_invoice_counts_as_linked(self):
        old = self.ids["INV2"]
        new = self.credit_invoice(300, 1)
        self.s.customer_payment(self.cust, 50)
        pay = self.s.one("SELECT Max(PaymentID) FROM CustomerPayments")
        self.c.execute("UPDATE CustomerPayments SET SalesInvoiceID = ? WHERE PaymentID = ?", (new, pay))
        self.assertAlmostEqual(self.open_of(new), 250, places=4)
        self.assertAlmostEqual(self.open_of(old), 164, places=4)

    def test_link_larger_than_the_invoice_goes_to_the_oldest(self):
        new = self.credit_invoice(30, 1)
        self.s.customer_payment(self.cust, 100)
        pay = self.s.one("SELECT Max(PaymentID) FROM CustomerPayments")
        self.s.insert("CustomerAllocations", PaymentID=pay, SalesInvoiceID=new, Amount=100, EmployeeID=1)
        self.assertAlmostEqual(self.open_of(new), 0, places=4)
        self.assertAlmostEqual(self.open_of(self.ids["INV2"]), 94, places=4)    # 164 - 70
        self.assertEqual(self.aging_totals("C"), self.balances("C"))

    def test_credit_left_when_payments_exceed(self):
        self.s.customer_payment(self.cust, 500)
        rows, credits = self.s.aging("C", self.today)
        self.assertNotIn(self.cust, {r[0] for r in rows})
        self.assertAlmostEqual(float(credits[self.cust]), 336, places=4)
        self.assertEqual(self.aging_totals("C"), self.balances("C"))

    def test_due_date(self):
        inv = self.credit_invoice(1000, 40)            # older than INV2: the earlier payments go to it first
        rows, _ = self.s.aging("C", self.today)
        due = {r[2]: r[5] for r in rows}
        self.assertEqual(due[inv], (TODAY - dt.timedelta(days=10)).isoformat())     # 40 days ago + 30 days
        self.c.execute("UPDATE SalesInvoices SET DueDate = ? WHERE SalesInvoiceID = ?", (day(-5), inv))
        rows, _ = self.s.aging("C", self.today)
        self.assertEqual({r[2]: r[5] for r in rows}[inv], (TODAY + dt.timedelta(days=5)).isoformat())

    def test_as_of_an_earlier_day(self):
        new = self.credit_invoice(300, 1)
        rows, _ = self.s.aging("C", (TODAY - dt.timedelta(days=2)).isoformat())
        self.assertNotIn(new, {r[2] for r in rows})

    def test_suppliers(self):
        sup = self.ids["S1"]
        self.s.payment(sup, 1000)
        self.assertEqual(self.aging_totals("S"), self.balances("S"))
        rows, _ = self.s.aging("S", self.today)
        self.assertAlmostEqual(float(rows[0][6]), 1855, places=4)


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class AgeColumnRuntimeTests(unittest.TestCase):

    def test_age_columns(self):
        h = H.Harness()
        try:
            h.load({"modAging": H.read_module("modAging")})
            for days, col in [(-5, 0), (0, 0), (1, 1), (30, 1), (31, 2), (60, 2), (61, 3), (90, 3), (91, 4), (400, 4)]:
                self.assertEqual(int(h.call("modAging", "AgeColumn", days)), col, days)
        finally:
            h.close()


class AgingCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modAging")

    def test_engine_and_rules(self):
        engine = proc(self.text, "ComputeAging")
        for part in ("qryAgingDebits", "qryAgingCredits", "qryAgingAllocations", 'rs!CreditType = "RETURN"',
                     "' 2. the rest pays the oldest due documents first"):
            self.assertIn(part, engine)
        self.assertIn("DateAdd('d', Nz(TermsDays, 0), DateValue(DocDate))", self.text)
        alloc = proc(self.text, "AllocatePayment")
        for part in ("payParty <> invParty", "Amount > payFree", "Amount > invFree", 'CanScreenAction("frmAllocation", "ADD"'):
            self.assertIn(part, alloc)
        self.assertIn('HasPermission("SUPPLIER_PAYMENTS")', proc(self.text, "KindParts"))
        sales = read("modSales")
        self.assertIn('MaxDaysLate("C", CustomerID, Date)', proc(sales, "CheckCustomer"))
        self.assertIn('rs!DueDate = DueDateFor("C", CustomerID, invDate)', proc(sales, "PostSaleFromCart"))
        self.assertIn('rs!DueDate = DueDateFor("S", SupplierID, docDate)', proc(read("modPurchases"), "PostPurchaseFromCart"))
        self.assertIn('"TestAging"', read("modTestAll"))
        self.assertIn("CREATE TABLE tmpAging", read("modPOS"))

    def test_schema_screens_report(self):
        for t, inv in [("CustomerAllocations", "SalesInvoiceID"), ("SupplierAllocations", "PurchaseInvoiceID")]:
            self.assertIn(inv, [f.name for f in table(t).fields])
        for t in ("SalesInvoices", "PurchaseInvoices"):
            self.assertIn("DueDate", [f.name for f in table(t).fields])
        for t in ("Customers", "Suppliers"):
            self.assertIn("PaymentTermsDays", [f.name for f in table(t).fields])
        self.assertIn("CreditBlockDays", [f.name for f in table("Settings").fields])
        self.assertEqual(F.SCREEN_PERMISSIONS["frmAging"], "REPORTS")
        self.assertEqual(F.SCREEN_PERMISSIONS["frmAllocation"], "CUSTOMER_PAYMENTS")
        self.assertIn("frmAllocation", {s[0] for s in SCREEN_LIST})
        for form in ("frmAging", "frmAllocation"):
            names = {c.name for c in MODELS[form].controls}
            for pname in re.findall(r"^Public Sub (\w+)\(ByVal frm As Access\.Form", self.text, re.M):
                if (form == "frmAging") != (pname.startswith("Aging") or pname == "PrintAging"):
                    continue
                for ctl in set(re.findall(r"frm!(\w+)", proc(self.text, pname))):
                    self.assertIn(ctl, names, f"{form} {pname}: {ctl}")
        for screen in ("frmCustomers", "frmSuppliers"):
            names = {c.name for c in MODELS[screen].controls}
            self.assertTrue({"btnAging", "btnAllocate"} <= names, screen)
        rpt = {r.name: r for r in RP.all_reports()}["rptAging"]
        self.assertEqual(rpt.record_source, "tmpAging")
        with open(os.path.join(ROOT, "tools", "generate.py"), encoding="utf-8") as fh:
            self.assertIn('"modAging"', fh.read())


class Static_modAging(VbaModuleChecks, unittest.TestCase):
    module_name = "modAging"
    vba = read("modAging")


if __name__ == "__main__":
    unittest.main()
