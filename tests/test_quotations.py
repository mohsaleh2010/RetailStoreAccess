"""Customer quotations (modQuotations, frmQuotation, rptQuotation, docs/75-Customer-Quotations.md):
  * the tables: Quotations / QuotationLines, the sequence QT-, the permission QUOTATIONS, the screen;
  * the mirror of the rules (tools/quotation_reference.py): expiry, what may be converted, the issued totals;
  * replayed on the demo data in SQLite: the list report (status, expired, what it became) and the print data;
  * the code: the invoice of a quotation accepts its prices and marks it inside the sale transaction, a deleted
    draft order gives its quotation back, a quoted price other than the customer's needs PRICE_OVERRIDE.
"""
import os
import re
import sys
import unittest
from datetime import date
from decimal import Decimal as D

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools"))
sys.path.insert(0, HERE)

import demo_data as DD
import forms as F
import master_en
import quotation_reference as R
from helpers import VbaModuleChecks
from reports_docs import quotation
from schema import SCREEN_LIST, table

MODELS = {m.name: m for m in F.all_forms()}
NOW = "2026-10-01 00:00:00"


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


VBA = read("modQuotations")


class SchemaTests(unittest.TestCase):

    def test_tables(self):
        f = {x.name: x for x in table("Quotations").fields}
        self.assertEqual(f["CustomerID"].fk, "Customers.CustomerID")
        self.assertTrue(f["ValidUntil"].required)
        self.assertEqual(f["Status"].rule, 'In ("DRAFT","ISSUED","ORDERED","INVOICED","REJECTED")')
        self.assertEqual(f["SalesOrderID"].fk, "SalesOrders.SalesOrderID")
        self.assertEqual(f["SalesInvoiceID"].fk, "SalesInvoices.SalesInvoiceID")
        self.assertFalse(f["SalesOrderID"].required or f["SalesInvoiceID"].required)
        f = {x.name: x for x in table("QuotationLines").fields}
        self.assertTrue(f["QuotationID"].on_delete_cascade)
        self.assertEqual(f["Quantity"].rule, ">0")

    def test_screen_permission_sequence(self):
        self.assertIn(("frmQuotation", "عروض الأسعار للعملاء", "المبيعات", "QUOTATIONS", True, True, True), SCREEN_LIST)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmQuotation"], "QUOTATIONS")
        self.assertIn("QUOTATIONS", [r[0] for r in table("Permissions").seed_rows])
        self.assertIn("QUOTATIONS", master_en.PERMISSION_NAMES_EN)
        self.assertIn("frmQuotation", master_en.SCREEN_TITLES_EN)
        self.assertIn(("QUOTATION", "QT-", 1, 6, "عروض الأسعار للعملاء"), table("Sequences").seed_rows)
        roles = {r for r, p in table("RolePermissions").seed_rows if p == "QUOTATIONS"}
        self.assertEqual(roles, {1, 2}, "the admin and the manager")

    def test_screens(self):
        m = MODELS["frmQuotation"]
        names = {c.name for c in m.controls}
        for b in ("btnIssue", "btnPrint", "btnToOrder", "btnToInvoice", "cboPayType", "btnUnissue", "btnReject",
                  "btnDelete", "btnCopy", "btnOpenDoc"):
            self.assertIn(b, names)
        self.assertIn("frmQuotationLines", MODELS)
        order = MODELS["frmSalesOrder"]
        btn = next(c for c in order.controls if c.name == "btnQuotations")
        self.assertIn('OpenScreen "frmQuotation", 0', str(btn.props) + str(order.code))

    def test_reports(self):
        keys = {r.key: r for r in F.REPORTS}
        self.assertEqual((keys["QUOTATIONS"].query, keys["QUOTATIONS"].needs, keys["QUOTATIONS"].date_column),
                         ("QuotationListQuery", "Dc", "QuoteDate"))
        rpt = quotation()
        self.assertEqual((rpt.name, rpt.record_source, rpt.group), ("rptQuotation", "qryQuotationPrint", "QuotationID"))


class MirrorTests(unittest.TestCase):

    def test_expiry(self):
        today = date(2026, 10, 10)
        self.assertEqual(R.state_for("ISSUED", date(2026, 10, 9), today), "EXPIRED")
        self.assertEqual(R.state_for("ISSUED", date(2026, 10, 10), today), "ISSUED", "valid on its last day")
        self.assertEqual(R.state_for("ORDERED", date(2026, 10, 1), today), "ORDERED", "a converted one never expires")
        self.assertEqual(R.state_for("DRAFT", date(2026, 10, 1), today), "DRAFT")

    def test_convert(self):
        self.assertIsNone(R.convert_problem("ISSUED"))
        for state in ("DRAFT", "EXPIRED", "ORDERED", "INVOICED", "REJECTED"):
            self.assertIsNotNone(R.convert_problem(state), state)

    def test_issue_totals(self):
        lines, (sub, tax, total) = R.issue_totals([(3, 20, True), (1, 10, False)], "0.15", True)
        self.assertEqual(total, D("70.00"), "VAT included: the customer pays the shelf price")
        self.assertEqual(sub + tax, total)
        self.assertEqual(lines[0][1], D("7.83"))
        lines, (sub, tax, total) = R.issue_totals([(3, 20, True)], "0.15", False)
        self.assertEqual((sub, tax, total), (D("60.00"), D("9.00"), D("69.00")))


class ReplayTests(unittest.TestCase):
    """Quotations in each state added to the demo data in the SQLite mirror of Access."""

    @classmethod
    def setUpClass(cls):
        cls.r = DD.simulate()
        c = cls.c = cls.r.db.con
        cust = cls.r.ids["cust"][1]
        cls.p1 = cls.r.ids["prod"][1]
        inv = c.execute("SELECT Min(SalesInvoiceID) FROM SalesInvoices").fetchone()[0]
        cls.inv_no = c.execute("SELECT InvoiceNumber FROM SalesInvoices WHERE SalesInvoiceID = ?", (inv,)).fetchone()[0]
        cls.ids = {}
        for no, status, valid, inv_id in (("QT-T1", "ISSUED", "2099-01-01 00:00:00", None),
                                          ("QT-T2", "ISSUED", "2020-01-01 00:00:00", None),
                                          ("QT-T3", "INVOICED", "2020-01-01 00:00:00", inv),
                                          ("QT-T4", "REJECTED", "2099-01-01 00:00:00", None)):
            c.execute("INSERT INTO Quotations (QuotationNumber, QuoteDate, CustomerID, ValidUntil, Status, SubTotal, Tax, "
                      "TotalAmount, SalesInvoiceID, EmployeeID, Terms, CreatedAt) VALUES (?, ?, ?, ?, ?, 52.17, 7.83, 60, ?, 1, "
                      "'TEST', ?)", (no, NOW, cust, valid, status, inv_id, NOW))
            cls.ids[no] = c.execute("SELECT Max(QuotationID) FROM Quotations").fetchone()[0]
            c.execute("INSERT INTO QuotationLines (QuotationID, ProductID, Quantity, UnitPrice, NetAmount, Tax, LineTotal, "
                      "CreatedAt) VALUES (?, ?, 3, 20, 52.17, 7.83, 60, ?)", (cls.ids[no], cls.p1, NOW))

    def q(self, sql, *args):
        return self.c.execute(sql, args).fetchall()

    def test_list_report(self):
        rows = {r[0]: r[1:] for r in self.q("SELECT QuotationNumber, StatusName, ConvertedNumber, IsConverted, TotalAmount "
                                            "FROM QuotationListQuery")}
        self.assertEqual(rows["QT-T1"][:3], ("صادر", "", 0))
        self.assertEqual(rows["QT-T2"][:3], ("منتهي", "", 0), "issued and past its date")
        self.assertEqual(rows["QT-T3"][:3], ("تحوّل إلى فاتورة", self.inv_no, 1), "a converted one is not expired")
        self.assertEqual(rows["QT-T4"][0], "رفضه العميل")
        self.assertEqual(rows["QT-T1"][3], 60)

    def test_print_data(self):
        rows = self.q("SELECT QuotationNumber, ProductName, Quantity, LineTotal, TotalAmount, CustomerName, EmployeeName "
                      "FROM qryQuotationPrint WHERE QuotationID = ?", self.ids["QT-T1"])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][2:5], (3, 60, 60))
        self.assertIsNotNone(rows[0][5])


class CodeTests(unittest.TestCase):

    def test_sale_accepts_and_marks_the_quotation(self):
        sale = proc(read("modSales"), "PostSaleFromCart")
        self.assertIn("Optional ByVal QuotationID As Long = 0", sale)
        self.assertIn("If IsQuotationPrice(QuotationID, rs!ProductID, rs!UnitID, rs!UnitPrice) Then", sale)
        mark = sale.index("MarkQuotationInvoiced db, QuotationID, invoiceID")
        self.assertLess(sale.index("ws.BeginTrans"), mark)
        self.assertLess(mark, sale.index("ws.CommitTrans"), "inside the sale transaction")
        body = proc(VBA, "InvoiceQuotation")
        self.assertIn("QuotationID:=QuotationID", body)
        self.assertIn('db.Execute "DELETE FROM tmpPOSLines"', body)
        self.assertIn("If QuotationID = 0 Then Exit Function", proc(VBA, "IsQuotationPrice"))

    def test_price_override_at_issue(self):
        body = proc(VBA, "IssueQuotation")
        self.assertIn("rs!UnitPrice <> QuotationDefaultPrice(cust, rs!ProductID)", body)
        self.assertIn('Not HasPermission("PRICE_OVERRIDE")', body)
        self.assertIn('CalcRun(0, Nz(SettingValue("PricesIncludeVAT"), True))', body)
        self.assertIn("ws.Rollback", body)

    def test_deleted_order_gives_the_quotation_back(self):
        body = proc(read("modSalesOrders"), "DeleteSalesOrder")
        self.assertLess(body.index("QuotationOrderDeleted SalesOrderID"), body.index("DELETE FROM SalesOrders"))
        self.assertIn("Status = 'ISSUED', SalesOrderID = Null", proc(VBA, "QuotationOrderDeleted"))

    def test_conversion_checks(self):
        for name in ("QuotationToOrder", "InvoiceQuotation"):
            self.assertIn("QuotationConvertProblem(QuotationID)", proc(VBA, name), name)
        self.assertIn("SalesOrderLineProblem(", proc(VBA, "QuotationToOrder"))
        self.assertIn('Case "EXPIRED"', proc(VBA, "QuotationConvertProblem"))

    def test_print(self):
        self.assertIn('OpenDocumentReport "rptQuotation", "[QuotationID] = "', proc(read("modReports"), "PrintQuotation"))

    def test_deletes_are_audited(self):
        self.assertIn('AuditDeleted "QUOTATION_DELETE"', proc(VBA, "DeleteQuotation"))
        self.assertIn('AuditDeleted "QUOTATION_LINE_DELETE"', proc(VBA, "QuotationRemoveLine"))

    def test_in_access_test_is_listed(self):
        self.assertIn('"TestQuotations"', read("modTestAll"))
        test = proc(VBA, "TestQuotations")
        self.assertIn("ws.Rollback", test)
        for call in ("CreateQuotation(", "IssueQuotation(", "QuotationToOrder(", "DeleteSalesOrder(", "InvoiceQuotation(",
                     "CopyQuotation(", "RejectQuotation(", "DeleteQuotation("):
            self.assertIn(call, test)


class Static_modQuotations(VbaModuleChecks, unittest.TestCase):
    module_name = "modQuotations"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
