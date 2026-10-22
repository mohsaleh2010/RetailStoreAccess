"""Multiple units per product (modProductUnits, frmProductUnits, docs/72-Product-Units.md):
  * the table and the line fields: ProductUnits (factor, own barcode and prices), UnitID / UnitFactor on the sales
    and purchase lines and their returns (factor 1 by default: the old lines are pieces);
  * the mirror of the rules (tools/product_units_reference.py): the price of a unit, the pieces and the cost a line
    moves;
  * replayed on the demo data in SQLite: a carton line gives the product reports pieces and the right cost, the
    printed documents show its unit;
  * the code: posting moves Quantity x UnitFactor pieces, returns keep the unit of the invoice line, the scan finds a
    unit barcode, the carts have a unit column, the e-invoice lines take the unit code of the line.
"""
import os
import re
import sys
import unittest
from decimal import Decimal as D

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools"))
sys.path.insert(0, HERE)

import demo_data as DD
import forms as F
import product_units_reference as R
from helpers import VbaModuleChecks
from schema import SCREEN_LIST, table

MODELS = {m.name: m for m in F.all_forms()}
LINE_TABLES = ["SalesInvoiceDetails", "SalesReturnDetails", "PurchaseInvoiceDetails", "PurchaseReturnDetails"]


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


VBA = read("modProductUnits")


class SchemaTests(unittest.TestCase):

    def test_product_units(self):
        t = table("ProductUnits")
        f = {x.name: x for x in t.fields}
        self.assertEqual(f["ProductID"].fk, "Products.ProductID")
        self.assertEqual(f["UnitID"].fk, "Units.UnitID")
        self.assertEqual(f["UnitFactor"].rule, ">0 And <>1")
        self.assertEqual(f["SellingPrice"].default, "0")
        idx = {i.name: i for i in t.indexes}
        self.assertTrue(idx["UX_ProductID_UnitID"].unique)
        self.assertTrue(idx["UX_Barcode"].unique and idx["UX_Barcode"].ignore_nulls)

    def test_line_units(self):
        for name in LINE_TABLES:
            f = {x.name: x for x in table(name).fields}
            self.assertEqual(f["UnitID"].fk, "Units.UnitID", name)
            self.assertFalse(f["UnitID"].required, name)
            self.assertEqual((f["UnitFactor"].default, f["UnitFactor"].rule), ("1", ">0"), name)

    def test_screen(self):
        self.assertIn(("frmProductUnits", "وحدات الأصناف", "المخزون", "PRODUCTS", True, True, True), SCREEN_LIST)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmProductUnits"], "PRODUCTS")
        screen = next(s for s in F.DATA_SCREENS if s.name == "frmProductUnits")
        self.assertEqual(screen.table, "ProductUnits")
        products = next(s for s in F.DATA_SCREENS if s.name == "frmProducts")
        self.assertIn(("btnProductUnits", "وحدات الصنف", "OpenProductUnits Me"), products.extra_buttons)


class MirrorTests(unittest.TestCase):

    def test_factor(self):
        units = {7: 24}
        self.assertEqual(R.unit_factor(1, None, units), 1)
        self.assertEqual(R.unit_factor(1, 1, units), 1)
        self.assertEqual(R.unit_factor(1, 7, units), 24)
        self.assertEqual(R.unit_factor(1, 9, units), 0, "not one of its units")

    def test_prices(self):
        self.assertEqual(R.unit_sale_price("4.50", 24), D("108.00"))
        self.assertEqual(R.unit_sale_price("4.50", 24, 100), D("100.00"), "the unit's own price")
        self.assertEqual(R.unit_sale_price("4.50", 1, 100), D("4.50"), "a piece keeps the piece price")
        self.assertEqual(R.unit_purchase_price("2.125", 12), D("25.5000"))

    def test_stock_and_cost(self):
        pieces, piece_cost, line_cost = R.sale_stock_move(2, 24, "3.1250")
        self.assertEqual((pieces, piece_cost, line_cost), (D(-48), D("3.1250"), D("75.0000")))
        self.assertEqual(2 * line_cost, -pieces * piece_cost, "Quantity x UnitCost = the cost of the pieces")
        self.assertEqual(R.purchase_stock_move(2, 12, 48), (D(24), D("2.0000")))


class ReplayTests(unittest.TestCase):
    """A carton sale and its return added to the demo data in the SQLite mirror of Access."""

    @classmethod
    def setUpClass(cls):
        cls.r = DD.simulate()
        c = cls.c = cls.r.db.con
        pid = cls.pid = cls.r.ids["prod"][1]
        c.execute("INSERT INTO Units (UnitName, IsActive) VALUES ('كرتونة', 1)")
        cls.carton = c.execute("SELECT Max(UnitID) FROM Units").fetchone()[0]
        c.execute("INSERT INTO ProductUnits (ProductID, UnitID, UnitFactor, Barcode, SellingPrice, PurchasePrice, "
                  "IsActive, CreatedAt) VALUES (?, ?, 12, 'CARTON-1', 0, 0, 1, '2026-10-01 00:00:00')", (pid, cls.carton))
        inv = c.execute("SELECT Max(SalesInvoiceID), Max(InvoiceNumber) FROM SalesInvoices").fetchone()
        cls.inv = inv[0]
        c.execute("INSERT INTO SalesInvoiceDetails (SalesInvoiceID, LineNumber, ProductID, Quantity, UnitPrice, Discount, "
                  "NetAmount, VATCategory, VATRate, Tax, LineTotal, UnitCost, UnitID, UnitFactor) "
                  "VALUES (?, 99, ?, 2, 100, 0, 173.91, 'S', 0.15, 26.09, 200, 36, ?, 12)", (cls.inv, pid, cls.carton))

    def q(self, sql, *args):
        return self.c.execute(sql, args).fetchall()

    def test_old_lines_are_pieces(self):
        self.assertEqual(self.q("SELECT Count(*) FROM SalesInvoiceDetails WHERE UnitFactor <> 1 AND LineNumber <> 99")[0][0], 0)
        self.assertEqual(self.q("SELECT Count(*) FROM PurchaseInvoiceDetails WHERE UnitFactor <> 1")[0][0], 0)

    def test_reports_count_pieces_at_the_right_cost(self):
        qty, cost = self.q("SELECT SignedQty, LineCost FROM qrySalesLineItems WHERE DocType = 'SALE' AND DocID = ? "
                           "AND ProductID = ? AND SignedQty = 24", self.inv, self.pid)[0]
        self.assertEqual((qty, cost), (24, 72), "2 cartons = 24 pieces, cost 2 x 36")

    def test_print_shows_the_unit_of_the_line(self):
        rows = self.q("SELECT LineNumber, UnitName, LineUnit FROM qrySalesDocPrint WHERE DocKind = 'SALE' AND DocID = ? "
                      "ORDER BY LineNumber", self.inv)
        self.assertEqual(rows[-1][0], 99)
        self.assertEqual(rows[-1][2], "كرتونة")
        self.assertTrue(all(r[2] is None for r in rows[:-1]), "a piece line prints no extra unit")
        self.assertIsNotNone(rows[-1][1], "the product's unit stays available")


class CodeTests(unittest.TestCase):

    def test_sale_moves_pieces(self):
        body = proc(read("modSales"), "PostSaleFromCart")
        self.assertIn("factors(n) = UnitFactorOf(rs!ProductID, rs!UnitID)", body, "the factor of now, not of the cart")
        self.assertIn("costs(n) = pieceCosts(n) * factors(n)", body)
        self.assertIn('ApplyStockMovement db, productIDs(i), -CalcLine(i, "QTY") * factors(i), TT_SALE, pieceCosts(i)', body)
        self.assertIn("rs!UnitFactor = factors(i)", body)
        self.assertIn("IsListedUnitPrice(rs!ProductID, rs!UnitID, rs!UnitPrice)", body)
        self.assertIn('qtys(i) = CalcLine(i, "QTY") * Factors(i)', proc(read("modSales"), "CheckStock"))

    def test_returns_keep_the_unit_of_the_invoice(self):
        ret = proc(read("modSales"), "PostSalesReturn")
        self.assertIn("factors(n) = Nz(rs!UnitFactor, 1)", ret)
        self.assertIn("qtys(i) * factors(i), TT_SALES_RETURN, costs(i) / factors(i)", ret)
        pret = proc(read("modPurchases"), "PostPurchaseReturn")
        self.assertIn("pieces(n) = rs!ReturnQty * factors(n)", pret)
        self.assertIn("rs!Quantity * factors(n), reclaim)", pret)

    def test_purchase_moves_pieces(self):
        body = proc(read("modPurchases"), "PostPurchaseFromCart")
        self.assertIn('CalcLine(i, "QTY") * factors(i), reclaim)', body)
        self.assertIn("UPDATE ProductUnits SET PurchasePrice = ", body, "the last carton price")

    def test_scan_finds_a_unit_barcode(self):
        self.assertIn("If Not FindProductCode(Code, pid, unit) Then", proc(read("modPOS"), "ScanCode"))
        self.assertIn("If Not FindProductCode(Code, pid, unit) Then", proc(read("modPurchaseScreens"), "PurScanCode"))
        self.assertIn("u.Barcode = ", proc(VBA, "FindProductCode"))

    def test_carts_have_a_unit(self):
        local = proc(read("modPOS"), "EnsureLocalTables")
        self.assertIn('ALTER TABLE tmpPOSLines ADD COLUMN UnitID LONG, UnitFactor CURRENCY', local)
        self.assertIn('ALTER TABLE tmpPurchaseLines ADD COLUMN UnitID LONG, UnitFactor CURRENCY', local)
        for form, handler in (("frmPOSLines", "LineChanged"), ("frmPurchaseLines", "PurLineChanged")):
            m = MODELS[form]
            combo = next(c for c in m.controls if c.name == "UnitID")
            self.assertEqual(combo.kind, "combo", form)
            self.assertIn(f'    {handler} Me, "UnitID"', m.code, form)
        self.assertIn('If FieldName = "UnitID" Then', proc(read("modPOS"), "LineChanged"))
        self.assertIn('If FieldName = "UnitID" Then', proc(read("modPurchaseScreens"), "PurLineChanged"))

    def test_note_invoices_stay_in_pieces(self):
        self.assertIn("StockNoteID > 0 And factors(n) <> 1", proc(read("modSales"), "PostSaleFromCart"))
        self.assertIn("StockNoteID > 0 And factors(n) <> 1", proc(read("modPurchases"), "PostPurchaseFromCart"))

    def test_e_invoice_lines_take_the_unit_of_the_line(self):
        self.assertIn("IIf(v.UnitID Is Null, u.ZatcaUnitCode, v.ZatcaUnitCode)", read("modZatcaXml"))
        for name in ("modEtaInvoice", "modEtaReceipt"):
            self.assertIn("IIf(v.UnitID Is Null, u.EtaUnitCode, v.EtaUnitCode)", read(name), name)

    def test_validation(self):
        self.assertIn("ValidateProductUnit(frm)", read("modForms"))
        self.assertIn("ProductBarcodeProblem(frm!Barcode.Value", read("modForms"))
        problem = proc(VBA, "ProductUnitProblem")
        self.assertIn("IsBaseUnit(CLng(ProductID), UnitID)", problem)
        self.assertIn("SELECT ProductID FROM Products WHERE Barcode = ", problem)

    def test_in_access_test_is_listed(self):
        self.assertIn('"TestProductUnits"', read("modTestAll"))
        test = proc(VBA, "TestProductUnits")
        self.assertIn("ws.Rollback", test)
        for call in ("PostPurchaseFromCart(", "PostSaleFromCart(", "PostSalesReturn(", "PostPurchaseReturn("):
            self.assertIn(call, test)


class Static_modProductUnits(VbaModuleChecks, unittest.TestCase):
    module_name = "modProductUnits"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
