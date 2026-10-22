"""Price lists (modPriceLists, frmPriceLists, frmPriceListItems, docs/73-Price-Lists.md):
  * the tables: PriceLists (discount rate), PriceListItems (unit, price, validity), Customers.PriceListID;
  * the mirror of the rules (tools/price_list_reference.py): the item of the unit, the piece item x the factor,
    the list discount, the valid item that starts last;
  * replayed on the demo data in SQLite: the report query of the lists;
  * the code: a new sale line, an order line and a delivery note line start at the customer's price, a line at
    that price is not a price change, the POS reprices its lines when the customer changes.
"""
import os
import re
import sys
import unittest
from datetime import date, timedelta
from decimal import Decimal as D

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools"))
sys.path.insert(0, HERE)

import demo_data as DD
import forms as F
import master_en
import price_list_reference as R
from helpers import VbaModuleChecks
from schema import SCREEN_LIST, table

TODAY = date(2026, 10, 22)


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


VBA = read("modPriceLists")


class SchemaTests(unittest.TestCase):

    def test_tables(self):
        f = {x.name: x for x in table("PriceLists").fields}
        self.assertEqual((f["DiscountPercent"].default, f["DiscountPercent"].rule), ("0", ">=0 And <1"))
        self.assertTrue(f["ListName"].required)
        self.assertTrue(next(i for i in table("PriceLists").indexes if i.name == "UX_ListName").unique)
        f = {x.name: x for x in table("PriceListItems").fields}
        self.assertEqual(f["PriceListID"].fk, "PriceLists.PriceListID")
        self.assertTrue(f["PriceListID"].on_delete_cascade, "the prices go with their list")
        self.assertEqual(f["ProductID"].fk, "Products.ProductID")
        self.assertFalse(f["ProductID"].on_delete_cascade)
        self.assertEqual(f["UnitID"].fk, "Units.UnitID")
        self.assertFalse(f["UnitID"].required, "Null = the smallest unit")
        self.assertEqual(f["UnitPrice"].rule, ">0")
        self.assertFalse(f["ValidFrom"].required or f["ValidTo"].required)
        self.assertIsNone(f["ValidFrom"].default, "no date = no limit")

    def test_customer_list(self):
        f = {x.name: x for x in table("Customers").fields}
        self.assertEqual(f["PriceListID"].fk, "PriceLists.PriceListID")
        self.assertFalse(f["PriceListID"].required, "an existing back-end is upgraded")
        customers = next(s for s in F.DATA_SCREENS if s.name == "frmCustomers")
        self.assertIn(("PriceListID", F.PRICE_LIST_ROWS), [(x.field, x.rows) for x in customers.fields if hasattr(x, "rows")])

    def test_screens(self):
        for name in ("frmPriceLists", "frmPriceListItems"):
            self.assertEqual(F.SCREEN_PERMISSIONS[name], "PRODUCTS")
            self.assertTrue(any(s[0] == name and s[3] == "PRODUCTS" for s in SCREEN_LIST), name)
            self.assertIn(name, master_en.SCREEN_TITLES_EN)
        lists = next(s for s in F.DATA_SCREENS if s.name == "frmPriceLists")
        self.assertIn(("btnListItems", "أسعار القائمة", "OpenPriceListItems Me"), lists.extra_buttons)
        products = next(s for s in F.DATA_SCREENS if s.name == "frmProducts")
        self.assertIn(("btnPriceLists", "قوائم الأسعار", 'OpenScreen "frmPriceLists"'), products.extra_buttons)

    def test_english_names(self):
        self.assertEqual(master_en.ENGLISH_NAMES["PriceLists"][:3], ("ListNameEn", "ListName", "PriceListID"))
        self.assertIn("[@PriceLists]", F.PRICE_LIST_ROWS)

    def test_report(self):
        rep = next(r for r in F.REPORTS if r.key == "PRICE_LISTS")
        self.assertEqual((rep.query, rep.report, rep.needs), ("PriceListQuery", "rptPriceLists", ""))


class MirrorTests(unittest.TestCase):
    """The numbers of TestPriceLists: a piece at 10, a carton of 12, a list with 10 % discount."""

    def price(self, items, unit=None, factor=1, usual=10, day=TODAY, list_id=1):
        return R.customer_price(items, list_id, D("0.1"), 7, unit, factor, usual, day)

    def test_without_a_list(self):
        self.assertEqual(self.price([], list_id=0), D("10.00"))

    def test_discount_for_a_product_without_a_price(self):
        self.assertEqual(self.price([]), D("9.00"))
        self.assertEqual(self.price([], unit=5, factor=12, usual=120), D("108.00"))

    def test_piece_price_and_carton(self):
        items = [R.Item(1, 1, 7, None, D(8))]
        self.assertEqual(self.price(items), D("8.00"))
        self.assertEqual(self.price(items, unit=5, factor=12, usual=120), D("96.00"), "piece item x 12, no discount")
        items.append(R.Item(2, 1, 7, 5, D(90)))
        self.assertEqual(self.price(items, unit=5, factor=12, usual=120), D("90.00"), "the carton's own item")

    def test_dates(self):
        d = TODAY
        items = [R.Item(1, 1, 7, None, D(8)),
                 R.Item(3, 1, 7, None, D(7), d - timedelta(2), d + timedelta(2)),
                 R.Item(4, 1, 7, None, D(6), d - timedelta(30), d - timedelta(10)),
                 R.Item(5, 1, 7, None, D(5), d + timedelta(5))]
        self.assertEqual(R.item_price(items, 1, 7, None, d), 7, "a promotion over the standing price")
        self.assertEqual(R.item_price(items, 1, 7, None, d - timedelta(20)), 6)
        self.assertEqual(R.item_price(items, 1, 7, None, d + timedelta(3)), 8, "after the promotion")
        self.assertEqual(R.item_price(items, 1, 7, None, d + timedelta(6)), 5, "the future price")
        self.assertIsNone(R.item_price(items, 2, 7, None, d), "another list")
        self.assertIsNone(R.item_price(items, 1, 7, 5, d), "a carton has no item")

    def test_problems(self):
        items = [R.Item(1, 1, 7, None, D(8))]
        self.assertEqual(R.item_problem(items, 1, 7, None, 0, None, None), "price")
        self.assertEqual(R.item_problem(items, 1, 7, None, 8, TODAY, TODAY - timedelta(1)), "dates")
        self.assertEqual(R.item_problem(items, 1, 7, None, 7, None, None), "same start")
        self.assertEqual(R.item_problem(items, 1, 7, None, 7, None, None, item_id=1), "", "the item itself")
        self.assertEqual(R.item_problem(items, 1, 7, None, 7, TODAY, None), "", "a promotion with a start date")
        self.assertEqual(R.item_problem(items, 1, 7, 5, 90, None, None), "", "another unit")


class ReplayTests(unittest.TestCase):
    """A wholesale list added to the demo data in the SQLite mirror of Access."""

    @classmethod
    def setUpClass(cls):
        cls.r = DD.simulate()
        c = cls.c = cls.r.db.con
        pid = cls.pid = cls.r.ids["prod"][1]
        c.execute("INSERT INTO Units (UnitName, IsActive) VALUES ('كرتونة', 1)")
        cls.carton = c.execute("SELECT Max(UnitID) FROM Units").fetchone()[0]
        c.execute("INSERT INTO ProductUnits (ProductID, UnitID, UnitFactor, SellingPrice, PurchasePrice, IsActive, "
                  "CreatedAt) VALUES (?, ?, 12, 0, 0, 1, '2026-10-01 00:00:00')", (pid, cls.carton))
        c.execute("INSERT INTO PriceLists (ListName, DiscountPercent, IsActive, CreatedAt) "
                  "VALUES ('جملة', 0.1, 1, '2026-10-01 00:00:00')")
        cls.list = c.execute("SELECT Max(PriceListID) FROM PriceLists").fetchone()[0]
        c.execute("INSERT INTO PriceLists (ListName, DiscountPercent, IsActive, CreatedAt) "
                  "VALUES ('موقوفة', 0, 0, '2026-10-01 00:00:00')")
        stopped = c.execute("SELECT Max(PriceListID) FROM PriceLists").fetchone()[0]
        far = "2099-01-01 00:00:00"
        for lst, unit, price, start, end in ((cls.list, None, 8, None, None), (cls.list, cls.carton, 90, None, None),
                                             (cls.list, None, 7, "2000-01-01 00:00:00", "2000-02-01 00:00:00"),
                                             (cls.list, None, 6, far, None), (stopped, None, 5, None, None)):
            c.execute("INSERT INTO PriceListItems (PriceListID, ProductID, UnitID, UnitPrice, ValidFrom, ValidTo, CreatedAt) "
                      "VALUES (?, ?, ?, ?, ?, ?, '2026-10-01 00:00:00')", (lst, pid, unit, price, start, end))
        cls.selling = c.execute("SELECT SellingPrice FROM Products WHERE ProductID = ?", (pid,)).fetchone()[0]

    def q(self, sql, *args):
        return self.c.execute(sql, args).fetchall()

    def test_report_rows(self):
        rows = self.q("SELECT UnitPrice, PriceUnit, UsualPrice, StatusName FROM PriceListQuery ORDER BY UnitPrice")
        self.assertEqual([r[0] for r in rows], [6, 7, 8, 90], "an inactive list is not shown")
        status = {r[0]: r[3] for r in rows}
        self.assertEqual((status[6], status[7], status[8]), ("لم يبدأ بعد", "انتهى سريانه", "ساري"))
        carton = rows[-1]
        self.assertEqual(carton[1], "كرتونة")
        self.assertAlmostEqual(carton[2], self.selling * 12, 2, "the usual carton price = piece price x 12")
        self.assertAlmostEqual(rows[2][2], self.selling, 2)
        self.assertIsNotNone(rows[2][1], "the smallest unit is named")

    def test_customer_link(self):
        cust = self.r.ids["cust"][1]
        self.c.execute("UPDATE Customers SET PriceListID = ? WHERE CustomerID = ?", (self.list, cust))
        self.assertEqual(self.q("SELECT l.PriceListID FROM Customers AS c INNER JOIN PriceLists AS l "
                                "ON c.PriceListID = l.PriceListID WHERE l.IsActive <> 0 AND c.CustomerID = ?", cust),
                         [(self.list,)])


class CodeTests(unittest.TestCase):

    def test_sale_price_goes_through_the_customer_list(self):
        body = proc(read("modProductUnits"), "UnitSalePrice")
        self.assertIn("Optional ByVal CustomerID As Long = 0", body)
        self.assertIn("UnitSalePrice = CustomerPrice(CustomerID, ProductID, UnitID, usual)", body)
        price = proc(VBA, "CustomerListPrice")
        self.assertIn("PriceListItemPrice(listID, ProductID, UnitID, Date)", price)
        self.assertIn("PriceListItemPrice(listID, ProductID, Null, Date)", price, "the piece item x the factor")
        self.assertIn("RoundMoney(CDec(v) * CDec(factor))", price)
        self.assertIn("RoundMoney(CDec(Price) * (1 - CDec(pct)))", proc(VBA, "ListDiscounted"))

    def test_item_that_starts_last(self):
        body = proc(VBA, "PriceListItemPrice")
        self.assertIn("SELECT TOP 1 UnitPrice FROM PriceListItems", body)
        self.assertIn("ORDER BY ValidFrom DESC, PriceListItemID DESC", body)
        self.assertIn('unitCond = "UnitID Is Null"', body)

    def test_list_price_is_not_a_price_change(self):
        body = proc(read("modSales"), "PostSaleFromCart")
        self.assertIn("If IsCustomerListPrice(CustomerID, rs!ProductID, rs!UnitID, rs!UnitPrice) Then", body)
        self.assertLess(body.index("IsCustomerListPrice("), body.index("IsListedUnitPrice("))
        changed = proc(read("modPOS"), "LineChanged")
        self.assertIn("IsCustomerListPrice(cust, sf!ProductID.Value, sf!UnitID.Value", changed)
        self.assertNotIn("SessionPrice(", changed, "the price put back is the customer's")

    def test_pos_prices_for_the_customer(self):
        pos = read("modPOS")
        add = proc(pos, "AddLine")
        self.assertIn("rs!UnitPrice = UnitSalePrice(ProductID, UnitID, CartCustomerID(frm))", add)
        self.assertIn("rs!ListPrice = rs!UnitPrice", add)
        self.assertIn('ALTER TABLE tmpPOSLines ADD COLUMN ListPrice CURRENCY', proc(pos, "EnsureLocalTables"))
        self.assertIn("RepriceCart frm", proc(pos, "CustomerChanged"))
        reprice = proc(pos, "RepriceCart")
        self.assertIn('"WHERE UnitPrice = ListPrice"', reprice, "a typed price stays")
        self.assertIn("dbOpenDynaset, dbSeeChanges", reprice)
        self.assertIn('If HasControl(frm, "cboCustomer") Then', proc(pos, "CartCustomerID"))

    def test_orders_and_notes(self):
        self.assertIn("OrderDefaultPrice = CustomerPrice(", proc(read("modSalesOrders"), "OrderDefaultPrice"))
        note = proc(read("modStockNotes"), "StockNoteLineBeforeUpdate")
        self.assertIn('If NoteKind(id) = "DELIVERY" Then', note)
        self.assertIn("CustomerPrice(NoteParty(id), CLng(frm!ProductID.Value), Null, frm!UnitPrice.Value)", note)

    def test_validation(self):
        self.assertIn("ValidatePriceListItem(frm)", read("modForms"))
        problem = proc(VBA, "PriceListItemProblem")
        self.assertIn("UnitFactorOf(CLng(ProductID), UnitID) <= 0", problem)
        self.assertIn("CDate(ValidTo) < CDate(ValidFrom)", problem)
        self.assertIn("LineUnitID(CLng(frm!ProductID.Value), frm!UnitID.Value)", proc(VBA, "ValidatePriceListItem"))

    def test_in_access_test_is_listed(self):
        self.assertIn('"TestPriceLists"', read("modTestAll"))
        test = proc(VBA, "TestPriceLists")
        self.assertIn("ws.Rollback", test)
        self.assertIn("PostSaleFromCart(", test)


class Static_modPriceLists(VbaModuleChecks, unittest.TestCase):
    module_name = "modPriceLists"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
