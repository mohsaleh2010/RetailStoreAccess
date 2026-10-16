"""Warehouses (modWarehouses, frmWarehouses, docs/57-Warehouses.md):
  * on the demo data replayed in the SQLite mirror: every move is in the main warehouse after the upgrade,
    the balance per warehouse adds up to the ledger of each product, a move in another warehouse shows there
    only, the report of balances per warehouse;
  * the rules of the VBA code: every document and move carries its warehouse, the checks read the stock of
    that warehouse, the stock count of one warehouse, the opening stock per warehouse;
  * schema, screens and registration."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
import demo_data as DD
import forms as F
import generate as G
import queries as Q
from schema import RULE_UPGRADES, SCREEN_LIST, table

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


VBA = read("modWarehouses")
SALES, PURCHASES, OPENING = read("modSales"), read("modPurchases"), read("modOpening")


class WarehouseStockTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.r = DD.simulate()
        cls.c = cls.r.db.con

    def test_every_move_in_the_main_warehouse(self):
        self.assertEqual(self.c.execute("SELECT COUNT(*) FROM InventoryTransactions WHERE WarehouseID <> 1").fetchone()[0], 0)
        self.assertGreater(self.c.execute("SELECT COUNT(*) FROM InventoryTransactions").fetchone()[0], 0)
        for t in ("SalesInvoices", "PurchaseInvoices"):
            self.assertEqual(self.c.execute(f"SELECT COUNT(*) FROM {t} WHERE WarehouseID <> 1").fetchone()[0], 0, t)

    def test_balances_add_up_to_the_ledger(self):
        per_wh = dict(self.c.execute('SELECT ProductID, Sum(StockQty) FROM "qryWarehouseStock" GROUP BY ProductID'))
        ledger = dict(self.c.execute('SELECT ProductID, LedgerQty FROM "qryProductLedger"'))
        self.assertEqual(per_wh.keys(), ledger.keys())
        for pid, qty in ledger.items():
            self.assertAlmostEqual(per_wh[pid], qty, places=3)

    def test_a_second_warehouse(self):
        c = self.c
        c.execute("INSERT INTO Warehouses (WarehouseCode, WarehouseName, WarehouseType) VALUES ('WH-2', 'فرع', 'BRANCH')")
        wh2 = c.execute("SELECT Max(WarehouseID) FROM Warehouses").fetchone()[0]
        pid = c.execute("SELECT Min(ProductID) FROM InventoryTransactions").fetchone()[0]
        try:
            c.execute("INSERT INTO InventoryTransactions (TransactionDate, ProductID, TransactionTypeID, Quantity, "
                      "UnitCost, QuantityAfter, ReferenceType, WarehouseID) VALUES ('2026-01-01 10:00:00', ?, 5, 7, "
                      "3, 0, 'MANUAL', ?)", (pid, wh2))
            rows = c.execute('SELECT WarehouseID, StockQty FROM "StockByWarehouseQuery" WHERE ProductID = ?',
                             (pid,)).fetchall()
            self.assertIn((wh2, 7), [(w, round(q, 3)) for w, q in rows])
            self.assertEqual({w for w, _ in rows} - {1, wh2}, set())
        finally:
            c.execute("DELETE FROM InventoryTransactions WHERE WarehouseID = ?", (wh2,))
            c.execute("DELETE FROM Warehouses WHERE WarehouseID = ?", (wh2,))


class WarehouseCodeTests(unittest.TestCase):

    def test_schema(self):
        t = table("Warehouses")
        self.assertEqual(t.seed_rows[0][:3], (1, "WH-1", "المخزن الرئيسي"))
        for name in ("InventoryTransactions", "SalesInvoices", "SalesReturns", "PurchaseInvoices", "PurchaseReturns",
                     "StockCounts", "OpeningStock"):
            f = next(x for x in table(name).fields if x.name == "WarehouseID")
            self.assertEqual((f.default, f.fk), ("1", "Warehouses.WarehouseID"), name)
        self.assertEqual([i.fields for i in table("OpeningStock").indexes if i.unique], [["ProductID", "WarehouseID"]])
        self.assertIsNone(next(x for x in table("Employees").fields if x.name == "WarehouseID").default)
        self.assertIn(("frmWarehouses", "المخازن", "المخزون", "PRODUCTS", True, True, True), SCREEN_LIST)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmWarehouses"], "PRODUCTS")
        self.assertIn("modWarehouses", G.STATIC_MODULES)
        self.assertIn('"TestWarehouses"', read("modTestAll"))
        self.assertIn("STOCK_WAREHOUSE", {e.key for e in F.REPORTS})
        self.assertNotIn(("Warehouses", "WarehouseType"), RULE_UPGRADES)

    def test_old_index_is_dropped(self):
        build = read("modBuildSchema")
        self.assertIn('DropOldIndex "OpeningStock", "UX_ProductID"', build)

    def test_every_move_carries_its_warehouse(self):
        apply = proc(SALES, "ApplyStockMovement")
        self.assertIn("If WarehouseID = 0 Then WarehouseID = CurrentWarehouseID()", apply)
        self.assertIn("rs!WarehouseID = WarehouseID", apply)
        self.assertIn("available = AvailableQty(ProductIDs(i), WarehouseID)", proc(SALES, "CheckStockAvailable"))
        self.assertIn("rs!WarehouseID = CurrentWarehouseID()", proc(SALES, "PostSaleFromCart"))
        self.assertIn("SELECT WarehouseID FROM SalesInvoices", proc(SALES, "PostSalesReturn"))
        purchase = proc(PURCHASES, "PostPurchaseFromCart")
        self.assertIn("rs!WarehouseID = WarehouseID", purchase)
        self.assertIn('invoiceID, invNo, "", True, WarehouseID', purchase)
        ret = proc(PURCHASES, "PostPurchaseReturn")
        self.assertIn("SELECT WarehouseID FROM PurchaseInvoices", ret)
        self.assertIn('retID, retNo, "", True, warehouse', ret)
        manual = proc(PURCHASES, "PostManualStock")
        self.assertIn('CheckStockAvailable(ids, qtys, 1, "خصم المخزون", WarehouseID)', manual)
        self.assertEqual(manual.count(", WarehouseID\n"), 2)
        count = proc(PURCHASES, "CreateStockCount")
        self.assertIn("rs!WarehouseID = WarehouseID", count)
        self.assertIn("i.WarehouseID = ", count)
        post = proc(PURCHASES, "PostStockCount")
        self.assertIn("sysQty = AvailableQty(rs!ProductID, warehouse)", post)
        self.assertIn('"", False, warehouse', post)

    def test_available_quantity(self):
        avail = proc(VBA, "AvailableQty")
        self.assertIn("WarehouseID = MAIN_WAREHOUSE_ID And OnlyMainWarehouse()", avail)
        self.assertIn("SELECT CurrentQuantity FROM Products", avail)
        self.assertIn("WarehouseStock(ProductID, WarehouseID)", avail)
        self.assertIn("AvailableQty(ProductID)", read("modPOS"))
        self.assertIn("AvailableQty(ProductID)", read("modTouchPOS"))

    def test_opening_stock_per_warehouse(self):
        self.assertIn("AND WarehouseID = \" & wh & \" AND OpeningStockID <>", proc(OPENING, "OpeningStockLineProblem"))
        self.assertIn('"", True, wh', proc(OPENING, "PostOpeningStock"))
        unpost = proc(OPENING, "UnpostOpeningStock")
        self.assertIn("remainQty = ", unpost)
        self.assertIn('" AND WarehouseID = " & wh', unpost)
        self.assertIn("rs.Fields(4).Value", proc(OPENING, "ImportOpeningStock"))
        self.assertIn("[كود المخزن]", proc(OPENING, "OpeningStockTemplateSql"))
        grid = {c.name for c in MODELS["frmOpeningStock"].controls}
        self.assertIn("WarehouseID", grid)

    def test_screens_use_the_warehouse(self):
        screens = read("modPurchaseScreens")
        self.assertIn("code, fx, Nz(frm!cboWarehouse.Value, 0))", screens)
        self.assertIn("refNo, Nz(frm!cboWarehouse.Value, 0))", screens)
        self.assertIn('CreateStockCount(frm!cboCategory.Value, "", newID, Nz(frm!cboWarehouse.Value, 0))', screens)
        for name in ("frmPurchaseInvoice", "frmInventory", "frmStockCount"):
            self.assertIn("cboWarehouse", {c.name for c in MODELS[name].controls}, name)
        users = {c.name for c in MODELS["frmUsers"].controls}
        self.assertIn("WarehouseID", users)

    def test_report_query(self):
        sql = Q.query("StockByWarehouseQuery").sql
        self.assertIn("[@Warehouses] AS w", sql)
        self.assertIn("q.StockQty <> 0", sql)


class StaticModule(VbaModuleChecks, unittest.TestCase):
    module_name = "modWarehouses"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
