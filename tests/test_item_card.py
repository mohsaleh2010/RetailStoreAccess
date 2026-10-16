"""The item card and the warehouse reports (docs/60-Item-Card-Warehouse-Reports.md, plan 2D):
  * replayed on the SQLite mirror (tools/sim.py): ItemCardQuery gives the opening line (quantity and value) and
    every move of a product in one warehouse or all, with the other party of each move; the valuation of each
    warehouse, the transfers in transit and the posted count differences;
  * the report centre (warehouse filter, the card opened from the product screen) and one open count per warehouse."""
import os
import re
import sys
import unittest
from decimal import Decimal as D

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools"))
sys.path.insert(0, HERE)

import forms as F
import queries as Q
import reports_catalog as RC
from access_sqlite import AccessOnSqlite, day
from helpers import VbaModuleChecks
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}
ENTRIES = {r.key: r for r in F.REPORTS}


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


class ItemCardTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        s = self.s = Store(self.db.con, now=day(20, 9))
        s.user = 1
        s.c.execute("INSERT INTO Warehouses (WarehouseID, WarehouseCode, WarehouseName, WarehouseType) "
                    "VALUES (2, 'WH-2', 'فرع', 'BRANCH')")
        self.pid = s.insert("Products", ProductCode="TEST-K1", ProductName="TEST صنف", CategoryID=1, UnitID=1,
                            SellingPrice=20, CurrentQuantity=0, AverageCost=0, PurchasePrice=0)
        self.hall = s.insert("StockParties", PartyCode="SP-2", PartyName="صالة الإنتاج", PartyType="PRODUCTION")
        s.apply_stock(self.pid, D(10), 5, D(6), "MANUAL", 0, "ADJ-1", recalc=True, warehouse=1)    # day 20
        s.now = day(10, 9)
        s.apply_stock(self.pid, D(10), 5, D(9), "MANUAL", 0, "ADJ-2", recalc=True, warehouse=1)    # avg 7.5
        s.now = day(5, 9)
        s.send_transfer(s.stock_transfer(1, 2, [(self.pid, 4)]))                                  # 4 x 7.5
        s.post_stock_document(s.stock_document("ISSUE", 2, self.hall, [(self.pid, 1)]))          # 1 x 7.5

    def card(self, start_days_ago, warehouse=0):
        self.db.params.update(PeriodStart=day(start_days_ago), PeriodEnd=day(-1), ProductID=self.pid,
                              WarehouseID=warehouse)
        rows = self.s.c.execute("SELECT SortKey, NetQty, NetValue, QtyIn, QtyOut, ValueIn, ValueOut, WarehouseName, "
                                "PartyName, MovementType, MoveID FROM ItemCardQuery ORDER BY SortKey, MovementDate, "
                                "MoveID").fetchall()
        return [(r[0], D(str(r[1])), D(str(r[2])).quantize(D("0.01"))) + tuple(r[3:]) for r in rows]

    def test_all_warehouses(self):
        rows = self.card(15)
        self.assertEqual(rows[0][:3], (0, 10, 60), "opening line: quantity and value before the period")
        self.assertEqual(len(rows), 1 + 4)       # ADJ-2, transfer out and in, the issue
        self.assertEqual(sum(r[1] for r in rows), 19)
        self.assertEqual(sum(r[2] for r in rows), D("142.50"), "19 x 7.5: the value balance follows the average")

    def test_one_warehouse(self):
        rows = self.card(30, 2)
        self.assertEqual(rows[0][:3], (0, 0, 0))
        self.assertEqual([(r[1], r[2]) for r in rows[1:]], [(4, D("30.00")), (-1, D("-7.50"))])
        self.assertEqual({r[7] for r in rows[1:]}, {"فرع"})
        main = self.card(30, 1)
        self.assertEqual(sum(r[1] for r in main), 16)

    def test_in_out_columns(self):
        rows = self.card(30, 1)
        out = [r for r in rows if r[1] < 0][0]
        self.assertEqual((out[3], out[4], out[5], out[6]), (0, 4, 0, 30))

    def test_party_of_each_move(self):
        rows = self.card(30)
        parties = [(r[1], r[8]) for r in rows if r[0] == 1]
        self.assertIn((-4, "فرع"), parties, "a transfer out shows the receiving warehouse")
        self.assertIn((4, "المخزن الرئيسي"), parties, "a transfer in shows the sending warehouse")
        self.assertIn((-1, "صالة الإنتاج"), parties, "a voucher shows its party")
        self.assertIn((10, None), parties, "a manual move has no party")

    def test_sale_shows_the_customer(self):
        s = self.s
        s.sale([(self.pid, 2)], customer=1)
        name = s.one("SELECT CustomerName FROM Customers WHERE CustomerID = 1")
        rows = self.card(30)
        self.assertIn((-2, name), [(r[1], r[8]) for r in rows])

    def test_card_value_matches_the_warehouse_valuation(self):
        rows = self.card(30, 2)
        val = self.s.one("SELECT StockCostValue FROM StockValueByWarehouseQuery WHERE WarehouseID = 2")
        self.assertEqual(sum(r[2] for r in rows), D(str(val)).quantize(D("0.01")))
        self.assertEqual(self.s.one("SELECT StockAccount FROM StockValueByWarehouseQuery WHERE WarehouseID = 2"),
                         140002)
        self.assertEqual(self.s.one("SELECT TotalQuantity FROM StockValueByWarehouseQuery WHERE WarehouseID = 1"), 16)


class WarehouseReportTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        s = self.s = Store(self.db.con, now=day(3, 9))
        s.user = 1
        s.c.execute("INSERT INTO Warehouses (WarehouseID, WarehouseCode, WarehouseName, WarehouseType, ConfirmReceipt) "
                    "VALUES (2, 'WH-2', 'فرع', 'BRANCH', 1)")
        self.pid = s.insert("Products", ProductCode="TEST-K2", ProductName="TEST صنف", CategoryID=1, UnitID=1,
                            SellingPrice=20, CurrentQuantity=0, AverageCost=0, PurchasePrice=0)
        s.apply_stock(self.pid, D(10), 5, D(6), "MANUAL", 0, "ADJ-1", recalc=True, warehouse=1)

    def test_transfers_in_transit(self):
        s = self.s
        tid = s.stock_transfer(1, 2, [(self.pid, 3)])
        s.send_transfer(tid)
        row = s.c.execute("SELECT FromName, ToName, TransferValue, DaysInTransit FROM TransfersInTransitQuery "
                          "WHERE StockTransferID = ?", (tid,)).fetchone()
        self.assertEqual(row, ("المخزن الرئيسي", "فرع", 18, 3))
        s.receive_transfer_stock(tid)
        self.assertEqual(s.one("SELECT COUNT(*) FROM TransfersInTransitQuery"), 0)

    def test_posted_count_differences(self):
        s = self.s
        cid, lines, value = s.stock_count(1, {self.pid: 7})
        row = s.c.execute("SELECT WarehouseName, ShortQty, ExtraQty, DiffValue FROM StockCountDiffQuery "
                          "WHERE StockCountID = ? AND ProductID = ?", (cid, self.pid)).fetchone()
        self.assertEqual(row, ("المخزن الرئيسي", 3, 0, -18))
        self.assertEqual(s.one("SELECT WarehouseName FROM StockCountQuery WHERE StockCountID = ? AND ProductID = ?",
                               cid, self.pid), "المخزن الرئيسي")


class ReportCentreTests(unittest.TestCase):

    def test_catalogue(self):
        for key, query, needs in (("ITEM_CARD", "ItemCardQuery", "PRW"), ("STOCK_WAREHOUSE", "StockByWarehouseQuery", "w"),
                                  ("STOCK_VALUE_WAREHOUSE", "StockValueByWarehouseQuery", ""),
                                  ("TRANSFERS_IN_TRANSIT", "TransfersInTransitQuery", ""),
                                  ("STOCK_COUNT_DIFF", "StockCountDiffQuery", "Dw")):
            self.assertEqual((ENTRIES[key].query, ENTRIES[key].needs), (query, needs), key)
            self.assertIn(key, {sp.key for sp in RC.LIST_SPECS}, key)
        names = {q.name: q for q in Q.QUERIES}
        self.assertEqual(set(names["ItemCardQuery"].params), {"PeriodStart", "PeriodEnd", "ProductID", "WarehouseID"})
        for key in ("STOCK_WAREHOUSE", "STOCK_COUNT_DIFF"):        # a column filter needs the column
            self.assertIn("WarehouseID", names[ENTRIES[key].query].sql.split("FROM")[0], key)

    def test_warehouse_choice(self):
        code = read("modScreens")
        self.assertIn("cboWarehouse", {c.name for c in MODELS["frmReportCenter"].controls})
        selected = proc(code, "ReportSelected")
        self.assertIn('frm!cboWarehouse.Enabled = HasNeed(needs, "W") Or HasNeed(needs, "w")', selected)
        prepare = proc(code, "PrepareReport")
        self.assertIn('AddFilter where, frm!cboWarehouse, "WarehouseID"', prepare)
        self.assertIn('SetQueryParam "WarehouseID", CLng(Nz(frm!cboWarehouse.Value, 0))', prepare)
        load = proc(code, "ReportCenterLoad")
        self.assertIn('Split(CStr(frm.OpenArgs), "|")', load)
        self.assertIn("frm!cboProduct.Value = CLng(Val(args(1)))", load)

    def test_product_screen_opens_the_card(self):
        code = "\n".join(MODELS["frmProducts"].code)
        self.assertIn('OpenScreen "frmReportCenter", 0, "ITEM_CARD|" & Nz(Me!ProductID, 0)', code)


class StockCountPerWarehouseTests(unittest.TestCase):

    def test_one_open_count_per_warehouse(self):
        create = proc(read("modPurchases"), "CreateStockCount")
        self.assertIn("WHERE Status = 'OPEN' AND WarehouseID = \" & WarehouseID", create)
        self.assertLess(create.index("WarehouseID = CurrentWarehouseID()"), create.index("Status = 'OPEN'"),
                        "the warehouse is known before the open count is looked for")
        load = proc(read("modPurchaseScreens"), "StockCountLoad")
        self.assertIn("Status = 'OPEN' AND WarehouseID = \" & _", load)
        self.assertIn("w.WarehouseName", proc(read("modPurchaseScreens"), "ShowStockCount"))

    def test_in_access_test(self):
        self.assertIn('"TestStockReports"', read("modTestAll"))
        test = proc(read("modWarehouses"), "TestStockReports")
        for part in ("ItemCardQuery", "StockValueByWarehouseQuery", "TransfersInTransitQuery", "ws.Rollback"):
            self.assertIn(part, test)


class StaticModule(VbaModuleChecks, unittest.TestCase):
    module_name = "modWarehouses"
    vba = read("modWarehouses")


if __name__ == "__main__":
    unittest.main()
