"""Purchase orders and the purchase suggestion (modPurchaseOrders, frmPurchaseOrder, docs/74-Purchase-Orders.md):
  * the tables: PurchaseOrders / PurchaseOrderLines, the order links on the receipt notes, Products.ReorderQuantity;
  * the mirror of the rules (tools/purchase_order_reference.py): the suggested quantity, the status of an order;
  * replayed on the demo data in SQLite: what is received / left of each line, what is on order, the suggestion
    and the open orders;
  * the code: receipt notes of an order are checked and keep the order in step, the suggestion makes one draft per
    supplier.
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
import master_en
import purchase_order_reference as R
from helpers import VbaModuleChecks
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


VBA = read("modPurchaseOrders")


class SchemaTests(unittest.TestCase):

    def test_tables(self):
        f = {x.name: x for x in table("PurchaseOrders").fields}
        self.assertEqual(f["SupplierID"].fk, "Suppliers.SupplierID")
        self.assertEqual((f["WarehouseID"].fk, f["WarehouseID"].default), ("Warehouses.WarehouseID", "1"))
        self.assertIn('"PARTIAL","RECEIVED","INVOICED"', f["Status"].rule)
        f = {x.name: x for x in table("PurchaseOrderLines").fields}
        self.assertTrue(f["PurchaseOrderID"].on_delete_cascade)
        self.assertEqual(f["Quantity"].rule, ">0")

    def test_note_links(self):
        f = {x.name: x for x in table("StockNotes").fields}
        self.assertEqual(f["PurchaseOrderID"].fk, "PurchaseOrders.PurchaseOrderID")
        f = {x.name: x for x in table("StockNoteLines").fields}
        self.assertEqual(f["PurchaseOrderLineID"].fk, "PurchaseOrderLines.PurchaseOrderLineID")
        self.assertFalse(f["PurchaseOrderLineID"].required)

    def test_reorder_quantity(self):
        f = {x.name: x for x in table("Products").fields}
        self.assertEqual((f["ReorderQuantity"].default, f["ReorderQuantity"].rule), ("0", ">=0"))
        products = next(s for s in F.DATA_SCREENS if s.name == "frmProducts")
        self.assertIn("ReorderQuantity", [x.field for x in products.fields if hasattr(x, "field")])

    def test_screen_permission_and_tile(self):
        self.assertIn(("frmPurchaseOrder", "أوامر الشراء", "المشتريات", "PURCHASE_ORDERS", True, True, True), SCREEN_LIST)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmPurchaseOrder"], "PURCHASE_ORDERS")
        self.assertIn("PURCHASE_ORDERS", [r[0] for r in table("Permissions").seed_rows])
        self.assertIn("PURCHASE_ORDERS", master_en.PERMISSION_NAMES_EN)
        inventory = next(m for m in MODELS.values() if m.name == "frmInventory")
        self.assertTrue(any(c.name == "btnPurchaseOrders" for c in inventory.controls))
        self.assertIn(("PURCHASE_ORDER", "PO-", 1, 6, "أوامر الشراء"), table("Sequences").seed_rows)
        m = MODELS["frmPurchaseOrder"]
        for b in ("btnApprove", "btnReceive", "btnSuggest", "btnCancel", "btnDelete"):
            self.assertTrue(any(c.name == b for c in m.controls), b)

    def test_reports(self):
        keys = {r.key: r for r in F.REPORTS}
        self.assertEqual(keys["PURCHASE_SUGGEST"].query, "PurchaseSuggestQuery")
        self.assertEqual((keys["PENDING_PURCHASE_ORDERS"].query, keys["PENDING_PURCHASE_ORDERS"].needs),
                         ("PendingPurchaseOrdersQuery", "Dw"))


class MirrorTests(unittest.TestCase):

    def test_suggest_qty(self):
        self.assertEqual(R.suggest_qty(4, 0, 10, 0), 16, "2 x 10 - 4")
        self.assertEqual(R.suggest_qty(4, 6, 10, 0), 10, "what is on order counts")
        self.assertEqual(R.suggest_qty(11, 0, 10, 0), 0, "above the level")
        self.assertEqual(R.suggest_qty(10, 0, 10, 0), 10, "at the level")
        self.assertEqual(R.suggest_qty(0, 0, 5, 50), 50, "the reorder quantity")
        self.assertEqual(R.suggest_qty(0, 0, 0, 50), 0, "no level, no suggestion")
        self.assertEqual(R.suggest_qty(-3, 0, 10, 0), 23, "a negative stock")

    def test_status(self):
        self.assertEqual(R.status_for(0, 0, 16), "APPROVED")
        self.assertEqual(R.status_for(10, 10, 6), "PARTIAL")
        self.assertEqual(R.status_for(16, 10, 0), "RECEIVED")
        self.assertEqual(R.status_for(16, 16, 0), "INVOICED")


class ReplayTests(unittest.TestCase):
    """An approved order of 16 + 50 with a receipt of 10 added to the demo data in the SQLite mirror of Access."""

    @classmethod
    def setUpClass(cls):
        cls.r = DD.simulate()
        c = cls.c = cls.r.db.con
        cls.sup = cls.r.ids["sup"][1]
        cls.p1, cls.p2 = cls.r.ids["prod"][1], cls.r.ids["prod"][2]
        c.execute("UPDATE Products SET MinimumQuantity = CurrentQuantity + 6, ReorderQuantity = 0, SupplierID = ? "
                  "WHERE ProductID = ?", (cls.sup, cls.p1))
        cls.stock1 = c.execute("SELECT CurrentQuantity FROM Products WHERE ProductID = ?", (cls.p1,)).fetchone()[0]
        c.execute("INSERT INTO PurchaseOrders (OrderNumber, OrderDate, SupplierID, WarehouseID, ExpectedDate, Status, "
                  "EmployeeID, CreatedAt) VALUES ('PO-T1', ?, ?, 1, '2026-09-01 00:00:00', 'APPROVED', 1, ?)",
                  (NOW, cls.sup, NOW))
        cls.po = c.execute("SELECT Max(PurchaseOrderID) FROM PurchaseOrders").fetchone()[0]
        for pid, q in ((cls.p1, 4), (cls.p2, 50)):
            c.execute("INSERT INTO PurchaseOrderLines (PurchaseOrderID, ProductID, Quantity, UnitCost, CreatedAt) "
                      "VALUES (?, ?, ?, 8, ?)", (cls.po, pid, q, NOW))
        line2 = c.execute("SELECT PurchaseOrderLineID FROM PurchaseOrderLines WHERE ProductID = ? AND PurchaseOrderID = ?",
                          (cls.p2, cls.po)).fetchone()[0]
        c.execute("INSERT INTO StockNotes (NoteNumber, NoteType, NoteDate, WarehouseID, SupplierID, PurchaseOrderID, Status, "
                  "EmployeeID, CreatedAt) VALUES ('GRN-T1', 'RECEIPT', ?, 1, ?, ?, 'INVOICED', 1, ?)",
                  (NOW, cls.sup, cls.po, NOW))
        note = c.execute("SELECT Max(StockNoteID) FROM StockNotes").fetchone()[0]
        c.execute("INSERT INTO StockNoteLines (StockNoteID, ProductID, Quantity, UnitPrice, UnitCost, PurchaseOrderLineID, "
                  "CreatedAt) VALUES (?, ?, 10, 8, 8, ?, ?)", (note, cls.p2, line2, NOW))

    def q(self, sql, *args):
        return self.c.execute(sql, args).fetchall()

    def test_line_status(self):
        rows = dict((r[0], r[1:]) for r in self.q(
            "SELECT ProductID, Quantity, ReceivedQty, InvoicedQty, RemainingQty, OrderStatus "
            "FROM qryPurchaseOrderLineStatus WHERE PurchaseOrderID = ?", self.po))
        self.assertEqual(rows[self.p2], (50, 10, 10, 40, "APPROVED"))
        self.assertEqual(rows[self.p1], (4, 0, 0, 4, "APPROVED"))

    def test_on_order_and_suggestion(self):
        on = dict(self.q("SELECT ProductID, OpenQty FROM qryOnOrderStock"))
        self.assertEqual((on[self.p1], on[self.p2]), (4, 40))
        rows = self.q("SELECT CurrentQuantity, OnOrderQty, MinimumQuantity, SuggestQty, SupplierName "
                      "FROM PurchaseSuggestQuery WHERE ProductID = ?", self.p1)
        self.assertEqual(len(rows), 1, "stock + 4 on order <= stock + 6")
        cur, on_order, minimum, suggest, name = rows[0]
        self.assertEqual(suggest, float(R.suggest_qty(cur, on_order, minimum, 0)))
        self.assertEqual(suggest, self.stock1 + 8)
        self.assertIsNotNone(name)

    def test_suggestion_matches_the_mirror_for_every_product(self):
        for pid, cur, on_order, minimum, reorder, suggest in self.q(
                "SELECT ProductID, CurrentQuantity, OnOrderQty, MinimumQuantity, ReorderQuantity, SuggestQty "
                "FROM PurchaseSuggestQuery"):
            self.assertEqual(D(str(suggest)), R.suggest_qty(cur, on_order, minimum, reorder), pid)
            self.assertGreater(suggest, 0, pid)

    def test_pending_orders(self):
        rows = self.q("SELECT ProductName, RemainingQty, RemainingValue, DaysLate FROM PendingPurchaseOrdersQuery "
                      "WHERE PurchaseOrderID = ? ORDER BY RemainingQty", self.po)
        self.assertEqual([(r[1], r[2]) for r in rows], [(4, 32), (40, 320)])
        self.assertTrue(all(r[3] > 0 for r in rows), "past the expected date")
        self.c.execute("UPDATE PurchaseOrders SET Status = 'CLOSED' WHERE PurchaseOrderID = ?", (self.po,))
        try:
            self.assertEqual(self.q("SELECT COUNT(*) FROM PendingPurchaseOrdersQuery WHERE PurchaseOrderID = ?",
                                    self.po)[0][0], 0)
            self.assertEqual(self.q("SELECT COUNT(*) FROM qryOnOrderStock WHERE ProductID = ?", self.p2)[0][0], 0)
        finally:
            self.c.execute("UPDATE PurchaseOrders SET Status = 'APPROVED' WHERE PurchaseOrderID = ?", (self.po,))


class CodeTests(unittest.TestCase):

    def test_notes_keep_the_order_in_step(self):
        notes = read("modStockNotes")
        self.assertIn("msg = PurchaseOrderNoteProblem(StockNoteID)", proc(notes, "PostStockNote"))
        for name in ("PostStockNote", "UnpostStockNote", "MarkNoteInvoiced"):
            self.assertIn("SyncPurchaseOrder db, NotePurchaseOrderID(StockNoteID)", proc(notes, name), name)
        self.assertIn('OpenScreen "frmPurchaseOrder", 0, orderID', proc(read("modSalesOrders"), "StockNoteOpenOrder"))

    def test_receipt_note_from_the_order(self):
        body = proc(VBA, "CreateOrderReceiptNote")
        self.assertIn('CreateStockNote("RECEIPT", wh, sup,', body)
        self.assertIn("nl!PurchaseOrderLineID = rs!PurchaseOrderLineID", body)
        self.assertIn("RemainingQty > 0", body)
        self.assertIn("rs!LineQty > ln!RemainingQty", proc(VBA, "PurchaseOrderNoteProblem"))

    def test_sync_uses_the_status_rule(self):
        self.assertIn("newStatus = PurchaseOrderStatusFor(received, invoiced, remaining)", proc(VBA, "SyncPurchaseOrder"))

    def test_suggestion(self):
        body = proc(VBA, "CreateSuggestedOrders")
        self.assertIn("FROM PurchaseSuggestQuery AS q", body)
        self.assertIn("WHERE o.Status = 'DRAFT'", body, "a product already in a draft is skipped")
        self.assertIn('NotePartyIsActive("RECEIPT", rs!SupplierID)', body)
        self.assertIn("ws.Rollback", body)

    def test_deletes_are_audited(self):
        self.assertIn('AuditDeleted "PURCHASE_ORDER_DELETE"', proc(VBA, "DeletePurchaseOrder"))
        self.assertIn('AuditDeleted "PURCHASE_ORDER_LINE_DELETE"', proc(VBA, "PurchaseOrderRemoveLine"))

    def test_in_access_test_is_listed(self):
        self.assertIn('"TestPurchaseOrders"', read("modTestAll"))
        test = proc(VBA, "TestPurchaseOrders")
        self.assertIn("ws.Rollback", test)
        for call in ("CreateSuggestedOrders(", "ApprovePurchaseOrder(", "CreateOrderReceiptNote(", "PostStockNote(",
                     "InvoiceStockNote(", "UnpostStockNote(", "CancelPurchaseOrder("):
            self.assertIn(call, test)


class Static_modPurchaseOrders(VbaModuleChecks, unittest.TestCase):
    module_name = "modPurchaseOrders"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
