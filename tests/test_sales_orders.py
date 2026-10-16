"""Customer orders with stock reservation (modSalesOrders, frmSalesOrder, docs/62-Sales-Orders.md, plan 4A):
  * replayed on the SQLite mirror (tools/sim.py): an approved order reserves its quantities in its warehouse, so
    a sale or another order cannot take them; delivery notes made from the order deliver it in parts, free the
    reservation, and the order follows them (partly delivered, delivered, invoiced); no note delivers more than
    what is left; cancelling frees what is left (closed when part was delivered);
  * the rules of the VBA code (CheckStockAvailable less the reservations, the notes keep the order in step), the
    screen and the reports."""
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
from forms_accounting import HUB_TILES
from helpers import VbaModuleChecks
from schema import SCREEN_LIST, table
from sim import SimError, Store

MODELS = {m.name: m for m in F.all_forms()}
ENTRIES = {r.key: r for r in F.REPORTS}


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


VBA = read("modSalesOrders")


class SalesOrderTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        s = self.s = Store(self.db.con, now=day(5, 9))
        s.user = 1
        self.pid = s.insert("Products", ProductCode="TEST-O1", ProductName="TEST صنف الطلبيات", CategoryID=1, UnitID=1,
                            SellingPrice=20, CurrentQuantity=0, AverageCost=0, PurchasePrice=8)
        self.cust = s.insert("Customers", CustomerName="TEST عميل الطلبيات", AllowCredit=1, CreditLimit=0,
                             CurrentBalance=0, IsSystem=0)
        s.apply_stock(self.pid, D(10), 5, D(6), "MANUAL", 0, "ADJ-1", recalc=True, warehouse=1)

    def order(self, q, approve=True):
        oid = self.s.sales_order(self.cust, 1, [(self.pid, q, 20)])
        if approve:
            self.s.approve_order(oid)
        return oid

    def reserved(self):
        return self.s.reserved(self.pid, 1)

    def test_approval_reserves_the_quantity(self):
        s = self.s
        oid = self.order(6, approve=False)
        self.assertEqual(self.reserved(), 0, "a draft reserves nothing")
        s.approve_order(oid)
        self.assertEqual((s.order_status(oid), self.reserved()), ("APPROVED", 6))
        self.assertEqual(s.one("SELECT ReservedQty FROM qryReservedStock WHERE ProductID = ?", self.pid), 6)

    def test_reserved_stock_is_not_sold_to_others(self):
        s = self.s
        self.order(6)
        with self.assertRaises(SimError):
            s.sale([(self.pid, 5)], customer=self.cust)
        s.sale([(self.pid, 4)], customer=self.cust)                   # the free 4
        self.assertEqual(s.stock(self.pid), 6)

    def test_an_order_beyond_the_free_stock_is_not_approved(self):
        s = self.s
        self.order(6)
        other = self.order(5, approve=False)
        with self.assertRaises(SimError):
            s.approve_order(other)
        self.assertEqual(s.order_status(other), "DRAFT")

    def test_partial_delivery_then_the_rest(self):
        s = self.s
        oid = self.order(6)
        nid = s.order_note(oid)
        self.assertEqual(s.one("SELECT Sum(Quantity) FROM StockNoteLines WHERE StockNoteID = ?", nid), 6,
                         "the note starts with what is left of the order")
        s.c.execute("UPDATE StockNoteLines SET Quantity = 4 WHERE StockNoteID = ?", (nid,))
        s.post_stock_note(nid)
        self.assertEqual((s.order_status(oid), self.reserved(), s.stock(self.pid)), ("PARTIAL", 2, 6))
        s.invoice_note(nid, credit=True)
        self.assertEqual(s.order_status(oid), "PARTIAL")
        rest = s.order_note(oid)
        self.assertEqual(s.one("SELECT Sum(Quantity) FROM StockNoteLines WHERE StockNoteID = ?", rest), 2)
        s.post_stock_note(rest)
        self.assertEqual((s.order_status(oid), self.reserved()), ("DELIVERED", 0))
        s.invoice_note(rest, credit=True)
        self.assertEqual((s.order_status(oid), s.stock(self.pid)), ("INVOICED", 4))
        self.assertEqual(s.one("SELECT Sum(Quantity) FROM SalesInvoiceDetails AS d INNER JOIN StockNotes AS n "
                               "ON d.SalesInvoiceID = n.SalesInvoiceID WHERE n.SalesOrderID = ?", oid), 6)
        with self.assertRaises(SimError):
            s.order_note(oid)

    def test_the_order_reserves_for_its_own_note(self):
        s = self.s
        s.apply_stock(self.pid, D(-4), 6, D(6), "MANUAL", 0, "ADJ-2", warehouse=1)    # 6 left, all reserved
        oid = self.order(6)
        nid = s.order_note(oid)
        s.post_stock_note(nid)                                       # its own reservation is free for it
        self.assertEqual((s.order_status(oid), s.stock(self.pid)), ("DELIVERED", 0))

    def test_no_more_than_what_is_left(self):
        s = self.s
        oid = self.order(6)
        nid = s.order_note(oid)
        s.c.execute("UPDATE StockNoteLines SET Quantity = 7 WHERE StockNoteID = ?", (nid,))
        with self.assertRaises(SimError):
            s.post_stock_note(nid)
        self.assertEqual(s.one("SELECT Status FROM StockNotes WHERE StockNoteID = ?", nid), "DRAFT")

    def test_unposting_a_note_reserves_again(self):
        s = self.s
        oid = self.order(6)
        nid = s.order_note(oid)
        s.post_stock_note(nid)
        s.unpost_stock_note(nid)
        self.assertEqual((s.order_status(oid), self.reserved()), ("APPROVED", 6))

    def test_cancel_frees_the_rest(self):
        s = self.s
        oid = self.order(6)
        nid = s.order_note(oid)
        with self.assertRaises(SimError):
            s.cancel_order(oid)                                      # its draft note first
        s.c.execute("UPDATE StockNoteLines SET Quantity = 1 WHERE StockNoteID = ?", (nid,))
        s.post_stock_note(nid)
        s.cancel_order(oid)
        self.assertEqual((s.order_status(oid), self.reserved()), ("CLOSED", 0))
        other = self.order(3)
        s.cancel_order(other)
        self.assertEqual(s.order_status(other), "CANCELLED")

    def test_reports(self):
        s = self.s
        s.now = day(1, 9)
        oid = s.sales_order(self.cust, 1, [(self.pid, 6, 20)], required=day(3))
        s.approve_order(oid)
        nid = s.order_note(oid)
        s.c.execute("UPDATE StockNoteLines SET Quantity = 2 WHERE StockNoteID = ?", (nid,))
        s.post_stock_note(nid)
        row = s.c.execute("SELECT CustomerName, Quantity, DeliveredQty, RemainingQty, RemainingValue, StatusName, "
                          "DaysLate FROM PendingOrdersQuery WHERE SalesOrderID = ?", (oid,)).fetchone()
        self.assertEqual(row, ("TEST عميل الطلبيات", 6, 2, 4, 80, "مُسلَّمة جزئيًا", 3))
        row = s.c.execute("SELECT OnHandQty, ReservedQty, FreeQty FROM ReservedStockQuery WHERE ProductID = ?",
                          (self.pid,)).fetchone()
        self.assertEqual(row, (8, 4, 4))


class SalesOrderCodeTests(unittest.TestCase):

    def test_schema(self):
        orders = {f.name: f for f in table("SalesOrders").fields}
        self.assertIn('"CLOSED"', orders["Status"].rule)
        self.assertEqual(orders["WarehouseID"].default, "1")
        lines = {f.name: f for f in table("SalesOrderLines").fields}
        self.assertTrue(lines["SalesOrderID"].on_delete_cascade)
        self.assertEqual({f.name: f for f in table("StockNotes").fields}["SalesOrderID"].fk, "SalesOrders.SalesOrderID")
        self.assertFalse({f.name: f for f in table("StockNoteLines").fields}["SalesOrderLineID"].required,
                         "an existing back-end gets the new field")
        self.assertIn(("frmSalesOrder", "طلبيات العملاء", "المبيعات", "SALES_ORDERS", True, True, True), SCREEN_LIST)
        self.assertIn(("SALES_ORDERS",), [(r[0],) for r in table("Permissions").seed_rows])

    def test_reservation_in_the_stock_check(self):
        check = proc(read("modSales"), "CheckStockAvailable")
        self.assertIn("Optional ByVal SalesOrderID As Long = 0", check)
        self.assertIn("available = AvailableQty(ProductIDs(i), WarehouseID) - reserved", check)
        self.assertIn("ReservedQty(ProductIDs(i), WarehouseID, SalesOrderID)", check)
        reserved = proc(VBA, "ReservedQty")
        self.assertIn("If WarehouseID = 0 Then WarehouseID = CurrentWarehouseID()", reserved)
        self.assertIn("SalesOrderID <> ", reserved)

    def test_the_notes_keep_the_order_in_step(self):
        notes = read("modStockNotes")
        post = proc(notes, "PostStockNote")
        self.assertIn("msg = OrderNoteProblem(StockNoteID)", post)
        self.assertIn('CheckStockAvailable(ids, qtys, n, "إذن التسليم", wh, orderID)', post)
        self.assertLess(post.index("SyncSalesOrder db, orderID"), post.index("ws.CommitTrans"))
        unpost = proc(notes, "UnpostStockNote")
        self.assertLess(unpost.index("SyncSalesOrder db"), unpost.index("ws.CommitTrans"))
        self.assertIn("SyncSalesOrder db", proc(notes, "MarkNoteInvoiced"))
        sync = proc(VBA, "SyncSalesOrder")
        for st in ('"PARTIAL"', '"DELIVERED"', '"INVOICED"', '"PREPARING"'):
            self.assertIn(st, sync)

    def test_order_actions(self):
        approve = proc(VBA, "ApproveSalesOrder")
        self.assertIn('CheckStockAvailable(ids, qtys, n, "اعتماد الطلبية وحجز كمياتها", wh, SalesOrderID)', approve)
        note = proc(VBA, "CreateOrderDeliveryNote")
        self.assertIn('HasPermission("STOCK_NOTES")', note)
        self.assertIn("nl!SalesOrderLineID = rs!SalesOrderLineID", note)
        self.assertLess(note.index("ws.BeginTrans"), note.index("CreateStockNote("))
        cancel = proc(VBA, "CancelSalesOrder")
        self.assertIn('newStatus = "CLOSED"', cancel)
        self.assertIn("AuditEdited", cancel)
        delete = proc(VBA, "DeleteSalesOrder")
        self.assertLess(delete.index("AuditSnapshot"), delete.index("DELETE FROM SalesOrders"))
        self.assertIn('AuditDeleted "SALES_ORDER_DELETE"', delete)

    def test_screen(self):
        frm = MODELS["frmSalesOrder"]
        names = {c.name for c in frm.controls}
        for name in ("cboOrder", "cboCustomer", "cboWarehouse", "txtRequired", "subLines", "btnApprove", "btnPrepare",
                     "btnDeliver", "btnNotes", "btnUnapprove", "btnCancel", "btnDelete"):
            self.assertIn(name, names)
        self.assertIn("Activate", frm.form_events)
        lines = {c.name: c for c in MODELS["frmSalesOrderLines"].controls}
        self.assertEqual(lines["txtDelivered"].source, "=OrderLineDelivered([SalesOrderLineID])")
        self.assertEqual(F.SCREEN_PERMISSIONS["frmSalesOrder"], "SALES_ORDERS")
        self.assertIn("frmSalesOrder", [t[3] for t in HUB_TILES])
        self.assertIn("btnOrder", {c.name for c in MODELS["frmStockNote"].controls})
        self.assertIn('OpenScreen "frmStockNote", 0, noteID', proc(VBA, "SalesOrderAction"))

    def test_reports(self):
        self.assertEqual((ENTRIES["PENDING_ORDERS"].query, ENTRIES["PENDING_ORDERS"].needs,
                          ENTRIES["PENDING_ORDERS"].date_column), ("PendingOrdersQuery", "Dcw", "OrderDate"))
        self.assertEqual((ENTRIES["RESERVED_STOCK"].query, ENTRIES["RESERVED_STOCK"].needs), ("ReservedStockQuery", "w"))
        keys = {sp.key for sp in RC.LIST_SPECS}
        self.assertTrue({"PENDING_ORDERS", "RESERVED_STOCK"} <= keys)
        names = {q.name: q for q in Q.QUERIES}
        for key in ("PENDING_ORDERS", "RESERVED_STOCK"):          # the filters need their columns
            head = names[ENTRIES[key].query].sql.split("FROM")[0]
            self.assertIn("WarehouseID", head, key)
        self.assertIn("CustomerID", names["PendingOrdersQuery"].sql.split("FROM")[0])

    def test_in_access_test(self):
        self.assertIn('"TestSalesOrders"', read("modTestAll"))
        test = proc(VBA, "TestSalesOrders")
        for part in ("ws.Rollback", "ReservedQty", "CreateOrderDeliveryNote", "CancelSalesOrder"):
            self.assertIn(part, test)


class StaticModule(VbaModuleChecks, unittest.TestCase):
    module_name = "modSalesOrders"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
