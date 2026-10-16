"""Delivery and receipt notes tied to invoices (modStockNotes, frmStockNote, docs/61-Delivery-Receipt-Notes.md):
  * replayed on the SQLite mirror (tools/sim.py): a delivery note takes the goods out at the average cost onto
    1430, its sales invoice does not move the stock again and clears 1430; a receipt note brings the goods in at
    the purchase price onto 2150 (updating the average cost), its purchase invoice clears 2150; a note gets one
    invoice and is not unposted after it;
  * the rules of the VBA code (the invoice engines skip the stock moves of a note) and the screen."""
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
from access_sqlite import AccessOnSqlite, day
from helpers import VbaModuleChecks
from schema import ACCOUNT_TREE, SCREEN_LIST, table
from sim import SimError, Store

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


VBA = read("modStockNotes")


class StockNoteTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        s = self.s = Store(self.db.con, now=day(5, 9))
        s.user = 1
        self.pid = s.insert("Products", ProductCode="TEST-N1", ProductName="TEST صنف", CategoryID=1, UnitID=1,
                            SellingPrice=20, CurrentQuantity=0, AverageCost=0, PurchasePrice=8)
        self.cust = s.insert("Customers", CustomerName="TEST عميل الأذونات", AllowCredit=1, CreditLimit=0,
                             CurrentBalance=0, IsSystem=0)
        self.sup = s.insert("Suppliers", SupplierName="TEST مورد الأذونات")
        s.apply_stock(self.pid, D(10), 5, D(6), "MANUAL", 0, "ADJ-1", recalc=True, warehouse=1)

    def balance(self, code):
        return D(str(self.s.one("SELECT Coalesce(Sum(Debit - Credit), 0) FROM JournalLines WHERE AccountCode = ?",
                                code))).quantize(D("0.01"))

    def moves(self):
        return self.s.one("SELECT COUNT(*) FROM InventoryTransactions WHERE ProductID = ?", self.pid)

    def test_receipt_note_then_purchase_invoice(self):
        s = self.s
        nid = s.stock_note("RECEIPT", 1, self.sup, [(self.pid, 10, 10)])
        s.post_stock_note(nid)
        self.assertEqual((s.stock(self.pid), s.avg(self.pid)), (20, 8))
        s.sync_journal()
        self.assertEqual((self.balance(2150), self.balance(1400)), (-100, 160))
        moves = self.moves()
        inv = s.invoice_note(nid, credit=True)
        self.assertEqual((s.stock(self.pid), self.moves()), (20, moves), "the invoice does not move the stock again")
        s.sync_journal()
        self.assertEqual(self.balance(2150), 0, "the invoice clears the goods received not invoiced")
        self.assertEqual(self.balance(1400), 160)
        self.assertEqual(self.balance(2100), D("-115.00"), "the supplier owes the invoice with VAT")
        self.assertEqual(s.one("SELECT Status FROM StockNotes WHERE StockNoteID = ?", nid), "INVOICED")
        self.assertEqual(s.one("SELECT PurchaseInvoiceID FROM StockNotes WHERE StockNoteID = ?", nid), inv)

    def test_delivery_note_then_sales_invoice(self):
        s = self.s
        nid = s.stock_note("DELIVERY", 1, self.cust, [(self.pid, 4, 25)])
        s.post_stock_note(nid)
        self.assertEqual(s.stock(self.pid), 6)
        s.sync_journal()
        self.assertEqual((self.balance(1430), self.balance(1400)), (24, 36))
        moves = self.moves()
        inv = s.invoice_note(nid, credit=True)
        self.assertEqual((s.stock(self.pid), self.moves()), (6, moves))
        self.assertEqual(s.one("SELECT Sum(LineTotal) FROM SalesInvoiceDetails WHERE SalesInvoiceID = ?", inv), 100,
                         "the invoice keeps the price of the note (4 x 25 with VAT)")
        s.sync_journal()
        self.assertEqual(self.balance(1430), 0, "the invoice clears the goods delivered not invoiced")
        self.assertEqual((self.balance(5100), self.balance(1400)), (24, 36), "cost of sales once, stock once")

    def test_note_sale_cost_is_the_cost_at_delivery(self):
        s = self.s
        nid = s.stock_note("DELIVERY", 1, self.cust, [(self.pid, 2, 20)])
        s.post_stock_note(nid)
        s.apply_stock(self.pid, D(8), 5, D(15), "MANUAL", 0, "ADJ-2", recalc=True, warehouse=1)   # average moves
        inv = s.invoice_note(nid)
        self.assertEqual(s.one("SELECT UnitCost FROM SalesInvoiceDetails WHERE SalesInvoiceID = ?", inv), 6)
        s.sync_journal()
        self.assertEqual(self.balance(1430), 0)

    def test_note_from_a_branch(self):
        s = self.s
        s.c.execute("INSERT INTO Warehouses (WarehouseID, WarehouseCode, WarehouseName, WarehouseType) "
                    "VALUES (2, 'WH-2', 'فرع', 'BRANCH')")
        nid = s.stock_note("RECEIPT", 2, self.sup, [(self.pid, 5, 8)])
        s.post_stock_note(nid)
        inv = s.invoice_note(nid)
        self.assertEqual(s.one("SELECT WarehouseID FROM PurchaseInvoices WHERE PurchaseInvoiceID = ?", inv), 2)
        s.sync_journal()
        self.assertEqual((self.balance(140002), self.balance(2150)), (40, 0))

    def test_unpost_and_one_invoice(self):
        s = self.s
        nid = s.stock_note("RECEIPT", 1, self.sup, [(self.pid, 10, 10)])
        s.post_stock_note(nid)
        s.unpost_stock_note(nid)
        self.assertEqual((s.stock(self.pid), s.avg(self.pid)), (10, 6))
        s.post_stock_note(nid)
        s.invoice_note(nid)
        with self.assertRaises(SimError):
            s.unpost_stock_note(nid)
        with self.assertRaises(SimError):
            s.invoice_note(nid)

    def test_draft_has_no_entry(self):
        s = self.s
        s.stock_note("DELIVERY", 1, self.cust, [(self.pid, 1, 20)])
        s.sync_journal()
        self.assertEqual(s.one("SELECT COUNT(*) FROM JournalEntries WHERE SourceType = 'DELIVERY_NOTE'"), 0)

    def test_report_and_item_card(self):
        s = self.s
        nid = s.stock_note("DELIVERY", 1, self.cust, [(self.pid, 1, 20)])
        s.post_stock_note(nid)
        row = s.c.execute("SELECT PartyName, NoteValue FROM NotesNotInvoicedQuery WHERE StockNoteID = ?",
                          (nid,)).fetchone()
        self.assertEqual(row, ("TEST عميل الأذونات", 6))
        self.assertEqual(s.one("SELECT m.PartyName FROM qryStockMoveParty AS m INNER JOIN InventoryTransactions AS t "
                               "ON m.TransactionID = t.TransactionID WHERE t.ReferenceType = 'STOCK_NOTE'"),
                         "TEST عميل الأذونات")
        s.invoice_note(nid)
        self.assertEqual(s.one("SELECT COUNT(*) FROM NotesNotInvoicedQuery"), 0)


class StockNoteCodeTests(unittest.TestCase):

    def test_schema(self):
        codes = {a[0] for a in ACCOUNT_TREE}
        self.assertTrue({1430, 2150} <= codes)
        self.assertTrue({"DELIVERY_NOTE", "RECEIPT_NOTE"} <= {r[0] for r in table("JournalSourceTypes").seed_rows})
        self.assertTrue({"DELIVERY_NOTE", "RECEIPT_NOTE"} <= {r[1] for r in table("TransactionTypes").seed_rows})
        self.assertTrue({"DELIVERY_NOTE", "RECEIPT_NOTE"} <= {r[0] for r in table("Sequences").seed_rows})
        self.assertIn("STOCK_NOTES", {r[0] for r in table("Permissions").seed_rows})
        self.assertIn("frmStockNote", {r[0] for r in SCREEN_LIST})
        self.assertEqual(F.SCREEN_PERMISSIONS["frmStockNote"], "STOCK_NOTES")
        self.assertIn("qryJournalStockNote", Q.JOURNAL_SOURCE_QUERIES)
        self.assertIn("qryJournalStockNote", read("modJournal"))
        self.assertIn("TestStockNotes", read("modTestAll"))
        self.assertIn("1430", read("modVat"))

    def test_invoice_engines_skip_the_stock_of_a_note(self):
        sale = proc(read("modSales"), "PostSaleFromCart")
        self.assertIn("Optional ByVal StockNoteID As Long = 0", sale)
        self.assertIn("If Len(msg) = 0 And StockNoteID = 0 Then msg = CheckStock(productIDs, n)", sale)
        self.assertIn('MarkNoteInvoiced db, StockNoteID, "SalesInvoiceID", invoiceID', sale)
        self.assertLess(sale.index("If StockNoteID = 0 Then\n        For i = 0 To n - 1\n            ApplyStockMovement"),
                        sale.index("ws.CommitTrans"))
        purchase = proc(read("modPurchases"), "PostPurchaseFromCart")
        self.assertIn("Optional ByVal StockNoteID As Long = 0", purchase)
        self.assertIn('MarkNoteInvoiced db, StockNoteID, "PurchaseInvoiceID", invoiceID', purchase)
        self.assertIn("If StockNoteID = 0 Then\n            unitCost = PurchaseUnitCost", purchase)
        self.assertLess(purchase.index("MarkNoteInvoiced"), purchase.index("ws.CommitTrans"))

    def test_post_unpost_delete_invoice(self):
        post = proc(VBA, "PostStockNote")
        for part in ('CheckStockAvailable(ids, qtys, n, "إذن التسليم", wh)', "ClosedPeriodProblem(Date)",
                     "TT_DELIVERY_NOTE", "TT_RECEIPT_NOTE", '"STOCK_NOTE"'):
            self.assertIn(part, post)
        unpost = proc(VBA, "UnpostStockNote")
        for part in ('st = "INVOICED"', "WeightedAverage(cur!CurrentQuantity, cur!AverageCost, -rs!Quantity, rs!UnitCost)",
                     'AuditEdited "STOCK_NOTE_UNPOST"'):
            self.assertIn(part, unpost)
        delete = proc(VBA, "DeleteStockNote")
        self.assertLess(delete.index('AuditSnapshot("StockNotes"'), delete.index("DELETE FROM StockNotes"))
        self.assertIn('AuditDeleted "STOCK_NOTE_DELETE"', delete)
        invoice = proc(VBA, "InvoiceStockNote")
        self.assertIn("StockNoteID:=StockNoteID", invoice)
        self.assertIn('Nz(DbValue("SELECT COUNT(*) FROM " & cartTable), 0) > 0', invoice,
                      "a cart with an unsaved invoice is not wiped")
        self.assertIn('cart!LineDiscount = 0', invoice, "local tables have no defaults")
        mark = proc(VBA, "MarkNoteInvoiced")
        self.assertIn("db.RecordsAffected <> 1", mark)

    def test_screen(self):
        names = {c.name for c in MODELS["frmStockNote"].controls}
        for pname in ("StockNoteLoad", "StockNoteShow", "StockNoteNew", "StockNoteAction", "StockNoteTypeChanged"):
            for ctl in set(re.findall(r"frm!(\w+)", proc(VBA, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        code = "\n".join(MODELS["frmStockNote"].code)
        for call in ('StockNoteAction Me, "POST"', 'StockNoteAction Me, "UNPOST"', 'StockNoteAction Me, "DELETE"',
                     'StockNoteAction Me, "INVOICE"', "StockNoteNew Me", "StockNoteShow Me", "StockNoteTypeChanged Me"):
            self.assertIn(call, code)
        self.assertIn('OpenScreen "frmStockNote"', "\n".join(MODELS["frmInventory"].code))
        lines = {c.name for c in MODELS["frmStockNoteLines"].controls}
        self.assertTrue({"ProductID", "Quantity", "UnitPrice", "UnitCost"} <= lines)


class StaticModule(VbaModuleChecks, unittest.TestCase):
    module_name = "modStockNotes"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
