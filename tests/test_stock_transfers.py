"""Transfers between warehouses (modWarehouses, frmStockTransfer, docs/58-Stock-Transfers.md):
  * replayed on the SQLite mirror (tools/sim.py): a transfer moves the goods out of the sending warehouse at the
    average cost and into the receiving one; its entry moves the value from the stock account of the sender to
    the one of the receiver (1400 = main, 140000 + number for the others); through a warehouse that confirms its
    receipts the value waits on 1410 until the receipt entry (STOCK_RECEIVE);
  * every stock entry (sale, return, purchase, manual move, count) uses the account of its warehouse;
  * the rules of the VBA code and the screen."""
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
from schema import ACCOUNT_TREE, SCREEN_LIST, table
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


VBA = read("modWarehouses")


class TransferTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        s = self.s = Store(self.db.con, now=day(5, 9))
        s.user = 1
        c = s.c
        c.execute("INSERT INTO Warehouses (WarehouseID, WarehouseCode, WarehouseName, WarehouseType) "
                  "VALUES (2, 'WH-2', 'فرع', 'BRANCH')")
        self.pid = s.insert("Products", ProductCode="TEST-T1", ProductName="TEST صنف", CategoryID=1, UnitID=1,
                            SellingPrice=20, CurrentQuantity=0, AverageCost=0, PurchasePrice=0)
        s.apply_stock(self.pid, D(10), 5, D(8), "MANUAL", 0, "ADJ-1", recalc=True, warehouse=1)

    def wh(self, w):
        return D(str(self.s.one("SELECT Coalesce(Sum(Quantity), 0) FROM InventoryTransactions "
                                "WHERE ProductID = ? AND WarehouseID = ?", self.pid, w)))

    def account(self, code):
        return D(str(self.s.one("SELECT Coalesce(Sum(Debit - Credit), 0) FROM JournalLines WHERE AccountCode = ?",
                                code))).quantize(D("0.01"))

    def test_at_once(self):
        s = self.s
        t = s.stock_transfer(1, 2, [(self.pid, 4)])
        s.send_transfer(t)
        self.assertEqual((self.wh(1), self.wh(2)), (6, 4))
        self.assertEqual(s.stock(self.pid), 10, "the total does not change")
        self.assertEqual(s.one("SELECT Status FROM StockTransfers WHERE StockTransferID = ?", t), "RECEIVED")
        s.sync_journal()
        self.assertEqual(self.account(1400), 48)            # 80 in, 32 out
        self.assertEqual(self.account(140002), 32)
        self.assertEqual(self.account(1410), 0)
        self.assertEqual(s.one("SELECT AccountName FROM Accounts WHERE AccountCode = 140002"), "فرع")

    def test_through_transit(self):
        s = self.s
        s.c.execute("UPDATE Warehouses SET ConfirmReceipt = 1 WHERE WarehouseID = 2")
        t = s.stock_transfer(1, 2, [(self.pid, 5)])
        s.send_transfer(t)
        self.assertEqual((self.wh(1), self.wh(2)), (5, 0))
        self.assertEqual(s.stock(self.pid), 5, "goods in transit are in no warehouse")
        s.sync_journal()
        self.assertEqual((self.account(1400), self.account(1410), self.account(140002)), (40, 40, 0))
        s.now = day(3, 11)
        s.receive_transfer_stock(t)
        self.assertEqual((self.wh(2), s.stock(self.pid)), (5, 10))
        s.sync_journal()
        self.assertEqual((self.account(1410), self.account(140002)), (0, 40))
        when = s.one("SELECT EntryDate FROM JournalEntries WHERE SourceType = 'STOCK_RECEIVE'")
        self.assertEqual(when[:10], day(3)[:10])

    def test_documents_use_the_account_of_their_warehouse(self):
        s = self.s
        s.apply_stock(self.pid, D(2), 5, D(8), "MANUAL", 0, "ADJ-2", warehouse=2)
        s.sync_journal()
        self.assertEqual(self.account(140002), 16)
        for name in ("qryJournalSale", "qryJournalSalesReturn", "qryJournalPurchase", "qryJournalPurchaseReturn"):
            self.assertIn("IIf(", Q.query(name).sql)
            self.assertIn(".WarehouseID = 1, 1400, 140000 + ", Q.query(name).sql, name)
        self.assertIn("k.CountWarehouse = 1, 1400", Q.query("qryJournalStock").sql)


class TransferCodeTests(unittest.TestCase):

    def test_schema(self):
        self.assertIn(1410, {a[0] for a in ACCOUNT_TREE})
        self.assertIn("[ToWarehouseID]<>[FromWarehouseID]", table("StockTransfers").rule)
        self.assertTrue({"STOCK_TRANSFER", "STOCK_RECEIVE"} <= {r[0] for r in table("JournalSourceTypes").seed_rows})
        self.assertTrue({"TRANSFER_OUT", "TRANSFER_IN"} <= {r[1] for r in table("TransactionTypes").seed_rows})
        self.assertTrue(table("TransactionTypes").seed_missing)
        self.assertIn("STOCK_TRANSFER", {r[0] for r in table("Permissions").seed_rows})
        self.assertIn("STOCK_TRANSFER", {r[0] for r in table("Sequences").seed_rows})
        self.assertIn(("frmStockTransfer", "التحويل بين المخازن", "المخزون", "STOCK_TRANSFER", True, True, True),
                      SCREEN_LIST)
        self.assertIn("qryJournalStockTransfer", Q.JOURNAL_SOURCE_QUERIES)
        self.assertIn("qryJournalStockTransfer", read("modJournal"))
        self.assertIn("140000 + w.WarehouseID", proc(read("modJournal"), "EnsureAccounts"))

    def test_send_receive_unpost(self):
        send = proc(VBA, "SendStockTransfer")
        for part in ('CheckStockAvailable(ids, qtys, n, "التحويل بين المخازن", fromWh)', "ClosedPeriodProblem(Date)",
                     "transit = WarehouseConfirmsReceipt(toWh)", "TT_TRANSFER_OUT", "If Not transit Then"):
            self.assertIn(part, send)
        receive = proc(VBA, "ReceiveStockTransfer")
        self.assertIn("CanReceiveIn(toWh)", receive)
        self.assertIn('AuditEdited "STOCK_TRANSFER_RECEIVE"', receive)
        unpost = proc(VBA, "UnpostStockTransfer")
        self.assertIn('<> "SENT"', unpost)
        self.assertIn("CurrentQuantity = CurrentQuantity + ", unpost)
        self.assertIn("DELETE FROM InventoryTransactions WHERE ReferenceType = 'TRANSFER'", unpost)
        delete = proc(VBA, "DeleteStockTransfer")
        self.assertLess(delete.index('AuditSnapshot("StockTransfers"'), delete.index("DELETE FROM StockTransfers"))
        self.assertIn('AuditDeleted "STOCK_TRANSFER_DELETE"', delete)

    def test_screen(self):
        names = {c.name for c in MODELS["frmStockTransfer"].controls}
        for pname in ("TransferLoad", "TransferShow", "TransferNew", "TransferAction"):
            for ctl in set(re.findall(r"frm!(\w+)", proc(VBA, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        code = "\n".join(MODELS["frmStockTransfer"].code)
        for call in ('TransferAction Me, "SEND"', 'TransferAction Me, "RECEIVE"', 'TransferAction Me, "UNPOST"',
                     'TransferAction Me, "DELETE"', "TransferNew Me", "TransferShow Me"):
            self.assertIn(call, code)
        self.assertIn('OpenScreen "frmStockTransfer"', "\n".join(MODELS["frmInventory"].code))


if __name__ == "__main__":
    unittest.main()
