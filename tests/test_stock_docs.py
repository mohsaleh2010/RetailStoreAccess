"""Stock issue and receipt vouchers (modStockDocs, frmStockDocument, frmStockParties, docs/59-Stock-Issue-Receipt.md):
  * replayed on the SQLite mirror (tools/sim.py): an issue takes the goods out of its warehouse at the average cost
    and charges the account of the party (its own, else the account of its type); a receipt brings them in at its
    typed cost, updates the average cost, and credits the party; unposting gives back the quantity and the average;
  * the rules of the VBA code and the screens."""
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
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


VBA = read("modStockDocs")


class StockDocumentTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        s = self.s = Store(self.db.con, now=day(5, 9))
        s.user = 1
        c = s.c
        c.execute("INSERT INTO Warehouses (WarehouseID, WarehouseCode, WarehouseName, WarehouseType) "
                  "VALUES (2, 'WH-2', 'فرع', 'BRANCH')")
        self.hall = s.insert("StockParties", PartyCode="SP-2", PartyName="صالة الإنتاج", PartyType="PRODUCTION")
        self.dept = s.insert("StockParties", PartyCode="SP-3", PartyName="قسم الصيانة", PartyType="DEPARTMENT")
        self.pid = s.insert("Products", ProductCode="TEST-D1", ProductName="TEST صنف", CategoryID=1, UnitID=1,
                            SellingPrice=20, CurrentQuantity=0, AverageCost=0, PurchasePrice=0)
        s.apply_stock(self.pid, D(10), 5, D(6), "MANUAL", 0, "ADJ-1", recalc=True, warehouse=1)

    def wh(self, w):
        return D(str(self.s.one("SELECT Coalesce(Sum(Quantity), 0) FROM InventoryTransactions "
                                "WHERE ProductID = ? AND WarehouseID = ?", self.pid, w)))

    def account(self, code):
        return D(str(self.s.one("SELECT Coalesce(Sum(Debit - Credit), 0) FROM JournalLines WHERE AccountCode = ?",
                                code))).quantize(D("0.01"))

    def test_issue_to_each_party_type(self):
        s = self.s
        for party, account in ((self.hall, 1420), (self.dept, 5640), (1, 5250)):
            d = s.stock_document("ISSUE", 1, party, [(self.pid, 1)])
            s.post_stock_document(d)
            s.sync_journal()
            self.assertEqual(self.account(account), 6, account)
        self.assertEqual(self.wh(1), 7)
        self.assertEqual(self.account(1400), 42)            # 60 in by the manual move, 18 out
        self.assertEqual(s.avg(self.pid), 6, "an issue does not change the average cost")

    def test_party_account_wins(self):
        s = self.s
        s.c.execute("UPDATE StockParties SET AccountCode = 5900 WHERE StockPartyID = ?", (self.dept,))
        s.post_stock_document(s.stock_document("ISSUE", 1, self.dept, [(self.pid, 2)]))
        s.sync_journal()
        self.assertEqual((self.account(5900), self.account(5640)), (12, 0))

    def test_issue_from_a_branch_uses_its_account(self):
        s = self.s
        s.apply_stock(self.pid, D(3), 5, D(6), "MANUAL", 0, "ADJ-2", warehouse=2)
        s.post_stock_document(s.stock_document("ISSUE", 2, self.hall, [(self.pid, 2)]))
        s.sync_journal()
        self.assertEqual((self.wh(2), self.account(140002), self.account(1420)), (1, 6, 12))

    def test_receipt_updates_the_average_and_unpost_gives_it_back(self):
        s = self.s
        d = s.stock_document("RECEIPT", 1, self.hall, [(self.pid, 10, 10)])
        s.post_stock_document(d)
        self.assertEqual((self.wh(1), s.avg(self.pid)), (20, 8))
        s.sync_journal()
        self.assertEqual((self.account(1400), self.account(1420)), (160, -100))
        s.unpost_stock_document(d)
        self.assertEqual((self.wh(1), s.stock(self.pid), s.avg(self.pid)), (10, 10, 6))
        self.assertEqual(s.one("SELECT Status FROM StockDocuments WHERE StockDocumentID = ?", d), "DRAFT")
        s.sync_journal()
        self.assertEqual((self.account(1400), self.account(1420)), (60, 0))

    def test_draft_has_no_entry(self):
        s = self.s
        s.stock_document("ISSUE", 1, self.hall, [(self.pid, 2)])
        s.sync_journal()
        self.assertEqual(s.one("SELECT COUNT(*) FROM JournalEntries WHERE SourceType = 'STOCK_ISSUE'"), 0)


class StockDocumentCodeTests(unittest.TestCase):

    def test_schema(self):
        codes = {a[0] for a in ACCOUNT_TREE}
        self.assertTrue({1420, 5250, 5640} <= codes)
        self.assertTrue({"STOCK_ISSUE", "STOCK_RECEIPT"} <= {r[0] for r in table("JournalSourceTypes").seed_rows})
        self.assertTrue({"PARTY_ISSUE", "PARTY_RECEIPT"} <= {r[1] for r in table("TransactionTypes").seed_rows})
        self.assertIn("STOCK_DOCS", {r[0] for r in table("Permissions").seed_rows})
        self.assertTrue({"STOCK_ISSUE", "STOCK_RECEIPT"} <= {r[0] for r in table("Sequences").seed_rows})
        self.assertEqual(table("StockParties").seed_rows[0][3], "WASTE")
        for screen in ("frmStockParties", "frmStockDocument"):
            self.assertIn(screen, {r[0] for r in SCREEN_LIST})
            self.assertEqual(F.SCREEN_PERMISSIONS[screen], "STOCK_DOCS")
        self.assertIn("qryJournalStockDocument", Q.JOURNAL_SOURCE_QUERIES)
        self.assertIn("qryJournalStockDocument", read("modJournal"))
        self.assertIn("TestStockDocs", read("modTestAll"))

    def test_post_unpost_delete(self):
        post = proc(VBA, "PostStockDocument")
        for part in ('CheckStockAvailable(ids, qtys, n, "الصرف المخزني", wh)', "ClosedPeriodProblem(Date)",
                     "TT_PARTY_ISSUE", "TT_PARTY_RECEIPT", '"STOCK_DOC"', "partyName, True, wh"):
            self.assertIn(part, post)
        unpost = proc(VBA, "UnpostStockDocument")
        for part in ('<> "POSTED"', "WeightedAverage(cur!CurrentQuantity, cur!AverageCost, -rs!Quantity, rs!UnitCost)",
                     "DELETE FROM InventoryTransactions WHERE ReferenceType = 'STOCK_DOC'", "CheckStockAvailable",
                     'AuditEdited "STOCK_DOC_UNPOST"'):
            self.assertIn(part, unpost)
        delete = proc(VBA, "DeleteStockDocument")
        self.assertLess(delete.index('AuditSnapshot("StockDocuments"'), delete.index("DELETE FROM StockDocuments"))
        self.assertIn('AuditDeleted "STOCK_DOC_DELETE"', delete)

    def test_screens(self):
        names = {c.name for c in MODELS["frmStockDocument"].controls}
        for pname in ("StockDocLoad", "StockDocShow", "StockDocNew", "StockDocAction"):
            for ctl in set(re.findall(r"frm!(\w+)", proc(VBA, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        code = "\n".join(MODELS["frmStockDocument"].code)
        for call in ('StockDocAction Me, "POST"', 'StockDocAction Me, "UNPOST"', 'StockDocAction Me, "DELETE"',
                     "StockDocNew Me", "StockDocShow Me", 'OpenScreen "frmStockParties"'):
            self.assertIn(call, code)
        self.assertIn('OpenScreen "frmStockDocument"', "\n".join(MODELS["frmInventory"].code))
        lines = {c.name for c in MODELS["frmStockDocumentLines"].controls}
        self.assertTrue({"ProductID", "Quantity", "UnitCost"} <= lines)


class StaticModule(VbaModuleChecks, unittest.TestCase):
    module_name = "modStockDocs"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
