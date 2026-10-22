"""Box custodians, transfers in transit and the custody box (modCash, docs/56-Cash-Custodians.md):
  * replayed on the SQLite mirror: a transfer to a box that confirms its receipts leaves the sending box at
    once and enters the receiving box only when received; meanwhile its journal entry is on 1195 (cash in
    transit), and the receipt entry (CASH_RECEIVE) moves it to the box;
  * the rules of the VBA code: who may pay, transfer and close from a box, receipt and cancel;
  * schema, screens, permissions and their registration."""
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
from relations import UNENFORCED
from schema import ACCOUNT_TREE, RULE_UPGRADES, SCREEN_LIST, table
from sim import Store

FORMS = {m.name: m for m in F.all_forms()}
MAIN, CASHIER = 1, 2


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


CASH = read("modCash")


class TransitTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        s = self.s = Store(self.db.con, now=day(3, 9))
        s.user = 1
        s.cash_voucher("IN", CASHIER, D(1000), "OWNER")
        s.c.execute("UPDATE CashBoxes SET ConfirmReceipt = 1 WHERE CashBoxID = ?", (MAIN,))

    def account(self, code):
        return D(str(self.s.one("SELECT Coalesce(Sum(Debit - Credit), 0) FROM JournalLines WHERE AccountCode = ?",
                                code))).quantize(D("0.01"))

    def test_in_transit_until_received(self):
        s = self.s
        v = s.cash_voucher("TRANSFER", CASHIER, D(400), "TRANSFER", to_box=MAIN)
        self.assertEqual(s.one("SELECT InTransit FROM CashVouchers WHERE CashVoucherID = ?", v), 1)
        self.assertEqual(s.cash_balance(CASHIER), 600)
        self.assertEqual(s.cash_balance(MAIN), 0, "not in the receiving box before its receipt")
        s.sync_journal()
        self.assertEqual(self.account(1195), 400)
        self.assertEqual(self.account(110000 + MAIN), 0)
        s.now = day(2, 10)
        s.receive_transfer(v)
        self.assertEqual(s.cash_balance(MAIN), 400)
        self.assertEqual(s.one("SELECT Max(MoveDate) FROM qryCashMovements WHERE CashBoxID = ?", MAIN)[:10], day(2)[:10],
                         "it enters the box on the day of its receipt")
        s.sync_journal()
        self.assertEqual(self.account(1195), 0)
        self.assertEqual(self.account(110000 + MAIN), 400)
        self.assertEqual(s.one("SELECT COUNT(*) FROM JournalEntries WHERE SourceType = 'CASH_RECEIVE' AND SourceID = ?",
                               v), 1)
        self.assertEqual(s.one("SELECT EntryDate FROM JournalEntries WHERE SourceType = 'CASH_RECEIVE'")[:10],
                         day(2)[:10])

    def test_a_box_without_confirmation_receives_at_once(self):
        s = self.s
        s.cash_voucher("TRANSFER", CASHIER, D(100), "TRANSFER", to_box=MAIN)          # waits for its receipt
        s.c.execute("UPDATE CashBoxes SET ConfirmReceipt = 0")
        w = s.cash_voucher("TRANSFER", CASHIER, D(50), "TRANSFER", to_box=MAIN)
        self.assertEqual(s.one("SELECT InTransit FROM CashVouchers WHERE CashVoucherID = ?", w), 0)
        self.assertEqual(s.cash_balance(MAIN), 50)
        s.sync_journal()
        self.assertEqual(self.account(1195), 100, "only the first transfer waits")
        self.assertEqual(s.one("SELECT COUNT(*) FROM JournalEntries WHERE SourceType = 'CASH_RECEIVE'"), 0)


class CustodyCodeTests(unittest.TestCase):

    def test_schema(self):
        boxes = {f.name: f for f in table("CashBoxes").fields}
        self.assertIn('"CUSTODY"', boxes["BoxType"].rule)
        self.assertIn(("CashBoxes", "BoxType"), RULE_UPGRADES)
        self.assertEqual(boxes["SalesRepID"].fk, "SalesReps.SalesRepID")
        self.assertEqual(boxes["ConfirmReceipt"].default, "False")
        users = table("CashBoxUsers")
        self.assertEqual([i.fields for i in users.indexes if i.unique], [["CashBoxID", "EmployeeID"]])
        vouchers = {f.name for f in table("CashVouchers").fields}
        self.assertTrue({"InTransit", "ReceivedAt", "ReceivedByID"} <= vouchers)
        self.assertIn(("CashVouchers", "ReceivedByID"), UNENFORCED)
        self.assertIn(1195, {a[0] for a in ACCOUNT_TREE})
        self.assertIn("CASH_RECEIVE", {r[0] for r in table("JournalSourceTypes").seed_rows})
        self.assertIn("CASH_ALL_BOXES", {r[0] for r in table("Permissions").seed_rows})

    def test_screens(self):
        self.assertEqual(F.SCREEN_PERMISSIONS["frmCashBoxUsers"], "CASH_ALL_BOXES")
        self.assertEqual(F.SCREEN_PERMISSIONS["frmCashTransfers"], "CASH_CLOSING")
        self.assertTrue({"frmCashBoxUsers", "frmCashTransfers"} <= {r[0] for r in SCREEN_LIST})
        self.assertIn('OpenScreen "frmCashTransfers", 0', "\n".join(FORMS["frmTreasury"].code))
        names = {c.name for c in FORMS["frmCashTransfers"].controls}
        for pname in ("TransfersLoad", "TransfersReceive", "TransfersCancel", "TransfersPrint", "SelectedTransfer"):
            for ctl in set(re.findall(r"frm!(\w+)", proc(CASH, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        code = "\n".join(FORMS["frmCashTransfers"].code)
        for call in ("TransfersLoad Me", "TransfersReceive Me", "TransfersCancel Me", "TransfersPrint Me"):
            self.assertIn(call, code)

    def test_who_uses_a_box(self):
        can = proc(CASH, "CanUseBox")
        self.assertIn('HasPermission("CASH_ALL_BOXES")', can)
        self.assertIn("Not BoxHasCustodians(", can)
        custodian = proc(CASH, "IsBoxCustodian")
        for part in ("FROM Employees WHERE", "FROM CashBoxUsers WHERE", "s.EmployeeID = "):
            self.assertIn(part, custodian)
        post = proc(CASH, "PostCashVoucher")
        self.assertIn('If Len(msg) = 0 And VoucherType <> "IN" Then msg = BoxUseProblem(BoxID)', post)
        self.assertIn("If Not ClosingAllowed(BoxID) Then", proc(CASH, "PostCashClosing"))
        self.assertIn("ClosingBoxes()", proc(CASH, "ClosingLoad"))
        self.assertIn("Not CanUseBox(CashBoxID)", read("modBank"))
        self.assertIn("Not CanUseBox(CashBoxID)", proc(read("modPayroll"), "PayPayroll"))

    def test_transfer_in_transit_code(self):
        self.assertIn("rs!InTransit = BoxConfirmsReceipt(ToBoxID)", proc(CASH, "InsertVoucher"))
        receive = proc(CASH, "ReceiveCashTransfer")
        self.assertIn("CanUseBox(toBox)", receive)
        self.assertIn("ClosedPeriodProblem(", receive)
        self.assertLess(receive.index('AuditSnapshot("CashVouchers"'), receive.index("UPDATE CashVouchers"))
        self.assertIn('AuditEdited "CASH_TRANSFER_RECEIVE"', receive)
        cancel = proc(CASH, "CancelCashTransfer")
        self.assertIn("SELECT ClosingID FROM CashVouchers", cancel)
        self.assertLess(cancel.index('AuditSnapshot("CashVouchers"'), cancel.index("DELETE FROM CashVouchers"))
        self.assertIn('AuditDeleted "CASH_TRANSFER_DELETE"', cancel)
        pending = proc(CASH, "PendingTransferBox")
        self.assertIn("InTransit <> 0 AND ReceivedAt Is Null", pending)
        self.assertIn('Case "CASH_RECEIVE":', read("modJournal"))

    def test_movements_and_journal_queries(self):
        moves = Q.query("qryCashMovements").sql
        self.assertIn("IIf(v.InTransit, v.ReceivedAt, v.VoucherDate)", moves)
        self.assertIn("(v.InTransit = 0 OR v.ReceivedAt Is Not Null)", moves)
        journal = Q.query("qryJournalCashVoucher").sql
        self.assertIn("IIf(v.InTransit, 1195, 110000 + v.ToCashBoxID)", journal)
        self.assertIn("'CASH_RECEIVE' AS SourceType", journal)


if __name__ == "__main__":
    unittest.main()
