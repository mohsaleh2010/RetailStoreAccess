"""The rep day and van sales (modVanSales, frmRepDay, docs/63-Van-Sales.md, plan 4B):
  * replayed on the SQLite mirror (tools/sim.py): a day is opened for a rep with a van warehouse and a custody box;
    loading is a transfer into the van, the office sells, takes returns and collects in the rep's name (the van and
    the custody box), the van count and the closing of the box give the goods and cash differences, and
    RepDaySummaryQuery adds the day up; the day is settled only when its transfers are complete, the van is
    counted and the box is closed;
  * the rules of the VBA code (working in the rep's name moves CurrentWarehouseID / CurrentCashBoxID /
    SalesRepFor / CanUseBox), the screen and the report."""
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
from sim import TT, SimError, Store

MODELS = {m.name: m for m in F.all_forms()}
ENTRIES = {r.key: r for r in F.REPORTS}


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


VBA = read("modVanSales")


class RepDayTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        s = self.s = Store(self.db.con, now=day(1, 8))
        s.user = 1
        self.pid = s.insert("Products", ProductCode="TEST-V1", ProductName="TEST صنف السيارة", CategoryID=1, UnitID=1,
                            SellingPrice=23, CurrentQuantity=0, AverageCost=0, PurchasePrice=8)
        self.rep = s.insert("SalesReps", RepCode="TEST-R", RepName="TEST مندوب", CommissionRate=0,
                            CommissionBase="SALES", IsActive=1)
        self.van = s.insert("Warehouses", WarehouseCode="TEST-VAN", WarehouseName="سيارة 1", WarehouseType="VAN",
                            SalesRepID=self.rep, ConfirmReceipt=0, IsActive=1)
        self.cust = s.insert("Customers", CustomerName="TEST عميل السيارة", AllowCredit=1, CreditLimit=0,
                             CurrentBalance=0, IsSystem=0)
        s.apply_stock(self.pid, D(10), 5, D(6), "MANUAL", 0, "ADJ-1", recalc=True, warehouse=1)

    def custody_box(self):
        return self.s.insert("CashBoxes", BoxName="TEST عهدة", BoxType="CUSTODY", OpeningBalance=0, IsActive=1,
                             SalesRepID=self.rep, ConfirmReceipt=0)

    def summary(self, did, *cols):
        return self.s.c.execute(f"SELECT {', '.join(cols)} FROM RepDaySummaryQuery WHERE RepDayID = ?", (did,)).fetchone()

    def van_count(self, diff):                       # a posted count of the van with one difference
        s = self.s
        cid = s.insert("StockCounts", CountNumber=s.next_number("STOCK_COUNT"), CountDate=s.now, Status="POSTED",
                       EmployeeID=s.user, PostedAt=s.now, WarehouseID=self.van)
        s.apply_stock(self.pid, D(diff), TT["ADJUSTMENT"], s.avg(self.pid), "STOCK_COUNT", cid, "CNT", warehouse=self.van)
        return cid

    def test_a_rep_needs_a_van_and_a_custody_box(self):
        with self.assertRaises(SimError):
            self.s.open_rep_day(self.rep)
        self.custody_box()
        did = self.s.open_rep_day(self.rep)
        with self.assertRaises(SimError):
            self.s.open_rep_day(self.rep)                     # one open day per rep
        self.assertEqual(self.s.one("SELECT WarehouseID FROM RepDays WHERE RepDayID = ?", did), self.van)

    def test_a_whole_day(self):
        s = self.s
        box = self.custody_box()
        did = s.open_rep_day(self.rep)
        s.now = day(1, 9)
        s.send_transfer(s.stock_transfer(1, self.van, [(self.pid, 6)]))           # load 6 x 6
        s.act_as_rep(did)
        s.now = day(1, 11)
        inv = s.sale([(self.pid, 2)], customer=self.cust)                           # cash, from the van
        self.assertEqual(s.one("SELECT WarehouseID FROM SalesInvoices WHERE SalesInvoiceID = ?", inv), self.van)
        self.assertEqual(s.one("SELECT CashBoxID FROM SalesInvoices WHERE SalesInvoiceID = ?", inv), box)
        credit = s.sale([(self.pid, 1)], customer=self.cust, credit=True)
        s.customer_payment(self.cust, 10)                                            # collected into the box
        s.sales_return(credit, {1: 1}, cash_refund=False)                           # back into the van
        self.assertEqual(s.one("SELECT Sum(Quantity) FROM InventoryTransactions WHERE WarehouseID = ?", self.van), 4)
        s.now = day(1, 17)
        self.van_count(-1)                                                           # one short
        s.send_transfer(s.stock_transfer(self.van, 1, [(self.pid, 3)]))             # unload the rest
        with self.assertRaises(SimError):
            s.settle_rep_day(did)                                                    # the box is not closed yet
        expected = s.cash_balance(box)
        self.assertEqual(expected, D("56.00"), "2 x 23 in cash and 10 collected")
        s.cash_closing(box, expected - 5, "MAIN", s.box_of_type("MAIN"), expected - 5)
        s.settle_rep_day(did)
        self.assertIsNone(s.acting, "settling stops working in the rep's name")
        self.assertEqual(self.summary(did, "LoadedValue", "SoldCost", "ReturnedCost", "UnloadedValue", "CountDiffValue"),
                         (36, 18, 6, 18, -6))
        self.assertEqual(self.summary(did, "InvoiceCount", "SalesTotal", "CashSales", "Collections", "ReturnsTotal"),
                         (2, 69, 46, 10, 23))
        self.assertEqual(self.summary(did, "CashExpected", "CashCounted", "CashDifference", "CashHandedOver", "StatusName"),
                         (56, 51, -5, 51, "تمت تسويته"))
        self.assertEqual(s.cash_balance(box), 0, "the cash went to the treasury")

    def balance(self, account):
        return round(self.s.one("SELECT Coalesce(Sum(Debit) - Sum(Credit), 0) FROM JournalLines WHERE AccountCode = ?",
                                account), 4)

    def test_shortages_are_an_advance_on_a_rep_with_a_user(self):
        """docs/66: the cash and the goods shortage of the day go to the rep's advances (1600), not to 5400 / 5200."""
        s = self.s
        s.c.execute("UPDATE SalesReps SET EmployeeID = 1 WHERE SalesRepID = ?", (self.rep,))
        box = self.custody_box()
        did = s.open_rep_day(self.rep)
        s.send_transfer(s.stock_transfer(1, self.van, [(self.pid, 4)]))              # 4 x 6
        s.act_as_rep(did)
        s.sale([(self.pid, 1)], customer=self.cust)                                  # 23 in cash
        self.van_count(-2)                                                           # two short: 12 at cost
        s.send_transfer(s.stock_transfer(self.van, 1, [(self.pid, 1)]))
        s.sync_journal()
        shortage_before, stock_diff_before = self.balance(5400), self.balance(5200)
        s.cash_closing(box, 20, "MAIN", s.box_of_type("MAIN"), 20)                  # 3 short
        s.settle_rep_day(did)
        s.sync_journal()
        self.assertEqual(s.one("SELECT Category FROM CashVouchers WHERE ClosingID IS NOT NULL AND CashBoxID = ? AND "
                               "VoucherType = 'OUT'", box), "ADVANCE")
        self.assertEqual(s.one("SELECT GoodsShortage FROM RepDays WHERE RepDayID = ?", did), 12)
        self.assertEqual(s.one("SELECT AdvanceBalance FROM qryAdvanceTotals WHERE EmployeeID = 1"), 15)
        self.assertEqual(self.balance(1600), 15)
        self.assertEqual(self.balance(5400), shortage_before)                       # not a box shortage
        self.assertEqual(self.balance(5200), stock_diff_before - 12)                # the count loss is taken back
        self.assertEqual(s.one("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0)

    def test_a_rep_without_a_user_keeps_the_shortage_general(self):
        s = self.s
        box = self.custody_box()
        did = s.open_rep_day(self.rep)
        self.van_count(-1)
        s.cash_closing(box, 0, "KEEP", None, 0)
        s.settle_rep_day(did)
        self.assertEqual(s.one("SELECT GoodsShortage FROM RepDays WHERE RepDayID = ?", did), 0)
        self.assertIsNone(s.one("SELECT AdvanceEmployeeID FROM RepDays WHERE RepDayID = ?", did))

    def test_settle_needs_the_transfers_done_and_the_van_counted(self):
        s = self.s
        box = self.custody_box()
        did = s.open_rep_day(self.rep)
        tid = s.stock_transfer(1, self.van, [(self.pid, 2)])                         # a draft load
        s.cash_closing(box, 0, "KEEP", None, 0)
        self.van_count(0)
        with self.assertRaises(SimError):
            s.settle_rep_day(did)
        s.send_transfer(tid)
        s.settle_rep_day(did)
        self.assertEqual(s.one("SELECT Status FROM RepDays WHERE RepDayID = ?", did), "SETTLED")

    def test_moves_outside_the_day_are_not_counted(self):
        s = self.s
        self.custody_box()
        s.send_transfer(s.stock_transfer(1, self.van, [(self.pid, 1)]))             # before the day
        s.now = day(1, 9)
        did = s.open_rep_day(self.rep)
        s.now = day(1, 10)
        s.send_transfer(s.stock_transfer(1, self.van, [(self.pid, 2)]))
        self.assertEqual(self.summary(did, "LoadedValue")[0], 12)


class RepDayCodeTests(unittest.TestCase):

    def test_schema(self):
        days = {f.name: f for f in table("RepDays").fields}
        self.assertEqual(days["Status"].rule, 'In ("OPEN","SETTLED")')
        self.assertEqual(days["WarehouseID"].fk, "Warehouses.WarehouseID")
        self.assertEqual(days["CashBoxID"].fk, "CashBoxes.CashBoxID")
        self.assertIn(("frmRepDay", "يوم المندوب والبيع من السيارة", "المبيعات", "REP_DAYS", True, True, True), SCREEN_LIST)
        self.assertIn("REP_DAYS", [r[0] for r in table("Permissions").seed_rows])

    def test_acting_as_the_rep(self):
        for module, name, value in (("modWarehouses", "CurrentWarehouseID", "ActingWarehouseID()"),
                                    ("modCash", "CurrentCashBoxID", "ActingCashBoxID()"),
                                    ("modSalesReps", "SalesRepFor", "ActingSalesRepID()")):
            code = proc(read(module), name)
            self.assertIn("If ActingRepDayID() > 0 Then", code, name)
            self.assertIn(f"{name} = {value}", code, name)
            self.assertLess(code.index("ActingRepDayID()"), code.index("DbValue("), name + ": before any look-up")
        self.assertIn("CLng(BoxID) = ActingCashBoxID()", proc(read("modCash"), "CanUseBox"))
        self.assertIn("StopActingAsRep", proc(VBA, "SettleRepDay"))
        self.assertIn("StopActingAsRep", "\n".join(MODELS["frmRepDay"].code))       # closing the screen stops it

    def test_settle_rules(self):
        settle = proc(VBA, "SettleRepDay")
        for part in ("Status In ('DRAFT', 'SENT')", "Status = 'POSTED' AND WarehouseID", "FROM CashClosings WHERE CashBoxID",
                     "PostedAt >= ", "ClosingDate >= "):
            self.assertIn(part, settle)
        self.assertIn("DayHasActivity(RepDayID)", proc(VBA, "DeleteRepDay"))
        self.assertIn('AuditDeleted "REP_DAY_DELETE"', proc(VBA, "DeleteRepDay"))
        unload = proc(VBA, "UnloadRepDay")
        self.assertLess(unload.index("ws.BeginTrans"), unload.index("CreateStockTransfer("))

    def test_screen_and_report(self):
        frm = MODELS["frmRepDay"]
        names = {c.name for c in frm.controls}
        for name in ("cboDay", "cboRep", "cboStore", "btnOpen", "btnLoad", "btnSale", "btnReturn", "btnPayment",
                     "btnCount", "btnUnload", "btnCash", "btnSettle", "lstSummary", "lblActing"):
            self.assertIn(name, names)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmRepDay"], "REP_DAYS")
        self.assertIn("frmRepDay", [t[3] for t in HUB_TILES])
        self.assertEqual((ENTRIES["REP_DAYS"].query, ENTRIES["REP_DAYS"].needs, ENTRIES["REP_DAYS"].date_column),
                         ("RepDaySummaryQuery", "D", "DayDate"))
        self.assertIn("REP_DAYS", {sp.key for sp in RC.LIST_SPECS})
        names = {q.name for q in Q.QUERIES}
        self.assertTrue({"qryRepDayStock", "qryRepDaySales", "qryRepDayReturns", "qryRepDayCollections"} <= names)

    def test_in_access_test(self):
        self.assertIn('"TestRepDays"', read("modTestAll"))
        test = proc(VBA, "TestRepDays")
        for part in ("ws.Rollback", "StartActingAsRep", "SettleRepDay", "PostCashClosing", "PostStockCount"):
            self.assertIn(part, test)


class StaticModule(VbaModuleChecks, unittest.TestCase):
    module_name = "modVanSales"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
