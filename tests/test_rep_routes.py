"""Routes, visits and rep limits (modRepRoutes, frmRepVisits, docs/64-Rep-Routes-Limits.md, plan 4C):
  * replayed on the SQLite mirror (tools/sim.py): the plan of a day comes from the route stops of its weekday, once;
    a visit outside the route; the result of a visit read from the documents of the customer that day;
    RouteExecutionQuery, RepVisitListQuery and RepVanStockQuery; the rep's price list and his credit limit in a sale;
  * the rules of the VBA code (the limits in modSales / modPOS / modSalesOrders), the screens and the reports."""
import datetime as dt
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


VBA = read("modRepRoutes")


class RouteTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        s = self.s = Store(self.db.con, now=day(1, 8))
        s.user = 1
        self.rep = s.insert("SalesReps", RepCode="TEST-R", RepName="TEST مندوب", CommissionRate=0, CommissionBase="SALES",
                            IsActive=1, CreditLimit=0, MaxDiscountPercent=0)
        self.c1, self.c2, self.c3 = (s.insert("Customers", CustomerName=f"TEST عميل {i}", AllowCredit=1, CreditLimit=0,
                                              CurrentBalance=0, IsSystem=0, SalesRepID=self.rep) for i in (1, 2, 3))
        self.wd = s.weekday(s.now)
        s.insert("RouteStops", SalesRepID=self.rep, CustomerID=self.c1, VisitWeekday=self.wd, StopOrder=2, IsActive=1)
        s.insert("RouteStops", SalesRepID=self.rep, CustomerID=self.c2, VisitWeekday=self.wd, StopOrder=1, IsActive=1)
        s.insert("RouteStops", SalesRepID=self.rep, CustomerID=self.c3, VisitWeekday=self.wd % 7 + 1, StopOrder=1,
                 IsActive=1)

    def visit(self, cust):
        return self.s.one("SELECT RepVisitID FROM RepVisits WHERE CustomerID = ?", cust)

    def test_weekday_is_vba_weekday(self):
        self.assertEqual(Store.weekday("2026-10-18"), 1)          # a Sunday
        self.assertEqual(Store.weekday("2026-10-17"), 7)          # a Saturday

    def test_the_plan_of_a_day(self):
        s = self.s
        self.assertEqual(s.create_visit_plan(self.rep, s.now), 2, "the customers of today's weekday only")
        self.assertEqual(s.create_visit_plan(self.rep, s.now), 0, "once")
        self.assertIsNone(self.visit(self.c3))
        s.add_visit(self.rep, self.c3, s.now)
        with self.assertRaises(SimError):
            s.add_visit(self.rep, self.c3, s.now)
        rows = s.c.execute("SELECT CustomerID, StopOrder, PlanName, ResultName FROM RepVisitListQuery "
                           "ORDER BY StopOrder").fetchall()
        self.assertEqual(rows, [(self.c2, 1, "من خط السير", "لم تتم"), (self.c1, 2, "من خط السير", "لم تتم"),
                                (self.c3, 3, "خارج الخط", "لم تتم")])
        with self.assertRaises(SimError):                               # a weekday without stops
            s.create_visit_plan(self.rep, (dt.date.fromisoformat(s.now[:10]) + dt.timedelta(days=2)).isoformat())

    def test_results_from_the_documents_and_the_execution_report(self):
        s = self.s
        s.create_visit_plan(self.rep, s.now)
        s.add_visit(self.rep, self.c3, s.now)
        s.insert("SalesOrders", OrderNumber="SO-T1", OrderDate=day(1, 10), CustomerID=self.c1, WarehouseID=1,
                 Status="DRAFT", EmployeeID=1)
        s.insert("CustomerPayments", PaymentNumber="RC-T1", CustomerID=self.c3, PaymentDate=day(1, 11), Amount=5,
                 PaymentMethodID=1, EmployeeID=1)
        s.insert("CustomerPayments", PaymentNumber="RC-T2", CustomerID=self.c2, PaymentDate=day(2, 11), Amount=5,
                 PaymentMethodID=1, EmployeeID=1)                       # another day: no result
        self.assertEqual(s.suggested_result(self.visit(self.c1)), "ORDER")
        self.assertEqual(s.suggested_result(self.visit(self.c3)), "PAYMENT")
        self.assertEqual(s.suggested_result(self.visit(self.c2)), "")
        with self.assertRaises(SimError):
            s.record_visit(self.visit(self.c2))
        s.record_visit(self.visit(self.c1))
        s.record_visit(self.visit(self.c3), "PAYMENT", "TEST")
        row = s.c.execute("SELECT VisitCount, PlannedCount, DoneCount, MissedCount, OrderCount, PaymentCount, DoneRate "
                          "FROM RouteExecutionQuery WHERE SalesRepID = ?", (self.rep,)).fetchone()
        self.assertEqual(row[:6], (3, 2, 2, 1, 1, 1))
        self.assertAlmostEqual(row[6], 2 / 3)

    def test_van_stock_per_rep(self):
        s = self.s
        pid = s.insert("Products", ProductCode="TEST-V", ProductName="TEST صنف", CategoryID=1, UnitID=1, SellingPrice=10,
                       CurrentQuantity=0, AverageCost=0, PurchasePrice=4)
        van = s.insert("Warehouses", WarehouseCode="TEST-VAN", WarehouseName="سيارة 1", WarehouseType="VAN",
                       SalesRepID=self.rep, ConfirmReceipt=0, IsActive=1)
        s.apply_stock(pid, D(5), 5, D(4), "MANUAL", 0, "ADJ-1", recalc=True, warehouse=1)
        s.send_transfer(s.stock_transfer(1, van, [(pid, 3)]))
        rows = s.c.execute("SELECT RepName, WarehouseName, StockQty, StockCostValue FROM RepVanStockQuery").fetchall()
        self.assertEqual(rows, [("TEST مندوب", "سيارة 1", 3, 12)], "the van only, not the main warehouse")


class RepLimitTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        s = self.s = Store(self.db.con, now=day(1, 8))
        s.user = 1
        self.rep = s.insert("SalesReps", RepCode="TEST-R", RepName="TEST مندوب", CommissionRate=0, CommissionBase="SALES",
                            IsActive=1, CreditLimit=100, MaxDiscountPercent=0.05, EmployeeID=1)
        self.cust = s.insert("Customers", CustomerName="TEST عميل", AllowCredit=1, CreditLimit=0, CurrentBalance=0,
                             IsSystem=0, SalesRepID=self.rep)
        self.pid = s.insert("Products", ProductCode="TEST-P", ProductName="TEST صنف", CategoryID=1, UnitID=1,
                            SellingPrice=20, CurrentQuantity=0, AverageCost=0, PurchasePrice=8)
        s.apply_stock(self.pid, D(20), 5, D(8), "MANUAL", 0, "ADJ-1", recalc=True, warehouse=1)
        s.insert("RepPrices", SalesRepID=self.rep, ProductID=self.pid, UnitPrice=18)

    def test_the_rep_price_list(self):
        s = self.s
        inv = s.sale([(self.pid, 2)], customer=self.cust)
        self.assertEqual(s.one("SELECT LineTotal FROM SalesInvoiceDetails WHERE SalesInvoiceID = ?", inv), 36)
        self.assertEqual(s.one("SELECT SalesRepID FROM SalesInvoices WHERE SalesInvoiceID = ?", inv), self.rep)
        s.c.execute("UPDATE SalesReps SET EmployeeID = Null")
        inv = s.sale([(self.pid, 1)], customer=self.cust)
        self.assertEqual(s.one("SELECT LineTotal FROM SalesInvoiceDetails WHERE SalesInvoiceID = ?", inv), 20,
                         "the user is not the rep: the selling price")

    def test_the_rep_credit_limit(self):
        s = self.s
        s.sale([(self.pid, 5)], customer=self.cust, credit=True)        # 90 on credit
        with self.assertRaises(SimError):
            s.sale([(self.pid, 1)], customer=self.cust, credit=True)    # 90 + 18 > 100
        s.sale([(self.pid, 1)], customer=self.cust)                     # cash is fine
        s.c.execute("UPDATE SalesReps SET CreditLimit = 0")
        s.sale([(self.pid, 1)], customer=self.cust, credit=True)        # 0 = no limit


class RepRoutesCodeTests(unittest.TestCase):

    def test_schema(self):
        visits = {f.name: f for f in table("RepVisits").fields}
        self.assertEqual(visits["Result"].rule, 'In ("PENDING","ORDER","SALE","PAYMENT","NO_SALE","CLOSED")')
        stops = {f.name: f for f in table("RouteStops").fields}
        self.assertEqual(stops["VisitWeekday"].rule, "Between 1 And 7")
        self.assertIn(["VisitDate", "SalesRepID", "CustomerID"], [ix.fields for ix in table("RepVisits").indexes])
        self.assertIn(["SalesRepID", "ProductID"], [ix.fields for ix in table("RepPrices").indexes])
        reps = {f.name: f for f in table("SalesReps").fields}
        self.assertEqual((reps["CreditLimit"].default, reps["MaxDiscountPercent"].default), ("0", "0"))
        for name in ("frmRepVisits", "frmRouteStops", "frmRepPrices"):
            self.assertIn(name, [r[0] for r in SCREEN_LIST])
            self.assertEqual(F.SCREEN_PERMISSIONS[name], "REP_ROUTES")
        self.assertIn("REP_ROUTES", [r[0] for r in table("Permissions").seed_rows])

    def test_the_limits_are_used(self):
        sales = read("modSales")
        self.assertIn("RepDiscountLimit(maxPct)", proc(sales, "CheckDiscountAllowed"))
        self.assertIn("RepCreditProblem(SalesRepFor(CustomerID), remaining)", proc(sales, "PostSaleFromCart"))
        self.assertIn("IsSessionRepPrice(rs!ProductID, rs!UnitPrice)", proc(sales, "PostSaleFromCart"))
        pos = read("modPOS")
        self.assertIn("SessionPrice(ProductID, p!SellingPrice)", proc(pos, "AddLine"))
        self.assertIn("IsSessionRepPrice(", proc(pos, "LineChanged"))
        self.assertIn("OrderDefaultPrice(id,", proc(read("modSalesOrders"), "SalesOrderLineBeforeUpdate"))
        session = proc(VBA, "SessionRepID")
        self.assertLess(session.index("ActingRepDayID()"), session.index("DbValue("), "the rep day comes first")

    def test_visits(self):
        plan = proc(VBA, "CreateVisitPlan")
        self.assertIn("VisitWeekdayOf(d)", plan)
        self.assertIn("NOT IN (SELECT v.CustomerID FROM", plan)
        self.assertIn("Weekday(VisitDate, vbSunday)", proc(VBA, "VisitWeekdayOf"))
        self.assertNotIn('"dddd"', VBA, "weekday names do not follow the Windows locale")
        delete = proc(VBA, "DeleteRepVisit")
        self.assertIn("<> \"PENDING\"", delete)
        self.assertIn('AuditDeleted "REP_VISIT_DELETE"', delete)
        self.assertIn('AuditEdited "REP_VISIT_RECORD"', proc(VBA, "RecordRepVisit"))

    def test_screens_and_reports(self):
        frm = MODELS["frmRepVisits"]
        names = {c.name for c in frm.controls}
        for name in ("cboRep", "txtDate", "btnPlan", "btnRoute", "btnPrices", "cboCustomer", "btnAddVisit",
                     "lstVisits", "cboResult", "txtVisitNotes", "btnRecord", "btnDeleteVisit", "btnExecution"):
            self.assertIn(name, names)
        self.assertIn("frmRepVisits", [t[3] for t in HUB_TILES])
        self.assertIn("frmRouteStops", MODELS)
        self.assertIn("frmRepPrices", MODELS)
        for key, query, needs, col in (("ROUTE_EXECUTION", "RouteExecutionQuery", "D", "VisitDate"),
                                       ("REP_VISITS", "RepVisitListQuery", "Dc", "VisitDate"),
                                       ("REP_VAN_STOCK", "RepVanStockQuery", "w", "")):
            e = ENTRIES[key]
            self.assertEqual((e.query, e.needs, e.date_column), (query, needs, col))
            self.assertIn(key, {sp.key for sp in RC.LIST_SPECS})
        self.assertTrue({"RouteExecutionQuery", "RepVisitListQuery", "RepVanStockQuery"} <= {q.name for q in Q.QUERIES})

    def test_in_access_test(self):
        self.assertIn('"TestRepRoutes"', read("modTestAll"))
        test = proc(VBA, "TestRepRoutes")
        for part in ("ws.Rollback", "CreateVisitPlan", "RecordRepVisit", "RouteExecutionQuery", "RepCreditProblem",
                     "PostSaleFromCart"):
            self.assertIn(part, test)


class StaticModule(VbaModuleChecks, unittest.TestCase):
    module_name = "modRepRoutes"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
