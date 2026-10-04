"""Phase 7 checks: purchases, purchase returns, payment vouchers, manual stock
movements and stocktaking.

  * the average-cost and unit-cost functions of modSales are EXECUTED with
    LibreOffice Basic and compared with tools/pricing.py;
  * the Phase 7 scenario (tools/purchases_reference.py) is replayed on the
    SQLite mirror with the same posting rules as modPurchases, and the saved
    queries (stock, supplier balance, profit, integrity) must agree with it;
  * the posting code is pinned to the ledger conventions of qrySupplierLedger;
  * static checks on modPurchases, modPurchaseScreens and modTestPurchases.

Run:  python3 -m unittest discover -s tests -v
"""

import os
import re
import unittest
from decimal import Decimal as D

from access_sqlite import AccessOnSqlite
from helpers import ROOT, VbaModuleChecks
import gen_test_purchases
import pricing as P
import purchases_reference as R
import vba_harness as H

TT = {"PURCHASE": 1, "SALE": 2, "PURCHASE_RETURN": 3, "SALES_RETURN": 4, "STOCK_IN": 5,
      "STOCK_OUT": 6, "ADJUSTMENT": 7, "OPENING": 8}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    """Body of one procedure."""
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


# --------------------------------------------------------------------------
# Python replay of the modPurchases posting rules on the SQLite mirror
# --------------------------------------------------------------------------
class Store:
    """Each method mirrors one VBA posting function (same names in comments)."""

    def __init__(self, db: AccessOnSqlite):
        self.db = db
        self.c = db.con

    def one(self, sql, *args):
        row = self.c.execute(sql, args).fetchone()
        return None if row is None else row[0]

    def insert(self, table, **values):
        cols = ", ".join(values)
        ph = ", ".join("?" for _ in values)
        return self.c.execute(f'INSERT INTO "{table}" ({cols}) VALUES ({ph})', list(values.values())).lastrowid

    def next_number(self, seq):                                   # modCommon.NextNumber
        n, prefix, pad = self.c.execute("SELECT NextValue, Prefix, PadLength FROM Sequences "
                                        "WHERE SequenceName = ?", (seq,)).fetchone()
        self.c.execute("UPDATE Sequences SET NextValue = NextValue + 1 WHERE SequenceName = ?", (seq,))
        return (prefix or "") + (str(n).zfill(pad) if pad else str(n))

    def apply_stock(self, pid, qty, tt, cost, ref_type, ref_id, ref_no, recalc=False):   # ApplyStockMovement
        cur, avg = self.c.execute("SELECT CurrentQuantity, AverageCost FROM Products WHERE ProductID = ?",
                                  (pid,)).fetchone()
        new = D(str(cur)) + qty
        if recalc:
            avg = P.weighted_average(D(str(cur)), D(str(avg)), qty, cost)
        self.c.execute("UPDATE Products SET CurrentQuantity = ?, AverageCost = ? WHERE ProductID = ?",
                       (float(new), float(avg), pid))
        self.insert("InventoryTransactions", TransactionDate=self.db.value(R_NOW), ProductID=pid,
                    TransactionTypeID=tt, Quantity=float(qty), UnitCost=float(cost), QuantityAfter=float(new),
                    ReferenceType=ref_type, ReferenceID=ref_id or None, ReferenceNumber=ref_no, EmployeeID=1)

    def adjust_supplier(self, sid, delta):                          # AdjustBalance "Suppliers"
        self.c.execute("UPDATE Suppliers SET CurrentBalance = CurrentBalance + ? WHERE SupplierID = ?",
                       (float(delta), sid))

    def purchase(self, sid, lines, credit, paid=None, charge_vat=True, supplier_no=None):   # PostPurchaseFromCart
        out = P.calc([P.LineIn(D(q), D(c), D(0), D("0.15") if charge_vat else D(0)) for _, q, c in lines],
                     D(0), False)
        tendered = (D(0) if credit else out.total) if paid is None else D(paid)
        paid_amt, remaining, _ = P.settle(out.total, tendered, credit)
        no = self.next_number("PURCHASE_INVOICE")
        inv = self.insert("PurchaseInvoices", InvoiceNumber=no, SupplierInvoiceNo=supplier_no,
                          InvoiceDate=self.db.value(R_NOW), SupplierID=sid, EmployeeID=1,
                          PaymentType="CREDIT" if credit else "CASH", PaymentMethodID=1,
                          SubTotal=float(out.subtotal), Discount=float(out.discount),
                          TaxableAmount=float(out.taxable), Tax=float(out.tax), TotalAmount=float(out.total),
                          PaidAmount=float(paid_amt), RemainingAmount=float(remaining))
        for i, ((pid, q, c), ln) in enumerate(zip(lines, out.lines)):
            self.insert("PurchaseInvoiceDetails", PurchaseInvoiceID=inv, LineNumber=i + 1, ProductID=pid,
                        Quantity=float(q), UnitCost=float(ln.unit_price), Discount=float(ln.discount),
                        NetAmount=float(ln.net), VATRate=0.15 if charge_vat else 0, Tax=float(ln.tax),
                        LineTotal=float(ln.total))
            cost = P.purchase_unit_cost(ln.net, ln.total, D(q), True)
            self.apply_stock(pid, D(q), TT["PURCHASE"], cost, "PURCHASE", inv, no, recalc=True)
            self.c.execute("UPDATE Products SET PurchasePrice = ? WHERE ProductID = ?", (float(ln.unit_price), pid))
        if remaining:
            self.adjust_supplier(sid, remaining)
        return inv

    def sale(self, pid, qty):                                       # PostSaleFromCart (cash)
        price, avg = self.c.execute("SELECT SellingPrice, AverageCost FROM Products WHERE ProductID = ?",
                                    (pid,)).fetchone()
        out = P.calc([P.LineIn(D(qty), D(str(price)))], D(0), True)
        ln = out.lines[0]
        no = self.next_number("SALES_INVOICE")
        inv = self.insert("SalesInvoices", InvoiceNumber=no, InvoiceDate=self.db.value(R_NOW), CustomerID=1,
                          EmployeeID=1, PaymentType="CASH", PaymentMethodID=1, SubTotal=float(out.subtotal),
                          Discount=float(out.discount), TaxableAmount=float(out.taxable), Tax=float(out.tax),
                          TotalAmount=float(out.total), PaidAmount=float(out.total), RemainingAmount=0,
                          AmountTendered=float(out.total), ChangeDue=0)
        self.insert("SalesInvoiceDetails", SalesInvoiceID=inv, LineNumber=1, ProductID=pid, Quantity=float(qty),
                    UnitPrice=float(ln.unit_price), Discount=float(ln.discount), NetAmount=float(ln.net),
                    VATRate=0.15, Tax=float(ln.tax), LineTotal=float(ln.total), UnitCost=avg)
        self.apply_stock(pid, -D(qty), TT["SALE"], D(str(avg)), "SALE", inv, no)
        return inv

    def purchase_return(self, inv, qty_by_detail, cash_refund=False):   # PostPurchaseReturn
        sid = self.one("SELECT SupplierID FROM PurchaseInvoices WHERE PurchaseInvoiceID = ?", inv)
        rows, total_sum, tax_sum = [], D(0), D(0)
        for detail, qty in qty_by_detail.items():
            pid, q, total, tax, net = self.c.execute(
                "SELECT ProductID, Quantity, LineTotal, Tax, NetAmount FROM PurchaseInvoiceDetails "
                "WHERE PurchaseDetailID = ?", (detail,)).fetchone()
            prev = self.c.execute("SELECT Sum(Quantity), Sum(LineTotal), Sum(Tax) FROM PurchaseReturnDetails "
                                  "WHERE PurchaseDetailID = ?", (detail,)).fetchone()
            pq, pt, px = (D(str(v or 0)) for v in prev)
            r_net, r_tax, r_total = P.return_amounts(D(str(q)), D(str(total)), D(str(tax)), pq, pt, px, D(qty))
            cost = P.purchase_unit_cost(D(str(net)), D(str(total)), D(str(q)), True)
            rows.append((detail, pid, D(qty), r_net, r_tax, r_total, cost))
            total_sum += r_total
            tax_sum += r_tax
        refunded = total_sum if cash_refund else D(0)
        no = self.next_number("PURCHASE_RETURN")
        ret = self.insert("PurchaseReturns", ReturnNumber=no, ReturnDate=self.db.value(R_NOW),
                          PurchaseInvoiceID=inv, SupplierID=sid, EmployeeID=1, Reason="TEST",
                          RefundType="CASH" if cash_refund else "CREDIT", SubTotal=float(total_sum - tax_sum),
                          Discount=0, TaxableAmount=float(total_sum - tax_sum), Tax=float(tax_sum),
                          TotalAmount=float(total_sum), RefundedAmount=float(refunded))
        for detail, pid, qty, r_net, r_tax, r_total, cost in rows:
            self.insert("PurchaseReturnDetails", PurchaseReturnID=ret, PurchaseDetailID=detail, ProductID=pid,
                        Quantity=float(qty), UnitCost=float(cost), Discount=0, NetAmount=float(r_net),
                        VATRate=0.15, Tax=float(r_tax), LineTotal=float(r_total))
            self.apply_stock(pid, -qty, TT["PURCHASE_RETURN"], cost, "PURCHASE_RETURN", ret, no, recalc=True)
        if refunded - total_sum:
            self.adjust_supplier(sid, refunded - total_sum)
        return ret, total_sum, tax_sum

    def payment(self, sid, amount):                                 # PostSupplierPayment
        no = self.next_number("SUPPLIER_PAYMENT")
        self.insert("SupplierPayments", PaymentNumber=no, SupplierID=sid, PaymentDate=self.db.value(R_NOW),
                    Amount=float(amount), PaymentMethodID=1, EmployeeID=1)
        self.adjust_supplier(sid, -D(amount))

    def manual(self, pid, tt, qty, cost=None):                      # PostManualStock
        avg = D(str(self.one("SELECT AverageCost FROM Products WHERE ProductID = ?", pid)))
        no = self.next_number("STOCK_ADJUST")
        if tt == TT["STOCK_OUT"]:
            self.apply_stock(pid, -D(qty), tt, avg, "MANUAL", None, no)
        else:
            self.apply_stock(pid, D(qty), tt, D(cost) if cost is not None else avg, "MANUAL", None, no, recalc=True)

    def stock_count(self, category, actual_by_product):             # CreateStockCount + PostStockCount
        no = self.next_number("STOCK_COUNT")
        cid = self.insert("StockCounts", CountNumber=no, CountDate=self.db.value(R_NOW), CategoryID=category,
                          Status="OPEN", EmployeeID=1)
        self.c.execute("INSERT INTO StockCountDetails (StockCountID, ProductID, SystemQuantity, ActualQuantity, "
                       "Difference, UnitCost, DifferenceValue) SELECT ?, ProductID, CurrentQuantity, NULL, 0, "
                       "AverageCost, 0 FROM Products WHERE IsActive = 1 AND CategoryID = ?", (cid, category))
        for pid, actual in actual_by_product.items():
            self.c.execute("UPDATE StockCountDetails SET ActualQuantity = ? WHERE StockCountID = ? AND ProductID = ?",
                           (float(actual), cid, pid))
        lines, value = 0, D(0)
        for detail, pid, actual in self.c.execute(
                "SELECT StockCountDetailID, ProductID, ActualQuantity FROM StockCountDetails "
                "WHERE StockCountID = ?", (cid,)).fetchall():
            sys_qty, cost = (D(str(v)) for v in self.c.execute(
                "SELECT CurrentQuantity, AverageCost FROM Products WHERE ProductID = ?", (pid,)).fetchone())
            act = sys_qty if actual is None else D(str(actual))
            diff = act - sys_qty
            val = P.r2(diff * cost)
            self.c.execute("UPDATE StockCountDetails SET SystemQuantity = ?, Difference = ?, UnitCost = ?, "
                           "DifferenceValue = ? WHERE StockCountDetailID = ?",
                           (float(sys_qty), float(diff), float(cost), float(val), detail))
            if diff:
                self.apply_stock(pid, diff, TT["ADJUSTMENT"], cost, "STOCK_COUNT", cid, no)
                lines += 1
                value += val
        self.c.execute("UPDATE StockCounts SET Status = 'POSTED' WHERE StockCountID = ?", (cid,))
        return cid, lines, value


R_NOW = None   # replaced in setUpClass by a Day() value inside the test period


class ScenarioOnMirrorTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        global R_NOW
        import queries as Q
        R_NOW = Q.Day(0, 1)
        cls.db = AccessOnSqlite()
        cls.db.load_fixture()
        cls.db.set_period(0)                     # today only: the scenario's documents
        cls.s = s = Store(cls.db)
        cls.e = R.expectations()
        cls.integrity_before = s.one("SELECT COUNT(*) FROM IntegrityCheckQuery")
        cls.profit_adj_before = s.one("SELECT InventoryAdjustments FROM ProfitQuery")
        cls.sid = s.insert("Suppliers", SupplierName="TEST مورد المرحلة 7", VATNumber="300000000000003",
                           OpeningBalance=0, CurrentBalance=0)
        cls.cat = s.insert("Categories", CategoryName="TEST جرد المرحلة 7")
        cls.pa = s.insert("Products", ProductCode="TEST-PUR-A", ProductName="TEST أ", CategoryID=cls.cat, UnitID=1,
                          SellingPrice=float(R.SALE_PRICE), CurrentQuantity=0, AverageCost=0, PurchasePrice=0)
        cls.pb = s.insert("Products", ProductCode="TEST-PUR-B", ProductName="TEST ب", CategoryID=cls.cat, UnitID=1,
                          SellingPrice=23, CurrentQuantity=0, AverageCost=0, PurchasePrice=0)
        log = cls.log = {}
        log["inv1"] = s.purchase(cls.sid, [(cls.pa, R.P1_QTY, R.P1_COST)], credit=True, paid=R.P1_PAID,
                                 supplier_no="S-1")
        log["after_p1"] = cls.snapshot()
        s.sale(cls.pa, R.SALE_QTY)
        log["after_sale"] = cls.snapshot()
        s.purchase(cls.sid, [(cls.pa, R.P2_QTY, R.P2_COST)], credit=False)
        log["after_p2"] = cls.snapshot()
        detail = s.one("SELECT PurchaseDetailID FROM PurchaseInvoiceDetails WHERE PurchaseInvoiceID = ?", log["inv1"])
        _, log["ret_total"], log["ret_tax"] = s.purchase_return(log["inv1"], {detail: R.RETURN_QTY})
        log["after_return"] = cls.snapshot()
        s.payment(cls.sid, R.PAYMENT)
        log["after_payment"] = cls.snapshot()
        s.manual(cls.pb, TT["OPENING"], R.B_OPENING_QTY, R.B_OPENING_COST)
        s.manual(cls.pb, TT["STOCK_OUT"], R.B_OUT_QTY)
        _, log["count_lines"], log["count_value"] = s.stock_count(cls.cat, {cls.pa: R.COUNT_A, cls.pb: R.COUNT_B})
        log["final"] = cls.snapshot()

    @classmethod
    def snapshot(cls):
        s = cls.s
        return {"stock_a": D(str(s.one("SELECT CurrentQuantity FROM Products WHERE ProductID = ?", cls.pa))),
                "avg_a": D(str(s.one("SELECT AverageCost FROM Products WHERE ProductID = ?", cls.pa))),
                "stock_b": D(str(s.one("SELECT CurrentQuantity FROM Products WHERE ProductID = ?", cls.pb))),
                "balance": D(str(s.one("SELECT CurrentBalance FROM Suppliers WHERE SupplierID = ?", cls.sid)))}

    def near(self, a, b):
        self.assertAlmostEqual(float(a), float(b), places=4)

    def test_spec_buy_100_sell_10_leaves_90(self):
        self.assertEqual(self.log["after_p1"]["stock_a"], 100)
        self.assertEqual(self.log["after_sale"]["stock_a"], 90)
        self.assertEqual(self.e["STOCK_AFTER_SALE"], 90)

    def test_weighted_average_cost_through_the_cycle(self):
        self.near(self.log["after_p1"]["avg_a"], self.e["AVG_1"])
        self.near(self.log["after_p2"]["avg_a"], self.e["AVG_2"])
        self.near(self.log["after_return"]["avg_a"], self.e["AVG_3"])
        self.assertEqual(self.e["AVG_2"], D("51"))            # (90 x 50 + 10 x 60) / 100

    def test_supplier_balance_steps(self):
        self.near(self.log["after_p1"]["balance"], self.e["P1_REMAINING"])
        self.near(self.log["after_p2"]["balance"], self.e["P1_REMAINING"])      # cash purchase
        self.near(self.log["after_return"]["balance"], self.e["BALANCE_AFTER_RETURN"])
        self.near(self.log["after_payment"]["balance"], self.e["BALANCE_FINAL"])
        self.near(self.log["ret_total"], self.e["RETURN_TOTAL"])

    def test_supplier_balance_query_agrees_with_cached_balance(self):
        bal, cached = self.s.c.execute("SELECT Balance, CachedBalance FROM SupplierBalanceQuery "
                                       "WHERE SupplierID = ?", (self.sid,)).fetchone()
        self.near(bal, self.e["BALANCE_FINAL"])
        self.near(bal, cached)

    def test_supplier_statement(self):
        self.db.params["SupplierID"] = self.sid
        rows = self.s.c.execute("SELECT EntryType, Debit, Credit FROM SupplierStatementQuery "
                                "WHERE SortKey = 1").fetchall()
        kinds = sorted(r[0] for r in rows)
        self.assertEqual(kinds, ["PAYMENT", "PURCHASE", "PURCHASE", "PURCHASE_RETURN"])
        self.near(sum(r[2] for r in rows) - sum(r[1] for r in rows), self.e["BALANCE_FINAL"])

    def test_stock_matches_ledger_and_count(self):
        final = self.log["final"]
        self.assertEqual(final["stock_a"], self.e["STOCK_FINAL_A"])
        self.assertEqual(final["stock_b"], self.e["B_STOCK"])
        self.assertEqual(self.log["count_lines"], self.e["COUNT_LINES"])
        self.near(self.log["count_value"], self.e["COUNT_VALUE"])
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM StockBalanceQuery WHERE QuantityMismatch <> 0"), 0)

    def test_integrity_check_finds_nothing_new(self):
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM IntegrityCheckQuery"), self.integrity_before)

    def test_profit_counts_stock_out_and_count_but_not_opening(self):
        # adjustments in the period: stock-out 2 x 10 and the count -2 x average (opening excluded)
        after = self.s.one("SELECT InventoryAdjustments FROM ProfitQuery")
        expected = -R.B_OUT_QTY * R.B_OPENING_COST + self.e["COUNT_DIFF_A"] * self.e["AVG_3"]
        self.assertAlmostEqual(float(after), float(expected), places=2)
        self.assertNotEqual(self.profit_adj_before, None)

    def test_vat_summary_includes_purchase_input_vat(self):
        purchase_vat = self.s.one("SELECT PurchaseVAT FROM VatSummaryQuery")
        p2_tax = P.calc([P.LineIn(R.P2_QTY, R.P2_COST)], D(0), False).tax
        self.near(purchase_vat, self.e["P1_TAX"] + p2_tax - self.e["RETURN_TAX"])

    def test_stock_count_query(self):
        rows = self.s.c.execute("SELECT ProductID, Difference, DifferenceValue, Status FROM StockCountQuery "
                                "WHERE CategoryID = ?", (self.cat,)).fetchall()
        self.assertEqual(len(rows), 2)
        by_pid = {r[0]: r for r in rows}
        self.near(by_pid[self.pa][2], self.e["COUNT_VALUE"])
        self.assertEqual(by_pid[self.pb][1], 0)
        self.assertEqual({r[3] for r in rows}, {"POSTED"})


# --------------------------------------------------------------------------
# The VBA posting code follows the same ledger conventions
# --------------------------------------------------------------------------
class PostingCodeTests(unittest.TestCase):
    text = read("modPurchases")

    def test_supplier_balance_deltas(self):
        self.assertIn('AdjustBalance db, "Suppliers", "SupplierID", SupplierID, remaining',
                      proc(self.text, "PostPurchaseFromCart"))
        self.assertIn('AdjustBalance db, "Suppliers", "SupplierID", supplierID, refunded - sumTotal',
                      proc(self.text, "PostPurchaseReturn"))
        self.assertIn('AdjustBalance db, "Suppliers", "SupplierID", SupplierID, -Amount',
                      proc(self.text, "PostSupplierPayment"))

    def test_stock_movement_types_and_average(self):
        body = proc(self.text, "PostPurchaseFromCart")
        self.assertRegex(body, r'ApplyStockMovement db, productIDs\(i\), CalcLine\(i, "QTY"\), TT_PURCHASE, '
                               r'unitCost, "PURCHASE",[\s\S]*?, True')
        self.assertIn("CalcRun(InvoiceDiscount, False)", body)          # purchase prices exclude VAT
        body = proc(self.text, "PostPurchaseReturn")
        self.assertRegex(body, r"ApplyStockMovement db, productIDs\(i\), -qtys\(i\), TT_PURCHASE_RETURN")
        self.assertIn('CheckStockAvailable(productIDs, qtys, n, "الإرجاع للمورد")', body)
        body = proc(self.text, "PostStockCount")
        self.assertIn("TT_ADJUSTMENT", body)
        self.assertNotRegex(body, r"TT_ADJUSTMENT[^\n]*True")            # a count never changes the average

    def test_every_posting_runs_in_one_transaction(self):
        for name in ("PostPurchaseFromCart", "PostPurchaseReturn", "PostSupplierPayment", "PostManualStock",
                     "CreateStockCount", "PostStockCount"):
            body = proc(self.text, name)
            self.assertEqual(body.count("ws.BeginTrans"), 1, name)
            self.assertEqual(body.count("ws.CommitTrans"), 1, name)
            self.assertIn("If inTrans Then ws.Rollback", body, name)
            self.assertNotIn("DLookup(", body, f"{name}: DLookup does not see the open transaction")

    def test_permissions_checked(self):
        for name, key in (("PostPurchaseFromCart", "PURCHASES"), ("PostPurchaseReturn", "PURCHASE_RETURN"),
                          ("PostSupplierPayment", "SUPPLIER_PAYMENTS"), ("PostManualStock", "INVENTORY_ADJUST"),
                          ("CreateStockCount", "STOCK_COUNT"), ("PostStockCount", "STOCK_COUNT"),
                          ("CancelStockCount", "STOCK_COUNT")):
            self.assertIn(f'HasPermission("{key}")', proc(self.text, name), name)

    def test_permission_keys_exist(self):
        from schema import table
        keys = {r[0] for r in table("Permissions").seed_rows}
        for module in ("modPurchases", "modPurchaseScreens", "modSales", "modPOS"):
            for key in re.findall(r'HasPermission\("(\w+)"\)', read(module)):
                self.assertIn(key, keys, f"{module}: {key}")

    def test_sequences_exist(self):
        from schema import table
        names = {r[0] for r in table("Sequences").seed_rows}
        for module in ("modPurchases", "modSales"):
            for seq in re.findall(r'NextNumber\("(\w+)"\)', read(module)):
                self.assertIn(seq, names, f"{module}: {seq}")


# --------------------------------------------------------------------------
# LibreOffice: execute the VBA average-cost functions
# --------------------------------------------------------------------------
DRIVER = r'''
Option VBASupport 1

Public Function RunAverage(cq As Double, ca As Double, q As Double, c As Double) As String
    RunAverage = Str(WeightedAverage(CCur(cq), CCur(ca), CCur(q), CCur(c)))
End Function

Public Function RunUnitCost(net As Double, total As Double, q As Double, registered As Boolean) As String
    RunUnitCost = Str(PurchaseUnitCost(CCur(net), CCur(total), CCur(q), registered))
End Function
'''


def plain(d) -> str:
    return format(D(d).normalize(), "f")


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class AverageCostRuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modCommon": H.read_module("modCommon"), "modSales": H.read_module("modSales"),
                    "Driver": DRIVER})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_average_cases(self):
        for label, cq, ca, q, c in P.AVERAGE_CASES:
            with self.subTest(label):
                got = self.h.call("Driver", "RunAverage", float(cq), float(ca), float(q), float(c)).strip()
                self.assertEqual(plain(got), plain(P.weighted_average(cq, ca, q, c)))

    def test_unit_cost(self):
        for net, total, q, reg in [(100, 115, 3, True), (100, 115, 3, False), (0, 0, 0, True), (575, 661.25, 7, True)]:
            got = self.h.call("Driver", "RunUnitCost", float(net), float(total), float(q), reg).strip()
            self.assertEqual(plain(got), plain(P.purchase_unit_cost(str(net), str(total), q, reg)), (net, q, reg))


# --------------------------------------------------------------------------
# Static checks
# --------------------------------------------------------------------------
class GeneratedTestModule(VbaModuleChecks, unittest.TestCase):
    module_name = "modTestPurchases"
    vba = gen_test_purchases.build_test_purchases_vba()

    def test_scenario_values_come_from_the_reference(self):
        e = R.expectations()
        for key in ("P1_TOTAL", "AVG_2", "AVG_3", "BALANCE_FINAL", "COUNT_VALUE"):
            self.assertIn(R.vba_number(e[key]), self.vba, key)
        for case in P.AVERAGE_CASES:
            self.assertIn("متوسط التكلفة: " + case[0], self.vba)


def _static(name):
    class T(VbaModuleChecks, unittest.TestCase):
        module_name = name
        vba = read(name)
    T.__name__ = T.__qualname__ = f"Static_{name}"
    return T


Static_modPurchases = _static("modPurchases")
Static_modPurchaseScreens = _static("modPurchaseScreens")
