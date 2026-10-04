"""Python replay of the VBA posting functions on the SQLite mirror of the back end.

Each method mirrors one VBA function (named in its comment) with the same
calculation engine (tools/pricing.py) and the same ledger conventions, so the
saved queries can be checked against a known history:
  * tests/test_purchases.py replays the Phase 7 scenario;
  * tools/demo_data.py replays the demo data and the expected results it
    produces are written into modDemoData, which verifies them in Access.

The rules that refuse a document in VBA (not enough stock, credit limit, ...)
raise SimError here, so a plan that Access would refuse cannot pass in Python.
"""

from decimal import Decimal as D
from typing import Dict, List, Optional, Tuple

import pricing as P

TT = {"PURCHASE": 1, "SALE": 2, "PURCHASE_RETURN": 3, "SALES_RETURN": 4, "STOCK_IN": 5,
      "STOCK_OUT": 6, "ADJUSTMENT": 7, "OPENING": 8}
VAT = D("0.15")


class SimError(ValueError):
    pass


def dec(v) -> D:
    return D(str(v if v is not None else 0))


class Store:

    def __init__(self, con, now: str = "2026-10-03 12:00:00", vat_registered: bool = True):
        self.c = con
        self.now = now                    # date/time given to the next documents
        self.vat_registered = vat_registered
        self.user = 1

    # ------------------------------------------------------------- helpers
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

    def stock(self, pid) -> D:
        return dec(self.one("SELECT CurrentQuantity FROM Products WHERE ProductID = ?", pid))

    def avg(self, pid) -> D:
        return dec(self.one("SELECT AverageCost FROM Products WHERE ProductID = ?", pid))

    def apply_stock(self, pid, qty, tt, cost, ref_type, ref_id, ref_no, recalc=False):   # ApplyStockMovement
        cur, avg = self.stock(pid), self.avg(pid)
        new = cur + qty
        if recalc:
            avg = P.weighted_average(cur, avg, qty, cost)
        self.c.execute("UPDATE Products SET CurrentQuantity = ?, AverageCost = ? WHERE ProductID = ?",
                       (float(new), float(avg), pid))
        self.insert("InventoryTransactions", TransactionDate=self.now, ProductID=pid, TransactionTypeID=tt,
                    Quantity=float(qty), UnitCost=float(cost), QuantityAfter=float(new), ReferenceType=ref_type,
                    ReferenceID=ref_id or None, ReferenceNumber=ref_no, EmployeeID=self.user)

    def adjust(self, table, key, kid, delta):                      # AdjustBalance
        self.c.execute(f"UPDATE {table} SET CurrentBalance = CurrentBalance + ? WHERE {key} = ?",
                       (float(delta), kid))

    def check_stock(self, needs: Dict[int, D]):                    # CheckStockAvailable (no override)
        for pid, need in needs.items():
            if need > self.stock(pid):
                raise SimError(f"not enough stock of product {pid}: need {need}, have {self.stock(pid)}")

    # ------------------------------------------------------------- purchases
    def purchase(self, sid, lines, credit, paid=None, charge_vat=True, supplier_no=None,
                 discount=0):                                       # PostPurchaseFromCart
        out = P.calc([P.LineIn(dec(q), dec(c), D(0), VAT if charge_vat else D(0)) for _, q, c in lines],
                     dec(discount), False)
        tendered = (D(0) if credit else out.total) if paid is None else dec(paid)
        paid_amt, remaining, _ = P.settle(out.total, tendered, credit)
        no = self.next_number("PURCHASE_INVOICE")
        inv = self.insert("PurchaseInvoices", InvoiceNumber=no, SupplierInvoiceNo=supplier_no, InvoiceDate=self.now,
                          SupplierID=sid, EmployeeID=self.user, PaymentType="CREDIT" if credit else "CASH",
                          PaymentMethodID=1, SubTotal=float(out.subtotal), Discount=float(out.discount),
                          TaxableAmount=float(out.taxable), Tax=float(out.tax), TotalAmount=float(out.total),
                          PaidAmount=float(paid_amt), RemainingAmount=float(remaining))
        for i, ((pid, q, c), ln) in enumerate(zip(lines, out.lines)):
            self.insert("PurchaseInvoiceDetails", PurchaseInvoiceID=inv, LineNumber=i + 1, ProductID=pid,
                        Quantity=float(q), UnitCost=float(ln.unit_price), Discount=float(ln.discount),
                        NetAmount=float(ln.net), VATRate=float(VAT if charge_vat else 0), Tax=float(ln.tax),
                        LineTotal=float(ln.total))
            cost = P.purchase_unit_cost(ln.net, ln.total, dec(q), self.vat_registered)
            self.apply_stock(pid, dec(q), TT["PURCHASE"], cost, "PURCHASE", inv, no, recalc=True)
            self.c.execute("UPDATE Products SET PurchasePrice = ? WHERE ProductID = ?", (float(ln.unit_price), pid))
        if remaining:
            self.adjust("Suppliers", "SupplierID", sid, remaining)
        return inv

    def purchase_return(self, inv, qty_by_detail, cash_refund=False):    # PostPurchaseReturn
        sid = self.one("SELECT SupplierID FROM PurchaseInvoices WHERE PurchaseInvoiceID = ?", inv)
        rows, total_sum, tax_sum = [], D(0), D(0)
        needs: Dict[int, D] = {}
        for detail, qty in qty_by_detail.items():
            pid, q, total, tax, net = self.c.execute(
                "SELECT ProductID, Quantity, LineTotal, Tax, NetAmount FROM PurchaseInvoiceDetails "
                "WHERE PurchaseDetailID = ?", (detail,)).fetchone()
            prev = self.c.execute("SELECT Sum(Quantity), Sum(LineTotal), Sum(Tax) FROM PurchaseReturnDetails "
                                  "WHERE PurchaseDetailID = ?", (detail,)).fetchone()
            pq, pt, px = (dec(v) for v in prev)
            r_net, r_tax, r_total = P.return_amounts(dec(q), dec(total), dec(tax), pq, pt, px, dec(qty))
            cost = P.purchase_unit_cost(dec(net), dec(total), dec(q), self.vat_registered)
            rows.append((detail, pid, dec(qty), r_net, r_tax, r_total, cost))
            needs[pid] = needs.get(pid, D(0)) + dec(qty)
            total_sum += r_total
            tax_sum += r_tax
        self.check_stock(needs)
        refunded = total_sum if cash_refund else D(0)
        no = self.next_number("PURCHASE_RETURN")
        ret = self.insert("PurchaseReturns", ReturnNumber=no, ReturnDate=self.now, PurchaseInvoiceID=inv,
                          SupplierID=sid, EmployeeID=self.user, Reason="DEMO", RefundType="CASH" if cash_refund
                          else "CREDIT", SubTotal=float(total_sum - tax_sum), Discount=0,
                          TaxableAmount=float(total_sum - tax_sum), Tax=float(tax_sum),
                          TotalAmount=float(total_sum), RefundedAmount=float(refunded))
        for detail, pid, qty, r_net, r_tax, r_total, cost in rows:
            self.insert("PurchaseReturnDetails", PurchaseReturnID=ret, PurchaseDetailID=detail, ProductID=pid,
                        Quantity=float(qty), UnitCost=float(cost), Discount=0, NetAmount=float(r_net),
                        VATRate=float(VAT), Tax=float(r_tax), LineTotal=float(r_total))
            self.apply_stock(pid, -qty, TT["PURCHASE_RETURN"], cost, "PURCHASE_RETURN", ret, no, recalc=True)
        if refunded - total_sum:
            self.adjust("Suppliers", "SupplierID", sid, refunded - total_sum)
        return ret, total_sum, tax_sum

    def payment(self, sid, amount):                                     # PostSupplierPayment
        no = self.next_number("SUPPLIER_PAYMENT")
        self.insert("SupplierPayments", PaymentNumber=no, SupplierID=sid, PaymentDate=self.now,
                    Amount=float(amount), PaymentMethodID=1, EmployeeID=self.user)
        self.adjust("Suppliers", "SupplierID", sid, -dec(amount))

    # ------------------------------------------------------------- sales
    def sale(self, lines: List[Tuple[int, object]], customer=1, credit=False, paid=None,
             invoice_discount=0):                                       # PostSaleFromCart
        prods = []
        for pid, q in lines:
            price, avg, cat = self.c.execute("SELECT SellingPrice, AverageCost, VATCategory FROM Products "
                                             "WHERE ProductID = ?", (pid,)).fetchone()
            prods.append((pid, dec(q), dec(price), dec(avg), cat or "S"))
        out = P.calc([P.LineIn(q, price, D(0), VAT if cat == "S" else D(0)) for _, q, price, _, cat in prods],
                     dec(invoice_discount), True)
        tendered = (D(0) if credit else out.total) if paid is None else dec(paid)
        paid_amt, remaining, change = P.settle(out.total, tendered, credit)
        allow, limit, bal, system = self.c.execute(
            "SELECT AllowCredit, CreditLimit, CurrentBalance, IsSystem FROM Customers WHERE CustomerID = ?",
            (customer,)).fetchone()
        if (credit or remaining > 0) and (not allow or system):
            raise SimError(f"customer {customer} may not buy on credit")
        if (credit or remaining > 0) and dec(limit) > 0 and dec(bal) + remaining > dec(limit):
            raise SimError(f"credit limit of customer {customer}")
        needs: Dict[int, D] = {}
        for pid, q, *_ in prods:
            needs[pid] = needs.get(pid, D(0)) + q
        self.check_stock(needs)
        no = self.next_number("SALES_INVOICE")
        inv = self.insert("SalesInvoices", InvoiceNumber=no, InvoiceDate=self.now, CustomerID=customer,
                          EmployeeID=self.user, PaymentType="CREDIT" if credit else "CASH", PaymentMethodID=1,
                          SubTotal=float(out.subtotal), Discount=float(out.discount),
                          TaxableAmount=float(out.taxable), Tax=float(out.tax), TotalAmount=float(out.total),
                          PaidAmount=float(paid_amt), RemainingAmount=float(remaining),
                          AmountTendered=float(tendered), ChangeDue=float(change))
        for i, ((pid, q, _, avg, cat), ln) in enumerate(zip(prods, out.lines)):
            self.insert("SalesInvoiceDetails", SalesInvoiceID=inv, LineNumber=i + 1, ProductID=pid,
                        Quantity=float(q), UnitPrice=float(ln.unit_price), Discount=float(ln.discount),
                        NetAmount=float(ln.net), VATCategory=cat, VATRate=float(VAT if cat == "S" else 0),
                        Tax=float(ln.tax), LineTotal=float(ln.total), UnitCost=float(avg))
            self.apply_stock(pid, -q, TT["SALE"], avg, "SALE", inv, no)
        if remaining:
            self.adjust("Customers", "CustomerID", customer, remaining)
        return inv

    def sales_return(self, inv, qty_by_line: Dict[int, object], cash_refund=True):   # PostSalesReturn
        customer = self.one("SELECT CustomerID FROM SalesInvoices WHERE SalesInvoiceID = ?", inv)
        if customer == 1:
            cash_refund = True
        rows, total_sum, tax_sum, disc_sum = [], D(0), D(0), D(0)
        for line_no, qty in qty_by_line.items():
            detail, pid, q, total, tax, unit, cost, cat = self.c.execute(
                "SELECT SalesDetailID, ProductID, Quantity, LineTotal, Tax, UnitPrice, UnitCost, VATCategory "
                "FROM SalesInvoiceDetails WHERE SalesInvoiceID = ? AND LineNumber = ?", (inv, line_no)).fetchone()
            prev = self.c.execute("SELECT Sum(Quantity), Sum(LineTotal), Sum(Tax) FROM SalesReturnDetails "
                                  "WHERE SalesDetailID = ?", (detail,)).fetchone()
            pq, pt, px = (dec(v) for v in prev)
            r_net, r_tax, r_total = P.return_amounts(dec(q), dec(total), dec(tax), pq, pt, px, dec(qty))
            disc = max(D(0), P.r2(dec(qty) * dec(unit)) - r_net)
            rows.append((detail, pid, dec(qty), dec(unit), disc, r_net, cat, r_tax, r_total, dec(cost)))
            total_sum += r_total
            tax_sum += r_tax
            disc_sum += disc
        refunded = total_sum if cash_refund else D(0)
        no = self.next_number("SALES_RETURN")
        net_sum = total_sum - tax_sum
        ret = self.insert("SalesReturns", ReturnNumber=no, ReturnDate=self.now, SalesInvoiceID=inv,
                          CustomerID=customer, EmployeeID=self.user, Reason="DEMO",
                          RefundType="CASH" if cash_refund else "CREDIT", SubTotal=float(net_sum + disc_sum),
                          Discount=float(disc_sum), TaxableAmount=float(net_sum), Tax=float(tax_sum),
                          TotalAmount=float(total_sum), RefundedAmount=float(refunded))
        for detail, pid, qty, unit, disc, r_net, cat, r_tax, r_total, cost in rows:
            self.insert("SalesReturnDetails", SalesReturnID=ret, SalesDetailID=detail, ProductID=pid,
                        Quantity=float(qty), UnitPrice=float(unit), Discount=float(disc), NetAmount=float(r_net),
                        VATCategory=cat, VATRate=float(VAT if cat == "S" else 0), Tax=float(r_tax),
                        LineTotal=float(r_total), UnitCost=float(cost), ReturnToStock=1)
            self.apply_stock(pid, qty, TT["SALES_RETURN"], cost, "SALES_RETURN", ret, no, recalc=True)
        if refunded - total_sum:
            self.adjust("Customers", "CustomerID", customer, refunded - total_sum)
        return ret, total_sum

    def customer_payment(self, cid, amount):                            # PostCustomerPayment
        no = self.next_number("CUSTOMER_PAYMENT")
        self.insert("CustomerPayments", PaymentNumber=no, CustomerID=cid, PaymentDate=self.now,
                    Amount=float(amount), PaymentMethodID=1, EmployeeID=self.user)
        self.adjust("Customers", "CustomerID", cid, -dec(amount))

    # ------------------------------------------------------------- inventory
    def manual(self, pid, tt, qty, cost=None):                          # PostManualStock
        no = self.next_number("STOCK_ADJUST")
        if tt == TT["STOCK_OUT"]:
            self.check_stock({pid: dec(qty)})
            self.apply_stock(pid, -dec(qty), tt, self.avg(pid), "MANUAL", None, no)
        else:
            self.apply_stock(pid, dec(qty), tt, dec(cost) if cost is not None else self.avg(pid), "MANUAL", None,
                             no, recalc=True)

    def stock_count(self, category, actual_by_product):                 # CreateStockCount + PostStockCount
        no = self.next_number("STOCK_COUNT")
        cid = self.insert("StockCounts", CountNumber=no, CountDate=self.now, CategoryID=category, Status="OPEN",
                          EmployeeID=self.user)
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
            sys_qty, cost = self.stock(pid), self.avg(pid)
            act = sys_qty if actual is None else dec(actual)
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

    def expense(self, type_id, amount, tax, description, method=1):     # frmExpenses (data screen)
        no = self.next_number("EXPENSE")
        return self.insert("Expenses", ExpenseNumber=no, ExpenseDate=self.now[:10] + " 00:00:00",
                           ExpenseTypeID=type_id, Amount=float(amount), Tax=float(tax),
                           TotalAmount=float(dec(amount) + dec(tax)), PaymentMethodID=method,
                           Description=description, EmployeeID=self.user)
