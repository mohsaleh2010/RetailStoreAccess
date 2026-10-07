"""Phase 4: saved queries (Access SQL), the test fixture and expected results.

Dialect rules (checked by tests/test_queries.py):
  * Access SQL only: IIf, Nz(x, 0), INNER/LEFT JOIN with nested parentheses,
    no CASE / COALESCE / LIMIT / TOP / COUNT(DISTINCT) / || inside saved queries.
  * Numeric Nz() is always wrapped in CCur() - Access otherwise types the
    column as text.
  * An alias never appears inside its own expression (Access "circular
    reference" error) and union branches have the same number of columns.
  * Parameters come from QDate('...') / QLong('...') (modQueryParams).
  * A query may only use queries defined above it (creation order).
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class Query:
    name: str
    caption: str
    sql: str
    params: List[str] = field(default_factory=list)


def period(col):
    return f"{col} >= QDate('PeriodStart') AND {col} < QDate('PeriodEnd')"


def nz(expr):
    return f"CCur(Nz({expr}, 0))"


P = ["PeriodStart", "PeriodEnd"]


# --------------------------------------------------------------------------
# Journal entries: one query per kind of operation gives the lines of its entry
# (SourceType, SourceID, SourceNumber, SourceDate, LineOrder, AccountCode, Debit, Credit, LineText).
# modJournal.SyncJournal creates / refreshes one entry per (SourceType, SourceID).
# --------------------------------------------------------------------------
def cash_account(a):
    """The cash box of a document (110000 + box), else 1190 (cash without a box) or 1200 (bank / card)."""
    return (f"IIf({a}.CashBoxID Is Null, IIf({a}.PaymentMethodID Is Null Or {a}.PaymentMethodID = 1, 1190, 1200), "
            f"110000 + {a}.CashBoxID)")


def jline(stype, key, number, when, party, order, account, debit, credit, text, source, where):
    """party: the customer / supplier / payee of the operation (the entry description)."""
    return (f"SELECT {stype} AS SourceType, {key} AS SourceID, {number} AS SourceNumber, {when} AS SourceDate, "
            f"{party} AS Party, {order} AS LineOrder, {account} AS AccountCode, {debit} AS Debit, "
            f"{credit} AS Credit, {text} AS LineText\nFROM {source}\nWHERE {where}")


def jquery(branches):
    return "\nUNION ALL\n".join(branches)


ZERO = "CCur(0)"
JOURNAL_SOURCE_QUERIES = ["qryJournalSale", "qryJournalSalesReturn", "qryJournalPurchase", "qryJournalPurchaseReturn",
                          "qryJournalPayments", "qryJournalExpense", "qryJournalCashVoucher", "qryJournalStock",
                          "qryJournalOpening", "qryJournalManual"]


def _sale():
    src = "SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID"
    cost = f"({src}) INNER JOIN qrySaleCost AS k ON h.SalesInvoiceID = k.SalesInvoiceID"
    k = ("'SALE'", "h.SalesInvoiceID", "h.InvoiceNumber", "h.InvoiceDate", "c.CustomerName")
    return jquery([
        jline(*k, 1, cash_account("h"), "h.PaidAmount", ZERO, "c.CustomerName", src, "h.PaidAmount <> 0"),
        jline(*k, 2, 1300, "h.RemainingAmount", ZERO, "c.CustomerName", src, "h.RemainingAmount <> 0"),
        jline(*k, 3, 4100, ZERO, "h.TaxableAmount", "'المبيعات'", src, "h.TaxableAmount <> 0"),
        jline(*k, 4, 2200, ZERO, "h.Tax", "'ضريبة المخرجات'", src, "h.Tax <> 0"),
        jline(*k, 5, 5100, "k.SaleCost", ZERO, "'تكلفة البضاعة المباعة'", cost, "k.SaleCost <> 0"),
        jline(*k, 6, 1400, ZERO, "k.SaleCost", "'المخزون'", cost, "k.SaleCost <> 0")])


def _sales_return():
    src = "SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID"
    cost = f"({src}) INNER JOIN qryReturnCost AS k ON r.SalesReturnID = k.SalesReturnID"
    k = ("'SALES_RETURN'", "r.SalesReturnID", "r.ReturnNumber", "r.ReturnDate", "c.CustomerName")
    return jquery([
        jline(*k, 1, 4110, "r.TaxableAmount", ZERO, "'مردودات المبيعات'", src, "r.TaxableAmount <> 0"),
        jline(*k, 2, 2200, "r.Tax", ZERO, "'ضريبة المخرجات'", src, "r.Tax <> 0"),
        jline(*k, 3, cash_account("r"), ZERO, "r.RefundedAmount", "c.CustomerName", src, "r.RefundedAmount <> 0"),
        jline(*k, 4, 1300, ZERO, "r.TotalAmount - r.RefundedAmount", "c.CustomerName", src,
              "r.TotalAmount - r.RefundedAmount <> 0"),
        jline(*k, 5, 1400, "k.ReturnCost", ZERO, "'المخزون'", cost, "k.ReturnCost <> 0"),
        jline(*k, 6, 5100, ZERO, "k.ReturnCost", "'تكلفة البضاعة المباعة'", cost, "k.ReturnCost <> 0")])


def _purchase():
    src = "PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID"
    k = ("'PURCHASE'", "h.PurchaseInvoiceID", "h.InvoiceNumber", "h.InvoiceDate", "s.SupplierName")
    return jquery([
        jline(*k, 1, 1400, "h.TaxableAmount", ZERO, "'المخزون'", src, "h.TaxableAmount <> 0"),
        jline(*k, 2, 1500, "h.Tax", ZERO, "'ضريبة المدخلات'", src, "h.Tax <> 0"),
        jline(*k, 3, cash_account("h"), ZERO, "h.PaidAmount", "s.SupplierName", src, "h.PaidAmount <> 0"),
        jline(*k, 4, 2100, ZERO, "h.RemainingAmount", "s.SupplierName", src, "h.RemainingAmount <> 0")])


def _purchase_return():
    src = "PurchaseReturns AS r INNER JOIN Suppliers AS s ON r.SupplierID = s.SupplierID"
    k = ("'PURCHASE_RETURN'", "r.PurchaseReturnID", "r.ReturnNumber", "r.ReturnDate", "s.SupplierName")
    return jquery([
        jline(*k, 1, cash_account("r"), "r.RefundedAmount", ZERO, "s.SupplierName", src, "r.RefundedAmount <> 0"),
        jline(*k, 2, 2100, "r.TotalAmount - r.RefundedAmount", ZERO, "s.SupplierName", src,
              "r.TotalAmount - r.RefundedAmount <> 0"),
        jline(*k, 3, 1400, ZERO, "r.TaxableAmount", "'المخزون'", src, "r.TaxableAmount <> 0"),
        jline(*k, 4, 1500, ZERO, "r.Tax", "'ضريبة المدخلات'", src, "r.Tax <> 0")])


def _payments():
    cs = "CustomerPayments AS p INNER JOIN Customers AS c ON p.CustomerID = c.CustomerID"
    ss = "SupplierPayments AS p INNER JOIN Suppliers AS s ON p.SupplierID = s.SupplierID"
    kc = ("'CUSTOMER_PAYMENT'", "p.PaymentID", "p.PaymentNumber", "p.PaymentDate", "c.CustomerName")
    ks = ("'SUPPLIER_PAYMENT'", "p.PaymentID", "p.PaymentNumber", "p.PaymentDate", "s.SupplierName")
    return jquery([
        jline(*kc, 1, cash_account("p"), "p.Amount", ZERO, "c.CustomerName", cs, "p.Amount <> 0"),
        jline(*kc, 2, 1300, ZERO, "p.Amount", "c.CustomerName", cs, "p.Amount <> 0"),
        jline(*ks, 1, 2100, "p.Amount", ZERO, "s.SupplierName", ss, "p.Amount <> 0"),
        jline(*ks, 2, cash_account("p"), ZERO, "p.Amount", "s.SupplierName", ss, "p.Amount <> 0")])


def _expense():
    # an expense recorded by a cash voucher is booked by the voucher's entry
    src = ("(Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID) "
           "LEFT JOIN CashVouchers AS v ON e.ExpenseID = v.ExpenseID")
    k = ("'EXPENSE'", "e.ExpenseID", "e.ExpenseNumber", "e.ExpenseDate", "t.ExpenseTypeName")
    return jquery([
        jline(*k, 1, "530000 + e.ExpenseTypeID", "e.Amount", ZERO, "t.ExpenseTypeName", src,
              "v.CashVoucherID Is Null AND e.Amount <> 0"),
        jline(*k, 2, 1500, "e.Tax", ZERO, "'ضريبة المدخلات'", src, "v.CashVoucherID Is Null AND e.Tax <> 0"),
        jline(*k, 3, cash_account("e"), ZERO, "e.TotalAmount", "e.Description", src,
              "v.CashVoucherID Is Null AND e.TotalAmount <> 0")])


def _cash_voucher():
    k = ("'CASH_VOUCHER'", "v.CashVoucherID", "v.VoucherNumber", "v.VoucherDate",
         "IIf(v.PartyName Is Null, v.Description, v.PartyName)")
    out_src = "CashVouchers AS v LEFT JOIN Expenses AS x ON v.ExpenseID = x.ExpenseID"
    out_account = ("IIf(v.Category = 'OWNER', 3100, IIf(v.Category = 'ADVANCE', 1600, IIf(v.Category = 'SHORTAGE', "
                   "5400, IIf(v.Category = 'EXPENSE' AND x.ExpenseTypeID Is Not Null, 530000 + x.ExpenseTypeID, "
                   "5900))))")
    return jquery([
        jline(*k, 1, "110000 + v.CashBoxID", "v.Amount", ZERO, "v.PartyName", "CashVouchers AS v",
              "v.VoucherType = 'IN'"),
        jline(*k, 2, "IIf(v.Category = 'OWNER', 3100, IIf(v.Category = 'OVERAGE', 4300, 4200))", ZERO, "v.Amount",
              "v.Description", "CashVouchers AS v", "v.VoucherType = 'IN'"),
        jline(*k, 1, out_account, "v.Amount", ZERO, "v.Description", out_src, "v.VoucherType = 'OUT'"),
        jline(*k, 2, "110000 + v.CashBoxID", ZERO, "v.Amount", "v.PartyName", out_src, "v.VoucherType = 'OUT'"),
        jline(*k, 1, "110000 + v.ToCashBoxID", "v.Amount", ZERO, "v.Description", "CashVouchers AS v",
              "v.VoucherType = 'TRANSFER'"),
        jline(*k, 2, "110000 + v.CashBoxID", ZERO, "v.Amount", "v.Description", "CashVouchers AS v",
              "v.VoucherType = 'TRANSFER'")])


def _stock():
    val = "i.Quantity * i.UnitCost"
    km = ("'STOCK_MOVE'", "i.TransactionID", "i.ReferenceNumber", "i.TransactionDate", "p.ProductName")
    kc = ("'STOCK_COUNT'", "k.StockCountID", "k.CountNumber", "k.CountDate", "'تسوية الجرد'")
    moves = "InventoryTransactions AS i INNER JOIN Products AS p ON i.ProductID = p.ProductID"
    manual = "i.ReferenceType = 'MANUAL' AND " + val + " <> 0"
    other = "IIf(i.TransactionTypeID = 8, 3900, 5200)"          # 8 = opening stock
    return jquery([
        jline(*km, 1, 1400, f"IIf({val} > 0, {val}, 0)", f"IIf({val} < 0, -{val}, 0)", "i.Notes", moves, manual),
        jline(*km, 2, other, f"IIf({val} < 0, -{val}, 0)", f"IIf({val} > 0, {val}, 0)", "i.Notes", moves, manual),
        jline(*kc, 1, 1400, "IIf(k.CountValue > 0, k.CountValue, 0)", "IIf(k.CountValue < 0, -k.CountValue, 0)",
              "'المخزون'", "qryStockCountValue AS k", "k.CountValue <> 0"),
        jline(*kc, 2, 5200, "IIf(k.CountValue < 0, -k.CountValue, 0)", "IIf(k.CountValue > 0, k.CountValue, 0)",
              "'فروقات الجرد'", "qryStockCountValue AS k", "k.CountValue <> 0")])


def _manual():
    k = ("'MANUAL'", "m.ManualEntryID", "m.EntryNumber", "m.EntryDate", "m.Description")
    return jline(*k, "m.LineNo", "m.LineAccount", "m.LineDebit", "m.LineCredit", "m.LineNote",
                 "qryManualEntryLines AS m", "m.LineDebit + m.LineCredit <> 0")


def _tree_rollup():
    # each account appears once, in the LevelNCode of its own level; the lines of its
    # sub-accounts carry the same code in that column - so every level is summed once
    parts = []
    for n in range(1, 6):
        parts.append(f"""SELECT d.Level{n}Code AS TreeCode, Sum(t.OpeningBalance) AS SumOpening, Sum(t.PeriodDebit) AS SumDebit,
       Sum(t.PeriodCredit) AS SumCredit, Sum(t.ClosingBalance) AS SumClosing
FROM TrialBalanceQuery AS t INNER JOIN Accounts AS d ON t.AccountCode = d.AccountCode
WHERE d.Level{n}Code Is Not Null
GROUP BY d.Level{n}Code""")
    return "\nUNION ALL\n".join(parts)


def IN_TREE(alias):
    """The account of the line is the chosen account or one of its sub-accounts."""
    return "(" + " OR ".join(f"{alias}.Level{n}Code = QLong('AccountCode')" for n in range(1, 6)) + ")"


ACCOUNT_TYPE_NAME = ("IIf(a.AccountType = 'ASSET', 'أصول', IIf(a.AccountType = 'LIABILITY', 'خصوم', "
                     "IIf(a.AccountType = 'EQUITY', 'حقوق ملكية', IIf(a.AccountType = 'REVENUE', 'إيرادات', "
                     "'مصروفات'))))")


def _opening():
    kb = ("'BOX_OPENING'", "b.CashBoxID", "b.BoxName", "b.OpeningDate", "b.BoxName")
    kc = ("'CUSTOMER_OPENING'", "c.CustomerID", "c.CustomerName", "c.CreatedAt", "c.CustomerName")
    ks = ("'SUPPLIER_OPENING'", "s.SupplierID", "s.SupplierName", "s.CreatedAt", "s.SupplierName")
    pos, neg = "IIf({0}.OpeningBalance > 0, {0}.OpeningBalance, 0)", "IIf({0}.OpeningBalance < 0, -{0}.OpeningBalance, 0)"
    return jquery([
        jline(*kb, 1, "110000 + b.CashBoxID", "b.OpeningBalance", ZERO, "b.BoxName", "CashBoxes AS b",
              "b.OpeningBalance <> 0"),
        jline(*kb, 2, 3900, ZERO, "b.OpeningBalance", "'رصيد افتتاحي'", "CashBoxes AS b", "b.OpeningBalance <> 0"),
        jline(*kc, 1, 1300, pos.format("c"), neg.format("c"), "c.CustomerName", "Customers AS c",
              "c.OpeningBalance <> 0"),
        jline(*kc, 2, 3900, neg.format("c"), pos.format("c"), "'رصيد افتتاحي'", "Customers AS c",
              "c.OpeningBalance <> 0"),
        jline(*ks, 1, 2100, neg.format("s"), pos.format("s"), "s.SupplierName", "Suppliers AS s",
              "s.OpeningBalance <> 0"),
        jline(*ks, 2, 3900, pos.format("s"), neg.format("s"), "'رصيد افتتاحي'", "Suppliers AS s",
              "s.OpeningBalance <> 0")])

QUERIES: List[Query] = [

    # ================================================================ SALES
    Query("qrySalesDocuments", "مستندات البيع: الفواتير (+) والمرتجعات (−) بقيم موقّعة", """
SELECT 'SALE' AS DocType, h.SalesInvoiceID AS DocID, h.InvoiceNumber AS DocNumber,
       h.InvoiceDate AS DocDate, h.CustomerID, h.EmployeeID, h.PaymentType,
       h.TaxableAmount AS NetAmount, h.Tax AS VATAmount, h.TotalAmount AS GrossAmount,
       h.PaidAmount AS SettledAmount, h.RemainingAmount AS OnAccount
FROM SalesInvoices AS h
UNION ALL
SELECT 'RETURN', r.SalesReturnID, r.ReturnNumber, r.ReturnDate, r.CustomerID, r.EmployeeID,
       r.RefundType, -r.TaxableAmount, -r.Tax, -r.TotalAmount, -r.RefundedAmount,
       -(r.TotalAmount - r.RefundedAmount)
FROM SalesReturns AS r"""),

    Query("qrySalesLineItems", "أسطر البيع والمرتجعات مع التكلفة (أساس تحليل المنتجات والأرباح)", """
SELECT 'SALE' AS DocType, h.SalesInvoiceID AS DocID, h.InvoiceNumber AS DocNumber,
       h.InvoiceDate AS DocDate, h.CustomerID, d.ProductID, d.Quantity AS SignedQty,
       d.Quantity AS SoldQty, CCur(0) AS ReturnedQty, d.NetAmount AS LineNet,
       d.Tax AS LineVAT, d.LineTotal AS LineGross, d.Quantity * d.UnitCost AS LineCost
FROM SalesInvoices AS h INNER JOIN SalesInvoiceDetails AS d
     ON h.SalesInvoiceID = d.SalesInvoiceID
UNION ALL
SELECT 'RETURN', r.SalesReturnID, r.ReturnNumber, r.ReturnDate, r.CustomerID, d.ProductID,
       -d.Quantity, CCur(0), d.Quantity, -d.NetAmount, -d.Tax, -d.LineTotal,
       IIf(d.ReturnToStock, -d.Quantity * d.UnitCost, 0)
FROM SalesReturns AS r INNER JOIN SalesReturnDetails AS d
     ON r.SalesReturnID = d.SalesReturnID"""),

    Query("qrySalesLinesInPeriod", "أسطر البيع والمرتجعات داخل الفترة", f"""
SELECT * FROM qrySalesLineItems
WHERE {period("DocDate")}""", P),

    Query("DailySalesQuery", "المبيعات اليومية: عدد الفواتير والمرتجعات والصافي والنقدي والآجل", """
SELECT DateValue(DocDate) AS SaleDate,
       Sum(IIf(DocType = 'SALE', 1, 0)) AS InvoiceCount,
       Sum(IIf(DocType = 'RETURN', 1, 0)) AS ReturnCount,
       Sum(IIf(DocType = 'SALE', NetAmount, 0)) AS SalesExVAT,
       -Sum(IIf(DocType = 'RETURN', NetAmount, 0)) AS ReturnsExVAT,
       Sum(NetAmount) AS NetSalesExVAT, Sum(VATAmount) AS NetVAT,
       Sum(GrossAmount) AS NetSalesTotal,
       Sum(IIf(DocType = 'SALE' AND PaymentType = 'CASH', GrossAmount, 0)) AS CashSales,
       Sum(IIf(DocType = 'SALE' AND PaymentType = 'CREDIT', GrossAmount, 0)) AS CreditSales,
       Sum(SettledAmount) AS CollectedAmount
FROM qrySalesDocuments
GROUP BY DateValue(DocDate)
ORDER BY DateValue(DocDate) DESC"""),

    Query("qrySalesMonthlyDocs", "تجميع شهري لمستندات البيع", """
SELECT Year(DocDate) AS SalesYear, Month(DocDate) AS SalesMonth,
       Sum(IIf(DocType = 'SALE', 1, 0)) AS InvoiceCount,
       Sum(IIf(DocType = 'RETURN', 1, 0)) AS ReturnCount,
       Sum(NetAmount) AS NetSalesExVAT, Sum(VATAmount) AS NetVAT,
       Sum(GrossAmount) AS NetSalesTotal
FROM qrySalesDocuments
GROUP BY Year(DocDate), Month(DocDate)"""),

    Query("qrySalesMonthlyCost", "تكلفة البضاعة المباعة شهريًا", """
SELECT Year(DocDate) AS SalesYear, Month(DocDate) AS SalesMonth, Sum(LineCost) AS MonthCost
FROM qrySalesLineItems
GROUP BY Year(DocDate), Month(DocDate)"""),

    Query("MonthlySalesQuery", "المبيعات الشهرية مع التكلفة ومجمل الربح", f"""
SELECT d.SalesYear, d.SalesMonth, d.InvoiceCount, d.ReturnCount, d.NetSalesExVAT,
       d.NetVAT, d.NetSalesTotal, {nz("c.MonthCost")} AS CostOfSales,
       d.NetSalesExVAT - {nz("c.MonthCost")} AS GrossProfit
FROM qrySalesMonthlyDocs AS d LEFT JOIN qrySalesMonthlyCost AS c
     ON (d.SalesYear = c.SalesYear AND d.SalesMonth = c.SalesMonth)
ORDER BY d.SalesYear DESC, d.SalesMonth DESC"""),

    Query("SalesByPeriodQuery", "فواتير ومرتجعات البيع خلال فترة مع العميل والكاشير", f"""
SELECT s.DocType, s.DocID, s.DocNumber, s.DocDate, s.CustomerID, c.CustomerName,
       e.EmployeeName, s.PaymentType, s.NetAmount, s.VATAmount, s.GrossAmount,
       s.SettledAmount, s.OnAccount
FROM (qrySalesDocuments AS s INNER JOIN Customers AS c ON s.CustomerID = c.CustomerID)
     INNER JOIN Employees AS e ON s.EmployeeID = e.EmployeeID
WHERE {period("s.DocDate")}
ORDER BY s.DocDate""", P),

    Query("SalesByProductQuery", "المبيعات حسب المنتج خلال فترة (كمية، صافي، تكلفة، ربح)", """
SELECT p.ProductID, p.ProductCode, p.ProductName, c.CategoryName,
       Sum(l.SoldQty) AS QtySold, Sum(l.ReturnedQty) AS QtyReturned, Sum(l.SignedQty) AS NetQty,
       Sum(l.LineNet) AS NetSales, Sum(l.LineVAT) AS SalesVAT, Sum(l.LineGross) AS SalesTotal,
       Sum(l.LineCost) AS CostOfSales, Sum(l.LineNet) - Sum(l.LineCost) AS GrossProfit
FROM (qrySalesLinesInPeriod AS l INNER JOIN Products AS p ON l.ProductID = p.ProductID)
     INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID
GROUP BY p.ProductID, p.ProductCode, p.ProductName, c.CategoryName
ORDER BY p.ProductName""", P),

    Query("BestSellingProductsQuery", "أفضل المنتجات مبيعًا خلال فترة (حسب صافي الكمية)", """
SELECT * FROM SalesByProductQuery
ORDER BY NetQty DESC, NetSales DESC""", P),

    Query("SalesByCategoryQuery", "المبيعات حسب التصنيف خلال فترة (للرسم الدائري)", """
SELECT CategoryName, Count(*) AS ProductCount, Sum(NetQty) AS CategoryQty,
       Sum(NetSales) AS CategoryNet, Sum(SalesTotal) AS CategoryTotal, Sum(GrossProfit) AS CategoryProfit
FROM SalesByProductQuery
GROUP BY CategoryName
ORDER BY Sum(SalesTotal) DESC""", P),

    Query("LeastSellingProductsQuery", "أقل المنتجات مبيعًا خلال فترة (تشمل المنتجات التي لم تُبع)", f"""
SELECT p.ProductID, p.ProductCode, p.ProductName, c.CategoryName, p.CurrentQuantity,
       {nz("s.NetQty")} AS NetQtySold, {nz("s.NetSales")} AS NetSalesAmount
FROM (Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID)
     LEFT JOIN SalesByProductQuery AS s ON p.ProductID = s.ProductID
WHERE p.IsActive = True
ORDER BY {nz("s.NetQty")}, p.ProductName""", P),

    # ============================================================ PURCHASES
    Query("qryPurchaseDocuments", "مستندات الشراء: الفواتير (+) والمرتجعات (−)", """
SELECT 'PURCHASE' AS DocType, h.PurchaseInvoiceID AS DocID, h.InvoiceNumber AS DocNumber,
       h.SupplierInvoiceNo AS SupplierRef, h.InvoiceDate AS DocDate, h.SupplierID,
       h.PaymentType, h.TaxableAmount AS NetAmount, h.Tax AS VATAmount,
       h.TotalAmount AS GrossAmount, h.PaidAmount AS SettledAmount,
       h.RemainingAmount AS OnAccount
FROM PurchaseInvoices AS h
UNION ALL
SELECT 'RETURN', r.PurchaseReturnID, r.ReturnNumber, Null, r.ReturnDate, r.SupplierID,
       r.RefundType, -r.TaxableAmount, -r.Tax, -r.TotalAmount, -r.RefundedAmount,
       -(r.TotalAmount - r.RefundedAmount)
FROM PurchaseReturns AS r"""),

    Query("PurchasesQuery", "فواتير ومرتجعات الشراء خلال فترة مع المورد", f"""
SELECT d.DocType, d.DocID, d.DocNumber, d.SupplierRef, d.DocDate, d.SupplierID,
       s.SupplierName, d.PaymentType, d.NetAmount, d.VATAmount, d.GrossAmount,
       d.SettledAmount, d.OnAccount
FROM qryPurchaseDocuments AS d INNER JOIN Suppliers AS s ON d.SupplierID = s.SupplierID
WHERE {period("d.DocDate")}
ORDER BY d.DocDate""", P),

    # ============================================================ INVENTORY
    Query("qryProductLedger", "رصيد كل منتج من دفتر حركة المخزون", """
SELECT ProductID, Sum(Quantity) AS LedgerQty, Max(TransactionDate) AS LastMovementDate
FROM InventoryTransactions
GROUP BY ProductID"""),

    Query("qryProductLastSale", "تاريخ آخر بيع لكل منتج", """
SELECT d.ProductID, Max(h.InvoiceDate) AS LastSaleDate
FROM SalesInvoiceDetails AS d INNER JOIN SalesInvoices AS h
     ON d.SalesInvoiceID = h.SalesInvoiceID
GROUP BY d.ProductID"""),

    Query("StockBalanceQuery", "المخزون الحالي: الكمية والقيمة بالتكلفة وبسعر البيع ومطابقتها مع الحركات", f"""
SELECT p.ProductID, p.ProductCode, p.Barcode, p.ProductName, c.CategoryName, u.UnitName,
       p.ProductLocation, p.CurrentQuantity, p.MinimumQuantity, p.AverageCost, p.SellingPrice,
       p.CurrentQuantity * p.AverageCost AS StockCostValue,
       p.CurrentQuantity * p.SellingPrice AS StockSalesValue,
       {nz("l.LedgerQty")} AS LedgerQuantity,
       p.CurrentQuantity - {nz("l.LedgerQty")} AS QuantityMismatch,
       IIf(p.CurrentQuantity <= p.MinimumQuantity, True, False) AS IsLowStock, p.IsActive
FROM ((Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID)
      INNER JOIN Units AS u ON p.UnitID = u.UnitID)
     LEFT JOIN qryProductLedger AS l ON p.ProductID = l.ProductID
ORDER BY p.ProductName"""),

    Query("LowStockQuery", "المنتجات منخفضة المخزون: CurrentQuantity <= MinimumQuantity", """
SELECT p.ProductID, p.ProductCode, p.Barcode, p.ProductName, c.CategoryName,
       p.CurrentQuantity, p.MinimumQuantity, p.MinimumQuantity - p.CurrentQuantity AS ShortageQty,
       s.SupplierName, s.Mobile AS SupplierMobile
FROM (Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID)
     LEFT JOIN Suppliers AS s ON p.SupplierID = s.SupplierID
WHERE p.IsActive = True AND p.TrackStock = True AND p.CurrentQuantity <= p.MinimumQuantity
ORDER BY p.MinimumQuantity - p.CurrentQuantity DESC, p.ProductName"""),

    Query("ProductMovementQuery", "حركة منتج خلال فترة مع رصيد أول المدة (الرصيد التراكمي في التقرير)", f"""
SELECT 1 AS SortKey, t.TransactionID, t.TransactionDate AS MovementDate,
       tt.TypeName AS MovementType, t.ReferenceNumber,
       IIf(t.Quantity > 0, t.Quantity, 0) AS QtyIn, IIf(t.Quantity < 0, -t.Quantity, 0) AS QtyOut,
       t.Quantity AS NetQty, t.UnitCost, t.Notes
FROM InventoryTransactions AS t INNER JOIN TransactionTypes AS tt
     ON t.TransactionTypeID = tt.TransactionTypeID
WHERE t.ProductID = QLong('ProductID') AND {period("t.TransactionDate")}
UNION ALL
SELECT 0, 0, QDate('PeriodStart'), 'رصيد أول المدة', Null,
       IIf({nz("Sum(o.Quantity)")} > 0, {nz("Sum(o.Quantity)")}, 0),
       IIf({nz("Sum(o.Quantity)")} < 0, -{nz("Sum(o.Quantity)")}, 0),
       {nz("Sum(o.Quantity)")}, Null, Null
FROM InventoryTransactions AS o
WHERE o.ProductID = QLong('ProductID') AND o.TransactionDate < QDate('PeriodStart')
ORDER BY SortKey, MovementDate, TransactionID""", P + ["ProductID"]),

    Query("SlowMovingProductsQuery", "المنتجات غير المتحركة: لها رصيد ولم تُبع منذ عدد الأيام المحدد في الإعدادات", """
SELECT p.ProductID, p.ProductCode, p.ProductName, c.CategoryName, p.CurrentQuantity,
       p.AverageCost, p.CurrentQuantity * p.AverageCost AS StockCostValue, ls.LastSaleDate,
       DateDiff('d', Nz(ls.LastSaleDate, p.CreatedAt), Date()) AS DaysWithoutSale
FROM (Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID)
     LEFT JOIN qryProductLastSale AS ls ON p.ProductID = ls.ProductID
WHERE p.IsActive = True AND p.CurrentQuantity > 0
  AND DateDiff('d', Nz(ls.LastSaleDate, p.CreatedAt), Date()) >=
      (SELECT SlowMovingDays FROM Settings)
ORDER BY DateDiff('d', Nz(ls.LastSaleDate, p.CreatedAt), Date()) DESC"""),

    Query("StockByCategoryQuery", "المخزون حسب التصنيف: عدد المنتجات والكمية والقيمة", """
SELECT c.CategoryID, c.CategoryName, Count(*) AS ProductCount,
       Sum(p.CurrentQuantity) AS TotalQuantity,
       Sum(p.CurrentQuantity * p.AverageCost) AS StockCostValue,
       Sum(p.CurrentQuantity * p.SellingPrice) AS StockSalesValue
FROM Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID
WHERE p.IsActive = True
GROUP BY c.CategoryID, c.CategoryName
ORDER BY c.CategoryName"""),

    Query("StockCountQuery", "تفاصيل جلسات الجرد: الكمية المسجلة والفعلية والفرق وقيمته", """
SELECT c.StockCountID, c.CountNumber, c.CountDate, c.Status, c.CategoryID, g.CategoryName,
       d.ProductID, p.ProductCode, p.ProductName, d.SystemQuantity, d.ActualQuantity, d.Difference,
       d.UnitCost, d.DifferenceValue, d.Notes
FROM ((StockCountDetails AS d INNER JOIN StockCounts AS c ON d.StockCountID = c.StockCountID)
      INNER JOIN Products AS p ON d.ProductID = p.ProductID)
     LEFT JOIN Categories AS g ON c.CategoryID = g.CategoryID
ORDER BY c.StockCountID, p.ProductName"""),

    # ============================================================ CUSTOMERS
    Query("qryCustomerLedger", "دفتر حساب العملاء: مدين (عليه) / دائن (له)", """
SELECT h.CustomerID, h.InvoiceDate AS EntryDate, 'SALE' AS EntryType,
       'فاتورة بيع' AS EntryTypeName, h.SalesInvoiceID AS DocID, h.InvoiceNumber AS DocNumber,
       h.TotalAmount AS Debit, h.PaidAmount AS Credit
FROM SalesInvoices AS h
UNION ALL
SELECT r.CustomerID, r.ReturnDate, 'SALES_RETURN', 'مرتجع بيع', r.SalesReturnID, r.ReturnNumber,
       r.RefundedAmount, r.TotalAmount
FROM SalesReturns AS r
UNION ALL
SELECT p.CustomerID, p.PaymentDate, 'PAYMENT', 'سند قبض', p.PaymentID, p.PaymentNumber,
       CCur(0), p.Amount
FROM CustomerPayments AS p
UNION ALL
SELECT c.CustomerID, c.CreatedAt, 'OPENING', 'رصيد افتتاحي', 0, '-',
       IIf(c.OpeningBalance > 0, c.OpeningBalance, 0),
       IIf(c.OpeningBalance < 0, -c.OpeningBalance, 0)
FROM Customers AS c
WHERE c.OpeningBalance <> 0"""),

    Query("qryCustomerLedgerTotals", "مجاميع حساب كل عميل", """
SELECT CustomerID, Sum(Debit) AS TotalDebit, Sum(Credit) AS TotalCredit,
       Max(EntryDate) AS LastEntryDate
FROM qryCustomerLedger
GROUP BY CustomerID"""),

    Query("CustomerBalanceQuery", "رصيد كل عميل محسوبًا من الحركات (موجب = عليه للمحل)", f"""
SELECT c.CustomerID, c.CustomerName, c.Mobile, c.CreditLimit, c.AllowCredit, c.IsActive,
       {nz("l.TotalDebit")} AS DebitTotal, {nz("l.TotalCredit")} AS CreditTotal,
       {nz("l.TotalDebit")} - {nz("l.TotalCredit")} AS Balance,
       c.CurrentBalance AS CachedBalance, l.LastEntryDate
FROM Customers AS c LEFT JOIN qryCustomerLedgerTotals AS l ON c.CustomerID = l.CustomerID
ORDER BY c.CustomerName"""),

    Query("CustomersWithDebtQuery", "العملاء الذين عليهم مبالغ مستحقة", """
SELECT * FROM CustomerBalanceQuery
WHERE Balance > 0
ORDER BY Balance DESC"""),

    Query("CustomerStatementQuery", "كشف حساب عميل لفترة: رصيد سابق ثم الحركات", f"""
SELECT 1 AS SortKey, l.EntryDate, l.EntryType, l.EntryTypeName, l.DocNumber, l.Debit, l.Credit
FROM qryCustomerLedger AS l
WHERE l.CustomerID = QLong('CustomerID') AND {period("l.EntryDate")}
UNION ALL
SELECT 0, QDate('PeriodStart'), 'BALANCE_FWD', 'رصيد سابق', '-',
       IIf({nz("Sum(o.Debit)")} - {nz("Sum(o.Credit)")} > 0, {nz("Sum(o.Debit)")} - {nz("Sum(o.Credit)")}, 0),
       IIf({nz("Sum(o.Debit)")} - {nz("Sum(o.Credit)")} < 0, {nz("Sum(o.Credit)")} - {nz("Sum(o.Debit)")}, 0)
FROM qryCustomerLedger AS o
WHERE o.CustomerID = QLong('CustomerID') AND o.EntryDate < QDate('PeriodStart')
ORDER BY SortKey, EntryDate""", P + ["CustomerID"]),

    # ============================================================ SUPPLIERS
    Query("qrySupplierLedger", "دفتر حساب الموردين: دائن (للمورد) / مدين (سُدِّد له)", """
SELECT h.SupplierID, h.InvoiceDate AS EntryDate, 'PURCHASE' AS EntryType,
       'فاتورة شراء' AS EntryTypeName, h.PurchaseInvoiceID AS DocID, h.InvoiceNumber AS DocNumber,
       h.PaidAmount AS Debit, h.TotalAmount AS Credit
FROM PurchaseInvoices AS h
UNION ALL
SELECT r.SupplierID, r.ReturnDate, 'PURCHASE_RETURN', 'مرتجع شراء', r.PurchaseReturnID,
       r.ReturnNumber, r.TotalAmount, r.RefundedAmount
FROM PurchaseReturns AS r
UNION ALL
SELECT p.SupplierID, p.PaymentDate, 'PAYMENT', 'سند صرف', p.PaymentID, p.PaymentNumber,
       p.Amount, CCur(0)
FROM SupplierPayments AS p
UNION ALL
SELECT s.SupplierID, s.CreatedAt, 'OPENING', 'رصيد افتتاحي', 0, '-',
       IIf(s.OpeningBalance < 0, -s.OpeningBalance, 0),
       IIf(s.OpeningBalance > 0, s.OpeningBalance, 0)
FROM Suppliers AS s
WHERE s.OpeningBalance <> 0"""),

    Query("qrySupplierLedgerTotals", "مجاميع حساب كل مورد", """
SELECT SupplierID, Sum(Debit) AS TotalDebit, Sum(Credit) AS TotalCredit,
       Max(EntryDate) AS LastEntryDate
FROM qrySupplierLedger
GROUP BY SupplierID"""),

    Query("SupplierBalanceQuery", "رصيد كل مورد محسوبًا من الحركات (موجب = مستحق للمورد)", f"""
SELECT s.SupplierID, s.SupplierName, s.ContactPerson, s.Mobile, s.IsActive,
       {nz("l.TotalDebit")} AS DebitTotal, {nz("l.TotalCredit")} AS CreditTotal,
       {nz("l.TotalCredit")} - {nz("l.TotalDebit")} AS Balance,
       s.CurrentBalance AS CachedBalance, l.LastEntryDate
FROM Suppliers AS s LEFT JOIN qrySupplierLedgerTotals AS l ON s.SupplierID = l.SupplierID
ORDER BY s.SupplierName"""),

    Query("SupplierStatementQuery", "كشف حساب مورد لفترة: رصيد سابق ثم الحركات", f"""
SELECT 1 AS SortKey, l.EntryDate, l.EntryType, l.EntryTypeName, l.DocNumber, l.Debit, l.Credit
FROM qrySupplierLedger AS l
WHERE l.SupplierID = QLong('SupplierID') AND {period("l.EntryDate")}
UNION ALL
SELECT 0, QDate('PeriodStart'), 'BALANCE_FWD', 'رصيد سابق', '-',
       IIf({nz("Sum(o.Debit)")} - {nz("Sum(o.Credit)")} > 0, {nz("Sum(o.Debit)")} - {nz("Sum(o.Credit)")}, 0),
       IIf({nz("Sum(o.Debit)")} - {nz("Sum(o.Credit)")} < 0, {nz("Sum(o.Credit)")} - {nz("Sum(o.Debit)")}, 0)
FROM qrySupplierLedger AS o
WHERE o.SupplierID = QLong('SupplierID') AND o.EntryDate < QDate('PeriodStart')
ORDER BY SortKey, EntryDate""", P + ["SupplierID"]),

    # ============================================================= EXPENSES
    Query("ExpensesQuery", "المصروفات خلال فترة", f"""
SELECT e.ExpenseID, e.ExpenseNumber, e.ExpenseDate, t.ExpenseTypeName, e.Amount, e.Tax,
       e.TotalAmount, pm.MethodName, e.Description, em.EmployeeName, e.ExpenseTypeID
FROM ((Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID)
      INNER JOIN Employees AS em ON e.EmployeeID = em.EmployeeID)
     LEFT JOIN PaymentMethods AS pm ON e.PaymentMethodID = pm.PaymentMethodID
WHERE {period("e.ExpenseDate")}
ORDER BY e.ExpenseDate""", P),

    Query("ExpensesByTypeQuery", "المصروفات مجمّعة حسب النوع خلال فترة", f"""
SELECT t.ExpenseTypeName, Count(*) AS ExpenseCount, Sum(e.Amount) AS AmountExVAT,
       Sum(e.Tax) AS InputVAT, Sum(e.TotalAmount) AS AmountTotal
FROM Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID
WHERE {period("e.ExpenseDate")}
GROUP BY t.ExpenseTypeName
ORDER BY Sum(e.TotalAmount) DESC""", P),

    # =============================================================== PROFIT
    Query("qryProfitSales", "صافي المبيعات وتكلفتها خلال الفترة", f"""
SELECT {nz("Sum(LineNet)")} AS PeriodNetSales, {nz("Sum(LineCost)")} AS PeriodCost
FROM qrySalesLinesInPeriod""", P),

    Query("qryProfitAdjustments", "قيمة فروقات المخزون (جرد، إضافة، خصم) خلال الفترة", f"""
SELECT {nz("Sum(t.Quantity * t.UnitCost)")} AS PeriodAdjustments
FROM InventoryTransactions AS t INNER JOIN TransactionTypes AS tt
     ON t.TransactionTypeID = tt.TransactionTypeID
WHERE tt.TypeCode IN ('STOCK_IN', 'STOCK_OUT', 'ADJUSTMENT') AND {period("t.TransactionDate")}""", P),

    Query("qryProfitExpenses", "المصروفات (بدون ضريبة) خلال الفترة", f"""
SELECT {nz("Sum(Amount)")} AS PeriodExpenses
FROM Expenses
WHERE {period("ExpenseDate")}""", P),

    Query("ProfitQuery", "الأرباح: صافي المبيعات − التكلفة = مجمل الربح؛ ثم ± فروقات المخزون − المصروفات = صافي الربح", """
SELECT QDate('PeriodStart') AS PeriodFrom, DateAdd('d', -1, QDate('PeriodEnd')) AS PeriodTo,
       s.PeriodNetSales AS NetSales, s.PeriodCost AS CostOfSales,
       s.PeriodNetSales - s.PeriodCost AS GrossProfit,
       a.PeriodAdjustments AS InventoryAdjustments, x.PeriodExpenses AS TotalExpenses,
       s.PeriodNetSales - s.PeriodCost + a.PeriodAdjustments - x.PeriodExpenses AS NetProfit
FROM qryProfitSales AS s, qryProfitAdjustments AS a, qryProfitExpenses AS x""", P),

    # ================================================================== VAT
    Query("qryVatOutput", "ضريبة المخرجات (المبيعات ناقص المرتجعات)", f"""
SELECT {nz("Sum(LineNet)")} AS TaxableSales, {nz("Sum(LineVAT)")} AS OutputVAT
FROM qrySalesLinesInPeriod""", P),

    Query("qryVatInputPurchases", "ضريبة المدخلات من المشتريات (ناقص المرتجعات)", f"""
SELECT {nz("Sum(NetAmount)")} AS TaxablePurchases, {nz("Sum(VATAmount)")} AS PurchaseVAT
FROM qryPurchaseDocuments
WHERE {period("DocDate")}""", P),

    Query("qryVatInputExpenses", "ضريبة المدخلات من المصروفات", f"""
SELECT {nz("Sum(Tax)")} AS ExpenseVAT
FROM Expenses
WHERE {period("ExpenseDate")}""", P),

    Query("VatSummaryQuery", "ملخص ضريبة القيمة المضافة للفترة (للإقرار الضريبي)", """
SELECT QDate('PeriodStart') AS PeriodFrom, DateAdd('d', -1, QDate('PeriodEnd')) AS PeriodTo,
       o.TaxableSales, o.OutputVAT, p.TaxablePurchases, p.PurchaseVAT, e.ExpenseVAT,
       p.PurchaseVAT + e.ExpenseVAT AS InputVAT,
       o.OutputVAT - p.PurchaseVAT - e.ExpenseVAT AS NetVATDue
FROM qryVatOutput AS o, qryVatInputPurchases AS p, qryVatInputExpenses AS e""", P),

    # ============================================================ DASHBOARD
    # Own parameters (set by modDashboard), so the dashboard never changes the
    # period chosen in the report centre: DashDay = today, DashMonth = first day
    # of the month, DashEnd = tomorrow (exclusive).
    Query("DashboardQuery", "مؤشرات لوحة التحكم في سجل واحد (اليوم، الشهر، الأرصدة، المخزون)", f"""
SELECT (SELECT {nz("Sum(d.GrossAmount)")} FROM qrySalesDocuments AS d
        WHERE d.DocDate >= QDate('DashDay') AND d.DocDate < QDate('DashEnd')) AS TodaySales,
       (SELECT Count(*) FROM SalesInvoices AS h
        WHERE h.InvoiceDate >= QDate('DashDay') AND h.InvoiceDate < QDate('DashEnd')) AS TodayInvoices,
       (SELECT {nz("Sum(d.GrossAmount)")} FROM qrySalesDocuments AS d
        WHERE d.DocDate >= QDate('DashMonth') AND d.DocDate < QDate('DashEnd')) AS MonthSales,
       (SELECT {nz("Sum(d.VATAmount)")} FROM qrySalesDocuments AS d
        WHERE d.DocDate >= QDate('DashMonth') AND d.DocDate < QDate('DashEnd')) AS MonthVAT,
       (SELECT Count(*) FROM SalesInvoices AS h
        WHERE h.InvoiceDate >= QDate('DashMonth') AND h.InvoiceDate < QDate('DashEnd')) AS MonthInvoices,
       (SELECT {nz("Sum(e.Amount)")} FROM Expenses AS e
        WHERE e.ExpenseDate >= QDate('DashMonth') AND e.ExpenseDate < QDate('DashEnd')) AS MonthExpenses,
       (SELECT {nz("Sum(c.CurrentBalance)")} FROM Customers AS c WHERE c.CurrentBalance > 0) AS CustomerDebt,
       (SELECT Count(*) FROM Customers AS c WHERE c.CurrentBalance > 0) AS DebtorCount,
       (SELECT {nz("Sum(s.CurrentBalance)")} FROM Suppliers AS s WHERE s.CurrentBalance > 0) AS SupplierDue,
       (SELECT {nz("Sum(p.CurrentQuantity * p.AverageCost)")} FROM Products AS p
        WHERE p.IsActive = True AND p.CurrentQuantity > 0) AS StockValue,
       (SELECT Count(*) FROM LowStockQuery) AS LowStockCount
FROM Settings AS st
WHERE st.SettingID = 1""", ["DashDay", "DashMonth", "DashEnd"]),

    Query("qryDashboardTopProducts", "صافي الكمية المباعة لكل منتج منذ بداية الشهر (لوحة التحكم)", """
SELECT l.ProductID, p.ProductName, Sum(l.SignedQty) AS NetQty, Sum(l.LineGross) AS NetSales
FROM qrySalesLineItems AS l INNER JOIN Products AS p ON l.ProductID = p.ProductID
WHERE l.DocDate >= QDate('DashMonth') AND l.DocDate < QDate('DashEnd')
GROUP BY l.ProductID, p.ProductName
HAVING Sum(l.SignedQty) > 0""", ["DashMonth", "DashEnd"]),

    # ============================================================ PRINTING
    Query("qrySalesDocPrint", "بيانات طباعة فواتير البيع والإشعارات الدائنة (سطر لكل صنف)", """
SELECT 'SALE' AS DocKind, h.SalesInvoiceID AS DocID, h.InvoiceNumber AS DocNumber,
       h.InvoiceDate AS DocDate, '' AS OriginalNumber, h.InvoiceSubType, h.PaymentType,
       h.CustomerID, c.CustomerName, c.VATNumber AS CustomerVAT, c.City AS CustomerCity,
       c.District AS CustomerDistrict, c.StreetName AS CustomerStreet,
       c.BuildingNo AS CustomerBuilding, c.PostalCode AS CustomerPostal, e.EmployeeName,
       h.SubTotal AS DocSubTotal, h.Discount AS DocDiscount, h.TaxableAmount, h.Tax AS DocTax,
       h.TotalAmount, h.PaidAmount, h.RemainingAmount, h.AmountTendered, h.ChangeDue,
       d.LineNumber, p.ProductName, p.ProductCode, u.UnitName, d.Quantity, d.UnitPrice,
       d.Discount AS LineDiscount, d.NetAmount, d.VATRate, d.Tax AS LineTax, d.LineTotal,
       h.OrderType, h.TableNo, h.OrderName, d.LineNote
FROM ((((SalesInvoices AS h INNER JOIN SalesInvoiceDetails AS d ON h.SalesInvoiceID = d.SalesInvoiceID)
       INNER JOIN Products AS p ON d.ProductID = p.ProductID)
      INNER JOIN Units AS u ON p.UnitID = u.UnitID)
     INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID)
    INNER JOIN Employees AS e ON h.EmployeeID = e.EmployeeID
UNION ALL
SELECT 'RETURN', r.SalesReturnID, r.ReturnNumber, r.ReturnDate, o.InvoiceNumber, r.InvoiceSubType,
       r.RefundType, r.CustomerID, c.CustomerName, c.VATNumber, c.City, c.District, c.StreetName,
       c.BuildingNo, c.PostalCode, e.EmployeeName, r.SubTotal, r.Discount, r.TaxableAmount, r.Tax,
       r.TotalAmount, r.RefundedAmount, r.TotalAmount - r.RefundedAmount, CCur(0), CCur(0),
       rd.ReturnDetailID, p.ProductName, p.ProductCode, u.UnitName, rd.Quantity, rd.UnitPrice,
       rd.Discount, rd.NetAmount, rd.VATRate, rd.Tax, rd.LineTotal, o.OrderType, o.TableNo, o.OrderName,
       od.LineNote
FROM ((((((SalesReturns AS r INNER JOIN SalesReturnDetails AS rd ON r.SalesReturnID = rd.SalesReturnID)
        INNER JOIN SalesInvoices AS o ON r.SalesInvoiceID = o.SalesInvoiceID)
        INNER JOIN SalesInvoiceDetails AS od ON rd.SalesDetailID = od.SalesDetailID)
       INNER JOIN Products AS p ON rd.ProductID = p.ProductID)
      INNER JOIN Units AS u ON p.UnitID = u.UnitID)
     INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID)
    INNER JOIN Employees AS e ON r.EmployeeID = e.EmployeeID"""),

    Query("qryPurchaseDocPrint", "بيانات طباعة فواتير الشراء ومرتجعاتها (سطر لكل صنف)", """
SELECT 'PURCHASE' AS DocKind, h.PurchaseInvoiceID AS DocID, h.InvoiceNumber AS DocNumber,
       h.InvoiceDate AS DocDate, h.SupplierInvoiceNo, '' AS OriginalNumber, h.PaymentType,
       '' AS Reason, h.SupplierID, s.SupplierName, s.VATNumber AS SupplierVAT,
       s.Mobile AS SupplierMobile, e.EmployeeName, h.SubTotal AS DocSubTotal,
       h.Discount AS DocDiscount, h.TaxableAmount, h.Tax AS DocTax, h.TotalAmount,
       h.PaidAmount, h.RemainingAmount, d.LineNumber, p.ProductCode, p.ProductName, u.UnitName,
       d.Quantity, d.UnitCost, d.Discount AS LineDiscount, d.NetAmount, d.VATRate,
       d.Tax AS LineTax, d.LineTotal
FROM ((((PurchaseInvoices AS h INNER JOIN PurchaseInvoiceDetails AS d
         ON h.PurchaseInvoiceID = d.PurchaseInvoiceID)
       INNER JOIN Products AS p ON d.ProductID = p.ProductID)
      INNER JOIN Units AS u ON p.UnitID = u.UnitID)
     INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID)
    INNER JOIN Employees AS e ON h.EmployeeID = e.EmployeeID
UNION ALL
SELECT 'RETURN', r.PurchaseReturnID, r.ReturnNumber, r.ReturnDate, o.SupplierInvoiceNo,
       o.InvoiceNumber, r.RefundType, r.Reason, r.SupplierID, s.SupplierName, s.VATNumber,
       s.Mobile, e.EmployeeName, r.SubTotal, r.Discount, r.TaxableAmount, r.Tax, r.TotalAmount,
       r.RefundedAmount, r.TotalAmount - r.RefundedAmount, od.LineNumber, p.ProductCode,
       p.ProductName, u.UnitName, rd.Quantity, rd.UnitCost, rd.Discount, rd.NetAmount,
       rd.VATRate, rd.Tax, rd.LineTotal
FROM ((((((PurchaseReturns AS r INNER JOIN PurchaseReturnDetails AS rd
           ON r.PurchaseReturnID = rd.PurchaseReturnID)
         INNER JOIN PurchaseInvoiceDetails AS od ON rd.PurchaseDetailID = od.PurchaseDetailID)
        INNER JOIN PurchaseInvoices AS o ON r.PurchaseInvoiceID = o.PurchaseInvoiceID)
       INNER JOIN Products AS p ON rd.ProductID = p.ProductID)
      INNER JOIN Units AS u ON p.UnitID = u.UnitID)
     INNER JOIN Suppliers AS s ON r.SupplierID = s.SupplierID)
    INNER JOIN Employees AS e ON r.EmployeeID = e.EmployeeID"""),

    Query("qryVoucherPrint", "بيانات طباعة سندات القبض (من العملاء) وسندات الصرف (للموردين)", """
SELECT 'RECEIPT' AS DocKind, p.PaymentID AS DocID, p.PaymentNumber AS DocNumber,
       p.PaymentDate AS DocDate, 1 AS LineNumber, c.CustomerName AS PartyName,
       c.Mobile AS PartyMobile, p.Amount, m.MethodName, p.Notes, e.EmployeeName,
       c.CurrentBalance AS PartyBalance
FROM ((CustomerPayments AS p INNER JOIN Customers AS c ON p.CustomerID = c.CustomerID)
      INNER JOIN PaymentMethods AS m ON p.PaymentMethodID = m.PaymentMethodID)
     INNER JOIN Employees AS e ON p.EmployeeID = e.EmployeeID
UNION ALL
SELECT 'PAYMENT', p.PaymentID, p.PaymentNumber, p.PaymentDate, 1, s.SupplierName, s.Mobile,
       p.Amount, m.MethodName, p.Notes, e.EmployeeName, s.CurrentBalance
FROM ((SupplierPayments AS p INNER JOIN Suppliers AS s ON p.SupplierID = s.SupplierID)
      INNER JOIN PaymentMethods AS m ON p.PaymentMethodID = m.PaymentMethodID)
     INNER JOIN Employees AS e ON p.EmployeeID = e.EmployeeID"""),

    # ============================================================= TREASURY
    Query("qryCashMovements", "كل حركات النقدية في الخزينة والصناديق: داخل (+) وخارج (−)", f"""
SELECT h.CashBoxID, h.InvoiceDate AS MoveDate, 'SALE' AS MoveType, 'فاتورة بيع' AS MoveTypeName,
       h.InvoiceNumber AS DocNumber, c.CustomerName AS PartyName, h.Notes AS Details,
       h.PaidAmount AS AmountIn, CCur(0) AS AmountOut, h.EmployeeID
FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID
WHERE h.CashBoxID Is Not Null AND h.PaidAmount <> 0
UNION ALL
SELECT r.CashBoxID, r.ReturnDate, 'SALES_RETURN', 'مرتجع بيع (رد نقدي)', r.ReturnNumber,
       c.CustomerName, r.Reason, CCur(0), r.RefundedAmount, r.EmployeeID
FROM SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID
WHERE r.CashBoxID Is Not Null AND r.RefundedAmount <> 0
UNION ALL
SELECT p.CashBoxID, p.PaymentDate, 'CUSTOMER_PAYMENT', 'سند قبض من عميل', p.PaymentNumber,
       c.CustomerName, p.Notes, p.Amount, CCur(0), p.EmployeeID
FROM CustomerPayments AS p INNER JOIN Customers AS c ON p.CustomerID = c.CustomerID
WHERE p.CashBoxID Is Not Null
UNION ALL
SELECT h.CashBoxID, h.InvoiceDate, 'PURCHASE', 'فاتورة شراء', h.InvoiceNumber,
       s.SupplierName, h.Notes, CCur(0), h.PaidAmount, h.EmployeeID
FROM PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID
WHERE h.CashBoxID Is Not Null AND h.PaidAmount <> 0
UNION ALL
SELECT r.CashBoxID, r.ReturnDate, 'PURCHASE_RETURN', 'مرتجع شراء (استرداد نقدي)', r.ReturnNumber,
       s.SupplierName, r.Reason, r.RefundedAmount, CCur(0), r.EmployeeID
FROM PurchaseReturns AS r INNER JOIN Suppliers AS s ON r.SupplierID = s.SupplierID
WHERE r.CashBoxID Is Not Null AND r.RefundedAmount <> 0
UNION ALL
SELECT p.CashBoxID, p.PaymentDate, 'SUPPLIER_PAYMENT', 'سند صرف لمورد', p.PaymentNumber,
       s.SupplierName, p.Notes, CCur(0), p.Amount, p.EmployeeID
FROM SupplierPayments AS p INNER JOIN Suppliers AS s ON p.SupplierID = s.SupplierID
WHERE p.CashBoxID Is Not Null
UNION ALL
SELECT e.CashBoxID, e.ExpenseDate, 'EXPENSE', 'مصروف', e.ExpenseNumber,
       t.ExpenseTypeName, e.Description, CCur(0), e.TotalAmount, e.EmployeeID
FROM Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID
WHERE e.CashBoxID Is Not Null
UNION ALL
SELECT v.CashBoxID, v.VoucherDate, 'CASH_IN',
       IIf(v.Category = 'OWNER', 'إيداع من المالك', IIf(v.Category = 'OVERAGE', 'زيادة في الصندوق',
           'سند قبض نقدية')),
       v.VoucherNumber, v.PartyName, v.Description, v.Amount, CCur(0), v.EmployeeID
FROM CashVouchers AS v
WHERE v.VoucherType = 'IN'
UNION ALL
SELECT v.CashBoxID, v.VoucherDate, 'CASH_OUT',
       IIf(v.Category = 'OWNER', 'تسوية مع المالك', IIf(v.Category = 'EXPENSE', 'مصروف (سند صرف)',
           IIf(v.Category = 'ADVANCE', 'سلفة موظف', IIf(v.Category = 'SHORTAGE', 'عجز في الصندوق',
           'سند صرف نقدية')))),
       v.VoucherNumber, v.PartyName, v.Description, CCur(0), v.Amount, v.EmployeeID
FROM CashVouchers AS v
WHERE v.VoucherType = 'OUT'
UNION ALL
SELECT v.CashBoxID, v.VoucherDate, 'TRANSFER_OUT', 'تحويل إلى صندوق آخر', v.VoucherNumber,
       b.BoxName, v.Description, CCur(0), v.Amount, v.EmployeeID
FROM CashVouchers AS v INNER JOIN CashBoxes AS b ON v.ToCashBoxID = b.CashBoxID
WHERE v.VoucherType = 'TRANSFER'
UNION ALL
SELECT v.ToCashBoxID, v.VoucherDate, 'TRANSFER_IN', 'تحويل من صندوق آخر', v.VoucherNumber,
       b.BoxName, v.Description, v.Amount, CCur(0), v.EmployeeID
FROM CashVouchers AS v INNER JOIN CashBoxes AS b ON v.CashBoxID = b.CashBoxID
WHERE v.VoucherType = 'TRANSFER'
UNION ALL
SELECT b.CashBoxID, b.OpeningDate, 'OPENING', 'رصيد افتتاحي', '-', b.BoxName, b.Notes,
       b.OpeningBalance, CCur(0), Null
FROM CashBoxes AS b
WHERE b.OpeningBalance <> 0"""),

    Query("qryCashBoxTotals", "إجمالي الداخل والخارج لكل صندوق", """
SELECT CashBoxID, Sum(AmountIn) AS BoxIn, Sum(AmountOut) AS BoxOut, Max(MoveDate) AS LastMoveDate
FROM qryCashMovements
GROUP BY CashBoxID"""),

    Query("CashBoxBalanceQuery", "أرصدة الخزينة والصناديق الآن", f"""
SELECT b.CashBoxID, b.BoxName, b.BoxType,
       IIf(b.BoxType = 'MAIN', 'خزينة رئيسية', 'صندوق كاشير') AS BoxTypeName, b.IsActive,
       {nz("t.BoxIn")} AS TotalIn, {nz("t.BoxOut")} AS TotalOut,
       {nz("t.BoxIn")} - {nz("t.BoxOut")} AS Balance, t.LastMoveDate
FROM CashBoxes AS b LEFT JOIN qryCashBoxTotals AS t ON b.CashBoxID = t.CashBoxID
ORDER BY b.BoxType DESC, b.BoxName"""),

    Query("CashStatementQuery", "حركة الخزينة / الصندوق لفترة: رصيد أول المدة ثم الحركات (0 = كل الصناديق)", f"""
SELECT 1 AS SortKey, m.MoveDate, m.MoveType, m.MoveTypeName, m.DocNumber, m.PartyName, m.Details,
       b.BoxName, m.AmountIn, m.AmountOut, m.CashBoxID
FROM qryCashMovements AS m INNER JOIN CashBoxes AS b ON m.CashBoxID = b.CashBoxID
WHERE (QLong('CashBoxID') = 0 OR m.CashBoxID = QLong('CashBoxID')) AND {period("m.MoveDate")}
UNION ALL
SELECT 0, QDate('PeriodStart'), 'BALANCE_FWD', 'رصيد أول المدة', '-', Null, Null, Null,
       IIf({nz("Sum(o.AmountIn)")} - {nz("Sum(o.AmountOut)")} > 0, {nz("Sum(o.AmountIn)")} - {nz("Sum(o.AmountOut)")}, 0),
       IIf({nz("Sum(o.AmountIn)")} - {nz("Sum(o.AmountOut)")} < 0, {nz("Sum(o.AmountOut)")} - {nz("Sum(o.AmountIn)")}, 0),
       QLong('CashBoxID')
FROM qryCashMovements AS o
WHERE (QLong('CashBoxID') = 0 OR o.CashBoxID = QLong('CashBoxID')) AND o.MoveDate < QDate('PeriodStart')
ORDER BY SortKey, MoveDate""", P + ["CashBoxID"]),

    Query("qryCashDays", "مقبوضات ومدفوعات كل يوم داخل الفترة", f"""
SELECT DateValue(m.MoveDate) AS CashDay, Sum(m.AmountIn) AS Receipts, Sum(m.AmountOut) AS Payments,
       Count(*) AS MoveCount
FROM qryCashMovements AS m
WHERE (QLong('CashBoxID') = 0 OR m.CashBoxID = QLong('CashBoxID')) AND {period("m.MoveDate")}
GROUP BY DateValue(m.MoveDate)""", P + ["CashBoxID"]),

    Query("qryCashDayOpening", "رصيد أول كل يوم من أيام الحركة (كل الحركات قبل ذلك اليوم)", """
SELECT d.CashDay, Sum(x.AmountIn - x.AmountOut) AS DayOpening
FROM qryCashDays AS d, qryCashMovements AS x
WHERE (QLong('CashBoxID') = 0 OR x.CashBoxID = QLong('CashBoxID')) AND x.MoveDate < d.CashDay
GROUP BY d.CashDay""", P + ["CashBoxID"]),

    # No subqueries here: a report that totals its columns wraps its record source in a
    # GROUP BY, and Access refuses subqueries in it (error 3612).
    Query("CashDailyQuery", "حركة الخزينة اليومية: رصيد أول اليوم والمقبوضات والمدفوعات ورصيد آخر اليوم", f"""
SELECT d.CashDay, {nz("o.DayOpening")} AS OpeningBalance, d.Receipts, d.Payments,
       {nz("o.DayOpening")} + d.Receipts - d.Payments AS ClosingBalance, d.MoveCount
FROM qryCashDays AS d LEFT JOIN qryCashDayOpening AS o ON d.CashDay = o.CashDay
ORDER BY d.CashDay""", P + ["CashBoxID"]),

    Query("CashClosingsQuery", "تصفيات يومية الكاشير خلال فترة (0 = كل الصناديق)", f"""
SELECT c.ClosingID, c.ClosingNumber, c.ClosingDate, b.BoxName, e.EmployeeName, c.PeriodStart,
       c.OpeningBalance, c.CashIn, c.CashOut, c.ExpectedBalance, c.CountedAmount, c.Difference,
       IIf(c.Destination = 'MAIN', 'الخزينة الرئيسية', IIf(c.Destination = 'OWNER', 'تسوية مع المالك',
           'يبقى في الصندوق')) AS DestinationName,
       t.BoxName AS ToBoxName, c.TransferAmount, c.KeptAmount, c.Notes, c.CashBoxID
FROM ((CashClosings AS c INNER JOIN CashBoxes AS b ON c.CashBoxID = b.CashBoxID)
      INNER JOIN Employees AS e ON c.EmployeeID = e.EmployeeID)
     LEFT JOIN CashBoxes AS t ON c.ToCashBoxID = t.CashBoxID
WHERE (QLong('CashBoxID') = 0 OR c.CashBoxID = QLong('CashBoxID')) AND {period("c.ClosingDate")}
ORDER BY c.ClosingDate""", P + ["CashBoxID"]),

    Query("qryCashClosingPrint", "بيانات طباعة تصفية الكاشير", """
SELECT c.ClosingID, c.ClosingNumber, c.ClosingDate, b.BoxName, e.EmployeeName, c.PeriodStart,
       c.OpeningBalance, c.CashIn, c.CashOut, c.ExpectedBalance, c.CountedAmount, c.Difference,
       IIf(c.Destination = 'MAIN', 'الخزينة الرئيسية', IIf(c.Destination = 'OWNER', 'تسوية مع المالك',
           'يبقى في الصندوق')) AS DestinationName,
       t.BoxName AS ToBoxName, c.TransferAmount, c.KeptAmount, c.Notes
FROM ((CashClosings AS c INNER JOIN CashBoxes AS b ON c.CashBoxID = b.CashBoxID)
      INNER JOIN Employees AS e ON c.EmployeeID = e.EmployeeID)
     LEFT JOIN CashBoxes AS t ON c.ToCashBoxID = t.CashBoxID"""),

    Query("qryCashVoucherPrint", "بيانات طباعة سندات قبض وصرف وتحويل النقدية", """
SELECT v.CashVoucherID AS DocID, v.VoucherNumber, v.VoucherDate, v.VoucherType,
       IIf(v.VoucherType = 'IN', 'سند قبض نقدية', IIf(v.VoucherType = 'OUT', 'سند صرف نقدية',
           'سند تحويل نقدية')) AS VoucherTitle,
       IIf(v.Category = 'OWNER', IIf(v.VoucherType = 'IN', 'إيداع من المالك', 'تسوية مع المالك'),
           IIf(v.Category = 'EXPENSE', 'مصروف', IIf(v.Category = 'ADVANCE', 'سلفة موظف',
           IIf(v.Category = 'SHORTAGE', 'عجز في الصندوق', IIf(v.Category = 'OVERAGE', 'زيادة في الصندوق',
           IIf(v.Category = 'TRANSFER', 'تحويل بين الصناديق', 'أخرى')))))) AS CategoryName,
       b.BoxName, t.BoxName AS ToBoxName, v.Amount, v.PartyName, v.Description,
       x.ExpenseTypeName, e.EmployeeName
FROM ((((CashVouchers AS v INNER JOIN CashBoxes AS b ON v.CashBoxID = b.CashBoxID)
        INNER JOIN Employees AS e ON v.EmployeeID = e.EmployeeID)
       LEFT JOIN CashBoxes AS t ON v.ToCashBoxID = t.CashBoxID)
      LEFT JOIN Expenses AS ex ON v.ExpenseID = ex.ExpenseID)
     LEFT JOIN ExpenseTypes AS x ON ex.ExpenseTypeID = x.ExpenseTypeID"""),

    # ============================================================ JOURNAL
    Query("qrySaleCost", "تكلفة كل فاتورة بيع", """
SELECT SalesInvoiceID, Sum(Quantity * UnitCost) AS SaleCost
FROM SalesInvoiceDetails
GROUP BY SalesInvoiceID"""),

    Query("qryReturnCost", "تكلفة ما عاد للمخزون من كل مرتجع بيع", """
SELECT SalesReturnID, Sum(IIf(ReturnToStock, Quantity * UnitCost, 0)) AS ReturnCost
FROM SalesReturnDetails
GROUP BY SalesReturnID"""),

    Query("qryStockCountValue", "قيمة فروقات كل جرد مُرحّل", """
SELECT ReferenceID AS StockCountID, Max(ReferenceNumber) AS CountNumber, Max(TransactionDate) AS CountDate,
       Sum(Quantity * UnitCost) AS CountValue
FROM InventoryTransactions
WHERE ReferenceType = 'STOCK_COUNT'
GROUP BY ReferenceID"""),

    Query("qryJournalSale", "أسطر قيود فواتير البيع", _sale()),
    Query("qryJournalSalesReturn", "أسطر قيود مرتجعات البيع", _sales_return()),
    Query("qryJournalPurchase", "أسطر قيود فواتير الشراء", _purchase()),
    Query("qryJournalPurchaseReturn", "أسطر قيود مرتجعات الشراء", _purchase_return()),
    Query("qryJournalPayments", "أسطر قيود سندات القبض من العملاء والصرف للموردين", _payments()),
    Query("qryJournalExpense", "أسطر قيود المصروفات (عدا المسجلة بسند نقدية)", _expense()),
    Query("qryJournalCashVoucher", "أسطر قيود سندات النقدية (قبض وصرف وتحويل)", _cash_voucher()),
    Query("qryJournalStock", "أسطر قيود حركات المخزون اليدوية وتسويات الجرد", _stock()),
    Query("qryJournalOpening", "أسطر قيود الأرصدة الافتتاحية للصناديق والعملاء والموردين", _opening()),
    Query("qryManualEntryLines", "أسطر القيود اليدوية مع رأس كل قيد", """
SELECT h.ManualEntryID, h.EntryNumber, h.EntryDate, h.Description, l.LineNumber AS LineNo,
       l.AccountCode AS LineAccount, l.Debit AS LineDebit, l.Credit AS LineCredit, l.LineText AS LineNote
FROM ManualEntries AS h INNER JOIN ManualEntryLines AS l ON h.ManualEntryID = l.ManualEntryID"""),
    Query("qryJournalManual", "أسطر القيود اليدوية", _manual()),

    Query("JournalLinesQuery", "قيود اليومية خلال فترة بأسطرها", f"""
SELECT e.EntryID, e.EntryNumber, e.EntryDate, e.SourceType, t.TypeName, e.SourceID, e.SourceNumber,
       e.Description, l.LineNumber, l.AccountCode, a.AccountName, l.LineText, l.Debit, l.Credit
FROM ((JournalEntries AS e INNER JOIN JournalSourceTypes AS t ON e.SourceType = t.SourceType)
      INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID)
     INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode
WHERE {period("e.EntryDate")}
ORDER BY e.EntryDate, e.EntryNumber, l.LineNumber""", P),

    Query("qryJournalEntryPrint", "بيانات طباعة قيد", """
SELECT e.EntryID, e.EntryNumber, e.EntryDate, t.TypeName, e.SourceNumber, e.Description, e.TotalDebit,
       l.LineNumber, l.AccountCode, a.AccountName, l.LineText, l.Debit, l.Credit
FROM ((JournalEntries AS e INNER JOIN JournalSourceTypes AS t ON e.SourceType = t.SourceType)
      INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID)
     INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode"""),

    Query("qryTrialBefore", "مجموع الحسابات قبل الفترة", """
SELECT l.AccountCode, Sum(l.Debit) AS DebitBefore, Sum(l.Credit) AS CreditBefore
FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID
WHERE e.EntryDate < QDate('PeriodStart')
GROUP BY l.AccountCode""", ["PeriodStart"]),

    Query("qryTrialPeriod", "حركة الحسابات خلال الفترة", f"""
SELECT l.AccountCode, Sum(l.Debit) AS SumDebit, Sum(l.Credit) AS SumCredit
FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID
WHERE {period("e.EntryDate")}
GROUP BY l.AccountCode""", P),

    Query("TrialBalanceQuery", "ميزان المراجعة: رصيد أول المدة وحركة الفترة والرصيد الختامي (المدين موجب)", f"""
SELECT a.AccountCode, a.AccountName, a.AccountType,
       {nz("b.DebitBefore")} - {nz("b.CreditBefore")} AS OpeningBalance,
       {nz("p.SumDebit")} AS PeriodDebit, {nz("p.SumCredit")} AS PeriodCredit,
       {nz("b.DebitBefore")} - {nz("b.CreditBefore")} + {nz("p.SumDebit")} - {nz("p.SumCredit")} AS ClosingBalance
FROM (Accounts AS a LEFT JOIN qryTrialBefore AS b ON a.AccountCode = b.AccountCode)
     LEFT JOIN qryTrialPeriod AS p ON a.AccountCode = p.AccountCode
WHERE b.AccountCode Is Not Null OR p.AccountCode Is Not Null
ORDER BY a.AccountCode""", P),

    # A main account takes all its sub-accounts: QLong('AccountCode') is one of the LevelNCode of the line.
    Query("qryStatementBefore", "رصيد الحساب المختار (مع حساباته التابعة) قبل بداية الفترة", f"""
SELECT {nz("Sum(o.Debit)")} - {nz("Sum(o.Credit)")} AS SumBefore
FROM (JournalLines AS o INNER JOIN JournalEntries AS f ON o.EntryID = f.EntryID)
     INNER JOIN Accounts AS b ON o.AccountCode = b.AccountCode
WHERE {IN_TREE("b")} AND f.EntryDate < QDate('PeriodStart')""", ["PeriodStart", "AccountCode"]),

    Query("AccountStatementQuery", "كشف حساب لفترة: رصيد أول المدة ثم كل سطر قيد (الحساب الرئيسي يشمل حساباته التابعة)", f"""
SELECT 1 AS SortKey, s.AccountCode AS StatementAccount, s.AccountName AS StatementName, e.EntryDate AS LineDate,
       e.EntryNumber AS EntryNo, e.EntryID AS EntryRef, k.TypeName AS KindName, e.SourceNumber AS DocNo,
       e.Description AS Details, a.AccountCode AS SubCode, a.AccountName AS SubName, l.Debit AS LineDebit,
       l.Credit AS LineCredit
FROM (((JournalLines AS l INNER JOIN JournalEntries AS e ON l.EntryID = e.EntryID)
      INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode)
     INNER JOIN JournalSourceTypes AS k ON e.SourceType = k.SourceType), Accounts AS s
WHERE s.AccountCode = QLong('AccountCode') AND {IN_TREE("a")} AND {period("e.EntryDate")}
UNION ALL
SELECT 0, s.AccountCode, s.AccountName, QDate('PeriodStart'), '-', 0, 'رصيد أول المدة', Null, Null, Null, Null,
       IIf(x.SumBefore > 0, x.SumBefore, 0), IIf(x.SumBefore < 0, -x.SumBefore, 0)
FROM Accounts AS s, qryStatementBefore AS x
WHERE s.AccountCode = QLong('AccountCode')
ORDER BY SortKey, LineDate, EntryNo""", P + ["AccountCode"]),

    Query("GeneralLedgerQuery", "دفتر الأستاذ لفترة: لكل حساب فرعي رصيد أول المدة ثم أسطر قيوده (0 = كل الحسابات)", f"""
SELECT 1 AS SortKey, a.TreeKey AS AccountKey, a.AccountCode AS LedgerCode, a.AccountName AS LedgerName,
       e.EntryDate AS LineDate, e.EntryNumber AS EntryNo, e.EntryID AS EntryRef, k.TypeName AS KindName,
       e.SourceNumber AS DocNo, e.Description AS Details, l.Debit AS LineDebit, l.Credit AS LineCredit
FROM ((JournalLines AS l INNER JOIN JournalEntries AS e ON l.EntryID = e.EntryID)
      INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode)
     INNER JOIN JournalSourceTypes AS k ON e.SourceType = k.SourceType
WHERE (QLong('AccountCode') = 0 OR {IN_TREE("a")}) AND {period("e.EntryDate")}
UNION ALL
SELECT 0, b.TreeKey, b.AccountCode, b.AccountName, QDate('PeriodStart'), '-', 0, 'رصيد أول المدة', Null, Null,
       IIf(t.OpeningBalance > 0, t.OpeningBalance, 0), IIf(t.OpeningBalance < 0, -t.OpeningBalance, 0)
FROM TrialBalanceQuery AS t INNER JOIN Accounts AS b ON t.AccountCode = b.AccountCode
WHERE QLong('AccountCode') = 0 OR {IN_TREE("b")}
ORDER BY AccountKey, SortKey, LineDate, EntryNo""", P + ["AccountCode"]),

    Query("qryTreeRollup", "أرصدة ميزان المراجعة مجمّعة على كل مستوى من شجرة الحسابات", _tree_rollup(), P),

    Query("TrialBalanceTreeQuery", "ميزان المراجعة بالمستويات: كل حساب رئيسي بمجموع حساباته التابعة", f"""
SELECT a.AccountCode, a.AccountName, {ACCOUNT_TYPE_NAME} AS TypeName, a.AccountLevel, a.TreeKey, a.IsPosting,
       r.SumOpening AS OpeningBalance, r.SumDebit AS PeriodDebit, r.SumCredit AS PeriodCredit,
       r.SumClosing AS ClosingBalance
FROM Accounts AS a INNER JOIN qryTreeRollup AS r ON a.AccountCode = r.TreeCode
ORDER BY a.TreeKey""", P),

    Query("AccountTreeQuery", "شجرة الحسابات: كل حساب بمستواه ونوعه وهل يقبل القيود", f"""
SELECT a.AccountCode, a.AccountName, {ACCOUNT_TYPE_NAME} AS TypeName, a.AccountLevel, a.TreeKey,
       a.ParentCode, IIf(a.IsPosting, 'فرعي', 'رئيسي') AS KindName, a.IsPosting, a.IsActive
FROM Accounts AS a
ORDER BY a.TreeKey"""),

    # ============================================================ INTEGRITY
    Query("qrySalesInvoiceLineTotals", "مجموع أسطر كل فاتورة بيع", """
SELECT SalesInvoiceID, Sum(LineTotal) AS LinesTotal
FROM SalesInvoiceDetails
GROUP BY SalesInvoiceID"""),

    Query("qryPurchaseInvoiceLineTotals", "مجموع أسطر كل فاتورة شراء", """
SELECT PurchaseInvoiceID, Sum(LineTotal) AS LinesTotal
FROM PurchaseInvoiceDetails
GROUP BY PurchaseInvoiceID"""),

    Query("qrySalesReturnedQty", "الكمية المرتجعة من كل سطر فاتورة بيع", """
SELECT SalesDetailID, Sum(Quantity) AS QtyReturned
FROM SalesReturnDetails
GROUP BY SalesDetailID"""),

    Query("qryPurchaseReturnedQty", "الكمية المرتجعة للمورد من كل سطر فاتورة شراء", """
SELECT PurchaseDetailID, Sum(Quantity) AS QtyReturned
FROM PurchaseReturnDetails
GROUP BY PurchaseDetailID"""),

    Query("IntegrityCheckQuery", "فحص سلامة البيانات: أي سطر هنا مشكلة يجب مراجعتها (النتيجة الفارغة = سليم)", f"""
SELECT 'STOCK_MISMATCH' AS IssueCode, 'الكمية المسجلة لا تطابق حركات المخزون' AS IssueText,
       'Products' AS SourceTable, p.ProductID AS RecordID,
       {nz("l.LedgerQty")} AS ExpectedValue, CCur(p.CurrentQuantity) AS ActualValue
FROM Products AS p LEFT JOIN qryProductLedger AS l ON p.ProductID = l.ProductID
WHERE p.CurrentQuantity <> {nz("l.LedgerQty")}
UNION ALL
SELECT 'NEGATIVE_STOCK', 'رصيد المنتج سالب', 'Products', p.ProductID, CCur(0),
       CCur(p.CurrentQuantity)
FROM Products AS p
WHERE p.CurrentQuantity < 0
UNION ALL
SELECT 'CUSTOMER_BALANCE', 'رصيد العميل المسجل لا يطابق الحركات', 'Customers', b.CustomerID,
       CCur(b.Balance), CCur(b.CachedBalance)
FROM CustomerBalanceQuery AS b
WHERE b.Balance <> b.CachedBalance
UNION ALL
SELECT 'SUPPLIER_BALANCE', 'رصيد المورد المسجل لا يطابق الحركات', 'Suppliers', b.SupplierID,
       CCur(b.Balance), CCur(b.CachedBalance)
FROM SupplierBalanceQuery AS b
WHERE b.Balance <> b.CachedBalance
UNION ALL
SELECT 'SALE_TOTAL', 'إجمالي فاتورة البيع لا يساوي مجموع أسطرها', 'SalesInvoices',
       h.SalesInvoiceID, {nz("t.LinesTotal")}, CCur(h.TotalAmount)
FROM SalesInvoices AS h LEFT JOIN qrySalesInvoiceLineTotals AS t
     ON h.SalesInvoiceID = t.SalesInvoiceID
WHERE h.TotalAmount <> {nz("t.LinesTotal")} OR t.SalesInvoiceID Is Null
UNION ALL
SELECT 'PURCHASE_TOTAL', 'إجمالي فاتورة الشراء لا يساوي مجموع أسطرها', 'PurchaseInvoices',
       h.PurchaseInvoiceID, {nz("t.LinesTotal")}, CCur(h.TotalAmount)
FROM PurchaseInvoices AS h LEFT JOIN qryPurchaseInvoiceLineTotals AS t
     ON h.PurchaseInvoiceID = t.PurchaseInvoiceID
WHERE h.TotalAmount <> {nz("t.LinesTotal")} OR t.PurchaseInvoiceID Is Null
UNION ALL
SELECT 'CREDIT_NOT_ALLOWED', 'بيع آجل لعميل غير مسموح له بالآجل', 'SalesInvoices',
       h.SalesInvoiceID, CCur(0), CCur(h.RemainingAmount)
FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID
WHERE h.RemainingAmount > 0 AND c.AllowCredit = False
UNION ALL
SELECT 'RETURN_CUSTOMER', 'عميل المرتجع يختلف عن عميل الفاتورة الأصلية', 'SalesReturns',
       r.SalesReturnID, CCur(h.CustomerID), CCur(r.CustomerID)
FROM SalesReturns AS r INNER JOIN SalesInvoices AS h ON r.SalesInvoiceID = h.SalesInvoiceID
WHERE r.CustomerID <> h.CustomerID
UNION ALL
SELECT 'RETURN_LINE_INVOICE', 'سطر المرتجع من فاتورة غير فاتورة المرتجع', 'SalesReturnDetails',
       d.ReturnDetailID, CCur(r.SalesInvoiceID), CCur(o.SalesInvoiceID)
FROM (SalesReturnDetails AS d INNER JOIN SalesReturns AS r ON d.SalesReturnID = r.SalesReturnID)
     INNER JOIN SalesInvoiceDetails AS o ON d.SalesDetailID = o.SalesDetailID
WHERE o.SalesInvoiceID <> r.SalesInvoiceID
UNION ALL
SELECT 'RETURN_LINE_PRODUCT', 'منتج سطر المرتجع يختلف عن منتج السطر الأصلي', 'SalesReturnDetails',
       d.ReturnDetailID, CCur(o.ProductID), CCur(d.ProductID)
FROM SalesReturnDetails AS d INNER JOIN SalesInvoiceDetails AS o
     ON d.SalesDetailID = o.SalesDetailID
WHERE d.ProductID <> o.ProductID
UNION ALL
SELECT 'SALE_RETURN_QTY', 'الكمية المرتجعة أكبر من الكمية المباعة', 'SalesInvoiceDetails',
       o.SalesDetailID, CCur(o.Quantity), CCur(q.QtyReturned)
FROM SalesInvoiceDetails AS o INNER JOIN qrySalesReturnedQty AS q
     ON o.SalesDetailID = q.SalesDetailID
WHERE q.QtyReturned > o.Quantity
UNION ALL
SELECT 'PURCHASE_RETURN_QTY', 'الكمية المرتجعة للمورد أكبر من الكمية المشتراة', 'PurchaseInvoiceDetails',
       o.PurchaseDetailID, CCur(o.Quantity), CCur(q.QtyReturned)
FROM PurchaseInvoiceDetails AS o INNER JOIN qryPurchaseReturnedQty AS q
     ON o.PurchaseDetailID = q.PurchaseDetailID
WHERE q.QtyReturned > o.Quantity
ORDER BY IssueCode, RecordID"""),
]


# --------------------------------------------------------------------------
# Test fixture: a small, fully posted business history.
# Values: Ref("KEY") = id of an earlier fixture row, Day(n, h) = n days ago at h:00.
# --------------------------------------------------------------------------
@dataclass(frozen=True)
class Ref:
    key: str


@dataclass(frozen=True)
class Day:
    days_ago: int
    hour: int = 0


@dataclass
class Row:
    key: str
    table: str
    values: Dict[str, object]


def _inv_tx(key, day, product, type_id, qty, cost, after, ref_type, ref_key, ref_no):
    return Row(key, "InventoryTransactions", {
        "TransactionDate": day, "ProductID": Ref(product), "TransactionTypeID": type_id,
        "Quantity": qty, "UnitCost": cost, "QuantityAfter": after, "ReferenceType": ref_type,
        "ReferenceID": Ref(ref_key) if ref_key else None, "ReferenceNumber": ref_no,
        "EmployeeID": 1})


FIXTURE: List[Row] = [
    # Treasury: main safe 10000 and a cashier box 500 (opening balances 60 days ago)
    Row("BOXM", "CashBoxes", {"BoxName": "TEST الخزينة", "BoxType": "MAIN", "OpeningBalance": 10000,
                              "OpeningDate": Day(60)}),
    Row("BOXC", "CashBoxes", {"BoxName": "TEST صندوق كاشير", "BoxType": "CASHIER", "OpeningBalance": 500,
                              "OpeningDate": Day(60)}),
    Row("S1", "Suppliers", {"SupplierName": "TEST مورد", "OpeningBalance": 0,
                            "CurrentBalance": 2855, "CreatedAt": Day(120)}),
    Row("C2", "Customers", {"CustomerName": "TEST عميل آجل", "OpeningBalance": 50,
                            "CurrentBalance": 164, "AllowCredit": True, "CreatedAt": Day(90)}),
    Row("CAT2", "Categories", {"CategoryName": "TEST مواد غذائية"}),
    Row("P1", "Products", {"ProductCode": "TEST-P1", "ProductName": "TEST منتج 1", "CategoryID": 1,
                           "UnitID": 1, "PurchasePrice": 60, "AverageCost": 60, "SellingPrice": 115,
                           "CurrentQuantity": 83, "MinimumQuantity": 5, "SupplierID": Ref("S1"),
                           "CreatedAt": Day(120)}),
    Row("P2", "Products", {"ProductCode": "TEST-P2", "ProductName": "TEST منتج 2",
                           "CategoryID": Ref("CAT2"), "UnitID": 1, "PurchasePrice": 10,
                           "AverageCost": 10, "SellingPrice": 23, "CurrentQuantity": 186,
                           "MinimumQuantity": 200, "SupplierID": Ref("S1"), "CreatedAt": Day(120)}),
    Row("P3", "Products", {"ProductCode": "TEST-P3", "ProductName": "TEST منتج 3", "CategoryID": 1,
                           "UnitID": 1, "PurchasePrice": 10, "AverageCost": 10, "SellingPrice": 20,
                           "CurrentQuantity": 20, "MinimumQuantity": 0, "CreatedAt": Day(120)}),
    _inv_tx("T1", Day(60, 8), "P3", 8, 20, 10, 20, "MANUAL", None, "TEST-OPEN"),

    # Purchase 100 x P1 @60 and 200 x P2 @10, VAT 15%, paid 5000 of 9200
    Row("PUR1", "PurchaseInvoices", {
        "InvoiceNumber": "TEST-PUR-1", "SupplierInvoiceNo": "S-100", "InvoiceDate": Day(50, 9),
        "SupplierID": Ref("S1"), "EmployeeID": 1, "PaymentType": "CREDIT", "PaymentMethodID": 1,
        "SubTotal": 8000, "Discount": 0, "TaxableAmount": 8000, "Tax": 1200, "TotalAmount": 9200,
        "PaidAmount": 5000, "RemainingAmount": 4200, "CashBoxID": Ref("BOXM")}),
    Row("PUR1L1", "PurchaseInvoiceDetails", {
        "PurchaseInvoiceID": Ref("PUR1"), "LineNumber": 1, "ProductID": Ref("P1"), "Quantity": 100,
        "UnitCost": 60, "Discount": 0, "NetAmount": 6000, "VATRate": 0.15, "Tax": 900,
        "LineTotal": 6900}),
    Row("PUR1L2", "PurchaseInvoiceDetails", {
        "PurchaseInvoiceID": Ref("PUR1"), "LineNumber": 2, "ProductID": Ref("P2"), "Quantity": 200,
        "UnitCost": 10, "Discount": 0, "NetAmount": 2000, "VATRate": 0.15, "Tax": 300,
        "LineTotal": 2300}),
    _inv_tx("T2", Day(50, 9), "P1", 1, 100, 60, 100, "PURCHASE", "PUR1", "TEST-PUR-1"),
    _inv_tx("T3", Day(50, 9), "P2", 1, 200, 10, 200, "PURCHASE", "PUR1", "TEST-PUR-1"),

    # Old cash sale (outside the 30-day test period): 5 x P2 @23 incl. VAT
    Row("INV0", "SalesInvoices", {
        "InvoiceNumber": "TEST-INV-0", "InvoiceDate": Day(45, 10), "CustomerID": 1,
        "EmployeeID": 1, "PaymentType": "CASH", "PaymentMethodID": 1, "SubTotal": 100,
        "Discount": 0, "TaxableAmount": 100, "Tax": 15, "TotalAmount": 115, "PaidAmount": 115,
        "RemainingAmount": 0, "AmountTendered": 115, "ChangeDue": 0, "CashBoxID": Ref("BOXC")}),
    Row("INV0L1", "SalesInvoiceDetails", {
        "SalesInvoiceID": Ref("INV0"), "LineNumber": 1, "ProductID": Ref("P2"), "Quantity": 5,
        "UnitPrice": 20, "Discount": 0, "NetAmount": 100, "VATRate": 0.15, "Tax": 15,
        "LineTotal": 115, "UnitCost": 10}),
    _inv_tx("T4", Day(45, 10), "P2", 2, -5, 10, 195, "SALE", "INV0", "TEST-INV-0"),

    # Cash sale: 10 x P1 @115 incl. VAT -> net 1000, VAT 150
    Row("INV1", "SalesInvoices", {
        "InvoiceNumber": "TEST-INV-1", "InvoiceDate": Day(10, 11), "CustomerID": 1,
        "EmployeeID": 1, "PaymentType": "CASH", "PaymentMethodID": 1, "SubTotal": 1000,
        "Discount": 0, "TaxableAmount": 1000, "Tax": 150, "TotalAmount": 1150,
        "PaidAmount": 1150, "RemainingAmount": 0, "AmountTendered": 1200, "ChangeDue": 50,
        "CashBoxID": Ref("BOXC")}),
    Row("INV1L1", "SalesInvoiceDetails", {
        "SalesInvoiceID": Ref("INV1"), "LineNumber": 1, "ProductID": Ref("P1"), "Quantity": 10,
        "UnitPrice": 100, "Discount": 0, "NetAmount": 1000, "VATRate": 0.15, "Tax": 150,
        "LineTotal": 1150, "UnitCost": 60}),
    _inv_tx("T5", Day(10, 11), "P1", 2, -10, 60, 90, "SALE", "INV1", "TEST-INV-1"),

    # Credit sale to C2: 2 x P1 + 10 x P2 = 460, paid 100
    Row("INV2", "SalesInvoices", {
        "InvoiceNumber": "TEST-INV-2", "InvoiceDate": Day(5, 12), "CustomerID": Ref("C2"),
        "EmployeeID": 1, "PaymentType": "CREDIT", "PaymentMethodID": 1, "SubTotal": 400,
        "Discount": 0, "TaxableAmount": 400, "Tax": 60, "TotalAmount": 460, "PaidAmount": 100,
        "RemainingAmount": 360, "AmountTendered": 100, "ChangeDue": 0, "CashBoxID": Ref("BOXC")}),
    Row("INV2L1", "SalesInvoiceDetails", {
        "SalesInvoiceID": Ref("INV2"), "LineNumber": 1, "ProductID": Ref("P1"), "Quantity": 2,
        "UnitPrice": 100, "Discount": 0, "NetAmount": 200, "VATRate": 0.15, "Tax": 30,
        "LineTotal": 230, "UnitCost": 60}),
    Row("INV2L2", "SalesInvoiceDetails", {
        "SalesInvoiceID": Ref("INV2"), "LineNumber": 2, "ProductID": Ref("P2"), "Quantity": 10,
        "UnitPrice": 20, "Discount": 0, "NetAmount": 200, "VATRate": 0.15, "Tax": 30,
        "LineTotal": 230, "UnitCost": 10}),
    _inv_tx("T6", Day(5, 12), "P1", 2, -2, 60, 88, "SALE", "INV2", "TEST-INV-2"),
    _inv_tx("T7", Day(5, 12), "P2", 2, -10, 10, 185, "SALE", "INV2", "TEST-INV-2"),

    # Payments
    Row("PAY1", "SupplierPayments", {
        "PaymentNumber": "TEST-PAY-1", "SupplierID": Ref("S1"), "PaymentDate": Day(4, 10),
        "Amount": 1000, "PaymentMethodID": 3, "EmployeeID": 1}),
    Row("RCV1", "CustomerPayments", {
        "PaymentNumber": "TEST-RCV-1", "CustomerID": Ref("C2"), "PaymentDate": Day(3, 10),
        "Amount": 200, "PaymentMethodID": 1, "EmployeeID": 1, "CashBoxID": Ref("BOXC")}),

    # C2 returns 2 x P2 (back to stock), credited to the account
    Row("CRN1", "SalesReturns", {
        "ReturnNumber": "TEST-CRN-1", "ReturnDate": Day(2, 13), "SalesInvoiceID": Ref("INV2"),
        "CustomerID": Ref("C2"), "EmployeeID": 1, "Reason": "TEST إرجاع العميل",
        "RefundType": "CREDIT", "SubTotal": 40, "Discount": 0, "TaxableAmount": 40, "Tax": 6,
        "TotalAmount": 46, "RefundedAmount": 0}),
    Row("CRN1L1", "SalesReturnDetails", {
        "SalesReturnID": Ref("CRN1"), "SalesDetailID": Ref("INV2L2"), "ProductID": Ref("P2"),
        "Quantity": 2, "UnitPrice": 20, "Discount": 0, "NetAmount": 40, "VATRate": 0.15,
        "Tax": 6, "LineTotal": 46, "UnitCost": 10, "ReturnToStock": True}),
    _inv_tx("T8", Day(2, 13), "P2", 4, 2, 10, 187, "SALES_RETURN", "CRN1", "TEST-CRN-1"),

    # 5 x P1 returned to the supplier, deducted from the supplier balance
    Row("PRT1", "PurchaseReturns", {
        "ReturnNumber": "TEST-PRT-1", "ReturnDate": Day(1, 10), "PurchaseInvoiceID": Ref("PUR1"),
        "SupplierID": Ref("S1"), "EmployeeID": 1, "Reason": "TEST عيب مصنعي",
        "RefundType": "CREDIT", "SubTotal": 300, "Discount": 0, "TaxableAmount": 300, "Tax": 45,
        "TotalAmount": 345, "RefundedAmount": 0}),
    Row("PRT1L1", "PurchaseReturnDetails", {
        "PurchaseReturnID": Ref("PRT1"), "PurchaseDetailID": Ref("PUR1L1"), "ProductID": Ref("P1"),
        "Quantity": 5, "UnitCost": 60, "Discount": 0, "NetAmount": 300, "VATRate": 0.15,
        "Tax": 45, "LineTotal": 345}),
    _inv_tx("T9", Day(1, 10), "P1", 3, -5, 60, 83, "PURCHASE_RETURN", "PRT1", "TEST-PRT-1"),

    # 1 x P2 damaged (manual stock-out)
    _inv_tx("T10", Day(1, 15), "P2", 6, -1, 10, 186, "MANUAL", None, "TEST-ADJ-1"),

    # An open stocktake (not posted, so no stock movement): P1 counted 80, P2 not counted yet
    Row("CNT1", "StockCounts", {"CountNumber": "TEST-CNT-1", "CountDate": Day(1, 18),
                                "Status": "OPEN", "EmployeeID": 1}),
    Row("CNT1L1", "StockCountDetails", {
        "StockCountID": Ref("CNT1"), "ProductID": Ref("P1"), "SystemQuantity": 83,
        "ActualQuantity": 80, "Difference": -3, "UnitCost": 60, "DifferenceValue": -180}),
    Row("CNT1L2", "StockCountDetails", {
        "StockCountID": Ref("CNT1"), "ProductID": Ref("P2"), "SystemQuantity": 186,
        "ActualQuantity": None, "Difference": 0, "UnitCost": 10, "DifferenceValue": 0}),

    # Expenses: electricity inside the period, rent outside it
    Row("EXP1", "Expenses", {
        "ExpenseNumber": "TEST-EXP-1", "ExpenseDate": Day(6), "ExpenseTypeID": 2, "Amount": 200,
        "Tax": 30, "TotalAmount": 230, "PaymentMethodID": 1, "Description": "TEST كهرباء",
        "EmployeeID": 1, "CashBoxID": Ref("BOXC")}),
    Row("EXP2", "Expenses", {
        "ExpenseNumber": "TEST-EXP-2", "ExpenseDate": Day(40), "ExpenseTypeID": 1, "Amount": 1000,
        "Tax": 0, "TotalAmount": 1000, "PaymentMethodID": 3, "Description": "TEST إيجار",
        "EmployeeID": 1}),

    # Cash voucher for an expense (35 days ago): 50 out of the cashier box, recorded as an expense
    Row("EXPV", "Expenses", {
        "ExpenseNumber": "TEST-EXP-V", "ExpenseDate": Day(35), "ExpenseTypeID": 9, "Amount": 50,
        "Tax": 0, "TotalAmount": 50, "PaymentMethodID": 1, "Description": "TEST نثريات", "EmployeeID": 1}),
    Row("V1", "CashVouchers", {
        "VoucherNumber": "TEST-COT-1", "VoucherDate": Day(35, 12), "VoucherType": "OUT",
        "CashBoxID": Ref("BOXC"), "Category": "EXPENSE", "Amount": 50, "PartyName": "TEST محل",
        "ExpenseID": Ref("EXPV"), "EmployeeID": 1}),
    # Cashier closing 4 days ago: book 500 + 115 + 1150 + 100 - 230 - 50 = 1585, counted 1570
    # (shortage 15), 1000 moved to the main safe, 570 kept as float
    Row("CL1", "CashClosings", {
        "ClosingNumber": "TEST-CLS-1", "ClosingDate": Day(4, 18), "CashBoxID": Ref("BOXC"),
        "EmployeeID": 1, "OpeningBalance": 0, "CashIn": 1865, "CashOut": 280,
        "ExpectedBalance": 1585, "CountedAmount": 1570, "Difference": -15, "Destination": "MAIN",
        "ToCashBoxID": Ref("BOXM"), "TransferAmount": 1000, "KeptAmount": 570}),
    Row("V2", "CashVouchers", {
        "VoucherNumber": "TEST-COT-2", "VoucherDate": Day(4, 18), "VoucherType": "OUT",
        "CashBoxID": Ref("BOXC"), "Category": "SHORTAGE", "Amount": 15, "ClosingID": Ref("CL1"),
        "EmployeeID": 1}),
    Row("V3", "CashVouchers", {
        "VoucherNumber": "TEST-TRF-1", "VoucherDate": Day(4, 18), "VoucherType": "TRANSFER",
        "CashBoxID": Ref("BOXC"), "ToCashBoxID": Ref("BOXM"), "Category": "TRANSFER", "Amount": 1000,
        "ClosingID": Ref("CL1"), "EmployeeID": 1}),
    # Owner puts in 2000, later takes 300
    Row("V4", "CashVouchers", {
        "VoucherNumber": "TEST-CIN-1", "VoucherDate": Day(2, 9), "VoucherType": "IN",
        "CashBoxID": Ref("BOXM"), "Category": "OWNER", "Amount": 2000, "PartyName": "TEST المالك",
        "EmployeeID": 1}),
    Row("V5", "CashVouchers", {
        "VoucherNumber": "TEST-COT-3", "VoucherDate": Day(1, 12), "VoucherType": "OUT",
        "CashBoxID": Ref("BOXM"), "Category": "OWNER", "Amount": 300, "PartyName": "TEST المالك",
        "EmployeeID": 1}),
    # Manual entry 3 days ago: salaries of the month accrued (3000 = 2500 to pay + 500 social insurance)
    Row("MJ1", "ManualEntries", {
        "EntryNumber": "TEST-MJ-1", "EntryDate": Day(3), "Description": "TEST رواتب الشهر المستحقة",
        "TotalAmount": 3000, "EmployeeID": 1}),
    Row("MJ1A", "ManualEntryLines", {"ManualEntryID": Ref("MJ1"), "LineNumber": 1, "AccountCode": 5500,
                                     "Debit": 3000, "Credit": 0}),
    Row("MJ1B", "ManualEntryLines", {"ManualEntryID": Ref("MJ1"), "LineNumber": 2, "AccountCode": 2310,
                                     "Debit": 0, "Credit": 2500, "LineText": "TEST صافي الرواتب"}),
    Row("MJ1C", "ManualEntryLines", {"ManualEntryID": Ref("MJ1"), "LineNumber": 3, "AccountCode": 2320,
                                     "Debit": 0, "Credit": 500}),
]

# Test period: the last 30 days including today
PERIOD_START_DAYS_AGO = 30


@dataclass
class Check:
    label: str
    sql: str            # may contain {ref:KEY} and {day:N}
    expected: float
    params: Optional[Dict[str, str]] = None   # TempVars set before the check (values are fixture keys)


CHECKS: List[Check] = [
    # Stock (spec: buy 100, sell 10 -> 90; further movements bring P1 to 83)
    Check("رصيد المنتج 1 = 100 شراء − 10 − 2 بيع − 5 مرتجع شراء = 83",
          "SELECT CurrentQuantity FROM StockBalanceQuery WHERE ProductID = {ref:P1}", 83),
    Check("رصيد المنتج 1 من دفتر الحركات = 83",
          "SELECT LedgerQuantity FROM StockBalanceQuery WHERE ProductID = {ref:P1}", 83),
    Check("رصيد المنتج 2 = 200 − 5 − 10 + 2 − 1 = 186",
          "SELECT LedgerQuantity FROM StockBalanceQuery WHERE ProductID = {ref:P2}", 186),
    Check("لا توجد فروقات بين الكمية المسجلة والحركات",
          "SELECT COUNT(*) FROM StockBalanceQuery WHERE QuantityMismatch <> 0", 0),
    Check("قيمة مخزون المنتج 1 بالتكلفة = 83 × 60",
          "SELECT StockCostValue FROM StockBalanceQuery WHERE ProductID = {ref:P1}", 4980),
    Check("منخفض المخزون: منتج واحد فقط",
          "SELECT COUNT(*) FROM LowStockQuery", 1),
    Check("منخفض المخزون: المنتج 2 (186 ≤ 200)",
          "SELECT ShortageQty FROM LowStockQuery WHERE ProductID = {ref:P2}", 14),
    Check("طباعة فاتورة الشراء: سطران",
          "SELECT COUNT(*) FROM qryPurchaseDocPrint WHERE DocKind = 'PURCHASE' AND DocID = {ref:PUR1}", 2),
    Check("طباعة مرتجع الشراء: رقم السطر الأصلي ورقم الفاتورة الأصلية",
          "SELECT COUNT(*) FROM qryPurchaseDocPrint WHERE DocKind = 'RETURN' AND DocID = {ref:PRT1} "
          "AND LineNumber = 1 AND OriginalNumber = 'TEST-PUR-1' AND RemainingAmount = 345", 1),
    Check("طباعة سند الصرف: المبلغ وطريقة الدفع",
          "SELECT Amount FROM qryVoucherPrint WHERE DocKind = 'PAYMENT' AND DocID = {ref:PAY1}", 1000),
    Check("طباعة سند القبض",
          "SELECT COUNT(*) FROM qryVoucherPrint WHERE DocKind = 'RECEIPT' AND DocID = {ref:RCV1}", 1),
    Check("الجرد المفتوح: عجز المنتج 1 = −3 × 60",
          "SELECT DifferenceValue FROM StockCountQuery WHERE ProductID = {ref:P1}", -180),
    Check("الجرد المفتوح: صنف واحد لم يُعدّ بعد",
          "SELECT COUNT(*) FROM StockCountQuery WHERE ActualQuantity Is Null", 1),
    Check("غير المتحركة: منتج واحد",
          "SELECT COUNT(*) FROM SlowMovingProductsQuery", 1),
    Check("غير المتحركة: المنتج 3 بقيمة 200",
          "SELECT StockCostValue FROM SlowMovingProductsQuery WHERE ProductID = {ref:P3}", 200),
    Check("المخزون حسب التصنيف: تصنيف عام = 83 + 20",
          "SELECT TotalQuantity FROM StockByCategoryQuery WHERE CategoryID = 1", 103),
    Check("المخزون حسب التصنيف: قيمة تصنيف عام = 4980 + 200",
          "SELECT StockCostValue FROM StockByCategoryQuery WHERE CategoryID = 1", 5180),
    Check("المخزون حسب التصنيف: قيمة التصنيف 2 = 186 × 10",
          "SELECT StockCostValue FROM StockByCategoryQuery WHERE CategoryID = {ref:CAT2}", 1860),
    Check("حركة المنتج 2: رصيد أول المدة = 195",
          "SELECT NetQty FROM ProductMovementQuery WHERE SortKey = 0", 195,
          {"ProductID": "P2"}),
    Check("حركة المنتج 2: 3 حركات + سطر الرصيد السابق",
          "SELECT COUNT(*) FROM ProductMovementQuery", 4, {"ProductID": "P2"}),
    Check("حركة المنتج 2: الرصيد الختامي = 186",
          "SELECT Sum(NetQty) FROM ProductMovementQuery", 186, {"ProductID": "P2"}),

    # Sales
    Check("المبيعات اليومية: يوم الفاتورة الآجلة = 460",
          "SELECT NetSalesTotal FROM DailySalesQuery WHERE SaleDate = DateValue({day:5})", 460),
    Check("المبيعات اليومية: الآجل في نفس اليوم = 460",
          "SELECT CreditSales FROM DailySalesQuery WHERE SaleDate = DateValue({day:5})", 460),
    Check("المبيعات اليومية: يوم الفاتورة النقدية = 1150 نقدًا",
          "SELECT CashSales FROM DailySalesQuery WHERE SaleDate = DateValue({day:10})", 1150),
    Check("المبيعات اليومية: يوم المرتجع = −46",
          "SELECT NetSalesTotal FROM DailySalesQuery WHERE SaleDate = DateValue({day:2})", -46),
    Check("المبيعات الشهرية: صافي كل الأشهر = 100 + 1000 + 400 − 40",
          "SELECT Sum(NetSalesExVAT) FROM MonthlySalesQuery", 1460),
    Check("المبيعات الشهرية: مجمل ربح كل الأشهر = 1460 − 850",
          "SELECT Sum(GrossProfit) FROM MonthlySalesQuery", 610),
    Check("المبيعات خلال الفترة: فاتورتان ومرتجع",
          "SELECT COUNT(*) FROM SalesByPeriodQuery", 3),
    Check("المبيعات خلال الفترة: الإجمالي = 1150 + 460 − 46",
          "SELECT Sum(GrossAmount) FROM SalesByPeriodQuery", 1564),
    Check("المبيعات حسب المنتج: المنتج 1 صافي كمية 12",
          "SELECT NetQty FROM SalesByProductQuery WHERE ProductID = {ref:P1}", 12),
    Check("المبيعات حسب المنتج: ربح المنتج 1 = 1200 − 720",
          "SELECT GrossProfit FROM SalesByProductQuery WHERE ProductID = {ref:P1}", 480),
    Check("المبيعات حسب المنتج: المنتج 2 مرتجع 2",
          "SELECT QtyReturned FROM SalesByProductQuery WHERE ProductID = {ref:P2}", 2),
    Check("المبيعات حسب المنتج: صافي مبيعات المنتج 2 = 200 − 40",
          "SELECT NetSales FROM SalesByProductQuery WHERE ProductID = {ref:P2}", 160),
    Check("المبيعات حسب التصنيف: المجموع = المبيعات حسب المنتج",
          "SELECT (SELECT Sum(CategoryTotal) FROM SalesByCategoryQuery) - "
          "(SELECT Sum(SalesTotal) FROM SalesByProductQuery) FROM Settings", 0),
    Check("المبيعات حسب التصنيف: صافي الكمية = 12 + 8",
          "SELECT Sum(CategoryQty) FROM SalesByCategoryQuery", 20),
    Check("الأكثر مبيعًا: المنتج 1",
          "SELECT TOP 1 ProductID FROM BestSellingProductsQuery ORDER BY NetQty DESC, NetSales DESC",
          "ref:P1"),
    Check("الأقل مبيعًا: المنتج 3 (لم يُبع)",
          "SELECT TOP 1 ProductID FROM LeastSellingProductsQuery ORDER BY NetQtySold, ProductName",
          "ref:P3"),

    # Purchases
    Check("المشتريات خلال الفترة: مرتجع الشراء فقط",
          "SELECT COUNT(*) FROM PurchasesQuery", 1),
    Check("المشتريات خلال الفترة: −345",
          "SELECT Sum(GrossAmount) FROM PurchasesQuery", -345),

    # Customers / suppliers
    Check("رصيد العميل الآجل = 50 + 460 − 100 − 46 − 200",
          "SELECT Balance FROM CustomerBalanceQuery WHERE CustomerID = {ref:C2}", 164),
    Check("رصيد العميل النقدي = 0",
          "SELECT Balance FROM CustomerBalanceQuery WHERE CustomerID = 1", 0),
    Check("العملاء المدينون: عميل واحد",
          "SELECT COUNT(*) FROM CustomersWithDebtQuery", 1),
    Check("كشف العميل: الرصيد السابق = 50",
          "SELECT Debit FROM CustomerStatementQuery WHERE SortKey = 0", 50, {"CustomerID": "C2"}),
    Check("كشف العميل: رصيد سابق + 3 حركات",
          "SELECT COUNT(*) FROM CustomerStatementQuery", 4, {"CustomerID": "C2"}),
    Check("كشف العميل: الرصيد الختامي = 164",
          "SELECT Sum(Debit) - Sum(Credit) FROM CustomerStatementQuery", 164, {"CustomerID": "C2"}),
    Check("رصيد المورد = 9200 − 5000 − 1000 − 345",
          "SELECT Balance FROM SupplierBalanceQuery WHERE SupplierID = {ref:S1}", 2855),
    Check("كشف المورد: الرصيد السابق = 4200",
          "SELECT Credit FROM SupplierStatementQuery WHERE SortKey = 0", 4200, {"SupplierID": "S1"}),
    Check("كشف المورد: الرصيد الختامي = 2855",
          "SELECT Sum(Credit) - Sum(Debit) FROM SupplierStatementQuery", 2855, {"SupplierID": "S1"}),

    # Expenses
    Check("المصروفات خلال الفترة: مصروف واحد",
          "SELECT COUNT(*) FROM ExpensesQuery", 1),
    Check("المصروفات خلال الفترة: 230 شامل الضريبة",
          "SELECT Sum(TotalAmount) FROM ExpensesQuery", 230),
    Check("المصروفات حسب النوع: الكهرباء 200",
          "SELECT AmountExVAT FROM ExpensesByTypeQuery", 200),

    # Profit (spec section 11)
    Check("الأرباح: صافي المبيعات = 1000 + 400 − 40",
          "SELECT NetSales FROM ProfitQuery", 1360),
    Check("الأرباح: تكلفة المبيعات = 600 + 220 − 20",
          "SELECT CostOfSales FROM ProfitQuery", 800),
    Check("الأرباح: مجمل الربح = 1360 − 800",
          "SELECT GrossProfit FROM ProfitQuery", 560),
    Check("الأرباح: فروقات المخزون = −10 (منتج تالف)",
          "SELECT InventoryAdjustments FROM ProfitQuery", -10),
    Check("الأرباح: المصروفات = 200",
          "SELECT TotalExpenses FROM ProfitQuery", 200),
    Check("الأرباح: صافي الربح = 560 − 10 − 200",
          "SELECT NetProfit FROM ProfitQuery", 350),

    # VAT
    Check("الضريبة: ضريبة المخرجات = 150 + 60 − 6",
          "SELECT OutputVAT FROM VatSummaryQuery", 204),
    Check("الضريبة: ضريبة المدخلات = −45 (مرتجع شراء) + 30 (مصروف)",
          "SELECT InputVAT FROM VatSummaryQuery", -15),
    Check("الضريبة: الصافي المستحق = 204 + 15",
          "SELECT NetVATDue FROM VatSummaryQuery", 219),

    # Printing
    Check("طباعة الفاتورة الآجلة: سطران",
          "SELECT COUNT(*) FROM qrySalesDocPrint WHERE DocKind = 'SALE' AND DocID = {ref:INV2}", 2),
    Check("طباعة الفاتورة الآجلة: مجموع الأسطر = 460",
          "SELECT Sum(LineTotal) FROM qrySalesDocPrint WHERE DocKind = 'SALE' AND DocID = {ref:INV2}", 460),
    Check("طباعة الإشعار الدائن: سطر واحد بقيمة 46",
          "SELECT Sum(LineTotal) FROM qrySalesDocPrint WHERE DocKind = 'RETURN' AND DocID = {ref:CRN1}", 46),

    # Treasury
    Check("رصيد صندوق الكاشير = 500 + 115 + 1150 + 100 + 200 − 230 − 50 − 15 − 1000",
          "SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = {ref:BOXC}", 770),
    Check("رصيد الخزينة = 10000 − 5000 + 1000 + 2000 − 300",
          "SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = {ref:BOXM}", 7700),
    Check("الصناديق المسجلة بدون حركة رصيدها صفر",
          "SELECT Sum(Balance) FROM CashBoxBalanceQuery WHERE CashBoxID <= 2", 0),
    Check("حركة صندوق الكاشير: رصيد أول المدة = 500 + 115 − 50",
          "SELECT AmountIn FROM CashStatementQuery WHERE SortKey = 0", 565, {"CashBoxID": "BOXC"}),
    Check("حركة صندوق الكاشير: 6 حركات في الفترة",
          "SELECT COUNT(*) FROM CashStatementQuery WHERE SortKey = 1", 6, {"CashBoxID": "BOXC"}),
    Check("حركة صندوق الكاشير: رصيد آخر المدة = 770",
          "SELECT Sum(AmountIn) - Sum(AmountOut) FROM CashStatementQuery", 770, {"CashBoxID": "BOXC"}),
    Check("حركة الخزينة: التحويل من الكاشير داخل = 1000",
          "SELECT AmountIn FROM CashStatementQuery WHERE MoveType = 'TRANSFER_IN'", 1000,
          {"CashBoxID": "BOXM"}),
    Check("يومية صندوق الكاشير يوم التصفية: رصيد أول اليوم = 565 + 1150 + 100 − 230",
          "SELECT OpeningBalance FROM CashDailyQuery WHERE CashDay = DateValue({day:4})", 1585,
          {"CashBoxID": "BOXC"}),
    Check("يومية صندوق الكاشير يوم التصفية: المدفوعات = 15 عجز + 1000 تحويل",
          "SELECT Payments FROM CashDailyQuery WHERE CashDay = DateValue({day:4})", 1015,
          {"CashBoxID": "BOXC"}),
    Check("يومية صندوق الكاشير يوم التصفية: رصيد آخر اليوم = 570",
          "SELECT ClosingBalance FROM CashDailyQuery WHERE CashDay = DateValue({day:4})", 570,
          {"CashBoxID": "BOXC"}),
    Check("يومية صندوق الكاشير: آخر يوم = الرصيد الحالي",
          "SELECT ClosingBalance FROM CashDailyQuery WHERE CashDay = DateValue({day:3})", 770,
          {"CashBoxID": "BOXC"}),
    Check("تصفيات الكاشير خلال الفترة: تصفية واحدة بعجز 15",
          "SELECT Difference FROM CashClosingsQuery", -15, {"CashBoxID": "BOXC"}),
    Check("طباعة سند صرف المصروف: نوع المصروف",
          "SELECT COUNT(*) FROM qryCashVoucherPrint WHERE DocID = {ref:V1} AND ExpenseTypeName = 'مصروفات أخرى'", 1),
    Check("طباعة سند التحويل: الصندوق المستلم",
          "SELECT COUNT(*) FROM qryCashVoucherPrint WHERE DocID = {ref:V3} AND ToBoxName = 'TEST الخزينة'", 1),

    # Journal entries (the lines the sync posts)
    Check("قيود Sale: كل قيد متوازن",
          "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalSale GROUP BY SourceType, SourceID "
          "HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0),
    Check("قيود SalesReturn: كل قيد متوازن",
          "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalSalesReturn GROUP BY SourceType, SourceID "
          "HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0),
    Check("قيود Purchase: كل قيد متوازن",
          "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalPurchase GROUP BY SourceType, SourceID "
          "HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0),
    Check("قيود PurchaseReturn: كل قيد متوازن",
          "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalPurchaseReturn GROUP BY SourceType, SourceID "
          "HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0),
    Check("قيود Payments: كل قيد متوازن",
          "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalPayments GROUP BY SourceType, SourceID "
          "HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0),
    Check("قيود Expense: كل قيد متوازن",
          "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalExpense GROUP BY SourceType, SourceID "
          "HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0),
    Check("قيود CashVoucher: كل قيد متوازن",
          "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalCashVoucher GROUP BY SourceType, SourceID "
          "HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0),
    Check("قيود Stock: كل قيد متوازن",
          "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalStock GROUP BY SourceType, SourceID "
          "HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0),
    Check("قيود Opening: كل قيد متوازن",
          "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalOpening GROUP BY SourceType, SourceID "
          "HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0),
    Check("قيد الفاتورة الآجلة: 6 أسطر (نقدي، عميل، مبيعات، ضريبة، تكلفة، مخزون)",
          "SELECT COUNT(*) FROM qryJournalSale WHERE SourceID = {ref:INV2}", 6),
    Check("قيد الفاتورة الآجلة: المتبقي على العميل 360 في ذمم العملاء",
          "SELECT Debit FROM qryJournalSale WHERE SourceID = {ref:INV2} AND AccountCode = 1300", 360),
    Check("قيد الفاتورة الآجلة: التكلفة = 2×60 + 10×10",
          "SELECT Debit FROM qryJournalSale WHERE SourceID = {ref:INV2} AND AccountCode = 5100", 220),
    Check("ذمم العملاء من القيود = أرصدة العملاء (164)",
          'SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 1300', 164),
    Check("ذمم الموردين من القيود = رصيد المورد (2855 دائن)",
          'SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 2100', -2855),
    Check("المخزون من القيود = قيمة المخزون بالتكلفة (7040)",
          'SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalStock) AS x WHERE AccountCode = 1400', 7040),
    Check("صندوق الكاشير من القيود = رصيده (770)",
          'SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalExpense UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalCashVoucher UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalStock UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 110000 + {ref:BOXC}', 770),
    Check("الخزينة من القيود = رصيدها (7700)",
          'SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalExpense UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalCashVoucher UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalStock UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 110000 + {ref:BOXM}', 7700),
    Check("ضريبة المخرجات من القيود = 15 + 150 + 60 − 6",
          'SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn) AS x WHERE AccountCode = 2200', -219),
    Check("مصروف سند النقدية لا يُقيَّد مرتين",
          "SELECT COUNT(*) FROM qryJournalExpense WHERE SourceID = {ref:EXPV}", 0),
    Check("القيد اليدوي: 3 أسطر متوازنة (3000)",
          "SELECT COUNT(*) FROM qryJournalManual WHERE SourceID = {ref:MJ1} AND SourceType = 'MANUAL'", 3),
    Check("قيود Manual: كل قيد متوازن",
          "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalManual GROUP BY SourceType, SourceID "
          "HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0),
    Check("سند صرف المصروف يُقيَّد على حساب نوع المصروف",
          "SELECT Debit FROM qryJournalCashVoucher WHERE SourceID = {ref:V1} AND AccountCode = 530009", 50),

    # Integrity
    Check("فحص السلامة: لا توجد مشكلات",
          "SELECT COUNT(*) FROM IntegrityCheckQuery", 0),
]

# Deliberate corruption -> IntegrityCheckQuery must report it
CORRUPTIONS = [
    ("UPDATE [Products] SET [CurrentQuantity] = 80 WHERE [ProductID] = {ref:P1}",
     Check("فحص السلامة يكتشف تلاعبًا بالكمية",
           "SELECT COUNT(*) FROM IntegrityCheckQuery WHERE IssueCode = 'STOCK_MISMATCH'", 1)),
    ("UPDATE [Customers] SET [CurrentBalance] = 0 WHERE [CustomerID] = {ref:C2}",
     Check("فحص السلامة يكتشف رصيد عميل خاطئ",
           "SELECT ExpectedValue FROM IntegrityCheckQuery WHERE IssueCode = 'CUSTOMER_BALANCE'", 164)),
    ("UPDATE [SalesReturns] SET [CustomerID] = 1 WHERE [SalesReturnID] = {ref:CRN1}",
     Check("فحص السلامة يكتشف مرتجعًا لعميل مختلف",
           "SELECT COUNT(*) FROM IntegrityCheckQuery WHERE IssueCode = 'RETURN_CUSTOMER'", 1)),
]


def query(name: str) -> Query:
    for q in QUERIES:
        if q.name == name:
            return q
    raise KeyError(name)
