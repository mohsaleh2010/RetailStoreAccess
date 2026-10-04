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
WHERE p.IsActive = True AND p.CurrentQuantity <= p.MinimumQuantity
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
       e.TotalAmount, pm.MethodName, e.Description, em.EmployeeName
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
       d.Discount AS LineDiscount, d.NetAmount, d.VATRate, d.Tax AS LineTax, d.LineTotal
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
       rd.Discount, rd.NetAmount, rd.VATRate, rd.Tax, rd.LineTotal
FROM (((((SalesReturns AS r INNER JOIN SalesReturnDetails AS rd ON r.SalesReturnID = rd.SalesReturnID)
        INNER JOIN SalesInvoices AS o ON r.SalesInvoiceID = o.SalesInvoiceID)
       INNER JOIN Products AS p ON rd.ProductID = p.ProductID)
      INNER JOIN Units AS u ON p.UnitID = u.UnitID)
     INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID)
    INNER JOIN Employees AS e ON r.EmployeeID = e.EmployeeID"""),

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
        "PaidAmount": 5000, "RemainingAmount": 4200}),
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
        "RemainingAmount": 0, "AmountTendered": 115, "ChangeDue": 0}),
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
        "PaidAmount": 1150, "RemainingAmount": 0, "AmountTendered": 1200, "ChangeDue": 50}),
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
        "RemainingAmount": 360, "AmountTendered": 100, "ChangeDue": 0}),
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
        "Amount": 200, "PaymentMethodID": 1, "EmployeeID": 1}),

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

    # Expenses: electricity inside the period, rent outside it
    Row("EXP1", "Expenses", {
        "ExpenseNumber": "TEST-EXP-1", "ExpenseDate": Day(6), "ExpenseTypeID": 2, "Amount": 200,
        "Tax": 30, "TotalAmount": 230, "PaymentMethodID": 1, "Description": "TEST كهرباء",
        "EmployeeID": 1}),
    Row("EXP2", "Expenses", {
        "ExpenseNumber": "TEST-EXP-2", "ExpenseDate": Day(40), "ExpenseTypeID": 1, "Amount": 1000,
        "Tax": 0, "TotalAmount": 1000, "PaymentMethodID": 3, "Description": "TEST إيجار",
        "EmployeeID": 1}),
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
