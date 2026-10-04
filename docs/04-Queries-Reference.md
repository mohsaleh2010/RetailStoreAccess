# مرجع الاستعلامات (Queries Reference)

> ملف مُولَّد تلقائيًا من `tools/queries.py` – لا تعدّله يدويًا.

عدد الاستعلامات: **44**. الاستعلامات التي تبدأ بـ `qry` مساعدة تستخدمها الاستعلامات الأخرى؛ البقية تُستخدم مباشرة في التقارير والنماذج. ⭐ = مطلوب بالاسم في البرومبت.

| # | الاستعلام | الوصف | المعاملات |
|---|---|---|---|
| 1 | [`qrySalesDocuments`](#qrysalesdocuments) | مستندات البيع: الفواتير (+) والمرتجعات (−) بقيم موقّعة |  |
| 2 | [`qrySalesLineItems`](#qrysaleslineitems) | أسطر البيع والمرتجعات مع التكلفة (أساس تحليل المنتجات والأرباح) |  |
| 3 | [`qrySalesLinesInPeriod`](#qrysaleslinesinperiod) | أسطر البيع والمرتجعات داخل الفترة | `PeriodStart`, `PeriodEnd` |
| 4 | [`DailySalesQuery`](#dailysalesquery) ⭐ | المبيعات اليومية: عدد الفواتير والمرتجعات والصافي والنقدي والآجل |  |
| 5 | [`qrySalesMonthlyDocs`](#qrysalesmonthlydocs) | تجميع شهري لمستندات البيع |  |
| 6 | [`qrySalesMonthlyCost`](#qrysalesmonthlycost) | تكلفة البضاعة المباعة شهريًا |  |
| 7 | [`MonthlySalesQuery`](#monthlysalesquery) ⭐ | المبيعات الشهرية مع التكلفة ومجمل الربح |  |
| 8 | [`SalesByPeriodQuery`](#salesbyperiodquery) | فواتير ومرتجعات البيع خلال فترة مع العميل والكاشير | `PeriodStart`, `PeriodEnd` |
| 9 | [`SalesByProductQuery`](#salesbyproductquery) | المبيعات حسب المنتج خلال فترة (كمية، صافي، تكلفة، ربح) | `PeriodStart`, `PeriodEnd` |
| 10 | [`BestSellingProductsQuery`](#bestsellingproductsquery) ⭐ | أفضل المنتجات مبيعًا خلال فترة (حسب صافي الكمية) | `PeriodStart`, `PeriodEnd` |
| 11 | [`LeastSellingProductsQuery`](#leastsellingproductsquery) | أقل المنتجات مبيعًا خلال فترة (تشمل المنتجات التي لم تُبع) | `PeriodStart`, `PeriodEnd` |
| 12 | [`qryPurchaseDocuments`](#qrypurchasedocuments) | مستندات الشراء: الفواتير (+) والمرتجعات (−) |  |
| 13 | [`PurchasesQuery`](#purchasesquery) | فواتير ومرتجعات الشراء خلال فترة مع المورد | `PeriodStart`, `PeriodEnd` |
| 14 | [`qryProductLedger`](#qryproductledger) | رصيد كل منتج من دفتر حركة المخزون |  |
| 15 | [`qryProductLastSale`](#qryproductlastsale) | تاريخ آخر بيع لكل منتج |  |
| 16 | [`StockBalanceQuery`](#stockbalancequery) ⭐ | المخزون الحالي: الكمية والقيمة بالتكلفة وبسعر البيع ومطابقتها مع الحركات |  |
| 17 | [`LowStockQuery`](#lowstockquery) ⭐ | المنتجات منخفضة المخزون: CurrentQuantity <= MinimumQuantity |  |
| 18 | [`ProductMovementQuery`](#productmovementquery) | حركة منتج خلال فترة مع رصيد أول المدة (الرصيد التراكمي في التقرير) | `PeriodStart`, `PeriodEnd`, `ProductID` |
| 19 | [`SlowMovingProductsQuery`](#slowmovingproductsquery) ⭐ | المنتجات غير المتحركة: لها رصيد ولم تُبع منذ عدد الأيام المحدد في الإعدادات |  |
| 20 | [`StockByCategoryQuery`](#stockbycategoryquery) | المخزون حسب التصنيف: عدد المنتجات والكمية والقيمة |  |
| 21 | [`qryCustomerLedger`](#qrycustomerledger) | دفتر حساب العملاء: مدين (عليه) / دائن (له) |  |
| 22 | [`qryCustomerLedgerTotals`](#qrycustomerledgertotals) | مجاميع حساب كل عميل |  |
| 23 | [`CustomerBalanceQuery`](#customerbalancequery) ⭐ | رصيد كل عميل محسوبًا من الحركات (موجب = عليه للمحل) |  |
| 24 | [`CustomersWithDebtQuery`](#customerswithdebtquery) | العملاء الذين عليهم مبالغ مستحقة |  |
| 25 | [`CustomerStatementQuery`](#customerstatementquery) | كشف حساب عميل لفترة: رصيد سابق ثم الحركات | `PeriodStart`, `PeriodEnd`, `CustomerID` |
| 26 | [`qrySupplierLedger`](#qrysupplierledger) | دفتر حساب الموردين: دائن (للمورد) / مدين (سُدِّد له) |  |
| 27 | [`qrySupplierLedgerTotals`](#qrysupplierledgertotals) | مجاميع حساب كل مورد |  |
| 28 | [`SupplierBalanceQuery`](#supplierbalancequery) ⭐ | رصيد كل مورد محسوبًا من الحركات (موجب = مستحق للمورد) |  |
| 29 | [`SupplierStatementQuery`](#supplierstatementquery) | كشف حساب مورد لفترة: رصيد سابق ثم الحركات | `PeriodStart`, `PeriodEnd`, `SupplierID` |
| 30 | [`ExpensesQuery`](#expensesquery) | المصروفات خلال فترة | `PeriodStart`, `PeriodEnd` |
| 31 | [`ExpensesByTypeQuery`](#expensesbytypequery) | المصروفات مجمّعة حسب النوع خلال فترة | `PeriodStart`, `PeriodEnd` |
| 32 | [`qryProfitSales`](#qryprofitsales) | صافي المبيعات وتكلفتها خلال الفترة | `PeriodStart`, `PeriodEnd` |
| 33 | [`qryProfitAdjustments`](#qryprofitadjustments) | قيمة فروقات المخزون (جرد، إضافة، خصم) خلال الفترة | `PeriodStart`, `PeriodEnd` |
| 34 | [`qryProfitExpenses`](#qryprofitexpenses) | المصروفات (بدون ضريبة) خلال الفترة | `PeriodStart`, `PeriodEnd` |
| 35 | [`ProfitQuery`](#profitquery) ⭐ | الأرباح: صافي المبيعات − التكلفة = مجمل الربح؛ ثم ± فروقات المخزون − المصروفات = صافي الربح | `PeriodStart`, `PeriodEnd` |
| 36 | [`qryVatOutput`](#qryvatoutput) | ضريبة المخرجات (المبيعات ناقص المرتجعات) | `PeriodStart`, `PeriodEnd` |
| 37 | [`qryVatInputPurchases`](#qryvatinputpurchases) | ضريبة المدخلات من المشتريات (ناقص المرتجعات) | `PeriodStart`, `PeriodEnd` |
| 38 | [`qryVatInputExpenses`](#qryvatinputexpenses) | ضريبة المدخلات من المصروفات | `PeriodStart`, `PeriodEnd` |
| 39 | [`VatSummaryQuery`](#vatsummaryquery) | ملخص ضريبة القيمة المضافة للفترة (للإقرار الضريبي) | `PeriodStart`, `PeriodEnd` |
| 40 | [`qrySalesInvoiceLineTotals`](#qrysalesinvoicelinetotals) | مجموع أسطر كل فاتورة بيع |  |
| 41 | [`qryPurchaseInvoiceLineTotals`](#qrypurchaseinvoicelinetotals) | مجموع أسطر كل فاتورة شراء |  |
| 42 | [`qrySalesReturnedQty`](#qrysalesreturnedqty) | الكمية المرتجعة من كل سطر فاتورة بيع |  |
| 43 | [`qryPurchaseReturnedQty`](#qrypurchasereturnedqty) | الكمية المرتجعة للمورد من كل سطر فاتورة شراء |  |
| 44 | [`IntegrityCheckQuery`](#integritycheckquery) | فحص سلامة البيانات: أي سطر هنا مشكلة يجب مراجعتها (النتيجة الفارغة = سليم) |  |

## بيانات الاختبار والنتائج المتوقعة

الفترة: آخر 30 يومًا حتى اليوم. التواريخ نسبية إلى يوم التشغيل.

| # | الاختبار | الاستعلام | المتوقع |
|---|---|---|---|
| 1 | رصيد المنتج 1 = 100 شراء − 10 − 2 بيع − 5 مرتجع شراء = 83 | `SELECT CurrentQuantity FROM StockBalanceQuery WHERE ProductID = {ref:P1}` | 83 |
| 2 | رصيد المنتج 1 من دفتر الحركات = 83 | `SELECT LedgerQuantity FROM StockBalanceQuery WHERE ProductID = {ref:P1}` | 83 |
| 3 | رصيد المنتج 2 = 200 − 5 − 10 + 2 − 1 = 186 | `SELECT LedgerQuantity FROM StockBalanceQuery WHERE ProductID = {ref:P2}` | 186 |
| 4 | لا توجد فروقات بين الكمية المسجلة والحركات | `SELECT COUNT(*) FROM StockBalanceQuery WHERE QuantityMismatch <> 0` | 0 |
| 5 | قيمة مخزون المنتج 1 بالتكلفة = 83 × 60 | `SELECT StockCostValue FROM StockBalanceQuery WHERE ProductID = {ref:P1}` | 4980 |
| 6 | منخفض المخزون: منتج واحد فقط | `SELECT COUNT(*) FROM LowStockQuery` | 1 |
| 7 | منخفض المخزون: المنتج 2 (186 ≤ 200) | `SELECT ShortageQty FROM LowStockQuery WHERE ProductID = {ref:P2}` | 14 |
| 8 | غير المتحركة: منتج واحد | `SELECT COUNT(*) FROM SlowMovingProductsQuery` | 1 |
| 9 | غير المتحركة: المنتج 3 بقيمة 200 | `SELECT StockCostValue FROM SlowMovingProductsQuery WHERE ProductID = {ref:P3}` | 200 |
| 10 | المخزون حسب التصنيف: تصنيف عام = 83 + 20 | `SELECT TotalQuantity FROM StockByCategoryQuery WHERE CategoryID = 1` | 103 |
| 11 | المخزون حسب التصنيف: قيمة تصنيف عام = 4980 + 200 | `SELECT StockCostValue FROM StockByCategoryQuery WHERE CategoryID = 1` | 5180 |
| 12 | المخزون حسب التصنيف: قيمة التصنيف 2 = 186 × 10 | `SELECT StockCostValue FROM StockByCategoryQuery WHERE CategoryID = {ref:CAT2}` | 1860 |
| 13 | حركة المنتج 2: رصيد أول المدة = 195 | `SELECT NetQty FROM ProductMovementQuery WHERE SortKey = 0` | 195 |
| 14 | حركة المنتج 2: 3 حركات + سطر الرصيد السابق | `SELECT COUNT(*) FROM ProductMovementQuery` | 4 |
| 15 | حركة المنتج 2: الرصيد الختامي = 186 | `SELECT Sum(NetQty) FROM ProductMovementQuery` | 186 |
| 16 | المبيعات اليومية: يوم الفاتورة الآجلة = 460 | `SELECT NetSalesTotal FROM DailySalesQuery WHERE SaleDate = DateValue({day:5})` | 460 |
| 17 | المبيعات اليومية: الآجل في نفس اليوم = 460 | `SELECT CreditSales FROM DailySalesQuery WHERE SaleDate = DateValue({day:5})` | 460 |
| 18 | المبيعات اليومية: يوم الفاتورة النقدية = 1150 نقدًا | `SELECT CashSales FROM DailySalesQuery WHERE SaleDate = DateValue({day:10})` | 1150 |
| 19 | المبيعات اليومية: يوم المرتجع = −46 | `SELECT NetSalesTotal FROM DailySalesQuery WHERE SaleDate = DateValue({day:2})` | -46 |
| 20 | المبيعات الشهرية: صافي كل الأشهر = 100 + 1000 + 400 − 40 | `SELECT Sum(NetSalesExVAT) FROM MonthlySalesQuery` | 1460 |
| 21 | المبيعات الشهرية: مجمل ربح كل الأشهر = 1460 − 850 | `SELECT Sum(GrossProfit) FROM MonthlySalesQuery` | 610 |
| 22 | المبيعات خلال الفترة: فاتورتان ومرتجع | `SELECT COUNT(*) FROM SalesByPeriodQuery` | 3 |
| 23 | المبيعات خلال الفترة: الإجمالي = 1150 + 460 − 46 | `SELECT Sum(GrossAmount) FROM SalesByPeriodQuery` | 1564 |
| 24 | المبيعات حسب المنتج: المنتج 1 صافي كمية 12 | `SELECT NetQty FROM SalesByProductQuery WHERE ProductID = {ref:P1}` | 12 |
| 25 | المبيعات حسب المنتج: ربح المنتج 1 = 1200 − 720 | `SELECT GrossProfit FROM SalesByProductQuery WHERE ProductID = {ref:P1}` | 480 |
| 26 | المبيعات حسب المنتج: المنتج 2 مرتجع 2 | `SELECT QtyReturned FROM SalesByProductQuery WHERE ProductID = {ref:P2}` | 2 |
| 27 | المبيعات حسب المنتج: صافي مبيعات المنتج 2 = 200 − 40 | `SELECT NetSales FROM SalesByProductQuery WHERE ProductID = {ref:P2}` | 160 |
| 28 | الأكثر مبيعًا: المنتج 1 | `SELECT TOP 1 ProductID FROM BestSellingProductsQuery ORDER BY NetQty DESC, NetSales DESC` | رقم P1 |
| 29 | الأقل مبيعًا: المنتج 3 (لم يُبع) | `SELECT TOP 1 ProductID FROM LeastSellingProductsQuery ORDER BY NetQtySold, ProductName` | رقم P3 |
| 30 | المشتريات خلال الفترة: مرتجع الشراء فقط | `SELECT COUNT(*) FROM PurchasesQuery` | 1 |
| 31 | المشتريات خلال الفترة: −345 | `SELECT Sum(GrossAmount) FROM PurchasesQuery` | -345 |
| 32 | رصيد العميل الآجل = 50 + 460 − 100 − 46 − 200 | `SELECT Balance FROM CustomerBalanceQuery WHERE CustomerID = {ref:C2}` | 164 |
| 33 | رصيد العميل النقدي = 0 | `SELECT Balance FROM CustomerBalanceQuery WHERE CustomerID = 1` | 0 |
| 34 | العملاء المدينون: عميل واحد | `SELECT COUNT(*) FROM CustomersWithDebtQuery` | 1 |
| 35 | كشف العميل: الرصيد السابق = 50 | `SELECT Debit FROM CustomerStatementQuery WHERE SortKey = 0` | 50 |
| 36 | كشف العميل: رصيد سابق + 3 حركات | `SELECT COUNT(*) FROM CustomerStatementQuery` | 4 |
| 37 | كشف العميل: الرصيد الختامي = 164 | `SELECT Sum(Debit) - Sum(Credit) FROM CustomerStatementQuery` | 164 |
| 38 | رصيد المورد = 9200 − 5000 − 1000 − 345 | `SELECT Balance FROM SupplierBalanceQuery WHERE SupplierID = {ref:S1}` | 2855 |
| 39 | كشف المورد: الرصيد السابق = 4200 | `SELECT Credit FROM SupplierStatementQuery WHERE SortKey = 0` | 4200 |
| 40 | كشف المورد: الرصيد الختامي = 2855 | `SELECT Sum(Credit) - Sum(Debit) FROM SupplierStatementQuery` | 2855 |
| 41 | المصروفات خلال الفترة: مصروف واحد | `SELECT COUNT(*) FROM ExpensesQuery` | 1 |
| 42 | المصروفات خلال الفترة: 230 شامل الضريبة | `SELECT Sum(TotalAmount) FROM ExpensesQuery` | 230 |
| 43 | المصروفات حسب النوع: الكهرباء 200 | `SELECT AmountExVAT FROM ExpensesByTypeQuery` | 200 |
| 44 | الأرباح: صافي المبيعات = 1000 + 400 − 40 | `SELECT NetSales FROM ProfitQuery` | 1360 |
| 45 | الأرباح: تكلفة المبيعات = 600 + 220 − 20 | `SELECT CostOfSales FROM ProfitQuery` | 800 |
| 46 | الأرباح: مجمل الربح = 1360 − 800 | `SELECT GrossProfit FROM ProfitQuery` | 560 |
| 47 | الأرباح: فروقات المخزون = −10 (منتج تالف) | `SELECT InventoryAdjustments FROM ProfitQuery` | -10 |
| 48 | الأرباح: المصروفات = 200 | `SELECT TotalExpenses FROM ProfitQuery` | 200 |
| 49 | الأرباح: صافي الربح = 560 − 10 − 200 | `SELECT NetProfit FROM ProfitQuery` | 350 |
| 50 | الضريبة: ضريبة المخرجات = 150 + 60 − 6 | `SELECT OutputVAT FROM VatSummaryQuery` | 204 |
| 51 | الضريبة: ضريبة المدخلات = −45 (مرتجع شراء) + 30 (مصروف) | `SELECT InputVAT FROM VatSummaryQuery` | -15 |
| 52 | الضريبة: الصافي المستحق = 204 + 15 | `SELECT NetVATDue FROM VatSummaryQuery` | 219 |
| 53 | فحص السلامة: لا توجد مشكلات | `SELECT COUNT(*) FROM IntegrityCheckQuery` | 0 |

## qrySalesDocuments

مستندات البيع: الفواتير (+) والمرتجعات (−) بقيم موقّعة

```sql
SELECT 'SALE' AS DocType, h.SalesInvoiceID AS DocID, h.InvoiceNumber AS DocNumber,
       h.InvoiceDate AS DocDate, h.CustomerID, h.EmployeeID, h.PaymentType,
       h.TaxableAmount AS NetAmount, h.Tax AS VATAmount, h.TotalAmount AS GrossAmount,
       h.PaidAmount AS SettledAmount, h.RemainingAmount AS OnAccount
FROM SalesInvoices AS h
UNION ALL
SELECT 'RETURN', r.SalesReturnID, r.ReturnNumber, r.ReturnDate, r.CustomerID, r.EmployeeID,
       r.RefundType, -r.TaxableAmount, -r.Tax, -r.TotalAmount, -r.RefundedAmount,
       -(r.TotalAmount - r.RefundedAmount)
FROM SalesReturns AS r
```

## qrySalesLineItems

أسطر البيع والمرتجعات مع التكلفة (أساس تحليل المنتجات والأرباح)

```sql
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
     ON r.SalesReturnID = d.SalesReturnID
```

## qrySalesLinesInPeriod

أسطر البيع والمرتجعات داخل الفترة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT * FROM qrySalesLineItems
WHERE DocDate >= QDate('PeriodStart') AND DocDate < QDate('PeriodEnd')
```

## DailySalesQuery

المبيعات اليومية: عدد الفواتير والمرتجعات والصافي والنقدي والآجل

```sql
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
ORDER BY DateValue(DocDate) DESC
```

## qrySalesMonthlyDocs

تجميع شهري لمستندات البيع

```sql
SELECT Year(DocDate) AS SalesYear, Month(DocDate) AS SalesMonth,
       Sum(IIf(DocType = 'SALE', 1, 0)) AS InvoiceCount,
       Sum(IIf(DocType = 'RETURN', 1, 0)) AS ReturnCount,
       Sum(NetAmount) AS NetSalesExVAT, Sum(VATAmount) AS NetVAT,
       Sum(GrossAmount) AS NetSalesTotal
FROM qrySalesDocuments
GROUP BY Year(DocDate), Month(DocDate)
```

## qrySalesMonthlyCost

تكلفة البضاعة المباعة شهريًا

```sql
SELECT Year(DocDate) AS SalesYear, Month(DocDate) AS SalesMonth, Sum(LineCost) AS MonthCost
FROM qrySalesLineItems
GROUP BY Year(DocDate), Month(DocDate)
```

## MonthlySalesQuery

المبيعات الشهرية مع التكلفة ومجمل الربح

```sql
SELECT d.SalesYear, d.SalesMonth, d.InvoiceCount, d.ReturnCount, d.NetSalesExVAT,
       d.NetVAT, d.NetSalesTotal, CCur(Nz(c.MonthCost, 0)) AS CostOfSales,
       d.NetSalesExVAT - CCur(Nz(c.MonthCost, 0)) AS GrossProfit
FROM qrySalesMonthlyDocs AS d LEFT JOIN qrySalesMonthlyCost AS c
     ON (d.SalesYear = c.SalesYear AND d.SalesMonth = c.SalesMonth)
ORDER BY d.SalesYear DESC, d.SalesMonth DESC
```

## SalesByPeriodQuery

فواتير ومرتجعات البيع خلال فترة مع العميل والكاشير

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT s.DocType, s.DocID, s.DocNumber, s.DocDate, s.CustomerID, c.CustomerName,
       e.EmployeeName, s.PaymentType, s.NetAmount, s.VATAmount, s.GrossAmount,
       s.SettledAmount, s.OnAccount
FROM (qrySalesDocuments AS s INNER JOIN Customers AS c ON s.CustomerID = c.CustomerID)
     INNER JOIN Employees AS e ON s.EmployeeID = e.EmployeeID
WHERE s.DocDate >= QDate('PeriodStart') AND s.DocDate < QDate('PeriodEnd')
ORDER BY s.DocDate
```

## SalesByProductQuery

المبيعات حسب المنتج خلال فترة (كمية، صافي، تكلفة، ربح)

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT p.ProductID, p.ProductCode, p.ProductName, c.CategoryName,
       Sum(l.SoldQty) AS QtySold, Sum(l.ReturnedQty) AS QtyReturned, Sum(l.SignedQty) AS NetQty,
       Sum(l.LineNet) AS NetSales, Sum(l.LineVAT) AS SalesVAT, Sum(l.LineGross) AS SalesTotal,
       Sum(l.LineCost) AS CostOfSales, Sum(l.LineNet) - Sum(l.LineCost) AS GrossProfit
FROM (qrySalesLinesInPeriod AS l INNER JOIN Products AS p ON l.ProductID = p.ProductID)
     INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID
GROUP BY p.ProductID, p.ProductCode, p.ProductName, c.CategoryName
ORDER BY p.ProductName
```

## BestSellingProductsQuery

أفضل المنتجات مبيعًا خلال فترة (حسب صافي الكمية)

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT * FROM SalesByProductQuery
ORDER BY NetQty DESC, NetSales DESC
```

## LeastSellingProductsQuery

أقل المنتجات مبيعًا خلال فترة (تشمل المنتجات التي لم تُبع)

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT p.ProductID, p.ProductCode, p.ProductName, c.CategoryName, p.CurrentQuantity,
       CCur(Nz(s.NetQty, 0)) AS NetQtySold, CCur(Nz(s.NetSales, 0)) AS NetSalesAmount
FROM (Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID)
     LEFT JOIN SalesByProductQuery AS s ON p.ProductID = s.ProductID
WHERE p.IsActive = True
ORDER BY CCur(Nz(s.NetQty, 0)), p.ProductName
```

## qryPurchaseDocuments

مستندات الشراء: الفواتير (+) والمرتجعات (−)

```sql
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
FROM PurchaseReturns AS r
```

## PurchasesQuery

فواتير ومرتجعات الشراء خلال فترة مع المورد

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT d.DocType, d.DocID, d.DocNumber, d.SupplierRef, d.DocDate, d.SupplierID,
       s.SupplierName, d.PaymentType, d.NetAmount, d.VATAmount, d.GrossAmount,
       d.SettledAmount, d.OnAccount
FROM qryPurchaseDocuments AS d INNER JOIN Suppliers AS s ON d.SupplierID = s.SupplierID
WHERE d.DocDate >= QDate('PeriodStart') AND d.DocDate < QDate('PeriodEnd')
ORDER BY d.DocDate
```

## qryProductLedger

رصيد كل منتج من دفتر حركة المخزون

```sql
SELECT ProductID, Sum(Quantity) AS LedgerQty, Max(TransactionDate) AS LastMovementDate
FROM InventoryTransactions
GROUP BY ProductID
```

## qryProductLastSale

تاريخ آخر بيع لكل منتج

```sql
SELECT d.ProductID, Max(h.InvoiceDate) AS LastSaleDate
FROM SalesInvoiceDetails AS d INNER JOIN SalesInvoices AS h
     ON d.SalesInvoiceID = h.SalesInvoiceID
GROUP BY d.ProductID
```

## StockBalanceQuery

المخزون الحالي: الكمية والقيمة بالتكلفة وبسعر البيع ومطابقتها مع الحركات

```sql
SELECT p.ProductID, p.ProductCode, p.Barcode, p.ProductName, c.CategoryName, u.UnitName,
       p.ProductLocation, p.CurrentQuantity, p.MinimumQuantity, p.AverageCost, p.SellingPrice,
       p.CurrentQuantity * p.AverageCost AS StockCostValue,
       p.CurrentQuantity * p.SellingPrice AS StockSalesValue,
       CCur(Nz(l.LedgerQty, 0)) AS LedgerQuantity,
       p.CurrentQuantity - CCur(Nz(l.LedgerQty, 0)) AS QuantityMismatch,
       IIf(p.CurrentQuantity <= p.MinimumQuantity, True, False) AS IsLowStock, p.IsActive
FROM ((Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID)
      INNER JOIN Units AS u ON p.UnitID = u.UnitID)
     LEFT JOIN qryProductLedger AS l ON p.ProductID = l.ProductID
ORDER BY p.ProductName
```

## LowStockQuery

المنتجات منخفضة المخزون: CurrentQuantity <= MinimumQuantity

```sql
SELECT p.ProductID, p.ProductCode, p.Barcode, p.ProductName, c.CategoryName,
       p.CurrentQuantity, p.MinimumQuantity, p.MinimumQuantity - p.CurrentQuantity AS ShortageQty,
       s.SupplierName, s.Mobile AS SupplierMobile
FROM (Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID)
     LEFT JOIN Suppliers AS s ON p.SupplierID = s.SupplierID
WHERE p.IsActive = True AND p.CurrentQuantity <= p.MinimumQuantity
ORDER BY p.MinimumQuantity - p.CurrentQuantity DESC, p.ProductName
```

## ProductMovementQuery

حركة منتج خلال فترة مع رصيد أول المدة (الرصيد التراكمي في التقرير)

المعاملات: `PeriodStart`, `PeriodEnd`, `ProductID`

```sql
SELECT 1 AS SortKey, t.TransactionID, t.TransactionDate AS MovementDate,
       tt.TypeName AS MovementType, t.ReferenceNumber,
       IIf(t.Quantity > 0, t.Quantity, 0) AS QtyIn, IIf(t.Quantity < 0, -t.Quantity, 0) AS QtyOut,
       t.Quantity AS NetQty, t.UnitCost, t.Notes
FROM InventoryTransactions AS t INNER JOIN TransactionTypes AS tt
     ON t.TransactionTypeID = tt.TransactionTypeID
WHERE t.ProductID = QLong('ProductID') AND t.TransactionDate >= QDate('PeriodStart') AND t.TransactionDate < QDate('PeriodEnd')
UNION ALL
SELECT 0, 0, QDate('PeriodStart'), 'رصيد أول المدة', Null,
       IIf(CCur(Nz(Sum(o.Quantity), 0)) > 0, CCur(Nz(Sum(o.Quantity), 0)), 0),
       IIf(CCur(Nz(Sum(o.Quantity), 0)) < 0, -CCur(Nz(Sum(o.Quantity), 0)), 0),
       CCur(Nz(Sum(o.Quantity), 0)), Null, Null
FROM InventoryTransactions AS o
WHERE o.ProductID = QLong('ProductID') AND o.TransactionDate < QDate('PeriodStart')
ORDER BY SortKey, MovementDate, TransactionID
```

## SlowMovingProductsQuery

المنتجات غير المتحركة: لها رصيد ولم تُبع منذ عدد الأيام المحدد في الإعدادات

```sql
SELECT p.ProductID, p.ProductCode, p.ProductName, c.CategoryName, p.CurrentQuantity,
       p.AverageCost, p.CurrentQuantity * p.AverageCost AS StockCostValue, ls.LastSaleDate,
       DateDiff('d', Nz(ls.LastSaleDate, p.CreatedAt), Date()) AS DaysWithoutSale
FROM (Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID)
     LEFT JOIN qryProductLastSale AS ls ON p.ProductID = ls.ProductID
WHERE p.IsActive = True AND p.CurrentQuantity > 0
  AND DateDiff('d', Nz(ls.LastSaleDate, p.CreatedAt), Date()) >=
      (SELECT SlowMovingDays FROM Settings)
ORDER BY DateDiff('d', Nz(ls.LastSaleDate, p.CreatedAt), Date()) DESC
```

## StockByCategoryQuery

المخزون حسب التصنيف: عدد المنتجات والكمية والقيمة

```sql
SELECT c.CategoryID, c.CategoryName, Count(*) AS ProductCount,
       Sum(p.CurrentQuantity) AS TotalQuantity,
       Sum(p.CurrentQuantity * p.AverageCost) AS StockCostValue,
       Sum(p.CurrentQuantity * p.SellingPrice) AS StockSalesValue
FROM Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID
WHERE p.IsActive = True
GROUP BY c.CategoryID, c.CategoryName
ORDER BY c.CategoryName
```

## qryCustomerLedger

دفتر حساب العملاء: مدين (عليه) / دائن (له)

```sql
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
WHERE c.OpeningBalance <> 0
```

## qryCustomerLedgerTotals

مجاميع حساب كل عميل

```sql
SELECT CustomerID, Sum(Debit) AS TotalDebit, Sum(Credit) AS TotalCredit,
       Max(EntryDate) AS LastEntryDate
FROM qryCustomerLedger
GROUP BY CustomerID
```

## CustomerBalanceQuery

رصيد كل عميل محسوبًا من الحركات (موجب = عليه للمحل)

```sql
SELECT c.CustomerID, c.CustomerName, c.Mobile, c.CreditLimit, c.AllowCredit, c.IsActive,
       CCur(Nz(l.TotalDebit, 0)) AS DebitTotal, CCur(Nz(l.TotalCredit, 0)) AS CreditTotal,
       CCur(Nz(l.TotalDebit, 0)) - CCur(Nz(l.TotalCredit, 0)) AS Balance,
       c.CurrentBalance AS CachedBalance, l.LastEntryDate
FROM Customers AS c LEFT JOIN qryCustomerLedgerTotals AS l ON c.CustomerID = l.CustomerID
ORDER BY c.CustomerName
```

## CustomersWithDebtQuery

العملاء الذين عليهم مبالغ مستحقة

```sql
SELECT * FROM CustomerBalanceQuery
WHERE Balance > 0
ORDER BY Balance DESC
```

## CustomerStatementQuery

كشف حساب عميل لفترة: رصيد سابق ثم الحركات

المعاملات: `PeriodStart`, `PeriodEnd`, `CustomerID`

```sql
SELECT 1 AS SortKey, l.EntryDate, l.EntryType, l.EntryTypeName, l.DocNumber, l.Debit, l.Credit
FROM qryCustomerLedger AS l
WHERE l.CustomerID = QLong('CustomerID') AND l.EntryDate >= QDate('PeriodStart') AND l.EntryDate < QDate('PeriodEnd')
UNION ALL
SELECT 0, QDate('PeriodStart'), 'BALANCE_FWD', 'رصيد سابق', '-',
       IIf(CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) > 0, CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)), 0),
       IIf(CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) < 0, CCur(Nz(Sum(o.Credit), 0)) - CCur(Nz(Sum(o.Debit), 0)), 0)
FROM qryCustomerLedger AS o
WHERE o.CustomerID = QLong('CustomerID') AND o.EntryDate < QDate('PeriodStart')
ORDER BY SortKey, EntryDate
```

## qrySupplierLedger

دفتر حساب الموردين: دائن (للمورد) / مدين (سُدِّد له)

```sql
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
WHERE s.OpeningBalance <> 0
```

## qrySupplierLedgerTotals

مجاميع حساب كل مورد

```sql
SELECT SupplierID, Sum(Debit) AS TotalDebit, Sum(Credit) AS TotalCredit,
       Max(EntryDate) AS LastEntryDate
FROM qrySupplierLedger
GROUP BY SupplierID
```

## SupplierBalanceQuery

رصيد كل مورد محسوبًا من الحركات (موجب = مستحق للمورد)

```sql
SELECT s.SupplierID, s.SupplierName, s.ContactPerson, s.Mobile, s.IsActive,
       CCur(Nz(l.TotalDebit, 0)) AS DebitTotal, CCur(Nz(l.TotalCredit, 0)) AS CreditTotal,
       CCur(Nz(l.TotalCredit, 0)) - CCur(Nz(l.TotalDebit, 0)) AS Balance,
       s.CurrentBalance AS CachedBalance, l.LastEntryDate
FROM Suppliers AS s LEFT JOIN qrySupplierLedgerTotals AS l ON s.SupplierID = l.SupplierID
ORDER BY s.SupplierName
```

## SupplierStatementQuery

كشف حساب مورد لفترة: رصيد سابق ثم الحركات

المعاملات: `PeriodStart`, `PeriodEnd`, `SupplierID`

```sql
SELECT 1 AS SortKey, l.EntryDate, l.EntryType, l.EntryTypeName, l.DocNumber, l.Debit, l.Credit
FROM qrySupplierLedger AS l
WHERE l.SupplierID = QLong('SupplierID') AND l.EntryDate >= QDate('PeriodStart') AND l.EntryDate < QDate('PeriodEnd')
UNION ALL
SELECT 0, QDate('PeriodStart'), 'BALANCE_FWD', 'رصيد سابق', '-',
       IIf(CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) > 0, CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)), 0),
       IIf(CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) < 0, CCur(Nz(Sum(o.Credit), 0)) - CCur(Nz(Sum(o.Debit), 0)), 0)
FROM qrySupplierLedger AS o
WHERE o.SupplierID = QLong('SupplierID') AND o.EntryDate < QDate('PeriodStart')
ORDER BY SortKey, EntryDate
```

## ExpensesQuery

المصروفات خلال فترة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT e.ExpenseID, e.ExpenseNumber, e.ExpenseDate, t.ExpenseTypeName, e.Amount, e.Tax,
       e.TotalAmount, pm.MethodName, e.Description, em.EmployeeName
FROM ((Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID)
      INNER JOIN Employees AS em ON e.EmployeeID = em.EmployeeID)
     LEFT JOIN PaymentMethods AS pm ON e.PaymentMethodID = pm.PaymentMethodID
WHERE e.ExpenseDate >= QDate('PeriodStart') AND e.ExpenseDate < QDate('PeriodEnd')
ORDER BY e.ExpenseDate
```

## ExpensesByTypeQuery

المصروفات مجمّعة حسب النوع خلال فترة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT t.ExpenseTypeName, Count(*) AS ExpenseCount, Sum(e.Amount) AS AmountExVAT,
       Sum(e.Tax) AS InputVAT, Sum(e.TotalAmount) AS AmountTotal
FROM Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID
WHERE e.ExpenseDate >= QDate('PeriodStart') AND e.ExpenseDate < QDate('PeriodEnd')
GROUP BY t.ExpenseTypeName
ORDER BY Sum(e.TotalAmount) DESC
```

## qryProfitSales

صافي المبيعات وتكلفتها خلال الفترة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT CCur(Nz(Sum(LineNet), 0)) AS PeriodNetSales, CCur(Nz(Sum(LineCost), 0)) AS PeriodCost
FROM qrySalesLinesInPeriod
```

## qryProfitAdjustments

قيمة فروقات المخزون (جرد، إضافة، خصم) خلال الفترة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT CCur(Nz(Sum(t.Quantity * t.UnitCost), 0)) AS PeriodAdjustments
FROM InventoryTransactions AS t INNER JOIN TransactionTypes AS tt
     ON t.TransactionTypeID = tt.TransactionTypeID
WHERE tt.TypeCode IN ('STOCK_IN', 'STOCK_OUT', 'ADJUSTMENT') AND t.TransactionDate >= QDate('PeriodStart') AND t.TransactionDate < QDate('PeriodEnd')
```

## qryProfitExpenses

المصروفات (بدون ضريبة) خلال الفترة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT CCur(Nz(Sum(Amount), 0)) AS PeriodExpenses
FROM Expenses
WHERE ExpenseDate >= QDate('PeriodStart') AND ExpenseDate < QDate('PeriodEnd')
```

## ProfitQuery

الأرباح: صافي المبيعات − التكلفة = مجمل الربح؛ ثم ± فروقات المخزون − المصروفات = صافي الربح

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT QDate('PeriodStart') AS PeriodFrom, DateAdd('d', -1, QDate('PeriodEnd')) AS PeriodTo,
       s.PeriodNetSales AS NetSales, s.PeriodCost AS CostOfSales,
       s.PeriodNetSales - s.PeriodCost AS GrossProfit,
       a.PeriodAdjustments AS InventoryAdjustments, x.PeriodExpenses AS TotalExpenses,
       s.PeriodNetSales - s.PeriodCost + a.PeriodAdjustments - x.PeriodExpenses AS NetProfit
FROM qryProfitSales AS s, qryProfitAdjustments AS a, qryProfitExpenses AS x
```

## qryVatOutput

ضريبة المخرجات (المبيعات ناقص المرتجعات)

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT CCur(Nz(Sum(LineNet), 0)) AS TaxableSales, CCur(Nz(Sum(LineVAT), 0)) AS OutputVAT
FROM qrySalesLinesInPeriod
```

## qryVatInputPurchases

ضريبة المدخلات من المشتريات (ناقص المرتجعات)

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT CCur(Nz(Sum(NetAmount), 0)) AS TaxablePurchases, CCur(Nz(Sum(VATAmount), 0)) AS PurchaseVAT
FROM qryPurchaseDocuments
WHERE DocDate >= QDate('PeriodStart') AND DocDate < QDate('PeriodEnd')
```

## qryVatInputExpenses

ضريبة المدخلات من المصروفات

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT CCur(Nz(Sum(Tax), 0)) AS ExpenseVAT
FROM Expenses
WHERE ExpenseDate >= QDate('PeriodStart') AND ExpenseDate < QDate('PeriodEnd')
```

## VatSummaryQuery

ملخص ضريبة القيمة المضافة للفترة (للإقرار الضريبي)

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT QDate('PeriodStart') AS PeriodFrom, DateAdd('d', -1, QDate('PeriodEnd')) AS PeriodTo,
       o.TaxableSales, o.OutputVAT, p.TaxablePurchases, p.PurchaseVAT, e.ExpenseVAT,
       p.PurchaseVAT + e.ExpenseVAT AS InputVAT,
       o.OutputVAT - p.PurchaseVAT - e.ExpenseVAT AS NetVATDue
FROM qryVatOutput AS o, qryVatInputPurchases AS p, qryVatInputExpenses AS e
```

## qrySalesInvoiceLineTotals

مجموع أسطر كل فاتورة بيع

```sql
SELECT SalesInvoiceID, Sum(LineTotal) AS LinesTotal
FROM SalesInvoiceDetails
GROUP BY SalesInvoiceID
```

## qryPurchaseInvoiceLineTotals

مجموع أسطر كل فاتورة شراء

```sql
SELECT PurchaseInvoiceID, Sum(LineTotal) AS LinesTotal
FROM PurchaseInvoiceDetails
GROUP BY PurchaseInvoiceID
```

## qrySalesReturnedQty

الكمية المرتجعة من كل سطر فاتورة بيع

```sql
SELECT SalesDetailID, Sum(Quantity) AS QtyReturned
FROM SalesReturnDetails
GROUP BY SalesDetailID
```

## qryPurchaseReturnedQty

الكمية المرتجعة للمورد من كل سطر فاتورة شراء

```sql
SELECT PurchaseDetailID, Sum(Quantity) AS QtyReturned
FROM PurchaseReturnDetails
GROUP BY PurchaseDetailID
```

## IntegrityCheckQuery

فحص سلامة البيانات: أي سطر هنا مشكلة يجب مراجعتها (النتيجة الفارغة = سليم)

```sql
SELECT 'STOCK_MISMATCH' AS IssueCode, 'الكمية المسجلة لا تطابق حركات المخزون' AS IssueText,
       'Products' AS SourceTable, p.ProductID AS RecordID,
       CCur(Nz(l.LedgerQty, 0)) AS ExpectedValue, CCur(p.CurrentQuantity) AS ActualValue
FROM Products AS p LEFT JOIN qryProductLedger AS l ON p.ProductID = l.ProductID
WHERE p.CurrentQuantity <> CCur(Nz(l.LedgerQty, 0))
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
       h.SalesInvoiceID, CCur(Nz(t.LinesTotal, 0)), CCur(h.TotalAmount)
FROM SalesInvoices AS h LEFT JOIN qrySalesInvoiceLineTotals AS t
     ON h.SalesInvoiceID = t.SalesInvoiceID
WHERE h.TotalAmount <> CCur(Nz(t.LinesTotal, 0)) OR t.SalesInvoiceID Is Null
UNION ALL
SELECT 'PURCHASE_TOTAL', 'إجمالي فاتورة الشراء لا يساوي مجموع أسطرها', 'PurchaseInvoices',
       h.PurchaseInvoiceID, CCur(Nz(t.LinesTotal, 0)), CCur(h.TotalAmount)
FROM PurchaseInvoices AS h LEFT JOIN qryPurchaseInvoiceLineTotals AS t
     ON h.PurchaseInvoiceID = t.PurchaseInvoiceID
WHERE h.TotalAmount <> CCur(Nz(t.LinesTotal, 0)) OR t.PurchaseInvoiceID Is Null
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
ORDER BY IssueCode, RecordID
```
