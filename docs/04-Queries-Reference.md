# مرجع الاستعلامات (Queries Reference)

> ملف مُولَّد تلقائيًا من `tools/queries.py` – لا تعدّله يدويًا.

عدد الاستعلامات: **60**. الاستعلامات التي تبدأ بـ `qry` مساعدة تستخدمها الاستعلامات الأخرى؛ البقية تُستخدم مباشرة في التقارير والنماذج. ⭐ = مطلوب بالاسم في البرومبت.

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
| 11 | [`SalesByCategoryQuery`](#salesbycategoryquery) | المبيعات حسب التصنيف خلال فترة (للرسم الدائري) | `PeriodStart`, `PeriodEnd` |
| 12 | [`LeastSellingProductsQuery`](#leastsellingproductsquery) | أقل المنتجات مبيعًا خلال فترة (تشمل المنتجات التي لم تُبع) | `PeriodStart`, `PeriodEnd` |
| 13 | [`qryPurchaseDocuments`](#qrypurchasedocuments) | مستندات الشراء: الفواتير (+) والمرتجعات (−) |  |
| 14 | [`PurchasesQuery`](#purchasesquery) | فواتير ومرتجعات الشراء خلال فترة مع المورد | `PeriodStart`, `PeriodEnd` |
| 15 | [`qryProductLedger`](#qryproductledger) | رصيد كل منتج من دفتر حركة المخزون |  |
| 16 | [`qryProductLastSale`](#qryproductlastsale) | تاريخ آخر بيع لكل منتج |  |
| 17 | [`StockBalanceQuery`](#stockbalancequery) ⭐ | المخزون الحالي: الكمية والقيمة بالتكلفة وبسعر البيع ومطابقتها مع الحركات |  |
| 18 | [`LowStockQuery`](#lowstockquery) ⭐ | المنتجات منخفضة المخزون: CurrentQuantity <= MinimumQuantity |  |
| 19 | [`ProductMovementQuery`](#productmovementquery) | حركة منتج خلال فترة مع رصيد أول المدة (الرصيد التراكمي في التقرير) | `PeriodStart`, `PeriodEnd`, `ProductID` |
| 20 | [`SlowMovingProductsQuery`](#slowmovingproductsquery) ⭐ | المنتجات غير المتحركة: لها رصيد ولم تُبع منذ عدد الأيام المحدد في الإعدادات |  |
| 21 | [`StockByCategoryQuery`](#stockbycategoryquery) | المخزون حسب التصنيف: عدد المنتجات والكمية والقيمة |  |
| 22 | [`StockCountQuery`](#stockcountquery) | تفاصيل جلسات الجرد: الكمية المسجلة والفعلية والفرق وقيمته |  |
| 23 | [`qryCustomerLedger`](#qrycustomerledger) | دفتر حساب العملاء: مدين (عليه) / دائن (له) |  |
| 24 | [`qryCustomerLedgerTotals`](#qrycustomerledgertotals) | مجاميع حساب كل عميل |  |
| 25 | [`CustomerBalanceQuery`](#customerbalancequery) ⭐ | رصيد كل عميل محسوبًا من الحركات (موجب = عليه للمحل) |  |
| 26 | [`CustomersWithDebtQuery`](#customerswithdebtquery) | العملاء الذين عليهم مبالغ مستحقة |  |
| 27 | [`CustomerStatementQuery`](#customerstatementquery) | كشف حساب عميل لفترة: رصيد سابق ثم الحركات | `PeriodStart`, `PeriodEnd`, `CustomerID` |
| 28 | [`qrySupplierLedger`](#qrysupplierledger) | دفتر حساب الموردين: دائن (للمورد) / مدين (سُدِّد له) |  |
| 29 | [`qrySupplierLedgerTotals`](#qrysupplierledgertotals) | مجاميع حساب كل مورد |  |
| 30 | [`SupplierBalanceQuery`](#supplierbalancequery) ⭐ | رصيد كل مورد محسوبًا من الحركات (موجب = مستحق للمورد) |  |
| 31 | [`SupplierStatementQuery`](#supplierstatementquery) | كشف حساب مورد لفترة: رصيد سابق ثم الحركات | `PeriodStart`, `PeriodEnd`, `SupplierID` |
| 32 | [`ExpensesQuery`](#expensesquery) | المصروفات خلال فترة | `PeriodStart`, `PeriodEnd` |
| 33 | [`ExpensesByTypeQuery`](#expensesbytypequery) | المصروفات مجمّعة حسب النوع خلال فترة | `PeriodStart`, `PeriodEnd` |
| 34 | [`qryProfitSales`](#qryprofitsales) | صافي المبيعات وتكلفتها خلال الفترة | `PeriodStart`, `PeriodEnd` |
| 35 | [`qryProfitAdjustments`](#qryprofitadjustments) | قيمة فروقات المخزون (جرد، إضافة، خصم) خلال الفترة | `PeriodStart`, `PeriodEnd` |
| 36 | [`qryProfitExpenses`](#qryprofitexpenses) | المصروفات (بدون ضريبة) خلال الفترة | `PeriodStart`, `PeriodEnd` |
| 37 | [`ProfitQuery`](#profitquery) ⭐ | الأرباح: صافي المبيعات − التكلفة = مجمل الربح؛ ثم ± فروقات المخزون − المصروفات = صافي الربح | `PeriodStart`, `PeriodEnd` |
| 38 | [`qryVatOutput`](#qryvatoutput) | ضريبة المخرجات (المبيعات ناقص المرتجعات) | `PeriodStart`, `PeriodEnd` |
| 39 | [`qryVatInputPurchases`](#qryvatinputpurchases) | ضريبة المدخلات من المشتريات (ناقص المرتجعات) | `PeriodStart`, `PeriodEnd` |
| 40 | [`qryVatInputExpenses`](#qryvatinputexpenses) | ضريبة المدخلات من المصروفات | `PeriodStart`, `PeriodEnd` |
| 41 | [`VatSummaryQuery`](#vatsummaryquery) | ملخص ضريبة القيمة المضافة للفترة (للإقرار الضريبي) | `PeriodStart`, `PeriodEnd` |
| 42 | [`DashboardQuery`](#dashboardquery) | مؤشرات لوحة التحكم في سجل واحد (اليوم، الشهر، الأرصدة، المخزون) | `DashDay`, `DashMonth`, `DashEnd` |
| 43 | [`qryDashboardTopProducts`](#qrydashboardtopproducts) | صافي الكمية المباعة لكل منتج منذ بداية الشهر (لوحة التحكم) | `DashMonth`, `DashEnd` |
| 44 | [`qrySalesDocPrint`](#qrysalesdocprint) | بيانات طباعة فواتير البيع والإشعارات الدائنة (سطر لكل صنف) |  |
| 45 | [`qryPurchaseDocPrint`](#qrypurchasedocprint) | بيانات طباعة فواتير الشراء ومرتجعاتها (سطر لكل صنف) |  |
| 46 | [`qryVoucherPrint`](#qryvoucherprint) | بيانات طباعة سندات القبض (من العملاء) وسندات الصرف (للموردين) |  |
| 47 | [`qryCashMovements`](#qrycashmovements) | كل حركات النقدية في الخزينة والصناديق: داخل (+) وخارج (−) |  |
| 48 | [`qryCashBoxTotals`](#qrycashboxtotals) | إجمالي الداخل والخارج لكل صندوق |  |
| 49 | [`CashBoxBalanceQuery`](#cashboxbalancequery) | أرصدة الخزينة والصناديق الآن |  |
| 50 | [`CashStatementQuery`](#cashstatementquery) | حركة الخزينة / الصندوق لفترة: رصيد أول المدة ثم الحركات (0 = كل الصناديق) | `PeriodStart`, `PeriodEnd`, `CashBoxID` |
| 51 | [`qryCashDays`](#qrycashdays) | مقبوضات ومدفوعات كل يوم داخل الفترة | `PeriodStart`, `PeriodEnd`, `CashBoxID` |
| 52 | [`CashDailyQuery`](#cashdailyquery) | حركة الخزينة اليومية: رصيد أول اليوم والمقبوضات والمدفوعات ورصيد آخر اليوم | `PeriodStart`, `PeriodEnd`, `CashBoxID` |
| 53 | [`CashClosingsQuery`](#cashclosingsquery) | تصفيات يومية الكاشير خلال فترة (0 = كل الصناديق) | `PeriodStart`, `PeriodEnd`, `CashBoxID` |
| 54 | [`qryCashClosingPrint`](#qrycashclosingprint) | بيانات طباعة تصفية الكاشير |  |
| 55 | [`qryCashVoucherPrint`](#qrycashvoucherprint) | بيانات طباعة سندات قبض وصرف وتحويل النقدية |  |
| 56 | [`qrySalesInvoiceLineTotals`](#qrysalesinvoicelinetotals) | مجموع أسطر كل فاتورة بيع |  |
| 57 | [`qryPurchaseInvoiceLineTotals`](#qrypurchaseinvoicelinetotals) | مجموع أسطر كل فاتورة شراء |  |
| 58 | [`qrySalesReturnedQty`](#qrysalesreturnedqty) | الكمية المرتجعة من كل سطر فاتورة بيع |  |
| 59 | [`qryPurchaseReturnedQty`](#qrypurchasereturnedqty) | الكمية المرتجعة للمورد من كل سطر فاتورة شراء |  |
| 60 | [`IntegrityCheckQuery`](#integritycheckquery) | فحص سلامة البيانات: أي سطر هنا مشكلة يجب مراجعتها (النتيجة الفارغة = سليم) |  |

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
| 8 | طباعة فاتورة الشراء: سطران | `SELECT COUNT(*) FROM qryPurchaseDocPrint WHERE DocKind = 'PURCHASE' AND DocID = {ref:PUR1}` | 2 |
| 9 | طباعة مرتجع الشراء: رقم السطر الأصلي ورقم الفاتورة الأصلية | `SELECT COUNT(*) FROM qryPurchaseDocPrint WHERE DocKind = 'RETURN' AND DocID = {ref:PRT1} AND LineNumber = 1 AND OriginalNumber = 'TEST-PUR-1' AND RemainingAmount = 345` | 1 |
| 10 | طباعة سند الصرف: المبلغ وطريقة الدفع | `SELECT Amount FROM qryVoucherPrint WHERE DocKind = 'PAYMENT' AND DocID = {ref:PAY1}` | 1000 |
| 11 | طباعة سند القبض | `SELECT COUNT(*) FROM qryVoucherPrint WHERE DocKind = 'RECEIPT' AND DocID = {ref:RCV1}` | 1 |
| 12 | الجرد المفتوح: عجز المنتج 1 = −3 × 60 | `SELECT DifferenceValue FROM StockCountQuery WHERE ProductID = {ref:P1}` | -180 |
| 13 | الجرد المفتوح: صنف واحد لم يُعدّ بعد | `SELECT COUNT(*) FROM StockCountQuery WHERE ActualQuantity Is Null` | 1 |
| 14 | غير المتحركة: منتج واحد | `SELECT COUNT(*) FROM SlowMovingProductsQuery` | 1 |
| 15 | غير المتحركة: المنتج 3 بقيمة 200 | `SELECT StockCostValue FROM SlowMovingProductsQuery WHERE ProductID = {ref:P3}` | 200 |
| 16 | المخزون حسب التصنيف: تصنيف عام = 83 + 20 | `SELECT TotalQuantity FROM StockByCategoryQuery WHERE CategoryID = 1` | 103 |
| 17 | المخزون حسب التصنيف: قيمة تصنيف عام = 4980 + 200 | `SELECT StockCostValue FROM StockByCategoryQuery WHERE CategoryID = 1` | 5180 |
| 18 | المخزون حسب التصنيف: قيمة التصنيف 2 = 186 × 10 | `SELECT StockCostValue FROM StockByCategoryQuery WHERE CategoryID = {ref:CAT2}` | 1860 |
| 19 | حركة المنتج 2: رصيد أول المدة = 195 | `SELECT NetQty FROM ProductMovementQuery WHERE SortKey = 0` | 195 |
| 20 | حركة المنتج 2: 3 حركات + سطر الرصيد السابق | `SELECT COUNT(*) FROM ProductMovementQuery` | 4 |
| 21 | حركة المنتج 2: الرصيد الختامي = 186 | `SELECT Sum(NetQty) FROM ProductMovementQuery` | 186 |
| 22 | المبيعات اليومية: يوم الفاتورة الآجلة = 460 | `SELECT NetSalesTotal FROM DailySalesQuery WHERE SaleDate = DateValue({day:5})` | 460 |
| 23 | المبيعات اليومية: الآجل في نفس اليوم = 460 | `SELECT CreditSales FROM DailySalesQuery WHERE SaleDate = DateValue({day:5})` | 460 |
| 24 | المبيعات اليومية: يوم الفاتورة النقدية = 1150 نقدًا | `SELECT CashSales FROM DailySalesQuery WHERE SaleDate = DateValue({day:10})` | 1150 |
| 25 | المبيعات اليومية: يوم المرتجع = −46 | `SELECT NetSalesTotal FROM DailySalesQuery WHERE SaleDate = DateValue({day:2})` | -46 |
| 26 | المبيعات الشهرية: صافي كل الأشهر = 100 + 1000 + 400 − 40 | `SELECT Sum(NetSalesExVAT) FROM MonthlySalesQuery` | 1460 |
| 27 | المبيعات الشهرية: مجمل ربح كل الأشهر = 1460 − 850 | `SELECT Sum(GrossProfit) FROM MonthlySalesQuery` | 610 |
| 28 | المبيعات خلال الفترة: فاتورتان ومرتجع | `SELECT COUNT(*) FROM SalesByPeriodQuery` | 3 |
| 29 | المبيعات خلال الفترة: الإجمالي = 1150 + 460 − 46 | `SELECT Sum(GrossAmount) FROM SalesByPeriodQuery` | 1564 |
| 30 | المبيعات حسب المنتج: المنتج 1 صافي كمية 12 | `SELECT NetQty FROM SalesByProductQuery WHERE ProductID = {ref:P1}` | 12 |
| 31 | المبيعات حسب المنتج: ربح المنتج 1 = 1200 − 720 | `SELECT GrossProfit FROM SalesByProductQuery WHERE ProductID = {ref:P1}` | 480 |
| 32 | المبيعات حسب المنتج: المنتج 2 مرتجع 2 | `SELECT QtyReturned FROM SalesByProductQuery WHERE ProductID = {ref:P2}` | 2 |
| 33 | المبيعات حسب المنتج: صافي مبيعات المنتج 2 = 200 − 40 | `SELECT NetSales FROM SalesByProductQuery WHERE ProductID = {ref:P2}` | 160 |
| 34 | المبيعات حسب التصنيف: المجموع = المبيعات حسب المنتج | `SELECT (SELECT Sum(CategoryTotal) FROM SalesByCategoryQuery) - (SELECT Sum(SalesTotal) FROM SalesByProductQuery) FROM Settings` | 0 |
| 35 | المبيعات حسب التصنيف: صافي الكمية = 12 + 8 | `SELECT Sum(CategoryQty) FROM SalesByCategoryQuery` | 20 |
| 36 | الأكثر مبيعًا: المنتج 1 | `SELECT TOP 1 ProductID FROM BestSellingProductsQuery ORDER BY NetQty DESC, NetSales DESC` | رقم P1 |
| 37 | الأقل مبيعًا: المنتج 3 (لم يُبع) | `SELECT TOP 1 ProductID FROM LeastSellingProductsQuery ORDER BY NetQtySold, ProductName` | رقم P3 |
| 38 | المشتريات خلال الفترة: مرتجع الشراء فقط | `SELECT COUNT(*) FROM PurchasesQuery` | 1 |
| 39 | المشتريات خلال الفترة: −345 | `SELECT Sum(GrossAmount) FROM PurchasesQuery` | -345 |
| 40 | رصيد العميل الآجل = 50 + 460 − 100 − 46 − 200 | `SELECT Balance FROM CustomerBalanceQuery WHERE CustomerID = {ref:C2}` | 164 |
| 41 | رصيد العميل النقدي = 0 | `SELECT Balance FROM CustomerBalanceQuery WHERE CustomerID = 1` | 0 |
| 42 | العملاء المدينون: عميل واحد | `SELECT COUNT(*) FROM CustomersWithDebtQuery` | 1 |
| 43 | كشف العميل: الرصيد السابق = 50 | `SELECT Debit FROM CustomerStatementQuery WHERE SortKey = 0` | 50 |
| 44 | كشف العميل: رصيد سابق + 3 حركات | `SELECT COUNT(*) FROM CustomerStatementQuery` | 4 |
| 45 | كشف العميل: الرصيد الختامي = 164 | `SELECT Sum(Debit) - Sum(Credit) FROM CustomerStatementQuery` | 164 |
| 46 | رصيد المورد = 9200 − 5000 − 1000 − 345 | `SELECT Balance FROM SupplierBalanceQuery WHERE SupplierID = {ref:S1}` | 2855 |
| 47 | كشف المورد: الرصيد السابق = 4200 | `SELECT Credit FROM SupplierStatementQuery WHERE SortKey = 0` | 4200 |
| 48 | كشف المورد: الرصيد الختامي = 2855 | `SELECT Sum(Credit) - Sum(Debit) FROM SupplierStatementQuery` | 2855 |
| 49 | المصروفات خلال الفترة: مصروف واحد | `SELECT COUNT(*) FROM ExpensesQuery` | 1 |
| 50 | المصروفات خلال الفترة: 230 شامل الضريبة | `SELECT Sum(TotalAmount) FROM ExpensesQuery` | 230 |
| 51 | المصروفات حسب النوع: الكهرباء 200 | `SELECT AmountExVAT FROM ExpensesByTypeQuery` | 200 |
| 52 | الأرباح: صافي المبيعات = 1000 + 400 − 40 | `SELECT NetSales FROM ProfitQuery` | 1360 |
| 53 | الأرباح: تكلفة المبيعات = 600 + 220 − 20 | `SELECT CostOfSales FROM ProfitQuery` | 800 |
| 54 | الأرباح: مجمل الربح = 1360 − 800 | `SELECT GrossProfit FROM ProfitQuery` | 560 |
| 55 | الأرباح: فروقات المخزون = −10 (منتج تالف) | `SELECT InventoryAdjustments FROM ProfitQuery` | -10 |
| 56 | الأرباح: المصروفات = 200 | `SELECT TotalExpenses FROM ProfitQuery` | 200 |
| 57 | الأرباح: صافي الربح = 560 − 10 − 200 | `SELECT NetProfit FROM ProfitQuery` | 350 |
| 58 | الضريبة: ضريبة المخرجات = 150 + 60 − 6 | `SELECT OutputVAT FROM VatSummaryQuery` | 204 |
| 59 | الضريبة: ضريبة المدخلات = −45 (مرتجع شراء) + 30 (مصروف) | `SELECT InputVAT FROM VatSummaryQuery` | -15 |
| 60 | الضريبة: الصافي المستحق = 204 + 15 | `SELECT NetVATDue FROM VatSummaryQuery` | 219 |
| 61 | طباعة الفاتورة الآجلة: سطران | `SELECT COUNT(*) FROM qrySalesDocPrint WHERE DocKind = 'SALE' AND DocID = {ref:INV2}` | 2 |
| 62 | طباعة الفاتورة الآجلة: مجموع الأسطر = 460 | `SELECT Sum(LineTotal) FROM qrySalesDocPrint WHERE DocKind = 'SALE' AND DocID = {ref:INV2}` | 460 |
| 63 | طباعة الإشعار الدائن: سطر واحد بقيمة 46 | `SELECT Sum(LineTotal) FROM qrySalesDocPrint WHERE DocKind = 'RETURN' AND DocID = {ref:CRN1}` | 46 |
| 64 | رصيد صندوق الكاشير = 500 + 115 + 1150 + 100 + 200 − 230 − 50 − 15 − 1000 | `SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = {ref:BOXC}` | 770 |
| 65 | رصيد الخزينة = 10000 − 5000 + 1000 + 2000 − 300 | `SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = {ref:BOXM}` | 7700 |
| 66 | الصناديق المسجلة بدون حركة رصيدها صفر | `SELECT Sum(Balance) FROM CashBoxBalanceQuery WHERE CashBoxID <= 2` | 0 |
| 67 | حركة صندوق الكاشير: رصيد أول المدة = 500 + 115 − 50 | `SELECT AmountIn FROM CashStatementQuery WHERE SortKey = 0` | 565 |
| 68 | حركة صندوق الكاشير: 6 حركات في الفترة | `SELECT COUNT(*) FROM CashStatementQuery WHERE SortKey = 1` | 6 |
| 69 | حركة صندوق الكاشير: رصيد آخر المدة = 770 | `SELECT Sum(AmountIn) - Sum(AmountOut) FROM CashStatementQuery` | 770 |
| 70 | حركة الخزينة: التحويل من الكاشير داخل = 1000 | `SELECT AmountIn FROM CashStatementQuery WHERE MoveType = 'TRANSFER_IN'` | 1000 |
| 71 | يومية صندوق الكاشير يوم التصفية: رصيد أول اليوم = 565 + 1150 + 100 − 230 | `SELECT OpeningBalance FROM CashDailyQuery WHERE CashDay = DateValue({day:4})` | 1585 |
| 72 | يومية صندوق الكاشير يوم التصفية: المدفوعات = 15 عجز + 1000 تحويل | `SELECT Payments FROM CashDailyQuery WHERE CashDay = DateValue({day:4})` | 1015 |
| 73 | يومية صندوق الكاشير يوم التصفية: رصيد آخر اليوم = 570 | `SELECT ClosingBalance FROM CashDailyQuery WHERE CashDay = DateValue({day:4})` | 570 |
| 74 | يومية صندوق الكاشير: آخر يوم = الرصيد الحالي | `SELECT ClosingBalance FROM CashDailyQuery WHERE CashDay = DateValue({day:3})` | 770 |
| 75 | تصفيات الكاشير خلال الفترة: تصفية واحدة بعجز 15 | `SELECT Difference FROM CashClosingsQuery` | -15 |
| 76 | طباعة سند صرف المصروف: نوع المصروف | `SELECT COUNT(*) FROM qryCashVoucherPrint WHERE DocID = {ref:V1} AND ExpenseTypeName = 'مصروفات أخرى'` | 1 |
| 77 | طباعة سند التحويل: الصندوق المستلم | `SELECT COUNT(*) FROM qryCashVoucherPrint WHERE DocID = {ref:V3} AND ToBoxName = 'TEST الخزينة'` | 1 |
| 78 | فحص السلامة: لا توجد مشكلات | `SELECT COUNT(*) FROM IntegrityCheckQuery` | 0 |

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

## SalesByCategoryQuery

المبيعات حسب التصنيف خلال فترة (للرسم الدائري)

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT CategoryName, Count(*) AS ProductCount, Sum(NetQty) AS CategoryQty,
       Sum(NetSales) AS CategoryNet, Sum(SalesTotal) AS CategoryTotal, Sum(GrossProfit) AS CategoryProfit
FROM SalesByProductQuery
GROUP BY CategoryName
ORDER BY Sum(SalesTotal) DESC
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
WHERE p.IsActive = True AND p.TrackStock = True AND p.CurrentQuantity <= p.MinimumQuantity
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

## StockCountQuery

تفاصيل جلسات الجرد: الكمية المسجلة والفعلية والفرق وقيمته

```sql
SELECT c.StockCountID, c.CountNumber, c.CountDate, c.Status, c.CategoryID, g.CategoryName,
       d.ProductID, p.ProductCode, p.ProductName, d.SystemQuantity, d.ActualQuantity, d.Difference,
       d.UnitCost, d.DifferenceValue, d.Notes
FROM ((StockCountDetails AS d INNER JOIN StockCounts AS c ON d.StockCountID = c.StockCountID)
      INNER JOIN Products AS p ON d.ProductID = p.ProductID)
     LEFT JOIN Categories AS g ON c.CategoryID = g.CategoryID
ORDER BY c.StockCountID, p.ProductName
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
       e.TotalAmount, pm.MethodName, e.Description, em.EmployeeName, e.ExpenseTypeID
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

## DashboardQuery

مؤشرات لوحة التحكم في سجل واحد (اليوم، الشهر، الأرصدة، المخزون)

المعاملات: `DashDay`, `DashMonth`, `DashEnd`

```sql
SELECT (SELECT CCur(Nz(Sum(d.GrossAmount), 0)) FROM qrySalesDocuments AS d
        WHERE d.DocDate >= QDate('DashDay') AND d.DocDate < QDate('DashEnd')) AS TodaySales,
       (SELECT Count(*) FROM SalesInvoices AS h
        WHERE h.InvoiceDate >= QDate('DashDay') AND h.InvoiceDate < QDate('DashEnd')) AS TodayInvoices,
       (SELECT CCur(Nz(Sum(d.GrossAmount), 0)) FROM qrySalesDocuments AS d
        WHERE d.DocDate >= QDate('DashMonth') AND d.DocDate < QDate('DashEnd')) AS MonthSales,
       (SELECT CCur(Nz(Sum(d.VATAmount), 0)) FROM qrySalesDocuments AS d
        WHERE d.DocDate >= QDate('DashMonth') AND d.DocDate < QDate('DashEnd')) AS MonthVAT,
       (SELECT Count(*) FROM SalesInvoices AS h
        WHERE h.InvoiceDate >= QDate('DashMonth') AND h.InvoiceDate < QDate('DashEnd')) AS MonthInvoices,
       (SELECT CCur(Nz(Sum(e.Amount), 0)) FROM Expenses AS e
        WHERE e.ExpenseDate >= QDate('DashMonth') AND e.ExpenseDate < QDate('DashEnd')) AS MonthExpenses,
       (SELECT CCur(Nz(Sum(c.CurrentBalance), 0)) FROM Customers AS c WHERE c.CurrentBalance > 0) AS CustomerDebt,
       (SELECT Count(*) FROM Customers AS c WHERE c.CurrentBalance > 0) AS DebtorCount,
       (SELECT CCur(Nz(Sum(s.CurrentBalance), 0)) FROM Suppliers AS s WHERE s.CurrentBalance > 0) AS SupplierDue,
       (SELECT CCur(Nz(Sum(p.CurrentQuantity * p.AverageCost), 0)) FROM Products AS p
        WHERE p.IsActive = True AND p.CurrentQuantity > 0) AS StockValue,
       (SELECT Count(*) FROM LowStockQuery) AS LowStockCount
FROM Settings AS st
WHERE st.SettingID = 1
```

## qryDashboardTopProducts

صافي الكمية المباعة لكل منتج منذ بداية الشهر (لوحة التحكم)

المعاملات: `DashMonth`, `DashEnd`

```sql
SELECT l.ProductID, p.ProductName, Sum(l.SignedQty) AS NetQty, Sum(l.LineGross) AS NetSales
FROM qrySalesLineItems AS l INNER JOIN Products AS p ON l.ProductID = p.ProductID
WHERE l.DocDate >= QDate('DashMonth') AND l.DocDate < QDate('DashEnd')
GROUP BY l.ProductID, p.ProductName
HAVING Sum(l.SignedQty) > 0
```

## qrySalesDocPrint

بيانات طباعة فواتير البيع والإشعارات الدائنة (سطر لكل صنف)

```sql
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
    INNER JOIN Employees AS e ON r.EmployeeID = e.EmployeeID
```

## qryPurchaseDocPrint

بيانات طباعة فواتير الشراء ومرتجعاتها (سطر لكل صنف)

```sql
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
    INNER JOIN Employees AS e ON r.EmployeeID = e.EmployeeID
```

## qryVoucherPrint

بيانات طباعة سندات القبض (من العملاء) وسندات الصرف (للموردين)

```sql
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
     INNER JOIN Employees AS e ON p.EmployeeID = e.EmployeeID
```

## qryCashMovements

كل حركات النقدية في الخزينة والصناديق: داخل (+) وخارج (−)

```sql
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
WHERE b.OpeningBalance <> 0
```

## qryCashBoxTotals

إجمالي الداخل والخارج لكل صندوق

```sql
SELECT CashBoxID, Sum(AmountIn) AS BoxIn, Sum(AmountOut) AS BoxOut, Max(MoveDate) AS LastMoveDate
FROM qryCashMovements
GROUP BY CashBoxID
```

## CashBoxBalanceQuery

أرصدة الخزينة والصناديق الآن

```sql
SELECT b.CashBoxID, b.BoxName, b.BoxType,
       IIf(b.BoxType = 'MAIN', 'خزينة رئيسية', 'صندوق كاشير') AS BoxTypeName, b.IsActive,
       CCur(Nz(t.BoxIn, 0)) AS TotalIn, CCur(Nz(t.BoxOut, 0)) AS TotalOut,
       CCur(Nz(t.BoxIn, 0)) - CCur(Nz(t.BoxOut, 0)) AS Balance, t.LastMoveDate
FROM CashBoxes AS b LEFT JOIN qryCashBoxTotals AS t ON b.CashBoxID = t.CashBoxID
ORDER BY b.BoxType DESC, b.BoxName
```

## CashStatementQuery

حركة الخزينة / الصندوق لفترة: رصيد أول المدة ثم الحركات (0 = كل الصناديق)

المعاملات: `PeriodStart`, `PeriodEnd`, `CashBoxID`

```sql
SELECT 1 AS SortKey, m.MoveDate, m.MoveType, m.MoveTypeName, m.DocNumber, m.PartyName, m.Details,
       b.BoxName, m.AmountIn, m.AmountOut, m.CashBoxID
FROM qryCashMovements AS m INNER JOIN CashBoxes AS b ON m.CashBoxID = b.CashBoxID
WHERE (QLong('CashBoxID') = 0 OR m.CashBoxID = QLong('CashBoxID')) AND m.MoveDate >= QDate('PeriodStart') AND m.MoveDate < QDate('PeriodEnd')
UNION ALL
SELECT 0, QDate('PeriodStart'), 'BALANCE_FWD', 'رصيد أول المدة', '-', Null, Null, Null,
       IIf(CCur(Nz(Sum(o.AmountIn), 0)) - CCur(Nz(Sum(o.AmountOut), 0)) > 0, CCur(Nz(Sum(o.AmountIn), 0)) - CCur(Nz(Sum(o.AmountOut), 0)), 0),
       IIf(CCur(Nz(Sum(o.AmountIn), 0)) - CCur(Nz(Sum(o.AmountOut), 0)) < 0, CCur(Nz(Sum(o.AmountOut), 0)) - CCur(Nz(Sum(o.AmountIn), 0)), 0),
       QLong('CashBoxID')
FROM qryCashMovements AS o
WHERE (QLong('CashBoxID') = 0 OR o.CashBoxID = QLong('CashBoxID')) AND o.MoveDate < QDate('PeriodStart')
ORDER BY SortKey, MoveDate
```

## qryCashDays

مقبوضات ومدفوعات كل يوم داخل الفترة

المعاملات: `PeriodStart`, `PeriodEnd`, `CashBoxID`

```sql
SELECT DateValue(m.MoveDate) AS CashDay, Sum(m.AmountIn) AS Receipts, Sum(m.AmountOut) AS Payments,
       Count(*) AS MoveCount
FROM qryCashMovements AS m
WHERE (QLong('CashBoxID') = 0 OR m.CashBoxID = QLong('CashBoxID')) AND m.MoveDate >= QDate('PeriodStart') AND m.MoveDate < QDate('PeriodEnd')
GROUP BY DateValue(m.MoveDate)
```

## CashDailyQuery

حركة الخزينة اليومية: رصيد أول اليوم والمقبوضات والمدفوعات ورصيد آخر اليوم

المعاملات: `PeriodStart`, `PeriodEnd`, `CashBoxID`

```sql
SELECT d.CashDay,
       (SELECT CCur(Nz(Sum(x.AmountIn - x.AmountOut), 0)) FROM qryCashMovements AS x
        WHERE (QLong('CashBoxID') = 0 OR x.CashBoxID = QLong('CashBoxID'))
          AND x.MoveDate < d.CashDay) AS OpeningBalance,
       d.Receipts, d.Payments,
       (SELECT CCur(Nz(Sum(y.AmountIn - y.AmountOut), 0)) FROM qryCashMovements AS y
        WHERE (QLong('CashBoxID') = 0 OR y.CashBoxID = QLong('CashBoxID'))
          AND y.MoveDate < d.CashDay) + d.Receipts - d.Payments AS ClosingBalance,
       d.MoveCount
FROM qryCashDays AS d
ORDER BY d.CashDay
```

## CashClosingsQuery

تصفيات يومية الكاشير خلال فترة (0 = كل الصناديق)

المعاملات: `PeriodStart`, `PeriodEnd`, `CashBoxID`

```sql
SELECT c.ClosingID, c.ClosingNumber, c.ClosingDate, b.BoxName, e.EmployeeName, c.PeriodStart,
       c.OpeningBalance, c.CashIn, c.CashOut, c.ExpectedBalance, c.CountedAmount, c.Difference,
       IIf(c.Destination = 'MAIN', 'الخزينة الرئيسية', IIf(c.Destination = 'OWNER', 'تسوية مع المالك',
           'يبقى في الصندوق')) AS DestinationName,
       t.BoxName AS ToBoxName, c.TransferAmount, c.KeptAmount, c.Notes, c.CashBoxID
FROM ((CashClosings AS c INNER JOIN CashBoxes AS b ON c.CashBoxID = b.CashBoxID)
      INNER JOIN Employees AS e ON c.EmployeeID = e.EmployeeID)
     LEFT JOIN CashBoxes AS t ON c.ToCashBoxID = t.CashBoxID
WHERE (QLong('CashBoxID') = 0 OR c.CashBoxID = QLong('CashBoxID')) AND c.ClosingDate >= QDate('PeriodStart') AND c.ClosingDate < QDate('PeriodEnd')
ORDER BY c.ClosingDate
```

## qryCashClosingPrint

بيانات طباعة تصفية الكاشير

```sql
SELECT c.ClosingID, c.ClosingNumber, c.ClosingDate, b.BoxName, e.EmployeeName, c.PeriodStart,
       c.OpeningBalance, c.CashIn, c.CashOut, c.ExpectedBalance, c.CountedAmount, c.Difference,
       IIf(c.Destination = 'MAIN', 'الخزينة الرئيسية', IIf(c.Destination = 'OWNER', 'تسوية مع المالك',
           'يبقى في الصندوق')) AS DestinationName,
       t.BoxName AS ToBoxName, c.TransferAmount, c.KeptAmount, c.Notes
FROM ((CashClosings AS c INNER JOIN CashBoxes AS b ON c.CashBoxID = b.CashBoxID)
      INNER JOIN Employees AS e ON c.EmployeeID = e.EmployeeID)
     LEFT JOIN CashBoxes AS t ON c.ToCashBoxID = t.CashBoxID
```

## qryCashVoucherPrint

بيانات طباعة سندات قبض وصرف وتحويل النقدية

```sql
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
     LEFT JOIN ExpenseTypes AS x ON ex.ExpenseTypeID = x.ExpenseTypeID
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
