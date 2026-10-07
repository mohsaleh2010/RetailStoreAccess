# مرجع الاستعلامات (Queries Reference)

> ملف مُولَّد تلقائيًا من `tools/queries.py` – لا تعدّله يدويًا.

عدد الاستعلامات: **99**. الاستعلامات التي تبدأ بـ `qry` مساعدة تستخدمها الاستعلامات الأخرى؛ البقية تُستخدم مباشرة في التقارير والنماذج. ⭐ = مطلوب بالاسم في البرومبت.

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
| 52 | [`qryCashDayOpening`](#qrycashdayopening) | رصيد أول كل يوم من أيام الحركة (كل الحركات قبل ذلك اليوم) | `PeriodStart`, `PeriodEnd`, `CashBoxID` |
| 53 | [`CashDailyQuery`](#cashdailyquery) | حركة الخزينة اليومية: رصيد أول اليوم والمقبوضات والمدفوعات ورصيد آخر اليوم | `PeriodStart`, `PeriodEnd`, `CashBoxID` |
| 54 | [`CashClosingsQuery`](#cashclosingsquery) | تصفيات يومية الكاشير خلال فترة (0 = كل الصناديق) | `PeriodStart`, `PeriodEnd`, `CashBoxID` |
| 55 | [`qryCashClosingPrint`](#qrycashclosingprint) | بيانات طباعة تصفية الكاشير |  |
| 56 | [`qryCashVoucherPrint`](#qrycashvoucherprint) | بيانات طباعة سندات قبض وصرف وتحويل النقدية |  |
| 57 | [`qrySaleCost`](#qrysalecost) | تكلفة كل فاتورة بيع |  |
| 58 | [`qryReturnCost`](#qryreturncost) | تكلفة ما عاد للمخزون من كل مرتجع بيع |  |
| 59 | [`qryStockCountValue`](#qrystockcountvalue) | قيمة فروقات كل جرد مُرحّل |  |
| 60 | [`qryJournalSale`](#qryjournalsale) | أسطر قيود فواتير البيع |  |
| 61 | [`qryJournalSalesReturn`](#qryjournalsalesreturn) | أسطر قيود مرتجعات البيع |  |
| 62 | [`qryJournalPurchase`](#qryjournalpurchase) | أسطر قيود فواتير الشراء |  |
| 63 | [`qryJournalPurchaseReturn`](#qryjournalpurchasereturn) | أسطر قيود مرتجعات الشراء |  |
| 64 | [`qryJournalPayments`](#qryjournalpayments) | أسطر قيود سندات القبض من العملاء والصرف للموردين |  |
| 65 | [`qryJournalExpense`](#qryjournalexpense) | أسطر قيود المصروفات (عدا المسجلة بسند نقدية) |  |
| 66 | [`qryJournalCashVoucher`](#qryjournalcashvoucher) | أسطر قيود سندات النقدية (قبض وصرف وتحويل) |  |
| 67 | [`qryJournalStock`](#qryjournalstock) | أسطر قيود حركات المخزون اليدوية وتسويات الجرد |  |
| 68 | [`qryJournalOpening`](#qryjournalopening) | أسطر قيود الأرصدة الافتتاحية للصناديق والعملاء والموردين |  |
| 69 | [`qryManualEntryLines`](#qrymanualentrylines) | أسطر القيود اليدوية مع رأس كل قيد |  |
| 70 | [`qryJournalManual`](#qryjournalmanual) | أسطر القيود اليدوية |  |
| 71 | [`qryYearCloseLines`](#qryyearcloselines) | أسطر قيود إقفال السنوات مع رأس كل إقفال |  |
| 72 | [`qryJournalYearClose`](#qryjournalyearclose) | أسطر قيود إقفال السنوات: الإيرادات والمصروفات إلى الأرباح المحتجزة |  |
| 73 | [`JournalLinesQuery`](#journallinesquery) | قيود اليومية خلال فترة بأسطرها | `PeriodStart`, `PeriodEnd` |
| 74 | [`qryJournalEntryPrint`](#qryjournalentryprint) | بيانات طباعة قيد |  |
| 75 | [`qryTrialBefore`](#qrytrialbefore) | مجموع الحسابات قبل الفترة | `PeriodStart` |
| 76 | [`qryTrialPeriod`](#qrytrialperiod) | حركة الحسابات خلال الفترة | `PeriodStart`, `PeriodEnd` |
| 77 | [`TrialBalanceQuery`](#trialbalancequery) | ميزان المراجعة: رصيد أول المدة وحركة الفترة والرصيد الختامي (المدين موجب) | `PeriodStart`, `PeriodEnd` |
| 78 | [`qryStatementBefore`](#qrystatementbefore) | رصيد الحساب المختار (مع حساباته التابعة) قبل بداية الفترة | `PeriodStart`, `AccountCode` |
| 79 | [`AccountStatementQuery`](#accountstatementquery) | كشف حساب لفترة: رصيد أول المدة ثم كل سطر قيد (الحساب الرئيسي يشمل حساباته التابعة) | `PeriodStart`, `PeriodEnd`, `AccountCode` |
| 80 | [`GeneralLedgerQuery`](#generalledgerquery) | دفتر الأستاذ لفترة: لكل حساب فرعي رصيد أول المدة ثم أسطر قيوده (0 = كل الحسابات) | `PeriodStart`, `PeriodEnd`, `AccountCode` |
| 81 | [`qryTreeRollup`](#qrytreerollup) | أرصدة ميزان المراجعة مجمّعة على كل مستوى من شجرة الحسابات | `PeriodStart`, `PeriodEnd` |
| 82 | [`TrialBalanceTreeQuery`](#trialbalancetreequery) | ميزان المراجعة بالمستويات: كل حساب رئيسي بمجموع حساباته التابعة | `PeriodStart`, `PeriodEnd` |
| 83 | [`qryIncomeMoves`](#qryincomemoves) | حركة الحسابات في الفترة بدون قيود إقفال السنة | `PeriodStart`, `PeriodEnd` |
| 84 | [`qryCompareMoves`](#qrycomparemoves) | حركة الحسابات في فترة المقارنة بدون قيود إقفال السنة | `CompareStart`, `CompareEnd` |
| 85 | [`qryIncomeAccounts`](#qryincomeaccounts) | حسابات قائمة الدخل: صافي حركة كل حساب إيرادات أو مصروفات في الفترة وفترة المقارنة | `PeriodStart`, `PeriodEnd`, `CompareStart`, `CompareEnd` |
| 86 | [`IncomeStatementQuery`](#incomestatementquery) | قائمة الدخل: الإيرادات والتكاليف والمصروفات ومجمل وصافي الربح، مع فترة المقارنة | `PeriodStart`, `PeriodEnd`, `CompareStart`, `CompareEnd` |
| 87 | [`qryBalanceAt`](#qrybalanceat) | رصيد كل حساب في نهاية الفترة (مدين موجب) | `PeriodEnd` |
| 88 | [`qryBalanceCompare`](#qrybalancecompare) | رصيد كل حساب في نهاية فترة المقارنة (مدين موجب) | `CompareEnd` |
| 89 | [`qryBalanceAccounts`](#qrybalanceaccounts) | حسابات الميزانية: رصيد كل حساب أصول أو خصوم أو حقوق ملكية (بطبيعته موجب) | `PeriodEnd`, `CompareEnd` |
| 90 | [`qryProfitAt`](#qryprofitat) | صافي ربح الفترات غير المقفلة حتى نهاية الفترة (مدين موجب) | `PeriodEnd` |
| 91 | [`qryProfitCompare`](#qryprofitcompare) | صافي ربح الفترات غير المقفلة حتى نهاية فترة المقارنة (مدين موجب) | `CompareEnd` |
| 92 | [`qryBalanceItems`](#qrybalanceitems) | بنود الميزانية بمجموعاتها، ومعها صافي الربح غير المقفل في الأرباح المحتجزة (32) | `PeriodEnd`, `CompareEnd` |
| 93 | [`BalanceSheetQuery`](#balancesheetquery) | الميزانية العمومية في نهاية الفترة: الأصول = الخصوم + حقوق الملكية، مع فترة المقارنة | `PeriodStart`, `PeriodEnd`, `CompareStart`, `CompareEnd` |
| 94 | [`AccountTreeQuery`](#accounttreequery) | شجرة الحسابات: كل حساب بمستواه ونوعه وهل يقبل القيود |  |
| 95 | [`qrySalesInvoiceLineTotals`](#qrysalesinvoicelinetotals) | مجموع أسطر كل فاتورة بيع |  |
| 96 | [`qryPurchaseInvoiceLineTotals`](#qrypurchaseinvoicelinetotals) | مجموع أسطر كل فاتورة شراء |  |
| 97 | [`qrySalesReturnedQty`](#qrysalesreturnedqty) | الكمية المرتجعة من كل سطر فاتورة بيع |  |
| 98 | [`qryPurchaseReturnedQty`](#qrypurchasereturnedqty) | الكمية المرتجعة للمورد من كل سطر فاتورة شراء |  |
| 99 | [`IntegrityCheckQuery`](#integritycheckquery) | فحص سلامة البيانات: أي سطر هنا مشكلة يجب مراجعتها (النتيجة الفارغة = سليم) |  |

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
| 78 | قيود Sale: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalSale GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 79 | قيود SalesReturn: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalSalesReturn GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 80 | قيود Purchase: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalPurchase GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 81 | قيود PurchaseReturn: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalPurchaseReturn GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 82 | قيود Payments: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalPayments GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 83 | قيود Expense: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalExpense GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 84 | قيود CashVoucher: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalCashVoucher GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 85 | قيود Stock: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalStock GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 86 | قيود Opening: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalOpening GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 87 | قيد الفاتورة الآجلة: 6 أسطر (نقدي، عميل، مبيعات، ضريبة، تكلفة، مخزون) | `SELECT COUNT(*) FROM qryJournalSale WHERE SourceID = {ref:INV2}` | 6 |
| 88 | قيد الفاتورة الآجلة: المتبقي على العميل 360 في ذمم العملاء | `SELECT Debit FROM qryJournalSale WHERE SourceID = {ref:INV2} AND AccountCode = 1300` | 360 |
| 89 | قيد الفاتورة الآجلة: التكلفة = 2×60 + 10×10 | `SELECT Debit FROM qryJournalSale WHERE SourceID = {ref:INV2} AND AccountCode = 5100` | 220 |
| 90 | ذمم العملاء من القيود = أرصدة العملاء (164) | `SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 1300` | 164 |
| 91 | ذمم الموردين من القيود = رصيد المورد (2855 دائن) | `SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 2100` | -2855 |
| 92 | المخزون من القيود = قيمة المخزون بالتكلفة (7040) | `SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalStock) AS x WHERE AccountCode = 1400` | 7040 |
| 93 | صندوق الكاشير من القيود = رصيده (770) | `SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalExpense UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalCashVoucher UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalStock UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 110000 + {ref:BOXC}` | 770 |
| 94 | الخزينة من القيود = رصيدها (7700) | `SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalExpense UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalCashVoucher UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalStock UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 110000 + {ref:BOXM}` | 7700 |
| 95 | ضريبة المخرجات من القيود = 15 + 150 + 60 − 6 | `SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn) AS x WHERE AccountCode = 2200` | -219 |
| 96 | مصروف سند النقدية لا يُقيَّد مرتين | `SELECT COUNT(*) FROM qryJournalExpense WHERE SourceID = {ref:EXPV}` | 0 |
| 97 | القيد اليدوي: 3 أسطر متوازنة (3000) | `SELECT COUNT(*) FROM qryJournalManual WHERE SourceID = {ref:MJ1} AND SourceType = 'MANUAL'` | 3 |
| 98 | قيود Manual: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalManual GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 99 | سند صرف المصروف يُقيَّد على حساب نوع المصروف | `SELECT Debit FROM qryJournalCashVoucher WHERE SourceID = {ref:V1} AND AccountCode = 530009` | 50 |
| 100 | فحص السلامة: لا توجد مشكلات | `SELECT COUNT(*) FROM IntegrityCheckQuery` | 0 |

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

## qryCashDayOpening

رصيد أول كل يوم من أيام الحركة (كل الحركات قبل ذلك اليوم)

المعاملات: `PeriodStart`, `PeriodEnd`, `CashBoxID`

```sql
SELECT d.CashDay, Sum(x.AmountIn - x.AmountOut) AS DayOpening
FROM qryCashDays AS d, qryCashMovements AS x
WHERE (QLong('CashBoxID') = 0 OR x.CashBoxID = QLong('CashBoxID')) AND x.MoveDate < d.CashDay
GROUP BY d.CashDay
```

## CashDailyQuery

حركة الخزينة اليومية: رصيد أول اليوم والمقبوضات والمدفوعات ورصيد آخر اليوم

المعاملات: `PeriodStart`, `PeriodEnd`, `CashBoxID`

```sql
SELECT d.CashDay, CCur(Nz(o.DayOpening, 0)) AS OpeningBalance, d.Receipts, d.Payments,
       CCur(Nz(o.DayOpening, 0)) + d.Receipts - d.Payments AS ClosingBalance, d.MoveCount
FROM qryCashDays AS d LEFT JOIN qryCashDayOpening AS o ON d.CashDay = o.CashDay
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

## qrySaleCost

تكلفة كل فاتورة بيع

```sql
SELECT SalesInvoiceID, Sum(Quantity * UnitCost) AS SaleCost
FROM SalesInvoiceDetails
GROUP BY SalesInvoiceID
```

## qryReturnCost

تكلفة ما عاد للمخزون من كل مرتجع بيع

```sql
SELECT SalesReturnID, Sum(IIf(ReturnToStock, Quantity * UnitCost, 0)) AS ReturnCost
FROM SalesReturnDetails
GROUP BY SalesReturnID
```

## qryStockCountValue

قيمة فروقات كل جرد مُرحّل

```sql
SELECT ReferenceID AS StockCountID, Max(ReferenceNumber) AS CountNumber, Max(TransactionDate) AS CountDate,
       Sum(Quantity * UnitCost) AS CountValue
FROM InventoryTransactions
WHERE ReferenceType = 'STOCK_COUNT'
GROUP BY ReferenceID
```

## qryJournalSale

أسطر قيود فواتير البيع

```sql
SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 1 AS LineOrder, IIf(h.CashBoxID Is Null, IIf(h.PaymentMethodID Is Null Or h.PaymentMethodID = 1, 1190, 1200), 110000 + h.CashBoxID) AS AccountCode, h.PaidAmount AS Debit, CCur(0) AS Credit, c.CustomerName AS LineText
FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID
WHERE h.PaidAmount <> 0
UNION ALL
SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 2 AS LineOrder, 1300 AS AccountCode, h.RemainingAmount AS Debit, CCur(0) AS Credit, c.CustomerName AS LineText
FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID
WHERE h.RemainingAmount <> 0
UNION ALL
SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 3 AS LineOrder, 4100 AS AccountCode, CCur(0) AS Debit, h.TaxableAmount AS Credit, 'المبيعات' AS LineText
FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID
WHERE h.TaxableAmount <> 0
UNION ALL
SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 4 AS LineOrder, 2200 AS AccountCode, CCur(0) AS Debit, h.Tax AS Credit, 'ضريبة المخرجات' AS LineText
FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID
WHERE h.Tax <> 0
UNION ALL
SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 5 AS LineOrder, 5100 AS AccountCode, k.SaleCost AS Debit, CCur(0) AS Credit, 'تكلفة البضاعة المباعة' AS LineText
FROM (SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID) INNER JOIN qrySaleCost AS k ON h.SalesInvoiceID = k.SalesInvoiceID
WHERE k.SaleCost <> 0
UNION ALL
SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 6 AS LineOrder, 1400 AS AccountCode, CCur(0) AS Debit, k.SaleCost AS Credit, 'المخزون' AS LineText
FROM (SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID) INNER JOIN qrySaleCost AS k ON h.SalesInvoiceID = k.SalesInvoiceID
WHERE k.SaleCost <> 0
```

## qryJournalSalesReturn

أسطر قيود مرتجعات البيع

```sql
SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 1 AS LineOrder, 4110 AS AccountCode, r.TaxableAmount AS Debit, CCur(0) AS Credit, 'مردودات المبيعات' AS LineText
FROM SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID
WHERE r.TaxableAmount <> 0
UNION ALL
SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 2 AS LineOrder, 2200 AS AccountCode, r.Tax AS Debit, CCur(0) AS Credit, 'ضريبة المخرجات' AS LineText
FROM SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID
WHERE r.Tax <> 0
UNION ALL
SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 3 AS LineOrder, IIf(r.CashBoxID Is Null, IIf(r.PaymentMethodID Is Null Or r.PaymentMethodID = 1, 1190, 1200), 110000 + r.CashBoxID) AS AccountCode, CCur(0) AS Debit, r.RefundedAmount AS Credit, c.CustomerName AS LineText
FROM SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID
WHERE r.RefundedAmount <> 0
UNION ALL
SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 4 AS LineOrder, 1300 AS AccountCode, CCur(0) AS Debit, r.TotalAmount - r.RefundedAmount AS Credit, c.CustomerName AS LineText
FROM SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID
WHERE r.TotalAmount - r.RefundedAmount <> 0
UNION ALL
SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 5 AS LineOrder, 1400 AS AccountCode, k.ReturnCost AS Debit, CCur(0) AS Credit, 'المخزون' AS LineText
FROM (SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID) INNER JOIN qryReturnCost AS k ON r.SalesReturnID = k.SalesReturnID
WHERE k.ReturnCost <> 0
UNION ALL
SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 6 AS LineOrder, 5100 AS AccountCode, CCur(0) AS Debit, k.ReturnCost AS Credit, 'تكلفة البضاعة المباعة' AS LineText
FROM (SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID) INNER JOIN qryReturnCost AS k ON r.SalesReturnID = k.SalesReturnID
WHERE k.ReturnCost <> 0
```

## qryJournalPurchase

أسطر قيود فواتير الشراء

```sql
SELECT 'PURCHASE' AS SourceType, h.PurchaseInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, s.SupplierName AS Party, 1 AS LineOrder, 1400 AS AccountCode, h.TaxableAmount AS Debit, CCur(0) AS Credit, 'المخزون' AS LineText
FROM PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID
WHERE h.TaxableAmount <> 0
UNION ALL
SELECT 'PURCHASE' AS SourceType, h.PurchaseInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, s.SupplierName AS Party, 2 AS LineOrder, 1500 AS AccountCode, h.Tax AS Debit, CCur(0) AS Credit, 'ضريبة المدخلات' AS LineText
FROM PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID
WHERE h.Tax <> 0
UNION ALL
SELECT 'PURCHASE' AS SourceType, h.PurchaseInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, s.SupplierName AS Party, 3 AS LineOrder, IIf(h.CashBoxID Is Null, IIf(h.PaymentMethodID Is Null Or h.PaymentMethodID = 1, 1190, 1200), 110000 + h.CashBoxID) AS AccountCode, CCur(0) AS Debit, h.PaidAmount AS Credit, s.SupplierName AS LineText
FROM PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID
WHERE h.PaidAmount <> 0
UNION ALL
SELECT 'PURCHASE' AS SourceType, h.PurchaseInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, s.SupplierName AS Party, 4 AS LineOrder, 2100 AS AccountCode, CCur(0) AS Debit, h.RemainingAmount AS Credit, s.SupplierName AS LineText
FROM PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID
WHERE h.RemainingAmount <> 0
```

## qryJournalPurchaseReturn

أسطر قيود مرتجعات الشراء

```sql
SELECT 'PURCHASE_RETURN' AS SourceType, r.PurchaseReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, s.SupplierName AS Party, 1 AS LineOrder, IIf(r.CashBoxID Is Null, IIf(r.PaymentMethodID Is Null Or r.PaymentMethodID = 1, 1190, 1200), 110000 + r.CashBoxID) AS AccountCode, r.RefundedAmount AS Debit, CCur(0) AS Credit, s.SupplierName AS LineText
FROM PurchaseReturns AS r INNER JOIN Suppliers AS s ON r.SupplierID = s.SupplierID
WHERE r.RefundedAmount <> 0
UNION ALL
SELECT 'PURCHASE_RETURN' AS SourceType, r.PurchaseReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, s.SupplierName AS Party, 2 AS LineOrder, 2100 AS AccountCode, r.TotalAmount - r.RefundedAmount AS Debit, CCur(0) AS Credit, s.SupplierName AS LineText
FROM PurchaseReturns AS r INNER JOIN Suppliers AS s ON r.SupplierID = s.SupplierID
WHERE r.TotalAmount - r.RefundedAmount <> 0
UNION ALL
SELECT 'PURCHASE_RETURN' AS SourceType, r.PurchaseReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, s.SupplierName AS Party, 3 AS LineOrder, 1400 AS AccountCode, CCur(0) AS Debit, r.TaxableAmount AS Credit, 'المخزون' AS LineText
FROM PurchaseReturns AS r INNER JOIN Suppliers AS s ON r.SupplierID = s.SupplierID
WHERE r.TaxableAmount <> 0
UNION ALL
SELECT 'PURCHASE_RETURN' AS SourceType, r.PurchaseReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, s.SupplierName AS Party, 4 AS LineOrder, 1500 AS AccountCode, CCur(0) AS Debit, r.Tax AS Credit, 'ضريبة المدخلات' AS LineText
FROM PurchaseReturns AS r INNER JOIN Suppliers AS s ON r.SupplierID = s.SupplierID
WHERE r.Tax <> 0
```

## qryJournalPayments

أسطر قيود سندات القبض من العملاء والصرف للموردين

```sql
SELECT 'CUSTOMER_PAYMENT' AS SourceType, p.PaymentID AS SourceID, p.PaymentNumber AS SourceNumber, p.PaymentDate AS SourceDate, c.CustomerName AS Party, 1 AS LineOrder, IIf(p.CashBoxID Is Null, IIf(p.PaymentMethodID Is Null Or p.PaymentMethodID = 1, 1190, 1200), 110000 + p.CashBoxID) AS AccountCode, p.Amount AS Debit, CCur(0) AS Credit, c.CustomerName AS LineText
FROM CustomerPayments AS p INNER JOIN Customers AS c ON p.CustomerID = c.CustomerID
WHERE p.Amount <> 0
UNION ALL
SELECT 'CUSTOMER_PAYMENT' AS SourceType, p.PaymentID AS SourceID, p.PaymentNumber AS SourceNumber, p.PaymentDate AS SourceDate, c.CustomerName AS Party, 2 AS LineOrder, 1300 AS AccountCode, CCur(0) AS Debit, p.Amount AS Credit, c.CustomerName AS LineText
FROM CustomerPayments AS p INNER JOIN Customers AS c ON p.CustomerID = c.CustomerID
WHERE p.Amount <> 0
UNION ALL
SELECT 'SUPPLIER_PAYMENT' AS SourceType, p.PaymentID AS SourceID, p.PaymentNumber AS SourceNumber, p.PaymentDate AS SourceDate, s.SupplierName AS Party, 1 AS LineOrder, 2100 AS AccountCode, p.Amount AS Debit, CCur(0) AS Credit, s.SupplierName AS LineText
FROM SupplierPayments AS p INNER JOIN Suppliers AS s ON p.SupplierID = s.SupplierID
WHERE p.Amount <> 0
UNION ALL
SELECT 'SUPPLIER_PAYMENT' AS SourceType, p.PaymentID AS SourceID, p.PaymentNumber AS SourceNumber, p.PaymentDate AS SourceDate, s.SupplierName AS Party, 2 AS LineOrder, IIf(p.CashBoxID Is Null, IIf(p.PaymentMethodID Is Null Or p.PaymentMethodID = 1, 1190, 1200), 110000 + p.CashBoxID) AS AccountCode, CCur(0) AS Debit, p.Amount AS Credit, s.SupplierName AS LineText
FROM SupplierPayments AS p INNER JOIN Suppliers AS s ON p.SupplierID = s.SupplierID
WHERE p.Amount <> 0
```

## qryJournalExpense

أسطر قيود المصروفات (عدا المسجلة بسند نقدية)

```sql
SELECT 'EXPENSE' AS SourceType, e.ExpenseID AS SourceID, e.ExpenseNumber AS SourceNumber, e.ExpenseDate AS SourceDate, t.ExpenseTypeName AS Party, 1 AS LineOrder, 530000 + e.ExpenseTypeID AS AccountCode, e.Amount AS Debit, CCur(0) AS Credit, t.ExpenseTypeName AS LineText
FROM (Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID) LEFT JOIN CashVouchers AS v ON e.ExpenseID = v.ExpenseID
WHERE v.CashVoucherID Is Null AND e.Amount <> 0
UNION ALL
SELECT 'EXPENSE' AS SourceType, e.ExpenseID AS SourceID, e.ExpenseNumber AS SourceNumber, e.ExpenseDate AS SourceDate, t.ExpenseTypeName AS Party, 2 AS LineOrder, 1500 AS AccountCode, e.Tax AS Debit, CCur(0) AS Credit, 'ضريبة المدخلات' AS LineText
FROM (Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID) LEFT JOIN CashVouchers AS v ON e.ExpenseID = v.ExpenseID
WHERE v.CashVoucherID Is Null AND e.Tax <> 0
UNION ALL
SELECT 'EXPENSE' AS SourceType, e.ExpenseID AS SourceID, e.ExpenseNumber AS SourceNumber, e.ExpenseDate AS SourceDate, t.ExpenseTypeName AS Party, 3 AS LineOrder, IIf(e.CashBoxID Is Null, IIf(e.PaymentMethodID Is Null Or e.PaymentMethodID = 1, 1190, 1200), 110000 + e.CashBoxID) AS AccountCode, CCur(0) AS Debit, e.TotalAmount AS Credit, e.Description AS LineText
FROM (Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID) LEFT JOIN CashVouchers AS v ON e.ExpenseID = v.ExpenseID
WHERE v.CashVoucherID Is Null AND e.TotalAmount <> 0
```

## qryJournalCashVoucher

أسطر قيود سندات النقدية (قبض وصرف وتحويل)

```sql
SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 1 AS LineOrder, 110000 + v.CashBoxID AS AccountCode, v.Amount AS Debit, CCur(0) AS Credit, v.PartyName AS LineText
FROM CashVouchers AS v
WHERE v.VoucherType = 'IN'
UNION ALL
SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 2 AS LineOrder, IIf(v.Category = 'OWNER', 3100, IIf(v.Category = 'OVERAGE', 4300, 4200)) AS AccountCode, CCur(0) AS Debit, v.Amount AS Credit, v.Description AS LineText
FROM CashVouchers AS v
WHERE v.VoucherType = 'IN'
UNION ALL
SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 1 AS LineOrder, IIf(v.Category = 'OWNER', 3100, IIf(v.Category = 'ADVANCE', 1600, IIf(v.Category = 'SHORTAGE', 5400, IIf(v.Category = 'EXPENSE' AND x.ExpenseTypeID Is Not Null, 530000 + x.ExpenseTypeID, 5900)))) AS AccountCode, v.Amount AS Debit, CCur(0) AS Credit, v.Description AS LineText
FROM CashVouchers AS v LEFT JOIN Expenses AS x ON v.ExpenseID = x.ExpenseID
WHERE v.VoucherType = 'OUT'
UNION ALL
SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 2 AS LineOrder, 110000 + v.CashBoxID AS AccountCode, CCur(0) AS Debit, v.Amount AS Credit, v.PartyName AS LineText
FROM CashVouchers AS v LEFT JOIN Expenses AS x ON v.ExpenseID = x.ExpenseID
WHERE v.VoucherType = 'OUT'
UNION ALL
SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 1 AS LineOrder, 110000 + v.ToCashBoxID AS AccountCode, v.Amount AS Debit, CCur(0) AS Credit, v.Description AS LineText
FROM CashVouchers AS v
WHERE v.VoucherType = 'TRANSFER'
UNION ALL
SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 2 AS LineOrder, 110000 + v.CashBoxID AS AccountCode, CCur(0) AS Debit, v.Amount AS Credit, v.Description AS LineText
FROM CashVouchers AS v
WHERE v.VoucherType = 'TRANSFER'
```

## qryJournalStock

أسطر قيود حركات المخزون اليدوية وتسويات الجرد

```sql
SELECT 'STOCK_MOVE' AS SourceType, i.TransactionID AS SourceID, i.ReferenceNumber AS SourceNumber, i.TransactionDate AS SourceDate, p.ProductName AS Party, 1 AS LineOrder, 1400 AS AccountCode, IIf(i.Quantity * i.UnitCost > 0, i.Quantity * i.UnitCost, 0) AS Debit, IIf(i.Quantity * i.UnitCost < 0, -i.Quantity * i.UnitCost, 0) AS Credit, i.Notes AS LineText
FROM InventoryTransactions AS i INNER JOIN Products AS p ON i.ProductID = p.ProductID
WHERE i.ReferenceType = 'MANUAL' AND i.Quantity * i.UnitCost <> 0
UNION ALL
SELECT 'STOCK_MOVE' AS SourceType, i.TransactionID AS SourceID, i.ReferenceNumber AS SourceNumber, i.TransactionDate AS SourceDate, p.ProductName AS Party, 2 AS LineOrder, IIf(i.TransactionTypeID = 8, 3900, 5200) AS AccountCode, IIf(i.Quantity * i.UnitCost < 0, -i.Quantity * i.UnitCost, 0) AS Debit, IIf(i.Quantity * i.UnitCost > 0, i.Quantity * i.UnitCost, 0) AS Credit, i.Notes AS LineText
FROM InventoryTransactions AS i INNER JOIN Products AS p ON i.ProductID = p.ProductID
WHERE i.ReferenceType = 'MANUAL' AND i.Quantity * i.UnitCost <> 0
UNION ALL
SELECT 'STOCK_COUNT' AS SourceType, k.StockCountID AS SourceID, k.CountNumber AS SourceNumber, k.CountDate AS SourceDate, 'تسوية الجرد' AS Party, 1 AS LineOrder, 1400 AS AccountCode, IIf(k.CountValue > 0, k.CountValue, 0) AS Debit, IIf(k.CountValue < 0, -k.CountValue, 0) AS Credit, 'المخزون' AS LineText
FROM qryStockCountValue AS k
WHERE k.CountValue <> 0
UNION ALL
SELECT 'STOCK_COUNT' AS SourceType, k.StockCountID AS SourceID, k.CountNumber AS SourceNumber, k.CountDate AS SourceDate, 'تسوية الجرد' AS Party, 2 AS LineOrder, 5200 AS AccountCode, IIf(k.CountValue < 0, -k.CountValue, 0) AS Debit, IIf(k.CountValue > 0, k.CountValue, 0) AS Credit, 'فروقات الجرد' AS LineText
FROM qryStockCountValue AS k
WHERE k.CountValue <> 0
```

## qryJournalOpening

أسطر قيود الأرصدة الافتتاحية للصناديق والعملاء والموردين

```sql
SELECT 'BOX_OPENING' AS SourceType, b.CashBoxID AS SourceID, b.BoxName AS SourceNumber, b.OpeningDate AS SourceDate, b.BoxName AS Party, 1 AS LineOrder, 110000 + b.CashBoxID AS AccountCode, b.OpeningBalance AS Debit, CCur(0) AS Credit, b.BoxName AS LineText
FROM CashBoxes AS b
WHERE b.OpeningBalance <> 0
UNION ALL
SELECT 'BOX_OPENING' AS SourceType, b.CashBoxID AS SourceID, b.BoxName AS SourceNumber, b.OpeningDate AS SourceDate, b.BoxName AS Party, 2 AS LineOrder, 3900 AS AccountCode, CCur(0) AS Debit, b.OpeningBalance AS Credit, 'رصيد افتتاحي' AS LineText
FROM CashBoxes AS b
WHERE b.OpeningBalance <> 0
UNION ALL
SELECT 'CUSTOMER_OPENING' AS SourceType, c.CustomerID AS SourceID, c.CustomerName AS SourceNumber, c.CreatedAt AS SourceDate, c.CustomerName AS Party, 1 AS LineOrder, 1300 AS AccountCode, IIf(c.OpeningBalance > 0, c.OpeningBalance, 0) AS Debit, IIf(c.OpeningBalance < 0, -c.OpeningBalance, 0) AS Credit, c.CustomerName AS LineText
FROM Customers AS c
WHERE c.OpeningBalance <> 0
UNION ALL
SELECT 'CUSTOMER_OPENING' AS SourceType, c.CustomerID AS SourceID, c.CustomerName AS SourceNumber, c.CreatedAt AS SourceDate, c.CustomerName AS Party, 2 AS LineOrder, 3900 AS AccountCode, IIf(c.OpeningBalance < 0, -c.OpeningBalance, 0) AS Debit, IIf(c.OpeningBalance > 0, c.OpeningBalance, 0) AS Credit, 'رصيد افتتاحي' AS LineText
FROM Customers AS c
WHERE c.OpeningBalance <> 0
UNION ALL
SELECT 'SUPPLIER_OPENING' AS SourceType, s.SupplierID AS SourceID, s.SupplierName AS SourceNumber, s.CreatedAt AS SourceDate, s.SupplierName AS Party, 1 AS LineOrder, 2100 AS AccountCode, IIf(s.OpeningBalance < 0, -s.OpeningBalance, 0) AS Debit, IIf(s.OpeningBalance > 0, s.OpeningBalance, 0) AS Credit, s.SupplierName AS LineText
FROM Suppliers AS s
WHERE s.OpeningBalance <> 0
UNION ALL
SELECT 'SUPPLIER_OPENING' AS SourceType, s.SupplierID AS SourceID, s.SupplierName AS SourceNumber, s.CreatedAt AS SourceDate, s.SupplierName AS Party, 2 AS LineOrder, 3900 AS AccountCode, IIf(s.OpeningBalance > 0, s.OpeningBalance, 0) AS Debit, IIf(s.OpeningBalance < 0, -s.OpeningBalance, 0) AS Credit, 'رصيد افتتاحي' AS LineText
FROM Suppliers AS s
WHERE s.OpeningBalance <> 0
```

## qryManualEntryLines

أسطر القيود اليدوية مع رأس كل قيد

```sql
SELECT h.ManualEntryID, h.EntryNumber, h.EntryDate, h.Description, l.LineNumber AS LineNo,
       l.AccountCode AS LineAccount, l.Debit AS LineDebit, l.Credit AS LineCredit, l.LineText AS LineNote
FROM ManualEntries AS h INNER JOIN ManualEntryLines AS l ON h.ManualEntryID = l.ManualEntryID
```

## qryJournalManual

أسطر القيود اليدوية

```sql
SELECT 'MANUAL' AS SourceType, m.ManualEntryID AS SourceID, m.EntryNumber AS SourceNumber, m.EntryDate AS SourceDate, m.Description AS Party, m.LineNo AS LineOrder, m.LineAccount AS AccountCode, m.LineDebit AS Debit, m.LineCredit AS Credit, m.LineNote AS LineText
FROM qryManualEntryLines AS m
WHERE m.LineDebit + m.LineCredit <> 0
```

## qryYearCloseLines

أسطر قيود إقفال السنوات مع رأس كل إقفال

```sql
SELECT h.YearClosingID, h.ClosingNumber, h.ClosingDate, h.Notes, l.LineNumber AS LineNo,
       l.AccountCode AS LineAccount, l.Debit AS LineDebit, l.Credit AS LineCredit, l.LineText AS LineNote
FROM FiscalYearClosings AS h INNER JOIN FiscalYearClosingLines AS l ON h.YearClosingID = l.YearClosingID
```

## qryJournalYearClose

أسطر قيود إقفال السنوات: الإيرادات والمصروفات إلى الأرباح المحتجزة

```sql
SELECT 'YEAR_CLOSE' AS SourceType, y.YearClosingID AS SourceID, y.ClosingNumber AS SourceNumber, y.ClosingDate AS SourceDate, y.Notes AS Party, y.LineNo AS LineOrder, y.LineAccount AS AccountCode, y.LineDebit AS Debit, y.LineCredit AS Credit, y.LineNote AS LineText
FROM qryYearCloseLines AS y
WHERE y.LineDebit + y.LineCredit <> 0
```

## JournalLinesQuery

قيود اليومية خلال فترة بأسطرها

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT e.EntryID, e.EntryNumber, e.EntryDate, e.SourceType, t.TypeName, e.SourceID, e.SourceNumber,
       e.Description, l.LineNumber, l.AccountCode, a.AccountName, l.LineText, l.Debit, l.Credit
FROM ((JournalEntries AS e INNER JOIN JournalSourceTypes AS t ON e.SourceType = t.SourceType)
      INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID)
     INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode
WHERE e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd')
ORDER BY e.EntryDate, e.EntryNumber, l.LineNumber
```

## qryJournalEntryPrint

بيانات طباعة قيد

```sql
SELECT e.EntryID, e.EntryNumber, e.EntryDate, t.TypeName, e.SourceNumber, e.Description, e.TotalDebit,
       l.LineNumber, l.AccountCode, a.AccountName, l.LineText, l.Debit, l.Credit
FROM ((JournalEntries AS e INNER JOIN JournalSourceTypes AS t ON e.SourceType = t.SourceType)
      INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID)
     INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode
```

## qryTrialBefore

مجموع الحسابات قبل الفترة

المعاملات: `PeriodStart`

```sql
SELECT l.AccountCode, Sum(l.Debit) AS DebitBefore, Sum(l.Credit) AS CreditBefore
FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID
WHERE e.EntryDate < QDate('PeriodStart')
GROUP BY l.AccountCode
```

## qryTrialPeriod

حركة الحسابات خلال الفترة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT l.AccountCode, Sum(l.Debit) AS SumDebit, Sum(l.Credit) AS SumCredit
FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID
WHERE e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd')
GROUP BY l.AccountCode
```

## TrialBalanceQuery

ميزان المراجعة: رصيد أول المدة وحركة الفترة والرصيد الختامي (المدين موجب)

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT a.AccountCode, a.AccountName, a.AccountType,
       CCur(Nz(b.DebitBefore, 0)) - CCur(Nz(b.CreditBefore, 0)) AS OpeningBalance,
       CCur(Nz(p.SumDebit, 0)) AS PeriodDebit, CCur(Nz(p.SumCredit, 0)) AS PeriodCredit,
       CCur(Nz(b.DebitBefore, 0)) - CCur(Nz(b.CreditBefore, 0)) + CCur(Nz(p.SumDebit, 0)) - CCur(Nz(p.SumCredit, 0)) AS ClosingBalance
FROM (Accounts AS a LEFT JOIN qryTrialBefore AS b ON a.AccountCode = b.AccountCode)
     LEFT JOIN qryTrialPeriod AS p ON a.AccountCode = p.AccountCode
WHERE b.AccountCode Is Not Null OR p.AccountCode Is Not Null
ORDER BY a.AccountCode
```

## qryStatementBefore

رصيد الحساب المختار (مع حساباته التابعة) قبل بداية الفترة

المعاملات: `PeriodStart`, `AccountCode`

```sql
SELECT CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) AS SumBefore
FROM (JournalLines AS o INNER JOIN JournalEntries AS f ON o.EntryID = f.EntryID)
     INNER JOIN Accounts AS b ON o.AccountCode = b.AccountCode
WHERE (b.Level1Code = QLong('AccountCode') OR b.Level2Code = QLong('AccountCode') OR b.Level3Code = QLong('AccountCode') OR b.Level4Code = QLong('AccountCode') OR b.Level5Code = QLong('AccountCode')) AND f.EntryDate < QDate('PeriodStart')
```

## AccountStatementQuery

كشف حساب لفترة: رصيد أول المدة ثم كل سطر قيد (الحساب الرئيسي يشمل حساباته التابعة)

المعاملات: `PeriodStart`, `PeriodEnd`, `AccountCode`

```sql
SELECT 1 AS SortKey, s.AccountCode AS StatementAccount, s.AccountName AS StatementName, e.EntryDate AS LineDate,
       e.EntryNumber AS EntryNo, e.EntryID AS EntryRef, k.TypeName AS KindName, e.SourceNumber AS DocNo,
       e.Description AS Details, a.AccountCode AS SubCode, a.AccountName AS SubName, l.Debit AS LineDebit,
       l.Credit AS LineCredit
FROM (((JournalLines AS l INNER JOIN JournalEntries AS e ON l.EntryID = e.EntryID)
      INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode)
     INNER JOIN JournalSourceTypes AS k ON e.SourceType = k.SourceType), Accounts AS s
WHERE s.AccountCode = QLong('AccountCode') AND (a.Level1Code = QLong('AccountCode') OR a.Level2Code = QLong('AccountCode') OR a.Level3Code = QLong('AccountCode') OR a.Level4Code = QLong('AccountCode') OR a.Level5Code = QLong('AccountCode')) AND e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd')
UNION ALL
SELECT 0, s.AccountCode, s.AccountName, QDate('PeriodStart'), '-', 0, 'رصيد أول المدة', Null, Null, Null, Null,
       IIf(x.SumBefore > 0, x.SumBefore, 0), IIf(x.SumBefore < 0, -x.SumBefore, 0)
FROM Accounts AS s, qryStatementBefore AS x
WHERE s.AccountCode = QLong('AccountCode')
ORDER BY SortKey, LineDate, EntryNo
```

## GeneralLedgerQuery

دفتر الأستاذ لفترة: لكل حساب فرعي رصيد أول المدة ثم أسطر قيوده (0 = كل الحسابات)

المعاملات: `PeriodStart`, `PeriodEnd`, `AccountCode`

```sql
SELECT 1 AS SortKey, a.TreeKey AS AccountKey, a.AccountCode AS LedgerCode, a.AccountName AS LedgerName,
       e.EntryDate AS LineDate, e.EntryNumber AS EntryNo, e.EntryID AS EntryRef, k.TypeName AS KindName,
       e.SourceNumber AS DocNo, e.Description AS Details, l.Debit AS LineDebit, l.Credit AS LineCredit
FROM ((JournalLines AS l INNER JOIN JournalEntries AS e ON l.EntryID = e.EntryID)
      INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode)
     INNER JOIN JournalSourceTypes AS k ON e.SourceType = k.SourceType
WHERE (QLong('AccountCode') = 0 OR (a.Level1Code = QLong('AccountCode') OR a.Level2Code = QLong('AccountCode') OR a.Level3Code = QLong('AccountCode') OR a.Level4Code = QLong('AccountCode') OR a.Level5Code = QLong('AccountCode'))) AND e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd')
UNION ALL
SELECT 0, b.TreeKey, b.AccountCode, b.AccountName, QDate('PeriodStart'), '-', 0, 'رصيد أول المدة', Null, Null,
       IIf(t.OpeningBalance > 0, t.OpeningBalance, 0), IIf(t.OpeningBalance < 0, -t.OpeningBalance, 0)
FROM TrialBalanceQuery AS t INNER JOIN Accounts AS b ON t.AccountCode = b.AccountCode
WHERE QLong('AccountCode') = 0 OR (b.Level1Code = QLong('AccountCode') OR b.Level2Code = QLong('AccountCode') OR b.Level3Code = QLong('AccountCode') OR b.Level4Code = QLong('AccountCode') OR b.Level5Code = QLong('AccountCode'))
ORDER BY AccountKey, SortKey, LineDate, EntryNo
```

## qryTreeRollup

أرصدة ميزان المراجعة مجمّعة على كل مستوى من شجرة الحسابات

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT d.Level1Code AS TreeCode, Sum(t.OpeningBalance) AS SumOpening, Sum(t.PeriodDebit) AS SumDebit,
       Sum(t.PeriodCredit) AS SumCredit, Sum(t.ClosingBalance) AS SumClosing
FROM TrialBalanceQuery AS t INNER JOIN Accounts AS d ON t.AccountCode = d.AccountCode
WHERE d.Level1Code Is Not Null
GROUP BY d.Level1Code
UNION ALL
SELECT d.Level2Code AS TreeCode, Sum(t.OpeningBalance) AS SumOpening, Sum(t.PeriodDebit) AS SumDebit,
       Sum(t.PeriodCredit) AS SumCredit, Sum(t.ClosingBalance) AS SumClosing
FROM TrialBalanceQuery AS t INNER JOIN Accounts AS d ON t.AccountCode = d.AccountCode
WHERE d.Level2Code Is Not Null
GROUP BY d.Level2Code
UNION ALL
SELECT d.Level3Code AS TreeCode, Sum(t.OpeningBalance) AS SumOpening, Sum(t.PeriodDebit) AS SumDebit,
       Sum(t.PeriodCredit) AS SumCredit, Sum(t.ClosingBalance) AS SumClosing
FROM TrialBalanceQuery AS t INNER JOIN Accounts AS d ON t.AccountCode = d.AccountCode
WHERE d.Level3Code Is Not Null
GROUP BY d.Level3Code
UNION ALL
SELECT d.Level4Code AS TreeCode, Sum(t.OpeningBalance) AS SumOpening, Sum(t.PeriodDebit) AS SumDebit,
       Sum(t.PeriodCredit) AS SumCredit, Sum(t.ClosingBalance) AS SumClosing
FROM TrialBalanceQuery AS t INNER JOIN Accounts AS d ON t.AccountCode = d.AccountCode
WHERE d.Level4Code Is Not Null
GROUP BY d.Level4Code
UNION ALL
SELECT d.Level5Code AS TreeCode, Sum(t.OpeningBalance) AS SumOpening, Sum(t.PeriodDebit) AS SumDebit,
       Sum(t.PeriodCredit) AS SumCredit, Sum(t.ClosingBalance) AS SumClosing
FROM TrialBalanceQuery AS t INNER JOIN Accounts AS d ON t.AccountCode = d.AccountCode
WHERE d.Level5Code Is Not Null
GROUP BY d.Level5Code
```

## TrialBalanceTreeQuery

ميزان المراجعة بالمستويات: كل حساب رئيسي بمجموع حساباته التابعة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT a.AccountCode, a.AccountName, IIf(a.AccountType = 'ASSET', 'أصول', IIf(a.AccountType = 'LIABILITY', 'خصوم', IIf(a.AccountType = 'EQUITY', 'حقوق ملكية', IIf(a.AccountType = 'REVENUE', 'إيرادات', 'مصروفات')))) AS TypeName, a.AccountLevel, a.TreeKey, a.IsPosting,
       r.SumOpening AS OpeningBalance, r.SumDebit AS PeriodDebit, r.SumCredit AS PeriodCredit,
       r.SumClosing AS ClosingBalance
FROM Accounts AS a INNER JOIN qryTreeRollup AS r ON a.AccountCode = r.TreeCode
ORDER BY a.TreeKey
```

## qryIncomeMoves

حركة الحسابات في الفترة بدون قيود إقفال السنة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT l.AccountCode, Sum(l.Debit) AS SumDebit, Sum(l.Credit) AS SumCredit
FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID
WHERE e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd') AND e.SourceType <> 'YEAR_CLOSE'
GROUP BY l.AccountCode
```

## qryCompareMoves

حركة الحسابات في فترة المقارنة بدون قيود إقفال السنة

المعاملات: `CompareStart`, `CompareEnd`

```sql
SELECT l.AccountCode, Sum(l.Debit) AS SumDebit, Sum(l.Credit) AS SumCredit
FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID
WHERE e.EntryDate >= QDate('CompareStart') AND e.EntryDate < QDate('CompareEnd') AND e.SourceType <> 'YEAR_CLOSE'
GROUP BY l.AccountCode
```

## qryIncomeAccounts

حسابات قائمة الدخل: صافي حركة كل حساب إيرادات أو مصروفات في الفترة وفترة المقارنة

المعاملات: `PeriodStart`, `PeriodEnd`, `CompareStart`, `CompareEnd`

```sql
SELECT a.AccountCode, a.AccountName, a.TreeKey,
       IIf(a.Level2Code = 41, 1, IIf(a.Level2Code = 51, 2, IIf(a.Level2Code = 52, 3,
           IIf(a.AccountType = 'REVENUE', 4, 5)))) AS SectionNo,
       IIf(a.AccountType = 'REVENUE', 1, -1) * (CCur(Nz(c.SumCredit, 0)) - CCur(Nz(c.SumDebit, 0))) AS CurrentAmount,
       IIf(a.AccountType = 'REVENUE', 1, -1) * (CCur(Nz(p.SumCredit, 0)) - CCur(Nz(p.SumDebit, 0))) AS PriorAmount
FROM (Accounts AS a LEFT JOIN qryIncomeMoves AS c ON a.AccountCode = c.AccountCode)
     LEFT JOIN qryCompareMoves AS p ON a.AccountCode = p.AccountCode
WHERE a.AccountType IN ('REVENUE', 'EXPENSE') AND (c.AccountCode Is Not Null OR p.AccountCode Is Not Null)
```

## IncomeStatementQuery

قائمة الدخل: الإيرادات والتكاليف والمصروفات ومجمل وصافي الربح، مع فترة المقارنة

المعاملات: `PeriodStart`, `PeriodEnd`, `CompareStart`, `CompareEnd`

```sql
SELECT q.SectionNo * 10 + 1 AS Block, q.TreeKey AS AccountKey, 'A' AS RowKind, q.AccountName AS Caption,
       q.AccountCode AS LineAccount, q.CurrentAmount AS CurrentValue, q.PriorAmount AS PriorValue
FROM qryIncomeAccounts AS q
UNION ALL
SELECT 10, '', 'H', 'إيرادات النشاط', Null, Null, Null
FROM Settings AS z WHERE z.SettingID = 1
UNION ALL
SELECT 12, '', 'T', 'صافي إيرادات النشاط', Null, CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.PriorAmount, 0)), 0))
FROM qryIncomeAccounts AS q
UNION ALL
SELECT 20, '', 'H', 'تكلفة المبيعات', Null, Null, Null
FROM Settings AS z WHERE z.SettingID = 1
UNION ALL
SELECT 22, '', 'T', 'إجمالي تكلفة المبيعات', Null, CCur(Nz(Sum(IIf(q.SectionNo = 2, 1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 2, 1 * q.PriorAmount, 0)), 0))
FROM qryIncomeAccounts AS q
UNION ALL
SELECT 30, '', 'H', 'المصروفات التشغيلية والإدارية', Null, Null, Null
FROM Settings AS z WHERE z.SettingID = 1
UNION ALL
SELECT 32, '', 'T', 'إجمالي المصروفات التشغيلية والإدارية', Null, CCur(Nz(Sum(IIf(q.SectionNo = 3, 1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 3, 1 * q.PriorAmount, 0)), 0))
FROM qryIncomeAccounts AS q
UNION ALL
SELECT 40, '', 'H', 'إيرادات أخرى', Null, Null, Null
FROM Settings AS z WHERE z.SettingID = 1
UNION ALL
SELECT 42, '', 'T', 'إجمالي الإيرادات الأخرى', Null, CCur(Nz(Sum(IIf(q.SectionNo = 4, 1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 4, 1 * q.PriorAmount, 0)), 0))
FROM qryIncomeAccounts AS q
UNION ALL
SELECT 50, '', 'H', 'مصروفات أخرى', Null, Null, Null
FROM Settings AS z WHERE z.SettingID = 1
UNION ALL
SELECT 52, '', 'T', 'إجمالي المصروفات الأخرى', Null, CCur(Nz(Sum(IIf(q.SectionNo = 5, 1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 5, 1 * q.PriorAmount, 0)), 0))
FROM qryIncomeAccounts AS q
UNION ALL
SELECT 25, '', 'R', 'مجمل الربح', Null, CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 2, -1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.PriorAmount, 0) + IIf(q.SectionNo = 2, -1 * q.PriorAmount, 0)), 0))
FROM qryIncomeAccounts AS q
UNION ALL
SELECT 35, '', 'R', 'الربح التشغيلي', Null, CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 2, -1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 3, -1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.PriorAmount, 0) + IIf(q.SectionNo = 2, -1 * q.PriorAmount, 0) + IIf(q.SectionNo = 3, -1 * q.PriorAmount, 0)), 0))
FROM qryIncomeAccounts AS q
UNION ALL
SELECT 60, '', 'R', 'صافي الربح (الخسارة)', Null, CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 2, -1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 3, -1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 4, 1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 5, -1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.PriorAmount, 0) + IIf(q.SectionNo = 2, -1 * q.PriorAmount, 0) + IIf(q.SectionNo = 3, -1 * q.PriorAmount, 0) + IIf(q.SectionNo = 4, 1 * q.PriorAmount, 0) + IIf(q.SectionNo = 5, -1 * q.PriorAmount, 0)), 0))
FROM qryIncomeAccounts AS q
ORDER BY Block, AccountKey
```

## qryBalanceAt

رصيد كل حساب في نهاية الفترة (مدين موجب)

المعاملات: `PeriodEnd`

```sql
SELECT l.AccountCode, Sum(l.Debit) - Sum(l.Credit) AS NetAt
FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID
WHERE e.EntryDate < QDate('PeriodEnd')
GROUP BY l.AccountCode
```

## qryBalanceCompare

رصيد كل حساب في نهاية فترة المقارنة (مدين موجب)

المعاملات: `CompareEnd`

```sql
SELECT l.AccountCode, Sum(l.Debit) - Sum(l.Credit) AS NetCompare
FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID
WHERE e.EntryDate < QDate('CompareEnd')
GROUP BY l.AccountCode
```

## qryBalanceAccounts

حسابات الميزانية: رصيد كل حساب أصول أو خصوم أو حقوق ملكية (بطبيعته موجب)

المعاملات: `PeriodEnd`, `CompareEnd`

```sql
SELECT a.AccountCode, a.AccountName, a.TreeKey, a.Level1Code, a.Level2Code,
       IIf(a.AccountType = 'ASSET', 1, -1) * CCur(Nz(b.NetAt, 0)) AS CurrentAmount,
       IIf(a.AccountType = 'ASSET', 1, -1) * CCur(Nz(c.NetCompare, 0)) AS PriorAmount
FROM (Accounts AS a LEFT JOIN qryBalanceAt AS b ON a.AccountCode = b.AccountCode)
     LEFT JOIN qryBalanceCompare AS c ON a.AccountCode = c.AccountCode
WHERE a.AccountType IN ('ASSET', 'LIABILITY', 'EQUITY') AND (CCur(Nz(b.NetAt, 0)) <> 0 OR CCur(Nz(c.NetCompare, 0)) <> 0)
```

## qryProfitAt

صافي ربح الفترات غير المقفلة حتى نهاية الفترة (مدين موجب)

المعاملات: `PeriodEnd`

```sql
SELECT CCur(Nz(Sum(b.NetAt), 0)) AS NetProfitSum
FROM qryBalanceAt AS b INNER JOIN Accounts AS a ON b.AccountCode = a.AccountCode
WHERE a.AccountType IN ('REVENUE', 'EXPENSE')
```

## qryProfitCompare

صافي ربح الفترات غير المقفلة حتى نهاية فترة المقارنة (مدين موجب)

المعاملات: `CompareEnd`

```sql
SELECT CCur(Nz(Sum(c.NetCompare), 0)) AS NetCompareSum
FROM qryBalanceCompare AS c INNER JOIN Accounts AS a ON c.AccountCode = a.AccountCode
WHERE a.AccountType IN ('REVENUE', 'EXPENSE')
```

## qryBalanceItems

بنود الميزانية بمجموعاتها، ومعها صافي الربح غير المقفل في الأرباح المحتجزة (32)

المعاملات: `PeriodEnd`, `CompareEnd`

```sql
SELECT q.Level1Code AS ClassNo, q.Level2Code AS GroupCode, q.CurrentAmount AS CurrentValue, q.PriorAmount AS PriorValue
FROM qryBalanceAccounts AS q
UNION ALL
SELECT 3, 32, -x.NetProfitSum, -y.NetCompareSum
FROM qryProfitAt AS x, qryProfitCompare AS y
```

## BalanceSheetQuery

الميزانية العمومية في نهاية الفترة: الأصول = الخصوم + حقوق الملكية، مع فترة المقارنة

المعاملات: `PeriodStart`, `PeriodEnd`, `CompareStart`, `CompareEnd`

```sql
SELECT q.Level1Code AS ClassNo, g.TreeKey AS GroupKey, 1 AS Pos, q.TreeKey AS AccountKey, 'A' AS RowKind,
       q.AccountName AS Caption, q.AccountCode AS LineAccount, q.CurrentAmount AS CurrentValue,
       q.PriorAmount AS PriorValue
FROM qryBalanceAccounts AS q INNER JOIN Accounts AS g ON q.Level2Code = g.AccountCode
UNION ALL
SELECT 3, g.TreeKey, 1, 'Z', 'A', 'صافي ربح (خسارة) الفترات غير المقفلة', Null, -x.NetProfitSum, -y.NetCompareSum
FROM Accounts AS g, qryProfitAt AS x, qryProfitCompare AS y
WHERE g.AccountCode = 32
UNION ALL
SELECT c.AccountCode, '', 0, '', 'C', c.AccountName, Null, Null, Null
FROM Accounts AS c
WHERE c.AccountCode IN (1, 2, 3)
UNION ALL
SELECT g.Level1Code, g.TreeKey, 0, '', 'G', g.AccountName, Null, Null, Null
FROM Accounts AS g
WHERE g.AccountCode IN (SELECT GroupCode FROM qryBalanceItems)
UNION ALL
SELECT g.Level1Code, g.TreeKey, 2, '', 'S', g.AccountName, Null, Sum(i.CurrentValue), Sum(i.PriorValue)
FROM Accounts AS g INNER JOIN qryBalanceItems AS i ON g.AccountCode = i.GroupCode
GROUP BY g.Level1Code, g.TreeKey, g.AccountName
UNION ALL
SELECT i.ClassNo, '~', 9, '', 'T', IIf(i.ClassNo = 1, 'إجمالي الأصول', IIf(i.ClassNo = 2, 'إجمالي الخصوم', 'إجمالي حقوق الملكية')), Null, Sum(i.CurrentValue), Sum(i.PriorValue)
FROM qryBalanceItems AS i
GROUP BY i.ClassNo
UNION ALL
SELECT 4, '', 9, '', 'T', 'إجمالي الخصوم وحقوق الملكية', Null, CCur(Nz(Sum(i.CurrentValue), 0)), CCur(Nz(Sum(i.PriorValue), 0))
FROM qryBalanceItems AS i
WHERE i.ClassNo IN (2, 3)
ORDER BY ClassNo, GroupKey, Pos, AccountKey
```

## AccountTreeQuery

شجرة الحسابات: كل حساب بمستواه ونوعه وهل يقبل القيود

```sql
SELECT a.AccountCode, a.AccountName, IIf(a.AccountType = 'ASSET', 'أصول', IIf(a.AccountType = 'LIABILITY', 'خصوم', IIf(a.AccountType = 'EQUITY', 'حقوق ملكية', IIf(a.AccountType = 'REVENUE', 'إيرادات', 'مصروفات')))) AS TypeName, a.AccountLevel, a.TreeKey,
       a.ParentCode, IIf(a.IsPosting, 'فرعي', 'رئيسي') AS KindName, a.IsPosting, a.IsActive
FROM Accounts AS a
ORDER BY a.TreeKey
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
