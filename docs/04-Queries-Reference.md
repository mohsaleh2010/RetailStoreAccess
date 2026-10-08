# مرجع الاستعلامات (Queries Reference)

> ملف مُولَّد تلقائيًا من `tools/queries.py` – لا تعدّله يدويًا.

عدد الاستعلامات: **197**. الاستعلامات التي تبدأ بـ `qry` مساعدة تستخدمها الاستعلامات الأخرى؛ البقية تُستخدم مباشرة في التقارير والنماذج. ⭐ = مطلوب بالاسم في البرومبت.

| # | الاستعلام | الوصف | المعاملات |
|---|---|---|---|
| 1 | [`qryLocAccounts1`](#qrylocaccounts1) | Accounts للواجهة الإنجليزية (خطوة 1) |  |
| 2 | [`qryLocAccounts`](#qrylocaccounts) | Accounts بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 3 | [`qryLocPaymentMethods1`](#qrylocpaymentmethods1) | PaymentMethods للواجهة الإنجليزية (خطوة 1) |  |
| 4 | [`qryLocPaymentMethods`](#qrylocpaymentmethods) | PaymentMethods بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 5 | [`qryLocJournalSourceTypes1`](#qrylocjournalsourcetypes1) | JournalSourceTypes للواجهة الإنجليزية (خطوة 1) |  |
| 6 | [`qryLocJournalSourceTypes`](#qrylocjournalsourcetypes) | JournalSourceTypes بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 7 | [`qryLocTransactionTypes1`](#qryloctransactiontypes1) | TransactionTypes للواجهة الإنجليزية (خطوة 1) |  |
| 8 | [`qryLocTransactionTypes`](#qryloctransactiontypes) | TransactionTypes بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 9 | [`qryLocRoles1`](#qrylocroles1) | Roles للواجهة الإنجليزية (خطوة 1) |  |
| 10 | [`qryLocRoles`](#qrylocroles) | Roles بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 11 | [`qryLocPermissions1`](#qrylocpermissions1) | Permissions للواجهة الإنجليزية (خطوة 1) |  |
| 12 | [`qryLocPermissions`](#qrylocpermissions) | Permissions بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 13 | [`qryLocScreens1`](#qrylocscreens1) | Screens للواجهة الإنجليزية (خطوة 1) |  |
| 14 | [`qryLocScreens`](#qrylocscreens) | Screens بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 15 | [`qryLocCategories1`](#qryloccategories1) | Categories للواجهة الإنجليزية (خطوة 1) |  |
| 16 | [`qryLocCategories`](#qryloccategories) | Categories بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 17 | [`qryLocUnits1`](#qrylocunits1) | Units للواجهة الإنجليزية (خطوة 1) |  |
| 18 | [`qryLocUnits`](#qrylocunits) | Units بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 19 | [`qryLocExpenseTypes1`](#qrylocexpensetypes1) | ExpenseTypes للواجهة الإنجليزية (خطوة 1) |  |
| 20 | [`qryLocExpenseTypes`](#qrylocexpensetypes) | ExpenseTypes بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 21 | [`qryLocCustomers1`](#qryloccustomers1) | Customers للواجهة الإنجليزية (خطوة 1) |  |
| 22 | [`qryLocCustomers`](#qryloccustomers) | Customers بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 23 | [`qryLocSuppliers1`](#qrylocsuppliers1) | Suppliers للواجهة الإنجليزية (خطوة 1) |  |
| 24 | [`qryLocSuppliers`](#qrylocsuppliers) | Suppliers بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 25 | [`qryLocCashBoxes1`](#qryloccashboxes1) | CashBoxes للواجهة الإنجليزية (خطوة 1) |  |
| 26 | [`qryLocCashBoxes`](#qryloccashboxes) | CashBoxes بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 27 | [`qryLocBanks1`](#qrylocbanks1) | Banks للواجهة الإنجليزية (خطوة 1) |  |
| 28 | [`qryLocBanks`](#qrylocbanks) | Banks بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 29 | [`qryLocCostCenters1`](#qryloccostcenters1) | CostCenters للواجهة الإنجليزية (خطوة 1) |  |
| 30 | [`qryLocCostCenters`](#qryloccostcenters) | CostCenters بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 31 | [`qryLocSalesReps1`](#qrylocsalesreps1) | SalesReps للواجهة الإنجليزية (خطوة 1) |  |
| 32 | [`qryLocSalesReps`](#qrylocsalesreps) | SalesReps بالأسماء الإنجليزية (الواجهة الإنجليزية) |  |
| 33 | [`qrySalesDocuments`](#qrysalesdocuments) | مستندات البيع: الفواتير (+) والمرتجعات (−) بقيم موقّعة |  |
| 34 | [`qrySalesLineItems`](#qrysaleslineitems) | أسطر البيع والمرتجعات مع التكلفة (أساس تحليل المنتجات والأرباح) |  |
| 35 | [`qrySalesLinesInPeriod`](#qrysaleslinesinperiod) | أسطر البيع والمرتجعات داخل الفترة | `PeriodStart`, `PeriodEnd` |
| 36 | [`DailySalesQuery`](#dailysalesquery) ⭐ | المبيعات اليومية: عدد الفواتير والمرتجعات والصافي والنقدي والآجل |  |
| 37 | [`qrySalesMonthlyDocs`](#qrysalesmonthlydocs) | تجميع شهري لمستندات البيع |  |
| 38 | [`qrySalesMonthlyCost`](#qrysalesmonthlycost) | تكلفة البضاعة المباعة شهريًا |  |
| 39 | [`MonthlySalesQuery`](#monthlysalesquery) ⭐ | المبيعات الشهرية مع التكلفة ومجمل الربح |  |
| 40 | [`SalesByPeriodQuery`](#salesbyperiodquery) | فواتير ومرتجعات البيع خلال فترة مع العميل والكاشير | `PeriodStart`, `PeriodEnd` |
| 41 | [`SalesByProductQuery`](#salesbyproductquery) | المبيعات حسب المنتج خلال فترة (كمية، صافي، تكلفة، ربح) | `PeriodStart`, `PeriodEnd` |
| 42 | [`BestSellingProductsQuery`](#bestsellingproductsquery) ⭐ | أفضل المنتجات مبيعًا خلال فترة (حسب صافي الكمية) | `PeriodStart`, `PeriodEnd` |
| 43 | [`SalesByCategoryQuery`](#salesbycategoryquery) | المبيعات حسب التصنيف خلال فترة (للرسم الدائري) | `PeriodStart`, `PeriodEnd` |
| 44 | [`LeastSellingProductsQuery`](#leastsellingproductsquery) | أقل المنتجات مبيعًا خلال فترة (تشمل المنتجات التي لم تُبع) | `PeriodStart`, `PeriodEnd` |
| 45 | [`qryPurchaseDocuments`](#qrypurchasedocuments) | مستندات الشراء: الفواتير (+) والمرتجعات (−) |  |
| 46 | [`PurchasesQuery`](#purchasesquery) | فواتير ومرتجعات الشراء خلال فترة مع المورد | `PeriodStart`, `PeriodEnd` |
| 47 | [`qryProductLedger`](#qryproductledger) | رصيد كل منتج من دفتر حركة المخزون |  |
| 48 | [`qryProductLastSale`](#qryproductlastsale) | تاريخ آخر بيع لكل منتج |  |
| 49 | [`StockBalanceQuery`](#stockbalancequery) ⭐ | المخزون الحالي: الكمية والقيمة بالتكلفة وبسعر البيع ومطابقتها مع الحركات |  |
| 50 | [`LowStockQuery`](#lowstockquery) ⭐ | المنتجات منخفضة المخزون: CurrentQuantity <= MinimumQuantity |  |
| 51 | [`ProductMovementQuery`](#productmovementquery) | حركة منتج خلال فترة مع رصيد أول المدة (الرصيد التراكمي في التقرير) | `PeriodStart`, `PeriodEnd`, `ProductID` |
| 52 | [`SlowMovingProductsQuery`](#slowmovingproductsquery) ⭐ | المنتجات غير المتحركة: لها رصيد ولم تُبع منذ عدد الأيام المحدد في الإعدادات |  |
| 53 | [`StockByCategoryQuery`](#stockbycategoryquery) | المخزون حسب التصنيف: عدد المنتجات والكمية والقيمة |  |
| 54 | [`StockCountQuery`](#stockcountquery) | تفاصيل جلسات الجرد: الكمية المسجلة والفعلية والفرق وقيمته |  |
| 55 | [`qryCustomerLedger`](#qrycustomerledger) | دفتر حساب العملاء: مدين (عليه) / دائن (له) |  |
| 56 | [`qryCustomerLedgerTotals`](#qrycustomerledgertotals) | مجاميع حساب كل عميل |  |
| 57 | [`CustomerBalanceQuery`](#customerbalancequery) ⭐ | رصيد كل عميل محسوبًا من الحركات (موجب = عليه للمحل) |  |
| 58 | [`CustomersWithDebtQuery`](#customerswithdebtquery) | العملاء الذين عليهم مبالغ مستحقة |  |
| 59 | [`CustomerStatementQuery`](#customerstatementquery) | كشف حساب عميل لفترة: رصيد سابق ثم الحركات | `PeriodStart`, `PeriodEnd`, `CustomerID` |
| 60 | [`qrySupplierLedger`](#qrysupplierledger) | دفتر حساب الموردين: دائن (للمورد) / مدين (سُدِّد له) |  |
| 61 | [`qrySupplierLedgerTotals`](#qrysupplierledgertotals) | مجاميع حساب كل مورد |  |
| 62 | [`SupplierBalanceQuery`](#supplierbalancequery) ⭐ | رصيد كل مورد محسوبًا من الحركات (موجب = مستحق للمورد) |  |
| 63 | [`qrySupplierFxMoves`](#qrysupplierfxmoves) | حركات أرصدة الموردين بعملة كل مستند (الرصيد الافتتاحي والشيكات بعملة البرنامج) |  |
| 64 | [`qryLatestRateDates`](#qrylatestratedates) | تاريخ آخر سعر لكل عملة |  |
| 65 | [`qryLatestRates`](#qrylatestrates) | آخر معامل لكل عملة |  |
| 66 | [`qrySupplierFxTotals`](#qrysupplierfxtotals) | رصيد كل مورد بكل عملة |  |
| 67 | [`SupplierFxBalanceQuery`](#supplierfxbalancequery) | أرصدة الموردين بالعملات الأجنبية: بالدفاتر، وبآخر سعر، وفرق العملة غير المحقق |  |
| 68 | [`SupplierStatementQuery`](#supplierstatementquery) | كشف حساب مورد لفترة: رصيد سابق ثم الحركات | `PeriodStart`, `PeriodEnd`, `SupplierID` |
| 69 | [`qryCustomerAllocSums`](#qrycustomerallocsums) | مجموع ما رُبط من كل سند بالفواتير |  |
| 70 | [`qryCustomerPaymentFree`](#qrycustomerpaymentfree) | سندات القبض: المربوط بالفواتير والباقي غير المربوط |  |
| 71 | [`qryCustomerInvoiceAlloc`](#qrycustomerinvoicealloc) | مجموع ما رُبط بكل فاتورة من السندات |  |
| 72 | [`qryCustomerInvoiceReturns`](#qrycustomerinvoicereturns) | مرتجعات كل فاتورة المخصومة من رصيد الحساب |  |
| 73 | [`qryCustomerInvoiceFree`](#qrycustomerinvoicefree) | الفواتير الآجلة: المتبقي وما رُبط بها وما يمكن ربطه |  |
| 74 | [`qrySupplierAllocSums`](#qrysupplierallocsums) | مجموع ما رُبط من كل سند بالفواتير |  |
| 75 | [`qrySupplierPaymentFree`](#qrysupplierpaymentfree) | سندات الصرف: المربوط بالفواتير والباقي غير المربوط |  |
| 76 | [`qrySupplierInvoiceAlloc`](#qrysupplierinvoicealloc) | مجموع ما رُبط بكل فاتورة من السندات |  |
| 77 | [`qrySupplierInvoiceReturns`](#qrysupplierinvoicereturns) | مرتجعات كل فاتورة المخصومة من رصيد الحساب |  |
| 78 | [`qrySupplierInvoiceFree`](#qrysupplierinvoicefree) | الفواتير الآجلة: المتبقي وما رُبط بها وما يمكن ربطه |  |
| 79 | [`qryAgingDebits`](#qryagingdebits) | المستحق على كل عميل وللمورد بالمستند: الفواتير الآجلة والرصيد الافتتاحي |  |
| 80 | [`qryAgingCredits`](#qryagingcredits) | ما يسدد المستحق: المرتجعات (على فاتورتها أولًا) والسندات والرصيد الافتتاحي الدائن |  |
| 81 | [`qryAgingAllocations`](#qryagingallocations) | ربط السندات بالفواتير (ومنها السند المسجل عن فاتورة قبل الربط) |  |
| 82 | [`ExpensesQuery`](#expensesquery) | المصروفات خلال فترة | `PeriodStart`, `PeriodEnd` |
| 83 | [`ExpensesByTypeQuery`](#expensesbytypequery) | المصروفات مجمّعة حسب النوع خلال فترة | `PeriodStart`, `PeriodEnd` |
| 84 | [`qryProfitSales`](#qryprofitsales) | صافي المبيعات وتكلفتها خلال الفترة | `PeriodStart`, `PeriodEnd` |
| 85 | [`qryProfitAdjustments`](#qryprofitadjustments) | قيمة فروقات المخزون (جرد، إضافة، خصم) خلال الفترة | `PeriodStart`, `PeriodEnd` |
| 86 | [`qryProfitExpenses`](#qryprofitexpenses) | المصروفات (بدون ضريبة) خلال الفترة | `PeriodStart`, `PeriodEnd` |
| 87 | [`ProfitQuery`](#profitquery) ⭐ | الأرباح: صافي المبيعات − التكلفة = مجمل الربح؛ ثم ± فروقات المخزون − المصروفات = صافي الربح | `PeriodStart`, `PeriodEnd` |
| 88 | [`qryVatOutput`](#qryvatoutput) | ضريبة المخرجات (المبيعات ناقص المرتجعات) | `PeriodStart`, `PeriodEnd` |
| 89 | [`qryVatInputPurchases`](#qryvatinputpurchases) | ضريبة المدخلات من المشتريات (ناقص المرتجعات) | `PeriodStart`, `PeriodEnd` |
| 90 | [`qryVatInputExpenses`](#qryvatinputexpenses) | ضريبة المدخلات من المصروفات | `PeriodStart`, `PeriodEnd` |
| 91 | [`VatSummaryQuery`](#vatsummaryquery) | ملخص ضريبة القيمة المضافة للفترة (للإقرار الضريبي) | `PeriodStart`, `PeriodEnd` |
| 92 | [`qryVatReturnLines`](#qryvatreturnlines) | أسطر الإقرار الضريبي: المبيعات والمشتريات والمصروفات بفئتها الضريبية |  |
| 93 | [`qryVatReturnTotals`](#qryvatreturntotals) | خانات الإقرار الضريبي للفترة محسوبة من المستندات (صف واحد) | `PeriodStart`, `PeriodEnd` |
| 94 | [`qryVatReturnHead`](#qryvatreturnhead) | الإقرار الضريبي المختار | `VatReturnID` |
| 95 | [`VatReturnQuery`](#vatreturnquery) | إقرار ضريبة القيمة المضافة بخانات نموذج الهيئة (1 إلى 16) | `VatReturnID` |
| 96 | [`DashboardQuery`](#dashboardquery) | مؤشرات لوحة التحكم في سجل واحد (اليوم، الشهر، الأرصدة، المخزون) | `DashDay`, `DashMonth`, `DashEnd` |
| 97 | [`qryRepDocs`](#qryrepdocs) | عمليات المندوبين: صافي المبيعات بدون الضريبة والتحصيل |  |
| 98 | [`qryRepPeriodTotals`](#qryrepperiodtotals) | مبيعات وتحصيل كل مندوب في الفترة | `PeriodStart`, `PeriodEnd` |
| 99 | [`qryRepTargetTotals`](#qryreptargettotals) | أهداف كل مندوب في أشهر الفترة | `PeriodStart`, `PeriodEnd` |
| 100 | [`qryRepCommissionPaid`](#qryrepcommissionpaid) | ما صُرف لكل مندوب من عمولاته (سندات صرف النقدية) |  |
| 101 | [`qryRepCommissionPosted`](#qryrepcommissionposted) | العمولات المرحَّلة لكل مندوب |  |
| 102 | [`RepPerformanceQuery`](#repperformancequery) | أداء المندوبين في الفترة: المبيعات والتحصيل والهدف والإنجاز والعمولة المتوقعة | `PeriodStart`, `PeriodEnd` |
| 103 | [`RepCustomersQuery`](#repcustomersquery) | عملاء كل مندوب وأرصدتهم |  |
| 104 | [`RepCommissionBalanceQuery`](#repcommissionbalancequery) | العمولات المستحقة لكل مندوب: المرحَّل والمصروف والباقي |  |
| 105 | [`CommissionSheetQuery`](#commissionsheetquery) | مسير العمولات المختار بأسطر المندوبين | `CommissionRunID` |
| 106 | [`qryIndicatorLines`](#qryindicatorlines) | أسطر القيود مع مجموعة الحساب (المستوى 2) لحساب المؤشرات المالية |  |
| 107 | [`FinancialIndicatorsQuery`](#financialindicatorsquery) | المؤشرات المالية (صف واحد): هامش الربح، دوران المخزون، فترة التحصيل، السيولة | `IndEnd`, `IndMonth`, `IndPrevMonth`, `IndYear`, `Ind90` |
| 108 | [`qryDashboardTopProducts`](#qrydashboardtopproducts) | صافي الكمية المباعة لكل منتج منذ بداية الشهر (لوحة التحكم) | `DashMonth`, `DashEnd` |
| 109 | [`qryEInvoiceDocs`](#qryeinvoicedocs) | مستندات البيع وحالة الفاتورة الإلكترونية (الفواتير والمرتجعات) |  |
| 110 | [`qrySalesDocPrint`](#qrysalesdocprint) | بيانات طباعة فواتير البيع والإشعارات الدائنة (سطر لكل صنف) |  |
| 111 | [`qryPurchaseDocPrint`](#qrypurchasedocprint) | بيانات طباعة فواتير الشراء ومرتجعاتها (سطر لكل صنف) |  |
| 112 | [`qryVoucherPrint`](#qryvoucherprint) | بيانات طباعة سندات القبض (من العملاء) وسندات الصرف (للموردين) |  |
| 113 | [`qryCashMovements`](#qrycashmovements) | كل حركات النقدية في الخزينة والصناديق: داخل (+) وخارج (−) |  |
| 114 | [`qryCashBoxTotals`](#qrycashboxtotals) | إجمالي الداخل والخارج لكل صندوق |  |
| 115 | [`CashBoxBalanceQuery`](#cashboxbalancequery) | أرصدة الخزينة والصناديق الآن |  |
| 116 | [`CashStatementQuery`](#cashstatementquery) | حركة الخزينة / الصندوق لفترة: رصيد أول المدة ثم الحركات (0 = كل الصناديق) | `PeriodStart`, `PeriodEnd`, `CashBoxID` |
| 117 | [`qryCashDays`](#qrycashdays) | مقبوضات ومدفوعات كل يوم داخل الفترة | `PeriodStart`, `PeriodEnd`, `CashBoxID` |
| 118 | [`qryCashDayOpening`](#qrycashdayopening) | رصيد أول كل يوم من أيام الحركة (كل الحركات قبل ذلك اليوم) | `PeriodStart`, `PeriodEnd`, `CashBoxID` |
| 119 | [`CashDailyQuery`](#cashdailyquery) | حركة الخزينة اليومية: رصيد أول اليوم والمقبوضات والمدفوعات ورصيد آخر اليوم | `PeriodStart`, `PeriodEnd`, `CashBoxID` |
| 120 | [`CashClosingsQuery`](#cashclosingsquery) | تصفيات يومية الكاشير خلال فترة (0 = كل الصناديق) | `PeriodStart`, `PeriodEnd`, `CashBoxID` |
| 121 | [`qryCashClosingPrint`](#qrycashclosingprint) | بيانات طباعة تصفية الكاشير |  |
| 122 | [`qryCashVoucherPrint`](#qrycashvoucherprint) | بيانات طباعة سندات قبض وصرف وتحويل النقدية |  |
| 123 | [`qrySaleCost`](#qrysalecost) | تكلفة كل فاتورة بيع |  |
| 124 | [`qryReturnCost`](#qryreturncost) | تكلفة ما عاد للمخزون من كل مرتجع بيع |  |
| 125 | [`qryStockCountValue`](#qrystockcountvalue) | قيمة فروقات كل جرد مُرحّل |  |
| 126 | [`qryJournalSale`](#qryjournalsale) | أسطر قيود فواتير البيع |  |
| 127 | [`qryJournalSalesReturn`](#qryjournalsalesreturn) | أسطر قيود مرتجعات البيع |  |
| 128 | [`qryJournalPurchase`](#qryjournalpurchase) | أسطر قيود فواتير الشراء |  |
| 129 | [`qryJournalPurchaseReturn`](#qryjournalpurchasereturn) | أسطر قيود مرتجعات الشراء |  |
| 130 | [`qryJournalPayments`](#qryjournalpayments) | أسطر قيود سندات القبض من العملاء والصرف للموردين |  |
| 131 | [`qryJournalExpense`](#qryjournalexpense) | أسطر قيود المصروفات (عدا المسجلة بسند نقدية) |  |
| 132 | [`qryJournalCashVoucher`](#qryjournalcashvoucher) | أسطر قيود سندات النقدية (قبض وصرف وتحويل) |  |
| 133 | [`qryJournalStock`](#qryjournalstock) | أسطر قيود حركات المخزون اليدوية وتسويات الجرد |  |
| 134 | [`qryJournalOpening`](#qryjournalopening) | أسطر قيود الأرصدة الافتتاحية للصناديق والعملاء والموردين |  |
| 135 | [`qryManualEntryLines`](#qrymanualentrylines) | أسطر القيود اليدوية مع رأس كل قيد |  |
| 136 | [`qryJournalManual`](#qryjournalmanual) | أسطر القيود اليدوية |  |
| 137 | [`qryYearCloseLines`](#qryyearcloselines) | أسطر قيود إقفال السنوات مع رأس كل إقفال |  |
| 138 | [`qryJournalYearClose`](#qryjournalyearclose) | أسطر قيود إقفال السنوات: الإيرادات والمصروفات إلى الأرباح المحتجزة |  |
| 139 | [`qryJournalVatReturn`](#qryjournalvatreturn) | أسطر قيود الإقرار الضريبي المعتمد (التسوية) وسداده |  |
| 140 | [`qryJournalCheque`](#qryjournalcheque) | أسطر قيود الشيكات: الاستلام أو الإصدار، ثم التحصيل أو الارتداد |  |
| 141 | [`qryJournalAsset`](#qryjournalasset) | أسطر قيود اقتناء الأصول الثابتة وبيعها أو استبعادها |  |
| 142 | [`qryDepreciationLines`](#qrydepreciationlines) | أسطر قيود الإهلاك الشهرية مع اسم الأصل |  |
| 143 | [`qryJournalDepreciation`](#qryjournaldepreciation) | أسطر قيود الإهلاك الشهرية: مصروف الإهلاك ومجمع الإهلاك لكل أصل |  |
| 144 | [`qryPayrollTotals`](#qrypayrolltotals) | مجاميع كل مسير رواتب لقيده |  |
| 145 | [`qryPayrollCenterTotals`](#qrypayrollcentertotals) | مجاميع كل مسير رواتب لكل مركز تكلفة لقيده |  |
| 146 | [`qryJournalPayroll`](#qryjournalpayroll) | أسطر قيود مسيرات الرواتب المرحَّلة وصرفها |  |
| 147 | [`qryCommissionCenterTotals`](#qrycommissioncentertotals) | عمولات كل مسير لكل مركز تكلفة لقيده |  |
| 148 | [`qryJournalCommission`](#qryjournalcommission) | أسطر قيود مسيرات العمولات المرحَّلة |  |
| 149 | [`qryJournalBankTx`](#qryjournalbanktx) | أسطر قيود الحركات البنكية: الإيداع والسحب وتسوية مدى والتحويل والحركات الأخرى |  |
| 150 | [`qryBankItemSums`](#qrybankitemsums) | صافي كل عملية على حساب كل بنك في القيود |  |
| 151 | [`qryBankItems`](#qrybankitems) | عمليات البنوك: المبلغ، وهل طابقت كشف البنك ومبلغها يوم المطابقة |  |
| 152 | [`qryBankTotals`](#qrybanktotals) | رصيد كل بنك في الدفاتر |  |
| 153 | [`BankBalanceQuery`](#bankbalancequery) | أرصدة البنوك في الدفاتر |  |
| 154 | [`qryAssetDepTotals`](#qryassetdeptotals) | مجموع إهلاك كل أصل في القيود الشهرية |  |
| 155 | [`FixedAssetsQuery`](#fixedassetsquery) | سجل الأصول الثابتة: التكلفة ومجمع الإهلاك والقيمة الدفترية والقسط الشهري |  |
| 156 | [`AuditTrailQuery`](#audittrailquery) | سجل التدقيق: كل عملية بمن قام بها ووقتها، وحقولها بالقيمة قبل وبعد |  |
| 157 | [`qryAdvanceMoves`](#qryadvancemoves) | حركات سلف الموظفين: الصرف والسداد النقدي والخصم من الرواتب |  |
| 158 | [`qryAdvanceTotals`](#qryadvancetotals) | رصيد سلف كل موظف |  |
| 159 | [`AdvanceBalanceQuery`](#advancebalancequery) | أرصدة سلف الموظفين |  |
| 160 | [`PayrollSheetQuery`](#payrollsheetquery) | مسير الرواتب المختار بأسطر الموظفين | `PayrollRunID` |
| 161 | [`ChequesQuery`](#chequesquery) | الشيكات الواردة والصادرة مع العميل أو المورد وحالتها |  |
| 162 | [`JournalLinesQuery`](#journallinesquery) | قيود اليومية خلال فترة بأسطرها | `PeriodStart`, `PeriodEnd` |
| 163 | [`qryJournalEntryPrint`](#qryjournalentryprint) | بيانات طباعة قيد |  |
| 164 | [`qryTrialBefore`](#qrytrialbefore) | مجموع الحسابات قبل الفترة | `PeriodStart` |
| 165 | [`qryTrialPeriod`](#qrytrialperiod) | حركة الحسابات خلال الفترة | `PeriodStart`, `PeriodEnd` |
| 166 | [`TrialBalanceQuery`](#trialbalancequery) | ميزان المراجعة: رصيد أول المدة وحركة الفترة والرصيد الختامي (المدين موجب) | `PeriodStart`, `PeriodEnd` |
| 167 | [`qryStatementBefore`](#qrystatementbefore) | رصيد الحساب المختار (مع حساباته التابعة) قبل بداية الفترة | `PeriodStart`, `AccountCode` |
| 168 | [`AccountStatementQuery`](#accountstatementquery) | كشف حساب لفترة: رصيد أول المدة ثم كل سطر قيد (الحساب الرئيسي يشمل حساباته التابعة) | `PeriodStart`, `PeriodEnd`, `AccountCode` |
| 169 | [`GeneralLedgerQuery`](#generalledgerquery) | دفتر الأستاذ لفترة: لكل حساب فرعي رصيد أول المدة ثم أسطر قيوده (0 = كل الحسابات) | `PeriodStart`, `PeriodEnd`, `AccountCode` |
| 170 | [`qryTreeRollup`](#qrytreerollup) | أرصدة ميزان المراجعة مجمّعة على كل مستوى من شجرة الحسابات | `PeriodStart`, `PeriodEnd` |
| 171 | [`TrialBalanceTreeQuery`](#trialbalancetreequery) | ميزان المراجعة بالمستويات: كل حساب رئيسي بمجموع حساباته التابعة | `PeriodStart`, `PeriodEnd` |
| 172 | [`qryIncomeMoves`](#qryincomemoves) | حركة الحسابات في الفترة بدون قيود إقفال السنة | `PeriodStart`, `PeriodEnd` |
| 173 | [`qryCompareMoves`](#qrycomparemoves) | حركة الحسابات في فترة المقارنة بدون قيود إقفال السنة | `CompareStart`, `CompareEnd` |
| 174 | [`qryIncomeAccounts`](#qryincomeaccounts) | حسابات قائمة الدخل: صافي حركة كل حساب إيرادات أو مصروفات في الفترة وفترة المقارنة | `PeriodStart`, `PeriodEnd`, `CompareStart`, `CompareEnd` |
| 175 | [`IncomeStatementQuery`](#incomestatementquery) | قائمة الدخل: الإيرادات والتكاليف والمصروفات ومجمل وصافي الربح، مع فترة المقارنة | `PeriodStart`, `PeriodEnd`, `CompareStart`, `CompareEnd` |
| 176 | [`qryCenterMoves`](#qrycentermoves) | صافي حركة كل حساب إيرادات أو مصروفات لكل مركز تكلفة في الفترة | `PeriodStart`, `PeriodEnd` |
| 177 | [`qryCenterNames`](#qrycenternames) | مراكز التكلفة ومعها «غير موزع» |  |
| 178 | [`qryCenterSums`](#qrycentersums) | الإيرادات وتكلفة المبيعات والمصروفات لكل مركز تكلفة |  |
| 179 | [`CostCenterProfitQuery`](#costcenterprofitquery) | قائمة الدخل لكل مركز تكلفة: الإيرادات، تكلفة المبيعات، مجمل الربح، المصروفات، صافي الربح | `PeriodStart`, `PeriodEnd` |
| 180 | [`CostCenterAccountsQuery`](#costcenteraccountsquery) | إيرادات ومصروفات كل مركز تكلفة بالحسابات | `PeriodStart`, `PeriodEnd` |
| 181 | [`qryBudgetMonths`](#qrybudgetmonths) | أشهر الموازنة: سطر لكل شهر من كل سطر موازنة |  |
| 182 | [`qryBudgetPlanned`](#qrybudgetplanned) | مبلغ الموازنة لكل سطر في أشهر الفترة | `PeriodStart`, `PeriodEnd` |
| 183 | [`qryBudgetActual`](#qrybudgetactual) | الفعلي لكل سطر موازنة في الفترة من القيود | `PeriodStart`, `PeriodEnd` |
| 184 | [`BudgetVsActualQuery`](#budgetvsactualquery) | الموازنة مقابل الفعلي في الفترة: الانحراف ونسبته، وهل هو ملائم | `PeriodStart`, `PeriodEnd` |
| 185 | [`qryBalanceAt`](#qrybalanceat) | رصيد كل حساب في نهاية الفترة (مدين موجب) | `PeriodEnd` |
| 186 | [`qryBalanceCompare`](#qrybalancecompare) | رصيد كل حساب في نهاية فترة المقارنة (مدين موجب) | `CompareEnd` |
| 187 | [`qryBalanceAccounts`](#qrybalanceaccounts) | حسابات الميزانية: رصيد كل حساب أصول أو خصوم أو حقوق ملكية (بطبيعته موجب) | `PeriodEnd`, `CompareEnd` |
| 188 | [`qryProfitAt`](#qryprofitat) | صافي ربح الفترات غير المقفلة حتى نهاية الفترة (مدين موجب) | `PeriodEnd` |
| 189 | [`qryProfitCompare`](#qryprofitcompare) | صافي ربح الفترات غير المقفلة حتى نهاية فترة المقارنة (مدين موجب) | `CompareEnd` |
| 190 | [`qryBalanceItems`](#qrybalanceitems) | بنود الميزانية بمجموعاتها، ومعها صافي الربح غير المقفل في الأرباح المحتجزة (32) | `PeriodEnd`, `CompareEnd` |
| 191 | [`BalanceSheetQuery`](#balancesheetquery) | الميزانية العمومية في نهاية الفترة: الأصول = الخصوم + حقوق الملكية، مع فترة المقارنة | `PeriodStart`, `PeriodEnd`, `CompareStart`, `CompareEnd` |
| 192 | [`AccountTreeQuery`](#accounttreequery) | شجرة الحسابات: كل حساب بمستواه ونوعه وهل يقبل القيود |  |
| 193 | [`qrySalesInvoiceLineTotals`](#qrysalesinvoicelinetotals) | مجموع أسطر كل فاتورة بيع |  |
| 194 | [`qryPurchaseInvoiceLineTotals`](#qrypurchaseinvoicelinetotals) | مجموع أسطر كل فاتورة شراء |  |
| 195 | [`qrySalesReturnedQty`](#qrysalesreturnedqty) | الكمية المرتجعة من كل سطر فاتورة بيع |  |
| 196 | [`qryPurchaseReturnedQty`](#qrypurchasereturnedqty) | الكمية المرتجعة للمورد من كل سطر فاتورة شراء |  |
| 197 | [`IntegrityCheckQuery`](#integritycheckquery) | فحص سلامة البيانات: أي سطر هنا مشكلة يجب مراجعتها (النتيجة الفارغة = سليم) |  |

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
| 61 | الإقرار الضريبي: ضريبة المبيعات الخاضعة (الخانة 1) = ضريبة المخرجات | `SELECT SalesStdVAT FROM qryVatReturnTotals` | 204 |
| 62 | الإقرار الضريبي: ضريبة المشتريات والمصروفات (الخانة 7) = ضريبة المدخلات | `SELECT PurchStdVAT FROM qryVatReturnTotals` | -15 |
| 63 | الإقرار الضريبي: صافي المبيعات الخاضعة (الخانة 1 مع التعديلات) = صافي المبيعات | `SELECT SalesStdAmount + SalesStdAdjust + SalesZeroAmount + SalesZeroAdjust + SalesExemptAmount + SalesExemptAdjust FROM qryVatReturnTotals` | 1360 |
| 64 | طباعة الفاتورة الآجلة: سطران | `SELECT COUNT(*) FROM qrySalesDocPrint WHERE DocKind = 'SALE' AND DocID = {ref:INV2}` | 2 |
| 65 | طباعة الفاتورة الآجلة: مجموع الأسطر = 460 | `SELECT Sum(LineTotal) FROM qrySalesDocPrint WHERE DocKind = 'SALE' AND DocID = {ref:INV2}` | 460 |
| 66 | طباعة الإشعار الدائن: سطر واحد بقيمة 46 | `SELECT Sum(LineTotal) FROM qrySalesDocPrint WHERE DocKind = 'RETURN' AND DocID = {ref:CRN1}` | 46 |
| 67 | رصيد صندوق الكاشير = 500 + 115 + 1150 + 100 + 200 − 230 − 50 − 15 − 1000 | `SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = {ref:BOXC}` | 770 |
| 68 | رصيد الخزينة = 10000 − 5000 + 1000 + 2000 − 300 | `SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = {ref:BOXM}` | 7700 |
| 69 | الصناديق المسجلة بدون حركة رصيدها صفر | `SELECT Sum(Balance) FROM CashBoxBalanceQuery WHERE CashBoxID <= 2` | 0 |
| 70 | حركة صندوق الكاشير: رصيد أول المدة = 500 + 115 − 50 | `SELECT AmountIn FROM CashStatementQuery WHERE SortKey = 0` | 565 |
| 71 | حركة صندوق الكاشير: 6 حركات في الفترة | `SELECT COUNT(*) FROM CashStatementQuery WHERE SortKey = 1` | 6 |
| 72 | حركة صندوق الكاشير: رصيد آخر المدة = 770 | `SELECT Sum(AmountIn) - Sum(AmountOut) FROM CashStatementQuery` | 770 |
| 73 | حركة الخزينة: التحويل من الكاشير داخل = 1000 | `SELECT AmountIn FROM CashStatementQuery WHERE MoveType = 'TRANSFER_IN'` | 1000 |
| 74 | يومية صندوق الكاشير يوم التصفية: رصيد أول اليوم = 565 + 1150 + 100 − 230 | `SELECT OpeningBalance FROM CashDailyQuery WHERE CashDay = DateValue({day:4})` | 1585 |
| 75 | يومية صندوق الكاشير يوم التصفية: المدفوعات = 15 عجز + 1000 تحويل | `SELECT Payments FROM CashDailyQuery WHERE CashDay = DateValue({day:4})` | 1015 |
| 76 | يومية صندوق الكاشير يوم التصفية: رصيد آخر اليوم = 570 | `SELECT ClosingBalance FROM CashDailyQuery WHERE CashDay = DateValue({day:4})` | 570 |
| 77 | يومية صندوق الكاشير: آخر يوم = الرصيد الحالي | `SELECT ClosingBalance FROM CashDailyQuery WHERE CashDay = DateValue({day:3})` | 770 |
| 78 | تصفيات الكاشير خلال الفترة: تصفية واحدة بعجز 15 | `SELECT Difference FROM CashClosingsQuery` | -15 |
| 79 | طباعة سند صرف المصروف: نوع المصروف | `SELECT COUNT(*) FROM qryCashVoucherPrint WHERE DocID = {ref:V1} AND ExpenseTypeName = 'مصروفات أخرى'` | 1 |
| 80 | طباعة سند التحويل: الصندوق المستلم | `SELECT COUNT(*) FROM qryCashVoucherPrint WHERE DocID = {ref:V3} AND ToBoxName = 'TEST الخزينة'` | 1 |
| 81 | قيود Sale: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalSale GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 82 | قيود SalesReturn: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalSalesReturn GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 83 | قيود Purchase: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalPurchase GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 84 | قيود PurchaseReturn: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalPurchaseReturn GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 85 | قيود Payments: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalPayments GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 86 | قيود Expense: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalExpense GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 87 | قيود CashVoucher: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalCashVoucher GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 88 | قيود Stock: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalStock GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 89 | قيود Opening: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalOpening GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 90 | قيد الفاتورة الآجلة: 6 أسطر (نقدي، عميل، مبيعات، ضريبة، تكلفة، مخزون) | `SELECT COUNT(*) FROM qryJournalSale WHERE SourceID = {ref:INV2}` | 6 |
| 91 | قيد الفاتورة الآجلة: المتبقي على العميل 360 في ذمم العملاء | `SELECT Debit FROM qryJournalSale WHERE SourceID = {ref:INV2} AND AccountCode = 1300` | 360 |
| 92 | قيد الفاتورة الآجلة: التكلفة = 2×60 + 10×10 | `SELECT Debit FROM qryJournalSale WHERE SourceID = {ref:INV2} AND AccountCode = 5100` | 220 |
| 93 | ذمم العملاء من القيود = أرصدة العملاء (164) | `SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 1300` | 164 |
| 94 | ذمم الموردين من القيود = رصيد المورد (2855 دائن) | `SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 2100` | -2855 |
| 95 | المخزون من القيود = قيمة المخزون بالتكلفة (7040) | `SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalStock) AS x WHERE AccountCode = 1400` | 7040 |
| 96 | صندوق الكاشير من القيود = رصيده (770) | `SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalExpense UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalCashVoucher UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalStock UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 110000 + {ref:BOXC}` | 770 |
| 97 | الخزينة من القيود = رصيدها (7700) | `SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalExpense UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalCashVoucher UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalStock UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 110000 + {ref:BOXM}` | 7700 |
| 98 | ضريبة المخرجات من القيود = 15 + 150 + 60 − 6 | `SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn) AS x WHERE AccountCode = 2200` | -219 |
| 99 | مصروف سند النقدية لا يُقيَّد مرتين | `SELECT COUNT(*) FROM qryJournalExpense WHERE SourceID = {ref:EXPV}` | 0 |
| 100 | القيد اليدوي: 3 أسطر متوازنة (3000) | `SELECT COUNT(*) FROM qryJournalManual WHERE SourceID = {ref:MJ1} AND SourceType = 'MANUAL'` | 3 |
| 101 | قيود Manual: كل قيد متوازن | `SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalManual GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x` | 0 |
| 102 | سند صرف المصروف يُقيَّد على حساب نوع المصروف | `SELECT Debit FROM qryJournalCashVoucher WHERE SourceID = {ref:V1} AND AccountCode = 530009` | 50 |
| 103 | فحص السلامة: لا توجد مشكلات | `SELECT COUNT(*) FROM IntegrityCheckQuery` | 0 |

## qryLocAccounts1

Accounts للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.AccountCode, t.AccountName AS LocArabicName, t.AccountNameEn, t.AccountType, t.ParentCode, t.IsActive, t.IsPosting, t.IsSystem, t.AccountLevel, t.TreeKey, t.Level1Code, t.Level2Code, t.Level3Code, t.Level4Code, t.Level5Code FROM Accounts AS t
```

## qryLocAccounts

Accounts بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT AccountCode, Nz(AccountNameEn, LocArabicName) AS AccountName, AccountNameEn, AccountType, ParentCode, IsActive, IsPosting, IsSystem, AccountLevel, TreeKey, Level1Code, Level2Code, Level3Code, Level4Code, Level5Code FROM qryLocAccounts1
```

## qryLocPaymentMethods1

PaymentMethods للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.PaymentMethodID, t.MethodName AS LocArabicName, t.MethodNameEn, t.ZatcaCode, t.SortOrder, t.IsActive FROM PaymentMethods AS t
```

## qryLocPaymentMethods

PaymentMethods بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT PaymentMethodID, Nz(MethodNameEn, LocArabicName) AS MethodName, MethodNameEn, ZatcaCode, SortOrder, IsActive FROM qryLocPaymentMethods1
```

## qryLocJournalSourceTypes1

JournalSourceTypes للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.SourceType, t.TypeName AS LocArabicName, t.TypeNameEn, t.SortOrder FROM JournalSourceTypes AS t
```

## qryLocJournalSourceTypes

JournalSourceTypes بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT SourceType, Nz(TypeNameEn, LocArabicName) AS TypeName, TypeNameEn, SortOrder FROM qryLocJournalSourceTypes1
```

## qryLocTransactionTypes1

TransactionTypes للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.TransactionTypeID, t.TypeCode, t.TypeName AS LocArabicName, t.TypeNameEn, t.Direction, t.IsManual FROM TransactionTypes AS t
```

## qryLocTransactionTypes

TransactionTypes بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT TransactionTypeID, TypeCode, Nz(TypeNameEn, LocArabicName) AS TypeName, TypeNameEn, Direction, IsManual FROM qryLocTransactionTypes1
```

## qryLocRoles1

Roles للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.RoleID, t.RoleCode, t.RoleName AS LocArabicName, t.RoleNameEn, t.Description FROM Roles AS t
```

## qryLocRoles

Roles بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT RoleID, RoleCode, Nz(RoleNameEn, LocArabicName) AS RoleName, RoleNameEn, Description FROM qryLocRoles1
```

## qryLocPermissions1

Permissions للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.PermissionKey, t.PermissionName AS LocArabicName, t.PermissionNameEn, t.ModuleName AS LocArabicModuleName, t.ModuleNameEn, t.SortOrder FROM Permissions AS t
```

## qryLocPermissions

Permissions بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT PermissionKey, Nz(PermissionNameEn, LocArabicName) AS PermissionName, PermissionNameEn, Nz(ModuleNameEn, LocArabicModuleName) AS ModuleName, ModuleNameEn, SortOrder FROM qryLocPermissions1
```

## qryLocScreens1

Screens للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.ScreenName, t.ScreenTitle AS LocArabicName, t.ScreenTitleEn, t.ModuleName AS LocArabicModuleName, t.ModuleNameEn, t.SortOrder, t.PermissionKey, t.HasAdd, t.HasEdit, t.HasDelete FROM Screens AS t
```

## qryLocScreens

Screens بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT ScreenName, Nz(ScreenTitleEn, LocArabicName) AS ScreenTitle, ScreenTitleEn, Nz(ModuleNameEn, LocArabicModuleName) AS ModuleName, ModuleNameEn, SortOrder, PermissionKey, HasAdd, HasEdit, HasDelete FROM qryLocScreens1
```

## qryLocCategories1

Categories للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.CategoryID, t.CategoryName AS LocArabicName, t.CategoryNameEn, t.Description, t.IsActive, t.ImagePath, t.TileColor, t.SortOrder, t.IsAddOn FROM Categories AS t
```

## qryLocCategories

Categories بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT CategoryID, Nz(CategoryNameEn, LocArabicName) AS CategoryName, CategoryNameEn, Description, IsActive, ImagePath, TileColor, SortOrder, IsAddOn FROM qryLocCategories1
```

## qryLocUnits1

Units للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.UnitID, t.UnitName AS LocArabicName, t.UnitNameEn, t.ZatcaUnitCode, t.IsActive FROM Units AS t
```

## qryLocUnits

Units بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT UnitID, Nz(UnitNameEn, LocArabicName) AS UnitName, UnitNameEn, ZatcaUnitCode, IsActive FROM qryLocUnits1
```

## qryLocExpenseTypes1

ExpenseTypes للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.ExpenseTypeID, t.ExpenseTypeName AS LocArabicName, t.ExpenseTypeNameEn, t.IsActive FROM ExpenseTypes AS t
```

## qryLocExpenseTypes

ExpenseTypes بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT ExpenseTypeID, Nz(ExpenseTypeNameEn, LocArabicName) AS ExpenseTypeName, ExpenseTypeNameEn, IsActive FROM qryLocExpenseTypes1
```

## qryLocCustomers1

Customers للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.CustomerID, t.CustomerName AS LocArabicName, t.CustomerNameEn, t.Mobile, t.Phone, t.Email, t.VATNumber, t.CRNumber, t.BuildingNo, t.StreetName, t.District, t.City, t.PostalCode, t.Address, t.OpeningBalance, t.CurrentBalance, t.AllowCredit, t.CreditLimit, t.PaymentTermsDays, t.IsSystem, t.IsActive, t.Notes, t.CreatedAt, t.SalesRepID FROM Customers AS t
```

## qryLocCustomers

Customers بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT CustomerID, Nz(CustomerNameEn, LocArabicName) AS CustomerName, CustomerNameEn, Mobile, Phone, Email, VATNumber, CRNumber, BuildingNo, StreetName, District, City, PostalCode, Address, OpeningBalance, CurrentBalance, AllowCredit, CreditLimit, PaymentTermsDays, IsSystem, IsActive, Notes, CreatedAt, SalesRepID FROM qryLocCustomers1
```

## qryLocSuppliers1

Suppliers للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.SupplierID, t.SupplierName AS LocArabicName, t.SupplierNameEn, t.ContactPerson, t.Mobile, t.Phone, t.Email, t.VATNumber, t.CRNumber, t.Address, t.City, t.OpeningBalance, t.CurrentBalance, t.PaymentTermsDays, t.IsActive, t.Notes, t.CreatedAt, t.CurrencyCode FROM Suppliers AS t
```

## qryLocSuppliers

Suppliers بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT SupplierID, Nz(SupplierNameEn, LocArabicName) AS SupplierName, SupplierNameEn, ContactPerson, Mobile, Phone, Email, VATNumber, CRNumber, Address, City, OpeningBalance, CurrentBalance, PaymentTermsDays, IsActive, Notes, CreatedAt, CurrencyCode FROM qryLocSuppliers1
```

## qryLocCashBoxes1

CashBoxes للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.CashBoxID, t.BoxName AS LocArabicName, t.BoxNameEn, t.BoxType, t.OpeningBalance, t.OpeningDate, t.IsActive, t.Notes, t.CreatedAt FROM CashBoxes AS t
```

## qryLocCashBoxes

CashBoxes بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT CashBoxID, Nz(BoxNameEn, LocArabicName) AS BoxName, BoxNameEn, BoxType, OpeningBalance, OpeningDate, IsActive, Notes, CreatedAt FROM qryLocCashBoxes1
```

## qryLocBanks1

Banks للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.BankID, t.BankName AS LocArabicName, t.BankNameEn, t.AccountNo, t.IBAN, t.OpeningBalance, t.OpeningDate, t.IsActive, t.Notes, t.CreatedAt FROM Banks AS t
```

## qryLocBanks

Banks بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT BankID, Nz(BankNameEn, LocArabicName) AS BankName, BankNameEn, AccountNo, IBAN, OpeningBalance, OpeningDate, IsActive, Notes, CreatedAt FROM qryLocBanks1
```

## qryLocCostCenters1

CostCenters للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.CostCenterID, t.CenterCode, t.CenterName AS LocArabicName, t.CenterNameEn, t.IsDefault, t.IsActive, t.Notes, t.CreatedAt FROM CostCenters AS t
```

## qryLocCostCenters

CostCenters بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT CostCenterID, CenterCode, Nz(CenterNameEn, LocArabicName) AS CenterName, CenterNameEn, IsDefault, IsActive, Notes, CreatedAt FROM qryLocCostCenters1
```

## qryLocSalesReps1

SalesReps للواجهة الإنجليزية (خطوة 1)

```sql
SELECT t.SalesRepID, t.RepCode, t.RepName AS LocArabicName, t.RepNameEn, t.Mobile, t.EmployeeID, t.Region, t.CostCenterID, t.CommissionRate, t.CommissionBase, t.IsActive, t.Notes, t.CreatedAt FROM SalesReps AS t
```

## qryLocSalesReps

SalesReps بالأسماء الإنجليزية (الواجهة الإنجليزية)

```sql
SELECT SalesRepID, RepCode, Nz(RepNameEn, LocArabicName) AS RepName, RepNameEn, Mobile, EmployeeID, Region, CostCenterID, CommissionRate, CommissionBase, IsActive, Notes, CreatedAt FROM qryLocSalesReps1
```

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
FROM (qrySalesDocuments AS s INNER JOIN [@Customers] AS c ON s.CustomerID = c.CustomerID)
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
     INNER JOIN [@Categories] AS c ON p.CategoryID = c.CategoryID
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
FROM (Products AS p INNER JOIN [@Categories] AS c ON p.CategoryID = c.CategoryID)
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
FROM qryPurchaseDocuments AS d INNER JOIN [@Suppliers] AS s ON d.SupplierID = s.SupplierID
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
FROM ((Products AS p INNER JOIN [@Categories] AS c ON p.CategoryID = c.CategoryID)
      INNER JOIN [@Units] AS u ON p.UnitID = u.UnitID)
     LEFT JOIN qryProductLedger AS l ON p.ProductID = l.ProductID
ORDER BY p.ProductName
```

## LowStockQuery

المنتجات منخفضة المخزون: CurrentQuantity <= MinimumQuantity

```sql
SELECT p.ProductID, p.ProductCode, p.Barcode, p.ProductName, c.CategoryName,
       p.CurrentQuantity, p.MinimumQuantity, p.MinimumQuantity - p.CurrentQuantity AS ShortageQty,
       s.SupplierName, s.Mobile AS SupplierMobile
FROM (Products AS p INNER JOIN [@Categories] AS c ON p.CategoryID = c.CategoryID)
     LEFT JOIN [@Suppliers] AS s ON p.SupplierID = s.SupplierID
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
FROM InventoryTransactions AS t INNER JOIN [@TransactionTypes] AS tt
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
FROM (Products AS p INNER JOIN [@Categories] AS c ON p.CategoryID = c.CategoryID)
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
FROM Products AS p INNER JOIN [@Categories] AS c ON p.CategoryID = c.CategoryID
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
     LEFT JOIN [@Categories] AS g ON c.CategoryID = g.CategoryID
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
FROM [@Customers] AS c
WHERE c.OpeningBalance <> 0
UNION ALL
SELECT q.CustomerID, q.IssueDate, 'CHEQUE', 'شيك وارد', q.ChequeID, q.ChequeNo, CCur(0), q.Amount
FROM Cheques AS q
WHERE q.Direction = 'IN'
UNION ALL
SELECT q.CustomerID, q.StatusDate, 'CHEQUE_BOUNCE', 'شيك مرتد', q.ChequeID, q.ChequeNo, q.Amount, CCur(0)
FROM Cheques AS q
WHERE q.Direction = 'IN' AND q.Status = 'BOUNCED'
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
FROM [@Customers] AS c LEFT JOIN qryCustomerLedgerTotals AS l ON c.CustomerID = l.CustomerID
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
FROM [@Suppliers] AS s
WHERE s.OpeningBalance <> 0
UNION ALL
SELECT q.SupplierID, q.IssueDate, 'CHEQUE', 'شيك صادر', q.ChequeID, q.ChequeNo, q.Amount, CCur(0)
FROM Cheques AS q
WHERE q.Direction = 'OUT'
UNION ALL
SELECT q.SupplierID, q.StatusDate, 'CHEQUE_BOUNCE', 'شيك مرتد', q.ChequeID, q.ChequeNo, CCur(0), q.Amount
FROM Cheques AS q
WHERE q.Direction = 'OUT' AND q.Status = 'BOUNCED'
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
FROM [@Suppliers] AS s LEFT JOIN qrySupplierLedgerTotals AS l ON s.SupplierID = l.SupplierID
ORDER BY s.SupplierName
```

## qrySupplierFxMoves

حركات أرصدة الموردين بعملة كل مستند (الرصيد الافتتاحي والشيكات بعملة البرنامج)

```sql
SELECT h.SupplierID, h.CurrencyCode, h.TotalAmount - h.PaidAmount AS BaseAmount,
       Round((h.TotalAmount - h.PaidAmount) / h.ExchangeRate, 2) AS FxAmount
FROM PurchaseInvoices AS h
UNION ALL
SELECT r.SupplierID, r.CurrencyCode, r.RefundedAmount - r.TotalAmount, Round((r.RefundedAmount - r.TotalAmount) / r.ExchangeRate, 2)
FROM PurchaseReturns AS r
UNION ALL
SELECT p.SupplierID, p.CurrencyCode, -p.Amount, Round(-p.Amount / p.ExchangeRate, 2)
FROM SupplierPayments AS p
UNION ALL
SELECT s.SupplierID, z.CurrencyCode, s.OpeningBalance, s.OpeningBalance
FROM Suppliers AS s, Settings AS z
WHERE z.SettingID = 1 AND s.OpeningBalance <> 0
UNION ALL
SELECT q.SupplierID, z.CurrencyCode, IIf(q.Status = 'BOUNCED', 0, -q.Amount), IIf(q.Status = 'BOUNCED', 0, -q.Amount)
FROM Cheques AS q, Settings AS z
WHERE z.SettingID = 1 AND q.Direction = 'OUT'
```

## qryLatestRateDates

تاريخ آخر سعر لكل عملة

```sql
SELECT CurrencyCode, Max(RateDate) AS LastRateDate
FROM CurrencyRates
GROUP BY CurrencyCode
```

## qryLatestRates

آخر معامل لكل عملة

```sql
SELECT r.CurrencyCode, r.RateDate, r.Rate
FROM CurrencyRates AS r INNER JOIN qryLatestRateDates AS d
     ON (r.CurrencyCode = d.CurrencyCode AND r.RateDate = d.LastRateDate)
```

## qrySupplierFxTotals

رصيد كل مورد بكل عملة

```sql
SELECT SupplierID, CurrencyCode, Sum(BaseAmount) AS BookBalance, Sum(FxAmount) AS FxBalance
FROM qrySupplierFxMoves
GROUP BY SupplierID, CurrencyCode
```

## SupplierFxBalanceQuery

أرصدة الموردين بالعملات الأجنبية: بالدفاتر، وبآخر سعر، وفرق العملة غير المحقق

```sql
SELECT s.SupplierName, t.SupplierID, t.CurrencyCode, t.FxBalance, t.BookBalance, l.Rate AS LastRate, l.RateDate AS LastRateDate,
       Round(t.FxBalance * CCur(Nz(l.Rate, 0)), 2) AS RevaluedBalance,
       Round(t.FxBalance * CCur(Nz(l.Rate, 0)), 2) - t.BookBalance AS FxDifference
FROM ((qrySupplierFxTotals AS t INNER JOIN [@Suppliers] AS s ON t.SupplierID = s.SupplierID)
      LEFT JOIN qryLatestRates AS l ON t.CurrencyCode = l.CurrencyCode)
      INNER JOIN Settings AS st ON st.SettingID = 1
WHERE t.CurrencyCode <> st.CurrencyCode AND (t.FxBalance <> 0 OR t.BookBalance <> 0)
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

## qryCustomerAllocSums

مجموع ما رُبط من كل سند بالفواتير

```sql
SELECT PaymentID, Sum(Amount) AS SumAllocated
FROM CustomerAllocations
GROUP BY PaymentID
```

## qryCustomerPaymentFree

سندات القبض: المربوط بالفواتير والباقي غير المربوط

```sql
SELECT p.PaymentID, p.CustomerID AS PartyID, p.PaymentNumber, p.PaymentDate, p.Amount,
       CCur(Nz(s.SumAllocated, 0)) AS Allocated, p.Amount - CCur(Nz(s.SumAllocated, 0)) AS Free
FROM CustomerPayments AS p LEFT JOIN qryCustomerAllocSums AS s ON p.PaymentID = s.PaymentID
```

## qryCustomerInvoiceAlloc

مجموع ما رُبط بكل فاتورة من السندات

```sql
SELECT SalesInvoiceID, Sum(Amount) AS SumAllocated
FROM CustomerAllocations
GROUP BY SalesInvoiceID
```

## qryCustomerInvoiceReturns

مرتجعات كل فاتورة المخصومة من رصيد الحساب

```sql
SELECT SalesInvoiceID, Sum(TotalAmount - RefundedAmount) AS SumReturned
FROM SalesReturns
GROUP BY SalesInvoiceID
```

## qryCustomerInvoiceFree

الفواتير الآجلة: المتبقي وما رُبط بها وما يمكن ربطه

```sql
SELECT h.SalesInvoiceID AS InvoiceID, h.CustomerID AS PartyID, h.InvoiceNumber, h.InvoiceDate, h.DueDate,
       h.RemainingAmount, CCur(Nz(a.SumAllocated, 0)) AS Allocated, CCur(Nz(r.SumReturned, 0)) AS Returned,
       h.RemainingAmount - CCur(Nz(a.SumAllocated, 0)) - CCur(Nz(r.SumReturned, 0)) AS Free
FROM (SalesInvoices AS h LEFT JOIN qryCustomerInvoiceAlloc AS a ON h.SalesInvoiceID = a.SalesInvoiceID)
     LEFT JOIN qryCustomerInvoiceReturns AS r ON h.SalesInvoiceID = r.SalesInvoiceID
WHERE h.RemainingAmount > 0
```

## qrySupplierAllocSums

مجموع ما رُبط من كل سند بالفواتير

```sql
SELECT PaymentID, Sum(Amount) AS SumAllocated
FROM SupplierAllocations
GROUP BY PaymentID
```

## qrySupplierPaymentFree

سندات الصرف: المربوط بالفواتير والباقي غير المربوط

```sql
SELECT p.PaymentID, p.SupplierID AS PartyID, p.PaymentNumber, p.PaymentDate, p.Amount,
       CCur(Nz(s.SumAllocated, 0)) AS Allocated, p.Amount - CCur(Nz(s.SumAllocated, 0)) AS Free
FROM SupplierPayments AS p LEFT JOIN qrySupplierAllocSums AS s ON p.PaymentID = s.PaymentID
```

## qrySupplierInvoiceAlloc

مجموع ما رُبط بكل فاتورة من السندات

```sql
SELECT PurchaseInvoiceID, Sum(Amount) AS SumAllocated
FROM SupplierAllocations
GROUP BY PurchaseInvoiceID
```

## qrySupplierInvoiceReturns

مرتجعات كل فاتورة المخصومة من رصيد الحساب

```sql
SELECT PurchaseInvoiceID, Sum(TotalAmount - RefundedAmount) AS SumReturned
FROM PurchaseReturns
GROUP BY PurchaseInvoiceID
```

## qrySupplierInvoiceFree

الفواتير الآجلة: المتبقي وما رُبط بها وما يمكن ربطه

```sql
SELECT h.PurchaseInvoiceID AS InvoiceID, h.SupplierID AS PartyID, h.InvoiceNumber, h.InvoiceDate, h.DueDate,
       h.RemainingAmount, CCur(Nz(a.SumAllocated, 0)) AS Allocated, CCur(Nz(r.SumReturned, 0)) AS Returned,
       h.RemainingAmount - CCur(Nz(a.SumAllocated, 0)) - CCur(Nz(r.SumReturned, 0)) AS Free
FROM (PurchaseInvoices AS h LEFT JOIN qrySupplierInvoiceAlloc AS a ON h.PurchaseInvoiceID = a.PurchaseInvoiceID)
     LEFT JOIN qrySupplierInvoiceReturns AS r ON h.PurchaseInvoiceID = r.PurchaseInvoiceID
WHERE h.RemainingAmount > 0
```

## qryAgingDebits

المستحق على كل عميل وللمورد بالمستند: الفواتير الآجلة والرصيد الافتتاحي

```sql
SELECT 'C' AS PartyKind, h.CustomerID AS PartyID, 'INVOICE' AS DocType, h.SalesInvoiceID AS DocID,
       h.InvoiceNumber AS DocNo, h.InvoiceDate AS DocDate, h.DueDate, c.PaymentTermsDays AS TermsDays,
       h.RemainingAmount AS Amount
FROM SalesInvoices AS h INNER JOIN [@Customers] AS c ON h.CustomerID = c.CustomerID
WHERE h.RemainingAmount > 0
UNION ALL
SELECT 'C', c.CustomerID, 'OPENING', c.CustomerID, 'رصيد افتتاحي', c.CreatedAt, c.CreatedAt, 0, c.OpeningBalance
FROM [@Customers] AS c
WHERE c.OpeningBalance > 0
UNION ALL
SELECT 'S', h.SupplierID, 'INVOICE', h.PurchaseInvoiceID, h.InvoiceNumber, h.InvoiceDate, h.DueDate,
       s.PaymentTermsDays, h.RemainingAmount
FROM PurchaseInvoices AS h INNER JOIN [@Suppliers] AS s ON h.SupplierID = s.SupplierID
WHERE h.RemainingAmount > 0
UNION ALL
SELECT 'S', s.SupplierID, 'OPENING', s.SupplierID, 'رصيد افتتاحي', s.CreatedAt, s.CreatedAt, 0, s.OpeningBalance
FROM [@Suppliers] AS s
WHERE s.OpeningBalance > 0
```

## qryAgingCredits

ما يسدد المستحق: المرتجعات (على فاتورتها أولًا) والسندات والرصيد الافتتاحي الدائن

```sql
SELECT 'C' AS PartyKind, r.CustomerID AS PartyID, 'RETURN' AS CreditType, r.SalesReturnID AS CreditID,
       r.ReturnNumber AS CreditNo, r.ReturnDate AS CreditDate, r.TotalAmount - r.RefundedAmount AS Amount,
       r.SalesInvoiceID AS TargetID
FROM SalesReturns AS r
WHERE r.TotalAmount - r.RefundedAmount > 0
UNION ALL
SELECT 'C', p.CustomerID, 'PAYMENT', p.PaymentID, p.PaymentNumber, p.PaymentDate, p.Amount, 0
FROM CustomerPayments AS p
UNION ALL
SELECT 'C', c.CustomerID, 'OPENING', c.CustomerID, 'رصيد افتتاحي', c.CreatedAt, -c.OpeningBalance, 0
FROM [@Customers] AS c
WHERE c.OpeningBalance < 0
UNION ALL
SELECT 'S', r.SupplierID, 'RETURN', r.PurchaseReturnID, r.ReturnNumber, r.ReturnDate, r.TotalAmount - r.RefundedAmount,
       r.PurchaseInvoiceID
FROM PurchaseReturns AS r
WHERE r.TotalAmount - r.RefundedAmount > 0
UNION ALL
SELECT 'S', p.SupplierID, 'PAYMENT', p.PaymentID, p.PaymentNumber, p.PaymentDate, p.Amount, 0
FROM SupplierPayments AS p
UNION ALL
SELECT 'S', s.SupplierID, 'OPENING', s.SupplierID, 'رصيد افتتاحي', s.CreatedAt, -s.OpeningBalance, 0
FROM [@Suppliers] AS s
WHERE s.OpeningBalance < 0
UNION ALL
SELECT 'C', q.CustomerID, 'CHEQUE', q.ChequeID, q.ChequeNo, q.IssueDate, q.Amount, 0
FROM Cheques AS q
WHERE q.Direction = 'IN' AND q.Status <> 'BOUNCED'
UNION ALL
SELECT 'S', q.SupplierID, 'CHEQUE', q.ChequeID, q.ChequeNo, q.IssueDate, q.Amount, 0
FROM Cheques AS q
WHERE q.Direction = 'OUT' AND q.Status <> 'BOUNCED'
```

## qryAgingAllocations

ربط السندات بالفواتير (ومنها السند المسجل عن فاتورة قبل الربط)

```sql
SELECT 'C' AS PartyKind, p.CustomerID AS PartyID, a.PaymentID, a.SalesInvoiceID AS InvoiceID, a.Amount
FROM CustomerAllocations AS a INNER JOIN CustomerPayments AS p ON a.PaymentID = p.PaymentID
UNION ALL
SELECT 'C', p.CustomerID, p.PaymentID, p.SalesInvoiceID, p.Amount
FROM CustomerPayments AS p LEFT JOIN qryCustomerAllocSums AS s ON p.PaymentID = s.PaymentID
WHERE p.SalesInvoiceID Is Not Null AND s.PaymentID Is Null
UNION ALL
SELECT 'S', p.SupplierID, a.PaymentID, a.PurchaseInvoiceID, a.Amount
FROM SupplierAllocations AS a INNER JOIN SupplierPayments AS p ON a.PaymentID = p.PaymentID
UNION ALL
SELECT 'S', p.SupplierID, p.PaymentID, p.PurchaseInvoiceID, p.Amount
FROM SupplierPayments AS p LEFT JOIN qrySupplierAllocSums AS s ON p.PaymentID = s.PaymentID
WHERE p.PurchaseInvoiceID Is Not Null AND s.PaymentID Is Null
```

## ExpensesQuery

المصروفات خلال فترة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT e.ExpenseID, e.ExpenseNumber, e.ExpenseDate, t.ExpenseTypeName, e.Amount, e.Tax,
       e.TotalAmount, pm.MethodName, e.Description, em.EmployeeName, e.ExpenseTypeID
FROM ((Expenses AS e INNER JOIN [@ExpenseTypes] AS t ON e.ExpenseTypeID = t.ExpenseTypeID)
      INNER JOIN Employees AS em ON e.EmployeeID = em.EmployeeID)
     LEFT JOIN [@PaymentMethods] AS pm ON e.PaymentMethodID = pm.PaymentMethodID
WHERE e.ExpenseDate >= QDate('PeriodStart') AND e.ExpenseDate < QDate('PeriodEnd')
ORDER BY e.ExpenseDate
```

## ExpensesByTypeQuery

المصروفات مجمّعة حسب النوع خلال فترة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT t.ExpenseTypeName, Count(*) AS ExpenseCount, Sum(e.Amount) AS AmountExVAT,
       Sum(e.Tax) AS InputVAT, Sum(e.TotalAmount) AS AmountTotal
FROM Expenses AS e INNER JOIN [@ExpenseTypes] AS t ON e.ExpenseTypeID = t.ExpenseTypeID
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

## qryVatReturnLines

أسطر الإقرار الضريبي: المبيعات والمشتريات والمصروفات بفئتها الضريبية

```sql
SELECT 'S' AS Side, d.VATCategory AS Category, h.InvoiceDate AS DocDate, d.NetAmount AS Amount,
       CCur(0) AS Adjust, d.Tax AS VAT
FROM SalesInvoices AS h INNER JOIN SalesInvoiceDetails AS d ON h.SalesInvoiceID = d.SalesInvoiceID
UNION ALL
SELECT 'S', d.VATCategory, r.ReturnDate, CCur(0), -d.NetAmount, -d.Tax
FROM SalesReturns AS r INNER JOIN SalesReturnDetails AS d ON r.SalesReturnID = d.SalesReturnID
UNION ALL
SELECT 'P', IIf(d.VATRate > 0, 'S', 'Z'), h.InvoiceDate, d.NetAmount, CCur(0), d.Tax
FROM PurchaseInvoices AS h INNER JOIN PurchaseInvoiceDetails AS d ON h.PurchaseInvoiceID = d.PurchaseInvoiceID
UNION ALL
SELECT 'P', IIf(d.VATRate > 0, 'S', 'Z'), r.ReturnDate, CCur(0), -d.NetAmount, -d.Tax
FROM PurchaseReturns AS r INNER JOIN PurchaseReturnDetails AS d ON r.PurchaseReturnID = d.PurchaseReturnID
UNION ALL
SELECT 'P', 'S', e.ExpenseDate, e.Amount, CCur(0), e.Tax
FROM Expenses AS e
WHERE e.Tax <> 0
```

## qryVatReturnTotals

خانات الإقرار الضريبي للفترة محسوبة من المستندات (صف واحد)

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'S', Amount, 0)), 0)) AS SalesStdAmount,
       CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'S', Adjust, 0)), 0)) AS SalesStdAdjust,
       CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'S', VAT, 0)), 0)) AS SalesStdVAT,
       CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'Z', Amount, 0)), 0)) AS SalesZeroAmount,
       CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'Z', Adjust, 0)), 0)) AS SalesZeroAdjust,
       CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'E', Amount, 0)), 0)) AS SalesExemptAmount,
       CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'E', Adjust, 0)), 0)) AS SalesExemptAdjust,
       CCur(Nz(Sum(IIf(Side = 'P' AND Category = 'S', Amount, 0)), 0)) AS PurchStdAmount,
       CCur(Nz(Sum(IIf(Side = 'P' AND Category = 'S', Adjust, 0)), 0)) AS PurchStdAdjust,
       CCur(Nz(Sum(IIf(Side = 'P' AND Category = 'S', VAT, 0)), 0)) AS PurchStdVAT,
       CCur(Nz(Sum(IIf(Side = 'P' AND Category = 'Z', Amount, 0)), 0)) AS PurchZeroAmount,
       CCur(Nz(Sum(IIf(Side = 'P' AND Category = 'Z', Adjust, 0)), 0)) AS PurchZeroAdjust
FROM qryVatReturnLines
WHERE DocDate >= QDate('PeriodStart') AND DocDate < QDate('PeriodEnd')
```

## qryVatReturnHead

الإقرار الضريبي المختار

المعاملات: `VatReturnID`

```sql
SELECT * FROM VatReturns
WHERE VatReturnID = QLong('VatReturnID')
```

## VatReturnQuery

إقرار ضريبة القيمة المضافة بخانات نموذج الهيئة (1 إلى 16)

المعاملات: `VatReturnID`

```sql
SELECT 1 AS BoxNo, 'المبيعات الخاضعة للنسبة الأساسية (15%)' AS BoxText, v.SalesStdAmount AS Amount, v.SalesStdAdjust AS Adjust, v.SalesStdVAT AS VAT, 'L' AS RowKind, v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 2, 'المبيعات للمواطنين (الخدمات الصحية الخاصة والتعليم الأهلي والمسكن الأول)', CCur(0), CCur(0), CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 3, 'المبيعات المحلية الخاضعة للنسبة الصفرية', v.SalesZeroAmount, v.SalesZeroAdjust, CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 4, 'الصادرات', CCur(0), CCur(0), CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 5, 'المبيعات المعفاة', v.SalesExemptAmount, v.SalesExemptAdjust, CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 6, 'إجمالي المبيعات', v.SalesStdAmount + v.SalesZeroAmount + v.SalesExemptAmount, v.SalesStdAdjust + v.SalesZeroAdjust + v.SalesExemptAdjust, v.SalesStdVAT, 'T', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 7, 'المشتريات الخاضعة للنسبة الأساسية (مع المصروفات بفاتورة ضريبية)', v.PurchStdAmount, v.PurchStdAdjust, v.PurchStdVAT, 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 8, 'الاستيرادات الخاضعة للنسبة الأساسية والمدفوعة ضريبتها في الجمارك', CCur(0), CCur(0), CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 9, 'الاستيرادات الخاضعة للضريبة بآلية الاحتساب العكسي', CCur(0), CCur(0), CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 10, 'المشتريات الخاضعة للنسبة الصفرية', v.PurchZeroAmount, v.PurchZeroAdjust, CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 11, 'المشتريات المعفاة', CCur(0), CCur(0), CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 12, 'إجمالي المشتريات', v.PurchStdAmount + v.PurchZeroAmount, v.PurchStdAdjust + v.PurchZeroAdjust, v.PurchStdVAT, 'T', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 13, 'إجمالي ضريبة القيمة المضافة المستحقة عن الفترة الحالية', Null, Null, v.SalesStdVAT - v.PurchStdVAT, 'N', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 14, 'تصحيحات من الفترات السابقة', Null, Null, v.Corrections, 'N', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 15, 'ضريبة القيمة المضافة المرحَّلة من الفترات السابقة (رصيد دائن)', Null, Null, v.CarriedCredit, 'N', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
UNION ALL
SELECT 16, 'صافي الضريبة المستحقة (سالب = مستردة)', Null, Null, v.NetDue, 'N', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount
FROM qryVatReturnHead AS v
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

## qryRepDocs

عمليات المندوبين: صافي المبيعات بدون الضريبة والتحصيل

```sql
SELECT h.SalesRepID, h.InvoiceDate AS DocDate, h.TaxableAmount AS NetSales, h.PaidAmount AS Collected
FROM SalesInvoices AS h
WHERE h.SalesRepID Is Not Null
UNION ALL
SELECT r.SalesRepID, r.ReturnDate, -r.TaxableAmount, -r.RefundedAmount
FROM SalesReturns AS r
WHERE r.SalesRepID Is Not Null
UNION ALL
SELECT p.SalesRepID, p.PaymentDate, CCur(0), p.Amount
FROM CustomerPayments AS p
WHERE p.SalesRepID Is Not Null
```

## qryRepPeriodTotals

مبيعات وتحصيل كل مندوب في الفترة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT SalesRepID, Sum(NetSales) AS SumSales, Sum(Collected) AS SumCollected
FROM qryRepDocs
WHERE DocDate >= QDate('PeriodStart') AND DocDate < QDate('PeriodEnd')
GROUP BY SalesRepID
```

## qryRepTargetTotals

أهداف كل مندوب في أشهر الفترة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT SalesRepID, Sum(TargetAmount) AS SumTarget
FROM SalesRepTargets
WHERE TargetYear * 100 + TargetMonth >= Year(QDate('PeriodStart')) * 100 + Month(QDate('PeriodStart'))
  AND TargetYear * 100 + TargetMonth <= Year(DateAdd('d', -1, QDate('PeriodEnd'))) * 100
                                        + Month(DateAdd('d', -1, QDate('PeriodEnd')))
GROUP BY SalesRepID
```

## qryRepCommissionPaid

ما صُرف لكل مندوب من عمولاته (سندات صرف النقدية)

```sql
SELECT SalesRepID, Sum(Amount) AS SumPaid
FROM CashVouchers
WHERE Category = 'COMMISSION' AND SalesRepID Is Not Null
GROUP BY SalesRepID
```

## qryRepCommissionPosted

العمولات المرحَّلة لكل مندوب

```sql
SELECT l.SalesRepID, Sum(l.Commission) AS SumPosted
FROM CommissionLines AS l INNER JOIN CommissionRuns AS r ON l.CommissionRunID = r.CommissionRunID
WHERE r.Status = 'POSTED'
GROUP BY l.SalesRepID
```

## RepPerformanceQuery

أداء المندوبين في الفترة: المبيعات والتحصيل والهدف والإنجاز والعمولة المتوقعة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT s.SalesRepID, s.RepCode, s.RepName, s.Region, s.IsActive, s.CommissionRate, s.CommissionBase,
       CCur(Nz(t.SumSales, 0)) AS NetSales, CCur(Nz(t.SumCollected, 0)) AS Collections, CCur(Nz(g.SumTarget, 0)) AS Target,
       IIf(CCur(Nz(g.SumTarget, 0)) = 0, Null, CCur(Nz(t.SumSales, 0)) / CCur(Nz(g.SumTarget, 0))) AS Achievement,
       IIf(s.CommissionBase = 'COLLECTION', CCur(Nz(t.SumCollected, 0)), CCur(Nz(t.SumSales, 0))) AS BaseAmount,
       IIf(IIf(s.CommissionBase = 'COLLECTION', CCur(Nz(t.SumCollected, 0)), CCur(Nz(t.SumSales, 0))) > 0,
           Round(IIf(s.CommissionBase = 'COLLECTION', CCur(Nz(t.SumCollected, 0)), CCur(Nz(t.SumSales, 0))) * s.CommissionRate, 2),
           0) AS Commission
FROM ([@SalesReps] AS s LEFT JOIN qryRepPeriodTotals AS t ON s.SalesRepID = t.SalesRepID)
     LEFT JOIN qryRepTargetTotals AS g ON s.SalesRepID = g.SalesRepID
WHERE s.IsActive = True OR CCur(Nz(t.SumSales, 0)) <> 0 OR CCur(Nz(t.SumCollected, 0)) <> 0
```

## RepCustomersQuery

عملاء كل مندوب وأرصدتهم

```sql
SELECT s.SalesRepID, s.RepName, c.CustomerID, c.CustomerName, c.Mobile, b.Balance
FROM ([@Customers] AS c INNER JOIN [@SalesReps] AS s ON c.SalesRepID = s.SalesRepID)
     INNER JOIN CustomerBalanceQuery AS b ON c.CustomerID = b.CustomerID
```

## RepCommissionBalanceQuery

العمولات المستحقة لكل مندوب: المرحَّل والمصروف والباقي

```sql
SELECT s.SalesRepID, s.RepCode, s.RepName, CCur(Nz(p.SumPosted, 0)) AS Posted, CCur(Nz(d.SumPaid, 0)) AS Paid,
       CCur(Nz(p.SumPosted, 0)) - CCur(Nz(d.SumPaid, 0)) AS Payable
FROM ([@SalesReps] AS s LEFT JOIN qryRepCommissionPosted AS p ON s.SalesRepID = p.SalesRepID)
     LEFT JOIN qryRepCommissionPaid AS d ON s.SalesRepID = d.SalesRepID
```

## CommissionSheetQuery

مسير العمولات المختار بأسطر المندوبين

المعاملات: `CommissionRunID`

```sql
SELECT r.CommissionRunID, r.RunNumber, r.RunMonth, r.Status, l.SalesRepID, l.RepName, l.NetSales, l.Collections,
       IIf(l.CommissionBase = 'COLLECTION', 'التحصيل', 'المبيعات') AS BaseName, l.BaseAmount, l.CommissionRate,
       l.Adjustment, l.Commission, l.Notes
FROM CommissionRuns AS r INNER JOIN CommissionLines AS l ON r.CommissionRunID = l.CommissionRunID
WHERE r.CommissionRunID = QLong('CommissionRunID')
```

## qryIndicatorLines

أسطر القيود مع مجموعة الحساب (المستوى 2) لحساب المؤشرات المالية

```sql
SELECT l.AccountCode, a.Level2Code, e.EntryDate, e.SourceType, l.Debit, l.Credit
FROM (JournalLines AS l INNER JOIN JournalEntries AS e ON l.EntryID = e.EntryID)
     INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode
```

## FinancialIndicatorsQuery

المؤشرات المالية (صف واحد): هامش الربح، دوران المخزون، فترة التحصيل، السيولة

المعاملات: `IndEnd`, `IndMonth`, `IndPrevMonth`, `IndYear`, `Ind90`

```sql
SELECT CCur(Nz(Sum(IIf(i.Level2Code = 41 AND i.EntryDate >= QDate('IndMonth') AND i.EntryDate < QDate('IndEnd') AND i.SourceType <> 'YEAR_CLOSE', i.Credit - i.Debit, 0)), 0)) AS MonthSales,
       CCur(Nz(Sum(IIf(i.Level2Code = 51 AND i.EntryDate >= QDate('IndMonth') AND i.EntryDate < QDate('IndEnd') AND i.SourceType <> 'YEAR_CLOSE', i.Debit - i.Credit, 0)), 0)) AS MonthCost,
       CCur(Nz(Sum(IIf(i.Level2Code = 41 AND i.EntryDate >= QDate('IndPrevMonth') AND i.EntryDate < QDate('IndMonth') AND i.SourceType <> 'YEAR_CLOSE', i.Credit - i.Debit, 0)), 0)) AS PrevSales,
       CCur(Nz(Sum(IIf(i.Level2Code = 51 AND i.EntryDate >= QDate('IndPrevMonth') AND i.EntryDate < QDate('IndMonth') AND i.SourceType <> 'YEAR_CLOSE', i.Debit - i.Credit, 0)), 0)) AS PrevCost,
       CCur(Nz(Sum(IIf(i.Level2Code = 51 AND i.EntryDate >= QDate('IndYear') AND i.EntryDate < QDate('IndEnd') AND i.SourceType <> 'YEAR_CLOSE', i.Debit - i.Credit, 0)), 0)) AS YearCost,
       CCur(Nz(Sum(IIf(i.AccountCode = 1400 AND i.EntryDate < QDate('IndYear'), i.Debit - i.Credit, 0)), 0)) AS StockStart,
       CCur(Nz(Sum(IIf(i.AccountCode = 1400 AND i.EntryDate < QDate('IndEnd'), i.Debit - i.Credit, 0)), 0)) AS StockEnd,
       CCur(Nz(Sum(IIf(i.AccountCode = 1300 AND i.EntryDate < QDate('IndEnd'), i.Debit - i.Credit, 0)), 0)) AS Receivables,
       CCur(Nz(Sum(IIf(i.AccountCode = 1300 AND i.SourceType = 'SALE' AND i.EntryDate >= QDate('Ind90') AND i.EntryDate < QDate('IndEnd'), i.Debit, 0)), 0)) AS CreditSales,
       CCur(Nz(Sum(IIf(i.Level2Code = 11 AND i.EntryDate < QDate('IndEnd'), i.Debit - i.Credit, 0)), 0)) AS CurrentAssets,
       CCur(Nz(Sum(IIf(i.Level2Code = 21 AND i.EntryDate < QDate('IndEnd'), i.Credit - i.Debit, 0)), 0)) AS CurrentLiabilities
FROM qryIndicatorLines AS i
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

## qryEInvoiceDocs

مستندات البيع وحالة الفاتورة الإلكترونية (الفواتير والمرتجعات)

```sql
SELECT 'SALE' AS DocKind, h.SalesInvoiceID AS DocID, h.InvoiceNumber AS DocNumber, h.InvoiceDate AS DocDate,
       'فاتورة بيع' AS DocKindName, c.CustomerName, h.TotalAmount, h.ZatcaStatus AS EStatus,
       IIf(h.ZatcaStatus = 'PENDING', 'بانتظار الإرسال', IIf(h.ZatcaStatus = 'REPORTED', 'مُبلَّغ', IIf(h.ZatcaStatus = 'CLEARED', 'معتمد', IIf(h.ZatcaStatus = 'WARNING', 'مقبول مع تحذير', IIf(h.ZatcaStatus = 'REJECTED', 'مرفوض', IIf(h.ZatcaStatus = 'SUBMITTED', 'مُرسل', IIf(h.ZatcaStatus = 'VALID', 'صالح', IIf(h.ZatcaStatus = 'INVALID', 'غير صالح', IIf(h.ZatcaStatus = 'CANCELLED', 'ملغى', 'لا يُرسل'))))))))) AS StatusName,
       Nz(h.EInvoiceAttempts, 0) AS Attempts, h.EInvoiceError AS LastError, h.ICV, h.InvoiceSubType
FROM SalesInvoices AS h INNER JOIN [@Customers] AS c ON h.CustomerID = c.CustomerID
UNION ALL
SELECT 'RETURN', r.SalesReturnID, r.ReturnNumber, r.ReturnDate, 'مرتجع بيع', c.CustomerName, r.TotalAmount,
       r.ZatcaStatus, IIf(r.ZatcaStatus = 'PENDING', 'بانتظار الإرسال', IIf(r.ZatcaStatus = 'REPORTED', 'مُبلَّغ', IIf(r.ZatcaStatus = 'CLEARED', 'معتمد', IIf(r.ZatcaStatus = 'WARNING', 'مقبول مع تحذير', IIf(r.ZatcaStatus = 'REJECTED', 'مرفوض', IIf(r.ZatcaStatus = 'SUBMITTED', 'مُرسل', IIf(r.ZatcaStatus = 'VALID', 'صالح', IIf(r.ZatcaStatus = 'INVALID', 'غير صالح', IIf(r.ZatcaStatus = 'CANCELLED', 'ملغى', 'لا يُرسل'))))))))),
       Nz(r.EInvoiceAttempts, 0), r.EInvoiceError, r.ICV, r.InvoiceSubType
FROM SalesReturns AS r INNER JOIN [@Customers] AS c ON r.CustomerID = c.CustomerID
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
      INNER JOIN [@Units] AS u ON p.UnitID = u.UnitID)
     INNER JOIN [@Suppliers] AS s ON h.SupplierID = s.SupplierID)
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
      INNER JOIN [@Units] AS u ON p.UnitID = u.UnitID)
     INNER JOIN [@Suppliers] AS s ON r.SupplierID = s.SupplierID)
    INNER JOIN Employees AS e ON r.EmployeeID = e.EmployeeID
```

## qryVoucherPrint

بيانات طباعة سندات القبض (من العملاء) وسندات الصرف (للموردين)

```sql
SELECT 'RECEIPT' AS DocKind, p.PaymentID AS DocID, p.PaymentNumber AS DocNumber,
       p.PaymentDate AS DocDate, 1 AS LineNumber, c.CustomerName AS PartyName,
       c.Mobile AS PartyMobile, p.Amount, m.MethodName, p.Notes, e.EmployeeName,
       c.CurrentBalance AS PartyBalance
FROM ((CustomerPayments AS p INNER JOIN [@Customers] AS c ON p.CustomerID = c.CustomerID)
      INNER JOIN [@PaymentMethods] AS m ON p.PaymentMethodID = m.PaymentMethodID)
     INNER JOIN Employees AS e ON p.EmployeeID = e.EmployeeID
UNION ALL
SELECT 'PAYMENT', p.PaymentID, p.PaymentNumber, p.PaymentDate, 1, s.SupplierName, s.Mobile,
       p.Amount, m.MethodName, p.Notes, e.EmployeeName, s.CurrentBalance
FROM ((SupplierPayments AS p INNER JOIN [@Suppliers] AS s ON p.SupplierID = s.SupplierID)
      INNER JOIN [@PaymentMethods] AS m ON p.PaymentMethodID = m.PaymentMethodID)
     INNER JOIN Employees AS e ON p.EmployeeID = e.EmployeeID
```

## qryCashMovements

كل حركات النقدية في الخزينة والصناديق: داخل (+) وخارج (−)

```sql
SELECT h.CashBoxID, h.InvoiceDate AS MoveDate, 'SALE' AS MoveType, 'فاتورة بيع' AS MoveTypeName,
       h.InvoiceNumber AS DocNumber, c.CustomerName AS PartyName, h.Notes AS Details,
       h.PaidAmount AS AmountIn, CCur(0) AS AmountOut, h.EmployeeID
FROM SalesInvoices AS h INNER JOIN [@Customers] AS c ON h.CustomerID = c.CustomerID
WHERE h.CashBoxID Is Not Null AND h.PaidAmount <> 0
UNION ALL
SELECT r.CashBoxID, r.ReturnDate, 'SALES_RETURN', 'مرتجع بيع (رد نقدي)', r.ReturnNumber,
       c.CustomerName, r.Reason, CCur(0), r.RefundedAmount, r.EmployeeID
FROM SalesReturns AS r INNER JOIN [@Customers] AS c ON r.CustomerID = c.CustomerID
WHERE r.CashBoxID Is Not Null AND r.RefundedAmount <> 0
UNION ALL
SELECT p.CashBoxID, p.PaymentDate, 'CUSTOMER_PAYMENT', 'سند قبض من عميل', p.PaymentNumber,
       c.CustomerName, p.Notes, p.Amount, CCur(0), p.EmployeeID
FROM CustomerPayments AS p INNER JOIN [@Customers] AS c ON p.CustomerID = c.CustomerID
WHERE p.CashBoxID Is Not Null
UNION ALL
SELECT h.CashBoxID, h.InvoiceDate, 'PURCHASE', 'فاتورة شراء', h.InvoiceNumber,
       s.SupplierName, h.Notes, CCur(0), h.PaidAmount, h.EmployeeID
FROM PurchaseInvoices AS h INNER JOIN [@Suppliers] AS s ON h.SupplierID = s.SupplierID
WHERE h.CashBoxID Is Not Null AND h.PaidAmount <> 0
UNION ALL
SELECT r.CashBoxID, r.ReturnDate, 'PURCHASE_RETURN', 'مرتجع شراء (استرداد نقدي)', r.ReturnNumber,
       s.SupplierName, r.Reason, r.RefundedAmount, CCur(0), r.EmployeeID
FROM PurchaseReturns AS r INNER JOIN [@Suppliers] AS s ON r.SupplierID = s.SupplierID
WHERE r.CashBoxID Is Not Null AND r.RefundedAmount <> 0
UNION ALL
SELECT p.CashBoxID, p.PaymentDate, 'SUPPLIER_PAYMENT', 'سند صرف لمورد', p.PaymentNumber,
       s.SupplierName, p.Notes, CCur(0), p.Amount, p.EmployeeID
FROM SupplierPayments AS p INNER JOIN [@Suppliers] AS s ON p.SupplierID = s.SupplierID
WHERE p.CashBoxID Is Not Null
UNION ALL
SELECT e.CashBoxID, e.ExpenseDate, 'EXPENSE', 'مصروف', e.ExpenseNumber,
       t.ExpenseTypeName, e.Description, CCur(0), e.TotalAmount, e.EmployeeID
FROM Expenses AS e INNER JOIN [@ExpenseTypes] AS t ON e.ExpenseTypeID = t.ExpenseTypeID
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
FROM CashVouchers AS v INNER JOIN [@CashBoxes] AS b ON v.ToCashBoxID = b.CashBoxID
WHERE v.VoucherType = 'TRANSFER'
UNION ALL
SELECT v.ToCashBoxID, v.VoucherDate, 'TRANSFER_IN', 'تحويل من صندوق آخر', v.VoucherNumber,
       b.BoxName, v.Description, v.Amount, CCur(0), v.EmployeeID
FROM CashVouchers AS v INNER JOIN [@CashBoxes] AS b ON v.CashBoxID = b.CashBoxID
WHERE v.VoucherType = 'TRANSFER'
UNION ALL
SELECT t.CashBoxID, t.TxDate, 'BANK_DEPOSIT', 'إيداع في البنك', t.TxNumber, k.BankName, t.Description,
       CCur(0), t.Amount, t.EmployeeID
FROM BankTransactions AS t INNER JOIN [@Banks] AS k ON t.BankID = k.BankID
WHERE t.TxType = 'DEPOSIT'
UNION ALL
SELECT t.CashBoxID, t.TxDate, 'BANK_WITHDRAW', 'سحب من البنك', t.TxNumber, k.BankName, t.Description,
       t.Amount, CCur(0), t.EmployeeID
FROM BankTransactions AS t INNER JOIN [@Banks] AS k ON t.BankID = k.BankID
WHERE t.TxType = 'WITHDRAW'
UNION ALL
SELECT r.CashBoxID, r.PaidDate, 'PAYROLL', 'صرف الرواتب', r.RunNumber, '-', r.Notes, CCur(0), r.PaidAmount,
       r.EmployeeID
FROM PayrollRuns AS r
WHERE r.Status = 'POSTED' AND r.PaidFrom = 'CASHBOX' AND r.PaidAmount <> 0
UNION ALL
SELECT a.CashBoxID, a.PurchaseDate, 'ASSET', 'شراء أصل ثابت', a.AssetCode, a.AssetName, a.Notes,
       CCur(0), a.Cost + a.InputVAT, a.EmployeeID
FROM FixedAssets AS a
WHERE a.SourceType = 'CASHBOX'
UNION ALL
SELECT a.DisposalCashBoxID, a.DisposalDate, 'ASSET_SALE', 'بيع أصل ثابت', a.AssetCode, a.AssetName, a.Notes,
       a.DisposalProceeds, CCur(0), a.EmployeeID
FROM FixedAssets AS a
WHERE a.Status = 'DISPOSED' AND a.DisposalTo = 'CASHBOX' AND a.DisposalProceeds <> 0
UNION ALL
SELECT b.CashBoxID, b.OpeningDate, 'OPENING', 'رصيد افتتاحي', '-', b.BoxName, b.Notes,
       b.OpeningBalance, CCur(0), Null
FROM [@CashBoxes] AS b
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
FROM [@CashBoxes] AS b LEFT JOIN qryCashBoxTotals AS t ON b.CashBoxID = t.CashBoxID
ORDER BY b.BoxType DESC, b.BoxName
```

## CashStatementQuery

حركة الخزينة / الصندوق لفترة: رصيد أول المدة ثم الحركات (0 = كل الصناديق)

المعاملات: `PeriodStart`, `PeriodEnd`, `CashBoxID`

```sql
SELECT 1 AS SortKey, m.MoveDate, m.MoveType, m.MoveTypeName, m.DocNumber, m.PartyName, m.Details,
       b.BoxName, m.AmountIn, m.AmountOut, m.CashBoxID
FROM qryCashMovements AS m INNER JOIN [@CashBoxes] AS b ON m.CashBoxID = b.CashBoxID
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
FROM ((CashClosings AS c INNER JOIN [@CashBoxes] AS b ON c.CashBoxID = b.CashBoxID)
      INNER JOIN Employees AS e ON c.EmployeeID = e.EmployeeID)
     LEFT JOIN [@CashBoxes] AS t ON c.ToCashBoxID = t.CashBoxID
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
FROM ((CashClosings AS c INNER JOIN [@CashBoxes] AS b ON c.CashBoxID = b.CashBoxID)
      INNER JOIN Employees AS e ON c.EmployeeID = e.EmployeeID)
     LEFT JOIN [@CashBoxes] AS t ON c.ToCashBoxID = t.CashBoxID
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
FROM ((((CashVouchers AS v INNER JOIN [@CashBoxes] AS b ON v.CashBoxID = b.CashBoxID)
        INNER JOIN Employees AS e ON v.EmployeeID = e.EmployeeID)
       LEFT JOIN [@CashBoxes] AS t ON v.ToCashBoxID = t.CashBoxID)
      LEFT JOIN Expenses AS ex ON v.ExpenseID = ex.ExpenseID)
     LEFT JOIN [@ExpenseTypes] AS x ON ex.ExpenseTypeID = x.ExpenseTypeID
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
SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 1 AS LineOrder, IIf(h.CashBoxID Is Null, IIf(h.PaymentMethodID Is Null Or h.PaymentMethodID = 1, 1190, IIf(h.BankID Is Null, 1200, 120000 + h.BankID)), 110000 + h.CashBoxID) AS AccountCode, h.PaidAmount AS Debit, CCur(0) AS Credit, c.CustomerName AS LineText, IIf(h.CostCenterID Is Null, 0, h.CostCenterID) AS CostCenter
FROM SalesInvoices AS h INNER JOIN [@Customers] AS c ON h.CustomerID = c.CustomerID
WHERE h.PaidAmount <> 0
UNION ALL
SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 2 AS LineOrder, 1300 AS AccountCode, h.RemainingAmount AS Debit, CCur(0) AS Credit, c.CustomerName AS LineText, IIf(h.CostCenterID Is Null, 0, h.CostCenterID) AS CostCenter
FROM SalesInvoices AS h INNER JOIN [@Customers] AS c ON h.CustomerID = c.CustomerID
WHERE h.RemainingAmount <> 0
UNION ALL
SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 3 AS LineOrder, 4100 AS AccountCode, CCur(0) AS Debit, h.TaxableAmount AS Credit, 'المبيعات' AS LineText, IIf(h.CostCenterID Is Null, 0, h.CostCenterID) AS CostCenter
FROM SalesInvoices AS h INNER JOIN [@Customers] AS c ON h.CustomerID = c.CustomerID
WHERE h.TaxableAmount <> 0
UNION ALL
SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 4 AS LineOrder, 2200 AS AccountCode, CCur(0) AS Debit, h.Tax AS Credit, 'ضريبة المخرجات' AS LineText, IIf(h.CostCenterID Is Null, 0, h.CostCenterID) AS CostCenter
FROM SalesInvoices AS h INNER JOIN [@Customers] AS c ON h.CustomerID = c.CustomerID
WHERE h.Tax <> 0
UNION ALL
SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 5 AS LineOrder, 5100 AS AccountCode, k.SaleCost AS Debit, CCur(0) AS Credit, 'تكلفة البضاعة المباعة' AS LineText, IIf(h.CostCenterID Is Null, 0, h.CostCenterID) AS CostCenter
FROM (SalesInvoices AS h INNER JOIN [@Customers] AS c ON h.CustomerID = c.CustomerID) INNER JOIN qrySaleCost AS k ON h.SalesInvoiceID = k.SalesInvoiceID
WHERE k.SaleCost <> 0
UNION ALL
SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 6 AS LineOrder, 1400 AS AccountCode, CCur(0) AS Debit, k.SaleCost AS Credit, 'المخزون' AS LineText, IIf(h.CostCenterID Is Null, 0, h.CostCenterID) AS CostCenter
FROM (SalesInvoices AS h INNER JOIN [@Customers] AS c ON h.CustomerID = c.CustomerID) INNER JOIN qrySaleCost AS k ON h.SalesInvoiceID = k.SalesInvoiceID
WHERE k.SaleCost <> 0
```

## qryJournalSalesReturn

أسطر قيود مرتجعات البيع

```sql
SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 1 AS LineOrder, 4110 AS AccountCode, r.TaxableAmount AS Debit, CCur(0) AS Credit, 'مردودات المبيعات' AS LineText, IIf(r.CostCenterID Is Null, 0, r.CostCenterID) AS CostCenter
FROM SalesReturns AS r INNER JOIN [@Customers] AS c ON r.CustomerID = c.CustomerID
WHERE r.TaxableAmount <> 0
UNION ALL
SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 2 AS LineOrder, 2200 AS AccountCode, r.Tax AS Debit, CCur(0) AS Credit, 'ضريبة المخرجات' AS LineText, IIf(r.CostCenterID Is Null, 0, r.CostCenterID) AS CostCenter
FROM SalesReturns AS r INNER JOIN [@Customers] AS c ON r.CustomerID = c.CustomerID
WHERE r.Tax <> 0
UNION ALL
SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 3 AS LineOrder, IIf(r.CashBoxID Is Null, IIf(r.PaymentMethodID Is Null Or r.PaymentMethodID = 1, 1190, IIf(r.BankID Is Null, 1200, 120000 + r.BankID)), 110000 + r.CashBoxID) AS AccountCode, CCur(0) AS Debit, r.RefundedAmount AS Credit, c.CustomerName AS LineText, IIf(r.CostCenterID Is Null, 0, r.CostCenterID) AS CostCenter
FROM SalesReturns AS r INNER JOIN [@Customers] AS c ON r.CustomerID = c.CustomerID
WHERE r.RefundedAmount <> 0
UNION ALL
SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 4 AS LineOrder, 1300 AS AccountCode, CCur(0) AS Debit, r.TotalAmount - r.RefundedAmount AS Credit, c.CustomerName AS LineText, IIf(r.CostCenterID Is Null, 0, r.CostCenterID) AS CostCenter
FROM SalesReturns AS r INNER JOIN [@Customers] AS c ON r.CustomerID = c.CustomerID
WHERE r.TotalAmount - r.RefundedAmount <> 0
UNION ALL
SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 5 AS LineOrder, 1400 AS AccountCode, k.ReturnCost AS Debit, CCur(0) AS Credit, 'المخزون' AS LineText, IIf(r.CostCenterID Is Null, 0, r.CostCenterID) AS CostCenter
FROM (SalesReturns AS r INNER JOIN [@Customers] AS c ON r.CustomerID = c.CustomerID) INNER JOIN qryReturnCost AS k ON r.SalesReturnID = k.SalesReturnID
WHERE k.ReturnCost <> 0
UNION ALL
SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 6 AS LineOrder, 5100 AS AccountCode, CCur(0) AS Debit, k.ReturnCost AS Credit, 'تكلفة البضاعة المباعة' AS LineText, IIf(r.CostCenterID Is Null, 0, r.CostCenterID) AS CostCenter
FROM (SalesReturns AS r INNER JOIN [@Customers] AS c ON r.CustomerID = c.CustomerID) INNER JOIN qryReturnCost AS k ON r.SalesReturnID = k.SalesReturnID
WHERE k.ReturnCost <> 0
```

## qryJournalPurchase

أسطر قيود فواتير الشراء

```sql
SELECT 'PURCHASE' AS SourceType, h.PurchaseInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, s.SupplierName AS Party, 1 AS LineOrder, 1400 AS AccountCode, h.TaxableAmount AS Debit, CCur(0) AS Credit, 'المخزون' AS LineText, 0 AS CostCenter
FROM PurchaseInvoices AS h INNER JOIN [@Suppliers] AS s ON h.SupplierID = s.SupplierID
WHERE h.TaxableAmount <> 0
UNION ALL
SELECT 'PURCHASE' AS SourceType, h.PurchaseInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, s.SupplierName AS Party, 2 AS LineOrder, 1500 AS AccountCode, h.Tax AS Debit, CCur(0) AS Credit, 'ضريبة المدخلات' AS LineText, 0 AS CostCenter
FROM PurchaseInvoices AS h INNER JOIN [@Suppliers] AS s ON h.SupplierID = s.SupplierID
WHERE h.Tax <> 0
UNION ALL
SELECT 'PURCHASE' AS SourceType, h.PurchaseInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, s.SupplierName AS Party, 3 AS LineOrder, IIf(h.CashBoxID Is Null, IIf(h.PaymentMethodID Is Null Or h.PaymentMethodID = 1, 1190, IIf(h.BankID Is Null, 1200, 120000 + h.BankID)), 110000 + h.CashBoxID) AS AccountCode, CCur(0) AS Debit, h.PaidAmount AS Credit, s.SupplierName AS LineText, 0 AS CostCenter
FROM PurchaseInvoices AS h INNER JOIN [@Suppliers] AS s ON h.SupplierID = s.SupplierID
WHERE h.PaidAmount <> 0
UNION ALL
SELECT 'PURCHASE' AS SourceType, h.PurchaseInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, s.SupplierName AS Party, 4 AS LineOrder, 2100 AS AccountCode, CCur(0) AS Debit, h.RemainingAmount AS Credit, s.SupplierName AS LineText, 0 AS CostCenter
FROM PurchaseInvoices AS h INNER JOIN [@Suppliers] AS s ON h.SupplierID = s.SupplierID
WHERE h.RemainingAmount <> 0
```

## qryJournalPurchaseReturn

أسطر قيود مرتجعات الشراء

```sql
SELECT 'PURCHASE_RETURN' AS SourceType, r.PurchaseReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, s.SupplierName AS Party, 1 AS LineOrder, IIf(r.CashBoxID Is Null, IIf(r.PaymentMethodID Is Null Or r.PaymentMethodID = 1, 1190, IIf(r.BankID Is Null, 1200, 120000 + r.BankID)), 110000 + r.CashBoxID) AS AccountCode, r.RefundedAmount AS Debit, CCur(0) AS Credit, s.SupplierName AS LineText, 0 AS CostCenter
FROM PurchaseReturns AS r INNER JOIN [@Suppliers] AS s ON r.SupplierID = s.SupplierID
WHERE r.RefundedAmount <> 0
UNION ALL
SELECT 'PURCHASE_RETURN' AS SourceType, r.PurchaseReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, s.SupplierName AS Party, 2 AS LineOrder, 2100 AS AccountCode, r.TotalAmount - r.RefundedAmount AS Debit, CCur(0) AS Credit, s.SupplierName AS LineText, 0 AS CostCenter
FROM PurchaseReturns AS r INNER JOIN [@Suppliers] AS s ON r.SupplierID = s.SupplierID
WHERE r.TotalAmount - r.RefundedAmount <> 0
UNION ALL
SELECT 'PURCHASE_RETURN' AS SourceType, r.PurchaseReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, s.SupplierName AS Party, 3 AS LineOrder, 1400 AS AccountCode, CCur(0) AS Debit, r.TaxableAmount AS Credit, 'المخزون' AS LineText, 0 AS CostCenter
FROM PurchaseReturns AS r INNER JOIN [@Suppliers] AS s ON r.SupplierID = s.SupplierID
WHERE r.TaxableAmount <> 0
UNION ALL
SELECT 'PURCHASE_RETURN' AS SourceType, r.PurchaseReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, s.SupplierName AS Party, 4 AS LineOrder, 1500 AS AccountCode, CCur(0) AS Debit, r.Tax AS Credit, 'ضريبة المدخلات' AS LineText, 0 AS CostCenter
FROM PurchaseReturns AS r INNER JOIN [@Suppliers] AS s ON r.SupplierID = s.SupplierID
WHERE r.Tax <> 0
```

## qryJournalPayments

أسطر قيود سندات القبض من العملاء والصرف للموردين

```sql
SELECT 'CUSTOMER_PAYMENT' AS SourceType, p.PaymentID AS SourceID, p.PaymentNumber AS SourceNumber, p.PaymentDate AS SourceDate, c.CustomerName AS Party, 1 AS LineOrder, IIf(p.CashBoxID Is Null, IIf(p.PaymentMethodID Is Null Or p.PaymentMethodID = 1, 1190, IIf(p.BankID Is Null, 1200, 120000 + p.BankID)), 110000 + p.CashBoxID) AS AccountCode, p.Amount AS Debit, CCur(0) AS Credit, c.CustomerName AS LineText, 0 AS CostCenter
FROM CustomerPayments AS p INNER JOIN [@Customers] AS c ON p.CustomerID = c.CustomerID
WHERE p.Amount <> 0
UNION ALL
SELECT 'CUSTOMER_PAYMENT' AS SourceType, p.PaymentID AS SourceID, p.PaymentNumber AS SourceNumber, p.PaymentDate AS SourceDate, c.CustomerName AS Party, 2 AS LineOrder, 1300 AS AccountCode, CCur(0) AS Debit, p.Amount AS Credit, c.CustomerName AS LineText, 0 AS CostCenter
FROM CustomerPayments AS p INNER JOIN [@Customers] AS c ON p.CustomerID = c.CustomerID
WHERE p.Amount <> 0
UNION ALL
SELECT 'SUPPLIER_PAYMENT' AS SourceType, p.PaymentID AS SourceID, p.PaymentNumber AS SourceNumber, p.PaymentDate AS SourceDate, s.SupplierName AS Party, 1 AS LineOrder, 2100 AS AccountCode, p.Amount AS Debit, CCur(0) AS Credit, s.SupplierName AS LineText, 0 AS CostCenter
FROM SupplierPayments AS p INNER JOIN [@Suppliers] AS s ON p.SupplierID = s.SupplierID
WHERE p.Amount <> 0
UNION ALL
SELECT 'SUPPLIER_PAYMENT' AS SourceType, p.PaymentID AS SourceID, p.PaymentNumber AS SourceNumber, p.PaymentDate AS SourceDate, s.SupplierName AS Party, 2 AS LineOrder, IIf(p.CashBoxID Is Null, IIf(p.PaymentMethodID Is Null Or p.PaymentMethodID = 1, 1190, IIf(p.BankID Is Null, 1200, 120000 + p.BankID)), 110000 + p.CashBoxID) AS AccountCode, CCur(0) AS Debit, p.Amount AS Credit, s.SupplierName AS LineText, 0 AS CostCenter
FROM SupplierPayments AS p INNER JOIN [@Suppliers] AS s ON p.SupplierID = s.SupplierID
WHERE p.Amount <> 0
```

## qryJournalExpense

أسطر قيود المصروفات (عدا المسجلة بسند نقدية)

```sql
SELECT 'EXPENSE' AS SourceType, e.ExpenseID AS SourceID, e.ExpenseNumber AS SourceNumber, e.ExpenseDate AS SourceDate, t.ExpenseTypeName AS Party, 1 AS LineOrder, 530000 + e.ExpenseTypeID AS AccountCode, e.Amount AS Debit, CCur(0) AS Credit, t.ExpenseTypeName AS LineText, IIf(e.CostCenterID Is Null, 0, e.CostCenterID) AS CostCenter
FROM (Expenses AS e INNER JOIN [@ExpenseTypes] AS t ON e.ExpenseTypeID = t.ExpenseTypeID) LEFT JOIN CashVouchers AS v ON e.ExpenseID = v.ExpenseID
WHERE v.CashVoucherID Is Null AND e.Amount <> 0
UNION ALL
SELECT 'EXPENSE' AS SourceType, e.ExpenseID AS SourceID, e.ExpenseNumber AS SourceNumber, e.ExpenseDate AS SourceDate, t.ExpenseTypeName AS Party, 2 AS LineOrder, 1500 AS AccountCode, e.Tax AS Debit, CCur(0) AS Credit, 'ضريبة المدخلات' AS LineText, IIf(e.CostCenterID Is Null, 0, e.CostCenterID) AS CostCenter
FROM (Expenses AS e INNER JOIN [@ExpenseTypes] AS t ON e.ExpenseTypeID = t.ExpenseTypeID) LEFT JOIN CashVouchers AS v ON e.ExpenseID = v.ExpenseID
WHERE v.CashVoucherID Is Null AND e.Tax <> 0
UNION ALL
SELECT 'EXPENSE' AS SourceType, e.ExpenseID AS SourceID, e.ExpenseNumber AS SourceNumber, e.ExpenseDate AS SourceDate, t.ExpenseTypeName AS Party, 3 AS LineOrder, IIf(e.CashBoxID Is Null, IIf(e.PaymentMethodID Is Null Or e.PaymentMethodID = 1, 1190, IIf(e.BankID Is Null, 1200, 120000 + e.BankID)), 110000 + e.CashBoxID) AS AccountCode, CCur(0) AS Debit, e.TotalAmount AS Credit, e.Description AS LineText, IIf(e.CostCenterID Is Null, 0, e.CostCenterID) AS CostCenter
FROM (Expenses AS e INNER JOIN [@ExpenseTypes] AS t ON e.ExpenseTypeID = t.ExpenseTypeID) LEFT JOIN CashVouchers AS v ON e.ExpenseID = v.ExpenseID
WHERE v.CashVoucherID Is Null AND e.TotalAmount <> 0
```

## qryJournalCashVoucher

أسطر قيود سندات النقدية (قبض وصرف وتحويل)

```sql
SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 1 AS LineOrder, 110000 + v.CashBoxID AS AccountCode, v.Amount AS Debit, CCur(0) AS Credit, v.PartyName AS LineText, IIf(v.CostCenterID Is Null, 0, v.CostCenterID) AS CostCenter
FROM CashVouchers AS v
WHERE v.VoucherType = 'IN'
UNION ALL
SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 2 AS LineOrder, IIf(v.Category = 'OWNER', 3100, IIf(v.Category = 'OVERAGE', 4300, IIf(v.Category = 'ADVANCE', 1600, 4200))) AS AccountCode, CCur(0) AS Debit, v.Amount AS Credit, v.Description AS LineText, IIf(v.CostCenterID Is Null, 0, v.CostCenterID) AS CostCenter
FROM CashVouchers AS v
WHERE v.VoucherType = 'IN'
UNION ALL
SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 1 AS LineOrder, IIf(v.Category = 'OWNER', 3100, IIf(v.Category = 'ADVANCE', 1600, IIf(v.Category = 'SHORTAGE', 5400, IIf(v.Category = 'COMMISSION', 2330, IIf(v.Category = 'EXPENSE' AND x.ExpenseTypeID Is Not Null, 530000 + x.ExpenseTypeID, 5900))))) AS AccountCode, v.Amount - CCur(Nz(x.Tax, 0)) AS Debit, CCur(0) AS Credit, v.Description AS LineText, IIf(v.CostCenterID Is Null, 0, v.CostCenterID) AS CostCenter
FROM CashVouchers AS v LEFT JOIN Expenses AS x ON v.ExpenseID = x.ExpenseID
WHERE v.VoucherType = 'OUT'
UNION ALL
SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 2 AS LineOrder, 110000 + v.CashBoxID AS AccountCode, CCur(0) AS Debit, v.Amount AS Credit, v.PartyName AS LineText, IIf(v.CostCenterID Is Null, 0, v.CostCenterID) AS CostCenter
FROM CashVouchers AS v LEFT JOIN Expenses AS x ON v.ExpenseID = x.ExpenseID
WHERE v.VoucherType = 'OUT'
UNION ALL
SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 3 AS LineOrder, 1500 AS AccountCode, x.Tax AS Debit, CCur(0) AS Credit, 'ضريبة المدخلات' AS LineText, IIf(v.CostCenterID Is Null, 0, v.CostCenterID) AS CostCenter
FROM CashVouchers AS v LEFT JOIN Expenses AS x ON v.ExpenseID = x.ExpenseID
WHERE v.VoucherType = 'OUT' AND x.Tax <> 0
UNION ALL
SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 1 AS LineOrder, 110000 + v.ToCashBoxID AS AccountCode, v.Amount AS Debit, CCur(0) AS Credit, v.Description AS LineText, IIf(v.CostCenterID Is Null, 0, v.CostCenterID) AS CostCenter
FROM CashVouchers AS v
WHERE v.VoucherType = 'TRANSFER'
UNION ALL
SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 2 AS LineOrder, 110000 + v.CashBoxID AS AccountCode, CCur(0) AS Debit, v.Amount AS Credit, v.Description AS LineText, IIf(v.CostCenterID Is Null, 0, v.CostCenterID) AS CostCenter
FROM CashVouchers AS v
WHERE v.VoucherType = 'TRANSFER'
```

## qryJournalStock

أسطر قيود حركات المخزون اليدوية وتسويات الجرد

```sql
SELECT 'STOCK_MOVE' AS SourceType, i.TransactionID AS SourceID, i.ReferenceNumber AS SourceNumber, i.TransactionDate AS SourceDate, p.ProductName AS Party, 1 AS LineOrder, 1400 AS AccountCode, IIf(i.Quantity * i.UnitCost > 0, i.Quantity * i.UnitCost, 0) AS Debit, IIf(i.Quantity * i.UnitCost < 0, -i.Quantity * i.UnitCost, 0) AS Credit, i.Notes AS LineText, 0 AS CostCenter
FROM InventoryTransactions AS i INNER JOIN Products AS p ON i.ProductID = p.ProductID
WHERE i.ReferenceType = 'MANUAL' AND i.Quantity * i.UnitCost <> 0
UNION ALL
SELECT 'STOCK_MOVE' AS SourceType, i.TransactionID AS SourceID, i.ReferenceNumber AS SourceNumber, i.TransactionDate AS SourceDate, p.ProductName AS Party, 2 AS LineOrder, IIf(i.TransactionTypeID = 8, 3900, 5200) AS AccountCode, IIf(i.Quantity * i.UnitCost < 0, -i.Quantity * i.UnitCost, 0) AS Debit, IIf(i.Quantity * i.UnitCost > 0, i.Quantity * i.UnitCost, 0) AS Credit, i.Notes AS LineText, 0 AS CostCenter
FROM InventoryTransactions AS i INNER JOIN Products AS p ON i.ProductID = p.ProductID
WHERE i.ReferenceType = 'MANUAL' AND i.Quantity * i.UnitCost <> 0
UNION ALL
SELECT 'STOCK_COUNT' AS SourceType, k.StockCountID AS SourceID, k.CountNumber AS SourceNumber, k.CountDate AS SourceDate, 'تسوية الجرد' AS Party, 1 AS LineOrder, 1400 AS AccountCode, IIf(k.CountValue > 0, k.CountValue, 0) AS Debit, IIf(k.CountValue < 0, -k.CountValue, 0) AS Credit, 'المخزون' AS LineText, 0 AS CostCenter
FROM qryStockCountValue AS k
WHERE k.CountValue <> 0
UNION ALL
SELECT 'STOCK_COUNT' AS SourceType, k.StockCountID AS SourceID, k.CountNumber AS SourceNumber, k.CountDate AS SourceDate, 'تسوية الجرد' AS Party, 2 AS LineOrder, 5200 AS AccountCode, IIf(k.CountValue < 0, -k.CountValue, 0) AS Debit, IIf(k.CountValue > 0, k.CountValue, 0) AS Credit, 'فروقات الجرد' AS LineText, 0 AS CostCenter
FROM qryStockCountValue AS k
WHERE k.CountValue <> 0
```

## qryJournalOpening

أسطر قيود الأرصدة الافتتاحية للصناديق والعملاء والموردين

```sql
SELECT 'BOX_OPENING' AS SourceType, b.CashBoxID AS SourceID, b.BoxName AS SourceNumber, b.OpeningDate AS SourceDate, b.BoxName AS Party, 1 AS LineOrder, 110000 + b.CashBoxID AS AccountCode, b.OpeningBalance AS Debit, CCur(0) AS Credit, b.BoxName AS LineText, 0 AS CostCenter
FROM [@CashBoxes] AS b
WHERE b.OpeningBalance <> 0
UNION ALL
SELECT 'BOX_OPENING' AS SourceType, b.CashBoxID AS SourceID, b.BoxName AS SourceNumber, b.OpeningDate AS SourceDate, b.BoxName AS Party, 2 AS LineOrder, 3900 AS AccountCode, CCur(0) AS Debit, b.OpeningBalance AS Credit, 'رصيد افتتاحي' AS LineText, 0 AS CostCenter
FROM [@CashBoxes] AS b
WHERE b.OpeningBalance <> 0
UNION ALL
SELECT 'CUSTOMER_OPENING' AS SourceType, c.CustomerID AS SourceID, c.CustomerName AS SourceNumber, c.CreatedAt AS SourceDate, c.CustomerName AS Party, 1 AS LineOrder, 1300 AS AccountCode, IIf(c.OpeningBalance > 0, c.OpeningBalance, 0) AS Debit, IIf(c.OpeningBalance < 0, -c.OpeningBalance, 0) AS Credit, c.CustomerName AS LineText, 0 AS CostCenter
FROM [@Customers] AS c
WHERE c.OpeningBalance <> 0
UNION ALL
SELECT 'CUSTOMER_OPENING' AS SourceType, c.CustomerID AS SourceID, c.CustomerName AS SourceNumber, c.CreatedAt AS SourceDate, c.CustomerName AS Party, 2 AS LineOrder, 3900 AS AccountCode, IIf(c.OpeningBalance < 0, -c.OpeningBalance, 0) AS Debit, IIf(c.OpeningBalance > 0, c.OpeningBalance, 0) AS Credit, 'رصيد افتتاحي' AS LineText, 0 AS CostCenter
FROM [@Customers] AS c
WHERE c.OpeningBalance <> 0
UNION ALL
SELECT 'SUPPLIER_OPENING' AS SourceType, s.SupplierID AS SourceID, s.SupplierName AS SourceNumber, s.CreatedAt AS SourceDate, s.SupplierName AS Party, 1 AS LineOrder, 2100 AS AccountCode, IIf(s.OpeningBalance < 0, -s.OpeningBalance, 0) AS Debit, IIf(s.OpeningBalance > 0, s.OpeningBalance, 0) AS Credit, s.SupplierName AS LineText, 0 AS CostCenter
FROM [@Suppliers] AS s
WHERE s.OpeningBalance <> 0
UNION ALL
SELECT 'SUPPLIER_OPENING' AS SourceType, s.SupplierID AS SourceID, s.SupplierName AS SourceNumber, s.CreatedAt AS SourceDate, s.SupplierName AS Party, 2 AS LineOrder, 3900 AS AccountCode, IIf(s.OpeningBalance > 0, s.OpeningBalance, 0) AS Debit, IIf(s.OpeningBalance < 0, -s.OpeningBalance, 0) AS Credit, 'رصيد افتتاحي' AS LineText, 0 AS CostCenter
FROM [@Suppliers] AS s
WHERE s.OpeningBalance <> 0
UNION ALL
SELECT 'BANK_OPENING' AS SourceType, k.BankID AS SourceID, k.BankName AS SourceNumber, k.OpeningDate AS SourceDate, k.BankName AS Party, 1 AS LineOrder, 120000 + k.BankID AS AccountCode, IIf(k.OpeningBalance > 0, k.OpeningBalance, 0) AS Debit, IIf(k.OpeningBalance < 0, -k.OpeningBalance, 0) AS Credit, k.BankName AS LineText, 0 AS CostCenter
FROM [@Banks] AS k
WHERE k.OpeningBalance <> 0
UNION ALL
SELECT 'BANK_OPENING' AS SourceType, k.BankID AS SourceID, k.BankName AS SourceNumber, k.OpeningDate AS SourceDate, k.BankName AS Party, 2 AS LineOrder, 3900 AS AccountCode, IIf(k.OpeningBalance < 0, -k.OpeningBalance, 0) AS Debit, IIf(k.OpeningBalance > 0, k.OpeningBalance, 0) AS Credit, 'رصيد افتتاحي' AS LineText, 0 AS CostCenter
FROM [@Banks] AS k
WHERE k.OpeningBalance <> 0
```

## qryManualEntryLines

أسطر القيود اليدوية مع رأس كل قيد

```sql
SELECT h.ManualEntryID, h.EntryNumber, h.EntryDate, h.Description, l.LineNumber AS LineNo,
       l.AccountCode AS LineAccount, l.Debit AS LineDebit, l.Credit AS LineCredit, l.LineText AS LineNote,
       IIf(l.CostCenterID Is Null, 0, l.CostCenterID) AS LineCenter
FROM ManualEntries AS h INNER JOIN ManualEntryLines AS l ON h.ManualEntryID = l.ManualEntryID
```

## qryJournalManual

أسطر القيود اليدوية

```sql
SELECT 'MANUAL' AS SourceType, m.ManualEntryID AS SourceID, m.EntryNumber AS SourceNumber, m.EntryDate AS SourceDate, m.Description AS Party, m.LineNo AS LineOrder, m.LineAccount AS AccountCode, m.LineDebit AS Debit, m.LineCredit AS Credit, m.LineNote AS LineText, m.LineCenter AS CostCenter
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
SELECT 'YEAR_CLOSE' AS SourceType, y.YearClosingID AS SourceID, y.ClosingNumber AS SourceNumber, y.ClosingDate AS SourceDate, y.Notes AS Party, y.LineNo AS LineOrder, y.LineAccount AS AccountCode, y.LineDebit AS Debit, y.LineCredit AS Credit, y.LineNote AS LineText, 0 AS CostCenter
FROM qryYearCloseLines AS y
WHERE y.LineDebit + y.LineCredit <> 0
```

## qryJournalVatReturn

أسطر قيود الإقرار الضريبي المعتمد (التسوية) وسداده

```sql
SELECT 'VAT_RETURN' AS SourceType, v.VatReturnID AS SourceID, v.ReturnNumber AS SourceNumber, v.FiledDate AS SourceDate, v.ReturnNumber AS Party, 1 AS LineOrder, 2200 AS AccountCode, IIf(v.SalesStdVAT > 0, v.SalesStdVAT, 0) AS Debit, IIf(v.SalesStdVAT < 0, -v.SalesStdVAT, 0) AS Credit, 'ضريبة المخرجات للفترة' AS LineText, 0 AS CostCenter
FROM VatReturns AS v
WHERE v.Status = 'FILED' AND v.SalesStdVAT <> 0
UNION ALL
SELECT 'VAT_RETURN' AS SourceType, v.VatReturnID AS SourceID, v.ReturnNumber AS SourceNumber, v.FiledDate AS SourceDate, v.ReturnNumber AS Party, 2 AS LineOrder, 1500 AS AccountCode, IIf(v.PurchStdVAT < 0, -v.PurchStdVAT, 0) AS Debit, IIf(v.PurchStdVAT > 0, v.PurchStdVAT, 0) AS Credit, 'ضريبة المدخلات للفترة' AS LineText, 0 AS CostCenter
FROM VatReturns AS v
WHERE v.Status = 'FILED' AND v.PurchStdVAT <> 0
UNION ALL
SELECT 'VAT_RETURN' AS SourceType, v.VatReturnID AS SourceID, v.ReturnNumber AS SourceNumber, v.FiledDate AS SourceDate, v.ReturnNumber AS Party, 3 AS LineOrder, 2200 AS AccountCode, IIf(v.Corrections > 0, v.Corrections, 0) AS Debit, IIf(v.Corrections < 0, -v.Corrections, 0) AS Credit, 'تصحيحات من الفترات السابقة' AS LineText, 0 AS CostCenter
FROM VatReturns AS v
WHERE v.Status = 'FILED' AND v.Corrections <> 0
UNION ALL
SELECT 'VAT_RETURN' AS SourceType, v.VatReturnID AS SourceID, v.ReturnNumber AS SourceNumber, v.FiledDate AS SourceDate, v.ReturnNumber AS Party, 4 AS LineOrder, 2250 AS AccountCode, IIf((v.SalesStdVAT - v.PurchStdVAT + v.Corrections) < 0, -(v.SalesStdVAT - v.PurchStdVAT + v.Corrections), 0) AS Debit, IIf((v.SalesStdVAT - v.PurchStdVAT + v.Corrections) > 0, (v.SalesStdVAT - v.PurchStdVAT + v.Corrections), 0) AS Credit, 'صافي ضريبة الفترة' AS LineText, 0 AS CostCenter
FROM VatReturns AS v
WHERE v.Status = 'FILED' AND (v.SalesStdVAT - v.PurchStdVAT + v.Corrections) <> 0
UNION ALL
SELECT 'VAT_PAYMENT' AS SourceType, v.VatReturnID AS SourceID, v.ReturnNumber AS SourceNumber, v.PaidDate AS SourceDate, v.ReturnNumber AS Party, 1 AS LineOrder, 2250 AS AccountCode, v.PaidAmount AS Debit, CCur(0) AS Credit, 'سداد ضريبة القيمة المضافة' AS LineText, 0 AS CostCenter
FROM VatReturns AS v
WHERE v.Status = 'FILED' AND v.PaidAmount <> 0
UNION ALL
SELECT 'VAT_PAYMENT' AS SourceType, v.VatReturnID AS SourceID, v.ReturnNumber AS SourceNumber, v.PaidDate AS SourceDate, v.ReturnNumber AS Party, 2 AS LineOrder, v.PaidAccount AS AccountCode, CCur(0) AS Debit, v.PaidAmount AS Credit, v.FilingRef AS LineText, 0 AS CostCenter
FROM VatReturns AS v
WHERE v.Status = 'FILED' AND v.PaidAmount <> 0
```

## qryJournalCheque

أسطر قيود الشيكات: الاستلام أو الإصدار، ثم التحصيل أو الارتداد

```sql
SELECT 'CHEQUE' AS SourceType, q.ChequeID AS SourceID, q.ChequeRef AS SourceNumber, q.IssueDate AS SourceDate, q.ChequeNo AS Party, 1 AS LineOrder, 1250 AS AccountCode, q.Amount AS Debit, CCur(0) AS Credit, 'شيك وارد تحت التحصيل' AS LineText, 0 AS CostCenter
FROM Cheques AS q
WHERE q.Direction = 'IN'
UNION ALL
SELECT 'CHEQUE' AS SourceType, q.ChequeID AS SourceID, q.ChequeRef AS SourceNumber, q.IssueDate AS SourceDate, q.ChequeNo AS Party, 2 AS LineOrder, 1300 AS AccountCode, CCur(0) AS Debit, q.Amount AS Credit, q.ChequeNo AS LineText, 0 AS CostCenter
FROM Cheques AS q
WHERE q.Direction = 'IN'
UNION ALL
SELECT 'CHEQUE' AS SourceType, q.ChequeID AS SourceID, q.ChequeRef AS SourceNumber, q.IssueDate AS SourceDate, q.ChequeNo AS Party, 1 AS LineOrder, 2100 AS AccountCode, q.Amount AS Debit, CCur(0) AS Credit, q.ChequeNo AS LineText, 0 AS CostCenter
FROM Cheques AS q
WHERE q.Direction = 'OUT'
UNION ALL
SELECT 'CHEQUE' AS SourceType, q.ChequeID AS SourceID, q.ChequeRef AS SourceNumber, q.IssueDate AS SourceDate, q.ChequeNo AS Party, 2 AS LineOrder, 2110 AS AccountCode, CCur(0) AS Debit, q.Amount AS Credit, 'شيك صادر' AS LineText, 0 AS CostCenter
FROM Cheques AS q
WHERE q.Direction = 'OUT'
UNION ALL
SELECT 'CHEQUE_STATUS' AS SourceType, q.ChequeID AS SourceID, q.ChequeRef AS SourceNumber, q.StatusDate AS SourceDate, q.ChequeNo AS Party, 1 AS LineOrder, 120000 + q.BankID AS AccountCode, q.Amount AS Debit, CCur(0) AS Credit, 'تحصيل شيك' AS LineText, 0 AS CostCenter
FROM Cheques AS q
WHERE q.Direction = 'IN' AND q.Status = 'COLLECTED'
UNION ALL
SELECT 'CHEQUE_STATUS' AS SourceType, q.ChequeID AS SourceID, q.ChequeRef AS SourceNumber, q.StatusDate AS SourceDate, q.ChequeNo AS Party, 2 AS LineOrder, 1250 AS AccountCode, CCur(0) AS Debit, q.Amount AS Credit, q.ChequeNo AS LineText, 0 AS CostCenter
FROM Cheques AS q
WHERE q.Direction = 'IN' AND q.Status = 'COLLECTED'
UNION ALL
SELECT 'CHEQUE_STATUS' AS SourceType, q.ChequeID AS SourceID, q.ChequeRef AS SourceNumber, q.StatusDate AS SourceDate, q.ChequeNo AS Party, 1 AS LineOrder, 1300 AS AccountCode, q.Amount AS Debit, CCur(0) AS Credit, 'شيك مرتد' AS LineText, 0 AS CostCenter
FROM Cheques AS q
WHERE q.Direction = 'IN' AND q.Status = 'BOUNCED'
UNION ALL
SELECT 'CHEQUE_STATUS' AS SourceType, q.ChequeID AS SourceID, q.ChequeRef AS SourceNumber, q.StatusDate AS SourceDate, q.ChequeNo AS Party, 2 AS LineOrder, 1250 AS AccountCode, CCur(0) AS Debit, q.Amount AS Credit, q.ChequeNo AS LineText, 0 AS CostCenter
FROM Cheques AS q
WHERE q.Direction = 'IN' AND q.Status = 'BOUNCED'
UNION ALL
SELECT 'CHEQUE_STATUS' AS SourceType, q.ChequeID AS SourceID, q.ChequeRef AS SourceNumber, q.StatusDate AS SourceDate, q.ChequeNo AS Party, 1 AS LineOrder, 2110 AS AccountCode, q.Amount AS Debit, CCur(0) AS Credit, q.ChequeNo AS LineText, 0 AS CostCenter
FROM Cheques AS q
WHERE q.Direction = 'OUT' AND q.Status = 'COLLECTED'
UNION ALL
SELECT 'CHEQUE_STATUS' AS SourceType, q.ChequeID AS SourceID, q.ChequeRef AS SourceNumber, q.StatusDate AS SourceDate, q.ChequeNo AS Party, 2 AS LineOrder, 120000 + q.BankID AS AccountCode, CCur(0) AS Debit, q.Amount AS Credit, 'صرف شيك' AS LineText, 0 AS CostCenter
FROM Cheques AS q
WHERE q.Direction = 'OUT' AND q.Status = 'COLLECTED'
UNION ALL
SELECT 'CHEQUE_STATUS' AS SourceType, q.ChequeID AS SourceID, q.ChequeRef AS SourceNumber, q.StatusDate AS SourceDate, q.ChequeNo AS Party, 1 AS LineOrder, 2110 AS AccountCode, q.Amount AS Debit, CCur(0) AS Credit, q.ChequeNo AS LineText, 0 AS CostCenter
FROM Cheques AS q
WHERE q.Direction = 'OUT' AND q.Status = 'BOUNCED'
UNION ALL
SELECT 'CHEQUE_STATUS' AS SourceType, q.ChequeID AS SourceID, q.ChequeRef AS SourceNumber, q.StatusDate AS SourceDate, q.ChequeNo AS Party, 2 AS LineOrder, 2100 AS AccountCode, CCur(0) AS Debit, q.Amount AS Credit, 'شيك مرتد' AS LineText, 0 AS CostCenter
FROM Cheques AS q
WHERE q.Direction = 'OUT' AND q.Status = 'BOUNCED'
```

## qryJournalAsset

أسطر قيود اقتناء الأصول الثابتة وبيعها أو استبعادها

```sql
SELECT 'ASSET' AS SourceType, a.AssetID AS SourceID, a.AssetCode AS SourceNumber, a.PurchaseDate AS SourceDate, a.AssetName AS Party, 1 AS LineOrder, a.AssetAccount AS AccountCode, a.Cost AS Debit, CCur(0) AS Credit, a.AssetName AS LineText, IIf(a.CostCenterID Is Null, 0, a.CostCenterID) AS CostCenter
FROM FixedAssets AS a
WHERE a.Cost <> 0
UNION ALL
SELECT 'ASSET' AS SourceType, a.AssetID AS SourceID, a.AssetCode AS SourceNumber, a.PurchaseDate AS SourceDate, a.AssetName AS Party, 2 AS LineOrder, 1500 AS AccountCode, a.InputVAT AS Debit, CCur(0) AS Credit, 'ضريبة المدخلات' AS LineText, IIf(a.CostCenterID Is Null, 0, a.CostCenterID) AS CostCenter
FROM FixedAssets AS a
WHERE a.InputVAT <> 0 AND a.SourceType <> 'OPENING'
UNION ALL
SELECT 'ASSET' AS SourceType, a.AssetID AS SourceID, a.AssetCode AS SourceNumber, a.PurchaseDate AS SourceDate, a.AssetName AS Party, 3 AS LineOrder, IIf(a.SourceType = 'BANK', 120000 + a.BankID, IIf(a.SourceType = 'CASHBOX', 110000 + a.CashBoxID, IIf(a.SourceType = 'ACCOUNT', a.CounterAccount, 3900))) AS AccountCode, CCur(0) AS Debit, IIf(a.SourceType = 'OPENING', a.Cost - a.OpeningAccumDep, a.Cost + a.InputVAT) AS Credit, a.Notes AS LineText, IIf(a.CostCenterID Is Null, 0, a.CostCenterID) AS CostCenter
FROM FixedAssets AS a
WHERE IIf(a.SourceType = 'OPENING', a.Cost - a.OpeningAccumDep, a.Cost + a.InputVAT) <> 0
UNION ALL
SELECT 'ASSET' AS SourceType, a.AssetID AS SourceID, a.AssetCode AS SourceNumber, a.PurchaseDate AS SourceDate, a.AssetName AS Party, 4 AS LineOrder, 1790 AS AccountCode, CCur(0) AS Debit, a.OpeningAccumDep AS Credit, 'إهلاك سابق' AS LineText, IIf(a.CostCenterID Is Null, 0, a.CostCenterID) AS CostCenter
FROM FixedAssets AS a
WHERE a.SourceType = 'OPENING' AND a.OpeningAccumDep <> 0
UNION ALL
SELECT 'ASSET_DISPOSAL' AS SourceType, a.AssetID AS SourceID, a.AssetCode AS SourceNumber, a.DisposalDate AS SourceDate, a.AssetName AS Party, 1 AS LineOrder, 1790 AS AccountCode, a.DisposalAccumDep AS Debit, CCur(0) AS Credit, 'مجمع إهلاك الأصل' AS LineText, IIf(a.CostCenterID Is Null, 0, a.CostCenterID) AS CostCenter
FROM FixedAssets AS a
WHERE a.Status = 'DISPOSED' AND a.DisposalAccumDep <> 0
UNION ALL
SELECT 'ASSET_DISPOSAL' AS SourceType, a.AssetID AS SourceID, a.AssetCode AS SourceNumber, a.DisposalDate AS SourceDate, a.AssetName AS Party, 2 AS LineOrder, IIf(a.DisposalTo = 'BANK', 120000 + a.DisposalBankID, 110000 + a.DisposalCashBoxID) AS AccountCode, a.DisposalProceeds AS Debit, CCur(0) AS Credit, 'ثمن بيع الأصل' AS LineText, IIf(a.CostCenterID Is Null, 0, a.CostCenterID) AS CostCenter
FROM FixedAssets AS a
WHERE a.Status = 'DISPOSED' AND a.DisposalProceeds <> 0
UNION ALL
SELECT 'ASSET_DISPOSAL' AS SourceType, a.AssetID AS SourceID, a.AssetCode AS SourceNumber, a.DisposalDate AS SourceDate, a.AssetName AS Party, 3 AS LineOrder, a.AssetAccount AS AccountCode, CCur(0) AS Debit, a.Cost AS Credit, a.AssetName AS LineText, IIf(a.CostCenterID Is Null, 0, a.CostCenterID) AS CostCenter
FROM FixedAssets AS a
WHERE a.Status = 'DISPOSED'
UNION ALL
SELECT 'ASSET_DISPOSAL' AS SourceType, a.AssetID AS SourceID, a.AssetCode AS SourceNumber, a.DisposalDate AS SourceDate, a.AssetName AS Party, 4 AS LineOrder, 4500 AS AccountCode, CCur(0) AS Debit, (a.DisposalProceeds + a.DisposalAccumDep - a.Cost) AS Credit, 'ربح بيع الأصل' AS LineText, IIf(a.CostCenterID Is Null, 0, a.CostCenterID) AS CostCenter
FROM FixedAssets AS a
WHERE a.Status = 'DISPOSED' AND (a.DisposalProceeds + a.DisposalAccumDep - a.Cost) > 0
UNION ALL
SELECT 'ASSET_DISPOSAL' AS SourceType, a.AssetID AS SourceID, a.AssetCode AS SourceNumber, a.DisposalDate AS SourceDate, a.AssetName AS Party, 5 AS LineOrder, 5650 AS AccountCode, -(a.DisposalProceeds + a.DisposalAccumDep - a.Cost) AS Debit, CCur(0) AS Credit, 'خسارة بيع / استبعاد الأصل' AS LineText, IIf(a.CostCenterID Is Null, 0, a.CostCenterID) AS CostCenter
FROM FixedAssets AS a
WHERE a.Status = 'DISPOSED' AND (a.DisposalProceeds + a.DisposalAccumDep - a.Cost) < 0
```

## qryDepreciationLines

أسطر قيود الإهلاك الشهرية مع اسم الأصل

```sql
SELECT r.RunID, r.RunNumber, r.RunMonth, d.LineNo, d.AssetID, d.Amount, a.AssetName,
       IIf(a.CostCenterID Is Null, 0, a.CostCenterID) AS LineCenter
FROM (DepreciationRuns AS r INNER JOIN AssetDepreciations AS d ON r.RunID = d.RunID)
     INNER JOIN FixedAssets AS a ON d.AssetID = a.AssetID
```

## qryJournalDepreciation

أسطر قيود الإهلاك الشهرية: مصروف الإهلاك ومجمع الإهلاك لكل أصل

```sql
SELECT 'DEPRECIATION' AS SourceType, d.RunID AS SourceID, d.RunNumber AS SourceNumber, d.RunMonth AS SourceDate, 'الإهلاك الشهري' AS Party, 2 * d.LineNo - 1 AS LineOrder, 5600 AS AccountCode, d.Amount AS Debit, CCur(0) AS Credit, d.AssetName AS LineText, d.LineCenter AS CostCenter
FROM qryDepreciationLines AS d
WHERE d.Amount <> 0
UNION ALL
SELECT 'DEPRECIATION' AS SourceType, d.RunID AS SourceID, d.RunNumber AS SourceNumber, d.RunMonth AS SourceDate, 'الإهلاك الشهري' AS Party, 2 * d.LineNo AS LineOrder, 1790 AS AccountCode, CCur(0) AS Debit, d.Amount AS Credit, d.AssetName AS LineText, d.LineCenter AS CostCenter
FROM qryDepreciationLines AS d
WHERE d.Amount <> 0
```

## qryPayrollTotals

مجاميع كل مسير رواتب لقيده

```sql
SELECT PayrollRunID, Sum(Basic + Housing - AbsenceDeduction) AS SumSalaries,
       Sum(OtherAllow + Overtime + Additions) AS SumAllowances, Sum(GosiEmployer) AS SumGosiER,
       Sum(GosiEmployee + GosiEmployer) AS SumGosi, Sum(AdvanceDeduction) AS SumAdvance,
       Sum(OtherDeduction) AS SumOtherDed, Sum(NetPay) AS SumNet, Count(*) AS LineCount
FROM PayrollLines
GROUP BY PayrollRunID
```

## qryPayrollCenterTotals

مجاميع كل مسير رواتب لكل مركز تكلفة لقيده

```sql
SELECT PayrollRunID, IIf(CostCenterID Is Null, 0, CostCenterID) AS CenterKey,
       Sum(Basic + Housing - AbsenceDeduction) AS SumSalaries, Sum(OtherAllow + Overtime + Additions) AS SumAllowances,
       Sum(GosiEmployer) AS SumGosiER, Sum(GosiEmployee + GosiEmployer) AS SumGosi, Sum(AdvanceDeduction) AS SumAdvance,
       Sum(OtherDeduction) AS SumOtherDed, Sum(NetPay) AS SumNet
FROM PayrollLines
GROUP BY PayrollRunID, IIf(CostCenterID Is Null, 0, CostCenterID)
```

## qryJournalPayroll

أسطر قيود مسيرات الرواتب المرحَّلة وصرفها

```sql
SELECT 'PAYROLL' AS SourceType, r.PayrollRunID AS SourceID, r.RunNumber AS SourceNumber, r.PayMonth AS SourceDate, 'مسير الرواتب' AS Party, 1 + 10 * t.CenterKey AS LineOrder, 5500 AS AccountCode, t.SumSalaries AS Debit, CCur(0) AS Credit, 'الرواتب' AS LineText, t.CenterKey AS CostCenter
FROM PayrollRuns AS r INNER JOIN qryPayrollCenterTotals AS t ON r.PayrollRunID = t.PayrollRunID
WHERE r.Status = 'POSTED' AND t.SumSalaries <> 0
UNION ALL
SELECT 'PAYROLL' AS SourceType, r.PayrollRunID AS SourceID, r.RunNumber AS SourceNumber, r.PayMonth AS SourceDate, 'مسير الرواتب' AS Party, 2 + 10 * t.CenterKey AS LineOrder, 5510 AS AccountCode, t.SumAllowances AS Debit, CCur(0) AS Credit, 'البدلات والإضافي' AS LineText, t.CenterKey AS CostCenter
FROM PayrollRuns AS r INNER JOIN qryPayrollCenterTotals AS t ON r.PayrollRunID = t.PayrollRunID
WHERE r.Status = 'POSTED' AND t.SumAllowances <> 0
UNION ALL
SELECT 'PAYROLL' AS SourceType, r.PayrollRunID AS SourceID, r.RunNumber AS SourceNumber, r.PayMonth AS SourceDate, 'مسير الرواتب' AS Party, 3 + 10 * t.CenterKey AS LineOrder, 5520 AS AccountCode, t.SumGosiER AS Debit, CCur(0) AS Credit, 'التأمينات - حصة المنشأة' AS LineText, t.CenterKey AS CostCenter
FROM PayrollRuns AS r INNER JOIN qryPayrollCenterTotals AS t ON r.PayrollRunID = t.PayrollRunID
WHERE r.Status = 'POSTED' AND t.SumGosiER <> 0
UNION ALL
SELECT 'PAYROLL' AS SourceType, r.PayrollRunID AS SourceID, r.RunNumber AS SourceNumber, r.PayMonth AS SourceDate, 'مسير الرواتب' AS Party, 4 + 10 * t.CenterKey AS LineOrder, 2320 AS AccountCode, CCur(0) AS Debit, t.SumGosi AS Credit, 'التأمينات المستحقة' AS LineText, t.CenterKey AS CostCenter
FROM PayrollRuns AS r INNER JOIN qryPayrollCenterTotals AS t ON r.PayrollRunID = t.PayrollRunID
WHERE r.Status = 'POSTED' AND t.SumGosi <> 0
UNION ALL
SELECT 'PAYROLL' AS SourceType, r.PayrollRunID AS SourceID, r.RunNumber AS SourceNumber, r.PayMonth AS SourceDate, 'مسير الرواتب' AS Party, 5 + 10 * t.CenterKey AS LineOrder, 1600 AS AccountCode, CCur(0) AS Debit, t.SumAdvance AS Credit, 'خصم السلف' AS LineText, t.CenterKey AS CostCenter
FROM PayrollRuns AS r INNER JOIN qryPayrollCenterTotals AS t ON r.PayrollRunID = t.PayrollRunID
WHERE r.Status = 'POSTED' AND t.SumAdvance <> 0
UNION ALL
SELECT 'PAYROLL' AS SourceType, r.PayrollRunID AS SourceID, r.RunNumber AS SourceNumber, r.PayMonth AS SourceDate, 'مسير الرواتب' AS Party, 6 + 10 * t.CenterKey AS LineOrder, 4200 AS AccountCode, CCur(0) AS Debit, t.SumOtherDed AS Credit, 'جزاءات وخصومات' AS LineText, t.CenterKey AS CostCenter
FROM PayrollRuns AS r INNER JOIN qryPayrollCenterTotals AS t ON r.PayrollRunID = t.PayrollRunID
WHERE r.Status = 'POSTED' AND t.SumOtherDed <> 0
UNION ALL
SELECT 'PAYROLL' AS SourceType, r.PayrollRunID AS SourceID, r.RunNumber AS SourceNumber, r.PayMonth AS SourceDate, 'مسير الرواتب' AS Party, 7 + 10 * t.CenterKey AS LineOrder, 2310 AS AccountCode, CCur(0) AS Debit, t.SumNet AS Credit, 'صافي الرواتب' AS LineText, t.CenterKey AS CostCenter
FROM PayrollRuns AS r INNER JOIN qryPayrollCenterTotals AS t ON r.PayrollRunID = t.PayrollRunID
WHERE r.Status = 'POSTED' AND t.SumNet <> 0
UNION ALL
SELECT 'PAYROLL_PAYMENT' AS SourceType, r.PayrollRunID AS SourceID, r.RunNumber AS SourceNumber, r.PaidDate AS SourceDate, 'صرف الرواتب' AS Party, 1 AS LineOrder, 2310 AS AccountCode, r.PaidAmount AS Debit, CCur(0) AS Credit, 'صافي الرواتب' AS LineText, 0 AS CostCenter
FROM PayrollRuns AS r
WHERE r.Status = 'POSTED' AND r.PaidAmount <> 0
UNION ALL
SELECT 'PAYROLL_PAYMENT' AS SourceType, r.PayrollRunID AS SourceID, r.RunNumber AS SourceNumber, r.PaidDate AS SourceDate, 'صرف الرواتب' AS Party, 2 AS LineOrder, IIf(r.PaidFrom = 'BANK', 120000 + r.BankID, 110000 + r.CashBoxID) AS AccountCode, CCur(0) AS Debit, r.PaidAmount AS Credit, 'صرف الرواتب' AS LineText, 0 AS CostCenter
FROM PayrollRuns AS r
WHERE r.Status = 'POSTED' AND r.PaidAmount <> 0
```

## qryCommissionCenterTotals

عمولات كل مسير لكل مركز تكلفة لقيده

```sql
SELECT CommissionRunID, IIf(CostCenterID Is Null, 0, CostCenterID) AS CenterKey, Sum(Commission) AS SumCommission
FROM CommissionLines
GROUP BY CommissionRunID, IIf(CostCenterID Is Null, 0, CostCenterID)
```

## qryJournalCommission

أسطر قيود مسيرات العمولات المرحَّلة

```sql
SELECT 'COMMISSION' AS SourceType, r.CommissionRunID AS SourceID, r.RunNumber AS SourceNumber, r.RunMonth AS SourceDate, 'عمولات المندوبين' AS Party, 1 + 10 * t.CenterKey AS LineOrder, 5530 AS AccountCode, t.SumCommission AS Debit, CCur(0) AS Credit, 'عمولات المندوبين' AS LineText, t.CenterKey AS CostCenter
FROM CommissionRuns AS r INNER JOIN qryCommissionCenterTotals AS t ON r.CommissionRunID = t.CommissionRunID
WHERE r.Status = 'POSTED' AND t.SumCommission <> 0
UNION ALL
SELECT 'COMMISSION' AS SourceType, r.CommissionRunID AS SourceID, r.RunNumber AS SourceNumber, r.RunMonth AS SourceDate, 'عمولات المندوبين' AS Party, 2 + 10 * t.CenterKey AS LineOrder, 2330 AS AccountCode, CCur(0) AS Debit, t.SumCommission AS Credit, 'عمولات مستحقة' AS LineText, t.CenterKey AS CostCenter
FROM CommissionRuns AS r INNER JOIN qryCommissionCenterTotals AS t ON r.CommissionRunID = t.CommissionRunID
WHERE r.Status = 'POSTED' AND t.SumCommission <> 0
```

## qryJournalBankTx

أسطر قيود الحركات البنكية: الإيداع والسحب وتسوية مدى والتحويل والحركات الأخرى

```sql
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 1 AS LineOrder, 120000 + t.BankID AS AccountCode, t.Amount AS Debit, CCur(0) AS Credit, 'إيداع نقدية' AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.TxType = 'DEPOSIT'
UNION ALL
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 2 AS LineOrder, 110000 + t.CashBoxID AS AccountCode, CCur(0) AS Debit, t.Amount AS Credit, 'إيداع في البنك' AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.TxType = 'DEPOSIT'
UNION ALL
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 1 AS LineOrder, 110000 + t.CashBoxID AS AccountCode, t.Amount AS Debit, CCur(0) AS Credit, 'سحب من البنك' AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.TxType = 'WITHDRAW'
UNION ALL
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 2 AS LineOrder, 120000 + t.BankID AS AccountCode, CCur(0) AS Debit, t.Amount AS Credit, 'سحب نقدية' AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.TxType = 'WITHDRAW'
UNION ALL
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 1 AS LineOrder, 120000 + t.BankID AS AccountCode, t.Amount - t.FeeAmount - t.FeeVAT AS Debit, CCur(0) AS Credit, 'صافي تسوية مدى' AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.TxType = 'SETTLEMENT' AND t.Amount - t.FeeAmount - t.FeeVAT <> 0
UNION ALL
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 2 AS LineOrder, 5610 AS AccountCode, t.FeeAmount AS Debit, CCur(0) AS Credit, 'عمولة مدى' AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.TxType = 'SETTLEMENT' AND t.FeeAmount <> 0
UNION ALL
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 3 AS LineOrder, 1500 AS AccountCode, t.FeeVAT AS Debit, CCur(0) AS Credit, 'ضريبة العمولة' AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.FeeVAT <> 0 AND (t.TxType = 'SETTLEMENT' OR t.TxType = 'OTHER_OUT')
UNION ALL
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 4 AS LineOrder, 1200 AS AccountCode, CCur(0) AS Debit, t.Amount AS Credit, 'تحصيلات مدى' AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.TxType = 'SETTLEMENT'
UNION ALL
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 1 AS LineOrder, 120000 + t.ToBankID AS AccountCode, t.Amount AS Debit, CCur(0) AS Credit, 'تحويل وارد' AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.TxType = 'TRANSFER'
UNION ALL
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 2 AS LineOrder, 120000 + t.BankID AS AccountCode, CCur(0) AS Debit, t.Amount AS Credit, 'تحويل صادر' AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.TxType = 'TRANSFER'
UNION ALL
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 1 AS LineOrder, 120000 + t.BankID AS AccountCode, t.Amount AS Debit, CCur(0) AS Credit, t.Reference AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.TxType = 'OTHER_IN'
UNION ALL
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 2 AS LineOrder, t.CounterAccount AS AccountCode, CCur(0) AS Debit, t.Amount AS Credit, t.Description AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.TxType = 'OTHER_IN'
UNION ALL
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 1 AS LineOrder, t.CounterAccount AS AccountCode, t.Amount - t.FeeVAT AS Debit, CCur(0) AS Credit, t.Description AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.TxType = 'OTHER_OUT'
UNION ALL
SELECT 'BANK_TX' AS SourceType, t.BankTxID AS SourceID, t.TxNumber AS SourceNumber, t.TxDate AS SourceDate, t.Description AS Party, 2 AS LineOrder, 120000 + t.BankID AS AccountCode, CCur(0) AS Debit, t.Amount AS Credit, t.Reference AS LineText, 0 AS CostCenter
FROM BankTransactions AS t
WHERE t.TxType = 'OTHER_OUT'
```

## qryBankItemSums

صافي كل عملية على حساب كل بنك في القيود

```sql
SELECT l.AccountCode - 120000 AS BankID, e.SourceType, e.SourceID, Max(e.EntryDate) AS ItemDate,
       Max(e.SourceNumber) AS ItemNumber, Max(e.Description) AS ItemText, Sum(l.Debit) - Sum(l.Credit) AS ItemAmount
FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID
WHERE l.AccountCode > 120000 AND l.AccountCode < 130000
GROUP BY l.AccountCode, e.SourceType, e.SourceID
```

## qryBankItems

عمليات البنوك: المبلغ، وهل طابقت كشف البنك ومبلغها يوم المطابقة

```sql
SELECT i.BankID, i.SourceType, i.SourceID, i.ItemDate, i.ItemNumber, i.ItemText, i.ItemAmount,
       t.TypeName, c.ReconciliationID, c.ClearedAmount, IIf(c.ClearingID Is Null, 0, 1) AS IsCleared
FROM (qryBankItemSums AS i INNER JOIN [@JournalSourceTypes] AS t ON i.SourceType = t.SourceType)
     LEFT JOIN BankClearings AS c ON (i.BankID = c.BankID AND i.SourceType = c.SourceType AND i.SourceID = c.SourceID)
WHERE i.ItemAmount <> 0
```

## qryBankTotals

رصيد كل بنك في الدفاتر

```sql
SELECT BankID, Sum(ItemAmount) AS BookBalance, Max(ItemDate) AS LastItemDate
FROM qryBankItemSums
GROUP BY BankID
```

## BankBalanceQuery

أرصدة البنوك في الدفاتر

```sql
SELECT k.BankID, k.BankName, k.AccountNo, k.IBAN, k.IsActive, CCur(Nz(t.BookBalance, 0)) AS Balance, t.LastItemDate
FROM [@Banks] AS k LEFT JOIN qryBankTotals AS t ON k.BankID = t.BankID
ORDER BY k.BankName
```

## qryAssetDepTotals

مجموع إهلاك كل أصل في القيود الشهرية

```sql
SELECT AssetID, Sum(Amount) AS SumDep, Count(*) AS DepCount
FROM AssetDepreciations
GROUP BY AssetID
```

## FixedAssetsQuery

سجل الأصول الثابتة: التكلفة ومجمع الإهلاك والقيمة الدفترية والقسط الشهري

```sql
SELECT a.AssetID, a.AssetCode, a.AssetName, a.AssetAccount, c.AccountName AS AssetGroup, a.PurchaseDate, a.Cost,
       a.SalvageValue, a.UsefulLifeMonths, a.DepStartDate, a.Status, IIf(a.Status = 'ACTIVE', 'قائم', 'مستبعد') AS StatusName,
       a.OpeningAccumDep + CCur(Nz(t.SumDep, 0)) AS AccumDep, a.Cost - a.OpeningAccumDep - CCur(Nz(t.SumDep, 0)) AS BookValue,
       Round((a.Cost - a.SalvageValue) / a.UsefulLifeMonths, 2) AS MonthlyDep, CCur(Nz(t.DepCount, 0)) AS DepMonths,
       a.DisposalDate, a.DisposalProceeds
FROM (FixedAssets AS a INNER JOIN [@Accounts] AS c ON a.AssetAccount = c.AccountCode)
     LEFT JOIN qryAssetDepTotals AS t ON a.AssetID = t.AssetID
```

## AuditTrailQuery

سجل التدقيق: كل عملية بمن قام بها ووقتها، وحقولها بالقيمة قبل وبعد

```sql
SELECT a.LogID, a.LogDate, a.EmployeeID, IIf(e.EmployeeName Is Null, '-', e.EmployeeName) AS UserName, a.ActionType,
       IIf(a.ActionType = 'ADD', 'إضافة', IIf(a.ActionType = 'EDIT', 'تعديل', IIf(a.ActionType = 'DELETE', 'حذف',
       IIf(a.ActionType = 'LOGIN', 'دخول', IIf(a.ActionType = 'LOGOUT', 'خروج', a.ActionType))))) AS ActionLabel,
       a.ObjectName, a.RecordID, a.RecordLabel, a.ComputerName, a.Details,
       c.LineNo, c.FieldCaption, c.OldValue, c.NewValue
FROM (AuditLog AS a LEFT JOIN Employees AS e ON a.EmployeeID = e.EmployeeID)
     LEFT JOIN AuditChanges AS c ON a.LogID = c.LogID
```

## qryAdvanceMoves

حركات سلف الموظفين: الصرف والسداد النقدي والخصم من الرواتب

```sql
SELECT v.AdvanceEmployeeID AS EmployeeID, v.VoucherDate AS MoveDate, IIf(v.VoucherType = 'OUT', v.Amount, -v.Amount) AS MoveAmount
FROM CashVouchers AS v
WHERE v.Category = 'ADVANCE' AND v.AdvanceEmployeeID Is Not Null
UNION ALL
SELECT l.EmployeeID, r.PayMonth, -l.AdvanceDeduction
FROM PayrollLines AS l INNER JOIN PayrollRuns AS r ON l.PayrollRunID = r.PayrollRunID
WHERE r.Status = 'POSTED' AND l.AdvanceDeduction <> 0
```

## qryAdvanceTotals

رصيد سلف كل موظف

```sql
SELECT EmployeeID, Sum(MoveAmount) AS AdvanceBalance
FROM qryAdvanceMoves
GROUP BY EmployeeID
```

## AdvanceBalanceQuery

أرصدة سلف الموظفين

```sql
SELECT e.EmployeeID, e.EmployeeName, e.AdvanceInstallment, CCur(Nz(t.AdvanceBalance, 0)) AS Balance
FROM Employees AS e LEFT JOIN qryAdvanceTotals AS t ON e.EmployeeID = t.EmployeeID
WHERE t.AdvanceBalance <> 0
```

## PayrollSheetQuery

مسير الرواتب المختار بأسطر الموظفين

المعاملات: `PayrollRunID`

```sql
SELECT r.PayrollRunID, r.RunNumber, r.PayMonth, r.Status, l.EmployeeName, l.Basic, l.Housing, l.OtherAllow, l.Overtime,
       l.Additions, l.Basic + l.Housing + l.OtherAllow + l.Overtime + l.Additions AS Gross, l.AbsenceDeduction,
       l.AdvanceDeduction, l.OtherDeduction, l.GosiEmployee, l.GosiEmployer, l.NetPay
FROM PayrollRuns AS r INNER JOIN PayrollLines AS l ON r.PayrollRunID = l.PayrollRunID
WHERE r.PayrollRunID = QLong('PayrollRunID')
```

## ChequesQuery

الشيكات الواردة والصادرة مع العميل أو المورد وحالتها

```sql
SELECT q.ChequeID, q.ChequeRef, q.Direction, IIf(q.Direction = 'IN', 'وارد', 'صادر') AS DirectionName,
       IIf(q.Direction = 'IN', c.CustomerName, s.SupplierName) AS PartyName, q.ChequeNo, q.DrawerBank,
       k.BankName, q.IssueDate, q.DueDate, q.Amount, q.Status,
       IIf(q.Status = 'PENDING', 'تحت التحصيل', IIf(q.Status = 'COLLECTED', IIf(q.Direction = 'IN', 'محصَّل', 'مصروف'),
           'مرتد')) AS StatusName, q.StatusDate, q.Notes
FROM ((Cheques AS q LEFT JOIN [@Customers] AS c ON q.CustomerID = c.CustomerID)
      LEFT JOIN [@Suppliers] AS s ON q.SupplierID = s.SupplierID)
     LEFT JOIN [@Banks] AS k ON q.BankID = k.BankID
```

## JournalLinesQuery

قيود اليومية خلال فترة بأسطرها

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT e.EntryID, e.EntryNumber, e.EntryDate, e.SourceType, t.TypeName, e.SourceID, e.SourceNumber,
       e.Description, l.LineNumber, l.AccountCode, a.AccountName, l.LineText, l.Debit, l.Credit
FROM ((JournalEntries AS e INNER JOIN [@JournalSourceTypes] AS t ON e.SourceType = t.SourceType)
      INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID)
     INNER JOIN [@Accounts] AS a ON l.AccountCode = a.AccountCode
WHERE e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd')
ORDER BY e.EntryDate, e.EntryNumber, l.LineNumber
```

## qryJournalEntryPrint

بيانات طباعة قيد

```sql
SELECT e.EntryID, e.EntryNumber, e.EntryDate, t.TypeName, e.SourceNumber, e.Description, e.TotalDebit,
       l.LineNumber, l.AccountCode, a.AccountName, l.LineText, l.Debit, l.Credit
FROM ((JournalEntries AS e INNER JOIN [@JournalSourceTypes] AS t ON e.SourceType = t.SourceType)
      INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID)
     INNER JOIN [@Accounts] AS a ON l.AccountCode = a.AccountCode
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
FROM ([@Accounts] AS a LEFT JOIN qryTrialBefore AS b ON a.AccountCode = b.AccountCode)
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
      INNER JOIN [@Accounts] AS a ON l.AccountCode = a.AccountCode)
     INNER JOIN [@JournalSourceTypes] AS k ON e.SourceType = k.SourceType), [@Accounts] AS s
WHERE s.AccountCode = QLong('AccountCode') AND (a.Level1Code = QLong('AccountCode') OR a.Level2Code = QLong('AccountCode') OR a.Level3Code = QLong('AccountCode') OR a.Level4Code = QLong('AccountCode') OR a.Level5Code = QLong('AccountCode')) AND e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd')
UNION ALL
SELECT 0, s.AccountCode, s.AccountName, QDate('PeriodStart'), '-', 0, 'رصيد أول المدة', Null, Null, Null, Null,
       IIf(x.SumBefore > 0, x.SumBefore, 0), IIf(x.SumBefore < 0, -x.SumBefore, 0)
FROM [@Accounts] AS s, qryStatementBefore AS x
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
      INNER JOIN [@Accounts] AS a ON l.AccountCode = a.AccountCode)
     INNER JOIN [@JournalSourceTypes] AS k ON e.SourceType = k.SourceType
WHERE (QLong('AccountCode') = 0 OR (a.Level1Code = QLong('AccountCode') OR a.Level2Code = QLong('AccountCode') OR a.Level3Code = QLong('AccountCode') OR a.Level4Code = QLong('AccountCode') OR a.Level5Code = QLong('AccountCode'))) AND e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd')
UNION ALL
SELECT 0, b.TreeKey, b.AccountCode, b.AccountName, QDate('PeriodStart'), '-', 0, 'رصيد أول المدة', Null, Null,
       IIf(t.OpeningBalance > 0, t.OpeningBalance, 0), IIf(t.OpeningBalance < 0, -t.OpeningBalance, 0)
FROM TrialBalanceQuery AS t INNER JOIN [@Accounts] AS b ON t.AccountCode = b.AccountCode
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
FROM [@Accounts] AS a INNER JOIN qryTreeRollup AS r ON a.AccountCode = r.TreeCode
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
FROM ([@Accounts] AS a LEFT JOIN qryIncomeMoves AS c ON a.AccountCode = c.AccountCode)
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

## qryCenterMoves

صافي حركة كل حساب إيرادات أو مصروفات لكل مركز تكلفة في الفترة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT IIf(l.CostCenterID Is Null, 0, l.CostCenterID) AS CenterKey, l.AccountCode, a.AccountName, a.TreeKey,
       a.Level2Code, IIf(a.AccountType = 'REVENUE', 1, -1) * (Sum(l.Credit) - Sum(l.Debit)) AS CenterAmount
FROM (JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID)
     INNER JOIN [@Accounts] AS a ON l.AccountCode = a.AccountCode
WHERE e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd') AND e.SourceType <> 'YEAR_CLOSE' AND a.AccountType IN ('REVENUE', 'EXPENSE')
GROUP BY IIf(l.CostCenterID Is Null, 0, l.CostCenterID), l.AccountCode, a.AccountName, a.TreeKey, a.Level2Code,
         a.AccountType
```

## qryCenterNames

مراكز التكلفة ومعها «غير موزع»

```sql
SELECT CostCenterID AS CenterKey, CenterCode, CenterName
FROM [@CostCenters]
UNION ALL
SELECT 0, '-', 'غير موزع'
FROM Settings AS z
WHERE z.SettingID = 1
```

## qryCenterSums

الإيرادات وتكلفة المبيعات والمصروفات لكل مركز تكلفة

```sql
SELECT CenterKey, Sum(IIf(Level2Code = 41 Or Level2Code = 42, CenterAmount, 0)) AS SumRevenue,
       Sum(IIf(Level2Code = 51, CenterAmount, 0)) AS SumCostOfSales,
       Sum(IIf(Level2Code = 52 Or Level2Code = 53, CenterAmount, 0)) AS SumExpenses
FROM qryCenterMoves
GROUP BY CenterKey
```

## CostCenterProfitQuery

قائمة الدخل لكل مركز تكلفة: الإيرادات، تكلفة المبيعات، مجمل الربح، المصروفات، صافي الربح

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT n.CenterKey, n.CenterCode, n.CenterName, s.SumRevenue AS Revenue, s.SumCostOfSales AS CostOfSales,
       s.SumRevenue - s.SumCostOfSales AS GrossProfit, s.SumExpenses AS Expenses,
       s.SumRevenue - s.SumCostOfSales - s.SumExpenses AS NetProfit
FROM qryCenterNames AS n INNER JOIN qryCenterSums AS s ON n.CenterKey = s.CenterKey
```

## CostCenterAccountsQuery

إيرادات ومصروفات كل مركز تكلفة بالحسابات

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT n.CenterKey, n.CenterName, m.AccountCode, m.AccountName, m.TreeKey,
       IIf(m.Level2Code = 41 Or m.Level2Code = 42, 'إيرادات', IIf(m.Level2Code = 51, 'تكلفة المبيعات', 'مصروفات'))
           AS SectionName, m.CenterAmount
FROM qryCenterNames AS n INNER JOIN qryCenterMoves AS m ON n.CenterKey = m.CenterKey
```

## qryBudgetMonths

أشهر الموازنة: سطر لكل شهر من كل سطر موازنة

```sql
SELECT l.BudgetLineID, h.BudgetYear, 1 AS MonthNo, h.BudgetYear * 100 + 1 AS MonthKey, l.M1 AS PlanAmount
FROM BudgetLines AS l INNER JOIN Budgets AS h ON l.BudgetID = h.BudgetID
UNION ALL
SELECT l.BudgetLineID, h.BudgetYear, 2, h.BudgetYear * 100 + 2, l.M2
FROM BudgetLines AS l INNER JOIN Budgets AS h ON l.BudgetID = h.BudgetID
UNION ALL
SELECT l.BudgetLineID, h.BudgetYear, 3, h.BudgetYear * 100 + 3, l.M3
FROM BudgetLines AS l INNER JOIN Budgets AS h ON l.BudgetID = h.BudgetID
UNION ALL
SELECT l.BudgetLineID, h.BudgetYear, 4, h.BudgetYear * 100 + 4, l.M4
FROM BudgetLines AS l INNER JOIN Budgets AS h ON l.BudgetID = h.BudgetID
UNION ALL
SELECT l.BudgetLineID, h.BudgetYear, 5, h.BudgetYear * 100 + 5, l.M5
FROM BudgetLines AS l INNER JOIN Budgets AS h ON l.BudgetID = h.BudgetID
UNION ALL
SELECT l.BudgetLineID, h.BudgetYear, 6, h.BudgetYear * 100 + 6, l.M6
FROM BudgetLines AS l INNER JOIN Budgets AS h ON l.BudgetID = h.BudgetID
UNION ALL
SELECT l.BudgetLineID, h.BudgetYear, 7, h.BudgetYear * 100 + 7, l.M7
FROM BudgetLines AS l INNER JOIN Budgets AS h ON l.BudgetID = h.BudgetID
UNION ALL
SELECT l.BudgetLineID, h.BudgetYear, 8, h.BudgetYear * 100 + 8, l.M8
FROM BudgetLines AS l INNER JOIN Budgets AS h ON l.BudgetID = h.BudgetID
UNION ALL
SELECT l.BudgetLineID, h.BudgetYear, 9, h.BudgetYear * 100 + 9, l.M9
FROM BudgetLines AS l INNER JOIN Budgets AS h ON l.BudgetID = h.BudgetID
UNION ALL
SELECT l.BudgetLineID, h.BudgetYear, 10, h.BudgetYear * 100 + 10, l.M10
FROM BudgetLines AS l INNER JOIN Budgets AS h ON l.BudgetID = h.BudgetID
UNION ALL
SELECT l.BudgetLineID, h.BudgetYear, 11, h.BudgetYear * 100 + 11, l.M11
FROM BudgetLines AS l INNER JOIN Budgets AS h ON l.BudgetID = h.BudgetID
UNION ALL
SELECT l.BudgetLineID, h.BudgetYear, 12, h.BudgetYear * 100 + 12, l.M12
FROM BudgetLines AS l INNER JOIN Budgets AS h ON l.BudgetID = h.BudgetID
```

## qryBudgetPlanned

مبلغ الموازنة لكل سطر في أشهر الفترة

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT BudgetLineID, Sum(PlanAmount) AS SumPlan
FROM qryBudgetMonths
WHERE BudgetYear = Year(QDate('PeriodStart')) AND MonthKey >= Year(QDate('PeriodStart')) * 100 + Month(QDate('PeriodStart'))
      AND MonthKey <= Year(DateAdd('d', -1, QDate('PeriodEnd'))) * 100 + Month(DateAdd('d', -1, QDate('PeriodEnd')))
GROUP BY BudgetLineID
```

## qryBudgetActual

الفعلي لكل سطر موازنة في الفترة من القيود

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT b.BudgetLineID, Sum(IIf(a.AccountType = 'REVENUE', l.Credit - l.Debit, l.Debit - l.Credit)) AS SumActual
FROM Budgets AS h, BudgetLines AS b, Accounts AS a, JournalEntries AS e, JournalLines AS l, Accounts AS d
WHERE h.BudgetYear = Year(QDate('PeriodStart')) AND b.BudgetID = h.BudgetID AND a.AccountCode = b.AccountCode
      AND l.EntryID = e.EntryID AND d.AccountCode = l.AccountCode AND (d.Level1Code = b.AccountCode OR d.Level2Code = b.AccountCode OR d.Level3Code = b.AccountCode OR d.Level4Code = b.AccountCode OR d.Level5Code = b.AccountCode)
      AND (b.CostCenterID Is Null OR l.CostCenterID = b.CostCenterID)
      AND e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd') AND e.SourceType <> 'YEAR_CLOSE'
GROUP BY b.BudgetLineID
```

## BudgetVsActualQuery

الموازنة مقابل الفعلي في الفترة: الانحراف ونسبته، وهل هو ملائم

المعاملات: `PeriodStart`, `PeriodEnd`

```sql
SELECT b.BudgetLineID, b.AccountCode, a.AccountName, a.TreeKey, a.AccountType,
       IIf(a.AccountType = 'REVENUE', 'الإيرادات', 'المصروفات') AS SectionName,
       IIf(c.CenterName Is Null, 'كل المراكز', c.CenterName) AS BudgetCenter,
       CCur(Nz(p.SumPlan, 0)) AS BudgetAmount, CCur(Nz(x.SumActual, 0)) AS ActualAmount,
       CCur(Nz(x.SumActual, 0)) - CCur(Nz(p.SumPlan, 0)) AS Variance,
       IIf(CCur(Nz(p.SumPlan, 0)) = 0, Null, (CCur(Nz(x.SumActual, 0)) - CCur(Nz(p.SumPlan, 0))) / CCur(Nz(p.SumPlan, 0))) AS VariancePct,
       IIf(CCur(Nz(x.SumActual, 0)) = CCur(Nz(p.SumPlan, 0)), 'مطابق', IIf((a.AccountType = 'REVENUE') = (CCur(Nz(x.SumActual, 0)) > CCur(Nz(p.SumPlan, 0))),
           'ملائم', 'غير ملائم')) AS VarianceNote
FROM ((((Budgets AS h INNER JOIN BudgetLines AS b ON h.BudgetID = b.BudgetID)
       INNER JOIN [@Accounts] AS a ON b.AccountCode = a.AccountCode)
      LEFT JOIN [@CostCenters] AS c ON b.CostCenterID = c.CostCenterID)
     LEFT JOIN qryBudgetPlanned AS p ON b.BudgetLineID = p.BudgetLineID)
     LEFT JOIN qryBudgetActual AS x ON b.BudgetLineID = x.BudgetLineID
WHERE h.BudgetYear = Year(QDate('PeriodStart'))
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
FROM ([@Accounts] AS a LEFT JOIN qryBalanceAt AS b ON a.AccountCode = b.AccountCode)
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
FROM qryBalanceAccounts AS q INNER JOIN [@Accounts] AS g ON q.Level2Code = g.AccountCode
UNION ALL
SELECT 3, g.TreeKey, 1, 'Z', 'A', 'صافي ربح (خسارة) الفترات غير المقفلة', Null, -x.NetProfitSum, -y.NetCompareSum
FROM [@Accounts] AS g, qryProfitAt AS x, qryProfitCompare AS y
WHERE g.AccountCode = 32
UNION ALL
SELECT c.AccountCode, '', 0, '', 'C', c.AccountName, Null, Null, Null
FROM [@Accounts] AS c
WHERE c.AccountCode IN (1, 2, 3)
UNION ALL
SELECT g.Level1Code, g.TreeKey, 0, '', 'G', g.AccountName, Null, Null, Null
FROM [@Accounts] AS g
WHERE g.AccountCode IN (SELECT GroupCode FROM qryBalanceItems)
UNION ALL
SELECT g.Level1Code, g.TreeKey, 2, '', 'S', g.AccountName, Null, Sum(i.CurrentValue), Sum(i.PriorValue)
FROM [@Accounts] AS g INNER JOIN qryBalanceItems AS i ON g.AccountCode = i.GroupCode
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
FROM [@Accounts] AS a
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
