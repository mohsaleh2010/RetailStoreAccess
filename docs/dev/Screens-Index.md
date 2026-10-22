# فهرس الشاشات (Screens Index)

> ملف مُولَّد تلقائيًا من `tools/forms*.py` بواسطة `tools/generate.py` – لا تعدّله يدويًا.

عدد الشاشات: **113**. شرح طريقة بناء الشاشات في [03-Forms.md](03-Forms.md).

| # | الشاشة | العنوان | النوع | الجدول | الصلاحية | تُفتح من |
|---|---|---|---|---|---|---|
| 1 | `frmMain` | نظام إدارة المحل | الشاشة الرئيسية | - | الكل | - |
| 2 | `frmProducts` | المنتجات | نافذة منبثقة | Products | PRODUCTS | frmInventory، frmMain، frmPurchaseInvoice، القائمة الجانبية |
| 3 | `frmCustomers` | العملاء | نافذة منبثقة | Customers | CUSTOMERS | frmMain، القائمة الجانبية |
| 4 | `frmSuppliers` | الموردون | نافذة منبثقة | Suppliers | SUPPLIERS | frmMain، القائمة الجانبية |
| 5 | `frmExpenses` | المصروفات | نافذة منبثقة | Expenses | EXPENSES | frmMain، frmRecurring، القائمة الجانبية |
| 6 | `frmCurrencies` | العملات | نافذة منبثقة | Currencies | CURRENCIES | frmAccounting |
| 7 | `frmCurrencyRates` | أسعار العملات | نافذة منبثقة | CurrencyRates | CURRENCIES | frmCurrencies |
| 8 | `frmSalesReps` | المندوبين | نافذة منبثقة | SalesReps | SALES_REPS | frmAccounting، frmCommissions |
| 9 | `frmRepTargets` | أهداف المندوبين | نافذة منبثقة | SalesRepTargets | SALES_REPS | frmSalesReps |
| 10 | `frmRouteStops` | خط سير المندوبين | نافذة منبثقة | RouteStops | REP_ROUTES | frmRepVisits |
| 11 | `frmRepPrices` | أسعار المندوبين | نافذة منبثقة | RepPrices | REP_ROUTES | frmRepVisits |
| 12 | `frmRecurring` | المصروفات المتكررة | نافذة منبثقة | RecurringExpenses | EXPENSES | frmAccounting |
| 13 | `frmUsers` | المستخدمون | نافذة منبثقة | Employees | USERS | frmMain، القائمة الجانبية |
| 14 | `frmCostCenters` | مراكز التكلفة | نافذة منبثقة | CostCenters | JOURNAL | frmAccounting، frmAccounts |
| 15 | `frmEmployeePay` | رواتب الموظفين | نافذة منبثقة | Employees | PAYROLL | frmPayroll |
| 16 | `frmCategories` | التصنيفات | نافذة منبثقة | Categories | PRODUCTS | frmSettings |
| 17 | `frmUnits` | وحدات القياس | نافذة منبثقة | Units | PRODUCTS | frmSettings |
| 18 | `frmProductUnits` | وحدات الأصناف | نافذة منبثقة | ProductUnits | PRODUCTS | - |
| 19 | `frmPriceLists` | قوائم الأسعار | نافذة منبثقة | PriceLists | PRODUCTS | frmProducts |
| 20 | `frmPriceListItems` | أسعار قوائم الأسعار | نافذة منبثقة | PriceListItems | PRODUCTS | - |
| 21 | `frmExpenseTypes` | أنواع المصروفات | نافذة منبثقة | ExpenseTypes | EXPENSES | frmExpenses، frmSettings |
| 22 | `frmCashBoxes` | الصناديق | نافذة منبثقة | CashBoxes | CASH_BOX | frmCashBoxUsers، frmTreasury |
| 23 | `frmCashBoxUsers` | أمناء الصناديق | نافذة منبثقة | CashBoxUsers | CASH_ALL_BOXES | frmCashBoxes |
| 24 | `frmWarehouses` | المخازن | نافذة منبثقة | Warehouses | PRODUCTS | frmInventory |
| 25 | `frmStockParties` | جهات الصرف والاستلام المخزني | نافذة منبثقة | StockParties | STOCK_DOCS | frmStockDocument |
| 26 | `frmBanks` | البنوك | نافذة منبثقة | Banks | BANKS | frmAccounting، frmTreasury |
| 27 | `frmAccounts` | دليل الحسابات | نافذة منبثقة | Accounts | JOURNAL | frmAccounting، frmJournal، frmLedger، frmManualEntry |
| 28 | `frmSettings` | الإعدادات | نافذة منبثقة | Settings | SETTINGS | frmMain، القائمة الجانبية |
| 29 | `frmLabelSettings` | إعدادات ملصقات الباركود | نافذة منبثقة | LabelSettings | PRODUCTS | frmBarcodeLabels، frmSettings |
| 30 | `frmSearch` | البحث المتقدم | شاشة كاملة | - | الكل | frmMain، القائمة الجانبية |
| 31 | `frmReportCenter` | التقارير | شاشة كاملة | - | REPORTS | frmMain، frmProducts، frmRepVisits، frmSalesReps، frmWarehouses، القائمة الجانبية |
| 32 | `frmPOSLines` | أسطر الفاتورة | شاشة فرعية في frmPOS | tmpPOSLines | مع الشاشة الأم | - |
| 33 | `frmPOS` | نقطة البيع | شاشة كاملة | - | SALES_POS | frmMain، القائمة الجانبية |
| 34 | `frmReturnLines` | أسطر المرتجع | شاشة فرعية في frmSalesReturn | tmpReturnLines | مع الشاشة الأم | - |
| 35 | `frmSalesReturn` | مرتجع مبيعات | نافذة منبثقة | - | SALES_RETURN | frmPOS، frmSalesInvoice |
| 36 | `frmCustomerPayment` | سند قبض | نافذة منبثقة | - | CUSTOMER_PAYMENTS | frmCustomers، frmPOS |
| 37 | `frmSalesInvoice` | فاتورة بيع | نافذة منبثقة | - | SALES_VIEW | - |
| 38 | `frmPurchaseLines` | أسطر فاتورة الشراء | شاشة فرعية في frmPurchaseInvoice | tmpPurchaseLines | مع الشاشة الأم | - |
| 39 | `frmPurchaseInvoice` | فاتورة مشتريات | شاشة كاملة | - | PURCHASES | frmInventory، frmMain، القائمة الجانبية |
| 40 | `frmPurchaseReturnLines` | أسطر مرتجع المشتريات | شاشة فرعية في frmPurchaseReturn | tmpPurchaseReturnLines | مع الشاشة الأم | - |
| 41 | `frmPurchaseReturn` | مرتجع مشتريات | نافذة منبثقة | - | PURCHASE_RETURN | frmPurchaseInvoice، frmPurchaseView |
| 42 | `frmSupplierPayment` | سند صرف | نافذة منبثقة | - | SUPPLIER_PAYMENTS | frmPurchaseInvoice، frmPurchaseView، frmSuppliers |
| 43 | `frmPurchaseView` | فاتورة شراء | نافذة منبثقة | - | PURCHASES | - |
| 44 | `frmInventory` | المخزون | شاشة كاملة | - | PRODUCTS | frmMain، frmWarehouses، القائمة الجانبية |
| 45 | `frmStockCountLines` | أسطر الجرد | شاشة فرعية في frmStockCount | StockCountDetails | مع الشاشة الأم | - |
| 46 | `frmStockCount` | الجرد | شاشة كاملة | - | STOCK_COUNT | frmInventory، frmMain، القائمة الجانبية |
| 47 | `frmStockTransferLines` | أسطر التحويل | شاشة فرعية في frmStockTransfer | StockTransferLines | مع الشاشة الأم | - |
| 48 | `frmStockTransfer` | التحويل بين المخازن | نافذة منبثقة | - | STOCK_TRANSFER | frmInventory، frmWarehouses |
| 49 | `frmStockDocumentLines` | أسطر السند المخزني | شاشة فرعية في frmStockDocument | StockDocumentLines | مع الشاشة الأم | - |
| 50 | `frmStockDocument` | سندات الصرف والاستلام المخزني | نافذة منبثقة | - | STOCK_DOCS | frmInventory، frmStockParties، frmWarehouses |
| 51 | `frmStockNoteLines` | أسطر إذن التسليم أو الاستلام | شاشة فرعية في frmStockNote | StockNoteLines | مع الشاشة الأم | - |
| 52 | `frmStockNote` | أذونات التسليم والاستلام | نافذة منبثقة | - | STOCK_NOTES | frmInventory |
| 53 | `frmLogin` | تسجيل الدخول | نافذة منبثقة | - | الكل | - |
| 54 | `frmChangePassword` | كلمة المرور | نافذة منبثقة | - | الكل | frmMain، frmUsers |
| 55 | `frmRolePermLines` | صلاحيات الدور | شاشة فرعية في frmRoles | tmpRolePermissions | مع الشاشة الأم | - |
| 56 | `frmRoles` | الأدوار والصلاحيات | نافذة منبثقة | - | USERS | frmUsers |
| 57 | `frmUserScreenLines` | شاشات المستخدم | شاشة فرعية في frmUserScreens | tmpUserScreens | مع الشاشة الأم | - |
| 58 | `frmUserScreens` | صلاحيات الشاشات | نافذة منبثقة | - | USERS | frmUsers |
| 59 | `frmActivation` | تفعيل البرنامج | نافذة منبثقة | - | الكل | frmSettings |
| 60 | `frmBackup` | النسخ الاحتياطي | نافذة منبثقة | - | BACKUP | frmMain، القائمة الجانبية |
| 61 | `frmLabelLines` | أسطر الملصقات | شاشة فرعية في frmBarcodeLabels | - | مع الشاشة الأم | - |
| 62 | `frmBarcodeLabels` | طباعة ملصقات الباركود | شاشة كاملة | - | PRODUCTS | frmInventory |
| 63 | `frmTouchLines` | أسطر الطلب | شاشة فرعية في frmCafePOS | tmpPOSLines | مع الشاشة الأم | - |
| 64 | `frmTouchPOS` | نقطة بيع المطعم | شاشة كاملة | - | SALES_POS | - |
| 65 | `frmTouchPay` | الدفع نقدًا | نافذة منبثقة | - | SALES_POS | - |
| 66 | `frmCafePOS` | نقطة بيع الكافيه | شاشة كاملة | - | SALES_POS | - |
| 67 | `frmCafeItem` | خيارات المشروب | نافذة منبثقة | - | SALES_POS | - |
| 68 | `frmTreasury` | الخزينة | شاشة كاملة | - | CASH_CLOSING | frmCashBoxes، frmExpenses، frmMain، القائمة الجانبية |
| 69 | `frmCashVoucher` | سند نقدية | نافذة منبثقة | - | CASH_BOX | frmCommissions |
| 70 | `frmCashClosing` | تصفية يومية الكاشير | نافذة منبثقة | - | CASH_CLOSING | - |
| 71 | `frmCashTransfers` | التحويلات بالطريق | نافذة منبثقة | - | CASH_CLOSING | frmTreasury |
| 72 | `frmJournal` | قيود اليومية | شاشة كاملة | - | JOURNAL | frmAccounting، frmAccounts، frmFinancials |
| 73 | `frmJournalEntry` | قيد يومية | نافذة منبثقة | - | JOURNAL | - |
| 74 | `frmManualLines` | أسطر القيد اليدوي | شاشة فرعية في frmManualEntry | tmpManualLines | مع الشاشة الأم | - |
| 75 | `frmManualEntry` | القيود اليدوية | شاشة كاملة | - | MANUAL_ENTRY | frmAccounting، frmAccounts، frmJournal، frmLedger |
| 76 | `frmLedger` | كشف حساب | شاشة كاملة | - | JOURNAL | frmAccounting، frmAccounts، frmBanks، frmJournal |
| 77 | `frmFinancials` | القوائم المالية | شاشة كاملة | - | REPORTS_PROFIT | frmAccounting، frmLedger |
| 78 | `frmPeriodClosing` | إقفال الفترات والسنة المالية | نافذة منبثقة | - | PERIOD_CLOSE | frmAccounting، frmFinancials، frmFiscalPeriods |
| 79 | `frmFiscalPeriods` | الفترات المالية | نافذة منبثقة | - | PERIOD_CLOSE | frmPeriodClosing |
| 80 | `frmVatReturn` | إقرار ضريبة القيمة المضافة | نافذة منبثقة | - | VAT_RETURN | frmAccounting، frmFinancials |
| 81 | `frmAging` | أعمار الديون | نافذة منبثقة | - | REPORTS | frmAccounting، frmCustomers، frmSuppliers |
| 82 | `frmAllocation` | ربط السداد بالفواتير | نافذة منبثقة | - | CUSTOMER_PAYMENTS | frmCustomers، frmSuppliers |
| 83 | `frmBankTx` | الحركات البنكية | نافذة منبثقة | - | BANKS | frmBanks |
| 84 | `frmBankRecon` | التسوية البنكية | نافذة منبثقة | - | BANKS | frmBankTx، frmBanks |
| 85 | `frmCheques` | الشيكات | نافذة منبثقة | - | CHEQUES | frmAccounting، frmBanks، frmTreasury |
| 86 | `frmAssets` | الأصول الثابتة | نافذة منبثقة | - | FIXED_ASSETS | frmAccounting، frmDepreciation، frmFinancials |
| 87 | `frmDepreciation` | الإهلاك الشهري | نافذة منبثقة | - | FIXED_ASSETS | frmAccounting، frmAssets |
| 88 | `frmAssetChanges` | إعادة تقييم الأصول ونقلها | نافذة منبثقة | - | FIXED_ASSETS | - |
| 89 | `frmPayrollLines` | أسطر مسير الرواتب | شاشة فرعية في frmPayroll | PayrollLines | مع الشاشة الأم | - |
| 90 | `frmPayroll` | مسير الرواتب | نافذة منبثقة | - | PAYROLL | frmAccounting، frmEmployeePay، frmFinancials |
| 91 | `frmBudgetLines` | أسطر الموازنة | شاشة فرعية في frmBudget | BudgetLines | مع الشاشة الأم | - |
| 92 | `frmBudget` | الموازنة التقديرية | نافذة منبثقة | - | BUDGET | frmAccounting، frmFinancials |
| 93 | `frmAccounting` | المحاسبة والمالية | شاشة كاملة | - | الكل | frmMain، القائمة الجانبية |
| 94 | `frmAuditLog` | سجل التدقيق | نافذة منبثقة | - | AUDIT_LOG | frmAccounting، frmUsers |
| 95 | `frmCommissionLines` | أسطر مسير العمولات | شاشة فرعية في frmCommissions | CommissionLines | مع الشاشة الأم | - |
| 96 | `frmCommissions` | عمولات المندوبين | نافذة منبثقة | - | SALES_REPS | frmAccounting، frmSalesReps |
| 97 | `frmEnglishNameLines` | الأسماء الإنجليزية | شاشة فرعية في frmEnglishNames | tmpEnglishNames | مع الشاشة الأم | - |
| 98 | `frmEnglishNames` | الأسماء الإنجليزية | نافذة منبثقة | - | SETTINGS | frmSettings |
| 99 | `frmEInvoices` | الفاتورة الإلكترونية | نافذة منبثقة | - | EINVOICE | frmSettings |
| 100 | `frmZatcaSetup` | إعداد ربط منصة فاتورة | نافذة منبثقة | - | SETTINGS | - |
| 101 | `frmEtaSetup` | إعداد ربط منظومة الإيصال الإلكتروني | نافذة منبثقة | - | SETTINGS | - |
| 102 | `frmOpeningAccountLines` | الأرصدة الافتتاحية للحسابات | شاشة فرعية في frmOpeningBalances | OpeningBalances | مع الشاشة الأم | - |
| 103 | `frmOpeningCustomers` | الأرصدة الافتتاحية | شاشة كاملة | Customers | OPENING_BALANCES | - |
| 104 | `frmOpeningSuppliers` | الأرصدة الافتتاحية | شاشة كاملة | Suppliers | OPENING_BALANCES | - |
| 105 | `frmOpeningBoxes` | الأرصدة الافتتاحية | شاشة كاملة | CashBoxes | OPENING_BALANCES | - |
| 106 | `frmOpeningBanks` | الأرصدة الافتتاحية | شاشة كاملة | Banks | OPENING_BALANCES | - |
| 107 | `frmOpeningStock` | الرصيد الافتتاحي للمخزون | شاشة كاملة | OpeningStock | OPENING_BALANCES | - |
| 108 | `frmOpeningBalances` | أرصدة بداية التشغيل | شاشة كاملة | - | OPENING_BALANCES | frmAccounting |
| 109 | `frmSalesOrderLines` | أسطر طلبية العميل | شاشة فرعية في frmSalesOrder | SalesOrderLines | مع الشاشة الأم | - |
| 110 | `frmSalesOrder` | طلبيات العملاء | نافذة منبثقة | - | SALES_ORDERS | frmAccounting، frmSalesReps |
| 111 | `frmRepDay` | يوم المندوب | شاشة كاملة | - | REP_DAYS | frmAccounting |
| 112 | `frmRepVisits` | زيارات المندوبين | شاشة كاملة | - | REP_ROUTES | frmAccounting |
| 113 | `frmSqlServer` | قاعدة البيانات على SQL Server | نافذة منبثقة | - | SQL_SERVER | frmBackup |
