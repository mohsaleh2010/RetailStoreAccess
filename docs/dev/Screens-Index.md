# فهرس الشاشات (Screens Index)

> ملف مُولَّد تلقائيًا من `tools/forms*.py` بواسطة `tools/generate.py` – لا تعدّله يدويًا.

عدد الشاشات: **104**. شرح طريقة بناء الشاشات في [03-Forms.md](03-Forms.md).

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
| 10 | `frmRecurring` | المصروفات المتكررة | نافذة منبثقة | RecurringExpenses | EXPENSES | frmAccounting |
| 11 | `frmUsers` | المستخدمون | نافذة منبثقة | Employees | USERS | frmMain، القائمة الجانبية |
| 12 | `frmCostCenters` | مراكز التكلفة | نافذة منبثقة | CostCenters | JOURNAL | frmAccounting، frmAccounts |
| 13 | `frmEmployeePay` | رواتب الموظفين | نافذة منبثقة | Employees | PAYROLL | frmPayroll |
| 14 | `frmCategories` | التصنيفات | نافذة منبثقة | Categories | PRODUCTS | frmSettings |
| 15 | `frmUnits` | وحدات القياس | نافذة منبثقة | Units | PRODUCTS | frmSettings |
| 16 | `frmExpenseTypes` | أنواع المصروفات | نافذة منبثقة | ExpenseTypes | EXPENSES | frmExpenses، frmSettings |
| 17 | `frmCashBoxes` | الصناديق | نافذة منبثقة | CashBoxes | CASH_BOX | frmCashBoxUsers، frmTreasury |
| 18 | `frmCashBoxUsers` | أمناء الصناديق | نافذة منبثقة | CashBoxUsers | CASH_ALL_BOXES | frmCashBoxes |
| 19 | `frmWarehouses` | المخازن | نافذة منبثقة | Warehouses | PRODUCTS | frmInventory |
| 20 | `frmStockParties` | جهات الصرف والاستلام المخزني | نافذة منبثقة | StockParties | STOCK_DOCS | frmStockDocument |
| 21 | `frmBanks` | البنوك | نافذة منبثقة | Banks | BANKS | frmAccounting، frmTreasury |
| 22 | `frmAccounts` | دليل الحسابات | نافذة منبثقة | Accounts | JOURNAL | frmAccounting، frmJournal، frmLedger، frmManualEntry |
| 23 | `frmSettings` | الإعدادات | نافذة منبثقة | Settings | SETTINGS | frmMain، القائمة الجانبية |
| 24 | `frmLabelSettings` | إعدادات ملصقات الباركود | نافذة منبثقة | LabelSettings | PRODUCTS | frmBarcodeLabels، frmSettings |
| 25 | `frmSearch` | البحث المتقدم | شاشة كاملة | - | الكل | frmMain، القائمة الجانبية |
| 26 | `frmReportCenter` | التقارير | شاشة كاملة | - | REPORTS | frmMain، frmProducts، frmSalesReps، frmWarehouses، القائمة الجانبية |
| 27 | `frmPOSLines` | أسطر الفاتورة | شاشة فرعية في frmPOS | tmpPOSLines | مع الشاشة الأم | - |
| 28 | `frmPOS` | نقطة البيع | شاشة كاملة | - | SALES_POS | frmMain، القائمة الجانبية |
| 29 | `frmReturnLines` | أسطر المرتجع | شاشة فرعية في frmSalesReturn | tmpReturnLines | مع الشاشة الأم | - |
| 30 | `frmSalesReturn` | مرتجع مبيعات | نافذة منبثقة | - | SALES_RETURN | frmPOS، frmSalesInvoice |
| 31 | `frmCustomerPayment` | سند قبض | نافذة منبثقة | - | CUSTOMER_PAYMENTS | frmCustomers، frmPOS |
| 32 | `frmSalesInvoice` | فاتورة بيع | نافذة منبثقة | - | SALES_VIEW | - |
| 33 | `frmPurchaseLines` | أسطر فاتورة الشراء | شاشة فرعية في frmPurchaseInvoice | tmpPurchaseLines | مع الشاشة الأم | - |
| 34 | `frmPurchaseInvoice` | فاتورة مشتريات | شاشة كاملة | - | PURCHASES | frmInventory، frmMain، القائمة الجانبية |
| 35 | `frmPurchaseReturnLines` | أسطر مرتجع المشتريات | شاشة فرعية في frmPurchaseReturn | tmpPurchaseReturnLines | مع الشاشة الأم | - |
| 36 | `frmPurchaseReturn` | مرتجع مشتريات | نافذة منبثقة | - | PURCHASE_RETURN | frmPurchaseInvoice، frmPurchaseView |
| 37 | `frmSupplierPayment` | سند صرف | نافذة منبثقة | - | SUPPLIER_PAYMENTS | frmPurchaseInvoice، frmPurchaseView، frmSuppliers |
| 38 | `frmPurchaseView` | فاتورة شراء | نافذة منبثقة | - | PURCHASES | - |
| 39 | `frmInventory` | المخزون | شاشة كاملة | - | PRODUCTS | frmMain، frmWarehouses، القائمة الجانبية |
| 40 | `frmStockCountLines` | أسطر الجرد | شاشة فرعية في frmStockCount | StockCountDetails | مع الشاشة الأم | - |
| 41 | `frmStockCount` | الجرد | شاشة كاملة | - | STOCK_COUNT | frmInventory، frmMain، القائمة الجانبية |
| 42 | `frmStockTransferLines` | أسطر التحويل | شاشة فرعية في frmStockTransfer | StockTransferLines | مع الشاشة الأم | - |
| 43 | `frmStockTransfer` | التحويل بين المخازن | نافذة منبثقة | - | STOCK_TRANSFER | frmInventory، frmWarehouses |
| 44 | `frmStockDocumentLines` | أسطر السند المخزني | شاشة فرعية في frmStockDocument | StockDocumentLines | مع الشاشة الأم | - |
| 45 | `frmStockDocument` | سندات الصرف والاستلام المخزني | نافذة منبثقة | - | STOCK_DOCS | frmInventory، frmStockParties، frmWarehouses |
| 46 | `frmStockNoteLines` | أسطر إذن التسليم أو الاستلام | شاشة فرعية في frmStockNote | StockNoteLines | مع الشاشة الأم | - |
| 47 | `frmStockNote` | أذونات التسليم والاستلام | نافذة منبثقة | - | STOCK_NOTES | frmInventory |
| 48 | `frmLogin` | تسجيل الدخول | نافذة منبثقة | - | الكل | - |
| 49 | `frmChangePassword` | كلمة المرور | نافذة منبثقة | - | الكل | frmMain، frmUsers |
| 50 | `frmRolePermLines` | صلاحيات الدور | شاشة فرعية في frmRoles | tmpRolePermissions | مع الشاشة الأم | - |
| 51 | `frmRoles` | الأدوار والصلاحيات | نافذة منبثقة | - | USERS | frmUsers |
| 52 | `frmUserScreenLines` | شاشات المستخدم | شاشة فرعية في frmUserScreens | tmpUserScreens | مع الشاشة الأم | - |
| 53 | `frmUserScreens` | صلاحيات الشاشات | نافذة منبثقة | - | USERS | frmUsers |
| 54 | `frmActivation` | تفعيل البرنامج | نافذة منبثقة | - | الكل | frmSettings |
| 55 | `frmBackup` | النسخ الاحتياطي | نافذة منبثقة | - | BACKUP | frmMain، القائمة الجانبية |
| 56 | `frmLabelLines` | أسطر الملصقات | شاشة فرعية في frmBarcodeLabels | - | مع الشاشة الأم | - |
| 57 | `frmBarcodeLabels` | طباعة ملصقات الباركود | شاشة كاملة | - | PRODUCTS | frmInventory |
| 58 | `frmTouchLines` | أسطر الطلب | شاشة فرعية في frmCafePOS | tmpPOSLines | مع الشاشة الأم | - |
| 59 | `frmTouchPOS` | نقطة بيع المطعم | شاشة كاملة | - | SALES_POS | - |
| 60 | `frmTouchPay` | الدفع نقدًا | نافذة منبثقة | - | SALES_POS | - |
| 61 | `frmCafePOS` | نقطة بيع الكافيه | شاشة كاملة | - | SALES_POS | - |
| 62 | `frmCafeItem` | خيارات المشروب | نافذة منبثقة | - | SALES_POS | - |
| 63 | `frmTreasury` | الخزينة | شاشة كاملة | - | CASH_CLOSING | frmCashBoxes، frmExpenses، frmMain، القائمة الجانبية |
| 64 | `frmCashVoucher` | سند نقدية | نافذة منبثقة | - | CASH_BOX | frmCommissions |
| 65 | `frmCashClosing` | تصفية يومية الكاشير | نافذة منبثقة | - | CASH_CLOSING | - |
| 66 | `frmCashTransfers` | التحويلات بالطريق | نافذة منبثقة | - | CASH_CLOSING | frmTreasury |
| 67 | `frmJournal` | قيود اليومية | شاشة كاملة | - | JOURNAL | frmAccounting، frmAccounts، frmFinancials |
| 68 | `frmJournalEntry` | قيد يومية | نافذة منبثقة | - | JOURNAL | - |
| 69 | `frmManualLines` | أسطر القيد اليدوي | شاشة فرعية في frmManualEntry | tmpManualLines | مع الشاشة الأم | - |
| 70 | `frmManualEntry` | القيود اليدوية | شاشة كاملة | - | MANUAL_ENTRY | frmAccounting، frmAccounts، frmJournal، frmLedger |
| 71 | `frmLedger` | كشف حساب | شاشة كاملة | - | JOURNAL | frmAccounting، frmAccounts، frmBanks، frmJournal |
| 72 | `frmFinancials` | القوائم المالية | شاشة كاملة | - | REPORTS_PROFIT | frmAccounting، frmLedger |
| 73 | `frmPeriodClosing` | إقفال الفترات والسنة المالية | نافذة منبثقة | - | PERIOD_CLOSE | frmAccounting، frmFinancials، frmFiscalPeriods |
| 74 | `frmFiscalPeriods` | الفترات المالية | نافذة منبثقة | - | PERIOD_CLOSE | frmPeriodClosing |
| 75 | `frmVatReturn` | إقرار ضريبة القيمة المضافة | نافذة منبثقة | - | VAT_RETURN | frmAccounting، frmFinancials |
| 76 | `frmAging` | أعمار الديون | نافذة منبثقة | - | REPORTS | frmAccounting، frmCustomers، frmSuppliers |
| 77 | `frmAllocation` | ربط السداد بالفواتير | نافذة منبثقة | - | CUSTOMER_PAYMENTS | frmCustomers، frmSuppliers |
| 78 | `frmBankTx` | الحركات البنكية | نافذة منبثقة | - | BANKS | frmBanks |
| 79 | `frmBankRecon` | التسوية البنكية | نافذة منبثقة | - | BANKS | frmBankTx، frmBanks |
| 80 | `frmCheques` | الشيكات | نافذة منبثقة | - | CHEQUES | frmAccounting، frmBanks، frmTreasury |
| 81 | `frmAssets` | الأصول الثابتة | نافذة منبثقة | - | FIXED_ASSETS | frmAccounting، frmDepreciation، frmFinancials |
| 82 | `frmDepreciation` | الإهلاك الشهري | نافذة منبثقة | - | FIXED_ASSETS | frmAccounting، frmAssets |
| 83 | `frmPayrollLines` | أسطر مسير الرواتب | شاشة فرعية في frmPayroll | PayrollLines | مع الشاشة الأم | - |
| 84 | `frmPayroll` | مسير الرواتب | نافذة منبثقة | - | PAYROLL | frmAccounting، frmEmployeePay، frmFinancials |
| 85 | `frmBudgetLines` | أسطر الموازنة | شاشة فرعية في frmBudget | BudgetLines | مع الشاشة الأم | - |
| 86 | `frmBudget` | الموازنة التقديرية | نافذة منبثقة | - | BUDGET | frmAccounting، frmFinancials |
| 87 | `frmAccounting` | المحاسبة والمالية | شاشة كاملة | - | الكل | frmMain، القائمة الجانبية |
| 88 | `frmAuditLog` | سجل التدقيق | نافذة منبثقة | - | AUDIT_LOG | frmAccounting، frmUsers |
| 89 | `frmCommissionLines` | أسطر مسير العمولات | شاشة فرعية في frmCommissions | CommissionLines | مع الشاشة الأم | - |
| 90 | `frmCommissions` | عمولات المندوبين | نافذة منبثقة | - | SALES_REPS | frmAccounting، frmSalesReps |
| 91 | `frmEnglishNameLines` | الأسماء الإنجليزية | شاشة فرعية في frmEnglishNames | tmpEnglishNames | مع الشاشة الأم | - |
| 92 | `frmEnglishNames` | الأسماء الإنجليزية | نافذة منبثقة | - | SETTINGS | frmSettings |
| 93 | `frmEInvoices` | الفاتورة الإلكترونية | نافذة منبثقة | - | EINVOICE | frmSettings |
| 94 | `frmZatcaSetup` | إعداد ربط منصة فاتورة | نافذة منبثقة | - | SETTINGS | - |
| 95 | `frmEtaSetup` | إعداد ربط منظومة الإيصال الإلكتروني | نافذة منبثقة | - | SETTINGS | - |
| 96 | `frmOpeningAccountLines` | الأرصدة الافتتاحية للحسابات | شاشة فرعية في frmOpeningBalances | OpeningBalances | مع الشاشة الأم | - |
| 97 | `frmOpeningCustomers` | الأرصدة الافتتاحية | شاشة كاملة | Customers | OPENING_BALANCES | - |
| 98 | `frmOpeningSuppliers` | الأرصدة الافتتاحية | شاشة كاملة | Suppliers | OPENING_BALANCES | - |
| 99 | `frmOpeningBoxes` | الأرصدة الافتتاحية | شاشة كاملة | CashBoxes | OPENING_BALANCES | - |
| 100 | `frmOpeningBanks` | الأرصدة الافتتاحية | شاشة كاملة | Banks | OPENING_BALANCES | - |
| 101 | `frmOpeningStock` | الرصيد الافتتاحي للمخزون | شاشة كاملة | OpeningStock | OPENING_BALANCES | - |
| 102 | `frmOpeningBalances` | أرصدة بداية التشغيل | شاشة كاملة | - | OPENING_BALANCES | frmAccounting |
| 103 | `frmSalesOrderLines` | أسطر طلبية العميل | شاشة فرعية في frmSalesOrder | SalesOrderLines | مع الشاشة الأم | - |
| 104 | `frmSalesOrder` | طلبيات العملاء | نافذة منبثقة | - | SALES_ORDERS | frmAccounting، frmSalesReps |
