# فهرس الشاشات (Screens Index)

> ملف مُولَّد تلقائيًا من `tools/forms*.py` بواسطة `tools/generate.py` – لا تعدّله يدويًا.

عدد الشاشات: **95**. شرح طريقة بناء الشاشات في [03-Forms.md](03-Forms.md).

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
| 20 | `frmBanks` | البنوك | نافذة منبثقة | Banks | BANKS | frmAccounting، frmTreasury |
| 21 | `frmAccounts` | دليل الحسابات | نافذة منبثقة | Accounts | JOURNAL | frmAccounting، frmJournal، frmLedger، frmManualEntry |
| 22 | `frmSettings` | الإعدادات | نافذة منبثقة | Settings | SETTINGS | frmMain، القائمة الجانبية |
| 23 | `frmLabelSettings` | إعدادات ملصقات الباركود | نافذة منبثقة | LabelSettings | PRODUCTS | frmBarcodeLabels، frmSettings |
| 24 | `frmSearch` | البحث المتقدم | شاشة كاملة | - | الكل | frmMain، القائمة الجانبية |
| 25 | `frmReportCenter` | التقارير | شاشة كاملة | - | REPORTS | frmMain، frmSalesReps، frmWarehouses، القائمة الجانبية |
| 26 | `frmPOSLines` | أسطر الفاتورة | شاشة فرعية في frmPOS | tmpPOSLines | مع الشاشة الأم | - |
| 27 | `frmPOS` | نقطة البيع | شاشة كاملة | - | SALES_POS | frmMain، القائمة الجانبية |
| 28 | `frmReturnLines` | أسطر المرتجع | شاشة فرعية في frmSalesReturn | tmpReturnLines | مع الشاشة الأم | - |
| 29 | `frmSalesReturn` | مرتجع مبيعات | نافذة منبثقة | - | SALES_RETURN | frmPOS، frmSalesInvoice |
| 30 | `frmCustomerPayment` | سند قبض | نافذة منبثقة | - | CUSTOMER_PAYMENTS | frmCustomers، frmPOS |
| 31 | `frmSalesInvoice` | فاتورة بيع | نافذة منبثقة | - | SALES_VIEW | - |
| 32 | `frmPurchaseLines` | أسطر فاتورة الشراء | شاشة فرعية في frmPurchaseInvoice | tmpPurchaseLines | مع الشاشة الأم | - |
| 33 | `frmPurchaseInvoice` | فاتورة مشتريات | شاشة كاملة | - | PURCHASES | frmInventory، frmMain، القائمة الجانبية |
| 34 | `frmPurchaseReturnLines` | أسطر مرتجع المشتريات | شاشة فرعية في frmPurchaseReturn | tmpPurchaseReturnLines | مع الشاشة الأم | - |
| 35 | `frmPurchaseReturn` | مرتجع مشتريات | نافذة منبثقة | - | PURCHASE_RETURN | frmPurchaseInvoice، frmPurchaseView |
| 36 | `frmSupplierPayment` | سند صرف | نافذة منبثقة | - | SUPPLIER_PAYMENTS | frmPurchaseInvoice، frmPurchaseView، frmSuppliers |
| 37 | `frmPurchaseView` | فاتورة شراء | نافذة منبثقة | - | PURCHASES | - |
| 38 | `frmInventory` | المخزون | شاشة كاملة | - | PRODUCTS | frmMain، frmWarehouses، القائمة الجانبية |
| 39 | `frmStockCountLines` | أسطر الجرد | شاشة فرعية في frmStockCount | StockCountDetails | مع الشاشة الأم | - |
| 40 | `frmStockCount` | الجرد | شاشة كاملة | - | STOCK_COUNT | frmInventory، frmMain، القائمة الجانبية |
| 41 | `frmLogin` | تسجيل الدخول | نافذة منبثقة | - | الكل | - |
| 42 | `frmChangePassword` | كلمة المرور | نافذة منبثقة | - | الكل | frmMain، frmUsers |
| 43 | `frmRolePermLines` | صلاحيات الدور | شاشة فرعية في frmRoles | tmpRolePermissions | مع الشاشة الأم | - |
| 44 | `frmRoles` | الأدوار والصلاحيات | نافذة منبثقة | - | USERS | frmUsers |
| 45 | `frmUserScreenLines` | شاشات المستخدم | شاشة فرعية في frmUserScreens | tmpUserScreens | مع الشاشة الأم | - |
| 46 | `frmUserScreens` | صلاحيات الشاشات | نافذة منبثقة | - | USERS | frmUsers |
| 47 | `frmActivation` | تفعيل البرنامج | نافذة منبثقة | - | الكل | frmSettings |
| 48 | `frmBackup` | النسخ الاحتياطي | نافذة منبثقة | - | BACKUP | frmMain، القائمة الجانبية |
| 49 | `frmLabelLines` | أسطر الملصقات | شاشة فرعية في frmBarcodeLabels | - | مع الشاشة الأم | - |
| 50 | `frmBarcodeLabels` | طباعة ملصقات الباركود | شاشة كاملة | - | PRODUCTS | frmInventory |
| 51 | `frmTouchLines` | أسطر الطلب | شاشة فرعية في frmCafePOS | tmpPOSLines | مع الشاشة الأم | - |
| 52 | `frmTouchPOS` | نقطة بيع المطعم | شاشة كاملة | - | SALES_POS | - |
| 53 | `frmTouchPay` | الدفع نقدًا | نافذة منبثقة | - | SALES_POS | - |
| 54 | `frmCafePOS` | نقطة بيع الكافيه | شاشة كاملة | - | SALES_POS | - |
| 55 | `frmCafeItem` | خيارات المشروب | نافذة منبثقة | - | SALES_POS | - |
| 56 | `frmTreasury` | الخزينة | شاشة كاملة | - | CASH_CLOSING | frmCashBoxes، frmExpenses، frmMain، القائمة الجانبية |
| 57 | `frmCashVoucher` | سند نقدية | نافذة منبثقة | - | CASH_BOX | frmCommissions |
| 58 | `frmCashClosing` | تصفية يومية الكاشير | نافذة منبثقة | - | CASH_CLOSING | - |
| 59 | `frmCashTransfers` | التحويلات بالطريق | نافذة منبثقة | - | CASH_CLOSING | frmTreasury |
| 60 | `frmJournal` | قيود اليومية | شاشة كاملة | - | JOURNAL | frmAccounting، frmAccounts، frmFinancials |
| 61 | `frmJournalEntry` | قيد يومية | نافذة منبثقة | - | JOURNAL | - |
| 62 | `frmManualLines` | أسطر القيد اليدوي | شاشة فرعية في frmManualEntry | tmpManualLines | مع الشاشة الأم | - |
| 63 | `frmManualEntry` | القيود اليدوية | شاشة كاملة | - | MANUAL_ENTRY | frmAccounting، frmAccounts، frmJournal، frmLedger |
| 64 | `frmLedger` | كشف حساب | شاشة كاملة | - | JOURNAL | frmAccounting، frmAccounts، frmBanks، frmJournal |
| 65 | `frmFinancials` | القوائم المالية | شاشة كاملة | - | REPORTS_PROFIT | frmAccounting، frmLedger |
| 66 | `frmPeriodClosing` | إقفال الفترات والسنة المالية | نافذة منبثقة | - | PERIOD_CLOSE | frmAccounting، frmFinancials، frmFiscalPeriods |
| 67 | `frmFiscalPeriods` | الفترات المالية | نافذة منبثقة | - | PERIOD_CLOSE | frmPeriodClosing |
| 68 | `frmVatReturn` | إقرار ضريبة القيمة المضافة | نافذة منبثقة | - | VAT_RETURN | frmAccounting، frmFinancials |
| 69 | `frmAging` | أعمار الديون | نافذة منبثقة | - | REPORTS | frmAccounting، frmCustomers، frmSuppliers |
| 70 | `frmAllocation` | ربط السداد بالفواتير | نافذة منبثقة | - | CUSTOMER_PAYMENTS | frmCustomers، frmSuppliers |
| 71 | `frmBankTx` | الحركات البنكية | نافذة منبثقة | - | BANKS | frmBanks |
| 72 | `frmBankRecon` | التسوية البنكية | نافذة منبثقة | - | BANKS | frmBankTx، frmBanks |
| 73 | `frmCheques` | الشيكات | نافذة منبثقة | - | CHEQUES | frmAccounting، frmBanks، frmTreasury |
| 74 | `frmAssets` | الأصول الثابتة | نافذة منبثقة | - | FIXED_ASSETS | frmAccounting، frmDepreciation، frmFinancials |
| 75 | `frmDepreciation` | الإهلاك الشهري | نافذة منبثقة | - | FIXED_ASSETS | frmAccounting، frmAssets |
| 76 | `frmPayrollLines` | أسطر مسير الرواتب | شاشة فرعية في frmPayroll | PayrollLines | مع الشاشة الأم | - |
| 77 | `frmPayroll` | مسير الرواتب | نافذة منبثقة | - | PAYROLL | frmAccounting، frmEmployeePay، frmFinancials |
| 78 | `frmBudgetLines` | أسطر الموازنة | شاشة فرعية في frmBudget | BudgetLines | مع الشاشة الأم | - |
| 79 | `frmBudget` | الموازنة التقديرية | نافذة منبثقة | - | BUDGET | frmAccounting، frmFinancials |
| 80 | `frmAccounting` | المحاسبة والمالية | شاشة كاملة | - | الكل | frmMain، القائمة الجانبية |
| 81 | `frmAuditLog` | سجل التدقيق | نافذة منبثقة | - | AUDIT_LOG | frmAccounting، frmUsers |
| 82 | `frmCommissionLines` | أسطر مسير العمولات | شاشة فرعية في frmCommissions | CommissionLines | مع الشاشة الأم | - |
| 83 | `frmCommissions` | عمولات المندوبين | نافذة منبثقة | - | SALES_REPS | frmAccounting، frmSalesReps |
| 84 | `frmEnglishNameLines` | الأسماء الإنجليزية | شاشة فرعية في frmEnglishNames | tmpEnglishNames | مع الشاشة الأم | - |
| 85 | `frmEnglishNames` | الأسماء الإنجليزية | نافذة منبثقة | - | SETTINGS | frmSettings |
| 86 | `frmEInvoices` | الفاتورة الإلكترونية | نافذة منبثقة | - | EINVOICE | frmSettings |
| 87 | `frmZatcaSetup` | إعداد ربط منصة فاتورة | نافذة منبثقة | - | SETTINGS | - |
| 88 | `frmEtaSetup` | إعداد ربط منظومة الإيصال الإلكتروني | نافذة منبثقة | - | SETTINGS | - |
| 89 | `frmOpeningAccountLines` | الأرصدة الافتتاحية للحسابات | شاشة فرعية في frmOpeningBalances | OpeningBalances | مع الشاشة الأم | - |
| 90 | `frmOpeningCustomers` | الأرصدة الافتتاحية | شاشة كاملة | Customers | OPENING_BALANCES | - |
| 91 | `frmOpeningSuppliers` | الأرصدة الافتتاحية | شاشة كاملة | Suppliers | OPENING_BALANCES | - |
| 92 | `frmOpeningBoxes` | الأرصدة الافتتاحية | شاشة كاملة | CashBoxes | OPENING_BALANCES | - |
| 93 | `frmOpeningBanks` | الأرصدة الافتتاحية | شاشة كاملة | Banks | OPENING_BALANCES | - |
| 94 | `frmOpeningStock` | الرصيد الافتتاحي للمخزون | شاشة كاملة | OpeningStock | OPENING_BALANCES | - |
| 95 | `frmOpeningBalances` | أرصدة بداية التشغيل | شاشة كاملة | - | OPENING_BALANCES | frmAccounting |
