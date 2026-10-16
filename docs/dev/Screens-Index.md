# فهرس الشاشات (Screens Index)

> ملف مُولَّد تلقائيًا من `tools/forms*.py` بواسطة `tools/generate.py` – لا تعدّله يدويًا.

عدد الشاشات: **94**. شرح طريقة بناء الشاشات في [03-Forms.md](03-Forms.md).

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
| 19 | `frmBanks` | البنوك | نافذة منبثقة | Banks | BANKS | frmAccounting، frmTreasury |
| 20 | `frmAccounts` | دليل الحسابات | نافذة منبثقة | Accounts | JOURNAL | frmAccounting، frmJournal، frmLedger، frmManualEntry |
| 21 | `frmSettings` | الإعدادات | نافذة منبثقة | Settings | SETTINGS | frmMain، القائمة الجانبية |
| 22 | `frmLabelSettings` | إعدادات ملصقات الباركود | نافذة منبثقة | LabelSettings | PRODUCTS | frmBarcodeLabels، frmSettings |
| 23 | `frmSearch` | البحث المتقدم | شاشة كاملة | - | الكل | frmMain، القائمة الجانبية |
| 24 | `frmReportCenter` | التقارير | شاشة كاملة | - | REPORTS | frmMain، frmSalesReps، القائمة الجانبية |
| 25 | `frmPOSLines` | أسطر الفاتورة | شاشة فرعية في frmPOS | tmpPOSLines | مع الشاشة الأم | - |
| 26 | `frmPOS` | نقطة البيع | شاشة كاملة | - | SALES_POS | frmMain، القائمة الجانبية |
| 27 | `frmReturnLines` | أسطر المرتجع | شاشة فرعية في frmSalesReturn | tmpReturnLines | مع الشاشة الأم | - |
| 28 | `frmSalesReturn` | مرتجع مبيعات | نافذة منبثقة | - | SALES_RETURN | frmPOS، frmSalesInvoice |
| 29 | `frmCustomerPayment` | سند قبض | نافذة منبثقة | - | CUSTOMER_PAYMENTS | frmCustomers، frmPOS |
| 30 | `frmSalesInvoice` | فاتورة بيع | نافذة منبثقة | - | SALES_VIEW | - |
| 31 | `frmPurchaseLines` | أسطر فاتورة الشراء | شاشة فرعية في frmPurchaseInvoice | tmpPurchaseLines | مع الشاشة الأم | - |
| 32 | `frmPurchaseInvoice` | فاتورة مشتريات | شاشة كاملة | - | PURCHASES | frmInventory، frmMain، القائمة الجانبية |
| 33 | `frmPurchaseReturnLines` | أسطر مرتجع المشتريات | شاشة فرعية في frmPurchaseReturn | tmpPurchaseReturnLines | مع الشاشة الأم | - |
| 34 | `frmPurchaseReturn` | مرتجع مشتريات | نافذة منبثقة | - | PURCHASE_RETURN | frmPurchaseInvoice، frmPurchaseView |
| 35 | `frmSupplierPayment` | سند صرف | نافذة منبثقة | - | SUPPLIER_PAYMENTS | frmPurchaseInvoice، frmPurchaseView، frmSuppliers |
| 36 | `frmPurchaseView` | فاتورة شراء | نافذة منبثقة | - | PURCHASES | - |
| 37 | `frmInventory` | المخزون | شاشة كاملة | - | PRODUCTS | frmMain، القائمة الجانبية |
| 38 | `frmStockCountLines` | أسطر الجرد | شاشة فرعية في frmStockCount | StockCountDetails | مع الشاشة الأم | - |
| 39 | `frmStockCount` | الجرد | شاشة كاملة | - | STOCK_COUNT | frmInventory، frmMain، القائمة الجانبية |
| 40 | `frmLogin` | تسجيل الدخول | نافذة منبثقة | - | الكل | - |
| 41 | `frmChangePassword` | كلمة المرور | نافذة منبثقة | - | الكل | frmMain، frmUsers |
| 42 | `frmRolePermLines` | صلاحيات الدور | شاشة فرعية في frmRoles | tmpRolePermissions | مع الشاشة الأم | - |
| 43 | `frmRoles` | الأدوار والصلاحيات | نافذة منبثقة | - | USERS | frmUsers |
| 44 | `frmUserScreenLines` | شاشات المستخدم | شاشة فرعية في frmUserScreens | tmpUserScreens | مع الشاشة الأم | - |
| 45 | `frmUserScreens` | صلاحيات الشاشات | نافذة منبثقة | - | USERS | frmUsers |
| 46 | `frmActivation` | تفعيل البرنامج | نافذة منبثقة | - | الكل | frmSettings |
| 47 | `frmBackup` | النسخ الاحتياطي | نافذة منبثقة | - | BACKUP | frmMain، القائمة الجانبية |
| 48 | `frmLabelLines` | أسطر الملصقات | شاشة فرعية في frmBarcodeLabels | - | مع الشاشة الأم | - |
| 49 | `frmBarcodeLabels` | طباعة ملصقات الباركود | شاشة كاملة | - | PRODUCTS | frmInventory |
| 50 | `frmTouchLines` | أسطر الطلب | شاشة فرعية في frmCafePOS | tmpPOSLines | مع الشاشة الأم | - |
| 51 | `frmTouchPOS` | نقطة بيع المطعم | شاشة كاملة | - | SALES_POS | - |
| 52 | `frmTouchPay` | الدفع نقدًا | نافذة منبثقة | - | SALES_POS | - |
| 53 | `frmCafePOS` | نقطة بيع الكافيه | شاشة كاملة | - | SALES_POS | - |
| 54 | `frmCafeItem` | خيارات المشروب | نافذة منبثقة | - | SALES_POS | - |
| 55 | `frmTreasury` | الخزينة | شاشة كاملة | - | CASH_CLOSING | frmCashBoxes، frmExpenses، frmMain، القائمة الجانبية |
| 56 | `frmCashVoucher` | سند نقدية | نافذة منبثقة | - | CASH_BOX | frmCommissions |
| 57 | `frmCashClosing` | تصفية يومية الكاشير | نافذة منبثقة | - | CASH_CLOSING | - |
| 58 | `frmCashTransfers` | التحويلات بالطريق | نافذة منبثقة | - | CASH_CLOSING | frmTreasury |
| 59 | `frmJournal` | قيود اليومية | شاشة كاملة | - | JOURNAL | frmAccounting، frmAccounts، frmFinancials |
| 60 | `frmJournalEntry` | قيد يومية | نافذة منبثقة | - | JOURNAL | - |
| 61 | `frmManualLines` | أسطر القيد اليدوي | شاشة فرعية في frmManualEntry | tmpManualLines | مع الشاشة الأم | - |
| 62 | `frmManualEntry` | القيود اليدوية | شاشة كاملة | - | MANUAL_ENTRY | frmAccounting، frmAccounts، frmJournal، frmLedger |
| 63 | `frmLedger` | كشف حساب | شاشة كاملة | - | JOURNAL | frmAccounting، frmAccounts، frmBanks، frmJournal |
| 64 | `frmFinancials` | القوائم المالية | شاشة كاملة | - | REPORTS_PROFIT | frmAccounting، frmLedger |
| 65 | `frmPeriodClosing` | إقفال الفترات والسنة المالية | نافذة منبثقة | - | PERIOD_CLOSE | frmAccounting، frmFinancials، frmFiscalPeriods |
| 66 | `frmFiscalPeriods` | الفترات المالية | نافذة منبثقة | - | PERIOD_CLOSE | frmPeriodClosing |
| 67 | `frmVatReturn` | إقرار ضريبة القيمة المضافة | نافذة منبثقة | - | VAT_RETURN | frmAccounting، frmFinancials |
| 68 | `frmAging` | أعمار الديون | نافذة منبثقة | - | REPORTS | frmAccounting، frmCustomers، frmSuppliers |
| 69 | `frmAllocation` | ربط السداد بالفواتير | نافذة منبثقة | - | CUSTOMER_PAYMENTS | frmCustomers، frmSuppliers |
| 70 | `frmBankTx` | الحركات البنكية | نافذة منبثقة | - | BANKS | frmBanks |
| 71 | `frmBankRecon` | التسوية البنكية | نافذة منبثقة | - | BANKS | frmBankTx، frmBanks |
| 72 | `frmCheques` | الشيكات | نافذة منبثقة | - | CHEQUES | frmAccounting، frmBanks، frmTreasury |
| 73 | `frmAssets` | الأصول الثابتة | نافذة منبثقة | - | FIXED_ASSETS | frmAccounting، frmDepreciation، frmFinancials |
| 74 | `frmDepreciation` | الإهلاك الشهري | نافذة منبثقة | - | FIXED_ASSETS | frmAccounting، frmAssets |
| 75 | `frmPayrollLines` | أسطر مسير الرواتب | شاشة فرعية في frmPayroll | PayrollLines | مع الشاشة الأم | - |
| 76 | `frmPayroll` | مسير الرواتب | نافذة منبثقة | - | PAYROLL | frmAccounting، frmEmployeePay، frmFinancials |
| 77 | `frmBudgetLines` | أسطر الموازنة | شاشة فرعية في frmBudget | BudgetLines | مع الشاشة الأم | - |
| 78 | `frmBudget` | الموازنة التقديرية | نافذة منبثقة | - | BUDGET | frmAccounting، frmFinancials |
| 79 | `frmAccounting` | المحاسبة والمالية | شاشة كاملة | - | الكل | frmMain، القائمة الجانبية |
| 80 | `frmAuditLog` | سجل التدقيق | نافذة منبثقة | - | AUDIT_LOG | frmAccounting، frmUsers |
| 81 | `frmCommissionLines` | أسطر مسير العمولات | شاشة فرعية في frmCommissions | CommissionLines | مع الشاشة الأم | - |
| 82 | `frmCommissions` | عمولات المندوبين | نافذة منبثقة | - | SALES_REPS | frmAccounting، frmSalesReps |
| 83 | `frmEnglishNameLines` | الأسماء الإنجليزية | شاشة فرعية في frmEnglishNames | tmpEnglishNames | مع الشاشة الأم | - |
| 84 | `frmEnglishNames` | الأسماء الإنجليزية | نافذة منبثقة | - | SETTINGS | frmSettings |
| 85 | `frmEInvoices` | الفاتورة الإلكترونية | نافذة منبثقة | - | EINVOICE | frmSettings |
| 86 | `frmZatcaSetup` | إعداد ربط منصة فاتورة | نافذة منبثقة | - | SETTINGS | - |
| 87 | `frmEtaSetup` | إعداد ربط منظومة الإيصال الإلكتروني | نافذة منبثقة | - | SETTINGS | - |
| 88 | `frmOpeningAccountLines` | الأرصدة الافتتاحية للحسابات | شاشة فرعية في frmOpeningBalances | OpeningBalances | مع الشاشة الأم | - |
| 89 | `frmOpeningCustomers` | الأرصدة الافتتاحية | شاشة كاملة | Customers | OPENING_BALANCES | - |
| 90 | `frmOpeningSuppliers` | الأرصدة الافتتاحية | شاشة كاملة | Suppliers | OPENING_BALANCES | - |
| 91 | `frmOpeningBoxes` | الأرصدة الافتتاحية | شاشة كاملة | CashBoxes | OPENING_BALANCES | - |
| 92 | `frmOpeningBanks` | الأرصدة الافتتاحية | شاشة كاملة | Banks | OPENING_BALANCES | - |
| 93 | `frmOpeningStock` | الرصيد الافتتاحي للمخزون | شاشة كاملة | OpeningStock | OPENING_BALANCES | - |
| 94 | `frmOpeningBalances` | أرصدة بداية التشغيل | شاشة كاملة | - | OPENING_BALANCES | frmAccounting |
