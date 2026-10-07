# فهرس الشاشات (Screens Index)

> ملف مُولَّد تلقائيًا من `tools/forms*.py` بواسطة `tools/generate.py` – لا تعدّله يدويًا.

عدد الشاشات: **73**. شرح طريقة بناء الشاشات في [03-Forms.md](03-Forms.md).

| # | الشاشة | العنوان | النوع | الجدول | الصلاحية | تُفتح من |
|---|---|---|---|---|---|---|
| 1 | `frmMain` | نظام إدارة المحل | الشاشة الرئيسية | - | الكل | - |
| 2 | `frmProducts` | المنتجات | نافذة منبثقة | Products | PRODUCTS | frmInventory، frmMain، frmPurchaseInvoice، القائمة الجانبية |
| 3 | `frmCustomers` | العملاء | نافذة منبثقة | Customers | CUSTOMERS | frmMain، القائمة الجانبية |
| 4 | `frmSuppliers` | الموردون | نافذة منبثقة | Suppliers | SUPPLIERS | frmMain، القائمة الجانبية |
| 5 | `frmExpenses` | المصروفات | نافذة منبثقة | Expenses | EXPENSES | frmMain، frmRecurring، القائمة الجانبية |
| 6 | `frmRecurring` | المصروفات المتكررة | نافذة منبثقة | RecurringExpenses | EXPENSES | frmAccounting |
| 7 | `frmUsers` | المستخدمون | نافذة منبثقة | Employees | USERS | frmMain، القائمة الجانبية |
| 8 | `frmCostCenters` | مراكز التكلفة | نافذة منبثقة | CostCenters | JOURNAL | frmAccounting، frmAccounts |
| 9 | `frmEmployeePay` | رواتب الموظفين | نافذة منبثقة | Employees | PAYROLL | frmPayroll |
| 10 | `frmCategories` | التصنيفات | نافذة منبثقة | Categories | PRODUCTS | frmSettings |
| 11 | `frmUnits` | وحدات القياس | نافذة منبثقة | Units | PRODUCTS | frmSettings |
| 12 | `frmExpenseTypes` | أنواع المصروفات | نافذة منبثقة | ExpenseTypes | EXPENSES | frmExpenses، frmSettings |
| 13 | `frmCashBoxes` | الصناديق | نافذة منبثقة | CashBoxes | CASH_BOX | frmTreasury |
| 14 | `frmBanks` | البنوك | نافذة منبثقة | Banks | BANKS | frmAccounting، frmTreasury |
| 15 | `frmAccounts` | دليل الحسابات | نافذة منبثقة | Accounts | JOURNAL | frmAccounting، frmJournal، frmLedger، frmManualEntry |
| 16 | `frmSettings` | الإعدادات | نافذة منبثقة | Settings | SETTINGS | frmMain، القائمة الجانبية |
| 17 | `frmLabelSettings` | إعدادات ملصقات الباركود | نافذة منبثقة | LabelSettings | PRODUCTS | frmBarcodeLabels، frmSettings |
| 18 | `frmSearch` | البحث المتقدم | شاشة كاملة | - | الكل | frmMain، القائمة الجانبية |
| 19 | `frmReportCenter` | التقارير | شاشة كاملة | - | REPORTS | frmMain، القائمة الجانبية |
| 20 | `frmPOSLines` | أسطر الفاتورة | شاشة فرعية في frmPOS | tmpPOSLines | مع الشاشة الأم | - |
| 21 | `frmPOS` | نقطة البيع | شاشة كاملة | - | SALES_POS | frmMain، القائمة الجانبية |
| 22 | `frmReturnLines` | أسطر المرتجع | شاشة فرعية في frmSalesReturn | tmpReturnLines | مع الشاشة الأم | - |
| 23 | `frmSalesReturn` | مرتجع مبيعات | نافذة منبثقة | - | SALES_RETURN | frmPOS، frmSalesInvoice |
| 24 | `frmCustomerPayment` | سند قبض | نافذة منبثقة | - | CUSTOMER_PAYMENTS | frmCustomers، frmPOS |
| 25 | `frmSalesInvoice` | فاتورة بيع | نافذة منبثقة | - | SALES_VIEW | - |
| 26 | `frmPurchaseLines` | أسطر فاتورة الشراء | شاشة فرعية في frmPurchaseInvoice | tmpPurchaseLines | مع الشاشة الأم | - |
| 27 | `frmPurchaseInvoice` | فاتورة مشتريات | شاشة كاملة | - | PURCHASES | frmInventory، frmMain، القائمة الجانبية |
| 28 | `frmPurchaseReturnLines` | أسطر مرتجع المشتريات | شاشة فرعية في frmPurchaseReturn | tmpPurchaseReturnLines | مع الشاشة الأم | - |
| 29 | `frmPurchaseReturn` | مرتجع مشتريات | نافذة منبثقة | - | PURCHASE_RETURN | frmPurchaseInvoice، frmPurchaseView |
| 30 | `frmSupplierPayment` | سند صرف | نافذة منبثقة | - | SUPPLIER_PAYMENTS | frmPurchaseInvoice، frmPurchaseView، frmSuppliers |
| 31 | `frmPurchaseView` | فاتورة شراء | نافذة منبثقة | - | PURCHASES | - |
| 32 | `frmInventory` | المخزون | شاشة كاملة | - | PRODUCTS | frmMain، القائمة الجانبية |
| 33 | `frmStockCountLines` | أسطر الجرد | شاشة فرعية في frmStockCount | StockCountDetails | مع الشاشة الأم | - |
| 34 | `frmStockCount` | الجرد | شاشة كاملة | - | STOCK_COUNT | frmInventory، frmMain، القائمة الجانبية |
| 35 | `frmLogin` | تسجيل الدخول | نافذة منبثقة | - | الكل | - |
| 36 | `frmChangePassword` | كلمة المرور | نافذة منبثقة | - | الكل | frmMain، frmUsers |
| 37 | `frmRolePermLines` | صلاحيات الدور | شاشة فرعية في frmRoles | tmpRolePermissions | مع الشاشة الأم | - |
| 38 | `frmRoles` | الأدوار والصلاحيات | نافذة منبثقة | - | USERS | frmUsers |
| 39 | `frmUserScreenLines` | شاشات المستخدم | شاشة فرعية في frmUserScreens | tmpUserScreens | مع الشاشة الأم | - |
| 40 | `frmUserScreens` | صلاحيات الشاشات | نافذة منبثقة | - | USERS | frmUsers |
| 41 | `frmActivation` | تفعيل البرنامج | نافذة منبثقة | - | الكل | frmSettings |
| 42 | `frmBackup` | النسخ الاحتياطي | نافذة منبثقة | - | BACKUP | frmMain، القائمة الجانبية |
| 43 | `frmLabelLines` | أسطر الملصقات | شاشة فرعية في frmBarcodeLabels | - | مع الشاشة الأم | - |
| 44 | `frmBarcodeLabels` | طباعة ملصقات الباركود | شاشة كاملة | - | PRODUCTS | frmInventory |
| 45 | `frmTouchLines` | أسطر الطلب | شاشة فرعية في frmCafePOS | tmpPOSLines | مع الشاشة الأم | - |
| 46 | `frmTouchPOS` | نقطة بيع المطعم | شاشة كاملة | - | SALES_POS | - |
| 47 | `frmTouchPay` | الدفع نقدًا | نافذة منبثقة | - | SALES_POS | - |
| 48 | `frmCafePOS` | نقطة بيع الكافيه | شاشة كاملة | - | SALES_POS | - |
| 49 | `frmCafeItem` | خيارات المشروب | نافذة منبثقة | - | SALES_POS | - |
| 50 | `frmTreasury` | الخزينة | شاشة كاملة | - | CASH_CLOSING | frmCashBoxes، frmExpenses، frmMain، القائمة الجانبية |
| 51 | `frmCashVoucher` | سند نقدية | نافذة منبثقة | - | CASH_BOX | - |
| 52 | `frmCashClosing` | تصفية يومية الكاشير | نافذة منبثقة | - | CASH_CLOSING | - |
| 53 | `frmJournal` | قيود اليومية | شاشة كاملة | - | JOURNAL | frmAccounting، frmAccounts، frmFinancials |
| 54 | `frmJournalEntry` | قيد يومية | نافذة منبثقة | - | JOURNAL | - |
| 55 | `frmManualLines` | أسطر القيد اليدوي | شاشة فرعية في frmManualEntry | tmpManualLines | مع الشاشة الأم | - |
| 56 | `frmManualEntry` | القيود اليدوية | شاشة كاملة | - | MANUAL_ENTRY | frmAccounting، frmAccounts، frmJournal، frmLedger |
| 57 | `frmLedger` | كشف حساب | شاشة كاملة | - | JOURNAL | frmAccounting، frmAccounts، frmBanks، frmJournal |
| 58 | `frmFinancials` | القوائم المالية | شاشة كاملة | - | REPORTS_PROFIT | frmAccounting، frmLedger |
| 59 | `frmPeriodClosing` | إقفال الفترات والسنة المالية | نافذة منبثقة | - | PERIOD_CLOSE | frmAccounting، frmFinancials |
| 60 | `frmVatReturn` | إقرار ضريبة القيمة المضافة | نافذة منبثقة | - | VAT_RETURN | frmAccounting، frmFinancials |
| 61 | `frmAging` | أعمار الديون | نافذة منبثقة | - | REPORTS | frmAccounting، frmCustomers، frmSuppliers |
| 62 | `frmAllocation` | ربط السداد بالفواتير | نافذة منبثقة | - | CUSTOMER_PAYMENTS | frmCustomers، frmSuppliers |
| 63 | `frmBankTx` | الحركات البنكية | نافذة منبثقة | - | BANKS | frmBanks |
| 64 | `frmBankRecon` | التسوية البنكية | نافذة منبثقة | - | BANKS | frmBankTx، frmBanks |
| 65 | `frmCheques` | الشيكات | نافذة منبثقة | - | CHEQUES | frmAccounting، frmBanks، frmTreasury |
| 66 | `frmAssets` | الأصول الثابتة | نافذة منبثقة | - | FIXED_ASSETS | frmAccounting، frmDepreciation، frmFinancials |
| 67 | `frmDepreciation` | الإهلاك الشهري | نافذة منبثقة | - | FIXED_ASSETS | frmAccounting، frmAssets |
| 68 | `frmPayrollLines` | أسطر مسير الرواتب | شاشة فرعية في frmPayroll | PayrollLines | مع الشاشة الأم | - |
| 69 | `frmPayroll` | مسير الرواتب | نافذة منبثقة | - | PAYROLL | frmAccounting، frmEmployeePay، frmFinancials |
| 70 | `frmBudgetLines` | أسطر الموازنة | شاشة فرعية في frmBudget | BudgetLines | مع الشاشة الأم | - |
| 71 | `frmBudget` | الموازنة التقديرية | نافذة منبثقة | - | BUDGET | frmAccounting، frmFinancials |
| 72 | `frmAccounting` | المحاسبة والمالية | شاشة كاملة | - | الكل | frmMain، القائمة الجانبية |
| 73 | `frmAuditLog` | سجل التدقيق | نافذة منبثقة | - | AUDIT_LOG | frmAccounting، frmUsers |
