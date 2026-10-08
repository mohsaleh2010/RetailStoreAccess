# مرجع الجداول (Tables Reference)

> ملف مُولَّد تلقائيًا من `tools/schema.py` بواسطة `tools/generate.py` – لا تعدّله يدويًا.

عدد الجداول: **69** | عدد الحقول: **877**

## الفهرس

1. [`Settings`](#settings) – إعدادات المحل
2. [`Sequences`](#sequences) – عدّادات الترقيم
3. [`Roles`](#roles) – الأدوار
4. [`Permissions`](#permissions) – الصلاحيات
5. [`RolePermissions`](#rolepermissions) – صلاحيات الأدوار
6. [`Employees`](#employees) – الموظفون والمستخدمون
7. [`Screens`](#screens) – الشاشات
8. [`UserScreens`](#userscreens) – صلاحيات الشاشات للمستخدم
9. [`Activations`](#activations) – تفعيل البرنامج
10. [`Categories`](#categories) – التصنيفات
11. [`Units`](#units) – وحدات القياس
12. [`PaymentMethods`](#paymentmethods) – طرق الدفع
13. [`Currencies`](#currencies) – العملات
14. [`CurrencyRates`](#currencyrates) – أسعار العملات
15. [`CashBoxes`](#cashboxes) – الخزينة والصناديق
16. [`Suppliers`](#suppliers) – الموردون
17. [`Customers`](#customers) – العملاء
18. [`Products`](#products) – المنتجات
19. [`SalesInvoices`](#salesinvoices) – فواتير المبيعات
20. [`SalesInvoiceDetails`](#salesinvoicedetails) – تفاصيل فواتير المبيعات
21. [`SalesReturns`](#salesreturns) – مرتجعات المبيعات
22. [`SalesReturnDetails`](#salesreturndetails) – تفاصيل مرتجعات المبيعات
23. [`PurchaseInvoices`](#purchaseinvoices) – فواتير المشتريات
24. [`PurchaseInvoiceDetails`](#purchaseinvoicedetails) – تفاصيل فواتير المشتريات
25. [`PurchaseReturns`](#purchasereturns) – مرتجعات المشتريات
26. [`PurchaseReturnDetails`](#purchasereturndetails) – تفاصيل مرتجعات المشتريات
27. [`CustomerPayments`](#customerpayments) – دفعات العملاء (سندات القبض)
28. [`SupplierPayments`](#supplierpayments) – دفعات الموردين (سندات الصرف)
29. [`Banks`](#banks) – البنوك
30. [`BankTransactions`](#banktransactions) – الحركات البنكية
31. [`Cheques`](#cheques) – الشيكات الواردة والصادرة
32. [`FixedAssets`](#fixedassets) – الأصول الثابتة
33. [`DepreciationRuns`](#depreciationruns) – قيود الإهلاك الشهرية
34. [`AssetDepreciations`](#assetdepreciations) – إهلاك كل أصل في كل شهر
35. [`CostCenters`](#costcenters) – مراكز التكلفة والفروع
36. [`SalesReps`](#salesreps) – المندوبين
37. [`SalesRepTargets`](#salesreptargets) – أهداف المندوبين
38. [`CommissionRuns`](#commissionruns) – مسيرات العمولات
39. [`CommissionLines`](#commissionlines) – أسطر مسير العمولات
40. [`Budgets`](#budgets) – الموازنات التقديرية
41. [`BudgetLines`](#budgetlines) – أسطر الموازنة
42. [`PayrollRuns`](#payrollruns) – مسيرات الرواتب
43. [`PayrollLines`](#payrolllines) – أسطر مسير الرواتب
44. [`BankReconciliations`](#bankreconciliations) – التسويات البنكية
45. [`BankClearings`](#bankclearings) – حركات الدفاتر المطابقة لكشف البنك
46. [`CustomerAllocations`](#customerallocations) – ربط سندات القبض بالفواتير
47. [`SupplierAllocations`](#supplierallocations) – ربط سندات الصرف بفواتير الشراء
48. [`ExpenseTypes`](#expensetypes) – أنواع المصروفات
49. [`Expenses`](#expenses) – المصروفات
50. [`RecurringExpenses`](#recurringexpenses) – المصروفات المتكررة
51. [`CashVouchers`](#cashvouchers) – سندات النقدية
52. [`CashClosings`](#cashclosings) – تصفية يومية الكاشير
53. [`Accounts`](#accounts) – دليل الحسابات (شجرة الحسابات)
54. [`JournalSourceTypes`](#journalsourcetypes) – أنواع مصادر القيود
55. [`JournalEntries`](#journalentries) – قيود اليومية
56. [`JournalLines`](#journallines) – أسطر القيود
57. [`PeriodClosings`](#periodclosings) – سجل إقفال الفترات
58. [`FiscalYearClosings`](#fiscalyearclosings) – إقفال السنوات المالية
59. [`FiscalYearClosingLines`](#fiscalyearclosinglines) – أسطر قيود إقفال السنوات
60. [`VatReturns`](#vatreturns) – إقرارات ضريبة القيمة المضافة
61. [`ManualEntries`](#manualentries) – القيود اليدوية
62. [`ManualEntryLines`](#manualentrylines) – أسطر القيود اليدوية
63. [`TransactionTypes`](#transactiontypes) – أنواع حركات المخزون
64. [`InventoryTransactions`](#inventorytransactions) – حركة المخزون
65. [`StockCounts`](#stockcounts) – جلسات الجرد
66. [`StockCountDetails`](#stockcountdetails) – تفاصيل الجرد
67. [`AuditLog`](#auditlog) – سجل العمليات
68. [`AuditChanges`](#auditchanges) – تفاصيل سجل التدقيق
69. [`LabelSettings`](#labelsettings) – إعدادات ملصقات الباركود

## Settings

**إعدادات المحل** – سجل واحد فقط يحتوي بيانات المحل الضريبية وإعدادات التشغيل.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **SettingID** 🔑 | Number (Long) |  | ✔ | `1` | `=1` |  | رقم الإعداد |
| 2 | StoreName | Short Text | 150 | ✔ |  |  |  | اسم المحل |
| 3 | StoreNameEn | Short Text | 150 |  |  |  |  | اسم المحل بالإنجليزية |
| 4 | VATNumber | Short Text | 15 |  |  | `Is Null Or Like "3#############3"` |  | الرقم الضريبي |
| 5 | CRNumber | Short Text | 20 |  |  |  |  | السجل التجاري |
| 6 | BuildingNo | Short Text | 10 |  |  |  |  | رقم المبنى |
| 7 | StreetName | Short Text | 100 |  |  |  |  | الشارع |
| 8 | District | Short Text | 100 |  |  |  |  | الحي |
| 9 | City | Short Text | 50 |  |  |  |  | المدينة |
| 10 | PostalCode | Short Text | 10 |  |  |  |  | الرمز البريدي |
| 11 | AdditionalNo | Short Text | 10 |  |  |  |  | الرقم الإضافي |
| 12 | CountryCode | Short Text | 2 | ✔ | `"SA"` |  |  | رمز الدولة |
| 13 | Phone | Short Text | 20 |  |  |  |  | الهاتف |
| 14 | Email | Short Text | 100 |  |  |  |  | البريد الإلكتروني |
| 15 | VATRate | Currency (نسبة) |  | ✔ | `0.15` | `>=0 And <1` |  | نسبة الضريبة |
| 16 | PricesIncludeVAT | Yes/No |  |  | `True` |  |  | الأسعار شاملة الضريبة |
| 17 | AllowNegativeStock | Yes/No |  |  | `False` |  |  | السماح بالبيع بالسالب |
| 18 | CurrencyCode | Short Text | 3 | ✔ | `"SAR"` |  |  | العملة |
| 19 | DefaultCustomerID | Number (Long) |  | ✔ | `1` |  | `Customers.CustomerID` | العميل الافتراضي |
| 20 | ZatcaPhase | Number (Byte) |  | ✔ | `1` | `In (1,2)` |  | مرحلة فاتورة |
| 21 | ZatcaEnvironment | Short Text | 20 |  |  |  |  | بيئة الربط مع الهيئة |
| 22 | LastInvoiceHash | Short Text | 255 |  |  |  |  | بصمة آخر مستند – تبدأ بالقيمة الافتراضية التي تحددها الهيئة لأول فاتورة |
| 23 | BackupFolder | Short Text | 255 |  |  |  |  | مجلد النسخ الاحتياطي |
| 24 | BackupKeepCount | Number (Integer) |  | ✔ | `30` | `>=1` |  | عدد النسخ المحتفظ بها |
| 25 | SlowMovingDays | Number (Integer) |  | ✔ | `90` | `>=1` |  | أيام عدم الحركة |
| 26 | ReceiptFooter | Short Text | 255 |  |  |  |  | تذييل الفاتورة |
| 27 | LogoPath | Short Text | 255 |  |  |  |  | مسار الشعار |
| 28 | UpdatedAt | Date/Time |  |  |  |  |  | آخر تعديل |
| 29 | POSMode | Short Text | 10 | ✔ | `"RETAIL"` | `In ("RETAIL","RESTAURANT","CAFE")` |  | شاشة البيع – RETAIL = المحلات (باركود)، RESTAURANT = المطاعم (لمس)، CAFE = الكافيهات (لمس) |
| 30 | ImagesFolder | Short Text | 255 |  |  |  |  | مجلد صور المنتجات – المسارات النسبية للصور تُقرأ منه؛ فارغ = مجلد Images بجانب ملف البيانات |
| 31 | InvoicePrintMode | Short Text | 10 | ✔ | `"PREVIEW"` | `In ("DIRECT","PREVIEW","NONE")` |  | الطباعة عند حفظ الفاتورة – DIRECT = طباعة مباشرة بدون معاينة، PREVIEW = عرض المعاينة، NONE = بدون طباعة |
| 32 | AllowAdminCompanyName | Yes/No |  |  | `False` |  |  | السماح لمدير النظام بتغيير اسم المحل |
| 33 | ClosedThrough | Date/Time |  |  |  |  |  | الفترة مقفلة حتى (لا يُضاف ولا يُعدَّل مستند بتاريخ حتى هذا اليوم) |
| 34 | DefaultBankID | Number (Long) |  |  |  |  | `Banks.BankID` | البنك الافتراضي للتحويلات البنكية – المبالغ المدفوعة أو المستلمة بطريقة «تحويل بنكي» تُقيَّد في حساب هذا البنك |
| 35 | GosiEmployeeRate | Currency (نسبة) |  | ✔ | `0.0975` | `>=0 And <1` |  | التأمينات: حصة الموظف السعودي |
| 36 | GosiEmployerRate | Currency (نسبة) |  | ✔ | `0.1175` | `>=0 And <1` |  | التأمينات: حصة المنشأة عن السعودي |
| 37 | GosiNonSaudiRate | Currency (نسبة) |  | ✔ | `0.02` | `>=0 And <1` |  | التأمينات: حصة المنشأة عن غير السعودي (الأخطار المهنية) |
| 38 | GosiMaxWage | Currency |  | ✔ | `45000` | `>=0` |  | التأمينات: الحد الأعلى للأجر الخاضع |
| 39 | CreditBlockDays | Number (Integer) |  |  | `0` | `>=0` |  | إيقاف البيع الآجل لعميل متأخر أكثر من (يوم) |

- المفتاح الأساسي: `SettingID`
- بيانات أساسية: 1 سجل

## Sequences

**عدّادات الترقيم** – يولّد أرقامًا تسلسلية بدون فجوات للفواتير والسندات (يُزاد داخل نفس معاملة الحفظ).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **SequenceName** 🔑 | Short Text | 30 | ✔ |  |  |  | اسم العدّاد |
| 2 | Prefix | Short Text | 10 |  |  |  |  | البادئة |
| 3 | NextValue | Number (Long) |  | ✔ | `1` | `>=1` |  | الرقم التالي |
| 4 | PadLength | Number (Byte) |  | ✔ | `6` | `Between 0 And 12` |  | عدد الخانات |
| 5 | Description | Short Text | 100 |  |  |  |  | الوصف |

- المفتاح الأساسي: `SequenceName`
- بيانات أساسية: 25 سجل

## Roles

**الأدوار** – مستويات الصلاحيات: مدير النظام، مدير، كاشير.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **RoleID** 🔑 | Number (Long) |  | ✔ |  |  |  | رقم الدور |
| 2 | RoleCode | Short Text | 20 | ✔ |  |  |  | رمز الدور |
| 3 | RoleName | Short Text | 50 | ✔ |  |  |  | اسم الدور |
| 4 | Description | Short Text | 255 |  |  |  |  | الوصف |

- المفتاح الأساسي: `RoleID`
- فهرس فريد: `RoleCode`
- بيانات أساسية: 3 سجل

## Permissions

**الصلاحيات** – قائمة الصلاحيات التي يمكن منحها للأدوار.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **PermissionKey** 🔑 | Short Text | 50 | ✔ |  |  |  | رمز الصلاحية |
| 2 | PermissionName | Short Text | 100 | ✔ |  |  |  | اسم الصلاحية |
| 3 | ModuleName | Short Text | 50 |  |  |  |  | القسم |
| 4 | SortOrder | Number (Integer) |  | ✔ | `0` |  |  | الترتيب |

- المفتاح الأساسي: `PermissionKey`
- بيانات أساسية: 36 سجل

## RolePermissions

**صلاحيات الأدوار** – ربط كل دور بالصلاحيات الممنوحة له (علاقة متعدد لمتعدد).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **RoleID** 🔑 | Number (Long) |  | ✔ |  |  | `Roles.RoleID` | الدور |
| 2 | **PermissionKey** 🔑 | Short Text | 50 | ✔ |  |  | `Permissions.PermissionKey` | الصلاحية |

- المفتاح الأساسي: `RoleID, PermissionKey`
- بيانات أساسية: 71 سجل

## Employees

**الموظفون والمستخدمون** – كل موظف هو مستخدم للنظام؛ لا يُحذف بل يُعطَّل للحفاظ على سجل عملياته.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **EmployeeID** 🔑 | AutoNumber |  |  |  |  |  | رقم الموظف |
| 2 | EmployeeName | Short Text | 100 | ✔ |  |  |  | اسم الموظف |
| 3 | JobTitle | Short Text | 50 |  |  |  |  | المسمى الوظيفي |
| 4 | Mobile | Short Text | 20 |  |  |  |  | الجوال |
| 5 | Username | Short Text | 30 | ✔ |  |  |  | اسم المستخدم |
| 6 | PasswordHash | Short Text | 64 |  |  |  |  | بصمة كلمة المرور – SHA-256 بصيغة hex؛ كلمة المرور نفسها لا تُحفظ أبدًا |
| 7 | PasswordSalt | Short Text | 32 |  |  |  |  | ملح التشفير |
| 8 | RoleID | Number (Long) |  | ✔ |  |  | `Roles.RoleID` | الدور |
| 9 | MaxDiscountPercent | Currency (نسبة) |  | ✔ | `0` | `>=0 And <=1` |  | أقصى نسبة خصم |
| 10 | MustChangePassword | Yes/No |  |  | `True` |  |  | يجب تغيير كلمة المرور |
| 11 | FailedLoginCount | Number (Integer) |  | ✔ | `0` |  |  | محاولات الدخول الفاشلة |
| 12 | LockedUntil | Date/Time |  |  |  |  |  | مقفل حتى |
| 13 | LastLoginAt | Date/Time |  |  |  |  |  | آخر دخول |
| 14 | IsActive | Yes/No |  |  | `True` |  |  | نشط |
| 15 | Notes | Long Text |  |  |  |  |  | ملاحظات |
| 16 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 17 | CashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | صندوق النقدية – تدخل فيه نقدية مبيعاته وسنداته؛ فارغ = أول صندوق كاشير نشط |
| 18 | IsDeveloper | Yes/No |  |  | `False` |  |  | المبرمج |
| 19 | CustomScreens | Yes/No |  |  | `False` |  |  | صلاحيات شاشات خاصة |
| 20 | OnPayroll | Yes/No |  |  | `False` |  |  | في مسير الرواتب |
| 21 | BasicSalary | Currency |  | ✔ | `0` | `>=0` |  | الراتب الأساسي |
| 22 | HousingAllowance | Currency |  | ✔ | `0` | `>=0` |  | بدل السكن |
| 23 | TransportAllowance | Currency |  | ✔ | `0` | `>=0` |  | بدل النقل |
| 24 | OtherAllowance | Currency |  | ✔ | `0` | `>=0` |  | بدلات أخرى |
| 25 | IsSaudi | Yes/No |  |  | `False` |  |  | سعودي (التأمينات بحصتي الموظف والمنشأة) |
| 26 | AdvanceInstallment | Currency |  | ✔ | `0` | `>=0` |  | قسط السلفة الشهري – 0 = يُخصم كل الرصيد في أول مسير |
| 27 | NationalID | Short Text | 15 |  |  |  |  | رقم الهوية / الإقامة |
| 28 | IBAN | Short Text | 34 |  |  |  |  | آيبان الموظف |
| 29 | HireDate | Date/Time (تاريخ) |  |  |  |  |  | تاريخ التعيين |
| 30 | CostCenterID | Number (Long) |  |  |  |  | `CostCenters.CostCenterID` | مركز التكلفة – مبيعاته ومسير راتبه على هذا المركز |

- المفتاح الأساسي: `EmployeeID`
- فهرس فريد: `Username`
- بيانات أساسية: 1 سجل

## Screens

**الشاشات** – كل شاشة في البرنامج، وما ينطبق عليها من إضافة وتعديل وحذف، وصلاحية الدور التي تفتحها.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **ScreenName** 🔑 | Short Text | 64 | ✔ |  |  |  | اسم الشاشة في Access |
| 2 | ScreenTitle | Short Text | 100 | ✔ |  |  |  | الشاشة |
| 3 | ModuleName | Short Text | 50 |  |  |  |  | القسم |
| 4 | SortOrder | Number (Integer) |  | ✔ | `0` |  |  | الترتيب |
| 5 | PermissionKey | Short Text | 50 |  |  |  |  | صلاحية الدور – فارغ = متاحة لكل المستخدمين؛ تُستخدم للمستخدم الذي ليست له صلاحيات شاشات خاصة |
| 6 | HasAdd | Yes/No |  |  | `False` |  |  | فيها إضافة / حفظ مستند |
| 7 | HasEdit | Yes/No |  |  | `False` |  |  | فيها تعديل |
| 8 | HasDelete | Yes/No |  |  | `False` |  |  | فيها حذف |

- المفتاح الأساسي: `ScreenName`
- بيانات أساسية: 58 سجل

## UserScreens

**صلاحيات الشاشات للمستخدم** – للمستخدم الذي فُعّلت له «صلاحيات شاشات خاصة»: الشاشات التي يفتحها، والإضافة والتعديل والحذف في كل شاشة.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **EmployeeID** 🔑 | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | المستخدم |
| 2 | **ScreenName** 🔑 | Short Text | 64 | ✔ |  |  | `Screens.ScreenName` | الشاشة |
| 3 | CanOpen | Yes/No |  |  | `False` |  |  | فتح |
| 4 | CanAdd | Yes/No |  |  | `False` |  |  | إضافة |
| 5 | CanEdit | Yes/No |  |  | `False` |  |  | تعديل |
| 6 | CanDelete | Yes/No |  |  | `False` |  |  | حذف |

- المفتاح الأساسي: `EmployeeID, ScreenName`

## Activations

**تفعيل البرنامج** – الأجهزة المفعّل عليها البرنامج: رقم الجهاز وكود التفعيل الصادر من المبرمج.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **ActivationID** 🔑 | AutoNumber |  |  |  |  |  | رقم التفعيل |
| 2 | MachineID | Short Text | 24 | ✔ |  |  |  | رقم الجهاز – بصمة لوحة الأم والمعالج وقرص النظام (modActivation.MachineID) |
| 3 | ActivationCode | Short Text | 30 | ✔ |  |  |  | كود التفعيل |
| 4 | ComputerName | Short Text | 64 |  |  |  |  | اسم الجهاز |
| 5 | ActivatedAt | Date/Time |  |  | `Now()` |  |  | تاريخ التفعيل |
| 6 | EmployeeID | Number (Long) |  |  |  |  | `Employees.EmployeeID` | فعّله |

- المفتاح الأساسي: `ActivationID`
- فهرس فريد: `MachineID`

## Categories

**التصنيفات** – تصنيفات المنتجات (إلكترونيات، مواد غذائية، ...).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **CategoryID** 🔑 | AutoNumber |  |  |  |  |  | رقم التصنيف |
| 2 | CategoryName | Short Text | 100 | ✔ |  |  |  | اسم التصنيف |
| 3 | Description | Short Text | 255 |  |  |  |  | الوصف |
| 4 | IsActive | Yes/No |  |  | `True` |  |  | نشط |
| 5 | ImagePath | Short Text | 255 |  |  |  |  | صورة التصنيف |
| 6 | TileColor | Short Text | 10 | ✔ | `"BLUE"` | `In ("BLUE","GREEN","ORANGE","PURPLE","RED","INDIGO","TEAL","PINK","BROWN","GREY")` |  | لون الزر |
| 7 | SortOrder | Number (Integer) |  | ✔ | `0` |  |  | ترتيب العرض |
| 8 | IsAddOn | Yes/No |  |  | `False` |  |  | فئة إضافات |

- المفتاح الأساسي: `CategoryID`
- فهرس فريد: `CategoryName`
- بيانات أساسية: 1 سجل

## Units

**وحدات القياس** – وحدات البيع (حبة، كرتون، كيلو، ...) مع رمزها في فاتورة الهيئة.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **UnitID** 🔑 | AutoNumber |  |  |  |  |  | رقم الوحدة |
| 2 | UnitName | Short Text | 30 | ✔ |  |  |  | اسم الوحدة |
| 3 | ZatcaUnitCode | Short Text | 10 |  |  |  |  | رمز الوحدة (UN/ECE) |
| 4 | IsActive | Yes/No |  |  | `True` |  |  | نشط |

- المفتاح الأساسي: `UnitID`
- فهرس فريد: `UnitName`
- بيانات أساسية: 8 سجل

## PaymentMethods

**طرق الدفع** – طرق دفع المبالغ المسددة مع رمزها في فاتورة الهيئة (UNTDID 4461).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **PaymentMethodID** 🔑 | Number (Long) |  | ✔ |  |  |  | رقم الطريقة |
| 2 | MethodName | Short Text | 50 | ✔ |  |  |  | طريقة الدفع |
| 3 | MethodNameEn | Short Text | 50 |  |  |  |  | الاسم بالإنجليزية – يظهر في الواجهة الإنجليزية |
| 4 | ZatcaCode | Short Text | 5 |  |  |  |  | رمز الهيئة |
| 5 | SortOrder | Number (Integer) |  | ✔ | `0` |  |  | الترتيب |
| 6 | IsActive | Yes/No |  |  | `True` |  |  | نشط |

- المفتاح الأساسي: `PaymentMethodID`
- فهرس فريد: `MethodName`
- بيانات أساسية: 4 سجل

## Currencies

**العملات** – عملات التعامل. عملة البرنامج هي Settings.CurrencyCode (الريال)، وكل المبالغ تُحفظ بها؛ المستند بعملة أخرى يحفظ عملته ومعامله ومبلغه بها.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **CurrencyCode** 🔑 | Short Text | 3 | ✔ |  | `Len([CurrencyCode])=3` |  | رمز العملة |
| 2 | CurrencyName | Short Text | 50 | ✔ |  |  |  | اسم العملة |
| 3 | CurrencyNameEn | Short Text | 50 |  |  |  |  | اسم العملة بالإنجليزية |
| 4 | Symbol | Short Text | 10 |  |  |  |  | الرمز المختصر |
| 5 | DecimalPlaces | Number (Byte) |  | ✔ | `2` | `Between 0 And 3` |  | عدد الخانات العشرية |
| 6 | SortOrder | Number (Integer) |  | ✔ | `0` |  |  | الترتيب |
| 7 | IsActive | Yes/No |  |  | `True` |  |  | نشط |

- المفتاح الأساسي: `CurrencyCode`
- فهرس فريد: `CurrencyName`
- بيانات أساسية: 11 سجل

## CurrencyRates

**أسعار العملات** – معامل كل عملة في تاريخ: قيمة وحدة واحدة منها بعملة البرنامج. المستند يأخذ آخر سعر في تاريخه أو قبله.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **CurrencyRateID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | CurrencyCode | Short Text | 3 | ✔ |  |  | `Currencies.CurrencyCode` | العملة |
| 3 | RateDate | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | التاريخ |
| 4 | Rate | Currency (نسبة) |  | ✔ | `1` | `>0` |  | المعامل |
| 5 | Notes | Short Text | 100 |  |  |  |  | ملاحظات |
| 6 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `CurrencyRateID`
- فهرس فريد: `CurrencyCode, RateDate`
- بيانات أساسية: 5 سجل

## CashBoxes

**الخزينة والصناديق** – الخزينة الرئيسية وصناديق الكاشير. الرصيد لا يُخزَّن: يُحسب من الحركات (qryCashMovements).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **CashBoxID** 🔑 | AutoNumber |  |  |  |  |  | رقم الصندوق |
| 2 | BoxName | Short Text | 50 | ✔ |  |  |  | اسم الصندوق |
| 3 | BoxType | Short Text | 10 | ✔ | `"CASHIER"` | `In ("MAIN","CASHIER")` |  | النوع |
| 4 | OpeningBalance | Currency |  | ✔ | `0` | `>=0` |  | الرصيد الافتتاحي |
| 5 | OpeningDate | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | تاريخ الرصيد الافتتاحي |
| 6 | IsActive | Yes/No |  |  | `True` |  |  | نشط |
| 7 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 8 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `CashBoxID`
- فهرس فريد: `BoxName`
- بيانات أساسية: 2 سجل

## Suppliers

**الموردون** – بيانات الموردين. الرصيد الحالي قيمة مساعدة تُحدَّث بالكود ويمكن إعادة احتسابها من الحركات.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **SupplierID** 🔑 | AutoNumber |  |  |  |  |  | رقم المورد |
| 2 | SupplierName | Short Text | 150 | ✔ |  |  |  | اسم المورد |
| 3 | ContactPerson | Short Text | 100 |  |  |  |  | الشخص المسؤول |
| 4 | Mobile | Short Text | 20 |  |  |  |  | الجوال |
| 5 | Phone | Short Text | 20 |  |  |  |  | الهاتف |
| 6 | Email | Short Text | 100 |  |  |  |  | البريد الإلكتروني |
| 7 | VATNumber | Short Text | 15 |  |  | `Is Null Or Like "3#############3"` |  | الرقم الضريبي |
| 8 | CRNumber | Short Text | 20 |  |  |  |  | السجل التجاري |
| 9 | Address | Short Text | 255 |  |  |  |  | العنوان |
| 10 | City | Short Text | 50 |  |  |  |  | المدينة |
| 11 | OpeningBalance | Currency |  | ✔ | `0` |  |  | الرصيد الافتتاحي – موجب = المحل مدين للمورد |
| 12 | CurrentBalance | Currency |  | ✔ | `0` |  |  | الرصيد الحالي – قيمة مساعدة؛ المرجع هو SupplierBalanceQuery |
| 13 | PaymentTermsDays | Number (Integer) |  |  | `30` | `>=0` |  | مدة السداد (يوم) |
| 14 | IsActive | Yes/No |  |  | `True` |  |  | نشط |
| 15 | Notes | Long Text |  |  |  |  |  | ملاحظات |
| 16 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 17 | CurrencyCode | Short Text | 3 |  | `"SAR"` |  | `Currencies.CurrencyCode` | عملة التعامل – تُقترح في فواتيره وسنداته |

- المفتاح الأساسي: `SupplierID`
- فهرس عادي: `SupplierName`
- فهرس عادي: `Mobile`

## Customers

**العملاء** – بيانات العملاء. السجل رقم 1 هو "عميل نقدي" الافتراضي ولا يُحذف.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **CustomerID** 🔑 | AutoNumber |  |  |  |  |  | رقم العميل |
| 2 | CustomerName | Short Text | 150 | ✔ |  |  |  | اسم العميل |
| 3 | Mobile | Short Text | 20 |  |  |  |  | الجوال |
| 4 | Phone | Short Text | 20 |  |  |  |  | الهاتف |
| 5 | Email | Short Text | 100 |  |  |  |  | البريد الإلكتروني |
| 6 | VATNumber | Short Text | 15 |  |  | `Is Null Or Like "3#############3"` |  | الرقم الضريبي – إذا وُجد تصدر للعميل فاتورة ضريبية B2B |
| 7 | CRNumber | Short Text | 20 |  |  |  |  | السجل التجاري |
| 8 | BuildingNo | Short Text | 10 |  |  |  |  | رقم المبنى |
| 9 | StreetName | Short Text | 100 |  |  |  |  | الشارع |
| 10 | District | Short Text | 100 |  |  |  |  | الحي |
| 11 | City | Short Text | 50 |  |  |  |  | المدينة |
| 12 | PostalCode | Short Text | 10 |  |  |  |  | الرمز البريدي |
| 13 | Address | Short Text | 255 |  |  |  |  | العنوان |
| 14 | OpeningBalance | Currency |  | ✔ | `0` |  |  | الرصيد الافتتاحي – موجب = العميل مدين للمحل |
| 15 | CurrentBalance | Currency |  | ✔ | `0` |  |  | الرصيد الحالي – قيمة مساعدة؛ المرجع هو CustomerBalanceQuery |
| 16 | AllowCredit | Yes/No |  |  | `True` |  |  | يسمح بالبيع الآجل |
| 17 | CreditLimit | Currency |  | ✔ | `0` | `>=0` |  | حد الائتمان – 0 = بدون حد |
| 18 | PaymentTermsDays | Number (Integer) |  |  | `30` | `>=0` |  | مدة السداد (يوم) |
| 19 | IsSystem | Yes/No |  |  | `False` |  |  | سجل نظام |
| 20 | IsActive | Yes/No |  |  | `True` |  |  | نشط |
| 21 | Notes | Long Text |  |  |  |  |  | ملاحظات |
| 22 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 23 | SalesRepID | Number (Long) |  |  |  |  | `SalesReps.SalesRepID` | المندوب – المندوب المسؤول عن العميل: تُنسب له فواتيره وتحصيلاته |

- المفتاح الأساسي: `CustomerID`
- فهرس عادي: `CustomerName`
- فهرس عادي: `Mobile`
- بيانات أساسية: 1 سجل

## Products

**المنتجات** – الأصناف وأسعارها. CurrentQuantity قيمة مساعدة؛ المرجع هو مجموع InventoryTransactions.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **ProductID** 🔑 | AutoNumber |  |  |  |  |  | رقم المنتج |
| 2 | ProductCode | Short Text | 30 | ✔ |  |  |  | كود المنتج |
| 3 | Barcode | Short Text | 50 |  |  |  |  | الباركود |
| 4 | ProductName | Short Text | 150 | ✔ |  |  |  | اسم المنتج |
| 5 | ProductNameEn | Short Text | 150 |  |  |  |  | الاسم بالإنجليزية |
| 6 | CategoryID | Number (Long) |  | ✔ | `1` |  | `Categories.CategoryID` | التصنيف |
| 7 | UnitID | Number (Long) |  | ✔ | `1` |  | `Units.UnitID` | الوحدة |
| 8 | PurchasePrice | Currency |  | ✔ | `0` | `>=0` |  | آخر سعر شراء – بدون ضريبة |
| 9 | AverageCost | Currency |  | ✔ | `0` | `>=0` |  | متوسط التكلفة – المتوسط المرجّح، يُحدَّث مع كل شراء |
| 10 | SellingPrice | Currency |  | ✔ | `0` | `>=0` |  | سعر البيع – شامل الضريبة إذا كان الإعداد PricesIncludeVAT مفعّلًا |
| 11 | VATCategory | Short Text | 1 | ✔ | `"S"` | `In ("S","Z","E")` |  | الفئة الضريبية |
| 12 | CurrentQuantity | Currency (كمية) |  | ✔ | `0` |  |  | الكمية الحالية |
| 13 | MinimumQuantity | Currency (كمية) |  | ✔ | `0` | `>=0` |  | حد إعادة الطلب |
| 14 | SupplierID | Number (Long) |  |  |  |  | `Suppliers.SupplierID` | المورد الافتراضي |
| 15 | ProductLocation | Short Text | 50 |  |  |  |  | مكان المنتج |
| 16 | Notes | Long Text |  |  |  |  |  | ملاحظات |
| 17 | IsActive | Yes/No |  |  | `True` |  |  | نشط |
| 18 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 19 | UpdatedAt | Date/Time |  |  |  |  |  | آخر تعديل |
| 20 | ImagePath | Short Text | 255 |  |  |  |  | صورة المنتج – لشاشات اللمس؛ مسار كامل أو اسم ملف في مجلد الصور |
| 21 | TrackStock | Yes/No |  |  | `True` |  |  | يتابع المخزون |
| 22 | SizePriceM | Currency |  |  |  | `>=0` |  | سعر الحجم الوسط |
| 23 | SizePriceL | Currency |  |  |  | `>=0` |  | سعر الحجم الكبير |

- المفتاح الأساسي: `ProductID`
- فهرس فريد: `ProductCode`
- فهرس فريد (يتجاهل الفارغ): `Barcode`
- فهرس عادي: `ProductName`

## SalesInvoices

**فواتير المبيعات** – رأس فاتورة البيع. لا تُحذف ولا تُعدَّل بعد الحفظ؛ التصحيح بمرتجع.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **SalesInvoiceID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | InvoiceNumber | Short Text | 20 | ✔ |  |  |  | رقم الفاتورة |
| 3 | InvoiceDate | Date/Time |  | ✔ | `Now()` |  |  | تاريخ ووقت الفاتورة |
| 4 | CustomerID | Number (Long) |  | ✔ | `1` |  | `Customers.CustomerID` | العميل |
| 5 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الكاشير |
| 6 | PaymentType | Short Text | 10 | ✔ | `"CASH"` | `In ("CASH","CREDIT")` |  | نوع البيع |
| 7 | PaymentMethodID | Number (Long) |  |  |  |  | `PaymentMethods.PaymentMethodID` | طريقة الدفع |
| 8 | SubTotal | Currency |  | ✔ | `0` | `>=0` |  | المجموع قبل الخصم – مجموع (الكمية × سعر الوحدة) بدون ضريبة |
| 9 | Discount | Currency |  | ✔ | `0` | `>=0` |  | الخصم – مجموع خصومات الأسطر (خصم الفاتورة يوزَّع على الأسطر) |
| 10 | TaxableAmount | Currency |  | ✔ | `0` | `>=0` |  | الخاضع للضريبة – SubTotal − Discount |
| 11 | Tax | Currency |  | ✔ | `0` | `>=0` |  | ضريبة القيمة المضافة – مجموع ضريبة الأسطر |
| 12 | TotalAmount | Currency |  | ✔ | `0` | `>=0` |  | الإجمالي شامل الضريبة – TaxableAmount + Tax |
| 13 | PaidAmount | Currency |  | ✔ | `0` | `>=0` |  | المدفوع – المبلغ المحتسب من الفاتورة (لا يتجاوز الإجمالي) |
| 14 | RemainingAmount | Currency |  | ✔ | `0` | `>=0` |  | المتبقي – يُضاف إلى رصيد العميل في البيع الآجل |
| 15 | AmountTendered | Currency |  | ✔ | `0` | `>=0` |  | المبلغ المستلم – ما سلّمه العميل نقدًا |
| 16 | ChangeDue | Currency |  | ✔ | `0` | `>=0` |  | الباقي للعميل |
| 17 | DueDate | Date/Time (تاريخ) |  |  |  |  |  | تاريخ الاستحقاق |
| 18 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 19 | InvoiceSubType | Short Text | 10 | ✔ | `"SIMPLIFIED"` | `In ("SIMPLIFIED","STANDARD")` |  | نوع الفاتورة الضريبية – مبسطة للأفراد B2C، ضريبية للمنشآت B2B (للعميل رقم ضريبي) |
| 20 | InvoiceTypeCode | Short Text | 3 | ✔ | `"388"` | `In ("388","381","383")` |  | رمز نوع المستند |
| 21 | InvoiceUUID | Short Text | 36 |  |  |  |  | المعرّف الفريد UUID |
| 22 | ICV | Number (Long) |  |  |  |  |  | عدّاد الفواتير ICV – تسلسل مشترك لكل المستندات المرسلة للهيئة |
| 23 | InvoiceHash | Short Text | 255 |  |  |  |  | بصمة المستند |
| 24 | PreviousInvoiceHash | Short Text | 255 |  |  |  |  | بصمة المستند السابق |
| 25 | QRCodeData | Long Text |  |  |  |  |  | بيانات رمز QR |
| 26 | ZatcaStatus | Short Text | 20 | ✔ | `"NOT_SENT"` | `In ("NOT_SENT","PENDING","REPORTED","CLEARED","WARNING","REJECTED")` |  | حالة الإرسال للهيئة |
| 27 | ZatcaSubmittedAt | Date/Time |  |  |  |  |  | تاريخ الإرسال للهيئة |
| 28 | ZatcaResponse | Long Text |  |  |  |  |  | رد الهيئة |
| 29 | SignedXmlPath | Short Text | 255 |  |  |  |  | مسار ملف XML الموقّع |
| 30 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 31 | OrderType | Short Text | 10 |  |  | `Is Null Or In ("DINE_IN","TAKEAWAY","DELIVERY")` |  | نوع الطلب |
| 32 | TableNo | Short Text | 10 |  |  |  |  | رقم الطاولة |
| 33 | DeliveryPhone | Short Text | 20 |  |  |  |  | جوال التوصيل |
| 34 | DeliveryAddress | Short Text | 255 |  |  |  |  | عنوان التوصيل |
| 35 | OrderName | Short Text | 50 |  |  |  |  | اسم العميل على الطلب |
| 36 | CashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | صندوق النقدية – يُملأ عند الدفع النقدي: المبلغ المدفوع يدخل هذا الصندوق |
| 37 | BankID | Number (Long) |  |  |  |  | `Banks.BankID` | البنك – المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك |
| 38 | CostCenterID | Number (Long) |  |  |  |  | `CostCenters.CostCenterID` | مركز التكلفة – من مركز الكاشير، وإلا المركز الافتراضي |
| 39 | SalesRepID | Number (Long) |  |  |  |  | `SalesReps.SalesRepID` | المندوب – من مندوب العميل، وإلا مندوب المستخدم |

- المفتاح الأساسي: `SalesInvoiceID`
- فهرس فريد: `InvoiceNumber`
- فهرس عادي: `InvoiceDate`
- فهرس فريد (يتجاهل الفارغ): `InvoiceUUID`
- فهرس عادي: `ZatcaStatus`
- قاعدة تحقق على مستوى الجدول: `[PaidAmount]<=[TotalAmount] And [RemainingAmount]=[TotalAmount]-[PaidAmount]` – المدفوع لا يتجاوز الإجمالي، والمتبقي = الإجمالي − المدفوع

## SalesInvoiceDetails

**تفاصيل فواتير المبيعات** – أسطر فاتورة البيع، مع حفظ تكلفة الصنف لحظة البيع لحساب الربح بدقة.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **SalesDetailID** 🔑 | AutoNumber |  |  |  |  |  | رقم السطر الداخلي |
| 2 | SalesInvoiceID | Number (Long) |  | ✔ |  |  | `SalesInvoices.SalesInvoiceID` | الفاتورة |
| 3 | LineNumber | Number (Integer) |  | ✔ |  |  |  | رقم السطر |
| 4 | ProductID | Number (Long) |  | ✔ |  |  | `Products.ProductID` | المنتج |
| 5 | Quantity | Currency (كمية) |  | ✔ |  | `>0` |  | الكمية |
| 6 | UnitPrice | Currency |  | ✔ | `0` | `>=0` |  | سعر الوحدة – بدون ضريبة |
| 7 | Discount | Currency |  | ✔ | `0` | `>=0` |  | الخصم – قيمة الخصم على السطر بدون ضريبة |
| 8 | NetAmount | Currency |  | ✔ | `0` | `>=0` |  | الصافي قبل الضريبة – Quantity × UnitPrice − Discount |
| 9 | VATCategory | Short Text | 1 | ✔ | `"S"` | `In ("S","Z","E")` |  | الفئة الضريبية |
| 10 | VATRate | Currency (نسبة) |  | ✔ | `0.15` | `>=0 And <1` |  | نسبة الضريبة |
| 11 | Tax | Currency |  | ✔ | `0` | `>=0` |  | الضريبة |
| 12 | LineTotal | Currency |  | ✔ | `0` | `>=0` |  | الإجمالي شامل الضريبة – NetAmount + Tax |
| 13 | UnitCost | Currency |  | ✔ | `0` | `>=0` |  | تكلفة الوحدة – AverageCost لحظة البيع |
| 14 | LineNote | Short Text | 100 |  |  |  |  | ملاحظة السطر – الحجم والخيارات (الكافيه) |

- المفتاح الأساسي: `SalesDetailID`
- فهرس فريد: `SalesInvoiceID, LineNumber`

## SalesReturns

**مرتجعات المبيعات** – إشعار دائن مرتبط بالفاتورة الأصلية؛ يعيد الكمية للمخزون ويعكس المبلغ.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **SalesReturnID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | ReturnNumber | Short Text | 20 | ✔ |  |  |  | رقم المرتجع |
| 3 | ReturnDate | Date/Time |  | ✔ | `Now()` |  |  | تاريخ المرتجع |
| 4 | SalesInvoiceID | Number (Long) |  | ✔ |  |  | `SalesInvoices.SalesInvoiceID` | الفاتورة الأصلية |
| 5 | CustomerID | Number (Long) |  | ✔ |  |  | `Customers.CustomerID` | العميل |
| 6 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 7 | Reason | Short Text | 255 | ✔ |  |  |  | سبب الإرجاع – إلزامي في الإشعار الدائن حسب متطلبات الهيئة |
| 8 | RefundType | Short Text | 10 | ✔ | `"CASH"` | `In ("CASH","CREDIT")` |  | طريقة رد المبلغ |
| 9 | PaymentMethodID | Number (Long) |  |  |  |  | `PaymentMethods.PaymentMethodID` | طريقة الرد |
| 10 | SubTotal | Currency |  | ✔ | `0` | `>=0` |  | المجموع قبل الخصم – مجموع (الكمية × سعر الوحدة) بدون ضريبة |
| 11 | Discount | Currency |  | ✔ | `0` | `>=0` |  | الخصم – مجموع خصومات الأسطر (خصم الفاتورة يوزَّع على الأسطر) |
| 12 | TaxableAmount | Currency |  | ✔ | `0` | `>=0` |  | الخاضع للضريبة – SubTotal − Discount |
| 13 | Tax | Currency |  | ✔ | `0` | `>=0` |  | ضريبة القيمة المضافة – مجموع ضريبة الأسطر |
| 14 | TotalAmount | Currency |  | ✔ | `0` | `>=0` |  | الإجمالي شامل الضريبة – TaxableAmount + Tax |
| 15 | RefundedAmount | Currency |  | ✔ | `0` | `>=0` |  | المبلغ المردود نقدًا |
| 16 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 17 | InvoiceSubType | Short Text | 10 | ✔ | `"SIMPLIFIED"` | `In ("SIMPLIFIED","STANDARD")` |  | نوع الفاتورة الضريبية – مبسطة للأفراد B2C، ضريبية للمنشآت B2B (للعميل رقم ضريبي) |
| 18 | InvoiceTypeCode | Short Text | 3 | ✔ | `"381"` | `In ("388","381","383")` |  | رمز نوع المستند |
| 19 | InvoiceUUID | Short Text | 36 |  |  |  |  | المعرّف الفريد UUID |
| 20 | ICV | Number (Long) |  |  |  |  |  | عدّاد الفواتير ICV – تسلسل مشترك لكل المستندات المرسلة للهيئة |
| 21 | InvoiceHash | Short Text | 255 |  |  |  |  | بصمة المستند |
| 22 | PreviousInvoiceHash | Short Text | 255 |  |  |  |  | بصمة المستند السابق |
| 23 | QRCodeData | Long Text |  |  |  |  |  | بيانات رمز QR |
| 24 | ZatcaStatus | Short Text | 20 | ✔ | `"NOT_SENT"` | `In ("NOT_SENT","PENDING","REPORTED","CLEARED","WARNING","REJECTED")` |  | حالة الإرسال للهيئة |
| 25 | ZatcaSubmittedAt | Date/Time |  |  |  |  |  | تاريخ الإرسال للهيئة |
| 26 | ZatcaResponse | Long Text |  |  |  |  |  | رد الهيئة |
| 27 | SignedXmlPath | Short Text | 255 |  |  |  |  | مسار ملف XML الموقّع |
| 28 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 29 | CashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | صندوق النقدية – الرد النقدي يخرج من هذا الصندوق |
| 30 | BankID | Number (Long) |  |  |  |  | `Banks.BankID` | البنك – المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك |
| 31 | CostCenterID | Number (Long) |  |  |  |  | `CostCenters.CostCenterID` | مركز التكلفة – مركز الفاتورة الأصلية |
| 32 | SalesRepID | Number (Long) |  |  |  |  | `SalesReps.SalesRepID` | المندوب – مندوب الفاتورة الأصلية |

- المفتاح الأساسي: `SalesReturnID`
- فهرس فريد: `ReturnNumber`
- فهرس عادي: `ReturnDate`
- فهرس فريد (يتجاهل الفارغ): `InvoiceUUID`
- قاعدة تحقق على مستوى الجدول: `[RefundedAmount]<=[TotalAmount]` – المبلغ المردود لا يتجاوز قيمة المرتجع

## SalesReturnDetails

**تفاصيل مرتجعات المبيعات** – الأسطر المرتجعة، كل سطر يشير إلى سطر الفاتورة الأصلي لمنع إرجاع أكثر مما بيع.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **ReturnDetailID** 🔑 | AutoNumber |  |  |  |  |  | رقم السطر الداخلي |
| 2 | SalesReturnID | Number (Long) |  | ✔ |  |  | `SalesReturns.SalesReturnID` | المرتجع |
| 3 | SalesDetailID | Number (Long) |  | ✔ |  |  | `SalesInvoiceDetails.SalesDetailID` | سطر الفاتورة الأصلي |
| 4 | ProductID | Number (Long) |  | ✔ |  |  | `Products.ProductID` | المنتج |
| 5 | Quantity | Currency (كمية) |  | ✔ |  | `>0` |  | الكمية المرتجعة |
| 6 | UnitPrice | Currency |  | ✔ | `0` | `>=0` |  | سعر الوحدة |
| 7 | Discount | Currency |  | ✔ | `0` | `>=0` |  | الخصم |
| 8 | NetAmount | Currency |  | ✔ | `0` | `>=0` |  | الصافي قبل الضريبة |
| 9 | VATCategory | Short Text | 1 | ✔ | `"S"` | `In ("S","Z","E")` |  | الفئة الضريبية |
| 10 | VATRate | Currency (نسبة) |  | ✔ | `0.15` | `>=0 And <1` |  | نسبة الضريبة |
| 11 | Tax | Currency |  | ✔ | `0` | `>=0` |  | الضريبة |
| 12 | LineTotal | Currency |  | ✔ | `0` | `>=0` |  | الإجمالي شامل الضريبة |
| 13 | UnitCost | Currency |  | ✔ | `0` | `>=0` |  | تكلفة الوحدة – نفس تكلفة سطر البيع الأصلي |
| 14 | ReturnToStock | Yes/No |  |  | `True` |  |  | يعاد للمخزون |

- المفتاح الأساسي: `ReturnDetailID`

## PurchaseInvoices

**فواتير المشتريات** – رأس فاتورة الشراء من المورد.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **PurchaseInvoiceID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | InvoiceNumber | Short Text | 20 | ✔ |  |  |  | رقم الفاتورة الداخلي |
| 3 | SupplierInvoiceNo | Short Text | 30 |  |  |  |  | رقم فاتورة المورد |
| 4 | InvoiceDate | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الفاتورة |
| 5 | SupplierID | Number (Long) |  | ✔ |  |  | `Suppliers.SupplierID` | المورد |
| 6 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 7 | PaymentType | Short Text | 10 | ✔ | `"CASH"` | `In ("CASH","CREDIT")` |  | نوع الشراء |
| 8 | PaymentMethodID | Number (Long) |  |  |  |  | `PaymentMethods.PaymentMethodID` | طريقة الدفع |
| 9 | SubTotal | Currency |  | ✔ | `0` | `>=0` |  | المجموع قبل الخصم – مجموع (الكمية × سعر الوحدة) بدون ضريبة |
| 10 | Discount | Currency |  | ✔ | `0` | `>=0` |  | الخصم – مجموع خصومات الأسطر (خصم الفاتورة يوزَّع على الأسطر) |
| 11 | TaxableAmount | Currency |  | ✔ | `0` | `>=0` |  | الخاضع للضريبة – SubTotal − Discount |
| 12 | Tax | Currency |  | ✔ | `0` | `>=0` |  | ضريبة القيمة المضافة – مجموع ضريبة الأسطر |
| 13 | TotalAmount | Currency |  | ✔ | `0` | `>=0` |  | الإجمالي شامل الضريبة – TaxableAmount + Tax |
| 14 | PaidAmount | Currency |  | ✔ | `0` | `>=0` |  | المدفوع |
| 15 | RemainingAmount | Currency |  | ✔ | `0` | `>=0` |  | المتبقي – يُضاف إلى رصيد المورد |
| 16 | DueDate | Date/Time (تاريخ) |  |  |  |  |  | تاريخ الاستحقاق |
| 17 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 18 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 19 | CashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | صندوق النقدية – المدفوع نقدًا يخرج من هذا الصندوق |
| 20 | BankID | Number (Long) |  |  |  |  | `Banks.BankID` | البنك – المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك |
| 21 | CurrencyCode | Short Text | 3 |  | `"SAR"` |  | `Currencies.CurrencyCode` | العملة – فارغ = عملة البرنامج (Settings.CurrencyCode) |
| 22 | ExchangeRate | Currency (نسبة) |  | ✔ | `1` | `>0` |  | معامل التحويل – قيمة وحدة واحدة من العملة بعملة البرنامج |
| 23 | ForeignAmount | Currency |  | ✔ | `0` | `>=0` |  | المبلغ بالعملة – الإجمالي بعملة المستند (0 للمستندات القديمة) |

- المفتاح الأساسي: `PurchaseInvoiceID`
- فهرس فريد: `InvoiceNumber`
- فهرس عادي: `InvoiceDate`
- فهرس عادي: `SupplierInvoiceNo`
- قاعدة تحقق على مستوى الجدول: `[PaidAmount]<=[TotalAmount] And [RemainingAmount]=[TotalAmount]-[PaidAmount]` – المدفوع لا يتجاوز الإجمالي، والمتبقي = الإجمالي − المدفوع

## PurchaseInvoiceDetails

**تفاصيل فواتير المشتريات** – أسطر فاتورة الشراء.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **PurchaseDetailID** 🔑 | AutoNumber |  |  |  |  |  | رقم السطر الداخلي |
| 2 | PurchaseInvoiceID | Number (Long) |  | ✔ |  |  | `PurchaseInvoices.PurchaseInvoiceID` | الفاتورة |
| 3 | LineNumber | Number (Integer) |  | ✔ |  |  |  | رقم السطر |
| 4 | ProductID | Number (Long) |  | ✔ |  |  | `Products.ProductID` | المنتج |
| 5 | Quantity | Currency (كمية) |  | ✔ |  | `>0` |  | الكمية |
| 6 | UnitCost | Currency |  | ✔ | `0` | `>=0` |  | تكلفة الوحدة – بدون ضريبة |
| 7 | Discount | Currency |  | ✔ | `0` | `>=0` |  | الخصم |
| 8 | NetAmount | Currency |  | ✔ | `0` | `>=0` |  | الصافي قبل الضريبة |
| 9 | VATRate | Currency (نسبة) |  | ✔ | `0.15` | `>=0 And <1` |  | نسبة الضريبة |
| 10 | Tax | Currency |  | ✔ | `0` | `>=0` |  | ضريبة المدخلات |
| 11 | LineTotal | Currency |  | ✔ | `0` | `>=0` |  | الإجمالي شامل الضريبة |

- المفتاح الأساسي: `PurchaseDetailID`
- فهرس فريد: `PurchaseInvoiceID, LineNumber`

## PurchaseReturns

**مرتجعات المشتريات** – إرجاع بضاعة للمورد مرتبط بفاتورة الشراء الأصلية.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **PurchaseReturnID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | ReturnNumber | Short Text | 20 | ✔ |  |  |  | رقم المرتجع |
| 3 | ReturnDate | Date/Time |  | ✔ | `Now()` |  |  | تاريخ المرتجع |
| 4 | PurchaseInvoiceID | Number (Long) |  | ✔ |  |  | `PurchaseInvoices.PurchaseInvoiceID` | فاتورة الشراء الأصلية |
| 5 | SupplierID | Number (Long) |  | ✔ |  |  | `Suppliers.SupplierID` | المورد |
| 6 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 7 | Reason | Short Text | 255 | ✔ |  |  |  | سبب الإرجاع |
| 8 | RefundType | Short Text | 10 | ✔ | `"CREDIT"` | `In ("CASH","CREDIT")` |  | طريقة الاسترداد |
| 9 | PaymentMethodID | Number (Long) |  |  |  |  | `PaymentMethods.PaymentMethodID` | طريقة الاسترداد |
| 10 | SubTotal | Currency |  | ✔ | `0` | `>=0` |  | المجموع قبل الخصم – مجموع (الكمية × سعر الوحدة) بدون ضريبة |
| 11 | Discount | Currency |  | ✔ | `0` | `>=0` |  | الخصم – مجموع خصومات الأسطر (خصم الفاتورة يوزَّع على الأسطر) |
| 12 | TaxableAmount | Currency |  | ✔ | `0` | `>=0` |  | الخاضع للضريبة – SubTotal − Discount |
| 13 | Tax | Currency |  | ✔ | `0` | `>=0` |  | ضريبة القيمة المضافة – مجموع ضريبة الأسطر |
| 14 | TotalAmount | Currency |  | ✔ | `0` | `>=0` |  | الإجمالي شامل الضريبة – TaxableAmount + Tax |
| 15 | RefundedAmount | Currency |  | ✔ | `0` | `>=0` |  | المبلغ المسترد نقدًا |
| 16 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 17 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 18 | CashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | صندوق النقدية – الاسترداد النقدي يدخل هذا الصندوق |
| 19 | BankID | Number (Long) |  |  |  |  | `Banks.BankID` | البنك – المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك |
| 20 | CurrencyCode | Short Text | 3 |  | `"SAR"` |  | `Currencies.CurrencyCode` | العملة – فارغ = عملة البرنامج (Settings.CurrencyCode) |
| 21 | ExchangeRate | Currency (نسبة) |  | ✔ | `1` | `>0` |  | معامل التحويل – قيمة وحدة واحدة من العملة بعملة البرنامج |
| 22 | ForeignAmount | Currency |  | ✔ | `0` | `>=0` |  | المبلغ بالعملة – الإجمالي بعملة المستند (0 للمستندات القديمة) |

- المفتاح الأساسي: `PurchaseReturnID`
- فهرس فريد: `ReturnNumber`
- فهرس عادي: `ReturnDate`
- قاعدة تحقق على مستوى الجدول: `[RefundedAmount]<=[TotalAmount]` – المبلغ المسترد لا يتجاوز قيمة المرتجع

## PurchaseReturnDetails

**تفاصيل مرتجعات المشتريات** – الأسطر المرتجعة للمورد.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **PurchaseReturnDetailID** 🔑 | AutoNumber |  |  |  |  |  | رقم السطر الداخلي |
| 2 | PurchaseReturnID | Number (Long) |  | ✔ |  |  | `PurchaseReturns.PurchaseReturnID` | المرتجع |
| 3 | PurchaseDetailID | Number (Long) |  | ✔ |  |  | `PurchaseInvoiceDetails.PurchaseDetailID` | سطر الفاتورة الأصلي |
| 4 | ProductID | Number (Long) |  | ✔ |  |  | `Products.ProductID` | المنتج |
| 5 | Quantity | Currency (كمية) |  | ✔ |  | `>0` |  | الكمية المرتجعة |
| 6 | UnitCost | Currency |  | ✔ | `0` | `>=0` |  | تكلفة الوحدة |
| 7 | Discount | Currency |  | ✔ | `0` | `>=0` |  | الخصم |
| 8 | NetAmount | Currency |  | ✔ | `0` | `>=0` |  | الصافي قبل الضريبة |
| 9 | VATRate | Currency (نسبة) |  | ✔ | `0.15` | `>=0 And <1` |  | نسبة الضريبة |
| 10 | Tax | Currency |  | ✔ | `0` | `>=0` |  | الضريبة |
| 11 | LineTotal | Currency |  | ✔ | `0` | `>=0` |  | الإجمالي شامل الضريبة |

- المفتاح الأساسي: `PurchaseReturnDetailID`

## CustomerPayments

**دفعات العملاء (سندات القبض)** – المبالغ المستلمة من العملاء لسداد أرصدتهم الآجلة.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **PaymentID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | PaymentNumber | Short Text | 20 | ✔ |  |  |  | رقم السند |
| 3 | CustomerID | Number (Long) |  | ✔ |  |  | `Customers.CustomerID` | العميل |
| 4 | PaymentDate | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الدفعة |
| 5 | Amount | Currency |  | ✔ | `0` | `>0` |  | المبلغ |
| 6 | PaymentMethodID | Number (Long) |  | ✔ | `1` |  | `PaymentMethods.PaymentMethodID` | طريقة الدفع |
| 7 | SalesInvoiceID | Number (Long) |  |  |  |  | `SalesInvoices.SalesInvoiceID` | عن فاتورة (اختياري) |
| 8 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 9 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 10 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 11 | CashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | صندوق النقدية – المبلغ النقدي يدخل هذا الصندوق |
| 12 | BankID | Number (Long) |  |  |  |  | `Banks.BankID` | البنك – المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك |
| 13 | CurrencyCode | Short Text | 3 |  | `"SAR"` |  | `Currencies.CurrencyCode` | العملة – فارغ = عملة البرنامج (Settings.CurrencyCode) |
| 14 | ExchangeRate | Currency (نسبة) |  | ✔ | `1` | `>0` |  | معامل التحويل – قيمة وحدة واحدة من العملة بعملة البرنامج |
| 15 | ForeignAmount | Currency |  | ✔ | `0` | `>=0` |  | المبلغ بالعملة – الإجمالي بعملة المستند (0 للمستندات القديمة) |
| 16 | SalesRepID | Number (Long) |  |  |  |  | `SalesReps.SalesRepID` | المندوب – المحصِّل: مندوب العميل، وإلا مندوب المستخدم |

- المفتاح الأساسي: `PaymentID`
- فهرس فريد: `PaymentNumber`
- فهرس عادي: `PaymentDate`

## SupplierPayments

**دفعات الموردين (سندات الصرف)** – المبالغ المدفوعة للموردين لسداد أرصدتهم.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **PaymentID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | PaymentNumber | Short Text | 20 | ✔ |  |  |  | رقم السند |
| 3 | SupplierID | Number (Long) |  | ✔ |  |  | `Suppliers.SupplierID` | المورد |
| 4 | PaymentDate | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الدفعة |
| 5 | Amount | Currency |  | ✔ | `0` | `>0` |  | المبلغ |
| 6 | PaymentMethodID | Number (Long) |  | ✔ | `1` |  | `PaymentMethods.PaymentMethodID` | طريقة الدفع |
| 7 | PurchaseInvoiceID | Number (Long) |  |  |  |  | `PurchaseInvoices.PurchaseInvoiceID` | عن فاتورة (اختياري) |
| 8 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 9 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 10 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 11 | CashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | صندوق النقدية – المبلغ النقدي يخرج من هذا الصندوق |
| 12 | BankID | Number (Long) |  |  |  |  | `Banks.BankID` | البنك – المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك |
| 13 | CurrencyCode | Short Text | 3 |  | `"SAR"` |  | `Currencies.CurrencyCode` | العملة – فارغ = عملة البرنامج (Settings.CurrencyCode) |
| 14 | ExchangeRate | Currency (نسبة) |  | ✔ | `1` | `>0` |  | معامل التحويل – قيمة وحدة واحدة من العملة بعملة البرنامج |
| 15 | ForeignAmount | Currency |  | ✔ | `0` | `>=0` |  | المبلغ بالعملة – الإجمالي بعملة المستند (0 للمستندات القديمة) |

- المفتاح الأساسي: `PaymentID`
- فهرس فريد: `PaymentNumber`
- فهرس عادي: `PaymentDate`

## Banks

**البنوك** – كل حساب بنكي للمحل. حسابه في الدليل 120000 + رقمه تحت «الحسابات البنكية» (1210).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **BankID** 🔑 | AutoNumber |  |  |  |  |  | رقم البنك |
| 2 | BankName | Short Text | 100 | ✔ |  |  |  | اسم البنك / الحساب |
| 3 | AccountNo | Short Text | 30 |  |  |  |  | رقم الحساب |
| 4 | IBAN | Short Text | 34 |  |  |  |  | الآيبان |
| 5 | OpeningBalance | Currency |  | ✔ | `0` |  |  | الرصيد الافتتاحي – رصيد الحساب في البنك عند بدء استخدام البرنامج |
| 6 | OpeningDate | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | تاريخ الرصيد الافتتاحي |
| 7 | IsActive | Yes/No |  |  | `True` |  |  | نشط |
| 8 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 9 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `BankID`
- فهرس فريد: `BankName`

## BankTransactions

**الحركات البنكية** – إيداع نقدية من صندوق، سحب إلى صندوق، تسوية تحصيلات مدى (بعمولتها)، تحويل بين بنكين، وحركات أخرى (عمولات، فوائد، قروض...) بحساب مقابل.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **BankTxID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | TxNumber | Short Text | 20 | ✔ |  |  |  | رقم الحركة |
| 3 | TxDate | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | التاريخ |
| 4 | TxType | Short Text | 12 | ✔ |  | `In ("DEPOSIT","WITHDRAW","SETTLEMENT","TRANSFER","OTHER_IN","OTHER_OUT")` |  | النوع |
| 5 | BankID | Number (Long) |  | ✔ |  |  | `Banks.BankID` | البنك |
| 6 | ToBankID | Number (Long) |  |  |  |  | `Banks.BankID` | إلى بنك |
| 7 | CashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | الصندوق |
| 8 | CounterAccount | Number (Long) |  |  |  |  | `Accounts.AccountCode` | الحساب المقابل |
| 9 | Amount | Currency |  | ✔ | `0` | `>0` |  | المبلغ – في تسوية مدى: إجمالي التحصيلات قبل العمولة |
| 10 | FeeAmount | Currency |  | ✔ | `0` | `>=0` |  | العمولة – تسوية مدى: عمولة البنك بدون ضريبة |
| 11 | FeeVAT | Currency |  | ✔ | `0` | `>=0` |  | ضريبة العمولة – ضريبة مدخلات على العمولة |
| 12 | Reference | Short Text | 40 |  |  |  |  | مرجع البنك |
| 13 | Description | Short Text | 255 |  |  |  |  | البيان |
| 14 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 15 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `BankTxID`
- فهرس فريد: `TxNumber`
- فهرس عادي: `TxDate`
- قاعدة تحقق على مستوى الجدول: `[FeeAmount]+[FeeVAT]<[Amount] Or [TxType]<>"SETTLEMENT"` – العمولة وضريبتها أقل من مبلغ التسوية

## Cheques

**الشيكات الواردة والصادرة** – شيك وارد من عميل (يسدد رصيده ويبقى في «شيكات تحت التحصيل» حتى يُحصَّل في البنك أو يرتد) أو صادر لمورد (يسدد رصيده ويبقى في «أوراق الدفع» حتى يصرفه البنك أو يرتد).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **ChequeID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | ChequeRef | Short Text | 20 | ✔ |  |  |  | رقم القيد الداخلي |
| 3 | Direction | Short Text | 3 | ✔ |  | `In ("IN","OUT")` |  | الاتجاه |
| 4 | CustomerID | Number (Long) |  |  |  |  | `Customers.CustomerID` | العميل (الوارد) |
| 5 | SupplierID | Number (Long) |  |  |  |  | `Suppliers.SupplierID` | المورد (الصادر) |
| 6 | ChequeNo | Short Text | 30 | ✔ |  |  |  | رقم الشيك |
| 7 | DrawerBank | Short Text | 100 |  |  |  |  | بنك الساحب (الوارد) |
| 8 | BankID | Number (Long) |  |  |  |  | `Banks.BankID` | بنكنا – الوارد: البنك الذي حُصِّل فيه؛ الصادر: البنك المسحوب عليه |
| 9 | IssueDate | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | تاريخ الاستلام / الإصدار |
| 10 | DueDate | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | تاريخ الاستحقاق |
| 11 | Amount | Currency |  | ✔ | `0` | `>0` |  | المبلغ |
| 12 | Status | Short Text | 10 | ✔ | `"PENDING"` | `In ("PENDING","COLLECTED","BOUNCED")` |  | الحالة |
| 13 | StatusDate | Date/Time (تاريخ) |  |  |  |  |  | تاريخ التحصيل / الارتداد |
| 14 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 15 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 16 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `ChequeID`
- فهرس فريد: `ChequeRef`
- فهرس عادي: `DueDate`
- فهرس عادي: `ChequeNo`
- قاعدة تحقق على مستوى الجدول: `([Direction]="IN" And [CustomerID] Is Not Null And [SupplierID] Is Null) Or ([Direction]="OUT" And [SupplierID] Is Not Null And [CustomerID] Is Null)` – الشيك الوارد لعميل، والصادر لمورد

## FixedAssets

**الأصول الثابتة** – سجل الأصول: التكلفة وحساب الأصل، والعمر الإنتاجي بالأشهر، والقيمة المتبقية، ومصدر الشراء. يُهلك بالقسط الثابت شهريًا، ويُستبعد عند بيعه أو تلفه.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **AssetID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | AssetCode | Short Text | 20 | ✔ |  |  |  | رقم الأصل |
| 3 | AssetName | Short Text | 150 | ✔ |  |  |  | اسم الأصل |
| 4 | AssetAccount | Number (Long) |  | ✔ |  |  | `Accounts.AccountCode` | حساب الأصل – حساب فرعي من الأصول غير المتداولة (12)، مثل الأثاث أو الأجهزة |
| 5 | CostCenterID | Number (Long) |  |  |  |  | `CostCenters.CostCenterID` | مركز التكلفة – يذهب إليه إهلاك الأصل وربح أو خسارة بيعه |
| 6 | PurchaseDate | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | تاريخ الشراء |
| 7 | Cost | Currency |  | ✔ | `0` | `>0` |  | التكلفة بدون الضريبة |
| 8 | InputVAT | Currency |  | ✔ | `0` | `>=0` |  | ضريبة المدخلات |
| 9 | SalvageValue | Currency |  | ✔ | `0` | `>=0` |  | القيمة المتبقية في آخر العمر |
| 10 | UsefulLifeMonths | Number (Integer) |  | ✔ |  | `>0` |  | العمر الإنتاجي (شهر) |
| 11 | DepStartDate | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | بداية الإهلاك (يُهلك من شهر هذا التاريخ) |
| 12 | SourceType | Short Text | 10 | ✔ | `"BANK"` | `In ("OPENING","BANK","CASHBOX","ACCOUNT")` |  | مصدر الشراء |
| 13 | BankID | Number (Long) |  |  |  |  | `Banks.BankID` | البنك |
| 14 | CashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | الصندوق |
| 15 | CounterAccount | Number (Long) |  |  |  |  | `Accounts.AccountCode` | الحساب الدائن (حساب آخر) |
| 16 | OpeningAccumDep | Currency |  | ✔ | `0` | `>=0` |  | إهلاك سابق (للأصول الموجودة قبل البرنامج) |
| 17 | Status | Short Text | 10 | ✔ | `"ACTIVE"` | `In ("ACTIVE","DISPOSED")` |  | الحالة |
| 18 | DisposalDate | Date/Time (تاريخ) |  |  |  |  |  | تاريخ البيع / الاستبعاد |
| 19 | DisposalProceeds | Currency |  | ✔ | `0` | `>=0` |  | ثمن البيع |
| 20 | DisposalTo | Short Text | 10 |  |  | `Is Null Or In ("NONE","BANK","CASHBOX")` |  | استلام الثمن |
| 21 | DisposalBankID | Number (Long) |  |  |  |  | `Banks.BankID` | بنك استلام الثمن |
| 22 | DisposalCashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | صندوق استلام الثمن |
| 23 | DisposalAccumDep | Currency |  | ✔ | `0` | `>=0` |  | مجمع الإهلاك يوم الاستبعاد |
| 24 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 25 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 26 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `AssetID`
- فهرس فريد: `AssetCode`
- فهرس عادي: `AssetAccount`
- قاعدة تحقق على مستوى الجدول: `[SalvageValue]+[OpeningAccumDep]<=[Cost]` – القيمة المتبقية والإهلاك السابق لا يتجاوزان التكلفة

## DepreciationRuns

**قيود الإهلاك الشهرية** – قيد إهلاك كل شهر: مصروف الإهلاك (5600) مدين ومجمع الإهلاك (1790) دائن، بسطر لكل أصل.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **RunID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | RunNumber | Short Text | 20 | ✔ |  |  |  | رقم القيد |
| 3 | RunMonth | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | الشهر (آخر يوم فيه) |
| 4 | TotalAmount | Currency |  | ✔ | `0` | `>=0` |  | إجمالي الإهلاك |
| 5 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 6 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `RunID`
- فهرس فريد: `RunNumber`
- فهرس فريد: `RunMonth`

## AssetDepreciations

**إهلاك كل أصل في كل شهر** – أسطر قيد الإهلاك الشهري.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **DepreciationID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | RunID | Number (Long) |  | ✔ |  |  | `DepreciationRuns.RunID` | قيد الإهلاك |
| 3 | LineNo | Number (Integer) |  | ✔ |  |  |  | رقم السطر في القيد |
| 4 | AssetID | Number (Long) |  | ✔ |  |  | `FixedAssets.AssetID` | الأصل |
| 5 | Amount | Currency |  | ✔ | `0` | `>0` |  | الإهلاك |

- المفتاح الأساسي: `DepreciationID`
- فهرس فريد: `RunID, AssetID`
- فهرس عادي: `AssetID`

## CostCenters

**مراكز التكلفة والفروع** – الفروع أو الأقسام (التجزئة، المطعم، المقهى...). تُوزَّع عليها الإيرادات والمصروفات في القيود، ومنها قائمة دخل لكل مركز.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **CostCenterID** 🔑 | AutoNumber |  |  |  |  |  | رقم المركز |
| 2 | CenterCode | Short Text | 10 | ✔ |  |  |  | رمز المركز |
| 3 | CenterName | Short Text | 100 | ✔ |  |  |  | اسم المركز |
| 4 | IsDefault | Yes/No |  |  | `False` |  |  | المركز الافتراضي |
| 5 | IsActive | Yes/No |  |  | `True` |  |  | نشط |
| 6 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 7 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `CostCenterID`
- فهرس فريد: `CenterCode`
- فهرس فريد: `CenterName`

## SalesReps

**المندوبين** – مندوبو المبيعات: العملاء المسندون لكل مندوب، ومبيعاته وتحصيلاته وهدفه الشهري وعمولته.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **SalesRepID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | RepCode | Short Text | 20 | ✔ |  |  |  | كود المندوب |
| 3 | RepName | Short Text | 100 | ✔ |  |  |  | اسم المندوب |
| 4 | RepNameEn | Short Text | 100 |  |  |  |  | الاسم بالإنجليزية |
| 5 | Mobile | Short Text | 20 |  |  |  |  | الجوال |
| 6 | EmployeeID | Number (Long) |  |  |  |  | `Employees.EmployeeID` | مستخدم البرنامج – مبيعات هذا المستخدم لعميل بلا مندوب تُنسب لهذا المندوب |
| 7 | Region | Short Text | 50 |  |  |  |  | المنطقة / خط السير |
| 8 | CostCenterID | Number (Long) |  |  |  |  | `CostCenters.CostCenterID` | مركز التكلفة – مركز قيد عمولته |
| 9 | CommissionRate | Currency (نسبة) |  | ✔ | `0` | `>=0 And <1` |  | نسبة العمولة |
| 10 | CommissionBase | Short Text | 10 | ✔ | `"SALES"` | `In ("SALES","COLLECTION")` |  | أساس العمولة |
| 11 | IsActive | Yes/No |  |  | `True` |  |  | نشط |
| 12 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 13 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `SalesRepID`
- فهرس فريد: `RepCode`
- فهرس فريد: `RepName`

## SalesRepTargets

**أهداف المندوبين** – الهدف الشهري لمبيعات كل مندوب (صافي المبيعات بدون الضريبة).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **TargetID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | SalesRepID | Number (Long) |  | ✔ |  |  | `SalesReps.SalesRepID` | المندوب |
| 3 | TargetYear | Number (Integer) |  | ✔ |  | `Between 2000 And 2100` |  | السنة |
| 4 | TargetMonth | Number (Byte) |  | ✔ | `1` | `Between 1 And 12` |  | الشهر |
| 5 | TargetAmount | Currency |  | ✔ | `0` | `>=0` |  | الهدف |

- المفتاح الأساسي: `TargetID`
- فهرس فريد: `SalesRepID, TargetYear, TargetMonth`

## CommissionRuns

**مسيرات العمولات** – عمولات المندوبين لشهر: مسودة تُعدَّل، ثم يُرحَّل قيدها (مصروف العمولات 5530 على عمولات مستحقة 2330)، وتُصرف بسند صرف نقدية من بند «صرف عمولة مندوب».

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **CommissionRunID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | RunNumber | Short Text | 20 | ✔ |  |  |  | رقم المسير |
| 3 | RunMonth | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | الشهر (آخر يوم) |
| 4 | Status | Short Text | 10 | ✔ | `"DRAFT"` | `In ("DRAFT","POSTED")` |  | الحالة |
| 5 | TotalAmount | Currency |  | ✔ | `0` | `>=0` |  | إجمالي العمولات |
| 6 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 7 | PostedAt | Date/Time |  |  |  |  |  | تاريخ الترحيل |
| 8 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 9 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `CommissionRunID`
- فهرس فريد: `RunNumber`
- فهرس فريد: `RunMonth`

## CommissionLines

**أسطر مسير العمولات** – عمولة كل مندوب في الشهر: صافي مبيعاته وتحصيلاته، والأساس والنسبة، والتعديل اليدوي.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **CommissionLineID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | CommissionRunID | Number (Long) |  | ✔ |  |  | `CommissionRuns.CommissionRunID` | المسير |
| 3 | SalesRepID | Number (Long) |  | ✔ |  |  | `SalesReps.SalesRepID` | المندوب |
| 4 | RepName | Short Text | 100 |  |  |  |  | اسم المندوب |
| 5 | CostCenterID | Number (Long) |  |  |  |  | `CostCenters.CostCenterID` | مركز التكلفة |
| 6 | NetSales | Currency |  | ✔ | `0` |  |  | صافي المبيعات |
| 7 | Collections | Currency |  | ✔ | `0` |  |  | التحصيل |
| 8 | CommissionBase | Short Text | 10 | ✔ | `"SALES"` |  |  | الأساس |
| 9 | BaseAmount | Currency |  | ✔ | `0` |  |  | مبلغ الأساس |
| 10 | CommissionRate | Currency (نسبة) |  | ✔ | `0` | `>=0 And <1` |  | النسبة |
| 11 | Adjustment | Currency |  | ✔ | `0` |  |  | تعديل (+/-) |
| 12 | Commission | Currency |  | ✔ | `0` | `>=0` |  | العمولة – الأساس × النسبة + التعديل، ولا تقل عن صفر |
| 13 | Notes | Short Text | 150 |  |  |  |  | ملاحظات |

- المفتاح الأساسي: `CommissionLineID`
- فهرس فريد: `CommissionRunID, SalesRepID`

## Budgets

**الموازنات التقديرية** – موازنة كل سنة: مبلغ شهري لكل حساب إيرادات أو مصروفات، ويمكن لكل مركز تكلفة.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **BudgetID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | BudgetYear | Number (Integer) |  | ✔ |  |  |  | السنة |
| 3 | BudgetName | Short Text | 100 | ✔ |  |  |  | اسم الموازنة |
| 4 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 5 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | أعدّها |
| 6 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `BudgetID`
- فهرس فريد: `BudgetYear`

## BudgetLines

**أسطر الموازنة** – حساب (ومركز تكلفة اختياري) ومبلغه في كل شهر. بلا مركز = الحساب في كل المراكز.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **BudgetLineID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | BudgetID | Number (Long) |  | ✔ |  |  | `Budgets.BudgetID` | الموازنة |
| 3 | AccountCode | Number (Long) |  | ✔ |  |  | `Accounts.AccountCode` | الحساب |
| 4 | CostCenterID | Number (Long) |  |  |  |  | `CostCenters.CostCenterID` | مركز التكلفة – فارغ = كل المراكز |
| 5 | M1 | Currency |  | ✔ | `0` | `>=0` |  | الشهر 1 |
| 6 | M2 | Currency |  | ✔ | `0` | `>=0` |  | الشهر 2 |
| 7 | M3 | Currency |  | ✔ | `0` | `>=0` |  | الشهر 3 |
| 8 | M4 | Currency |  | ✔ | `0` | `>=0` |  | الشهر 4 |
| 9 | M5 | Currency |  | ✔ | `0` | `>=0` |  | الشهر 5 |
| 10 | M6 | Currency |  | ✔ | `0` | `>=0` |  | الشهر 6 |
| 11 | M7 | Currency |  | ✔ | `0` | `>=0` |  | الشهر 7 |
| 12 | M8 | Currency |  | ✔ | `0` | `>=0` |  | الشهر 8 |
| 13 | M9 | Currency |  | ✔ | `0` | `>=0` |  | الشهر 9 |
| 14 | M10 | Currency |  | ✔ | `0` | `>=0` |  | الشهر 10 |
| 15 | M11 | Currency |  | ✔ | `0` | `>=0` |  | الشهر 11 |
| 16 | M12 | Currency |  | ✔ | `0` | `>=0` |  | الشهر 12 |
| 17 | Notes | Short Text | 150 |  |  |  |  | ملاحظات |

- المفتاح الأساسي: `BudgetLineID`
- فهرس عادي: `BudgetID`
- فهرس عادي: `AccountCode`

## PayrollRuns

**مسيرات الرواتب** – مسير كل شهر: مسودة تُعدَّل، ثم يُرحَّل قيده في آخر يوم من الشهر، ثم يُسجَّل صرفه من بنك أو صندوق.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **PayrollRunID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | RunNumber | Short Text | 20 | ✔ |  |  |  | رقم المسير |
| 3 | PayMonth | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | الشهر (آخر يوم فيه) |
| 4 | Status | Short Text | 10 | ✔ | `"DRAFT"` | `In ("DRAFT","POSTED")` |  | الحالة |
| 5 | PaidDate | Date/Time (تاريخ) |  |  |  |  |  | تاريخ الصرف |
| 6 | PaidFrom | Short Text | 10 |  |  | `Is Null Or In ("BANK","CASHBOX")` |  | الصرف من |
| 7 | BankID | Number (Long) |  |  |  |  | `Banks.BankID` | البنك |
| 8 | CashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | الصندوق |
| 9 | PaidAmount | Currency |  | ✔ | `0` | `>=0` |  | المبلغ المصروف |
| 10 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 11 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | أعدّه |
| 12 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `PayrollRunID`
- فهرس فريد: `RunNumber`
- فهرس فريد: `PayMonth`

## PayrollLines

**أسطر مسير الرواتب** – سطر كل موظف: الراتب والبدلات والإضافي والخصومات والتأمينات وصافي الراتب.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **PayrollLineID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | PayrollRunID | Number (Long) |  | ✔ |  |  | `PayrollRuns.PayrollRunID` | المسير |
| 3 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 4 | EmployeeName | Short Text | 100 |  |  |  |  | اسم الموظف |
| 5 | CostCenterID | Number (Long) |  |  |  |  | `CostCenters.CostCenterID` | مركز التكلفة – مركز الموظف يوم إنشاء المسير |
| 6 | IsSaudi | Yes/No |  |  | `False` |  |  | سعودي |
| 7 | Basic | Currency |  | ✔ | `0` | `>=0` |  | الأساسي |
| 8 | Housing | Currency |  | ✔ | `0` | `>=0` |  | بدل السكن |
| 9 | OtherAllow | Currency |  | ✔ | `0` | `>=0` |  | بدلات أخرى (النقل وغيره) |
| 10 | Overtime | Currency |  | ✔ | `0` | `>=0` |  | العمل الإضافي |
| 11 | Additions | Currency |  | ✔ | `0` | `>=0` |  | مكافآت وإضافات |
| 12 | AbsenceDeduction | Currency |  | ✔ | `0` | `>=0` |  | خصم الغياب |
| 13 | AdvanceDeduction | Currency |  | ✔ | `0` | `>=0` |  | خصم السلفة |
| 14 | OtherDeduction | Currency |  | ✔ | `0` | `>=0` |  | خصومات أخرى (جزاءات) |
| 15 | GosiWage | Currency |  | ✔ | `0` | `>=0` |  | الأجر الخاضع للتأمينات |
| 16 | GosiEmployee | Currency |  | ✔ | `0` | `>=0` |  | التأمينات - حصة الموظف |
| 17 | GosiEmployer | Currency |  | ✔ | `0` | `>=0` |  | التأمينات - حصة المنشأة |
| 18 | NetPay | Currency |  | ✔ | `0` |  |  | صافي الراتب – قد يكون سالبًا في المسودة؛ الترحيل يرفضه (الخصومات أكبر من الراتب) |
| 19 | Notes | Short Text | 150 |  |  |  |  | ملاحظات |

- المفتاح الأساسي: `PayrollLineID`
- فهرس فريد: `PayrollRunID, EmployeeID`

## BankReconciliations

**التسويات البنكية** – مطابقة كشف البنك في تاريخ مع الدفاتر: رصيد الكشف، والرصيد في الدفاتر، والحركات غير الظاهرة في الكشف.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **ReconciliationID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | ReconNumber | Short Text | 20 | ✔ |  |  |  | رقم التسوية |
| 3 | BankID | Number (Long) |  | ✔ |  |  | `Banks.BankID` | البنك |
| 4 | StatementDate | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | تاريخ كشف البنك |
| 5 | StatementBalance | Currency |  | ✔ | `0` |  |  | رصيد كشف البنك |
| 6 | BookBalance | Currency |  | ✔ | `0` |  |  | الرصيد في الدفاتر في التاريخ |
| 7 | Outstanding | Currency |  | ✔ | `0` |  |  | حركات لم تظهر في الكشف (صافي) |
| 8 | Difference | Currency |  | ✔ | `0` |  |  | الفرق |
| 9 | Status | Short Text | 10 | ✔ | `"OPEN"` | `In ("OPEN","DONE")` |  | الحالة |
| 10 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 11 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 12 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `ReconciliationID`
- فهرس فريد: `ReconNumber`
- فهرس عادي: `BankID`

## BankClearings

**حركات الدفاتر المطابقة لكشف البنك** – كل عملية قيدها على حساب البنك ظهرت في كشف البنك: نوع العملية ورقمها ومبلغها يوم المطابقة.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **ClearingID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | ReconciliationID | Number (Long) |  | ✔ |  |  | `BankReconciliations.ReconciliationID` | التسوية |
| 3 | BankID | Number (Long) |  | ✔ |  |  | `Banks.BankID` | البنك |
| 4 | SourceType | Short Text | 20 | ✔ |  |  |  | نوع العملية |
| 5 | SourceID | Number (Long) |  | ✔ |  |  |  | رقم العملية |
| 6 | ClearedAmount | Currency |  | ✔ | `0` |  |  | المبلغ يوم المطابقة (مدين موجب) |
| 7 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `ClearingID`
- فهرس فريد: `BankID, SourceType, SourceID`

## CustomerAllocations

**ربط سندات القبض بالفواتير** – كم من سند القبض سدّد كل فاتورة آجلة. ما لا يُربط بفاتورة يسدد أقدم الفواتير استحقاقًا.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **AllocationID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | PaymentID | Number (Long) |  | ✔ |  |  | `CustomerPayments.PaymentID` | سند القبض |
| 3 | SalesInvoiceID | Number (Long) |  | ✔ |  |  | `SalesInvoices.SalesInvoiceID` | الفاتورة |
| 4 | Amount | Currency |  | ✔ | `0` | `>0` |  | المبلغ |
| 5 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 6 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `AllocationID`
- فهرس فريد: `PaymentID, SalesInvoiceID`
- فهرس عادي: `SalesInvoiceID`

## SupplierAllocations

**ربط سندات الصرف بفواتير الشراء** – كم من سند الصرف سدّد كل فاتورة شراء آجلة. ما لا يُربط بفاتورة يسدد أقدم الفواتير استحقاقًا.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **AllocationID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | PaymentID | Number (Long) |  | ✔ |  |  | `SupplierPayments.PaymentID` | سند الصرف |
| 3 | PurchaseInvoiceID | Number (Long) |  | ✔ |  |  | `PurchaseInvoices.PurchaseInvoiceID` | الفاتورة |
| 4 | Amount | Currency |  | ✔ | `0` | `>0` |  | المبلغ |
| 5 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 6 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `AllocationID`
- فهرس فريد: `PaymentID, PurchaseInvoiceID`
- فهرس عادي: `PurchaseInvoiceID`

## ExpenseTypes

**أنواع المصروفات** – قائمة ثابتة لأنواع المصروفات لضمان دقة التقارير المجمّعة.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **ExpenseTypeID** 🔑 | AutoNumber |  |  |  |  |  | رقم النوع |
| 2 | ExpenseTypeName | Short Text | 50 | ✔ |  |  |  | نوع المصروف |
| 3 | IsActive | Yes/No |  |  | `True` |  |  | نشط |

- المفتاح الأساسي: `ExpenseTypeID`
- فهرس فريد: `ExpenseTypeName`
- بيانات أساسية: 9 سجل

## Expenses

**المصروفات** – مصروفات المحل التشغيلية؛ المبلغ بدون ضريبة والضريبة منفصلة (ضريبة مدخلات).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **ExpenseID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | ExpenseNumber | Short Text | 20 | ✔ |  |  |  | رقم المصروف |
| 3 | ExpenseDate | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | تاريخ المصروف |
| 4 | ExpenseTypeID | Number (Long) |  | ✔ |  |  | `ExpenseTypes.ExpenseTypeID` | نوع المصروف |
| 5 | Amount | Currency |  | ✔ | `0` | `>0` |  | المبلغ قبل الضريبة |
| 6 | Tax | Currency |  | ✔ | `0` | `>=0` |  | ضريبة المدخلات – فقط إذا كانت لدى المحل فاتورة ضريبية بالمصروف |
| 7 | TotalAmount | Currency |  | ✔ | `0` | `>=0` |  | الإجمالي |
| 8 | PaymentMethodID | Number (Long) |  |  |  |  | `PaymentMethods.PaymentMethodID` | طريقة الدفع |
| 9 | SupplierInvoiceRef | Short Text | 30 |  |  |  |  | رقم فاتورة المصروف |
| 10 | Description | Short Text | 255 |  |  |  |  | الوصف |
| 11 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 12 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 13 | CashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | صُرف من صندوق – المصروف النقدي يخرج من هذا الصندوق؛ فارغ = لم يُدفع من صندوق |
| 14 | BankID | Number (Long) |  |  |  |  | `Banks.BankID` | البنك – المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك |
| 15 | CostCenterID | Number (Long) |  |  |  |  | `CostCenters.CostCenterID` | مركز التكلفة |
| 16 | RecurringID | Number (Long) |  |  |  |  | `RecurringExpenses.RecurringID` | من مصروف متكرر – أنشأه البرنامج من المصروف المتكرر بتاريخ استحقاقه |
| 17 | CurrencyCode | Short Text | 3 |  | `"SAR"` |  | `Currencies.CurrencyCode` | العملة – فارغ = عملة البرنامج (Settings.CurrencyCode) |
| 18 | ExchangeRate | Currency (نسبة) |  | ✔ | `1` | `>0` |  | معامل التحويل – قيمة وحدة واحدة من العملة بعملة البرنامج |
| 19 | ForeignAmount | Currency |  | ✔ | `0` | `>=0` |  | المبلغ بالعملة (بدون الضريبة) – الإجمالي بعملة المستند (0 للمستندات القديمة) |
| 20 | ForeignTax | Currency |  | ✔ | `0` | `>=0` |  | الضريبة بالعملة |

- المفتاح الأساسي: `ExpenseID`
- فهرس فريد: `ExpenseNumber`
- فهرس عادي: `ExpenseDate`
- قاعدة تحقق على مستوى الجدول: `[TotalAmount]=[Amount]+[Tax]` – الإجمالي = المبلغ + الضريبة

## RecurringExpenses

**المصروفات المتكررة** – مصروف ثابت يتكرر (الإيجار، الكهرباء، الاشتراكات): يُنشئ البرنامج المصروف في تاريخ استحقاقه، مرة واحدة لكل تاريخ.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **RecurringID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | RecurringName | Short Text | 100 | ✔ |  |  |  | اسم المصروف |
| 3 | ExpenseTypeID | Number (Long) |  | ✔ |  |  | `ExpenseTypes.ExpenseTypeID` | نوع المصروف |
| 4 | Amount | Currency |  | ✔ | `0` | `>0` |  | المبلغ قبل الضريبة |
| 5 | Tax | Currency |  | ✔ | `0` | `>=0` |  | ضريبة المدخلات |
| 6 | PaymentMethodID | Number (Long) |  | ✔ |  |  | `PaymentMethods.PaymentMethodID` | طريقة الدفع |
| 7 | CashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | من صندوق – للدفع النقدي؛ فارغ = صندوق من يُنشئ المصروف |
| 8 | BankID | Number (Long) |  |  |  |  | `Banks.BankID` | من بنك – للتحويل البنكي؛ فارغ = البنك الافتراضي |
| 9 | CostCenterID | Number (Long) |  |  |  |  | `CostCenters.CostCenterID` | مركز التكلفة |
| 10 | Frequency | Short Text | 10 | ✔ | `"MONTHLY"` | `In ("MONTHLY","QUARTERLY","YEARLY")` |  | التكرار |
| 11 | DueDay | Number (Integer) |  | ✔ | `1` | `Between 1 And 28` |  | يوم الاستحقاق |
| 12 | StartDate | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | يبدأ من |
| 13 | EndDate | Date/Time (تاريخ) |  |  |  |  |  | ينتهي في |
| 14 | NextDueDate | Date/Time (تاريخ) |  |  |  |  |  | الاستحقاق التالي |
| 15 | LastCreatedDate | Date/Time (تاريخ) |  |  |  |  |  | آخر مصروف أُنشئ |
| 16 | Description | Short Text | 255 |  |  |  |  | الوصف |
| 17 | IsActive | Yes/No |  |  | `True` |  |  | نشط |
| 18 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `RecurringID`
- فهرس فريد: `RecurringName`
- فهرس عادي: `NextDueDate`

## CashVouchers

**سندات النقدية** – قبض نقدية لصندوق، أو صرف منه، أو تحويل بين صندوقين (ومنها ترحيل يومية الكاشير).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **CashVoucherID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | VoucherNumber | Short Text | 20 | ✔ |  |  |  | رقم السند |
| 3 | VoucherDate | Date/Time |  | ✔ | `Now()` |  |  | التاريخ |
| 4 | VoucherType | Short Text | 10 | ✔ |  | `In ("IN","OUT","TRANSFER")` |  | نوع السند |
| 5 | CashBoxID | Number (Long) |  | ✔ |  |  | `CashBoxes.CashBoxID` | الصندوق – القبض يدخله، والصرف والتحويل يخرجان منه |
| 6 | ToCashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | إلى صندوق – للتحويل فقط |
| 7 | Category | Short Text | 10 | ✔ | `"OTHER"` | `In ("OTHER","OWNER","EXPENSE","ADVANCE","SHORTAGE","OVERAGE","TRANSFER","COMMISSION")` |  | البند |
| 8 | Amount | Currency |  | ✔ | `0` | `>0` |  | المبلغ |
| 9 | PartyName | Short Text | 100 |  |  |  |  | المستلم / المسلِّم |
| 10 | Description | Short Text | 255 |  |  |  |  | البيان |
| 11 | ExpenseID | Number (Long) |  |  |  |  | `Expenses.ExpenseID` | المصروف المسجَّل – صرف بند مصروف يسجل مصروفًا بنفس المبلغ في المصروفات |
| 12 | ClosingID | Number (Long) |  |  |  |  | `CashClosings.ClosingID` | تصفية الكاشير |
| 13 | AdvanceEmployeeID | Number (Long) |  |  |  |  | `Employees.EmployeeID` | الموظف صاحب السلفة – سلفة موظف (صرف) أو سدادها نقدًا (قبض) |
| 14 | CostCenterID | Number (Long) |  |  |  |  | `CostCenters.CostCenterID` | مركز التكلفة |
| 15 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | الموظف |
| 16 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 17 | SalesRepID | Number (Long) |  |  |  |  | `SalesReps.SalesRepID` | المندوب – صرف عمولة المندوب |

- المفتاح الأساسي: `CashVoucherID`
- فهرس فريد: `VoucherNumber`
- فهرس عادي: `VoucherDate`
- فهرس عادي: `CashBoxID`
- قاعدة تحقق على مستوى الجدول: `[VoucherType]<>"TRANSFER" Or ([ToCashBoxID] Is Not Null And [ToCashBoxID]<>[CashBoxID])` – التحويل يحتاج صندوقًا آخر غير صندوق الصرف

## CashClosings

**تصفية يومية الكاشير** – جرد نقدية صندوق الكاشير في نهاية الوردية وترحيلها للخزينة الرئيسية أو تسويتها مع المالك.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **ClosingID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | ClosingNumber | Short Text | 20 | ✔ |  |  |  | رقم التصفية |
| 3 | ClosingDate | Date/Time |  | ✔ | `Now()` |  |  | تاريخ التصفية |
| 4 | CashBoxID | Number (Long) |  | ✔ |  |  | `CashBoxes.CashBoxID` | الصندوق |
| 5 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | أجراها |
| 6 | PeriodStart | Date/Time |  |  |  |  |  | من (آخر تصفية) |
| 7 | OpeningBalance | Currency |  | ✔ | `0` |  |  | رصيد البداية |
| 8 | CashIn | Currency |  | ✔ | `0` | `>=0` |  | المقبوضات |
| 9 | CashOut | Currency |  | ✔ | `0` | `>=0` |  | المدفوعات |
| 10 | ExpectedBalance | Currency |  | ✔ | `0` |  |  | الرصيد الدفتري |
| 11 | CountedAmount | Currency |  | ✔ | `0` | `>=0` |  | النقدية الفعلية |
| 12 | Difference | Currency |  | ✔ | `0` |  |  | الفرق – سالب = عجز، موجب = زيادة |
| 13 | Destination | Short Text | 10 | ✔ | `"MAIN"` | `In ("MAIN","OWNER","KEEP")` |  | الترحيل إلى |
| 14 | ToCashBoxID | Number (Long) |  |  |  |  | `CashBoxes.CashBoxID` | الخزينة المستلمة |
| 15 | TransferAmount | Currency |  | ✔ | `0` | `>=0` |  | المبلغ المرحَّل |
| 16 | KeptAmount | Currency |  | ✔ | `0` | `>=0` |  | المتبقي في الصندوق (عهدة) |
| 17 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 18 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `ClosingID`
- فهرس فريد: `ClosingNumber`
- فهرس عادي: `CashBoxID, ClosingDate`
- قاعدة تحقق على مستوى الجدول: `[TransferAmount]+[KeptAmount]=[CountedAmount]` – المرحَّل + المتبقي = النقدية الفعلية

## Accounts

**دليل الحسابات (شجرة الحسابات)** – شجرة من خمسة مستويات على الأكثر: الحسابات الرئيسية (تجميعية) والحسابات الفرعية التي تُرحَّل إليها القيود. حسابات الصناديق (110000 + رقم الصندوق) وأنواع المصروفات (530000 + رقم النوع) تُنشأ تلقائيًا.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **AccountCode** 🔑 | Number (Long) |  | ✔ |  | `>0` |  | رقم الحساب |
| 2 | AccountName | Short Text | 100 | ✔ |  |  |  | اسم الحساب |
| 3 | AccountNameEn | Short Text | 100 |  |  |  |  | الاسم بالإنجليزية – يظهر في الواجهة الإنجليزية والقوائم المالية بها |
| 4 | AccountType | Short Text | 10 | ✔ |  | `In ("ASSET","LIABILITY","EQUITY","REVENUE","EXPENSE")` |  | نوع الحساب |
| 5 | ParentCode | Number (Long) |  |  |  |  |  | الحساب الرئيسي – فارغ للحسابات الخمسة في المستوى الأول فقط |
| 6 | IsActive | Yes/No |  |  | `True` |  |  | نشط |
| 7 | IsPosting | Yes/No |  |  | `True` |  |  | حساب فرعي (يقبل القيود) |
| 8 | IsSystem | Yes/No |  |  | `False` |  |  | حساب أساسي في النظام |
| 9 | AccountLevel | Number (Byte) |  | ✔ | `1` |  |  | المستوى |
| 10 | TreeKey | Short Text | 60 |  |  |  |  | مفتاح الترتيب في الشجرة – يحسبه البرنامج (modAccounts.RebuildAccountTree): رقم كل مستوى بعشر خانات |
| 11 | Level1Code | Number (Long) |  |  |  |  |  | حساب المستوى 1 |
| 12 | Level2Code | Number (Long) |  |  |  |  |  | حساب المستوى 2 |
| 13 | Level3Code | Number (Long) |  |  |  |  |  | حساب المستوى 3 |
| 14 | Level4Code | Number (Long) |  |  |  |  |  | حساب المستوى 4 |
| 15 | Level5Code | Number (Long) |  |  |  |  |  | حساب المستوى 5 |

- المفتاح الأساسي: `AccountCode`
- فهرس عادي: `ParentCode`
- فهرس عادي: `TreeKey`
- بيانات أساسية: 80 سجل

## JournalSourceTypes

**أنواع مصادر القيود** – أنواع العمليات التي يُنشأ عنها قيد آلي، ومنها يُعرف أصل القيد.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **SourceType** 🔑 | Short Text | 20 | ✔ |  |  |  | نوع العملية |
| 2 | TypeName | Short Text | 50 | ✔ |  |  |  | الاسم |
| 3 | TypeNameEn | Short Text | 50 |  |  |  |  | الاسم بالإنجليزية |
| 4 | SortOrder | Number (Integer) |  | ✔ | `0` |  |  | الترتيب |

- المفتاح الأساسي: `SourceType`
- بيانات أساسية: 27 سجل

## JournalEntries

**قيود اليومية** – قيد آلي لكل عملية، مربوط بأصلها (SourceType + SourceID). يُحدَّث إذا تغيرت العملية.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **EntryID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | EntryNumber | Short Text | 20 | ✔ |  |  |  | رقم القيد |
| 3 | EntryDate | Date/Time |  | ✔ |  |  |  | تاريخ القيد |
| 4 | SourceType | Short Text | 20 | ✔ |  |  | `JournalSourceTypes.SourceType` | نوع العملية |
| 5 | SourceID | Number (Long) |  | ✔ |  |  |  | رقم العملية الداخلي |
| 6 | SourceNumber | Short Text | 20 |  |  |  |  | رقم مستند العملية |
| 7 | Description | Short Text | 255 |  |  |  |  | البيان |
| 8 | TotalDebit | Currency |  | ✔ | `0` | `>=0` |  | إجمالي المدين |
| 9 | TotalCredit | Currency |  | ✔ | `0` | `>=0` |  | إجمالي الدائن |
| 10 | LineCount | Number (Integer) |  | ✔ | `0` |  |  | عدد الأسطر |
| 11 | Signature | Currency |  | ✔ | `0` | `>=0` |  | بصمة القيد – تكشف تغيّر العملية بعد إنشاء القيد |
| 12 | UpdatedAt | Date/Time |  |  |  |  |  | آخر تحديث |
| 13 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 14 | CurrencyCode | Short Text | 3 |  | `"SAR"` |  | `Currencies.CurrencyCode` | العملة – فارغ = عملة البرنامج (Settings.CurrencyCode) |
| 15 | ExchangeRate | Currency (نسبة) |  | ✔ | `1` | `>0` |  | معامل التحويل – قيمة وحدة واحدة من العملة بعملة البرنامج |
| 16 | ForeignAmount | Currency |  | ✔ | `0` | `>=0` |  | مبلغ المستند بعملته – الإجمالي بعملة المستند (0 للمستندات القديمة) |

- المفتاح الأساسي: `EntryID`
- فهرس فريد: `EntryNumber`
- فهرس فريد: `SourceType, SourceID`
- فهرس عادي: `EntryDate`
- قاعدة تحقق على مستوى الجدول: `[TotalDebit]=[TotalCredit]` – القيد غير متوازن: المدين يجب أن يساوي الدائن

## JournalLines

**أسطر القيود** – الطرف المدين والطرف الدائن لكل قيد.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **JournalLineID** 🔑 | AutoNumber |  |  |  |  |  | رقم السطر الداخلي |
| 2 | EntryID | Number (Long) |  | ✔ |  |  | `JournalEntries.EntryID` | القيد |
| 3 | LineNumber | Number (Integer) |  | ✔ |  |  |  | رقم السطر |
| 4 | AccountCode | Number (Long) |  | ✔ |  |  | `Accounts.AccountCode` | الحساب |
| 5 | Debit | Currency |  | ✔ | `0` | `>=0` |  | مدين |
| 6 | Credit | Currency |  | ✔ | `0` | `>=0` |  | دائن |
| 7 | LineText | Short Text | 255 |  |  |  |  | البيان |
| 8 | CostCenterID | Number (Long) |  |  |  |  | `CostCenters.CostCenterID` | مركز التكلفة – من مستند العملية؛ فارغ = غير موزع |

- المفتاح الأساسي: `JournalLineID`
- فهرس فريد: `EntryID, LineNumber`
- فهرس عادي: `AccountCode`

## PeriodClosings

**سجل إقفال الفترات** – كل إقفال أو إعادة فتح لفترة أو سنة مالية: التاريخ الجديد للإقفال والسابق، ومن قام به والسبب.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **PeriodClosingID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | ActionType | Short Text | 12 | ✔ |  | `In ("CLOSE","REOPEN","YEAR_CLOSE","YEAR_OPEN")` |  | العملية |
| 3 | ClosedThrough | Date/Time |  |  |  |  |  | مقفلة حتى (بعد العملية) |
| 4 | PreviousThrough | Date/Time |  |  |  |  |  | مقفلة حتى (قبل العملية) |
| 5 | FiscalYear | Number (Integer) |  |  |  |  |  | السنة المالية |
| 6 | Notes | Short Text | 255 |  |  |  |  | السبب / ملاحظات |
| 7 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | قام به |
| 8 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `PeriodClosingID`

## FiscalYearClosings

**إقفال السنوات المالية** – قيد إقفال كل سنة: أرصدة الإيرادات والمصروفات في 31 ديسمبر تُقفل في الأرباح المحتجزة (3300). أسطره محفوظة كما كانت يوم الإقفال، ويُرحَّل لليومية كعملية «قيد إقفال السنة».

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **YearClosingID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | FiscalYear | Number (Integer) |  | ✔ |  |  |  | السنة المالية |
| 3 | ClosingNumber | Short Text | 20 | ✔ |  |  |  | رقم الإقفال |
| 4 | ClosingDate | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | تاريخ الإقفال |
| 5 | NetProfit | Currency |  | ✔ | `0` |  |  | صافي ربح (خسارة) السنة |
| 6 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | أقفلها |
| 7 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 8 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `YearClosingID`
- فهرس فريد: `FiscalYear`
- فهرس فريد: `ClosingNumber`

## FiscalYearClosingLines

**أسطر قيود إقفال السنوات** – لكل حساب إيرادات أو مصروفات رصيده معكوسًا، ثم صافي الربح في الأرباح المحتجزة.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **YearClosingLineID** 🔑 | AutoNumber |  |  |  |  |  | رقم السطر الداخلي |
| 2 | YearClosingID | Number (Long) |  | ✔ |  |  | `FiscalYearClosings.YearClosingID` | إقفال السنة |
| 3 | LineNumber | Number (Integer) |  | ✔ |  |  |  | رقم السطر |
| 4 | AccountCode | Number (Long) |  | ✔ |  |  | `Accounts.AccountCode` | الحساب |
| 5 | Debit | Currency |  | ✔ | `0` | `>=0` |  | مدين |
| 6 | Credit | Currency |  | ✔ | `0` | `>=0` |  | دائن |
| 7 | LineText | Short Text | 150 |  |  |  |  | البيان |

- المفتاح الأساسي: `YearClosingLineID`
- فهرس فريد: `YearClosingID, LineNumber`

## VatReturns

**إقرارات ضريبة القيمة المضافة** – إقرار كل فترة ضريبية (شهر أو ربع سنة) بخانات نموذج هيئة الزكاة والضريبة والجمارك. المسودة تُحسب من المستندات، وعند الاعتماد تُحفظ قيمها كما هي ويُنشأ قيد التسوية (ضريبة المخرجات والمدخلات إلى حساب التسوية 2250)، ثم قيد السداد عند تسجيله.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **VatReturnID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | ReturnNumber | Short Text | 20 | ✔ |  |  |  | رقم الإقرار – VAT-yyyymmdd (آخر يوم في الفترة) |
| 3 | PeriodFrom | Date/Time (تاريخ) |  | ✔ |  |  |  | بداية الفترة |
| 4 | PeriodTo | Date/Time (تاريخ) |  | ✔ |  |  |  | نهاية الفترة |
| 5 | Status | Short Text | 10 | ✔ | `"DRAFT"` | `In ("DRAFT","FILED")` |  | الحالة |
| 6 | SalesStdAmount | Currency |  | ✔ | `0` | `>=0` |  | 1- المبيعات الخاضعة للنسبة الأساسية |
| 7 | SalesStdAdjust | Currency |  | ✔ | `0` |  |  | 1- تعديلات (مرتجعات) المبيعات الخاضعة |
| 8 | SalesStdVAT | Currency |  | ✔ | `0` |  |  | 1- ضريبة المبيعات الخاضعة |
| 9 | SalesZeroAmount | Currency |  | ✔ | `0` | `>=0` |  | 3- المبيعات بنسبة صفرية |
| 10 | SalesZeroAdjust | Currency |  | ✔ | `0` |  |  | 3- تعديلات المبيعات بنسبة صفرية |
| 11 | SalesExemptAmount | Currency |  | ✔ | `0` | `>=0` |  | 5- المبيعات المعفاة |
| 12 | SalesExemptAdjust | Currency |  | ✔ | `0` |  |  | 5- تعديلات المبيعات المعفاة |
| 13 | PurchStdAmount | Currency |  | ✔ | `0` | `>=0` |  | 7- المشتريات والمصروفات الخاضعة للنسبة الأساسية |
| 14 | PurchStdAdjust | Currency |  | ✔ | `0` |  |  | 7- تعديلات (مرتجعات) المشتريات الخاضعة |
| 15 | PurchStdVAT | Currency |  | ✔ | `0` |  |  | 7- ضريبة المشتريات الخاضعة |
| 16 | PurchZeroAmount | Currency |  | ✔ | `0` | `>=0` |  | 10- المشتريات بنسبة صفرية |
| 17 | PurchZeroAdjust | Currency |  | ✔ | `0` |  |  | 10- تعديلات المشتريات بنسبة صفرية |
| 18 | Corrections | Currency |  | ✔ | `0` |  |  | 14- تصحيحات من الفترات السابقة – موجبة تزيد الضريبة المستحقة، سالبة تنقصها |
| 19 | CarriedCredit | Currency |  | ✔ | `0` | `>=0` |  | 15- الرصيد الدائن المرحَّل من الفترات السابقة |
| 20 | NetDue | Currency |  | ✔ | `0` |  |  | 16- صافي الضريبة المستحقة (سالب = مستردة) – SalesStdVAT − PurchStdVAT + Corrections − CarriedCredit |
| 21 | FiledDate | Date/Time (تاريخ) |  |  |  |  |  | تاريخ الاعتماد (تاريخ قيد التسوية) |
| 22 | FilingRef | Short Text | 30 |  |  |  |  | رقم الإقرار لدى الهيئة |
| 23 | PaidDate | Date/Time (تاريخ) |  |  |  |  |  | تاريخ السداد |
| 24 | PaidAmount | Currency |  | ✔ | `0` | `>=0` |  | المبلغ المسدد |
| 25 | PaidAccount | Number (Long) |  |  |  |  | `Accounts.AccountCode` | حساب السداد (البنك) |
| 26 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 27 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | أعدّه / اعتمده |
| 28 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `VatReturnID`
- فهرس فريد: `ReturnNumber`
- فهرس عادي: `PeriodFrom`
- قاعدة تحقق على مستوى الجدول: `[PeriodTo]>=[PeriodFrom] And ([PaidAmount]=0 Or [PaidAmount]<=[NetDue])` – نهاية الفترة قبل بدايتها، أو المسدد أكبر من الضريبة المستحقة

## ManualEntries

**القيود اليدوية** – قيد يكتبه المحاسب بنفسه (مستحقات، تسويات، رأس المال، أرصدة افتتاحية...). يُرحَّل لليومية كأي عملية أخرى (SourceType = MANUAL)، ويُعدَّل أو يُحذف فيتبعه قيده.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **ManualEntryID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | EntryNumber | Short Text | 20 | ✔ |  |  |  | رقم القيد اليدوي |
| 3 | EntryDate | Date/Time (تاريخ) |  | ✔ | `Date()` |  |  | تاريخ القيد |
| 4 | Description | Short Text | 255 | ✔ |  |  |  | البيان |
| 5 | Reference | Short Text | 50 |  |  |  |  | المرجع – رقم مستند خارجي: فاتورة، عقد، كشف بنك... |
| 6 | ReversalOfID | Number (Long) |  |  |  |  |  | عكس القيد – القيد اليدوي الذي يعكسه هذا القيد |
| 7 | TotalAmount | Currency |  | ✔ | `0` | `>=0` |  | إجمالي القيد |
| 8 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | أدخله |
| 9 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |
| 10 | UpdatedAt | Date/Time |  |  |  |  |  | آخر تعديل |
| 11 | CurrencyCode | Short Text | 3 |  | `"SAR"` |  | `Currencies.CurrencyCode` | العملة – فارغ = عملة البرنامج (Settings.CurrencyCode) |
| 12 | ExchangeRate | Currency (نسبة) |  | ✔ | `1` | `>0` |  | معامل التحويل – قيمة وحدة واحدة من العملة بعملة البرنامج |
| 13 | ForeignAmount | Currency |  | ✔ | `0` | `>=0` |  | إجمالي القيد بالعملة – الإجمالي بعملة المستند (0 للمستندات القديمة) |

- المفتاح الأساسي: `ManualEntryID`
- فهرس فريد: `EntryNumber`
- فهرس عادي: `EntryDate`

## ManualEntryLines

**أسطر القيود اليدوية** – الطرف المدين والطرف الدائن للقيد اليدوي؛ كل سطر مدين أو دائن فقط.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **ManualLineID** 🔑 | AutoNumber |  |  |  |  |  | رقم السطر الداخلي |
| 2 | ManualEntryID | Number (Long) |  | ✔ |  |  | `ManualEntries.ManualEntryID` | القيد اليدوي |
| 3 | LineNumber | Number (Integer) |  | ✔ |  |  |  | رقم السطر |
| 4 | AccountCode | Number (Long) |  | ✔ |  |  | `Accounts.AccountCode` | الحساب |
| 5 | Debit | Currency |  | ✔ | `0` | `>=0` |  | مدين |
| 6 | Credit | Currency |  | ✔ | `0` | `>=0` |  | دائن |
| 7 | LineText | Short Text | 150 |  |  |  |  | بيان السطر |
| 8 | CostCenterID | Number (Long) |  |  |  |  | `CostCenters.CostCenterID` | مركز التكلفة |
| 9 | ForeignDebit | Currency |  | ✔ | `0` | `>=0` |  | مدين بالعملة |
| 10 | ForeignCredit | Currency |  | ✔ | `0` | `>=0` |  | دائن بالعملة |

- المفتاح الأساسي: `ManualLineID`
- فهرس فريد: `ManualEntryID, LineNumber`
- فهرس عادي: `AccountCode`
- قاعدة تحقق على مستوى الجدول: `([Debit]=0 Or [Credit]=0) And [Debit]+[Credit]>0` – كل سطر مدين أو دائن فقط، وبمبلغ أكبر من صفر

## TransactionTypes

**أنواع حركات المخزون** – أنواع الحركة وإشارتها: +1 تزيد المخزون، −1 تنقصه، 0 تسوية بالإشارة.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **TransactionTypeID** 🔑 | Number (Long) |  | ✔ |  |  |  | رقم النوع |
| 2 | TypeCode | Short Text | 20 | ✔ |  |  |  | رمز النوع |
| 3 | TypeName | Short Text | 50 | ✔ |  |  |  | نوع الحركة |
| 4 | TypeNameEn | Short Text | 50 |  |  |  |  | الاسم بالإنجليزية |
| 5 | Direction | Number (Integer) |  | ✔ |  | `In (-1,0,1)` |  | الاتجاه |
| 6 | IsManual | Yes/No |  |  | `False` |  |  | متاح للإدخال اليدوي |

- المفتاح الأساسي: `TransactionTypeID`
- فهرس فريد: `TypeCode`
- بيانات أساسية: 8 سجل

## InventoryTransactions

**حركة المخزون** – دفتر أستاذ المخزون: الكمية مخزنة بإشارتها، ورصيد أي صنف = مجموع Quantity.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **TransactionID** 🔑 | AutoNumber |  |  |  |  |  | رقم الحركة |
| 2 | TransactionDate | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الحركة |
| 3 | ProductID | Number (Long) |  | ✔ |  |  | `Products.ProductID` | المنتج |
| 4 | TransactionTypeID | Number (Long) |  | ✔ |  |  | `TransactionTypes.TransactionTypeID` | نوع الحركة |
| 5 | Quantity | Currency (كمية) |  | ✔ |  | `<>0` |  | الكمية (+/−) |
| 6 | UnitCost | Currency |  | ✔ | `0` | `>=0` |  | تكلفة الوحدة |
| 7 | QuantityAfter | Currency (كمية) |  | ✔ | `0` |  |  | الرصيد بعد الحركة |
| 8 | ReferenceType | Short Text | 20 |  |  |  |  | نوع المستند – SALE, SALES_RETURN, PURCHASE, PURCHASE_RETURN, STOCK_COUNT, MANUAL |
| 9 | ReferenceID | Number (Long) |  |  |  |  |  | رقم المستند الداخلي |
| 10 | ReferenceNumber | Short Text | 20 |  |  |  |  | رقم المستند |
| 11 | EmployeeID | Number (Long) |  |  |  |  | `Employees.EmployeeID` | الموظف |
| 12 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 13 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

- المفتاح الأساسي: `TransactionID`
- فهرس عادي: `TransactionDate`
- فهرس عادي: `ReferenceType, ReferenceID`

## StockCounts

**جلسات الجرد** – رأس عملية الجرد؛ تبقى مفتوحة حتى الترحيل الذي ينشئ حركات التسوية.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **StockCountID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | CountNumber | Short Text | 20 | ✔ |  |  |  | رقم الجرد |
| 3 | CountDate | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الجرد |
| 4 | CategoryID | Number (Long) |  |  |  |  | `Categories.CategoryID` | تصنيف محدد (اختياري) |
| 5 | Status | Short Text | 10 | ✔ | `"OPEN"` | `In ("OPEN","POSTED","CANCELLED")` |  | الحالة |
| 6 | EmployeeID | Number (Long) |  | ✔ |  |  | `Employees.EmployeeID` | أجراه |
| 7 | PostedAt | Date/Time |  |  |  |  |  | تاريخ الترحيل |
| 8 | PostedByID | Number (Long) |  |  |  |  | `Employees.EmployeeID` | رحّله |
| 9 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |

- المفتاح الأساسي: `StockCountID`
- فهرس فريد: `CountNumber`

## StockCountDetails

**تفاصيل الجرد** – الكمية المسجلة والفعلية والفرق لكل صنف في جلسة الجرد.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **StockCountDetailID** 🔑 | AutoNumber |  |  |  |  |  | رقم السطر الداخلي |
| 2 | StockCountID | Number (Long) |  | ✔ |  |  | `StockCounts.StockCountID` | الجرد |
| 3 | ProductID | Number (Long) |  | ✔ |  |  | `Products.ProductID` | المنتج |
| 4 | SystemQuantity | Currency (كمية) |  | ✔ | `0` |  |  | الكمية المسجلة |
| 5 | ActualQuantity | Currency (كمية) |  |  |  | `Is Null Or >=0` |  | الكمية الفعلية |
| 6 | Difference | Currency (كمية) |  | ✔ | `0` |  |  | الفرق – ActualQuantity − SystemQuantity |
| 7 | UnitCost | Currency |  | ✔ | `0` | `>=0` |  | تكلفة الوحدة |
| 8 | DifferenceValue | Currency |  | ✔ | `0` |  |  | قيمة الفرق – سالب = عجز، موجب = زيادة |
| 9 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |

- المفتاح الأساسي: `StockCountDetailID`
- فهرس فريد: `StockCountID, ProductID`

## AuditLog

**سجل العمليات** – سجل التدقيق: الدخول والخروج، وإضافة أي سجل أو تعديله أو حذفه (والحقول في AuditChanges)، وعمليات المستندات والعمليات الحساسة (تجاوز المخزون، تعديل الأسعار، النسخ الاحتياطي).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **LogID** 🔑 | AutoNumber |  |  |  |  |  | رقم السجل |
| 2 | LogDate | Date/Time |  | ✔ | `Now()` |  |  | التاريخ |
| 3 | EmployeeID | Number (Long) |  |  |  |  | `Employees.EmployeeID` | الموظف |
| 4 | ActionType | Short Text | 30 | ✔ |  |  |  | نوع العملية |
| 5 | ObjectName | Short Text | 50 |  |  |  |  | الكائن |
| 6 | RecordID | Short Text | 30 |  |  |  |  | رقم السجل المتأثر |
| 7 | Details | Long Text |  |  |  |  |  | التفاصيل |
| 8 | ComputerName | Short Text | 50 |  |  |  |  | اسم الجهاز |
| 9 | RecordLabel | Short Text | 100 |  |  |  |  | السجل – اسم السجل أو رقمه المقروء (المنتج، رقم المصروف...) |

- المفتاح الأساسي: `LogID`
- فهرس عادي: `LogDate`
- فهرس عادي: `ActionType`

## AuditChanges

**تفاصيل سجل التدقيق** – حقول كل عملية في سجل التدقيق: القيمة قبل وبعد (الإضافة: بعد فقط، الحذف: قبل فقط).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **ChangeID** 🔑 | AutoNumber |  |  |  |  |  | رقم داخلي |
| 2 | LogID | Number (Long) |  | ✔ |  |  | `AuditLog.LogID` | العملية |
| 3 | LineNo | Number (Integer) |  | ✔ | `0` |  |  | الترتيب |
| 4 | FieldName | Short Text | 64 | ✔ |  |  |  | الحقل |
| 5 | FieldCaption | Short Text | 100 |  |  |  |  | اسم الحقل |
| 6 | OldValue | Short Text | 255 |  |  |  |  | القيمة قبل |
| 7 | NewValue | Short Text | 255 |  |  |  |  | القيمة بعد |

- المفتاح الأساسي: `ChangeID`
- فهرس عادي: `LogID`

## LabelSettings

**إعدادات ملصقات الباركود** – سجل واحد: مقاس الملصق والورق والهوامش، وحجم الباركود، والنصوص أعلاه وأسفله.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **LabelSettingID** 🔑 | Number (Long) |  | ✔ | `1` | `=1` |  | رقم الإعداد |
| 2 | PrinterName | Short Text | 255 |  |  |  |  | طابعة الملصقات – فارغ = الطابعة الافتراضية |
| 3 | LabelWidth | Currency (كمية) |  | ✔ | `38` | `Between 15 And 210` |  | عرض الملصق (مم) |
| 4 | LabelHeight | Currency (كمية) |  | ✔ | `25` | `Between 10 And 297` |  | ارتفاع الملصق (مم) |
| 5 | LabelsAcross | Number (Byte) |  | ✔ | `1` | `Between 1 And 10` |  | عدد الملصقات في الصف |
| 6 | ColumnGap | Currency (كمية) |  | ✔ | `2` | `Between 0 And 50` |  | المسافة بين الأعمدة (مم) |
| 7 | RowGap | Currency (كمية) |  | ✔ | `0` | `Between 0 And 50` |  | المسافة بين الصفوف (مم) |
| 8 | MarginTop | Currency (كمية) |  | ✔ | `0` | `Between 0 And 50` |  | الهامش العلوي (مم) |
| 9 | MarginBottom | Currency (كمية) |  | ✔ | `0` | `Between 0 And 50` |  | الهامش السفلي (مم) |
| 10 | MarginLeft | Currency (كمية) |  | ✔ | `0` | `Between 0 And 50` |  | الهامش الأيسر (مم) |
| 11 | MarginRight | Currency (كمية) |  | ✔ | `0` | `Between 0 And 50` |  | الهامش الأيمن (مم) |
| 12 | BarHeight | Currency (كمية) |  | ✔ | `10` | `Between 3 And 100` |  | ارتفاع الباركود (مم) |
| 13 | BarWidth | Currency (كمية) |  | ✔ | `0.25` | `Between 0.1 And 1` |  | عرض أرفع خط (مم) |
| 14 | TopLine1 | Short Text | 10 | ✔ | `"STORE"` | `In ("NONE","STORE","NAME","PRICE","CODE","BARCODE")` |  | السطر الأول أعلى الباركود |
| 15 | TopLine2 | Short Text | 10 | ✔ | `"NAME"` | `In ("NONE","STORE","NAME","PRICE","CODE","BARCODE")` |  | السطر الثاني أعلى الباركود |
| 16 | BottomLine1 | Short Text | 10 | ✔ | `"BARCODE"` | `In ("NONE","STORE","NAME","PRICE","CODE","BARCODE")` |  | السطر الأول أسفل الباركود |
| 17 | BottomLine2 | Short Text | 10 | ✔ | `"PRICE"` | `In ("NONE","STORE","NAME","PRICE","CODE","BARCODE")` |  | السطر الثاني أسفل الباركود |
| 18 | ShortName | Short Text | 30 |  |  |  |  | الاسم المختصر للمحل – يُطبع إذا اخترت «الاسم المختصر» |
| 19 | FontSize | Number (Byte) |  | ✔ | `7` | `Between 5 And 16` |  | حجم الخط |

- المفتاح الأساسي: `LabelSettingID`
- بيانات أساسية: 1 سجل
