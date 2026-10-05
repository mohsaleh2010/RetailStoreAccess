# مرجع الجداول (Tables Reference)

> ملف مُولَّد تلقائيًا من `tools/schema.py` بواسطة `tools/generate.py` – لا تعدّله يدويًا.

عدد الجداول: **30** | عدد الحقول: **365**

## الفهرس

1. [`Settings`](#settings) – إعدادات المحل
2. [`Sequences`](#sequences) – عدّادات الترقيم
3. [`Roles`](#roles) – الأدوار
4. [`Permissions`](#permissions) – الصلاحيات
5. [`RolePermissions`](#rolepermissions) – صلاحيات الأدوار
6. [`Employees`](#employees) – الموظفون والمستخدمون
7. [`Categories`](#categories) – التصنيفات
8. [`Units`](#units) – وحدات القياس
9. [`PaymentMethods`](#paymentmethods) – طرق الدفع
10. [`Suppliers`](#suppliers) – الموردون
11. [`Customers`](#customers) – العملاء
12. [`Products`](#products) – المنتجات
13. [`SalesInvoices`](#salesinvoices) – فواتير المبيعات
14. [`SalesInvoiceDetails`](#salesinvoicedetails) – تفاصيل فواتير المبيعات
15. [`SalesReturns`](#salesreturns) – مرتجعات المبيعات
16. [`SalesReturnDetails`](#salesreturndetails) – تفاصيل مرتجعات المبيعات
17. [`PurchaseInvoices`](#purchaseinvoices) – فواتير المشتريات
18. [`PurchaseInvoiceDetails`](#purchaseinvoicedetails) – تفاصيل فواتير المشتريات
19. [`PurchaseReturns`](#purchasereturns) – مرتجعات المشتريات
20. [`PurchaseReturnDetails`](#purchasereturndetails) – تفاصيل مرتجعات المشتريات
21. [`CustomerPayments`](#customerpayments) – دفعات العملاء (سندات القبض)
22. [`SupplierPayments`](#supplierpayments) – دفعات الموردين (سندات الصرف)
23. [`ExpenseTypes`](#expensetypes) – أنواع المصروفات
24. [`Expenses`](#expenses) – المصروفات
25. [`TransactionTypes`](#transactiontypes) – أنواع حركات المخزون
26. [`InventoryTransactions`](#inventorytransactions) – حركة المخزون
27. [`StockCounts`](#stockcounts) – جلسات الجرد
28. [`StockCountDetails`](#stockcountdetails) – تفاصيل الجرد
29. [`AuditLog`](#auditlog) – سجل العمليات
30. [`LabelSettings`](#labelsettings) – إعدادات ملصقات الباركود

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
- بيانات أساسية: 11 سجل

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
- بيانات أساسية: 22 سجل

## RolePermissions

**صلاحيات الأدوار** – ربط كل دور بالصلاحيات الممنوحة له (علاقة متعدد لمتعدد).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **RoleID** 🔑 | Number (Long) |  | ✔ |  |  | `Roles.RoleID` | الدور |
| 2 | **PermissionKey** 🔑 | Short Text | 50 | ✔ |  |  | `Permissions.PermissionKey` | الصلاحية |

- المفتاح الأساسي: `RoleID, PermissionKey`
- بيانات أساسية: 44 سجل

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

- المفتاح الأساسي: `EmployeeID`
- فهرس فريد: `Username`
- بيانات أساسية: 1 سجل

## Categories

**التصنيفات** – تصنيفات المنتجات (إلكترونيات، مواد غذائية، ...).

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **CategoryID** 🔑 | AutoNumber |  |  |  |  |  | رقم التصنيف |
| 2 | CategoryName | Short Text | 100 | ✔ |  |  |  | اسم التصنيف |
| 3 | Description | Short Text | 255 |  |  |  |  | الوصف |
| 4 | IsActive | Yes/No |  |  | `True` |  |  | نشط |

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
| 3 | ZatcaCode | Short Text | 5 |  |  |  |  | رمز الهيئة |
| 4 | SortOrder | Number (Integer) |  | ✔ | `0` |  |  | الترتيب |
| 5 | IsActive | Yes/No |  |  | `True` |  |  | نشط |

- المفتاح الأساسي: `PaymentMethodID`
- فهرس فريد: `MethodName`
- بيانات أساسية: 4 سجل

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
| 13 | IsActive | Yes/No |  |  | `True` |  |  | نشط |
| 14 | Notes | Long Text |  |  |  |  |  | ملاحظات |
| 15 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

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
| 18 | IsSystem | Yes/No |  |  | `False` |  |  | سجل نظام |
| 19 | IsActive | Yes/No |  |  | `True` |  |  | نشط |
| 20 | Notes | Long Text |  |  |  |  |  | ملاحظات |
| 21 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

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
| 17 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 18 | InvoiceSubType | Short Text | 10 | ✔ | `"SIMPLIFIED"` | `In ("SIMPLIFIED","STANDARD")` |  | نوع الفاتورة الضريبية – مبسطة للأفراد B2C، ضريبية للمنشآت B2B (للعميل رقم ضريبي) |
| 19 | InvoiceTypeCode | Short Text | 3 | ✔ | `"388"` | `In ("388","381","383")` |  | رمز نوع المستند |
| 20 | InvoiceUUID | Short Text | 36 |  |  |  |  | المعرّف الفريد UUID |
| 21 | ICV | Number (Long) |  |  |  |  |  | عدّاد الفواتير ICV – تسلسل مشترك لكل المستندات المرسلة للهيئة |
| 22 | InvoiceHash | Short Text | 255 |  |  |  |  | بصمة المستند |
| 23 | PreviousInvoiceHash | Short Text | 255 |  |  |  |  | بصمة المستند السابق |
| 24 | QRCodeData | Long Text |  |  |  |  |  | بيانات رمز QR |
| 25 | ZatcaStatus | Short Text | 20 | ✔ | `"NOT_SENT"` | `In ("NOT_SENT","PENDING","REPORTED","CLEARED","WARNING","REJECTED")` |  | حالة الإرسال للهيئة |
| 26 | ZatcaSubmittedAt | Date/Time |  |  |  |  |  | تاريخ الإرسال للهيئة |
| 27 | ZatcaResponse | Long Text |  |  |  |  |  | رد الهيئة |
| 28 | SignedXmlPath | Short Text | 255 |  |  |  |  | مسار ملف XML الموقّع |
| 29 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

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
| 16 | Notes | Short Text | 255 |  |  |  |  | ملاحظات |
| 17 | CreatedAt | Date/Time |  | ✔ | `Now()` |  |  | تاريخ الإنشاء |

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

- المفتاح الأساسي: `PaymentID`
- فهرس فريد: `PaymentNumber`
- فهرس عادي: `PaymentDate`

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

- المفتاح الأساسي: `ExpenseID`
- فهرس فريد: `ExpenseNumber`
- فهرس عادي: `ExpenseDate`
- قاعدة تحقق على مستوى الجدول: `[TotalAmount]=[Amount]+[Tax]` – الإجمالي = المبلغ + الضريبة

## TransactionTypes

**أنواع حركات المخزون** – أنواع الحركة وإشارتها: +1 تزيد المخزون، −1 تنقصه، 0 تسوية بالإشارة.

| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |
|---|---|---|---|---|---|---|---|---|
| 1 | **TransactionTypeID** 🔑 | Number (Long) |  | ✔ |  |  |  | رقم النوع |
| 2 | TypeCode | Short Text | 20 | ✔ |  |  |  | رمز النوع |
| 3 | TypeName | Short Text | 50 | ✔ |  |  |  | نوع الحركة |
| 4 | Direction | Number (Integer) |  | ✔ |  | `In (-1,0,1)` |  | الاتجاه |
| 5 | IsManual | Yes/No |  |  | `False` |  |  | متاح للإدخال اليدوي |

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

**سجل العمليات** – يسجل الدخول والخروج والعمليات الحساسة (تجاوز المخزون، تعديل الأسعار، النسخ الاحتياطي).

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

- المفتاح الأساسي: `LogID`
- فهرس عادي: `LogDate`
- فهرس عادي: `ActionType`

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
