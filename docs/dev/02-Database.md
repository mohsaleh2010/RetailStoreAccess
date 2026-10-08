# قاعدة البيانات: كيف تُفهم وكيف تُعدَّل

- **مصدر الحقيقة:** `tools/schema.py`.
- **المرجع الكامل المولَّد:** [../02-Tables-Reference.md](../02-Tables-Reference.md)، وفيه كل حقل ونوعه وقاعدته، و[../03-Relationships-Reference.md](../03-Relationships-Reference.md) للعلاقات.
- **هذا الملف:** يشرح الصورة الكبيرة والقواعد التي يجب احترامها.

## 1) المبادئ
- **البيانات في ملف منفصل:** `RetailStore_BE.accdb`. الواجهة ترتبط بجداوله.
- **لا حذف للتاريخ:**
  - الفواتير والمرتجعات والسندات لا تُعدَّل ولا تُحذف بعد الحفظ. التصحيح يكون بمستند عكسي.
  - الموظفون والعملاء والمنتجات لا تُحذف إن كان لها حركات، بل تُعطَّل (`IsActive`).
- **الأرصدة لا تُخزَّن كمرجع، بل تُحسب من الحركات:**
  - المخزون = مجموع `InventoryTransactions`.
  - الصندوق = حركات `qryCashMovements`.
  - الحسابات = `JournalLines`.
  - الحقول المساعدة، مثل `CurrentQuantity` و`CurrentBalance`، تُحدَّث بالكود ويمكن إعادة احتسابها.
- **كل مستند يُحفظ في معاملة واحدة** (`BeginTrans` / `CommitTrans`)، وأرقامه من `Sequences` بلا فجوات (`NextNumber`).
- **العلاقات مفروضة التكامل:**
  - الحذف المتتالي فقط من رأس المستند إلى أسطره.
  - المفاتيح النصية فقط تتحدث تلقائيًا.
- **العملات:** كل حقول المبالغ **بعملة البرنامج (الريال)**. المستند بعملة أخرى يحفظ أيضًا `CurrencyCode` و`ExchangeRate` و`ForeignAmount`، فلا يتغير أي استعلام أو قيد. التفاصيل في `docs/36-Currencies.md`.
- **التحديث الآمن:** `BuildSchema` يضيف الجداول والحقول الناقصة إلى ملف بيانات موجود. لذلك أي حقل جديد في جدول موجود يجب أن يكون اختياريًا، أو له قيمة افتراضية.

## 2) الجداول حسب المجال
| المجال | الجداول | ملاحظات |
|---|---|---|
| **النظام** | `Settings` (سجل واحد)، `Sequences`، `Roles`، `Permissions`، `RolePermissions`، `Employees`، `Screens`، `UserScreens`، `Activations`، `AuditLog`، `AuditChanges`، `LabelSettings` | كل مستخدم هو موظف. المبرمج (`IsDeveloper`) مخفي. كلمات المرور SHA-256 مع salt |
| **البيانات الأساسية** | `Categories`، `Units`، `PaymentMethods`، `Products`، `Customers`، `Suppliers`، `ExpenseTypes`، `CashBoxes`، `Banks`، `CostCenters` | العميل رقم 1 هو «عميل نقدي» |
| **المبيعات** | `SalesInvoices` + `SalesInvoiceDetails`، `SalesReturns` + `SalesReturnDetails`، `CustomerPayments`، `CustomerAllocations` | تكلفة الصنف تُحفظ في السطر لحظة البيع |
| **المندوبين** | `SalesReps`، `SalesRepTargets`، `CommissionRuns` + `CommissionLines` | `SalesRepID` على العملاء والفواتير والمرتجعات وسندات القبض وسندات النقدية |
| **المشتريات والمخزون** | `PurchaseInvoices` + `Details`، `PurchaseReturns` + `Details`، `SupplierPayments`، `SupplierAllocations`، `TransactionTypes`، `InventoryTransactions`، `StockCounts` + `StockCountDetails` | المخزون دفتر أستاذ بالكمية وإشارتها |
| **الخزينة والبنوك** | `CashVouchers`، `CashClosings`، `BankTransactions`، `BankReconciliations` + `BankClearings`، `Cheques` | الصناديق والبنوك لها حسابات في الدليل |
| **المحاسبة** | `Accounts` (شجرة 5 مستويات)، `JournalSourceTypes`، `JournalEntries` + `JournalLines`، `ManualEntries` + `ManualEntryLines`، `PeriodClosings`، `FiscalYearClosings` + `Lines`، `VatReturns` | القيود الآلية تُبنى من المستندات |
| **التوسعات** | `FixedAssets`، `DepreciationRuns` + `AssetDepreciations`، `PayrollRuns` + `PayrollLines`، `Budgets` + `BudgetLines`، `Expenses`، `RecurringExpenses` | |

الجداول المؤقتة المحلية في ملف الواجهة (`tmpPOSLines`، `tmpPurchaseLines`، `tmpAging`، `tmpLedger`...) تُنشأ في `modPOS.EnsureLocalTables`. **ليس لها قيم افتراضية.**

## 3) النموذج المحاسبي
- **قيد آلي لكل مستند:** `SyncJournal` (`modJournal`) يقرأ استعلام مصدر لكل نوع مستند (`qryJournalSale`، `qryJournalExpense`...)، ثم:
  - يُنشئ القيد الناقص؛
  - يحدّث القيد الذي تغيّر توقيعه؛
  - يحذف القيد الذي حُذف مستنده.
- **مطابقة القوائم:** قائمة الاستعلامات `JOURNAL_SOURCE_QUERIES` في `tools/queries.py` يجب أن تطابق `SOURCE_QUERIES` في `modJournal`.
- **الحسابات المهمة:**

  | الحساب | الرقم |
  |---|---|
  | الصناديق | 110000 + رقم الصندوق |
  | البنوك | 120000 + رقم البنك |
  | تحصيلات مدى والمحافظ | 1200 |
  | العملاء | 1300 |
  | ضريبة المدخلات | 1500 |
  | السلف | 1600 |
  | مجمع الإهلاك | 1790 |
  | الموردون | 2100 |
  | أوراق الدفع | 2110 |
  | الرواتب والتأمينات | 2310 / 2320 |
  | عمولات المندوبين المستحقة / مصروفها | 2330 / 5530 |
  | الأرباح المحتجزة | 3300 |
  | الإيرادات | فئة 4 |
  | المصروفات | فئة 5 |

- **مراكز التكلفة:** على `JournalLines.CostCenterID`.
- **الفترات المقفلة:** `Settings.ClosedThrough`. أي تعديل بتاريخ مقفل يرفضه `ClosedPeriodProblem`.

## 4) الصلاحيات
- **صلاحيات الدور:** في `Permissions` و`RolePermissions`. ترى المدير والمدير التنفيذي والكاشير في `_role_permissions` في schema.
- **صلاحية كل شاشة:** في `SCREEN_LIST` (جدول `Screens`)، ومعها هل تنطبق عليها الإضافة والتعديل والحذف.
- **صلاحيات خاصة لمستخدم:** في `UserScreens`، عندما تُفعَّل له «صلاحيات شاشات خاصة».
- **مدير النظام والمبرمج:** كل الصلاحيات دائمًا.

## 5) حدود Access التي تؤثر على التصميم
- **32 فهرسًا لكل جدول:** كل علاقة مفروضة تُحسب على الجدولين.
  - لذلك حقول «سجّله الموظف» في جداول السجلات لها `fk` بلا علاقة مفروضة (`relations.UNENFORCED`).
  - الاختبار يُبقي كل جدول 4 فهارس أو أكثر تحت الحد.
- **أطوال النصوص:** حقل Short Text حتى 255 حرفًا. الأطول يكون `memo` (Long Text).
- **التواريخ:** نوع `DATE` تاريخ فقط، و`DATETIME` تاريخ ووقت.

## 6) خطوات إضافة جدول أو حقل
1. **عرّفه في `tools/schema.py`:** النوع، والعنوان العربي، والقاعدة ونصها العربي، و`fk=` إن كان يشير لجدول آخر، والفهارس.
2. **حقل جديد في جدول موجود:** اجعله `required=False`، أو أعطه `default`.
3. **جدول له أسطر:** علاقة الأسطر بالرأس `cascade=True`، وأضف الزوج إلى قائمة الحذف المتتالي في `tests/test_relations.py`.
4. **حدّث الاختبارات المعتمدة على العدد:** عدد العلاقات في `test_relations.py`، والجداول المستثناة، وغيرها.
5. **شغّل `python3 tools/generate.py`:** يتحدث المرجع وكود `BuildSchema` تلقائيًا.
6. **ثم اختبر:** `python3 -m unittest discover -s tests`.
