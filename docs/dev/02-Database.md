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
- **الأسماء الإنجليزية:** للجداول في `tools/master_en.py` (الحسابات، وطرق الدفع، وأنواع القيود وحركات المخزون، والأدوار، والصلاحيات، والشاشات، والتصنيفات، والوحدات، وأنواع المصروفات، والعملاء، والموردون، والصناديق، والبنوك، ومراكز التكلفة، والمندوبون) حقل اسم إنجليزي اختياري (`AccountNameEn`، `CustomerNameEn`...)، وللصلاحيات والشاشات أيضًا اسم المجموعة `ModuleNameEn` (`EXTRA_NAMES`). الواجهة الإنجليزية تقرأ الاستعلام `qryLoc<الجدول>`، وفيه نفس الأعمدة والاسم بالإنجليزية. التفاصيل في `docs/39-English-Master-Data.md` و`docs/40-English-Party-Names.md`.
- **العملات:** كل حقول المبالغ **بعملة البرنامج (الريال)**. المستند بعملة أخرى يحفظ أيضًا `CurrencyCode` و`ExchangeRate` و`ForeignAmount`، فلا يتغير أي استعلام أو قيد. التفاصيل في `docs/36-Currencies.md`.
- **التحديث الآمن:** `BuildSchema` يضيف الجداول والحقول الناقصة إلى ملف بيانات موجود. لذلك أي حقل جديد في جدول موجود يجب أن يكون اختياريًا، أو له قيمة افتراضية.

## 2) الجداول حسب المجال
| المجال | الجداول | ملاحظات |
|---|---|---|
| **النظام** | `Settings` (سجل واحد)، `Sequences`، `Roles`، `Permissions`، `RolePermissions`، `Employees`، `Screens`، `UserScreens`، `Activations`، `AuditLog`، `AuditChanges`، `LabelSettings` | كل مستخدم هو موظف. المبرمج (`IsDeveloper`) مخفي. كلمات المرور SHA-256 مع salt |
| **البيانات الأساسية** | `Categories`، `Units`، `PaymentMethods`، `Products`، `Customers`، `Suppliers`، `ExpenseTypes`، `CashBoxes`، `Banks`، `CostCenters` | العميل رقم 1 هو «عميل نقدي» |
| **المبيعات** | `SalesInvoices` + `SalesInvoiceDetails`، `SalesReturns` + `SalesReturnDetails`، `CustomerPayments`، `CustomerAllocations` | تكلفة الصنف تُحفظ في السطر لحظة البيع |
| **المندوبين** | `SalesReps`، `SalesRepTargets`، `CommissionRuns` + `CommissionLines` | `SalesRepID` على العملاء والفواتير والمرتجعات وسندات القبض وسندات النقدية |
| **المشتريات والمخزون** | `PurchaseInvoices` + `Details`، `PurchaseReturns` + `Details`، `SupplierPayments`، `SupplierAllocations`، `TransactionTypes`، `InventoryTransactions`، `StockCounts` + `StockCountDetails`، `StockTransfers` + `StockTransferLines` (التحويل بين المخازن، `docs/58`: مسودة ثم بالطريق ثم مُستلَم)، `StockParties` و`StockDocuments` + `StockDocumentLines` (جهات وسندات الصرف والاستلام، `docs/59`: مسودة ثم مُرحَّل)، `StockNotes` + `StockNoteLines` (أذونات التسليم والاستلام، `docs/61`: مسودة ثم مُرحَّل ثم صدرت فاتورته `SalesInvoiceID` / `PurchaseInvoiceID`، و`SalesOrderID` / `SalesOrderLineID` لإذن من طلبية)، `RepDays` (يوم المندوب، `docs/63`: مفتوح ثم تمت تسويته، يحفظ جرد السيارة وتصفية الصندوق، وعجز البضاعة المحمَّل على المندوب سلفة `GoodsShortage` / `AdvanceEmployeeID`، `docs/66`)، `RouteStops` و`RepVisits` و`RepPrices` (خط السير وزيارات اليوم ونتيجتها وقائمة أسعار المندوب، `docs/64`؛ وحدّا `SalesReps.CreditLimit` و`MaxDiscountPercent`)، `SalesOrders` + `SalesOrderLines` (طلبيات العملاء، `docs/62`: مسودة ثم معتمدة تحجز ثم قيد التجهيز ثم مُسلَّمة جزئيًا أو كليًا ثم مفوترة، أو ملغاة / مغلقة؛ المُسلَّم = أسطر الأذونات المرحّلة)، `Warehouses` (المخازن، `docs/57`: كل حركة ومستند ورصيد افتتاحي وجرد يحمل `WarehouseID`، القيمة الافتراضية 1 = المخزن الرئيسي) | المخزون دفتر أستاذ بالكمية وإشارتها، ورصيد الصنف في المخزن = مجموع حركاته فيه؛ وقيمة الحركة = الكمية × `UnitCost` (كارت الصنف `ItemCardQuery`، `docs/60`)؛ وجرد مفتوح واحد لكل مخزن |
| **الخزينة والبنوك** | `CashVouchers` (التحويل لصندوق يؤكد الاستلام: `InTransit` ثم `ReceivedAt`، حساب 1195 بينهما)، `CashClosings`، `CashBoxUsers` (أمناء كل صندوق، `docs/56`)، `BankTransactions`، `BankReconciliations` + `BankClearings`، `Cheques` | الصناديق والبنوك لها حسابات في الدليل |
| **المحاسبة** | `Accounts` (شجرة 5 مستويات)، `JournalSourceTypes`، `JournalEntries` + `JournalLines`، `ManualEntries` + `ManualEntryLines`، `PeriodClosings`، `FiscalPeriods` (الفترات المالية بالترتيب؛ إقفال الفترة مفتاحه سنة نهايتها، `docs/54`)، `FiscalYearClosings` + `Lines`، `VatReturns`، `OpeningBalances` (أرصدة الحسابات العامة يوم بدء التشغيل `Settings.OpeningDate`، مقابلها 3900 «إجمالي الأرصدة الافتتاحية»، `docs/52`)، `OpeningStock` (كمية وتكلفة كل صنف يوم بدء التشغيل؛ `TransactionID` = حركة المخزون المرحّلة، `docs/53`) | القيود الآلية تُبنى من المستندات |
| **التوسعات** | `FixedAssets` (`DepMethod` / `DepRate` / `DisposalVAT`، `docs/65`)، `DepreciationRuns` + `AssetDepreciations` (`DepCenter` وقت الإهلاك)، `AssetRevaluations` و`AssetTransfers` (إعادة التقييم والنقل بين مراكز التكلفة، `docs/65`)، `PayrollRuns` + `PayrollLines`، `Budgets` + `BudgetLines`، `Expenses`، `RecurringExpenses` | |

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
- **الفاتورة الإلكترونية:** حالة كل فاتورة ومرتجع في `ZatcaStatus` (السعودية ومصر)، وكل طلب للمنظومة في `EInvoiceLog` (بلا علاقة: `DocKind` + `DocID`). التفعيل والبيئة في `Settings.EInvoiceEnabled` و`EInvoiceEnvironment` (`docs/45`).
- **تسجيل الجهاز لدى هيئة الزكاة:** حقول `Settings.Zatca*` (الرقم التسلسلي، وطلب الشهادة، وشهادتا الامتثال والفعلية وكلمتاهما السريتان، والبيئة والمرحلة: `COMPLIANCE` ثم `CHECKED` ثم `PRODUCTION`). المفتاح الخاص ملف خارج قاعدة البيانات، والكلمات السرية لا تُكتب في `EInvoiceLog` (`docs/47`).
- **الإيصال الإلكتروني المصري:** بيانات الجهاز في `Settings.Eta*`، وسلسلة الإيصالات في `Settings.EtaLastUUID`. لكل فاتورة ومرتجع `EtaUUID` و`EtaPreviousUUID` و`EtaSubmissionId`، والإيصال نفسه في `EInvoiceXml`. كود الصنف في `Products.EtaItemType` و`EtaItemCode`، ورمز الوحدة في `Units.EtaUnitCode`، والرقم القومي في `Customers.NationalID` (`docs/48`).
- **الفاتورة الإلكترونية المصرية:** بيانات دخول البرنامج وبرنامج التوقيع في `Settings.EtaErp*` و`EtaSigner*` و`EtaTokenPin`، والرقم الطويل للمستند في `EtaLongId` (`docs/49`).
- **دولة التشغيل:** `Settings.CountryCode` (SA أو EG) تحدد عملة البرنامج ونسبة الضريبة وشكل الرقم الضريبي (`docs/44`). شرط حقل تغيّر بعد الإصدار الأول يوضع في `schema.RULE_UPGRADES` حتى يصل إلى ملف بيانات موجود.
- **الفترات المقفلة:** `Settings.ClosedThrough`. أي تعديل بتاريخ مقفل يرفضه `ClosedPeriodProblem`.
- **الوحدات المتعددة (`docs/72`):** `ProductUnits` (الصنف، والوحدة، و`UnitFactor` = عدد الحبات فيها، وباركود وسعر بيع وشراء خاصّان، 0 = سعر الحبة × العدد). أسطر البيع والشراء ومرتجعاتهما فيها `UnitID` (فارغ = الوحدة الصغرى) و`UnitFactor` (1 افتراضيًا)، والكمية والأسعار والتكلفة بوحدة السطر. المخزون والمتوسط بالحبة.
- **قوائم الأسعار (`docs/73`):** `PriceLists` (الاسم، و`DiscountPercent` خصم الأصناف غير المسعّرة، نسبة من 0 إلى أقل من 1) و`PriceListItems` (القائمة مع الحذف المتتالي، والصنف، و`UnitID` فارغ = الوحدة الصغرى، و`UnitPrice` شامل الضريبة، و`ValidFrom` / `ValidTo` فارغ = بلا حد؛ الساري الذي بدأ آخرًا يغلب). `Customers.PriceListID` = قائمة العميل (فئته).
- **أوامر الشراء (`docs/74`):** `PurchaseOrders` (المورد، والمخزن الذي يستلم، و`ExpectedDate`، والحالة DRAFT / APPROVED / PARTIAL / RECEIVED / INVOICED / CANCELLED / CLOSED) و`PurchaseOrderLines` (الكمية، و`UnitCost` بدون الضريبة). إذن الاستلام منه: `StockNotes.PurchaseOrderID` و`StockNoteLines.PurchaseOrderLineID`. `Products.ReorderQuantity` = كمية الاقتراح (0 = ضعف الحد ناقص الرصيد وما تحت الطلب).
- **SQL Server، الاستعلامات (`docs/71`):** على الخادم فقط (ليست في ملف Access): جدول `QueryParams` (صف لكل نسخة Access مفتوحة: `HostName` + `HostProcess`، وعمود لكل معامل `QDate` / `QLong`)، والعرض `QueryParamsNow` (معاملات هذا الجهاز، صف واحد دائمًا)، والعرض `QueriesVersion`. ثم عرض لكل استعلام محفوظ في `dbo` و`en`.
- **SQL Server:** `Settings.ServerLiveSince` = منذ متى البيانات الفعلية على الخادم. يُكتب على الخادم وفي ملف Access عند أول ربط. يمنع «نقل البيانات» فوقها، ويحذّر عند الدخول جهازًا يعمل على ملف Access القديم (`docs/70`). على الخادم: كل `Execute` وكل Recordset قابل للتعديل بالخيار `dbSeeChanges`، ورقم السجل الجديد يُقرأ بعد `Update`، وحقل نعم/لا يُقارن بـ `<> 0` أو `= 0`.

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
