# جرد واجهات Blazor (المرحلة 1 من إعادة التصميم بـ MudBlazor)

> وثيقة عمل لخطة إعادة تصميم الواجهات. لم يُعدَّل أي كود في هذه المرحلة.
> عمود «الحالة» يُحدَّث بعد كل دفعة: ⬜ لم تبدأ · 🟨 جارية · ✅ معتمدة.

## قرارات صاحب المشروع (2026-11-26)
- **اللغتان باقيتان**: `MudRTLProvider` مربوط بلغة الواجهة (عربي RTL / إنجليزي LTR)، وكل النصوص عبر `L.T`.
- **النطاق**: الويب أولًا (`RetailStore.UI` + `RetailStore.Web`)؛ `RetailStore.PosApp` و`RetailStore.RepApp` يُقرَّر فيهما بعد اكتمال الويب.
- **Git**: فرع واحد `claude/gifted-cray-56ahux` وcommit لكل دفعة.
- **العملة**: من إعدادات البلد (ريال / جنيه)؛ **التصدير**: CSV يفتح في Excel، بلا مكتبات جديدة.

## 1) البيئة الحالية

| البند | القيمة |
|---|---|
| إصدار .NET | `net10.0` لكل المشاريع؛ `RetailStore.Pos` = `net10.0-windows10.0.19041.0` (MAUI)، `RetailStore.App` = `net10.0-android` (MAUI) |
| نوع مشروع الويب | Blazor Web App (`RetailStore.Web`) والصفحات في Razor Class Library مشتركة (`RetailStore.UI`) |
| وضع العرض | `InteractiveServer` على `<Routes>` و`<HeadOutlet>` (`App.razor`)، ما عدا صفحتي الدخول وتغيير كلمة المرور: نماذج ثابتة (Static SSR) لأنها تكتب الكوكي |
| مكتبات الواجهة | **لا يوجد Bootstrap ولا أي مكتبة**: CSS يدوي فقط (`RetailStore.UI/wwwroot/retailstore.css` 212 سطرًا، `posapp.css` 65، `repapp.css` 64، `ReconnectModal.razor.css`) |
| اللغة والاتجاه | عربي RTL وإنجليزي LTR بالتبديل (`UiLanguage`: `L.Code` / `L.Dir`، كوكي، `/lang/{code}`)؛ كل نص عربي في Razor يمر بـ `L.T("...")` (يفحصه `tests/test_web.py`) |
| التنقل | `NavMenu.razor`: قائمة مسطحة (الرئيسية + 4 شاشات بيانات + 21 صفحة عمل من `WorkPages.All`) تظهر حسب صلاحية كل شاشة، بلا تجميع ولا أيقونات |
| تطبيقات أخرى بواجهات Razor | `RetailStore.PosApp` (نقطة البيع على Windows، 8 صفحات) و`RetailStore.RepApp` (تطبيق المندوب على Android، 10 صفحات)، لكل منهما Layout وCSS خاص |

## 2) جرد الصفحات — الويب (`RetailStore.UI` + `RetailStore.Web`)

| # | المسار | الملف | الموديول | النوع | الأسطر | الحالة |
|---|---|---|---|---|---|---|
| 1 | `/` | `Pages/Home.razor` | الرئيسية | D | 147 | ⬜ |
| 2 | `/account/login` | `RetailStore.Web/Components/Pages/Account/Login.razor` | إعدادات/دخول | D (SSR) | 50 | ⬜ |
| 3 | `/account/change-password` | `RetailStore.Web/Components/Pages/Account/ChangePassword.razor` | إعدادات/دخول | D (SSR) | 65 | ⬜ |
| 4 | `/products` | `Pages/ProductsPage.razor` → `Components/DataScreen.razor` | مخازن | A | 4 (+421) | ⬜ |
| 5 | `/customers` | `Pages/CustomersPage.razor` → `DataScreen` | مبيعات | A | 4 | ⬜ |
| 6 | `/suppliers` | `Pages/SuppliersPage.razor` → `DataScreen` | مشتريات | A | 4 | ⬜ |
| 7 | `/warehouses` | `Pages/WarehousesPage.razor` → `DataScreen` | مخازن | A | 4 | ⬜ |
| 8 | `/pos` | `Pages/Sales/PosPage.razor` | مبيعات | B | 361 | ⬜ |
| 9 | `/sales` | `Pages/Sales/SalesListPage.razor` | مبيعات | C (قائمة مستندات) | 100 | ⬜ |
| 10 | `/sales/{Id}` | `Pages/Sales/InvoicePage.razor` | مبيعات | B (عرض/طباعة) | 107 | ⬜ |
| 11 | `/payments` | `Pages/Sales/PaymentsPage.razor` | مبيعات | B (سند) | 192 | ⬜ |
| 12 | `/sales-returns` | `Pages/Office/SalesReturnsPage.razor` | مبيعات | B | 228 | ⬜ |
| 13 | `/purchase` | `Pages/Purchases/PurchasePage.razor` | مشتريات | B | 347 | ⬜ |
| 14 | `/purchases` | `Pages/Purchases/PurchasesListPage.razor` | مشتريات | C (قائمة مستندات) | 102 | ⬜ |
| 15 | `/purchases/{Id}` | `Pages/Purchases/PurchaseViewPage.razor` | مشتريات | B (عرض/طباعة) | 110 | ⬜ |
| 16 | `/supplier-payments` | `Pages/Purchases/SupplierPaymentsPage.razor` | مشتريات | B (سند) | 193 | ⬜ |
| 17 | `/purchase-returns` | `Pages/Office/PurchaseReturnsPage.razor` | مشتريات | B | 254 | ⬜ |
| 18 | `/stock` | `Pages/Stock/StockPage.razor` | مخازن | C (رصيد + كارت صنف) | 177 | ⬜ |
| 19 | `/transfers` | `Pages/Stock/TransfersPage.razor` | مخازن | B | 371 | ⬜ |
| 20 | `/treasury` | `Pages/Treasury/TreasuryPage.razor` | حسابات (خزينة) | C (أرصدة + كشف صندوق) | 218 | ⬜ |
| 21 | `/cash-vouchers` | `Pages/Treasury/CashVouchersPage.razor` | حسابات (خزينة) | B (سند) | 210 | ⬜ |
| 22 | `/cash-closing` | `Pages/Office/CashClosingPage.razor` | حسابات (خزينة) | B | 213 | ⬜ |
| 23 | `/bank-transactions` | `Pages/Treasury/BankTxPage.razor` | حسابات (بنوك) | B | 291 | ⬜ |
| 24 | `/bank-reconciliation` | `Pages/Office/BankReconPage.razor` | حسابات (بنوك) | B (خاصة) | 316 | ⬜ |
| 25 | `/journal` | `Pages/Accounting/JournalPage.razor` | حسابات | C (قيود + سطورها) | 171 | ⬜ |
| 26 | `/manual-entries` | `Pages/Office/ManualEntriesPage.razor` | حسابات | B | 287 | ⬜ |
| 27 | `/ledger` | `Pages/Accounting/LedgerPage.razor` | تقارير | C | 122 | ⬜ |
| 28 | `/trial-balance` | `Pages/Accounting/TrialBalancePage.razor` | تقارير | C | 95 | ⬜ |
| 29 | `/financials` | `Pages/Accounting/FinancialsPage.razor` | تقارير | C | 115 | ⬜ |
| 30 | `/reports` | `Pages/Reports/ReportsPage.razor` | تقارير | C (مركز التقارير: كل تقارير `forms.REPORTS`) | 278 | ⬜ |
| 31 | `/not-found` | `RetailStore.Web/Components/Pages/NotFound.razor` | عام | D | 6 | ⬜ |
| 32 | `/Error` | `RetailStore.Web/Components/Pages/Error.razor` | عام | D (SSR) | 22 | ⬜ |

مكونات وتخطيط الويب: `Layout/MainLayout.razor` (44)، `Layout/NavMenu.razor` (40)، `Layout/MustChangePasswordGuard.razor`، `Components/DataScreen.razor` (421، الشاشة العامة للبيانات الأساسية)، `Components/QrImage.razor`، `Components/RedirectToLogin.razor`، `RetailStore.Web/Components/Layout/AccountLayout.razor` و`ReconnectModal.razor`.

## 3) جرد الصفحات — نقطة البيع على Windows (`RetailStore.PosApp`، MAUI)

| # | المسار | الملف | الموديول | النوع | الأسطر | الحالة |
|---|---|---|---|---|---|---|
| P1 | `/login` | `Pages/LoginPage.razor` | دخول | D | 49 | ⬜ |
| P2 | `/` | `Pages/SalePage.razor` | مبيعات | B | 287 | ⬜ |
| P3 | `/sales` | `Pages/SalesPage.razor` | مبيعات | C | 63 | ⬜ |
| P4 | `/return` | `Pages/ReturnPage.razor` | مبيعات | B | 223 | ⬜ |
| P5 | `/payment` | `Pages/PaymentPage.razor` | مبيعات | B (سند) | 134 | ⬜ |
| P6 | `/closing` | `Pages/ClosingPage.razor` | خزينة | B | 241 | ⬜ |
| P7 | `/settings` | `Pages/SettingsPage.razor` | إعدادات | D | 73 | ⬜ |
| P8 | `/database` | `Pages/DatabasePage.razor` | إعدادات | D | 113 | ⬜ |

## 4) جرد الصفحات — تطبيق المندوب (`RetailStore.RepApp`، MAUI Android)

| # | المسار | الملف | الموديول | النوع | الأسطر | الحالة |
|---|---|---|---|---|---|---|
| R1 | `/login` | `Pages/LoginPage.razor` | دخول | D | 51 | ⬜ |
| R2 | `/` | `Pages/TodayPage.razor` | لوحة اليوم | D | 64 | ⬜ |
| R3 | `/customers` | `Pages/CustomersPage.razor` | مبيعات | A (قائمة فقط) | 36 | ⬜ |
| R4 | `/customers/{Id}` | `Pages/CustomerPage.razor` | مبيعات | D | 91 | ⬜ |
| R5 | `/orders` | `Pages/OrdersPage.razor` | مبيعات | C | 66 | ⬜ |
| R6 | `/orders/new/{CustomerId}` | `Pages/NewOrderPage.razor` | مبيعات | B | 115 | ⬜ |
| R7 | `/invoice/{CustomerId}` | `Pages/InvoicePage.razor` | مبيعات | B | 145 | ⬜ |
| R8 | `/payment/{CustomerId}` | `Pages/PaymentPage.razor` | مبيعات | B (سند) | 95 | ⬜ |
| R9 | `/settings` | `Pages/SettingsPage.razor` | إعدادات | D | 69 | ⬜ |
| R10 | `/password` | `Pages/PasswordPage.razor` | إعدادات | D | 36 | ⬜ |

## 5) ملاحظات

### هيكلية
- **شاشات البيانات الأساسية الأربع شاشة واحدة**: الأصناف والعملاء والموردون والمخازن كلها `DataScreen<T>` بمواصفات مولّدة (`Generated/*.g.cs` من `tools/gen_web.py`). إعادة تصميم `DataScreen` تغيّر الأربع مرة واحدة، والحقول والتحقق والترتيب تأتي من المواصفة لا من الصفحة.
- **لا يوجد Bootstrap** أصلًا؛ «إزالة Bootstrap» في الخطة = إزالة `retailstore.css` القديم بعد النقل.
- **اللغتان**: الواجهة ثنائية (عربي RTL / إنجليزي LTR)، فيجب ربط `MudRTLProvider` بـ `L.Dir` لا تثبيته على RTL.
- **صفحتا الدخول وتغيير كلمة المرور Static SSR** (تكتبان الكوكي): مكونات MudBlazor التفاعلية لا تعمل فيها؛ تحتاجان HTML عاديًا بأنماط MudBlazor أو نموذجًا خاصًا.
- **تكرار مقصود**: البيع والمرتجع والقبض والتصفية موجودة في الويب ونقطة بيع Windows وتطبيق المندوب بثلاث تنفيذات واجهة منفصلة (المنطق مشترك في `RetailStore.Data` / `PosCart` / `SaleDraft`).
- `/ledger` و`/trial-balance` تفتحان بصلاحية واحدة (`frmLedger`)، و`/sales` و`/purchases` قوائم مستندات أقرب لنوع C.

### أداء
- القوائم محدودة بعدد من السيرفر بلا ترقيم: `DataScreenService.MaxListRows = 1000` (رسالة «ضيّق البحث»)، وقوائم الفواتير والسندات والقيود والتحويلات بـ `Take(take)`. `MudDataGrid` مع `ServerData` يحتاج دالة بترقيم (skip/take/ترتيب) في الخدمات = تعديل خلفي (يُسأل عنه قبل التنفيذ).
- **مركز التقارير** (`ReportService`) يقرأ كل صفوف التقرير في الذاكرة ويعرضها في جدول واحد، والتصدير CSV بـ `data:` URI (الملف كله داخل الصفحة)؛ مع تقرير كبير قد تثقل الصفحة. الحل في الواجهة فقط: `Virtualize` / ترقيم على الذاكرة.
- `NavMenu` يستدعي `CanOpenScreenAsync` لكل شاشة (25 استعلامًا) في كل تحميل للتخطيط.

### تحقق المدخلات
- لا توجد صفة `required` ولا `EditForm`/`DataAnnotations` في أي صفحة: التحقق كله في الخدمات وبنفس رسائل VBA العربية وترتيبها (قاعدة المشروع، يفحصها `test_web`). أي تحقق في `MudForm` يجب أن يكون إضافيًا (حقول إلزامية فقط) ولا يغيّر الرسائل ولا ترتيب الفحوص.

### قيود يجب احترامها أثناء التصميم
- كل نص عربي في Razor داخل `L.T("...")` (حتى `Label=` و`Placeholder=` في MudBlazor) وله ترجمة في `tools/i18n_web.py` (`test_web` يفحص ذلك ويسرد الناقص).
- العملة من إعدادات البلد (`SA` ريال / `EG` جنيه)، لا يُكتب «جنيه» أو «ريال» ثابتًا.
- `RetailStore.UI` مكتبة مشتركة مع MAUI لاحقًا؛ و`RetailStore.Pos` / `RetailStore.App` لا يُبنيان في السحابة (يحتاجان Windows / workload).
- .NET 10 SDK موجود في الحاوية (10.0.115) ويُبنى به الويب ويُشغَّل للقطات.

## 6) المرحلة 2: نظام التصميم والتخطيط (✅ تمت الموافقة 2026-11-26)
- **MudBlazor 9.14.0** في `RetailStore.UI`؛ `AddMudServices` في `RetailStore.Web/Program.cs` (الرسائل أسفل اليسار، 3 ثوانٍ).
- **`Theme/ErpTheme.cs`**: ألوان الخطة مع تعديلات التباين: Secondary `#0B7A6B`، Warning `#B45309`، نص غامق على Tertiary، ونص غامق على Primary (وSecondary) في الوضع الداكن؛ خط **Cairo** (400/600/700 من Google Fonts)؛ زوايا 8px؛ القائمة 260px.
- **`Layout/MainLayout.razor`**: `MudRTLProvider` مربوط بـ `L.Dir` (عربي RTL / إنجليزي LTR)، `MudThemeProvider` بالوضع الليلي من كوكي `theme` (يقرؤها `App.razor` ويمررها `Routes` مثل اللغة، ويكتبها `erp.js` عند التبديل مع `data-theme` على `<html>`)، شريط علوي: القائمة، اسم المحل، **بحث عن شاشة** (`MudAutocomplete` على شاشات المستخدم)، الوضع الليلي، اللغة، قائمة المستخدم (تغيير كلمة المرور، الخروج). القائمة الجانبية `Responsive`: ثابتة على الكمبيوتر ومنبثقة على الموبايل.
- **`Layout/NavGroups.cs`**: المبيعات، المشتريات، المخازن، الخزينة والبنوك، الحسابات، التقارير، بأيقونات Material؛ الشاشات وصلاحياتها كما هي من `ScreenSpecs` و`WorkPages` (صلاحية كل نموذج تُسأل مرة واحدة)، وأي مسار جديد غير مصنّف يظهر في آخر مجموعة.
- **`wwwroot/erp.css`**: الخط، أرقام `tabular-nums`، الشريط والقائمة، الطباعة؛ ومتغيرات `retailstore.css` مربوطة بألوان الثيم فتتبع الصفحات القديمة الوضع الليلي إلى أن يُعاد تصميمها (`retailstore.css` صار يتبع `data-theme` لا إعداد نظام التشغيل).
- صفحتا الدخول وتغيير كلمة المرور (SSR) لم تتغيرا بعد سوى الخط والوضع الليلي.

## 7) المرحلة 3: المكونات المشتركة واختصارات الكيبورد (✅ تمت الموافقة 2026-11-26)
المكونات في `RetailStore.UI/Components/Shared` (namespace `RetailStore.UI.Components.Shared`، مضافة إلى `_Imports.razor` في المشروعين). كل مكون يبدأ بتعليق قصير فيه مثال استخدام، والنصوص الممرَّرة له تمر بـ `L.T` عند صاحب الصفحة، ونصوصه الداخلية بـ `L.T` داخله.

| المكون | الوظيفة والقرار |
|---|---|
| `PageHeader` | العنوان (`h1` الذي يركّز عليه الموجّه) و`MudBreadcrumbs` (الرئيسية / المجموعة / العنوان عبر `Group`، أو `Breadcrumbs` كاملة)، والوصف، و`Status` (شارة الحالة) و`Actions` (الأزرار)، و`PageTitle` |
| `EntityGrid<T>` | غلاف `MudDataGrid`: Dense + Hover + Striped، بحث سريع (`SearchText`، يطوي أشكال الألف والتاء المربوطة والياء)، فلاتر الأعمدة والترتيب المتعدد، صفحات 10/25/50/100، تصدير CSV يفتح في Excel (`Export` + `GridCsv`: UTF-8 بـ BOM، والنص الذي يبدأ بـ `= + - @` يبقى نصًا)، شريط تحميل، و`EmptyState` عند عدم وجود بيانات. `ServerData` مُمرَّر لـ `MudDataGrid` لكن **الترقيم من الخادم مؤجّل للمرحلة 7** (الخدمات لا ترقّم بعد) |
| `EntityDialog` | قالب `MudDialog` بـ `MudForm`: الحقول الإلزامية تُفحص أولًا، ثم `OnSave` يعيد "" أو رسالة الخدمة نفسها (بلا تغيير في الفحوص ولا ترتيبها)، وزرا حفظ / إلغاء ثابتان، F2 يحفظ وEsc يغلق (`EntityDialog.Options`) |
| `LookupField<T>` | `MudAutocomplete` بالكود أو الاسم: الكود المطابق أولًا ثم الأكواد التي تبدأ بالنص ثم الأسماء (`LookupMatch`)، أو `SearchFunc` لخدمة؛ `InfoTemplate` لعرض الرصيد أو السعر بجوار الاسم |
| `InvoiceLinesEditor<TProduct>` | الصنف والكمية والسعر والخصم والضريبة؛ مبالغ كل سطر والإجماليات من **محرك الفواتير نفسه** `InvoiceCalculator.Calc` (= `modSales.CalcRun`) لا من الشاشة، ورسالته تظهر عند خطأ سطر؛ Enter للخانة التالية وبعد آخر خانة سطر جديد، وDelete في خانة فارغة (أو Ctrl+Delete) يحذف السطر، واختيار الصنف ينقل للكمية |
| `FilterPanel` | لوحة قابلة للطي: الفترة `MudDateRangePicker` (ميلادي دائمًا: `ErpFormat.DateCulture`، لأن ثقافة ar-SA هجرية)، واختيارات التقرير (`ChildContent`)، وزرا «عرض» و«مسح»؛ `FilterPanel.ThisMonth()` الفترة الافتراضية |
| `MoneyText` / `QtyText` | خانتان عشريتان وفواصل الآلاف وأرقام لاتينية، والسالب أحمر، و`ShowCurrency` يضيف عملة البرنامج (`UiCurrency`: `Settings.CurrencyCode`، وإلا عملة البلد: ريال / جنيه)؛ الكمية حتى 3 خانات مع الوحدة |
| `StatusChip` | لون وأيقونة ونص لكل حالة (`StatusStyles`): DRAFT، POSTED، APPROVED، PREPARING، PARTIAL، PARTIAL_PAID، PAID، UNPAID، DELIVERED، RECEIVED، INVOICED، IN_TRANSIT، PENDING، OPEN، CLOSED، EXPIRED، REJECTED، CANCELLED؛ غير المعروفة تظهر بكودها |
| `ConfirmDialog` | `ConfirmDialog.AskAsync(...)`: عنوان ورسالة واضحة وزر بلون الإجراء (أحمر للحذف، Primary للترحيل)؛ Esc أو الخلفية = لا |
| `StatCard` | رقم كبير + عنوان + أيقونة بلون الثيم + نسبة التغيّر (سهم أخضر / أحمر، و`DownIsGood` للمصروفات) أو سطر وصف، و`Href` يفتح شاشة |
| `EmptyState` | أيقونة ونص وزر إجراء |
| `KeyboardShortcuts` | يسجّل اختصارات الصفحة في `ShortcutService` |

**خدمة الاختصارات** (`Services/ShortcutService.cs` + `erp.js`): مستمع واحد في المتصفح لكل جلسة، وكومة معالجات: آخر من سجّل المفتاح يأخذه (النافذة المفتوحة تأخذ F2 من الصفحة وتعيده عند إغلاقها). **F2** حفظ، **F3** بحث (بلا معالج: يذهب لمربع بحث الصفحة `data-erp-search`)، **F4** جديد، **Esc** إغلاق (النافذة المفتوحة تُغلق نفسها: `MudDialogProvider CloseOnEscapeKey`)، **Ctrl+P** طباعة (بلا معالج: طباعة المتصفح). المتصفح لا يمنع إلا المفاتيح التي لها معالج.

**نصوص MudBlazor** (عدد الصفحات، فلاتر الأعمدة، «مطلوب»، التواريخ): بالعربية في الواجهة العربية عبر `Services/ErpMudLocalizer.cs`، وبالإنجليزية الأصلية في الإنجليزية.

**صفحة العرض** `/dev/components` (`RetailStore.Web/Components/Pages/Dev`): كل المكونات ببيانات تجريبية من الصفحة نفسها (لا من القاعدة)، في بيئة Development فقط (غيرها: «الصفحة غير موجودة»)، وليست في القائمة.

**الاختبارات:** `tests/test_web_components.py` (وجود المكونات وأمثلتها، لا ألوان hex، خصائص منطقية للاتجاهين، العملة من الإعدادات، الأسطر من المحرك، الاختصارات، صفحة العرض في Development فقط) و`ComponentsTests.cs` في xUnit (التنسيق، البحث، CSV، الأسطر مقابل `InvoiceCalculator`، الحالات، كومة الاختصارات، النصوص العربية لـ MudBlazor).

**مشكلة وُجدت وحُلّت:** جدول الأسطر الأعرض من الشاشة في RTL كان يوسّع الصفحة كلها على 768px (تمرير أفقي)؛ الحل `contain: paint` على حاوية الجدول فيُمرَّر داخلها فقط. وفي الوضع الداكن كانت الصفوف المخططة فاتحة جدًا؛ صار لها `TableStriped` / `TableHover` في `ErpTheme`.
