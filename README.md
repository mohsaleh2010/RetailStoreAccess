# RetailStoreAccess – نظام إدارة محل تجاري (Microsoft Access)

نظام متكامل لإدارة محل تجاري: المنتجات، المخزون، المبيعات (POS)، المشتريات، العملاء، الموردون، المصروفات، الأرباح، التقارير، الصلاحيات، والنسخ الاحتياطي، مع دعم متطلبات ضريبة القيمة المضافة والفوترة الإلكترونية (فاتورة) في السعودية.

## العمل على المشروع من Claude Desktop

افتح المجلد في تبويب **Code** في Claude Desktop. التجهيز والأوامر وطريقة العمل في [docs/00-Claude-Desktop.md](docs/00-Claude-Desktop.md).

## للمطوّرين ولأي ذكاء اصطناعي آخر

| الملف | المحتوى |
|---|---|
| [AGENTS.md](AGENTS.md) | تعليمات أي أداة ذكاء اصطناعي، مثل Claude وChatGPT/Codex وGemini وCursor وCopilot: القواعد والأوامر وخطوات إضافة ميزة. Claude يقرؤها عبر [CLAUDE.md](CLAUDE.md) |
| [docs/dev/01-Architecture.md](docs/dev/01-Architecture.md) | بنية المشروع: كيف يولّد Python كود Access، وخطوات البناء، ودور كل وحدة |
| [docs/dev/02-Database.md](docs/dev/02-Database.md) | قاعدة البيانات: المبادئ، والجداول حسب المجال، والنموذج المحاسبي، والصلاحيات، وحدود Access |
| [docs/dev/03-Forms.md](docs/dev/03-Forms.md) | الشاشات: أنواعها، وخاصية Tag، والأحداث العامة، والتسمية، وقواعد الشكل |
| [docs/dev/04-Workflow.md](docs/dev/04-Workflow.md) | طريقة العمل والتفكير، وكل قواعد Access مع سبب كل قاعدة، وتشخيص الأخطاء |
| [docs/dev/Screens-Index.md](docs/dev/Screens-Index.md) | فهرس كل الشاشات (مولَّد) |

## مراحل التنفيذ

| المرحلة | الموضوع | الحالة | الملف |
|---|---|---|---|
| 1 | تحليل النظام | ✅ تمت الموافقة | [docs/01-System-Analysis.md](docs/01-System-Analysis.md) |
| 2 | تصميم الجداول | ✅ تمت الموافقة | [docs/02-Table-Design.md](docs/02-Table-Design.md) · [مرجع الجداول](docs/02-Tables-Reference.md) |
| 3 | العلاقات | ✅ تمت الموافقة | [docs/03-Relationships.md](docs/03-Relationships.md) · [مرجع العلاقات](docs/03-Relationships-Reference.md) |
| 4 | الاستعلامات | ✅ تمت الموافقة | [docs/04-Queries.md](docs/04-Queries.md) · [مرجع الاستعلامات](docs/04-Queries-Reference.md) |
| 5 | النماذج | ✅ تمت الموافقة | [docs/05-Forms.md](docs/05-Forms.md) |
| 6 | نظام المبيعات | ✅ تمت الموافقة | [docs/06-Sales.md](docs/06-Sales.md) |
| 7 | المشتريات والمخزون | ✅ تمت الموافقة | [docs/07-Purchases-Inventory.md](docs/07-Purchases-Inventory.md) |
| 8 | التقارير | ✅ تمت الموافقة | [docs/08-Reports.md](docs/08-Reports.md) |
| 9 | لوحة التحكم | ✅ تمت الموافقة | [docs/09-Dashboard.md](docs/09-Dashboard.md) |
| 10 | المستخدمون والصلاحيات والنسخ الاحتياطي | ✅ تمت الموافقة | [docs/10-Security.md](docs/10-Security.md) |
| 11 | الاختبار الشامل والبيانات التجريبية | ✅ تمت الموافقة | [docs/11-Testing-Demo.md](docs/11-Testing-Demo.md) |
| 12 | دليل الاستخدام | ✅ تمت الموافقة | [docs/12-User-Guide.md](docs/12-User-Guide.md) |
| + | طباعة ملصقات الباركود | ✅ بانتظار الموافقة | [docs/13-Barcode-Labels.md](docs/13-Barcode-Labels.md) |
| + | لوحة التحكم الجديدة والإحصائيات بالرسوم البيانية | ✅ بانتظار الموافقة | [docs/14-Dashboard-Charts.md](docs/14-Dashboard-Charts.md) |
| + | نقطة بيع المطاعم (شاشة لمس) | ✅ تمت الموافقة | [docs/15-Restaurant-POS.md](docs/15-Restaurant-POS.md) |
| + | نقطة بيع الكافيهات (شاشة لمس) | ✅ بانتظار الموافقة | [docs/16-Cafe-POS.md](docs/16-Cafe-POS.md) |
| + | الخزينة والصناديق وتصفية الكاشير وتقارير المصروفات | ✅ بانتظار الموافقة | [docs/17-Treasury.md](docs/17-Treasury.md) |
| + | قيود اليومية ودليل الحسابات وميزان المراجعة | ✅ بانتظار الموافقة | [docs/18-Journal.md](docs/18-Journal.md) |
| + | صلاحيات الشاشات لكل مستخدم، وحساب المبرمج، وتفعيل البرنامج على جهاز محدد | ✅ بانتظار الموافقة | [docs/19-Permissions-Activation.md](docs/19-Permissions-Activation.md) |
| + | شجرة الحسابات (5 مستويات) والقيود اليدوية وميزان المراجعة بالمستويات | ✅ بانتظار الموافقة | [docs/20-Accounts-Manual-Entries.md](docs/20-Accounts-Manual-Entries.md) |
| + | كشف الحساب ودفتر الأستاذ | ✅ تمت الموافقة | [docs/21-Ledger.md](docs/21-Ledger.md) |
| + | القوائم المالية: قائمة الدخل والميزانية العمومية مع المقارنة | ✅ تمت الموافقة | [docs/22-Financial-Statements.md](docs/22-Financial-Statements.md) |
| + | إقفال الفترات وإقفال السنة المالية | ✅ تمت الموافقة | [docs/23-Period-Closing.md](docs/23-Period-Closing.md) |
| + | إقرار ضريبة القيمة المضافة: خانات نموذج الهيئة، الاعتماد وقيد التسوية، السداد | ✅ تمت الموافقة | [docs/24-VAT-Return.md](docs/24-VAT-Return.md) |
| + | أعمار الديون، تاريخ الاستحقاق، ربط السداد بالفواتير، إيقاف الآجل للمتأخرين | ✅ تمت الموافقة | [docs/25-Aging.md](docs/25-Aging.md) |
| + | البنوك، تسوية مدى بعمولتها، الإيداع والسحب، التسوية البنكية | ✅ تمت الموافقة | [docs/26-Banks.md](docs/26-Banks.md) |
| + | الشيكات الواردة والصادرة: تحت التحصيل، محصَّل، مرتد | ✅ تمت الموافقة | [docs/27-Cheques.md](docs/27-Cheques.md) |
| + | الأصول الثابتة والإهلاك الشهري والبيع أو الاستبعاد | ✅ تمت الموافقة | [docs/28-Fixed-Assets.md](docs/28-Fixed-Assets.md) |
| + | الرواتب: مسير شهري، التأمينات، خصم السلف، القيد والصرف | ✅ تمت الموافقة | [docs/29-Payroll.md](docs/29-Payroll.md) |
| + | مراكز التكلفة والفروع: توزيع القيود وقائمة دخل لكل مركز | ✅ تمت الموافقة | [docs/30-Cost-Centers.md](docs/30-Cost-Centers.md) |
| + | الموازنة التقديرية: شهرية لكل حساب ومركز، والمقارنة بالفعلي والانحراف | ✅ تمت الموافقة | [docs/31-Budget.md](docs/31-Budget.md) |
| + | المصروفات المتكررة، وشاشة «المحاسبة والمالية» لكل شاشات الحسابات | ✅ تمت الموافقة | [docs/32-Recurring-Expenses.md](docs/32-Recurring-Expenses.md) |
| + | سجل التدقيق: من أضاف أو عدّل أو حذف، والقيم قبل وبعد | ✅ تمت الموافقة | [docs/33-Audit-Trail.md](docs/33-Audit-Trail.md) |
| + | المؤشرات المالية في لوحة التحكم: هامش الربح، دوران المخزون، فترة التحصيل، السيولة | ✅ بانتظار الموافقة | [docs/34-Financial-Indicators.md](docs/34-Financial-Indicators.md) |
| + | خطة العملات المتعددة والمندوبين والواجهة الإنجليزية | 📋 خطة | [docs/35-Plan-Currency-Reps-Language.md](docs/35-Plan-Currency-Reps-Language.md) |
| + | العملات المتعددة: عملة ومعامل لكل مستند، والترحيل بالمكافئ بالريال | ✅ بانتظار الموافقة | [docs/36-Currencies.md](docs/36-Currencies.md) |
| + | المندوبين: عملاء المندوب، والأهداف الشهرية، ومسير العمولات وقيده وصرفه، وتقارير الأداء | ✅ بانتظار الموافقة | [docs/37-Sales-Reps.md](docs/37-Sales-Reps.md) |
| + | الواجهة الإنجليزية: ملف واجهة بالإنجليزية من اليسار لليمين على نفس البيانات، والرسائل والتقارير مترجمة | ✅ بانتظار الموافقة | [docs/38-English-Interface.md](docs/38-English-Interface.md) |

## هيكل المستودع

| المسار | المحتوى |
|---|---|
| `dist/vba/` | ملفات VBA الجاهزة للاستيراد في Access (Windows-1256) |
| `src/vba/` | نفس الملفات بترميز UTF-8 للقراءة |
| `docs/` | توثيق كل مرحلة |
| `tools/schema.py` | تعريف الجداول (المصدر الوحيد) |
| `tools/relations.py` | العلاقات المشتقة من الجداول وسيناريو اختبارها |
| `tools/queries.py` | الاستعلامات وبيانات الاختبار والنتائج المتوقعة |
| `tools/forms.py` | تعريف الشاشات وتخطيطها، قوالب البحث، قائمة التقارير |
| `tools/forms_sales.py`, `tools/reports.py` | شاشات المبيعات، وتقارير الفاتورة |
| `tools/reports_catalog.py`, `tools/reports_docs.py`, `tools/tafqeet.py` | التقارير المنسقة لمركز التقارير، والمستندات، ومرجع المبلغ بالحروف |
| `tools/forms_labels.py`, `tools/barcode_reference.py` | شاشة ملصقات الباركود، ومرجع ترميز EAN-13 و Code 128 |
| `tools/demo_data.py`, `tools/sim.py` | البيانات التجريبية وخطتها، وإعادة تنفيذ دوال الحفظ بـ Python للتحقق |
| `tools/forms_security.py`, `tools/security_reference.py` | شاشات الدخول والمستخدمين والصلاحيات والنسخ، ومرجع SHA-256 والنسخ |
| `dist/tools/BuildFrontEnd.vbs` | **بناء ملف الواجهة `RetailStore_FE.accdb` جاهزًا بنقرة مزدوجة**: يستورد كل الوحدات ويترجمها ويشغّل أوامر البناء |
| `dist/tools/EnableShiftKey.vbs` | إعادة تفعيل مفتاح Shift إذا تعذر دخول المدير |
| `tools/forms_purchases.py`, `tools/purchases_reference.py` | شاشات المشتريات والمخزون والجرد، وسيناريو المرحلة 7 بنتائجه المتوقعة |
| `tools/pricing.py`, `tools/zatca_reference.py`, `tools/qr_reference.py` | المراجع الحسابية: الفاتورة، حمولة QR للهيئة، مولّد QR |
| `tests/vba_harness.py` | تشغيل كود VBA الحسابي فعليًا عبر LibreOffice للتحقق منه |
| `tools/generate.py` | يولّد الكود والمرجع: `python3 tools/generate.py` |
| `tests/` | الاختبارات الآلية: `python3 -m unittest discover -s tests -v` |

## الطريقة الأسرع: بناء ملف الواجهة تلقائيًا
على جهاز Windows عليه Microsoft Access 2010 أو أحدث:
1. حمّل المستودع كاملًا (Code ← Download ZIP) وفك الضغط.
2. انقر نقرًا مزدوجًا على `dist\tools\BuildFrontEnd.vbs` واختر مكان الملف، والافتراضي `dist\RetailStore_FE.accdb`.
3. السكربت يقوم بما يلي:
   - ينشئ ملفًا جديدًا ويستورد **كل** الوحدات من `dist\vba`.
   - يترجمها بأمر Debug ← Compile.
   - يشغّل `BuildSchema` ← `BuildRelationships` ← `BuildQueries` ← `BuildForms` ← `BuildReports`، ثم يترجم مرة ثانية.
   - ينشئ ملف البيانات `RetailStore_BE.accdb` بجانب الواجهة، أو يحدّثه إن كان موجودًا مع الحفاظ على بياناته.
   - بعد كل خطوة يعرض Access رسالة: اضغط **موافق**.
4. في النهاية يسألك:
   - هل تريد تحميل البيانات التجريبية (للتدريب فقط)؟
   - هل تريد تشغيل كل الاختبارات؟
   - هل تريد التحويل إلى وضع المستخدم النهائي؟
5. كل خطوة تُكتب في `BuildFrontEnd.log` بجانب الملف. إذا فشلت خطوة، أرسل لي هذا الملف ونص الرسالة.

> **شرط اللغة العربية:** في إعدادات Windows يجب أن تكون «لغة البرامج غير الداعمة لـ Unicode» العربية.
>
> **عند التحديث لاحقًا:** شغّل السكربت مرة أخرى ووافق على استبدال الواجهة. ملف البيانات لا يُحذف.
>
> **أول دخول:** `admin` بدون كلمة مرور.

## ترتيب التثبيت في Access (يدويًا)

| # | الملف المستورد | الأمر في نافذة Immediate | الفحص |
|---|---|---|---|
| 1 | `dist/vba/modBuildSchema.bas` | `BuildSchema` | `VerifySchema` |
| 2 | `dist/vba/modBuildRelations.bas` | `BuildRelationships` | `TestRelationships` |
| 3 | `dist/vba/modQueryParams.bas` (دائمة) ثم `dist/vba/modBuildQueries.bas` | `BuildQueries` | `TestQueries` |
| 4 | `modCommon`, `modStartup`, `modForms`, `modScreens`, `modAppData` (دائمة) ثم `modBuildForms` | `BuildForms` | `TestForms` |
| 5 | `modZatca`, `modQRCode`, `modSales`, `modPOS` (دائمة) ثم `modBuildReports`, `modTestSales` | `BuildQueries`, `BuildForms`, `BuildReports` | `TestSales` |
| 6 | `modPurchases`, `modPurchaseScreens` (دائمة) ثم `modTestPurchases` | `BuildQueries`, `BuildForms` | `TestPurchases` |
| 7 | `modReports` (دائمة) ثم `modBuildReports` | `BuildQueries`, `BuildForms`, `BuildReports` | `TestReports` |
| 8 | `modDashboard` (دائمة) | `BuildQueries`, `BuildForms` | `TestDashboard` |
| 9 | `modSecurity`, `modSecurityScreens`, `modBackup` (دائمة) ثم `modTestSecurity` | `BuildForms`, `BuildReports` | `TestSecurity` |
| 10 | `modTestAll`, `modDemoData` | `LoadDemoData` (اختياري، للتدريب) | **`RunAllTests`** (كل الاختبارات) |
| 11 | `modLabels` (دائمة) | `BuildSchema`, `BuildForms`, `BuildReports` | `TestLabels` |
| 12 | `modCharts` (دائمة) | `BuildQueries`, `BuildForms`, `BuildReports` | `TestDashboard`, `TestReports` |
| 13 | `modTouchPOS` (دائمة) | `BuildSchema` (يضيف الحقول الناقصة), `BuildQueries`, `BuildForms`, `BuildReports` | `TestTouchPOS` |
| 14 | `modCash` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestCash` |
| 15 | `modJournal` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestJournal` |
| 16 | `modActivation` (دائمة) | `BuildSchema` (يطلب كلمة مرور المبرمج), `BuildRelations`, `BuildForms` | `TestSecurity` |
| 17 | `modAccounts`, `modManualEntry` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestJournal` |
| 18 | `modLedger` (دائمة) | `BuildSchema`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestJournal` |
| 19 | `modFinancials` (دائمة) | `BuildSchema`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestJournal` |
| 20 | `modClosing` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms` | `TestJournal` |
| 21 | `modVat` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestJournal`, `TestQueries` |
| 22 | `modAging` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestAging` |
| 23 | `modBank` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms` | `TestBank` |
| 24 | `modCheque` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms` | `TestCheques` |
| 25 | `modAssets` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestAssets` |
| 26 | `modPayroll` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestPayroll` |
| 27 | `modCostCenters` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestCostCenters` |
| 28 | `modBudget` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestBudget` |
| 29 | `modRecurring` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildForms` | `TestRecurring` |
| 30 | `modAudit` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestAudit` |
| 31 | `modIndicators` (دائمة) | `BuildQueries`, `BuildForms` | `TestIndicators` |
| 32 | `modCurrency` (دائمة) | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestCurrency` |
| 33 | `modSalesReps` (دائمة)، واستبدال `modSales` و`modCash` و`modJournal` و`modForms` و`modScreens` | `BuildSchema`, `BuildRelations`, `BuildQueries`, `BuildForms`, `BuildReports` | `TestSalesReps` |
| 34 | `modLang` و`modLangData1..4` (مولَّدة، جديدة)، واستبدال كل الوحدات (الأسهل: `BuildFrontEnd.vbs`) | `BuildQueries`, `BuildForms`, `BuildReports` (وللإنجليزية قبلها `SetInterfaceLanguage "EN"`) | `TestLang` |

> عند تحديث وحدة موجودة: احذفها أولًا من محرر VBA ثم استورد النسخة الجديدة.
>
> **من المرحلة 10:** البرنامج يتطلب تسجيل الدخول. أول دخول `admin` بدون كلمة مرور، ثم يُطلب تعيينها.

### الاختبارات الآلية
```
python3 tools/generate.py
python3 -m unittest discover -s tests -v
```
اختبارات `test_vba_runtime` و `test_purchases` و `test_reports` و `test_security` تشغّل كود VBA الحسابي فعليًا وتحتاج LibreOffice (`soffice` و `python3-uno`)، وتُتخطى تلقائيًا إن لم يكن مثبتًا.
