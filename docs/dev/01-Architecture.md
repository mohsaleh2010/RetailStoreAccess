# بنية المشروع: كيف يُبنى البرنامج

هذا الملف يشرح المشروع لأي مطوّر، أو لأي أداة ذكاء اصطناعي (Claude أو غيره). اقرأه أولًا، ثم:
- [02-Database.md](02-Database.md): قاعدة البيانات.
- [03-Forms.md](03-Forms.md): الشاشات.
- [04-Workflow.md](04-Workflow.md): طريقة العمل والتفكير.

## 1) الفكرة الأساسية
البرنامج **Microsoft Access**، لكن **لا شيء يُرسم يدويًا داخل Access**. كل شيء مكتوب كودًا في هذا المستودع:

```
tools/*.py  (التعريفات: الجداول، الاستعلامات، الشاشات، التقارير)
     │  python3 tools/generate.py
     ▼
src/vba/*.bas  (UTF-8، نسخة للقراءة)   +   dist/vba/*.bas  (Windows-1256 + CRLF، للاستيراد)
     │  dist/tools/BuildFrontEnd.vbs  (على Windows مع Access)
     ▼
RetailStore_FE.accdb  (البرنامج: الكود، الشاشات، التقارير، الاستعلامات، روابط الجداول)
RetailStore_BE.accdb  (البيانات فقط: الجداول والعلاقات)
```

- **الواجهة (FE):** تُبنى من جديد في كل تحديث، ويمكن استبدالها بأمان.
- **البيانات (BE):** تُنشأ مرة واحدة. بعد ذلك يحدّثها `BuildSchema`: يضيف الجداول والحقول الجديدة دون أن يمس البيانات الموجودة.
- **التسليم للعميل:** يُسلَّم ملف **ACCDE**، وهي واجهة مترجمة لا يمكن قراءة كودها.

## 2) خطوات البناء داخل Access (بالترتيب)
| الخطوة | ماذا تفعل | من أين تأتي |
|---|---|---|
| استيراد الوحدات | كل ملفات `dist/vba/*.bas` | المولّد، والوحدات المكتوبة يدويًا |
| Debug > Compile | ترجمة VBA. أي خطأ هنا يوقف كل شيء | — |
| `BuildSchema` | ينشئ ملف البيانات أو يحدّثه، ويضيف البيانات الأساسية، ويطلب كلمة مرور حساب المبرمج أول مرة | `tools/schema.py` |
| `BuildRelationships` | العلاقات مع فرض التكامل المرجعي | `fk=` في schema + `tools/relations.py` |
| `BuildQueries` | الاستعلامات المحفوظة | `tools/queries.py` |
| `BuildForms` | كل الشاشات مع كودها | `tools/forms*.py` |
| `BuildReports` | كل التقارير | `tools/reports*.py` |
| ترجمة ثانية | لكود الشاشات والتقارير | — |
| `LoadDemoData` / `RunAllTests` | بيانات تجريبية واختبارات داخل Access (اختياري) | `tools/demo_data.py` والوحدات |

`BuildFrontEnd.vbs` يشغّل كل هذه الخطوات، ويكتب نتيجة كل خطوة في `BuildFrontEnd.log`.

## 3) نوعان من وحدات VBA
| النوع | أمثلة | أين تعدّله |
|---|---|---|
| **مولَّدة** (أول سطر في التعليق: `GENERATED FILE`) | `modBuildSchema`، `modBuildRelations`، `modBuildQueries`، `modBuildForms`، `modBuildReports`، `modAppData`، `modDemoData`، `modQRCode`، `modTestSales`، `modTestPurchases`، `modTestSecurity` | في `tools/*.py` ثم `python3 tools/generate.py`. **لا تعدّلها يدويًا** |
| **مكتوبة يدويًا** (القائمة في `STATIC_MODULES` في `tools/generate.py`) | باقي الوحدات | في `src/vba/modX.bas`. المولّد ينسخها إلى `dist/vba` بالترميز الصحيح |

## 4) دور كل وحدة مكتوبة يدويًا
| الوحدة | الدور |
|---|---|
| `modCommon` | الأدوات المشتركة: الألوان، والرسائل، والترقيم `NextNumber`، والتقريب، و`SqlDate`/`SqlText`، والإعدادات، و`LogAction` |
| `modQueryParams` | معاملات الاستعلامات `QDate`/`QLong` و`SetPeriod` |
| `modStartup` | بدء التشغيل، ووضع المستخدم ووضع المطوّر |
| `modForms` | سلوك كل شاشات البيانات العامة: التحميل، والحفظ، والتحقق، والحذف، والقوائم، والبحث |
| `modScreens` | الشاشة الرئيسية، والبحث، ومركز التقارير |
| `modDashboard`، `modCharts`، `modIndicators` | بطاقات لوحة التحكم، والرسوم البيانية، والمؤشرات المالية (من القيود) |
| `modSales`، `modPOS`، `modTouchPOS`، `modZatca` | محرك حساب الفاتورة، ونقاط البيع (محلات، ومطاعم، وكافيهات)، ورمز QR للفاتورة الإلكترونية |
| `modPurchases`، `modPurchaseScreens` | المشتريات والمرتجعات والمخزون والجرد |
| `modReports`، `modLabels` | دوال التقارير وملصقات الباركود |
| `modSecurity`، `modSecurityScreens`، `modActivation`، `modBackup` | المستخدمون، وكلمات المرور (SHA-256)، والصلاحيات، وتفعيل الجهاز، والنسخ الاحتياطي |
| `modCash`، `modBank`، `modCheque` | الصناديق، والبنوك، والتسوية البنكية، والشيكات |
| `modJournal`، `modAccounts`، `modManualEntry`، `modLedger`، `modFinancials`، `modClosing`، `modVat` | المحاسبة: القيود الآلية، وشجرة الحسابات، والقيود اليدوية، وكشف الحساب، والقوائم المالية، والإقفال، والإقرار الضريبي |
| `modAging`، `modAssets`، `modPayroll`، `modCostCenters`، `modBudget`، `modRecurring` | أعمار الديون، والأصول والإهلاك، والرواتب، ومراكز التكلفة، والموازنة، والمصروفات المتكررة |
| `modAudit` | سجل التدقيق: من أضاف أو عدّل أو حذف، والقيم قبل وبعد |
| `modCurrency` | العملات: عملة البرنامج، وأسعار العملات، والتحويل `ToBase`/`FromBase`، وحقلا العملة والمعامل في الشاشات، وختم القيود بعملة المستند |
| `modSalesReps` | المندوبين: مندوب المستند `SalesRepFor`، ومسير العمولات الشهري (إنشاء، ترحيل، إلغاء، حذف)، وصرف العمولة بسند نقدية |
| `modZatcaXml` / `modZatcaData` | ملف الفاتورة السعودية (UBL 2.1) بالصيغة القياسية، والبصمة، والتوقيع بـ OpenSSL، وقراءة الشهادة، ورمز QR بتسعة حقول، وشاشة `frmZatcaSetup` (`docs/46`) |
| `modZatcaApi` | منصة فاتورة: تسجيل الجهاز (CSR بـ OpenSSL، ورمز OTP، وفحوص الامتثال، والشهادة الفعلية)، وإرسال المستندات للتبليغ أو الاعتماد، وقراءة ردود الهيئة (`docs/47`) |
| `modEtaReceipt` | الإيصال الإلكتروني المصري: بناء الإيصال (JSON) ومعرّفه (SHA-256 لتسلسل المصلحة) وسلسلته، والدخول بحساب جهاز نقطة البيع، والإرسال ومتابعة الحالة، وشاشة `frmEtaSetup` (`docs/48`) |
| `modEInvoice` / `modHttp` | أساس الفاتورة الإلكترونية: حالة المستند، وسجل `EInvoiceLog`، وشاشة `frmEInvoices`، والإرسال عبر منظومة الدولة بالاسم (`Application.Run`)، وHTTPS وقراءة JSON (`docs/45`) |
| `modCountry` | دولة التشغيل (السعودية أو مصر): العملة والضريبة والرقم الضريبي وعنوان المستند، وتغيير الدولة قبل أول عملية فقط (`docs/44`) |
| `modEnglishNames` | شاشة الأسماء الإنجليزية `frmEnglishNames`: الأسماء الناقصة من كل الجداول، واقتراح `Transliterate` (نسخته بـ Python `tools/translit.py`)، وحفظ مع سجل التدقيق (`docs/42`) |
| `modLang`، `modLangData1..n` (مولَّدة) | لغة الواجهة: `Tr` يترجم النص بالقاموس، و`UiAlign` و`MSG_RTL` للاتجاه، و`SetInterfaceLanguage` يختار لغة ملف الواجهة |
| `modTestAll` | `RunAllTests`: كل اختبارات Access بملخص واحد |

## 5) لماذا Python؟
- **التعريف في مكان واحد:** الجدول يُكتب مرة في `schema.py`، ومنه يُولَّد كود الإنشاء، والوثائق، وقاعدة SQLite للاختبار.
- **الاختبار بدون Access:** المنطق المحاسبي والمخزني مكرر في Python (`tools/sim.py` وملفات `*_reference.py`) ويُختبر على SQLite. استعلامات Access تُشغَّل على SQLite بعد ترجمتها (`tests/access_sqlite.py`).
- **تنفيذ VBA فعليًا:** بعض الوحدات الحسابية تُشغَّل في LibreOffice Basic عبر `tests/vba_harness.py`، مثل SHA-256 وQR وأعمار الديون وجدول المصروفات المتكررة.
- **حد هذه الاختبارات:** لا شيء يُترجم داخل Access هنا. أخطاء Access الحقيقية تُكتشف عند البناء على Windows، وتتحول كل واحدة منها إلى قاعدة في [04-Workflow.md](04-Workflow.md) وإلى اختبار يمنع تكرارها.

## 6) المراجع المولَّدة (تتحدث تلقائيًا)
| الملف | المحتوى |
|---|---|
| `docs/02-Tables-Reference.md` | كل جدول وحقل وقاعدة وفهرس |
| `docs/03-Relationships-Reference.md` | كل العلاقات ومخطط ER |
| `docs/04-Queries-Reference.md` | كل الاستعلامات |
| `docs/dev/Screens-Index.md` | كل الشاشات، ومن أين تُفتح، وجدولها، وصلاحيتها |
