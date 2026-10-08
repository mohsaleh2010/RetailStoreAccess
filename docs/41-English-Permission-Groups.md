# أسماء المجموعات بالإنجليزية في شاشتي الصلاحيات

**الفكرة:** في شاشتي **الأدوار والصلاحيات** و**صلاحيات الشاشات للمستخدمين** عمود «القسم» يجمع الصلاحيات والشاشات (المبيعات، الحسابات…). بعد `docs/39` صارت أسماء الصلاحيات والشاشات إنجليزية في الواجهة الإنجليزية، وبقي اسم المجموعة عربيًا. صار له اسم إنجليزي بنفس طريقة `docs/39`.

## 1) الحقول الجديدة (يضيفها `BuildSchema`)
| الجدول | الحقل الجديد |
|---|---|
| `Permissions` (الصلاحيات) | `ModuleNameEn` |
| `Screens` (الشاشات) | `ModuleNameEn` |

`BuildSchema` يملؤه للمجموعات العشر، ويملأ الحقل الفارغ فقط:

| بالعربية | بالإنجليزية |
|---|---|
| المبيعات | Sales |
| العملاء | Customers |
| المشتريات | Purchases |
| الموردون | Suppliers |
| المخزون | Inventory |
| المصروفات | Expenses |
| الخزينة | Treasury |
| الحسابات | Accounting |
| التقارير | Reports |
| النظام | System |

**الواجهة العربية** لا تتغير. **الواجهة الإنجليزية** تعرض الاسم الإنجليزي، وإن كان فارغًا تعرض العربي. ترتيب الصفوف في الشاشتين لم يتغير.

## 2) كيف يعمل (للمطوّر)
- الترجمات في `tools/master_en.py`: `MODULE_NAMES_EN`، و`EXTRA_NAMES`. هذا الأخير يحمل أسماء إضافية لجدول من `ENGLISH_NAMES`، مرتبة **بالقيمة العربية**، لا بالمفتاح.
- `qryLocPermissions` و`qryLocScreens` صار فيهما عمود `ModuleName` بالاسم الإنجليزي أيضًا. الخطوة الأولى تسمّي العمود العربي `LocArabicModuleName`، حتى لا يساوي اسم العمود حقلًا في تعبيره.
- الشاشتان تقرآن `[@Permissions]` و`[@Screens]` من قبل (`modSecurityScreens`)، فلم يتغير فيهما كود.
- `SeedEnglishNames` في `modBuildSchema`، مثلًا:
  `UPDATE Permissions SET ModuleNameEn = 'Sales' WHERE ModuleName = 'المبيعات' AND ModuleNameEn Is Null`

## 3) الاختبارات
`tests/test_master_en.py`:
- لكل مجموعة في الصلاحيات والشاشات الجاهزة اسم إنجليزي.
- `BuildSchema` يملأ الفارغ فقط.
- المجموعة بالإنجليزية في الواجهة الإنجليزية على نسخة SQLite، وبالعربية في العربية.
- كل الاستعلامات المحفوظة تعمل في نسختها الإنجليزية.

## 4) التثبيت والاختبار
1. استبدل الوحدتين المولَّدتين `modBuildSchema` و`modBuildQueries` (الصف 38 في `README.md`)، أو أعد البناء بـ `BuildFrontEnd.vbs`.
2. شغّل `BuildSchema` ثم `BuildQueries`، في كل ملف واجهة. لا حاجة لـ `BuildForms`.
3. شغّل `RunAllTests`.
4. **في الملف الإنجليزي:** افتح «Roles and permissions» ثم «Screen permissions of users». عمود القسم يعرض Sales وAccounting وSystem…
5. **في الملف العربي:** كما كان.
