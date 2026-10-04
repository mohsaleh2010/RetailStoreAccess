# RetailStoreAccess – نظام إدارة محل تجاري (Microsoft Access)

نظام متكامل لإدارة محل تجاري: المنتجات، المخزون، المبيعات (POS)، المشتريات، العملاء، الموردون، المصروفات، الأرباح، التقارير، الصلاحيات، والنسخ الاحتياطي، مع دعم متطلبات ضريبة القيمة المضافة والفوترة الإلكترونية (فاتورة) في السعودية.

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
| 12 | دليل الاستخدام | ✅ بانتظار الموافقة | [docs/12-User-Guide.md](docs/12-User-Guide.md) |

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
| `tools/demo_data.py`, `tools/sim.py` | البيانات التجريبية وخطتها، وإعادة تنفيذ دوال الحفظ بـ Python للتحقق |
| `tools/forms_security.py`, `tools/security_reference.py` | شاشات الدخول والمستخدمين والصلاحيات والنسخ، ومرجع SHA-256 والنسخ |
| `dist/tools/EnableShiftKey.vbs` | إعادة تفعيل مفتاح Shift إذا تعذر دخول المدير |
| `tools/forms_purchases.py`, `tools/purchases_reference.py` | شاشات المشتريات والمخزون والجرد، وسيناريو المرحلة 7 بنتائجه المتوقعة |
| `tools/pricing.py`, `tools/zatca_reference.py`, `tools/qr_reference.py` | المراجع الحسابية: الفاتورة، حمولة QR للهيئة، مولّد QR |
| `tests/vba_harness.py` | تشغيل كود VBA الحسابي فعليًا عبر LibreOffice للتحقق منه |
| `tools/generate.py` | يولّد الكود والمرجع: `python3 tools/generate.py` |
| `tests/` | الاختبارات الآلية: `python3 -m unittest discover -s tests -v` |

## ترتيب التثبيت في Access

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

> عند تحديث وحدة موجودة: احذفها أولًا من محرر VBA ثم استورد النسخة الجديدة.
>
> **من المرحلة 10:** البرنامج يتطلب تسجيل الدخول. أول دخول `admin` بدون كلمة مرور، ثم يُطلب تعيينها.

### الاختبارات الآلية
```
python3 tools/generate.py
python3 -m unittest discover -s tests -v
```
اختبارات `test_vba_runtime` و `test_purchases` و `test_reports` و `test_security` تشغّل كود VBA الحسابي فعليًا وتحتاج LibreOffice (`soffice` و `python3-uno`)، وتُتخطى تلقائيًا إن لم يكن مثبتًا.
