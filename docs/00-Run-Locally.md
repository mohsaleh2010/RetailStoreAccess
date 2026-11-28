# تنزيل المشروع وتشغيل كل تطبيق فيه على جهاز Windows

دليل خطوة بخطوة لصاحب المشروع: تنزيل المستودع كله، ثم تشغيل كل جزء منه وحده. الأوامر تُكتب في **PowerShell**، والمسارات على أن المشروع في `D:\RetailStore\RetailStoreAccess` (غيّرها إن نزّلته في مكان آخر).

| الجزء | ما هو | يحتاج | القسم |
|---|---|---|---|
| Access | البرنامج الأصلي للمكتب (`RetailStore_FE.accdb` + `RetailStore_BE.accdb`) | Microsoft Access 2010 أو أحدث | 3 |
| الموقع | نظام الويب والـ API (`RetailStore.Web`) | .NET 10 SDK | 4 |
| نقطة البيع | برنامج الكاشير على Windows (`RetailStore.Pos`) | .NET 10 + أداة `maui-windows` | 5 |
| تطبيق المندوب | تطبيق Android (`RetailStore.App`) | .NET 10 + أداة `maui-android` + Java 17 | 6 |
| خدمة مهام المكتب | خدمة Windows للقيود والفاتورة الإلكترونية (`RetailStore.OfficeService`) | .NET 10 + قاعدة SQL Server | 7 |
| المعاينة | صفحات نقطة البيع والمندوب في المتصفح، للتطوير فقط (`RetailStore.AppPreview`) | .NET 10 | 8 |

## 1) ما يُثبَّت مرة واحدة
1. **Git:** من git-scm.com. للتأكد: `git --version`.
2. **.NET 10 SDK:** من dotnet.microsoft.com، أو مع Visual Studio 2026. للتأكد: `dotnet --version` يبدأ بـ `10.`.
3. **أدوات MAUI** (لنقطة البيع وتطبيق المندوب فقط)، من Visual Studio Installer («.NET Multi-platform App UI development») أو:
   ```bash
   dotnet workload install maui-windows maui-android
   ```
   للتأكد: `dotnet workload list` يذكر `maui-windows` و`android`.
4. **Python 3** (لمولّد الكود والاختبارات فقط): من python.org مع «Add python.exe to PATH». للتأكد: `python --version`.
5. **Microsoft Access 2010 أو أحدث** لجزء Access، و**لغة البرامج غير الداعمة لـ Unicode = العربية** في إعدادات Windows (Region ← Administrative).
6. **Java 17 أو أحدث** لبناء تطبيق Android (يُثبَّت مع أداة MAUI لـ Android في Visual Studio).

## 2) تنزيل المشروع وتحديثه
**أول مرة** (المستودع خاص: تفتح نافذة دخول GitHub، ادخل بحسابك):
```bash
cd D:\RetailStore
```
```bash
git clone -b claude/gifted-cray-56ahux https://github.com/mohsaleh2010/RetailStoreAccess.git
```
```bash
cd D:\RetailStore\RetailStoreAccess
```

**كل مرة بعدها** (لأخذ آخر ما رُفع على الفرع):
```bash
cd D:\RetailStore\RetailStoreAccess
```
```bash
git pull
```

**للتأكد من النسخة:** `git log --oneline -3` يعرض آخر ثلاث عمليات، والأولى هي آخر ما ذكره Claude في تقريره.

**اختياري: المولّد والاختبارات** (يتأكد أن الكود المولّد يطابق المصدر):
```bash
python tools\generate.py
```
```bash
python -m unittest discover -s tests
```

## 3) Access
1. انقر نقرًا مزدوجًا على `dist\tools\BuildFrontEnd.vbs` واختر مكان ملف الواجهة (الافتراضي `dist\RetailStore_FE.accdb`).
2. يستورد كل الوحدات ويبني الجداول والعلاقات والاستعلامات والشاشات والتقارير؛ بعد كل خطوة رسالة: اضغط «موافق».
3. ينشئ ملف البيانات `RetailStore_BE.accdb` بجانب الواجهة، أو يحدّثه إن وُجد دون مساس ببياناته.
4. في النهاية يسأل: البيانات التجريبية (نعم للتجربة)، وتشغيل كل الاختبارات، ووضع المستخدم النهائي.
5. افتح `RetailStore_FE.accdb` وادخل بـ `admin`.
6. إن فشلت خطوة: أرسل `BuildFrontEnd.log` (بجانب الملف) ونص الرسالة.
- **عند كل تحديث:** `git pull` ثم شغّل `BuildFrontEnd.vbs` مرة أخرى ووافق على استبدال الواجهة.
- **نسخة التسليم (ACCDE بالسر الحقيقي):** من حزمة `release\RetailStore`، لا من `dist` (`docs/88` القسم 3).

## 4) الموقع (الويب والـ API) على جهازك
```bash
cd D:\RetailStore\RetailStoreAccess\web\src\RetailStore.Web
```
```bash
dotnet run
```
1. أول مرة يأخذ دقيقة لتنزيل الحزم، ثم يكتب `Now listening on: http://localhost:5137`.
2. افتح المتصفح على http://localhost:5137.
3. تُنشأ قاعدة تجربة SQLite في `App_Data\retailstore-dev.db` بالبيانات الأساسية فقط.
4. ادخل بـ `admin` بلا كلمة مرور، ويطلب كلمة مرور جديدة.
5. أضف أصنافًا وعملاء من شاشاتهم لتجرّب البيع.
6. للإيقاف: Ctrl+C في نافذة PowerShell.
- **ليصل إليه الجوال أو جهاز آخر في الشبكة:** شغّله هكذا، واسمح بالمنفذ 5137 في جدار حماية Windows عند السؤال:
  ```bash
  dotnet run --urls http://0.0.0.0:5137
  ```
  وعنوانه للأجهزة الأخرى `http://<عنوان جهازك>:5137` (عنوان الجهاز من الأمر `ipconfig`، مثل `192.168.1.10`).
- **على قاعدة SQL Server بدل التجربة:** في `appsettings.Development.json` (أو `appsettings.json`) اجعل `Database:Provider` = `SqlServer` واكتب سلسلة الاتصال في `ConnectionStrings:SqlServer`؛ القاعدة تُنشأ بسكربت `dist\sql\RetailStore_Schema.sql` ثم `RetailStore_Queries.sql` (`docs/67`، `docs/71`)، لا من الموقع.
- **الموقع على الاستضافة:** بأداة `tools\deploy_web.py` (`docs/88`).

## 5) نقطة البيع على Windows
**البناء:**
```bash
cd D:\RetailStore\RetailStoreAccess\web
```
```bash
dotnet publish src\RetailStore.Pos -f net10.0-windows10.0.19041.0 -c Release
```
ينتهي بسطر `RetailStore.Pos ... succeeded`، والبرنامج في:
`web\src\RetailStore.Pos\bin\Release\net10.0-windows10.0.19041.0\win-x64\publish\RetailStore.Pos.exe`

**التشغيل:**
```bash
& "D:\RetailStore\RetailStoreAccess\web\src\RetailStore.Pos\bin\Release\net10.0-windows10.0.19041.0\win-x64\publish\RetailStore.Pos.exe"
```
أول تشغيل يفتح «إعدادات قاعدة البيانات». اختر واحدًا:

| الاختيار | ماذا تكتب | متى |
|---|---|---|
| محلية ← ملف تجربة SQLite | مسار ملف (يُنشأ بالبيانات الأساسية إن لم يوجد). **لترى أصناف الموقع المحلي نفسها** اكتب مسار قاعدته: `D:\RetailStore\RetailStoreAccess\web\src\RetailStore.Web\App_Data\retailstore-dev.db` | للتجربة |
| محلية ← SQL Server | الخادم (مثل `.\SQLEXPRESS`) والقاعدة وطريقة الدخول | في المحل، القاعدة نفسها التي يربطها Access |
| خادم بعيد | عنوان الموقع `https://...` | الكاشير يصل للبرنامج بالإنترنت |

ثم «رمز الجهاز» (مثل `POS1`)، و«اختبار الاتصال»، و«حفظ»، والدخول باسم مستخدم من البرنامج.
- إعدادات الجهاز وفواتيره في `pos.db` لمستخدم Windows؛ تغيير القاعدة بعد ذلك من شاشة «إعدادات قاعدة البيانات» للمدير فقط، ويُرفض ما دامت على الجهاز فواتير لم تُرسل.
- البيع يعمل بلا اتصال؛ المرتجع وسند القبض والتصفية تحتاج الاتصال (`docs/83`، `docs/86`).
- **التحديث:** `git pull` ثم أمر البناء مرة أخرى (أغلق البرنامج قبلها).
- **أو من Visual Studio 2026:** افتح `web\RetailStore.Windows.slnx`، واختر `RetailStore.Pos`، ثم F5.

## 6) تطبيق المندوب على Android
**قبل البناء:** الموقع يعمل ويصل إليه الجوال (القسم 4 بـ `--urls http://0.0.0.0:5137` داخل الشبكة، أو موقع الاستضافة بـ HTTPS)، وللمندوب مستخدم مربوط به في شاشة المندوبين بصلاحيتي «طلبيات العملاء» و«خط السير وزيارات المندوبين» ومخزن سيارة (`docs/78`).

**البناء** (أول أمر مرة واحدة: ينزّل أدوات Android):
```bash
cd D:\RetailStore\RetailStoreAccess\web
```
```bash
dotnet build src\RetailStore.App -t:InstallAndroidDependencies -f net10.0-android -p:AcceptAndroidSDKLicenses=True
```
```bash
dotnet publish src\RetailStore.App -f net10.0-android -c Release
```
الملف الناتج: `web\src\RetailStore.App\bin\Release\net10.0-android\publish\com.retailstore.rep-Signed.apk`

**التثبيت والتشغيل:**
1. انقل ملف `apk` إلى الجوال (كابل، أو واتساب لنفسك، أو Google Drive) وافتحه، واسمح بالتثبيت من مصادر غير معروفة.
2. افتح التطبيق: اكتب عنوان الموقع (مثل `http://192.168.1.10:5137` داخل المحل، أو `https://...`)، ثم اسم المندوب وكلمة مروره.
- **أو من Visual Studio 2026:** افتح `web\RetailStore.Mobile.slnx`، ووصّل الجوال بوضع المطوّر (أو محاكي Android)، ثم F5.

## 7) خدمة مهام المكتب (القيود والفاتورة الإلكترونية بلا Access مفتوح)
تحتاج قاعدة SQL Server (`docs/85`).
```bash
cd D:\RetailStore\RetailStoreAccess
```
```bash
dotnet publish web\src\RetailStore.OfficeService -c Release -r win-x64 --self-contained -o C:\RetailStore\OfficeService
```
1. في `C:\RetailStore\OfficeService\appsettings.json` اكتب سلسلة الاتصال في `ConnectionStrings:SqlServer`.
2. للتجربة دون تثبيتها خدمة:
   ```bash
   C:\RetailStore\OfficeService\RetailStore.OfficeService.exe --check
   ```
   يطبع اسم الخدمة وحالة الفاتورة الإلكترونية وما ينتظر؛ و`--run-now` يشغّل المهام مرة واحدة.
3. التثبيت خدمة تبدأ مع Windows (PowerShell كمسؤول) وكتابة اسمها في شاشة «مهام المكتب» في Access: الخطوات في `docs/85` القسم 2.

## 8) المعاينة في المتصفح (للتطوير فقط)
تعرض صفحات نقطة البيع وتطبيق المندوب في المتصفح بلا بناء MAUI، لرؤية الواجهة بسرعة:
```bash
cd D:\RetailStore\RetailStoreAccess\web\src\RetailStore.AppPreview
```
```bash
dotnet run
```
نقطة البيع على http://localhost:5150. ولتطبيق المندوب: `dotnet run --launch-profile rep` ثم http://localhost:5151. لا تُسلَّم للعميل.

## 9) إذا ظهر خطأ
- **بناء .NET:** انسخ آخر 20 سطرًا من نافذة PowerShell (السطر الذي فيه `error`) وأرسلها.
- **Access:** نص الرسالة، و`BuildFrontEnd.log` أو `RunAllTests.log`، أو السطر الذي يقف عنده Debug ← Compile.
- **نقطة البيع أو التطبيق:** صورة الشاشة ونص الرسالة.
