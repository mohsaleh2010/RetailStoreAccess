# الفاتورة الإلكترونية بـ C# وخدمة Windows لمهام المكتب (المرحلة 7أ)

**الهدف (الخطة 7، `docs/51`):** أن تُبنى القيود وتُوقَّع الفواتير الإلكترونية وتُرسل وتُتابَع **بلا Access مفتوح**، مع بقاء Access كما هو تمامًا. في 6ح (`docs/84`) كان جهاز Access واحد في المكتب يقوم بـ«مهام المكتب» من مؤقت لوحة التحكم، فإذا أُغلق Access توقفت. الآن يقوم بها برنامج صغير يعمل **خدمةً في Windows** (`RetailStore.OfficeService`) يبدأ مع الجهاز ولا يحتاج مستخدمًا.

## 1) ما الذي نُقل إلى C#
نُقلت منطقيات الوحدات الأربع بالحرف، بفحوصها وترتيبها ورسائلها العربية نفسها:

| في Access | في C# | ماذا يفعل |
|---|---|---|
| `modZatcaXml` | `RetailStore.Domain/EInvoicing/ZatcaDocument.cs` | مستند UBL 2.1 مكتوبًا بالصيغة القانونية (C14N)، وبصمته (hash)، وحقول الشهادة، وكتلة التوقيع XAdES، ورمز QR بتسعة حقول، و**التوقيع بمفتاح الجهاز** (secp256k1، SHA-256) بـ .NET نفسه بدل برنامج OpenSSL |
| `modZatcaApi` | `EInvoiceReply.cs` (`ZatcaReplies`) و`EInvoiceService.ZatcaSendAsync` | الإبلاغ (المبسطة) والاعتماد (الضريبية) وقراءة ردود الهيئة، وحفظ المستند المختوم من الهيئة ورمزه |
| `modEtaReceipt` | `EtaDocument.cs` و`EInvoiceService.EtaSendAsync` / `EtaRefreshAsync` | الإيصال الإلكتروني المصري 1.2: JSON، والتسلسل (serialization)، والـ UUID، وسلسلة الإيصالات `Settings.EtaLastUUID`، ورمز الدخول للجهاز، والإرسال والمتابعة |
| `modEtaInvoice` | `EtaDocument.cs` و`EInvoiceService.EtaInvoiceSendAsync` / `EtaInvoiceRefreshAsync` / `EtaCancelAsync` | الفاتورة الإلكترونية المصرية للشركات: 1.0 موقّعة ببرنامج التوقيع (فلاشة التوقيع) أو 0.9 بلا توقيع في التجريبي، والإرسال والمتابعة والإلغاء |
| `modEInvoice` | `EInvoiceService.SendAsync` / `SendPendingAsync` / `RefreshSubmittedAsync` | إرسال مستند، وكل «بانتظار الإرسال» من الأقدم مع التوقف عند أول انقطاع، ومتابعة المرسل في مصر؛ و`EInvoiceLog` لكل طلب |
| `modOfficeSync` | `RetailStore.Data/Services/OfficeSyncService.cs` | تشغيل واحد بالترتيب: القيود (`SyncJournal`) ثم الإرسال ثم المتابعة، وصف في `OfficeSyncLog` و`Settings.OfficeSyncLastAt` / `OfficeSyncLastResult` |

**الحقول نفسها:** يكتب C# ما يكتبه Access بالضبط: `ZatcaStatus`، و`EInvoiceXml`، و`InvoiceHash` و`PreviousInvoiceHash`، و`QRCodeData`، و`EtaUUID` و`EtaPreviousUUID` و`EtaSubmissionId` و`EtaLongId`، و`EInvoiceAttempts` و`EInvoiceError`، و`Settings.LastInvoiceHash` و`EtaLastUUID`، و`EInvoiceLog`. لذلك يرى Access (شاشة «الفاتورة الإلكترونية» وتقريرها وإقرار الضريبة) ما أرسلته الخدمة كأنه أرسله بنفسه، والعكس.

**السلسلة محمية:** التوقيع (السعودية) وبناء الإيصال (مصر) يأخذان قفل السلسلة داخل معاملة واحدة كما يفعل Access (`UPDATE Settings SET LastInvoiceHash = LastInvoiceHash`)، فلو وقّع Access والخدمة في اللحظة نفسها لا تنكسر السلسلة.

**ما بقي في Access وحده:** تسجيل الجهاز لدى الهيئة (طلب الشهادة وفحوص الامتثال والشهادة الفعلية) وإعداد الربط المصري، فهي خطوات تُعمل مرة واحدة من شاشتي الإعداد. والخدمة تستعمل نتيجتها من الإعدادات.

**التطابق مع Access مضمون بالاختبارات:** المستندات والبصمات وحقول الشهادة ورمز QR وقراءة كل رد من ردود المنصتين مقارنة حرفًا بحرف بمراجع Python (`tools/zatca_xml_reference.py`، و`eta_receipt_reference.py`، و`eta_invoice_reference.py`، و`zatca_api_reference.py`) التي تُقارن بها وحدات VBA نفسها.

## 2) خدمة Windows لمهام المكتب
`web/src/RetailStore.OfficeService` (Worker Service، .NET 10):
- كل دقيقة (`OfficeService:CheckSeconds`) تنظر: هل **اسمها** هو المكتوب في `Settings.OfficeSyncComputer`؟ وهل مرت `OfficeSyncMinutes` دقيقة منذ آخر تشغيل (`OfficeSyncRules.IsDue`، القاعدة نفسها في Access)؟ إن نعم شغّلت مهام المكتب (`AUTO`).
- **اسم الخدمة:** `OfficeService:Name` في `appsettings.json`، وإن ترك فارغًا فهو `SERVICE-` واسم الجهاز (مثل `SERVICE-OFFICE-PC`). لا يساوي اسم أي جهاز أبدًا، فلا يعمل Access والخدمة تلقائيًا معًا ولا تُرسل فاتورة مرتين. لنقل المهام من Access إلى الخدمة يكفي كتابة اسم الخدمة في شاشة «مهام المكتب» في Access بدل اسم الجهاز؛ وللرجوع إلى Access اضغط «هذا الجهاز».
- لا تتوقف عند مشكلة: المشكلة تُكتب في `OfficeSyncLog` (يظهر في شاشة «مهام المكتب» في Access) وفي `OfficeSyncLastResult` (ينبّه به الموقع المدير كما في `docs/84`)، وفي سجل أحداث Windows.
- تعمل على القاعدة نفسها: SQL Server (`Database:Provider = SqlServer` وسلسلة الاتصال `ConnectionStrings:SqlServer`)، أو ملف SQLite للتجربة. القيود تُقرأ من عروض الاستعلامات على الخادم (`docs/71`) كما يقرؤها الموقع.

### التثبيت (على جهاز المكتب أو الخادم)
1. انشر البرنامج: `dotnet publish web/src/RetailStore.OfficeService -c Release -r win-x64 --self-contained -o C:\RetailStore\OfficeService`.
2. في `C:\RetailStore\OfficeService\appsettings.json` اكتب سلسلة الاتصال بقاعدة SQL Server (المستخدم يحتاج القراءة والكتابة فقط).
3. جرّب من سطر الأوامر قبل التثبيت:
   - `RetailStore.OfficeService.exe --check`: اسم الخدمة، وهل الفاتورة الإلكترونية جاهزة، وهل **يوقّع مفتاح الجهاز** على هذا الجهاز وتتحقق منه الشهادة، وعدد المنتظر، وهل حان التشغيل.
   - `RetailStore.OfficeService.exe --run-now`: تشغيل واحد يدوي (`MANUAL`) يطبع ما تم.
4. ثبّت الخدمة (موجّه أوامر بصلاحية المدير):
   ```
   sc.exe create "RetailStore Office Service" binPath= "C:\RetailStore\OfficeService\RetailStore.OfficeService.exe" start= delayed-auto
   sc.exe start "RetailStore Office Service"
   ```
5. في Access، شاشة «مهام المكتب»: اكتب في «جهاز مهام المكتب» اسم الخدمة الذي طبعه `--check`، واترك 15 دقيقة، ثم «حفظ الإعداد». بعدها يمكن إغلاق Access.

**ملفات الجهاز:**
- **السعودية:** ملف المفتاح الخاص (`Settings.ZatcaKeyFile`) يجب أن يكون **على الجهاز الذي تعمل عليه الخدمة** بالمسار نفسه، ويقرؤه حساب الخدمة. يحتاج Windows 10 أو Windows Server 2016 فما بعد (منحنى secp256k1)، ولا يحتاج OpenSSL.
- **مصر:** برنامج التوقيع وفلاشة التوقيع (`EtaSignerPath`) على جهاز الخدمة. بعض برامج الفلاشة تحتاج مستخدمًا داخلًا؛ إن لم يعمل التوقيع من الخدمة فشغّلها بحساب المستخدم (`sc.exe config ... obj= .\User password= ...`) أو اترك مهام المكتب على Access في ذلك الجهاز.

## 3) الأمان
- لا تُسجَّل في `EInvoiceLog` الكلمات السرية: طلب رمز الدخول لا يُحفظ، ورمز الدخول في الرد يُستبدل بـ `***`، كما في Access.
- `appsettings.json` في المستودع بلا سلسلة اتصال؛ تُكتب على جهاز التثبيت فقط.
- مفتاح الاختبار في `tools/gen_web_einvoice.py` لشهادة اختبار ذاتية التوقيع (`TST-EGS-RetailStore`)، للاختبارات فقط.

## 4) الاختبارات
- **xUnit** (`web/tests/RetailStore.Tests/EInvoiceTests.cs`، 13 اختبارًا):
  - المستندات والبصمات والقيم وحقول الشهادة وكتلة التوقيع ورمز QR مطابقة لمراجع Python (`EInvoiceVectors.g.cs` من `tools/gen_web_einvoice.py`)؛ وقراءة ردود الهيئة والمصلحة لكل حالة (0، 200، 202، 400، 401، 429، 5xx، غير متوقع).
  - التوقيع بمفتاح الجهاز يتحقق منه بمفتاح الشهادة كما يتحقق من توقيع OpenSSL.
  - على قاعدة اختبار بمنصة وهمية: بيع ← توقيع في سلسلة ← إبلاغ (السعودية)، والاعتماد وحفظ المستند المختوم، وبقاء المستند «بانتظار الإرسال» عند الانقطاع أو نقص البيانات؛ والإيصال المصري في سلسلة ← إرسال ← متابعة؛ والفاتورة المصرية الموقعة ← إرسال ← متابعة ← إلغاء؛ وتشغيل مهام المكتب (القيود والإرسال والسجل والإعدادات واسم الخدمة).
- **Python** (`tests/test_einvoice_csharp.py`): كل رسالة عربية في C# موجودة في وحدات VBA، وترتيب خطوات السلسلة والإرسال ومهام المكتب كترتيب Access، وعدم تسجيل الأسرار، ومشروع الخدمة.
- **يدويًا على جهازك:** `--check` ثم `--run-now` في البيئة التجريبية (Developer Portal للسعودية، أو ما قبل الإنتاج لمصر)، ثم شاشة «الفاتورة الإلكترونية» في Access لترى الحالات والسجل.

## 5) في Access
**لا شيء يتغير:** لا وحدة جديدة ولا حقل جديد ولا `BuildSchema`. فقط اكتب اسم الخدمة في شاشة «مهام المكتب» عند تشغيلها.
