# العمل على المشروع من Claude Desktop

يشرح هذا الملف كيف تجهّز جهازك لتكمل تطوير البرنامج مع Claude في تطبيق **Claude Desktop** (تبويب **Code**)، ثم كيف تبني البرنامج في Access.

## الملفات التي يقرؤها Claude
| الملف | الغرض |
|---|---|
| `AGENTS.md` (في أعلى المستودع) | تعليمات المشروع لأي ذكاء اصطناعي: طريقة العمل، والأوامر، وخريطة الملفات، وقواعد Access التي سببت أخطاء سابقًا |
| `CLAUDE.md` | يقرؤه Claude تلقائيًا عند فتح المجلد، ويضم `AGENTS.md` |
| `docs/dev/*.md` | شرح البنية، وقاعدة البيانات، والشاشات، وطريقة العمل والتفكير |
| `.claude/skills/*/SKILL.md` | أوامر جاهزة تكتبها في المحادثة: `/run-tests` و`/new-feature` و`/access-error` |
| `.claude/settings.json` | يسمح لـ Claude بتشغيل التوليد والاختبارات وأوامر git للقراءة دون أن يسألك كل مرة |
| `docs/NN-*.md` | شرح كل ميزة بالعربية، و`README.md` يعرض حالة كل ميزة |

## 1) التجهيز (مرة واحدة)
| # | الخطوة |
|---|---|
| 1 | ثبّت **Claude Desktop** من claude.ai/download وسجّل الدخول |
| 2 | ثبّت **Git** (git-scm.com) و**Python 3.11 أو أحدث** (python.org، وفعّل «Add python.exe to PATH» أثناء التثبيت) |
| 3 | انسخ المستودع: `git clone https://github.com/mohsaleh2010/RetailStoreAccess.git` ثم `cd RetailStoreAccess` ثم `git checkout claude/retail-store-management-access-enwx66` |
| 4 | **اختياري:** مكتبات الاختبارات الإضافية: `pip install -r requirements-dev.txt`. بدونها تتخطى بعض الاختبارات نفسها ولا تفشل |
| 5 | **اختياري:** LibreOffice لتشغيل اختبارات تنفّذ كود VBA فعليًا. على Windows تتخطى هذه الاختبارات نفسها غالبًا، وهذا طبيعي |
| 6 | في Claude Desktop افتح تبويب **Code**، واختر **مجلدًا محليًا**، ثم اختر مجلد `RetailStoreAccess` |

## 2) أول تشغيل
اكتب لـ Claude:
> شغّل /run-tests

سيولّد الملفات ويشغّل كل الاختبارات (نحو 750 اختبارًا، مدة دقيقتين تقريبًا). يجب أن تنتهي بـ `OK`.

يمكنك تشغيلها بنفسك من الطرفية:
```
python tools/generate.py
python -m unittest discover -s tests
```
على Windows اكتب `python`. على Mac وLinux اكتب `python3`.

## 3) طريقة العمل اليومية
- **ميزة جديدة:**
  > /new-feature سجل التدقيق: من أضاف أو عدّل أو حذف أي مستند، مع القيم قبل وبعد

  ينفّذها Claude كاملة، مع الجداول والشاشات والاختبارات والتوثيق، ثم يرسل لك تقريرًا بالعربية فيه:
  - ما الذي أُضيف؛
  - الوحدات التي تستوردها أو تستبدلها؛
  - أوامر Build التي تشغّلها؛
  - طريقة الاختبار.

  بعدها ينتظر موافقتك.
- **خطأ ظهر في Access:**
  > /access-error

  ثم الصق نص الرسالة، أو صورتها، ومحتوى `BuildFrontEnd.log` أو `RunAllTests.log`. يصلح Claude السبب، ويضيف اختبارًا يمنع تكراره.
- **الحفظ:** يحفظ Claude كل تعديل في git ويرفعه إلى الفرع. لا يفتح Pull Request إلا إذا طلبت.

## 4) بناء البرنامج في Access (على Windows فقط)
1. بعد أي تعديل، شغّل `dist\tools\BuildFrontEnd.vbs` بنقرة مزدوجة.
2. يُنشئ `RetailStore_FE.accdb` كاملًا:
   - يستورد الوحدات؛
   - يترجم الكود؛
   - يشغّل `BuildSchema` ← `BuildRelationships` ← `BuildQueries` ← `BuildForms` ← `BuildReports`.
3. ملف البيانات `RetailStore_BE.accdb` يُنشأ أول مرة، ثم يُحدَّث بعد ذلك ولا يُستبدل أبدًا.
4. يسألك بعد ذلك عن البيانات التجريبية، والاختبارات (`RunAllTests`)، ووضع المستخدم.
5. إذا فشلت خطوة، يكتب السبب في `BuildFrontEnd.log`. أرسل محتواه لـ Claude مع `/access-error`.

ملفات Access المبنية (`*.accdb` و`*.accde`) والسجلات مستثناة من git عبر `.gitignore`، فلن تُرفع إلى المستودع.

## 5) قواعد مهمة
- **`LICENSE_SECRET`** في `modActivation` سرّ التفعيل:
  - لا تنشره، ولا تطلب من Claude طباعته.
  - غيّره قبل التسليم.
  - سلّم العميل ملف **ACCDE**، واحتفظ لنفسك بنسخة ACCDB.
- **لا تعدّل ملفات `dist/vba` يدويًا:** يكتبها المولّد بترميز Windows-1256 وأسطر CRLF.
- **الوحدات المولَّدة** (مثل `modBuildForms` و`modBuildQueries`) لا تُعدَّل يدويًا. التعديل في `tools/*.py` ثم `python tools/generate.py`.
- **كلمة مرور المبرمج:** إن نسيتها، شغّل `ResetDeveloperPassword` في نسخة ACCDB. التفاصيل في `docs/19-Permissions-Activation.md`.

## 6) استخدام ذكاء اصطناعي آخر غير Claude
كل التعليمات موجودة في ملفات عادية داخل المستودع، فيمكن لأي أداة أن تكمل العمل:

| الأداة | ما تفعله |
|---|---|
| **ChatGPT / Codex، وCursor، وGitHub Copilot، وأدوات أخرى كثيرة** | تقرأ `AGENTS.md` تلقائيًا من أعلى المستودع |
| **Gemini CLI** | اطلب منه في أول رسالة قراءة `AGENTS.md`، أو انسخه إلى ملف `GEMINI.md` |
| **محادثة عادية بلا وصول للملفات** | ارفع لها `AGENTS.md` والملفات الأربعة في `docs/dev/` في أول المحادثة، ثم الملفات التي تخص الميزة |

ابدأ أي محادثة جديدة بجملة مثل:
> اقرأ AGENTS.md وملفات docs/dev ثم نفّذ: …

الأوامر `/run-tests` و`/new-feature` و`/access-error` خاصة بـ Claude. خطواتها مكتوبة نصًا في قسم «Standard procedures» في `AGENTS.md`، فيمكنك أن تطلبها من أي أداة بالكلمات.

