# نقل المشروع إلى جهازك والعمل عليه بـ Claude Code CLI

المشروع كله، بقواعده وحالته، محفوظ داخل المستودع. لذلك تستطيع أن تكمل العمل من الطرفية على جهازك بالأمر `claude`، من حيث توقّفت الجلسة السحابية تمامًا.

## 1) التجهيز (مرة واحدة)
| # | الخطوة |
|---|---|
| 1 | ثبّت **Git** (git-scm.com). على Windows يلزم **Git for Windows**، لأن Claude Code يستعمل Git Bash |
| 2 | ثبّت **Python 3.11 أو أحدث** (python.org). على Windows فعّل «Add python.exe to PATH» |
| 3 | ثبّت **Claude Code CLI**:<br>• Windows (PowerShell): `irm https://claude.ai/install.ps1 \| iex`<br>• Mac وLinux: `curl -fsSL https://claude.ai/install.sh \| bash`<br>• أو عبر Node.js 18+: `npm install -g @anthropic-ai/claude-code` |
| 4 | انسخ المستودع وانتقل إلى فرع العمل:<br>`git clone https://github.com/mohsaleh2010/RetailStoreAccess.git`<br>`cd RetailStoreAccess`<br>`git checkout claude/retail-store-management-access-enwx66` |
| 5 | **اختياري:** `pip install -r requirements-dev.txt` لاختبارات إضافية. بدونها تُتخطى بعض الاختبارات ولا تفشل |
| 6 | داخل المجلد اكتب `claude`، وسجّل الدخول بحسابك في أول مرة |

**إذا كان لديك نسخة قديمة من المستودع:** `git fetch origin` ثم `git checkout claude/retail-store-management-access-enwx66` ثم `git pull`.

## 2) ما يقرؤه Claude تلقائيًا
| الملف | الغرض |
|---|---|
| `CLAUDE.md` ← `AGENTS.md` | القواعد: الرد بالعربية، وميزة واحدة في كل مرة، والتقرير والانتظار، وقواعد Access |
| [`docs/dev/05-Progress.md`](dev/05-Progress.md) | **حالة العمل:** آخر عملية، وما ينتظر الموافقة، والخطوة التالية. يُحدَّث بعد كل عملية |
| `docs/dev/01..04` | البنية، وقاعدة البيانات، والشاشات، وطريقة العمل |
| `.claude/skills/*` | الأوامر `/continue` و`/run-tests` و`/new-feature` و`/access-error` |
| `.claude/settings.json` | يسمح بالتوليد والاختبارات وأوامر git دون سؤالك كل مرة |

## 3) أول تشغيل على جهازك
1. اكتب في `claude`:
   > /run-tests

   يجب أن تنتهي بـ `OK` (نحو 1220 اختبارًا، 7 دقائق تقريبًا). على Windows تُتخطى اختبارات LibreOffice غالبًا، وهذا طبيعي.
2. ثم:
   > /continue

   يقرأ Claude ملف الحالة ويخبرك أين توقّف العمل، ثم يبدأ الخطوة التالية.

## 4) العمل اليومي
| تريد | اكتب |
|---|---|
| متابعة الخطة من حيث توقّفت | `/continue` |
| الموافقة على ما سُلّم والانتقال للتالي | `موافق` |
| ميزة جديدة من خارج الخطة | `/new-feature` ثم وصفها |
| خطأ ظهر في Access | `/access-error` ثم الصق الرسالة أو `BuildFrontEnd.log` أو `RunAllTests.log` |
| تشغيل الاختبارات | `/run-tests` |

**بعد كل عملية** يحدّث Claude ملفات Markdown: وثيقة الميزة، و`README.md`، وسجل التنفيذ في `docs/51`، و`AGENTS.md` و`docs/dev/*`، وملف الحالة `docs/dev/05-Progress.md`. ثم يحفظ في git ويرفع إلى الفرع.

**الجلسة السحابية والجهاز معًا:** كلاهما يعمل على الفرع نفسه. قبل البدء في أي منهما اكتب `git pull`، حتى لا يعمل أحدهما على نسخة قديمة.

## 5) بناء البرنامج في Access
كما في [00-Claude-Desktop.md](00-Claude-Desktop.md) القسم 4: شغّل `dist\tools\BuildFrontEnd.vbs`، وأرسل `BuildFrontEnd.log` إلى Claude إذا فشلت خطوة.

## 6) قواعد مهمة
- **`LICENSE_SECRET`** في `modActivation` لا يُنشر ولا يُطبع. غيّره قبل التسليم، وسلّم العميل ملف **ACCDE**.
- **لا تعدّل `dist/vba` ولا الوحدات المولَّدة يدويًا.** التعديل في `tools/*.py` ثم `python tools/generate.py`.
