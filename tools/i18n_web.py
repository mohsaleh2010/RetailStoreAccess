"""The English interface of the web system (docs/76): the same dictionary as Access (i18n_en.EN) plus the texts
only the web has (WEB_EN). The web translates with the same rule as modLang.Tr (Domain/Localization/Translator.cs).

Every Arabic text of the web is a key source:
  * the hand-written C# (.cs) of web/src: every Arabic string literal (messages, built with + from pieces);
  * the Razor pages (.razor): every T("...") call; no other Arabic is allowed in them (LANGUAGE_NAMES aside);
  * the generated data screens (gen_web.WEB_SCREENS): captions, column headers, labels, hints, value lists.
The English texts follow the rules of i18n_en.EN (ASCII, none of " ' [ ] ; | & = < >)."""
import os
import re
from typing import Dict, List

import i18n
from i18n_en import EN

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEB_SRC = os.path.join(ROOT, "web", "src")
LANGUAGE_NAMES = {"العربية"}           # the link to the Arabic interface is written in Arabic in both languages

WEB_EN = {
    'أطول من': 'is longer than',
    'أوقف الخادم الجلسة مؤقتًا': 'The server paused the session',
    'إعادة التحميل': 'Reload',
    'إعادة المحاولة': 'Retry',
    'اختر سجلًا من القائمة أو اضغط جديد': 'Choose a record from the list or press New',
    'استئناف': 'Resume',
    'الرئيسية': 'Home',
    'السجل غير موجود': 'The record does not exist',
    'الصفحة غير موجودة': 'Page not found',
    'القيمة في حقل': 'The value in the field',
    'تأكيد كلمة المرور': 'Confirm the password',
    'تعذر استئناف الجلسة. أعد المحاولة أو أعد تحميل الصفحة': 'The session could not be resumed. Retry or reload the page',
    'تعذر الاتصال بالخادم. أعد المحاولة أو أعد تحميل الصفحة': 'Could not reach the server. Retry or reload the page',
    'تعذر الاتصال. محاولة جديدة بعد': 'Could not connect. Next attempt in',
    'ثانية': 'seconds',
    'جارٍ إعادة الاتصال بالخادم': 'Reconnecting to the server',
    'حدث خطأ غير متوقع': 'An unexpected error occurred',
    'خارج الحدود المسموحة': 'is outside the allowed limits',
    'خطأ': 'Error',
    'ضيّق البحث لعرض باقي السجلات': 'Narrow the search to see the other records',
    'غير صحيحة': 'is not valid',
    'غيّر مستخدم آخر هذا السجل الآن. أعد فتحه ثم عدّل': 'Another user has just changed this record. Open it again, then edit it',
    'قيمة غير مقبولة في حقل': 'A value that is not accepted in the field',
    'لا توجد سجلات': 'No records',
    'لا توجد شاشات متاحة لك على الويب بعد. راجع مدير النظام': 'No web screens are open to you yet. Ask the system administrator',
    'لا يمكن حذف هذا السجل لأنه مستخدم في عمليات سابقة. هل تريد تعطيله بدلًا من ذلك': 'This record cannot be deleted because earlier operations use it. Deactivate it instead',
    'مرحبًا': 'Welcome',
    'نسخة الويب: البيانات الأساسية على نفس قاعدة بيانات البرنامج': 'Web edition: the master data on the same database as the program',
    'يجب تغيير كلمة المرور قبل المتابعة': 'Change the password before you go on',
    'يختفي من القوائم ويبقى في التقارير': 'it leaves the lists and stays in the reports',
}

_CS_LITERAL = re.compile(r'"((?:[^"\\\n]|\\.)*)"')
_T_CALL = re.compile(r'\bT\("((?:[^"\\\n]|\\.)*)"\)')


def web_sources() -> List[str]:
    out = []
    for folder, dirs, files in os.walk(WEB_SRC):
        dirs[:] = [d for d in dirs if d not in ("bin", "obj", "Generated", "wwwroot")]
        out += [os.path.join(folder, f) for f in files if f.endswith((".cs", ".razor"))]
    return sorted(out)


def _read(path: str) -> str:
    with open(path, encoding="utf-8-sig") as fh:
        return fh.read()


def razor_outside_t(text: str) -> List[str]:
    """Arabic left in a Razor file once the T("...") calls and the language names are taken out."""
    rest = _T_CALL.sub("", text)
    for name in LANGUAGE_NAMES:
        rest = rest.replace('"' + name + '"', "")
    return [line.strip() for line in rest.splitlines() if i18n.has_arabic(line)]


def code_strings() -> List[str]:
    out = []
    for path in web_sources():
        text = _read(path)
        if path.endswith(".razor"):
            out += _T_CALL.findall(text)
            code = text.split("@code", 1)[1] if "@code" in text else ""
            out += [s for s in _CS_LITERAL.findall(code) if s not in LANGUAGE_NAMES]
        else:
            out += _CS_LITERAL.findall(text)
    return [s for s in out if i18n.has_arabic(s)]


def screen_strings() -> List[str]:
    import forms as F
    import gen_web
    out = []
    for _route, s in gen_web.web_screens():
        out += [s.caption, s.subtitle] + [h for h, _w in s.list_headers]
        t = gen_web.table(s.table)
        for item in s.fields:
            if isinstance(item, F.Info):
                out.append(item.text)
                continue
            f = next(x for x in t.fields if x.name == item.field)
            out += [item.label or f.caption, item.hint]
            out += [text for _v, text in (gen_web._value_list(item.rows) or [])]
    return [x for x in out if i18n.has_arabic(x)]


def strings() -> List[str]:
    return code_strings() + screen_strings()


def keys() -> set:
    out = set()
    for s in strings():
        out.update(i18n.fragments(s))
    return out


def dictionary() -> Dict[str, str]:
    return {**EN, **WEB_EN}
