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
    'آخر مزامنة': 'Last sync',
    'أرصدة عملائي': 'Balances of my customers',
    'أضف أسطر الطلبية': 'Add the lines of the order',
    'أطول من': 'is longer than',
    'أوقف الخادم الجلسة مؤقتًا': 'The server paused the session',
    'إعادة الإرسال': 'Send again',
    'إعادة التحميل': 'Reload',
    'إعادة المحاولة': 'Retry',
    'ابحث عن صنف أو اكتب الباركود': 'Search a product or type the barcode',
    'اختر الطابعة في الإعدادات': 'Choose the printer in the settings',
    'اختر سجلًا من القائمة أو اضغط جديد': 'Choose a record from the list or press New',
    'ادخل باسمه وزامن أولًا': 'Log in as that user and sync first',
    'اربط الطابعة أولًا من إعدادات البلوتوث في الجوال': 'First pair the printer in the Bluetooth settings of the phone',
    'استئناف': 'Resume',
    'اكتب عنوان الخادم': 'Type the address of the server',
    'الرئيسية': 'Home',
    'السجل غير موجود': 'The record does not exist',
    'الصفحة غير موجودة': 'Page not found',
    'الطلبيات': 'Orders',
    'العميل غير موجود على الجهاز. زامن أولًا': 'The customer is not on this phone. Sync first',
    'القيمة في حقل': 'The value in the field',
    'انتهت صلاحية الدخول. سجّل الدخول من جديد': 'The login has expired. Log in again',
    'بعد الخروج تبقى البيانات والعمليات غير المرسلة على الجهاز حتى يدخل المستخدم نفسه': 'After logging out the data and the operations not sent stay on the phone until the same user logs in',
    'تأكيد كلمة المرور': 'Confirm the password',
    'تاريخ التسليم المطلوب': 'Required delivery date',
    'تطبيق المندوب': 'Rep app',
    'تعذر استئناف الجلسة. أعد المحاولة أو أعد تحميل الصفحة': 'The session could not be resumed. Retry or reload the page',
    'تعذر الاتصال بالخادم. أعد المحاولة أو أعد تحميل الصفحة': 'Could not reach the server. Retry or reload the page',
    'تعذر الاتصال. محاولة جديدة بعد': 'Could not connect. Next attempt in',
    'تعذر الوصول إلى الخادم. تأكد من الإنترنت ومن عنوان الخادم': 'Could not reach the server. Check the internet and the address of the server',
    'تعذر الوصول إلى الخادم. ستُرسل العمليات عند عودة الاتصال': 'Could not reach the server. The operations will be sent when the connection is back',
    'تعذرت الطباعة': 'Printing failed',
    'تُرسل إلى المكتب عند المزامنة، ويعتمدها المكتب': 'It goes to the office at the sync, and the office approves it',
    'ثانية': 'seconds',
    'جارٍ إعادة الاتصال بالخادم': 'Reconnecting to the server',
    'جارٍ المزامنة': 'Syncing',
    'حدث خطأ غير متوقع': 'An unexpected error occurred',
    'حذف من الجهاز': 'Delete from the phone',
    'حفظ الطلبية': 'Save the order',
    'حقل غير معروف': 'Unknown field',
    'حُفظت الطلبية على الجهاز': 'The order is saved on the phone',
    'خارج الحدود المسموحة': 'is outside the allowed limits',
    'خط سير اليوم': 'Route of today',
    'خطأ': 'Error',
    'رجوع إلى العميل': 'Back to the customer',
    'رفضها البرنامج': 'Refused by the program',
    'رقم الطلبية في الجهاز أطول من 40 حرفًا': 'The number of the order on the phone is longer than 40 characters',
    'رقم الطلبية في الجهاز مستعمل لمستخدم آخر': 'The number of the order on the phone is used by another user',
    'زيارات اليوم': 'Visits of today',
    'زيارة اليوم': 'Visit of today',
    'زيارة خارج خط السير': 'Visit outside the route',
    'زيارة رفضها البرنامج': 'Visit refused by the program',
    'سجّل الدخول للمزامنة': 'Log in to sync',
    'ضيّق البحث لعرض باقي السجلات': 'Narrow the search to see the other records',
    'طابعة بلوتوث': 'Bluetooth printer',
    'طباعة نسخة العميل': 'Print the customer copy',
    'طلبية عميل': 'Customer order',
    'طلبية وليست فاتورة ضريبية': 'An order, not a tax invoice',
    'عرض الورق': 'Paper width',
    'على الجهاز عمليات لم تُرسل للمستخدم': 'This phone has operations not sent of the user',
    'عملائي': 'My customers',
    'عمليات رفضها البرنامج': 'Operations refused by the program',
    'عنوان الخادم': 'Server address',
    'غير صحيحة': 'is not valid',
    'غيّر مستخدم آخر هذا السجل الآن. أعد فتحه ثم عدّل': 'Another user has just changed this record. Open it again, then edit it',
    'في السيارة': 'In the van',
    'قيمة غير مقبولة في حقل': 'A value that is not accepted in the field',
    'لا اتصال بالإنترنت. ستُرسل العمليات عند عودته': 'No internet connection. The operations will be sent when it is back',
    'لا تملك صلاحية هذه العملية. راجع مدير النظام': 'You do not have the permission for this operation. Contact the system administrator',
    'لا توجد سجلات': 'No records',
    'لا توجد شاشات متاحة لك على الويب بعد. راجع مدير النظام': 'No web screens are open to you yet. Ask the system administrator',
    'لا زيارات لهذا اليوم بعد': 'No visits for this day yet',
    'لا شيء': 'Nothing',
    'لا مخزن سيارة للمندوب. أضفه في شاشة المخازن (النوع سيارة': 'The rep has no van warehouse. Add it in the warehouses screen (type van',
    'لا يمكن حذف هذا السجل لأنه مستخدم في عمليات سابقة. هل تريد تعطيله بدلًا من ذلك': 'This record cannot be deleted because earlier operations use it. Deactivate it instead',
    'لم تُرسل إلى المكتب بعد': 'Not sent to the office yet',
    'لم تُسجَّل': 'Not recorded',
    'مخزن السيارة': 'Van warehouse',
    'مرحبًا': 'Welcome',
    'مزامنة': 'Sync',
    'من مستندات اليوم': 'From the documents of the day',
    'نسخة الويب: البيانات الأساسية على نفس قاعدة بيانات البرنامج': 'Web edition: the master data on the same database as the program',
    'هذا المستخدم ليس مندوبًا نشطًا. اربطه بمندوب في شاشة المندوبين': 'This user is not an active sales rep. Link the user to a rep in the sales reps screen',
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
