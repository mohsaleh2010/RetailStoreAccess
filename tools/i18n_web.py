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
    # phase 6, batch 5 of the redesign: reports (/trial-balance, /financials, /reports and the groups of forms.REPORT_GROUPS)
    "عدد الحسابات": "Accounts", "كل التقارير": "All reports", "المندوبون": "Sales reps",
    "الخزينة والمصروفات": "Treasury and expenses", "الضريبة والفاتورة الإلكترونية": "VAT and e-invoicing",
    "الرقابة": "Control",
    # phase 6, batch 4b of the redesign: banks and entries (/bank-transactions, /bank-reconciliation, /journal, /manual-entries)
    'ملخص الحركة': 'Transaction summary',
    'ملخص التسوية': 'Reconciliation summary',
    'سطور القيد': 'Entry lines',
    'أسطر القيد': 'Entry lines',
    'أضف سطرًا لكل حساب': 'Add a line for each account',
    'ملخص القيد': 'Entry summary',
    'عدد الأسطر': 'Lines',
    # phase 6, batch 4a of the redesign: the treasury pages (/treasury, /cash-vouchers, /cash-closing)
    'اختر صندوقًا من القائمة لعرض حركته': 'Choose a cash box in the list to see its moves',
    'موقوف': 'Inactive',
    'لا توجد صناديق يمكنك تصفيتها': 'There is no cash box you can close',
    'لا توجد حركة منذ آخر تصفية': 'No moves since the last closing',
    'ملخص التصفية': 'Closing summary',
    'حركة الصندوق من آخر تصفية': 'Box moves since the last closing',
    'لا توجد تحويلات بانتظار الاستلام': 'No transfers waiting for their receipt',
    # phase 6, batch 2 of the redesign: the purchase pages (/purchase, /purchases, /purchases/{id}, /supplier-payments,
    # /purchase-returns)
    'رقم السند أو اسم المورد': 'Voucher number or supplier name',
    'رقم الفاتورة أو اسم المورد': 'Invoice number or supplier name',
    'رصيد المورد': 'Supplier balance',
    # phase 6, batch 1 of the redesign: the sales pages (/sales, /sales/{id}, /payments, /sales-returns)
    'إجمالي الفاتورة': 'Invoice total',
    'المبلغ بالعملة': 'Amount in currency',
    'تاريخ الاستحقاق': 'Due date',
    'تظهر أحدث 200 سجل فقط: ضيّق الفترة أو ابحث': 'Only the latest 200 records are shown: narrow the period or search',
    'رصيد العميل': 'Customer balance',
    'رقم السند أو اسم العميل': 'Voucher number or customer name',
    'رقم الفاتورة أو اسم العميل': 'Invoice number or customer name',
    'عدد الفواتير': 'Invoices',
    'لا توجد سندات في هذه الفترة': 'No vouchers in this period',
    'لا توجد فواتير في هذه الفترة': 'No invoices in this period',
    'لا توجد مرتجعات في هذه الفترة': 'No returns in this period',
    'للبحث': 'to search',
    'مرتجع جديد': 'New return',
    'ملخص السند': 'Voucher summary',
    # the shared components of the redesign (docs/ERP-Redesign-Prompt.md, phase 3) and the sample data of /dev/components
    'أضف أول فاتورة لتظهر هنا': 'Add the first invoice to see it here',
    'أضف صنفًا بزر جديد': 'Add an item with the New button',
    'أي من': 'is any of',
    'إخفاء': 'Hide',
    'إخفاء الكل': 'Hide all',
    'إضافة فلتر': 'Add filter',
    'إظهار الكل': 'Show all',
    'إلغاء الترتيب': 'Unsort',
    'إنقاص': 'Decrement',
    'الأعمدة': 'Columns',
    'التطوير': 'Development',
    'الخانة التالية': 'next cell',
    'الشرط': 'Operator',
    'الصفحة الأخيرة': 'Last page',
    'الصفحة الأولى': 'First page',
    'الصفحة التالية': 'Next page',
    'الصفحة السابقة': 'Previous page',
    'الصنف بالكود أو الاسم': 'Item by code or name',
    'العمود': 'Column',
    'المكونات المشتركة لكل الشاشات ببيانات تجريبية': 'The shared components of every screen with sample data',
    'بالكود أو الاسم': 'By code or name',
    'بعد الترحيل لا يعدل المستند إلا بإلغاء ترحيله': 'Once posted, the document changes only after its posting is undone',
    'تأكيد': 'Confirm',
    'تحديث البيانات': 'Refresh data',
    'ترتيب': 'Sort',
    'تطبيق': 'Apply',
    'تم الترحيل': 'Posted',
    'تم التسليم': 'Delivered',
    'جار التحميل': 'Loading',
    'جرّب كلمة بحث أخرى': 'Try another search word',
    'جزئي': 'Partial',
    'حذف السطر': 'Delete the line',
    'حذف السطر من خانة فارغة': 'deletes the line from an empty cell',
    'حذف الصنف': 'Delete the item',
    'حذف الفلتر': 'Remove filter',
    'خيارات التقرير': 'Report options',
    'خيارات العمود': 'Column options',
    'سيُحذف الصنف نهائيًا ولا يمكن التراجع': 'The item will be deleted for good, with no undo',
    'صفوف في الصفحة': 'Rows per page',
    'صنف جديد': 'New item',
    'عن الفترة السابقة': 'from the period before',
    'غير فارغ': 'is not empty',
    'غير مدفوع': 'Unpaid',
    'غير مطابق': 'is not',
    'فتح الفلاتر': 'Open filters',
    'فلتر': 'Filter',
    'في أو بعد': 'is on or after',
    'في أو قبل': 'is on or before',
    'في الطريق': 'In transit',
    'قيمة الفلتر': 'Filter value',
    'لا توجد بيانات': 'No data',
    'لا توجد فواتير': 'No invoices',
    'لا توجد نتائج': 'No results',
    'لا يحتوي': 'does not contain',
    'لا يساوي': 'not equals',
    'مدفوع جزئيًا': 'Partly paid',
    'مسار الصفحة': 'Breadcrumb',
    'مسح الفلتر': 'Clear filter',
    'معرض المكونات': 'Components gallery',
    'يبدأ بـ': 'starts with',
    'يحتوي': 'contains',
    'ينتهي بـ': 'ends with',
    'أرز بسمتي 5 كجم': 'Basmati rice 5 kg',
    'حليب طويل الأجل 1 لتر': 'Long life milk 1 l',
    'زيت دوار الشمس 1.5 لتر': 'Sunflower oil 1.5 l',
    'سكر أبيض 1 كجم': 'White sugar 1 kg',
    'شاي أسود 100 كيس': 'Black tea 100 bags',
    'حبة': 'Piece',
    'علبة': 'Box',
    'كرتون': 'Carton',
    'كيس': 'Bag',
    # the frame of the MudBlazor redesign (docs/ERP-Redesign-Prompt.md, phase 2)
    'الخزينة والبنوك': 'Treasury and banks',
    'الوضع الفاتح': 'Light mode',
    'الوضع الليلي': 'Dark mode',
    'بحث عن شاشة': 'Find a screen',
    # the rest of the web and the invoice and receipt of the rep app (docs/87)
    'أضف أصناف الفاتورة': 'Add the items of the invoice',
    'إرجاع أصناف فاتورة شراء للمورد: نقدًا أو خصمًا من المستحق له': 'Return the items of a purchase invoice to the supplier: in cash or off what is owed to him',
    'إشعار دائن من رقم الفاتورة: الكمية المرتجعة، والرد نقدًا أو الخصم من الرصيد': 'A credit note from the invoice number: the quantity returned, paid back in cash or taken off the balance',
    'اكتب رقم الفاتورة ثم «بحث': 'Type the invoice number, then Search',
    'اكتب رقم فاتورة الشراء ثم «بحث': 'Type the purchase invoice number, then Search',
    'التسويات': 'Reconciliations',
    'الرصيد الدفتري للصندوق من آخر تصفية، والنقدية الفعلية، والعجز أو الزيادة، والترحيل للخزينة': 'The book balance of the box since its last closing, the counted cash, the shortage or overage, and the transfer to the safe',
    'الرصيد المعدل': 'Adjusted balance',
    'الفاتورة والتحصيل بعد فتح يوم المندوب في البرنامج': 'Invoices and receipts once the rep day is open in the program',
    'الفاتورة والتحصيل تحتاج الاتصال بالبرنامج، والطلبيات والزيارات تعمل بدونه': 'Invoices and receipts need the connection to the program, orders and visits work without it',
    'المشترى': 'Bought',
    'بالمخزون': 'In stock',
    'بدء التسوية': 'Start the reconciliation',
    'تعذر الوصول إلى البرنامج. أعد الإرسال عند عودة الاتصال: لن يتكرر المستند': 'The program could not be reached. Send again when the connection is back: the document will not be made twice',
    'حذف التسوية': 'Delete the reconciliation',
    'حركة بنكية': 'Bank transaction',
    'حفظ سند القبض': 'Save the receipt',
    'خصم من المستحق للمورد': 'Off what is owed to the supplier',
    'رقم الفاتورة أو رقم فاتورة المورد': 'Invoice number or supplier invoice number',
    'سطر جديد': 'New line',
    'عمليات ظهرت في هذا الكشف': 'Operations in this statement',
    'عمليات لم تظهر في كشف بعد': 'Operations not in a statement yet',
    'قيد يومية يدوي بأسطره المدينة والدائنة، وتعديله وحذفه وعكسه': 'A manual journal entry with its debit and credit lines, changed, deleted or reversed',
    'لا يوجد يوم مفتوح لك: افتح يوم المندوب في البرنامج (شاشة يوم المندوب)، ثم أعد المحاولة': 'You have no open day: open the rep day in the program (the rep day screen), then try again',
    'مرتجع للمورد': 'Return to the supplier',
    'مرتجعات المبيعات': 'Sales returns',
    'مرتجعات المشتريات': 'Purchase returns',
    'مطابقة': 'Tick',
    'مطابقة كشف البنك مع الدفاتر واعتماد التسوية': 'Match the bank statement with the books and approve the reconciliation',
    # returns, receipts and cashier closings on the Windows POS (docs/86)
    'إلى المخزون': 'To stock',
    'اكتب رقم الفاتورة': 'Type the invoice number',
    'الرصيد الدفتري': 'Book balance',
    'المبلغ المردود نقدًا': 'Cash paid back',
    'المرتجع سابقًا': 'Returned before',
    'تعذر الاتصال بقاعدة البيانات. المرتجع وسند القبض والتصفية تحتاج الاتصال، والبيع يعمل بدونه':
        'The database cannot be reached. Returns, receipts and closings need the connection, selling works without it',
    'توقيع الكاشير': 'Cashier signature',
    'توقيع المستلم': 'Receiver signature',
    'رقم المستند في الجهاز أطول من 40 حرفًا': 'The number of the document on the device is longer than 40 characters',
    'رقم المستند في الجهاز مستعمل لمستخدم آخر': 'The number of the document on the device is used by another user',
    'رقم المستند في الجهاز مطلوب': 'The number of the document on the device is required',
    'طباعة مرة أخرى': 'Print again',
    'على الجهاز فواتير لك لم تصل إلى البرنامج: أرسلها أولًا (أو احذف المرفوضة)، ثم صفِّ الصندوق':
        'This computer has invoices of yours that have not reached the program: send them first (or delete the refused ones), then close the box',
    'عن الفاتورة': 'For invoice',
    'هذه الفاتورة على الجهاز ولم تصل إلى البرنامج بعد. أرسلها أولًا، ثم أرجعها':
        'This invoice is on this computer and has not reached the program yet. Send it first, then return it',
    # the office tasks of Access on the dashboard (docs/84)
    'افتح شاشة مهام المكتب في': 'Open the office tasks screen in',
    'توقفت مهام المكتب في Access: آخر تشغيل': 'The office tasks of Access stopped: last run',
    'على الجهاز': 'on the computer',
    'لا يعمل جهاز مهام المكتب تلقائيًا، وتنتظر فواتير إلكترونية الإرسال': 'No office computer runs the tasks automatically, and e-invoices wait to be sent',
    # plan 6G (docs/83): the point of sale on Windows and its database settings
    'آخر تحديث للأسعار والأرصدة': 'Prices and balances last read',
    'أول تشغيل على هذا الجهاز: اختر مكان بيانات البرنامج. بعد الحفظ لا يغيّرها إلا مدير النظام': 'First run on this PC: choose where the data of the program is. Once saved, only the administrator changes it',
    'إرسال وتحديث': 'Send and refresh',
    'إعادة الإرسال بالموافقة على المخزون': 'Send again, approving the stock',
    'إعدادات قاعدة البيانات': 'Database settings',
    'إعدادات قاعدة البيانات لمدير النظام فقط': 'The database settings are for the administrator only',
    'اجعل لكل جهاز رمزًا مختلفًا في إعدادات قاعدة البيانات': 'Give each PC its own code in the database settings',
    'اختر ملف قاعدة التجربة': 'Choose the trial database file',
    'اختر نوع قاعدة البيانات: محلية أو بعيدة': 'Choose the kind of database: local or remote',
    'اسم الخادم': 'Server name',
    'اضبط إعدادات قاعدة البيانات أولًا': 'Set the database settings first',
    'اكتب اسم الخادم، مثل': 'Type the server name, such as',
    'اكتب اسم قاعدة البيانات': 'Type the database name',
    'اكتب اسم مستخدم SQL Server، أو اختر الدخول بحساب': 'Type the SQL Server user, or choose to log in with the account of',
    'اكتب كلمة مرور مستخدم': 'Type the password of the user of',
    'الإجمالي قبل الضريبة': 'Total before VAT',
    'الاتصال ناجح': 'The connection works',
    'الباركود أو الرمز أو جزء من الاسم': 'Barcode, code or part of the name',
    'الخادم البعيد يحتاج HTTPS. يُقبل http داخل شبكة المحل فقط': 'A remote server needs HTTPS. http is accepted inside the store network only',
    'الدخول بحساب': 'Log in with the account of',
    'العنوان لا يرد بخادم البرنامج': 'The address does not answer as the server of the program',
    'الوثوق بشهادة خادم المحل': 'Trust the certificate of the store server',
    'انتهى الدخول على الخادم: اخرج وادخل من جديد لتحديث الأسعار والأرصدة': 'The login on the server ended: log out and in again to refresh the prices and balances',
    'بعد الخروج تبقى فواتير الجهاز وتُرسل، ويدخل كاشير آخر': 'After logging out the invoices of the PC stay and are sent, and another cashier logs in',
    'بلا اتصال': 'Offline',
    'بلا اتصال بقاعدة البيانات يدخل من دخل على هذا الجهاز من قبل، ويبيع، وتُرسل فواتيره عند عودة الاتصال': 'Without the database, a user who logged in on this PC before logs in and sells, and his invoices are sent when the connection returns',
    'تبقى الفواتير المرسلة هنا سبعة أيام للطباعة، وهي في البرنامج بأرقامها نفسها': 'Sent invoices stay here seven days for printing, and the program has them with the same numbers',
    'تعذر الاتصال بقاعدة البيانات': 'Could not reach the database',
    'تعذر الاتصال بقاعدة البيانات. ستُرسل الفواتير عند عودة الاتصال': 'Could not reach the database. The invoices will be sent when the connection returns',
    'تعذر الاتصال بقاعدة البيانات، وآخر دخول لهذا المستخدم على البرنامج منذ أكثر من 30 يومًا': 'Could not reach the database, and this user last logged in on the program more than 30 days ago',
    'تعذر الاتصال بقاعدة البيانات، ولا توجد بيانات محفوظة لهذا المستخدم على الجهاز': 'Could not reach the database, and no data of this user is kept on this PC',
    'تعذر الاتصال بقاعدة البيانات، ولم يدخل هذا المستخدم بكلمة المرور هذه على هذا الجهاز من قبل': 'Could not reach the database, and this user never logged in on this PC with this password',
    'جارٍ الإرسال': 'Sending',
    'خادم البرنامج في مكان آخر، عبر الـ API بعنوانه وبـ HTTPS. إن انقطع الإنترنت يبيع الجهاز ويحفظ الفواتير، وتُرسل عند عودته': 'The server of the program elsewhere, through the API at its HTTPS address. Without the internet the PC keeps selling and keeps the invoices, sent when it returns',
    'خادم بعيد': 'Remote server',
    'رقم الفاتورة في الجهاز أطول من 40 حرفًا': 'The number of the invoice on the device is longer than 40 characters',
    'رقم الفاتورة في الجهاز مستعمل لمستخدم آخر': 'The number of the invoice on the device is used by another user',
    'رقم الفاتورة في الجهاز مطلوب': 'The number of the invoice on the device is required',
    'رقم الفاتورة من الجهاز غير صالح: رمز الجهاز ثم شرطة ثم أرقام': 'The invoice number of the device is not valid: the device code, a dash, then digits',
    'رمز الجهاز': 'Device code',
    'رمز الجهاز من حرف إلى ثمانية حروف إنجليزية أو أرقام، مثل': 'The device code is one to eight English letters or digits, such as',
    'رمز الجهاز هو بادئة فواتير البرنامج. اختر رمزًا آخر': 'The device code is the prefix of the invoices of the program. Choose another code',
    'طابعة إيصالات ESC/POS معرّفة في Windows، والدرج موصول بها': 'An ESC/POS receipt printer installed in Windows, with the drawer plugged into it',
    'طابعة الإيصالات': 'Receipt printer',
    'على الجهاز فواتير لم تصل إلى قاعدة البيانات الحالية. أرسلها أولًا، ثم غيّر قاعدة البيانات': 'The PC has invoices that did not reach the current database. Send them first, then change the database',
    'على هذا الجهاز أو على شبكة المحل، وهو القاعدة نفسها التي يربط بها Access والموقع. يعمل البيع إن انقطع الإنترنت': 'on this PC or the store network, the same database Access and the website link to. Selling works without the internet',
    'عنوان الخادم غير صالح، مثل': 'The server address is not valid, such as',
    'فتح درج النقود مع البيع النقدي': 'Open the cash drawer on a cash sale',
    'فواتير الجهاز': 'Invoices of the PC',
    'كل فواتير الجهاز': 'All invoices of the PC',
    'لا توجد فواتير على هذا الجهاز': 'No invoices on this PC',
    'فواتير بانتظار دخول الكاشير الذي باعها': 'Invoices waiting for the cashier who made them to log in',
    'فواتير رفضها البرنامج': 'Invoices refused by the program',
    'قاعدة البيانات لا تحتوي إعدادات البرنامج': 'The database has no settings of the program',
    'قاعدة بيانات محلية': 'Local database',
    'لا يوجد صنف بهذا الرمز': 'No product has this code',
    'لم تُرسل إلى البرنامج بعد': 'Not sent to the program yet',
    'متصل': 'Online',
    'مستعمل لفاتورة أخرى': 'is used by another invoice',
    'ملف تجربة': 'Trial file',
    'ملف قاعدة التجربة': 'Trial database file',
    'هذا الإجراء لمدير النظام فقط': 'This action is for the administrator only',
    'وصلت إلى البرنامج': 'Reached the program',
    'يبدأ به رقم كل فاتورة من هذا الجهاز، مثل POS1-000001. اجعل لكل جهاز رمزًا مختلفًا': 'Every invoice number of this PC starts with it, such as POS1-000001. Give each PC its own code',
    'يجب تغيير كلمة المرور أولًا من البرنامج أو الموقع، ثم الدخول هنا': 'Change the password first in the program or the website, then log in here',
    # plan 6F (docs/82): the financial indicators and the report centre on the web
    'التقرير': 'Report',
    'المؤشرات المالية': 'Financial indicators',
    'تصدير إلى': 'Export to',
    'تقارير البرنامج كلها من استعلاماته: المبيعات والمخزون والخزينة والحسابات': 'Every report of the program from its queries: sales, stock, treasury and accounts',
    'مجمل الربح ± فروقات المخزون − المصروفات. الرصيد الافتتاحي لا يدخل في الربح': 'Gross profit plus or minus stock differences less expenses. The opening balance is not part of the profit',
    'يُعرض هذا التقرير في صفحة القوائم المالية مع فترة المقارنة': 'This report is shown on the financial statements page with its comparison period',
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
    'ادخل باسمه وزامن أولًا': 'Log in as that user and sync first',
    'اربط الطابعة أولًا من إعدادات البلوتوث في الجوال': 'First pair the printer in the Bluetooth settings of the phone',
    'استئناف': 'Resume',
    'اكتب عنوان الخادم': 'Type the address of the server',
    'الرئيسية': 'Home',
    'السجل غير موجود': 'The record does not exist',
    'السعر قبل الضريبة': 'Price before VAT',
    'الصفحة غير موجودة': 'Page not found',
    'الطلبيات': 'Orders',
    'العميل غير موجود على الجهاز. زامن أولًا': 'The customer is not on this phone. Sync first',
    'الفواتير بتاريخها وطباعتها مع رمز الفاتورة الإلكترونية': 'The invoices by date and their print with the e-invoice code',
    'القيمة في حقل': 'The value in the field',
    'امسح الباركود أو ابحث عن صنف لإضافته': 'Scan the barcode or search a product to add it',
    'انتهت صلاحية الدخول. سجّل الدخول من جديد': 'The login has expired. Log in again',
    'بعد الخروج تبقى البيانات والعمليات غير المرسلة على الجهاز حتى يدخل المستخدم نفسه': 'After logging out the data and the operations not sent stay on the phone until the same user logs in',
    'تأكيد كلمة المرور': 'Confirm the password',
    'تاريخ التسليم المطلوب': 'Required delivery date',
    'تحصيل من العملاء نقدًا أو بالبطاقة أو التحويل': 'Collect from customers in cash, by card or by transfer',
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
    'حفظ الفاتورة': 'Save the invoice',
    'حقل غير معروف': 'Unknown field',
    'حُفظت الطلبية على الجهاز': 'The order is saved on the phone',
    'خارج الحدود المسموحة': 'is outside the allowed limits',
    'خصم الفاتورة': 'Invoice discount',
    'خط سير اليوم': 'Route of today',
    'خطأ': 'Error',
    'رجوع إلى العميل': 'Back to the customer',
    'رفضها البرنامج': 'Refused by the program',
    'رقم الطلبية في الجهاز أطول من 40 حرفًا': 'The number of the order on the phone is longer than 40 characters',
    'رقم الطلبية في الجهاز مستعمل لمستخدم آخر': 'The number of the order on the phone is used by another user',
    'رمز الفاتورة الإلكترونية': 'E-invoice code',
    'زيارات اليوم': 'Visits of today',
    'زيارة اليوم': 'Visit of today',
    'زيارة خارج خط السير': 'Visit outside the route',
    'زيارة رفضها البرنامج': 'Visit refused by the program',
    'سجّل الدخول للمزامنة': 'Log in to sync',
    'سند قبض جديد': 'New receipt voucher',
    'سندات القبض': 'Receipt vouchers',
    'ضيّق البحث لعرض باقي السجلات': 'Narrow the search to see the other records',
    'طابعة بلوتوث': 'Bluetooth printer',
    'طباعة نسخة العميل': 'Print the customer copy',
    'طلبية عميل': 'Customer order',
    'طلبية وليست فاتورة ضريبية': 'An order, not a tax invoice',
    'عرض الفاتورة وطباعتها': 'View and print the invoice',
    'عرض الورق': 'Paper width',
    'على الجهاز عمليات لم تُرسل للمستخدم': 'This phone has operations not sent of the user',
    'عملائي': 'My customers',
    'عمليات رفضها البرنامج': 'Operations refused by the program',
    'عنوان الخادم': 'Server address',
    'غير صحيحة': 'is not valid',
    'غيّر مستخدم آخر هذا السجل الآن. أعد فتحه ثم عدّل': 'Another user has just changed this record. Open it again, then edit it',
    'فاتورة بيع: الأصناف بالباركود أو الاسم، والعميل، والدفع': 'A sales invoice: products by barcode or name, the customer and the payment',
    'فواتير المبيعات': 'Sales invoices',
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
    'لا يوجد صنف بهذا الرمز أو الاسم': 'No product has this code or name',
    'لم تُرسل إلى المكتب بعد': 'Not sent to the office yet',
    'لم تُسجَّل': 'Not recorded',
    'مخزن السيارة': 'Van warehouse',
    'مرحبًا': 'Welcome',
    'مزامنة': 'Sync',
    'من مستندات اليوم': 'From the documents of the day',
    'نسخة الويب: البيانات الأساسية والمبيعات على نفس قاعدة بيانات البرنامج': 'Web edition: the master data on the same database as the program',
    'نوع الدفع': 'Payment type',
    'هذا المستخدم ليس مندوبًا نشطًا. اربطه بمندوب في شاشة المندوبين': 'This user is not an active sales rep. Link the user to a rep in the sales reps screen',
    'يجب تغيير كلمة المرور قبل المتابعة': 'Change the password before you go on',
    'يختفي من القوائم ويبقى في التقارير': 'it leaves the lists and stays in the reports',
    # plan 6D (docs/80): purchases, stock and transfers on the web
    'أسعار الشراء بدون الضريبة': 'Purchase prices without VAT',
    'إرسال': 'Send',
    'إرسال الأصناف من مخزن إلى آخر واستلامها': 'Send items from one warehouse to another and receive them',
    'اختر صنفًا من القائمة لعرض حركته': 'Choose an item from the list to show its moves',
    'الأصناف بسعر الشراء ووحدتها، والمورد، والدفع، والمخزن': 'Items at their purchase price and unit, the supplier, payment and warehouse',
    'الأصناف تحت حد الطلب': 'Items at or under the minimum',
    'الدفع للموردين نقدًا أو بالتحويل': 'Payments to suppliers in cash or by transfer',
    'الكميات بالوحدة الصغرى للصنف': 'Quantities in the smallest unit of the item',
    'المرجع': 'Reference',
    'تم حفظ التحويل': 'Transfer saved',
    'تمت العملية على التحويل': 'Done on transfer',
    'رصيد كل صنف في كل مخزن، وكارت الصنف': 'The stock of each item in each warehouse, and the item card',
    'سعر البيع الجديد': 'New selling price',
    'سند صرف جديد': 'New payment voucher',
    'سندات الصرف': 'Payment vouchers',
    'عرض الفاتورة': 'Show the invoice',
    'فواتير المشتريات': 'Purchase invoices',
    'فواتير الموردين بتاريخها وأسطرها': 'Supplier invoices with their dates and lines',
    'محجوز': 'Reserved',
    'أرصدة البنوك': 'Bank balances',
    'أرصدة البنوك، والإيداع والسحب والتحويل وتسوية مدى': 'Bank balances, deposits, withdrawals, transfers and Mada settlements',
    'أصناف تحت حد الطلب': 'Products at or under their minimum',
    'إجمالي النقدية': 'Total cash',
    'إلى البنك': 'To bank',
    'إلى الصندوق': 'To box',
    'اختر قيدًا من القائمة لعرض سطوره': 'Choose an entry in the list to see its lines',
    'اضغط لعرض الأصناف': 'Click to see the products',
    'الحساب المقابل': 'Counter account',
    'الخارج': 'Out',
    'الداخل': 'In',
    'الرصيد من القيود: يتحدث بعد تحديث القيود': 'The balance comes from the journal: it changes once the journal is updated',
    'الطرف الآخر': 'Other side',
    'القيود المبنية من العمليات، وسطور كل قيد': 'The entries built from the operations, and the lines of each entry',
    'المقارنة': 'Comparison',
    'الميزانية': 'Balance sheet',
    'النقدية في الصناديق': 'Cash in the boxes',
    'تحويلات بانتظار الاستلام': 'Transfers waiting for their receipt',
    'تم تحديث القيود': 'The journal was updated',
    'تم حذف الحركة': 'Deleted the transaction',
    'تم حفظ السند': 'Saved the voucher',
    'حركة الصندوق': 'Movements of the box',
    'حركة جديدة': 'New transaction',
    'خارج': 'Out',
    'داخل': 'In',
    'رصيد أول المدة وحركة الفترة والرصيد الختامي لكل حساب': 'The opening balance, the movements of the period and the closing balance of each account',
    'رصيد أول المدة، ثم كل سطر قيد بالرصيد بعده': 'The opening balance, then each journal line with the balance after it',
    'رصيد كل صندوق وحركته، واستلام التحويلات بين الصناديق': 'The balance and movements of each box, and the receipt of transfers between boxes',
    'سندات النقدية والتحويل': 'Cash and transfer vouchers',
    'صافي ربح الشهر': 'Net profit of the month',
    'قائمة الدخل والميزانية من القيود، مع فترة المقارنة': 'The income statement and the balance sheet from the journal, with a comparison period',
    'قبض وصرف نقدية من الصناديق، والتحويل بينها': 'Cash in and out of the boxes, and transfers between them',
    'من القيود': 'From the journal',
    # web redesign, phase 4: the pilot screens (templates A / B / C / D)
    'آخر 30 يومًا': 'Last 30 days',
    'آخر الفواتير': 'Latest invoices',
    'إجراءات': 'Actions',
    'إضافة جديد': 'Add new',
    'إيقاف السجل': 'Deactivate the record',
    'اضغط إضافة جديد لإضافة أول سجل': 'Press Add new to add the first record',
    'الفاتورة فارغة': 'The invoice is empty',
    'بحث بالكود أو الاسم': 'Search by code or name',
    'سطور القيود': 'Journal lines',
    'شاشاتك': 'Your screens',
    'كل الفواتير': 'All invoices',
    'لا توجد فواتير في آخر 30 يومًا': 'No invoices in the last 30 days',
    'لا توجد مبيعات في آخر 12 شهرًا': 'No sales in the last 12 months',
    'لا توجد مبيعات في آخر 30 يومًا': 'No sales in the last 30 days',
    'ملخص الفاتورة': 'Invoice summary',
    # phase 6, batch 3 (stock)
    'التحويلات': 'Transfers',
    'حركة الصنف': 'Item moves',
    'رصيد المخزون': 'Stock on hand',
    'عدد الأصناف': 'Number of items',
    'ملخص التحويل': 'Transfer summary',
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
    import gen_web_reports
    return code_strings() + screen_strings() + gen_web_reports.report_strings()


def keys() -> set:
    out = set()
    for s in strings():
        out.update(i18n.fragments(s))
    return out


def dictionary() -> Dict[str, str]:
    return {**EN, **WEB_EN}
