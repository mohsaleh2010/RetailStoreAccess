"""Single source of truth for the RetailStore back-end schema.

tools/generate.py turns this file into:
  * src/vba/modBuildSchema.bas   (readable UTF-8 copy of the VBA builder)
  * dist/vba/modBuildSchema.bas  (Windows-1256 / CRLF copy to import into Access)
  * docs/02-Tables-Reference.md  (field-by-field reference)

Field kinds
-----------
AUTO      AutoNumber (Long Integer, increment)
LONG      Long Integer
INT       Integer
BYTE      Byte
MONEY     Currency, shown as #,##0.00        (amounts in SAR)
QTY       Currency, shown as #,##0.###       (exact quantities, up to 4 decimals)
RATE      Currency, shown as 0.00%           (0.15 = 15%)
DATE      Date/Time, date only
DATETIME  Date/Time, date and time
BOOL      Yes/No
TEXT      Short Text (size = max length)
MEMO      Long Text

Defaults and validation rules are written in Access expression syntax.
"""

from dataclasses import dataclass, field as dc_field
from typing import List, Optional, Tuple

ZATCA_INITIAL_PIH = (
    "NWZlY2ViNjZmZmM4NmYzOGQ5NTI3ODZjNmQ2OTZjNzljMmRiYzIzOWRkNGU5MWI0NjcyOWQ3M2EyN2ZiNTdlOQ=="
)

VAT_RULE = 'Is Null Or Like "3#############3"'
VAT_TEXT = "الرقم الضريبي 15 رقمًا ويبدأ وينتهي بالرقم 3"


@dataclass
class Field:
    name: str
    kind: str
    caption: str
    size: int = 0
    required: bool = False
    default: Optional[str] = None
    rule: Optional[str] = None
    rule_text: Optional[str] = None
    fk: Optional[str] = None          # "Table.Field"
    on_delete_cascade: bool = False   # used by Phase 3 (relationships)
    note: str = ""                    # extra explanation for the reference doc


@dataclass
class Index:
    name: str
    fields: List[str]
    unique: bool = False
    ignore_nulls: bool = False
    primary: bool = False


@dataclass
class Table:
    name: str
    caption: str
    purpose: str
    fields: List[Field]
    pk: List[str]
    indexes: List[Index] = dc_field(default_factory=list)
    rule: Optional[str] = None
    rule_text: Optional[str] = None
    seed_columns: List[str] = dc_field(default_factory=list)
    seed_rows: List[Tuple] = dc_field(default_factory=list)
    # True: BuildSchema also adds seed rows that are missing from an existing back-end
    # (matched by the first seed column), e.g. new sequences and permissions.
    seed_missing: bool = False


# --------------------------------------------------------------------------
# Small helpers so the table definitions below stay short and readable
# --------------------------------------------------------------------------
def auto(name, caption):
    return Field(name, "AUTO", caption)


def text(name, size, caption, required=False, default=None, rule=None, rule_text=None, note=""):
    return Field(name, "TEXT", caption, size=size, required=required, default=default,
                 rule=rule, rule_text=rule_text, note=note)


def memo(name, caption):
    return Field(name, "MEMO", caption)


def long_(name, caption, required=False, default=None, rule=None, rule_text=None, fk=None,
          cascade=False, note=""):
    return Field(name, "LONG", caption, required=required, default=default, rule=rule,
                 rule_text=rule_text, fk=fk, on_delete_cascade=cascade, note=note)


def int_(name, caption, required=False, default=None, rule=None, rule_text=None):
    return Field(name, "INT", caption, required=required, default=default, rule=rule,
                 rule_text=rule_text)


def byte_(name, caption, default=None, rule=None, rule_text=None):
    return Field(name, "BYTE", caption, required=True, default=default, rule=rule,
                 rule_text=rule_text)


def money(name, caption, required=True, default="0", rule=">=0",
          rule_text="المبلغ لا يمكن أن يكون سالبًا", note=""):
    return Field(name, "MONEY", caption, required=required, default=default, rule=rule,
                 rule_text=rule_text if rule else None, note=note)


def qty(name, caption, required=True, default="0", rule=None, rule_text=None, note=""):
    return Field(name, "QTY", caption, required=required, default=default, rule=rule,
                 rule_text=rule_text, note=note)


def rate(name, caption, default=None, rule=">=0 And <1",
         rule_text="النسبة يجب أن تكون بين 0% و 100%"):
    return Field(name, "RATE", caption, required=True, default=default, rule=rule,
                 rule_text=rule_text)


def date_(name, caption, required=True, default="Date()"):
    return Field(name, "DATE", caption, required=required, default=default)


def datetime_(name, caption, required=False, default=None):
    return Field(name, "DATETIME", caption, required=required, default=default)


def fx_fields(foreign_caption="المبلغ بالعملة"):
    """Currency of a document (modCurrency): the amount fields stay in the program currency (SAR);
    these say in which currency the document was entered, at which rate, and its total there."""
    return [Field("CurrencyCode", "TEXT", "العملة", size=3, default='"SAR"', fk="Currencies.CurrencyCode",
                  note="فارغ = عملة البرنامج (Settings.CurrencyCode)"),
            Field("ExchangeRate", "RATE", "معامل التحويل", required=True, default="1", rule=">0",
                  rule_text="المعامل يجب أن يكون أكبر من صفر", note="قيمة وحدة واحدة من العملة بعملة البرنامج"),
            money("ForeignAmount", foreign_caption, note="الإجمالي بعملة المستند (0 للمستندات القديمة)")]


def created_at():
    return datetime_("CreatedAt", "تاريخ الإنشاء", required=True, default="Now()")


def bool_(name, caption, default="False"):
    return Field(name, "BOOL", caption, default=default)


def is_active():
    return bool_("IsActive", "نشط", "True")


def pk_index(*fields):
    return Index("PrimaryKey", list(fields), unique=True, primary=True)


def ux(*fields, ignore_nulls=False):
    return Index("UX_" + "_".join(fields), list(fields), unique=True, ignore_nulls=ignore_nulls)


def ix(*fields):
    return Index("IX_" + "_".join(fields), list(fields))


def doc_totals():
    """Header totals shared by every sales/purchase document."""
    return [
        money("SubTotal", "المجموع قبل الخصم", note="مجموع (الكمية × سعر الوحدة) بدون ضريبة"),
        money("Discount", "الخصم", note="مجموع خصومات الأسطر (خصم الفاتورة يوزَّع على الأسطر)"),
        money("TaxableAmount", "الخاضع للضريبة", note="SubTotal − Discount"),
        money("Tax", "ضريبة القيمة المضافة", note="مجموع ضريبة الأسطر"),
        money("TotalAmount", "الإجمالي شامل الضريبة", note="TaxableAmount + Tax"),
    ]


def zatca_fields(type_code_default):
    """Fields needed for ZATCA e-invoicing (phase 1 now, phase 2 later)."""
    return [
        text("InvoiceSubType", 10, "نوع الفاتورة الضريبية", required=True, default='"SIMPLIFIED"',
             rule='In ("SIMPLIFIED","STANDARD")',
             rule_text="SIMPLIFIED = فاتورة مبسطة، STANDARD = فاتورة ضريبية",
             note="مبسطة للأفراد B2C، ضريبية للمنشآت B2B (للعميل رقم ضريبي)"),
        text("InvoiceTypeCode", 3, "رمز نوع المستند", required=True, default=f'"{type_code_default}"',
             rule='In ("388","381","383")', rule_text="388 فاتورة، 381 إشعار دائن، 383 إشعار مدين"),
        text("InvoiceUUID", 36, "المعرّف الفريد UUID"),
        long_("ICV", "عدّاد الفواتير ICV", note="تسلسل مشترك لكل المستندات المرسلة للهيئة"),
        text("InvoiceHash", 255, "بصمة المستند"),
        text("PreviousInvoiceHash", 255, "بصمة المستند السابق"),
        memo("QRCodeData", "بيانات رمز QR"),
        text("ZatcaStatus", 20, "حالة الإرسال للهيئة", required=True, default='"NOT_SENT"',
             rule='In ("NOT_SENT","PENDING","REPORTED","CLEARED","WARNING","REJECTED")',
             rule_text="حالة غير معروفة"),
        datetime_("ZatcaSubmittedAt", "تاريخ الإرسال للهيئة"),
        memo("ZatcaResponse", "رد الهيئة"),
        text("SignedXmlPath", 255, "مسار ملف XML الموقّع"),
    ]


PAYMENT_TYPE_RULE = 'In ("CASH","CREDIT")'
PAYMENT_TYPE_TEXT = "CASH = نقدي، CREDIT = آجل"
PAID_RULE = "[PaidAmount]<=[TotalAmount] And [RemainingAmount]=[TotalAmount]-[PaidAmount]"
PAID_TEXT = "المدفوع لا يتجاوز الإجمالي، والمتبقي = الإجمالي − المدفوع"
QTY_POSITIVE = ">0"
QTY_POSITIVE_TEXT = "الكمية يجب أن تكون أكبر من صفر"
VAT_CAT_RULE = 'In ("S","Z","E")'
VAT_CAT_TEXT = "S = خاضع 15%، Z = نسبة صفرية، E = معفى"

# --------------------------------------------------------------------------
# Tables
# --------------------------------------------------------------------------
TABLES: List[Table] = [

    # ---------------------------------------------------------------- System
    Table(
        "Settings", "إعدادات المحل",
        "سجل واحد فقط يحتوي بيانات المحل الضريبية وإعدادات التشغيل.",
        [
            long_("SettingID", "رقم الإعداد", required=True, default="1", rule="=1",
                  rule_text="يسمح بسجل إعدادات واحد فقط"),
            text("StoreName", 150, "اسم المحل", required=True),
            text("StoreNameEn", 150, "اسم المحل بالإنجليزية"),
            text("VATNumber", 15, "الرقم الضريبي", rule=VAT_RULE, rule_text=VAT_TEXT),
            text("CRNumber", 20, "السجل التجاري"),
            text("BuildingNo", 10, "رقم المبنى"),
            text("StreetName", 100, "الشارع"),
            text("District", 100, "الحي"),
            text("City", 50, "المدينة"),
            text("PostalCode", 10, "الرمز البريدي"),
            text("AdditionalNo", 10, "الرقم الإضافي"),
            text("CountryCode", 2, "رمز الدولة", required=True, default='"SA"'),
            text("Phone", 20, "الهاتف"),
            text("Email", 100, "البريد الإلكتروني"),
            rate("VATRate", "نسبة الضريبة", default="0.15"),
            bool_("PricesIncludeVAT", "الأسعار شاملة الضريبة", "True"),
            bool_("AllowNegativeStock", "السماح بالبيع بالسالب", "False"),
            text("CurrencyCode", 3, "العملة", required=True, default='"SAR"'),
            long_("DefaultCustomerID", "العميل الافتراضي", required=True, default="1",
                  fk="Customers.CustomerID"),
            byte_("ZatcaPhase", "مرحلة فاتورة", default="1", rule="In (1,2)",
                  rule_text="المرحلة 1 أو 2"),
            text("ZatcaEnvironment", 20, "بيئة الربط مع الهيئة"),
            text("LastInvoiceHash", 255, "بصمة آخر مستند",
                 note="تبدأ بالقيمة الافتراضية التي تحددها الهيئة لأول فاتورة"),
            text("BackupFolder", 255, "مجلد النسخ الاحتياطي"),
            int_("BackupKeepCount", "عدد النسخ المحتفظ بها", required=True, default="30",
                 rule=">=1", rule_text="احتفظ بنسخة واحدة على الأقل"),
            int_("SlowMovingDays", "أيام عدم الحركة", required=True, default="90",
                 rule=">=1", rule_text="عدد الأيام يجب أن يكون 1 أو أكثر"),
            text("ReceiptFooter", 255, "تذييل الفاتورة"),
            text("LogoPath", 255, "مسار الشعار"),
            datetime_("UpdatedAt", "آخر تعديل"),
            # added after the first release (BuildSchema adds missing fields to existing tables)
            text("POSMode", 10, "شاشة البيع", required=True, default='"RETAIL"',
                 rule='In ("RETAIL","RESTAURANT","CAFE")', rule_text="اختر شاشة البيع من القائمة",
                 note="RETAIL = المحلات (باركود)، RESTAURANT = المطاعم (لمس)، CAFE = الكافيهات (لمس)"),
            text("ImagesFolder", 255, "مجلد صور المنتجات",
                 note="المسارات النسبية للصور تُقرأ منه؛ فارغ = مجلد Images بجانب ملف البيانات"),
            text("InvoicePrintMode", 10, "الطباعة عند حفظ الفاتورة", required=True, default='"PREVIEW"',
                 rule='In ("DIRECT","PREVIEW","NONE")', rule_text="اختر طريقة الطباعة من القائمة",
                 note="DIRECT = طباعة مباشرة بدون معاينة، PREVIEW = عرض المعاينة، NONE = بدون طباعة"),
            bool_("AllowAdminCompanyName", "السماح لمدير النظام بتغيير اسم المحل", "False"),
            datetime_("ClosedThrough", "الفترة مقفلة حتى (لا يُضاف ولا يُعدَّل مستند بتاريخ حتى هذا اليوم)"),
            long_("DefaultBankID", "البنك الافتراضي للتحويلات البنكية", fk="Banks.BankID",
                  note="المبالغ المدفوعة أو المستلمة بطريقة «تحويل بنكي» تُقيَّد في حساب هذا البنك"),
            rate("GosiEmployeeRate", "التأمينات: حصة الموظف السعودي", default="0.0975"),
            rate("GosiEmployerRate", "التأمينات: حصة المنشأة عن السعودي", default="0.1175"),
            rate("GosiNonSaudiRate", "التأمينات: حصة المنشأة عن غير السعودي (الأخطار المهنية)", default="0.02"),
            money("GosiMaxWage", "التأمينات: الحد الأعلى للأجر الخاضع", default="45000"),
            int_("CreditBlockDays", "إيقاف البيع الآجل لعميل متأخر أكثر من (يوم)", default="0", rule=">=0",
                 rule_text="عدد الأيام لا يكون سالبًا"),
        ],
        pk=["SettingID"],
        seed_columns=["SettingID", "StoreName", "CountryCode", "VATRate", "PricesIncludeVAT",
                      "AllowNegativeStock", "CurrencyCode", "DefaultCustomerID", "ZatcaPhase",
                      "LastInvoiceHash", "BackupKeepCount", "SlowMovingDays", "ReceiptFooter"],
        seed_rows=[(1, "اسم المحل", "SA", 0.15, True, False, "SAR", 1, 1, ZATCA_INITIAL_PIH,
                    30, 90, "شكرًا لزيارتكم")],
    ),

    Table(
        "Sequences", "عدّادات الترقيم",
        "يولّد أرقامًا تسلسلية بدون فجوات للفواتير والسندات (يُزاد داخل نفس معاملة الحفظ).",
        [
            text("SequenceName", 30, "اسم العدّاد", required=True),
            text("Prefix", 10, "البادئة"),
            long_("NextValue", "الرقم التالي", required=True, default="1", rule=">=1",
                  rule_text="الرقم التالي يجب أن يكون 1 أو أكثر"),
            byte_("PadLength", "عدد الخانات", default="6", rule="Between 0 And 12",
                  rule_text="من 0 إلى 12 خانة"),
            text("Description", 100, "الوصف"),
        ],
        pk=["SequenceName"],
        seed_columns=["SequenceName", "Prefix", "NextValue", "PadLength", "Description"],
        seed_rows=[
            ("SALES_INVOICE", "INV-", 1, 6, "فواتير المبيعات"),
            ("SALES_RETURN", "CRN-", 1, 6, "مرتجعات المبيعات (إشعار دائن)"),
            ("PURCHASE_INVOICE", "PUR-", 1, 6, "فواتير المشتريات"),
            ("PURCHASE_RETURN", "PRT-", 1, 6, "مرتجعات المشتريات"),
            ("CUSTOMER_PAYMENT", "RCV-", 1, 6, "سندات القبض من العملاء"),
            ("SUPPLIER_PAYMENT", "PAY-", 1, 6, "سندات الصرف للموردين"),
            ("EXPENSE", "EXP-", 1, 6, "المصروفات"),
            ("SALES_REP", "REP-", 1, 3, "أكواد المندوبين"),
            ("COMMISSION_RUN", "COM-", 1, 5, "مسيرات العمولات"),
            ("STOCK_COUNT", "CNT-", 1, 5, "جلسات الجرد"),
            ("STOCK_ADJUST", "ADJ-", 1, 6, "حركات المخزون اليدوية"),
            ("PRODUCT_CODE", "P", 1, 5, "أكواد المنتجات"),
            ("ZATCA_ICV", None, 1, 0, "عدّاد ICV لمستندات فاتورة"),
            ("CASH_IN", "CIN-", 1, 6, "سندات قبض النقدية (الخزينة)"),
            ("CASH_OUT", "COT-", 1, 6, "سندات صرف النقدية (الخزينة)"),
            ("CASH_TRANSFER", "TRF-", 1, 6, "التحويل بين الصناديق"),
            ("CASH_CLOSING", "CLS-", 1, 6, "تصفية يومية الكاشير"),
            ("JOURNAL", "JV-", 1, 6, "قيود اليومية"),
            ("MANUAL_ENTRY", "MJ-", 1, 6, "القيود اليدوية"),
            ("BANK_TX", "BNK-", 1, 6, "الحركات البنكية"),
            ("BANK_RECON", "REC-", 1, 6, "التسويات البنكية"),
            ("CHEQUE", "CHQ-", 1, 6, "الشيكات الواردة والصادرة"),
            ("FIXED_ASSET", "FA-", 1, 5, "الأصول الثابتة"),
            ("PAYROLL", "PAY-", 1, 5, "مسيرات الرواتب"),
            ("DEPRECIATION", "DEP-", 1, 6, "قيود الإهلاك الشهرية"),
        ],
        seed_missing=True,
    ),

    Table(
        "Roles", "الأدوار",
        "مستويات الصلاحيات: مدير النظام، مدير، كاشير.",
        [
            long_("RoleID", "رقم الدور", required=True),
            text("RoleCode", 20, "رمز الدور", required=True),
            text("RoleName", 50, "اسم الدور", required=True),
            text("Description", 255, "الوصف"),
        ],
        pk=["RoleID"],
        indexes=[ux("RoleCode")],
        seed_columns=["RoleID", "RoleCode", "RoleName", "Description"],
        seed_rows=[
            (1, "ADMIN", "مدير النظام", "جميع الصلاحيات"),
            (2, "MANAGER", "مدير", "المبيعات والمشتريات والمخزون والتقارير"),
            (3, "CASHIER", "كاشير", "المبيعات والعملاء فقط"),
        ],
    ),

    Table(
        "Permissions", "الصلاحيات",
        "قائمة الصلاحيات التي يمكن منحها للأدوار.",
        [
            text("PermissionKey", 50, "رمز الصلاحية", required=True),
            text("PermissionName", 100, "اسم الصلاحية", required=True),
            text("ModuleName", 50, "القسم"),
            int_("SortOrder", "الترتيب", required=True, default="0"),
        ],
        pk=["PermissionKey"],
        seed_columns=["PermissionKey", "PermissionName", "ModuleName", "SortOrder"],
        seed_rows=[
            ("SALES_POS", "نقطة البيع", "المبيعات", 10),
            ("SALES_VIEW", "عرض وإعادة طباعة الفواتير", "المبيعات", 11),
            ("SALES_RETURN", "مرتجعات المبيعات", "المبيعات", 12),
            ("PRICE_OVERRIDE", "تعديل سعر البيع في الفاتورة", "المبيعات", 13),
            ("DISCOUNT_OVERRIDE", "خصم أعلى من الحد المسموح", "المبيعات", 14),
            ("ALLOW_NEGATIVE_STOCK", "البيع بكمية أكبر من المتوفر", "المبيعات", 15),
            ("CUSTOMERS", "إدارة العملاء", "العملاء", 20),
            ("CUSTOMER_PAYMENTS", "سندات القبض", "العملاء", 21),
            ("PURCHASES", "فواتير المشتريات", "المشتريات", 30),
            ("PURCHASE_RETURN", "مرتجعات المشتريات", "المشتريات", 31),
            ("SUPPLIERS", "إدارة الموردين", "الموردون", 40),
            ("SUPPLIER_PAYMENTS", "سندات الصرف", "الموردون", 41),
            ("PRODUCTS", "إدارة المنتجات والأسعار", "المخزون", 50),
            ("INVENTORY_ADJUST", "إضافة وخصم مخزون يدوي", "المخزون", 51),
            ("STOCK_COUNT", "الجرد", "المخزون", 52),
            ("EXPENSES", "المصروفات", "المصروفات", 60),
            ("CASH_BOX", "الخزينة: سندات القبض والصرف والتحويل والصناديق", "الخزينة", 65),
            ("CASH_CLOSING", "تصفية يومية الكاشير", "الخزينة", 66),
            ("JOURNAL", "قيود اليومية ودليل الحسابات وميزان المراجعة", "الحسابات", 75),
            ("MANUAL_ENTRY", "القيود اليدوية: إضافة وتعديل وحذف", "الحسابات", 76),
            ("PERIOD_CLOSE", "إقفال الفترات والسنة المالية وإعادة فتحها", "الحسابات", 77),
            ("VAT_RETURN", "إقرار ضريبة القيمة المضافة: الاعتماد والسداد", "الحسابات", 78),
            ("BANKS", "البنوك: الحسابات والحركات البنكية والتسوية البنكية", "الخزينة", 67),
            ("CHEQUES", "الشيكات الواردة والصادرة: التسجيل والتحصيل والارتداد", "الخزينة", 68),
            ("FIXED_ASSETS", "الأصول الثابتة والإهلاك", "الحسابات", 79),
            ("PAYROLL", "مسير الرواتب: الإعداد والترحيل والصرف", "الحسابات", 80),
            ("BUDGET", "الموازنة التقديرية: الإعداد والمقارنة بالفعلي", "الحسابات", 81),
            ("CURRENCIES", "العملات وأسعارها", "الحسابات", 84),
            ("SALES_REPS", "المندوبين: البيانات والأهداف والعمولات وتقاريرهم", "المبيعات", 15),
            ("REPORTS", "التقارير التشغيلية", "التقارير", 70),
            ("REPORTS_PROFIT", "تقارير الأرباح والضريبة", "التقارير", 71),
            ("DASHBOARD_FINANCIAL", "الأرقام المالية في لوحة التحكم", "التقارير", 72),
            ("SETTINGS", "إعدادات المحل", "النظام", 80),
            ("USERS", "المستخدمون والصلاحيات", "النظام", 81),
            ("BACKUP", "النسخ الاحتياطي", "النظام", 82),
            ("AUDIT_LOG", "سجل التدقيق: من أضاف أو عدّل أو حذف، والقيم قبل وبعد", "النظام", 83),
        ],
        seed_missing=True,
    ),

    Table(
        "RolePermissions", "صلاحيات الأدوار",
        "ربط كل دور بالصلاحيات الممنوحة له (علاقة متعدد لمتعدد).",
        [
            long_("RoleID", "الدور", required=True, fk="Roles.RoleID", cascade=True),
            text("PermissionKey", 50, "الصلاحية", required=True),
        ],
        pk=["RoleID", "PermissionKey"],
        seed_columns=["RoleID", "PermissionKey"],
        seed_rows=[],   # filled below
    ),

    Table(
        "Employees", "الموظفون والمستخدمون",
        "كل موظف هو مستخدم للنظام؛ لا يُحذف بل يُعطَّل للحفاظ على سجل عملياته.",
        [
            auto("EmployeeID", "رقم الموظف"),
            text("EmployeeName", 100, "اسم الموظف", required=True),
            text("JobTitle", 50, "المسمى الوظيفي"),
            text("Mobile", 20, "الجوال"),
            text("Username", 30, "اسم المستخدم", required=True),
            text("PasswordHash", 64, "بصمة كلمة المرور",
                 note="SHA-256 بصيغة hex؛ كلمة المرور نفسها لا تُحفظ أبدًا"),
            text("PasswordSalt", 32, "ملح التشفير"),
            long_("RoleID", "الدور", required=True, fk="Roles.RoleID"),
            rate("MaxDiscountPercent", "أقصى نسبة خصم", default="0",
                 rule=">=0 And <=1", rule_text="النسبة بين 0% و 100%"),
            bool_("MustChangePassword", "يجب تغيير كلمة المرور", "True"),
            int_("FailedLoginCount", "محاولات الدخول الفاشلة", required=True, default="0"),
            datetime_("LockedUntil", "مقفل حتى"),
            datetime_("LastLoginAt", "آخر دخول"),
            is_active(),
            memo("Notes", "ملاحظات"),
            created_at(),
            long_("CashBoxID", "صندوق النقدية", fk="CashBoxes.CashBoxID",
                  note="تدخل فيه نقدية مبيعاته وسنداته؛ فارغ = أول صندوق كاشير نشط"),
            bool_("IsDeveloper", "المبرمج", "False"),
            bool_("CustomScreens", "صلاحيات شاشات خاصة", "False"),
            # payroll (modPayroll)
            bool_("OnPayroll", "في مسير الرواتب", "False"),
            money("BasicSalary", "الراتب الأساسي"),
            money("HousingAllowance", "بدل السكن"),
            money("TransportAllowance", "بدل النقل"),
            money("OtherAllowance", "بدلات أخرى"),
            bool_("IsSaudi", "سعودي (التأمينات بحصتي الموظف والمنشأة)", "False"),
            money("AdvanceInstallment", "قسط السلفة الشهري", note="0 = يُخصم كل الرصيد في أول مسير"),
            text("NationalID", 15, "رقم الهوية / الإقامة"),
            text("IBAN", 34, "آيبان الموظف"),
            date_("HireDate", "تاريخ التعيين", required=False, default=None),
            long_("CostCenterID", "مركز التكلفة", fk="CostCenters.CostCenterID", note="مبيعاته ومسير راتبه على هذا المركز"),
        ],
        pk=["EmployeeID"],
        indexes=[ux("Username")],
        seed_columns=["EmployeeID", "EmployeeName", "JobTitle", "Username", "RoleID",
                      "MaxDiscountPercent", "MustChangePassword", "IsActive"],
        seed_rows=[(1, "مدير النظام", "مدير النظام", "admin", 1, 1, True, True)],
    ),

    Table(
        "Screens", "الشاشات",
        "كل شاشة في البرنامج، وما ينطبق عليها من إضافة وتعديل وحذف، وصلاحية الدور التي تفتحها.",
        [
            text("ScreenName", 64, "اسم الشاشة في Access", required=True),
            text("ScreenTitle", 100, "الشاشة", required=True),
            text("ModuleName", 50, "القسم"),
            int_("SortOrder", "الترتيب", required=True, default="0"),
            text("PermissionKey", 50, "صلاحية الدور",
                 note="فارغ = متاحة لكل المستخدمين؛ تُستخدم للمستخدم الذي ليست له صلاحيات شاشات خاصة"),
            bool_("HasAdd", "فيها إضافة / حفظ مستند", "False"),
            bool_("HasEdit", "فيها تعديل", "False"),
            bool_("HasDelete", "فيها حذف", "False"),
        ],
        pk=["ScreenName"],
        seed_columns=["ScreenName", "ScreenTitle", "ModuleName", "SortOrder", "PermissionKey",
                      "HasAdd", "HasEdit", "HasDelete"],
        seed_rows=[],   # filled below from SCREEN_LIST
        seed_missing=True,
    ),

    Table(
        "UserScreens", "صلاحيات الشاشات للمستخدم",
        "للمستخدم الذي فُعّلت له «صلاحيات شاشات خاصة»: الشاشات التي يفتحها، والإضافة والتعديل والحذف في كل شاشة.",
        [
            long_("EmployeeID", "المستخدم", required=True, fk="Employees.EmployeeID"),
            text("ScreenName", 64, "الشاشة", required=True),
            bool_("CanOpen", "فتح", "False"),
            bool_("CanAdd", "إضافة", "False"),
            bool_("CanEdit", "تعديل", "False"),
            bool_("CanDelete", "حذف", "False"),
        ],
        pk=["EmployeeID", "ScreenName"],
    ),

    Table(
        "Activations", "تفعيل البرنامج",
        "الأجهزة المفعّل عليها البرنامج: رقم الجهاز وكود التفعيل الصادر من المبرمج.",
        [
            auto("ActivationID", "رقم التفعيل"),
            text("MachineID", 24, "رقم الجهاز", required=True,
                 note="بصمة لوحة الأم والمعالج وقرص النظام (modActivation.MachineID)"),
            text("ActivationCode", 30, "كود التفعيل", required=True),
            text("ComputerName", 64, "اسم الجهاز"),
            datetime_("ActivatedAt", "تاريخ التفعيل", default="Now()"),
            long_("EmployeeID", "فعّله", fk="Employees.EmployeeID"),
        ],
        pk=["ActivationID"],
        indexes=[ux("MachineID")],
    ),

    # ----------------------------------------------------------- Master data
    Table(
        "Categories", "التصنيفات",
        "تصنيفات المنتجات (إلكترونيات، مواد غذائية، ...).",
        [
            auto("CategoryID", "رقم التصنيف"),
            text("CategoryName", 100, "اسم التصنيف", required=True),
            text("Description", 255, "الوصف"),
            is_active(),
            # touch screens (restaurant / café)
            text("ImagePath", 255, "صورة التصنيف"),
            text("TileColor", 10, "لون الزر", required=True, default='"BLUE"',
                 rule='In ("BLUE","GREEN","ORANGE","PURPLE","RED","INDIGO","TEAL","PINK","BROWN","GREY")',
                 rule_text="اختر اللون من القائمة"),
            int_("SortOrder", "ترتيب العرض", required=True, default="0"),
            bool_("IsAddOn", "فئة إضافات", default="False"),
        ],
        pk=["CategoryID"],
        indexes=[ux("CategoryName")],
        seed_columns=["CategoryID", "CategoryName", "Description"],
        seed_rows=[(1, "عام", "تصنيف افتراضي")],
    ),

    Table(
        "Units", "وحدات القياس",
        "وحدات البيع (حبة، كرتون، كيلو، ...) مع رمزها في فاتورة الهيئة.",
        [
            auto("UnitID", "رقم الوحدة"),
            text("UnitName", 30, "اسم الوحدة", required=True),
            text("ZatcaUnitCode", 10, "رمز الوحدة (UN/ECE)"),
            is_active(),
        ],
        pk=["UnitID"],
        indexes=[ux("UnitName")],
        seed_columns=["UnitID", "UnitName", "ZatcaUnitCode"],
        seed_rows=[
            (1, "حبة", "PCE"), (2, "علبة", "BX"), (3, "كرتون", "CT"), (4, "باكيت", "PK"),
            (5, "كيلو", "KGM"), (6, "لتر", "LTR"), (7, "متر", "MTR"), (8, "طقم", "SET"),
        ],
    ),

    Table(
        "PaymentMethods", "طرق الدفع",
        "طرق دفع المبالغ المسددة مع رمزها في فاتورة الهيئة (UNTDID 4461).",
        [
            long_("PaymentMethodID", "رقم الطريقة", required=True),
            text("MethodName", 50, "طريقة الدفع", required=True),
            text("MethodNameEn", 50, "الاسم بالإنجليزية", note="يظهر في الواجهة الإنجليزية"),
            text("ZatcaCode", 5, "رمز الهيئة"),
            int_("SortOrder", "الترتيب", required=True, default="0"),
            is_active(),
        ],
        pk=["PaymentMethodID"],
        indexes=[ux("MethodName")],
        seed_columns=["PaymentMethodID", "MethodName", "ZatcaCode", "SortOrder"],
        seed_rows=[
            (1, "نقدي", "10", 1),
            (2, "مدى / بطاقة", "48", 2),
            (3, "تحويل بنكي", "42", 3),
            (4, "محفظة إلكترونية", "1", 4),
        ],
    ),

    Table(
        "Currencies", "العملات",
        "عملات التعامل. عملة البرنامج هي Settings.CurrencyCode (الريال)، وكل المبالغ تُحفظ بها؛ "
        "المستند بعملة أخرى يحفظ عملته ومعامله ومبلغه بها.",
        [
            text("CurrencyCode", 3, "رمز العملة", required=True, rule="Len([CurrencyCode])=3",
                 rule_text="رمز العملة 3 أحرف (ISO 4217) مثل USD"),
            text("CurrencyName", 50, "اسم العملة", required=True),
            text("CurrencyNameEn", 50, "اسم العملة بالإنجليزية"),
            text("Symbol", 10, "الرمز المختصر"),
            byte_("DecimalPlaces", "عدد الخانات العشرية", default="2", rule="Between 0 And 3",
                  rule_text="من 0 إلى 3 خانات"),
            int_("SortOrder", "الترتيب", required=True, default="0"),
            is_active(),
        ],
        pk=["CurrencyCode"],
        indexes=[ux("CurrencyName")],
        seed_columns=["CurrencyCode", "CurrencyName", "CurrencyNameEn", "Symbol", "DecimalPlaces", "SortOrder"],
        seed_rows=[
            ("SAR", "ريال سعودي", "Saudi Riyal", "ر.س", 2, 1),
            ("USD", "دولار أمريكي", "US Dollar", "$", 2, 2),
            ("EUR", "يورو", "Euro", "€", 2, 3),
            ("AED", "درهم إماراتي", "UAE Dirham", "د.إ", 2, 4),
            ("KWD", "دينار كويتي", "Kuwaiti Dinar", "د.ك", 3, 5),
            ("BHD", "دينار بحريني", "Bahraini Dinar", "د.ب", 3, 6),
            ("QAR", "ريال قطري", "Qatari Riyal", "ر.ق", 2, 7),
            ("OMR", "ريال عماني", "Omani Rial", "ر.ع", 3, 8),
            ("EGP", "جنيه مصري", "Egyptian Pound", "ج.م", 2, 9),
            ("GBP", "جنيه إسترليني", "Pound Sterling", "£", 2, 10),
            ("CNY", "يوان صيني", "Chinese Yuan", "¥", 2, 11),
        ],
        seed_missing=True,
    ),

    Table(
        "CurrencyRates", "أسعار العملات",
        "معامل كل عملة في تاريخ: قيمة وحدة واحدة منها بعملة البرنامج. المستند يأخذ آخر سعر في تاريخه أو قبله.",
        [
            auto("CurrencyRateID", "رقم داخلي"),
            Field("CurrencyCode", "TEXT", "العملة", size=3, required=True, fk="Currencies.CurrencyCode"),
            date_("RateDate", "التاريخ"),
            Field("Rate", "RATE", "المعامل", required=True, default="1", rule=">0",
                  rule_text="المعامل يجب أن يكون أكبر من صفر"),
            text("Notes", 100, "ملاحظات"),
            created_at(),
        ],
        pk=["CurrencyRateID"],
        indexes=[ux("CurrencyCode", "RateDate")],
        seed_columns=["CurrencyRateID", "CurrencyCode", "RateDate", "Rate", "Notes"],
        seed_rows=[
            # the currencies pegged to the dollar; the others are entered by the user
            (1, "USD", "2000-01-01", 3.75, "سعر الربط الرسمي"),
            (2, "AED", "2000-01-01", 1.0211, "مرتبط بالدولار"),
            (3, "BHD", "2000-01-01", 9.9734, "مرتبط بالدولار"),
            (4, "QAR", "2000-01-01", 1.0302, "مرتبط بالدولار"),
            (5, "OMR", "2000-01-01", 9.7529, "مرتبط بالدولار"),
        ],
        seed_missing=True,
    ),

    Table(
        "CashBoxes", "الخزينة والصناديق",
        "الخزينة الرئيسية وصناديق الكاشير. الرصيد لا يُخزَّن: يُحسب من الحركات (qryCashMovements).",
        [
            auto("CashBoxID", "رقم الصندوق"),
            text("BoxName", 50, "اسم الصندوق", required=True),
            text("BoxType", 10, "النوع", required=True, default='"CASHIER"',
                 rule='In ("MAIN","CASHIER")', rule_text="MAIN = خزينة رئيسية، CASHIER = صندوق كاشير"),
            money("OpeningBalance", "الرصيد الافتتاحي"),
            date_("OpeningDate", "تاريخ الرصيد الافتتاحي"),
            is_active(),
            text("Notes", 255, "ملاحظات"),
            created_at(),
        ],
        pk=["CashBoxID"],
        indexes=[ux("BoxName")],
        seed_columns=["CashBoxID", "BoxName", "BoxType"],
        seed_rows=[(1, "الخزينة الرئيسية", "MAIN"), (2, "صندوق الكاشير", "CASHIER")],
    ),

    Table(
        "Suppliers", "الموردون",
        "بيانات الموردين. الرصيد الحالي قيمة مساعدة تُحدَّث بالكود ويمكن إعادة احتسابها من الحركات.",
        [
            auto("SupplierID", "رقم المورد"),
            text("SupplierName", 150, "اسم المورد", required=True),
            text("ContactPerson", 100, "الشخص المسؤول"),
            text("Mobile", 20, "الجوال"),
            text("Phone", 20, "الهاتف"),
            text("Email", 100, "البريد الإلكتروني"),
            text("VATNumber", 15, "الرقم الضريبي", rule=VAT_RULE, rule_text=VAT_TEXT),
            text("CRNumber", 20, "السجل التجاري"),
            text("Address", 255, "العنوان"),
            text("City", 50, "المدينة"),
            money("OpeningBalance", "الرصيد الافتتاحي", rule=None,
                  note="موجب = المحل مدين للمورد"),
            money("CurrentBalance", "الرصيد الحالي", rule=None,
                  note="قيمة مساعدة؛ المرجع هو SupplierBalanceQuery"),
            int_("PaymentTermsDays", "مدة السداد (يوم)", default="30", rule=">=0",
                 rule_text="عدد الأيام لا يكون سالبًا"),
            is_active(),
            memo("Notes", "ملاحظات"),
            created_at(),
            Field("CurrencyCode", "TEXT", "عملة التعامل", size=3, default='"SAR"', fk="Currencies.CurrencyCode",
                  note="تُقترح في فواتيره وسنداته"),
        ],
        pk=["SupplierID"],
        indexes=[ix("SupplierName"), ix("Mobile")],
    ),

    Table(
        "Customers", "العملاء",
        "بيانات العملاء. السجل رقم 1 هو \"عميل نقدي\" الافتراضي ولا يُحذف.",
        [
            auto("CustomerID", "رقم العميل"),
            text("CustomerName", 150, "اسم العميل", required=True),
            text("Mobile", 20, "الجوال"),
            text("Phone", 20, "الهاتف"),
            text("Email", 100, "البريد الإلكتروني"),
            text("VATNumber", 15, "الرقم الضريبي", rule=VAT_RULE, rule_text=VAT_TEXT,
                 note="إذا وُجد تصدر للعميل فاتورة ضريبية B2B"),
            text("CRNumber", 20, "السجل التجاري"),
            text("BuildingNo", 10, "رقم المبنى"),
            text("StreetName", 100, "الشارع"),
            text("District", 100, "الحي"),
            text("City", 50, "المدينة"),
            text("PostalCode", 10, "الرمز البريدي"),
            text("Address", 255, "العنوان"),
            money("OpeningBalance", "الرصيد الافتتاحي", rule=None,
                  note="موجب = العميل مدين للمحل"),
            money("CurrentBalance", "الرصيد الحالي", rule=None,
                  note="قيمة مساعدة؛ المرجع هو CustomerBalanceQuery"),
            bool_("AllowCredit", "يسمح بالبيع الآجل", "True"),
            money("CreditLimit", "حد الائتمان", note="0 = بدون حد"),
            int_("PaymentTermsDays", "مدة السداد (يوم)", default="30", rule=">=0",
                 rule_text="عدد الأيام لا يكون سالبًا"),
            bool_("IsSystem", "سجل نظام", "False"),
            is_active(),
            memo("Notes", "ملاحظات"),
            created_at(),
            long_("SalesRepID", "المندوب", fk="SalesReps.SalesRepID", note="المندوب المسؤول عن العميل: تُنسب له فواتيره وتحصيلاته"),
        ],
        pk=["CustomerID"],
        indexes=[ix("CustomerName"), ix("Mobile")],
        seed_columns=["CustomerID", "CustomerName", "AllowCredit", "IsSystem"],
        seed_rows=[(1, "عميل نقدي", False, True)],
    ),

    Table(
        "Products", "المنتجات",
        "الأصناف وأسعارها. CurrentQuantity قيمة مساعدة؛ المرجع هو مجموع InventoryTransactions.",
        [
            auto("ProductID", "رقم المنتج"),
            text("ProductCode", 30, "كود المنتج", required=True),
            text("Barcode", 50, "الباركود"),
            text("ProductName", 150, "اسم المنتج", required=True),
            text("ProductNameEn", 150, "الاسم بالإنجليزية"),
            long_("CategoryID", "التصنيف", required=True, default="1", fk="Categories.CategoryID"),
            long_("UnitID", "الوحدة", required=True, default="1", fk="Units.UnitID"),
            money("PurchasePrice", "آخر سعر شراء", note="بدون ضريبة"),
            money("AverageCost", "متوسط التكلفة", note="المتوسط المرجّح، يُحدَّث مع كل شراء"),
            money("SellingPrice", "سعر البيع",
                  note="شامل الضريبة إذا كان الإعداد PricesIncludeVAT مفعّلًا"),
            text("VATCategory", 1, "الفئة الضريبية", required=True, default='"S"',
                 rule=VAT_CAT_RULE, rule_text=VAT_CAT_TEXT),
            qty("CurrentQuantity", "الكمية الحالية"),
            qty("MinimumQuantity", "حد إعادة الطلب", rule=">=0",
                rule_text="الحد الأدنى لا يمكن أن يكون سالبًا"),
            long_("SupplierID", "المورد الافتراضي", fk="Suppliers.SupplierID"),
            text("ProductLocation", 50, "مكان المنتج"),
            memo("Notes", "ملاحظات"),
            is_active(),
            created_at(),
            datetime_("UpdatedAt", "آخر تعديل"),
            text("ImagePath", 255, "صورة المنتج", note="لشاشات اللمس؛ مسار كامل أو اسم ملف في مجلد الصور"),
            bool_("TrackStock", "يتابع المخزون", default="True"),
            # café sizes: SellingPrice = small; a price for medium / large makes the product sized
            money("SizePriceM", "سعر الحجم الوسط", required=False, default=None),
            money("SizePriceL", "سعر الحجم الكبير", required=False, default=None),
        ],
        pk=["ProductID"],
        indexes=[ux("ProductCode"), ux("Barcode", ignore_nulls=True), ix("ProductName")],
    ),

    # ----------------------------------------------------------------- Sales
    Table(
        "SalesInvoices", "فواتير المبيعات",
        "رأس فاتورة البيع. لا تُحذف ولا تُعدَّل بعد الحفظ؛ التصحيح بمرتجع.",
        [
            auto("SalesInvoiceID", "رقم داخلي"),
            text("InvoiceNumber", 20, "رقم الفاتورة", required=True),
            datetime_("InvoiceDate", "تاريخ ووقت الفاتورة", required=True, default="Now()"),
            long_("CustomerID", "العميل", required=True, default="1", fk="Customers.CustomerID"),
            long_("EmployeeID", "الكاشير", required=True, fk="Employees.EmployeeID"),
            text("PaymentType", 10, "نوع البيع", required=True, default='"CASH"',
                 rule=PAYMENT_TYPE_RULE, rule_text=PAYMENT_TYPE_TEXT),
            long_("PaymentMethodID", "طريقة الدفع", fk="PaymentMethods.PaymentMethodID"),
            *doc_totals(),
            money("PaidAmount", "المدفوع", note="المبلغ المحتسب من الفاتورة (لا يتجاوز الإجمالي)"),
            money("RemainingAmount", "المتبقي", note="يُضاف إلى رصيد العميل في البيع الآجل"),
            money("AmountTendered", "المبلغ المستلم", note="ما سلّمه العميل نقدًا"),
            money("ChangeDue", "الباقي للعميل"),
            date_("DueDate", "تاريخ الاستحقاق", required=False, default=None),
            text("Notes", 255, "ملاحظات"),
            *zatca_fields("388"),
            created_at(),
            # restaurant / café orders
            text("OrderType", 10, "نوع الطلب", rule='Is Null Or In ("DINE_IN","TAKEAWAY","DELIVERY")',
                 rule_text="نوع الطلب: داخلي أو سفري أو توصيل"),
            text("TableNo", 10, "رقم الطاولة"),
            text("DeliveryPhone", 20, "جوال التوصيل"),
            text("DeliveryAddress", 255, "عنوان التوصيل"),
            text("OrderName", 50, "اسم العميل على الطلب"),
            long_("CashBoxID", "صندوق النقدية", fk="CashBoxes.CashBoxID",
                  note="يُملأ عند الدفع النقدي: المبلغ المدفوع يدخل هذا الصندوق"),
            long_("BankID", "البنك", fk="Banks.BankID", note="المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك"),
            long_("CostCenterID", "مركز التكلفة", fk="CostCenters.CostCenterID", note="من مركز الكاشير، وإلا المركز الافتراضي"),
            long_("SalesRepID", "المندوب", fk="SalesReps.SalesRepID", note="من مندوب العميل، وإلا مندوب المستخدم"),
        ],
        pk=["SalesInvoiceID"],
        indexes=[ux("InvoiceNumber"), ix("InvoiceDate"), ux("InvoiceUUID", ignore_nulls=True),
                 ix("ZatcaStatus")],
        rule=PAID_RULE, rule_text=PAID_TEXT,
    ),

    Table(
        "SalesInvoiceDetails", "تفاصيل فواتير المبيعات",
        "أسطر فاتورة البيع، مع حفظ تكلفة الصنف لحظة البيع لحساب الربح بدقة.",
        [
            auto("SalesDetailID", "رقم السطر الداخلي"),
            long_("SalesInvoiceID", "الفاتورة", required=True,
                  fk="SalesInvoices.SalesInvoiceID", cascade=True),
            int_("LineNumber", "رقم السطر", required=True),
            long_("ProductID", "المنتج", required=True, fk="Products.ProductID"),
            qty("Quantity", "الكمية", rule=QTY_POSITIVE, rule_text=QTY_POSITIVE_TEXT,
                default=None),
            money("UnitPrice", "سعر الوحدة", note="بدون ضريبة"),
            money("Discount", "الخصم", note="قيمة الخصم على السطر بدون ضريبة"),
            money("NetAmount", "الصافي قبل الضريبة", note="Quantity × UnitPrice − Discount"),
            text("VATCategory", 1, "الفئة الضريبية", required=True, default='"S"',
                 rule=VAT_CAT_RULE, rule_text=VAT_CAT_TEXT),
            rate("VATRate", "نسبة الضريبة", default="0.15"),
            money("Tax", "الضريبة"),
            money("LineTotal", "الإجمالي شامل الضريبة", note="NetAmount + Tax"),
            money("UnitCost", "تكلفة الوحدة", note="AverageCost لحظة البيع"),
            text("LineNote", 100, "ملاحظة السطر", note="الحجم والخيارات (الكافيه)"),
        ],
        pk=["SalesDetailID"],
        indexes=[ux("SalesInvoiceID", "LineNumber")],
    ),

    Table(
        "SalesReturns", "مرتجعات المبيعات",
        "إشعار دائن مرتبط بالفاتورة الأصلية؛ يعيد الكمية للمخزون ويعكس المبلغ.",
        [
            auto("SalesReturnID", "رقم داخلي"),
            text("ReturnNumber", 20, "رقم المرتجع", required=True),
            datetime_("ReturnDate", "تاريخ المرتجع", required=True, default="Now()"),
            long_("SalesInvoiceID", "الفاتورة الأصلية", required=True,
                  fk="SalesInvoices.SalesInvoiceID"),
            long_("CustomerID", "العميل", required=True, fk="Customers.CustomerID"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            text("Reason", 255, "سبب الإرجاع", required=True,
                 note="إلزامي في الإشعار الدائن حسب متطلبات الهيئة"),
            text("RefundType", 10, "طريقة رد المبلغ", required=True, default='"CASH"',
                 rule=PAYMENT_TYPE_RULE,
                 rule_text="CASH = رد نقدي، CREDIT = خصم من رصيد العميل"),
            long_("PaymentMethodID", "طريقة الرد", fk="PaymentMethods.PaymentMethodID"),
            *doc_totals(),
            money("RefundedAmount", "المبلغ المردود نقدًا"),
            text("Notes", 255, "ملاحظات"),
            *zatca_fields("381"),
            created_at(),
            long_("CashBoxID", "صندوق النقدية", fk="CashBoxes.CashBoxID",
                  note="الرد النقدي يخرج من هذا الصندوق"),
            long_("BankID", "البنك", fk="Banks.BankID", note="المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك"),
            long_("CostCenterID", "مركز التكلفة", fk="CostCenters.CostCenterID", note="مركز الفاتورة الأصلية"),
            long_("SalesRepID", "المندوب", fk="SalesReps.SalesRepID", note="مندوب الفاتورة الأصلية"),
        ],
        pk=["SalesReturnID"],
        indexes=[ux("ReturnNumber"), ix("ReturnDate"), ux("InvoiceUUID", ignore_nulls=True)],
        rule="[RefundedAmount]<=[TotalAmount]",
        rule_text="المبلغ المردود لا يتجاوز قيمة المرتجع",
    ),

    Table(
        "SalesReturnDetails", "تفاصيل مرتجعات المبيعات",
        "الأسطر المرتجعة، كل سطر يشير إلى سطر الفاتورة الأصلي لمنع إرجاع أكثر مما بيع.",
        [
            auto("ReturnDetailID", "رقم السطر الداخلي"),
            long_("SalesReturnID", "المرتجع", required=True,
                  fk="SalesReturns.SalesReturnID", cascade=True),
            long_("SalesDetailID", "سطر الفاتورة الأصلي", required=True,
                  fk="SalesInvoiceDetails.SalesDetailID"),
            long_("ProductID", "المنتج", required=True, fk="Products.ProductID"),
            qty("Quantity", "الكمية المرتجعة", rule=QTY_POSITIVE, rule_text=QTY_POSITIVE_TEXT,
                default=None),
            money("UnitPrice", "سعر الوحدة"),
            money("Discount", "الخصم"),
            money("NetAmount", "الصافي قبل الضريبة"),
            text("VATCategory", 1, "الفئة الضريبية", required=True, default='"S"',
                 rule=VAT_CAT_RULE, rule_text=VAT_CAT_TEXT),
            rate("VATRate", "نسبة الضريبة", default="0.15"),
            money("Tax", "الضريبة"),
            money("LineTotal", "الإجمالي شامل الضريبة"),
            money("UnitCost", "تكلفة الوحدة", note="نفس تكلفة سطر البيع الأصلي"),
            bool_("ReturnToStock", "يعاد للمخزون", "True"),
        ],
        pk=["ReturnDetailID"],
    ),

    # ------------------------------------------------------------- Purchases
    Table(
        "PurchaseInvoices", "فواتير المشتريات",
        "رأس فاتورة الشراء من المورد.",
        [
            auto("PurchaseInvoiceID", "رقم داخلي"),
            text("InvoiceNumber", 20, "رقم الفاتورة الداخلي", required=True),
            text("SupplierInvoiceNo", 30, "رقم فاتورة المورد"),
            datetime_("InvoiceDate", "تاريخ الفاتورة", required=True, default="Now()"),
            long_("SupplierID", "المورد", required=True, fk="Suppliers.SupplierID"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            text("PaymentType", 10, "نوع الشراء", required=True, default='"CASH"',
                 rule=PAYMENT_TYPE_RULE, rule_text=PAYMENT_TYPE_TEXT),
            long_("PaymentMethodID", "طريقة الدفع", fk="PaymentMethods.PaymentMethodID"),
            *doc_totals(),
            money("PaidAmount", "المدفوع"),
            money("RemainingAmount", "المتبقي", note="يُضاف إلى رصيد المورد"),
            date_("DueDate", "تاريخ الاستحقاق", required=False, default=None),
            text("Notes", 255, "ملاحظات"),
            created_at(),
            long_("CashBoxID", "صندوق النقدية", fk="CashBoxes.CashBoxID",
                  note="المدفوع نقدًا يخرج من هذا الصندوق"),
            long_("BankID", "البنك", fk="Banks.BankID", note="المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك"),
            *fx_fields(),
        ],
        pk=["PurchaseInvoiceID"],
        indexes=[ux("InvoiceNumber"), ix("InvoiceDate"), ix("SupplierInvoiceNo")],
        rule=PAID_RULE, rule_text=PAID_TEXT,
    ),

    Table(
        "PurchaseInvoiceDetails", "تفاصيل فواتير المشتريات",
        "أسطر فاتورة الشراء.",
        [
            auto("PurchaseDetailID", "رقم السطر الداخلي"),
            long_("PurchaseInvoiceID", "الفاتورة", required=True,
                  fk="PurchaseInvoices.PurchaseInvoiceID", cascade=True),
            int_("LineNumber", "رقم السطر", required=True),
            long_("ProductID", "المنتج", required=True, fk="Products.ProductID"),
            qty("Quantity", "الكمية", rule=QTY_POSITIVE, rule_text=QTY_POSITIVE_TEXT,
                default=None),
            money("UnitCost", "تكلفة الوحدة", note="بدون ضريبة"),
            money("Discount", "الخصم"),
            money("NetAmount", "الصافي قبل الضريبة"),
            rate("VATRate", "نسبة الضريبة", default="0.15"),
            money("Tax", "ضريبة المدخلات"),
            money("LineTotal", "الإجمالي شامل الضريبة"),
        ],
        pk=["PurchaseDetailID"],
        indexes=[ux("PurchaseInvoiceID", "LineNumber")],
    ),

    Table(
        "PurchaseReturns", "مرتجعات المشتريات",
        "إرجاع بضاعة للمورد مرتبط بفاتورة الشراء الأصلية.",
        [
            auto("PurchaseReturnID", "رقم داخلي"),
            text("ReturnNumber", 20, "رقم المرتجع", required=True),
            datetime_("ReturnDate", "تاريخ المرتجع", required=True, default="Now()"),
            long_("PurchaseInvoiceID", "فاتورة الشراء الأصلية", required=True,
                  fk="PurchaseInvoices.PurchaseInvoiceID"),
            long_("SupplierID", "المورد", required=True, fk="Suppliers.SupplierID"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            text("Reason", 255, "سبب الإرجاع", required=True),
            text("RefundType", 10, "طريقة الاسترداد", required=True, default='"CREDIT"',
                 rule=PAYMENT_TYPE_RULE,
                 rule_text="CASH = استرداد نقدي، CREDIT = خصم من رصيد المورد"),
            long_("PaymentMethodID", "طريقة الاسترداد", fk="PaymentMethods.PaymentMethodID"),
            *doc_totals(),
            money("RefundedAmount", "المبلغ المسترد نقدًا"),
            text("Notes", 255, "ملاحظات"),
            created_at(),
            long_("CashBoxID", "صندوق النقدية", fk="CashBoxes.CashBoxID",
                  note="الاسترداد النقدي يدخل هذا الصندوق"),
            long_("BankID", "البنك", fk="Banks.BankID", note="المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك"),
            *fx_fields(),
        ],
        pk=["PurchaseReturnID"],
        indexes=[ux("ReturnNumber"), ix("ReturnDate")],
        rule="[RefundedAmount]<=[TotalAmount]",
        rule_text="المبلغ المسترد لا يتجاوز قيمة المرتجع",
    ),

    Table(
        "PurchaseReturnDetails", "تفاصيل مرتجعات المشتريات",
        "الأسطر المرتجعة للمورد.",
        [
            auto("PurchaseReturnDetailID", "رقم السطر الداخلي"),
            long_("PurchaseReturnID", "المرتجع", required=True,
                  fk="PurchaseReturns.PurchaseReturnID", cascade=True),
            long_("PurchaseDetailID", "سطر الفاتورة الأصلي", required=True,
                  fk="PurchaseInvoiceDetails.PurchaseDetailID"),
            long_("ProductID", "المنتج", required=True, fk="Products.ProductID"),
            qty("Quantity", "الكمية المرتجعة", rule=QTY_POSITIVE, rule_text=QTY_POSITIVE_TEXT,
                default=None),
            money("UnitCost", "تكلفة الوحدة"),
            money("Discount", "الخصم"),
            money("NetAmount", "الصافي قبل الضريبة"),
            rate("VATRate", "نسبة الضريبة", default="0.15"),
            money("Tax", "الضريبة"),
            money("LineTotal", "الإجمالي شامل الضريبة"),
        ],
        pk=["PurchaseReturnDetailID"],
    ),

    # -------------------------------------------------------------- Payments
    Table(
        "CustomerPayments", "دفعات العملاء (سندات القبض)",
        "المبالغ المستلمة من العملاء لسداد أرصدتهم الآجلة.",
        [
            auto("PaymentID", "رقم داخلي"),
            text("PaymentNumber", 20, "رقم السند", required=True),
            long_("CustomerID", "العميل", required=True, fk="Customers.CustomerID"),
            datetime_("PaymentDate", "تاريخ الدفعة", required=True, default="Now()"),
            money("Amount", "المبلغ", rule=">0", rule_text="المبلغ يجب أن يكون أكبر من صفر"),
            long_("PaymentMethodID", "طريقة الدفع", required=True, default="1",
                  fk="PaymentMethods.PaymentMethodID"),
            long_("SalesInvoiceID", "عن فاتورة (اختياري)", fk="SalesInvoices.SalesInvoiceID"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            text("Notes", 255, "ملاحظات"),
            created_at(),
            long_("CashBoxID", "صندوق النقدية", fk="CashBoxes.CashBoxID",
                  note="المبلغ النقدي يدخل هذا الصندوق"),
            long_("BankID", "البنك", fk="Banks.BankID", note="المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك"),
            *fx_fields(),
            long_("SalesRepID", "المندوب", fk="SalesReps.SalesRepID", note="المحصِّل: مندوب العميل، وإلا مندوب المستخدم"),
        ],
        pk=["PaymentID"],
        indexes=[ux("PaymentNumber"), ix("PaymentDate")],
    ),

    Table(
        "SupplierPayments", "دفعات الموردين (سندات الصرف)",
        "المبالغ المدفوعة للموردين لسداد أرصدتهم.",
        [
            auto("PaymentID", "رقم داخلي"),
            text("PaymentNumber", 20, "رقم السند", required=True),
            long_("SupplierID", "المورد", required=True, fk="Suppliers.SupplierID"),
            datetime_("PaymentDate", "تاريخ الدفعة", required=True, default="Now()"),
            money("Amount", "المبلغ", rule=">0", rule_text="المبلغ يجب أن يكون أكبر من صفر"),
            long_("PaymentMethodID", "طريقة الدفع", required=True, default="1",
                  fk="PaymentMethods.PaymentMethodID"),
            long_("PurchaseInvoiceID", "عن فاتورة (اختياري)",
                  fk="PurchaseInvoices.PurchaseInvoiceID"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            text("Notes", 255, "ملاحظات"),
            created_at(),
            long_("CashBoxID", "صندوق النقدية", fk="CashBoxes.CashBoxID",
                  note="المبلغ النقدي يخرج من هذا الصندوق"),
            long_("BankID", "البنك", fk="Banks.BankID", note="المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك"),
            *fx_fields(),
        ],
        pk=["PaymentID"],
        indexes=[ux("PaymentNumber"), ix("PaymentDate")],
    ),

    # ----------------------------------------------------------------- Banks
    Table(
        "Banks", "البنوك",
        "كل حساب بنكي للمحل. حسابه في الدليل 120000 + رقمه تحت «الحسابات البنكية» (1210).",
        [
            auto("BankID", "رقم البنك"),
            text("BankName", 100, "اسم البنك / الحساب", required=True),
            text("AccountNo", 30, "رقم الحساب"),
            text("IBAN", 34, "الآيبان"),
            money("OpeningBalance", "الرصيد الافتتاحي", rule=None, note="رصيد الحساب في البنك عند بدء استخدام البرنامج"),
            date_("OpeningDate", "تاريخ الرصيد الافتتاحي"),
            is_active(),
            text("Notes", 255, "ملاحظات"),
            created_at(),
        ],
        pk=["BankID"],
        indexes=[ux("BankName")],
    ),

    Table(
        "BankTransactions", "الحركات البنكية",
        "إيداع نقدية من صندوق، سحب إلى صندوق، تسوية تحصيلات مدى (بعمولتها)، تحويل بين بنكين، "
        "وحركات أخرى (عمولات، فوائد، قروض...) بحساب مقابل.",
        [
            auto("BankTxID", "رقم داخلي"),
            text("TxNumber", 20, "رقم الحركة", required=True),
            date_("TxDate", "التاريخ"),
            text("TxType", 12, "النوع", required=True,
                 rule='In ("DEPOSIT","WITHDRAW","SETTLEMENT","TRANSFER","OTHER_IN","OTHER_OUT")',
                 rule_text="إيداع، سحب، تسوية مدى، تحويل، وارد آخر، صادر آخر"),
            long_("BankID", "البنك", required=True, fk="Banks.BankID"),
            long_("ToBankID", "إلى بنك", fk="Banks.BankID"),
            long_("CashBoxID", "الصندوق", fk="CashBoxes.CashBoxID"),
            long_("CounterAccount", "الحساب المقابل", fk="Accounts.AccountCode"),
            money("Amount", "المبلغ", rule=">0", rule_text="المبلغ يجب أن يكون أكبر من صفر",
                  note="في تسوية مدى: إجمالي التحصيلات قبل العمولة"),
            money("FeeAmount", "العمولة", note="تسوية مدى: عمولة البنك بدون ضريبة"),
            money("FeeVAT", "ضريبة العمولة", note="ضريبة مدخلات على العمولة"),
            text("Reference", 40, "مرجع البنك"),
            text("Description", 255, "البيان"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            created_at(),
        ],
        pk=["BankTxID"],
        indexes=[ux("TxNumber"), ix("TxDate")],
        rule="[FeeAmount]+[FeeVAT]<[Amount] Or [TxType]<>\"SETTLEMENT\"",
        rule_text="العمولة وضريبتها أقل من مبلغ التسوية",
    ),

    Table(
        "Cheques", "الشيكات الواردة والصادرة",
        "شيك وارد من عميل (يسدد رصيده ويبقى في «شيكات تحت التحصيل» حتى يُحصَّل في البنك أو يرتد) "
        "أو صادر لمورد (يسدد رصيده ويبقى في «أوراق الدفع» حتى يصرفه البنك أو يرتد).",
        [
            auto("ChequeID", "رقم داخلي"),
            text("ChequeRef", 20, "رقم القيد الداخلي", required=True),
            text("Direction", 3, "الاتجاه", required=True, rule='In ("IN","OUT")', rule_text="IN = وارد، OUT = صادر"),
            long_("CustomerID", "العميل (الوارد)", fk="Customers.CustomerID"),
            long_("SupplierID", "المورد (الصادر)", fk="Suppliers.SupplierID"),
            text("ChequeNo", 30, "رقم الشيك", required=True),
            text("DrawerBank", 100, "بنك الساحب (الوارد)"),
            long_("BankID", "بنكنا", fk="Banks.BankID",
                  note="الوارد: البنك الذي حُصِّل فيه؛ الصادر: البنك المسحوب عليه"),
            date_("IssueDate", "تاريخ الاستلام / الإصدار"),
            date_("DueDate", "تاريخ الاستحقاق"),
            money("Amount", "المبلغ", rule=">0", rule_text="المبلغ يجب أن يكون أكبر من صفر"),
            text("Status", 10, "الحالة", required=True, default='"PENDING"', rule='In ("PENDING","COLLECTED","BOUNCED")',
                 rule_text="PENDING = تحت التحصيل، COLLECTED = محصَّل / مصروف، BOUNCED = مرتد"),
            date_("StatusDate", "تاريخ التحصيل / الارتداد", required=False, default=None),
            text("Notes", 255, "ملاحظات"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            created_at(),
        ],
        pk=["ChequeID"],
        indexes=[ux("ChequeRef"), ix("DueDate"), ix("ChequeNo")],
        rule="([Direction]=\"IN\" And [CustomerID] Is Not Null And [SupplierID] Is Null) Or "
             "([Direction]=\"OUT\" And [SupplierID] Is Not Null And [CustomerID] Is Null)",
        rule_text="الشيك الوارد لعميل، والصادر لمورد",
    ),

    # ---------------------------------------------------------- Fixed assets
    Table(
        "FixedAssets", "الأصول الثابتة",
        "سجل الأصول: التكلفة وحساب الأصل، والعمر الإنتاجي بالأشهر، والقيمة المتبقية، ومصدر الشراء. "
        "يُهلك بالقسط الثابت شهريًا، ويُستبعد عند بيعه أو تلفه.",
        [
            auto("AssetID", "رقم داخلي"),
            text("AssetCode", 20, "رقم الأصل", required=True),
            text("AssetName", 150, "اسم الأصل", required=True),
            long_("AssetAccount", "حساب الأصل", required=True, fk="Accounts.AccountCode",
                  note="حساب فرعي من الأصول غير المتداولة (12)، مثل الأثاث أو الأجهزة"),
            long_("CostCenterID", "مركز التكلفة", fk="CostCenters.CostCenterID", note="يذهب إليه إهلاك الأصل وربح أو خسارة بيعه"),
            date_("PurchaseDate", "تاريخ الشراء"),
            money("Cost", "التكلفة بدون الضريبة", rule=">0", rule_text="التكلفة أكبر من صفر"),
            money("InputVAT", "ضريبة المدخلات"),
            money("SalvageValue", "القيمة المتبقية في آخر العمر"),
            int_("UsefulLifeMonths", "العمر الإنتاجي (شهر)", required=True, rule=">0",
                 rule_text="العمر الإنتاجي شهر على الأقل"),
            date_("DepStartDate", "بداية الإهلاك (يُهلك من شهر هذا التاريخ)"),
            text("SourceType", 10, "مصدر الشراء", required=True, default='"BANK"', rule='In ("OPENING","BANK","CASHBOX","ACCOUNT")',
                 rule_text="رصيد افتتاحي، بنك، صندوق، أو حساب آخر"),
            long_("BankID", "البنك", fk="Banks.BankID"),
            long_("CashBoxID", "الصندوق", fk="CashBoxes.CashBoxID"),
            long_("CounterAccount", "الحساب الدائن (حساب آخر)", fk="Accounts.AccountCode"),
            money("OpeningAccumDep", "إهلاك سابق (للأصول الموجودة قبل البرنامج)"),
            text("Status", 10, "الحالة", required=True, default='"ACTIVE"', rule='In ("ACTIVE","DISPOSED")',
                 rule_text="ACTIVE = قائم، DISPOSED = مستبعد"),
            date_("DisposalDate", "تاريخ البيع / الاستبعاد", required=False, default=None),
            money("DisposalProceeds", "ثمن البيع"),
            text("DisposalTo", 10, "استلام الثمن", rule='Is Null Or In ("NONE","BANK","CASHBOX")',
                 rule_text="بدون ثمن، بنك، صندوق"),
            long_("DisposalBankID", "بنك استلام الثمن", fk="Banks.BankID"),
            long_("DisposalCashBoxID", "صندوق استلام الثمن", fk="CashBoxes.CashBoxID"),
            money("DisposalAccumDep", "مجمع الإهلاك يوم الاستبعاد"),
            text("Notes", 255, "ملاحظات"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            created_at(),
        ],
        pk=["AssetID"],
        indexes=[ux("AssetCode"), ix("AssetAccount")],
        rule="[SalvageValue]+[OpeningAccumDep]<=[Cost]",
        rule_text="القيمة المتبقية والإهلاك السابق لا يتجاوزان التكلفة",
    ),

    Table(
        "DepreciationRuns", "قيود الإهلاك الشهرية",
        "قيد إهلاك كل شهر: مصروف الإهلاك (5600) مدين ومجمع الإهلاك (1790) دائن، بسطر لكل أصل.",
        [
            auto("RunID", "رقم داخلي"),
            text("RunNumber", 20, "رقم القيد", required=True),
            date_("RunMonth", "الشهر (آخر يوم فيه)"),
            money("TotalAmount", "إجمالي الإهلاك"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            created_at(),
        ],
        pk=["RunID"],
        indexes=[ux("RunNumber"), ux("RunMonth")],
    ),

    Table(
        "AssetDepreciations", "إهلاك كل أصل في كل شهر",
        "أسطر قيد الإهلاك الشهري.",
        [
            auto("DepreciationID", "رقم داخلي"),
            long_("RunID", "قيد الإهلاك", required=True, fk="DepreciationRuns.RunID", cascade=True),
            int_("LineNo", "رقم السطر في القيد", required=True),
            long_("AssetID", "الأصل", required=True, fk="FixedAssets.AssetID"),
            money("Amount", "الإهلاك", rule=">0", rule_text="الإهلاك أكبر من صفر"),
        ],
        pk=["DepreciationID"],
        indexes=[ux("RunID", "AssetID"), ix("AssetID")],
    ),

    # ---------------------------------------------------------- Cost centres
    Table(
        "CostCenters", "مراكز التكلفة والفروع",
        "الفروع أو الأقسام (التجزئة، المطعم، المقهى...). تُوزَّع عليها الإيرادات والمصروفات في القيود، "
        "ومنها قائمة دخل لكل مركز.",
        [
            auto("CostCenterID", "رقم المركز"),
            text("CenterCode", 10, "رمز المركز", required=True),
            text("CenterName", 100, "اسم المركز", required=True),
            bool_("IsDefault", "المركز الافتراضي", "False"),
            is_active(),
            text("Notes", 255, "ملاحظات"),
            created_at(),
        ],
        pk=["CostCenterID"],
        indexes=[ux("CenterCode"), ux("CenterName")],
    ),

    # ----------------------------------------------------------------- Budget
    Table(
        "SalesReps", "المندوبين",
        "مندوبو المبيعات: العملاء المسندون لكل مندوب، ومبيعاته وتحصيلاته وهدفه الشهري وعمولته.",
        [
            auto("SalesRepID", "رقم داخلي"),
            text("RepCode", 20, "كود المندوب", required=True),
            text("RepName", 100, "اسم المندوب", required=True),
            text("RepNameEn", 100, "الاسم بالإنجليزية"),
            text("Mobile", 20, "الجوال"),
            long_("EmployeeID", "مستخدم البرنامج", fk="Employees.EmployeeID",
                  note="مبيعات هذا المستخدم لعميل بلا مندوب تُنسب لهذا المندوب"),
            text("Region", 50, "المنطقة / خط السير"),
            long_("CostCenterID", "مركز التكلفة", fk="CostCenters.CostCenterID", note="مركز قيد عمولته"),
            rate("CommissionRate", "نسبة العمولة", default="0"),
            text("CommissionBase", 10, "أساس العمولة", required=True, default='"SALES"', rule='In ("SALES","COLLECTION")',
                 rule_text="SALES = صافي المبيعات، COLLECTION = التحصيل"),
            is_active(),
            text("Notes", 255, "ملاحظات"),
            created_at(),
        ],
        pk=["SalesRepID"],
        indexes=[ux("RepCode"), ux("RepName")],
    ),

    Table(
        "SalesRepTargets", "أهداف المندوبين",
        "الهدف الشهري لمبيعات كل مندوب (صافي المبيعات بدون الضريبة).",
        [
            auto("TargetID", "رقم داخلي"),
            long_("SalesRepID", "المندوب", required=True, fk="SalesReps.SalesRepID", cascade=True),
            int_("TargetYear", "السنة", required=True, rule="Between 2000 And 2100", rule_text="سنة غير صحيحة"),
            byte_("TargetMonth", "الشهر", default="1", rule="Between 1 And 12", rule_text="الشهر من 1 إلى 12"),
            money("TargetAmount", "الهدف"),
        ],
        pk=["TargetID"],
        indexes=[ux("SalesRepID", "TargetYear", "TargetMonth")],
    ),

    Table(
        "CommissionRuns", "مسيرات العمولات",
        "عمولات المندوبين لشهر: مسودة تُعدَّل، ثم يُرحَّل قيدها (مصروف العمولات 5530 على عمولات مستحقة 2330)، "
        "وتُصرف بسند صرف نقدية من بند «صرف عمولة مندوب».",
        [
            auto("CommissionRunID", "رقم داخلي"),
            text("RunNumber", 20, "رقم المسير", required=True),
            date_("RunMonth", "الشهر (آخر يوم)"),
            text("Status", 10, "الحالة", required=True, default='"DRAFT"', rule='In ("DRAFT","POSTED")',
                 rule_text="DRAFT = مسودة، POSTED = مرحَّل"),
            money("TotalAmount", "إجمالي العمولات"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            datetime_("PostedAt", "تاريخ الترحيل"),
            text("Notes", 255, "ملاحظات"),
            created_at(),
        ],
        pk=["CommissionRunID"],
        indexes=[ux("RunNumber"), ux("RunMonth")],
    ),

    Table(
        "CommissionLines", "أسطر مسير العمولات",
        "عمولة كل مندوب في الشهر: صافي مبيعاته وتحصيلاته، والأساس والنسبة، والتعديل اليدوي.",
        [
            auto("CommissionLineID", "رقم داخلي"),
            long_("CommissionRunID", "المسير", required=True, fk="CommissionRuns.CommissionRunID", cascade=True),
            long_("SalesRepID", "المندوب", required=True, fk="SalesReps.SalesRepID"),
            text("RepName", 100, "اسم المندوب"),
            long_("CostCenterID", "مركز التكلفة", fk="CostCenters.CostCenterID"),
            money("NetSales", "صافي المبيعات", rule=None),
            money("Collections", "التحصيل", rule=None),
            text("CommissionBase", 10, "الأساس", required=True, default='"SALES"'),
            money("BaseAmount", "مبلغ الأساس", rule=None),
            rate("CommissionRate", "النسبة", default="0"),
            money("Adjustment", "تعديل (+/-)", rule=None),
            money("Commission", "العمولة", note="الأساس × النسبة + التعديل، ولا تقل عن صفر"),
            text("Notes", 150, "ملاحظات"),
        ],
        pk=["CommissionLineID"],
        indexes=[ux("CommissionRunID", "SalesRepID")],
    ),

    Table(
        "Budgets", "الموازنات التقديرية",
        "موازنة كل سنة: مبلغ شهري لكل حساب إيرادات أو مصروفات، ويمكن لكل مركز تكلفة.",
        [
            auto("BudgetID", "رقم داخلي"),
            int_("BudgetYear", "السنة", required=True),
            text("BudgetName", 100, "اسم الموازنة", required=True),
            text("Notes", 255, "ملاحظات"),
            long_("EmployeeID", "أعدّها", required=True, fk="Employees.EmployeeID"),
            created_at(),
        ],
        pk=["BudgetID"],
        indexes=[ux("BudgetYear")],
    ),

    Table(
        "BudgetLines", "أسطر الموازنة",
        "حساب (ومركز تكلفة اختياري) ومبلغه في كل شهر. بلا مركز = الحساب في كل المراكز.",
        [
            auto("BudgetLineID", "رقم داخلي"),
            long_("BudgetID", "الموازنة", required=True, fk="Budgets.BudgetID", cascade=True),
            long_("AccountCode", "الحساب", required=True, fk="Accounts.AccountCode"),
            long_("CostCenterID", "مركز التكلفة", fk="CostCenters.CostCenterID", note="فارغ = كل المراكز"),
            money("M1", "الشهر 1"),
            money("M2", "الشهر 2"),
            money("M3", "الشهر 3"),
            money("M4", "الشهر 4"),
            money("M5", "الشهر 5"),
            money("M6", "الشهر 6"),
            money("M7", "الشهر 7"),
            money("M8", "الشهر 8"),
            money("M9", "الشهر 9"),
            money("M10", "الشهر 10"),
            money("M11", "الشهر 11"),
            money("M12", "الشهر 12"),
            text("Notes", 150, "ملاحظات"),
        ],
        pk=["BudgetLineID"],
        indexes=[ix("BudgetID"), ix("AccountCode")],
    ),

    # --------------------------------------------------------------- Payroll
    Table(
        "PayrollRuns", "مسيرات الرواتب",
        "مسير كل شهر: مسودة تُعدَّل، ثم يُرحَّل قيده في آخر يوم من الشهر، ثم يُسجَّل صرفه من بنك أو صندوق.",
        [
            auto("PayrollRunID", "رقم داخلي"),
            text("RunNumber", 20, "رقم المسير", required=True),
            date_("PayMonth", "الشهر (آخر يوم فيه)"),
            text("Status", 10, "الحالة", required=True, default='"DRAFT"', rule='In ("DRAFT","POSTED")',
                 rule_text="DRAFT = مسودة، POSTED = مرحَّل"),
            date_("PaidDate", "تاريخ الصرف", required=False, default=None),
            text("PaidFrom", 10, "الصرف من", rule='Is Null Or In ("BANK","CASHBOX")', rule_text="بنك أو صندوق"),
            long_("BankID", "البنك", fk="Banks.BankID"),
            long_("CashBoxID", "الصندوق", fk="CashBoxes.CashBoxID"),
            money("PaidAmount", "المبلغ المصروف"),
            text("Notes", 255, "ملاحظات"),
            long_("EmployeeID", "أعدّه", required=True, fk="Employees.EmployeeID"),
            created_at(),
        ],
        pk=["PayrollRunID"],
        indexes=[ux("RunNumber"), ux("PayMonth")],
    ),

    Table(
        "PayrollLines", "أسطر مسير الرواتب",
        "سطر كل موظف: الراتب والبدلات والإضافي والخصومات والتأمينات وصافي الراتب.",
        [
            auto("PayrollLineID", "رقم داخلي"),
            long_("PayrollRunID", "المسير", required=True, fk="PayrollRuns.PayrollRunID", cascade=True),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            text("EmployeeName", 100, "اسم الموظف"),
            long_("CostCenterID", "مركز التكلفة", fk="CostCenters.CostCenterID", note="مركز الموظف يوم إنشاء المسير"),
            bool_("IsSaudi", "سعودي", "False"),
            money("Basic", "الأساسي"),
            money("Housing", "بدل السكن"),
            money("OtherAllow", "بدلات أخرى (النقل وغيره)"),
            money("Overtime", "العمل الإضافي"),
            money("Additions", "مكافآت وإضافات"),
            money("AbsenceDeduction", "خصم الغياب"),
            money("AdvanceDeduction", "خصم السلفة"),
            money("OtherDeduction", "خصومات أخرى (جزاءات)"),
            money("GosiWage", "الأجر الخاضع للتأمينات"),
            money("GosiEmployee", "التأمينات - حصة الموظف"),
            money("GosiEmployer", "التأمينات - حصة المنشأة"),
            money("NetPay", "صافي الراتب", rule=None,
                  note="قد يكون سالبًا في المسودة؛ الترحيل يرفضه (الخصومات أكبر من الراتب)"),
            text("Notes", 150, "ملاحظات"),
        ],
        pk=["PayrollLineID"],
        indexes=[ux("PayrollRunID", "EmployeeID")],
    ),

    Table(
        "BankReconciliations", "التسويات البنكية",
        "مطابقة كشف البنك في تاريخ مع الدفاتر: رصيد الكشف، والرصيد في الدفاتر، والحركات غير الظاهرة في الكشف.",
        [
            auto("ReconciliationID", "رقم داخلي"),
            text("ReconNumber", 20, "رقم التسوية", required=True),
            long_("BankID", "البنك", required=True, fk="Banks.BankID"),
            date_("StatementDate", "تاريخ كشف البنك"),
            money("StatementBalance", "رصيد كشف البنك", rule=None),
            money("BookBalance", "الرصيد في الدفاتر في التاريخ", rule=None),
            money("Outstanding", "حركات لم تظهر في الكشف (صافي)", rule=None),
            money("Difference", "الفرق", rule=None),
            text("Status", 10, "الحالة", required=True, default='"OPEN"', rule='In ("OPEN","DONE")',
                 rule_text="OPEN = جارية، DONE = معتمدة"),
            text("Notes", 255, "ملاحظات"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            created_at(),
        ],
        pk=["ReconciliationID"],
        indexes=[ux("ReconNumber"), ix("BankID")],
    ),

    Table(
        "BankClearings", "حركات الدفاتر المطابقة لكشف البنك",
        "كل عملية قيدها على حساب البنك ظهرت في كشف البنك: نوع العملية ورقمها ومبلغها يوم المطابقة.",
        [
            auto("ClearingID", "رقم داخلي"),
            long_("ReconciliationID", "التسوية", required=True, fk="BankReconciliations.ReconciliationID",
                  cascade=True),
            long_("BankID", "البنك", required=True, fk="Banks.BankID"),
            text("SourceType", 20, "نوع العملية", required=True),
            long_("SourceID", "رقم العملية", required=True),
            money("ClearedAmount", "المبلغ يوم المطابقة (مدين موجب)", rule=None),
            created_at(),
        ],
        pk=["ClearingID"],
        indexes=[ux("BankID", "SourceType", "SourceID")],
    ),

    Table(
        "CustomerAllocations", "ربط سندات القبض بالفواتير",
        "كم من سند القبض سدّد كل فاتورة آجلة. ما لا يُربط بفاتورة يسدد أقدم الفواتير استحقاقًا.",
        [
            auto("AllocationID", "رقم داخلي"),
            long_("PaymentID", "سند القبض", required=True, fk="CustomerPayments.PaymentID", cascade=True),
            long_("SalesInvoiceID", "الفاتورة", required=True, fk="SalesInvoices.SalesInvoiceID"),
            money("Amount", "المبلغ", rule=">0", rule_text="المبلغ يجب أن يكون أكبر من صفر"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            created_at(),
        ],
        pk=["AllocationID"],
        indexes=[ux("PaymentID", "SalesInvoiceID"), ix("SalesInvoiceID")],
    ),

    Table(
        "SupplierAllocations", "ربط سندات الصرف بفواتير الشراء",
        "كم من سند الصرف سدّد كل فاتورة شراء آجلة. ما لا يُربط بفاتورة يسدد أقدم الفواتير استحقاقًا.",
        [
            auto("AllocationID", "رقم داخلي"),
            long_("PaymentID", "سند الصرف", required=True, fk="SupplierPayments.PaymentID", cascade=True),
            long_("PurchaseInvoiceID", "الفاتورة", required=True, fk="PurchaseInvoices.PurchaseInvoiceID"),
            money("Amount", "المبلغ", rule=">0", rule_text="المبلغ يجب أن يكون أكبر من صفر"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            created_at(),
        ],
        pk=["AllocationID"],
        indexes=[ux("PaymentID", "PurchaseInvoiceID"), ix("PurchaseInvoiceID")],
    ),

    # -------------------------------------------------------------- Expenses
    Table(
        "ExpenseTypes", "أنواع المصروفات",
        "قائمة ثابتة لأنواع المصروفات لضمان دقة التقارير المجمّعة.",
        [
            auto("ExpenseTypeID", "رقم النوع"),
            text("ExpenseTypeName", 50, "نوع المصروف", required=True),
            is_active(),
        ],
        pk=["ExpenseTypeID"],
        indexes=[ux("ExpenseTypeName")],
        seed_columns=["ExpenseTypeID", "ExpenseTypeName"],
        seed_rows=[
            (1, "الإيجار"), (2, "الكهرباء"), (3, "المياه"), (4, "الإنترنت والاتصالات"),
            (5, "النقل"), (6, "الصيانة"), (7, "الرواتب"), (8, "المستلزمات"),
            (9, "مصروفات أخرى"),
        ],
    ),

    Table(
        "Expenses", "المصروفات",
        "مصروفات المحل التشغيلية؛ المبلغ بدون ضريبة والضريبة منفصلة (ضريبة مدخلات).",
        [
            auto("ExpenseID", "رقم داخلي"),
            text("ExpenseNumber", 20, "رقم المصروف", required=True),
            date_("ExpenseDate", "تاريخ المصروف"),
            long_("ExpenseTypeID", "نوع المصروف", required=True, fk="ExpenseTypes.ExpenseTypeID"),
            money("Amount", "المبلغ قبل الضريبة", rule=">0",
                  rule_text="المبلغ يجب أن يكون أكبر من صفر"),
            money("Tax", "ضريبة المدخلات", note="فقط إذا كانت لدى المحل فاتورة ضريبية بالمصروف"),
            money("TotalAmount", "الإجمالي"),
            long_("PaymentMethodID", "طريقة الدفع", fk="PaymentMethods.PaymentMethodID"),
            text("SupplierInvoiceRef", 30, "رقم فاتورة المصروف"),
            text("Description", 255, "الوصف"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            created_at(),
            long_("CashBoxID", "صُرف من صندوق", fk="CashBoxes.CashBoxID",
                  note="المصروف النقدي يخرج من هذا الصندوق؛ فارغ = لم يُدفع من صندوق"),
            long_("BankID", "البنك", fk="Banks.BankID", note="المبلغ المحوَّل بنكيًا يُقيَّد في حساب هذا البنك"),
            long_("CostCenterID", "مركز التكلفة", fk="CostCenters.CostCenterID"),
            long_("RecurringID", "من مصروف متكرر", fk="RecurringExpenses.RecurringID",
                  note="أنشأه البرنامج من المصروف المتكرر بتاريخ استحقاقه"),
            *fx_fields("المبلغ بالعملة (بدون الضريبة)"), money("ForeignTax", "الضريبة بالعملة"),
        ],
        pk=["ExpenseID"],
        indexes=[ux("ExpenseNumber"), ix("ExpenseDate")],
        rule="[TotalAmount]=[Amount]+[Tax]",
        rule_text="الإجمالي = المبلغ + الضريبة",
    ),

    Table(
        "RecurringExpenses", "المصروفات المتكررة",
        "مصروف ثابت يتكرر (الإيجار، الكهرباء، الاشتراكات): يُنشئ البرنامج المصروف في تاريخ استحقاقه، "
        "مرة واحدة لكل تاريخ.",
        [
            auto("RecurringID", "رقم داخلي"),
            text("RecurringName", 100, "اسم المصروف", required=True),
            long_("ExpenseTypeID", "نوع المصروف", required=True, fk="ExpenseTypes.ExpenseTypeID"),
            money("Amount", "المبلغ قبل الضريبة", rule=">0", rule_text="المبلغ يجب أن يكون أكبر من صفر"),
            money("Tax", "ضريبة المدخلات"),
            long_("PaymentMethodID", "طريقة الدفع", required=True, fk="PaymentMethods.PaymentMethodID"),
            long_("CashBoxID", "من صندوق", fk="CashBoxes.CashBoxID", note="للدفع النقدي؛ فارغ = صندوق من يُنشئ المصروف"),
            long_("BankID", "من بنك", fk="Banks.BankID", note="للتحويل البنكي؛ فارغ = البنك الافتراضي"),
            long_("CostCenterID", "مركز التكلفة", fk="CostCenters.CostCenterID"),
            text("Frequency", 10, "التكرار", required=True, default='"MONTHLY"',
                 rule='In ("MONTHLY","QUARTERLY","YEARLY")',
                 rule_text="MONTHLY = شهري، QUARTERLY = كل 3 أشهر، YEARLY = سنوي"),
            int_("DueDay", "يوم الاستحقاق", required=True, default="1", rule="Between 1 And 28",
                 rule_text="يوم الاستحقاق من 1 إلى 28"),
            date_("StartDate", "يبدأ من"),
            date_("EndDate", "ينتهي في", required=False, default=None),
            date_("NextDueDate", "الاستحقاق التالي", required=False, default=None),
            date_("LastCreatedDate", "آخر مصروف أُنشئ", required=False, default=None),
            text("Description", 255, "الوصف"),
            is_active(),
            created_at(),
        ],
        pk=["RecurringID"],
        indexes=[ux("RecurringName"), ix("NextDueDate")],
    ),

    # -------------------------------------------------------------- Treasury
    Table(
        "CashVouchers", "سندات النقدية",
        "قبض نقدية لصندوق، أو صرف منه، أو تحويل بين صندوقين (ومنها ترحيل يومية الكاشير).",
        [
            auto("CashVoucherID", "رقم داخلي"),
            text("VoucherNumber", 20, "رقم السند", required=True),
            datetime_("VoucherDate", "التاريخ", required=True, default="Now()"),
            text("VoucherType", 10, "نوع السند", required=True,
                 rule='In ("IN","OUT","TRANSFER")',
                 rule_text="IN = قبض، OUT = صرف، TRANSFER = تحويل بين صندوقين"),
            long_("CashBoxID", "الصندوق", required=True, fk="CashBoxes.CashBoxID",
                  note="القبض يدخله، والصرف والتحويل يخرجان منه"),
            long_("ToCashBoxID", "إلى صندوق", fk="CashBoxes.CashBoxID", note="للتحويل فقط"),
            text("Category", 10, "البند", required=True, default='"OTHER"',
                 rule='In ("OTHER","OWNER","EXPENSE","ADVANCE","SHORTAGE","OVERAGE","TRANSFER","COMMISSION")',
                 rule_text="اختر البند من القائمة"),
            money("Amount", "المبلغ", rule=">0", rule_text="المبلغ يجب أن يكون أكبر من صفر"),
            text("PartyName", 100, "المستلم / المسلِّم"),
            text("Description", 255, "البيان"),
            long_("ExpenseID", "المصروف المسجَّل", fk="Expenses.ExpenseID",
                  note="صرف بند مصروف يسجل مصروفًا بنفس المبلغ في المصروفات"),
            long_("ClosingID", "تصفية الكاشير", fk="CashClosings.ClosingID"),
            long_("AdvanceEmployeeID", "الموظف صاحب السلفة", fk="Employees.EmployeeID",
                  note="سلفة موظف (صرف) أو سدادها نقدًا (قبض)"),
            long_("CostCenterID", "مركز التكلفة", fk="CostCenters.CostCenterID"),
            long_("EmployeeID", "الموظف", required=True, fk="Employees.EmployeeID"),
            created_at(),
            long_("SalesRepID", "المندوب", fk="SalesReps.SalesRepID", note="صرف عمولة المندوب"),
        ],
        pk=["CashVoucherID"],
        indexes=[ux("VoucherNumber"), ix("VoucherDate"), ix("CashBoxID")],
        rule='[VoucherType]<>"TRANSFER" Or ([ToCashBoxID] Is Not Null And [ToCashBoxID]<>[CashBoxID])',
        rule_text="التحويل يحتاج صندوقًا آخر غير صندوق الصرف",
    ),

    Table(
        "CashClosings", "تصفية يومية الكاشير",
        "جرد نقدية صندوق الكاشير في نهاية الوردية وترحيلها للخزينة الرئيسية أو تسويتها مع المالك.",
        [
            auto("ClosingID", "رقم داخلي"),
            text("ClosingNumber", 20, "رقم التصفية", required=True),
            datetime_("ClosingDate", "تاريخ التصفية", required=True, default="Now()"),
            long_("CashBoxID", "الصندوق", required=True, fk="CashBoxes.CashBoxID"),
            long_("EmployeeID", "أجراها", required=True, fk="Employees.EmployeeID"),
            datetime_("PeriodStart", "من (آخر تصفية)"),
            money("OpeningBalance", "رصيد البداية", rule=None),
            money("CashIn", "المقبوضات"),
            money("CashOut", "المدفوعات"),
            money("ExpectedBalance", "الرصيد الدفتري", rule=None),
            money("CountedAmount", "النقدية الفعلية"),
            money("Difference", "الفرق", rule=None, note="سالب = عجز، موجب = زيادة"),
            text("Destination", 10, "الترحيل إلى", required=True, default='"MAIN"',
                 rule='In ("MAIN","OWNER","KEEP")',
                 rule_text="MAIN = الخزينة الرئيسية، OWNER = تسوية مع المالك، KEEP = يبقى في الصندوق"),
            long_("ToCashBoxID", "الخزينة المستلمة", fk="CashBoxes.CashBoxID"),
            money("TransferAmount", "المبلغ المرحَّل"),
            money("KeptAmount", "المتبقي في الصندوق (عهدة)"),
            text("Notes", 255, "ملاحظات"),
            created_at(),
        ],
        pk=["ClosingID"],
        indexes=[ux("ClosingNumber"), ix("CashBoxID", "ClosingDate")],
        rule="[TransferAmount]+[KeptAmount]=[CountedAmount]",
        rule_text="المرحَّل + المتبقي = النقدية الفعلية",
    ),

    # ------------------------------------------------------------ Accounting
    Table(
        "Accounts", "دليل الحسابات (شجرة الحسابات)",
        "شجرة من خمسة مستويات على الأكثر: الحسابات الرئيسية (تجميعية) والحسابات الفرعية التي تُرحَّل إليها القيود. "
        "حسابات الصناديق (110000 + رقم الصندوق) وأنواع المصروفات (530000 + رقم النوع) تُنشأ تلقائيًا.",
        [
            long_("AccountCode", "رقم الحساب", required=True, rule=">0", rule_text="رقم الحساب أكبر من صفر"),
            text("AccountName", 100, "اسم الحساب", required=True),
            text("AccountNameEn", 100, "الاسم بالإنجليزية", note="يظهر في الواجهة الإنجليزية والقوائم المالية بها"),
            text("AccountType", 10, "نوع الحساب", required=True,
                 rule='In ("ASSET","LIABILITY","EQUITY","REVENUE","EXPENSE")',
                 rule_text="أصول، خصوم، حقوق ملكية، إيرادات، مصروفات"),
            long_("ParentCode", "الحساب الرئيسي", note="فارغ للحسابات الخمسة في المستوى الأول فقط"),
            is_active(),
            bool_("IsPosting", "حساب فرعي (يقبل القيود)", "True"),
            bool_("IsSystem", "حساب أساسي في النظام", "False"),
            byte_("AccountLevel", "المستوى", default="1"),
            text("TreeKey", 60, "مفتاح الترتيب في الشجرة",
                 note="يحسبه البرنامج (modAccounts.RebuildAccountTree): رقم كل مستوى بعشر خانات"),
            long_("Level1Code", "حساب المستوى 1"),
            long_("Level2Code", "حساب المستوى 2"),
            long_("Level3Code", "حساب المستوى 3"),
            long_("Level4Code", "حساب المستوى 4"),
            long_("Level5Code", "حساب المستوى 5"),
        ],
        pk=["AccountCode"],
        indexes=[ix("ParentCode"), ix("TreeKey")],
        seed_columns=["AccountCode", "AccountName", "AccountType", "ParentCode", "IsPosting", "IsSystem"],
        seed_rows=[],   # filled below from ACCOUNT_TREE
        seed_missing=True,
    ),

    Table(
        "JournalSourceTypes", "أنواع مصادر القيود",
        "أنواع العمليات التي يُنشأ عنها قيد آلي، ومنها يُعرف أصل القيد.",
        [
            text("SourceType", 20, "نوع العملية", required=True),
            text("TypeName", 50, "الاسم", required=True),
            text("TypeNameEn", 50, "الاسم بالإنجليزية"),
            int_("SortOrder", "الترتيب", required=True, default="0"),
        ],
        pk=["SourceType"],
        seed_columns=["SourceType", "TypeName", "SortOrder"],
        seed_rows=[
            ("SALE", "فاتورة بيع", 1), ("SALES_RETURN", "مرتجع بيع", 2),
            ("PURCHASE", "فاتورة شراء", 3), ("PURCHASE_RETURN", "مرتجع شراء", 4),
            ("CUSTOMER_PAYMENT", "سند قبض من عميل", 5), ("SUPPLIER_PAYMENT", "سند صرف لمورد", 6),
            ("EXPENSE", "مصروف", 7), ("CASH_VOUCHER", "سند نقدية", 8),
            ("STOCK_MOVE", "حركة مخزون يدوية", 9), ("STOCK_COUNT", "تسوية جرد", 10),
            ("BOX_OPENING", "رصيد افتتاحي لصندوق", 11), ("CUSTOMER_OPENING", "رصيد افتتاحي لعميل", 12),
            ("SUPPLIER_OPENING", "رصيد افتتاحي لمورد", 13), ("MANUAL", "قيد يدوي", 14),
            ("YEAR_CLOSE", "قيد إقفال السنة", 15), ("VAT_RETURN", "تسوية إقرار ضريبة القيمة المضافة", 16),
            ("VAT_PAYMENT", "سداد ضريبة القيمة المضافة", 17),
            ("BANK_OPENING", "رصيد افتتاحي لبنك", 18), ("BANK_TX", "حركة بنكية", 19),
            ("CHEQUE", "شيك وارد أو صادر", 20), ("CHEQUE_STATUS", "تحصيل أو ارتداد شيك", 21),
            ("ASSET", "اقتناء أصل ثابت", 22), ("ASSET_DISPOSAL", "بيع أو استبعاد أصل ثابت", 23),
            ("DEPRECIATION", "قيد الإهلاك الشهري", 24),
            ("PAYROLL", "قيد مسير الرواتب", 25), ("PAYROLL_PAYMENT", "صرف الرواتب", 26),
            ("COMMISSION", "قيد عمولات المندوبين", 27),
        ],
        seed_missing=True,
    ),

    Table(
        "JournalEntries", "قيود اليومية",
        "قيد آلي لكل عملية، مربوط بأصلها (SourceType + SourceID). يُحدَّث إذا تغيرت العملية.",
        [
            auto("EntryID", "رقم داخلي"),
            text("EntryNumber", 20, "رقم القيد", required=True),
            datetime_("EntryDate", "تاريخ القيد", required=True),
            text("SourceType", 20, "نوع العملية", required=True),
            long_("SourceID", "رقم العملية الداخلي", required=True),
            text("SourceNumber", 20, "رقم مستند العملية"),
            text("Description", 255, "البيان"),
            money("TotalDebit", "إجمالي المدين"),
            money("TotalCredit", "إجمالي الدائن"),
            int_("LineCount", "عدد الأسطر", required=True, default="0"),
            money("Signature", "بصمة القيد", note="تكشف تغيّر العملية بعد إنشاء القيد"),
            datetime_("UpdatedAt", "آخر تحديث"),
            created_at(),
            *fx_fields("مبلغ المستند بعملته"),
        ],
        pk=["EntryID"],
        indexes=[ux("EntryNumber"), ux("SourceType", "SourceID"), ix("EntryDate")],
        rule="[TotalDebit]=[TotalCredit]",
        rule_text="القيد غير متوازن: المدين يجب أن يساوي الدائن",
    ),

    Table(
        "JournalLines", "أسطر القيود",
        "الطرف المدين والطرف الدائن لكل قيد.",
        [
            auto("JournalLineID", "رقم السطر الداخلي"),
            long_("EntryID", "القيد", required=True, fk="JournalEntries.EntryID", cascade=True),
            int_("LineNumber", "رقم السطر", required=True),
            long_("AccountCode", "الحساب", required=True, fk="Accounts.AccountCode"),
            money("Debit", "مدين"),
            money("Credit", "دائن"),
            text("LineText", 255, "البيان"),
            long_("CostCenterID", "مركز التكلفة", fk="CostCenters.CostCenterID", note="من مستند العملية؛ فارغ = غير موزع"),
        ],
        pk=["JournalLineID"],
        indexes=[ux("EntryID", "LineNumber"), ix("AccountCode")],
    ),

    Table(
        "PeriodClosings", "سجل إقفال الفترات",
        "كل إقفال أو إعادة فتح لفترة أو سنة مالية: التاريخ الجديد للإقفال والسابق، ومن قام به والسبب.",
        [
            auto("PeriodClosingID", "رقم داخلي"),
            text("ActionType", 12, "العملية", required=True,
                 rule='In ("CLOSE","REOPEN","YEAR_CLOSE","YEAR_OPEN")',
                 rule_text="إقفال، إعادة فتح، إقفال سنة، إعادة فتح سنة"),
            datetime_("ClosedThrough", "مقفلة حتى (بعد العملية)"),
            datetime_("PreviousThrough", "مقفلة حتى (قبل العملية)"),
            int_("FiscalYear", "السنة المالية"),
            text("Notes", 255, "السبب / ملاحظات"),
            long_("EmployeeID", "قام به", required=True, fk="Employees.EmployeeID"),
            created_at(),
        ],
        pk=["PeriodClosingID"],
    ),

    Table(
        "FiscalYearClosings", "إقفال السنوات المالية",
        "قيد إقفال كل سنة: أرصدة الإيرادات والمصروفات في 31 ديسمبر تُقفل في الأرباح المحتجزة (3300). "
        "أسطره محفوظة كما كانت يوم الإقفال، ويُرحَّل لليومية كعملية «قيد إقفال السنة».",
        [
            auto("YearClosingID", "رقم داخلي"),
            int_("FiscalYear", "السنة المالية", required=True),
            text("ClosingNumber", 20, "رقم الإقفال", required=True),
            date_("ClosingDate", "تاريخ الإقفال"),
            money("NetProfit", "صافي ربح (خسارة) السنة", rule=None, rule_text=None),
            long_("EmployeeID", "أقفلها", required=True, fk="Employees.EmployeeID"),
            text("Notes", 255, "ملاحظات"),
            created_at(),
        ],
        pk=["YearClosingID"],
        indexes=[ux("FiscalYear"), ux("ClosingNumber")],
    ),

    Table(
        "FiscalYearClosingLines", "أسطر قيود إقفال السنوات",
        "لكل حساب إيرادات أو مصروفات رصيده معكوسًا، ثم صافي الربح في الأرباح المحتجزة.",
        [
            auto("YearClosingLineID", "رقم السطر الداخلي"),
            long_("YearClosingID", "إقفال السنة", required=True, fk="FiscalYearClosings.YearClosingID", cascade=True),
            int_("LineNumber", "رقم السطر", required=True),
            long_("AccountCode", "الحساب", required=True, fk="Accounts.AccountCode"),
            money("Debit", "مدين"),
            money("Credit", "دائن"),
            text("LineText", 150, "البيان"),
        ],
        pk=["YearClosingLineID"],
        indexes=[ux("YearClosingID", "LineNumber")],
    ),

    Table(
        "VatReturns", "إقرارات ضريبة القيمة المضافة",
        "إقرار كل فترة ضريبية (شهر أو ربع سنة) بخانات نموذج هيئة الزكاة والضريبة والجمارك. "
        "المسودة تُحسب من المستندات، وعند الاعتماد تُحفظ قيمها كما هي ويُنشأ قيد التسوية "
        "(ضريبة المخرجات والمدخلات إلى حساب التسوية 2250)، ثم قيد السداد عند تسجيله.",
        [
            auto("VatReturnID", "رقم داخلي"),
            text("ReturnNumber", 20, "رقم الإقرار", required=True, note="VAT-yyyymmdd (آخر يوم في الفترة)"),
            date_("PeriodFrom", "بداية الفترة", default=None),
            date_("PeriodTo", "نهاية الفترة", default=None),
            text("Status", 10, "الحالة", required=True, default='"DRAFT"', rule='In ("DRAFT","FILED")',
                 rule_text="DRAFT = مسودة، FILED = معتمد"),
            money("SalesStdAmount", "1- المبيعات الخاضعة للنسبة الأساسية"),
            money("SalesStdAdjust", "1- تعديلات (مرتجعات) المبيعات الخاضعة", rule=None, rule_text=None),
            money("SalesStdVAT", "1- ضريبة المبيعات الخاضعة", rule=None, rule_text=None),
            money("SalesZeroAmount", "3- المبيعات بنسبة صفرية"),
            money("SalesZeroAdjust", "3- تعديلات المبيعات بنسبة صفرية", rule=None, rule_text=None),
            money("SalesExemptAmount", "5- المبيعات المعفاة"),
            money("SalesExemptAdjust", "5- تعديلات المبيعات المعفاة", rule=None, rule_text=None),
            money("PurchStdAmount", "7- المشتريات والمصروفات الخاضعة للنسبة الأساسية"),
            money("PurchStdAdjust", "7- تعديلات (مرتجعات) المشتريات الخاضعة", rule=None, rule_text=None),
            money("PurchStdVAT", "7- ضريبة المشتريات الخاضعة", rule=None, rule_text=None),
            money("PurchZeroAmount", "10- المشتريات بنسبة صفرية"),
            money("PurchZeroAdjust", "10- تعديلات المشتريات بنسبة صفرية", rule=None, rule_text=None),
            money("Corrections", "14- تصحيحات من الفترات السابقة", rule=None, rule_text=None,
                  note="موجبة تزيد الضريبة المستحقة، سالبة تنقصها"),
            money("CarriedCredit", "15- الرصيد الدائن المرحَّل من الفترات السابقة"),
            money("NetDue", "16- صافي الضريبة المستحقة (سالب = مستردة)", rule=None, rule_text=None,
                  note="SalesStdVAT − PurchStdVAT + Corrections − CarriedCredit"),
            date_("FiledDate", "تاريخ الاعتماد (تاريخ قيد التسوية)", required=False, default=None),
            text("FilingRef", 30, "رقم الإقرار لدى الهيئة"),
            date_("PaidDate", "تاريخ السداد", required=False, default=None),
            money("PaidAmount", "المبلغ المسدد"),
            long_("PaidAccount", "حساب السداد (البنك)", fk="Accounts.AccountCode"),
            text("Notes", 255, "ملاحظات"),
            long_("EmployeeID", "أعدّه / اعتمده", required=True, fk="Employees.EmployeeID"),
            created_at(),
        ],
        pk=["VatReturnID"],
        indexes=[ux("ReturnNumber"), ix("PeriodFrom")],
        rule="[PeriodTo]>=[PeriodFrom] And ([PaidAmount]=0 Or [PaidAmount]<=[NetDue])",
        rule_text="نهاية الفترة قبل بدايتها، أو المسدد أكبر من الضريبة المستحقة",
    ),

    Table(
        "ManualEntries", "القيود اليدوية",
        "قيد يكتبه المحاسب بنفسه (مستحقات، تسويات، رأس المال، أرصدة افتتاحية...). يُرحَّل لليومية "
        "كأي عملية أخرى (SourceType = MANUAL)، ويُعدَّل أو يُحذف فيتبعه قيده.",
        [
            auto("ManualEntryID", "رقم داخلي"),
            text("EntryNumber", 20, "رقم القيد اليدوي", required=True),
            date_("EntryDate", "تاريخ القيد"),
            text("Description", 255, "البيان", required=True),
            text("Reference", 50, "المرجع", note="رقم مستند خارجي: فاتورة، عقد، كشف بنك..."),
            long_("ReversalOfID", "عكس القيد", note="القيد اليدوي الذي يعكسه هذا القيد"),
            money("TotalAmount", "إجمالي القيد"),
            long_("EmployeeID", "أدخله", required=True, fk="Employees.EmployeeID"),
            created_at(),
            datetime_("UpdatedAt", "آخر تعديل"),
            *fx_fields("إجمالي القيد بالعملة"),
        ],
        pk=["ManualEntryID"],
        indexes=[ux("EntryNumber"), ix("EntryDate")],
    ),

    Table(
        "ManualEntryLines", "أسطر القيود اليدوية",
        "الطرف المدين والطرف الدائن للقيد اليدوي؛ كل سطر مدين أو دائن فقط.",
        [
            auto("ManualLineID", "رقم السطر الداخلي"),
            long_("ManualEntryID", "القيد اليدوي", required=True, fk="ManualEntries.ManualEntryID", cascade=True),
            int_("LineNumber", "رقم السطر", required=True),
            long_("AccountCode", "الحساب", required=True, fk="Accounts.AccountCode"),
            money("Debit", "مدين"),
            money("Credit", "دائن"),
            text("LineText", 150, "بيان السطر"),
            long_("CostCenterID", "مركز التكلفة", fk="CostCenters.CostCenterID"),
            money("ForeignDebit", "مدين بالعملة"), money("ForeignCredit", "دائن بالعملة"),
        ],
        pk=["ManualLineID"],
        indexes=[ux("ManualEntryID", "LineNumber"), ix("AccountCode")],
        rule="([Debit]=0 Or [Credit]=0) And [Debit]+[Credit]>0",
        rule_text="كل سطر مدين أو دائن فقط، وبمبلغ أكبر من صفر",
    ),

    # ------------------------------------------------------------- Inventory
    Table(
        "TransactionTypes", "أنواع حركات المخزون",
        "أنواع الحركة وإشارتها: +1 تزيد المخزون، −1 تنقصه، 0 تسوية بالإشارة.",
        [
            long_("TransactionTypeID", "رقم النوع", required=True),
            text("TypeCode", 20, "رمز النوع", required=True),
            text("TypeName", 50, "نوع الحركة", required=True),
            text("TypeNameEn", 50, "الاسم بالإنجليزية"),
            int_("Direction", "الاتجاه", required=True, rule="In (-1,0,1)",
                 rule_text="الاتجاه 1 أو -1 أو 0"),
            bool_("IsManual", "متاح للإدخال اليدوي", "False"),
        ],
        pk=["TransactionTypeID"],
        indexes=[ux("TypeCode")],
        seed_columns=["TransactionTypeID", "TypeCode", "TypeName", "Direction", "IsManual"],
        seed_rows=[
            (1, "PURCHASE", "شراء", 1, False),
            (2, "SALE", "بيع", -1, False),
            (3, "PURCHASE_RETURN", "مرتجع شراء", -1, False),
            (4, "SALES_RETURN", "مرتجع بيع", 1, False),
            (5, "STOCK_IN", "إضافة مخزون", 1, True),
            (6, "STOCK_OUT", "خصم مخزون", -1, True),
            (7, "ADJUSTMENT", "تسوية جرد", 0, False),
            (8, "OPENING", "رصيد افتتاحي", 1, True),
        ],
    ),

    Table(
        "InventoryTransactions", "حركة المخزون",
        "دفتر أستاذ المخزون: الكمية مخزنة بإشارتها، ورصيد أي صنف = مجموع Quantity.",
        [
            auto("TransactionID", "رقم الحركة"),
            datetime_("TransactionDate", "تاريخ الحركة", required=True, default="Now()"),
            long_("ProductID", "المنتج", required=True, fk="Products.ProductID"),
            long_("TransactionTypeID", "نوع الحركة", required=True,
                  fk="TransactionTypes.TransactionTypeID"),
            qty("Quantity", "الكمية (+/−)", rule="<>0", rule_text="الكمية لا يمكن أن تكون صفرًا",
                default=None),
            money("UnitCost", "تكلفة الوحدة"),
            qty("QuantityAfter", "الرصيد بعد الحركة"),
            text("ReferenceType", 20, "نوع المستند",
                 note="SALE, SALES_RETURN, PURCHASE, PURCHASE_RETURN, STOCK_COUNT, MANUAL"),
            long_("ReferenceID", "رقم المستند الداخلي"),
            text("ReferenceNumber", 20, "رقم المستند"),
            long_("EmployeeID", "الموظف", fk="Employees.EmployeeID"),
            text("Notes", 255, "ملاحظات"),
            created_at(),
        ],
        pk=["TransactionID"],
        indexes=[ix("TransactionDate"), ix("ReferenceType", "ReferenceID")],
    ),

    Table(
        "StockCounts", "جلسات الجرد",
        "رأس عملية الجرد؛ تبقى مفتوحة حتى الترحيل الذي ينشئ حركات التسوية.",
        [
            auto("StockCountID", "رقم داخلي"),
            text("CountNumber", 20, "رقم الجرد", required=True),
            datetime_("CountDate", "تاريخ الجرد", required=True, default="Now()"),
            long_("CategoryID", "تصنيف محدد (اختياري)", fk="Categories.CategoryID"),
            text("Status", 10, "الحالة", required=True, default='"OPEN"',
                 rule='In ("OPEN","POSTED","CANCELLED")',
                 rule_text="OPEN مفتوح، POSTED مُرحّل، CANCELLED ملغى"),
            long_("EmployeeID", "أجراه", required=True, fk="Employees.EmployeeID"),
            datetime_("PostedAt", "تاريخ الترحيل"),
            long_("PostedByID", "رحّله", fk="Employees.EmployeeID"),
            text("Notes", 255, "ملاحظات"),
        ],
        pk=["StockCountID"],
        indexes=[ux("CountNumber")],
    ),

    Table(
        "StockCountDetails", "تفاصيل الجرد",
        "الكمية المسجلة والفعلية والفرق لكل صنف في جلسة الجرد.",
        [
            auto("StockCountDetailID", "رقم السطر الداخلي"),
            long_("StockCountID", "الجرد", required=True,
                  fk="StockCounts.StockCountID", cascade=True),
            long_("ProductID", "المنتج", required=True, fk="Products.ProductID"),
            qty("SystemQuantity", "الكمية المسجلة"),
            qty("ActualQuantity", "الكمية الفعلية", required=False, default=None,
                rule="Is Null Or >=0", rule_text="الكمية الفعلية لا يمكن أن تكون سالبة"),
            qty("Difference", "الفرق", note="ActualQuantity − SystemQuantity"),
            money("UnitCost", "تكلفة الوحدة"),
            money("DifferenceValue", "قيمة الفرق", rule=None, note="سالب = عجز، موجب = زيادة"),
            text("Notes", 255, "ملاحظات"),
        ],
        pk=["StockCountDetailID"],
        indexes=[ux("StockCountID", "ProductID")],
    ),

    # ------------------------------------------------------------- Audit
    Table(
        "AuditLog", "سجل العمليات",
        "سجل التدقيق: الدخول والخروج، وإضافة أي سجل أو تعديله أو حذفه (والحقول في AuditChanges)، "
        "وعمليات المستندات والعمليات الحساسة (تجاوز المخزون، تعديل الأسعار، النسخ الاحتياطي).",
        [
            auto("LogID", "رقم السجل"),
            datetime_("LogDate", "التاريخ", required=True, default="Now()"),
            long_("EmployeeID", "الموظف", fk="Employees.EmployeeID"),
            text("ActionType", 30, "نوع العملية", required=True),
            text("ObjectName", 50, "الكائن"),
            text("RecordID", 30, "رقم السجل المتأثر"),
            memo("Details", "التفاصيل"),
            text("ComputerName", 50, "اسم الجهاز"),
            text("RecordLabel", 100, "السجل", note="اسم السجل أو رقمه المقروء (المنتج، رقم المصروف...)"),
        ],
        pk=["LogID"],
        indexes=[ix("LogDate"), ix("ActionType")],
    ),

    Table(
        "AuditChanges", "تفاصيل سجل التدقيق",
        "حقول كل عملية في سجل التدقيق: القيمة قبل وبعد (الإضافة: بعد فقط، الحذف: قبل فقط).",
        [
            auto("ChangeID", "رقم داخلي"),
            long_("LogID", "العملية", required=True, fk="AuditLog.LogID", cascade=True),
            int_("LineNo", "الترتيب", required=True, default="0"),
            text("FieldName", 64, "الحقل", required=True),
            text("FieldCaption", 100, "اسم الحقل"),
            text("OldValue", 255, "القيمة قبل"),
            text("NewValue", 255, "القيمة بعد"),
        ],
        pk=["ChangeID"],
        indexes=[ix("LogID")],
    ),

    # ------------------------------------------------------------- Barcode labels
    Table(
        "LabelSettings", "إعدادات ملصقات الباركود",
        "سجل واحد: مقاس الملصق والورق والهوامش، وحجم الباركود، والنصوص أعلاه وأسفله.",
        [
            long_("LabelSettingID", "رقم الإعداد", required=True, default="1", rule="=1",
                  rule_text="يسمح بسجل إعدادات واحد فقط"),
            text("PrinterName", 255, "طابعة الملصقات", note="فارغ = الطابعة الافتراضية"),
            qty("LabelWidth", "عرض الملصق (مم)", default="38", rule="Between 15 And 210",
                rule_text="عرض الملصق من 15 إلى 210 مم"),
            qty("LabelHeight", "ارتفاع الملصق (مم)", default="25", rule="Between 10 And 297",
                rule_text="ارتفاع الملصق من 10 إلى 297 مم"),
            byte_("LabelsAcross", "عدد الملصقات في الصف", default="1", rule="Between 1 And 10",
                  rule_text="من 1 إلى 10 ملصقات في الصف"),
            qty("ColumnGap", "المسافة بين الأعمدة (مم)", default="2", rule="Between 0 And 50",
                rule_text="من 0 إلى 50 مم"),
            qty("RowGap", "المسافة بين الصفوف (مم)", default="0", rule="Between 0 And 50",
                rule_text="من 0 إلى 50 مم"),
            qty("MarginTop", "الهامش العلوي (مم)", default="0", rule="Between 0 And 50",
                rule_text="من 0 إلى 50 مم"),
            qty("MarginBottom", "الهامش السفلي (مم)", default="0", rule="Between 0 And 50",
                rule_text="من 0 إلى 50 مم"),
            qty("MarginLeft", "الهامش الأيسر (مم)", default="0", rule="Between 0 And 50",
                rule_text="من 0 إلى 50 مم"),
            qty("MarginRight", "الهامش الأيمن (مم)", default="0", rule="Between 0 And 50",
                rule_text="من 0 إلى 50 مم"),
            qty("BarHeight", "ارتفاع الباركود (مم)", default="10", rule="Between 3 And 100",
                rule_text="ارتفاع الباركود من 3 إلى 100 مم"),
            qty("BarWidth", "عرض أرفع خط (مم)", default="0.25", rule="Between 0.1 And 1",
                rule_text="عرض أرفع خط من 0.1 إلى 1 مم (المعتاد 0.25 - 0.33)"),
            text("TopLine1", 10, "السطر الأول أعلى الباركود", required=True, default='"STORE"',
                 rule="In (\"NONE\",\"STORE\",\"NAME\",\"PRICE\",\"CODE\",\"BARCODE\")", rule_text="اختر من القائمة"),
            text("TopLine2", 10, "السطر الثاني أعلى الباركود", required=True, default='"NAME"',
                 rule="In (\"NONE\",\"STORE\",\"NAME\",\"PRICE\",\"CODE\",\"BARCODE\")", rule_text="اختر من القائمة"),
            text("BottomLine1", 10, "السطر الأول أسفل الباركود", required=True, default='"BARCODE"',
                 rule="In (\"NONE\",\"STORE\",\"NAME\",\"PRICE\",\"CODE\",\"BARCODE\")", rule_text="اختر من القائمة"),
            text("BottomLine2", 10, "السطر الثاني أسفل الباركود", required=True, default='"PRICE"',
                 rule="In (\"NONE\",\"STORE\",\"NAME\",\"PRICE\",\"CODE\",\"BARCODE\")", rule_text="اختر من القائمة"),
            text("ShortName", 30, "الاسم المختصر للمحل", note="يُطبع إذا اخترت «الاسم المختصر»"),
            byte_("FontSize", "حجم الخط", default="7", rule="Between 5 And 16",
                  rule_text="حجم الخط من 5 إلى 16"),
        ],
        pk=["LabelSettingID"],
        seed_columns=["LabelSettingID"],
        seed_rows=[(1,)],
    ),
]


# --------------------------------------------------------------------------
# Role -> permission mapping (Administrator gets everything)
# --------------------------------------------------------------------------
def _role_permissions():
    perms = [r[0] for r in table("Permissions").seed_rows]
    cashier = ["SALES_POS", "SALES_VIEW", "CUSTOMERS", "CUSTOMER_PAYMENTS", "CASH_CLOSING"]
    manager_excluded = {"SETTINGS", "USERS", "BACKUP", "ALLOW_NEGATIVE_STOCK", "PERIOD_CLOSE", "AUDIT_LOG"}
    rows = [(1, p) for p in perms]
    rows += [(2, p) for p in perms if p not in manager_excluded]
    rows += [(3, p) for p in cashier]
    return rows


# --------------------------------------------------------------------------
# Screens of the permissions screen: (form, title, section, role permission, add, edit, delete)
# add = a new record or saving a new document; documents are never edited or deleted after saving.
# --------------------------------------------------------------------------
SCREEN_LIST = [
    ("frmPOS", "نقطة البيع (المحلات)", "المبيعات", "SALES_POS", True, False, False),
    ("frmTouchPOS", "نقطة البيع (المطاعم)", "المبيعات", "SALES_POS", True, False, False),
    ("frmCafePOS", "نقطة البيع (الكافيهات)", "المبيعات", "SALES_POS", True, False, False),
    ("frmSalesInvoice", "عرض الفواتير وإعادة طباعتها", "المبيعات", "SALES_VIEW", False, False, False),
    ("frmSalesReturn", "مرتجعات المبيعات", "المبيعات", "SALES_RETURN", True, False, False),
    ("frmCustomers", "العملاء", "العملاء", "CUSTOMERS", True, True, True),
    ("frmCustomerPayment", "سندات القبض من العملاء", "العملاء", "CUSTOMER_PAYMENTS", True, False, False),
    ("frmPurchaseInvoice", "فواتير المشتريات", "المشتريات", "PURCHASES", True, False, False),
    ("frmPurchaseView", "عرض فواتير المشتريات", "المشتريات", "PURCHASES", False, False, False),
    ("frmPurchaseReturn", "مرتجعات المشتريات", "المشتريات", "PURCHASE_RETURN", True, False, False),
    ("frmSuppliers", "الموردون", "الموردون", "SUPPLIERS", True, True, True),
    ("frmSupplierPayment", "سندات الصرف للموردين", "الموردون", "SUPPLIER_PAYMENTS", True, False, False),
    ("frmProducts", "المنتجات والأسعار", "المخزون", "PRODUCTS", True, True, True),
    ("frmCategories", "التصنيفات", "المخزون", "PRODUCTS", True, True, True),
    ("frmUnits", "الوحدات", "المخزون", "PRODUCTS", True, True, True),
    ("frmInventory", "المخزون والحركات اليدوية", "المخزون", "PRODUCTS", True, False, False),
    ("frmStockCount", "الجرد", "المخزون", "STOCK_COUNT", True, False, False),
    ("frmBarcodeLabels", "ملصقات الباركود", "المخزون", "PRODUCTS", False, False, False),
    ("frmLabelSettings", "إعدادات الملصقات", "المخزون", "PRODUCTS", False, True, False),
    ("frmExpenses", "المصروفات", "المصروفات", "EXPENSES", True, True, True),
    ("frmExpenseTypes", "أنواع المصروفات", "المصروفات", "EXPENSES", True, True, True),
    ("frmRecurring", "المصروفات المتكررة", "المصروفات", "EXPENSES", True, True, True),
    ("frmTreasury", "الخزينة", "الخزينة", "CASH_CLOSING", False, False, False),
    ("frmCashVoucher", "سندات النقدية والتحويل", "الخزينة", "CASH_BOX", True, False, False),
    ("frmCashClosing", "تصفية يومية الكاشير", "الخزينة", "CASH_CLOSING", True, False, False),
    ("frmCashBoxes", "الصناديق", "الخزينة", "CASH_BOX", True, True, True),
    ("frmAccounting", "المحاسبة والمالية", "الحسابات", None, False, False, False),
    ("frmJournal", "قيود اليومية", "الحسابات", "JOURNAL", False, False, False),
    ("frmAccounts", "دليل الحسابات", "الحسابات", "JOURNAL", True, True, True),
    ("frmManualEntry", "القيود اليدوية", "الحسابات", "MANUAL_ENTRY", True, True, True),
    ("frmLedger", "كشف حساب ودفتر الأستاذ", "الحسابات", "JOURNAL", False, False, False),
    ("frmFinancials", "القوائم المالية", "الحسابات", "REPORTS_PROFIT", False, False, False),
    ("frmPeriodClosing", "إقفال الفترات والسنة المالية", "الحسابات", "PERIOD_CLOSE", False, False, False),
    ("frmVatReturn", "إقرار ضريبة القيمة المضافة", "الحسابات", "VAT_RETURN", True, True, True),
    ("frmAging", "أعمار الديون (العملاء والموردون)", "التقارير", "REPORTS", False, False, False),
    ("frmBanks", "البنوك", "الخزينة", "BANKS", True, True, True),
    ("frmBankTx", "الحركات البنكية", "الخزينة", "BANKS", True, False, True),
    ("frmBankRecon", "التسوية البنكية", "الخزينة", "BANKS", True, True, False),
    ("frmCheques", "الشيكات الواردة والصادرة", "الخزينة", "CHEQUES", True, True, True),
    ("frmAssets", "الأصول الثابتة", "الحسابات", "FIXED_ASSETS", True, True, True),
    ("frmDepreciation", "الإهلاك الشهري", "الحسابات", "FIXED_ASSETS", True, False, True),
    ("frmPayroll", "مسير الرواتب", "الحسابات", "PAYROLL", True, True, True),
    ("frmCostCenters", "مراكز التكلفة والفروع", "الحسابات", "JOURNAL", True, True, True),
    ("frmBudget", "الموازنة التقديرية", "الحسابات", "BUDGET", True, True, True),
    ("frmCurrencies", "العملات", "الحسابات", "CURRENCIES", True, True, True),
    ("frmSalesReps", "المندوبين", "المبيعات", "SALES_REPS", True, True, True),
    ("frmRepTargets", "أهداف المندوبين", "المبيعات", "SALES_REPS", True, True, True),
    ("frmCommissions", "عمولات المندوبين", "المبيعات", "SALES_REPS", True, True, True),
    ("frmCurrencyRates", "أسعار العملات", "الحسابات", "CURRENCIES", True, True, True),
    ("frmAllocation", "ربط السداد بالفواتير", "العملاء", "CUSTOMER_PAYMENTS", True, False, True),
    ("frmReportCenter", "التقارير", "التقارير", "REPORTS", False, False, False),
    ("frmSearch", "البحث", "النظام", None, False, False, False),
    ("frmSettings", "إعدادات المحل", "النظام", "SETTINGS", False, True, False),
    ("frmUsers", "المستخدمون", "النظام", "USERS", True, True, False),
    ("frmRoles", "الأدوار والصلاحيات", "النظام", "USERS", False, True, False),
    ("frmUserScreens", "صلاحيات الشاشات للمستخدمين", "النظام", "USERS", False, True, False),
    ("frmAuditLog", "سجل التدقيق", "النظام", "AUDIT_LOG", False, False, False),
    ("frmBackup", "النسخ الاحتياطي", "النظام", "BACKUP", False, False, False),
]


# --------------------------------------------------------------------------
# Chart of accounts: (code, name, type, parent, posting, system)
#   posting = False: a main (summary) account - entries go to its sub-accounts;
#   system = True: used by the automatic entries (modJournal) or the tree itself - not deleted.
# Codes of the accounts used by the automatic entries never change.
# --------------------------------------------------------------------------
A, L, E, R, X = "ASSET", "LIABILITY", "EQUITY", "REVENUE", "EXPENSE"
ACCOUNT_TREE = [
    (1, "الأصول", A, None, False, True),
    (11, "الأصول المتداولة", A, 1, False, True),
    (1100, "النقدية بالخزينة والصناديق", A, 11, False, True),
    (110001, "الخزينة الرئيسية", A, 1100, True, True),
    (110002, "صندوق الكاشير", A, 1100, True, True),
    (1190, "نقدية غير موزعة على صندوق", A, 11, True, True),
    (1200, "مدى والمحافظ تحت التسوية (البنك والشبكة)", A, 11, True, True),
    (1210, "الحسابات البنكية", A, 11, False, True),
    (1250, "شيكات تحت التحصيل", A, 11, True, True),
    (1300, "ذمم العملاء", A, 11, True, True),
    (1310, "أوراق القبض", A, 11, True, False),
    (1350, "مخصص الديون المشكوك في تحصيلها", A, 11, True, False),
    (1400, "المخزون", A, 11, True, True),
    (1500, "ضريبة القيمة المضافة - مدخلات", A, 11, True, True),
    (1600, "سلف الموظفين", A, 11, True, True),
    (1610, "العهد النقدية", A, 11, True, False),
    (1650, "مصروفات مدفوعة مقدمًا", A, 11, True, False),
    (1660, "إيرادات مستحقة", A, 11, True, False),
    (1690, "تأمينات لدى الغير", A, 11, True, False),
    (12, "الأصول غير المتداولة", A, 1, False, True),
    (1710, "الأثاث والتجهيزات", A, 12, True, False),
    (1720, "الأجهزة والحاسبات وأنظمة نقاط البيع", A, 12, True, False),
    (1730, "السيارات ووسائل النقل", A, 12, True, False),
    (1740, "الديكورات وتحسينات المحل", A, 12, True, False),
    (1790, "مجمع إهلاك الأصول الثابتة", A, 12, True, True),
    (1800, "البرامج والتراخيص", A, 12, True, False),
    (2, "الخصوم", L, None, False, True),
    (21, "الخصوم المتداولة", L, 2, False, True),
    (2100, "ذمم الموردين", L, 21, True, True),
    (2110, "أوراق الدفع (شيكات صادرة)", L, 21, True, True),
    (2200, "ضريبة القيمة المضافة - مخرجات", L, 21, True, True),
    (2250, "ضريبة القيمة المضافة - التسوية والسداد", L, 21, True, True),
    (2300, "مصروفات مستحقة", L, 21, True, False),
    (2310, "رواتب وأجور مستحقة", L, 21, True, True),
    (2320, "التأمينات الاجتماعية المستحقة", L, 21, True, True),
    (2330, "عمولات مستحقة للمندوبين", L, 21, True, True),
    (2400, "دفعات مقدمة من العملاء", L, 21, True, False),
    (2500, "قروض قصيرة الأجل", L, 21, True, False),
    (2600, "الزكاة المستحقة", L, 21, True, False),
    (22, "الخصوم غير المتداولة", L, 2, False, True),
    (2700, "قروض طويلة الأجل", L, 22, True, False),
    (2800, "مخصص مكافأة نهاية الخدمة", L, 22, True, False),
    (3, "حقوق الملكية", E, None, False, True),
    (31, "رأس المال وجاري المالك", E, 3, False, True),
    (3200, "رأس المال", E, 31, True, False),
    (3100, "جاري المالك", E, 31, True, True),
    (3900, "أرصدة افتتاحية", E, 31, True, True),
    (32, "الأرباح المحتجزة ونتائج الأعمال", E, 3, False, True),
    (3300, "الأرباح المحتجزة", E, 32, True, True),
    (3400, "صافي ربح (خسارة) العام", E, 32, True, False),
    (4, "الإيرادات", R, None, False, True),
    (41, "إيرادات النشاط", R, 4, False, True),
    (4100, "المبيعات", R, 41, True, True),
    (4110, "مردودات المبيعات", R, 41, True, True),
    (4120, "الخصم المسموح به", R, 41, True, False),
    (42, "إيرادات أخرى", R, 4, False, True),
    (4200, "إيرادات متنوعة", R, 42, True, True),
    (4300, "زيادة الصناديق", R, 42, True, True),
    (4400, "الخصم المكتسب", R, 42, True, False),
    (4500, "أرباح بيع الأصول الثابتة", R, 42, True, True),
    (5, "المصروفات", X, None, False, True),
    (51, "تكلفة المبيعات", X, 5, False, True),
    (5100, "تكلفة البضاعة المباعة", X, 51, True, True),
    (5200, "فروقات وتسويات المخزون", X, 51, True, True),
    (52, "المصروفات التشغيلية والإدارية", X, 5, False, True),
    (5300, "المصروفات حسب النوع", X, 52, False, True),
    (5500, "الرواتب والأجور", X, 52, True, True),
    (5510, "البدلات والحوافز", X, 52, True, True),
    (5520, "التأمينات الاجتماعية", X, 52, True, True),
    (5530, "عمولات المندوبين", X, 52, True, True),
    (5600, "إهلاك الأصول الثابتة", X, 52, True, True),
    (5610, "عمولات البنوك ونقاط البيع", X, 52, True, True),
    (5620, "الرسوم الحكومية والتراخيص", X, 52, True, False),
    (5630, "الدعاية والإعلان", X, 52, True, False),
    (53, "مصروفات أخرى", X, 5, False, True),
    (5400, "عجز الصناديق", X, 53, True, True),
    (5650, "خسائر بيع واستبعاد الأصول الثابتة", X, 53, True, True),
    (5700, "الديون المعدومة", X, 53, True, False),
    (5800, "الزكاة", X, 53, True, False),
    (5900, "مصروفات متنوعة", X, 53, True, True),
]


def table(name: str) -> Table:
    for t in TABLES:
        if t.name == name:
            return t
    raise KeyError(name)


table("RolePermissions").seed_rows = _role_permissions()
table("Accounts").seed_rows = list(ACCOUNT_TREE)
table("Screens").seed_rows = [(name, title, module, 10 * (i + 1), key, add, edit, delete)
                              for i, (name, title, module, key, add, edit, delete) in enumerate(SCREEN_LIST)]
# ScreenName in UserScreens points to Screens (text key)
table("UserScreens").fields[1].fk = "Screens.ScreenName"
# PermissionKey in RolePermissions points to Permissions (text key)
table("RolePermissions").fields[1].fk = "Permissions.PermissionKey"
table("RolePermissions").fields[1].on_delete_cascade = True
# the kind of operation of an entry points to JournalSourceTypes (text key)
table("JournalEntries").fields[3].fk = "JournalSourceTypes.SourceType"
