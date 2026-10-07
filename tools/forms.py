"""Phase 5: screen definitions and layout.

Every screen is described here; layout() turns it into a flat list of
controls with positions in twips (567 twips = 1 cm). Positions are logical:
x is measured from the START edge (the right edge, the form is right-to-left).
gen_forms.py turns the result into VBA that builds the forms in Access.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from schema import table

CM = 567


def cm(x: float) -> int:
    return int(round(x * CM))


# --------------------------------------------------------------------------
# Icons: Segoe MDL2 Assets code points (Windows 10/11)
# --------------------------------------------------------------------------
ICONS = {
    "sales": 0xE7BF, "purchases": 0xE896, "inventory": 0xE7B8, "products": 0xE8EC,
    "customers": 0xE716, "suppliers": 0xE77B, "expenses": 0xE8C7, "stocktake": 0xE8EF,
    "reports": 0xE8A5, "search": 0xE721, "settings": 0xE713, "users": 0xE8D7,
    "backup": 0xE8B7, "logout": 0xE7E8, "home": 0xE80F, "category": 0xE8FD,
    "treasury": 0xE825, "journal": 0xE8F1,
}


@dataclass
class Control:
    kind: str            # rect, label, icon, text, combo, check, list, button
    name: str
    x: int
    y: int
    w: int
    h: int
    props: Dict[str, object] = field(default_factory=dict)
    events: List[str] = field(default_factory=list)     # e.g. ["Click"]
    parent: str = ""     # attached label -> owner control
    source: str = ""     # bound field
    decorative: bool = False   # may overlap other controls (backgrounds, icons)


@dataclass
class Sym:
    """A VBA symbol (constant or expression) emitted as-is."""
    code: str


# --------------------------------------------------------------------------
# Screen definitions
# --------------------------------------------------------------------------
@dataclass
class Fld:
    field: str
    label: Optional[str] = None
    span: int = 1
    locked: bool = False
    rows: Optional[str] = None          # combo row source (SQL) or value list "a;b;c;d"
    widths: Optional[str] = None        # combo column widths in cm, e.g. "0;5"
    hook: bool = False                  # AfterUpdate -> FieldChanged
    button: Optional[Tuple[str, str, str]] = None   # (name, caption, handler call)
    hint: str = ""


@dataclass
class Info:
    """A full-width information label inside the field grid."""
    name: str
    text: str = " "


@dataclass
class DataScreen:
    name: str
    table: str
    caption: str
    subtitle: str
    icon: str
    fields: List[object]
    kind: str = "LIST"                   # LIST (list + detail) or SINGLE (one record)
    list_select: str = ""                # columns after the key, "t." alias
    list_from: str = ""
    list_order: str = ""
    list_headers: List[Tuple[str, float]] = field(default_factory=list)   # (header, width cm)
    search: List[str] = field(default_factory=list)
    active: str = ""                     # e.g. "t.IsActive"
    seq: str = ""                        # "SEQUENCE:Field"
    unique: List[str] = field(default_factory=list)
    extra_buttons: List[Tuple[str, str, str]] = field(default_factory=list)
    allow_add: bool = True
    allow_delete: bool = True
    record_source: str = ""

    @property
    def pk(self):
        return table(self.table).pk[0]

    def list_template(self) -> str:
        return (f"SELECT t.{self.pk}, {self.list_select} FROM {self.list_from} "
                f"WHERE ({{ACTIVE}}) AND ({{SEARCH}}) ORDER BY {self.list_order}")

    def tag(self) -> str:
        parts = [f"KIND={self.kind}", f"TABLE={self.table}", f"PK={self.pk}"]
        if self.kind == "LIST":
            parts += [f"LIST={self.list_template()}", "SEARCH=" + ",".join(self.search)]
        if self.active:
            parts.append(f"ACTIVE={self.active}")
        if self.seq:
            parts.append(f"SEQ={self.seq}")
        if self.unique:
            parts.append("UNIQUE=" + ",".join(self.unique))
        return "|".join(parts)


CATEGORY_ROWS = "SELECT CategoryID, CategoryName FROM Categories ORDER BY CategoryName"
UNIT_ROWS = "SELECT UnitID, UnitName FROM Units ORDER BY UnitName"
SUPPLIER_ROWS = "SELECT SupplierID, SupplierName FROM Suppliers ORDER BY SupplierName"
CUSTOMER_ROWS = "SELECT CustomerID, CustomerName FROM Customers ORDER BY CustomerName"
PRODUCT_ROWS = ("SELECT ProductID, ProductName & ' (' & ProductCode & ')' AS Item "
                "FROM Products ORDER BY ProductName")
EXPENSE_TYPE_ROWS = "SELECT ExpenseTypeID, ExpenseTypeName FROM ExpenseTypes ORDER BY ExpenseTypeName"
PAYMENT_ROWS = "SELECT PaymentMethodID, MethodName FROM PaymentMethods ORDER BY SortOrder"
CASHBOX_ROWS = "SELECT CashBoxID, BoxName FROM CashBoxes ORDER BY BoxType DESC, BoxName"
BOX_TYPES = "MAIN;خزينة رئيسية;CASHIER;صندوق كاشير"
ACCOUNT_TYPES = "ASSET;أصول;LIABILITY;خصوم;EQUITY;حقوق ملكية;REVENUE;إيرادات;EXPENSE;مصروفات"
# main (summary) accounts only: a sub-account always hangs under a main account
ACCOUNT_ROWS = ("SELECT AccountCode, Space((AccountLevel - 1) * 3) & AccountName AS Account FROM Accounts "
                "WHERE IsPosting = False ORDER BY TreeKey")
ROLE_ROWS = "SELECT RoleID, RoleName FROM Roles ORDER BY RoleID"
VAT_CATEGORY_LIST = "S;خاضع للضريبة 15%;Z;نسبة صفرية;E;معفى من الضريبة"


TILE_COLORS = ("BLUE;أزرق;GREEN;أخضر;ORANGE;برتقالي;PURPLE;بنفسجي;RED;أحمر;INDIGO;نيلي;TEAL;فيروزي;"
               "PINK;وردي;BROWN;بني;GREY;رمادي")
POS_MODES = "RETAIL;المحلات (باركود);RESTAURANT;المطاعم (شاشة لمس);CAFE;الكافيهات (شاشة لمس)"

INVOICE_PRINT_MODES = ("DIRECT;طباعة مباشرة بدون معاينة;PREVIEW;عرض معاينة الطباعة;"
                       "NONE;بدون طباعة")

LABEL_LINES = ("NONE;بدون;STORE;الاسم المختصر للمحل;NAME;اسم المنتج;PRICE;السعر;CODE;كود المنتج;"
               "BARCODE;رقم الباركود")

DATA_SCREENS: List[DataScreen] = [
    DataScreen(
        "frmProducts", "Products", "المنتجات", "إضافة وتعديل الأصناف والأسعار", "products",
        list_select="t.ProductCode AS [الكود], t.ProductName AS [المنتج], t.CurrentQuantity AS [الكمية]",
        list_from="Products AS t", list_order="t.ProductName",
        list_headers=[("الكود", 2.0), ("المنتج", 4.9), ("الكمية", 1.5)],
        search=["t.ProductName", "t.ProductCode", "t.Barcode", "t.ProductNameEn"],
        active="t.IsActive", seq="PRODUCT_CODE:ProductCode", unique=["ProductCode", "Barcode"],
        fields=[
            Fld("ProductCode", hint="يُولَّد تلقائيًا إذا تُرك فارغًا"), Fld("Barcode"),
            Fld("ProductName", span=2), Fld("ProductNameEn", span=2),
            Fld("CategoryID", rows=CATEGORY_ROWS), Fld("UnitID", rows=UNIT_ROWS),
            Fld("SupplierID", rows=SUPPLIER_ROWS), Fld("VATCategory", rows=VAT_CATEGORY_LIST, widths="0;4.5", hook=True),
            Fld("SellingPrice", label="سعر البيع", hook=True), Fld("PurchasePrice"),
            Info("lblPriceInfo"),
            Fld("AverageCost", locked=True), Fld("CurrentQuantity", locked=True),
            Fld("MinimumQuantity"), Fld("ProductLocation"),
            Fld("IsActive"), Info("lblStockNote", "الكمية تتغير فقط من المشتريات والمبيعات والجرد"),
            Fld("TrackStock", hint="ألغِ العلامة للوجبات والمشروبات التي تُحضَّر عند الطلب: تُباع بلا رصيد"),
            Fld("SizePriceM", hint="الكافيه: سعر البيع = الصغير؛ سعر الوسط أو الكبير يجعل للمشروب أحجامًا"),
            Fld("SizePriceL"),
            Fld("ImagePath", hint="صورة الزر في شاشة اللمس: مسار كامل أو اسم ملف في مجلد الصور",
                button=("btnBrowseImage", "استعراض", 'BrowseFile Me, "ImagePath"')),
            Fld("Notes", span=2),
        ]),
    DataScreen(
        "frmCustomers", "Customers", "العملاء", "بيانات العملاء وأرصدتهم", "customers",
        list_select="t.CustomerName AS [العميل], t.Mobile AS [الجوال], t.CurrentBalance AS [الرصيد]",
        list_from="Customers AS t", list_order="t.CustomerName",
        list_headers=[("العميل", 4.4), ("الجوال", 2.4), ("الرصيد", 1.6)],
        search=["t.CustomerName", "t.Mobile", "t.Phone", "t.VATNumber"],
        active="t.IsActive",
        extra_buttons=[("btnPayment", "سند قبض", 'OpenScreen "frmCustomerPayment", 6, Me!CustomerID'),
                       ("btnStatement", "كشف حساب", 'PrintPartyStatement "C", Me!CustomerID')],
        fields=[
            Fld("CustomerName", span=2), Fld("Mobile"), Fld("Phone"),
            Fld("Email"), Fld("VATNumber", hint="للعملاء المنشآت (فاتورة ضريبية)"),
            Fld("CRNumber"), Fld("City"), Fld("District"), Fld("StreetName"),
            Fld("BuildingNo"), Fld("PostalCode"), Fld("Address", span=2),
            Fld("OpeningBalance", hint="يُقفل بعد أول عملية"), Fld("CurrentBalance", locked=True),
            Fld("AllowCredit"), Fld("CreditLimit", hint="0 = بدون حد"),
            Fld("IsActive"), Info("lblBalanceNote", "الرصيد الموجب = مبلغ مستحق على العميل"),
            Fld("Notes", span=2),
        ]),
    DataScreen(
        "frmSuppliers", "Suppliers", "الموردون", "بيانات الموردين وأرصدتهم", "suppliers",
        list_select="t.SupplierName AS [المورد], t.Mobile AS [الجوال], t.CurrentBalance AS [الرصيد]",
        list_from="Suppliers AS t", list_order="t.SupplierName",
        list_headers=[("المورد", 4.4), ("الجوال", 2.4), ("الرصيد", 1.6)],
        search=["t.SupplierName", "t.ContactPerson", "t.Mobile", "t.VATNumber"],
        active="t.IsActive",
        extra_buttons=[("btnPayment", "سند صرف", 'OpenScreen "frmSupplierPayment", 7, Me!SupplierID'),
                       ("btnStatement", "كشف حساب", 'PrintPartyStatement "S", Me!SupplierID')],
        fields=[
            Fld("SupplierName", span=2), Fld("ContactPerson"), Fld("Mobile"),
            Fld("Phone"), Fld("Email"), Fld("VATNumber"), Fld("CRNumber"),
            Fld("City"), Info("lblSupplierNote", " "), Fld("Address", span=2),
            Fld("OpeningBalance", hint="يُقفل بعد أول عملية"), Fld("CurrentBalance", locked=True),
            Fld("IsActive"), Info("lblBalanceNote", "الرصيد الموجب = مبلغ مستحق للمورد"),
            Fld("Notes", span=2),
        ]),
    DataScreen(
        "frmExpenses", "Expenses", "المصروفات", "تسجيل مصروفات المحل", "expenses",
        list_select=("t.ExpenseNumber AS [الرقم], t.ExpenseDate AS [التاريخ], "
                     "x.ExpenseTypeName AS [النوع], t.TotalAmount AS [المبلغ]"),
        list_from="Expenses AS t INNER JOIN ExpenseTypes AS x ON t.ExpenseTypeID = x.ExpenseTypeID",
        list_order="t.ExpenseDate DESC, t.ExpenseID DESC",
        list_headers=[("الرقم", 2.0), ("التاريخ", 2.1), ("النوع", 2.6), ("المبلغ", 1.7)],
        search=["t.ExpenseNumber", "t.Description", "x.ExpenseTypeName", "t.SupplierInvoiceRef"],
        seq="EXPENSE:ExpenseNumber", unique=["ExpenseNumber"],
        extra_buttons=[("btnExpenseTypes", "أنواع المصروفات", 'OpenScreen "frmExpenseTypes"'),
                       ("btnTreasury", "الخزينة", 'OpenScreen "frmTreasury"')],
        fields=[
            Fld("ExpenseNumber", locked=True, hint="يُولَّد عند الحفظ"), Fld("ExpenseDate"),
            Fld("ExpenseTypeID", rows=EXPENSE_TYPE_ROWS,
                button=("btnNewType", "نوع جديد", 'AddExpenseType Me, "ExpenseTypeID"')),
            Fld("PaymentMethodID", rows=PAYMENT_ROWS, hook=True),
            Fld("Amount", hook=True), Fld("Tax", hook=True,
                                         button=("btnCalcVat", "احسب 15%", "CalcExpenseVat Me")),
            Fld("TotalAmount", locked=True), Fld("SupplierInvoiceRef"),
            Fld("CashBoxID", rows=CASHBOX_ROWS, hint="المصروف النقدي يُخصم من هذا الصندوق (يُختار صندوقك تلقائيًا)"),
            Info("lblCashNote", "الدفع النقدي يُخصم من الصندوق"),
            Fld("Description", span=2),
        ]),
    DataScreen(
        "frmUsers", "Employees", "المستخدمون", "الموظفون وأسماء الدخول والأدوار", "users",
        list_select="t.Username AS [المستخدم], t.EmployeeName AS [الاسم], r.RoleName AS [الدور]",
        list_from="Employees AS t INNER JOIN Roles AS r ON t.RoleID = r.RoleID",
        list_order="t.EmployeeName",
        list_headers=[("المستخدم", 2.4), ("الاسم", 3.8), ("الدور", 2.2)],
        search=["t.EmployeeName", "t.Username", "t.Mobile"], active="t.IsActive", unique=["Username"],
        allow_delete=False,
        extra_buttons=[("btnSetPassword", "كلمة المرور", 'OpenScreen "frmChangePassword", 10, Me!EmployeeID'),
                       ("btnUnlock", "فك القفل", "UnlockUser Me"),
                       ("btnRoles", "صلاحيات الأدوار", 'OpenScreen "frmRoles", 10'),
                       ("btnUserScreens", "صلاحيات الشاشات", 'OpenScreen "frmUserScreens", 10')],
        fields=[
            Fld("EmployeeName", span=2), Fld("Username", hint="بدون مسافات، 3 أحرف على الأقل"),
            Fld("RoleID", rows=ROLE_ROWS, widths="0;4"),
            Fld("JobTitle"), Fld("Mobile"),
            Fld("MaxDiscountPercent", hint="أقصى خصم بدون موافقة (مثال 5%)"), Fld("IsActive"),
            Fld("CashBoxID", rows=CASHBOX_ROWS,
                hint="نقدية مبيعات المستخدم وسنداته تدخل هذا الصندوق؛ فارغ = أول صندوق كاشير"),
            Info("lblCashBoxNote", "المدير المسؤول عن الخزينة: اختر له الخزينة الرئيسية"),
            Fld("MustChangePassword"), Fld("LastLoginAt", locked=True),
            Fld("FailedLoginCount", locked=True), Fld("LockedUntil", locked=True),
            Info("lblPasswordState"),
            Fld("Notes", span=2),
        ]),
    DataScreen(
        "frmCategories", "Categories", "التصنيفات", "تصنيفات المنتجات", "category",
        list_select="t.CategoryName AS [التصنيف]", list_from="Categories AS t",
        list_order="t.CategoryName", list_headers=[("التصنيف", 8.4)],
        search=["t.CategoryName", "t.Description"], active="t.IsActive", unique=["CategoryName"],
        fields=[Fld("CategoryName", span=2), Fld("Description", span=2), Fld("IsActive"),
                Fld("SortOrder", hint="ترتيب الزر في شاشة اللمس (الأصغر أولًا)"),
                Fld("IsAddOn", hint="الكافيه: أصناف هذه الفئة تظهر كإضافات للمشروب (حليب، شوت إضافي...)"),
                Fld("TileColor", rows=TILE_COLORS),
                Fld("ImagePath", hint="صورة الزر في شاشة اللمس",
                    button=("btnBrowseImage", "استعراض", 'BrowseFile Me, "ImagePath"'))]),
    DataScreen(
        "frmUnits", "Units", "وحدات القياس", "وحدات بيع المنتجات", "category",
        list_select="t.UnitName AS [الوحدة], t.ZatcaUnitCode AS [الرمز]", list_from="Units AS t",
        list_order="t.UnitName", list_headers=[("الوحدة", 5.4), ("الرمز", 3.0)],
        search=["t.UnitName", "t.ZatcaUnitCode"], active="t.IsActive", unique=["UnitName"],
        fields=[Fld("UnitName"), Fld("ZatcaUnitCode", hint="مثال: PCE للحبة، KGM للكيلو"),
                Fld("IsActive")]),
    DataScreen(
        "frmExpenseTypes", "ExpenseTypes", "أنواع المصروفات", "قائمة أنواع المصروفات", "expenses",
        list_select="t.ExpenseTypeName AS [النوع]", list_from="ExpenseTypes AS t",
        list_order="t.ExpenseTypeName", list_headers=[("النوع", 8.4)],
        search=["t.ExpenseTypeName"], active="t.IsActive", unique=["ExpenseTypeName"],
        fields=[Fld("ExpenseTypeName", span=2), Fld("IsActive")]),
    DataScreen(
        "frmCashBoxes", "CashBoxes", "الصناديق", "الخزينة الرئيسية وصناديق الكاشير", "treasury",
        list_select="t.BoxName AS [الصندوق], IIf(t.BoxType = 'MAIN', 'خزينة', 'كاشير') AS [النوع]",
        list_from="CashBoxes AS t", list_order="t.BoxType DESC, t.BoxName",
        list_headers=[("الصندوق", 5.6), ("النوع", 2.8)],
        search=["t.BoxName", "t.Notes"], active="t.IsActive", unique=["BoxName"],
        extra_buttons=[("btnTreasury", "الخزينة", 'OpenScreen "frmTreasury"')],
        fields=[Fld("BoxName", span=2), Fld("BoxType", rows=BOX_TYPES, widths="0;5"),
                Fld("IsActive"),
                Fld("OpeningBalance", hint="النقدية الموجودة في الصندوق عند بدء استخدام البرنامج"),
                Fld("OpeningDate"),
                Info("lblBoxNote", "الرصيد لا يُكتب يدويًا: يُحسب من المبيعات والسندات والمصروفات"),
                Fld("Notes", span=2)]),
    DataScreen(
        "frmAccounts", "Accounts", "دليل الحسابات", "شجرة الحسابات: الحسابات الرئيسية والفرعية", "journal",
        list_select="t.AccountCode AS [الرقم], Space((t.AccountLevel - 1) * 3) & t.AccountName AS [الحساب], "
                    "IIf(t.IsPosting, 'فرعي', 'رئيسي') AS [النوع]",
        list_from="Accounts AS t", list_order="t.TreeKey",
        list_headers=[("الرقم", 1.8), ("الحساب", 5.4), ("النوع", 1.2)],
        search=["t.AccountName"], active="t.IsActive", unique=["AccountCode"],
        extra_buttons=[("btnJournal", "قيود اليومية", 'OpenScreen "frmJournal"'),
                       ("btnStatement", "كشف حساب", 'OpenScreen "frmLedger", 0, Me!AccountCode'),
                       ("btnManual", "قيد يدوي", 'OpenScreen "frmManualEntry"')],
        fields=[Fld("AccountCode", hint="رقم جديد لا يتكرر؛ لا يتغير بعد الحفظ"),
                Fld("ParentCode", rows=ACCOUNT_ROWS, widths="0;7", hook=True,
                    hint="الحساب الرئيسي الذي يتبعه (نوع الحساب يتبعه تلقائيًا)"),
                Fld("AccountName", span=2),
                Fld("AccountType", rows=ACCOUNT_TYPES, widths="0;5"),
                Fld("IsPosting", hint="فرعي = تُكتب عليه القيود؛ رئيسي = يجمع حساباته التابعة فقط"),
                Fld("IsActive"), Fld("IsSystem", locked=True), Fld("AccountLevel", locked=True),
                Info("lblAccountNote", "الحسابات الأساسية (المعلَّمة) تستخدمها القيود الآلية: لا تُحذف ولا يتغير نوعها. "
                                       "القيود اليدوية من زر «قيد يدوي».")]),
    DataScreen(
        "frmSettings", "Settings", "الإعدادات", "بيانات المحل الضريبية وإعدادات التشغيل", "settings",
        kind="SINGLE", allow_add=False, allow_delete=False,
        record_source="SELECT * FROM Settings WHERE SettingID = 1",
        extra_buttons=[("btnCategories", "التصنيفات", 'OpenScreen "frmCategories"'),
                       ("btnUnits", "الوحدات", 'OpenScreen "frmUnits"'),
                       ("btnExpenseTypes", "أنواع المصروفات", 'OpenScreen "frmExpenseTypes"'),
                       ("btnLabelSettings", "ملصقات الباركود", 'OpenScreen "frmLabelSettings"'),
                       ("btnActivation", "تفعيل البرنامج", 'OpenScreen "frmActivation", 10')],
        fields=[
            Fld("StoreName"), Fld("StoreNameEn"),
            Fld("VATNumber", hint="15 رقمًا يبدأ وينتهي بـ 3"), Fld("CRNumber"),
            Fld("BuildingNo"), Fld("StreetName"), Fld("District"), Fld("City"),
            Fld("PostalCode"), Fld("AdditionalNo"), Fld("Phone"), Fld("Email"),
            Fld("VATRate"), Fld("PricesIncludeVAT"),
            Fld("AllowNegativeStock"), Fld("SlowMovingDays"),
            Fld("BackupFolder", button=("btnBrowseBackup", "استعراض", 'BrowseFolder Me, "BackupFolder"')),
            Fld("BackupKeepCount"),
            Fld("LogoPath", button=("btnBrowseLogo", "استعراض", 'BrowseFile Me, "LogoPath"')),
            Fld("ReceiptFooter"),
            Fld("POSMode", rows=POS_MODES, widths="0;6",
                hint="الشاشة التي يفتحها زر المبيعات"),
            Fld("ImagesFolder", hint="فارغ = مجلد Images بجانب ملف البيانات",
                button=("btnBrowseImages", "استعراض", 'BrowseFolder Me, "ImagesFolder"')),
            Fld("InvoicePrintMode", rows=INVOICE_PRINT_MODES, widths="0;6",
                hint="عند حفظ فاتورة البيع أو المرتجع"),
            Info("lblStoreNameNote"),
            Fld("AllowAdminCompanyName", hint="يظهر للمبرمج فقط"),
        ]),
    DataScreen(
        "frmLabelSettings", "LabelSettings", "إعدادات ملصقات الباركود",
        "مقاس الملصق والورق والهوامش، وحجم الباركود، والنصوص أعلاه وأسفله", "settings",
        kind="SINGLE", allow_add=False, allow_delete=False,
        record_source="SELECT * FROM LabelSettings WHERE LabelSettingID = 1",
        fields=[
            Info("lblInfoPaper", "الملصق والورق (بالمليمتر). مقاس ورق الطابعة نفسه يُضبط من إعدادات الطابعة في Windows"),
            Fld("PrinterName", rows="PRINTERS", hint="اتركه فارغًا للطباعة على الطابعة الافتراضية"),
            Fld("LabelsAcross", hint="1 لطابعة الملصقات، وأكثر لورق A4 فيه أعمدة ملصقات"),
            Fld("LabelWidth", hint="مثال: 38 أو 40 أو 50"), Fld("LabelHeight", hint="مثال: 25 أو 30"),
            Fld("ColumnGap"), Fld("RowGap"),
            Fld("MarginTop"), Fld("MarginBottom"), Fld("MarginRight"), Fld("MarginLeft"),
            Info("lblInfoBar", "الباركود والنصوص"),
            Fld("BarHeight"), Fld("BarWidth", hint="0.25 مناسب لطابعات 203 نقطة/بوصة، وكبّره إذا صعبت القراءة"),
            Fld("TopLine1", rows=LABEL_LINES), Fld("TopLine2", rows=LABEL_LINES),
            Fld("BottomLine1", rows=LABEL_LINES), Fld("BottomLine2", rows=LABEL_LINES),
            Fld("ShortName", hint="مثال: النخبة. فارغ = اسم المحل من الإعدادات"), Fld("FontSize"),
        ]),
]


@dataclass
class NavItem:
    key: str
    caption: str
    icon: str
    target: str          # form name, or "" for the logout action
    phase: int


NAV_ITEMS: List[NavItem] = [
    NavItem("Sales", "المبيعات", "sales", "frmPOS", 6),
    NavItem("Purchases", "المشتريات", "purchases", "frmPurchaseInvoice", 7),
    NavItem("Inventory", "المخزون", "inventory", "frmInventory", 7),
    NavItem("Products", "المنتجات", "products", "frmProducts", 5),
    NavItem("Customers", "العملاء", "customers", "frmCustomers", 5),
    NavItem("Suppliers", "الموردون", "suppliers", "frmSuppliers", 5),
    NavItem("Expenses", "المصروفات", "expenses", "frmExpenses", 5),
    NavItem("Treasury", "الخزينة", "treasury", "frmTreasury", 0),
    NavItem("Journal", "قيود اليومية", "journal", "frmJournal", 0),
    NavItem("StockCount", "الجرد", "stocktake", "frmStockCount", 7),
    NavItem("Reports", "التقارير", "reports", "frmReportCenter", 5),
    NavItem("Search", "البحث", "search", "frmSearch", 5),
    NavItem("Settings", "الإعدادات", "settings", "frmSettings", 5),
    NavItem("Users", "المستخدمون", "users", "frmUsers", 10),
    NavItem("Backup", "نسخة احتياطية", "backup", "frmBackup", 10),
    NavItem("Logout", "تسجيل الخروج", "logout", "", 0),
]


# Permission needed to open each screen ("" = everyone logged in). Subforms open
# with their parent and are not listed. modAppData.ScreenPermission is generated
# from this table and OpenScreen refuses a screen the user may not open.
SCREEN_PERMISSIONS = {
    "frmMain": "", "frmSearch": "", "frmLogin": "", "frmChangePassword": "",
    "frmPOS": "SALES_POS", "frmSalesInvoice": "SALES_VIEW", "frmSalesReturn": "SALES_RETURN",
    "frmCustomerPayment": "CUSTOMER_PAYMENTS", "frmCustomers": "CUSTOMERS",
    "frmSuppliers": "SUPPLIERS", "frmSupplierPayment": "SUPPLIER_PAYMENTS",
    "frmPurchaseInvoice": "PURCHASES", "frmPurchaseView": "PURCHASES", "frmPurchaseReturn": "PURCHASE_RETURN",
    "frmProducts": "PRODUCTS", "frmCategories": "PRODUCTS", "frmUnits": "PRODUCTS", "frmInventory": "PRODUCTS",
    "frmStockCount": "STOCK_COUNT", "frmExpenses": "EXPENSES", "frmExpenseTypes": "EXPENSES",
    "frmReportCenter": "REPORTS", "frmSettings": "SETTINGS", "frmUsers": "USERS", "frmRoles": "USERS",
    "frmUserScreens": "USERS", "frmActivation": "",
    "frmBackup": "BACKUP",
    "frmBarcodeLabels": "PRODUCTS", "frmLabelSettings": "PRODUCTS",
    "frmTouchPOS": "SALES_POS", "frmTouchPay": "SALES_POS", "frmCafePOS": "SALES_POS",
    "frmCafeItem": "SALES_POS",
    "frmTreasury": "CASH_CLOSING", "frmCashClosing": "CASH_CLOSING",
    "frmCashVoucher": "CASH_BOX", "frmCashBoxes": "CASH_BOX",
    "frmJournal": "JOURNAL", "frmJournalEntry": "JOURNAL", "frmAccounts": "JOURNAL",
    "frmManualEntry": "MANUAL_ENTRY", "frmLedger": "JOURNAL",
}


# --------------------------------------------------------------------------
# Search templates and report catalogue (runtime data, emitted to modAppData)
# Tokens: {LIKE} quoted Like pattern, {NUM} exact id or -1, {FROM}/{TO} date literals
# --------------------------------------------------------------------------
@dataclass
class SearchKind:
    key: str
    caption: str
    sql: str
    widths: List[float]     # cm per column, first = hidden id (0)


SEARCH_KINDS: List[SearchKind] = [
    SearchKind("PRODUCT", "المنتجات",
               "SELECT p.ProductID, p.ProductCode AS [الكود], p.Barcode AS [الباركود], "
               "p.ProductName AS [المنتج], p.CurrentQuantity AS [الكمية], p.SellingPrice AS [السعر] "
               "FROM Products AS p WHERE p.ProductName Like {LIKE} OR p.ProductCode Like {LIKE} "
               "OR p.Barcode Like {LIKE} OR p.ProductNameEn Like {LIKE} OR p.ProductID = {NUM} "
               "ORDER BY p.ProductName",
               [0, 2.5, 3.5, 9, 2.5, 2.5]),
    SearchKind("CUSTOMER", "العملاء",
               "SELECT c.CustomerID, c.CustomerName AS [العميل], c.Mobile AS [الجوال], "
               "c.VATNumber AS [الرقم الضريبي], c.CurrentBalance AS [الرصيد] "
               "FROM Customers AS c WHERE c.CustomerName Like {LIKE} OR c.Mobile Like {LIKE} "
               "OR c.Phone Like {LIKE} OR c.VATNumber Like {LIKE} OR c.CustomerID = {NUM} "
               "ORDER BY c.CustomerName",
               [0, 9, 3.5, 4.5, 3]),
    SearchKind("SUPPLIER", "الموردون",
               "SELECT s.SupplierID, s.SupplierName AS [المورد], s.ContactPerson AS [المسؤول], "
               "s.Mobile AS [الجوال], s.CurrentBalance AS [الرصيد] "
               "FROM Suppliers AS s WHERE s.SupplierName Like {LIKE} OR s.ContactPerson Like {LIKE} "
               "OR s.Mobile Like {LIKE} OR s.VATNumber Like {LIKE} OR s.SupplierID = {NUM} "
               "ORDER BY s.SupplierName",
               [0, 8, 5, 3.5, 3]),
    SearchKind("SALE", "فواتير البيع",
               "SELECT h.SalesInvoiceID, h.InvoiceNumber AS [رقم الفاتورة], h.InvoiceDate AS [التاريخ], "
               "c.CustomerName AS [العميل], h.TotalAmount AS [الإجمالي], h.RemainingAmount AS [المتبقي] "
               "FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID "
               "WHERE (h.InvoiceNumber Like {LIKE} OR c.CustomerName Like {LIKE} OR c.Mobile Like {LIKE}) "
               "AND h.InvoiceDate >= {FROM} AND h.InvoiceDate < {TO} ORDER BY h.InvoiceDate DESC",
               [0, 3.5, 4, 8, 3, 3]),
    SearchKind("PURCHASE", "فواتير الشراء",
               "SELECT h.PurchaseInvoiceID, h.InvoiceNumber AS [رقم الفاتورة], "
               "h.SupplierInvoiceNo AS [فاتورة المورد], h.InvoiceDate AS [التاريخ], "
               "s.SupplierName AS [المورد], h.TotalAmount AS [الإجمالي] "
               "FROM PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID "
               "WHERE (h.InvoiceNumber Like {LIKE} OR h.SupplierInvoiceNo Like {LIKE} "
               "OR s.SupplierName Like {LIKE}) AND h.InvoiceDate >= {FROM} AND h.InvoiceDate < {TO} "
               "ORDER BY h.InvoiceDate DESC",
               [0, 3.5, 3.5, 4, 7, 3]),
]


@dataclass
class ReportEntry:
    key: str
    title: str
    query: str
    report: str
    needs: str = ""
    date_column: str = ""


REPORTS: List[ReportEntry] = [
    ReportEntry("STATISTICS", "الإحصائيات والرسوم البيانية", "SalesByCategoryQuery", "rptStatistics", "P"),
    ReportEntry("DAILY_SALES", "المبيعات اليومية", "DailySalesQuery", "rptDailySales", "D", "SaleDate"),
    ReportEntry("MONTHLY_SALES", "المبيعات الشهرية", "MonthlySalesQuery", "rptMonthlySales", "$"),
    ReportEntry("SALES_PERIOD", "المبيعات حسب فترة", "SalesByPeriodQuery", "rptSalesByPeriod", "Pc"),
    ReportEntry("SALES_PRODUCT", "المبيعات حسب المنتج", "SalesByProductQuery", "rptSalesByProduct", "Pr$"),
    ReportEntry("BEST_SELLING", "أفضل المنتجات مبيعًا", "BestSellingProductsQuery", "rptBestSelling", "P$"),
    ReportEntry("LEAST_SELLING", "أقل المنتجات مبيعًا", "LeastSellingProductsQuery", "rptLeastSelling", "P"),
    ReportEntry("PURCHASES", "المشتريات", "PurchasesQuery", "rptPurchases", "Ps"),
    ReportEntry("STOCK", "المخزون الحالي", "StockBalanceQuery", "rptStockBalance"),
    ReportEntry("LOW_STOCK", "المنتجات منخفضة المخزون", "LowStockQuery", "rptLowStock"),
    ReportEntry("PRODUCT_MOVEMENT", "حركة منتج", "ProductMovementQuery", "rptProductMovement", "PR"),
    ReportEntry("CUSTOMER_STATEMENT", "كشف حساب عميل", "CustomerStatementQuery", "rptCustomerStatement", "PC"),
    ReportEntry("SUPPLIER_STATEMENT", "كشف حساب مورد", "SupplierStatementQuery", "rptSupplierStatement", "PS"),
    ReportEntry("EXPENSES", "المصروفات (تفصيلي)", "ExpensesQuery", "rptExpenses", "Pe"),
    ReportEntry("EXPENSES_BY_TYPE", "المصروفات (إجمالي حسب النوع)", "ExpensesByTypeQuery", "rptExpensesByType", "P"),
    ReportEntry("CASH_STATEMENT", "حركة الخزينة / الصندوق (تفصيلي)", "CashStatementQuery", "rptCashStatement", "Pb#"),
    ReportEntry("CASH_DAILY", "حركة الخزينة اليومية (أول اليوم وآخره)", "CashDailyQuery", "rptCashDaily", "Pb#"),
    ReportEntry("CASH_BALANCES", "أرصدة الخزينة والصناديق", "CashBoxBalanceQuery", "rptCashBalances", "#"),
    ReportEntry("CASH_CLOSINGS", "تصفيات يومية الكاشير", "CashClosingsQuery", "rptCashClosings", "Pb#"),
    ReportEntry("PROFIT", "الأرباح", "ProfitQuery", "rptProfit", "P$"),
    ReportEntry("JOURNAL", "قيود اليومية", "JournalLinesQuery", "rptJournal", "PJ"),
    ReportEntry("TRIAL_BALANCE", "ميزان المراجعة", "TrialBalanceQuery", "rptTrialBalance", "PJ"),
    ReportEntry("TRIAL_BALANCE_TREE", "ميزان المراجعة بالمستويات", "TrialBalanceTreeQuery", "rptTrialBalanceTree",
                "PJ"),
    ReportEntry("ACCOUNT_TREE", "دليل الحسابات (شجرة الحسابات)", "AccountTreeQuery", "rptAccountTree", "J"),
    ReportEntry("GENERAL_LEDGER", "دفتر الأستاذ (كل الحسابات)", "GeneralLedgerQuery", "rptGeneralLedger", "PJ"),
    ReportEntry("SLOW_MOVING", "المنتجات غير المتحركة", "SlowMovingProductsQuery", "rptSlowMoving"),
    ReportEntry("STOCK_BY_CATEGORY", "المخزون حسب التصنيف", "StockByCategoryQuery", "rptStockByCategory"),
    ReportEntry("VAT_SUMMARY", "ملخص ضريبة القيمة المضافة", "VatSummaryQuery", "rptVatSummary", "P$"),
    ReportEntry("CUSTOMER_BALANCES", "أرصدة العملاء", "CustomerBalanceQuery", "rptCustomerBalances"),
    ReportEntry("SUPPLIER_BALANCES", "أرصدة الموردين", "SupplierBalanceQuery", "rptSupplierBalances"),
    ReportEntry("INTEGRITY", "فحص سلامة البيانات", "IntegrityCheckQuery", "rptIntegrityCheck"),
]


# --------------------------------------------------------------------------
# Layout
# --------------------------------------------------------------------------
@dataclass
class FormModel:
    name: str
    caption: str
    width: int
    height: int
    popup: bool
    record_source: str = ""
    tag: str = ""
    allow_add: bool = True
    allow_edit: bool = True
    controls: List[Control] = field(default_factory=list)
    form_events: List[str] = field(default_factory=list)
    code: List[str] = field(default_factory=list)       # module lines
    form_props: Dict[str, object] = field(default_factory=dict)   # extra form properties
    # window fitting (FitControls): name -> (mx, mw, my, mh) in 1/1000 of the extra width/height
    fit: Dict[str, tuple] = field(default_factory=dict)

    def add(self, c: Control) -> Control:
        self.controls.append(c)
        return c


BUTTON_W, BUTTON_H, GAP = cm(2.4), cm(0.85), cm(0.2)


def shrink_area(m: FormModel, names, k: int):
    """Make the named list/grid controls k twips shorter and move everything below them up.
    The screens are designed at their smallest size; fit_window grows them with the window."""
    bottom = min(c.y + c.h for c in m.controls if c.name in names)
    for c in m.controls:
        if c.name in names:
            c.h -= k
        elif c.y >= bottom:
            c.y -= k
    m.height -= k
    for c in m.controls:                        # full-height backgrounds (side bar)
        if c.y == 0 and c.h > m.height:
            c.h = m.height


def fit_window(m: FormModel, split_x=None, bottom_y=None, stretch_w=(), stretch_h=(), extra=None):
    """How each control follows a window larger than the design (modForms.FitControls):
    controls right of split_x move with the right edge, controls below bottom_y move with the
    bottom edge, stretch_w / stretch_h grow; extra gives (mx, mw, my, mh) per mille directly."""
    for c in m.controls:
        mx = mw = my = mh = 0
        if c.name == "boxTitle" or c.name in stretch_w:
            mw = 1000
        elif split_x is not None and c.x >= split_x:
            mx = 1000
        if c.name in stretch_h:
            mh = 1000
        elif bottom_y is not None and c.y >= bottom_y:
            my = 1000
        if extra and c.name in extra:
            mx, mw, my, mh = extra[c.name]
        if mx or mw or my or mh:
            m.fit[c.name] = (mx, mw, my, mh)


def title_band(m: FormModel, title: str, subtitle: str, icon: str):
    m.add(Control("rect", "boxTitle", 0, 0, m.width, cm(1.5), {"BackColor": Sym("CLR_PRIMARY")},
                  decorative=True))
    m.add(Control("icon", "icoTitle", cm(0.4), cm(0.3), cm(0.9), cm(0.9),
                  {"Caption": Sym(f"ChrW(&H{ICONS[icon]:X})"), "FontSize": 20,
                   "ForeColor": Sym("CLR_SURFACE")}, decorative=True))
    m.add(Control("label", "lblTitle", cm(1.5), cm(0.18), cm(14), cm(0.75),
                  {"Caption": title, "FontSize": 16, "FontBold": True,
                   "ForeColor": Sym("CLR_SURFACE")}, decorative=True))
    m.add(Control("label", "lblSubtitle", cm(1.5), cm(0.9), cm(14), cm(0.5),
                  {"Caption": subtitle, "FontSize": 9, "ForeColor": Sym("CLR_SIDEBAR_TEXT")},
                  decorative=True))


def button(m: FormModel, name, caption, x, y, style="secondary", w=BUTTON_W, h=BUTTON_H,
           call=None):
    m.add(Control("button", name, x, y, w, h, {"Caption": caption, "Style": style},
                  events=["Click"]))
    if call:
        m.code += [f"Private Sub {name}_Click()", f"    {call}", "End Sub"]


def kind_of(table_name: str, field_name: str):
    for f in table(table_name).fields:
        if f.name == field_name:
            return f
    raise KeyError(f"{table_name}.{field_name}")


def input_control(m: FormModel, screen: DataScreen, fld: Fld, x, y, w, h, multiline=False):
    f = kind_of(screen.table, fld.field)
    props: Dict[str, object] = {}
    if f.kind == "BOOL":
        c = Control("check", fld.field, x, y + cm(0.15), cm(0.5), cm(0.5), props,
                    source=fld.field)
    elif fld.rows == "PRINTERS":                 # filled when the screen opens (modLabels)
        props.update({"RowSource": "", "RowSourceType": "Value List", "ColumnCount": 1,
                      "ColumnWidths": fld.widths or "8"})
        c = Control("combo", fld.field, x, y, w, h, props, source=fld.field)
    elif fld.rows:
        props["RowSource"] = fld.rows
        props["ColumnCount"] = 2
        props["ColumnWidths"] = fld.widths or "0;6"
        if not fld.rows.lstrip().upper().startswith("SELECT"):
            props["RowSourceType"] = "Value List"
        c = Control("combo", fld.field, x, y, w, h, props, source=fld.field)
    else:
        if f.kind in ("MONEY",):
            props["Format"] = "#,##0.00"
        elif f.kind == "QTY":
            props["Format"] = "#,##0.###"
        elif f.kind == "RATE":
            props["Format"] = "0.00%"
        elif f.kind in ("DATE", "DATETIME"):
            props["Format"] = "yyyy/mm/dd"
        if multiline:
            props["EnterKeyBehavior"] = True
            props["ScrollBars"] = 2
        c = Control("text", fld.field, x, y, w, h, props, source=fld.field)
    if fld.locked:
        c.props["Locked"] = True
        c.props["TabStop"] = False
    if fld.hook:
        c.events.append("AfterUpdate")
        m.code += [f"Private Sub {fld.field}_AfterUpdate()",
                   f'    FieldChanged Me, "{fld.field}"', "End Sub"]
    m.add(c)
    return c, f


def layout_data_screen(s: DataScreen) -> FormModel:
    width = cm(27.0)
    m = FormModel(s.name, s.caption, width, 0, popup=True,
                  record_source=s.record_source or f"SELECT * FROM {s.table}",
                  tag=s.tag(), allow_add=s.allow_add)
    title_band(m, s.caption, s.subtitle, s.icon)

    # toolbar
    y = cm(1.8)
    x = cm(0.4)
    toolbar = []
    if s.allow_add:
        toolbar.append(("btnNew", "جديد", "secondary", 'FormAction Me, "NEW"'))
    toolbar.append(("btnSave", "حفظ", "primary", 'FormAction Me, "SAVE"'))
    toolbar.append(("btnUndo", "تراجع", "secondary", 'FormAction Me, "UNDO"'))
    if s.allow_delete:
        toolbar.append(("btnDelete", "حذف", "danger", 'FormAction Me, "DELETE"'))
    for name, caption, style, call in toolbar:
        button(m, name, caption, x, y, style, call=call)
        x += BUTTON_W + GAP
    for name, caption, call in s.extra_buttons:
        button(m, name, caption, x, y, "secondary", w=cm(3.0), call=call)
        x += cm(3.0) + GAP
    button(m, "btnClose", "إغلاق", width - cm(0.4) - BUTTON_W, y, "secondary",
           call='FormAction Me, "CLOSE"')

    top = cm(3.0)
    if s.kind == "LIST":
        lx, lw = cm(0.4), cm(8.8)
        m.add(Control("label", "lblSearch", lx, top, cm(5.5), cm(0.5),
                      {"Caption": "بحث (F3)", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")},
                      decorative=True))
        m.add(Control("label", "lblCount", lx + cm(5.6), top, lw - cm(5.6), cm(0.5),
                      {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED"),
                       "TextAlign": 3}, decorative=True))
        m.add(Control("text", "txtSearch", lx, top + cm(0.5), lw, cm(0.8), {}, events=["Change"]))
        m.code += ["Private Sub txtSearch_Change()", "    RefreshList Me", "End Sub"]
        list_top = top + cm(1.5)
        if s.active:
            m.add(Control("check", "chkShowInactive", lx, list_top + cm(0.05), cm(0.5), cm(0.5),
                          {"DefaultValue": "False"}, events=["AfterUpdate"]))
            m.add(Control("label", "lblShowInactive", lx + cm(0.6), list_top, cm(5), cm(0.6),
                          {"Caption": "إظهار غير النشط", "FontSize": 9,
                           "ForeColor": Sym("CLR_MUTED")}, decorative=True))
            m.code += ["Private Sub chkShowInactive_AfterUpdate()", "    RefreshList Me", "End Sub"]
            list_top += cm(0.8)
        widths = ";".join(["0"] + [f"{w:g}" for _, w in s.list_headers])
        list_ctl = m.add(Control("list", "lstItems", lx, list_top, lw, 0,
                                 {"ColumnCount": len(s.list_headers) + 1, "ColumnWidths": widths,
                                  "ColumnHeads": True}, events=["AfterUpdate"]))
        m.code += ["Private Sub lstItems_AfterUpdate()", "    ListPick Me", "End Sub"]
        dx, dw = cm(9.6), width - cm(0.4) - cm(9.6)
        label_w = cm(3.0)
    else:
        list_ctl = None
        dx, dw = cm(0.4), width - cm(0.8)
        label_w = cm(4.0)

    # field grid: two columns, label on the start side of each input
    col_gap = cm(0.4)
    col_w = (dw - col_gap) // 2
    row_h, ctl_h = cm(1.0), cm(0.75)
    y = top
    col = 0
    for item in s.fields:
        if isinstance(item, Info):
            if col == 1:
                cx = dx + col_w + col_gap
                m.add(Control("label", item.name, cx, y, col_w, ctl_h,
                              {"Caption": item.text, "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
                y += row_h
                col = 0
            else:
                m.add(Control("label", item.name, dx, y, dw, ctl_h,
                              {"Caption": item.text, "FontSize": 10, "ForeColor": Sym("CLR_ACCENT"),
                               "FontBold": True}))
                y += row_h
            continue
        f = kind_of(s.table, item.field)
        span = 2 if item.span == 2 else 1
        if span == 2 and col == 1:
            y += row_h
            col = 0
        cx = dx + (col_w + col_gap) * col
        w_total = dw if span == 2 else col_w
        tall = f.kind == "MEMO" or (span == 2 and item.field == "Description")
        h = cm(1.6) if tall else ctl_h
        input_w = w_total - label_w - cm(0.1)
        if item.button:
            input_w -= cm(2.3)
        ctl, f = input_control(m, s, item, cx + label_w + cm(0.1), y, input_w, h, multiline=tall)
        caption = item.label or f.caption
        seq_field = s.seq.split(":")[1] if s.seq else ""
        if (f.required and f.default is None and f.kind not in ("AUTO", "BOOL")
                and not item.locked and item.field != seq_field):
            caption += " *"
        m.add(Control("label", "lbl" + item.field, cx, y, label_w, ctl_h,
                      {"Caption": caption, "FontSize": 10, "ForeColor": Sym("CLR_MUTED")},
                      parent=item.field))
        if item.hint:
            ctl.props["ControlTipText"] = item.hint
            ctl.props["StatusBarText"] = item.hint
        if item.button:
            bname, bcaption, call = item.button
            button(m, bname, bcaption, cx + w_total - cm(2.2), y, "secondary", w=cm(2.2),
                   h=ctl_h, call=call)
        advance = h + (row_h - ctl_h)
        if span == 2:
            y += advance
            col = 0
        elif col == 0:
            col = 1
            row_advance = advance
        else:
            y += max(advance, row_advance)
            col = 0
    if col == 1:
        y += row_advance
    status_y = y + cm(0.2)
    m.add(Control("label", "lblStatus", dx, status_y, dw, cm(0.6),
                  {"Caption": " ", "FontSize": 10, "FontBold": True,
                   "ForeColor": Sym("CLR_MUTED")}))
    m.height = max(status_y + cm(1.2), cm(14.5))
    if list_ctl:
        list_ctl.h = m.height - list_ctl.y - cm(0.5)

    m.form_events = ["Load", "Current", "BeforeUpdate", "AfterUpdate", "Error", "KeyDown", "Unload"]
    m.code = [
        "Private Sub Form_Load()", "    FormLoad Me", "End Sub",
        "Private Sub Form_Current()", "    FormCurrent Me", "End Sub",
        "Private Sub Form_BeforeUpdate(Cancel As Integer)", "    Cancel = Not FormBeforeUpdate(Me)", "End Sub",
        "Private Sub Form_AfterUpdate()", "    FormAfterUpdate Me", "End Sub",
        "Private Sub Form_Error(DataErr As Integer, Response As Integer)",
        "    Response = FormError(Me, DataErr)", "End Sub",
        "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)",
        "    FormKeyDown Me, KeyCode, Shift", "End Sub",
        "Private Sub Form_Unload(Cancel As Integer)", "    Cancel = Not FormUnload(Me)", "End Sub",
    ] + m.code
    return m


DASHBOARD_TILES = [("TODAY", "مبيعات اليوم"), ("MONTH", "مبيعات الشهر"), ("PROFIT", "صافي ربح الشهر (تقريبي)"),
                   ("LOW", "منتجات منخفضة المخزون"), ("DEBT", "ديون العملاء"), ("DUE", "مستحقات الموردين"),
                   ("STOCK", "قيمة المخزون بالتكلفة"), ("EXPENSES", "مصروفات الشهر")]


# figure cards 1-4: icon and colour of the square
KPI_STYLE = {"TODAY": ("sales", (67, 160, 71)), "MONTH": ("reports", (30, 136, 229)),
             "PROFIT": ("expenses", (229, 57, 53)), "LOW": ("inventory", (251, 140, 0))}

# launcher tiles: (NAV_ITEMS key, caption, colour, light tile)
LAUNCH_TILES = [
    ("Sales", "المبيعات", (67, 160, 71), False), ("Purchases", "المشتريات", (30, 136, 229), False),
    ("Inventory", "المخزون", (251, 140, 0), False), ("Products", "المنتجات", (142, 36, 170), False),
    ("Customers", "العملاء", (229, 57, 53), False), ("Suppliers", "الموردون", (57, 73, 171), False),
    ("Expenses", "المصروفات", (0, 137, 123), False), ("Reports", "التقارير", (216, 27, 96), False),
    ("Settings", "الإعدادات", (232, 236, 243), True), ("Users", "المستخدمون", (232, 236, 243), True),
    ("Treasury", "الخزينة", (0, 121, 107), False), ("Logout", "تسجيل الخروج", (244, 81, 30), False),
]


def layout_main() -> FormModel:
    width, height = cm(33.5), cm(19.0)
    m = FormModel("frmMain", "نظام إدارة المحل", width, height, popup=False, allow_add=False,
                  allow_edit=False)
    side_w = cm(6.2)
    m.add(Control("rect", "boxSidebar", 0, 0, side_w, height, {"BackColor": Sym("CLR_PRIMARY")},
                  decorative=True))
    m.add(Control("icon", "icoApp", cm(0.4), cm(0.45), cm(1.0), cm(1.0),
                  {"Caption": Sym(f"ChrW(&H{ICONS['home']:X})"), "FontSize": 22,
                   "ForeColor": Sym("CLR_SURFACE")}, decorative=True))
    m.add(Control("label", "lblAppTitle", cm(1.5), cm(0.4), side_w - cm(1.7), cm(0.75),
                  {"Caption": "نظام إدارة المحل", "FontSize": 15, "FontBold": True,
                   "ForeColor": Sym("CLR_SURFACE")}, decorative=True))
    m.add(Control("label", "lblStoreName", cm(1.5), cm(1.15), side_w - cm(1.7), cm(0.55),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_SIDEBAR_TEXT")},
                  decorative=True))
    y = cm(2.3)
    for item in NAV_ITEMS:
        name = f"btnNav{item.key}"
        call = (f'OpenScreen "{item.target}", {item.phase}' if item.target else "LogoutUser")
        button(m, name, item.caption, cm(0.25), y, "nav", w=side_w - cm(0.5), h=cm(0.83),
               call=call)
        if item.target:
            m.controls[-1].props["Tag"] = item.target       # MainLoad disables what the user may not open
        icon = m.add(Control("icon", f"ico{item.key}", cm(0.45), y + cm(0.1), cm(0.8),
                             cm(0.65), {"Caption": Sym(f"ChrW(&H{ICONS[item.icon]:X})"),
                                        "FontSize": 13, "ForeColor": Sym("CLR_SIDEBAR_TEXT")},
                             events=["Click"], decorative=True))
        m.code += [f"Private Sub {icon.name}_Click()", f"    {call}", "End Sub"]
        y += cm(0.9)

    cx = side_w + cm(0.8)
    cw = width - cx - cm(0.8)
    # header: title and date on the right, buttons and user on the left (Arabic reading order)
    m.add(Control("label", "lblWelcome", cx + cw - cm(12), cm(0.5), cm(12), cm(0.95),
                  {"Caption": "لوحة التحكم", "FontSize": 20, "FontBold": True, "TextAlign": 3,
                   "ForeColor": Sym("CLR_TEXT")}))
    m.add(Control("label", "lblToday", cx + cw - cm(12), cm(1.55), cm(12), cm(0.6),
                  {"Caption": " ", "FontSize": 11, "ForeColor": Sym("CLR_MUTED"), "TextAlign": 3}))
    button(m, "btnRefresh", "تحديث", cx, cm(0.6), "secondary", w=cm(2.4), h=cm(0.8),
           call="DashboardRefresh Me")
    button(m, "btnChangePassword", "كلمة المرور", cx + cm(2.6), cm(0.6), "secondary", w=cm(2.7), h=cm(0.8),
           call='OpenScreen "frmChangePassword", 10')
    m.add(Control("label", "lblUpdated", cx + cm(5.5), cm(0.75), cm(6.0), cm(0.55),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED"), "TextAlign": 1}))
    m.add(Control("label", "lblUser", cx, cm(1.55), cm(11.5), cm(0.6),
                  {"Caption": " ", "FontSize": 11, "ForeColor": Sym("CLR_MUTED"), "TextAlign": 1}))
    gap = cm(0.4)
    tile_w = (cw - gap * 3) // 4
    extra = {"boxSidebar": (0, 0, 0, 1000), "lblWelcome": (1000, 0, 0, 0), "lblToday": (1000, 0, 0, 0)}

    def col_x(i):                                # first card on the right (Arabic reading order)
        return cx + (3 - i % 4) * (tile_w + gap)

    def col_fit(i, my=0, mh=0):
        return ((3 - i % 4) * 250, 250, my, mh)

    # 4 large figure cards with a coloured icon (tiles 1-4)
    for i, (key, caption) in enumerate(DASHBOARD_TILES[:4]):
        n, tx, ty = i + 1, col_x(i), cm(2.5)
        icon_key, rgb = KPI_STYLE[key]
        m.add(Control("rect", f"boxTile{n}", tx, ty, tile_w, cm(2.4), {"BackColor": Sym("CLR_SURFACE")},
                      decorative=True))
        m.add(Control("rect", f"boxKpiIcon{n}", tx + tile_w - cm(1.9), ty + cm(0.4), cm(1.6), cm(1.6),
                      {"BackColor": Sym(f"RGB({rgb[0]}, {rgb[1]}, {rgb[2]})")}, decorative=True))
        m.add(Control("icon", f"icoKpi{n}", tx + tile_w - cm(1.9), ty + cm(0.65), cm(1.6), cm(1.1),
                      {"Caption": Sym(f"ChrW(&H{ICONS[icon_key]:X})"), "FontSize": 22, "TextAlign": 2,
                       "ForeColor": Sym("CLR_SURFACE")}, decorative=True))
        text_w = tile_w - cm(2.4)
        m.add(Control("label", f"lblTileTitle{n}", tx + cm(0.3), ty + cm(0.15), text_w, cm(0.55),
                      {"Caption": caption, "FontSize": 10, "ForeColor": Sym("CLR_MUTED")}, events=["Click"]))
        m.add(Control("label", f"lblTileValue{n}", tx + cm(0.3), ty + cm(0.7), text_w, cm(1.0),
                      {"Caption": "-", "FontSize": 20, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")},
                      events=["Click"]))
        m.add(Control("label", f"lblTileSub{n}", tx + cm(0.3), ty + cm(1.75), text_w, cm(0.5),
                      {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
        for part in ("boxTile", "boxKpiIcon", "icoKpi", "lblTileTitle", "lblTileValue", "lblTileSub"):
            extra[f"{part}{n}"] = col_fit(i)

    # 12 coloured launcher tiles (3 rows share the extra height)
    by_key = {item.key: item for item in NAV_ITEMS}
    row_h, row_gap, top = cm(2.4), cm(0.3), cm(5.2)
    for i, (key, caption, rgb, light) in enumerate(LAUNCH_TILES):
        item = by_key[key]
        tx, ty, row = col_x(i), top + (i // 4) * (row_h + row_gap), i // 4
        call = f'OpenScreen "{item.target}", {item.phase}' if item.target else "LogoutUser"
        fore = "CLR_PRIMARY" if light else "CLR_SURFACE"
        m.add(Control("rect", f"boxNav{key}", tx, ty, tile_w, row_h,
                      {"BackColor": Sym(f"RGB({rgb[0]}, {rgb[1]}, {rgb[2]})")}, decorative=True))
        m.add(Control("icon", f"icoTile{key}", tx, ty + cm(0.3), tile_w, cm(1.15),
                      {"Caption": Sym(f"ChrW(&H{ICONS[item.icon]:X})"), "FontSize": 26, "TextAlign": 2,
                       "ForeColor": Sym(fore)}, decorative=True))
        m.add(Control("label", f"lblTile{key}", tx, ty + cm(1.5), tile_w, cm(0.7),
                      {"Caption": caption, "FontSize": 13, "FontBold": True, "TextAlign": 2,
                       "ForeColor": Sym("CLR_TEXT" if light else "CLR_SURFACE")}, decorative=True))
        button(m, f"btnTile{key}", caption, tx, ty, "secondary", w=tile_w, h=row_h, call=call)
        m.controls[-1].props["Transparent"] = True       # the coloured tile under it shows through
        if item.target:
            m.controls[-1].props["Tag"] = item.target    # MainLoad disables what the user may not open
        for name in (f"boxNav{key}", f"icoTile{key}", f"lblTile{key}", f"btnTile{key}"):
            extra[name] = col_fit(i, row * 333, 333)

    # 4 smaller figure cards (tiles 5-8)
    sy = top + 3 * row_h + 2 * row_gap + cm(0.3)
    for i, (key, caption) in enumerate(DASHBOARD_TILES[4:]):
        n, tx = i + 5, col_x(i)
        m.add(Control("rect", f"boxTile{n}", tx, sy, tile_w, cm(1.85), {"BackColor": Sym("CLR_SURFACE")},
                      decorative=True))
        m.add(Control("label", f"lblTileTitle{n}", tx + cm(0.3), sy + cm(0.1), tile_w - cm(0.6), cm(0.5),
                      {"Caption": caption, "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}, events=["Click"]))
        m.add(Control("label", f"lblTileValue{n}", tx + cm(0.3), sy + cm(0.62), tile_w - cm(0.6), cm(0.72),
                      {"Caption": "-", "FontSize": 15, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")},
                      events=["Click"]))
        m.add(Control("label", f"lblTileSub{n}", tx + cm(0.3), sy + cm(1.35), tile_w - cm(0.6), cm(0.45),
                      {"Caption": " ", "FontSize": 8, "ForeColor": Sym("CLR_MUTED")}))
        for part in ("boxTile", "lblTileTitle", "lblTileValue", "lblTileSub"):
            extra[f"{part}{n}"] = col_fit(i, 1000)
    for i, (key, _) in enumerate(DASHBOARD_TILES):
        for part in ("Title", "Value"):
            m.code += [f"Private Sub lblTile{part}{i + 1}_Click()", f'    DashboardTileClick "{key}"', "End Sub"]

    height = max(sy + cm(1.85) + cm(1.0), y + cm(0.3))     # y: below the last side-menu button
    m.height = height
    m.controls[0].h = height                     # side bar
    m.add(Control("label", "lblIntegrity", cx, height - cm(0.8), cw, cm(0.55),
                  {"Caption": " ", "FontSize": 10, "FontBold": True,
                   "ForeColor": Sym("CLR_MUTED")}))
    extra["lblIntegrity"] = (0, 1000, 1000, 0)
    m.form_events = ["Open", "Load", "Activate", "Timer"]
    m.form_props = {"TimerInterval": 300000}          # refresh the dashboard every 5 minutes
    m.code = ["Private Sub Form_Open(Cancel As Integer)", "    Cancel = Not MainOpen(Me)", "End Sub",
              "Private Sub Form_Load()", "    MainLoad Me", "End Sub",
              "Private Sub Form_Activate()", "    DashboardActivate Me", "End Sub",
              "Private Sub Form_Timer()", "    DashboardRefresh Me", "End Sub"] + m.code
    fit_window(m, extra=extra)
    return m


def labelled(m: FormModel, name, caption, ctl: Control, label_h=cm(0.5)):
    m.add(Control("label", "lbl" + name[3:], ctl.x, ctl.y - label_h - cm(0.05), ctl.w, label_h,
                  {"Caption": caption, "FontSize": 9, "ForeColor": Sym("CLR_MUTED")},
                  parent=ctl.name))


def layout_search() -> FormModel:
    width, height = cm(27.0), cm(17.0)
    m = FormModel("frmSearch", "البحث المتقدم", width, height, popup=False, allow_add=False)
    title_band(m, "البحث المتقدم", "ابحث بالاسم أو الكود أو الباركود أو رقم الفاتورة أو الجوال",
               "search")
    y = cm(2.45)
    kinds = ";".join(f"{k.key};{k.caption}" for k in SEARCH_KINDS)
    cbo = m.add(Control("combo", "cboKind", cm(0.4), y, cm(4.2), cm(0.8),
                        {"RowSource": kinds, "ColumnCount": 2, "ColumnWidths": "0;4",
                         "DefaultValue": '"PRODUCT"'}, events=["AfterUpdate"]))
    labelled(m, "cboKind", "ابحث في", cbo)
    txt = m.add(Control("text", "txtText", cm(4.9), y, cm(8.6), cm(0.8), {}, events=["AfterUpdate"]))
    labelled(m, "txtText", "كلمة البحث (اسم، كود، باركود، رقم، جوال)", txt)
    d1 = m.add(Control("text", "txtFrom", cm(13.8), y, cm(3.2), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtFrom", "من تاريخ", d1)
    d2 = m.add(Control("text", "txtTo", cm(17.3), y, cm(3.2), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtTo", "إلى تاريخ", d2)
    button(m, "btnSearch", "بحث", cm(20.8), y - cm(0.03), "primary", call="RunSearch Me")
    button(m, "btnClear", "مسح", cm(23.4), y - cm(0.03), "secondary", w=cm(1.6),
           call="SearchClear Me")
    button(m, "btnClose", "رجوع", width - cm(0.4) - cm(1.6), y - cm(0.03), "secondary",
           w=cm(1.6), call='DoCmd.Close acForm, Me.Name')
    m.add(Control("list", "lstResults", cm(0.4), cm(3.7), width - cm(0.8), height - cm(3.7) - cm(1.2),
                  {"ColumnHeads": True, "ColumnCount": 6}, events=["DblClick"]))
    m.add(Control("label", "lblCount", cm(0.4), height - cm(0.95), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 10, "ForeColor": Sym("CLR_MUTED")}))
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    SearchLoad Me", "End Sub",
               "Private Sub cboKind_AfterUpdate()", "    SearchKindChanged Me", "End Sub",
               "Private Sub txtText_AfterUpdate()", "    RunSearch Me", "End Sub",
               "Private Sub lstResults_DblClick(Cancel As Integer)", "    SearchOpen Me", "End Sub"]
              + m.code)
    shrink_area(m, ("lstResults",), cm(1.7))
    fit_window(m, split_x=cm(20.7), bottom_y=cm(14.0), stretch_w=("lstResults", "lblCount"),
               stretch_h=("lstResults",))
    return m


def layout_report_center() -> FormModel:
    width, height = cm(27.0), cm(17.0)
    m = FormModel("frmReportCenter", "التقارير", width, height, popup=False, allow_add=False)
    title_band(m, "مركز التقارير", "اختر التقرير ثم حدد الفترة أو العميل أو المنتج", "reports")
    m.add(Control("list", "lstReports", cm(0.4), cm(1.9), cm(9.0), height - cm(2.4),
                  {"RowSourceType": "Value List", "ColumnCount": 2, "ColumnWidths": "0;8.6",
                   "ColumnHeads": False, "FontSize": 11}, events=["AfterUpdate", "DblClick"]))
    px, pw = cm(10.0), width - cm(10.0) - cm(0.4)
    m.add(Control("label", "lblReportTitle", px, cm(1.9), pw, cm(0.85),
                  {"Caption": " ", "FontSize": 16, "FontBold": True,
                   "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("label", "lblNeeds", px, cm(2.8), pw, cm(0.6),
                  {"Caption": " ", "FontSize": 10, "ForeColor": Sym("CLR_MUTED")}))
    y = cm(4.1)
    d1 = m.add(Control("text", "txtFrom", px, y, cm(4.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtFrom", "من تاريخ", d1)
    d2 = m.add(Control("text", "txtTo", px + cm(4.4), y, cm(4.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtTo", "إلى تاريخ", d2)
    y = cm(5.2)
    for i, (name, caption, which) in enumerate([("btnToday", "اليوم", "TODAY"),
                                                ("btnThisMonth", "هذا الشهر", "MONTH"),
                                                ("btnLastMonth", "الشهر الماضي", "LASTMONTH"),
                                                ("btnThisYear", "هذه السنة", "YEAR")]):
        button(m, name, caption, px + i * cm(3.1), y, "secondary", w=cm(2.9), h=cm(0.75),
               call=f'SetQuickPeriod Me, "{which}"')
    y = cm(6.9)
    for i, (name, caption, rows) in enumerate([("cboCustomer", "العميل", CUSTOMER_ROWS),
                                               ("cboSupplier", "المورد", SUPPLIER_ROWS),
                                               ("cboProduct", "المنتج", PRODUCT_ROWS),
                                               ("cboCashBox", "الخزينة / الصندوق", CASHBOX_ROWS),
                                               ("cboExpenseType", "نوع المصروف", EXPENSE_TYPE_ROWS)]):
        cx = px if i % 2 == 0 else px + cm(8.5)          # two columns
        c = m.add(Control("combo", name, cx, y, cm(8.1), cm(0.8),
                          {"RowSource": rows, "ColumnCount": 2, "ColumnWidths": "0;8"}))
        labelled(m, name, caption, c)
        if i % 2 == 1 or i == 4:
            y += cm(1.45)
    button(m, "btnRun", "عرض التقرير", px, y + cm(0.3), "primary", w=cm(5.0), h=cm(1.0),
           call="RunReport Me")
    button(m, "btnPdf", "حفظ PDF", px + cm(5.2), y + cm(0.3), "secondary", w=cm(3.0), h=cm(1.0),
           call='ExportReport Me, "PDF"')
    button(m, "btnExcel", "تصدير Excel", px + cm(8.4), y + cm(0.3), "secondary", w=cm(3.0), h=cm(1.0),
           call='ExportReport Me, "XLSX"')
    button(m, "btnClose", "رجوع", width - cm(0.4) - cm(2.0), y + cm(0.3), "secondary",
           w=cm(2.0), h=cm(1.0), call='DoCmd.Close acForm, Me.Name')
    m.add(Control("label", "lblPhaseNote", px, y + cm(1.7), pw, cm(1.0),
                  {"Caption": "يُعرض التقرير للمعاينة ومنها الطباعة. «حفظ PDF» و«تصدير Excel» يحفظان "
                              "الملف في مجلد Reports بجانب ملف البرنامج.", "FontSize": 9,
                   "ForeColor": Sym("CLR_MUTED")}))
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    ReportCenterLoad Me", "End Sub",
               "Private Sub lstReports_AfterUpdate()", "    ReportSelected Me", "End Sub",
               "Private Sub lstReports_DblClick(Cancel As Integer)", "    RunReport Me", "End Sub"]
              + m.code)
    shrink_area(m, ("lstReports",), cm(1.7))
    fit_window(m, split_x=cm(24.5), stretch_w=("lblReportTitle", "lblNeeds", "lblPhaseNote"),
               stretch_h=("lstReports",))
    return m


def all_forms() -> List[FormModel]:
    from forms_sales import sales_forms
    from forms_purchases import purchase_forms
    from forms_security import security_forms
    from forms_labels import label_forms
    from forms_touch import touch_forms
    from forms_cash import cash_forms
    from forms_journal import journal_forms
    return ([layout_main()] + [layout_data_screen(s) for s in DATA_SCREENS]
            + [layout_search(), layout_report_center()] + sales_forms() + purchase_forms()
            + security_forms() + label_forms() + touch_forms() + cash_forms() + journal_forms())
