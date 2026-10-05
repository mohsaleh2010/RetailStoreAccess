"""Phase 11: the demo data required by the specification, as a dated plan of
documents that modDemoData posts in Access through the REAL posting functions
(PostPurchaseFromCart, PostSaleFromCart, ...), never by writing the tables
directly (masters and expenses are entered like their screens would).

    5 categories, 20 products, 10 customers, 5 suppliers, 3 employees,
    5 purchase invoices, 10 sales invoices, 6 expenses, and a little of
    everything else: a sales return, a purchase return, receipt and payment
    vouchers, a damaged-stock entry and a posted stocktake.

simulate() replays the plan on the SQLite mirror with tools/sim.py (the same
rules as the VBA). The expected end state (stock and average cost of every
product, every balance, totals) is written into modDemoData, which checks the
result in Access after loading (VerifyDemoData).

All names are invented. Day = days before the day the data is loaded.
"""

import os
import sys
from dataclasses import dataclass
from decimal import Decimal as D
from typing import Dict, List, Optional, Tuple

STORE = {"StoreName": "متجر النخبة للتجزئة", "StoreNameEn": "Al Nukhba Retail Store",
         "VATNumber": "310123456700003", "CRNumber": "1010654321", "BuildingNo": "2345",
         "StreetName": "طريق الملك فهد", "District": "العليا", "City": "الرياض", "PostalCode": "12211",
         "Phone": "0112345678"}

CATEGORIES = ["مواد غذائية", "مشروبات", "منظفات", "عناية شخصية", "أدوات منزلية"]

# Units: 1 حبة, 2 علبة, 3 كرتون, 4 باكيت


@dataclass
class Product:
    name: str
    category: int        # index in CATEGORIES (0-based)
    unit: int
    cost: str            # purchase price without VAT
    price: str           # selling price including VAT
    minimum: int
    supplier: int        # index in SUPPLIERS
    barcode: str = ""


def ean13(body12: str) -> str:
    total = sum(int(d) * (3 if i % 2 else 1) for i, d in enumerate(body12))
    return body12 + str((10 - total % 10) % 10)


PRODUCTS: List[Product] = [
    Product("أرز بسمتي 5 كجم", 0, 4, "38.00", "52.00", 10, 0),
    Product("سكر أبيض 2 كجم", 0, 4, "9.50", "13.50", 15, 0),
    Product("زيت دوار الشمس 1.5 لتر", 0, 1, "14.00", "19.95", 12, 0),
    Product("معكرونة 450 جم", 0, 1, "2.60", "3.95", 30, 0),
    Product("تمر سكري 1 كجم", 0, 2, "22.00", "34.50", 8, 0),
    Product("مياه معدنية 330 مل × 40", 1, 3, "11.00", "16.50", 10, 1),
    Product("عصير برتقال 1 لتر", 1, 1, "4.20", "6.50", 20, 1),
    Product("حليب طازج 2 لتر", 1, 1, "7.80", "11.00", 15, 1),
    Product("قهوة عربية 500 جم", 1, 4, "26.00", "39.00", 6, 1),
    Product("شاي أكياس × 100", 1, 2, "9.00", "14.25", 10, 1),
    Product("منظف أرضيات 3 لتر", 2, 1, "11.50", "17.25", 8, 2),
    Product("سائل غسيل صحون 1 لتر", 2, 1, "5.30", "8.50", 12, 2),
    Product("مسحوق غسيل 3 كجم", 2, 4, "24.00", "36.00", 6, 2),
    Product("مناديل ورقية × 5", 2, 4, "8.40", "12.95", 15, 2),
    Product("شامبو 400 مل", 3, 1, "12.00", "18.95", 8, 3),
    Product("معجون أسنان 100 مل", 3, 1, "6.10", "9.50", 12, 3),
    Product("صابون يدين 500 مل", 3, 1, "5.20", "8.25", 10, 3),
    Product("أكياس نفايات كبيرة × 30", 4, 4, "7.00", "10.95", 10, 4),
    Product("ورق ألمنيوم 30 م", 4, 1, "9.20", "14.50", 6, 4),
    Product("بطاريات AA × 4", 4, 4, "8.00", "13.00", 5, 4),
]
for _i, _p in enumerate(PRODUCTS, 1):
    _p.barcode = ean13(f"6281000{_i:05d}")
SLOW_PRODUCT = 20          # never sold, created 150 days ago -> "slow moving" report


@dataclass
class Supplier:
    name: str
    contact: str
    mobile: str
    vat: str
    city: str


SUPPLIERS = [
    Supplier("مؤسسة الوفرة للمواد الغذائية", "سالم العمري", "0551110001", "300112233400003", "الرياض"),
    Supplier("شركة الينابيع للمشروبات", "ماجد الحربي", "0551110002", "300223344500003", "جدة"),
    Supplier("مؤسسة النقاء للمنظفات", "تركي الغامدي", "0551110003", "300334455600003", "الدمام"),
    Supplier("شركة العناية الذهبية", "هند السبيعي", "0551110004", "300445566700003", "الرياض"),
    Supplier("مؤسسة البيت العصري للأدوات", "وليد الشمري", "0551110005", "300556677800003", "الرياض"),
]


@dataclass
class Customer:
    name: str
    mobile: str
    credit: bool
    limit: int
    vat: str = ""
    city: str = "الرياض"


CUSTOMERS = [
    Customer("مطعم الديرة", "0501000001", True, 5000, "300667788900003"),
    Customer("مؤسسة الضيافة للتموين", "0501000002", True, 10000, "300778899000003"),
    Customer("مقهى الركن", "0501000003", True, 3000, "300889900100003"),
    Customer("أحمد محمد العتيبي", "0501000004", True, 1000),
    Customer("سارة عبدالله القحطاني", "0501000005", True, 500),
    Customer("خالد إبراهيم الشهري", "0501000006", False, 0),
    Customer("نورة سعد الدوسري", "0501000007", False, 0),
    Customer("فهد عبدالرحمن الحربي", "0501000008", True, 1500),
    Customer("منى علي الزهراني", "0501000009", False, 0),
    Customer("عبدالله ناصر المطيري", "0501000010", True, 2000),
]

# (name, job, username, role, max discount)
EMPLOYEES = [
    ("مشرف الفرع (تجريبي)", "مدير فرع", "manager", 2, "0.10"),
    ("كاشير الصباح (تجريبي)", "كاشير", "cashier1", 3, "0.02"),
    ("كاشير المساء (تجريبي)", "كاشير", "cashier2", 3, "0.02"),
]
DEMO_PASSWORD = "Demo@2026"          # temporary: every demo user must change it at first login
DEMO_MARK = "DEMO"                   # Notes / Description of every demo master record

# ---------------------------------------------------------------------------
# The plan. Product numbers are 1-based positions in PRODUCTS, customers and
# suppliers 1-based positions in their lists; customer 0 = the cash customer.
# user: None = the administrator who loads the data, else a demo username.
# ---------------------------------------------------------------------------
PLAN = [
    dict(op="cash_in", day=31, hour=8, box="MAIN", category="OWNER", amount="30000.00", party="المالك",
         text="رأس مال تشغيل (تجريبي)"),
    dict(op="purchase", day=30, hour=9, supplier=1, ref="INV-7781", credit=True, paid=1000,
         lines=[(1, 40), (2, 40), (3, 50), (4, 120), (5, 30)]),
    dict(op="purchase", day=29, hour=10, supplier=2, ref="SP-20451", credit=False,
         lines=[(6, 40), (7, 80), (8, 25), (9, 25), (10, 40)]),
    dict(op="expense", day=28, hour=9, type=1, amount="4500.00", tax="0", text="إيجار المحل للشهر"),
    dict(op="purchase", day=27, hour=11, supplier=3, ref="N-3390", credit=True, paid=0,
         lines=[(11, 30), (12, 40), (13, 8), (14, 60)]),
    dict(op="purchase", day=25, hour=12, supplier=4, ref="GC-118", credit=True, paid=500,
         lines=[(15, 30), (16, 15), (17, 40)]),
    dict(op="purchase", day=22, hour=10, supplier=5, ref="BA-5521", credit=False, discount="20.00",
         lines=[(18, 40), (19, 25), (20, 20)]),
    dict(op="sale", day=20, hour=10, user="cashier1", lines=[(4, 10), (2, 2), (3, 1)]),
    dict(op="expense", day=20, hour=12, type=2, amount="380.00", tax="57.00", text="فاتورة الكهرباء"),
    dict(op="expense", day=19, hour=12, type=3, amount="120.00", tax="18.00", text="فاتورة المياه"),
    dict(op="sale", day=18, hour=11, customer=1, credit=True, paid=0, lines=[(1, 10), (3, 6), (6, 5)]),
    dict(op="expense", day=16, hour=13, type=4, amount="299.00", tax="44.85", text="الإنترنت والهاتف"),
    dict(op="sale", day=15, hour=17, user="cashier2", lines=[(7, 6), (16, 2), (15, 1)]),
    dict(op="supplier_payment", day=15, hour=18, supplier=1, amount="1500.00"),
    dict(op="closing", day=15, hour=23, box="CASHIER", short="5.00", keep="100.00"),
    dict(op="purchase_return", day=14, hour=10, invoice=3, lines=[(1, 3)]),
    dict(op="sale", day=12, hour=13, customer=4, credit=True, paid=50, lines=[(5, 2), (9, 1)]),
    dict(op="expense", day=11, hour=9, type=5, amount="150.00", tax="0", text="نقل بضاعة"),
    dict(op="sale", day=10, hour=19, discount="5.00", lines=[(13, 2), (11, 1), (14, 3)]),
    dict(op="supplier_payment", day=9, hour=10, supplier=3, amount="800.00"),
    dict(op="sale", day=8, hour=12, customer=2, credit=True, paid=1000,
         lines=[(1, 15), (2, 25), (4, 60), (6, 20)]),
    dict(op="stock_out", day=7, hour=9, product=12, qty=1, text="تالف - عبوة مكسورة"),
    dict(op="cash_out", day=7, hour=20, box="MAIN", category="OWNER", amount="2000.00", party="المالك",
         text="مسحوبات المالك (تجريبي)"),
    dict(op="customer_payment", day=6, hour=11, customer=1, amount="300.00"),
    dict(op="expense", day=6, hour=15, type=6, amount="260.00", tax="39.00", text="صيانة ثلاجة العرض"),
    dict(op="sale", day=5, hour=18, user="cashier1", lines=[(8, 4), (10, 2), (12, 3), (17, 2)]),
    dict(op="sales_return", day=4, hour=10, invoice=6, lines=[(4, 2)]),
    dict(op="sale", day=3, hour=16, customer=8, credit=True, paid=0, lines=[(9, 2), (5, 1), (19, 1)]),
    dict(op="customer_payment", day=2, hour=10, customer=2, amount="600.00"),
    dict(op="stock_count", day=2, hour=21, category=5, actual={19: -1}),
    dict(op="sale", day=1, hour=20, user="cashier2", lines=[(7, 10), (8, 6), (18, 2)]),
    dict(op="closing", day=1, hour=23, box="CASHIER", short="0", keep="100.00"),
    dict(op="sale", day=0, hour=0, user="cashier1", lines=[(3, 2), (16, 3), (14, 2)]),
]
# sales_return / purchase_return "invoice" = the n-th sale / purchase of the plan (1-based)
# cash_in / cash_out: a cash voucher of the box type (MAIN / CASHIER) by the administrator
# closing: cashier closing of the box: counted = book balance - short, all but "keep" goes to the main safe
# stock_count "actual": product -> difference to the counted quantity (others counted as recorded)


def demo_counts():
    ops = [p["op"] for p in PLAN]
    return {"Categories": len(CATEGORIES), "Products": len(PRODUCTS), "Customers": len(CUSTOMERS),
            "Suppliers": len(SUPPLIERS), "Employees": len(EMPLOYEES), "SalesInvoices": ops.count("sale"),
            "PurchaseInvoices": ops.count("purchase"), "Expenses": ops.count("expense")}


# ---------------------------------------------------------------------------
# simulation on the SQLite mirror
# ---------------------------------------------------------------------------
def _mirror():
    tests = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tests")
    if tests not in sys.path:
        sys.path.insert(0, tests)
    from access_sqlite import AccessOnSqlite, day
    return AccessOnSqlite(), day


@dataclass
class Result:
    db: object
    store: object
    ids: Dict[str, Dict[int, int]]
    stock: Dict[int, D]
    avg: Dict[int, D]
    customer_balance: Dict[int, D]
    supplier_balance: Dict[int, D]
    sales_total: D
    purchases_total: D
    count_lines: int
    count_value: D
    cash_balance: Dict[str, D]


def simulate() -> Result:
    import sim
    db, day = _mirror()
    c = db.con
    s = sim.Store(c)
    ids: Dict[str, Dict[int, int]] = {"cat": {}, "prod": {}, "cust": {}, "sup": {}, "emp": {}}
    for i, name in enumerate(CATEGORIES, 1):
        ids["cat"][i] = s.insert("Categories", CategoryName=name, Description=DEMO_MARK)
    for i, sp in enumerate(SUPPLIERS, 1):
        ids["sup"][i] = s.insert("Suppliers", SupplierName=sp.name, ContactPerson=sp.contact, Mobile=sp.mobile,
                                 VATNumber=sp.vat, City=sp.city, OpeningBalance=0, CurrentBalance=0,
                                 Notes=DEMO_MARK, CreatedAt=day(60))
    for i, p in enumerate(PRODUCTS, 1):
        ids["prod"][i] = s.insert("Products", ProductCode=s.next_number("PRODUCT_CODE"), Barcode=p.barcode,
                                  ProductName=p.name, CategoryID=ids["cat"][p.category + 1], UnitID=p.unit,
                                  PurchasePrice=float(D(p.cost)), AverageCost=0, SellingPrice=float(D(p.price)),
                                  CurrentQuantity=0, MinimumQuantity=p.minimum,
                                  SupplierID=ids["sup"][p.supplier + 1], Notes=DEMO_MARK,
                                  CreatedAt=day(150 if i == SLOW_PRODUCT else 60))
    for i, cu in enumerate(CUSTOMERS, 1):
        ids["cust"][i] = s.insert("Customers", CustomerName=cu.name, Mobile=cu.mobile, VATNumber=cu.vat or None,
                                  City=cu.city, AllowCredit=1 if cu.credit else 0, CreditLimit=cu.limit,
                                  OpeningBalance=0, CurrentBalance=0, Notes=DEMO_MARK, CreatedAt=day(60))
    for i, (name, job, user, role, disc) in enumerate(EMPLOYEES, 1):
        ids["emp"][i] = s.insert("Employees", EmployeeName=name, JobTitle=job, Username=user, RoleID=role,
                                 MaxDiscountPercent=float(D(disc)), Notes=DEMO_MARK)
    users = {e[2]: ids["emp"][i] for i, e in enumerate(EMPLOYEES, 1)}

    sales, purchases = [], []
    count_lines, count_value = 0, D(0)
    for step in PLAN:
        s.now = day(step["day"], step["hour"])
        s.user = users.get(step.get("user"), 1)
        op = step["op"]
        if op == "purchase":
            lines = [(ids["prod"][n], q, PRODUCTS[n - 1].cost) for n, q in step["lines"]]
            purchases.append(s.purchase(ids["sup"][step["supplier"]], lines, step["credit"], step.get("paid"),
                                        True, step["ref"], step.get("discount", 0)))
        elif op == "sale":
            cust = ids["cust"][step["customer"]] if step.get("customer") else 1
            sales.append(s.sale([(ids["prod"][n], q) for n, q in step["lines"]], cust, step.get("credit", False),
                                step.get("paid"), step.get("discount", 0)))
        elif op == "sales_return":
            s.sales_return(sales[step["invoice"] - 1], dict(step["lines"]), cash_refund=False)
        elif op == "purchase_return":
            inv = purchases[step["invoice"] - 1]
            details = {s.one("SELECT PurchaseDetailID FROM PurchaseInvoiceDetails WHERE PurchaseInvoiceID = ? "
                             "AND LineNumber = ?", inv, ln): q for ln, q in step["lines"]}
            s.purchase_return(inv, details)
        elif op == "supplier_payment":
            s.payment(ids["sup"][step["supplier"]], step["amount"])
        elif op == "customer_payment":
            s.customer_payment(ids["cust"][step["customer"]], step["amount"])
        elif op == "expense":
            s.expense(step["type"], step["amount"], step["tax"], step["text"])
        elif op == "stock_out":
            s.manual(ids["prod"][step["product"]], sim.TT["STOCK_OUT"], step["qty"])
        elif op == "stock_count":
            actual = {ids["prod"][n]: s.stock(ids["prod"][n]) + d for n, d in step["actual"].items()}
            _, count_lines, count_value = s.stock_count(ids["cat"][step["category"]], actual)
        elif op in ("cash_in", "cash_out"):
            s.cash_voucher("IN" if op == "cash_in" else "OUT", s.box_of_type(step["box"]), D(step["amount"]),
                           step["category"], party=step["party"], text=step["text"])
        elif op == "closing":
            box = s.box_of_type(step["box"])
            counted = s.cash_balance(box) - D(step["short"])
            transfer = max(D(0), counted - D(step["keep"]))
            s.cash_closing(box, counted, "MAIN", s.box_of_type("MAIN"), transfer)
        else:
            raise ValueError(op)
    s.user = 1
    return Result(
        db, s, ids,
        stock={n: s.stock(pid) for n, pid in ids["prod"].items()},
        avg={n: s.avg(pid) for n, pid in ids["prod"].items()},
        customer_balance={n: sim.dec(s.one("SELECT CurrentBalance FROM Customers WHERE CustomerID = ?", cid))
                          for n, cid in ids["cust"].items()},
        supplier_balance={n: sim.dec(s.one("SELECT CurrentBalance FROM Suppliers WHERE SupplierID = ?", sid))
                          for n, sid in ids["sup"].items()},
        sales_total=sim.dec(s.one("SELECT Sum(TotalAmount) FROM SalesInvoices")),
        purchases_total=sim.dec(s.one("SELECT Sum(TotalAmount) FROM PurchaseInvoices")),
        count_lines=count_lines, count_value=count_value,
        cash_balance={kind: s.cash_balance(s.box_of_type(kind)) for kind in ("MAIN", "CASHIER")})


# Order in which RemoveDemoData empties the tables (children before parents).
DOCUMENT_TABLES = ["CashVouchers", "CashClosings", "SalesReturnDetails", "SalesReturns", "CustomerPayments", "SalesInvoiceDetails", "SalesInvoices",
                   "PurchaseReturnDetails", "PurchaseReturns", "SupplierPayments", "PurchaseInvoiceDetails",
                   "PurchaseInvoices", "StockCountDetails", "StockCounts", "InventoryTransactions", "Expenses"]
DOCUMENT_SEQUENCES = ["SALES_INVOICE", "SALES_RETURN", "PURCHASE_INVOICE", "PURCHASE_RETURN", "CUSTOMER_PAYMENT",
                      "SUPPLIER_PAYMENT", "EXPENSE", "STOCK_COUNT", "STOCK_ADJUST", "PRODUCT_CODE", "ZATCA_ICV",
                      "CASH_IN", "CASH_OUT", "CASH_TRANSFER", "CASH_CLOSING"]
