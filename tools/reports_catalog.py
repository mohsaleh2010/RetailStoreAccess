"""Phase 8: the formatted reports of the report centre (forms.REPORTS).

Every catalogue entry gets a report built from its saved query:
  * LIST reports: title block and column captions on every page (page header),
    one line per record (detail, alternate shading), totals and the record
    count in the report footer, printing time and "page x of y" in the page footer;
  * CARD reports (profit, VAT): one record shown as labelled rows.
Sorting is done by the report (Access reports ignore the ORDER BY of the query).
Widths are in cm from the start (right) edge; the last column absorbs rounding.
"""

from dataclasses import dataclass
from typing import List, Optional, Tuple

from forms import Control, Sym, cm, REPORTS
from reports import (ReportModel, MONEY, SEC_DETAIL, SEC_PAGE_HEADER, SEC_PAGE_FOOTER, SEC_RPT_FOOTER,
                     txt, lbl, hline)

QTY = "#,##0.###"
INT = "#,##0"
PCT = "0.0%"
PORTRAIT_W, LANDSCAPE_W = 19.0, 27.4
NO_DATA = "لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة."


@dataclass
class Col:
    caption: str
    source: str               # field name or "=expression"
    width: float              # cm
    fmt: Optional[str] = None
    total: bool = False       # Sum in the report footer
    running: bool = False     # running sum over the report (balances)
    grow: bool = False


@dataclass
class ListSpec:
    key: str                  # forms.REPORTS key
    cols: List[Col]
    sorts: List[Tuple[str, bool]]
    landscape: bool = False
    no_data: str = NO_DATA
    summary: List[Tuple[str, str]] = None   # (caption, expression) rows in the report footer


@dataclass
class CardSpec:
    key: str
    rows: List[Tuple[str, str, str, bool]]      # caption, source, format, bold
    note: str = ""


def doc_type(field="DocType"):
    return f'=IIf([{field}]="RETURN","مرتجع","فاتورة")'


def pay_type(field="PaymentType"):
    return f'=IIf([{field}]="CREDIT","آجل","نقدي")'


def sales_product_cols():
    return [Col("الكود", "ProductCode", 2.2), Col("المنتج", "ProductName", 5.4, grow=True),
            Col("التصنيف", "CategoryName", 2.6), Col("المباع", "QtySold", 1.8, QTY, True),
            Col("المرتجع", "QtyReturned", 1.8, QTY, True), Col("الصافي", "NetQty", 1.8, QTY, True),
            Col("المبيعات بدون ضريبة", "NetSales", 2.5, MONEY, True),
            Col("الضريبة", "SalesVAT", 2.1, MONEY, True), Col("شامل الضريبة", "SalesTotal", 2.5, MONEY, True),
            Col("التكلفة", "CostOfSales", 2.3, MONEY, True), Col("الربح", "GrossProfit", 2.4, MONEY, True)]


def statement_cols(balance_expr):
    return [Col("التاريخ", "=GDate([EntryDate],True)", 3.0), Col("البيان", "EntryTypeName", 3.6),
            Col("رقم المستند", "DocNumber", 3.0), Col("مدين", "Debit", 3.0, MONEY, True),
            Col("دائن", "Credit", 3.0, MONEY, True), Col("الرصيد", balance_expr, 3.4, MONEY, running=True)]


LIST_SPECS: List[ListSpec] = [
    ListSpec("DAILY_SALES", [
        Col("التاريخ", "=GDate([SaleDate])", 2.4), Col("الفواتير", "InvoiceCount", 1.6, INT, True),
        Col("المرتجعات", "ReturnCount", 1.6, INT, True),
        Col("المبيعات بدون ضريبة", "SalesExVAT", 2.6, MONEY, True),
        Col("المرتجعات بدون ضريبة", "ReturnsExVAT", 2.6, MONEY, True),
        Col("الصافي بدون ضريبة", "NetSalesExVAT", 2.6, MONEY, True), Col("الضريبة", "NetVAT", 2.2, MONEY, True),
        Col("الصافي شامل الضريبة", "NetSalesTotal", 2.7, MONEY, True), Col("نقدي", "CashSales", 2.4, MONEY, True),
        Col("آجل", "CreditSales", 2.4, MONEY, True), Col("المحصّل", "CollectedAmount", 2.3, MONEY, True)],
        [("SaleDate", False)], landscape=True),
    ListSpec("MONTHLY_SALES", [
        Col("الشهر", '=[SalesYear] & "/" & Format([SalesMonth],"00")', 2.0),
        Col("الفواتير", "InvoiceCount", 1.7, INT, True), Col("المرتجعات", "ReturnCount", 1.7, INT, True),
        Col("الصافي بدون ضريبة", "NetSalesExVAT", 2.6, MONEY, True), Col("الضريبة", "NetVAT", 2.2, MONEY, True),
        Col("شامل الضريبة", "NetSalesTotal", 2.6, MONEY, True), Col("التكلفة", "CostOfSales", 2.6, MONEY, True),
        Col("مجمل الربح", "GrossProfit", 3.6, MONEY, True)],
        [("SalesYear", False), ("SalesMonth", False)]),
    ListSpec("SALES_PERIOD", [
        Col("النوع", doc_type(), 1.5), Col("الرقم", "DocNumber", 2.5), Col("التاريخ", "=GDate([DocDate],True)", 2.9),
        Col("العميل", "CustomerName", 4.2, grow=True), Col("الموظف", "EmployeeName", 2.6),
        Col("الدفع", pay_type(), 1.4), Col("بدون ضريبة", "NetAmount", 2.4, MONEY, True),
        Col("الضريبة", "VATAmount", 2.1, MONEY, True), Col("الإجمالي", "GrossAmount", 2.5, MONEY, True),
        Col("المدفوع", "SettledAmount", 2.5, MONEY, True), Col("على الحساب", "OnAccount", 2.8, MONEY, True)],
        [("DocDate", False), ("DocNumber", False)], landscape=True),
    ListSpec("SALES_PRODUCT", sales_product_cols(), [("ProductName", False)], landscape=True),
    ListSpec("BEST_SELLING", sales_product_cols(), [("NetQty", True), ("NetSales", True)], landscape=True),
    ListSpec("LEAST_SELLING", [
        Col("الكود", "ProductCode", 2.4), Col("المنتج", "ProductName", 6.4, grow=True),
        Col("التصنيف", "CategoryName", 3.0), Col("المخزون الحالي", "CurrentQuantity", 2.3, QTY, True),
        Col("صافي المباع", "NetQtySold", 2.3, QTY, True), Col("صافي المبيعات", "NetSalesAmount", 2.6, MONEY, True)],
        [("NetQtySold", False), ("ProductName", False)]),
    ListSpec("PURCHASES", [
        Col("النوع", doc_type(), 1.5), Col("الرقم", "DocNumber", 2.5), Col("فاتورة المورد", "SupplierRef", 2.4),
        Col("التاريخ", "=GDate([DocDate])", 2.2), Col("المورد", "SupplierName", 4.4, grow=True),
        Col("الدفع", pay_type(), 1.4), Col("بدون ضريبة", "NetAmount", 2.5, MONEY, True),
        Col("الضريبة", "VATAmount", 2.2, MONEY, True), Col("الإجمالي", "GrossAmount", 2.6, MONEY, True),
        Col("المدفوع", "SettledAmount", 2.6, MONEY, True), Col("على الحساب", "OnAccount", 3.1, MONEY, True)],
        [("DocDate", False), ("DocNumber", False)], landscape=True),
    ListSpec("STOCK", [
        Col("الكود", "ProductCode", 2.3), Col("المنتج", "ProductName", 5.6, grow=True),
        Col("التصنيف", "CategoryName", 2.7), Col("الوحدة", "UnitName", 1.6),
        Col("الكمية", "CurrentQuantity", 1.9, QTY, True), Col("الحد الأدنى", "MinimumQuantity", 1.8, QTY),
        Col("متوسط التكلفة", "AverageCost", 2.2, MONEY), Col("سعر البيع", "SellingPrice", 2.1, MONEY),
        Col("القيمة بالتكلفة", "StockCostValue", 2.6, MONEY, True),
        Col("القيمة بسعر البيع", "StockSalesValue", 2.6, MONEY, True),
        Col("الحالة", '=IIf([IsLowStock],"منخفض","") & IIf([QuantityMismatch]<>0," / فرق!","")', 2.0)],
        [("CategoryName", False), ("ProductName", False)], landscape=True),
    ListSpec("LOW_STOCK", [
        Col("الكود", "ProductCode", 2.1), Col("المنتج", "ProductName", 4.6, grow=True),
        Col("التصنيف", "CategoryName", 2.3), Col("المتوفر", "CurrentQuantity", 1.7, QTY),
        Col("الحد الأدنى", "MinimumQuantity", 1.7, QTY), Col("النقص", "ShortageQty", 1.6, QTY, True),
        Col("المورد", "SupplierName", 3.0), Col("جوال المورد", "SupplierMobile", 2.0)],
        [("ShortageQty", True), ("ProductName", False)],
        no_data="لا توجد منتجات منخفضة المخزون."),
    ListSpec("PRODUCT_MOVEMENT", [
        Col("التاريخ", "=GDate([MovementDate],True)", 2.9), Col("الحركة", "MovementType", 2.6),
        Col("المستند", "ReferenceNumber", 2.5), Col("وارد", "QtyIn", 1.8, QTY, True),
        Col("صادر", "QtyOut", 1.8, QTY, True), Col("الرصيد", "NetQty", 2.0, QTY, running=True),
        Col("التكلفة", "UnitCost", 2.0, MONEY), Col("ملاحظات", "Notes", 3.4, grow=True)],
        [("SortKey", False), ("MovementDate", False), ("TransactionID", False)]),
    ListSpec("CUSTOMER_STATEMENT", statement_cols("=[Debit]-[Credit]"),
             [("SortKey", False), ("EntryDate", False)]),
    ListSpec("SUPPLIER_STATEMENT", statement_cols("=[Credit]-[Debit]"),
             [("SortKey", False), ("EntryDate", False)]),
    ListSpec("EXPENSES", [
        Col("الرقم", "ExpenseNumber", 2.2), Col("التاريخ", "=GDate([ExpenseDate])", 2.1),
        Col("النوع", "ExpenseTypeName", 2.6), Col("الوصف", "Description", 4.4, grow=True),
        Col("الدفع", "MethodName", 1.8), Col("المبلغ", "Amount", 2.0, MONEY, True),
        Col("الضريبة", "Tax", 1.8, MONEY, True), Col("الإجمالي", "TotalAmount", 2.1, MONEY, True)],
        [("ExpenseDate", False), ("ExpenseNumber", False)]),
    ListSpec("EXPENSES_BY_TYPE", [
        Col("نوع المصروف", "ExpenseTypeName", 6.0), Col("العدد", "ExpenseCount", 2.4, INT, True),
        Col("بدون ضريبة", "AmountExVAT", 3.4, MONEY, True), Col("الضريبة", "InputVAT", 3.2, MONEY, True),
        Col("الإجمالي", "AmountTotal", 4.0, MONEY, True)],
        [("AmountTotal", True)]),
    ListSpec("CASH_STATEMENT", [
        Col("التاريخ", "=GDate([MoveDate],True)", 3.0), Col("الحركة", "MoveTypeName", 3.8),
        Col("المستند", "DocNumber", 2.6), Col("الجهة", "PartyName", 4.0, grow=True),
        Col("البيان", "Details", 3.8, grow=True), Col("الصندوق", "BoxName", 2.8),
        Col("مقبوض", "AmountIn", 2.4, MONEY), Col("مدفوع", "AmountOut", 2.4, MONEY),
        Col("الرصيد", "=[AmountIn]-[AmountOut]", 2.6, MONEY, running=True)],
        [("SortKey", False), ("MoveDate", False)], landscape=True,
        no_data="لا توجد حركة نقدية للاختيارات المحددة.",
        summary=[("رصيد أول المدة", "=Sum(IIf([SortKey]=0,[AmountIn]-[AmountOut],0))"),
                 ("المقبوضات", "=Sum(IIf([SortKey]=1,[AmountIn],0))"),
                 ("المدفوعات والمصروفات", "=Sum(IIf([SortKey]=1,[AmountOut],0))"),
                 ("رصيد آخر المدة", "=Sum([AmountIn]-[AmountOut])")]),
    ListSpec("CASH_DAILY", [
        Col("اليوم", "=GDate([CashDay])", 3.0), Col("رصيد أول اليوم", "OpeningBalance", 3.4, MONEY),
        Col("المقبوضات", "Receipts", 3.4, MONEY, True), Col("المدفوعات", "Payments", 3.4, MONEY, True),
        Col("رصيد آخر اليوم", "ClosingBalance", 3.6, MONEY), Col("الحركات", "MoveCount", 2.2, INT, True)],
        [("CashDay", False)], no_data="لا توجد حركة نقدية في هذه الفترة."),
    ListSpec("FIXED_ASSETS", [
        Col("الرقم", "AssetCode", 2.0), Col("الأصل", "AssetName", 5.2, grow=True), Col("المجموعة", "AssetGroup", 3.6),
        Col("الشراء", "=GDate([PurchaseDate])", 2.2), Col("العمر (شهر)", "UsefulLifeMonths", 1.8, INT),
        Col("التكلفة", "Cost", 2.6, MONEY, True), Col("القسط الشهري", "MonthlyDep", 2.4, MONEY, True),
        Col("مجمع الإهلاك", "AccumDep", 2.6, MONEY, True), Col("القيمة الدفترية", "BookValue", 2.6, MONEY, True),
        Col("الحالة", "StatusName", 2.4)],
        [("Status", False), ("AssetCode", False)], landscape=True, no_data="لا توجد أصول ثابتة."),
    ListSpec("CASH_BALANCES", [
        Col("الصندوق", "BoxName", 5.0, grow=True), Col("النوع", "BoxTypeName", 2.6),
        Col("إجمالي الداخل", "TotalIn", 2.9, MONEY, True), Col("إجمالي الخارج", "TotalOut", 2.9, MONEY, True),
        Col("الرصيد", "Balance", 3.0, MONEY, True), Col("آخر حركة", "=GDate([LastMoveDate])", 2.6)],
        [("BoxType", True), ("BoxName", False)]),
    ListSpec("CASH_CLOSINGS", [
        Col("الرقم", "ClosingNumber", 2.4), Col("التاريخ", "=GDate([ClosingDate],True)", 2.9),
        Col("الصندوق", "BoxName", 3.0), Col("أجراها", "EmployeeName", 2.8),
        Col("الدفتري", "ExpectedBalance", 2.6, MONEY, True), Col("الفعلي", "CountedAmount", 2.6, MONEY, True),
        Col("الفرق", "Difference", 2.2, MONEY, True), Col("الترحيل", "DestinationName", 3.0),
        Col("المرحَّل", "TransferAmount", 2.6, MONEY, True), Col("المتبقي", "KeptAmount", 3.3, MONEY, True)],
        [("ClosingDate", False)], landscape=True, no_data="لا توجد تصفيات في هذه الفترة."),
    ListSpec("JOURNAL", [
        Col("رقم القيد", "EntryNumber", 2.4), Col("التاريخ", "=GDate([EntryDate])", 2.2),
        Col("العملية", "TypeName", 3.0), Col("المستند", "SourceNumber", 2.6), Col("الحساب", "AccountCode", 1.8),
        Col("اسم الحساب", "AccountName", 4.4), Col("البيان", "LineText", 5.0, grow=True),
        Col("مدين", "=IIf([Debit]=0,Null,[Debit])", 2.6, MONEY, True),
        Col("دائن", "=IIf([Credit]=0,Null,[Credit])", 3.4, MONEY, True)],
        [("EntryDate", False), ("EntryNumber", False), ("LineNumber", False)], landscape=True,
        no_data="لا توجد قيود في هذه الفترة."),
    ListSpec("TRIAL_BALANCE", [
        Col("الحساب", "AccountCode", 2.0), Col("اسم الحساب", "AccountName", 5.6, grow=True),
        Col("رصيد أول المدة", "OpeningBalance", 2.8, MONEY, True), Col("مدين الفترة", "PeriodDebit", 2.8, MONEY, True),
        Col("دائن الفترة", "PeriodCredit", 2.8, MONEY, True), Col("الرصيد الختامي", "ClosingBalance", 3.0, MONEY, True)],
        [("AccountCode", False)], no_data="لا توجد قيود.",
        summary=[("مجموع الأرصدة (صفر = متوازن)", "=Sum([ClosingBalance])")]),
    ListSpec("TRIAL_BALANCE_TREE", [
        Col("الحساب", "AccountCode", 2.0), Col("اسم الحساب", "=Space(([AccountLevel]-1)*3) & [AccountName]", 6.4,
                                               grow=True),
        Col("رصيد أول المدة", "OpeningBalance", 2.7, MONEY), Col("مدين الفترة", "PeriodDebit", 2.7, MONEY),
        Col("دائن الفترة", "PeriodCredit", 2.7, MONEY), Col("الرصيد الختامي", "ClosingBalance", 2.9, MONEY)],
        [("TreeKey", False)], no_data="لا توجد قيود.",
        summary=[("مدين الفترة (المستوى الأول)", "=Sum(IIf([AccountLevel]=1,[PeriodDebit],0))"),
                 ("دائن الفترة (المستوى الأول)", "=Sum(IIf([AccountLevel]=1,[PeriodCredit],0))"),
                 ("مجموع أرصدة المستوى الأول (صفر = متوازن)", "=Sum(IIf([AccountLevel]=1,[ClosingBalance],0))")]),
    ListSpec("ACCOUNT_TREE", [
        Col("رقم الحساب", "AccountCode", 2.4), Col("الحساب", "=Space(([AccountLevel]-1)*3) & [AccountName]", 8.0,
                                                  grow=True),
        Col("المستوى", "AccountLevel", 1.6, INT), Col("النوع", "TypeName", 2.4), Col("رئيسي / فرعي", "KindName", 2.4)],
        [("TreeKey", False)], no_data="لا توجد حسابات."),
    ListSpec("SLOW_MOVING", [
        Col("الكود", "ProductCode", 2.2), Col("المنتج", "ProductName", 5.0, grow=True),
        Col("التصنيف", "CategoryName", 2.5), Col("الكمية", "CurrentQuantity", 1.8, QTY, True),
        Col("متوسط التكلفة", "AverageCost", 2.0, MONEY), Col("القيمة", "StockCostValue", 2.2, MONEY, True),
        Col("آخر بيع", "=GDate([LastSaleDate])", 2.0), Col("أيام", "DaysWithoutSale", 1.3, INT)],
        [("DaysWithoutSale", True)], no_data="لا توجد منتجات راكدة."),
    ListSpec("STOCK_BY_CATEGORY", [
        Col("التصنيف", "CategoryName", 6.0), Col("المنتجات", "ProductCount", 2.4, INT, True),
        Col("الكمية", "TotalQuantity", 3.0, QTY, True), Col("القيمة بالتكلفة", "StockCostValue", 3.6, MONEY, True),
        Col("القيمة بسعر البيع", "StockSalesValue", 4.0, MONEY, True)],
        [("CategoryName", False)]),
    ListSpec("CUSTOMER_BALANCES", [
        Col("العميل", "CustomerName", 5.0, grow=True), Col("الجوال", "Mobile", 2.4),
        Col("حد الائتمان", "CreditLimit", 2.2, MONEY), Col("مدين", "DebitTotal", 2.4, MONEY, True),
        Col("دائن", "CreditTotal", 2.4, MONEY, True), Col("الرصيد", "Balance", 2.6, MONEY, True),
        Col("آخر حركة", "=GDate([LastEntryDate])", 2.0)],
        [("CustomerName", False)]),
    ListSpec("SUPPLIER_BALANCES", [
        Col("المورد", "SupplierName", 5.0, grow=True), Col("المسؤول", "ContactPerson", 2.9),
        Col("الجوال", "Mobile", 2.3), Col("مدين", "DebitTotal", 2.2, MONEY, True),
        Col("دائن", "CreditTotal", 2.2, MONEY, True), Col("الرصيد", "Balance", 2.4, MONEY, True),
        Col("آخر حركة", "=GDate([LastEntryDate])", 2.0)],
        [("SupplierName", False)]),
    ListSpec("INTEGRITY", [
        Col("الرمز", "IssueCode", 3.4), Col("المشكلة", "IssueText", 6.8, grow=True),
        Col("الجدول", "SourceTable", 3.0), Col("السجل", "RecordID", 1.6, INT),
        Col("المتوقع", "ExpectedValue", 2.1, QTY), Col("الفعلي", "ActualValue", 2.1, QTY)],
        [("IssueCode", False), ("RecordID", False)],
        no_data="فحص سلامة البيانات: لا توجد أي مشكلات."),
]

CARD_SPECS: List[CardSpec] = [
    CardSpec("PROFIT", [
        ("صافي المبيعات (بدون الضريبة)", "NetSales", MONEY, False),
        ("تكلفة البضاعة المباعة", "CostOfSales", MONEY, False),
        ("مجمل الربح", "GrossProfit", MONEY, True),
        ("نسبة مجمل الربح", "=IIf([NetSales]=0,0,[GrossProfit]/[NetSales])", PCT, False),
        ("فروقات المخزون (جرد، إضافة، خصم)", "InventoryAdjustments", MONEY, False),
        ("المصروفات (بدون الضريبة)", "TotalExpenses", MONEY, False),
        ("صافي الربح", "NetProfit", MONEY, True)],
        "صافي الربح = مجمل الربح ± فروقات المخزون − المصروفات. الرصيد الافتتاحي لا يدخل في الربح."),
    CardSpec("VAT_SUMMARY", [
        ("المبيعات الخاضعة للضريبة (بعد المرتجعات)", "TaxableSales", MONEY, False),
        ("ضريبة المخرجات", "OutputVAT", MONEY, True),
        ("المشتريات الخاضعة للضريبة (بعد المرتجعات)", "TaxablePurchases", MONEY, False),
        ("ضريبة مدخلات المشتريات", "PurchaseVAT", MONEY, False),
        ("ضريبة مدخلات المصروفات", "ExpenseVAT", MONEY, False),
        ("إجمالي ضريبة المدخلات", "InputVAT", MONEY, True),
        ("صافي الضريبة المستحقة", "NetVATDue", MONEY, True)],
        "الموجب مستحق للهيئة، والسالب رصيد مسترد. راجع الأرقام مع محاسبك قبل تقديم الإقرار."),
]


# --------------------------------------------------------------------------
# layout
# --------------------------------------------------------------------------
def entry(key):
    return next(r for r in REPORTS if r.key == key)


def title_block(m: ReportModel, title: str, w: int, with_band: bool) -> int:
    P = SEC_PAGE_HEADER
    half = w // 2
    txt(m, P, "txtStoreName", '=Nz(SettingValue("StoreName"),"")', 0, cm(0.05), half, cm(0.6), 11, True)
    txt(m, P, "txtStoreVat", '=IIf(Len(Nz(SettingValue("VATNumber"),""))>0,"الرقم الضريبي: " & '
                             'SettingValue("VATNumber"),"")', half, cm(0.05), w - half, cm(0.6), 9, align=1)
    lbl(m, P, "lblTitle", title, 0, cm(0.7), w, cm(0.85), 16, True, align=2)
    txt(m, P, "txtCriteria", "=ReportCriteria()", 0, cm(1.6), w, cm(0.5), 10, align=2)
    return cm(2.2)


def page_footer(m: ReportModel, w: int):
    F = SEC_PAGE_FOOTER
    split = w * 6 // 10
    txt(m, F, "txtPrinted", "=ReportPrintedAt()", 0, cm(0.1), split, cm(0.45), 8)
    txt(m, F, "txtPage", '="صفحة " & [Page] & " من " & [Pages]', split, cm(0.1), w - split, cm(0.45), 8,
        align=1)


def list_report(spec: ListSpec) -> ReportModel:
    e = entry(spec.key)
    w = cm(LANDSCAPE_W if spec.landscape else PORTRAIT_W)
    summary = spec.summary or []
    m = ReportModel(e.report, e.title, w,
                    {SEC_PAGE_HEADER: cm(2.95), SEC_DETAIL: cm(0.56),
                     SEC_RPT_FOOTER: cm(0.8) + cm(0.65) * len(summary) + (cm(0.2) if summary else 0),
                     SEC_PAGE_FOOTER: cm(0.6)},
                    record_source=e.query, group="", sorts=list(spec.sorts), landscape=spec.landscape,
                    page_setup=True, no_data=spec.no_data)
    y = title_block(m, e.title, w, True)
    m.add(SEC_PAGE_HEADER, Control("rect", "boxColumns", 0, y, w, cm(0.65),
                                   {"BackStyle": 1, "BackColor": Sym("CLR_SECONDARY")}, decorative=True))
    x, first_total = 0, None
    for i, c in enumerate(spec.cols):
        cw = cm(c.width) if i < len(spec.cols) - 1 else w - x
        lbl(m, SEC_PAGE_HEADER, f"lblCol{i + 1}", c.caption, x, y + cm(0.08), cw, cm(0.5), 8, True, align=2)
        t = txt(m, SEC_DETAIL, f"txtCol{i + 1}", c.source, x, cm(0.03), cw, cm(0.5), 8,
                align=2 if c.fmt else 0, fmt=c.fmt, grow=c.grow)
        if c.running:
            t.props["RunningSum"] = 2
        if c.total:
            if first_total is None:
                first_total = x
            src = c.source[1:] if c.source.startswith("=") else f"[{c.source}]"
            txt(m, SEC_RPT_FOOTER, f"txtTotal{i + 1}", f"=Sum({src})", x, cm(0.15), cw, cm(0.5), 8, True,
                align=2, fmt=c.fmt)
        x += cw
    assert x == w, spec.key
    hline(m, SEC_RPT_FOOTER, "lnTotals", cm(0.05), w)
    count_w = first_total if first_total else w
    txt(m, SEC_RPT_FOOTER, "txtCount", '="الإجمالي (" & Count(*) & " سجل)"', 0, cm(0.15), count_w, cm(0.5),
        8, True)
    for i, (caption, expr) in enumerate(summary):
        y = cm(0.95) + i * cm(0.65)
        lbl(m, SEC_RPT_FOOTER, f"lblSum{i + 1}", caption, w - cm(10.0), y, cm(5.5), cm(0.55), 11, True)
        txt(m, SEC_RPT_FOOTER, f"txtSum{i + 1}", expr, w - cm(4.5), y, cm(4.5), cm(0.55), 11, True, align=2,
            fmt=MONEY)
    page_footer(m, w)
    m.events.append('SetSecProp 0, "AlternateBackColor", 15921906')     # light grey rows
    return m


def card_report(spec: CardSpec) -> ReportModel:
    e = entry(spec.key)
    w = cm(PORTRAIT_W)
    row_h = cm(0.75)
    heights = {SEC_PAGE_HEADER: cm(2.3), SEC_DETAIL: row_h * len(spec.rows) + cm(0.2),
               SEC_PAGE_FOOTER: cm(0.6)}
    if spec.note:
        heights[SEC_RPT_FOOTER] = cm(1.1)
    m = ReportModel(e.report, e.title, w, heights, record_source=e.query, group="", sorts=[],
                    page_setup=True, no_data=NO_DATA)
    title_block(m, e.title, w, False)
    for i, (caption, source, fmt, bold) in enumerate(spec.rows):
        y = cm(0.1) + i * row_h
        lbl(m, SEC_DETAIL, f"lblRow{i + 1}", caption, cm(2.0), y, cm(9.5), cm(0.6), 11, bold)
        txt(m, SEC_DETAIL, f"txtRow{i + 1}", source, cm(11.6), y, cm(5.4), cm(0.6), 12 if bold else 11, bold,
            align=2, fmt=fmt)
    if spec.note:
        lbl(m, SEC_RPT_FOOTER, "lblNote", spec.note, 0, cm(0.2), w, cm(0.8), 9, align=2)
    page_footer(m, w)
    return m


def catalog_reports() -> List[ReportModel]:
    return [list_report(s) for s in LIST_SPECS] + [card_report(s) for s in CARD_SPECS]
