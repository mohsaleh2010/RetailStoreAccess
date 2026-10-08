"""Phase 7 screens: purchase invoice, purchase return, payment voucher, purchase
invoice view, inventory (balances + manual movements) and stocktaking.
Layout conventions are the same as forms.py (twips, x from the start edge)."""

from typing import List, Tuple

from forms import Control, FormModel, Sym, button, cm, currency_pair, fit_window, labelled, shrink_area, title_band, CATEGORY_ROWS
from forms_sales import LOCKED, PAYMENT_ROWS, PAYMENT_TYPES, grid_row, header_labels

PURCHASE_PRODUCT_ROWS = ("SELECT ProductID, ProductName & ' - ' & ProductCode AS Item, PurchasePrice "
                         "FROM Products WHERE IsActive = True ORDER BY ProductName")
ACTIVE_SUPPLIER_ROWS = ("SELECT SupplierID, SupplierName FROM Suppliers WHERE IsActive = True "
                        "ORDER BY SupplierName")
PURCHASE_REFUND_TYPES = "CREDIT;خصم من رصيد المورد;CASH;استرداد نقدي من المورد"
MANUAL_TYPE_ROWS = ("SELECT t.TransactionTypeID, t.TypeName FROM [@TransactionTypes] AS t WHERE t.IsManual = True "
                    "ORDER BY t.TransactionTypeID")
COUNT_ROWS = ("SELECT StockCountID, CountNumber, CountDate, IIf(Status = 'OPEN', 'مفتوح', "
              "IIf(Status = 'POSTED', 'مُرحّل', 'ملغى')) AS StatusName FROM StockCounts "
              "ORDER BY StockCountID DESC")
COUNT_LINES_SQL = ("SELECT d.StockCountDetailID, d.StockCountID, d.ProductID, d.SystemQuantity, "
                   "d.ActualQuantity, d.Difference, d.UnitCost, d.DifferenceValue, d.Notes, "
                   "p.ProductCode, p.ProductName, p.Barcode FROM StockCountDetails AS d "
                   "INNER JOIN Products AS p ON d.ProductID = p.ProductID")


def check_with_label(m: FormModel, name: str, caption: str, x: int, y: int, w: int, events):
    """Check box with its caption beside it (labelled() would put it above, too narrow)."""
    c = m.add(Control("check", name, x, y + cm(0.15), cm(0.5), cm(0.5), {}, events=events))
    m.add(Control("label", "lbl" + name[3:], x + cm(0.65), y, w - cm(0.65), cm(0.8),
                  {"Caption": caption, "FontSize": 10, "ForeColor": Sym("CLR_TEXT")}, parent=name))
    return c


# ------------------------------------------------------------------ purchase invoice
PURCHASE_TITLES = ["#", "الصنف", "الكمية", "تكلفة الوحدة", "الخصم", "الإجمالي", "سعر البيع",
                   "سعر بيع جديد", ""]


def layout_purchase_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.75)
    m = FormModel("frmPurchaseLines", "أسطر فاتورة الشراء", cm(21.2), row_h, popup=False,
                  record_source="SELECT * FROM tmpPurchaseLines ORDER BY LineNo", allow_add=False)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0}
    heads = grid_row(m, [
        ("LineNo", "LineNo", 0.9, dict(LOCKED), []),
        ("ProductName", "ProductName", 5.6, dict(LOCKED), []),
        ("Quantity", "Quantity", 1.9, {"Format": "#,##0.###"}, ["AfterUpdate"]),
        ("UnitCost", "UnitCost", 2.3, {"Format": "#,##0.00##"}, ["AfterUpdate"]),
        ("LineDiscount", "LineDiscount", 1.9, {"Format": "#,##0.00"}, ["AfterUpdate"]),
        ("LineTotal", "LineTotal", 2.5, {**LOCKED, "Format": "#,##0.00", "FontBold": True}, []),
        ("SellingPrice", "SellingPrice", 2.2, {**LOCKED, "Format": "#,##0.00"}, []),
        ("NewSellingPrice", "NewSellingPrice", 2.4, {"Format": "#,##0.00"}, ["AfterUpdate"]),
        ("btnRemove", "", 0.8, {"_kind": "button", "Caption": Sym(f"ChrW(&H{0xE74D:X})"),
                               "Style": "danger", "FontName": Sym("ICON_FONT")}, ["Click"]),
    ], row_h)
    m.code = []
    for f in ("Quantity", "UnitCost", "LineDiscount", "NewSellingPrice"):
        m.code += [f"Private Sub {f}_AfterUpdate()", f'    PurLineChanged Me, "{f}"', "End Sub"]
    m.code += ["Private Sub btnRemove_Click()", "    PurRemoveLine Me", "End Sub"]
    return m, heads


def layout_purchase_invoice(line_heads) -> FormModel:
    width, height = cm(33.5), cm(16.3)
    m = FormModel("frmPurchaseInvoice", "فاتورة مشتريات", width, height, popup=False, allow_add=False)
    title_band(m, "فاتورة مشتريات",
               "التكلفة بدون ضريبة  |  F9 حفظ  |  F5 فاتورة جديدة  |  F4 بحث بالاسم  |  F2 الباركود",
               "purchases")
    y = cm(2.3)
    bc = m.add(Control("text", "txtBarcode", cm(0.4), y, cm(8.0), cm(1.0), {"FontSize": 16},
                       events=["KeyDown"]))
    labelled(m, "txtBarcode", "الباركود أو كود المنتج (Enter)", bc)
    q = m.add(Control("text", "txtQty", cm(8.7), y, cm(2.0), cm(1.0),
                      {"FontSize": 16, "Format": "#,##0.###", "DefaultValue": "1"}))
    labelled(m, "txtQty", "الكمية", q)
    p = m.add(Control("combo", "cboProduct", cm(11.0), y, cm(10.6), cm(1.0),
                      {"RowSource": PURCHASE_PRODUCT_ROWS, "ColumnCount": 3, "ColumnWidths": "0;8;2",
                       "FontSize": 13}, events=["AfterUpdate"]))
    labelled(m, "cboProduct", "أو ابحث باسم المنتج (F4)", p)

    header_labels(m, cm(0.4), cm(3.6), PURCHASE_TITLES, line_heads)
    m.add(Control("subform", "subLines", cm(0.4), cm(4.25), cm(21.2), cm(9.3),
                  {"SourceObject": "frmPurchaseLines"}))
    m.add(Control("label", "lblStatus", cm(0.4), cm(13.75), cm(21.2), cm(0.7),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnSave", "حفظ (F9)", "primary", 3.2, "SavePurchase Me"),
            ("btnNewInvoice", "فاتورة جديدة (F5)", "secondary", 3.8, "NewPurchase Me"),
            ("btnReturn", "مرتجع مشتريات", "secondary", 3.4, 'OpenScreen "frmPurchaseReturn", 7'),
            ("btnPayment", "سند صرف", "secondary", 2.6,
             'OpenScreen "frmSupplierPayment", 7, Me!cboSupplier.Value'),
            ("btnNewProduct", "منتج جديد", "secondary", 2.8, 'OpenScreen "frmProducts"'),
            ("btnLastInvoice", "آخر فاتورة", "secondary", 2.8, "PurShowLast Me")]:
        button(m, name, caption, bx, cm(14.9), style, w=cm(w), h=cm(1.1), call=call)
        bx += cm(w) + cm(0.2)

    px, pw, half = cm(22.0), cm(11.1), cm(5.35)
    x2 = px + pw - half
    c = m.add(Control("combo", "cboSupplier", px, cm(2.3), pw, cm(0.8),
                      {"RowSource": ACTIVE_SUPPLIER_ROWS, "ColumnCount": 2, "ColumnWidths": "0;10"},
                      events=["AfterUpdate"]))
    labelled(m, "cboSupplier", "المورد *", c)
    m.add(Control("label", "lblSupplierInfo", px, cm(3.15), pw, cm(0.5),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    c = m.add(Control("text", "txtSupplierInvoiceNo", px, cm(4.3), half, cm(0.8), {}))
    labelled(m, "txtSupplierInvoiceNo", "رقم فاتورة المورد", c)
    c = m.add(Control("text", "txtInvoiceDate", x2, cm(4.3), half, cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtInvoiceDate", "تاريخ الفاتورة", c)
    c = m.add(Control("combo", "cboPaymentType", px, cm(5.75), half, cm(0.8),
                      {"RowSource": PAYMENT_TYPES, "ColumnCount": 2, "ColumnWidths": "0;5"},
                      events=["AfterUpdate"]))
    labelled(m, "cboPaymentType", "نوع الشراء", c)
    c = m.add(Control("combo", "cboPaymentMethod", x2, cm(5.75), half, cm(0.8),
                      {"RowSource": PAYMENT_ROWS, "ColumnCount": 2, "ColumnWidths": "0;5"}))
    labelled(m, "cboPaymentMethod", "طريقة الدفع", c)
    c = m.add(Control("text", "txtInvoiceDiscount", px, cm(7.2), half, cm(0.8),
                      {"Format": "#,##0.00"}, events=["AfterUpdate"]))
    labelled(m, "txtInvoiceDiscount", "خصم على الفاتورة (بدون ضريبة)", c)
    check_with_label(m, "chkChargeVAT", "المورد يحتسب الضريبة", x2, cm(7.2), half, ["AfterUpdate"])
    c = m.add(Control("text", "txtNotes", px, cm(8.65), half, cm(0.8), {}))
    labelled(m, "txtNotes", "ملاحظات", c)
    currency_pair(m, x2, cm(8.65), cbo_w=cm(2.6), rate_w=half - cm(2.8), call="PurCurrencyPicked Me")

    m.add(Control("rect", "boxTotals", px, cm(9.6), pw, cm(3.75), {"BackColor": Sym("CLR_SURFACE")},
                  decorative=True))
    ty = cm(9.7)
    for name, caption in [("SubTotal", "المجموع قبل الخصم والضريبة"), ("Discount", "الخصم"),
                          ("Tax", "ضريبة المدخلات")]:
        m.add(Control("label", f"lblCap{name}", px + cm(0.3), ty, cm(6.5), cm(0.6),
                      {"Caption": caption, "FontSize": 11, "ForeColor": Sym("CLR_MUTED")}))
        m.add(Control("label", f"lbl{name}", px + cm(6.9), ty, cm(3.9), cm(0.6),
                      {"Caption": "0.00", "FontSize": 12, "FontBold": True, "TextAlign": 3,
                       "ForeColor": Sym("CLR_TEXT")}))
        ty += cm(0.6)
    m.add(Control("label", "lblCapTotal", px + cm(0.3), cm(11.65), cm(5.0), cm(0.6),
                  {"Caption": "الإجمالي شامل الضريبة", "FontSize": 12, "FontBold": True,
                   "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("label", "lblTotal", px + cm(5.4), cm(11.5), cm(5.4), cm(1.0),
                  {"Caption": "0.00", "FontSize": 24, "FontBold": True, "TextAlign": 3,
                   "ForeColor": Sym("CLR_ACCENT")}))
    m.add(Control("label", "lblItems", px + cm(0.3), cm(12.65), cm(10.5), cm(0.6),
                  {"Caption": " ", "FontSize": 10, "TextAlign": 2, "ForeColor": Sym("CLR_MUTED")}))
    c = m.add(Control("text", "txtPaid", px, cm(14.0), half, cm(0.9),
                      {"FontSize": 13, "Format": "#,##0.00"}, events=["AfterUpdate"]))
    labelled(m, "txtPaid", "المدفوع للمورد الآن", c)
    m.add(Control("label", "lblRemaining", x2, cm(14.05), half, cm(0.8),
                  {"Caption": " ", "FontSize": 12, "FontBold": True, "ForeColor": Sym("CLR_WARNING")}))
    button(m, "btnLabels", "طباعة الباركود", px, cm(14.9), "secondary", w=cm(4.0), h=cm(1.1),
           call="LabelsFromPurchaseScreen Me")
    button(m, "btnClose", "إغلاق", px + pw - cm(3.0), cm(14.9), "secondary", w=cm(3.0), h=cm(1.1),
           call="DoCmd.Close acForm, Me.Name")
    fit_window(m, split_x=cm(21.8), bottom_y=cm(13.4), stretch_w=("subLines", "lblStatus", "cboProduct"),
               stretch_h=("subLines",))

    m.form_events = ["Load", "KeyDown", "Unload"]
    m.code = ([
        "Private Sub Form_Load()", "    PurchaseLoad Me", "End Sub",
        "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)",
        "    PurchaseKeyDown Me, KeyCode, Shift", "End Sub",
        "Private Sub Form_Unload(Cancel As Integer)", "    Cancel = Not PurchaseUnload(Me)", "End Sub",
        "Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)",
        "    PurBarcodeKeyDown Me, KeyCode", "End Sub",
        "Private Sub cboProduct_AfterUpdate()", "    PurProductPicked Me", "End Sub",
        "Private Sub cboSupplier_AfterUpdate()", "    PurSupplierChanged Me", "End Sub",
        "Private Sub cboPaymentType_AfterUpdate()", "    PurPaymentTypeChanged Me", "End Sub",
        "Private Sub txtInvoiceDiscount_AfterUpdate()", "    RecalcPurchase Me", "End Sub",
        "Private Sub chkChargeVAT_AfterUpdate()", "    RecalcPurchase Me", "End Sub",
        "Private Sub txtPaid_AfterUpdate()", "    RecalcPurchase Me", "End Sub",
    ] + m.code)
    return m


# ------------------------------------------------------------------ purchase return
PURCHASE_RETURN_TITLES = ["الصنف", "الكمية المشتراة", "مرتجع سابقًا", "المتاح للإرجاع", "الرصيد الحالي",
                          "الكمية المرتجعة", "قيمة المرتجع"]


def layout_purchase_return_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.75)
    m = FormModel("frmPurchaseReturnLines", "أسطر مرتجع المشتريات", cm(26.2), row_h, popup=False,
                  record_source="SELECT * FROM tmpPurchaseReturnLines ORDER BY PurchaseDetailID",
                  allow_add=False)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0}
    heads = grid_row(m, [
        ("ProductName", "ProductName", 8.0, dict(LOCKED), []),
        ("BoughtQty", "BoughtQty", 2.5, {**LOCKED, "Format": "#,##0.###"}, []),
        ("ReturnedQty", "ReturnedQty", 2.5, {**LOCKED, "Format": "#,##0.###"}, []),
        ("AvailableQty", "AvailableQty", 2.5, {**LOCKED, "Format": "#,##0.###"}, []),
        ("InStock", "InStock", 2.5, {**LOCKED, "Format": "#,##0.###"}, []),
        ("ReturnQty", "ReturnQty", 2.8, {"Format": "#,##0.###", "FontBold": True}, ["AfterUpdate"]),
        ("ReturnAmount", "ReturnAmount", 3.4, {**LOCKED, "Format": "#,##0.00", "FontBold": True}, []),
    ], row_h)
    m.code = ["Private Sub ReturnQty_AfterUpdate()", "    PurReturnLineChanged Me", "End Sub"]
    return m, heads


def layout_purchase_return(line_heads) -> FormModel:
    width, height = cm(27.0), cm(17.0)
    m = FormModel("frmPurchaseReturn", "مرتجع مشتريات", width, height, popup=True, allow_add=False)
    title_band(m, "مرتجع مشتريات", "اختر فاتورة الشراء ثم حدد الكميات التي تعود للمورد", "purchases")
    m.add(Control("text", "txtInvoiceID", width - cm(1.0), cm(0.1), cm(0.6), cm(0.4),
                  {"Visible": False}, decorative=True))
    c = m.add(Control("text", "txtInvoiceNo", cm(0.4), cm(2.3), cm(5.0), cm(0.8), {"FontSize": 12},
                      events=["AfterUpdate"]))
    labelled(m, "txtInvoiceNo", "رقمنا (PUR-) أو رقم فاتورة المورد", c)
    button(m, "btnFind", "بحث", cm(5.6), cm(2.3), "secondary", w=cm(2.0), h=cm(0.8),
           call="PurReturnFind Me")
    m.add(Control("label", "lblInvoiceInfo", cm(7.9), cm(2.4), width - cm(8.3), cm(0.7),
                  {"Caption": " ", "FontSize": 10, "ForeColor": Sym("CLR_TEXT")}))
    header_labels(m, cm(0.4), cm(3.45), PURCHASE_RETURN_TITLES, line_heads)
    m.add(Control("subform", "subReturnLines", cm(0.4), cm(4.05), cm(26.2), cm(7.6),
                  {"SourceObject": "frmPurchaseReturnLines"}))
    c = m.add(Control("text", "txtReason", cm(0.4), cm(12.35), cm(11.6), cm(0.8), {}))
    labelled(m, "txtReason", "سبب الإرجاع *", c)
    c = m.add(Control("combo", "cboRefundType", cm(12.3), cm(12.35), cm(6.6), cm(0.8),
                      {"RowSource": PURCHASE_REFUND_TYPES, "ColumnCount": 2, "ColumnWidths": "0;6"}))
    labelled(m, "cboRefundType", "طريقة الاسترداد", c)
    c = m.add(Control("combo", "cboPaymentMethod", cm(19.2), cm(12.35), cm(7.4), cm(0.8),
                      {"RowSource": PAYMENT_ROWS, "ColumnCount": 2, "ColumnWidths": "0;6"}))
    labelled(m, "cboPaymentMethod", "طريقة الاستلام", c)
    m.add(Control("label", "lblReturnTotalCap", cm(0.4), cm(13.6), cm(4.0), cm(0.8),
                  {"Caption": "قيمة المرتجع:", "FontSize": 13, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("label", "lblReturnTotal", cm(4.5), cm(13.4), cm(6.0), cm(1.1),
                  {"Caption": "0.00", "FontSize": 20, "FontBold": True, "ForeColor": Sym("CLR_DANGER")}))
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnReturnAll", "إرجاع الكل", "secondary", 3.2, "PurReturnAll Me"),
            ("btnSaveReturn", "حفظ المرتجع", "primary", 3.6, "SavePurchaseReturn Me"),
            ("btnSaveReturnPrint", "حفظ وطباعة", "primary", 3.6, "SavePurchaseReturn Me, True")]:
        button(m, name, caption, bx, cm(15.3), style, w=cm(w), h=cm(1.0), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(15.3), "secondary", w=cm(2.6),
           h=cm(1.0), call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    PurReturnLoad Me", "End Sub",
               "Private Sub txtInvoiceNo_AfterUpdate()", "    PurReturnFind Me", "End Sub"]
              + m.code)
    return m


# ------------------------------------------------------------------ payment voucher
def layout_supplier_payment() -> FormModel:
    width, height = cm(16.0), cm(10.2)
    m = FormModel("frmSupplierPayment", "سند صرف", width, height, popup=True, allow_add=False)
    title_band(m, "سند صرف", "تسجيل دفعة لمورد من حسابه", "suppliers")
    c = m.add(Control("combo", "cboSupplier", cm(0.4), cm(2.4), width - cm(0.8), cm(0.8),
                      {"RowSource": ACTIVE_SUPPLIER_ROWS, "ColumnCount": 2, "ColumnWidths": "0;10"},
                      events=["AfterUpdate"]))
    labelled(m, "cboSupplier", "المورد", c)
    m.add(Control("label", "lblBalance", cm(0.4), cm(3.3), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    c = m.add(Control("text", "txtAmount", cm(0.4), cm(4.6), cm(4.4), cm(0.9),
                      {"FontSize": 14, "Format": "#,##0.00"}))
    labelled(m, "txtAmount", "المبلغ *", c)
    currency_pair(m, cm(5.0), cm(4.6), h=cm(0.9), cbo_w=cm(2.6), rate_w=cm(2.4))
    c = m.add(Control("combo", "cboPaymentMethod", cm(10.4), cm(4.6), width - cm(10.8), cm(0.9),
                      {"RowSource": PAYMENT_ROWS, "ColumnCount": 2, "ColumnWidths": "0;6"}))
    labelled(m, "cboPaymentMethod", "طريقة الدفع", c)
    c = m.add(Control("text", "txtNotes", cm(0.4), cm(6.2), width - cm(0.8), cm(0.8), {}))
    labelled(m, "txtNotes", "ملاحظات", c)
    button(m, "btnSave", "حفظ السند", cm(0.4), cm(8.4), "primary", w=cm(3.4), h=cm(1.0),
           call="SaveSupplierPayment Me")
    button(m, "btnSavePrint", "حفظ وطباعة", cm(4.0), cm(8.4), "primary", w=cm(3.4), h=cm(1.0),
           call="SaveSupplierPayment Me, True")
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(8.4), "secondary", w=cm(2.6),
           h=cm(1.0), call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    SupplierPaymentLoad Me", "End Sub",
               "Private Sub cboSupplier_AfterUpdate()", "    SupplierPaymentChanged Me", "End Sub"]
              + m.code)
    return m


# ------------------------------------------------------------------ purchase invoice view
def layout_purchase_view() -> FormModel:
    width, height = cm(27.0), cm(15.6)
    m = FormModel("frmPurchaseView", "فاتورة شراء", width, height, popup=True, allow_add=False)
    title_band(m, "فاتورة شراء", "عرض فقط - التصحيح يكون بمرتجع مشتريات", "purchases")
    m.add(Control("text", "txtInvoiceID", width - cm(1.0), cm(0.1), cm(0.6), cm(0.4),
                  {"Visible": False}, decorative=True))
    m.add(Control("text", "txtSupplierID", width - cm(1.7), cm(0.1), cm(0.6), cm(0.4),
                  {"Visible": False}, decorative=True))
    m.add(Control("label", "lblHeader", cm(0.4), cm(1.8), width - cm(0.8), cm(0.7),
                  {"Caption": " ", "FontSize": 10, "ForeColor": Sym("CLR_TEXT")}))
    m.add(Control("list", "lstLines", cm(0.4), cm(2.7), width - cm(0.8), cm(6.6),
                  {"ColumnCount": 7, "ColumnWidths": "1;9;2.5;3.5;2.5;3;4", "ColumnHeads": True}))
    m.add(Control("label", "lblReturnsCap", cm(0.4), cm(9.5), cm(12), cm(0.55),
                  {"Caption": "المرتجعات على هذه الفاتورة (نقر مزدوج للطباعة)", "FontSize": 10,
                   "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstReturns", cm(0.4), cm(10.1), width - cm(0.8), cm(2.3),
                  {"ColumnCount": 5, "ColumnWidths": "0;4;5;3.5;12", "ColumnHeads": True},
                  events=["DblClick"]))
    m.add(Control("label", "lblTotals", cm(0.4), cm(12.6), width - cm(0.8), cm(0.7),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnPrint", "طباعة", "primary", 2.6, 'PrintPurchaseDocument "PURCHASE", Me!txtInvoiceID.Value'),
            ("btnReturn", "مرتجع", "secondary", 2.6,
             'OpenScreen "frmPurchaseReturn", 7, Me!txtInvoiceID.Value'),
            ("btnPayment", "سند صرف", "secondary", 2.8,
             'OpenScreen "frmSupplierPayment", 7, Me!txtSupplierID.Value'),
            ("btnLabels", "طباعة الباركود", "secondary", 3.4, "LabelsForPurchase Me!txtInvoiceID.Value, Me")]:
        button(m, name, caption, bx, cm(13.9), style, w=cm(w), h=cm(1.0), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(13.9), "secondary", w=cm(2.6),
           h=cm(1.0), call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    PurchaseViewLoad Me", "End Sub",
               "Private Sub lstReturns_DblClick(Cancel As Integer)",
               '    PrintPurchaseDocument "RETURN", Me!lstReturns.Value', "End Sub"] + m.code)
    return m


# ------------------------------------------------------------------ inventory
def layout_inventory() -> FormModel:
    width, height = cm(33.5), cm(18.6)
    m = FormModel("frmInventory", "المخزون", width, height, popup=False, allow_add=False)
    title_band(m, "المخزون", "أرصدة المنتجات وحركاتها، والرصيد الافتتاحي والإضافة والخصم اليدوي",
               "inventory")
    y = cm(2.3)
    c = m.add(Control("text", "txtSearch", cm(0.4), y, cm(7.0), cm(0.8), {}, events=["AfterUpdate"]))
    labelled(m, "txtSearch", "بحث بالاسم أو الكود أو الباركود", c)
    c = m.add(Control("combo", "cboCategory", cm(7.7), y, cm(4.6), cm(0.8),
                      {"RowSource": CATEGORY_ROWS, "ColumnCount": 2, "ColumnWidths": "0;4.5"},
                      events=["AfterUpdate"]))
    labelled(m, "cboCategory", "التصنيف", c)
    check_with_label(m, "chkLowOnly", "منخفضة المخزون فقط", cm(12.6), y, cm(4.6), ["AfterUpdate"])
    button(m, "btnRefresh", "تحديث", cm(17.4), y - cm(0.03), "secondary", w=cm(2.0),
           call="InventoryRefresh Me")
    m.add(Control("list", "lstProducts", cm(0.4), cm(3.5), cm(20.6), cm(12.0),
                  {"ColumnCount": 8, "ColumnWidths": "0;2.4;6.5;3;2;2;2.2;2.3", "ColumnHeads": True},
                  events=["AfterUpdate", "DblClick"]))
    m.add(Control("label", "lblInvTotals", cm(0.4), cm(15.7), cm(20.6), cm(0.6),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnStockCount", "الجرد", "secondary", 2.6, 'OpenScreen "frmStockCount", 7'),
            ("btnPurchase", "فاتورة مشتريات", "secondary", 3.6, 'OpenScreen "frmPurchaseInvoice", 7'),
            ("btnLowReport", "تقرير النواقص", "secondary", 3.4,
             'OpenReportOrQuery "rptLowStock", "LowStockQuery", ""'),
            ("btnStockReport", "تقرير المخزون", "secondary", 3.4,
             'OpenReportOrQuery "rptStockBalance", "StockBalanceQuery", ""'),
            ("btnLabels", "طباعة باركود", "secondary", 3.4,
             'OpenScreen "frmBarcodeLabels", 0, Me!lstProducts.Value')]:
        button(m, name, caption, bx, cm(16.6), style, w=cm(w), h=cm(1.1), call=call)
        bx += cm(w) + cm(0.2)

    px, pw = cm(21.4), cm(11.7)
    m.add(Control("label", "lblProductName", px, cm(2.3), pw, cm(0.8),
                  {"Caption": "اختر منتجًا من القائمة", "FontSize": 14, "FontBold": True,
                   "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("label", "lblProductStock", px, cm(3.15), pw, cm(0.6),
                  {"Caption": " ", "FontSize": 10, "ForeColor": Sym("CLR_TEXT")}))
    m.add(Control("label", "lblManualCap", px, cm(4.0), pw, cm(0.6),
                  {"Caption": "حركة يدوية على المنتج المحدد", "FontSize": 11, "FontBold": True,
                   "ForeColor": Sym("CLR_MUTED")}))
    c = m.add(Control("combo", "cboMoveType", px, cm(5.25), cm(5.6), cm(0.8),
                      {"RowSource": MANUAL_TYPE_ROWS, "ColumnCount": 2, "ColumnWidths": "0;5"},
                      events=["AfterUpdate"]))
    labelled(m, "cboMoveType", "نوع الحركة", c)
    c = m.add(Control("text", "txtMoveQty", px + cm(5.9), cm(5.25), cm(2.7), cm(0.8),
                      {"Format": "#,##0.###"}))
    labelled(m, "txtMoveQty", "الكمية", c)
    c = m.add(Control("text", "txtMoveCost", px + cm(8.9), cm(5.25), cm(2.8), cm(0.8),
                      {"Format": "#,##0.00##"}))
    labelled(m, "txtMoveCost", "تكلفة الوحدة", c)
    c = m.add(Control("text", "txtMoveNotes", px, cm(6.7), cm(8.6), cm(0.8), {}))
    labelled(m, "txtMoveNotes", "السبب / ملاحظات", c)
    button(m, "btnPostMove", "حفظ الحركة", px + cm(8.9), cm(6.67), "primary", w=cm(2.8),
           call="PostInventoryMove Me")
    m.add(Control("label", "lblMovesCap", px, cm(7.85), pw, cm(0.6),
                  {"Caption": "آخر حركات المنتج", "FontSize": 11, "FontBold": True,
                   "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstMoves", px, cm(8.5), pw, cm(7.8),
                  {"ColumnCount": 5, "ColumnWidths": "2.6;2.6;1.8;2;2.4", "ColumnHeads": True}))
    button(m, "btnClose", "إغلاق", px + pw - cm(3.0), cm(16.6), "secondary", w=cm(3.0), h=cm(1.1),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    InventoryLoad Me", "End Sub",
               "Private Sub txtSearch_AfterUpdate()", "    InventoryRefresh Me", "End Sub",
               "Private Sub cboCategory_AfterUpdate()", "    InventoryRefresh Me", "End Sub",
               "Private Sub chkLowOnly_AfterUpdate()", "    InventoryRefresh Me", "End Sub",
               "Private Sub lstProducts_AfterUpdate()", "    InventoryProductPicked Me", "End Sub",
               "Private Sub lstProducts_DblClick(Cancel As Integer)",
               '    OpenScreen "frmProducts", 0, Me!lstProducts.Value', "End Sub",
               "Private Sub cboMoveType_AfterUpdate()", "    InventoryMoveTypeChanged Me", "End Sub"]
              + m.code)
    shrink_area(m, ("lstProducts", "lstMoves"), cm(2.1))
    fit_window(m, split_x=cm(21.2), bottom_y=cm(13.4), stretch_w=("lstProducts", "lblInvTotals"),
               stretch_h=("lstProducts", "lstMoves"))
    return m


# ------------------------------------------------------------------ stocktaking
COUNT_TITLES = ["الكود", "المنتج", "الكمية المسجلة", "الكمية الفعلية", "الفرق", "قيمة الفرق", "ملاحظات"]


def layout_count_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.75)
    m = FormModel("frmStockCountLines", "أسطر الجرد", cm(32.7), row_h, popup=False,
                  record_source=COUNT_LINES_SQL + " WHERE d.StockCountID = 0 ORDER BY p.ProductName",
                  allow_add=False)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0}
    heads = grid_row(m, [
        ("ProductCode", "ProductCode", 2.6, dict(LOCKED), []),
        ("ProductName", "ProductName", 8.6, dict(LOCKED), []),
        ("SystemQuantity", "SystemQuantity", 3.0, {**LOCKED, "Format": "#,##0.###"}, []),
        ("ActualQuantity", "ActualQuantity", 3.0, {"Format": "#,##0.###", "FontBold": True},
         ["AfterUpdate"]),
        ("Difference", "Difference", 3.0, {**LOCKED, "Format": "#,##0.###"}, []),
        ("DifferenceValue", "DifferenceValue", 3.2, {**LOCKED, "Format": "#,##0.00"}, []),
        ("Notes", "Notes", 7.4, {}, []),
    ], row_h)
    m.code = ["Private Sub ActualQuantity_AfterUpdate()", "    CountLineChanged Me", "End Sub"]
    return m, heads


def layout_stock_count(line_heads) -> FormModel:
    width, height = cm(33.5), cm(18.6)
    m = FormModel("frmStockCount", "الجرد", width, height, popup=False, allow_add=False)
    title_band(m, "الجرد", "ابدأ جردًا، أدخل الكمية الفعلية أو امسح الباركود، ثم رحّل الفروقات",
               "stocktake")
    y = cm(2.3)
    c = m.add(Control("combo", "cboCount", cm(0.4), y, cm(9.0), cm(0.8),
                      {"RowSource": COUNT_ROWS, "ColumnCount": 4, "ColumnWidths": "0;3;3.5;2.3"},
                      events=["AfterUpdate"]))
    labelled(m, "cboCount", "جلسة الجرد (الأحدث أولًا)", c)
    c = m.add(Control("combo", "cboCategory", cm(9.7), y, cm(5.0), cm(0.8),
                      {"RowSource": CATEGORY_ROWS, "ColumnCount": 2, "ColumnWidths": "0;5"}))
    labelled(m, "cboCategory", "تصنيف الجرد الجديد (فارغ = الكل)", c)
    button(m, "btnNewCount", "جرد جديد", cm(15.0), y - cm(0.03), "primary", w=cm(3.0),
           call="NewStockCount Me")
    m.add(Control("label", "lblCountInfo", cm(18.3), y, width - cm(18.7), cm(0.8),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_TEXT")}))
    y = cm(3.85)
    c = m.add(Control("text", "txtCountBarcode", cm(0.4), y, cm(7.0), cm(0.8), {"FontSize": 12},
                      events=["KeyDown"]))
    labelled(m, "txtCountBarcode", "امسح الباركود (يضيف 1) أو 3*الكود", c)
    check_with_label(m, "chkDiffOnly", "الفروقات فقط", cm(7.7), y, cm(3.6), ["AfterUpdate"])
    button(m, "btnRefreshSystem", "تحديث الكميات المسجلة", cm(11.6), y - cm(0.03), "secondary", w=cm(4.6),
           call="CountRefreshSystem Me")
    header_labels(m, cm(0.4), cm(4.9), COUNT_TITLES, line_heads)
    m.add(Control("subform", "subCountLines", cm(0.4), cm(5.5), cm(32.7), cm(9.85),
                  {"SourceObject": "frmStockCountLines"}))
    m.add(Control("label", "lblCountSummary", cm(0.4), cm(15.55), cm(32.7), cm(0.7),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnPostCount", "ترحيل الجرد", "primary", 3.4, "PostCountScreen Me"),
            ("btnCancelCount", "إلغاء الجرد", "danger", 3.0, "CancelCountScreen Me"),
            ("btnCountReport", "طباعة الجرد", "secondary", 3.4, "PrintStockCount Me!cboCount.Value")]:
        button(m, name, caption, bx, cm(16.6), style, w=cm(w), h=cm(1.1), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(3.0), cm(16.6), "secondary", w=cm(3.0),
           h=cm(1.1), call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    StockCountLoad Me", "End Sub",
               "Private Sub cboCount_AfterUpdate()", "    StockCountPicked Me", "End Sub",
               "Private Sub txtCountBarcode_KeyDown(KeyCode As Integer, Shift As Integer)",
               "    CountBarcodeKeyDown Me, KeyCode", "End Sub",
               "Private Sub chkDiffOnly_AfterUpdate()", "    CountFilterChanged Me", "End Sub"]
              + m.code)
    shrink_area(m, ("subCountLines",), cm(2.1))
    fit_window(m, split_x=cm(29.0), bottom_y=cm(13.3),
               stretch_w=("subCountLines", "lblCountSummary", "lblCountInfo"), stretch_h=("subCountLines",))
    return m


def purchase_forms() -> List[FormModel]:
    lines, heads = layout_purchase_lines()
    ret_lines, ret_heads = layout_purchase_return_lines()
    count_lines, count_heads = layout_count_lines()
    # subforms first: the parent forms embed them by name
    return [lines, layout_purchase_invoice(heads), ret_lines, layout_purchase_return(ret_heads),
            layout_supplier_payment(), layout_purchase_view(), layout_inventory(),
            count_lines, layout_stock_count(count_heads)]
