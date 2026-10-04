"""Phase 6 screens: point of sale, sales return, receipt voucher, invoice view.
Layout conventions are the same as forms.py (twips, x from the start edge)."""

from typing import List, Tuple

from forms import (Control, FormModel, Sym, ICONS, button, cm, labelled, title_band,
                   CUSTOMER_ROWS)

PAYMENT_TYPES = "CASH;نقدي;CREDIT;آجل"
REFUND_TYPES = "CASH;رد نقدي;CREDIT;خصم من رصيد العميل"
PAYMENT_ROWS = "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder"
POS_PRODUCT_ROWS = ("SELECT ProductID, ProductName & ' - ' & ProductCode AS Item, SellingPrice "
                    "FROM Products WHERE IsActive = True ORDER BY ProductName")
CREDIT_CUSTOMER_ROWS = ("SELECT CustomerID, CustomerName FROM Customers WHERE IsSystem = False "
                        "AND IsActive = True ORDER BY CustomerName")


def grid_row(m: FormModel, columns: List[Tuple[str, str, float, dict, list]], row_h: int):
    """Continuous-form row: (name, field, width cm, props, events). Returns header positions."""
    x = cm(0.05)
    heads = []
    for name, fld, width, props, events in columns:
        w = cm(width)
        kind = props.pop("_kind", "text")
        if kind == "check":
            c = Control("check", name, x + (w - cm(0.5)) // 2, cm(0.12), cm(0.5), cm(0.5), props,
                        events=events, source=fld)
        elif kind == "button":
            c = Control("button", name, x, cm(0.03), w, row_h - cm(0.06), props, events=events)
        else:
            c = Control("text", name, x, 0, w, row_h, props, events=events, source=fld)
        m.add(c)
        heads.append((x, w))
        x += w + cm(0.05)
    return heads


def header_labels(m: FormModel, sub_x: int, y: int, titles: List[str], heads):
    for i, (title, (x, w)) in enumerate(zip(titles, heads)):
        if title:
            m.add(Control("label", f"lblCol{i + 1}", sub_x + x, y, w, cm(0.55),
                          {"Caption": title, "FontSize": 9, "FontBold": True,
                           "ForeColor": Sym("CLR_MUTED"), "TextAlign": 2}))


LOCKED = {"Locked": True, "TabStop": False}


# ------------------------------------------------------------------ cart lines
POS_COLUMNS_TITLES = ["#", "الصنف", "الكمية", "السعر", "خصم السطر", "الإجمالي", "المتوفر", ""]


def layout_pos_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.75)
    m = FormModel("frmPOSLines", "أسطر الفاتورة", cm(21.2), row_h, popup=False,
                  record_source="SELECT * FROM tmpPOSLines ORDER BY LineNo", allow_add=False)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0}
    heads = grid_row(m, [
        ("LineNo", "LineNo", 1.0, dict(LOCKED), []),
        ("ProductName", "ProductName", 7.5, dict(LOCKED), []),
        ("Quantity", "Quantity", 2.2, {"Format": "#,##0.###"}, ["AfterUpdate"]),
        ("UnitPrice", "UnitPrice", 2.6, {"Format": "#,##0.00"}, ["AfterUpdate"]),
        ("LineDiscount", "LineDiscount", 2.2, {"Format": "#,##0.00"}, ["AfterUpdate"]),
        ("LineTotal", "LineTotal", 2.8, {**LOCKED, "Format": "#,##0.00", "FontBold": True}, []),
        ("Available", "Available", 1.6, {**LOCKED, "Format": "#,##0.###"}, []),
        ("btnRemove", "", 0.8, {"_kind": "button", "Caption": Sym(f"ChrW(&H{0xE74D:X})"),
                               "Style": "danger", "FontName": Sym("ICON_FONT")}, ["Click"]),
    ], row_h)
    m.code = []
    for f in ("Quantity", "UnitPrice", "LineDiscount"):
        m.code += [f"Private Sub {f}_AfterUpdate()", f'    LineChanged Me, "{f}"', "End Sub"]
    m.code += ["Private Sub btnRemove_Click()", "    RemoveCurrentLine Me", "End Sub"]
    return m, heads


def layout_pos(line_heads) -> FormModel:
    width, height = cm(33.5), cm(18.6)
    m = FormModel("frmPOS", "نقطة البيع", width, height, popup=False, allow_add=False)
    title_band(m, "نقطة البيع",
               "F9 حفظ  |  F12 حفظ وطباعة  |  F5 فاتورة جديدة  |  F4 بحث بالاسم  |  F8 المبلغ المدفوع  |  F2 الباركود",
               "sales")
    y = cm(2.3)
    bc = m.add(Control("text", "txtBarcode", cm(0.4), y, cm(8.0), cm(1.0), {"FontSize": 16},
                       events=["KeyDown"]))
    labelled(m, "txtBarcode", "الباركود أو كود المنتج (Enter)", bc)
    q = m.add(Control("text", "txtQty", cm(8.7), y, cm(2.0), cm(1.0),
                      {"FontSize": 16, "Format": "#,##0.###", "DefaultValue": "1"}))
    labelled(m, "txtQty", "الكمية", q)
    p = m.add(Control("combo", "cboProduct", cm(11.0), y, cm(10.6), cm(1.0),
                      {"RowSource": POS_PRODUCT_ROWS, "ColumnCount": 3, "ColumnWidths": "0;8;2",
                       "FontSize": 13}, events=["AfterUpdate"]))
    labelled(m, "cboProduct", "أو ابحث باسم المنتج (F4)", p)

    sub_x, sub_y = cm(0.4), cm(4.25)
    header_labels(m, sub_x, cm(3.6), POS_COLUMNS_TITLES, line_heads)
    m.add(Control("subform", "subLines", sub_x, sub_y, cm(21.2), cm(11.0),
                  {"SourceObject": "frmPOSLines"}))
    m.add(Control("label", "lblStatus", cm(0.4), cm(15.45), cm(21.2), cm(0.7),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    bx, by, bh = cm(0.4), cm(16.6), cm(1.1)
    for name, caption, style, w, call in [
            ("btnSave", "حفظ (F9)", "primary", 3.4, "SavePOS Me, False"),
            ("btnSavePrint", "حفظ وطباعة (F12)", "primary", 4.2, "SavePOS Me, True"),
            ("btnNewSale", "فاتورة جديدة (F5)", "secondary", 3.8, "NewSale Me"),
            ("btnReturn", "مرتجع", "secondary", 2.6, 'OpenScreen "frmSalesReturn", 6'),
            ("btnPayment", "سند قبض", "secondary", 2.6,
             'OpenScreen "frmCustomerPayment", 6, Me!cboCustomer.Value'),
            ("btnReprint", "إعادة طباعة", "secondary", 2.8, "ReprintLast Me")]:
        button(m, name, caption, bx, by, style, w=cm(w), h=bh, call=call)
        bx += cm(w) + cm(0.2)

    px, pw = cm(22.0), cm(11.1)
    c = m.add(Control("combo", "cboCustomer", px, cm(2.3), pw, cm(0.8),
                      {"RowSource": CUSTOMER_ROWS, "ColumnCount": 2, "ColumnWidths": "0;10"},
                      events=["AfterUpdate"]))
    labelled(m, "cboCustomer", "العميل", c)
    m.add(Control("label", "lblCustomerInfo", px, cm(3.15), pw, cm(0.5),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    half = cm(5.35)
    c = m.add(Control("combo", "cboPaymentType", px, cm(4.3), half, cm(0.8),
                      {"RowSource": PAYMENT_TYPES, "ColumnCount": 2, "ColumnWidths": "0;5"},
                      events=["AfterUpdate"]))
    labelled(m, "cboPaymentType", "نوع البيع", c)
    c = m.add(Control("combo", "cboPaymentMethod", px + pw - half, cm(4.3), half, cm(0.8),
                      {"RowSource": PAYMENT_ROWS, "ColumnCount": 2, "ColumnWidths": "0;5"},
                      events=["AfterUpdate"]))
    labelled(m, "cboPaymentMethod", "طريقة الدفع", c)
    c = m.add(Control("text", "txtInvoiceDiscount", px, cm(5.75), half, cm(0.8),
                      {"Format": "#,##0.00"}, events=["AfterUpdate"]))
    labelled(m, "txtInvoiceDiscount", "خصم على الفاتورة", c)
    c = m.add(Control("text", "txtNotes", px + pw - half, cm(5.75), half, cm(0.8), {}))
    labelled(m, "txtNotes", "ملاحظات", c)

    m.add(Control("rect", "boxTotals", px, cm(6.85), pw, cm(5.8), {"BackColor": Sym("CLR_SURFACE")},
                  decorative=True))
    ty = cm(7.0)
    for name, caption in [("SubTotal", "المجموع قبل الخصم والضريبة"), ("Discount", "الخصم"),
                          ("Tax", "ضريبة القيمة المضافة")]:
        m.add(Control("label", f"lblCap{name}", px + cm(0.3), ty, cm(6.5), cm(0.6),
                      {"Caption": caption, "FontSize": 11, "ForeColor": Sym("CLR_MUTED")}))
        m.add(Control("label", f"lbl{name}", px + cm(6.9), ty, cm(3.9), cm(0.6),
                      {"Caption": "0.00", "FontSize": 12, "FontBold": True, "TextAlign": 3,
                       "ForeColor": Sym("CLR_TEXT")}))
        ty += cm(0.7)
    m.add(Control("label", "lblCapTotal", px + cm(0.3), cm(9.25), cm(10.5), cm(0.6),
                  {"Caption": "الإجمالي شامل الضريبة", "FontSize": 12, "FontBold": True,
                   "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("label", "lblTotal", px + cm(0.3), cm(9.9), cm(10.5), cm(1.7),
                  {"Caption": "0.00", "FontSize": 34, "FontBold": True, "TextAlign": 2,
                   "ForeColor": Sym("CLR_ACCENT")}))
    m.add(Control("label", "lblItems", px + cm(0.3), cm(11.75), cm(10.5), cm(0.6),
                  {"Caption": " ", "FontSize": 10, "TextAlign": 2, "ForeColor": Sym("CLR_MUTED")}))
    c = m.add(Control("text", "txtTendered", px, cm(13.4), pw, cm(1.0),
                      {"FontSize": 16, "Format": "#,##0.00"}, events=["AfterUpdate"]))
    labelled(m, "txtTendered", "المبلغ المدفوع (F8) - اتركه فارغًا إذا دفع المبلغ بالضبط", c)
    m.add(Control("label", "lblChange", px, cm(14.55), pw, cm(0.8),
                  {"Caption": " ", "FontSize": 14, "FontBold": True, "TextAlign": 2,
                   "ForeColor": Sym("CLR_SUCCESS")}))
    m.add(Control("label", "lblLastInvoiceCap", px, cm(15.55), cm(3.0), cm(0.55),
                  {"Caption": "آخر فاتورة:", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("label", "lblLastInvoice", px + cm(3.1), cm(15.55), cm(5.0), cm(0.55),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    button(m, "btnClose", "إغلاق", px + pw - cm(3.0), cm(16.6), "secondary", w=cm(3.0), h=cm(1.1),
           call="DoCmd.Close acForm, Me.Name")

    m.form_events = ["Load", "KeyDown", "Unload"]
    m.code = ([
        "Private Sub Form_Load()", "    POSLoad Me", "End Sub",
        "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)",
        "    POSKeyDown Me, KeyCode, Shift", "End Sub",
        "Private Sub Form_Unload(Cancel As Integer)", "    Cancel = Not POSUnload(Me)", "End Sub",
        "Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)",
        "    BarcodeKeyDown Me, KeyCode", "End Sub",
        "Private Sub cboProduct_AfterUpdate()", "    ProductPicked Me", "End Sub",
        "Private Sub cboCustomer_AfterUpdate()", "    CustomerChanged Me", "End Sub",
        "Private Sub cboPaymentType_AfterUpdate()", "    PaymentTypeChanged Me", "End Sub",
        "Private Sub cboPaymentMethod_AfterUpdate()", "    PaymentMethodChanged Me", "End Sub",
        "Private Sub txtInvoiceDiscount_AfterUpdate()", "    RecalcPOS Me", "End Sub",
        "Private Sub txtTendered_AfterUpdate()", "    RecalcPOS Me", "End Sub",
    ] + m.code)
    return m


# ------------------------------------------------------------------ returns
RETURN_TITLES = ["الصنف", "الكمية المباعة", "مرتجع سابقًا", "المتاح للإرجاع", "الكمية المرتجعة",
                 "يعود للمخزون", "قيمة المرتجع"]


def layout_return_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.75)
    m = FormModel("frmReturnLines", "أسطر المرتجع", cm(26.2), row_h, popup=False,
                  record_source="SELECT * FROM tmpReturnLines ORDER BY SalesDetailID", allow_add=False)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0}
    heads = grid_row(m, [
        ("ProductName", "ProductName", 9.0, dict(LOCKED), []),
        ("SoldQty", "SoldQty", 2.6, {**LOCKED, "Format": "#,##0.###"}, []),
        ("ReturnedQty", "ReturnedQty", 2.6, {**LOCKED, "Format": "#,##0.###"}, []),
        ("AvailableQty", "AvailableQty", 2.6, {**LOCKED, "Format": "#,##0.###"}, []),
        ("ReturnQty", "ReturnQty", 2.8, {"Format": "#,##0.###", "FontBold": True}, ["AfterUpdate"]),
        ("ReturnToStock", "ReturnToStock", 2.6, {"_kind": "check"}, ["AfterUpdate"]),
        ("ReturnAmount", "ReturnAmount", 3.4, {**LOCKED, "Format": "#,##0.00", "FontBold": True}, []),
    ], row_h)
    m.code = ["Private Sub ReturnQty_AfterUpdate()", "    ReturnLineChanged Me", "End Sub",
              "Private Sub ReturnToStock_AfterUpdate()", "    ReturnLineChanged Me", "End Sub"]
    return m, heads


def layout_sales_return(line_heads) -> FormModel:
    width, height = cm(27.0), cm(17.0)
    m = FormModel("frmSalesReturn", "مرتجع مبيعات", width, height, popup=True, allow_add=False)
    title_band(m, "مرتجع مبيعات (إشعار دائن)", "اختر الفاتورة الأصلية ثم حدد الكميات المرتجعة",
               "sales")
    m.add(Control("text", "txtInvoiceID", width - cm(1.0), cm(0.1), cm(0.6), cm(0.4),
                  {"Visible": False}, decorative=True))
    c = m.add(Control("text", "txtInvoiceNo", cm(0.4), cm(2.3), cm(5.0), cm(0.8), {"FontSize": 12},
                      events=["AfterUpdate"]))
    labelled(m, "txtInvoiceNo", "رقم الفاتورة الأصلية", c)
    button(m, "btnFind", "بحث", cm(5.6), cm(2.3), "secondary", w=cm(2.0), h=cm(0.8),
           call="ReturnFindInvoice Me")
    m.add(Control("label", "lblInvoiceInfo", cm(7.9), cm(2.4), width - cm(8.3), cm(0.7),
                  {"Caption": " ", "FontSize": 10, "ForeColor": Sym("CLR_TEXT")}))
    header_labels(m, cm(0.4), cm(3.45), RETURN_TITLES, line_heads)
    m.add(Control("subform", "subReturnLines", cm(0.4), cm(4.05), cm(26.2), cm(7.6),
                  {"SourceObject": "frmReturnLines"}))
    c = m.add(Control("text", "txtReason", cm(0.4), cm(12.35), cm(11.6), cm(0.8), {}))
    labelled(m, "txtReason", "سبب الإرجاع *", c)
    c = m.add(Control("combo", "cboRefundType", cm(12.3), cm(12.35), cm(6.6), cm(0.8),
                      {"RowSource": REFUND_TYPES, "ColumnCount": 2, "ColumnWidths": "0;6"}))
    labelled(m, "cboRefundType", "طريقة رد المبلغ", c)
    c = m.add(Control("combo", "cboPaymentMethod", cm(19.2), cm(12.35), cm(7.4), cm(0.8),
                      {"RowSource": PAYMENT_ROWS, "ColumnCount": 2, "ColumnWidths": "0;6"}))
    labelled(m, "cboPaymentMethod", "طريقة الرد", c)
    m.add(Control("label", "lblReturnTotalCap", cm(0.4), cm(13.6), cm(4.0), cm(0.8),
                  {"Caption": "قيمة المرتجع:", "FontSize": 13, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("label", "lblReturnTotal", cm(4.5), cm(13.4), cm(6.0), cm(1.1),
                  {"Caption": "0.00", "FontSize": 20, "FontBold": True, "ForeColor": Sym("CLR_DANGER")}))
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnReturnAll", "إرجاع الكل", "secondary", 3.2, "ReturnAll Me"),
            ("btnSaveReturn", "حفظ المرتجع", "primary", 3.6, "SaveReturn Me, False"),
            ("btnSaveReturnPrint", "حفظ وطباعة", "primary", 3.6, "SaveReturn Me, True")]:
        button(m, name, caption, bx, cm(15.3), style, w=cm(w), h=cm(1.0), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(15.3), "secondary", w=cm(2.6),
           h=cm(1.0), call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    ReturnLoad Me", "End Sub",
               "Private Sub txtInvoiceNo_AfterUpdate()", "    ReturnFindInvoice Me", "End Sub"]
              + m.code)
    return m


# ------------------------------------------------------------------ receipt voucher
def layout_customer_payment() -> FormModel:
    width, height = cm(16.0), cm(10.2)
    m = FormModel("frmCustomerPayment", "سند قبض", width, height, popup=True, allow_add=False)
    title_band(m, "سند قبض", "تسجيل دفعة من عميل على حسابه", "customers")
    c = m.add(Control("combo", "cboCustomer", cm(0.4), cm(2.4), width - cm(0.8), cm(0.8),
                      {"RowSource": CREDIT_CUSTOMER_ROWS, "ColumnCount": 2, "ColumnWidths": "0;10"},
                      events=["AfterUpdate"]))
    labelled(m, "cboCustomer", "العميل", c)
    m.add(Control("label", "lblBalance", cm(0.4), cm(3.3), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    c = m.add(Control("text", "txtAmount", cm(0.4), cm(4.6), cm(7.4), cm(0.9),
                      {"FontSize": 14, "Format": "#,##0.00"}))
    labelled(m, "txtAmount", "المبلغ *", c)
    c = m.add(Control("combo", "cboPaymentMethod", cm(8.2), cm(4.6), width - cm(8.6), cm(0.9),
                      {"RowSource": PAYMENT_ROWS, "ColumnCount": 2, "ColumnWidths": "0;6"}))
    labelled(m, "cboPaymentMethod", "طريقة الدفع", c)
    c = m.add(Control("text", "txtNotes", cm(0.4), cm(6.2), width - cm(0.8), cm(0.8), {}))
    labelled(m, "txtNotes", "ملاحظات", c)
    button(m, "btnSave", "حفظ السند", cm(0.4), cm(8.4), "primary", w=cm(3.4), h=cm(1.0),
           call="SavePayment Me")
    button(m, "btnSavePrint", "حفظ وطباعة", cm(4.0), cm(8.4), "primary", w=cm(3.4), h=cm(1.0),
           call="SavePayment Me, True")
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(8.4), "secondary", w=cm(2.6),
           h=cm(1.0), call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    PaymentLoad Me", "End Sub",
               "Private Sub cboCustomer_AfterUpdate()", "    PaymentCustomerChanged Me", "End Sub"]
              + m.code)
    return m


# ------------------------------------------------------------------ invoice view
def layout_sales_invoice() -> FormModel:
    width, height = cm(27.0), cm(15.6)
    m = FormModel("frmSalesInvoice", "فاتورة بيع", width, height, popup=True, allow_add=False)
    title_band(m, "فاتورة بيع", "عرض فقط - التصحيح يكون بمرتجع", "sales")
    m.add(Control("text", "txtInvoiceID", width - cm(1.0), cm(0.1), cm(0.6), cm(0.4),
                  {"Visible": False}, decorative=True))
    m.add(Control("label", "lblHeader", cm(0.4), cm(1.8), width - cm(0.8), cm(0.7),
                  {"Caption": " ", "FontSize": 10, "ForeColor": Sym("CLR_TEXT")}))
    m.add(Control("list", "lstLines", cm(0.4), cm(2.7), width - cm(0.8), cm(6.6),
                  {"ColumnCount": 7, "ColumnWidths": "1;9;2.5;3.5;2.5;3;4", "ColumnHeads": True}))
    m.add(Control("label", "lblReturnsCap", cm(0.4), cm(9.5), cm(8), cm(0.55),
                  {"Caption": "المرتجعات على هذه الفاتورة", "FontSize": 10, "FontBold": True,
                   "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstReturns", cm(0.4), cm(10.1), width - cm(0.8), cm(2.3),
                  {"ColumnCount": 4, "ColumnWidths": "4;5;3.5;12", "ColumnHeads": True}))
    m.add(Control("label", "lblTotals", cm(0.4), cm(12.6), width - cm(0.8), cm(0.7),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnPrint", "طباعة", "primary", 2.8, 'PrintSalesDocument "SALE", Me!txtInvoiceID.Value'),
            ("btnPrintA4", "طباعة A4", "secondary", 2.8,
             'PrintSalesDocument "SALE", Me!txtInvoiceID.Value, True'),
            ("btnReturn", "مرتجع", "secondary", 2.6, 'OpenScreen "frmSalesReturn", 6, Me!txtInvoiceID.Value')]:
        button(m, name, caption, bx, cm(13.9), style, w=cm(w), h=cm(1.0), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(13.9), "secondary", w=cm(2.6),
           h=cm(1.0), call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = ["Private Sub Form_Load()", "    InvoiceViewLoad Me", "End Sub"] + m.code
    return m


def sales_forms() -> List[FormModel]:
    lines, pos_heads = layout_pos_lines()
    ret_lines, ret_heads = layout_return_lines()
    # subforms first: the parent forms embed them by name
    return [lines, layout_pos(pos_heads), ret_lines, layout_sales_return(ret_heads),
            layout_customer_payment(), layout_sales_invoice()]
