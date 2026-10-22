"""Customer orders (docs/62, plan 4A): frmSalesOrder and its lines frmSalesOrderLines.
Layout conventions are the same as forms.py (twips, x from the start edge)."""

from typing import List, Tuple

from forms import Control, FormModel, Sym, button, cm, fit_window, labelled, title_band
from forms_purchases import TRANSFER_PRODUCTS
from forms_sales import LOCKED, grid_row, header_labels

SALES_ORDER_TITLES = ["الصنف", "الكمية المطلوبة", "سعر البيع", "المُسلَّم", "قيمة السطر"]


def layout_sales_order_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.75)
    m = FormModel("frmSalesOrderLines", "أسطر طلبية العميل", cm(20.6), row_h, popup=False,
                  record_source="SELECT * FROM SalesOrderLines WHERE SalesOrderID = 0", allow_add=True)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0, "AllowDeletions": False}
    heads = grid_row(m, [
        ("ProductID", "ProductID", 7.8, {"_kind": "combo", "RowSource": TRANSFER_PRODUCTS, "ColumnCount": 2,
                                         "ColumnWidths": "0;7.6", "BoundColumn": 1, "LimitToList": True}, []),
        ("Quantity", "Quantity", 2.4, {"Format": "#,##0.###"}, []),
        ("UnitPrice", "UnitPrice", 2.8, {"Format": "#,##0.00##"}, []),
        ("txtDelivered", "=OrderLineDelivered([SalesOrderLineID])", 2.4, {**LOCKED, "Format": "#,##0.###"}, []),
        ("txtValue", "=Nz([Quantity],0)*Nz([UnitPrice],0)", 3.2, {**LOCKED, "Format": "#,##0.00"}, []),
        ("btnRemove", "", 0.8, {"_kind": "button", "Caption": Sym(f"ChrW(&H{0xE74D:X})"),
                               "Style": "danger", "FontName": Sym("ICON_FONT")}, ["Click"]),
    ], row_h)
    m.form_events = ["BeforeUpdate", "AfterUpdate"]
    m.code = ["Private Sub Form_BeforeUpdate(Cancel As Integer)", "    SalesOrderLineBeforeUpdate Me, Cancel", "End Sub",
              "Private Sub Form_AfterUpdate()", "    SalesOrderLineAfterUpdate Me", "End Sub",
              "Private Sub btnRemove_Click()", "    SalesOrderRemoveLine Me", "End Sub"]
    return m, heads


def layout_sales_order(line_heads) -> FormModel:
    width, height = cm(21.4), cm(18.0)
    m = FormModel("frmSalesOrder", "طلبيات العملاء", width, height, popup=True, allow_add=False)
    title_band(m, "طلبيات العملاء",
               "«اعتماد» يحجز الكمية في المخزن، ثم «إذن التسليم» من الطلبية، ثم الفاتورة من الإذن",
               "sales")
    c = m.add(Control("combo", "cboOrder", cm(0.4), cm(2.3), cm(16.6), cm(0.8),
                      {"RowSource": "", "ColumnCount": 6, "ColumnWidths": "0;2.4;2.2;5.6;2.8;3.2", "LimitToList": True},
                      events=["AfterUpdate"]))
    labelled(m, "cboOrder", "الطلبية (الأحدث أولًا)", c)
    y = cm(3.8)
    c = m.add(Control("combo", "cboCustomer", cm(0.4), y, cm(5.0), cm(0.8),
                      {"RowSource": "", "ColumnCount": 2, "ColumnWidths": "0;4.8", "LimitToList": True}))
    labelled(m, "cboCustomer", "العميل", c)
    c = m.add(Control("combo", "cboWarehouse", cm(5.6), y, cm(3.4), cm(0.8),
                      {"RowSource": "", "ColumnCount": 2, "ColumnWidths": "0;3.2", "LimitToList": True}))
    labelled(m, "cboWarehouse", "المخزن", c)
    c = m.add(Control("text", "txtRequired", cm(9.2), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtRequired", "موعد التسليم", c)
    c = m.add(Control("text", "txtNotes", cm(12.4), y, cm(4.8), cm(0.8), {}))
    labelled(m, "txtNotes", "ملاحظات", c)
    button(m, "btnNew", "طلبية جديدة", cm(17.4), y - cm(0.03), "primary", w=cm(3.6), call="SalesOrderNew Me")
    m.add(Control("label", "lblInfo", cm(0.4), cm(4.9), width - cm(0.8), cm(0.7),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    header_labels(m, cm(0.4), cm(5.75), SALES_ORDER_TITLES, line_heads)
    m.add(Control("subform", "subLines", cm(0.4), cm(6.35), cm(20.6), cm(7.0),
                  {"SourceObject": "frmSalesOrderLines"}))
    bx = cm(0.4)                                        # the way of the order
    for name, caption, style, w, call in [
            ("btnApprove", "اعتماد وحجز", "primary", 3.4, 'SalesOrderAction Me, "APPROVE"'),
            ("btnPrepare", "بدء التجهيز", "secondary", 3.0, 'SalesOrderAction Me, "PREPARE"'),
            ("btnDeliver", "إذن التسليم", "primary", 3.4, 'SalesOrderAction Me, "DELIVER"'),
            ("btnNotes", "أذونات الطلبية", "secondary", 3.2, "SalesOrderOpenNote Me")]:
        button(m, name, caption, bx, cm(14.1), style, w=cm(w), call=call)
        bx += cm(w) + cm(0.2)
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnUnapprove", "إعادة لمسودة", "secondary", 3.2, 'SalesOrderAction Me, "UNAPPROVE"'),
            ("btnCancel", "إلغاء الطلبية", "danger", 3.2, 'SalesOrderAction Me, "CANCEL"'),
            ("btnDelete", "حذف المسودة", "danger", 3.2, 'SalesOrderAction Me, "DELETE"')]:
        button(m, name, caption, bx, cm(16.1), style, w=cm(w), h=cm(1.1), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(16.1), "secondary", w=cm(2.6), h=cm(1.1),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load", "Activate"]
    m.code = (["Private Sub Form_Load()", "    SalesOrderLoad Me", "End Sub",
               "Private Sub Form_Activate()", "    SalesOrderShow Me", "End Sub",       # back from its delivery note
               "Private Sub cboOrder_AfterUpdate()", "    SalesOrderShow Me", "End Sub"] + m.code)
    return m


PURCHASE_ORDER_TITLES = ["الصنف", "الكمية المطلوبة", "سعر الشراء", "المستلم", "قيمة السطر"]


def layout_purchase_order_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.75)
    m = FormModel("frmPurchaseOrderLines", "أسطر أمر الشراء", cm(20.6), row_h, popup=False,
                  record_source="SELECT * FROM PurchaseOrderLines WHERE PurchaseOrderID = 0", allow_add=True)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0, "AllowDeletions": False}
    heads = grid_row(m, [
        ("ProductID", "ProductID", 7.8, {"_kind": "combo", "RowSource": TRANSFER_PRODUCTS, "ColumnCount": 2,
                                         "ColumnWidths": "0;7.6", "BoundColumn": 1, "LimitToList": True}, []),
        ("Quantity", "Quantity", 2.4, {"Format": "#,##0.###"}, []),
        ("UnitCost", "UnitCost", 2.8, {"Format": "#,##0.00##"}, []),
        ("txtReceived", "=PurchaseOrderLineReceived([PurchaseOrderLineID])", 2.4, {**LOCKED, "Format": "#,##0.###"}, []),
        ("txtValue", "=Nz([Quantity],0)*Nz([UnitCost],0)", 3.2, {**LOCKED, "Format": "#,##0.00"}, []),
        ("btnRemove", "", 0.8, {"_kind": "button", "Caption": Sym(f"ChrW(&H{0xE74D:X})"),
                               "Style": "danger", "FontName": Sym("ICON_FONT")}, ["Click"]),
    ], row_h)
    m.form_events = ["BeforeUpdate", "AfterUpdate"]
    m.code = ["Private Sub Form_BeforeUpdate(Cancel As Integer)", "    PurchaseOrderLineBeforeUpdate Me, Cancel", "End Sub",
              "Private Sub Form_AfterUpdate()", "    PurchaseOrderLineAfterUpdate Me", "End Sub",
              "Private Sub btnRemove_Click()", "    PurchaseOrderRemoveLine Me", "End Sub"]
    return m, heads


def layout_purchase_order(line_heads) -> FormModel:
    width, height = cm(21.4), cm(18.0)
    m = FormModel("frmPurchaseOrder", "أوامر الشراء", width, height, popup=True, allow_add=False)
    title_band(m, "أوامر الشراء",
               "«اعتماد» يجعل الكمية تحت الطلب، ثم «إذن الاستلام» من الأمر، ثم فاتورة المشتريات من الإذن",
               "purchases")
    c = m.add(Control("combo", "cboOrder", cm(0.4), cm(2.3), cm(16.6), cm(0.8),
                      {"RowSource": "", "ColumnCount": 6, "ColumnWidths": "0;2.4;2.2;5.6;2.8;3.2", "LimitToList": True},
                      events=["AfterUpdate"]))
    labelled(m, "cboOrder", "أمر الشراء (الأحدث أولًا)", c)
    y = cm(3.8)
    c = m.add(Control("combo", "cboSupplier", cm(0.4), y, cm(5.0), cm(0.8),
                      {"RowSource": "", "ColumnCount": 2, "ColumnWidths": "0;4.8", "LimitToList": True}))
    labelled(m, "cboSupplier", "المورد", c)
    c = m.add(Control("combo", "cboWarehouse", cm(5.6), y, cm(3.4), cm(0.8),
                      {"RowSource": "", "ColumnCount": 2, "ColumnWidths": "0;3.2", "LimitToList": True}))
    labelled(m, "cboWarehouse", "المخزن", c)
    c = m.add(Control("text", "txtExpected", cm(9.2), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtExpected", "موعد الاستلام", c)
    c = m.add(Control("text", "txtNotes", cm(12.4), y, cm(4.8), cm(0.8), {}))
    labelled(m, "txtNotes", "ملاحظات", c)
    button(m, "btnNew", "أمر جديد", cm(17.4), y - cm(0.03), "primary", w=cm(3.6), call="PurchaseOrderNew Me")
    m.add(Control("label", "lblInfo", cm(0.4), cm(4.9), width - cm(0.8), cm(0.7),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    header_labels(m, cm(0.4), cm(5.75), PURCHASE_ORDER_TITLES, line_heads)
    m.add(Control("subform", "subLines", cm(0.4), cm(6.35), cm(20.6), cm(7.0),
                  {"SourceObject": "frmPurchaseOrderLines"}))
    bx = cm(0.4)                                        # the way of the order
    for name, caption, style, w, call in [
            ("btnApprove", "اعتماد", "primary", 2.6, 'PurchaseOrderAction Me, "APPROVE"'),
            ("btnReceive", "إذن الاستلام", "primary", 3.2, 'PurchaseOrderAction Me, "RECEIVE"'),
            ("btnNotes", "أذونات الأمر", "secondary", 3.0, "PurchaseOrderOpenNote Me"),
            ("btnSuggest", "أوامر من الاقتراح", "secondary", 3.8, "PurchaseOrderSuggest Me"),
            ("btnSuggestReport", "اقتراح الشراء", "secondary", 3.2,
             'OpenScreen "frmReportCenter", 0, "PURCHASE_SUGGEST"'),
            ("btnPending", "المفتوحة", "secondary", 2.6, 'OpenScreen "frmReportCenter", 0, "PENDING_PURCHASE_ORDERS"')]:
        button(m, name, caption, bx, cm(14.1), style, w=cm(w), call=call)
        bx += cm(w) + cm(0.2)
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnUnapprove", "إعادة لمسودة", "secondary", 3.2, 'PurchaseOrderAction Me, "UNAPPROVE"'),
            ("btnCancel", "إلغاء الأمر", "danger", 3.2, 'PurchaseOrderAction Me, "CANCEL"'),
            ("btnDelete", "حذف المسودة", "danger", 3.2, 'PurchaseOrderAction Me, "DELETE"')]:
        button(m, name, caption, bx, cm(16.1), style, w=cm(w), h=cm(1.1), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(16.1), "secondary", w=cm(2.6), h=cm(1.1),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load", "Activate"]
    m.code = (["Private Sub Form_Load()", "    PurchaseOrderLoad Me", "End Sub",
               "Private Sub Form_Activate()", "    PurchaseOrderShow Me", "End Sub",    # back from its receipt note
               "Private Sub cboOrder_AfterUpdate()", "    PurchaseOrderShow Me", "End Sub"] + m.code)
    return m


def order_forms() -> List[FormModel]:
    lines, heads = layout_sales_order_lines()
    po_lines, po_heads = layout_purchase_order_lines()
    # the subform first: the parent form embeds it by name
    return [lines, layout_sales_order(heads), po_lines, layout_purchase_order(po_heads), layout_rep_day(),
            layout_rep_visits()]


# ---------------------------------------------------------------- the rep day and van sales (docs/63)
def layout_rep_day() -> FormModel:
    width, height = cm(24.0), cm(16.4)
    m = FormModel("frmRepDay", "يوم المندوب", width, height, popup=False, allow_add=False)
    title_band(m, "يوم المندوب والبيع من السيارة",
               "تحميل السيارة، والبيع والمرتجع والتحصيل باسم المندوب، ثم جرد السيارة وتصفية صندوق العهدة",
               "customers")
    c = m.add(Control("combo", "cboDay", cm(0.4), cm(2.3), cm(15.0), cm(0.8),
                      {"RowSource": "", "ColumnCount": 6, "ColumnWidths": "0;2.4;2.2;4.4;3.4;2.4", "LimitToList": True},
                      events=["AfterUpdate"]))
    labelled(m, "cboDay", "اليوم (الأحدث أولًا)", c)
    y = cm(3.8)
    c = m.add(Control("combo", "cboRep", cm(0.4), y, cm(5.0), cm(0.8),
                      {"RowSource": "", "ColumnCount": 2, "ColumnWidths": "0;4.8", "LimitToList": True}))
    labelled(m, "cboRep", "المندوب", c)
    c = m.add(Control("combo", "cboStore", cm(5.6), y, cm(4.0), cm(0.8),
                      {"RowSource": "", "ColumnCount": 2, "ColumnWidths": "0;3.8", "LimitToList": True}))
    labelled(m, "cboStore", "يحمّل من المخزن", c)
    button(m, "btnOpen", "فتح يوم", cm(9.8), y - cm(0.03), "primary", w=cm(3.0), call="RepDayOpen Me")
    button(m, "btnDeleteDay", "حذف اليوم", cm(13.0), y - cm(0.03), "danger", w=cm(2.8), call='RepDayAction Me, "DELETE"')
    m.add(Control("label", "lblInfo", cm(0.4), cm(4.9), width - cm(0.8), cm(0.7),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("label", "lblActing", cm(0.4), cm(5.6), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_DANGER")}))
    m.add(Control("list", "lstSummary", cm(0.4), cm(6.4), width - cm(0.8), cm(5.8),
                  {"RowSourceType": "Value List", "ColumnCount": 2, "ColumnWidths": "14;6", "ColumnHeads": False}))
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnLoad", "تحميل السيارة", "primary", 3.4, 'RepDayAction Me, "LOAD"'),
            ("btnSale", "فاتورة بيع", "secondary", 3.0, 'RepDayAction Me, "SALE"'),
            ("btnReturn", "مرتجع", "secondary", 2.6, 'RepDayAction Me, "RETURN"'),
            ("btnPayment", "سند قبض", "secondary", 2.6, 'RepDayAction Me, "PAYMENT"'),
            ("btnRefresh", "تحديث", "secondary", 2.4, "RepDayShow Me")]:
        button(m, name, caption, bx, cm(12.6), style, w=cm(w), call=call)
        bx += cm(w) + cm(0.2)
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnCount", "جرد السيارة", "secondary", 3.2, 'RepDayAction Me, "COUNT"'),
            ("btnUnload", "تفريغ الباقي", "secondary", 3.2, 'RepDayAction Me, "UNLOAD"'),
            ("btnCash", "تصفية الصندوق", "secondary", 3.4, 'RepDayAction Me, "CASH"'),
            ("btnSettle", "إغلاق اليوم", "primary", 3.2, 'RepDayAction Me, "SETTLE"'),
            ("btnPrint", "طباعة", "secondary", 2.4, "RepDayPrint Me")]:
        button(m, name, caption, bx, cm(14.4), style, w=cm(w), h=cm(1.1), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(14.4), "secondary", w=cm(2.6), h=cm(1.1),
           call="DoCmd.Close acForm, Me.Name")
    fit_window(m, bottom_y=cm(12.5), stretch_w=("lblInfo", "lblActing", "lstSummary"), stretch_h=("lstSummary",),
               extra={"btnClose": (1000, 0, 1000, 0)})
    m.form_events = ["Load", "Activate", "Unload"]
    m.code = (["Private Sub Form_Load()", "    RepDayLoad Me", "End Sub",
               "Private Sub Form_Activate()", "    RepDayShow Me", "End Sub",          # back from a sale or a count
               "Private Sub Form_Unload(Cancel As Integer)", "    StopActingAsRep", "End Sub",
               "Private Sub cboDay_AfterUpdate()", "    RepDayShow Me", "End Sub"] + m.code)
    return m


# ---------------------------------------------------------------- routes and visits of the reps (docs/64)
def layout_rep_visits() -> FormModel:
    width, height = cm(24.0), cm(16.4)
    m = FormModel("frmRepVisits", "زيارات المندوبين", width, height, popup=False, allow_add=False)
    title_band(m, "زيارات المندوبين وخطة اليوم",
               "خطة اليوم من خط السير، ثم تسجيل كل زيارة ونتيجتها: طلبية أو بيع أو تحصيل أو لا شيء",
               "customers")
    y = cm(2.3)
    c = m.add(Control("combo", "cboRep", cm(0.4), y, cm(5.0), cm(0.8),
                      {"RowSource": "", "ColumnCount": 2, "ColumnWidths": "0;4.8", "LimitToList": True},
                      events=["AfterUpdate"]))
    labelled(m, "cboRep", "المندوب", c)
    c = m.add(Control("text", "txtDate", cm(5.6), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}, events=["AfterUpdate"]))
    labelled(m, "txtDate", "اليوم", c)
    button(m, "btnPlan", "خطة اليوم من خط السير", cm(8.8), y - cm(0.03), "primary", w=cm(4.6), call="RepVisitsPlan Me")
    button(m, "btnRoute", "خط السير", cm(13.6), y - cm(0.03), "secondary", w=cm(2.8),
           call='OpenScreen "frmRouteStops", 0')
    button(m, "btnPrices", "أسعار المندوب", cm(16.6), y - cm(0.03), "secondary", w=cm(3.2),
           call='OpenScreen "frmRepPrices", 0')
    y = cm(3.8)
    c = m.add(Control("combo", "cboCustomer", cm(0.4), y, cm(8.2), cm(0.8),
                      {"RowSource": "", "ColumnCount": 2, "ColumnWidths": "0;8", "LimitToList": True}))
    labelled(m, "cboCustomer", "عميل خارج خط السير", c)
    button(m, "btnAddVisit", "إضافة زيارة", cm(8.8), y - cm(0.03), "secondary", w=cm(3.2), call="RepVisitsAdd Me")
    m.add(Control("label", "lblInfo", cm(0.4), cm(4.9), width - cm(0.8), cm(0.7),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("list", "lstVisits", cm(0.4), cm(5.7), width - cm(0.8), cm(6.4),
                  {"ColumnCount": 7, "ColumnWidths": "0;1.2;7;2.6;2.6;2.4;6.4", "ColumnHeads": True},
                  events=["AfterUpdate"]))
    y = cm(12.8)
    c = m.add(Control("combo", "cboResult", cm(0.4), y, cm(4.0), cm(0.8),
                      {"RowSourceType": "Value List", "RowSource": "", "ColumnCount": 2, "ColumnWidths": "0;3.8",
                       "LimitToList": True}))
    labelled(m, "cboResult", "نتيجة الزيارة", c)
    c = m.add(Control("text", "txtVisitNotes", cm(4.6), y, cm(8.0), cm(0.8), {}))
    labelled(m, "txtVisitNotes", "ملاحظات الزيارة", c)
    button(m, "btnRecord", "تسجيل الزيارة", cm(12.8), y - cm(0.03), "primary", w=cm(3.4), call="RepVisitsRecord Me")
    button(m, "btnDeleteVisit", "حذف الزيارة", cm(16.4), y - cm(0.03), "danger", w=cm(3.0), call="RepVisitsDelete Me")
    bx = cm(0.4)
    for name, caption, w, call in [
            ("btnExecution", "تقرير تنفيذ خط السير", 4.6, 'OpenScreen "frmReportCenter", 0, "ROUTE_EXECUTION"'),
            ("btnVisitsReport", "تقرير الزيارات", 3.4, 'OpenScreen "frmReportCenter", 0, "REP_VISITS"'),
            ("btnVanStock", "رصيد السيارات", 3.4, 'OpenScreen "frmReportCenter", 0, "REP_VAN_STOCK"')]:
        button(m, name, caption, bx, cm(14.4), "secondary", w=cm(w), h=cm(1.1), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(14.4), "secondary", w=cm(2.6), h=cm(1.1),
           call="DoCmd.Close acForm, Me.Name")
    fit_window(m, bottom_y=cm(12.2), stretch_w=("lblInfo", "lstVisits"), stretch_h=("lstVisits",),
               extra={"btnClose": (1000, 0, 1000, 0)})
    m.form_events = ["Load", "Activate"]
    m.code = (["Private Sub Form_Load()", "    RepVisitsLoad Me", "End Sub",
               "Private Sub Form_Activate()", "    RepVisitsShow Me", "End Sub",     # back from the route screen
               "Private Sub cboRep_AfterUpdate()", "    RepVisitsShow Me", "End Sub",
               "Private Sub txtDate_AfterUpdate()", "    RepVisitsShow Me", "End Sub",
               "Private Sub lstVisits_AfterUpdate()", "    RepVisitPicked Me", "End Sub"] + m.code)
    return m
