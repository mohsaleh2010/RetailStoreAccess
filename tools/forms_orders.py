"""Customer orders (docs/62, plan 4A): frmSalesOrder and its lines frmSalesOrderLines.
Layout conventions are the same as forms.py (twips, x from the start edge)."""

from typing import List, Tuple

from forms import Control, FormModel, Sym, button, cm, labelled, title_band
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


def order_forms() -> List[FormModel]:
    lines, heads = layout_sales_order_lines()
    # the subform first: the parent form embeds it by name
    return [lines, layout_sales_order(heads)]
