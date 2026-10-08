"""The e-invoicing screen (frmEInvoices, docs/45-EInvoice-Foundation.md): the sales invoices and returns with
their e-invoicing status, the requests of the chosen document, and sending again. The logic is in modEInvoice."""

from typing import List

from forms import Control, FormModel, Sym, button, cm, labelled, title_band
from forms_security import check_with_label

ENVIRONMENTS = "TEST;تجريبية;SIMULATION;محاكاة (السعودية);PRODUCTION;فعلية"
STATUS_FILTER = ("ATTENTION;تحتاج متابعة (بانتظار الإرسال، مرفوضة، تحذير);SENT;أُرسلت;"
                 "NOT_SENT;قبل التفعيل;ALL;الكل")


def layout_einvoices() -> FormModel:
    width, height = cm(27.0), cm(18.4)
    m = FormModel("frmEInvoices", "الفاتورة الإلكترونية", width, height, popup=True, allow_add=False)
    title_band(m, "الفاتورة الإلكترونية", "حالة إرسال فواتير البيع والمرتجعات للمنظومة، وإعادة الإرسال", "sales")
    y = cm(2.3)
    c = m.add(Control("text", "txtFrom", cm(0.4), y, cm(2.8), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtFrom", "من", c)
    c = m.add(Control("text", "txtTo", cm(3.4), y, cm(2.8), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtTo", "إلى", c)
    c = m.add(Control("combo", "cboStatus", cm(6.4), y, cm(8.0), cm(0.8),
                      {"RowSourceType": "Value List", "RowSource": STATUS_FILTER, "ColumnCount": 2,
                       "ColumnWidths": "0;7.8", "LimitToList": True}, events=["AfterUpdate"]))
    labelled(m, "cboStatus", "المستندات", c)
    button(m, "btnShow", "عرض", cm(14.6), y, "primary", w=cm(3.0), h=cm(0.8), call="EInvoicesShow Me")
    # the e-invoicing settings (Settings.EInvoiceEnabled / EInvoiceEnvironment), saved at once (modEInvoice)
    c = m.add(Control("combo", "cboEnvironment", cm(18.0), y, cm(4.4), cm(0.8),
                      {"RowSourceType": "Value List", "RowSource": ENVIRONMENTS, "ColumnCount": 2,
                       "ColumnWidths": "0;4.2", "LimitToList": True}, events=["AfterUpdate"]))
    labelled(m, "cboEnvironment", "البيئة", c)
    check_with_label(m, "chkEnabled", "تفعيل الإرسال", cm(22.8), y, cm(3.8), events=["AfterUpdate"])
    m.add(Control("label", "lblSummary", cm(0.4), cm(3.3), width - cm(0.8), cm(0.55),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("list", "lstDocs", cm(0.4), cm(4.0), width - cm(0.8), cm(7.4),
                  {"ColumnCount": 9, "ColumnWidths": "0;2.8;3.2;2.6;4.4;2.6;3.2;1.8;5.6", "ColumnHeads": True},
                  events=["AfterUpdate"]))
    m.add(Control("label", "lblLogTitle", cm(0.4), cm(11.5), width - cm(0.8), cm(0.55),
                  {"Caption": "طلبات المستند المختار وردود المنظومة", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstLog", cm(0.4), cm(12.1), width - cm(0.8), cm(4.0),
                  {"ColumnCount": 6, "ColumnWidths": "0;3.6;2.4;2.4;1.8;15.4", "ColumnHeads": True}))
    y = cm(16.6)
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnSendPicked", "إرسال المختار", "primary", 3.6, "EInvoicesSendPicked Me"),
            ("btnSendAll", "إرسال كل المعلّق", "secondary", 3.8, "EInvoicesSendAll Me")]:
        button(m, name, caption, bx, y, style, w=cm(w), h=cm(0.9), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), y, "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    EInvoicesLoad Me", "End Sub",
               "Private Sub cboStatus_AfterUpdate()", "    EInvoicesShow Me", "End Sub",
               "Private Sub lstDocs_AfterUpdate()", "    EInvoicesPick Me", "End Sub",
               "Private Sub cboEnvironment_AfterUpdate()", "    EInvoiceSettingChanged Me", "End Sub",
               "Private Sub chkEnabled_AfterUpdate()", "    EInvoiceSettingChanged Me", "End Sub"] + m.code)
    return m


def einvoice_forms() -> List[FormModel]:
    return [layout_einvoices()]
