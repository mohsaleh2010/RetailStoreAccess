"""Receivables and payables by age (frmAging) and the links between payments and invoices
(frmAllocation). The logic is in modAging. Layout conventions are the same as forms.py."""

from typing import List

from forms import Control, FormModel, Sym, button, cm, labelled, title_band

KINDS = "C;العملاء;S;الموردون"


def _kind_combo(m: FormModel, x, y, event: str):
    c = m.add(Control("combo", "cboKind", x, y, cm(3.2), cm(0.8),
                      {"RowSourceType": "Value List", "RowSource": KINDS, "ColumnCount": 2,
                       "ColumnWidths": "0;3", "LimitToList": True}, events=["AfterUpdate"]))
    labelled(m, "cboKind", "النوع", c)
    m.code += ["Private Sub cboKind_AfterUpdate()", f"    {event} Me", "End Sub"]
    return c


def layout_aging() -> FormModel:
    width, height = cm(27.0), cm(18.4)
    m = FormModel("frmAging", "أعمار الديون", width, height, popup=True, allow_add=False)
    title_band(m, "أعمار الديون", "المتبقي من كل فاتورة آجلة حسب تأخيرها عن تاريخ الاستحقاق", "reports")
    y = cm(2.3)
    _kind_combo(m, cm(0.4), y, "AgingRefresh")
    c = m.add(Control("text", "txtAsOf", cm(3.8), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtAsOf", "في يوم", c)
    button(m, "btnShow", "عرض", cm(7.0), y, "primary", w=cm(2.4), h=cm(0.8), call="AgingRefresh Me")
    m.add(Control("label", "lblTotals", cm(0.4), cm(3.3), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("list", "lstParties", cm(0.4), cm(4.0), width - cm(0.8), cm(6.4),
                  {"ColumnCount": 10, "ColumnWidths": "0;6;2.6;2.6;2.4;2.4;2.4;2.4;2.4;2", "ColumnHeads": True},
                  events=["AfterUpdate", "DblClick"]))
    m.add(Control("label", "lblInfo", cm(0.4), cm(10.5), width - cm(0.8), cm(0.55),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstDocs", cm(0.4), cm(11.2), width - cm(0.8), cm(4.8),
                  {"ColumnCount": 7, "ColumnWidths": "0;3;3.2;2.6;2.6;2.4;3", "ColumnHeads": True}))
    x, y = cm(0.4), cm(16.6)
    for name, caption, style, w, call in [
            ("btnPrint", "طباعة", "primary", 2.4, "PrintAging Me"),
            ("btnAllocate", "ربط السداد بالفواتير", "secondary", 4.0, "AgingOpenAllocation Me"),
            ("btnStatement", "كشف حساب", "secondary", 2.8, "AgingStatement Me")]:
        button(m, name, caption, x, y, style, w=cm(w), h=cm(0.9), call=call)
        x += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), y, "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    AgingLoad Me", "End Sub",
               "Private Sub lstParties_AfterUpdate()", "    AgingPartyChanged Me", "End Sub",
               "Private Sub lstParties_DblClick(Cancel As Integer)", "    AgingOpenAllocation Me", "End Sub"] + m.code)
    return m


def layout_allocation() -> FormModel:
    width, height = cm(26.0), cm(18.4)
    m = FormModel("frmAllocation", "ربط السداد بالفواتير", width, height, popup=True, allow_add=False)
    title_band(m, "ربط السداد بالفواتير",
               "اختر السند ثم الفاتورة التي يسددها. ما لا يُربط يسدد أقدم الفواتير استحقاقًا", "customers")
    y = cm(2.3)
    _kind_combo(m, cm(0.4), y, "AllocationKindChanged")
    c = m.add(Control("combo", "cboParty", cm(3.8), y, cm(7.0), cm(0.8),
                      {"RowSourceType": "Table/Query", "RowSource": "", "ColumnCount": 2, "ColumnWidths": "0;6.8",
                       "LimitToList": True}, events=["AfterUpdate"]))
    labelled(m, "cboParty", "العميل / المورد", c)
    button(m, "btnAuto", "ربط تلقائي بالأقدم", cm(11.0), y, "secondary", w=cm(3.6), h=cm(0.8),
           call="DoAutoAllocate Me")
    half = (width - cm(1.0)) // 2
    m.add(Control("label", "lblPayCap", cm(0.4), cm(3.4), half, cm(0.55),
                  {"Caption": "السندات", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_TEXT")}))
    m.add(Control("list", "lstPayments", cm(0.4), cm(4.0), half, cm(6.4),
                  {"ColumnCount": 6, "ColumnWidths": "0;2.4;2.2;2.2;2.2;2.2", "ColumnHeads": True},
                  events=["AfterUpdate"]))
    m.add(Control("label", "lblInvCap", cm(0.6) + half, cm(3.4), half, cm(0.55),
                  {"Caption": "الفواتير المفتوحة", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_TEXT")}))
    m.add(Control("list", "lstInvoices", cm(0.6) + half, cm(4.0), half, cm(6.4),
                  {"ColumnCount": 7, "ColumnWidths": "0;2.2;1.9;1.9;1.9;2;1.9", "ColumnHeads": True},
                  events=["AfterUpdate"]))
    y = cm(11.2)
    c = m.add(Control("text", "txtAmount", cm(0.4), y, cm(3.2), cm(0.8), {"Format": "#,##0.00"}))
    labelled(m, "txtAmount", "المبلغ المربوط", c)
    button(m, "btnAllocate", "ربط بالفاتورة", cm(3.8), y, "primary", w=cm(3.2), h=cm(0.8), call="DoAllocate Me")
    m.add(Control("label", "lblAllocCap", cm(0.4), cm(12.3), cm(12.0), cm(0.55),
                  {"Caption": "الربط المسجل", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_TEXT")}))
    m.add(Control("list", "lstAllocations", cm(0.4), cm(12.9), width - cm(0.8), cm(3.6),
                  {"ColumnCount": 5, "ColumnWidths": "0;3;3;3;3", "ColumnHeads": True}))
    button(m, "btnRemove", "إلغاء الربط", cm(0.4), cm(16.8), "danger", w=cm(3.0), h=cm(0.9),
           call="DoRemoveAllocation Me")
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(16.8), "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    AllocationLoad Me", "End Sub",
               "Private Sub cboParty_AfterUpdate()", "    AllocationRefresh Me", "End Sub",
               "Private Sub lstPayments_AfterUpdate()", "    AllocationPicked Me", "End Sub",
               "Private Sub lstInvoices_AfterUpdate()", "    AllocationPicked Me", "End Sub"] + m.code)
    return m


def aging_forms() -> List[FormModel]:
    return [layout_aging(), layout_allocation()]
