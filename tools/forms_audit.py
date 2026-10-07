"""The audit trail screen (frmAuditLog): the operations of a period, filtered by user, kind and table,
and the fields of the chosen operation with their values before and after. The logic is in modAudit.
Layout conventions are the same as forms.py."""

from typing import List

from forms import Control, FormModel, Sym, button, cm, labelled, title_band
from forms_cash import table_combo

ACTIONS = ("ADD;إضافة;EDIT;تعديل;DELETE;حذف;DOCS;عمليات المستندات;LOGIN;الدخول والخروج")
USER_ROWS = "SELECT EmployeeID, EmployeeName FROM Employees ORDER BY EmployeeName"


def layout_audit_log() -> FormModel:
    width, height = cm(27.0), cm(18.4)
    m = FormModel("frmAuditLog", "سجل التدقيق", width, height, popup=True, allow_add=False)
    title_band(m, "سجل التدقيق", "من أضاف أو عدّل أو حذف، ومتى، ومن أي جهاز، والقيم قبل وبعد", "users")
    y = cm(2.3)
    c = m.add(Control("text", "txtFrom", cm(0.4), y, cm(2.8), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtFrom", "من", c)
    c = m.add(Control("text", "txtTo", cm(3.4), y, cm(2.8), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtTo", "إلى", c)
    c = table_combo(m, "cboUser", cm(6.4), y, cm(4.0), events=(), rows=USER_ROWS)
    labelled(m, "cboUser", "المستخدم (فارغ = الكل)", c)
    c = m.add(Control("combo", "cboAction", cm(10.6), y, cm(3.8), cm(0.8),
                      {"RowSourceType": "Value List", "RowSource": ACTIONS, "ColumnCount": 2,
                       "ColumnWidths": "0;3.6", "LimitToList": True}))
    labelled(m, "cboAction", "العملية (فارغ = الكل)", c)
    c = m.add(Control("combo", "cboTable", cm(14.6), y, cm(3.8), cm(0.8),
                      {"RowSourceType": "Table/Query", "RowSource": "", "ColumnCount": 1, "ColumnWidths": "3.6",
                       "LimitToList": True}))
    labelled(m, "cboTable", "الجدول (فارغ = الكل)", c)
    c = m.add(Control("text", "txtSearch", cm(18.6), y, cm(4.6), cm(0.8), {}))
    labelled(m, "txtSearch", "رقم أو اسم السجل", c)
    button(m, "btnShow", "عرض", cm(23.4), y, "primary", w=cm(3.2), h=cm(0.8), call="AuditScreenShow Me")
    m.add(Control("label", "lblCount", cm(0.4), cm(3.3), width - cm(0.8), cm(0.55),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("list", "lstLog", cm(0.4), cm(4.0), width - cm(0.8), cm(6.6),
                  {"ColumnCount": 8, "ColumnWidths": "0;3.2;3.2;3.2;3.2;1.6;5.6;3", "ColumnHeads": True},
                  events=["AfterUpdate"]))
    m.add(Control("label", "lblDetails", cm(0.4), cm(10.7), width - cm(0.8), cm(0.55),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstChanges", cm(0.4), cm(11.4), width - cm(0.8), cm(4.6),
                  {"ColumnCount": 4, "ColumnWidths": "0;5;9.5;9.5", "ColumnHeads": True}))
    y = cm(16.6)
    button(m, "btnPrint", "طباعة الفترة", cm(0.4), y, "secondary", w=cm(3.2), h=cm(0.9), call="PrintAuditLog Me")
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), y, "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    AuditScreenLoad Me", "End Sub",
               "Private Sub lstLog_AfterUpdate()", "    AuditScreenPick Me", "End Sub"] + m.code)
    return m


def audit_forms() -> List[FormModel]:
    return [layout_audit_log()]
