"""The office tasks screen (frmOfficeSync, plan 6H, docs/84): Access stays in the office next to the web, the
Windows POS and the rep app on the same database. One Access computer brings the journal up to date with the
documents made outside Access and sends their pending e-invoices, by itself every few minutes (the dashboard
timer) or by a button. Logic: modOfficeSync."""

from typing import List

from forms import Control, FormModel, Sym, button, cm, labelled, title_band


def layout_office_sync() -> FormModel:
    width, height = cm(22.0), cm(15.6)
    m = FormModel("frmOfficeSync", "مهام المكتب", width, height, popup=True, allow_add=False)
    title_band(m, "مهام المكتب", "تحديث القيود وإرسال الفواتير الإلكترونية لمستندات الويب ونقاط البيع "
               "وتطبيق المندوب", "backup")
    m.add(Control("label", "lblThisComputer", cm(0.4), cm(1.8), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_TEXT")}))
    y = cm(3.1)
    c = m.add(Control("text", "txtComputer", cm(0.4), y, cm(6.0), cm(0.8), {}))
    labelled(m, "txtComputer", "جهاز مهام المكتب (فارغ = يدويًا فقط)", c)
    button(m, "btnThisComputer", "هذا الجهاز", cm(6.6), y, "secondary", w=cm(2.8), h=cm(0.8),
           call="OfficeSyncUseThisComputer Me")
    c = m.add(Control("text", "txtMinutes", cm(9.8), y, cm(3.0), cm(0.8), {"Format": "0"}))
    labelled(m, "txtMinutes", "كل (دقيقة)", c)
    button(m, "btnSave", "حفظ الإعداد", cm(13.0), y, "secondary", w=cm(3.0), h=cm(0.8), call="OfficeSyncSave Me")
    m.add(Control("label", "lblState", cm(0.4), cm(4.2), width - cm(0.8), cm(0.65),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("label", "lblWaiting", cm(0.4), cm(4.9), width - cm(0.8), cm(0.65),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    button(m, "btnRunNow", "تشغيل الآن", cm(0.4), cm(5.75), "primary", w=cm(3.4), h=cm(0.9), call="OfficeSyncRunNow Me")
    button(m, "btnRefresh", "تحديث", cm(4.0), cm(5.75), "secondary", w=cm(2.6), h=cm(0.9), call="OfficeSyncShow Me")
    button(m, "btnEInvoices", "الفاتورة الإلكترونية", cm(6.8), cm(5.75), "secondary", w=cm(4.0), h=cm(0.9),
           call='OpenScreen "frmEInvoices", 0')
    m.add(Control("list", "lstLog", cm(0.4), cm(7.0), width - cm(0.8), cm(7.2),
                  {"RowSourceType": "Table/Query", "RowSource": "", "ColumnCount": 9,
                   "ColumnWidths": "3.2;2.6;1.6;1.5;1.5;1.5;1.6;1.6;6", "ColumnHeads": True}))
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(14.45), "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = ["Private Sub Form_Load()", "    OfficeSyncLoad Me", "End Sub"] + m.code
    return m


def office_forms() -> List[FormModel]:
    return [layout_office_sync()]
