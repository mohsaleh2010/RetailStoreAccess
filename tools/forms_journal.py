"""Journal screens: review of the entries of a period (frmJournal) and one entry
(frmJournalEntry). The logic is in modJournal.
Layout conventions are the same as forms.py (twips, x from the start edge)."""

from typing import List

from forms import Control, FormModel, Sym, button, cm, fit_window, labelled, shrink_area, title_band

LINE_WIDTHS = "1.8;5.0;6.4;2.6;2.6"


def layout_journal() -> FormModel:
    width, height = cm(27.0), cm(17.6)
    m = FormModel("frmJournal", "قيود اليومية", width, height, popup=False, allow_add=False)
    title_band(m, "قيود اليومية", "قيد لكل عملية مربوط بأصلها: افتح القيد أو أصل العملية", "journal")
    y = cm(2.3)
    c = m.add(Control("text", "txtFrom", cm(0.4), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtFrom", "من تاريخ", c)
    c = m.add(Control("text", "txtTo", cm(3.6), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtTo", "إلى تاريخ", c)
    for i, (name, caption, which) in enumerate([("btnToday", "اليوم", "TODAY"), ("btnThisMonth", "هذا الشهر", "MONTH"),
                                                ("btnLastMonth", "الشهر الماضي", "LASTMONTH"),
                                                ("btnThisYear", "هذه السنة", "YEAR")]):
        button(m, name, caption, cm(6.8) + i * cm(2.25), y, "secondary", w=cm(2.1), h=cm(0.8),
               call=f'JournalQuickPeriod Me, "{which}"')
    c = m.add(Control("combo", "cboSourceType", cm(15.9), y, cm(4.2), cm(0.8),
                      {"RowSource": "", "RowSourceType": "Value List", "ColumnCount": 2, "ColumnWidths": "0;4",
                       "LimitToList": True}, events=["AfterUpdate"]))
    labelled(m, "cboSourceType", "نوع العملية", c)
    c = m.add(Control("text", "txtSearch", cm(20.3), y, cm(3.6), cm(0.8), {}, events=["AfterUpdate"]))
    labelled(m, "txtSearch", "بحث (رقم / بيان)", c)
    button(m, "btnShow", "عرض", cm(24.1), y, "primary", w=cm(2.5), h=cm(0.8), call="JournalRefresh Me")

    m.add(Control("list", "lstEntries", cm(0.4), cm(3.4), width - cm(0.8), cm(6.4),
                  {"ColumnCount": 7, "ColumnWidths": "0;2.6;2.4;3.4;2.8;11.6;3.0", "ColumnHeads": True},
                  events=["AfterUpdate", "DblClick"]))
    m.add(Control("label", "lblTotals", cm(0.4), cm(9.9), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("label", "lblLinesCap", cm(0.4), cm(10.6), cm(12), cm(0.55),
                  {"Caption": "أسطر القيد المختار", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstLines", cm(0.4), cm(11.2), width - cm(0.8), cm(4.1),
                  {"ColumnCount": 5, "ColumnWidths": "2.2;6.0;10.6;3.2;3.2", "ColumnHeads": True}))
    y = cm(15.6)
    x = cm(0.4)
    for name, caption, style, w, call in [
            ("btnOpenEntry", "فتح القيد", "primary", 2.8, "JournalOpenEntry Me"),
            ("btnOpenSource", "فتح أصل العملية", "primary", 3.6, "OpenJournalSource Me!lstEntries.Value"),
            ("btnSync", "تحديث القيود", "secondary", 3.0, "JournalSync Me"),
            ("btnPrint", "طباعة اليومية", "secondary", 3.0, 'PrintJournal Me, "JOURNAL"'),
            ("btnTrial", "ميزان المراجعة", "secondary", 3.0, 'PrintJournal Me, "TRIAL"'),
            ("btnAccounts", "دليل الحسابات", "secondary", 3.0, 'OpenScreen "frmAccounts"')]:
        button(m, name, caption, x, y, style, w=cm(w), h=cm(0.9), call=call)
        x += cm(w) + cm(0.2)
    button(m, "btnClose", "رجوع", width - cm(0.4) - cm(2.4), y, "secondary", w=cm(2.4), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.add(Control("label", "lblSync", cm(0.4), cm(16.7), width - cm(0.8), cm(0.55),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    JournalLoad Me", "End Sub",
               "Private Sub lstEntries_AfterUpdate()", "    JournalEntryPicked Me", "End Sub",
               "Private Sub lstEntries_DblClick(Cancel As Integer)", "    JournalOpenEntry Me", "End Sub",
               "Private Sub cboSourceType_AfterUpdate()", "    JournalRefresh Me", "End Sub",
               "Private Sub txtSearch_AfterUpdate()", "    JournalRefresh Me", "End Sub"]
              + m.code)
    shrink_area(m, ("lstEntries",), cm(1.7))
    fit_window(m, split_x=cm(24.0), bottom_y=cm(8.0),
               stretch_w=("lstEntries", "lstLines", "lblTotals", "lblSync"), stretch_h=("lstEntries",))
    return m


def layout_journal_entry() -> FormModel:
    width, height = cm(21.0), cm(11.4)
    m = FormModel("frmJournalEntry", "قيد يومية", width, height, popup=True, allow_add=False)
    title_band(m, "قيد يومية", "القيد الآلي للعملية", "journal")
    m.add(Control("text", "txtEntryID", width - cm(1.0), cm(0.1), cm(0.6), cm(0.4), {"Visible": False},
                  decorative=True))
    m.add(Control("label", "lblHeader", cm(0.4), cm(1.8), width - cm(0.8), cm(0.65),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_TEXT")}))
    m.add(Control("label", "lblDescription", cm(0.4), cm(2.5), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 10, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstLines", cm(0.4), cm(3.3), width - cm(0.8), cm(5.0),
                  {"ColumnCount": 5, "ColumnWidths": "2.0;5.0;6.4;3.0;3.0", "ColumnHeads": True}))
    m.add(Control("label", "lblTotals", cm(0.4), cm(8.45), width - cm(0.8), cm(0.65),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    button(m, "btnOpenSource", "فتح أصل العملية", cm(0.4), cm(9.9), "primary", w=cm(3.8), h=cm(1.0),
           call="OpenJournalSource Me!txtEntryID.Value")
    button(m, "btnPrint", "طباعة القيد", cm(4.4), cm(9.9), "secondary", w=cm(3.0), h=cm(1.0),
           call="PrintJournalEntry Me!txtEntryID.Value")
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(9.9), "secondary", w=cm(2.6), h=cm(1.0),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = ["Private Sub Form_Load()", "    JournalEntryLoad Me", "End Sub"] + m.code
    return m


def journal_forms() -> List[FormModel]:
    return [layout_journal(), layout_journal_entry()]
