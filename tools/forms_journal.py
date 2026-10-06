"""Journal screens: review of the entries of a period (frmJournal) and one entry
(frmJournalEntry). The logic is in modJournal.
Layout conventions are the same as forms.py (twips, x from the start edge)."""

from typing import List

from typing import Tuple

from forms import Control, FormModel, Sym, button, cm, fit_window, labelled, shrink_area, title_band
from forms_sales import LOCKED, grid_row, header_labels

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
            ("btnAccounts", "دليل الحسابات", "secondary", 3.0, 'OpenScreen "frmAccounts"'),
            ("btnManual", "قيد يدوي", "secondary", 2.4, 'OpenScreen "frmManualEntry"')]:
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


# ------------------------------------------------------------- manual entries
POSTING_ACCOUNTS = ("SELECT AccountCode, AccountCode & '  ' & AccountName AS Account FROM Accounts "
                    "WHERE IsPosting = True AND IsActive = True ORDER BY TreeKey")
MANUAL_FIND_ROWS = ("SELECT ManualEntryID, EntryNumber, EntryDate, Description FROM ManualEntries "
                    "ORDER BY EntryDate DESC, ManualEntryID DESC")
MANUAL_TITLES = ["#", "الحساب", "مدين", "دائن", "بيان السطر", ""]


def layout_manual_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.75)
    m = FormModel("frmManualLines", "أسطر القيد اليدوي", cm(26.2), row_h, popup=False,
                  record_source="SELECT * FROM tmpManualLines ORDER BY LineNo", allow_add=True)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0}
    heads = grid_row(m, [
        ("LineNo", "LineNo", 0.9, dict(LOCKED), []),
        ("AccountCode", "AccountCode", 9.3, {"_kind": "combo", "RowSource": POSTING_ACCOUNTS, "ColumnCount": 2,
                                             "ColumnWidths": "0;9", "BoundColumn": 1, "LimitToList": True},
         ["AfterUpdate"]),
        ("Debit", "Debit", 3.0, {"Format": "#,##0.00"}, ["AfterUpdate"]),
        ("Credit", "Credit", 3.0, {"Format": "#,##0.00"}, ["AfterUpdate"]),
        ("LineText", "LineText", 8.7, {}, ["AfterUpdate"]),
        ("btnRemove", "", 0.8, {"_kind": "button", "Caption": Sym(f"ChrW(&H{0xE74D:X})"),
                               "Style": "danger", "FontName": Sym("ICON_FONT")}, ["Click"]),
    ], row_h)
    m.code = []
    for f in ("AccountCode", "Debit", "Credit", "LineText"):
        m.code += [f"Private Sub {f}_AfterUpdate()", f'    ManualLineChanged Me, "{f}"', "End Sub"]
    m.code += ["Private Sub btnRemove_Click()", "    ManualRemoveLine Me", "End Sub"]
    return m, heads


def layout_manual_entry(heads) -> FormModel:
    width, height = cm(27.0), cm(17.0)
    m = FormModel("frmManualEntry", "القيود اليدوية", width, height, popup=False, allow_add=False)
    title_band(m, "القيود اليدوية", "قيد يكتبه المحاسب: المدين = الدائن، ويُرحَّل لقيود اليومية عند الحفظ",
               "journal")
    y = cm(2.3)
    c = m.add(Control("combo", "cboFind", cm(0.4), y, cm(12.0), cm(0.8),
                      {"RowSource": MANUAL_FIND_ROWS, "ColumnCount": 4, "ColumnWidths": "0;2.6;2.6;8",
                       "LimitToList": True}, events=["AfterUpdate"]))
    labelled(m, "cboFind", "فتح قيد محفوظ", c)
    m.add(Control("text", "txtEntryID", cm(12.6), y, cm(0.5), cm(0.8), {"Visible": False}))
    m.add(Control("text", "txtReversalOf", cm(13.3), y, cm(0.5), cm(0.8), {"Visible": False}))
    y = cm(3.7)
    c = m.add(Control("text", "txtNumber", cm(0.4), y, cm(3.2), cm(0.8), {**LOCKED, "FontBold": True}))
    labelled(m, "txtNumber", "رقم القيد", c)
    c = m.add(Control("text", "txtDate", cm(3.8), y, cm(3.2), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtDate", "التاريخ", c)
    c = m.add(Control("text", "txtReference", cm(7.2), y, cm(4.2), cm(0.8), {}))
    labelled(m, "txtReference", "المرجع (اختياري)", c)
    c = m.add(Control("text", "txtDescription", cm(11.6), y, width - cm(12.0), cm(0.8), {}))
    labelled(m, "txtDescription", "البيان", c)
    header_labels(m, cm(0.4), cm(4.85), MANUAL_TITLES, heads)
    m.add(Control("subform", "subLines", cm(0.4), cm(5.45), cm(26.2), cm(7.6),
                  {"SourceObject": "frmManualLines"}))
    m.add(Control("label", "lblTotals", cm(0.4), cm(13.25), width - cm(0.8), cm(0.65),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("label", "lblStatus", cm(0.4), cm(13.95), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    x, y = cm(0.4), cm(15.4)
    for name, caption, style, w, call in [
            ("btnNew", "قيد جديد", "secondary", 2.6, "ManualNew Me"),
            ("btnSave", "حفظ القيد", "primary", 2.8, "SaveManualEntry Me"),
            ("btnReverse", "قيد عكسي", "secondary", 2.6, "ReverseManualEntry Me"),
            ("btnPrint", "طباعة", "secondary", 2.2, "PrintManualEntry Me"),
            ("btnInJournal", "عرض في اليومية", "secondary", 3.2, "ManualOpenInJournal Me"),
            ("btnAccounts", "دليل الحسابات", "secondary", 3.0, 'OpenScreen "frmAccounts"'),
            ("btnDelete", "حذف القيد", "danger", 2.6, "DeleteManualEntry Me")]:
        button(m, name, caption, x, y, style, w=cm(w), h=cm(0.9), call=call)
        x += cm(w) + cm(0.2)
    button(m, "btnClose", "رجوع", width - cm(0.4) - cm(2.4), y, "secondary", w=cm(2.4), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    ManualEntryLoad Me", "End Sub",
               "Private Sub cboFind_AfterUpdate()", "    ManualFindPicked Me", "End Sub"] + m.code)
    fit_window(m, split_x=cm(24.0), bottom_y=cm(13.0),
               stretch_w=("subLines", "lblTotals", "lblStatus", "txtDescription"), stretch_h=("subLines",))
    return m


def journal_forms() -> List[FormModel]:
    lines, heads = layout_manual_lines()
    return [layout_journal(), layout_journal_entry(), lines, layout_manual_entry(heads)]
