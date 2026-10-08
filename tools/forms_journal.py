"""Journal screens: review of the entries of a period (frmJournal) and one entry
(frmJournalEntry). The logic is in modJournal.
Layout conventions are the same as forms.py (twips, x from the start edge)."""

from typing import List

from typing import Tuple

from forms import Control, FormModel, Sym, button, cm, currency_pair, fit_window, labelled, shrink_area, title_band
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
            ("btnSync", "تحديث القيود", "secondary", 2.6, "JournalSync Me"),
            ("btnPrint", "طباعة اليومية", "secondary", 2.6, 'PrintJournal Me, "JOURNAL"'),
            ("btnTrial", "ميزان المراجعة", "secondary", 2.8, 'PrintJournal Me, "TRIAL"'),
            ("btnAccounts", "دليل الحسابات", "secondary", 2.8, 'OpenScreen "frmAccounts"'),
            ("btnLedger", "كشف حساب", "secondary", 2.4, 'OpenScreen "frmLedger"'),
            ("btnManual", "قيد يدوي", "secondary", 2.2, 'OpenScreen "frmManualEntry"')]:
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
POSTING_ACCOUNTS = ("SELECT a.AccountCode, a.AccountCode & '  ' & a.AccountName AS Account FROM [@Accounts] AS a "
                    "WHERE IsPosting = True AND IsActive = True ORDER BY TreeKey")
MANUAL_FIND_ROWS = ("SELECT ManualEntryID, EntryNumber, EntryDate, Description FROM ManualEntries "
                    "ORDER BY EntryDate DESC, ManualEntryID DESC")
MANUAL_TITLES = ["#", "الحساب", "مدين", "دائن", "بيان السطر", "مركز التكلفة", ""]


CENTER_ROWS = "SELECT CostCenterID, CenterName FROM CostCenters WHERE IsActive = True ORDER BY CenterCode"


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
        ("LineText", "LineText", 5.6, {}, ["AfterUpdate"]),
        ("LineCenter", "LineCenter", 3.1, {"_kind": "combo", "RowSource": CENTER_ROWS, "ColumnCount": 2,
                                           "ColumnWidths": "0;3", "BoundColumn": 1, "LimitToList": True}, []),
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
    c = m.add(Control("text", "txtDescription", cm(11.6), y, width - cm(17.4), cm(0.8), {}))
    labelled(m, "txtDescription", "البيان", c)
    currency_pair(m, width - cm(5.6), y, cbo_w=cm(2.6), rate_w=cm(2.4), call="ManualCurrencyPicked Me")
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
    fit_window(m, split_x=cm(21.3), bottom_y=cm(13.0),   # the currency pair follows the right edge
               stretch_w=("subLines", "lblTotals", "lblStatus", "txtDescription"), stretch_h=("subLines",))
    return m


# ------------------------------------------------------------- account statement
LEDGER_ACCOUNTS = ("SELECT a.AccountCode, a.AccountCode & '  ' & Space((a.AccountLevel - 1) * 2) & a.AccountName "
                   "AS Account FROM [@Accounts] AS a ORDER BY a.TreeKey")
LEDGER_LIST_WIDTHS = "0;2.2;2.4;3.0;2.6;5.6;3.0;2.4;2.4;2.6"


def layout_ledger() -> FormModel:
    width, height = cm(27.0), cm(17.0)
    m = FormModel("frmLedger", "كشف حساب", width, height, popup=False, allow_add=False)
    title_band(m, "كشف حساب ودفتر الأستاذ",
               "حركة أي حساب برصيد أول المدة والرصيد بعد كل قيد؛ الحساب الرئيسي يشمل حساباته التابعة", "journal")
    y = cm(2.3)
    c = m.add(Control("combo", "cboAccount", cm(0.4), y, cm(9.0), cm(0.8),
                      {"RowSource": LEDGER_ACCOUNTS, "ColumnCount": 2, "ColumnWidths": "0;9", "LimitToList": True},
                      events=["AfterUpdate"]))
    labelled(m, "cboAccount", "الحساب", c)
    c = m.add(Control("text", "txtFrom", cm(9.6), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtFrom", "من تاريخ", c)
    c = m.add(Control("text", "txtTo", cm(12.8), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtTo", "إلى تاريخ", c)
    for i, (name, caption, which) in enumerate([("btnThisMonth", "هذا الشهر", "MONTH"),
                                                ("btnLastMonth", "الشهر الماضي", "LASTMONTH"),
                                                ("btnThisYear", "هذه السنة", "YEAR")]):
        button(m, name, caption, cm(16.0) + i * cm(2.25), y, "secondary", w=cm(2.1), h=cm(0.8),
               call=f'LedgerQuickPeriod Me, "{which}"')
    button(m, "btnShow", "عرض", cm(23.0), y, "primary", w=cm(2.5), h=cm(0.8), call="LedgerRefresh Me")
    # the four figures of the statement
    part = (width - cm(0.8)) // 4
    for i, (key, caption) in enumerate([("Opening", "رصيد أول المدة"), ("Debit", "مدين الفترة"),
                                        ("Credit", "دائن الفترة"), ("Closing", "الرصيد الختامي")]):
        x = cm(0.4) + i * part
        m.add(Control("label", f"lblCap{key}", x, cm(3.45), part - cm(0.2), cm(0.5),
                      {"Caption": caption, "FontSize": 9, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
        m.add(Control("label", f"lbl{key}", x, cm(3.95), part - cm(0.2), cm(0.75),
                      {"Caption": "-", "FontSize": 14, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("list", "lstLines", cm(0.4), cm(4.95), width - cm(0.8), cm(8.6),
                  {"ColumnCount": 10, "ColumnWidths": LEDGER_LIST_WIDTHS, "ColumnHeads": True},
                  events=["DblClick"]))
    m.add(Control("label", "lblInfo", cm(0.4), cm(13.7), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    x, y = cm(0.4), cm(14.6)
    for name, caption, style, w, call in [
            ("btnOpenEntry", "فتح القيد", "primary", 2.6, "LedgerOpenEntry Me"),
            ("btnOpenSource", "فتح أصل العملية", "primary", 3.6, "LedgerOpenSource Me"),
            ("btnPrintStatement", "طباعة كشف الحساب", "secondary", 3.8, 'PrintLedger Me, "STATEMENT"'),
            ("btnPrintLedger", "دفتر الأستاذ", "secondary", 2.8, 'PrintLedger Me, "LEDGER"'),
            ("btnManual", "قيد يدوي", "secondary", 2.4, 'OpenScreen "frmManualEntry"'),
            ("btnAccounts", "دليل الحسابات", "secondary", 2.8, 'OpenScreen "frmAccounts"'),
            ("btnFinancials", "القوائم المالية", "secondary", 3.2, 'OpenScreen "frmFinancials"')]:
        button(m, name, caption, x, y, style, w=cm(w), h=cm(0.9), call=call)
        x += cm(w) + cm(0.2)
    button(m, "btnClose", "رجوع", width - cm(0.4) - cm(2.4), y, "secondary", w=cm(2.4), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    LedgerLoad Me", "End Sub",
               "Private Sub cboAccount_AfterUpdate()", "    LedgerRefresh Me", "End Sub",
               "Private Sub lstLines_DblClick(Cancel As Integer)", "    LedgerOpenEntry Me", "End Sub"] + m.code)
    shrink_area(m, ("lstLines",), cm(1.6))
    fit_window(m, split_x=cm(22.9), bottom_y=cm(11.0), stretch_w=("lstLines", "lblInfo"), stretch_h=("lstLines",))
    return m


# ------------------------------------------------------------- financial statements
STATEMENT_KINDS = "INCOME;قائمة الدخل;BALANCE;الميزانية العمومية"
FIN_LIST_WIDTHS = "0;10.6;3.8;3.8;3.8;4.2"


def layout_financials() -> FormModel:
    width, height = cm(27.0), cm(17.0)
    m = FormModel("frmFinancials", "القوائم المالية", width, height, popup=False, allow_add=False)
    title_band(m, "القوائم المالية", "قائمة الدخل والميزانية العمومية من القيود، مع فترة المقارنة", "reports")
    y = cm(2.3)
    c = m.add(Control("combo", "cboStatement", cm(0.4), y, cm(4.6), cm(0.8),
                      {"RowSource": STATEMENT_KINDS, "RowSourceType": "Value List", "ColumnCount": 2,
                       "ColumnWidths": "0;4.5", "LimitToList": True}, events=["AfterUpdate"]))
    labelled(m, "cboStatement", "القائمة", c)
    c = m.add(Control("text", "txtFrom", cm(5.2), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtFrom", "من تاريخ", c)
    c = m.add(Control("text", "txtTo", cm(8.4), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtTo", "إلى تاريخ (الميزانية في هذا اليوم)", c)
    for i, (name, caption, which) in enumerate([("btnThisMonth", "هذا الشهر", "MONTH"),
                                                ("btnLastMonth", "الشهر الماضي", "LASTMONTH"),
                                                ("btnThisYear", "هذه السنة", "YEAR"),
                                                ("btnLastYear", "السنة الماضية", "LASTYEAR")]):
        button(m, name, caption, cm(11.6) + i * cm(2.6), y, "secondary", w=cm(2.45), h=cm(0.8),
               call=f'FinancialsQuickPeriod Me, "{which}"')
    button(m, "btnShow", "عرض", cm(22.2), y, "primary", w=cm(2.5), h=cm(0.8), call="FinancialsRefresh Me")
    m.add(Control("label", "lblCompare", cm(0.4), cm(3.3), width - cm(0.8), cm(0.55),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    part = (width - cm(0.8)) // 3
    for i in range(1, 4):
        x = cm(0.4) + (i - 1) * part
        m.add(Control("label", f"lblCap{i}", x, cm(3.95), part - cm(0.2), cm(0.5),
                      {"Caption": " ", "FontSize": 9, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
        m.add(Control("label", f"lblVal{i}", x, cm(4.5), part - cm(0.2), cm(0.75),
                      {"Caption": "-", "FontSize": 14, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("list", "lstRows", cm(0.4), cm(5.45), width - cm(0.8), cm(8.4),
                  {"ColumnCount": 6, "ColumnWidths": FIN_LIST_WIDTHS, "ColumnHeads": True}, events=["DblClick"]))
    m.add(Control("label", "lblInfo", cm(0.4), cm(14.0), width - cm(0.8), cm(0.55),
                  {"Caption": "نقر مزدوج على حساب يفتح كشف حسابه للفترة نفسها", "FontSize": 9,
                   "ForeColor": Sym("CLR_MUTED")}))
    x, y = cm(0.4), cm(14.8)
    for name, caption, style, w, call in [
            ("btnPrint", "طباعة القائمة", "primary", 3.2, "PrintFinancials Me"),
            ("btnLedger", "كشف حساب", "secondary", 2.8, "FinancialsOpenLedger Me"),
            ("btnTrial", "ميزان المراجعة", "secondary", 2.8, 'OpenScreen "frmJournal"'),
            ("btnClosing", "إقفال الفترات", "secondary", 2.7, 'OpenScreen "frmPeriodClosing", 0'),
            ("btnVat", "الإقرار الضريبي", "secondary", 2.9, 'OpenScreen "frmVatReturn", 0'),
            ("btnAssets", "الأصول الثابتة", "secondary", 3.0, 'OpenScreen "frmAssets", 0'),
            ("btnPayroll", "الرواتب", "secondary", 2.2, 'OpenScreen "frmPayroll"'),
            ("btnBudget", "الموازنة", "secondary", 2.2, 'OpenScreen "frmBudget"')]:
        button(m, name, caption, x, y, style, w=cm(w), h=cm(0.9), call=call)
        x += cm(w) + cm(0.2)
    button(m, "btnClose", "رجوع", width - cm(0.4) - cm(2.4), y, "secondary", w=cm(2.4), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    FinancialsLoad Me", "End Sub",
               "Private Sub cboStatement_AfterUpdate()", "    FinancialsRefresh Me", "End Sub",
               "Private Sub lstRows_DblClick(Cancel As Integer)", "    FinancialsOpenLedger Me", "End Sub"] + m.code)
    shrink_area(m, ("lstRows",), cm(1.6))
    fit_window(m, split_x=cm(22.1), bottom_y=cm(11.0), stretch_w=("lstRows", "lblInfo", "lblCompare"),
               stretch_h=("lstRows",))
    return m


# ------------------------------------------------------------- period closing
CLOSING_HISTORY = ("SELECT p.PeriodClosingID, IIf(p.ActionType = 'CLOSE', 'إقفال فترة', IIf(p.ActionType = 'REOPEN', "
                   "'إعادة فتح', IIf(p.ActionType = 'YEAR_CLOSE', 'إقفال سنة', 'إعادة فتح سنة'))) AS [العملية], "
                   "p.ClosedThrough AS [مقفلة حتى], p.FiscalYear AS [السنة], e.EmployeeName AS [بواسطة], "
                   "p.CreatedAt AS [في], p.Notes AS [السبب] FROM PeriodClosings AS p INNER JOIN Employees AS e ON "
                   "p.EmployeeID = e.EmployeeID ORDER BY p.PeriodClosingID DESC")


def layout_period_closing() -> FormModel:
    width, height = cm(22.0), cm(17.0)
    m = FormModel("frmPeriodClosing", "إقفال الفترات والسنة المالية", width, height, popup=True, allow_add=False)
    title_band(m, "إقفال الفترات والسنة المالية",
               "بعد الإقفال لا يُضاف ولا يُعدَّل ولا يُحذف أي مستند بتاريخ مقفل", "journal")
    m.add(Control("label", "lblState", cm(0.4), cm(1.85), width - cm(0.8), cm(0.8),
                  {"Caption": " ", "FontSize": 14, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    # a period
    m.add(Control("label", "lblPeriodCap", cm(0.4), cm(2.85), cm(12.0), cm(0.6),
                  {"Caption": "إقفال فترة (شهر أو أكثر)", "FontSize": 11, "FontBold": True,
                   "ForeColor": Sym("CLR_TEXT")}))
    c = m.add(Control("text", "txtThrough", cm(0.4), cm(4.0), cm(3.4), cm(0.85), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtThrough", "مقفلة حتى يوم", c)
    c = m.add(Control("text", "txtNotes", cm(4.0), cm(4.0), width - cm(4.4), cm(0.85), {}))
    labelled(m, "txtNotes", "السبب / ملاحظات (مطلوب لإعادة الفتح)", c)
    button(m, "btnClosePeriod", "إقفال حتى هذا اليوم", cm(0.4), cm(5.1), "primary", w=cm(4.6), h=cm(0.9),
           call="DoClosePeriod Me")
    button(m, "btnReopenPeriod", "إعادة الفتح إلى هذا اليوم", cm(5.2), cm(5.1), "danger", w=cm(5.2), h=cm(0.9),
           call="DoReopenPeriod Me")
    # a fiscal year
    m.add(Control("label", "lblYearCap", cm(0.4), cm(6.4), cm(12.0), cm(0.6),
                  {"Caption": "إقفال السنة المالية (الإيرادات والمصروفات إلى الأرباح المحتجزة)", "FontSize": 11,
                   "FontBold": True, "ForeColor": Sym("CLR_TEXT")}))
    c = m.add(Control("combo", "cboYear", cm(0.4), cm(7.55), cm(3.4), cm(0.85),
                      {"RowSourceType": "Value List", "RowSource": "", "ColumnCount": 1, "ColumnWidths": "3",
                       "LimitToList": True}, events=["AfterUpdate"]))
    labelled(m, "cboYear", "السنة", c)
    m.add(Control("label", "lblYearInfo", cm(4.0), cm(7.6), width - cm(4.4), cm(0.75),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    button(m, "btnCloseYear", "إقفال السنة", cm(0.4), cm(8.65), "primary", w=cm(4.6), h=cm(0.9), call="DoCloseYear Me")
    button(m, "btnReopenYear", "إعادة فتح السنة", cm(5.2), cm(8.65), "danger", w=cm(5.2), h=cm(0.9),
           call="DoReopenYear Me")
    m.add(Control("label", "lblHistoryCap", cm(0.4), cm(9.95), cm(12.0), cm(0.6),
                  {"Caption": "سجل الإقفال وإعادة الفتح", "FontSize": 10, "FontBold": True,
                   "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstHistory", cm(0.4), cm(10.6), width - cm(0.8), cm(4.6),
                  {"RowSourceType": "Table/Query", "RowSource": CLOSING_HISTORY, "ColumnCount": 7,
                   "ColumnWidths": "0;2.8;2.6;1.4;3.4;3.6;6", "ColumnHeads": True}))
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(15.6), "secondary", w=cm(2.6), h=cm(1.0),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    PeriodClosingLoad Me", "End Sub",
               "Private Sub cboYear_AfterUpdate()", "    PeriodClosingRefresh Me", "End Sub"] + m.code)
    return m


# ------------------------------------------------------------------ VAT return
VAT_HISTORY = ("SELECT VatReturnID, ReturnNumber AS [الإقرار], Format(PeriodFrom, 'yyyy/mm/dd') AS [من], "
               "Format(PeriodTo, 'yyyy/mm/dd') AS [إلى], "
               "IIf(Status = 'FILED', 'معتمد', 'مسودة') AS [الحالة], Format(NetDue, '#,##0.00') AS [الصافي], "
               "Format(PaidAmount, '#,##0.00') AS [المسدد], Format(FiledDate, 'yyyy/mm/dd') AS [اعتُمد في] FROM VatReturns "
               "ORDER BY PeriodFrom DESC")


def layout_vat_return() -> FormModel:
    width, height = cm(26.0), cm(19.4)
    m = FormModel("frmVatReturn", "إقرار ضريبة القيمة المضافة", width, height, popup=True, allow_add=False)
    title_band(m, "إقرار ضريبة القيمة المضافة",
               "خانات نموذج هيئة الزكاة والضريبة والجمارك من المستندات، ثم الاعتماد وقيد التسوية والسداد", "reports")
    y = cm(2.3)
    c = m.add(Control("text", "txtFrom", cm(0.4), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtFrom", "بداية الفترة الضريبية", c)
    c = m.add(Control("text", "txtTo", cm(3.6), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtTo", "نهاية الفترة", c)
    button(m, "btnLastMonth", "الشهر الماضي", cm(6.8), y, "secondary", w=cm(2.6), h=cm(0.8),
           call='VatQuickPeriod Me, "LASTMONTH"')
    button(m, "btnLastQuarter", "الربع الماضي", cm(9.6), y, "secondary", w=cm(2.6), h=cm(0.8),
           call='VatQuickPeriod Me, "LASTQUARTER"')
    button(m, "btnCalc", "احسب", cm(12.4), y, "primary", w=cm(2.4), h=cm(0.8), call="VatCalculate Me")
    m.add(Control("text", "txtReturnID", cm(15.0), y, cm(1.0), cm(0.8), {"Visible": False}))
    m.add(Control("label", "lblState", cm(0.4), cm(3.25), width - cm(0.8), cm(1.0),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("list", "lstBoxes", cm(0.4), cm(4.35), width - cm(0.8), cm(6.9),
                  {"RowSourceType": "Value List", "RowSource": "", "ColumnCount": 5,
                   "ColumnWidths": "1.2;13.4;3.4;3.4;3.4", "ColumnHeads": True}))
    # boxes 14 and 15, box 16
    y = cm(12.0)
    c = m.add(Control("text", "txtCorrections", cm(0.4), y, cm(3.4), cm(0.8), {"Format": "#,##0.00"},
                      events=["AfterUpdate"]))
    labelled(m, "txtCorrections", "14- تصحيحات سابقة (+/-)", c)
    c = m.add(Control("text", "txtCarried", cm(4.0), y, cm(3.4), cm(0.8), {"Format": "#,##0.00"},
                      events=["AfterUpdate"]))
    labelled(m, "txtCarried", "15- رصيد دائن مرحَّل", c)
    m.add(Control("label", "lblNetDue", cm(7.6), y, cm(10.0), cm(0.8),
                  {"Caption": " ", "FontSize": 14, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    button(m, "btnSaveDraft", "حفظ مسودة", cm(17.8), y, "secondary", w=cm(2.6), h=cm(0.8), call="VatSaveDraft Me")
    button(m, "btnPrint", "طباعة", cm(20.6), y, "secondary", w=cm(2.2), h=cm(0.8), call="PrintVatReturn Me")
    button(m, "btnDeleteDraft", "حذف المسودة", cm(23.0), y, "danger", w=cm(2.6), h=cm(0.8), call="VatDeleteDraft Me")
    # filing
    y = cm(13.55)
    c = m.add(Control("text", "txtFiledDate", cm(0.4), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtFiledDate", "تاريخ الاعتماد", c)
    c = m.add(Control("text", "txtFilingRef", cm(3.6), y, cm(3.6), cm(0.8), {}))
    labelled(m, "txtFilingRef", "رقم الإقرار لدى الهيئة", c)
    c = m.add(Control("text", "txtNotes", cm(7.4), y, cm(5.2), cm(0.8), {}))
    labelled(m, "txtNotes", "سبب إلغاء الاعتماد", c)
    button(m, "btnFile", "اعتماد الإقرار", cm(12.8), y, "primary", w=cm(3.0), h=cm(0.8), call="VatFile Me")
    button(m, "btnUnfile", "إلغاء الاعتماد", cm(16.0), y, "danger", w=cm(2.8), h=cm(0.8), call="VatUnfile Me")
    button(m, "btnEntry", "قيد التسوية", cm(19.0), y, "secondary", w=cm(2.8), h=cm(0.8), call="VatOpenEntry Me")
    # payment
    y = cm(15.1)
    c = m.add(Control("text", "txtPaidDate", cm(0.4), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtPaidDate", "تاريخ السداد", c)
    c = m.add(Control("text", "txtPaidAmount", cm(3.6), y, cm(3.0), cm(0.8), {"Format": "#,##0.00"}))
    labelled(m, "txtPaidAmount", "المبلغ المسدد", c)
    c = m.add(Control("combo", "cboPayAccount", cm(6.8), y, cm(5.8), cm(0.8),
                      {"RowSourceType": "Table/Query", "RowSource": "", "ColumnCount": 2, "ColumnWidths": "0;5.6",
                       "LimitToList": True}))
    labelled(m, "cboPayAccount", "سُدِّدت من حساب", c)
    button(m, "btnPay", "تسجيل السداد", cm(12.8), y, "primary", w=cm(3.0), h=cm(0.8), call="VatPay Me")
    button(m, "btnUnpay", "إلغاء السداد", cm(16.0), y, "danger", w=cm(2.8), h=cm(0.8), call="VatUnpay Me")
    # the returns
    m.add(Control("label", "lblHistoryCap", cm(0.4), cm(16.1), cm(12.0), cm(0.55),
                  {"Caption": "الإقرارات المحفوظة (اختر إقرارًا لعرضه)", "FontSize": 9, "FontBold": True,
                   "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstReturns", cm(0.4), cm(16.7), width - cm(3.4), cm(2.4),
                  {"RowSourceType": "Table/Query", "RowSource": VAT_HISTORY, "ColumnCount": 8,
                   "ColumnWidths": "0;3.2;2.6;2.6;2;3;3;2.6", "ColumnHeads": True}, events=["AfterUpdate"]))
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(18.1), "secondary", w=cm(2.6), h=cm(1.0),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    VatReturnLoad Me", "End Sub",
               "Private Sub txtCorrections_AfterUpdate()", "    VatShowNet Me", "End Sub",
               "Private Sub txtCarried_AfterUpdate()", "    VatShowNet Me", "End Sub",
               "Private Sub lstReturns_AfterUpdate()", "    VatPickReturn Me", "End Sub"] + m.code)
    return m


def journal_forms() -> List[FormModel]:
    lines, heads = layout_manual_lines()
    return [layout_journal(), layout_journal_entry(), lines, layout_manual_entry(heads), layout_ledger(),
            layout_financials(), layout_period_closing(), layout_vat_return()]
