"""Treasury screens: the treasury (boxes, the day of a box), the cash voucher
(cash in / cash out / transfer) and the cashier closing. The logic is in modCash.
Layout conventions are the same as forms.py (twips, x from the start edge)."""

from typing import List

from forms import Control, FormModel, Sym, button, cm, fit_window, labelled, shrink_area, title_band

ACTIVE_EXPENSE_TYPES = ("SELECT ExpenseTypeID, ExpenseTypeName FROM ExpenseTypes WHERE IsActive = True "
                        "ORDER BY ExpenseTypeName")


def value_list_combo(m: FormModel, name, x, y, w, h=cm(0.8), events=("AfterUpdate",), font=None):
    props = {"RowSource": "", "RowSourceType": "Value List", "ColumnCount": 2, "ColumnWidths": "0;7",
             "LimitToList": True}
    if font:
        props["FontSize"] = font
    return m.add(Control("combo", name, x, y, w, h, props, events=list(events)))


def table_combo(m: FormModel, name, x, y, w, h=cm(0.8), events=("AfterUpdate",), rows=""):
    return m.add(Control("combo", name, x, y, w, h,
                         {"RowSource": rows, "ColumnCount": 2, "ColumnWidths": "0;7", "LimitToList": True},
                         events=list(events)))


def figure(m: FormModel, key: str, caption: str, x, y, w, value_font=16):
    """A small figure card: caption on top, value below (lbl<key>Cap / lbl<key>)."""
    m.add(Control("rect", f"box{key}", x, y, w, cm(1.7), {"BackColor": Sym("CLR_SURFACE")}, decorative=True))
    m.add(Control("label", f"lbl{key}Cap", x + cm(0.25), y + cm(0.15), w - cm(0.5), cm(0.5),
                  {"Caption": caption, "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("label", f"lbl{key}", x + cm(0.25), y + cm(0.7), w - cm(0.5), cm(0.85),
                  {"Caption": "-", "FontSize": value_font, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))


# ------------------------------------------------------------------ treasury
def layout_treasury() -> FormModel:
    width, height = cm(27.0), cm(17.6)
    m = FormModel("frmTreasury", "الخزينة", width, height, popup=False, allow_add=False, allow_edit=True)
    title_band(m, "الخزينة والصناديق", "أرصدة الخزينة الرئيسية وصناديق الكاشير وحركة كل يوم", "treasury")
    x = cm(0.4)
    for name, caption, w, call in [
            ("btnCashIn", "سند قبض نقدية", 3.2, 'TreasuryOpen Me, "IN"'),
            ("btnCashOut", "سند صرف نقدية", 3.2, 'TreasuryOpen Me, "OUT"'),
            ("btnTransfer", "تحويل بين الصناديق", 3.6, 'TreasuryOpen Me, "TRANSFER"'),
            ("btnClosing", "تصفية يومية الكاشير", 3.8, 'TreasuryOpen Me, "CLOSING"'),
            ("btnBoxes", "الصناديق", 2.4, 'OpenScreen "frmCashBoxes"'),
            ("btnBanks", "البنوك", 2.4, 'OpenScreen "frmBanks"'),
            ("btnCheques", "الشيكات", 2.4, 'OpenScreen "frmCheques", 0, "IN"')]:
        button(m, name, caption, x, cm(1.8), "primary" if name == "btnClosing" else "secondary",
               w=cm(w), call=call)
        x += cm(w) + cm(0.2)
    button(m, "btnClose", "رجوع", width - cm(0.4) - cm(2.4), cm(1.8), "secondary",
           call="DoCmd.Close acForm, Me.Name")

    m.add(Control("label", "lblBoxesCap", cm(0.4), cm(3.0), cm(9.0), cm(0.55),
                  {"Caption": "الخزينة والصناديق (الرصيد الآن)", "FontSize": 10, "FontBold": True,
                   "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstBoxes", cm(0.4), cm(3.6), cm(9.0), height - cm(4.1),
                  {"ColumnCount": 4, "ColumnWidths": "0;4.2;2.3;2.3", "ColumnHeads": True, "FontSize": 11},
                  events=["AfterUpdate"]))

    px, pw = cm(10.0), width - cm(10.0) - cm(0.4)
    m.add(Control("label", "lblBoxName", px, cm(3.0), pw, cm(0.8),
                  {"Caption": " ", "FontSize": 15, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    y = cm(4.5)
    c = m.add(Control("text", "txtDay", px, y, cm(3.2), cm(0.8), {"Format": "yyyy/mm/dd"},
                      events=["AfterUpdate"]))
    labelled(m, "txtDay", "اليوم", c)
    button(m, "btnPrevDay", "اليوم السابق", px + cm(3.4), y, "secondary", w=cm(2.6), h=cm(0.8),
           call="TreasuryDayStep Me, -1")
    button(m, "btnNextDay", "اليوم التالي", px + cm(6.2), y, "secondary", w=cm(2.6), h=cm(0.8),
           call="TreasuryDayStep Me, 1")
    button(m, "btnToday", "اليوم", px + cm(9.0), y, "secondary", w=cm(1.8), h=cm(0.8),
           call="TreasuryToday Me")
    button(m, "btnPrintDay", "طباعة حركة اليوم", px + cm(11.0), y, "primary", w=cm(3.6), h=cm(0.8),
           call="PrintCashDay Me")

    gap = cm(0.3)
    fw = (pw - 3 * gap) // 4
    for i, (key, caption) in enumerate([("Opening", "رصيد أول اليوم"), ("In", "المقبوضات"),
                                        ("Out", "المدفوعات والمصروفات"), ("Closing", "رصيد آخر اليوم")]):
        figure(m, key, caption, px + i * (fw + gap), cm(5.6), fw)
    m.add(Control("label", "lblCurrent", px, cm(7.5), pw, cm(0.6),
                  {"Caption": " ", "FontSize": 10, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstMoves", px, cm(8.2), pw, height - cm(8.7),
                  {"ColumnCount": 6, "ColumnWidths": "1.5;4.4;2.9;4.4;1.7;1.7", "ColumnHeads": True}))
    m.form_events = ["Load", "Activate"]
    m.code = (["Private Sub Form_Load()", "    TreasuryLoad Me", "End Sub",
               "Private Sub Form_Activate()", "    TreasuryActivate Me", "End Sub",
               "Private Sub lstBoxes_AfterUpdate()", "    TreasuryShow Me", "End Sub",
               "Private Sub txtDay_AfterUpdate()", "    TreasuryShow Me", "End Sub"]
              + m.code)
    shrink_area(m, ("lstBoxes", "lstMoves"), cm(1.7))
    fit_window(m, split_x=cm(24.0), stretch_w=("lblBoxName", "lblCurrent", "lstMoves"),
               stretch_h=("lstBoxes", "lstMoves"))
    return m


# ------------------------------------------------------------------ cash voucher
def layout_cash_voucher() -> FormModel:
    width, height = cm(16.0), cm(12.0)
    m = FormModel("frmCashVoucher", "سند نقدية", width, height, popup=True, allow_add=False)
    title_band(m, "سند نقدية", "قبض نقدية في صندوق، أو صرف منه، أو تحويل بين صندوقين", "treasury")
    half = cm(7.4)
    c = value_list_combo(m, "cboVoucherType", cm(0.4), cm(2.4), half)
    labelled(m, "cboVoucherType", "نوع السند", c)
    c = value_list_combo(m, "cboCategory", cm(8.2), cm(2.4), half)
    labelled(m, "cboCategory", "البند", c)
    c = table_combo(m, "cboBox", cm(0.4), cm(3.9), half)
    labelled(m, "cboBox", "الصندوق", c)
    c = table_combo(m, "cboToBox", cm(8.2), cm(3.9), half, events=())
    labelled(m, "cboToBox", "إلى صندوق", c)
    m.add(Control("label", "lblBoxBalance", cm(0.4), cm(4.8), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    c = table_combo(m, "cboExpenseType", cm(0.4), cm(6.0), cm(5.2), events=(), rows=ACTIVE_EXPENSE_TYPES)
    labelled(m, "cboExpenseType", "نوع المصروف *", c)
    button(m, "btnNewExpenseType", "نوع جديد", cm(5.8), cm(6.0), "secondary", w=cm(2.0), h=cm(0.8),
           call='AddExpenseType Me, "cboExpenseType"')
    c = m.add(Control("text", "txtAmount", cm(8.2), cm(6.0), half, cm(0.9),
                      {"FontSize": 14, "Format": "#,##0.00"}))
    labelled(m, "txtAmount", "المبلغ *", c)
    c = m.add(Control("text", "txtParty", cm(0.4), cm(7.6), width - cm(0.8), cm(0.8), {}))
    labelled(m, "txtParty", "المستلم", c)
    c = m.add(Control("text", "txtDescription", cm(0.4), cm(9.0), width - cm(0.8), cm(0.8), {}))
    labelled(m, "txtDescription", "البيان", c)
    button(m, "btnSave", "حفظ السند", cm(0.4), cm(10.4), "primary", w=cm(3.4), h=cm(1.0),
           call="SaveCashVoucher Me")
    button(m, "btnSavePrint", "حفظ وطباعة", cm(4.0), cm(10.4), "primary", w=cm(3.4), h=cm(1.0),
           call="SaveCashVoucher Me, True")
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(10.4), "secondary", w=cm(2.6),
           h=cm(1.0), call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    VoucherLoad Me", "End Sub",
               "Private Sub cboVoucherType_AfterUpdate()", "    VoucherTypeChanged Me", "End Sub",
               "Private Sub cboCategory_AfterUpdate()", "    VoucherCategoryChanged Me", "End Sub",
               "Private Sub cboBox_AfterUpdate()", "    VoucherBoxChanged Me", "End Sub"]
              + m.code)
    return m


# ------------------------------------------------------------------ cashier closing
def layout_cash_closing() -> FormModel:
    width, height = cm(18.0), cm(16.4)
    m = FormModel("frmCashClosing", "تصفية يومية الكاشير", width, height, popup=True, allow_add=False)
    title_band(m, "تصفية يومية الكاشير", "عدّ النقدية وترحيلها للخزينة الرئيسية أو تسويتها مع المالك",
               "treasury")
    full = width - cm(0.8)
    c = table_combo(m, "cboBox", cm(0.4), cm(2.4), cm(8.4))
    labelled(m, "cboBox", "الصندوق", c)
    m.add(Control("label", "lblPeriod", cm(9.2), cm(2.5), cm(8.4), cm(0.7),
                  {"Caption": " ", "FontSize": 10, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("text", "txtExpected", width - cm(1.0), cm(0.1), cm(0.6), cm(0.4),
                  {"Visible": False, "Format": "0.00"}, decorative=True))
    m.add(Control("list", "lstBreakdown", cm(0.4), cm(3.5), full, cm(3.6),
                  {"ColumnCount": 4, "ColumnWidths": "8.2;1.6;3.6;3.6", "ColumnHeads": True}))
    gap = cm(0.3)
    fw = (full - 3 * gap) // 4
    for i, (key, caption) in enumerate([("Opening", "رصيد البداية"), ("CashIn", "المقبوضات"),
                                        ("CashOut", "المدفوعات"), ("Expected", "الرصيد الدفتري (المفروض)")]):
        figure(m, key, caption, cm(0.4) + i * (fw + gap), cm(7.3), fw, value_font=14)
    c = m.add(Control("text", "txtCounted", cm(0.4), cm(9.7), cm(5.5), cm(1.0),
                      {"FontSize": 16, "Format": "#,##0.00"}, events=["AfterUpdate"]))
    labelled(m, "txtCounted", "النقدية الفعلية بالعدّ *", c)
    m.add(Control("label", "lblDifference", cm(6.2), cm(9.8), width - cm(6.6), cm(0.8),
                  {"Caption": " ", "FontSize": 13, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    c = value_list_combo(m, "cboDestination", cm(0.4), cm(11.5), cm(5.5))
    labelled(m, "cboDestination", "الترحيل إلى", c)
    c = table_combo(m, "cboToBox", cm(6.2), cm(11.5), cm(5.5), events=())
    labelled(m, "cboToBox", "الخزينة المستلمة", c)
    c = m.add(Control("text", "txtTransfer", cm(12.0), cm(11.5), width - cm(12.4), cm(0.8),
                      {"Format": "#,##0.00"}, events=["AfterUpdate"]))
    labelled(m, "txtTransfer", "المبلغ المرحَّل", c)
    m.add(Control("label", "lblKept", cm(0.4), cm(12.45), full, cm(0.6),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    c = m.add(Control("text", "txtNotes", cm(0.4), cm(13.7), full, cm(0.8), {}))
    labelled(m, "txtNotes", "ملاحظات", c)
    button(m, "btnSave", "حفظ التصفية", cm(0.4), cm(14.9), "primary", w=cm(3.6), h=cm(1.0),
           call="SaveCashClosing Me")
    button(m, "btnSavePrint", "حفظ وطباعة", cm(4.2), cm(14.9), "primary", w=cm(3.4), h=cm(1.0),
           call="SaveCashClosing Me, True")
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(14.9), "secondary", w=cm(2.6),
           h=cm(1.0), call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    ClosingLoad Me", "End Sub",
               "Private Sub cboBox_AfterUpdate()", "    ClosingBoxChanged Me", "End Sub",
               "Private Sub txtCounted_AfterUpdate()", "    ClosingCountChanged Me", "End Sub",
               "Private Sub cboDestination_AfterUpdate()", "    ClosingDestinationChanged Me", "End Sub",
               "Private Sub txtTransfer_AfterUpdate()", "    ClosingRecalc Me", "End Sub"]
              + m.code)
    return m


def cash_forms() -> List[FormModel]:
    return [layout_treasury(), layout_cash_voucher(), layout_cash_closing()]
