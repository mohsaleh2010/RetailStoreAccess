"""Bank screens: bank transactions (frmBankTx) and the bank reconciliation (frmBankRecon).
The banks themselves are a data screen (frmBanks, forms.py). The logic is in modBank.
Layout conventions are the same as forms.py."""

from typing import List

from forms import BANK_ROWS, CASHBOX_ROWS, Control, FormModel, Sym, button, cm, labelled, title_band
from forms_cash import figure, table_combo, value_list_combo

TX_TYPES = ("DEPOSIT;إيداع نقدية من صندوق;WITHDRAW;سحب نقدية إلى صندوق;SETTLEMENT;تسوية تحصيلات مدى;"
            "TRANSFER;تحويل إلى بنك آخر;OTHER_IN;وارد آخر (فوائد أو قرض);OTHER_OUT;صادر آخر (رسوم بنكية وغيرها)")
COUNTER_ROWS = ("SELECT AccountCode, AccountCode & '  ' & AccountName FROM Accounts WHERE IsPosting = True AND "
                "IsActive = True AND Nz(Level3Code, 0) NOT IN (1100, 1210) AND AccountCode NOT IN (1190, 1200, 1300, "
                "2100) ORDER BY TreeKey")
TX_LIST = ("SELECT t.BankTxID, t.TxNumber AS [الرقم], Format(t.TxDate, 'yyyy/mm/dd') AS [التاريخ], "
           "IIf(t.TxType = 'DEPOSIT', 'إيداع', IIf(t.TxType = 'WITHDRAW', 'سحب', IIf(t.TxType = 'SETTLEMENT', "
           "'تسوية مدى', IIf(t.TxType = 'TRANSFER', 'تحويل', IIf(t.TxType = 'OTHER_IN', 'وارد', 'صادر'))))) AS [النوع], "
           "k.BankName AS [البنك], Format(t.Amount, '#,##0.00') AS [المبلغ], t.Description AS [البيان] "
           "FROM BankTransactions AS t INNER JOIN Banks AS k ON t.BankID = k.BankID "
           "ORDER BY t.TxDate DESC, t.BankTxID DESC")


def layout_bank_tx() -> FormModel:
    width, height = cm(26.0), cm(18.0)
    m = FormModel("frmBankTx", "الحركات البنكية", width, height, popup=True, allow_add=False)
    title_band(m, "الحركات البنكية",
               "إيداع وسحب، تسوية تحصيلات مدى بعمولتها، التحويل بين البنوك، والحركات الأخرى", "treasury")
    y = cm(2.3)
    c = value_list_combo(m, "cboType", cm(0.4), y, cm(5.6))
    c.props.update({"RowSource": TX_TYPES, "ColumnWidths": "0;5.4"})
    labelled(m, "cboType", "نوع الحركة", c)
    c = m.add(Control("text", "txtDate", cm(6.2), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtDate", "التاريخ", c)
    c = table_combo(m, "cboBank", cm(9.4), y, cm(5.2), rows=BANK_ROWS)
    labelled(m, "cboBank", "البنك", c)
    c = table_combo(m, "cboToBank", cm(14.8), y, cm(5.2), rows=BANK_ROWS)
    labelled(m, "cboToBank", "إلى بنك (التحويل)", c)
    c = table_combo(m, "cboBox", cm(20.2), y, cm(5.4), rows=CASHBOX_ROWS)
    labelled(m, "cboBox", "الصندوق (إيداع / سحب)", c)
    y = cm(3.85)
    c = m.add(Control("text", "txtAmount", cm(0.4), y, cm(3.2), cm(0.8), {"Format": "#,##0.00"}, events=["AfterUpdate"]))
    labelled(m, "txtAmount", "المبلغ", c)
    c = m.add(Control("text", "txtFee", cm(3.8), y, cm(2.8), cm(0.8), {"Format": "#,##0.00"}, events=["AfterUpdate"]))
    labelled(m, "txtFee", "عمولة مدى", c)
    c = m.add(Control("text", "txtFeeVAT", cm(6.8), y, cm(2.8), cm(0.8), {"Format": "#,##0.00"}, events=["AfterUpdate"]))
    labelled(m, "txtFeeVAT", "ضريبة العمولة / الرسوم", c)
    c = table_combo(m, "cboCounter", cm(9.8), y, cm(7.0), rows=COUNTER_ROWS, events=())
    c.props["ColumnWidths"] = "0;6.8"
    labelled(m, "cboCounter", "الحساب المقابل (الحركات الأخرى)", c)
    c = m.add(Control("text", "txtReference", cm(17.0), y, cm(3.2), cm(0.8), {}))
    labelled(m, "txtReference", "مرجع البنك", c)
    c = m.add(Control("text", "txtDescription", cm(20.4), y, cm(5.2), cm(0.8), {}))
    labelled(m, "txtDescription", "البيان", c)
    m.add(Control("label", "lblInfo", cm(0.4), cm(4.9), cm(19.6), cm(1.0),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    button(m, "btnSave", "حفظ الحركة", cm(20.4), cm(5.0), "primary", w=cm(5.2), h=cm(0.9), call="SaveBankTx Me")
    m.add(Control("label", "lblListCap", cm(0.4), cm(6.2), cm(12.0), cm(0.55),
                  {"Caption": "الحركات المسجلة", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstTx", cm(0.4), cm(6.8), width - cm(0.8), cm(9.6),
                  {"RowSourceType": "Table/Query", "RowSource": TX_LIST, "ColumnCount": 7,
                   "ColumnWidths": "0;2.6;2.4;2.4;4;3;10", "ColumnHeads": True}))
    button(m, "btnDelete", "حذف الحركة", cm(0.4), cm(16.7), "danger", w=cm(3.0), h=cm(0.9),
           call="DeleteSelectedBankTx Me")
    button(m, "btnRecon", "التسوية البنكية", cm(3.6), cm(16.7), "secondary", w=cm(3.4), h=cm(0.9),
           call='OpenScreen "frmBankRecon", 0, Me!cboBank.Value')
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(16.7), "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    BankTxLoad Me", "End Sub",
               "Private Sub cboType_AfterUpdate()", "    BankTxTypeChanged Me", "End Sub"]
              + [line for ctl in ("cboBank", "cboToBank", "cboBox", "txtAmount", "txtFee", "txtFeeVAT")
                 for line in (f"Private Sub {ctl}_AfterUpdate()", "    BankTxRefresh Me", "End Sub")]
              + m.code)
    return m


def layout_bank_recon() -> FormModel:
    width, height = cm(27.0), cm(19.2)
    m = FormModel("frmBankRecon", "التسوية البنكية", width, height, popup=True, allow_add=False)
    title_band(m, "التسوية البنكية", "مطابقة كشف البنك مع الدفاتر: علِّم العمليات الظاهرة في الكشف", "treasury")
    y = cm(2.3)
    c = table_combo(m, "cboBank", cm(0.4), y, cm(5.4), rows=BANK_ROWS)
    labelled(m, "cboBank", "البنك", c)
    c = m.add(Control("text", "txtStatementDate", cm(6.0), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtStatementDate", "تاريخ كشف البنك", c)
    c = m.add(Control("text", "txtStatementBalance", cm(9.2), y, cm(3.4), cm(0.8), {"Format": "#,##0.00"}))
    labelled(m, "txtStatementBalance", "رصيد الكشف في هذا التاريخ", c)
    button(m, "btnStart", "بدء / تحديث التسوية", cm(12.8), y, "primary", w=cm(4.2), h=cm(0.8), call="BankReconStart Me")
    m.add(Control("text", "txtReconID", cm(17.2), y, cm(0.6), cm(0.8), {"Visible": False}))
    part = (width - cm(0.8)) // 4
    for i, (key, caption) in enumerate([("Book", "الرصيد في الدفاتر"), ("Outstanding", "عمليات لم تظهر في الكشف"),
                                        ("Adjusted", "الدفاتر بعد استبعادها"), ("Difference", "الفرق مع الكشف")]):
        figure(m, key, caption, cm(0.4) + i * part, cm(3.35), part - cm(0.2), value_font=14)
    m.add(Control("label", "lblState", cm(0.4), cm(5.15), width - cm(0.8), cm(0.95),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    half = (width - cm(1.0)) // 2
    m.add(Control("label", "lblOpenCap", cm(0.4), cm(6.2), half, cm(0.55),
                  {"Caption": "عمليات الدفاتر غير المطابقة (حتى تاريخ الكشف)", "FontSize": 10, "FontBold": True,
                   "ForeColor": Sym("CLR_TEXT")}))
    m.add(Control("list", "lstOpen", cm(0.4), cm(6.8), half, cm(7.4),
                  {"ColumnCount": 7, "ColumnWidths": "0;2;2.2;2.2;2.8;1.8;1.8", "ColumnHeads": True, "MultiSelect": 2},
                  events=["DblClick"]))
    m.add(Control("label", "lblClearedCap", cm(0.6) + half, cm(6.2), half, cm(0.55),
                  {"Caption": "المطابقة في هذه التسوية", "FontSize": 10, "FontBold": True,
                   "ForeColor": Sym("CLR_TEXT")}))
    m.add(Control("list", "lstCleared", cm(0.6) + half, cm(6.8), half, cm(7.4),
                  {"ColumnCount": 6, "ColumnWidths": "0;2.2;2.6;2.6;2.4;2.4", "ColumnHeads": True, "MultiSelect": 2},
                  events=["DblClick"]))
    y = cm(14.4)
    button(m, "btnClear", "مطابقة المحدد", cm(0.4), y, "primary", w=cm(3.4), h=cm(0.9), call="BankReconClear Me")
    button(m, "btnAddTx", "حركة بنكية جديدة", cm(4.0), y, "secondary", w=cm(3.8), h=cm(0.9), call="BankReconAddTx Me")
    button(m, "btnRefresh", "تحديث", cm(8.0), y, "secondary", w=cm(2.2), h=cm(0.9), call="BankReconRefresh Me")
    button(m, "btnUnclear", "إلغاء المطابقة", cm(0.6) + half, y, "secondary", w=cm(3.4), h=cm(0.9),
           call="BankReconUnclear Me")
    m.add(Control("label", "lblReconsCap", cm(0.4), cm(15.5), cm(12.0), cm(0.55),
                  {"Caption": "تسويات هذا البنك", "FontSize": 9, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstRecons", cm(0.4), cm(16.1), cm(13.0), cm(2.7),
                  {"RowSourceType": "Table/Query", "RowSource": "", "ColumnCount": 5,
                   "ColumnWidths": "0;2.8;2.8;3.4;2.4", "ColumnHeads": True}, events=["AfterUpdate"]))
    x = cm(13.6)
    for name, caption, style, w, call in [("btnFinish", "اعتماد التسوية", "primary", 3.4, "BankReconFinish Me"),
                                          ("btnReopen", "إعادة فتح", "secondary", 2.6, "BankReconReopen Me"),
                                          ("btnDeleteRecon", "حذف الجارية", "danger", 2.8, "BankReconDelete Me")]:
        button(m, name, caption, x, cm(16.1), style, w=cm(w), h=cm(0.9), call=call)
        x += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(17.9), "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    BankReconLoad Me", "End Sub",
               "Private Sub cboBank_AfterUpdate()", "    BankReconBankChanged Me", "End Sub",
               "Private Sub lstRecons_AfterUpdate()", "    BankReconPick Me", "End Sub",
               "Private Sub lstOpen_DblClick(Cancel As Integer)", "    BankReconClear Me", "End Sub",
               "Private Sub lstCleared_DblClick(Cancel As Integer)", "    BankReconUnclear Me", "End Sub"] + m.code)
    return m


# ------------------------------------------------------------------ cheques
DIRECTIONS = "IN;شيكات واردة (من العملاء);OUT;شيكات صادرة (للموردين)"
SHOW = "PENDING;تحت التحصيل;DUE;مستحقة خلال 7 أيام أو فات استحقاقها;COLLECTED;المحصَّلة / المصروفة;BOUNCED;المرتدة;ALL;الكل"


def layout_cheques() -> FormModel:
    width, height = cm(27.0), cm(18.6)
    m = FormModel("frmCheques", "الشيكات", width, height, popup=True, allow_add=False)
    title_band(m, "الشيكات الواردة والصادرة",
               "تسجيل الشيك يسدد رصيد العميل أو المورد، ثم يُحصَّل في البنك أو يرتد", "treasury")
    y = cm(2.3)
    c = value_list_combo(m, "cboDirection", cm(0.4), y, cm(6.0))
    c.props.update({"RowSource": DIRECTIONS, "ColumnWidths": "0;5.8"})
    labelled(m, "cboDirection", "النوع", c)
    c = value_list_combo(m, "cboShow", cm(6.6), y, cm(6.4))
    c.props.update({"RowSource": SHOW, "ColumnWidths": "0;6.2"})
    labelled(m, "cboShow", "عرض", c)
    m.add(Control("label", "lblTotals", cm(13.2), y, width - cm(13.6), cm(0.8),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    # a new cheque
    m.add(Control("label", "lblNewCap", cm(0.4), cm(3.4), cm(12.0), cm(0.55),
                  {"Caption": "تسجيل شيك جديد", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_TEXT")}))
    y = cm(4.55)
    c = table_combo(m, "cboParty", cm(0.4), y, cm(6.0), events=())
    labelled(m, "cboParty", "العميل", c)
    c = m.add(Control("text", "txtChequeNo", cm(6.6), y, cm(3.0), cm(0.8), {}))
    labelled(m, "txtChequeNo", "رقم الشيك", c)
    c = m.add(Control("text", "txtDrawerBank", cm(9.8), y, cm(4.0), cm(0.8), {}))
    labelled(m, "txtDrawerBank", "بنك الساحب (الوارد)", c)
    c = table_combo(m, "cboBank", cm(14.0), y, cm(4.4), rows=BANK_ROWS, events=())
    labelled(m, "cboBank", "بنكنا (الصادر: المسحوب عليه)", c)
    c = m.add(Control("text", "txtIssueDate", cm(18.6), y, cm(2.6), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtIssueDate", "تاريخ الشيك", c)
    c = m.add(Control("text", "txtDueDate", cm(21.4), y, cm(2.6), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtDueDate", "الاستحقاق", c)
    c = m.add(Control("text", "txtAmount", cm(24.2), y, cm(2.4), cm(0.8), {"Format": "#,##0.00"}))
    labelled(m, "txtAmount", "المبلغ", c)
    y = cm(6.0)
    c = m.add(Control("text", "txtNotes", cm(0.4), y, cm(13.4), cm(0.8), {}))
    labelled(m, "txtNotes", "ملاحظات", c)
    button(m, "btnSave", "تسجيل الشيك", cm(14.0), y, "primary", w=cm(4.4), h=cm(0.8), call="SaveCheque Me")
    m.add(Control("list", "lstCheques", cm(0.4), cm(7.2), width - cm(0.8), cm(8.2),
                  {"ColumnCount": 9, "ColumnWidths": "0;2.4;2.8;5.4;2.4;2.6;2.4;2.4;4.6", "ColumnHeads": True}))
    # the selected cheque
    y = cm(16.3)
    c = m.add(Control("text", "txtActionDate", cm(0.4), y, cm(2.8), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtActionDate", "تاريخ العملية", c)
    c = table_combo(m, "cboActionBank", cm(3.4), y, cm(4.4), rows=BANK_ROWS, events=())
    labelled(m, "cboActionBank", "البنك (التحصيل / الصرف)", c)
    x = cm(8.0)
    for name, caption, style, w, call in [("btnCollect", "تحصيل في البنك", "primary", 3.4, "CollectCheque Me"),
                                          ("btnBounce", "ارتداد", "danger", 2.2, "BounceCheque Me"),
                                          ("btnUndo", "إلغاء الحالة", "secondary", 2.8, "UndoCheque Me"),
                                          ("btnDelete", "حذف", "danger", 1.8, "DeleteSelectedCheque Me")]:
        button(m, name, caption, x, y, style, w=cm(w), h=cm(0.8), call=call)
        x += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(17.4), "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    ChequesLoad Me", "End Sub",
               "Private Sub cboDirection_AfterUpdate()", "    ChequesDirectionChanged Me", "End Sub",
               "Private Sub cboShow_AfterUpdate()", "    ChequesRefresh Me", "End Sub"] + m.code)
    return m


def bank_forms() -> List[FormModel]:
    return [layout_bank_tx(), layout_bank_recon(), layout_cheques()]
