"""Fixed assets (frmAssets) and the monthly depreciation (frmDepreciation). The logic is in modAssets.
Layout conventions are the same as forms.py."""

from typing import List

from forms import BANK_ROWS, CASHBOX_ROWS, CENTER_ROWS, Control, FormModel, Sym, button, cm, labelled, title_band
from forms_bank import COUNTER_ROWS
from forms_cash import table_combo, value_list_combo

ASSET_ACCOUNTS = ("SELECT AccountCode, AccountCode & '  ' & AccountName FROM Accounts WHERE IsPosting = True AND "
                  "IsActive = True AND AccountType = 'ASSET' AND Level2Code = 12 AND AccountCode <> 1790 ORDER BY TreeKey")
SOURCES = "BANK;من البنك;CASHBOX;من صندوق;ACCOUNT;على حساب آخر (مستحقات أو قرض...);OPENING;موجود قبل البرنامج (رصيد افتتاحي)"
DISPOSAL_TO = "BANK;بيع - الثمن في البنك;CASHBOX;بيع - الثمن في صندوق;NONE;استبعاد بدون ثمن (تلف أو فقد)"
ASSET_LIST = ("SELECT AssetID, AssetCode AS [الرقم], AssetName AS [الأصل], AssetGroup AS [المجموعة], "
              "Format(PurchaseDate, 'yyyy/mm/dd') AS [الشراء], Format(q.Cost, '#,##0.00') AS [التكلفة], "
              "Format(AccumDep, '#,##0.00') AS [مجمع الإهلاك], Format(BookValue, '#,##0.00') AS [القيمة الدفترية], "
              "q.StatusName AS [الحالة] FROM FixedAssetsQuery AS q ORDER BY q.Status, q.AssetCode")
RUN_LIST = ("SELECT RunID, RunNumber AS [القيد], Format(RunMonth, 'yyyy/mm') AS [الشهر], "
            "Format(TotalAmount, '#,##0.00') AS [الإهلاك], Format(CreatedAt, 'yyyy/mm/dd') AS [سُجِّل في] "
            "FROM DepreciationRuns ORDER BY RunMonth DESC")


def _text(m, name, caption, x, y, w, fmt=None, events=()):
    c = m.add(Control("text", name, x, y, w, cm(0.8), {"Format": fmt} if fmt else {}, events=list(events)))
    labelled(m, name, caption, c)
    return c


def layout_assets() -> FormModel:
    width, height = cm(27.0), cm(19.4)
    m = FormModel("frmAssets", "الأصول الثابتة", width, height, popup=True, allow_add=False)
    title_band(m, "الأصول الثابتة", "سجل الأصول وقيد شرائها، والإهلاك بالقسط الثابت، والبيع أو الاستبعاد", "journal")
    y = cm(2.3)
    _text(m, "txtAssetName", "اسم الأصل", cm(0.4), y, cm(6.4))
    c = table_combo(m, "cboAssetAccount", cm(7.0), y, cm(5.4), rows=ASSET_ACCOUNTS, events=())
    c.props["ColumnWidths"] = "0;5.2"
    labelled(m, "cboAssetAccount", "حساب الأصل (المجموعة)", c)
    _text(m, "txtPurchaseDate", "تاريخ الشراء", cm(12.6), y, cm(2.8), "yyyy/mm/dd")
    _text(m, "txtCost", "التكلفة بدون الضريبة", cm(15.6), y, cm(3.0), "#,##0.00", ["AfterUpdate"])
    _text(m, "txtInputVAT", "ضريبة المدخلات", cm(18.8), y, cm(2.6), "#,##0.00")
    _text(m, "txtSalvage", "القيمة المتبقية", cm(21.6), y, cm(2.6), "#,##0.00", ["AfterUpdate"])
    _text(m, "txtLife", "العمر (شهر)", cm(24.4), y, cm(2.2), "0", ["AfterUpdate"])
    m.add(Control("text", "txtAssetID", cm(26.7), cm(1.6), cm(0.2), cm(0.4), {"Visible": False}))
    y = cm(3.75)
    _text(m, "txtDepStart", "بداية الإهلاك", cm(0.4), y, cm(2.8), "yyyy/mm/dd")
    c = value_list_combo(m, "cboSource", cm(3.4), y, cm(5.6))
    c.props.update({"RowSource": SOURCES, "ColumnWidths": "0;5.4"})
    labelled(m, "cboSource", "مصدر الشراء", c)
    c = table_combo(m, "cboBank", cm(9.2), y, cm(3.6), rows=BANK_ROWS, events=())
    labelled(m, "cboBank", "البنك", c)
    c = table_combo(m, "cboBox", cm(13.0), y, cm(3.4), rows=CASHBOX_ROWS, events=())
    labelled(m, "cboBox", "الصندوق", c)
    c = table_combo(m, "cboCounter", cm(16.6), y, cm(5.6), rows=COUNTER_ROWS, events=())
    c.props["ColumnWidths"] = "0;5.4"
    labelled(m, "cboCounter", "الحساب الدائن", c)
    _text(m, "txtOpeningAccum", "إهلاك سابق", cm(22.4), y, cm(4.2), "#,##0.00")
    y = cm(5.2)
    _text(m, "txtNotes", "ملاحظات", cm(0.4), y, cm(7.4))
    c = table_combo(m, "cboCenter", cm(8.0), y, cm(3.4), rows=CENTER_ROWS, events=())
    labelled(m, "cboCenter", "مركز التكلفة", c)
    x = cm(11.6)
    for name, caption, style, w, call in [("btnSave", "حفظ الأصل", "primary", 3.0, "AssetSave Me"),
                                          ("btnNew", "أصل جديد", "secondary", 2.6, "AssetNew Me"),
                                          ("btnDelete", "حذف", "danger", 1.8, "AssetDelete Me"),
                                          ("btnDepreciation", "الإهلاك الشهري", "secondary", 3.2,
                                           'OpenScreen "frmDepreciation", 0'),
                                          ("btnPrint", "طباعة السجل", "secondary", 2.8, "PrintAssets Me")]:
        button(m, name, caption, x, y, style, w=cm(w), h=cm(0.8), call=call)
        x += cm(w) + cm(0.2)
    m.add(Control("label", "lblAssetInfo", cm(0.4), cm(6.2), width - cm(0.8), cm(1.0),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("list", "lstAssets", cm(0.4), cm(7.3), width - cm(0.8), cm(8.0),
                  {"RowSourceType": "Table/Query", "RowSource": ASSET_LIST, "ColumnCount": 9,
                   "ColumnWidths": "0;2.2;6;4.4;2.4;2.8;2.8;3;1.8", "ColumnHeads": True}, events=["AfterUpdate"]))
    # selling / scrapping the asset shown
    m.add(Control("label", "lblDisposeCap", cm(0.4), cm(15.5), cm(12.0), cm(0.55),
                  {"Caption": "بيع الأصل المعروض أو استبعاده", "FontSize": 10, "FontBold": True,
                   "ForeColor": Sym("CLR_TEXT")}))
    y = cm(16.75)
    _text(m, "txtDisposalDate", "التاريخ", cm(0.4), y, cm(2.8), "yyyy/mm/dd")
    c = value_list_combo(m, "cboDisposalTo", cm(3.4), y, cm(5.6), events=())
    c.props.update({"RowSource": DISPOSAL_TO, "ColumnWidths": "0;5.4"})
    labelled(m, "cboDisposalTo", "الطريقة", c)
    _text(m, "txtProceeds", "ثمن البيع", cm(9.2), y, cm(2.6), "#,##0.00")
    c = table_combo(m, "cboDisposalBank", cm(12.0), y, cm(3.4), rows=BANK_ROWS, events=())
    labelled(m, "cboDisposalBank", "البنك", c)
    c = table_combo(m, "cboDisposalBox", cm(15.6), y, cm(3.2), rows=CASHBOX_ROWS, events=())
    labelled(m, "cboDisposalBox", "الصندوق", c)
    button(m, "btnDispose", "بيع / استبعاد", cm(19.0), y, "danger", w=cm(3.0), h=cm(0.8), call="AssetDispose Me")
    button(m, "btnUndoDispose", "إلغاء الاستبعاد", cm(22.2), y, "secondary", w=cm(3.2), h=cm(0.8),
           call="AssetUndoDisposal Me")
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(18.0), "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    AssetsLoad Me", "End Sub",
               "Private Sub cboSource_AfterUpdate()", "    AssetSourceChanged Me", "End Sub",
               "Private Sub lstAssets_AfterUpdate()", "    AssetPick Me", "End Sub"]
              + [line for ctl in ("txtCost", "txtSalvage", "txtLife")
                 for line in (f"Private Sub {ctl}_AfterUpdate()", "    AssetRefresh Me", "End Sub")]
              + m.code)
    return m


def layout_depreciation() -> FormModel:
    width, height = cm(20.0), cm(16.4)
    m = FormModel("frmDepreciation", "الإهلاك الشهري", width, height, popup=True, allow_add=False)
    title_band(m, "الإهلاك الشهري", "قيد إهلاك كل شهر بالترتيب: مصروف الإهلاك ومجمع الإهلاك لكل أصل", "journal")
    y = cm(2.3)
    _text(m, "txtThrough", "حتى شهر (أي يوم فيه)", cm(0.4), y, cm(3.4), "yyyy/mm/dd")
    button(m, "btnRun", "تسجيل الإهلاك حتى هذا الشهر", cm(4.0), y, "primary", w=cm(6.0), h=cm(0.8),
           call="DepreciationRun Me")
    button(m, "btnUndo", "حذف آخر شهر", cm(10.2), y, "danger", w=cm(3.0), h=cm(0.8), call="DepreciationUndo Me")
    button(m, "btnAssets", "الأصول", cm(13.4), y, "secondary", w=cm(2.4), h=cm(0.8), call='OpenScreen "frmAssets", 0')
    m.add(Control("label", "lblNext", cm(0.4), cm(3.4), width - cm(0.8), cm(0.7),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("list", "lstPreview", cm(0.4), cm(4.2), width - cm(0.8), cm(5.4),
                  {"RowSourceType": "Value List", "RowSource": "", "ColumnCount": 2, "ColumnWidths": "14;4.4",
                   "ColumnHeads": True}))
    m.add(Control("label", "lblRunsCap", cm(0.4), cm(9.8), cm(12.0), cm(0.55),
                  {"Caption": "قيود الإهلاك المسجلة", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstRuns", cm(0.4), cm(10.4), width - cm(0.8), cm(4.4),
                  {"RowSourceType": "Table/Query", "RowSource": RUN_LIST, "ColumnCount": 5,
                   "ColumnWidths": "0;3.4;3;4;4", "ColumnHeads": True}))
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(15.1), "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = ["Private Sub Form_Load()", "    DepreciationLoad Me", "End Sub"] + m.code
    return m


def asset_forms() -> List[FormModel]:
    return [layout_assets(), layout_depreciation()]
