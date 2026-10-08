"""Sales rep commissions (frmCommissions with the lines subform frmCommissionLines on CommissionLines).
The logic is in modSalesReps. The reps and their targets are the data screens frmSalesReps and
frmRepTargets (forms.py). Layout as forms.py."""

from typing import List, Tuple

from forms import COMMISSION_BASES, Control, FormModel, Sym, button, cm, labelled, title_band
from forms_cash import table_combo
from forms_sales import LOCKED, grid_row, header_labels

EDITABLE = ("CommissionBase", "CommissionRate", "Adjustment", "Notes")
LINE_TITLES = ["المندوب", "صافي المبيعات", "التحصيل", "الأساس", "مبلغ الأساس", "النسبة", "تعديل (+/-)", "العمولة",
               "ملاحظات"]
RUN_ROWS = ("SELECT CommissionRunID, Format(RunMonth, 'yyyy/mm') & '  ' & IIf(Status = 'DRAFT', 'مسودة', 'مرحَّل') "
            "FROM CommissionRuns ORDER BY RunMonth DESC")


def layout_commission_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.75)
    m = FormModel("frmCommissionLines", "أسطر مسير العمولات", cm(24.2), row_h, popup=False,
                  record_source="SELECT * FROM CommissionLines WHERE CommissionRunID = 0 ORDER BY RepName",
                  allow_add=False)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0}
    money = {"Format": "#,##0.00"}
    base = {"_kind": "combo", "RowSource": COMMISSION_BASES, "RowSourceType": "Value List", "ColumnCount": 2,
            "ColumnWidths": "0;2.6", "LimitToList": True}
    cols = [("RepName", 4.6, dict(LOCKED)), ("NetSales", 2.6, {**LOCKED, **money}),
            ("Collections", 2.6, {**LOCKED, **money}), ("CommissionBase", 2.8, base),
            ("BaseAmount", 2.6, {**LOCKED, **money}), ("CommissionRate", 1.6, {"Format": "0.00%"}),
            ("Adjustment", 2.2, dict(money)), ("Commission", 2.4, {**LOCKED, **money, "FontBold": True}),
            ("Notes", 2.25, {})]
    heads = grid_row(m, [(f, f, w, props, ["AfterUpdate"] if f in EDITABLE[:3] else []) for f, w, props in cols],
                     row_h)
    m.form_events = ["BeforeUpdate", "AfterUpdate"]          # the audit trail of the edited lines (modAudit)
    m.code = ["Private Sub Form_BeforeUpdate(Cancel As Integer)",
              '    AuditFormBefore Me, "CommissionLines", "CommissionLineID"',
              "End Sub", "Private Sub Form_AfterUpdate()", "    AuditFormAfter Me", "End Sub"]
    for f in EDITABLE[:3]:
        m.code += [f"Private Sub {f}_AfterUpdate()", "    CommissionLineChanged Me", "End Sub"]
    return m, heads


def layout_commissions(heads) -> FormModel:
    width, height = cm(25.0), cm(16.2)
    m = FormModel("frmCommissions", "عمولات المندوبين", width, height, popup=True, allow_add=False)
    title_band(m, "عمولات المندوبين", "مسير كل شهر من مبيعات المندوبين أو تحصيلاتهم، ثم ترحيل قيده؛ "
               "الصرف بسند صرف نقدية", "customers")
    y = cm(2.3)
    c = table_combo(m, "cboRun", cm(0.4), y, cm(4.6), rows=RUN_ROWS)
    labelled(m, "cboRun", "المسيرات", c)
    c = m.add(Control("text", "txtMonth", cm(5.2), y, cm(2.8), cm(0.8), {"Format": "yyyy/mm"}))
    labelled(m, "txtMonth", "شهر جديد", c)
    button(m, "btnCreate", "إنشاء مسير الشهر", cm(8.2), y, "primary", w=cm(3.6), h=cm(0.8), call="CommissionsCreate Me")
    x = cm(12.0)
    for name, caption, style, w, call in [("btnRebuild", "إعادة الإنشاء", "secondary", 2.8, "CommissionsRebuild Me"),
                                          ("btnPost", "ترحيل المسير", "primary", 2.8, "CommissionsPost Me"),
                                          ("btnUnpost", "إلغاء الترحيل", "secondary", 2.8, "CommissionsUnpost Me"),
                                          ("btnDelete", "حذف المسودة", "danger", 2.6, "CommissionsDelete Me")]:
        button(m, name, caption, x, y, style, w=cm(w), h=cm(0.8), call=call)
        x += cm(w) + cm(0.15)
    m.add(Control("text", "txtRunID", cm(24.7), cm(1.6), cm(0.2), cm(0.4), {"Visible": False}))
    m.add(Control("label", "lblState", cm(0.4), cm(3.35), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    header_labels(m, cm(0.4), cm(4.1), LINE_TITLES, heads)
    m.add(Control("subform", "subLines", cm(0.4), cm(4.7), cm(24.2), cm(8.6), {"SourceObject": "frmCommissionLines"}))
    m.add(Control("label", "lblTotals", cm(0.4), cm(13.5), width - cm(0.8), cm(0.65),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    y = cm(14.9)
    button(m, "btnPrint", "طباعة المسير", cm(0.4), y, "secondary", w=cm(2.8), h=cm(0.8), call="PrintCommissionRun Me")
    button(m, "btnReps", "المندوبين", cm(3.4), y, "secondary", w=cm(2.8), h=cm(0.8), call='OpenScreen "frmSalesReps"')
    button(m, "btnPayVoucher", "سند صرف عمولة", cm(6.4), y, "secondary", w=cm(3.2), h=cm(0.8),
           call='OpenScreen "frmCashVoucher", 0, "OUT"')
    button(m, "btnClose", "رجوع", width - cm(0.4) - cm(2.6), y, "secondary", w=cm(2.6), h=cm(0.8),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    CommissionsLoad Me", "End Sub",
               "Private Sub cboRun_AfterUpdate()", "    CommissionsPick Me", "End Sub"] + m.code)
    return m


def sales_rep_forms() -> List[FormModel]:
    lines, heads = layout_commission_lines()
    return [lines, layout_commissions(heads)]
