"""Monthly payroll (frmPayroll with the lines subform frmPayrollLines on PayrollLines). The logic is in modPayroll.
The salary data of each employee is the data screen frmEmployeePay (forms.py). Layout as forms.py."""

from typing import List, Tuple

from forms import BANK_ROWS, CASHBOX_ROWS, Control, FormModel, Sym, button, cm, labelled, title_band
from forms_cash import table_combo, value_list_combo
from forms_sales import LOCKED, grid_row, header_labels

EDITABLE = ("Overtime", "Additions", "AbsenceDeduction", "AdvanceDeduction", "OtherDeduction", "Notes")
LINE_TITLES = ["الموظف", "الأساسي", "السكن", "بدلات أخرى", "الإضافي", "مكافآت", "غياب", "سلفة", "جزاءات",
               "التأمينات", "الصافي", "ملاحظات"]
RUN_ROWS = ("SELECT PayrollRunID, Format(PayMonth, 'yyyy/mm') & '  ' & IIf(Status = 'DRAFT', 'مسودة', "
            "IIf(PaidAmount <> 0, 'مرحَّل ومصروف', 'مرحَّل')) FROM PayrollRuns ORDER BY PayMonth DESC")
PAID_FROM = "BANK;من البنك;CASHBOX;من صندوق"


def layout_payroll_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.75)
    m = FormModel("frmPayrollLines", "أسطر مسير الرواتب", cm(26.2), row_h, popup=False,
                  record_source="SELECT * FROM PayrollLines WHERE PayrollRunID = 0 ORDER BY EmployeeName", allow_add=False)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0}
    money = {"Format": "#,##0.00"}
    cols = [("EmployeeName", 4.4, dict(LOCKED)), ("Basic", 2.0, {**LOCKED, **money}),
            ("Housing", 2.0, {**LOCKED, **money}), ("OtherAllow", 2.0, {**LOCKED, **money}),
            ("Overtime", 1.8, dict(money)), ("Additions", 1.8, dict(money)), ("AbsenceDeduction", 1.8, dict(money)),
            ("AdvanceDeduction", 1.8, dict(money)), ("OtherDeduction", 1.8, dict(money)),
            ("GosiEmployee", 1.8, {**LOCKED, **money}), ("NetPay", 2.2, {**LOCKED, **money, "FontBold": True}),
            ("Notes", 2.2, {})]
    heads = grid_row(m, [(f, f, w, props, ["AfterUpdate"] if f in EDITABLE else []) for f, w, props in cols], row_h)
    m.code = []
    for f in EDITABLE:
        m.code += [f"Private Sub {f}_AfterUpdate()", "    PayrollLineChanged Me", "End Sub"]
    return m, heads


def layout_payroll(heads) -> FormModel:
    width, height = cm(27.0), cm(18.4)
    m = FormModel("frmPayroll", "مسير الرواتب", width, height, popup=True, allow_add=False)
    title_band(m, "مسير الرواتب", "مسير كل شهر: الرواتب والبدلات والإضافي والخصومات والتأمينات، ثم الترحيل والصرف",
               "users")
    y = cm(2.3)
    c = table_combo(m, "cboRun", cm(0.4), y, cm(5.0), rows=RUN_ROWS)
    labelled(m, "cboRun", "المسيرات", c)
    c = m.add(Control("text", "txtMonth", cm(5.6), y, cm(3.0), cm(0.8), {"Format": "yyyy/mm"}))
    labelled(m, "txtMonth", "شهر جديد", c)
    button(m, "btnCreate", "إنشاء مسير الشهر", cm(8.8), y, "primary", w=cm(3.6), h=cm(0.8), call="PayrollCreate Me")
    x = cm(12.6)
    for name, caption, style, w, call in [("btnRebuild", "إعادة الإنشاء", "secondary", 2.8, "PayrollRebuild Me"),
                                          ("btnPost", "ترحيل المسير", "primary", 2.8, "PayrollPost Me"),
                                          ("btnUnpost", "إلغاء الترحيل", "secondary", 2.8, "PayrollUnpost Me"),
                                          ("btnDelete", "حذف المسودة", "danger", 2.6, "PayrollDelete Me")]:
        button(m, name, caption, x, y, style, w=cm(w), h=cm(0.8), call=call)
        x += cm(w) + cm(0.15)
    m.add(Control("text", "txtRunID", cm(26.7), cm(1.6), cm(0.2), cm(0.4), {"Visible": False}))
    m.add(Control("label", "lblState", cm(0.4), cm(3.35), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    header_labels(m, cm(0.4), cm(4.1), LINE_TITLES, heads)
    m.add(Control("subform", "subLines", cm(0.4), cm(4.7), cm(26.2), cm(9.4), {"SourceObject": "frmPayrollLines"}))
    m.add(Control("label", "lblTotals", cm(0.4), cm(14.3), width - cm(0.8), cm(0.65),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    y = cm(15.6)
    c = m.add(Control("text", "txtPaidDate", cm(0.4), y, cm(2.8), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtPaidDate", "تاريخ الصرف", c)
    c = value_list_combo(m, "cboPaidFrom", cm(3.4), y, cm(3.0), events=())
    c.props.update({"RowSource": PAID_FROM, "ColumnWidths": "0;2.8"})
    labelled(m, "cboPaidFrom", "الصرف", c)
    c = table_combo(m, "cboBank", cm(6.6), y, cm(4.0), rows=BANK_ROWS, events=())
    labelled(m, "cboBank", "البنك", c)
    c = table_combo(m, "cboBox", cm(10.8), y, cm(3.6), rows=CASHBOX_ROWS, events=())
    labelled(m, "cboBox", "الصندوق", c)
    button(m, "btnPay", "تسجيل الصرف", cm(14.6), y, "primary", w=cm(3.0), h=cm(0.8), call="PayrollPay Me")
    button(m, "btnUndoPay", "إلغاء الصرف", cm(17.8), y, "secondary", w=cm(2.8), h=cm(0.8), call="PayrollUndoPay Me")
    button(m, "btnPrint", "طباعة المسير", cm(20.8), y, "secondary", w=cm(2.8), h=cm(0.8), call="PrintPayroll Me")
    button(m, "btnEmployees", "رواتب الموظفين", cm(0.4), cm(17.15), "secondary", w=cm(3.4), h=cm(0.8),
           call='OpenScreen "frmEmployeePay"')
    # the GOSI rates (Settings)
    x = cm(4.0)
    for name, caption, w, fmt in [("txtGosiEE", "حصة الموظف السعودي", 2.8, "0.00%"),
                                  ("txtGosiER", "حصة المنشأة (سعودي)", 2.8, "0.00%"),
                                  ("txtGosiNonSaudi", "المنشأة (غير سعودي)", 2.8, "0.00%"),
                                  ("txtGosiMax", "الحد الأعلى للأجر", 2.8, "#,##0")]:
        c = m.add(Control("text", name, x, cm(17.15), cm(w), cm(0.8), {"Format": fmt}))
        labelled(m, name, caption, c)
        x += cm(w) + cm(0.2)
    button(m, "btnSaveGosi", "حفظ نسب التأمينات", x, cm(17.15), "secondary", w=cm(3.6), h=cm(0.8),
           call="SaveGosiRates Me")
    button(m, "btnClose", "رجوع", width - cm(0.4) - cm(2.6), cm(17.15), "secondary", w=cm(2.6), h=cm(0.8),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    PayrollLoad Me", "End Sub",
               "Private Sub cboRun_AfterUpdate()", "    PayrollPick Me", "End Sub"] + m.code)
    return m


def payroll_forms() -> List[FormModel]:
    lines, heads = layout_payroll_lines()
    return [lines, layout_payroll(heads)]
