"""The budget (frmBudget with the lines subform frmBudgetLines on BudgetLines). The logic is in modBudget.
Layout conventions are the same as forms.py."""

from typing import List, Tuple

from forms import CENTER_ROWS, Control, FormModel, Sym, button, cm, labelled, title_band
from forms_cash import table_combo
from forms_sales import grid_row, header_labels

BUDGET_ACCOUNTS = ("SELECT a.AccountCode, a.AccountCode & '  ' & a.AccountName FROM [@Accounts] AS a WHERE AccountType IN "
                   "('REVENUE', 'EXPENSE') AND IsActive = True ORDER BY TreeKey")
BUDGET_ROWS = "SELECT BudgetID, BudgetYear & '  ' & BudgetName FROM Budgets ORDER BY BudgetYear DESC"
MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
LINE_TITLES = ["الحساب", "المركز"] + MONTHS + ["السنة"]


def layout_budget_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.7)
    m = FormModel("frmBudgetLines", "أسطر الموازنة", cm(26.2), row_h, popup=False,
                  record_source="SELECT * FROM BudgetLines WHERE BudgetID = 0 ORDER BY AccountCode", allow_add=True)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0}
    money = {"Format": "#,##0"}
    total = "=Nz([M1],0)+Nz([M2],0)+Nz([M3],0)+Nz([M4],0)+Nz([M5],0)+Nz([M6],0)+Nz([M7],0)+Nz([M8],0)+Nz([M9],0)" \
            "+Nz([M10],0)+Nz([M11],0)+Nz([M12],0)"
    cols = [("AccountCode", "AccountCode", 3.9, {"_kind": "combo", "RowSource": BUDGET_ACCOUNTS, "ColumnCount": 2,
                                                 "ColumnWidths": "0;5.5", "BoundColumn": 1, "LimitToList": True}),
            ("CostCenterID", "CostCenterID", 2.1, {"_kind": "combo", "RowSource": CENTER_ROWS, "ColumnCount": 2,
                                                   "ColumnWidths": "0;3", "BoundColumn": 1, "LimitToList": True})]
    cols += [(f"M{k}", f"M{k}", 1.45, dict(money)) for k in range(1, 13)]
    cols += [("txtYearTotal", total, 1.7, {"Locked": True, "TabStop": False, "Format": "#,##0", "FontBold": True}),
             ("BudgetID", "BudgetID", 0.05, {"Visible": False})]
    heads = grid_row(m, [(n, f, w, p, []) for n, f, w, p in cols], row_h)
    m.form_events = ["BeforeInsert", "BeforeUpdate", "AfterUpdate"]
    m.code = ["Private Sub Form_BeforeInsert(Cancel As Integer)", "    BudgetLineInsert Me", "End Sub",
              "Private Sub Form_BeforeUpdate(Cancel As Integer)", "    BudgetLineCheck Me, Cancel", "End Sub",
              "Private Sub Form_AfterUpdate()", "    BudgetLineSaved Me", "End Sub"]
    return m, heads


def layout_budget(heads) -> FormModel:
    width, height = cm(27.0), cm(19.2)
    m = FormModel("frmBudget", "الموازنة التقديرية", width, height, popup=True, allow_add=False)
    title_band(m, "الموازنة التقديرية", "مبلغ شهري لكل حساب إيرادات أو مصروفات (ويمكن لكل مركز)، ثم المقارنة بالفعلي",
               "reports")
    y = cm(2.3)
    c = table_combo(m, "cboBudget", cm(0.4), y, cm(4.6), rows=BUDGET_ROWS)
    labelled(m, "cboBudget", "الموازنة", c)
    c = m.add(Control("text", "txtYear", cm(5.2), y, cm(1.8), cm(0.8), {"Format": "0"}))
    labelled(m, "txtYear", "سنة جديدة", c)
    button(m, "btnCreate", "إنشاء موازنة", cm(7.2), y, "primary", w=cm(2.8), h=cm(0.8), call="BudgetCreate Me")
    button(m, "btnAddAccounts", "إضافة كل الحسابات", cm(10.2), y, "secondary", w=cm(3.4), h=cm(0.8),
           call="BudgetAddAccounts Me")
    c = m.add(Control("text", "txtFillYear", cm(13.8), y, cm(1.8), cm(0.8), {"Format": "0"}))
    labelled(m, "txtFillYear", "فعلي سنة", c)
    c = m.add(Control("text", "txtPercent", cm(15.8), y, cm(1.8), cm(0.8), {"Format": "0%"}))
    labelled(m, "txtPercent", "زيادة", c)
    button(m, "btnFill", "ملء من الفعلي", cm(17.8), y, "secondary", w=cm(3.0), h=cm(0.8), call="BudgetFill Me")
    c = m.add(Control("text", "txtAnnual", cm(21.0), y, cm(2.2), cm(0.8), {"Format": "#,##0"}))
    labelled(m, "txtAnnual", "مبلغ سنوي", c)
    button(m, "btnSpread", "توزيع على الأشهر", cm(23.4), y, "secondary", w=cm(3.2), h=cm(0.8), call="BudgetSpread Me")
    m.add(Control("text", "txtBudgetID", cm(26.7), cm(1.6), cm(0.2), cm(0.4), {"Visible": False}))
    header_labels(m, cm(0.4), cm(3.45), LINE_TITLES, heads)
    m.add(Control("subform", "subLines", cm(0.4), cm(4.05), cm(26.2), cm(6.9), {"SourceObject": "frmBudgetLines"}))
    m.add(Control("label", "lblTotals", cm(0.4), cm(11.1), cm(20.0), cm(0.6),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    button(m, "btnDelete", "حذف الموازنة", width - cm(0.4) - cm(3.0), cm(11.0), "danger", w=cm(3.0), h=cm(0.75),
           call="BudgetDelete Me")
    # budget against actual
    y = cm(12.55)
    c = m.add(Control("text", "txtFrom", cm(0.4), y, cm(2.8), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtFrom", "مقارنة من", c)
    c = m.add(Control("text", "txtTo", cm(3.4), y, cm(2.8), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtTo", "إلى", c)
    button(m, "btnCompare", "الموازنة مقابل الفعلي", cm(6.4), y, "primary", w=cm(4.2), h=cm(0.8), call="BudgetCompare Me")
    button(m, "btnPrint", "طباعة المقارنة", cm(10.8), y, "secondary", w=cm(3.0), h=cm(0.8), call="PrintBudgetVariance Me")
    m.add(Control("label", "lblVariance", cm(14.0), y, width - cm(14.4), cm(0.8),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("list", "lstVariance", cm(0.4), cm(13.6), width - cm(0.8), cm(4.2),
                  {"ColumnCount": 9, "ColumnWidths": "0;2;6;3.4;3;3;3;1.8;2.4", "ColumnHeads": True}))
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(18.0), "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    BudgetLoad Me", "End Sub",
               "Private Sub cboBudget_AfterUpdate()", "    BudgetPick Me", "End Sub"] + m.code)
    return m


def budget_forms() -> List[FormModel]:
    lines, heads = layout_budget_lines()
    return [lines, layout_budget(heads)]
