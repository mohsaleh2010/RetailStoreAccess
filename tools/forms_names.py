"""The English names screen (frmEnglishNames, docs/42-English-Names-Screen.md): the names of customers, suppliers,
products, boxes, banks, accounts ... with their English name, a suggested transliteration for the empty ones,
then one save. The logic is in modEnglishNames; the lines are in the local table tmpEnglishNames."""

from typing import List, Tuple

from forms import Control, FormModel, Sym, button, cm, labelled, title_band
from forms_sales import LOCKED, grid_row, header_labels
from forms_security import check_with_label

NAME_TITLES = ["الجدول", "الاسم بالعربية", "الاسم بالإنجليزية"]


def layout_name_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.7)
    m = FormModel("frmEnglishNameLines", "الأسماء الإنجليزية", cm(21.6), row_h, popup=False,
                  record_source="SELECT * FROM tmpEnglishNames ORDER BY LineNo", allow_add=False)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0, "AllowDeletions": False}
    heads = grid_row(m, [
        ("TableTitle", "TableTitle", 4.4, dict(LOCKED), []),
        ("ArabicName", "ArabicName", 8.3, dict(LOCKED), []),
        ("EnglishName", "EnglishName", 8.3, {"TextAlign": 1}, []),
    ], row_h)
    m.code = []
    return m, heads


def layout_english_names(heads) -> FormModel:
    width, height = cm(22.8), cm(17.6)
    m = FormModel("frmEnglishNames", "الأسماء الإنجليزية", width, height, popup=True, allow_add=False)
    title_band(m, "الأسماء الإنجليزية", "اسم إنجليزي لكل عميل ومورد ومنتج وصندوق وحساب: يظهر في الواجهة الإنجليزية",
               "settings")
    y = cm(2.3)
    c = m.add(Control("combo", "cboTable", cm(0.4), y, cm(6.4), cm(0.8),
                      {"RowSourceType": "Value List", "RowSource": "", "ColumnCount": 2, "ColumnWidths": "0;6",
                       "LimitToList": True}, events=["AfterUpdate"]))
    labelled(m, "cboTable", "الجدول", c)
    check_with_label(m, "chkMissing", "بدون اسم إنجليزي فقط", cm(7.2), y, cm(5.4), events=["AfterUpdate"])
    m.add(Control("label", "lblCount", cm(12.8), y, width - cm(13.2), cm(0.8),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    header_labels(m, cm(0.4), cm(3.45), NAME_TITLES, heads)
    m.add(Control("subform", "subNames", cm(0.4), cm(4.05), cm(21.6), cm(10.6), {"SourceObject": "frmEnglishNameLines"}))
    m.add(Control("label", "lblNote", cm(0.4), cm(14.8), width - cm(0.8), cm(0.9),
                  {"Caption": "زر الاقتراح يكتب الاسم العربي بحروف لاتينية في الخانات الفارغة فقط. راجعه وعدّله "
                              "ثم احفظ. الاسم الفارغ يظهر بالعربية في الواجهة الإنجليزية.",
                   "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    bx, y = cm(0.4), cm(16.0)
    for name, caption, style, w, call in [
            ("btnSave", "حفظ", "primary", 3.4, "SaveEnglishNames Me"),
            ("btnSuggest", "اقتراح للفارغ", "secondary", 3.4, "EnglishNamesSuggest Me"),
            ("btnShow", "تحديث القائمة", "secondary", 3.4, "EnglishNamesShow Me")]:
        button(m, name, caption, bx, y, style, w=cm(w), h=cm(1.0), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), y, "secondary", w=cm(2.6), h=cm(1.0),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    EnglishNamesLoad Me", "End Sub",
               "Private Sub cboTable_AfterUpdate()", "    EnglishNamesShow Me", "End Sub",
               "Private Sub chkMissing_AfterUpdate()", "    EnglishNamesShow Me", "End Sub"] + m.code)
    return m


def names_forms() -> List[FormModel]:
    lines, heads = layout_name_lines()
    return [lines, layout_english_names(heads)]
