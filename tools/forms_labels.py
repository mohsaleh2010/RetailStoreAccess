"""Barcode label screens:
  frmBarcodeLabels + frmLabelLines   choose products and the number of labels, preview, print
  (frmLabelSettings is a DataScreen in forms.py: label size, margins, text lines)
"""

from typing import List, Tuple

from forms import Control, FormModel, Sym, button, cm, fit_window, labelled, title_band
from forms_sales import LOCKED, grid_row, header_labels

LABEL_TITLES = ["الصنف", "الباركود المطبوع", "السعر", "عدد الملصقات", ""]
LABEL_PRODUCT_ROWS = ("SELECT ProductID, ProductName, SellingPrice FROM Products "
                      "WHERE IsActive = True ORDER BY ProductName")


def layout_label_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.75)
    # the record source is set in Form_Open, after the local tables exist (LabelLinesOpen)
    m = FormModel("frmLabelLines", "أسطر الملصقات", cm(19.0), row_h, popup=False, allow_add=False)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0}
    heads = grid_row(m, [
        ("ProductName", "ProductName", 8.6, dict(LOCKED), []),
        ("LabelCode", "LabelCode", 4.4, dict(LOCKED), []),
        ("Price", "Price", 2.4, {**LOCKED, "Format": "#,##0.00"}, []),
        ("Copies", "Copies", 2.4, {"Format": "0", "FontBold": True}, ["AfterUpdate"]),
        ("btnRemove", "", 0.8, {"_kind": "button", "Caption": Sym(f"ChrW(&H{0xE74D:X})"),
                               "Style": "danger", "FontName": Sym("ICON_FONT")}, ["Click"]),
    ], row_h)
    m.add(Control("text", "LineNo", cm(18.9), 0, cm(0.05), row_h, {"Visible": False}, source="LineNo",
                  decorative=True))
    m.form_events = ["Open"]
    m.code = ["Private Sub Form_Open(Cancel As Integer)", "    LabelLinesOpen Me", "End Sub",
              "Private Sub Copies_AfterUpdate()", "    LabelCopiesChanged Me", "End Sub",
              "Private Sub btnRemove_Click()", "    RemoveLabelLine Me", "End Sub"]
    return m, heads


def layout_barcode_labels(line_heads) -> FormModel:
    width, height = cm(20.0), cm(14.6)
    m = FormModel("frmBarcodeLabels", "طباعة ملصقات الباركود", width, height, popup=False, allow_add=False)
    title_band(m, "طباعة ملصقات الباركود",
               "امسح الباركود أو اختر الصنف بالاسم، وحدد عدد الملصقات  |  Enter للإضافة", "products")
    y = cm(2.3)
    c = m.add(Control("text", "txtBarcode", cm(0.4), y, cm(6.0), cm(0.9), {"FontSize": 14},
                      events=["KeyDown"]))
    labelled(m, "txtBarcode", "الباركود أو كود المنتج (Enter)", c)
    c = m.add(Control("text", "txtCopies", cm(6.7), y, cm(2.2), cm(0.9),
                      {"FontSize": 14, "Format": "0", "DefaultValue": "1"}))
    labelled(m, "txtCopies", "عدد الملصقات", c)
    c = m.add(Control("combo", "cboProduct", cm(9.2), y, cm(10.4), cm(0.9),
                      {"RowSource": LABEL_PRODUCT_ROWS, "ColumnCount": 3, "ColumnWidths": "0;8;2",
                       "FontSize": 12}, events=["AfterUpdate"]))
    labelled(m, "cboProduct", "أو اختر المنتج بالاسم", c)

    header_labels(m, cm(0.4), cm(3.55), LABEL_TITLES, line_heads)
    m.add(Control("subform", "subLines", cm(0.4), cm(4.2), cm(19.2), cm(7.6),
                  {"SourceObject": "frmLabelLines"}))
    m.add(Control("label", "lblStatus", cm(0.4), cm(12.0), cm(19.2), cm(0.6),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    bx, by, bh = cm(0.4), cm(13.1), cm(1.1)
    for name, caption, style, w, call in [
            ("btnPreview", "معاينة", "primary", 2.8, "PrintLabels Me, True"),
            ("btnPrint", "طباعة", "primary", 2.8, "PrintLabels Me, False"),
            ("btnClear", "مسح القائمة", "secondary", 3.0, "ClearLabels Me"),
            ("btnSettings", "إعدادات الملصق", "secondary", 3.4, 'OpenScreen "frmLabelSettings"')]:
        button(m, name, caption, bx, by, style, w=cm(w), h=bh, call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), by, "secondary", w=cm(2.6), h=bh,
           call="DoCmd.Close acForm, Me.Name")
    fit_window(m, split_x=cm(16.0), bottom_y=cm(11.9), stretch_w=("subLines", "lblStatus", "cboProduct"),
               stretch_h=("subLines",))
    m.form_events = ["Load"]
    m.code = ([
        "Private Sub Form_Load()", "    LabelsLoad Me", "End Sub",
        "Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)",
        "    LabelBarcodeKeyDown Me, KeyCode", "End Sub",
        "Private Sub cboProduct_AfterUpdate()", "    LabelProductPicked Me", "End Sub",
    ] + m.code)
    return m


def label_forms() -> List[FormModel]:
    lines, heads = layout_label_lines()
    return [lines, layout_barcode_labels(heads)]
