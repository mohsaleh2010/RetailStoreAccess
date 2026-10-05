"""Restaurant and café touch screens (modTouchPOS):
  frmTouchPOS     restaurant: order type, category column, product grid with pictures, order panel
  frmCafePOS      café: category strip, larger product cards, customer name for the cup
  frmCafeItem     café: size, add-ons, notes and quantity of a drink
  frmTouchLines   cart rows with large + / - buttons (cart table tmpPOSLines, as the shop POS)
  frmTouchPay     cash payment number pad
Every control name the shop POS code uses (modPOS: RecalcPOS, SavePOS, ResetSaleHeader...) exists here,
hidden where the touch screen has no use for it, so the same code prices and posts the order.
"""

from typing import List, Tuple

from forms import (Control, FormModel, Sym, ICONS, button, cm, fit_window, labelled, title_band)
from forms_sales import (CUSTOMER_ROWS, LOCKED, PAYMENT_ROWS, PAYMENT_TYPES, grid_row, header_labels)

CAT_TILES, PRODUCT_TILES = 7, 16          # restaurant grid (modTouchPOS counts the tiles on the screen)
CAFE_CAT_TILES, CAFE_PRODUCT_TILES, ADDON_TILES = 5, 15, 8
TABLES = ";".join(str(n) for n in range(1, 41))
TOUCH_FONT = 14


def layout_touch_lines() -> Tuple[FormModel, list]:
    row_h = cm(1.15)
    m = FormModel("frmTouchLines", "أسطر الطلب", cm(11.0), row_h, popup=False,
                  record_source="SELECT * FROM tmpPOSLines ORDER BY LineNo", allow_add=False)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0}
    big = {"FontSize": 16, "FontBold": True}
    heads = grid_row(m, [
        ("btnRemove", "", 0.9, {"_kind": "button", "Caption": Sym(f"ChrW(&H{0xE74D:X})"),
                               "Style": "danger", "FontName": Sym("ICON_FONT")}, ["Click"]),
        ("LineTotal", "LineTotal", 2.1, {**LOCKED, "Format": "#,##0.00", "FontBold": True, "FontSize": 12}, []),
        ("btnPlus", "", 1.0, {"_kind": "button", "Caption": "+", "Style": "secondary", **big}, ["Click"]),
        ("Quantity", "Quantity", 1.1, {**LOCKED, "Format": "#,##0.###", "TextAlign": 2, **big}, []),
        ("btnMinus", "", 1.0, {"_kind": "button", "Caption": "-", "Style": "secondary", **big}, ["Click"]),
        ("ProductName", "ProductName", 4.4, {**LOCKED, "FontSize": 12, "FontBold": True}, []),
    ], row_h)
    m.code = ["Private Sub btnPlus_Click()", "    TouchQtyStep Me, 1", "End Sub",
              "Private Sub btnMinus_Click()", "    TouchQtyStep Me, -1", "End Sub",
              "Private Sub btnRemove_Click()", "    RemoveCurrentLine Me", "End Sub"]
    return m, heads


def hidden(m: FormModel, kind, name, i, props=None, source="", x0=cm(16.0)):
    """Controls the shop POS code needs but the touch screen does not show."""
    p = {"Visible": False}
    p.update(props or {})
    return m.add(Control(kind, name, x0 + i * cm(0.35), cm(0.1), cm(0.3), cm(0.3), p,
                         source=source, decorative=True))


def layout_touch_pos(line_heads) -> FormModel:
    W, H = cm(33.5), cm(16.4)
    m = FormModel("frmTouchPOS", "نقطة بيع المطعم", W, H, popup=False, allow_add=False)
    title_band(m, "نقطة بيع المطعم", "اختر الفئة ثم الصنف  |  + و - لتعديل الكمية  |  نقدي أو بطاقة للدفع",
               "sales")
    extra = {}
    button(m, "btnClose", "إغلاق", W - cm(2.7), cm(0.3), "secondary", w=cm(2.4), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    button(m, "btnReprint", "إعادة طباعة آخر فاتورة", W - cm(7.6), cm(0.3), "secondary", w=cm(4.7), h=cm(0.9),
           call="ReprintLast Me")
    extra["btnClose"] = extra["btnReprint"] = (1000, 0, 0, 0)

    # controls of the shop POS code (hidden)
    for i, (kind, name, props) in enumerate([
            ("text", "txtBarcode", {}), ("text", "txtQty", {}), ("text", "txtInvoiceDiscount", {}),
            ("text", "txtTendered", {}), ("text", "txtNotes", {}), ("text", "txtOrderType", {}),
            ("combo", "cboPaymentType", {"RowSource": PAYMENT_TYPES, "ColumnCount": 2, "ColumnWidths": "0;5"}),
            ("combo", "cboPaymentMethod", {"RowSource": PAYMENT_ROWS, "ColumnCount": 2, "ColumnWidths": "0;5"}),
            ("label", "lblChange", {"Caption": " "}), ("label", "lblLastInvoice", {"Caption": " "}),
            ("label", "lblCustomerInfo", {"Caption": " "})]):
        hidden(m, kind, name, i, props)

    # ---------------------------------------------------------------- order panel (left)
    x0, pw = cm(0.3), cm(11.0)
    m.add(Control("label", "lblOrderTitle", x0, cm(1.65), pw, cm(0.6),
                  {"Caption": " ", "FontSize": 13, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY"), "TextAlign": 3}))
    bw = (pw - 2 * cm(0.1)) // 3
    for i, (name, caption, kind) in enumerate([("btnTypeDelivery", "توصيل", "DELIVERY"),
                                               ("btnTypeTakeaway", "سفري", "TAKEAWAY"),
                                               ("btnTypeDineIn", "داخلي", "DINE_IN")]):
        button(m, name, caption, x0 + i * (bw + cm(0.1)), cm(2.3), "secondary", w=bw, h=cm(1.3),
               call=f'SetOrderType Me, "{kind}"')
        m.controls[-1].props["FontSize"] = 16
    c = m.add(Control("combo", "cboTable", x0 + cm(6.0), cm(4.25), cm(5.0), cm(1.0),
                      {"RowSource": TABLES, "RowSourceType": "Value List", "ColumnCount": 1, "ColumnWidths": "3",
                       "FontSize": 16, "LimitToList": False}, events=["AfterUpdate"]))
    labelled(m, "cboTable", "رقم الطاولة", c)
    m.add(Control("label", "lblTakeaway", x0, cm(4.25), pw, cm(1.0),
                  {"Caption": "طلب سفري: يُغلَّف ويُسلَّم للعميل", "FontSize": 14, "FontBold": True, "TextAlign": 2,
                   "ForeColor": Sym("CLR_MUTED")}, decorative=True))
    # delivery: shown instead of the table number (same place)
    c = m.add(Control("combo", "cboCustomer", x0 + cm(5.6), cm(4.25), cm(5.4), cm(0.85),
                      {"RowSource": CUSTOMER_ROWS, "ColumnCount": 2, "ColumnWidths": "0;10", "FontSize": 12},
                      events=["AfterUpdate"], decorative=True))
    labelled(m, "cboCustomer", "العميل (اختياري)", c)
    c = m.add(Control("text", "txtDeliveryPhone", x0, cm(4.25), cm(5.4), cm(0.85), {"FontSize": 14},
                      decorative=True))
    labelled(m, "txtDeliveryPhone", "جوال التوصيل *", c)
    c = m.add(Control("text", "txtDeliveryAddress", x0, cm(5.65), pw, cm(0.8), {"FontSize": 12},
                      decorative=True))
    labelled(m, "txtDeliveryAddress", "عنوان التوصيل *", c, label_h=cm(0.4))
    for name in ("lblCustomer", "lblDeliveryPhone", "lblDeliveryAddress"):
        next(x for x in m.controls if x.name == name).decorative = True

    header_labels(m, x0, cm(6.6), ["", "الإجمالي", "", "الكمية", "", "الصنف"], line_heads)
    m.add(Control("subform", "subLines", x0, cm(7.15), pw, cm(4.5), {"SourceObject": "frmTouchLines"}))
    extra["subLines"] = (0, 0, 0, 1000)

    ty = cm(11.8)
    m.add(Control("rect", "boxTotals", x0, ty, pw, cm(2.0), {"BackColor": Sym("CLR_SURFACE")}, decorative=True))
    for i, (name, caption) in enumerate([("SubTotal", "قبل الخصم والضريبة"), ("Discount", "الخصم"),
                                         ("Tax", "ضريبة القيمة المضافة")]):
        y = ty + cm(0.1) + i * cm(0.5)
        m.add(Control("label", f"lblCap{name}", x0 + cm(7.3), y, cm(3.6), cm(0.45),
                      {"Caption": caption, "FontSize": 10, "ForeColor": Sym("CLR_MUTED"), "TextAlign": 3}))
        m.add(Control("label", f"lbl{name}", x0 + cm(4.7), y, cm(2.5), cm(0.45),
                      {"Caption": "0.00", "FontSize": 11, "FontBold": True, "TextAlign": 3, "ForeColor": Sym("CLR_TEXT")}))
    m.add(Control("label", "lblItems", x0 + cm(4.7), ty + cm(1.6), cm(6.2), cm(0.38),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED"), "TextAlign": 3}))
    m.add(Control("label", "lblCapTotal", x0 + cm(0.2), ty + cm(0.05), cm(4.3), cm(0.45),
                  {"Caption": "الإجمالي شامل الضريبة", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("label", "lblTotal", x0 + cm(0.2), ty + cm(0.55), cm(4.3), cm(1.3),
                  {"Caption": "0.00", "FontSize": 26, "FontBold": True, "TextAlign": 2, "ForeColor": Sym("CLR_ACCENT")}))
    m.add(Control("label", "lblStatus", x0, cm(13.9), pw, cm(0.5),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    by, bh = cm(14.5), cm(1.6)
    button(m, "btnPayCash", "نقدي", x0 + cm(6.7), by, "primary", w=cm(4.3), h=bh, call="TouchPayCash Me")
    button(m, "btnPayCard", "مدى / بطاقة", x0 + cm(2.4), by, "nav", w=cm(4.2), h=bh, call="TouchPayCard Me")
    button(m, "btnCancelOrder", "إلغاء", x0, by, "danger", w=cm(2.3), h=bh, call="NewSale Me")
    for name in ("btnPayCash", "btnPayCard", "btnCancelOrder"):
        next(x for x in m.controls if x.name == name).props["FontSize"] = 16
    for c in m.controls:                       # everything under the cart follows the bottom edge
        if c.x < cm(11.4) and c.y >= ty and c.name not in extra:
            extra[c.name] = (0, 0, 1000, 0)

    # ---------------------------------------------------------------- products (middle)
    gx, gw, gy = cm(11.6), cm(16.5), cm(2.3)
    gap_x, gap_y = cm(0.25), cm(0.2)
    tw = (gw - 3 * gap_x) // 4
    th = (cm(14.3) - gy - 3 * gap_y) // 4
    m.add(Control("label", "lblCategoryTitle", gx, cm(1.65), gw, cm(0.6),
                  {"Caption": " ", "FontSize": 14, "FontBold": True, "ForeColor": Sym("CLR_TEXT"), "TextAlign": 3}))
    extra["lblCategoryTitle"] = (0, 1000, 0, 0)
    for i in range(PRODUCT_TILES):
        n, row, col = i + 1, i // 4, 3 - i % 4                 # first product at the top right
        x, y = gx + col * (tw + gap_x), gy + row * (th + gap_y)
        parts = [
            Control("rect", f"boxProd{n}", x, y, tw, th, {"BackColor": Sym("CLR_SURFACE")}, decorative=True),
            Control("rect", f"boxProdStrip{n}", x, y, tw, cm(0.18), {"BackColor": Sym("CLR_PRIMARY")},
                    decorative=True),
            Control("image", f"imgProd{n}", x + cm(0.3), y + cm(0.3), tw - cm(0.6), th - cm(1.35), {},
                    decorative=True),
            Control("label", f"lblProd{n}", x + cm(0.1), y + th - cm(1.0), tw - cm(0.2), cm(0.5),
                    {"Caption": " ", "FontSize": 11, "FontBold": True, "TextAlign": 2, "ForeColor": Sym("CLR_TEXT")},
                    decorative=True),
            Control("label", f"lblPrice{n}", x + cm(0.1), y + th - cm(0.5), tw - cm(0.2), cm(0.45),
                    {"Caption": " ", "FontSize": 11, "FontBold": True, "TextAlign": 2, "ForeColor": Sym("CLR_ACCENT")},
                    decorative=True)]
        for c in parts:
            m.add(c)
        button(m, f"btnProd{n}", " ", x, y, "secondary", w=tw, h=th, call=f"ProductTileClick Me, {n}")
        m.controls[-1].props["Transparent"] = True
        for c in parts + [m.controls[-1]]:
            extra[c.name] = (col * 250, 250, row * 250, 250)
    button(m, "btnProdNext", "التالي", gx, cm(14.5), "secondary", w=cm(3.0), h=cm(1.6), call="ProductPage Me, 1")
    m.add(Control("label", "lblProdPage", gx + cm(3.2), cm(14.95), gw - cm(6.4), cm(0.7),
                  {"Caption": " ", "FontSize": 12, "TextAlign": 2, "ForeColor": Sym("CLR_MUTED")}))
    button(m, "btnProdPrev", "السابق", gx + gw - cm(3.0), cm(14.5), "secondary", w=cm(3.0), h=cm(1.6),
           call="ProductPage Me, -1")
    extra["btnProdNext"] = (0, 0, 1000, 0)
    extra["lblProdPage"] = (0, 1000, 1000, 0)
    extra["btnProdPrev"] = (1000, 0, 1000, 0)

    # ---------------------------------------------------------------- categories (right)
    cx, cw = cm(28.4), cm(4.8)
    m.add(Control("label", "lblCatHeader", cx, cm(1.65), cw, cm(0.6),
                  {"Caption": "الفئات", "FontSize": 14, "FontBold": True, "TextAlign": 3, "ForeColor": Sym("CLR_TEXT")}))
    extra["lblCatHeader"] = (1000, 0, 0, 0)
    ch = cm(1.55)
    for i in range(CAT_TILES):
        n, y = i + 1, gy + i * (ch + cm(0.15))
        parts = [
            Control("rect", f"boxCat{n}", cx, y, cw, ch, {"BackColor": Sym("CLR_PRIMARY")}, decorative=True),
            Control("image", f"imgCat{n}", cx + cw - cm(1.45), y + cm(0.12), cm(1.3), ch - cm(0.24), {},
                    decorative=True),
            Control("label", f"lblCat{n}", cx + cm(0.15), y + cm(0.4), cw - cm(1.75), cm(0.75),
                    {"Caption": " ", "FontSize": 13, "TextAlign": 3, "ForeColor": Sym("CLR_SURFACE")}, decorative=True)]
        for c in parts:
            m.add(c)
        button(m, f"btnCat{n}", " ", cx, y, "secondary", w=cw, h=ch, call=f"CategoryTileClick Me, {n}")
        m.controls[-1].props["Transparent"] = True
        for c in parts + [m.controls[-1]]:
            extra[c.name] = (1000, 0, 0, 0)
    half = (cw - cm(0.2)) // 2
    button(m, "btnCatDown", "التالية", cx, cm(14.5), "secondary", w=half, h=cm(1.6), call="CategoryPage Me, 1")
    button(m, "btnCatUp", "السابقة", cx + half + cm(0.2), cm(14.5), "secondary", w=half, h=cm(1.6),
           call="CategoryPage Me, -1")
    extra["btnCatDown"] = extra["btnCatUp"] = (1000, 0, 1000, 0)

    fit_window(m, extra=extra)
    m.form_events = ["Load", "Unload"]
    m.code = ([
        "Private Sub Form_Load()", "    TouchLoad Me", "End Sub",
        "Private Sub Form_Unload(Cancel As Integer)", "    Cancel = Not POSUnload(Me)", "End Sub",
        "Private Sub cboTable_AfterUpdate()", "    TableChanged Me", "End Sub",
        "Private Sub cboCustomer_AfterUpdate()", "    DeliveryCustomerChanged Me", "End Sub",
    ] + m.code)
    return m


def layout_touch_pay() -> FormModel:
    W = cm(15.6)                               # the title band text is 14 cm wide
    m = FormModel("frmTouchPay", "الدفع نقدًا", W, cm(14.4), popup=True, allow_add=False)
    title_band(m, "الدفع نقدًا", "اكتب المبلغ المستلم أو اختر مبلغًا جاهزًا", "sales")
    x0, iw = cm(0.4), W - cm(0.8)
    hidden(m, "text", "txtPayInput", 0, x0=cm(15.0))
    for y, cap, name, size, color in [(cm(1.8), "الإجمالي", "lblPayTotal", 20, "CLR_TEXT"),
                                      (cm(2.9), "المبلغ المستلم", "lblPayAmount", 24, "CLR_ACCENT")]:
        m.add(Control("label", "lblCap" + name[3:], x0 + cm(7.4), y + cm(0.2), iw - cm(7.4), cm(0.6),
                      {"Caption": cap, "FontSize": 13, "ForeColor": Sym("CLR_MUTED"), "TextAlign": 3}))
        m.add(Control("label", name, x0, y, cm(7.2), cm(1.0),
                      {"Caption": "0.00", "FontSize": size, "FontBold": True, "TextAlign": 1, "ForeColor": Sym(color)}))
    m.add(Control("label", "lblPayChange", x0, cm(4.1), iw, cm(0.6),
                  {"Caption": " ", "FontSize": 13, "FontBold": True, "TextAlign": 2, "ForeColor": Sym("CLR_SUCCESS")}))
    qw = (iw - 4 * cm(0.15)) // 5
    for i, (caption, amount) in enumerate([("500", 500), ("200", 200), ("100", 100), ("50", 50), ("بالضبط", 0)]):
        button(m, f"btnQuick{i + 1}", caption, x0 + i * (qw + cm(0.15)), cm(4.9), "secondary", w=qw, h=cm(1.2),
               call=f"PayQuick Me, {amount}")
        m.controls[-1].props["FontSize"] = 14
    kw, kh = (iw - 2 * cm(0.15)) // 3, cm(1.35)
    keys = [["9", "8", "7"], ["6", "5", "4"], ["3", "2", "1"], ["<", "0", "."]]
    names = {"<": "Back", ".": "Dot"}
    for r, row in enumerate(keys):
        for c, key in enumerate(row):
            caption = "حذف" if key == "<" else key
            button(m, f"btnKey{names.get(key, key)}", caption, x0 + c * (kw + cm(0.15)), cm(6.35) + r * (kh + cm(0.15)),
                   "secondary", w=kw, h=kh, call=f'PayKey Me, "{key}"')
            m.controls[-1].props["FontSize"] = 20
    by, bh = cm(12.45), cm(1.5)
    button(m, "btnPayConfirm", "تأكيد الدفع", x0 + iw - cm(5.9), by, "primary", w=cm(5.9), h=bh, call="PayConfirm Me")
    button(m, "btnPayCancel", "إلغاء", x0 + cm(2.6), by, "secondary", w=cm(2.6), h=bh, call="PayCancel Me")
    button(m, "btnPayClear", "مسح", x0, by, "danger", w=cm(2.4), h=bh, call='PayKey Me, "C"')
    for name in ("btnPayConfirm", "btnPayCancel", "btnPayClear"):
        next(x for x in m.controls if x.name == name).props["FontSize"] = 15
    m.form_events = ["Load"]
    m.code = ["Private Sub Form_Load()", "    PayLoad Me", "End Sub"] + m.code
    return m


ENGINE_CONTROLS = [
    ("text", "txtBarcode", {}), ("text", "txtQty", {}), ("text", "txtInvoiceDiscount", {}),
    ("text", "txtTendered", {}), ("text", "txtNotes", {}), ("text", "txtOrderType", {}),
    ("combo", "cboPaymentType", {"RowSource": PAYMENT_TYPES, "ColumnCount": 2, "ColumnWidths": "0;5"}),
    ("combo", "cboPaymentMethod", {"RowSource": PAYMENT_ROWS, "ColumnCount": 2, "ColumnWidths": "0;5"}),
    ("combo", "cboCustomer", {"RowSource": CUSTOMER_ROWS, "ColumnCount": 2, "ColumnWidths": "0;10"}),
    ("label", "lblChange", {"Caption": " "}), ("label", "lblLastInvoice", {"Caption": " "}),
    ("label", "lblCustomerInfo", {"Caption": " "})]


def totals_and_payment(m: FormModel, x0, pw, ty, extra):
    """Totals box, status line and the three payment buttons (shared by both touch screens)."""
    m.add(Control("rect", "boxTotals", x0, ty, pw, cm(2.0), {"BackColor": Sym("CLR_SURFACE")}, decorative=True))
    for i, (name, caption) in enumerate([("SubTotal", "قبل الخصم والضريبة"), ("Discount", "الخصم"),
                                         ("Tax", "ضريبة القيمة المضافة")]):
        y = ty + cm(0.1) + i * cm(0.5)
        m.add(Control("label", f"lblCap{name}", x0 + cm(7.3), y, cm(3.6), cm(0.45),
                      {"Caption": caption, "FontSize": 10, "ForeColor": Sym("CLR_MUTED"), "TextAlign": 3}))
        m.add(Control("label", f"lbl{name}", x0 + cm(4.7), y, cm(2.5), cm(0.45),
                      {"Caption": "0.00", "FontSize": 11, "FontBold": True, "TextAlign": 3, "ForeColor": Sym("CLR_TEXT")}))
    m.add(Control("label", "lblItems", x0 + cm(4.7), ty + cm(1.6), cm(6.2), cm(0.38),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED"), "TextAlign": 3}))
    m.add(Control("label", "lblCapTotal", x0 + cm(0.2), ty + cm(0.05), cm(4.3), cm(0.45),
                  {"Caption": "الإجمالي شامل الضريبة", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("label", "lblTotal", x0 + cm(0.2), ty + cm(0.55), cm(4.3), cm(1.3),
                  {"Caption": "0.00", "FontSize": 26, "FontBold": True, "TextAlign": 2, "ForeColor": Sym("CLR_ACCENT")}))
    m.add(Control("label", "lblStatus", x0, ty + cm(2.1), pw, cm(0.5),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    by, bh = ty + cm(2.7), cm(1.6)
    button(m, "btnPayCash", "نقدي", x0 + cm(6.7), by, "primary", w=cm(4.3), h=bh, call="TouchPayCash Me")
    button(m, "btnPayCard", "مدى / بطاقة", x0 + cm(2.4), by, "nav", w=cm(4.2), h=bh, call="TouchPayCard Me")
    button(m, "btnCancelOrder", "إلغاء", x0, by, "danger", w=cm(2.3), h=bh, call="NewSale Me")
    for name in ("btnPayCash", "btnPayCard", "btnCancelOrder"):
        next(x for x in m.controls if x.name == name).props["FontSize"] = 16
    for c in m.controls:
        if c.x < x0 + pw and c.y >= ty and c.name not in extra:
            extra[c.name] = (0, 0, 1000, 0)


def product_grid(m: FormModel, gx, gw, gy, gbottom, cols, rows, extra, name_size=11):
    gap_x, gap_y = cm(0.25), cm(0.2)
    tw = (gw - (cols - 1) * gap_x) // cols
    th = (gbottom - gy - (rows - 1) * gap_y) // rows
    for i in range(cols * rows):
        n, row, col = i + 1, i // cols, cols - 1 - i % cols
        x, y = gx + col * (tw + gap_x), gy + row * (th + gap_y)
        parts = [
            Control("rect", f"boxProd{n}", x, y, tw, th, {"BackColor": Sym("CLR_SURFACE")}, decorative=True),
            Control("rect", f"boxProdStrip{n}", x, y, tw, cm(0.18), {"BackColor": Sym("CLR_PRIMARY")},
                    decorative=True),
            Control("image", f"imgProd{n}", x + cm(0.3), y + cm(0.3), tw - cm(0.6), th - cm(1.35), {},
                    decorative=True),
            Control("label", f"lblProd{n}", x + cm(0.1), y + th - cm(1.0), tw - cm(0.2), cm(0.5),
                    {"Caption": " ", "FontSize": name_size, "FontBold": True, "TextAlign": 2,
                     "ForeColor": Sym("CLR_TEXT")}, decorative=True),
            Control("label", f"lblPrice{n}", x + cm(0.1), y + th - cm(0.5), tw - cm(0.2), cm(0.45),
                    {"Caption": " ", "FontSize": 11, "FontBold": True, "TextAlign": 2, "ForeColor": Sym("CLR_ACCENT")},
                    decorative=True)]
        for c in parts:
            m.add(c)
        button(m, f"btnProd{n}", " ", x, y, "secondary", w=tw, h=th, call=f"ProductTileClick Me, {n}")
        m.controls[-1].props["Transparent"] = True
        f = (col * 1000 // cols, 1000 // cols, row * 1000 // rows, 1000 // rows)
        for c in parts + [m.controls[-1]]:
            extra[c.name] = f
    button(m, "btnProdNext", "التالي", gx, cm(14.5), "secondary", w=cm(3.0), h=cm(1.6), call="ProductPage Me, 1")
    m.add(Control("label", "lblProdPage", gx + cm(3.2), cm(14.95), gw - cm(6.4), cm(0.7),
                  {"Caption": " ", "FontSize": 12, "TextAlign": 2, "ForeColor": Sym("CLR_MUTED")}))
    button(m, "btnProdPrev", "السابق", gx + gw - cm(3.0), cm(14.5), "secondary", w=cm(3.0), h=cm(1.6),
           call="ProductPage Me, -1")
    extra["btnProdNext"] = (0, 0, 1000, 0)
    extra["lblProdPage"] = (0, 1000, 1000, 0)
    extra["btnProdPrev"] = (1000, 0, 1000, 0)


def layout_cafe_pos(line_heads) -> FormModel:
    W, H = cm(33.5), cm(16.4)
    m = FormModel("frmCafePOS", "نقطة بيع الكافيه", W, H, popup=False, allow_add=False)
    title_band(m, "نقطة بيع الكافيه", "اختر المشروب: الحجم والإضافات تظهر تلقائيًا  |  نقدي أو بطاقة للدفع", "sales")
    extra = {}
    button(m, "btnClose", "إغلاق", W - cm(2.7), cm(0.3), "secondary", w=cm(2.4), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    button(m, "btnReprint", "إعادة طباعة آخر فاتورة", W - cm(7.6), cm(0.3), "secondary", w=cm(4.7), h=cm(0.9),
           call="ReprintLast Me")
    extra["btnClose"] = extra["btnReprint"] = (1000, 0, 0, 0)
    for i, (kind, name, props) in enumerate(ENGINE_CONTROLS):
        hidden(m, kind, name, i, props)

    # order panel (left)
    x0, pw = cm(0.3), cm(11.0)
    m.add(Control("label", "lblOrderTitle", x0, cm(1.65), pw, cm(0.6),
                  {"Caption": " ", "FontSize": 13, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY"), "TextAlign": 3}))
    bw = (pw - cm(0.1)) // 2
    for i, (name, caption, kind) in enumerate([("btnTypeTakeaway", "سفري", "TAKEAWAY"),
                                               ("btnTypeDineIn", "محلي", "DINE_IN")]):
        button(m, name, caption, x0 + i * (bw + cm(0.1)), cm(2.3), "secondary", w=bw, h=cm(1.3),
               call=f'SetOrderType Me, "{kind}"')
        m.controls[-1].props["FontSize"] = 16
    c = m.add(Control("text", "txtOrderName", x0, cm(4.25), pw, cm(1.0), {"FontSize": 16}, events=["AfterUpdate"]))
    labelled(m, "txtOrderName", "اسم العميل على الكوب (اختياري)", c)
    header_labels(m, x0, cm(5.45), ["", "الإجمالي", "", "الكمية", "", "الصنف"], line_heads)
    m.add(Control("subform", "subLines", x0, cm(6.0), pw, cm(5.65), {"SourceObject": "frmTouchLines"}))
    extra["subLines"] = (0, 0, 0, 1000)
    totals_and_payment(m, x0, pw, cm(11.8), extra)

    # categories strip (top) and product cards
    gx, gw = cm(11.6), cm(21.6)
    arrow = cm(1.4)
    button(m, "btnCatDown", "التالية", gx, cm(1.65), "secondary", w=arrow, h=cm(1.55), call="CategoryPage Me, 1")
    button(m, "btnCatUp", "السابقة", gx + gw - arrow, cm(1.65), "secondary", w=arrow, h=cm(1.55),
           call="CategoryPage Me, -1")
    extra["btnCatUp"] = (1000, 0, 0, 0)
    sx, sw = gx + arrow + cm(0.2), gw - 2 * arrow - cm(0.4)
    cw = (sw - (CAFE_CAT_TILES - 1) * cm(0.2)) // CAFE_CAT_TILES
    for i in range(CAFE_CAT_TILES):
        n, col = i + 1, CAFE_CAT_TILES - 1 - i
        x, y, ch = sx + col * (cw + cm(0.2)), cm(1.65), cm(1.55)
        parts = [
            Control("rect", f"boxCat{n}", x, y, cw, ch, {"BackColor": Sym("CLR_PRIMARY")}, decorative=True),
            Control("image", f"imgCat{n}", x + cw - cm(1.05), y + cm(0.2), cm(0.9), ch - cm(0.4), {},
                    decorative=True),
            Control("label", f"lblCat{n}", x + cm(0.1), y + cm(0.45), cw - cm(1.2), cm(0.7),
                    {"Caption": " ", "FontSize": 12, "TextAlign": 2, "ForeColor": Sym("CLR_SURFACE")}, decorative=True)]
        for c in parts:
            m.add(c)
        button(m, f"btnCat{n}", " ", x, y, "secondary", w=cw, h=ch, call=f"CategoryTileClick Me, {n}")
        m.controls[-1].props["Transparent"] = True
        f = (col * 1000 // CAFE_CAT_TILES, 1000 // CAFE_CAT_TILES, 0, 0)
        for c in parts + [m.controls[-1]]:
            extra[c.name] = f
    m.add(Control("label", "lblCategoryTitle", gx, cm(3.35), gw, cm(0.6),
                  {"Caption": " ", "FontSize": 14, "FontBold": True, "ForeColor": Sym("CLR_TEXT"), "TextAlign": 3}))
    extra["lblCategoryTitle"] = (0, 1000, 0, 0)
    product_grid(m, gx, gw, cm(4.05), cm(14.3), 5, 3, extra, name_size=12)

    fit_window(m, extra=extra)
    m.form_events = ["Load", "Unload"]
    m.code = ([
        "Private Sub Form_Load()", "    TouchLoad Me", "End Sub",
        "Private Sub Form_Unload(Cancel As Integer)", "    Cancel = Not POSUnload(Me)", "End Sub",
        "Private Sub txtOrderName_AfterUpdate()", "    OrderNameChanged Me", "End Sub",
    ] + m.code)
    return m


def layout_cafe_item() -> FormModel:
    W = cm(15.6)
    m = FormModel("frmCafeItem", "خيارات المشروب", W, cm(13.6), popup=True, allow_add=False)
    title_band(m, "المشروب", "اختر الحجم والإضافات، ثم «إضافة للطلب»", "sales")
    for i, name in enumerate(("txtItemProduct", "txtSize", "txtAddOns", "txtNoteFlags", "txtItemQty")):
        hidden(m, "text", name, i, x0=cm(13.2))
    x0, iw = cm(0.4), W - cm(0.8)
    m.add(Control("image", "imgItem", x0 + cm(11.0), cm(1.75), cm(3.8), cm(2.45), {}, decorative=True))
    m.add(Control("label", "lblCapSize", x0, cm(1.75), cm(10.8), cm(0.55),
                  {"Caption": "الحجم", "FontSize": 12, "FontBold": True, "TextAlign": 3, "ForeColor": Sym("CLR_TEXT")}))
    sw = (cm(10.8) - 2 * cm(0.15)) // 3
    for i, code in enumerate(("L", "M", "S")):                 # small on the right
        button(m, f"btnSize{code}", {"S": "صغير", "M": "وسط", "L": "كبير"}[code], x0 + i * (sw + cm(0.15)), cm(2.4),
               "secondary", w=sw, h=cm(1.8), call=f'ItemSize Me, "{code}"')
        m.controls[-1].props["FontSize"] = 14
    m.add(Control("label", "lblCapAddOns", x0, cm(4.4), iw, cm(0.55),
                  {"Caption": "الإضافات", "FontSize": 12, "FontBold": True, "TextAlign": 3, "ForeColor": Sym("CLR_TEXT")}))
    aw = (iw - 3 * cm(0.2)) // 4
    for i in range(ADDON_TILES):
        n, row, col = i + 1, i // 4, 3 - i % 4
        button(m, f"btnAdd{n}", " ", x0 + col * (aw + cm(0.2)), cm(5.0) + row * cm(1.4), "secondary", w=aw, h=cm(1.25),
               call=f"ItemToggleAddOn Me, {n}")
    m.add(Control("label", "lblCapNotes", x0, cm(7.85), iw, cm(0.55),
                  {"Caption": "ملاحظات", "FontSize": 12, "FontBold": True, "TextAlign": 3, "ForeColor": Sym("CLR_TEXT")}))
    for i in range(4):
        n, col = i + 1, 3 - i
        button(m, f"btnNote{n}", " ", x0 + col * (aw + cm(0.2)), cm(8.45), "secondary", w=aw, h=cm(1.2),
               call=f"ItemToggleNote Me, {n}")
    button(m, "btnQtyPlus", "+", x0, cm(9.95), "secondary", w=cm(1.6), h=cm(1.4), call="ItemQty Me, 1")
    m.add(Control("label", "lblItemQty", x0 + cm(1.7), cm(10.2), cm(1.6), cm(0.9),
                  {"Caption": "1", "FontSize": 20, "FontBold": True, "TextAlign": 2, "ForeColor": Sym("CLR_TEXT")}))
    button(m, "btnQtyMinus", "-", x0 + cm(3.4), cm(9.95), "secondary", w=cm(1.6), h=cm(1.4), call="ItemQty Me, -1")
    for name in ("btnQtyPlus", "btnQtyMinus"):
        next(x for x in m.controls if x.name == name).props["FontSize"] = 20
    m.add(Control("label", "lblItemTotal", x0 + cm(5.2), cm(10.2), iw - cm(5.2), cm(0.9),
                  {"Caption": " ", "FontSize": 18, "FontBold": True, "TextAlign": 3, "ForeColor": Sym("CLR_ACCENT")}))
    button(m, "btnItemAdd", "إضافة للطلب", x0 + cm(5.7), cm(11.7), "primary", w=iw - cm(5.7), h=cm(1.6),
           call="ItemConfirm Me")
    button(m, "btnItemCancel", "إلغاء", x0, cm(11.7), "secondary", w=cm(5.5), h=cm(1.6), call="ItemCancel Me")
    for name in ("btnItemAdd", "btnItemCancel"):
        next(x for x in m.controls if x.name == name).props["FontSize"] = 16
    m.form_events = ["Load"]
    m.code = ["Private Sub Form_Load()", "    ItemLoad Me", "End Sub"] + m.code
    return m


def touch_forms() -> List[FormModel]:
    lines, heads = layout_touch_lines()
    return [lines, layout_touch_pos(heads), layout_touch_pay(), layout_cafe_pos(heads), layout_cafe_item()]
