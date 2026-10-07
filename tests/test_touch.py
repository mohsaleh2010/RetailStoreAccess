"""Restaurant touch screen (modTouchPOS, frmTouchPOS/frmTouchLines/frmTouchPay):
helpers run through LibreOffice, the screen has every control the shared shop-POS code uses,
schema upgrade safety, stock tracking, order information on the receipt."""
import os
import re
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools"))
sys.path.insert(0, HERE)

import forms as F
import forms_touch as FT
import queries as Q
import vba_harness as H
from helpers import VbaModuleChecks
from schema import table

DRIVER = r'''
Option VBASupport 1

Public Function Ping() As Long
    Ping = 42
End Function

Public Function RunPad(ByVal keys As String) As String
    Dim s As String, i As Long
    For i = 1 To Len(keys)
        s = PayAppend(s, Mid$(keys, i, 1))
    Next
    RunPad = s
End Function
'''


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def body(module, header):
    text = read(module)
    return text.split(header)[1].split("\nEnd ")[0]


FORMS = {m.name: m for m in F.all_forms()}


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class TouchRuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modCommon": H.read_module("modCommon"), "modTouchPOS": H.read_module("modTouchPOS"),
                    "Driver": DRIVER})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def call(self, *a):
        return self.h.call("modTouchPOS", *a)

    def test_order_text(self):
        self.assertEqual(self.call("OrderTypeText", "DINE_IN", "7"), "طلب داخلي - طاولة 7")
        self.assertEqual(self.call("OrderTypeText", "DINE_IN", ""), "طلب داخلي")
        self.assertEqual(self.call("OrderTypeText", "TAKEAWAY", ""), "طلب سفري")
        self.assertEqual(self.call("OrderTypeText", "DELIVERY", ""), "طلب توصيل")
        self.assertEqual(self.call("OrderTypeText", "", ""), "")

    def test_pages(self):
        for items, per, exp in [(0, 7, 1), (7, 7, 1), (8, 7, 2), (33, 16, 3)]:
            self.assertEqual(self.call("PageCount", items, per), exp)

    def test_number_pad(self):
        for keys, exp in [("12", "12"), ("1.5", "1.5"), ("1.255", "1.25"), ("..5", "0.5"), ("0", "0"),
                          ("05", "5"), ("12<", "1"), ("12C3", "3"), ("1234567890", "123456789")]:
            self.assertEqual(self.h.call("Driver", "RunPad", keys), exp, keys)
        self.assertEqual(self.call("PayAmountOf", "", 57.5), 57.5)
        self.assertEqual(self.call("PayAmountOf", "100", 57.5), 100)

    def test_image_paths(self):
        self.assertEqual(self.call("ImageCandidate", "a.png", "D:\\Img"), "D:\\Img\\a.png")
        self.assertEqual(self.call("ImageCandidate", "a.png", "D:\\Img\\"), "D:\\Img\\a.png")
        self.assertEqual(self.call("ImageCandidate", "C:\\x\\a.png", "D:\\Img"), "C:\\x\\a.png")
        self.assertEqual(self.call("ImageCandidate", "\\\\srv\\a.png", "D:\\Img"), "\\\\srv\\a.png")
        self.assertEqual(self.call("ImageCandidate", "  ", "D:\\Img"), "")

    def test_cafe_helpers(self):
        self.assertEqual(self.call("ToggleInList", ",", "5"), ",5,")
        self.assertEqual(self.call("ToggleInList", ",5,7,", "5"), ",7,")
        self.assertEqual(self.call("ToggleInList", "", "3"), ",3,")
        self.assertTrue(self.call("InList", ",12,3,", "3"))
        self.assertFalse(self.call("InList", ",12,3,", "1"))
        self.assertEqual(self.call("SizeName", "L"), "كبير")
        self.assertEqual(self.call("SizeName", "S"), "صغير")
        self.assertEqual(self.call("CafeLineNote", "كبير", ",3,1,"), "كبير، بدون سكر، ثلج قليل")
        self.assertEqual(self.call("CafeLineNote", "وسط", ","), "وسط")
        self.assertEqual(self.call("OrderTypeText", "TAKEAWAY", "", "سارة"), "طلب سفري  -  الاسم: سارة")
        self.assertEqual(self.call("CafeNote", 5), "")

    def test_tile_colours_cover_the_choices(self):
        codes = F.TILE_COLORS.split(";")[0::2]
        values = {self.call("TileColorValue", c) for c in codes}
        self.assertEqual(len(values), len(codes))


class ScreenTests(unittest.TestCase):

    def names(self, form):
        return {c.name for c in FORMS[form].controls}

    def test_touch_screens_have_every_control_the_shop_pos_code_uses(self):
        used = set()
        for header in ("Public Sub ResetSaleHeader(", "Public Sub RecalcPOS(", "Public Function SavePOS(",
                       "Public Sub CustomerChanged(", "Public Sub NewSale(", "Public Sub ReprintLast(",
                       "Private Sub SetPOSStatus(", "Public Sub AddLine("):
            used |= set(re.findall(r"frm!(\w+)", body("modPOS", header)))
        for form in ("frmTouchPOS", "frmCafePOS"):
            self.assertEqual(used - self.names(form), set(), form)

    def test_screens_have_the_controls_of_modtouchpos(self):
        text = read("modTouchPOS").split("Public Function TestTouchPOS(")[0]
        names = set().union(*(self.names(f) for f in ("frmTouchPOS", "frmTouchPay", "frmCafePOS", "frmCafeItem")))
        for name in set(re.findall(r"frm!(\w+)", text)):
            self.assertIn(name, names)
        for form, cats, prods in (("frmTouchPOS", FT.CAT_TILES, FT.PRODUCT_TILES),
                                  ("frmCafePOS", FT.CAFE_CAT_TILES, FT.CAFE_PRODUCT_TILES)):
            own = self.names(form)
            for prefix, count in (("boxCat", cats), ("imgCat", cats), ("lblCat", cats), ("btnCat", cats),
                                  ("boxProd", prods), ("imgProd", prods), ("boxProdStrip", prods),
                                  ("lblProd", prods), ("lblPrice", prods), ("btnProd", prods)):
                self.assertIn(f'"{prefix}"', text)
                for i in range(1, count + 1):
                    self.assertIn(f"{prefix}{i}", own, form)
                self.assertNotIn(f"{prefix}{count + 1}", own, form)       # TileCount stops at the last one
            for btn in ("btnTypeDineIn", "btnTypeTakeaway", "btnPayCash", "btnCatUp", "btnCatDown",
                        "btnProdPrev", "btnProdNext"):
                self.assertIn(btn, own, form)
        self.assertIn("btnTypeDelivery", self.names("frmTouchPOS"))
        item = self.names("frmCafeItem")
        for i in range(1, FT.ADDON_TILES + 1):
            self.assertIn(f"btnAdd{i}", item)
        self.assertIn(f"ADDON_TILES As Long = {FT.ADDON_TILES}", text)
        for name in ("btnSizeS", "btnSizeM", "btnSizeL", "btnNote1", "btnNote4", "imgItem", "lblItemTotal"):
            self.assertIn(name, item)

    def test_touch_sized_buttons(self):
        for form in ("frmTouchPOS", "frmTouchPay", "frmCafePOS", "frmCafeItem"):
            for c in FORMS[form].controls:
                if c.kind == "button" and not c.name.startswith(("btnClose", "btnReprint")):
                    self.assertGreaterEqual(c.h, F.cm(1.15), f"{form}.{c.name}")
                    self.assertGreaterEqual(c.w, F.cm(0.9), f"{form}.{c.name}")
        for c in FORMS["frmTouchLines"].controls:
            if c.kind == "button":
                self.assertGreaterEqual(c.h, F.cm(1.05), c.name)

    def test_tiles_are_transparent_buttons_over_pictures(self):
        ctl = {c.name: c for c in FORMS["frmTouchPOS"].controls}
        for i in range(1, FT.PRODUCT_TILES + 1):
            b = ctl[f"btnProd{i}"]
            self.assertTrue(b.props["Transparent"])
            for part in ("boxProd", "imgProd", "lblProd", "lblPrice"):
                c = ctl[f"{part}{i}"]
                self.assertTrue(c.decorative)
                self.assertTrue(b.x <= c.x and c.x + c.w <= b.x + b.w and b.y <= c.y and c.y + c.h <= b.y + b.h,
                                part + str(i))

    def test_sales_button_opens_the_chosen_screen(self):
        self.assertIn('If FormName = "frmPOS" Then FormName = SalesScreenName()', read("modStartup"))
        self.assertEqual(F.SCREEN_PERMISSIONS["frmTouchPOS"], "SALES_POS")
        modes = set(F.POS_MODES.split(";")[0::2])
        rule = next(f for f in table("Settings").fields if f.name == "POSMode").rule
        for mode in modes:
            self.assertIn(f'"{mode}"', rule)


class DataTests(unittest.TestCase):

    def test_fields_added_to_existing_tables_can_be_upgraded(self):
        # BuildSchema adds missing fields to an existing back-end: each must be optional or have a default
        added = {"Settings": ["POSMode", "ImagesFolder"],
                 "Categories": ["ImagePath", "TileColor", "SortOrder", "IsAddOn"],
                 "Products": ["ImagePath", "TrackStock", "SizePriceM", "SizePriceL"],
                 "SalesInvoices": ["OrderType", "TableNo", "DeliveryPhone", "DeliveryAddress", "OrderName"],
                 "SalesInvoiceDetails": ["LineNote"]}
        for t, names in added.items():
            fields = {f.name: f for f in table(t).fields}
            for n in names:
                f = fields[n]
                self.assertTrue(not f.required or f.default is not None, f"{t}.{n}")
        schema = read("modBuildSchema")
        self.assertIn("If FieldExistsIn(tdf, FieldName) Then", schema)
        self.assertIn('tdf.Fields(FieldName).ValidationRule = ""', schema)       # a dropped rule goes too
        self.assertIn('m_db.Execute "UPDATE [" & tdf.Name & "] SET [" & FieldName & "] = " & DefaultValue', schema)

    def test_order_types_match(self):
        rule = next(f for f in table("SalesInvoices").fields if f.name == "OrderType").rule
        code = read("modTouchPOS")
        for kind in ("DINE_IN", "TAKEAWAY", "DELIVERY"):
            self.assertIn(f'"{kind}"', rule)
            self.assertIn(f'Case "{kind}"', code)
        self.assertIn("h.OrderType, h.TableNo", Q.query("qrySalesDocPrint").sql)

    def test_made_to_order_products_skip_stock(self):
        self.assertIn("If Not ProductTracksStock(ProductID) Then Exit Sub", body("modSales", "Public Sub ApplyStockMovement("))
        self.assertIn("If Not ProductTracksStock(ProductIDs(i)) Then need = 0",
                      body("modSales", "Public Function CheckStockAvailable("))
        self.assertIn("p.TrackStock = True", Q.query("LowStockQuery").sql)
        self.assertIn('scope = "IsActive = True AND TrackStock = True"', read("modPurchases"))

    def test_order_details_are_posted(self):
        sig = body("modSales", "Public Function PostSaleFromCart(")
        for f in ("OrderType", "TableNo", "DeliveryPhone", "DeliveryAddress", "OrderName", "LineNote"):
            self.assertIn(f"rs!{f} =", sig)
        self.assertIn("TouchOrderProblem(frm)", body("modPOS", "Public Function SavePOS("))
        self.assertIn('"TestTouchPOS"', read("modTestAll"))


class CafeTests(unittest.TestCase):

    def test_sized_drinks_open_the_options_window(self):
        click = body("modTouchPOS", "Public Sub ProductTileClick(")
        self.assertIn("If ProductHasSizes(CLng(id)) Then", click)
        self.assertIn("CafeOpenItem frm, CLng(id)", click)
        self.assertIn('IsCafe = (frm.Name = "frmCafePOS")', read("modTouchPOS"))

    def test_size_prices_need_no_price_permission(self):
        post = body("modSales", "Public Function PostSaleFromCart(")
        self.assertIn("rs!UnitPrice <> Nz(rs!SizePriceM, -1) And rs!UnitPrice <> Nz(rs!SizePriceL, -1)", post)
        self.assertIn("p.SizePriceM, p.SizePriceL", post)

    def test_add_on_categories_stay_out_of_the_category_tiles(self):
        text = read("modTouchPOS")
        self.assertIn("c.IsAddOn = False AND EXISTS", body("modTouchPOS", "Public Sub LoadCategories("))
        self.assertIn("c.IsAddOn = False AND", body("modTouchPOS", "Public Sub TouchLoad("))
        self.assertIn("WHERE c.IsAddOn = True", body("modTouchPOS", "Public Sub ItemLoad("))
        self.assertIn('"   + "', body("modTouchPOS", "Public Sub ItemConfirm("))

    def test_notes_reach_the_cart_and_the_receipt(self):
        self.assertIn('ALTER TABLE tmpPOSLines ADD COLUMN LineNote TEXT(100)', read("modPOS"))
        self.assertIn("AND LineNote Is Null", body("modPOS", "Public Sub AddLine("))
        self.assertIn("d.LineNote", Q.query("qrySalesDocPrint").sql)
        self.assertIn("h.OrderName", Q.query("qrySalesDocPrint").sql)
        import reports as RP
        rcpt = next(m for m in RP.all_reports() if m.name == "rptSalesReceipt")
        srcs = [c.source for cs in rcpt.controls.values() for c in cs]
        self.assertIn(RP.PRODUCT_WITH_NOTE, srcs)
        self.assertTrue(any("OrderTypeText([OrderType],[TableNo],[OrderName])" in s for s in srcs))

    def test_cafe_mode_opens_the_cafe_screen(self):
        name = body("modTouchPOS", "Public Function SalesScreenName(")
        self.assertIn('If FormExists("frmCafePOS") Then SalesScreenName = "frmCafePOS"', name)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmCafePOS"], "SALES_POS")


class Static_modTouchPOS(VbaModuleChecks, unittest.TestCase):
    module_name = "modTouchPOS"
    vba = read("modTouchPOS")
