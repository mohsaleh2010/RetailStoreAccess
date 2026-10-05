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

    def test_tile_colours_cover_the_choices(self):
        codes = F.TILE_COLORS.split(";")[0::2]
        values = {self.call("TileColorValue", c) for c in codes}
        self.assertEqual(len(values), len(codes))


class ScreenTests(unittest.TestCase):

    def names(self, form):
        return {c.name for c in FORMS[form].controls}

    def test_touch_screen_has_every_control_the_shop_pos_code_uses(self):
        used = set()
        for header in ("Public Sub ResetSaleHeader(", "Public Sub RecalcPOS(", "Public Function SavePOS(",
                       "Public Sub CustomerChanged(", "Public Sub NewSale(", "Public Sub ReprintLast(",
                       "Private Sub SetPOSStatus(", "Public Sub AddLine("):
            used |= set(re.findall(r"frm!(\w+)", body("modPOS", header)))
        missing = used - self.names("frmTouchPOS")
        self.assertEqual(missing, set())

    def test_touch_screen_has_the_controls_of_modtouchpos(self):
        text = read("modTouchPOS").split("Public Function TestTouchPOS(")[0]
        names = self.names("frmTouchPOS") | self.names("frmTouchPay")
        for name in set(re.findall(r"frm!(\w+)", text)):
            self.assertIn(name, names)
        for prefix, count in (("boxCat", FT.CAT_TILES), ("imgCat", FT.CAT_TILES), ("lblCat", FT.CAT_TILES),
                              ("btnCat", FT.CAT_TILES), ("boxProd", FT.PRODUCT_TILES), ("imgProd", FT.PRODUCT_TILES),
                              ("boxProdStrip", FT.PRODUCT_TILES), ("lblProd", FT.PRODUCT_TILES),
                              ("lblPrice", FT.PRODUCT_TILES), ("btnProd", FT.PRODUCT_TILES)):
            self.assertIn(f'"{prefix}"', text)
            for i in range(1, count + 1):
                self.assertIn(f"{prefix}{i}", names)
        self.assertIn(f"CAT_TILES As Long = {FT.CAT_TILES}", text)
        self.assertIn(f"PRODUCT_TILES As Long = {FT.PRODUCT_TILES}", text)
        for btn in ("btnTypeDineIn", "btnTypeTakeaway", "btnTypeDelivery", "btnPayCash"):
            self.assertIn(btn, names)

    def test_touch_sized_buttons(self):
        for form in ("frmTouchPOS", "frmTouchPay"):
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
        added = {"Settings": ["POSMode", "ImagesFolder"], "Categories": ["ImagePath", "TileColor", "SortOrder"],
                 "Products": ["ImagePath", "TrackStock"],
                 "SalesInvoices": ["OrderType", "TableNo", "DeliveryPhone", "DeliveryAddress"]}
        for t, names in added.items():
            fields = {f.name: f for f in table(t).fields}
            for n in names:
                f = fields[n]
                self.assertTrue(not f.required or f.default is not None, f"{t}.{n}")
        schema = read("modBuildSchema")
        self.assertIn("If FieldExistsIn(tdf, FieldName) Then Exit Sub", schema)
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
        for f in ("OrderType", "TableNo", "DeliveryPhone", "DeliveryAddress"):
            self.assertIn(f"rs!{f} =", sig)
        self.assertIn("TouchOrderProblem(frm)", body("modPOS", "Public Function SavePOS("))
        self.assertIn('"TestTouchPOS"', read("modTestAll"))


class Static_modTouchPOS(VbaModuleChecks, unittest.TestCase):
    module_name = "modTouchPOS"
    vba = read("modTouchPOS")
