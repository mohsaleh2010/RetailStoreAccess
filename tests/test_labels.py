"""Barcode labels:
  * the reference encoders (tools/barcode_reference.py) are decoded by a real barcode reader
    (zxing-cpp, when installed): EAN-13 for valid 13-digit codes, Code 128 for the rest;
  * the VBA in modLabels gives exactly the same modules (run through LibreOffice);
  * label settings, screens and report wiring; static checks on modLabels.
"""
import os
import re
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools"))
sys.path.insert(0, HERE)

import barcode_reference as B
import demo_data as D
import forms as F
import reports as RP
import vba_harness as H
from helpers import VbaModuleChecks
from schema import table

SAMPLES = ([p.barcode for p in D.PRODUCTS] +
           ["P00021", "P00001", "ABC-123", "12345", "123456", "0", "42", "Hello World!", "a b~c",
            "4006381333931", "5901234123457", "6281000000015", "62810000000141", "000000000000"])


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


class ReferenceTests(unittest.TestCase):

    def test_known_ean13(self):
        self.assertTrue(B.ean13_check_ok("5901234123457"))
        self.assertFalse(B.ean13_check_ok("5901234123458"))
        m = B.barcode_modules("5901234123457")
        self.assertEqual(len(m), 95)
        self.assertTrue(m.startswith("101") and m.endswith("101") and m[45:50] == "01010")

    def test_code128_shapes(self):
        for text in ("P00021", "12345", "Hello World!"):
            m = B.code128_modules(text)
            self.assertTrue(m.startswith("11010") and m.endswith("1100011101011"), text)
            self.assertEqual((len(m) - 13) % 11, 0, text)
        self.assertEqual(B.barcode_modules("منتج"), "")
        self.assertEqual(B.barcode_modules(""), "")

    def test_a_real_reader_decodes_every_sample(self):
        try:
            import zxingcpp  # noqa: F401
            from PIL import Image  # noqa: F401
        except ImportError:
            self.skipTest("zxing-cpp / Pillow not installed")
        import zxingcpp
        with tempfile.TemporaryDirectory() as tmp:
            for code in SAMPLES:
                with self.subTest(code=code):
                    img = B.render_png(B.barcode_modules(code), os.path.join(tmp, "b.png"))
                    got = zxingcpp.read_barcodes(img)
                    self.assertTrue(got, code)
                    self.assertEqual(got[0].text, code)
                    expected = "EAN13" if B.ean13_check_ok(code) else "Code128"
                    self.assertEqual(got[0].format.name, expected, code)


DRIVER = r'''
Option VBASupport 1

Public Function Ping() As Long
    Ping = 42
End Function
'''


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class LabelRuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modCommon": H.read_module("modCommon"), "modLabels": H.read_module("modLabels"),
                    "Driver": DRIVER})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_vba_modules_match_the_reference(self):
        for code in SAMPLES:
            with self.subTest(code=code):
                self.assertEqual(self.h.call("modLabels", "BarcodePattern", code), B.barcode_modules(code))

    def test_vba_helpers(self):
        self.assertTrue(self.h.call("modLabels", "EanCheckOk", "6281000000014"))
        self.assertFalse(self.h.call("modLabels", "EanCheckOk", "6281000000015"))
        self.assertEqual(self.h.call("modLabels", "BarcodePattern", "منتج"), "")
        self.assertEqual(self.h.call("modLabels", "UnitsToCopies", 2.5), 3)
        self.assertEqual(self.h.call("modLabels", "UnitsToCopies", 0.0), 1)
        self.assertEqual(self.h.call("modLabels", "UnitsToCopies", 99999.0), 500)
        self.assertEqual(self.h.call("modLabels", "LabelCode", "منتج", "P00001"), "P00001")
        self.assertEqual(self.h.call("modLabels", "LabelCode", " 6281000000014 ", "P00001"), "6281000000014")

    def test_in_access_test_vectors_match_the_reference(self):
        body = read("modLabels")
        for code in ("6281000000014", "P00021", "12345"):
            literal = re.search(r'BarcodePattern\("' + code + r'"\) = ((?:"[01]+"(?: & _\s+)?)+)', body)
            self.assertIsNotNone(literal, code)
            joined = "".join(re.findall(r'"([01]+)"', literal.group(1)))
            self.assertEqual(joined, B.barcode_modules(code), code)


class WiringTests(unittest.TestCase):

    def test_settings_table_and_choices(self):
        t = table("LabelSettings")
        codes = set(F.LABEL_LINES.split(";")[0::2])
        self.assertEqual(codes, {"NONE", "STORE", "NAME", "PRICE", "CODE", "BARCODE"})
        for f in t.fields:
            if f.name.endswith(("Line1", "Line2")):
                self.assertIn(f.default.strip('"'), codes, f.name)
                for c in codes:
                    self.assertIn(f'"{c}"', f.rule, f.name)
        self.assertEqual(t.seed_rows, [(1,)])

    def test_settings_screen_covers_every_field(self):
        screen = next(s for s in F.DATA_SCREENS if s.name == "frmLabelSettings")
        shown = {f.field for f in screen.fields if isinstance(f, F.Fld)}
        self.assertEqual(shown, {f.name for f in table("LabelSettings").fields} - {"LabelSettingID"})

    def test_layout_and_report_names_match(self):
        body = read("modLabels")
        rpt = next(m for m in RP.all_reports() if m.name == "rptBarcodeLabels")
        names = {c.name for cs in rpt.controls.values() for c in cs}
        for name in re.findall(r'r\.Controls\("(\w+)"\)', body):
            self.assertIn(name, names)
        for name in re.findall(r'"(txt\w+|boxBar)"', body):
            self.assertIn(name, names)
        self.assertIn("EnsureLabelTables", rpt.prepare)

    def test_buttons_reach_the_label_screen(self):
        forms = {m.name: m for m in F.all_forms()}
        for form, button in [("frmPurchaseInvoice", "btnLabels"), ("frmPurchaseView", "btnLabels"),
                             ("frmInventory", "btnLabels"), ("frmSettings", "btnLabelSettings"),
                             ("frmBarcodeLabels", "btnSettings")]:
            self.assertIn(button, {c.name for c in forms[form].controls}, form)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmBarcodeLabels"], "PRODUCTS")
        self.assertIn('"TestLabels"', read("modTestAll"))


class Static_modLabels(VbaModuleChecks, unittest.TestCase):
    module_name = "modLabels"
    vba = read("modLabels")
