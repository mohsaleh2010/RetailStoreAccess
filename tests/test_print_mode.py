"""Settings.InvoicePrintMode: print the sales invoice directly, preview it, or don't print after saving."""
import os
import re
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools"))

import forms as F
from schema import table


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


class PrintModeTests(unittest.TestCase):

    def test_setting_is_upgrade_safe_and_matches_the_choices(self):
        f = next(f for f in table("Settings").fields if f.name == "InvoicePrintMode")
        self.assertTrue(f.required and f.default == '"PREVIEW"')
        codes = F.INVOICE_PRINT_MODES.split(";")[0::2]
        self.assertEqual(set(codes), {"DIRECT", "PREVIEW", "NONE"})
        for c in codes:
            self.assertIn(f'"{c}"', f.rule)
            self.assertIn(f'Case "{c}"' if c != "PREVIEW" else "Case Else", read("modPOS"))

    def test_settings_screen_shows_it(self):
        screen = next(s for s in F.DATA_SCREENS if s.name == "frmSettings")
        self.assertIn("InvoicePrintMode", {f.field for f in screen.fields if isinstance(f, F.Fld)})

    def test_saving_honours_the_setting(self):
        body = read("modPOS")
        self.assertIn('If PrintAfter Then PrintAfterSave "SALE", newID', body)
        self.assertIn('If PrintAfter Then PrintAfterSave "RETURN", newID', body)
        self.assertIn('Case "DIRECT": SavePrintView = acViewNormal', body)
        self.assertIn('Case "NONE": SavePrintView = -1', body)
        self.assertIn("If v >= 0 Then PrintSalesDocument DocKind, DocID, False, v", body)
        self.assertIn("DoCmd.OpenReport rpt, PrintView, ,", body)
        # the touch screens print through SavePOS, so they follow the same setting
        self.assertEqual(len(re.findall(r"SavePOS frm, True", read("modTouchPOS"))), 2)


if __name__ == "__main__":
    unittest.main()
