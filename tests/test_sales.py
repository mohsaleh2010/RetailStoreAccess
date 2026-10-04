"""Phase 6 checks that do not need LibreOffice: QR reference against the
`qrcode` library and OpenCV (skipped when not installed), ZATCA payload,
pricing invariants, report layout and bindings, and the generated modules.

Run:  python3 -m unittest discover -s tests -v
"""

import random
import re
import string
import unittest
from decimal import Decimal as D

from access_sqlite import AccessOnSqlite
from helpers import VbaModuleChecks
import forms as F
import gen_qr
import gen_reports
import gen_test_sales
import pricing as P
import qr_reference as R
import reports as RP
import zatca_reference as Z

try:
    import qrcode
    HAVE_QRCODE = True
except ImportError:
    HAVE_QRCODE = False
try:
    import cv2
    import numpy as np
    HAVE_CV2 = True
except ImportError:
    HAVE_CV2 = False


class ZatcaReferenceTests(unittest.TestCase):

    def test_official_example(self):
        self.assertEqual(Z.qr_payload("Bobs Records", "310122393500003", "2022-04-25T15:30:00Z", 1000, 150),
                         "AQxCb2JzIFJlY29yZHMCDzMxMDEyMjM5MzUwMDAwMwMUMjAyMi0wNC0yNVQxNTozMDowMFoEBzEwMDAuMDAF"
                         "BjE1MC4wMA==")

    def test_round_trip_arabic(self):
        payload = Z.qr_payload("متجر الاختبار", "300000000000003", "2026-10-04T09:15:00Z", 64.2, 8.37)
        self.assertEqual(Z.decode_payload(payload),
                         {1: "متجر الاختبار", 2: "300000000000003", 3: "2026-10-04T09:15:00Z",
                          4: "64.20", 5: "8.37"})


class QrReferenceTests(unittest.TestCase):

    @unittest.skipUnless(HAVE_QRCODE, "qrcode library not installed")
    def test_identical_to_qrcode_library(self):
        random.seed(3)
        for length in (1, 25, 90, 140, 230, 400, 660):
            data = "".join(random.choice(string.ascii_letters + string.digits + "+/=")
                           for _ in range(length)).encode()
            for mask in range(8):
                mine, version, _ = R.encode(data, mask=mask)
                q = qrcode.QRCode(version=version, error_correction=qrcode.constants.ERROR_CORRECT_M,
                                  mask_pattern=mask, border=0)
                q.add_data(qrcode.util.QRData(data, mode=qrcode.util.MODE_8BIT_BYTE))
                q.make(fit=False)
                self.assertEqual(mine, [[1 if x else 0 for x in row] for row in q.modules],
                                 f"length {length} mask {mask}")

    @unittest.skipUnless(HAVE_CV2, "OpenCV not installed")
    def test_decodable_by_opencv(self):
        detector = cv2.QRCodeDetector()
        for text in (Z.qr_payload("متجر الاختبار", "300000000000003", "2026-10-04T09:15:00Z", 1150, 150),
                     "Q" * 400):
            mat, _, _ = R.encode(text.encode())
            img = np.kron(np.pad(1 - np.array(mat, dtype=np.uint8), 4, constant_values=1) * 255,
                          np.ones((8, 8), dtype=np.uint8))
            decoded, _, _ = detector.detectAndDecode(img)
            self.assertEqual(decoded, text)

    def test_capacity_limit(self):
        with self.assertRaises(ValueError):
            R.choose_version(700)
        self.assertEqual(R.choose_version(666), 20)


class PricingTests(unittest.TestCase):

    def test_invariants(self):
        for case in P.CASES:
            o = P.calc(*P.case_inputs(case)[:2], case[1])
            self.assertEqual(o.subtotal - o.discount, o.taxable, case[0])
            self.assertEqual(o.taxable + o.tax, o.total, case[0])
            for ln in o.lines:
                self.assertGreaterEqual(ln.discount, 0, case[0])
                self.assertEqual(ln.net + ln.tax, ln.total, case[0])

    def test_vat_inclusive_customer_pays_shelf_price(self):
        for qty, price in (("1", "11.50"), ("3", "9.99"), ("7", "0.05"), ("2", "23.00")):
            o = P.calc([P.LineIn(D(qty), D(price))], D(0), True)
            self.assertEqual(o.total, P.r2(D(qty) * D(price)))

    def test_invoice_discount_is_fully_spread(self):
        lines = [P.LineIn(D("2"), D("25.00")), P.LineIn(D("1"), D("7.95")), P.LineIn(D("5"), D("3.45"))]
        o = P.calc(lines, D("10"), True)
        self.assertEqual(sum(P.r2(l.qty * l.price) for l in lines) - D("10"), o.total)

    def test_returns_never_exceed_the_sale(self):
        for _, (q, t, x), rets in P.RETURN_CASES:
            pq = pt = px = D(0)
            for r in rets:
                _, tax, total = P.return_amounts(D(q), D(t), D(x), pq, pt, px, D(r))
                pq, pt, px = pq + D(r), pt + total, px + tax
            if pq == D(q):
                self.assertEqual((pt, px), (D(t), D(x)))

    def test_cash_sale_must_be_paid(self):
        with self.assertRaises(P.PricingError):
            P.settle(D(87), D(50), credit=False)
        self.assertEqual(P.settle(D(87), D(100), credit=False), (D(87), D(0), D(13)))


def overlaps(a, b):
    return a.x < b.x + b.w and b.x < a.x + a.w and a.y < b.y + b.h and b.y < a.y + a.h


class ReportLayoutTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        db = AccessOnSqlite()
        db.load_fixture()
        cur = db.con.execute('SELECT * FROM "qrySalesDocPrint"')
        cls.columns = {d[0] for d in cur.description}

    def test_controls_inside_sections_and_not_overlapping(self):
        for m in RP.sales_reports():
            for sec, ctls in m.controls.items():
                for c in ctls:
                    with self.subTest(report=m.name, control=c.name):
                        self.assertGreaterEqual(c.x, 0)
                        self.assertLessEqual(c.x + c.w, m.width)
                        self.assertLessEqual(c.y + c.h, m.heights[sec])
                solid = [c for c in ctls if not c.decorative and c.props.get("Visible", True)]
                for i, a in enumerate(solid):
                    for b in solid[i + 1:]:
                        self.assertFalse(overlaps(a, b), f"{m.name}: {a.name} / {b.name}")

    def test_fields_exist_in_print_query(self):
        for m in RP.sales_reports():
            for ctls in m.controls.values():
                for c in ctls:
                    if c.kind != "text":
                        continue
                    refs = re.findall(r"\[(\w+)\]", c.source) if c.source.startswith("=") else [c.source]
                    for ref in refs:
                        self.assertIn(ref, self.columns, f"{m.name}.{c.name}")

    def test_qr_box_is_centred_and_large_enough(self):
        for m in RP.sales_reports():
            box = next(c for c in m.controls[RP.SEC_FOOTER] if c.name == "boxQR")
            self.assertEqual(box.x * 2 + box.w, m.width - (m.width - box.w) % 2)
            self.assertGreaterEqual(box.w, F.cm(3.0), "ZATCA QR should be readable (>= 3 cm)")
            self.assertTrue(any("DrawDocumentQR" in line for line in m.code))

    def test_receipt_fits_80mm_paper(self):
        self.assertLessEqual(RP.receipt().width, F.cm(7.6))


class GeneratedQRModuleTests(VbaModuleChecks, unittest.TestCase):
    module_name = "modQRCode"
    vba = gen_qr.build_qr_vba()

    def test_tables_copied_from_reference(self):
        for v, (ec, b1, d1, b2, d2) in R.M_BLOCKS.items():
            self.assertIn(f"Case {v}: t = Array({ec}, {b1}, {d1}, {b2}, {d2})", self.vba)

    def test_no_array_assignment_between_arrays(self):
        """Copy-vs-reference semantics differ between engines: copy explicitly."""
        self.assertIsNone(re.search(r"^\s+\w+ = (?:m_mod|candidate|baseM|nxt)\s*$", self.vba, re.M))


class GeneratedReportsModuleTests(VbaModuleChecks, unittest.TestCase):
    module_name = "modBuildReports"
    vba = gen_reports.build_reports_vba()


class GeneratedSalesTestsModuleTests(VbaModuleChecks, unittest.TestCase):
    module_name = "modTestSales"
    vba = gen_test_sales.build_test_sales_vba()

    def test_every_pricing_case_is_replayed(self):
        for case in P.CASES:
            self.assertIn("حساب: " + case[0], self.vba)


def _static(name):
    from helpers import ROOT
    import os
    class T(VbaModuleChecks, unittest.TestCase):
        module_name = name
        with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
            vba = fh.read()
    T.__name__ = T.__qualname__ = f"Static_{name}"
    return T


Static_modZatca = _static("modZatca")
Static_modSales = _static("modSales")
Static_modPOS = _static("modPOS")

if __name__ == "__main__":
    unittest.main()
