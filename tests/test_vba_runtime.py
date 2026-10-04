"""Phase 6: EXECUTE the VBA calculation modules (modZatca, modQRCode, modSales
engine) with LibreOffice Basic and compare every result with the Python
references. Skipped automatically when LibreOffice is not installed.

Run:  python3 -m unittest discover -s tests -v
"""

import datetime as dt
import random
import string
import unittest
from decimal import Decimal as D

import helpers  # noqa: F401  (puts tools/ on the import path)
import vba_harness as H
import gen_qr
import pricing as P
import qr_reference as R
import zatca_reference as Z

DRIVER = r'''
Option VBASupport 1

Public Function RunCase(ByVal spec As String, ByVal invDisc As Double, ByVal incl As Boolean) As String
    Dim lines As Variant, f As Variant, i As Long, s As String, msg As String
    CalcReset
    lines = Split(spec, ";")
    For i = 0 To UBound(lines)
        f = Split(lines(i), ",")
        CalcAddLine CCur(Val(f(0))), CCur(Val(f(1))), CCur(Val(f(2))), CCur(Val(f(3)))
    Next
    msg = CalcRun(CCur(invDisc), incl)
    If Len(msg) > 0 Then
        RunCase = "ERR:" & msg
        Exit Function
    End If
    For i = 0 To CalcCount() - 1
        s = s & Str(CalcLine(i, "UNIT")) & "|" & Str(CalcLine(i, "DISCOUNT")) & "|" & Str(CalcLine(i, "NET")) _
              & "|" & Str(CalcLine(i, "TAX")) & "|" & Str(CalcLine(i, "TOTAL")) & ";"
    Next
    RunCase = s & "H" & Str(CalcTotal("SUBTOTAL")) & "|" & Str(CalcTotal("DISCOUNT")) & "|" & _
              Str(CalcTotal("TAXABLE")) & "|" & Str(CalcTotal("TAX")) & "|" & Str(CalcTotal("TOTAL"))
End Function

Public Function RunReturn(sq As Double, st As Double, sx As Double, pq As Double, pt As Double, _
                          px As Double, q As Double) As String
    Dim n As Currency, t As Currency, tt As Currency, msg As String
    msg = ReturnAmounts(CCur(sq), CCur(st), CCur(sx), CCur(pq), CCur(pt), CCur(px), CCur(q), n, t, tt)
    RunReturn = msg & "#" & Str(tt) & "|" & Str(t)
End Function

Public Function RunSettle(total As Double, tendered As Double, credit As Boolean) As String
    Dim p As Currency, r As Currency, c As Currency, msg As String
    msg = Settle(CCur(total), CCur(tendered), credit, p, r, c)
    If Len(msg) > 0 Then
        RunSettle = "ERR"
    Else
        RunSettle = Str(p) & "|" & Str(r) & "|" & Str(c)
    End If
End Function

Public Function RunQR(ByVal seller As String, ByVal vat As String, ByVal d As Double, _
                      ByVal total As Double, ByVal tax As Double) As String
    RunQR = BuildZatcaQR(seller, vat, CDate(d), CCur(total), CCur(tax))
End Function

Public Function RunRound(ByVal v As Double, ByVal digits As Integer) As String
    RunRound = Str(RoundMoney(v, digits))
End Function
'''


def plain(d: D) -> str:
    """Decimal without trailing zeros and without exponent: 100.00 -> '100'."""
    return format(d.normalize(), "f")


def norm(text):
    text = text.strip()
    return plain(D(text)) if text else text


def ole_date(d: dt.datetime) -> float:
    delta = d - dt.datetime(1899, 12, 30)
    return delta.days + delta.seconds / 86400


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class VbaRuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modCommon": H.read_module("modCommon"), "modZatca": H.read_module("modZatca"),
                    "modQRCode": gen_qr.build_qr_vba(), "modSales": H.read_module("modSales"),
                    "Driver": DRIVER})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def call(self, module, func, *args):
        return self.h.call(module, func, *args)

    # ------------------------------------------------------------- ZATCA
    def test_official_zatca_example(self):
        """Example from the ZATCA e-invoicing guideline (seller 'Bobs Records')."""
        got = self.call("Driver", "RunQR", "Bobs Records", "310122393500003",
                        ole_date(dt.datetime(2022, 4, 25, 18, 30)), 1000.0, 150.0)
        self.assertEqual(got, "AQxCb2JzIFJlY29yZHMCDzMxMDEyMjM5MzUwMDAwMwMUMjAyMi0wNC0yNVQxNTozMDow"
                              "MFoEBzEwMDAuMDAFBjE1MC4wMA==")

    def test_arabic_seller_name(self):
        local = dt.datetime(2026, 10, 4, 12, 15)
        for seller, total, tax in [("متجر الاختبار", 1150, 150), ("مؤسسة النخبة للتجارة | Elite", 64.2, 8.37),
                                   ("ش", 0.35, 0.05)]:
            got = self.call("Driver", "RunQR", seller, "300000000000003", ole_date(local), float(total),
                            float(tax))
            exp = Z.qr_payload(seller, "300000000000003", Z.timestamp(local), total, tax)
            self.assertEqual(got, exp, seller)
            self.assertEqual(Z.decode_payload(got)[1], seller)

    def test_utf8_including_surrogate_pairs(self):
        for text in ["abc", "ريال سعودي", "café", "€", "😀 مرحبا"]:
            got = self.call("modZatca", "Base64Text", text)
            import base64
            self.assertEqual(got, base64.b64encode(text.encode("utf-8")).decode(), text)

    def test_long_seller_name_is_cut_to_255_bytes(self):
        seller = "متجر " * 80                      # 400 bytes in UTF-8
        got = self.call("Driver", "RunQR", seller, "300000000000003",
                        ole_date(dt.datetime(2026, 1, 1, 10)), 1.0, 0.13)
        name = Z.decode_payload(got)[1]
        self.assertLessEqual(len(name.encode("utf-8")), 255)
        self.assertTrue(seller.startswith(name))

    def test_amount_format(self):
        for v, exp in [(1000, "1000.00"), (0.5, "0.50"), (12.345, "12.35"), (0.004, "0.00"),
                       (1234567.891, "1234567.89")]:
            self.assertEqual(self.call("modZatca", "ZatcaAmount", v), exp, v)

    def test_timestamp_is_utc(self):
        self.assertEqual(self.call("modZatca", "ZatcaTimestamp", ole_date(dt.datetime(2026, 10, 4, 1, 5, 9))),
                         "2026-10-03T22:05:09Z")

    # ------------------------------------------------------------- QR
    def test_qr_matrices_identical_to_reference(self):
        random.seed(11)
        texts = ["A", "HELLO WORLD", Z.qr_payload("متجر الاختبار", "300000000000003",
                                                  "2026-10-04T09:15:00Z", 1150, 150)]
        texts += ["".join(random.choice(string.ascii_letters + string.digits + "+/=") for _ in range(n))
                  for n in (17, 42, 77, 106, 134, 180, 213, 260, 330, 380, 470, 560, 666)]
        for t in texts:
            with self.subTest(length=len(t)):
                matrix, version, mask = R.encode(t.encode())
                self.assertEqual(self.call("modQRCode", "QRChecksum", t, -1), R.checksum(matrix))
                self.assertEqual(self.call("modQRCode", "QRLastInfo"), f"V{version}-M{mask}")

    def test_qr_every_mask(self):
        t = "AQxCb2JzIFJlY29yZHMCDzMxMDEyMjM5MzUwMDAwMwMUMjAyMi0wNC0yNVQxNTozMDowMFoEBzEwMDAuMDAF"
        for m in range(8):
            self.assertEqual(self.call("modQRCode", "QRChecksum", t, m),
                             R.checksum(R.encode(t.encode(), mask=m)[0]), m)

    # ------------------------------------------------------------- pricing
    def test_pricing_cases(self):
        for case in P.CASES:
            label, incl, inv, lines = case
            with self.subTest(label):
                got = self.call("Driver", "RunCase", ";".join(",".join(x) for x in lines), float(inv), incl)
                self.assertFalse(got.startswith("ERR"), got)
                o = P.calc(*P.case_inputs(case)[:2], incl)
                exp = [[plain(v) for v in (l.unit_price, l.discount, l.net, l.tax, l.total)]
                       for l in o.lines]
                exp.append([plain(v) for v in (o.subtotal, o.discount, o.taxable, o.tax, o.total)])
                parts = got.split(";")
                parts[-1] = parts[-1][1:]
                self.assertEqual([[norm(v) for v in p.split("|")] for p in parts], exp)

    def test_pricing_rejects_bad_input(self):
        for spec, inv in [("0,10,0,0.15", 0), ("1,10,11,0.15", 0), ("1,10,0,0.15", 11), ("1,-1,0,0.15", 0)]:
            self.assertTrue(self.call("Driver", "RunCase", spec, float(inv), True).startswith("ERR"), spec)

    def test_return_cases(self):
        for label, (q, t, x), rets in P.RETURN_CASES:
            pq = pt = px = D(0)
            for r in rets:
                net, tax, total = P.return_amounts(D(q), D(t), D(x), pq, pt, px, D(r))
                got = self.call("Driver", "RunReturn", float(q), float(t), float(x), float(pq), float(pt),
                                float(px), float(r))
                msg, values = got.split("#")
                self.assertEqual(msg, "", label)
                self.assertEqual([norm(v) for v in values.split("|")],
                                 [plain(total), plain(tax)], f"{label} {r}")
                pq, pt, px = pq + D(r), pt + total, px + tax
        self.assertNotEqual(self.call("Driver", "RunReturn", 3.0, 29.97, 3.91, 2.0, 19.98, 2.6, 2.0)
                            .split("#")[0], "", "returning more than sold must fail")

    def test_settle(self):
        def settle(*args):
            got = self.call("Driver", "RunSettle", *args)
            return got if got == "ERR" else "|".join(norm(v) for v in got.split("|"))
        self.assertEqual(settle(87.0, 100.0, False), "87|0|13")       # cash: change 13
        self.assertEqual(settle(87.0, 50.0, False), "ERR")            # cash must be paid in full
        self.assertEqual(settle(460.0, 100.0, True), "100|360|0")     # credit: 360 on account
        self.assertEqual(settle(460.0, 0.0, True), "0|460|0")

    def test_round_money_half_up(self):
        for v, d, exp in [(2.345, 2, "2.35"), (-2.345, 2, "-2.35"), (2.344, 2, "2.34"), (0.125, 2, "0.13"),
                          (8.68695652, 4, "8.687")]:
            self.assertEqual(norm(self.call("Driver", "RunRound", v, d)), exp, v)


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class VbaCompileTests(unittest.TestCase):
    """Every VBA module of the project must at least compile (LibreOffice Basic,
    VBA mode, with the Access-only syntax neutralised by vba_harness.prepare)."""

    def test_every_module_compiles(self):
        import glob
        import os
        import re
        h = H.Harness()
        try:
            h.add_module("Ping", "Option VBASupport 1\nFunction Ping() As Long\nPing = 42\nEnd Function\n")
            h.add_module("HarnessStubs", H.STUBS)
            failures = []
            for path in sorted(glob.glob(os.path.join(H.ROOT, "src", "vba", "*.bas"))):
                name = os.path.basename(path)[:-4]
                with open(path, encoding="utf-8") as fh:
                    code = H.prepare(fh.read())
                h.add_module(name, code)
                if h.call("Ping", "Ping") != 42:
                    parts = re.split(r"\n(?=(?:Public |Private )?(?:Sub|Function) )", code)
                    bad = []
                    for part in parts[1:]:
                        h.add_module(name, parts[0] + "\n" + part)
                        if h.call("Ping", "Ping") != 42:
                            bad.append(re.match(r"(?:Public |Private )?(?:Sub|Function) (\w+)", part).group(1))
                    failures.append(f"{name}: {bad or 'declarations'}")
                h.lib.removeByName(name)
            self.assertEqual(failures, [])
        finally:
            h.close()


if __name__ == "__main__":
    unittest.main()
