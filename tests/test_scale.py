"""Phase 7B (docs/86-Windows-POS-Documents.md): the Windows POS completed.
  * scale barcodes: the settings (LabelSettings, frmLabelSettings), the Python mirror (tools/scale_reference.py), the VBA
    (modScale, run in LibreOffice against the mirror when it is installed) and the Access POS (modPOS.ScanCode), the C#
    (RetailStore.Domain.Sales.ScaleBarcode) with the same messages;
  * returns, receipts and cashier closings of the POS: ClientRef on SalesReturns / CustomerPayments / CashClosings (a
    document sent again after a lost answer is added once), the C# of PostSalesReturn / PostCashClosing in the order of
    Access, the API operations and the screens of the POS."""
import os
import re
import sys
import unittest
from decimal import Decimal as D

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, HERE)

import forms as F
import scale_reference as R
import vba_harness as H
from helpers import VbaModuleChecks
from schema import table


def read(rel):
    with open(os.path.join(ROOT, rel), encoding="utf-8-sig") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


VBA = read("src/vba/modScale.bas")
POS = read("src/vba/modPOS.bas")
WEB = "web/src/"


class MirrorTests(unittest.TestCase):

    def test_check_digit(self):
        self.assertEqual(R.check_digit("629104150021"), 3)
        self.assertEqual(R.check_digit("400638133393"), 1)            # 4006381333931, a known EAN-13

    def test_parts(self):
        code = R.make("20", "12345", 1250, 5)
        self.assertEqual(R.code_parts(code, "20,21", 5), ("12345", 1250))
        self.assertIsNone(R.code_parts(code[:12] + str((int(code[12]) + 1) % 10), "20", 5), "a wrong check digit")
        self.assertIsNone(R.code_parts(code, "21", 5), "another prefix")
        self.assertIsNone(R.code_parts(code, "", 5), "no prefix: no scale")
        self.assertIsNone(R.code_parts(code, "20", 7), "item digits from 4 to 6")
        self.assertIsNone(R.code_parts("20123450125", "20", 5), "13 digits")
        self.assertEqual(R.prefixes(" 20, x, 2 , 1234,"), ["20", "2"])

    def test_quantity(self):
        self.assertEqual(R.quantity(1250, "WEIGHT"), D("1.25"))
        self.assertEqual(R.quantity(1000, "PRICE", "3.00"), D("3.333"))
        self.assertEqual(R.quantity(1000, "PRICE", "6.00"), D("1.667"))
        self.assertEqual(R.quantity(1000, "PRICE", 0), 0)
        self.assertEqual(R.item_codes("00042"), ["00042", "42"])
        self.assertEqual(R.item_codes("12345"), ["12345"])
        self.assertEqual(R.item_codes("00000"), ["00000"])

    def test_vectors_reach_xunit(self):
        vectors = read("web/tests/RetailStore.Tests/Generated/Vectors.g.cs")
        for row in R.vector_results():
            self.assertIn(f'("{row[0]}", "{row[1]}", {row[2]}, "{row[3]}"', vectors)


class SettingsTests(unittest.TestCase):

    def test_fields(self):
        fields = {f.name: f for f in table("LabelSettings").fields}
        self.assertEqual(fields["ScalePrefixes"].size, 20)
        self.assertEqual((fields["ScaleItemDigits"].default, fields["ScaleItemDigits"].rule), ("5", "Between 4 And 6"))
        self.assertEqual((fields["ScaleValueKind"].default, fields["ScaleValueKind"].rule), ('"WEIGHT"', 'In ("WEIGHT","PRICE")'))

    def test_screen(self):
        screen = next(m for m in F.DATA_SCREENS if m.name == "frmLabelSettings")
        names = [getattr(f, "field", "") for f in screen.fields]
        for n in ("ScalePrefixes", "ScaleItemDigits", "ScaleValueKind"):
            self.assertIn(n, names)


class VbaCodeTests(unittest.TestCase):

    def test_the_pos_reads_a_scale_label_after_the_codes_of_products_and_units(self):
        scan = proc(POS, "ScanCode")
        self.assertLess(scan.index("FindProductCode(Code, pid, unit)"), scan.index("ScaleLine(Code, CartCustomerID(frm), pid, qty, msg)"))

    def test_the_price_of_a_price_label_is_the_customer_price(self):
        self.assertIn("UnitSalePrice(ProductID, Null, CustomerID)", proc(VBA, "ScaleLine"))
        self.assertIn("Int(CDec(scaleVal) / 100 / CDec(price) * 1000 + 0.5) / 1000", VBA, "rounded half up to 3 decimals")

    def test_messages_are_those_of_the_csharp(self):
        cs = read(WEB + "RetailStore.Domain/Sales/ScaleBarcode.cs")
        for m in re.findall(r'"([^"]*[؀-ۿ][^"]*)"', cs):
            with self.subTest(m):
                self.assertIn(m, VBA)

    def test_in_access_test(self):
        test = proc(VBA, "TestScale")
        self.assertIn("ws.Rollback", test)
        self.assertIn('"TestScale"', read("src/vba/modTestAll.bas"))
        self.assertIn('"modScale"', read("tools/generate.py"))


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class ScaleRuntimeTests(unittest.TestCase):
    """modScale run by LibreOffice Basic: DbValue reads the settings of the probe module."""

    PROBE = '''Option VBASupport 1
Option Explicit
Public g_Prefixes As Variant
Public g_Digits As Variant

Public Function DbValue(ByVal Sql As String) As Variant
    If InStr(Sql, "ScalePrefixes") > 0 Then
        DbValue = g_Prefixes
    ElseIf InStr(Sql, "ScaleItemDigits") > 0 Then
        DbValue = g_Digits
    Else
        DbValue = Null
    End If
End Function

Public Function SqlText(ByVal s As Variant) As String
    SqlText = "'" & s & "'"
End Function

Public Function Probe(ByVal Code As String, ByVal Prefixes As String, ByVal Digits As Long) As String
    Dim itemNo As String, v As Long
    g_Prefixes = Prefixes
    g_Digits = Digits
    If ScaleCodeParts(Code, itemNo, v) Then
        Probe = itemNo & "|" & v
    Else
        Probe = "-"
    End If
End Function
'''

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        code = re.sub(r"\n'-+\n' In-Access test.*", "\n", VBA, flags=re.S)            # the test needs Access
        cls.h.load({"modScale": code, "ScaleProbe": cls.PROBE})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_same_as_the_python_mirror(self):
        for code, setting, digits, _kind, _price in R.VECTORS:
            with self.subTest(code=code, setting=setting):
                parts = R.code_parts(code, setting, digits)
                want = "-" if parts is None else f"{parts[0]}|{parts[1]}"
                self.assertEqual(self.h.call("ScaleProbe", "Probe", code, setting, digits), want)
        for digits12 in ("629104150021", "400638133393", "209876101250", "000000000000"):
            self.assertEqual(self.h.call("modScale", "ScaleCheckDigit", digits12), R.check_digit(digits12))


class PosDocumentTests(unittest.TestCase):
    """Returns, receipts and closings of the Windows POS (docs/86)."""

    def test_a_document_is_sent_once(self):
        for t in ("SalesReturns", "CustomerPayments", "CashClosings"):
            with self.subTest(t):
                self.assertEqual(next(f for f in table(t).fields if f.name == "ClientRef").size, 40)
                self.assertTrue(any(ix.fields == ["ClientRef"] and ix.unique and ix.ignore_nulls for ix in table(t).indexes))
        pos = read(WEB + "RetailStore.Data/Services/PosService.cs")
        for check in ("r.ClientRef == reference", "p.ClientRef == reference", "c.ClientRef == reference"):
            self.assertIn(check, pos)
        self.assertIn("PostSalesReturnAsync(user, request.ToReturn(), reference)", pos)
        self.assertIn("request.ConfirmOverBalance, reference)", pos)
        self.assertIn("PostCashClosingAsync(user, request.ToClosing(), reference)", pos)

    def test_the_return_follows_the_order_of_access(self):
        code = read(WEB + "RetailStore.Data/Services/SalesService.cs")
        ret = code[code.index("public async Task<SaleResult> PostSalesReturnAsync"):code.index("private static async Task SetReturnZatcaFieldsAsync")]
        order = ["frmSalesReturn", "ScreenAction.Add", "سبب الإرجاع مطلوب", "الفاتورة الأصلية غير موجودة.", "DefaultCustomerID",
                 "سطر لا ينتمي إلى الفاتورة الأصلية.", "InvoiceCalculator.ReturnAmounts(", "حدد الكمية المرتجعة",
                 "BeginTransactionAsync", "\"SALES_RETURN\")", "SetReturnZatcaFieldsAsync", "SalesReturnDetails.Add",
                 "ApplyStockMovementAsync", "AdjustBalanceAsync", "\"SALES_RETURN\", \"SalesReturns\"", "CommitAsync"]
        self.assertEqual([ret.index(x) for x in order], sorted(ret.index(x) for x in order))
        self.assertIn("l.D.UnitCost / l.Factor", ret, "the stock comes back at the piece cost of the invoice line (docs/72)")
        self.assertIn("recalcAverage: true", ret)

    def test_the_closing_follows_the_order_of_access(self):
        code = read(WEB + "RetailStore.Data/Services/TreasuryService.cs")
        post = code[code.index("public async Task<TreasuryResult> PostCashClosingAsync"):code.index("private static async Task<CashVouchers> InsertClosingVoucherAsync")]
        order = ["ClosingForm", "ScreenAction.Add", "\"CASH_CLOSING\"", "اختر الصندوق.", "ClosingAllowedAsync", "لا تكون سالبة",
                 "\"KEEP\" ? 0", "المبلغ المرحَّل", "اختر الخزينة التي تستلم المبلغ.", "لا يمكن الترحيل إلى نفس الصندوق.",
                 "اختر جهة الترحيل.", "ClosingTotalsAsync", "BeginTransactionAsync", "NextNumberAsync(db, \"CASH_CLOSING\")", "\"SHORTAGE\"",
                 "\"ADVANCE\"", "\"OVERAGE\"", "\"TRANSFER\", transfer", "\"OWNER\", transfer", "CommitAsync", "\"CASH_CLOSING\", \"CashClosings\""]
        self.assertEqual([post.index(x) for x in order], sorted(post.index(x) for x in order))

    def test_the_operations_of_the_api(self):
        api = read(WEB + "RetailStore.Web/Api/PosApi.cs")
        for route, key in (('"/returns/{number}"', '"SALES_RETURN", "frmSalesReturn"'),
                           ('"/returns"', '"SALES_RETURN", "frmSalesReturn", ScreenAction.Add'),
                           ('"/payments"', '"CUSTOMER_PAYMENTS", "frmCustomerPayment"'),
                           ('"/closing"', '"CASH_CLOSING", "frmCashClosing"'),
                           ('"/closing/{box:int}"', '"CASH_CLOSING", "frmCashClosing"')):
            with self.subTest(route):
                at = api.index(route)
                self.assertIn(key, api[at:at + 700])

    def test_the_screens_of_the_pos(self):
        layout = read(WEB + "RetailStore.PosApp/Layout/PosLayout.razor")
        for page, problem in (("return", "ReturnProblem"), ("payment", "PaymentProblem"), ("closing", "ClosingProblem")):
            with self.subTest(page):
                self.assertIn(f'@page "/{page}"', read(WEB + f"RetailStore.PosApp/Pages/{page.capitalize()}Page.razor"))
                self.assertIn(f'yield return new("{page}",', layout)  # a screen of the app bar (redesign batch 7B)
                self.assertIn(problem, layout)
        session = read(WEB + "RetailStore.PosApp/Offline/PosSession.cs")
        closing = session[session.index("public async Task<(SaveStatus Status, string Message, PosDocument? Document)> SaveClosingAsync"):]
        self.assertLess(closing.index("SendBeforeClosingAsync()"), closing.index("SendClosingAsync("),
                        "the sales of the user reach the program before his closing")
        self.assertIn("ScanScale(code)", read(WEB + "RetailStore.PosApp/Pages/SalePage.razor"))


class Static_modScale(VbaModuleChecks, unittest.TestCase):
    module_name = "modScale"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
