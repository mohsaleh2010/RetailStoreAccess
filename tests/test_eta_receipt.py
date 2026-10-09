"""The Egyptian e-receipt (modEtaReceipt, docs/48-ETA-EReceipt.md).

tools/eta_receipt_reference.py is checked against the ETA rules (the serialization algorithm, a valid JSON, the
UUID = SHA-256 of the serialization); the VBA is compared with it in LibreOffice; the code checks hold the API
rules (endpoints, token, secrets, chain)."""
import hashlib
import json
import os
import re
import unittest
import urllib.parse

from helpers import ROOT, VbaModuleChecks
import generate as G
import schema
import vba_harness as H
import eta_receipt_reference as E


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


VBA = read("modEtaReceipt")
EINVOICE = read("modEInvoice")
SELLER = dict(rin="123456789", trade_name='محل "النور" & أولاده', branch_code="0", governate="القاهرة", city="مدينة نصر",
              street="عباس العقاد", building="12", postal="11765", device_serial="POS-0001", activity_code="4711")
ITEMS = [dict(internal_code="P-1", description="شاي\tأخضر", item_type="EGS", item_code="EG-123456789-1001",
              unit_type="EA", quantity=2, unit_price=50.4348, total_sale=100.87, discount=0.87, net_sale=100,
              tax=14, rate=14, total=114),
         dict(internal_code="P-2", description="Sugar 1kg", item_type="GS1", item_code="6221234567890",
              unit_type="KGM", quantity=1.5, unit_price=20, total_sale=30, discount=0, net_sale=30, tax=4.2,
              rate=14, total=34.2)]
SALE = dict(SELLER, date_time="2026-10-09T08:15:30Z", number="INV-00012", previous_uuid="a" * 64, reference_uuid="",
            buyer_type="P", buyer_id="", buyer_name="", buyer_mobile="", items=ITEMS, total_sales=130.87,
            total_discount=0.87, net_amount=130, tax_total=18.2, total_amount=148.2, payment_method="C")
RETURN = dict(SALE, number="RET-00003", previous_uuid="b" * 64, reference_uuid="c" * 64, buyer_type="B",
              buyer_id="987654321", buyer_name="شركة الأمل", buyer_mobile="01012345678", items=ITEMS[1:],
              total_sales=30, total_discount=0, net_amount=30, tax_total=4.2, total_amount=34.2, payment_method="V")


def body(name):
    return re.search(r"Public (Function|Sub) " + name + r"\b.*?\nEnd (Function|Sub)", VBA, re.S).group(0)


class ReferenceTests(unittest.TestCase):

    def test_numbers(self):
        self.assertEqual([E.num(x) for x in (12.5, 100, "0.00125", 0, -3.10, "1.000004", 50.4348)],
                         ["12.5", "100", "0.00125", "0", "-3.1", "1", "50.4348"])

    def test_serialization_algorithm(self):
        """Names upper case; an array is its name then, per element, the name and the element."""
        text = '{"a":"x y","b":[{"c":2},{"c":3.5}],"d":{"e":""},"f":[]}'
        self.assertEqual(E.serialize(E.parse(text)), '"A""x y""B""B""C""2""B""C""3.5""D""E""""F"')

    def test_receipts_are_json_and_hash_their_serialization(self):
        for r in (SALE, RETURN):
            uuid, text = E.signed_receipt(r)
            doc = json.loads(text)
            self.assertEqual(doc["header"]["uuid"], uuid)
            self.assertRegex(uuid, "^[0-9a-f]{64}$")
            doc["header"]["uuid"] = ""
            unsigned = E.receipt_json(r)
            self.assertEqual(json.loads(unsigned), doc)
            self.assertEqual(uuid, hashlib.sha256(E.serialize(E.parse(unsigned)).encode()).hexdigest())
        sale, ret = json.loads(E.receipt_json(SALE)), json.loads(E.receipt_json(RETURN))
        self.assertEqual(sale["documentType"], {"receiptType": "S", "typeVersion": "1.2"})
        self.assertEqual(ret["documentType"]["receiptType"], "R")
        self.assertNotIn("referenceUUID", sale["header"])
        self.assertEqual(ret["header"]["referenceUUID"], "c" * 64)
        self.assertEqual(sale["itemData"][0]["commercialDiscountData"], [{"amount": 0.87, "description": "Discount"}])
        self.assertNotIn("commercialDiscountData", sale["itemData"][1])
        self.assertEqual(sale["itemData"][0]["taxableItems"], [{"taxType": "T1", "amount": 14, "subType": "V009", "rate": 14}])
        self.assertEqual(sale["seller"]["branchAddress"]["country"], "EG")
        self.assertEqual(sale["buyer"], {"type": "P"})
        self.assertEqual(ret["buyer"]["id"], "987654321")
        self.assertEqual(sale["taxTotals"], [{"taxType": "T1", "amount": 18.2}])

    def test_qr_link(self):
        self.assertEqual(E.qr_link("PROD", "f" * 64, "2026-10-09T08:15:30Z", 148.2, "123456789"),
                         "https://invoicing.eta.gov.eg/receipts/search/" + "f" * 64 +
                         "/share/2026-10-09T08:15:30Z#Total:148.2,IssuerRIN:123456789")

    def test_replies(self):
        ok = '{"submissionId": "SUB1", "acceptedDocuments": [{"uuid": "u", "receiptNumber": "1"}], "rejectedDocuments": []}'
        self.assertEqual(E.submit_result(202, ok), ("OK", "SUBMITTED", "", "SUB1"))
        bad = ('{"submissionId": null, "acceptedDocuments": [], "rejectedDocuments": [{"uuid": "u", "error": '
               '{"code": "Invalid", "message": "Validation error", "details": [{"propertyPath": "itemData[0].itemCode", '
               '"message": "Item code not found"}]}}]}')
        self.assertEqual(E.submit_result(202, bad)[:3],
                         ("REJECTED", "REJECTED", "Validation error (itemData[0].itemCode: Item code not found)"))
        self.assertEqual(E.submit_result(503, "")[:2], ("NETWORK", ""))
        self.assertEqual(E.status_result('{"receipts": [{"uuid": "u", "status": "Valid"}], "status": "Valid"}'),
                         ("VALID", ""))
        self.assertEqual(E.status_result('{"receipts": [{"status": "Invalid", "errors": [{"propertyPath": "x", '
                                         '"message": "m"}]}]}'), ("INVALID", "x: m"))
        self.assertEqual(E.status_result('{"status": "InProgress"}'), ("", ""))


REPLIES = [
    (0, ""),
    (202, '{"submissionId": "SUB1", "acceptedDocuments": [{"uuid": "u", "receiptNumber": "1"}], "rejectedDocuments": []}'),
    (202, '{"acceptedDocuments": [], "rejectedDocuments": [{"uuid": "u", "error": {"code": "Invalid", "message": '
          '"Validation error", "details": [{"propertyPath": "itemData[0].itemCode", "message": "Item code not found"}, '
          '{"target": "buyer", "message": "m2"}, {"code": "C3", "message": "m3"}, {"code": "C4", "message": "m4"}]}}]}'),
    (202, '{"acceptedDocuments": [], "rejectedDocuments": [{"uuid": "u", "error": {"code": "E1"}}]}'),
    (202, '{"acceptedDocuments": [], "rejectedDocuments": [{"uuid": "u"}]}'),
    (202, "{}"),
    (400, '{"error": {"message": "Bad structure", "details": [{"propertyPath": "header", "message": "x"}]}}'),
    (400, "plain"), (401, ""), (403, ""), (429, ""), (500, "oops"), (404, "nf")]
STATUSES = ['{"receipts": [{"status": "Valid"}]}', '{"status": "valid"}', '{"status": "InProgress"}', "{}",
            '{"receipts": [{"status": "Invalid", "errors": [{"propertyPath": "a", "message": "b"}, {"code": "c", '
            '"message": "d"}]}]}', '{"receipts": [{"status": "Invalid"}]}']

DRIVER = '''Option Explicit
Public Function RunNum(ByVal v As Variant) As String
    RunNum = EtaNum(v)
End Function
Public Function RunItem(ByVal a1 As String, ByVal a2 As String, ByVal a3 As String, ByVal a4 As String, ByVal a5 As String, _
                        ByVal q As Double, ByVal p As Double, ByVal ts As Double, ByVal d As Double, ByVal ns As Double, _
                        ByVal t As Double, ByVal r As Double, ByVal tot As Double) As String
    RunItem = EtaItemJson(a1, a2, a3, a4, a5, q, p, ts, d, ns, t, r, tot)
End Function
Public Function RunSeller(ByVal a1 As String, ByVal a2 As String, ByVal a3 As String, ByVal a4 As String, ByVal a5 As String, _
                          ByVal a6 As String, ByVal a7 As String, ByVal a8 As String, ByVal a9 As String, ByVal a10 As String) As String
    RunSeller = EtaSellerJson(a1, a2, a3, a4, a5, a6, a7, a8, a9, a10)
End Function
Public Function RunBuyer(ByVal a1 As String, ByVal a2 As String, ByVal a3 As String, ByVal a4 As String) As String
    RunBuyer = EtaBuyerJson(a1, a2, a3, a4)
End Function
Public Function RunReceipt(ByVal dt As String, ByVal num As String, ByVal prev As String, ByVal ref As String, _
                           ByVal seller As String, ByVal buyer As String, ByVal items As String, ByVal ts As Double, _
                           ByVal td As Double, ByVal na As Double, ByVal tt As Double, ByVal ta As Double, _
                           ByVal pm As String) As String
    RunReceipt = EtaReceiptJson(dt, num, "", prev, ref, seller, buyer, items, ts, td, na, tt, ta, pm)
End Function
Public Function RunSerialize(ByVal s As String) As String
    RunSerialize = EtaSerialize(s)
End Function
Public Function RunUuid(ByVal s As String) As String
    RunUuid = EtaReceiptUuid(s)
End Function
Public Function RunWithUuid(ByVal s As String, ByVal u As String) As String
    RunWithUuid = EtaWithUuid(s, u)
End Function
Public Function RunQr(ByVal env As String, ByVal u As String, ByVal dt As String, ByVal total As Double, ByVal rin As String) As String
    RunQr = EtaQrLink(env, u, dt, total, rin)
End Function
Public Function RunSubmit(ByVal st As Long, ByVal b As String) As String
    Dim ds As String, msg As String, sid As String, r As String
    r = EtaSubmitResult(st, b, ds, msg, sid)
    RunSubmit = r & "|" & ds & "|" & msg & "|" & sid
End Function
Public Function RunStatus(ByVal b As String) As String
    Dim msg As String, r As String
    r = EtaStatusResult(b, msg)
    RunStatus = r & "|" & msg
End Function
Public Function RunUrlEncode(ByVal s As String) As String
    RunUrlEncode = EtaUrlEncode(s)
End Function
Public Function RunMask(ByVal s As String) As String
    RunMask = EtaMaskToken(s)
End Function
Public Function RunPay(ByVal t As String, ByVal m As Long) As String
    RunPay = EtaPaymentCode(t, m)
End Function
Public Function RunEnv(ByVal e As String) As String
    RunEnv = EtaEnv(e) & "|" & EtaTokenUrl(EtaEnv(e)) & "|" & EtaApiUrl(EtaEnv(e)) & "|" & EtaPortalUrl(EtaEnv(e))
End Function
'''


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class VbaRuntimeTests(unittest.TestCase):
    """The VBA builds the same receipt, serialization and UUID as the reference."""

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modCommon": H.read_module("modCommon"), "modZatca": H.read_module("modZatca"),
                    "modSecurity": H.read_module("modSecurity"), "modHttp": H.read_module("modHttp"),
                    "modEtaReceipt": VBA, "Driver": DRIVER})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def vba_receipt(self, r):
        items = ",".join(self.h.call("Driver", "RunItem", it["internal_code"], it["description"], it["item_type"],
                                     it["item_code"], it["unit_type"], float(it["quantity"]), float(it["unit_price"]),
                                     float(it["total_sale"]), float(it["discount"]), float(it["net_sale"]),
                                     float(it["tax"]), float(it["rate"]), float(it["total"])) for it in r["items"])
        seller = self.h.call("Driver", "RunSeller", r["rin"], r["trade_name"], r["branch_code"], r["governate"], r["city"],
                             r["street"], r["building"], r["postal"], r["device_serial"], r["activity_code"])
        buyer = self.h.call("Driver", "RunBuyer", r["buyer_type"], r["buyer_id"], r["buyer_name"], r["buyer_mobile"])
        return self.h.call("Driver", "RunReceipt", r["date_time"], r["number"], r["previous_uuid"], r["reference_uuid"],
                           seller, buyer, items, float(r["total_sales"]), float(r["total_discount"]),
                           float(r["net_amount"]), float(r["tax_total"]), float(r["total_amount"]), r["payment_method"])

    def test_numbers(self):
        for x in (12.5, 100.0, 0.00125, 0.0, -3.1, 50.4348, 148.2, 0.87, 1e6):
            self.assertEqual(self.h.call("Driver", "RunNum", x), E.num(x), x)

    def test_receipts_serialization_and_uuid(self):
        for r in (SALE, RETURN, dict(SALE, postal="", branch_code="7")):
            text = self.vba_receipt(r)
            self.assertEqual(text, E.receipt_json(r))
            self.assertEqual(self.h.call("Driver", "RunSerialize", text), E.serialize(E.parse(text)))
            uuid = self.h.call("Driver", "RunUuid", text)
            self.assertEqual(uuid, E.receipt_uuid(text))
            self.assertEqual(self.h.call("Driver", "RunWithUuid", text, uuid), E.receipt_json(r, uuid))
        escaped = '{"a":"q\\"b\\\\c\\n\\u0041","b":[{"c":-1.5},{"c":true}],"d":{}}'
        self.assertEqual(self.h.call("Driver", "RunSerialize", escaped), E.serialize(E.parse(escaped)))

    def test_qr_environments_and_payment(self):
        self.assertEqual(self.h.call("Driver", "RunQr", "PREPROD", "f" * 64, "2026-10-09T08:15:30Z", 148.2, "123456789"),
                         E.qr_link("PREPROD", "f" * 64, "2026-10-09T08:15:30Z", 148.2, "123456789"))
        for env in ("TEST", "SIMULATION", "PRODUCTION"):
            e = E.eta_env(env)
            self.assertEqual(self.h.call("Driver", "RunEnv", env),
                             "|".join([e, E.TOKEN_URLS[e], E.API_URLS[e], E.PORTAL_URLS[e]]))
        for t, m in (("CASH", 1), ("CASH", 2), ("CASH", 3), ("CREDIT", 1), ("CASH", 9)):
            self.assertEqual(self.h.call("Driver", "RunPay", t, m), E.payment_code(t, m))

    def test_replies(self):
        for status, text in REPLIES:
            with self.subTest(status=status, body=text[:40]):
                self.assertEqual(self.h.call("Driver", "RunSubmit", status, text), "|".join(E.submit_result(status, text)))
        for text in STATUSES:
            with self.subTest(body=text[:40]):
                self.assertEqual(self.h.call("Driver", "RunStatus", text), "|".join(E.status_result(text)))

    def test_url_encode_and_token_mask(self):
        for s in ("abc-._~", "a b&c=d+e/f", "سر#1"):
            self.assertEqual(self.h.call("Driver", "RunUrlEncode", s), urllib.parse.quote(s, safe="-._~"))
        reply = '{"access_token": "eyJhbGci.x.y", "expires_in": 3600, "token_type": "Bearer"}'
        self.assertEqual(self.h.call("Driver", "RunMask", reply), reply.replace("eyJhbGci.x.y", "***"))


class CodeTests(unittest.TestCase):

    def test_in_access_test_uuid(self):
        uuid = re.search(r'TEST_RECEIPT_UUID As String = "([0-9a-f]{64})"', VBA).group(1)
        item = dict(internal_code="P1", description="Tea", item_type="EGS", item_code="EG-123456789-1", unit_type="EA",
                    quantity=2, unit_price=10, total_sale=20, discount=0, net_sale=20, tax="2.8", rate=14, total="22.8")
        r = dict(date_time="2026-10-09T10:00:00Z", number="INV-1", previous_uuid="", reference_uuid="", rin="123456789",
                 trade_name="Store", branch_code="0", governate="Cairo", city="Nasr City", street="Abbas", building="12",
                 postal="", device_serial="POS-1", activity_code="4711", buyer_type="P", buyer_id="", buyer_name="",
                 buyer_mobile="", items=[item], total_sales=20, total_discount=0, net_amount=20, tax_total="2.8",
                 total_amount="22.8", payment_method="C")
        self.assertEqual(uuid, E.receipt_uuid(E.receipt_json(r)))
        self.assertIn('EtaItemJson("P1", "Tea", "EGS", "EG-123456789-1", "EA", 2, 10, 20, 0, 20, 2.8, 14, 22.8)', VBA)

    def test_token(self):
        token = body("EtaToken")
        self.assertIn('"grant_type=client_credentials&client_id=" & EtaUrlEncode(', token)
        for h in ('"posserial: "', '"pososversion: "', '"presharedkey: "', '"Content-Type: application/x-www-form-urlencoded"'):
            self.assertIn(h, token)
        log = token[token.index("LogEInvoice"):]
        self.assertIn('LogEInvoice "", 0, "TOKEN", url, status, result, Left$(msg, 255), "", EtaMaskToken(reply), ms', log)
        self.assertIn("m_tokenEnv = CurrentEtaEnv()", token)                # a token per environment

    def test_sending(self):
        send = body("EtaSendDocument")
        self.assertLess(send.index("EtaDataProblem("), send.index("EtaPrepareDocument("))
        self.assertLess(send.index("EtaPrepareDocument("), send.index("EtaToken("))
        self.assertIn('"/api/v1/receiptsubmissions"', send)
        self.assertIn('"{""receipts"":[" & json & "]}"', send)
        self.assertIn('"Authorization: Bearer " & token', send)
        self.assertLess(send.index("LogEInvoice DocKind"), send.index("SetEInvoiceStatus"))
        self.assertIn("SET EtaSubmissionId", send)
        refresh = body("EtaRefreshStatus")
        self.assertIn('"/details?PageNo=1&PageSize=10"', refresh)
        self.assertIn('"GET"', refresh)

    def test_chain(self):
        prep = body("EtaPrepareDocument")
        self.assertLess(prep.index("SELECT EtaUUID, EtaPreviousUUID, ZatcaStatus"), prep.index("ws.BeginTrans"))
        self.assertIn('Txt(rs!ZatcaStatus) = "REJECTED"', prep)                # only a rejected receipt is rebuilt
        self.assertLess(prep.index("UPDATE Settings SET EtaLastUUID = EtaLastUUID"),            # lock first
                        prep.index("SELECT EtaLastUUID FROM Settings"))
        self.assertLess(prep.index("EtaReceiptUuid(json)"), prep.index("EtaWithUuid(json, uuid)"))
        self.assertIn("rs!EInvoiceXml = json", prep)
        self.assertIn("rs!QRCodeData = EtaQrLink(", prep)
        self.assertIn("ws.CommitTrans", prep)

    def test_ready_and_data_checks(self):
        ready = body("EtaReadyProblem")
        self.assertIn('AppCountry() <> "EG"', ready)
        self.assertIn("EtaClientSecret", ready)
        problem = body("EtaDataProblem")
        self.assertIn("p.EtaItemCode IS NULL", problem)
        self.assertIn("o.EtaUUID", problem)                                   # a return after its sale
        self.assertIn("BUYER_ID_LIMIT", problem)
        seller = body("EtaSellerProblem")
        self.assertIn('TaxNumberProblem(s!VATNumber, "EG")', seller)
        save = body("EtaSetupSave")
        self.assertIn('HasPermission("SETTINGS")', save)
        self.assertLess(save.index("AuditSnapshot("), save.index("rs.Update"))

    def test_einvoice_screen(self):
        setup = re.search(r"Public Sub EInvoiceSetupOpen.*?End Sub", EINVOICE, re.S).group(0)
        self.assertIn('OpenScreen "frmEtaSetup"', setup)
        self.assertIn('OpenScreen "frmZatcaSetup"', setup)
        refresh = re.search(r"Public Function RefreshSubmittedEInvoices.*?End Function", EINVOICE, re.S).group(0)
        self.assertIn("EStatus = 'SUBMITTED'", refresh)
        self.assertIn('Application.Run("EtaRefreshStatus"', refresh)

    def test_schema_screens_and_registration(self):
        def names(t):
            return {f.name for f in schema.table(t).fields}
        for t in ("SalesInvoices", "SalesReturns"):
            self.assertTrue({"EtaUUID", "EtaPreviousUUID", "EtaSubmissionId"} <= names(t))
        self.assertTrue({"EtaClientId", "EtaClientSecret", "EtaPosSerial", "EtaPosOsVersion", "EtaPreSharedKey",
                         "EtaBranchCode", "EtaActivityCode", "EtaGovernate", "EtaLastUUID"} <= names("Settings"))
        self.assertTrue({"EtaItemType", "EtaItemCode"} <= names("Products"))
        self.assertIn("EtaUnitCode", names("Units"))
        self.assertIn("NationalID", names("Customers"))
        import forms_einvoice
        m = forms_einvoice.layout_eta_setup()
        ctl = {c.name: c for c in m.controls}
        for c in ("txtClientId", "txtClientSecret", "txtPosSerial", "txtPosOs", "txtPreSharedKey", "txtBranchCode",
                  "txtActivityCode", "txtGovernate", "lblReady"):
            self.assertIn(c, ctl)
        self.assertEqual(ctl["txtClientSecret"].props.get("InputMask"), "Password")
        self.assertIn("    EtaSetupLoad Me", m.code)
        self.assertIn('"TestEtaReceipt"', read("modTestAll"))
        self.assertIn("modEtaReceipt", G.STATIC_MODULES)
        self.assertIn(("frmEtaSetup", "إعداد ربط منظومة الإيصال الإلكتروني", "النظام", "SETTINGS", False, True, False),
                      schema.SCREEN_LIST)


class StaticEtaReceipt(VbaModuleChecks, unittest.TestCase):
    module_name = "modEtaReceipt"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
