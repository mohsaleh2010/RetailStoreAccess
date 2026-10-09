"""The Egyptian e-invoice to businesses (modEtaInvoice, docs/49-ETA-EInvoice.md).

tools/eta_invoice_reference.py builds valid JSON documents whose signed text is the ETA serialization; the VBA is
compared with it in LibreOffice; the code checks hold the API rules (endpoints, token, signature, cancellation)."""
import json
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
import generate as G
import schema
import vba_harness as H
import eta_invoice_reference as I
import eta_receipt_reference as E


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


VBA = read("modEtaInvoice")
RECEIPT = read("modEtaReceipt")
EINVOICE = read("modEInvoice")
ISSUER = ("issuer", "B", "123456789", 'محل "النور" & أولاده', "0", "القاهرة", "مدينة نصر", "عباس العقاد", "12", "11765")
RECEIVER = ("receiver", "B", "987654321", "شركة الأمل", "", "الجيزة", "الدقي", "التحرير", "5", "")
LINES = [dict(description="شاي\tأخضر", item_type="EGS", item_code="EG-123456789-1001", unit_type="EA", quantity=2,
              internal_code="P-1", unit_value=50.4348, sales_total=100.87, discount=0.87, net_total=100, tax=14, rate=14,
              total=114),
         dict(description="Sugar", item_type="GS1", item_code="6221234567890", unit_type="KGM", quantity=1.5,
              internal_code="P-2", unit_value=20, sales_total=30, discount=0, net_total=30, tax=4.2, rate=14, total=34.2)]


def doc(doc_type="I", version="1.0", reference=""):
    return dict(doc_type=doc_type, version=version, date_time="2026-10-09T08:15:30Z", activity_code="4711",
                internal_id="INV-00012", issuer=I.party_json(*ISSUER), receiver=I.party_json(*RECEIVER),
                lines=[I.line_json(x) for x in LINES], reference_uuid=reference, total_discount=0.87,
                total_sales=130.87, net_amount=130, tax_total=18.2, total_amount=148.2)


def body(text, name):
    return re.search(r"Public (Function|Sub) " + name + r"\b.*?\nEnd (Function|Sub)", text, re.S).group(0)


class ReferenceTests(unittest.TestCase):

    def test_document(self):
        d = json.loads(I.invoice_json(doc()))
        self.assertEqual(list(d)[:7], ["issuer", "receiver", "documentType", "documentTypeVersion", "dateTimeIssued",
                                       "taxpayerActivityCode", "internalID"])
        self.assertEqual(d["issuer"]["address"]["branchID"], "0")
        self.assertNotIn("branchID", d["receiver"]["address"])
        self.assertEqual(d["receiver"]["id"], "987654321")
        self.assertNotIn("references", d)
        line = d["invoiceLines"][0]
        self.assertEqual(line["unitValue"], {"currencySold": "EGP", "amountEGP": 50.4348})
        self.assertEqual(line["discount"], {"rate": 0, "amount": 0.87})
        self.assertEqual(line["taxableItems"], [{"taxType": "T1", "amount": 14, "subType": "V009", "rate": 14}])
        credit = json.loads(I.invoice_json(doc("C", reference="u" * 10)))
        self.assertEqual(credit["documentType"], "C")
        self.assertEqual(credit["references"], ["u" * 10])

    def test_signature_is_appended_and_not_serialized(self):
        unsigned = I.invoice_json(doc())
        signed = json.loads(I.with_signature(unsigned, "MIIG+/=="))
        self.assertEqual(signed["signatures"], [{"signatureType": "I", "value": "MIIG+/=="}])
        del signed["signatures"]
        self.assertEqual(signed, json.loads(unsigned))
        text = E.serialize(E.parse(unsigned))
        self.assertTrue(text.startswith('"ISSUER""ADDRESS""BRANCHID""0""COUNTRY""EG"'))
        self.assertIn('"INVOICELINES""INVOICELINES""DESCRIPTION"', text)
        self.assertNotIn("SIGNATURES", text)
        refs = E.serialize(E.parse(I.invoice_json(doc("C", reference="abc"))))
        self.assertIn('"REFERENCES""REFERENCES""abc"', refs)

    def test_signer_command_and_qr(self):
        self.assertEqual(I.signer_command("", r"C:\Signer\sign.exe", "in.txt", "out.txt", "1234"),
                         r'"C:\Signer\sign.exe" "in.txt" "out.txt" "1234"')
        self.assertEqual(I.signer_command("-i {IN} -o {OUT}", "s.exe", "a", "b", "p"), '"s.exe" -i a -o b')
        self.assertEqual(I.qr_link("PROD", "U1", "L1"), "https://invoicing.eta.gov.eg/documents/U1/share/L1")

    def test_replies(self):
        ok = '{"submissionId": "S1", "acceptedDocuments": [{"uuid": "U1", "longId": "L1", "internalId": "INV-1"}]}'
        self.assertEqual(I.submit_result(202, ok), ("OK", "SUBMITTED", "", "S1", "U1", "L1"))
        self.assertEqual(I.status_result('{"status": "Valid"}'), ("VALID", ""))
        self.assertEqual(I.status_result('{"status": "Cancelled"}'), ("CANCELLED", ""))
        invalid = ('{"status": "Invalid", "validationResults": {"validationSteps": [{"name": "Structure", "status": "Valid"},'
                   ' {"name": "Signature", "status": "Invalid", "error": {"error": "Bad signature", "innerError": '
                   '[{"error": "Certificate not trusted"}]}}, {"name": "Codes", "status": "Invalid", "error": '
                   '{"errorCode": "4041"}}]}}')
        self.assertEqual(I.status_result(invalid), ("INVALID", "Signature: Certificate not trusted; Codes: 4041"))


REPLIES = [
    (0, ""), (202, '{"submissionId": "S1", "acceptedDocuments": [{"uuid": "U1", "longId": "L1"}], "rejectedDocuments": []}'),
    (202, '{"acceptedDocuments": [{"uuid": "U2"}]}'),
    (202, '{"acceptedDocuments": [], "rejectedDocuments": [{"internalId": "1", "error": {"code": "Invalid", "message": '
          '"Validation error", "details": [{"propertyPath": "invoiceLines[0].itemCode", "message": "not found"}]}}]}'),
    (202, '{"rejectedDocuments": [{"internalId": "1"}]}'), (202, "{}"),
    (400, '{"error": {"message": "Bad structure"}}'), (400, "x"), (401, ""), (403, ""), (429, ""), (502, ""), (409, "dup")]
STATUSES = ['{"status": "Valid"}', '{"status": "Submitted"}', '{"status": "cancelled"}', "{}",
            '{"status": "Invalid", "validationResults": {"validationSteps": [{"name": "A", "status": "Invalid", '
            '"error": {"error": "e1"}}, {"name": "B", "status": "Invalid", "error": {"innerError": [{"error": "e2"}]}}, '
            '{"name": "C", "status": "Valid"}, {"name": "D", "status": "Invalid", "error": {"errorCode": "c4"}}, '
            '{"name": "E", "status": "Invalid", "error": {"error": "e5"}}]}}']

DRIVER = '''Option Explicit
Public Function RunParty(ByVal a1 As String, ByVal a2 As String, ByVal a3 As String, ByVal a4 As String, ByVal a5 As String, _
                         ByVal a6 As String, ByVal a7 As String, ByVal a8 As String, ByVal a9 As String, ByVal a10 As String) As String
    RunParty = EtaPartyJson(a1, a2, a3, a4, a5, a6, a7, a8, a9, a10)
End Function
Public Function RunLine(ByVal a1 As String, ByVal a2 As String, ByVal a3 As String, ByVal a4 As String, ByVal q As Double, _
                        ByVal a6 As String, ByVal uv As Double, ByVal st As Double, ByVal d As Double, ByVal nt As Double, _
                        ByVal t As Double, ByVal r As Double, ByVal tot As Double) As String
    RunLine = EtaLineJson(a1, a2, a3, a4, q, a6, uv, st, d, nt, t, r, tot)
End Function
Public Function RunDoc(ByVal dt As String, ByVal v As String, ByVal dtm As String, ByVal act As String, ByVal iid As String, _
                       ByVal iss As String, ByVal rcv As String, ByVal lns As String, ByVal ref As String, ByVal td As Double, _
                       ByVal ts As Double, ByVal na As Double, ByVal tt As Double, ByVal ta As Double) As String
    RunDoc = EtaInvoiceJson(dt, v, dtm, act, iid, iss, rcv, lns, ref, td, ts, na, tt, ta)
End Function
Public Function RunSign(ByVal s As String, ByVal sig As String) As String
    RunSign = EtaWithSignature(s, sig)
End Function
Public Function RunSerialize(ByVal s As String) As String
    RunSerialize = EtaSerialize(s)
End Function
Public Function RunCmd(ByVal t As String, ByVal e As String, ByVal i As String, ByVal o As String, ByVal p As String) As String
    RunCmd = EtaSignerCommand(t, e, i, o, p)
End Function
Public Function RunQr(ByVal env As String, ByVal u As String, ByVal l As String) As String
    RunQr = EtaInvoiceQrLink(env, u, l)
End Function
Public Function RunSubmit(ByVal st As Long, ByVal b As String) As String
    Dim ds As String, msg As String, sid As String, u As String, l As String, r As String
    r = EtaDocSubmitResult(st, b, ds, msg, sid, u, l)
    RunSubmit = r & "|" & ds & "|" & msg & "|" & sid & "|" & u & "|" & l
End Function
Public Function RunStatus(ByVal b As String) As String
    Dim msg As String, r As String
    r = EtaDocStatusResult(b, msg)
    RunStatus = r & "|" & msg
End Function
'''


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class VbaRuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modCommon": H.read_module("modCommon"), "modZatca": H.read_module("modZatca"),
                    "modSecurity": H.read_module("modSecurity"), "modHttp": H.read_module("modHttp"),
                    "modEtaReceipt": RECEIPT, "modEtaInvoice": VBA, "Driver": DRIVER})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def vba_doc(self, d):
        lines = ",".join(self.h.call("Driver", "RunLine", x["description"], x["item_type"], x["item_code"], x["unit_type"],
                                     float(x["quantity"]), x["internal_code"], float(x["unit_value"]),
                                     float(x["sales_total"]), float(x["discount"]), float(x["net_total"]),
                                     float(x["tax"]), float(x["rate"]), float(x["total"])) for x in LINES)
        issuer = self.h.call("Driver", "RunParty", *ISSUER)
        receiver = self.h.call("Driver", "RunParty", *RECEIVER)
        self.assertEqual(issuer, d["issuer"])
        self.assertEqual(receiver, d["receiver"])
        return self.h.call("Driver", "RunDoc", d["doc_type"], d["version"], d["date_time"], d["activity_code"],
                           d["internal_id"], issuer, receiver, lines, d["reference_uuid"], float(d["total_discount"]),
                           float(d["total_sales"]), float(d["net_amount"]), float(d["tax_total"]),
                           float(d["total_amount"]))

    def test_documents_signature_and_serialization(self):
        for d in (doc(), doc("C", "0.9", "f" * 64)):
            text = self.vba_doc(d)
            self.assertEqual(text, I.invoice_json(d))
            self.assertEqual(self.h.call("Driver", "RunSerialize", text), E.serialize(E.parse(text)))
            self.assertEqual(self.h.call("Driver", "RunSign", text, "MIIG+/=="), I.with_signature(text, "MIIG+/=="))

    def test_command_qr_and_replies(self):
        for t in ("", "-i {IN} -o {OUT} -p {PIN}"):
            self.assertEqual(self.h.call("Driver", "RunCmd", t, r"C:\S\s.exe", r"C:\T\in.txt", r"C:\T\out.txt", "12"),
                             I.signer_command(t, r"C:\S\s.exe", r"C:\T\in.txt", r"C:\T\out.txt", "12"))
        self.assertEqual(self.h.call("Driver", "RunQr", "PREPROD", "U", "L"), I.qr_link("PREPROD", "U", "L"))
        for status, text in REPLIES:
            with self.subTest(status=status, body=text[:40]):
                self.assertEqual(self.h.call("Driver", "RunSubmit", status, text), "|".join(I.submit_result(status, text)))
        for text in STATUSES:
            with self.subTest(body=text[:40]):
                self.assertEqual(self.h.call("Driver", "RunStatus", text), "|".join(I.status_result(text)))


class CodeTests(unittest.TestCase):

    def test_in_access_test_hash(self):
        h = re.search(r'TEST_DOCUMENT_HASH As String = "([0-9a-f]{64})"', VBA).group(1)
        d = dict(doc_type="C", version="1.0", date_time="2026-10-09T10:00:00Z", activity_code="4711", internal_id="RET-1",
                 issuer=I.party_json("issuer", "B", "123456789", "Store", "0", "Cairo", "Nasr City", "Abbas", "12", ""),
                 receiver=I.party_json("receiver", "B", "987654321", "Buyer", "", "Giza", "Dokki", "Tahrir", "5", ""),
                 lines=[I.line_json(dict(description="Tea", item_type="EGS", item_code="EG-123456789-1", unit_type="EA",
                                         quantity=2, internal_code="P1", unit_value=10, sales_total=20, discount=0,
                                         net_total=20, tax="2.8", rate=14, total="22.8"))],
                 reference_uuid="u-1", total_discount=0, total_sales=20, net_amount=20, tax_total="2.8",
                 total_amount="22.8")
        self.assertEqual(h, E.receipt_uuid(I.invoice_json(d)))
        self.assertIn('EtaLineJson("Tea", "EGS", "EG-123456789-1", "EA", 2, "P1", 10, 20, 0, 20, 2.8, 14, 22.8)', VBA)
        self.assertEqual(re.search(r'Public Const ETA_SIGNER_ARGS As String = "(.*)"', VBA).group(1).replace('""', '"'),
                         I.DEFAULT_SIGNER_ARGS)

    def test_routing(self):
        send = body(RECEIPT, "EtaSendDocument")
        self.assertLess(send.index("EtaIsInvoice(DocKind, DocID)"), send.index("EtaDataProblem("))
        self.assertIn("EtaInvoiceSend(DocKind, DocID)", send)
        self.assertIn("EtaInvoiceRefresh(DocKind, DocID)", body(RECEIPT, "EtaRefreshStatus"))
        self.assertIn('= "STANDARD")', body(VBA, "EtaIsInvoice"))

    def test_signing_and_sending(self):
        prep = body(VBA, "EtaInvoicePrepare")
        self.assertIn('st <> "REJECTED" And st <> "INVALID"', prep)            # built once, again only when refused
        self.assertLess(prep.index('EtaInvoiceDocumentJson(DocKind, DocID, "1.0")'), prep.index("EtaSignDocument(EtaSerialize(json)"))
        self.assertIn('EtaInvoiceDocumentJson(DocKind, DocID, "0.9")', prep)
        self.assertIn('CurrentEnv() = "PROD" Or', body(VBA, "EtaSignedVersion"))   # production is always signed
        self.assertIn('CurrentEnv() = "PROD" And Len(Txt(SettingValue("EtaSignerPath"))) = 0', body(VBA, "EtaInvoiceDataProblem"))
        sign = body(VBA, "EtaSignDocument")
        self.assertIn('.Run(cmd, 0, True)', sign)                                 # hidden, waits for the program
        self.assertIn("Kill inFile", sign)
        send = body(VBA, "EtaInvoiceSend")
        self.assertIn('"/api/v1.0/documentsubmissions"', send)
        self.assertIn('"{""documents"":["', send)
        self.assertLess(send.index("LogEInvoice DocKind"), send.index("SetEInvoiceStatus"))
        self.assertIn("rs!QRCodeData = EtaInvoiceQrLink(", send)
        token = body(VBA, "EtaErpToken")
        self.assertIn('LogEInvoice "", 0, "TOKEN_ERP", url, status, result, Left$(msg, 255), "", EtaMaskToken(reply), ms', token)
        self.assertNotIn("posserial", token)
        refresh = body(VBA, "EtaInvoiceRefresh")
        self.assertIn('"/api/v1.0/documents/" & uuid & "/details"', refresh)
        cancel = body(VBA, "EtaCancelDocument")
        self.assertIn('"/api/v1.0/documents/state/" & uuid & "/state"', cancel)
        self.assertIn('HttpSend("PUT"', cancel)
        self.assertIn('EtaJsonText("status", "cancelled")', cancel)
        self.assertLess(cancel.index('<> "VALID" Then'), cancel.index("HttpSend("))

    def test_screens(self):
        cancel = body(EINVOICE, "EInvoicesCancelPicked")
        self.assertIn('CanScreenAction(frm.Name, "EDIT")', cancel)
        self.assertIn('Application.Run("EtaCancelDocument"', cancel)
        import forms_einvoice
        m = forms_einvoice.layout_eta_setup()
        ctl = {c.name: c for c in m.controls}
        for c in ("txtErpClientId", "txtErpClientSecret", "txtSignerPath", "txtTokenPin", "txtSignerArgs"):
            self.assertIn(c, ctl)
        for c in ("txtErpClientSecret", "txtTokenPin"):
            self.assertEqual(ctl[c].props.get("InputMask"), "Password")
        names = [c.name for c in forms_einvoice.layout_einvoices().controls]
        self.assertIn("btnCancelDoc", names)
        fields = {f.name for f in schema.table("Settings").fields}
        self.assertTrue({"EtaErpClientId", "EtaErpClientSecret", "EtaSignerPath", "EtaSignerArgs", "EtaTokenPin"} <= fields)
        for t in ("SalesInvoices", "SalesReturns"):
            self.assertIn("EtaLongId", {f.name for f in schema.table(t).fields})
        # every control of the setup screen is loaded and saved
        controls = re.search(r"Private Function SetupControls.*?End Function", RECEIPT, re.S).group(0)
        for c in ctl:
            if c.startswith("txt"):
                self.assertIn(f'"{c}"', controls)

    def test_registered(self):
        self.assertIn('"TestEtaInvoice"', read("modTestAll"))
        self.assertIn("modEtaInvoice", G.STATIC_MODULES)


class StaticEtaInvoice(VbaModuleChecks, unittest.TestCase):
    module_name = "modEtaInvoice"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
