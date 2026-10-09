"""Onboarding of the device with ZATCA and sending (modZatcaApi, docs/47-ZATCA-Onboarding-Sending.md).

tools/zatca_api_reference.py mirrors the CSR configuration and the reading of ZATCA's replies. OpenSSL builds a CSR
from that configuration and cryptography checks what ZATCA reads in it; the VBA is compared with the reference in
LibreOffice; the code checks hold the API rules (endpoints, headers, secrets, the order of the steps)."""
import os
import re
import shutil
import subprocess
import tempfile
import unittest

from helpers import ROOT, VbaModuleChecks
import generate as G
import schema
import vba_harness as H
import zatca_api_reference as A

try:
    from cryptography import x509
    from cryptography.x509.oid import NameOID, ObjectIdentifier
except ImportError:                                    # optional (requirements-dev.txt)
    x509 = None


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


VBA = read("modZatcaApi")
CSR_ARGS = ("SIMULATION", "3cf5ee18-ee25-44ea-a444-2c37ba7f28be", "399999999900003", "1234 King Fahd, Olaya, Riyadh",
            "Retail #1 $x", "EGS-3cf5ee18", "Main  branch", "مؤسسة النور")
REPLIES = [
    (0, "", True), (200, "{}", True), (200, "{}", False),
    (202, '{"validationResults": {"warningMessages": [{"code": "BR-KSA-08", "message": "w1"}, '
          '{"code": "BR-KSA-09", "message": "w2"}]}}', False),
    (400, '{"validationResults": {"errorMessages": [{"code": "BR-01", "message": "e1"}, {"code": "BR-02", "message": "e2"},'
          ' {"code": "BR-03", "message": "e3"}, {"code": "BR-04", "message": "e4"}]}}', True),
    (400, "Bad request", False), (401, "", False), (403, "", True), (429, "", False), (503, "", False), (404, "nf", False)]


def body(name):
    return re.search(r"Public (Function|Sub) " + name + r"\b.*?End (Function|Sub)", VBA, re.S).group(0)


class ReferenceTests(unittest.TestCase):

    def test_environments(self):
        self.assertEqual(set(A.BASE_URLS), {"TEST", "SIMULATION", "PRODUCTION"})
        self.assertTrue(A.BASE_URLS["PRODUCTION"].endswith("/e-invoicing/core"))
        self.assertEqual(A.TEMPLATE_NAMES["TEST"], "TSTZATCA-Code-Signing")
        self.assertEqual(len(A.SAMPLES), 6)

    def test_config_value(self):
        self.assertEqual(A.config_value("  a #b\n$c\t d "), "a b c d")

    def test_reply_results(self):
        self.assertEqual(A.reply_result(200, "{}", True)[:2], ("OK", "CLEARED"))
        self.assertEqual(A.reply_result(200, "{}", False)[:2], ("OK", "REPORTED"))
        self.assertEqual(A.reply_result(400, REPLIES[4][1], True), ("REJECTED", "REJECTED", "BR-01: e1; BR-02: e2; BR-03: e3"))
        self.assertEqual(A.reply_result(503, "", False)[:2], ("NETWORK", ""))     # stays pending: sent again later
        self.assertEqual(A.reply_result(401, "", False)[:2], ("ERROR", ""))


@unittest.skipUnless(shutil.which("openssl") and x509, "openssl or cryptography is not installed")
class CsrTests(unittest.TestCase):
    """OpenSSL makes the CSR ZATCA expects from the configuration."""

    def test_csr_fields(self):
        with tempfile.TemporaryDirectory() as d:
            cnf, key, csr = (os.path.join(d, n) for n in ("c.cnf", "k.pem", "r.csr"))
            with open(cnf, "w", encoding="utf-8", newline="\n") as fh:
                fh.write(A.csr_config(*CSR_ARGS))
            subprocess.run(["openssl", "ecparam", "-name", "secp256k1", "-genkey", "-noout", "-out", key], check=True,
                           capture_output=True)
            subprocess.run(["openssl", "req", "-new", "-sha256", "-key", key, "-config", cnf, "-out", csr], check=True,
                           capture_output=True)
            with open(csr, "rb") as fh:
                req = x509.load_pem_x509_csr(fh.read())
        self.assertTrue(req.is_signature_valid)
        subject = {a.oid: a.value for a in req.subject}
        self.assertEqual(subject[NameOID.COUNTRY_NAME], "SA")
        self.assertEqual(subject[NameOID.ORGANIZATION_NAME], "مؤسسة النور")
        self.assertEqual(subject[NameOID.ORGANIZATIONAL_UNIT_NAME], "Main branch")
        self.assertEqual(subject[NameOID.COMMON_NAME], "EGS-3cf5ee18")
        template = req.extensions.get_extension_for_oid(ObjectIdentifier("1.3.6.1.4.1.311.20.2")).value.value
        self.assertIn(b"PREZATCA-Code-Signing", template)
        san = req.extensions.get_extension_for_class(x509.SubjectAlternativeName).value
        names = {a.oid.dotted_string: a.value for a in san.get_values_for_type(x509.DirectoryName)[0]}
        self.assertEqual(names["2.5.4.4"], "1-RetailStoreAccess|2-1.0|3-3cf5ee18-ee25-44ea-a444-2c37ba7f28be")   # SN
        self.assertEqual(names["0.9.2342.19200300.100.1.1"], "399999999900003")                                  # UID
        self.assertEqual(names["2.5.4.12"], "1100")                                                              # title
        self.assertEqual(names["2.5.4.26"], "1234 King Fahd, Olaya, Riyadh")                                     # address
        self.assertEqual(names["2.5.4.15"], "Retail 1 x")                                    # no comment nor variable


DRIVER = '''Option Explicit
Public Function RunCsr(ByVal Env As String, ByVal Serial As String, ByVal Vat As String, ByVal Location As String, _
                       ByVal Industry As String, ByVal CommonName As String, ByVal Branch As String, _
                       ByVal Organization As String) As String
    RunCsr = ZatcaCsrConfig(Env, Serial, Vat, Location, Industry, CommonName, Branch, Organization)
End Function
Public Function RunReply(ByVal HttpStatus As Long, ByVal Body As String, ByVal IsStandard As Boolean) As String
    Dim st As String, msg As String, r As String
    r = ZatcaReplyResult(HttpStatus, Body, IsStandard, st, msg)
    RunReply = r & "|" & st & "|" & msg
End Function
Public Function RunValue(ByVal s As String) As String
    RunValue = ZatcaConfigValue(s)
End Function
Public Function RunMask(ByVal s As String) As String
    RunMask = MaskSecret(s)
End Function
Public Function RunCleared(ByVal s As String) As String
    RunCleared = ClearedQR(s)
End Function
Public Function RunUrl(ByVal Env As String) As String
    RunUrl = ZatcaBaseUrl(Env) & "|" & ZatcaTemplateName(Env)
End Function
'''


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class VbaRuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modCommon": H.read_module("modCommon"), "modHttp": H.read_module("modHttp"),
                    "modZatcaApi": VBA, "Driver": DRIVER})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_csr_config(self):
        self.assertEqual(self.h.call("Driver", "RunCsr", *CSR_ARGS), A.csr_config(*CSR_ARGS))
        for env in A.BASE_URLS:
            self.assertEqual(self.h.call("Driver", "RunUrl", env), A.BASE_URLS[env] + "|" + A.TEMPLATE_NAMES[env])

    def test_config_value(self):
        for s in ("  a #b\n$c\t d ", "x", "", "a\r\nb"):
            self.assertEqual(self.h.call("Driver", "RunValue", s), A.config_value(s))

    def test_replies(self):
        for status, text, standard in REPLIES:
            with self.subTest(status=status, standard=standard):
                self.assertEqual(self.h.call("Driver", "RunReply", status, text, standard),
                                 "|".join(A.reply_result(status, text, standard)))

    def test_secret_is_masked_and_cleared_qr(self):
        reply = '{"requestID": 123, "binarySecurityToken": "TUlJQ", "secret": "Xy9+abc="}'
        self.assertEqual(self.h.call("Driver", "RunMask", reply), reply.replace("Xy9+abc=", "***"))
        self.assertEqual(self.h.call("Driver", "RunMask", "{}"), "{}")
        xml = ('<cac:AdditionalDocumentReference><cbc:ID>QR</cbc:ID><cac:Attachment><cbc:EmbeddedDocumentBinaryObject '
               'mimeCode="text/plain">AQ1Z</cbc:EmbeddedDocumentBinaryObject></cac:Attachment>')
        self.assertEqual(self.h.call("Driver", "RunCleared", xml), "AQ1Z")
        self.assertEqual(self.h.call("Driver", "RunCleared", "<x/>"), "")


class CodeTests(unittest.TestCase):

    def test_post_headers(self):
        post = body("ZatcaPost")
        for h in ('"Accept-Version: V2"', '"Content-Type: application/json"', '"Authorization: Basic " & Base64Text(Token & ":" & Secret)'):
            self.assertIn(h, post)
        self.assertIn("ZatcaBaseUrl(CurrentEnv()) & Path", post)

    def test_onboarding_steps(self):
        step1 = body("ZatcaRequestCompliance")
        self.assertIn('AppCountry() <> "SA"', step1)
        self.assertIn('"/compliance"', step1)
        self.assertIn('"OTP: " & Trim$(Otp)', step1)
        self.assertIn("Base64Text(csrPem)", step1)
        self.assertIn('Format$(Now, "yyyymmddhhnnss")', step1)               # the registered key is not overwritten
        self.assertLess(step1.index('"/compliance"'), step1.index('SaveSettings Array("ZatcaDeviceSerial"'))
        self.assertIn('"ZatcaOnboardStage", "COMPLIANCE"', step1)
        self.assertIn('"ZatcaProductionToken", Null', step1)                 # a new registration starts again
        step2 = body("ZatcaComplianceChecks")
        self.assertIn('"/compliance/invoices"', step2)
        self.assertEqual(step2.count('"0100000"') + step2.count('"0200000"'), 6)
        self.assertIn("pih = hash", step2)
        self.assertLess(step2.index("If failed > 0"), step2.index('SaveSettings Array("ZatcaOnboardStage", "CHECKED")'))
        step3 = body("ZatcaRequestProduction")
        self.assertIn('<> "CHECKED" Then', step3)
        self.assertIn('"/production/csids"', step3)
        self.assertIn('"compliance_request_id"', step3)
        self.assertIn('"ZatcaCertificate", TokenCertificate(', step3)
        self.assertIn('"ZatcaOnboardStage", "PRODUCTION"', step3)
        ui = body("ZatcaOnboardStep")
        self.assertIn('HasPermission("SETTINGS")', ui)
        save = re.search(r"Private Sub SaveSettings.*?End Sub", VBA, re.S).group(0)
        self.assertLess(save.index("AuditSnapshot("), save.index("rs.Update"))
        self.assertIn('AuditEdited "EDIT", "Settings"', save)

    def test_secrets_are_not_logged(self):
        for name in ("ZatcaRequestCompliance", "ZatcaRequestProduction"):
            b = body(name)
            log = b[b.index("LogEInvoice "):b.index("\n    If ", b.index("LogEInvoice "))]
            self.assertIn("MaskSecret(reply)", log, name)
            self.assertNotRegex(log, r"[,(]\s*reply\s*,", name)
        self.assertNotIn("LogEInvoice", body("ZatcaPost"))                   # the Authorization header is not logged

    def test_sending(self):
        send = body("ZatcaSendDocument")
        self.assertLess(send.index("ZatcaPrepareDocument(DocKind, DocID)"), send.index("ZatcaPost("))
        self.assertIn('"/invoices/clearance/single"', send)
        self.assertIn('"/invoices/reporting/single"', send)
        clear = send[send.index('"/invoices/clearance/single"'):send.index('"/invoices/reporting/single"')]
        self.assertIn('"Clearance-Status: 1"', clear)                         # standard documents only
        self.assertEqual(send.count("Clearance-Status"), 1)
        self.assertIn("ZatcaReplyResult(status, reply, standard, docStatus, msg)", send)
        self.assertLess(send.index("LogEInvoice DocKind"), send.index("SetEInvoiceStatus"))
        self.assertIn('JsonGet(reply, "clearedInvoice")', send)
        self.assertIn("rs!EInvoiceXml = xml", send)                          # the cleared document is the invoice
        ready = body("ZatcaReadyProblem")
        for s in ('<> "PRODUCTION"', "ZatcaOnboardEnv", "EInvoiceEnvironment", "Dir$("):
            self.assertIn(s, ready)

    def test_schema_and_screen(self):
        fields = {f.name for f in schema.table("Settings").fields}
        for f in ("ZatcaDeviceSerial", "ZatcaBranchName", "ZatcaIndustry", "ZatcaCsr", "ZatcaComplianceToken",
                  "ZatcaComplianceSecret", "ZatcaRequestId", "ZatcaProductionToken", "ZatcaProductionSecret",
                  "ZatcaOnboardEnv", "ZatcaOnboardStage"):
            self.assertIn(f, fields)
        form = read("modBuildForms") if os.path.exists(os.path.join(ROOT, "src", "vba", "modBuildForms.bas")) else ""
        import forms_einvoice
        m = forms_einvoice.layout_zatca_setup()
        names = {c.name for c in m.controls}
        for c in ("txtOtp", "txtBranch", "txtIndustry", "lblStage", "btnStep1", "btnStep2", "btnStep3"):
            self.assertIn(c, names)
        for n in (1, 2, 3):
            self.assertIn(f"    ZatcaOnboardStep Me, {n}", m.code)
        self.assertIn("    ZatcaOnboardShow Me", m.code)
        self.assertTrue(form)

    def test_registered(self):
        self.assertIn('"TestZatcaApi"', read("modTestAll"))
        self.assertIn("modZatcaApi", G.STATIC_MODULES)


class StaticZatcaApi(VbaModuleChecks, unittest.TestCase):
    module_name = "modZatcaApi"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
