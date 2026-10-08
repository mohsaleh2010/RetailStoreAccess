"""The ZATCA phase 2 document (modZatcaXml, modZatcaData, docs/46-ZATCA-Invoice-XML.md).

tools/zatca_xml_reference.py is checked against the real thing: lxml C14N (the document is written canonical,
and its hash equals the DOM removal + C14N that ZATCA does), OpenSSL signatures verified by cryptography, and the
certificate fields read by cryptography. The VBA is then compared with the reference in LibreOffice."""
import base64
import datetime
import hashlib
import os
import re
import shutil
import subprocess
import tempfile
import unittest

from helpers import ROOT, VbaModuleChecks
import generate as G
import gen_zatca
import schema
import vba_harness as H
import zatca_xml_reference as Z

try:
    from lxml import etree
except ImportError:                                    # optional (requirements-dev.txt)
    etree = None
try:
    from cryptography import x509
    from cryptography.x509.oid import NameOID
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import ec
except ImportError:
    x509 = None


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


VBA = read("modZatcaXml")
SELLER = dict(name="مؤسسة النور & أولاده", vat="399999999900003", crn="1010010000", street="الملك فهد", building="1234",
              plot="5678", district="العليا", city="الرياض", postal="12345")
BUYER = dict(name="شركة <الأمل>", vat="311111111100003", street="التحلية", building="4321", district="السلامة",
             city="جدة", postal="23456")
LINE = dict(id=1, qty=2, unit_code="PCE", net="100", discount="0.87", tax="15", total="115",
            name='منتج "أ" <1>', category="S", percent=15, price="50.4348")
SAMPLE = dict(id="INV-0001", uuid="3cf5ee18-ee25-44ea-a444-2c37ba7f28be", issue_date="2026-10-08",
              issue_time="14:05:09", type_code="388", sub_type="0200000", icv=7, pih=Z.INITIAL_PIH, seller=SELLER,
              buyer=None, payment_code="10", tax_total="15", subtotals=[("100", "15", "S", 15)], line_ext="100",
              tax_excl="100", tax_incl="115", lines=[LINE])
CREDIT = dict(SAMPLE, id="RET-0002", type_code="381", sub_type="0100000", billing_ref="INV-0001", buyer=BUYER,
              delivery_date="2026-10-08", payment_code="30", reason="إرجاع بضاعة",
              lines=[LINE, dict(LINE, id=2, qty="1.5", net="30", discount="0", tax="4.5", total="34.5", price="20")],
              subtotals=[("130", "19.5", "S", 15)], line_ext="130", tax_excl="130", tax_incl="149.5", tax_total="19.5")
NS = {"ext": "urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2",
      "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
      "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"}


def dom_hash(doc: str) -> str:
    """What ZATCA does: remove the three parts from the DOM (their neighbouring text stays), C14N, SHA-256."""
    root = etree.fromstring(doc.encode())
    for n in root.xpath("//ext:UBLExtensions|//cac:Signature|//cac:AdditionalDocumentReference[cbc:ID='QR']",
                        namespaces=NS):
        parent, prev, tail = n.getparent(), n.getprevious(), n.tail or ""
        if prev is not None:
            prev.tail = (prev.tail or "") + tail
        else:
            parent.text = (parent.text or "") + tail
        parent.remove(n)
    return base64.b64encode(hashlib.sha256(etree.tostring(root, method="c14n")).digest()).decode()


def make_certificate():
    key = ec.generate_private_key(ec.SECP256K1())
    name = x509.Name([x509.NameAttribute(NameOID.COUNTRY_NAME, "SA"), x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Test"),
                      x509.NameAttribute(NameOID.COMMON_NAME, "TST-EGS")])
    now = datetime.datetime(2026, 1, 1)
    cert = (x509.CertificateBuilder().subject_name(name).issuer_name(name).public_key(key.public_key())
            .serial_number(0x1234567890ABCDEF1234567890).not_valid_before(now)
            .not_valid_after(now + datetime.timedelta(days=365)).sign(key, hashes.SHA256()))
    return key, cert, base64.b64encode(cert.public_bytes(serialization.Encoding.DER)).decode()


@unittest.skipUnless(etree, "lxml is not installed")
class CanonicalTests(unittest.TestCase):

    def test_documents_are_written_canonical(self):
        for d in (SAMPLE, CREDIT):
            xml = Z.invoice_xml(d)
            self.assertEqual(etree.tostring(etree.fromstring(xml.encode()), method="c14n").decode(), xml)

    def test_hash_is_the_dom_and_c14n_hash_before_and_after_signing(self):
        for d in (SAMPLE, CREDIT):
            xml = Z.invoice_xml(d)
            h = Z.invoice_hash(xml)
            self.assertEqual(dom_hash(xml), h)
            info = dict(issuer="CN=TST", serial="1", hash="abc")
            doc = Z.signed_document(xml, Z.ubl_extensions(h, "SIG", "CERT", "2026-10-08T11:05:10", info), "QRDATA")
            self.assertEqual(dom_hash(doc), h)
            self.assertEqual(Z.invoice_hash(doc), h)


@unittest.skipUnless(x509, "cryptography is not installed")
class CertificateAndSignatureTests(unittest.TestCase):

    def test_certificate_fields(self):
        key, cert, b64 = make_certificate()
        info = Z.certificate_info(b64)
        self.assertEqual(info["serial"], str(cert.serial_number))
        self.assertEqual(info["issuer"], "CN=TST-EGS, O=Test, C=SA")
        self.assertEqual(info["public_key"], key.public_key().public_bytes(
            serialization.Encoding.DER, serialization.PublicFormat.SubjectPublicKeyInfo))
        self.assertEqual(info["signature"], cert.signature)
        self.assertEqual(info["hash"], base64.b64encode(hashlib.sha256(b64.encode()).hexdigest().encode()).decode())

    @unittest.skipUnless(shutil.which("openssl"), "openssl is not installed")
    def test_openssl_signs_the_hash_bytes(self):
        """openssl dgst -sha256 -sign over the 32 hash bytes = ECDSA-SHA256 of the hash (what ZATCA verifies)."""
        key, _, _ = make_certificate()
        h = Z.invoice_hash(Z.invoice_xml(SAMPLE))
        with tempfile.TemporaryDirectory() as d:
            with open(os.path.join(d, "key.pem"), "wb") as fh:
                fh.write(key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.TraditionalOpenSSL,
                                           serialization.NoEncryption()))
            with open(os.path.join(d, "hash.bin"), "wb") as fh:
                fh.write(base64.b64decode(h))
            subprocess.run(["openssl", "dgst", "-sha256", "-sign", os.path.join(d, "key.pem"), "-out",
                            os.path.join(d, "sig.der"), os.path.join(d, "hash.bin")], check=True)
            sig = open(os.path.join(d, "sig.der"), "rb").read()
        key.public_key().verify(sig, base64.b64decode(h), ec.ECDSA(hashes.SHA256()))

    def test_qr_code_tags(self):
        _, _, b64 = make_certificate()
        info = Z.certificate_info(b64)
        qr = base64.b64decode(Z.qr_code("محل", "399999999900003", "2026-10-08T14:05:09", 115, 15, "HASH", "SIG",
                                        info["public_key"], info["signature"]))
        tags, pos = {}, 0
        while pos < len(qr):
            tag, n = qr[pos], qr[pos + 1]
            pos += 2
            if n == 0x81:
                n, pos = qr[pos], pos + 1
            tags[tag] = qr[pos:pos + n]
            pos += n
        self.assertEqual(sorted(tags), list(range(1, 10)))
        self.assertEqual(tags[1].decode(), "محل")
        self.assertEqual(tags[4], b"115.00")
        self.assertEqual(tags[8], info["public_key"])
        self.assertEqual(tags[9], info["signature"])


class ReferenceTests(unittest.TestCase):

    def test_initial_hash_and_templates(self):
        self.assertEqual(Z.INITIAL_PIH, schema.ZATCA_INITIAL_PIH)
        self.assertEqual(base64.b64decode(Z.INITIAL_PIH).decode(), hashlib.sha256(b"0").hexdigest())
        for t in (Z.SIGNED_PROPERTIES_FOR_HASH, Z.SIGNED_PROPERTIES_FOR_XML):
            self.assertIn('Id="xadesSignedProperties"', t)
        # the document keeps the indentation of the hashed form (the validator recomputes it)
        hashed = [ln.strip() for ln in Z.SIGNED_PROPERTIES_FOR_HASH.split("\n")]
        written = [ln.strip() for ln in Z.SIGNED_PROPERTIES_FOR_XML.split("\n")]
        self.assertEqual([re.sub(r' xmlns:\w+="[^"]*"', "", ln) for ln in hashed], written)
        self.assertEqual([len(a) - len(a.lstrip()) for a in Z.SIGNED_PROPERTIES_FOR_HASH.split("\n")],
                         [len(a) - len(a.lstrip()) for a in Z.SIGNED_PROPERTIES_FOR_XML.split("\n")])

    def test_sample_hash_of_the_in_access_test(self):
        self.assertIn(f'SAMPLE_HASH As String = "{Z.invoice_hash(Z.invoice_xml(SAMPLE))}"', VBA)

    def test_price_and_quantity(self):
        self.assertEqual([Z.price(x) for x in ("8.6957", "10", "2.5", "1.2340")], ["8.6957", "10.00", "2.50", "1.234"])
        self.assertEqual(Z.quantity("1.5"), "1.500000")


DRIVER = r'''Option VBASupport 1
Public Function RunSample() As String
    RunSample = TestSampleXml()
End Function
Public Function RunCredit() As String
    Dim lines As String
    lines = ZxLine(1, 2, "PCE", 100, 0.87, 15, 115, "منتج ""أ"" <1>", "S", 0.15, 50.4348, "SAR") & vbLf & _
            ZxLine(2, 1.5, "PCE", 30, 0, 4.5, 34.5, "منتج ""أ"" <1>", "S", 0.15, 20, "SAR")
    RunCredit = ZxInvoice("RET-0002", "3cf5ee18-ee25-44ea-a444-2c37ba7f28be", "2026-10-08", "14:05:09", "381", "0100000", _
        "SAR", "INV-0001", 7, ZATCA_INITIAL_PIH, _
        ZxParty("Supplier", "مؤسسة النور & أولاده", "399999999900003", "1010010000", "الملك فهد", "1234", "5678", "العليا", _
                "الرياض", "12345", "SA"), _
        ZxParty("Customer", "شركة <الأمل>", "311111111100003", "", "التحلية", "4321", "", "السلامة", "جدة", "23456", "SA"), _
        "2026-10-08", "30", "إرجاع بضاعة", 19.5, ZxTaxSubtotal(130, 19.5, "S", 0.15, "SAR"), 130, 130, 149.5, lines)
End Function
Public Function RunHash(ByVal Xml As String) As String
    RunHash = ZatcaInvoiceHash(Xml)
End Function
Public Function RunCert(ByVal B64 As String) As String
    Dim i As String, s As String, p As String, c As String, h As String, m As String
    m = ZatcaCertInfo(B64, i, s, p, c, h)
    RunCert = m & "|" & i & "|" & s & "|" & p & "|" & c & "|" & h
End Function
Public Function RunExt(ByVal Hash As String, ByVal Sig As String, ByVal Cert As String, ByVal T As String, _
                       ByVal CertHash As String, ByVal Issuer As String, ByVal Serial As String) As String
    RunExt = ZatcaExtensions(Hash, Sig, Cert, T, CertHash, Issuer, Serial)
End Function
Public Function RunQR(ByVal Hash As String, ByVal Sig As String, ByVal Pub As String, ByVal CertSig As String) As String
    RunQR = ZatcaQRCode("مؤسسة النور & أولاده", "399999999900003", "2026-10-08T14:05:09", 115, 15, Hash, Sig, Pub, CertSig)
End Function
Public Function RunDoc(ByVal Xml As String, ByVal Ext As String, ByVal QR As String) As String
    RunDoc = ZatcaSignedDocument(Xml, Ext, QR)
End Function
'''


@unittest.skipUnless(H.available() and x509, "LibreOffice (soffice + python3-uno) or cryptography is not installed")
class VbaRuntimeTests(unittest.TestCase):
    """The VBA builds the same bytes as the reference."""

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modCommon": H.read_module("modCommon"), "modZatca": H.read_module("modZatca"),
                    "modSecurity": H.read_module("modSecurity"), "modZatcaData": H.read_module("modZatcaData"),
                    "modZatcaXml": VBA, "Driver": DRIVER})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_documents(self):
        self.assertEqual(self.h.call("Driver", "RunSample"), Z.invoice_xml(SAMPLE))
        self.assertEqual(self.h.call("Driver", "RunCredit"), Z.invoice_xml(CREDIT))

    def test_hash(self):
        for d in (SAMPLE, CREDIT):
            xml = Z.invoice_xml(d)
            self.assertEqual(self.h.call("Driver", "RunHash", xml), Z.invoice_hash(xml))

    def test_certificate_signature_block_qr_and_file(self):
        _, _, b64 = make_certificate()
        info = Z.certificate_info(b64)
        parts = self.h.call("Driver", "RunCert", b64).split("|")
        self.assertEqual(parts, ["", info["issuer"], info["serial"], base64.b64encode(info["public_key"]).decode(),
                                 base64.b64encode(info["signature"]).decode(), info["hash"]])
        xml = Z.invoice_xml(SAMPLE)
        h = Z.invoice_hash(xml)
        ext = self.h.call("Driver", "RunExt", h, "SIGNATURE==", b64, "2026-10-08T11:05:10", info["hash"],
                          info["issuer"], info["serial"])
        self.assertEqual(ext, Z.ubl_extensions(h, "SIGNATURE==", b64, "2026-10-08T11:05:10", info))
        qr = self.h.call("Driver", "RunQR", h, "SIGNATURE==", base64.b64encode(info["public_key"]).decode(),
                         base64.b64encode(info["signature"]).decode())
        self.assertEqual(qr, Z.qr_code(SELLER["name"], SELLER["vat"], "2026-10-08T14:05:09", 115, 15, h, "SIGNATURE==",
                                       info["public_key"], info["signature"]))
        doc = self.h.call("Driver", "RunDoc", xml, ext, qr)
        self.assertEqual(doc, Z.signed_document(xml, ext, qr))
        self.assertEqual(self.h.call("Driver", "RunHash", doc), h)


class CodeTests(unittest.TestCase):

    def test_generated_templates(self):
        self.assertEqual(read("modZatcaData").rstrip("\n"), gen_zatca.build_zatca_data_vba().rstrip("\n"))

    def test_chain_is_moved_once_and_in_order(self):
        body = re.search(r"Public Function ZatcaPrepareDocument.*?End Function", VBA, re.S).group(0)
        self.assertLess(body.index("SELECT EInvoiceXml FROM"), body.index("ws.BeginTrans"))          # signed once
        self.assertLess(body.index("UPDATE Settings SET LastInvoiceHash = LastInvoiceHash"),          # lock first
                        body.index("SELECT LastInvoiceHash FROM Settings"))
        self.assertLess(body.index("rs!PreviousInvoiceHash = pih"), body.index("UPDATE Settings SET LastInvoiceHash = '"))
        self.assertIn("rs!EInvoiceXml = xml", body)                                                     # a memo, not SQL
        self.assertIn("ws.CommitTrans", body)

    def test_signing_and_setup(self):
        sign = re.search(r"Public Function ZatcaSignHash.*?End Function", VBA, re.S).group(0)
        self.assertIn('"dgst -sha256 -sign "', sign)
        self.assertIn("Base64Decode(HashB64, b)", sign)                                    # the 32 hash bytes
        test_cert = re.search(r"Public Sub ZatcaSetupTestCertificate.*?End Sub", VBA, re.S).group(0)
        self.assertIn('<> "TEST" Then', test_cert)
        save = re.search(r"Public Sub ZatcaSetupSave.*?End Sub", VBA, re.S).group(0)
        self.assertIn('HasPermission("SETTINGS")', save)
        self.assertLess(save.index("AuditSnapshot("), save.index("rs.Update"))
        self.assertIn('AuditEdited "EDIT", "Settings"', save)

    def test_integer_division_is_not_mixed_with_multiplication(self):
        """VBA binds * tighter than \\: Len(s) \\ 4 * 3 is Len(s) \\ 12 (a base64 buffer far too small)."""
        for name in G.STATIC_MODULES:
            for ln in read(name).splitlines():
                code = ln.split("'")[0]
                with self.subTest(module=name, line=ln.strip()[:60]):
                    self.assertIsNone(re.search(r"\\\s*[\w.$]+\s*\*", code))

    def test_registered(self):
        self.assertIn('"TestZatcaXml"', read("modTestAll"))
        self.assertIn("modZatcaXml", G.STATIC_MODULES)
        self.assertIn(("frmZatcaSetup", "إعداد ربط منصة فاتورة", "النظام", "SETTINGS", False, True, False),
                      schema.SCREEN_LIST)


class StaticZatcaXml(VbaModuleChecks, unittest.TestCase):
    module_name = "modZatcaXml"
    vba = VBA


class StaticZatcaData(VbaModuleChecks, unittest.TestCase):
    module_name = "modZatcaData"
    vba = gen_zatca.build_zatca_data_vba()


if __name__ == "__main__":
    unittest.main()
