"""E-invoicing in C# and the office service (plan 7A, docs/85-EInvoice-CSharp.md).

The C# (RetailStore.Domain/EInvoicing, RetailStore.Data/Services/EInvoiceService.cs, OfficeSyncService.cs) follows
modEInvoice, modZatcaXml, modZatcaApi, modEtaReceipt, modEtaInvoice and modOfficeSync: the same messages, the same order of
the steps, the same chain lock. Its outputs are compared with the Python mirrors by xUnit (EInvoiceVectors.g.cs, written by
tools/gen_web_einvoice.py); here the code itself is checked."""
import json
import os
import re
import unittest

from helpers import ROOT
import gen_web_einvoice as GE
import zatca_xml_reference as ZX

WEB = "web/src/"


def read(rel):
    with open(os.path.join(ROOT, rel), encoding="utf-8-sig") as fh:
        return fh.read()


def vba(*names):
    return "".join(read(f"src/vba/{n}.bas") for n in names)


def arabic_literals(code):
    return {m.replace("\\r\\n", "").replace("\\n", "") for m in re.findall(r'"([^"]*[؀-ۿ][^"]*)"', code)}


SERVICE = read(WEB + "RetailStore.Data/Services/EInvoiceService.cs")
OFFICE = read(WEB + "RetailStore.Data/Services/OfficeSyncService.cs")
DOMAIN = read(WEB + "RetailStore.Domain/EInvoicing/ZatcaDocument.cs") + read(WEB + "RetailStore.Domain/EInvoicing/EtaDocument.cs") + \
    read(WEB + "RetailStore.Domain/EInvoicing/EInvoiceReply.cs")
VBA = vba("modEInvoice", "modZatcaXml", "modZatcaApi", "modEtaReceipt", "modEtaInvoice", "modOfficeSync")


def method(code, name):
    start = code.index(name)
    nxt = re.search(r"\n    /// <summary>|\n    (public|private) ", code[start + len(name):])
    return code[start:start + len(name) + (nxt.start() if nxt else len(code))]


class MessagesTests(unittest.TestCase):

    def test_the_messages_are_those_of_access(self):
        messages = arabic_literals(SERVICE + OFFICE + DOMAIN)
        self.assertGreater(len(messages), 60)
        for m in sorted(messages):
            with self.subTest(m):
                self.assertIn(m, VBA)


class OrderTests(unittest.TestCase):

    def assertInOrder(self, text, order):
        self.assertEqual([text.index(x) for x in order], sorted(text.index(x) for x in order), order)

    def test_the_zatca_chain_is_moved_inside_one_transaction(self):
        body = method(SERVICE, "public async Task<string> ZatcaPrepareAsync")
        self.assertInOrder(body, ["!string.IsNullOrEmpty(h.Xml)", "ZatcaDataProblemAsync", "BeginTransactionAsync",
                                  'LockChainAsync(db, "LastInvoiceHash")', "x.LastInvoiceHash", "ZatcaDocumentXmlAsync(db,",
                                  "ZatcaDocument.Sign(", "SaveDocAsync", "SetProperty(x => x.LastInvoiceHash, hash)", "CommitAsync"])

    def test_the_eta_chain_is_moved_inside_one_transaction(self):
        body = method(SERVICE, "public async Task<string> EtaReceiptPrepareAsync")
        self.assertInOrder(body, ['h.Status == "REJECTED"', "BeginTransactionAsync", 'LockChainAsync(db, "EtaLastUUID")',
                                  "EtaReceiptJsonAsync(db,", "EtaDocument.ReceiptUuid", "EtaDocument.WithUuid", "SaveDocAsync",
                                  "x.EtaLastUUID == oldUuid", "CommitAsync"])
        self.assertIn("UPDATE Settings SET EtaLastUUID = EtaLastUUID WHERE SettingID = 1", SERVICE)
        self.assertIn("UPDATE Settings SET LastInvoiceHash = LastInvoiceHash WHERE SettingID = 1", SERVICE)
        self.assertIn('UPDATE Settings SET EtaLastUUID = EtaLastUUID WHERE SettingID = 1"', vba("modEtaReceipt").replace(
            '", dbFailOnError', '"'))

    def test_sending_follows_send_einvoice(self):
        body = method(SERVICE, "public async Task<EInvoiceReply> SendAsync")
        self.assertInOrder(body, ["s.EInvoiceEnabled", "ReadyProblem", "EtaSendAsync", "EInvoiceAttempts", "EInvoiceError"])
        pending = method(SERVICE, "public async Task<(int Sent, int Failed, string Problem)> SendPendingAsync")
        self.assertIn('r.Result is "NETWORK" or "ERROR"', pending)
        self.assertIn("ThenByDescending(d => d.Kind", SERVICE, "qryEInvoiceDocs ORDER BY DocDate, DocKind DESC, DocID")

    def test_the_office_run_follows_run_office_sync(self):
        body = method(OFFICE, "public async Task<OfficeSyncRun> RunAsync")
        self.assertInOrder(body, ["journal.SyncAsync", "EInvoiceEnabled", "ReadyProblemAsync", "PendingCountAsync",
                                  "SendPendingAsync", "RefreshSubmittedAsync", "WriteLogAsync"])
        run = method(vba("modOfficeSync"), "Public Function RunOfficeSync")
        self.assertInOrder(run, ["SyncJournal", "EInvoiceEnabled", "EInvoiceReadyProblem", "PendingEInvoiceCount",
                                 "SendPendingEInvoices", "RefreshSubmittedEInvoices", "WriteOfficeSyncLog"])
        for part in ("القيود: ", " جديد، ", " معدّل، ", " محذوف. ", "الفواتير الإلكترونية: أُرسلت ", "، لم تُرسل ", "، تمت متابعة "):
            self.assertIn(part, OFFICE)
        self.assertIn("OfficeSyncRules.IsDue(", OFFICE)


class SafetyTests(unittest.TestCase):

    def test_secrets_are_not_logged(self):
        token = method(SERVICE, "private async Task<(string Token, EInvoiceReply? Problem)> TokenAsync")
        self.assertIn('erp ? "TOKEN_ERP" : "TOKEN", url, status, result, msg, "", masked', token)
        self.assertIn('reply.Replace(token, "***")', token)

    def test_the_office_service(self):
        project = read(WEB + "RetailStore.OfficeService/RetailStore.OfficeService.csproj")
        self.assertIn('Sdk="Microsoft.NET.Sdk.Worker"', project)
        self.assertIn("Microsoft.Extensions.Hosting.WindowsServices", project)
        self.assertIn("AddWindowsService", read(WEB + "RetailStore.OfficeService/Program.cs"))
        settings = json.loads(read(WEB + "RetailStore.OfficeService/appsettings.json"))
        self.assertEqual(settings["ConnectionStrings"], {"SqlServer": "", "Sqlite": ""}, "never commit a connection string")
        self.assertIn("RetailStore.OfficeService.csproj", read("web/RetailStore.slnx"))
        worker = read(WEB + "RetailStore.OfficeService/OfficeWorker.cs")
        self.assertIn('office.RunAsync("AUTO", name, user)', worker)
        self.assertIn('"SERVICE-" + machine', OFFICE, "the name never equals a computer name: Access does not run a second time")

    def test_signing_is_done_in_dotnet_as_openssl_does(self):
        self.assertIn("ec.SignData(Convert.FromBase64String(invoiceHashB64), HashAlgorithmName.SHA256", DOMAIN)
        self.assertIn("DSASignatureFormat.Rfc3279DerSequence", DOMAIN)
        self.assertIn('dgst -sha256 -sign', vba("modZatcaXml"))

    def test_the_templates_are_those_of_the_mirror(self):
        """The C# builds the XAdES texts from pieces; the vectors compare them, and here the template texts too."""
        for piece in ZX.UBL_EXTENSIONS.split("\n"):
            with self.subTest(piece):
                self.assertIn(json.dumps(piece, ensure_ascii=False), DOMAIN)

    def test_the_test_key_is_only_a_test_key(self):
        self.assertIn("TST-EGS-RetailStore", __import__("base64").b64decode(GE.TEST_CERT).decode("latin-1"))
        self.assertNotIn("PRIVATE KEY", SERVICE + OFFICE + DOMAIN)


if __name__ == "__main__":
    unittest.main()
