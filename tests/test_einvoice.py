"""The e-invoicing foundation (modEInvoice, modHttp, docs/45-EInvoice-Foundation.md): the status of every sales
document, the log of every request, the adapter of each country called by Application.Run, the follow-up screen,
and the JSON reader of the replies (tools/json_reference.py is its Python mirror)."""
import json
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite
import forms as F
import generate as G
import json_reference as J
import queries as Q
import vba_harness as H
from schema import EINVOICE_STATUSES, RULE_UPGRADES, SCREEN_LIST, table


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


EINV = read("modEInvoice")
HTTP = read("modHttp")

SAMPLE = json.dumps({
    "validationResults": {"status": "WARNING", "warningMessages": [{"code": "BR-KSA-08", "message": "تنبيه \"مهم\""}],
                          "errorMessages": []},
    "reportingStatus": "REPORTED", "clearedInvoice": None, "amount": 12.5, "count": 3, "ok": True,
    "list": [[1, 2], {"a": "\\/\n\t"}, "x"], "unicode": "é ☃ 😀", "empty": {},
}, ensure_ascii=False)
SAMPLE_ASCII = json.dumps(json.loads(SAMPLE), ensure_ascii=True, indent=2)
PATHS = ["validationResults.status", "validationResults.warningMessages[0].code",
         "validationResults.warningMessages[0].message", "validationResults.errorMessages", "reportingStatus",
         "clearedInvoice", "amount", "count", "ok", "list[0]", "list[0][1]", "list[1].a", "list[2]", "list[3]",
         "unicode", "empty", "missing", "validationResults.missing", "", "list[1].b"]


def expected(text, path):
    """json_get by the json module: the value at path, numbers and booleans as written, objects as JSON."""
    value = json.loads(text)
    segs = path.replace("[", ".#").replace("]", "").replace("..", ".").strip(".")
    for seg in (segs.split(".") if segs else []):
        if seg.startswith("#"):
            i = int(seg[1:])
            if not isinstance(value, list) or i >= len(value):
                return "MISSING"
            value = value[i]
        else:
            if not isinstance(value, dict) or seg not in value:
                return "MISSING"
            value = value[seg]
    return value


class JsonReferenceTests(unittest.TestCase):

    def test_same_as_the_json_module(self):
        for text in (SAMPLE, SAMPLE_ASCII):
            for path in PATHS:
                with self.subTest(path=path):
                    got, want = J.json_get(text, path), expected(text, path)
                    if want == "MISSING" or want is None:
                        self.assertIsNone(got)
                    elif isinstance(want, str):
                        self.assertEqual(got, want)
                    else:
                        self.assertEqual(json.loads(got), want)

    def test_escape(self):
        for s in ['a"b\\c', "line\nnext\r\ttab", "\x01\x1f", "عربي 😀"]:
            self.assertEqual(json.loads('"' + J.json_escape(s) + '"'), s)


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class JsonRuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modHttp": HTTP, "modEInvoice": EINV, "Driver": (
            "Option VBASupport 1\n"
            "Public Function RunGet(ByVal Json As String, ByVal Path As String) As String\n"
            "    Dim v As Variant\n    v = JsonGet(Json, Path)\n"
            "    If IsNull(v) Then\n        RunGet = \"<NULL>\"\n    Else\n        RunGet = v\n    End If\n"
            "End Function\n"
            "Public Function RunSplit(ByVal Reply As String) As String\n"
            "    Dim m As String, c As String\n    c = SplitResult(Reply, m)\n    RunSplit = c & \"/\" & m\n"
            "End Function\n")})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_same_as_the_python_mirror(self):
        for text in (SAMPLE, SAMPLE_ASCII):
            for path in PATHS:
                with self.subTest(path=path):
                    want = J.json_get(text, path)
                    self.assertEqual(self.h.call("Driver", "RunGet", text, path), "<NULL>" if want is None else want)

    def test_escape(self):
        for s in ['a"b\\c', "line\nnext\r\ttab", "\x01\x1f", "عربي"]:
            self.assertEqual(self.h.call("modHttp", "JsonEscape", s), J.json_escape(s))

    def test_split_result(self):
        self.assertEqual(self.h.call("Driver", "RunSplit", "REJECTED|السبب"), "REJECTED/السبب")
        self.assertEqual(self.h.call("Driver", "RunSplit", "ok|"), "OK/")
        self.assertEqual(self.h.call("Driver", "RunSplit", "strange"), "ERROR/strange")


class SchemaTests(unittest.TestCase):

    def test_status_and_attempts_on_both_documents(self):
        for tbl in ("SalesInvoices", "SalesReturns"):
            fields = {f.name: f for f in table(tbl).fields}
            self.assertEqual(fields["ZatcaStatus"].rule, "In (" + ",".join(f'"{s}"' for s in EINVOICE_STATUSES) + ")")
            self.assertIn((tbl, "ZatcaStatus"), RULE_UPGRADES)            # an existing back-end gets the new values
            self.assertFalse(fields["EInvoiceAttempts"].required)
            self.assertFalse(fields["EInvoiceError"].required)
        self.assertTrue({"NOT_SENT", "PENDING", "REPORTED", "CLEARED", "SUBMITTED", "VALID", "INVALID"}
                        <= set(EINVOICE_STATUSES))

    def test_log_and_settings(self):
        log = {f.name for f in table("EInvoiceLog").fields}
        self.assertTrue({"LoggedAt", "Country", "DocKind", "DocID", "Action", "HttpStatus", "Result", "Message",
                         "RequestBody", "ResponseBody"} <= log)
        settings = {f.name: f for f in table("Settings").fields}
        self.assertEqual(settings["EInvoiceEnabled"].default, "False")
        self.assertEqual(settings["EInvoiceEnvironment"].rule, 'In ("TEST","SIMULATION","PRODUCTION")')

    def test_permission_and_screen(self):
        self.assertIn("EINVOICE", {r[0] for r in table("Permissions").seed_rows})
        self.assertIn(("frmEInvoices", "الفاتورة الإلكترونية", "المبيعات", "EINVOICE", False, True, False), SCREEN_LIST)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmEInvoices"], "EINVOICE")
        settings = next(m for m in F.all_forms() if m.name == "frmSettings")
        self.assertTrue(any(c.name == "btnEInvoices" for c in settings.controls))


class QueryTests(unittest.TestCase):

    def test_documents_query_runs(self):
        db = AccessOnSqlite()
        db.load_fixture()
        rows = db.con.execute("SELECT DocKind, EStatus, StatusName FROM qryEInvoiceDocs").fetchall()
        self.assertTrue(rows)
        self.assertTrue({r[0] for r in rows} <= {"SALE", "RETURN"})
        db.con.execute("UPDATE SalesInvoices SET ZatcaStatus = 'CLEARED'")
        names = {r[0] for r in db.con.execute("SELECT StatusName FROM qryEInvoiceDocs WHERE DocKind = 'SALE'")}
        self.assertEqual(names, {"معتمد"})
        sql = next(q for q in Q.QUERIES if q.name == "qryEInvoiceDocs").sql
        self.assertNotIn("&", sql)                                      # no & in saved queries
        for status in EINVOICE_STATUSES[1:]:
            self.assertIn(f"'{status}'", sql)


class CodeTests(unittest.TestCase):

    def test_adapters_are_called_by_name(self):
        """Each country is a module of its own, added later without changing this one."""
        self.assertIn('AdapterProc = "Eta" & Name', EINV)
        self.assertIn('AdapterProc = "Zatca" & Name', EINV)
        self.assertIn('Application.Run(AdapterProc("SendDocument"), DocKind, DocID)', EINV)
        self.assertIn('Application.Run(AdapterProc("ReadyProblem"))', EINV)
        send = re.search(r"Public Function SendEInvoice.*?End Function", EINV, re.S).group(0)
        self.assertLess(send.index("EInvoiceEnabled()"), send.index("Application.Run"))
        self.assertLess(send.index("EInvoiceReadyProblem()"), send.index("Application.Run"))
        self.assertIn("EInvoiceAttempts = Nz(EInvoiceAttempts, 0) + 1", send)

    def test_sending_after_a_sale_never_fails_it(self):
        sales = read("modSales")
        for kind, after in (("SALE", 'LogAction "SALE"'), ("RETURN", 'LogAction "SALES_RETURN"')):
            self.assertLess(sales.index(after), sales.index(f'EInvoiceAfterSave "{kind}"'))
        hook = re.search(r"Public Sub EInvoiceAfterSave.*?End Sub", EINV, re.S).group(0)
        self.assertIn("On Error GoTo Done", hook)
        self.assertLess(hook.index("If Not EInvoiceEnabled() Then Exit Sub"), hook.index("SendEInvoice"))

    def test_new_documents_wait_when_enabled(self):
        self.assertIn('If Nz(SettingValue("EInvoiceEnabled"), False) Or', read("modZatca"))

    def test_enabling_needs_permission_and_a_ready_platform(self):
        body = re.search(r"Public Sub EInvoiceSettingChanged.*?End Sub", EINV, re.S).group(0)
        self.assertIn('HasPermission("SETTINGS")', body)
        self.assertIn("EInvoiceReadyProblem()", body)
        self.assertLess(body.index("AuditSnapshot("), body.index("UPDATE Settings"))
        self.assertIn('AuditEdited "EDIT", "Settings"', body)

    def test_screen_actions_need_permission(self):
        for proc in ("EInvoicesSendPicked", "EInvoicesSendAll"):
            body = re.search(rf"Public Sub {proc}.*?End Sub", EINV, re.S).group(0)
            self.assertIn('CanScreenAction(frm.Name, "EDIT")', body)

    def test_http(self):
        self.assertIn('CreateObject("MSXML2.ServerXMLHTTP.6.0")', HTTP)
        body = re.search(r"Public Function HttpSend.*?End Function", HTTP, re.S).group(0)
        self.assertIn("On Error GoTo EH", body)                         # a network problem is a message

    def test_in_access_test_is_run(self):
        self.assertIn('"TestEInvoice"', read("modTestAll"))
        self.assertTrue({"modHttp", "modEInvoice"} <= set(G.STATIC_MODULES))


class StaticEInvoice(VbaModuleChecks, unittest.TestCase):
    module_name = "modEInvoice"
    vba = EINV


class StaticHttp(VbaModuleChecks, unittest.TestCase):
    module_name = "modHttp"
    vba = HTTP


if __name__ == "__main__":
    unittest.main()
