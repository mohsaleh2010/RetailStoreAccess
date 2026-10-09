"""Phase F of the e-invoicing plan (docs/50-EInvoice-Reports-VAT-EG.md):
  * the report of the e-documents (EInvoiceReportQuery, rptEInvoiceDocs): every status in its group, a return
    negative, opened from the report centre and from frmEInvoices with the filter of the screen;
  * the warning of the VAT return when documents of the period are not accepted (EInvoiceOpenCount);
  * the Egyptian VAT return (form 10): the same figures in 14 lines (VatReturnQueryEG, rptVatReturnEG), the
    captions and keys of modVat equal to tools/queries.py, the report chosen by the operating country."""
import os
import re
import unittest

from helpers import ROOT
from access_sqlite import AccessOnSqlite
from sim import Store
from access_sqlite import day
import forms as F
import queries as Q
import reports as RP
from reports_catalog import LIST_SPECS
from schema import EINVOICE_STATUSES


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


VAT = read("modVat")
EINV = read("modEInvoice")
GROUPS = {"REJECTED": 1, "INVALID": 1, "PENDING": 2, "SUBMITTED": 2, "REPORTED": 3, "CLEARED": 3, "WARNING": 3,
          "VALID": 3, "CANCELLED": 4, "NOT_SENT": 5}


class EDocumentsReportTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()

    def test_every_status_has_its_group(self):
        self.assertEqual(set(GROUPS), set(EINVOICE_STATUSES))
        sale = self.db.con.execute("SELECT Min(SalesInvoiceID) FROM SalesInvoices").fetchone()[0]
        for status, group in GROUPS.items():
            self.db.con.execute("UPDATE SalesInvoices SET ZatcaStatus = ? WHERE SalesInvoiceID = ?", (status, sale))
            got = self.db.con.execute("SELECT GroupNo FROM EInvoiceReportQuery WHERE DocKind = 'SALE' AND DocID = ?",
                                      (sale,)).fetchone()[0]
            self.assertEqual(got, group, status)

    def test_a_return_is_negative(self):
        rows = self.db.con.execute("SELECT DocKind, TotalAmount, SignedTotal FROM EInvoiceReportQuery").fetchall()
        self.assertIn("RETURN", {r[0] for r in rows})
        for kind, total, signed in rows:
            self.assertAlmostEqual(signed, -total if kind == "RETURN" else total, places=4)

    def test_report_centre_and_screen(self):
        entry = next(r for r in F.REPORTS if r.key == "EINVOICE_DOCS")
        self.assertEqual((entry.query, entry.report, entry.needs, entry.date_column),
                         ("EInvoiceReportQuery", "rptEInvoiceDocs", "D", "DocDate"))
        spec = next(s for s in LIST_SPECS if s.key == "EINVOICE_DOCS")
        self.assertEqual(spec.sorts[0], ("GroupNo", False))
        fields = {c.source for c in spec.cols} | {e for _, e in spec.summary}
        self.assertTrue(any("[GroupNo]=3" in f for f in fields))
        self.assertIn("rptEInvoiceDocs", {r.name for r in RP.all_reports()})
        model = next(m for m in F.all_forms() if m.name == "frmEInvoices")
        btn = next(c for c in model.controls if c.name == "btnPrint")
        self.assertIn("EInvoicesPrint Me", "\n".join(model.code) + str(btn.props))

    def test_screen_and_report_use_one_filter(self):
        self.assertIn("EInvoiceStatusWhere(", proc(EINV, "EInvoicesShow"))
        body = proc(EINV, "EInvoicesPrint")
        for part in ("EInvoiceStatusWhere(", '"rptEInvoiceDocs", "EInvoiceReportQuery"', "[DocDate] <"):
            self.assertIn(part, body)
        for choice in ("ATTENTION", "SENT", "NOT_SENT"):
            self.assertIn(f'Case "{choice}"', proc(EINV, "EInvoiceStatusWhere"))

    def test_open_documents_of_the_period(self):
        body = proc(EINV, "EInvoiceOpenCount")
        self.assertIn("('PENDING','SUBMITTED','REJECTED','INVALID')", body)
        self.assertIn('DateAdd("d", 1, ToDate)', body)
        note = proc(VAT, "VatEInvoiceNote")
        self.assertIn("EInvoiceOpenCount(", note)
        self.assertIn("VatEInvoiceNote(", proc(VAT, "VatCalculate"))
        self.assertIn("VatEInvoiceNote(", proc(VAT, "VatShowReturn"))


class EgyptVatReturnTests(unittest.TestCase):

    def test_layouts(self):
        self.assertEqual([b for b, *_ in Q.VAT_BOXES_EG], list(range(1, 15)))
        for layout in (Q.VAT_LAYOUT_SA, Q.VAT_LAYOUT_EG):
            keys = [k for k, _ in layout]
            for key in ("SS", "SZ", "SE", "ST", "PS", "PZ", "PT", "DUE", "COR", "CAR", "NET"):
                self.assertEqual(keys.count(key), 1, key)            # every recorded figure once
            self.assertEqual(keys[-1], "NET")
        self.assertNotIn("15%", "".join(c for _, c in Q.VAT_LAYOUT_EG))

    def test_vba_keys_and_captions_match(self):
        for country, layout in (("SA", Q.VAT_LAYOUT_SA), ("EG", Q.VAT_LAYOUT_EG)):
            const = re.search(rf'Private Const VAT_KEYS_{country} As String = "(.*)"', VAT).group(1)
            self.assertEqual(const.split(","), [k for k, _ in layout], country)
        eg = dict((int(n), c) for n, c in re.findall(r'Case (\d+): VatBoxCaptionEG = "(.*)"',
                                                      proc(VAT, "VatBoxCaptionEG")))
        self.assertEqual(eg, {b: c for b, c, *_ in Q.VAT_BOXES_EG})
        rows = proc(VAT, "VatBoxRows")
        for key in Q.VAT_FIGURES:
            if key != "0":
                self.assertIn(f'Case "{key}"', rows)
        self.assertIn('If AppCountry() = "EG" Then', proc(VAT, "VatBoxCaption"))

    def test_same_figures_as_the_saudi_return(self):
        db = AccessOnSqlite()
        db.load_fixture()
        s = Store(db.con, now=day(0, 9))
        s.sync_journal()
        db.params.update({"PeriodStart": "2000-01-01 00:00:00", "PeriodEnd": "2100-01-01 00:00:00"})
        cur = db.con.execute("SELECT * FROM qryVatReturnTotals")
        t = dict(zip([d[0] for d in cur.description], cur.fetchone()))
        vid, net = s.file_vat_return(t, "2000-01-01 00:00:00", day(1), day(0), corrections=7, carried=3)
        db.params["VatReturnID"] = vid
        sa = {r[0]: r[1:] for r in db.con.execute("SELECT BoxNo, Amount, Adjust, VAT FROM VatReturnQuery")}
        eg = {r[0]: r[1:] for r in db.con.execute("SELECT BoxNo, Amount, Adjust, VAT FROM VatReturnQueryEG")}
        self.assertEqual(sorted(eg), list(range(1, 15)))
        sa_key = {k: i for i, (k, _) in enumerate(Q.VAT_LAYOUT_SA, 1) if k != "0"}
        for i, (key, _) in enumerate(Q.VAT_LAYOUT_EG, 1):
            if key == "0":
                self.assertEqual(eg[i], (0, 0, 0))
            else:
                self.assertEqual(eg[i], sa[sa_key[key]], key)
        self.assertAlmostEqual(eg[14][2], float(net), places=4)

    def test_report_and_screen_follow_the_country(self):
        reports = {r.name: r for r in RP.all_reports()}
        self.assertEqual(reports["rptVatReturnEG"].record_source, "VatReturnQueryEG")
        sources = [c.source or "" for cs in reports["rptVatReturnEG"].controls.values() for c in cs]
        sources += [v for cs in reports["rptVatReturnEG"].controls.values() for c in cs for v in c.props.values()
                    if isinstance(v, str)] + [reports["rptVatReturnEG"].caption]
        self.assertTrue(any("[BoxNo]=14" in x for x in sources))
        self.assertTrue(any("المصلحة" in x for x in sources))
        self.assertTrue(any("نموذج 10" in x for x in sources))
        printing = proc(VAT, "PrintVatReturn")
        self.assertIn('OpenReportOrQuery "rptVatReturnEG", "VatReturnQueryEG"', printing)
        load = proc(VAT, "VatReturnLoad")
        self.assertIn("frm!btnLastQuarter.Visible = False", load)
        self.assertIn("frm!lblFilingRef.Caption", load)
        names = {c.name for m in F.all_forms() if m.name == "frmVatReturn" for c in m.controls}
        self.assertTrue({"btnLastQuarter", "lblFilingRef"} <= names)

    def test_in_access_test_is_run(self):
        self.assertIn('"TestEInvoiceReports"', read("modTestAll"))
        body = proc(EINV, "TestEInvoiceReports")
        for part in ("ws.Rollback", "CountryCode = 'EG'", "VatBoxCount() = 14", "VatBoxCount() = 16",
                     "EInvoiceOpenCount("):
            self.assertIn(part, body)


if __name__ == "__main__":
    unittest.main()
