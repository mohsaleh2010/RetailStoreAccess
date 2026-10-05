"""Phase 8 checks: every formatted report (catalogue, documents, sales invoices).

  * layout: inside the page and the sections, no overlapping controls,
    portrait <= 19 cm / landscape <= 27.4 cm;
  * every field, sort field and function used by a report exists (fields are
    taken from the saved query itself, run on the SQLite mirror);
  * every report-centre entry has its report, on the same query;
  * running balances in the statements end on the balance of the account;
  * the amount-in-words function is EXECUTED with LibreOffice and compared
    with tools/tafqeet.py.

Run:  python3 -m unittest discover -s tests -v
"""

import os
import random
import re
import unittest

from access_sqlite import AccessOnSqlite
from helpers import ROOT, VbaModuleChecks
import forms as F
import reports as RP
import reports_catalog as RC
import tafqeet as T
import vba_harness as H

MODELS = RP.all_reports()
ACCESS_FUNCTIONS = {"Nz", "IIf", "Format", "Sum", "Count", "Len", "IsNull", "Trim"}
PROJECT_FUNCTIONS = {"SettingValue", "GDate", "ReportCriteria", "ReportPrintedAt", "AmountInWords",
                     "LabelCode", "LabelPrice", "OrderTypeText"}
REPORT_PROPERTIES = {"Page", "Pages"}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def overlaps(a, b):
    return a.x < b.x + b.w and b.x < a.x + a.w and a.y < b.y + b.h and b.y < a.y + a.h


def all_controls(m):
    for sec, ctls in m.controls.items():
        for c in ctls:
            yield sec, c


class LayoutTests(unittest.TestCase):

    def test_names_unique(self):
        names = [m.name for m in MODELS]
        self.assertEqual(len(names), len(set(names)))
        for m in MODELS:
            ctl = [c.name for _, c in all_controls(m)]
            self.assertEqual(len(ctl), len(set(ctl)), m.name)

    def test_inside_page_and_sections(self):
        for m in MODELS:
            if m.page_setup:
                self.assertLessEqual(m.width, F.cm(RC.LANDSCAPE_W if m.landscape else RC.PORTRAIT_W), m.name)
            for sec, c in all_controls(m):
                with self.subTest(report=m.name, control=c.name):
                    self.assertIn(sec, m.heights)
                    self.assertGreaterEqual(c.x, 0)
                    self.assertGreater(c.w, 0)
                    self.assertLessEqual(c.x + c.w, m.width)
                    self.assertLessEqual(c.y + c.h, m.heights[sec])

    def test_no_overlapping_controls(self):
        for m in MODELS:
            for sec, ctls in m.controls.items():
                solid = [c for c in ctls if not c.decorative and c.props.get("Visible", True)]
                for i, a in enumerate(solid):
                    for b in solid[i + 1:]:
                        self.assertFalse(overlaps(a, b), f"{m.name}: {a.name} / {b.name}")

    def test_labels_have_text(self):
        for m in MODELS:
            for _, c in all_controls(m):
                if c.kind == "label":
                    self.assertTrue(c.props["Caption"].strip(), f"{m.name}.{c.name}")


# unbound reports that draw their own content (modCharts); the report-centre query is only the
# fallback shown when the report is missing
DRAWN = {"rptStatistics"}


class CatalogueTests(unittest.TestCase):

    def test_every_report_centre_entry_has_its_report(self):
        by_name = {m.name: m for m in MODELS}
        for r in F.REPORTS:
            with self.subTest(r.key):
                self.assertIn(r.report, by_name)
                if r.report in DRAWN:
                    self.assertEqual(by_name[r.report].record_source, "")
                else:
                    self.assertEqual(by_name[r.report].record_source, r.query)
        self.assertEqual({s.key for s in RC.LIST_SPECS} | {s.key for s in RC.CARD_SPECS},
                         {r.key for r in F.REPORTS if r.report not in DRAWN})

    def test_low_stock_report_exists(self):
        """Required by the specification: "Low Stock Products Report"."""
        m = next(m for m in MODELS if m.name == "rptLowStock")
        self.assertEqual(m.record_source, "LowStockQuery")
        self.assertEqual(m.sorts[0], ("ShortageQty", True))

    def test_every_list_and_document_report_handles_no_data(self):
        for m in MODELS:
            if m.page_setup and m.name not in DRAWN:
                self.assertIn("m_rpt.OnNoData = EP", m.events, m.name)
                self.assertTrue(any("ReportNoData" in line for line in m.code), m.name)

    def test_page_footer_and_criteria_on_catalogue_reports(self):
        for m in RC.catalog_reports():
            sources = [c.source for _, c in all_controls(m) if c.kind == "text"]
            self.assertIn("=ReportCriteria()", sources, m.name)
            self.assertIn("=ReportPrintedAt()", sources, m.name)
            self.assertTrue(any("[Pages]" in s for s in sources), m.name)


class SourceTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.db = db = AccessOnSqlite()
        db.load_fixture()
        db.params.update({"CustomerID": db.ids["C2"], "SupplierID": db.ids["S1"], "ProductID": db.ids["P1"]})
        cls.columns = {}
        # local front-end tables of the label report (modLabels.EnsureLabelTables)
        db.con.execute("CREATE TABLE tmpLabelQueue (LineNo INTEGER PRIMARY KEY, ProductID INTEGER, "
                       "ProductName TEXT, LabelCode TEXT, Price NUMERIC, Copies INTEGER)")
        db.con.execute("CREATE TABLE tmpLabelNumbers (N INTEGER PRIMARY KEY)")
        for m in MODELS:
            if m.record_source not in cls.columns:
                src = m.record_source
                if not src:                         # unbound (rptStatistics draws its own content)
                    cls.columns[src] = set()
                    continue
                sql = f"SELECT * FROM ({src})" if src.upper().startswith("SELECT ") else f'SELECT * FROM "{src}"'
                cur = db.con.execute(sql)
                cls.columns[m.record_source] = {d[0] for d in cur.description}

    def test_fields_exist_in_record_source(self):
        for m in MODELS:
            cols = self.columns[m.record_source]
            for _, c in all_controls(m):
                if c.kind != "text":
                    continue
                refs = re.findall(r"\[(\w+)\]", c.source) if c.source.startswith("=") else [c.source]
                for ref in refs:
                    if ref not in REPORT_PROPERTIES:
                        self.assertIn(ref, cols, f"{m.name}.{c.name}")
            for f, _ in m.sorts:
                self.assertIn(f, cols, f"{m.name}: sort {f}")
            if m.group:
                self.assertIn(m.group, cols, f"{m.name}: group {m.group}")

    def test_functions_exist(self):
        public = set()
        for name in ("modCommon", "modReports", "modLabels", "modTouchPOS"):
            public |= set(re.findall(r"^Public Function (\w+)\(", read(name), re.M))
        self.assertTrue(PROJECT_FUNCTIONS <= public, PROJECT_FUNCTIONS - public)
        for m in MODELS:
            for _, c in all_controls(m):
                if c.kind == "text" and c.source.startswith("="):
                    code = re.sub(r'"[^"]*"', '""', c.source)
                    for fn in re.findall(r"\b([A-Za-z]\w*)\(", code):
                        self.assertIn(fn, ACCESS_FUNCTIONS | PROJECT_FUNCTIONS, f"{m.name}.{c.name}: {fn}")

    def test_card_reports_read_one_row(self):
        for m in RC.catalog_reports():
            if m.record_source in ("ProfitQuery", "VatSummaryQuery"):
                rows = self.db.con.execute(f'SELECT COUNT(*) FROM "{m.record_source}"').fetchone()[0]
                self.assertEqual(rows, 1, m.name)

    def test_statement_running_balance_ends_on_the_account_balance(self):
        """rptCustomerStatement / rptSupplierStatement add Debit-Credit (Credit-Debit) line by line."""
        db = self.db
        db.set_period(400)                     # the whole fixture history
        last = db.con.execute("SELECT Sum(Debit - Credit) FROM CustomerStatementQuery").fetchone()[0]
        bal = db.con.execute("SELECT Balance FROM CustomerBalanceQuery WHERE CustomerID = ?",
                             (db.ids["C2"],)).fetchone()[0]
        self.assertAlmostEqual(last, bal, places=2)
        last = db.con.execute("SELECT Sum(Credit - Debit) FROM SupplierStatementQuery").fetchone()[0]
        bal = db.con.execute("SELECT Balance FROM SupplierBalanceQuery WHERE SupplierID = ?",
                             (db.ids["S1"],)).fetchone()[0]
        self.assertAlmostEqual(last, bal, places=2)
        spec = {s.key: s for s in RC.LIST_SPECS}
        self.assertEqual(spec["CUSTOMER_STATEMENT"].cols[-1].source, "=[Debit]-[Credit]")
        self.assertEqual(spec["SUPPLIER_STATEMENT"].cols[-1].source, "=[Credit]-[Debit]")
        self.assertTrue(spec["CUSTOMER_STATEMENT"].cols[-1].running)


class AmountInWordsReference(unittest.TestCase):

    def test_known_texts(self):
        self.assertEqual(T.amount_in_words(1250.5), "فقط ألف ومائتان وخمسون ريال سعودي وخمسون هللة لا غير")
        self.assertEqual(T.amount_in_words(21), "فقط واحد وعشرون ريال سعودي لا غير")
        self.assertEqual(T.amount_in_words(2000), "فقط ألفان ريال سعودي لا غير")
        self.assertEqual(T.amount_in_words(3175), "فقط ثلاثة آلاف ومائة وخمسة وسبعون ريال سعودي لا غير")
        self.assertEqual(T.amount_in_words(11000), "فقط أحد عشر ألف ريال سعودي لا غير")
        self.assertEqual(T.amount_in_words(0.05), "فقط خمسة هللة لا غير")
        self.assertEqual(T.amount_in_words(3500000), "فقط ثلاثة ملايين وخمسمائة ألف ريال سعودي لا غير")


DRIVER = r'''
Option VBASupport 1
Public Function RunWords(ByVal v As Double) As String
    RunWords = AmountInWords(CCur(v))
End Function
'''


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class AmountInWordsRuntime(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modCommon": H.read_module("modCommon"), "modReports": H.read_module("modReports"),
                    "Driver": DRIVER})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_cases_and_random_amounts(self):
        rnd = random.Random(8)
        amounts = list(T.CASES) + [round(rnd.uniform(0, 999_999_999), 2) for _ in range(60)] + \
            [rnd.randint(1, 120) for _ in range(40)]
        for a in amounts:
            with self.subTest(a):
                self.assertEqual(self.h.call("Driver", "RunWords", float(a)), T.amount_in_words(a))


class StaticModReports(VbaModuleChecks, unittest.TestCase):
    module_name = "modReports"
    vba = read("modReports")
