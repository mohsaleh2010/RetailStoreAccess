"""Plan 6F (docs/82): the financial indicators and the report centre of the web.

  * every report of forms.REPORTS reaches the web: a layout of reports_catalog (or WEB_ONLY_SPECS), or a page (LINKS);
  * the Access expressions of the layouts become C# (gen_web_reports), an unknown one is refused;
  * the saved queries reach the SQLite development database as views, and the parameters the web sets cover every
    parameter of the report queries;
  * ReportService follows modScreens.PrepareReport (its checks in their order, its Arabic messages), the indicators follow
    modIndicators (formulas of tools/indicators_reference.py, texts of the cards)."""
import os
import re
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools"))

import forms as F  # noqa: E402
import gen_sql_views as SV  # noqa: E402
import gen_web_reports as G  # noqa: E402
import indicators_reference as IR  # noqa: E402
import queries as Q  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def read(rel):
    with open(os.path.join(ROOT, rel), encoding="utf-8-sig") as fh:
        return fh.read()


def arabic_literals(code):
    code = re.sub(r"//[^\n]*", "", code)
    return {m for m in re.findall(r'"([^"]*[؀-ۿ][^"]*)"', code)}


class CatalogueTests(unittest.TestCase):

    def test_every_report_of_the_centre_reaches_the_web(self):
        lists, cards = G.specs()
        for r in F.REPORTS:
            with self.subTest(r.key):
                self.assertTrue(r.key in lists or r.key in cards or r.key in G.LINKS)
        text = read("web/src/RetailStore.Data/Generated/Reports.g.cs")
        self.assertEqual(text, G.build_reports())
        for r in F.REPORTS:
            self.assertIn(f'new("{r.key}", ', text)

    def test_the_expressions_become_csharp(self):
        self.assertEqual(G.column_lambda("ProductName"), 'r => r.Value("ProductName")')
        self.assertEqual(G.column_lambda("=GDate([EntryDate],True)"), 'r => (object)ReportRow.GDate(r.Value("EntryDate"), true)')
        self.assertEqual(G.column_lambda('=IIf([PaymentType]="CREDIT","آجل","نقدي")'),
                         'r => (object)((r.Text("PaymentType") == "CREDIT") ? "آجل" : "نقدي")')
        self.assertEqual(G.column_lambda("=[Debit]-[Credit]"), 'r => (object)(r.Num("Debit") - r.Num("Credit"))')
        self.assertEqual(G.column_lambda("=IIf([LineDebit]=0,Null,[LineDebit])"),
                         'r => ((r.Num("LineDebit") == 0m) ? null : r.Value("LineDebit"))')
        self.assertEqual(G.sum_lambda("=Sum(IIf([SortKey]=1,[QtyIn],0))"),
                         'rows => rows.Sum(r => ((r.Num("SortKey") == 1m) ? r.Num("QtyIn") : 0m))')
        for bad in ("=DLookup([X],[Y])", "=[A] Like [B]", '=IIf([A]>"x",1,2)'):
            with self.subTest(bad):
                with self.assertRaises(ValueError):
                    G.column_lambda(bad)
        with self.assertRaises(ValueError):
            G.sum_lambda("=Count(*)")

    def test_report_texts_need_english(self):
        strings = G.report_strings()
        self.assertIn("المبيعات اليومية", strings)
        self.assertIn("آجل", strings)                   # a text of an expression

    def test_the_saved_queries_are_sqlite_views(self):
        text = read("web/src/RetailStore.Data/Generated/SavedQueries.g.cs")
        self.assertEqual(text, G.build_saved_queries())
        for q in Q.QUERIES:
            self.assertIn(f'("{q.name}", ', text)
        self.assertNotIn("[@", text, "the names of the master data are resolved")
        runner = read("web/src/RetailStore.Data/Reports/SavedQueryRunner.cs")
        functions = set(re.findall(r'CreateFunction(?:<[^>]*>)?\("(\w+)"', runner))
        harness = set(re.findall(r'create_function\("(\w+)"', read("tests/access_sqlite.py")))
        self.assertEqual(functions, harness, "the functions of tests/access_sqlite.py")

    def test_the_parameters_of_the_web_cover_the_report_queries(self):
        runner = read("web/src/RetailStore.Data/Reports/SavedQueryRunner.cs")
        web = re.search(r"WebParams =\s*\[([^\]]*)\]", runner).group(1)
        web = set(re.findall(r'"(\w+)"', web))
        self.assertLessEqual(web, set(SV.query_params()), "every one is a column of dbo.QueryParams")
        by = {q.name: q for q in Q.QUERIES}

        def deps(name, seen):
            for m in re.findall(r"\b(\w+)\b", by[name].sql):
                if m in by and m not in seen:
                    seen.add(m)
                    deps(m, seen)
            return seen

        for r in F.REPORTS:
            if r.key in G.LINKS:
                continue
            used = set()
            for n in deps(r.query, {r.query}):
                used |= set(re.findall(r"Q(?:Date|Long)\('(\w+)'\)", by[n].sql))
            with self.subTest(r.key):
                self.assertLessEqual(used, web)
        self.assertIn('"MERGE [dbo].[QueryParams] AS t', runner)
        self.assertIn("RollbackAsync", runner, "the parameters of a run never stay for another page")
        self.assertIn("!Names.Contains(query)", runner, "only saved query names reach the SQL")


class ServiceTests(unittest.TestCase):

    def test_the_report_service_follows_prepare_report(self):
        code = read("web/src/RetailStore.Data/Services/ReportService.cs")
        vba = read("src/vba/modScreens.bas") + read("src/vba/modReports.bas")
        # RequireChoice: ShowWarning "اختر " & Caption & "." with the captions of its calls
        vba += "".join(f"اختر {c}." for c in re.findall(r'RequireChoice\(frm!\w+, HasNeed\(needs, "\w"\), "\w+", "([^"]+)"\)', vba))
        for m in arabic_literals(code):
            with self.subTest(m):
                self.assertIn(m, vba)
        run = code[code.index("public async Task<ReportResult> RunAsync"):code.index("public static List<ReportRow> Sort")]
        order = ['"REPORTS_PROFIT"', '"AUDIT_LOG"', '"CASH_BOX"', '"JOURNAL"', "journal.SyncAsync", "أدخل تاريخ البداية",
                 "تاريخ النهاية يجب", '"اختر العميل."', '"اختر المورد."', '"اختر المنتج."', "SavedQueryRunner.RunAsync",
                 '"REPORT"']
        self.assertEqual([run.index(x) for x in order], sorted(run.index(x) for x in order))

    def test_the_indicators_follow_mod_indicators(self):
        code = read("web/src/RetailStore.Domain/Accounting/FinancialIndicators.cs")
        vba = read("src/vba/modIndicators.bas")
        for m in arabic_literals(code):
            with self.subTest(m):
                self.assertIn(m, vba)
        self.assertEqual(IR.gross_margin(1000, 750), 0.25)
        self.assertEqual(IR.stock_turnover(1200, 200, 400), 4)
        self.assertEqual(IR.collection_days(3000, 9000), 30)
        self.assertEqual(IR.current_ratio(3000, 2000), 1.5)
        service = read("web/src/RetailStore.Data/Services/IndicatorsService.cs")
        self.assertIn('"DASHBOARD_FINANCIAL"', service)
        self.assertIn("journal.SyncAsync", service, "the journal is brought up to date first")
        self.assertIn("CacheMinutes = 10", service)
        query = next(q for q in Q.QUERIES if q.name == "FinancialIndicatorsQuery").sql
        for code_ in ("41", "51", "11", "21", "1300", "1400", "140001", "149999", "'YEAR_CLOSE'", "'SALE'"):
            self.assertIn(code_, query)
            self.assertIn(code_.strip("'"), service)

    def test_the_pages_and_the_api(self):
        pages = read("web/src/RetailStore.UI/Services/WorkPages.cs")
        self.assertIn('new("reports", "frmReportCenter"', pages)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmReportCenter"], "REPORTS")
        api = read("web/src/RetailStore.Web/Api/ReportsApi.cs")
        self.assertIn('.Needs("REPORTS", "frmReportCenter")', api)
        self.assertIn('.Needs("DASHBOARD_FINANCIAL")', api)
        self.assertIn("ReportsApi.Map(secured);", read("web/src/RetailStore.Web/Api/ApiEndpoints.cs"))
        home = read("web/src/RetailStore.UI/Pages/Home.razor")
        self.assertIn("Indicators.IndicatorsAsync", home)


if __name__ == "__main__":
    unittest.main()
