"""Financial indicators of the dashboard (modIndicators, FinancialIndicatorsQuery, cards 9-12 of frmMain):
the query runs on the SQLite mirror after the fixture's journal is built and agrees with the journal; the
VBA formulas really run in LibreOffice and match tools/indicators_reference.py."""
import datetime as dt
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day, TODAY
import forms as F
import indicators_reference as R
import vba_harness as H
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


def stamp(d: dt.date) -> str:
    return d.isoformat() + " 00:00:00"


class IndicatorQueryTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.c = self.db.con
        self.s = Store(self.c, now=day(0, 9))
        self.s.sync_journal()
        first = TODAY.replace(day=1)
        prev = (first - dt.timedelta(days=1)).replace(day=1)
        p = self.db.params                                     # = SetIndicatorParams TODAY
        p["IndEnd"] = stamp(TODAY + dt.timedelta(days=1))
        p["IndMonth"] = stamp(first)
        p["IndPrevMonth"] = stamp(prev)
        p["IndYear"] = stamp(TODAY - dt.timedelta(days=364))
        p["Ind90"] = stamp(TODAY - dt.timedelta(days=89))
        cur = self.c.execute("SELECT * FROM FinancialIndicatorsQuery")
        self.row = dict(zip([d[0] for d in cur.description], cur.fetchone()))

    def journal(self, where, expr, *args):
        return round(self.s.one(f"SELECT Coalesce(Sum({expr}), 0) FROM JournalLines AS l INNER JOIN JournalEntries AS e "
                                f"ON l.EntryID = e.EntryID INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode "
                                f"WHERE {where}", *args), 4)

    def test_one_row_that_agrees_with_the_journal(self):
        p, r = self.db.params, self.row
        self.assertAlmostEqual(r["CurrentAssets"], self.journal("a.Level2Code = 11 AND e.EntryDate < ?",
                                                                "l.Debit - l.Credit", p["IndEnd"]), places=2)
        self.assertAlmostEqual(r["CurrentLiabilities"], self.journal("a.Level2Code = 21 AND e.EntryDate < ?",
                                                                     "l.Credit - l.Debit", p["IndEnd"]), places=2)
        self.assertAlmostEqual(r["StockEnd"], self.journal("l.AccountCode = 1400", "l.Debit - l.Credit"), places=2)
        self.assertAlmostEqual(r["Receivables"], self.journal("l.AccountCode = 1300", "l.Debit - l.Credit"), places=2)
        self.assertAlmostEqual(r["MonthSales"], self.journal(
            "a.Level2Code = 41 AND e.EntryDate >= ? AND e.SourceType <> 'YEAR_CLOSE'", "l.Credit - l.Debit",
            p["IndMonth"]), places=2)
        self.assertAlmostEqual(r["YearCost"], self.journal(
            "a.Level2Code = 51 AND e.EntryDate >= ? AND e.SourceType <> 'YEAR_CLOSE'", "l.Debit - l.Credit",
            p["IndYear"]), places=2)
        self.assertGreater(r["YearCost"], 0)                   # the fixture sells in the last 365 days
        self.assertGreater(r["CurrentAssets"], 0)

    def test_receivables_match_the_customer_balances(self):
        self.assertAlmostEqual(self.row["Receivables"],
                               self.s.one("SELECT Coalesce(Sum(Balance), 0) FROM CustomerBalanceQuery"), places=2)

    def test_formulas_on_the_row(self):
        r = self.row
        margin = R.gross_margin(r["MonthSales"], r["MonthCost"])
        self.assertTrue(margin is None or margin < 1)
        turnover = R.stock_turnover(r["YearCost"], r["StockStart"], r["StockEnd"])
        if turnover is not None:
            self.assertGreater(turnover, 0)
            self.assertAlmostEqual(R.stock_days(turnover) * turnover, 365)


class ReferenceTests(unittest.TestCase):

    def test_formulas(self):
        self.assertEqual(R.gross_margin(1000, 750), 0.25)
        self.assertIsNone(R.gross_margin(0, 10))
        self.assertEqual(R.stock_turnover(1200, 200, 400), 4)
        self.assertIsNone(R.stock_turnover(100, 0, 0))
        self.assertEqual(R.collection_days(3000, 9000), 30)
        self.assertEqual(R.collection_days(-5, 900), 0)
        self.assertIsNone(R.collection_days(100, 0))
        self.assertEqual(R.current_ratio(3000, 2000), 1.5)
        self.assertIsNone(R.current_ratio(100, 0))


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class FormulaRuntimeTests(unittest.TestCase):

    def test_vba_matches_the_reference(self):
        h = H.Harness()
        try:
            # a Currency value does not come back through UNO: these wrappers return a Double
            # (-999 for Null, "no meaning")
            wrap = "Option Explicit\n" + "\n".join(
                f"Public Function W{n}({', '.join(f'ByVal p{i} As Double' for i in range(k))}) As Double\n"
                f"    Dim v As Variant\n    v = {n}({', '.join(f'CCur(p{i})' for i in range(k))})\n"
                f"    If IsNull(v) Then\n        W{n} = -999\n    Else\n        W{n} = CDbl(v)\n    End If\n"
                "End Function" for n, k in [("GrossMargin", 2), ("StockTurnover", 3), ("CollectionDays", 2),
                                            ("CurrentRatio", 2)])
            h.load({"modIndicators": H.read_module("modIndicators"), "modWrap": wrap})

            def same(name, args, want):
                got = h.call("modWrap", "W" + name, *[float(a) for a in args])
                if want is None:
                    self.assertEqual(got, -999, (name, args))
                else:
                    self.assertAlmostEqual(got, want, places=4, msg=(name, args))

            for args in [(1000, 750), (500, 600), (0, 10), (123.45, 67.8)]:
                same("GrossMargin", args, R.gross_margin(*args))
            for args in [(1200, 200, 400), (100, 0, 0), (0, 50, 50), (5000, 1000, 3000)]:
                same("StockTurnover", args, R.stock_turnover(*args))
            for args in [(3000, 9000), (-5, 900), (100, 0), (12500, 40000)]:
                same("CollectionDays", args, R.collection_days(*args))
            for args in [(3000, 2000), (100, 0), (900, 1000)]:
                same("CurrentRatio", args, R.current_ratio(*args))
        finally:
            h.close()


class IndicatorCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modIndicators")

    def test_parameters_and_cache(self):
        params = proc(self.text, "SetIndicatorParams")
        for part in ('"IndEnd", AsOf + 1', '"IndMonth", DateSerial(Year(AsOf), Month(AsOf), 1)',
                     '"IndPrevMonth", DateSerial(Year(AsOf), Month(AsOf) - 1, 1)', '"IndYear", AsOf - 364',
                     '"Ind90", AsOf - (COLLECTION_DAYS_SPAN - 1)'):
            self.assertIn(part, params)
        main = proc(self.text, "FinancialIndicators")
        self.assertIn("DateDiff(\"n\", m_time, Now) < CACHE_MINUTES", main)
        self.assertLess(main.index("SyncJournal"), main.index("FinancialIndicatorsQuery"))

    def test_no_division_inside_iif(self):
        # IIf evaluates both branches: a division there fails when the divisor is 0
        for line in self.text.splitlines():
            for m in re.finditer(r"IIf\(([^()]|\([^()]*\))*\)", line):
                self.assertNotIn("/", m.group(0), line.strip())

    def test_dashboard_cards(self):
        dash = read("modDashboard")
        self.assertIn("ShowIndicators frm, financial, Force", proc(dash, "DashboardRefresh"))
        show = proc(dash, "ShowIndicators")
        self.assertIn("If Not financial Then", show)
        self.assertIn("ind = FinancialIndicators(Date, Force)", show)
        keys = [k for k, _ in F.DASHBOARD_TILES]
        self.assertEqual(keys[8:], ["MARGIN", "TURNOVER", "COLLECTION", "LIQUIDITY"])
        main = MODELS["frmMain"]
        names = {c.name for c in main.controls}
        for n in range(9, 13):
            for part in ("boxTile", "lblTileTitle", "lblTileValue", "lblTileSub"):
                self.assertIn(f"{part}{n}", names)
        self.assertIn('DashboardRefresh Me, True', "\n".join(main.code))
        self.assertLessEqual(main.height, F.cm(17.0))
        self.assertIn('"TestIndicators"', read("modTestAll"))


class Static_modIndicators(VbaModuleChecks, unittest.TestCase):
    module_name = "modIndicators"
    vba = read("modIndicators")


if __name__ == "__main__":
    unittest.main()
