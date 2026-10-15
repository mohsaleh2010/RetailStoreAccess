"""The fiscal periods (modFiscal, frmFiscalPeriods, docs/54-Fiscal-Periods.md):
  * the opening balances report (OpeningBalancesReportQuery) on the demo data replayed in the SQLite mirror:
    the accounts balance, the customers and suppliers match accounts 1300 / 2100, the stock matches 1400;
  * the rules of the VBA code: periods in order without overlap or gap, the key of a closing = the year
    the period ends in, the year closing (modClosing) on the days of the period, the quick periods."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import day
import demo_data as DD
import forms as F
import generate as G
from schema import SCREEN_LIST, table

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


VBA = read("modFiscal")
CLOSING = read("modClosing")


class OpeningReportTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.r = DD.simulate()
        cls.c = cls.r.db.con

    def report(self, start):
        self.r.db.params["PeriodStart"] = start
        return self.c.execute('SELECT SectionNo, ItemCode, DebitBalance, CreditBalance, ItemQty '
                              'FROM "OpeningBalancesReportQuery"').fetchall()

    def account(self, code, start):
        return self.c.execute("SELECT Coalesce(Sum(l.Debit - l.Credit), 0) FROM JournalLines AS l INNER JOIN "
                              "JournalEntries AS e ON l.EntryID = e.EntryID WHERE l.AccountCode = ? AND e.EntryDate < ?",
                              (code, start)).fetchone()[0]

    def test_sections_match_their_accounts(self):
        start = day(-1)                                   # tomorrow: every operation is before it
        rows = self.report(start)
        by = {n: [r for r in rows if r[0] == n] for n in (1, 2, 3, 4)}
        for n in (1, 2, 3, 4):
            self.assertTrue(by[n], f"section {n} has rows")
        debit = sum(r[2] for r in by[1])
        credit = sum(r[3] for r in by[1])
        self.assertAlmostEqual(debit, credit, places=2)                  # the trial balance is balanced
        self.assertAlmostEqual(sum(r[2] - r[3] for r in by[2]), self.account(1300, start), places=2)
        self.assertAlmostEqual(sum(r[3] - r[2] for r in by[3]), -self.account(2100, start), places=2)
        self.assertAlmostEqual(sum(r[2] for r in by[4]), self.account(1400, start), delta=1.0)
        self.assertTrue(all(r[4] != 0 for r in by[4]))

    def test_nothing_before_the_first_operation(self):
        self.assertEqual(self.report(day(5000)), [])


class FiscalCodeTests(unittest.TestCase):

    def test_schema_screen_and_registration(self):
        t = table("FiscalPeriods")
        self.assertEqual([f.name for f in t.fields][:4], ["FiscalPeriodID", "PeriodName", "StartDate", "EndDate"])
        self.assertIn("[EndDate]>=[StartDate]", t.rule)
        self.assertIn(("frmFiscalPeriods", "الفترات المالية", "الحسابات", "PERIOD_CLOSE", True, False, True), SCREEN_LIST)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmFiscalPeriods"], "PERIOD_CLOSE")
        self.assertIn("modFiscal", G.STATIC_MODULES)
        self.assertIn('"TestFiscalPeriods"', read("modTestAll"))
        closing = {c.name for c in MODELS["frmPeriodClosing"].controls}
        self.assertTrue({"btnFiscalPeriods", "btnOpeningSheet"} <= closing)
        self.assertIn('OpenScreen "frmFiscalPeriods", 0', "\n".join(MODELS["frmPeriodClosing"].code))
        self.assertIn("OPENING_BALANCES", {e.key for e in F.REPORTS})

    def test_screen_controls_exist(self):
        names = {c.name for c in MODELS["frmFiscalPeriods"].controls}
        for pname in re.findall(r"^Public Sub (\w+)\(ByVal frm As Access\.Form", VBA, re.M):
            for ctl in set(re.findall(r"frm!(\w+)", proc(VBA, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        code = "\n".join(MODELS["frmFiscalPeriods"].code)
        for call in ("FiscalPeriodsLoad Me", "FiscalAddPeriod Me", "FiscalDeleteLast Me", "FiscalPrintSelected Me, True",
                     "FiscalPrintSelected Me, False"):
            self.assertIn(call, code)

    def test_period_rules(self):
        body = proc(VBA, "FiscalPeriodProblem")
        for part in ("DateAdd(\"d\", 1, DateValue(last))", "Year(EndDay) = Year(last)", "DateAdd(\"m\", 24, StartDay)",
                     "Max(ClosingDate) FROM FiscalYearClosings", "Day(StartDay) <> 1 Or Month(StartDay) <> 1",
                     "SourceType <> 'YEAR_CLOSE'"):
            self.assertIn(part, body)
        delete = proc(VBA, "DeleteLastFiscalPeriod")
        self.assertIn("ORDER BY StartDate DESC", delete)
        self.assertIn("YearIsClosed(Year(endDay))", delete)
        self.assertLess(delete.index('AuditSnapshot("FiscalPeriods"'), delete.index("DELETE FROM FiscalPeriods"))
        self.assertIn('AuditDeleted "FISCAL_PERIOD_DELETE"', delete)
        for name in ("AddFiscalPeriod", "DeleteLastFiscalPeriod"):
            self.assertIn("CanFiscal(", proc(VBA, name))

    def test_keys(self):
        for name, default in (("FiscalYearStart", "DateSerial(FiscalKey, 1, 1)"),
                              ("FiscalYearEnd", "DateSerial(FiscalKey, 12, 31)")):
            body = proc(VBA, name)
            self.assertIn("WHERE Year(EndDate) = ", body)
            self.assertIn(default, body)
        keys = proc(VBA, "FiscalKeys")
        self.assertIn("lastCalendar = Year(first) - 1", keys)
        self.assertIn("ORDER BY StartDate", keys)

    def test_closing_uses_the_period_days(self):
        for name in ("YearNetProfit", "YearCloseProblem", "CloseFiscalYear", "ReopenFiscalYear", "PeriodClosingLoad"):
            body = proc(CLOSING, name)
            self.assertNotIn("DateSerial(FiscalYear", body, name)
            self.assertNotIn("DateSerial(y", body, name)
        self.assertIn("yearEnd = FiscalYearEnd(FiscalYear)", proc(CLOSING, "CloseFiscalYear"))
        self.assertIn('DateAdd("d", -1, FiscalYearStart(FiscalYear))', proc(CLOSING, "ReopenFiscalYear"))
        self.assertIn("FiscalKeys()", proc(CLOSING, "YearCloseProblem"))
        quick = proc(read("modScreens"), "SetQuickPeriod")
        self.assertIn("FiscalPeriodOf Date, fyStart, fyEnd", quick)
        self.assertIn('FiscalPeriodOf DateAdd("d", -1, fyStart), fyStart, fyEnd', quick)

    def test_opening_sheet_is_the_day_before(self):
        body = proc(VBA, "PrintOpeningBalanceSheet")
        self.assertIn('asOf = DateAdd("d", -1, FiscalYearStart(FiscalKey))', body)
        self.assertIn('PrepareFinancials("BALANCE_SHEET", asOf, asOf)', body)
        report = proc(VBA, "PrintOpeningBalances")
        self.assertIn("SetPeriod StartDay, StartDay", report)
        self.assertIn('"rptOpeningBalances", "OpeningBalancesReportQuery"', report)


class StaticModule(VbaModuleChecks, unittest.TestCase):
    module_name = "modFiscal"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
