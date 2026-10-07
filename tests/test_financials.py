"""Financial statements (IncomeStatementQuery, BalanceSheetQuery, modFinancials, frmFinancials,
rptIncomeStatement, rptBalanceSheet) on the SQLite mirror after the journal sync:
  * income statement: each section = its accounts, gross / operating / net profit from the
    sections, net profit = revenue - expenses of the period, the comparison column = the
    comparison period;
  * balance sheet: assets = liabilities + equity (now and at the comparison date), equity
    holds the profit not closed yet, every group total = its accounts, rows in printed order;
  * the screen, the reports, the report centre and the comparison rule in the VBA."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day
import forms as F
import queries as Q
import reports as RP
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}
REPORTS = {m.name: m for m in RP.all_reports()}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


class StatementTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        Store(self.db.con, now=day(0, 9)).sync_journal()
        self.db.set_period(30)                                   # 30 days ago .. today
        self.db.params["CompareStart"] = day(60)
        self.db.params["CompareEnd"] = day(30)                  # exclusive: 60 .. 31 days ago

    def rows(self, sql):
        return self.db.con.execute(sql).fetchall()

    def moves(self, start, end, types):
        return self.db.scalar(
            "SELECT Sum(l.Credit) - Sum(l.Debit) FROM (JournalLines AS l INNER JOIN JournalEntries AS e ON "
            "l.EntryID = e.EntryID) INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode "
            f"WHERE e.EntryDate >= '{start}' AND e.EntryDate < '{end}' AND a.AccountType IN ({types})") or 0

    def test_income_sections_and_results(self):
        r = {b: (cur, pri) for b, cur, pri in self.rows(
            "SELECT Block, CurrentValue, PriorValue FROM IncomeStatementQuery WHERE RowKind <> 'A' AND RowKind <> 'H'")}
        acc = {}
        for b, cur, pri in self.rows("SELECT Block, CurrentValue, PriorValue FROM IncomeStatementQuery WHERE RowKind = 'A'"):
            s = acc.setdefault(b, [0, 0])
            s[0] += cur
            s[1] += pri
        for n in range(1, 6):
            got = acc.get(n * 10 + 1, [0, 0])
            self.assertAlmostEqual(r[n * 10 + 2][0], got[0], places=4, msg=n)
            self.assertAlmostEqual(r[n * 10 + 2][1], got[1], places=4, msg=n)
        for col in (0, 1):
            s = {n: r[n * 10 + 2][col] for n in range(1, 6)}
            self.assertAlmostEqual(r[25][col], s[1] - s[2], places=4)
            self.assertAlmostEqual(r[35][col], s[1] - s[2] - s[3], places=4)
            self.assertAlmostEqual(r[60][col], s[1] - s[2] - s[3] + s[4] - s[5], places=4)
        types = "'REVENUE', 'EXPENSE'"
        self.assertAlmostEqual(r[60][0], self.moves(day(30), day(-1), types), places=4)
        self.assertAlmostEqual(r[60][1], self.moves(day(60), day(30), types), places=4)
        self.assertEqual(r[12][0], 1360)                         # sales 1400 - returns 40
        self.assertEqual(r[25][0], 550)

    def test_income_order(self):
        blocks = [b for b, in self.rows("SELECT Block FROM IncomeStatementQuery")]
        self.assertEqual(blocks, sorted(blocks))
        self.assertEqual([b for b in blocks if b % 10 != 1],
                         [10, 12, 20, 22, 25, 30, 32, 35, 40, 42, 50, 52, 60])

    def test_balance_sheet_balances(self):
        for col in ("CurrentValue", "PriorValue"):
            assets = self.db.scalar(f"SELECT {col} FROM BalanceSheetQuery WHERE ClassNo = 1 AND RowKind = 'T'")
            other = self.db.scalar(f"SELECT {col} FROM BalanceSheetQuery WHERE ClassNo = 4")
            self.assertAlmostEqual(assets, other, places=4, msg=col)
        self.assertEqual(self.db.scalar("SELECT CurrentValue FROM BalanceSheetQuery WHERE ClassNo = 1 AND RowKind = 'T'"),
                         14859)
        profit = self.db.scalar("SELECT CurrentValue FROM BalanceSheetQuery WHERE AccountKey = 'Z'")
        self.assertAlmostEqual(profit, self.moves("1900-01-01", day(-1), "'REVENUE', 'EXPENSE'"), places=4)
        self.assertEqual(self.db.scalar("SELECT CurrentValue FROM BalanceSheetQuery WHERE LineAccount = 1300"), 164)
        self.assertEqual(self.db.scalar("SELECT CurrentValue FROM BalanceSheetQuery WHERE LineAccount = 2100"), 2855,
                         "liabilities with their credit balance positive")

    def test_balance_group_totals_and_order(self):
        rows = self.rows("SELECT ClassNo, GroupKey, Pos, RowKind, CurrentValue FROM BalanceSheetQuery")
        totals = {}
        for cls, gk, pos, kind, cur in rows:
            if kind == "A":
                totals[gk] = totals.get(gk, 0) + cur
        for cls, gk, pos, kind, cur in rows:
            if kind == "S":
                self.assertAlmostEqual(cur, totals[gk], places=4, msg=gk)
        keys = [(r[0], r[1], r[2]) for r in rows]
        self.assertEqual(keys, sorted(keys))
        self.assertEqual([r[3] for r in rows][:2], ["C", "G"])
        self.assertEqual(rows[-1][3], "T")


class FinancialCodeTests(unittest.TestCase):

    def test_comparison_rule(self):
        body = proc(read("modQueryParams"), "SetComparePeriod")
        for part in ("If Day(FromDate) = 1 Then", 'cFrom = DateAdd("m", -k, FromDate)',
                     'If cTo >= FromDate Then cTo = DateAdd("d", -1, FromDate)', 'cTo = DateAdd("d", -1, FromDate)',
                     'TempVars.Add "CompareEnd", DateAdd("d", 1, cTo)'):
            self.assertIn(part, body)

    def test_screen_and_reports(self):
        text = read("modFinancials")
        names = {c.name for c in MODELS["frmFinancials"].controls}
        for pname in re.findall(r"^(?:Public|Private) (?:Sub|Function) (\w+)\(ByVal frm As Access\.Form", text, re.M):
            for ctl in set(re.findall(r"frm!(\w+)", proc(text, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        self.assertIn("SyncJournal()", proc(text, "FinancialsLoad"))
        self.assertIn("SetComparePeriod FromDate, ToDate", proc(text, "PrepareFinancials"))
        self.assertIn('PrepareFinancials(CStr(r(0)), fromDate, toDate)', read("modScreens"))
        keys = {r.key: r for r in F.REPORTS}
        for key, rpt in (("INCOME_STATEMENT", "rptIncomeStatement"), ("BALANCE_SHEET", "rptBalanceSheet")):
            self.assertEqual(keys[key].report, rpt)
            for need in "PJ$F":
                self.assertIn(need, keys[key].needs)
            self.assertIn(rpt, REPORTS)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmFinancials"], "REPORTS_PROFIT")

    def test_in_access_tests(self):
        body = proc(read("modJournal"), "TestJournal")
        self.assertIn("الميزانية متوازنة", body)
        self.assertIn("صافي الربح في قائمة الدخل", body)


class Static_modFinancials(VbaModuleChecks, unittest.TestCase):
    module_name = "modFinancials"
    vba = read("modFinancials")


if __name__ == "__main__":
    unittest.main()
