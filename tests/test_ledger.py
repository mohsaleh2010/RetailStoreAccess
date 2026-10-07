"""Account statement and general ledger (AccountStatementQuery, GeneralLedgerQuery, modLedger,
frmLedger, rptAccountStatement, rptGeneralLedger), on the SQLite mirror after the journal sync:
  * a sub-account: opening + debit - credit = its balance in the trial balance (customers 164);
  * a main account takes all its sub-accounts (cash, current assets) and equals the trial balance
    by levels; the period is respected (opening before, lines inside);
  * the general ledger has one group per sub-account with movement, each with its opening row,
    and adds up to the trial balance; filtered on a main account it keeps its sub-accounts only;
  * the screen, the reports and the VBA keep the same rules."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day
import forms as F
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


class LedgerQueryTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        Store(self.db.con, now=day(0, 9)).sync_journal()
        self.db.set_period(30)

    def statement(self, code):
        self.db.params["AccountCode"] = code
        rows = self.db.con.execute("SELECT SortKey, LineDebit, LineCredit, StatementAccount, StatementName, SubCode "
                                   "FROM AccountStatementQuery").fetchall()
        opening = [r for r in rows if r[0] == 0]
        self.assertEqual(len(opening), 1, "one opening row")
        lines = [r for r in rows if r[0] == 1]
        return (round(opening[0][1] - opening[0][2], 4), round(sum(r[1] for r in lines), 4),
                round(sum(r[2] for r in lines), 4), rows)

    def trial(self, code, query="TrialBalanceTreeQuery"):
        return self.db.con.execute(
            f"SELECT OpeningBalance, PeriodDebit, PeriodCredit, ClosingBalance FROM {query} WHERE AccountCode = ?",
            (code,)).fetchone()

    def test_sub_account(self):
        opening, debit, credit, rows = self.statement(1300)
        o, d, c, closing = self.trial(1300, "TrialBalanceQuery")
        self.assertEqual((opening, debit, credit), (round(o, 4), round(d, 4), round(c, 4)))
        self.assertEqual(round(opening + debit - credit, 4), 164)
        self.assertTrue(all(r[3] == 1300 and r[4] == "ذمم العملاء" for r in rows))

    def test_main_account_takes_its_sub_accounts(self):
        for code in (1100, 11, 2, 5):
            with self.subTest(code):
                opening, debit, credit, rows = self.statement(code)
                o, d, c, closing = self.trial(code)
                self.assertAlmostEqual(opening, o, places=4)
                self.assertAlmostEqual(debit, d, places=4)
                self.assertAlmostEqual(credit, c, places=4)
                self.assertAlmostEqual(opening + debit - credit, closing, places=4)
        rows = self.statement(1100)[3]
        self.assertEqual({r[5] for r in rows if r[0] == 1}, {110000 + self.db.ids["BOXC"], 110000 + self.db.ids["BOXM"]})

    def test_period(self):
        self.db.set_period(3)                       # the manual entry (3 days ago) is inside
        opening, debit, credit, rows = self.statement(5500)
        self.assertEqual((opening, debit, credit), (0, 3000, 0))
        self.db.set_period(1)                       # ...and now before the period: the opening balance
        opening, debit, credit, rows = self.statement(5500)
        self.assertEqual((opening, debit, credit), (3000, 0, 0))

    def test_account_without_entries(self):
        opening, debit, credit, rows = self.statement(1250)
        self.assertEqual((opening, debit, credit, len(rows)), (0, 0, 0, 1))

    def test_general_ledger(self):
        con = self.db.con
        self.db.params["AccountCode"] = 0
        groups = {}
        for code, sort, d, c in con.execute("SELECT LedgerCode, SortKey, LineDebit, LineCredit FROM GeneralLedgerQuery"):
            g = groups.setdefault(code, [0, 0])
            g[0] += sort == 0
            g[1] += d - c
        trial = dict(con.execute("SELECT AccountCode, ClosingBalance FROM TrialBalanceQuery").fetchall())
        self.assertEqual(set(groups), set(trial))
        for code, (openings, balance) in groups.items():
            self.assertEqual(openings, 1, code)
            self.assertAlmostEqual(balance, trial[code], places=4, msg=code)
        keys = [r[0] for r in con.execute("SELECT AccountKey FROM GeneralLedgerQuery")]
        self.assertEqual(keys, sorted(keys), "accounts in the order of the tree")
        self.db.params["AccountCode"] = 1100
        codes = {r[0] for r in con.execute("SELECT LedgerCode FROM GeneralLedgerQuery")}
        self.assertEqual(codes, {110000 + self.db.ids["BOXC"], 110000 + self.db.ids["BOXM"]})


class LedgerCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modLedger")

    def test_fill_and_balance(self):
        body = proc(self.text, "FillLedger")
        self.assertIn('SetQueryParam "AccountCode", AccountCode', body)
        self.assertIn("AccountStatementQuery ORDER BY SortKey, LineDate, EntryNo", body)
        self.assertIn("balance = balance + Nz(rs!LineDebit, 0) - Nz(rs!LineCredit, 0)", body)
        self.assertLess(body.index('DELETE FROM tmpLedger'), body.index("OpenRecordset"))
        self.assertIn("SyncJournal()", proc(self.text, "LedgerLoad"))
        prt = proc(self.text, "PrintLedger")
        self.assertIn('"rptAccountStatement", "AccountStatementQuery"', prt)
        self.assertIn('"rptGeneralLedger", "GeneralLedgerQuery"', prt)
        self.assertIn('SetQueryParam "AccountCode", account', prt)
        self.assertIn('SetQueryParam "AccountCode", 0', read("modScreens"), "report centre: every account")

    def test_screen_controls_exist(self):
        names = {c.name for c in MODELS["frmLedger"].controls}
        for pname in re.findall(r"^(?:Public|Private) (?:Sub|Function) (\w+)\(ByVal frm As Access\.Form", self.text, re.M):
            for ctl in set(re.findall(r"frm!(\w+)", proc(self.text, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        tmp = re.search(r'CREATE TABLE tmpLedger \((.*?)\)", dbFailOnError', read("modPOS"), re.S).group(1)
        for col in ("EntryRef", "DateText", "EntryNo", "KindName", "DocNo", "Details", "SubName", "DebitText",
                    "CreditText", "BalanceText"):
            self.assertIn(col, tmp)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmLedger"], "JOURNAL")

    def test_reports(self):
        for name, group in (("rptAccountStatement", "StatementAccount"), ("rptGeneralLedger", "AccountKey")):
            m = REPORTS[name]
            self.assertEqual(m.group, group)
            balance = [c for cs in m.controls.values() for c in cs if c.source == "=[LineDebit]-[LineCredit]"]
            self.assertEqual(len(balance), 1)
            self.assertEqual(balance[0].props.get("RunningSum"), 1, "running balance per account")
        self.assertIn("GENERAL_LEDGER", {r.key for r in F.REPORTS})

    def test_balance_text(self):
        body = proc(self.text, "BalanceText")
        self.assertIn('" مدين"', body)
        self.assertIn('" دائن"', body)


class Static_modLedger(VbaModuleChecks, unittest.TestCase):
    module_name = "modLedger"
    vba = read("modLedger")


if __name__ == "__main__":
    unittest.main()
