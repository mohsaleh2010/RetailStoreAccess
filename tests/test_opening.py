"""The start-up opening balances (modOpening, frmOpeningBalances, docs/52-Opening-Balances.md):
  * on the SQLite mirror: the general accounts (OpeningBalances) post ACCOUNT_OPENING entries against the
    opening balances account 3900 on Settings.OpeningDate; customers and suppliers move to that date (before
    it is set they keep the day of their card); qryOpeningTotals is the balance of 3900 by source and is zero
    when the balances entered are balanced;
  * the schema, the permission, the screen and its grids, the accounts allowed, the rules of the VBA code."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day
import forms as F
import generate as G
import queries as Q
from schema import ACCOUNT_TREE, SCREEN_LIST, table
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


VBA = read("modOpening")


class OpeningJournalTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.s = Store(self.db.con, now=day(0, 9))

    def one(self, sql, *args):
        return self.db.con.execute(sql, args).fetchone()[0]

    def total_3900(self):
        return round(self.one("SELECT Coalesce(Sum(SumDebit - SumCredit), 0) FROM qryOpeningTotals"), 4)

    def test_customers_keep_their_card_day_until_the_date_is_set(self):
        self.s.sync_journal()
        cust = self.db.con.execute("SELECT CustomerID, CreatedAt FROM Customers WHERE OpeningBalance <> 0").fetchone()
        self.assertIsNotNone(cust, "the fixture has a customer with an opening balance")
        on = self.one("SELECT EntryDate FROM JournalEntries WHERE SourceType = 'CUSTOMER_OPENING' AND SourceID = ?",
                      cust[0])
        self.assertEqual(on[:10], cust[1][:10])
        start = day(200)
        self.db.con.execute("UPDATE Settings SET OpeningDate = ?", (start,))
        self.s.sync_journal()
        dates = {r[0][:10] for r in self.db.con.execute(
            "SELECT EntryDate FROM JournalEntries WHERE SourceType IN ('CUSTOMER_OPENING', 'SUPPLIER_OPENING')")}
        self.assertEqual(dates, {start[:10]})

    def test_account_opening_against_3900(self):
        self.db.con.execute("UPDATE Settings SET OpeningDate = ?", (day(200),))
        self.s.sync_journal()
        before = self.total_3900()
        self.db.con.execute("INSERT INTO OpeningBalances (AccountCode, Debit, Credit, LineText) VALUES (3200, 0, 5000, 'x')")
        self.db.con.execute("INSERT INTO OpeningBalances (AccountCode, Debit, Credit, LineText) VALUES (1650, 1200, 0, 'y')")
        self.s.sync_journal()
        rows = self.db.con.execute(
            "SELECT e.EntryDate, l.AccountCode, l.Debit, l.Credit FROM JournalEntries AS e INNER JOIN JournalLines AS l "
            "ON l.EntryID = e.EntryID WHERE e.SourceType = 'ACCOUNT_OPENING' ORDER BY e.SourceID, l.AccountCode").fetchall()
        self.assertEqual(len(rows), 4)
        self.assertEqual({r[0][:10] for r in rows}, {day(200)[:10]})
        self.assertEqual(self.one("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0)
        self.assertAlmostEqual(self.total_3900(), before + 5000 - 1200, places=4)
        sources = {r[0] for r in self.db.con.execute("SELECT SourceType FROM qryOpeningTotals")}
        self.assertIn("ACCOUNT_OPENING", sources)

    def test_balanced_when_the_other_side_is_entered(self):
        self.db.con.execute("UPDATE Settings SET OpeningDate = ?", (day(200),))
        self.s.sync_journal()
        diff = self.total_3900()
        if diff > 0:
            self.db.con.execute("INSERT INTO OpeningBalances (AccountCode, Debit, Credit) VALUES (3200, ?, 0)", (diff,))
        elif diff < 0:
            self.db.con.execute("INSERT INTO OpeningBalances (AccountCode, Debit, Credit) VALUES (3200, 0, ?)", (-diff,))
        self.s.sync_journal()
        self.assertEqual(self.total_3900(), 0)

    def test_without_a_date_no_account_entry(self):
        self.db.con.execute("UPDATE Settings SET OpeningDate = Null")
        self.db.con.execute("INSERT INTO OpeningBalances (AccountCode, Debit, Credit) VALUES (3200, 0, 10)")
        self.s.sync_journal()
        self.assertEqual(self.one("SELECT COUNT(*) FROM JournalEntries WHERE SourceType = 'ACCOUNT_OPENING'"), 0)


class OpeningSchemaTests(unittest.TestCase):

    def test_schema_and_names(self):
        self.assertIn("OpeningDate", {f.name for f in table("Settings").fields})
        t = table("OpeningBalances")
        self.assertEqual([f.name for f in t.fields][:5], ["OpeningLineID", "AccountCode", "Debit", "Credit", "LineText"])
        self.assertIn("[Debit]=0 Or [Credit]=0", t.rule)
        self.assertIn(("ACCOUNT_OPENING", "رصيد افتتاحي لحساب", 28), table("JournalSourceTypes").seed_rows)
        self.assertIn((3900, "إجمالي الأرصدة الافتتاحية"), [(r[0], r[1]) for r in ACCOUNT_TREE])
        upgrade = read("modBuildSchema")
        self.assertIn("SET [AccountName] = 'إجمالي الأرصدة الافتتاحية' WHERE [AccountCode] = 3900 "
                      "AND [AccountName] = 'أرصدة افتتاحية'", upgrade)

    def test_permission_and_screen(self):
        self.assertIn("OPENING_BALANCES", [r[0] for r in table("Permissions").seed_rows])
        self.assertIn(("frmOpeningBalances", "أرصدة بداية التشغيل", "الحسابات", "OPENING_BALANCES", True, True, True),
                      SCREEN_LIST)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmOpeningBalances"], "OPENING_BALANCES")
        hub = {c.props.get("Tag") for c in MODELS["frmAccounting"].controls}
        self.assertIn("frmOpeningBalances", hub)
        self.assertIn('Case "ACCOUNT_OPENING":  OpenScreen "frmOpeningBalances"', read("modJournal"))

    def test_screen_controls_exist(self):
        names = {c.name for c in MODELS["frmOpeningBalances"].controls}
        for pname in re.findall(r"^Public Sub (\w+)\(ByVal frm As Access\.Form", VBA, re.M):
            body = proc(VBA, pname)
            if pname.startswith(("OpeningLine", "OpeningParty", "OpeningGrid", "OpeningRemove")):
                continue                          # the grids
            for ctl in set(re.findall(r"frm!(\w+)", body)):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        grids = re.findall(r'grid = "(\w+)"', proc(VBA, "OpeningShowPart"))
        self.assertEqual(sorted(grids), sorted(["frmOpeningAccountLines", "frmOpeningCustomers", "frmOpeningSuppliers",
                                                "frmOpeningBoxes", "frmOpeningBanks"]))
        for g in grids:
            self.assertIn(g, MODELS)
            self.assertIn("OpeningGridAfterUpdate Me", "\n".join(MODELS[g].code))
            self.assertIn("BeforeUpdate", MODELS[g].form_events)

    def test_party_grids_bind_the_fields_the_code_uses(self):
        for name, tbl in [("frmOpeningCustomers", "Customers"), ("frmOpeningSuppliers", "Suppliers"),
                          ("frmOpeningBoxes", "CashBoxes"), ("frmOpeningBanks", "Banks")]:
            sql = MODELS[name].record_source
            self.assertIn(f"FROM {tbl}", sql)
            self.assertIn("OpeningBalance", sql)
            combo = next(c for c in MODELS[name].controls if c.name == "cboName")
            self.assertIn("[@", combo.props["RowSource"])           # the English name in an English front-end
            if tbl in ("Customers", "Suppliers"):
                self.assertIn("CurrentBalance", sql)
                self.assertIn("CreatedAt", sql)
            else:
                self.assertIn("OpeningDate", sql)

    def test_accounts_allowed(self):
        rows = next(c for c in MODELS["frmOpeningAccountLines"].controls if c.name == "AccountCode").props["RowSource"]
        self.assertIn("NOT IN (1300, 2100, 1400, 3900)", rows)
        body = proc(VBA, "OpeningAccountProblem")
        for case in ("Case 1300", "Case 2100", "Case 1400", "Case OPENING_ACCOUNT", "Case 110001 To 119999",
                     "Case 120001 To 129999"):
            self.assertIn(case, body)


class OpeningCodeTests(unittest.TestCase):

    def test_date_rules(self):
        body = proc(VBA, "OpeningDateProblem")
        for part in ("FirstDocumentDate()", "ClosedPeriodProblem(NewDate)", "ClosedPeriodProblem(old)"):
            self.assertIn(part, body)
        save = proc(VBA, "SaveOpeningDate")
        for part in ('CanOpening("EDIT")', "OpeningDateProblem(NewDate)", "UPDATE CashBoxes SET OpeningDate",
                     "UPDATE Banks SET OpeningDate", 'AuditEdited "OPENING_DATE"', "SyncJournal()"):
            self.assertIn(part, save)
        self.assertLess(save.index("SyncJournal()"), save.index("OpeningDateProblem(NewDate)"))

    def test_grids_check_and_audit(self):
        line = proc(VBA, "OpeningLineBeforeUpdate")
        for part in ("CanOpening(", "ClosedPeriodProblem(", "OpeningAccountProblem(", 'AuditFormBefore frm, "OpeningBalances"'):
            self.assertIn(part, line)
        party = proc(VBA, "OpeningPartyBeforeUpdate")
        for part in ("CanOpening(", "ClosedPeriodProblem(oldDay)", "frm!CurrentBalance.Value", "AuditFormBefore frm, TableName"):
            self.assertIn(part, party)
        remove = proc(VBA, "OpeningRemoveLine")
        self.assertIn('AuditSnapshot("OpeningBalances"', remove)
        self.assertIn('AuditDeleted "OPENING_DELETE"', remove)
        self.assertLess(remove.index("AuditSnapshot("), remove.index("DELETE FROM OpeningBalances"))
        self.assertIn("OpeningEntryDate(frm!CreatedAt.Value)", proc(read("modClosing"), "ClosedRecordProblem"))

    def test_registered(self):
        self.assertIn("modOpening", G.STATIC_MODULES)
        self.assertIn('"TestOpeningBalances"', read("modTestAll"))
        self.assertIn("qryJournalOpening", Q.JOURNAL_SOURCE_QUERIES)
        for q in ("qryCustomerOpening", "qrySupplierOpening", "qryOpeningAccountLines", "qryOpeningTotals"):
            self.assertIn(q, {x.name for x in Q.QUERIES})


class StaticModule(VbaModuleChecks, unittest.TestCase):
    module_name = "modOpening"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
