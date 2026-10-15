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
import forms_opening as FO
import generate as G
import queries as Q
from schema import ACCOUNT_TREE, SCREEN_LIST, table
from sim import Store
from decimal import Decimal as D

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


def first_document_sql():
    """FirstDocumentDate of modOpening, with its constants: an opening stock move is not a document."""
    body = proc(VBA, "FirstDocumentDate")
    sources = re.search(r'OPENING_SOURCES As String = "(.*?)"', VBA).group(1)
    sql = "".join(re.findall(r'"((?:[^"]|"")*)"', body.split("=", 1)[1]))
    return sql.replace("SourceType NOT IN (", "SourceType NOT IN (" + sources).replace(
        "TransactionTypeID = ", "TransactionTypeID = 8")


class OpeningStockTests(unittest.TestCase):
    """The stock part: PostOpeningStock mirrored on the SQLite copy (one move "opening" per line on the
    start-up date, the line keeps the move), its entry against 3900, and the first document date."""

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.s = Store(self.db.con, now=day(0, 9))
        self.start = day(-400)
        self.db.con.execute("UPDATE Settings SET OpeningDate = ?", (self.start,))

    def one(self, sql, *args):
        row = self.db.con.execute(sql, args).fetchone()
        return row[0] if row else None

    def free_product(self):
        pid = self.one("SELECT ProductID FROM Products WHERE TrackStock = 1 AND ProductID NOT IN "
                       "(SELECT ProductID FROM InventoryTransactions) ORDER BY ProductID")
        if pid is None:                         # every product of the fixture has moves: take a copy
            self.db.con.execute("INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, TrackStock) "
                                "SELECT 'OPEN-T', 'opening test', CategoryID, UnitID, 1 FROM Products LIMIT 1")
            pid = self.one("SELECT Max(ProductID) FROM Products")
        return pid

    def post(self):                              # PostOpeningStock
        rows = self.db.con.execute("SELECT OpeningStockID, ProductID, Quantity, UnitCost FROM OpeningStock "
                                   "WHERE TransactionID IS NULL AND Quantity > 0 ORDER BY OpeningStockID").fetchall()
        for line, pid, q, cost in rows:
            if self.one("SELECT COUNT(*) FROM InventoryTransactions WHERE ProductID = ?", pid):
                continue
            self.s.apply_stock(pid, D(str(q)), 8, D(str(cost)), "MANUAL", 0, "ADJ-T", recalc=True)
            tid = self.one("SELECT Max(TransactionID) FROM InventoryTransactions WHERE ProductID = ?", pid)
            self.db.con.execute("UPDATE InventoryTransactions SET TransactionDate = ? WHERE TransactionID = ?",
                                (self.start, tid))
            self.db.con.execute("UPDATE OpeningStock SET TransactionID = ? WHERE OpeningStockID = ?", (tid, line))

    def test_post_dates_the_move_and_posts_against_3900(self):
        self.s.sync_journal()
        first = self.one(first_document_sql())
        pid = self.free_product()
        self.db.con.execute("INSERT INTO OpeningStock (ProductID, Quantity, UnitCost) VALUES (?, 12, 2.5)", (pid,))
        self.post()
        tid = self.one("SELECT TransactionID FROM OpeningStock WHERE ProductID = ?", pid)
        self.assertIsNotNone(tid)
        self.assertEqual(self.one("SELECT TransactionDate FROM InventoryTransactions WHERE TransactionID = ?", tid)[:10],
                         self.start[:10])
        self.assertEqual(self.s.stock(pid), 12)
        self.assertEqual(self.s.avg(pid), D("2.5"))
        self.s.sync_journal()
        lines = dict(self.db.con.execute(
            "SELECT l.AccountCode, l.Debit - l.Credit FROM JournalLines AS l INNER JOIN JournalEntries AS e "
            "ON l.EntryID = e.EntryID WHERE e.SourceType = 'STOCK_MOVE' AND e.SourceID = ?", (tid,)).fetchall())
        self.assertEqual(lines, {1400: 30, 3900: -30})
        self.assertEqual(self.one(first_document_sql()), first, "an opening stock move is not the first document")
        self.assertIn("STOCK_MOVE", {r[0] for r in self.db.con.execute("SELECT SourceType FROM qryOpeningTotals")})

    def test_a_product_with_moves_is_not_posted(self):
        pid = self.one("SELECT ProductID FROM InventoryTransactions ORDER BY ProductID")
        self.db.con.execute("INSERT INTO OpeningStock (ProductID, Quantity, UnitCost) VALUES (?, 5, 1)", (pid,))
        self.post()
        self.assertIsNone(self.one("SELECT TransactionID FROM OpeningStock WHERE ProductID = ?", pid))

    def test_vba_rules(self):
        post = proc(VBA, "PostOpeningStock")
        for part in ('OpeningStockReady("ADD")', "ProductHasMoves(pid)", "TT_OPENING", '"MANUAL"',
                     "SET TransactionDate = \" & SqlDate(startDay)", "UPDATE OpeningStock SET TransactionID",
                     "ws.Rollback", "SyncJournal()"):
            self.assertIn(part, post)
        ready = proc(VBA, "OpeningStockReady")
        self.assertIn("ClosedPeriodProblem(OpeningDateValue())", ready)
        unpost = proc(VBA, "UnpostOpeningStock")
        for part in ("TransactionID <> ", 'AuditSnapshot("InventoryTransactions"', 'AuditDeleted "OPENING_STOCK_UNPOST"',
                     "CurrentQuantity = 0"):
            self.assertIn(part, unpost)
        self.assertLess(unpost.index("AuditSnapshot("), unpost.index("DELETE FROM InventoryTransactions"))
        remove = proc(VBA, "OpeningStockRemoveLine")
        self.assertLess(remove.index('AuditSnapshot("OpeningStock"'), remove.index("DELETE FROM OpeningStock"))
        row = proc(VBA, "ImportOpeningStockRow")
        self.assertLess(row.index('"POSTED"'), row.index('"MOVES"'))      # a posted line is not "has moves"
        self.assertIn("ProductCode = ", row)
        self.assertIn("Barcode = ", row)
        self.assertIn("UPDATE InventoryTransactions SET TransactionDate", proc(VBA, "SaveOpeningDate"))
        before = proc(VBA, "OpeningStockBeforeUpdate")
        self.assertIn("frm!TransactionID.Value", before)
        self.assertIn('AuditFormBefore frm, "OpeningStock"', before)
        show = proc(VBA, "OpeningShowPart")
        for name, _, _ in FO.STOCK_BUTTONS:
            self.assertIn(f'frm!{name}.Visible = (Part = "STOCK")', show)

    def test_import_by_position_and_template_columns(self):
        imp = proc(VBA, "ImportOpeningStock")
        self.assertIn("rs.Fields.Count >= 4", imp)              # the template: code, name, quantity, cost
        self.assertIn("DropImportTables db", imp)
        self.assertIn('"tmpOpeningImport"', imp)
        sql = proc(VBA, "OpeningStockTemplateSql")
        self.assertLess(sql.index("[كود الصنف]"), sql.index("[اسم الصنف]"))
        self.assertLess(sql.index("[الكمية الافتتاحية]"), sql.index("[تكلفة الوحدة الافتتاحية]"))

    def test_exported_files_go_into_the_reports_folder(self):
        for name in os.listdir(os.path.join(ROOT, "src", "vba")):
            text_ = read(name[:-4])
            for m in re.finditer(r'ReportsFolder\(\) & "(.)', text_):
                self.assertEqual(m.group(1), "\\", f"{name}: the file name needs a backslash after the folder")


class OpeningSchemaTests(unittest.TestCase):

    def test_stock_schema_and_grid(self):
        t = table("OpeningStock")
        self.assertEqual([f.name for f in t.fields][:5], ["OpeningStockID", "ProductID", "Quantity", "UnitCost",
                                                          "TransactionID"])
        self.assertIn("ProductID", [f for i in t.indexes if i.unique for f in i.fields])
        grid = MODELS["frmOpeningStock"]
        self.assertIn("FROM OpeningStock", grid.record_source)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmOpeningStock"], "OPENING_BALANCES")
        self.assertIn('grid = "frmOpeningStock"', proc(VBA, "OpeningShowPart"))
        names = {c.name for c in MODELS["frmOpeningBalances"].controls}
        for name, _, call in FO.STOCK_BUTTONS:
            self.assertIn(name, names)
            self.assertIn(f"Public Sub {call.split()[0]}(ByVal frm As Access.Form)", VBA)

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
            if pname.startswith(("OpeningLine", "OpeningParty", "OpeningGrid", "OpeningRemove", "OpeningStockBefore",
                                 "OpeningStockRemove")):
                continue                          # the grids
            for ctl in set(re.findall(r"frm!(\w+)", body)):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        grids = re.findall(r'grid = "(\w+)"', proc(VBA, "OpeningShowPart"))
        self.assertEqual(sorted(grids), sorted(["frmOpeningAccountLines", "frmOpeningCustomers", "frmOpeningSuppliers",
                                                "frmOpeningBoxes", "frmOpeningBanks", "frmOpeningStock"]))
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
