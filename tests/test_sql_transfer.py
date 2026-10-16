"""Plan 5B (docs/68): moving the back-end to SQL Server from the program (modSqlServer, frmSqlServer).

  * the pure VBA parts (SqlLiteral, ServerConnect, SplitSqlBatches, InsertBatchSql) are the Python mirror
    (tools/sql_transfer_reference.py), checked in LibreOffice;
  * the move itself (CopyTable's INSERT batches) runs into a database built from the SQL Server script pieces:
    the demo data and the fixture keep every rule and relationship of the server, and the comparison
    (rows, amount totals, first character of the texts) finds them equal - and finds a changed amount or
    Arabic text that became "?";
  * the code: the order of the move, the link only after a clean comparison, the start-up check, the screen.
"""
import datetime as dt
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite
import demo_data as DD
import forms as F
import generate as GEN
import sql_transfer_reference as R
import vba_harness as H
from schema import SCREEN_LIST, TABLES, table


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


VBA = read("modSqlServer")
SCRIPT = os.path.join(ROOT, "dist", "sql", "RetailStore_Schema.sql")


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


LITERALS = [(None, 10, "NULL"), (True, 1, "1"), (False, 1, "0"), (-1, 1, "1"), ("O'Brien", 10, "N'O''Brien'"),
            ("مؤسسة النور", 12, "N'مؤسسة النور'"), (12.5, 5, "12.5"), (-3, 4, "-3"), (0, 5, "0"),
            (1234567.8912, 5, "1234567.8912"), ("2026-03-09 14:05:07", 8, "'20260309 14:05:07'"),
            (dt.datetime(2026, 1, 2), 8, "'20260102 00:00:00'")]


class MirrorTests(unittest.TestCase):

    def test_literals(self):
        for v, kind, expected in LITERALS:
            self.assertEqual(R.sql_literal(v, kind), expected, (v, kind))

    def test_connection(self):
        self.assertEqual(R.server_connect(".\\SQLEXPRESS", "RetailStore", "", "", "ODBC Driver 18 for SQL Server"),
                         "ODBC;DRIVER={ODBC Driver 18 for SQL Server};SERVER=.\\SQLEXPRESS;DATABASE=RetailStore;"
                         "Trusted_Connection=Yes;TrustServerCertificate=Yes;APP=RetailStore;")
        self.assertEqual(R.server_connect("srv", "db", "shop", "p1", "ODBC Driver 17 for SQL Server"),
                         "ODBC;DRIVER={ODBC Driver 17 for SQL Server};SERVER=srv;DATABASE=db;UID=shop;PWD=p1;"
                         "APP=RetailStore;")

    def test_batches_of_the_script(self):
        with open(SCRIPT, encoding="utf-8-sig") as fh:
            text = fh.read()
        batches = R.split_sql_batches(text)
        self.assertEqual(len(batches), len(re.findall(r"^GO\r?$", text, re.M)))
        self.assertFalse(any(re.search(r"^GO$", b, re.M) for b in batches))
        self.assertIn("CREATE TABLE [dbo].[Settings]", batches[1])
        self.assertEqual(R.split_sql_batches("A\r\nGO\r\nB\nC\ngo\n\n"), ["A\n", "B\nC\n"])

    def test_insert_batch(self):
        self.assertEqual(R.insert_batch_sql("Units", "[UnitID], [UnitName]", "(1, N'حبة'),\n(2, N'علبة')", True),
                         "SET IDENTITY_INSERT [dbo].[Units] ON;\nINSERT INTO [dbo].[Units] ([UnitID], [UnitName]) VALUES "
                         "(1, N'حبة'),\n(2, N'علبة');\nSET IDENTITY_INSERT [dbo].[Units] OFF;\n")

    def test_tsql_like_of_the_test_server(self):
        self.assertTrue(R.tsql_like("3[0-9][0-9]3", "3123"))
        self.assertFalse(R.tsql_like("3[0-9][0-9]3", "31a3"))
        self.assertTrue(R.tsql_like("a%", "ABC"))


class TransferTests(unittest.TestCase):
    """The move into the server database (its rules, defaults and relationships, on SQLite)."""

    def check(self, source):
        server = R.server_database()
        copied = R.transfer(source, server)
        self.assertEqual(server.execute("PRAGMA foreign_key_check").fetchall(), [])
        self.assertEqual([t for t, ok in R.compare(source, server) if not ok], [])
        return server, copied

    def test_fixture(self):
        db = AccessOnSqlite()
        db.load_fixture()
        _, copied = self.check(db.con)
        self.assertEqual(sum(copied.values()),
                         sum(db.con.execute(f'SELECT COUNT(*) FROM "{t.name}"').fetchone()[0] for t in TABLES))

    def test_demo_data(self):
        r = DD.simulate()
        server, copied = self.check(r.db.con)
        self.assertGreater(copied["JournalLines"], 50)
        self.assertEqual(server.execute("SELECT StoreName FROM Settings").fetchone(),
                         r.db.con.execute("SELECT StoreName FROM Settings").fetchone(), "Arabic as it was")

    def test_the_comparison_finds_differences(self):
        r = DD.simulate()
        server = R.server_database()
        R.transfer(r.db.con, server)
        server.execute("UPDATE Customers SET CustomerName = '?????' WHERE CustomerID = (SELECT Max(CustomerID) FROM Customers)")
        server.execute("UPDATE Products SET SellingPrice = SellingPrice + 0.01 WHERE ProductID = (SELECT Min(ProductID) FROM Products)")
        bad = {t for t, ok in R.compare(r.db.con, server) if not ok}
        self.assertEqual(bad, {"Customers", "Products"})

    def test_rules_of_the_server_refuse_bad_data(self):
        db = AccessOnSqlite()
        db.load_fixture()
        db.con.execute("UPDATE Products SET SellingPrice = -1 WHERE ProductID = (SELECT Min(ProductID) FROM Products)")
        with self.assertRaises(Exception):
            R.transfer(db.con, R.server_database())


class CodeTests(unittest.TestCase):

    def test_registered(self):
        self.assertIn("modSqlServer", GEN.STATIC_MODULES)
        self.assertIn('"TestSqlServer"', read("modTestAll"))
        self.assertIn("Public Function SchemaTableList() As String", read("modBuildSchema"))
        perms = {r[0]: r for r in table("Permissions").seed_rows}
        self.assertIn("SQL_SERVER", perms)
        self.assertEqual({r for r, k in table("RolePermissions").seed_rows if k == "SQL_SERVER"}, {1}, "admin only")
        self.assertIn(("frmSqlServer", "قاعدة البيانات على SQL Server", "النظام", "SQL_SERVER", False, False, False),
                      SCREEN_LIST)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmSqlServer"], "SQL_SERVER")
        backup = next(m for m in F.all_forms() if m.name == "frmBackup")
        self.assertIn("btnSqlServer", {c.name for c in backup.controls})

    def test_screen_controls_exist(self):
        m = next(m for m in F.all_forms() if m.name == "frmSqlServer")
        names = {c.name for c in m.controls}
        for pname in re.findall(r"^(?:Public|Private) (?:Sub|Function) (\w+)\(ByVal frm As Access\.Form", VBA, re.M):
            for ctl in set(re.findall(r"frm!(\w+)", proc(VBA, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")

    def test_the_move_in_order(self):
        body = proc(VBA, "TransferAllData")
        steps = [body.index(s) for s in ("IsLinkedToServer()", "ServerTablesMissing", "NOCHECK CONSTRAINT ALL",
                                         "DELETE FROM", "CopyTable(", "WITH CHECK CHECK CONSTRAINT ALL")]
        self.assertEqual(steps, sorted(steps))
        self.assertIn('HasPermission("SQL_SERVER")', body)
        copy = proc(VBA, "CopyTable")
        self.assertIn("TableHasIdentity", copy)
        self.assertIn("SqlLiteral(f.Value, f.Type)", copy)
        self.assertIn("InsertBatchSql(TableName, cols, rows, hasIdentity)", copy)

    def test_link_only_after_a_clean_comparison(self):
        body = proc(VBA, "LinkToServer")
        self.assertLess(body.index("CompareAllData(Connect, rows, diffs)"), body.index("RelinkTable"))
        self.assertIn('SaveSqlSetting "SqlAccessPath", AccessBackEndPath()', body)
        self.assertIn('"dbo." & names(i)', body)
        self.assertIn('";DATABASE=" & BackEndPath, names(i)', proc(VBA, "LinkToAccess"))

    def test_the_comparison_checks_texts_and_amounts(self):
        body = proc(VBA, "CompareTable")
        self.assertIn("AscW(Nz([", body)
        self.assertIn("UNICODE(SUBSTRING(ISNULL([", body)
        self.assertIn("If f.Type = dbCurrency Then", body)

    def test_login_checks_the_server_first(self):
        login = proc(read("modSecurityScreens"), "LoginLoad")
        self.assertLess(login.index("CheckServerAtStartup()"), login.index("SettingValue("))
        self.assertIn("If Not IsLinkedToServer() Then Exit Function", proc(VBA, "CheckServerAtStartup"))

    def test_no_password_is_kept_in_the_settings(self):
        self.assertNotIn('SaveSqlSetting "SqlPassword"', VBA)
        self.assertNotRegex(proc(VBA, "SaveScreenSettings"), r"txtPassword")


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class RuntimeTests(unittest.TestCase):
    """The VBA itself, run in LibreOffice: the same results as the Python mirror."""

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modSqlServer": VBA})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_literals(self):
        for v, kind, expected in LITERALS:
            if isinstance(v, (dt.datetime, str)) and kind == 8:
                continue                                # a Date cannot cross the UNO bridge as such
            with self.subTest(v=v, kind=kind):
                self.assertEqual(self.h.call("modSqlServer", "SqlLiteral", v, kind), expected)

    def test_connection(self):
        for args in [(".\\SQLEXPRESS", "RetailStore", "", "", "ODBC Driver 18 for SQL Server"),
                     ("srv", "db", "shop", "p1", "ODBC Driver 17 for SQL Server")]:
            self.assertEqual(self.h.call("modSqlServer", "ServerConnect", *args), R.server_connect(*args))

    def test_batches(self):
        with open(SCRIPT, encoding="utf-8-sig") as fh:
            text = fh.read()
        self.assertEqual(self.h.call("modSqlServer", "SqlBatchCount", text), len(R.split_sql_batches(text)))
        self.assertEqual(self.h.call("modSqlServer", "SqlBatchCount", "A\r\nGO\r\nB\nC\ngo\n\n"), 2)

    def test_insert_batch(self):
        args = ("Units", "[UnitID], [UnitName]", "(1, N'حبة')", True)
        self.assertEqual(self.h.call("modSqlServer", "InsertBatchSql", *args).replace("\r\n", "\n"),
                         R.insert_batch_sql(*args))


class Static_modSqlServer(VbaModuleChecks, unittest.TestCase):
    module_name = "modSqlServer"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
