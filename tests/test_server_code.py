"""Plan 5C (docs/70): the program runs on SQL Server.

  * every statement and updatable recordset on the program's tables carries dbSeeChanges (error 3622 without it);
  * no AutoNumber is read between AddNew and Update (SQL Server numbers the row only when it is saved);
  * no Yes/No field is compared with True / -1 in SQL (SQL Server keeps 1): "<> 0" / "= 0";
  * backup and restore of the server database (BACKUP DATABASE, RESTORE VERIFYONLY, RESTORE from master);
  * the real data on the server is never overwritten by a transfer, and a front-end on the old Access file is warned;
  * BuildSchema never relinks a server front-end; READ_COMMITTED_SNAPSHOT in the script.
Each check is shown to fail on the code before the change.
"""
import glob
import os
import re
import unittest

from helpers import ROOT
import generate as GEN
import server_code as S
import sql_transfer_reference as R
import vba_harness as H
from schema import table


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


def modules():
    for path in sorted(glob.glob(os.path.join(ROOT, "src", "vba", "*.bas"))):
        name = os.path.basename(path)[:-4]
        with open(path, encoding="utf-8") as fh:
            yield name, fh.read()


SQL = read("modSqlServer")
BACKUP = read("modBackup")
SCRIPT = os.path.join(ROOT, "dist", "sql", "RetailStore_Schema.sql")


class SeeChangesTests(unittest.TestCase):

    def test_the_transformation(self):
        cases = {
            'db.Execute "DELETE FROM T WHERE a = \'1,2\'"': 'db.Execute "DELETE FROM T WHERE a = \'1,2\'", dbSeeChanges',
            'CurrentDb.Execute "UPDATE T SET a = " & _\n    f(1, 2), dbFailOnError   \' x, y':
                'CurrentDb.Execute "UPDATE T SET a = " & _\n    f(1, 2), dbFailOnError + dbSeeChanges   \' x, y',
            'Set rs = db.OpenRecordset("SELECT a, b FROM T")':
                'Set rs = db.OpenRecordset("SELECT a, b FROM T", dbOpenDynaset, dbSeeChanges)',
            'Set rs = db.OpenRecordset("T", dbOpenDynaset)': 'Set rs = db.OpenRecordset("T", dbOpenDynaset, dbSeeChanges)',
            'Set rs = db.OpenRecordset("T", dbOpenDynaset, dbAppendOnly)':
                'Set rs = db.OpenRecordset("T", dbOpenDynaset, dbAppendOnly + dbSeeChanges)',
            'Set rs = db.OpenRecordset("SELECT 1", dbOpenSnapshot)': 'Set rs = db.OpenRecordset("SELECT 1", dbOpenSnapshot)',
            's = "db.Execute x"': 's = "db.Execute x"',
            "' db.Execute x": "' db.Execute x",
        }
        for before, after in cases.items():
            self.assertEqual(S.see_changes(before), after)
            self.assertEqual(S.see_changes(after), after, "applied twice = once")

    def test_every_module_on_the_tables(self):
        for name, text in modules():
            if name in S.BUILD_ON_FILE:
                continue
            with self.subTest(module=name):
                self.assertEqual(S.missing_see_changes(text), [])

    def test_no_option_on_the_temp_table_ddl(self):
        for name, text in modules():
            with self.subTest(module=name):
                self.assertEqual(S.ddl_with_see_changes(text), [])
        self.assertEqual(S.see_changes('db.Execute "CREATE TABLE tmpX (N LONG)", dbFailOnError'),
                         'db.Execute "CREATE TABLE tmpX (N LONG)", dbFailOnError')

    def test_the_check_finds_the_old_code(self):
        self.assertEqual(len(S.missing_see_changes('CurrentDb.Execute "DELETE FROM Units", dbFailOnError\n'
                                                   'Set rs = CurrentDb.OpenRecordset("AuditLog", dbOpenDynaset, dbAppendOnly)')), 2)

    def test_generated_modules_go_through_it(self):
        for name in ("modDemoData", "modTestSales", "modTestPurchases", "modTestSecurity", "modBuildForms"):
            self.assertIn("dbSeeChanges", read(name), name)
        self.assertNotIn("dbSeeChanges", read("modBuildSchema"), "the back-end file itself")

    def test_pass_through_querydefs_unchanged(self):
        self.assertIn("    qd.Execute dbFailOnError\n", proc(SQL, "ServerExec"))


class NewNumberTests(unittest.TestCase):

    def test_no_autonumber_read_before_update(self):
        autos = S.auto_fields()
        for name, text in modules():
            with self.subTest(module=name):
                self.assertEqual(S.ids_read_before_update(text, autos), [])

    def test_the_check_finds_the_old_audit_code(self):
        old = "    rs.AddNew\n    rs!LogDate = Now\n    logID = rs!LogID\n    rs.Update\n"
        self.assertEqual(S.ids_read_before_update(old, S.auto_fields()), ["3: logID = rs!LogID"])

    def test_audit_reads_the_number_after_saving(self):
        body = proc(read("modAudit"), "WriteAudit")
        self.assertLess(body.index("rs.Update"), body.index("rs.Bookmark = rs.LastModified"))
        self.assertLess(body.index("rs.Bookmark = rs.LastModified"), body.index("logID = rs!LogID"))


class BooleanTests(unittest.TestCase):

    def test_the_conversion(self):
        self.assertEqual(S.server_booleans('x = "SELECT * FROM T WHERE IsActive = True AND p.TrackStock=True"'),
                         'x = "SELECT * FROM T WHERE IsActive <> 0 AND p.TrackStock <> 0"')
        self.assertEqual(S.server_booleans('x = " AND [IsDeveloper] = False"'), 'x = " AND [IsDeveloper] = 0"')
        kept = 'x = "UPDATE Settings SET PricesIncludeVAT = True, IsActive = False WHERE SettingID = 1"'
        self.assertEqual(S.server_booleans(kept), kept, "an assignment is converted by the server")
        self.assertEqual(S.server_booleans("If rs!IsActive = True Then"), "If rs!IsActive = True Then", "VBA code")

    def test_no_comparison_with_true_in_sql(self):
        for name, text in modules():
            with self.subTest(module=name):
                self.assertEqual([text[s:e] for s, e, _ in S.bool_comparisons(text)], [])

    def test_saved_queries(self):
        import queries as Q
        names = "|".join(S.bool_fields())
        for q in Q.QUERIES:
            sql = q[1] if isinstance(q, tuple) else q.sql
            self.assertNotRegex(sql, rf"\b({names})\]?\s*=\s*(True|-1)\b", q)


class ServerBackupTests(unittest.TestCase):

    def test_mirror(self):
        c = R.server_connect("srv", "RetailStore", "", "", "ODBC Driver 18 for SQL Server")
        self.assertEqual(R.connect_database(c), "RetailStore")
        self.assertEqual(R.connect_database(R.connect_to_database(c, "master")), "master")
        self.assertEqual(R.connect_to_database(c, "master"), c.replace("DATABASE=RetailStore", "DATABASE=master"))
        self.assertEqual(R.sql_name("a]b"), "[a]]b]")
        self.assertEqual(R.backup_database_sql("RetailStore", "D:\\B\\it's.bak"),
                         "BACKUP DATABASE [RetailStore] TO DISK = N'D:\\B\\it''s.bak' "
                         "WITH COPY_ONLY, INIT, CHECKSUM, NAME = N'RetailStore';")
        restore = R.restore_database_sql("RetailStore", "x.bak")
        self.assertTrue(restore.startswith("ALTER DATABASE [RetailStore] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;"))
        self.assertTrue(restore.endswith("END CATCH;\nALTER DATABASE [RetailStore] SET MULTI_USER;"))
        self.assertIn("    THROW;", restore)

    def test_the_backup_goes_to_the_server(self):
        now = proc(BACKUP, "BackupNow")
        self.assertLess(now.index("If IsLinkedToServer() Then"), now.index("source = BackendFilePath()"))
        self.assertIn("ServerBackupNow(BackupFile, Folder, Suffix)", now)
        restore = proc(BACKUP, "RestoreBackup")
        self.assertLess(restore.index("ServerRestore(BackupFile)"), restore.index("target = BackendFilePath()"))
        srv = proc(BACKUP, "ServerBackupNow")
        self.assertLess(srv.index("BackupDatabaseSql(dbName, BackupFile)"), srv.index("BackupIsReadable(BackupFile)"))
        self.assertIn('".bak"', srv)
        self.assertIn("VerifyBackupSql(FilePath)", proc(BACKUP, "BackupIsReadable"))
        rst = proc(BACKUP, "ServerRestore")
        self.assertLess(rst.index('ServerBackupNow(safety, , "_BeforeRestore")'), rst.index("RestoreDatabaseSql("))
        self.assertIn('ConnectToDatabase(cn, "master")', rst)

    def test_the_screen_lists_the_server_backups(self):
        screens = read("modSecurityScreens")
        self.assertIn("ServerBackupRows()", proc(screens, "BackupRefreshList"))
        self.assertIn("file = frm!lstBackups.Value", proc(screens, "BackupRestoreSelected"))
        self.assertIn("DataLocation()", proc(screens, "BackupLoad"))


class LiveDataTests(unittest.TestCase):

    def test_the_field(self):
        f = next(f for f in table("Settings").fields if f.name == "ServerLiveSince")
        self.assertEqual((f.kind, f.required, f.default), ("DATETIME", False, None))

    def test_no_transfer_over_the_real_data(self):
        body = proc(SQL, "TransferAllData")
        self.assertLess(body.index("ServerLiveSince(Connect)"), body.index("NOCHECK CONSTRAINT ALL"))

    def test_link(self):
        body = proc(SQL, "LinkToServer")
        live = body.index("live = Not IsNull(ServerLiveSince(Connect))")
        self.assertLess(live, body.index("CompareAllData(Connect, rows, diffs)"))
        self.assertIn("If Not live Then", body)
        self.assertLess(body.index("SET ServerLiveSince = GETDATE()"), body.index("RelinkTable"))
        self.assertLess(body.index("UPDATE Settings SET ServerLiveSince = Now()"), body.index("RelinkTable"))
        back = proc(SQL, "LinkToAccess")
        self.assertIn("SET ServerLiveSince = NULL", back)
        self.assertIn("UPDATE Settings SET ServerLiveSince = Null", back)
        self.assertIn("msg = LinkToAccess(, release)", proc(SQL, "SqlServerBackToAccess"))

    def test_login_warns_on_the_old_access_file(self):
        login = proc(read("modSecurityScreens"), "LoginLoad")
        self.assertLess(login.index("CheckServerAtStartup()"), login.index("StaleAccessWarning()"))
        self.assertIn("If IsLinkedToServer() Then Exit Function", proc(SQL, "StaleAccessWarning"))

    def test_build_schema_never_relinks_a_server_front_end(self):
        body = proc(read("modBuildSchema"), "BuildSchema")
        self.assertLess(body.index("If TablesOnServer() Then"), body.index("LinkBackEnd BackEndPath"))

    def test_images_and_audit_captions_on_the_server(self):
        self.assertIn("If IsLinkedToServer() Then be = AccessBackEndPath()", proc(read("modTouchPOS"), "ImagesFolder"))
        self.assertIn('";DATABASE=" & AccessBackEndPath()', proc(read("modAudit"), "BackEndDatabase"))


class ScriptTests(unittest.TestCase):

    def test_read_committed_snapshot(self):
        with open(SCRIPT, encoding="utf-8-sig") as fh:
            text = fh.read()
        options = text.index("ALTER DATABASE CURRENT SET READ_COMMITTED_SNAPSHOT ON WITH ROLLBACK IMMEDIATE;")
        guard = text.rindex("is_read_committed_snapshot_on = 0", 0, options)
        self.assertLess(options - guard, 120, "only once: guarded")
        self.assertGreater(options, text.index("7. Relationships"))
        self.assertIn("ServerLiveSince", text)


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class RuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modSqlServer": SQL})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_same_as_the_mirror(self):
        c = R.server_connect("srv", "RetailStore", "u", "p", "ODBC Driver 17 for SQL Server")
        self.assertEqual(self.h.call("modSqlServer", "ConnectDatabase", c), R.connect_database(c))
        self.assertEqual(self.h.call("modSqlServer", "ConnectToDatabase", c, "master"), R.connect_to_database(c, "master"))
        self.assertEqual(self.h.call("modSqlServer", "SqlName", "a]b"), R.sql_name("a]b"))
        self.assertEqual(self.h.call("modSqlServer", "BackupDatabaseSql", "RetailStore", "D:\\x.bak"),
                         R.backup_database_sql("RetailStore", "D:\\x.bak"))
        self.assertEqual(self.h.call("modSqlServer", "VerifyBackupSql", "x.bak"), R.verify_backup_sql("x.bak"))
        self.assertEqual(self.h.call("modSqlServer", "RestoreDatabaseSql", "RetailStore", "x.bak"),
                         R.restore_database_sql("RetailStore", "x.bak"))


if __name__ == "__main__":
    unittest.main()
