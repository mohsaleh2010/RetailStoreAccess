"""Delivery (docs/88, phase 7D): the release secret and package, the ACCDE script, the web deployment by FTP.

  * tools/release.py: a secret made once, never printed, only in the copy of modActivation in release/ (Windows-1256,
    CRLF, nothing else changed); the activation codes of the release differ from the development ones;
  * dist/tools/MakeAccde.vbs: SysCmd 603, refuses an open file, ASCII + CRLF;
  * tools/deploy_web.py: the production settings, the files sent (never App_Data nor the development files),
    app_offline.htm first and removed last (kept after a failure), the T-SQL batches;
  * the site: keys kept in App_Data, GET /health; the SQL script does not fail on a shared host;
  * release/ and appsettings.Production.json are never committed.
"""
import ftplib
import io
import json
import os
import re
import shutil
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from unittest import mock

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
sys.path.insert(0, os.path.join(ROOT, "tools"))

import activation_reference  # noqa: E402
import deploy_web  # noqa: E402
import release  # noqa: E402


def read(*parts, encoding="utf-8"):
    with open(os.path.join(ROOT, *parts), encoding=encoding) as fh:
        return fh.read()


class ReleaseSecretTests(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmp)
        self.secret_file = os.path.join(self.tmp, "release", "license-secret.txt")

    def test_made_once_long_and_simple(self):
        self.assertTrue(release.make_secret(self.secret_file))
        first = release.read_secret(self.secret_file)
        self.assertFalse(release.make_secret(self.secret_file), "never replaced: the codes given would stop working")
        self.assertEqual(first, release.read_secret(self.secret_file))
        self.assertEqual(len(first), release.SECRET_LENGTH)
        self.assertRegex(first, r"^[A-Za-z0-9]+$")
        self.assertEqual(release.secret_problem(first, activation_reference.vba_secret()), "")
        self.assertNotEqual(release.new_secret(), release.new_secret())

    def test_problems(self):
        dev = activation_reference.vba_secret()
        self.assertIn("32", release.secret_problem("short", dev))
        self.assertIn("quote", release.secret_problem('a' * 40 + '"', dev))
        self.assertIn("quote", release.secret_problem('a' * 40 + ' b', dev))
        if len(dev) >= release.MIN_LENGTH and re.fullmatch(r"[A-Za-z0-9_\-.+*/!@#$%^()=~]+", dev):
            self.assertIn("development", release.secret_problem(dev, dev))

    def test_missing_secret_file(self):
        with self.assertRaises(release.ReleaseError):
            release.read_secret(self.secret_file)


class ReleasePackageTests(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmp)
        self.secret = release.new_secret()
        self.package = os.path.join(self.tmp, "RetailStore")

    def build(self):
        return release.build_package(self.secret, package=self.package)

    def test_only_the_secret_line_changes(self):
        self.build()
        with open(os.path.join(ROOT, "dist", "vba", release.MODULE), "rb") as fh:
            original = fh.read()
        with open(os.path.join(self.package, "vba", release.MODULE), "rb") as fh:
            made = fh.read()
        self.assertEqual(release.secret_in(made), self.secret)
        old_lines, new_lines = original.split(b"\r\n"), made.split(b"\r\n")
        self.assertEqual(len(old_lines), len(new_lines), "CRLF kept")
        changed = [i for i, (a, b) in enumerate(zip(old_lines, new_lines)) if a != b]
        self.assertEqual(len(changed), 1)
        self.assertIn(b"Private Const LICENSE_SECRET As String", new_lines[changed[0]])
        made.decode("cp1256")

    def test_the_package_holds_everything_to_build(self):
        self.build()
        for folder in ("vba", "tools", "sql"):
            self.assertEqual(sorted(os.listdir(os.path.join(ROOT, "dist", folder))),
                             sorted(os.listdir(os.path.join(self.package, folder))), folder)
        self.assertTrue(os.path.exists(os.path.join(self.package, "tools", "MakeAccde.vbs")))
        self.assertTrue(os.path.exists(os.path.join(self.package, "88-Delivery-Installation.md")))
        for name in os.listdir(os.path.join(ROOT, "dist", "vba")):
            if name != release.MODULE:
                with open(os.path.join(ROOT, "dist", "vba", name), "rb") as a, \
                        open(os.path.join(self.package, "vba", name), "rb") as b:
                    self.assertEqual(a.read(), b.read(), name)

    def test_check(self):
        self.assertIn("not made", release.check_package(self.secret, self.package))
        self.build()
        self.assertEqual(release.check_package(self.secret, self.package), "")
        self.assertIn("another", release.check_package(release.new_secret(), self.package))
        shutil.copy2(os.path.join(ROOT, "dist", "vba", release.MODULE), os.path.join(self.package, "vba", release.MODULE))
        self.assertIn("development", release.check_package(self.secret, self.package))

    def test_rebuild_replaces_the_old_package(self):
        self.build()
        stale = os.path.join(self.package, "vba", "old.bas")
        open(stale, "w").close()
        self.build()
        self.assertFalse(os.path.exists(stale))

    def test_refuses_the_development_secret(self):
        with self.assertRaises(release.ReleaseError):
            release.build_package("x", package=self.package)

    def test_codes_of_the_release(self):
        machine = activation_reference.machine_id_from("BOARD-1|CPU-2|DISK-3")
        code = activation_reference.activation_code(machine, self.secret)
        self.assertTrue(activation_reference.is_valid(machine, code, self.secret))
        self.assertFalse(activation_reference.is_valid(machine, code, activation_reference.vba_secret()))


class ReleaseCommandTests(unittest.TestCase):
    """The commands never print the secret."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmp)
        self.saved = (release.SECRET_FILE, release.PACKAGE)
        release.SECRET_FILE = os.path.join(self.tmp, "license-secret.txt")
        release.PACKAGE = os.path.join(self.tmp, "RetailStore")
        self.addCleanup(self.restore)

    def restore(self):
        release.SECRET_FILE, release.PACKAGE = self.saved

    def run_main(self, *args):
        out = io.StringIO()
        with redirect_stdout(out):
            code = release.main(list(args))
        return code, out.getvalue()

    def test_secret_package_check_code(self):
        code, text = self.run_main("check")
        self.assertEqual(code, 2)
        self.assertIn("release.py secret", text)
        code, text = self.run_main("secret")
        self.assertEqual(code, 0)
        secret = release.read_secret()
        self.assertEqual(os.path.dirname(release.SECRET_FILE), self.tmp)
        self.assertIn("not made", self.run_main("check")[1])
        code, text = self.run_main("package")
        self.assertEqual(code, 0)
        code, text2 = self.run_main("check")
        self.assertEqual(code, 0)
        code, text3 = self.run_main("code", "ABCD-1234-ABCD-1234", "bad")
        self.assertIn(activation_reference.activation_code("ABCD-1234-ABCD-1234", secret), text3)
        self.assertIn("bad: not a machine id", text3)
        code, text4 = self.run_main("secret")
        self.assertIn("kept", text4)
        self.assertEqual(release.read_secret(), secret)
        for out in (text, text2, text3, text4):
            self.assertNotIn(secret, out)

    def test_usage(self):
        code, text = self.run_main()
        self.assertEqual(code, 1)
        self.assertIn("release.py secret", text)

    def test_source_has_no_secret(self):
        text = read("tools", "release.py")
        self.assertNotIn(activation_reference.vba_secret(), text)
        self.assertNotIn("print(secret", text)
        self.assertNotIn("{secret}", text)


class MakeAccdeTests(unittest.TestCase):

    def setUp(self):
        with open(os.path.join(ROOT, "dist", "tools", "MakeAccde.vbs"), "rb") as fh:
            self.data = fh.read()
        self.text = self.data.decode("ascii")

    def test_crlf_ascii(self):
        self.assertNotIn(b"\n", self.data.replace(b"\r\n", b""))

    def test_makes_the_accde(self):
        self.assertIn("Const acSysCmdMakeAccde = 603", self.text)
        self.assertIn("app.SysCmd acSysCmdMakeAccde, src, dst", self.text)
        self.assertIn(".laccdb", self.text, "an open file is refused")
        self.assertIn("app.Quit 2", self.text)
        self.assertIn("Option Explicit", self.text)
        self.assertNotIn(activation_reference.vba_secret(), self.text)


class FakeFtp:
    def __init__(self, fail_at=None):
        self.calls, self.files, self.fail_at = [], {}, fail_at

    def storbinary(self, command, fh):
        path = command[5:]
        if self.fail_at and path.endswith(self.fail_at):
            raise OSError("connection lost")
        self.files[path] = fh.read()
        self.calls.append(("STOR", path))

    def mkd(self, path):
        self.calls.append(("MKD", path))

    def voidcmd(self, command):
        return "200 OK"

    def size(self, path):
        return len(self.files[path])

    def delete(self, path):
        self.files.pop(path, None)
        self.calls.append(("DELE", path))


SETTINGS = {"sql_server": "sql1.host.net,1433", "sql_database": "shop", "sql_user": "u1", "sql_password": "p=w#1",
            "ftp_host": "ftp.host.net", "ftp_user": "f1", "ftp_dir": "/site/wwwroot", "ftp_tls": "yes",
            "site_url": "https://erp.example.com", "api_docs": "no", "jwt_key": "k" * 64}


class DeploySettingsTests(unittest.TestCase):

    def test_problems(self):
        self.assertEqual(deploy_web.settings_problem(SETTINGS), "")
        self.assertIn("https", deploy_web.settings_problem({**SETTINGS, "site_url": "http://erp.example.com"}))
        self.assertIn("missing", deploy_web.settings_problem({**SETTINGS, "ftp_user": ""}))
        self.assertIn("';'", deploy_web.settings_problem({**SETTINGS, "sql_password": "a;b"}))
        self.assertIn("token", deploy_web.settings_problem({**SETTINGS, "jwt_key": "short"}))

    def test_production_settings(self):
        p = deploy_web.production_settings(SETTINGS)
        self.assertEqual(p["Database"]["Provider"], "SqlServer")
        self.assertEqual(p["Api"]["JwtKey"], "k" * 64)
        self.assertFalse(p["Api"]["Docs"])
        self.assertFalse(p["Hosting"]["ForwardedHeaders"])
        cs = p["ConnectionStrings"]["SqlServer"]
        self.assertTrue(cs.startswith("Server=sql1.host.net,1433;Database=shop;User Id=u1;Password=p=w#1;Encrypt=True;"))
        self.assertIn("TrustServerCertificate=True", cs)
        # the keys the site reads (Program.cs, ApiSecurity.cs)
        program = read("web", "src", "RetailStore.Web", "Program.cs")
        for key in ('"Database:Provider"', 'GetConnectionString("SqlServer")', '"Hosting:ForwardedHeaders"'):
            self.assertIn(key, program)
        self.assertIn('config["Api:JwtKey"]', read("web", "src", "RetailStore.Web", "Api", "ApiSecurity.cs"))

    def test_jwt_key(self):
        key = deploy_web.new_jwt_key()
        self.assertEqual(len(key), 64)
        self.assertRegex(key, r"^[A-Za-z0-9]+$")

    def test_settings_file_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "release", "web-deploy.json")
            with self.assertRaises(deploy_web.DeployError):
                deploy_web.load_settings(path)
            deploy_web.save_settings(SETTINGS, path)
            self.assertEqual(deploy_web.load_settings(path), SETTINGS)

    def test_publish_command(self):
        self.assertEqual(deploy_web.publish_command("out", False)[-2:], ["--self-contained", "true"])
        command = deploy_web.publish_command("out", True)
        self.assertEqual(command[-2:], ["--self-contained", "false"])
        self.assertIn("win-x64", command)
        self.assertTrue(os.path.exists(deploy_web.PROJECT))


class DeployUploadTests(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmp)
        files = ["RetailStore.Web.dll", "RetailStore.Web.exe", "RetailStore.Web.pdb", "web.config", "appsettings.json",
                 "appsettings.Development.json", "appsettings.Production.json", "wwwroot/app.css", "wwwroot/lib/x.js",
                 "App_Data/keys/key-1.xml", "App_Data/retailstore-dev.db", "ar/RetailStore.UI.resources.dll"]
        for rel in files:
            path = os.path.join(self.tmp, *rel.split("/"))
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w") as fh:
                fh.write(rel)

    def test_plan(self):
        rels = [rel for rel, _ in deploy_web.upload_plan(self.tmp)]
        self.assertIn("wwwroot/lib/x.js", rels)
        self.assertIn("appsettings.Production.json", rels)
        self.assertIn("ar/RetailStore.UI.resources.dll", rels)
        for never in ("App_Data/keys/key-1.xml", "App_Data/retailstore-dev.db", "appsettings.Development.json",
                      "RetailStore.Web.pdb"):
            self.assertNotIn(never, rels)

    def test_offline_first_removed_last(self):
        ftp = FakeFtp()
        sent = deploy_web.upload(ftp, SETTINGS, self.tmp, log=lambda _t: None)
        self.assertEqual(sent, 8)
        self.assertEqual(ftp.calls[0], ("STOR", "/site/wwwroot/app_offline.htm"))
        self.assertEqual(ftp.calls[-1], ("DELE", "/site/wwwroot/app_offline.htm"))
        self.assertIn(("MKD", "/site/wwwroot/wwwroot/lib"), ftp.calls)
        self.assertLess(ftp.calls.index(("MKD", "/site/wwwroot/wwwroot")), ftp.calls.index(("MKD", "/site/wwwroot/wwwroot/lib")))
        self.assertNotIn("/site/wwwroot/app_offline.htm", ftp.files)
        self.assertEqual(ftp.files["/site/wwwroot/wwwroot/app.css"], b"wwwroot/app.css")

    def test_failure_keeps_the_site_offline(self):
        ftp = FakeFtp(fail_at="web.config")
        with self.assertRaises(deploy_web.DeployError):
            deploy_web.upload(ftp, SETTINGS, self.tmp, log=lambda _t: None)
        self.assertIn("/site/wwwroot/app_offline.htm", ftp.files, "half a program is never started")

    def test_refuses_an_unpublished_folder(self):
        os.remove(os.path.join(self.tmp, "appsettings.Production.json"))
        with self.assertRaises(deploy_web.DeployError):
            deploy_web.upload(FakeFtp(), SETTINGS, self.tmp)

    def test_stdout_log(self):
        config = ('<?xml version="1.0" encoding="utf-8"?><configuration><system.webServer><handlers>'
                  '<add name="aspNetCore" path="*" verb="*" modules="AspNetCoreModuleV2" resourceType="Unspecified" />'
                  '</handlers><aspNetCore processPath=".\\RetailStore.Web.exe" stdoutLogEnabled="false" '
                  'stdoutLogFile=".\\logs\\stdout" hostingModel="inprocess" /></system.webServer></configuration>')
        path = os.path.join(self.tmp, "web.config")
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(config)
        deploy_web.set_stdout_log(self.tmp, True)
        with open(path, encoding="utf-8") as fh:
            text = fh.read()
        self.assertIn('stdoutLogEnabled="true"', text)
        self.assertIn('stdoutLogFile=".\\App_Data\\logs\\stdout"', text)
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(config.replace("inprocess", "outofprocess"))
        with self.assertRaises(deploy_web.DeployError):
            deploy_web.set_stdout_log(self.tmp, False)

    def test_app_offline_page(self):
        self.assertIn('dir="rtl"', deploy_web.APP_OFFLINE)
        self.assertEqual(deploy_web.APP_OFFLINE.encode("ascii").decode("ascii"), deploy_web.APP_OFFLINE)

    def test_health_url(self):
        self.assertEqual(deploy_web.health_url({"site_url": "https://erp.example.com/"}), "https://erp.example.com/health")


class FakeCursor:
    def __init__(self, log):
        self.log = log

    def execute(self, sql):
        self.log.append(sql)


class FakeConnection:
    def __init__(self):
        self.batches, self.commits = [], 0

    def cursor(self):
        return FakeCursor(self.batches)

    def commit(self):
        self.commits += 1


class DeploySqlTests(unittest.TestCase):

    def test_split(self):
        self.assertEqual(deploy_web.split_batches("A\nGO\n\n go \nB\nC\nGO"), ["A", "B\nC"])
        self.assertEqual(deploy_web.split_batches("SELECT 'GOOD'\n"), ["SELECT 'GOOD'"])

    def test_both_scripts(self):
        connection = FakeConnection()
        files = [os.path.join(ROOT, "dist", "sql", name) for name in ("RetailStore_Schema.sql", "RetailStore_Queries.sql")]
        count = deploy_web.run_scripts(connection, files, log=lambda _t: None)
        self.assertEqual(count, connection.commits)
        self.assertGreater(count, 100)
        self.assertFalse(any(re.match(r"(?im)^\s*GO\s*$", b) for b in connection.batches))


class SharedHostingTests(unittest.TestCase):

    def test_read_committed_snapshot_never_stops_the_script(self):
        text = read("dist", "sql", "RetailStore_Schema.sql", encoding="utf-8-sig")
        options = text.index("ALTER DATABASE CURRENT SET READ_COMMITTED_SNAPSHOT ON")
        self.assertLess(text.rindex("BEGIN TRY", 0, options), options)
        self.assertLess(options, text.index("END TRY", options))
        catch = text[text.index("BEGIN CATCH", options):text.index("END CATCH", options)]
        self.assertIn("PRINT", catch)

    def test_site_keeps_its_keys_and_answers_health(self):
        program = read("web", "src", "RetailStore.Web", "Program.cs")
        self.assertIn("PersistKeysToFileSystem", program)
        self.assertIn('Path.Combine(builder.Environment.ContentRootPath, "App_Data", "keys")', program)
        self.assertIn("app.MapHealth();", program)
        health = read("web", "src", "RetailStore.Web", "Health.cs")
        self.assertIn('app.MapGet("/health"', health)
        self.assertIn(".AllowAnonymous()", health)
        self.assertNotIn("ex.Message", health, "the health answer tells nothing about the server")

    def test_repository_settings_hold_no_secret(self):
        settings = json.loads(read("web", "src", "RetailStore.Web", "appsettings.json"))
        self.assertEqual(settings["Api"]["JwtKey"], "")
        self.assertEqual(settings["ConnectionStrings"]["SqlServer"], "")

    def test_gitignore(self):
        ignore = read(".gitignore").splitlines()
        self.assertIn("release/", ignore)
        self.assertIn("appsettings.Production.json", ignore)
        self.assertIn("App_Data/", read("web", ".gitignore").splitlines())
        self.assertIn("release", deploy_web.SETTINGS_FILE.replace(os.sep, "/"))
        self.assertTrue(release.SECRET_FILE.replace(os.sep, "/").endswith("release/license-secret.txt"))


class WindowsPathTests(unittest.TestCase):
    """On Windows the repository and the temporary folder can be on two drives (E: and C:): os.path.relpath then
    raises ValueError. The tools print such a path whole; no file builds a path by replacing '/' in __file__."""

    def test_shown_on_another_drive(self):
        for module in (release, deploy_web):
            with mock.patch("os.path.relpath", side_effect=ValueError("path is on mount 'C:', start on mount 'E:'")):
                self.assertEqual(module.shown(os.path.join(ROOT, "release", "x")),
                                 os.path.abspath(os.path.join(ROOT, "release", "x")))
            self.assertEqual(module.shown(os.path.join(ROOT, "release", "x")), os.path.join("release", "x"))

    def test_commands_on_another_drive(self):
        with tempfile.TemporaryDirectory() as tmp:
            saved = (release.SECRET_FILE, release.PACKAGE)
            release.SECRET_FILE = os.path.join(tmp, "license-secret.txt")
            release.PACKAGE = os.path.join(tmp, "RetailStore")
            try:
                release.make_secret()
                out = io.StringIO()
                with mock.patch("os.path.relpath", side_effect=ValueError("another drive")), redirect_stdout(out):
                    self.assertEqual(release.main(["package"]), 0)
                self.assertIn(release.PACKAGE, out.getvalue())
                with mock.patch("os.path.relpath", side_effect=ValueError("another drive")):
                    with self.assertRaises(deploy_web.DeployError) as caught:
                        deploy_web.load_settings(os.path.join(tmp, "web-deploy.json"))
                self.assertIn(tmp, str(caught.exception))
            finally:
                release.SECRET_FILE, release.PACKAGE = saved

    def test_no_path_built_with_a_slash_replace(self):
        pattern = re.compile(r"__file__\.replace\(")
        for folder in ("tools", "tests"):
            for name in os.listdir(os.path.join(ROOT, folder)):
                if name.endswith(".py") and name != "test_delivery.py":
                    with open(os.path.join(ROOT, folder, name), encoding="utf-8") as fh:
                        self.assertIsNone(pattern.search(fh.read()), f"{folder}/{name}: use os.path.join")


class InitAnswersTests(unittest.TestCase):
    """init asks a wrong answer again instead of stopping at the end and losing every answer (owner's PC)."""

    def run_init(self, answers, secrets_=("pw",), old=None):
        typed, hidden, said = list(answers), list(secrets_), []
        settings = deploy_web.ask_settings(old or {}, ask=lambda _q: typed.pop(0),
                                           ask_secret=lambda _q: hidden.pop(0), log=said.append)
        self.assertEqual(typed, [], "every answer used")
        return settings, said

    BASE = ["sql1.host.net", "shop", "u1", "ftp.host.net", "f1", "/"]

    def test_site_without_https_is_completed(self):
        settings, said = self.run_init(self.BASE + ["y", "store.example.net/", "y"])
        self.assertEqual(settings["site_url"], "https://store.example.net")
        self.assertEqual(said, [])
        self.assertEqual(deploy_web.settings_problem(settings), "")

    def test_wrong_answers_asked_again(self):
        settings, said = self.run_init(self.BASE + ["x", "y", "http://store.example.net", "https://store.example.net", "no"])
        self.assertEqual(settings["ftp_tls"], "yes")
        self.assertEqual(settings["api_docs"], "no")
        self.assertEqual(len(said), 2)
        self.assertIn("yes or no", said[0])
        self.assertIn("https://", said[1])

    def test_enter_keeps_the_old_values_and_the_key(self):
        old = {"sql_server": "s", "sql_database": "d", "sql_user": "u", "sql_password": "old", "ftp_host": "f",
               "ftp_user": "fu", "ftp_dir": "/", "ftp_tls": "yes", "site_url": "https://a.example.net",
               "api_docs": "no", "jwt_key": "k" * 64}
        settings, _ = self.run_init([""] * 9, secrets_=("",), old=old)
        self.assertEqual(settings, old)
        settings, _ = self.run_init([""] * 9, secrets_=("new",), old=old)
        self.assertEqual(settings["sql_password"], "new")
        self.assertEqual(settings["jwt_key"], "k" * 64)

    def test_normalize(self):
        self.assertEqual(deploy_web.normalize_site_url("HTTPS://Erp.Example.com/"), ("HTTPS://Erp.Example.com", ""))
        self.assertTrue(deploy_web.normalize_site_url("http://a.b")[1])
        self.assertTrue(deploy_web.normalize_site_url("")[1])
        self.assertTrue(deploy_web.normalize_site_url("a b")[1])
        self.assertEqual(deploy_web.yes_no_value("نعم"), "yes")
        self.assertIsNone(deploy_web.yes_no_value("x"))
        # the keys y and n of an Arabic keyboard
        self.assertEqual(deploy_web.yes_no_value("غ"), "yes")
        self.assertEqual(deploy_web.yes_no_value("ى"), "no")

    def test_sql_confirmation_is_never_a_silent_no(self):
        """The owner answered غ (the y key) and the command ended without a word."""
        for answers, runs in ((["غ"], True), (["?", "y"], True), (["n"], False)):
            typed, said = list(answers), []
            with mock.patch.object(deploy_web, "sql_connect", return_value=FakeSqlConnection()) as connect, \
                    mock.patch.object(deploy_web, "run_scripts") as run:
                deploy_web.cmd_sql(mock.Mock(yes=False), SETTINGS, ask=lambda _q: typed.pop(0), log=said.append)
            self.assertEqual(typed, [])
            self.assertEqual(run.called, runs, answers)
            self.assertEqual(connect.called, runs)
            if not runs:
                self.assertIn("Nothing was run.", said)
            if answers[0] == "?":
                self.assertIn("  answer yes or no (y / n)", said)

    def test_ctrl_c_is_not_a_traceback(self):
        out = io.StringIO()
        with mock.patch.object(deploy_web, "cmd_init", side_effect=KeyboardInterrupt), redirect_stdout(out):
            self.assertEqual(deploy_web.main(["init"]), 130)
        self.assertIn("Cancelled", out.getvalue())


LOGIN_FAILED = ("(18456, b\"Login failed for user 'shop_user'.DB-Lib error message 20018, severity 14:\\n"
                "General SQL Server error: Check messages from the SQL Server\\nDB-Lib error message 20002, severity 9:"
                "\\nTDS server connection failed (sql1.host.net)\\n\")")


class FakeSqlConnection:
    def __init__(self, row=("shop", "16.0.4135.4", 1), fail_on=None):
        self.row, self.fail_on, self.closed, self.rolled_back, self.batches = row, fail_on, False, 0, []

    def cursor(self):
        connection = self

        class Cursor:
            def execute(self, sql):
                if connection.fail_on and connection.fail_on in sql:
                    raise RuntimeError("(2714, b'There is already an object named X in the database.')")
                connection.batches.append(sql)

            def fetchone(self):
                return connection.row
        return Cursor()

    def commit(self):
        pass

    def rollback(self):
        self.rolled_back += 1

    def close(self):
        self.closed = True


class SqlErrorTests(unittest.TestCase):
    """deploy_web.py sql on the owner's PC: error 18456 came out as a traceback of pymssql."""

    def refused(self, text):
        def connect(**_kw):
            raise RuntimeError(text)
        return connect

    def test_login_failed_explained(self):
        with self.assertRaises(deploy_web.DeployError) as caught:
            deploy_web.sql_connect(SETTINGS, self.refused(LOGIN_FAILED))
        message = str(caught.exception)
        self.assertIn("18456", message)
        self.assertIn("password", message)
        self.assertNotIn("p=w#1", message, "the password is never shown")

    def test_other_errors(self):
        self.assertEqual(deploy_web.sql_error_help(RuntimeError("(4060, b'Cannot open database shop')")),
                         deploy_web.SQL_DATABASE_HELP)
        self.assertIn("port", deploy_web.sql_error_help(RuntimeError("(20009, b'Unable to connect')")))

    def test_connect_arguments(self):
        seen = {}

        def connect(**kw):
            seen.update(kw)
            return FakeSqlConnection()
        deploy_web.sql_connect({**SETTINGS, "sql_server": " sql1.host.net , 1533 "}, connect)
        self.assertEqual((seen["server"], seen["port"], seen["database"]), ("sql1.host.net", "1533", "shop"))
        self.assertTrue(seen["autocommit"], "ALTER DATABASE is refused inside a transaction (errors 226 / 3998)")
        deploy_web.sql_connect({**SETTINGS, "sql_server": "sql1.host.net"}, connect)
        self.assertEqual(seen["port"], "1433")

    def test_sql_check_all_permissions(self):
        said = []
        ones = (1,) * len(deploy_web.PERMISSION_CHECKS)
        connection = FakeSqlConnection(row=("shop", "16.0.1220.3") + ones)
        self.assertTrue(deploy_web.cmd_sql_check(None, SETTINGS, connect=lambda **_kw: connection, log=said.append))
        self.assertIn("Connected to shop", said[0])
        self.assertIn("Next:", said[1])
        self.assertEqual(len(said), 2)
        self.assertTrue(connection.closed)
        query = connection.batches[0]
        self.assertNotIn("IS_ROLEMEMBER", query, "a user with the needed permissions but not db_owner is fine")
        self.assertEqual(query.count("HAS_PERMS_BY_NAME"), len(deploy_web.PERMISSION_CHECKS) + 1)

    def test_sql_check_read_write_and_create_table_only(self):
        """The owner's host: read / write and CREATE TABLE, not db_owner: only what is really missing is named."""
        said = []
        values = {"CREATE TABLE": 1, "ALTER on dbo": 1, "REFERENCES on dbo": 1, "CREATE VIEW": 0, "schema [en]": None,
                  "SELECT": 1, "INSERT": 1, "UPDATE": 1, "DELETE": 1}
        row = ("shop", "16.0") + tuple(values[label] for label, _e, _w in deploy_web.PERMISSION_CHECKS)
        connection = FakeSqlConnection(row=row)
        self.assertFalse(deploy_web.cmd_sql_check(None, SETTINGS, connect=lambda **_kw: connection, log=said.append))
        missing = [line for line in said if line.startswith("  - ")]
        self.assertEqual([m.split(":")[0] for m in missing], ["  - CREATE VIEW", "  - schema [en]"])
        self.assertIn("db_ddladmin", said[1])

    def test_permissions_cover_the_scripts(self):
        schema = read("dist", "sql", "RetailStore_Schema.sql", encoding="utf-8-sig")
        views = read("dist", "sql", "RetailStore_Queries.sql", encoding="utf-8-sig")
        labels = [label for label, _e, _w in deploy_web.PERMISSION_CHECKS]
        if "CREATE TABLE" in schema:
            self.assertIn("CREATE TABLE", labels)
        if "ALTER TABLE" in schema or "CREATE INDEX" in schema:
            self.assertIn("ALTER on dbo", labels)
        if "FOREIGN KEY" in schema:
            self.assertIn("REFERENCES on dbo", labels)
        if "CREATE OR ALTER VIEW" in views:
            self.assertIn("CREATE VIEW", labels)
        if "CREATE SCHEMA [en]" in views:
            self.assertIn("schema [en]", labels)

    def test_main_prints_no_traceback(self):
        out = io.StringIO()
        with mock.patch.object(deploy_web, "load_settings", return_value=SETTINGS), \
                mock.patch.object(deploy_web, "sql_connect", side_effect=deploy_web.DeployError(deploy_web.SQL_LOGIN_HELP)), \
                redirect_stdout(out):
            self.assertEqual(deploy_web.main(["sql-check"]), 2)
            self.assertEqual(deploy_web.main(["sql", "--yes"]), 2)
        self.assertIn("Error: SQL Server refused the login", out.getvalue())

    def test_failed_part_named(self):
        connection = FakeSqlConnection(fail_on="CREATE VIEW")
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "x.sql")
            with open(path, "w", encoding="utf-8") as fh:
                fh.write("CREATE TABLE A (x INT)\nGO\n-- the view\nCREATE VIEW V AS SELECT 1 AS x\nGO\n")
            with self.assertRaises(deploy_web.DeployError) as caught:
                deploy_web.run_scripts(connection, [path], log=lambda _t: None)
        message = str(caught.exception)
        self.assertIn("x.sql, part 2 of 2 (CREATE VIEW V", message)
        self.assertEqual(deploy_web.first_statement("/*----\n  8. Database options\n----*/\nIF EXISTS (SELECT 1)"),
                         "IF EXISTS (SELECT 1)")
        self.assertIn("2714", message)
        self.assertEqual(connection.rolled_back, 1)


class FtpsTests(unittest.TestCase):
    """The owner's host: TimeoutError in ssl unwrap() after the first STOR of ftplib.FTP_TLS."""

    class Conn:
        def __init__(self, log, unwrap_error=TimeoutError("The read operation timed out")):
            self.log, self.sent, self.unwrap_error = log, b"", unwrap_error

        def sendall(self, data):
            self.sent += data

        def settimeout(self, seconds):
            self.log.append(("timeout", seconds))

        def unwrap(self):
            self.log.append("unwrap")
            if self.unwrap_error:
                raise self.unwrap_error
            return self

        def shutdown(self, how):
            self.log.append("shutdown")

        def close(self):
            self.log.append("close")

    def store(self, unwrap_error):
        ftp = deploy_web.SafeFtpTls()
        log = []
        conn = self.Conn(log, unwrap_error)
        ftp.voidcmd = lambda c: log.append(c)
        ftp.transfercmd = lambda c, rest=None: (log.append(c), conn)[1]
        ftp.voidresp = lambda: (log.append("226"), "226 Transfer complete")[1]
        answer = ftp.storbinary("STOR /a.txt", io.BytesIO(b"x" * 200000))
        return answer, conn, log

    def test_store_survives_a_server_without_tls_close(self):
        for error in (TimeoutError("The read operation timed out"), None):
            answer, conn, log = self.store(error)
            self.assertEqual(answer, "226 Transfer complete")
            self.assertEqual(len(conn.sent), 200000)
            self.assertEqual(log, ["TYPE I", "STOR /a.txt", ("timeout", deploy_web.SafeFtpTls.UNWRAP_SECONDS),
                                   "unwrap", "shutdown", "226", "close"],
                             "the socket is closed only after the 226 (closing at once cut the file)")

    def test_size_checked(self):
        ftp = mock.Mock()
        ftp.size.return_value = 10
        deploy_web.check_size(ftp, "/a", 10)
        with self.assertRaises(ftplib.error_reply):
            deploy_web.check_size(ftp, "/a", 11)
        ftp.size.side_effect = ftplib.error_perm("500 SIZE not understood")
        deploy_web.check_size(ftp, "/a", 11)

    def test_data_connection_reuses_the_tls_session(self):
        ftp = deploy_web.SafeFtpTls()
        ftp._prot_p = True
        ftp.host = "ftp.host.net"
        ftp.sock = mock.Mock(session="SESSION")
        ftp.context = mock.Mock()
        ftp.context.wrap_socket.return_value = "wrapped"
        with mock.patch("ftplib.FTP.ntransfercmd", return_value=("raw", 10)):
            self.assertEqual(ftp.ntransfercmd("STOR x"), ("wrapped", 10))
        ftp.context.wrap_socket.assert_called_once_with("raw", server_hostname="ftp.host.net", session="SESSION")

    def test_ftps_by_default(self):
        with mock.patch.object(deploy_web.SafeFtpTls, "connect"), mock.patch.object(deploy_web.SafeFtpTls, "login"), \
                mock.patch.object(deploy_web.SafeFtpTls, "prot_p") as prot:
            ftp = deploy_web.ftp_connect(SETTINGS, "pw")
        self.assertIsInstance(ftp, deploy_web.SafeFtpTls)
        self.assertTrue(prot.called)

    def test_errors_explained(self):
        with tempfile.TemporaryDirectory() as tmp:
            for name in ("RetailStore.Web.dll", "appsettings.Production.json", "web.config"):
                with open(os.path.join(tmp, name), "w") as fh:
                    fh.write(name)

            class Failing(FakeFtp):
                def storbinary(self, command, fh):
                    raise TimeoutError("The read operation timed out")
            with self.assertRaises(deploy_web.DeployError) as caught:
                deploy_web.upload(Failing(), SETTINGS, tmp, log=lambda _t: None)
            self.assertIn("app_offline.htm", str(caught.exception))
            self.assertIn("plain FTP", str(caught.exception))
            with self.assertRaises(deploy_web.DeployError) as caught:
                deploy_web.upload(FakeFtp(fail_at="web.config"), SETTINGS, tmp, log=lambda _t: None)
            self.assertIn("stopped at web.config (2 of 3 files sent)", str(caught.exception))
            self.assertIn("upload  again", str(caught.exception))


def _ftps_server_available():
    try:
        import pyftpdlib  # noqa: F401
        import OpenSSL  # noqa: F401
    except ImportError:
        return False
    return shutil.which("openssl") is not None


@unittest.skipUnless(_ftps_server_available(), "pip install pyftpdlib pyopenssl (and openssl on the PATH)")
class FtpsServerTests(unittest.TestCase):
    """A real FTPS server (pyftpdlib, TLS on the control and data connections), also one that never sends its TLS
    close like the owner's host: every byte arrives, App_Data is never sent, app_offline.htm is removed."""

    @classmethod
    def setUpClass(cls):
        import subprocess
        cls.tmp = tempfile.mkdtemp()
        key, crt = os.path.join(cls.tmp, "k.pem"), os.path.join(cls.tmp, "c.pem")
        subprocess.run(["openssl", "req", "-x509", "-newkey", "rsa:2048", "-nodes", "-keyout", key, "-out", crt,
                        "-days", "1", "-subj", "/CN=localhost"], check=True, capture_output=True)
        cls.pem = os.path.join(cls.tmp, "both.pem")
        with open(cls.pem, "w") as out:
            out.write(open(key).read() + open(crt).read())

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp)

    def serve(self, no_tls_close):
        import threading
        import OpenSSL.SSL
        from pyftpdlib.authorizers import DummyAuthorizer
        from pyftpdlib.handlers import TLS_FTPHandler
        from pyftpdlib.servers import FTPServer
        if no_tls_close:
            patcher = mock.patch.object(OpenSSL.SSL.Connection, "shutdown", lambda _self: True)
            patcher.start()
            self.addCleanup(patcher.stop)
        root = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, root)
        os.makedirs(os.path.join(root, "site"))
        auth = DummyAuthorizer()
        auth.add_user("u", "p", root, perm="elradfmwMT")

        class Handler(TLS_FTPHandler):
            pass
        Handler.certfile, Handler.authorizer = self.pem, auth
        Handler.tls_control_required = Handler.tls_data_required = True
        Handler.log_prefix = ""
        server = FTPServer(("127.0.0.1", 0), Handler)
        thread = threading.Thread(target=server.serve_forever, kwargs={"timeout": 0.2}, daemon=True)
        thread.start()
        self.addCleanup(server.close_all)
        return root, server.address[1]

    def publish(self):
        pub = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, pub)
        files = {"RetailStore.Web.dll": 3_000_000, "appsettings.Production.json": 200, "wwwroot/css/app.css": 5000,
                 "App_Data/keys/key.xml": 10, "web.config": 400}
        for rel, size in files.items():
            path = os.path.join(pub, *rel.split("/"))
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "wb") as fh:
                fh.write(os.urandom(size))
        return pub, files

    def check(self, no_tls_close):
        import logging
        logging.getLogger("pyftpdlib").setLevel(logging.ERROR)
        root, port = self.serve(no_tls_close)
        pub, files = self.publish()
        ftp = deploy_web.SafeFtpTls(timeout=20)
        ftp.connect("127.0.0.1", port)
        ftp.login("u", "p")
        ftp.prot_p()
        try:
            sent = deploy_web.upload(ftp, {"ftp_dir": "/site"}, pub, log=lambda _t: None)
        finally:
            ftp.quit()
        site = os.path.join(root, "site")
        self.assertEqual(sent, 4)
        for rel, size in files.items():
            path = os.path.join(site, *rel.split("/"))
            if rel.startswith("App_Data"):
                self.assertFalse(os.path.exists(path))
            else:
                with open(path, "rb") as a, open(os.path.join(pub, *rel.split("/")), "rb") as b:
                    self.assertEqual(a.read(), b.read(), rel)
        self.assertFalse(os.path.exists(os.path.join(site, "app_offline.htm")))

    def test_upload(self):
        self.check(no_tls_close=False)

    def test_upload_to_a_server_without_tls_close(self):
        self.check(no_tls_close=True)


class DocsTests(unittest.TestCase):

    def test_guide_and_rows(self):
        guide = read("docs", "88-Delivery-Installation.md")
        for word in ("deploy_web.py", "release.py", "MakeAccde.vbs", "/health", "app_offline.htm", "BuildFrontEnd.vbs",
                     "RetailStore_Schema.sql", "RetailStore_Queries.sql", "RetailStore.OfficeService"):
            self.assertIn(word, guide)
        self.assertIn("docs/88-Delivery-Installation.md", read("README.md"))
        self.assertNotIn(activation_reference.vba_secret(), guide)


if __name__ == "__main__":
    unittest.main()
