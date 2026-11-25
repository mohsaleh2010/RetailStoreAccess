"""Delivery (docs/88, phase 7D): the release secret and package, the ACCDE script, the web deployment by FTP.

  * tools/release.py: a secret made once, never printed, only in the copy of modActivation in release/ (Windows-1256,
    CRLF, nothing else changed); the activation codes of the release differ from the development ones;
  * dist/tools/MakeAccde.vbs: SysCmd 603, refuses an open file, ASCII + CRLF;
  * tools/deploy_web.py: the production settings, the files sent (never App_Data nor the development files),
    app_offline.htm first and removed last (kept after a failure), the T-SQL batches;
  * the site: keys kept in App_Data, GET /health; the SQL script does not fail on a shared host;
  * release/ and appsettings.Production.json are never committed.
"""
import io
import json
import os
import re
import shutil
import sys
import tempfile
import unittest
from contextlib import redirect_stdout

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
        with self.assertRaises(OSError):
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
        text = open(path, encoding="utf-8").read()
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
