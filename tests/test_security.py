"""Phase 10 checks: users, passwords, permissions and backup.

  * SHA-256 and the password hash are EXECUTED in LibreOffice and compared with
    Python's hashlib (tools/security_reference.py), as are the backup file
    names and the retention rule;
  * every screen has a permission entry and every permission exists;
  * the login code keeps the security rules (generic message, lock-out,
    no stored passwords, transaction-safe permission checks);
  * static checks on modSecurity, modSecurityScreens, modBackup, modTestSecurity.

Run:  python3 -m unittest discover -s tests -v
"""

import datetime as dt
import os
import random
import re
import unittest

from helpers import ROOT, VbaModuleChecks
import forms as F
import gen_test_security
import security_reference as S
import vba_harness as H
from schema import table

MODELS = F.all_forms()


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


DRIVER = r'''
Option VBASupport 1
Public Function RunSha(ByVal s As String) As String
    RunSha = Sha256Text(s)
End Function
Public Function RunPwd(ByVal p As String, ByVal salt As String) As String
    RunPwd = PasswordHash(p, salt)
End Function
Public Function RunName(ByVal baseName As String, ByVal d As Double) As String
    RunName = BackupFileName(baseName, CDate(d))
End Function
Public Function RunOld(ByVal names As String, ByVal keep As Long) As String
    RunOld = OldBackups(names, keep)
End Function
'''


def ole_date(d: dt.datetime) -> float:
    delta = d - dt.datetime(1899, 12, 30)
    return delta.days + delta.seconds / 86400


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class SecurityRuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modCommon": H.read_module("modCommon"), "modZatca": H.read_module("modZatca"),
                    "modSecurity": H.read_module("modSecurity"), "modBackup": H.read_module("modBackup"),
                    "Driver": DRIVER})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_sha256_vectors(self):
        for text in S.SHA_VECTORS:
            with self.subTest(len(text)):
                self.assertEqual(self.h.call("Driver", "RunSha", text), S.sha256_hex(text.encode()))

    def test_sha256_random_lengths(self):
        """Every padding case (lengths around 55/56/64/119/120 bytes) and random Arabic text."""
        rnd = random.Random(10)
        alphabet = "abcXYZ0123 ؟،ابتثجحخدذرزسشصضطظعغفقكلمنهوي"
        for n in list(range(50, 70)) + [119, 120, 121, 127, 128, 129] + [rnd.randint(0, 300) for _ in range(10)]:
            text = "".join(rnd.choice(alphabet) for _ in range(n))
            self.assertEqual(self.h.call("Driver", "RunSha", text), S.sha256_hex(text.encode()), n)

    def test_password_hash(self):
        for password, salt in S.PASSWORD_VECTORS:
            self.assertEqual(self.h.call("Driver", "RunPwd", password, salt), S.password_hash(password, salt))

    def test_backup_names_and_retention(self):
        when = dt.datetime(2026, 1, 9, 23, 59, 58)
        self.assertEqual(self.h.call("Driver", "RunName", "RetailStore_BE", ole_date(when)),
                         S.backup_file_name("RetailStore_BE", when))
        rnd = random.Random(3)
        names = [S.backup_file_name("B", dt.datetime(2026, 1, 1) + dt.timedelta(minutes=rnd.randint(0, 99999)))
                 for _ in range(12)]
        for keep in (1, 3, 12, 20):
            got = self.h.call("Driver", "RunOld", "|".join(names), keep)
            self.assertEqual(got.split("|") if got else [], S.backups_to_delete(names, keep), keep)


class PermissionMapTests(unittest.TestCase):

    def test_every_screen_has_a_permission_entry(self):
        subforms = {c.props["SourceObject"] for m in MODELS for c in m.controls if c.kind == "subform"}
        screens = {m.name for m in MODELS} - subforms
        self.assertEqual(screens, set(F.SCREEN_PERMISSIONS), "SCREEN_PERMISSIONS must list every screen")

    def test_permission_keys_exist(self):
        keys = {r[0] for r in table("Permissions").seed_rows}
        for form, key in F.SCREEN_PERMISSIONS.items():
            if key:
                self.assertIn(key, keys, form)
        for name in ("modSecurity", "modSecurityScreens", "modBackup", "modDashboard", "modScreens", "modReports"):
            for key in re.findall(r'HasPermission\("(\w+)"\)', read(name)):
                self.assertIn(key, keys, f"{name}: {key}")

    def test_admin_screens_closed_to_cashier(self):
        cashier = {p for r, p in table("RolePermissions").seed_rows if r == 3}
        for form in ("frmUsers", "frmRoles", "frmBackup", "frmSettings", "frmPurchaseInvoice", "frmReportCenter"):
            self.assertNotIn(F.SCREEN_PERMISSIONS[form], cashier, form)
        self.assertIn(F.SCREEN_PERMISSIONS["frmPOS"], cashier)

    def test_profit_reports_need_the_profit_permission(self):
        flagged = {r.key for r in F.REPORTS if "$" in r.needs}
        self.assertTrue({"PROFIT", "VAT_SUMMARY", "MONTHLY_SALES", "SALES_PRODUCT", "BEST_SELLING"} <= flagged)

    def test_nav_buttons_carry_their_target(self):
        main = next(m for m in MODELS if m.name == "frmMain")
        for item in F.NAV_ITEMS:
            btn = next(c for c in main.controls if c.name == f"btnNav{item.key}")
            self.assertEqual(btn.props.get("Tag", ""), item.target, item.key)


class SecurityCodeTests(unittest.TestCase):
    sec = read("modSecurity")

    def test_constants_match_reference(self):
        for name, value in (("PASSWORD_ITERATIONS", S.ITERATIONS), ("MIN_PASSWORD_LENGTH", S.MIN_PASSWORD_LENGTH),
                            ("MAX_FAILED_LOGINS", S.MAX_FAILED_LOGINS), ("LOCK_MINUTES", S.LOCK_MINUTES)):
            self.assertIn(f"Public Const {name} As Long = {value}", self.sec)

    def test_login_does_not_reveal_user_names(self):
        body = proc(self.sec, "LoginUser")
        generic = '"اسم المستخدم أو كلمة المرور غير صحيحة."'
        self.assertGreaterEqual(body.count(generic), 2, "unknown user and wrong password: same message")

    def test_passwords_are_never_stored(self):
        for name in ("modSecurity", "modSecurityScreens"):
            text = read(name)
            self.assertNotRegex(text, r"rs!PasswordHash = (?!PasswordHash\()", name)
        self.assertIn("rs!PasswordHash = PasswordHash(NewPassword, salt)", proc(self.sec, "SetUserPassword"))

    def test_no_default_user(self):
        body = proc(read("modCommon"), "CurrentUserID")
        self.assertNotIn("Else CurrentUserID = 1", body)
        self.assertNotIn('TempVars.Add "UserID", 1', read("modStartup"))

    def test_permission_check_sees_open_transactions(self):
        body = proc(read("modCommon"), "HasPermission")
        self.assertNotIn("DLookup", body)
        self.assertNotIn("DCount", body)

    def test_every_in_access_test_runs_as_a_user(self):
        import gen_forms
        import gen_reports
        import gen_test_purchases
        import gen_test_sales
        for text in (gen_forms.build_forms_vba(), gen_reports.build_reports_vba(),
                     gen_test_sales.build_test_sales_vba(), gen_test_purchases.build_test_purchases_vba(),
                     gen_test_security.build_test_security_vba(), read("modDashboard")):
            self.assertIn("EnsureTestUser", text)

    def test_user_mode_blocks_shift_and_starts_on_login(self):
        text = read("modStartup")
        self.assertIn('SetDbProp db, "AllowBypassKey", dbBoolean, Developer', text)
        self.assertIn('SetDbProp db, "StartupForm", dbText, "frmLogin"', text)
        with open(os.path.join(ROOT, "dist", "tools", "EnableShiftKey.vbs"), "rb") as fh:
            data = fh.read()
        data.decode("ascii")
        self.assertIn(b'"AllowBypassKey"', data)

    def test_restore_makes_a_safety_copy_first(self):
        body = proc(read("modBackup"), "RestoreBackup")
        self.assertLess(body.index('BackupNow(safety, , "_BeforeRestore")'), body.index("fso.CopyFile BackupFile"))
        self.assertIn("BackupIsReadable(BackupFile)", body)


def _static(name, text=None):
    class T(VbaModuleChecks, unittest.TestCase):
        module_name = name
        vba = text if text is not None else read(name)
    T.__name__ = T.__qualname__ = f"Static_{name}"
    return T


Static_modSecurity = _static("modSecurity")
Static_modSecurityScreens = _static("modSecurityScreens")
Static_modBackup = _static("modBackup")
Static_modTestSecurity = _static("modTestSecurity", gen_test_security.build_test_security_vba())
