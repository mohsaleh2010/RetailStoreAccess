"""Screen permissions per user, the programmer account and copy protection.

  * the screens list (schema.SCREEN_LIST -> table Screens) covers every screen of the side menu,
    agrees with the data screens (add / delete) and with the role permission of each screen;
  * every place that saves (data screens, documents, permission screens) asks CanScreenAction;
  * HasPermission / CanOpenScreen: programmer always, own screens for a user with CustomScreens;
  * the programmer is created by BuildSchema, hidden from the users screen, may log in once
    without a password, and alone decides if the administrator may rename the shop;
  * activation: machine id and codes are EXECUTED in LibreOffice and compared with
    tools/activation_reference.py; login and frmMain need an activated computer.
"""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
import activation_reference as A
import forms as F
import vba_harness as H
from schema import SCREEN_LIST, table

MODELS = {m.name: m for m in F.all_forms()}


def read(name, folder="src"):
    sub = os.path.join(ROOT, folder, "vba", name + ".bas")
    with open(sub, encoding="utf-8" if folder == "src" else "cp1256") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


class ScreenListTests(unittest.TestCase):

    def test_every_menu_screen_is_listed(self):
        listed = {s[0] for s in SCREEN_LIST}
        for item in F.NAV_ITEMS:
            if item.target:
                self.assertIn(item.target, listed, item.key)
        for name in listed:
            self.assertIn(name, MODELS, name)

    def test_role_permission_matches_the_screen_map(self):
        for name, title, module, key, *_ in SCREEN_LIST:
            self.assertEqual(F.SCREEN_PERMISSIONS[name] or None, key, name)

    def test_data_screens_actions(self):
        data = {d.name: d for d in F.DATA_SCREENS}
        for name, title, module, key, add, edit, delete in SCREEN_LIST:
            if name in data:
                d = data[name]
                self.assertEqual(add, d.allow_add, f"{name}: add")
                self.assertEqual(delete, d.allow_delete, f"{name}: delete")
                self.assertTrue(edit, f"{name}: a data screen edits its records")

    def test_seeded_and_added_to_existing_back_ends(self):
        t = table("Screens")
        self.assertTrue(t.seed_missing)
        self.assertEqual(len(t.seed_rows), len(SCREEN_LIST))
        self.assertEqual(len({r[0] for r in t.seed_rows}), len(SCREEN_LIST))


class EnforcementTests(unittest.TestCase):

    def test_documents_check_add_before_saving(self):
        saves = {"modPOS": ["SavePOS", "SaveReturn", "SavePayment"],
                 "modPurchaseScreens": ["SavePurchase", "SavePurchaseReturn", "SaveSupplierPayment",
                                        "PostInventoryMove", "PostCountScreen", "NewStockCount"],
                 "modCash": ["SaveCashVoucher", "SaveCashClosing"]}
        for module, names in saves.items():
            text = read(module)
            for name in names:
                self.assertIn('If Not CanScreenAction(frm.Name, "ADD") Then Exit', proc(text, name), name)

    def test_document_screens_with_add_are_covered(self):
        """A listed document screen with "add" has a save procedure that checks it."""
        data = {d.name for d in F.DATA_SCREENS}
        checked = set()
        for module in ("modPOS", "modPurchaseScreens", "modCash"):
            checked |= {m for m in re.findall(r"Public Function (\w+)\(ByVal frm", read(module))}
        doc_screens = [s for s in SCREEN_LIST if s[4] and s[0] not in data]
        self.assertEqual({s[0] for s in doc_screens},
                         {"frmPOS", "frmTouchPOS", "frmCafePOS", "frmSalesReturn", "frmCustomerPayment",
                          "frmPurchaseInvoice", "frmPurchaseReturn", "frmSupplierPayment", "frmInventory",
                          "frmStockCount", "frmCashVoucher", "frmCashClosing", "frmManualEntry",
                          "frmVatReturn", "frmAllocation", "frmBankTx", "frmBankRecon"})

    def test_data_screens_check_add_edit_delete(self):
        text = read("modForms")
        self.assertIn('If Not CanScreenAction(frm.Name, IIf(frm.NewRecord, "ADD", "EDIT")) Then Exit Function',
                      proc(text, "FormBeforeUpdate"))
        self.assertIn('If Not CanScreenAction(frm.Name, "DELETE") Then Exit Sub', proc(text, "DeleteRecord"))
        self.assertIn('CanScreenAction(frm.Name, "DELETE", True)', proc(text, "FormCurrent"))
        self.assertIn("LockForActions frm", proc(text, "FormCurrent"))
        self.assertIn('frm.AllowAdditions = False', proc(text, "FormLoad"))

    def test_permission_screens_check_edit(self):
        text = read("modSecurityScreens")
        self.assertIn('If Not CanScreenAction(frm.Name, "EDIT") Then Exit Function', proc(text, "SaveRolePermissions"))
        self.assertIn('If Not CanScreenAction(frm.Name, "EDIT") Then Exit Function', proc(text, "SaveUserScreens"))

    def test_rules(self):
        sec = read("modSecurity")
        has = proc(read("modCommon"), "HasPermission")
        self.assertLess(has.index("IsDeveloper"), has.index("CustomScreens"), "the programmer first")
        self.assertIn("u.CanOpen = True AND", has)
        self.assertIn("roleID <> ADMIN_ROLE_ID And", has, "an administrator is never limited")
        opener = proc(sec, "CanOpenScreen")
        self.assertLess(opener.index("IsDeveloper()"), opener.index("UsesCustomScreens()"))
        self.assertIn('If FormName = "frmActivation" Then', opener)
        custom = proc(sec, "UsesCustomScreens")
        self.assertIn("IsDeveloper = False AND RoleID <> ", custom)
        action = proc(sec, "CanScreenAction")
        self.assertIn('"SELECT CanOpen AND Can" & word', action)
        self.assertIn('"SELECT Has" & word', action)
        save = proc(read("modSecurityScreens"), "SaveUserScreensFor")
        for text in ('"لا يمكنك تغيير صلاحياتك بنفسك."', 'HasPermission("USERS")', "ws.BeginTrans",
                     "CanOpen AND CanAdd AND HasAdd"):
            self.assertIn(text, save)

    def test_grid_controls_exist(self):
        text = read("modSecurityScreens")
        names = {c.name for c in MODELS["frmUserScreens"].controls}
        for pname in ("UserScreensLoad", "UserScreensPicked", "UserScreensCustomChanged", "UserScreensFromRole",
                      "UserScreensAll", "SaveUserScreens"):
            for ctl in set(re.findall(r"frm!(\w+)", proc(text, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        body = proc(text, "ShowUserScreensState")
        for ctl in set(re.findall(r"frm!(\w+)", body)):
            self.assertIn(ctl, names, ctl)
        lines = {c.name for c in MODELS["frmUserScreenLines"].controls}
        self.assertLessEqual({"CanOpen", "CanAdd", "CanEdit", "CanDelete"}, lines)
        tmp = re.search(r'CREATE TABLE tmpUserScreens \((.*?)\)", dbFailOnError', read("modPOS"), re.S).group(1)
        for col in ("ScreenName", "ScreenTitle", "HasAdd", "CanOpen", "CanDelete", "ActionsNote"):
            self.assertIn(col, tmp)


class DeveloperTests(unittest.TestCase):

    def test_created_by_build_schema(self):
        text = read("modBuildSchema")
        self.assertIn("EnsureDeveloperUser", text)
        body = proc(text, "EnsureDeveloperUser")
        self.assertIn("[IsDeveloper] = True", body)
        self.assertIn("'developer', 1, 1, False, True, True, '\" & salt & \"', '\" & _", body)
        self.assertIn('Application.Run("PasswordHash", pwd, salt)', body)
        self.assertIn("If Len(pwd) = 0 Then", body, "cancel: no account rather than one without a password")

    def test_hidden_from_users_screen(self):
        text = read("modForms")
        self.assertIn('frm.RecordSource = "SELECT * FROM Employees WHERE IsDeveloper = False"', proc(text, "FormLoad"))
        self.assertIn('"t.IsDeveloper = False AND ("', proc(text, "RefreshList"))
        import forms_security as FS
        self.assertIn("e.IsDeveloper = False", FS.USER_ROWS)

    def test_never_without_a_password(self):
        body = proc(read("modSecurity"), "LoginUser")
        self.assertIn("If rs!EmployeeID = 1 And Len(Password) = 0 Then", body, "only the first admin login")

    def test_store_name_permission(self):
        sec = read("modSecurity")
        body = proc(sec, "CanChangeStoreName")
        self.assertIn("IsDeveloper()", body)
        self.assertIn('SettingValue("AllowAdminCompanyName")', body)
        forms_text = read("modForms")
        validate = proc(forms_text, "ValidateSettings")
        self.assertIn("Not CanChangeStoreName()", validate)
        self.assertIn("If Not IsDeveloper() And Nz(frm!AllowAdminCompanyName.Value", validate)
        load = proc(forms_text, "SettingsScreenLoad")
        self.assertIn("frm!AllowAdminCompanyName.Visible = IsDeveloper()", load)
        names = {c.name for c in MODELS["frmSettings"].controls}
        self.assertLessEqual({"StoreName", "StoreNameEn", "AllowAdminCompanyName", "lblStoreNameNote"}, names)
        self.assertEqual([f.default for f in table("Settings").fields if f.name == "AllowAdminCompanyName"],
                         ["False"])


DRIVER = r'''
Option VBASupport 1
Public Function RunMachine(ByVal s As String) As String
    RunMachine = MachineIDFrom(s)
End Function
Public Function RunCode(ByVal s As String) As String
    RunCode = ActivationCodeFor(s)
End Function
Public Function RunValid(ByVal m As String, ByVal c As String) As Boolean
    RunValid = IsValidCode(m, c)
End Function
'''


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class ActivationRuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modCommon": H.read_module("modCommon"), "modZatca": H.read_module("modZatca"),
                    "modSecurity": H.read_module("modSecurity"), "modActivation": H.read_module("modActivation"),
                    "Driver": DRIVER})
        cls.secret = A.vba_secret()

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_machine_id(self):
        for raw in ("", "To be filled by O.E.M.|BFEBFBFF000906EA|5A3C11F2", "ABC|def|1"):
            self.assertEqual(self.h.call("Driver", "RunMachine", raw), A.machine_id_from(raw))

    def test_codes(self):
        for raw in ("1", "2", "board|cpu|disk"):
            mid = A.machine_id_from(raw)
            code = A.activation_code(mid, self.secret)
            self.assertEqual(self.h.call("Driver", "RunCode", mid), code)
            self.assertTrue(self.h.call("Driver", "RunValid", mid, code.lower().replace("-", " ")))
            self.assertFalse(self.h.call("Driver", "RunValid", A.machine_id_from(raw + "x"), code))
            self.assertFalse(self.h.call("Driver", "RunValid", mid, code[:-1] + ("0" if code[-1] != "0" else "1")))


class ActivationTests(unittest.TestCase):

    def test_reference(self):
        mid = A.machine_id_from("board|cpu|disk")
        self.assertRegex(mid, r"^[0-9A-F]{4}(-[0-9A-F]{4}){3}$")
        code = A.activation_code(mid, "secret")
        self.assertRegex(code, r"^[0-9A-F]{5}(-[0-9A-F]{5}){3}$")
        self.assertNotEqual(code, A.activation_code(mid, "other secret"))
        self.assertTrue(A.is_valid(mid.lower(), code.replace("-", ""), "secret"))

    def test_login_and_main_need_an_activated_computer(self):
        text = read("modSecurityScreens")
        login = proc(text, "DoLogin")
        self.assertIn("If Not IsDeveloper() And Not IsActivated() Then", login)
        self.assertLess(login.index("IsActivated()"), login.index('DoCmd.OpenForm "frmMain"'))
        self.assertIn('DoCmd.OpenForm "frmActivation", acNormal, , , , acDialog, "LOGIN"', login)
        self.assertIn("(IsDeveloper() Or IsActivated())", proc(text, "MainOpen"))

    def test_only_administrators_activate_and_only_the_programmer_makes_codes(self):
        text = read("modActivation")
        self.assertIn("If Not IsAdministrator() Then", proc(text, "SaveActivation"))
        self.assertIn("If Not IsValidCode(ForMachineID, Code) Then", proc(text, "SaveActivation"))
        self.assertIn("If Not IsDeveloper() Then", proc(text, "GenerateActivationCode"))
        self.assertIn("dev = IsDeveloper()", proc(text, "ActivationLoad"))

    def test_screen_controls_exist(self):
        text = read("modActivation")
        names = {c.name for c in MODELS["frmActivation"].controls}
        for pname in re.findall(r"^Public Sub (\w+)\(ByVal frm As Access\.Form", text, re.M):
            for ctl in set(re.findall(r"frm!(\w+)", proc(text, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")

    def test_schema(self):
        t = table("Activations")
        self.assertIn(["MachineID"], [ix.fields for ix in t.indexes if ix.unique])
        self.assertIn("modActivation", read("modTestAll") + open(os.path.join(ROOT, "tools", "generate.py")).read())


class Static_modActivation(VbaModuleChecks, unittest.TestCase):
    module_name = "modActivation"
    vba = read("modActivation")


if __name__ == "__main__":
    unittest.main()
