"""Phase 10 screens: login, change password, roles and permissions, backup.
The users screen is a regular data screen (forms.DATA_SCREENS, frmUsers)."""

from typing import List, Tuple

from forms import Control, FormModel, Sym, button, cm, labelled, title_band
from forms_sales import LOCKED, grid_row, header_labels

ROLE_ROWS = "SELECT RoleID, RoleName FROM Roles ORDER BY RoleID"
PASSWORD = {"InputMask": "Password"}


def check_with_label(m: FormModel, name: str, caption: str, x: int, y: int, w: int, events=()):
    m.add(Control("check", name, x, y + cm(0.15), cm(0.5), cm(0.5), {}, events=list(events)))
    m.add(Control("label", "lbl" + name[3:], x + cm(0.65), y, w - cm(0.65), cm(0.8),
                  {"Caption": caption, "FontSize": 10, "ForeColor": Sym("CLR_TEXT")}, parent=name))


def layout_login() -> FormModel:
    width, height = cm(16.0), cm(9.6)
    m = FormModel("frmLogin", "تسجيل الدخول", width, height, popup=True, allow_add=False)
    title_band(m, "تسجيل الدخول", "نظام إدارة المحل", "users")
    m.add(Control("label", "lblStoreName", cm(0.4), cm(1.75), width - cm(0.8), cm(0.75),
                  {"Caption": " ", "FontSize": 13, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    c = m.add(Control("text", "txtUsername", cm(0.4), cm(3.1), width - cm(0.8), cm(0.9), {"FontSize": 13},
                      events=["KeyDown"]))
    labelled(m, "txtUsername", "اسم المستخدم", c)
    c = m.add(Control("text", "txtPassword", cm(0.4), cm(4.6), width - cm(0.8), cm(0.9),
                      {"FontSize": 13, **PASSWORD}, events=["KeyDown"]))
    labelled(m, "txtPassword", "كلمة المرور", c)
    m.add(Control("label", "lblMessage", cm(0.4), cm(5.65), width - cm(0.8), cm(1.2),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_DANGER")}))
    button(m, "btnLogin", "دخول", cm(0.4), cm(7.4), "primary", w=cm(5.4), h=cm(1.0), call="DoLogin Me")
    button(m, "btnExit", "خروج", width - cm(5.8), cm(7.4), "secondary", w=cm(5.4), h=cm(1.0), call="LoginExit")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    LoginLoad Me", "End Sub",
               "Private Sub txtUsername_KeyDown(KeyCode As Integer, Shift As Integer)",
               "    LoginKeyDown Me, KeyCode, False", "End Sub",
               "Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)",
               "    LoginKeyDown Me, KeyCode, True", "End Sub"] + m.code)
    return m


def layout_change_password() -> FormModel:
    width, height = cm(16.0), cm(10.4)
    m = FormModel("frmChangePassword", "كلمة المرور", width, height, popup=True, allow_add=False)
    title_band(m, "كلمة المرور", "تغيير أو تعيين كلمة المرور", "users")
    m.add(Control("label", "lblFor", cm(0.4), cm(1.75), width - cm(0.8), cm(0.65),
                  {"Caption": " ", "FontSize": 11, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    y = cm(3.05)
    for name, caption in [("txtOld", "كلمة المرور الحالية"), ("txtNew", "كلمة المرور الجديدة"),
                          ("txtConfirm", "تأكيد كلمة المرور الجديدة")]:
        c = m.add(Control("text", name, cm(0.4), y, width - cm(0.8), cm(0.85), {"FontSize": 12, **PASSWORD}))
        labelled(m, name, caption, c)
        y += cm(1.45)
    check_with_label(m, "chkMustChange", "يغيّرها المستخدم عند أول دخول (كلمة مؤقتة)", cm(0.4), cm(7.25),
                     width - cm(0.8))
    m.add(Control("label", "lblRules", cm(0.4), cm(8.1), width - cm(0.8), cm(0.55),
                  {"Caption": "6 أحرف على الأقل، ولا تساوي اسم المستخدم", "FontSize": 9,
                   "ForeColor": Sym("CLR_MUTED")}))
    button(m, "btnSave", "حفظ", cm(0.4), cm(9.0), "primary", w=cm(5.4), h=cm(1.0),
           call="SaveChangedPassword Me")
    button(m, "btnCancel", "إلغاء", width - cm(5.8), cm(9.0), "secondary", w=cm(5.4), h=cm(1.0),
           call="CancelChangePassword Me")
    m.form_events = ["Load"]
    m.code = ["Private Sub Form_Load()", "    ChangePasswordLoad Me", "End Sub"] + m.code
    return m


ROLE_TITLES = ["ممنوحة", "الصلاحية", "القسم"]


def layout_role_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.7)
    m = FormModel("frmRolePermLines", "صلاحيات الدور", cm(14.8), row_h, popup=False,
                  record_source="SELECT * FROM tmpRolePermissions ORDER BY SortOrder", allow_add=False)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0}
    heads = grid_row(m, [
        ("Granted", "Granted", 1.6, {"_kind": "check"}, []),
        ("PermissionName", "PermissionName", 9.5, dict(LOCKED), []),
        ("ModuleName", "ModuleName", 3.4, dict(LOCKED), []),
    ], row_h)
    m.code = []
    return m, heads


def layout_roles(heads) -> FormModel:
    width, height = cm(16.0), cm(16.0)
    m = FormModel("frmRoles", "الأدوار والصلاحيات", width, height, popup=True, allow_add=False)
    title_band(m, "الأدوار والصلاحيات", "حدد ما يستطيع كل دور فعله", "users")
    c = m.add(Control("combo", "cboRole", cm(0.4), cm(2.3), cm(7.0), cm(0.8),
                      {"RowSource": ROLE_ROWS, "ColumnCount": 2, "ColumnWidths": "0;6"}, events=["AfterUpdate"]))
    labelled(m, "cboRole", "الدور", c)
    m.add(Control("label", "lblRoleInfo", cm(7.7), cm(2.35), width - cm(8.1), cm(0.7),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    header_labels(m, cm(0.4), cm(3.45), ROLE_TITLES, heads)
    m.add(Control("subform", "subPermissions", cm(0.4), cm(4.05), cm(14.8), cm(9.6),
                  {"SourceObject": "frmRolePermLines"}))
    m.add(Control("label", "lblLockedNote", cm(0.4), cm(13.8), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 9, "FontBold": True, "ForeColor": Sym("CLR_WARNING")}))
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnSaveRole", "حفظ الصلاحيات", "primary", 3.8, "SaveRolePermissions Me"),
            ("btnAll", "تحديد الكل", "secondary", 3.0, "RoleSelectAll Me, True"),
            ("btnNone", "إلغاء الكل", "secondary", 3.0, "RoleSelectAll Me, False")]:
        button(m, name, caption, bx, cm(14.6), style, w=cm(w), h=cm(1.0), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(14.6), "secondary", w=cm(2.6), h=cm(1.0),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    RolesLoad Me", "End Sub",
               "Private Sub cboRole_AfterUpdate()", "    RolePicked Me", "End Sub"] + m.code)
    return m


# ------------------------------------------------------------- screens of a user
USER_ROWS = ("SELECT e.EmployeeID, e.EmployeeName & '  (' & e.Username & ')', r.RoleName FROM Employees AS e "
             "INNER JOIN Roles AS r ON e.RoleID = r.RoleID WHERE e.IsDeveloper = False ORDER BY e.EmployeeName")
USER_SCREEN_TITLES = ["فتح", "الشاشة", "القسم", "إضافة / حفظ", "تعديل", "حذف", "ما ينطبق عليها"]
USER_SCREEN_CHECKS = ("CanOpen", "CanAdd", "CanEdit", "CanDelete")


def layout_user_screen_lines() -> Tuple[FormModel, list]:
    row_h = cm(0.7)
    m = FormModel("frmUserScreenLines", "شاشات المستخدم", cm(19.6), row_h, popup=False,
                  record_source="SELECT * FROM tmpUserScreens ORDER BY SortOrder", allow_add=False)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0}
    heads = grid_row(m, [
        ("CanOpen", "CanOpen", 1.3, {"_kind": "check"}, ["AfterUpdate"]),
        ("ScreenTitle", "ScreenTitle", 5.6, dict(LOCKED), []),
        ("ModuleName", "ModuleName", 2.4, dict(LOCKED), []),
        ("CanAdd", "CanAdd", 1.9, {"_kind": "check"}, ["AfterUpdate"]),
        ("CanEdit", "CanEdit", 1.4, {"_kind": "check"}, ["AfterUpdate"]),
        ("CanDelete", "CanDelete", 1.4, {"_kind": "check"}, ["AfterUpdate"]),
        ("ActionsNote", "ActionsNote", 5.25, {**LOCKED, "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}, []),
    ], row_h)
    m.code = []
    for f in USER_SCREEN_CHECKS:
        m.code += [f"Private Sub {f}_AfterUpdate()", f'    UserScreenLineChanged Me, "{f}"', "End Sub"]
    return m, heads


def layout_user_screens(heads) -> FormModel:
    width, height = cm(20.4), cm(17.4)
    m = FormModel("frmUserScreens", "صلاحيات الشاشات", width, height, popup=True, allow_add=False)
    title_band(m, "صلاحيات الشاشات", "الشاشات التي يفتحها كل مستخدم، والإضافة والتعديل والحذف في كل شاشة",
               "users")
    c = m.add(Control("combo", "cboUser", cm(0.4), cm(2.3), cm(8.0), cm(0.8),
                      {"RowSource": USER_ROWS, "ColumnCount": 3, "ColumnWidths": "0;6;2.5"},
                      events=["AfterUpdate"]))
    labelled(m, "cboUser", "المستخدم", c)
    m.add(Control("label", "lblUserInfo", cm(8.7), cm(2.35), width - cm(9.1), cm(0.7),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("check", "chkCustom", cm(0.4), cm(3.45), cm(0.5), cm(0.5), {}, events=["AfterUpdate"]))
    m.add(Control("label", "lblCustom", cm(1.05), cm(3.3), width - cm(1.45), cm(0.8),
                  {"Caption": "صلاحيات شاشات خاصة بهذا المستخدم (بدل صلاحيات دوره)", "FontSize": 10,
                   "FontBold": True, "ForeColor": Sym("CLR_TEXT")}, parent="chkCustom"))
    header_labels(m, cm(0.4), cm(4.3), USER_SCREEN_TITLES, heads)
    m.add(Control("subform", "subScreens", cm(0.4), cm(4.9), cm(19.6), cm(9.6),
                  {"SourceObject": "frmUserScreenLines"}))
    m.add(Control("label", "lblNote", cm(0.4), cm(14.6), width - cm(0.8), cm(1.0),
                  {"Caption": " ", "FontSize": 9, "FontBold": True, "ForeColor": Sym("CLR_WARNING")}))
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnSaveScreens", "حفظ", "primary", 2.8, "SaveUserScreens Me"),
            ("btnFromRole", "من صلاحيات الدور", "secondary", 3.8, "UserScreensFromRole Me"),
            ("btnAll", "كل الشاشات", "secondary", 3.0, "UserScreensAll Me, True"),
            ("btnNone", "إلغاء الكل", "secondary", 2.8, "UserScreensAll Me, False")]:
        button(m, name, caption, bx, cm(15.9), style, w=cm(w), h=cm(1.0), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(15.9), "secondary", w=cm(2.6), h=cm(1.0),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    UserScreensLoad Me", "End Sub",
               "Private Sub cboUser_AfterUpdate()", "    UserScreensPicked Me", "End Sub",
               "Private Sub chkCustom_AfterUpdate()", "    UserScreensCustomChanged Me", "End Sub"] + m.code)
    return m


# ------------------------------------------------------------- activation
def layout_activation() -> FormModel:
    width, height = cm(19.0), cm(17.6)
    m = FormModel("frmActivation", "تفعيل البرنامج", width, height, popup=True, allow_add=False)
    title_band(m, "تفعيل البرنامج", "يعمل البرنامج على الأجهزة المفعّلة فقط بكود من المبرمج", "settings")
    m.add(Control("label", "lblState", cm(0.4), cm(1.85), width - cm(0.8), cm(0.8),
                  {"Caption": " ", "FontSize": 13, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    c = m.add(Control("text", "txtMachineID", cm(0.4), cm(3.4), cm(9.0), cm(0.9),
                      {"FontSize": 14, "FontBold": True, "Locked": True, "TextAlign": 2}))
    labelled(m, "txtMachineID", "رقم هذا الجهاز (أرسله للمبرمج)", c)
    button(m, "btnCopyID", "نسخ الرقم", cm(9.6), cm(3.4), "secondary", w=cm(3.0), h=cm(0.9),
           call="CopyMachineID Me")
    c = m.add(Control("text", "txtCode", cm(0.4), cm(5.0), cm(9.0), cm(0.9),
                      {"FontSize": 14, "TextAlign": 2}))
    labelled(m, "txtCode", "كود التفعيل", c)
    button(m, "btnActivate", "تفعيل هذا الجهاز", cm(9.6), cm(5.0), "primary", w=cm(4.4), h=cm(0.9),
           call="ActivateThisMachine Me")
    m.add(Control("label", "lblListCap", cm(0.4), cm(6.3), cm(12.0), cm(0.6),
                  {"Caption": "الأجهزة المفعّلة", "FontSize": 10, "FontBold": True,
                   "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstMachines", cm(0.4), cm(6.95), width - cm(0.8), cm(3.6),
                  {"RowSourceType": "Table/Query", "ColumnCount": 4, "ColumnWidths": "0;5;5.5;4",
                   "ColumnHeads": True,
                   "RowSource": "SELECT ActivationID, ComputerName AS [الجهاز], MachineID AS [رقم الجهاز], "
                                "ActivatedAt AS [تاريخ التفعيل] FROM Activations ORDER BY ActivatedAt"}))
    button(m, "btnRemove", "إلغاء تفعيل الجهاز المحدد", cm(0.4), cm(10.75), "danger", w=cm(5.6), h=cm(0.9),
           call="RemoveSelectedActivation Me")
    # the programmer only: codes for the computers of customers, and the shop name permission
    m.add(Control("rect", "boxDeveloper", cm(0.4), cm(11.95), width - cm(0.8), cm(3.75),
                  {"BackColor": Sym("CLR_SURFACE"), "BorderColor": Sym("CLR_BORDER")}, decorative=True))
    m.add(Control("label", "lblDevCap", cm(0.7), cm(12.1), width - cm(1.4), cm(0.6),
                  {"Caption": "للمبرمج: توليد كود تفعيل لجهاز عميل", "FontSize": 10, "FontBold": True,
                   "ForeColor": Sym("CLR_PRIMARY")}))
    c = m.add(Control("text", "txtForMachine", cm(0.7), cm(13.35), cm(7.4), cm(0.85),
                      {"FontSize": 12, "TextAlign": 2}))
    labelled(m, "txtForMachine", "رقم جهاز العميل", c)
    button(m, "btnGenerate", "توليد الكود", cm(8.3), cm(13.35), "primary", w=cm(3.0), h=cm(0.85),
           call="GenerateActivationCode Me")
    c = m.add(Control("text", "txtGenerated", cm(11.5), cm(13.35), cm(6.8), cm(0.85),
                      {"FontSize": 12, "FontBold": True, "Locked": True, "TextAlign": 2}))
    labelled(m, "txtGenerated", "كود التفعيل", c)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(16.2), "secondary", w=cm(2.6), h=cm(1.0),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = ["Private Sub Form_Load()", "    ActivationLoad Me", "End Sub"] + m.code
    return m


def layout_backup() -> FormModel:
    width, height = cm(20.0), cm(14.8)
    m = FormModel("frmBackup", "النسخ الاحتياطي", width, height, popup=True, allow_add=False)
    title_band(m, "النسخ الاحتياطي", "نسخة من ملف البيانات بالتاريخ والوقت، والاستعادة عند الحاجة", "backup")
    m.add(Control("label", "lblDataFile", cm(0.4), cm(1.8), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_TEXT")}))
    m.add(Control("label", "lblFolder", cm(0.4), cm(2.5), cm(13.5), cm(0.6),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_TEXT")}))
    button(m, "btnChooseFolder", "تغيير المجلد", cm(14.1), cm(2.42), "secondary", w=cm(2.6), h=cm(0.75),
           call="BackupChooseFolder Me")
    button(m, "btnOpenFolder", "فتح المجلد", cm(17.0), cm(2.42), "secondary", w=cm(2.6), h=cm(0.75),
           call="BackupOpenFolder")
    m.add(Control("label", "lblLast", cm(0.4), cm(3.3), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    button(m, "btnBackupNow", "نسخة احتياطية الآن", cm(0.4), cm(4.1), "primary", w=cm(5.4), h=cm(1.1),
           call="BackupRun Me")
    m.add(Control("label", "lblListCap", cm(0.4), cm(5.55), cm(12.0), cm(0.6),
                  {"Caption": "النسخ الموجودة (الأحدث أولًا)", "FontSize": 10, "FontBold": True,
                   "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstBackups", cm(0.4), cm(6.2), width - cm(0.8), cm(6.5),
                  {"RowSourceType": "Value List", "ColumnCount": 3, "ColumnWidths": "10;5;3.5",
                   "ColumnHeads": True}))
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnRestore", "استعادة النسخة المحددة", "danger", 5.4, "BackupRestoreSelected Me"),
            ("btnDevMode", "وضع المطوّر", "secondary", 3.6, "DeveloperModeFromApp")]:
        button(m, name, caption, bx, cm(13.2), style, w=cm(w), h=cm(1.0), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(13.2), "secondary", w=cm(2.6), h=cm(1.0),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = ["Private Sub Form_Load()", "    BackupLoad Me", "End Sub"] + m.code
    return m


def security_forms() -> List[FormModel]:
    lines, heads = layout_role_lines()
    screen_lines, screen_heads = layout_user_screen_lines()
    return [layout_login(), layout_change_password(), lines, layout_roles(heads),
            screen_lines, layout_user_screens(screen_heads), layout_activation(), layout_backup()]
