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
    return [layout_login(), layout_change_password(), lines, layout_roles(heads), layout_backup()]
