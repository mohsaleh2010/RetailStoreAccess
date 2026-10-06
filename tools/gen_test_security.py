"""Generate modTestSecurity.bas: the in-Access tests of Phase 10.
Hash vectors come from tools/security_reference.py (Python hashlib)."""

import datetime as dt
import re

import security_reference as S
from generate_common import vba_str


def vba_text_expr(text: str) -> str:
    """VBA expression for any text: characters outside Windows-1256 become ChrW(UTF-16 unit)."""
    parts, run = [], ""
    for ch in text:
        try:
            ch.encode("cp1256")
            run += ch
        except UnicodeEncodeError:
            if run:
                parts.append(vba_str(run))
                run = ""
            data = ch.encode("utf-16-be")
            for i in range(0, len(data), 2):
                parts.append(f"ChrW(&H{int.from_bytes(data[i:i + 2], 'big'):X})")
    if run or not parts:
        parts.append(vba_str(run))
    return " & ".join(parts)


def vector_lines():
    out = []
    for text in S.SHA_VECTORS:
        if len(text) > 100:
            src = f'String$({len(text)}, "{text[0]}")'
        else:
            src = vba_text_expr(text)
        label = "SHA-256: " + (text[:20].encode("cp1256", "replace").decode("cp1256") if text else "(فارغ)") + \
            f" ({len(text)})"
        out.append(f"    Call Record(Sha256Text({src}) = {vba_str(S.sha256_hex(text.encode()))}, {vba_str(label)})")
    for password, salt in S.PASSWORD_VECTORS:
        out.append(f"    Call Record(PasswordHash({vba_str(password)}, {vba_str(salt)}) = "
                   f"{vba_str(S.password_hash(password, salt))}, {vba_str('بصمة كلمة المرور: ' + password)})")
    when = dt.datetime(2026, 10, 4, 9, 5, 7)
    out.append(f'    Call Record(BackupFileName("RetailStore_BE", DateSerial(2026, 10, 4) + TimeSerial(9, 5, 7)) = '
               f'{vba_str(S.backup_file_name("RetailStore_BE", when))}, "اسم ملف النسخة بالتاريخ والوقت")')
    names = ["B_2026-10-03_090000.accdb", "B_2026-09-30_235959.accdb", "B_2026-10-04_080000.accdb",
             "B_2026-10-01_120000.accdb"]
    out.append(f'    Call Record(OldBackups({vba_str("|".join(names))}, 2) = '
               f'{vba_str("|".join(S.backups_to_delete(names, 2)))}, "حذف النسخ الأقدم مع إبقاء الأحدث")')
    return "\n".join(out)


TEMPLATE = r'''Attribute VB_Name = "modTestSecurity"
'==============================================================================
' modTestSecurity  -  Retail Store Management System (Phase 10 tests)
'
' GENERATED FILE - do not edit by hand.  Source: tools/gen_test_security.py
'
'   TestSecurity   1) SHA-256 and password-hash vectors (same as Python hashlib),
'                     backup file names and retention
'                  2) users and permissions inside a transaction that is rolled
'                     back: password rules, login, wrong passwords and lock-out,
'                     inactive user, user without password, cashier permissions,
'                     role permission changes, the last administrator
'                  3) a real backup into a temporary folder, verified and deleted
'                  4) screens of a user (frmUserScreens), the programmer and the
'                     activation of a made-up computer, rolled back
'==============================================================================
Option Compare Database
Option Explicit

Private m_passed As Long
Private m_failed As Long
Private m_report As String

Public Function TestSecurity() As Boolean
    Dim savedUser As Long
    m_passed = 0: m_failed = 0: m_report = ""
    Debug.Print "=== TestSecurity  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Calendar = vbCalGreg
    EnsureTestUser
    savedUser = CurrentUserID()
    g_SilentMode = True
    g_AutoAnswer = True

    TestVectors
    TestUsers
    TempVars.Add "UserID", savedUser
    TestBackupFile
    TestScreens
    TempVars.Add "UserID", savedUser

    g_SilentMode = False
    Debug.Print "--- نجح: " & m_passed & " | فشل: " & m_failed
    If m_failed = 0 Then
        TestMsg "جميع اختبارات الأمان ناجحة (" & m_passed & " اختبارًا)." & vbCrLf & _
               "لم تُترك أي بيانات اختبار.", vbInformation + MSG_RTL, "TestSecurity"
        TestSecurity = True
    Else
        TestMsg "نجح " & m_passed & " وفشل " & m_failed & ":" & vbCrLf & vbCrLf & Left$(m_report, 900), _
               vbExclamation + MSG_RTL, "TestSecurity"
    End If
End Function

'------------------------------------------------------------------------------
' 1) vectors
'------------------------------------------------------------------------------
Private Sub TestVectors()
@@VECTORS@@
    Call Record(Len(NewSalt()) = 32 And NewSalt() <> NewSalt(), "ملح عشوائي مختلف في كل مرة")
End Sub

'------------------------------------------------------------------------------
' 2) users and permissions, rolled back at the end
'------------------------------------------------------------------------------
Private Sub TestUsers()
    Dim ws As DAO.Workspace, db As DAO.Database, inTrans As Boolean, uid As Long, id As Long
    Dim msg As String, i As Long, adminID As Long, cashierPerms As Long
    On Error GoTo EH
    Set db = CurrentDb
    adminID = CurrentUserID()
    Set ws = DBEngine.Workspaces(0)
    cashierPerms = DCount("*", "RolePermissions", "RoleID = 3")      ' restored by the rollback
    ws.BeginTrans
    inTrans = True

    db.Execute "INSERT INTO Employees (EmployeeName, Username, RoleID, MaxDiscountPercent, IsActive) " & _
               "VALUES ('TEST كاشير', 'test_cashier', 3, 0, True)", dbFailOnError
    uid = DbValue("SELECT EmployeeID FROM Employees WHERE Username = 'test_cashier'")

    ' --- password rules
    Call Record(Len(PasswordProblem("12345", "x")) > 0, "رفض كلمة مرور أقل من 6 أحرف")
    Call Record(Len(PasswordProblem("test_cashier", "test_cashier")) > 0, "رفض كلمة مرور = اسم المستخدم")
    Call Record(Len(PasswordProblem(" abc123", "x")) > 0, "رفض مسافة في البداية")
    Call Record(Len(PasswordProblem("Kashier#26", "test_cashier")) = 0, "قبول كلمة مرور صحيحة")

    ' --- a user without a password cannot log in
    Call Record(Len(LoginUser("test_cashier", "", id)) > 0 And id = 0, "مستخدم بلا كلمة مرور لا يدخل")
    msg = SetUserPassword(uid, "Kashier#26", True)
    Call Record(Len(msg) = 0, "تعيين كلمة مرور مؤقتة " & msg)
    Call Record(Nz(DbValue("SELECT PasswordHash FROM Employees WHERE EmployeeID = " & uid), "") = _
                PasswordHash("Kashier#26", Nz(DbValue("SELECT PasswordSalt FROM Employees WHERE EmployeeID = " & uid), "")) _
                And InStr(Nz(DbValue("SELECT PasswordHash FROM Employees WHERE EmployeeID = " & uid), ""), "Kashier") = 0, _
                "تُحفظ البصمة فقط وليس كلمة المرور")

    ' --- wrong passwords and lock-out
    For i = 1 To MAX_FAILED_LOGINS - 1
        msg = LoginUser("test_cashier", "wrong-" & i, id)
    Next
    Call Record(id = 0 And _
                DbValue("SELECT FailedLoginCount FROM Employees WHERE EmployeeID = " & uid) = MAX_FAILED_LOGINS - 1, _
                "عدّ المحاولات الخاطئة")
    msg = LoginUser("test_cashier", "wrong-last", id)
    Call Record(id = 0 And Not IsNull(DbValue("SELECT LockedUntil FROM Employees WHERE EmployeeID = " & uid)), _
                "قفل الحساب بعد " & MAX_FAILED_LOGINS & " محاولات خاطئة")
    Call Record(Len(LoginUser("test_cashier", "Kashier#26", id)) > 0 And id = 0, "الحساب المقفل يرفض حتى كلمة المرور الصحيحة")
    db.Execute "UPDATE Employees SET LockedUntil = DateAdd('n', -1, Now()) WHERE EmployeeID = " & uid, dbFailOnError
    msg = LoginUser("test_cashier", "Kashier#26", id)
    Call Record(Len(msg) = 0 And id = uid And CurrentUserID() = uid, "الدخول بعد انتهاء القفل " & msg)
    Call Record(DbValue("SELECT FailedLoginCount FROM Employees WHERE EmployeeID = " & uid) = 0 And _
                IsNull(DbValue("SELECT LockedUntil FROM Employees WHERE EmployeeID = " & uid)), "تصفير العداد عند الدخول")
    Call Record(NeedsPasswordChange(uid), "كلمة المرور المؤقتة تتطلب التغيير")
    Call Record(Len(LoginUser("no_such_user", "x", id)) > 0, "اسم مستخدم غير موجود يُرفض")
    Call Record(LoginUser("no_such_user", "x", id) = LoginUser("test_cashier", "bad-pass", id), _
                "نفس الرسالة لاسم خاطئ ولكلمة خاطئة (لا يكشف وجود المستخدم)")
    db.Execute "UPDATE Employees SET FailedLoginCount = 0 WHERE EmployeeID = " & uid, dbFailOnError
    msg = LoginUser("test_cashier", "Kashier#26", id)

    ' --- the cashier's permissions (logged in as the cashier now)
    Call Record(HasPermission("SALES_POS") And Not HasPermission("REPORTS_PROFIT") And Not HasPermission("USERS"), _
                "صلاحيات الكاشير: البيع نعم، الأرباح والمستخدمون لا")
    Call Record(CanOpenScreen("frmPOS", True) And Not CanOpenScreen("frmSettings", True) And _
                Not CanOpenScreen("frmUsers", True) And Not CanOpenScreen("frmBackup", True), _
                "الكاشير يفتح نقطة البيع ولا يفتح الإعدادات والمستخدمين والنسخ")
    Call Record(Len(SaveRolePermissionSet(3, "SALES_POS")) > 0, "الكاشير لا يغير الصلاحيات")

    ' --- own password change
    Call Record(Len(ChangeOwnPassword("bad", "NewPass#1", "NewPass#1", True)) > 0, "رفض تغيير بكلمة حالية خاطئة")
    Call Record(Len(ChangeOwnPassword("Kashier#26", "NewPass#1", "Other#1", True)) > 0, "رفض تأكيد غير مطابق")
    msg = ChangeOwnPassword("Kashier#26", "NewPass#1", "NewPass#1", True)
    Call Record(Len(msg) = 0 And Not NeedsPasswordChange(uid), "تغيير كلمة المرور يلغي شرط التغيير " & msg)
    Call Record(Len(LoginUser("test_cashier", "NewPass#1", id)) = 0, "الدخول بكلمة المرور الجديدة")

    ' --- back to the administrator: role permissions and inactive users
    TempVars.Add "UserID", adminID
    msg = SaveRolePermissionSet(3, "SALES_POS,SALES_VIEW,CUSTOMERS,CUSTOMER_PAYMENTS,REPORTS")
    Call Record(Len(msg) = 0 And DbValue("SELECT COUNT(*) FROM RolePermissions WHERE RoleID = 3") = 5, "المدير يضيف صلاحية للكاشير " & msg)
    Call Record(Len(SaveRolePermissionSet(ADMIN_ROLE_ID, "SALES_POS")) > 0 And _
                DbValue("SELECT COUNT(*) FROM RolePermissions WHERE RoleID = " & ADMIN_ROLE_ID) = _
                DbValue("SELECT COUNT(*) FROM Permissions"), _
                "صلاحيات مدير النظام لا تُقيَّد")
    db.Execute "UPDATE Employees SET IsActive = False WHERE EmployeeID = " & uid, dbFailOnError
    Call Record(Len(LoginUser("test_cashier", "NewPass#1", id)) > 0 And id = 0, "المستخدم المعطّل لا يدخل")
    Call Record(CurrentUserID() = adminID, "الدخول الفاشل لا يغير المستخدم الحالي")

    ws.Rollback
    inTrans = False
    TempVars.Add "UserID", adminID
    Call Record(DCount("*", "Employees", "Username = 'test_cashier'") = 0 And _
                DCount("*", "RolePermissions", "RoleID = 3") = cashierPerms, "التراجع عن كل بيانات الاختبار")
    Exit Sub
EH:
    Fail "خطأ غير متوقع " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If inTrans Then ws.Rollback
    TempVars.Add "UserID", adminID
End Sub

'------------------------------------------------------------------------------
' 3) a real backup in a temporary folder
'------------------------------------------------------------------------------
Private Sub TestBackupFile()
    Dim folder As String, file As String, msg As String
    On Error GoTo EH
    folder = Environ$("TEMP") & "\RetailStoreBackupTest"
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    msg = BackupNow(file, folder)
    Call Record(Len(msg) = 0 And Len(Dir$(file)) > 0, "نسخة احتياطية فعلية " & msg)
    Call Record(BackupIsReadable(file), "النسخة تُفتح وفيها جداول البرنامج")
    Call Record(InStr(file, "_" & Format$(Date, "yyyy-mm-dd") & "_") > 0, "اسم النسخة بتاريخ اليوم")
    If Len(file) > 0 And Len(Dir$(file)) > 0 Then Kill file
    ' a test copy is not a real backup: keep "last backup" (shown on exit) truthful
    CurrentDb.Execute "DELETE FROM AuditLog WHERE ActionType = 'BACKUP' AND Details = " & SqlText(file), dbFailOnError
    RmDir folder
    Exit Sub
EH:
    Fail "النسخ الاحتياطي: خطأ " & Err.Number & ": " & Err.Description
End Sub

'------------------------------------------------------------------------------
' 4) screens of a user, the programmer and activation, rolled back at the end
'------------------------------------------------------------------------------
Private Sub TestScreens()
    Dim ws As DAO.Workspace, db As DAO.Database, inTrans As Boolean, uid As Long, dev As Long, adminID As Long
    Dim msg As String, mid As String
    On Error GoTo EH
    Set db = CurrentDb
    adminID = CurrentUserID()
    EnsureLocalTables
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    db.Execute "INSERT INTO Employees (EmployeeName, Username, RoleID, MaxDiscountPercent, IsActive) " & _
               "VALUES ('TEST شاشات', 'test_screens', 3, 0, True)", dbFailOnError
    uid = DbValue("SELECT EmployeeID FROM Employees WHERE Username = 'test_screens'")
    db.Execute "INSERT INTO Employees (EmployeeName, Username, RoleID, MaxDiscountPercent, IsActive, IsDeveloper) " & _
               "VALUES ('TEST مبرمج', 'test_dev', 3, 0, True, True)", dbFailOnError
    dev = DbValue("SELECT EmployeeID FROM Employees WHERE Username = 'test_dev'")

    ' --- the grid starts from the role; then own screens: no customers, products without add / delete
    FillUserScreens uid, True
    Call Record(DbValue("SELECT CanOpen AND CanAdd FROM tmpUserScreens WHERE ScreenName = 'frmPOS'") And _
                Not DbValue("SELECT CanOpen FROM tmpUserScreens WHERE ScreenName = 'frmSettings'"), _
                "شاشات المستخدم تبدأ من صلاحيات دوره")
    db.Execute "UPDATE tmpUserScreens SET CanOpen = False, CanAdd = False, CanEdit = False, CanDelete = False " & _
               "WHERE ScreenName = 'frmCustomers'", dbFailOnError
    db.Execute "UPDATE tmpUserScreens SET CanOpen = True, CanAdd = False, CanEdit = True, CanDelete = False " & _
               "WHERE ScreenName = 'frmProducts'", dbFailOnError
    msg = SaveUserScreensFor(uid, True)
    Call Record(Len(msg) = 0 And UsesCustomScreens(uid), "حفظ شاشات خاصة للمستخدم " & msg)
    TempVars.Add "UserID", uid
    Call Record(Not CanOpenScreen("frmCustomers", True) And CanOpenScreen("frmProducts", True) And _
                CanOpenScreen("frmPOS", True), "يفتح الشاشات المحددة له فقط")
    Call Record(Not HasPermission("CUSTOMERS") And HasPermission("PRODUCTS") And HasPermission("SALES_POS") And _
                Not HasPermission("REPORTS_PROFIT"), "صلاحيات الشاشات تتبع اختياره، والصلاحيات الخاصة من دوره")
    Call Record(Not CanScreenAction("frmProducts", "ADD", True) And CanScreenAction("frmProducts", "EDIT", True) And _
                Not CanScreenAction("frmProducts", "DELETE", True), "المنتجات: تعديل بدون إضافة أو حذف")
    Call Record(CanScreenAction("frmPOS", "ADD", True) And CanScreenAction("frmSalesInvoice", "EDIT", True), _
                "الإجراء غير الموجود في الشاشة لا يمنع شيئًا")
    Call Record(Len(SaveUserScreensFor(uid, False)) > 0, "المستخدم لا يغير صلاحياته")
    TempVars.Add "UserID", adminID
    Call Record(Len(SaveUserScreensFor(adminID, True)) > 0, "شاشات مدير النظام لا تُقيَّد")
    msg = SaveUserScreensFor(uid, False)
    TempVars.Add "UserID", uid
    Call Record(Len(msg) = 0 And CanOpenScreen("frmCustomers", True) And _
                DbValue("SELECT COUNT(*) FROM UserScreens WHERE EmployeeID = " & uid) = 0, "الرجوع لصلاحيات الدور " & msg)

    ' --- the programmer, and the shop name
    TempVars.Add "UserID", dev
    Call Record(IsDeveloper() And HasPermission("USERS") And HasPermission("REPORTS_PROFIT") And _
                CanOpenScreen("frmBackup", True) And CanChangeStoreName(), "المبرمج يملك كل الصلاحيات ويغير اسم المحل")
    TempVars.Add "UserID", adminID
    db.Execute "UPDATE Settings SET AllowAdminCompanyName = False", dbFailOnError
    Call Record(Not IsDeveloper() And Not CanChangeStoreName(), "مدير النظام لا يغير اسم المحل بدون إذن المبرمج")
    db.Execute "UPDATE Settings SET AllowAdminCompanyName = True", dbFailOnError
    Call Record(CanChangeStoreName(), "بإذن المبرمج يغير مدير النظام اسم المحل")

    ' --- activation of a made-up computer
    mid = "@@MID@@"
    Call Record(ActivationCodeFor(mid) = "@@CODE@@", "كود التفعيل يطابق المرجع (tools/activation_reference.py)")
    Call Record(Len(SaveActivation(mid, "11111-22222-33333-44444", "TEST-PC")) > 0, "رفض كود تفعيل خاطئ")
    Call Record(Not IsValidCode("1A2B-3C4D-5E6F-7A8C", ActivationCodeFor(mid)), "الكود لا يصلح لجهاز آخر")
    msg = SaveActivation(mid, LCase$(ActivationCodeFor(mid)), "TEST-PC")
    Call Record(Len(msg) = 0 And DbValue("SELECT COUNT(*) FROM Activations WHERE MachineID = '" & mid & "'") = 1, _
                "تفعيل جهاز بالكود الصحيح " & msg)
    TempVars.Add "UserID", uid
    Call Record(Len(SaveActivation(mid, ActivationCodeFor(mid), "TEST-PC")) > 0, "التفعيل لمدير النظام فقط")
    Call Record(MachineID() Like "????-????-????-????", "رقم هذا الجهاز " & MachineID())

    ws.Rollback
    inTrans = False
    TempVars.Add "UserID", adminID
    Call Record(DCount("*", "Employees", "Username = 'test_screens' OR Username = 'test_dev'") = 0 And _
                DCount("*", "Activations", "MachineID = '" & mid & "'") = 0, "التراجع عن بيانات اختبار الشاشات والتفعيل")
    Exit Sub
EH:
    Fail "الشاشات والتفعيل: خطأ " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If inTrans Then ws.Rollback
    TempVars.Add "UserID", adminID
End Sub

Private Sub Record(ByVal Passed As Boolean, ByVal Label As String)
    If Passed Then
        m_passed = m_passed + 1
        Debug.Print "[OK] " & Label
    Else
        Fail Label
    End If
End Sub

Private Sub Fail(ByVal Msg As String)
    m_failed = m_failed + 1
    m_report = m_report & "- " & Msg & vbCrLf
    Debug.Print "[X] " & Msg
End Sub
'''


def build_test_security_vba() -> str:
    import activation_reference as A
    mid = "1A2B-3C4D-5E6F-7A8B"
    text = TEMPLATE.replace("@@VECTORS@@", vector_lines()).replace("@@MID@@", mid) \
        .replace("@@CODE@@", A.activation_code(mid, A.vba_secret()))
    assert not re.search(r"@@[A-Z_]+@@", text)
    return text
