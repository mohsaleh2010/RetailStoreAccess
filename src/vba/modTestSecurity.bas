Attribute VB_Name = "modTestSecurity"
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
    Call Record(Sha256Text("") = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855", "SHA-256: (فارغ) (0)")
    Call Record(Sha256Text("abc") = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", "SHA-256: abc (3)")
    Call Record(Sha256Text("abcdbcdecdefdefgefghfghighijhijkijkljklmmnomnopnopq") = "d17ddb2e3c6e7ee4ea00838f71742fb81c9a8bd785cdf422a4340ea0aa13ddce", "SHA-256: abcdbcdecdefdefgefgh (51)")
    Call Record(Sha256Text("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa") = "9f4390f8d30c2dd92ec9f095b65e2b9ae9b0a925a5258e241c9f1e910f734318", "SHA-256: aaaaaaaaaaaaaaaaaaaa (55)")
    Call Record(Sha256Text("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa") = "b35439a4ac6f0948b6d6f9e3c6af0f5f590ce20f1bde7090ef7970686ec6738a", "SHA-256: aaaaaaaaaaaaaaaaaaaa (56)")
    Call Record(Sha256Text("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa") = "ffe054fe7ae0cb6dc65c3af9b61d5209f439851db43d0ba5997337df154668eb", "SHA-256: aaaaaaaaaaaaaaaaaaaa (64)")
    Call Record(Sha256Text(String$(1000, "a")) = "41edece42d63e8d9bf515a9ba6932e1c20cbc9f5a5d134645adb5db1b9737ea3", "SHA-256: aaaaaaaaaaaaaaaaaaaa (1000)")
    Call Record(Sha256Text("متجر الاختبار") = "fba445173dbe1c333f7c5f370bb40e8b24849419ea03ce46c47ed4aeacc04883", "SHA-256: متجر الاختبار (13)")
    Call Record(Sha256Text("كلمة-سر 123 " & ChrW(&HD83D) & ChrW(&HDE00)) = "60cb98d4b7b085f1ff33d28bb78b37bc6f30a700879624efd777fc555ddf5455", "SHA-256: كلمة-سر 123 ? (13)")
    Call Record(PasswordHash("Admin@2026", "0F1E2D3C4B5A69788796A5B4C3D2E1F0") = "5139079eda9d2161c3572eaebe60ffb450a722363e76ac2e3c9a09a2bfc9696a", "بصمة كلمة المرور: Admin@2026")
    Call Record(PasswordHash("كلمة سر عربية", "00000000000000000000000000000000") = "c67d96dc5473ce557d25bf78d067fca97d10abba359bbad81a3aa8da77c8831e", "بصمة كلمة المرور: كلمة سر عربية")
    Call Record(PasswordHash("123456", "ABCDEFABCDEFABCDEFABCDEFABCDEF12") = "aa10a07a24547d2ccbedde456c580ed3ca0b6cb94539ac619532d851ff299f7b", "بصمة كلمة المرور: 123456")
    Call Record(BackupFileName("RetailStore_BE", DateSerial(2026, 10, 4) + TimeSerial(9, 5, 7)) = "RetailStore_BE_2026-10-04_090507.accdb", "اسم ملف النسخة بالتاريخ والوقت")
    Call Record(OldBackups("B_2026-10-03_090000.accdb|B_2026-09-30_235959.accdb|B_2026-10-04_080000.accdb|B_2026-10-01_120000.accdb", 2) = "B_2026-09-30_235959.accdb|B_2026-10-01_120000.accdb", "حذف النسخ الأقدم مع إبقاء الأحدث")
    Call Record(Len(NewSalt()) = 32 And NewSalt() <> NewSalt(), "ملح عشوائي مختلف في كل مرة")
End Sub

'------------------------------------------------------------------------------
' 2) users and permissions, rolled back at the end
'------------------------------------------------------------------------------
Private Sub TestUsers()
    Dim ws As DAO.Workspace, db As DAO.Database, inTrans As Boolean, uid As Long, id As Long
    Dim msg As String, i As Long, adminID As Long
    On Error GoTo EH
    Set db = CurrentDb
    adminID = CurrentUserID()
    Set ws = DBEngine.Workspaces(0)
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
                DCount("*", "RolePermissions", "RoleID = 3") = 4, "التراجع عن كل بيانات الاختبار")
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
