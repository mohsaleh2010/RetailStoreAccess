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
    Debug.Print "--- ‰ÃÕ: " & m_passed & " | ›‘·: " & m_failed
    If m_failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·√„«‰ ‰«ÃÕ… (" & m_passed & " «Œ »«—«)." & vbCrLf & _
               "·„  ı —ﬂ √Ì »Ì«‰«  «Œ »«—.", vbInformation + MSG_RTL, "TestSecurity"
        TestSecurity = True
    Else
        TestMsg "‰ÃÕ " & m_passed & " Ê›‘· " & m_failed & ":" & vbCrLf & vbCrLf & Left$(m_report, 900), _
               vbExclamation + MSG_RTL, "TestSecurity"
    End If
End Function

'------------------------------------------------------------------------------
' 1) vectors
'------------------------------------------------------------------------------
Private Sub TestVectors()
    Call Record(Sha256Text("") = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855", "SHA-256: (›«—€) (0)")
    Call Record(Sha256Text("abc") = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", "SHA-256: abc (3)")
    Call Record(Sha256Text("abcdbcdecdefdefgefghfghighijhijkijkljklmmnomnopnopq") = "d17ddb2e3c6e7ee4ea00838f71742fb81c9a8bd785cdf422a4340ea0aa13ddce", "SHA-256: abcdbcdecdefdefgefgh (51)")
    Call Record(Sha256Text("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa") = "9f4390f8d30c2dd92ec9f095b65e2b9ae9b0a925a5258e241c9f1e910f734318", "SHA-256: aaaaaaaaaaaaaaaaaaaa (55)")
    Call Record(Sha256Text("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa") = "b35439a4ac6f0948b6d6f9e3c6af0f5f590ce20f1bde7090ef7970686ec6738a", "SHA-256: aaaaaaaaaaaaaaaaaaaa (56)")
    Call Record(Sha256Text("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa") = "ffe054fe7ae0cb6dc65c3af9b61d5209f439851db43d0ba5997337df154668eb", "SHA-256: aaaaaaaaaaaaaaaaaaaa (64)")
    Call Record(Sha256Text(String$(1000, "a")) = "41edece42d63e8d9bf515a9ba6932e1c20cbc9f5a5d134645adb5db1b9737ea3", "SHA-256: aaaaaaaaaaaaaaaaaaaa (1000)")
    Call Record(Sha256Text("„ Ã— «·«Œ »«—") = "fba445173dbe1c333f7c5f370bb40e8b24849419ea03ce46c47ed4aeacc04883", "SHA-256: „ Ã— «·«Œ »«— (13)")
    Call Record(Sha256Text("ﬂ·„…-”— 123 " & ChrW(&HD83D) & ChrW(&HDE00)) = "60cb98d4b7b085f1ff33d28bb78b37bc6f30a700879624efd777fc555ddf5455", "SHA-256: ﬂ·„…-”— 123 ? (13)")
    Call Record(PasswordHash("Admin@2026", "0F1E2D3C4B5A69788796A5B4C3D2E1F0") = "5139079eda9d2161c3572eaebe60ffb450a722363e76ac2e3c9a09a2bfc9696a", "»’„… ﬂ·„… «·„—Ê—: Admin@2026")
    Call Record(PasswordHash("ﬂ·„… ”— ⁄—»Ì…", "00000000000000000000000000000000") = "c67d96dc5473ce557d25bf78d067fca97d10abba359bbad81a3aa8da77c8831e", "»’„… ﬂ·„… «·„—Ê—: ﬂ·„… ”— ⁄—»Ì…")
    Call Record(PasswordHash("123456", "ABCDEFABCDEFABCDEFABCDEFABCDEF12") = "aa10a07a24547d2ccbedde456c580ed3ca0b6cb94539ac619532d851ff299f7b", "»’„… ﬂ·„… «·„—Ê—: 123456")
    Call Record(BackupFileName("RetailStore_BE", DateSerial(2026, 10, 4) + TimeSerial(9, 5, 7)) = "RetailStore_BE_2026-10-04_090507.accdb", "«”„ „·› «·‰”Œ… »«· «—ÌŒ Ê«·Êﬁ ")
    Call Record(OldBackups("B_2026-10-03_090000.accdb|B_2026-09-30_235959.accdb|B_2026-10-04_080000.accdb|B_2026-10-01_120000.accdb", 2) = "B_2026-09-30_235959.accdb|B_2026-10-01_120000.accdb", "Õ–› «·‰”Œ «·√ﬁœ„ „⁄ ≈»ﬁ«¡ «·√ÕœÀ")
    Call Record(Len(NewSalt()) = 32 And NewSalt() <> NewSalt(), "„·Õ ⁄‘Ê«∆Ì „Œ ·› ›Ì ﬂ· „—…")
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
               "VALUES ('TEST ﬂ«‘Ì—', 'test_cashier', 3, 0, True)", dbFailOnError
    uid = DbValue("SELECT EmployeeID FROM Employees WHERE Username = 'test_cashier'")

    ' --- password rules
    Call Record(Len(PasswordProblem("12345", "x")) > 0, "—›÷ ﬂ·„… „—Ê— √ﬁ· „‰ 6 √Õ—›")
    Call Record(Len(PasswordProblem("test_cashier", "test_cashier")) > 0, "—›÷ ﬂ·„… „—Ê— = «”„ «·„” Œœ„")
    Call Record(Len(PasswordProblem(" abc123", "x")) > 0, "—›÷ „”«›… ›Ì «·»œ«Ì…")
    Call Record(Len(PasswordProblem("Kashier#26", "test_cashier")) = 0, "ﬁ»Ê· ﬂ·„… „—Ê— ’ÕÌÕ…")

    ' --- a user without a password cannot log in
    Call Record(Len(LoginUser("test_cashier", "", id)) > 0 And id = 0, "„” Œœ„ »·« ﬂ·„… „—Ê— ·« ÌœŒ·")
    msg = SetUserPassword(uid, "Kashier#26", True)
    Call Record(Len(msg) = 0, " ⁄ÌÌ‰ ﬂ·„… „—Ê— „ƒﬁ … " & msg)
    Call Record(Nz(DbValue("SELECT PasswordHash FROM Employees WHERE EmployeeID = " & uid), "") = _
                PasswordHash("Kashier#26", Nz(DbValue("SELECT PasswordSalt FROM Employees WHERE EmployeeID = " & uid), "")) _
                And InStr(Nz(DbValue("SELECT PasswordHash FROM Employees WHERE EmployeeID = " & uid), ""), "Kashier") = 0, _
                " ıÕ›Ÿ «·»’„… ›ﬁÿ Ê·Ì” ﬂ·„… «·„—Ê—")

    ' --- wrong passwords and lock-out
    For i = 1 To MAX_FAILED_LOGINS - 1
        msg = LoginUser("test_cashier", "wrong-" & i, id)
    Next
    Call Record(id = 0 And _
                DbValue("SELECT FailedLoginCount FROM Employees WHERE EmployeeID = " & uid) = MAX_FAILED_LOGINS - 1, _
                "⁄œ¯ «·„Õ«Ê·«  «·Œ«ÿ∆…")
    msg = LoginUser("test_cashier", "wrong-last", id)
    Call Record(id = 0 And Not IsNull(DbValue("SELECT LockedUntil FROM Employees WHERE EmployeeID = " & uid)), _
                "ﬁ›· «·Õ”«» »⁄œ " & MAX_FAILED_LOGINS & " „Õ«Ê·«  Œ«ÿ∆…")
    Call Record(Len(LoginUser("test_cashier", "Kashier#26", id)) > 0 And id = 0, "«·Õ”«» «·„ﬁ›· Ì—›÷ Õ Ï ﬂ·„… «·„—Ê— «·’ÕÌÕ…")
    db.Execute "UPDATE Employees SET LockedUntil = DateAdd('n', -1, Now()) WHERE EmployeeID = " & uid, dbFailOnError
    msg = LoginUser("test_cashier", "Kashier#26", id)
    Call Record(Len(msg) = 0 And id = uid And CurrentUserID() = uid, "«·œŒÊ· »⁄œ «‰ Â«¡ «·ﬁ›· " & msg)
    Call Record(DbValue("SELECT FailedLoginCount FROM Employees WHERE EmployeeID = " & uid) = 0 And _
                IsNull(DbValue("SELECT LockedUntil FROM Employees WHERE EmployeeID = " & uid)), " ’›Ì— «·⁄œ«œ ⁄‰œ «·œŒÊ·")
    Call Record(NeedsPasswordChange(uid), "ﬂ·„… «·„—Ê— «·„ƒﬁ …   ÿ·» «· €ÌÌ—")
    Call Record(Len(LoginUser("no_such_user", "x", id)) > 0, "«”„ „” Œœ„ €Ì— „ÊÃÊœ Ìı—›÷")
    Call Record(LoginUser("no_such_user", "x", id) = LoginUser("test_cashier", "bad-pass", id), _
                "‰›” «·—”«·… ·«”„ Œ«ÿ∆ Ê·ﬂ·„… Œ«ÿ∆… (·« Ìﬂ‘› ÊÃÊœ «·„” Œœ„)")
    db.Execute "UPDATE Employees SET FailedLoginCount = 0 WHERE EmployeeID = " & uid, dbFailOnError
    msg = LoginUser("test_cashier", "Kashier#26", id)

    ' --- the cashier's permissions (logged in as the cashier now)
    Call Record(HasPermission("SALES_POS") And Not HasPermission("REPORTS_PROFIT") And Not HasPermission("USERS"), _
                "’·«ÕÌ«  «·ﬂ«‘Ì—: «·»Ì⁄ ‰⁄„° «·√—»«Õ Ê«·„” Œœ„Ê‰ ·«")
    Call Record(CanOpenScreen("frmPOS", True) And Not CanOpenScreen("frmSettings", True) And _
                Not CanOpenScreen("frmUsers", True) And Not CanOpenScreen("frmBackup", True), _
                "«·ﬂ«‘Ì— Ì› Õ ‰ﬁÿ… «·»Ì⁄ Ê·« Ì› Õ «·≈⁄œ«œ«  Ê«·„” Œœ„Ì‰ Ê«·‰”Œ")
    Call Record(Len(SaveRolePermissionSet(3, "SALES_POS")) > 0, "«·ﬂ«‘Ì— ·« Ì€Ì— «·’·«ÕÌ« ")

    ' --- own password change
    Call Record(Len(ChangeOwnPassword("bad", "NewPass#1", "NewPass#1", True)) > 0, "—›÷  €ÌÌ— »ﬂ·„… Õ«·Ì… Œ«ÿ∆…")
    Call Record(Len(ChangeOwnPassword("Kashier#26", "NewPass#1", "Other#1", True)) > 0, "—›÷  √ﬂÌœ €Ì— „ÿ«»ﬁ")
    msg = ChangeOwnPassword("Kashier#26", "NewPass#1", "NewPass#1", True)
    Call Record(Len(msg) = 0 And Not NeedsPasswordChange(uid), " €ÌÌ— ﬂ·„… «·„—Ê— Ì·€Ì ‘—ÿ «· €ÌÌ— " & msg)
    Call Record(Len(LoginUser("test_cashier", "NewPass#1", id)) = 0, "«·œŒÊ· »ﬂ·„… «·„—Ê— «·ÃœÌœ…")

    ' --- back to the administrator: role permissions and inactive users
    TempVars.Add "UserID", adminID
    msg = SaveRolePermissionSet(3, "SALES_POS,SALES_VIEW,CUSTOMERS,CUSTOMER_PAYMENTS,REPORTS")
    Call Record(Len(msg) = 0 And DbValue("SELECT COUNT(*) FROM RolePermissions WHERE RoleID = 3") = 5, "«·„œÌ— Ì÷Ì› ’·«ÕÌ… ··ﬂ«‘Ì— " & msg)
    Call Record(Len(SaveRolePermissionSet(ADMIN_ROLE_ID, "SALES_POS")) > 0 And _
                DbValue("SELECT COUNT(*) FROM RolePermissions WHERE RoleID = " & ADMIN_ROLE_ID) = _
                DbValue("SELECT COUNT(*) FROM Permissions"), _
                "’·«ÕÌ«  „œÌ— «·‰Ÿ«„ ·«  ıﬁÌÛ¯œ")
    db.Execute "UPDATE Employees SET IsActive = False WHERE EmployeeID = " & uid, dbFailOnError
    Call Record(Len(LoginUser("test_cashier", "NewPass#1", id)) > 0 And id = 0, "«·„” Œœ„ «·„⁄ÿ¯· ·« ÌœŒ·")
    Call Record(CurrentUserID() = adminID, "«·œŒÊ· «·›«‘· ·« Ì€Ì— «·„” Œœ„ «·Õ«·Ì")

    ws.Rollback
    inTrans = False
    TempVars.Add "UserID", adminID
    Call Record(DCount("*", "Employees", "Username = 'test_cashier'") = 0 And _
                DCount("*", "RolePermissions", "RoleID = 3") = cashierPerms, "«· —«Ã⁄ ⁄‰ ﬂ· »Ì«‰«  «·«Œ »«—")
    Exit Sub
EH:
    Fail "Œÿ√ €Ì— „ Êﬁ⁄ " & Err.Number & ": " & Err.Description
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
    Call Record(Len(msg) = 0 And Len(Dir$(file)) > 0, "‰”Œ… «Õ Ì«ÿÌ… ›⁄·Ì… " & msg)
    Call Record(BackupIsReadable(file), "«·‰”Œ…  ı› Õ Ê›ÌÂ« Ãœ«Ê· «·»—‰«„Ã")
    Call Record(InStr(file, "_" & Format$(Date, "yyyy-mm-dd") & "_") > 0, "«”„ «·‰”Œ… » «—ÌŒ «·ÌÊ„")
    If Len(file) > 0 And Len(Dir$(file)) > 0 Then Kill file
    ' a test copy is not a real backup: keep "last backup" (shown on exit) truthful
    CurrentDb.Execute "DELETE FROM AuditLog WHERE ActionType = 'BACKUP' AND Details = " & SqlText(file), dbFailOnError
    RmDir folder
    Exit Sub
EH:
    Fail "«·‰”Œ «·«Õ Ì«ÿÌ: Œÿ√ " & Err.Number & ": " & Err.Description
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
               "VALUES ('TEST ‘«‘« ', 'test_screens', 3, 0, True)", dbFailOnError
    uid = DbValue("SELECT EmployeeID FROM Employees WHERE Username = 'test_screens'")
    db.Execute "INSERT INTO Employees (EmployeeName, Username, RoleID, MaxDiscountPercent, IsActive, IsDeveloper) " & _
               "VALUES ('TEST „»—„Ã', 'test_dev', 3, 0, True, True)", dbFailOnError
    dev = DbValue("SELECT EmployeeID FROM Employees WHERE Username = 'test_dev'")

    ' --- the grid starts from the role; then own screens: no customers, products without add / delete
    FillUserScreens uid, True
    Call Record(DbValue("SELECT CanOpen AND CanAdd FROM tmpUserScreens WHERE ScreenName = 'frmPOS'") And _
                Not DbValue("SELECT CanOpen FROM tmpUserScreens WHERE ScreenName = 'frmSettings'"), _
                "‘«‘«  «·„” Œœ„  »œ√ „‰ ’·«ÕÌ«  œÊ—Â")
    db.Execute "UPDATE tmpUserScreens SET CanOpen = False, CanAdd = False, CanEdit = False, CanDelete = False " & _
               "WHERE ScreenName = 'frmCustomers'", dbFailOnError
    db.Execute "UPDATE tmpUserScreens SET CanOpen = True, CanAdd = False, CanEdit = True, CanDelete = False " & _
               "WHERE ScreenName = 'frmProducts'", dbFailOnError
    msg = SaveUserScreensFor(uid, True)
    Call Record(Len(msg) = 0 And UsesCustomScreens(uid), "Õ›Ÿ ‘«‘«  Œ«’… ··„” Œœ„ " & msg)
    TempVars.Add "UserID", uid
    Call Record(Not CanOpenScreen("frmCustomers", True) And CanOpenScreen("frmProducts", True) And _
                CanOpenScreen("frmPOS", True), "Ì› Õ «·‘«‘«  «·„Õœœ… ·Â ›ﬁÿ")
    Call Record(Not HasPermission("CUSTOMERS") And HasPermission("PRODUCTS") And HasPermission("SALES_POS") And _
                Not HasPermission("REPORTS_PROFIT"), "’·«ÕÌ«  «·‘«‘«    »⁄ «Œ Ì«—Â° Ê«·’·«ÕÌ«  «·Œ«’… „‰ œÊ—Â")
    Call Record(Not CanScreenAction("frmProducts", "ADD", True) And CanScreenAction("frmProducts", "EDIT", True) And _
                Not CanScreenAction("frmProducts", "DELETE", True), "«·„‰ Ã« :  ⁄œÌ· »œÊ‰ ≈÷«›… √Ê Õ–›")
    Call Record(CanScreenAction("frmPOS", "ADD", True) And CanScreenAction("frmSalesInvoice", "EDIT", True), _
                "«·≈Ã—«¡ €Ì— «·„ÊÃÊœ ›Ì «·‘«‘… ·« Ì„‰⁄ ‘Ì∆«")
    Call Record(Len(SaveUserScreensFor(uid, False)) > 0, "«·„” Œœ„ ·« Ì€Ì— ’·«ÕÌ« Â")
    TempVars.Add "UserID", adminID
    Call Record(Len(SaveUserScreensFor(adminID, True)) > 0, "‘«‘«  „œÌ— «·‰Ÿ«„ ·«  ıﬁÌÛ¯œ")
    msg = SaveUserScreensFor(uid, False)
    TempVars.Add "UserID", uid
    Call Record(Len(msg) = 0 And CanOpenScreen("frmCustomers", True) And _
                DbValue("SELECT COUNT(*) FROM UserScreens WHERE EmployeeID = " & uid) = 0, "«·—ÃÊ⁄ ·’·«ÕÌ«  «·œÊ— " & msg)

    ' --- the programmer, and the shop name
    TempVars.Add "UserID", dev
    Call Record(IsDeveloper() And HasPermission("USERS") And HasPermission("REPORTS_PROFIT") And _
                CanOpenScreen("frmBackup", True) And CanChangeStoreName(), "«·„»—„Ã Ì„·ﬂ ﬂ· «·’·«ÕÌ«  ÊÌ€Ì— «”„ «·„Õ·")
    TempVars.Add "UserID", adminID
    db.Execute "UPDATE Settings SET AllowAdminCompanyName = False", dbFailOnError
    Call Record(Not IsDeveloper() And Not CanChangeStoreName(), "„œÌ— «·‰Ÿ«„ ·« Ì€Ì— «”„ «·„Õ· »œÊ‰ ≈–‰ «·„»—„Ã")
    db.Execute "UPDATE Settings SET AllowAdminCompanyName = True", dbFailOnError
    Call Record(CanChangeStoreName(), "»≈–‰ «·„»—„Ã Ì€Ì— „œÌ— «·‰Ÿ«„ «”„ «·„Õ·")

    ' --- activation of a made-up computer
    mid = "1A2B-3C4D-5E6F-7A8B"
    Call Record(ActivationCodeFor(mid) = "7AAE2-F50AF-20481-9EBAB", "ﬂÊœ «· ›⁄Ì· Ìÿ«»ﬁ «·„—Ã⁄ (tools/activation_reference.py)")
    Call Record(Len(SaveActivation(mid, "11111-22222-33333-44444", "TEST-PC")) > 0, "—›÷ ﬂÊœ  ›⁄Ì· Œ«ÿ∆")
    Call Record(Not IsValidCode("1A2B-3C4D-5E6F-7A8C", ActivationCodeFor(mid)), "«·ﬂÊœ ·« Ì’·Õ ·ÃÂ«“ ¬Œ—")
    msg = SaveActivation(mid, LCase$(ActivationCodeFor(mid)), "TEST-PC")
    Call Record(Len(msg) = 0 And DbValue("SELECT COUNT(*) FROM Activations WHERE MachineID = '" & mid & "'") = 1, _
                " ›⁄Ì· ÃÂ«“ »«·ﬂÊœ «·’ÕÌÕ " & msg)
    TempVars.Add "UserID", uid
    Call Record(Len(SaveActivation(mid, ActivationCodeFor(mid), "TEST-PC")) > 0, "«· ›⁄Ì· ·„œÌ— «·‰Ÿ«„ ›ﬁÿ")
    Call Record(MachineID() Like "????-????-????-????", "—ﬁ„ Â–« «·ÃÂ«“ " & MachineID())

    ws.Rollback
    inTrans = False
    TempVars.Add "UserID", adminID
    Call Record(DCount("*", "Employees", "Username = 'test_screens' OR Username = 'test_dev'") = 0 And _
                DCount("*", "Activations", "MachineID = '" & mid & "'") = 0, "«· —«Ã⁄ ⁄‰ »Ì«‰«  «Œ »«— «·‘«‘«  Ê«· ›⁄Ì·")
    Exit Sub
EH:
    Fail "«·‘«‘«  Ê«· ›⁄Ì·: Œÿ√ " & Err.Number & ": " & Err.Description
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
