Attribute VB_Name = "modSecurityScreens"
'==============================================================================
' modSecurityScreens  -  Retail Store Management System (Phase 10)
'
' Behaviour of the screens:
'   frmLogin            login (start-up form in user mode)
'   frmMain             MainOpen: nobody logged in -> login; nav buttons follow permissions
'   frmChangePassword   own password, forced change, or set by the administrator
'   frmUsers            users (data screen; extra buttons and status line here)
'   frmRoles            permissions of each role (lines in local table tmpRolePermissions)
'   frmBackup           backup / restore
' Logic is in modSecurity and modBackup.
'==============================================================================
Option Compare Database
Option Explicit

'==============================================================================
' Login / logout
'==============================================================================
Public Sub LoginLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    If Not g_SilentMode Then                 ' (tests open every screen while logged in)
        On Error Resume Next
        If Not IsDeveloperMode() Then HideAccessUI
        TempVars.Remove "UserID"
        On Error GoTo 0
    End If
    frm!lblStoreName.Caption = Nz(SettingValue("StoreName"), APP_TITLE)
    frm!txtUsername.Value = GetSetting("RetailStore", "Login", "LastUser", "")
    If IsNull(DbValue("SELECT PasswordHash FROM Employees WHERE EmployeeID = 1")) And _
       IsNull(DbValue("SELECT LastLoginAt FROM Employees WHERE EmployeeID = 1")) Then
        frm!lblMessage.Caption = "«·œŒÊ· «·√Ê·: «”„ «·„” Œœ„ admin »œÊ‰ ﬂ·„… „—Ê—° Ê”Ìıÿ·» „‰ﬂ  ⁄ÌÌ‰Â«."
        frm!lblMessage.ForeColor = CLR_PRIMARY
        frm!txtUsername.Value = "admin"
    End If
    If Len(Nz(frm!txtUsername.Value, "")) > 0 Then
        SafeFocus frm!txtPassword
    Else
        SafeFocus frm!txtUsername
    End If
End Sub

Public Sub LoginKeyDown(ByVal frm As Access.Form, ByRef KeyCode As Integer, ByVal IsPassword As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    KeyCode = 0
    If IsPassword Then
        DoLogin frm
    Else
        SafeFocus frm!txtPassword
    End If
End Sub

Public Function DoLogin(ByVal frm As Access.Form) As Boolean
    Dim msg As String, id As Long, user As String
    user = Trim$(Nz(frm!txtUsername.Value, ""))
    If Len(user) = 0 Then
        frm!lblMessage.Caption = "«ﬂ » «”„ «·„” Œœ„."
        SafeFocus frm!txtUsername
        Exit Function
    End If
    msg = LoginUser(user, Nz(frm!txtPassword.Value, ""), id)
    frm!txtPassword.Value = Null
    If Len(msg) > 0 Then
        frm!lblMessage.Caption = msg
        frm!lblMessage.ForeColor = CLR_DANGER
        SafeFocus frm!txtPassword
        Exit Function
    End If
    SaveSetting "RetailStore", "Login", "LastUser", user
    DoLogin = True
    If g_SilentMode Then Exit Function          ' tests stop here
    DoCmd.Close acForm, frm.Name
    If NeedsPasswordChange(id) Then
        DoCmd.OpenForm "frmChangePassword", acNormal, , , , acDialog, "FORCE"
        If NeedsPasswordChange(id) Then          ' closed without choosing a password
            LogoutUser False
            Exit Function
        End If
    End If
    DoCmd.OpenForm "frmMain"
End Function

Public Sub LoginExit()
    Application.Quit acQuitSaveNone
End Sub

Public Function MainOpen(ByVal frm As Access.Form) As Boolean
    ' frmMain opens only for a logged-in user (tests run as the administrator).
    If CurrentUserID() > 0 Or g_SilentMode Then
        MainOpen = True
    Else
        DoCmd.OpenForm "frmLogin"
    End If
End Function

Public Sub ApplyNavPermissions(ByVal frm As Access.Form)
    ' Side-menu buttons of screens the user may not open are disabled.
    Dim ctl As Access.Control
    For Each ctl In frm.Controls
        If Left$(ctl.Name, 6) = "btnNav" And Len(ctl.Tag) > 0 Then
            ctl.Enabled = CanOpenScreen(ctl.Tag, True)
        End If
    Next
End Sub

Public Sub LogoutUser(Optional ByVal Ask As Boolean = True)
    Dim i As Long
    If Ask Then
        If Not AskYesNo("Â·  —Ìœ  ”ÃÌ· «·Œ—ÊÃø") Then Exit Sub
        OfferBackupOnExit
    End If
    LogAction "LOGOUT"
    For i = Forms.Count - 1 To 0 Step -1
        DoCmd.Close acForm, Forms(i).Name, acSaveNo
    Next
    On Error Resume Next
    TempVars.Remove "UserID"
    TempVars.Remove "LowStockAlertShown"
    On Error GoTo 0
    DoCmd.OpenForm "frmLogin"
End Sub

'==============================================================================
' Password
' OpenArgs: Null = own password; "FORCE" = must choose one now; EmployeeID = set by admin
'==============================================================================
Public Sub ChangePasswordLoad(ByVal frm As Access.Form)
    Dim target As Long, forced As Boolean
    forced = (Nz(frm.OpenArgs, "") = "FORCE")
    If IsNumeric(Nz(frm.OpenArgs, "")) Then target = CLng(frm.OpenArgs)
    If target > 0 And target <> CurrentUserID() Then
        If Not HasPermission("USERS") Then
            ShowWarning "·«  „·ﬂ ’·«ÕÌ…  ⁄ÌÌ‰ ﬂ·„«  «·„—Ê—."
            frm!btnSave.Enabled = False
            Exit Sub
        End If
        SafeFocus frm!txtNew                       ' a focused control cannot be hidden
        frm!lblFor.Caption = " ⁄ÌÌ‰ ﬂ·„… „—Ê—: " & Nz(DLookup("EmployeeName & ' (' & Username & ')'", "Employees", _
                             "EmployeeID = " & target), "")
        frm!txtOld.Visible = False
        frm!lblOld.Visible = False
        frm!chkMustChange.Value = True
    Else
        SafeFocus frm!txtNew
        frm!lblFor.Caption = IIf(forced, "«Œ — ﬂ·„… „—Ê— ÃœÌœ… ··„ «»⁄…: ", " €ÌÌ— ﬂ·„… «·„—Ê—: ") & CurrentUserName()
        frm!txtOld.Visible = Not forced And Not IsNull(DbValue("SELECT PasswordHash FROM Employees WHERE EmployeeID = " & _
                                                              CurrentUserID()))
        frm!lblOld.Visible = frm!txtOld.Visible
        frm!chkMustChange.Visible = False
        frm!lblMustChange.Visible = False
    End If
    frm!lblRules.Caption = MIN_PASSWORD_LENGTH & " √Õ—› ⁄·Ï «·√ﬁ·° Ê·«  ”«ÊÌ «”„ «·„” Œœ„"
End Sub

Public Function SaveChangedPassword(ByVal frm As Access.Form) As Boolean
    Dim msg As String, target As Long
    If IsNumeric(Nz(frm.OpenArgs, "")) Then target = CLng(frm.OpenArgs)
    If target > 0 And target <> CurrentUserID() Then
        If Nz(frm!txtNew.Value, "") <> Nz(frm!txtConfirm.Value, "") Then
            msg = " √ﬂÌœ ﬂ·„… «·„—Ê— ·« Ìÿ«»ﬁ ﬂ·„… «·„—Ê— «·ÃœÌœ…."
        ElseIf Not HasPermission("USERS") Then
            msg = "·«  „·ﬂ ’·«ÕÌ…  ⁄ÌÌ‰ ﬂ·„«  «·„—Ê—."
        Else
            msg = SetUserPassword(target, Nz(frm!txtNew.Value, ""), Nz(frm!chkMustChange.Value, True))
        End If
    Else
        msg = ChangeOwnPassword(Nz(frm!txtOld.Value, ""), Nz(frm!txtNew.Value, ""), Nz(frm!txtConfirm.Value, ""), _
                                frm!txtOld.Visible)
    End If
    frm!txtNew.Value = Null
    frm!txtConfirm.Value = Null
    frm!txtOld.Value = Null
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    ShowInfo " „ Õ›Ÿ ﬂ·„… «·„—Ê—."
    SaveChangedPassword = True
    DoCmd.Close acForm, frm.Name
End Function

Public Sub CancelChangePassword(ByVal frm As Access.Form)
    DoCmd.Close acForm, frm.Name
End Sub

'==============================================================================
' Users screen (frmUsers, a data screen on Employees)
'==============================================================================
Public Sub UserCurrent(ByVal frm As Access.Form)
    Dim state As String
    If frm.NewRecord Then
        state = "»⁄œ «·Õ›Ÿ ⁄Ì¯‰ ﬂ·„… «·„—Ê— „‰ “— ´ﬂ·„… «·„—Ê—ª."
    ElseIf IsNull(frm!PasswordHash.Value) Then
        state = "·„  ı⁄ÌÛ¯‰ ﬂ·„… „—Ê— »⁄œ° Ê·« Ì” ÿÌ⁄ Â–« «·„” Œœ„ «·œŒÊ·."
    ElseIf Nz(frm!LockedUntil.Value, 0) > Now Then
        state = "«·Õ”«» „ﬁ›· Õ Ï " & Format$(frm!LockedUntil.Value, "hh:nn") & " »”»» „Õ«Ê·«  Œ«ÿ∆…."
    ElseIf frm!MustChangePassword.Value Then
        state = "ﬂ·„… „—Ê— „ƒﬁ …: ”Ìıÿ·»  €ÌÌ—Â« ⁄‰œ «·œŒÊ·."
    Else
        state = "ﬂ·„… «·„—Ê— „⁄Ì¯‰…."
    End If
    frm!lblPasswordState.Caption = state
End Sub

Public Sub UnlockUser(ByVal frm As Access.Form)
    If frm.NewRecord Or IsNull(frm!EmployeeID.Value) Then Exit Sub
    If Not HasPermission("USERS") Then
        ShowWarning "·«  „·ﬂ ’·«ÕÌ… ≈œ«—… «·„” Œœ„Ì‰."
        Exit Sub
    End If
    CurrentDb.Execute "UPDATE Employees SET FailedLoginCount = 0, LockedUntil = Null WHERE EmployeeID = " & _
                      frm!EmployeeID.Value, dbFailOnError
    LogAction "USER_UNLOCK", "Employees", CStr(frm!EmployeeID.Value)
    frm.Refresh
    UserCurrent frm
    ShowInfo " „ ›ﬂ ﬁ›· «·Õ”«»."
End Sub

'==============================================================================
' Roles and permissions
'==============================================================================
Public Sub RolesLoad(ByVal frm As Access.Form)
    EnsureLocalTables
    frm!cboRole.Value = 3                       ' cashier: the role most often adjusted
    RolePicked frm
End Sub

Public Sub RolePicked(ByVal frm As Access.Form)
    Dim roleID As Long, locked As Boolean
    roleID = Nz(frm!cboRole.Value, 0)
    CurrentDb.Execute "DELETE FROM tmpRolePermissions", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpRolePermissions (PermissionKey, PermissionName, ModuleName, SortOrder, Granted) " & _
        "SELECT p.PermissionKey, p.PermissionName, p.ModuleName, p.SortOrder, " & _
        "IIf(r.PermissionKey Is Null, False, True) FROM Permissions AS p LEFT JOIN " & _
        "(SELECT PermissionKey FROM RolePermissions WHERE RoleID = " & roleID & ") AS r " & _
        "ON p.PermissionKey = r.PermissionKey", dbFailOnError
    frm!subPermissions.Form.Requery
    frm!lblRoleInfo.Caption = Nz(DLookup("Description", "Roles", "RoleID = " & roleID), " ") & "   (" & _
        DCount("*", "Employees", "RoleID = " & roleID & " AND IsActive = True") & " „” Œœ„)"
    locked = (roleID = ADMIN_ROLE_ID)
    frm!subPermissions.Form.AllowEdits = Not locked
    SafeFocus frm!cboRole
    frm!btnSaveRole.Enabled = Not locked
    frm!btnAll.Enabled = Not locked
    frm!btnNone.Enabled = Not locked
    frm!lblLockedNote.Caption = IIf(locked, "œÊ— „œÌ— «·‰Ÿ«„ Ì„·ﬂ ﬂ· «·’·«ÕÌ«  œ«∆„« Ê·« Ì„ﬂ‰  ﬁÌÌœÂ.", " ")
End Sub

Public Sub RoleSelectAll(ByVal frm As Access.Form, ByVal Granted As Boolean)
    CurrentDb.Execute "UPDATE tmpRolePermissions SET Granted = " & IIf(Granted, "True", "False"), dbFailOnError
    frm!subPermissions.Form.Requery
End Sub

Public Function SaveRolePermissions(ByVal frm As Access.Form) As Boolean
    Dim rs As DAO.Recordset, keys As String, msg As String
    If frm!subPermissions.Form.Dirty Then frm!subPermissions.Form.Dirty = False
    Set rs = CurrentDb.OpenRecordset("SELECT PermissionKey FROM tmpRolePermissions WHERE Granted = True", dbOpenSnapshot)
    Do Until rs.EOF
        keys = keys & IIf(Len(keys) > 0, ",", "") & rs!PermissionKey
        rs.MoveNext
    Loop
    rs.Close
    msg = SaveRolePermissionSet(Nz(frm!cboRole.Value, 0), keys)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    ShowInfo " „ Õ›Ÿ ’·«ÕÌ«  «·œÊ—.  ”—Ì ⁄‰œ «·› Õ «· «·Ì ·ﬂ· ‘«‘…."
    SaveRolePermissions = True
End Function

Public Function SaveRolePermissionSet(ByVal RoleID As Long, ByVal KeyList As String) As String
    ' Replaces the permissions of a role with KeyList ("KEY1,KEY2"). The administrator role is fixed.
    Dim ws As DAO.Workspace, db As DAO.Database, k As Variant, inTrans As Boolean
    On Error GoTo EH
    If Not HasPermission("USERS") Then
        SaveRolePermissionSet = "·«  „·ﬂ ’·«ÕÌ… ≈œ«—… «·’·«ÕÌ« ."
        Exit Function
    End If
    If RoleID = ADMIN_ROLE_ID Then
        SaveRolePermissionSet = "œÊ— „œÌ— «·‰Ÿ«„ Ì„·ﬂ ﬂ· «·’·«ÕÌ«  œ«∆„«."
        Exit Function
    End If
    If IsNull(DbValue("SELECT RoleID FROM Roles WHERE RoleID = " & RoleID)) Then
        SaveRolePermissionSet = "«Œ — «·œÊ—."
        Exit Function
    End If
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    db.Execute "DELETE FROM RolePermissions WHERE RoleID = " & RoleID, dbFailOnError
    If Len(KeyList) > 0 Then
        For Each k In Split(KeyList, ",")
            db.Execute "INSERT INTO RolePermissions (RoleID, PermissionKey) VALUES (" & RoleID & ", " & _
                       SqlText(Trim$(CStr(k))) & ")", dbFailOnError
        Next
    End If
    ws.CommitTrans
    inTrans = False
    LogAction "ROLE_PERMISSIONS", "Roles", CStr(RoleID), KeyList
    Exit Function
EH:
    SaveRolePermissionSet = " ⁄–— «·Õ›Ÿ: " & Err.Description
    If inTrans Then ws.Rollback
End Function

'==============================================================================
' Backup screen
'==============================================================================
Public Sub BackupLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    frm!lblDataFile.Caption = "„·› «·»Ì«‰« : " & BackendFilePath()
    BackupRefreshList frm
End Sub

Public Sub BackupRefreshList(ByVal frm As Access.Form)
    Dim folder As String, names As String, v As Variant, rows As String, i As Long, arr() As String
    Dim lastDate As Variant
    folder = BackupFolderPath()
    frm!lblFolder.Caption = "«·„Ã·œ: " & folder
    lastDate = LastBackupDate()
    frm!lblLast.Caption = "¬Œ— ‰”Œ…: " & IIf(IsNull(lastDate), "·«  ÊÃœ", GDate(lastDate, True)) & _
        "    ÌıÕ ›Ÿ »¬Œ— " & Nz(SettingValue("BackupKeepCount"), 30) & " ‰”Œ…"
    names = BackupList(folder, BaseNameOfFile(BackendFilePath()))
    rows = """«·„·›"";""«· «—ÌŒ"";""«·ÕÃ„"""
    If Len(names) > 0 Then
        arr = Split(names, "|")
        For i = UBound(arr) To 0 Step -1         ' newest first (names sort by time)
            v = folder & "\" & arr(i)
            rows = rows & ";""" & arr(i) & """;""" & GDate(FileDateTime(v), True) & """;""" & _
                   Format$(FileLen(v) / 1048576, "0.0") & " MB"""
        Next
    End If
    frm!lstBackups.RowSource = rows
End Sub

Private Function BaseNameOfFile(ByVal FullPath As String) As String
    Dim f As String
    f = Mid$(FullPath, InStrRev(FullPath, "\") + 1)
    If InStrRev(f, ".") > 0 Then f = Left$(f, InStrRev(f, ".") - 1)
    BaseNameOfFile = f
End Function

Public Function BackupRun(ByVal frm As Access.Form) As Boolean
    Dim msg As String, file As String
    DoCmd.Hourglass True
    msg = BackupNow(file)
    DoCmd.Hourglass False
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        ShowInfo " „  «·‰”Œ… «·«Õ Ì«ÿÌ… »‰Ã«Õ Ê „ «· Õﬁﬁ „‰Â«:" & vbCrLf & file
        BackupRun = True
    End If
    BackupRefreshList frm
End Function

Public Sub BackupRestoreSelected(ByVal frm As Access.Form)
    Dim file As String, msg As String
    If IsNull(frm!lstBackups.Value) Then
        ShowWarning "«Œ — «·‰”Œ… „‰ «·ﬁ«∆„… √Ê·«."
        Exit Sub
    End If
    file = BackupFolderPath() & "\" & frm!lstBackups.Value
    If Not AskYesNo("” ı” »œ· ﬂ· «·»Ì«‰«  «·Õ«·Ì… »«·‰”Œ…:" & vbCrLf & frm!lstBackups.Value & vbCrLf & vbCrLf & _
                    "ﬂ· „« ”ıÃ¯· »⁄œ  «—ÌŒ Â–Â «·‰”Œ… ”Ìı›ﬁœ ( ıÕ›Ÿ ‰”Œ… Êﬁ«∆Ì… „‰ «·»Ì«‰«  «·Õ«·Ì… √Ê·«)." & _
                    vbCrLf & "ÌÃ» √‰ ÌﬂÊ‰ «·»—‰«„Ã „€·ﬁ« ⁄·Ï ﬂ· «·√ÃÂ“… «·√Œ—Ï. Â·  —Ìœ «·„ «»⁄…ø") Then Exit Sub
    If Not AskYesNo(" √ﬂÌœ √ŒÌ—: «” ⁄«œ… «·‰”Œ… " & frm!lstBackups.Value & "ø") Then Exit Sub
    DoCmd.Hourglass True
    msg = RestoreBackup(file)
    DoCmd.Hourglass False
    If Len(msg) > 0 Then
        ShowError msg
        Exit Sub
    End If
    ShowInfo " „  «·«” ⁄«œ…. ”Ìı€·ﬁ «·»—‰«„Ã «·¬‰° «› ÕÂ „‰ ÃœÌœ."
    Application.Quit acQuitSaveNone
End Sub

Public Sub BackupOpenFolder()
    Application.FollowHyperlink BackupFolderPath()
End Sub

Public Sub BackupChooseFolder(ByVal frm As Access.Form)
    If Not HasPermission("SETTINGS") And Not HasPermission("BACKUP") Then Exit Sub
    With Application.FileDialog(4)          ' msoFileDialogFolderPicker
        .Title = "«Œ — „Ã·œ «·‰”Œ «·«Õ Ì«ÿÌ (Ì›÷· ﬁ—’ ¬Œ— √Ê ›·«‘… √Ê „Ã·œ „ “«„‰ „⁄ «·”Õ«»…)"
        If .Show Then
            CurrentDb.Execute "UPDATE Settings SET BackupFolder = " & SqlText(.SelectedItems(1)) & _
                              " WHERE SettingID = 1", dbFailOnError
            LogAction "BACKUP_FOLDER", "Settings", "1", .SelectedItems(1)
        End If
    End With
    BackupRefreshList frm
End Sub

Public Sub DeveloperModeFromApp()
    ' Only the administrator role can bring back the Access interface (and the Shift key).
    If Not IsAdministrator() Then
        ShowWarning "Ê÷⁄ «·„ÿÊ¯— ·„œÌ— «·‰Ÿ«„ ›ﬁÿ."
        Exit Sub
    End If
    If Not AskYesNo("”Ìı⁄«œ ≈ŸÂ«— √œÊ«  Access ÊÌı”„Õ »„› «Õ Shift ⁄‰œ «·› Õ. Â·  —Ìœ «·„ «»⁄…ø") Then Exit Sub
    LogAction "DEVELOPER_MODE"
    InstallDeveloperMode
End Sub
