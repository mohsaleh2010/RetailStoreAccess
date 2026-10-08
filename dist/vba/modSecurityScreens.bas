Attribute VB_Name = "modSecurityScreens"
'==============================================================================
' modSecurityScreens  -  Retail Store Management System (Phase 10)
'
' Behaviour of the screens:
'   frmLogin            login (start-up form in user mode)
'   frmMain             MainOpen: nobody logged in, or this computer is not activated -> login;
'                       nav buttons follow permissions
'   frmChangePassword   own password, forced change, or set by the administrator
'   frmUsers            users (data screen; extra buttons and status line here)
'   frmRoles            permissions of each role (lines in local table tmpRolePermissions)
'   frmUserScreens      screens of one user, with add / edit / delete in each (tmpUserScreens)
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
        MaximizeAccessWindow                 ' full screen from the first window
        If Not IsDeveloperMode() Then HideAccessUI
        TempVars.Remove "UserID"
        On Error GoTo 0
    End If
    frm!lblStoreName.Caption = Tr(Nz(SettingValue("StoreName"), APP_TITLE))
    frm!txtUsername.Value = GetSetting("RetailStore", "Login", "LastUser", "")
    If IsNull(DbValue("SELECT PasswordHash FROM Employees WHERE EmployeeID = 1")) Then
        frm!lblMessage.Caption = Tr("«·œŒÊ· «·√Ê·: «”„ «·„” Œœ„ admin »œÊ‰ ﬂ·„… „—Ê—° Ê”Ìıÿ·» „‰ﬂ  ⁄ÌÌ‰Â«.")
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
        frm!lblMessage.Caption = Tr("«ﬂ » «”„ «·„” Œœ„.")
        SafeFocus frm!txtUsername
        Exit Function
    End If
    msg = LoginUser(user, Nz(frm!txtPassword.Value, ""), id)
    frm!txtPassword.Value = Null
    If Len(msg) > 0 Then
        frm!lblMessage.Caption = Tr(msg)
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
    If Not IsDeveloper() And Not IsActivated() Then  ' copy protection (modActivation)
        If IsAdministrator() Then
            DoCmd.OpenForm "frmActivation", acNormal, , , , acDialog, "LOGIN"
        Else
            ShowWarning "«·»—‰«„Ã €Ì— „›⁄¯· ⁄·Ï Â–« «·ÃÂ«“." & vbCrLf & "ÌÃ» √‰ Ì›⁄¯·Â „œÌ— «·‰Ÿ«„ ﬁ»· «·«” Œœ«„."
        End If
        If Not IsActivated() Then
            LogoutUser False
            Exit Function
        End If
    End If
    DoCmd.OpenForm "frmMain"
    MaximizeScreen "frmMain"
End Function

Public Sub LoginExit()
    Application.Quit acQuitSaveNone
End Sub

Public Function MainOpen(ByVal frm As Access.Form) As Boolean
    ' frmMain opens only for a logged-in user (tests run as the administrator).
    If g_SilentMode Then
        MainOpen = True
    ElseIf CurrentUserID() > 0 And (IsDeveloper() Or IsActivated()) Then
        MainOpen = True
    Else
        DoCmd.OpenForm "frmLogin"
    End If
End Function

Public Sub ApplyNavPermissions(ByVal frm As Access.Form)
    ' Side-menu buttons and dashboard tiles of screens the user may not open are disabled;
    ' such a tile is also shown in grey.
    Dim ctl As Access.Control, allowed As Boolean
    For Each ctl In frm.Controls
        If (Left$(ctl.Name, 6) = "btnNav" Or Left$(ctl.Name, 7) = "btnTile") And Len(ctl.Tag) > 0 Then
            allowed = CanOpenScreen(ctl.Tag, True)
            ctl.Enabled = allowed
            If Left$(ctl.Name, 7) = "btnTile" And Not allowed Then
                frm.Controls("boxNav" & Mid$(ctl.Name, 8)).BackColor = RGB(205, 210, 218)
            ElseIf Left$(ctl.Name, 6) = "btnNav" And Not allowed Then
                ' the side-menu button is transparent: its caption and icon show it is closed
                frm.Controls("lblNav" & Mid$(ctl.Name, 7)).ForeColor = RGB(120, 134, 156)
                frm.Controls("ico" & Mid$(ctl.Name, 7)).ForeColor = RGB(120, 134, 156)
            End If
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
    TempVars.Remove "RecurringAlertShown"
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
        frm!lblFor.Caption = Tr(" ⁄ÌÌ‰ ﬂ·„… „—Ê—: " & Nz(DLookup("EmployeeName & ' (' & Username & ')'", "Employees", _
                             "EmployeeID = " & target), ""))
        frm!txtOld.Visible = False
        frm!lblOld.Visible = False
        frm!chkMustChange.Value = True
    Else
        SafeFocus frm!txtNew
        frm!lblFor.Caption = Tr(IIf(forced, "«Œ — ﬂ·„… „—Ê— ÃœÌœ… ··„ «»⁄…: ", " €ÌÌ— ﬂ·„… «·„—Ê—: ") & CurrentUserName())
        frm!txtOld.Visible = Not forced And Not IsNull(DbValue("SELECT PasswordHash FROM Employees WHERE EmployeeID = " & _
                                                              CurrentUserID()))
        frm!lblOld.Visible = frm!txtOld.Visible
        frm!chkMustChange.Visible = False
        frm!lblMustChange.Visible = False
    End If
    frm!lblRules.Caption = Tr(MIN_PASSWORD_LENGTH & " √Õ—› ⁄·Ï «·√ﬁ·° Ê·«  ”«ÊÌ «”„ «·„” Œœ„")
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
    frm!lblPasswordState.Caption = Tr(state)
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
    CurrentDb.Execute Tr("INSERT INTO tmpRolePermissions (PermissionKey, PermissionName, ModuleName, SortOrder, Granted) " & _
        "SELECT p.PermissionKey, p.PermissionName, p.ModuleName, p.SortOrder, " & _
        "IIf(r.PermissionKey Is Null, False, True) FROM [@Permissions] AS p LEFT JOIN " & _
        "(SELECT PermissionKey FROM RolePermissions WHERE RoleID = " & roleID & ") AS r " & _
        "ON p.PermissionKey = r.PermissionKey"), dbFailOnError      ' the names in the interface language
    frm!subPermissions.Form.Requery
    frm!lblRoleInfo.Caption = Tr(Nz(DLookup("Description", "Roles", "RoleID = " & roleID), " ") & "   (" & _
        DCount("*", "Employees", "RoleID = " & roleID & " AND IsActive = True AND IsDeveloper = False") & " „” Œœ„)")
    locked = (roleID = ADMIN_ROLE_ID)
    frm!subPermissions.Form.AllowEdits = Not locked
    SafeFocus frm!cboRole
    frm!btnSaveRole.Enabled = Not locked
    frm!btnAll.Enabled = Not locked
    frm!btnNone.Enabled = Not locked
    frm!lblLockedNote.Caption = Tr(IIf(locked, "œÊ— „œÌ— «·‰Ÿ«„ Ì„·ﬂ ﬂ· «·’·«ÕÌ«  œ«∆„« Ê·« Ì„ﬂ‰  ﬁÌÌœÂ.", " "))
End Sub

Public Sub RoleSelectAll(ByVal frm As Access.Form, ByVal Granted As Boolean)
    CurrentDb.Execute "UPDATE tmpRolePermissions SET Granted = " & IIf(Granted, "True", "False"), dbFailOnError
    frm!subPermissions.Form.Requery
End Sub

Public Function SaveRolePermissions(ByVal frm As Access.Form) As Boolean
    Dim rs As DAO.Recordset, keys As String, msg As String
    If Not CanScreenAction(frm.Name, "EDIT") Then Exit Function      ' frmUserScreens
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
' Screens of a user (frmUserScreens + frmUserScreenLines on the local table tmpUserScreens)
' A user with "’·«ÕÌ«  ‘«‘«  Œ«’…" (Employees.CustomScreens) opens only the screens of his
' rows in UserScreens, with the add / edit / delete chosen there; any other user follows his role.
'==============================================================================
Public Sub UserScreensLoad(ByVal frm As Access.Form)
    EnsureLocalTables
    frm!cboUser.Value = Nz(DbValue("SELECT TOP 1 EmployeeID FROM Employees WHERE IsDeveloper = False AND RoleID <> " & _
                                   ADMIN_ROLE_ID & " ORDER BY EmployeeName, EmployeeID"), frm!cboUser.ItemData(0))
    UserScreensPicked frm
End Sub

Public Sub UserScreensPicked(ByVal frm As Access.Form)
    Dim uid As Long, custom As Boolean
    uid = Nz(frm!cboUser.Value, 0)
    custom = UsesCustomScreens(uid)
    FillUserScreens uid, Not custom
    frm!chkCustom.Value = custom
    ShowUserScreensState frm
End Sub

Public Sub UserScreensCustomChanged(ByVal frm As Access.Form)
    ' switched on: the grid starts from what the user has now; switched off: back to the role
    If Not Nz(frm!chkCustom.Value, False) Then FillUserScreens Nz(frm!cboUser.Value, 0), True
    ShowUserScreensState frm
End Sub

Public Sub UserScreensFromRole(ByVal frm As Access.Form)
    FillUserScreens Nz(frm!cboUser.Value, 0), True
    ShowUserScreensState frm
End Sub

Public Sub UserScreensAll(ByVal frm As Access.Form, ByVal Granted As Boolean)
    Dim g As String
    g = IIf(Granted, "True", "False")
    CurrentDb.Execute "UPDATE tmpUserScreens SET CanOpen = " & g & ", CanAdd = " & g & " AND HasAdd, CanEdit = " & g & _
                      " AND HasEdit, CanDelete = " & g & " AND HasDelete", dbFailOnError
    frm!subScreens.Form.Requery
End Sub

Public Sub UserScreenLineChanged(ByVal sf As Access.Form, ByVal FieldName As String)
    ' An action the screen does not have stays off; an action opens the screen; closing the screen
    ' removes its actions.
    If FieldName = "CanOpen" Then
        If Not Nz(sf!CanOpen.Value, False) Then
            sf!CanAdd.Value = False
            sf!CanEdit.Value = False
            sf!CanDelete.Value = False
        End If
    ElseIf Nz(sf(FieldName).Value, False) Then
        If Not Nz(sf("Has" & Mid$(FieldName, 4)).Value, False) Then
            sf(FieldName).Value = False
            ShowInfo "Â–« «·≈Ã—«¡ €Ì— „ÊÃÊœ ›Ì ‘«‘… ´" & sf!ScreenTitle.Value & "ª."
        Else
            sf!CanOpen.Value = True
        End If
    End If
    sf.Dirty = False
End Sub

Public Sub FillUserScreens(ByVal EmployeeID As Long, ByVal FromRole As Boolean)
    ' tmpUserScreens: every screen with what the user may do there - his own rows (UserScreens), or
    ' what his role gives (the role permission of the screen, and every action of an opened screen).
    Dim db As DAO.Database, roleID As Long
    Set db = CurrentDb
    roleID = Nz(DbValue("SELECT RoleID FROM Employees WHERE EmployeeID = " & EmployeeID), 0)
    db.Execute "DELETE FROM tmpUserScreens", dbFailOnError
    db.Execute Tr("INSERT INTO tmpUserScreens (ScreenName, ScreenTitle, ModuleName, SortOrder, HasAdd, HasEdit, " & _
        "HasDelete, CanOpen, CanAdd, CanEdit, CanDelete) SELECT s.ScreenName, s.ScreenTitle, s.ModuleName, s.SortOrder, " & _
        "s.HasAdd, s.HasEdit, s.HasDelete, False, False, False, False FROM [@Screens] AS s"), dbFailOnError
    If Not FromRole Then
        db.Execute "UPDATE tmpUserScreens AS t INNER JOIN UserScreens AS u ON t.ScreenName = u.ScreenName " & _
            "SET t.CanOpen = u.CanOpen, t.CanAdd = u.CanAdd AND t.HasAdd, t.CanEdit = u.CanEdit AND t.HasEdit, " & _
            "t.CanDelete = u.CanDelete AND t.HasDelete WHERE u.EmployeeID = " & EmployeeID, dbFailOnError
    Else
        If roleID = ADMIN_ROLE_ID Then
            db.Execute "UPDATE tmpUserScreens SET CanOpen = True", dbFailOnError
        Else
            db.Execute "UPDATE tmpUserScreens AS t INNER JOIN Screens AS s ON t.ScreenName = s.ScreenName " & _
                "SET t.CanOpen = True WHERE s.PermissionKey Is Null OR s.PermissionKey IN " & _
                "(SELECT PermissionKey FROM RolePermissions WHERE RoleID = " & roleID & ")", dbFailOnError
        End If
        db.Execute "UPDATE tmpUserScreens SET CanAdd = CanOpen AND HasAdd, CanEdit = CanOpen AND HasEdit, " & _
                   "CanDelete = CanOpen AND HasDelete", dbFailOnError
    End If
    db.Execute "UPDATE tmpUserScreens SET ActionsNote = IIf(HasAdd Or HasEdit Or HasDelete, Mid(IIf(HasAdd, " & _
        "'° ≈÷«›…', '') & IIf(HasEdit, '°  ⁄œÌ·', '') & IIf(HasDelete, '° Õ–›', ''), 3), '› Õ Ê⁄—÷ ›ﬁÿ')", dbFailOnError
End Sub

Private Sub ShowUserScreensState(ByVal frm As Access.Form)
    Dim uid As Long, isAdmin As Boolean, custom As Boolean
    uid = Nz(frm!cboUser.Value, 0)
    isAdmin = (Nz(DbValue("SELECT RoleID FROM Employees WHERE EmployeeID = " & uid), 0) = ADMIN_ROLE_ID)
    custom = Nz(frm!chkCustom.Value, False) And Not isAdmin
    SafeFocus frm!cboUser                         ' a button that has the focus cannot be disabled
    frm!subScreens.Form.Requery
    frm!subScreens.Form.AllowEdits = custom
    frm!chkCustom.Enabled = Not isAdmin
    frm!btnFromRole.Enabled = custom
    frm!btnAll.Enabled = custom
    frm!btnNone.Enabled = custom
    frm!btnSaveScreens.Enabled = Not isAdmin
    frm!lblUserInfo.Caption = Tr("«·œÊ—: " & Nz(DbValue(Tr("SELECT r.RoleName FROM Employees AS e INNER JOIN [@Roles] AS r " & _
                              "ON e.RoleID = r.RoleID WHERE e.EmployeeID = " & uid)), "-"))
    If isAdmin Then
        frm!lblNote.Caption = Tr("„œÌ— «·‰Ÿ«„ Ì› Õ ﬂ· «·‘«‘«  »ﬂ· «·’·«ÕÌ«  œ«∆„«.")
    ElseIf custom Then
        frm!lblNote.Caption = Tr("Õœœ «·‘«‘«  Ê«·≈Ã—«¡«  À„ «÷€ÿ ´Õ›Ÿª. «·’·«ÕÌ«  «·Œ«’… ( ﬁ«—Ì— «·√—»«Õ°  ⁄œÌ· «·”⁄—° " & _
                              "«·Œ’„...)  »ﬁÏ „‰ œÊ— «·„” Œœ„ ›Ì ‘«‘… «·√œÊ«—.")
    Else
        frm!lblNote.Caption = Tr("«·„” Œœ„ Ì »⁄ ’·«ÕÌ«  œÊ—Â ﬂ„« ›Ì «·ﬁ«∆„…. ›⁄¯· ´’·«ÕÌ«  ‘«‘«  Œ«’…ª · ÕœÌœ ‘«‘« Â ÊÕœÂ.")
    End If
End Sub

Public Function SaveUserScreens(ByVal frm As Access.Form) As Boolean
    Dim msg As String
    If Not CanScreenAction(frm.Name, "EDIT") Then Exit Function      ' frmUserScreens
    If frm!subScreens.Form.Dirty Then frm!subScreens.Form.Dirty = False
    msg = SaveUserScreensFor(Nz(frm!cboUser.Value, 0), Nz(frm!chkCustom.Value, False))
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    ShowUserScreensState frm
    ShowInfo " „ Õ›Ÿ ’·«ÕÌ«  «·‘«‘« .  ”—Ì „‰ «·œŒÊ· «· «·Ì ··„” Œœ„."
    SaveUserScreens = True
End Function

Public Function SaveUserScreensFor(ByVal EmployeeID As Long, ByVal Custom As Boolean) As String
    ' Saves the grid (tmpUserScreens) for the user; Custom = False: back to the screens of his role.
    Dim ws As DAO.Workspace, db As DAO.Database, inTrans As Boolean, roleID As Variant
    On Error GoTo EH
    If Not HasPermission("USERS") Then
        SaveUserScreensFor = "·«  „·ﬂ ’·«ÕÌ… ≈œ«—… «·„” Œœ„Ì‰."
        Exit Function
    End If
    roleID = DbValue("SELECT RoleID FROM Employees WHERE EmployeeID = " & EmployeeID & " AND IsDeveloper = False")
    If IsNull(roleID) Then
        SaveUserScreensFor = "«Œ — «·„” Œœ„."
    ElseIf roleID = ADMIN_ROLE_ID Then
        SaveUserScreensFor = "„œÌ— «·‰Ÿ«„ Ì› Õ ﬂ· «·‘«‘«  »ﬂ· «·’·«ÕÌ«  œ«∆„«."
    ElseIf EmployeeID = CurrentUserID() Then
        SaveUserScreensFor = "·« Ì„ﬂ‰ﬂ  €ÌÌ— ’·«ÕÌ« ﬂ »‰›”ﬂ."
    ElseIf Custom And Nz(DbValue("SELECT COUNT(*) FROM tmpUserScreens WHERE CanOpen = True"), 0) = 0 Then
        SaveUserScreensFor = "«Œ — ‘«‘… Ê«Õœ… ⁄·Ï «·√ﬁ·° √Ê √·€ˆ ´’·«ÕÌ«  ‘«‘«  Œ«’…ª."
    End If
    If Len(SaveUserScreensFor) > 0 Then Exit Function
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    db.Execute "DELETE FROM UserScreens WHERE EmployeeID = " & EmployeeID, dbFailOnError
    If Custom Then
        db.Execute "INSERT INTO UserScreens (EmployeeID, ScreenName, CanOpen, CanAdd, CanEdit, CanDelete) SELECT " & _
            EmployeeID & ", ScreenName, CanOpen, CanOpen AND CanAdd AND HasAdd, CanOpen AND CanEdit AND HasEdit, " & _
            "CanOpen AND CanDelete AND HasDelete FROM tmpUserScreens", dbFailOnError
    End If
    db.Execute "UPDATE Employees SET CustomScreens = " & IIf(Custom, "True", "False") & " WHERE EmployeeID = " & _
               EmployeeID, dbFailOnError
    ws.CommitTrans
    inTrans = False
    LogAction "USER_SCREENS", "Employees", CStr(EmployeeID), IIf(Custom, "custom", "role")
    Exit Function
EH:
    SaveUserScreensFor = " ⁄–— «·Õ›Ÿ: " & Err.Description
    If inTrans Then ws.Rollback
End Function

'==============================================================================
' Backup screen
'==============================================================================
Public Sub BackupLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    frm!lblDataFile.Caption = Tr("„·› «·»Ì«‰« : " & BackendFilePath())
    BackupRefreshList frm
End Sub

Public Sub BackupRefreshList(ByVal frm As Access.Form)
    Dim folder As String, names As String, v As Variant, rows As String, i As Long, arr() As String
    Dim lastDate As Variant
    folder = BackupFolderPath()
    frm!lblFolder.Caption = Tr("«·„Ã·œ: " & folder)
    lastDate = LastBackupDate()
    frm!lblLast.Caption = Tr("¬Œ— ‰”Œ…: " & IIf(IsNull(lastDate), "·«  ÊÃœ", GDate(lastDate, True)) & _
        "    ÌıÕ ›Ÿ »¬Œ— " & Nz(SettingValue("BackupKeepCount"), 30) & " ‰”Œ…")
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
    frm!lstBackups.RowSource = Tr(rows)
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
