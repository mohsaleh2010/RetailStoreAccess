Attribute VB_Name = "modStartup"
'==============================================================================
' modStartup  -  Retail Store Management System (Phase 5)
'
' Application start-up, hiding the Access interface, and screen navigation.
'
'   InstallUserMode        end-user look: opens on frmLogin, no ribbon,
'                          no navigation pane, no design shortcuts.
'                          Close and reopen the file to see it.
'   InstallDeveloperMode   restores the full Access interface.
'   User mode also disables the SHIFT bypass key (AllowBypassKey). To get back:
'   frmBackup > "Ê÷⁄ «·„ÿÊ¯—" (administrator), or dist\tools\EnableShiftKey.vbs.
'==============================================================================
Option Compare Database
Option Explicit

Private Const PROP_DEV_MODE As String = "RS_DeveloperMode"

'------------------------------------------------------------------------------
' Called by frmMain when it opens
'------------------------------------------------------------------------------
Public Sub AppStartup()
    Calendar = vbCalGreg
    On Error Resume Next
    Application.SetOption "Use Hijri Calendar", False
    On Error GoTo 0
    If Not IsDeveloperMode() Then HideAccessUI
End Sub

Public Function IsDeveloperMode() As Boolean
    On Error Resume Next
    IsDeveloperMode = True
    IsDeveloperMode = CurrentDb.Properties(PROP_DEV_MODE)
End Function

Public Sub HideAccessUI()
    On Error Resume Next
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    DoCmd.NavigateTo "acNavigationCategoryObjectType"
    DoCmd.RunCommand acCmdWindowHide
End Sub

Public Sub ShowAccessUI()
    On Error Resume Next
    DoCmd.ShowToolbar "Ribbon", acToolbarYes
    DoCmd.SelectObject acForm, "frmMain", True
End Sub

'------------------------------------------------------------------------------
' Start-up properties
'------------------------------------------------------------------------------
Public Sub InstallUserMode()
    ApplyStartupProperties False
    MsgBox " „ ÷»ÿ Ê÷⁄ «·„” Œœ„ «·‰Â«∆Ì." & vbCrLf & vbCrLf & _
           "√€·ﬁ «·„·› Ê«› ÕÂ „‰ ÃœÌœ ·ÌŸÂ— «·‰Ÿ«„ »œÊ‰ √œÊ«  Access." & vbCrLf & _
           "··⁄Êœ… ·Ê÷⁄ «· ÿÊÌ—: «› Õ «·„·› „⁄ «·÷€ÿ ⁄·Ï „› «Õ SHIFT À„ ‘€¯· InstallDeveloperMode.", _
           vbInformation + MSG_RTL, APP_TITLE
End Sub

Public Sub InstallDeveloperMode()
    ApplyStartupProperties True
    ShowAccessUI
    MsgBox " „ ÷»ÿ Ê÷⁄ «· ÿÊÌ—: ﬂ· √œÊ«  Access Ÿ«Â—….", vbInformation + MSG_RTL, APP_TITLE
End Sub

Private Sub ApplyStartupProperties(ByVal Developer As Boolean)
    Dim db As DAO.Database
    Set db = CurrentDb
    SetDbProp db, "AppTitle", dbText, APP_TITLE
    SetDbProp db, "StartupForm", dbText, "frmLogin"
    SetDbProp db, "AllowBypassKey", dbBoolean, Developer   ' user mode: SHIFT no longer skips the start-up
    SetDbProp db, "StartupShowDBWindow", dbBoolean, Developer
    SetDbProp db, "StartupShowStatusBar", dbBoolean, True
    SetDbProp db, "AllowBuiltInToolbars", dbBoolean, Developer
    SetDbProp db, "AllowFullMenus", dbBoolean, Developer
    SetDbProp db, "AllowShortcutMenus", dbBoolean, Developer
    SetDbProp db, "AllowToolbarChanges", dbBoolean, Developer
    SetDbProp db, "AllowSpecialKeys", dbBoolean, Developer
    SetDbProp db, "UseMDIMode", dbByte, 0                ' tabbed documents
    SetDbProp db, "ShowDocumentTabs", dbBoolean, Developer
    SetDbProp db, PROP_DEV_MODE, dbBoolean, Developer
    Application.RefreshTitleBar
End Sub

Private Sub SetDbProp(ByVal db As DAO.Database, ByVal PropName As String, _
                      ByVal PropType As Integer, ByVal PropValue As Variant)
    Dim prp As DAO.Property
    On Error Resume Next
    Set prp = db.Properties(PropName)
    On Error GoTo 0
    If prp Is Nothing Then
        db.Properties.Append db.CreateProperty(PropName, PropType, PropValue)
    Else
        prp.Value = PropValue
    End If
End Sub

'------------------------------------------------------------------------------
' Navigation
'------------------------------------------------------------------------------
Public Function OpenScreen(ByVal FormName As String, Optional ByVal PhaseNo As Integer = 0, _
                           Optional ByVal RecordID As Variant) As Boolean
    If IsMissing(RecordID) Then RecordID = Null
    If Not FormExists(FormName) Then
        If PhaseNo > 0 Then
            ShowInfo "Â–Â «·‘«‘… ” ﬂÊ‰ „ «Õ… »⁄œ  ‰›Ì– «·„—Õ·… " & PhaseNo & "."
        Else
            ShowError "«·‘«‘… €Ì— „ÊÃÊœ…: " & FormName
        End If
        Exit Function
    End If
    If Not CanOpenScreen(FormName) Then Exit Function          ' permission of the user's role
    If IsFormOpen(FormName) Then DoCmd.Close acForm, FormName
    DoCmd.OpenForm FormName, acNormal, , , , acWindowNormal, RecordID
    OpenScreen = True
End Function

Public Sub ExitApplication()
    If AskYesNo("Â·  —Ìœ «·Œ—ÊÃ „‰ «·‰Ÿ«„ø") Then
        OfferBackupOnExit
        LogAction "LOGOUT"
        Application.Quit acQuitSaveNone
    End If
End Sub
