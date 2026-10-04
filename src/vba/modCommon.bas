Attribute VB_Name = "modCommon"
'==============================================================================
' modCommon  -  Retail Store Management System (Phase 5)
'
' Shared helpers used by every screen: theme, messages, numbering, rounding,
' SQL literals, settings, audit log and object checks.
' Hand-written module (not generated). Readable copy: src/vba/modCommon.bas
'==============================================================================
Option Compare Database
Option Explicit

'------------------------------------------------------------------------------
' Theme  (VBA colour = R + G*256 + B*65536; hex value in the comment)
'------------------------------------------------------------------------------
Public Const CLR_PRIMARY As Long = 6240799           ' #1F3A5F  navy - title bars, sidebar
Public Const CLR_PRIMARY_HOVER As Long = 8277803     ' #2B4F7E
Public Const CLR_PRIMARY_PRESSED As Long = 4664086   ' #162B47
Public Const CLR_ACCENT As Long = 7371790            ' #0E7C70  teal - main action (Save)
Public Const CLR_ACCENT_HOVER As Long = 9082386      ' #12968A
Public Const CLR_BACKGROUND As Long = 16315891       ' #F3F5F8  form background
Public Const CLR_SURFACE As Long = 16777215          ' #FFFFFF  inputs
Public Const CLR_BORDER As Long = 14537675           ' #CBD3DD
Public Const CLR_TEXT As Long = 2695966              ' #1E2329
Public Const CLR_MUTED As Long = 8022879             ' #5F6B7A  labels, hints
Public Const CLR_LOCKED As Long = 16118254           ' #EEF1F5  read-only inputs
Public Const CLR_DANGER As Long = 3029680            ' #B03A2E  delete
Public Const CLR_DANGER_HOVER As Long = 3819721      ' #C9483A
Public Const CLR_SUCCESS As Long = 3308846           ' #2E7D32
Public Const CLR_WARNING As Long = 23450             ' #9A5B00
Public Const CLR_SIDEBAR_TEXT As Long = 15918812     ' #DCE6F2
Public Const CLR_SECONDARY As Long = 15788516        ' #E4E9F0  secondary buttons
Public Const CLR_SECONDARY_HOVER As Long = 15129555  ' #D3DBE6

Public Const FONT_NAME As String = "Segoe UI"
Public Const ICON_FONT As String = "Segoe MDL2 Assets"   ' Windows 10/11 icon font
Public Const APP_TITLE As String = "نظام إدارة المحل"
Public Const MSG_RTL As Long = &H180000                  ' vbMsgBoxRight + vbMsgBoxRtlReading

'------------------------------------------------------------------------------
' Test support: when g_SilentMode is True no message box is shown, the text
' goes to g_LastMessage and questions return g_AutoAnswer.
'------------------------------------------------------------------------------
Public g_SilentMode As Boolean
Public g_AutoAnswer As Boolean
Public g_LastMessage As String

' RunAllTests (modTestAll) collects the result message of every test instead of
' showing one message box per test.
Public g_CollectTests As Boolean
Public g_TestSummary As String

'------------------------------------------------------------------------------
' Messages
'------------------------------------------------------------------------------
Public Sub ShowInfo(ByVal Text As String, Optional ByVal Title As String = "")
    ShowMessage Text, vbInformation, Title
End Sub

Public Sub ShowWarning(ByVal Text As String, Optional ByVal Title As String = "")
    ShowMessage Text, vbExclamation, Title
End Sub

Public Sub ShowError(ByVal Text As String, Optional ByVal Title As String = "")
    ShowMessage Text, vbCritical, Title
End Sub

Public Function AskYesNo(ByVal Text As String, Optional ByVal Title As String = "") As Boolean
    g_LastMessage = Text
    If g_SilentMode Then
        Debug.Print "[?] " & Text & " -> " & IIf(g_AutoAnswer, "نعم", "لا")
        AskYesNo = g_AutoAnswer
        Exit Function
    End If
    If Len(Title) = 0 Then Title = APP_TITLE
    AskYesNo = (MsgBox(Text, vbQuestion + vbYesNo + vbDefaultButton2 + MSG_RTL, Title) = vbYes)
End Function

Public Sub TestMsg(ByVal Text As String, ByVal Style As Long, Optional ByVal Title As String = "")
    ' The final message of an in-Access test (TestSales, TestForms, ...).
    If g_CollectTests Then
        g_TestSummary = g_TestSummary & "- " & Title & ": " & Replace(Left$(Text, 700), vbCrLf, " | ") & vbCrLf
    Else
        MsgBox Text, Style, Title
    End If
End Sub

Private Sub ShowMessage(ByVal Text As String, ByVal Icon As VbMsgBoxStyle, ByVal Title As String)
    g_LastMessage = Text
    If g_SilentMode Then
        Debug.Print "[i] " & Text
        Exit Sub
    End If
    If Len(Title) = 0 Then Title = APP_TITLE
    MsgBox Text, Icon + MSG_RTL, Title
End Sub

'------------------------------------------------------------------------------
' Current user: set by the login screen (modSecurity.LoginUser). 0 = nobody,
' so without a login no permission is granted.
'------------------------------------------------------------------------------
Public Function CurrentUserID() As Long
    Dim v As Variant
    On Error Resume Next
    v = TempVars("UserID")
    If IsNumeric(v) Then CurrentUserID = CLng(v)
End Function

Public Function CurrentUserName() As String
    CurrentUserName = Nz(DLookup("EmployeeName", "Employees", "EmployeeID = " & CurrentUserID()), "")
End Function

Public Function HasPermission(ByVal PermissionKey As String) As Boolean
    ' True when the role of the current user has the permission (table RolePermissions).
    Dim roleID As Variant
    ' DbValue reads through CurrentDb, so it also sees changes made inside an open transaction.
    roleID = DbValue("SELECT RoleID FROM Employees WHERE EmployeeID = " & CurrentUserID() & " AND IsActive = True")
    If IsNull(roleID) Then Exit Function
    HasPermission = Nz(DbValue("SELECT COUNT(*) FROM RolePermissions WHERE RoleID = " & roleID & _
                               " AND PermissionKey = " & SqlText(PermissionKey)), 0) > 0
End Function

'------------------------------------------------------------------------------
' Numbering: next value of a row in the Sequences table (INV-000001, P00001...)
' Call it inside the same transaction that saves the document.
'------------------------------------------------------------------------------
Public Function NextNumber(ByVal SequenceName As String) As String
    Dim rs As DAO.Recordset, n As Long, prefix As String, pad As Integer
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [Sequences] WHERE [SequenceName] = " & _
                                     SqlText(SequenceName), dbOpenDynaset)
    If rs.EOF Then
        rs.Close
        Err.Raise vbObjectError + 700, "NextNumber", "Sequence not found: " & SequenceName
    End If
    rs.Edit
    n = rs!NextValue
    prefix = Nz(rs!Prefix, "")
    pad = Nz(rs!PadLength, 0)
    rs!NextValue = n + 1
    rs.Update
    rs.Close
    If pad > 0 Then
        NextNumber = prefix & Format$(n, String$(pad, "0"))
    Else
        NextNumber = prefix & CStr(n)
    End If
End Function

'------------------------------------------------------------------------------
' Rounding half away from zero. VBA Round() uses banker's rounding:
' Round(2.345, 2) = 2.34, but invoices need 2.35.
'------------------------------------------------------------------------------
Public Function RoundMoney(ByVal Value As Variant, Optional ByVal Digits As Integer = 2) As Currency
    Dim factor As Variant, x As Variant
    factor = CDec(10 ^ Digits)
    x = CDec(Nz(Value, 0)) * factor
    If x >= 0 Then
        x = Fix(x + CDec(0.5))
    Else
        x = -Fix(-x + CDec(0.5))
    End If
    RoundMoney = CCur(x / factor)
End Function

'------------------------------------------------------------------------------
' SQL literals (always Gregorian, whatever the Windows calendar is)
'------------------------------------------------------------------------------
Public Function SqlDate(ByVal d As Date) As String
    Dim saved As Integer
    saved = Calendar
    Calendar = vbCalGreg
    SqlDate = "#" & Format$(d, "yyyy-mm-dd hh:nn:ss") & "#"
    Calendar = saved
End Function

Public Function SqlText(ByVal s As Variant) As String
    SqlText = "'" & Replace(Nz(s, ""), "'", "''") & "'"
End Function

Public Function LikePattern(ByVal s As String) As String
    ' Text for  Field Like '*...*'  with the Access wildcards escaped.
    s = Replace(s, "[", "[[]")
    s = Replace(s, "*", "[*]")
    s = Replace(s, "?", "[?]")
    s = Replace(s, "#", "[#]")
    LikePattern = "'*" & Replace(s, "'", "''") & "*'"
End Function

'------------------------------------------------------------------------------
' Settings and audit log
'------------------------------------------------------------------------------
Public Function SettingValue(ByVal FieldName As String) As Variant
    SettingValue = DbValue("SELECT [" & FieldName & "] FROM [Settings] WHERE [SettingID] = 1")
End Function

Public Function DbValue(ByVal Sql As String) As Variant
    ' First column of the first row, or Null. Unlike DLookup it reads through CurrentDb,
    ' so it also sees changes made earlier in the same (not yet committed) transaction.
    Dim rs As DAO.Recordset
    DbValue = Null
    Set rs = CurrentDb.OpenRecordset(Sql, dbOpenSnapshot)
    If Not rs.EOF Then DbValue = rs(0).Value
    rs.Close
End Function

Public Sub LogAction(ByVal ActionType As String, Optional ByVal ObjectName As String = "", _
                     Optional ByVal RecordID As String = "", Optional ByVal Details As String = "")
    On Error Resume Next   ' logging must never stop the user's work
    CurrentDb.Execute "INSERT INTO [AuditLog] ([EmployeeID], [ActionType], [ObjectName], " & _
        "[RecordID], [Details], [ComputerName]) VALUES (" & CurrentUserID() & ", " & _
        SqlText(Left$(ActionType, 30)) & ", " & SqlText(Left$(ObjectName, 50)) & ", " & _
        SqlText(Left$(RecordID, 30)) & ", " & SqlText(Details) & ", " & _
        SqlText(Left$(Environ$("COMPUTERNAME"), 50)) & ")", dbFailOnError
End Sub

'------------------------------------------------------------------------------
' Object checks
'------------------------------------------------------------------------------
Public Function FormExists(ByVal FormName As String) As Boolean
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If StrComp(obj.Name, FormName, vbTextCompare) = 0 Then
            FormExists = True
            Exit Function
        End If
    Next
End Function

Public Function ReportExists(ByVal ReportName As String) As Boolean
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllReports
        If StrComp(obj.Name, ReportName, vbTextCompare) = 0 Then
            ReportExists = True
            Exit Function
        End If
    Next
End Function

Public Function IsFormOpen(ByVal FormName As String) As Boolean
    If FormExists(FormName) Then IsFormOpen = CurrentProject.AllForms(FormName).IsLoaded
End Function
