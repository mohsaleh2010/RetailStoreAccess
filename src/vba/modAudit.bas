Attribute VB_Name = "modAudit"
'==============================================================================
' modAudit  -  Retail Store Management System
'
' Audit trail: who added, changed or deleted a record, when, on which computer,
' and the value of every field before and after.
'   AuditLog       one row per operation (also the logins and the document operations of LogAction)
'   AuditChanges   the fields of an operation: caption, old value, new value
' Where it is recorded:
'   data screens (modForms)     ADD / EDIT with the changed fields, DELETE with the whole record
'   bound grids                 budget lines, payroll lines (AuditFormBefore / AuditFormAfter)
'   documents deleted in code   assets, depreciation runs, bank operations, reconciliations,
'                               cheques, manual entries, payroll runs, budgets, VAT returns
'                               (AuditSnapshot before the delete, AuditDeleted after it)
'   documents changed in code   assets, manual entries, VAT drafts (AuditSnapshot + AuditEdited)
' Screen: frmAuditLog (permission AUDIT_LOG); report AUDIT_TRAIL.
' Password fields are never written, only that they changed.
'==============================================================================
Option Compare Database
Option Explicit

Private Const MAX_TEXT As Long = 255

Private m_pending As Object        ' form name -> Array(action, table, key field, changes)
Private m_captions As Object       ' "Table.Field" -> caption
Private m_be As DAO.Database       ' the back-end, read-only, for the field captions

'------------------------------------------------------------------------------
' Bound screens: collect the changes before the save, write them after it
'------------------------------------------------------------------------------
Public Sub AuditFormBefore(ByVal frm As Access.Form, Optional ByVal TableName As String = "", _
                           Optional ByVal KeyField As String = "")
    ' Call when the record is about to be saved (all checks passed).
    Dim ctl As Access.Control, changes As Collection, src As String, oldV As Variant, newV As Variant
    On Error GoTo Done                                ' the audit must never stop a save
    If Pending().Exists(frm.Name) Then Pending().Remove frm.Name     ' left by a save Access refused
    If Len(TableName) = 0 Then TableName = TagValue(frm, "TABLE")
    If Len(KeyField) = 0 Then KeyField = TagValue(frm, "PK")
    If Len(TableName) = 0 Then Exit Sub
    Set changes = New Collection
    For Each ctl In frm.Controls
        src = BoundField(ctl)
        If Len(src) > 0 And Not SkippedField(src) Then
            newV = ctl.Value
            If frm.NewRecord Then
                If Not IsEmptyValue(newV) Then AddChange changes, TableName, src, frm, Null, newV, ctl
            Else
                oldV = ctl.OldValue
                If Not SameValue(oldV, newV) Then AddChange changes, TableName, src, frm, oldV, newV, ctl
            End If
        End If
    Next
    If changes.Count = 0 And Not frm.NewRecord Then Exit Sub
    Pending().Item(frm.Name) = Array(IIf(frm.NewRecord, "ADD", "EDIT"), TableName, KeyField, changes)
Done:
End Sub

Public Sub AuditFormAfter(ByVal frm As Access.Form)
    ' Call after the record was saved (Form_AfterUpdate).
    Dim p As Variant, id As String
    On Error GoTo Done
    If Not Pending().Exists(frm.Name) Then Exit Sub
    p = Pending().Item(frm.Name)
    Pending().Remove frm.Name
    id = CStr(Nz(frm(CStr(p(2))).Value, ""))
    WriteAudit CStr(p(0)), CStr(p(1)), id, RecordLabel(frm), p(3)
Done:
End Sub

Public Sub AuditFormCancel(ByVal frm As Access.Form)
    ' The save was refused after AuditFormBefore: forget the collected changes.
    On Error Resume Next
    If Pending().Exists(frm.Name) Then Pending().Remove frm.Name
End Sub

'------------------------------------------------------------------------------
' Changes made in code: a snapshot of the record before, compared after
'------------------------------------------------------------------------------
Public Function AuditSnapshot(ByVal TableName As String, ByVal KeyField As String, ByVal id As Variant) As Collection
    ' Every field of one record: Array(field, caption, value). Empty collection when not found.
    Dim rs As DAO.Recordset, f As DAO.Field, snap As Collection
    Set snap = New Collection
    Set AuditSnapshot = snap
    On Error GoTo Done
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [" & TableName & "] WHERE [" & KeyField & "] = " & _
                                     KeyLiteral(id), dbOpenSnapshot)
    If Not rs.EOF Then
        For Each f In rs.Fields
            If Not SkippedField(f.Name) And Not IsSecretField(f.Name) Then
                snap.Add Array(f.Name, FieldCaption(TableName, f.Name), f.Value)
            End If
        Next
    End If
    rs.Close
Done:
End Function

Public Sub AuditDeleted(ByVal ActionType As String, ByVal TableName As String, ByVal id As Variant, _
                       ByVal Before As Collection, Optional ByVal Label As String = "")
    ' After a delete: the whole record as it was.
    Dim changes As Collection, item As Variant
    On Error GoTo Done
    Set changes = New Collection
    For Each item In Before
        If Not IsEmptyValue(item(2)) Then changes.Add Array(item(0), item(1), AuditValueText(item(2)), "")
    Next
    If Len(Label) = 0 Then Label = SnapshotLabel(Before)
    WriteAudit ActionType, TableName, CStr(id), Label, changes
Done:
End Sub

Public Sub AuditEdited(ByVal ActionType As String, ByVal TableName As String, ByVal KeyField As String, _
                       ByVal id As Variant, ByVal Before As Collection)
    ' After a change made in code: the fields that differ from the snapshot.
    Dim after As Collection, changes As Collection, i As Long
    On Error GoTo Done
    Set after = AuditSnapshot(TableName, KeyField, id)
    Set changes = New Collection
    For i = 1 To after.Count
        If i <= Before.Count Then
            If Not SameValue(Before(i)(2), after(i)(2)) Then
                changes.Add Array(after(i)(0), after(i)(1), AuditValueText(Before(i)(2)), AuditValueText(after(i)(2)))
            End If
        End If
    Next
    If changes.Count > 0 Then WriteAudit ActionType, TableName, CStr(id), SnapshotLabel(after), changes
Done:
End Sub

'------------------------------------------------------------------------------
' Writing
'------------------------------------------------------------------------------
Private Sub WriteAudit(ByVal ActionType As String, ByVal TableName As String, ByVal RecordID As String, _
                       ByVal Label As String, ByVal changes As Collection)
    Dim db As DAO.Database, rs As DAO.Recordset, logID As Long, item As Variant, n As Long
    Set db = CurrentDb
    Set rs = db.OpenRecordset("AuditLog", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!LogDate = Now
    rs!EmployeeID = CurrentUserID()
    rs!ActionType = Left$(ActionType, 30)
    rs!ObjectName = Left$(TableName, 50)
    If Len(RecordID) > 0 Then rs!RecordID = Left$(RecordID, 30)
    If Len(Label) > 0 Then rs!RecordLabel = Left$(Label, 100)
    rs!ComputerName = Left$(Environ$("COMPUTERNAME"), 50)
    logID = rs!LogID
    rs.Update
    rs.Close
    Set rs = db.OpenRecordset("AuditChanges", dbOpenDynaset, dbAppendOnly)
    For Each item In changes
        n = n + 1
        rs.AddNew
        rs!LogID = logID
        rs!LineNo = n
        rs!FieldName = Left$(item(0), 64)
        rs!FieldCaption = Left$(item(1), 100)
        If Len(item(2)) > 0 Then rs!OldValue = Left$(item(2), MAX_TEXT)
        If Len(item(3)) > 0 Then rs!NewValue = Left$(item(3), MAX_TEXT)
        rs.Update
    Next
    rs.Close
End Sub

Private Sub AddChange(ByVal changes As Collection, ByVal TableName As String, ByVal FieldName As String, _
                      ByVal frm As Access.Form, ByVal OldV As Variant, ByVal NewV As Variant, _
                      ByVal ctl As Access.Control)
    Dim caption As String
    caption = LabelCaption(frm, ctl.Name)
    If Len(caption) = 0 Then caption = FieldCaption(TableName, FieldName)
    If IsSecretField(FieldName) Then
        changes.Add Array(FieldName, caption, "", "(تغيّرت)")
    Else
        changes.Add Array(FieldName, caption, ShownText(ctl, OldV), ShownText(ctl, NewV))
    End If
End Sub

'------------------------------------------------------------------------------
' Values as text
'------------------------------------------------------------------------------
Public Function AuditValueText(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbBoolean
            If v Then
                AuditValueText = "نعم"
            Else
                AuditValueText = "لا"
            End If
        Case vbDate
            AuditValueText = GDate(v, CDbl(v) <> Int(CDbl(v)))
        Case vbString
            AuditValueText = v
        Case Else
            AuditValueText = CStr(v)
    End Select
    If Len(AuditValueText) > MAX_TEXT Then AuditValueText = Left$(AuditValueText, MAX_TEXT - 3) & "..."
End Function

Private Function ShownText(ByVal ctl As Access.Control, ByVal v As Variant) As String
    ' A combo box shows the text of the chosen row, not its hidden number.
    Dim i As Long
    ShownText = AuditValueText(v)
    If IsNull(v) Then Exit Function
    If ctl.ControlType <> acComboBox Then Exit Function
    On Error GoTo Done
    If ctl.ColumnCount < 2 Then Exit Function
    For i = 0 To ctl.ListCount - 1
        If CStr(ctl.ItemData(i)) = CStr(v) Then
            ShownText = Left$(Nz(ctl.Column(1, i), "") & " (" & CStr(v) & ")", MAX_TEXT)
            Exit Function
        End If
    Next
Done:
End Function

Public Function SameValue(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsEmptyValue(a) Or IsEmptyValue(b) Then
        SameValue = (IsEmptyValue(a) And IsEmptyValue(b))
    ElseIf VarType(a) = vbString Or VarType(b) = vbString Then
        SameValue = (CStr(a) = CStr(b))
    Else
        SameValue = (a = b)
    End If
End Function

Private Function IsEmptyValue(ByVal v As Variant) As Boolean
    If IsNull(v) Or IsEmpty(v) Then
        IsEmptyValue = True
    ElseIf VarType(v) = vbString Then
        IsEmptyValue = (Len(v) = 0)
    End If
End Function

'------------------------------------------------------------------------------
' Fields and captions
'------------------------------------------------------------------------------
Private Function BoundField(ByVal ctl As Access.Control) As String
    ' The field a text box, combo box or check box is bound to ("" for anything else).
    Dim src As String
    On Error GoTo Done
    Select Case ctl.ControlType
        Case acTextBox, acComboBox, acCheckBox
            src = Nz(ctl.ControlSource, "")
            If Len(src) > 0 And Left$(src, 1) <> "=" Then BoundField = src
    End Select
Done:
End Function

Private Function SkippedField(ByVal FieldName As String) As Boolean
    Select Case FieldName
        Case "CreatedAt", "UpdatedAt", "PasswordSalt"
            SkippedField = True
    End Select
End Function

Private Function IsSecretField(ByVal FieldName As String) As Boolean
    IsSecretField = (FieldName = "PasswordHash" Or FieldName = "PasswordSalt")
End Function

Private Function LabelCaption(ByVal frm As Access.Form, ByVal CtlName As String) As String
    ' The caption of the label next to a field on the data screens ("lbl" & field), without " *".
    Dim c As String
    On Error GoTo Done
    c = Trim$(frm.Controls("lbl" & CtlName).Caption)
    If Right$(c, 1) = "*" Then c = Trim$(Left$(c, Len(c) - 1))
    LabelCaption = c
Done:
End Function

Public Function FieldCaption(ByVal TableName As String, ByVal FieldName As String) As String
    ' The Arabic caption BuildSchema gave the field in the back-end (cached), else the field name.
    Dim key As String, c As String, tdf As DAO.TableDef
    key = TableName & "." & FieldName
    If m_captions Is Nothing Then Set m_captions = CreateObject("Scripting.Dictionary")
    If m_captions.Exists(key) Then
        FieldCaption = m_captions(key)
        Exit Function
    End If
    c = FieldName
    On Error Resume Next
    If m_be Is Nothing Then Set m_be = BackEndDatabase()
    If Not m_be Is Nothing Then
        Set tdf = m_be.TableDefs(TableName)
        c = tdf.Fields(FieldName).Properties("Caption").Value
        If Len(c) = 0 Then c = FieldName
    End If
    On Error GoTo 0
    m_captions(key) = c
    FieldCaption = c
End Function

Private Function BackEndDatabase() As DAO.Database
    ' The data file of the linked tables, opened read-only.
    Dim db As DAO.Database, tdf As DAO.TableDef, cnn As String
    On Error GoTo Done
    Set db = CurrentDb
    Set tdf = db.TableDefs("AuditLog")
    cnn = tdf.Connect
    If InStr(cnn, "DATABASE=") > 0 Then
        Set BackEndDatabase = DBEngine.OpenDatabase(Mid$(cnn, InStr(cnn, "DATABASE=") + 9), False, True)
    Else
        Set BackEndDatabase = db                 ' tables in this file (development)
    End If
Done:
End Function

Private Function RecordLabel(ByVal frm As Access.Form) As String
    ' A readable name of the record: its first name / number / code field.
    Dim ctl As Access.Control, src As String, pass As Long, ends As Variant
    On Error GoTo Done
    ends = Array("Name", "Number", "Code")
    For pass = 0 To 2
        For Each ctl In frm.Controls
            src = BoundField(ctl)
            If Len(src) > Len(ends(pass)) Then
                If Right$(src, Len(ends(pass))) = ends(pass) And Not IsNull(ctl.Value) Then
                    RecordLabel = Left$(CStr(ctl.Value), 100)
                    Exit Function
                End If
            End If
        Next
    Next
Done:
End Function

Private Function SnapshotLabel(ByVal snap As Collection) As String
    Dim item As Variant, pass As Long, ends As Variant, f As String
    ends = Array("Name", "Number", "Code")
    For pass = 0 To 2
        For Each item In snap
            f = item(0)
            If Len(f) > Len(ends(pass)) Then
                If Right$(f, Len(ends(pass))) = ends(pass) And Not IsNull(item(2)) Then
                    SnapshotLabel = Left$(AuditValueText(item(2)), 100)
                    Exit Function
                End If
            End If
        Next
    Next
End Function

Private Function KeyLiteral(ByVal id As Variant) As String
    If IsNumeric(id) Then
        KeyLiteral = CStr(id)
    Else
        KeyLiteral = SqlText(id)
    End If
End Function

Private Function Pending() As Object
    If m_pending Is Nothing Then Set m_pending = CreateObject("Scripting.Dictionary")
    Set Pending = m_pending
End Function

'------------------------------------------------------------------------------
' frmAuditLog
'------------------------------------------------------------------------------
Public Sub AuditScreenLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    frm!txtFrom.Value = Date - 7
    frm!txtTo.Value = Date
    frm!cboTable.RowSource = "SELECT DISTINCT ObjectName FROM AuditLog WHERE ObjectName Is Not Null ORDER BY ObjectName"
    AuditScreenShow frm
End Sub

Public Sub AuditScreenShow(ByVal frm As Access.Form)
    Dim where As String, txt As String, n As Long
    If Not IsDate(frm!txtFrom.Value) Or Not IsDate(frm!txtTo.Value) Then
        ShowWarning "أدخل تاريخ البداية وتاريخ النهاية."
        Exit Sub
    End If
    where = "a.LogDate >= " & SqlDate(DateValue(frm!txtFrom.Value)) & " AND a.LogDate < " & _
            SqlDate(DateValue(frm!txtTo.Value) + 1)
    If Not IsNull(frm!cboUser.Value) Then where = where & " AND a.EmployeeID = " & CLng(frm!cboUser.Value)
    If Not IsNull(frm!cboTable.Value) Then where = where & " AND a.ObjectName = " & SqlText(frm!cboTable.Value)
    Select Case Nz(frm!cboAction.Value, "")
        Case "ADD", "EDIT", "DELETE"
            where = where & " AND a.ActionType = " & SqlText(frm!cboAction.Value)
        Case "LOGIN"
            where = where & " AND a.ActionType IN ('LOGIN', 'LOGOUT', 'LOGIN_FAILED')"
        Case "DOCS"
            where = where & " AND a.ActionType NOT IN ('ADD', 'EDIT', 'DELETE', 'LOGIN', 'LOGOUT', 'LOGIN_FAILED', 'REPORT')"
    End Select
    txt = Trim$(Nz(frm!txtSearch.Value, ""))
    If Len(txt) > 0 Then
        where = where & " AND (a.RecordID = " & SqlText(txt) & " OR a.RecordLabel LIKE " & SqlText("*" & txt & "*") & _
                " OR a.Details LIKE " & SqlText("*" & txt & "*") & ")"
    End If
    frm!lstLog.RowSource = "SELECT a.LogID, GDate(a.LogDate, True) AS [الوقت], Nz(e.EmployeeName, '-') AS [المستخدم], " & _
        "ActionName(a.ActionType) AS [العملية], a.ObjectName AS [الجدول], a.RecordID AS [الرقم], " & _
        "a.RecordLabel AS [السجل], a.ComputerName AS [الجهاز] FROM AuditLog AS a LEFT JOIN Employees AS e ON " & _
        "a.EmployeeID = e.EmployeeID WHERE " & where & " ORDER BY a.LogID DESC"
    frm!lstChanges.RowSource = ""
    n = frm!lstLog.ListCount - 1
    If n < 0 Then n = 0
    frm!lblCount.Caption = n & " عملية"
End Sub

Public Sub AuditScreenPick(ByVal frm As Access.Form)
    Dim id As Variant, details As String
    id = frm!lstLog.Value
    If IsNull(id) Then Exit Sub
    frm!lstChanges.RowSource = "SELECT ChangeID, FieldCaption AS [الحقل], Nz(OldValue, '-') AS [قبل], " & _
        "Nz(NewValue, '-') AS [بعد] FROM AuditChanges WHERE LogID = " & CLng(id) & " ORDER BY LineNo"
    details = Nz(DbValue("SELECT Details FROM AuditLog WHERE LogID = " & CLng(id)), "")
    frm!lblDetails.Caption = IIf(Len(details) = 0, " ", Left$(details, 250))
End Sub

Public Function ActionName(ByVal ActionType As Variant) As String
    ' The Arabic name of an AuditLog action (used by the screen and the report).
    Select Case Nz(ActionType, "")
        Case "ADD": ActionName = "إضافة"
        Case "EDIT": ActionName = "تعديل"
        Case "DELETE": ActionName = "حذف"
        Case "LOGIN": ActionName = "دخول"
        Case "LOGOUT": ActionName = "خروج"
        Case "LOGIN_FAILED": ActionName = "دخول خاطئ"
        Case "REPORT": ActionName = "تقرير"
        Case Else
            If Right$(Nz(ActionType, ""), 7) = "_DELETE" Then
                ActionName = "حذف (" & ActionType & ")"
            ElseIf Right$(Nz(ActionType, ""), 5) = "_EDIT" Then
                ActionName = "تعديل (" & ActionType & ")"
            Else
                ActionName = Nz(ActionType, "")
            End If
    End Select
End Function

Public Sub PrintAuditLog(ByVal frm As Access.Form)
    If Not IsDate(frm!txtFrom.Value) Or Not IsDate(frm!txtTo.Value) Then
        ShowWarning "أدخل تاريخ البداية وتاريخ النهاية."
        Exit Sub
    End If
    OpenReportOrQuery "rptAuditTrail", "AuditTrailQuery", "[LogDate] >= " & SqlDate(DateValue(frm!txtFrom.Value)) & _
                      " AND [LogDate] < " & SqlDate(DateValue(frm!txtTo.Value) + 1), _
                      PeriodText(DateValue(frm!txtFrom.Value), DateValue(frm!txtTo.Value))
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Public Function TestAudit() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, passed As Long, failed As Long, report As String
    Dim id As Long, before As Collection, logID As Variant, lastLog As Long
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestAudit  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Check AuditValueText(True) = "نعم" And AuditValueText(Null) = "" And AuditValueText(12.5) = CStr(12.5), "القيم كنص", _
          passed, failed, report
    Check SameValue(Null, "") And Not SameValue(1, 2) And SameValue("5", 5), "مقارنة القيم", passed, failed, report
    Check ActionName("EDIT") = "تعديل" And ActionName("CHEQUE_DELETE") = "حذف (CHEQUE_DELETE)", "أسماء العمليات", _
          passed, failed, report
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    lastLog = Nz(DMax("LogID", "AuditLog"), 0)
    CurrentDb.Execute "INSERT INTO CostCenters (CenterCode, CenterName, IsDefault, IsActive) VALUES ('TAUD', " & _
                      "'TEST مركز التدقيق', False, True)", dbFailOnError
    id = DbValue("SELECT CostCenterID FROM CostCenters WHERE CenterCode = 'TAUD'")
    Set before = AuditSnapshot("CostCenters", "CostCenterID", id)
    Check before.Count >= 4, "لقطة السجل قبل التعديل (" & before.Count & " حقول)", passed, failed, report
    CurrentDb.Execute "UPDATE CostCenters SET CenterName = 'TEST مركز معدَّل' WHERE CostCenterID = " & id, dbFailOnError
    AuditEdited "EDIT", "CostCenters", "CostCenterID", id, before
    logID = DbValue("SELECT Max(LogID) FROM AuditLog WHERE ActionType = 'EDIT' AND ObjectName = 'CostCenters' " & _
                    "AND RecordID = '" & id & "' AND LogID > " & lastLog)
    Check Not IsNull(logID) And Nz(DbValue("SELECT COUNT(*) FROM AuditChanges WHERE LogID = " & Nz(logID, 0)), 0) = 1 And _
          Nz(DbValue("SELECT OldValue FROM AuditChanges WHERE FieldName = 'CenterName' AND LogID = " & Nz(logID, 0)), "") = _
          "TEST مركز التدقيق" And _
          Nz(DbValue("SELECT EmployeeID FROM AuditLog WHERE LogID = " & Nz(logID, 0)), 0) = CurrentUserID(), _
          "التعديل: الحقل المتغير فقط، بقيمته قبل وبعد، ومن عدّله", passed, failed, report
    Set before = AuditSnapshot("CostCenters", "CostCenterID", id)
    CurrentDb.Execute "DELETE FROM CostCenters WHERE CostCenterID = " & id, dbFailOnError
    AuditDeleted "DELETE", "CostCenters", id, before
    logID = DbValue("SELECT Max(LogID) FROM AuditLog WHERE ActionType = 'DELETE' AND ObjectName = 'CostCenters' " & _
                    "AND LogID > " & lastLog)
    Check Nz(DbValue("SELECT RecordLabel FROM AuditLog WHERE LogID = " & Nz(logID, 0)), "") = "TEST مركز معدَّل" And _
          Nz(DbValue("SELECT COUNT(*) FROM AuditChanges WHERE NewValue Is Null AND LogID = " & Nz(logID, 0)), 0) >= 3, _
          "الحذف: السجل كله كما كان قبل الحذف", passed, failed, report
    Check Nz(DbValue("SELECT COUNT(*) FROM AuditTrailQuery WHERE LogID = " & Nz(logID, 0)), 0) >= 3, _
          "تقرير سجل التدقيق يعرض الحقول", passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    Check False, "خطأ: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات سجل التدقيق ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestAudit"
        TestAudit = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestAudit"
    End If
End Function

Private Sub Check(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
                  ByRef report As String)
    If ok Then
        passed = passed + 1
        Debug.Print "[OK] " & Title
    Else
        failed = failed + 1
        report = report & "- " & Title & vbCrLf
        Debug.Print "[X]  " & Title
    End If
End Sub
