Attribute VB_Name = "modEInvoice"
'==============================================================================
' modEInvoice  -  Retail Store Management System (e-invoicing, shared by both countries)
'
' docs/45-EInvoice-Foundation.md. Every sales invoice and return has an e-invoicing status
' (SalesInvoices / SalesReturns.ZatcaStatus):
'   NOT_SENT   e-invoicing was not enabled when the document was saved
'   PENDING    to send
'   REPORTED / CLEARED (Saudi Arabia), SUBMITTED -> VALID / INVALID (Egypt), WARNING, REJECTED
' SendEInvoice hands one document to the platform of the operating country (modCountry) with
' Application.Run, so each platform is a module of its own:
'   Saudi Arabia  ZatcaReadyProblem() / ZatcaSendDocument(DocKind, DocID)   (phase C)
'   Egypt         EtaReadyProblem()   / EtaSendDocument(DocKind, DocID)     (phase D)
' ...ReadyProblem returns "" when the platform can send (certificate, credentials...); ...SendDocument
' returns "RESULT|message" (OK, WARNING, REJECTED, ERROR, NETWORK), sets the status with
' SetEInvoiceStatus and writes every request with LogEInvoice. frmEInvoices follows and resends.
'==============================================================================
Option Compare Database
Option Explicit

Private Const ERR_NO_PROCEDURE As Long = 2517     ' Application.Run: the procedure does not exist

'==============================================================================
' Settings and platform
'==============================================================================
Public Function EInvoiceEnabled() As Boolean
    EInvoiceEnabled = Nz(SettingValue("EInvoiceEnabled"), False)
End Function

Public Function EInvoicePlatform() As String
    ' ZATCA (Saudi Arabia) or ETA (Egypt)
    If AppCountry() = "EG" Then
        EInvoicePlatform = "ETA"
    Else
        EInvoicePlatform = "ZATCA"
    End If
End Function

Public Function EInvoicePlatformName() As String
    If AppCountry() = "EG" Then
        EInvoicePlatformName = "„‰ŸÊ„… „’·Õ… «·÷—«∆» «·„’—Ì…"
    Else
        EInvoicePlatformName = "„‰’… ›« Ê—… (ÂÌ∆… «·“ﬂ«… Ê«·÷—Ì»… Ê«·Ã„«—ﬂ)"
    End If
End Function

Private Function AdapterProc(ByVal Name As String) As String
    ' ZatcaSendDocument / EtaSendDocument ...
    If AppCountry() = "EG" Then
        AdapterProc = "Eta" & Name
    Else
        AdapterProc = "Zatca" & Name
    End If
End Function

Public Function EInvoiceReadyProblem() As String
    ' "" when the platform of the country is installed and set up (certificate, credentials...).
    Dim reply As Variant
    On Error Resume Next
    reply = Application.Run(AdapterProc("ReadyProblem"))
    If Err.Number = ERR_NO_PROCEDURE Or Err.Number = 2465 Or Err.Number = 438 Then
        EInvoiceReadyProblem = "—»ÿ " & EInvoicePlatformName() & " ·„ ÌıÀ»Û¯  »⁄œ ›Ì «·»—‰«„Ã."
    ElseIf Err.Number <> 0 Then
        EInvoiceReadyProblem = " ⁄–¯— ›Õ’ ≈⁄œ«œ «·›« Ê—… «·≈·ﬂ —Ê‰Ì…: " & Err.Description
    Else
        EInvoiceReadyProblem = Nz(reply, "")
    End If
End Function

'==============================================================================
' Sending
'==============================================================================
Public Function DocTableOf(ByVal DocKind As String) As String
    If DocKind = "RETURN" Then
        DocTableOf = "SalesReturns"
    Else
        DocTableOf = "SalesInvoices"
    End If
End Function

Public Function DocKeyOf(ByVal DocKind As String) As String
    If DocKind = "RETURN" Then
        DocKeyOf = "SalesReturnID"
    Else
        DocKeyOf = "SalesInvoiceID"
    End If
End Function

Public Function SplitResult(ByVal Reply As String, ByRef Message As String) As String
    ' "REJECTED|the reason" -> "REJECTED" and the reason. An unknown reply is an ERROR.
    Dim bar As Long, code As String
    bar = InStr(Reply, "|")
    If bar > 0 Then
        code = UCase$(Left$(Reply, bar - 1))
        Message = Mid$(Reply, bar + 1)
    Else
        code = UCase$(Reply)
        Message = ""
    End If
    Select Case code
        Case "OK", "WARNING", "REJECTED", "ERROR", "NETWORK"
            SplitResult = code
        Case Else
            SplitResult = "ERROR"
            Message = Reply
    End Select
End Function

Public Function SendEInvoice(ByVal DocKind As String, ByVal DocID As Long, ByRef Message As String) As String
    ' Sends one document to the platform of the country: OK, WARNING, REJECTED, ERROR or NETWORK.
    Dim reply As Variant, result As String
    Message = ""
    If Not EInvoiceEnabled() Then
        Message = "«·›« Ê—… «·≈·ﬂ —Ê‰Ì… €Ì— „›⁄¯·… ›Ì «·≈⁄œ«œ« ."
        SendEInvoice = "ERROR"
        Exit Function
    End If
    Message = EInvoiceReadyProblem()
    If Len(Message) > 0 Then
        SendEInvoice = "ERROR"
        Exit Function
    End If
    On Error Resume Next
    reply = Application.Run(AdapterProc("SendDocument"), DocKind, DocID)
    If Err.Number <> 0 Then
        result = "ERROR"
        Message = " ⁄–¯— «·≈—”«·: " & Err.Description
    Else
        On Error GoTo 0
        result = SplitResult(Nz(reply, ""), Message)
    End If
    On Error GoTo 0
    CurrentDb.Execute "UPDATE " & DocTableOf(DocKind) & " SET EInvoiceAttempts = Nz(EInvoiceAttempts, 0) + 1, " & _
                      "EInvoiceError = " & SqlText(Left$(IIf(result = "OK", "", Message), 255)) & _
                      " WHERE " & DocKeyOf(DocKind) & " = " & DocID, dbFailOnError
    SendEInvoice = result
End Function

Public Function SendPendingEInvoices(ByRef Sent As Long, ByRef Failed As Long) As String
    ' Every PENDING document, oldest first. Stops at the first network problem (the next would fail too).
    ' Returns "" or the last problem.
    Dim rs As DAO.Recordset, result As String, msg As String, lastProblem As String
    Sent = 0
    Failed = 0
    Set rs = CurrentDb.OpenRecordset("SELECT DocKind, DocID FROM qryEInvoiceDocs WHERE EStatus = 'PENDING' " & _
                                     "ORDER BY DocDate, DocKind DESC, DocID", dbOpenSnapshot)
    Do Until rs.EOF
        result = SendEInvoice(rs!DocKind, rs!DocID, msg)
        If result = "OK" Or result = "WARNING" Then
            Sent = Sent + 1
        Else
            Failed = Failed + 1
            lastProblem = msg
            If result = "NETWORK" Or result = "ERROR" Then Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
    SendPendingEInvoices = lastProblem
End Function

Public Sub EInvoiceAfterSave(ByVal DocKind As String, ByVal DocID As Long)
    ' From the sales posting (modSales), after the commit: send at once when the platform is ready.
    ' A failure leaves the document PENDING for frmEInvoices; it never fails the sale.
    Dim msg As String
    On Error GoTo Done
    If Not EInvoiceEnabled() Then Exit Sub
    If Nz(DbValue("SELECT ZatcaStatus FROM " & DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & DocID), _
          "") <> "PENDING" Then Exit Sub
    If Len(EInvoiceReadyProblem()) > 0 Then Exit Sub
    SendEInvoice DocKind, DocID, msg
Done:
End Sub

'==============================================================================
' For the platform modules: status and log
'==============================================================================
Public Sub SetEInvoiceStatus(ByVal DocKind As String, ByVal DocID As Long, ByVal Status As String, _
                             Optional ByVal Response As String = "")
    CurrentDb.Execute "UPDATE " & DocTableOf(DocKind) & " SET ZatcaStatus = '" & Status & "', ZatcaSubmittedAt = Now(), " & _
                      "ZatcaResponse = " & SqlText(Left$(Response, 60000)) & " WHERE " & DocKeyOf(DocKind) & " = " & _
                      DocID, dbFailOnError
End Sub

Public Sub LogEInvoice(ByVal DocKind As String, ByVal DocID As Long, ByVal Action As String, ByVal Endpoint As String, _
                       ByVal HttpStatus As Long, ByVal Result As String, ByVal Message As String, _
                       ByVal RequestBody As String, ByVal ResponseBody As String, ByVal DurationMs As Long)
    ' One request to the platform and its reply (EInvoiceLog). DocKind "" = a request of the device.
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("EInvoiceLog", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!LoggedAt = Now
    rs!Country = AppCountry()
    If Len(DocKind) > 0 Then
        rs!DocKind = DocKind
        rs!DocID = DocID
        rs!DocNumber = DbValue("SELECT DocNumber FROM qryEInvoiceDocs WHERE DocKind = '" & DocKind & _
                               "' AND DocID = " & DocID)
    End If
    rs!Action = Left$(Action, 20)
    If Len(Endpoint) > 0 Then rs!Endpoint = Left$(Endpoint, 255)
    rs!HttpStatus = HttpStatus
    rs!Result = Result
    If Len(Message) > 0 Then rs!Message = Left$(Message, 255)
    If Len(RequestBody) > 0 Then rs!RequestBody = RequestBody
    If Len(ResponseBody) > 0 Then rs!ResponseBody = ResponseBody
    rs!DurationMs = DurationMs
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Close
End Sub

Public Function EInvoiceOverdueCount() As Long
    ' Documents still PENDING 20 hours after their date (Saudi Arabia: a simplified invoice is reported
    ' within 24 hours).
    EInvoiceOverdueCount = Nz(DbValue("SELECT COUNT(*) FROM qryEInvoiceDocs WHERE EStatus = 'PENDING' AND " & _
                                      "DocDate < DateAdd('h', -20, Now())"), 0)
End Function

'==============================================================================
' frmEInvoices
'==============================================================================
Public Sub EInvoicesLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    frm!txtFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    frm!txtTo.Value = Date
    frm!cboStatus.Value = "ATTENTION"
    frm!cboEnvironment.Value = Nz(SettingValue("EInvoiceEnvironment"), "TEST")
    frm!chkEnabled.Value = EInvoiceEnabled()
    EInvoicesShow frm
End Sub

Public Sub EInvoiceSettingChanged(ByVal frm As Access.Form)
    ' The environment and the switch of frmEInvoices, saved at once. Store settings permission; sending is
    ' enabled only when the platform of the country is installed and set up (EInvoiceReadyProblem).
    Dim problem As String, before As Collection
    If Not HasPermission("SETTINGS") Then
        ShowWarning " €ÌÌ— ≈⁄œ«œ«  «·›« Ê—… «·≈·ﬂ —Ê‰Ì… ÌÕ «Ã ’·«ÕÌ… ≈⁄œ«œ«  «·„Õ·."
        problem = "x"
    ElseIf Nz(frm!chkEnabled.Value, False) And Not EInvoiceEnabled() Then
        problem = EInvoiceReadyProblem()
        If Len(problem) > 0 Then ShowWarning "·« Ì„ﬂ‰  ›⁄Ì· «·›« Ê—… «·≈·ﬂ —Ê‰Ì… «·¬‰:" & vbCrLf & problem
    End If
    If Len(problem) > 0 Then
        frm!cboEnvironment.Value = Nz(SettingValue("EInvoiceEnvironment"), "TEST")
        frm!chkEnabled.Value = EInvoiceEnabled()
        Exit Sub
    End If
    Set before = AuditSnapshot("Settings", "SettingID", 1)
    CurrentDb.Execute "UPDATE Settings SET EInvoiceEnvironment = '" & Nz(frm!cboEnvironment.Value, "TEST") & _
                      "', EInvoiceEnabled = " & IIf(Nz(frm!chkEnabled.Value, False), "True", "False") & _
                      " WHERE SettingID = 1", dbFailOnError
    AuditEdited "EDIT", "Settings", "SettingID", 1, before
    EInvoicesShow frm
End Sub

Public Sub EInvoicesShow(ByVal frm As Access.Form)
    Dim sql As String, where As String, n As Long, overdue As Long, info As String
    Calendar = vbCalGreg
    Select Case Nz(frm!cboStatus.Value, "ALL")
        Case "ATTENTION": where = "d.EStatus IN ('PENDING','REJECTED','INVALID','WARNING')"
        Case "SENT": where = "d.EStatus IN ('REPORTED','CLEARED','SUBMITTED','VALID')"
        Case "NOT_SENT": where = "d.EStatus = 'NOT_SENT'"
        Case Else: where = "True"
    End Select
    If IsDate(frm!txtFrom.Value) Then where = where & " AND d.DocDate >= " & SqlDate(DateValue(frm!txtFrom.Value))
    If IsDate(frm!txtTo.Value) Then
        where = where & " AND d.DocDate < " & SqlDate(DateAdd("d", 1, DateValue(frm!txtTo.Value)))
    End If
    sql = Tr("SELECT d.DocKind & ':' & d.DocID AS DocKey, d.DocNumber AS [—ﬁ„ «·„” ‰œ], Format(d.DocDate, 'yyyy/mm/dd hh:nn') AS [«· «—ÌŒ], " & _
             "d.DocKindName AS [‰Ê⁄ «·„” ‰œ], d.CustomerName AS [«·⁄„Ì·], Format(d.TotalAmount, '#,##0.00') AS [«·≈Ã„«·Ì], " & _
             "d.StatusName AS [Õ«·… «·≈—”«·], d.Attempts AS [⁄œœ «·„Õ«Ê·« ], d.LastError AS [¬Œ— Œÿ√] " & _
             "FROM qryEInvoiceDocs AS d WHERE {WHERE} ORDER BY d.DocDate DESC, d.DocID DESC")
    frm!lstDocs.RowSource = Replace(sql, "{WHERE}", where)
    frm!lstLog.RowSource = ""
    n = frm!lstDocs.ListCount - 1
    If n < 0 Then n = 0
    overdue = EInvoiceOverdueCount()
    If EInvoiceEnabled() Then
        info = EInvoicePlatformName() & " - " & EnvironmentName()
    Else
        info = "«·›« Ê—… «·≈·ﬂ —Ê‰Ì… €Ì— „›⁄¯·… («·≈⁄œ«œ« )"
    End If
    info = info & "    |    " & n & " „” ‰œ"
    If overdue > 0 Then info = info & "    |    „ √Œ—… √ﬂÀ— „‰ 20 ”«⁄…: " & overdue
    frm!lblSummary.Caption = Tr(info)
    frm!lblSummary.ForeColor = IIf(overdue > 0, CLR_DANGER, CLR_PRIMARY)
End Sub

Public Function EnvironmentName() As String
    Select Case Nz(SettingValue("EInvoiceEnvironment"), "TEST")
        Case "PRODUCTION": EnvironmentName = "«·»Ì∆… «·›⁄·Ì…"
        Case "SIMULATION": EnvironmentName = "»Ì∆… «·„Õ«ﬂ«…"
        Case Else: EnvironmentName = "«·»Ì∆… «· Ã—Ì»Ì…"
    End Select
End Function

Private Function PickedDoc(ByVal frm As Access.Form, ByRef DocKind As String, ByRef DocID As Long) As Boolean
    ' lstDocs key: "SALE:12" / "RETURN:3"
    Dim key As String
    key = Nz(frm!lstDocs.Value, "")
    If InStr(key, ":") = 0 Then Exit Function
    DocKind = Left$(key, InStr(key, ":") - 1)
    DocID = CLng(Mid$(key, InStr(key, ":") + 1))
    PickedDoc = True
End Function

Public Sub EInvoicesPick(ByVal frm As Access.Form)
    Dim kind As String, id As Long
    If Not PickedDoc(frm, kind, id) Then Exit Sub
    frm!lstLog.RowSource = Tr("SELECT l.LogID, Format(l.LoggedAt, 'yyyy/mm/dd hh:nn:ss') AS [«·Êﬁ ], l.Action AS [«·⁄„·Ì…], " & _
        "l.Result AS [‰ ÌÃ… «·ÿ·»], l.HttpStatus AS [—„“ «·—œ], l.Message AS [—”«·… «·„‰ŸÊ„…] FROM EInvoiceLog AS l " & _
        "WHERE l.DocKind = '" & kind & "' AND l.DocID = " & id & " ORDER BY l.LogID DESC")
End Sub

Public Sub EInvoicesSendPicked(ByVal frm As Access.Form)
    Dim kind As String, id As Long, result As String, msg As String
    If Not CanScreenAction(frm.Name, "EDIT") Then Exit Sub
    If Not PickedDoc(frm, kind, id) Then
        ShowWarning "«Œ — „” ‰œ« „‰ «·ﬁ«∆„…."
        Exit Sub
    End If
    DoCmd.Hourglass True
    result = SendEInvoice(kind, id, msg)
    DoCmd.Hourglass False
    If result = "OK" Then
        ShowInfo " „ «·≈—”«·."
    ElseIf result = "WARNING" Then
        ShowInfo " „ «·≈—”«· „⁄  Õ–Ì—: " & msg
    Else
        ShowWarning "·„ Ìıﬁ»· «·„” ‰œ: " & msg
    End If
    EInvoicesShow frm
End Sub

Public Sub EInvoicesSendAll(ByVal frm As Access.Form)
    Dim sent As Long, failed As Long, msg As String
    If Not CanScreenAction(frm.Name, "EDIT") Then Exit Sub
    If Not EInvoiceEnabled() Then
        ShowWarning "«·›« Ê—… «·≈·ﬂ —Ê‰Ì… €Ì— „›⁄¯·… ›Ì «·≈⁄œ«œ« ."
        Exit Sub
    End If
    DoCmd.Hourglass True
    msg = SendPendingEInvoices(sent, failed)
    DoCmd.Hourglass False
    If failed = 0 Then
        ShowInfo "√ı—”· " & sent & " „” ‰œ."
    Else
        ShowWarning "√ı—”· " & sent & " Ê·„ Ìı—”· " & failed & "." & vbCrLf & msg
    End If
    EInvoicesShow frm
End Sub

'==============================================================================
' In-Access test (changes rolled back)
'==============================================================================
Public Function TestEInvoice() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, passed As Long, failed As Long, report As String, msg As String
    Dim id As Long, logs As Long
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestEInvoice  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Record JsonGet("{""a"": {""b"": [1, {""c"": ""x\""y""}]}}", "a.b[1].c") = "x""y" And _
           IsNull(JsonGet("{""a"": null}", "a")) And JsonGet("{""n"": 12.50}", "n") = "12.50", "ﬁ—«¡… —œ JSON", _
           passed, failed, report
    Record JsonEscape("a""b\c" & vbLf) = "a\""b\\c\n", "‰’ œ«Œ· JSON", passed, failed, report
    Record SplitResult("REJECTED|”»»", msg) = "REJECTED" And msg = "”»»" And SplitResult("??", msg) = "ERROR", _
           "ﬁ—«¡… ‰ ÌÃ… «·„‰ŸÊ„…", passed, failed, report
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    id = Nz(DbValue("SELECT Max(SalesInvoiceID) FROM SalesInvoices"), 0)
    If id > 0 Then
        CurrentDb.Execute "UPDATE Settings SET EInvoiceEnabled = False WHERE SettingID = 1", dbFailOnError
        Record SendEInvoice("SALE", id, msg) = "ERROR" And Len(msg) > 0, "·« ≈—”«· ﬁ»· «· ›⁄Ì·", passed, failed, report
        SetEInvoiceStatus "SALE", id, "REJECTED", "{""test"": 1}"
        Record DbValue("SELECT ZatcaStatus FROM SalesInvoices WHERE SalesInvoiceID = " & id) = "REJECTED", _
               " €ÌÌ— Õ«·… «·„” ‰œ", passed, failed, report
        logs = Nz(DbValue("SELECT COUNT(*) FROM EInvoiceLog"), 0)
        LogEInvoice "SALE", id, "SEND", "https://example.test", 400, "REJECTED", "TEST", "{}", "{}", 5
        Record Nz(DbValue("SELECT COUNT(*) FROM EInvoiceLog"), 0) = logs + 1 And _
               Nz(DbValue("SELECT DocNumber FROM EInvoiceLog WHERE LogID = (SELECT Max(LogID) FROM EInvoiceLog)"), "") = _
               Nz(DbValue("SELECT InvoiceNumber FROM SalesInvoices WHERE SalesInvoiceID = " & id), "-"), _
               "”Ã· «·≈—”«· »—ﬁ„ «·„” ‰œ", passed, failed, report
        CurrentDb.Execute "UPDATE SalesInvoices SET ZatcaStatus = 'PENDING', InvoiceDate = DateAdd('h', -30, Now()) " & _
                          "WHERE SalesInvoiceID = " & id, dbFailOnError
        Record EInvoiceOverdueCount() >= 1, " ‰»ÌÂ «·„” ‰œ«  «·„ √Œ—…", passed, failed, report
    Else
        Debug.Print "[--] ·«  ÊÃœ ›Ê« Ì— »Ì⁄"
    End If
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    Record False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·›« Ê—… «·≈·ﬂ —Ê‰Ì… ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, _
                "TestEInvoice"
        TestEInvoice = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestEInvoice"
    End If
End Function

Private Sub Record(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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
