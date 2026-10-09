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
        EInvoicePlatformName = "منظومة مصلحة الضرائب المصرية"
    Else
        EInvoicePlatformName = "منصة فاتورة (هيئة الزكاة والضريبة والجمارك)"
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
        EInvoiceReadyProblem = "ربط " & EInvoicePlatformName() & " لم يُثبَّت بعد في البرنامج."
    ElseIf Err.Number <> 0 Then
        EInvoiceReadyProblem = "تعذّر فحص إعداد الفاتورة الإلكترونية: " & Err.Description
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
        Message = "الفاتورة الإلكترونية غير مفعّلة في الإعدادات."
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
        Message = "تعذّر الإرسال: " & Err.Description
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
        ShowWarning "تغيير إعدادات الفاتورة الإلكترونية يحتاج صلاحية إعدادات المحل."
        problem = "x"
    ElseIf Nz(frm!chkEnabled.Value, False) And Not EInvoiceEnabled() Then
        problem = EInvoiceReadyProblem()
        If Len(problem) > 0 Then ShowWarning "لا يمكن تفعيل الفاتورة الإلكترونية الآن:" & vbCrLf & problem
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
    sql = Tr("SELECT d.DocKind & ':' & d.DocID AS DocKey, d.DocNumber AS [رقم المستند], Format(d.DocDate, 'yyyy/mm/dd hh:nn') AS [التاريخ], " & _
             "d.DocKindName AS [نوع المستند], d.CustomerName AS [العميل], Format(d.TotalAmount, '#,##0.00') AS [الإجمالي], " & _
             "d.StatusName AS [حالة الإرسال], d.Attempts AS [عدد المحاولات], d.LastError AS [آخر خطأ] " & _
             "FROM qryEInvoiceDocs AS d WHERE {WHERE} ORDER BY d.DocDate DESC, d.DocID DESC")
    frm!lstDocs.RowSource = Replace(sql, "{WHERE}", where)
    frm!lstLog.RowSource = ""
    n = frm!lstDocs.ListCount - 1
    If n < 0 Then n = 0
    overdue = EInvoiceOverdueCount()
    If EInvoiceEnabled() Then
        info = EInvoicePlatformName() & " - " & EnvironmentName()
    Else
        info = "الفاتورة الإلكترونية غير مفعّلة (الإعدادات)"
    End If
    info = info & "    |    " & n & " مستند"
    If overdue > 0 Then info = info & "    |    متأخرة أكثر من 20 ساعة: " & overdue
    frm!lblSummary.Caption = Tr(info)
    frm!lblSummary.ForeColor = IIf(overdue > 0, CLR_DANGER, CLR_PRIMARY)
End Sub

Public Sub EInvoiceSetupOpen()
    ' The setup of the platform of the country: Saudi Arabia (modZatcaApi), Egypt (modEtaReceipt).
    If AppCountry() = "EG" Then
        OpenScreen "frmEtaSetup"
    Else
        OpenScreen "frmZatcaSetup"
    End If
End Sub

Public Function RefreshSubmittedEInvoices(ByRef Checked As Long, ByRef Changed As Long) As String
    ' Egypt: reads the result of every SUBMITTED receipt (EtaRefreshStatus). Stops at the first network problem.
    ' Returns "" or the last problem.
    Dim rs As DAO.Recordset, reply As Variant, msg As String, result As String, lastProblem As String
    Checked = 0
    Changed = 0
    If AppCountry() <> "EG" Then Exit Function          ' Saudi Arabia answers at once: nothing to follow
    Set rs = CurrentDb.OpenRecordset("SELECT DocKind, DocID FROM qryEInvoiceDocs WHERE EStatus = 'SUBMITTED' " & _
                                     "ORDER BY DocDate, DocKind DESC, DocID", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        reply = Application.Run("EtaRefreshStatus", CStr(rs!DocKind), CLng(rs!DocID))
        If Err.Number <> 0 Then
            reply = "ERROR|" & Err.Description
        End If
        On Error GoTo 0
        result = SplitResult(Nz(reply, ""), msg)
        Checked = Checked + 1
        If Nz(DbValue("SELECT EStatus FROM qryEInvoiceDocs WHERE DocKind = '" & rs!DocKind & "' AND DocID = " & _
                      rs!DocID), "") <> "SUBMITTED" Then Changed = Changed + 1
        If result = "NETWORK" Or result = "ERROR" Then
            lastProblem = msg
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
    RefreshSubmittedEInvoices = lastProblem
End Function

Public Sub EInvoicesCancelPicked(ByVal frm As Access.Form)
    ' Egypt: cancels the chosen valid e-invoice at ETA (modEtaInvoice.EtaCancelDocument).
    Dim kind As String, id As Long, reason As String, reply As Variant, msg As String, result As String
    If Not CanScreenAction(frm.Name, "EDIT") Then Exit Sub
    If AppCountry() <> "EG" Then
        ShowInfo "في السعودية يُصحَّح المستند بإشعار دائن، ولا يُلغى."
        Exit Sub
    End If
    If Not PickedDoc(frm, kind, id) Then
        ShowWarning "اختر مستندًا من القائمة."
        Exit Sub
    End If
    reason = Trim$(InputBox(Tr("سبب إلغاء المستند لدى المصلحة:"), Tr("إلغاء المستند")))
    If Len(reason) = 0 Then Exit Sub
    DoCmd.Hourglass True
    On Error Resume Next
    reply = Application.Run("EtaCancelDocument", kind, id, reason)
    If Err.Number <> 0 Then reply = "ERROR|" & Err.Description
    On Error GoTo 0
    DoCmd.Hourglass False
    result = SplitResult(Nz(reply, ""), msg)
    If result = "OK" Then
        ShowInfo "أُلغي المستند لدى المصلحة. سجّل مرتجعًا في البرنامج إن لزم."
    Else
        ShowWarning "لم يُلغَ المستند: " & msg
    End If
    EInvoicesShow frm
End Sub

Public Sub EInvoicesRefresh(ByVal frm As Access.Form)
    Dim checked As Long, changed As Long, msg As String
    If Not CanScreenAction(frm.Name, "EDIT") Then Exit Sub
    If AppCountry() <> "EG" Then
        ShowInfo "منصة فاتورة ترد عند الإرسال، فلا حاجة لتحديث الحالة."
        Exit Sub
    End If
    DoCmd.Hourglass True
    msg = RefreshSubmittedEInvoices(checked, changed)
    DoCmd.Hourglass False
    If Len(msg) > 0 Then
        ShowWarning "تعذّر تحديث الحالة: " & msg
    Else
        ShowInfo "رُوجع " & checked & " إيصال، وتغيّرت حالة " & changed & "."
    End If
    EInvoicesShow frm
End Sub

Public Function EnvironmentName() As String
    Select Case Nz(SettingValue("EInvoiceEnvironment"), "TEST")
        Case "PRODUCTION": EnvironmentName = "البيئة الفعلية"
        Case "SIMULATION": EnvironmentName = "بيئة المحاكاة"
        Case Else: EnvironmentName = "البيئة التجريبية"
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
    frm!lstLog.RowSource = Tr("SELECT l.LogID, Format(l.LoggedAt, 'yyyy/mm/dd hh:nn:ss') AS [الوقت], l.Action AS [العملية], " & _
        "l.Result AS [نتيجة الطلب], l.HttpStatus AS [رمز الرد], l.Message AS [رسالة المنظومة] FROM EInvoiceLog AS l " & _
        "WHERE l.DocKind = '" & kind & "' AND l.DocID = " & id & " ORDER BY l.LogID DESC")
End Sub

Public Sub EInvoicesSendPicked(ByVal frm As Access.Form)
    Dim kind As String, id As Long, result As String, msg As String
    If Not CanScreenAction(frm.Name, "EDIT") Then Exit Sub
    If Not PickedDoc(frm, kind, id) Then
        ShowWarning "اختر مستندًا من القائمة."
        Exit Sub
    End If
    DoCmd.Hourglass True
    result = SendEInvoice(kind, id, msg)
    DoCmd.Hourglass False
    If result = "OK" Then
        ShowInfo "تم الإرسال."
    ElseIf result = "WARNING" Then
        ShowInfo "تم الإرسال مع تحذير: " & msg
    Else
        ShowWarning "لم يُقبل المستند: " & msg
    End If
    EInvoicesShow frm
End Sub

Public Sub EInvoicesSendAll(ByVal frm As Access.Form)
    Dim sent As Long, failed As Long, msg As String
    If Not CanScreenAction(frm.Name, "EDIT") Then Exit Sub
    If Not EInvoiceEnabled() Then
        ShowWarning "الفاتورة الإلكترونية غير مفعّلة في الإعدادات."
        Exit Sub
    End If
    DoCmd.Hourglass True
    msg = SendPendingEInvoices(sent, failed)
    DoCmd.Hourglass False
    If failed = 0 Then
        ShowInfo "أُرسل " & sent & " مستند."
    Else
        ShowWarning "أُرسل " & sent & " ولم يُرسل " & failed & "." & vbCrLf & msg
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
           IsNull(JsonGet("{""a"": null}", "a")) And JsonGet("{""n"": 12.50}", "n") = "12.50", "قراءة رد JSON", _
           passed, failed, report
    Record JsonEscape("a""b\c" & vbLf) = "a\""b\\c\n", "نص داخل JSON", passed, failed, report
    Record SplitResult("REJECTED|سبب", msg) = "REJECTED" And msg = "سبب" And SplitResult("??", msg) = "ERROR", _
           "قراءة نتيجة المنظومة", passed, failed, report
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    id = Nz(DbValue("SELECT Max(SalesInvoiceID) FROM SalesInvoices"), 0)
    If id > 0 Then
        CurrentDb.Execute "UPDATE Settings SET EInvoiceEnabled = False WHERE SettingID = 1", dbFailOnError
        Record SendEInvoice("SALE", id, msg) = "ERROR" And Len(msg) > 0, "لا إرسال قبل التفعيل", passed, failed, report
        SetEInvoiceStatus "SALE", id, "REJECTED", "{""test"": 1}"
        Record DbValue("SELECT ZatcaStatus FROM SalesInvoices WHERE SalesInvoiceID = " & id) = "REJECTED", _
               "تغيير حالة المستند", passed, failed, report
        logs = Nz(DbValue("SELECT COUNT(*) FROM EInvoiceLog"), 0)
        LogEInvoice "SALE", id, "SEND", "https://example.test", 400, "REJECTED", "TEST", "{}", "{}", 5
        Record Nz(DbValue("SELECT COUNT(*) FROM EInvoiceLog"), 0) = logs + 1 And _
               Nz(DbValue("SELECT DocNumber FROM EInvoiceLog WHERE LogID = (SELECT Max(LogID) FROM EInvoiceLog)"), "") = _
               Nz(DbValue("SELECT InvoiceNumber FROM SalesInvoices WHERE SalesInvoiceID = " & id), "-"), _
               "سجل الإرسال برقم المستند", passed, failed, report
        CurrentDb.Execute "UPDATE SalesInvoices SET ZatcaStatus = 'PENDING', InvoiceDate = DateAdd('h', -30, Now()) " & _
                          "WHERE SalesInvoiceID = " & id, dbFailOnError
        Record EInvoiceOverdueCount() >= 1, "تنبيه المستندات المتأخرة", passed, failed, report
    Else
        Debug.Print "[--] لا توجد فواتير بيع"
    End If
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    Record False, "خطأ: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات الفاتورة الإلكترونية ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, _
                "TestEInvoice"
        TestEInvoice = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestEInvoice"
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
