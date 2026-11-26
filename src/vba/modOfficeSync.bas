Attribute VB_Name = "modOfficeSync"
'==============================================================================
' modOfficeSync  -  Retail Store Management System
'
' Plan 6H (docs/84): Access is kept for the office. The web, the Windows POS and the rep app write their
' documents to the same database but leave two things to Access: the journal (SyncJournal) and the sending
' of the e-invoices (modEInvoice). One Access computer of the office does both:
'   Settings.OfficeSyncComputer  the name of that computer (empty = by the button only)
'   Settings.OfficeSyncMinutes   every how many minutes (0 = by the button only)
'   OfficeSyncTick               from the dashboard timer (modDashboard.DashboardRefresh): runs when this is
'                                the office computer and the time has come
'   RunOfficeSync                SyncJournal, then the pending e-invoices sent (SendPendingEInvoices) and,
'                                in Egypt, the submitted ones followed (RefreshSubmittedEInvoices);
'                                each run is a row of OfficeSyncLog and Settings.OfficeSyncLastAt / LastResult
' The web reads OfficeSyncLastAt and warns its administrators when the office tasks stopped.
' Screen: frmOfficeSync (permission OFFICE_SYNC, from the backup screen).
'==============================================================================
Option Compare Database
Option Explicit

Private m_running As Boolean

'==============================================================================
' Rules
'==============================================================================
Public Function ThisComputerName() As String
    ThisComputerName = Left$(Environ$("COMPUTERNAME"), 64)
End Function

Public Function IsOfficeSyncComputer() As Boolean
    ' This computer is the one named in the settings (not case sensitive).
    Dim named As String
    named = Trim$(Nz(SettingValue("OfficeSyncComputer"), ""))
    If Len(named) = 0 Then Exit Function
    IsOfficeSyncComputer = (StrComp(named, ThisComputerName(), vbTextCompare) = 0)
End Function

Public Function OfficeSyncDue(ByVal LastAt As Variant, ByVal Minutes As Long, ByVal NowAt As Date) As Boolean
    ' The automatic run is due: never ran, or Minutes have passed since the last run. 0 minutes = never.
    If Minutes <= 0 Then Exit Function
    If IsNull(LastAt) Then
        OfficeSyncDue = True
        Exit Function
    End If
    If Not IsDate(LastAt) Then
        OfficeSyncDue = True
        Exit Function
    End If
    OfficeSyncDue = (WholeMinutes(CDate(LastAt), NowAt) >= Minutes)
End Function

Public Function OfficeSyncStale(ByVal LastAt As Variant, ByVal Minutes As Long, ByVal NowAt As Date) As Boolean
    ' The office tasks stopped: more than twice the period and 10 minutes without a run (the web warns).
    If Minutes <= 0 Then Exit Function
    If IsNull(LastAt) Then
        OfficeSyncStale = True
        Exit Function
    End If
    OfficeSyncStale = (WholeMinutes(CDate(LastAt), NowAt) > 2 * Minutes + 10)
End Function

Public Function WholeMinutes(ByVal FromAt As Date, ByVal ToAt As Date) As Long
    ' The minute boundaries crossed (= OfficeSyncRules.WholeMinutes), counted on whole seconds as integers.
    ' A date is a Double: 10:00 less 15 minutes is not exactly 9:45, and DateDiff("n") on it was a minute off.
    WholeMinutes = CLng(MinuteIndex(ToAt) - MinuteIndex(FromAt))
End Function

Public Function MinuteIndex(ByVal At As Date) As Double
    Dim secs As Double
    secs = Round(CDbl(At) * 86400#, 0)
    MinuteIndex = Int(secs / 60#)
End Function

Public Function OfficeSyncMinutes() As Long
    OfficeSyncMinutes = CLng(Nz(SettingValue("OfficeSyncMinutes"), 0))
End Function

Public Function PendingEInvoiceCount() As Long
    ' Documents waiting for the platform, whoever made them (Access, the web, a POS device).
    If Not EInvoiceEnabled() Then Exit Function
    PendingEInvoiceCount = CLng(Nz(DbValue("SELECT Count(*) FROM qryEInvoiceDocs WHERE EStatus = 'PENDING'"), 0))
End Function

Private Function JoinProblem(ByVal First As String, ByVal Second As String) As String
    If Len(First) = 0 Then
        JoinProblem = Second
    ElseIf Len(Second) = 0 Then
        JoinProblem = First
    Else
        JoinProblem = First & " - " & Second
    End If
End Function

'==============================================================================
' The run
'==============================================================================
Public Sub OfficeSyncTick()
    ' The dashboard timer (every 5 minutes) and its Activate: the automatic run of the office computer.
    ' Never shows a message: the result is in OfficeSyncLog and the screen.
    On Error GoTo Done
    If m_running Then Exit Sub
    If CurrentUserID() = 0 Then Exit Sub
    If Not IsOfficeSyncComputer() Then Exit Sub
    If Not OfficeSyncDue(SettingValue("OfficeSyncLastAt"), OfficeSyncMinutes(), Now) Then Exit Sub
    RunOfficeSync "AUTO"
Done:
End Sub

Public Function RunOfficeSync(ByVal RunKind As String, Optional ByRef Summary As String) As String
    ' Returns "" or the problem; Summary = what was done, in Arabic.
    Dim added As Long, updated As Long, removed As Long, sent As Long, notSent As Long
    Dim checked As Long, changed As Long, problem As String, msg As String
    Summary = ""
    If m_running Then
        RunOfficeSync = "مهام المكتب تعمل الآن."
        Exit Function
    End If
    m_running = True
    On Error GoTo EH
    problem = SyncJournal(added, updated, removed)
    If EInvoiceEnabled() Then
        msg = EInvoiceReadyProblem()
        If Len(msg) > 0 Then
            If PendingEInvoiceCount() > 0 Then problem = JoinProblem(problem, msg)
        Else
            msg = SendPendingEInvoices(sent, notSent)
            problem = JoinProblem(problem, msg)
            msg = RefreshSubmittedEInvoices(checked, changed)
            problem = JoinProblem(problem, msg)
        End If
    End If
Finish:
    On Error Resume Next
    Summary = "القيود: " & added & " جديد، " & updated & " معدّل، " & removed & " محذوف. " & _
              "الفواتير الإلكترونية: أُرسلت " & sent & "، لم تُرسل " & notSent & "، تمت متابعة " & checked & "."
    WriteOfficeSyncLog RunKind, added, updated, removed, sent, notSent, checked, problem
    On Error GoTo 0
    m_running = False
    RunOfficeSync = problem
    Exit Function
EH:
    problem = JoinProblem(problem, "تعذّر إكمال مهام المكتب: " & Err.Description)
    Resume Finish
End Function

Private Sub WriteOfficeSyncLog(ByVal RunKind As String, ByVal added As Long, ByVal updated As Long, _
                               ByVal removed As Long, ByVal sent As Long, ByVal notSent As Long, _
                               ByVal checked As Long, ByVal problem As String)
    Dim db As DAO.Database, stamp As Date
    Set db = CurrentDb
    stamp = Now
    db.Execute "INSERT INTO OfficeSyncLog (RunAt, ComputerName, RunKind, JournalAdded, JournalUpdated, JournalRemoved, " & _
               "EInvoicesSent, EInvoicesFailed, EInvoicesChecked, Problem, EmployeeID) VALUES (" & SqlDate(stamp) & ", " & _
               SqlText(ThisComputerName()) & ", " & SqlText(RunKind) & ", " & added & ", " & updated & ", " & removed & ", " & _
               sent & ", " & notSent & ", " & checked & ", " & SqlText(Left$(problem, 255)) & ", " & _
               IIf(CurrentUserID() = 0, "Null", CStr(CurrentUserID())) & ")", dbFailOnError + dbSeeChanges
    db.Execute "UPDATE Settings SET OfficeSyncLastAt = " & SqlDate(stamp) & ", OfficeSyncLastResult = " & _
               SqlText(Left$(problem, 255)) & " WHERE SettingID = 1", dbFailOnError + dbSeeChanges
End Sub

'==============================================================================
' Screen frmOfficeSync
'==============================================================================
Public Sub OfficeSyncLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    frm!txtComputer.Value = Nz(SettingValue("OfficeSyncComputer"), "")
    frm!txtMinutes.Value = OfficeSyncMinutes()
    OfficeSyncShow frm
End Sub

Public Sub OfficeSyncShow(ByVal frm As Access.Form)
    Dim lastAt As Variant, lastResult As String, minutes As Long, state As String, waiting As Long
    Calendar = vbCalGreg
    frm!lblThisComputer.Caption = Tr("هذا الجهاز: " & ThisComputerName())
    lastAt = SettingValue("OfficeSyncLastAt")
    lastResult = Nz(SettingValue("OfficeSyncLastResult"), "")
    minutes = OfficeSyncMinutes()
    If IsNull(lastAt) Then
        state = "لم تعمل مهام المكتب بعد."
    Else
        state = "آخر تشغيل: " & GDate(lastAt, True)
        If Len(lastResult) > 0 Then state = state & " - " & lastResult
    End If
    If Len(Trim$(Nz(SettingValue("OfficeSyncComputer"), ""))) = 0 Or minutes = 0 Then
        state = state & " (التشغيل التلقائي متوقف)"
    ElseIf Not IsOfficeSyncComputer() Then
        state = state & " (تعمل تلقائيًا على جهاز آخر)"
    End If
    frm!lblState.Caption = Tr(state)
    If OfficeSyncStale(lastAt, minutes, Now) Or Len(lastResult) > 0 Then
        frm!lblState.ForeColor = CLR_DANGER
    Else
        frm!lblState.ForeColor = CLR_SUCCESS
    End If
    waiting = PendingEInvoiceCount()
    If EInvoiceEnabled() Then
        frm!lblWaiting.Caption = Tr("فواتير إلكترونية تنتظر الإرسال: " & waiting)
    Else
        frm!lblWaiting.Caption = Tr("الفاتورة الإلكترونية غير مفعّلة: المهام تحدّث القيود فقط.")
    End If
    frm!lstLog.RowSource = Tr("SELECT TOP 200 GDate(RunAt, True) AS [الوقت], ComputerName AS [الجهاز], " & _
        "IIf(RunKind = 'AUTO', 'تلقائي', 'يدوي') AS [الطريقة], JournalAdded AS [قيود جديدة], " & _
        "JournalUpdated AS [قيود معدّلة], JournalRemoved AS [قيود محذوفة], EInvoicesSent AS [أُرسلت], " & _
        "EInvoicesFailed AS [لم تُرسل], Nz(Problem, 'بلا مشكلة') AS [نتيجة التشغيل] FROM OfficeSyncLog " & _
        "ORDER BY RunAt DESC, OfficeSyncLogID DESC")
End Sub

Public Sub OfficeSyncUseThisComputer(ByVal frm As Access.Form)
    frm!txtComputer.Value = ThisComputerName()
End Sub

Public Sub OfficeSyncSave(ByVal frm As Access.Form)
    Dim minutes As Variant, before As Collection
    minutes = frm!txtMinutes.Value
    If Not IsNumeric(Nz(minutes, "x")) Then
        ShowWarning "اكتب عدد الدقائق من 0 إلى 1440."
        Exit Sub
    End If
    If CDbl(minutes) < 0 Or CDbl(minutes) > 1440 Or CDbl(minutes) <> Int(CDbl(minutes)) Then
        ShowWarning "اكتب عدد الدقائق من 0 إلى 1440."
        Exit Sub
    End If
    Set before = AuditSnapshot("Settings", "SettingID", 1)
    CurrentDb.Execute "UPDATE Settings SET OfficeSyncComputer = " & SqlText(Trim$(Nz(frm!txtComputer.Value, ""))) & _
                      ", OfficeSyncMinutes = " & CLng(minutes) & " WHERE SettingID = 1", dbFailOnError + dbSeeChanges
    AuditEdited "OFFICE_SYNC_SETTINGS", "Settings", "SettingID", 1, before
    ShowInfo "تم حفظ إعداد مهام المكتب."
    OfficeSyncShow frm
End Sub

Public Sub OfficeSyncRunNow(ByVal frm As Access.Form)
    Dim problem As String, summary As String
    DoCmd.Hourglass True
    problem = RunOfficeSync("MANUAL", summary)
    DoCmd.Hourglass False
    OfficeSyncShow frm
    If Len(problem) > 0 Then
        ShowWarning summary & vbCrLf & vbCrLf & problem
    Else
        ShowInfo summary
    End If
End Sub

'==============================================================================
' Test (inside a transaction that is rolled back)
'==============================================================================
Public Function TestOfficeSync() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, passed As Long, failed As Long, report As String
    Dim msg As String, summary As String, before As Long, t0 As Date
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestOfficeSync  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    t0 = DateSerial(2026, 11, 16) + TimeSerial(10, 0, 0)
    Check OfficeSyncDue(Null, 15, t0) And Not OfficeSyncDue(Null, 0, t0), "أول تشغيل مستحق، و0 دقيقة = يدويًا فقط", _
          passed, failed, report
    Check Not OfficeSyncDue(t0 - TimeSerial(0, 14, 0), 15, t0) And OfficeSyncDue(t0 - TimeSerial(0, 15, 0), 15, t0), _
          "التشغيل التالي بعد عدد الدقائق", passed, failed, report
    Check Not OfficeSyncStale(t0 - TimeSerial(0, 40, 0), 15, t0) And OfficeSyncStale(t0 - TimeSerial(0, 41, 0), 15, t0), _
          "التوقف بعد ضعف المدة وعشر دقائق", passed, failed, report
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "UPDATE Settings SET OfficeSyncComputer = 'NOT-THIS-PC-TEST', OfficeSyncMinutes = 15, " & _
                      "OfficeSyncLastAt = Null, EInvoiceEnabled = 0 WHERE SettingID = 1", dbFailOnError + dbSeeChanges
    before = CLng(Nz(DbValue("SELECT Count(*) FROM OfficeSyncLog"), 0))
    Check Not IsOfficeSyncComputer(), "جهاز آخر ليس جهاز مهام المكتب", passed, failed, report
    OfficeSyncTick
    Check CLng(Nz(DbValue("SELECT Count(*) FROM OfficeSyncLog"), 0)) = before, "لا تشغيل تلقائي على جهاز آخر", _
          passed, failed, report
    CurrentDb.Execute "UPDATE Settings SET OfficeSyncComputer = " & SqlText(LCase$(ThisComputerName())) & _
                      " WHERE SettingID = 1", dbFailOnError + dbSeeChanges
    Check IsOfficeSyncComputer(), "هذا الجهاز هو جهاز مهام المكتب (بلا تفرقة بين الحروف)", passed, failed, report
    OfficeSyncTick
    Check CLng(Nz(DbValue("SELECT Count(*) FROM OfficeSyncLog WHERE RunKind = 'AUTO'"), 0)) >= 1 And _
          Not IsNull(SettingValue("OfficeSyncLastAt")), "التشغيل التلقائي يكتب سجله ووقته", passed, failed, report
    before = CLng(Nz(DbValue("SELECT Count(*) FROM OfficeSyncLog"), 0))
    OfficeSyncTick
    Check CLng(Nz(DbValue("SELECT Count(*) FROM OfficeSyncLog"), 0)) = before, "لا تشغيل قبل مرور المدة", _
          passed, failed, report
    msg = RunOfficeSync("MANUAL", summary)
    Check Len(summary) > 0 And CLng(Nz(DbValue("SELECT Count(*) FROM OfficeSyncLog"), 0)) = before + 1 And _
          Nz(DbValue("SELECT Count(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0) = 0, _
          "التشغيل اليدوي يحدّث القيود ويكتب سجله " & msg, passed, failed, report
Undo:
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
        TestMsg "جميع اختبارات مهام المكتب ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestOfficeSync"
        TestOfficeSync = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestOfficeSync"
    End If
End Function

Private Sub Check(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
                  ByRef report As String)
    If ok Then
        passed = passed + 1
        Debug.Print "[OK] " & Title
    Else
        failed = failed + 1
        Debug.Print "[XX] " & Title
        report = report & "- " & Title & vbCrLf
    End If
End Sub
