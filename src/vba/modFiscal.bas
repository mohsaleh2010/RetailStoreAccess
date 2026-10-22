Attribute VB_Name = "modFiscal"
'==============================================================================
' modFiscal  -  Retail Store Management System
'
' The fiscal periods (docs/54-Fiscal-Periods.md, screen frmFiscalPeriods, permission PERIOD_CLOSE):
'   FiscalPeriods      the years (or periods) the accountant adds, in order: each one starts the
'                      day after the end of the one before, so there is no overlap and no gap.
'   The fiscal key     a closing (FiscalYearClosings.FiscalYear) is keyed by the year its period
'                      ends in, so two periods never end in the same calendar year. Before the first
'                      period the calendar years stay (key = the year, 1 January - 31 December):
'                      the first period starts on 1 January, or before the first entry.
'   FiscalYearStart / FiscalYearEnd   the days of a key, used by the year closing (modClosing),
'                      the quick periods "this year / last year" (SetQuickPeriod), the opening
'                      balances report and the opening balance sheet of a period.
'==============================================================================
Option Compare Database
Option Explicit

'------------------------------------------------------------------------------
' The periods and their keys
'------------------------------------------------------------------------------
Public Function FirstPeriodStart() As Variant
    FirstPeriodStart = DbValue("SELECT Min(StartDate) FROM FiscalPeriods")
End Function

Public Function LastPeriodEnd() As Variant
    LastPeriodEnd = DbValue("SELECT Max(EndDate) FROM FiscalPeriods")
End Function

Public Function FiscalYearStart(ByVal FiscalKey As Long) As Date
    ' The first day of the period whose key is FiscalKey; 1 January for a calendar year.
    Dim v As Variant
    v = DbValue("SELECT StartDate FROM FiscalPeriods WHERE Year(EndDate) = " & FiscalKey)
    If IsDate(v) Then
        FiscalYearStart = DateValue(v)
    Else
        FiscalYearStart = DateSerial(FiscalKey, 1, 1)
    End If
End Function

Public Function FiscalYearEnd(ByVal FiscalKey As Long) As Date
    ' The last day of the period whose key is FiscalKey; 31 December for a calendar year.
    Dim v As Variant
    v = DbValue("SELECT EndDate FROM FiscalPeriods WHERE Year(EndDate) = " & FiscalKey)
    If IsDate(v) Then
        FiscalYearEnd = DateValue(v)
    Else
        FiscalYearEnd = DateSerial(FiscalKey, 12, 31)
    End If
End Function

Public Function FiscalKeyName(ByVal FiscalKey As Long) As String
    Dim v As Variant
    v = DbValue("SELECT PeriodName FROM FiscalPeriods WHERE Year(EndDate) = " & FiscalKey)
    If IsNull(v) Then
        FiscalKeyName = CStr(FiscalKey)
    Else
        FiscalKeyName = v
    End If
End Function

Public Function FiscalKeys() As String
    ' Every key, oldest first, separated by ";": the calendar years with entries before the first
    ' period (or until last year when there is no period), then the periods.
    Dim y As Long, lastCalendar As Long, rs As DAO.Recordset, keys As String, first As Variant
    first = FirstPeriodStart()
    If IsDate(first) Then
        lastCalendar = Year(first) - 1
    Else
        lastCalendar = Year(Date) - 1
    End If
    For y = FirstEntryYear() To lastCalendar
        keys = keys & IIf(Len(keys) > 0, ";", "") & y
    Next
    Set rs = CurrentDb.OpenRecordset("SELECT EndDate FROM FiscalPeriods ORDER BY StartDate", dbOpenSnapshot)
    Do Until rs.EOF
        keys = keys & IIf(Len(keys) > 0, ";", "") & Year(rs!EndDate)
        rs.MoveNext
    Loop
    rs.Close
    FiscalKeys = keys
End Function

Public Function FiscalKeyKnown(ByVal FiscalKey As Long) As Boolean
    FiscalKeyKnown = InStr(";" & FiscalKeys() & ";", ";" & FiscalKey & ";") > 0
End Function

Public Function FiscalPeriodOf(ByVal When As Date, ByRef StartDay As Date, ByRef EndDay As Date) As Boolean
    ' The period that holds the day When (True), else its calendar year (False).
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT StartDate, EndDate FROM FiscalPeriods WHERE StartDate <= " & _
                                     SqlDate(DateValue(When)) & " AND EndDate >= " & SqlDate(DateValue(When)), _
                                     dbOpenSnapshot, dbSeeChanges)
    If rs.EOF Then
        StartDay = DateSerial(Year(When), 1, 1)
        EndDay = DateSerial(Year(When), 12, 31)
    Else
        StartDay = DateValue(rs!StartDate)
        EndDay = DateValue(rs!EndDate)
        FiscalPeriodOf = True
    End If
    rs.Close
End Function

'------------------------------------------------------------------------------
' Adding and deleting a period
'------------------------------------------------------------------------------
Public Function CanFiscal(ByVal Action As String) As String
    If Not HasPermission("PERIOD_CLOSE") Then
        CanFiscal = "لا تملك صلاحية الفترات المالية (إقفال الفترات)."
    ElseIf Not CanScreenAction("frmFiscalPeriods", Action, True) Then
        CanFiscal = "لا تملك صلاحية " & IIf(Action = "ADD", "الإضافة", "الحذف") & " في شاشة الفترات المالية."
    End If
End Function

Public Function FiscalPeriodProblem(ByVal PeriodName As String, ByVal StartDay As Variant, ByVal EndDay As Variant) As String
    ' "" when the period can be added after the last one.
    Dim last As Variant, closedEnd As Variant, earliest As Variant
    If Len(Trim$(PeriodName)) = 0 Then
        FiscalPeriodProblem = "اكتب اسم الفترة."
        Exit Function
    End If
    If Not IsDate(StartDay) Or Not IsDate(EndDay) Then
        FiscalPeriodProblem = "أدخل بداية الفترة ونهايتها."
        Exit Function
    End If
    StartDay = DateValue(StartDay)
    EndDay = DateValue(EndDay)
    If EndDay < StartDay Then
        FiscalPeriodProblem = "نهاية الفترة لا تكون قبل بدايتها."
    ElseIf DateAdd("m", 24, StartDay) <= EndDay Then
        FiscalPeriodProblem = "الفترة المالية لا تزيد على 24 شهرًا."
    ElseIf Not IsNull(DbValue("SELECT FiscalPeriodID FROM FiscalPeriods WHERE PeriodName = " & SqlText(Trim$(PeriodName)))) Then
        FiscalPeriodProblem = "توجد فترة بهذا الاسم."
    End If
    If Len(FiscalPeriodProblem) > 0 Then Exit Function
    last = LastPeriodEnd()
    If IsDate(last) Then
        If StartDay <> DateAdd("d", 1, DateValue(last)) Then
            FiscalPeriodProblem = "تبدأ الفترة الجديدة في " & GDate(DateAdd("d", 1, DateValue(last))) & _
                                  "، اليوم التالي لنهاية آخر فترة: لا تداخل ولا فجوة بين الفترات."
        ElseIf Year(EndDay) = Year(last) Then
            FiscalPeriodProblem = "لا تنتهي فترتان في سنة ميلادية واحدة: إقفال الفترة يُعرف بسنة نهايتها."
        End If
        Exit Function
    End If
    ' the first period: after the last closed year; on 1 January, or before the first entry
    closedEnd = DbValue("SELECT Max(ClosingDate) FROM FiscalYearClosings")
    If IsDate(closedEnd) Then
        If StartDay <= DateValue(closedEnd) Then
            FiscalPeriodProblem = "تبدأ الفترة الأولى بعد آخر سنة مقفلة (" & GDate(closedEnd) & ")."
            Exit Function
        End If
    End If
    If Day(StartDay) <> 1 Or Month(StartDay) <> 1 Then
        earliest = DbValue("SELECT Min(EntryDate) FROM JournalEntries WHERE SourceType <> 'YEAR_CLOSE'")
        If IsDate(earliest) Then
            If DateValue(earliest) < StartDay Then
                FiscalPeriodProblem = "توجد قيود قبل " & GDate(StartDay) & " (أولها " & GDate(earliest) & "). " & _
                    "ابدأ الفترة الأولى في 1 يناير، فتبقى السنوات قبلها ميلادية، أو قبل أول قيد."
            End If
        End If
    End If
End Function

Public Function AddFiscalPeriod(ByVal PeriodName As String, ByVal StartDay As Variant, ByVal EndDay As Variant, _
                                Optional ByVal Notes As String = "") As String
    AddFiscalPeriod = CanFiscal("ADD")
    If Len(AddFiscalPeriod) = 0 Then AddFiscalPeriod = FiscalPeriodProblem(PeriodName, StartDay, EndDay)
    If Len(AddFiscalPeriod) > 0 Then Exit Function
    CurrentDb.Execute "INSERT INTO FiscalPeriods (PeriodName, StartDate, EndDate, Notes) VALUES (" & _
                      SqlText(Trim$(PeriodName)) & ", " & SqlDate(DateValue(StartDay)) & ", " & _
                      SqlDate(DateValue(EndDay)) & ", " & SqlText(Trim$(Notes)) & ")", dbFailOnError + dbSeeChanges
    LogAction "FISCAL_PERIOD_ADD", "FiscalPeriods", Trim$(PeriodName), GDate(StartDay) & " - " & GDate(EndDay)
End Function

Public Function DeleteLastFiscalPeriod() As String
    ' Only the last period, and only while it is not closed.
    Dim id As Variant, endDay As Variant, before As Collection
    DeleteLastFiscalPeriod = CanFiscal("DELETE")
    If Len(DeleteLastFiscalPeriod) > 0 Then Exit Function
    id = DbValue("SELECT TOP 1 FiscalPeriodID FROM FiscalPeriods ORDER BY StartDate DESC")
    If IsNull(id) Then
        DeleteLastFiscalPeriod = "لا توجد فترات."
        Exit Function
    End If
    endDay = DbValue("SELECT EndDate FROM FiscalPeriods WHERE FiscalPeriodID = " & id)
    If YearIsClosed(Year(endDay)) Then
        DeleteLastFiscalPeriod = "الفترة مقفلة. أعد فتحها أولًا من شاشة إقفال الفترات."
        Exit Function
    End If
    Set before = AuditSnapshot("FiscalPeriods", "FiscalPeriodID", id)
    CurrentDb.Execute "DELETE FROM FiscalPeriods WHERE FiscalPeriodID = " & CLng(id), dbFailOnError + dbSeeChanges
    AuditDeleted "FISCAL_PERIOD_DELETE", "FiscalPeriods", id, before
End Function

'------------------------------------------------------------------------------
' Opening balances of a period
'------------------------------------------------------------------------------
Public Sub PrintOpeningBalances(ByVal StartDay As Date, Optional ByVal Title As String = "")
    ' Every balance at the start of StartDay: the accounts, the customers, the suppliers, the stock.
    Dim msg As String, criteria As String
    If Not HasPermission("JOURNAL") Then
        ShowWarning "التقارير المحاسبية تحتاج صلاحية «قيود اليومية»."
        Exit Sub
    End If
    msg = SyncJournal()
    If Len(msg) > 0 Then ShowWarning msg
    SetPeriod StartDay, StartDay
    criteria = "في بداية يوم " & GDate(StartDay)
    If Len(Title) > 0 Then criteria = Title & "    " & criteria
    OpenReportOrQuery "rptOpeningBalances", "OpeningBalancesReportQuery", "", criteria
End Sub

Public Sub PrintOpeningBalanceSheet(ByVal FiscalKey As Long)
    ' The balance sheet on the day before the period starts.
    Dim msg As String, criteria As String, asOf As Date
    If FiscalKey = 0 Then
        ShowWarning "اختر السنة أو الفترة المالية."
        Exit Sub
    End If
    If Not HasPermission("JOURNAL") Then
        ShowWarning "التقارير المحاسبية تحتاج صلاحية «قيود اليومية»."
        Exit Sub
    End If
    msg = SyncJournal()
    If Len(msg) > 0 Then ShowWarning msg
    asOf = DateAdd("d", -1, FiscalYearStart(FiscalKey))
    criteria = PrepareFinancials("BALANCE_SHEET", asOf, asOf)
    OpenReportOrQuery "rptBalanceSheet", "BalanceSheetQuery", "", _
                      "الميزانية الافتتاحية لـ " & FiscalKeyName(FiscalKey) & "    " & criteria
End Sub

'------------------------------------------------------------------------------
' Screen frmFiscalPeriods
'------------------------------------------------------------------------------
Public Sub FiscalPeriodsLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    FiscalPeriodsRefresh frm
End Sub

Public Sub FiscalPeriodsRefresh(ByVal frm As Access.Form)
    Dim last As Variant, startDay As Date, endDay As Date, first As Variant
    frm!lstPeriods.RowSource = Tr("SELECT f.FiscalPeriodID, f.PeriodName AS [الفترة], " & _
        "Format(f.StartDate, 'yyyy/mm/dd') AS [من], Format(f.EndDate, 'yyyy/mm/dd') AS [إلى], " & _
        "IIf((SELECT COUNT(*) FROM FiscalYearClosings AS c WHERE c.FiscalYear = Year(f.EndDate)) > 0, " & _
        "'مقفلة', 'مفتوحة') AS [الحالة] FROM FiscalPeriods AS f ORDER BY f.StartDate")
    last = LastPeriodEnd()
    If IsDate(last) Then
        startDay = DateAdd("d", 1, DateValue(last))
        frm!txtStart.Locked = True
        frm!lblHint.Caption = Tr("الفترة الجديدة تبدأ في اليوم التالي لنهاية آخر فترة، ولا تنتهي فترتان في سنة ميلادية " & _
                                 "واحدة. تُحذف آخر فترة فقط، وهي غير مقفلة.")
    Else
        first = DbValue("SELECT Min(EntryDate) FROM JournalEntries WHERE SourceType <> 'YEAR_CLOSE'")
        If IsDate(first) Then
            startDay = DateSerial(Year(first), 1, 1)
        Else
            startDay = DateSerial(Year(Date), 1, 1)
        End If
        frm!txtStart.Locked = False
        frm!lblHint.Caption = Tr("الفترة الأولى تبدأ في 1 يناير (والسنوات قبلها ميلادية) أو قبل أول قيد. " & _
                                 "بعدها تُضاف الفترات بالترتيب، ولا تزيد الفترة على 24 شهرًا.")
    End If
    endDay = DateAdd("d", -1, DateAdd("yyyy", 1, startDay))
    frm!txtStart.Value = startDay
    frm!txtEnd.Value = endDay
    If Year(startDay) = Year(endDay) Then
        frm!txtName.Value = CStr(Year(endDay))
    Else
        frm!txtName.Value = Year(startDay) & "/" & Year(endDay)
    End If
End Sub

Public Sub FiscalAddPeriod(ByVal frm As Access.Form)
    Dim msg As String
    msg = AddFiscalPeriod(Nz(frm!txtName.Value, ""), frm!txtStart.Value, frm!txtEnd.Value)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    ShowInfo "تمت إضافة الفترة " & frm!txtName.Value & "."
    FiscalPeriodsRefresh frm
End Sub

Public Sub FiscalDeleteLast(ByVal frm As Access.Form)
    Dim msg As String
    If Not AskYesNo("حذف آخر فترة مالية؟") Then Exit Sub
    msg = DeleteLastFiscalPeriod()
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    FiscalPeriodsRefresh frm
End Sub

Public Sub FiscalPrintSelected(ByVal frm As Access.Form, ByVal Sheet As Boolean)
    Dim id As Variant, key As Long
    id = frm!lstPeriods.Value
    If IsNull(id) Then
        ShowWarning "اختر فترة من القائمة."
        Exit Sub
    End If
    key = Year(DbValue("SELECT EndDate FROM FiscalPeriods WHERE FiscalPeriodID = " & id))
    If Sheet Then
        PrintOpeningBalanceSheet key
    Else
        PrintOpeningBalances FiscalYearStart(key), FiscalKeyName(key)
    End If
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Public Function TestFiscalPeriods() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, passed As Long, failed As Long, report As String
    Dim lastEnd As Variant, firstDay As Date, s As Date, e As Date, key As Long
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestFiscalPeriods  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Record FiscalPeriodProblem("", Date, Date) <> "", "اسم الفترة مطلوب", passed, failed, report
    Record FiscalPeriodProblem("X", Date, DateAdd("d", -1, Date)) <> "", "النهاية لا تكون قبل البداية", passed, failed, report
    Record FiscalPeriodProblem("X", Date, DateAdd("m", 25, Date)) <> "", "لا تزيد على 24 شهرًا", passed, failed, report
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    lastEnd = LastPeriodEnd()
    If IsDate(lastEnd) Then
        firstDay = DateAdd("d", 1, DateValue(lastEnd))
    Else
        firstDay = DateSerial(Year(Date) + 5, 1, 1)                ' far after every entry and closing
        CurrentDb.Execute "INSERT INTO FiscalPeriods (PeriodName, StartDate, EndDate) VALUES ('TEST-0', " & _
                          SqlDate(DateSerial(Year(Date) - 30, 1, 1)) & ", " & SqlDate(DateAdd("d", -1, firstDay)) & ")", _
                          dbFailOnError + dbSeeChanges                              ' a first period, so the rules of the next apply
    End If
    Record FiscalPeriodProblem("TEST-GAP", DateAdd("d", 1, firstDay), DateAdd("yyyy", 1, firstDay)) <> "", _
           "رفض الفجوة بعد آخر فترة", passed, failed, report
    Record FiscalPeriodProblem("TEST-OVERLAP", DateAdd("d", -1, firstDay), DateAdd("yyyy", 1, firstDay)) <> "", _
           "رفض التداخل مع آخر فترة", passed, failed, report
    s = firstDay
    e = DateAdd("d", -1, DateAdd("yyyy", 1, s))
    Record AddFiscalPeriod("TEST-FP", s, e) = "", "إضافة فترة بعد آخر فترة", passed, failed, report
    key = Year(e)
    Record FiscalYearStart(key) = s And FiscalYearEnd(key) = e And FiscalKeyName(key) = "TEST-FP", _
           "بداية الفترة ونهايتها واسمها بمفتاحها", passed, failed, report
    Record FiscalKeyKnown(key), "الفترة في قائمة الإقفال", passed, failed, report
    Record FiscalPeriodOf(DateAdd("d", 10, s), s, e) And s = FiscalYearStart(key), "الفترة التي تحتوي يومًا", _
           passed, failed, report
    Record YearCloseProblem(key) <> "" Or FiscalYearEnd(key) < Date, "لا تُقفل فترة قبل انتهائها", passed, failed, report
    Record DeleteLastFiscalPeriod() = "" And Not FiscalKeyKnown(key), "حذف آخر فترة", passed, failed, report
    Record FiscalYearStart(2001) = DateSerial(2001, 1, 1) Or Not IsNull(DbValue("SELECT FiscalPeriodID FROM " & _
           "FiscalPeriods WHERE Year(EndDate) = 2001")), "السنة الميلادية بلا فترة", passed, failed, report
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
        TestMsg "جميع اختبارات الفترات المالية ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, _
                "TestFiscalPeriods"
        TestFiscalPeriods = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
                "TestFiscalPeriods"
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
