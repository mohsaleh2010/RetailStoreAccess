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
                                     dbOpenSnapshot)
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
        CanFiscal = "·«  „·ﬂ ’·«ÕÌ… «·› —«  «·„«·Ì… (≈ﬁ›«· «·› —« )."
    ElseIf Not CanScreenAction("frmFiscalPeriods", Action, True) Then
        CanFiscal = "·«  „·ﬂ ’·«ÕÌ… " & IIf(Action = "ADD", "«·≈÷«›…", "«·Õ–›") & " ›Ì ‘«‘… «·› —«  «·„«·Ì…."
    End If
End Function

Public Function FiscalPeriodProblem(ByVal PeriodName As String, ByVal StartDay As Variant, ByVal EndDay As Variant) As String
    ' "" when the period can be added after the last one.
    Dim last As Variant, closedEnd As Variant, earliest As Variant
    If Len(Trim$(PeriodName)) = 0 Then
        FiscalPeriodProblem = "«ﬂ » «”„ «·› —…."
        Exit Function
    End If
    If Not IsDate(StartDay) Or Not IsDate(EndDay) Then
        FiscalPeriodProblem = "√œŒ· »œ«Ì… «·› —… Ê‰Â«Ì Â«."
        Exit Function
    End If
    StartDay = DateValue(StartDay)
    EndDay = DateValue(EndDay)
    If EndDay < StartDay Then
        FiscalPeriodProblem = "‰Â«Ì… «·› —… ·«  ﬂÊ‰ ﬁ»· »œ«Ì Â«."
    ElseIf DateAdd("m", 24, StartDay) <= EndDay Then
        FiscalPeriodProblem = "«·› —… «·„«·Ì… ·«  “Ìœ ⁄·Ï 24 ‘Â—«."
    ElseIf Not IsNull(DbValue("SELECT FiscalPeriodID FROM FiscalPeriods WHERE PeriodName = " & SqlText(Trim$(PeriodName)))) Then
        FiscalPeriodProblem = " ÊÃœ › —… »Â–« «·«”„."
    End If
    If Len(FiscalPeriodProblem) > 0 Then Exit Function
    last = LastPeriodEnd()
    If IsDate(last) Then
        If StartDay <> DateAdd("d", 1, DateValue(last)) Then
            FiscalPeriodProblem = " »œ√ «·› —… «·ÃœÌœ… ›Ì " & GDate(DateAdd("d", 1, DateValue(last))) & _
                                  "° «·ÌÊ„ «· «·Ì ·‰Â«Ì… ¬Œ— › —…: ·«  œ«Œ· Ê·« ›ÃÊ… »Ì‰ «·› —« ."
        ElseIf Year(EndDay) = Year(last) Then
            FiscalPeriodProblem = "·«  ‰ ÂÌ › — «‰ ›Ì ”‰… „Ì·«œÌ… Ê«Õœ…: ≈ﬁ›«· «·› —… Ìı⁄—› »”‰… ‰Â«Ì Â«."
        End If
        Exit Function
    End If
    ' the first period: after the last closed year; on 1 January, or before the first entry
    closedEnd = DbValue("SELECT Max(ClosingDate) FROM FiscalYearClosings")
    If IsDate(closedEnd) Then
        If StartDay <= DateValue(closedEnd) Then
            FiscalPeriodProblem = " »œ√ «·› —… «·√Ê·Ï »⁄œ ¬Œ— ”‰… „ﬁ›·… (" & GDate(closedEnd) & ")."
            Exit Function
        End If
    End If
    If Day(StartDay) <> 1 Or Month(StartDay) <> 1 Then
        earliest = DbValue("SELECT Min(EntryDate) FROM JournalEntries WHERE SourceType <> 'YEAR_CLOSE'")
        If IsDate(earliest) Then
            If DateValue(earliest) < StartDay Then
                FiscalPeriodProblem = " ÊÃœ ﬁÌÊœ ﬁ»· " & GDate(StartDay) & " (√Ê·Â« " & GDate(earliest) & "). " & _
                    "«»œ√ «·› —… «·√Ê·Ï ›Ì 1 Ì‰«Ì—° › »ﬁÏ «·”‰Ê«  ﬁ»·Â« „Ì·«œÌ…° √Ê ﬁ»· √Ê· ﬁÌœ."
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
                      SqlDate(DateValue(EndDay)) & ", " & SqlText(Trim$(Notes)) & ")", dbFailOnError
    LogAction "FISCAL_PERIOD_ADD", "FiscalPeriods", Trim$(PeriodName), GDate(StartDay) & " - " & GDate(EndDay)
End Function

Public Function DeleteLastFiscalPeriod() As String
    ' Only the last period, and only while it is not closed.
    Dim id As Variant, endDay As Variant, before As Collection
    DeleteLastFiscalPeriod = CanFiscal("DELETE")
    If Len(DeleteLastFiscalPeriod) > 0 Then Exit Function
    id = DbValue("SELECT TOP 1 FiscalPeriodID FROM FiscalPeriods ORDER BY StartDate DESC")
    If IsNull(id) Then
        DeleteLastFiscalPeriod = "·«  ÊÃœ › —« ."
        Exit Function
    End If
    endDay = DbValue("SELECT EndDate FROM FiscalPeriods WHERE FiscalPeriodID = " & id)
    If YearIsClosed(Year(endDay)) Then
        DeleteLastFiscalPeriod = "«·› —… „ﬁ›·…. √⁄œ › ÕÂ« √Ê·« „‰ ‘«‘… ≈ﬁ›«· «·› —« ."
        Exit Function
    End If
    Set before = AuditSnapshot("FiscalPeriods", "FiscalPeriodID", id)
    CurrentDb.Execute "DELETE FROM FiscalPeriods WHERE FiscalPeriodID = " & CLng(id), dbFailOnError
    AuditDeleted "FISCAL_PERIOD_DELETE", "FiscalPeriods", id, before
End Function

'------------------------------------------------------------------------------
' Opening balances of a period
'------------------------------------------------------------------------------
Public Sub PrintOpeningBalances(ByVal StartDay As Date, Optional ByVal Title As String = "")
    ' Every balance at the start of StartDay: the accounts, the customers, the suppliers, the stock.
    Dim msg As String, criteria As String
    If Not HasPermission("JOURNAL") Then
        ShowWarning "«· ﬁ«—Ì— «·„Õ«”»Ì…  Õ «Ã ’·«ÕÌ… ´ﬁÌÊœ «·ÌÊ„Ì…ª."
        Exit Sub
    End If
    msg = SyncJournal()
    If Len(msg) > 0 Then ShowWarning msg
    SetPeriod StartDay, StartDay
    criteria = "›Ì »œ«Ì… ÌÊ„ " & GDate(StartDay)
    If Len(Title) > 0 Then criteria = Title & "    " & criteria
    OpenReportOrQuery "rptOpeningBalances", "OpeningBalancesReportQuery", "", criteria
End Sub

Public Sub PrintOpeningBalanceSheet(ByVal FiscalKey As Long)
    ' The balance sheet on the day before the period starts.
    Dim msg As String, criteria As String, asOf As Date
    If FiscalKey = 0 Then
        ShowWarning "«Œ — «·”‰… √Ê «·› —… «·„«·Ì…."
        Exit Sub
    End If
    If Not HasPermission("JOURNAL") Then
        ShowWarning "«· ﬁ«—Ì— «·„Õ«”»Ì…  Õ «Ã ’·«ÕÌ… ´ﬁÌÊœ «·ÌÊ„Ì…ª."
        Exit Sub
    End If
    msg = SyncJournal()
    If Len(msg) > 0 Then ShowWarning msg
    asOf = DateAdd("d", -1, FiscalYearStart(FiscalKey))
    criteria = PrepareFinancials("BALANCE_SHEET", asOf, asOf)
    OpenReportOrQuery "rptBalanceSheet", "BalanceSheetQuery", "", _
                      "«·„Ì“«‰Ì… «·«›  «ÕÌ… ·‹ " & FiscalKeyName(FiscalKey) & "    " & criteria
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
    frm!lstPeriods.RowSource = Tr("SELECT f.FiscalPeriodID, f.PeriodName AS [«·› —…], " & _
        "Format(f.StartDate, 'yyyy/mm/dd') AS [„‰], Format(f.EndDate, 'yyyy/mm/dd') AS [≈·Ï], " & _
        "IIf((SELECT COUNT(*) FROM FiscalYearClosings AS c WHERE c.FiscalYear = Year(f.EndDate)) > 0, " & _
        "'„ﬁ›·…', '„› ÊÕ…') AS [«·Õ«·…] FROM FiscalPeriods AS f ORDER BY f.StartDate")
    last = LastPeriodEnd()
    If IsDate(last) Then
        startDay = DateAdd("d", 1, DateValue(last))
        frm!txtStart.Locked = True
        frm!lblHint.Caption = Tr("«·› —… «·ÃœÌœ…  »œ√ ›Ì «·ÌÊ„ «· «·Ì ·‰Â«Ì… ¬Œ— › —…° Ê·«  ‰ ÂÌ › — «‰ ›Ì ”‰… „Ì·«œÌ… " & _
                                 "Ê«Õœ….  ıÕ–› ¬Œ— › —… ›ﬁÿ° ÊÂÌ €Ì— „ﬁ›·….")
    Else
        first = DbValue("SELECT Min(EntryDate) FROM JournalEntries WHERE SourceType <> 'YEAR_CLOSE'")
        If IsDate(first) Then
            startDay = DateSerial(Year(first), 1, 1)
        Else
            startDay = DateSerial(Year(Date), 1, 1)
        End If
        frm!txtStart.Locked = False
        frm!lblHint.Caption = Tr("«·› —… «·√Ê·Ï  »œ√ ›Ì 1 Ì‰«Ì— (Ê«·”‰Ê«  ﬁ»·Â« „Ì·«œÌ…) √Ê ﬁ»· √Ê· ﬁÌœ. " & _
                                 "»⁄œÂ«  ı÷«› «·› —«  »«· — Ì»° Ê·«  “Ìœ «·› —… ⁄·Ï 24 ‘Â—«.")
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
    ShowInfo " „  ≈÷«›… «·› —… " & frm!txtName.Value & "."
    FiscalPeriodsRefresh frm
End Sub

Public Sub FiscalDeleteLast(ByVal frm As Access.Form)
    Dim msg As String
    If Not AskYesNo("Õ–› ¬Œ— › —… „«·Ì…ø") Then Exit Sub
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
        ShowWarning "«Œ — › —… „‰ «·ﬁ«∆„…."
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
    Record FiscalPeriodProblem("", Date, Date) <> "", "«”„ «·› —… „ÿ·Ê»", passed, failed, report
    Record FiscalPeriodProblem("X", Date, DateAdd("d", -1, Date)) <> "", "«·‰Â«Ì… ·«  ﬂÊ‰ ﬁ»· «·»œ«Ì…", passed, failed, report
    Record FiscalPeriodProblem("X", Date, DateAdd("m", 25, Date)) <> "", "·«  “Ìœ ⁄·Ï 24 ‘Â—«", passed, failed, report
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
                          dbFailOnError                              ' a first period, so the rules of the next apply
    End If
    Record FiscalPeriodProblem("TEST-GAP", DateAdd("d", 1, firstDay), DateAdd("yyyy", 1, firstDay)) <> "", _
           "—›÷ «·›ÃÊ… »⁄œ ¬Œ— › —…", passed, failed, report
    Record FiscalPeriodProblem("TEST-OVERLAP", DateAdd("d", -1, firstDay), DateAdd("yyyy", 1, firstDay)) <> "", _
           "—›÷ «· œ«Œ· „⁄ ¬Œ— › —…", passed, failed, report
    s = firstDay
    e = DateAdd("d", -1, DateAdd("yyyy", 1, s))
    Record AddFiscalPeriod("TEST-FP", s, e) = "", "≈÷«›… › —… »⁄œ ¬Œ— › —…", passed, failed, report
    key = Year(e)
    Record FiscalYearStart(key) = s And FiscalYearEnd(key) = e And FiscalKeyName(key) = "TEST-FP", _
           "»œ«Ì… «·› —… Ê‰Â«Ì Â« Ê«”„Â« »„› «ÕÂ«", passed, failed, report
    Record FiscalKeyKnown(key), "«·› —… ›Ì ﬁ«∆„… «·≈ﬁ›«·", passed, failed, report
    Record FiscalPeriodOf(DateAdd("d", 10, s), s, e) And s = FiscalYearStart(key), "«·› —… «· Ì  Õ ÊÌ ÌÊ„«", _
           passed, failed, report
    Record YearCloseProblem(key) <> "" Or FiscalYearEnd(key) < Date, "·«  ıﬁ›· › —… ﬁ»· «‰ Â«∆Â«", passed, failed, report
    Record DeleteLastFiscalPeriod() = "" And Not FiscalKeyKnown(key), "Õ–› ¬Œ— › —…", passed, failed, report
    Record FiscalYearStart(2001) = DateSerial(2001, 1, 1) Or Not IsNull(DbValue("SELECT FiscalPeriodID FROM " & _
           "FiscalPeriods WHERE Year(EndDate) = 2001")), "«·”‰… «·„Ì·«œÌ… »·« › —…", passed, failed, report
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
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·› —«  «·„«·Ì… ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, _
                "TestFiscalPeriods"
        TestFiscalPeriods = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
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
