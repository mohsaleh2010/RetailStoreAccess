Attribute VB_Name = "modIndicators"
'==============================================================================
' modIndicators  -  Retail Store Management System
'
' Financial indicators of the dashboard (cards 9-12 of frmMain), from the journal so
' they agree with the financial statements (FinancialIndicatorsQuery):
'   gross margin        (net sales - cost of sales) / net sales, this month (and last month)
'   stock turnover      cost of sales of the last 365 days / average stock (1400)
'   collection period   customer debts (1300) / credit sales per day of the last 90 days
'   current ratio       current assets (group 11) / current liabilities (group 21)
' The journal is brought up to date first; the figures are kept 10 minutes (the
' Refresh button recomputes them at once). Python reference: tools/indicators_reference.py.
'==============================================================================
Option Compare Database
Option Explicit

Private Const CACHE_MINUTES As Long = 10
Private Const COLLECTION_DAYS_SPAN As Long = 90

Private m_values As Variant         ' Array of 4 x Array(value text, sub text, colour)
Private m_time As Date
Private m_asOf As Date

'------------------------------------------------------------------------------
' Formulas (Null = no meaning: nothing to divide by)
'------------------------------------------------------------------------------
Public Function GrossMargin(ByVal Sales As Currency, ByVal Cost As Currency) As Variant
    GrossMargin = Null
    If Sales > 0 Then GrossMargin = CDbl(Sales - Cost) / CDbl(Sales)     ' Double: no 4-decimal Currency
End Function

Public Function StockTurnover(ByVal YearCost As Currency, ByVal StockStart As Currency, ByVal StockEnd As Currency) As Variant
    Dim average As Double
    StockTurnover = Null
    average = (CDbl(StockStart) + CDbl(StockEnd)) / 2
    If average > 0 And YearCost > 0 Then StockTurnover = CDbl(YearCost) / average
End Function

Public Function CollectionDays(ByVal Receivables As Currency, ByVal CreditSales As Currency, _
                               Optional ByVal Days As Long = COLLECTION_DAYS_SPAN) As Variant
    CollectionDays = Null
    If Receivables < 0 Then Receivables = 0
    If CreditSales > 0 Then CollectionDays = CDbl(Receivables) / (CDbl(CreditSales) / Days)
End Function

Public Function CurrentRatio(ByVal CurrentAssets As Currency, ByVal CurrentLiabilities As Currency) As Variant
    CurrentRatio = Null
    If CurrentLiabilities > 0 Then CurrentRatio = CDbl(CurrentAssets) / CDbl(CurrentLiabilities)
End Function

'------------------------------------------------------------------------------
' The four cards
'------------------------------------------------------------------------------
Public Sub SetIndicatorParams(ByVal AsOf As Date)
    AsOf = DateValue(AsOf)
    TempVars.Add "IndEnd", AsOf + 1
    TempVars.Add "IndMonth", DateSerial(Year(AsOf), Month(AsOf), 1)
    TempVars.Add "IndPrevMonth", DateSerial(Year(AsOf), Month(AsOf) - 1, 1)
    TempVars.Add "IndYear", AsOf - 364
    TempVars.Add "Ind90", AsOf - (COLLECTION_DAYS_SPAN - 1)
End Sub

Public Function FinancialIndicators(ByVal AsOf As Date, Optional ByVal Force As Boolean = False) As Variant
    ' Array of 4 x Array(value text, sub text, colour): margin, turnover, collection, liquidity.
    Dim rs As DAO.Recordset, v As Variant, prev As Variant, out(3) As Variant, days As Variant
    If Not Force And IsArray(m_values) Then
        If m_asOf = DateValue(AsOf) And DateDiff("n", m_time, Now) < CACHE_MINUTES Then
            FinancialIndicators = m_values
            Exit Function
        End If
    End If
    SyncJournal                                       ' the documents saved since the last time
    SetIndicatorParams AsOf
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM FinancialIndicatorsQuery", dbOpenSnapshot)

    v = GrossMargin(rs!MonthSales, rs!MonthCost)
    prev = GrossMargin(rs!PrevSales, rs!PrevCost)
    If IsNull(v) Then
        out(0) = Array("-", "·« „»Ì⁄«  Â–« «·‘Â—", CLR_MUTED)
    Else
        out(0) = Array(Format$(v, "0.0%"), "«·‘Â— «·„«÷Ì: " & PercentText(prev), IIf(v < 0, CLR_DANGER, CLR_PRIMARY))
    End If

    v = StockTurnover(rs!YearCost, rs!StockStart, rs!StockEnd)
    If IsNull(v) Then
        out(1) = Array("-", "·«  ﬂ·›… „»Ì⁄«  √Ê ·« „Œ“Ê‰", CLR_MUTED)
    Else
        days = 365 / v
        out(1) = Array(Format$(v, "0.0") & " „—…/”‰…", "«·„Œ“Ê‰ Ìﬂ›Ì ‰ÕÊ " & Format$(days, "0") & " ÌÊ„«", CLR_PRIMARY)
    End If

    v = CollectionDays(rs!Receivables, rs!CreditSales)
    If IsNull(v) Then
        out(2) = Array("-", "·« „»Ì⁄«  ¬Ã·… ›Ì ¬Œ— " & COLLECTION_DAYS_SPAN & " ÌÊ„«", CLR_MUTED)
    Else
        out(2) = Array(Format$(v, "0") & " ÌÊ„«", "œÌÊ‰ «·⁄„·«¡ ˜ «·„»Ì⁄«  «·¬Ã·… «·ÌÊ„Ì…", _
                       IIf(v > 60, CLR_DANGER, IIf(v > 30, CLR_WARNING, CLR_PRIMARY)))
    End If

    v = CurrentRatio(rs!CurrentAssets, rs!CurrentLiabilities)
    If IsNull(v) Then
        out(3) = Array("-", "·« Œ’Ê„ „ œ«Ê·…", CLR_MUTED)
    Else
        out(3) = Array(Format$(v, "0.00"), "«·√’Ê· «·„ œ«Ê·… ˜ «·Œ’Ê„ «·„ œ«Ê·…", _
                       IIf(v < 1, CLR_DANGER, IIf(v < 1.5, CLR_WARNING, CLR_SUCCESS)))
    End If
    rs.Close
    m_values = out
    m_time = Now
    m_asOf = DateValue(AsOf)
    FinancialIndicators = m_values
End Function

Private Function PercentText(ByVal v As Variant) As String
    If IsNull(v) Then
        PercentText = "-"
    Else
        PercentText = Format$(v, "0.0%")
    End If
End Function

Public Sub ForgetIndicators()
    ' The next dashboard refresh recomputes them (after a closing, a restore...).
    m_values = Empty
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Public Function TestIndicators() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, passed As Long, failed As Long, report As String, msg As String
    Dim before As DAO.Recordset, after As DAO.Recordset, id As Long, a0 As Currency, l0 As Currency, s0 As Currency
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestIndicators  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Check GrossMargin(1000, 750) = 0.25 And IsNull(GrossMargin(0, 10)), "Â«„‘ «·—»Õ", passed, failed, report
    Check StockTurnover(1200, 200, 400) = 4 And IsNull(StockTurnover(100, 0, 0)), "œÊ—«‰ «·„Œ“Ê‰", passed, failed, report
    Check CollectionDays(3000, 9000) = 30 And IsNull(CollectionDays(100, 0)) And CollectionDays(-5, 900) = 0, _
          "› —… «· Õ’Ì·", passed, failed, report
    Check CurrentRatio(3000, 2000) = 1.5 And IsNull(CurrentRatio(100, 0)), "‰”»… «·”ÌÊ·…", passed, failed, report
    msg = SyncJournal()
    SetIndicatorParams Date
    Set before = CurrentDb.OpenRecordset("SELECT * FROM FinancialIndicatorsQuery", dbOpenSnapshot)
    a0 = before!CurrentAssets
    l0 = before!CurrentLiabilities
    s0 = before!StockEnd
    before.Close
    Check a0 = Nz(DbValue("SELECT Sum(l.Debit) - Sum(l.Credit) FROM JournalLines AS l INNER JOIN Accounts AS a ON " & _
          "l.AccountCode = a.AccountCode WHERE a.Level2Code = 11"), 0), "«·√’Ê· «·„ œ«Ê·… = √—’œ… «·„Ã„Ê⁄… 11 " & msg, _
          passed, failed, report
    If ClosedThroughDate() >= Date Then
        Debug.Print "[--] «·› —… „ﬁ›·… Õ Ï «·ÌÊ„"
        GoTo Done
    End If
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    EnsureLocalTables
    CurrentDb.Execute "DELETE FROM tmpManualLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpManualLines (AccountCode, Debit, Credit) VALUES (1400, 1000, 0)", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpManualLines (AccountCode, Debit, Credit) VALUES (2100, 0, 1000)", dbFailOnError
    msg = PostManualEntry(0, Date, "TEST-IND", "", Null, id)
    msg = msg & SyncJournal()
    Set after = CurrentDb.OpenRecordset("SELECT * FROM FinancialIndicatorsQuery", dbOpenSnapshot)
    Check Len(msg) = 0 And after!StockEnd = s0 + 1000 And after!CurrentAssets = a0 + 1000 And _
          after!CurrentLiabilities = l0 + 1000, "ﬁÌœ „Œ“Ê‰ ⁄·Ï „Ê—œ Ì—›⁄ «·„Œ“Ê‰ Ê«·√’Ê· Ê«·Œ’Ê„ «·„ œ«Ê·… " & msg, _
          passed, failed, report
    after.Close
    CurrentDb.Execute "DELETE FROM tmpManualLines", dbFailOnError
    ws.Rollback
    inTrans = False
    ForgetIndicators
    GoTo Done
EH:
    Check False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·„ƒ‘—«  «·„«·Ì… ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestIndicators"
        TestIndicators = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestIndicators"
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
