Attribute VB_Name = "modClosing"
'==============================================================================
' modClosing  -  Retail Store Management System
'
' Period and fiscal year closing (screen frmPeriodClosing, permission PERIOD_CLOSE):
'   Settings.ClosedThrough   the last closed day: nothing dated until it may be
'                            added, changed or deleted (expenses, manual entries,
'                            opening balances...). SyncJournal never changes an
'                            entry dated in the closed period either.
'   ClosePeriod / ReopenPeriod     move that day forward / back (a reason is kept
'                            in PeriodClosings).
'   CloseFiscalYear          the year closing entry on 31 December: every revenue and
'                            expense account back to zero, the net profit (loss) to
'                            retained earnings 3300 (FiscalYearClosings + lines,
'                            journal source YEAR_CLOSE); the year is closed through
'                            31 December. Years are closed in order.
'   ReopenFiscalYear         removes the last year closing and opens that year.
' The income statement leaves the year closing entries out (qryIncomeMoves).
' Same rules as tools/sim.py (Store.close_year).
'==============================================================================
Option Compare Database
Option Explicit

Public Const RETAINED_EARNINGS As Long = 3300

'------------------------------------------------------------------------------
' The closed period
'------------------------------------------------------------------------------
Public Function ClosedThroughDate() As Date
    ' The last closed day, 0 when nothing is closed.
    Dim v As Variant
    v = SettingValue("ClosedThrough")
    If IsDate(v) Then ClosedThroughDate = DateValue(v)
End Function

Public Function ClosedPeriodProblem(ByVal When As Variant) As String
    ' "" when the date is open (or empty).
    Dim closed As Date
    If Not IsDate(When) Then Exit Function
    closed = ClosedThroughDate()
    If closed > 0 And DateValue(When) <= closed Then
        ClosedPeriodProblem = "التاريخ " & GDate(When) & " في فترة مقفلة (الإقفال حتى " & GDate(closed) & ")." & vbCrLf & _
                              "لا يُضاف ولا يُعدَّل ولا يُحذف فيها شيء إلا بعد إعادة فتح الفترة."
    End If
End Function

Public Function ClosedRecordProblem(ByVal frm As Access.Form, Optional ByVal Deleting As Boolean = False) As String
    ' Data screens: a record that is (or would become) part of a closed period.
    Dim p As String
    Select Case TagValue(frm, "TABLE")
        Case "Expenses"
            If Not frm.NewRecord Then p = ClosedPeriodProblem(frm!ExpenseDate.OldValue)
            If Len(p) = 0 And Not Deleting Then p = ClosedPeriodProblem(frm!ExpenseDate.Value)
        Case "Customers", "Suppliers"                    ' the opening balance is an entry on CreatedAt
            If Not frm.NewRecord Then
                If Deleting Or Nz(frm!OpeningBalance.Value, 0) <> Nz(frm!OpeningBalance.OldValue, 0) Then
                    p = ClosedPeriodProblem(frm!CreatedAt.Value)
                End If
            End If
        Case "CashBoxes", "Banks"
            If Not frm.NewRecord Then
                If Deleting Or Nz(frm!OpeningBalance.Value, 0) <> Nz(frm!OpeningBalance.OldValue, 0) Or _
                   Nz(frm!OpeningDate.Value, 0) <> Nz(frm!OpeningDate.OldValue, 0) Then
                    p = ClosedPeriodProblem(frm!OpeningDate.OldValue)
                End If
            End If
            If Len(p) = 0 And Not Deleting And Nz(frm!OpeningBalance.Value, 0) <> 0 Then
                p = ClosedPeriodProblem(frm!OpeningDate.Value)
            End If
    End Select
    ClosedRecordProblem = p
End Function

Private Function CanClose() As String
    If Not HasPermission("PERIOD_CLOSE") Then CanClose = "لا تملك صلاحية إقفال الفترات."
End Function

Private Function JournalReady() As String
    ' Every operation must have its entry before a period closes: later it could not get one.
    Dim msg As String
    msg = SyncJournal()
    If Len(msg) > 0 Then JournalReady = "أصلح هذا أولًا، ثم أقفل:" & vbCrLf & msg
End Function

Private Sub LogClosing(ByVal ActionType As String, ByVal Through As Variant, ByVal Previous As Date, _
                       ByVal FiscalYear As Variant, ByVal Notes As String)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("PeriodClosings", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!ActionType = ActionType
    rs!ClosedThrough = Through
    If Previous > 0 Then rs!PreviousThrough = Previous
    rs!FiscalYear = FiscalYear
    If Len(Trim$(Notes)) > 0 Then rs!Notes = Left$(Trim$(Notes), 255)
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Close
    LogAction "PERIOD_" & ActionType, "Settings", IIf(IsDate(Through), GDate(Through), ""), Notes
End Sub

Private Sub SetClosedThrough(ByVal Through As Variant)
    CurrentDb.Execute "UPDATE Settings SET ClosedThrough = " & IIf(IsDate(Through), SqlDate(Through), "Null") & _
                      " WHERE SettingID = 1", dbFailOnError
End Sub

Public Function ClosePeriod(ByVal Through As Date, ByVal Notes As String) As String
    Dim previous As Date
    ClosePeriod = CanClose()
    If Len(ClosePeriod) > 0 Then Exit Function
    Through = DateValue(Through)
    previous = ClosedThroughDate()
    If Through >= Date Then
        ClosePeriod = "يُقفل ما قبل اليوم فقط: اختر تاريخًا قبل " & GDate(Date) & "."
    ElseIf Through <= previous Then
        ClosePeriod = "الفترة مقفلة حتى " & GDate(previous) & " بالفعل. لإرجاعها استخدم «إعادة الفتح»."
    Else
        ClosePeriod = JournalReady()
    End If
    If Len(ClosePeriod) > 0 Then Exit Function
    SetClosedThrough Through
    LogClosing "CLOSE", Through, previous, Null, Notes
End Function

Public Function ReopenPeriod(ByVal Through As Variant, ByVal Notes As String) As String
    ' Through: the new last closed day (earlier than now), or Null to open everything.
    Dim previous As Date, lastYearEnd As Variant
    ReopenPeriod = CanClose()
    If Len(ReopenPeriod) > 0 Then Exit Function
    previous = ClosedThroughDate()
    lastYearEnd = DbValue("SELECT Max(ClosingDate) FROM FiscalYearClosings")
    If previous = 0 Then
        ReopenPeriod = "لا توجد فترة مقفلة."
    ElseIf IsDate(Through) And DateValue(Nz(Through, 0)) >= previous Then
        ReopenPeriod = "اختر تاريخًا قبل " & GDate(previous) & "، أو اترك التاريخ فارغًا لفتح كل الفترات."
    ElseIf IsDate(lastYearEnd) And (Not IsDate(Through) Or Nz(Through, 0) < Nz(lastYearEnd, 0)) Then
        ReopenPeriod = "السنة " & Year(lastYearEnd) & " مقفلة بقيد إقفال. أعد فتح السنة أولًا."
    ElseIf Len(Trim$(Notes)) = 0 Then
        ReopenPeriod = "اكتب سبب إعادة الفتح."
    End If
    If Len(ReopenPeriod) > 0 Then Exit Function
    If IsDate(Through) Then Through = DateValue(Through)
    SetClosedThrough Through
    LogClosing "REOPEN", Through, previous, Null, Notes
End Function

'------------------------------------------------------------------------------
' Fiscal year (calendar year)
'------------------------------------------------------------------------------
Public Function FirstEntryYear() As Long
    FirstEntryYear = Year(Nz(DbValue("SELECT Min(EntryDate) FROM JournalEntries WHERE SourceType <> 'YEAR_CLOSE'"), Date))
End Function

Public Function YearIsClosed(ByVal FiscalYear As Long) As Boolean
    YearIsClosed = Not IsNull(DbValue("SELECT YearClosingID FROM FiscalYearClosings WHERE FiscalYear = " & FiscalYear))
End Function

Public Function YearNetProfit(ByVal FiscalYear As Long) As Currency
    ' Revenue - expenses still open until 31 December of the year (credit positive).
    YearNetProfit = Nz(DbValue("SELECT Sum(l.Credit) - Sum(l.Debit) FROM (JournalLines AS l INNER JOIN JournalEntries AS e " & _
        "ON l.EntryID = e.EntryID) INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode WHERE e.EntryDate < " & _
        SqlDate(DateSerial(FiscalYear + 1, 1, 1)) & " AND a.AccountType IN ('REVENUE', 'EXPENSE')"), 0)
End Function

Public Function YearCloseProblem(ByVal FiscalYear As Long) As String
    Dim y As Long
    YearCloseProblem = CanClose()
    If Len(YearCloseProblem) > 0 Then Exit Function
    If YearIsClosed(FiscalYear) Then
        YearCloseProblem = "السنة " & FiscalYear & " مقفلة بالفعل."
    ElseIf FiscalYear >= Year(Date) Then
        YearCloseProblem = "تُقفل السنة بعد انتهائها (بعد 31 ديسمبر " & FiscalYear & ")."
    Else
        For y = FirstEntryYear() To FiscalYear - 1
            If Not YearIsClosed(y) And Nz(DbValue("SELECT COUNT(*) FROM JournalEntries WHERE EntryDate >= " & _
                    SqlDate(DateSerial(y, 1, 1)) & " AND EntryDate < " & SqlDate(DateSerial(y + 1, 1, 1))), 0) > 0 Then
                YearCloseProblem = "أقفل السنة " & y & " أولًا: السنوات تُقفل بالترتيب."
                Exit Function
            End If
        Next
    End If
End Function

Public Function CloseFiscalYear(ByVal FiscalYear As Long, ByVal Notes As String) As String
    Dim ws As DAO.Workspace, db As DAO.Database, rs As DAO.Recordset, h As DAO.Recordset, l As DAO.Recordset
    Dim inTrans As Boolean, id As Long, n As Long, net As Currency, profit As Currency, yearEnd As Date
    Dim previous As Date
    On Error GoTo EH
    CloseFiscalYear = YearCloseProblem(FiscalYear)
    If Len(CloseFiscalYear) = 0 Then CloseFiscalYear = JournalReady()
    If Len(CloseFiscalYear) > 0 Then Exit Function
    yearEnd = DateSerial(FiscalYear, 12, 31)
    previous = ClosedThroughDate()
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    Set h = db.OpenRecordset("FiscalYearClosings", dbOpenDynaset)
    h.AddNew
    h!FiscalYear = FiscalYear
    h!ClosingNumber = "FY-" & FiscalYear
    h!ClosingDate = yearEnd
    h!NetProfit = 0
    h!EmployeeID = CurrentUserID()
    h!Notes = Left$("إقفال السنة " & FiscalYear & IIf(Len(Trim$(Notes)) > 0, " - " & Trim$(Notes), ""), 255)
    h.Update
    h.Bookmark = h.LastModified
    id = h!YearClosingID
    ' every revenue / expense account with a balance on 31 December, reversed
    Set rs = db.OpenRecordset("SELECT l.AccountCode, Sum(l.Debit) - Sum(l.Credit) AS Net FROM (JournalLines AS l " & _
        "INNER JOIN JournalEntries AS e ON l.EntryID = e.EntryID) INNER JOIN Accounts AS a ON l.AccountCode = " & _
        "a.AccountCode WHERE e.EntryDate < " & SqlDate(DateSerial(FiscalYear + 1, 1, 1)) & " AND a.AccountType IN " & _
        "('REVENUE', 'EXPENSE') GROUP BY l.AccountCode HAVING Sum(l.Debit) <> Sum(l.Credit) ORDER BY l.AccountCode", _
        dbOpenSnapshot)
    Set l = db.OpenRecordset("FiscalYearClosingLines", dbOpenDynaset, dbAppendOnly)
    Do Until rs.EOF
        net = rs!Net
        n = n + 1
        AddClosingLine l, id, n, rs!AccountCode, IIf(net < 0, -net, 0), IIf(net > 0, net, 0), "إقفال السنة " & FiscalYear
        profit = profit - net
        rs.MoveNext
    Loop
    rs.Close
    If n > 0 And profit <> 0 Then
        n = n + 1
        AddClosingLine l, id, n, RETAINED_EARNINGS, IIf(profit < 0, -profit, 0), IIf(profit > 0, profit, 0), _
                       IIf(profit >= 0, "صافي ربح السنة ", "صافي خسارة السنة ") & FiscalYear
    End If
    l.Close
    h.Edit
    h!NetProfit = profit
    h.Update
    h.Close
    If yearEnd > previous Then SetClosedThrough yearEnd
    LogClosing "YEAR_CLOSE", IIf(yearEnd > previous, yearEnd, previous), previous, FiscalYear, Notes
    ws.CommitTrans
    inTrans = False
    CloseFiscalYear = SyncJournal()                    ' the closing entry (YEAR_CLOSE)
    Exit Function
EH:
    CloseFiscalYear = "تعذر إقفال السنة: " & Err.Description
    If inTrans Then ws.Rollback
End Function

Private Sub AddClosingLine(ByVal l As DAO.Recordset, ByVal ClosingID As Long, ByVal LineNumber As Long, _
                           ByVal AccountCode As Long, ByVal Debit As Currency, ByVal Credit As Currency, _
                           ByVal LineText As String)
    l.AddNew
    l!YearClosingID = ClosingID
    l!LineNumber = LineNumber
    l!AccountCode = AccountCode
    l!Debit = Debit
    l!Credit = Credit
    l!LineText = Left$(LineText, 150)
    l.Update
End Sub

Public Function ReopenFiscalYear(ByVal FiscalYear As Long, ByVal Notes As String) As String
    Dim previous As Date, opened As Date
    ReopenFiscalYear = CanClose()
    If Len(ReopenFiscalYear) > 0 Then Exit Function
    If Not YearIsClosed(FiscalYear) Then
        ReopenFiscalYear = "السنة " & FiscalYear & " غير مقفلة."
    ElseIf Nz(DbValue("SELECT Max(FiscalYear) FROM FiscalYearClosings"), 0) <> FiscalYear Then
        ReopenFiscalYear = "تُعاد فتح آخر سنة مقفلة فقط (" & DbValue("SELECT Max(FiscalYear) FROM FiscalYearClosings") & ")."
    ElseIf Len(Trim$(Notes)) = 0 Then
        ReopenFiscalYear = "اكتب سبب إعادة فتح السنة."
    End If
    If Len(ReopenFiscalYear) > 0 Then Exit Function
    previous = ClosedThroughDate()
    opened = DateSerial(FiscalYear, 1, 0)              ' 31 December of the year before
    CurrentDb.Execute "DELETE FROM FiscalYearClosings WHERE FiscalYear = " & FiscalYear, dbFailOnError   ' lines cascade
    If previous > opened Then SetClosedThrough opened
    LogClosing "YEAR_OPEN", IIf(previous > opened, opened, previous), previous, FiscalYear, Notes
    ReopenFiscalYear = SyncJournal()                   ' removes the closing entry
End Function

'------------------------------------------------------------------------------
' Screen frmPeriodClosing
'------------------------------------------------------------------------------
Public Sub PeriodClosingLoad(ByVal frm As Access.Form)
    Dim rows As String, y As Long
    Calendar = vbCalGreg
    For y = Year(Date) - 1 To FirstEntryYear() Step -1
        rows = rows & IIf(Len(rows) > 0, ";", "") & y
    Next
    If Len(rows) = 0 Then rows = CStr(Year(Date) - 1)
    frm!cboYear.RowSourceType = "Value List"
    frm!cboYear.RowSource = rows
    frm!cboYear.Value = CLng(Split(rows, ";")(0))
    frm!txtThrough.Value = DateSerial(Year(Date), Month(Date), 0)          ' end of last month
    PeriodClosingRefresh frm
End Sub

Public Sub PeriodClosingRefresh(ByVal frm As Access.Form)
    Dim closed As Date, y As Long
    closed = ClosedThroughDate()
    If closed > 0 Then
        frm!lblState.Caption = "الفترة مقفلة حتى " & GDate(closed)
        frm!lblState.ForeColor = CLR_WARNING
    Else
        frm!lblState.Caption = "لا توجد فترة مقفلة"
        frm!lblState.ForeColor = CLR_SUCCESS
    End If
    y = Nz(frm!cboYear.Value, Year(Date) - 1)
    If YearIsClosed(y) Then
        frm!lblYearInfo.Caption = "السنة " & y & " مقفلة: صافي الربح المرحَّل للأرباح المحتجزة " & _
            Format$(Nz(DbValue("SELECT NetProfit FROM FiscalYearClosings WHERE FiscalYear = " & y), 0), "#,##0.00")
    Else
        frm!lblYearInfo.Caption = "السنة " & y & " غير مقفلة: صافي الربح (الخسارة) حتى 31 ديسمبر " & _
                                  Format$(YearNetProfit(y), "#,##0.00")
    End If
    frm!lstHistory.Requery
End Sub

Private Function Done(ByVal frm As Access.Form, ByVal Msg As String, ByVal Success As String) As Boolean
    If Len(Msg) > 0 Then
        ShowWarning Msg
    Else
        frm!txtNotes.Value = Null
        ShowInfo Success
        Done = True
    End If
    PeriodClosingRefresh frm
End Function

Public Sub DoClosePeriod(ByVal frm As Access.Form)
    If Not IsDate(frm!txtThrough.Value) Then
        ShowWarning "اختر آخر يوم في الفترة التي تُقفل."
        Exit Sub
    End If
    If Not AskYesNo("إقفال كل المستندات حتى " & GDate(frm!txtThrough.Value) & "؟" & vbCrLf & _
                    "بعدها لا يُضاف ولا يُعدَّل ولا يُحذف أي مستند بتاريخ حتى هذا اليوم.") Then Exit Sub
    Done frm, ClosePeriod(frm!txtThrough.Value, Nz(frm!txtNotes.Value, "")), "تم إقفال الفترة حتى " & GDate(frm!txtThrough.Value) & "."
End Sub

Public Sub DoReopenPeriod(ByVal frm As Access.Form)
    Dim target As String
    target = IIf(IsDate(frm!txtThrough.Value), "حتى " & GDate(Nz(frm!txtThrough.Value, 0)), "بالكامل (لا تبقى فترة مقفلة)")
    If Not AskYesNo("إعادة الفتح: تصبح الفترة مقفلة " & target & "؟") Then Exit Sub
    Done frm, ReopenPeriod(frm!txtThrough.Value, Nz(frm!txtNotes.Value, "")), "تمت إعادة فتح الفترة."
End Sub

Public Sub DoCloseYear(ByVal frm As Access.Form)
    Dim y As Long
    y = Nz(frm!cboYear.Value, 0)
    If y = 0 Then Exit Sub
    If Not AskYesNo("إقفال السنة " & y & "؟" & vbCrLf & "يُقيَّد صافي الربح (الخسارة) " & Format$(YearNetProfit(y), "#,##0.00") & _
                    " في الأرباح المحتجزة، وتُقفل السنة كلها.") Then Exit Sub
    Done frm, CloseFiscalYear(y, Nz(frm!txtNotes.Value, "")), "تم إقفال السنة " & y & "."
End Sub

Public Sub DoReopenYear(ByVal frm As Access.Form)
    Dim y As Long
    y = Nz(frm!cboYear.Value, 0)
    If y = 0 Then Exit Sub
    If Not AskYesNo("إعادة فتح السنة " & y & "؟ يُحذف قيد إقفالها وتُفتح فترتها.") Then Exit Sub
    Done frm, ReopenFiscalYear(y, Nz(frm!txtNotes.Value, "")), "تمت إعادة فتح السنة " & y & "."
End Sub
