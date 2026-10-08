Attribute VB_Name = "modRecurring"
'==============================================================================
' modRecurring  -  Retail Store Management System
'
' Recurring expenses (RecurringExpenses, screen frmRecurring): rent, electricity,
' subscriptions. Each one has an amount, a frequency (monthly, every 3 months,
' yearly) and a due day (1-28). CreateDueExpenses writes an ordinary expense
' (Expenses.RecurringID) for every due date up to a day, once per date, and
' moves NextDueDate on; the expense then gets its journal entry like any other.
' On the first dashboard of a session the user with the EXPENSES permission is
' told how many are due and may create them at once (RecurringAlert).
' A due date in a closed period is not created: that recurring expense stops
' there until the period is reopened (or its next date is moved on the screen).
'==============================================================================
Option Compare Database
Option Explicit

Private Const MAX_PER_RUN As Long = 36       ' one run never creates more than 3 years of one expense

'------------------------------------------------------------------------------
' Schedule
'------------------------------------------------------------------------------
Public Function MonthsOf(ByVal Frequency As String) As Long
    Select Case Frequency
        Case "QUARTERLY": MonthsOf = 3
        Case "YEARLY": MonthsOf = 12
        Case Else: MonthsOf = 1
    End Select
End Function

Public Function FirstDueDate(ByVal StartDate As Date, ByVal DueDay As Long) As Date
    ' The first due day on or after the start date.
    Dim d As Date
    d = DateSerial(Year(StartDate), Month(StartDate), DueDay)
    If d < Int(StartDate) Then d = DateSerial(Year(StartDate), Month(StartDate) + 1, DueDay)
    FirstDueDate = d
End Function

Public Function DueAfter(ByVal Due As Date, ByVal Frequency As String, ByVal DueDay As Long) As Date
    DueAfter = DateSerial(Year(Due), Month(Due) + MonthsOf(Frequency), DueDay)
End Function

Public Function DuePeriodText(ByVal Due As Date, ByVal Frequency As String) As String
    ' "2026/10" for a month, "2026/10 - 2026/12" for a quarter, "2026" for a year.
    Dim saved As Integer
    saved = Calendar
    Calendar = vbCalGreg
    Select Case Frequency
        Case "QUARTERLY"
            DuePeriodText = Format$(Due, "yyyy/mm") & " - " & Format$(DateSerial(Year(Due), Month(Due) + 2, 1), "yyyy/mm")
        Case "YEARLY"
            DuePeriodText = Format$(Due, "yyyy")
        Case Else
            DuePeriodText = Format$(Due, "yyyy/mm")
    End Select
    Calendar = saved
End Function

'------------------------------------------------------------------------------
' Creating the expenses that are due
'------------------------------------------------------------------------------
Public Function DueRecurringCount(Optional ByVal AsOf As Variant) As Long
    If IsMissing(AsOf) Then AsOf = Date
    DueRecurringCount = Nz(DbValue("SELECT COUNT(*) FROM RecurringExpenses WHERE IsActive = True AND " & _
        "NextDueDate <= " & SqlDate(AsOf) & " AND (EndDate Is Null OR NextDueDate <= EndDate)"), 0)
End Function

Public Function CreateDueExpenses(ByVal AsOf As Date, ByRef Created As Long, Optional ByVal OnlyID As Long = 0) As String
    ' Every active recurring expense (or only OnlyID): one expense per due date up to AsOf.
    ' Returns "" or the reasons some could not be created (one line each).
    Dim rs As DAO.Recordset, problems As String, msg As String, n As Long
    Created = 0
    If Not HasPermission("EXPENSES") Then
        CreateDueExpenses = "ليست لديك صلاحية المصروفات."
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM RecurringExpenses WHERE IsActive = True AND NextDueDate <= " & _
        SqlDate(AsOf) & IIf(OnlyID > 0, " AND RecurringID = " & OnlyID, "") & " ORDER BY NextDueDate, RecurringID", _
        dbOpenDynaset)
    Do Until rs.EOF
        msg = CreateDueOf(rs, AsOf, n)
        Created = Created + n
        If Len(msg) > 0 Then problems = problems & rs!RecurringName & ": " & msg & vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    CreateDueExpenses = problems
End Function

Private Function CreateDueOf(ByVal rs As DAO.Recordset, ByVal AsOf As Date, ByRef n As Long) As String
    Dim due As Date, closedMsg As String
    n = 0
    due = rs!NextDueDate
    Do While due <= AsOf And n < MAX_PER_RUN
        If Not IsNull(rs!EndDate) Then
            If due > rs!EndDate Then Exit Do
        End If
        closedMsg = ClosedPeriodProblem(due)
        If Len(closedMsg) > 0 Then
            CreateDueOf = "استحقاق " & GDate(due) & " في فترة مقفلة، لم يُنشأ."
            Exit Do
        End If
        ' never twice for one date (a second user, or a run that stopped half way)
        If IsNull(DbValue("SELECT ExpenseID FROM Expenses WHERE RecurringID = " & rs!RecurringID & _
                          " AND ExpenseDate = " & SqlDate(due))) Then
            AddRecurringExpense rs, due
            n = n + 1
        End If
        rs.Edit
        rs!LastCreatedDate = due
        due = DueAfter(due, rs!Frequency, rs!DueDay)
        rs!NextDueDate = due
        rs.Update
    Loop
    If n > 0 Then LogAction "RECURRING_CREATE", "RecurringExpenses", CStr(rs!RecurringID), n & " expense(s)"
End Function

Private Sub AddRecurringExpense(ByVal r As DAO.Recordset, ByVal Due As Date)
    ' An ordinary expense, as the expense screen saves it (ValidateExpense).
    Dim e As DAO.Recordset, number As String, total As Currency, tries As Long
    total = r!Amount + r!Tax
    Do
        number = NextNumber("EXPENSE")
        tries = tries + 1
    Loop While Nz(DbValue("SELECT COUNT(*) FROM Expenses WHERE ExpenseNumber = " & SqlText(number)), 0) > 0 And tries < 1000
    Set e = CurrentDb.OpenRecordset("Expenses", dbOpenDynaset, dbAppendOnly)
    e.AddNew
    e!ExpenseNumber = number
    e!ExpenseDate = Due
    e!ExpenseTypeID = r!ExpenseTypeID
    e!Amount = r!Amount
    e!Tax = r!Tax
    e!TotalAmount = total
    e!PaymentMethodID = r!PaymentMethodID
    If r!PaymentMethodID = CASH_METHOD_ID Then
        If IsNull(r!CashBoxID) Then
            e!CashBoxID = CurrentCashBoxID()
        Else
            e!CashBoxID = r!CashBoxID
        End If
    End If
    If r!PaymentMethodID = BANK_TRANSFER_METHOD_ID Then
        If IsNull(r!BankID) Then
            e!BankID = BankFor(r!PaymentMethodID, total)
        Else
            e!BankID = r!BankID
        End If
    End If
    If IsNull(r!CostCenterID) Then
        e!CostCenterID = CostCenterFor()
    Else
        e!CostCenterID = r!CostCenterID
    End If
    e!Description = Left$(r!RecurringName & " - " & DuePeriodText(Due, r!Frequency) & _
                          IIf(Len(Nz(r!Description, "")) > 0, " - " & r!Description, ""), 255)
    e!EmployeeID = CurrentUserID()
    e!RecurringID = r!RecurringID
    e.Update
    e.Close
End Sub

Public Sub RecurringAlert()
    ' First dashboard of a session: offer to create what is due.
    Dim n As Long, created As Long, msg As String
    On Error GoTo Done
    If Not IsNull(TempVars("RecurringAlertShown")) Then Exit Sub
    TempVars.Add "RecurringAlertShown", True
    If Not HasPermission("EXPENSES") Then Exit Sub
    n = DueRecurringCount()
    If n = 0 Then Exit Sub
    If Not AskYesNo("يوجد " & n & " مصروف متكرر مستحق (إيجار، كهرباء...)." & vbCrLf & _
                    "هل تريد إنشاء المصروفات المستحقة الآن؟") Then Exit Sub
    msg = CreateDueExpenses(Date, created)
    ShowRecurringResult created, msg
Done:
End Sub

Private Sub ShowRecurringResult(ByVal Created As Long, ByVal Problems As String)
    If Len(Problems) = 0 Then
        ShowInfo "تم إنشاء " & Created & " مصروف من المصروفات المتكررة."
    Else
        ShowWarning "تم إنشاء " & Created & " مصروف." & vbCrLf & vbCrLf & "لم يُنشأ:" & vbCrLf & Problems
    End If
End Sub

'------------------------------------------------------------------------------
' frmRecurring (data screen: modForms calls these for the RecurringExpenses table)
'------------------------------------------------------------------------------
Public Function ValidateRecurring(ByVal frm As Access.Form) As Boolean
    Dim scheduleChanged As Boolean
    If Nz(frm!Amount.Value, 0) <= 0 Then
        ShowWarning "اكتب المبلغ."
        SafeFocus frm!Amount
        Exit Function
    End If
    If Not IsNull(frm!EndDate.Value) Then
        If frm!EndDate.Value < frm!StartDate.Value Then
            ShowWarning "تاريخ النهاية قبل تاريخ البداية."
            SafeFocus frm!EndDate
            Exit Function
        End If
    End If
    If Nz(frm!PaymentMethodID.Value, 0) <> CASH_METHOD_ID Then frm!CashBoxID.Value = Null
    If Nz(frm!PaymentMethodID.Value, 0) <> BANK_TRANSFER_METHOD_ID Then frm!BankID.Value = Null
    If frm.NewRecord Then
        scheduleChanged = True
    Else
        scheduleChanged = Nz(frm!StartDate.OldValue, 0) <> Nz(frm!StartDate.Value, 0) Or _
                          Nz(frm!DueDay.OldValue, 0) <> Nz(frm!DueDay.Value, 0) Or _
                          Nz(frm!Frequency.OldValue, "") <> Nz(frm!Frequency.Value, "")
    End If
    If scheduleChanged Or IsNull(frm!NextDueDate.Value) Then
        ' after the last created one, else from the start date
        If IsNull(frm!LastCreatedDate.Value) Then
            frm!NextDueDate.Value = FirstDueDate(frm!StartDate.Value, frm!DueDay.Value)
        Else
            frm!NextDueDate.Value = DueAfter(DateSerial(Year(frm!LastCreatedDate.Value), Month(frm!LastCreatedDate.Value), 1), _
                                             frm!Frequency.Value, frm!DueDay.Value)
        End If
    End If
    ValidateRecurring = True
End Function

Public Sub RecurringCurrent(ByVal frm As Access.Form)
    Dim n As Long, total As Currency, info As String
    If frm.NewRecord Then
        info = "يُنشأ المصروف في يوم الاستحقاق من كل فترة، بدءًا من تاريخ البداية."
    Else
        n = Nz(DbValue("SELECT COUNT(*) FROM Expenses WHERE RecurringID = " & Nz(frm!RecurringID.Value, 0)), 0)
        total = Nz(DbValue("SELECT Sum(TotalAmount) FROM Expenses WHERE RecurringID = " & Nz(frm!RecurringID.Value, 0)), 0)
        info = "أُنشئ منه " & n & " مصروف بإجمالي " & Format$(total, "#,##0.00")
        If Not frm!IsActive.Value Then
            info = info & "   (متوقف)"
        ElseIf Not IsNull(frm!NextDueDate.Value) Then
            If frm!NextDueDate.Value <= Date Then info = info & "   - مستحق الآن: اضغط «إنشاء المستحق الآن»"
        End If
    End If
    frm!lblRecurringInfo.Caption = Tr(info)
End Sub

Public Sub RecurringCreateNow(ByVal frm As Access.Form)
    Dim created As Long, msg As String
    If frm.Dirty Then
        If Not SaveRecord(frm) Then Exit Sub
    End If
    If DueRecurringCount() = 0 Then
        ShowInfo "لا توجد مصروفات متكررة مستحقة حتى اليوم."
        Exit Sub
    End If
    If Not AskYesNo("إنشاء كل المصروفات المتكررة المستحقة حتى اليوم (" & GDate(Date) & ")؟") Then Exit Sub
    msg = CreateDueExpenses(Date, created)
    frm.Requery
    RefreshList frm
    ShowRecurringResult created, msg
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Public Function TestRecurring() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, passed As Long, failed As Long, report As String, msg As String
    Dim id As Long, created As Long, start As Date, expected As Long, nextDue As Date
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestRecurring  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Check FirstDueDate(DateSerial(2026, 1, 10), 5) = DateSerial(2026, 2, 5) And _
          FirstDueDate(DateSerial(2026, 1, 5), 5) = DateSerial(2026, 1, 5), "أول استحقاق في يوم الاستحقاق أو بعد البداية", _
          passed, failed, report
    Check DueAfter(DateSerial(2026, 11, 28), "QUARTERLY", 28) = DateSerial(2027, 2, 28) And _
          DueAfter(DateSerial(2026, 3, 1), "YEARLY", 1) = DateSerial(2027, 3, 1), "الاستحقاق التالي ربع سنوي وسنوي", _
          passed, failed, report
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    start = DateSerial(Year(Date), Month(Date) - 2, 1)                  ' three monthly dates until today
    If ClosedThroughDate() >= start Then
        Debug.Print "[--] الفترة مقفلة: الاختبار يحتاج آخر ثلاثة أشهر مفتوحة"
        GoTo Undo
    End If
    CurrentDb.Execute "INSERT INTO RecurringExpenses (RecurringName, ExpenseTypeID, Amount, Tax, PaymentMethodID, Frequency, DueDay, " & _
        "StartDate, NextDueDate, IsActive) VALUES ('TEST إيجار', " & _
        Nz(DbValue("SELECT Min(ExpenseTypeID) FROM ExpenseTypes"), 1) & ", 1000, 150, " & BANK_TRANSFER_METHOD_ID & _
        ", 'MONTHLY', 1, " & SqlDate(start) & ", " & SqlDate(FirstDueDate(start, 1)) & ", True)", dbFailOnError
    id = DbValue("SELECT RecurringID FROM RecurringExpenses WHERE RecurringName = 'TEST إيجار'")
    expected = 3
    Check DueRecurringCount() >= 1, "المصروف المستحق يظهر في العدد", passed, failed, report
    msg = CreateDueExpenses(Date, created, id)
    nextDue = DbValue("SELECT NextDueDate FROM RecurringExpenses WHERE RecurringID = " & id)
    Check Len(msg) = 0 And created = expected And _
          Nz(DbValue("SELECT Sum(TotalAmount) FROM Expenses WHERE RecurringID = " & id), 0) = expected * 1150 And _
          nextDue = DateSerial(Year(Date), Month(Date) + 1, 1), "مصروف لكل شهر حتى اليوم، والاستحقاق التالي الشهر القادم " & msg, _
          passed, failed, report
    msg = CreateDueExpenses(Date, created, id)
    Check Len(msg) = 0 And created = 0, "التشغيل الثاني لا يكرر المصروفات", passed, failed, report
    msg = SyncJournal()
    Check Len(msg) = 0 And Nz(DbValue("SELECT COUNT(*) FROM JournalEntries AS j INNER JOIN Expenses AS e ON j.SourceID = " & _
          "e.ExpenseID WHERE j.SourceType = 'EXPENSE' AND e.RecurringID = " & id), 0) = expected And _
          Nz(DbValue("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0) = 0, _
          "لكل مصروف قيد متوازن " & msg, passed, failed, report
    CurrentDb.Execute "UPDATE RecurringExpenses SET NextDueDate = " & SqlDate(start) & " WHERE RecurringID = " & id, dbFailOnError
    msg = CreateDueExpenses(Date, created, id)
    Check created = 0 And Nz(DbValue("SELECT COUNT(*) FROM Expenses WHERE RecurringID = " & id), 0) = expected, _
          "التاريخ الذي له مصروف لا يُنشأ مرة أخرى", passed, failed, report
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
        TestMsg "جميع اختبارات المصروفات المتكررة ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestRecurring"
        TestRecurring = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestRecurring"
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
