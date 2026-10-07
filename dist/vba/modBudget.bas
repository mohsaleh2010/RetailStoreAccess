Attribute VB_Name = "modBudget"
'==============================================================================
' modBudget  -  Retail Store Management System
'
' The budget (Budgets + BudgetLines, screen frmBudget, permission BUDGET):
'   one budget per year; a line per revenue or expense account (a main account
'   takes its sub-accounts), optionally per cost centre, with an amount for each
'   month (M1..M12).
'   AddAllAccounts     a line for every active revenue / expense sub-account
'   FillFromActual     the months from the actual of a year (+ a percentage)
'   SpreadAnnual       an annual amount over the twelve months
' BudgetVsActualQuery: for a period (the budget of the year of its start, whole
' months), budget, actual from the journal, variance, its percentage and whether it
' is favourable (revenue above / expenses below the budget).
'==============================================================================
Option Compare Database
Option Explicit

Private Function CanBudget(ByVal Action As String) As String
    If Not HasPermission("BUDGET") Then
        CanBudget = "·«  „·ﬂ ’·«ÕÌ… «·„Ê«“‰…."
    ElseIf Not CanScreenAction("frmBudget", Action, True) Then
        CanBudget = "·«  „·ﬂ Â–Â «·’·«ÕÌ… ›Ì ‘«‘… «·„Ê«“‰…."
    End If
End Function

Public Function BudgetAccountProblem(ByVal AccountCode As Variant) As String
    If IsNull(AccountCode) Then
        BudgetAccountProblem = "«Œ — «·Õ”«»."
    ElseIf IsNull(DbValue("SELECT AccountCode FROM Accounts WHERE AccountCode = " & CLng(AccountCode) & _
                          " AND AccountType IN ('REVENUE', 'EXPENSE')")) Then
        BudgetAccountProblem = "«·„Ê«“‰… ·Õ”«»«  «·≈Ì—«œ«  Ê«·„’—Ê›«  ›ﬁÿ."
    End If
End Function

Public Function BudgetLineProblem(ByVal BudgetID As Long, ByVal LineID As Long, ByVal AccountCode As Variant, _
                                  ByVal CostCenterID As Variant) As String
    BudgetLineProblem = BudgetAccountProblem(AccountCode)
    If Len(BudgetLineProblem) > 0 Then Exit Function
    If Nz(DbValue("SELECT COUNT(*) FROM BudgetLines WHERE BudgetID = " & BudgetID & " AND BudgetLineID <> " & LineID & _
                  " AND AccountCode = " & CLng(AccountCode) & " AND " & _
                  IIf(IsNull(CostCenterID), "CostCenterID Is Null", "CostCenterID = " & CLng(Nz(CostCenterID, 0)))), 0) > 0 Then
        BudgetLineProblem = "«·Õ”«» " & AccountCode & " „ÊÃÊœ ›Ì «·„Ê«“‰… ·‰›” «·„—ﬂ“."
    End If
End Function

Public Function CreateBudget(ByVal BudgetYear As Long, ByVal BudgetName As String, ByRef BudgetID As Long) As String
    Dim rs As DAO.Recordset
    BudgetID = 0
    CreateBudget = CanBudget("ADD")
    If Len(CreateBudget) > 0 Then Exit Function
    If BudgetYear < 2000 Or BudgetYear > 2100 Then
        CreateBudget = "«ﬂ » «·”‰…."
    ElseIf Not IsNull(DbValue("SELECT BudgetID FROM Budgets WHERE BudgetYear = " & BudgetYear)) Then
        CreateBudget = " ÊÃœ „Ê«“‰… ·”‰… " & BudgetYear & " »«·›⁄·."
    End If
    If Len(CreateBudget) > 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("Budgets", dbOpenDynaset)
    rs.AddNew
    rs!BudgetYear = BudgetYear
    rs!BudgetName = Left$(IIf(Len(Trim$(BudgetName)) > 0, Trim$(BudgetName), "„Ê«“‰… " & BudgetYear), 100)
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Bookmark = rs.LastModified
    BudgetID = rs!BudgetID
    rs.Close
    LogAction "BUDGET_CREATE", "Budgets", CStr(BudgetYear)
End Function

Public Function DeleteBudget(ByVal BudgetID As Long) As String
    DeleteBudget = CanBudget("DELETE")
    If Len(DeleteBudget) > 0 Then Exit Function
    CurrentDb.Execute "DELETE FROM Budgets WHERE BudgetID = " & BudgetID, dbFailOnError       ' lines cascade
    LogAction "BUDGET_DELETE", "Budgets", CStr(BudgetID)
End Function

Public Function AddAllAccounts(ByVal BudgetID As Long, ByRef Added As Long) As String
    ' A line (all centres) for every active revenue / expense sub-account not in the budget yet.
    Added = 0
    AddAllAccounts = CanBudget("ADD")
    If Len(AddAllAccounts) > 0 Then Exit Function
    Added = Nz(DbValue("SELECT COUNT(*) FROM Accounts WHERE IsPosting = True AND IsActive = True AND AccountType IN " & _
                       "('REVENUE', 'EXPENSE') AND AccountCode NOT IN (SELECT AccountCode FROM BudgetLines WHERE BudgetID = " & _
                       BudgetID & ")"), 0)
    CurrentDb.Execute "INSERT INTO BudgetLines (BudgetID, AccountCode, M1, M2, M3, M4, M5, M6, M7, M8, M9, M10, M11, M12) " & _
        "SELECT " & BudgetID & ", AccountCode, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 FROM Accounts WHERE IsPosting = True AND " & _
        "IsActive = True AND AccountType IN ('REVENUE', 'EXPENSE') AND AccountCode NOT IN (SELECT AccountCode FROM BudgetLines " & _
        "WHERE BudgetID = " & BudgetID & ")", dbFailOnError
End Function

Public Function ActualOfMonth(ByVal AccountCode As Long, ByVal CostCenterID As Variant, ByVal Y As Long, _
                              ByVal M As Long) As Currency
    ' The actual of an account (with its sub-accounts) in a month, revenue and expenses positive.
    Dim t As String, n As Long, tree As String
    For n = 1 To 5
        tree = tree & IIf(n > 1, " OR ", "") & "d.Level" & n & "Code = " & AccountCode
    Next
    t = Nz(DbValue("SELECT AccountType FROM Accounts WHERE AccountCode = " & AccountCode), "EXPENSE")
    ActualOfMonth = Nz(DbValue("SELECT Sum(" & IIf(t = "REVENUE", "l.Credit - l.Debit", "l.Debit - l.Credit") & ") FROM " & _
        "(JournalLines AS l INNER JOIN JournalEntries AS e ON l.EntryID = e.EntryID) INNER JOIN Accounts AS d ON " & _
        "l.AccountCode = d.AccountCode WHERE (" & tree & ") AND e.SourceType <> 'YEAR_CLOSE' AND e.EntryDate >= " & _
        SqlDate(DateSerial(Y, M, 1)) & " AND e.EntryDate < " & SqlDate(DateSerial(Y, M + 1, 1)) & _
        IIf(IsNull(CostCenterID), "", " AND l.CostCenterID = " & CLng(Nz(CostCenterID, 0)))), 0)
End Function

Public Function FillFromActual(ByVal BudgetID As Long, ByVal FromYear As Long, ByVal Percent As Double) As String
    ' Every line: each month = the actual of the same month of FromYear x (1 + Percent).
    Dim rs As DAO.Recordset, m As Long
    FillFromActual = CanBudget("EDIT")
    If Len(FillFromActual) > 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM BudgetLines WHERE BudgetID = " & BudgetID, dbOpenDynaset)
    Do Until rs.EOF
        rs.Edit
        For m = 1 To 12
            rs.Fields("M" & m).Value = Round(ActualOfMonth(rs!AccountCode, rs!CostCenterID, FromYear, m) * (1 + Percent), 0)
            If rs.Fields("M" & m).Value < 0 Then rs.Fields("M" & m).Value = 0
        Next
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    LogAction "BUDGET_FILL", "Budgets", CStr(BudgetID), FromYear & " " & Format$(Percent, "0%")
End Function

Public Function SpreadAnnual(ByVal BudgetLineID As Long, ByVal Annual As Currency) As String
    ' The annual amount in twelve equal months (the rounding in December).
    Dim rs As DAO.Recordset, m As Long, part As Currency
    SpreadAnnual = CanBudget("EDIT")
    If Len(SpreadAnnual) = 0 And Annual < 0 Then SpreadAnnual = "«·„»·€ ·« ÌﬂÊ‰ ”«·»«."
    If Len(SpreadAnnual) > 0 Then Exit Function
    part = Round(Annual / 12, 2)
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM BudgetLines WHERE BudgetLineID = " & BudgetLineID, dbOpenDynaset)
    If rs.EOF Then
        SpreadAnnual = "«Œ — ”ÿ—« „‰ «·„Ê«“‰…."
    Else
        rs.Edit
        For m = 1 To 11
            rs.Fields("M" & m).Value = part
        Next
        rs!M12 = Annual - 11 * part
        rs.Update
    End If
    rs.Close
End Function

'------------------------------------------------------------------------------
' Screen frmBudget (subform frmBudgetLines on BudgetLines)
'------------------------------------------------------------------------------
Private Function ShownBudget(ByVal frm As Access.Form) As Long
    ShownBudget = Nz(frm!txtBudgetID.Value, 0)
End Function

Public Sub BudgetLoad(ByVal frm As Access.Form)
    Dim id As Variant
    Calendar = vbCalGreg
    SyncJournal
    frm!txtYear.Value = Year(Date)
    frm!txtFillYear.Value = Year(Date) - 1
    frm!txtPercent.Value = 0
    frm!txtFrom.Value = DateSerial(Year(Date), 1, 1)
    frm!txtTo.Value = DateSerial(Year(Date), Month(Date) + 1, 0)
    id = DbValue("SELECT BudgetID FROM Budgets WHERE BudgetYear = " & Year(Date))
    If IsNull(id) Then id = DbValue("SELECT TOP 1 BudgetID FROM Budgets ORDER BY BudgetYear DESC")
    BudgetShow frm, Nz(id, 0)
End Sub

Public Sub BudgetShow(ByVal frm As Access.Form, ByVal BudgetID As Long)
    frm!txtBudgetID.Value = IIf(BudgetID = 0, Null, BudgetID)
    frm!cboBudget.Requery
    frm!cboBudget.Value = IIf(BudgetID = 0, Null, BudgetID)
    frm!subLines.Form.RecordSource = "SELECT * FROM BudgetLines WHERE BudgetID = " & BudgetID & " ORDER BY AccountCode"
    frm!subLines.Form.AllowAdditions = (BudgetID > 0)
    If BudgetID > 0 Then
        frm!txtFrom.Value = DateSerial(DbValue("SELECT BudgetYear FROM Budgets WHERE BudgetID = " & BudgetID), 1, 1)
    End If
    BudgetTotals frm
End Sub

Public Sub BudgetTotals(ByVal frm As Access.Form)
    Dim id As Long, rev As Currency, costs As Currency, sumSql As String
    id = ShownBudget(frm)
    If id = 0 Then
        frm!lblTotals.Caption = "√‰‘∆ „Ê«“‰… ··”‰…."
        Exit Sub
    End If
    sumSql = "Sum(b.M1 + b.M2 + b.M3 + b.M4 + b.M5 + b.M6 + b.M7 + b.M8 + b.M9 + b.M10 + b.M11 + b.M12)"
    rev = Nz(DbValue("SELECT " & sumSql & " FROM BudgetLines AS b INNER JOIN Accounts AS a ON b.AccountCode = a.AccountCode " & _
                     "WHERE b.BudgetID = " & id & " AND a.AccountType = 'REVENUE'"), 0)
    costs = Nz(DbValue("SELECT " & sumSql & " FROM BudgetLines AS b INNER JOIN Accounts AS a ON b.AccountCode = a.AccountCode " & _
                      "WHERE b.BudgetID = " & id & " AND a.AccountType = 'EXPENSE'"), 0)
    frm!lblTotals.Caption = "„Ê«“‰… «·”‰…: «·≈Ì—«œ«  " & Format$(rev, "#,##0") & "   «·„’—Ê›«  " & Format$(costs, "#,##0") & _
                            "   ’«›Ì «·—»Õ «·„ Êﬁ⁄ " & Format$(rev - costs, "#,##0")
End Sub

Public Sub BudgetLineInsert(ByVal lines As Access.Form)
    ' A new line of the subform belongs to the budget shown.
    lines!BudgetID.Value = lines.Parent!txtBudgetID.Value
End Sub

Public Sub BudgetLineCheck(ByVal lines As Access.Form, ByRef Cancel As Integer)
    Dim msg As String, m As Long
    msg = BudgetLineProblem(Nz(lines!BudgetID.Value, 0), Nz(lines!BudgetLineID.Value, 0), lines!AccountCode.Value, _
                            lines!CostCenterID.Value)
    For m = 1 To 12
        If IsNull(lines("M" & m).Value) Then lines("M" & m).Value = 0
        If Len(msg) = 0 And lines("M" & m).Value < 0 Then msg = "„»«·€ «·„Ê«“‰… ·«  ﬂÊ‰ ”«·»…."
    Next
    If Len(msg) = 0 And Len(CanBudget(IIf(lines.NewRecord, "ADD", "EDIT"))) > 0 Then msg = CanBudget(IIf(lines.NewRecord, "ADD", "EDIT"))
    If Len(msg) > 0 Then
        ShowWarning msg
        Cancel = True
    End If
End Sub

Public Sub BudgetLineSaved(ByVal lines As Access.Form)
    BudgetTotals lines.Parent
End Sub

Public Sub BudgetPick(ByVal frm As Access.Form)
    If Not IsNull(frm!cboBudget.Value) Then BudgetShow frm, CLng(frm!cboBudget.Value)
End Sub

Public Sub BudgetCreate(ByVal frm As Access.Form)
    Dim msg As String, id As Long, n As Long
    msg = CreateBudget(CLng(Nz(frm!txtYear.Value, 0)), "", id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    AddAllAccounts id, n
    BudgetShow frm, id
    ShowInfo "√ı‰‘∆  «·„Ê«“‰… »”ÿ— ·ﬂ· Õ”«» ≈Ì—«œ«  Ê„’—Ê›«  (" & n & "). «ﬂ » «·„»«·€° √Ê «‰”ŒÂ« „‰ ›⁄·Ì ”‰… ”«»ﬁ…."
End Sub

Public Sub BudgetAddAccounts(ByVal frm As Access.Form)
    Dim msg As String, n As Long
    If ShownBudget(frm) = 0 Then Exit Sub
    msg = AddAllAccounts(ShownBudget(frm), n)
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        ShowInfo "√ı÷Ì› " & n & " Õ”«»."
    End If
    BudgetShow frm, ShownBudget(frm)
End Sub

Public Sub BudgetFill(ByVal frm As Access.Form)
    Dim msg As String
    If ShownBudget(frm) = 0 Then Exit Sub
    If Not IsNumeric(frm!txtFillYear.Value) Then
        ShowWarning "«ﬂ » «·”‰… «· Ì Ìı‰”Œ ›⁄·ÌÂ«."
        Exit Sub
    End If
    If Not AskYesNo("„·¡ ﬂ· √‘Â— «·„Ê«“‰… „‰ ›⁄·Ì ”‰… " & frm!txtFillYear.Value & " »“Ì«œ… " & _
                    Format$(Nz(frm!txtPercent.Value, 0), "0%") & "ø  ı” »œ· «·„»«·€ «·Õ«·Ì….") Then Exit Sub
    DoCmd.Hourglass True
    msg = FillFromActual(ShownBudget(frm), CLng(frm!txtFillYear.Value), Nz(frm!txtPercent.Value, 0))
    DoCmd.Hourglass False
    If Len(msg) > 0 Then ShowWarning msg
    BudgetShow frm, ShownBudget(frm)
End Sub

Public Sub BudgetSpread(ByVal frm As Access.Form)
    Dim msg As String, id As Variant
    id = frm!subLines.Form!BudgetLineID.Value
    If IsNull(id) Or Not IsNumeric(frm!txtAnnual.Value) Then
        ShowWarning "«Œ — ”ÿ—« ›Ì «·ÃœÊ· Ê«ﬂ » «·„»·€ «·”‰ÊÌ."
        Exit Sub
    End If
    msg = SpreadAnnual(CLng(id), CCur(frm!txtAnnual.Value))
    If Len(msg) > 0 Then ShowWarning msg
    frm!subLines.Form.Requery
    BudgetTotals frm
End Sub

Public Sub BudgetDelete(ByVal frm As Access.Form)
    Dim msg As String
    If ShownBudget(frm) = 0 Then Exit Sub
    If Not AskYesNo("Õ–› «·„Ê«“‰… ﬂ·Â« »√”ÿ—Â«ø") Then Exit Sub
    msg = DeleteBudget(ShownBudget(frm))
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        BudgetShow frm, 0
    End If
End Sub

Private Function ComparePeriodOK(ByVal frm As Access.Form) As Boolean
    If Not IsDate(frm!txtFrom.Value) Or Not IsDate(frm!txtTo.Value) Then
        ShowWarning "«ﬂ » «·› —…."
    ElseIf DateValue(frm!txtTo.Value) < DateValue(frm!txtFrom.Value) Then
        ShowWarning "‰Â«Ì… «·› —… ﬁ»· »œ«Ì Â«."
    ElseIf Year(frm!txtTo.Value) <> Year(frm!txtFrom.Value) Then
        ShowWarning "«·„ﬁ«—‰… œ«Œ· ”‰… Ê«Õœ… („Ê«“‰… «·”‰…)."
    Else
        ComparePeriodOK = True
    End If
End Function

Public Sub BudgetCompare(ByVal frm As Access.Form)
    Dim budget As Currency, actual As Currency
    If Not ComparePeriodOK(frm) Then Exit Sub
    SyncJournal
    SetPeriod DateValue(frm!txtFrom.Value), DateValue(frm!txtTo.Value)
    frm!lstVariance.RowSource = "SELECT BudgetLineID, AccountCode AS [«·Õ”«»], AccountName AS [«”„ «·Õ”«»], BudgetCenter " & _
        "AS [«·„—ﬂ“], Format(BudgetAmount, '#,##0') AS [«·„Ê«“‰…], Format(ActualAmount, '#,##0') AS [«·›⁄·Ì], " & _
        "Format(Variance, '#,##0') AS [«·«‰Õ—«›], Format(VariancePct, '0%') AS [%], VarianceNote AS [ ] " & _
        "FROM BudgetVsActualQuery WHERE BudgetAmount <> 0 OR ActualAmount <> 0 ORDER BY AccountType DESC, TreeKey"
    budget = Nz(DbValue("SELECT Sum(IIf(AccountType = 'REVENUE', BudgetAmount, -BudgetAmount)) FROM BudgetVsActualQuery"), 0)
    actual = Nz(DbValue("SELECT Sum(IIf(AccountType = 'REVENUE', ActualAmount, -ActualAmount)) FROM BudgetVsActualQuery"), 0)
    frm!lblVariance.Caption = "’«›Ì «·—»Õ: «·„Ê«“‰… " & Format$(budget, "#,##0") & "   «·›⁄·Ì " & Format$(actual, "#,##0") & _
                              "   «·«‰Õ—«› " & Format$(actual - budget, "#,##0")
    frm!lblVariance.ForeColor = IIf(actual >= budget, CLR_SUCCESS, CLR_DANGER)
End Sub

Public Sub PrintBudgetVariance(ByVal frm As Access.Form)
    If Not ComparePeriodOK(frm) Then Exit Sub
    SyncJournal
    SetPeriod DateValue(frm!txtFrom.Value), DateValue(frm!txtTo.Value)
    LogAction "REPORT", "BUDGET_VS_ACTUAL"
    OpenReportOrQuery "rptBudgetVsActual", "BudgetVsActualQuery", "BudgetAmount <> 0 OR ActualAmount <> 0", _
                      PeriodText(DateValue(frm!txtFrom.Value), DateValue(frm!txtTo.Value))
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Private Sub CheckBudget(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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

Public Function TestBudget() As Boolean
    Dim passed As Long, failed As Long, report As String, msg As String, ws As DAO.Workspace, inTrans As Boolean
    Dim id As Long, n As Long, y As Long, lineID As Long, actual As Currency
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestBudget  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    y = Year(Date) + 50                                      ' a year no budget uses
    msg = CreateBudget(y, "TEST", id)
    CheckBudget Len(msg) = 0 And id > 0, "≈‰‘«¡ „Ê«“‰… " & msg, passed, failed, report
    CheckBudget Len(CreateBudget(y, "TEST", n)) > 0, "„Ê«“‰… Ê«Õœ… ·ﬂ· ”‰…", passed, failed, report
    msg = AddAllAccounts(id, n)
    CheckBudget Len(msg) = 0 And n > 0 And Nz(DbValue("SELECT COUNT(*) FROM BudgetLines WHERE BudgetID = " & id), 0) = n, _
                "”ÿ— ·ﬂ· Õ”«» ≈Ì—«œ«  Ê„’—Ê›«  (" & n & ")", passed, failed, report
    lineID = DbValue("SELECT BudgetLineID FROM BudgetLines WHERE BudgetID = " & id & " AND AccountCode = 4100")
    msg = SpreadAnnual(lineID, 1000)
    CheckBudget Len(msg) = 0 And DbValue("SELECT M1 FROM BudgetLines WHERE BudgetLineID = " & lineID) = 83.33 And _
                DbValue("SELECT M12 FROM BudgetLines WHERE BudgetLineID = " & lineID) = 83.37, _
                " Ê“Ì⁄ «·„»·€ «·”‰ÊÌ ⁄·Ï «·√‘Â—", passed, failed, report
    CheckBudget Len(BudgetLineProblem(id, 0, 4100, Null)) > 0 And Len(BudgetLineProblem(id, 0, 1300, Null)) > 0, _
                "·« Ì ﬂ—— «·Õ”«»° Ê·« Õ”«»«  «·„Ì“«‰Ì…", passed, failed, report
    ' compare with the actual: the budget of this year moved to the test year is the same months
    CurrentDb.Execute "UPDATE Budgets SET BudgetYear = " & Year(Date) & " WHERE BudgetID = " & id & _
                      " AND " & Year(Date) & " NOT IN (SELECT BudgetYear FROM Budgets)", dbFailOnError
    If DbValue("SELECT BudgetYear FROM Budgets WHERE BudgetID = " & id) = Year(Date) Then
        SetPeriod DateSerial(Year(Date), 1, 1), DateSerial(Year(Date), 12, 31)
        actual = 0
        For n = 1 To 12
            actual = actual + ActualOfMonth(4100, Null, Year(Date), n)
        Next
        CheckBudget Nz(DbValue("SELECT BudgetAmount FROM BudgetVsActualQuery WHERE BudgetLineID = " & lineID), 0) = 1000 And _
                    Nz(DbValue("SELECT ActualAmount FROM BudgetVsActualQuery WHERE BudgetLineID = " & lineID), 0) = actual, _
                    "«·„Ê«“‰… „ﬁ«»· «·›⁄·Ì ··„»Ì⁄« ", passed, failed, report
    End If
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckBudget False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·„Ê«“‰… ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestBudget"
        TestBudget = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestBudget"
    End If
End Function
