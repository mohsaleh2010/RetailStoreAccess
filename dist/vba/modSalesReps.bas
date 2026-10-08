Attribute VB_Name = "modSalesReps"
'==============================================================================
' modSalesReps  -  Retail Store Management System
'
' Sales representatives (SalesReps, screens frmSalesReps / frmRepTargets /
' frmCommissions, permission SALES_REPS):
'   SalesRepFor      the rep of a new sale or customer payment: the active rep of
'                    the customer, else the active rep linked to the current user.
'                    A sales return takes the rep of its invoice (modSales).
'   targets          a monthly target per rep (SalesRepTargets), on net sales.
'   commissions      a monthly run (CommissionRuns + CommissionLines): a line per
'                    rep from RepPerformanceQuery (net sales without VAT or the
'                    collections, x the rep's rate), edited as a draft, then
'                    posted on the last day of the month (journal source COMMISSION):
'                    5530 commissions expense (rep's cost centre) against 2330
'                    commissions payable. Paid by a cash voucher of category
'                    COMMISSION (modCash), which debits 2330.
'   commission       = Max(0, Round(base x rate, 2) + adjustment)
' Same amounts as tools/sales_reps_reference.py.
'==============================================================================
Option Compare Database
Option Explicit

Private Function CommissionRunField(ByVal RunID As Long, ByVal FieldName As String) As Variant
    CommissionRunField = DbValue("SELECT " & FieldName & " FROM CommissionRuns WHERE CommissionRunID = " & RunID)
End Function

Private Function CanCommissions(ByVal Action As String) As String
    If Not HasPermission("SALES_REPS") Then
        CanCommissions = "·«  „·ﬂ ’·«ÕÌ… «·„‰œÊ»Ì‰ Ê⁄„Ê·« Â„."
    ElseIf Not CanScreenAction("frmCommissions", Action, True) Then
        CanCommissions = "·«  „·ﬂ Â–Â «·’·«ÕÌ… ›Ì ‘«‘… ⁄„Ê·«  «·„‰œÊ»Ì‰."
    End If
End Function

'------------------------------------------------------------------------------
' The rep of a document
'------------------------------------------------------------------------------
Public Function SalesRepFor(ByVal CustomerID As Long) As Variant
    ' The active rep of the customer, else the active rep linked to the current user, else Null.
    Dim v As Variant
    v = DbValue("SELECT c.SalesRepID FROM Customers AS c INNER JOIN SalesReps AS s ON c.SalesRepID = s.SalesRepID " & _
                "WHERE s.IsActive = True AND c.CustomerID = " & CustomerID)
    If IsNull(v) Then
        v = DbValue("SELECT Min(SalesRepID) FROM SalesReps WHERE IsActive = True AND EmployeeID = " & CurrentUserID())
    End If
    SalesRepFor = v
End Function

'------------------------------------------------------------------------------
' Amounts
'------------------------------------------------------------------------------
Public Function CalcCommission(ByVal BaseAmount As Currency, ByVal CommissionRate As Double, _
                               ByVal Adjustment As Currency) As Currency
    Dim c As Currency
    c = RoundMoney(CDbl(BaseAmount) * CommissionRate) + Adjustment            ' modCommon
    If c < 0 Then c = 0
    CalcCommission = c
End Function

Public Function RepPayable(ByVal SalesRepID As Long) As Currency
    ' Posted commissions not paid yet.
    RepPayable = Nz(DbValue("SELECT Payable FROM RepCommissionBalanceQuery WHERE SalesRepID = " & SalesRepID), 0)
End Function

Public Function CommissionVoucherProblem(ByVal VoucherType As String, ByVal Category As String, _
                                         ByVal SalesRepID As Variant, ByVal Amount As Currency) As String
    ' A cash voucher paying a commission (modCash.SaveCashVoucher): the rep and his payable balance.
    If Category <> "COMMISSION" Then Exit Function
    If VoucherType <> "OUT" Then
        CommissionVoucherProblem = "’—› «·⁄„Ê·… ÌﬂÊ‰ »”‰œ ’—›."
    ElseIf IsNull(SalesRepID) Then
        CommissionVoucherProblem = "«Œ — «·„‰œÊ»."
    ElseIf Amount > RepPayable(CLng(SalesRepID)) Then
        CommissionVoucherProblem = "«·„»·€ √ﬂ»— „‰ ⁄„Ê·«  «·„‰œÊ» «·„” Õﬁ… (" & _
                                   Format$(RepPayable(CLng(SalesRepID)), "#,##0.00") & ")."
    End If
End Function

Private Sub SaveCommissionAmounts(ByVal rs As DAO.Recordset)
    ' The base and the commission of the current record of rs (inside its Edit / AddNew).
    If rs!CommissionBase = "COLLECTION" Then
        rs!BaseAmount = rs!Collections
    Else
        rs!BaseAmount = rs!NetSales
    End If
    rs!Commission = CalcCommission(rs!BaseAmount, CDbl(rs!CommissionRate), rs!Adjustment)
End Sub

Private Sub UpdateRunTotal(ByVal RunID As Long)
    CurrentDb.Execute "UPDATE CommissionRuns SET TotalAmount = " & _
        Str$(Nz(DbValue("SELECT Sum(Commission) FROM CommissionLines WHERE CommissionRunID = " & RunID), 0)) & _
        " WHERE CommissionRunID = " & RunID, dbFailOnError
End Sub

Private Sub AddCommissionLines(ByVal RunID As Long, ByVal RunMonth As Date)
    ' A line for each rep with sales or collections in the month (RepPerformanceQuery).
    Dim db As DAO.Database, p As DAO.Recordset, l As DAO.Recordset
    Set db = CurrentDb
    SetPeriod DateSerial(Year(RunMonth), Month(RunMonth), 1), RunMonth
    Set p = db.OpenRecordset("SELECT p.*, s.CostCenterID FROM RepPerformanceQuery AS p INNER JOIN SalesReps AS s " & _
                             "ON p.SalesRepID = s.SalesRepID WHERE p.NetSales <> 0 OR p.Collections <> 0 " & _
                             "ORDER BY p.RepName", dbOpenSnapshot)
    Set l = db.OpenRecordset("CommissionLines", dbOpenDynaset, dbAppendOnly)
    Do Until p.EOF
        l.AddNew
        l!CommissionRunID = RunID
        l!SalesRepID = p!SalesRepID
        l!RepName = Left$(p!RepName, 100)
        l!CostCenterID = p!CostCenterID
        l!NetSales = Nz(p!NetSales, 0)
        l!Collections = Nz(p!Collections, 0)
        l!CommissionBase = Nz(p!CommissionBase, "SALES")
        l!CommissionRate = Nz(p!CommissionRate, 0)
        l!Adjustment = 0
        SaveCommissionAmounts l
        l.Update
        p.MoveNext
    Loop
    p.Close
    l.Close
    UpdateRunTotal RunID
End Sub

'------------------------------------------------------------------------------
' Draft, posting
'------------------------------------------------------------------------------
Public Function CreateCommissionRun(ByVal AnyDayOfMonth As Date, ByRef RunID As Long) As String
    Dim m As Date, rs As DAO.Recordset
    RunID = 0
    CreateCommissionRun = CanCommissions("ADD")
    If Len(CreateCommissionRun) > 0 Then Exit Function
    m = MonthEnd(AnyDayOfMonth)                                      ' modAssets
    If Not IsNull(DbValue("SELECT CommissionRunID FROM CommissionRuns WHERE RunMonth = " & SqlDate(m))) Then
        CreateCommissionRun = "ÌÊÃœ „”Ì— ⁄„Ê·«  ·‘Â— " & Format$(m, "yyyy/mm") & " »«·›⁄·."
        Exit Function
    End If
    If DateSerial(Year(m), Month(m), 1) > Date Then
        CreateCommissionRun = "·« Ìı‰‘√ „”Ì— ·‘Â— ·„ Ì»œ√."
        Exit Function
    End If
    CreateCommissionRun = ClosedPeriodProblem(m)
    If Len(CreateCommissionRun) > 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("CommissionRuns", dbOpenDynaset)
    rs.AddNew
    rs!RunNumber = NextNumber("COMMISSION_RUN")
    rs!RunMonth = m
    rs!Status = "DRAFT"
    rs!TotalAmount = 0
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Bookmark = rs.LastModified
    RunID = rs!CommissionRunID
    rs.Close
    AddCommissionLines RunID, m
    LogAction "COMMISSION_CREATE", "CommissionRuns", Format$(m, "yyyy-mm")
End Function

Public Function RebuildCommissionRun(ByVal RunID As Long) As String
    ' The draft again from the documents of the month (the adjustments are lost).
    RebuildCommissionRun = CanCommissions("EDIT")
    If Len(RebuildCommissionRun) = 0 And Nz(CommissionRunField(RunID, "Status"), "") <> "DRAFT" Then
        RebuildCommissionRun = "«·„”Ì— ·Ì” „”Êœ…."
    End If
    If Len(RebuildCommissionRun) > 0 Then Exit Function
    CurrentDb.Execute "DELETE FROM CommissionLines WHERE CommissionRunID = " & RunID, dbFailOnError
    AddCommissionLines RunID, CommissionRunField(RunID, "RunMonth")
End Function

Public Sub RecalcCommissionRun(ByVal RunID As Long)
    ' The commission of every line of a draft (after the lines were edited).
    Dim rs As DAO.Recordset
    If Nz(CommissionRunField(RunID, "Status"), "") <> "DRAFT" Then Exit Sub
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM CommissionLines WHERE CommissionRunID = " & RunID, dbOpenDynaset)
    Do Until rs.EOF
        rs.Edit
        SaveCommissionAmounts rs
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    UpdateRunTotal RunID
End Sub

Public Function PostCommissionRun(ByVal RunID As Long) As String
    Dim rm As Variant
    PostCommissionRun = CanCommissions("EDIT")
    If Len(PostCommissionRun) > 0 Then Exit Function
    If Nz(CommissionRunField(RunID, "Status"), "") <> "DRAFT" Then
        PostCommissionRun = "«·„”Ì— €Ì— „ÊÃÊœ √Ê „—ÕÛ¯·."
        Exit Function
    End If
    rm = CommissionRunField(RunID, "RunMonth")
    RecalcCommissionRun RunID
    If Nz(DbValue("SELECT COUNT(*) FROM CommissionLines WHERE CommissionRunID = " & RunID), 0) = 0 Then
        PostCommissionRun = "«·„”Ì— »·« √”ÿ—."
    ElseIf Nz(DbValue("SELECT COUNT(*) FROM CommissionLines WHERE (CommissionRate < 0 OR CommissionRate >= 1) AND " & _
                      "CommissionRunID = " & RunID), 0) > 0 Then
        PostCommissionRun = "‰”»… «·⁄„Ê·… »Ì‰ 0% Ê 100%."
    Else
        PostCommissionRun = ClosedPeriodProblem(rm)
    End If
    If Len(PostCommissionRun) > 0 Then Exit Function
    CurrentDb.Execute "UPDATE CommissionRuns SET Status = 'POSTED', PostedAt = Now(), EmployeeID = " & CurrentUserID() & _
                      " WHERE CommissionRunID = " & RunID, dbFailOnError
    LogAction "COMMISSION_POST", "CommissionRuns", Format$(rm, "yyyy-mm")
    SyncJournal
End Function

Public Function UnpostCommissionRun(ByVal RunID As Long) As String
    Dim rs As DAO.Recordset, p As String
    UnpostCommissionRun = CanCommissions("EDIT")
    If Len(UnpostCommissionRun) > 0 Then Exit Function
    If Nz(CommissionRunField(RunID, "Status"), "") <> "POSTED" Then
        UnpostCommissionRun = "«·„”Ì— €Ì— „—ÕÛ¯·."
        Exit Function
    End If
    ' a rep paid more than what stays posted without this run
    Set rs = CurrentDb.OpenRecordset("SELECT SalesRepID, RepName, Commission FROM CommissionLines WHERE CommissionRunID = " & _
                                     RunID, dbOpenSnapshot)
    Do Until rs.EOF Or Len(p) > 0
        If RepPayable(rs!SalesRepID) < rs!Commission Then
            p = rs!RepName & ": ’ı—› ·Â Ã“¡ „‰ Â–Â «·⁄„Ê·« . √·€ˆ ”‰œ «·’—› √Ê·«."
        End If
        rs.MoveNext
    Loop
    rs.Close
    If Len(p) = 0 Then p = ClosedPeriodProblem(CommissionRunField(RunID, "RunMonth"))
    UnpostCommissionRun = p
    If Len(p) > 0 Then Exit Function
    CurrentDb.Execute "UPDATE CommissionRuns SET Status = 'DRAFT', PostedAt = Null WHERE CommissionRunID = " & RunID, _
                      dbFailOnError
    LogAction "COMMISSION_UNPOST", "CommissionRuns", CStr(RunID)
    SyncJournal
End Function

Public Function DeleteCommissionRun(ByVal RunID As Long) As String
    DeleteCommissionRun = CanCommissions("DELETE")
    If Len(DeleteCommissionRun) = 0 And Nz(CommissionRunField(RunID, "Status"), "") <> "DRAFT" Then
        DeleteCommissionRun = " ıÕ–› «·„”Êœ… ›ﬁÿ: √·€ˆ «· —ÕÌ· √Ê·«."
    End If
    If Len(DeleteCommissionRun) > 0 Then Exit Function
    Dim auditBefore As Collection
    Set auditBefore = AuditSnapshot("CommissionRuns", "CommissionRunID", RunID)      ' the record as it was (modAudit)
    CurrentDb.Execute "DELETE FROM CommissionRuns WHERE CommissionRunID = " & RunID, dbFailOnError   ' lines cascade
    AuditDeleted "COMMISSION_DELETE", "CommissionRuns", RunID, auditBefore
End Function

'------------------------------------------------------------------------------
' Screen frmCommissions (subform frmCommissionLines on CommissionLines)
'------------------------------------------------------------------------------
Private Function ShownCommissionRun(ByVal frm As Access.Form) As Long
    ShownCommissionRun = Nz(frm!txtRunID.Value, 0)
End Function

Public Sub CommissionsLoad(ByVal frm As Access.Form)
    Dim last As Variant
    Calendar = vbCalGreg
    frm!txtMonth.Value = DateSerial(Year(Date), Month(Date), 1)
    last = Null
    If Not IsNull(frm.OpenArgs) Then
        If IsNumeric(frm.OpenArgs) Then last = DbValue("SELECT CommissionRunID FROM CommissionRuns WHERE CommissionRunID = " & _
                                                        CLng(frm.OpenArgs))
    End If
    If IsNull(last) Then last = DbValue("SELECT TOP 1 CommissionRunID FROM CommissionRuns ORDER BY RunMonth DESC")
    If Not IsNull(last) Then
        CommissionsShow frm, CLng(last)
    Else
        CommissionsShow frm, 0
    End If
End Sub

Public Sub CommissionsShow(ByVal frm As Access.Form, ByVal RunID As Long)
    Dim status As String, draft As Boolean
    frm!txtRunID.Value = IIf(RunID = 0, Null, RunID)
    frm!cboRun.Requery
    frm!cboRun.Value = IIf(RunID = 0, Null, RunID)
    frm!subLines.Form.RecordSource = "SELECT * FROM CommissionLines WHERE CommissionRunID = " & RunID & " ORDER BY RepName"
    status = Nz(CommissionRunField(RunID, "Status"), "")
    draft = (status = "DRAFT")
    frm!subLines.Form.AllowEdits = draft
    frm!btnRebuild.Enabled = draft
    frm!btnPost.Enabled = draft
    frm!btnDelete.Enabled = draft
    frm!btnUnpost.Enabled = (status = "POSTED")
    frm!btnPrint.Enabled = (RunID <> 0)
    CommissionsTotals frm
End Sub

Public Sub CommissionsTotals(ByVal frm As Access.Form)
    Dim id As Long, rs As DAO.Recordset, state As String
    id = ShownCommissionRun(frm)
    If id = 0 Then
        frm!lblState.Caption = Tr("«Œ — «·‘Â— À„ ´≈‰‘«¡ „”Ì— «·‘Â—ª.")
        frm!lblTotals.Caption = Tr(" ")
        Exit Sub
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT COUNT(*) AS LineCount, Sum(NetSales) AS SumSales, Sum(Collections) AS SumCollected, " & _
                                     "Sum(Commission) AS SumCommission FROM CommissionLines WHERE CommissionRunID = " & id, _
                                     dbOpenSnapshot)
    If Nz(rs!LineCount, 0) > 0 Then
        frm!lblTotals.Caption = Tr(rs!LineCount & " „‰œÊ»   ’«›Ì «·„»Ì⁄«  " & Format$(Nz(rs!SumSales, 0), "#,##0.00") & _
            "   «· Õ’Ì· " & Format$(Nz(rs!SumCollected, 0), "#,##0.00") & "   ≈Ã„«·Ì «·⁄„Ê·«  " & _
            Format$(Nz(rs!SumCommission, 0), "#,##0.00"))
    Else
        frm!lblTotals.Caption = Tr("·« „»Ì⁄«  Ê·«  Õ’Ì·«  ··„‰œÊ»Ì‰ ›Ì Â–« «·‘Â—.")
    End If
    rs.Close
    Select Case Nz(CommissionRunField(id, "Status"), "")
        Case "DRAFT"
            state = "„”Êœ… ⁄„Ê·«  " & Format$(CommissionRunField(id, "RunMonth"), "yyyy/mm") & _
                    ": ⁄œ¯· «·‰”»… √Ê «· ⁄œÌ· À„ ´ —ÕÌ· «·„”Ì—ª. «·’—› »”‰œ ’—› ‰ﬁœÌ… „‰ »‰œ ´’—› ⁄„Ê·… „‰œÊ»ª."
        Case "POSTED"
            state = "⁄„Ê·«  " & Format$(CommissionRunField(id, "RunMonth"), "yyyy/mm") & " „—ÕÛ¯·… ›Ì " & _
                    GDate(CommissionRunField(id, "PostedAt"))
    End Select
    frm!lblState.Caption = Tr(state)
End Sub

Public Sub CommissionLineChanged(ByVal lines As Access.Form)
    ' The subform: the commission of the line as its rate or adjustment is typed.
    Dim baseAmount As Currency
    If lines.NewRecord Then Exit Sub
    If Nz(lines!CommissionBase.Value, "SALES") = "COLLECTION" Then
        baseAmount = Nz(lines!Collections.Value, 0)
    Else
        baseAmount = Nz(lines!NetSales.Value, 0)
    End If
    lines!BaseAmount.Value = baseAmount
    lines!Commission.Value = CalcCommission(baseAmount, CDbl(Nz(lines!CommissionRate.Value, 0)), Nz(lines!Adjustment.Value, 0))
    If lines.Dirty Then lines.Dirty = False
    UpdateRunTotal Nz(lines.Parent!txtRunID.Value, 0)
    CommissionsTotals lines.Parent
End Sub

Public Sub CommissionsPick(ByVal frm As Access.Form)
    If Not IsNull(frm!cboRun.Value) Then CommissionsShow frm, CLng(frm!cboRun.Value)
End Sub

Public Sub CommissionsCreate(ByVal frm As Access.Form)
    Dim id As Long, msg As String
    If Not IsDate(frm!txtMonth.Value) Then
        ShowWarning "«Œ — «·‘Â—."
        Exit Sub
    End If
    msg = CreateCommissionRun(DateValue(frm!txtMonth.Value), id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    CommissionsShow frm, id
End Sub

Private Sub CommissionsDone(ByVal frm As Access.Form, ByVal Msg As String, ByVal Success As String)
    If Len(Msg) > 0 Then
        ShowWarning Msg
    Else
        ShowInfo Success
    End If
    CommissionsShow frm, ShownCommissionRun(frm)
End Sub

Public Sub CommissionsRebuild(ByVal frm As Access.Form)
    If ShownCommissionRun(frm) = 0 Then Exit Sub
    If Not AskYesNo("≈⁄«œ… ≈‰‘«¡ √”ÿ— «·„”Ì— „‰ „»Ì⁄«  Ê Õ’Ì·«  «·‘Â—ø  ÷Ì⁄ «· ⁄œÌ·«  «·„ﬂ Ê»… ›ÌÂ.") Then Exit Sub
    CommissionsDone frm, RebuildCommissionRun(ShownCommissionRun(frm)), " „  ≈⁄«œ… ≈‰‘«¡ «·√”ÿ—."
End Sub

Public Sub CommissionsPost(ByVal frm As Access.Form)
    If ShownCommissionRun(frm) = 0 Then Exit Sub
    If Not AskYesNo(" —ÕÌ· «·„”Ì— Ê≈‰‘«¡ ﬁÌœ «·⁄„Ê·« ø") Then Exit Sub
    CommissionsDone frm, PostCommissionRun(ShownCommissionRun(frm)), " „  —ÕÌ· „”Ì— «·⁄„Ê·« ."
End Sub

Public Sub CommissionsUnpost(ByVal frm As Access.Form)
    If ShownCommissionRun(frm) = 0 Then Exit Sub
    If Not AskYesNo("≈·€«¡  —ÕÌ· «·„”Ì— ÊÕ–› ﬁÌœÂø") Then Exit Sub
    CommissionsDone frm, UnpostCommissionRun(ShownCommissionRun(frm)), "⁄«œ «·„”Ì— „”Êœ…."
End Sub

Public Sub CommissionsDelete(ByVal frm As Access.Form)
    Dim msg As String
    If ShownCommissionRun(frm) = 0 Then Exit Sub
    If Not AskYesNo("Õ–› „”Êœ… „”Ì— «·⁄„Ê·« ø") Then Exit Sub
    msg = DeleteCommissionRun(ShownCommissionRun(frm))
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        CommissionsShow frm, 0
    End If
End Sub

Public Sub PrintCommissionRun(ByVal frm As Access.Form)
    If ShownCommissionRun(frm) = 0 Then Exit Sub
    SetQueryParam "CommissionRunID", ShownCommissionRun(frm)
    LogAction "REPORT", "COMMISSION_RUN", CStr(ShownCommissionRun(frm))
    OpenReportOrQuery "rptCommissionRun", "CommissionSheetQuery", "", _
                      "‘Â— " & Format$(CommissionRunField(ShownCommissionRun(frm), "RunMonth"), "yyyy/mm")
End Sub

'------------------------------------------------------------------------------
' Screen frmSalesReps: the rep's numbers this month under the fields
'------------------------------------------------------------------------------
Public Sub SalesRepCurrent(ByVal frm As Access.Form)
    Dim rs As DAO.Recordset, info As String
    If frm.NewRecord Or IsNull(frm!SalesRepID.Value) Then
        frm!lblRepInfo.Caption = Tr(" ")
        Exit Sub
    End If
    SetPeriod DateSerial(Year(Date), Month(Date), 1), Date
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM RepPerformanceQuery WHERE SalesRepID = " & frm!SalesRepID.Value, _
                                     dbOpenSnapshot)
    If Not rs.EOF Then
        info = "Â–« «·‘Â—: ’«›Ì «·„»Ì⁄«  " & Format$(Nz(rs!NetSales, 0), "#,##0.00") & "   «· Õ’Ì· " & _
               Format$(Nz(rs!Collections, 0), "#,##0.00")
        If Not IsNull(rs!Achievement) Then info = info & "   «·≈‰Ã«“ " & Format$(rs!Achievement, "0%")
    End If
    rs.Close
    info = info & "   «·⁄„Ê·«  «·„” Õﬁ… " & Format$(RepPayable(frm!SalesRepID.Value), "#,##0.00") & _
           "   «·⁄„·«¡ " & Nz(DbValue("SELECT COUNT(*) FROM Customers WHERE SalesRepID = " & frm!SalesRepID.Value), 0)
    frm!lblRepInfo.Caption = Tr(Trim$(info))
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Private Sub CheckRep(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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

Public Function TestSalesReps() As Boolean
    Dim passed As Long, failed As Long, report As String, msg As String, ws As DAO.Workspace, inTrans As Boolean
    Dim rep As Long, cust As Long, payID As Long, runID As Long, box As Long, voucher As Long, rm As Date
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestSalesReps  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    rm = MonthEnd(Date)
    If ClosedThroughDate() >= Date Or Not IsNull(DbValue("SELECT CommissionRunID FROM CommissionRuns WHERE RunMonth = " & _
                                                          SqlDate(rm))) Then
        Debug.Print "[--] ··‘Â— «·Õ«·Ì „”Ì— ⁄„Ê·«  √Ê «·› —… „ﬁ›·…"
        GoTo Undo
    End If
    CheckRep CalcCommission(1000, 0.025, 0) = 25 And CalcCommission(1000, 0.025, -40) = 0 And _
             CalcCommission(-500, 0.1, 0) = 0, "Õ”«» «·⁄„Ê·… Ê·«  ﬁ· ⁄‰ ’›—", passed, failed, report
    CurrentDb.Execute "UPDATE SalesReps SET IsActive = False", dbFailOnError
    CurrentDb.Execute "INSERT INTO SalesReps (RepCode, RepName, CommissionRate, CommissionBase, IsActive) VALUES " & _
                      "('TST-REP', 'TEST „‰œÊ»', 0.1, 'COLLECTION', True)", dbFailOnError
    rep = DbValue("SELECT SalesRepID FROM SalesReps WHERE RepCode = 'TST-REP'")
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, OpeningBalance, CurrentBalance, AllowCredit, CreditLimit, IsActive, " & _
                      "SalesRepID) VALUES ('TEST ⁄„Ì· «·„‰œÊ»', 0, 1000, True, 0, True, " & rep & ")", dbFailOnError
    cust = DbValue("SELECT CustomerID FROM Customers WHERE CustomerName = 'TEST ⁄„Ì· «·„‰œÊ»'")
    CheckRep SalesRepFor(cust) = rep, "„‰œÊ» «·⁄„Ì·", passed, failed, report
    CheckRep IsNull(SalesRepFor(1)), "⁄„Ì· »·« „‰œÊ» Ê„” Œœ„ »·« „‰œÊ»", passed, failed, report
    CurrentDb.Execute "UPDATE SalesReps SET EmployeeID = " & CurrentUserID() & " WHERE SalesRepID = " & rep, dbFailOnError
    CheckRep SalesRepFor(1) = rep, "„‰œÊ» «·„” Œœ„ «·Õ«·Ì", passed, failed, report

    msg = PostCustomerPayment(cust, 800, CASH_METHOD_ID, "TEST", payID)
    CheckRep Len(msg) = 0 And Nz(DbValue("SELECT SalesRepID FROM CustomerPayments WHERE PaymentID = " & payID), 0) = rep, _
             "”‰œ «·ﬁ»÷ ÌÕ„· «·„‰œÊ» " & msg, passed, failed, report
    msg = CreateCommissionRun(Date, runID)
    CheckRep Len(msg) = 0 And runID > 0 And _
             Nz(DbValue("SELECT Commission FROM CommissionLines WHERE CommissionRunID = " & runID & " AND SalesRepID = " & rep), 0) = 80, _
             "„”Êœ… «·⁄„Ê·« : 10% „‰ «· Õ’Ì· " & msg, passed, failed, report
    CheckRep Len(CreateCommissionRun(Date, voucher)) > 0, "„”Ì— Ê«Õœ ··‘Â—", passed, failed, report
    CurrentDb.Execute "UPDATE CommissionLines SET Adjustment = 20 WHERE CommissionRunID = " & runID & " AND SalesRepID = " & rep, _
                      dbFailOnError
    msg = PostCommissionRun(runID)
    CheckRep Len(msg) = 0 And Nz(CommissionRunField(runID, "TotalAmount"), 0) >= 100 And RepPayable(rep) = 100 And _
             Not IsNull(DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType = 'COMMISSION' AND SourceID = " & runID)), _
             " —ÕÌ· «·⁄„Ê·« : «· ⁄œÌ· Ê«·„” Õﬁ Ê«·ﬁÌœ " & msg, passed, failed, report
    CheckRep Len(DeleteCommissionRun(runID)) > 0, "·« ÌıÕ–› „”Ì— „—ÕÛ¯·", passed, failed, report
    CheckRep Len(CommissionVoucherProblem("OUT", "COMMISSION", rep, 150)) > 0 And _
             Len(CommissionVoucherProblem("OUT", "COMMISSION", Null, 50)) > 0 And _
             Len(CommissionVoucherProblem("OUT", "COMMISSION", rep, 60)) = 0, "’—› «·⁄„Ê·… ·« Ì Ã«Ê“ «·„” Õﬁ", _
             passed, failed, report
    box = Nz(DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE IsActive = True"), 0)
    CurrentDb.Execute "INSERT INTO CashVouchers (VoucherNumber, VoucherDate, VoucherType, CashBoxID, Category, Amount, " & _
        "Description, EmployeeID, SalesRepID) VALUES ('TEST-COM', " & SqlDate(Date) & ", 'OUT', " & box & _
        ", 'COMMISSION', 60, 'TEST', " & CurrentUserID() & ", " & rep & ")", dbFailOnError
    CheckRep RepPayable(rep) = 40, "”‰œ «·’—› Ì‰ﬁ’ «·„” Õﬁ", passed, failed, report
    CheckRep Len(UnpostCommissionRun(runID)) > 0, "·« Ìı·€Ï  —ÕÌ· ⁄„Ê·«  ’ı—› „‰Â«", passed, failed, report
    CurrentDb.Execute "DELETE FROM CashVouchers WHERE VoucherNumber = 'TEST-COM'", dbFailOnError
    msg = UnpostCommissionRun(runID) & DeleteCommissionRun(runID)
    CheckRep Len(msg) = 0 And IsNull(DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType = 'COMMISSION' AND " & _
             "SourceID = " & runID)), "≈·€«¡ «· —ÕÌ· ÌÕ–› «·ﬁÌœ À„  ıÕ–› «·„”Êœ… " & msg, passed, failed, report
    SyncJournal
    CheckRep Nz(DbValue("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0) = 0, _
             "ﬂ· «·ﬁÌÊœ „ Ê«“‰…", passed, failed, report
Undo:
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckRep False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·„‰œÊ»Ì‰ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestSalesReps"
        TestSalesReps = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestSalesReps"
    End If
End Function
