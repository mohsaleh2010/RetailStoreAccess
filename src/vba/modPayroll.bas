Attribute VB_Name = "modPayroll"
'==============================================================================
' modPayroll  -  Retail Store Management System
'
' Monthly payroll (PayrollRuns + PayrollLines, screen frmPayroll, permission PAYROLL):
'   CreatePayroll   a draft for a month: a line for each active employee "on payroll"
'                   with his salary, allowances and the advance installment to deduct.
'   the draft is edited (overtime, additions, absence, advance, penalties); each line:
'     GOSI wage  = basic + housing, at most Settings.GosiMaxWage
'     employee   = Saudi: wage x GosiEmployeeRate; others: 0
'     employer   = wage x GosiEmployerRate (Saudi) or GosiNonSaudiRate (others)
'     net        = gross - absence - advance - penalties - employee GOSI
'   PostPayroll     the entry on the last day of the month (journal source PAYROLL):
'                   5500 salaries + 5510 allowances + 5520 employer GOSI against
'                   2320 GOSI payable + 1600 advances + 4200 penalties + 2310 net payable.
'   PayPayroll      the payment from a bank or a cash box (PAYROLL_PAYMENT): 2310 against it.
' Advances: cash vouchers of category ADVANCE carry the employee (AdvanceEmployeeID);
' AdvanceBalance = given - paid back in cash - deducted in posted payrolls.
' Same amounts as tools/sim.py (Store.payroll_line).
'==============================================================================
Option Compare Database
Option Explicit

Private Function RunField(ByVal RunID As Long, ByVal FieldName As String) As Variant
    RunField = DbValue("SELECT " & FieldName & " FROM PayrollRuns WHERE PayrollRunID = " & RunID)
End Function

Private Function CanPayroll(ByVal Action As String) As String
    If Not HasPermission("PAYROLL") Then
        CanPayroll = "لا تملك صلاحية مسير الرواتب."
    ElseIf Not CanScreenAction("frmPayroll", Action, True) Then
        CanPayroll = "لا تملك هذه الصلاحية في شاشة مسير الرواتب."
    End If
End Function

'------------------------------------------------------------------------------
' Amounts
'------------------------------------------------------------------------------
Public Function CalcPayLine(ByVal Basic As Currency, ByVal Housing As Currency, ByVal OtherAllow As Currency, _
                            ByVal Overtime As Currency, ByVal Additions As Currency, ByVal Absence As Currency, _
                            ByVal Advance As Currency, ByVal OtherDed As Currency, ByVal IsSaudi As Boolean, _
                            ByRef Wage As Currency, ByRef GosiEE As Currency, ByRef GosiER As Currency) As Currency
    ' Returns the net salary; Wage / GosiEE / GosiER are filled.
    Dim maxWage As Currency
    maxWage = Nz(SettingValue("GosiMaxWage"), 45000)
    Wage = Basic + Housing
    If maxWage > 0 And Wage > maxWage Then Wage = maxWage
    If IsSaudi Then
        GosiEE = Round(Wage * Nz(SettingValue("GosiEmployeeRate"), 0.0975), 2)
        GosiER = Round(Wage * Nz(SettingValue("GosiEmployerRate"), 0.1175), 2)
    Else
        GosiEE = 0
        GosiER = Round(Wage * Nz(SettingValue("GosiNonSaudiRate"), 0.02), 2)
    End If
    CalcPayLine = Basic + Housing + OtherAllow + Overtime + Additions - Absence - Advance - OtherDed - GosiEE
End Function

Public Function AdvanceBalance(ByVal EmployeeID As Long) As Currency
    ' Advances not paid back yet (posted payrolls only).
    AdvanceBalance = Nz(DbValue("SELECT AdvanceBalance FROM qryAdvanceTotals WHERE EmployeeID = " & EmployeeID), 0)
End Function

Private Sub SaveLineAmounts(ByVal rs As DAO.Recordset)
    ' Recomputes GOSI and the net of the current record of rs (inside its Edit / AddNew).
    Dim wage As Currency, ee As Currency, er As Currency
    rs!NetPay = CalcPayLine(rs!Basic, rs!Housing, rs!OtherAllow, rs!Overtime, rs!Additions, rs!AbsenceDeduction, _
                            rs!AdvanceDeduction, rs!OtherDeduction, rs!IsSaudi, wage, ee, er)
    rs!GosiWage = wage
    rs!GosiEmployee = ee
    rs!GosiEmployer = er
End Sub

Private Sub AddLines(ByVal RunID As Long)
    ' A line for every active employee on payroll, with the advance installment.
    Dim db As DAO.Database, e As DAO.Recordset, l As DAO.Recordset, advance As Currency, room As Currency
    Dim wage As Currency, ee As Currency, er As Currency
    Set db = CurrentDb
    Set e = db.OpenRecordset("SELECT * FROM Employees WHERE OnPayroll = True AND IsActive = True AND BasicSalary > 0 " & _
                             "ORDER BY EmployeeName", dbOpenSnapshot)
    Set l = db.OpenRecordset("PayrollLines", dbOpenDynaset, dbAppendOnly)
    Do Until e.EOF
        l.AddNew
        l!PayrollRunID = RunID
        l!EmployeeID = e!EmployeeID
        l!EmployeeName = Left$(e!EmployeeName, 100)
        l!IsSaudi = Nz(e!IsSaudi, False)
        l!Basic = Nz(e!BasicSalary, 0)
        l!Housing = Nz(e!HousingAllowance, 0)
        l!OtherAllow = Nz(e!TransportAllowance, 0) + Nz(e!OtherAllowance, 0)
        l!Overtime = 0: l!Additions = 0: l!AbsenceDeduction = 0: l!OtherDeduction = 0
        ' the advance: the installment (or all of it), not more than the salary leaves
        advance = AdvanceBalance(e!EmployeeID)
        If Nz(e!AdvanceInstallment, 0) > 0 And advance > e!AdvanceInstallment Then advance = e!AdvanceInstallment
        room = CalcPayLine(l!Basic, l!Housing, l!OtherAllow, 0, 0, 0, 0, 0, l!IsSaudi, wage, ee, er)
        If advance > room Then advance = room
        If advance < 0 Then advance = 0
        l!AdvanceDeduction = advance
        SaveLineAmounts l
        l.Update
        e.MoveNext
    Loop
    e.Close
    l.Close
End Sub

'------------------------------------------------------------------------------
' Draft, posting, payment
'------------------------------------------------------------------------------
Public Function CreatePayroll(ByVal AnyDayOfMonth As Date, ByRef RunID As Long) As String
    Dim m As Date, rs As DAO.Recordset
    RunID = 0
    CreatePayroll = CanPayroll("ADD")
    If Len(CreatePayroll) > 0 Then Exit Function
    m = MonthEnd(AnyDayOfMonth)                                      ' modAssets
    If Not IsNull(DbValue("SELECT PayrollRunID FROM PayrollRuns WHERE PayMonth = " & SqlDate(m))) Then
        CreatePayroll = "يوجد مسير لشهر " & Format$(m, "yyyy/mm") & " بالفعل."
        Exit Function
    End If
    If DateSerial(Year(m), Month(m), 1) > Date Then
        CreatePayroll = "لا يُنشأ مسير لشهر لم يبدأ."
        Exit Function
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM Employees WHERE OnPayroll = True AND IsActive = True AND BasicSalary > 0"), 0) = 0 Then
        CreatePayroll = "لا يوجد موظفون في مسير الرواتب: حدّد «في مسير الرواتب» والراتب الأساسي في شاشة المستخدمين."
        Exit Function
    End If
    CreatePayroll = ClosedPeriodProblem(m)
    If Len(CreatePayroll) > 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("PayrollRuns", dbOpenDynaset)
    rs.AddNew
    rs!RunNumber = NextNumber("PAYROLL")
    rs!PayMonth = m
    rs!Status = "DRAFT"
    rs!PaidAmount = 0
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Bookmark = rs.LastModified
    RunID = rs!PayrollRunID
    rs.Close
    AddLines RunID
    LogAction "PAYROLL_CREATE", "PayrollRuns", Format$(m, "yyyy-mm")
End Function

Public Function RebuildPayroll(ByVal RunID As Long) As String
    ' The draft again from the employees' data (the changes made in it are lost).
    RebuildPayroll = CanPayroll("EDIT")
    If Len(RebuildPayroll) = 0 And Nz(RunField(RunID, "Status"), "") <> "DRAFT" Then RebuildPayroll = "المسير ليس مسودة."
    If Len(RebuildPayroll) > 0 Then Exit Function
    CurrentDb.Execute "DELETE FROM PayrollLines WHERE PayrollRunID = " & RunID, dbFailOnError
    AddLines RunID
End Function

Public Function RecalcPayroll(ByVal RunID As Long) As String
    ' GOSI and net of every line of a draft (after the lines were edited).
    Dim rs As DAO.Recordset
    If Nz(RunField(RunID, "Status"), "") <> "DRAFT" Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM PayrollLines WHERE PayrollRunID = " & RunID, dbOpenDynaset)
    Do Until rs.EOF
        rs.Edit
        SaveLineAmounts rs
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
End Function

Public Function PayrollProblem(ByVal RunID As Long) As String
    ' The checks of a draft before posting ("" = fine).
    Dim rs As DAO.Recordset, p As String
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM PayrollLines WHERE PayrollRunID = " & RunID & " ORDER BY EmployeeName", _
                                     dbOpenSnapshot)
    If rs.EOF Then p = "المسير بلا أسطر."
    Do Until rs.EOF Or Len(p) > 0
        If rs!Overtime < 0 Or rs!Additions < 0 Or rs!AbsenceDeduction < 0 Or rs!AdvanceDeduction < 0 Or rs!OtherDeduction < 0 Then
            p = rs!EmployeeName & ": المبالغ لا تكون سالبة."
        ElseIf rs!AbsenceDeduction > rs!Basic + rs!Housing Then
            p = rs!EmployeeName & ": خصم الغياب أكبر من الأساسي والسكن."
        ElseIf rs!AdvanceDeduction > AdvanceBalance(rs!EmployeeID) Then
            p = rs!EmployeeName & ": خصم السلفة أكبر من رصيد سلفه (" & Format$(AdvanceBalance(rs!EmployeeID), "#,##0.00") & ")."
        ElseIf rs!NetPay < 0 Then
            p = rs!EmployeeName & ": الخصومات أكبر من الراتب."
        End If
        rs.MoveNext
    Loop
    rs.Close
    PayrollProblem = p
End Function

Public Function PostPayroll(ByVal RunID As Long) As String
    Dim pm As Variant
    PostPayroll = CanPayroll("EDIT")
    If Len(PostPayroll) > 0 Then Exit Function
    If Nz(RunField(RunID, "Status"), "") <> "DRAFT" Then
        PostPayroll = "المسير غير موجود أو مرحَّل."
        Exit Function
    End If
    pm = RunField(RunID, "PayMonth")
    RecalcPayroll RunID
    PostPayroll = PayrollProblem(RunID)
    If Len(PostPayroll) = 0 Then PostPayroll = ClosedPeriodProblem(pm)
    If Len(PostPayroll) > 0 Then Exit Function
    CurrentDb.Execute "UPDATE PayrollRuns SET Status = 'POSTED', EmployeeID = " & CurrentUserID() & _
                      " WHERE PayrollRunID = " & RunID, dbFailOnError
    LogAction "PAYROLL_POST", "PayrollRuns", Format$(pm, "yyyy-mm")
    SyncJournal
End Function

Public Function UnpostPayroll(ByVal RunID As Long) As String
    UnpostPayroll = CanPayroll("EDIT")
    If Len(UnpostPayroll) > 0 Then Exit Function
    If Nz(RunField(RunID, "Status"), "") <> "POSTED" Then
        UnpostPayroll = "المسير غير مرحَّل."
    ElseIf Nz(RunField(RunID, "PaidAmount"), 0) <> 0 Then
        UnpostPayroll = "المسير مصروف: ألغِ الصرف أولًا."
    ElseIf Not IsNull(DbValue("SELECT TOP 1 PayrollRunID FROM PayrollRuns WHERE Status = 'POSTED' AND PayMonth > " & _
                              SqlDate(RunField(RunID, "PayMonth")))) Then
        UnpostPayroll = "يوجد مسير مرحَّل بعده: خصومات السلف فيه تعتمد على هذا المسير."
    Else
        UnpostPayroll = ClosedPeriodProblem(RunField(RunID, "PayMonth"))
    End If
    If Len(UnpostPayroll) > 0 Then Exit Function
    CurrentDb.Execute "UPDATE PayrollRuns SET Status = 'DRAFT' WHERE PayrollRunID = " & RunID, dbFailOnError
    LogAction "PAYROLL_UNPOST", "PayrollRuns", CStr(RunID)
    SyncJournal
End Function

Public Function DeletePayroll(ByVal RunID As Long) As String
    DeletePayroll = CanPayroll("DELETE")
    If Len(DeletePayroll) = 0 And Nz(RunField(RunID, "Status"), "") <> "DRAFT" Then
        DeletePayroll = "تُحذف المسودة فقط: ألغِ الترحيل أولًا."
    End If
    If Len(DeletePayroll) > 0 Then Exit Function
    CurrentDb.Execute "DELETE FROM PayrollRuns WHERE PayrollRunID = " & RunID, dbFailOnError       ' lines cascade
    LogAction "PAYROLL_DELETE", "PayrollRuns", CStr(RunID)
End Function

Public Function PayPayroll(ByVal RunID As Long, ByVal PaidDate As Variant, ByVal PaidFrom As String, _
                           ByVal BankID As Variant, ByVal CashBoxID As Variant) As String
    Dim net As Currency, pm As Date
    PayPayroll = CanPayroll("EDIT")
    If Len(PayPayroll) > 0 Then Exit Function
    If Nz(RunField(RunID, "Status"), "") <> "POSTED" Then
        PayPayroll = "رحّل المسير أولًا، ثم سجّل صرفه."
        Exit Function
    End If
    If Nz(RunField(RunID, "PaidAmount"), 0) <> 0 Then
        PayPayroll = "المسير مصروف بالفعل."
        Exit Function
    End If
    net = Nz(DbValue("SELECT SumNet FROM qryPayrollTotals WHERE PayrollRunID = " & RunID), 0)
    pm = RunField(RunID, "PayMonth")
    If Not IsDate(PaidDate) Then
        PayPayroll = "اكتب تاريخ الصرف."
    ElseIf DateValue(PaidDate) > Date Then
        PayPayroll = "تاريخ الصرف بعد اليوم."
    ElseIf DateValue(PaidDate) < DateSerial(Year(pm), Month(pm), 1) Then
        PayPayroll = "تاريخ الصرف قبل شهر المسير."
    ElseIf PaidFrom = "BANK" Then
        If Nz(DbValue("SELECT COUNT(*) FROM Banks WHERE IsActive = True AND BankID = " & CLng(Nz(BankID, 0))), 0) = 0 Then
            PayPayroll = "اختر البنك."
        End If
    ElseIf PaidFrom = "CASHBOX" Then
        If IsNull(CashBoxID) Then
            PayPayroll = "اختر الصندوق."
        ElseIf CashBoxBalance(CLng(CashBoxID), DateValue(PaidDate) + 1) < net Then
            PayPayroll = "رصيد الصندوق في " & GDate(PaidDate) & " أقل من صافي الرواتب (" & Format$(net, "#,##0.00") & ")."
        End If
    Else
        PayPayroll = "اختر الصرف من بنك أو صندوق."
    End If
    If Len(PayPayroll) = 0 Then PayPayroll = ClosedPeriodProblem(PaidDate)
    If Len(PayPayroll) > 0 Then Exit Function
    CurrentDb.Execute "UPDATE PayrollRuns SET PaidDate = " & SqlDate(DateValue(PaidDate)) & ", PaidFrom = " & SqlText(PaidFrom) & _
        ", BankID = " & IIf(PaidFrom = "BANK", CLng(Nz(BankID, 0)), "Null") & ", CashBoxID = " & _
        IIf(PaidFrom = "CASHBOX", CLng(Nz(CashBoxID, 0)), "Null") & ", PaidAmount = " & Str$(net) & _
        " WHERE PayrollRunID = " & RunID, dbFailOnError
    LogAction "PAYROLL_PAY", "PayrollRuns", CStr(RunID), Format$(net, "0.00")
    SyncJournal
End Function

Public Function UndoPayrollPayment(ByVal RunID As Long) As String
    UndoPayrollPayment = CanPayroll("EDIT")
    If Len(UndoPayrollPayment) > 0 Then Exit Function
    If Nz(RunField(RunID, "PaidAmount"), 0) = 0 Then
        UndoPayrollPayment = "المسير غير مصروف."
        Exit Function
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM BankClearings WHERE SourceType = 'PAYROLL_PAYMENT' AND SourceID = " & RunID), 0) > 0 Then
        UndoPayrollPayment = "الصرف مطابق في تسوية بنكية: ألغِ مطابقته أولًا."
        Exit Function
    End If
    UndoPayrollPayment = ClosedPeriodProblem(RunField(RunID, "PaidDate"))
    If Len(UndoPayrollPayment) > 0 Then Exit Function
    CurrentDb.Execute "UPDATE PayrollRuns SET PaidDate = Null, PaidFrom = Null, BankID = Null, CashBoxID = Null, " & _
                      "PaidAmount = 0 WHERE PayrollRunID = " & RunID, dbFailOnError
    LogAction "PAYROLL_UNPAY", "PayrollRuns", CStr(RunID)
    SyncJournal
End Function

'------------------------------------------------------------------------------
' Screen frmPayroll (subform frmPayrollLines on PayrollLines)
'------------------------------------------------------------------------------
Private Function ShownRun(ByVal frm As Access.Form) As Long
    ShownRun = Nz(frm!txtRunID.Value, 0)
End Function

Public Sub PayrollLoad(ByVal frm As Access.Form)
    Dim last As Variant
    Calendar = vbCalGreg
    frm!txtMonth.Value = DateSerial(Year(Date), Month(Date), 1)
    frm!txtPaidDate.Value = Date
    frm!cboPaidFrom.Value = "BANK"
    frm!cboBank.Value = SettingValue("DefaultBankID")
    frm!txtGosiEE.Value = SettingValue("GosiEmployeeRate")
    frm!txtGosiER.Value = SettingValue("GosiEmployerRate")
    frm!txtGosiNonSaudi.Value = SettingValue("GosiNonSaudiRate")
    frm!txtGosiMax.Value = SettingValue("GosiMaxWage")
    last = DbValue("SELECT TOP 1 PayrollRunID FROM PayrollRuns ORDER BY PayMonth DESC")
    If Not IsNull(last) Then
        PayrollShow frm, CLng(last)
    Else
        PayrollShow frm, 0
    End If
End Sub

Public Sub PayrollShow(ByVal frm As Access.Form, ByVal RunID As Long)
    Dim status As String, draft As Boolean
    frm!txtRunID.Value = IIf(RunID = 0, Null, RunID)
    frm!cboRun.Requery
    frm!cboRun.Value = IIf(RunID = 0, Null, RunID)
    frm!subLines.Form.RecordSource = "SELECT * FROM PayrollLines WHERE PayrollRunID = " & RunID & " ORDER BY EmployeeName"
    status = Nz(RunField(RunID, "Status"), "")
    draft = (status = "DRAFT")
    frm!subLines.Form.AllowEdits = draft
    frm!btnRebuild.Enabled = draft
    frm!btnPost.Enabled = draft
    frm!btnDelete.Enabled = draft
    frm!btnUnpost.Enabled = (status = "POSTED")
    frm!btnPay.Enabled = (status = "POSTED")
    frm!btnUndoPay.Enabled = (status = "POSTED")
    PayrollTotals frm
End Sub

Public Sub PayrollTotals(ByVal frm As Access.Form)
    Dim id As Long, rs As DAO.Recordset, state As String
    id = ShownRun(frm)
    If id = 0 Then
        frm!lblState.Caption = "اختر الشهر ثم «إنشاء مسير الشهر»."
        frm!lblTotals.Caption = " "
        Exit Sub
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM qryPayrollTotals WHERE PayrollRunID = " & id, dbOpenSnapshot)
    If Not rs.EOF Then
        frm!lblTotals.Caption = rs!LineCount & " موظف   الرواتب " & Format$(rs!SumSalaries, "#,##0.00") & "   البدلات " & _
            Format$(rs!SumAllowances, "#,##0.00") & "   خصم السلف " & Format$(rs!SumAdvance, "#,##0.00") & "   التأمينات " & _
            Format$(rs!SumGosi, "#,##0.00") & "   صافي الرواتب " & Format$(rs!SumNet, "#,##0.00")
    Else
        frm!lblTotals.Caption = "المسير بلا أسطر."
    End If
    rs.Close
    Select Case Nz(RunField(id, "Status"), "")
        Case "DRAFT"
            state = "مسودة مسير " & Format$(RunField(id, "PayMonth"), "yyyy/mm") & ": عدّل الإضافي والخصومات ثم «ترحيل المسير»."
        Case "POSTED"
            state = "مسير " & Format$(RunField(id, "PayMonth"), "yyyy/mm") & " مرحَّل"
            If Nz(RunField(id, "PaidAmount"), 0) <> 0 Then
                state = state & " ومصروف في " & GDate(RunField(id, "PaidDate"))
            Else
                state = state & "، لم يُصرف بعد"
            End If
    End Select
    frm!lblState.Caption = state
End Sub

Public Sub PayrollLineChanged(ByVal lines As Access.Form)
    ' The subform: GOSI and net of the line as its amounts are typed.
    Dim wage As Currency, ee As Currency, er As Currency
    If lines.NewRecord Then Exit Sub
    lines!NetPay.Value = CalcPayLine(Nz(lines!Basic.Value, 0), Nz(lines!Housing.Value, 0), Nz(lines!OtherAllow.Value, 0), _
                                   Nz(lines!Overtime.Value, 0), Nz(lines!Additions.Value, 0), Nz(lines!AbsenceDeduction.Value, 0), _
                                   Nz(lines!AdvanceDeduction.Value, 0), Nz(lines!OtherDeduction.Value, 0), _
                                   Nz(lines!IsSaudi.Value, False), wage, ee, er)
    lines!GosiWage.Value = wage
    lines!GosiEmployee.Value = ee
    lines!GosiEmployer.Value = er
    If lines.Dirty Then lines.Dirty = False
    PayrollTotals lines.Parent
End Sub

Public Sub PayrollPick(ByVal frm As Access.Form)
    If Not IsNull(frm!cboRun.Value) Then PayrollShow frm, CLng(frm!cboRun.Value)
End Sub

Public Sub PayrollCreate(ByVal frm As Access.Form)
    Dim id As Long, msg As String
    If Not IsDate(frm!txtMonth.Value) Then
        ShowWarning "اختر الشهر."
        Exit Sub
    End If
    msg = CreatePayroll(DateValue(frm!txtMonth.Value), id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    PayrollShow frm, id
End Sub

Private Sub PayrollDone(ByVal frm As Access.Form, ByVal Msg As String, ByVal Success As String)
    If Len(Msg) > 0 Then
        ShowWarning Msg
    Else
        ShowInfo Success
    End If
    PayrollShow frm, ShownRun(frm)
End Sub

Public Sub PayrollRebuild(ByVal frm As Access.Form)
    If ShownRun(frm) = 0 Then Exit Sub
    If Not AskYesNo("إعادة إنشاء أسطر المسير من بيانات الموظفين؟ تضيع التعديلات المكتوبة فيه.") Then Exit Sub
    PayrollDone frm, RebuildPayroll(ShownRun(frm)), "تمت إعادة إنشاء الأسطر."
End Sub

Public Sub PayrollPost(ByVal frm As Access.Form)
    If ShownRun(frm) = 0 Then Exit Sub
    If Not AskYesNo("ترحيل المسير وإنشاء قيد الرواتب؟") Then Exit Sub
    PayrollDone frm, PostPayroll(ShownRun(frm)), "تم ترحيل المسير."
End Sub

Public Sub PayrollUnpost(ByVal frm As Access.Form)
    If ShownRun(frm) = 0 Then Exit Sub
    If Not AskYesNo("إلغاء ترحيل المسير وحذف قيده؟") Then Exit Sub
    PayrollDone frm, UnpostPayroll(ShownRun(frm)), "عاد المسير مسودة."
End Sub

Public Sub PayrollDelete(ByVal frm As Access.Form)
    Dim msg As String
    If ShownRun(frm) = 0 Then Exit Sub
    If Not AskYesNo("حذف مسودة المسير؟") Then Exit Sub
    msg = DeletePayroll(ShownRun(frm))
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        PayrollShow frm, 0
    End If
End Sub

Public Sub PayrollPay(ByVal frm As Access.Form)
    If ShownRun(frm) = 0 Then Exit Sub
    PayrollDone frm, PayPayroll(ShownRun(frm), frm!txtPaidDate.Value, Nz(frm!cboPaidFrom.Value, ""), frm!cboBank.Value, _
                                frm!cboBox.Value), "تم تسجيل صرف الرواتب."
End Sub

Public Sub PayrollUndoPay(ByVal frm As Access.Form)
    If ShownRun(frm) = 0 Then Exit Sub
    If Not AskYesNo("إلغاء صرف الرواتب وحذف قيده؟") Then Exit Sub
    PayrollDone frm, UndoPayrollPayment(ShownRun(frm)), "تم إلغاء الصرف."
End Sub

Public Sub SaveGosiRates(ByVal frm As Access.Form)
    ' The GOSI rates of the coming payrolls (a posted payroll keeps its amounts).
    Dim v As Variant
    If Len(CanPayroll("EDIT")) > 0 Then
        ShowWarning CanPayroll("EDIT")
        Exit Sub
    End If
    For Each v In Array(frm!txtGosiEE.Value, frm!txtGosiER.Value, frm!txtGosiNonSaudi.Value)
        If Not IsNumeric(v) Then
            ShowWarning "اكتب النسب الثلاث."
            Exit Sub
        End If
        If v < 0 Or v >= 1 Then
            ShowWarning "النسبة بين 0% و 100%."
            Exit Sub
        End If
    Next
    If Not IsNumeric(frm!txtGosiMax.Value) Or Nz(frm!txtGosiMax.Value, 0) < 0 Then
        ShowWarning "اكتب الحد الأعلى للأجر الخاضع (0 = بلا حد)."
        Exit Sub
    End If
    CurrentDb.Execute "UPDATE Settings SET GosiEmployeeRate = " & Str$(frm!txtGosiEE.Value) & ", GosiEmployerRate = " & _
                      Str$(frm!txtGosiER.Value) & ", GosiNonSaudiRate = " & Str$(frm!txtGosiNonSaudi.Value) & _
                      ", GosiMaxWage = " & Str$(frm!txtGosiMax.Value) & " WHERE SettingID = 1", dbFailOnError
    LogAction "GOSI_RATES", "Settings"
    ShowInfo "تم حفظ نسب التأمينات. تُطبَّق على المسودات عند تعديلها أو ترحيلها."
End Sub

Public Sub PrintPayroll(ByVal frm As Access.Form)
    If ShownRun(frm) = 0 Then Exit Sub
    SetQueryParam "PayrollRunID", ShownRun(frm)
    LogAction "REPORT", "PAYROLL", CStr(ShownRun(frm))
    OpenReportOrQuery "rptPayroll", "PayrollSheetQuery", "", "شهر " & Format$(RunField(ShownRun(frm), "PayMonth"), "yyyy/mm")
End Sub

'------------------------------------------------------------------------------
' Cash voucher of an employee advance (modCash.SaveCashVoucher)
'------------------------------------------------------------------------------
Public Function AdvanceEmployeeProblem(ByVal VoucherType As String, ByVal Category As String, _
                                       ByVal EmployeeID As Variant, ByVal Amount As Currency) As String
    If Category <> "ADVANCE" Then Exit Function
    If IsNull(EmployeeID) Then
        AdvanceEmployeeProblem = "اختر الموظف صاحب السلفة."
    ElseIf VoucherType = "IN" And Amount > AdvanceBalance(CLng(EmployeeID)) Then
        AdvanceEmployeeProblem = "المبلغ أكبر من رصيد سلف الموظف (" & Format$(AdvanceBalance(CLng(EmployeeID)), "#,##0.00") & ")."
    End If
End Function

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Private Sub CheckPay(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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

Public Function TestPayroll() As Boolean
    Dim passed As Long, failed As Long, report As String, msg As String, ws As DAO.Workspace, inTrans As Boolean
    Dim emp As Long, runID As Long, box As Long, voucher As Long, wage As Currency, ee As Currency, er As Currency
    Dim net As Currency, pm As Date, adv As Currency
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestPayroll  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    pm = DateSerial(Year(Date), Month(Date), 0)                      ' last month
    If ClosedThroughDate() >= pm Or Not IsNull(DbValue("SELECT PayrollRunID FROM PayrollRuns WHERE PayMonth = " & _
                                                          SqlDate(pm))) Then
        Debug.Print "[--] للشهر الماضي مسير أو الفترة مقفلة"
        GoTo Undo
    End If
    net = CalcPayLine(5000, 1250, 500, 0, 0, 0, 0, 0, True, wage, ee, er)
    CheckPay wage = 6250 And ee = 609.38 And er = 734.38 And net = 6140.62, "التأمينات للسعودي وصافي الراتب", _
             passed, failed, report
    net = CalcPayLine(50000, 0, 0, 0, 0, 0, 0, 0, False, wage, ee, er)
    CheckPay wage = Nz(SettingValue("GosiMaxWage"), 45000) And ee = 0, "غير السعودي والحد الأعلى للأجر", passed, failed, report

    CurrentDb.Execute "UPDATE Employees SET OnPayroll = False", dbFailOnError
    emp = CurrentUserID()
    CurrentDb.Execute "UPDATE Employees SET OnPayroll = True, BasicSalary = 4000, HousingAllowance = 1000, " & _
                      "TransportAllowance = 400, OtherAllowance = 0, IsSaudi = True, AdvanceInstallment = 500 " & _
                      "WHERE EmployeeID = " & emp, dbFailOnError
    box = Nz(DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE IsActive = True"), 0)
    adv = AdvanceBalance(emp)
    CurrentDb.Execute "INSERT INTO CashVouchers (VoucherNumber, VoucherDate, VoucherType, CashBoxID, Category, Amount, " & _
        "Description, EmployeeID, AdvanceEmployeeID) VALUES ('TEST-ADV', " & SqlDate(pm - 10) & ", 'OUT', " & box & _
        ", 'ADVANCE', 1200, 'TEST', " & emp & ", " & emp & ")", dbFailOnError
    CheckPay AdvanceBalance(emp) = adv + 1200, "سلفة الموظف من سند الصرف", passed, failed, report
    msg = CreatePayroll(pm, runID)
    CheckPay Len(msg) = 0 And runID > 0 And Nz(DbValue("SELECT COUNT(*) FROM PayrollLines WHERE PayrollRunID = " & runID), 0) = 1 And _
             Nz(DbValue("SELECT AdvanceDeduction FROM PayrollLines WHERE PayrollRunID = " & runID), 0) = 500, _
             "مسودة المسير بقسط السلفة " & msg, passed, failed, report
    CurrentDb.Execute "UPDATE PayrollLines SET AdvanceDeduction = 999999 WHERE PayrollRunID = " & runID, dbFailOnError
    RecalcPayroll runID
    CheckPay Len(PostPayroll(runID)) > 0, "لا يُخصم أكثر من رصيد السلف", passed, failed, report
    CurrentDb.Execute "UPDATE PayrollLines SET AdvanceDeduction = 500, Overtime = 100 WHERE PayrollRunID = " & runID, dbFailOnError
    msg = PostPayroll(runID)
    net = Nz(DbValue("SELECT SumNet FROM qryPayrollTotals WHERE PayrollRunID = " & runID), 0)
    CheckPay Len(msg) = 0 And net = 4000 + 1000 + 400 + 100 - 500 - 487.5 And AdvanceBalance(emp) = adv + 700 And _
             Not IsNull(DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType = 'PAYROLL' AND SourceID = " & runID)), _
             "ترحيل المسير: الصافي وخصم السلفة والقيد " & msg, passed, failed, report
    CheckPay Len(DeletePayroll(runID)) > 0, "لا يُحذف مسير مرحَّل", passed, failed, report
    CurrentDb.Execute "INSERT INTO CashVouchers (VoucherNumber, VoucherDate, VoucherType, CashBoxID, Category, Amount, " & _
        "Description, EmployeeID) VALUES ('TEST-CIN', " & SqlDate(pm) & ", 'IN', " & box & ", 'OTHER', " & Str$(net) & _
        ", 'TEST', " & emp & ")", dbFailOnError
    msg = PayPayroll(runID, Date, "CASHBOX", Null, box)
    CheckPay Len(msg) = 0 And Not IsNull(DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType = 'PAYROLL_PAYMENT' " & _
             "AND SourceID = " & runID)), "صرف الرواتب من الصندوق " & msg, passed, failed, report
    CheckPay Len(UnpostPayroll(runID)) > 0, "لا يُلغى ترحيل مسير مصروف", passed, failed, report
    msg = UndoPayrollPayment(runID) & UnpostPayroll(runID)
    CheckPay Len(msg) = 0 And AdvanceBalance(emp) = adv + 1200 And _
             IsNull(DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType LIKE 'PAYROLL*' AND SourceID = " & runID)), _
             "إلغاء الصرف والترحيل يحذف القيدين " & msg, passed, failed, report
    CheckPay Nz(DbValue("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0) = 0, _
             "كل القيود متوازنة", passed, failed, report
Undo:
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckPay False, "خطأ: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات الرواتب ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestPayroll"
        TestPayroll = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestPayroll"
    End If
End Function
