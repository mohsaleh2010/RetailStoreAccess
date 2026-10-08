Attribute VB_Name = "modCheque"
'==============================================================================
' modCheque  -  Retail Store Management System
'
' Received and issued cheques (table Cheques, screen frmCheques, permission CHEQUES):
'   received (IN) from a customer: pays his balance at once and waits in 1250
'       "cheques under collection"; collected -> the bank, bounced -> back on the
'       customer's balance.
'   issued (OUT) to a supplier: pays his balance at once and waits in 2110
'       "notes payable"; paid by the bank -> the bank, bounced -> back to the supplier.
' Journal sources CHEQUE (the receipt / issue) and CHEQUE_STATUS (the collection /
' bounce); the customer and supplier ledgers and the aging show them too.
' CurrentBalance of the customer / supplier follows (modSales.AdjustBalance).
'==============================================================================
Option Compare Database
Option Explicit

Private Function ChequeField(ByVal ChequeID As Long, ByVal FieldName As String) As Variant
    ChequeField = DbValue("SELECT " & FieldName & " FROM Cheques WHERE ChequeID = " & ChequeID)
End Function

Private Function CanCheques(ByVal Action As String) As String
    If Not HasPermission("CHEQUES") Then
        CanCheques = "لا تملك صلاحية الشيكات."
    ElseIf Not CanScreenAction("frmCheques", Action, True) Then
        CanCheques = "لا تملك صلاحية " & IIf(Action = "ADD", "تسجيل الشيكات", IIf(Action = "EDIT", "تحصيل الشيكات وارتدادها", _
                     "حذف الشيكات")) & "."
    End If
End Function

Private Sub MoveBalance(ByVal Direction As String, ByVal PartyID As Long, ByVal Delta As Currency)
    ' Delta > 0: the party owes more (customer) / is owed more (supplier).
    If Direction = "IN" Then
        AdjustBalance CurrentDb, "Customers", "CustomerID", PartyID, Delta
    Else
        AdjustBalance CurrentDb, "Suppliers", "SupplierID", PartyID, Delta
    End If
End Sub

Private Function PartyOf(ByVal ChequeID As Long) As Long
    PartyOf = Nz(ChequeField(ChequeID, "IIf(Direction = 'IN', CustomerID, SupplierID)"), 0)
End Function

'------------------------------------------------------------------------------
' Receive / issue
'------------------------------------------------------------------------------
Public Function ChequeProblem(ByVal Direction As String, ByVal PartyID As Variant, ByVal ChequeNo As String, _
                              ByVal BankID As Variant, ByVal IssueDate As Variant, ByVal DueDate As Variant, _
                              ByVal Amount As Currency) As String
    Dim p As String
    If Direction <> "IN" And Direction <> "OUT" Then
        p = "اختر نوع الشيك: وارد أو صادر."
    ElseIf IsNull(PartyID) Then
        p = IIf(Direction = "IN", "اختر العميل.", "اختر المورد.")
    ElseIf Direction = "IN" And CLng(PartyID) = Nz(SettingValue("DefaultCustomerID"), 1) Then
        p = "لا يُسجَّل شيك على العميل النقدي."
    ElseIf Len(Trim$(ChequeNo)) = 0 Then
        p = "اكتب رقم الشيك."
    ElseIf Amount <= 0 Then
        p = "المبلغ يجب أن يكون أكبر من صفر."
    ElseIf Not IsDate(IssueDate) Or Not IsDate(DueDate) Then
        p = "اكتب تاريخ " & IIf(Direction = "IN", "الاستلام", "الإصدار") & " وتاريخ الاستحقاق."
    ElseIf DateValue(IssueDate) > Date Then
        p = "تاريخ " & IIf(Direction = "IN", "الاستلام", "الإصدار") & " بعد اليوم."
    ElseIf DateValue(DueDate) < DateValue(IssueDate) Then
        p = "تاريخ الاستحقاق قبل تاريخ " & IIf(Direction = "IN", "الاستلام", "الإصدار") & "."
    ElseIf Direction = "OUT" And IsNull(BankID) Then
        p = "اختر البنك المسحوب عليه الشيك."
    ElseIf Not IsNull(BankID) Then
        If Nz(DbValue("SELECT COUNT(*) FROM Banks WHERE IsActive = True AND BankID = " & CLng(BankID)), 0) = 0 Then
            p = "البنك غير نشط."
        End If
    End If
    If Len(p) = 0 And Direction = "IN" Then
        If Nz(DbValue("SELECT COUNT(*) FROM Cheques WHERE Direction = 'IN' AND CustomerID = " & CLng(PartyID) & _
                      " AND ChequeNo = " & SqlText(Trim$(ChequeNo))), 0) > 0 Then p = "هذا الشيك مسجل من قبل لنفس العميل."
    End If
    If Len(p) = 0 Then p = ClosedPeriodProblem(IssueDate)
    ChequeProblem = p
End Function

Public Function PostCheque(ByVal Direction As String, ByVal PartyID As Variant, ByVal ChequeNo As String, _
                           ByVal DrawerBank As String, ByVal BankID As Variant, ByVal IssueDate As Variant, _
                           ByVal DueDate As Variant, ByVal Amount As Currency, ByVal Notes As String, _
                           ByRef NewID As Long) As String
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean, ref As String
    On Error GoTo EH
    NewID = 0
    PostCheque = CanCheques("ADD")
    If Len(PostCheque) = 0 Then PostCheque = ChequeProblem(Direction, PartyID, ChequeNo, BankID, IssueDate, DueDate, Amount)
    If Len(PostCheque) > 0 Then Exit Function
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    ref = NextNumber("CHEQUE")
    Set rs = db.OpenRecordset("Cheques", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!ChequeRef = ref
    rs!Direction = Direction
    If Direction = "IN" Then rs!CustomerID = CLng(PartyID) Else rs!SupplierID = CLng(PartyID)
    rs!ChequeNo = Left$(Trim$(ChequeNo), 30)
    If Len(Trim$(DrawerBank)) > 0 Then rs!DrawerBank = Left$(Trim$(DrawerBank), 100)
    If Not IsNull(BankID) Then rs!BankID = CLng(BankID)
    rs!IssueDate = DateValue(IssueDate)
    rs!DueDate = DateValue(DueDate)
    rs!Amount = Amount
    rs!Status = "PENDING"
    If Len(Trim$(Notes)) > 0 Then rs!Notes = Left$(Trim$(Notes), 255)
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Bookmark = rs.LastModified
    NewID = rs!ChequeID
    rs.Close
    MoveBalance Direction, CLng(PartyID), -Amount                  ' the cheque pays the balance
    ws.CommitTrans
    inTrans = False
    LogAction "CHEQUE_" & Direction, "Cheques", ref, ChequeNo & " " & Format$(Amount, "0.00")
    PostCheque = SyncJournal()
    Exit Function
EH:
    PostCheque = "تعذر حفظ الشيك: " & Err.Description
    NewID = 0
    If inTrans Then ws.Rollback
End Function

Public Function DeleteCheque(ByVal ChequeID As Long) As String
    ' A cheque still under collection, recorded in an open period.
    Dim direction As Variant, amount As Currency
    DeleteCheque = CanCheques("DELETE")
    If Len(DeleteCheque) > 0 Then Exit Function
    direction = ChequeField(ChequeID, "Direction")
    If IsNull(direction) Then
        DeleteCheque = "الشيك غير موجود."
        Exit Function
    End If
    If ChequeField(ChequeID, "Status") <> "PENDING" Then
        DeleteCheque = "يُحذف الشيك تحت التحصيل فقط: ألغِ تحصيله أو ارتداده أولًا."
        Exit Function
    End If
    DeleteCheque = ClosedPeriodProblem(ChequeField(ChequeID, "IssueDate"))
    If Len(DeleteCheque) > 0 Then Exit Function
    amount = ChequeField(ChequeID, "Amount")
    MoveBalance direction, PartyOf(ChequeID), amount
    Dim auditBefore As Collection
    Set auditBefore = AuditSnapshot("Cheques", "ChequeID", ChequeID)        ' the record as it was (modAudit)
    CurrentDb.Execute "DELETE FROM Cheques WHERE ChequeID = " & ChequeID, dbFailOnError
    AuditDeleted "CHEQUE_DELETE", "Cheques", ChequeID, auditBefore
    DeleteCheque = SyncJournal()
End Function

'------------------------------------------------------------------------------
' Collect / bounce / undo
'------------------------------------------------------------------------------
Public Function SetChequeStatus(ByVal ChequeID As Long, ByVal NewStatus As String, ByVal StatusDate As Variant, _
                                ByVal BankID As Variant) As String
    ' NewStatus COLLECTED (the bank collected / paid it) or BOUNCED.
    Dim direction As Variant, amount As Currency, due As Variant
    SetChequeStatus = CanCheques("EDIT")
    If Len(SetChequeStatus) > 0 Then Exit Function
    direction = ChequeField(ChequeID, "Direction")
    If IsNull(direction) Then
        SetChequeStatus = "اختر الشيك."
        Exit Function
    End If
    If ChequeField(ChequeID, "Status") <> "PENDING" Then
        SetChequeStatus = "الشيك ليس تحت التحصيل."
        Exit Function
    End If
    If NewStatus <> "COLLECTED" And NewStatus <> "BOUNCED" Then Exit Function
    If Not IsDate(StatusDate) Then
        SetChequeStatus = "اكتب تاريخ " & IIf(NewStatus = "BOUNCED", "الارتداد.", IIf(direction = "IN", "التحصيل.", "الصرف."))
        Exit Function
    End If
    due = ChequeField(ChequeID, "DueDate")
    If DateValue(StatusDate) > Date Then
        SetChequeStatus = "التاريخ بعد اليوم."
    ElseIf DateValue(StatusDate) < DateValue(ChequeField(ChequeID, "IssueDate")) Then
        SetChequeStatus = "التاريخ قبل تاريخ الشيك."
    ElseIf NewStatus = "COLLECTED" And DateValue(StatusDate) < DateValue(due) Then
        SetChequeStatus = "الشيك مستحق في " & GDate(due) & ": لا يُحصَّل قبل استحقاقه."
    ElseIf NewStatus = "COLLECTED" And IsNull(BankID) Then
        SetChequeStatus = "اختر البنك."
    End If
    If Len(SetChequeStatus) = 0 And NewStatus = "COLLECTED" Then
        If Nz(DbValue("SELECT COUNT(*) FROM Banks WHERE IsActive = True AND BankID = " & CLng(BankID)), 0) = 0 Then
            SetChequeStatus = "البنك غير نشط."
        End If
    End If
    If Len(SetChequeStatus) = 0 Then SetChequeStatus = ClosedPeriodProblem(StatusDate)
    If Len(SetChequeStatus) > 0 Then Exit Function
    amount = ChequeField(ChequeID, "Amount")
    CurrentDb.Execute "UPDATE Cheques SET Status = " & SqlText(NewStatus) & ", StatusDate = " & SqlDate(DateValue(StatusDate)) & _
                      IIf(NewStatus = "COLLECTED", ", BankID = " & CLng(Nz(BankID, 0)), "") & _
                      " WHERE ChequeID = " & ChequeID, dbFailOnError
    If NewStatus = "BOUNCED" Then MoveBalance direction, PartyOf(ChequeID), amount     ' owed again
    LogAction "CHEQUE_" & NewStatus, "Cheques", CStr(ChequeID)
    SetChequeStatus = SyncJournal()
End Function

Public Function UndoChequeStatus(ByVal ChequeID As Long) As String
    ' Back under collection (a mistake): not when its bank movement was reconciled.
    Dim status As String
    UndoChequeStatus = CanCheques("EDIT")
    If Len(UndoChequeStatus) > 0 Then Exit Function
    status = Nz(ChequeField(ChequeID, "Status"), "")
    If status <> "COLLECTED" And status <> "BOUNCED" Then
        UndoChequeStatus = "الشيك تحت التحصيل."
        Exit Function
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM BankClearings WHERE SourceType = 'CHEQUE_STATUS' AND SourceID = " & ChequeID), 0) > 0 Then
        UndoChequeStatus = "حركة الشيك مطابقة في تسوية بنكية: ألغِ مطابقتها أولًا."
        Exit Function
    End If
    UndoChequeStatus = ClosedPeriodProblem(ChequeField(ChequeID, "StatusDate"))
    If Len(UndoChequeStatus) > 0 Then Exit Function
    If status = "BOUNCED" Then MoveBalance ChequeField(ChequeID, "Direction"), PartyOf(ChequeID), -ChequeField(ChequeID, "Amount")
    CurrentDb.Execute "UPDATE Cheques SET Status = 'PENDING', StatusDate = Null WHERE ChequeID = " & ChequeID, dbFailOnError
    LogAction "CHEQUE_UNDO", "Cheques", CStr(ChequeID)
    UndoChequeStatus = SyncJournal()
End Function

'------------------------------------------------------------------------------
' Screen frmCheques (OpenArgs "IN" / "OUT")
'------------------------------------------------------------------------------
Public Sub ChequesLoad(ByVal frm As Access.Form)
    Dim msg As String
    Calendar = vbCalGreg
    msg = SyncJournal()
    If Len(msg) > 0 Then ShowWarning msg
    frm!cboDirection.Value = IIf(Nz(frm.OpenArgs, "IN") = "OUT", "OUT", "IN")
    frm!cboShow.Value = "PENDING"
    frm!txtIssueDate.Value = Date
    frm!txtDueDate.Value = Date
    frm!txtActionDate.Value = Date
    frm!cboActionBank.Value = SettingValue("DefaultBankID")
    ChequesDirectionChanged frm
End Sub

Public Sub ChequesDirectionChanged(ByVal frm As Access.Form)
    If frm!cboDirection.Value = "IN" Then
        frm!cboParty.RowSource = Tr("SELECT c.CustomerID, c.CustomerName FROM [@Customers] AS c WHERE c.IsActive = True AND " & _
                                 "c.CustomerID <> " & Nz(SettingValue("DefaultCustomerID"), 1) & " ORDER BY c.CustomerName")
        frm!lblParty.Caption = Tr("العميل")
        frm!btnCollect.Caption = Tr("تحصيل في البنك")
    Else
        frm!cboParty.RowSource = Tr("SELECT s.SupplierID, s.SupplierName FROM [@Suppliers] AS s WHERE s.IsActive = True ORDER BY s.SupplierName")
        frm!lblParty.Caption = Tr("المورد")
        frm!btnCollect.Caption = Tr("صرفه البنك")
    End If
    frm!cboParty.Value = Null
    frm!txtDrawerBank.Enabled = (frm!cboDirection.Value = "IN")
    ChequesRefresh frm
End Sub

Public Sub ChequesRefresh(ByVal frm As Access.Form)
    Dim where As String, kind As String, pending As Currency, soon As Currency, late As Currency
    kind = Nz(frm!cboDirection.Value, "IN")
    where = "Direction = " & SqlText(kind)
    Select Case Nz(frm!cboShow.Value, "PENDING")
        Case "PENDING": where = where & " AND Status = 'PENDING'"
        Case "DUE": where = where & " AND Status = 'PENDING' AND DueDate <= " & SqlDate(Date + 7)
        Case "COLLECTED": where = where & " AND Status = 'COLLECTED'"
        Case "BOUNCED": where = where & " AND Status = 'BOUNCED'"
    End Select
    frm!lstCheques.RowSource = Tr("SELECT ChequeID, ChequeRef AS [القيد], ChequeNo AS [رقم الشيك], PartyName AS [" & _
        IIf(kind = "IN", "العميل", "المورد") & "], Format(DueDate, 'yyyy/mm/dd') AS [الاستحقاق], Format(q.Amount, '#,##0.00') " & _
        "AS [مبلغ الشيك], StatusName AS [الحالة], Format(StatusDate, 'yyyy/mm/dd') AS [في], Nz(BankName, DrawerBank) AS [البنك] " & _
        "FROM ChequesQuery AS q WHERE " & where & " ORDER BY DueDate, ChequeID")
    pending = Nz(DbValue("SELECT Sum(Amount) FROM Cheques WHERE Status = 'PENDING' AND Direction = " & SqlText(kind)), 0)
    soon = Nz(DbValue("SELECT Sum(Amount) FROM Cheques WHERE Status = 'PENDING' AND Direction = " & SqlText(kind) & _
                      " AND DueDate <= " & SqlDate(Date + 7)), 0)
    late = Nz(DbValue("SELECT Sum(Amount) FROM Cheques WHERE Status = 'PENDING' AND Direction = " & SqlText(kind) & _
                      " AND DueDate < " & SqlDate(Date)), 0)
    frm!lblTotals.Caption = Tr(IIf(kind = "IN", "شيكات تحت التحصيل: ", "شيكات صادرة لم تُصرف: ") & Format$(pending, "#,##0.00") & _
        "    مستحقة خلال 7 أيام: " & Format$(soon, "#,##0.00") & "    فات استحقاقها: " & Format$(late, "#,##0.00"))
    frm!lblTotals.ForeColor = IIf(late > 0, CLR_DANGER, CLR_PRIMARY)
End Sub

Public Sub SaveCheque(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    msg = PostCheque(Nz(frm!cboDirection.Value, ""), frm!cboParty.Value, Nz(frm!txtChequeNo.Value, ""), _
                     Nz(frm!txtDrawerBank.Value, ""), frm!cboBank.Value, frm!txtIssueDate.Value, frm!txtDueDate.Value, _
                     CCur(Nz(frm!txtAmount.Value, 0)), Nz(frm!txtNotes.Value, ""), id)
    If id = 0 Then
        ShowWarning msg
        Exit Sub
    End If
    If Len(msg) > 0 Then ShowWarning msg
    ShowInfo "تم تسجيل الشيك " & ChequeField(id, "ChequeRef") & "."
    frm!txtChequeNo.Value = Null
    frm!txtAmount.Value = Null
    frm!txtNotes.Value = Null
    ChequesRefresh frm
End Sub

Private Function PickedCheque(ByVal frm As Access.Form) As Long
    PickedCheque = Nz(frm!lstCheques.Value, 0)
    If PickedCheque = 0 Then ShowWarning "اختر الشيك من القائمة."
End Function

Private Sub ChequeDone(ByVal frm As Access.Form, ByVal Msg As String, ByVal Success As String)
    If Len(Msg) > 0 Then
        ShowWarning Msg
    Else
        ShowInfo Success
    End If
    ChequesRefresh frm
End Sub

Public Sub CollectCheque(ByVal frm As Access.Form)
    Dim id As Long, msg As String
    id = PickedCheque(frm)
    If id = 0 Then Exit Sub
    msg = SetChequeStatus(id, "COLLECTED", frm!txtActionDate.Value, frm!cboActionBank.Value)
    If ChequeField(id, "Status") = "COLLECTED" Then
        ChequeDone frm, "", IIf(frm!cboDirection.Value = "IN", "تم تحصيل الشيك في البنك.", "تم تسجيل صرف الشيك.")
    Else
        ChequeDone frm, msg, ""
    End If
End Sub

Public Sub BounceCheque(ByVal frm As Access.Form)
    Dim id As Long, msg As String
    id = PickedCheque(frm)
    If id = 0 Then Exit Sub
    If Not AskYesNo("تسجيل ارتداد الشيك؟ يعود مبلغه على " & IIf(frm!cboDirection.Value = "IN", "العميل.", "المورد.")) Then Exit Sub
    msg = SetChequeStatus(id, "BOUNCED", frm!txtActionDate.Value, Null)
    If ChequeField(id, "Status") = "BOUNCED" Then
        ChequeDone frm, "", "تم تسجيل ارتداد الشيك."
    Else
        ChequeDone frm, msg, ""
    End If
End Sub

Public Sub UndoCheque(ByVal frm As Access.Form)
    Dim id As Long, msg As String
    id = PickedCheque(frm)
    If id = 0 Then Exit Sub
    If Not AskYesNo("إرجاع الشيك إلى «تحت التحصيل» وحذف قيد تحصيله أو ارتداده؟") Then Exit Sub
    msg = UndoChequeStatus(id)
    If ChequeField(id, "Status") = "PENDING" Then
        ChequeDone frm, "", "عاد الشيك تحت التحصيل."
    Else
        ChequeDone frm, msg, ""
    End If
End Sub

Public Sub DeleteSelectedCheque(ByVal frm As Access.Form)
    Dim id As Long, msg As String
    id = PickedCheque(frm)
    If id = 0 Then Exit Sub
    If Not AskYesNo("حذف الشيك المحدد وقيده؟") Then Exit Sub
    msg = DeleteCheque(id)
    If IsNull(ChequeField(id, "ChequeID")) Then
        ChequeDone frm, "", "تم حذف الشيك."
    Else
        ChequeDone frm, msg, ""
    End If
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Private Sub CheckCheque(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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

Public Function TestCheques() As Boolean
    Dim passed As Long, failed As Long, report As String, msg As String, ws As DAO.Workspace, inTrans As Boolean
    Dim cust As Long, supp As Long, bank As Long, id As Long, id2 As Long, before As Currency, sBefore As Currency
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestCheques  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    If ClosedThroughDate() >= Date Then
        Debug.Print "[--] الفترة مقفلة حتى اليوم"
        GoTo Undo
    End If
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, AllowCredit, CreditLimit, Notes) VALUES ('TEST عميل شيكات', True, 0, " & _
                      "'TEST')", dbFailOnError
    cust = DbValue("SELECT CustomerID FROM Customers WHERE CustomerName = 'TEST عميل شيكات'")
    CurrentDb.Execute "INSERT INTO Suppliers (SupplierName, Notes) VALUES ('TEST مورد شيكات', 'TEST')", dbFailOnError
    supp = DbValue("SELECT SupplierID FROM Suppliers WHERE SupplierName = 'TEST مورد شيكات'")
    CurrentDb.Execute "INSERT INTO Banks (BankName, OpeningBalance, OpeningDate, IsActive) VALUES ('TEST بنك الشيكات', 0, " & _
                      SqlDate(Date) & ", True)", dbFailOnError
    bank = DbValue("SELECT BankID FROM Banks WHERE BankName = 'TEST بنك الشيكات'")
    before = AccountBalance(1250)

    msg = PostCheque("IN", cust, "TEST-1001", "بنك الساحب", Null, Date, Date, 500, "", id)
    CheckCheque Len(msg) = 0 And id > 0 And Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerID = " & cust), 0) = -500 _
                And AccountBalance(1250) = before + 500, "شيك وارد يسدد رصيد العميل ويبقى تحت التحصيل " & msg, passed, failed, report
    CheckCheque Len(PostCheque("IN", cust, "TEST-1001", "", Null, Date, Date, 500, "", id2)) > 0, _
                "لا يتكرر الشيك نفسه للعميل", passed, failed, report
    msg = SetChequeStatus(id, "COLLECTED", Date, bank)
    CheckCheque Len(msg) = 0 And AccountBalance(1250) = before And BankBookBalance(bank) = 500, _
                "تحصيل الشيك: من تحت التحصيل إلى البنك " & msg, passed, failed, report
    msg = UndoChequeStatus(id)
    msg = msg & SetChequeStatus(id, "BOUNCED", Date, Null)
    CheckCheque Len(msg) = 0 And BankBookBalance(bank) = 0 And AccountBalance(1250) = before And _
                Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerID = " & cust), 0) = 0, _
                "الشيك المرتد يعود على العميل " & msg, passed, failed, report

    sBefore = AccountBalance(2110)
    msg = PostCheque("OUT", supp, "TEST-2001", "", bank, Date, Date + 30, 300, "", id2)
    CheckCheque Len(msg) = 0 And Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierID = " & supp), 0) = -300 And _
                AccountBalance(2110) = sBefore - 300, "شيك صادر يسدد رصيد المورد ويبقى في أوراق الدفع " & msg, _
                passed, failed, report
    CheckCheque Len(SetChequeStatus(id2, "COLLECTED", Date, bank)) > 0, "لا يُصرف الشيك قبل استحقاقه", passed, failed, report
    msg = DeleteCheque(id2)
    CheckCheque Len(msg) = 0 And Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierID = " & supp), 0) = 0 And _
                AccountBalance(2110) = sBefore, "حذف الشيك تحت التحصيل يعيد الرصيد " & msg, passed, failed, report
    CheckCheque Nz(DbValue("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0) = 0 And _
                AccountBalance(1300) = Nz(DbValue("SELECT Sum(CurrentBalance) FROM Customers"), 0) And _
                -AccountBalance(2100) = Nz(DbValue("SELECT Sum(CurrentBalance) FROM Suppliers"), 0), _
                "القيود متوازنة، وحسابا العملاء والموردين = أرصدتهم", passed, failed, report
Undo:
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckCheque False, "خطأ: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات الشيكات ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestCheques"
        TestCheques = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestCheques"
    End If
End Function
