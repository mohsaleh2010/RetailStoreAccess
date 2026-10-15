Attribute VB_Name = "modOpening"
'==============================================================================
' modOpening  -  Retail Store Management System
'
' The start-up opening balances (docs/52-Opening-Balances.md, screen frmOpeningBalances,
' permission OPENING_BALANCES):
'   Settings.OpeningDate  the start-up date: every opening balance is an entry of that day
'                         (customers and suppliers: qryCustomerOpening / qrySupplierOpening,
'                         the cash boxes and banks get it as their OpeningDate).
'                         Not after the first document, not in a closed period.
'   OpeningBalances       the general accounts (capital, loans, accruals...): each line is an entry
'                         ACCOUNT_OPENING against the opening balances account 3900.
'   Account 3900          "total opening balances": the other side of every opening balance. Its
'                         balance is zero when the balances entered are balanced; a difference stays
'                         there, visible on the screen by source (qryOpeningTotals).
' The customers, suppliers, cash boxes and banks keep their OpeningBalance field; the screen edits it
' in one grid per part, with the closed period check and the audit trail.
'==============================================================================
Option Compare Database
Option Explicit

Public Const OPENING_ACCOUNT As Long = 3900
Private Const OPENING_SOURCES As String = "'BOX_OPENING','CUSTOMER_OPENING','SUPPLIER_OPENING','BANK_OPENING','ACCOUNT_OPENING'"

'------------------------------------------------------------------------------
' Rules
'------------------------------------------------------------------------------
Public Function OpeningDateValue() As Variant
    ' The start-up date, Null when it was never set.
    Dim v As Variant
    v = SettingValue("OpeningDate")
    If IsDate(v) Then
        OpeningDateValue = DateValue(v)
    Else
        OpeningDateValue = Null
    End If
End Function

Public Function OpeningEntryDate(ByVal CreatedAt As Variant) As Variant
    ' The day of the opening balance of a customer or supplier: the start-up date, else the day of the card.
    OpeningEntryDate = OpeningDateValue()
    If IsNull(OpeningEntryDate) Then OpeningEntryDate = CreatedAt
End Function

Public Function CanOpening(ByVal Action As String) As String
    If Not HasPermission("OPENING_BALANCES") Then
        CanOpening = "لا تملك صلاحية أرصدة بداية التشغيل."
    ElseIf Not CanScreenAction("frmOpeningBalances", Action, True) Then
        CanOpening = "لا تملك صلاحية " & IIf(Action = "ADD", "الإضافة", IIf(Action = "EDIT", "التعديل", "الحذف")) & _
                     " في شاشة أرصدة بداية التشغيل."
    End If
End Function

Public Function FirstDocumentDate() As Variant
    ' The day of the first operation that is not an opening balance (from the journal), Null when none.
    FirstDocumentDate = DbValue("SELECT Min(EntryDate) FROM JournalEntries WHERE SourceType NOT IN (" & _
                                OPENING_SOURCES & ")")
End Function

Public Function OpeningDateProblem(ByVal NewDate As Variant) As String
    ' "" when the start-up date can be NewDate.
    Dim first As Variant, old As Variant
    If Not IsDate(NewDate) Then
        OpeningDateProblem = "أدخل تاريخ بدء التشغيل."
        Exit Function
    End If
    first = FirstDocumentDate()
    If IsDate(first) Then
        If DateValue(NewDate) > DateValue(first) Then
            OpeningDateProblem = "تاريخ بدء التشغيل لا يكون بعد أول مستند (" & GDate(first) & ")."
            Exit Function
        End If
    End If
    OpeningDateProblem = ClosedPeriodProblem(NewDate)
    If Len(OpeningDateProblem) > 0 Then Exit Function
    old = OpeningDateValue()
    If IsDate(old) Then OpeningDateProblem = ClosedPeriodProblem(old)       ' the entries of the old date move
End Function

Public Function OpeningAccountProblem(ByVal AccountCode As Variant) As String
    ' The general accounts only: the others have their opening balance in their own part.
    Dim code As Long
    If IsNull(AccountCode) Then
        OpeningAccountProblem = "اختر الحساب."
        Exit Function
    End If
    code = CLng(AccountCode)
    Select Case code
        Case 1300: OpeningAccountProblem = "رصيد العملاء يُدخل في قسم «العملاء»: رصيد لكل عميل."
        Case 2100: OpeningAccountProblem = "رصيد الموردين يُدخل في قسم «الموردون»: رصيد لكل مورد."
        Case 1400: OpeningAccountProblem = "رصيد المخزون يُدخل من شاشة المخزون: كمية وتكلفة لكل صنف."
        Case OPENING_ACCOUNT: OpeningAccountProblem = "حساب إجمالي الأرصدة الافتتاحية هو مقابل كل الأرصدة، ولا يُدخل له رصيد."
        Case 110001 To 119999: OpeningAccountProblem = "رصيد الصندوق يُدخل في قسم «الصناديق»."
        Case 120001 To 129999: OpeningAccountProblem = "رصيد البنك يُدخل في قسم «البنوك»."
        Case Else
            If Nz(DbValue("SELECT IsPosting FROM Accounts WHERE AccountCode = " & code), False) = False Then
                OpeningAccountProblem = "اختر حسابًا فرعيًا (حساب تُسجَّل عليه قيود)."
            End If
    End Select
End Function

Public Function SaveOpeningDate(ByVal NewDate As Variant) As String
    ' The start-up date; the cash boxes and banks that had the old date (or none) move with it.
    Dim old As Variant, before As Collection, cond As String
    SaveOpeningDate = CanOpening("EDIT")
    If Len(SaveOpeningDate) > 0 Then Exit Function
    SaveOpeningDate = SyncJournal()               ' the first document must be known
    If Len(SaveOpeningDate) > 0 Then Exit Function
    SaveOpeningDate = OpeningDateProblem(NewDate)
    If Len(SaveOpeningDate) > 0 Then Exit Function
    old = OpeningDateValue()
    Set before = AuditSnapshot("Settings", "SettingID", 1)
    CurrentDb.Execute "UPDATE Settings SET OpeningDate = " & SqlDate(DateValue(NewDate)) & " WHERE SettingID = 1", _
                      dbFailOnError
    cond = "OpeningDate Is Null"
    If IsDate(old) Then cond = cond & " OR OpeningDate = " & SqlDate(old)
    CurrentDb.Execute "UPDATE CashBoxes SET OpeningDate = " & SqlDate(DateValue(NewDate)) & " WHERE " & cond, dbFailOnError
    CurrentDb.Execute "UPDATE Banks SET OpeningDate = " & SqlDate(DateValue(NewDate)) & " WHERE " & cond, dbFailOnError
    AuditEdited "OPENING_DATE", "Settings", "SettingID", 1, before
    SaveOpeningDate = SyncJournal()
End Function

Public Function OpeningBalanceTotal() As Currency
    ' The balance of account 3900 (debit - credit); zero when the opening balances are balanced.
    OpeningBalanceTotal = Nz(DbValue("SELECT Sum(SumDebit - SumCredit) FROM qryOpeningTotals"), 0)
End Function

Public Function OpeningTotalText(ByVal Total As Currency) As String
    Dim acc As String
    acc = "حساب إجمالي الأرصدة الافتتاحية (" & OPENING_ACCOUNT & ")"
    If Total = 0 Then
        OpeningTotalText = acc & ": الرصيد صفر، الأرصدة الافتتاحية متوازنة."
    ElseIf Total > 0 Then
        OpeningTotalText = acc & ": رصيده " & Format$(Total, "#,##0.00") & " مدين." & vbCrLf & _
                           "الأرصدة الدائنة (رأس المال والالتزامات) أكبر من المدينة بهذا المبلغ."
    Else
        OpeningTotalText = acc & ": رصيده " & Format$(-Total, "#,##0.00") & " دائن." & vbCrLf & _
                           "الأرصدة المدينة (الأصول) أكبر من الدائنة بهذا المبلغ: أكمل رأس المال أو الالتزامات."
    End If
End Function

'------------------------------------------------------------------------------
' Screen frmOpeningBalances
'------------------------------------------------------------------------------
Public Sub OpeningLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    frm!txtOpeningDate.Value = OpeningDateValue()
    OpeningShowPart frm, "ACCOUNTS"
    OpeningRefresh frm
End Sub

Public Sub OpeningSaveDate(ByVal frm As Access.Form)
    Dim msg As String
    If Not IsDate(frm!txtOpeningDate.Value) Then
        ShowWarning "أدخل تاريخ بدء التشغيل."
        SafeFocus frm!txtOpeningDate
        Exit Sub
    End If
    If Not AskYesNo("كل الأرصدة الافتتاحية تصبح بتاريخ " & GDate(frm!txtOpeningDate.Value) & ". متابعة؟") Then Exit Sub
    msg = SaveOpeningDate(frm!txtOpeningDate.Value)
    If Len(msg) > 0 Then
        ShowWarning msg
        frm!txtOpeningDate.Value = OpeningDateValue()
        Exit Sub
    End If
    ShowInfo "تم حفظ تاريخ بدء التشغيل."
    OpeningShowPart frm, Nz(frm!txtPart.Value, "ACCOUNTS")
    OpeningRefresh frm
End Sub

Public Sub OpeningRefresh(ByVal frm As Access.Form)
    ' The balance of 3900 by source, after bringing the entries up to date.
    Dim msg As String, total As Currency
    msg = SyncJournal()
    If Len(msg) > 0 Then ShowWarning msg
    frm!lstSources.RowSource = Tr("SELECT TypeName AS [المصدر], Format(SumDebit, '#,##0.00') AS [مدين], " & _
                                  "Format(SumCredit, '#,##0.00') AS [دائن] FROM qryOpeningTotals ORDER BY SourceType")
    total = OpeningBalanceTotal()
    frm!lblTotal.Caption = Tr(OpeningTotalText(total))
    frm!lblTotal.ForeColor = IIf(total = 0, CLR_SUCCESS, CLR_DANGER)
End Sub

Public Sub OpeningShowPart(ByVal frm As Access.Form, ByVal Part As String)
    ' One grid per part: the general accounts, the customers, the suppliers, the cash boxes, the banks.
    Dim grid As String, caps As Variant, hint As String
    Select Case Part
        Case "CUSTOMERS"
            grid = "frmOpeningCustomers"
            caps = Array("العميل", "الرصيد الافتتاحي", "", "الرصيد الحالي")
            hint = "رصيد كل عميل يوم بدء التشغيل: موجب = العميل مدين للمحل (عليه)، سالب = دائن (له)."
        Case "SUPPLIERS"
            grid = "frmOpeningSuppliers"
            caps = Array("المورد", "الرصيد الافتتاحي", "", "الرصيد الحالي")
            hint = "رصيد كل مورد يوم بدء التشغيل: موجب = المحل مدين للمورد (له)، سالب = دائن (عليه)."
        Case "BOXES"
            grid = "frmOpeningBoxes"
            caps = Array("الصندوق", "الرصيد الافتتاحي", "", "تاريخ الرصيد")
            hint = "النقدية في كل صندوق يوم بدء التشغيل."
        Case "BANKS"
            grid = "frmOpeningBanks"
            caps = Array("البنك", "الرصيد الافتتاحي", "", "تاريخ الرصيد")
            hint = "رصيد كل حساب بنكي يوم بدء التشغيل: موجب = لنا في البنك، سالب = سحب على المكشوف."
        Case Else
            Part = "ACCOUNTS"
            grid = "frmOpeningAccountLines"
            caps = Array("الحساب", "مدين", "دائن", "البيان")
            hint = "رصيد كل حساب عام يوم بدء التشغيل: رأس المال، والقروض، والمستحقات، والأصول الأخرى. " & _
                   "العملاء والموردون والصناديق والبنوك في أقسامها، والمخزون من شاشة المخزون."
    End Select
    frm!txtPart.Value = Part
    If frm!subGrid.SourceObject <> grid Then frm!subGrid.SourceObject = grid
    frm!lblCol1.Caption = Tr(caps(0))
    frm!lblCol2.Caption = Tr(caps(1))
    frm!lblCol3.Caption = Tr(IIf(Len(caps(2)) = 0, " ", caps(2)))
    frm!lblCol4.Caption = Tr(caps(3))
    If IsNull(OpeningDateValue()) Then
        hint = "احفظ تاريخ بدء التشغيل أولًا، ثم أدخل الأرصدة."
        frm!subGrid.Enabled = False
    Else
        frm!subGrid.Enabled = True
    End If
    frm!lblHint.Caption = Tr(hint)
End Sub

'------------------------------------------------------------------------------
' Grids
'------------------------------------------------------------------------------
Public Sub OpeningLineBeforeUpdate(ByVal frm As Access.Form, ByRef Cancel As Integer)
    ' A line of the general accounts (OpeningBalances).
    Dim msg As String
    msg = CanOpening(IIf(frm.NewRecord, "ADD", "EDIT"))
    If Len(msg) = 0 And IsNull(OpeningDateValue()) Then msg = "احفظ تاريخ بدء التشغيل أولًا."
    If Len(msg) = 0 Then msg = ClosedPeriodProblem(OpeningDateValue())
    If Len(msg) = 0 Then msg = OpeningAccountProblem(frm!AccountCode.Value)
    If Len(msg) = 0 Then
        If Nz(frm!Debit.Value, 0) < 0 Or Nz(frm!Credit.Value, 0) < 0 Or _
           (Nz(frm!Debit.Value, 0) <> 0 And Nz(frm!Credit.Value, 0) <> 0) Or _
           Nz(frm!Debit.Value, 0) + Nz(frm!Credit.Value, 0) = 0 Then
            msg = "اكتب المبلغ في المدين أو في الدائن فقط."
        End If
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
        Cancel = True
        Exit Sub
    End If
    frm!Debit.Value = Nz(frm!Debit.Value, 0)
    frm!Credit.Value = Nz(frm!Credit.Value, 0)
    AuditFormBefore frm, "OpeningBalances", "OpeningLineID"
End Sub

Public Sub OpeningPartyBeforeUpdate(ByVal frm As Access.Form, ByRef Cancel As Integer, ByVal TableName As String, _
                                    ByVal KeyField As String)
    ' The opening balance of a customer, supplier, cash box or bank.
    Dim msg As String, oldDay As Variant, delta As Currency
    msg = CanOpening("EDIT")
    If Len(msg) = 0 And IsNull(OpeningDateValue()) Then msg = "احفظ تاريخ بدء التشغيل أولًا."
    If Len(msg) = 0 Then
        If TableName = "Customers" Or TableName = "Suppliers" Then
            oldDay = OpeningEntryDate(frm!CreatedAt.Value)
        Else
            oldDay = frm!OpeningDate.Value
        End If
        msg = ClosedPeriodProblem(oldDay)
        If Len(msg) = 0 Then msg = ClosedPeriodProblem(OpeningDateValue())
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
        Cancel = True
        Exit Sub
    End If
    frm!OpeningBalance.Value = Nz(frm!OpeningBalance.Value, 0)
    If TableName = "Customers" Or TableName = "Suppliers" Then
        ' the stored current balance follows (the reference stays the balance queries)
        delta = Nz(frm!OpeningBalance.Value, 0) - Nz(frm!OpeningBalance.OldValue, 0)
        frm!CurrentBalance.Value = Nz(frm!CurrentBalance.Value, 0) + delta
    ElseIf IsNull(frm!OpeningDate.Value) Then
        frm!OpeningDate.Value = OpeningDateValue()
    End If
    AuditFormBefore frm, TableName, KeyField
End Sub

Public Sub OpeningGridAfterUpdate(ByVal frm As Access.Form)
    AuditFormAfter frm
    On Error Resume Next                      ' the grid can be opened alone
    frm.Parent!lblTotal.Caption = Tr("تغيرت الأرصدة: اضغط «تحديث الإجمالي».")
    frm.Parent!lblTotal.ForeColor = CLR_MUTED
End Sub

Public Sub OpeningRemoveLine(ByVal frm As Access.Form)
    Dim id As Variant, msg As String, before As Collection
    If frm.NewRecord Then
        If frm.Dirty Then frm.Undo
        Exit Sub
    End If
    id = frm!OpeningLineID.Value
    If IsNull(id) Then Exit Sub
    msg = CanOpening("DELETE")
    If Len(msg) = 0 Then msg = ClosedPeriodProblem(OpeningDateValue())
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    If Not AskYesNo("حذف رصيد هذا الحساب؟") Then Exit Sub
    If frm.Dirty Then frm.Undo
    Set before = AuditSnapshot("OpeningBalances", "OpeningLineID", id)
    CurrentDb.Execute "DELETE FROM OpeningBalances WHERE OpeningLineID = " & CLng(id), dbFailOnError
    AuditDeleted "OPENING_DELETE", "OpeningBalances", id, before
    frm.Requery
    OpeningGridAfterUpdate frm
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Public Function TestOpeningBalances() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, passed As Long, failed As Long, report As String
    Dim first As Variant, startDay As Date, totalBefore As Currency, entry As Variant, msg As String
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestOpeningBalances  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Record OpeningAccountProblem(1300) <> "" And OpeningAccountProblem(OPENING_ACCOUNT) <> "" And _
           OpeningAccountProblem(110001) <> "" And OpeningAccountProblem(3200) = "", _
           "الحسابات المسموحة في الأرصدة العامة", passed, failed, report
    Record OpeningDateProblem(Null) <> "", "التاريخ مطلوب", passed, failed, report
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    msg = SyncJournal()
    first = FirstDocumentDate()
    If IsDate(first) Then
        Record OpeningDateProblem(DateAdd("d", 1, DateValue(first))) <> "", "لا يكون بعد أول مستند", passed, failed, report
        startDay = DateAdd("d", -1, DateValue(first))
    Else
        startDay = Date
    End If
    If Len(ClosedPeriodProblem(startDay)) > 0 Then
        Debug.Print "[--] يوم البدء في فترة مقفلة: تخطي القيد"
    Else
        Record SaveOpeningDate(startDay) = "" And OpeningDateValue() = startDay, "حفظ تاريخ بدء التشغيل", _
               passed, failed, report
        totalBefore = OpeningBalanceTotal()
        CurrentDb.Execute "INSERT INTO OpeningBalances (AccountCode, Debit, Credit, LineText) VALUES (3200, 0, 1000, 'TEST')", _
                          dbFailOnError
        msg = SyncJournal()
        entry = DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType = 'ACCOUNT_OPENING' AND SourceID = " & _
                        "(SELECT Max(OpeningLineID) FROM OpeningBalances)")
        Record Not IsNull(entry), "قيد الرصيد الافتتاحي للحساب", passed, failed, report
        If Not IsNull(entry) Then
            Record DateValue(DbValue("SELECT EntryDate FROM JournalEntries WHERE EntryID = " & entry)) = startDay And _
                   Nz(DbValue("SELECT Debit FROM JournalLines WHERE EntryID = " & entry & " AND AccountCode = 3900"), 0) = 1000, _
                   "بتاريخ بدء التشغيل ومقابله حساب إجمالي الأرصدة الافتتاحية", passed, failed, report
        End If
        Record OpeningBalanceTotal() = totalBefore + 1000, "إجمالي الأرصدة الافتتاحية يتغير بالرصيد", _
               passed, failed, report
        Record IsNull(DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType IN ('CUSTOMER_OPENING', " & _
                              "'SUPPLIER_OPENING') AND DateValue(EntryDate) <> " & SqlDate(startDay))), _
               "أرصدة العملاء والموردين بتاريخ بدء التشغيل", passed, failed, report
    End If
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
        TestMsg "جميع اختبارات أرصدة بداية التشغيل ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, _
                "TestOpeningBalances"
        TestOpeningBalances = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
                "TestOpeningBalances"
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
