Attribute VB_Name = "modManualEntry"
'==============================================================================
' modManualEntry  -  Retail Store Management System
'
' Manual journal entries (accruals, adjustments, capital, opening balances...):
'   tables ManualEntries + ManualEntryLines, screen frmManualEntry (lines in the
'   local table tmpManualLines while typing).
' A saved manual entry is an operation like the others: SyncJournal turns it into
' a journal entry (SourceType "MANUAL", source query qryJournalManual); editing it
' rebuilds that entry under the same JV number, deleting it removes the entry.
'   ManualEntryProblem   rules of the lines ("" = the entry can be saved)
'   PostManualEntry      saves a new or changed entry from tmpManualLines
'   RemoveManualEntry    deletes an entry (its lines go with it)
' Permission: MANUAL_ENTRY (frmUserScreens: add / edit / delete).
'==============================================================================
Option Compare Database
Option Explicit

'------------------------------------------------------------------------------
' Rules and saving
'------------------------------------------------------------------------------
Public Function ManualEntryProblem(ByVal EntryDate As Variant, ByVal Description As String) As String
    ' Checks the header and the lines typed in tmpManualLines.
    Dim rs As DAO.Recordset, n As Long, debit As Currency, credit As Currency, acc As Variant
    If Not IsDate(EntryDate) Then
        ManualEntryProblem = "اكتب تاريخ القيد."
        Exit Function
    End If
    If Len(Trim$(Description)) = 0 Then
        ManualEntryProblem = "اكتب بيان القيد."
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT LineNo, AccountCode, Debit, Credit FROM tmpManualLines ORDER BY LineNo", _
                                     dbOpenSnapshot)
    Do Until rs.EOF
        If IsNull(rs!AccountCode) And Nz(rs!Debit, 0) = 0 And Nz(rs!Credit, 0) = 0 Then
            ' an empty line is ignored
        ElseIf IsNull(rs!AccountCode) Then
            ManualEntryProblem = "اختر الحساب في كل سطر فيه مبلغ."
        ElseIf Nz(rs!Debit, 0) < 0 Or Nz(rs!Credit, 0) < 0 Then
            ManualEntryProblem = "المبالغ لا تكون سالبة: اكتب المبلغ في الجانب الآخر."
        ElseIf Nz(rs!Debit, 0) > 0 And Nz(rs!Credit, 0) > 0 Then
            ManualEntryProblem = "السطر يكون مدينًا أو دائنًا، وليس الاثنين."
        ElseIf Nz(rs!Debit, 0) = 0 And Nz(rs!Credit, 0) = 0 Then
            ManualEntryProblem = "اكتب المبلغ مدينًا أو دائنًا في كل سطر فيه حساب."
        Else
            acc = rs!AccountCode
            If IsNull(DbValue("SELECT AccountCode FROM Accounts WHERE AccountCode = " & acc & " AND IsPosting = True " & _
                              "AND IsActive = True")) Then
                ManualEntryProblem = "الحساب " & acc & " حساب رئيسي أو معطّل: القيود على الحسابات الفرعية النشطة فقط."
            Else
                n = n + 1
                debit = debit + Nz(rs!Debit, 0)
                credit = credit + Nz(rs!Credit, 0)
            End If
        End If
        If Len(ManualEntryProblem) > 0 Then
            rs.Close
            Exit Function
        End If
        rs.MoveNext
    Loop
    rs.Close
    If n < 2 Then
        ManualEntryProblem = "القيد سطران على الأقل: طرف مدين وطرف دائن."
    ElseIf debit <> credit Then
        ManualEntryProblem = "القيد غير متوازن: المدين " & Format$(debit, "#,##0.00") & " والدائن " & _
                             Format$(credit, "#,##0.00") & " (الفرق " & Format$(Abs(debit - credit), "#,##0.00") & ")."
    End If
End Function

Public Function PostManualEntry(ByVal ManualEntryID As Long, ByVal EntryDate As Variant, ByVal Description As String, _
                                ByVal Reference As String, ByVal ReversalOf As Variant, ByRef NewID As Long) As String
    ' "" on success; NewID = the saved entry. ManualEntryID = 0: a new entry.
    Dim ws As DAO.Workspace, db As DAO.Database, h As DAO.Recordset, rs As DAO.Recordset, inTrans As Boolean
    Dim n As Long, total As Currency
    On Error GoTo EH
    NewID = 0
    If Not HasPermission("MANUAL_ENTRY") Then
        PostManualEntry = "لا تملك صلاحية القيود اليدوية."
        Exit Function
    End If
    PostManualEntry = ManualEntryProblem(EntryDate, Description)
    If Len(PostManualEntry) = 0 Then PostManualEntry = ClosedPeriodProblem(EntryDate)          ' modClosing
    If Len(PostManualEntry) = 0 And ManualEntryID > 0 Then
        PostManualEntry = ClosedPeriodProblem(DbValue("SELECT EntryDate FROM ManualEntries WHERE ManualEntryID = " & _
                                                      ManualEntryID))
    End If
    If Len(PostManualEntry) > 0 Then Exit Function
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    If ManualEntryID = 0 Then
        Set h = db.OpenRecordset("ManualEntries", dbOpenDynaset)
        h.AddNew
        h!EntryNumber = NextNumber("MANUAL_ENTRY")
        h!EmployeeID = CurrentUserID()
        If Not IsNull(ReversalOf) Then h!ReversalOfID = ReversalOf
    Else
        Set h = db.OpenRecordset("SELECT * FROM ManualEntries WHERE ManualEntryID = " & ManualEntryID, dbOpenDynaset)
        If h.EOF Then
            h.Close
            ws.Rollback
            PostManualEntry = "القيد غير موجود."
            Exit Function
        End If
        h.Edit
        h!UpdatedAt = Now
        db.Execute "DELETE FROM ManualEntryLines WHERE ManualEntryID = " & ManualEntryID, dbFailOnError
    End If
    h!EntryDate = DateValue(EntryDate)
    h!Description = Left$(Trim$(Description), 255)
    h!Reference = IIf(Len(Trim$(Reference)) = 0, Null, Left$(Trim$(Reference), 50))
    h.Update
    h.Bookmark = h.LastModified
    NewID = h!ManualEntryID
    h.Close

    Set rs = db.OpenRecordset("SELECT AccountCode, Debit, Credit, LineText FROM tmpManualLines " & _
                              "WHERE AccountCode Is Not Null ORDER BY LineNo", dbOpenSnapshot)
    Set h = db.OpenRecordset("ManualEntryLines", dbOpenDynaset, dbAppendOnly)
    Do Until rs.EOF
        n = n + 1
        h.AddNew
        h!ManualEntryID = NewID
        h!LineNumber = n
        h!AccountCode = rs!AccountCode
        h!Debit = Nz(rs!Debit, 0)
        h!Credit = Nz(rs!Credit, 0)
        If Len(Trim$(Nz(rs!LineText, ""))) > 0 Then h!LineText = Left$(Trim$(rs!LineText), 150)
        h.Update
        total = total + Nz(rs!Debit, 0)
        rs.MoveNext
    Loop
    rs.Close
    h.Close
    db.Execute "UPDATE ManualEntries SET TotalAmount = " & Str$(total) & " WHERE ManualEntryID = " & NewID, dbFailOnError
    ws.CommitTrans
    inTrans = False
    LogAction IIf(ManualEntryID = 0, "MANUAL_ENTRY_ADD", "MANUAL_ENTRY_EDIT"), "ManualEntries", CStr(NewID)
    Exit Function
EH:
    PostManualEntry = "تعذر حفظ القيد: " & Err.Description
    NewID = 0
    If inTrans Then ws.Rollback
End Function

Public Function RemoveManualEntry(ByVal ManualEntryID As Long) As String
    If Not HasPermission("MANUAL_ENTRY") Then
        RemoveManualEntry = "لا تملك صلاحية القيود اليدوية."
        Exit Function
    End If
    If IsNull(DbValue("SELECT ManualEntryID FROM ManualEntries WHERE ManualEntryID = " & ManualEntryID)) Then
        RemoveManualEntry = "القيد غير موجود."
        Exit Function
    End If
    RemoveManualEntry = ClosedPeriodProblem(DbValue("SELECT EntryDate FROM ManualEntries WHERE ManualEntryID = " & _
                                                    ManualEntryID))                                   ' modClosing
    If Len(RemoveManualEntry) > 0 Then Exit Function
    CurrentDb.Execute "DELETE FROM ManualEntries WHERE ManualEntryID = " & ManualEntryID, dbFailOnError   ' lines cascade
    LogAction "MANUAL_ENTRY_DELETE", "ManualEntries", CStr(ManualEntryID)
End Function

Public Function JournalEntryOfManual(ByVal ManualEntryID As Long) As Variant
    JournalEntryOfManual = DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType = 'MANUAL' AND SourceID = " & _
                                   ManualEntryID)
End Function

'------------------------------------------------------------------------------
' Screen frmManualEntry (OpenArgs: ManualEntryID to open)
'------------------------------------------------------------------------------
Public Sub ManualEntryLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    EnsureLocalTables
    If Nz(frm.OpenArgs, 0) > 0 Then
        ManualOpen frm, CLng(frm.OpenArgs)
    Else
        ManualNew frm
    End If
End Sub

Public Sub ManualNew(ByVal frm As Access.Form)
    CurrentDb.Execute "DELETE FROM tmpManualLines", dbFailOnError
    frm!txtEntryID.Value = Null
    frm!txtReversalOf.Value = Null
    frm!txtNumber.Value = "جديد"
    frm!txtDate.Value = Date
    frm!txtDescription.Value = Null
    frm!txtReference.Value = Null
    frm!cboFind.Value = Null
    frm!subLines.Form.Requery
    ManualRecalc frm
    ManualButtons frm
    frm!lblStatus.Caption = "اكتب التاريخ والبيان، ثم سطرًا لكل حساب: المبلغ في المدين أو في الدائن."
    SafeFocus frm!txtDescription
End Sub

Public Sub ManualOpen(ByVal frm As Access.Form, ByVal ManualEntryID As Long)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM ManualEntries WHERE ManualEntryID = " & ManualEntryID, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        ShowWarning "القيد غير موجود."
        ManualNew frm
        Exit Sub
    End If
    frm!txtEntryID.Value = rs!ManualEntryID
    frm!txtReversalOf.Value = rs!ReversalOfID
    frm!txtNumber.Value = rs!EntryNumber
    frm!txtDate.Value = rs!EntryDate
    frm!txtDescription.Value = rs!Description
    frm!txtReference.Value = rs!Reference
    rs.Close
    CurrentDb.Execute "DELETE FROM tmpManualLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpManualLines (AccountCode, Debit, Credit, LineText) SELECT AccountCode, Debit, " & _
                      "Credit, LineText FROM ManualEntryLines WHERE ManualEntryID = " & ManualEntryID & _
                      " ORDER BY LineNumber", dbFailOnError
    frm!subLines.Form.Requery
    ManualRecalc frm
    ManualButtons frm
    If Len(ClosedPeriodProblem(frm!txtDate.Value)) > 0 Then
        frm!lblStatus.Caption = "قيد في فترة مقفلة: للعرض فقط. يمكن عمل قيد عكسي بتاريخ مفتوح."
    Else
        frm!lblStatus.Caption = "قيد محفوظ: عدّل ثم احفظ، فيتحدث قيده في اليومية بنفس رقمه."
    End If
End Sub

Public Sub ManualFindPicked(ByVal frm As Access.Form)
    If IsNull(frm!cboFind.Value) Then Exit Sub
    ManualOpen frm, frm!cboFind.Value
End Sub

Private Sub ManualButtons(ByVal frm As Access.Form)
    Dim saved As Boolean
    saved = Not IsNull(frm!txtEntryID.Value)
    SafeFocus frm!txtDescription                  ' a button that has the focus cannot be disabled
    frm!btnDelete.Enabled = saved
    frm!btnReverse.Enabled = saved
    frm!btnPrint.Enabled = saved
    frm!btnInJournal.Enabled = saved
End Sub

Public Sub ManualLineChanged(ByVal sf As Access.Form, ByVal FieldName As String)
    ' one side per line: an amount on one side clears the other
    If FieldName = "Debit" And Nz(sf!Debit.Value, 0) <> 0 Then sf!Credit.Value = 0
    If FieldName = "Credit" And Nz(sf!Credit.Value, 0) <> 0 Then sf!Debit.Value = 0
    If IsNull(sf!Debit.Value) Then sf!Debit.Value = 0
    If IsNull(sf!Credit.Value) Then sf!Credit.Value = 0
    If sf.Dirty Then sf.Dirty = False
    ManualRecalc sf.Parent
End Sub

Public Sub ManualRemoveLine(ByVal sf As Access.Form)
    If sf.NewRecord Then
        sf.Undo
        Exit Sub
    End If
    CurrentDb.Execute "DELETE FROM tmpManualLines WHERE LineNo = " & sf!LineNo.Value, dbFailOnError
    sf.Requery
    ManualRecalc sf.Parent
End Sub

Public Sub ManualRecalc(ByVal frm As Access.Form)
    Dim debit As Currency, credit As Currency
    debit = Nz(DbValue("SELECT Sum(Debit) FROM tmpManualLines"), 0)
    credit = Nz(DbValue("SELECT Sum(Credit) FROM tmpManualLines"), 0)
    frm!lblTotals.Caption = "المدين: " & Format$(debit, "#,##0.00") & "     الدائن: " & Format$(credit, "#,##0.00") & _
                            "     " & IIf(debit = credit And debit > 0, "متوازن", "الفرق: " & Format$(Abs(debit - credit), "#,##0.00"))
    frm!lblTotals.ForeColor = IIf(debit = credit And debit > 0, CLR_SUCCESS, CLR_DANGER)
End Sub

Public Function SaveManualEntry(ByVal frm As Access.Form) As Boolean
    Dim id As Long, msg As String, newID As Long, number As String, sync As String
    id = Nz(frm!txtEntryID.Value, 0)
    If Not CanScreenAction(frm.Name, IIf(id = 0, "ADD", "EDIT")) Then Exit Function      ' frmUserScreens
    If frm!subLines.Form.Dirty Then frm!subLines.Form.Dirty = False
    msg = PostManualEntry(id, frm!txtDate.Value, Nz(frm!txtDescription.Value, ""), Nz(frm!txtReference.Value, ""), _
                          frm!txtReversalOf.Value, newID)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    sync = SyncJournal()                                    ' the journal follows at once
    ManualOpen frm, newID
    frm!cboFind.Requery
    number = Nz(frm!txtNumber.Value, "")
    msg = "تم حفظ القيد " & number & "."
    If Len(sync) = 0 Then
        msg = msg & vbCrLf & "قيد اليومية: " & Nz(DbValue("SELECT EntryNumber FROM JournalEntries WHERE EntryID = " & _
                                                          Nz(JournalEntryOfManual(newID), 0)), "-")
    Else
        msg = msg & vbCrLf & sync
    End If
    ShowInfo msg
    SaveManualEntry = True
End Function

Public Sub DeleteManualEntry(ByVal frm As Access.Form)
    Dim id As Long, msg As String
    id = Nz(frm!txtEntryID.Value, 0)
    If id = 0 Then Exit Sub
    If Not CanScreenAction(frm.Name, "DELETE") Then Exit Sub
    If Not AskYesNo("حذف القيد " & frm!txtNumber.Value & " نهائيًا؟ يُحذف قيده من اليومية أيضًا.") Then Exit Sub
    msg = RemoveManualEntry(id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    SyncJournal
    frm!cboFind.Requery
    ManualNew frm
    frm!lblStatus.Caption = "تم حذف القيد."
End Sub

Public Sub ReverseManualEntry(ByVal frm As Access.Form)
    ' a new entry with debit and credit swapped, dated today, for review before saving
    Dim id As Long, number As String
    id = Nz(frm!txtEntryID.Value, 0)
    If id = 0 Then Exit Sub
    number = Nz(frm!txtNumber.Value, "")
    CurrentDb.Execute "DELETE FROM tmpManualLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpManualLines (AccountCode, Debit, Credit, LineText) SELECT AccountCode, Credit, " & _
                      "Debit, LineText FROM ManualEntryLines WHERE ManualEntryID = " & id & " ORDER BY LineNumber", dbFailOnError
    frm!txtDescription.Value = Left$("عكس القيد " & number & ": " & Nz(frm!txtDescription.Value, ""), 255)
    frm!txtReversalOf.Value = id
    frm!txtEntryID.Value = Null
    frm!txtNumber.Value = "جديد"
    frm!txtDate.Value = Date
    frm!cboFind.Value = Null
    frm!subLines.Form.Requery
    ManualRecalc frm
    ManualButtons frm
    frm!lblStatus.Caption = "قيد عكسي جديد للقيد " & number & ": راجع التاريخ والأسطر ثم اضغط «حفظ القيد»."
End Sub

Public Sub PrintManualEntry(ByVal frm As Access.Form)
    Dim id As Long
    id = Nz(frm!txtEntryID.Value, 0)
    If id = 0 Then Exit Sub
    SyncJournal
    PrintJournalEntry JournalEntryOfManual(id)
End Sub

Public Sub ManualOpenInJournal(ByVal frm As Access.Form)
    Dim id As Long, entry As Variant
    id = Nz(frm!txtEntryID.Value, 0)
    If id = 0 Then Exit Sub
    SyncJournal
    entry = JournalEntryOfManual(id)
    If IsNull(entry) Then
        ShowWarning "لم يُنشأ قيد اليومية لهذا القيد بعد. اضغط «تحديث القيود» في شاشة قيود اليومية."
        Exit Sub
    End If
    OpenScreen "frmJournalEntry", 0, entry
End Sub
