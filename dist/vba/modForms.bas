Attribute VB_Name = "modForms"
'==============================================================================
' modForms  -  Retail Store Management System (Phase 5)
'
' Behaviour shared by every data-entry screen built by modBuildForms.
' Each form keeps its settings in its Tag property, for example:
'   KIND=LIST|TABLE=Products|PK=ProductID|LIST=SELECT ...|SEARCH=t.ProductName,t.Barcode
'   |ACTIVE=t.IsActive|SEQ=PRODUCT_CODE:ProductCode|UNIQUE=ProductCode,Barcode
' The form modules only contain one-line event procedures that call this module.
'==============================================================================
Option Compare Database
Option Explicit

Private Const ERR_RELATED_RECORDS As Long = 3200
Private m_editable As Object        ' form name -> "|ctl|ctl|": bound controls the user normally types in

'------------------------------------------------------------------------------
' Form events
'------------------------------------------------------------------------------
Public Sub FormLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    If TagValue(frm, "TABLE") = "LabelSettings" Then LabelSettingsLoad frm      ' printer list (modLabels)
    If TagValue(frm, "TABLE") = "Employees" And Not IsDeveloper() Then
        frm.RecordSource = "SELECT * FROM Employees WHERE IsDeveloper = False"  ' the programmer stays hidden
    End If
    If TagValue(frm, "TABLE") = "Settings" Then SettingsScreenLoad frm
    If TagValue(frm, "TABLE") = "Accounts" Then RebuildAccountTree                    ' tree order (modAccounts)
    If Not CanScreenAction(frm.Name, "ADD", True) Then
        frm.AllowAdditions = False
        If ControlExists(frm, "btnNew") Then SetEnabled frm!btnNew, False
    End If
    If TagValue(frm, "KIND") = "LIST" Then RefreshList frm
    If Not IsNull(frm.OpenArgs) Then
        GoToRecord frm, frm.OpenArgs
    ElseIf frm.Recordset.RecordCount = 0 And frm.AllowAdditions Then
        DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
    End If
End Sub

Public Sub FitControls(ByVal frm As Access.Form, ByVal DesignW As Long, ByVal DesignH As Long, _
                       ByVal MinDH As Long, ByVal Mirror As Boolean, ByVal Spec As String)
    ' Form_Resize of the large screens: the screen is designed at its smallest size; extra window
    ' width/height goes to the lists and grids, and the buttons follow the window edges.
    ' Spec: "name,Left,Top,Width,Height,mx,mw,my,mh;..." (m* = share of the extra size, per mille)
    Dim dw As Long, dh As Long, newW As Long, newH As Long, items As Variant, f As Variant, i As Long
    Dim L As Long, T As Long, W As Long, H As Long
    On Error Resume Next                         ' resizing must never interrupt the user
    dw = frm.InsideWidth - DesignW
    If dw < 0 Then dw = 0
    dh = frm.InsideHeight - DesignH
    If dh < MinDH Then dh = MinDH
    newW = DesignW + dw
    newH = DesignH + dh
    If newW > frm.Width Then frm.Width = newW    ' grow first, so the controls fit while moving
    If newH > frm.Section(0).Height Then frm.Section(0).Height = newH
    items = Split(Spec, ";")
    For i = LBound(items) To UBound(items)
        f = Split(items(i), ",")
        If UBound(f) = 8 Then
            W = CLng(f(3)) + dw * CLng(f(6)) \ 1000
            H = CLng(f(4)) + dh * CLng(f(8)) \ 1000
            L = CLng(f(1)) + dw * CLng(f(5)) \ 1000
            T = CLng(f(2)) + dh * CLng(f(7)) \ 1000
            If Mirror Then L = newW - L - W
            frm.Controls(CStr(f(0))).Move L, T, W, H
        End If
    Next
    frm.Width = newW                             ' then shrink to the new size
    frm.Section(0).Height = newH
End Sub

Public Sub FormCurrent(ByVal frm As Access.Form)
    Dim pk As String
    pk = TagValue(frm, "PK")
    If ControlExists(frm, "lstItems") Then
        If frm.NewRecord Then
            frm!lstItems.Value = Null
        Else
            frm!lstItems.Value = frm(pk).Value
        End If
    End If
    If frm.NewRecord Then
        SetStatus frm, "”Ã· ÃœÌœ - √œŒ· «·»Ì«‰«  À„ «÷€ÿ Õ›Ÿ", CLR_ACCENT
    Else
        SetStatus frm, "", CLR_MUTED
    End If
    If ControlExists(frm, "btnDelete") Then
        SetEnabled frm!btnDelete, Not frm.NewRecord And CanScreenAction(frm.Name, "DELETE", True)
    End If
    LockForActions frm

    Select Case TagValue(frm, "TABLE")
        Case "Products": UpdatePriceInfo frm
        Case "Customers", "Suppliers": LockPartnerFields frm
        Case "Employees"                                            ' frmEmployeePay reads it too
            If frm.Name = "frmUsers" Then UserCurrent frm           ' modSecurityScreens
    End Select
End Sub

Public Function FormBeforeUpdate(ByVal frm As Access.Form) As Boolean
    Dim closedMsg As String
    If Not CanScreenAction(frm.Name, IIf(frm.NewRecord, "ADD", "EDIT")) Then Exit Function
    closedMsg = ClosedRecordProblem(frm)                    ' a closed period (modClosing)
    If Len(closedMsg) > 0 Then
        ShowWarning closedMsg
        Exit Function
    End If
    If Not CheckRequired(frm) Then Exit Function
    Select Case TagValue(frm, "TABLE")
        Case "Products"
            If Not ValidateProduct(frm) Then Exit Function
        Case "Customers", "Suppliers"
            If Not ValidatePartner(frm) Then Exit Function
        Case "Expenses"
            If Not ValidateExpense(frm) Then Exit Function
        Case "Settings"
            If Not ValidateSettings(frm) Then Exit Function
        Case "Employees"
            If frm.Name = "frmUsers" Then
                If Not ValidateEmployee(frm) Then Exit Function    ' modSecurity
            End If
        Case "Accounts"
            If Not ValidateAccount(frm) Then Exit Function         ' modAccounts
    End Select
    If Not CheckUnique(frm) Then Exit Function
    If Not AssignSequence(frm) Then Exit Function      ' last: a refused save wastes no number
    If HasRecordField(frm, "UpdatedAt") Then frm("UpdatedAt").Value = Now
    FormBeforeUpdate = True
End Function

Public Sub FormAfterUpdate(ByVal frm As Access.Form)
    Dim pk As String
    pk = TagValue(frm, "PK")
    LogAction "SAVE", TagValue(frm, "TABLE"), CStr(Nz(frm(pk).Value, ""))
    If TagValue(frm, "TABLE") = "Accounts" Then RebuildAccountTree                    ' level and place in the tree
    RefreshList frm
    SetStatus frm, " „ «·Õ›Ÿ", CLR_SUCCESS
End Sub

Public Function FormError(ByVal frm As Access.Form, ByVal DataErr As Integer) As Integer
    Dim msg As String
    Select Case DataErr
        Case 3022
            msg = "·« Ì„ﬂ‰ «·Õ›Ÿ:  ÊÃœ ﬁÌ„… „ﬂ——… ›Ì Õﬁ· ÌÃ» √‰ ÌﬂÊ‰ ›—Ìœ« („À· «·ﬂÊœ √Ê «·»«—ﬂÊœ)."
        Case 3314, 3162, 2169
            msg = "ÌÊÃœ Õﬁ· „ÿ·Ê» ›«—€. √ﬂ„· «·»Ì«‰«  À„ «Õ›Ÿ."
        Case 3201
            msg = "«Œ — ﬁÌ„… ’ÕÌÕ… „‰ «·ﬁ«∆„…."
        Case ERR_RELATED_RECORDS
            msg = "·« Ì„ﬂ‰ «·Õ–› √Ê «· ⁄œÌ· ·ÊÃÊœ ⁄„·Ì«  „— »ÿ… »Â–« «·”Ã·."
        Case 2113, 2279
            msg = "«·ﬁÌ„… «·„œŒ·… €Ì— ’ÕÌÕ… ·Â–« «·Õﬁ·."
        Case 2237
            msg = "«Œ — ﬁÌ„… „‰ «·ﬁ«∆„…."
        Case Else
            FormError = acDataErrDisplay   ' includes the Arabic validation texts of the tables
            Exit Function
    End Select
    ShowError msg
    FormError = acDataErrContinue
End Function

Public Sub FormKeyDown(ByVal frm As Access.Form, ByRef KeyCode As Integer, ByVal Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyEscape
            KeyCode = 0
            FormAction frm, "CLOSE"
        Case KeyCode = vbKeyS And (Shift And acCtrlMask) <> 0
            KeyCode = 0
            FormAction frm, "SAVE"
        Case KeyCode = vbKeyF2 And frm.AllowAdditions
            KeyCode = 0
            FormAction frm, "NEW"
        Case KeyCode = vbKeyF3
            If ControlExists(frm, "txtSearch") Then
                KeyCode = 0
                SafeFocus frm!txtSearch
            End If
    End Select
End Sub

Public Function FormUnload(ByVal frm As Access.Form) As Boolean
    FormUnload = True
    If frm.Dirty Then
        If AskYesNo(" ÊÃœ  ⁄œÌ·«  €Ì— „Õ›ÊŸ…. Â·  —Ìœ Õ›ŸÂ« ﬁ»· «·≈€·«ﬁø") Then
            FormUnload = SaveRecord(frm)
        Else
            frm.Undo
        End If
    End If
End Function

'------------------------------------------------------------------------------
' Toolbar buttons: NEW, SAVE, UNDO, DELETE, CLOSE
'------------------------------------------------------------------------------
Public Sub FormAction(ByVal frm As Access.Form, ByVal Action As String)
    Select Case Action
        Case "NEW"
            If Not SaveRecord(frm) Then Exit Sub
            DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
            FocusFirstInput frm
        Case "SAVE"
            If Not frm.Dirty Then
                SetStatus frm, "·«  ÊÃœ  ⁄œÌ·«  ··Õ›Ÿ", CLR_MUTED
            ElseIf SaveRecord(frm) Then
                ShowInfo " „ «·Õ›Ÿ »‰Ã«Õ."
            End If
        Case "UNDO"
            frm.Undo
            SetStatus frm, " „ «· —«Ã⁄ ⁄‰ «· ⁄œÌ·« ", CLR_MUTED
        Case "DELETE"
            DeleteRecord frm
        Case "CLOSE"
            If FormUnload(frm) Then
                If frm.Dirty Then frm.Undo
                DoCmd.Close acForm, frm.Name, acSaveNo
            End If
    End Select
End Sub

Public Function SaveRecord(ByVal frm As Access.Form) As Boolean
    ' Saves the current record; False when validation (or Access) refused it.
    If Not frm.Dirty Then
        SaveRecord = True
        Exit Function
    End If
    On Error Resume Next
    frm.Dirty = False
    On Error GoTo 0
    SaveRecord = Not frm.Dirty
End Function

Private Sub DeleteRecord(ByVal frm As Access.Form)
    Dim table As String, pk As String, id As Variant, activeField As String, errNo As Long

    If frm.NewRecord Then
        frm.Undo
        Exit Sub
    End If
    If Not CanScreenAction(frm.Name, "DELETE") Then Exit Sub
    If Len(ClosedRecordProblem(frm, True)) > 0 Then          ' a closed period (modClosing)
        ShowWarning ClosedRecordProblem(frm, True)
        Exit Sub
    End If
    table = TagValue(frm, "TABLE")
    pk = TagValue(frm, "PK")
    id = frm(pk).Value
    If table = "Accounts" Then
        If Len(AccountDeleteProblem(CLng(id))) > 0 Then
            ShowWarning AccountDeleteProblem(CLng(id))
            Exit Sub
        End If
    End If
    If HasRecordField(frm, "IsSystem") Then
        If Nz(frm("IsSystem").Value, False) Then
            ShowWarning "Â–« ”Ã· √”«”Ì ›Ì «·‰Ÿ«„ Ê·« Ì„ﬂ‰ Õ–›Â."
            Exit Sub
        End If
    End If
    If Not AskYesNo("Â·  —Ìœ Õ–› Â–« «·”Ã· ‰Â«∆Ì«ø") Then Exit Sub
    If frm.Dirty Then frm.Undo

    On Error Resume Next
    CurrentDb.Execute "DELETE FROM [" & table & "] WHERE [" & pk & "] = " & id, dbFailOnError
    errNo = Err.Number
    On Error GoTo 0

    If errNo = 0 Then
        LogAction "DELETE", table, CStr(id)
        frm.Requery
        RefreshList frm
        SetStatus frm, " „ «·Õ–›", CLR_SUCCESS
        Exit Sub
    End If

    activeField = Mid$(TagValue(frm, "ACTIVE"), InStr(TagValue(frm, "ACTIVE"), ".") + 1)
    If errNo = ERR_RELATED_RECORDS And Len(activeField) > 0 Then
        If AskYesNo("·« Ì„ﬂ‰ Õ–› Â–« «·”Ã· ·√‰Â „” Œœ„ ›Ì ⁄„·Ì«  ”«»ﬁ…." & vbCrLf & _
                    "Â·  —Ìœ  ⁄ÿÌ·Â »œ·« „‰ –·ﬂø (ÌŒ ›Ì „‰ «·ﬁÊ«∆„ ÊÌ»ﬁÏ ›Ì «· ﬁ«—Ì—)") Then
            frm(activeField).Value = False      ' through the form, so validation runs
            If SaveRecord(frm) Then
                LogAction "DEACTIVATE", table, CStr(id)
                SetStatus frm, " „  ⁄ÿÌ· «·”Ã·", CLR_WARNING
            End If
        End If
    ElseIf errNo = ERR_RELATED_RECORDS Then
        ShowWarning "·« Ì„ﬂ‰ Õ–› Â–« «·”Ã· ·√‰Â „” Œœ„ ›Ì ⁄„·Ì«  ”«»ﬁ…."
    Else
        ShowError " ⁄–— «·Õ–› (Œÿ√ " & errNo & ")."
    End If
End Sub

'------------------------------------------------------------------------------
' Record list (lstItems + txtSearch + chkShowInactive)
'------------------------------------------------------------------------------
Public Sub RefreshList(ByVal frm As Access.Form)
    Dim sql As String, activeCond As String, searchCond As String, txt As String
    Dim f As Variant, n As Long
    If Not ControlExists(frm, "lstItems") Then Exit Sub

    activeCond = "True"
    If Len(TagValue(frm, "ACTIVE")) > 0 Then
        If Not Nz(frm!chkShowInactive.Value, False) Then activeCond = TagValue(frm, "ACTIVE") & " = True"
    End If

    searchCond = "True"
    txt = Trim$(SearchText(frm))
    If Len(txt) > 0 Then
        searchCond = ""
        For Each f In Split(TagValue(frm, "SEARCH"), ",")
            If Len(searchCond) > 0 Then searchCond = searchCond & " OR "
            searchCond = searchCond & f & " Like " & LikePattern(txt)
        Next
    End If

    If TagValue(frm, "TABLE") = "Employees" And Not IsDeveloper() Then
        searchCond = "t.IsDeveloper = False AND (" & searchCond & ")"        ' the programmer stays hidden
    End If
    sql = Replace(TagValue(frm, "LIST"), "{ACTIVE}", activeCond)
    sql = Replace(sql, "{SEARCH}", searchCond)
    frm!lstItems.RowSource = sql
    n = frm!lstItems.ListCount + (frm!lstItems.ColumnHeads * 1)   ' ColumnHeads True = -1
    If ControlExists(frm, "lblCount") Then frm!lblCount.Caption = n & " ”Ã·"
    If Not frm.NewRecord Then frm!lstItems.Value = frm(TagValue(frm, "PK")).Value
End Sub

Public Sub ListPick(ByVal frm As Access.Form)
    If IsNull(frm!lstItems.Value) Then Exit Sub
    If Not SaveRecord(frm) Then Exit Sub
    GoToRecord frm, frm!lstItems.Value
End Sub

Private Function SearchText(ByVal frm As Access.Form) As String
    ' While typing, .Text holds the new text; .Value is updated only on exit.
    If Not ControlExists(frm, "txtSearch") Then Exit Function
    On Error Resume Next
    SearchText = Nz(frm!txtSearch.Value, "")
    SearchText = frm!txtSearch.Text
End Function

Public Sub GoToRecord(ByVal frm As Access.Form, ByVal RecordID As Variant)
    Dim rs As DAO.Recordset
    If Not IsNumeric(RecordID) Then Exit Sub
    Set rs = frm.RecordsetClone
    rs.FindFirst "[" & TagValue(frm, "PK") & "] = " & CLng(RecordID)
    If Not rs.NoMatch Then frm.Bookmark = rs.Bookmark
End Sub

'------------------------------------------------------------------------------
' Generic validation
'------------------------------------------------------------------------------
Private Function CheckRequired(ByVal frm As Access.Form) As Boolean
    Dim ctl As Access.Control, src As String, seqField As String
    seqField = Mid$(TagValue(frm, "SEQ"), InStr(TagValue(frm, "SEQ"), ":") + 1)
    For Each ctl In frm.Controls
        src = BoundField(ctl)
        If Len(src) > 0 And StrComp(src, seqField, vbTextCompare) <> 0 Then
            If FieldIsRequired(frm, src) Then
                If Len(Trim$(Nz(ctl.Value, ""))) = 0 Then
                    ShowWarning "«·Õﬁ· ´" & CaptionOf(ctl) & "ª „ÿ·Ê»."
                    SafeFocus ctl
                    Exit Function
                End If
            End If
        End If
    Next
    CheckRequired = True
End Function

Private Function CheckUnique(ByVal frm As Access.Form) As Boolean
    Dim f As Variant, v As Variant, pkValue As Variant, table As String, pk As String
    table = TagValue(frm, "TABLE")
    pk = TagValue(frm, "PK")
    ' a new record has no current row in frm.Recordset (error 3021): read the form itself
    pkValue = 0
    If Not frm.NewRecord Then pkValue = Nz(frm(pk).Value, 0)
    For Each f In Split(TagValue(frm, "UNIQUE"), ",")
        If Len(f) > 0 Then
            v = frm(f).Value
            If Not IsNull(v) Then
                If DCount("*", table, "[" & f & "] = " & SqlText(v) & " AND [" & pk & "] <> " & _
                          pkValue) > 0 Then
                    ShowWarning "«·ﬁÌ„… ´" & v & "ª ›Ì Õﬁ· ´" & CaptionOf(frm(f)) & _
                                "ª „” Œœ„… ·”Ã· ¬Œ—."
                    SafeFocus frm(f)
                    Exit Function
                End If
            End If
        End If
    Next
    CheckUnique = True
End Function

Private Function AssignSequence(ByVal frm As Access.Form) As Boolean
    ' SEQ=SEQUENCE_NAME:FieldName  -> fills the field with the next free number
    Dim spec As String, seqName As String, f As String, candidate As String, tries As Integer
    AssignSequence = True
    spec = TagValue(frm, "SEQ")
    If Len(spec) = 0 Then Exit Function
    seqName = Left$(spec, InStr(spec, ":") - 1)
    f = Mid$(spec, InStr(spec, ":") + 1)
    If Len(Trim$(Nz(frm(f).Value, ""))) > 0 Then Exit Function
    Do
        candidate = NextNumber(seqName)
        tries = tries + 1
    Loop While DCount("*", TagValue(frm, "TABLE"), "[" & f & "] = " & SqlText(candidate)) > 0 _
               And tries < 1000
    frm(f).Value = candidate
End Function

'------------------------------------------------------------------------------
' Screen-specific rules
'------------------------------------------------------------------------------
Private Function ValidateProduct(ByVal frm As Access.Form) As Boolean
    Dim cost As Currency
    If frm.NewRecord Then frm!CurrentQuantity.Value = 0
    If Nz(frm!CurrentQuantity.Value, 0) = 0 And Nz(frm!AverageCost.Value, 0) = 0 Then
        frm!AverageCost.Value = Nz(frm!PurchasePrice.Value, 0)   ' no stock yet
    End If
    cost = Nz(frm!AverageCost.Value, 0)
    If Nz(frm!SellingPrice.Value, 0) < cost And cost > 0 Then
        If Not AskYesNo("”⁄— «·»Ì⁄ (" & Format$(frm!SellingPrice.Value, "#,##0.00") & _
                        ") √ﬁ· „‰ «· ﬂ·›… (" & Format$(cost, "#,##0.00") & ")." & vbCrLf & _
                        "Â·  —Ìœ «·Õ›Ÿ ⁄·Ï √Ì Õ«·ø") Then
            SafeFocus frm!SellingPrice
            Exit Function
        End If
    End If
    ValidateProduct = True
End Function

Private Function ValidatePartner(ByVal frm As Access.Form) As Boolean
    Dim mobile As String, email As String
    mobile = Trim$(Nz(frm!Mobile.Value, ""))
    If Len(mobile) > 0 And Not (mobile Like "05########" Or mobile Like "+9665########" _
                                Or mobile Like "9665########") Then
        If Not AskYesNo("—ﬁ„ «·ÃÊ«· ´" & mobile & "ª ·Ì” »’Ì€… ”⁄ÊœÌ… (05xxxxxxxx)." & vbCrLf & _
                        "Â·  —Ìœ «·Õ›Ÿ ⁄·Ï √Ì Õ«·ø") Then
            SafeFocus frm!Mobile
            Exit Function
        End If
    End If
    email = Trim$(Nz(frm!Email.Value, ""))
    If Len(email) > 0 And Not email Like "*?@?*.?*" Then
        ShowWarning "«·»—Ìœ «·≈·ﬂ —Ê‰Ì €Ì— ’ÕÌÕ."
        SafeFocus frm!Email
        Exit Function
    End If
    If HasRecordField(frm, "IsSystem") Then
        If Nz(frm("IsSystem").Value, False) Then
            frm!AllowCredit.Value = False
            frm!OpeningBalance.Value = 0
        End If
    End If
    ' Without movements the balance is just the opening balance.
    If frm.NewRecord Or Not PartnerHasMovements(frm) Then
        frm!CurrentBalance.Value = Nz(frm!OpeningBalance.Value, 0)
    End If
    ValidatePartner = True
End Function

Private Function ValidateExpense(ByVal frm As Access.Form) As Boolean
    If Not frm.NewRecord Then
        If Nz(DbValue("SELECT COUNT(*) FROM CashVouchers WHERE ExpenseID = " & Nz(frm!ExpenseID.Value, 0)), 0) > 0 Then
            ShowWarning "Â–« «·„’—Ê› „”Ã· „‰ ”‰œ ’—› ‰ﬁœÌ… ›Ì «·Œ“Ì‰… Ê·« Ìı⁄œÛ¯· „‰ Â‰«." & vbCrLf & _
                        "·· ’ÕÌÕ ”Ã¯· ”‰œ ﬁ»÷ ‰ﬁœÌ… »«·„»·€ À„ ”‰œ ’—› ÃœÌœ."
            Exit Function
        End If
    End If
    frm!TotalAmount.Value = Nz(frm!Amount.Value, 0) + Nz(frm!Tax.Value, 0)
    If IsNull(frm!EmployeeID.Value) Then frm!EmployeeID.Value = CurrentUserID()
    If frm.NewRecord And IsNull(frm!CostCenterID.Value) Then frm!CostCenterID.Value = CostCenterFor()   ' modCostCenters
    ' a cash expense leaves a cash box (modCash); other payment methods do not
    If Nz(frm!PaymentMethodID.Value, 0) = CASH_METHOD_ID Then
        If IsNull(frm!CashBoxID.Value) Then frm!CashBoxID.Value = CurrentCashBoxID()
    Else
        frm!CashBoxID.Value = Null
    End If
    ' a bank transfer leaves the default bank (modBank)
    If Nz(frm!PaymentMethodID.Value, 0) <> BANK_TRANSFER_METHOD_ID Then
        frm!BankID.Value = Null
    ElseIf IsNull(frm!BankID.Value) Then
        frm!BankID.Value = BankFor(frm!PaymentMethodID.Value, Nz(frm!TotalAmount.Value, 0))
    End If
    If frm!ExpenseDate.Value > Date Then
        If Not AskYesNo(" «—ÌŒ «·„’—Ê› ›Ì «·„” ﬁ»·. Â·  —Ìœ «·Õ›Ÿ ⁄·Ï √Ì Õ«·ø") Then
            SafeFocus frm!ExpenseDate
            Exit Function
        End If
    End If
    ValidateExpense = True
End Function

Public Sub SettingsScreenLoad(ByVal frm As Access.Form)
    ' The shop name: the programmer, or an administrator when the programmer allows it.
    ' The permission itself (AllowAdminCompanyName) is shown to the programmer only.
    Dim allowed As Boolean
    allowed = CanChangeStoreName()
    frm!StoreName.Locked = Not allowed
    frm!StoreNameEn.Locked = Not allowed
    If ControlExists(frm, "lblStoreNameNote") Then
        frm!lblStoreNameNote.Caption = IIf(allowed, " ", " €ÌÌ— «”„ «·„Õ· Ì „ ⁄‰ ÿ—Ìﬁ «·„»—„Ã.")
    End If
    frm!AllowAdminCompanyName.Visible = IsDeveloper()
End Sub

Private Function ValidateSettings(ByVal frm As Access.Form) As Boolean
    If Not CanChangeStoreName() And (Nz(frm!StoreName.Value, "") <> Nz(frm!StoreName.OldValue, "") Or _
                                     Nz(frm!StoreNameEn.Value, "") <> Nz(frm!StoreNameEn.OldValue, "")) Then
        ShowWarning " €ÌÌ— «”„ «·„Õ· Ì „ ⁄‰ ÿ—Ìﬁ «·„»—„Ã."
        Exit Function
    End If
    If Not IsDeveloper() And Nz(frm!AllowAdminCompanyName.Value, False) <> Nz(frm!AllowAdminCompanyName.OldValue, False) Then
        ShowWarning "Â–« «·ŒÌ«— ··„»—„Ã ›ﬁÿ."
        Exit Function
    End If
    If Len(Nz(frm!VATNumber.Value, "")) = 0 Then
        If Not AskYesNo("·„ ÌıœŒ· «·—ﬁ„ «·÷—Ì»Ì° Ê·‰  ﬂÊ‰ «·›Ê« Ì— ›Ê« Ì— ÷—Ì»Ì… ‰Ÿ«„Ì…." & vbCrLf & _
                        "Â·  —Ìœ «·Õ›Ÿ ⁄·Ï √Ì Õ«·ø") Then
            SafeFocus frm!VATNumber
            Exit Function
        End If
    End If
    If Nz(frm!VATRate.Value, 0) <> 0.15 Then
        If Not AskYesNo("‰”»… «·÷—Ì»… ·Ì”  15%. Â· √‰  „ √ﬂœø") Then
            SafeFocus frm!VATRate
            Exit Function
        End If
    End If
    ValidateSettings = True
End Function

Private Function PartnerHasMovements(ByVal frm As Access.Form) As Boolean
    Dim id As Long
    id = Nz(frm(TagValue(frm, "PK")).Value, 0)
    If TagValue(frm, "TABLE") = "Customers" Then
        PartnerHasMovements = DCount("*", "SalesInvoices", "CustomerID = " & id) + _
                              DCount("*", "SalesReturns", "CustomerID = " & id) + _
                              DCount("*", "CustomerPayments", "CustomerID = " & id) > 0
    Else
        PartnerHasMovements = DCount("*", "PurchaseInvoices", "SupplierID = " & id) + _
                              DCount("*", "PurchaseReturns", "SupplierID = " & id) + _
                              DCount("*", "SupplierPayments", "SupplierID = " & id) > 0
    End If
End Function

Private Sub LockPartnerFields(ByVal frm As Access.Form)
    ' Opening balance can only change while the customer/supplier has no movements.
    Dim locked As Boolean, isSystem As Boolean
    If Not frm.NewRecord Then locked = PartnerHasMovements(frm)
    SetLocked frm!OpeningBalance, locked
    If HasRecordField(frm, "IsSystem") Then
        If Not frm.NewRecord Then isSystem = Nz(frm("IsSystem").Value, False)
        SetLocked frm!AllowCredit, isSystem
        SetLocked frm!IsActive, isSystem
    End If
End Sub

'------------------------------------------------------------------------------
' Field hooks (AfterUpdate of selected fields) and extra buttons
'------------------------------------------------------------------------------
Public Sub FieldChanged(ByVal frm As Access.Form, ByVal FieldName As String)
    Select Case TagValue(frm, "TABLE")
        Case "Accounts"
            If FieldName = "ParentCode" Then AccountParentChanged frm   ' modAccounts
        Case "Products"
            UpdatePriceInfo frm
        Case "Expenses"
            frm!TotalAmount.Value = Nz(frm!Amount.Value, 0) + Nz(frm!Tax.Value, 0)
            If FieldName = "PaymentMethodID" Then
                If Nz(frm!PaymentMethodID.Value, 0) = CASH_METHOD_ID Then
                    If IsNull(frm!CashBoxID.Value) Then frm!CashBoxID.Value = CurrentCashBoxID()
                Else
                    frm!CashBoxID.Value = Null
                End If
            End If
    End Select
End Sub

Public Sub CalcExpenseVat(ByVal frm As Access.Form)
    frm!Tax.Value = RoundMoney(Nz(frm!Amount.Value, 0) * Nz(SettingValue("VATRate"), 0.15))
    FieldChanged frm, "Tax"
End Sub

Public Sub BrowseFolder(ByVal frm As Access.Form, ByVal FieldName As String)
    With Application.FileDialog(4)          ' msoFileDialogFolderPicker
        .Title = "«Œ — «·„Ã·œ"
        If .Show Then frm(FieldName).Value = .SelectedItems(1)
    End With
End Sub

Public Sub BrowseFile(ByVal frm As Access.Form, ByVal FieldName As String)
    With Application.FileDialog(3)          ' msoFileDialogFilePicker
        .Title = "«Œ — «·’Ê—…"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "«·’Ê—", "*.png;*.jpg;*.jpeg;*.bmp"
        If .Show Then frm(FieldName).Value = .SelectedItems(1)
    End With
End Sub

Private Sub UpdatePriceInfo(ByVal frm As Access.Form)
    Dim price As Currency, rate As Double, net As Currency, tax As Currency
    If Not ControlExists(frm, "lblPriceInfo") Then Exit Sub
    price = Nz(frm!SellingPrice.Value, 0)
    rate = Nz(SettingValue("VATRate"), 0.15)
    If Nz(frm!VATCategory.Value, "S") <> "S" Then rate = 0
    If Nz(SettingValue("PricesIncludeVAT"), True) Then
        tax = RoundMoney(price * rate / (1 + rate))
        net = price - tax
    Else
        net = price
        tax = RoundMoney(price * rate)
    End If
    frm!lblPriceInfo.Caption = "«·”⁄— »œÊ‰ ÷—Ì»…: " & Format$(net, "#,##0.00") & _
                               "    «·÷—Ì»…: " & Format$(tax, "#,##0.00") & _
                               "    «·”⁄— ··⁄„Ì·: " & Format$(net + tax, "#,##0.00")
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Public Function TagValue(ByVal frm As Access.Form, ByVal Key As String) As String
    Dim part As Variant, p As Long
    For Each part In Split(frm.Tag, "|")
        p = InStr(part, "=")
        If p > 0 Then
            If StrComp(Left$(part, p - 1), Key, vbTextCompare) = 0 Then
                TagValue = Mid$(part, p + 1)
                Exit Function
            End If
        End If
    Next
End Function

Public Function ControlExists(ByVal frm As Access.Form, ByVal ControlName As String) As Boolean
    Dim ctl As Access.Control
    On Error Resume Next
    Set ctl = frm.Controls(ControlName)
    ControlExists = Not ctl Is Nothing
End Function

Private Function HasRecordField(ByVal frm As Access.Form, ByVal FieldName As String) As Boolean
    Dim fld As DAO.Field
    On Error Resume Next
    Set fld = frm.Recordset.Fields(FieldName)
    HasRecordField = Not fld Is Nothing
End Function

Private Function FieldIsRequired(ByVal frm As Access.Form, ByVal FieldName As String) As Boolean
    Dim fld As DAO.Field
    On Error Resume Next
    Set fld = frm.Recordset.Fields(FieldName)
    If fld Is Nothing Then Exit Function
    If fld.Type = dbBoolean Then Exit Function
    If (fld.Attributes And dbAutoIncrField) <> 0 Then Exit Function
    FieldIsRequired = fld.Required
End Function

Private Sub LockForActions(ByVal frm As Access.Form)
    ' A user with his own screen permissions and no "edit" in this screen sees the saved records
    ' read-only; a new record (when he may add) stays open for typing.
    Dim ctl As Access.Control, names As String, part As Variant, readOnly As Boolean
    If Not UsesCustomScreens() Then Exit Sub
    If m_editable Is Nothing Then Set m_editable = CreateObject("Scripting.Dictionary")
    If Not m_editable.Exists(frm.Name) Then
        names = "|"
        For Each ctl In frm.Controls
            If Len(BoundField(ctl)) > 0 Then
                If Not ctl.Locked Then names = names & ctl.Name & "|"
            End If
        Next
        m_editable(frm.Name) = names
    End If
    readOnly = Not frm.NewRecord And Not CanScreenAction(frm.Name, "EDIT", True)
    For Each part In Split(m_editable(frm.Name), "|")
        If Len(part) > 0 Then frm.Controls(CStr(part)).Locked = readOnly
    Next
End Sub

Private Function BoundField(ByVal ctl As Access.Control) As String
    ' Name of the table field bound to the control, "" for unbound/calculated controls.
    Dim src As String
    On Error Resume Next
    src = ctl.ControlSource
    If Len(src) > 0 And Left$(src, 1) <> "=" Then BoundField = src
End Function

Private Function CaptionOf(ByVal ctl As Access.Control) As String
    On Error Resume Next
    CaptionOf = ctl.Name
    CaptionOf = Replace(ctl.Controls(0).Caption, " *", "")
End Function

Private Sub SetStatus(ByVal frm As Access.Form, ByVal Text As String, ByVal Color As Long)
    If Not ControlExists(frm, "lblStatus") Then Exit Sub
    If Len(Text) = 0 Then Text = " "      ' an empty label caption is not allowed
    frm!lblStatus.Caption = Text
    frm!lblStatus.ForeColor = Color
End Sub

Private Sub SetLocked(ByVal ctl As Access.Control, ByVal IsLocked As Boolean)
    ctl.Locked = IsLocked
    On Error Resume Next      ' check boxes have no BackColor
    ctl.BackColor = IIf(IsLocked, CLR_LOCKED, CLR_SURFACE)
End Sub

Private Sub SetEnabled(ByVal ctl As Access.Control, ByVal IsEnabled As Boolean)
    On Error Resume Next      ' a control cannot be disabled while it has the focus
    ctl.Enabled = IsEnabled
End Sub

Public Sub SafeFocus(ByVal ctl As Access.Control)
    On Error Resume Next      ' fails on hidden forms (tests) - harmless
    ctl.SetFocus
End Sub

Private Sub FocusFirstInput(ByVal frm As Access.Form)
    Dim ctl As Access.Control, best As Access.Control
    For Each ctl In frm.Controls
        If Len(BoundField(ctl)) > 0 Then
            If ctl.Enabled And Not ctl.Locked And ctl.TabStop Then
                If best Is Nothing Then
                    Set best = ctl
                ElseIf ctl.TabIndex < best.TabIndex Then
                    Set best = ctl
                End If
            End If
        End If
    Next
    If Not best Is Nothing Then SafeFocus best
End Sub
