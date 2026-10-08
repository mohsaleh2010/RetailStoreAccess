Attribute VB_Name = "modBank"
'==============================================================================
' modBank  -  Retail Store Management System
'
' Banks (table Banks, screen frmBanks): each bank account has its account in the
' chart 120000 + BankID under 1210 "bank accounts".
'   Bank transfers (payment method 3) are booked in the default bank
'   (Settings.DefaultBankID): BankFor gives the bank a document pays into / out of.
'   Mada and wallets stay in 1200 until the bank settles them.
' Bank transactions (BankTransactions, screen frmBankTx, journal source BANK_TX):
'   DEPOSIT     cash box -> bank            WITHDRAW   bank -> cash box
'   SETTLEMENT  Mada settlement: the collections leave 1200, the bank gets them
'               less the fee (5610) and its VAT (1500)
'   TRANSFER    bank -> bank                OTHER_IN / OTHER_OUT  with the account
'               chosen (interest, loans, owner, bank charges and their VAT)
' Bank reconciliation (BankReconciliations, BankClearings, screen frmBankRecon):
'   the operations on the bank account in the books are ticked when they appear in
'   the bank statement; book balance - operations not in the statement yet must
'   equal the balance of the statement.
'==============================================================================
Option Compare Database
Option Explicit

Public Const BANK_TRANSFER_METHOD_ID As Long = 3      ' PaymentMethods:  ÕÊÌ· »‰ﬂÌ
Public Const MADA_CLEARING As Long = 1200
Public Const BANK_CHARGES As Long = 5610

Public Function BankAccount(ByVal BankID As Long) As Long
    BankAccount = 120000 + BankID
End Function

Public Function BankFor(ByVal PaymentMethodID As Variant, ByVal Amount As Currency) As Variant
    ' The bank a document pays into / out of: a bank transfer goes to the default bank (Null = none).
    Dim bank As Variant
    BankFor = Null
    If Nz(PaymentMethodID, 0) <> BANK_TRANSFER_METHOD_ID Or Amount = 0 Then Exit Function
    bank = SettingValue("DefaultBankID")
    If IsNull(bank) Then Exit Function
    If Nz(DbValue("SELECT COUNT(*) FROM Banks WHERE IsActive = True AND BankID = " & CLng(bank)), 0) > 0 Then BankFor = CLng(bank)
End Function

Public Function BankBookBalance(ByVal BankID As Long, Optional ByVal AsOf As Variant) As Currency
    ' The balance of the bank in the books (all its journal lines, or until a day included).
    Dim sql As String
    If IsMissing(AsOf) Then AsOf = Null
    sql = "SELECT Sum(ItemAmount) FROM qryBankItems WHERE BankID = " & BankID
    If Not IsNull(AsOf) Then sql = sql & " AND ItemDate < " & SqlDate(DateValue(AsOf) + 1)
    BankBookBalance = Nz(DbValue(sql), 0)
End Function

Public Function MadaPending() As Currency
    ' Mada and wallet collections not settled by the bank yet.
    MadaPending = AccountBalance(MADA_CLEARING)
End Function

'------------------------------------------------------------------------------
' Bank transactions
'------------------------------------------------------------------------------
Private Function ActiveBank(ByVal BankID As Variant) As Boolean
    If IsNull(BankID) Then Exit Function
    ActiveBank = Nz(DbValue("SELECT COUNT(*) FROM Banks WHERE IsActive = True AND BankID = " & CLng(BankID)), 0) > 0
End Function

Public Function CounterAccountProblem(ByVal AccountCode As Variant) As String
    ' Other bank transactions: any active sub-account except the cash boxes, the banks and Mada (they have
    ' their own transactions) and customers / suppliers (their balances come from their documents).
    Dim ok As Variant
    If IsNull(AccountCode) Then
        CounterAccountProblem = "«Œ — «·Õ”«» «·„ﬁ«»·."
        Exit Function
    End If
    ok = DbValue("SELECT AccountCode FROM Accounts WHERE AccountCode = " & CLng(AccountCode) & " AND IsPosting = True " & _
                 "AND IsActive = True AND Nz(Level3Code, 0) NOT IN (1100, 1210) AND AccountCode NOT IN (1190, 1200, 1300, 2100)")
    If IsNull(ok) Then
        CounterAccountProblem = "«·Õ”«» «·„ﬁ«»· Õ”«» ›—⁄Ì ‰‘ÿ° Ê·Ì” ’‰œÊﬁ« Ê·« »‰ﬂ« Ê·« „œÏ Ê·« «·⁄„·«¡ √Ê «·„Ê—œÌ‰." & vbCrLf & _
                                "··≈Ìœ«⁄ Ê«·”Õ» Ê«· ÕÊÌ· Ê ”ÊÌ… „œÏ «Œ — ‰Ê⁄ «·Õ—ﬂ… «·Œ«’ »Â«."
    End If
End Function

Public Function BankTxProblem(ByVal TxType As String, ByVal TxDate As Variant, ByVal BankID As Variant, _
                              ByVal ToBankID As Variant, ByVal CashBoxID As Variant, ByVal CounterAccount As Variant, _
                              ByVal Amount As Currency, ByVal FeeAmount As Currency, ByVal FeeVAT As Currency) As String
    Dim p As String
    If Not IsDate(TxDate) Then
        p = "«ﬂ »  «—ÌŒ «·Õ—ﬂ…."
    ElseIf DateValue(TxDate) > Date Then
        p = " «—ÌŒ «·Õ—ﬂ… »⁄œ «·ÌÊ„."
    ElseIf Not ActiveBank(BankID) Then
        p = "«Œ — «·»‰ﬂ."
    ElseIf Amount <= 0 Then
        p = "«·„»·€ ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—."
    ElseIf FeeAmount < 0 Or FeeVAT < 0 Then
        p = "«·⁄„Ê·… Ê÷—Ì» Â« ·«  ﬂÊ‰«‰ ”«·» Ì‰."
    Else
        Select Case TxType
            Case "DEPOSIT", "WITHDRAW"
                If IsNull(CashBoxID) Then
                    p = "«Œ — «·’‰œÊﬁ."
                ElseIf Not BoxIsActive(CashBoxID) Then
                    p = "«·’‰œÊﬁ €Ì— ‰‘ÿ."
                ElseIf TxType = "DEPOSIT" And CashBoxBalance(CLng(CashBoxID), DateValue(TxDate) + 1) < Amount Then
                    p = "—’Ìœ «·’‰œÊﬁ ›Ì " & GDate(TxDate) & " √ﬁ· „‰ «·„»·€ «·„ÊœÛ⁄."
                End If
            Case "SETTLEMENT"
                If FeeAmount + FeeVAT >= Amount Then p = "«·⁄„Ê·… Ê÷—Ì» Â« √ﬁ· „‰ „»·€ «· ”ÊÌ…."
            Case "TRANSFER"
                If Not ActiveBank(ToBankID) Then
                    p = "«Œ — «·»‰ﬂ «·„ÕÊÛ¯· ≈·ÌÂ."
                ElseIf CLng(ToBankID) = CLng(BankID) Then
                    p = "«· ÕÊÌ· »Ì‰ »‰ﬂÌ‰ „Œ ·›Ì‰."
                End If
            Case "OTHER_IN", "OTHER_OUT"
                p = CounterAccountProblem(CounterAccount)
                If Len(p) = 0 And TxType = "OTHER_OUT" And FeeVAT >= Amount And FeeVAT > 0 Then
                    p = "÷—Ì»… «·„œŒ·«  √ﬁ· „‰ «·„»·€."
                End If
            Case Else
                p = "«Œ — ‰Ê⁄ «·Õ—ﬂ…."
        End Select
    End If
    If Len(p) = 0 Then p = ClosedPeriodProblem(TxDate)
    BankTxProblem = p
End Function

Public Function PostBankTx(ByVal TxType As String, ByVal TxDate As Variant, ByVal BankID As Variant, _
                           ByVal ToBankID As Variant, ByVal CashBoxID As Variant, ByVal CounterAccount As Variant, _
                           ByVal Amount As Currency, ByVal FeeAmount As Currency, ByVal FeeVAT As Currency, _
                           ByVal Reference As String, ByVal Description As String, ByRef NewID As Long) As String
    Dim rs As DAO.Recordset, number As String
    NewID = 0
    If Not HasPermission("BANKS") Then
        PostBankTx = "·«  „·ﬂ ’·«ÕÌ… «·»‰Êﬂ."
        Exit Function
    End If
    If Not CanScreenAction("frmBankTx", "ADD", True) Then
        PostBankTx = "·«  „·ﬂ ’·«ÕÌ… «·≈÷«›… ›Ì ‘«‘… «·Õ—ﬂ«  «·»‰ﬂÌ…."
        Exit Function
    End If
    PostBankTx = BankTxProblem(TxType, TxDate, BankID, ToBankID, CashBoxID, CounterAccount, Amount, FeeAmount, FeeVAT)
    If Len(PostBankTx) > 0 Then Exit Function
    number = NextNumber("BANK_TX")
    Set rs = CurrentDb.OpenRecordset("BankTransactions", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!TxNumber = number
    rs!TxDate = DateValue(TxDate)
    rs!TxType = TxType
    rs!BankID = CLng(BankID)
    If TxType = "TRANSFER" Then rs!ToBankID = CLng(ToBankID)
    If TxType = "DEPOSIT" Or TxType = "WITHDRAW" Then rs!CashBoxID = CLng(CashBoxID)
    If TxType = "OTHER_IN" Or TxType = "OTHER_OUT" Then rs!CounterAccount = CLng(CounterAccount)
    rs!Amount = Amount
    rs!FeeAmount = IIf(TxType = "SETTLEMENT", FeeAmount, 0)
    rs!FeeVAT = IIf(TxType = "SETTLEMENT" Or TxType = "OTHER_OUT", FeeVAT, 0)
    If Len(Trim$(Reference)) > 0 Then rs!Reference = Left$(Trim$(Reference), 40)
    rs!Description = Left$(IIf(Len(Trim$(Description)) > 0, Trim$(Description), BankTxTypeName(TxType)), 255)
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Bookmark = rs.LastModified
    NewID = rs!BankTxID
    rs.Close
    LogAction "BANK_TX", "BankTransactions", number, TxType & " " & Format$(Amount, "0.00")
    PostBankTx = SyncJournal()
End Function

Public Function DeleteBankTx(ByVal BankTxID As Long) As String
    Dim d As Variant
    If Not HasPermission("BANKS") Or Not CanScreenAction("frmBankTx", "DELETE", True) Then
        DeleteBankTx = "·«  „·ﬂ ’·«ÕÌ… Õ–› «·Õ—ﬂ«  «·»‰ﬂÌ…."
        Exit Function
    End If
    d = DbValue("SELECT TxDate FROM BankTransactions WHERE BankTxID = " & BankTxID)
    If IsNull(d) Then
        DeleteBankTx = "«·Õ—ﬂ… €Ì— „ÊÃÊœ…."
        Exit Function
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM BankClearings WHERE SourceType = 'BANK_TX' AND SourceID = " & BankTxID), 0) > 0 Then
        DeleteBankTx = "«·Õ—ﬂ… „ÿ«»ﬁ… ›Ì  ”ÊÌ… »‰ﬂÌ…: √·€ˆ „ÿ«»ﬁ Â« √Ê·«."
        Exit Function
    End If
    DeleteBankTx = ClosedPeriodProblem(d)
    If Len(DeleteBankTx) > 0 Then Exit Function
    Dim auditBefore As Collection
    Set auditBefore = AuditSnapshot("BankTransactions", "BankTxID", BankTxID)        ' the record as it was (modAudit)
    CurrentDb.Execute "DELETE FROM BankTransactions WHERE BankTxID = " & BankTxID, dbFailOnError
    AuditDeleted "BANK_TX_DELETE", "BankTransactions", BankTxID, auditBefore
    DeleteBankTx = SyncJournal()
End Function

Public Function BankTxTypeName(ByVal TxType As String) As String
    Select Case TxType
        Case "DEPOSIT": BankTxTypeName = "≈Ìœ«⁄ ‰ﬁœÌ… ›Ì «·»‰ﬂ"
        Case "WITHDRAW": BankTxTypeName = "”Õ» ‰ﬁœÌ… „‰ «·»‰ﬂ"
        Case "SETTLEMENT": BankTxTypeName = " ”ÊÌ…  Õ’Ì·«  „œÏ"
        Case "TRANSFER": BankTxTypeName = " ÕÊÌ· »Ì‰ «·»‰Êﬂ"
        Case "OTHER_IN": BankTxTypeName = "Ê«—œ ¬Œ— ··»‰ﬂ"
        Case "OTHER_OUT": BankTxTypeName = "’«œ— ¬Œ— „‰ «·»‰ﬂ"
    End Select
End Function

'------------------------------------------------------------------------------
' Bank reconciliation
'------------------------------------------------------------------------------
Private Function ReconField(ByVal ReconID As Long, ByVal FieldName As String) As Variant
    ReconField = DbValue("SELECT " & FieldName & " FROM BankReconciliations WHERE ReconciliationID = " & ReconID)
End Function

Public Function ReconFigures(ByVal ReconID As Long, ByRef Book As Currency, ByRef Outstanding As Currency, _
                             ByRef Difference As Currency) As Boolean
    ' Book balance on the statement day, the operations until that day not in a statement yet, and the
    ' difference: statement balance - (book - outstanding). False when the reconciliation does not exist.
    Dim bank As Variant, d As Variant
    bank = ReconField(ReconID, "BankID")
    If IsNull(bank) Then Exit Function
    d = ReconField(ReconID, "StatementDate")
    Book = BankBookBalance(bank, d)
    Outstanding = Nz(DbValue("SELECT Sum(ItemAmount) FROM qryBankItems WHERE BankID = " & bank & " AND IsCleared = 0 " & _
                             "AND ItemDate < " & SqlDate(DateValue(d) + 1)), 0)
    Difference = Nz(ReconField(ReconID, "StatementBalance"), 0) - (Book - Outstanding)
    ReconFigures = True
End Function

Public Function OpenReconciliation(ByVal BankID As Long, ByVal StatementDate As Variant, _
                                   ByVal StatementBalance As Currency, ByRef ReconID As Long) As String
    ' The open reconciliation of the bank (a new one, or the open one with this date and balance).
    Dim lastDone As Variant, rs As DAO.Recordset
    ReconID = 0
    If Not HasPermission("BANKS") Then
        OpenReconciliation = "·«  „·ﬂ ’·«ÕÌ… «·»‰Êﬂ."
        Exit Function
    End If
    If Not ActiveBank(BankID) Then
        OpenReconciliation = "«Œ — «·»‰ﬂ."
        Exit Function
    End If
    If Not IsDate(StatementDate) Then
        OpenReconciliation = "«ﬂ »  «—ÌŒ ﬂ‘› «·»‰ﬂ."
        Exit Function
    End If
    If DateValue(StatementDate) > Date Then
        OpenReconciliation = " «—ÌŒ «·ﬂ‘› »⁄œ «·ÌÊ„."
        Exit Function
    End If
    lastDone = DbValue("SELECT Max(StatementDate) FROM BankReconciliations WHERE Status = 'DONE' AND BankID = " & BankID)
    If Not IsNull(lastDone) Then
        If DateValue(StatementDate) < DateValue(lastDone) Then
            OpenReconciliation = "¬Œ—  ”ÊÌ… „⁄ „œ… ·Â–« «·»‰ﬂ ›Ì " & GDate(lastDone) & ": «Œ —  «—ÌŒ« »⁄œÂ."
            Exit Function
        End If
    End If
    ReconID = Nz(DbValue("SELECT ReconciliationID FROM BankReconciliations WHERE Status = 'OPEN' AND BankID = " & BankID), 0)
    If ReconID = 0 Then
        If Not CanScreenAction("frmBankRecon", "ADD", True) Then
            OpenReconciliation = "·«  „·ﬂ ’·«ÕÌ… «·≈÷«›… ›Ì ‘«‘… «· ”ÊÌ… «·»‰ﬂÌ…."
            Exit Function
        End If
        Set rs = CurrentDb.OpenRecordset("BankReconciliations", dbOpenDynaset, dbAppendOnly)
        rs.AddNew
        rs!ReconNumber = NextNumber("BANK_RECON")
        rs!BankID = BankID
        rs!StatementDate = DateValue(StatementDate)
        rs!StatementBalance = StatementBalance
        rs!Status = "OPEN"
        rs!EmployeeID = CurrentUserID()
        rs.Update
        rs.Bookmark = rs.LastModified
        ReconID = rs!ReconciliationID
        rs.Close
    Else
        CurrentDb.Execute "UPDATE BankReconciliations SET StatementDate = " & SqlDate(DateValue(StatementDate)) & _
                          ", StatementBalance = " & Str$(StatementBalance) & " WHERE ReconciliationID = " & ReconID, dbFailOnError
    End If
End Function

Private Function ReconIsOpen(ByVal ReconID As Long) As Boolean
    ReconIsOpen = (Nz(ReconField(ReconID, "Status"), "") = "OPEN")
End Function

Public Function ClearBankItem(ByVal ReconID As Long, ByVal SourceType As String, ByVal SourceID As Long) As String
    ' The operation appears in the bank statement of this reconciliation.
    Dim bank As Long, amount As Variant
    If Not ReconIsOpen(ReconID) Then
        ClearBankItem = "«· ”ÊÌ… €Ì— „ÊÃÊœ… √Ê „⁄ „œ…."
        Exit Function
    End If
    bank = ReconField(ReconID, "BankID")
    amount = DbValue("SELECT ItemAmount FROM qryBankItems WHERE BankID = " & bank & " AND SourceType = " & SqlText(SourceType) & _
                     " AND SourceID = " & SourceID & " AND IsCleared = 0")
    If IsNull(amount) Then
        ClearBankItem = "«·⁄„·Ì… €Ì— „ÊÃÊœ… √Ê „ÿ«»ﬁ… „‰ ﬁ»·."
        Exit Function
    End If
    CurrentDb.Execute "INSERT INTO BankClearings (ReconciliationID, BankID, SourceType, SourceID, ClearedAmount) VALUES (" & _
                      ReconID & ", " & bank & ", " & SqlText(SourceType) & ", " & SourceID & ", " & Str$(CCur(amount)) & ")", _
                      dbFailOnError
End Function

Public Function UnclearBankItem(ByVal ReconID As Long, ByVal SourceType As String, ByVal SourceID As Long) As String
    If Not ReconIsOpen(ReconID) Then
        UnclearBankItem = "«· ”ÊÌ… €Ì— „ÊÃÊœ… √Ê „⁄ „œ…."
        Exit Function
    End If
    CurrentDb.Execute "DELETE FROM BankClearings WHERE ReconciliationID = " & ReconID & " AND SourceType = " & _
                      SqlText(SourceType) & " AND SourceID = " & SourceID, dbFailOnError
End Function

Public Function FinishReconciliation(ByVal ReconID As Long) As String
    Dim book As Currency, outstanding As Currency, diff As Currency
    If Not ReconIsOpen(ReconID) Then
        FinishReconciliation = "«· ”ÊÌ… €Ì— „ÊÃÊœ… √Ê „⁄ „œ…."
        Exit Function
    End If
    If Not CanScreenAction("frmBankRecon", "EDIT", True) Then
        FinishReconciliation = "·«  „·ﬂ ’·«ÕÌ… «⁄ „«œ «· ”ÊÌ… «·»‰ﬂÌ…."
        Exit Function
    End If
    ReconFigures ReconID, book, outstanding, diff
    If diff <> 0 Then
        FinishReconciliation = "ÌÊÃœ ›—ﬁ " & Format$(diff, "#,##0.00") & " »Ì‰ ﬂ‘› «·»‰ﬂ Ê«·œ›« —." & vbCrLf & _
                               "ÿ«»ﬁ «·⁄„·Ì«  «·Ÿ«Â—… ›Ì «·ﬂ‘›° Ê”Ã¯· „« ›Ì «·ﬂ‘› Ê·„ Ìı”ÃÛ¯· (⁄„Ê·« ° ›Ê«∆œ...) ﬂÕ—ﬂ… »‰ﬂÌ…."
        Exit Function
    End If
    CurrentDb.Execute "UPDATE BankReconciliations SET Status = 'DONE', BookBalance = " & Str$(book) & ", Outstanding = " & _
                      Str$(outstanding) & ", Difference = 0, EmployeeID = " & CurrentUserID() & _
                      " WHERE ReconciliationID = " & ReconID, dbFailOnError
    LogAction "BANK_RECON", "BankReconciliations", CStr(ReconID)
End Function

Public Function ReopenReconciliation(ByVal ReconID As Long) As String
    Dim bank As Variant, lastID As Variant
    bank = ReconField(ReconID, "BankID")
    If IsNull(bank) Then
        ReopenReconciliation = "«· ”ÊÌ… €Ì— „ÊÃÊœ…."
        Exit Function
    End If
    If Not CanScreenAction("frmBankRecon", "EDIT", True) Then
        ReopenReconciliation = "·«  „·ﬂ ’·«ÕÌ…  ⁄œÌ· «· ”ÊÌ… «·»‰ﬂÌ…."
        Exit Function
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM BankReconciliations WHERE Status = 'OPEN' AND BankID = " & bank), 0) > 0 Then
        ReopenReconciliation = " ÊÃœ  ”ÊÌ… Ã«—Ì… ·Â–« «·»‰ﬂ: «⁄ „œÂ« √Ê «Õ–›Â« √Ê·«."
        Exit Function
    End If
    lastID = DbValue("SELECT TOP 1 ReconciliationID FROM BankReconciliations WHERE BankID = " & bank & _
                     " ORDER BY StatementDate DESC, ReconciliationID DESC")
    If Nz(lastID, 0) <> ReconID Then
        ReopenReconciliation = " ı⁄«œ › Õ ¬Œ—  ”ÊÌ… ··»‰ﬂ ›ﬁÿ."
        Exit Function
    End If
    CurrentDb.Execute "UPDATE BankReconciliations SET Status = 'OPEN' WHERE ReconciliationID = " & ReconID, dbFailOnError
    LogAction "BANK_RECON_REOPEN", "BankReconciliations", CStr(ReconID)
End Function

Public Function DeleteReconciliation(ByVal ReconID As Long) As String
    If Not ReconIsOpen(ReconID) Then
        DeleteReconciliation = " ıÕ–› «· ”ÊÌ… «·Ã«—Ì… ›ﬁÿ (√⁄œ › Õ «·„⁄ „œ… √Ê·«)."
        Exit Function
    End If
    Dim auditBefore As Collection
    Set auditBefore = AuditSnapshot("BankReconciliations", "ReconciliationID", ReconID)        ' the record as it was (modAudit)
    CurrentDb.Execute "DELETE FROM BankReconciliations WHERE ReconciliationID = " & ReconID, dbFailOnError   ' clearings cascade
    AuditDeleted "BANK_RECON_DELETE", "BankReconciliations", ReconID, auditBefore
End Function

Public Function ChangedClearings(ByVal BankID As Long) As Long
    ' Operations ticked in a statement whose amount in the books changed since.
    ChangedClearings = Nz(DbValue("SELECT COUNT(*) FROM qryBankItems WHERE BankID = " & BankID & _
                                  " AND IsCleared = 1 AND ClearedAmount <> ItemAmount"), 0)
End Function

'------------------------------------------------------------------------------
' Screen frmBankTx (OpenArgs: the bank)
'------------------------------------------------------------------------------
Public Sub BankTxLoad(ByVal frm As Access.Form)
    Dim msg As String
    Calendar = vbCalGreg
    msg = SyncJournal()
    If Len(msg) > 0 Then ShowWarning msg
    frm!cboType.Value = "DEPOSIT"
    frm!txtDate.Value = Date
    If Nz(frm.OpenArgs, 0) > 0 Then
        frm!cboBank.Value = CLng(frm.OpenArgs)
    Else
        frm!cboBank.Value = SettingValue("DefaultBankID")
    End If
    BankTxTypeChanged frm
End Sub

Public Sub BankTxTypeChanged(ByVal frm As Access.Form)
    Dim t As String
    t = Nz(frm!cboType.Value, "")
    frm!cboToBank.Enabled = (t = "TRANSFER")
    frm!cboBox.Enabled = (t = "DEPOSIT" Or t = "WITHDRAW")
    frm!cboCounter.Enabled = (t = "OTHER_IN" Or t = "OTHER_OUT")
    frm!txtFee.Enabled = (t = "SETTLEMENT")
    frm!txtFeeVAT.Enabled = (t = "SETTLEMENT" Or t = "OTHER_OUT")
    If t = "SETTLEMENT" And Nz(frm!txtAmount.Value, 0) = 0 And MadaPending() > 0 Then frm!txtAmount.Value = MadaPending()
    If t = "OTHER_OUT" And IsNull(frm!cboCounter.Value) Then frm!cboCounter.Value = BANK_CHARGES
    BankTxRefresh frm
End Sub

Public Sub BankTxRefresh(ByVal frm As Access.Form)
    Dim info As String, net As Currency
    If Not IsNull(frm!cboBank.Value) Then
        info = "—’Ìœ «·»‰ﬂ ›Ì «·œ›« —: " & Format$(BankBookBalance(frm!cboBank.Value), "#,##0.00")
    End If
    info = info & "     Õ’Ì·«  „œÏ ·„  ı”ÊÛ¯: " & Format$(MadaPending(), "#,##0.00")
    If Not IsNull(frm!cboBox.Value) And frm!cboBox.Enabled Then
        info = info & "    —’Ìœ «·’‰œÊﬁ: " & Format$(CashBoxBalance(frm!cboBox.Value), "#,##0.00")
    End If
    If frm!cboType.Value = "SETTLEMENT" Then
        net = Nz(frm!txtAmount.Value, 0) - Nz(frm!txtFee.Value, 0) - Nz(frm!txtFeeVAT.Value, 0)
        info = info & vbCrLf & "’«›Ì „« ÌœŒ· «·»‰ﬂ: " & Format$(net, "#,##0.00")
    End If
    frm!lblInfo.Caption = Tr(info)
    frm!lstTx.Requery
End Sub

Public Sub SaveBankTx(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If frm!cboType.Value = "SETTLEMENT" And Nz(frm!txtAmount.Value, 0) > MadaPending() Then
        If Not AskYesNo("«·„»·€ √ﬂ»— „‰  Õ’Ì·«  „œÏ «· Ì ·„  ı”ÊÛ¯ (" & Format$(MadaPending(), "#,##0.00") & ")." & vbCrLf & _
                        "Â·  —Ìœ «·Õ›Ÿ ⁄·Ï √Ì Õ«·ø") Then Exit Sub
    End If
    msg = PostBankTx(Nz(frm!cboType.Value, ""), frm!txtDate.Value, frm!cboBank.Value, frm!cboToBank.Value, frm!cboBox.Value, _
                     frm!cboCounter.Value, CCur(Nz(frm!txtAmount.Value, 0)), CCur(Nz(frm!txtFee.Value, 0)), _
                     CCur(Nz(frm!txtFeeVAT.Value, 0)), Nz(frm!txtReference.Value, ""), Nz(frm!txtDescription.Value, ""), id)
    If id = 0 Then
        ShowWarning msg
        Exit Sub
    End If
    If Len(msg) > 0 Then ShowWarning msg
    ShowInfo " „ Õ›Ÿ «·Õ—ﬂ… " & DbValue("SELECT TxNumber FROM BankTransactions WHERE BankTxID = " & id) & "."
    frm!txtAmount.Value = Null
    frm!txtFee.Value = Null
    frm!txtFeeVAT.Value = Null
    frm!txtReference.Value = Null
    frm!txtDescription.Value = Null
    BankTxRefresh frm
End Sub

Public Sub DeleteSelectedBankTx(ByVal frm As Access.Form)
    Dim msg As String
    If IsNull(frm!lstTx.Value) Then
        ShowWarning "«Œ — «·Õ—ﬂ… „‰ «·ﬁ«∆„…."
        Exit Sub
    End If
    If Not AskYesNo("Õ–› «·Õ—ﬂ… «·»‰ﬂÌ… «·„Õœœ… ÊﬁÌœÂ«ø") Then Exit Sub
    msg = DeleteBankTx(frm!lstTx.Value)
    If Len(msg) > 0 Then ShowWarning msg
    BankTxRefresh frm
End Sub

'------------------------------------------------------------------------------
' Screen frmBankRecon (OpenArgs: the bank)
'------------------------------------------------------------------------------
Private Function ShownRecon(ByVal frm As Access.Form) As Long
    ShownRecon = Nz(frm!txtReconID.Value, 0)
End Function

Public Sub BankReconLoad(ByVal frm As Access.Form)
    Dim msg As String
    Calendar = vbCalGreg
    msg = SyncJournal()
    If Len(msg) > 0 Then ShowWarning msg
    If Nz(frm.OpenArgs, 0) > 0 Then
        frm!cboBank.Value = CLng(frm.OpenArgs)
    Else
        frm!cboBank.Value = SettingValue("DefaultBankID")
    End If
    BankReconBankChanged frm
End Sub

Public Sub BankReconBankChanged(ByVal frm As Access.Form)
    ' The open reconciliation of the bank, if there is one.
    Dim id As Variant
    frm!txtReconID.Value = Null
    If Not IsNull(frm!cboBank.Value) Then
        id = DbValue("SELECT ReconciliationID FROM BankReconciliations WHERE Status = 'OPEN' AND BankID = " & frm!cboBank.Value)
        If Not IsNull(id) Then
            frm!txtReconID.Value = id
            frm!txtStatementDate.Value = ReconField(id, "StatementDate")
            frm!txtStatementBalance.Value = ReconField(id, "StatementBalance")
        Else
            frm!txtStatementDate.Value = DateSerial(Year(Date), Month(Date), 0)
            frm!txtStatementBalance.Value = Null
        End If
        frm!lstRecons.RowSource = Tr("SELECT ReconciliationID, ReconNumber AS [«· ”ÊÌ…], Format(StatementDate, 'yyyy/mm/dd') " & _
            "AS [ «—ÌŒ «·ﬂ‘›], Format(StatementBalance, '#,##0.00') AS [—’Ìœ «·ﬂ‘›], IIf(r.Status = 'DONE', '„⁄ „œ…', " & _
            "'Ã«—Ì…') AS [Õ«·… «· ”ÊÌ…] FROM BankReconciliations AS r WHERE r.BankID = " & frm!cboBank.Value & " ORDER BY StatementDate DESC")
    End If
    BankReconRefresh frm
End Sub

Public Sub BankReconStart(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If IsNull(frm!cboBank.Value) Then
        ShowWarning "«Œ — «·»‰ﬂ."
        Exit Sub
    End If
    If IsNull(frm!txtStatementBalance.Value) Then
        ShowWarning "«ﬂ » «·—’Ìœ «·Ÿ«Â— ›Ì ﬂ‘› «·»‰ﬂ ›Ì Â–« «· «—ÌŒ."
        SafeFocus frm!txtStatementBalance
        Exit Sub
    End If
    msg = OpenReconciliation(frm!cboBank.Value, frm!txtStatementDate.Value, CCur(frm!txtStatementBalance.Value), id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    frm!txtReconID.Value = id
    frm!lstRecons.Requery
    BankReconRefresh frm
End Sub

Public Sub BankReconRefresh(ByVal frm As Access.Form)
    Dim id As Long, book As Currency, outstanding As Currency, diff As Currency, bank As Long, cutoff As String
    Dim changed As Long, state As String
    id = ShownRecon(frm)
    bank = Nz(frm!cboBank.Value, 0)
    frm!lstOpen.RowSource = Tr("")
    frm!lstCleared.RowSource = Tr("")
    If id = 0 Then
        frm!lblBook.Caption = "-": frm!lblOutstanding.Caption = "-": frm!lblAdjusted.Caption = "-": frm!lblDifference.Caption = "-"
        frm!lblState.Caption = Tr("«Œ — «·»‰ﬂ° Ê«ﬂ »  «—ÌŒ ﬂ‘› «·»‰ﬂ Ê—’ÌœÂ° À„ ´»œ¡ «· ”ÊÌ…ª.")
        frm!lblState.ForeColor = CLR_PRIMARY
        Exit Sub
    End If
    cutoff = SqlDate(DateValue(ReconField(id, "StatementDate")) + 1)
    frm!lstOpen.RowSource = Tr("SELECT SourceType & '|' & SourceID AS ItemKey, Format(ItemDate, 'yyyy/mm/dd') AS [«· «—ÌŒ], " & _
        "TypeName AS [«·⁄„·Ì…], ItemNumber AS [«·—ﬁ„], ItemText AS [«·»Ì«‰], IIf(ItemAmount > 0, Format(ItemAmount, " & _
        "'#,##0.00'), '') AS [Ê«—œ], IIf(ItemAmount < 0, Format(-ItemAmount, '#,##0.00'), '') AS [’«œ—] FROM qryBankItems " & _
        "WHERE BankID = " & bank & " AND IsCleared = 0 AND ItemDate < " & cutoff & " ORDER BY ItemDate, ItemNumber")
    frm!lstCleared.RowSource = Tr("SELECT SourceType & '|' & SourceID AS ItemKey, Format(ItemDate, 'yyyy/mm/dd') AS [«· «—ÌŒ], " & _
        "TypeName AS [«·⁄„·Ì…], ItemNumber AS [«·—ﬁ„], IIf(ItemAmount > 0, Format(ItemAmount, '#,##0.00'), '') AS [Ê«—œ], " & _
        "IIf(ItemAmount < 0, Format(-ItemAmount, '#,##0.00'), '') AS [’«œ—] FROM qryBankItems WHERE BankID = " & bank & _
        " AND ReconciliationID = " & id & " ORDER BY ItemDate, ItemNumber")
    ReconFigures id, book, outstanding, diff
    frm!lblBook.Caption = Tr(Format$(book, "#,##0.00"))
    frm!lblOutstanding.Caption = Tr(Format$(outstanding, "#,##0.00"))
    frm!lblAdjusted.Caption = Tr(Format$(book - outstanding, "#,##0.00"))
    frm!lblDifference.Caption = Tr(Format$(diff, "#,##0.00"))
    frm!lblDifference.ForeColor = IIf(diff = 0, CLR_SUCCESS, CLR_DANGER)
    state = "«· ”ÊÌ… " & ReconField(id, "ReconNumber") & IIf(ReconIsOpen(id), " Ã«—Ì…", " „⁄ „œ…") & _
            ": «Œ — «·⁄„·Ì«  «·Ÿ«Â—… ›Ì ﬂ‘› «·»‰ﬂ À„ ´„ÿ«»ﬁ…ª."
    changed = ChangedClearings(bank)
    If changed > 0 Then state = state & vbCrLf & changed & " ⁄„·Ì… „ÿ«»ﬁ…  €Ì¯— „»·€Â« ›Ì «·œ›« — »⁄œ «·„ÿ«»ﬁ…: —«Ã⁄Â«."
    frm!lblState.Caption = Tr(state)
    frm!lblState.ForeColor = IIf(changed > 0, CLR_WARNING, CLR_PRIMARY)
End Sub

Private Sub ApplyToSelected(ByVal frm As Access.Form, ByVal lst As Object, ByVal Clearing As Boolean)
    Dim v As Variant, parts() As String, msg As String, n As Long
    If ShownRecon(frm) = 0 Then
        ShowWarning "«»œ√ «· ”ÊÌ… √Ê·«."
        Exit Sub
    End If
    For Each v In lst.ItemsSelected
        parts = Split(lst.ItemData(v), "|")
        If Clearing Then
            msg = ClearBankItem(ShownRecon(frm), parts(0), CLng(parts(1)))
        Else
            msg = UnclearBankItem(ShownRecon(frm), parts(0), CLng(parts(1)))
        End If
        If Len(msg) > 0 Then Exit For
        n = n + 1
    Next
    If Len(msg) > 0 Then ShowWarning msg
    If n = 0 And Len(msg) = 0 Then ShowWarning "«Œ — ⁄„·Ì… √Ê √ﬂÀ— „‰ «·ﬁ«∆„…."
    BankReconRefresh frm
End Sub

Public Sub BankReconClear(ByVal frm As Access.Form)
    ApplyToSelected frm, frm!lstOpen, True
End Sub

Public Sub BankReconUnclear(ByVal frm As Access.Form)
    ApplyToSelected frm, frm!lstCleared, False
End Sub

Public Sub BankReconFinish(ByVal frm As Access.Form)
    Dim msg As String
    If ShownRecon(frm) = 0 Then Exit Sub
    msg = FinishReconciliation(ShownRecon(frm))
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        ShowInfo " „ «⁄ „«œ «· ”ÊÌ… «·»‰ﬂÌ…."
    End If
    frm!lstRecons.Requery
    BankReconRefresh frm
End Sub

Public Sub BankReconPick(ByVal frm As Access.Form)
    ' A reconciliation from the list (an approved one is shown read only).
    If IsNull(frm!lstRecons.Value) Then Exit Sub
    frm!txtReconID.Value = frm!lstRecons.Value
    frm!txtStatementDate.Value = ReconField(frm!lstRecons.Value, "StatementDate")
    frm!txtStatementBalance.Value = ReconField(frm!lstRecons.Value, "StatementBalance")
    BankReconRefresh frm
End Sub

Public Sub BankReconReopen(ByVal frm As Access.Form)
    Dim msg As String
    If ShownRecon(frm) = 0 Then Exit Sub
    If Not AskYesNo("≈⁄«œ… › Õ «· ”ÊÌ… «·„⁄ „œ…ø") Then Exit Sub
    msg = ReopenReconciliation(ShownRecon(frm))
    If Len(msg) > 0 Then ShowWarning msg
    frm!lstRecons.Requery
    BankReconRefresh frm
End Sub

Public Sub BankReconDelete(ByVal frm As Access.Form)
    Dim msg As String
    If ShownRecon(frm) = 0 Then Exit Sub
    If Not AskYesNo("Õ–› «· ”ÊÌ… «·Ã«—Ì… Êﬂ· „ÿ«»ﬁ« Â«ø") Then Exit Sub
    msg = DeleteReconciliation(ShownRecon(frm))
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        frm!txtReconID.Value = Null
    End If
    frm!lstRecons.Requery
    BankReconRefresh frm
End Sub

Public Sub BankReconAddTx(ByVal frm As Access.Form)
    ' A charge, interest... in the statement but not in the books.
    If IsNull(frm!cboBank.Value) Then Exit Sub
    OpenScreen "frmBankTx", 0, frm!cboBank.Value
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Private Sub CheckBank(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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

Public Function TestBank() As Boolean
    Dim passed As Long, failed As Long, report As String, msg As String, ws As DAO.Workspace, inTrans As Boolean
    Dim bank As Long, bank2 As Long, box As Long, id As Long, recon As Long, mada As Currency
    Dim book As Currency, outstanding As Currency, diff As Currency
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestBank  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    If ClosedThroughDate() >= Date Then
        Debug.Print "[--] «·› —… „ﬁ›·… Õ Ï «·ÌÊ„"
        GoTo Undo
    End If
    CurrentDb.Execute "INSERT INTO Banks (BankName, OpeningBalance, OpeningDate, IsActive) VALUES ('TEST »‰ﬂ 1', 1000, " & _
                      SqlDate(Date) & ", True)", dbFailOnError
    bank = DbValue("SELECT BankID FROM Banks WHERE BankName = 'TEST »‰ﬂ 1'")
    CurrentDb.Execute "INSERT INTO Banks (BankName, OpeningBalance, OpeningDate, IsActive) VALUES ('TEST »‰ﬂ 2', 0, " & _
                      SqlDate(Date) & ", True)", dbFailOnError
    bank2 = DbValue("SELECT BankID FROM Banks WHERE BankName = 'TEST »‰ﬂ 2'")
    msg = SyncJournal()
    CheckBank AccountBalance(BankAccount(bank)) = 1000 And BankBookBalance(bank) = 1000, _
              "Õ”«» «·»‰ﬂ ›Ì «·œ·Ì· »—’ÌœÂ «·«›  «ÕÌ " & msg, passed, failed, report

    box = CurrentCashBoxID()
    If box = 0 Then box = Nz(DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE IsActive = True"), 0)
    CurrentDb.Execute "INSERT INTO CashVouchers (VoucherNumber, VoucherDate, VoucherType, CashBoxID, Category, Amount, " & _
        "Description, EmployeeID) VALUES ('TEST-CIN', " & SqlDate(Date) & ", 'IN', " & box & ", 'OTHER', 500, 'TEST', " & _
        CurrentUserID() & ")", dbFailOnError
    msg = PostBankTx("DEPOSIT", Date, bank, Null, box, Null, 300, 0, 0, "", "TEST", id)
    CheckBank Len(msg) = 0 And id > 0 And BankBookBalance(bank) = 1300, "≈Ìœ«⁄ ‰ﬁœÌ… „‰ «·’‰œÊﬁ ›Ì «·»‰ﬂ " & msg, _
              passed, failed, report
    CheckBank Len(PostBankTx("DEPOSIT", Date, bank, Null, box, Null, 999999, 0, 0, "", "TEST", id)) > 0, _
              "·« ÌıÊœÛ⁄ √ﬂÀ— „‰ —’Ìœ «·’‰œÊﬁ", passed, failed, report

    mada = MadaPending()
    msg = PostBankTx("SETTLEMENT", Date, bank, Null, Null, Null, 100, 1.5, 0.23, "", "TEST", id)
    CheckBank Len(msg) = 0 And BankBookBalance(bank) = 1398.27 And MadaPending() = mada - 100, _
              " ”ÊÌ… „œÏ: «·’«›Ì ··»‰ﬂ Ê«·⁄„Ê·… „’—Ê› " & msg, passed, failed, report
    msg = PostBankTx("TRANSFER", Date, bank, bank2, Null, Null, 200, 0, 0, "", "TEST", id)
    CheckBank Len(msg) = 0 And BankBookBalance(bank2) = 200 And BankBookBalance(bank) = 1198.27, _
              " ÕÊÌ· »Ì‰ »‰ﬂÌ‰ " & msg, passed, failed, report
    CheckBank Len(PostBankTx("OTHER_OUT", Date, bank, Null, Null, 1300, 10, 0, 0, "", "TEST", id)) > 0, _
              "«·Õ—ﬂ… «·√Œ—Ï ·«  ﬂÊ‰ ⁄·Ï Õ”«» «·⁄„·«¡", passed, failed, report
    msg = PostBankTx("OTHER_OUT", Date, bank, Null, Null, BANK_CHARGES, 11.5, 0, 1.5, "", "TEST", id)
    CheckBank Len(msg) = 0 And BankBookBalance(bank) = 1186.77, "—”Ê„ »‰ﬂÌ… „⁄ ÷—Ì» Â« " & msg, passed, failed, report
    CheckBank Nz(DbValue("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0) = 0, _
              "ﬂ· «·ﬁÌÊœ „ Ê«“‰…", passed, failed, report

    ' reconciliation: the statement shows the opening, the deposit and the settlement only
    msg = OpenReconciliation(bank, Date, 1398.27, recon)
    CheckBank Len(msg) = 0 And recon > 0, "»œ¡ «· ”ÊÌ… «·»‰ﬂÌ… " & msg, passed, failed, report
    msg = ClearBankItem(recon, "BANK_OPENING", bank)
    msg = msg & ClearBankItem(recon, "BANK_TX", DbValue("SELECT BankTxID FROM BankTransactions WHERE TxType = 'DEPOSIT' AND BankID = " & bank))
    msg = msg & ClearBankItem(recon, "BANK_TX", DbValue("SELECT BankTxID FROM BankTransactions WHERE TxType = 'SETTLEMENT' AND BankID = " & bank))
    ReconFigures recon, book, outstanding, diff
    CheckBank Len(msg) = 0 And book = 1186.77 And outstanding = -211.5 And diff = 0, _
              "«·œ›« — - «·Õ—ﬂ«  €Ì— «·Ÿ«Â—… = —’Ìœ «·ﬂ‘› " & msg, passed, failed, report
    CheckBank Len(FinishReconciliation(recon)) = 0 And Not ReconIsOpen(recon), "«⁄ „«œ «· ”ÊÌ… »·« ›—ﬁ", passed, failed, report
    CheckBank Len(DeleteBankTx(DbValue("SELECT BankTxID FROM BankTransactions WHERE TxType = 'DEPOSIT' AND BankID = " & bank))) > 0, _
              "·«  ıÕ–› Õ—ﬂ… „ÿ«»ﬁ…", passed, failed, report
Undo:
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckBank False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·»‰Êﬂ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestBank"
        TestBank = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestBank"
    End If
End Function
