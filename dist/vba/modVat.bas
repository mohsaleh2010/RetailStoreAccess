Attribute VB_Name = "modVat"
'==============================================================================
' modVat  -  Retail Store Management System
'
' The VAT return (screen frmVatReturn, permission VAT_RETURN, table VatReturns):
'   the boxes of the return form of the Zakat, Tax and Customs Authority for a
'   tax period (a month or a quarter), from the documents (qryVatReturnTotals):
'     1  standard-rated sales (lines of category S), returns as adjustments
'     3  zero-rated sales (Z), 5 exempt sales (E), 6 total sales
'     7  standard-rated purchases (lines with VAT) and expenses with a tax invoice
'    10  zero-rated purchases, 12 total purchases
'    13  the VAT of the period, 14 corrections of earlier periods,
'    15  the credit carried from the return before, 16 net VAT due (negative = refund)
'   Boxes the shop does not record (2, 4, 8, 9, 11) stay zero.
'   SaveVatDraft     the figures of the period saved as a draft (printing needs it)
'   FileVatReturn    the figures are frozen and the settlement entry is made on
'                    the filing day (journal source VAT_RETURN): output VAT 2200 and
'                    input VAT 1500 closed into the settlement account 2250.
'   PayVatReturn     the payment entry (VAT_PAYMENT): 2250 to the bank account.
'   UnfileVatReturn  back to a draft (the last filed return, not paid, not closed).
' Same rules as tools/sim.py (Store.file_vat_return).
'==============================================================================
Option Compare Database
Option Explicit

Public Const VAT_SETTLEMENT As Long = 2250
Private Const VAT_FIELDS As String = "SalesStdAmount,SalesStdAdjust,SalesStdVAT,SalesZeroAmount,SalesZeroAdjust," & _
    "SalesExemptAmount,SalesExemptAdjust,PurchStdAmount,PurchStdAdjust,PurchStdVAT,PurchZeroAmount,PurchZeroAdjust"

'------------------------------------------------------------------------------
' Figures
'------------------------------------------------------------------------------
Public Function VatPeriodTotals(ByVal FromDate As Date, ByVal ToDate As Date) As Object
    ' The boxes of the period from the documents: a dictionary field name -> amount.
    Dim d As Object, rs As DAO.Recordset, f As Variant
    Set d = CreateObject("Scripting.Dictionary")
    SetPeriod FromDate, ToDate
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM qryVatReturnTotals", dbOpenSnapshot)
    For Each f In Split(VAT_FIELDS, ",")
        If rs.EOF Then d(CStr(f)) = CCur(0) Else d(CStr(f)) = CCur(Nz(rs.Fields(CStr(f)).Value, 0))
    Next
    rs.Close
    d("Corrections") = CCur(0)
    d("CarriedCredit") = CCur(0)
    d("NetDue") = CCur(0)
    Set VatPeriodTotals = d
End Function

Public Function VatReturnFigures(ByVal VatReturnID As Long) As Object
    ' The boxes saved in a return.
    Dim d As Object, rs As DAO.Recordset, f As Variant
    Set d = CreateObject("Scripting.Dictionary")
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM VatReturns WHERE VatReturnID = " & VatReturnID, dbOpenSnapshot)
    If Not rs.EOF Then
        For Each f In Split(VAT_FIELDS & ",Corrections,CarriedCredit,NetDue", ",")
            d(CStr(f)) = CCur(Nz(rs.Fields(CStr(f)).Value, 0))
        Next
    End If
    rs.Close
    Set VatReturnFigures = d
End Function

Public Function VatNetDue(ByVal d As Object) As Currency
    ' Box 16 = box 13 (output VAT - input VAT) + corrections - carried credit.
    VatNetDue = d("SalesStdVAT") - d("PurchStdVAT") + d("Corrections") - d("CarriedCredit")
End Function

Public Function SuggestedCarriedCredit(ByVal FromDate As Date, Optional ByVal ExceptID As Long = 0) As Currency
    ' The refund of the filed return just before the period, carried to this one.
    Dim net As Variant
    net = DbValue("SELECT TOP 1 NetDue FROM VatReturns WHERE Status = 'FILED' AND PeriodTo < " & SqlDate(FromDate) & _
                  " AND VatReturnID <> " & ExceptID & " ORDER BY PeriodTo DESC")
    If Nz(net, 0) < 0 Then SuggestedCarriedCredit = -net
End Function

Public Function VatOverlap(ByVal FromDate As Date, ByVal ToDate As Date, ByVal ExceptID As Long) As String
    ' The number of a return whose period overlaps, "" when none.
    VatOverlap = Nz(DbValue("SELECT TOP 1 ReturnNumber FROM VatReturns WHERE PeriodFrom <= " & SqlDate(ToDate) & _
                            " AND PeriodTo >= " & SqlDate(FromDate) & " AND VatReturnID <> " & ExceptID & _
                            " ORDER BY PeriodFrom"), "")
End Function

Public Function VatDrift(ByVal VatReturnID As Long) As Currency
    ' A filed return: how much the VAT of its documents changed after filing (for box 14 of the next one).
    Dim saved As Object, live As Object, p As Variant
    p = DbValue("SELECT PeriodFrom FROM VatReturns WHERE VatReturnID = " & VatReturnID)
    If IsNull(p) Then Exit Function
    Set saved = VatReturnFigures(VatReturnID)
    Set live = VatPeriodTotals(p, DbValue("SELECT PeriodTo FROM VatReturns WHERE VatReturnID = " & VatReturnID))
    VatDrift = (live("SalesStdVAT") - live("PurchStdVAT")) - (saved("SalesStdVAT") - saved("PurchStdVAT"))
End Function

'------------------------------------------------------------------------------
' Draft, filing, payment
'------------------------------------------------------------------------------
Private Function CanVat(ByVal Action As String) As String
    If Not HasPermission("VAT_RETURN") Then
        CanVat = "·«  „·ﬂ ’·«ÕÌ… «·≈ﬁ—«— «·÷—Ì»Ì."
    ElseIf Not CanScreenAction("frmVatReturn", Action, True) Then
        CanVat = "·«  „·ﬂ ’·«ÕÌ… " & IIf(Action = "ADD", "«·Õ›Ÿ Ê«·«⁄ „«œ", IIf(Action = "EDIT", " ”ÃÌ· «·”œ«œ", "«·Õ–› Ê≈·€«¡ «·«⁄ „«œ")) & _
                 " ›Ì ‘«‘… «·≈ﬁ—«— «·÷—Ì»Ì."
    End If
End Function

Private Function ReturnStatus(ByVal VatReturnID As Long) As String
    ReturnStatus = Nz(DbValue("SELECT Status FROM VatReturns WHERE VatReturnID = " & VatReturnID), "")
End Function

Public Function SaveVatDraft(ByVal FromDate As Date, ByVal ToDate As Date, ByVal Corrections As Currency, _
                             ByVal CarriedCredit As Variant, ByRef VatReturnID As Long) As String
    ' A new draft (VatReturnID = 0) or the draft VatReturnID with the figures of the documents now.
    ' CarriedCredit Null: the refund of the return before.
    SaveVatDraft = CanVat(IIf(VatReturnID = 0, "ADD", "EDIT"))
    If Len(SaveVatDraft) = 0 Then SaveVatDraft = WriteVatDraft(FromDate, ToDate, Corrections, CarriedCredit, VatReturnID)
End Function

Private Function WriteVatDraft(ByVal FromDate As Date, ByVal ToDate As Date, ByVal Corrections As Currency, _
                               ByVal CarriedCredit As Variant, ByRef VatReturnID As Long) As String
    Dim d As Object, rs As DAO.Recordset, f As Variant, other As String
    FromDate = DateValue(FromDate)
    ToDate = DateValue(ToDate)
    If ToDate < FromDate Then
        WriteVatDraft = "‰Â«Ì… «·› —… ﬁ»· »œ«Ì Â«."
        Exit Function
    End If
    If VatReturnID > 0 And ReturnStatus(VatReturnID) <> "DRAFT" Then
        WriteVatDraft = "«·≈ﬁ—«— „⁄ „œ: ·«   €Ì— ﬁÌ„Â. √·€ˆ «·«⁄ „«œ √Ê·« ≈‰ ﬂ«‰ «· €ÌÌ— „ﬁ’Êœ«."
        Exit Function
    End If
    other = VatOverlap(FromDate, ToDate, VatReturnID)
    If Len(other) > 0 Then
        WriteVatDraft = "«·› —…   œ«Œ· „⁄ «·≈ﬁ—«— " & other & "." & vbCrLf & "·ﬂ· › —… ÷—Ì»Ì… ≈ﬁ—«— Ê«Õœ."
        Exit Function
    End If
    If IsNull(CarriedCredit) Then CarriedCredit = SuggestedCarriedCredit(FromDate, VatReturnID)
    If CarriedCredit < 0 Then
        WriteVatDraft = "«·—’Ìœ «·„—ÕÛ¯· («·Œ«‰… 15) ·« ÌﬂÊ‰ ”«·»«."
        Exit Function
    End If
    Set d = VatPeriodTotals(FromDate, ToDate)
    d("Corrections") = Corrections
    d("CarriedCredit") = CCur(CarriedCredit)
    d("NetDue") = VatNetDue(d)

    If VatReturnID = 0 Then
        Set rs = CurrentDb.OpenRecordset("VatReturns", dbOpenDynaset)
        rs.AddNew
        rs!Status = "DRAFT"
    Else
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM VatReturns WHERE VatReturnID = " & VatReturnID, dbOpenDynaset)
        rs.Edit
    End If
    rs!ReturnNumber = "VAT-" & Format$(ToDate, "yyyymmdd")
    rs!PeriodFrom = FromDate
    rs!PeriodTo = ToDate
    For Each f In d.Keys
        rs.Fields(CStr(f)).Value = d(f)
    Next
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Bookmark = rs.LastModified
    VatReturnID = rs!VatReturnID
    rs.Close
    LogAction "VAT_DRAFT", "VatReturns", "VAT-" & Format$(ToDate, "yyyymmdd")
End Function

Public Function DeleteVatDraft(ByVal VatReturnID As Long) As String
    DeleteVatDraft = CanVat("DELETE")
    If Len(DeleteVatDraft) > 0 Then Exit Function
    If ReturnStatus(VatReturnID) <> "DRAFT" Then
        DeleteVatDraft = " ıÕ–› «·„”Êœ… ›ﬁÿ. ·≈ﬁ—«— „⁄ „œ «” Œœ„ ´≈·€«¡ «·«⁄ „«œª √Ê·«."
        Exit Function
    End If
    CurrentDb.Execute "DELETE FROM VatReturns WHERE VatReturnID = " & VatReturnID, dbFailOnError
    LogAction "VAT_DELETE", "VatReturns", CStr(VatReturnID)
End Function

Private Function EntryOf(ByVal SourceType As String, ByVal VatReturnID As Long) As Variant
    EntryOf = DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType = " & SqlText(SourceType) & _
                      " AND SourceID = " & VatReturnID)
End Function

Public Function FileVatReturn(ByVal VatReturnID As Long, ByVal FiledDate As Date, ByVal FilingRef As String) As String
    ' The draft with the figures of the documents now, frozen; the settlement entry on FiledDate.
    Dim periodFrom As Date, periodTo As Date, msg As String, d As Object, later As String
    FileVatReturn = CanVat("ADD")
    If Len(FileVatReturn) > 0 Then Exit Function
    If ReturnStatus(VatReturnID) <> "DRAFT" Then
        FileVatReturn = "«·≈ﬁ—«— €Ì— „ÊÃÊœ √Ê „⁄ „œ »«·›⁄·."
        Exit Function
    End If
    periodFrom = DbValue("SELECT PeriodFrom FROM VatReturns WHERE VatReturnID = " & VatReturnID)
    periodTo = DbValue("SELECT PeriodTo FROM VatReturns WHERE VatReturnID = " & VatReturnID)
    FiledDate = DateValue(FiledDate)
    If FiledDate <= periodTo Then
        FileVatReturn = "Ìı⁄ „œ «·≈ﬁ—«— »⁄œ «‰ Â«¡ › — Â: «Œ —  «—ÌŒ« »⁄œ " & GDate(periodTo) & "."
    ElseIf FiledDate > Date Then
        FileVatReturn = " «—ÌŒ «·«⁄ „«œ »⁄œ «·ÌÊ„."
    Else
        FileVatReturn = ClosedPeriodProblem(FiledDate)
    End If
    If Len(FileVatReturn) > 0 Then Exit Function
    later = Nz(DbValue("SELECT TOP 1 ReturnNumber FROM VatReturns WHERE Status = 'FILED' AND PeriodFrom > " & _
                       SqlDate(periodTo) & " ORDER BY PeriodFrom"), "")
    If Len(later) > 0 Then
        FileVatReturn = "ÌÊÃœ ≈ﬁ—«— „⁄ „œ »⁄œ Â–Â «·› —… (" & later & "): «·≈ﬁ—«—«   ı⁄ „œ »«· — Ì»."
        Exit Function
    End If
    ' the figures of the documents at the moment of filing
    Set d = VatReturnFigures(VatReturnID)
    msg = WriteVatDraft(periodFrom, periodTo, d("Corrections"), d("CarriedCredit"), VatReturnID)
    If Len(msg) > 0 Then
        FileVatReturn = msg
        Exit Function
    End If
    CurrentDb.Execute "UPDATE VatReturns SET Status = 'FILED', FiledDate = " & SqlDate(FiledDate) & ", FilingRef = " & _
                      IIf(Len(Trim$(FilingRef)) > 0, SqlText(Left$(Trim$(FilingRef), 30)), "Null") & _
                      ", EmployeeID = " & CurrentUserID() & " WHERE VatReturnID = " & VatReturnID, dbFailOnError
    msg = SyncJournal()
    Set d = VatReturnFigures(VatReturnID)
    If IsNull(EntryOf("VAT_RETURN", VatReturnID)) And d("SalesStdVAT") - d("PurchStdVAT") + d("Corrections") <> 0 Then
        CurrentDb.Execute "UPDATE VatReturns SET Status = 'DRAFT', FiledDate = Null WHERE VatReturnID = " & VatReturnID, _
                          dbFailOnError
        FileVatReturn = " ⁄–— ≈‰‘«¡ ﬁÌœ «· ”ÊÌ…:" & vbCrLf & msg
        Exit Function
    End If
    LogAction "VAT_FILE", "VatReturns", "VAT-" & Format$(periodTo, "yyyymmdd"), FilingRef
End Function

Public Function UnfileVatReturn(ByVal VatReturnID As Long, ByVal Reason As String) As String
    Dim periodTo As Variant, later As String
    UnfileVatReturn = CanVat("DELETE")
    If Len(UnfileVatReturn) > 0 Then Exit Function
    If ReturnStatus(VatReturnID) <> "FILED" Then
        UnfileVatReturn = "«·≈ﬁ—«— €Ì— „⁄ „œ."
        Exit Function
    End If
    If Len(Trim$(Reason)) = 0 Then
        UnfileVatReturn = "«ﬂ » ”»» ≈·€«¡ «·«⁄ „«œ ›Ì «·„·«ÕŸ« ."
        Exit Function
    End If
    If Nz(DbValue("SELECT PaidAmount FROM VatReturns WHERE VatReturnID = " & VatReturnID), 0) <> 0 Then
        UnfileVatReturn = "«·≈ﬁ—«— „”œœ: √·€ˆ «·”œ«œ √Ê·«."
        Exit Function
    End If
    periodTo = DbValue("SELECT PeriodTo FROM VatReturns WHERE VatReturnID = " & VatReturnID)
    later = Nz(DbValue("SELECT TOP 1 ReturnNumber FROM VatReturns WHERE Status = 'FILED' AND PeriodFrom > " & _
                       SqlDate(periodTo) & " ORDER BY PeriodFrom"), "")
    If Len(later) > 0 Then
        UnfileVatReturn = "√·€ˆ «⁄ „«œ «·≈ﬁ—«— «··«Õﬁ (" & later & ") √Ê·«: —’ÌœÂ «·„—ÕÛ¯· ﬁœ Ì⁄ „œ ⁄·Ï Â–« «·≈ﬁ—«—."
        Exit Function
    End If
    UnfileVatReturn = ClosedPeriodProblem(DbValue("SELECT FiledDate FROM VatReturns WHERE VatReturnID = " & VatReturnID))
    If Len(UnfileVatReturn) > 0 Then Exit Function
    CurrentDb.Execute "UPDATE VatReturns SET Status = 'DRAFT', FiledDate = Null, Notes = " & _
                      SqlText(Left$(Trim$(Reason), 255)) & " WHERE VatReturnID = " & VatReturnID, dbFailOnError
    SyncJournal                                        ' removes the settlement entry
    LogAction "VAT_UNFILE", "VatReturns", CStr(VatReturnID), Reason
End Function

Public Function VatPaymentAccountProblem(ByVal AccountCode As Long) As String
    ' The bank (or another cash account of the current assets that is not a cash box: the balance
    ' of a box comes from its vouchers, so a box pays the VAT through the bank only).
    Dim ok As Variant
    ok = DbValue("SELECT AccountCode FROM Accounts WHERE AccountCode = " & AccountCode & " AND IsPosting = True AND " & _
                 "AccountType = 'ASSET' AND Level2Code = 11 AND Nz(Level3Code, 0) <> 1100 AND AccountCode NOT IN " & _
                 "(1300, 1400, 1500, 1600)")
    If IsNull(ok) Then VatPaymentAccountProblem = "«Œ — Õ”«» «·»‰ﬂ «·–Ì ”ıœˆ¯œ  „‰Â «·÷—Ì»…."
End Function

Public Function PayVatReturn(ByVal VatReturnID As Long, ByVal PaidDate As Date, ByVal Amount As Currency, _
                             ByVal AccountCode As Long) As String
    Dim filed As Variant, due As Currency
    PayVatReturn = CanVat("EDIT")
    If Len(PayVatReturn) > 0 Then Exit Function
    If ReturnStatus(VatReturnID) <> "FILED" Then
        PayVatReturn = "«⁄ „œ «·≈ﬁ—«— √Ê·«° À„ ”Ã¯· «·”œ«œ."
        Exit Function
    End If
    due = Nz(DbValue("SELECT NetDue FROM VatReturns WHERE VatReturnID = " & VatReturnID), 0)
    filed = DbValue("SELECT FiledDate FROM VatReturns WHERE VatReturnID = " & VatReturnID)
    PaidDate = DateValue(PaidDate)
    If due <= 0 Then
        PayVatReturn = "·«  ÊÃœ ÷—Ì»… „” Õﬁ… ··”œ«œ ›Ì Â–« «·≈ﬁ—«— («·—’Ìœ «·œ«∆‰ Ìı—ÕÛ¯· ··≈ﬁ—«— «· «·Ì)."
    ElseIf Amount <= 0 Or Amount > due Then
        PayVatReturn = "«·„»·€ «·„”œœ »Ì‰ 0.01 Ê " & Format$(due, "#,##0.00") & "."
    ElseIf PaidDate < DateValue(filed) Then
        PayVatReturn = " «—ÌŒ «·”œ«œ ﬁ»·  «—ÌŒ «⁄ „«œ «·≈ﬁ—«— (" & GDate(filed) & ")."
    ElseIf PaidDate > Date Then
        PayVatReturn = " «—ÌŒ «·”œ«œ »⁄œ «·ÌÊ„."
    Else
        PayVatReturn = VatPaymentAccountProblem(AccountCode)
    End If
    If Len(PayVatReturn) = 0 Then PayVatReturn = ClosedPeriodProblem(PaidDate)
    If Len(PayVatReturn) = 0 Then
        PayVatReturn = ClosedPeriodProblem(DbValue("SELECT PaidDate FROM VatReturns WHERE VatReturnID = " & VatReturnID))
    End If
    If Len(PayVatReturn) > 0 Then Exit Function
    CurrentDb.Execute "UPDATE VatReturns SET PaidDate = " & SqlDate(PaidDate) & ", PaidAmount = " & _
                      Str$(Amount) & ", PaidAccount = " & AccountCode & " WHERE VatReturnID = " & VatReturnID, dbFailOnError
    SyncJournal
    LogAction "VAT_PAY", "VatReturns", CStr(VatReturnID), Format$(Amount, "0.00")
End Function

Public Function CancelVatPayment(ByVal VatReturnID As Long) As String
    Dim paid As Variant
    CancelVatPayment = CanVat("EDIT")
    If Len(CancelVatPayment) > 0 Then Exit Function
    paid = DbValue("SELECT PaidDate FROM VatReturns WHERE VatReturnID = " & VatReturnID & " AND PaidAmount <> 0")
    If IsNull(paid) Then
        CancelVatPayment = "·« ÌÊÃœ ”œ«œ „”Ã· ·Â–« «·≈ﬁ—«—."
        Exit Function
    End If
    CancelVatPayment = ClosedPeriodProblem(paid)
    If Len(CancelVatPayment) > 0 Then Exit Function
    CurrentDb.Execute "UPDATE VatReturns SET PaidDate = Null, PaidAmount = 0, PaidAccount = Null WHERE VatReturnID = " & _
                      VatReturnID, dbFailOnError
    SyncJournal
    LogAction "VAT_UNPAY", "VatReturns", CStr(VatReturnID)
End Function

'------------------------------------------------------------------------------
' Screen frmVatReturn (OpenArgs: the return to show)
'------------------------------------------------------------------------------
Public Function VatBoxCaption(ByVal Box As Long) As String
    ' Same captions as VAT_BOXES in tools/queries.py (the report).
    Select Case Box
        Case 1: VatBoxCaption = "«·„»Ì⁄«  «·Œ«÷⁄… ··‰”»… «·√”«”Ì… (15%)"
        Case 2: VatBoxCaption = "«·„»Ì⁄«  ··„Ê«ÿ‰Ì‰ («·Œœ„«  «·’ÕÌ… «·Œ«’… Ê«· ⁄·Ì„ «·√Â·Ì Ê«·„”ﬂ‰ «·√Ê·)"
        Case 3: VatBoxCaption = "«·„»Ì⁄«  «·„Õ·Ì… «·Œ«÷⁄… ··‰”»… «·’›—Ì…"
        Case 4: VatBoxCaption = "«·’«œ—« "
        Case 5: VatBoxCaption = "«·„»Ì⁄«  «·„⁄›«…"
        Case 6: VatBoxCaption = "≈Ã„«·Ì «·„»Ì⁄« "
        Case 7: VatBoxCaption = "«·„‘ —Ì«  «·Œ«÷⁄… ··‰”»… «·√”«”Ì… („⁄ «·„’—Ê›«  »›« Ê—… ÷—Ì»Ì…)"
        Case 8: VatBoxCaption = "«·«” Ì—«œ«  «·Œ«÷⁄… ··‰”»… «·√”«”Ì… Ê«·„œ›Ê⁄… ÷—Ì» Â« ›Ì «·Ã„«—ﬂ"
        Case 9: VatBoxCaption = "«·«” Ì—«œ«  «·Œ«÷⁄… ··÷—Ì»… »¬·Ì… «·«Õ ”«» «·⁄ﬂ”Ì"
        Case 10: VatBoxCaption = "«·„‘ —Ì«  «·Œ«÷⁄… ··‰”»… «·’›—Ì…"
        Case 11: VatBoxCaption = "«·„‘ —Ì«  «·„⁄›«…"
        Case 12: VatBoxCaption = "≈Ã„«·Ì «·„‘ —Ì« "
        Case 13: VatBoxCaption = "≈Ã„«·Ì ÷—Ì»… «·ﬁÌ„… «·„÷«›… «·„” Õﬁ… ⁄‰ «·› —… «·Õ«·Ì…"
        Case 14: VatBoxCaption = " ’ÕÌÕ«  „‰ «·› —«  «·”«»ﬁ…"
        Case 15: VatBoxCaption = "÷—Ì»… «·ﬁÌ„… «·„÷«›… «·„—ÕÛ¯·… „‰ «·› —«  «·”«»ﬁ… (—’Ìœ œ«∆‰)"
        Case 16: VatBoxCaption = "’«›Ì «·÷—Ì»… «·„” Õﬁ… (”«·» = „” —œ…)"
    End Select
End Function

Private Function ListItem(ByVal Value As Variant) As String
    If IsNull(Value) Then
        ListItem = """"""
    ElseIf VarType(Value) = vbCurrency Then
        ListItem = """" & Format$(Value, "#,##0.00") & """"
    Else
        ListItem = """" & Replace(CStr(Value), """", "'") & """"
    End If
End Function

Public Function VatBoxRows(ByVal d As Object) As String
    ' The value list of the 16 boxes (with the column heads) for the list of the screen.
    Dim rows As String, box As Long, amount As Variant, adjust As Variant, vat As Variant
    rows = """«·»‰œ"";""«·Ê’›"";""«·„»·€"";""«· ⁄œÌ·« "";""«·÷—Ì»…"""
    For box = 1 To 16
        amount = CCur(0): adjust = CCur(0): vat = CCur(0)
        Select Case box
            Case 1: amount = d("SalesStdAmount"): adjust = d("SalesStdAdjust"): vat = d("SalesStdVAT")
            Case 3: amount = d("SalesZeroAmount"): adjust = d("SalesZeroAdjust")
            Case 5: amount = d("SalesExemptAmount"): adjust = d("SalesExemptAdjust")
            Case 6
                amount = d("SalesStdAmount") + d("SalesZeroAmount") + d("SalesExemptAmount")
                adjust = d("SalesStdAdjust") + d("SalesZeroAdjust") + d("SalesExemptAdjust")
                vat = d("SalesStdVAT")
            Case 7: amount = d("PurchStdAmount"): adjust = d("PurchStdAdjust"): vat = d("PurchStdVAT")
            Case 10: amount = d("PurchZeroAmount"): adjust = d("PurchZeroAdjust")
            Case 12
                amount = d("PurchStdAmount") + d("PurchZeroAmount")
                adjust = d("PurchStdAdjust") + d("PurchZeroAdjust")
                vat = d("PurchStdVAT")
            Case 13: amount = Null: adjust = Null: vat = d("SalesStdVAT") - d("PurchStdVAT")
            Case 14: amount = Null: adjust = Null: vat = d("Corrections")
            Case 15: amount = Null: adjust = Null: vat = d("CarriedCredit")
            Case 16: amount = Null: adjust = Null: vat = VatNetDue(d)
        End Select
        rows = rows & ";" & ListItem(CStr(box)) & ";" & ListItem(VatBoxCaption(box)) & ";" & ListItem(amount) & ";" & _
               ListItem(adjust) & ";" & ListItem(vat)
    Next
    VatBoxRows = rows
End Function

Private Function ShownID(ByVal frm As Access.Form) As Long
    ShownID = Nz(frm!txtReturnID.Value, 0)
End Function

Public Sub VatReturnLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    frm!cboPayAccount.RowSource = "SELECT AccountCode, AccountName FROM Accounts WHERE IsPosting = True AND " & _
        "AccountType = 'ASSET' AND Level2Code = 11 AND Nz(Level3Code, 0) <> 1100 AND AccountCode NOT IN " & _
        "(1300, 1400, 1500, 1600) ORDER BY TreeKey"
    If Nz(frm.OpenArgs, 0) > 0 Then
        VatShowReturn frm, CLng(frm.OpenArgs)
    Else
        VatQuickPeriod frm, "LASTMONTH"
    End If
End Sub

Public Sub VatQuickPeriod(ByVal frm As Access.Form, ByVal Which As String)
    Dim y As Integer, q As Integer
    If Which = "LASTQUARTER" Then
        y = Year(Date)
        q = (Month(Date) - 1) \ 3                      ' this quarter starts at month 3q + 1
        frm!txtFrom.Value = DateSerial(y, 3 * q - 2, 1)
        frm!txtTo.Value = DateSerial(y, 3 * q + 1, 0)
    Else
        SetQuickPeriod frm, Which
    End If
    VatCalculate frm
End Sub

Public Sub VatCalculate(ByVal frm As Access.Form)
    ' The return of the period: the saved one if there is one for exactly this period, else the figures now.
    Dim fromDate As Date, toDate As Date, id As Variant, other As String, d As Object
    If Not IsDate(frm!txtFrom.Value) Or Not IsDate(frm!txtTo.Value) Then
        ShowWarning "√œŒ· »œ«Ì… «·› —… «·÷—Ì»Ì… Ê‰Â«Ì Â«."
        Exit Sub
    End If
    fromDate = DateValue(frm!txtFrom.Value)
    toDate = DateValue(frm!txtTo.Value)
    If toDate < fromDate Then
        ShowWarning "‰Â«Ì… «·› —… ﬁ»· »œ«Ì Â«."
        Exit Sub
    End If
    id = DbValue("SELECT VatReturnID FROM VatReturns WHERE PeriodFrom = " & SqlDate(fromDate) & " AND PeriodTo = " & _
                 SqlDate(toDate))
    If Not IsNull(id) Then
        If ReturnStatus(CLng(id)) = "DRAFT" Then          ' a draft follows the documents
            SaveVatDraft fromDate, toDate, Nz(DbValue("SELECT Corrections FROM VatReturns WHERE VatReturnID = " & id), 0), _
                         Nz(DbValue("SELECT CarriedCredit FROM VatReturns WHERE VatReturnID = " & id), 0), CLng(id)
        End If
        VatShowReturn frm, CLng(id)
        Exit Sub
    End If
    Set d = VatPeriodTotals(fromDate, toDate)
    d("CarriedCredit") = SuggestedCarriedCredit(fromDate)
    frm!txtReturnID.Value = Null
    frm!txtCorrections.Value = 0
    frm!txtCarried.Value = d("CarriedCredit")
    frm!lstBoxes.RowSource = VatBoxRows(d)
    other = VatOverlap(fromDate, toDate, 0)
    If Len(other) > 0 Then
        VatState frm, "«·› —…   œ«Œ· „⁄ «·≈ﬁ—«— " & other & ": «Œ — › —… ·« ≈ﬁ—«— ·Â«.", CLR_DANGER
    Else
        VatState frm, "≈ﬁ—«— €Ì— „Õ›ÊŸ: «·√—ﬁ«„ „‰ «·„” ‰œ«  «·¬‰. «Õ›ŸÂ „”Êœ… √Ê «⁄ „œÂ.", CLR_PRIMARY
    End If
    VatShowNet frm
    VatButtons frm, ""
End Sub

Public Sub VatShowReturn(ByVal frm As Access.Form, ByVal VatReturnID As Long)
    Dim rs As DAO.Recordset, status As String, drift As Currency, info As String
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM VatReturns WHERE VatReturnID = " & VatReturnID, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        ShowWarning "«·≈ﬁ—«— €Ì— „ÊÃÊœ."
        Exit Sub
    End If
    status = rs!Status
    frm!txtReturnID.Value = VatReturnID
    frm!txtFrom.Value = rs!PeriodFrom
    frm!txtTo.Value = rs!PeriodTo
    frm!txtCorrections.Value = rs!Corrections
    frm!txtCarried.Value = rs!CarriedCredit
    frm!txtFiledDate.Value = IIf(IsNull(rs!FiledDate), Date, rs!FiledDate)
    frm!txtFilingRef.Value = rs!FilingRef
    frm!txtPaidDate.Value = IIf(IsNull(rs!PaidDate), Date, rs!PaidDate)
    frm!txtPaidAmount.Value = IIf(Nz(rs!PaidAmount, 0) <> 0, rs!PaidAmount, IIf(rs!NetDue > 0, rs!NetDue, Null))
    frm!cboPayAccount.Value = Nz(rs!PaidAccount, 1200)
    If status = "FILED" Then
        info = "«·≈ﬁ—«— " & rs!ReturnNumber & " „⁄ „œ ›Ì " & GDate(rs!FiledDate)
        If Nz(rs!PaidAmount, 0) <> 0 Then
            info = info & "° „”œœ " & Format$(rs!PaidAmount, "#,##0.00") & " ›Ì " & GDate(rs!PaidDate)
        ElseIf rs!NetDue > 0 Then
            info = info & "° €Ì— „”œœ"
        End If
    Else
        info = "„”Êœ… «·≈ﬁ—«— " & rs!ReturnNumber & " («·√—ﬁ«„ „‰ «·„” ‰œ«  «·¬‰)"
    End If
    rs.Close
    frm!lstBoxes.RowSource = VatBoxRows(VatReturnFigures(VatReturnID))
    If status = "FILED" Then
        drift = VatDrift(VatReturnID)
        If drift <> 0 Then info = info & vbCrLf & " €Ì—  ÷—Ì»… „” ‰œ«  «·› —… »⁄œ «·«⁄ „«œ »„ﬁœ«— " & _
                                  Format$(drift, "#,##0.00") & ": √÷›Â ›Ì  ’ÕÌÕ«  «·≈ﬁ—«— «· «·Ì («·Œ«‰… 14)."
    End If
    VatState frm, info, IIf(status = "FILED", CLR_SUCCESS, CLR_PRIMARY)
    VatShowNet frm
    VatButtons frm, status
    frm!lstReturns.Requery
End Sub

Private Sub VatState(ByVal frm As Access.Form, ByVal Text As String, ByVal Color As Long)
    frm!lblState.Caption = Text
    frm!lblState.ForeColor = Color
End Sub

Public Sub VatShowNet(ByVal frm As Access.Form)
    ' Box 16 as the corrections and the carried credit are typed.
    Dim d As Object, net As Currency
    If ShownID(frm) = 0 And (Not IsDate(frm!txtFrom.Value) Or Not IsDate(frm!txtTo.Value)) Then Exit Sub
    If ShownID(frm) > 0 Then
        Set d = VatReturnFigures(ShownID(frm))
    Else
        Set d = VatPeriodTotals(DateValue(frm!txtFrom.Value), DateValue(frm!txtTo.Value))
    End If
    If ReturnStatus(ShownID(frm)) <> "FILED" Then
        d("Corrections") = CCur(Nz(frm!txtCorrections.Value, 0))
        d("CarriedCredit") = CCur(Nz(frm!txtCarried.Value, 0))
        frm!lstBoxes.RowSource = VatBoxRows(d)
    End If
    net = VatNetDue(d)
    If net >= 0 Then
        frm!lblNetDue.Caption = "’«›Ì «·÷—Ì»… «·„” Õﬁ…: " & Format$(net, "#,##0.00")
        frm!lblNetDue.ForeColor = CLR_DANGER
    Else
        frm!lblNetDue.Caption = "÷—Ì»… „” —œ… ( ı—ÕÛ¯· ··≈ﬁ—«— «· «·Ì): " & Format$(-net, "#,##0.00")
        frm!lblNetDue.ForeColor = CLR_SUCCESS
    End If
End Sub

Private Sub VatButtons(ByVal frm As Access.Form, ByVal Status As String)
    ' Status: "" (not saved), DRAFT or FILED.
    Dim filed As Boolean
    filed = (Status = "FILED")
    frm!txtCorrections.Locked = filed
    frm!txtCarried.Locked = filed
    frm!btnSaveDraft.Enabled = Not filed
    frm!btnDeleteDraft.Enabled = (Status = "DRAFT")
    frm!btnFile.Enabled = Not filed
    frm!btnUnfile.Enabled = filed
    frm!btnPay.Enabled = filed
    frm!btnUnpay.Enabled = filed
End Sub

Private Function EnsureSaved(ByVal frm As Access.Form) As Long
    ' The return shown, saved as a draft first when it is not saved yet (0 = could not).
    Dim id As Long, msg As String
    id = ShownID(frm)
    If ReturnStatus(id) <> "FILED" Then
        If Not IsDate(frm!txtFrom.Value) Or Not IsDate(frm!txtTo.Value) Then
            ShowWarning "√œŒ· »œ«Ì… «·› —… «·÷—Ì»Ì… Ê‰Â«Ì Â«."
            Exit Function
        End If
        msg = SaveVatDraft(frm!txtFrom.Value, frm!txtTo.Value, CCur(Nz(frm!txtCorrections.Value, 0)), _
                           CCur(Nz(frm!txtCarried.Value, 0)), id)
        If Len(msg) > 0 Then
            ShowWarning msg
            Exit Function
        End If
    End If
    EnsureSaved = id
End Function

Public Sub VatSaveDraft(ByVal frm As Access.Form)
    Dim id As Long
    id = EnsureSaved(frm)
    If id = 0 Then Exit Sub
    VatShowReturn frm, id
    ShowInfo " „ Õ›Ÿ „”Êœ… «·≈ﬁ—«—."
End Sub

Public Sub VatDeleteDraft(ByVal frm As Access.Form)
    Dim msg As String
    If ShownID(frm) = 0 Then Exit Sub
    If Not AskYesNo("Õ–› „”Êœ… «·≈ﬁ—«—ø") Then Exit Sub
    msg = DeleteVatDraft(ShownID(frm))
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    frm!txtReturnID.Value = Null
    frm!lstReturns.Requery
    VatCalculate frm
End Sub

Public Sub VatFile(ByVal frm As Access.Form)
    Dim id As Long, msg As String, periodTo As Date, number As String
    If Not IsDate(frm!txtFiledDate.Value) Then
        ShowWarning "√œŒ·  «—ÌŒ «⁄ „«œ «·≈ﬁ—«— (ÌÊ„  ﬁœÌ„Â ··ÂÌ∆…)."
        SafeFocus frm!txtFiledDate
        Exit Sub
    End If
    If Not AskYesNo("«⁄ „«œ «·≈ﬁ—«—ø" & vbCrLf & " ıÕ›Ÿ ﬁÌ„Â ﬂ„« ÂÌ «·¬‰° ÊÌı‰‘√ ﬁÌœ «· ”ÊÌ… » «—ÌŒ " & _
                    GDate(frm!txtFiledDate.Value) & ".") Then Exit Sub
    id = EnsureSaved(frm)
    If id = 0 Then Exit Sub
    msg = FileVatReturn(id, frm!txtFiledDate.Value, Nz(frm!txtFilingRef.Value, ""))
    VatShowReturn frm, id
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    periodTo = DbValue("SELECT PeriodTo FROM VatReturns WHERE VatReturnID = " & id)
    number = Nz(DbValue("SELECT ReturnNumber FROM VatReturns WHERE VatReturnID = " & id), "")
    ShowInfo " „ «⁄ „«œ «·≈ﬁ—«— " & number & "."
    ' the documents of the period should not change any more
    If HasPermission("PERIOD_CLOSE") And ClosedThroughDate() < periodTo Then
        If AskYesNo("≈ﬁ›«· «·› —… Õ Ï " & GDate(periodTo) & " Õ Ï ·«   €Ì— „” ‰œ«  «·≈ﬁ—«—ø") Then
            msg = ClosePeriod(periodTo, "»⁄œ «⁄ „«œ «·≈ﬁ—«— «·÷—Ì»Ì " & number)
            If Len(msg) > 0 Then
                ShowWarning msg
            Else
                ShowInfo " „ ≈ﬁ›«· «·› —… Õ Ï " & GDate(periodTo) & "."
            End If
        End If
    End If
End Sub

Public Sub VatUnfile(ByVal frm As Access.Form)
    Dim msg As String
    If ShownID(frm) = 0 Then Exit Sub
    If Not AskYesNo("≈·€«¡ «⁄ „«œ «·≈ﬁ—«—ø ÌıÕ–› ﬁÌœ «· ”ÊÌ… ÊÌ⁄Êœ «·≈ﬁ—«— „”Êœ….") Then Exit Sub
    msg = UnfileVatReturn(ShownID(frm), Nz(frm!txtNotes.Value, ""))
    If Len(msg) > 0 Then ShowWarning msg
    VatShowReturn frm, ShownID(frm)
End Sub

Public Sub VatPay(ByVal frm As Access.Form)
    Dim msg As String
    If ShownID(frm) = 0 Then Exit Sub
    If Not IsDate(frm!txtPaidDate.Value) Or Nz(frm!txtPaidAmount.Value, 0) <= 0 Or IsNull(frm!cboPayAccount.Value) Then
        ShowWarning "√œŒ·  «—ÌŒ «·”œ«œ Ê«·„»·€ ÊÕ”«» «·»‰ﬂ."
        Exit Sub
    End If
    msg = PayVatReturn(ShownID(frm), frm!txtPaidDate.Value, CCur(frm!txtPaidAmount.Value), CLng(frm!cboPayAccount.Value))
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        ShowInfo " „  ”ÃÌ· ”œ«œ «·÷—Ì»…."
    End If
    VatShowReturn frm, ShownID(frm)
End Sub

Public Sub VatUnpay(ByVal frm As Access.Form)
    Dim msg As String
    If ShownID(frm) = 0 Then Exit Sub
    If Not AskYesNo("≈·€«¡ ”œ«œ «·≈ﬁ—«—ø ÌıÕ–› ﬁÌœ «·”œ«œ.") Then Exit Sub
    msg = CancelVatPayment(ShownID(frm))
    If Len(msg) > 0 Then ShowWarning msg
    VatShowReturn frm, ShownID(frm)
End Sub

Public Sub VatPickReturn(ByVal frm As Access.Form)
    If IsNull(frm!lstReturns.Value) Then Exit Sub
    VatShowReturn frm, CLng(frm!lstReturns.Value)
End Sub

Public Sub VatOpenEntry(ByVal frm As Access.Form)
    Dim entry As Variant
    If ShownID(frm) = 0 Then Exit Sub
    entry = EntryOf("VAT_RETURN", ShownID(frm))
    If IsNull(entry) Then
        ShowWarning "·« ÌÊÃœ ﬁÌœ  ”ÊÌ…: «·≈ﬁ—«— €Ì— „⁄ „œ."
    Else
        OpenScreen "frmJournalEntry", 0, entry
    End If
End Sub

Public Sub PrintVatReturn(ByVal frm As Access.Form)
    Dim id As Long
    id = EnsureSaved(frm)
    If id = 0 Then Exit Sub
    VatShowReturn frm, id
    SetQueryParam "VatReturnID", id
    LogAction "REPORT", "VAT_RETURN", CStr(id)
    OpenReportOrQuery "rptVatReturn", "VatReturnQuery", "", _
                      PeriodText(DbValue("SELECT PeriodFrom FROM VatReturns WHERE VatReturnID = " & id), _
                                 DbValue("SELECT PeriodTo FROM VatReturns WHERE VatReturnID = " & id))
End Sub
