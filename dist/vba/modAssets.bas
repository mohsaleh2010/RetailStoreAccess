Attribute VB_Name = "modAssets"
'==============================================================================
' modAssets  -  Retail Store Management System
'
' Fixed assets (FixedAssets, screen frmAssets, permission FIXED_ASSETS):
'   bought  the asset account (a sub-account of 12: furniture, devices, cars...) and
'           the input VAT against a bank, a cash box or another account (journal
'           source ASSET); an asset owned before the program comes from the
'           opening balances (3900) with its depreciation until then (1790).
'   sold / scrapped  the accumulated depreciation and the price (with its output VAT
'           2200 when the sale is taxed) against the value, gain 4500 / loss 5650; the
'           revaluation surplus left goes to retained earnings 3300 (ASSET_DISPOSAL).
' Monthly depreciation (DepreciationRuns + AssetDepreciations, screen frmDepreciation,
' journal source DEPRECIATION), from the month of DepStartDate, never below the
' salvage value, not in the month of the disposal; 5600 debit / 1790 credit, a line per
' asset on the cost centre of the asset that month. Months are recorded in order.
'   SL straight line: (cost - salvage) / months.
'   DB declining balance: the book value at the start of the month x the yearly rate / 12
'      (DepRate, 0 = twice the straight line rate), all that is left in the last month of its life.
' Revaluation (AssetRevaluations, ASSET_REVAL, docs/65): a new value and the months left;
' the depreciation goes on from the new value. Transfer between cost centres
' (AssetTransfers, ASSET_TRANSFER): the value and the accumulated depreciation move.
' Same amounts as tools/sim.py (Store.depreciate, Store.revalue, Store.transfer_asset).
'==============================================================================
Option Compare Database
Option Explicit

Public Const ACCUM_DEPRECIATION As Long = 1790

Public Function MonthEnd(ByVal AnyDay As Date) As Date
    MonthEnd = DateSerial(Year(AnyDay), Month(AnyDay) + 1, 0)
End Function

Private Function AssetField(ByVal AssetID As Long, ByVal FieldName As String) As Variant
    AssetField = DbValue("SELECT " & FieldName & " FROM FixedAssets WHERE AssetID = " & AssetID)
End Function

Private Function CanAssets(ByVal Screen As String, ByVal Action As String) As String
    If Not HasPermission("FIXED_ASSETS") Then
        CanAssets = "·«  „·ﬂ ’·«ÕÌ… «·√’Ê· «·À«» …."
    ElseIf Not CanScreenAction(Screen, Action, True) Then
        CanAssets = "·«  „·ﬂ Â–Â «·’·«ÕÌ… ›Ì ‘«‘… " & IIf(Screen = "frmAssets", "«·√’Ê· «·À«» ….", "«·≈Â·«ﬂ «·‘Â—Ì.")
    End If
End Function

Public Function AssetDepreciated(ByVal AssetID As Long) As Currency
    ' Depreciation until now: the one before the program and the monthly entries.
    AssetDepreciated = Nz(AssetField(AssetID, "OpeningAccumDep"), 0) + _
                       Nz(DbValue("SELECT Sum(Amount) FROM AssetDepreciations WHERE AssetID = " & AssetID), 0)
End Function

Public Function AssetValue(ByVal AssetID As Long) As Currency
    ' The value in the asset account: the cost and the revaluations.
    AssetValue = Nz(AssetField(AssetID, "Cost"), 0) + _
                 Nz(DbValue("SELECT Sum(ValueChange) FROM AssetRevaluations WHERE AssetID = " & AssetID), 0)
End Function

Public Function AssetBookValue(ByVal AssetID As Long) As Currency
    AssetBookValue = AssetValue(AssetID) - AssetDepreciated(AssetID)
End Function

Public Function LastAssetEvent(ByVal AssetID As Long) As Variant
    ' The date of the last revaluation or transfer of the asset (Null = none).
    Dim a As Variant, b As Variant
    a = DbValue("SELECT Max(RevalDate) FROM AssetRevaluations WHERE AssetID = " & AssetID)
    b = DbValue("SELECT Max(TransferDate) FROM AssetTransfers WHERE AssetID = " & AssetID)
    If IsNull(a) Then
        LastAssetEvent = b
    ElseIf IsNull(b) Then
        LastAssetEvent = a
    ElseIf b > a Then
        LastAssetEvent = b
    Else
        LastAssetEvent = a
    End If
End Function

Private Function DepreciatedFrom(ByVal AssetID As Long, ByVal AnyDay As Date) As Variant
    ' The last month with depreciation of the asset from the month of AnyDay (Null = none).
    DepreciatedFrom = DbValue("SELECT Max(r.RunMonth) FROM AssetDepreciations AS d INNER JOIN DepreciationRuns AS r ON " & _
                              "d.RunID = r.RunID WHERE d.AssetID = " & AssetID & " AND r.RunMonth >= " & _
                              SqlDate(DateSerial(Year(AnyDay), Month(AnyDay), 1)))
End Function

Public Function MonthlyDepreciation(ByVal Cost As Currency, ByVal Salvage As Currency, ByVal Months As Long) As Currency
    If Months > 0 Then MonthlyDepreciation = Round((Cost - Salvage) / Months, 2)
End Function

'------------------------------------------------------------------------------
' Assets
'------------------------------------------------------------------------------
Public Function AssetAccountProblem(ByVal AccountCode As Variant) As String
    If IsNull(AccountCode) Then
        AssetAccountProblem = "«Œ — Õ”«» «·√’·."
    ElseIf IsNull(DbValue("SELECT AccountCode FROM Accounts WHERE AccountCode = " & CLng(AccountCode) & " AND IsPosting <> 0 " & _
                          "AND IsActive <> 0 AND AccountType = 'ASSET' AND Level2Code = 12 AND AccountCode <> 1790")) Then
        AssetAccountProblem = "Õ”«» «·√’· Õ”«» ›—⁄Ì „‰ «·√’Ê· €Ì— «·„ œ«Ê·… (12)° €Ì— „Ã„⁄ «·≈Â·«ﬂ."
    End If
End Function

Public Function AssetProblem(ByVal AssetID As Long, ByVal AccountCode As Variant, ByVal PurchaseDate As Variant, _
                             ByVal Cost As Currency, ByVal InputVAT As Currency, ByVal Salvage As Currency, _
                             ByVal LifeMonths As Long, ByVal DepStart As Variant, ByVal SourceType As String, _
                             ByVal BankID As Variant, ByVal CashBoxID As Variant, ByVal CounterAccount As Variant, _
                             ByVal OpeningAccum As Currency, Optional ByVal DepMethod As String = "SL", _
                             Optional ByVal DepRate As Double = 0) As String
    Dim p As String
    p = AssetAccountProblem(AccountCode)
    If Len(p) > 0 Then
    ElseIf DepMethod <> "SL" And DepMethod <> "DB" Then
        p = "«Œ — ÿ—Ìﬁ… «·≈Â·«ﬂ."
    ElseIf DepRate < 0 Or DepRate > 1 Then
        p = "‰”»… «·ﬁ”ÿ «·„ ‰«ﬁ’ »Ì‰ 0% Ê100%."
    ElseIf DepMethod = "SL" And DepRate <> 0 Then
        p = "‰”»… «·ﬁ”ÿ «·„ ‰«ﬁ’ ··ÿ—Ìﬁ… «·„ ‰«ﬁ’… ›ﬁÿ."
    ElseIf Not IsDate(PurchaseDate) Or Not IsDate(DepStart) Then
        p = "«ﬂ »  «—ÌŒ «·‘—«¡ Ê»œ«Ì… «·≈Â·«ﬂ."
    ElseIf DateValue(PurchaseDate) > Date Then
        p = " «—ÌŒ «·‘—«¡ »⁄œ «·ÌÊ„."
    ElseIf DateValue(DepStart) < DateSerial(Year(PurchaseDate), Month(PurchaseDate), 1) Then
        p = "»œ«Ì… «·≈Â·«ﬂ ﬁ»· ‘Â— «·‘—«¡."
    ElseIf Cost <= 0 Then
        p = "«· ﬂ·›… √ﬂ»— „‰ ’›—."
    ElseIf InputVAT < 0 Or Salvage < 0 Or OpeningAccum < 0 Then
        p = "«·÷—Ì»… Ê«·ﬁÌ„… «·„ »ﬁÌ… Ê«·≈Â·«ﬂ «·”«»ﬁ ·«  ﬂÊ‰ ”«·»…."
    ElseIf Salvage + OpeningAccum > Cost Then
        p = "«·ﬁÌ„… «·„ »ﬁÌ… Ê«·≈Â·«ﬂ «·”«»ﬁ ·« Ì Ã«Ê“«‰ «· ﬂ·›…."
    ElseIf LifeMonths <= 0 Then
        p = "«·⁄„— «·≈‰ «ÃÌ ‘Â— ⁄·Ï «·√ﬁ·."
    Else
        Select Case SourceType
            Case "OPENING"
                If InputVAT <> 0 Then p = "«·√’· «·„ÊÃÊœ ﬁ»· «·»—‰«„Ã »·« ÷—Ì»… „œŒ·«  (”ıÃˆ¯·  Êﬁ  ‘—«∆Â)."
            Case "BANK"
                If IsNull(BankID) Then
                    p = "«Œ — «·»‰ﬂ «·–Ì œı›⁄ „‰Â."
                ElseIf Nz(DbValue("SELECT COUNT(*) FROM Banks WHERE IsActive <> 0 AND BankID = " & CLng(BankID)), 0) = 0 Then
                    p = "«·»‰ﬂ €Ì— ‰‘ÿ."
                End If
            Case "CASHBOX"
                If IsNull(CashBoxID) Then
                    p = "«Œ — «·’‰œÊﬁ «·–Ì œı›⁄ „‰Â."
                ElseIf Not BoxIsActive(CashBoxID) Then
                    p = "«·’‰œÊﬁ €Ì— ‰‘ÿ."
                ElseIf CashBoxBalance(CLng(CashBoxID), DateValue(PurchaseDate) + 1) + _
                       IIf(AssetID > 0, Nz(AssetField(AssetID, "IIf(SourceType = 'CASHBOX', Cost + InputVAT, 0)"), 0), 0) _
                       < Cost + InputVAT Then
                    p = "—’Ìœ «·’‰œÊﬁ ›Ì " & GDate(PurchaseDate) & " √ﬁ· „‰ «·„»·€."
                End If
            Case "ACCOUNT"
                p = CounterAccountProblem(CounterAccount)               ' modBank
            Case Else
                p = "«Œ — „’œ— «·‘—«¡."
        End Select
    End If
    If Len(p) = 0 And SourceType <> "OPENING" And OpeningAccum <> 0 Then
        p = "«·≈Â·«ﬂ «·”«»ﬁ ··√’Ê· «·„ÊÃÊœ… ﬁ»· «·»—‰«„Ã ›ﬁÿ („’œ— «·‘—«¡: —’Ìœ «›  «ÕÌ)."
    End If
    If Len(p) = 0 Then p = ClosedPeriodProblem(PurchaseDate)
    AssetProblem = p
End Function

Private Function AssetLocked(ByVal AssetID As Long) As String
    ' An asset with monthly depreciation, a revaluation or a transfer, or disposed keeps its data.
    If Nz(AssetField(AssetID, "Status"), "") = "DISPOSED" Then
        AssetLocked = "«·√’· „” »⁄œ: √·€ˆ «·«” »⁄«œ √Ê·«."
    ElseIf Not IsNull(LastAssetEvent(AssetID)) Then
        AssetLocked = "⁄·Ï «·√’· ≈⁄«œ…  ﬁÌÌ„ √Ê ‰ﬁ·: ·«   €Ì— »Ì«‰« Â «·„«·Ì… Ê·« ÌıÕ–›." & vbCrLf & _
                      "·· ’ÕÌÕ «Õ–›Â« „‰ ‘«‘… ≈⁄«œ… «· ﬁÌÌ„ Ê«·‰ﬁ· »œ¡« „‰ «·√ŒÌ—…."
    ElseIf Nz(DbValue("SELECT COUNT(*) FROM AssetDepreciations WHERE AssetID = " & AssetID), 0) > 0 Then
        AssetLocked = "⁄·Ï «·√’· ﬁÌÊœ ≈Â·«ﬂ: ·«   €Ì— »Ì«‰« Â «·„«·Ì… Ê·« ÌıÕ–›." & vbCrLf & _
                      "·· ’ÕÌÕ «Õ–› ﬁÌÊœ «·≈Â·«ﬂ »œ¡« „‰ «·√ŒÌ—."
    Else
        AssetLocked = ClosedPeriodProblem(AssetField(AssetID, "PurchaseDate"))
    End If
End Function

Public Function SaveAsset(ByRef AssetID As Long, ByVal AssetName As String, ByVal AccountCode As Variant, _
                          ByVal PurchaseDate As Variant, ByVal Cost As Currency, ByVal InputVAT As Currency, _
                          ByVal Salvage As Currency, ByVal LifeMonths As Long, ByVal DepStart As Variant, _
                          ByVal SourceType As String, ByVal BankID As Variant, ByVal CashBoxID As Variant, _
                          ByVal CounterAccount As Variant, ByVal OpeningAccum As Currency, ByVal Notes As String, _
                          Optional ByVal DepMethod As String = "SL", Optional ByVal DepRate As Double = 0) As String
    ' A new asset (AssetID = 0) or the changes of one without depreciation.
    Dim rs As DAO.Recordset
    SaveAsset = CanAssets("frmAssets", IIf(AssetID = 0, "ADD", "EDIT"))
    If Len(SaveAsset) = 0 And Len(Trim$(AssetName)) = 0 Then SaveAsset = "«ﬂ » «”„ «·√’·."
    If Len(SaveAsset) = 0 And AssetID > 0 Then SaveAsset = AssetLocked(AssetID)
    If Len(SaveAsset) = 0 Then
        SaveAsset = AssetProblem(AssetID, AccountCode, PurchaseDate, Cost, InputVAT, Salvage, LifeMonths, DepStart, SourceType, _
                                 BankID, CashBoxID, CounterAccount, OpeningAccum, DepMethod, DepRate)
    End If
    If Len(SaveAsset) > 0 Then Exit Function
    Dim auditBefore As Collection
    If AssetID = 0 Then
        Set rs = CurrentDb.OpenRecordset("FixedAssets", dbOpenDynaset, dbSeeChanges)
        rs.AddNew
        rs!AssetCode = NextNumber("FIXED_ASSET")
        rs!Status = "ACTIVE"
    Else
        Set auditBefore = AuditSnapshot("FixedAssets", "AssetID", AssetID)       ' modAudit
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM FixedAssets WHERE AssetID = " & AssetID, dbOpenDynaset, dbSeeChanges)
        rs.Edit
    End If
    rs!AssetName = Left$(Trim$(AssetName), 150)
    rs!AssetAccount = CLng(AccountCode)
    rs!PurchaseDate = DateValue(PurchaseDate)
    rs!Cost = Cost
    rs!InputVAT = InputVAT
    rs!SalvageValue = Salvage
    rs!UsefulLifeMonths = LifeMonths
    rs!DepStartDate = DateValue(DepStart)
    rs!DepMethod = DepMethod
    rs!DepRate = DepRate
    rs!SourceType = SourceType
    rs!BankID = IIf(SourceType = "BANK", BankID, Null)
    rs!CashBoxID = IIf(SourceType = "CASHBOX", CashBoxID, Null)
    rs!CounterAccount = IIf(SourceType = "ACCOUNT", CounterAccount, Null)
    rs!OpeningAccumDep = OpeningAccum
    rs!Notes = IIf(Len(Trim$(Notes)) > 0, Left$(Trim$(Notes), 255), Null)
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Bookmark = rs.LastModified
    AssetID = rs!AssetID
    rs.Close
    LogAction "FIXED_ASSET", "FixedAssets", CStr(AssetID), AssetName
    If Not auditBefore Is Nothing Then AuditEdited "FIXED_ASSET_EDIT", "FixedAssets", "AssetID", AssetID, auditBefore
    SyncJournal                                        ' the purchase entry
End Function

Public Function DeleteAsset(ByVal AssetID As Long) As String
    DeleteAsset = CanAssets("frmAssets", "DELETE")
    If Len(DeleteAsset) = 0 Then DeleteAsset = AssetLocked(AssetID)
    If Len(DeleteAsset) > 0 Then Exit Function
    Dim auditBefore As Collection
    Set auditBefore = AuditSnapshot("FixedAssets", "AssetID", AssetID)        ' the record as it was (modAudit)
    CurrentDb.Execute "DELETE FROM FixedAssets WHERE AssetID = " & AssetID, dbFailOnError + dbSeeChanges
    AuditDeleted "FIXED_ASSET_DELETE", "FixedAssets", AssetID, auditBefore
    DeleteAsset = SyncJournal()
End Function

Public Function DisposeAsset(ByVal AssetID As Long, ByVal DisposalDate As Variant, ByVal Proceeds As Currency, _
                             ByVal DisposalTo As String, ByVal BankID As Variant, ByVal CashBoxID As Variant, _
                             Optional ByVal WithVAT As Boolean = False) As String
    ' Sold or scrapped: no depreciation from the month of DisposalDate. WithVAT: the price is before the
    ' output VAT, which the buyer pays on top (to the bank / cash box) and goes to 2200.
    Dim later As Variant, vat As Currency
    DisposeAsset = CanAssets("frmAssets", "EDIT")
    If Len(DisposeAsset) > 0 Then Exit Function
    If Nz(AssetField(AssetID, "Status"), "") <> "ACTIVE" Then
        DisposeAsset = "«·√’· €Ì— „ÊÃÊœ √Ê „” »⁄œ."
    ElseIf Not IsDate(DisposalDate) Then
        DisposeAsset = "«ﬂ »  «—ÌŒ «·»Ì⁄ √Ê «·«” »⁄«œ."
    ElseIf DateValue(DisposalDate) > Date Then
        DisposeAsset = "«· «—ÌŒ »⁄œ «·ÌÊ„."
    ElseIf DateValue(DisposalDate) < DateValue(AssetField(AssetID, "PurchaseDate")) Then
        DisposeAsset = "«· «—ÌŒ ﬁ»·  «—ÌŒ «·‘—«¡."
    ElseIf Proceeds < 0 Then
        DisposeAsset = "À„‰ «·»Ì⁄ ·« ÌﬂÊ‰ ”«·»«."
    ElseIf DisposalTo = "NONE" And Proceeds <> 0 Then
        DisposeAsset = "«Œ — √Ì‰ «” ı·„ À„‰ «·»Ì⁄ (»‰ﬂ √Ê ’‰œÊﬁ)."
    ElseIf DisposalTo = "BANK" And IsNull(BankID) Then
        DisposeAsset = "«Œ — «·»‰ﬂ."
    ElseIf DisposalTo = "CASHBOX" And IsNull(CashBoxID) Then
        DisposeAsset = "«Œ — «·’‰œÊﬁ."
    ElseIf DisposalTo <> "NONE" And DisposalTo <> "BANK" And DisposalTo <> "CASHBOX" Then
        DisposeAsset = "«Œ — ÿ—Ìﬁ… «·«” »⁄«œ."
    ElseIf DisposalTo <> "NONE" And Proceeds = 0 Then
        DisposeAsset = "«ﬂ » À„‰ «·»Ì⁄° √Ê «Œ — ´«” »⁄«œ »œÊ‰ À„‰ª."
    ElseIf WithVAT And Proceeds = 0 Then
        DisposeAsset = "«·÷—Ì»… ⁄·Ï À„‰ «·»Ì⁄ ›ﬁÿ."
    End If
    If Len(DisposeAsset) = 0 Then
        later = DepreciatedFrom(AssetID, DateValue(DisposalDate))
        If Not IsNull(later) Then DisposeAsset = "··√’· ≈Â·«ﬂ „”Ã· Õ Ï " & GDate(later) & _
                                                 ": «Õ–› ﬁÌÊœ «·≈Â·«ﬂ „‰ ‘Â— «·«” »⁄«œ √Ê·«."
    End If
    If Len(DisposeAsset) = 0 Then
        later = LastAssetEvent(AssetID)
        If Not IsNull(later) Then
            If DateValue(DisposalDate) < later Then DisposeAsset = "··√’· ≈⁄«œ…  ﬁÌÌ„ √Ê ‰ﬁ· ›Ì " & GDate(later) & _
                                                                   ": «· «—ÌŒ ·« ÌﬂÊ‰ ﬁ»·Â."
        End If
    End If
    If Len(DisposeAsset) = 0 Then DisposeAsset = ClosedPeriodProblem(DisposalDate)
    If Len(DisposeAsset) > 0 Then Exit Function
    If WithVAT Then vat = RoundMoney(Proceeds * Nz(SettingValue("VATRate"), 0))
    CurrentDb.Execute "UPDATE FixedAssets SET Status = 'DISPOSED', DisposalDate = " & SqlDate(DateValue(DisposalDate)) & _
        ", DisposalProceeds = " & Str$(Proceeds) & ", DisposalVAT = " & Str$(vat) & ", DisposalTo = " & SqlText(DisposalTo) & ", DisposalBankID = " & _
        IIf(DisposalTo = "BANK", CLng(Nz(BankID, 0)), "Null") & ", DisposalCashBoxID = " & _
        IIf(DisposalTo = "CASHBOX", CLng(Nz(CashBoxID, 0)), "Null") & ", DisposalAccumDep = " & Str$(AssetDepreciated(AssetID)) & _
        " WHERE AssetID = " & AssetID, dbFailOnError + dbSeeChanges
    LogAction "FIXED_ASSET_DISPOSE", "FixedAssets", CStr(AssetID), Format$(Proceeds, "0.00")
    DisposeAsset = SyncJournal()
End Function

Public Function UndoDisposal(ByVal AssetID As Long) As String
    UndoDisposal = CanAssets("frmAssets", "EDIT")
    If Len(UndoDisposal) > 0 Then Exit Function
    If Nz(AssetField(AssetID, "Status"), "") <> "DISPOSED" Then
        UndoDisposal = "«·√’· €Ì— „” »⁄œ."
        Exit Function
    End If
    UndoDisposal = ClosedPeriodProblem(AssetField(AssetID, "DisposalDate"))
    If Len(UndoDisposal) > 0 Then Exit Function
    CurrentDb.Execute "UPDATE FixedAssets SET Status = 'ACTIVE', DisposalDate = Null, DisposalProceeds = 0, DisposalTo = Null, " & _
                      "DisposalBankID = Null, DisposalCashBoxID = Null, DisposalAccumDep = 0, DisposalVAT = 0 WHERE AssetID = " & AssetID, dbFailOnError + dbSeeChanges
    LogAction "FIXED_ASSET_UNDISPOSE", "FixedAssets", CStr(AssetID)
    UndoDisposal = SyncJournal()
End Function

'------------------------------------------------------------------------------
' Monthly depreciation
'------------------------------------------------------------------------------
Public Function LastRunMonth() As Variant
    LastRunMonth = DbValue("SELECT Max(RunMonth) FROM DepreciationRuns")
End Function

Public Function NextRunMonth() As Variant
    ' The month to record next: after the last one, else the first month an asset starts (Null = none).
    Dim last As Variant, first As Variant
    last = LastRunMonth()
    If Not IsNull(last) Then
        NextRunMonth = MonthEnd(DateAdd("m", 1, DateSerial(Year(last), Month(last), 1)))
    Else
        first = DbValue("SELECT Min(DepStartDate) FROM FixedAssets")
        If Not IsNull(first) Then NextRunMonth = MonthEnd(first)
    End If
End Function

Public Function DepreciationFor(ByVal AssetID As Long, ByVal RunMonth As Date) As Currency
    ' The depreciation of the asset in the month that ends on RunMonth (with the entries before it).
    ' After a revaluation: from its new value over the months left, starting in its month.
    Dim rs As DAO.Recordset, lr As DAO.Recordset, remaining As Currency, monthly As Currency, baseValue As Currency
    Dim startMonth As Date, months As Long, yearRate As Double
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM FixedAssets WHERE AssetID = " & AssetID, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Exit Function
    End If
    If DateValue(rs!DepStartDate) <= RunMonth And DateValue(rs!PurchaseDate) <= RunMonth And _
       (rs!Status = "ACTIVE" Or Nz(rs!DisposalDate, 0) > RunMonth) Then
        remaining = AssetValue(AssetID) - rs!SalvageValue - AssetDepreciated(AssetID)
        Set lr = CurrentDb.OpenRecordset("SELECT TOP 1 RevalDate, NewValue, MonthsLeft FROM AssetRevaluations WHERE AssetID = " & _
                                         AssetID & " AND RevalDate <= " & SqlDate(RunMonth) & _
                                         " ORDER BY RevalDate DESC, RevaluationID DESC", dbOpenSnapshot)
        If lr.EOF Then
            startMonth = DateSerial(Year(rs!DepStartDate), Month(rs!DepStartDate), 1)
            months = rs!UsefulLifeMonths
            baseValue = rs!Cost
        Else
            startMonth = DateSerial(Year(lr!RevalDate), Month(lr!RevalDate), 1)
            months = lr!MonthsLeft
            baseValue = lr!NewValue
        End If
        lr.Close
        If Nz(rs!DepMethod, "SL") = "DB" Then
            yearRate = Nz(rs!DepRate, 0)
            If yearRate <= 0 Then yearRate = 24 / rs!UsefulLifeMonths          ' twice the straight line rate
            monthly = Round(CDbl(remaining + rs!SalvageValue) * yearRate / 12, 2)  ' on the book value at the start
            If RunMonth >= MonthEnd(DateAdd("m", months - 1, startMonth)) Then monthly = remaining   ' its last month
        Else
            monthly = MonthlyDepreciation(baseValue, rs!SalvageValue, months)
        End If
        If monthly > remaining Then monthly = remaining
        If remaining - monthly < 1 Then monthly = remaining           ' the rounding cents in the last month
        If monthly > 0 Then DepreciationFor = monthly
    End If
    rs.Close
End Function

Public Function RecordDepreciation(ByVal RunMonth As Date, ByRef RunID As Long) As String
    ' The entry of one month (the next one to record).
    Dim db As DAO.Database, rs As DAO.Recordset, d As DAO.Recordset, nextMonth As Variant, amount As Currency
    Dim total As Currency, n As Long
    RunID = 0
    RecordDepreciation = CanAssets("frmDepreciation", "ADD")
    If Len(RecordDepreciation) > 0 Then Exit Function
    RunMonth = MonthEnd(RunMonth)
    nextMonth = NextRunMonth()
    If IsNull(nextMonth) Then
        RecordDepreciation = "·«  ÊÃœ √’Ê· À«» …."
    ElseIf RunMonth <> nextMonth Then
        RecordDepreciation = "«·‘Â— «· «·Ì ··≈Â·«ﬂ ÂÊ " & Format$(nextMonth, "yyyy/mm") & ": «·√‘Â—  ı”ÃÛ¯· »«· — Ì»."
    ElseIf RunMonth > Date Then
        RecordDepreciation = "Ìı”ÃÛ¯· ≈Â·«ﬂ «·‘Â— ›Ì ¬Œ— ÌÊ„ „‰Â √Ê »⁄œÂ."
    Else
        RecordDepreciation = ClosedPeriodProblem(RunMonth)
    End If
    If Len(RecordDepreciation) > 0 Then Exit Function
    Set db = CurrentDb
    Set rs = db.OpenRecordset("DepreciationRuns", dbOpenDynaset, dbSeeChanges)
    rs.AddNew
    rs!RunNumber = NextNumber("DEPRECIATION")
    rs!RunMonth = RunMonth
    rs!TotalAmount = 0
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Bookmark = rs.LastModified
    RunID = rs!RunID
    rs.Close
    Set rs = db.OpenRecordset("SELECT AssetID, CostCenterID FROM FixedAssets ORDER BY AssetID", dbOpenSnapshot)
    Set d = db.OpenRecordset("AssetDepreciations", dbOpenDynaset, dbAppendOnly + dbSeeChanges)
    Do Until rs.EOF
        amount = DepreciationFor(rs!AssetID, RunMonth)
        If amount > 0 Then
            n = n + 1
            d.AddNew
            d!RunID = RunID
            d!LineNo = n
            d!AssetID = rs!AssetID
            d!Amount = amount
            d!DepCenter = Nz(rs!CostCenterID, 0)               ' stays there after a transfer
            d.Update
            total = total + amount
        End If
        rs.MoveNext
    Loop
    rs.Close
    d.Close
    db.Execute "UPDATE DepreciationRuns SET TotalAmount = " & Str$(total) & " WHERE RunID = " & RunID, dbFailOnError + dbSeeChanges
    LogAction "DEPRECIATION", "DepreciationRuns", Format$(RunMonth, "yyyy-mm"), Format$(total, "0.00")
    RecordDepreciation = SyncJournal()
End Function

Public Function RecordDepreciationThrough(ByVal LastMonth As Date, ByRef Months As Long) As String
    ' Every month not recorded yet until LastMonth.
    Dim m As Variant, id As Long, msg As String
    Months = 0
    m = NextRunMonth()
    Do While Not IsNull(m)
        If m > MonthEnd(LastMonth) Then Exit Do
        msg = RecordDepreciation(m, id)
        If id = 0 Then
            RecordDepreciationThrough = msg
            Exit Function
        End If
        Months = Months + 1
        m = NextRunMonth()
    Loop
    If Months = 0 Then RecordDepreciationThrough = "·«  ÊÃœ √‘Â— ·„  ı”ÃÛ¯· Õ Ï " & Format$(LastMonth, "yyyy/mm") & "."
End Function

Public Function DeleteLastRun() As String
    Dim last As Variant
    DeleteLastRun = CanAssets("frmDepreciation", "DELETE")
    If Len(DeleteLastRun) > 0 Then Exit Function
    last = DbValue("SELECT TOP 1 RunID FROM DepreciationRuns ORDER BY RunMonth DESC")
    If IsNull(last) Then
        DeleteLastRun = "·«  ÊÃœ ﬁÌÊœ ≈Â·«ﬂ."
        Exit Function
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM AssetDepreciations AS d INNER JOIN FixedAssets AS a ON d.AssetID = a.AssetID " & _
                  "WHERE d.RunID = " & last & " AND a.Status = 'DISPOSED'"), 0) > 0 Then
        DeleteLastRun = "›Ì «·ﬁÌœ √’· „” »⁄œ »⁄œÂ: √·€ˆ «·«” »⁄«œ √Ê·«."
        Exit Function
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM AssetRevaluations AS v, DepreciationRuns AS r WHERE r.RunID = " & last & _
                  " AND v.RevalDate > r.RunMonth AND v.AssetID IN (SELECT AssetID FROM AssetDepreciations WHERE RunID = " & _
                  last & ")"), 0) + _
       Nz(DbValue("SELECT COUNT(*) FROM AssetTransfers AS v, DepreciationRuns AS r WHERE r.RunID = " & last & _
                  " AND v.TransferDate > r.RunMonth AND v.AssetID IN (SELECT AssetID FROM AssetDepreciations WHERE RunID = " & _
                  last & ")"), 0) > 0 Then
        DeleteLastRun = "·√’· ›Ì «·ﬁÌœ ≈⁄«œ…  ﬁÌÌ„ √Ê ‰ﬁ· »⁄œÂ: «Õ–›Â« √Ê·«."
        Exit Function
    End If
    DeleteLastRun = ClosedPeriodProblem(DbValue("SELECT RunMonth FROM DepreciationRuns WHERE RunID = " & last))
    If Len(DeleteLastRun) > 0 Then Exit Function
    Dim auditBefore As Collection
    Set auditBefore = AuditSnapshot("DepreciationRuns", "RunID", last)        ' the record as it was (modAudit)
    CurrentDb.Execute "DELETE FROM DepreciationRuns WHERE RunID = " & last, dbFailOnError + dbSeeChanges      ' lines cascade
    AuditDeleted "DEPRECIATION_DELETE", "DepreciationRuns", last, auditBefore
    DeleteLastRun = SyncJournal()
End Function

'------------------------------------------------------------------------------
' Revaluation and transfer between cost centres (docs/65)
'------------------------------------------------------------------------------
Private Function AssetEventProblem(ByVal AssetID As Long, ByVal EventDate As Variant) As String
    ' A revaluation or a transfer: an asset in use, a day it owned it, no depreciation recorded from that
    ' month on, not before its last revaluation or transfer, in an open period.
    Dim p As String, later As Variant
    If Nz(AssetField(AssetID, "Status"), "") <> "ACTIVE" Then
        p = "«Œ — √’·« ﬁ«∆„« (€Ì— „” »⁄œ)."
    ElseIf Not IsDate(EventDate) Then
        p = "«ﬂ » «· «—ÌŒ."
    ElseIf DateValue(EventDate) > Date Then
        p = "«· «—ÌŒ »⁄œ «·ÌÊ„."
    ElseIf DateValue(EventDate) < DateValue(AssetField(AssetID, "PurchaseDate")) Then
        p = "«· «—ÌŒ ﬁ»·  «—ÌŒ «·‘—«¡."
    End If
    If Len(p) = 0 Then
        later = DepreciatedFrom(AssetID, DateValue(EventDate))
        If Not IsNull(later) Then p = "··√’· ≈Â·«ﬂ „”Ã· Õ Ï " & GDate(later) & _
                                     ": «Õ–› ﬁÌÊœ «·≈Â·«ﬂ „‰ ‘Â— Â–« «· «—ÌŒ √Ê·«."
    End If
    If Len(p) = 0 Then
        later = LastAssetEvent(AssetID)
        If Not IsNull(later) Then
            If DateValue(EventDate) < later Then p = "··√’· ≈⁄«œ…  ﬁÌÌ„ √Ê ‰ﬁ· ›Ì " & GDate(later) & ": «· «—ÌŒ ·« ÌﬂÊ‰ ﬁ»·Â."
        End If
    End If
    If Len(p) = 0 Then p = ClosedPeriodProblem(EventDate)
    AssetEventProblem = p
End Function

Public Function RevalMonthsLeft(ByVal AssetID As Long, ByVal OnDay As Date) As Long
    ' The months of life left from the month of OnDay (the default of a revaluation, at least 1).
    Dim n As Long, rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 RevalDate, MonthsLeft FROM AssetRevaluations WHERE AssetID = " & AssetID & _
                                     " ORDER BY RevalDate DESC, RevaluationID DESC", dbOpenSnapshot)
    If rs.EOF Then
        n = Nz(AssetField(AssetID, "UsefulLifeMonths"), 0) - DateDiff("m", Nz(AssetField(AssetID, "DepStartDate"), OnDay), OnDay)
    Else
        n = rs!MonthsLeft - DateDiff("m", rs!RevalDate, OnDay)
    End If
    rs.Close
    If n < 1 Then n = 1
    RevalMonthsLeft = n
End Function

Public Function RevalueAsset(ByVal AssetID As Long, ByVal RevalDate As Variant, ByVal NewValue As Currency, _
                             ByVal MonthsLeft As Long, ByVal Notes As String, ByRef RevalID As Long) As String
    ' A rise first takes back the impairment losses of the asset (5660), the rest goes to the surplus (3500);
    ' a fall first uses the surplus of the asset, the rest is a loss (5660).
    Dim p As String, before As Currency, amount As Currency, part As Currency, surplus As Currency, loss As Currency
    Dim rs As DAO.Recordset
    RevalID = 0
    p = CanAssets("frmAssetChanges", "ADD")
    If Len(p) = 0 Then p = AssetEventProblem(AssetID, RevalDate)
    If Len(p) = 0 Then
        before = AssetBookValue(AssetID)
        If NewValue < Nz(AssetField(AssetID, "SalvageValue"), 0) Then
            p = "«·ﬁÌ„… «·ÃœÌœ… √ﬁ· „‰ «·ﬁÌ„… «·„ »ﬁÌ… ··√’·."
        ElseIf NewValue = before Then
            p = "«·ﬁÌ„… «·ÃœÌœ…  ”«ÊÌ «·ﬁÌ„… «·œ› —Ì… " & Format$(before, "#,##0.00") & "."
        ElseIf MonthsLeft <= 0 Then
            p = "«·⁄„— «·»«ﬁÌ ‘Â— ⁄·Ï «·√ﬁ·."
        End If
    End If
    If Len(p) > 0 Then
        RevalueAsset = p
        Exit Function
    End If
    amount = NewValue - before
    If amount > 0 Then
        part = Nz(DbValue("SELECT Sum(LossPart) FROM AssetRevaluations WHERE AssetID = " & AssetID), 0)
        If part > amount Then part = amount
        If part < 0 Then part = 0
        loss = -part
        surplus = amount - part
    Else
        part = Nz(DbValue("SELECT Sum(SurplusPart) FROM AssetRevaluations WHERE AssetID = " & AssetID), 0)
        If part > -amount Then part = -amount
        If part < 0 Then part = 0
        surplus = -part
        loss = -amount - part
    End If
    Set rs = CurrentDb.OpenRecordset("AssetRevaluations", dbOpenDynaset, dbSeeChanges)
    rs.AddNew
    rs!RevalNumber = NextNumber("ASSET_REVAL")
    rs!AssetID = AssetID
    rs!RevalDate = DateValue(RevalDate)
    rs!BookValueBefore = before
    rs!NewValue = NewValue
    rs!ValueChange = amount
    rs!SurplusPart = surplus
    rs!LossPart = loss
    rs!MonthsLeft = MonthsLeft
    rs!RevalCenter = Nz(AssetField(AssetID, "CostCenterID"), 0)
    rs!Notes = IIf(Len(Trim$(Notes)) > 0, Left$(Trim$(Notes), 255), Null)
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Bookmark = rs.LastModified
    RevalID = rs!RevaluationID
    rs.Close
    LogAction "ASSET_REVAL", "AssetRevaluations", CStr(RevalID), Format$(amount, "0.00")
    RevalueAsset = SyncJournal()
End Function

Public Function DeleteLastRevaluation(ByVal AssetID As Long) As String
    Dim id As Variant, rs As DAO.Recordset, p As String, later As Variant, auditBefore As Collection
    p = CanAssets("frmAssetChanges", "DELETE")
    If Len(p) = 0 And Nz(AssetField(AssetID, "Status"), "") <> "ACTIVE" Then p = "«·√’· „” »⁄œ: √·€ˆ «·«” »⁄«œ √Ê·«."
    If Len(p) = 0 Then
        id = DbValue("SELECT TOP 1 RevaluationID FROM AssetRevaluations WHERE AssetID = " & AssetID & _
                     " ORDER BY RevalDate DESC, RevaluationID DESC")
        If IsNull(id) Then p = "·«  ÊÃœ ≈⁄«œ…  ﬁÌÌ„ ··√’·."
    End If
    If Len(p) = 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT RevalDate FROM AssetRevaluations WHERE RevaluationID = " & id, _
                                         dbOpenSnapshot, dbSeeChanges)
        later = DepreciatedFrom(AssetID, rs!RevalDate)
        If Not IsNull(later) Then
            p = "··√’· ≈Â·«ﬂ „”Ã· Õ Ï " & GDate(later) & ": «Õ–› ﬁÌÊœ «·≈Â·«ﬂ „‰ ‘Â— ≈⁄«œ… «· ﬁÌÌ„ √Ê·«."
        ElseIf Nz(DbValue("SELECT COUNT(*) FROM AssetTransfers WHERE AssetID = " & AssetID & " AND TransferDate >= " & _
                          SqlDate(rs!RevalDate)), 0) > 0 Then
            p = "··√’· ‰ﬁ· ›Ì ÌÊ„ ≈⁄«œ… «· ﬁÌÌ„ √Ê »⁄œÂ: «Õ–› «·‰ﬁ· √Ê·«."
        Else
            p = ClosedPeriodProblem(rs!RevalDate)
        End If
        rs.Close
    End If
    If Len(p) > 0 Then
        DeleteLastRevaluation = p
        Exit Function
    End If
    Set auditBefore = AuditSnapshot("AssetRevaluations", "RevaluationID", id)      ' modAudit
    CurrentDb.Execute "DELETE FROM AssetRevaluations WHERE RevaluationID = " & id, dbFailOnError + dbSeeChanges
    AuditDeleted "ASSET_REVAL_DELETE", "AssetRevaluations", id, auditBefore
    DeleteLastRevaluation = SyncJournal()
End Function

Public Function TransferAsset(ByVal AssetID As Long, ByVal TransferDate As Variant, ByVal ToCenter As Variant, _
                              ByVal Notes As String, ByRef TransferID As Long) As String
    ' The asset goes to another branch / cost centre (Null = none): its value and accumulated depreciation
    ' move with an entry, its depreciation from the month of the transfer goes to the new centre.
    Dim p As String, fromCenter As Variant, rs As DAO.Recordset
    TransferID = 0
    p = CanAssets("frmAssetChanges", "ADD")
    If Len(p) = 0 Then p = AssetEventProblem(AssetID, TransferDate)
    If Len(p) = 0 Then
        fromCenter = AssetField(AssetID, "CostCenterID")
        If Nz(fromCenter, 0) = Nz(ToCenter, 0) Then
            p = "«·√’· ⁄·Ï Â–« «·„—ﬂ“."
        ElseIf Not IsNull(ToCenter) Then
            If Nz(DbValue("SELECT COUNT(*) FROM CostCenters WHERE IsActive <> 0 AND CostCenterID = " & CLng(ToCenter)), 0) = 0 Then
                p = "„—ﬂ“ «· ﬂ·›… €Ì— ‰‘ÿ."
            End If
        End If
    End If
    If Len(p) > 0 Then
        TransferAsset = p
        Exit Function
    End If
    ' the depreciation recorded before stays on the centre it was on
    CurrentDb.Execute "UPDATE AssetDepreciations SET DepCenter = " & CLng(Nz(fromCenter, 0)) & " WHERE AssetID = " & AssetID & _
                      " AND DepCenter Is Null", dbFailOnError + dbSeeChanges
    Set rs = CurrentDb.OpenRecordset("AssetTransfers", dbOpenDynaset, dbSeeChanges)
    rs.AddNew
    rs!TransferNumber = NextNumber("ASSET_TRANSFER")
    rs!AssetID = AssetID
    rs!TransferDate = DateValue(TransferDate)
    rs!FromCenterID = fromCenter
    rs!ToCenterID = ToCenter
    rs!AssetValue = AssetValue(AssetID)
    rs!AccumDep = AssetDepreciated(AssetID)
    rs!Notes = IIf(Len(Trim$(Notes)) > 0, Left$(Trim$(Notes), 255), Null)
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Bookmark = rs.LastModified
    TransferID = rs!TransferID
    rs.Close
    CurrentDb.Execute "UPDATE FixedAssets SET CostCenterID = " & IIf(IsNull(ToCenter), "Null", CLng(Nz(ToCenter, 0))) & _
                      " WHERE AssetID = " & AssetID, dbFailOnError + dbSeeChanges
    LogAction "ASSET_TRANSFER", "AssetTransfers", CStr(TransferID), CStr(Nz(ToCenter, 0))
    TransferAsset = SyncJournal()
End Function

Public Function DeleteLastTransfer(ByVal AssetID As Long) As String
    ' The asset goes back to the centre it came from.
    Dim id As Variant, rs As DAO.Recordset, p As String, later As Variant, fromCenter As Variant, auditBefore As Collection
    p = CanAssets("frmAssetChanges", "DELETE")
    If Len(p) = 0 And Nz(AssetField(AssetID, "Status"), "") <> "ACTIVE" Then p = "«·√’· „” »⁄œ: √·€ˆ «·«” »⁄«œ √Ê·«."
    If Len(p) = 0 Then
        id = DbValue("SELECT TOP 1 TransferID FROM AssetTransfers WHERE AssetID = " & AssetID & _
                     " ORDER BY TransferDate DESC, TransferID DESC")
        If IsNull(id) Then p = "·« ÌÊÃœ ‰ﬁ· ··√’·."
    End If
    If Len(p) = 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT TransferDate, FromCenterID FROM AssetTransfers WHERE TransferID = " & _
                                         id, dbOpenSnapshot)
        fromCenter = rs!FromCenterID
        later = DepreciatedFrom(AssetID, rs!TransferDate)
        If Not IsNull(later) Then
            p = "··√’· ≈Â·«ﬂ „”Ã· Õ Ï " & GDate(later) & ": «Õ–› ﬁÌÊœ «·≈Â·«ﬂ „‰ ‘Â— «·‰ﬁ· √Ê·«."
        ElseIf Nz(DbValue("SELECT COUNT(*) FROM AssetRevaluations WHERE AssetID = " & AssetID & " AND RevalDate > " & _
                          SqlDate(rs!TransferDate)), 0) > 0 Then
            p = "··√’· ≈⁄«œ…  ﬁÌÌ„ »⁄œ «·‰ﬁ·: «Õ–›Â« √Ê·«."
        Else
            p = ClosedPeriodProblem(rs!TransferDate)
        End If
        rs.Close
    End If
    If Len(p) > 0 Then
        DeleteLastTransfer = p
        Exit Function
    End If
    Set auditBefore = AuditSnapshot("AssetTransfers", "TransferID", id)            ' modAudit
    CurrentDb.Execute "DELETE FROM AssetTransfers WHERE TransferID = " & id, dbFailOnError + dbSeeChanges
    CurrentDb.Execute "UPDATE FixedAssets SET CostCenterID = " & IIf(IsNull(fromCenter), "Null", CLng(Nz(fromCenter, 0))) & _
                      " WHERE AssetID = " & AssetID, dbFailOnError + dbSeeChanges
    AuditDeleted "ASSET_TRANSFER_DELETE", "AssetTransfers", id, auditBefore
    DeleteLastTransfer = SyncJournal()
End Function

'------------------------------------------------------------------------------
' Screen frmAssets (OpenArgs: the asset to show)
'------------------------------------------------------------------------------
Private Function ShownAsset(ByVal frm As Access.Form) As Long
    ShownAsset = Nz(frm!txtAssetID.Value, 0)
End Function

Public Sub AssetsLoad(ByVal frm As Access.Form)
    Dim msg As String
    Calendar = vbCalGreg
    msg = SyncJournal()
    If Len(msg) > 0 Then ShowWarning msg
    AssetNew frm
    If Nz(frm.OpenArgs, 0) > 0 Then AssetShow frm, CLng(frm.OpenArgs)
End Sub

Public Sub AssetNew(ByVal frm As Access.Form)
    frm!txtAssetID.Value = Null
    frm!txtAssetName.Value = Null
    frm!cboAssetAccount.Value = Null
    frm!txtPurchaseDate.Value = Date
    frm!txtDepStart.Value = Date
    frm!txtCost.Value = Null
    frm!txtInputVAT.Value = Null
    frm!txtSalvage.Value = 0
    frm!txtLife.Value = 60
    frm!cboDepMethod.Value = "SL"
    frm!txtDepRate.Value = 0
    frm!cboSource.Value = "BANK"
    frm!cboBank.Value = SettingValue("DefaultBankID")
    frm!cboBox.Value = Null
    frm!cboCounter.Value = Null
    frm!txtOpeningAccum.Value = 0
    frm!txtNotes.Value = Null
    frm!cboCenter.Value = CostCenterFor()
    frm!txtDisposalDate.Value = Date
    frm!txtProceeds.Value = Null
    frm!cboDisposalTo.Value = "BANK"
    frm!chkDisposalVAT.Value = False
    AssetSourceChanged frm
    AssetRefresh frm
End Sub

Public Sub AssetSourceChanged(ByVal frm As Access.Form)
    Dim s As String
    s = Nz(frm!cboSource.Value, "")
    frm!cboBank.Enabled = (s = "BANK")
    frm!cboBox.Enabled = (s = "CASHBOX")
    frm!cboCounter.Enabled = (s = "ACCOUNT")
    frm!txtOpeningAccum.Enabled = (s = "OPENING")
    frm!txtInputVAT.Enabled = (s <> "OPENING")
    frm!txtDepRate.Enabled = (Nz(frm!cboDepMethod.Value, "SL") = "DB")
End Sub

Public Sub AssetMethodChanged(ByVal frm As Access.Form)
    If Nz(frm!cboDepMethod.Value, "SL") = "SL" Then frm!txtDepRate.Value = 0
    AssetSourceChanged frm
    AssetRefresh frm
End Sub

Public Sub AssetShow(ByVal frm As Access.Form, ByVal AssetID As Long)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM FixedAssets WHERE AssetID = " & AssetID, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If
    frm!txtAssetID.Value = AssetID
    frm!txtAssetName.Value = rs!AssetName
    frm!cboAssetAccount.Value = rs!AssetAccount
    frm!txtPurchaseDate.Value = rs!PurchaseDate
    frm!txtDepStart.Value = rs!DepStartDate
    frm!txtCost.Value = rs!Cost
    frm!txtInputVAT.Value = rs!InputVAT
    frm!txtSalvage.Value = rs!SalvageValue
    frm!txtLife.Value = rs!UsefulLifeMonths
    frm!cboDepMethod.Value = Nz(rs!DepMethod, "SL")
    frm!txtDepRate.Value = Nz(rs!DepRate, 0)
    frm!cboSource.Value = rs!SourceType
    frm!cboBank.Value = rs!BankID
    frm!cboBox.Value = rs!CashBoxID
    frm!cboCounter.Value = rs!CounterAccount
    frm!txtOpeningAccum.Value = rs!OpeningAccumDep
    frm!txtNotes.Value = rs!Notes
    frm!cboCenter.Value = rs!CostCenterID
    If rs!Status = "DISPOSED" Then
        frm!txtDisposalDate.Value = rs!DisposalDate
        frm!txtProceeds.Value = rs!DisposalProceeds
        frm!cboDisposalTo.Value = rs!DisposalTo
        frm!chkDisposalVAT.Value = (Nz(rs!DisposalVAT, 0) <> 0)
    End If
    rs.Close
    AssetSourceChanged frm
    AssetRefresh frm
End Sub

Public Sub AssetRefresh(ByVal frm As Access.Form)
    Dim id As Long, info As String, rs As DAO.Recordset
    id = ShownAsset(frm)
    frm!lstAssets.Requery
    If id = 0 Then
        frm!lblAssetInfo.Caption = Tr("√’· ÃœÌœ: ﬁ”ÿ «·‘Â— «·√Ê· " & Format$(FirstMonthDepreciation(Nz(frm!cboDepMethod.Value, "SL"), _
            Nz(frm!txtDepRate.Value, 0), Nz(frm!txtCost.Value, 0), Nz(frm!txtSalvage.Value, 0), Nz(frm!txtLife.Value, 0)), _
            "#,##0.00"))
        frm!lblAssetInfo.ForeColor = CLR_PRIMARY
        Exit Sub
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM FixedAssetsQuery WHERE AssetID = " & id, dbOpenSnapshot)
    If Not rs.EOF Then
        info = rs!AssetCode & "   „Ã„⁄ «·≈Â·«ﬂ " & Format$(rs!AccumDep, "#,##0.00") & "   «·ﬁÌ„… «·œ› —Ì… " & _
               Format$(rs!BookValue, "#,##0.00") & "   ﬁ”ÿ «·‘Â— " & Format$(rs!MonthlyDep, "#,##0.00") & "   (" & _
               rs!DepMonths & " ‘Â—« „”Ã·…° " & rs!MethodName & ")"
        If rs!RevalTotal <> 0 Then info = info & "   ≈⁄«œ… «· ﬁÌÌ„ " & Format$(rs!RevalTotal, "#,##0.00")
        If rs!Status = "DISPOSED" Then
            info = info & vbCrLf & "„” »⁄œ ›Ì " & GDate(rs!DisposalDate) & _
                   IIf(Nz(rs!DisposalProceeds, 0) > 0, " »À„‰ " & Format$(rs!DisposalProceeds, "#,##0.00"), "")
        End If
        frm!lblAssetInfo.ForeColor = IIf(rs!Status = "DISPOSED", CLR_MUTED, CLR_PRIMARY)
    End If
    rs.Close
    frm!lblAssetInfo.Caption = Tr(info)
End Sub

Public Function FirstMonthDepreciation(ByVal DepMethod As String, ByVal DepRate As Double, ByVal Cost As Currency, _
                                       ByVal Salvage As Currency, ByVal Months As Long) As Currency
    ' What a new asset will depreciate in its first month (screen hint).
    Dim yearRate As Double, monthly As Currency
    If Months <= 0 Or Cost <= Salvage Then Exit Function
    If DepMethod = "DB" Then
        yearRate = DepRate
        If yearRate <= 0 Then yearRate = 24 / Months
        monthly = Round(CDbl(Cost) * yearRate / 12, 2)
        If monthly > Cost - Salvage Then monthly = Cost - Salvage
    Else
        monthly = MonthlyDepreciation(Cost, Salvage, Months)
    End If
    FirstMonthDepreciation = monthly
End Function

Public Sub AssetPick(ByVal frm As Access.Form)
    If IsNull(frm!lstAssets.Value) Then Exit Sub
    AssetShow frm, CLng(frm!lstAssets.Value)
End Sub

Public Sub AssetSave(ByVal frm As Access.Form)
    Dim id As Long, msg As String, isNew As Boolean
    id = ShownAsset(frm)
    isNew = (id = 0)
    msg = SaveAsset(id, Nz(frm!txtAssetName.Value, ""), frm!cboAssetAccount.Value, frm!txtPurchaseDate.Value, _
                    CCur(Nz(frm!txtCost.Value, 0)), CCur(Nz(frm!txtInputVAT.Value, 0)), CCur(Nz(frm!txtSalvage.Value, 0)), _
                    CLng(Nz(frm!txtLife.Value, 0)), frm!txtDepStart.Value, Nz(frm!cboSource.Value, ""), frm!cboBank.Value, _
                    frm!cboBox.Value, frm!cboCounter.Value, CCur(Nz(frm!txtOpeningAccum.Value, 0)), Nz(frm!txtNotes.Value, ""), _
                    Nz(frm!cboDepMethod.Value, "SL"), CDbl(Nz(frm!txtDepRate.Value, 0)))
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    ' the cost centre of the asset (modCostCenters): its depreciation and its sale go there
    CurrentDb.Execute "UPDATE FixedAssets SET CostCenterID = " & IIf(IsNull(frm!cboCenter.Value), "Null", _
                      Nz(frm!cboCenter.Value, 0)) & " WHERE AssetID = " & id, dbFailOnError + dbSeeChanges
    SyncJournal
    ShowInfo IIf(isNew, " „ Õ›Ÿ «·√’· ÊﬁÌœ ‘—«∆Â.", " „ Õ›Ÿ «· ⁄œÌ·.")
    AssetShow frm, id
End Sub

Public Sub AssetDelete(ByVal frm As Access.Form)
    Dim msg As String
    If ShownAsset(frm) = 0 Then Exit Sub
    If Not AskYesNo("Õ–› «·√’· ÊﬁÌœ ‘—«∆Âø") Then Exit Sub
    msg = DeleteAsset(ShownAsset(frm))
    If IsNull(AssetField(ShownAsset(frm), "AssetID")) Then
        AssetNew frm
    Else
        ShowWarning msg
    End If
End Sub

Public Sub AssetDispose(ByVal frm As Access.Form)
    Dim msg As String
    If ShownAsset(frm) = 0 Then
        ShowWarning "«Œ — «·√’· „‰ «·ﬁ«∆„…."
        Exit Sub
    End If
    If Not AskYesNo("»Ì⁄ √Ê «” »⁄«œ «·√’· ›Ì " & GDate(Nz(frm!txtDisposalDate.Value, Date)) & "ø") Then Exit Sub
    msg = DisposeAsset(ShownAsset(frm), frm!txtDisposalDate.Value, CCur(Nz(frm!txtProceeds.Value, 0)), _
                       Nz(frm!cboDisposalTo.Value, ""), frm!cboDisposalBank.Value, frm!cboDisposalBox.Value, _
                       Nz(frm!chkDisposalVAT.Value, False))
    If AssetField(ShownAsset(frm), "Status") <> "DISPOSED" Then ShowWarning msg
    AssetShow frm, ShownAsset(frm)
End Sub

Public Sub AssetUndoDisposal(ByVal frm As Access.Form)
    Dim msg As String
    If ShownAsset(frm) = 0 Then Exit Sub
    If Not AskYesNo("≈·€«¡ »Ì⁄ / «” »⁄«œ «·√’·ø") Then Exit Sub
    msg = UndoDisposal(ShownAsset(frm))
    If AssetField(ShownAsset(frm), "Status") <> "ACTIVE" Then ShowWarning msg
    AssetShow frm, ShownAsset(frm)
End Sub

Public Sub AssetOpenChanges(ByVal frm As Access.Form)
    If ShownAsset(frm) = 0 Then
        ShowWarning "«Œ — «·√’· „‰ «·ﬁ«∆„…."
        Exit Sub
    End If
    OpenScreen "frmAssetChanges", 0, ShownAsset(frm)
End Sub

Public Sub PrintAssets(ByVal frm As Access.Form)
    LogAction "REPORT", "FIXED_ASSETS"
    OpenReportOrQuery "rptFixedAssets", "FixedAssetsQuery", "", "›Ì " & GDate(Date)
End Sub

'------------------------------------------------------------------------------
' Screen frmDepreciation
'------------------------------------------------------------------------------
Public Sub DepreciationLoad(ByVal frm As Access.Form)
    Dim msg As String
    Calendar = vbCalGreg
    msg = SyncJournal()
    If Len(msg) > 0 Then ShowWarning msg
    frm!txtThrough.Value = DateSerial(Year(Date), Month(Date), 0)          ' end of last month
    DepreciationRefresh frm
End Sub

Public Sub DepreciationRefresh(ByVal frm As Access.Form)
    Dim nextMonth As Variant, rows As String, rs As DAO.Recordset, amount As Currency, total As Currency
    nextMonth = NextRunMonth()
    If IsNull(nextMonth) Then
        frm!lblNext.Caption = Tr("·«  ÊÃœ √’Ê· À«» ….")
        frm!lstPreview.RowSource = Tr("")
    Else
        ' what the next month will record
        rows = """«·√’·"";""«·≈Â·«ﬂ"""
        Set rs = CurrentDb.OpenRecordset("SELECT AssetID, AssetName FROM FixedAssets ORDER BY AssetID", dbOpenSnapshot)
        Do Until rs.EOF
            amount = DepreciationFor(rs!AssetID, nextMonth)
            If amount > 0 Then
                rows = rows & ";""" & Replace(rs!AssetName, """", "'") & """;""" & Format$(amount, "#,##0.00") & """"
                total = total + amount
            End If
            rs.MoveNext
        Loop
        rs.Close
        frm!lstPreview.RowSource = Tr(rows)
        frm!lblNext.Caption = Tr("«·‘Â— «· «·Ì ·· ”ÃÌ·: " & Format$(nextMonth, "yyyy/mm") & "   «·≈Ã„«·Ì " & Format$(total, "#,##0.00") & _
                              IIf(IsNull(LastRunMonth()), "", "   (¬Œ— ‘Â— „”Ã· " & Format$(Nz(LastRunMonth(), 0), "yyyy/mm") & ")"))
    End If
    frm!lstRuns.Requery
End Sub

Public Sub DepreciationRun(ByVal frm As Access.Form)
    Dim msg As String, n As Long
    If Not IsDate(frm!txtThrough.Value) Then
        ShowWarning "«Œ — ¬Œ— ‘Â— Ìı”ÃÛ¯· ≈Â·«ﬂÂ."
        Exit Sub
    End If
    If Not AskYesNo(" ”ÃÌ· «·≈Â·«ﬂ ·ﬂ· ‘Â— ·„ Ìı”ÃÛ¯· Õ Ï " & Format$(frm!txtThrough.Value, "yyyy/mm") & "ø") Then Exit Sub
    msg = RecordDepreciationThrough(DateValue(frm!txtThrough.Value), n)
    If n = 0 Then
        ShowWarning msg
    Else
        If Len(msg) > 0 Then ShowWarning msg
        ShowInfo " „  ”ÃÌ· ≈Â·«ﬂ " & n & " ‘Â—."
    End If
    DepreciationRefresh frm
End Sub

Public Sub DepreciationUndo(ByVal frm As Access.Form)
    Dim msg As String, before As Variant
    before = LastRunMonth()
    If IsNull(before) Then Exit Sub
    If Not AskYesNo("Õ–› ﬁÌœ ≈Â·«ﬂ ‘Â— " & Format$(before, "yyyy/mm") & "ø") Then Exit Sub
    msg = DeleteLastRun()
    If Nz(LastRunMonth(), 0) = before Then ShowWarning msg
    DepreciationRefresh frm
End Sub

'------------------------------------------------------------------------------
' Screen frmAssetChanges (OpenArgs: the asset): revaluations and transfers
'------------------------------------------------------------------------------
Private Function ChangesAsset(ByVal frm As Access.Form) As Long
    ChangesAsset = Nz(frm!txtAssetID.Value, 0)
End Function

Public Sub AssetChangesLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    frm!txtAssetID.Value = CLng(Nz(frm.OpenArgs, 0))
    frm!txtRevalDate.Value = Date
    frm!txtTransferDate.Value = Date
    AssetChangesRefresh frm
End Sub

Public Sub AssetChangesRefresh(ByVal frm As Access.Form)
    Dim id As Long, rs As DAO.Recordset, info As String, center As Variant, rows As String
    id = ChangesAsset(frm)
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM FixedAssetsQuery WHERE AssetID = " & id, dbOpenSnapshot)
    If rs.EOF Then
        info = "«Œ — «·√’· „‰ ‘«‘… «·√’Ê· «·À«» …."
    Else
        center = DbValue(Tr("SELECT c.CenterName FROM [@CostCenters] AS c INNER JOIN FixedAssets AS a ON c.CostCenterID = " & _
                            "a.CostCenterID WHERE a.AssetID = " & id))
        info = rs!AssetCode & "  " & rs!AssetName & "   «·ﬁÌ„… «·œ› —Ì… " & Format$(rs!BookValue, "#,##0.00") & _
               "   „—ﬂ“ «· ﬂ·›…: " & Nz(center, "-") & "   (" & rs!StatusName & ")"
    End If
    rs.Close
    frm!lblAsset.Caption = Tr(info)
    frm!txtNewValue.Value = Null
    frm!txtMonthsLeft.Value = IIf(id > 0, RevalMonthsLeft(id, Nz(frm!txtRevalDate.Value, Date)), Null)
    rows = Tr("SELECT r.RevaluationID, r.RevalNumber AS [«·—ﬁ„], Format(r.RevalDate, 'yyyy/mm/dd') " & _
        "AS [«· «—ÌŒ], Format(r.BookValueBefore, '#,##0.00') AS [«·ﬁÌ„… «·œ› —Ì… ﬁ»·Â«], Format(r.NewValue, '#,##0.00') AS " & _
        "[«·ﬁÌ„… »⁄œ «· ﬁÌÌ„], Format(r.ValueChange, '#,##0.00') AS [«·›—ﬁ], Format(r.SurplusPart, '#,##0.00') AS [›«∆÷ «· ﬁÌÌ„], " & _
        "Format(r.LossPart, '#,##0.00') AS [Œ”«—… «·«‰Œ›«÷], r.MonthsLeft AS [«·⁄„— «·»«ﬁÌ] FROM AssetRevaluations AS r " & _
        "WHERE r.AssetID = {ID} ORDER BY r.RevalDate DESC, r.RevaluationID DESC")
    frm!lstRevals.RowSource = Replace(rows, "{ID}", CStr(id))
    rows = Tr("SELECT t.TransferID, t.TransferNumber AS [«·—ﬁ„], Format(t.TransferDate, " & _
        "'yyyy/mm/dd') AS [«· «—ÌŒ], IIf(f.CenterName Is Null, '-', f.CenterName) AS [„‰ „—ﬂ“], IIf(g.CenterName Is Null, '-', " & _
        "g.CenterName) AS [≈·Ï „—ﬂ“], Format(t.AssetValue, '#,##0.00') AS [ﬁÌ„… «·√’·], Format(t.AccumDep, '#,##0.00') AS " & _
        "[„Ã„⁄ «·≈Â·«ﬂ] FROM (AssetTransfers AS t LEFT JOIN [@CostCenters] AS f ON t.FromCenterID = f.CostCenterID) LEFT JOIN " & _
        "[@CostCenters] AS g ON t.ToCenterID = g.CostCenterID WHERE t.AssetID = {ID} ORDER BY t.TransferDate DESC, " & _
        "t.TransferID DESC")
    frm!lstTransfers.RowSource = Replace(rows, "{ID}", CStr(id))
End Sub

Public Sub AssetChangesRevalue(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If IsNull(frm!txtNewValue.Value) Then
        ShowWarning "«ﬂ » «·ﬁÌ„… «·ÃœÌœ… ··√’·."
        Exit Sub
    End If
    If Not AskYesNo("≈⁄«œ…  ﬁÌÌ„ «·√’· ≈·Ï " & Format$(frm!txtNewValue.Value, "#,##0.00") & "ø") Then Exit Sub
    msg = RevalueAsset(ChangesAsset(frm), frm!txtRevalDate.Value, CCur(Nz(frm!txtNewValue.Value, 0)), _
                       CLng(Nz(frm!txtMonthsLeft.Value, 0)), Nz(frm!txtRevalNotes.Value, ""), id)
    If id = 0 Then
        ShowWarning msg
    Else
        frm!txtRevalNotes.Value = Null
        ShowInfo " „  ≈⁄«œ… «· ﬁÌÌ„ ÊﬁÌœÂ«."
    End If
    AssetChangesRefresh frm
End Sub

Public Sub AssetChangesUndoReval(ByVal frm As Access.Form)
    Dim msg As String
    If Not AskYesNo("Õ–› ¬Œ— ≈⁄«œ…  ﬁÌÌ„ ··√’·ø") Then Exit Sub
    msg = DeleteLastRevaluation(ChangesAsset(frm))
    If Len(msg) > 0 Then ShowWarning msg
    AssetChangesRefresh frm
End Sub

Public Sub AssetChangesTransfer(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If Not AskYesNo("‰ﬁ· «·√’· ≈·Ï „—ﬂ“ «· ﬂ·›… «·„Œ «—ø") Then Exit Sub
    msg = TransferAsset(ChangesAsset(frm), frm!txtTransferDate.Value, frm!cboToCenter.Value, Nz(frm!txtTransferNotes.Value, ""), id)
    If id = 0 Then
        ShowWarning msg
    Else
        frm!txtTransferNotes.Value = Null
        ShowInfo " „ ‰ﬁ· «·√’· ÊﬁÌœÂ."
    End If
    AssetChangesRefresh frm
End Sub

Public Sub AssetChangesUndoTransfer(ByVal frm As Access.Form)
    Dim msg As String
    If Not AskYesNo("Õ–› ¬Œ— ‰ﬁ· ··√’· Ê≈—Ã«⁄Â ≈·Ï „—ﬂ“Â «·”«»ﬁø") Then Exit Sub
    msg = DeleteLastTransfer(ChangesAsset(frm))
    If Len(msg) > 0 Then ShowWarning msg
    AssetChangesRefresh frm
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Private Sub CheckAsset(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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

Public Function TestAssets() As Boolean
    Dim passed As Long, failed As Long, report As String, msg As String, ws As DAO.Workspace, inTrans As Boolean
    Dim id As Long, runID As Long, start As Date, n As Long, accum As Currency, gainBefore As Currency
    Dim id2 As Long, revID As Long, trID As Long, centerA As Variant, surplusBefore As Currency, lossBefore As Currency
    Dim vatBefore As Currency
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestAssets  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    start = DateSerial(Year(Date), Month(Date) - 3, 1)                 ' three months ago
    If ClosedThroughDate() >= start Or Nz(DbValue("SELECT COUNT(*) FROM FixedAssets"), 0) > 0 Then
        Debug.Print "[--]  ÊÃœ √’Ê· À«» … √Ê «·› —… „ﬁ›·…: «·«Œ »«— ÌÕ «Ã ﬁ«⁄œ… »·« √’Ê·"
        GoTo Undo
    End If
    CheckAsset MonthlyDepreciation(12000, 0, 60) = 200 And MonthlyDepreciation(1000, 100, 7) = 128.57, _
               "«·ﬁ”ÿ «·‘Â—Ì »«·ﬁ”ÿ «·À«» ", passed, failed, report
    CheckAsset Len(SaveAsset(id, "TEST", 1790, start, 100, 0, 0, 12, start, "OPENING", Null, Null, Null, 0, "")) > 0, _
               "„Ã„⁄ «·≈Â·«ﬂ ·Ì” Õ”«» √’·", passed, failed, report
    id = 0
    msg = SaveAsset(id, "TEST ÃÂ«“", 1720, start, 1200, 0, 100, 10, start, "OPENING", Null, Null, Null, 200, "")
    CheckAsset Len(msg) = 0 And id > 0 And AccountBalance(1720) >= 1200, "√’· „ÊÃÊœ ﬁ»· «·»—‰«„Ã „⁄ ≈Â·«ﬂÂ «·”«»ﬁ " & msg, _
               passed, failed, report
    msg = RecordDepreciationThrough(DateSerial(Year(Date), Month(Date), 0), n)
    accum = AssetDepreciated(id)
    CheckAsset Len(msg) = 0 And n = 3 And accum = 200 + 3 * 110, "≈Â·«ﬂ À·«À… √‘Â— »«· — Ì» " & msg, passed, failed, report
    CheckAsset DepreciationFor(id, DateSerial(Year(Date), Month(Date) + 6, 0)) <= 110, "«·ﬁ”ÿ ·« Ì Ã«Ê“ «·»«ﬁÌ", _
               passed, failed, report
    CheckAsset Len(SaveAsset(id, "TEST ÃÂ«“", 1720, start, 1500, 0, 100, 10, start, "OPENING", Null, Null, Null, 200, "")) > 0, _
               "·«   €Ì—  ﬂ·›… √’· ⁄·ÌÂ ≈Â·«ﬂ", passed, failed, report
    gainBefore = AccountBalance(5650)
    msg = DisposeAsset(id, Date, 0, "NONE", Null, Null)
    CheckAsset Len(msg) = 0 And AccountBalance(5650) = gainBefore + (1200 - accum), _
               "«·«” »⁄«œ »·« À„‰: «·»«ﬁÌ Œ”«—… " & msg, passed, failed, report
    CheckAsset Len(DeleteLastRun()) > 0, "·« ÌıÕ–› ﬁÌœ ≈Â·«ﬂ ›ÌÂ √’· „” »⁄œ »⁄œÂ", passed, failed, report
    msg = UndoDisposal(id) & DeleteLastRun()
    CheckAsset Len(msg) = 0 And AssetDepreciated(id) = 200 + 2 * 110, "≈·€«¡ «·«” »⁄«œ ÊÕ–› ¬Œ— ﬁÌœ ≈Â·«ﬂ " & msg, _
               passed, failed, report
    ' docs/65: declining balance, revaluation, transfer, output VAT on the sale
    CheckAsset FirstMonthDepreciation("DB", 0, 1200, 0, 24) = 100 And FirstMonthDepreciation("DB", 0.5, 1200, 0, 24) = 50, _
               "«·ﬁ”ÿ «·„ ‰«ﬁ’: ÷⁄› ‰”»… «·ﬁ”ÿ «·À«» ° √Ê «·‰”»… «·„ﬂ Ê»…", passed, failed, report
    id2 = 0
    start = DateSerial(Year(NextRunMonth()), Month(NextRunMonth()), 1)          ' last month: the next one to record
    msg = SaveAsset(id2, "TEST „ ‰«ﬁ’", 1720, start, 1200, 0, 0, 24, start, "OPENING", Null, Null, Null, 0, "", "DB", 0)
    CheckAsset Len(msg) = 0 And DepreciationFor(id2, MonthEnd(start)) = 100, "√’· »«·ﬁ”ÿ «·„ ‰«ﬁ’: «·‘Â— «·√Ê· 100 " & msg, _
               passed, failed, report
    msg = RecordDepreciationThrough(MonthEnd(start), n)
    CheckAsset Len(msg) = 0 And DepreciationFor(id2, MonthEnd(DateAdd("m", 1, start))) = 91.67, _
               "«·‘Â— «·À«‰Ì ⁄·Ï «·ﬁÌ„… «·œ› —Ì… 1100 " & msg, passed, failed, report
    CheckAsset Len(SaveAsset(id2, "TEST „ ‰«ﬁ’", 1720, start, 1200, 0, 0, 24, start, "OPENING", Null, Null, Null, 0, "", _
               "SL", 0.3)) > 0, "‰”»… „ ‰«ﬁ’… „⁄ «·ﬁ”ÿ «·À«»  „—›Ê÷…", passed, failed, report
    surplusBefore = AccountBalance(3500)
    msg = RevalueAsset(id2, DateAdd("m", 1, start), 1300, 10, "", revID)
    CheckAsset Len(msg) = 0 And revID > 0 And AssetBookValue(id2) = 1300 And AccountBalance(3500) = surplusBefore - 200, _
               "≈⁄«œ… «· ﬁÌÌ„ »“Ì«œ… ≈·Ï ›«∆÷ ≈⁄«œ… «· ﬁÌÌ„ " & msg, passed, failed, report
    CheckAsset Len(RevalueAsset(id2, start, 1000, 10, "", n)) > 0, "·« ≈⁄«œ…  ﬁÌÌ„ ›Ì ‘Â— ⁄·ÌÂ ≈Â·«ﬂ", passed, failed, report
    lossBefore = AccountBalance(5660)
    msg = RevalueAsset(id2, DateAdd("m", 1, start), 1000, 10, "", revID)
    CheckAsset Len(msg) = 0 And AccountBalance(3500) = surplusBefore And AccountBalance(5660) = lossBefore + 100, _
               "«·‰ﬁ’ ÌŒ’„ «·›«∆÷ À„ Œ”«—… «‰Œ›«÷ " & msg, passed, failed, report
    CheckAsset DepreciationFor(id2, MonthEnd(DateAdd("m", 1, start))) = 83.33, _
               "«·≈Â·«ﬂ »⁄œÂ« „‰ «·ﬁÌ„… «·ÃœÌœ… Ê«·⁄„— «·»«ﬁÌ", passed, failed, report
    centerA = DbValue("SELECT Min(CostCenterID) FROM CostCenters WHERE IsActive <> 0")
    If Not IsNull(centerA) Then
        CurrentDb.Execute "UPDATE FixedAssets SET CostCenterID = Null WHERE AssetID = " & id2, dbFailOnError + dbSeeChanges
        msg = TransferAsset(id2, DateAdd("m", 1, start), centerA, "", trID)
        CheckAsset Len(msg) = 0 And trID > 0 And Nz(AssetField(id2, "CostCenterID"), 0) = centerA And _
                   Nz(DbValue("SELECT Sum(Debit) - Sum(Credit) FROM JournalLines WHERE AccountCode = 1720 AND CostCenterID = " & _
                              centerA), 0) >= 1000, "‰ﬁ· «·√’·: ﬁÌ„ Â ⁄·Ï «·„—ﬂ“ «·ÃœÌœ " & msg, passed, failed, report
        CheckAsset Len(DeleteLastRevaluation(id2)) > 0, "·«  ıÕ–› ≈⁄«œ…  ﬁÌÌ„ ﬁ»· «·‰ﬁ· «·–Ì »⁄œÂ«", passed, failed, report
        msg = DeleteLastTransfer(id2)
        CheckAsset Len(msg) = 0 And IsNull(AssetField(id2, "CostCenterID")), "Õ–› «·‰ﬁ· Ì⁄Ìœ «·√’· ·„—ﬂ“Â " & msg, _
                   passed, failed, report
    End If
    msg = DeleteLastRevaluation(id2) & DeleteLastRevaluation(id2)
    CheckAsset Len(msg) = 0 And AssetBookValue(id2) = 1100 And AccountBalance(3500) = surplusBefore, _
               "Õ–› ≈⁄«œ Ì «· ﬁÌÌ„ „‰ «·√ŒÌ—… " & msg, passed, failed, report
    vatBefore = AccountBalance(2200)
    msg = DisposeAsset(id2, Date, 1000, "CASHBOX", Null, DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE IsActive <> 0"), True)
    CheckAsset Len(msg) = 0 And AccountBalance(2200) = vatBefore - RoundMoney(1000 * Nz(SettingValue("VATRate"), 0)), _
               "»Ì⁄ «·√’· „⁄ ÷—Ì»… «·„Œ—Ã«  " & msg, passed, failed, report
    CheckAsset Nz(DbValue("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0) = 0, _
               "ﬂ· «·ﬁÌÊœ „ Ê«“‰…", passed, failed, report
Undo:
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckAsset False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·√’Ê· «·À«» … ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestAssets"
        TestAssets = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestAssets"
    End If
End Function
