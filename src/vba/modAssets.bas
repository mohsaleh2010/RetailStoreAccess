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
        CanAssets = "لا تملك صلاحية الأصول الثابتة."
    ElseIf Not CanScreenAction(Screen, Action, True) Then
        CanAssets = "لا تملك هذه الصلاحية في شاشة " & IIf(Screen = "frmAssets", "الأصول الثابتة.", "الإهلاك الشهري.")
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
        AssetAccountProblem = "اختر حساب الأصل."
    ElseIf IsNull(DbValue("SELECT AccountCode FROM Accounts WHERE AccountCode = " & CLng(AccountCode) & " AND IsPosting = True " & _
                          "AND IsActive = True AND AccountType = 'ASSET' AND Level2Code = 12 AND AccountCode <> 1790")) Then
        AssetAccountProblem = "حساب الأصل حساب فرعي من الأصول غير المتداولة (12)، غير مجمع الإهلاك."
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
        p = "اختر طريقة الإهلاك."
    ElseIf DepRate < 0 Or DepRate > 1 Then
        p = "نسبة القسط المتناقص بين 0% و100%."
    ElseIf DepMethod = "SL" And DepRate <> 0 Then
        p = "نسبة القسط المتناقص للطريقة المتناقصة فقط."
    ElseIf Not IsDate(PurchaseDate) Or Not IsDate(DepStart) Then
        p = "اكتب تاريخ الشراء وبداية الإهلاك."
    ElseIf DateValue(PurchaseDate) > Date Then
        p = "تاريخ الشراء بعد اليوم."
    ElseIf DateValue(DepStart) < DateSerial(Year(PurchaseDate), Month(PurchaseDate), 1) Then
        p = "بداية الإهلاك قبل شهر الشراء."
    ElseIf Cost <= 0 Then
        p = "التكلفة أكبر من صفر."
    ElseIf InputVAT < 0 Or Salvage < 0 Or OpeningAccum < 0 Then
        p = "الضريبة والقيمة المتبقية والإهلاك السابق لا تكون سالبة."
    ElseIf Salvage + OpeningAccum > Cost Then
        p = "القيمة المتبقية والإهلاك السابق لا يتجاوزان التكلفة."
    ElseIf LifeMonths <= 0 Then
        p = "العمر الإنتاجي شهر على الأقل."
    Else
        Select Case SourceType
            Case "OPENING"
                If InputVAT <> 0 Then p = "الأصل الموجود قبل البرنامج بلا ضريبة مدخلات (سُجِّلت وقت شرائه)."
            Case "BANK"
                If IsNull(BankID) Then
                    p = "اختر البنك الذي دُفع منه."
                ElseIf Nz(DbValue("SELECT COUNT(*) FROM Banks WHERE IsActive = True AND BankID = " & CLng(BankID)), 0) = 0 Then
                    p = "البنك غير نشط."
                End If
            Case "CASHBOX"
                If IsNull(CashBoxID) Then
                    p = "اختر الصندوق الذي دُفع منه."
                ElseIf Not BoxIsActive(CashBoxID) Then
                    p = "الصندوق غير نشط."
                ElseIf CashBoxBalance(CLng(CashBoxID), DateValue(PurchaseDate) + 1) + _
                       IIf(AssetID > 0, Nz(AssetField(AssetID, "IIf(SourceType = 'CASHBOX', Cost + InputVAT, 0)"), 0), 0) _
                       < Cost + InputVAT Then
                    p = "رصيد الصندوق في " & GDate(PurchaseDate) & " أقل من المبلغ."
                End If
            Case "ACCOUNT"
                p = CounterAccountProblem(CounterAccount)               ' modBank
            Case Else
                p = "اختر مصدر الشراء."
        End Select
    End If
    If Len(p) = 0 And SourceType <> "OPENING" And OpeningAccum <> 0 Then
        p = "الإهلاك السابق للأصول الموجودة قبل البرنامج فقط (مصدر الشراء: رصيد افتتاحي)."
    End If
    If Len(p) = 0 Then p = ClosedPeriodProblem(PurchaseDate)
    AssetProblem = p
End Function

Private Function AssetLocked(ByVal AssetID As Long) As String
    ' An asset with monthly depreciation, a revaluation or a transfer, or disposed keeps its data.
    If Nz(AssetField(AssetID, "Status"), "") = "DISPOSED" Then
        AssetLocked = "الأصل مستبعد: ألغِ الاستبعاد أولًا."
    ElseIf Not IsNull(LastAssetEvent(AssetID)) Then
        AssetLocked = "على الأصل إعادة تقييم أو نقل: لا تتغير بياناته المالية ولا يُحذف." & vbCrLf & _
                      "للتصحيح احذفها من شاشة إعادة التقييم والنقل بدءًا من الأخيرة."
    ElseIf Nz(DbValue("SELECT COUNT(*) FROM AssetDepreciations WHERE AssetID = " & AssetID), 0) > 0 Then
        AssetLocked = "على الأصل قيود إهلاك: لا تتغير بياناته المالية ولا يُحذف." & vbCrLf & _
                      "للتصحيح احذف قيود الإهلاك بدءًا من الأخير."
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
    If Len(SaveAsset) = 0 And Len(Trim$(AssetName)) = 0 Then SaveAsset = "اكتب اسم الأصل."
    If Len(SaveAsset) = 0 And AssetID > 0 Then SaveAsset = AssetLocked(AssetID)
    If Len(SaveAsset) = 0 Then
        SaveAsset = AssetProblem(AssetID, AccountCode, PurchaseDate, Cost, InputVAT, Salvage, LifeMonths, DepStart, SourceType, _
                                 BankID, CashBoxID, CounterAccount, OpeningAccum, DepMethod, DepRate)
    End If
    If Len(SaveAsset) > 0 Then Exit Function
    Dim auditBefore As Collection
    If AssetID = 0 Then
        Set rs = CurrentDb.OpenRecordset("FixedAssets", dbOpenDynaset)
        rs.AddNew
        rs!AssetCode = NextNumber("FIXED_ASSET")
        rs!Status = "ACTIVE"
    Else
        Set auditBefore = AuditSnapshot("FixedAssets", "AssetID", AssetID)       ' modAudit
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM FixedAssets WHERE AssetID = " & AssetID, dbOpenDynaset)
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
    CurrentDb.Execute "DELETE FROM FixedAssets WHERE AssetID = " & AssetID, dbFailOnError
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
        DisposeAsset = "الأصل غير موجود أو مستبعد."
    ElseIf Not IsDate(DisposalDate) Then
        DisposeAsset = "اكتب تاريخ البيع أو الاستبعاد."
    ElseIf DateValue(DisposalDate) > Date Then
        DisposeAsset = "التاريخ بعد اليوم."
    ElseIf DateValue(DisposalDate) < DateValue(AssetField(AssetID, "PurchaseDate")) Then
        DisposeAsset = "التاريخ قبل تاريخ الشراء."
    ElseIf Proceeds < 0 Then
        DisposeAsset = "ثمن البيع لا يكون سالبًا."
    ElseIf DisposalTo = "NONE" And Proceeds <> 0 Then
        DisposeAsset = "اختر أين استُلم ثمن البيع (بنك أو صندوق)."
    ElseIf DisposalTo = "BANK" And IsNull(BankID) Then
        DisposeAsset = "اختر البنك."
    ElseIf DisposalTo = "CASHBOX" And IsNull(CashBoxID) Then
        DisposeAsset = "اختر الصندوق."
    ElseIf DisposalTo <> "NONE" And DisposalTo <> "BANK" And DisposalTo <> "CASHBOX" Then
        DisposeAsset = "اختر طريقة الاستبعاد."
    ElseIf DisposalTo <> "NONE" And Proceeds = 0 Then
        DisposeAsset = "اكتب ثمن البيع، أو اختر «استبعاد بدون ثمن»."
    ElseIf WithVAT And Proceeds = 0 Then
        DisposeAsset = "الضريبة على ثمن البيع فقط."
    End If
    If Len(DisposeAsset) = 0 Then
        later = DepreciatedFrom(AssetID, DateValue(DisposalDate))
        If Not IsNull(later) Then DisposeAsset = "للأصل إهلاك مسجل حتى " & GDate(later) & _
                                                 ": احذف قيود الإهلاك من شهر الاستبعاد أولًا."
    End If
    If Len(DisposeAsset) = 0 Then
        later = LastAssetEvent(AssetID)
        If Not IsNull(later) Then
            If DateValue(DisposalDate) < later Then DisposeAsset = "للأصل إعادة تقييم أو نقل في " & GDate(later) & _
                                                                   ": التاريخ لا يكون قبله."
        End If
    End If
    If Len(DisposeAsset) = 0 Then DisposeAsset = ClosedPeriodProblem(DisposalDate)
    If Len(DisposeAsset) > 0 Then Exit Function
    If WithVAT Then vat = RoundMoney(Proceeds * Nz(SettingValue("VATRate"), 0))
    CurrentDb.Execute "UPDATE FixedAssets SET Status = 'DISPOSED', DisposalDate = " & SqlDate(DateValue(DisposalDate)) & _
        ", DisposalProceeds = " & Str$(Proceeds) & ", DisposalVAT = " & Str$(vat) & ", DisposalTo = " & SqlText(DisposalTo) & ", DisposalBankID = " & _
        IIf(DisposalTo = "BANK", CLng(Nz(BankID, 0)), "Null") & ", DisposalCashBoxID = " & _
        IIf(DisposalTo = "CASHBOX", CLng(Nz(CashBoxID, 0)), "Null") & ", DisposalAccumDep = " & Str$(AssetDepreciated(AssetID)) & _
        " WHERE AssetID = " & AssetID, dbFailOnError
    LogAction "FIXED_ASSET_DISPOSE", "FixedAssets", CStr(AssetID), Format$(Proceeds, "0.00")
    DisposeAsset = SyncJournal()
End Function

Public Function UndoDisposal(ByVal AssetID As Long) As String
    UndoDisposal = CanAssets("frmAssets", "EDIT")
    If Len(UndoDisposal) > 0 Then Exit Function
    If Nz(AssetField(AssetID, "Status"), "") <> "DISPOSED" Then
        UndoDisposal = "الأصل غير مستبعد."
        Exit Function
    End If
    UndoDisposal = ClosedPeriodProblem(AssetField(AssetID, "DisposalDate"))
    If Len(UndoDisposal) > 0 Then Exit Function
    CurrentDb.Execute "UPDATE FixedAssets SET Status = 'ACTIVE', DisposalDate = Null, DisposalProceeds = 0, DisposalTo = Null, " & _
                      "DisposalBankID = Null, DisposalCashBoxID = Null, DisposalAccumDep = 0, DisposalVAT = 0 WHERE AssetID = " & AssetID, dbFailOnError
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
        RecordDepreciation = "لا توجد أصول ثابتة."
    ElseIf RunMonth <> nextMonth Then
        RecordDepreciation = "الشهر التالي للإهلاك هو " & Format$(nextMonth, "yyyy/mm") & ": الأشهر تُسجَّل بالترتيب."
    ElseIf RunMonth > Date Then
        RecordDepreciation = "يُسجَّل إهلاك الشهر في آخر يوم منه أو بعده."
    Else
        RecordDepreciation = ClosedPeriodProblem(RunMonth)
    End If
    If Len(RecordDepreciation) > 0 Then Exit Function
    Set db = CurrentDb
    Set rs = db.OpenRecordset("DepreciationRuns", dbOpenDynaset)
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
    Set d = db.OpenRecordset("AssetDepreciations", dbOpenDynaset, dbAppendOnly)
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
    db.Execute "UPDATE DepreciationRuns SET TotalAmount = " & Str$(total) & " WHERE RunID = " & RunID, dbFailOnError
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
    If Months = 0 Then RecordDepreciationThrough = "لا توجد أشهر لم تُسجَّل حتى " & Format$(LastMonth, "yyyy/mm") & "."
End Function

Public Function DeleteLastRun() As String
    Dim last As Variant
    DeleteLastRun = CanAssets("frmDepreciation", "DELETE")
    If Len(DeleteLastRun) > 0 Then Exit Function
    last = DbValue("SELECT TOP 1 RunID FROM DepreciationRuns ORDER BY RunMonth DESC")
    If IsNull(last) Then
        DeleteLastRun = "لا توجد قيود إهلاك."
        Exit Function
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM AssetDepreciations AS d INNER JOIN FixedAssets AS a ON d.AssetID = a.AssetID " & _
                  "WHERE d.RunID = " & last & " AND a.Status = 'DISPOSED'"), 0) > 0 Then
        DeleteLastRun = "في القيد أصل مستبعد بعده: ألغِ الاستبعاد أولًا."
        Exit Function
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM AssetRevaluations AS v, DepreciationRuns AS r WHERE r.RunID = " & last & _
                  " AND v.RevalDate > r.RunMonth AND v.AssetID IN (SELECT AssetID FROM AssetDepreciations WHERE RunID = " & _
                  last & ")"), 0) + _
       Nz(DbValue("SELECT COUNT(*) FROM AssetTransfers AS v, DepreciationRuns AS r WHERE r.RunID = " & last & _
                  " AND v.TransferDate > r.RunMonth AND v.AssetID IN (SELECT AssetID FROM AssetDepreciations WHERE RunID = " & _
                  last & ")"), 0) > 0 Then
        DeleteLastRun = "لأصل في القيد إعادة تقييم أو نقل بعده: احذفها أولًا."
        Exit Function
    End If
    DeleteLastRun = ClosedPeriodProblem(DbValue("SELECT RunMonth FROM DepreciationRuns WHERE RunID = " & last))
    If Len(DeleteLastRun) > 0 Then Exit Function
    Dim auditBefore As Collection
    Set auditBefore = AuditSnapshot("DepreciationRuns", "RunID", last)        ' the record as it was (modAudit)
    CurrentDb.Execute "DELETE FROM DepreciationRuns WHERE RunID = " & last, dbFailOnError      ' lines cascade
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
        p = "اختر أصلًا قائمًا (غير مستبعد)."
    ElseIf Not IsDate(EventDate) Then
        p = "اكتب التاريخ."
    ElseIf DateValue(EventDate) > Date Then
        p = "التاريخ بعد اليوم."
    ElseIf DateValue(EventDate) < DateValue(AssetField(AssetID, "PurchaseDate")) Then
        p = "التاريخ قبل تاريخ الشراء."
    End If
    If Len(p) = 0 Then
        later = DepreciatedFrom(AssetID, DateValue(EventDate))
        If Not IsNull(later) Then p = "للأصل إهلاك مسجل حتى " & GDate(later) & _
                                     ": احذف قيود الإهلاك من شهر هذا التاريخ أولًا."
    End If
    If Len(p) = 0 Then
        later = LastAssetEvent(AssetID)
        If Not IsNull(later) Then
            If DateValue(EventDate) < later Then p = "للأصل إعادة تقييم أو نقل في " & GDate(later) & ": التاريخ لا يكون قبله."
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
            p = "القيمة الجديدة أقل من القيمة المتبقية للأصل."
        ElseIf NewValue = before Then
            p = "القيمة الجديدة تساوي القيمة الدفترية " & Format$(before, "#,##0.00") & "."
        ElseIf MonthsLeft <= 0 Then
            p = "العمر الباقي شهر على الأقل."
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
    Set rs = CurrentDb.OpenRecordset("AssetRevaluations", dbOpenDynaset)
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
    If Len(p) = 0 And Nz(AssetField(AssetID, "Status"), "") <> "ACTIVE" Then p = "الأصل مستبعد: ألغِ الاستبعاد أولًا."
    If Len(p) = 0 Then
        id = DbValue("SELECT TOP 1 RevaluationID FROM AssetRevaluations WHERE AssetID = " & AssetID & _
                     " ORDER BY RevalDate DESC, RevaluationID DESC")
        If IsNull(id) Then p = "لا توجد إعادة تقييم للأصل."
    End If
    If Len(p) = 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT RevalDate FROM AssetRevaluations WHERE RevaluationID = " & id, _
                                         dbOpenSnapshot)
        later = DepreciatedFrom(AssetID, rs!RevalDate)
        If Not IsNull(later) Then
            p = "للأصل إهلاك مسجل حتى " & GDate(later) & ": احذف قيود الإهلاك من شهر إعادة التقييم أولًا."
        ElseIf Nz(DbValue("SELECT COUNT(*) FROM AssetTransfers WHERE AssetID = " & AssetID & " AND TransferDate >= " & _
                          SqlDate(rs!RevalDate)), 0) > 0 Then
            p = "للأصل نقل في يوم إعادة التقييم أو بعده: احذف النقل أولًا."
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
    CurrentDb.Execute "DELETE FROM AssetRevaluations WHERE RevaluationID = " & id, dbFailOnError
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
            p = "الأصل على هذا المركز."
        ElseIf Not IsNull(ToCenter) Then
            If Nz(DbValue("SELECT COUNT(*) FROM CostCenters WHERE IsActive = True AND CostCenterID = " & CLng(ToCenter)), 0) = 0 Then
                p = "مركز التكلفة غير نشط."
            End If
        End If
    End If
    If Len(p) > 0 Then
        TransferAsset = p
        Exit Function
    End If
    ' the depreciation recorded before stays on the centre it was on
    CurrentDb.Execute "UPDATE AssetDepreciations SET DepCenter = " & CLng(Nz(fromCenter, 0)) & " WHERE AssetID = " & AssetID & _
                      " AND DepCenter Is Null", dbFailOnError
    Set rs = CurrentDb.OpenRecordset("AssetTransfers", dbOpenDynaset)
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
                      " WHERE AssetID = " & AssetID, dbFailOnError
    LogAction "ASSET_TRANSFER", "AssetTransfers", CStr(TransferID), CStr(Nz(ToCenter, 0))
    TransferAsset = SyncJournal()
End Function

Public Function DeleteLastTransfer(ByVal AssetID As Long) As String
    ' The asset goes back to the centre it came from.
    Dim id As Variant, rs As DAO.Recordset, p As String, later As Variant, fromCenter As Variant, auditBefore As Collection
    p = CanAssets("frmAssetChanges", "DELETE")
    If Len(p) = 0 And Nz(AssetField(AssetID, "Status"), "") <> "ACTIVE" Then p = "الأصل مستبعد: ألغِ الاستبعاد أولًا."
    If Len(p) = 0 Then
        id = DbValue("SELECT TOP 1 TransferID FROM AssetTransfers WHERE AssetID = " & AssetID & _
                     " ORDER BY TransferDate DESC, TransferID DESC")
        If IsNull(id) Then p = "لا يوجد نقل للأصل."
    End If
    If Len(p) = 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT TransferDate, FromCenterID FROM AssetTransfers WHERE TransferID = " & _
                                         id, dbOpenSnapshot)
        fromCenter = rs!FromCenterID
        later = DepreciatedFrom(AssetID, rs!TransferDate)
        If Not IsNull(later) Then
            p = "للأصل إهلاك مسجل حتى " & GDate(later) & ": احذف قيود الإهلاك من شهر النقل أولًا."
        ElseIf Nz(DbValue("SELECT COUNT(*) FROM AssetRevaluations WHERE AssetID = " & AssetID & " AND RevalDate > " & _
                          SqlDate(rs!TransferDate)), 0) > 0 Then
            p = "للأصل إعادة تقييم بعد النقل: احذفها أولًا."
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
    CurrentDb.Execute "DELETE FROM AssetTransfers WHERE TransferID = " & id, dbFailOnError
    CurrentDb.Execute "UPDATE FixedAssets SET CostCenterID = " & IIf(IsNull(fromCenter), "Null", CLng(Nz(fromCenter, 0))) & _
                      " WHERE AssetID = " & AssetID, dbFailOnError
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
        frm!lblAssetInfo.Caption = Tr("أصل جديد: قسط الشهر الأول " & Format$(FirstMonthDepreciation(Nz(frm!cboDepMethod.Value, "SL"), _
            Nz(frm!txtDepRate.Value, 0), Nz(frm!txtCost.Value, 0), Nz(frm!txtSalvage.Value, 0), Nz(frm!txtLife.Value, 0)), _
            "#,##0.00"))
        frm!lblAssetInfo.ForeColor = CLR_PRIMARY
        Exit Sub
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM FixedAssetsQuery WHERE AssetID = " & id, dbOpenSnapshot)
    If Not rs.EOF Then
        info = rs!AssetCode & "   مجمع الإهلاك " & Format$(rs!AccumDep, "#,##0.00") & "   القيمة الدفترية " & _
               Format$(rs!BookValue, "#,##0.00") & "   قسط الشهر " & Format$(rs!MonthlyDep, "#,##0.00") & "   (" & _
               rs!DepMonths & " شهرًا مسجلة، " & rs!MethodName & ")"
        If rs!RevalTotal <> 0 Then info = info & "   إعادة التقييم " & Format$(rs!RevalTotal, "#,##0.00")
        If rs!Status = "DISPOSED" Then
            info = info & vbCrLf & "مستبعد في " & GDate(rs!DisposalDate) & _
                   IIf(Nz(rs!DisposalProceeds, 0) > 0, " بثمن " & Format$(rs!DisposalProceeds, "#,##0.00"), "")
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
                      Nz(frm!cboCenter.Value, 0)) & " WHERE AssetID = " & id, dbFailOnError
    SyncJournal
    ShowInfo IIf(isNew, "تم حفظ الأصل وقيد شرائه.", "تم حفظ التعديل.")
    AssetShow frm, id
End Sub

Public Sub AssetDelete(ByVal frm As Access.Form)
    Dim msg As String
    If ShownAsset(frm) = 0 Then Exit Sub
    If Not AskYesNo("حذف الأصل وقيد شرائه؟") Then Exit Sub
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
        ShowWarning "اختر الأصل من القائمة."
        Exit Sub
    End If
    If Not AskYesNo("بيع أو استبعاد الأصل في " & GDate(Nz(frm!txtDisposalDate.Value, Date)) & "؟") Then Exit Sub
    msg = DisposeAsset(ShownAsset(frm), frm!txtDisposalDate.Value, CCur(Nz(frm!txtProceeds.Value, 0)), _
                       Nz(frm!cboDisposalTo.Value, ""), frm!cboDisposalBank.Value, frm!cboDisposalBox.Value, _
                       Nz(frm!chkDisposalVAT.Value, False))
    If AssetField(ShownAsset(frm), "Status") <> "DISPOSED" Then ShowWarning msg
    AssetShow frm, ShownAsset(frm)
End Sub

Public Sub AssetUndoDisposal(ByVal frm As Access.Form)
    Dim msg As String
    If ShownAsset(frm) = 0 Then Exit Sub
    If Not AskYesNo("إلغاء بيع / استبعاد الأصل؟") Then Exit Sub
    msg = UndoDisposal(ShownAsset(frm))
    If AssetField(ShownAsset(frm), "Status") <> "ACTIVE" Then ShowWarning msg
    AssetShow frm, ShownAsset(frm)
End Sub

Public Sub AssetOpenChanges(ByVal frm As Access.Form)
    If ShownAsset(frm) = 0 Then
        ShowWarning "اختر الأصل من القائمة."
        Exit Sub
    End If
    OpenScreen "frmAssetChanges", 0, ShownAsset(frm)
End Sub

Public Sub PrintAssets(ByVal frm As Access.Form)
    LogAction "REPORT", "FIXED_ASSETS"
    OpenReportOrQuery "rptFixedAssets", "FixedAssetsQuery", "", "في " & GDate(Date)
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
        frm!lblNext.Caption = Tr("لا توجد أصول ثابتة.")
        frm!lstPreview.RowSource = Tr("")
    Else
        ' what the next month will record
        rows = """الأصل"";""الإهلاك"""
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
        frm!lblNext.Caption = Tr("الشهر التالي للتسجيل: " & Format$(nextMonth, "yyyy/mm") & "   الإجمالي " & Format$(total, "#,##0.00") & _
                              IIf(IsNull(LastRunMonth()), "", "   (آخر شهر مسجل " & Format$(Nz(LastRunMonth(), 0), "yyyy/mm") & ")"))
    End If
    frm!lstRuns.Requery
End Sub

Public Sub DepreciationRun(ByVal frm As Access.Form)
    Dim msg As String, n As Long
    If Not IsDate(frm!txtThrough.Value) Then
        ShowWarning "اختر آخر شهر يُسجَّل إهلاكه."
        Exit Sub
    End If
    If Not AskYesNo("تسجيل الإهلاك لكل شهر لم يُسجَّل حتى " & Format$(frm!txtThrough.Value, "yyyy/mm") & "؟") Then Exit Sub
    msg = RecordDepreciationThrough(DateValue(frm!txtThrough.Value), n)
    If n = 0 Then
        ShowWarning msg
    Else
        If Len(msg) > 0 Then ShowWarning msg
        ShowInfo "تم تسجيل إهلاك " & n & " شهر."
    End If
    DepreciationRefresh frm
End Sub

Public Sub DepreciationUndo(ByVal frm As Access.Form)
    Dim msg As String, before As Variant
    before = LastRunMonth()
    If IsNull(before) Then Exit Sub
    If Not AskYesNo("حذف قيد إهلاك شهر " & Format$(before, "yyyy/mm") & "؟") Then Exit Sub
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
        info = "اختر الأصل من شاشة الأصول الثابتة."
    Else
        center = DbValue(Tr("SELECT c.CenterName FROM [@CostCenters] AS c INNER JOIN FixedAssets AS a ON c.CostCenterID = " & _
                            "a.CostCenterID WHERE a.AssetID = " & id))
        info = rs!AssetCode & "  " & rs!AssetName & "   القيمة الدفترية " & Format$(rs!BookValue, "#,##0.00") & _
               "   مركز التكلفة: " & Nz(center, "-") & "   (" & rs!StatusName & ")"
    End If
    rs.Close
    frm!lblAsset.Caption = Tr(info)
    frm!txtNewValue.Value = Null
    frm!txtMonthsLeft.Value = IIf(id > 0, RevalMonthsLeft(id, Nz(frm!txtRevalDate.Value, Date)), Null)
    rows = Tr("SELECT r.RevaluationID, r.RevalNumber AS [الرقم], Format(r.RevalDate, 'yyyy/mm/dd') " & _
        "AS [التاريخ], Format(r.BookValueBefore, '#,##0.00') AS [القيمة الدفترية قبلها], Format(r.NewValue, '#,##0.00') AS " & _
        "[القيمة بعد التقييم], Format(r.ValueChange, '#,##0.00') AS [الفرق], Format(r.SurplusPart, '#,##0.00') AS [فائض التقييم], " & _
        "Format(r.LossPart, '#,##0.00') AS [خسارة الانخفاض], r.MonthsLeft AS [العمر الباقي] FROM AssetRevaluations AS r " & _
        "WHERE r.AssetID = {ID} ORDER BY r.RevalDate DESC, r.RevaluationID DESC")
    frm!lstRevals.RowSource = Replace(rows, "{ID}", CStr(id))
    rows = Tr("SELECT t.TransferID, t.TransferNumber AS [الرقم], Format(t.TransferDate, " & _
        "'yyyy/mm/dd') AS [التاريخ], IIf(f.CenterName Is Null, '-', f.CenterName) AS [من مركز], IIf(g.CenterName Is Null, '-', " & _
        "g.CenterName) AS [إلى مركز], Format(t.AssetValue, '#,##0.00') AS [قيمة الأصل], Format(t.AccumDep, '#,##0.00') AS " & _
        "[مجمع الإهلاك] FROM (AssetTransfers AS t LEFT JOIN [@CostCenters] AS f ON t.FromCenterID = f.CostCenterID) LEFT JOIN " & _
        "[@CostCenters] AS g ON t.ToCenterID = g.CostCenterID WHERE t.AssetID = {ID} ORDER BY t.TransferDate DESC, " & _
        "t.TransferID DESC")
    frm!lstTransfers.RowSource = Replace(rows, "{ID}", CStr(id))
End Sub

Public Sub AssetChangesRevalue(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If IsNull(frm!txtNewValue.Value) Then
        ShowWarning "اكتب القيمة الجديدة للأصل."
        Exit Sub
    End If
    If Not AskYesNo("إعادة تقييم الأصل إلى " & Format$(frm!txtNewValue.Value, "#,##0.00") & "؟") Then Exit Sub
    msg = RevalueAsset(ChangesAsset(frm), frm!txtRevalDate.Value, CCur(Nz(frm!txtNewValue.Value, 0)), _
                       CLng(Nz(frm!txtMonthsLeft.Value, 0)), Nz(frm!txtRevalNotes.Value, ""), id)
    If id = 0 Then
        ShowWarning msg
    Else
        frm!txtRevalNotes.Value = Null
        ShowInfo "تمت إعادة التقييم وقيدها."
    End If
    AssetChangesRefresh frm
End Sub

Public Sub AssetChangesUndoReval(ByVal frm As Access.Form)
    Dim msg As String
    If Not AskYesNo("حذف آخر إعادة تقييم للأصل؟") Then Exit Sub
    msg = DeleteLastRevaluation(ChangesAsset(frm))
    If Len(msg) > 0 Then ShowWarning msg
    AssetChangesRefresh frm
End Sub

Public Sub AssetChangesTransfer(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If Not AskYesNo("نقل الأصل إلى مركز التكلفة المختار؟") Then Exit Sub
    msg = TransferAsset(ChangesAsset(frm), frm!txtTransferDate.Value, frm!cboToCenter.Value, Nz(frm!txtTransferNotes.Value, ""), id)
    If id = 0 Then
        ShowWarning msg
    Else
        frm!txtTransferNotes.Value = Null
        ShowInfo "تم نقل الأصل وقيده."
    End If
    AssetChangesRefresh frm
End Sub

Public Sub AssetChangesUndoTransfer(ByVal frm As Access.Form)
    Dim msg As String
    If Not AskYesNo("حذف آخر نقل للأصل وإرجاعه إلى مركزه السابق؟") Then Exit Sub
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
        Debug.Print "[--] توجد أصول ثابتة أو الفترة مقفلة: الاختبار يحتاج قاعدة بلا أصول"
        GoTo Undo
    End If
    CheckAsset MonthlyDepreciation(12000, 0, 60) = 200 And MonthlyDepreciation(1000, 100, 7) = 128.57, _
               "القسط الشهري بالقسط الثابت", passed, failed, report
    CheckAsset Len(SaveAsset(id, "TEST", 1790, start, 100, 0, 0, 12, start, "OPENING", Null, Null, Null, 0, "")) > 0, _
               "مجمع الإهلاك ليس حساب أصل", passed, failed, report
    id = 0
    msg = SaveAsset(id, "TEST جهاز", 1720, start, 1200, 0, 100, 10, start, "OPENING", Null, Null, Null, 200, "")
    CheckAsset Len(msg) = 0 And id > 0 And AccountBalance(1720) >= 1200, "أصل موجود قبل البرنامج مع إهلاكه السابق " & msg, _
               passed, failed, report
    msg = RecordDepreciationThrough(DateSerial(Year(Date), Month(Date), 0), n)
    accum = AssetDepreciated(id)
    CheckAsset Len(msg) = 0 And n = 3 And accum = 200 + 3 * 110, "إهلاك ثلاثة أشهر بالترتيب " & msg, passed, failed, report
    CheckAsset DepreciationFor(id, DateSerial(Year(Date), Month(Date) + 6, 0)) <= 110, "القسط لا يتجاوز الباقي", _
               passed, failed, report
    CheckAsset Len(SaveAsset(id, "TEST جهاز", 1720, start, 1500, 0, 100, 10, start, "OPENING", Null, Null, Null, 200, "")) > 0, _
               "لا تتغير تكلفة أصل عليه إهلاك", passed, failed, report
    gainBefore = AccountBalance(5650)
    msg = DisposeAsset(id, Date, 0, "NONE", Null, Null)
    CheckAsset Len(msg) = 0 And AccountBalance(5650) = gainBefore + (1200 - accum), _
               "الاستبعاد بلا ثمن: الباقي خسارة " & msg, passed, failed, report
    CheckAsset Len(DeleteLastRun()) > 0, "لا يُحذف قيد إهلاك فيه أصل مستبعد بعده", passed, failed, report
    msg = UndoDisposal(id) & DeleteLastRun()
    CheckAsset Len(msg) = 0 And AssetDepreciated(id) = 200 + 2 * 110, "إلغاء الاستبعاد وحذف آخر قيد إهلاك " & msg, _
               passed, failed, report
    ' docs/65: declining balance, revaluation, transfer, output VAT on the sale
    CheckAsset FirstMonthDepreciation("DB", 0, 1200, 0, 24) = 100 And FirstMonthDepreciation("DB", 0.5, 1200, 0, 24) = 50, _
               "القسط المتناقص: ضعف نسبة القسط الثابت، أو النسبة المكتوبة", passed, failed, report
    id2 = 0
    start = DateSerial(Year(NextRunMonth()), Month(NextRunMonth()), 1)          ' last month: the next one to record
    msg = SaveAsset(id2, "TEST متناقص", 1720, start, 1200, 0, 0, 24, start, "OPENING", Null, Null, Null, 0, "", "DB", 0)
    CheckAsset Len(msg) = 0 And DepreciationFor(id2, MonthEnd(start)) = 100, "أصل بالقسط المتناقص: الشهر الأول 100 " & msg, _
               passed, failed, report
    msg = RecordDepreciationThrough(MonthEnd(start), n)
    CheckAsset Len(msg) = 0 And DepreciationFor(id2, MonthEnd(DateAdd("m", 1, start))) = 91.67, _
               "الشهر الثاني على القيمة الدفترية 1100 " & msg, passed, failed, report
    CheckAsset Len(SaveAsset(id2, "TEST متناقص", 1720, start, 1200, 0, 0, 24, start, "OPENING", Null, Null, Null, 0, "", _
               "SL", 0.3)) > 0, "نسبة متناقصة مع القسط الثابت مرفوضة", passed, failed, report
    surplusBefore = AccountBalance(3500)
    msg = RevalueAsset(id2, DateAdd("m", 1, start), 1300, 10, "", revID)
    CheckAsset Len(msg) = 0 And revID > 0 And AssetBookValue(id2) = 1300 And AccountBalance(3500) = surplusBefore - 200, _
               "إعادة التقييم بزيادة إلى فائض إعادة التقييم " & msg, passed, failed, report
    CheckAsset Len(RevalueAsset(id2, start, 1000, 10, "", n)) > 0, "لا إعادة تقييم في شهر عليه إهلاك", passed, failed, report
    lossBefore = AccountBalance(5660)
    msg = RevalueAsset(id2, DateAdd("m", 1, start), 1000, 10, "", revID)
    CheckAsset Len(msg) = 0 And AccountBalance(3500) = surplusBefore And AccountBalance(5660) = lossBefore + 100, _
               "النقص يخصم الفائض ثم خسارة انخفاض " & msg, passed, failed, report
    CheckAsset DepreciationFor(id2, MonthEnd(DateAdd("m", 1, start))) = 83.33, _
               "الإهلاك بعدها من القيمة الجديدة والعمر الباقي", passed, failed, report
    centerA = DbValue("SELECT Min(CostCenterID) FROM CostCenters WHERE IsActive = True")
    If Not IsNull(centerA) Then
        CurrentDb.Execute "UPDATE FixedAssets SET CostCenterID = Null WHERE AssetID = " & id2, dbFailOnError
        msg = TransferAsset(id2, DateAdd("m", 1, start), centerA, "", trID)
        CheckAsset Len(msg) = 0 And trID > 0 And Nz(AssetField(id2, "CostCenterID"), 0) = centerA And _
                   Nz(DbValue("SELECT Sum(Debit) - Sum(Credit) FROM JournalLines WHERE AccountCode = 1720 AND CostCenterID = " & _
                              centerA), 0) >= 1000, "نقل الأصل: قيمته على المركز الجديد " & msg, passed, failed, report
        CheckAsset Len(DeleteLastRevaluation(id2)) > 0, "لا تُحذف إعادة تقييم قبل النقل الذي بعدها", passed, failed, report
        msg = DeleteLastTransfer(id2)
        CheckAsset Len(msg) = 0 And IsNull(AssetField(id2, "CostCenterID")), "حذف النقل يعيد الأصل لمركزه " & msg, _
                   passed, failed, report
    End If
    msg = DeleteLastRevaluation(id2) & DeleteLastRevaluation(id2)
    CheckAsset Len(msg) = 0 And AssetBookValue(id2) = 1100 And AccountBalance(3500) = surplusBefore, _
               "حذف إعادتي التقييم من الأخيرة " & msg, passed, failed, report
    vatBefore = AccountBalance(2200)
    msg = DisposeAsset(id2, Date, 1000, "CASHBOX", Null, DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE IsActive = True"), True)
    CheckAsset Len(msg) = 0 And AccountBalance(2200) = vatBefore - RoundMoney(1000 * Nz(SettingValue("VATRate"), 0)), _
               "بيع الأصل مع ضريبة المخرجات " & msg, passed, failed, report
    CheckAsset Nz(DbValue("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0) = 0, _
               "كل القيود متوازنة", passed, failed, report
Undo:
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckAsset False, "خطأ: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات الأصول الثابتة ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestAssets"
        TestAssets = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestAssets"
    End If
End Function
