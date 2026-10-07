Attribute VB_Name = "modAssets"
'==============================================================================
' modAssets  -  Retail Store Management System
'
' Fixed assets (FixedAssets, screen frmAssets, permission FIXED_ASSETS):
'   bought  the asset account (a sub-account of 12: furniture, devices, cars...) and
'           the input VAT against a bank, a cash box or another account (journal
'           source ASSET); an asset owned before the program comes from the
'           opening balances (3900) with its depreciation until then (1790).
'   sold / scrapped  the accumulated depreciation and the price against the cost,
'           gain 4500 / loss 5650 (journal source ASSET_DISPOSAL).
' Monthly depreciation (DepreciationRuns + AssetDepreciations, screen frmDepreciation,
' journal source DEPRECIATION): straight line, (cost - salvage) / months, from the
' month of DepStartDate, never below the salvage value, not in the month of the
' disposal; 5600 debit / 1790 credit, a line per asset. Months are recorded in order.
' Same amounts as tools/sim.py (Store.depreciate).
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
                             ByVal OpeningAccum As Currency) As String
    Dim p As String
    p = AssetAccountProblem(AccountCode)
    If Len(p) > 0 Then
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
    ' An asset with monthly depreciation or disposed keeps its data.
    If Nz(AssetField(AssetID, "Status"), "") = "DISPOSED" Then
        AssetLocked = "الأصل مستبعد: ألغِ الاستبعاد أولًا."
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
                          ByVal CounterAccount As Variant, ByVal OpeningAccum As Currency, ByVal Notes As String) As String
    ' A new asset (AssetID = 0) or the changes of one without depreciation.
    Dim rs As DAO.Recordset
    SaveAsset = CanAssets("frmAssets", IIf(AssetID = 0, "ADD", "EDIT"))
    If Len(SaveAsset) = 0 And Len(Trim$(AssetName)) = 0 Then SaveAsset = "اكتب اسم الأصل."
    If Len(SaveAsset) = 0 And AssetID > 0 Then SaveAsset = AssetLocked(AssetID)
    If Len(SaveAsset) = 0 Then
        SaveAsset = AssetProblem(AssetID, AccountCode, PurchaseDate, Cost, InputVAT, Salvage, LifeMonths, DepStart, SourceType, _
                                 BankID, CashBoxID, CounterAccount, OpeningAccum)
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
                             ByVal DisposalTo As String, ByVal BankID As Variant, ByVal CashBoxID As Variant) As String
    ' Sold or scrapped: no depreciation from the month of DisposalDate.
    Dim later As Variant
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
    End If
    If Len(DisposeAsset) = 0 Then
        later = DbValue("SELECT Max(r.RunMonth) FROM AssetDepreciations AS d INNER JOIN DepreciationRuns AS r ON d.RunID = " & _
                        "r.RunID WHERE d.AssetID = " & AssetID & " AND r.RunMonth >= " & _
                        SqlDate(DateSerial(Year(DisposalDate), Month(DisposalDate), 1)))
        If Not IsNull(later) Then DisposeAsset = "للأصل إهلاك مسجل حتى " & GDate(later) & _
                                                 ": احذف قيود الإهلاك من شهر الاستبعاد أولًا."
    End If
    If Len(DisposeAsset) = 0 Then DisposeAsset = ClosedPeriodProblem(DisposalDate)
    If Len(DisposeAsset) > 0 Then Exit Function
    CurrentDb.Execute "UPDATE FixedAssets SET Status = 'DISPOSED', DisposalDate = " & SqlDate(DateValue(DisposalDate)) & _
        ", DisposalProceeds = " & Str$(Proceeds) & ", DisposalTo = " & SqlText(DisposalTo) & ", DisposalBankID = " & _
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
                      "DisposalBankID = Null, DisposalCashBoxID = Null, DisposalAccumDep = 0 WHERE AssetID = " & AssetID, dbFailOnError
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
    Dim rs As DAO.Recordset, remaining As Currency, monthly As Currency
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM FixedAssets WHERE AssetID = " & AssetID, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Exit Function
    End If
    If DateValue(rs!DepStartDate) <= RunMonth And DateValue(rs!PurchaseDate) <= RunMonth And _
       (rs!Status = "ACTIVE" Or Nz(rs!DisposalDate, 0) > RunMonth) Then
        monthly = MonthlyDepreciation(rs!Cost, rs!SalvageValue, rs!UsefulLifeMonths)
        remaining = rs!Cost - rs!SalvageValue - AssetDepreciated(AssetID)
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
    Set rs = db.OpenRecordset("SELECT AssetID FROM FixedAssets ORDER BY AssetID", dbOpenSnapshot)
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
    DeleteLastRun = ClosedPeriodProblem(DbValue("SELECT RunMonth FROM DepreciationRuns WHERE RunID = " & last))
    If Len(DeleteLastRun) > 0 Then Exit Function
    Dim auditBefore As Collection
    Set auditBefore = AuditSnapshot("DepreciationRuns", "RunID", last)        ' the record as it was (modAudit)
    CurrentDb.Execute "DELETE FROM DepreciationRuns WHERE RunID = " & last, dbFailOnError      ' lines cascade
    AuditDeleted "DEPRECIATION_DELETE", "DepreciationRuns", last, auditBefore
    DeleteLastRun = SyncJournal()
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
        frm!lblAssetInfo.Caption = "أصل جديد: القسط الشهري " & Format$(MonthlyDepreciation(Nz(frm!txtCost.Value, 0), _
            Nz(frm!txtSalvage.Value, 0), Nz(frm!txtLife.Value, 0)), "#,##0.00")
        frm!lblAssetInfo.ForeColor = CLR_PRIMARY
        Exit Sub
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM FixedAssetsQuery WHERE AssetID = " & id, dbOpenSnapshot)
    If Not rs.EOF Then
        info = rs!AssetCode & "   مجمع الإهلاك " & Format$(rs!AccumDep, "#,##0.00") & "   القيمة الدفترية " & _
               Format$(rs!BookValue, "#,##0.00") & "   القسط الشهري " & Format$(rs!MonthlyDep, "#,##0.00") & "   (" & _
               rs!DepMonths & " شهرًا مسجلة)"
        If rs!Status = "DISPOSED" Then
            info = info & vbCrLf & "مستبعد في " & GDate(rs!DisposalDate) & _
                   IIf(Nz(rs!DisposalProceeds, 0) > 0, " بثمن " & Format$(rs!DisposalProceeds, "#,##0.00"), "")
        End If
        frm!lblAssetInfo.ForeColor = IIf(rs!Status = "DISPOSED", CLR_MUTED, CLR_PRIMARY)
    End If
    rs.Close
    frm!lblAssetInfo.Caption = info
End Sub

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
                    frm!cboBox.Value, frm!cboCounter.Value, CCur(Nz(frm!txtOpeningAccum.Value, 0)), Nz(frm!txtNotes.Value, ""))
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
                       Nz(frm!cboDisposalTo.Value, ""), frm!cboDisposalBank.Value, frm!cboDisposalBox.Value)
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
        frm!lblNext.Caption = "لا توجد أصول ثابتة."
        frm!lstPreview.RowSource = ""
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
        frm!lstPreview.RowSource = rows
        frm!lblNext.Caption = "الشهر التالي للتسجيل: " & Format$(nextMonth, "yyyy/mm") & "   الإجمالي " & Format$(total, "#,##0.00") & _
                              IIf(IsNull(LastRunMonth()), "", "   (آخر شهر مسجل " & Format$(Nz(LastRunMonth(), 0), "yyyy/mm") & ")")
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
