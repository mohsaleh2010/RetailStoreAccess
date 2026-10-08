Attribute VB_Name = "modCurrency"
'==============================================================================
' modCurrency  -  Retail Store Management System
'
' Several currencies. The program currency is Settings.CurrencyCode (SAR): every
' amount field of every document is kept in it, so the journal, VAT, balances and
' reports never change. A document entered in another currency keeps:
'   CurrencyCode    the currency (required; SAR = the program currency, rate 1)
'   ExchangeRate    the value of ONE unit of it in SAR (proposed from CurrencyRates,
'                   the last rate on or before the document date; can be changed)
'   ForeignAmount   its total in that currency
' Documents: purchase invoices and returns, supplier and customer payments, expenses,
' manual entries (lines in the currency, converted line by line). Their journal
' entries are posted in SAR (the equivalent) and carry the currency, the rate and the
' foreign amount (StampJournalCurrencies, called by SyncJournal).
' Screens: frmCurrencies and frmCurrencyRates (data screens); the document screens have
' cboCurrency + txtRate (CurrencyPicked / CurrencyChoice).
'==============================================================================
Option Compare Database
Option Explicit

Public Const CURRENCY_ROWS As String = "SELECT CurrencyCode, CurrencyCode & '  ' & CurrencyName AS Currency FROM Currencies " & _
                                       "WHERE IsActive = True ORDER BY SortOrder, CurrencyCode"

'------------------------------------------------------------------------------
' Currencies and rates
'------------------------------------------------------------------------------
Public Function BaseCurrency() As String
    BaseCurrency = Nz(SettingValue("CurrencyCode"), "SAR")
    If Len(BaseCurrency) = 0 Then BaseCurrency = "SAR"
End Function

Public Function IsBaseCurrency(ByVal Code As Variant) As Boolean
    If IsNull(Code) Then
        IsBaseCurrency = True
    ElseIf Len(Trim$(Code)) = 0 Then
        IsBaseCurrency = True
    Else
        IsBaseCurrency = (UCase$(Trim$(Code)) = UCase$(BaseCurrency()))
    End If
End Function

Public Function RateOn(ByVal Code As Variant, ByVal When As Variant) As Currency
    ' The rate of the currency on a day: the last one on or before it (1 for SAR, 0 when none).
    Dim d As Date
    If IsBaseCurrency(Code) Then
        RateOn = 1
        Exit Function
    End If
    If IsDate(When) Then
        d = DateValue(When)
    Else
        d = Date
    End If
    RateOn = Nz(DbValue("SELECT TOP 1 Rate FROM CurrencyRates WHERE CurrencyCode = " & SqlText(Code) & _
                        " AND RateDate <= " & SqlDate(d) & " ORDER BY RateDate DESC"), 0)
End Function

Public Function CurrencyProblem(ByVal Code As Variant, ByVal Rate As Variant) As String
    ' "" when the currency and its rate can be used in a document.
    If IsNull(Code) Then
        CurrencyProblem = "«Œ — «·⁄„·…."
    ElseIf Len(Trim$(Code)) = 0 Then
        CurrencyProblem = "«Œ — «·⁄„·…."
    ElseIf IsNull(DbValue("SELECT CurrencyCode FROM Currencies WHERE IsActive = True AND CurrencyCode = " & SqlText(Code))) Then
        CurrencyProblem = "«·⁄„·… " & Code & " €Ì— „ÊÃÊœ… √Ê €Ì— ‰‘ÿ…."
    ElseIf IsBaseCurrency(Code) Then
        If Nz(Rate, 1) <> 1 Then CurrencyProblem = "„⁄«„· ⁄„·… «·»—‰«„Ã " & BaseCurrency() & " Ì”«ÊÌ 1."
    ElseIf Not IsNumeric(Nz(Rate, "")) Then
        CurrencyProblem = "«ﬂ » „⁄«„· «· ÕÊÌ· ··⁄„·… " & Code & " (ﬁÌ„… ÊÕœ… Ê«Õœ… »‹ " & BaseCurrency() & ")."
    ElseIf CDbl(Rate) <= 0 Then
        CurrencyProblem = "«ﬂ » „⁄«„· «· ÕÊÌ· ··⁄„·… " & Code & " (ﬁÌ„… ÊÕœ… Ê«Õœ… »‹ " & BaseCurrency() & ")."
    End If
End Function

Public Function ToBase(ByVal Amount As Currency, ByVal Rate As Double) As Currency
    ' An amount in the document currency -> the program currency (2 decimals).
    ToBase = RoundMoney(CDbl(Amount) * Rate)
End Function

Public Function FromBase(ByVal Amount As Currency, ByVal Rate As Double) As Currency
    If Rate <= 0 Then
        FromBase = Amount
    Else
        FromBase = RoundMoney(CDbl(Amount) / Rate)
    End If
End Function

Public Function SaveCurrencyRate(ByVal Code As String, ByVal RateDate As Variant, ByVal Rate As Double) As String
    ' Adds or replaces the rate of a currency on a day.
    Dim id As Variant
    If IsBaseCurrency(Code) Then
        SaveCurrencyRate = "⁄„·… «·»—‰«„Ã ·«  Õ «Ã ”⁄—«."
        Exit Function
    End If
    If Not IsDate(RateDate) Or Rate <= 0 Then
        SaveCurrencyRate = "«ﬂ » «· «—ÌŒ Ê«·„⁄«„· (√ﬂ»— „‰ ’›—)."
        Exit Function
    End If
    id = DbValue("SELECT CurrencyRateID FROM CurrencyRates WHERE CurrencyCode = " & SqlText(Code) & _
                 " AND RateDate = " & SqlDate(DateValue(RateDate)))
    If IsNull(id) Then
        CurrentDb.Execute "INSERT INTO CurrencyRates (CurrencyCode, RateDate, Rate) VALUES (" & SqlText(Code) & ", " & _
                          SqlDate(DateValue(RateDate)) & ", " & Str$(Rate) & ")", dbFailOnError
    Else
        CurrentDb.Execute "UPDATE CurrencyRates SET Rate = " & Str$(Rate) & " WHERE CurrencyRateID = " & id, dbFailOnError
    End If
    LogAction "CURRENCY_RATE", "CurrencyRates", Code, GDate(RateDate) & " = " & Rate
End Function

'------------------------------------------------------------------------------
' Document screens: cboCurrency + txtRate
'------------------------------------------------------------------------------
Public Sub CurrencyReset(ByVal frm As Access.Form, Optional ByVal Code As Variant)
    ' The currency of a new document: the given one (a supplier's), else SAR.
    If IsMissing(Code) Then Code = Null
    If IsNull(Code) Then Code = BaseCurrency()
    frm!cboCurrency.Value = Code
    CurrencyPicked frm
End Sub

Public Sub CurrencyPicked(ByVal frm As Access.Form, Optional ByVal When As Variant)
    ' A currency was chosen: propose its rate on the document day; SAR keeps 1 and locks it.
    Dim isBase As Boolean
    If IsMissing(When) Then When = Date
    If IsNull(frm!cboCurrency.Value) Then frm!cboCurrency.Value = BaseCurrency()
    isBase = IsBaseCurrency(frm!cboCurrency.Value)
    If isBase Then
        frm!txtRate.Value = 1
    Else
        frm!txtRate.Value = IIf(RateOn(frm!cboCurrency.Value, When) = 0, Null, RateOn(frm!cboCurrency.Value, When))
    End If
    frm!txtRate.Locked = isBase
    frm!txtRate.BackColor = IIf(isBase, CLR_LOCKED, CLR_SURFACE)
End Sub

Public Function CurrencyChoice(ByVal frm As Access.Form, ByRef Code As String, ByRef Rate As Double) As String
    ' The currency and rate of the screen, checked ("" = fine).
    CurrencyChoice = CurrencyProblem(frm!cboCurrency.Value, frm!txtRate.Value)
    If Len(CurrencyChoice) > 0 Then Exit Function
    Code = frm!cboCurrency.Value
    If IsBaseCurrency(Code) Then
        Rate = 1
    Else
        Rate = CDbl(frm!txtRate.Value)
    End If
End Function

Public Sub ExpenseCurrencyChanged(ByVal frm As Access.Form, ByVal FieldName As String)
    ' frmExpenses: in another currency the amount and tax are typed in it and converted to SAR.
    Select Case FieldName
        Case "CurrencyCode", "ExpenseDate"
            If IsBaseCurrency(frm!CurrencyCode.Value) Then
                frm!ExchangeRate.Value = 1
            ElseIf FieldName = "CurrencyCode" Or Nz(frm!ExchangeRate.Value, 1) = 1 Then
                If RateOn(frm!CurrencyCode.Value, frm!ExpenseDate.Value) > 0 Then
                    frm!ExchangeRate.Value = RateOn(frm!CurrencyCode.Value, frm!ExpenseDate.Value)
                End If
            End If
    End Select
    If Not IsBaseCurrency(frm!CurrencyCode.Value) Then
        Select Case FieldName
            Case "CurrencyCode", "ExpenseDate", "ExchangeRate", "ForeignAmount", "ForeignTax"
                frm!Amount.Value = ToBase(Nz(frm!ForeignAmount.Value, 0), Nz(frm!ExchangeRate.Value, 0))
                frm!Tax.Value = ToBase(Nz(frm!ForeignTax.Value, 0), Nz(frm!ExchangeRate.Value, 0))
        End Select
    End If
End Sub

Public Function ExpenseCurrencyOK(ByVal frm As Access.Form) As Boolean
    ' Before saving an expense: SAR keeps its amounts (foreign = the same); another currency needs a
    ' rate and its typed amounts, from which the SAR amounts are computed.
    Dim msg As String
    If IsNull(frm!CurrencyCode.Value) Then frm!CurrencyCode.Value = BaseCurrency()
    If IsBaseCurrency(frm!CurrencyCode.Value) Then
        frm!ExchangeRate.Value = 1
        frm!ForeignAmount.Value = Nz(frm!Amount.Value, 0)
        frm!ForeignTax.Value = Nz(frm!Tax.Value, 0)
        ExpenseCurrencyOK = True
        Exit Function
    End If
    msg = CurrencyProblem(frm!CurrencyCode.Value, frm!ExchangeRate.Value)
    If Len(msg) = 0 And Nz(frm!ForeignAmount.Value, 0) <= 0 Then msg = "«ﬂ » «·„»·€ »⁄„·… «·„’—Ê› (" & frm!CurrencyCode.Value & ")."
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    frm!Amount.Value = ToBase(frm!ForeignAmount.Value, frm!ExchangeRate.Value)
    frm!Tax.Value = ToBase(Nz(frm!ForeignTax.Value, 0), frm!ExchangeRate.Value)
    ExpenseCurrencyOK = True
End Function

Public Function CurrencyNote(ByVal Code As String, ByVal Rate As Double, ByVal ForeignTotal As Currency) As String
    ' "1,000.00 USD ◊ 3.7500 = 3,750.00 SAR" for the status lines of the screens.
    If IsBaseCurrency(Code) Then Exit Function
    CurrencyNote = Format$(ForeignTotal, "#,##0.00") & " " & Code & " ◊ " & Format$(Rate, "0.0000") & " = " & _
                   Format$(ToBase(ForeignTotal, Rate), "#,##0.00") & " " & BaseCurrency()
End Function

'------------------------------------------------------------------------------
' Journal: the entries of documents in a currency say so (called by SyncJournal)
'------------------------------------------------------------------------------
Public Sub StampJournalCurrencies()
    Dim specs As Variant, spec As Variant, parts As Variant
    specs = Array("PURCHASE|PurchaseInvoices|PurchaseInvoiceID", "PURCHASE_RETURN|PurchaseReturns|PurchaseReturnID", _
                  "SUPPLIER_PAYMENT|SupplierPayments|PaymentID", "CUSTOMER_PAYMENT|CustomerPayments|PaymentID", _
                  "EXPENSE|Expenses|ExpenseID", "MANUAL|ManualEntries|ManualEntryID")
    For Each spec In specs
        parts = Split(spec, "|")
        CurrentDb.Execute "UPDATE JournalEntries AS e INNER JOIN [" & parts(1) & "] AS d ON e.SourceID = d.[" & parts(2) & _
            "] SET e.CurrencyCode = d.CurrencyCode, e.ExchangeRate = d.ExchangeRate, e.ForeignAmount = d.ForeignAmount " & _
            "WHERE e.SourceType = '" & parts(0) & "' AND d.CurrencyCode Is Not Null AND (e.CurrencyCode Is Null OR " & _
            "e.CurrencyCode <> d.CurrencyCode OR e.ExchangeRate <> d.ExchangeRate OR e.ForeignAmount <> d.ForeignAmount)", _
            dbFailOnError
    Next
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Public Function TestCurrency() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, passed As Long, failed As Long, report As String, msg As String
    Dim supplier As Long, payID As Long, entry As Variant, manualID As Long, before As Currency
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestCurrency  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Check IsBaseCurrency(Null) And IsBaseCurrency("SAR") And Not IsBaseCurrency("USD"), "⁄„·… «·»—‰«„Ã", _
          passed, failed, report
    Check ToBase(100, 3.75) = 375 And FromBase(375, 3.75) = 100 And ToBase(0.1, 3.75) = 0.38, "«· ÕÊÌ· Ê«· ﬁ—Ì»", _
          passed, failed, report
    Check Len(CurrencyProblem("SAR", 1)) = 0 And Len(CurrencyProblem("SAR", 2)) > 0 And _
          Len(CurrencyProblem("USD", Null)) > 0 And Len(CurrencyProblem("XXX", 1)) > 0, "«· Õﬁﬁ „‰ «·⁄„·… Ê«·„⁄«„·", _
          passed, failed, report
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    If ClosedThroughDate() >= Date Then
        Debug.Print "[--] «·› —… „ﬁ›·… Õ Ï «·ÌÊ„"
        GoTo Undo
    End If
    msg = SaveCurrencyRate("USD", Date, 3.76)
    Check Len(msg) = 0 And RateOn("USD", Date) = 3.76 And RateOn("USD", Date - 1) <> 3.76, "”⁄— «·ÌÊ„ Ê√ÕœÀ ”⁄— ﬁ»·Â " & msg, _
          passed, failed, report
    supplier = Nz(DbValue("SELECT Min(SupplierID) FROM Suppliers WHERE IsActive = True"), 0)
    before = Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierID = " & supplier), 0)
    msg = PostSupplierPayment(supplier, 100, 1, "TEST-FX", payID, "USD", 3.76)
    msg = msg & SyncJournal()
    entry = DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType = 'SUPPLIER_PAYMENT' AND SourceID = " & payID)
    Check Len(msg) = 0 And Nz(DbValue("SELECT Amount FROM SupplierPayments WHERE PaymentID = " & payID), 0) = 376 And _
          Nz(DbValue("SELECT ForeignAmount FROM SupplierPayments WHERE PaymentID = " & payID), 0) = 100 And _
          Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierID = " & supplier), 0) = before - 376, _
          "”‰œ ’—› 100 œÊ·«— = 376 —Ì«· ⁄·Ï Õ”«» «·„Ê—œ " & msg, passed, failed, report
    Check Nz(DbValue("SELECT TotalDebit FROM JournalEntries WHERE EntryID = " & Nz(entry, 0)), 0) = 376 And _
          Nz(DbValue("SELECT CurrencyCode FROM JournalEntries WHERE EntryID = " & Nz(entry, 0)), "") = "USD" And _
          Nz(DbValue("SELECT ExchangeRate FROM JournalEntries WHERE EntryID = " & Nz(entry, 0)), 0) = 3.76 And _
          Nz(DbValue("SELECT ForeignAmount FROM JournalEntries WHERE EntryID = " & Nz(entry, 0)), 0) = 100, _
          "«·ﬁÌœ »«·—Ì«· ÊÌÕ„· «·⁄„·… Ê«·„⁄«„· Ê«·„»·€ »«·œÊ·«—", passed, failed, report
    EnsureLocalTables
    CurrentDb.Execute "DELETE FROM tmpManualLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpManualLines (AccountCode, Debit, Credit) VALUES (5900, 33.33, 0)", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpManualLines (AccountCode, Debit, Credit) VALUES (5900, 33.33, 0)", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpManualLines (AccountCode, Debit, Credit) VALUES (2300, 0, 66.66)", dbFailOnError
    msg = PostManualEntry(0, Date, "TEST-FX", "", Null, manualID, "USD", 3.7777)
    Check Len(msg) = 0 And Nz(DbValue("SELECT Sum(Debit) - Sum(Credit) FROM ManualEntryLines WHERE ManualEntryID = " & _
          manualID), 1) = 0 And Nz(DbValue("SELECT Sum(ForeignDebit) FROM ManualEntryLines WHERE ManualEntryID = " & _
          manualID), 0) = 66.66, "ﬁÌœ ÌœÊÌ »«·œÊ·«—: „ Ê«“‰ »«·—Ì«· »⁄œ «· ﬁ—Ì»° Ê„»«·€Â »«·œÊ·«— „Õ›ÊŸ… " & msg, _
          passed, failed, report
    CurrentDb.Execute "DELETE FROM tmpManualLines", dbFailOnError
Undo:
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    Check False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·⁄„·«  ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestCurrency"
        TestCurrency = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestCurrency"
    End If
End Function

Private Sub Check(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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
