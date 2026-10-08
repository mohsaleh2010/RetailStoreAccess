Attribute VB_Name = "modReports"
'==============================================================================
' modReports  -  Retail Store Management System (Phase 8)
'
' Functions used by the formatted reports (built by modBuildReports):
'   ReportCriteria    the "period / customer / product" line under the title
'   GDate             Gregorian date text, whatever the Windows calendar is
'   ReportPrintedAt   printing time in the page footer
'   AmountInWords     amount in Arabic words for vouchers (same as tools/tafqeet.py)
'   ReportNoData      message instead of an empty report
' Printing of single documents:
'   PrintPurchaseDocument, PrintVoucher, PrintStockCount, PrintPartyStatement
'==============================================================================
Option Compare Database
Option Explicit

Private Const ONES_TEXT As String = "|Ê«Õœ|«À‰«‰|À·«À…|√—»⁄…|Œ„”…|” …|”»⁄…|À„«‰Ì…| ”⁄…|⁄‘—…|√Õœ ⁄‘—|" & _
    "«À‰« ⁄‘—|À·«À… ⁄‘—|√—»⁄… ⁄‘—|Œ„”… ⁄‘—|” … ⁄‘—|”»⁄… ⁄‘—|À„«‰Ì… ⁄‘—| ”⁄… ⁄‘—"
Private Const TENS_TEXT As String = "||⁄‘—Ê‰|À·«ÀÊ‰|√—»⁄Ê‰|Œ„”Ê‰|” Ê‰|”»⁄Ê‰|À„«‰Ê‰| ”⁄Ê‰"
Private Const HUNDREDS_TEXT As String = "|„«∆…|„«∆ «‰|À·«À„«∆…|√—»⁄„«∆…|Œ„”„«∆…|” „«∆…|”»⁄„«∆…|" & _
    "À„«‰„«∆…| ”⁄„«∆…"

'------------------------------------------------------------------------------
' Report text helpers
'------------------------------------------------------------------------------
Public Function ReportCriteria() As String
    ' Set by OpenReportOrQuery before a report opens ("" when the report has no choices).
    On Error Resume Next
    ReportCriteria = Tr(Nz(TempVars("ReportCriteria"), ""))
End Function

Public Function GDate(ByVal Value As Variant, Optional ByVal WithTime As Boolean = False) As String
    ' Saudi PCs may use the Hijri calendar by default; reports always show Gregorian dates.
    Dim saved As Integer
    If IsNull(Value) Or IsEmpty(Value) Then Exit Function
    If Not IsDate(Value) Then Exit Function
    saved = Calendar
    Calendar = vbCalGreg
    If WithTime Then
        GDate = Format$(CDate(Value), "yyyy/mm/dd hh:nn")
    Else
        GDate = Format$(CDate(Value), "yyyy/mm/dd")
    End If
    Calendar = saved
End Function

Public Function ReportPrintedAt() As String
    ReportPrintedAt = Tr("ÿı»⁄ ›Ì " & GDate(Now, True) & "  »Ê«”ÿ… ") & CurrentUserName()
End Function

Public Sub ReportNoData(ByRef Cancel As Integer, ByVal Message As String)
    ShowInfo Message
    Cancel = True
End Sub

'------------------------------------------------------------------------------
' Amount in words (Arabic)
'------------------------------------------------------------------------------
Public Function AmountInWords(ByVal Amount As Variant, Optional ByVal CurrencyCode As String = "") As String
    ' In the language of the interface (modLang) and the program currency (modCountry: riyal or pound).
    If Len(CurrencyCode) = 0 Then CurrencyCode = BaseCurrency()
    If UiEnglish() Then
        AmountInWords = AmountInWordsEn(Amount, CurrencyCode)
    Else
        AmountInWords = AmountInWordsAr(Amount, CurrencyCode)
    End If
End Function

Private Function CurrencyNames(ByVal CurrencyCode As String) As Variant
    ' Arabic unit, Arabic sub-unit, Arabic zero, English unit (one, many), English sub-unit (one, many)
    If CurrencyCode = "EGP" Then
        CurrencyNames = Array("Ã‰ÌÂ „’—Ì", "ﬁ—‘", "’›— Ã‰ÌÂ", "Egyptian Pound", "Egyptian Pounds", "Piaster", "Piasters")
    Else
        CurrencyNames = Array("—Ì«· ”⁄ÊœÌ", "Â··…", "’›— —Ì«·", "Saudi Riyal", "Saudi Riyals", "Halala", "Halalas")
    End If
End Function

Public Function AmountInWordsAr(ByVal Amount As Variant, Optional ByVal CurrencyCode As String = "SAR") As String
    ' 1250.5 -> "›ﬁÿ √·› Ê„«∆ «‰ ÊŒ„”Ê‰ —Ì«· ”⁄ÊœÌ ÊŒ„”Ê‰ Â··… ·« €Ì—" (... Ã‰ÌÂ „’—Ì ÊŒ„”Ê‰ ﬁ—‘ ... for EGP)
    Dim a As Currency, units As Long, cents As Long, text As String, names As Variant
    If IsNull(Amount) Then Exit Function
    names = CurrencyNames(CurrencyCode)
    a = RoundMoney(Abs(CCur(Amount)))
    If a >= 1000000000 Then
        AmountInWordsAr = Format$(a, "#,##0.00") & " " & names(0)
        Exit Function
    End If
    units = CLng(Fix(a))
    cents = CLng((a - units) * 100)
    If units = 0 And cents = 0 Then
        AmountInWordsAr = names(2)
        Exit Function
    End If
    If units > 0 Then text = NumberWords(units) & " " & names(0)
    If cents > 0 Then
        If Len(text) > 0 Then text = text & " Ê"
        text = text & NumberWords(cents) & " " & names(1)
    End If
    AmountInWordsAr = "›ﬁÿ " & text & " ·« €Ì—"
End Function

Public Function AmountInWordsEn(ByVal Amount As Variant, Optional ByVal CurrencyCode As String = "SAR") As String
    ' 1250.5 -> "Only one thousand two hundred fifty Saudi Riyals and fifty Halalas" (tools/tafqeet.py)
    Dim a As Currency, units As Long, cents As Long, words As String, names As Variant
    If IsNull(Amount) Then Exit Function
    names = CurrencyNames(CurrencyCode)
    a = RoundMoney(Abs(CCur(Amount)))
    If a >= 1000000000 Then
        AmountInWordsEn = Format$(a, "#,##0.00") & " " & names(4)
        Exit Function
    End If
    units = CLng(Fix(a))
    cents = CLng((a - units) * 100)
    If units = 0 And cents = 0 Then
        AmountInWordsEn = "Zero " & names(4)
        Exit Function
    End If
    If units > 0 Then words = NumberWordsEn(units) & " " & IIf(units = 1, names(3), names(4))
    If cents > 0 Then
        If Len(words) > 0 Then words = words & " and "
        words = words & NumberWordsEn(cents) & " " & IIf(cents = 1, names(5), names(6))
    End If
    AmountInWordsEn = "Only " & words
End Function

Private Function NumberWordsEn(ByVal n As Long) As String
    Dim parts As String
    If n >= 1000000 Then parts = Below1000En(n \ 1000000) & " million"
    If (n \ 1000) Mod 1000 > 0 Then parts = Trim$(parts & " " & Below1000En((n \ 1000) Mod 1000) & " thousand")
    If n Mod 1000 > 0 Then parts = Trim$(parts & " " & Below1000En(n Mod 1000))
    NumberWordsEn = parts
End Function

Private Function Below1000En(ByVal n As Long) As String
    Dim ones As Variant, tens As Variant, parts As String, r As Long
    ones = Split("zero one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen " & _
                 "sixteen seventeen eighteen nineteen", " ")
    tens = Split("- - twenty thirty forty fifty sixty seventy eighty ninety", " ")
    If n >= 100 Then parts = ones(n \ 100) & " hundred"
    r = n Mod 100
    If r >= 20 Then
        parts = Trim$(parts & " " & tens(r \ 10))
        If r Mod 10 > 0 Then parts = parts & "-" & ones(r Mod 10)
    ElseIf r > 0 Then
        parts = Trim$(parts & " " & ones(r))
    End If
    Below1000En = parts
End Function

Private Function NumberWords(ByVal n As Long) As String
    Dim millions As Long, thousands As Long, units As Long, parts As String
    If n = 0 Then
        NumberWords = "’›—"
        Exit Function
    End If
    millions = n \ 1000000
    thousands = (n \ 1000) Mod 1000
    units = n Mod 1000
    If millions > 0 Then parts = ScaleWords(millions, "„·ÌÊ‰", "„·ÌÊ‰«‰", "„·«ÌÌ‰")
    If thousands > 0 Then parts = JoinWords(parts, ScaleWords(thousands, "√·›", "√·›«‰", "¬·«›"))
    If units > 0 Then parts = JoinWords(parts, Below1000(units))
    NumberWords = parts
End Function

Private Function ScaleWords(ByVal n As Long, ByVal One As String, ByVal Two As String, ByVal Few As String) As String
    If n = 1 Then
        ScaleWords = One
    ElseIf n = 2 Then
        ScaleWords = Two
    ElseIf n <= 10 Then
        ScaleWords = Below1000(n) & " " & Few
    Else
        ScaleWords = Below1000(n) & " " & One
    End If
End Function

Private Function Below1000(ByVal n As Long) As String
    Dim h As Long, r As Long, t As Long, u As Long, parts As String, tens As String
    h = n \ 100
    r = n Mod 100
    If h > 0 Then parts = WordAt(HUNDREDS_TEXT, h)
    If r > 0 Then
        If r < 20 Then
            parts = JoinWords(parts, WordAt(ONES_TEXT, r))
        Else
            t = r \ 10
            u = r Mod 10
            tens = WordAt(TENS_TEXT, t)
            If u > 0 Then tens = WordAt(ONES_TEXT, u) & " Ê" & tens
            parts = JoinWords(parts, tens)
        End If
    End If
    Below1000 = parts
End Function

Private Function WordAt(ByVal List As String, ByVal Index As Long) As String
    Dim items() As String
    items = Split(List, "|")
    WordAt = items(Index)
End Function

Private Function JoinWords(ByVal Left1 As String, ByVal Right1 As String) As String
    If Len(Left1) = 0 Then
        JoinWords = Right1
    Else
        JoinWords = Left1 & " Ê" & Right1
    End If
End Function

'------------------------------------------------------------------------------
' Printing single documents
'------------------------------------------------------------------------------
Public Sub PrintPurchaseDocument(ByVal DocKind As String, ByVal DocID As Variant)
    ' DocKind: "PURCHASE" (invoice) or "RETURN" (purchase return)
    If IsNull(DocID) Then Exit Sub
    OpenDocumentReport "rptPurchaseDocument", "[DocKind] = '" & DocKind & "' AND [DocID] = " & CLng(DocID)
End Sub

Public Sub PrintVoucher(ByVal DocKind As String, ByVal DocID As Variant)
    ' DocKind: "RECEIPT" (customer payment) or "PAYMENT" (supplier payment)
    If IsNull(DocID) Then Exit Sub
    OpenDocumentReport "rptVoucher", "[DocKind] = '" & DocKind & "' AND [DocID] = " & CLng(DocID)
End Sub

Public Sub PrintStockCount(ByVal StockCountID As Variant)
    If IsNull(StockCountID) Then
        ShowWarning "«Œ — Ã·”… «·Ã—œ √Ê·«."
        Exit Sub
    End If
    OpenDocumentReport "rptStockCount", "[StockCountID] = " & CLng(StockCountID)
End Sub

Public Sub PrintPartyStatement(ByVal PartyKind As String, ByVal PartyID As Variant)
    ' Account statement of the current customer ("C") or supplier ("S") for this year.
    Dim partyName As Variant
    If IsNull(PartyID) Then
        ShowWarning "«Õ›Ÿ «·”Ã· √Ê·«."
        Exit Sub
    End If
    Calendar = vbCalGreg
    SetPeriod DateSerial(Year(Date), 1, 1), Date
    If PartyKind = "C" Then
        SetQueryParam "CustomerID", CLng(PartyID)
        partyName = DbValue(Tr("SELECT c.CustomerName FROM [@Customers] AS c WHERE c.CustomerID = " & CLng(PartyID)))
        OpenReportOrQuery "rptCustomerStatement", "CustomerStatementQuery", "", _
                          PeriodText(DateSerial(Year(Date), 1, 1), Date) & "    «·⁄„Ì·: " & Nz(partyName, "")
    Else
        SetQueryParam "SupplierID", CLng(PartyID)
        partyName = DbValue(Tr("SELECT s.SupplierName FROM [@Suppliers] AS s WHERE s.SupplierID = " & CLng(PartyID)))
        OpenReportOrQuery "rptSupplierStatement", "SupplierStatementQuery", "", _
                          PeriodText(DateSerial(Year(Date), 1, 1), Date) & "    «·„Ê—œ: " & Nz(partyName, "")
    End If
End Sub

Public Function PeriodText(ByVal FromDate As Date, ByVal ToDate As Date) As String
    PeriodText = "«·› —…: „‰ " & GDate(FromDate) & " ≈·Ï " & GDate(ToDate)
End Function

Private Sub OpenDocumentReport(ByVal ReportName As String, ByVal WhereCondition As String)
    If Not ReportExists(ReportName) Then
        ShowWarning " ﬁ—Ì— «·ÿ»«⁄… €Ì— „ÊÃÊœ: " & ReportName & vbCrLf & "‘€¯· BuildReports."
        Exit Sub
    End If
    On Error GoTo EH
    TempVars.Add "ReportCriteria", ""
    DoCmd.OpenReport ReportName, POS_PRINT_VIEW, , WhereCondition
    Exit Sub
EH:
    If Err.Number <> 2501 Then ShowError " ⁄–— › Õ «· ﬁ—Ì—: " & Err.Description   ' 2501 = cancelled (no data)
End Sub

Public Function ReportsFolder() As String
    ' Exported PDF / Excel files: a "Reports" folder next to the front-end file.
    Dim folder As String
    folder = CurrentProject.Path & "\Reports"
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    ReportsFolder = folder
End Function
