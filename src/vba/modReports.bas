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

Private Const ONES_TEXT As String = "|واحد|اثنان|ثلاثة|أربعة|خمسة|ستة|سبعة|ثمانية|تسعة|عشرة|أحد عشر|" & _
    "اثنا عشر|ثلاثة عشر|أربعة عشر|خمسة عشر|ستة عشر|سبعة عشر|ثمانية عشر|تسعة عشر"
Private Const TENS_TEXT As String = "||عشرون|ثلاثون|أربعون|خمسون|ستون|سبعون|ثمانون|تسعون"
Private Const HUNDREDS_TEXT As String = "|مائة|مائتان|ثلاثمائة|أربعمائة|خمسمائة|ستمائة|سبعمائة|" & _
    "ثمانمائة|تسعمائة"

'------------------------------------------------------------------------------
' Report text helpers
'------------------------------------------------------------------------------
Public Function ReportCriteria() As String
    ' Set by OpenReportOrQuery before a report opens ("" when the report has no choices).
    On Error Resume Next
    ReportCriteria = Nz(TempVars("ReportCriteria"), "")
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
    ReportPrintedAt = "طُبع في " & GDate(Now, True) & "  بواسطة " & CurrentUserName()
End Function

Public Sub ReportNoData(ByRef Cancel As Integer, ByVal Message As String)
    ShowInfo Message
    Cancel = True
End Sub

'------------------------------------------------------------------------------
' Amount in words (Arabic)
'------------------------------------------------------------------------------
Public Function AmountInWords(ByVal Amount As Variant) As String
    ' 1250.5 -> "فقط ألف ومائتان وخمسون ريال سعودي وخمسون هللة لا غير"
    Dim a As Currency, riyals As Long, halalas As Long, text As String
    If IsNull(Amount) Then Exit Function
    a = RoundMoney(Abs(CCur(Amount)))
    If a >= 1000000000 Then
        AmountInWords = Format$(a, "#,##0.00") & " ريال سعودي"
        Exit Function
    End If
    riyals = CLng(Fix(a))
    halalas = CLng((a - riyals) * 100)
    If riyals = 0 And halalas = 0 Then
        AmountInWords = "صفر ريال"
        Exit Function
    End If
    If riyals > 0 Then text = NumberWords(riyals) & " ريال سعودي"
    If halalas > 0 Then
        If Len(text) > 0 Then text = text & " و"
        text = text & NumberWords(halalas) & " هللة"
    End If
    AmountInWords = "فقط " & text & " لا غير"
End Function

Private Function NumberWords(ByVal n As Long) As String
    Dim millions As Long, thousands As Long, units As Long, parts As String
    If n = 0 Then
        NumberWords = "صفر"
        Exit Function
    End If
    millions = n \ 1000000
    thousands = (n \ 1000) Mod 1000
    units = n Mod 1000
    If millions > 0 Then parts = ScaleWords(millions, "مليون", "مليونان", "ملايين")
    If thousands > 0 Then parts = JoinWords(parts, ScaleWords(thousands, "ألف", "ألفان", "آلاف"))
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
            If u > 0 Then tens = WordAt(ONES_TEXT, u) & " و" & tens
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
        JoinWords = Left1 & " و" & Right1
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
        ShowWarning "اختر جلسة الجرد أولًا."
        Exit Sub
    End If
    OpenDocumentReport "rptStockCount", "[StockCountID] = " & CLng(StockCountID)
End Sub

Public Sub PrintPartyStatement(ByVal PartyKind As String, ByVal PartyID As Variant)
    ' Account statement of the current customer ("C") or supplier ("S") for this year.
    Dim partyName As Variant
    If IsNull(PartyID) Then
        ShowWarning "احفظ السجل أولًا."
        Exit Sub
    End If
    Calendar = vbCalGreg
    SetPeriod DateSerial(Year(Date), 1, 1), Date
    If PartyKind = "C" Then
        SetQueryParam "CustomerID", CLng(PartyID)
        partyName = DLookup("CustomerName", "Customers", "CustomerID = " & CLng(PartyID))
        OpenReportOrQuery "rptCustomerStatement", "CustomerStatementQuery", "", _
                          PeriodText(DateSerial(Year(Date), 1, 1), Date) & "    العميل: " & Nz(partyName, "")
    Else
        SetQueryParam "SupplierID", CLng(PartyID)
        partyName = DLookup("SupplierName", "Suppliers", "SupplierID = " & CLng(PartyID))
        OpenReportOrQuery "rptSupplierStatement", "SupplierStatementQuery", "", _
                          PeriodText(DateSerial(Year(Date), 1, 1), Date) & "    المورد: " & Nz(partyName, "")
    End If
End Sub

Public Function PeriodText(ByVal FromDate As Date, ByVal ToDate As Date) As String
    PeriodText = "الفترة: من " & GDate(FromDate) & " إلى " & GDate(ToDate)
End Function

Private Sub OpenDocumentReport(ByVal ReportName As String, ByVal WhereCondition As String)
    If Not ReportExists(ReportName) Then
        ShowWarning "تقرير الطباعة غير موجود: " & ReportName & vbCrLf & "شغّل BuildReports."
        Exit Sub
    End If
    On Error GoTo EH
    TempVars.Add "ReportCriteria", ""
    DoCmd.OpenReport ReportName, POS_PRINT_VIEW, , WhereCondition
    Exit Sub
EH:
    If Err.Number <> 2501 Then ShowError "تعذر فتح التقرير: " & Err.Description   ' 2501 = cancelled (no data)
End Sub

Public Function ReportsFolder() As String
    ' Exported PDF / Excel files: a "Reports" folder next to the front-end file.
    Dim folder As String
    folder = CurrentProject.Path & "\Reports"
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    ReportsFolder = folder
End Function
