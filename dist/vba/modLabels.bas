Attribute VB_Name = "modLabels"
'==============================================================================
' modLabels  -  Retail Store Management System: barcode labels
'
'   Encoding     BarcodePattern(code) -> "1010..." one character per module:
'                EAN-13 for 13 digits with a valid check digit, otherwise Code 128
'                (set C for digit pairs, set B for the rest). Same output as
'                tools/barcode_reference.py, which a real barcode reader decodes.
'   Drawing      DrawBarcode, called from the Print event of rptBarcodeLabels.
'   Print list   local tables tmpLabelQueue (product, copies) and tmpLabelNumbers
'                (1..500, one report row per copy).
'   Layout       ApplyLabelLayout opens rptBarcodeLabels in design view and applies
'                LabelSettings: label size, labels per row, gaps, margins, printer,
'                the text lines above/below the barcode and the font size.
'   Screens      frmBarcodeLabels + frmLabelLines, frmLabelSettings.
'   Test         TestLabels (also part of RunAllTests).
'==============================================================================
Option Compare Database
Option Explicit

Private Const LABEL_REPORT As String = "rptBarcodeLabels"
Private Const MAX_COPIES As Long = 500
Private Const TWIPS_PER_MM As Single = 56.7
Private Const PR_HORIZONTAL_LAYOUT As Long = 1953     ' acPRHorizontalColumnLayout

Private m_c128 As Variant
Private m_barMM As Single

'------------------------------------------------------------------------------
' Encoding
'------------------------------------------------------------------------------
Public Function BarcodePattern(ByVal Code As String) As String
    Code = Trim$(Code)
    If EanCheckOk(Code) Then
        BarcodePattern = Ean13Pattern(Code)
    Else
        BarcodePattern = Code128Pattern(Code)
    End If
End Function

Public Function EanCheckOk(ByVal Code As String) As Boolean
    Dim i As Long, total As Long, d As Long
    If Len(Code) <> 13 Then Exit Function
    For i = 1 To 13
        If Mid$(Code, i, 1) < "0" Or Mid$(Code, i, 1) > "9" Then Exit Function
    Next
    For i = 1 To 12
        d = CLng(Mid$(Code, i, 1))
        If i Mod 2 = 0 Then
            total = total + 3 * d
        Else
            total = total + d
        End If
    Next
    EanCheckOk = ((10 - total Mod 10) Mod 10 = CLng(Mid$(Code, 13, 1)))
End Function

Private Function EanPart(ByVal Set1 As String, ByVal Digit As Long) As String
    Dim parts As Variant
    Select Case Set1
        Case "L"
            parts = Split("0001101,0011001,0010011,0111101,0100011,0110001,0101111,0111011,0110111,0001011", ",")
        Case "G"
            parts = Split("0100111,0110011,0011011,0100001,0011101,0111001,0000101,0010001,0001001,0010111", ",")
        Case Else
            parts = Split("1110010,1100110,1101100,1000010,1011100,1001110,1010000,1000100,1001000,1110100", ",")
    End Select
    EanPart = parts(Digit)
End Function

Private Function Ean13Pattern(ByVal Code As String) As String
    Dim parity As Variant, p As String, i As Long, s As String
    parity = Split("LLLLLL,LLGLGG,LLGGLG,LLGGGL,LGLLGG,LGGLLG,LGGGLL,LGLGLG,LGLGGL,LGGLGL", ",")
    p = parity(CLng(Left$(Code, 1)))
    s = "101"
    For i = 1 To 6
        s = s & EanPart(Mid$(p, i, 1), CLng(Mid$(Code, i + 1, 1)))
    Next
    s = s & "01010"
    For i = 8 To 13
        s = s & EanPart("R", CLng(Mid$(Code, i, 1)))
    Next
    Ean13Pattern = s & "101"
End Function

Private Sub InitC128()
    ' bar/space widths of the Code 128 values 0..106 (106 = stop)
    If IsArray(m_c128) Then Exit Sub
    m_c128 = Split("212222 222122 222221 121223 121322 131222 122213 122312 132212 221213 " & _
        "221312 231212 112232 122132 122231 113222 123122 123221 223211 221132 " & _
        "221231 213212 223112 312131 311222 321122 321221 312212 322112 322211 " & _
        "212123 212321 232121 111323 131123 131321 112313 132113 132311 211313 " & _
        "231113 231311 112133 112331 132131 113123 113321 133121 313121 211331 " & _
        "231131 213113 213311 213131 311123 311321 331121 312113 312311 332111 " & _
        "314111 221411 431111 111224 111422 121124 121421 141122 141221 112214 " & _
        "112412 122114 122411 142112 142211 241211 221114 413111 241112 134111 " & _
        "111242 121142 121241 114212 124112 124211 411212 421112 421211 212141 " & _
        "214121 412121 111143 111341 131141 114113 114311 411113 411311 113141 " & _
        "114131 311141 411131 211412 211214 211232 2331112", " ")
End Sub

Public Function Code128Pattern(ByVal Text As String) As String
    Dim vals() As Long, n As Long, i As Long, ch As Long, allDigits As Boolean
    Dim pairs As Long, chk As Long, s As String
    If Len(Text) = 0 Then Exit Function
    allDigits = True
    For i = 1 To Len(Text)
        ch = AscW(Mid$(Text, i, 1))
        If ch < 32 Or ch > 126 Then Exit Function         ' not printable ASCII (e.g. Arabic)
        If ch < 48 Or ch > 57 Then allDigits = False
    Next
    ReDim vals(0 To Len(Text) + 3)
    If allDigits And Len(Text) >= 2 Then
        vals(0) = 105                                    ' start C: two digits per symbol
        n = 1
        pairs = Len(Text) \ 2
        For i = 1 To pairs
            vals(n) = CLng(Mid$(Text, 2 * i - 1, 2))
            n = n + 1
        Next
        If Len(Text) Mod 2 = 1 Then
            vals(n) = 100                                ' code B for the last digit
            vals(n + 1) = AscW(Right$(Text, 1)) - 32
            n = n + 2
        End If
    Else
        vals(0) = 104                                    ' start B
        n = 1
        For i = 1 To Len(Text)
            vals(n) = AscW(Mid$(Text, i, 1)) - 32
            n = n + 1
        Next
    End If
    chk = vals(0)
    For i = 1 To n - 1
        chk = chk + i * vals(i)
    Next
    vals(n) = chk Mod 103
    vals(n + 1) = 106
    n = n + 2
    InitC128
    For i = 0 To n - 1
        s = s & WidthsToModules(CStr(m_c128(vals(i))))
    Next
    Code128Pattern = s
End Function

Private Function WidthsToModules(ByVal Widths As String) As String
    Dim i As Long, s As String, isBar As Boolean
    isBar = True
    For i = 1 To Len(Widths)
        If isBar Then
            s = s & String$(CLng(Mid$(Widths, i, 1)), "1")
        Else
            s = s & String$(CLng(Mid$(Widths, i, 1)), "0")
        End If
        isBar = Not isBar
    Next
    WidthsToModules = s
End Function

Public Function LabelCode(ByVal Barcode As Variant, ByVal ProductCode As Variant) As String
    ' The product barcode when it can be encoded, otherwise the product code.
    Dim c As String
    c = Trim$(Nz(Barcode, ""))
    If Len(c) > 0 Then
        If Len(BarcodePattern(c)) > 0 Then
            LabelCode = c
            Exit Function
        End If
    End If
    LabelCode = Trim$(Nz(ProductCode, ""))
End Function

Public Function LabelPrice(ByVal Price As Variant) As String
    LabelPrice = Format$(Nz(Price, 0), "#,##0.00") & " —.”"
End Function

Public Function UnitsToCopies(ByVal Units As Double) As Long
    ' one label per unit; a fraction (weight) still needs one label
    Dim c As Long
    c = -Int(-Units)
    If c < 1 Then c = 1
    If c > MAX_COPIES Then c = MAX_COPIES
    UnitsToCopies = c
End Function

'------------------------------------------------------------------------------
' Drawing (Print event of the label section)
'------------------------------------------------------------------------------
Public Sub DrawBarcode(ByVal rpt As Access.Report, ByVal Code As String, ByVal LeftTwips As Single, _
                       ByVal TopTwips As Single, ByVal WidthTwips As Single, ByVal HeightTwips As Single, _
                       ByVal ModuleMM As Single)
    Dim p As String, n As Long, m As Single, x0 As Single, i As Long, j As Long
    p = BarcodePattern(Code)
    n = Len(p)
    If n = 0 Or WidthTwips <= 0 Or HeightTwips <= 0 Then Exit Sub
    m = ModuleMM * TWIPS_PER_MM
    If m * (n + 14) > WidthTwips Then m = WidthTwips / (n + 14)   ' 7-module quiet zone on each side
    x0 = LeftTwips + (WidthTwips - m * n) / 2
    i = 1
    Do While i <= n
        If Mid$(p, i, 1) = "1" Then
            j = i
            Do While j < n
                If Mid$(p, j + 1, 1) <> "1" Then Exit Do
                j = j + 1
            Loop
            rpt.Line (x0 + (i - 1) * m, TopTwips)-(x0 + j * m, TopTwips + HeightTwips), 0, BF
            i = j + 1
        Else
            i = i + 1
        End If
    Loop
End Sub

Public Sub LabelReportOpen()
    m_barMM = Nz(LabelSetting("BarWidth"), 0.25)
End Sub

Public Function LabelBarWidth() As Single
    If m_barMM <= 0 Then LabelReportOpen
    LabelBarWidth = m_barMM
End Function

'------------------------------------------------------------------------------
' Settings and local tables
'------------------------------------------------------------------------------
Public Function LabelSetting(ByVal FieldName As String) As Variant
    On Error Resume Next
    LabelSetting = Null
    LabelSetting = DbValue("SELECT [" & FieldName & "] FROM LabelSettings WHERE LabelSettingID = 1")
End Function

Private Function LabelTableExists(ByVal TableName As String) As Boolean
    Dim db As DAO.Database, tdf As DAO.TableDef
    Set db = CurrentDb
    For Each tdf In db.TableDefs
        If StrComp(tdf.Name, TableName, vbTextCompare) = 0 And Len(tdf.Connect) = 0 Then
            LabelTableExists = True
            Exit Function
        End If
    Next
End Function

Public Sub EnsureLabelTables()
    Dim db As DAO.Database, rs As DAO.Recordset, i As Long
    Set db = CurrentDb
    If Not LabelTableExists("tmpLabelQueue") Then
        db.Execute "CREATE TABLE tmpLabelQueue (LineNo COUNTER CONSTRAINT pkLabelQueue PRIMARY KEY, " & _
                   "ProductID LONG, ProductName TEXT(150), LabelCode TEXT(50), Price CURRENCY, Copies LONG)", _
                   dbFailOnError
    End If
    If Not LabelTableExists("tmpLabelNumbers") Then
        db.Execute "CREATE TABLE tmpLabelNumbers (N LONG CONSTRAINT pkLabelNumbers PRIMARY KEY)", dbFailOnError
        Set rs = db.OpenRecordset("tmpLabelNumbers", dbOpenDynaset, dbAppendOnly)
        For i = 1 To MAX_COPIES
            rs.AddNew
            rs!N = i
            rs.Update
        Next
        rs.Close
    End If
End Sub

'------------------------------------------------------------------------------
' Print list
'------------------------------------------------------------------------------
Public Function AddLabelProduct(ByVal ProductID As Long, ByVal Copies As Long) As String
    ' "" when added (copies are added to the same product's line), otherwise the message.
    Dim rs As DAO.Recordset, p As DAO.Recordset
    EnsureLabelTables
    If Copies < 1 Then Copies = 1
    Set p = CurrentDb.OpenRecordset("SELECT ProductName, Barcode, ProductCode, SellingPrice FROM Products " & _
                                    "WHERE ProductID = " & ProductID, dbOpenSnapshot)
    If p.EOF Then
        p.Close
        AddLabelProduct = "«·„‰ Ã €Ì— „ÊÃÊœ."
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tmpLabelQueue WHERE ProductID = " & ProductID, dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!ProductID = ProductID
        rs!ProductName = Left$(p!ProductName, 150)
        rs!LabelCode = Left$(LabelCode(p!Barcode, p!ProductCode), 50)
        rs!Price = Nz(p!SellingPrice, 0)
        If Copies > MAX_COPIES Then Copies = MAX_COPIES
        rs!Copies = Copies
    Else
        rs.Edit
        If rs!Copies + Copies > MAX_COPIES Then
            rs!Copies = MAX_COPIES
        Else
            rs!Copies = rs!Copies + Copies
        End If
    End If
    rs.Update
    rs.Close
    p.Close
End Function

Public Function LabelLineCount() As Long
    EnsureLabelTables
    LabelLineCount = Nz(DbValue("SELECT COUNT(*) FROM tmpLabelQueue"), 0)
End Function

Public Function LabelTotalCopies() As Long
    EnsureLabelTables
    LabelTotalCopies = Nz(DbValue("SELECT SUM(Copies) FROM tmpLabelQueue"), 0)
End Function

Public Function AddPurchaseLabels(ByVal PurchaseInvoiceID As Long) As Long
    ' one label per purchased unit of every product of the invoice; returns the products added
    Dim rs As DAO.Recordset, n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT ProductID, SUM(Quantity) AS Units FROM PurchaseInvoiceDetails " & _
                                     "WHERE PurchaseInvoiceID = " & PurchaseInvoiceID & " GROUP BY ProductID", _
                                     dbOpenSnapshot)
    Do Until rs.EOF
        If Len(AddLabelProduct(rs!ProductID, UnitsToCopies(rs!Units))) = 0 Then n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    AddPurchaseLabels = n
End Function

Public Function AddCartLabels() As Long
    ' the lines of the purchase invoice that is being entered (not saved yet)
    Dim rs As DAO.Recordset, n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT ProductID, SUM(Quantity) AS Units FROM tmpPurchaseLines " & _
                                     "GROUP BY ProductID", dbOpenSnapshot)
    Do Until rs.EOF
        If Len(AddLabelProduct(rs!ProductID, UnitsToCopies(Nz(rs!Units, 1)))) = 0 Then n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    AddCartLabels = n
End Function

Public Sub ClearLabelQueue()
    EnsureLabelTables
    CurrentDb.Execute "DELETE FROM tmpLabelQueue", dbFailOnError
End Sub

'------------------------------------------------------------------------------
' Report layout from LabelSettings
'------------------------------------------------------------------------------
Private Function MM(ByVal FieldName As String, ByVal DefaultMM As Double) As Long
    MM = CLng(CDbl(Nz(LabelSetting(FieldName), DefaultMM)) * TWIPS_PER_MM)
End Function

Private Function LineSource(ByVal Kind As String) As String
    Dim s As String
    Select Case Kind
        Case "STORE"
            s = Nz(LabelSetting("ShortName"), "")
            If Len(s) = 0 Then s = Nz(SettingValue("StoreName"), "")
            LineSource = "=""" & Replace(s, """", """""") & """"
        Case "NAME"
            LineSource = "=[ProductName]"
        Case "PRICE"
            LineSource = "=LabelPrice([SellingPrice])"
        Case "CODE"
            LineSource = "=[ProductCode]"
        Case Else
            LineSource = "=LabelCode([Barcode],[ProductCode])"
    End Select
End Function

Private Sub PlaceLine(ByVal ctl As Access.Control, ByVal Kind As String, ByVal TopTwips As Long, _
                      ByVal LabelW As Long, ByVal Pad As Long, ByVal LineH As Long, ByVal FontSize As Long)
    If Kind = "NONE" Or Len(Kind) = 0 Then
        ctl.Visible = False
        Exit Sub
    End If
    ctl.ControlSource = LineSource(Kind)
    ctl.FontSize = FontSize
    ctl.Visible = True
    ctl.Move Pad, TopTwips, LabelW - 2 * Pad, LineH
End Sub

Public Function ApplyLabelLayout() As String
    ' "" when rptBarcodeLabels now matches LabelSettings, otherwise the message.
    Dim r As Access.Report, w As Long, h As Long, pad As Long, lineH As Long, fs As Long, prn As String
    Dim y As Long, yb As Long, barH As Long, room As Long, kinds(1 To 4) As String, i As Long, ctl As Variant
    On Error GoTo EH
    If Not ReportExists(LABEL_REPORT) Then
        ApplyLabelLayout = " ﬁ—Ì— «·„·’ﬁ«  €Ì— „ÊÃÊœ. ‘€¯· BuildReports."
        Exit Function
    End If
    EnsureLabelTables
    w = MM("LabelWidth", 38)
    h = MM("LabelHeight", 25)
    pad = CLng(TWIPS_PER_MM)                                  ' 1 mm inside the label
    fs = Nz(LabelSetting("FontSize"), 7)
    lineH = CLng(fs * 20 * 1.35)
    kinds(1) = Nz(LabelSetting("TopLine1"), "NONE")
    kinds(2) = Nz(LabelSetting("TopLine2"), "NONE")
    kinds(3) = Nz(LabelSetting("BottomLine1"), "NONE")
    kinds(4) = Nz(LabelSetting("BottomLine2"), "NONE")
    prn = Nz(LabelSetting("PrinterName"), "")

    If CurrentProject.AllReports(LABEL_REPORT).IsLoaded Then DoCmd.Close acReport, LABEL_REPORT, acSaveNo
    DoCmd.OpenReport LABEL_REPORT, acViewDesign, , , acHidden
    Set r = Reports(LABEL_REPORT)
    If Len(prn) > 0 Then                                      ' the printer first: it resets the page setup
        Set r.Printer = Application.Printers(prn)
    Else
        r.UseDefaultPrinter = True
    End If
    With r.Printer
        .TopMargin = MM("MarginTop", 0)
        .BottomMargin = MM("MarginBottom", 0)
        .LeftMargin = MM("MarginLeft", 0)
        .RightMargin = MM("MarginRight", 0)
        .DefaultSize = False
        .ItemsAcross = Nz(LabelSetting("LabelsAcross"), 1)
        .ColumnSpacing = MM("ColumnGap", 2)
        .RowSpacing = MM("RowGap", 0)
        .ItemSizeWidth = w
        .ItemSizeHeight = h
        .ItemLayout = PR_HORIZONTAL_LAYOUT
    End With
    For Each ctl In Array("txtTop1", "txtTop2", "txtBottom1", "txtBottom2", "txtCode", "boxBar")
        r.Controls(CStr(ctl)).Move 0, 0, 60, 60                 ' out of the way while the label shrinks
    Next
    r.Width = w
    r.Section(0).Height = h

    y = pad
    PlaceLine r.Controls("txtTop1"), kinds(1), y, w, pad, lineH, fs
    If r.Controls("txtTop1").Visible Then y = y + lineH
    PlaceLine r.Controls("txtTop2"), kinds(2), y, w, pad, lineH, fs
    If r.Controls("txtTop2").Visible Then y = y + lineH
    yb = h - pad
    If kinds(4) <> "NONE" Then yb = yb - lineH
    PlaceLine r.Controls("txtBottom2"), kinds(4), yb, w, pad, lineH, fs
    If kinds(3) <> "NONE" Then yb = yb - lineH
    PlaceLine r.Controls("txtBottom1"), kinds(3), yb, w, pad, lineH, fs
    room = yb - y - 2 * 30
    barH = MM("BarHeight", 10)
    If barH > room Then barH = room
    If barH < 60 Then barH = 60
    r.Controls("boxBar").Move pad, y + 30 + (room - barH) \ 2, w - 2 * pad, barH
    r.Controls("txtCode").Move 0, 0, 60, 60
    r.Section(0).Height = h
    DoCmd.Close acReport, LABEL_REPORT, acSaveYes
    LabelReportOpen
    Exit Function
EH:
    ApplyLabelLayout = " ⁄–— ÷»ÿ  ﬁ—Ì— «·„·’ﬁ« : " & Err.Description & " (" & Err.Number & ")"
    If Err.Number = 5 Or Err.Number = 9 Then
        ApplyLabelLayout = "«·ÿ«»⁄… ´" & prn & "ª €Ì— „ÊÃÊœ… ⁄·Ï Â–« «·ÃÂ«“. «Œ —Â« „‰ ÃœÌœ ›Ì ≈⁄œ«œ«  «·„·’ﬁ« ."
    End If
    On Error Resume Next
    DoCmd.Close acReport, LABEL_REPORT, acSaveNo
End Function

Public Function PrintLabelQueue(ByVal Preview As Boolean) As String
    Dim msg As String
    If LabelTotalCopies() = 0 Then
        PrintLabelQueue = "√÷› ’‰›« Ê«Õœ« ⁄·Ï «·√ﬁ· ≈·Ï ﬁ«∆„… «·ÿ»«⁄…."
        Exit Function
    End If
    msg = ApplyLabelLayout()
    If Len(msg) > 0 Then
        PrintLabelQueue = msg
        Exit Function
    End If
    If Preview Then
        DoCmd.OpenReport LABEL_REPORT, acViewPreview
    Else
        DoCmd.OpenReport LABEL_REPORT, acViewNormal
        LogAction "PRINT_LABELS", "Products", "", "Labels=" & LabelTotalCopies()
    End If
End Function

'------------------------------------------------------------------------------
' Screen: frmBarcodeLabels (+ subform frmLabelLines)
'------------------------------------------------------------------------------
Public Sub LabelLinesOpen(ByVal frm As Access.Form)
    EnsureLabelTables
    frm.RecordSource = "SELECT * FROM tmpLabelQueue ORDER BY LineNo"
End Sub

Public Sub LabelsLoad(ByVal frm As Access.Form)
    Dim args As String
    Calendar = vbCalGreg
    EnsureLabelTables
    frm!txtCopies.Value = 1
    args = Nz(frm.OpenArgs, "")
    If Left$(args, 9) = "PURCHASE:" Then
        AddPurchaseLabels CLng(Mid$(args, 10))
    ElseIf args = "CART" Then
        AddCartLabels
    ElseIf Len(args) > 0 And IsNumeric(args) Then
        AddLabelProduct CLng(args), 1
    End If
    LabelsRefresh frm
    SafeFocus frm!txtBarcode
End Sub

Public Sub LabelsRefresh(ByVal frm As Access.Form)
    frm!subLines.Form.Requery
    UpdateLabelStatus frm
End Sub

Public Sub UpdateLabelStatus(ByVal frm As Access.Form)
    frm!lblStatus.Caption = "«·√’‰«›: " & LabelLineCount() & "    ⁄œœ «·„·’ﬁ« : " & LabelTotalCopies()
End Sub

Private Function LabelCopiesTyped(ByVal frm As Access.Form) As Long
    Dim v As Variant
    v = frm!txtCopies.Value
    If IsNumeric(v) Then LabelCopiesTyped = CLng(v)
    If LabelCopiesTyped < 1 Then LabelCopiesTyped = 1
End Function

Public Sub LabelBarcodeKeyDown(ByVal frm As Access.Form, ByRef KeyCode As Integer)
    Dim code As String, id As Variant
    If KeyCode <> vbKeyReturn Then Exit Sub
    KeyCode = 0
    code = Trim$(Nz(frm!txtBarcode.Value, ""))
    If Len(code) = 0 Then Exit Sub
    id = DLookup("ProductID", "Products", "Barcode = " & SqlText(code) & " OR ProductCode = " & SqlText(code))
    If IsNull(id) Then
        frm!lblStatus.Caption = "·« ÌÊÃœ „‰ Ã »«·»«—ﬂÊœ √Ê «·ﬂÊœ: " & code
        frm!txtBarcode.Value = Null
        SafeFocus frm!txtBarcode
        Exit Sub
    End If
    AddLabelProduct CLng(id), LabelCopiesTyped(frm)
    frm!txtBarcode.Value = Null
    LabelsRefresh frm
    SafeFocus frm!txtBarcode
End Sub

Public Sub LabelProductPicked(ByVal frm As Access.Form)
    If IsNull(frm!cboProduct.Value) Then Exit Sub
    AddLabelProduct CLng(frm!cboProduct.Value), LabelCopiesTyped(frm)
    frm!cboProduct.Value = Null
    LabelsRefresh frm
    SafeFocus frm!txtBarcode
End Sub

Public Sub LabelCopiesChanged(ByVal frm As Access.Form)
    ' subform: 1..500 labels per product
    Dim v As Variant
    v = frm!Copies.Value
    If Not IsNumeric(v) Then v = 1
    If CLng(v) < 1 Then v = 1
    If CLng(v) > MAX_COPIES Then v = MAX_COPIES
    frm!Copies.Value = CLng(v)
    If frm.Dirty Then frm.Dirty = False
    UpdateLabelStatus frm.Parent
End Sub

Public Sub RemoveLabelLine(ByVal frm As Access.Form)
    If IsNull(frm!LineNo.Value) Then Exit Sub
    CurrentDb.Execute "DELETE FROM tmpLabelQueue WHERE LineNo = " & frm!LineNo.Value, dbFailOnError
    frm.Requery
    UpdateLabelStatus frm.Parent
End Sub

Public Sub ClearLabels(ByVal frm As Access.Form)
    If LabelLineCount() = 0 Then Exit Sub
    If Not AskYesNo("Â·  —Ìœ „”Õ ﬁ«∆„… «·„·’ﬁ« ø") Then Exit Sub
    ClearLabelQueue
    LabelsRefresh frm
End Sub

Public Sub PrintLabels(ByVal frm As Access.Form, ByVal Preview As Boolean)
    Dim msg As String
    If frm!subLines.Form.Dirty Then frm!subLines.Form.Dirty = False
    msg = PrintLabelQueue(Preview)
    If Len(msg) > 0 Then ShowWarning msg
End Sub

Public Sub LabelsForPurchase(ByVal PurchaseInvoiceID As Variant, Optional ByVal FromForm As Variant)
    ' From a saved purchase invoice; a pop-up caller is closed so the label screen is not behind it.
    If IsNull(PurchaseInvoiceID) Or IsEmpty(PurchaseInvoiceID) Then
        ShowInfo "«Œ — ›« Ê—… ‘—«¡ √Ê·«."
        Exit Sub
    End If
    If Not IsMissing(FromForm) Then
        If FromForm.PopUp Then DoCmd.Close acForm, FromForm.Name
    End If
    OpenScreen "frmBarcodeLabels", 0, "PURCHASE:" & PurchaseInvoiceID
End Sub

Public Sub LabelsFromPurchaseScreen(ByVal frm As Access.Form)
    ' Purchase invoice screen: the lines being entered, or the last saved invoice.
    EnsureLocalTables
    If DCount("*", "tmpPurchaseLines") > 0 Then
        OpenScreen "frmBarcodeLabels", 0, "CART"
    ElseIf Not IsNull(TempVars("LastPurchaseID")) Then
        LabelsForPurchase TempVars("LastPurchaseID")
    Else
        ShowInfo "√÷› √’‰«› «·›« Ê—… √Ê «Õ›Ÿ ›« Ê—… √Ê·«° À„ «ÿ»⁄ «·»«—ﬂÊœ."
    End If
End Sub

'------------------------------------------------------------------------------
' Screen: frmLabelSettings (printer list)
'------------------------------------------------------------------------------
Public Sub LabelSettingsLoad(ByVal frm As Access.Form)
    Dim p As Variant, s As String
    On Error Resume Next
    For Each p In Application.Printers
        s = s & """" & Replace(p.DeviceName, """", "") & """;"
    Next
    frm!PrinterName.RowSourceType = "Value List"
    frm!PrinterName.ColumnCount = 1
    frm!PrinterName.ColumnWidths = ""
    frm!PrinterName.RowSource = s
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests)
'------------------------------------------------------------------------------
Public Function TestLabels() As Boolean
    Dim passed As Long, failed As Long, report As String, pid As Variant, msg As String
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestLabels  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    CheckLabel BarcodePattern("6281000000014") = "10100100110001001011001101001110001101000110101010111" & _
               "001011100101110010111001011001101011100101", "EAN-13: 6281000000014", passed, failed, report
    CheckLabel BarcodePattern("P00021") = "11010010000111011101101001110110010011101100100111011" & _
               "001100111001010011100110110010100001100011101011", "Code 128: P00021", passed, failed, report
    CheckLabel BarcodePattern("12345") = "1101001110010110011100100010110001011110111011011100100111010110001100011101011", _
               "Code 128 √—ﬁ«„: 12345", passed, failed, report
    CheckLabel BarcodePattern("6281000000015") = Code128Pattern("6281000000015"), _
               "—ﬁ„ EAN Œ«ÿ∆ Ìıÿ»⁄ Code 128", passed, failed, report
    CheckLabel Len(BarcodePattern("„‰ Ã")) = 0, "‰’ ⁄—»Ì ·« Ìı—„Û¯“", passed, failed, report
    CheckLabel LabelCode("„‰ Ã", "P00001") = "P00001", "»œÊ‰ »«—ﬂÊœ ’«·Õ Ìıÿ»⁄ ﬂÊœ «·„‰ Ã", passed, failed, report
    CheckLabel UnitsToCopies(2.5) = 3 And UnitsToCopies(0) = 1 And UnitsToCopies(10000) = MAX_COPIES, _
               "⁄œœ «·„·’ﬁ«  „‰ «·ﬂ„Ì…", passed, failed, report
    CheckLabel Not IsNull(LabelSetting("LabelWidth")), "ÃœÊ· ≈⁄œ«œ«  «·„·’ﬁ«  „ÊÃÊœ (BuildSchema)", _
               passed, failed, report

    pid = DMin("ProductID", "Products")
    If IsNull(pid) Then
        Debug.Print "[i] ·«  ÊÃœ „‰ Ã« :  ŒÿÌ «Œ »«— ﬁ«∆„… «·ÿ»«⁄…"
    ElseIf LabelLineCount() > 0 Then
        Debug.Print "[i] ﬁ«∆„… «·„·’ﬁ«  ·Ì”  ›«—€…:  ŒÿÌ «Œ »«— ﬁ«∆„… «·ÿ»«⁄…"
    Else
        AddLabelProduct pid, 2
        AddLabelProduct pid, 3
        CheckLabel LabelLineCount() = 1 And LabelTotalCopies() = 5, "≈÷«›… ‰›” «·’‰›  Ã„⁄ «·⁄œœ (5)", _
                   passed, failed, report
        msg = ApplyLabelLayout()
        CheckLabel Len(msg) = 0, "÷»ÿ  ﬁ—Ì— «·„·’ﬁ«  „‰ «·≈⁄œ«œ«  " & msg, passed, failed, report
        On Error Resume Next
        Err.Clear
        DoCmd.OpenReport LABEL_REPORT, acViewPreview, , , acHidden
        CheckLabel Err.Number = 0, "„⁄«Ì‰… «·„·’ﬁ«  " & Err.Description, passed, failed, report
        DoCmd.Close acReport, LABEL_REPORT, acSaveNo
        On Error GoTo 0
        ClearLabelQueue
        CheckLabel LabelLineCount() = 0, "„”Õ ﬁ«∆„… «·„·’ﬁ« ", passed, failed, report
    End If
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  „·’ﬁ«  «·»«—ﬂÊœ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestLabels"
        TestLabels = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
                "TestLabels"
    End If
End Function

Private Sub CheckLabel(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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
