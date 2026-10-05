Attribute VB_Name = "modCharts"
'==============================================================================
' modCharts  -  Retail Store Management System: statistics report with charts
'
'   rptStatistics draws three panels in the Print event of its detail section
'   (no chart control or add-in needed; prints the same on every PC):
'     1. bar chart   net sales of the 12 months that end with the report period
'                    (MonthlySalesQuery, VAT included, after returns)
'     2. pie chart   net sales by category within the period, with a legend
'                    (SalesByCategoryQuery)
'     3. top 10      best-selling products within the period, with bars
'                    (BestSellingProductsQuery)
'   The pure helpers (NiceMax, ChartMonths, ChartColor) are checked against
'   tools/charts_reference.py through LibreOffice.
'==============================================================================
Option Compare Database
Option Explicit

Private Const PI As Double = 3.14159265358979
Private Const TW As Single = 567                  ' twips per cm
Private Const CHART_FONT As String = "Tahoma"
Private Const CLR_GRID As Long = 15329769          ' #E9E9E9
Private Const CLR_PANEL As Long = 14079702         ' #D6D6D6 panel border

'------------------------------------------------------------------------------
' Pure helpers
'------------------------------------------------------------------------------
Public Function NiceMax(ByVal Value As Double) As Double
    ' Smallest 1, 2, 2.5 or 5 x 10^n that is >= Value (the top of the value axis).
    Dim p As Double, steps As Variant, i As Long
    If Value <= 0 Then
        NiceMax = 1
        Exit Function
    End If
    steps = Array(1, 2, 2.5, 5, 10)
    p = 10 ^ Int(Log(Value) / Log(10#))
    For i = 0 To 4
        If steps(i) * p >= Value * 0.999999 Then
            NiceMax = steps(i) * p
            Exit Function
        End If
    Next
    NiceMax = 10 * p
End Function

Public Function ChartMonths(ByVal EndDate As Date) As String
    ' "yyyymm,..." oldest first: the 12 months ending with the month of EndDate.
    Dim i As Long, d As Date, s As String
    For i = 11 To 0 Step -1
        d = DateSerial(Year(EndDate), Month(EndDate) - i, 1)
        s = s & IIf(Len(s) > 0, ",", "") & CStr(Year(d) * 100 + Month(d))
    Next
    ChartMonths = s
End Function

Public Function MonthShortName(ByVal MonthNo As Long) As String
    Dim names As Variant
    names = Split("Ì‰«Ì—,›»—«Ì—,„«—”,√»—Ì·,„«ÌÊ,ÌÊ‰ÌÊ,ÌÊ·ÌÊ,√€”ÿ”,”» „»—,√ﬂ Ê»—,‰Ê›„»—,œÌ”„»—", ",")
    MonthShortName = names(MonthNo - 1)
End Function

Public Function ChartColor(ByVal Index As Long) As Long
    ' Category colours (index from 0); the 9th and later categories share "other" grey.
    Select Case Index
        Case 0: ChartColor = RGB(30, 136, 229)
        Case 1: ChartColor = RGB(67, 160, 71)
        Case 2: ChartColor = RGB(251, 140, 0)
        Case 3: ChartColor = RGB(142, 36, 170)
        Case 4: ChartColor = RGB(229, 57, 53)
        Case 5: ChartColor = RGB(0, 137, 123)
        Case 6: ChartColor = RGB(216, 27, 96)
        Case 7: ChartColor = RGB(57, 73, 171)
        Case Else: ChartColor = RGB(158, 158, 158)
    End Select
End Function

'------------------------------------------------------------------------------
' Drawing helpers
'------------------------------------------------------------------------------
Private Sub PutText(ByVal rpt As Access.Report, ByVal Text As String, ByVal X As Single, ByVal Y As Single, _
                    ByVal Align As Integer, ByVal Size As Integer, ByVal Bold As Boolean, ByVal Color As Long)
    ' Align: 0 = text starts at X, 1 = centred on X, 2 = text ends at X (Arabic, right aligned)
    Dim w As Single
    rpt.FontName = CHART_FONT
    rpt.FontSize = Size
    rpt.FontBold = Bold
    rpt.ForeColor = Color
    w = rpt.TextWidth(Text)
    If Align = 1 Then
        rpt.CurrentX = X - w / 2
    ElseIf Align = 2 Then
        rpt.CurrentX = X - w
    Else
        rpt.CurrentX = X
    End If
    rpt.CurrentY = Y
    rpt.Print Text
End Sub

Private Sub Box(ByVal rpt As Access.Report, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, _
                ByVal Y2 As Single, ByVal Color As Long, ByVal Filled As Boolean)
    If Filled Then
        rpt.Line (X1, Y1)-(X2, Y2), Color, BF
    Else
        rpt.Line (X1, Y1)-(X2, Y2), Color, B
    End If
End Sub

Private Sub Panel(ByVal rpt As Access.Report, ByVal L As Single, ByVal T As Single, ByVal W As Single, _
                  ByVal H As Single, ByVal Title As String, ByVal Subtitle As String)
    Box rpt, L, T, L + W, T + H, CLR_PANEL, False
    PutText rpt, Title, L + W - 0.3 * TW, T + 0.15 * TW, 2, 12, True, CLR_PRIMARY
    If Len(Subtitle) > 0 Then PutText rpt, Subtitle, L + 0.3 * TW, T + 0.25 * TW, 0, 8, False, CLR_MUTED
End Sub

Private Function Amount(ByVal Value As Double) As String
    Amount = Format$(Value, "#,##0")
End Function

'------------------------------------------------------------------------------
' The report (Print event of the detail section)
'------------------------------------------------------------------------------
Public Sub DrawStatistics(ByVal rpt As Access.Report, ByVal L As Single, ByVal T As Single, _
                          ByVal W As Single, ByVal H As Single)
    Dim endDate As Date, h1 As Single, h2 As Single, gap As Single
    On Error GoTo EH
    Calendar = vbCalGreg
    endDate = Date
    If Not IsNull(TempVars("PeriodEnd")) Then endDate = CDate(TempVars("PeriodEnd")) - 1
    gap = 0.4 * TW
    h1 = H * 0.37
    h2 = H * 0.3
    DrawMonthlyBars rpt, L, T, W, h1, endDate
    DrawCategoryPie rpt, L, T + h1 + gap, W, h2
    DrawTopProducts rpt, L, T + h1 + h2 + 2 * gap, W, H - h1 - h2 - 2 * gap
    Exit Sub
EH:
    PutText rpt, " ⁄–— —”„ «·≈Õ’«∆Ì« : " & Err.Description, L + W - 0.3 * TW, T + 0.3 * TW, 2, 10, True, CLR_DANGER
End Sub

Private Sub DrawMonthlyBars(ByVal rpt As Access.Report, ByVal L As Single, ByVal T As Single, _
                            ByVal W As Single, ByVal H As Single, ByVal EndDate As Date)
    Dim months As Variant, vals(0 To 11) As Double, i As Long, ym As Long, top As Double, v As Double
    Dim plotL As Single, plotR As Single, plotT As Single, plotB As Single, slot As Single, x As Single
    Dim barH As Single, total As Double
    months = Split(ChartMonths(EndDate), ",")
    For i = 0 To 11
        ym = CLng(months(i))
        vals(i) = Nz(DbValue("SELECT NetSalesTotal FROM MonthlySalesQuery WHERE SalesYear = " & (ym \ 100) & _
                             " AND SalesMonth = " & (ym Mod 100)), 0)
        If vals(i) > top Then top = vals(i)
        total = total + vals(i)
    Next
    Panel rpt, L, T, W, H, "«·„»Ì⁄«  «·‘Â—Ì… (¬Œ— 12 ‘Â—«)", _
          "‘«„· «·÷—Ì»… »⁄œ «·„— Ã⁄«  - «·„Ã„Ê⁄ " & Amount(total)
    top = NiceMax(top)
    plotL = L + 0.4 * TW
    plotR = L + W - 1.8 * TW                         ' value axis on the right (Arabic layout)
    plotT = T + 1.1 * TW
    plotB = T + H - 0.9 * TW
    For i = 0 To 4                                   ' grid lines and values: 0, 25, 50, 75, 100 %
        x = plotB - (plotB - plotT) * i / 4
        rpt.Line (plotL, x)-(plotR, x), IIf(i = 0, CLR_MUTED, CLR_GRID)
        PutText rpt, Amount(top * i / 4), plotR + 0.15 * TW, x - 0.17 * TW, 0, 7, False, CLR_MUTED
    Next
    slot = (plotR - plotL) / 12
    For i = 0 To 11                                  ' oldest month on the right
        ym = CLng(months(i))
        x = plotR - (i + 1) * slot
        v = vals(i)
        If v > 0 Then
            barH = (plotB - plotT) * v / top
            Box rpt, x + slot * 0.18, plotB - barH, x + slot * 0.82, plotB, _
                IIf(i = 11, RGB(67, 160, 71), RGB(30, 136, 229)), True
            PutText rpt, Amount(v), x + slot / 2, plotB - barH - 0.38 * TW, 1, 6, False, CLR_TEXT
        End If
        PutText rpt, MonthShortName(ym Mod 100), x + slot / 2, plotB + 0.08 * TW, 1, 7, False, CLR_TEXT
        If i = 0 Or (ym Mod 100) = 1 Then
            PutText rpt, CStr(ym \ 100), x + slot / 2, plotB + 0.42 * TW, 1, 6, False, CLR_MUTED
        End If
    Next
End Sub

Private Sub DrawCategoryPie(ByVal rpt As Access.Report, ByVal L As Single, ByVal T As Single, _
                            ByVal W As Single, ByVal H As Single)
    Dim rs As DAO.Recordset, names(0 To 8) As String, vals(0 To 8) As Double, n As Long, total As Double
    Dim cx As Single, cy As Single, r As Single, a0 As Double, a1 As Double, i As Long, y As Single
    Dim legendR As Single
    Set rs = CurrentDb.OpenRecordset("SELECT CategoryName, CategoryTotal FROM SalesByCategoryQuery " & _
                                     "WHERE CategoryTotal > 0 ORDER BY CategoryTotal DESC", dbOpenSnapshot)
    Do Until rs.EOF
        If n < 8 Then
            names(n) = Nz(rs!CategoryName, "-")
            vals(n) = rs!CategoryTotal
            n = n + 1
        Else
            names(8) = "√Œ—Ï"
            vals(8) = vals(8) + rs!CategoryTotal
        End If
        total = total + rs!CategoryTotal
        rs.MoveNext
    Loop
    rs.Close
    If vals(8) > 0 Then n = 9
    Panel rpt, L, T, W, H, "«·„»Ì⁄«  Õ”» «· ’‰Ì›", "Œ·«· › —… «· ﬁ—Ì—° ‘«„· «·÷—Ì»…"
    If total <= 0 Then
        PutText rpt, "·«  ÊÃœ „»Ì⁄«  ›Ì Â–Â «·› —….", L + W / 2, T + H / 2, 1, 11, False, CLR_MUTED
        Exit Sub
    End If
    r = (H - 1.4 * TW) / 2
    cx = L + 0.8 * TW + r
    cy = T + 0.9 * TW + r
    rpt.FillStyle = 0                                 ' solid
    a0 = 0
    For i = 0 To n - 1
        a1 = a0 + 2 * PI * vals(i) / total
        rpt.FillColor = ChartColor(i)
        If vals(i) >= total * 0.9999 Then
            rpt.Circle (cx, cy), r, ChartColor(i)
        ElseIf a1 - a0 > 0.0001 Then
            ' negative angles draw the radii, so the slice is a filled sector
            If a1 > 2 * PI - 0.000001 Then a1 = 2 * PI - 0.000001
            rpt.Circle (cx, cy), r, RGB(255, 255, 255), -(a0 + 0.000001), -a1
        End If
        a0 = a1
    Next
    rpt.FillStyle = 1                                 ' transparent again
    legendR = L + W - 0.5 * TW                        ' legend on the right: colour, name, amount, share
    y = T + 1.0 * TW
    For i = 0 To n - 1
        Box rpt, legendR - 0.4 * TW, y + 0.05 * TW, legendR, y + 0.4 * TW, ChartColor(i), True
        PutText rpt, names(i), legendR - 0.6 * TW, y, 2, 9, False, CLR_TEXT
        PutText rpt, Amount(vals(i)) & "   (" & Format$(vals(i) / total, "0.0%") & ")", _
                cx + r + 1.0 * TW, y, 0, 9, True, CLR_TEXT
        y = y + 0.55 * TW
    Next
End Sub

Private Sub DrawTopProducts(ByVal rpt As Access.Report, ByVal L As Single, ByVal T As Single, _
                            ByVal W As Single, ByVal H As Single)
    Dim rs As DAO.Recordset, y As Single, rowH As Single, i As Long, top As Double, R As Single
    Dim xQty As Single, xSales As Single, barL As Single, barR As Single
    Panel rpt, L, T, W, H, "√›÷· 10 „‰ Ã«  „»Ì⁄«", "Õ”» ’«›Ì «·ﬂ„Ì… Œ·«· › —… «· ﬁ—Ì—"
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 10 ProductName, NetQty, SalesTotal FROM BestSellingProductsQuery " & _
                                     "WHERE NetQty > 0 ORDER BY NetQty DESC, SalesTotal DESC", dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        PutText rpt, "·«  ÊÃœ „»Ì⁄«  ›Ì Â–Â «·› —….", L + W / 2, T + H / 2, 1, 11, False, CLR_MUTED
        Exit Sub
    End If
    Do Until rs.EOF
        If Nz(rs!SalesTotal, 0) > top Then top = rs!SalesTotal
        rs.MoveNext
    Loop
    rs.MoveFirst
    R = L + W - 0.3 * TW
    xQty = L + W * 0.5
    xSales = L + W * 0.37
    barL = L + 0.4 * TW
    barR = L + W * 0.23                              ' the bars stay clear of the amounts
    rowH = (H - 1.6 * TW) / 11
    y = T + 0.9 * TW
    Box rpt, L + 0.2 * TW, y, L + W - 0.2 * TW, y + rowH, RGB(232, 236, 243), True
    PutText rpt, "#", R, y + 0.08 * TW, 2, 9, True, CLR_TEXT
    PutText rpt, "«·„‰ Ã", R - 0.8 * TW, y + 0.08 * TW, 2, 9, True, CLR_TEXT
    PutText rpt, "«·ﬂ„Ì…", xQty, y + 0.08 * TW, 2, 9, True, CLR_TEXT
    PutText rpt, "«·„»Ì⁄« ", xSales, y + 0.08 * TW, 2, 9, True, CLR_TEXT
    i = 0
    Do Until rs.EOF
        i = i + 1
        y = y + rowH
        If i Mod 2 = 0 Then Box rpt, L + 0.2 * TW, y, L + W - 0.2 * TW, y + rowH, RGB(247, 248, 250), True
        PutText rpt, CStr(i), R, y + 0.08 * TW, 2, 9, False, CLR_MUTED
        PutText rpt, Left$(Nz(rs!ProductName, ""), 40), R - 0.8 * TW, y + 0.08 * TW, 2, 9, False, CLR_TEXT
        PutText rpt, Format$(Nz(rs!NetQty, 0), "#,##0.###"), xQty, y + 0.08 * TW, 2, 9, False, CLR_TEXT
        PutText rpt, Format$(Nz(rs!SalesTotal, 0), "#,##0.00"), xSales, y + 0.08 * TW, 2, 9, True, CLR_TEXT
        If top > 0 And Nz(rs!SalesTotal, 0) > 0 Then
            Box rpt, barR - (barR - barL) * rs!SalesTotal / top, y + rowH * 0.25, barR, y + rowH * 0.75, _
                ChartColor(0), True
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub
