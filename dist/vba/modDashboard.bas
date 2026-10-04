Attribute VB_Name = "modDashboard"
'==============================================================================
' modDashboard  -  Retail Store Management System (Phase 9)
'
' The dashboard of frmMain:
'   8 tiles    today / month sales, month net profit, low-stock products,
'              customer debts, supplier dues, stock value, month expenses
'   3 lists    latest sales invoices, low-stock products, best sellers this month
' Figures come from DashboardQuery and qryDashboardTopProducts (own parameters
' DashDay / DashMonth / DashEnd, so the report-centre period is never changed);
' the month profit comes from ProfitQuery with the report period saved and restored.
'
' Without the DASHBOARD_FINANCIAL permission (e.g. a cashier) the amounts are
' hidden; counts stay visible and the invoice list shows the user's own invoices.
' Refresh: on opening, on returning to the main screen (at most every 20 s),
' every 5 minutes (form timer) and with the Refresh button.
'==============================================================================
Option Compare Database
Option Explicit

Private Const HIDDEN_AMOUNT As String = "ïïïï"
Private m_lastRefresh As Date

Public Sub DashboardRefresh(ByVal frm As Access.Form)
    Dim rs As DAO.Recordset, financial As Boolean, lowCount As Long
    On Error GoTo EH
    Calendar = vbCalGreg
    SetDashboardParams Date
    financial = HasPermission("DASHBOARD_FINANCIAL")
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM DashboardQuery", dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If
    lowCount = Nz(rs!LowStockCount, 0)
    SetTile frm, 1, Money(rs!TodaySales, financial), rs!TodayInvoices & " ›« Ê—… «·ÌÊ„"
    SetTile frm, 2, Money(rs!MonthSales, financial), _
            IIf(financial, "„‰Â« ÷—Ì»… " & Format$(Nz(rs!MonthVAT, 0), "#,##0.00"), rs!MonthInvoices & " ›« Ê—…")
    If financial Then
        SetTile frm, 3, Money(MonthNetProfit(), True), "»⁄œ «·„’—Ê›«  Ê›—Êﬁ«  «·„Œ“Ê‰"
    Else
        SetTile frm, 3, HIDDEN_AMOUNT, "Ì ÿ·» ’·«ÕÌ… «·√—ﬁ«„ «·„«·Ì…"
    End If
    SetTile frm, 4, CStr(lowCount), IIf(lowCount > 0, "«÷€ÿ ·⁄—÷ «· ﬁ—Ì—", "·«  ÊÃœ ‰Ê«ﬁ’")
    frm!lblTileValue4.ForeColor = IIf(lowCount > 0, CLR_DANGER, CLR_SUCCESS)
    SetTile frm, 5, Money(rs!CustomerDebt, financial), rs!DebtorCount & " ⁄„Ì· ⁄·ÌÂ —’Ìœ"
    SetTile frm, 6, Money(rs!SupplierDue, financial), "„” Õﬁ ··„Ê—œÌ‰"
    SetTile frm, 7, Money(rs!StockValue, financial), "»„ Ê”ÿ «· ﬂ·›…"
    SetTile frm, 8, Money(rs!MonthExpenses, financial), "»œÊ‰ «·÷—Ì»…"
    rs.Close

    frm!lstRecentSales.RowSource = "SELECT TOP 15 h.SalesInvoiceID, h.InvoiceNumber AS [«·—ﬁ„], " & _
        "Format(h.InvoiceDate, 'hh:nn') AS [«·Êﬁ ], c.CustomerName AS [«·⁄„Ì·], " & _
        IIf(financial, "h.TotalAmount", "'-'") & " AS [«·≈Ã„«·Ì] " & _
        "FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID " & _
        "WHERE h.InvoiceDate >= " & SqlDate(Date) & _
        IIf(financial, "", " AND h.EmployeeID = " & CurrentUserID()) & _
        " ORDER BY h.SalesInvoiceID DESC"
    frm!lstLowStock.RowSource = "SELECT TOP 30 ProductID, ProductName AS [«·„‰ Ã], CurrentQuantity AS [«·„ Ê›—], " & _
        "MinimumQuantity AS [«·Õœ] FROM LowStockQuery ORDER BY ShortageQty DESC, ProductName"
    frm!lstTopProducts.RowSource = "SELECT TOP 10 ProductID, ProductName AS [«·„‰ Ã], NetQty AS [«·ﬂ„Ì…], " & _
        IIf(financial, "NetSales", "'-'") & " AS [«·„»Ì⁄« ] FROM qryDashboardTopProducts " & _
        "ORDER BY NetQty DESC, ProductName"
    frm!lblUpdated.Caption = "¬Œ—  ÕœÌÀ: " & Format$(Now, "hh:nn")
    m_lastRefresh = Now
    RefreshIntegrityStatus frm
    Exit Sub
EH:
    frm!lblUpdated.Caption = " ⁄–—  ÕœÌÀ «·„ƒ‘—« : " & Err.Description
End Sub

Public Sub DashboardActivate(ByVal frm As Access.Form)
    ' Returning from another screen (a sale, a purchase...) shows fresh figures.
    If DateDiff("s", m_lastRefresh, Now) >= 20 Then DashboardRefresh frm
End Sub

Public Sub DashboardTileClick(ByVal TileKey As String)
    ' Each tile opens the report behind its figure.
    Dim firstDay As Date, where As String
    Calendar = vbCalGreg
    firstDay = DateSerial(Year(Date), Month(Date), 1)
    If TileKey <> "LOW" And Not HasPermission("REPORTS") Then
        ShowWarning "·«  „·ﬂ ’·«ÕÌ… «· ﬁ«—Ì—."
        Exit Sub
    End If
    Select Case TileKey
        Case "TODAY", "MONTH"
            If TileKey = "TODAY" Then firstDay = Date
            where = "[SaleDate] >= " & SqlDate(firstDay) & " AND [SaleDate] < " & SqlDate(Date + 1)
            OpenReportOrQuery "rptDailySales", "DailySalesQuery", where, PeriodText(firstDay, Date)
        Case "PROFIT", "EXPENSES"
            If Not HasPermission("REPORTS_PROFIT") And TileKey = "PROFIT" Then
                ShowWarning "·«  „·ﬂ ’·«ÕÌ…  ﬁ«—Ì— «·√—»«Õ."
                Exit Sub
            End If
            SetPeriod firstDay, Date
            If TileKey = "PROFIT" Then
                OpenReportOrQuery "rptProfit", "ProfitQuery", "", PeriodText(firstDay, Date)
            Else
                OpenReportOrQuery "rptExpenses", "ExpensesQuery", "", PeriodText(firstDay, Date)
            End If
        Case "LOW":   OpenReportOrQuery "rptLowStock", "LowStockQuery", ""
        Case "DEBT":  OpenReportOrQuery "rptCustomerBalances", "CustomerBalanceQuery", "[Balance] > 0"
        Case "DUE":   OpenReportOrQuery "rptSupplierBalances", "SupplierBalanceQuery", "[Balance] > 0"
        Case "STOCK": OpenReportOrQuery "rptStockBalance", "StockBalanceQuery", "[IsActive] = True"
    End Select
End Sub

Public Function LowStockAlert() As Boolean
    ' Once per session, after the application opens: offers the low-stock report.
    Dim n As Long
    On Error Resume Next
    If TempVars("LowStockAlertShown") = True Then Exit Function
    TempVars.Add "LowStockAlertShown", True
    n = Nz(DCount("*", "LowStockQuery"), 0)
    On Error GoTo 0
    If n = 0 Then Exit Function
    LowStockAlert = True
    If AskYesNo(" ‰»ÌÂ «·„Œ“Ê‰: ÌÊÃœ " & n & " „‰ Ã Ê’·  ﬂ„Ì Â ≈·Ï Õœ ≈⁄«œ… «·ÿ·» √Ê √ﬁ·." & vbCrLf & _
                "Â·  —Ìœ ⁄—÷  ﬁ—Ì— «·„‰ Ã«  „‰Œ›÷… «·„Œ“Ê‰ø") Then
        OpenReportOrQuery "rptLowStock", "LowStockQuery", ""
    End If
End Function

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Public Sub SetDashboardParams(ByVal Today As Date)
    TempVars.Add "DashDay", DateValue(Today)
    TempVars.Add "DashMonth", DateSerial(Year(Today), Month(Today), 1)
    TempVars.Add "DashEnd", DateValue(Today) + 1
End Sub

Private Function MonthNetProfit() As Currency
    ' ProfitQuery reads the report period: use this month, then put the user's period back.
    Dim savedStart As Variant, savedEnd As Variant
    savedStart = TempVars("PeriodStart")
    savedEnd = TempVars("PeriodEnd")
    SetPeriod DateSerial(Year(Date), Month(Date), 1), Date
    MonthNetProfit = Nz(DbValue("SELECT NetProfit FROM ProfitQuery"), 0)
    If IsNull(savedStart) Then
        TempVars.Remove "PeriodStart"
        TempVars.Remove "PeriodEnd"
    Else
        TempVars.Add "PeriodStart", savedStart
        TempVars.Add "PeriodEnd", savedEnd
    End If
End Function

Private Sub SetTile(ByVal frm As Access.Form, ByVal Index As Integer, ByVal ValueText As String, _
                    ByVal SubText As String)
    frm.Controls("lblTileValue" & Index).Caption = ValueText
    frm.Controls("lblTileSub" & Index).Caption = IIf(Len(SubText) = 0, " ", SubText)
End Sub

Private Function Money(ByVal Value As Variant, ByVal Visible As Boolean) As String
    If Visible Then Money = Format$(Nz(Value, 0), "#,##0.00") Else Money = HIDDEN_AMOUNT
End Function

'------------------------------------------------------------------------------
' In-Access test (Immediate window: TestDashboard). Reads only; changes nothing.
'------------------------------------------------------------------------------
Public Function TestDashboard() As Boolean
    Dim frm As Access.Form, opened As Boolean, i As Integer, passed As Long, failed As Long
    Dim report As String, cap As String, ok As Boolean, savedStart As Variant, savedEnd As Variant
    On Error GoTo EH
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    savedStart = TempVars("PeriodStart")
    savedEnd = TempVars("PeriodEnd")
    SetPeriod DateSerial(2020, 1, 1), DateSerial(2020, 1, 31)
    If IsFormOpen("frmMain") Then
        Set frm = Forms("frmMain")
        DashboardRefresh frm
    Else
        DoCmd.OpenForm "frmMain", acNormal, , , , acHidden
        opened = True
        Set frm = Forms("frmMain")
    End If
    For i = 1 To 8
        cap = Replace(frm.Controls("lblTileValue" & i).Caption, ",", "")
        ok = IsNumeric(cap) Or cap = HIDDEN_AMOUNT
        CheckResult ok, "«·»ÿ«ﬁ… " & i & "  ⁄—÷ —ﬁ„« (" & cap & ")", passed, failed, report
    Next
    CheckResult CLng(frm!lblTileValue4.Caption) = DCount("*", "LowStockQuery"), _
                "⁄œœ «·„‰ Ã«  „‰Œ›÷… «·„Œ“Ê‰ =  ﬁ—Ì— «·‰Ê«ﬁ’", passed, failed, report
    CheckResult Left$(frm!lblUpdated.Caption, 9) = "¬Œ—  ÕœÌÀ", "Êﬁ  ¬Œ—  ÕœÌÀ", passed, failed, report
    CheckResult DbValue("SELECT TodayInvoices FROM DashboardQuery") = _
                DCount("*", "SalesInvoices", "InvoiceDate >= " & SqlDate(Date)), _
                "⁄œœ ›Ê« Ì— «·ÌÊ„", passed, failed, report
    CheckResult TempVars("PeriodStart") = DateSerial(2020, 1, 1) And TempVars("PeriodEnd") = DateSerial(2020, 2, 1), _
                "·ÊÕ… «· Õﬂ„ ·«  €Ì¯— › —… „—ﬂ“ «· ﬁ«—Ì—", passed, failed, report
    CheckResult frm!lstLowStock.ListCount - 1 = DCount("*", "LowStockQuery") Or DCount("*", "LowStockQuery") > 30, _
                "ﬁ«∆„… «·‰Ê«ﬁ’", passed, failed, report
    If opened Then DoCmd.Close acForm, "frmMain", acSaveNo
    GoSub Restore
    g_SilentMode = False
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  ·ÊÕ… «· Õﬂ„ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestDashboard"
        TestDashboard = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & report, vbExclamation + MSG_RTL, "TestDashboard"
    End If
    Exit Function
Restore:
    If IsNull(savedStart) Then
        TempVars.Remove "PeriodStart"
        TempVars.Remove "PeriodEnd"
    Else
        TempVars.Add "PeriodStart", savedStart
        TempVars.Add "PeriodEnd", savedEnd
    End If
    Return
EH:
    g_SilentMode = False
    TestMsg "Œÿ√ " & Err.Number & ": " & Err.Description, vbCritical + MSG_RTL, "TestDashboard"
End Function

Private Sub CheckResult(ByVal Passed As Boolean, ByVal Label As String, ByRef PassedCount As Long, _
                        ByRef FailedCount As Long, ByRef Report As String)
    If Passed Then
        PassedCount = PassedCount + 1
        Debug.Print "[OK] " & Label
    Else
        FailedCount = FailedCount + 1
        Report = Report & "- " & Label & vbCrLf
        Debug.Print "[X] " & Label
    End If
End Sub
