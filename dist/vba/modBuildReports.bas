Attribute VB_Name = "modBuildReports"
'==============================================================================
' modBuildReports  -  Retail Store Management System (Phases 6 and 8: reports)
'
' GENERATED FILE - do not edit by hand.
' Source: tools/reports.py, tools/reports_docs.py, tools/reports_catalog.py
'
'   BuildReports   (re)creates every report:
'                  - sales invoice / credit note: rptSalesReceipt (80 mm) and
'                    rptSalesInvoiceA4, with the ZATCA QR code (DrawDocumentQR);
'                  - documents: purchase invoice / return, vouchers, stocktake sheet;
'                  - the report centre catalogue (one report per saved query).
'   TestReports    opens every report in preview (hidden) and checks it.
' Text functions used by the reports are in modReports.
'==============================================================================
Option Compare Database
Option Explicit

Private Const MIRROR_LAYOUT As Boolean = False      ' same switch as modBuildForms
Private Const EP As String = "[Event Procedure]"

Private m_rpt As Access.Report
Private m_tmp As String
Private m_width As Long
Private m_built As Long
Private m_failed As Long
Private m_report As String
Private m_passed As Long
Private Const REPORT_NAMES As String = "rptSalesReceipt,rptSalesInvoiceA4,rptPurchaseDocument,rptVoucher,rptStockCount,rptBarcodeLabels,rptStatistics,rptCashVoucher,rptCashClosing,rptJournalEntry,rptAccountStatement,rptGeneralLedger,rptIncomeStatement,rptBalanceSheet,rptVatReturn,rptAging,rptPayroll,rptDailySales,rptMonthlySales,rptSalesByPeriod,rptSalesByProduct,rptBestSelling,rptLeastSelling,rptPurchases,rptStockBalance,rptLowStock,rptProductMovement,rptCustomerStatement,rptSupplierStatement,rptExpenses,rptExpensesByType,rptCashStatement,rptCashDaily,rptCostCenterProfit,rptCostCenterAccounts,rptBudgetVsActual,rptAuditTrail,rptFixedAssets,rptCashBalances,rptCashClosings,rptJournal,rptTrialBalance,rptTrialBalanceTree,rptAccountTree,rptSlowMoving,rptStockByCategory,rptCustomerBalances,rptSupplierBalances,rptIntegrityCheck,rptProfit,rptVatSummary"

Public Function BuildReports() As Boolean
    Dim i As Long
    On Error GoTo EH
    m_built = 0: m_failed = 0: m_report = ""
    For i = Reports.Count - 1 To 0 Step -1
        DoCmd.Close acReport, Reports(i).Name, acSaveNo
    Next
    DoCmd.Echo False, "Ã«—Ì »‰«¡ «· ﬁ«—Ì—..."
    BuildReport_rptSalesReceipt
    BuildReport_rptSalesInvoiceA4
    BuildReport_rptPurchaseDocument
    BuildReport_rptVoucher
    BuildReport_rptStockCount
    BuildReport_rptBarcodeLabels
    BuildReport_rptStatistics
    BuildReport_rptCashVoucher
    BuildReport_rptCashClosing
    BuildReport_rptJournalEntry
    BuildReport_rptAccountStatement
    BuildReport_rptGeneralLedger
    BuildReport_rptIncomeStatement
    BuildReport_rptBalanceSheet
    BuildReport_rptVatReturn
    BuildReport_rptAging
    BuildReport_rptPayroll
    BuildReport_rptDailySales
    BuildReport_rptMonthlySales
    BuildReport_rptSalesByPeriod
    BuildReport_rptSalesByProduct
    BuildReport_rptBestSelling
    BuildReport_rptLeastSelling
    BuildReport_rptPurchases
    BuildReport_rptStockBalance
    BuildReport_rptLowStock
    BuildReport_rptProductMovement
    BuildReport_rptCustomerStatement
    BuildReport_rptSupplierStatement
    BuildReport_rptExpenses
    BuildReport_rptExpensesByType
    BuildReport_rptCashStatement
    BuildReport_rptCashDaily
    BuildReport_rptCostCenterProfit
    BuildReport_rptCostCenterAccounts
    BuildReport_rptBudgetVsActual
    BuildReport_rptAuditTrail
    BuildReport_rptFixedAssets
    BuildReport_rptCashBalances
    BuildReport_rptCashClosings
    BuildReport_rptJournal
    BuildReport_rptTrialBalance
    BuildReport_rptTrialBalanceTree
    BuildReport_rptAccountTree
    BuildReport_rptSlowMoving
    BuildReport_rptStockByCategory
    BuildReport_rptCustomerBalances
    BuildReport_rptSupplierBalances
    BuildReport_rptIntegrityCheck
    BuildReport_rptProfit
    BuildReport_rptVatSummary
    DoCmd.Echo True
    Application.RefreshDatabaseWindow
    If m_failed = 0 Then
        MsgBox " „ »‰«¡ «· ﬁ«—Ì— »‰Ã«Õ (" & m_built & ").", vbInformation + MSG_RTL, "BuildReports"
        BuildReports = True
    Else
        MsgBox "›‘· »‰«¡ " & m_failed & "  ﬁ—Ì—:" & vbCrLf & m_report, vbExclamation + MSG_RTL, "BuildReports"
    End If
    Exit Function
EH:
    DoCmd.Echo True
    MsgBox "Œÿ√ " & Err.Number & ": " & Err.Description, vbCritical + MSG_RTL, "BuildReports"
End Function

Public Function TestReports() As Boolean
    ' Opens every report hidden in preview (with this year's period and the first customer,
    ' supplier and product as parameters) and checks the amount-in-words function.
    Dim f As Variant
    m_passed = 0: m_failed = 0: m_report = ""
    Debug.Print "=== TestReports  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    SetPeriod DateSerial(Year(Date), 1, 1), Date
    SetQueryParam "CustomerID", Nz(DMin("CustomerID", "Customers"), 0)
    SetQueryParam "SupplierID", Nz(DMin("SupplierID", "Suppliers"), 0)
    SetQueryParam "ProductID", Nz(DMin("ProductID", "Products"), 0)
    For Each f In Split(REPORT_NAMES, ",")
        CheckReportOpens CStr(f)
    Next
    RecordR AmountInWords(0.5) = "›ﬁÿ Œ„”Ê‰ Â··… ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 0.5"
    RecordR AmountInWords(1) = "›ﬁÿ Ê«Õœ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 1"
    RecordR AmountInWords(2) = "›ﬁÿ «À‰«‰ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 2"
    RecordR AmountInWords(3) = "›ﬁÿ À·«À… —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 3"
    RecordR AmountInWords(10) = "›ﬁÿ ⁄‘—… —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 10"
    RecordR AmountInWords(11) = "›ﬁÿ √Õœ ⁄‘— —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 11"
    RecordR AmountInWords(12) = "›ﬁÿ «À‰« ⁄‘— —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 12"
    RecordR AmountInWords(19) = "›ﬁÿ  ”⁄… ⁄‘— —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 19"
    RecordR AmountInWords(20) = "›ﬁÿ ⁄‘—Ê‰ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 20"
    RecordR AmountInWords(21) = "›ﬁÿ Ê«Õœ Ê⁄‘—Ê‰ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 21"
    RecordR AmountInWords(99) = "›ﬁÿ  ”⁄… Ê ”⁄Ê‰ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 99"
    RecordR AmountInWords(100) = "›ﬁÿ „«∆… —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 100"
    RecordR AmountInWords(101) = "›ﬁÿ „«∆… ÊÊ«Õœ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 101"
    RecordR AmountInWords(115) = "›ﬁÿ „«∆… ÊŒ„”… ⁄‘— —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 115"
    RecordR AmountInWords(200) = "›ﬁÿ „«∆ «‰ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 200"
    RecordR AmountInWords(999) = "›ﬁÿ  ”⁄„«∆… Ê ”⁄… Ê ”⁄Ê‰ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 999"
    RecordR AmountInWords(1000) = "›ﬁÿ √·› —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 1000"
    RecordR AmountInWords(1001) = "›ﬁÿ √·› ÊÊ«Õœ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 1001"
    RecordR AmountInWords(1250.5) = "›ﬁÿ √·› Ê„«∆ «‰ ÊŒ„”Ê‰ —Ì«· ”⁄ÊœÌ ÊŒ„”Ê‰ Â··… ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 1250.5"
    RecordR AmountInWords(2000) = "›ﬁÿ √·›«‰ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 2000"
    RecordR AmountInWords(2500) = "›ﬁÿ √·›«‰ ÊŒ„”„«∆… —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 2500"
    RecordR AmountInWords(3000) = "›ﬁÿ À·«À… ¬·«› —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 3000"
    RecordR AmountInWords(10000) = "›ﬁÿ ⁄‘—… ¬·«› —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 10000"
    RecordR AmountInWords(11000) = "›ﬁÿ √Õœ ⁄‘— √·› —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 11000"
    RecordR AmountInWords(12345.67) = "›ﬁÿ «À‰« ⁄‘— √·› ÊÀ·«À„«∆… ÊŒ„”… Ê√—»⁄Ê‰ —Ì«· ”⁄ÊœÌ Ê”»⁄… Ê” Ê‰ Â··… ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 12345.67"
    RecordR AmountInWords(100000) = "›ﬁÿ „«∆… √·› —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 100000"
    RecordR AmountInWords(101000) = "›ﬁÿ „«∆… ÊÊ«Õœ √·› —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 101000"
    RecordR AmountInWords(999999.99) = "›ﬁÿ  ”⁄„«∆… Ê ”⁄… Ê ”⁄Ê‰ √·› Ê ”⁄„«∆… Ê ”⁄… Ê ”⁄Ê‰ —Ì«· ”⁄ÊœÌ Ê ”⁄… Ê ”⁄Ê‰ Â··… ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 999999.99"
    RecordR AmountInWords(1000000) = "›ﬁÿ „·ÌÊ‰ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 1000000"
    RecordR AmountInWords(2000000) = "›ﬁÿ „·ÌÊ‰«‰ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 2000000"
    RecordR AmountInWords(3500000) = "›ﬁÿ À·«À… „·«ÌÌ‰ ÊŒ„”„«∆… √·› —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 3500000"
    RecordR AmountInWords(11000000) = "›ﬁÿ √Õœ ⁄‘— „·ÌÊ‰ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 11000000"
    RecordR AmountInWords(123456789.01) = "›ﬁÿ „«∆… ÊÀ·«À… Ê⁄‘—Ê‰ „·ÌÊ‰ Ê√—»⁄„«∆… Ê” … ÊŒ„”Ê‰ √·› Ê”»⁄„«∆… Ê ”⁄… ÊÀ„«‰Ê‰ —Ì«· ”⁄ÊœÌ ÊÊ«Õœ Â··… ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 123456789.01"
    RecordR AmountInWords(0.05) = "›ﬁÿ Œ„”… Â··… ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 0.05"
    RecordR AmountInWords(0.11) = "›ﬁÿ √Õœ ⁄‘— Â··… ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 0.11"
    RecordR AmountInWords(4750) = "›ﬁÿ √—»⁄… ¬·«› Ê”»⁄„«∆… ÊŒ„”Ê‰ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 4750"
    RecordR AmountInWords(3175) = "›ﬁÿ À·«À… ¬·«› Ê„«∆… ÊŒ„”… Ê”»⁄Ê‰ —Ì«· ”⁄ÊœÌ ·« €Ì—", "«·„»·€ »«·Õ—Ê›: 3175"
    TempVars.Add "ReportCriteria", ""
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & m_passed & " | ›‘·: " & m_failed
    If m_failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «· ﬁ«—Ì— ‰«ÃÕ… (" & m_passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestReports"
        TestReports = True
    Else
        TestMsg "‰ÃÕ " & m_passed & " Ê›‘· " & m_failed & ":" & vbCrLf & vbCrLf & Left$(m_report, 900), _
               vbExclamation + MSG_RTL, "TestReports"
    End If
End Function

Private Sub CheckReportOpens(ByVal ReportName As String)
    On Error GoTo EH
    If Not ReportExists(ReportName) Then
        RecordR False, "«· ﬁ—Ì— €Ì— „ÊÃÊœ: " & ReportName & " (‘€¯· BuildReports)"
        Exit Sub
    End If
    TempVars.Add "ReportCriteria", "TEST"
    DoCmd.OpenReport ReportName, acViewPreview, , , acHidden
    DoCmd.Close acReport, ReportName, acSaveNo
    RecordR True, "«· ﬁ—Ì— " & ReportName & " Ì› Õ"
    Exit Sub
EH:
    If Err.Number = 2501 Then              ' cancelled by Report_NoData: no data yet, not an error
        RecordR True, "«· ﬁ—Ì— " & ReportName & " Ì› Õ (·«  ÊÃœ »Ì«‰«  »⁄œ)"
    Else
        RecordR False, ReportName & ": Œÿ√ " & Err.Number & " - " & Err.Description
    End If
    On Error Resume Next
    DoCmd.Close acReport, ReportName, acSaveNo
End Sub

Private Sub RecordR(ByVal Passed As Boolean, ByVal Label As String)
    If Passed Then
        m_passed = m_passed + 1
        Debug.Print "[OK] " & Label
    Else
        m_failed = m_failed + 1
        m_report = m_report & "- " & Label & vbCrLf
        Debug.Print "[X] " & Label
    End If
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Private Sub StartReport(ByVal FinalName As String, ByVal Caption As String, ByVal RecordSource As String, _
                        ByVal ReportWidth As Long, ByVal GroupField As String, ByVal SortFields As String, _
                        ByVal Landscape As Boolean, ByVal PageSetup As Boolean)
    ' GroupField: document reports group on it (header 5 = secHeader, footer 6 = secTotals).
    ' SortFields: "Field1,-Field2" (a leading "-" sorts descending). Reports ignore ORDER BY.
    Dim f As Variant, level As Long, fieldName As String
    If ReportExists(FinalName) Then DoCmd.DeleteObject acReport, FinalName
    Set m_rpt = CreateReport()
    m_tmp = m_rpt.Name
    m_width = ReportWidth
    SetRptProp "Orientation", 1                     ' right-to-left
    m_rpt.RecordSource = RecordSource
    m_rpt.Caption = Caption
    If Len(GroupField) > 0 Then
        CreateGroupLevel m_tmp, GroupField, True, True
        level = 1
    End If
    If Len(SortFields) > 0 Then
        For Each f In Split(SortFields, ",")
            fieldName = CStr(f)
            If Left$(fieldName, 1) = "-" Then fieldName = Mid$(fieldName, 2)
            CreateGroupLevel m_tmp, fieldName, False, False
            If Left$(CStr(f), 1) = "-" Then m_rpt.GroupLevel(level).SortOrder = True
            level = level + 1
        Next
    End If
    m_rpt.Width = ReportWidth
    If Len(GroupField) > 0 Then
        m_rpt.Section(5).Name = "secHeader"
        m_rpt.Section(6).Name = "secTotals"
        m_rpt.Section(6).KeepTogether = True
    End If
    If PageSetup Then SetPage Landscape
    m_rpt.HasModule = True
End Sub

Private Sub SetSection(ByVal SectionIndex As Integer, ByVal SectionHeight As Long)
    EnsureSection SectionIndex
    m_rpt.Section(SectionIndex).Height = SectionHeight
    m_rpt.Section(SectionIndex).Visible = True
End Sub

Private Sub EnsureSection(ByVal SectionIndex As Integer)
    ' Page and report header/footer sections come in pairs and are switched on with RunCommand.
    Dim h As Long, missing As Boolean
    On Error Resume Next
    h = m_rpt.Section(SectionIndex).Height
    missing = (Err.Number <> 0)
    On Error GoTo 0
    If Not missing Then Exit Sub
    DoCmd.SelectObject acReport, m_tmp
    If SectionIndex = acPageHeader Or SectionIndex = acPageFooter Then
        DoCmd.RunCommand acCmdPageHdrFtr
    ElseIf SectionIndex = acHeader Or SectionIndex = acFooter Then
        DoCmd.RunCommand acCmdReportHdrFtr
    End If
End Sub

Private Sub SetPage(ByVal Landscape As Boolean)
    ' A4 with 0.8 cm side margins. Needs a printer driver; without one the layout still works.
    On Error Resume Next
    m_rpt.Printer.PaperSize = 9                     ' A4
    If Landscape Then m_rpt.Printer.Orientation = 2 Else m_rpt.Printer.Orientation = 1
    m_rpt.Printer.LeftMargin = 454
    m_rpt.Printer.RightMargin = 454
    m_rpt.Printer.TopMargin = 567
    m_rpt.Printer.BottomMargin = 567
End Sub

Private Sub HideSection(ByVal SectionIndex As Integer)
    On Error Resume Next                             ' the section may not exist
    m_rpt.Section(SectionIndex).Height = 0
    m_rpt.Section(SectionIndex).Visible = False
End Sub

Private Sub SetSecProp(ByVal SectionIndex As Integer, ByVal PropName As String, ByVal Value As Variant)
    On Error Resume Next
    m_rpt.Section(SectionIndex).Properties(PropName).Value = Value
End Sub

Private Function NewRptCtl(ByVal CtlType As AcControlType, ByVal SectionIndex As Integer, _
                           ByVal CtlName As String, ByVal L As Long, ByVal T As Long, ByVal W As Long, _
                           ByVal H As Long, Optional ByVal ColumnName As String = "") As Access.Control
    Dim x As Long
    If MIRROR_LAYOUT Then x = m_width - L - W Else x = L
    Set NewRptCtl = CreateReportControl(m_tmp, CtlType, SectionIndex, "", ColumnName, x, T, W, H)
    NewRptCtl.Name = CtlName
End Function

Private Function RText(ByVal SectionIndex As Integer, ByVal CtlName As String, ByVal Source As String, _
                       ByVal L As Long, ByVal T As Long, ByVal W As Long, ByVal H As Long, _
                       ByVal FontSize As Integer, ByVal Bold As Boolean, ByVal TextAlign As Integer) As Access.Control
    Dim c As Access.Control
    If Left$(Source, 1) = "=" Then
        Set c = NewRptCtl(acTextBox, SectionIndex, CtlName, L, T, W, H)
        c.ControlSource = Source
    Else
        Set c = NewRptCtl(acTextBox, SectionIndex, CtlName, L, T, W, H, Source)
    End If
    c.FontName = FONT_NAME
    c.FontSize = FontSize
    c.FontBold = Bold
    c.TextAlign = TextAlign
    c.BorderStyle = 0
    c.BackStyle = 0
    c.ForeColor = 0
    Set RText = c
End Function

Private Function RLabel(ByVal SectionIndex As Integer, ByVal CtlName As String, ByVal Caption As String, _
                        ByVal L As Long, ByVal T As Long, ByVal W As Long, ByVal H As Long, _
                        ByVal FontSize As Integer, ByVal Bold As Boolean, ByVal TextAlign As Integer) As Access.Control
    Dim c As Access.Control
    Set c = NewRptCtl(acLabel, SectionIndex, CtlName, L, T, W, H)
    c.Caption = Caption
    c.FontName = FONT_NAME
    c.FontSize = FontSize
    c.FontBold = Bold
    c.TextAlign = TextAlign
    c.ForeColor = 0
    Set RLabel = c
End Function

Private Function RLine(ByVal SectionIndex As Integer, ByVal CtlName As String, ByVal T As Long, _
                       ByVal W As Long) As Access.Control
    Set RLine = NewRptCtl(acLine, SectionIndex, CtlName, 0, T, W, 0)
End Function

Private Function RBox(ByVal SectionIndex As Integer, ByVal CtlName As String, ByVal L As Long, _
                      ByVal T As Long, ByVal W As Long, ByVal H As Long) As Access.Control
    Dim c As Access.Control
    Set c = NewRptCtl(acRectangle, SectionIndex, CtlName, L, T, W, H)
    c.BorderStyle = 0
    c.BackStyle = 0
    Set RBox = c
End Function

Private Sub SetRptProp(ByVal PropName As String, ByVal Value As Variant)
    On Error Resume Next
    m_rpt.Properties(PropName).Value = Value
End Sub

Private Sub SetCtl(ByVal c As Access.Control, ByVal PropName As String, ByVal Value As Variant)
    On Error Resume Next
    c.Properties(PropName).Value = Value
End Sub

Private Sub FinishReport(ByVal FinalName As String, ByVal Code As String)
    Dim mdl As Access.Module, i As Long, hasExplicit As Boolean
    Set mdl = m_rpt.Module
    For i = 1 To mdl.CountOfDeclarationLines
        If Trim$(mdl.Lines(i, 1)) = "Option Explicit" Then hasExplicit = True
    Next
    If Not hasExplicit Then mdl.InsertLines mdl.CountOfDeclarationLines + 1, "Option Explicit"
    mdl.AddFromString Code
    DoCmd.Close acReport, m_tmp, acSaveYes
    DoCmd.Rename FinalName, acReport, m_tmp
    m_built = m_built + 1
    Debug.Print "  + " & FinalName
End Sub

Private Sub AbortReport(ByVal FinalName As String, ByVal ErrNumber As Long, ByVal ErrText As String)
    m_failed = m_failed + 1
    m_report = m_report & "- " & FinalName & ": " & ErrNumber & " " & ErrText & vbCrLf
    Debug.Print "[X] " & FinalName & ": " & ErrNumber & " " & ErrText
    On Error Resume Next
    DoCmd.Close acReport, m_tmp, acSaveNo
End Sub

'------------------------------------------------------------------------------
' Generated: one procedure per report
'------------------------------------------------------------------------------
Private Sub BuildReport_rptSalesReceipt()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSalesReceipt", "›« Ê—… (Õ—«—Ì 80 „„)", "qrySalesDocPrint", 4196, "DocID", "LineNumber", False, False
    SetSection 0, 539
    SetSection 5, 3827
    SetSection 6, 4734
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(0, "txtProduct", "=[ProductName] & IIf(Len(Nz([LineNote],""""))>0,"" - "" & [LineNote],"""")", 0, 0, 4196, 255, 9, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtQtyPrice", "=[Quantity] & "" ◊ "" & Format([LineTotal]/[Quantity],""#,##0.00"")", 113, 266, 2835, 238, 8, False, 0)
    Set c = RText(0, "txtLineTotal", "LineTotal", 2948, 266, 1247, 238, 9, True, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 4196, 340, 12, True, 2)
    Set c = RText(5, "txtStoreNameEn", "=Nz(SettingValue(""StoreNameEn""),"""")", 0, 397, 4196, 255, 9, False, 2)
    Set c = RText(5, "txtStoreVat", "=""«·—ﬁ„ «·÷—Ì»Ì: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 652, 4196, 227, 9, True, 2)
    Set c = RText(5, "txtStoreCR", "=""”. : "" & Nz(SettingValue(""CRNumber""),""-"") & ""   Â« ›: "" & Nz(SettingValue(""Phone""),""-"")", 0, 879, 4196, 227, 8, False, 2)
    Set c = RText(5, "txtStoreAddress", "=Trim(Nz(SettingValue(""City""),"""") & "" "" & Nz(SettingValue(""District""),"""") & "" "" & Nz(SettingValue(""StreetName""),""""))", 0, 1106, 4196, 227, 8, False, 2)
    Set c = RLine(5, "lnHeader1", 1361, 4196)
    Set c = RText(5, "txtTitle", "=IIf([DocKind]=""RETURN"",""≈‘⁄«— œ«∆‰"",IIf([InvoiceSubType]=""STANDARD"",""›« Ê—… ÷—Ì»Ì…"",""›« Ê—… ÷—Ì»Ì… „»”ÿ…""))", 0, 1418, 4196, 312, 12, True, 2)
    Set c = RText(5, "txtTitleEn", "=IIf([DocKind]=""RETURN"",""Credit Note"",IIf([InvoiceSubType]=""STANDARD"",""Tax Invoice"",""Simplified Tax Invoice""))", 0, 1730, 4196, 227, 8, False, 2)
    Set c = RText(5, "txtDocNumber", "=""—ﬁ„: "" & [DocNumber]", 0, 1985, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtDocDate", "=""«· «—ÌŒ: "" & Format([DocDate],""yyyy/mm/dd hh:nn"")", 0, 2212, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtOriginal", "=IIf(Len(Nz([OriginalNumber],""""))>0,""⁄‰ «·›« Ê—…: "" & [OriginalNumber],OrderTypeText([OrderType],[TableNo],[OrderName]))", 0, 2439, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtCashier", "=""«·ﬂ«‘Ì—: "" & [EmployeeName]", 0, 2666, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtCustomer", "=IIf([CustomerID]=Nz(SettingValue(""DefaultCustomerID""),1),"""",""«·⁄„Ì·: "" & [CustomerName])", 0, 2893, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtCustomerVat", "=IIf(Len(Nz([CustomerVAT],""""))>0,""«·—ﬁ„ «·÷—Ì»Ì ··⁄„Ì·: "" & [CustomerVAT],"""")", 0, 3120, 4196, 227, 8, False, 0)
    Set c = RLine(5, "lnHeader2", 3375, 4196)
    Set c = RLabel(5, "lblColItem", "«·’‰›", 0, 3404, 2211, 255, 8, True, 0)
    Set c = RLabel(5, "lblColQty", "«·ﬂ„Ì… ◊ «·”⁄—", 2211, 3404, 1134, 255, 8, True, 0)
    Set c = RLabel(5, "lblColTotal", "«·≈Ã„«·Ì", 3345, 3404, 850, 255, 8, True, 1)
    Set c = RLine(5, "lnHeader3", 3688, 4196)
    Set c = RLine(6, "lnTotals1", 45, 4196)
    Set c = RLabel(6, "lblCapTaxableAmount", "«·≈Ã„«·Ì »œÊ‰ «·÷—Ì»…", 0, 113, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtTaxableAmount", "TaxableAmount", 2665, 113, 1531, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapDocDiscount", "«·Œ’„", 0, 397, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtDocDiscount", "DocDiscount", 2665, 397, 1531, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtCapDocTax", "=""÷—Ì»… «·ﬁÌ„… «·„÷«›… "" & Format(Nz(SettingValue(""VATRate""),0.15),""0%"")", 0, 681, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtDocTax", "DocTax", 2665, 681, 1531, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapTotalAmount", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", 0, 965, 2608, 255, 11, True, 0)
    Set c = RText(6, "txtTotalAmount", "TotalAmount", 2665, 965, 1531, 255, 11, True, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(6, "lnTotals2", 1277, 4196)
    Set c = RText(6, "txtCapPaid", "=IIf([DocKind]=""RETURN"",""«·„»·€ «·„—œÊœ"",""«·„œ›Ê⁄"")", 0, 1334, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtValPaid", "PaidAmount", 2665, 1334, 1531, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtCapChange", "=IIf([ChangeDue]>0,""«·»«ﬁÌ ··⁄„Ì·"","""")", 0, 1589, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtValChange", "=IIf([ChangeDue]>0,Format([ChangeDue],""#,##0.00""),"""")", 2665, 1589, 1531, 255, 9, False, 1)
    Set c = RText(6, "txtCapRemaining", "=IIf([RemainingAmount]<>0,IIf([DocKind]=""RETURN"",""Œ’„ „‰ —’Ìœ «·⁄„Ì·"",""«·„ »ﬁÌ ⁄·Ï «·Õ”«»""),"""")", 0, 1844, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtValRemaining", "=IIf([RemainingAmount]<>0,Format([RemainingAmount],""#,##0.00""),"""")", 2665, 1844, 1531, 255, 9, False, 1)
    Set c = RBox(6, "boxQR", 1077, 2184, 2041, 2041)
    Set c = RText(6, "txtDocKind", "DocKind", 0, 2184, 284, 227, 9, False, 0)
    SetCtl c, "Visible", False
    Set c = RText(6, "txtDocID", "DocID", 0, 2439, 284, 227, 9, False, 0)
    SetCtl c, "Visible", False
    Set c = RText(6, "txtFooterNote", "=Nz(SettingValue(""ReceiptFooter""),"""")", 0, 4282, 4196, 255, 9, False, 2)
    m_rpt.Section(6).OnPrint = EP
    s = ""
    s = s & "Private Sub secTotals_Print(Cancel As Integer, PrintCount As Integer)" & vbCrLf
    s = s & "    If Me.HasData = 0 Then Exit Sub   ' no document: the fields have no value (2427)" & vbCrLf
    s = s & "    DrawDocumentQR Me, Me!txtDocKind.Value, Me!txtDocID.Value, Me!boxQR.Left, _" & vbCrLf
    s = s & "                   Me!boxQR.Top, Me!boxQR.Width" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSalesReceipt", s
    Exit Sub
EH:
    AbortReport "rptSalesReceipt", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptSalesInvoiceA4()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSalesInvoiceA4", "›« Ê—… ÷—Ì»Ì… (A4)", "qrySalesDocPrint", 10773, "DocID", "LineNumber", False, False
    SetSection 0, 340
    SetSection 5, 3856
    SetSection 6, 4649
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(0, "txtCol1", "LineNumber", 0, 28, 454, 284, 8, False, 2)
    Set c = RText(0, "txtCol2", "=[ProductName] & IIf(Len(Nz([LineNote],""""))>0,"" - "" & [LineNote],"""")", 454, 28, 3515, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "Quantity", 3969, 28, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol4", "UnitPrice", 4876, 28, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "LineDiscount", 6010, 28, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "NetAmount", 6917, 28, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "VATRate", 8051, 28, 680, 284, 8, False, 2)
    SetCtl c, "Format", "0%"
    Set c = RText(0, "txtCol8", "LineTax", 8731, 28, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "LineTotal", 9638, 28, 1135, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 5273, 397, 14, True, 0)
    Set c = RText(5, "txtStoreNameEn", "=Nz(SettingValue(""StoreNameEn""),"""")", 0, 454, 5273, 284, 10, False, 0)
    Set c = RText(5, "txtStoreVat", "=""«·—ﬁ„ «·÷—Ì»Ì: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 738, 5273, 284, 10, True, 0)
    Set c = RText(5, "txtStoreCR", "=""«·”Ã· «· Ã«—Ì: "" & Nz(SettingValue(""CRNumber""),"""")", 0, 1022, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtStoreAddress", "=Trim(Nz(SettingValue(""BuildingNo""),"""") & "" "" & Nz(SettingValue(""StreetName""),"""") & "" - "" & Nz(SettingValue(""District""),"""") & "" - "" & Nz(SettingValue(""City""),"""") & "" "" & Nz(SettingValue(""PostalCode""),""""))", 0, 1306, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtStorePhone", "=""Â« ›: "" & Nz(SettingValue(""Phone""),"""")", 0, 1590, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtTitle", "=IIf([DocKind]=""RETURN"",""≈‘⁄«— œ«∆‰"",IIf([InvoiceSubType]=""STANDARD"",""›« Ê—… ÷—Ì»Ì…"",""›« Ê—… ÷—Ì»Ì… „»”ÿ…""))", 5500, 57, 5273, 454, 16, True, 1)
    Set c = RText(5, "txtTitleEn", "=IIf([DocKind]=""RETURN"",""Credit Note"",IIf([InvoiceSubType]=""STANDARD"",""Tax Invoice"",""Simplified Tax Invoice""))", 5500, 539, 5273, 284, 10, False, 1)
    Set c = RText(5, "txtDocNumber", "=""—ﬁ„ «·›« Ê—…: "" & [DocNumber]", 5500, 850, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtDocDate", "=""«· «—ÌŒ: "" & Format([DocDate],""yyyy/mm/dd hh:nn"")", 5500, 1105, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtOriginal", "=IIf(Len(Nz([OriginalNumber],""""))>0,""⁄‰ «·›« Ê—…: "" & [OriginalNumber],OrderTypeText([OrderType],[TableNo],[OrderName]))", 5500, 1360, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtPaymentType", "=""ÿ—Ìﬁ… «·»Ì⁄: "" & IIf([PaymentType]=""CREDIT"",""¬Ã·"",""‰ﬁœÌ"")", 5500, 1615, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtCashier", "=""«·„ÊŸ›: "" & [EmployeeName]", 5500, 1870, 5273, 255, 9, False, 1)
    Set c = RBox(5, "boxBuyer", 0, 2183, 10773, 1106)
    SetCtl c, "BorderStyle", 1
    Set c = RLabel(5, "lblBuyer", "»Ì«‰«  «·„‘ —Ì", 113, 2211, 2268, 255, 9, True, 0)
    Set c = RText(5, "txtBuyerName", "=""«·«”„: "" & [CustomerName]", 113, 2467, 5103, 255, 9, False, 0)
    Set c = RText(5, "txtBuyerVat", "=IIf(Len(Nz([CustomerVAT],""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & [CustomerVAT],"""")", 5443, 2467, 5216, 255, 9, False, 0)
    Set c = RText(5, "txtBuyerAddress", "=Trim(Nz([CustomerBuilding],"""") & "" "" & Nz([CustomerStreet],"""") & "" "" & Nz([CustomerDistrict],"""") & "" "" & Nz([CustomerCity],"""") & "" "" & Nz([CustomerPostal],""""))", 113, 2750, 10546, 255, 9, False, 0)
    Set c = RBox(5, "boxColumns", 0, 3430, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(5, "lblCol1", "#", 0, 3487, 454, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol2", "«·’‰›", 454, 3487, 3515, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol3", "«·ﬂ„Ì…", 3969, 3487, 907, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol4", "”⁄— «·ÊÕœ…", 4876, 3487, 1134, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol5", "«·Œ’„", 6010, 3487, 907, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol6", "«·’«›Ì", 6917, 3487, 1134, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol7", "‰”»… «·÷—Ì»…", 8051, 3487, 680, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol8", "«·÷—Ì»…", 8731, 3487, 907, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol9", "«·≈Ã„«·Ì", 9638, 3487, 1135, 255, 8, True, 2)
    Set c = RLine(6, "lnTotals1", 45, 10773)
    Set c = RLabel(6, "lblCapDocSubTotal", "«·≈Ã„«·Ì ﬁ»· «·Œ’„", 6464, 113, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocSubTotal", "DocSubTotal", 8959, 113, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapDocDiscount", "«·Œ’„", 6464, 397, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocDiscount", "DocDiscount", 8959, 397, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapTaxableAmount", "«·≈Ã„«·Ì «·Œ«÷⁄ ··÷—Ì»…", 6464, 681, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtTaxableAmount", "TaxableAmount", 8959, 681, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapDocTax", "÷—Ì»… «·ﬁÌ„… «·„÷«›…", 6464, 965, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocTax", "DocTax", 8959, 965, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapTotalAmount", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", 6464, 1249, 2495, 255, 10, True, 0)
    Set c = RText(6, "txtTotalAmount", "TotalAmount", 8959, 1249, 1814, 255, 10, True, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapPaidAmount", "«·„œ›Ê⁄ / «·„—œÊœ", 6464, 1533, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtPaidAmount", "PaidAmount", 8959, 1533, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapRemainingAmount", "«·„ »ﬁÌ", 6464, 1817, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtRemainingAmount", "RemainingAmount", 8959, 1817, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RBox(6, "boxQR", 4366, 2211, 2041, 2041)
    Set c = RText(6, "txtDocKind", "DocKind", 0, 2211, 284, 227, 9, False, 0)
    SetCtl c, "Visible", False
    Set c = RText(6, "txtDocID", "DocID", 0, 2466, 284, 227, 9, False, 0)
    SetCtl c, "Visible", False
    Set c = RText(6, "txtFooterNote", "=Nz(SettingValue(""ReceiptFooter""),"""")", 0, 4309, 10773, 255, 9, False, 2)
    m_rpt.Section(6).OnPrint = EP
    s = ""
    s = s & "Private Sub secTotals_Print(Cancel As Integer, PrintCount As Integer)" & vbCrLf
    s = s & "    If Me.HasData = 0 Then Exit Sub   ' no document: the fields have no value (2427)" & vbCrLf
    s = s & "    DrawDocumentQR Me, Me!txtDocKind.Value, Me!txtDocID.Value, Me!boxQR.Left, _" & vbCrLf
    s = s & "                   Me!boxQR.Top, Me!boxQR.Width" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSalesInvoiceA4", s
    Exit Sub
EH:
    AbortReport "rptSalesInvoiceA4", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptPurchaseDocument()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptPurchaseDocument", "›« Ê—… / „— Ã⁄ „‘ —Ì« ", "qryPurchaseDocPrint", 10773, "DocID", "LineNumber", False, True
    SetSection 0, 340
    SetSection 5, 3345
    SetSection 6, 2948
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(0, "txtCol1", "LineNumber", 0, 23, 454, 284, 8, False, 2)
    Set c = RText(0, "txtCol2", "ProductCode", 454, 23, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "ProductName", 1701, 23, 2835, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol4", "UnitName", 4536, 23, 794, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "Quantity", 5330, 23, 850, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "UnitCost", 6180, 23, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00##"
    Set c = RText(0, "txtCol7", "LineDiscount", 7314, 23, 850, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "LineTax", 8164, 23, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "LineTotal", 9071, 23, 1702, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 5273, 425, 14, True, 0)
    Set c = RText(5, "txtStoreVat", "=""«·—ﬁ„ «·÷—Ì»Ì: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 510, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtStorePhone", "=""Â« ›: "" & Nz(SettingValue(""Phone""),""-"") & ""   "" & Nz(SettingValue(""City""),"""")", 0, 794, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtTitle", "=IIf([DocKind]=""RETURN"",""„— Ã⁄ „‘ —Ì« "",""›« Ê—… „‘ —Ì« "")", 5500, 57, 5273, 482, 16, True, 1)
    Set c = RText(5, "txtDocNumber", "=""«·—ﬁ„: "" & [DocNumber]", 5500, 567, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtDocDate", "=""«· «—ÌŒ: "" & GDate([DocDate],True)", 5500, 822, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtSupplierRef", "=""—ﬁ„ ›« Ê—… «·„Ê—œ: "" & Nz([SupplierInvoiceNo],""-"")", 5500, 1077, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtOriginal", "=IIf(Len(Nz([OriginalNumber],""""))>0,""⁄‰ ›« Ê—… «·‘—«¡: "" & [OriginalNumber],"""")", 5500, 1332, 5273, 255, 9, False, 1)
    Set c = RBox(5, "boxSupplier", 0, 1701, 10773, 1106)
    SetCtl c, "BorderStyle", 1
    Set c = RText(5, "txtSupplier", "=""«·„Ê—œ: "" & [SupplierName]", 113, 1758, 5103, 284, 10, True, 0)
    Set c = RText(5, "txtSupplierVat", "=IIf(Len(Nz([SupplierVAT],""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & [SupplierVAT],"""")", 5330, 1758, 5330, 284, 9, False, 0)
    Set c = RText(5, "txtPayment", "=IIf([DocKind]=""RETURN"",IIf([PaymentType]=""CASH"",""«” —œ«œ ‰ﬁœÌ"",""Œ’„ „‰ —’Ìœ «·„Ê—œ""),IIf([PaymentType]=""CREDIT"",""‘—«¡ ¬Ã·"",""‘—«¡ ‰ﬁœÌ""))", 113, 2098, 5103, 284, 9, False, 0)
    Set c = RText(5, "txtEmployee", "=""«·„ÊŸ›: "" & [EmployeeName]", 5330, 2098, 5330, 284, 9, False, 0)
    Set c = RText(5, "txtReason", "=IIf(Len(Nz([Reason],""""))>0,""”»» «·≈—Ã«⁄: "" & [Reason],"""")", 113, 2438, 10546, 284, 9, False, 0)
    Set c = RBox(5, "boxColumns", 0, 2920, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(5, "lblCol1", "#", 0, 2965, 454, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol2", "«·ﬂÊœ", 454, 2965, 1247, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol3", "«·’‰›", 1701, 2965, 2835, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol4", "«·ÊÕœ…", 4536, 2965, 794, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol5", "«·ﬂ„Ì…", 5330, 2965, 850, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol6", " ﬂ·›… «·ÊÕœ…", 6180, 2965, 1134, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol7", "«·Œ’„", 7314, 2965, 850, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol8", "«·÷—Ì»…", 8164, 2965, 907, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol9", "«·≈Ã„«·Ì", 9071, 2965, 1702, 284, 8, True, 2)
    Set c = RLine(6, "lnTotals", 45, 10773)
    Set c = RLabel(6, "lblCapDocSubTotal", "«·≈Ã„«·Ì ﬁ»· «·Œ’„", 6464, 113, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocSubTotal", "DocSubTotal", 8959, 113, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapDocDiscount", "«·Œ’„", 6464, 397, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocDiscount", "DocDiscount", 8959, 397, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapTaxableAmount", "«·Œ«÷⁄ ··÷—Ì»…", 6464, 681, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtTaxableAmount", "TaxableAmount", 8959, 681, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapDocTax", "÷—Ì»… «·ﬁÌ„… «·„÷«›…", 6464, 965, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocTax", "DocTax", 8959, 965, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapTotalAmount", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", 6464, 1249, 2495, 255, 10, True, 0)
    Set c = RText(6, "txtTotalAmount", "TotalAmount", 8959, 1249, 1814, 255, 10, True, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtCapPaid", "=IIf([DocKind]=""RETURN"",""«·„” —œ ‰ﬁœ«"",""«·„œ›Ê⁄"")", 6464, 1533, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtPaidAmount", "PaidAmount", 8959, 1533, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtCapRemaining", "=IIf([DocKind]=""RETURN"",""Œ’„ „‰ —’Ìœ «·„Ê—œ"",""«·„ »ﬁÌ ⁄·Ï «·Õ”«»"")", 6464, 1817, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtRemainingAmount", "RemainingAmount", 8959, 1817, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtWords", "=AmountInWords([TotalAmount])", 0, 113, 6237, 567, 10, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RLabel(6, "lblSign1", "«·„” ·„: ....................", 0, 2438, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign2", "√„Ì‰ «·„Œ“‰: ....................", 3591, 2438, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign3", "«·„œÌ—: ....................", 7182, 2438, 3591, 312, 10, False, 0)
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""«·„” ‰œ €Ì— „ÊÃÊœ.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptPurchaseDocument", s
    Exit Sub
EH:
    AbortReport "rptPurchaseDocument", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptVoucher()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptVoucher", "”‰œ ﬁ»÷ / ”‰œ ’—›", "qryVoucherPrint", 10773, "DocID", "LineNumber", False, True
    SetSection 0, 0
    SetSection 5, 3969
    SetSection 6, 907
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 5386, 425, 14, True, 0)
    Set c = RText(5, "txtStoreVat", "=""«·—ﬁ„ «·÷—Ì»Ì: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 510, 5386, 284, 9, False, 0)
    Set c = RText(5, "txtStorePhone", "=""Â« ›: "" & Nz(SettingValue(""Phone""),""-"") & ""   "" & Nz(SettingValue(""City""),"""")", 0, 794, 5386, 284, 9, False, 0)
    Set c = RText(5, "txtTitle", "=IIf([DocKind]=""RECEIPT"",""”‰œ ﬁ»÷"",""”‰œ ’—›"")", 5386, 57, 5387, 539, 18, True, 1)
    Set c = RText(5, "txtDocNumber", "=""—ﬁ„ «·”‰œ: "" & [DocNumber]", 5386, 624, 5387, 284, 10, False, 1)
    Set c = RText(5, "txtDocDate", "=""«· «—ÌŒ: "" & GDate([DocDate],True)", 5386, 936, 5387, 284, 10, False, 1)
    Set c = RLine(5, "lnTop", 1304, 10773)
    Set c = RText(5, "txtParty", "=IIf([DocKind]=""RECEIPT"",""«” ·„‰« „‰: "",""’—›‰« ≈·Ï: "") & [PartyName]", 0, 1474, 10773, 397, 13, True, 0)
    Set c = RBox(5, "boxAmount", 0, 1984, 3402, 680)
    SetCtl c, "BorderStyle", 1
    Set c = RText(5, "txtAmount", "=Format([Amount],""#,##0.00"") & "" —Ì«·""", 57, 2041, 3289, 567, 18, True, 2)
    Set c = RText(5, "txtWords", "=AmountInWords([Amount])", 3572, 2070, 7201, 567, 11, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(5, "txtMethod", "=""ÿ—Ìﬁ… «·œ›⁄: "" & [MethodName]", 0, 2835, 5386, 284, 10, False, 0)
    Set c = RText(5, "txtEmployee", "=""«·„ÊŸ›: "" & [EmployeeName]", 5386, 2835, 5387, 284, 10, False, 0)
    Set c = RText(5, "txtNotes", "=""«·»Ì«‰: "" & Nz([Notes],""-"")", 0, 3175, 10773, 312, 10, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(5, "txtBalance", "=IIf([DocKind]=""RECEIPT"",""«·—’Ìœ «·„ »ﬁÌ ⁄·Ï «·⁄„Ì· Õ«·Ì«: "",""«·—’Ìœ «·„” Õﬁ ··„Ê—œ Õ«·Ì«: "") & Format([PartyBalance],""#,##0.00"")", 0, 3572, 10773, 284, 9, False, 0)
    Set c = RLabel(6, "lblSign1", "«·„” ·„: ....................", 0, 284, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign2", "«·„Õ«”»: ....................", 3591, 284, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign3", "«·„œÌ—: ....................", 7182, 284, 3591, 312, 10, False, 0)
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""«·”‰œ €Ì— „ÊÃÊœ.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptVoucher", s
    Exit Sub
EH:
    AbortReport "rptVoucher", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptStockCount()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptStockCount", "Ê—ﬁ… «·Ã—œ", "StockCountQuery", 10773, "StockCountID", "ProductName", False, True
    SetSection 0, 340
    SetSection 5, 2041
    SetSection 6, 1474
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(0, "txtCol1", "=1", 0, 23, 510, 284, 8, False, 0)
    SetCtl c, "RunningSum", 1
    Set c = RText(0, "txtCol2", "ProductCode", 510, 23, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "ProductName", 1871, 23, 3402, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol4", "SystemQuantity", 5273, 23, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "ActualQuantity", 6520, 23, 1304, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "=IIf(IsNull([ActualQuantity]),Null,[Difference])", 7824, 23, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol7", "=IIf(IsNull([ActualQuantity]),Null,[DifferenceValue])", 8958, 23, 1815, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 5273, 425, 14, True, 0)
    Set c = RText(5, "txtStoreVat", "=""«·—ﬁ„ «·÷—Ì»Ì: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 510, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtStorePhone", "=""Â« ›: "" & Nz(SettingValue(""Phone""),""-"") & ""   "" & Nz(SettingValue(""City""),"""")", 0, 794, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtTitle", "=""Ê—ﬁ… Ã—œ "" & IIf([Status]=""POSTED"",""(„ı—Õ¯·)"",IIf([Status]=""CANCELLED"",""(„·€Ï)"",""(„› ÊÕ)""))", 5500, 57, 5273, 482, 16, True, 1)
    Set c = RText(5, "txtCountNumber", "=""—ﬁ„ «·Ã—œ: "" & [CountNumber] & ""    «· «—ÌŒ: "" & GDate([CountDate])", 5500, 567, 5273, 284, 9, False, 1)
    Set c = RText(5, "txtCategory", "=""«· ’‰Ì›: "" & Nz([CategoryName],""ﬂ· «·„‰ Ã« "")", 5500, 879, 5273, 284, 9, False, 1)
    Set c = RBox(5, "boxColumns", 0, 1644, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(5, "lblCol1", "#", 0, 1689, 510, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol2", "«·ﬂÊœ", 510, 1689, 1361, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol3", "«·’‰›", 1871, 1689, 3402, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol4", "«·„”Ã·", 5273, 1689, 1247, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol5", "«·›⁄·Ì", 6520, 1689, 1304, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol6", "«·›—ﬁ", 7824, 1689, 1134, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol7", "ﬁÌ„… «·›—ﬁ", 8958, 1689, 1815, 284, 8, True, 2)
    Set c = RLine(6, "lnTotals", 45, 10773)
    Set c = RText(6, "txtSummary", "="" „ ⁄œ¯ "" & Count([ActualQuantity]) & "" „‰ "" & Count(*) & "" ’‰›    «·⁄Ã“: "" & Format(-Sum(IIf([DifferenceValue]<0,[DifferenceValue],0)),""#,##0.00"") & ""    «·“Ì«œ…: "" & Format(Sum(IIf([DifferenceValue]>0,[DifferenceValue],0)),""#,##0.00"")", 0, 113, 10773, 312, 10, True, 0)
    Set c = RLabel(6, "lblSign1", "«·ﬁ«∆„ »«·Ã—œ: ....................", 0, 850, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign2", "«·„—«Ã⁄: ....................", 3591, 850, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign3", "«·„œÌ—: ....................", 7182, 850, 3591, 312, 10, False, 0)
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""Ã·”… «·Ã—œ €Ì— „ÊÃÊœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptStockCount", s
    Exit Sub
EH:
    AbortReport "rptStockCount", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptBarcodeLabels()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    Application.Run "EnsureLabelTables"
    StartReport "rptBarcodeLabels", "„·’ﬁ«  «·»«—ﬂÊœ", "SELECT q.LineNo, n.N, p.ProductName, p.ProductCode, p.Barcode, p.SellingPrice FROM tmpLabelQueue AS q, tmpLabelNumbers AS n, Products AS p WHERE p.ProductID = q.ProductID AND n.N <= q.Copies", 2155, "", "LineNo,N", False, False
    SetSection 0, 1418
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(0, "txtTop1", "=""«·„ Ã—""", 57, 57, 2041, 227, 7, False, 2)
    Set c = RText(0, "txtTop2", "=[ProductName]", 57, 284, 2041, 227, 7, False, 2)
    Set c = RBox(0, "boxBar", 57, 539, 2041, 510)
    SetCtl c, "BorderStyle", 0
    SetCtl c, "BackStyle", 0
    Set c = RText(0, "txtBottom1", "=LabelCode([Barcode],[ProductCode])", 57, 907, 2041, 227, 7, False, 2)
    Set c = RText(0, "txtBottom2", "=LabelPrice([SellingPrice])", 57, 1134, 2041, 227, 7, True, 2)
    Set c = RText(0, "txtCode", "=LabelCode([Barcode],[ProductCode])", 0, 0, 57, 57, 6, False, 0)
    SetCtl c, "Visible", False
    m_rpt.OnNoData = EP
    m_rpt.Section(0).Name = "secLabel"
    m_rpt.Section(0).OnPrint = EP
    m_rpt.OnOpen = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""ﬁ«∆„… «·„·’ﬁ«  ›«—€….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Report_Open(Cancel As Integer)" & vbCrLf
    s = s & "    LabelReportOpen" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub secLabel_Print(Cancel As Integer, PrintCount As Integer)" & vbCrLf
    s = s & "    If Me.HasData = 0 Then Exit Sub" & vbCrLf
    s = s & "    DrawBarcode Me, Nz(Me!txtCode.Value, """"), Me!boxBar.Left, Me!boxBar.Top, Me!boxBar.Width, _" & vbCrLf
    s = s & "                Me!boxBar.Height, LabelBarWidth()" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptBarcodeLabels", s
    Exit Sub
EH:
    AbortReport "rptBarcodeLabels", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptStatistics()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptStatistics", "«·≈Õ’«∆Ì«  Ê«·—”Ê„ «·»Ì«‰Ì…", "", 10773, "", "", False, True
    SetSection 3, 1134
    SetSection 4, 397
    SetSection 0, 13381
    HideSection 1
    HideSection 2
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 5500, 57, 5273, 397, 13, True, 3)
    Set c = RLabel(3, "lblTitle", "«·≈Õ’«∆Ì«  Ê«·—”Ê„ «·»Ì«‰Ì…", 0, 57, 5273, 425, 16, True, 0)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 539, 10773, 284, 9, False, 2)
    Set c = RLine(3, "lnHeader", 964, 10773)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6237, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6237, 57, 4536, 255, 8, False, 3)
    Set c = RBox(0, "boxCharts", 0, 0, 10773, 13381)
    SetCtl c, "BorderStyle", 0
    SetCtl c, "BackStyle", 0
    m_rpt.Section(0).Name = "secCharts"
    m_rpt.Section(0).OnPrint = EP
    s = ""
    s = s & "Private Sub secCharts_Print(Cancel As Integer, PrintCount As Integer)" & vbCrLf
    s = s & "    DrawStatistics Me, Me!boxCharts.Left, Me!boxCharts.Top, Me!boxCharts.Width, Me!boxCharts.Height" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptStatistics", s
    Exit Sub
EH:
    AbortReport "rptStatistics", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptCashVoucher()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptCashVoucher", "”‰œ ‰ﬁœÌ…", "qryCashVoucherPrint", 10773, "DocID", "", False, True
    SetSection 0, 0
    SetSection 5, 4309
    SetSection 6, 907
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 5386, 425, 14, True, 0)
    Set c = RText(5, "txtStoreVat", "=""«·—ﬁ„ «·÷—Ì»Ì: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 510, 5386, 284, 9, False, 0)
    Set c = RText(5, "txtStorePhone", "=""Â« ›: "" & Nz(SettingValue(""Phone""),""-"") & ""   "" & Nz(SettingValue(""City""),"""")", 0, 794, 5386, 284, 9, False, 0)
    Set c = RText(5, "txtTitle", "VoucherTitle", 5386, 57, 5387, 539, 18, True, 1)
    Set c = RText(5, "txtDocNumber", "=""—ﬁ„ «·”‰œ: "" & [VoucherNumber]", 5386, 624, 5387, 284, 10, False, 1)
    Set c = RText(5, "txtDocDate", "=""«· «—ÌŒ: "" & GDate([VoucherDate],True)", 5386, 936, 5387, 284, 10, False, 1)
    Set c = RLine(5, "lnTop", 1304, 10773)
    Set c = RText(5, "txtParty", "=IIf([VoucherType]=""IN"",""«” ·„‰« „‰: "",IIf([VoucherType]=""OUT"",""’—›‰« ≈·Ï: "",""ÕıÊˆ¯· ≈·Ï: "")) & IIf([VoucherType]=""TRANSFER"",[ToBoxName],Nz([PartyName],""-""))", 0, 1474, 10773, 397, 13, True, 0)
    Set c = RBox(5, "boxAmount", 0, 1984, 3402, 680)
    SetCtl c, "BorderStyle", 1
    Set c = RText(5, "txtAmount", "=Format([Amount],""#,##0.00"") & "" —Ì«·""", 57, 2041, 3289, 567, 18, True, 2)
    Set c = RText(5, "txtWords", "=AmountInWords([Amount])", 3572, 2070, 7201, 567, 11, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(5, "txtCategory", "=""«·»‰œ: "" & [CategoryName] & IIf(IsNull([ExpenseTypeName]),"""","" - "" & [ExpenseTypeName])", 0, 2835, 5386, 284, 10, False, 0)
    Set c = RText(5, "txtBox", "=IIf([VoucherType]=""IN"",""«·’‰œÊﬁ: "",""„‰ ’‰œÊﬁ: "") & [BoxName]", 5386, 2835, 5387, 284, 10, False, 0)
    Set c = RText(5, "txtNotes", "=""«·»Ì«‰: "" & Nz([Description],""-"")", 0, 3175, 10773, 312, 10, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(5, "txtEmployee", "=""«·„ÊŸ›: "" & [EmployeeName]", 0, 3572, 10773, 284, 9, False, 0)
    Set c = RLabel(6, "lblSign1", "«·„” ·„: ....................", 0, 284, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign2", "√„Ì‰ «·’‰œÊﬁ: ....................", 3591, 284, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign3", "«·„œÌ—: ....................", 7182, 284, 3591, 312, 10, False, 0)
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""«·”‰œ €Ì— „ÊÃÊœ.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptCashVoucher", s
    Exit Sub
EH:
    AbortReport "rptCashVoucher", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptCashClosing()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptCashClosing", " ’›Ì… ÌÊ„Ì… «·ﬂ«‘Ì—", "qryCashClosingPrint", 10773, "ClosingID", "", False, True
    SetSection 0, 0
    SetSection 5, 7031
    SetSection 6, 907
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 5386, 425, 14, True, 0)
    Set c = RText(5, "txtStoreVat", "=""«·—ﬁ„ «·÷—Ì»Ì: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 510, 5386, 284, 9, False, 0)
    Set c = RText(5, "txtStorePhone", "=""Â« ›: "" & Nz(SettingValue(""Phone""),""-"") & ""   "" & Nz(SettingValue(""City""),"""")", 0, 794, 5386, 284, 9, False, 0)
    Set c = RText(5, "txtTitle", "="" ’›Ì… ÌÊ„Ì… «·ﬂ«‘Ì—""", 5386, 57, 5387, 539, 18, True, 1)
    Set c = RText(5, "txtDocNumber", "=""—ﬁ„ «· ’›Ì…: "" & [ClosingNumber]", 5386, 624, 5387, 284, 10, False, 1)
    Set c = RText(5, "txtDocDate", "=""«· «—ÌŒ: "" & GDate([ClosingDate],True)", 5386, 936, 5387, 284, 10, False, 1)
    Set c = RLine(5, "lnTop", 1304, 10773)
    Set c = RText(5, "txtBox", "=""«·’‰œÊﬁ: "" & [BoxName] & ""    «·ﬂ«‘Ì—: "" & [EmployeeName]", 0, 1474, 10773, 397, 13, True, 0)
    Set c = RText(5, "txtPeriod", "=""«·› —…: "" & IIf(IsNull([PeriodStart]),""„‰ »œ«Ì… «·’‰œÊﬁ"",""„‰ "" & GDate([PeriodStart],True)) & "" ≈·Ï "" & GDate([ClosingDate],True)", 0, 1928, 10773, 284, 10, False, 0)
    Set c = RLabel(5, "lblCap1", "—’Ìœ «·»œ«Ì…", 1134, 2381, 5386, 340, 11, False, 0)
    Set c = RText(5, "txtVal1", "OpeningBalance", 6577, 2381, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(5, "lblCap2", "«·„ﬁ»Ê÷«  (+)", 1134, 2835, 5386, 340, 11, False, 0)
    Set c = RText(5, "txtVal2", "CashIn", 6577, 2835, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(5, "lblCap3", "«·„œ›Ê⁄«  (-)", 1134, 3289, 5386, 340, 11, False, 0)
    Set c = RText(5, "txtVal3", "CashOut", 6577, 3289, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(5, "lblCap4", "«·—’Ìœ «·œ› —Ì («·„›—Ê÷)", 1134, 3743, 5386, 340, 11, True, 0)
    Set c = RText(5, "txtVal4", "ExpectedBalance", 6577, 3743, 3062, 340, 12, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(5, "lblCap5", "«·‰ﬁœÌ… «·›⁄·Ì… »«·⁄œ¯", 1134, 4197, 5386, 340, 11, True, 0)
    Set c = RText(5, "txtVal5", "CountedAmount", 6577, 4197, 3062, 340, 12, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(5, "txtCap6", "=IIf([Difference]<0,""«·⁄Ã“"",IIf([Difference]>0,""«·“Ì«œ…"",""«·›—ﬁ""))", 1134, 4651, 5386, 340, 11, True, 0)
    Set c = RText(5, "txtVal6", "=IIf([Difference]<0,-[Difference],[Difference])", 6577, 4651, 3062, 340, 12, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(5, "txtCap7", "=""«·„—ÕÛ¯· ≈·Ï: "" & IIf(IsNull([ToBoxName]),[DestinationName],[ToBoxName])", 1134, 5105, 5386, 340, 11, False, 0)
    Set c = RText(5, "txtVal7", "TransferAmount", 6577, 5105, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(5, "lblCap8", "«·„ »ﬁÌ ›Ì «·’‰œÊﬁ (⁄Âœ…)", 1134, 5559, 5386, 340, 11, False, 0)
    Set c = RText(5, "txtVal8", "KeptAmount", 6577, 5559, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(5, "txtNotes", "=""„·«ÕŸ« : "" & Nz([Notes],""-"")", 0, 6126, 10773, 312, 10, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RLabel(6, "lblSign1", "«·ﬂ«‘Ì—: ....................", 0, 284, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign2", "«·„” ·„: ....................", 3591, 284, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign3", "«·„œÌ—: ....................", 7182, 284, 3591, 312, 10, False, 0)
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""«· ’›Ì… €Ì— „ÊÃÊœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptCashClosing", s
    Exit Sub
EH:
    AbortReport "rptCashClosing", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptJournalEntry()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptJournalEntry", "ﬁÌœ ÌÊ„Ì…", "qryJournalEntryPrint", 10773, "EntryID", "LineNumber", False, True
    SetSection 0, 340
    SetSection 5, 2495
    SetSection 6, 1474
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(0, "txtCol1", "AccountCode", 0, 23, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "AccountName", 1247, 23, 2948, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "LineText", 4195, 23, 3402, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "=IIf([Debit]=0,Null,[Debit])", 7597, 23, 1588, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "=IIf([Credit]=0,Null,[Credit])", 9185, 23, 1588, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 5386, 425, 14, True, 0)
    Set c = RText(5, "txtStoreVat", "=""«·—ﬁ„ «·÷—Ì»Ì: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 510, 5386, 284, 9, False, 0)
    Set c = RText(5, "txtStorePhone", "=""Â« ›: "" & Nz(SettingValue(""Phone""),""-"") & ""   "" & Nz(SettingValue(""City""),"""")", 0, 794, 5386, 284, 9, False, 0)
    Set c = RText(5, "txtTitle", "=""ﬁÌœ ÌÊ„Ì…""", 5386, 57, 5387, 539, 18, True, 1)
    Set c = RText(5, "txtDocNumber", "=""—ﬁ„ «·ﬁÌœ: "" & [EntryNumber]", 5386, 624, 5387, 284, 10, False, 1)
    Set c = RText(5, "txtDocDate", "=""«· «—ÌŒ: "" & GDate([EntryDate],True)", 5386, 936, 5387, 284, 10, False, 1)
    Set c = RText(5, "txtSource", "=""«·⁄„·Ì…: "" & [TypeName] & "" "" & Nz([SourceNumber],"""")", 0, 1304, 10773, 312, 11, True, 0)
    Set c = RText(5, "txtDescription", "=""«·»Ì«‰: "" & Nz([Description],""-"")", 0, 1644, 10773, 284, 9, False, 0)
    Set c = RBox(5, "boxColumns", 0, 2041, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(5, "lblCol1", "«·Õ”«»", 0, 2086, 1247, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol2", "«”„ «·Õ”«»", 1247, 2086, 2948, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol3", "«·»Ì«‰", 4195, 2086, 3402, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol4", "„œÌ‰", 7597, 2086, 1588, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol5", "œ«∆‰", 9185, 2086, 1588, 284, 8, True, 2)
    Set c = RLine(6, "lnTotals", 45, 10773)
    Set c = RText(6, "txtTotalCaption", "=""«·≈Ã„«·Ì""", 0, 113, 7598, 312, 10, True, 0)
    Set c = RText(6, "txtTotalDebit", "=Sum([Debit])", 7598, 113, 1588, 312, 10, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtTotalCredit", "=Sum([Credit])", 9186, 113, 1587, 312, 10, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblSign1", "«·„Õ«”»: ....................", 0, 794, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign2", "«·„—«Ã⁄: ....................", 3591, 794, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign3", "«·„œÌ—: ....................", 7182, 794, 3591, 312, 10, False, 0)
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""«·ﬁÌœ €Ì— „ÊÃÊœ.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptJournalEntry", s
    Exit Sub
EH:
    AbortReport "rptJournalEntry", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptAccountStatement()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptAccountStatement", "ﬂ‘› Õ”«»", "AccountStatementQuery", 10773, "StatementAccount", "SortKey,LineDate,EntryNo", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 0, 312
    SetSection 5, 454
    SetSection 6, 1134
    HideSection 1
    HideSection 2
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "ﬂ‘› Õ”«»", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«· «—ÌŒ", 0, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·ﬁÌœ", 1134, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·⁄„·Ì…", 2268, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·„” ‰œ", 3686, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·»Ì«‰", 4820, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·Õ”«»", 6294, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "„œÌ‰", 7315, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "œ«∆‰", 8449, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "«·—’Ìœ", 9583, 1292, 1190, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(0, "txtCol1", "=GDate([LineDate])", 0, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "EntryNo", 1134, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "KindName", 2268, 17, 1418, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "DocNo", 3686, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "Details", 4820, 17, 1474, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol6", "SubName", 6294, 17, 1021, 284, 8, False, 0)
    Set c = RText(0, "txtCol7", "=IIf([LineDebit]=0,Null,[LineDebit])", 7315, 17, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "=IIf([LineCredit]=0,Null,[LineCredit])", 8449, 17, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "=[LineDebit]-[LineCredit]", 9583, 17, 1190, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    SetCtl c, "RunningSum", 1
    Set c = RText(5, "txtAccount", "=""«·Õ”«»: "" & [StatementAccount] & ""  "" & [StatementName]", 0, 68, 10773, 340, 11, True, 0)
    Set c = RLine(6, "lnTotals", 28, 10773)
    Set c = RLabel(6, "lblSum1", "—’Ìœ √Ê· «·„œ…", 0, 85, 2693, 284, 9, True, 2)
    Set c = RText(6, "txtSum1", "=Sum(IIf([SortKey]=0,[LineDebit]-[LineCredit],0))", 0, 397, 2693, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblSum2", "„œÌ‰ «·› —…", 2693, 85, 2693, 284, 9, True, 2)
    Set c = RText(6, "txtSum2", "=Sum(IIf([SortKey]=1,[LineDebit],0))", 2693, 397, 2693, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblSum3", "œ«∆‰ «·› —…", 5386, 85, 2693, 284, 9, True, 2)
    Set c = RText(6, "txtSum3", "=Sum(IIf([SortKey]=1,[LineCredit],0))", 5386, 397, 2693, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblSum4", "«·—’Ìœ «·Œ «„Ì", 8079, 85, 2693, 284, 9, True, 2)
    Set c = RText(6, "txtSum4", "=Sum([LineDebit]-[LineCredit])", 8079, 397, 2693, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtNature", "=IIf(Sum([LineDebit]-[LineCredit])>=0,""«·—’Ìœ „œÌ‰"",""«·—’Ìœ œ«∆‰"")", 8079, 737, 2694, 284, 9, False, 2)
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ ﬁÌÊœ ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptAccountStatement", s
    Exit Sub
EH:
    AbortReport "rptAccountStatement", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptGeneralLedger()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptGeneralLedger", "œ› — «·√” «–", "GeneralLedgerQuery", 10773, "AccountKey", "SortKey,LineDate,EntryNo", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 0, 312
    SetSection 5, 454
    SetSection 6, 1134
    HideSection 1
    HideSection 2
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "œ› — «·√” «–", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«· «—ÌŒ", 0, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·ﬁÌœ", 1134, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·⁄„·Ì…", 2268, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·„” ‰œ", 3686, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·»Ì«‰", 4820, 1292, 2495, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "„œÌ‰", 7315, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "œ«∆‰", 8449, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "«·—’Ìœ", 9583, 1292, 1190, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(0, "txtCol1", "=GDate([LineDate])", 0, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "EntryNo", 1134, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "KindName", 2268, 17, 1418, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "DocNo", 3686, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "Details", 4820, 17, 2495, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol6", "=IIf([LineDebit]=0,Null,[LineDebit])", 7315, 17, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "=IIf([LineCredit]=0,Null,[LineCredit])", 8449, 17, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "=[LineDebit]-[LineCredit]", 9583, 17, 1190, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    SetCtl c, "RunningSum", 1
    Set c = RText(5, "txtAccount", "=[LedgerCode] & ""  "" & [LedgerName]", 0, 68, 10773, 340, 11, True, 0)
    Set c = RLine(6, "lnTotals", 28, 10773)
    Set c = RLabel(6, "lblSum1", "—’Ìœ √Ê· «·„œ…", 0, 85, 2693, 284, 9, True, 2)
    Set c = RText(6, "txtSum1", "=Sum(IIf([SortKey]=0,[LineDebit]-[LineCredit],0))", 0, 397, 2693, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblSum2", "„œÌ‰ «·› —…", 2693, 85, 2693, 284, 9, True, 2)
    Set c = RText(6, "txtSum2", "=Sum(IIf([SortKey]=1,[LineDebit],0))", 2693, 397, 2693, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblSum3", "œ«∆‰ «·› —…", 5386, 85, 2693, 284, 9, True, 2)
    Set c = RText(6, "txtSum3", "=Sum(IIf([SortKey]=1,[LineCredit],0))", 5386, 397, 2693, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblSum4", "«·—’Ìœ «·Œ «„Ì", 8079, 85, 2693, 284, 9, True, 2)
    Set c = RText(6, "txtSum4", "=Sum([LineDebit]-[LineCredit])", 8079, 397, 2693, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtNature", "=IIf(Sum([LineDebit]-[LineCredit])>=0,""«·—’Ìœ „œÌ‰"",""«·—’Ìœ œ«∆‰"")", 8079, 737, 2694, 284, 9, False, 2)
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ ﬁÌÊœ ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptGeneralLedger", s
    Exit Sub
EH:
    AbortReport "rptGeneralLedger", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptIncomeStatement()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptIncomeStatement", "ﬁ«∆„… «·œŒ·", "IncomeStatementQuery", 10773, "", "Block,AccountKey", False, True
    SetSection 3, 2041
    SetSection 4, 340
    SetSection 0, 340
    HideSection 1
    HideSection 2
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "ﬁ«∆„… «·œŒ·", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 709)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblItem", "«·»‰œ", 0, 1445, 4649, 284, 9, True, 2)
    Set c = RLabel(3, "lblCurrent", "«·› —… «·Õ«·Ì…", 4649, 1292, 3062, 284, 9, True, 2)
    Set c = RLabel(3, "lblPrior", "› —… «·„ﬁ«—‰…", 7711, 1292, 3062, 284, 9, True, 2)
    Set c = RLabel(3, "lblCol1", "«·Õ”«»", 4649, 1616, 1531, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·„Ã„Ê⁄", 6180, 1616, 1531, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·Õ”«»", 7711, 1616, 1531, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·„Ã„Ê⁄", 9242, 1616, 1531, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(0, "txtItem", "=IIf([RowKind]=""A"",""      "" & [Caption],IIf([RowKind]=""S"",""≈Ã„«·Ì "" & [Caption],[Caption]))", 0, 28, 4649, 284, 9, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol1", "=IIf([RowKind]=""A"",[CurrentValue],Null)", 4649, 28, 1531, 284, 9, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol2", "=IIf([RowKind]=""S"" Or [RowKind]=""T"" Or [RowKind]=""R"",[CurrentValue],Null)", 6180, 28, 1531, 284, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol3", "=IIf([RowKind]=""A"",[PriorValue],Null)", 7711, 28, 1531, 284, 9, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol4", "=IIf([RowKind]=""S"" Or [RowKind]=""T"" Or [RowKind]=""R"",[PriorValue],Null)", 9242, 28, 1531, 284, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ ﬁÌÊœ ··› —… «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptIncomeStatement", s
    Exit Sub
EH:
    AbortReport "rptIncomeStatement", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptBalanceSheet()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptBalanceSheet", "«·„Ì“«‰Ì… «·⁄„Ê„Ì… (ﬁ«∆„… «·„—ﬂ“ «·„«·Ì)", "BalanceSheetQuery", 10773, "", "ClassNo,GroupKey,Pos,AccountKey", False, True
    SetSection 3, 2041
    SetSection 4, 340
    SetSection 0, 340
    HideSection 1
    HideSection 2
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·„Ì“«‰Ì… «·⁄„Ê„Ì… (ﬁ«∆„… «·„—ﬂ“ «·„«·Ì)", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 709)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblItem", "«·»‰œ", 0, 1445, 4649, 284, 9, True, 2)
    Set c = RLabel(3, "lblCurrent", "«·› —… «·Õ«·Ì…", 4649, 1292, 3062, 284, 9, True, 2)
    Set c = RLabel(3, "lblPrior", "› —… «·„ﬁ«—‰…", 7711, 1292, 3062, 284, 9, True, 2)
    Set c = RLabel(3, "lblCol1", "«·Õ”«»", 4649, 1616, 1531, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·„Ã„Ê⁄", 6180, 1616, 1531, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·Õ”«»", 7711, 1616, 1531, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·„Ã„Ê⁄", 9242, 1616, 1531, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(0, "txtItem", "=IIf([RowKind]=""A"",""      "" & [Caption],IIf([RowKind]=""S"",""≈Ã„«·Ì "" & [Caption],[Caption]))", 0, 28, 4649, 284, 9, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol1", "=IIf([RowKind]=""A"",[CurrentValue],Null)", 4649, 28, 1531, 284, 9, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol2", "=IIf([RowKind]=""S"" Or [RowKind]=""T"" Or [RowKind]=""R"",[CurrentValue],Null)", 6180, 28, 1531, 284, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol3", "=IIf([RowKind]=""A"",[PriorValue],Null)", 7711, 28, 1531, 284, 9, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol4", "=IIf([RowKind]=""S"" Or [RowKind]=""T"" Or [RowKind]=""R"",[PriorValue],Null)", 9242, 28, 1531, 284, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ ﬁÌÊœ ··› —… «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptBalanceSheet", s
    Exit Sub
EH:
    AbortReport "rptBalanceSheet", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptVatReturn()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptVatReturn", "≈ﬁ—«— ÷—Ì»… «·ﬁÌ„… «·„÷«›…", "VatReturnQuery", 10773, "", "BoxNo", False, True
    SetSection 3, 2098
    SetSection 4, 340
    SetSection 2, 1247
    SetSection 0, 397
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "≈ﬁ—«— ÷—Ì»… «·ﬁÌ„… «·„÷«›…", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RText(3, "txtReturn", "=""«·≈ﬁ—«—: "" & [ReturnNumber] & ""    «·Õ«·…: "" & IIf([Status]=""FILED"",""„⁄ „œ ›Ì "" & GDate([FiledDate]),""„”Êœ… (€Ì— „⁄ „œ)"") & IIf(IsNull([FilingRef]),"""",""    —ﬁ„ «·≈ﬁ—«— ·œÏ «·ÂÌ∆…: "" & [FilingRef])", 0, 1247, 10773, 312, 10, True, 0)
    Set c = RBox(3, "boxColumns", 0, 1644, 10773, 397)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·»‰œ", 0, 1701, 680, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·Ê’›", 680, 1701, 4651, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·„»·€ (»œÊ‰ «·÷—Ì»…)", 5331, 1701, 1814, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«· ⁄œÌ·«  («·„— Ã⁄« )", 7145, 1701, 1814, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "÷—Ì»… «·ﬁÌ„… «·„÷«›…", 8959, 1701, 1814, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RLine(2, "lnNet", 57, 10773)
    Set c = RText(2, "txtNet", "=IIf(Sum(IIf([BoxNo]=16,[VAT],0))>=0,""’«›Ì «·÷—Ì»… «·„” Õﬁ… ··”œ«œ: "",""÷—Ì»… „” —œ…  ı—ÕÛ¯· ··≈ﬁ—«— «· «·Ì: "") & Format(IIf(Sum(IIf([BoxNo]=16,[VAT],0))>=0,1,-1)*Sum(IIf([BoxNo]=16,[VAT],0)),""#,##0.00"")", 0, 170, 10773, 397, 13, True, 0)
    Set c = RLabel(2, "lblSign1", "«·„Õ«”»: ....................", 0, 794, 5386, 312, 10, False, 0)
    Set c = RLabel(2, "lblSign2", "«·„œÌ—: ....................", 5386, 794, 5386, 312, 10, False, 0)
    Set c = RText(0, "txtCol1", "BoxNo", 0, 45, 680, 312, 9, False, 0)
    Set c = RText(0, "txtCol2", "BoxText", 680, 45, 4651, 312, 9, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "Amount", 5331, 45, 1814, 312, 9, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol4", "Adjust", 7145, 45, 1814, 312, 9, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "VAT", 8959, 45, 1814, 312, 9, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(0, "lnRow", 374, 10773)
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""«·≈ﬁ—«— €Ì— „ÊÃÊœ.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptVatReturn", s
    Exit Sub
EH:
    AbortReport "rptVatReturn", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptAging()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptAging", "√⁄„«— «·œÌÊ‰", "tmpAging", 15536, "PartyID", "PartyName,DueDate,LineNo", True, True
    SetSection 3, 1701
    SetSection 4, 340
    SetSection 2, 567
    SetSection 0, 312
    SetSection 5, 425
    SetSection 6, 454
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "√⁄„«— «·œÌÊ‰", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·„” ‰œ", 0, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·—ﬁ„", 1361, 1292, 1588, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«· «—ÌŒ", 2949, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·«” Õﬁ«ﬁ", 4083, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "√Ì«„ «· √ŒÌ—", 5217, 1292, 907, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "€Ì— „” Õﬁ", 6124, 1292, 1344, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "1-30 ÌÊ„«", 7468, 1292, 1344, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "31-60", 8812, 1292, 1344, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "61-90", 10156, 1292, 1344, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "√ﬂÀ— „‰ 90", 11500, 1292, 1344, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "—’Ìœ œ«∆‰", 12844, 1292, 1344, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol12", "«·„ »ﬁÌ", 14188, 1292, 1348, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RLine(2, "lntxtTotal", 23, 15536)
    Set c = RText(2, "txtTotalCaption", "=""«·≈Ã„«·Ì""", 0, 85, 6124, 312, 9, True, 0)
    Set c = RText(2, "txtTotal1", "=Sum([NotDue])", 6124, 85, 1344, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal2", "=Sum([Days30])", 7468, 85, 1344, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal3", "=Sum([Days60])", 8812, 85, 1344, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal4", "=Sum([Days90])", 10156, 85, 1344, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([Over90])", 11500, 85, 1344, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([Credit])", 12844, 85, 1344, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal7", "=Sum([OpenAmount])", 14188, 85, 1348, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol1", "DocTypeName", 0, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "DocNo", 1361, 17, 1588, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "=GDate([DocDate])", 2949, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "=GDate([DueDate])", 4083, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "=IIf([DaysLate]>0,[DaysLate],Null)", 5217, 17, 907, 284, 8, False, 0)
    Set c = RText(0, "txtCol6", "=IIf([NotDue]=0,Null,[NotDue])", 6124, 17, 1344, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "=IIf([Days30]=0,Null,[Days30])", 7468, 17, 1344, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "=IIf([Days60]=0,Null,[Days60])", 8812, 17, 1344, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "=IIf([Days90]=0,Null,[Days90])", 10156, 17, 1344, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "=IIf([Over90]=0,Null,[Over90])", 11500, 17, 1344, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "=IIf([Credit]=0,Null,[Credit])", 12844, 17, 1344, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol12", "=IIf([OpenAmount]=0,Null,[OpenAmount])", 14188, 17, 1348, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(5, "txtParty", "PartyName", 0, 68, 15536, 312, 11, True, 0)
    Set c = RLine(6, "lntxtSum", 23, 15536)
    Set c = RText(6, "txtSumCaption", "=""≈Ã„«·Ì "" & [PartyName]", 0, 85, 6124, 312, 9, True, 0)
    Set c = RText(6, "txtSum1", "=Sum([NotDue])", 6124, 85, 1344, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtSum2", "=Sum([Days30])", 7468, 85, 1344, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtSum3", "=Sum([Days60])", 8812, 85, 1344, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtSum4", "=Sum([Days90])", 10156, 85, 1344, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtSum5", "=Sum([Over90])", 11500, 85, 1344, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtSum6", "=Sum([Credit])", 12844, 85, 1344, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtSum7", "=Sum([OpenAmount])", 14188, 85, 1348, 312, 9, True, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ √—’œ… „› ÊÕ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptAging", s
    Exit Sub
EH:
    AbortReport "rptAging", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptPayroll()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptPayroll", "„”Ì— «·—Ê« »", "PayrollSheetQuery", 15536, "", "EmployeeName", True, True
    SetSection 3, 1701
    SetSection 4, 340
    SetSection 2, 1361
    SetSection 0, 340
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "„”Ì— «·—Ê« »", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·„ÊŸ›", 0, 1292, 2381, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·√”«”Ì", 2381, 1292, 1195, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·”ﬂ‰", 3576, 1292, 1195, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "»œ·«  √Œ—Ï", 4771, 1292, 1195, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·≈÷«›Ì", 5966, 1292, 1195, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "„ﬂ«›¬ ", 7161, 1292, 1195, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "«·≈Ã„«·Ì", 8356, 1292, 1195, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "€Ì«»", 9551, 1292, 1195, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "”·›…", 10746, 1292, 1195, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "Ã“«¡« ", 11941, 1292, 1195, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "«· √„Ì‰« ", 13136, 1292, 1195, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol12", "«·’«›Ì", 14331, 1292, 1205, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtSum2", "=Sum([Basic])", 2381, 85, 1195, 312, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtSum3", "=Sum([Housing])", 3576, 85, 1195, 312, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtSum4", "=Sum([OtherAllow])", 4771, 85, 1195, 312, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtSum5", "=Sum([Overtime])", 5966, 85, 1195, 312, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtSum6", "=Sum([Additions])", 7161, 85, 1195, 312, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtSum7", "=Sum([Gross])", 8356, 85, 1195, 312, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtSum8", "=Sum([AbsenceDeduction])", 9551, 85, 1195, 312, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtSum9", "=Sum([AdvanceDeduction])", 10746, 85, 1195, 312, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtSum10", "=Sum([OtherDeduction])", 11941, 85, 1195, 312, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtSum11", "=Sum([GosiEmployee])", 13136, 85, 1195, 312, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtSum12", "=Sum([NetPay])", 14331, 85, 1205, 312, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtSumCaption", "=""«·≈Ã„«·Ì""", 0, 85, 2381, 312, 9, True, 0)
    Set c = RLabel(2, "lblSign1", "√⁄œ¯Â: ....................", 0, 737, 3591, 312, 10, False, 0)
    Set c = RLabel(2, "lblSign2", "—«Ã⁄Â: ....................", 3591, 737, 3591, 312, 10, False, 0)
    Set c = RLabel(2, "lblSign3", "«⁄ „œÂ: ....................", 7182, 737, 3591, 312, 10, False, 0)
    Set c = RText(0, "txtCol1", "EmployeeName", 0, 17, 2381, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "Basic", 2381, 17, 1195, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol3", "Housing", 3576, 17, 1195, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol4", "OtherAllow", 4771, 17, 1195, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "Overtime", 5966, 17, 1195, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "Additions", 7161, 17, 1195, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "Gross", 8356, 17, 1195, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "AbsenceDeduction", 9551, 17, 1195, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "AdvanceDeduction", 10746, 17, 1195, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "OtherDeduction", 11941, 17, 1195, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "GosiEmployee", 13136, 17, 1195, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol12", "NetPay", 14331, 17, 1205, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""«·„”Ì— »·« √”ÿ—.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptPayroll", s
    Exit Sub
EH:
    AbortReport "rptPayroll", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptDailySales()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptDailySales", "«·„»Ì⁄«  «·ÌÊ„Ì…", "DailySalesQuery", 15536, "", "SaleDate", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·„»Ì⁄«  «·ÌÊ„Ì…", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«· «—ÌŒ", 0, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·›Ê« Ì—", 1361, 1292, 907, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·„— Ã⁄« ", 2268, 1292, 907, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·„»Ì⁄«  »œÊ‰ ÷—Ì»…", 3175, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·„— Ã⁄«  »œÊ‰ ÷—Ì»…", 4649, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·’«›Ì »œÊ‰ ÷—Ì»…", 6123, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "«·÷—Ì»…", 7597, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "«·’«›Ì ‘«„· «·÷—Ì»…", 8844, 1292, 1531, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "‰ﬁœÌ", 10375, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "¬Ã·", 11736, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "«·„Õ’¯·", 13097, 1292, 2439, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal2", "=Sum([InvoiceCount])", 1361, 85, 907, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(2, "txtTotal3", "=Sum([ReturnCount])", 2268, 85, 907, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(2, "txtTotal4", "=Sum([SalesExVAT])", 3175, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([ReturnsExVAT])", 4649, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([NetSalesExVAT])", 6123, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal7", "=Sum([NetVAT])", 7597, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([NetSalesTotal])", 8844, 85, 1531, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal9", "=Sum([CashSales])", 10375, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal10", "=Sum([CreditSales])", 11736, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal11", "=Sum([CollectedAmount])", 13097, 85, 2439, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 1361, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=GDate([SaleDate])", 0, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "InvoiceCount", 1361, 17, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol3", "ReturnCount", 2268, 17, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol4", "SalesExVAT", 3175, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "ReturnsExVAT", 4649, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "NetSalesExVAT", 6123, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "NetVAT", 7597, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "NetSalesTotal", 8844, 17, 1531, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "CashSales", 10375, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "CreditSales", 11736, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "CollectedAmount", 13097, 17, 2439, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptDailySales", s
    Exit Sub
EH:
    AbortReport "rptDailySales", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptMonthlySales()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptMonthlySales", "«·„»Ì⁄«  «·‘Â—Ì…", "MonthlySalesQuery", 10773, "", "SalesYear,SalesMonth", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·„»Ì⁄«  «·‘Â—Ì…", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·‘Â—", 0, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·›Ê« Ì—", 1134, 1292, 964, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·„— Ã⁄« ", 2098, 1292, 964, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·’«›Ì »œÊ‰ ÷—Ì»…", 3062, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·÷—Ì»…", 4536, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "‘«„· «·÷—Ì»…", 5783, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "«· ﬂ·›…", 7257, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "„Ã„· «·—»Õ", 8731, 1292, 2042, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal2", "=Sum([InvoiceCount])", 1134, 85, 964, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(2, "txtTotal3", "=Sum([ReturnCount])", 2098, 85, 964, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(2, "txtTotal4", "=Sum([NetSalesExVAT])", 3062, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([NetVAT])", 4536, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([NetSalesTotal])", 5783, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal7", "=Sum([CostOfSales])", 7257, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([GrossProfit])", 8731, 85, 2042, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 1134, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=[SalesYear] & ""/"" & Format([SalesMonth],""00"")", 0, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "InvoiceCount", 1134, 17, 964, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol3", "ReturnCount", 2098, 17, 964, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol4", "NetSalesExVAT", 3062, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "NetVAT", 4536, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "NetSalesTotal", 5783, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "CostOfSales", 7257, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "GrossProfit", 8731, 17, 2042, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptMonthlySales", s
    Exit Sub
EH:
    AbortReport "rptMonthlySales", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptSalesByPeriod()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSalesByPeriod", "«·„»Ì⁄«  Õ”» › —…", "SalesByPeriodQuery", 15536, "", "DocDate,DocNumber", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·„»Ì⁄«  Õ”» › —…", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·‰Ê⁄", 0, 1292, 850, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·—ﬁ„", 850, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«· «—ÌŒ", 2268, 1292, 1644, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·⁄„Ì·", 3912, 1292, 2381, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·„ÊŸ›", 6293, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·œ›⁄", 7767, 1292, 794, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "»œÊ‰ ÷—Ì»…", 8561, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "«·÷—Ì»…", 9922, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "«·≈Ã„«·Ì", 11113, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "«·„œ›Ê⁄", 12531, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "⁄·Ï «·Õ”«»", 13949, 1292, 1587, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal7", "=Sum([NetAmount])", 8561, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([VATAmount])", 9922, 85, 1191, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal9", "=Sum([GrossAmount])", 11113, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal10", "=Sum([SettledAmount])", 12531, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal11", "=Sum([OnAccount])", 13949, 85, 1587, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 8561, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=IIf([DocType]=""RETURN"",""„— Ã⁄"",""›« Ê—…"")", 0, 17, 850, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "DocNumber", 850, 17, 1418, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "=GDate([DocDate],True)", 2268, 17, 1644, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "CustomerName", 3912, 17, 2381, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol5", "EmployeeName", 6293, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol6", "=IIf([PaymentType]=""CREDIT"",""¬Ã·"",""‰ﬁœÌ"")", 7767, 17, 794, 284, 8, False, 0)
    Set c = RText(0, "txtCol7", "NetAmount", 8561, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "VATAmount", 9922, 17, 1191, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "GrossAmount", 11113, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "SettledAmount", 12531, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "OnAccount", 13949, 17, 1587, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSalesByPeriod", s
    Exit Sub
EH:
    AbortReport "rptSalesByPeriod", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptSalesByProduct()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSalesByProduct", "«·„»Ì⁄«  Õ”» «·„‰ Ã", "SalesByProductQuery", 15536, "", "ProductName", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·„»Ì⁄«  Õ”» «·„‰ Ã", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·ﬂÊœ", 0, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·„‰ Ã", 1247, 1292, 3062, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«· ’‰Ì›", 4309, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·„»«⁄", 5783, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·„— Ã⁄", 6804, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·’«›Ì", 7825, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "«·„»Ì⁄«  »œÊ‰ ÷—Ì»…", 8846, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "«·÷—Ì»…", 10264, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "‘«„· «·÷—Ì»…", 11455, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "«· ﬂ·›…", 12873, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "«·—»Õ", 14177, 1292, 1359, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([QtySold])", 5783, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal5", "=Sum([QtyReturned])", 6804, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal6", "=Sum([NetQty])", 7825, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal7", "=Sum([NetSales])", 8846, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([SalesVAT])", 10264, 85, 1191, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal9", "=Sum([SalesTotal])", 11455, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal10", "=Sum([CostOfSales])", 12873, 85, 1304, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal11", "=Sum([GrossProfit])", 14177, 85, 1359, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 5783, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ProductCode", 0, 17, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductName", 1247, 17, 3062, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "CategoryName", 4309, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "QtySold", 5783, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "QtyReturned", 6804, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "NetQty", 7825, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol7", "NetSales", 8846, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "SalesVAT", 10264, 17, 1191, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "SalesTotal", 11455, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "CostOfSales", 12873, 17, 1304, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "GrossProfit", 14177, 17, 1359, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSalesByProduct", s
    Exit Sub
EH:
    AbortReport "rptSalesByProduct", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptBestSelling()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptBestSelling", "√›÷· «·„‰ Ã«  „»Ì⁄«", "BestSellingProductsQuery", 15536, "", "-NetQty,-NetSales", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "√›÷· «·„‰ Ã«  „»Ì⁄«", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·ﬂÊœ", 0, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·„‰ Ã", 1247, 1292, 3062, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«· ’‰Ì›", 4309, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·„»«⁄", 5783, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·„— Ã⁄", 6804, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·’«›Ì", 7825, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "«·„»Ì⁄«  »œÊ‰ ÷—Ì»…", 8846, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "«·÷—Ì»…", 10264, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "‘«„· «·÷—Ì»…", 11455, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "«· ﬂ·›…", 12873, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "«·—»Õ", 14177, 1292, 1359, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([QtySold])", 5783, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal5", "=Sum([QtyReturned])", 6804, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal6", "=Sum([NetQty])", 7825, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal7", "=Sum([NetSales])", 8846, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([SalesVAT])", 10264, 85, 1191, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal9", "=Sum([SalesTotal])", 11455, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal10", "=Sum([CostOfSales])", 12873, 85, 1304, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal11", "=Sum([GrossProfit])", 14177, 85, 1359, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 5783, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ProductCode", 0, 17, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductName", 1247, 17, 3062, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "CategoryName", 4309, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "QtySold", 5783, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "QtyReturned", 6804, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "NetQty", 7825, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol7", "NetSales", 8846, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "SalesVAT", 10264, 17, 1191, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "SalesTotal", 11455, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "CostOfSales", 12873, 17, 1304, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "GrossProfit", 14177, 17, 1359, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptBestSelling", s
    Exit Sub
EH:
    AbortReport "rptBestSelling", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptLeastSelling()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptLeastSelling", "√ﬁ· «·„‰ Ã«  „»Ì⁄«", "LeastSellingProductsQuery", 10773, "", "NetQtySold,ProductName", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "√ﬁ· «·„‰ Ã«  „»Ì⁄«", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·ﬂÊœ", 0, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·„‰ Ã", 1361, 1292, 3629, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«· ’‰Ì›", 4990, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·„Œ“Ê‰ «·Õ«·Ì", 6691, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "’«›Ì «·„»«⁄", 7995, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "’«›Ì «·„»Ì⁄« ", 9299, 1292, 1474, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([CurrentQuantity])", 6691, 85, 1304, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal5", "=Sum([NetQtySold])", 7995, 85, 1304, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal6", "=Sum([NetSalesAmount])", 9299, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 6691, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ProductCode", 0, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductName", 1361, 17, 3629, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "CategoryName", 4990, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "CurrentQuantity", 6691, 17, 1304, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "NetQtySold", 7995, 17, 1304, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "NetSalesAmount", 9299, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptLeastSelling", s
    Exit Sub
EH:
    AbortReport "rptLeastSelling", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptPurchases()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptPurchases", "«·„‘ —Ì« ", "PurchasesQuery", 15536, "", "DocDate,DocNumber", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·„‘ —Ì« ", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·‰Ê⁄", 0, 1292, 850, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·—ﬁ„", 850, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "›« Ê—… «·„Ê—œ", 2268, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«· «—ÌŒ", 3629, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·„Ê—œ", 4876, 1292, 2495, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·œ›⁄", 7371, 1292, 794, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "»œÊ‰ ÷—Ì»…", 8165, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "«·÷—Ì»…", 9583, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "«·≈Ã„«·Ì", 10830, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "«·„œ›Ê⁄", 12304, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "⁄·Ï «·Õ”«»", 13778, 1292, 1758, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal7", "=Sum([NetAmount])", 8165, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([VATAmount])", 9583, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal9", "=Sum([GrossAmount])", 10830, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal10", "=Sum([SettledAmount])", 12304, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal11", "=Sum([OnAccount])", 13778, 85, 1758, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 8165, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=IIf([DocType]=""RETURN"",""„— Ã⁄"",""›« Ê—…"")", 0, 17, 850, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "DocNumber", 850, 17, 1418, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "SupplierRef", 2268, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "=GDate([DocDate])", 3629, 17, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "SupplierName", 4876, 17, 2495, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol6", "=IIf([PaymentType]=""CREDIT"",""¬Ã·"",""‰ﬁœÌ"")", 7371, 17, 794, 284, 8, False, 0)
    Set c = RText(0, "txtCol7", "NetAmount", 8165, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "VATAmount", 9583, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "GrossAmount", 10830, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "SettledAmount", 12304, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "OnAccount", 13778, 17, 1758, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptPurchases", s
    Exit Sub
EH:
    AbortReport "rptPurchases", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptStockBalance()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptStockBalance", "«·„Œ“Ê‰ «·Õ«·Ì", "StockBalanceQuery", 15536, "", "CategoryName,ProductName", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·„Œ“Ê‰ «·Õ«·Ì", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·ﬂÊœ", 0, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·„‰ Ã", 1304, 1292, 3175, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«· ’‰Ì›", 4479, 1292, 1531, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·ÊÕœ…", 6010, 1292, 907, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·ﬂ„Ì…", 6917, 1292, 1077, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·Õœ «·√œ‰Ï", 7994, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "„ Ê”ÿ «· ﬂ·›…", 9015, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "”⁄— «·»Ì⁄", 10262, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "«·ﬁÌ„… »«· ﬂ·›…", 11453, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "«·ﬁÌ„… »”⁄— «·»Ì⁄", 12927, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "«·Õ«·…", 14401, 1292, 1135, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal5", "=Sum([CurrentQuantity])", 6917, 85, 1077, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal9", "=Sum([StockCostValue])", 11453, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal10", "=Sum([StockSalesValue])", 12927, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 6917, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ProductCode", 0, 17, 1304, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductName", 1304, 17, 3175, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "CategoryName", 4479, 17, 1531, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "UnitName", 6010, 17, 907, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "CurrentQuantity", 6917, 17, 1077, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "MinimumQuantity", 7994, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol7", "AverageCost", 9015, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "SellingPrice", 10262, 17, 1191, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "StockCostValue", 11453, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "StockSalesValue", 12927, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "=IIf([IsLowStock],""„‰Œ›÷"","""") & IIf([QuantityMismatch]<>0,"" / ›—ﬁ!"","""")", 14401, 17, 1135, 284, 8, False, 0)
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptStockBalance", s
    Exit Sub
EH:
    AbortReport "rptStockBalance", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptLowStock()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptLowStock", "«·„‰ Ã«  „‰Œ›÷… «·„Œ“Ê‰", "LowStockQuery", 10773, "", "-ShortageQty,ProductName", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·„‰ Ã«  „‰Œ›÷… «·„Œ“Ê‰", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·ﬂÊœ", 0, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·„‰ Ã", 1191, 1292, 2608, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«· ’‰Ì›", 3799, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·„ Ê›—", 5103, 1292, 964, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·Õœ «·√œ‰Ï", 6067, 1292, 964, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·‰ﬁ’", 7031, 1292, 907, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "«·„Ê—œ", 7938, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "ÃÊ«· «·„Ê—œ", 9639, 1292, 1134, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal6", "=Sum([ShortageQty])", 7031, 85, 907, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 7031, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ProductCode", 0, 17, 1191, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductName", 1191, 17, 2608, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "CategoryName", 3799, 17, 1304, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "CurrentQuantity", 5103, 17, 964, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "MinimumQuantity", 6067, 17, 964, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "ShortageQty", 7031, 17, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol7", "SupplierName", 7938, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol8", "SupplierMobile", 9639, 17, 1134, 284, 8, False, 0)
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ „‰ Ã«  „‰Œ›÷… «·„Œ“Ê‰.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptLowStock", s
    Exit Sub
EH:
    AbortReport "rptLowStock", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptProductMovement()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptProductMovement", "Õ—ﬂ… „‰ Ã", "ProductMovementQuery", 10773, "", "SortKey,MovementDate,TransactionID", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "Õ—ﬂ… „‰ Ã", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«· «—ÌŒ", 0, 1292, 1644, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·Õ—ﬂ…", 1644, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·„” ‰œ", 3118, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "Ê«—œ", 4536, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "’«œ—", 5557, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·—’Ìœ", 6578, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "«· ﬂ·›…", 7712, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "„·«ÕŸ« ", 8846, 1292, 1927, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([QtyIn])", 4536, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal5", "=Sum([QtyOut])", 5557, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 4536, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=GDate([MovementDate],True)", 0, 17, 1644, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "MovementType", 1644, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "ReferenceNumber", 3118, 17, 1418, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "QtyIn", 4536, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "QtyOut", 5557, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "NetQty", 6578, 17, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    SetCtl c, "RunningSum", 2
    Set c = RText(0, "txtCol7", "UnitCost", 7712, 17, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "Notes", 8846, 17, 1927, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptProductMovement", s
    Exit Sub
EH:
    AbortReport "rptProductMovement", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptCustomerStatement()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptCustomerStatement", "ﬂ‘› Õ”«» ⁄„Ì·", "CustomerStatementQuery", 10773, "", "SortKey,EntryDate", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "ﬂ‘› Õ”«» ⁄„Ì·", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«· «—ÌŒ", 0, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·»Ì«‰", 1701, 1292, 2041, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "—ﬁ„ «·„” ‰œ", 3742, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "„œÌ‰", 5443, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "œ«∆‰", 7144, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·—’Ìœ", 8845, 1292, 1928, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([Debit])", 5443, 85, 1701, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([Credit])", 7144, 85, 1701, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 5443, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=GDate([EntryDate],True)", 0, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "EntryTypeName", 1701, 17, 2041, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "DocNumber", 3742, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "Debit", 5443, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "Credit", 7144, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "=[Debit]-[Credit]", 8845, 17, 1928, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    SetCtl c, "RunningSum", 2
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptCustomerStatement", s
    Exit Sub
EH:
    AbortReport "rptCustomerStatement", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptSupplierStatement()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSupplierStatement", "ﬂ‘› Õ”«» „Ê—œ", "SupplierStatementQuery", 10773, "", "SortKey,EntryDate", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "ﬂ‘› Õ”«» „Ê—œ", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«· «—ÌŒ", 0, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·»Ì«‰", 1701, 1292, 2041, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "—ﬁ„ «·„” ‰œ", 3742, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "„œÌ‰", 5443, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "œ«∆‰", 7144, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·—’Ìœ", 8845, 1292, 1928, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([Debit])", 5443, 85, 1701, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([Credit])", 7144, 85, 1701, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 5443, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=GDate([EntryDate],True)", 0, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "EntryTypeName", 1701, 17, 2041, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "DocNumber", 3742, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "Debit", 5443, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "Credit", 7144, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "=[Credit]-[Debit]", 8845, 17, 1928, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    SetCtl c, "RunningSum", 2
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSupplierStatement", s
    Exit Sub
EH:
    AbortReport "rptSupplierStatement", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptExpenses()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptExpenses", "«·„’—Ê›«  ( ›’Ì·Ì)", "ExpensesQuery", 10773, "", "ExpenseDate,ExpenseNumber", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·„’—Ê›«  ( ›’Ì·Ì)", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·—ﬁ„", 0, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«· «—ÌŒ", 1247, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·‰Ê⁄", 2438, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·Ê’›", 3912, 1292, 2495, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·œ›⁄", 6407, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·„»·€", 7428, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "«·÷—Ì»…", 8562, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "«·≈Ã„«·Ì", 9583, 1292, 1190, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal6", "=Sum([Amount])", 7428, 85, 1134, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal7", "=Sum([Tax])", 8562, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([TotalAmount])", 9583, 85, 1190, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 7428, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ExpenseNumber", 0, 17, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "=GDate([ExpenseDate])", 1247, 17, 1191, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "ExpenseTypeName", 2438, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "Description", 3912, 17, 2495, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol5", "MethodName", 6407, 17, 1021, 284, 8, False, 0)
    Set c = RText(0, "txtCol6", "Amount", 7428, 17, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "Tax", 8562, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "TotalAmount", 9583, 17, 1190, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptExpenses", s
    Exit Sub
EH:
    AbortReport "rptExpenses", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptExpensesByType()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptExpensesByType", "«·„’—Ê›«  (≈Ã„«·Ì Õ”» «·‰Ê⁄)", "ExpensesByTypeQuery", 10773, "", "-AmountTotal", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·„’—Ê›«  (≈Ã„«·Ì Õ”» «·‰Ê⁄)", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "‰Ê⁄ «·„’—Ê›", 0, 1292, 3402, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·⁄œœ", 3402, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "»œÊ‰ ÷—Ì»…", 4763, 1292, 1928, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·÷—Ì»…", 6691, 1292, 1814, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·≈Ã„«·Ì", 8505, 1292, 2268, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal2", "=Sum([ExpenseCount])", 3402, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(2, "txtTotal3", "=Sum([AmountExVAT])", 4763, 85, 1928, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal4", "=Sum([InputVAT])", 6691, 85, 1814, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([AmountTotal])", 8505, 85, 2268, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 3402, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ExpenseTypeName", 0, 17, 3402, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ExpenseCount", 3402, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol3", "AmountExVAT", 4763, 17, 1928, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol4", "InputVAT", 6691, 17, 1814, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "AmountTotal", 8505, 17, 2268, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptExpensesByType", s
    Exit Sub
EH:
    AbortReport "rptExpensesByType", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptCashStatement()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptCashStatement", "Õ—ﬂ… «·Œ“Ì‰… / «·’‰œÊﬁ ( ›’Ì·Ì)", "CashStatementQuery", 15536, "", "SortKey,MoveDate", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 2043
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "Õ—ﬂ… «·Œ“Ì‰… / «·’‰œÊﬁ ( ›’Ì·Ì)", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«· «—ÌŒ", 0, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·Õ—ﬂ…", 1701, 1292, 2155, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·„” ‰œ", 3856, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·ÃÂ…", 5330, 1292, 2268, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·»Ì«‰", 7598, 1292, 2155, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·’‰œÊﬁ", 9753, 1292, 1588, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "„ﬁ»Ê÷", 11341, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "„œ›Ê⁄", 12702, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "«·—’Ìœ", 14063, 1292, 1473, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 15536, 284, 8, True, 0)
    Set c = RLabel(2, "lblSum1", "—’Ìœ √Ê· «·„œ…", 9866, 539, 3118, 312, 11, True, 0)
    Set c = RText(2, "txtSum1", "=Sum(IIf([SortKey]=0,[AmountIn]-[AmountOut],0))", 12984, 539, 2552, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(2, "lblSum2", "«·„ﬁ»Ê÷« ", 9866, 908, 3118, 312, 11, True, 0)
    Set c = RText(2, "txtSum2", "=Sum(IIf([SortKey]=1,[AmountIn],0))", 12984, 908, 2552, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(2, "lblSum3", "«·„œ›Ê⁄«  Ê«·„’—Ê›« ", 9866, 1277, 3118, 312, 11, True, 0)
    Set c = RText(2, "txtSum3", "=Sum(IIf([SortKey]=1,[AmountOut],0))", 12984, 1277, 2552, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(2, "lblSum4", "—’Ìœ ¬Œ— «·„œ…", 9866, 1646, 3118, 312, 11, True, 0)
    Set c = RText(2, "txtSum4", "=Sum([AmountIn]-[AmountOut])", 12984, 1646, 2552, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol1", "=GDate([MoveDate],True)", 0, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "MoveTypeName", 1701, 17, 2155, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "DocNumber", 3856, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "PartyName", 5330, 17, 2268, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol5", "Details", 7598, 17, 2155, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol6", "BoxName", 9753, 17, 1588, 284, 8, False, 0)
    Set c = RText(0, "txtCol7", "AmountIn", 11341, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "AmountOut", 12702, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "=[AmountIn]-[AmountOut]", 14063, 17, 1473, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    SetCtl c, "RunningSum", 2
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ Õ—ﬂ… ‰ﬁœÌ… ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptCashStatement", s
    Exit Sub
EH:
    AbortReport "rptCashStatement", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptCashDaily()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptCashDaily", "Õ—ﬂ… «·Œ“Ì‰… «·ÌÊ„Ì… (√Ê· «·ÌÊ„ Ê¬Œ—Â)", "CashDailyQuery", 10773, "", "CashDay", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "Õ—ﬂ… «·Œ“Ì‰… «·ÌÊ„Ì… (√Ê· «·ÌÊ„ Ê¬Œ—Â)", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·ÌÊ„", 0, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "—’Ìœ √Ê· «·ÌÊ„", 1701, 1292, 1928, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·„ﬁ»Ê÷« ", 3629, 1292, 1928, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·„œ›Ê⁄« ", 5557, 1292, 1928, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "—’Ìœ ¬Œ— «·ÌÊ„", 7485, 1292, 2041, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·Õ—ﬂ« ", 9526, 1292, 1247, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal3", "=Sum([Receipts])", 3629, 85, 1928, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal4", "=Sum([Payments])", 5557, 85, 1928, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([MoveCount])", 9526, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 3629, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=GDate([CashDay])", 0, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "OpeningBalance", 1701, 17, 1928, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol3", "Receipts", 3629, 17, 1928, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol4", "Payments", 5557, 17, 1928, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "ClosingBalance", 7485, 17, 2041, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "MoveCount", 9526, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ Õ—ﬂ… ‰ﬁœÌ… ›Ì Â–Â «·› —….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptCashDaily", s
    Exit Sub
EH:
    AbortReport "rptCashDaily", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptCostCenterProfit()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptCostCenterProfit", "ﬁ«∆„… «·œŒ· Õ”» „—ﬂ“ «· ﬂ·›…", "CostCenterProfitQuery", 15536, "", "CenterCode", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "ﬁ«∆„… «·œŒ· Õ”» „—ﬂ“ «· ﬂ·›…", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·—„“", 0, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·„—ﬂ“", 1134, 1292, 3062, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·≈Ì—«œ« ", 4196, 1292, 1814, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", " ﬂ·›… «·„»Ì⁄« ", 6010, 1292, 1814, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "„Ã„· «·—»Õ", 7824, 1292, 1814, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·„’—Ê›« ", 9638, 1292, 1814, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "’«›Ì «·—»Õ", 11452, 1292, 4084, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal3", "=Sum([Revenue])", 4196, 85, 1814, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal4", "=Sum([CostOfSales])", 6010, 85, 1814, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([GrossProfit])", 7824, 85, 1814, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([Expenses])", 9638, 85, 1814, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal7", "=Sum([NetProfit])", 11452, 85, 4084, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 4196, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "CenterCode", 0, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "CenterName", 1134, 17, 3062, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "Revenue", 4196, 17, 1814, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol4", "CostOfSales", 6010, 17, 1814, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "GrossProfit", 7824, 17, 1814, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "Expenses", 9638, 17, 1814, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "NetProfit", 11452, 17, 4084, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ ≈Ì—«œ«  √Ê „’—Ê›«  ›Ì «·› —….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptCostCenterProfit", s
    Exit Sub
EH:
    AbortReport "rptCostCenterProfit", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptCostCenterAccounts()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptCostCenterAccounts", "≈Ì—«œ«  Ê„’—Ê›«  ﬂ· „—ﬂ“  ﬂ·›…", "CostCenterAccountsQuery", 10773, "", "CenterKey,TreeKey", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "≈Ì—«œ«  Ê„’—Ê›«  ﬂ· „—ﬂ“  ﬂ·›…", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·„—ﬂ“", 0, 1292, 2268, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·ﬁ”„", 2268, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·Õ”«»", 3969, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«”„ «·Õ”«»", 5103, 1292, 3402, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·„»·€", 8505, 1292, 2268, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 10773, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "CenterName", 0, 17, 2268, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "SectionName", 2268, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "AccountCode", 3969, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "AccountName", 5103, 17, 3402, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol5", "CenterAmount", 8505, 17, 2268, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ ≈Ì—«œ«  √Ê „’—Ê›«  ›Ì «·› —….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptCostCenterAccounts", s
    Exit Sub
EH:
    AbortReport "rptCostCenterAccounts", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptBudgetVsActual()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptBudgetVsActual", "«·„Ê«“‰… „ﬁ«»· «·›⁄·Ì", "BudgetVsActualQuery", 15536, "", "-AccountType,TreeKey", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·„Ê«“‰… „ﬁ«»· «·›⁄·Ì", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·»‰œ", 0, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·Õ”«»", 1361, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«”„ «·Õ”«»", 2382, 1292, 3062, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·„—ﬂ“", 5444, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·„Ê«“‰…", 7145, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·›⁄·Ì", 8846, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "«·«‰Õ—«›", 10547, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "«·‰”»…", 12248, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "«· ﬁÌÌ„", 13269, 1292, 2267, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal5", "=Sum([BudgetAmount])", 7145, 85, 1701, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([ActualAmount])", 8846, 85, 1701, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 7145, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "SectionName", 0, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "AccountCode", 1361, 17, 1021, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "AccountName", 2382, 17, 3062, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol4", "BudgetCenter", 5444, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "BudgetAmount", 7145, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "ActualAmount", 8846, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "Variance", 10547, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "VariancePct", 12248, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "0.0%"
    Set c = RText(0, "txtCol9", "VarianceNote", 13269, 17, 2267, 284, 8, False, 0)
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ „Ê«“‰… ·”‰… «·› —….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptBudgetVsActual", s
    Exit Sub
EH:
    AbortReport "rptBudgetVsActual", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptAuditTrail()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptAuditTrail", "”Ã· «· œﬁÌﬁ", "AuditTrailQuery", 15536, "", "-LogID,LineNo", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "”Ã· «· œﬁÌﬁ", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·Êﬁ ", 0, 1292, 1644, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·„” Œœ„", 1644, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·⁄„·Ì…", 3118, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·ÃœÊ·", 4479, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·—ﬁ„", 5953, 1292, 794, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·”Ã·", 6747, 1292, 1814, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "«·Õﬁ·", 8561, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "ﬁ»·", 10262, 1292, 2041, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "»⁄œ", 12303, 1292, 3233, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 15536, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=GDate([LogDate], True)", 0, 17, 1644, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "UserName", 1644, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "ActionLabel", 3118, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "ObjectName", 4479, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "RecordID", 5953, 17, 794, 284, 8, False, 0)
    Set c = RText(0, "txtCol6", "RecordLabel", 6747, 17, 1814, 284, 8, False, 0)
    Set c = RText(0, "txtCol7", "FieldCaption", 8561, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol8", "OldValue", 10262, 17, 2041, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol9", "NewValue", 12303, 17, 3233, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ ⁄„·Ì«  ›Ì «·› —….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptAuditTrail", s
    Exit Sub
EH:
    AbortReport "rptAuditTrail", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptFixedAssets()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptFixedAssets", "”Ã· «·√’Ê· «·À«» …", "FixedAssetsQuery", 15536, "", "Status,AssetCode", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "”Ã· «·√’Ê· «·À«» …", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·—ﬁ„", 0, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·√’·", 1134, 1292, 2948, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·„Ã„Ê⁄…", 4082, 1292, 2041, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·‘—«¡", 6123, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·⁄„— (‘Â—)", 7370, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«· ﬂ·›…", 8391, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "«·ﬁ”ÿ «·‘Â—Ì", 9865, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "„Ã„⁄ «·≈Â·«ﬂ", 11226, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "«·ﬁÌ„… «·œ› —Ì…", 12700, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "«·Õ«·…", 14174, 1292, 1362, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal6", "=Sum([Cost])", 8391, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal7", "=Sum([MonthlyDep])", 9865, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([AccumDep])", 11226, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal9", "=Sum([BookValue])", 12700, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 8391, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "AssetCode", 0, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "AssetName", 1134, 17, 2948, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "AssetGroup", 4082, 17, 2041, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "=GDate([PurchaseDate])", 6123, 17, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "UsefulLifeMonths", 7370, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol6", "Cost", 8391, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "MonthlyDep", 9865, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "AccumDep", 11226, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "BookValue", 12700, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "StatusName", 14174, 17, 1362, 284, 8, False, 0)
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ √’Ê· À«» ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptFixedAssets", s
    Exit Sub
EH:
    AbortReport "rptFixedAssets", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptCashBalances()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptCashBalances", "√—’œ… «·Œ“Ì‰… Ê«·’‰«œÌﬁ", "CashBoxBalanceQuery", 10773, "", "-BoxType,BoxName", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "√—’œ… «·Œ“Ì‰… Ê«·’‰«œÌﬁ", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·’‰œÊﬁ", 0, 1292, 2835, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·‰Ê⁄", 2835, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "≈Ã„«·Ì «·œ«Œ·", 4309, 1292, 1644, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "≈Ã„«·Ì «·Œ«—Ã", 5953, 1292, 1644, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·—’Ìœ", 7597, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "¬Œ— Õ—ﬂ…", 9298, 1292, 1475, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal3", "=Sum([TotalIn])", 4309, 85, 1644, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal4", "=Sum([TotalOut])", 5953, 85, 1644, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([Balance])", 7597, 85, 1701, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 4309, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "BoxName", 0, 17, 2835, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol2", "BoxTypeName", 2835, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "TotalIn", 4309, 17, 1644, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol4", "TotalOut", 5953, 17, 1644, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "Balance", 7597, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "=GDate([LastMoveDate])", 9298, 17, 1475, 284, 8, False, 0)
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptCashBalances", s
    Exit Sub
EH:
    AbortReport "rptCashBalances", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptCashClosings()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptCashClosings", " ’›Ì«  ÌÊ„Ì… «·ﬂ«‘Ì—", "CashClosingsQuery", 15536, "", "ClosingDate", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", " ’›Ì«  ÌÊ„Ì… «·ﬂ«‘Ì—", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·—ﬁ„", 0, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«· «—ÌŒ", 1361, 1292, 1644, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·’‰œÊﬁ", 3005, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "√Ã—«Â«", 4706, 1292, 1588, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·œ› —Ì", 6294, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·›⁄·Ì", 7768, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "«·›—ﬁ", 9242, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "«· —ÕÌ·", 10489, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "«·„—ÕÛ¯·", 12190, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "«·„ »ﬁÌ", 13664, 1292, 1872, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal5", "=Sum([ExpectedBalance])", 6294, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([CountedAmount])", 7768, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal7", "=Sum([Difference])", 9242, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal9", "=Sum([TransferAmount])", 12190, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal10", "=Sum([KeptAmount])", 13664, 85, 1872, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 6294, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ClosingNumber", 0, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "=GDate([ClosingDate],True)", 1361, 17, 1644, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "BoxName", 3005, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "EmployeeName", 4706, 17, 1588, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "ExpectedBalance", 6294, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "CountedAmount", 7768, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "Difference", 9242, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "DestinationName", 10489, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol9", "TransferAmount", 12190, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "KeptAmount", 13664, 17, 1872, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ  ’›Ì«  ›Ì Â–Â «·› —….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptCashClosings", s
    Exit Sub
EH:
    AbortReport "rptCashClosings", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptJournal()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptJournal", "ﬁÌÊœ «·ÌÊ„Ì…", "JournalLinesQuery", 15536, "", "EntryDate,EntryNumber,LineNumber", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "ﬁÌÊœ «·ÌÊ„Ì…", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "—ﬁ„ «·ﬁÌœ", 0, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«· «—ÌŒ", 1361, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·⁄„·Ì…", 2608, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·„” ‰œ", 4309, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·Õ”«»", 5783, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«”„ «·Õ”«»", 6804, 1292, 2495, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "«·»Ì«‰", 9299, 1292, 2835, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "„œÌ‰", 12134, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "œ«∆‰", 13608, 1292, 1928, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal8", "=Sum(IIf([Debit]=0,Null,[Debit]))", 12134, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal9", "=Sum(IIf([Credit]=0,Null,[Credit]))", 13608, 85, 1928, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 12134, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "EntryNumber", 0, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "=GDate([EntryDate])", 1361, 17, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "TypeName", 2608, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "SourceNumber", 4309, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "AccountCode", 5783, 17, 1021, 284, 8, False, 0)
    Set c = RText(0, "txtCol6", "AccountName", 6804, 17, 2495, 284, 8, False, 0)
    Set c = RText(0, "txtCol7", "LineText", 9299, 17, 2835, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol8", "=IIf([Debit]=0,Null,[Debit])", 12134, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "=IIf([Credit]=0,Null,[Credit])", 13608, 17, 1928, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ ﬁÌÊœ ›Ì Â–Â «·› —….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptJournal", s
    Exit Sub
EH:
    AbortReport "rptJournal", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptTrialBalance()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptTrialBalance", "„Ì“«‰ «·„—«Ã⁄…", "TrialBalanceQuery", 10773, "", "AccountCode", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 936
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "„Ì“«‰ «·„—«Ã⁄…", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·Õ”«»", 0, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«”„ «·Õ”«»", 1134, 1292, 3175, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "—’Ìœ √Ê· «·„œ…", 4309, 1292, 1588, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "„œÌ‰ «·› —…", 5897, 1292, 1588, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "œ«∆‰ «·› —…", 7485, 1292, 1588, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·—’Ìœ «·Œ «„Ì", 9073, 1292, 1700, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal3", "=Sum([OpeningBalance])", 4309, 85, 1588, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal4", "=Sum([PeriodDebit])", 5897, 85, 1588, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([PeriodCredit])", 7485, 85, 1588, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([ClosingBalance])", 9073, 85, 1700, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 4309, 284, 8, True, 0)
    Set c = RLabel(2, "lblSum1", "„Ã„Ê⁄ «·√—’œ… (’›— = „ Ê«“‰)", 5103, 539, 3118, 312, 11, True, 0)
    Set c = RText(2, "txtSum1", "=Sum([ClosingBalance])", 8221, 539, 2552, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol1", "AccountCode", 0, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "AccountName", 1134, 17, 3175, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "OpeningBalance", 4309, 17, 1588, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol4", "PeriodDebit", 5897, 17, 1588, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "PeriodCredit", 7485, 17, 1588, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "ClosingBalance", 9073, 17, 1700, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ ﬁÌÊœ.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptTrialBalance", s
    Exit Sub
EH:
    AbortReport "rptTrialBalance", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptTrialBalanceTree()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptTrialBalanceTree", "„Ì“«‰ «·„—«Ã⁄… »«·„” ÊÌ« ", "TrialBalanceTreeQuery", 10773, "", "TreeKey", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 1674
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "„Ì“«‰ «·„—«Ã⁄… »«·„” ÊÌ« ", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·Õ”«»", 0, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«”„ «·Õ”«»", 1134, 1292, 3629, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "—’Ìœ √Ê· «·„œ…", 4763, 1292, 1531, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "„œÌ‰ «·› —…", 6294, 1292, 1531, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "œ«∆‰ «·› —…", 7825, 1292, 1531, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·—’Ìœ «·Œ «„Ì", 9356, 1292, 1417, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 10773, 284, 8, True, 0)
    Set c = RLabel(2, "lblSum1", "„œÌ‰ «·› —… («·„” ÊÏ «·√Ê·)", 5103, 539, 3118, 312, 11, True, 0)
    Set c = RText(2, "txtSum1", "=Sum(IIf([AccountLevel]=1,[PeriodDebit],0))", 8221, 539, 2552, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(2, "lblSum2", "œ«∆‰ «·› —… («·„” ÊÏ «·√Ê·)", 5103, 908, 3118, 312, 11, True, 0)
    Set c = RText(2, "txtSum2", "=Sum(IIf([AccountLevel]=1,[PeriodCredit],0))", 8221, 908, 2552, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(2, "lblSum3", "„Ã„Ê⁄ √—’œ… «·„” ÊÏ «·√Ê· (’›— = „ Ê«“‰)", 5103, 1277, 3118, 312, 11, True, 0)
    Set c = RText(2, "txtSum3", "=Sum(IIf([AccountLevel]=1,[ClosingBalance],0))", 8221, 1277, 2552, 312, 11, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol1", "AccountCode", 0, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "=Space(([AccountLevel]-1)*3) & [AccountName]", 1134, 17, 3629, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "OpeningBalance", 4763, 17, 1531, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol4", "PeriodDebit", 6294, 17, 1531, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "PeriodCredit", 7825, 17, 1531, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "ClosingBalance", 9356, 17, 1417, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ ﬁÌÊœ.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptTrialBalanceTree", s
    Exit Sub
EH:
    AbortReport "rptTrialBalanceTree", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptAccountTree()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptAccountTree", "œ·Ì· «·Õ”«»«  (‘Ã—… «·Õ”«»« )", "AccountTreeQuery", 10773, "", "TreeKey", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "œ·Ì· «·Õ”«»«  (‘Ã—… «·Õ”«»« )", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "—ﬁ„ «·Õ”«»", 0, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·Õ”«»", 1361, 1292, 4536, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·„” ÊÏ", 5897, 1292, 907, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·‰Ê⁄", 6804, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "—∆Ì”Ì / ›—⁄Ì", 8165, 1292, 2608, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 10773, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "AccountCode", 0, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "=Space(([AccountLevel]-1)*3) & [AccountName]", 1361, 17, 4536, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "AccountLevel", 5897, 17, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol4", "TypeName", 6804, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "KindName", 8165, 17, 2608, 284, 8, False, 0)
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ Õ”«»« .""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptAccountTree", s
    Exit Sub
EH:
    AbortReport "rptAccountTree", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptSlowMoving()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSlowMoving", "«·„‰ Ã«  €Ì— «·„ Õ—ﬂ…", "SlowMovingProductsQuery", 10773, "", "-DaysWithoutSale", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·„‰ Ã«  €Ì— «·„ Õ—ﬂ…", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·ﬂÊœ", 0, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·„‰ Ã", 1247, 1292, 2835, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«· ’‰Ì›", 4082, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·ﬂ„Ì…", 5500, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "„ Ê”ÿ «· ﬂ·›…", 6521, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·ﬁÌ„…", 7655, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "¬Œ— »Ì⁄", 8902, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "√Ì«„", 10036, 1292, 737, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([CurrentQuantity])", 5500, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal6", "=Sum([StockCostValue])", 7655, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 5500, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ProductCode", 0, 17, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductName", 1247, 17, 2835, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "CategoryName", 4082, 17, 1418, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "CurrentQuantity", 5500, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "AverageCost", 6521, 17, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "StockCostValue", 7655, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "=GDate([LastSaleDate])", 8902, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol8", "DaysWithoutSale", 10036, 17, 737, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ „‰ Ã«  —«ﬂœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSlowMoving", s
    Exit Sub
EH:
    AbortReport "rptSlowMoving", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptStockByCategory()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptStockByCategory", "«·„Œ“Ê‰ Õ”» «· ’‰Ì›", "StockByCategoryQuery", 10773, "", "CategoryName", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·„Œ“Ê‰ Õ”» «· ’‰Ì›", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«· ’‰Ì›", 0, 1292, 3402, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·„‰ Ã« ", 3402, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·ﬂ„Ì…", 4763, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·ﬁÌ„… »«· ﬂ·›…", 6464, 1292, 2041, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·ﬁÌ„… »”⁄— «·»Ì⁄", 8505, 1292, 2268, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal2", "=Sum([ProductCount])", 3402, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(2, "txtTotal3", "=Sum([TotalQuantity])", 4763, 85, 1701, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal4", "=Sum([StockCostValue])", 6464, 85, 2041, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([StockSalesValue])", 8505, 85, 2268, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 3402, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "CategoryName", 0, 17, 3402, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductCount", 3402, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol3", "TotalQuantity", 4763, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol4", "StockCostValue", 6464, 17, 2041, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "StockSalesValue", 8505, 17, 2268, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptStockByCategory", s
    Exit Sub
EH:
    AbortReport "rptStockByCategory", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptCustomerBalances()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptCustomerBalances", "√—’œ… «·⁄„·«¡", "CustomerBalanceQuery", 10773, "", "CustomerName", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "√—’œ… «·⁄„·«¡", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·⁄„Ì·", 0, 1292, 2835, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·ÃÊ«·", 2835, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "Õœ «·«∆ „«‰", 4196, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "„œÌ‰", 5443, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "œ«∆‰", 6804, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·—’Ìœ", 8165, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "¬Œ— Õ—ﬂ…", 9639, 1292, 1134, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([DebitTotal])", 5443, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([CreditTotal])", 6804, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([Balance])", 8165, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 5443, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "CustomerName", 0, 17, 2835, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol2", "Mobile", 2835, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "CreditLimit", 4196, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol4", "DebitTotal", 5443, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "CreditTotal", 6804, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "Balance", 8165, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "=GDate([LastEntryDate])", 9639, 17, 1134, 284, 8, False, 0)
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptCustomerBalances", s
    Exit Sub
EH:
    AbortReport "rptCustomerBalances", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptSupplierBalances()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSupplierBalances", "√—’œ… «·„Ê—œÌ‰", "SupplierBalanceQuery", 10773, "", "SupplierName", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "√—’œ… «·„Ê—œÌ‰", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·„Ê—œ", 0, 1292, 2835, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·„”ƒÊ·", 2835, 1292, 1644, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·ÃÊ«·", 4479, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "„œÌ‰", 5783, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "œ«∆‰", 7030, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·—’Ìœ", 8277, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "¬Œ— Õ—ﬂ…", 9638, 1292, 1135, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([DebitTotal])", 5783, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([CreditTotal])", 7030, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([Balance])", 8277, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 5783, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "SupplierName", 0, 17, 2835, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol2", "ContactPerson", 2835, 17, 1644, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "Mobile", 4479, 17, 1304, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "DebitTotal", 5783, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "CreditTotal", 7030, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "Balance", 8277, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "=GDate([LastEntryDate])", 9638, 17, 1135, 284, 8, False, 0)
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSupplierBalances", s
    Exit Sub
EH:
    AbortReport "rptSupplierBalances", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptIntegrityCheck()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptIntegrityCheck", "›Õ’ ”·«„… «·»Ì«‰« ", "IntegrityCheckQuery", 10773, "", "IssueCode,RecordID", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "›Õ’ ”·«„… «·»Ì«‰« ", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "«·—„“", 0, 1292, 1928, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "«·„‘ﬂ·…", 1928, 1292, 3856, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "«·ÃœÊ·", 5784, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "«·”Ã·", 7485, 1292, 907, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "«·„ Êﬁ⁄", 8392, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "«·›⁄·Ì", 9583, 1292, 1190, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""«·≈Ã„«·Ì ("" & Count(*) & "" ”Ã·)""", 0, 85, 10773, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "IssueCode", 0, 17, 1928, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "IssueText", 1928, 17, 3856, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "SourceTable", 5784, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "RecordID", 7485, 17, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol5", "ExpectedValue", 8392, 17, 1191, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "ActualValue", 9583, 17, 1190, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""›Õ’ ”·«„… «·»Ì«‰« : ·«  ÊÃœ √Ì „‘ﬂ·« .""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptIntegrityCheck", s
    Exit Sub
EH:
    AbortReport "rptIntegrityCheck", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptProfit()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptProfit", "«·√—»«Õ", "ProfitQuery", 10773, "", "", False, True
    SetSection 3, 1304
    SetSection 4, 340
    SetSection 2, 624
    SetSection 0, 3088
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "«·√—»«Õ", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RLabel(2, "lblNote", "’«›Ì «·—»Õ = „Ã„· «·—»Õ ± ›—Êﬁ«  «·„Œ“Ê‰ - «·„’—Ê›« . «·—’Ìœ «·«›  «ÕÌ ·« ÌœŒ· ›Ì «·—»Õ.", 0, 113, 10773, 454, 9, False, 2)
    Set c = RLabel(0, "lblRow1", "’«›Ì «·„»Ì⁄«  (»œÊ‰ «·÷—Ì»…)", 1134, 57, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow1", "NetSales", 6577, 57, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow2", " ﬂ·›… «·»÷«⁄… «·„»«⁄…", 1134, 482, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow2", "CostOfSales", 6577, 482, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow3", "„Ã„· «·—»Õ", 1134, 907, 5386, 340, 11, True, 0)
    Set c = RText(0, "txtRow3", "GrossProfit", 6577, 907, 3062, 340, 12, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow4", "‰”»… „Ã„· «·—»Õ", 1134, 1332, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow4", "=IIf([NetSales]=0,0,[GrossProfit]/[NetSales])", 6577, 1332, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "0.0%"
    Set c = RLabel(0, "lblRow5", "›—Êﬁ«  «·„Œ“Ê‰ (Ã—œ° ≈÷«›…° Œ’„)", 1134, 1757, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow5", "InventoryAdjustments", 6577, 1757, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow6", "«·„’—Ê›«  (»œÊ‰ «·÷—Ì»…)", 1134, 2182, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow6", "TotalExpenses", 6577, 2182, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow7", "’«›Ì «·—»Õ", 1134, 2607, 5386, 340, 11, True, 0)
    Set c = RText(0, "txtRow7", "NetProfit", 6577, 2607, 3062, 340, 12, True, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptProfit", s
    Exit Sub
EH:
    AbortReport "rptProfit", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptVatSummary()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptVatSummary", "„·Œ’ ÷—Ì»… «·ﬁÌ„… «·„÷«›…", "VatSummaryQuery", 10773, "", "", False, True
    SetSection 3, 1304
    SetSection 4, 340
    SetSection 2, 624
    SetSection 0, 3088
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "„·Œ’ ÷—Ì»… «·ﬁÌ„… «·„÷«›…", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""’›Õ… "" & [Page] & "" „‰ "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RLabel(2, "lblNote", "«·„ÊÃ» „” Õﬁ ··ÂÌ∆…° Ê«·”«·» —’Ìœ „” —œ. —«Ã⁄ «·√—ﬁ«„ „⁄ „Õ«”»ﬂ ﬁ»·  ﬁœÌ„ «·≈ﬁ—«—.", 0, 113, 10773, 454, 9, False, 2)
    Set c = RLabel(0, "lblRow1", "«·„»Ì⁄«  «·Œ«÷⁄… ··÷—Ì»… (»⁄œ «·„— Ã⁄« )", 1134, 57, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow1", "TaxableSales", 6577, 57, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow2", "÷—Ì»… «·„Œ—Ã« ", 1134, 482, 5386, 340, 11, True, 0)
    Set c = RText(0, "txtRow2", "OutputVAT", 6577, 482, 3062, 340, 12, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow3", "«·„‘ —Ì«  «·Œ«÷⁄… ··÷—Ì»… (»⁄œ «·„— Ã⁄« )", 1134, 907, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow3", "TaxablePurchases", 6577, 907, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow4", "÷—Ì»… „œŒ·«  «·„‘ —Ì« ", 1134, 1332, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow4", "PurchaseVAT", 6577, 1332, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow5", "÷—Ì»… „œŒ·«  «·„’—Ê›« ", 1134, 1757, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow5", "ExpenseVAT", 6577, 1757, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow6", "≈Ã„«·Ì ÷—Ì»… «·„œŒ·« ", 1134, 2182, 5386, 340, 11, True, 0)
    Set c = RText(0, "txtRow6", "InputVAT", 6577, 2182, 3062, 340, 12, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow7", "’«›Ì «·÷—Ì»… «·„” Õﬁ…", 1134, 2607, 5386, 340, 11, True, 0)
    Set c = RText(0, "txtRow7", "NetVATDue", 6577, 2607, 3062, 340, 12, True, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""·«  ÊÃœ »Ì«‰«  ·⁄—÷Â« ›Ì Â–« «· ﬁ—Ì— ··«Œ Ì«—«  «·„Õœœ….""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptVatSummary", s
    Exit Sub
EH:
    AbortReport "rptVatSummary", Err.Number, Err.Description
End Sub
