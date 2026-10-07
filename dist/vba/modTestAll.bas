Attribute VB_Name = "modTestAll"
'==============================================================================
' modTestAll  -  Retail Store Management System (Phase 11)
'
'   RunAllTests   runs every in-Access test in order and shows ONE summary:
'                 VerifySchema, TestRelationships, TestQueries (empty database
'                 only), TestForms, TestSales, TestPurchases, TestReports,
'                 TestDashboard, TestSecurity, TestLabels, TestTouchPOS, TestCash, TestJournal, TestAging, TestBank, TestCheques, TestAssets, TestPayroll, TestCostCenters, TestBudget and VerifyDemoData right after the
'                 demo data was loaded. Each test leaves no data behind.
'                 The summary is also written to the Immediate window (Ctrl+G).
'==============================================================================
Option Compare Database
Option Explicit

Public Function RunAllTests() As Boolean
    Dim names As Variant, i As Long, ok As Boolean, passed As Long, failed As Long, skipped As Long
    Dim lines As String, started As Single, result As Variant, demoState As String
    Calendar = vbCalGreg
    EnsureTestUser
    started = Timer
    g_TestSummary = ""
    g_CollectTests = True
    DoCmd.Hourglass True
    names = Array("VerifySchema", "TestRelationships", "TestQueries", "TestForms", "TestSales", "TestPurchases", _
                  "TestReports", "TestDashboard", "TestSecurity", "TestLabels", "TestTouchPOS", "TestCash", "TestJournal", "TestAging", "TestBank", "TestCheques", "TestAssets", "TestPayroll", "TestCostCenters", "TestBudget", _
                  "VerifyDemoData")
    For i = LBound(names) To UBound(names)
        If SkipReason(CStr(names(i))) <> "" Then
            lines = lines & "[--] " & names(i) & ": " & SkipReason(CStr(names(i))) & vbCrLf
            skipped = skipped + 1
        Else
            ok = False
            On Error Resume Next
            Err.Clear
            result = Application.Run(CStr(names(i)))
            If Err.Number <> 0 Then
                lines = lines & "[X]  " & names(i) & ": ·„ Ì⁄„· (" & Err.Description & ")" & vbCrLf
                Err.Clear
            Else
                ok = (result = True)
                lines = lines & IIf(ok, "[OK] ", "[X]  ") & names(i) & vbCrLf
            End If
            On Error GoTo 0
            If ok Then passed = passed + 1 Else failed = failed + 1
        End If
        g_SilentMode = False
    Next
    g_CollectTests = False
    DoCmd.Hourglass False
    Debug.Print "=== RunAllTests ===" & vbCrLf & lines & vbCrLf & g_TestSummary
    WriteTestLog "=== RunAllTests " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ===" & vbCrLf & lines & vbCrLf & _
                 g_TestSummary
    MsgBox IIf(failed = 0, "ﬂ· «·«Œ »«—«  ‰«ÃÕ….", "ÌÊÃœ " & failed & " «Œ »«— ›«‘·.") & vbCrLf & _
           "‰«ÃÕ: " & passed & "   ›«‘·: " & failed & "   „ ŒÿÏ: " & skipped & "   «·„œ…: " & _
           Format$(Timer - started, "0") & " À" & vbCrLf & vbCrLf & lines & vbCrLf & _
           "«· ›«’Ì· ›Ì ‰«›–… Immediate (Ctrl+G) Ê›Ì «·„·› RunAllTests.log »Ã«‰» «·»—‰«„Ã.", IIf(failed = 0, vbInformation, vbExclamation) + MSG_RTL, _
           "RunAllTests"
    RunAllTests = (failed = 0)
End Function

Private Sub WriteTestLog(ByVal Text As String)
    ' RunAllTests.log next to the front-end (UTF-8), to send when a test fails.
    Dim st As Object
    On Error Resume Next
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                                  ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText Text
    st.SaveToFile CurrentProject.Path & "\RunAllTests.log", 2      ' adSaveCreateOverWrite
    st.Close
End Sub

Private Function SkipReason(ByVal TestName As String) As String
    Dim loaded As Variant
    Select Case TestName
        Case "TestQueries"
            ' its expected numbers are computed for an empty database
            If Nz(DbValue("SELECT COUNT(*) FROM Products"), 0) > 0 Or _
               Nz(DbValue("SELECT COUNT(*) FROM Suppliers"), 0) > 0 Or _
               Nz(DbValue("SELECT COUNT(*) FROM Customers"), 0) > 1 Then
                SkipReason = "Ì⁄„· ⁄·Ï ﬁ«⁄œ… ›«—€… ›ﬁÿ (‘€¯·Â ﬁ»· ≈œŒ«· √Ì »Ì«‰« )"
            End If
        Case "VerifyDemoData"
            loaded = DMax("LogDate", "AuditLog", "ActionType = 'DEMO_LOADED'")
            If IsNull(loaded) Then
                SkipReason = "·«  ÊÃœ »Ì«‰«   Ã—Ì»Ì…"
            ElseIf Not IsNull(DMax("LogDate", "AuditLog", "ActionType = 'DEMO_REMOVED' AND LogDate > " & SqlDate(loaded))) Then
                SkipReason = "Õı–›  «·»Ì«‰«  «· Ã—Ì»Ì…"
            ElseIf Nz(DbValue("SELECT COUNT(*) FROM SalesInvoices WHERE CreatedAt > " & SqlDate(loaded)), 0) + _
                   Nz(DbValue("SELECT COUNT(*) FROM PurchaseInvoices WHERE CreatedAt > " & SqlDate(loaded)), 0) > 0 Then
                SkipReason = "√ıœŒ·  „” ‰œ«  »⁄œ «·»Ì«‰«  «· Ã—Ì»Ì… › €Ì—  «·√—ﬁ«„ «·„ Êﬁ⁄…"
            End If
    End Select
End Function
