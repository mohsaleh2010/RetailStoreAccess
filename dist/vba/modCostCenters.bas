Attribute VB_Name = "modCostCenters"
'==============================================================================
' modCostCenters  -  Retail Store Management System
'
' Cost centres / branches (CostCenters, screen frmCostCenters): each journal line
' carries the centre of its document (JournalLines.CostCenterID):
'   sales               the centre of the cashier, else the default centre
'   sales returns       the centre of the original invoice
'   expenses, cash vouchers   the centre chosen (default: as sales)
'   fixed assets        the centre of the asset (depreciation, gain / loss)
'   payroll             the centre of each employee
'   manual entries      a centre per line
' Documents saved before the centres had none ("not allocated").
' Reports: CostCenterProfitQuery (an income statement per centre) and
' CostCenterAccountsQuery (the accounts of each centre).
'==============================================================================
Option Compare Database
Option Explicit

Public Function CostCenterFor(Optional ByVal EmployeeID As Long = 0) As Variant
    ' The centre of the employee (the current user by default) if active, else the default centre, else Null.
    Dim v As Variant
    If EmployeeID = 0 Then EmployeeID = CurrentUserID()
    v = DbValue("SELECT e.CostCenterID FROM Employees AS e INNER JOIN CostCenters AS c ON e.CostCenterID = " & _
                "c.CostCenterID WHERE c.IsActive = True AND e.EmployeeID = " & EmployeeID)
    If IsNull(v) Then v = DbValue("SELECT Min(CostCenterID) FROM CostCenters WHERE IsDefault = True AND IsActive = True")
    CostCenterFor = v
End Function

Public Function CostCenterRows() As String
    CostCenterRows = "SELECT CostCenterID, CenterName FROM CostCenters WHERE IsActive = True ORDER BY CenterCode"
End Function

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Public Function TestCostCenters() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, passed As Long, failed As Long, report As String, msg As String
    Dim center As Long, other As Long, voucher As Long, box As Long, entry As Variant
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestCostCenters  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    If ClosedThroughDate() >= Date Then
        Debug.Print "[--] «·› —… „ﬁ›·… Õ Ï «·ÌÊ„"
        GoTo Undo
    End If
    CurrentDb.Execute "UPDATE CostCenters SET IsDefault = False", dbFailOnError
    CurrentDb.Execute "INSERT INTO CostCenters (CenterCode, CenterName, IsDefault, IsActive) VALUES ('TST1', 'TEST „—ﬂ“ 1', " & _
                      "True, True)", dbFailOnError
    center = DbValue("SELECT CostCenterID FROM CostCenters WHERE CenterCode = 'TST1'")
    CurrentDb.Execute "INSERT INTO CostCenters (CenterCode, CenterName, IsDefault, IsActive) VALUES ('TST2', 'TEST „—ﬂ“ 2', " & _
                      "False, True)", dbFailOnError
    other = DbValue("SELECT CostCenterID FROM CostCenters WHERE CenterCode = 'TST2'")
    CurrentDb.Execute "UPDATE Employees SET CostCenterID = Null WHERE EmployeeID = " & CurrentUserID(), dbFailOnError
    Record CostCenterFor() = center, "»·« „—ﬂ“ ··„ÊŸ›: «·„—ﬂ“ «·«› —«÷Ì", passed, failed, report
    CurrentDb.Execute "UPDATE Employees SET CostCenterID = " & other & " WHERE EmployeeID = " & CurrentUserID(), dbFailOnError
    Record CostCenterFor() = other, "„—ﬂ“ «·„ÊŸ› ﬁ»· «·«› —«÷Ì", passed, failed, report
    box = Nz(DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE IsActive = True"), 0)
    msg = PostCashVoucher("IN", box, Null, "OTHER", 50, "TEST", "TEST-CC", Null, voucher)
    msg = msg & SyncJournal()
    entry = DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType = 'CASH_VOUCHER' AND SourceID = " & voucher)
    Record Len(msg) = 0 And Nz(DbValue("SELECT COUNT(*) FROM JournalLines WHERE EntryID = " & Nz(entry, 0) & _
                                       " AND CostCenterID = " & other), 0) = 2, "√”ÿ— ﬁÌœ «·”‰œ ⁄·Ï „—ﬂ“ «·„ÊŸ› " & msg, _
           passed, failed, report
    SetPeriod Date, Date
    Record Nz(DbValue("SELECT Revenue FROM CostCenterProfitQuery WHERE CenterKey = " & other), 0) >= 50, _
           "ﬁ«∆„… «·œŒ· ··„—ﬂ“  ‘„· «·≈Ì—«œ", passed, failed, report
Undo:
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    Record False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  „—«ﬂ“ «· ﬂ·›… ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestCostCenters"
        TestCostCenters = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestCostCenters"
    End If
End Function

Private Sub Record(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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
