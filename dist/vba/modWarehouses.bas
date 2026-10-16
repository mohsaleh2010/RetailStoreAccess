Attribute VB_Name = "modWarehouses"
'==============================================================================
' modWarehouses  -  Retail Store Management System (warehouses, docs/57)
'
' Every stock move carries its warehouse (InventoryTransactions.WarehouseID), and so do the
' documents that move stock (sales, returns, purchases, stock counts, opening stock). The balance
' of a product in a warehouse is the sum of its moves there; Products.CurrentQuantity stays the
' total of every warehouse and the average cost stays one per product.
'
'   CurrentWarehouseID   warehouse of the logged-in user (Employees.WarehouseID), else the main one
'   WarehouseStock       quantity of a product in a warehouse (the sum of its moves there)
'   AvailableQty         what can leave a warehouse; while every move is in the main warehouse
'                        it is Products.CurrentQuantity, exactly as before warehouses
'   WarehouseProblem     "" when the warehouse can be used by a document
' Screen: frmWarehouses (built by modBuildForms).
'==============================================================================
Option Compare Database
Option Explicit

Public Const MAIN_WAREHOUSE_ID As Long = 1

Public Function CurrentWarehouseID() As Long
    Dim v As Variant
    v = DbValue("SELECT e.WarehouseID FROM Employees AS e INNER JOIN Warehouses AS w ON e.WarehouseID = w.WarehouseID " & _
                "WHERE e.EmployeeID = " & CurrentUserID() & " AND w.IsActive = True")
    If IsNull(v) Then v = DbValue("SELECT Min(WarehouseID) FROM Warehouses WHERE IsActive = True AND WarehouseType = 'MAIN'")
    If IsNull(v) Then v = DbValue("SELECT Min(WarehouseID) FROM Warehouses WHERE IsActive = True")
    CurrentWarehouseID = Nz(v, MAIN_WAREHOUSE_ID)
End Function

Public Function WarehouseIsActive(ByVal WarehouseID As Variant) As Boolean
    If IsNull(WarehouseID) Then Exit Function
    WarehouseIsActive = Nz(DbValue("SELECT COUNT(*) FROM Warehouses WHERE IsActive = True AND WarehouseID = " & _
                                   CLng(WarehouseID)), 0) > 0
End Function

Public Function WarehouseProblem(ByVal WarehouseID As Variant) As String
    If Not WarehouseIsActive(WarehouseID) Then WarehouseProblem = "«Œ — «·„Œ“‰ („Œ“‰ ‰‘ÿ)."
End Function

Public Function WarehouseName(ByVal WarehouseID As Variant) As String
    If IsNull(WarehouseID) Then Exit Function
    WarehouseName = Nz(DbValue(Tr("SELECT w.WarehouseName FROM [@Warehouses] AS w WHERE w.WarehouseID = " & _
                                  CLng(WarehouseID))), "")
End Function

Public Function OnlyMainWarehouse() As Boolean
    ' True while no stock move was made outside the main warehouse.
    OnlyMainWarehouse = Nz(DbValue("SELECT COUNT(*) FROM InventoryTransactions WHERE WarehouseID <> " & _
                                   MAIN_WAREHOUSE_ID), 0) = 0
End Function

Public Function WarehouseStock(ByVal ProductID As Long, ByVal WarehouseID As Long) As Currency
    WarehouseStock = Nz(DbValue("SELECT Sum(Quantity) FROM InventoryTransactions WHERE ProductID = " & ProductID & _
                                " AND WarehouseID = " & WarehouseID), 0)
End Function

Public Function AvailableQty(ByVal ProductID As Long, Optional ByVal WarehouseID As Long = 0) As Currency
    If WarehouseID = 0 Then WarehouseID = CurrentWarehouseID()
    If WarehouseID = MAIN_WAREHOUSE_ID And OnlyMainWarehouse() Then
        AvailableQty = Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & ProductID), 0)
    Else
        AvailableQty = WarehouseStock(ProductID, WarehouseID)
    End If
End Function

Public Function WarehouseRows() As String
    ' The active warehouses for a combo box (through Tr: an English front-end shows the English names).
    WarehouseRows = Tr("SELECT w.WarehouseID, w.WarehouseName FROM [@Warehouses] AS w WHERE w.IsActive = True " & _
                       "ORDER BY w.WarehouseID")
End Function

'==============================================================================
' Test (inside a transaction that is rolled back)
'==============================================================================
Public Function TestWarehouses() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, wh2 As Long, pid As Long, refNo As String, msg As String
    Dim passed As Long, failed As Long, report As String, countID As Long, lines As Long, netValue As Currency
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestWarehouses  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    CheckWh WarehouseIsActive(MAIN_WAREHOUSE_ID), "«·„Œ“‰ «·—∆Ì”Ì „ÊÃÊœ Ê‰‘ÿ", passed, failed, report
    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "INSERT INTO Warehouses (WarehouseCode, WarehouseName, WarehouseType) VALUES " & _
                      "('TEST-WH', 'TEST „Œ“‰ ›—⁄', 'BRANCH')", dbFailOnError
    wh2 = DbValue("SELECT Max(WarehouseID) FROM Warehouses")
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, SellingPrice, " & _
                      "CurrentQuantity, AverageCost) VALUES ('TEST-WH-P', 'TEST ’‰› «·„Œ«“‰', " & _
                      Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1) & ", " & _
                      Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1) & ", 20, 0, 0)", dbFailOnError
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-WH-P'")

    msg = PostManualStock(pid, TT_STOCK_IN, 10, 5, "TEST", refNo, MAIN_WAREHOUSE_ID)
    CheckWh Len(msg) = 0, "≈÷«›… 10 ··„Œ“‰ «·—∆Ì”Ì " & msg, passed, failed, report
    msg = PostManualStock(pid, TT_STOCK_IN, 4, 5, "TEST", refNo, wh2)
    CheckWh Len(msg) = 0, "≈÷«›… 4 ·„Œ“‰ «·›—⁄ " & msg, passed, failed, report
    CheckWh WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 10 And WarehouseStock(pid, wh2) = 4, _
            "—’Ìœ ﬂ· „Œ“‰ = Õ—ﬂ« Â ›ÌÂ", passed, failed, report
    CheckWh Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & pid), 0) = 14, _
            "—’Ìœ «·’‰› = „Ã„Ê⁄ «·„Œ«“‰", passed, failed, report
    CheckWh AvailableQty(pid, wh2) = 4 And AvailableQty(pid, MAIN_WAREHOUSE_ID) = 10, _
            "«·„ «Õ ··’—› „‰ ﬂ· „Œ“‰", passed, failed, report
    CheckWh Nz(DbValue("SELECT Sum(StockQty) FROM StockByWarehouseQuery WHERE ProductID = " & pid & _
                       " AND WarehouseID = " & wh2), 0) = 4, " ﬁ—Ì— √—’œ… «·„Œ«“‰", passed, failed, report

    ' the user's warehouse
    CurrentDb.Execute "UPDATE Employees SET WarehouseID = " & wh2 & " WHERE EmployeeID = " & CurrentUserID(), dbFailOnError
    CheckWh CurrentWarehouseID() = wh2, "„Œ“‰ «·„” Œœ„", passed, failed, report
    msg = PostManualStock(pid, TT_STOCK_OUT, 1, Null, "TEST  «·›", refNo)
    CheckWh Len(msg) = 0 And WarehouseStock(pid, wh2) = 3 And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 10, _
            "«·Œ’„ „‰ „Œ“‰ «·„” Œœ„ " & msg, passed, failed, report
    CurrentDb.Execute "UPDATE Employees SET WarehouseID = Null WHERE EmployeeID = " & CurrentUserID(), dbFailOnError

    ' a stock count of one warehouse compares with that warehouse only
    CurrentDb.Execute "UPDATE StockCounts SET Status = 'CANCELLED' WHERE Status = 'OPEN'", dbFailOnError
    msg = CreateStockCount(Null, "TEST", countID, wh2)
    CheckWh Len(msg) = 0 And Nz(DbValue("SELECT SystemQuantity FROM StockCountDetails WHERE StockCountID = " & _
            countID & " AND ProductID = " & pid), -1) = 3, "«·Ã—œ Ìﬁ—√ —’Ìœ „Œ“‰Â " & msg, passed, failed, report
    CurrentDb.Execute "UPDATE StockCountDetails SET ActualQuantity = 2 WHERE StockCountID = " & countID & _
                      " AND ProductID = " & pid, dbFailOnError
    CurrentDb.Execute "UPDATE StockCountDetails SET ActualQuantity = SystemQuantity WHERE StockCountID = " & countID & _
                      " AND ProductID <> " & pid, dbFailOnError
    msg = PostStockCount(countID, False, lines, netValue)
    CheckWh Len(msg) = 0 And WarehouseStock(pid, wh2) = 2 And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 10, _
            " ”ÊÌ… «·Ã—œ ›Ì „Œ“‰Â ›ﬁÿ " & msg, passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckWh False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·„Œ«“‰ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestWarehouses"
        TestWarehouses = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
                "TestWarehouses"
    End If
End Function

Private Sub CheckWh(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
                    ByRef report As String)
    If ok Then
        passed = passed + 1
        Debug.Print "  [OK] " & Title
    Else
        failed = failed + 1
        report = report & "- " & Title & vbCrLf
        Debug.Print "  [X]  " & Title
    End If
End Sub
