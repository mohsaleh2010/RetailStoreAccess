Attribute VB_Name = "modDemoData"
'==============================================================================
' modDemoData  -  Retail Store Management System (Phase 11: demo data)
'
' GENERATED FILE - do not edit by hand.  Source: tools/demo_data.py, tools/gen_demo.py
'
'   LoadDemoData     enters the demo store: 20 „‰ Ã«° 5  ’‰Ì›« ° 10 ⁄„·«¡° 5 „Ê—œÌ‰° 3 „ÊŸ›Ì‰° 10 ›Ê« Ì— »Ì⁄° 5 ›Ê« Ì— ‘—«¡° 6 „’—Ê›« .
'                    Every document goes through the real posting functions, in ONE
'                    transaction (all or nothing), then gets its planned date.
'                    Only on a database without documents.
'   VerifyDemoData   compares the result with the Python replay of the same plan.
'   RemoveDemoData   removes the demo before real use (backup first). Refused
'                    once real documents were entered after the demo.
' Demo users (manager, cashier1, cashier2) have the temporary password
' "Demo@2026" and must change it at their first login.
'==============================================================================
Option Compare Database
Option Explicit

Private Const DEMO_PASSWORD As String = "Demo@2026"
Private Const DEMO_MARK As String = "DEMO"

Private m_sales(1 To 10) As Long
Private m_purchases(1 To 5) As Long
Private m_adminID As Long
Private m_passed As Long
Private m_failed As Long
Private m_report As String

'==============================================================================
' Load
'==============================================================================
Public Function LoadDemoData() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, problem As String, errText As String
    On Error GoTo EH
    Calendar = vbCalGreg
    EnsureTestUser
    m_adminID = CurrentUserID()
    problem = DemoBlocker()
    If Len(problem) > 0 Then
        MsgBox problem, vbExclamation + MSG_RTL, "LoadDemoData"
        Exit Function
    End If
    If MsgBox("”Ì „ ≈œŒ«· „ Ã—  Ã—Ì»Ì ﬂ«„·:" & vbCrLf & "20 „‰ Ã«° 5  ’‰Ì›« ° 10 ⁄„·«¡° 5 „Ê—œÌ‰° 3 „ÊŸ›Ì‰° 10 ›Ê« Ì— »Ì⁄° 5 ›Ê« Ì— ‘—«¡° 6 „’—Ê›« ." & vbCrLf & vbCrLf & _
              "·· œ—Ì» Ê«· Ã—»… ›ﬁÿ. ﬁ»· «·«” Œœ«„ «·›⁄·Ì «Õ–›Â« »«·√„— RemoveDemoData." & vbCrLf & _
              "Â·  —Ìœ «·„ «»⁄…ø", vbQuestion + vbYesNo + vbDefaultButton2 + MSG_RTL, "LoadDemoData") <> vbYes Then
        Exit Function
    End If
    DoCmd.Hourglass True
    EnsureLocalTables
    ClearWorkTables
    g_SilentMode = True
    g_AutoAnswer = True
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    DemoSettings
    DemoMasters
    DemoSteps1
    DemoSteps2
    DemoSteps3
    DemoSteps4
    DemoSteps5
    DemoSteps6
    TempVars.Add "UserID", m_adminID
    Check SyncJournal(), "ﬁÌÊœ «·ÌÊ„Ì…"
    ws.CommitTrans
    inTrans = False
    TempVars.Add "UserID", m_adminID
    LogAction "DEMO_LOADED", "", "", "20 „‰ Ã«° 5  ’‰Ì›« ° 10 ⁄„·«¡° 5 „Ê—œÌ‰° 3 „ÊŸ›Ì‰° 10 ›Ê« Ì— »Ì⁄° 5 ›Ê« Ì— ‘—«¡° 6 „’—Ê›« "
    g_SilentMode = False
    DoCmd.Hourglass False
    LoadDemoData = VerifyDemoData()
    Exit Function
EH:
    errText = Err.Description
    On Error Resume Next
    If inTrans Then ws.Rollback
    TempVars.Add "UserID", m_adminID
    ClearWorkTables
    g_SilentMode = False
    DoCmd.Hourglass False
    MsgBox " ⁄–— ≈œŒ«· «·»Ì«‰«  «· Ã—Ì»Ì…° Ê·„ ÌıÕ›Ÿ „‰Â« √Ì ‘Ì¡:" & vbCrLf & vbCrLf & errText, _
           vbCritical + MSG_RTL, "LoadDemoData"
End Function

Private Function DemoBlocker() As String
    ' Reasons not to load: documents already exist, demo already there, settings differ from the plan.
    Dim t As Variant
    If Not IsAdministrator() Then
        DemoBlocker = "≈œŒ«· «·»Ì«‰«  «· Ã—Ì»Ì… ·„œÌ— «·‰Ÿ«„ ›ﬁÿ."
        Exit Function
    End If
    For Each t In Array("SalesInvoices", "PurchaseInvoices", "InventoryTransactions", "Expenses", _
                        "CustomerPayments", "SupplierPayments", "StockCounts", "CashVouchers", "CashClosings")
        If Nz(DbValue("SELECT COUNT(*) FROM [" & t & "]"), 0) > 0 Then
            DemoBlocker = " ÊÃœ »Ì«‰«  ›⁄·Ì… (" & t & ")." & vbCrLf & _
                          "«·»Ì«‰«  «· Ã—Ì»Ì…  ıœŒ· ›Ì ﬁ«⁄œ… »œÊ‰ „” ‰œ«  ›ﬁÿ (‰”Œ… ·· œ—Ì»)."
            Exit Function
        End If
    Next
    If Nz(DbValue("SELECT COUNT(*) FROM Products WHERE Barcode = " & SqlText("6281000000014")), 0) > 0 Or _
       Nz(DbValue("SELECT COUNT(*) FROM Employees WHERE Username = 'manager' OR Username = 'cashier1' " & _
                  "OR Username = 'cashier2'"), 0) > 0 Then
        DemoBlocker = "ÌÊÃœ Ã“¡ „‰ «·»Ì«‰«  «· Ã—Ì»Ì… „”»ﬁ« („‰ Ã √Ê „” Œœ„ »‰›” «·«”„)."
        Exit Function
    End If
    If Nz(SettingValue("VATRate"), 0) <> 0.15 Or Not Nz(SettingValue("PricesIncludeVAT"), False) Then
        DemoBlocker = "«·»Ì«‰«  «· Ã—Ì»Ì… „»‰Ì… ⁄·Ï ‰”»… ÷—Ì»… 15% Ê√”⁄«— ‘«„·… «·÷—Ì»…." & vbCrLf & _
                      "√⁄œ Â–Ì‰ «·≈⁄œ«œÌ‰ À„ Õ«Ê· „—… √Œ—Ï."
    End If
End Function

Private Sub DemoSettings()
    ' The demo store details are used only when the store is not set up yet.
    If Not IsNull(SettingValue("VATNumber")) Then Exit Sub
    CurrentDb.Execute "UPDATE Settings SET StoreName = '„ Ã— «·‰Œ»… ·· Ã“∆…', StoreNameEn = 'Al Nukhba Retail Store', VATNumber = '310123456700003', CRNumber = '1010654321', BuildingNo = '2345', StreetName = 'ÿ—Ìﬁ «·„·ﬂ ›Âœ', District = '«·⁄·Ì«', City = '«·—Ì«÷', PostalCode = '12211', Phone = '0112345678' WHERE SettingID = 1", dbFailOnError
End Sub

Private Sub DemoMasters()
    CurrentDb.Execute "INSERT INTO Categories (CategoryName, Description) VALUES (" & SqlText("„Ê«œ €–«∆Ì…") & ", 'DEMO')", dbFailOnError
    CurrentDb.Execute "INSERT INTO Categories (CategoryName, Description) VALUES (" & SqlText("„‘—Ê»« ") & ", 'DEMO')", dbFailOnError
    CurrentDb.Execute "INSERT INTO Categories (CategoryName, Description) VALUES (" & SqlText("„‰Ÿ›« ") & ", 'DEMO')", dbFailOnError
    CurrentDb.Execute "INSERT INTO Categories (CategoryName, Description) VALUES (" & SqlText("⁄‰«Ì… ‘Œ’Ì…") & ", 'DEMO')", dbFailOnError
    CurrentDb.Execute "INSERT INTO Categories (CategoryName, Description) VALUES (" & SqlText("√œÊ«  „‰“·Ì…") & ", 'DEMO')", dbFailOnError
    CurrentDb.Execute "INSERT INTO Suppliers (SupplierName, ContactPerson, Mobile, VATNumber, City, Notes, CreatedAt) VALUES (" & SqlText("„ƒ””… «·Ê›—… ··„Ê«œ «·€–«∆Ì…") & ", " & SqlText("”«·„ «·⁄„—Ì") & ", '0551110001', '300112233400003', " & SqlText("«·—Ì«÷") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Suppliers (SupplierName, ContactPerson, Mobile, VATNumber, City, Notes, CreatedAt) VALUES (" & SqlText("‘—ﬂ… «·Ì‰«»Ì⁄ ··„‘—Ê»« ") & ", " & SqlText("„«Ãœ «·Õ—»Ì") & ", '0551110002', '300223344500003', " & SqlText("Ãœ…") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Suppliers (SupplierName, ContactPerson, Mobile, VATNumber, City, Notes, CreatedAt) VALUES (" & SqlText("„ƒ””… «·‰ﬁ«¡ ··„‰Ÿ›« ") & ", " & SqlText(" —ﬂÌ «·€«„œÌ") & ", '0551110003', '300334455600003', " & SqlText("«·œ„«„") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Suppliers (SupplierName, ContactPerson, Mobile, VATNumber, City, Notes, CreatedAt) VALUES (" & SqlText("‘—ﬂ… «·⁄‰«Ì… «·–Â»Ì…") & ", " & SqlText("Â‰œ «·”»Ì⁄Ì") & ", '0551110004', '300445566700003', " & SqlText("«·—Ì«÷") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Suppliers (SupplierName, ContactPerson, Mobile, VATNumber, City, Notes, CreatedAt) VALUES (" & SqlText("„ƒ””… «·»Ì  «·⁄’—Ì ··√œÊ« ") & ", " & SqlText("Ê·Ìœ «·‘„—Ì") & ", '0551110005', '300556677800003', " & SqlText("«·—Ì«÷") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000014', " & SqlText("√—“ »”„ Ì 5 ﬂÃ„") & ", " & CategoryIDOf("„Ê«œ €–«∆Ì…") & ", 4, 38.00, 52.00, 10, " & SupplierIDOf("„ƒ””… «·Ê›—… ··„Ê«œ «·€–«∆Ì…") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000021', " & SqlText("”ﬂ— √»Ì÷ 2 ﬂÃ„") & ", " & CategoryIDOf("„Ê«œ €–«∆Ì…") & ", 4, 9.50, 13.50, 15, " & SupplierIDOf("„ƒ””… «·Ê›—… ··„Ê«œ «·€–«∆Ì…") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000038', " & SqlText("“Ì  œÊ«— «·‘„” 1.5 · —") & ", " & CategoryIDOf("„Ê«œ €–«∆Ì…") & ", 1, 14.00, 19.95, 12, " & SupplierIDOf("„ƒ””… «·Ê›—… ··„Ê«œ «·€–«∆Ì…") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000045', " & SqlText("„⁄ﬂ—Ê‰… 450 Ã„") & ", " & CategoryIDOf("„Ê«œ €–«∆Ì…") & ", 1, 2.60, 3.95, 30, " & SupplierIDOf("„ƒ””… «·Ê›—… ··„Ê«œ «·€–«∆Ì…") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000052', " & SqlText(" „— ”ﬂ—Ì 1 ﬂÃ„") & ", " & CategoryIDOf("„Ê«œ €–«∆Ì…") & ", 2, 22.00, 34.50, 8, " & SupplierIDOf("„ƒ””… «·Ê›—… ··„Ê«œ «·€–«∆Ì…") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000069', " & SqlText("„Ì«Â „⁄œ‰Ì… 330 „· ◊ 40") & ", " & CategoryIDOf("„‘—Ê»« ") & ", 3, 11.00, 16.50, 10, " & SupplierIDOf("‘—ﬂ… «·Ì‰«»Ì⁄ ··„‘—Ê»« ") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000076', " & SqlText("⁄’Ì— »— ﬁ«· 1 · —") & ", " & CategoryIDOf("„‘—Ê»« ") & ", 1, 4.20, 6.50, 20, " & SupplierIDOf("‘—ﬂ… «·Ì‰«»Ì⁄ ··„‘—Ê»« ") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000083', " & SqlText("Õ·Ì» ÿ«“Ã 2 · —") & ", " & CategoryIDOf("„‘—Ê»« ") & ", 1, 7.80, 11.00, 15, " & SupplierIDOf("‘—ﬂ… «·Ì‰«»Ì⁄ ··„‘—Ê»« ") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000090', " & SqlText("ﬁÂÊ… ⁄—»Ì… 500 Ã„") & ", " & CategoryIDOf("„‘—Ê»« ") & ", 4, 26.00, 39.00, 6, " & SupplierIDOf("‘—ﬂ… «·Ì‰«»Ì⁄ ··„‘—Ê»« ") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000106', " & SqlText("‘«Ì √ﬂÌ«” ◊ 100") & ", " & CategoryIDOf("„‘—Ê»« ") & ", 2, 9.00, 14.25, 10, " & SupplierIDOf("‘—ﬂ… «·Ì‰«»Ì⁄ ··„‘—Ê»« ") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000113', " & SqlText("„‰Ÿ› √—÷Ì«  3 · —") & ", " & CategoryIDOf("„‰Ÿ›« ") & ", 1, 11.50, 17.25, 8, " & SupplierIDOf("„ƒ””… «·‰ﬁ«¡ ··„‰Ÿ›« ") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000120', " & SqlText("”«∆· €”Ì· ’ÕÊ‰ 1 · —") & ", " & CategoryIDOf("„‰Ÿ›« ") & ", 1, 5.30, 8.50, 12, " & SupplierIDOf("„ƒ””… «·‰ﬁ«¡ ··„‰Ÿ›« ") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000137', " & SqlText("„”ÕÊﬁ €”Ì· 3 ﬂÃ„") & ", " & CategoryIDOf("„‰Ÿ›« ") & ", 4, 24.00, 36.00, 6, " & SupplierIDOf("„ƒ””… «·‰ﬁ«¡ ··„‰Ÿ›« ") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000144', " & SqlText("„‰«œÌ· Ê—ﬁÌ… ◊ 5") & ", " & CategoryIDOf("„‰Ÿ›« ") & ", 4, 8.40, 12.95, 15, " & SupplierIDOf("„ƒ””… «·‰ﬁ«¡ ··„‰Ÿ›« ") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000151', " & SqlText("‘«„»Ê 400 „·") & ", " & CategoryIDOf("⁄‰«Ì… ‘Œ’Ì…") & ", 1, 12.00, 18.95, 8, " & SupplierIDOf("‘—ﬂ… «·⁄‰«Ì… «·–Â»Ì…") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000168', " & SqlText("„⁄ÃÊ‰ √”‰«‰ 100 „·") & ", " & CategoryIDOf("⁄‰«Ì… ‘Œ’Ì…") & ", 1, 6.10, 9.50, 12, " & SupplierIDOf("‘—ﬂ… «·⁄‰«Ì… «·–Â»Ì…") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000175', " & SqlText("’«»Ê‰ ÌœÌ‰ 500 „·") & ", " & CategoryIDOf("⁄‰«Ì… ‘Œ’Ì…") & ", 1, 5.20, 8.25, 10, " & SupplierIDOf("‘—ﬂ… «·⁄‰«Ì… «·–Â»Ì…") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000182', " & SqlText("√ﬂÌ«” ‰›«Ì«  ﬂ»Ì—… ◊ 30") & ", " & CategoryIDOf("√œÊ«  „‰“·Ì…") & ", 4, 7.00, 10.95, 10, " & SupplierIDOf("„ƒ””… «·»Ì  «·⁄’—Ì ··√œÊ« ") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000199', " & SqlText("Ê—ﬁ √·„‰ÌÊ„ 30 „") & ", " & CategoryIDOf("√œÊ«  „‰“·Ì…") & ", 1, 9.20, 14.50, 6, " & SupplierIDOf("„ƒ””… «·»Ì  «·⁄’—Ì ··√œÊ« ") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000205', " & SqlText("»ÿ«—Ì«  AA ◊ 4") & ", " & CategoryIDOf("√œÊ«  „‰“·Ì…") & ", 4, 8.00, 13.00, 5, " & SupplierIDOf("„ƒ””… «·»Ì  «·⁄’—Ì ··√œÊ« ") & ", 'DEMO', " & SqlDate(Date - 150) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("„ÿ⁄„ «·œÌ—…") & ", '0501000001', '300667788900003', " & SqlText("«·—Ì«÷") & ", True, 5000, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("„ƒ””… «·÷Ì«›… ·· „ÊÌ‰") & ", '0501000002', '300778899000003', " & SqlText("«·—Ì«÷") & ", True, 10000, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("„ﬁÂÏ «·—ﬂ‰") & ", '0501000003', '300889900100003', " & SqlText("«·—Ì«÷") & ", True, 3000, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("√Õ„œ „Õ„œ «·⁄ Ì»Ì") & ", '0501000004', Null, " & SqlText("«·—Ì«÷") & ", True, 1000, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("”«—… ⁄»œ«··Â «·ﬁÕÿ«‰Ì") & ", '0501000005', Null, " & SqlText("«·—Ì«÷") & ", True, 500, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("Œ«·œ ≈»—«ÂÌ„ «·‘Â—Ì") & ", '0501000006', Null, " & SqlText("«·—Ì«÷") & ", False, 0, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("‰Ê—… ”⁄œ «·œÊ”—Ì") & ", '0501000007', Null, " & SqlText("«·—Ì«÷") & ", False, 0, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("›Âœ ⁄»œ«·—Õ„‰ «·Õ—»Ì") & ", '0501000008', Null, " & SqlText("«·—Ì«÷") & ", True, 1500, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("„‰Ï ⁄·Ì «·“Â—«‰Ì") & ", '0501000009', Null, " & SqlText("«·—Ì«÷") & ", False, 0, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("⁄»œ«··Â ‰«’— «·„ÿÌ—Ì") & ", '0501000010', Null, " & SqlText("«·—Ì«÷") & ", True, 2000, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Employees (EmployeeName, JobTitle, Username, RoleID, MaxDiscountPercent, Notes) VALUES (" & SqlText("„‘—› «·›—⁄ ( Ã—Ì»Ì)") & ", " & SqlText("„œÌ— ›—⁄") & ", 'manager', 2, 0.10, 'DEMO')", dbFailOnError
    Check SetUserPassword(UserIDOf("manager"), DEMO_PASSWORD, True), "ﬂ·„… „—Ê— manager"
    CurrentDb.Execute "INSERT INTO Employees (EmployeeName, JobTitle, Username, RoleID, MaxDiscountPercent, Notes) VALUES (" & SqlText("ﬂ«‘Ì— «·’»«Õ ( Ã—Ì»Ì)") & ", " & SqlText("ﬂ«‘Ì—") & ", 'cashier1', 3, 0.02, 'DEMO')", dbFailOnError
    Check SetUserPassword(UserIDOf("cashier1"), DEMO_PASSWORD, True), "ﬂ·„… „—Ê— cashier1"
    CurrentDb.Execute "INSERT INTO Employees (EmployeeName, JobTitle, Username, RoleID, MaxDiscountPercent, Notes) VALUES (" & SqlText("ﬂ«‘Ì— «·„”«¡ ( Ã—Ì»Ì)") & ", " & SqlText("ﬂ«‘Ì—") & ", 'cashier2', 3, 0.02, 'DEMO')", dbFailOnError
    Check SetUserPassword(UserIDOf("cashier2"), DEMO_PASSWORD, True), "ﬂ·„… „—Ê— cashier2"
End Sub

Private Sub DemoSteps1()
    Dim when As Date, id As Long, refNo As String, adjusted As Long, netValue As Currency, counted As Currency
    ' --- cash_in (day -31)
    when = DemoWhen(31, 8)
    SetDemoUser ""
    Check PostCashVoucher("IN", BoxOfType("MAIN"), Null, "OWNER", 30000.00, "«·„«·ﬂ", "—√” „«·  ‘€Ì· ( Ã—Ì»Ì)", Null, id), "”‰œ ‰ﬁœÌ…"
    MoveDoc "CASH_VOUCHER", id, when
    ' --- purchase (day -30)
    when = DemoWhen(30, 9)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    PurLine "6281000000014", 40, "38.00"
    PurLine "6281000000021", 40, "9.50"
    PurLine "6281000000038", 50, "14.00"
    PurLine "6281000000045", 120, "2.60"
    PurLine "6281000000052", 30, "22.00"
    Check PostPurchaseFromCart(SupplierIDOf("„ƒ””… «·Ê›—… ··„Ê«œ «·€–«∆Ì…"), "INV-7781", DateValue(when), "CREDIT", 1, True, 0.00, 1000, "", id), "›« Ê—… ‘—«¡ 1"
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    MoveDoc "PURCHASE", id, when
    m_purchases(1) = id
    ' --- purchase (day -29)
    when = DemoWhen(29, 10)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    PurLine "6281000000069", 40, "11.00"
    PurLine "6281000000076", 80, "4.20"
    PurLine "6281000000083", 25, "7.80"
    PurLine "6281000000090", 25, "26.00"
    PurLine "6281000000106", 40, "9.00"
    Check PostPurchaseFromCart(SupplierIDOf("‘—ﬂ… «·Ì‰«»Ì⁄ ··„‘—Ê»« "), "SP-20451", DateValue(when), "CASH", 1, True, 0.00, Null, "", id), "›« Ê—… ‘—«¡ 2"
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    MoveDoc "PURCHASE", id, when
    m_purchases(2) = id
    ' --- expense (day -28)
    when = DemoWhen(28, 9)
    SetDemoUser ""
    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, TotalAmount, PaymentMethodID, Description, EmployeeID, CreatedAt, CashBoxID) VALUES (" & SqlText(NextNumber("EXPENSE")) & ", " & SqlDate(DateValue(when)) & ", 1, 4500.00, 0.00, 4500.00, 1, " & SqlText("≈ÌÃ«— «·„Õ· ··‘Â—") & ", " & CurrentUserID() & ", " & SqlDate(when) & ", " & CurrentCashBoxID() & ")", dbFailOnError
    ' --- purchase (day -27)
    when = DemoWhen(27, 11)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    PurLine "6281000000113", 30, "11.50"
    PurLine "6281000000120", 40, "5.30"
    PurLine "6281000000137", 8, "24.00"
    PurLine "6281000000144", 60, "8.40"
    Check PostPurchaseFromCart(SupplierIDOf("„ƒ””… «·‰ﬁ«¡ ··„‰Ÿ›« "), "N-3390", DateValue(when), "CREDIT", 1, True, 0.00, 0, "", id), "›« Ê—… ‘—«¡ 3"
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    MoveDoc "PURCHASE", id, when
    m_purchases(3) = id
    ' --- purchase (day -25)
    when = DemoWhen(25, 12)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    PurLine "6281000000151", 30, "12.00"
    PurLine "6281000000168", 15, "6.10"
    PurLine "6281000000175", 40, "5.20"
    Check PostPurchaseFromCart(SupplierIDOf("‘—ﬂ… «·⁄‰«Ì… «·–Â»Ì…"), "GC-118", DateValue(when), "CREDIT", 1, True, 0.00, 500, "", id), "›« Ê—… ‘—«¡ 4"
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    MoveDoc "PURCHASE", id, when
    m_purchases(4) = id
End Sub

Private Sub DemoSteps2()
    Dim when As Date, id As Long, refNo As String, adjusted As Long, netValue As Currency, counted As Currency
    ' --- purchase (day -22)
    when = DemoWhen(22, 10)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    PurLine "6281000000182", 40, "7.00"
    PurLine "6281000000199", 25, "9.20"
    PurLine "6281000000205", 20, "8.00"
    Check PostPurchaseFromCart(SupplierIDOf("„ƒ””… «·»Ì  «·⁄’—Ì ··√œÊ« "), "BA-5521", DateValue(when), "CASH", 1, True, 20.00, Null, "", id), "›« Ê—… ‘—«¡ 5"
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    MoveDoc "PURCHASE", id, when
    m_purchases(5) = id
    ' --- sale (day -20)
    when = DemoWhen(20, 10)
    SetDemoUser "cashier1"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000045", 10
    CartLine "6281000000021", 2
    CartLine "6281000000038", 1
    Check PostSaleFromCart(1, "CASH", 1, 0.00, Null, "", id), "›« Ê—… »Ì⁄ 1"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(1) = id
    ' --- expense (day -20)
    when = DemoWhen(20, 12)
    SetDemoUser ""
    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, TotalAmount, PaymentMethodID, Description, EmployeeID, CreatedAt, CashBoxID) VALUES (" & SqlText(NextNumber("EXPENSE")) & ", " & SqlDate(DateValue(when)) & ", 2, 380.00, 57.00, 437.00, 1, " & SqlText("›« Ê—… «·ﬂÂ—»«¡") & ", " & CurrentUserID() & ", " & SqlDate(when) & ", " & CurrentCashBoxID() & ")", dbFailOnError
    ' --- expense (day -19)
    when = DemoWhen(19, 12)
    SetDemoUser ""
    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, TotalAmount, PaymentMethodID, Description, EmployeeID, CreatedAt, CashBoxID) VALUES (" & SqlText(NextNumber("EXPENSE")) & ", " & SqlDate(DateValue(when)) & ", 3, 120.00, 18.00, 138.00, 1, " & SqlText("›« Ê—… «·„Ì«Â") & ", " & CurrentUserID() & ", " & SqlDate(when) & ", " & CurrentCashBoxID() & ")", dbFailOnError
    ' --- sale (day -18)
    when = DemoWhen(18, 11)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000014", 10
    CartLine "6281000000038", 6
    CartLine "6281000000069", 5
    Check PostSaleFromCart(CustomerIDOf("„ÿ⁄„ «·œÌ—…"), "CREDIT", 1, 0.00, 0, "", id), "›« Ê—… »Ì⁄ 2"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(2) = id
    ' --- expense (day -16)
    when = DemoWhen(16, 13)
    SetDemoUser ""
    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, TotalAmount, PaymentMethodID, Description, EmployeeID, CreatedAt, CashBoxID) VALUES (" & SqlText(NextNumber("EXPENSE")) & ", " & SqlDate(DateValue(when)) & ", 4, 299.00, 44.85, 343.85, 1, " & SqlText("«·≈‰ —‰  Ê«·Â« ›") & ", " & CurrentUserID() & ", " & SqlDate(when) & ", " & CurrentCashBoxID() & ")", dbFailOnError
End Sub

Private Sub DemoSteps3()
    Dim when As Date, id As Long, refNo As String, adjusted As Long, netValue As Currency, counted As Currency
    ' --- sale (day -15)
    when = DemoWhen(15, 17)
    SetDemoUser "cashier2"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000076", 6
    CartLine "6281000000168", 2
    CartLine "6281000000151", 1
    Check PostSaleFromCart(1, "CASH", 1, 0.00, Null, "", id), "›« Ê—… »Ì⁄ 3"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(3) = id
    ' --- supplier_payment (day -15)
    when = DemoWhen(15, 18)
    SetDemoUser ""
    Check PostSupplierPayment(SupplierIDOf("„ƒ””… «·Ê›—… ··„Ê«œ «·€–«∆Ì…"), 1500.00, 1, "œ›⁄… „‰ «·Õ”«»", id), "”‰œ ’—›"
    MoveDoc "SUPPLIER_PAYMENT", id, when
    ' --- closing (day -15)
    when = DemoWhen(15, 23)
    SetDemoUser ""
    counted = CashBoxBalance(BoxOfType("CASHIER")) - 5.00
    Check PostCashClosing(BoxOfType("CASHIER"), counted, "MAIN", BoxOfType("MAIN"), IIf(counted > 100.00, counted - 100.00, 0), " ’›Ì…  Ã—Ì»Ì…", id), " ’›Ì… «·ﬂ«‘Ì—"
    MoveDoc "CASH_CLOSING", id, when
    ' --- purchase_return (day -14)
    when = DemoWhen(14, 10)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpPurchaseReturnLines (PurchaseDetailID, ProductID, ReturnQty) SELECT PurchaseDetailID, ProductID, 3 FROM PurchaseInvoiceDetails WHERE PurchaseInvoiceID = " & m_purchases(3) & " AND LineNumber = 1", dbFailOnError
    Check PostPurchaseReturn(m_purchases(3), "⁄Ì» „’‰⁄Ì ( Ã—Ì»Ì)", "CREDIT", Null, id), "„— Ã⁄ „‘ —Ì« "
    CurrentDb.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError
    MoveDoc "PURCHASE_RETURN", id, when
    ' --- sale (day -12)
    when = DemoWhen(12, 13)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000052", 2
    CartLine "6281000000090", 1
    Check PostSaleFromCart(CustomerIDOf("√Õ„œ „Õ„œ «·⁄ Ì»Ì"), "CREDIT", 1, 0.00, 50, "", id), "›« Ê—… »Ì⁄ 4"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(4) = id
    ' --- expense (day -11)
    when = DemoWhen(11, 9)
    SetDemoUser ""
    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, TotalAmount, PaymentMethodID, Description, EmployeeID, CreatedAt, CashBoxID) VALUES (" & SqlText(NextNumber("EXPENSE")) & ", " & SqlDate(DateValue(when)) & ", 5, 150.00, 0.00, 150.00, 1, " & SqlText("‰ﬁ· »÷«⁄…") & ", " & CurrentUserID() & ", " & SqlDate(when) & ", " & CurrentCashBoxID() & ")", dbFailOnError
End Sub

Private Sub DemoSteps4()
    Dim when As Date, id As Long, refNo As String, adjusted As Long, netValue As Currency, counted As Currency
    ' --- sale (day -10)
    when = DemoWhen(10, 19)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000137", 2
    CartLine "6281000000113", 1
    CartLine "6281000000144", 3
    Check PostSaleFromCart(1, "CASH", 1, 5.00, Null, "", id), "›« Ê—… »Ì⁄ 5"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(5) = id
    ' --- supplier_payment (day -9)
    when = DemoWhen(9, 10)
    SetDemoUser ""
    Check PostSupplierPayment(SupplierIDOf("„ƒ””… «·‰ﬁ«¡ ··„‰Ÿ›« "), 800.00, 1, "œ›⁄… „‰ «·Õ”«»", id), "”‰œ ’—›"
    MoveDoc "SUPPLIER_PAYMENT", id, when
    ' --- sale (day -8)
    when = DemoWhen(8, 12)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000014", 15
    CartLine "6281000000021", 25
    CartLine "6281000000045", 60
    CartLine "6281000000069", 20
    Check PostSaleFromCart(CustomerIDOf("„ƒ””… «·÷Ì«›… ·· „ÊÌ‰"), "CREDIT", 1, 0.00, 1000, "", id), "›« Ê—… »Ì⁄ 6"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(6) = id
    ' --- stock_out (day -7)
    when = DemoWhen(7, 9)
    SetDemoUser ""
    Check PostManualStock(ProductIDOf("6281000000120"), TT_STOCK_OUT, 1, Null, " «·› - ⁄»Ê… „ﬂ”Ê—…", refNo), "Œ’„ „Œ“Ê‰"
    MoveManual refNo, when
    ' --- cash_out (day -7)
    when = DemoWhen(7, 20)
    SetDemoUser ""
    Check PostCashVoucher("OUT", BoxOfType("MAIN"), Null, "OWNER", 2000.00, "«·„«·ﬂ", "„”ÕÊ»«  «·„«·ﬂ ( Ã—Ì»Ì)", Null, id), "”‰œ ‰ﬁœÌ…"
    MoveDoc "CASH_VOUCHER", id, when
    ' --- customer_payment (day -6)
    when = DemoWhen(6, 11)
    SetDemoUser ""
    Check PostCustomerPayment(CustomerIDOf("„ÿ⁄„ «·œÌ—…"), 300.00, 1, "œ›⁄… „‰ «·Õ”«»", id), "”‰œ ﬁ»÷"
    MoveDoc "CUSTOMER_PAYMENT", id, when
End Sub

Private Sub DemoSteps5()
    Dim when As Date, id As Long, refNo As String, adjusted As Long, netValue As Currency, counted As Currency
    ' --- expense (day -6)
    when = DemoWhen(6, 15)
    SetDemoUser ""
    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, TotalAmount, PaymentMethodID, Description, EmployeeID, CreatedAt, CashBoxID) VALUES (" & SqlText(NextNumber("EXPENSE")) & ", " & SqlDate(DateValue(when)) & ", 6, 260.00, 39.00, 299.00, 1, " & SqlText("’Ì«‰… À·«Ã… «·⁄—÷") & ", " & CurrentUserID() & ", " & SqlDate(when) & ", " & CurrentCashBoxID() & ")", dbFailOnError
    ' --- sale (day -5)
    when = DemoWhen(5, 18)
    SetDemoUser "cashier1"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000083", 4
    CartLine "6281000000106", 2
    CartLine "6281000000120", 3
    CartLine "6281000000175", 2
    Check PostSaleFromCart(1, "CASH", 1, 0.00, Null, "", id), "›« Ê—… »Ì⁄ 7"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(7) = id
    ' --- sales_return (day -4)
    when = DemoWhen(4, 10)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpReturnLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpReturnLines (SalesDetailID, ProductID, ReturnQty, ReturnToStock) SELECT SalesDetailID, ProductID, 2, True FROM SalesInvoiceDetails WHERE SalesInvoiceID = " & m_sales(6) & " AND LineNumber = 4", dbFailOnError
    Check PostSalesReturn(m_sales(6), " ·› ›Ì «· €·Ì› ( Ã—Ì»Ì)", "CREDIT", Null, id), "„— Ã⁄ „»Ì⁄« "
    CurrentDb.Execute "DELETE FROM tmpReturnLines", dbFailOnError
    MoveDoc "SALES_RETURN", id, when
    ' --- sale (day -3)
    when = DemoWhen(3, 16)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000090", 2
    CartLine "6281000000052", 1
    CartLine "6281000000199", 1
    Check PostSaleFromCart(CustomerIDOf("›Âœ ⁄»œ«·—Õ„‰ «·Õ—»Ì"), "CREDIT", 1, 0.00, 0, "", id), "›« Ê—… »Ì⁄ 8"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(8) = id
    ' --- customer_payment (day -2)
    when = DemoWhen(2, 10)
    SetDemoUser ""
    Check PostCustomerPayment(CustomerIDOf("„ƒ””… «·÷Ì«›… ·· „ÊÌ‰"), 600.00, 1, "œ›⁄… „‰ «·Õ”«»", id), "”‰œ ﬁ»÷"
    MoveDoc "CUSTOMER_PAYMENT", id, when
    ' --- stock_count (day -2)
    when = DemoWhen(2, 21)
    SetDemoUser ""
    Check CreateStockCount(CategoryIDOf("√œÊ«  „‰“·Ì…"), "Ã—œ  Ã—Ì»Ì", id), "≈‰‘«¡ «·Ã—œ"
    CurrentDb.Execute "UPDATE StockCountDetails SET ActualQuantity = SystemQuantity + (-1) WHERE StockCountID = " & id & " AND ProductID = " & ProductIDOf("6281000000199"), dbFailOnError
    Check PostStockCount(id, False, adjusted, netValue), " —ÕÌ· «·Ã—œ"
    MoveDoc "STOCK_COUNT", id, when
End Sub

Private Sub DemoSteps6()
    Dim when As Date, id As Long, refNo As String, adjusted As Long, netValue As Currency, counted As Currency
    ' --- sale (day -1)
    when = DemoWhen(1, 20)
    SetDemoUser "cashier2"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000076", 10
    CartLine "6281000000083", 6
    CartLine "6281000000182", 2
    Check PostSaleFromCart(1, "CASH", 1, 0.00, Null, "", id), "›« Ê—… »Ì⁄ 9"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(9) = id
    ' --- closing (day -1)
    when = DemoWhen(1, 23)
    SetDemoUser ""
    counted = CashBoxBalance(BoxOfType("CASHIER")) - 0.00
    Check PostCashClosing(BoxOfType("CASHIER"), counted, "MAIN", BoxOfType("MAIN"), IIf(counted > 100.00, counted - 100.00, 0), " ’›Ì…  Ã—Ì»Ì…", id), " ’›Ì… «·ﬂ«‘Ì—"
    MoveDoc "CASH_CLOSING", id, when
    ' --- sale (day -0)
    when = DemoWhen(0, 0)
    SetDemoUser "cashier1"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000038", 2
    CartLine "6281000000168", 3
    CartLine "6281000000144", 2
    Check PostSaleFromCart(1, "CASH", 1, 0.00, Null, "", id), "›« Ê—… »Ì⁄ 10"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(10) = id
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Private Function DemoWhen(ByVal DaysAgo As Long, ByVal HourOfDay As Long) As Date
    If DaysAgo = 0 Then
        DemoWhen = DateAdd("n", -2, Now)                     ' today: just now
    Else
        DemoWhen = DateValue(Date - DaysAgo) + TimeSerial(HourOfDay, 0, 0)
    End If
End Function

Private Sub SetDemoUser(ByVal Username As String)
    If Len(Username) = 0 Then
        TempVars.Add "UserID", m_adminID
    Else
        TempVars.Add "UserID", UserIDOf(Username)
    End If
End Sub

Private Sub Check(ByVal Msg As String, ByVal Label As String)
    If Len(Msg) > 0 Then Err.Raise vbObjectError + 811, "LoadDemoData", Label & ": " & Msg
End Sub

Private Function ProductIDOf(ByVal Barcode As String) As Long
    ProductIDOf = DbValue("SELECT ProductID FROM Products WHERE Barcode = " & SqlText(Barcode))
End Function

Private Function CustomerIDOf(ByVal PartyName As String) As Long
    CustomerIDOf = DbValue("SELECT CustomerID FROM Customers WHERE CustomerName = " & SqlText(PartyName))
End Function

Private Function SupplierIDOf(ByVal PartyName As String) As Long
    SupplierIDOf = DbValue("SELECT SupplierID FROM Suppliers WHERE SupplierName = " & SqlText(PartyName))
End Function

Private Function CategoryIDOf(ByVal CategoryName As String) As Long
    CategoryIDOf = DbValue("SELECT CategoryID FROM Categories WHERE CategoryName = " & SqlText(CategoryName))
End Function

Private Function BoxOfType(ByVal BoxType As String) As Long
    BoxOfType = Nz(DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE IsActive = True AND BoxType = " & SqlText(BoxType)), 0)
End Function

Private Function UserIDOf(ByVal Username As String) As Long
    UserIDOf = DbValue("SELECT EmployeeID FROM Employees WHERE Username = " & SqlText(Username))
End Function

Private Sub CartLine(ByVal Barcode As String, ByVal Qty As Long)
    CurrentDb.Execute "INSERT INTO tmpPOSLines (ProductID, ProductCode, ProductName, Quantity, UnitPrice, " & _
        "LineDiscount, Available) SELECT ProductID, ProductCode, ProductName, " & Qty & ", SellingPrice, 0, " & _
        "CurrentQuantity FROM Products WHERE Barcode = " & SqlText(Barcode), dbFailOnError
End Sub

Private Sub PurLine(ByVal Barcode As String, ByVal Qty As Long, ByVal Cost As String)
    CurrentDb.Execute "INSERT INTO tmpPurchaseLines (ProductID, ProductCode, ProductName, Quantity, UnitCost, " & _
        "LineDiscount, SellingPrice) SELECT ProductID, ProductCode, ProductName, " & Qty & ", " & Cost & _
        ", 0, SellingPrice FROM Products WHERE Barcode = " & SqlText(Barcode), dbFailOnError
End Sub

Private Sub ClearWorkTables()
    Dim t As Variant
    For Each t In Array("tmpPOSLines", "tmpReturnLines", "tmpPurchaseLines", "tmpPurchaseReturnLines")
        CurrentDb.Execute "DELETE FROM " & t, dbFailOnError
    Next
End Sub

Private Sub MoveDoc(ByVal Kind As String, ByVal DocID As Long, ByVal When As Date)
    ' Puts a posted document (and its stock movements) on its planned date.
    Dim tbl As String, key As String, fld As String, ref As String, rs As DAO.Recordset, vatNo As Variant
    Select Case Kind
        Case "SALE":             tbl = "SalesInvoices": key = "SalesInvoiceID": fld = "InvoiceDate"
        Case "SALES_RETURN":     tbl = "SalesReturns": key = "SalesReturnID": fld = "ReturnDate"
        Case "PURCHASE":         tbl = "PurchaseInvoices": key = "PurchaseInvoiceID": fld = "InvoiceDate"
        Case "PURCHASE_RETURN":  tbl = "PurchaseReturns": key = "PurchaseReturnID": fld = "ReturnDate"
        Case "CUSTOMER_PAYMENT": tbl = "CustomerPayments": key = "PaymentID": fld = "PaymentDate"
        Case "SUPPLIER_PAYMENT": tbl = "SupplierPayments": key = "PaymentID": fld = "PaymentDate"
        Case "STOCK_COUNT":      tbl = "StockCounts": key = "StockCountID": fld = "CountDate"
        Case "CASH_VOUCHER":     tbl = "CashVouchers": key = "CashVoucherID": fld = "VoucherDate"
        Case "CASH_CLOSING":     tbl = "CashClosings": key = "ClosingID": fld = "ClosingDate"
    End Select
    If Kind = "STOCK_COUNT" Then
        CurrentDb.Execute "UPDATE StockCounts SET CountDate = " & SqlDate(When) & ", PostedAt = " & SqlDate(When) & _
                          " WHERE StockCountID = " & DocID, dbFailOnError
    Else
        CurrentDb.Execute "UPDATE " & tbl & " SET " & fld & " = " & SqlDate(When) & ", CreatedAt = " & SqlDate(When) & _
                          " WHERE " & key & " = " & DocID, dbFailOnError
    End If
    If Kind = "CASH_CLOSING" Then          ' the shortage / transfer vouchers of the closing
        CurrentDb.Execute "UPDATE CashVouchers SET VoucherDate = " & SqlDate(When) & ", CreatedAt = " & SqlDate(When) & _
                          " WHERE ClosingID = " & DocID, dbFailOnError
    End If
    Select Case Kind
        Case "SALE", "SALES_RETURN", "PURCHASE", "PURCHASE_RETURN", "STOCK_COUNT"
            CurrentDb.Execute "UPDATE InventoryTransactions SET TransactionDate = " & SqlDate(When) & ", CreatedAt = " & _
                              SqlDate(When) & " WHERE ReferenceType = " & SqlText(Kind) & " AND ReferenceID = " & DocID, _
                              dbFailOnError
    End Select
    If Kind = "SALE" Or Kind = "SALES_RETURN" Then
        ' the ZATCA QR code carries the time stamp of the document
        vatNo = SettingValue("VATNumber")
        If Not IsNull(vatNo) Then
            Set rs = CurrentDb.OpenRecordset("SELECT QRCodeData, TotalAmount, Tax FROM " & tbl & " WHERE " & key & _
                                             " = " & DocID, dbOpenDynaset)
            rs.Edit
            rs!QRCodeData = BuildZatcaQR(Nz(SettingValue("StoreName"), ""), CStr(vatNo), When, rs!TotalAmount, rs!Tax)
            rs.Update
            rs.Close
        End If
    End If
End Sub

Private Sub MoveManual(ByVal RefNumber As String, ByVal When As Date)
    CurrentDb.Execute "UPDATE InventoryTransactions SET TransactionDate = " & SqlDate(When) & ", CreatedAt = " & _
                      SqlDate(When) & " WHERE ReferenceType = 'MANUAL' AND ReferenceNumber = " & SqlText(RefNumber), _
                      dbFailOnError
End Sub

'==============================================================================
' Verify
'==============================================================================
Public Function VerifyDemoData() As Boolean
    Dim n As Long
    m_passed = 0: m_failed = 0: m_report = ""
    Debug.Print "=== VerifyDemoData  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Calendar = vbCalGreg
    Expect DbValue("SELECT COUNT(*) FROM Products WHERE Notes = '" & DEMO_MARK & "'") = 20, "20 „‰ Ã«"
    Expect DbValue("SELECT COUNT(*) FROM Categories WHERE Description = '" & DEMO_MARK & "'") = 5, "5  ’‰Ì›« "
    Expect DbValue("SELECT COUNT(*) FROM Customers WHERE Notes = '" & DEMO_MARK & "'") = 10, "10 ⁄„·«¡"
    Expect DbValue("SELECT COUNT(*) FROM Suppliers WHERE Notes = '" & DEMO_MARK & "'") = 5, "5 „Ê—œÌ‰"
    Expect DbValue("SELECT COUNT(*) FROM Employees WHERE Notes = '" & DEMO_MARK & "' AND MustChangePassword = True " & _
                   "AND PasswordHash Is Not Null") = 3, "3 „ÊŸ›Ì‰ »ﬂ·„… „—Ê— „ƒﬁ …"
    Expect DbValue("SELECT COUNT(*) FROM SalesInvoices") = 10, "10 ›Ê« Ì— »Ì⁄"
    Expect DbValue("SELECT COUNT(*) FROM PurchaseInvoices") = 5, "5 ›Ê« Ì— ‘—«¡"
    Expect DbValue("SELECT COUNT(*) FROM Expenses") = 6, "6 „’—Ê›« "
    Expect Nz(DbValue("SELECT Sum(TotalAmount) FROM SalesInvoices"), 0) = CCur(3289.90), "≈Ã„«·Ì «·„»Ì⁄«  3289.90"
    Expect Nz(DbValue("SELECT Sum(TotalAmount) FROM PurchaseInvoices"), 0) = CCur(9332.83), "≈Ã„«·Ì «·„‘ —Ì«  9332.83"
    Expect ProductValue("6281000000014", "CurrentQuantity") = 15 And ProductValue("6281000000014", "AverageCost") = CCur(38), "√—“ »”„ Ì 5 ﬂÃ„: «·ﬂ„Ì… 15 Ê«·„ Ê”ÿ 38"
    Expect ProductValue("6281000000021", "CurrentQuantity") = 13 And ProductValue("6281000000021", "AverageCost") = CCur(9.5), "”ﬂ— √»Ì÷ 2 ﬂÃ„: «·ﬂ„Ì… 13 Ê«·„ Ê”ÿ 9.5"
    Expect ProductValue("6281000000038", "CurrentQuantity") = 41 And ProductValue("6281000000038", "AverageCost") = CCur(14), "“Ì  œÊ«— «·‘„” 1.5 · —: «·ﬂ„Ì… 41 Ê«·„ Ê”ÿ 14"
    Expect ProductValue("6281000000045", "CurrentQuantity") = 50 And ProductValue("6281000000045", "AverageCost") = CCur(2.6), "„⁄ﬂ—Ê‰… 450 Ã„: «·ﬂ„Ì… 50 Ê«·„ Ê”ÿ 2.6"
    Expect ProductValue("6281000000052", "CurrentQuantity") = 27 And ProductValue("6281000000052", "AverageCost") = CCur(22), " „— ”ﬂ—Ì 1 ﬂÃ„: «·ﬂ„Ì… 27 Ê«·„ Ê”ÿ 22"
    Expect ProductValue("6281000000069", "CurrentQuantity") = 17 And ProductValue("6281000000069", "AverageCost") = CCur(11), "„Ì«Â „⁄œ‰Ì… 330 „· ◊ 40: «·ﬂ„Ì… 17 Ê«·„ Ê”ÿ 11"
    Expect ProductValue("6281000000076", "CurrentQuantity") = 64 And ProductValue("6281000000076", "AverageCost") = CCur(4.2), "⁄’Ì— »— ﬁ«· 1 · —: «·ﬂ„Ì… 64 Ê«·„ Ê”ÿ 4.2"
    Expect ProductValue("6281000000083", "CurrentQuantity") = 15 And ProductValue("6281000000083", "AverageCost") = CCur(7.8), "Õ·Ì» ÿ«“Ã 2 · —: «·ﬂ„Ì… 15 Ê«·„ Ê”ÿ 7.8"
    Expect ProductValue("6281000000090", "CurrentQuantity") = 22 And ProductValue("6281000000090", "AverageCost") = CCur(26), "ﬁÂÊ… ⁄—»Ì… 500 Ã„: «·ﬂ„Ì… 22 Ê«·„ Ê”ÿ 26"
    Expect ProductValue("6281000000106", "CurrentQuantity") = 38 And ProductValue("6281000000106", "AverageCost") = CCur(9), "‘«Ì √ﬂÌ«” ◊ 100: «·ﬂ„Ì… 38 Ê«·„ Ê”ÿ 9"
    Expect ProductValue("6281000000113", "CurrentQuantity") = 26 And ProductValue("6281000000113", "AverageCost") = CCur(11.5), "„‰Ÿ› √—÷Ì«  3 · —: «·ﬂ„Ì… 26 Ê«·„ Ê”ÿ 11.5"
    Expect ProductValue("6281000000120", "CurrentQuantity") = 36 And ProductValue("6281000000120", "AverageCost") = CCur(5.3), "”«∆· €”Ì· ’ÕÊ‰ 1 · —: «·ﬂ„Ì… 36 Ê«·„ Ê”ÿ 5.3"
    Expect ProductValue("6281000000137", "CurrentQuantity") = 6 And ProductValue("6281000000137", "AverageCost") = CCur(24), "„”ÕÊﬁ €”Ì· 3 ﬂÃ„: «·ﬂ„Ì… 6 Ê«·„ Ê”ÿ 24"
    Expect ProductValue("6281000000144", "CurrentQuantity") = 55 And ProductValue("6281000000144", "AverageCost") = CCur(8.4), "„‰«œÌ· Ê—ﬁÌ… ◊ 5: «·ﬂ„Ì… 55 Ê«·„ Ê”ÿ 8.4"
    Expect ProductValue("6281000000151", "CurrentQuantity") = 29 And ProductValue("6281000000151", "AverageCost") = CCur(12), "‘«„»Ê 400 „·: «·ﬂ„Ì… 29 Ê«·„ Ê”ÿ 12"
    Expect ProductValue("6281000000168", "CurrentQuantity") = 10 And ProductValue("6281000000168", "AverageCost") = CCur(6.1), "„⁄ÃÊ‰ √”‰«‰ 100 „·: «·ﬂ„Ì… 10 Ê«·„ Ê”ÿ 6.1"
    Expect ProductValue("6281000000175", "CurrentQuantity") = 38 And ProductValue("6281000000175", "AverageCost") = CCur(5.2), "’«»Ê‰ ÌœÌ‰ 500 „·: «·ﬂ„Ì… 38 Ê«·„ Ê”ÿ 5.2"
    Expect ProductValue("6281000000182", "CurrentQuantity") = 38 And ProductValue("6281000000182", "AverageCost") = CCur(6.7913), "√ﬂÌ«” ‰›«Ì«  ﬂ»Ì—… ◊ 30: «·ﬂ„Ì… 38 Ê«·„ Ê”ÿ 6.7913"
    Expect ProductValue("6281000000199", "CurrentQuantity") = 23 And ProductValue("6281000000199", "AverageCost") = CCur(8.9252), "Ê—ﬁ √·„‰ÌÊ„ 30 „: «·ﬂ„Ì… 23 Ê«·„ Ê”ÿ 8.9252"
    Expect ProductValue("6281000000205", "CurrentQuantity") = 20 And ProductValue("6281000000205", "AverageCost") = CCur(7.761), "»ÿ«—Ì«  AA ◊ 4: «·ﬂ„Ì… 20 Ê«·„ Ê”ÿ 7.761"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("„ÿ⁄„ «·œÌ—…")), -1) = CCur(422.20), "—’Ìœ „ÿ⁄„ «·œÌ—… = 422.20"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("„ƒ””… «·÷Ì«›… ·· „ÊÌ‰")), -1) = CCur(51.50), "—’Ìœ „ƒ””… «·÷Ì«›… ·· „ÊÌ‰ = 51.50"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("„ﬁÂÏ «·—ﬂ‰")), -1) = CCur(0.00), "—’Ìœ „ﬁÂÏ «·—ﬂ‰ = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("√Õ„œ „Õ„œ «·⁄ Ì»Ì")), -1) = CCur(58.00), "—’Ìœ √Õ„œ „Õ„œ «·⁄ Ì»Ì = 58.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("”«—… ⁄»œ«··Â «·ﬁÕÿ«‰Ì")), -1) = CCur(0.00), "—’Ìœ ”«—… ⁄»œ«··Â «·ﬁÕÿ«‰Ì = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("Œ«·œ ≈»—«ÂÌ„ «·‘Â—Ì")), -1) = CCur(0.00), "—’Ìœ Œ«·œ ≈»—«ÂÌ„ «·‘Â—Ì = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("‰Ê—… ”⁄œ «·œÊ”—Ì")), -1) = CCur(0.00), "—’Ìœ ‰Ê—… ”⁄œ «·œÊ”—Ì = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("›Âœ ⁄»œ«·—Õ„‰ «·Õ—»Ì")), -1) = CCur(127.00), "—’Ìœ ›Âœ ⁄»œ«·—Õ„‰ «·Õ—»Ì = 127.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("„‰Ï ⁄·Ì «·“Â—«‰Ì")), -1) = CCur(0.00), "—’Ìœ „‰Ï ⁄·Ì «·“Â—«‰Ì = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("⁄»œ«··Â ‰«’— «·„ÿÌ—Ì")), -1) = CCur(0.00), "—’Ìœ ⁄»œ«··Â ‰«’— «·„ÿÌ—Ì = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierName = " & SqlText("„ƒ””… «·Ê›—… ··„Ê«œ «·€–«∆Ì…")), -1) = CCur(1607.80), "—’Ìœ „ƒ””… «·Ê›—… ··„Ê«œ «·€–«∆Ì… = 1607.80"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierName = " & SqlText("‘—ﬂ… «·Ì‰«»Ì⁄ ··„‘—Ê»« ")), -1) = CCur(0.00), "—’Ìœ ‘—ﬂ… «·Ì‰«»Ì⁄ ··„‘—Ê»«  = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierName = " & SqlText("„ƒ””… «·‰ﬁ«¡ ··„‰Ÿ›« ")), -1) = CCur(601.27), "—’Ìœ „ƒ””… «·‰ﬁ«¡ ··„‰Ÿ›«  = 601.27"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierName = " & SqlText("‘—ﬂ… «·⁄‰«Ì… «·–Â»Ì…")), -1) = CCur(258.43), "—’Ìœ ‘—ﬂ… «·⁄‰«Ì… «·–Â»Ì… = 258.43"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierName = " & SqlText("„ƒ””… «·»Ì  «·⁄’—Ì ··√œÊ« ")), -1) = CCur(0.00), "—’Ìœ „ƒ””… «·»Ì  «·⁄’—Ì ··√œÊ«  = 0.00"
    Expect CashBoxBalance(BoxOfType("MAIN")) = CCur(17705.40), "—’Ìœ «·Œ“Ì‰… «·—∆Ì”Ì… = 17705.40"
    Expect CashBoxBalance(BoxOfType("CASHIER")) = CCur(194.30), "—’Ìœ ’‰œÊﬁ «·ﬂ«‘Ì— = 194.30"
    Expect DbValue("SELECT COUNT(*) FROM LowStockQuery AS l INNER JOIN Products AS p ON l.ProductID = p.ProductID " & _
                   "WHERE p.Notes = '" & DEMO_MARK & "'") = 4, "4 „‰ Ã«  „‰Œ›÷… «·„Œ“Ê‰ (·· ‰»ÌÂ Ê«· ﬁ—Ì—)"
    If Nz(SettingValue("SlowMovingDays"), 90) = 90 Then
        Expect DbValue("SELECT COUNT(*) FROM SlowMovingProductsQuery AS s INNER JOIN Products AS p ON s.ProductID = " & _
                       "p.ProductID WHERE p.Notes = '" & DEMO_MARK & "'") = 1, "1 „‰ Ã €Ì— „ Õ—ﬂ"
    End If
    Expect DbValue("SELECT COUNT(*) FROM StockCounts WHERE Status = 'POSTED'") = 1, "Ã—œ „ı—Õ¯· Ê«Õœ"
    Expect DbValue("SELECT COUNT(*) FROM IntegrityCheckQuery") = 0, "›Õ’ ”·«„… «·»Ì«‰« : ·«  ÊÃœ √Ì „‘ﬂ·…"
    n = Nz(DbValue("SELECT COUNT(*) FROM JournalEntries"), 0)
    Expect n = 34, "34 ﬁÌœ ÌÊ„Ì… («·›⁄·Ì: " & n & ")"
    ExpectJournal "SALE", 10, CCur(-2289.51)
    ExpectJournal "SALES_RETURN", 1, CCur(22.00)
    ExpectJournal "PURCHASE", 5, CCur(8115.50)
    ExpectJournal "PURCHASE_RETURN", 1, CCur(-34.50)
    ExpectJournal "CUSTOMER_PAYMENT", 2, CCur(0.00)
    ExpectJournal "SUPPLIER_PAYMENT", 2, CCur(0.00)
    ExpectJournal "EXPENSE", 6, CCur(0.00)
    ExpectJournal "CASH_VOUCHER", 5, CCur(0.00)
    ExpectJournal "STOCK_MOVE", 1, CCur(-5.30)
    ExpectJournal "STOCK_COUNT", 1, CCur(-8.93)
    ExpectJournal "BOX_OPENING", 0, CCur(0.00)
    ExpectJournal "CUSTOMER_OPENING", 0, CCur(0.00)
    ExpectJournal "SUPPLIER_OPENING", 0, CCur(0.00)
    ExpectJournal "MANUAL", 0, CCur(0.00)
    Expect DbValue("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit") = 0, "ﬂ· «·ﬁÌÊœ „ Ê«“‰…"
    Expect AccountBalance(1300) = Nz(DbValue("SELECT Sum(CurrentBalance) FROM Customers"), 0), "Õ”«» –„„ «·⁄„·«¡ = √—’œ… «·⁄„·«¡"
    Expect Round(AccountBalance(1400), 2) = CCur(5799.27), "Õ”«» «·„Œ“Ê‰ ›Ì «·ﬁÌÊœ = 5799.27 («·›⁄·Ì: " & _
           Format$(AccountBalance(1400), "0.00##") & ")"
    Debug.Print "--- ‰ÃÕ: " & m_passed & " | ›‘·: " & m_failed
    If m_failed = 0 Then
        TestMsg "«·»Ì«‰«  «· Ã—Ì»Ì… „ÿ«»ﬁ…  „«„« ··‰ «∆Ã «·„Õ”Ê»… „”»ﬁ« (" & m_passed & " ›Õ’«)." & vbCrLf & _
                "«·„” Œœ„Ê‰ «· Ã—Ì»ÌÊ‰: manager Ê cashier1 Ê cashier2° Êﬂ·„… «·„—Ê— «·„ƒﬁ …: " & DEMO_PASSWORD, _
                vbInformation + MSG_RTL, "VerifyDemoData"
        VerifyDemoData = True
    Else
        TestMsg "‰ÃÕ " & m_passed & " Ê›‘· " & m_failed & ":" & vbCrLf & vbCrLf & Left$(m_report, 900), _
                vbExclamation + MSG_RTL, "VerifyDemoData"
    End If
End Function

Private Function ProductValue(ByVal Barcode As String, ByVal FieldName As String) As Currency
    ProductValue = Nz(DbValue("SELECT " & FieldName & " FROM Products WHERE Barcode = " & SqlText(Barcode)), -1)
End Function

Private Sub ExpectJournal(ByVal Kind As String, ByVal Entries As Long, ByVal Stock As Currency)
    ' Entries of one kind of operation and their effect on the stock account 1400.
    Dim n As Long, v As Currency
    n = Nz(DbValue("SELECT COUNT(*) FROM JournalEntries WHERE SourceType = '" & Kind & "'"), 0)
    v = Nz(DbValue("SELECT Sum(l.Debit) - Sum(l.Credit) FROM JournalLines AS l INNER JOIN JournalEntries AS e " & _
                   "ON l.EntryID = e.EntryID WHERE e.SourceType = '" & Kind & "' AND l.AccountCode = 1400"), 0)
    Expect n = Entries And Round(v, 2) = Stock, "ﬁÌÊœ " & Kind & ": " & Entries & " ﬁÌœ° «·„Œ“Ê‰ " & _
           Format$(Stock, "0.00") & " («·›⁄·Ì: " & n & " ﬁÌœ° " & Format$(v, "0.00##") & ")"
End Sub

Private Sub Expect(ByVal Passed As Boolean, ByVal Label As String)
    If Passed Then
        m_passed = m_passed + 1
        Debug.Print "[OK] " & Label
    Else
        m_failed = m_failed + 1
        m_report = m_report & "- " & Label & vbCrLf
        Debug.Print "[X] " & Label
    End If
End Sub

'==============================================================================
' Remove (before going live)
'==============================================================================
Public Function RemoveDemoData() As Boolean
    Dim loaded As Variant, later As Long, t As Variant, ws As DAO.Workspace, inTrans As Boolean
    Dim db As DAO.Database, msg As String, file As String, errText As String
    On Error GoTo EH
    Calendar = vbCalGreg
    EnsureTestUser
    If Not IsAdministrator() Then
        MsgBox "Õ–› «·»Ì«‰«  «· Ã—Ì»Ì… ·„œÌ— «·‰Ÿ«„ ›ﬁÿ.", vbExclamation + MSG_RTL, "RemoveDemoData"
        Exit Function
    End If
    loaded = DMax("LogDate", "AuditLog", "ActionType = 'DEMO_LOADED'")
    If IsNull(loaded) Then
        MsgBox "·„  ıœŒ· »Ì«‰«   Ã—Ì»Ì… ›Ì Â–Â «·ﬁ«⁄œ….", vbInformation + MSG_RTL, "RemoveDemoData"
        Exit Function
    End If
    For Each t In Array("SalesInvoices", "SalesReturns", "PurchaseInvoices", "PurchaseReturns", "CustomerPayments", _
                        "SupplierPayments", "Expenses", "CashVouchers", "CashClosings")
        later = later + Nz(DbValue("SELECT COUNT(*) FROM [" & t & "] WHERE CreatedAt > " & SqlDate(loaded)), 0)
    Next
    later = later + Nz(DbValue("SELECT COUNT(*) FROM StockCounts WHERE CountDate > " & SqlDate(loaded)), 0)
    If later > 0 Then
        MsgBox " ÊÃœ " & later & " „” ‰œ«  √ıœŒ·  »⁄œ «·»Ì«‰«  «· Ã—Ì»Ì…° ›·‰  ıÕ–› √Ì »Ì«‰« ." & vbCrLf & _
               "(«·Õ–› Â‰« ·« Ì„Ì“ »Ì‰ «· Ã—Ì»Ì Ê«·ÕﬁÌﬁÌ° ·–·ﬂ Ìı—›÷ ·Õ„«Ì… »Ì«‰« ﬂ.)", vbExclamation + MSG_RTL, _
               "RemoveDemoData"
        Exit Function
    End If
    If MsgBox("”Ì „ Õ–› ﬂ· «·„” ‰œ«  Ê«·»Ì«‰«  «· Ã—Ì»Ì… Ê≈⁄«œ…  —ﬁÌ„ «·„” ‰œ«  „‰ 1." & vbCrLf & _
              " ıÕ›Ÿ ‰”Œ… «Õ Ì«ÿÌ… √Ê·«. Â·  —Ìœ «·„ «»⁄…ø", vbQuestion + vbYesNo + vbDefaultButton2 + MSG_RTL, _
              "RemoveDemoData") <> vbYes Then Exit Function
    msg = BackupNow(file, , "_BeforeDemoRemoval")
    If Len(msg) > 0 Then
        MsgBox " ⁄–—  «·‰”Œ… «·«Õ Ì«ÿÌ…° ›·„ ÌıÕ–› ‘Ì¡:" & vbCrLf & msg, vbExclamation + MSG_RTL, "RemoveDemoData"
        Exit Function
    End If
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    db.Execute "DELETE FROM JournalLines", dbFailOnError
    db.Execute "DELETE FROM JournalEntries", dbFailOnError
    db.Execute "DELETE FROM CashVouchers", dbFailOnError
    db.Execute "DELETE FROM CashClosings", dbFailOnError
    db.Execute "DELETE FROM SalesReturnDetails", dbFailOnError
    db.Execute "DELETE FROM SalesReturns", dbFailOnError
    db.Execute "DELETE FROM CustomerPayments", dbFailOnError
    db.Execute "DELETE FROM SalesInvoiceDetails", dbFailOnError
    db.Execute "DELETE FROM SalesInvoices", dbFailOnError
    db.Execute "DELETE FROM PurchaseReturnDetails", dbFailOnError
    db.Execute "DELETE FROM PurchaseReturns", dbFailOnError
    db.Execute "DELETE FROM SupplierPayments", dbFailOnError
    db.Execute "DELETE FROM PurchaseInvoiceDetails", dbFailOnError
    db.Execute "DELETE FROM PurchaseInvoices", dbFailOnError
    db.Execute "DELETE FROM StockCountDetails", dbFailOnError
    db.Execute "DELETE FROM StockCounts", dbFailOnError
    db.Execute "DELETE FROM InventoryTransactions", dbFailOnError
    db.Execute "DELETE FROM Expenses", dbFailOnError
    db.Execute "DELETE FROM Products WHERE Notes = 'DEMO'", dbFailOnError
    db.Execute "UPDATE Products SET CurrentQuantity = 0", dbFailOnError
    db.Execute "DELETE FROM Customers WHERE Notes = 'DEMO'", dbFailOnError
    db.Execute "UPDATE Customers SET CurrentBalance = OpeningBalance", dbFailOnError
    db.Execute "DELETE FROM Suppliers WHERE Notes = 'DEMO' AND SupplierID NOT IN (SELECT SupplierID FROM Products WHERE SupplierID Is Not Null)", dbFailOnError
    db.Execute "UPDATE Suppliers SET CurrentBalance = OpeningBalance", dbFailOnError
    db.Execute "DELETE FROM Categories WHERE Description = 'DEMO' AND CategoryID NOT IN (SELECT CategoryID FROM Products)", dbFailOnError
    db.Execute "DELETE FROM AuditLog WHERE EmployeeID IN (SELECT EmployeeID FROM Employees WHERE Notes = 'DEMO' AND Username IN ('manager', 'cashier1', 'cashier2'))", dbFailOnError
    db.Execute "DELETE FROM Employees WHERE Notes = 'DEMO' AND Username IN ('manager', 'cashier1', 'cashier2')", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'SALES_INVOICE'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'SALES_RETURN'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'PURCHASE_INVOICE'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'PURCHASE_RETURN'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'CUSTOMER_PAYMENT'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'SUPPLIER_PAYMENT'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'EXPENSE'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'STOCK_COUNT'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'STOCK_ADJUST'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'ZATCA_ICV'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'CASH_IN'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'CASH_OUT'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'CASH_TRANSFER'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'CASH_CLOSING'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'JOURNAL'", dbFailOnError
    db.Execute "UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'PRODUCT_CODE' AND (SELECT COUNT(*) FROM Products) = 0", dbFailOnError
    ws.CommitTrans
    inTrans = False
    LogAction "DEMO_REMOVED", "", "", file
    MsgBox " „ Õ–› «·»Ì«‰«  «· Ã—Ì»Ì…. «·ﬁ«⁄œ… Ã«Â“… ·»Ì«‰« ﬂ «·›⁄·Ì…." & vbCrLf & _
           "«·‰”Œ… ﬁ»· «·Õ–›: " & file, vbInformation + MSG_RTL, "RemoveDemoData"
    RemoveDemoData = True
    Exit Function
EH:
    errText = Err.Description
    On Error Resume Next
    If inTrans Then ws.Rollback
    MsgBox " ⁄–— «·Õ–› Ê·„ Ì €Ì— ‘Ì¡: " & errText, vbCritical + MSG_RTL, "RemoveDemoData"
End Function
