Attribute VB_Name = "modDemoData"
'==============================================================================
' modDemoData  -  Retail Store Management System (Phase 11: demo data)
'
' GENERATED FILE - do not edit by hand.  Source: tools/demo_data.py, tools/gen_demo.py
'
'   LoadDemoData     enters the demo store: 20 منتجًا، 5 تصنيفات، 10 عملاء، 5 موردين، 3 موظفين، 10 فواتير بيع، 5 فواتير شراء، 6 مصروفات.
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
    If MsgBox("سيتم إدخال متجر تجريبي كامل:" & vbCrLf & "20 منتجًا، 5 تصنيفات، 10 عملاء، 5 موردين، 3 موظفين، 10 فواتير بيع، 5 فواتير شراء، 6 مصروفات." & vbCrLf & vbCrLf & _
              "للتدريب والتجربة فقط. قبل الاستخدام الفعلي احذفها بالأمر RemoveDemoData." & vbCrLf & _
              "هل تريد المتابعة؟", vbQuestion + vbYesNo + vbDefaultButton2 + MSG_RTL, "LoadDemoData") <> vbYes Then
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
    Check SyncJournal(), "قيود اليومية"
    ws.CommitTrans
    inTrans = False
    TempVars.Add "UserID", m_adminID
    LogAction "DEMO_LOADED", "", "", "20 منتجًا، 5 تصنيفات، 10 عملاء، 5 موردين، 3 موظفين، 10 فواتير بيع، 5 فواتير شراء، 6 مصروفات"
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
    MsgBox "تعذر إدخال البيانات التجريبية، ولم يُحفظ منها أي شيء:" & vbCrLf & vbCrLf & errText, _
           vbCritical + MSG_RTL, "LoadDemoData"
End Function

Private Function DemoBlocker() As String
    ' Reasons not to load: documents already exist, demo already there, settings differ from the plan.
    Dim t As Variant
    If Not IsAdministrator() Then
        DemoBlocker = "إدخال البيانات التجريبية لمدير النظام فقط."
        Exit Function
    End If
    For Each t In Array("SalesInvoices", "PurchaseInvoices", "InventoryTransactions", "Expenses", _
                        "CustomerPayments", "SupplierPayments", "StockCounts", "CashVouchers", "CashClosings")
        If Nz(DbValue("SELECT COUNT(*) FROM [" & t & "]"), 0) > 0 Then
            DemoBlocker = "توجد بيانات فعلية (" & t & ")." & vbCrLf & _
                          "البيانات التجريبية تُدخل في قاعدة بدون مستندات فقط (نسخة للتدريب)."
            Exit Function
        End If
    Next
    If Nz(DbValue("SELECT COUNT(*) FROM Products WHERE Barcode = " & SqlText("6281000000014")), 0) > 0 Or _
       Nz(DbValue("SELECT COUNT(*) FROM Employees WHERE Username = 'manager' OR Username = 'cashier1' " & _
                  "OR Username = 'cashier2'"), 0) > 0 Then
        DemoBlocker = "يوجد جزء من البيانات التجريبية مسبقًا (منتج أو مستخدم بنفس الاسم)."
        Exit Function
    End If
    If Nz(SettingValue("VATRate"), 0) <> 0.15 Or Not Nz(SettingValue("PricesIncludeVAT"), False) Then
        DemoBlocker = "البيانات التجريبية مبنية على نسبة ضريبة 15% وأسعار شاملة الضريبة." & vbCrLf & _
                      "أعد هذين الإعدادين ثم حاول مرة أخرى."
    End If
End Function

Private Sub DemoSettings()
    ' The demo store details are used only when the store is not set up yet.
    If Not IsNull(SettingValue("VATNumber")) Then Exit Sub
    CurrentDb.Execute "UPDATE Settings SET StoreName = 'متجر النخبة للتجزئة', StoreNameEn = 'Al Nukhba Retail Store', VATNumber = '310123456700003', CRNumber = '1010654321', BuildingNo = '2345', StreetName = 'طريق الملك فهد', District = 'العليا', City = 'الرياض', PostalCode = '12211', Phone = '0112345678' WHERE SettingID = 1", dbFailOnError
End Sub

Private Sub DemoMasters()
    CurrentDb.Execute "INSERT INTO Categories (CategoryName, Description) VALUES (" & SqlText("مواد غذائية") & ", 'DEMO')", dbFailOnError
    CurrentDb.Execute "INSERT INTO Categories (CategoryName, Description) VALUES (" & SqlText("مشروبات") & ", 'DEMO')", dbFailOnError
    CurrentDb.Execute "INSERT INTO Categories (CategoryName, Description) VALUES (" & SqlText("منظفات") & ", 'DEMO')", dbFailOnError
    CurrentDb.Execute "INSERT INTO Categories (CategoryName, Description) VALUES (" & SqlText("عناية شخصية") & ", 'DEMO')", dbFailOnError
    CurrentDb.Execute "INSERT INTO Categories (CategoryName, Description) VALUES (" & SqlText("أدوات منزلية") & ", 'DEMO')", dbFailOnError
    CurrentDb.Execute "INSERT INTO Suppliers (SupplierName, ContactPerson, Mobile, VATNumber, City, Notes, CreatedAt) VALUES (" & SqlText("مؤسسة الوفرة للمواد الغذائية") & ", " & SqlText("سالم العمري") & ", '0551110001', '300112233400003', " & SqlText("الرياض") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Suppliers (SupplierName, ContactPerson, Mobile, VATNumber, City, Notes, CreatedAt) VALUES (" & SqlText("شركة الينابيع للمشروبات") & ", " & SqlText("ماجد الحربي") & ", '0551110002', '300223344500003', " & SqlText("جدة") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Suppliers (SupplierName, ContactPerson, Mobile, VATNumber, City, Notes, CreatedAt) VALUES (" & SqlText("مؤسسة النقاء للمنظفات") & ", " & SqlText("تركي الغامدي") & ", '0551110003', '300334455600003', " & SqlText("الدمام") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Suppliers (SupplierName, ContactPerson, Mobile, VATNumber, City, Notes, CreatedAt) VALUES (" & SqlText("شركة العناية الذهبية") & ", " & SqlText("هند السبيعي") & ", '0551110004', '300445566700003', " & SqlText("الرياض") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Suppliers (SupplierName, ContactPerson, Mobile, VATNumber, City, Notes, CreatedAt) VALUES (" & SqlText("مؤسسة البيت العصري للأدوات") & ", " & SqlText("وليد الشمري") & ", '0551110005', '300556677800003', " & SqlText("الرياض") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000014', " & SqlText("أرز بسمتي 5 كجم") & ", " & CategoryIDOf("مواد غذائية") & ", 4, 38.00, 52.00, 10, " & SupplierIDOf("مؤسسة الوفرة للمواد الغذائية") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000021', " & SqlText("سكر أبيض 2 كجم") & ", " & CategoryIDOf("مواد غذائية") & ", 4, 9.50, 13.50, 15, " & SupplierIDOf("مؤسسة الوفرة للمواد الغذائية") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000038', " & SqlText("زيت دوار الشمس 1.5 لتر") & ", " & CategoryIDOf("مواد غذائية") & ", 1, 14.00, 19.95, 12, " & SupplierIDOf("مؤسسة الوفرة للمواد الغذائية") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000045', " & SqlText("معكرونة 450 جم") & ", " & CategoryIDOf("مواد غذائية") & ", 1, 2.60, 3.95, 30, " & SupplierIDOf("مؤسسة الوفرة للمواد الغذائية") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000052', " & SqlText("تمر سكري 1 كجم") & ", " & CategoryIDOf("مواد غذائية") & ", 2, 22.00, 34.50, 8, " & SupplierIDOf("مؤسسة الوفرة للمواد الغذائية") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000069', " & SqlText("مياه معدنية 330 مل × 40") & ", " & CategoryIDOf("مشروبات") & ", 3, 11.00, 16.50, 10, " & SupplierIDOf("شركة الينابيع للمشروبات") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000076', " & SqlText("عصير برتقال 1 لتر") & ", " & CategoryIDOf("مشروبات") & ", 1, 4.20, 6.50, 20, " & SupplierIDOf("شركة الينابيع للمشروبات") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000083', " & SqlText("حليب طازج 2 لتر") & ", " & CategoryIDOf("مشروبات") & ", 1, 7.80, 11.00, 15, " & SupplierIDOf("شركة الينابيع للمشروبات") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000090', " & SqlText("قهوة عربية 500 جم") & ", " & CategoryIDOf("مشروبات") & ", 4, 26.00, 39.00, 6, " & SupplierIDOf("شركة الينابيع للمشروبات") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000106', " & SqlText("شاي أكياس × 100") & ", " & CategoryIDOf("مشروبات") & ", 2, 9.00, 14.25, 10, " & SupplierIDOf("شركة الينابيع للمشروبات") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000113', " & SqlText("منظف أرضيات 3 لتر") & ", " & CategoryIDOf("منظفات") & ", 1, 11.50, 17.25, 8, " & SupplierIDOf("مؤسسة النقاء للمنظفات") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000120', " & SqlText("سائل غسيل صحون 1 لتر") & ", " & CategoryIDOf("منظفات") & ", 1, 5.30, 8.50, 12, " & SupplierIDOf("مؤسسة النقاء للمنظفات") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000137', " & SqlText("مسحوق غسيل 3 كجم") & ", " & CategoryIDOf("منظفات") & ", 4, 24.00, 36.00, 6, " & SupplierIDOf("مؤسسة النقاء للمنظفات") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000144', " & SqlText("مناديل ورقية × 5") & ", " & CategoryIDOf("منظفات") & ", 4, 8.40, 12.95, 15, " & SupplierIDOf("مؤسسة النقاء للمنظفات") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000151', " & SqlText("شامبو 400 مل") & ", " & CategoryIDOf("عناية شخصية") & ", 1, 12.00, 18.95, 8, " & SupplierIDOf("شركة العناية الذهبية") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000168', " & SqlText("معجون أسنان 100 مل") & ", " & CategoryIDOf("عناية شخصية") & ", 1, 6.10, 9.50, 12, " & SupplierIDOf("شركة العناية الذهبية") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000175', " & SqlText("صابون يدين 500 مل") & ", " & CategoryIDOf("عناية شخصية") & ", 1, 5.20, 8.25, 10, " & SupplierIDOf("شركة العناية الذهبية") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000182', " & SqlText("أكياس نفايات كبيرة × 30") & ", " & CategoryIDOf("أدوات منزلية") & ", 4, 7.00, 10.95, 10, " & SupplierIDOf("مؤسسة البيت العصري للأدوات") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000199', " & SqlText("ورق ألمنيوم 30 م") & ", " & CategoryIDOf("أدوات منزلية") & ", 1, 9.20, 14.50, 6, " & SupplierIDOf("مؤسسة البيت العصري للأدوات") & ", 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & SqlText(NextNumber("PRODUCT_CODE")) & ", '6281000000205', " & SqlText("بطاريات AA × 4") & ", " & CategoryIDOf("أدوات منزلية") & ", 4, 8.00, 13.00, 5, " & SupplierIDOf("مؤسسة البيت العصري للأدوات") & ", 'DEMO', " & SqlDate(Date - 150) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("مطعم الديرة") & ", '0501000001', '300667788900003', " & SqlText("الرياض") & ", True, 5000, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("مؤسسة الضيافة للتموين") & ", '0501000002', '300778899000003', " & SqlText("الرياض") & ", True, 10000, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("مقهى الركن") & ", '0501000003', '300889900100003', " & SqlText("الرياض") & ", True, 3000, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("أحمد محمد العتيبي") & ", '0501000004', Null, " & SqlText("الرياض") & ", True, 1000, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("سارة عبدالله القحطاني") & ", '0501000005', Null, " & SqlText("الرياض") & ", True, 500, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("خالد إبراهيم الشهري") & ", '0501000006', Null, " & SqlText("الرياض") & ", False, 0, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("نورة سعد الدوسري") & ", '0501000007', Null, " & SqlText("الرياض") & ", False, 0, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("فهد عبدالرحمن الحربي") & ", '0501000008', Null, " & SqlText("الرياض") & ", True, 1500, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("منى علي الزهراني") & ", '0501000009', Null, " & SqlText("الرياض") & ", False, 0, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, CreditLimit, Notes, CreatedAt) VALUES (" & SqlText("عبدالله ناصر المطيري") & ", '0501000010', Null, " & SqlText("الرياض") & ", True, 2000, 'DEMO', " & SqlDate(Date - 60) & ")", dbFailOnError
    CurrentDb.Execute "INSERT INTO Employees (EmployeeName, JobTitle, Username, RoleID, MaxDiscountPercent, Notes) VALUES (" & SqlText("مشرف الفرع (تجريبي)") & ", " & SqlText("مدير فرع") & ", 'manager', 2, 0.10, 'DEMO')", dbFailOnError
    Check SetUserPassword(UserIDOf("manager"), DEMO_PASSWORD, True), "كلمة مرور manager"
    CurrentDb.Execute "INSERT INTO Employees (EmployeeName, JobTitle, Username, RoleID, MaxDiscountPercent, Notes) VALUES (" & SqlText("كاشير الصباح (تجريبي)") & ", " & SqlText("كاشير") & ", 'cashier1', 3, 0.02, 'DEMO')", dbFailOnError
    Check SetUserPassword(UserIDOf("cashier1"), DEMO_PASSWORD, True), "كلمة مرور cashier1"
    CurrentDb.Execute "INSERT INTO Employees (EmployeeName, JobTitle, Username, RoleID, MaxDiscountPercent, Notes) VALUES (" & SqlText("كاشير المساء (تجريبي)") & ", " & SqlText("كاشير") & ", 'cashier2', 3, 0.02, 'DEMO')", dbFailOnError
    Check SetUserPassword(UserIDOf("cashier2"), DEMO_PASSWORD, True), "كلمة مرور cashier2"
End Sub

Private Sub DemoSteps1()
    Dim when As Date, id As Long, refNo As String, adjusted As Long, netValue As Currency, counted As Currency
    ' --- cash_in (day -31)
    when = DemoWhen(31, 8)
    SetDemoUser ""
    Check PostCashVoucher("IN", BoxOfType("MAIN"), Null, "OWNER", 30000.00, "المالك", "رأس مال تشغيل (تجريبي)", Null, id), "سند نقدية"
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
    Check PostPurchaseFromCart(SupplierIDOf("مؤسسة الوفرة للمواد الغذائية"), "INV-7781", DateValue(when), "CREDIT", 1, True, 0.00, 1000, "", id), "فاتورة شراء 1"
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
    Check PostPurchaseFromCart(SupplierIDOf("شركة الينابيع للمشروبات"), "SP-20451", DateValue(when), "CASH", 1, True, 0.00, Null, "", id), "فاتورة شراء 2"
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    MoveDoc "PURCHASE", id, when
    m_purchases(2) = id
    ' --- expense (day -28)
    when = DemoWhen(28, 9)
    SetDemoUser ""
    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, TotalAmount, PaymentMethodID, Description, EmployeeID, CreatedAt, CashBoxID) VALUES (" & SqlText(NextNumber("EXPENSE")) & ", " & SqlDate(DateValue(when)) & ", 1, 4500.00, 0.00, 4500.00, 1, " & SqlText("إيجار المحل للشهر") & ", " & CurrentUserID() & ", " & SqlDate(when) & ", " & CurrentCashBoxID() & ")", dbFailOnError
    ' --- purchase (day -27)
    when = DemoWhen(27, 11)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    PurLine "6281000000113", 30, "11.50"
    PurLine "6281000000120", 40, "5.30"
    PurLine "6281000000137", 8, "24.00"
    PurLine "6281000000144", 60, "8.40"
    Check PostPurchaseFromCart(SupplierIDOf("مؤسسة النقاء للمنظفات"), "N-3390", DateValue(when), "CREDIT", 1, True, 0.00, 0, "", id), "فاتورة شراء 3"
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
    Check PostPurchaseFromCart(SupplierIDOf("شركة العناية الذهبية"), "GC-118", DateValue(when), "CREDIT", 1, True, 0.00, 500, "", id), "فاتورة شراء 4"
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
    Check PostPurchaseFromCart(SupplierIDOf("مؤسسة البيت العصري للأدوات"), "BA-5521", DateValue(when), "CASH", 1, True, 20.00, Null, "", id), "فاتورة شراء 5"
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
    Check PostSaleFromCart(1, "CASH", 1, 0.00, Null, "", id), "فاتورة بيع 1"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(1) = id
    ' --- expense (day -20)
    when = DemoWhen(20, 12)
    SetDemoUser ""
    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, TotalAmount, PaymentMethodID, Description, EmployeeID, CreatedAt, CashBoxID) VALUES (" & SqlText(NextNumber("EXPENSE")) & ", " & SqlDate(DateValue(when)) & ", 2, 380.00, 57.00, 437.00, 1, " & SqlText("فاتورة الكهرباء") & ", " & CurrentUserID() & ", " & SqlDate(when) & ", " & CurrentCashBoxID() & ")", dbFailOnError
    ' --- expense (day -19)
    when = DemoWhen(19, 12)
    SetDemoUser ""
    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, TotalAmount, PaymentMethodID, Description, EmployeeID, CreatedAt, CashBoxID) VALUES (" & SqlText(NextNumber("EXPENSE")) & ", " & SqlDate(DateValue(when)) & ", 3, 120.00, 18.00, 138.00, 1, " & SqlText("فاتورة المياه") & ", " & CurrentUserID() & ", " & SqlDate(when) & ", " & CurrentCashBoxID() & ")", dbFailOnError
    ' --- sale (day -18)
    when = DemoWhen(18, 11)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000014", 10
    CartLine "6281000000038", 6
    CartLine "6281000000069", 5
    Check PostSaleFromCart(CustomerIDOf("مطعم الديرة"), "CREDIT", 1, 0.00, 0, "", id), "فاتورة بيع 2"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(2) = id
    ' --- expense (day -16)
    when = DemoWhen(16, 13)
    SetDemoUser ""
    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, TotalAmount, PaymentMethodID, Description, EmployeeID, CreatedAt, CashBoxID) VALUES (" & SqlText(NextNumber("EXPENSE")) & ", " & SqlDate(DateValue(when)) & ", 4, 299.00, 44.85, 343.85, 1, " & SqlText("الإنترنت والهاتف") & ", " & CurrentUserID() & ", " & SqlDate(when) & ", " & CurrentCashBoxID() & ")", dbFailOnError
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
    Check PostSaleFromCart(1, "CASH", 1, 0.00, Null, "", id), "فاتورة بيع 3"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(3) = id
    ' --- supplier_payment (day -15)
    when = DemoWhen(15, 18)
    SetDemoUser ""
    Check PostSupplierPayment(SupplierIDOf("مؤسسة الوفرة للمواد الغذائية"), 1500.00, 1, "دفعة من الحساب", id), "سند صرف"
    MoveDoc "SUPPLIER_PAYMENT", id, when
    ' --- closing (day -15)
    when = DemoWhen(15, 23)
    SetDemoUser ""
    counted = CashBoxBalance(BoxOfType("CASHIER")) - 5.00
    Check PostCashClosing(BoxOfType("CASHIER"), counted, "MAIN", BoxOfType("MAIN"), IIf(counted > 100.00, counted - 100.00, 0), "تصفية تجريبية", id), "تصفية الكاشير"
    MoveDoc "CASH_CLOSING", id, when
    ' --- purchase_return (day -14)
    when = DemoWhen(14, 10)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpPurchaseReturnLines (PurchaseDetailID, ProductID, ReturnQty) SELECT PurchaseDetailID, ProductID, 3 FROM PurchaseInvoiceDetails WHERE PurchaseInvoiceID = " & m_purchases(3) & " AND LineNumber = 1", dbFailOnError
    Check PostPurchaseReturn(m_purchases(3), "عيب مصنعي (تجريبي)", "CREDIT", Null, id), "مرتجع مشتريات"
    CurrentDb.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError
    MoveDoc "PURCHASE_RETURN", id, when
    ' --- sale (day -12)
    when = DemoWhen(12, 13)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000052", 2
    CartLine "6281000000090", 1
    Check PostSaleFromCart(CustomerIDOf("أحمد محمد العتيبي"), "CREDIT", 1, 0.00, 50, "", id), "فاتورة بيع 4"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(4) = id
    ' --- expense (day -11)
    when = DemoWhen(11, 9)
    SetDemoUser ""
    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, TotalAmount, PaymentMethodID, Description, EmployeeID, CreatedAt, CashBoxID) VALUES (" & SqlText(NextNumber("EXPENSE")) & ", " & SqlDate(DateValue(when)) & ", 5, 150.00, 0.00, 150.00, 1, " & SqlText("نقل بضاعة") & ", " & CurrentUserID() & ", " & SqlDate(when) & ", " & CurrentCashBoxID() & ")", dbFailOnError
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
    Check PostSaleFromCart(1, "CASH", 1, 5.00, Null, "", id), "فاتورة بيع 5"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(5) = id
    ' --- supplier_payment (day -9)
    when = DemoWhen(9, 10)
    SetDemoUser ""
    Check PostSupplierPayment(SupplierIDOf("مؤسسة النقاء للمنظفات"), 800.00, 1, "دفعة من الحساب", id), "سند صرف"
    MoveDoc "SUPPLIER_PAYMENT", id, when
    ' --- sale (day -8)
    when = DemoWhen(8, 12)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000014", 15
    CartLine "6281000000021", 25
    CartLine "6281000000045", 60
    CartLine "6281000000069", 20
    Check PostSaleFromCart(CustomerIDOf("مؤسسة الضيافة للتموين"), "CREDIT", 1, 0.00, 1000, "", id), "فاتورة بيع 6"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(6) = id
    ' --- stock_out (day -7)
    when = DemoWhen(7, 9)
    SetDemoUser ""
    Check PostManualStock(ProductIDOf("6281000000120"), TT_STOCK_OUT, 1, Null, "تالف - عبوة مكسورة", refNo), "خصم مخزون"
    MoveManual refNo, when
    ' --- cash_out (day -7)
    when = DemoWhen(7, 20)
    SetDemoUser ""
    Check PostCashVoucher("OUT", BoxOfType("MAIN"), Null, "OWNER", 2000.00, "المالك", "مسحوبات المالك (تجريبي)", Null, id), "سند نقدية"
    MoveDoc "CASH_VOUCHER", id, when
    ' --- customer_payment (day -6)
    when = DemoWhen(6, 11)
    SetDemoUser ""
    Check PostCustomerPayment(CustomerIDOf("مطعم الديرة"), 300.00, 1, "دفعة من الحساب", id), "سند قبض"
    MoveDoc "CUSTOMER_PAYMENT", id, when
End Sub

Private Sub DemoSteps5()
    Dim when As Date, id As Long, refNo As String, adjusted As Long, netValue As Currency, counted As Currency
    ' --- expense (day -6)
    when = DemoWhen(6, 15)
    SetDemoUser ""
    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, TotalAmount, PaymentMethodID, Description, EmployeeID, CreatedAt, CashBoxID) VALUES (" & SqlText(NextNumber("EXPENSE")) & ", " & SqlDate(DateValue(when)) & ", 6, 260.00, 39.00, 299.00, 1, " & SqlText("صيانة ثلاجة العرض") & ", " & CurrentUserID() & ", " & SqlDate(when) & ", " & CurrentCashBoxID() & ")", dbFailOnError
    ' --- sale (day -5)
    when = DemoWhen(5, 18)
    SetDemoUser "cashier1"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000083", 4
    CartLine "6281000000106", 2
    CartLine "6281000000120", 3
    CartLine "6281000000175", 2
    Check PostSaleFromCart(1, "CASH", 1, 0.00, Null, "", id), "فاتورة بيع 7"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(7) = id
    ' --- sales_return (day -4)
    when = DemoWhen(4, 10)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpReturnLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpReturnLines (SalesDetailID, ProductID, ReturnQty, ReturnToStock) SELECT SalesDetailID, ProductID, 2, True FROM SalesInvoiceDetails WHERE SalesInvoiceID = " & m_sales(6) & " AND LineNumber = 4", dbFailOnError
    Check PostSalesReturn(m_sales(6), "تلف في التغليف (تجريبي)", "CREDIT", Null, id), "مرتجع مبيعات"
    CurrentDb.Execute "DELETE FROM tmpReturnLines", dbFailOnError
    MoveDoc "SALES_RETURN", id, when
    ' --- sale (day -3)
    when = DemoWhen(3, 16)
    SetDemoUser ""
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000090", 2
    CartLine "6281000000052", 1
    CartLine "6281000000199", 1
    Check PostSaleFromCart(CustomerIDOf("فهد عبدالرحمن الحربي"), "CREDIT", 1, 0.00, 0, "", id), "فاتورة بيع 8"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(8) = id
    ' --- customer_payment (day -2)
    when = DemoWhen(2, 10)
    SetDemoUser ""
    Check PostCustomerPayment(CustomerIDOf("مؤسسة الضيافة للتموين"), 600.00, 1, "دفعة من الحساب", id), "سند قبض"
    MoveDoc "CUSTOMER_PAYMENT", id, when
    ' --- stock_count (day -2)
    when = DemoWhen(2, 21)
    SetDemoUser ""
    Check CreateStockCount(CategoryIDOf("أدوات منزلية"), "جرد تجريبي", id), "إنشاء الجرد"
    CurrentDb.Execute "UPDATE StockCountDetails SET ActualQuantity = SystemQuantity + (-1) WHERE StockCountID = " & id & " AND ProductID = " & ProductIDOf("6281000000199"), dbFailOnError
    Check PostStockCount(id, False, adjusted, netValue), "ترحيل الجرد"
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
    Check PostSaleFromCart(1, "CASH", 1, 0.00, Null, "", id), "فاتورة بيع 9"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    MoveDoc "SALE", id, when
    m_sales(9) = id
    ' --- closing (day -1)
    when = DemoWhen(1, 23)
    SetDemoUser ""
    counted = CashBoxBalance(BoxOfType("CASHIER")) - 0.00
    Check PostCashClosing(BoxOfType("CASHIER"), counted, "MAIN", BoxOfType("MAIN"), IIf(counted > 100.00, counted - 100.00, 0), "تصفية تجريبية", id), "تصفية الكاشير"
    MoveDoc "CASH_CLOSING", id, when
    ' --- sale (day -0)
    when = DemoWhen(0, 0)
    SetDemoUser "cashier1"
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartLine "6281000000038", 2
    CartLine "6281000000168", 3
    CartLine "6281000000144", 2
    Check PostSaleFromCart(1, "CASH", 1, 0.00, Null, "", id), "فاتورة بيع 10"
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
    Expect DbValue("SELECT COUNT(*) FROM Products WHERE Notes = '" & DEMO_MARK & "'") = 20, "20 منتجًا"
    Expect DbValue("SELECT COUNT(*) FROM Categories WHERE Description = '" & DEMO_MARK & "'") = 5, "5 تصنيفات"
    Expect DbValue("SELECT COUNT(*) FROM Customers WHERE Notes = '" & DEMO_MARK & "'") = 10, "10 عملاء"
    Expect DbValue("SELECT COUNT(*) FROM Suppliers WHERE Notes = '" & DEMO_MARK & "'") = 5, "5 موردين"
    Expect DbValue("SELECT COUNT(*) FROM Employees WHERE Notes = '" & DEMO_MARK & "' AND MustChangePassword = True " & _
                   "AND PasswordHash Is Not Null") = 3, "3 موظفين بكلمة مرور مؤقتة"
    Expect DbValue("SELECT COUNT(*) FROM SalesInvoices") = 10, "10 فواتير بيع"
    Expect DbValue("SELECT COUNT(*) FROM PurchaseInvoices") = 5, "5 فواتير شراء"
    Expect DbValue("SELECT COUNT(*) FROM Expenses") = 6, "6 مصروفات"
    Expect Nz(DbValue("SELECT Sum(TotalAmount) FROM SalesInvoices"), 0) = CCur(3289.90), "إجمالي المبيعات 3289.90"
    Expect Nz(DbValue("SELECT Sum(TotalAmount) FROM PurchaseInvoices"), 0) = CCur(9332.83), "إجمالي المشتريات 9332.83"
    Expect ProductValue("6281000000014", "CurrentQuantity") = 15 And ProductValue("6281000000014", "AverageCost") = CCur(38), "أرز بسمتي 5 كجم: الكمية 15 والمتوسط 38"
    Expect ProductValue("6281000000021", "CurrentQuantity") = 13 And ProductValue("6281000000021", "AverageCost") = CCur(9.5), "سكر أبيض 2 كجم: الكمية 13 والمتوسط 9.5"
    Expect ProductValue("6281000000038", "CurrentQuantity") = 41 And ProductValue("6281000000038", "AverageCost") = CCur(14), "زيت دوار الشمس 1.5 لتر: الكمية 41 والمتوسط 14"
    Expect ProductValue("6281000000045", "CurrentQuantity") = 50 And ProductValue("6281000000045", "AverageCost") = CCur(2.6), "معكرونة 450 جم: الكمية 50 والمتوسط 2.6"
    Expect ProductValue("6281000000052", "CurrentQuantity") = 27 And ProductValue("6281000000052", "AverageCost") = CCur(22), "تمر سكري 1 كجم: الكمية 27 والمتوسط 22"
    Expect ProductValue("6281000000069", "CurrentQuantity") = 17 And ProductValue("6281000000069", "AverageCost") = CCur(11), "مياه معدنية 330 مل × 40: الكمية 17 والمتوسط 11"
    Expect ProductValue("6281000000076", "CurrentQuantity") = 64 And ProductValue("6281000000076", "AverageCost") = CCur(4.2), "عصير برتقال 1 لتر: الكمية 64 والمتوسط 4.2"
    Expect ProductValue("6281000000083", "CurrentQuantity") = 15 And ProductValue("6281000000083", "AverageCost") = CCur(7.8), "حليب طازج 2 لتر: الكمية 15 والمتوسط 7.8"
    Expect ProductValue("6281000000090", "CurrentQuantity") = 22 And ProductValue("6281000000090", "AverageCost") = CCur(26), "قهوة عربية 500 جم: الكمية 22 والمتوسط 26"
    Expect ProductValue("6281000000106", "CurrentQuantity") = 38 And ProductValue("6281000000106", "AverageCost") = CCur(9), "شاي أكياس × 100: الكمية 38 والمتوسط 9"
    Expect ProductValue("6281000000113", "CurrentQuantity") = 26 And ProductValue("6281000000113", "AverageCost") = CCur(11.5), "منظف أرضيات 3 لتر: الكمية 26 والمتوسط 11.5"
    Expect ProductValue("6281000000120", "CurrentQuantity") = 36 And ProductValue("6281000000120", "AverageCost") = CCur(5.3), "سائل غسيل صحون 1 لتر: الكمية 36 والمتوسط 5.3"
    Expect ProductValue("6281000000137", "CurrentQuantity") = 6 And ProductValue("6281000000137", "AverageCost") = CCur(24), "مسحوق غسيل 3 كجم: الكمية 6 والمتوسط 24"
    Expect ProductValue("6281000000144", "CurrentQuantity") = 55 And ProductValue("6281000000144", "AverageCost") = CCur(8.4), "مناديل ورقية × 5: الكمية 55 والمتوسط 8.4"
    Expect ProductValue("6281000000151", "CurrentQuantity") = 29 And ProductValue("6281000000151", "AverageCost") = CCur(12), "شامبو 400 مل: الكمية 29 والمتوسط 12"
    Expect ProductValue("6281000000168", "CurrentQuantity") = 10 And ProductValue("6281000000168", "AverageCost") = CCur(6.1), "معجون أسنان 100 مل: الكمية 10 والمتوسط 6.1"
    Expect ProductValue("6281000000175", "CurrentQuantity") = 38 And ProductValue("6281000000175", "AverageCost") = CCur(5.2), "صابون يدين 500 مل: الكمية 38 والمتوسط 5.2"
    Expect ProductValue("6281000000182", "CurrentQuantity") = 38 And ProductValue("6281000000182", "AverageCost") = CCur(6.7913), "أكياس نفايات كبيرة × 30: الكمية 38 والمتوسط 6.7913"
    Expect ProductValue("6281000000199", "CurrentQuantity") = 23 And ProductValue("6281000000199", "AverageCost") = CCur(8.9252), "ورق ألمنيوم 30 م: الكمية 23 والمتوسط 8.9252"
    Expect ProductValue("6281000000205", "CurrentQuantity") = 20 And ProductValue("6281000000205", "AverageCost") = CCur(7.761), "بطاريات AA × 4: الكمية 20 والمتوسط 7.761"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("مطعم الديرة")), -1) = CCur(422.20), "رصيد مطعم الديرة = 422.20"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("مؤسسة الضيافة للتموين")), -1) = CCur(51.50), "رصيد مؤسسة الضيافة للتموين = 51.50"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("مقهى الركن")), -1) = CCur(0.00), "رصيد مقهى الركن = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("أحمد محمد العتيبي")), -1) = CCur(58.00), "رصيد أحمد محمد العتيبي = 58.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("سارة عبدالله القحطاني")), -1) = CCur(0.00), "رصيد سارة عبدالله القحطاني = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("خالد إبراهيم الشهري")), -1) = CCur(0.00), "رصيد خالد إبراهيم الشهري = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("نورة سعد الدوسري")), -1) = CCur(0.00), "رصيد نورة سعد الدوسري = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("فهد عبدالرحمن الحربي")), -1) = CCur(127.00), "رصيد فهد عبدالرحمن الحربي = 127.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("منى علي الزهراني")), -1) = CCur(0.00), "رصيد منى علي الزهراني = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerName = " & SqlText("عبدالله ناصر المطيري")), -1) = CCur(0.00), "رصيد عبدالله ناصر المطيري = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierName = " & SqlText("مؤسسة الوفرة للمواد الغذائية")), -1) = CCur(1607.80), "رصيد مؤسسة الوفرة للمواد الغذائية = 1607.80"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierName = " & SqlText("شركة الينابيع للمشروبات")), -1) = CCur(0.00), "رصيد شركة الينابيع للمشروبات = 0.00"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierName = " & SqlText("مؤسسة النقاء للمنظفات")), -1) = CCur(601.27), "رصيد مؤسسة النقاء للمنظفات = 601.27"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierName = " & SqlText("شركة العناية الذهبية")), -1) = CCur(258.43), "رصيد شركة العناية الذهبية = 258.43"
    Expect Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierName = " & SqlText("مؤسسة البيت العصري للأدوات")), -1) = CCur(0.00), "رصيد مؤسسة البيت العصري للأدوات = 0.00"
    Expect CashBoxBalance(BoxOfType("MAIN")) = CCur(17705.40), "رصيد الخزينة الرئيسية = 17705.40"
    Expect CashBoxBalance(BoxOfType("CASHIER")) = CCur(194.30), "رصيد صندوق الكاشير = 194.30"
    Expect DbValue("SELECT COUNT(*) FROM LowStockQuery AS l INNER JOIN Products AS p ON l.ProductID = p.ProductID " & _
                   "WHERE p.Notes = '" & DEMO_MARK & "'") = 4, "4 منتجات منخفضة المخزون (للتنبيه والتقرير)"
    If Nz(SettingValue("SlowMovingDays"), 90) = 90 Then
        Expect DbValue("SELECT COUNT(*) FROM SlowMovingProductsQuery AS s INNER JOIN Products AS p ON s.ProductID = " & _
                       "p.ProductID WHERE p.Notes = '" & DEMO_MARK & "'") = 1, "1 منتج غير متحرك"
    End If
    Expect DbValue("SELECT COUNT(*) FROM StockCounts WHERE Status = 'POSTED'") = 1, "جرد مُرحّل واحد"
    Expect DbValue("SELECT COUNT(*) FROM IntegrityCheckQuery") = 0, "فحص سلامة البيانات: لا توجد أي مشكلة"
    n = Nz(DbValue("SELECT COUNT(*) FROM JournalEntries"), 0)
    Expect n = 34, "34 قيد يومية (الفعلي: " & n & ")"
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
    Expect DbValue("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit") = 0, "كل القيود متوازنة"
    Expect AccountBalance(1300) = Nz(DbValue("SELECT Sum(CurrentBalance) FROM Customers"), 0), "حساب ذمم العملاء = أرصدة العملاء"
    Expect Round(AccountBalance(1400), 2) = CCur(5799.27), "حساب المخزون في القيود = 5799.27 (الفعلي: " & _
           Format$(AccountBalance(1400), "0.00##") & ")"
    Debug.Print "--- نجح: " & m_passed & " | فشل: " & m_failed
    If m_failed = 0 Then
        TestMsg "البيانات التجريبية مطابقة تمامًا للنتائج المحسوبة مسبقًا (" & m_passed & " فحصًا)." & vbCrLf & _
                "المستخدمون التجريبيون: manager و cashier1 و cashier2، وكلمة المرور المؤقتة: " & DEMO_PASSWORD, _
                vbInformation + MSG_RTL, "VerifyDemoData"
        VerifyDemoData = True
    Else
        TestMsg "نجح " & m_passed & " وفشل " & m_failed & ":" & vbCrLf & vbCrLf & Left$(m_report, 900), _
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
    Expect n = Entries And Round(v, 2) = Stock, "قيود " & Kind & ": " & Entries & " قيد، المخزون " & _
           Format$(Stock, "0.00") & " (الفعلي: " & n & " قيد، " & Format$(v, "0.00##") & ")"
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
        MsgBox "حذف البيانات التجريبية لمدير النظام فقط.", vbExclamation + MSG_RTL, "RemoveDemoData"
        Exit Function
    End If
    loaded = DMax("LogDate", "AuditLog", "ActionType = 'DEMO_LOADED'")
    If IsNull(loaded) Then
        MsgBox "لم تُدخل بيانات تجريبية في هذه القاعدة.", vbInformation + MSG_RTL, "RemoveDemoData"
        Exit Function
    End If
    For Each t In Array("SalesInvoices", "SalesReturns", "PurchaseInvoices", "PurchaseReturns", "CustomerPayments", _
                        "SupplierPayments", "Expenses", "CashVouchers", "CashClosings")
        later = later + Nz(DbValue("SELECT COUNT(*) FROM [" & t & "] WHERE CreatedAt > " & SqlDate(loaded)), 0)
    Next
    later = later + Nz(DbValue("SELECT COUNT(*) FROM StockCounts WHERE CountDate > " & SqlDate(loaded)), 0)
    If later > 0 Then
        MsgBox "توجد " & later & " مستندات أُدخلت بعد البيانات التجريبية، فلن تُحذف أي بيانات." & vbCrLf & _
               "(الحذف هنا لا يميز بين التجريبي والحقيقي، لذلك يُرفض لحماية بياناتك.)", vbExclamation + MSG_RTL, _
               "RemoveDemoData"
        Exit Function
    End If
    If MsgBox("سيتم حذف كل المستندات والبيانات التجريبية وإعادة ترقيم المستندات من 1." & vbCrLf & _
              "تُحفظ نسخة احتياطية أولًا. هل تريد المتابعة؟", vbQuestion + vbYesNo + vbDefaultButton2 + MSG_RTL, _
              "RemoveDemoData") <> vbYes Then Exit Function
    msg = BackupNow(file, , "_BeforeDemoRemoval")
    If Len(msg) > 0 Then
        MsgBox "تعذرت النسخة الاحتياطية، فلم يُحذف شيء:" & vbCrLf & msg, vbExclamation + MSG_RTL, "RemoveDemoData"
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
    MsgBox "تم حذف البيانات التجريبية. القاعدة جاهزة لبياناتك الفعلية." & vbCrLf & _
           "النسخة قبل الحذف: " & file, vbInformation + MSG_RTL, "RemoveDemoData"
    RemoveDemoData = True
    Exit Function
EH:
    errText = Err.Description
    On Error Resume Next
    If inTrans Then ws.Rollback
    MsgBox "تعذر الحذف ولم يتغير شيء: " & errText, vbCritical + MSG_RTL, "RemoveDemoData"
End Function
