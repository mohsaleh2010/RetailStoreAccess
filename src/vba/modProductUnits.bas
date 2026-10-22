Attribute VB_Name = "modProductUnits"
'==============================================================================
' modProductUnits  -  Retail Store Management System
'
' Multiple units per product (docs/72, module 1 of docs/69). Products.UnitID is the smallest unit (a piece):
' the stock, the average cost and every stock move stay in it. ProductUnits adds bigger units of a product
' (a carton = 24 pieces): their factor, their own barcode and their own selling / purchase price.
' A sales or purchase line (and its return) is kept in the unit it was sold or bought in:
'     Quantity, UnitPrice / UnitCost  in that unit           UnitID  (Null = the smallest unit)
'     UnitFactor                      pieces in that unit    stock move = Quantity x UnitFactor
' so the amounts, the VAT and the e-invoice lines are those of the document, and the cost of a sale line
' (UnitCost = AverageCost x UnitFactor) keeps Quantity x UnitCost right in every report.
'   UnitFactorOf        1 for the smallest unit, the factor of a unit of the product, 0 = not one of its units
'   UnitSalePrice       the price of a new sale line in a unit (the unit's own price, else the piece price x factor;
'                       a customer's price list first, modPriceLists)
'   UnitPurchasePrice   the same for a purchase line (without VAT)
'   FindProductCode     a scanned code: the product's barcode / code (smallest unit) or a unit barcode
'   ValidateProductUnit the screen frmProductUnits (modForms.FormBeforeUpdate)
'==============================================================================
Option Compare Database
Option Explicit

'------------------------------------------------------------------------------
' Units of a product
'------------------------------------------------------------------------------
Public Function BaseUnitOf(ByVal ProductID As Long) As Long
    BaseUnitOf = Nz(DbValue("SELECT UnitID FROM Products WHERE ProductID = " & ProductID), 0)
End Function

Public Function IsBaseUnit(ByVal ProductID As Long, ByVal UnitID As Variant) As Boolean
    ' No unit, or the product's own (smallest) unit.
    If IsNull(UnitID) Or IsEmpty(UnitID) Then
        IsBaseUnit = True
    ElseIf CLng(UnitID) = 0 Then
        IsBaseUnit = True
    Else
        IsBaseUnit = (CLng(UnitID) = BaseUnitOf(ProductID))
    End If
End Function

Public Function UnitFactorOf(ByVal ProductID As Long, ByVal UnitID As Variant) As Currency
    ' Pieces in one UnitID of the product: 1 for its smallest unit, 0 when the unit is not one of its units.
    If IsBaseUnit(ProductID, UnitID) Then
        UnitFactorOf = 1
    Else
        UnitFactorOf = Nz(DbValue("SELECT UnitFactor FROM ProductUnits WHERE IsActive <> 0 AND ProductID = " & _
                                  ProductID & " AND UnitID = " & CLng(UnitID)), 0)
    End If
End Function

Public Function LineUnitID(ByVal ProductID As Long, ByVal UnitID As Variant) As Variant
    ' What a line keeps: Null for the smallest unit, else the unit.
    If IsBaseUnit(ProductID, UnitID) Then LineUnitID = Null Else LineUnitID = CLng(UnitID)
End Function

Private Function OwnUnitPrice(ByVal ProductID As Long, ByVal UnitID As Variant, ByVal FieldName As String) As Currency
    If IsBaseUnit(ProductID, UnitID) Then Exit Function
    OwnUnitPrice = Nz(DbValue("SELECT " & FieldName & " FROM ProductUnits WHERE IsActive <> 0 AND ProductID = " & _
                              ProductID & " AND UnitID = " & CLng(UnitID)), 0)
End Function

Public Function UnitSalePrice(ByVal ProductID As Long, ByVal UnitID As Variant, Optional ByVal CustomerID As Long = 0) _
                              As Currency
    ' A new sale line: the piece price (the rep's price list, modRepRoutes) or, for a bigger unit, its own price,
    ' else the piece price x its factor; for a customer with a price list, the price of his list (modPriceLists).
    Dim piece As Currency, own As Currency, usual As Currency
    piece = SessionPrice(ProductID, Nz(DbValue("SELECT SellingPrice FROM Products WHERE ProductID = " & ProductID), 0))
    If IsBaseUnit(ProductID, UnitID) Then
        usual = piece
    Else
        own = OwnUnitPrice(ProductID, UnitID, "SellingPrice")
        If own > 0 Then
            usual = own
        Else
            usual = RoundMoney(CDec(piece) * CDec(UnitFactorOf(ProductID, UnitID)))
        End If
    End If
    UnitSalePrice = CustomerPrice(CustomerID, ProductID, UnitID, usual)
End Function

Public Function IsListedUnitPrice(ByVal ProductID As Long, ByVal UnitID As Variant, ByVal Price As Currency) As Boolean
    ' A bigger unit sold at its own price, or at the piece price (or the rep's) x its factor: not a price change.
    Dim factor As Currency, piece As Currency
    If IsBaseUnit(ProductID, UnitID) Then Exit Function          ' the piece keeps the checks of modPOS / modSales
    factor = UnitFactorOf(ProductID, UnitID)
    If factor <= 0 Then Exit Function
    If Price = OwnUnitPrice(ProductID, UnitID, "SellingPrice") Then
        IsListedUnitPrice = True
        Exit Function
    End If
    piece = Nz(DbValue("SELECT SellingPrice FROM Products WHERE ProductID = " & ProductID), 0)
    IsListedUnitPrice = (Price = RoundMoney(CDec(piece) * CDec(factor))) Or _
                        (Price = RoundMoney(CDec(SessionPrice(ProductID, piece)) * CDec(factor)))
End Function

Public Function UnitPurchasePrice(ByVal ProductID As Long, ByVal UnitID As Variant) As Currency
    ' A new purchase line (without VAT): the unit's own purchase price, else the last piece price x its factor.
    Dim own As Currency
    own = OwnUnitPrice(ProductID, UnitID, "PurchasePrice")
    If own > 0 Then
        UnitPurchasePrice = own
    Else
        UnitPurchasePrice = RoundMoney(CDec(Nz(DbValue("SELECT PurchasePrice FROM Products WHERE ProductID = " & _
                                                       ProductID), 0)) * CDec(UnitFactorOf(ProductID, UnitID)), 4)
    End If
End Function

Public Function UnitNameOf(ByVal ProductID As Long, ByVal UnitID As Variant) As String
    ' The name of the unit of a line, in the interface language.
    Dim u As Long
    If IsBaseUnit(ProductID, UnitID) Then u = BaseUnitOf(ProductID) Else u = CLng(UnitID)
    UnitNameOf = Nz(DbValue(Tr("SELECT u.UnitName FROM [@Units] AS u WHERE u.UnitID = " & u)), "")
End Function

Public Function FindProductCode(ByVal Code As String, ByRef ProductID As Long, ByRef UnitID As Variant) As Boolean
    ' A scanned or typed code: the barcode or code of a product (its smallest unit), else the barcode of one of
    ' its units. False when nothing active has it.
    Dim v As Variant, rs As DAO.Recordset
    ProductID = 0
    UnitID = Null
    v = DLookup("ProductID", "Products", "IsActive <> 0 AND (Barcode = " & SqlText(Code) & _
                " OR ProductCode = " & SqlText(Code) & ")")
    If Not IsNull(v) Then
        ProductID = CLng(v)
        FindProductCode = True
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT u.ProductID, u.UnitID FROM ProductUnits AS u INNER JOIN Products AS p " & _
                                     "ON u.ProductID = p.ProductID WHERE u.IsActive <> 0 AND p.IsActive <> 0 AND " & _
                                     "u.Barcode = " & SqlText(Code), dbOpenSnapshot)
    If Not rs.EOF Then
        ProductID = rs!ProductID
        UnitID = rs!UnitID
        FindProductCode = True
    End If
    rs.Close
End Function

'------------------------------------------------------------------------------
' Screens
'------------------------------------------------------------------------------
Public Function ProductBarcodeProblem(ByVal Barcode As Variant, ByVal ProductID As Variant) As String
    ' frmProducts: a product barcode may not be the barcode of a unit (the scan would find two things).
    If Len(Trim$(Nz(Barcode, ""))) = 0 Then Exit Function
    If Not IsNull(DbValue("SELECT ProductUnitID FROM ProductUnits WHERE Barcode = " & SqlText(Trim$(Barcode)))) Then
        ProductBarcodeProblem = "الباركود " & Trim$(Barcode) & " مستخدم لوحدة من وحدات الأصناف."
    End If
End Function

Public Function ProductUnitProblem(ByVal ProductID As Variant, ByVal UnitID As Variant, ByVal UnitFactor As Variant, _
                                   ByVal Barcode As Variant) As String
    ' "" when the unit can be saved.
    Dim code As String
    If IsNull(ProductID) Or IsNull(UnitID) Then
        ProductUnitProblem = "اختر الصنف والوحدة."
    ElseIf IsBaseUnit(CLng(ProductID), UnitID) Then
        ProductUnitProblem = "هذه هي الوحدة الصغرى للصنف نفسه (من شاشة المنتجات): اختر وحدة أكبر مثل الكرتونة."
    ElseIf Nz(UnitFactor, 0) <= 0 Or Nz(UnitFactor, 0) = 1 Then
        ProductUnitProblem = "عدد الوحدات الصغرى في الوحدة يجب أن يكون أكبر من صفر وغير 1."
    Else
        code = Trim$(Nz(Barcode, ""))
        If Len(code) > 0 Then
            If Not IsNull(DbValue("SELECT ProductID FROM Products WHERE Barcode = " & SqlText(code) & _
                                  " OR ProductCode = " & SqlText(code))) Then
                ProductUnitProblem = "الباركود " & code & " مستخدم باركودًا أو كودًا لصنف."
            End If
        End If
    End If
End Function

Public Function ValidateProductUnit(ByVal frm As Access.Form) As Boolean
    Dim msg As String
    msg = ProductUnitProblem(frm!ProductID.Value, frm!UnitID.Value, frm!UnitFactor.Value, frm!Barcode.Value)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    If Len(Trim$(Nz(frm!Barcode.Value, ""))) = 0 Then frm!Barcode.Value = Null
    ValidateProductUnit = True
End Function

Public Sub OpenProductUnits(ByVal frm As Access.Form)
    ' From frmProducts: the units of this product; a new unit starts on it.
    Dim pid As Long
    If frm.Dirty Then frm.Dirty = False
    pid = Nz(frm!ProductID.Value, 0)
    If pid = 0 Then
        ShowWarning "احفظ الصنف أولًا."
        Exit Sub
    End If
    OpenScreen "frmProductUnits", 0, DbValue("SELECT Min(ProductUnitID) FROM ProductUnits WHERE ProductID = " & pid)
    If IsFormOpen("frmProductUnits") Then Forms("frmProductUnits").Controls("ProductID").DefaultValue = CStr(pid)
End Sub

'------------------------------------------------------------------------------
' In-Access test (Immediate window: TestProductUnits). Works inside a transaction that is rolled back.
'------------------------------------------------------------------------------
Private Sub CheckPu(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
                    ByRef Report As String)
    If ok Then
        passed = passed + 1
        Debug.Print "  OK   " & Title
    Else
        failed = failed + 1
        Debug.Print "  FAIL " & Title
        Report = Report & "- " & Title & vbCrLf
    End If
End Sub

Private Function StockOf(ByVal ProductID As Long) As Currency
    StockOf = Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & ProductID), 0)
End Function

Public Function TestProductUnits() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, pid As Long, piece As Long, carton As Long, uid As Variant
    Dim passed As Long, failed As Long, Report As String, msg As String, sup As Long, cust As Long
    Dim invID As Long, saleID As Long, retID As Long, detail As Long, found As Long
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestProductUnits  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    EnsureLocalTables
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError + dbSeeChanges
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError + dbSeeChanges
    piece = Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1)
    CurrentDb.Execute "INSERT INTO Units (UnitName, IsActive) VALUES ('TEST كرتونة', True)", dbFailOnError + dbSeeChanges
    carton = DbValue("SELECT UnitID FROM Units WHERE UnitName = 'TEST كرتونة'")
    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, SellingPrice, " & _
                      "PurchasePrice, CurrentQuantity, AverageCost) VALUES ('TEST-PU-P', 'TEST-PU-1', 'TEST صنف الوحدات', " & _
                      Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1) & ", " & piece & ", 5, 2, 0, 0)", _
                      dbFailOnError + dbSeeChanges
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-PU-P'")
    sup = Nz(DbValue("SELECT Min(SupplierID) FROM Suppliers WHERE IsActive <> 0"), 0)
    cust = Nz(DbValue("SELECT Min(CustomerID) FROM Customers WHERE IsActive <> 0"), 0)
    CheckPu sup > 0 And cust > 0, "يوجد مورد وعميل نشطان", passed, failed, Report

    ' the unit and its checks
    CheckPu Len(ProductUnitProblem(pid, piece, 12, Null)) > 0, "الوحدة الصغرى لا تُضاف وحدةً", passed, failed, Report
    CheckPu Len(ProductUnitProblem(pid, carton, 1, Null)) > 0, "العدد 1 مرفوض", passed, failed, Report
    CheckPu Len(ProductUnitProblem(pid, carton, 12, "TEST-PU-1")) > 0, "باركود صنف لا يكون باركود وحدة", passed, failed, Report
    CheckPu Len(ProductUnitProblem(pid, carton, 12, "TEST-PU-12")) = 0, "وحدة صحيحة", passed, failed, Report
    CurrentDb.Execute "INSERT INTO ProductUnits (ProductID, UnitID, UnitFactor, Barcode, SellingPrice, PurchasePrice) " & _
                      "VALUES (" & pid & ", " & carton & ", 12, 'TEST-PU-12', 0, 0)", dbFailOnError + dbSeeChanges
    CheckPu Len(ProductBarcodeProblem("TEST-PU-12", pid)) > 0, "باركود وحدة لا يكون باركود صنف", passed, failed, Report
    CheckPu UnitFactorOf(pid, Null) = 1 And UnitFactorOf(pid, piece) = 1 And UnitFactorOf(pid, carton) = 12, _
            "عدد الحبات في الوحدة", passed, failed, Report
    CheckPu UnitSalePrice(pid, carton) = 60 And UnitPurchasePrice(pid, carton) = 24, _
            "سعر الكرتونة = سعر الحبة × 12 حين لا سعر لها", passed, failed, Report
    CurrentDb.Execute "UPDATE ProductUnits SET SellingPrice = 55, PurchasePrice = 22 WHERE ProductID = " & pid, _
                      dbFailOnError + dbSeeChanges
    CheckPu UnitSalePrice(pid, carton) = 55 And UnitPurchasePrice(pid, carton) = 22 And IsListedUnitPrice(pid, carton, 55) _
            And IsListedUnitPrice(pid, carton, 60) And Not IsListedUnitPrice(pid, carton, 50), _
            "سعر الكرتونة الخاص", passed, failed, Report
    CheckPu FindProductCode("TEST-PU-12", found, uid) And found = pid And Nz(uid, 0) = carton, _
            "مسح باركود الكرتونة", passed, failed, Report
    CheckPu FindProductCode("TEST-PU-1", found, uid) And found = pid And IsNull(uid), "مسح باركود الحبة", _
            passed, failed, Report

    ' buy 2 cartons at 24 each: 24 pieces at 2 each
    CurrentDb.Execute "INSERT INTO tmpPurchaseLines (ProductID, ProductCode, ProductName, Quantity, UnitCost, LineDiscount, " & _
                      "LineTotal, SellingPrice, UnitID, UnitFactor) VALUES (" & pid & ", 'TEST-PU-P', 'x', 2, 24, 0, 0, 55, " & _
                      carton & ", 12)", dbFailOnError + dbSeeChanges
    msg = PostPurchaseFromCart(sup, "TEST-PU-INV", Date, "CREDIT", Null, False, 0, Null, "TEST", invID)
    CheckPu Len(msg) = 0 And StockOf(pid) = 24 And Nz(DbValue("SELECT AverageCost FROM Products WHERE ProductID = " & pid), 0) = 2, _
            "شراء كرتونتين = 24 حبة بتكلفة 2 للحبة " & msg, passed, failed, Report
    CheckPu Nz(DbValue("SELECT Quantity * 100 + UnitFactor FROM PurchaseInvoiceDetails WHERE PurchaseInvoiceID = " & invID), 0) = 212 _
            And Nz(DbValue("SELECT UnitID FROM PurchaseInvoiceDetails WHERE PurchaseInvoiceID = " & invID), 0) = carton, _
            "سطر الشراء بالكرتونة: الكمية 2 والعدد 12", passed, failed, Report

    ' sell 1 carton: 12 pieces leave, the line keeps the carton and its cost
    CurrentDb.Execute "INSERT INTO tmpPOSLines (ProductID, ProductCode, ProductName, Quantity, UnitPrice, LineDiscount, " & _
                      "LineTotal, Available, UnitID, UnitFactor) VALUES (" & pid & ", 'TEST-PU-P', 'x', 1, 55, 0, 0, 0, " & _
                      carton & ", 12)", dbFailOnError + dbSeeChanges
    msg = PostSaleFromCart(cust, "CASH", 1, 0, Null, "TEST", saleID)
    CheckPu Len(msg) = 0 And StockOf(pid) = 12, "بيع كرتونة يخصم 12 حبة " & msg, passed, failed, Report
    detail = Nz(DbValue("SELECT SalesDetailID FROM SalesInvoiceDetails WHERE SalesInvoiceID = " & saleID), 0)
    CheckPu Nz(DbValue("SELECT UnitCost FROM SalesInvoiceDetails WHERE SalesDetailID = " & detail), 0) = 24 And _
            Nz(DbValue("SELECT LineTotal FROM SalesInvoiceDetails WHERE SalesDetailID = " & detail), 0) = 55, _
            "سطر البيع: سعر الكرتونة وتكلفتها", passed, failed, Report
    CheckPu Nz(DbValue("SELECT Sum(SignedQty) FROM qrySalesLineItems WHERE DocType = 'SALE' AND DocID = " & saleID), 0) = 12 _
            And Nz(DbValue("SELECT Sum(LineCost) FROM qrySalesLineItems WHERE DocType = 'SALE' AND DocID = " & saleID), 0) = 24, _
            "تقارير الأصناف بالحبة والتكلفة صحيحة", passed, failed, Report
    CheckPu Nz(DbValue("SELECT Quantity FROM InventoryTransactions WHERE ReferenceType = 'SALE' AND ReferenceID = " & saleID), 0) = -12 _
            And Nz(DbValue("SELECT UnitCost FROM InventoryTransactions WHERE ReferenceType = 'SALE' AND ReferenceID = " & saleID), 0) = 2, _
            "حركة المخزون بالحبة وتكلفتها", passed, failed, Report

    ' the carton comes back
    CurrentDb.Execute "DELETE FROM tmpReturnLines", dbFailOnError + dbSeeChanges
    CurrentDb.Execute "INSERT INTO tmpReturnLines (SalesDetailID, ProductID, ProductName, SoldQty, ReturnedQty, AvailableQty, " & _
                      "ReturnQty, ReturnToStock, ReturnAmount) VALUES (" & detail & ", " & pid & ", 'x', 1, 0, 1, 1, True, 0)", _
                      dbFailOnError + dbSeeChanges
    msg = PostSalesReturn(saleID, "TEST", "CASH", 1, retID)
    CheckPu Len(msg) = 0 And StockOf(pid) = 24 And _
            Nz(DbValue("SELECT UnitFactor FROM SalesReturnDetails WHERE SalesReturnID = " & retID), 0) = 12, _
            "مرتجع الكرتونة يعيد 12 حبة " & msg, passed, failed, Report

    ' one carton back to the supplier
    CurrentDb.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError + dbSeeChanges
    CurrentDb.Execute "INSERT INTO tmpPurchaseReturnLines (PurchaseDetailID, ProductID, ProductName, BoughtQty, ReturnedQty, " & _
                      "AvailableQty, InStock, ReturnQty, ReturnAmount) SELECT PurchaseDetailID, ProductID, 'x', 2, 0, 2, 2, 1, 0 " & _
                      "FROM PurchaseInvoiceDetails WHERE PurchaseInvoiceID = " & invID, dbFailOnError + dbSeeChanges
    msg = PostPurchaseReturn(invID, "TEST", "CREDIT", Null, retID)
    CheckPu Len(msg) = 0 And StockOf(pid) = 12 And _
            Nz(DbValue("SELECT UnitFactor FROM PurchaseReturnDetails WHERE PurchaseReturnID = " & retID), 0) = 12, _
            "مرتجع شراء كرتونة يخصم 12 حبة " & msg, passed, failed, Report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckPu False, "خطأ: " & Err.Description, passed, failed, Report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات وحدات الأصناف ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestProductUnits"
        TestProductUnits = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & Report, vbExclamation + MSG_RTL, _
                "TestProductUnits"
    End If
End Function
