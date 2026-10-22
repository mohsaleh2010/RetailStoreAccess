Attribute VB_Name = "modPriceLists"
'==============================================================================
' modPriceLists  -  Retail Store Management System
'
' Price lists (docs/73, module 2 of docs/69): retail / wholesale / distributor. A customer has one list
' (Customers.PriceListID, Null = the usual selling price); the list is his price group.
'   PriceListItems      the price of a product in a unit of it (UnitID Null = its smallest unit), VAT included
'                       like the selling price, valid ValidFrom..ValidTo (Null = no limit); among the items
'                       valid on a day the one that starts last wins (a promotion over a standing price)
'   DiscountPercent     a product without an item is sold at its usual price less this rate (0 = none)
' The price of a new sale line for a customer (modProductUnits.UnitSalePrice):
'   1. the item of its unit in the customer's list
'   2. a bigger unit without an item: the item of the smallest unit x the factor
'   3. else the usual price (the rep's price list, the unit's own price, the piece price x factor) less the
'      discount of the list
' A line at that price is not a price change (IsCustomerListPrice: modSales, modPOS). The POS reprices the
' lines still at the price it gave them when the customer changes (modPOS.RepriceCart); an order line and a
' delivery note line start at it (modSalesOrders, modStockNotes).
'==============================================================================
Option Compare Database
Option Explicit

'------------------------------------------------------------------------------
' The list of a customer and its prices
'------------------------------------------------------------------------------
Public Function CustomerPriceListID(ByVal CustomerID As Long) As Long
    ' The active price list of the customer, else 0.
    If CustomerID = 0 Then Exit Function
    CustomerPriceListID = Nz(DbValue("SELECT l.PriceListID FROM Customers AS c INNER JOIN PriceLists AS l " & _
                                     "ON c.PriceListID = l.PriceListID WHERE l.IsActive <> 0 AND c.CustomerID = " & _
                                     CustomerID), 0)
End Function

Public Function PriceListItemPrice(ByVal PriceListID As Long, ByVal ProductID As Long, ByVal UnitID As Variant, _
                                   ByVal OnDate As Date) As Variant
    ' The price of the product in the unit on a day: the item valid that day that starts last, else Null.
    Dim unitCond As String
    PriceListItemPrice = Null
    If PriceListID = 0 Then Exit Function
    If IsBaseUnit(ProductID, UnitID) Then
        unitCond = "UnitID Is Null"
    Else
        unitCond = "UnitID = " & CLng(UnitID)
    End If
    PriceListItemPrice = DbValue("SELECT TOP 1 UnitPrice FROM PriceListItems WHERE PriceListID = " & PriceListID & _
                                 " AND ProductID = " & ProductID & " AND " & unitCond & _
                                 " AND (ValidFrom Is Null OR ValidFrom <= " & SqlDate(DateValue(OnDate)) & ")" & _
                                 " AND (ValidTo Is Null OR ValidTo >= " & SqlDate(DateValue(OnDate)) & ")" & _
                                 " ORDER BY ValidFrom DESC, PriceListItemID DESC")
End Function

Public Function CustomerListPrice(ByVal CustomerID As Long, ByVal ProductID As Long, ByVal UnitID As Variant) As Variant
    ' The price of the customer's list today: the item of the unit, else (a bigger unit) the item of the smallest
    ' unit x the factor. Null when the list has no price for the product.
    Dim listID As Long, v As Variant, factor As Currency
    CustomerListPrice = Null
    listID = CustomerPriceListID(CustomerID)
    If listID = 0 Then Exit Function
    v = PriceListItemPrice(listID, ProductID, UnitID, Date)
    If Not IsNull(v) Then
        CustomerListPrice = CCur(v)
        Exit Function
    End If
    If IsBaseUnit(ProductID, UnitID) Then Exit Function
    factor = UnitFactorOf(ProductID, UnitID)
    If factor <= 0 Then Exit Function
    v = PriceListItemPrice(listID, ProductID, Null, Date)
    If Not IsNull(v) Then CustomerListPrice = RoundMoney(CDec(v) * CDec(factor))
End Function

Public Function ListDiscounted(ByVal CustomerID As Long, ByVal Price As Currency) As Currency
    ' A price without an item in the customer's list, less the discount of the list.
    Dim pct As Double, listID As Long
    ListDiscounted = Price
    listID = CustomerPriceListID(CustomerID)
    If listID = 0 Then Exit Function
    pct = Nz(DbValue("SELECT DiscountPercent FROM PriceLists WHERE PriceListID = " & listID), 0)
    If pct > 0 Then ListDiscounted = RoundMoney(CDec(Price) * (1 - CDec(pct)))
End Function

Public Function CustomerPrice(ByVal CustomerID As Long, ByVal ProductID As Long, ByVal UnitID As Variant, _
                              ByVal UsualPrice As Currency) As Currency
    ' The price for the customer: his list, else the usual price less the discount of his list.
    Dim v As Variant
    v = CustomerListPrice(CustomerID, ProductID, UnitID)
    If IsNull(v) Then
        CustomerPrice = ListDiscounted(CustomerID, UsualPrice)
    Else
        CustomerPrice = CCur(v)
    End If
End Function

Public Function IsCustomerListPrice(ByVal CustomerID As Long, ByVal ProductID As Long, ByVal UnitID As Variant, _
                                    ByVal Price As Currency) As Boolean
    ' A line of a customer with a price list sold at the price the list gives: not a price change.
    If CustomerPriceListID(CustomerID) = 0 Then Exit Function
    IsCustomerListPrice = (Price = UnitSalePrice(ProductID, UnitID, CustomerID))
End Function

'------------------------------------------------------------------------------
' The screens
'------------------------------------------------------------------------------
Public Function PriceListItemProblem(ByVal PriceListID As Variant, ByVal ProductID As Variant, ByVal UnitID As Variant, _
                                     ByVal UnitPrice As Variant, ByVal ValidFrom As Variant, ByVal ValidTo As Variant, _
                                     ByVal PriceListItemID As Variant) As String
    ' "" when the price can be saved.
    Dim sameStart As String, unitCond As String
    If IsNull(PriceListID) Or IsNull(ProductID) Then
        PriceListItemProblem = "«Œ — «·ﬁ«∆„… Ê«·’‰›."
        Exit Function
    End If
    If Nz(UnitPrice, 0) <= 0 Then
        PriceListItemProblem = "«·”⁄— ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—."
        Exit Function
    End If
    If UnitFactorOf(CLng(ProductID), UnitID) <= 0 Then
        PriceListItemProblem = "Â–Â «·ÊÕœ… ·Ì”  „‰ ÊÕœ«  «·’‰›. ⁄—¯›Â« „‰ ´ÊÕœ«  «·’‰›ª ›Ì ‘«‘… «·„‰ Ã« ."
        Exit Function
    End If
    If Not IsNull(ValidFrom) And Not IsNull(ValidTo) Then
        If CDate(ValidTo) < CDate(ValidFrom) Then
            PriceListItemProblem = " «—ÌŒ ´”«—Ì Õ Ïª ﬁ»· ´”«—Ì „‰ª."
            Exit Function
        End If
    End If
    If IsBaseUnit(CLng(ProductID), UnitID) Then
        unitCond = "UnitID Is Null"
    Else
        unitCond = "UnitID = " & CLng(UnitID)
    End If
    If IsNull(ValidFrom) Then
        sameStart = "ValidFrom Is Null"
    Else
        sameStart = "ValidFrom = " & SqlDate(DateValue(CDate(ValidFrom)))
    End If
    If Not IsNull(DbValue("SELECT PriceListItemID FROM PriceListItems WHERE PriceListID = " & CLng(PriceListID) & _
                          " AND ProductID = " & CLng(ProductID) & " AND " & unitCond & " AND " & sameStart & _
                          " AND PriceListItemID <> " & Nz(PriceListItemID, 0))) Then
        PriceListItemProblem = "··’‰› »Â–Â «·ÊÕœ… ”⁄— ›Ì «·ﬁ«∆„… Ì»œ√ ›Ì «· «—ÌŒ ‰›”Â. ⁄œ¯·Â »œ· ≈÷«›… ”⁄— ¬Œ—."
    End If
End Function

Public Function ValidatePriceListItem(ByVal frm As Access.Form) As Boolean
    Dim msg As String
    msg = PriceListItemProblem(frm!PriceListID.Value, frm!ProductID.Value, frm!UnitID.Value, frm!UnitPrice.Value, _
                               frm!ValidFrom.Value, frm!ValidTo.Value, frm!PriceListItemID.Value)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    frm!UnitID.Value = LineUnitID(CLng(frm!ProductID.Value), frm!UnitID.Value)   ' the smallest unit is kept as Null
    ValidatePriceListItem = True
End Function

Public Sub OpenPriceListItems(ByVal frm As Access.Form)
    ' From frmPriceLists: the prices of this list; a new price starts on it.
    Dim listID As Long
    If frm.Dirty Then frm.Dirty = False
    listID = Nz(frm!PriceListID.Value, 0)
    If listID = 0 Then
        ShowWarning "«Õ›Ÿ «·ﬁ«∆„… √Ê·«."
        Exit Sub
    End If
    OpenScreen "frmPriceListItems", 0, DbValue("SELECT Min(PriceListItemID) FROM PriceListItems WHERE PriceListID = " & listID)
    If IsFormOpen("frmPriceListItems") Then Forms("frmPriceListItems").Controls("PriceListID").DefaultValue = CStr(listID)
End Sub

'------------------------------------------------------------------------------
' In-Access test (Immediate window: TestPriceLists). Works inside a transaction that is rolled back.
'------------------------------------------------------------------------------
Private Sub CheckPl(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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

Public Function TestPriceLists() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, pid As Long, piece As Long, carton As Long, listID As Long
    Dim passed As Long, failed As Long, Report As String, msg As String, cust As Long, saleID As Long
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestPriceLists  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    EnsureLocalTables
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError + dbSeeChanges
    piece = Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1)
    CurrentDb.Execute "INSERT INTO Units (UnitName, IsActive) VALUES ('TEST ﬂ— Ê‰… «·ﬁ«∆„…', True)", dbFailOnError + dbSeeChanges
    carton = DbValue("SELECT UnitID FROM Units WHERE UnitName = 'TEST ﬂ— Ê‰… «·ﬁ«∆„…'")
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, SellingPrice, PurchasePrice, " & _
                      "CurrentQuantity, AverageCost, TrackStock) VALUES ('TEST-PL-P', 'TEST ’‰› «·ﬁ«∆„…', " & _
                      Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1) & ", " & piece & ", 10, 5, 0, 5, False)", _
                      dbFailOnError + dbSeeChanges
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-PL-P'")
    CurrentDb.Execute "INSERT INTO ProductUnits (ProductID, UnitID, UnitFactor, SellingPrice, PurchasePrice) VALUES (" & _
                      pid & ", " & carton & ", 12, 0, 0)", dbFailOnError + dbSeeChanges
    CurrentDb.Execute "INSERT INTO PriceLists (ListName, DiscountPercent, IsActive) VALUES ('TEST Ã„·…', 0.1, True)", _
                      dbFailOnError + dbSeeChanges
    listID = DbValue("SELECT PriceListID FROM PriceLists WHERE ListName = 'TEST Ã„·…'")
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, AllowCredit, IsActive) VALUES ('TEST ⁄„Ì· «·Ã„·…', True, True)", _
                      dbFailOnError + dbSeeChanges
    cust = DbValue("SELECT CustomerID FROM Customers WHERE CustomerName = 'TEST ⁄„Ì· «·Ã„·…'")

    ' without a list: the usual price
    CheckPl UnitSalePrice(pid, Null, cust) = 10 And UnitSalePrice(pid, carton, cust) = 120, _
            "»·« ﬁ«∆„…: ”⁄— «·»Ì⁄ «·„⁄ «œ", passed, failed, Report
    CurrentDb.Execute "UPDATE Customers SET PriceListID = " & listID & " WHERE CustomerID = " & cust, dbFailOnError + dbSeeChanges
    CheckPl CustomerPriceListID(cust) = listID, "ﬁ«∆„… «·⁄„Ì·", passed, failed, Report
    CheckPl UnitSalePrice(pid, Null, cust) = 9 And UnitSalePrice(pid, carton, cust) = 108 And UnitSalePrice(pid, Null) = 10, _
            "’‰› »·« ”⁄— ›Ì «·ﬁ«∆„…: «·„⁄ «œ ‰«ﬁ’ Œ’„ «·ﬁ«∆„… 10%", passed, failed, Report

    ' the checks of a price
    CheckPl Len(PriceListItemProblem(listID, pid, Null, 0, Null, Null, Null)) > 0, "«·”⁄— ’›— „—›Ê÷", passed, failed, Report
    CheckPl Len(PriceListItemProblem(listID, pid, Null, 8, Date, Date - 1, Null)) > 0, "‰Â«Ì… ﬁ»· «·»œ«Ì… „—›Ê÷…", _
            passed, failed, Report
    CheckPl Len(PriceListItemProblem(listID, pid, Null, 8, Null, Null, Null)) = 0, "”⁄— ’ÕÌÕ", passed, failed, Report
    CurrentDb.Execute "INSERT INTO PriceListItems (PriceListID, ProductID, UnitPrice) VALUES (" & listID & ", " & pid & _
                      ", 8)", dbFailOnError + dbSeeChanges
    CheckPl Len(PriceListItemProblem(listID, pid, piece, 7, Null, Null, Null)) > 0, _
            "”⁄—«‰ ··ÊÕœ… ‰›”Â« Ì»œ¬‰ „⁄« „—›Ê÷«‰", passed, failed, Report
    CheckPl UnitSalePrice(pid, Null, cust) = 8 And UnitSalePrice(pid, carton, cust) = 96, _
            "”⁄— «·Õ»… ›Ì «·ﬁ«∆„…∫ «·ﬂ— Ê‰… = ”⁄—Â« ◊ 12", passed, failed, Report
    CurrentDb.Execute "INSERT INTO PriceListItems (PriceListID, ProductID, UnitID, UnitPrice) VALUES (" & listID & ", " & pid & _
                      ", " & carton & ", 90)", dbFailOnError + dbSeeChanges
    CheckPl UnitSalePrice(pid, carton, cust) = 90, "”⁄— «·ﬂ— Ê‰… ›Ì «·ﬁ«∆„…", passed, failed, Report

    ' dates: a promotion over the standing price, then an old and a future price
    CurrentDb.Execute "INSERT INTO PriceListItems (PriceListID, ProductID, UnitPrice, ValidFrom, ValidTo) VALUES (" & _
                      listID & ", " & pid & ", 7, " & SqlDate(Date - 2) & ", " & SqlDate(Date + 2) & ")", _
                      dbFailOnError + dbSeeChanges
    CurrentDb.Execute "INSERT INTO PriceListItems (PriceListID, ProductID, UnitPrice, ValidFrom, ValidTo) VALUES (" & _
                      listID & ", " & pid & ", 6, " & SqlDate(Date - 30) & ", " & SqlDate(Date - 10) & ")", _
                      dbFailOnError + dbSeeChanges
    CurrentDb.Execute "INSERT INTO PriceListItems (PriceListID, ProductID, UnitPrice, ValidFrom) VALUES (" & _
                      listID & ", " & pid & ", 5, " & SqlDate(Date + 5) & ")", dbFailOnError + dbSeeChanges
    CheckPl UnitSalePrice(pid, Null, cust) = 7, "«·⁄—÷ «·”«—Ì «·ÌÊ„ Ì€·» «·”⁄— «·œ«∆„", passed, failed, Report
    CheckPl PriceListItemPrice(listID, pid, Null, Date - 20) = 6 And PriceListItemPrice(listID, pid, Null, Date + 6) = 5 _
            And PriceListItemPrice(listID, pid, Null, Date + 3) = 8, "«·”⁄— ›Ì ﬂ· ÌÊ„ » «—ÌŒ ”—Ì«‰Â", passed, failed, Report
    CheckPl IsCustomerListPrice(cust, pid, Null, 7) And Not IsCustomerListPrice(cust, pid, Null, 8) _
            And Not IsCustomerListPrice(1, pid, Null, 7), "”⁄— «·ﬁ«∆„… ·Ì”  ⁄œÌ·« ··”⁄—", passed, failed, Report
    CurrentDb.Execute "UPDATE PriceLists SET IsActive = False WHERE PriceListID = " & listID, dbFailOnError + dbSeeChanges
    CheckPl UnitSalePrice(pid, Null, cust) = 10, "ﬁ«∆„… „ÊﬁÊ›… ·«  ı” ⁄„·", passed, failed, Report
    CurrentDb.Execute "UPDATE PriceLists SET IsActive = True WHERE PriceListID = " & listID, dbFailOnError + dbSeeChanges

    ' a sale at the prices of the list, without the permission to change prices
    CurrentDb.Execute "INSERT INTO tmpPOSLines (ProductID, ProductCode, ProductName, Quantity, UnitPrice, LineDiscount, " & _
                      "LineTotal, Available, UnitID, UnitFactor, ListPrice) VALUES (" & pid & ", 'TEST-PL-P', 'x', 2, 7, 0, 0, 0, " & _
                      "Null, 1, 7)", dbFailOnError + dbSeeChanges
    CurrentDb.Execute "INSERT INTO tmpPOSLines (ProductID, ProductCode, ProductName, Quantity, UnitPrice, LineDiscount, " & _
                      "LineTotal, Available, UnitID, UnitFactor, ListPrice) VALUES (" & pid & ", 'TEST-PL-P', 'x', 1, 90, 0, 0, 0, " & _
                      carton & ", 12, 90)", dbFailOnError + dbSeeChanges
    msg = PostSaleFromCart(cust, "CASH", 1, 0, Null, "TEST", saleID)
    CheckPl Len(msg) = 0 And Nz(DbValue("SELECT Sum(LineTotal) FROM SalesInvoiceDetails WHERE SalesInvoiceID = " & saleID), 0) = 104, _
            "»Ì⁄ »√”⁄«— «·ﬁ«∆„… (2 ◊ 7 + ﬂ— Ê‰… 90) " & msg, passed, failed, Report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckPl False, "Œÿ√: " & Err.Description, passed, failed, Report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  ﬁÊ«∆„ «·√”⁄«— ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestPriceLists"
        TestPriceLists = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & Report, vbExclamation + MSG_RTL, _
                "TestPriceLists"
    End If
End Function
