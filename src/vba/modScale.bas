Attribute VB_Name = "modScale"
'==============================================================================
' modScale  -  Retail Store Management System
'
' Scale barcodes (docs/86): the label a scale prints for weighed goods is an EAN-13 of 13 digits:
'     prefix (LabelSettings.ScalePrefixes, e.g. "20,21")  +  item code (ScaleItemDigits digits = the product code)
'     +  the weight in grams or the price in cents / piasters (ScaleValueKind WEIGHT / PRICE)  +  the check digit.
' The product is the active one whose code is the item code (as printed, else without its leading zeros).
' A price label gives the quantity: the price / the price of a new line for the customer, rounded to 3 decimals.
' The POS reads it when the code is not the barcode of a product or unit (modPOS.ScanCode); the Windows POS reads
' it the same way (RetailStore.Domain.Sales.ScaleBarcode). Python mirror: tools/scale_reference.py.
'   ScaleCheckDigit     the EAN-13 check digit of 12 digits
'   ScaleCodeParts      the item code and the value of a scale barcode of the settings
'   ScaleProductID      the product of an item code
'   ScaleLine           a scanned code -> the product and the quantity of the line (or why not)
'==============================================================================
Option Compare Database
Option Explicit

Public Function ScaleCheckDigit(ByVal Digits12 As String) As Long
    ' Weights 1, 3, 1, 3 ... from the left.
    Dim i As Long, total As Long
    For i = 1 To 12
        If i Mod 2 = 0 Then
            total = total + 3 * CLng(Mid$(Digits12, i, 1))
        Else
            total = total + CLng(Mid$(Digits12, i, 1))
        End If
    Next
    ScaleCheckDigit = (10 - total Mod 10) Mod 10
End Function

Private Function ScaleSetting(ByVal FieldName As String) As Variant
    ScaleSetting = DbValue("SELECT " & FieldName & " FROM LabelSettings WHERE LabelSettingID = 1")
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) Like "#") Then Exit Function
    Next
    IsDigits = True
End Function

Public Function ScaleCodeParts(ByVal Code As String, ByRef ItemCode As String, ByRef ScaleValue As Long) As Boolean
    ' True when Code is a scale barcode of the settings with a right check digit.
    Dim parts() As String, i As Long, pfx As String, digits As Long
    ItemCode = ""
    ScaleValue = 0
    Code = Trim$(Code)
    If Len(Code) <> 13 Then Exit Function
    If Not IsDigits(Code) Then Exit Function
    digits = Nz(ScaleSetting("ScaleItemDigits"), 5)
    If digits < 4 Or digits > 6 Then Exit Function
    parts = Split(Nz(ScaleSetting("ScalePrefixes"), ""), ",")
    For i = LBound(parts) To UBound(parts)
        pfx = Trim$(parts(i))
        If Len(pfx) >= 1 And Len(pfx) <= 3 And IsDigits(pfx) Then
            If Left$(Code, Len(pfx)) = pfx And Len(pfx) + digits <= 11 Then
                If ScaleCheckDigit(Left$(Code, 12)) <> CLng(Right$(Code, 1)) Then Exit Function
                ItemCode = Mid$(Code, Len(pfx) + 1, digits)
                ScaleValue = CLng(Mid$(Code, Len(pfx) + digits + 1, 12 - Len(pfx) - digits))
                ScaleCodeParts = True
                Exit Function
            End If
        End If
    Next
End Function

Public Function ScaleProductID(ByVal ItemCode As String) As Long
    ' The active product whose code is the item code, else the item code without its leading zeros; 0 = none.
    Dim v As Variant, stripped As String
    v = DbValue("SELECT Min(ProductID) FROM Products WHERE IsActive <> 0 AND ProductCode = " & SqlText(ItemCode))
    If IsNull(v) Then
        stripped = ItemCode
        Do While Left$(stripped, 1) = "0"
            stripped = Mid$(stripped, 2)
        Loop
        If Len(stripped) > 0 And stripped <> ItemCode Then
            v = DbValue("SELECT Min(ProductID) FROM Products WHERE IsActive <> 0 AND ProductCode = " & SqlText(stripped))
        End If
    End If
    ScaleProductID = Nz(v, 0)
End Function

Public Function ScaleLine(ByVal Code As String, ByVal CustomerID As Long, ByRef ProductID As Long, ByRef Qty As Currency, _
                          ByRef Problem As String) As Boolean
    ' True when Code is a scale barcode: then ProductID and Qty of the line, or Problem.
    Dim itemNo As String, scaleVal As Long, price As Currency
    ProductID = 0
    Qty = 0
    Problem = ""
    If Not ScaleCodeParts(Code, itemNo, scaleVal) Then Exit Function
    ScaleLine = True
    ProductID = ScaleProductID(itemNo)
    If ProductID = 0 Then
        Problem = "باركود الميزان " & Code & ": لا يوجد صنف نشط بالكود " & itemNo & "."
        Exit Function
    End If
    If Nz(ScaleSetting("ScaleValueKind"), "WEIGHT") = "PRICE" Then
        price = UnitSalePrice(ProductID, Null, CustomerID)                ' modProductUnits
        If price <= 0 Then
            Problem = "باركود الميزان " & Code & ": سعر الصنف صفر، فلا تُحسب الكمية من السعر."
            ProductID = 0
            Exit Function
        End If
        Qty = Int(CDec(scaleVal) / 100 / CDec(price) * 1000 + 0.5) / 1000
    Else
        Qty = CCur(scaleVal) / 1000
    End If
    If Qty <= 0 Then
        Problem = "باركود الميزان " & Code & ": الكمية صفر."
        ProductID = 0
    End If
End Function

'------------------------------------------------------------------------------
' In-Access test (Immediate window: TestScale). Works inside a transaction that is rolled back.
'------------------------------------------------------------------------------
Private Sub CheckSc(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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

Private Function WithCheck(ByVal Digits12 As String) As String
    WithCheck = Digits12 & CStr(ScaleCheckDigit(Digits12))
End Function

Public Function TestScale() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, pid As Long, found As Long, q As Currency, msg As String
    Dim passed As Long, failed As Long, Report As String, itemNo As String, scaleVal As Long
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestScale  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "UPDATE LabelSettings SET ScalePrefixes = '20, 21', ScaleItemDigits = 5, ScaleValueKind = 'WEIGHT' " & _
                      "WHERE LabelSettingID = 1", dbFailOnError + dbSeeChanges
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, SellingPrice, PurchasePrice, " & _
                      "CurrentQuantity, AverageCost) VALUES ('98761', 'TEST صنف الميزان', " & _
                      Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1) & ", " & _
                      Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1) & ", 25, 10, 0, 0)", dbFailOnError + dbSeeChanges
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = '98761'")

    CheckSc ScaleCheckDigit("629104150021") = 3, "رقم التحقق EAN-13", passed, failed, Report
    CheckSc ScaleCodeParts(WithCheck("209876101250"), itemNo, scaleVal) And itemNo = "98761" And scaleVal = 1250, _
            "قراءة الرمز والوزن", passed, failed, Report
    CheckSc Not ScaleCodeParts(Left$(WithCheck("209876101250"), 12) & "0", itemNo, scaleVal) Or _
            ScaleCheckDigit("209876101250") = 0, "رقم تحقق خطأ = ليس باركود ميزان", passed, failed, Report
    CheckSc Not ScaleCodeParts(WithCheck("229876101250"), itemNo, scaleVal), "بادئة ليست في الإعدادات", passed, failed, Report
    CheckSc ScaleLine(WithCheck("219876101250"), 0, found, q, msg) And found = pid And q = 1.25 And Len(msg) = 0, _
            "سطر بالوزن: 1.25 كجم", passed, failed, Report
    CheckSc ScaleLine(WithCheck("201111101250"), 0, found, q, msg) And found = 0 And Len(msg) > 0, _
            "رمز صنف غير موجود", passed, failed, Report
    CurrentDb.Execute "UPDATE LabelSettings SET ScaleValueKind = 'PRICE' WHERE LabelSettingID = 1", dbFailOnError + dbSeeChanges
    CheckSc ScaleLine(WithCheck("209876101250"), 0, found, q, msg) And found = pid And q = 0.5, _
            "سطر بالسعر: 12.50 / 25 = 0.5", passed, failed, Report
    CurrentDb.Execute "UPDATE LabelSettings SET ScalePrefixes = Null WHERE LabelSettingID = 1", dbFailOnError + dbSeeChanges
    CheckSc Not ScaleLine(WithCheck("209876101250"), 0, found, q, msg), "بلا بادئات = لا ميزان", passed, failed, Report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckSc False, "خطأ: " & Err.Description, passed, failed, Report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات باركود الميزان ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestScale"
        TestScale = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & Report, vbExclamation + MSG_RTL, _
                "TestScale"
    End If
End Function
