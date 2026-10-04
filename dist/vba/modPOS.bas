Attribute VB_Name = "modPOS"
'==============================================================================
' modPOS  -  Retail Store Management System (Phase 6)
'
' Behaviour of the sales screens:
'   frmPOS + frmPOSLines            point of sale (cart in local table tmpPOSLines)
'   frmSalesReturn + frmReturnLines credit note (lines in local table tmpReturnLines)
'   frmCustomerPayment              receipt voucher
'   frmSalesInvoice                 read-only invoice view (search, reprint, return)
' The cart tables live in the front-end file, so every cashier PC has its own.
' Calculations and posting are in modSales.
'==============================================================================
Option Compare Database
Option Explicit

' 2 = acViewPreview shows the invoice first; 0 = acViewNormal prints immediately.
Public Const POS_PRINT_VIEW As Integer = 2

'------------------------------------------------------------------------------
' Local working tables
'------------------------------------------------------------------------------
Public Sub EnsureLocalTables()
    If Not LocalTableExists("tmpPOSLines") Then
        CurrentDb.Execute "CREATE TABLE tmpPOSLines (LineNo COUNTER CONSTRAINT pkPOSLines PRIMARY KEY, " & _
            "ProductID LONG, ProductCode TEXT(30), ProductName TEXT(150), Quantity CURRENCY, " & _
            "UnitPrice CURRENCY, LineDiscount CURRENCY, LineTotal CURRENCY, Available CURRENCY)", dbFailOnError
    End If
    If Not LocalTableExists("tmpReturnLines") Then
        CurrentDb.Execute "CREATE TABLE tmpReturnLines (SalesDetailID LONG CONSTRAINT pkReturnLines " & _
            "PRIMARY KEY, ProductID LONG, ProductName TEXT(150), SoldQty CURRENCY, ReturnedQty CURRENCY, " & _
            "AvailableQty CURRENCY, ReturnQty CURRENCY, ReturnToStock BIT, ReturnAmount CURRENCY)", dbFailOnError
    End If
End Sub

Private Function LocalTableExists(ByVal TableName As String) As Boolean
    Dim tdf As DAO.TableDef
    For Each tdf In CurrentDb.TableDefs
        If StrComp(tdf.Name, TableName, vbTextCompare) = 0 And Len(tdf.Connect) = 0 Then
            LocalTableExists = True
            Exit Function
        End If
    Next
End Function

'==============================================================================
' Point of sale
'==============================================================================
Public Sub POSLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    EnsureLocalTables
    If DCount("*", "tmpPOSLines") > 0 Then
        If Not AskYesNo(" ÊÃœ ›« Ê—… €Ì— „ﬂ „·… „‰ Ã·”… ”«»ﬁ…. Â·  —Ìœ «” ﬂ„«·Â«ø") Then
            CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
        End If
    End If
    ResetSaleHeader frm
    RecalcPOS frm
    SafeFocus frm!txtBarcode
End Sub

Public Sub ResetSaleHeader(ByVal frm As Access.Form)
    frm!cboCustomer.Value = Nz(SettingValue("DefaultCustomerID"), 1)
    frm!cboPaymentType.Value = "CASH"
    frm!cboPaymentMethod.Value = 1
    frm!txtInvoiceDiscount.Value = 0
    frm!txtTendered.Value = Null
    frm!txtNotes.Value = Null
    frm!txtQty.Value = 1
    CustomerChanged frm
End Sub

Public Sub POSKeyDown(ByVal frm As Access.Form, ByRef KeyCode As Integer, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyF9:  KeyCode = 0: SavePOS frm, False
        Case vbKeyF12: KeyCode = 0: SavePOS frm, True
        Case vbKeyF5:  KeyCode = 0: NewSale frm
        Case vbKeyF4:  KeyCode = 0: SafeFocus frm!cboProduct
        Case vbKeyF8:  KeyCode = 0: SafeFocus frm!txtTendered
        Case vbKeyF2:  KeyCode = 0: SafeFocus frm!txtBarcode
        Case vbKeyEscape
            KeyCode = 0
            DoCmd.Close acForm, frm.Name
    End Select
End Sub

Public Sub BarcodeKeyDown(ByVal frm As Access.Form, ByRef KeyCode As Integer)
    ' Scanners send the code followed by Enter.
    Dim code As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    KeyCode = 0
    code = Trim$(frm!txtBarcode.Text)
    frm!txtBarcode.Text = ""
    If Len(code) > 0 Then ScanCode frm, code
End Sub

Public Function ScanCode(ByVal frm As Access.Form, ByVal Code As String) As Boolean
    ' "3*6281234567890" adds 3 units. Matches the barcode or the product code.
    Dim qty As Currency, id As Variant, p As Long
    qty = Nz(frm!txtQty.Value, 1)
    p = InStr(Code, "*")
    If p > 1 Then
        If IsNumeric(Left$(Code, p - 1)) Then qty = CCur(Left$(Code, p - 1))
        Code = Mid$(Code, p + 1)
    End If
    id = DLookup("ProductID", "Products", "IsActive = True AND (Barcode = " & SqlText(Code) & _
                 " OR ProductCode = " & SqlText(Code) & ")")
    If IsNull(id) Then
        Beep
        SetPOSStatus frm, "·„ Ì „ «·⁄ÀÊ— ⁄·Ï „‰ Ã »«·»«—ﬂÊœ √Ê «·ﬂÊœ: " & Code, CLR_DANGER
        Exit Function
    End If
    If qty <= 0 Then qty = 1
    AddLine frm, CLng(id), qty
    frm!txtQty.Value = 1
    ScanCode = True
End Function

Public Sub ProductPicked(ByVal frm As Access.Form)
    If IsNull(frm!cboProduct.Value) Then Exit Sub
    AddLine frm, CLng(frm!cboProduct.Value), Nz(frm!txtQty.Value, 1)
    frm!cboProduct.Value = Null
    frm!txtQty.Value = 1
    SafeFocus frm!txtBarcode
End Sub

Public Sub AddLine(ByVal frm As Access.Form, ByVal ProductID As Long, ByVal Qty As Currency)
    Dim rs As DAO.Recordset, p As DAO.Recordset, total As Currency
    Set p = CurrentDb.OpenRecordset("SELECT ProductCode, ProductName, SellingPrice, CurrentQuantity " & _
                                    "FROM Products WHERE ProductID = " & ProductID, dbOpenSnapshot)
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tmpPOSLines WHERE ProductID = " & ProductID & _
                                     " AND LineDiscount = 0", dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!ProductID = ProductID
        rs!ProductCode = p!ProductCode
        rs!ProductName = p!ProductName
        rs!Quantity = Qty
        rs!UnitPrice = p!SellingPrice
        rs!LineDiscount = 0
    Else
        rs.Edit
        rs!Quantity = rs!Quantity + Qty
    End If
    rs!Available = p!CurrentQuantity
    total = rs!Quantity
    rs.Update
    rs.Close
    RecalcPOS frm
    If total > p!CurrentQuantity Then
        SetPOSStatus frm, " ‰»ÌÂ: «·ﬂ„Ì… «·„ Ê›—… „‰ ´" & p!ProductName & "ª ÂÌ " & p!CurrentQuantity, CLR_WARNING
    Else
        SetPOSStatus frm, " „  ≈÷«›…: " & p!ProductName & "  (" & total & ")", CLR_SUCCESS
    End If
    p.Close
End Sub

Public Sub LineChanged(ByVal sf As Access.Form, ByVal FieldName As String)
    ' AfterUpdate of Quantity / UnitPrice / LineDiscount in frmPOSLines.
    Dim price As Variant
    If Nz(sf!Quantity.Value, 0) <= 0 Then
        ShowWarning "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ √ﬂ»— „‰ ’›—."
        sf!Quantity.Value = 1
    End If
    If Nz(sf!LineDiscount.Value, 0) < 0 Then sf!LineDiscount.Value = 0
    If FieldName = "UnitPrice" Then
        price = DLookup("SellingPrice", "Products", "ProductID = " & sf!ProductID.Value)
        If sf!UnitPrice.Value <> price And Not HasPermission("PRICE_OVERRIDE") Then
            ShowWarning "·«  „·ﬂ ’·«ÕÌ…  ⁄œÌ· ”⁄— «·»Ì⁄."
            sf!UnitPrice.Value = price
        End If
    End If
    If sf.Dirty Then sf.Dirty = False
    RecalcPOS sf.Parent
End Sub

Public Sub RemoveCurrentLine(ByVal sf As Access.Form)
    If sf.NewRecord Or IsNull(sf!LineNo.Value) Then Exit Sub
    CurrentDb.Execute "DELETE FROM tmpPOSLines WHERE LineNo = " & sf!LineNo.Value, dbFailOnError
    RecalcPOS sf.Parent
    SetPOSStatus sf.Parent, " „ Õ–› «·”ÿ—", CLR_MUTED
End Sub

Public Sub RecalcPOS(ByVal frm As Access.Form)
    ' Recalculates every line and the totals with the same engine used for posting.
    Dim rs As DAO.Recordset, msg As String, i As Long, n As Long, rate As Currency, vatRate As Currency
    Dim total As Currency, tendered As Currency
    vatRate = Nz(SettingValue("VATRate"), 0.15)
    CalcReset
    ' read through a snapshot of the join, write through the cart table alone (same order)
    Set rs = CurrentDb.OpenRecordset("SELECT t.LineNo, t.Quantity, t.UnitPrice, t.LineDiscount, p.VATCategory " & _
                                     "FROM tmpPOSLines AS t LEFT JOIN Products AS p ON t.ProductID = p.ProductID " & _
                                     "ORDER BY t.LineNo", dbOpenSnapshot)
    Do Until rs.EOF
        If Nz(rs!VATCategory, "S") = "S" Then rate = vatRate Else rate = 0
        CalcAddLine Nz(rs!Quantity, 0), Nz(rs!UnitPrice, 0), Nz(rs!LineDiscount, 0), rate
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    If n > 0 Then msg = CalcRun(Nz(frm!txtInvoiceDiscount.Value, 0), Nz(SettingValue("PricesIncludeVAT"), True))
    If n > 0 And Len(msg) = 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT LineNo, LineTotal FROM tmpPOSLines ORDER BY LineNo", dbOpenDynaset)
        For i = 0 To n - 1
            rs.Edit
            rs!LineTotal = CalcLine(i, "TOTAL")
            rs.Update
            rs.MoveNext
        Next
        rs.Close
    End If
    frm!subLines.Form.Requery

    If n = 0 Or Len(msg) > 0 Then
        frm!lblSubTotal.Caption = "0.00"
        frm!lblDiscount.Caption = "0.00"
        frm!lblTax.Caption = "0.00"
        frm!lblTotal.Caption = "0.00"
        frm!lblChange.Caption = " "
        frm!lblItems.Caption = n & " ’‰›"
        If Len(msg) > 0 Then SetPOSStatus frm, msg, CLR_DANGER
        Exit Sub
    End If
    total = CalcTotal("TOTAL")
    frm!lblSubTotal.Caption = Format$(CalcTotal("SUBTOTAL"), "#,##0.00")
    frm!lblDiscount.Caption = Format$(CalcTotal("DISCOUNT"), "#,##0.00")
    frm!lblTax.Caption = Format$(CalcTotal("TAX"), "#,##0.00")
    frm!lblTotal.Caption = Format$(total, "#,##0.00")
    frm!lblItems.Caption = n & " ’‰›"
    If IsNull(frm!txtTendered.Value) Then
        frm!lblChange.Caption = " "
    Else
        tendered = frm!txtTendered.Value
        If frm!cboPaymentType.Value = "CREDIT" Then
            frm!lblChange.Caption = "⁄·Ï «·Õ”«»: " & Format$(total - tendered, "#,##0.00")
            frm!lblChange.ForeColor = CLR_WARNING
        ElseIf tendered >= total Then
            frm!lblChange.Caption = "«·»«ﬁÌ ··⁄„Ì·: " & Format$(tendered - total, "#,##0.00")
            frm!lblChange.ForeColor = CLR_SUCCESS
        Else
            frm!lblChange.Caption = "‰«ﬁ’: " & Format$(total - tendered, "#,##0.00")
            frm!lblChange.ForeColor = CLR_DANGER
        End If
    End If
End Sub

Public Sub CustomerChanged(ByVal frm As Access.Form)
    Dim rs As DAO.Recordset
    If IsNull(frm!cboCustomer.Value) Then frm!cboCustomer.Value = Nz(SettingValue("DefaultCustomerID"), 1)
    Set rs = CurrentDb.OpenRecordset("SELECT CurrentBalance, CreditLimit, AllowCredit FROM Customers " & _
                                     "WHERE CustomerID = " & frm!cboCustomer.Value, dbOpenSnapshot)
    If rs.EOF Then
        frm!lblCustomerInfo.Caption = " "
    ElseIf frm!cboCustomer.Value = Nz(SettingValue("DefaultCustomerID"), 1) Then
        frm!lblCustomerInfo.Caption = "»Ì⁄ ‰ﬁœÌ"
        frm!cboPaymentType.Value = "CASH"
    Else
        frm!lblCustomerInfo.Caption = "«·—’Ìœ: " & Format$(rs!CurrentBalance, "#,##0.00") & _
            IIf(rs!CreditLimit > 0, "   Õœ «·«∆ „«‰: " & Format$(rs!CreditLimit, "#,##0.00"), "") & _
            IIf(rs!AllowCredit, "", "   (€Ì— „”„ÊÕ »«·¬Ã·)")
        If Not rs!AllowCredit Then frm!cboPaymentType.Value = "CASH"
    End If
    rs.Close
End Sub

Public Sub PaymentTypeChanged(ByVal frm As Access.Form)
    If frm!cboPaymentType.Value = "CREDIT" Then
        If frm!cboCustomer.Value = Nz(SettingValue("DefaultCustomerID"), 1) Then
            ShowWarning "«·»Ì⁄ «·¬Ã· ÌÕ «Ã ⁄„Ì·« „”Ã·«. «Œ — «·⁄„Ì· √Ê·«."
            frm!cboPaymentType.Value = "CASH"
            SafeFocus frm!cboCustomer
        End If
    End If
    RecalcPOS frm
End Sub

Public Sub PaymentMethodChanged(ByVal frm As Access.Form)
    ' Card / transfer: the exact total is paid.
    If Nz(frm!cboPaymentMethod.Value, 1) <> 1 And frm!cboPaymentType.Value = "CASH" Then
        frm!txtTendered.Value = CalcTotal("TOTAL")
    End If
    RecalcPOS frm
End Sub

Public Function SavePOS(ByVal frm As Access.Form, ByVal PrintAfter As Boolean) As Boolean
    Dim msg As String, newID As Long, invNo As String, change As Currency
    If frm!subLines.Form.Dirty Then frm!subLines.Form.Dirty = False
    msg = PostSaleFromCart(Nz(frm!cboCustomer.Value, 1), Nz(frm!cboPaymentType.Value, "CASH"), _
                           frm!cboPaymentMethod.Value, Nz(frm!txtInvoiceDiscount.Value, 0), _
                           frm!txtTendered.Value, Nz(frm!txtNotes.Value, ""), newID)
    If Len(msg) > 0 Then
        SetPOSStatus frm, msg, CLR_DANGER
        ShowWarning msg
        Exit Function
    End If
    invNo = DLookup("InvoiceNumber", "SalesInvoices", "SalesInvoiceID = " & newID)
    change = DLookup("ChangeDue", "SalesInvoices", "SalesInvoiceID = " & newID)
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    ResetSaleHeader frm
    RecalcPOS frm
    SetPOSStatus frm, " „ Õ›Ÿ «·›« Ê—… " & invNo & IIf(change > 0, "  -  «·»«ﬁÌ ··⁄„Ì·: " & _
                 Format$(change, "#,##0.00"), ""), CLR_SUCCESS
    frm!lblLastInvoice.Caption = invNo
    If PrintAfter Then PrintSalesDocument "SALE", newID
    SafeFocus frm!txtBarcode
    SavePOS = True
End Function

Public Sub NewSale(ByVal frm As Access.Form)
    If DCount("*", "tmpPOSLines") > 0 Then
        If Not AskYesNo("Â·  —Ìœ ≈·€«¡ «·›« Ê—… «·Õ«·Ì… Ê«·»œ¡ »›« Ê—… ÃœÌœ…ø") Then Exit Sub
    End If
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    ResetSaleHeader frm
    RecalcPOS frm
    SetPOSStatus frm, "›« Ê—… ÃœÌœ…", CLR_MUTED
    SafeFocus frm!txtBarcode
End Sub

Public Function POSUnload(ByVal frm As Access.Form) As Boolean
    POSUnload = True
    If DCount("*", "tmpPOSLines") > 0 Then
        POSUnload = AskYesNo("«·›« Ê—… «·Õ«·Ì… ·„  ıÕ›Ÿ (” »ﬁÏ „Õ›ÊŸ… „ƒﬁ « ·Â–« «·ÃÂ«“). Â·  —Ìœ «·Œ—ÊÃø")
    End If
End Function

Public Sub ReprintLast(ByVal frm As Access.Form)
    Dim id As Variant
    id = DLookup("SalesInvoiceID", "SalesInvoices", "InvoiceNumber = " & SqlText(Nz(frm!lblLastInvoice.Caption, "")))
    If IsNull(id) Then id = DMax("SalesInvoiceID", "SalesInvoices", "EmployeeID = " & CurrentUserID())
    If IsNull(id) Then
        ShowInfo "·«  ÊÃœ ›« Ê—… ”«»ﬁ…."
    Else
        PrintSalesDocument "SALE", CLng(id)
    End If
End Sub

Private Sub SetPOSStatus(ByVal frm As Access.Form, ByVal Text As String, ByVal Color As Long)
    frm!lblStatus.Caption = IIf(Len(Text) = 0, " ", Text)
    frm!lblStatus.ForeColor = Color
End Sub

'==============================================================================
' Sales return (credit note)
'==============================================================================
Public Sub ReturnLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    EnsureLocalTables
    CurrentDb.Execute "DELETE FROM tmpReturnLines", dbFailOnError
    frm!subReturnLines.Form.Requery
    frm!txtInvoiceID.Value = Null
    If Not IsNull(frm.OpenArgs) Then LoadInvoiceForReturn frm, CLng(frm.OpenArgs)
End Sub

Public Sub ReturnFindInvoice(ByVal frm As Access.Form)
    Dim id As Variant
    id = DLookup("SalesInvoiceID", "SalesInvoices", "InvoiceNumber = " & SqlText(Trim$(Nz(frm!txtInvoiceNo.Value, ""))))
    If IsNull(id) Then
        ShowWarning "·«  ÊÃœ ›« Ê—… »Â–« «·—ﬁ„."
        Exit Sub
    End If
    LoadInvoiceForReturn frm, CLng(id)
End Sub

Public Sub LoadInvoiceForReturn(ByVal frm As Access.Form, ByVal SalesInvoiceID As Long)
    Dim rs As DAO.Recordset
    CurrentDb.Execute "DELETE FROM tmpReturnLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpReturnLines (SalesDetailID, ProductID, ProductName, SoldQty, ReturnedQty, " & _
        "AvailableQty, ReturnQty, ReturnToStock, ReturnAmount) " & _
        "SELECT d.SalesDetailID, d.ProductID, p.ProductName, d.Quantity, CCur(Nz(r.QtyReturned, 0)), " & _
        "d.Quantity - CCur(Nz(r.QtyReturned, 0)), 0, True, 0 " & _
        "FROM (SalesInvoiceDetails AS d INNER JOIN Products AS p ON d.ProductID = p.ProductID) " & _
        "LEFT JOIN qrySalesReturnedQty AS r ON d.SalesDetailID = r.SalesDetailID " & _
        "WHERE d.SalesInvoiceID = " & SalesInvoiceID, dbFailOnError
    Set rs = CurrentDb.OpenRecordset("SELECT h.*, c.CustomerName FROM SalesInvoices AS h INNER JOIN Customers " & _
                                     "AS c ON h.CustomerID = c.CustomerID WHERE h.SalesInvoiceID = " & _
                                     SalesInvoiceID, dbOpenSnapshot)
    frm!txtInvoiceID.Value = SalesInvoiceID
    frm!txtInvoiceNo.Value = rs!InvoiceNumber
    frm!lblInvoiceInfo.Caption = "«· «—ÌŒ: " & Format$(rs!InvoiceDate, "yyyy/mm/dd hh:nn") & "    «·⁄„Ì·: " & _
        rs!CustomerName & "    «·≈Ã„«·Ì: " & Format$(rs!TotalAmount, "#,##0.00") & "    «·‰Ê⁄: " & _
        IIf(rs!PaymentType = "CREDIT", "¬Ã·", "‰ﬁœÌ")
    If rs!CustomerID = Nz(SettingValue("DefaultCustomerID"), 1) Then
        frm!cboRefundType.Value = "CASH"
        frm!cboRefundType.Locked = True
    Else
        frm!cboRefundType.Locked = False
        frm!cboRefundType.Value = IIf(rs!PaymentType = "CREDIT", "CREDIT", "CASH")
    End If
    rs.Close
    frm!subReturnLines.Form.Requery
    RecalcReturn frm
End Sub

Public Sub ReturnLineChanged(ByVal sf As Access.Form)
    If Nz(sf!ReturnQty.Value, 0) < 0 Then sf!ReturnQty.Value = 0
    If Nz(sf!ReturnQty.Value, 0) > Nz(sf!AvailableQty.Value, 0) Then
        ShowWarning "«·ﬂ„Ì… «·„— Ã⁄… √ﬂ»— „‰ «·„ «Õ ··≈—Ã«⁄ (" & sf!AvailableQty.Value & ")."
        sf!ReturnQty.Value = sf!AvailableQty.Value
    End If
    If sf.Dirty Then sf.Dirty = False
    RecalcReturn sf.Parent
End Sub

Public Sub ReturnAll(ByVal frm As Access.Form)
    CurrentDb.Execute "UPDATE tmpReturnLines SET ReturnQty = AvailableQty", dbFailOnError
    RecalcReturn frm
End Sub

Public Sub RecalcReturn(ByVal frm As Access.Form)
    ' Shows the credit-note value of each line with the same rule used for posting.
    Dim rs As DAO.Recordset, d As DAO.Recordset, total As Currency, net As Currency, tax As Currency
    Dim lineTotal As Currency, msg As String
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tmpReturnLines", dbOpenDynaset)
    Do Until rs.EOF
        lineTotal = 0
        If Nz(rs!ReturnQty, 0) > 0 Then
            Set d = CurrentDb.OpenRecordset("SELECT d.Quantity, d.LineTotal, d.Tax, CCur(Nz(r.QtyReturned, 0)) AS PQ, " & _
                "(SELECT CCur(Nz(Sum(LineTotal), 0)) FROM SalesReturnDetails WHERE SalesDetailID = d.SalesDetailID) AS PT, " & _
                "(SELECT CCur(Nz(Sum(Tax), 0)) FROM SalesReturnDetails WHERE SalesDetailID = d.SalesDetailID) AS PX " & _
                "FROM SalesInvoiceDetails AS d LEFT JOIN qrySalesReturnedQty AS r ON d.SalesDetailID = r.SalesDetailID " & _
                "WHERE d.SalesDetailID = " & rs!SalesDetailID, dbOpenSnapshot)
            msg = ReturnAmounts(d!Quantity, d!LineTotal, d!Tax, d!PQ, d!PT, d!PX, rs!ReturnQty, net, tax, lineTotal)
            d.Close
            If Len(msg) > 0 Then lineTotal = 0
        End If
        rs.Edit
        rs!ReturnAmount = lineTotal
        rs.Update
        total = total + lineTotal
        rs.MoveNext
    Loop
    rs.Close
    frm!subReturnLines.Form.Requery
    frm!lblReturnTotal.Caption = Format$(total, "#,##0.00")
End Sub

Public Function SaveReturn(ByVal frm As Access.Form, ByVal PrintAfter As Boolean) As Boolean
    Dim msg As String, newID As Long, retNo As String
    If IsNull(frm!txtInvoiceID.Value) Then
        ShowWarning "«Œ — «·›« Ê—… «·√’·Ì… √Ê·«."
        Exit Function
    End If
    If frm!subReturnLines.Form.Dirty Then frm!subReturnLines.Form.Dirty = False
    msg = PostSalesReturn(frm!txtInvoiceID.Value, Nz(frm!txtReason.Value, ""), Nz(frm!cboRefundType.Value, "CASH"), _
                          frm!cboPaymentMethod.Value, newID)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    retNo = DLookup("ReturnNumber", "SalesReturns", "SalesReturnID = " & newID)
    ShowInfo " „ Õ›Ÿ «·„— Ã⁄ " & retNo & "  »ﬁÌ„… " & Format$(DLookup("TotalAmount", "SalesReturns", _
             "SalesReturnID = " & newID), "#,##0.00")
    If PrintAfter Then PrintSalesDocument "RETURN", newID
    LoadInvoiceForReturn frm, frm!txtInvoiceID.Value
    frm!txtReason.Value = Null
    SaveReturn = True
End Function

'==============================================================================
' Receipt voucher
'==============================================================================
Public Sub PaymentLoad(ByVal frm As Access.Form)
    If Not IsNull(frm.OpenArgs) Then frm!cboCustomer.Value = CLng(frm.OpenArgs)
    frm!cboPaymentMethod.Value = 1
    PaymentCustomerChanged frm
End Sub

Public Sub PaymentCustomerChanged(ByVal frm As Access.Form)
    If IsNull(frm!cboCustomer.Value) Then
        frm!lblBalance.Caption = " "
    Else
        frm!lblBalance.Caption = "«·—’Ìœ «·„” Õﬁ: " & Format$(Nz(DLookup("CurrentBalance", "Customers", _
                                 "CustomerID = " & frm!cboCustomer.Value), 0), "#,##0.00")
    End If
End Sub

Public Function SavePayment(ByVal frm As Access.Form) As Boolean
    Dim msg As String, newID As Long
    If IsNull(frm!cboCustomer.Value) Then
        ShowWarning "«Œ — «·⁄„Ì·."
        Exit Function
    End If
    msg = PostCustomerPayment(frm!cboCustomer.Value, Nz(frm!txtAmount.Value, 0), Nz(frm!cboPaymentMethod.Value, 1), _
                              Nz(frm!txtNotes.Value, ""), newID)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    ShowInfo " „ Õ›Ÿ ”‰œ «·ﬁ»÷ " & DLookup("PaymentNumber", "CustomerPayments", "PaymentID = " & newID)
    frm!txtAmount.Value = Null
    frm!txtNotes.Value = Null
    PaymentCustomerChanged frm
    SavePayment = True
End Function

'==============================================================================
' Invoice view
'==============================================================================
Public Sub InvoiceViewLoad(ByVal frm As Access.Form)
    Dim rs As DAO.Recordset, id As Long
    Calendar = vbCalGreg
    If IsNull(frm.OpenArgs) Then
        ShowWarning "·„  ıÕœœ «·›« Ê—…."
        Exit Sub
    End If
    id = CLng(frm.OpenArgs)
    Set rs = CurrentDb.OpenRecordset("SELECT h.*, c.CustomerName, e.EmployeeName FROM (SalesInvoices AS h " & _
        "INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID) INNER JOIN Employees AS e " & _
        "ON h.EmployeeID = e.EmployeeID WHERE h.SalesInvoiceID = " & id, dbOpenSnapshot)
    If rs.EOF Then
        ShowWarning "«·›« Ê—… €Ì— „ÊÃÊœ…."
        Exit Sub
    End If
    frm!txtInvoiceID.Value = id
    frm!lblTitle.Caption = "›« Ê—… " & rs!InvoiceNumber
    frm!lblHeader.Caption = "«· «—ÌŒ: " & Format$(rs!InvoiceDate, "yyyy/mm/dd hh:nn") & "    «·⁄„Ì·: " & _
        rs!CustomerName & "    «·ﬂ«‘Ì—: " & rs!EmployeeName & "    " & IIf(rs!PaymentType = "CREDIT", "¬Ã·", "‰ﬁœÌ") & _
        "    " & IIf(rs!InvoiceSubType = "STANDARD", "›« Ê—… ÷—Ì»Ì…", "›« Ê—… ÷—Ì»Ì… „»”ÿ…")
    frm!lblTotals.Caption = "ﬁ»· «·÷—Ì»…: " & Format$(rs!TaxableAmount, "#,##0.00") & "    «·÷—Ì»…: " & _
        Format$(rs!Tax, "#,##0.00") & "    «·≈Ã„«·Ì: " & Format$(rs!TotalAmount, "#,##0.00") & _
        "    «·„œ›Ê⁄: " & Format$(rs!PaidAmount, "#,##0.00") & "    «·„ »ﬁÌ: " & Format$(rs!RemainingAmount, "#,##0.00")
    rs.Close
    frm!lstLines.RowSource = "SELECT d.LineNumber AS [#], p.ProductName AS [«·’‰›], d.Quantity AS [«·ﬂ„Ì…], " & _
        "d.UnitPrice AS [«·”⁄— »œÊ‰ ÷—Ì»…], d.Discount AS [«·Œ’„], d.Tax AS [«·÷—Ì»…], d.LineTotal AS [«·≈Ã„«·Ì] " & _
        "FROM SalesInvoiceDetails AS d INNER JOIN Products AS p ON d.ProductID = p.ProductID " & _
        "WHERE d.SalesInvoiceID = " & id & " ORDER BY d.LineNumber"
    frm!lstReturns.RowSource = "SELECT ReturnNumber AS [«·„— Ã⁄], ReturnDate AS [«· «—ÌŒ], TotalAmount AS [«·ﬁÌ„…], " & _
        "Reason AS [«·”»»] FROM SalesReturns WHERE SalesInvoiceID = " & id & " ORDER BY ReturnDate"
End Sub

'==============================================================================
' Printing
'==============================================================================
Public Sub PrintSalesDocument(ByVal DocKind As String, ByVal DocID As Long, Optional ByVal ForceA4 As Boolean = False)
    Dim rpt As String, subType As String
    If DocKind = "SALE" Then
        subType = Nz(DLookup("InvoiceSubType", "SalesInvoices", "SalesInvoiceID = " & DocID), "SIMPLIFIED")
    Else
        subType = Nz(DLookup("InvoiceSubType", "SalesReturns", "SalesReturnID = " & DocID), "SIMPLIFIED")
    End If
    ' B2B (standard) invoices show the buyer details: always A4.
    If ForceA4 Or subType = "STANDARD" Then rpt = "rptSalesInvoiceA4" Else rpt = "rptSalesReceipt"
    If Not ReportExists(rpt) Then
        ShowWarning " ﬁ—Ì— «·ÿ»«⁄… €Ì— „ÊÃÊœ: " & rpt & vbCrLf & "‘€¯· BuildReports."
        Exit Sub
    End If
    DoCmd.OpenReport rpt, POS_PRINT_VIEW, , "[DocKind] = '" & DocKind & "' AND [DocID] = " & DocID
End Sub

Public Sub DrawDocumentQR(ByVal rpt As Access.Report, ByVal DocKind As Variant, ByVal DocID As Variant, _
                          ByVal LeftTwips As Long, ByVal TopTwips As Long, ByVal SizeTwips As Long)
    ' Called from the Print event of the totals section of the invoice reports.
    Dim qr As Variant
    If IsNull(DocID) Then Exit Sub
    If DocKind = "RETURN" Then
        qr = DLookup("QRCodeData", "SalesReturns", "SalesReturnID = " & DocID)
    Else
        qr = DLookup("QRCodeData", "SalesInvoices", "SalesInvoiceID = " & DocID)
    End If
    If Not IsNull(qr) Then DrawQR rpt, CStr(qr), LeftTwips, TopTwips, SizeTwips
End Sub
