Attribute VB_Name = "modPurchaseScreens"
'==============================================================================
' modPurchaseScreens  -  Retail Store Management System (Phases 7-8)
'
' Behaviour of the purchase and inventory screens:
'   frmPurchaseInvoice + frmPurchaseLines            purchase invoice (cart in tmpPurchaseLines)
'   frmPurchaseReturn + frmPurchaseReturnLines       goods returned to the supplier
'   frmSupplierPayment                               payment voucher
'   frmPurchaseView                                  read-only purchase invoice
'   frmInventory                                     balances, movements, manual movements
'   frmStockCount + frmStockCountLines               stocktaking
' Posting is in modPurchases; the calculation engine is in modSales.
'==============================================================================
Option Compare Database
Option Explicit

Private Const COUNT_LINES_SQL As String = "SELECT d.StockCountDetailID, d.StockCountID, d.ProductID, " & _
    "d.SystemQuantity, d.ActualQuantity, d.Difference, d.UnitCost, d.DifferenceValue, d.Notes, " & _
    "p.ProductCode, p.ProductName, p.Barcode FROM StockCountDetails AS d " & _
    "INNER JOIN Products AS p ON d.ProductID = p.ProductID"

'==============================================================================
' Purchase invoice
'==============================================================================
Public Sub PurchaseLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    EnsureLocalTables
    If DCount("*", "tmpPurchaseLines") > 0 Then
        If Not AskYesNo(" ÊÃœ ›« Ê—… „‘ —Ì«  €Ì— „ﬂ „·… „‰ Ã·”… ”«»ﬁ…. Â·  —Ìœ «” ﬂ„«·Â«ø") Then
            CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
        End If
    End If
    ResetPurchaseHeader frm
    RecalcPurchase frm
    SafeFocus frm!cboSupplier
End Sub

Private Sub ResetPurchaseHeader(ByVal frm As Access.Form)
    frm!cboSupplier.Value = Null
    frm!txtSupplierInvoiceNo.Value = Null
    frm!txtInvoiceDate.Value = Date
    frm!cboPaymentType.Value = "CASH"
    frm!cboPaymentMethod.Value = 1
    frm!txtInvoiceDiscount.Value = 0
    frm!chkChargeVAT.Value = True
    frm!txtPaid.Value = Null
    frm!txtPaid.Locked = True                ' a cash purchase is paid in full
    frm!txtNotes.Value = Null
    frm!txtQty.Value = 1
    frm!lblSupplierInfo.Caption = " "
End Sub

Public Sub PurchaseKeyDown(ByVal frm As Access.Form, ByRef KeyCode As Integer, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyF9: KeyCode = 0: SavePurchase frm
        Case vbKeyF5: KeyCode = 0: NewPurchase frm
        Case vbKeyF4: KeyCode = 0: SafeFocus frm!cboProduct
        Case vbKeyF2: KeyCode = 0: SafeFocus frm!txtBarcode
        Case vbKeyEscape
            KeyCode = 0
            DoCmd.Close acForm, frm.Name
    End Select
End Sub

Public Sub PurBarcodeKeyDown(ByVal frm As Access.Form, ByRef KeyCode As Integer)
    Dim code As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    KeyCode = 0
    code = Trim$(frm!txtBarcode.Text)
    frm!txtBarcode.Text = ""
    If Len(code) > 0 Then PurScanCode frm, code
End Sub

Public Function PurScanCode(ByVal frm As Access.Form, ByVal Code As String) As Boolean
    ' Same input rules as the point of sale: barcode or product code, "12*code" adds 12.
    Dim qty As Currency, id As Variant
    qty = Nz(frm!txtQty.Value, 1)
    SplitQtyPrefix Code, qty
    id = DLookup("ProductID", "Products", "IsActive = True AND (Barcode = " & SqlText(Code) & _
                 " OR ProductCode = " & SqlText(Code) & ")")
    If IsNull(id) Then
        Beep
        SetPurStatus frm, "·„ Ì „ «·⁄ÀÊ— ⁄·Ï „‰ Ã »«·»«—ﬂÊœ √Ê «·ﬂÊœ: " & Code & _
                     "  (√÷›Â „‰ “— ´„‰ Ã ÃœÌœª)", CLR_DANGER
        Exit Function
    End If
    PurAddLine frm, CLng(id), qty
    frm!txtQty.Value = 1
    PurScanCode = True
End Function

Public Sub PurProductPicked(ByVal frm As Access.Form)
    If IsNull(frm!cboProduct.Value) Then Exit Sub
    PurAddLine frm, CLng(frm!cboProduct.Value), Nz(frm!txtQty.Value, 1)
    frm!cboProduct.Value = Null
    frm!txtQty.Value = 1
    SafeFocus frm!txtBarcode
End Sub

Public Sub PurAddLine(ByVal frm As Access.Form, ByVal ProductID As Long, ByVal Qty As Currency)
    ' New line at the last purchase price; the same product again adds to its line.
    Dim rs As DAO.Recordset, p As DAO.Recordset
    If Qty <= 0 Then Qty = 1
    Set p = CurrentDb.OpenRecordset("SELECT ProductCode, ProductName, PurchasePrice, SellingPrice " & _
                                    "FROM Products WHERE ProductID = " & ProductID, dbOpenSnapshot)
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tmpPurchaseLines WHERE ProductID = " & ProductID & _
                                     " AND LineDiscount = 0", dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!ProductID = ProductID
        rs!ProductCode = p!ProductCode
        rs!ProductName = p!ProductName
        rs!Quantity = Qty
        rs!UnitCost = p!PurchasePrice
        rs!LineDiscount = 0
        rs!SellingPrice = p!SellingPrice
    Else
        rs.Edit
        rs!Quantity = rs!Quantity + Qty
    End If
    rs.Update
    rs.Close
    RecalcPurchase frm
    SetPurStatus frm, " „  ≈÷«›…: " & p!ProductName, CLR_SUCCESS
    p.Close
End Sub

Public Sub PurLineChanged(ByVal sf As Access.Form, ByVal FieldName As String)
    ' AfterUpdate of Quantity / UnitCost / LineDiscount / NewSellingPrice in frmPurchaseLines.
    If Nz(sf!Quantity.Value, 0) <= 0 Then
        ShowWarning "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ √ﬂ»— „‰ ’›—."
        sf!Quantity.Value = 1
    End If
    If Nz(sf!UnitCost.Value, 0) < 0 Then
        ShowWarning "«· ﬂ·›… ·« Ì„ﬂ‰ √‰  ﬂÊ‰ ”«·»…."
        sf!UnitCost.Value = 0
    End If
    If Nz(sf!LineDiscount.Value, 0) < 0 Then sf!LineDiscount.Value = 0
    If FieldName = "NewSellingPrice" And Not IsNull(sf!NewSellingPrice.Value) Then
        If sf!NewSellingPrice.Value < 0 Then
            sf!NewSellingPrice.Value = Null
        ElseIf Not HasPermission("PRODUCTS") Then
            ShowWarning "·«  „·ﬂ ’·«ÕÌ…  ⁄œÌ· √”⁄«— «·»Ì⁄."
            sf!NewSellingPrice.Value = Null
        ElseIf sf!NewSellingPrice.Value < Nz(sf!UnitCost.Value, 0) Then
            ShowWarning " ‰»ÌÂ: ”⁄— «·»Ì⁄ «·ÃœÌœ √ﬁ· „‰  ﬂ·›… «·‘—«¡."
        End If
    End If
    If sf.Dirty Then sf.Dirty = False
    RecalcPurchase sf.Parent
End Sub

Public Sub PurRemoveLine(ByVal sf As Access.Form)
    If sf.NewRecord Or IsNull(sf!LineNo.Value) Then Exit Sub
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines WHERE LineNo = " & sf!LineNo.Value, dbFailOnError
    RecalcPurchase sf.Parent
    SetPurStatus sf.Parent, " „ Õ–› «·”ÿ—", CLR_MUTED
End Sub

Public Sub RecalcPurchase(ByVal frm As Access.Form)
    ' Same engine as the posting (prices WITHOUT VAT).
    Dim rs As DAO.Recordset, msg As String, i As Long, n As Long, rate As Currency, vatRate As Currency
    Dim chargeVat As Boolean, total As Currency
    vatRate = Nz(SettingValue("VATRate"), 0.15)
    chargeVat = Nz(frm!chkChargeVAT.Value, True)
    CalcReset
    Set rs = CurrentDb.OpenRecordset("SELECT t.LineNo, t.Quantity, t.UnitCost, t.LineDiscount, p.VATCategory " & _
                                     "FROM tmpPurchaseLines AS t LEFT JOIN Products AS p ON t.ProductID = p.ProductID " & _
                                     "ORDER BY t.LineNo", dbOpenSnapshot)
    Do Until rs.EOF
        If chargeVat And Nz(rs!VATCategory, "S") = "S" Then rate = vatRate Else rate = 0
        CalcAddLine Nz(rs!Quantity, 0), Nz(rs!UnitCost, 0), Nz(rs!LineDiscount, 0), rate
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    If n > 0 Then msg = CalcRun(Nz(frm!txtInvoiceDiscount.Value, 0), False)
    If n > 0 And Len(msg) = 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT LineNo, LineTotal FROM tmpPurchaseLines ORDER BY LineNo", dbOpenDynaset)
        For i = 0 To n - 1
            rs.Edit
            rs!LineTotal = CalcLine(i, "TOTAL")
            rs.Update
            rs.MoveNext
        Next
        rs.Close
    End If
    frm!subLines.Form.Requery
    frm!lblItems.Caption = n & " ’‰›"

    If n = 0 Or Len(msg) > 0 Then
        frm!lblSubTotal.Caption = "0.00"
        frm!lblDiscount.Caption = "0.00"
        frm!lblTax.Caption = "0.00"
        frm!lblTotal.Caption = "0.00"
        frm!lblRemaining.Caption = " "
        If Len(msg) > 0 Then SetPurStatus frm, msg, CLR_DANGER
        Exit Sub
    End If
    total = CalcTotal("TOTAL")
    frm!lblSubTotal.Caption = Format$(CalcTotal("SUBTOTAL"), "#,##0.00")
    frm!lblDiscount.Caption = Format$(CalcTotal("DISCOUNT"), "#,##0.00")
    frm!lblTax.Caption = Format$(CalcTotal("TAX"), "#,##0.00")
    frm!lblTotal.Caption = Format$(total, "#,##0.00")
    If frm!cboPaymentType.Value = "CREDIT" Then
        frm!lblRemaining.Caption = "⁄·Ï «·Õ”«»: " & Format$(total - Nz(frm!txtPaid.Value, 0), "#,##0.00")
    Else
        frm!lblRemaining.Caption = " ıœ›⁄ »«·ﬂ«„·"
    End If
End Sub

Public Sub PurSupplierChanged(ByVal frm As Access.Form)
    Dim rs As DAO.Recordset
    If IsNull(frm!cboSupplier.Value) Then
        frm!lblSupplierInfo.Caption = " "
        Exit Sub
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT CurrentBalance, VATNumber FROM Suppliers WHERE SupplierID = " & _
                                     frm!cboSupplier.Value, dbOpenSnapshot)
    If Not rs.EOF Then
        frm!lblSupplierInfo.Caption = "«·„” Õﬁ ··„Ê—œ: " & Format$(rs!CurrentBalance, "#,##0.00") & "    " & _
            IIf(IsNull(rs!VATNumber), "€Ì— „”Ã· »«·÷—Ì»…", "«·—ﬁ„ «·÷—Ì»Ì: " & rs!VATNumber)
        ' an unregistered supplier cannot charge VAT
        frm!chkChargeVAT.Value = Not IsNull(rs!VATNumber)
    End If
    rs.Close
    RecalcPurchase frm
End Sub

Public Sub PurPaymentTypeChanged(ByVal frm As Access.Form)
    If frm!cboPaymentType.Value = "CREDIT" Then
        frm!txtPaid.Locked = False
        frm!txtPaid.Value = 0
    Else
        frm!txtPaid.Value = Null
        frm!txtPaid.Locked = True
    End If
    RecalcPurchase frm
End Sub

Public Function SavePurchase(ByVal frm As Access.Form) As Boolean
    Dim msg As String, newID As Long, invNo As String, paidNow As Variant
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Function      ' frmUserScreens
    If frm!subLines.Form.Dirty Then frm!subLines.Form.Dirty = False
    If IsNull(frm!cboSupplier.Value) Then
        ShowWarning "«Œ — «·„Ê—œ."
        SafeFocus frm!cboSupplier
        Exit Function
    End If
    If frm!cboPaymentType.Value = "CREDIT" Then paidNow = Nz(frm!txtPaid.Value, 0) Else paidNow = Null
    msg = PostPurchaseFromCart(frm!cboSupplier.Value, Nz(frm!txtSupplierInvoiceNo.Value, ""), _
                               frm!txtInvoiceDate.Value, Nz(frm!cboPaymentType.Value, "CASH"), _
                               frm!cboPaymentMethod.Value, Nz(frm!chkChargeVAT.Value, True), _
                               Nz(frm!txtInvoiceDiscount.Value, 0), paidNow, Nz(frm!txtNotes.Value, ""), newID)
    If Len(msg) > 0 Then
        SetPurStatus frm, msg, CLR_DANGER
        ShowWarning msg
        Exit Function
    End If
    invNo = DLookup("InvoiceNumber", "PurchaseInvoices", "PurchaseInvoiceID = " & newID)
    TempVars("LastPurchaseID") = newID
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    ResetPurchaseHeader frm
    RecalcPurchase frm
    SetPurStatus frm, " „ Õ›Ÿ ›« Ê—… «·‘—«¡ " & invNo & " Ê ÕœÌÀ «·„Œ“Ê‰ Ê„ Ê”ÿ «· ﬂ·›…", CLR_SUCCESS
    SafeFocus frm!cboSupplier
    SavePurchase = True
End Function

Public Sub NewPurchase(ByVal frm As Access.Form)
    If DCount("*", "tmpPurchaseLines") > 0 Then
        If Not AskYesNo("Â·  —Ìœ ≈·€«¡ «·›« Ê—… «·Õ«·Ì… Ê«·»œ¡ »›« Ê—… ÃœÌœ…ø") Then Exit Sub
    End If
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    ResetPurchaseHeader frm
    RecalcPurchase frm
    SetPurStatus frm, "›« Ê—… ÃœÌœ…", CLR_MUTED
    SafeFocus frm!cboSupplier
End Sub

Public Function PurchaseUnload(ByVal frm As Access.Form) As Boolean
    PurchaseUnload = True
    If DCount("*", "tmpPurchaseLines") > 0 Then
        PurchaseUnload = AskYesNo("›« Ê—… «·„‘ —Ì«  «·Õ«·Ì… ·„  ıÕ›Ÿ (” »ﬁÏ „Õ›ÊŸ… „ƒﬁ « ·Â–« «·ÃÂ«“). Â·  —Ìœ «·Œ—ÊÃø")
    End If
End Function

Public Sub PurShowLast(ByVal frm As Access.Form)
    Dim id As Variant
    On Error Resume Next
    id = TempVars("LastPurchaseID")
    On Error GoTo 0
    If IsNull(id) Or IsEmpty(id) Then id = DMax("PurchaseInvoiceID", "PurchaseInvoices", "EmployeeID = " & CurrentUserID())
    If IsNull(id) Then
        ShowInfo "·«  ÊÃœ ›« Ê—… ‘—«¡ ”«»ﬁ…."
    Else
        OpenScreen "frmPurchaseView", 7, CLng(id)
    End If
End Sub

Private Sub SetPurStatus(ByVal frm As Access.Form, ByVal Text As String, ByVal Color As Long)
    frm!lblStatus.Caption = IIf(Len(Text) = 0, " ", Text)
    frm!lblStatus.ForeColor = Color
End Sub

Private Sub SplitQtyPrefix(ByRef Code As String, ByRef Qty As Currency)
    ' "3*6281234567890" -> Code = "6281234567890", Qty = 3
    Dim p As Long
    p = InStr(Code, "*")
    If p > 1 Then
        If IsNumeric(Left$(Code, p - 1)) Then Qty = CCur(Left$(Code, p - 1))
        Code = Mid$(Code, p + 1)
    End If
    If Qty <= 0 Then Qty = 1
End Sub

'==============================================================================
' Purchase return
'==============================================================================
Public Sub PurReturnLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    EnsureLocalTables
    CurrentDb.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError
    frm!subReturnLines.Form.Requery
    frm!txtInvoiceID.Value = Null
    frm!cboRefundType.Value = "CREDIT"
    frm!cboPaymentMethod.Value = 1
    If Not IsNull(frm.OpenArgs) Then LoadPurchaseForReturn frm, CLng(frm.OpenArgs)
End Sub

Public Sub PurReturnFind(ByVal frm As Access.Form)
    Dim id As Variant, typed As String
    typed = SqlText(Trim$(Nz(frm!txtInvoiceNo.Value, "")))
    id = DMax("PurchaseInvoiceID", "PurchaseInvoices", "InvoiceNumber = " & typed & " OR SupplierInvoiceNo = " & typed)
    If IsNull(id) Then
        ShowWarning "·«  ÊÃœ ›« Ê—… ‘—«¡ »Â–« «·—ﬁ„."
        Exit Sub
    End If
    LoadPurchaseForReturn frm, CLng(id)
End Sub

Public Sub LoadPurchaseForReturn(ByVal frm As Access.Form, ByVal PurchaseInvoiceID As Long)
    Dim rs As DAO.Recordset
    CurrentDb.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpPurchaseReturnLines (PurchaseDetailID, ProductID, ProductName, BoughtQty, " & _
        "ReturnedQty, AvailableQty, InStock, ReturnQty, ReturnAmount) " & _
        "SELECT d.PurchaseDetailID, d.ProductID, p.ProductName, d.Quantity, CCur(Nz(r.QtyReturned, 0)), " & _
        "d.Quantity - CCur(Nz(r.QtyReturned, 0)), p.CurrentQuantity, 0, 0 " & _
        "FROM (PurchaseInvoiceDetails AS d INNER JOIN Products AS p ON d.ProductID = p.ProductID) " & _
        "LEFT JOIN qryPurchaseReturnedQty AS r ON d.PurchaseDetailID = r.PurchaseDetailID " & _
        "WHERE d.PurchaseInvoiceID = " & PurchaseInvoiceID, dbFailOnError
    Set rs = CurrentDb.OpenRecordset("SELECT h.*, s.SupplierName FROM PurchaseInvoices AS h INNER JOIN Suppliers " & _
                                     "AS s ON h.SupplierID = s.SupplierID WHERE h.PurchaseInvoiceID = " & _
                                     PurchaseInvoiceID, dbOpenSnapshot)
    frm!txtInvoiceID.Value = PurchaseInvoiceID
    frm!txtInvoiceNo.Value = rs!InvoiceNumber
    frm!lblInvoiceInfo.Caption = "«· «—ÌŒ: " & Format$(rs!InvoiceDate, "yyyy/mm/dd") & "    «·„Ê—œ: " & _
        rs!SupplierName & "    ›« Ê—… «·„Ê—œ: " & Nz(rs!SupplierInvoiceNo, "-") & "    «·≈Ã„«·Ì: " & _
        Format$(rs!TotalAmount, "#,##0.00") & "    " & IIf(rs!PaymentType = "CREDIT", "¬Ã·", "‰ﬁœÌ")
    ' still owed on the invoice -> deduct from the account; paid in cash -> cash refund by default
    frm!cboRefundType.Value = IIf(rs!RemainingAmount > 0, "CREDIT", "CASH")
    rs.Close
    frm!subReturnLines.Form.Requery
    RecalcPurchaseReturn frm
End Sub

Public Sub PurReturnLineChanged(ByVal sf As Access.Form)
    If Nz(sf!ReturnQty.Value, 0) < 0 Then sf!ReturnQty.Value = 0
    If Nz(sf!ReturnQty.Value, 0) > Nz(sf!AvailableQty.Value, 0) Then
        ShowWarning "«·ﬂ„Ì… «·„— Ã⁄… √ﬂ»— „‰ «·„ «Õ ··≈—Ã«⁄ (" & sf!AvailableQty.Value & ")."
        sf!ReturnQty.Value = sf!AvailableQty.Value
    End If
    If sf.Dirty Then sf.Dirty = False
    RecalcPurchaseReturn sf.Parent
End Sub

Public Sub PurReturnAll(ByVal frm As Access.Form)
    CurrentDb.Execute "UPDATE tmpPurchaseReturnLines SET ReturnQty = AvailableQty", dbFailOnError
    RecalcPurchaseReturn frm
End Sub

Public Sub RecalcPurchaseReturn(ByVal frm As Access.Form)
    Dim rs As DAO.Recordset, d As DAO.Recordset, total As Currency, net As Currency, tax As Currency
    Dim lineTotal As Currency, msg As String
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tmpPurchaseReturnLines", dbOpenDynaset)
    Do Until rs.EOF
        lineTotal = 0
        If Nz(rs!ReturnQty, 0) > 0 Then
            Set d = CurrentDb.OpenRecordset("SELECT d.Quantity, d.LineTotal, d.Tax, CCur(Nz(r.QtyReturned, 0)) AS PQ, " & _
                "(SELECT CCur(Nz(Sum(LineTotal), 0)) FROM PurchaseReturnDetails WHERE PurchaseDetailID = d.PurchaseDetailID) AS PT, " & _
                "(SELECT CCur(Nz(Sum(Tax), 0)) FROM PurchaseReturnDetails WHERE PurchaseDetailID = d.PurchaseDetailID) AS PX " & _
                "FROM PurchaseInvoiceDetails AS d LEFT JOIN qryPurchaseReturnedQty AS r " & _
                "ON d.PurchaseDetailID = r.PurchaseDetailID WHERE d.PurchaseDetailID = " & rs!PurchaseDetailID, dbOpenSnapshot)
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

Public Function SavePurchaseReturn(ByVal frm As Access.Form, Optional ByVal PrintAfter As Boolean = False) As Boolean
    Dim msg As String, newID As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Function      ' frmUserScreens
    If IsNull(frm!txtInvoiceID.Value) Then
        ShowWarning "«Œ — ›« Ê—… «·‘—«¡ √Ê·«."
        Exit Function
    End If
    If frm!subReturnLines.Form.Dirty Then frm!subReturnLines.Form.Dirty = False
    msg = PostPurchaseReturn(frm!txtInvoiceID.Value, Nz(frm!txtReason.Value, ""), Nz(frm!cboRefundType.Value, "CREDIT"), _
                             frm!cboPaymentMethod.Value, newID)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    ShowInfo " „ Õ›Ÿ „— Ã⁄ «·„‘ —Ì«  " & DLookup("ReturnNumber", "PurchaseReturns", "PurchaseReturnID = " & newID) & _
             "  »ﬁÌ„… " & Format$(DLookup("TotalAmount", "PurchaseReturns", "PurchaseReturnID = " & newID), "#,##0.00")
    If PrintAfter Then PrintPurchaseDocument "RETURN", newID
    LoadPurchaseForReturn frm, frm!txtInvoiceID.Value
    frm!txtReason.Value = Null
    SavePurchaseReturn = True
End Function

'==============================================================================
' Payment voucher
'==============================================================================
Public Sub SupplierPaymentLoad(ByVal frm As Access.Form)
    If Not IsNull(frm.OpenArgs) Then frm!cboSupplier.Value = CLng(frm.OpenArgs)
    frm!cboPaymentMethod.Value = 1
    SupplierPaymentChanged frm
End Sub

Public Sub SupplierPaymentChanged(ByVal frm As Access.Form)
    If IsNull(frm!cboSupplier.Value) Then
        frm!lblBalance.Caption = " "
    Else
        frm!lblBalance.Caption = "«·„” Õﬁ ··„Ê—œ: " & Format$(Nz(DLookup("CurrentBalance", "Suppliers", _
                                 "SupplierID = " & frm!cboSupplier.Value), 0), "#,##0.00")
    End If
End Sub

Public Function SaveSupplierPayment(ByVal frm As Access.Form, Optional ByVal PrintAfter As Boolean = False) As Boolean
    Dim msg As String, newID As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Function      ' frmUserScreens
    If IsNull(frm!cboSupplier.Value) Then
        ShowWarning "«Œ — «·„Ê—œ."
        Exit Function
    End If
    msg = PostSupplierPayment(frm!cboSupplier.Value, Nz(frm!txtAmount.Value, 0), Nz(frm!cboPaymentMethod.Value, 1), _
                              Nz(frm!txtNotes.Value, ""), newID)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    ShowInfo " „ Õ›Ÿ ”‰œ «·’—› " & DLookup("PaymentNumber", "SupplierPayments", "PaymentID = " & newID)
    If PrintAfter Then PrintVoucher "PAYMENT", newID
    frm!txtAmount.Value = Null
    frm!txtNotes.Value = Null
    SupplierPaymentChanged frm
    SaveSupplierPayment = True
End Function

'==============================================================================
' Purchase invoice view
'==============================================================================
Public Sub PurchaseViewLoad(ByVal frm As Access.Form)
    Dim rs As DAO.Recordset, id As Long
    Calendar = vbCalGreg
    If IsNull(frm.OpenArgs) Then
        ShowWarning "·„  ıÕœœ «·›« Ê—…."
        Exit Sub
    End If
    id = CLng(frm.OpenArgs)
    Set rs = CurrentDb.OpenRecordset("SELECT h.*, s.SupplierName, e.EmployeeName FROM (PurchaseInvoices AS h " & _
        "INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID) INNER JOIN Employees AS e " & _
        "ON h.EmployeeID = e.EmployeeID WHERE h.PurchaseInvoiceID = " & id, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        ShowWarning "«·›« Ê—… €Ì— „ÊÃÊœ…."
        Exit Sub
    End If
    frm!txtInvoiceID.Value = id
    frm!txtSupplierID.Value = rs!SupplierID
    frm!lblTitle.Caption = "›« Ê—… ‘—«¡ " & rs!InvoiceNumber
    frm!lblHeader.Caption = "«· «—ÌŒ: " & Format$(rs!InvoiceDate, "yyyy/mm/dd") & "    «·„Ê—œ: " & rs!SupplierName & _
        "    ›« Ê—… «·„Ê—œ: " & Nz(rs!SupplierInvoiceNo, "-") & "    «·„ÊŸ›: " & rs!EmployeeName & "    " & _
        IIf(rs!PaymentType = "CREDIT", "¬Ã·", "‰ﬁœÌ")
    frm!lblTotals.Caption = "ﬁ»· «·÷—Ì»…: " & Format$(rs!TaxableAmount, "#,##0.00") & "    «·÷—Ì»…: " & _
        Format$(rs!Tax, "#,##0.00") & "    «·≈Ã„«·Ì: " & Format$(rs!TotalAmount, "#,##0.00") & _
        "    «·„œ›Ê⁄: " & Format$(rs!PaidAmount, "#,##0.00") & "    «·„ »ﬁÌ: " & Format$(rs!RemainingAmount, "#,##0.00")
    rs.Close
    frm!lstLines.RowSource = "SELECT d.LineNumber AS [#], p.ProductName AS [«·’‰›], d.Quantity AS [«·ﬂ„Ì…], " & _
        "d.UnitCost AS [ ﬂ·›… «·ÊÕœ…], d.Discount AS [«·Œ’„], d.Tax AS [«·÷—Ì»…], d.LineTotal AS [«·≈Ã„«·Ì] " & _
        "FROM PurchaseInvoiceDetails AS d INNER JOIN Products AS p ON d.ProductID = p.ProductID " & _
        "WHERE d.PurchaseInvoiceID = " & id & " ORDER BY d.LineNumber"
    frm!lstReturns.RowSource = "SELECT PurchaseReturnID, ReturnNumber AS [«·„— Ã⁄], ReturnDate AS [«· «—ÌŒ], " & _
        "TotalAmount AS [«·ﬁÌ„…], Reason AS [«·”»»] FROM PurchaseReturns WHERE PurchaseInvoiceID = " & id & _
        " ORDER BY ReturnDate"
End Sub

'==============================================================================
' Inventory
'==============================================================================
Public Sub InventoryLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    frm!cboMoveType.Value = TT_STOCK_IN
    InventoryRefresh frm
    InventoryMoveTypeChanged frm
    SafeFocus frm!txtSearch
End Sub

Public Sub InventoryRefresh(ByVal frm As Access.Form)
    Dim where As String, rs As DAO.Recordset, keep As Variant
    keep = frm!lstProducts.Value
    where = "p.IsActive = True"
    If Len(Trim$(Nz(frm!txtSearch.Value, ""))) > 0 Then
        where = where & " AND (p.ProductName Like " & LikePattern(Trim$(frm!txtSearch.Value)) & _
                " OR p.ProductCode Like " & LikePattern(Trim$(frm!txtSearch.Value)) & _
                " OR p.Barcode Like " & LikePattern(Trim$(frm!txtSearch.Value)) & ")"
    End If
    If Not IsNull(frm!cboCategory.Value) Then where = where & " AND p.CategoryID = " & CLng(frm!cboCategory.Value)
    If Nz(frm!chkLowOnly.Value, False) Then where = where & " AND p.CurrentQuantity <= p.MinimumQuantity"
    frm!lstProducts.RowSource = "SELECT p.ProductID, p.ProductCode AS [«·ﬂÊœ], p.ProductName AS [«·„‰ Ã], " & _
        "c.CategoryName AS [«· ’‰Ì›], p.CurrentQuantity AS [«·ﬂ„Ì…], p.MinimumQuantity AS [«·Õœ «·√œ‰Ï], " & _
        "p.AverageCost AS [„ Ê”ÿ «· ﬂ·›…], CCur(p.CurrentQuantity * p.AverageCost) AS [ﬁÌ„… «·„Œ“Ê‰] " & _
        "FROM Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID " & _
        "WHERE " & where & " ORDER BY p.ProductName"
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) AS N, Sum(p.CurrentQuantity * p.AverageCost) AS V, " & _
        "Sum(IIf(p.CurrentQuantity <= p.MinimumQuantity, 1, 0)) AS L FROM Products AS p WHERE " & where, dbOpenSnapshot)
    frm!lblInvTotals.Caption = "⁄œœ «·„‰ Ã« : " & rs!N & "    ﬁÌ„… «·„Œ“Ê‰ »«· ﬂ·›…: " & _
        Format$(Nz(rs!V, 0), "#,##0.00") & "    „‰Œ›÷… «·„Œ“Ê‰: " & Nz(rs!L, 0)
    rs.Close
    If Not IsNull(keep) Then
        frm!lstProducts.Value = keep
        InventoryProductPicked frm
    End If
End Sub

Public Sub InventoryProductPicked(ByVal frm As Access.Form)
    Dim rs As DAO.Recordset, id As Variant
    id = frm!lstProducts.Value
    If IsNull(id) Then Exit Sub
    Set rs = CurrentDb.OpenRecordset("SELECT ProductName, ProductCode, CurrentQuantity, MinimumQuantity, AverageCost, " & _
                                     "PurchasePrice FROM Products WHERE ProductID = " & id, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If
    frm!lblProductName.Caption = rs!ProductName & "  (" & rs!ProductCode & ")"
    frm!lblProductStock.Caption = "«·ﬂ„Ì…: " & Format$(rs!CurrentQuantity, "#,##0.###") & "    „ Ê”ÿ «· ﬂ·›…: " & _
        Format$(rs!AverageCost, "#,##0.00##") & "    «·ﬁÌ„…: " & _
        Format$(RoundMoney(CDec(rs!CurrentQuantity) * CDec(rs!AverageCost)), "#,##0.00") & _
        IIf(rs!CurrentQuantity <= rs!MinimumQuantity, "    („‰Œ›÷: «·Õœ " & rs!MinimumQuantity & ")", "")
    frm!lblProductStock.ForeColor = IIf(rs!CurrentQuantity <= rs!MinimumQuantity, CLR_DANGER, CLR_TEXT)
    If Nz(frm!cboMoveType.Value, 0) <> TT_STOCK_OUT Then
        frm!txtMoveCost.Value = IIf(rs!AverageCost > 0, rs!AverageCost, rs!PurchasePrice)
    Else
        frm!txtMoveCost.Value = rs!AverageCost
    End If
    rs.Close
    frm!lstMoves.RowSource = "SELECT TOP 50 t.TransactionDate AS [«· «—ÌŒ], tt.TypeName AS [«·‰Ê⁄], " & _
        "t.Quantity AS [«·ﬂ„Ì…], t.QuantityAfter AS [«·—’Ìœ »⁄œÂ«], t.ReferenceNumber AS [«·„” ‰œ] " & _
        "FROM InventoryTransactions AS t INNER JOIN TransactionTypes AS tt " & _
        "ON t.TransactionTypeID = tt.TransactionTypeID WHERE t.ProductID = " & id & _
        " ORDER BY t.TransactionID DESC"
End Sub

Public Sub InventoryMoveTypeChanged(ByVal frm As Access.Form)
    ' Stock out always leaves at the average cost, so the cost box is read-only for it.
    frm!txtMoveCost.Locked = (Nz(frm!cboMoveType.Value, 0) = TT_STOCK_OUT)
    frm!txtMoveCost.BackColor = IIf(frm!txtMoveCost.Locked, CLR_LOCKED, CLR_SURFACE)
    If Not IsNull(frm!lstProducts.Value) Then InventoryProductPicked frm
End Sub

Public Function PostInventoryMove(ByVal frm As Access.Form) As Boolean
    Dim msg As String, refNo As String
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Function      ' frmUserScreens
    If IsNull(frm!lstProducts.Value) Then
        ShowWarning "«Œ — «·„‰ Ã „‰ «·ﬁ«∆„… √Ê·«."
        Exit Function
    End If
    If IsNull(frm!cboMoveType.Value) Then
        ShowWarning "«Œ — ‰Ê⁄ «·Õ—ﬂ…."
        Exit Function
    End If
    msg = PostManualStock(frm!lstProducts.Value, frm!cboMoveType.Value, Nz(frm!txtMoveQty.Value, 0), _
                          frm!txtMoveCost.Value, Nz(frm!txtMoveNotes.Value, ""), refNo)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    ShowInfo " „ Õ›Ÿ «·Õ—ﬂ… " & refNo
    frm!txtMoveQty.Value = Null
    frm!txtMoveNotes.Value = Null
    InventoryRefresh frm
    PostInventoryMove = True
End Function

'==============================================================================
' Stocktaking
'==============================================================================
Public Sub StockCountLoad(ByVal frm As Access.Form)
    Dim id As Variant
    Calendar = vbCalGreg
    frm!cboCount.Requery
    id = DbValue("SELECT StockCountID FROM StockCounts WHERE Status = 'OPEN'")
    If IsNull(id) Then id = DMax("StockCountID", "StockCounts")
    If Not IsNull(frm.OpenArgs) Then id = CLng(frm.OpenArgs)
    frm!cboCount.Value = id
    ShowStockCount frm
End Sub

Public Sub StockCountPicked(ByVal frm As Access.Form)
    ShowStockCount frm
End Sub

Public Sub CountFilterChanged(ByVal frm As Access.Form)
    ShowStockCount frm
End Sub

Private Sub ShowStockCount(ByVal frm As Access.Form)
    Dim rs As DAO.Recordset, id As Long, isOpen As Boolean, sql As String
    id = Nz(frm!cboCount.Value, 0)
    sql = COUNT_LINES_SQL & " WHERE d.StockCountID = " & id
    If Nz(frm!chkDiffOnly.Value, False) Then sql = sql & " AND d.ActualQuantity Is Not Null AND d.Difference <> 0"
    frm!subCountLines.Form.RecordSource = sql & " ORDER BY p.ProductName"
    Set rs = CurrentDb.OpenRecordset("SELECT c.*, g.CategoryName FROM StockCounts AS c LEFT JOIN Categories AS g " & _
                                     "ON c.CategoryID = g.CategoryID WHERE c.StockCountID = " & id, dbOpenSnapshot)
    If rs.EOF Then
        frm!lblCountInfo.Caption = "·« ÌÊÃœ Ã—œ. «Œ — «· ’‰Ì› (√Ê « —ﬂÂ ··ﬂ·) À„ «÷€ÿ ´Ã—œ ÃœÌœª."
    Else
        isOpen = (rs!Status = "OPEN")
        frm!lblCountInfo.Caption = rs!CountNumber & "    " & Format$(rs!CountDate, "yyyy/mm/dd") & "    " & _
            IIf(isOpen, "„› ÊÕ", IIf(rs!Status = "POSTED", "„ı—Õ¯· " & Format$(Nz(rs!PostedAt, rs!CountDate), _
            "yyyy/mm/dd"), "„·€Ï")) & "    «· ’‰Ì›: " & Nz(rs!CategoryName, "«·ﬂ·")
    End If
    rs.Close
    frm!subCountLines.Form.AllowEdits = isOpen
    SafeFocus frm!cboCount               ' a focused control cannot be disabled
    frm!txtCountBarcode.Enabled = isOpen
    frm!btnPostCount.Enabled = isOpen
    frm!btnCancelCount.Enabled = isOpen
    frm!btnRefreshSystem.Enabled = isOpen
    RefreshCountSummary frm
End Sub

Private Sub RefreshCountSummary(ByVal frm As Access.Form)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) AS N, Sum(IIf(ActualQuantity Is Null, 0, 1)) AS C, " & _
        "Sum(IIf(DifferenceValue < 0, DifferenceValue, 0)) AS S, Sum(IIf(DifferenceValue > 0, DifferenceValue, 0)) AS P " & _
        "FROM StockCountDetails WHERE StockCountID = " & Nz(frm!cboCount.Value, 0), dbOpenSnapshot)
    If Nz(rs!N, 0) = 0 Then
        frm!lblCountSummary.Caption = " "
    Else
        frm!lblCountSummary.Caption = " „ ⁄œ¯ " & Nz(rs!C, 0) & " „‰ " & rs!N & " ’‰›    «·⁄Ã“: " & _
            Format$(-Nz(rs!S, 0), "#,##0.00") & "    «·“Ì«œ…: " & Format$(Nz(rs!P, 0), "#,##0.00") & _
            "    «·’«›Ì: " & Format$(Nz(rs!S, 0) + Nz(rs!P, 0), "#,##0.00")
    End If
    rs.Close
End Sub

Public Function NewStockCount(ByVal frm As Access.Form) As Boolean
    Dim msg As String, newID As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Function      ' frmUserScreens
    If Not AskYesNo("»œ¡ Ã—œ ÃœÌœ " & IIf(IsNull(frm!cboCategory.Value), "·ﬂ· «·„‰ Ã«  «·‰‘ÿ…", _
                    "· ’‰Ì› ´" & frm!cboCategory.Column(1) & "ª") & "ø" & vbCrLf & _
                    "Ìı›÷Û¯· «·Ã—œ Ê«·„Õ· „€·ﬁ √Ê »⁄œ ¬Œ— ›« Ê—…° À„ «· —ÕÌ· „»«‘—….") Then Exit Function
    msg = CreateStockCount(frm!cboCategory.Value, "", newID)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    frm!cboCount.Requery
    frm!cboCount.Value = newID
    ShowStockCount frm
    SafeFocus frm!txtCountBarcode
    NewStockCount = True
End Function

Public Sub CountLineChanged(ByVal sf As Access.Form)
    ' AfterUpdate of ActualQuantity in frmStockCountLines.
    Dim diff As Currency
    If Not IsNull(sf!ActualQuantity.Value) Then
        If sf!ActualQuantity.Value < 0 Then
            ShowWarning "«·ﬂ„Ì… «·›⁄·Ì… ·« Ì„ﬂ‰ √‰  ﬂÊ‰ ”«·»…."
            sf!ActualQuantity.Value = Null
        End If
    End If
    If IsNull(sf!ActualQuantity.Value) Then
        diff = 0
    Else
        diff = sf!ActualQuantity.Value - Nz(sf!SystemQuantity.Value, 0)
    End If
    sf!Difference.Value = diff
    sf!DifferenceValue.Value = RoundMoney(CDec(diff) * CDec(Nz(sf!UnitCost.Value, 0)))
    If sf.Dirty Then sf.Dirty = False
    RefreshCountSummary sf.Parent
End Sub

Public Sub CountBarcodeKeyDown(ByVal frm As Access.Form, ByRef KeyCode As Integer)
    Dim code As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    KeyCode = 0
    code = Trim$(frm!txtCountBarcode.Text)
    frm!txtCountBarcode.Text = ""
    If Len(code) > 0 Then CountScanCode frm, code
End Sub

Public Function CountScanCode(ByVal frm As Access.Form, ByVal Code As String) As Boolean
    ' Each scan adds 1 (or n with "n*code") to the counted quantity of that product.
    Dim qty As Currency, id As Long, rs As DAO.Recordset, detailID As Variant, diff As Currency
    id = Nz(frm!cboCount.Value, 0)
    If Nz(DbValue("SELECT Status FROM StockCounts WHERE StockCountID = " & id), "") <> "OPEN" Then
        ShowWarning "«Œ — Ã—œ« „› ÊÕ« √Ê·«."
        Exit Function
    End If
    qty = 1
    SplitQtyPrefix Code, qty
    detailID = DbValue("SELECT d.StockCountDetailID FROM StockCountDetails AS d INNER JOIN Products AS p " & _
                       "ON d.ProductID = p.ProductID WHERE d.StockCountID = " & id & " AND (p.Barcode = " & _
                       SqlText(Code) & " OR p.ProductCode = " & SqlText(Code) & ")")
    If IsNull(detailID) Then
        Beep
        If IsNull(DLookup("ProductID", "Products", "Barcode = " & SqlText(Code) & " OR ProductCode = " & SqlText(Code))) Then
            ShowWarning "·„ Ì „ «·⁄ÀÊ— ⁄·Ï „‰ Ã »«·»«—ﬂÊœ √Ê «·ﬂÊœ: " & Code
        Else
            ShowWarning "«·„‰ Ã „ÊÃÊœ ·ﬂ‰Â ·Ì” ÷„‰ Â–« «·Ã—œ ( ’‰Ì› ¬Œ— √Ê €Ì— ‰‘ÿ)."
        End If
        Exit Function
    End If
    If frm!subCountLines.Form.Dirty Then frm!subCountLines.Form.Dirty = False
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM StockCountDetails WHERE StockCountDetailID = " & detailID, dbOpenDynaset)
    rs.Edit
    rs!ActualQuantity = Nz(rs!ActualQuantity, 0) + qty
    diff = rs!ActualQuantity - rs!SystemQuantity
    rs!Difference = diff
    rs!DifferenceValue = RoundMoney(CDec(diff) * CDec(rs!UnitCost))
    rs.Update
    rs.Close
    frm!subCountLines.Form.Requery
    On Error Resume Next                 ' positioning is a convenience only
    frm!subCountLines.Form.Recordset.FindFirst "StockCountDetailID = " & detailID
    On Error GoTo 0
    RefreshCountSummary frm
    CountScanCode = True
End Function

Public Sub CountRefreshSystem(ByVal frm As Access.Form)
    ' Refreshes the recorded quantities and costs with the current stock (sales made while
    ' counting) and recomputes the differences. Posting does the same automatically.
    Dim rs As DAO.Recordset, diff As Currency
    If frm!subCountLines.Form.Dirty Then frm!subCountLines.Form.Dirty = False
    Set rs = CurrentDb.OpenRecordset("SELECT d.SystemQuantity, d.ActualQuantity, d.Difference, d.UnitCost, " & _
        "d.DifferenceValue, p.CurrentQuantity, p.AverageCost FROM StockCountDetails AS d INNER JOIN Products AS p " & _
        "ON d.ProductID = p.ProductID WHERE d.StockCountID = " & Nz(frm!cboCount.Value, 0), dbOpenDynaset)
    Do Until rs.EOF
        If IsNull(rs!ActualQuantity) Then diff = 0 Else diff = rs!ActualQuantity - rs!CurrentQuantity
        rs.Edit
        rs!SystemQuantity = rs!CurrentQuantity
        rs!UnitCost = rs!AverageCost
        rs!Difference = diff
        rs!DifferenceValue = RoundMoney(CDec(diff) * CDec(rs!AverageCost))
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    frm!subCountLines.Form.Requery
    RefreshCountSummary frm
End Sub

Public Function PostCountScreen(ByVal frm As Access.Form) As Boolean
    Dim msg As String, id As Long, uncounted As Long, lines As Long, value As Currency
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Function      ' frmUserScreens
    id = Nz(frm!cboCount.Value, 0)
    If frm!subCountLines.Form.Dirty Then frm!subCountLines.Form.Dirty = False
    uncounted = Nz(DbValue("SELECT COUNT(*) FROM StockCountDetails WHERE StockCountID = " & id & _
                           " AND ActualQuantity Is Null"), 0)
    If Not AskYesNo(" —ÕÌ· «·Ã—œ ÌÕÊ¯· «·›—Êﬁ«  ≈·Ï Õ—ﬂ«   ”ÊÌ… ›Ì «·„Œ“Ê‰ Ê·« Ì„ﬂ‰ «· —«Ã⁄ ⁄‰Â." & vbCrLf & _
                    " ıÕ”» «·›—Êﬁ«  ⁄·Ï «·ﬂ„Ì«  «·Õ«·Ì… ·ÕŸ… «· —ÕÌ·." & vbCrLf & _
                    IIf(uncounted > 0, "«·√’‰«› «· Ì ·„  ı⁄œ¯ (" & uncounted & ")  »ﬁÏ ﬂ„« ÂÌ∫ " & _
                    "· ’›Ì— ’‰› €Ì— „ÊÃÊœ «ﬂ » 0 ›Ì ﬂ„Ì Â «·›⁄·Ì…." & vbCrLf, "") & vbCrLf & "Â·  —Ìœ «· —ÕÌ·ø") Then
        Exit Function
    End If
    msg = PostStockCount(id, False, lines, value)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    ShowInfo " „  —ÕÌ· «·Ã—œ: " & lines & "  ”ÊÌ…° ’«›Ì ﬁÌ„… «·›—Êﬁ«  " & Format$(value, "#,##0.00")
    frm!cboCount.Requery
    ShowStockCount frm
    PostCountScreen = True
End Function

Public Function CancelCountScreen(ByVal frm As Access.Form) As Boolean
    Dim msg As String
    If Not AskYesNo("≈·€«¡ «·Ã—œ «·Õ«·Ìø ·‰   €Ì— √Ì ﬂ„Ì… ›Ì «·„Œ“Ê‰.") Then Exit Function
    msg = CancelStockCount(Nz(frm!cboCount.Value, 0))
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    frm!cboCount.Requery
    ShowStockCount frm
    CancelCountScreen = True
End Function
