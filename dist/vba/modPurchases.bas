Attribute VB_Name = "modPurchases"
'==============================================================================
' modPurchases  -  Retail Store Management System (Phase 7)
'
' Posting of the purchase and inventory documents, each in ONE transaction:
'   PostPurchaseFromCart   purchase invoice from tmpPurchaseLines
'   PostPurchaseReturn     goods returned to the supplier from tmpPurchaseReturnLines
'   PostSupplierPayment    payment voucher (”‰œ ’—›)
'   PostManualStock        opening balance / stock in / stock out
'   CreateStockCount       opens a stocktake with the recorded quantities
'   PostStockCount         turns the counted differences into ADJUSTMENT movements
'   CancelStockCount       cancels an open stocktake
' Functions return "" on success or the reason in Arabic.
'
' Supplier account (same as qrySupplierLedger, balance = Credit - Debit):
'   purchase  Credit = Total,  Debit = Paid       -> balance + Remaining
'   return    Debit  = Total,  Credit = Refunded  -> balance + (Refunded - Total)
'   payment   Debit  = Amount                     -> balance - Amount
'==============================================================================
Option Compare Database
Option Explicit

'------------------------------------------------------------------------------
' Purchase invoice
'------------------------------------------------------------------------------
Public Function PostPurchaseFromCart(ByVal SupplierID As Long, ByVal SupplierInvoiceNo As String, _
                                     ByVal InvoiceDate As Variant, ByVal PaymentType As String, _
                                     ByVal PaymentMethodID As Variant, ByVal ChargeVat As Boolean, _
                                     ByVal InvoiceDiscount As Currency, ByVal PaidAmount As Variant, _
                                     ByVal Notes As String, ByRef NewInvoiceID As Long) As String
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean
    Dim n As Long, i As Long, msg As String, vatRate As Currency, rate As Currency
    Dim productIDs() As Long, newPrices() As Variant, isCredit As Boolean, reclaim As Boolean
    Dim tend As Currency, paid As Currency, remaining As Currency, change As Currency
    Dim invoiceID As Long, invNo As String, docDate As Date, unitCost As Currency, priceChanges As String

    On Error GoTo EH
    NewInvoiceID = 0
    Calendar = vbCalGreg
    Set db = CurrentDb
    If Not HasPermission("PURCHASES") Then
        PostPurchaseFromCart = "·«  „·ﬂ ’·«ÕÌ…  ”ÃÌ· ›Ê« Ì— «·„‘ —Ì« ."
        Exit Function
    End If
    msg = CheckSupplier(SupplierID)
    If Len(msg) = 0 Then msg = PurchaseDocDate(InvoiceDate, docDate)
    If Len(msg) = 0 And Len(Trim$(SupplierInvoiceNo)) > 0 Then
        If Not IsNull(DbValue("SELECT InvoiceNumber FROM PurchaseInvoices WHERE SupplierID = " & SupplierID & _
                              " AND SupplierInvoiceNo = " & SqlText(Trim$(SupplierInvoiceNo)))) Then
            msg = "›« Ê—… «·„Ê—œ —ﬁ„ " & Trim$(SupplierInvoiceNo) & " „”Ã·… „”»ﬁ« ·‰›” «·„Ê—œ."
        End If
    End If
    If Len(msg) > 0 Then
        PostPurchaseFromCart = msg
        Exit Function
    End If
    vatRate = Nz(SettingValue("VATRate"), 0.15)
    reclaim = Len(Nz(SettingValue("VATNumber"), "")) > 0

    ' ---- 1. cart -> calculation engine (purchase prices are typed WITHOUT VAT)
    Set rs = db.OpenRecordset( _
        "SELECT t.ProductID, t.Quantity, t.UnitCost, t.LineDiscount, t.NewSellingPrice, p.ProductName, " & _
        "p.VATCategory, p.SellingPrice, p.IsActive " & _
        "FROM tmpPurchaseLines AS t LEFT JOIN Products AS p ON t.ProductID = p.ProductID " & _
        "ORDER BY t.LineNo", dbOpenSnapshot)
    CalcReset
    Do Until rs.EOF
        If IsNull(rs!ProductName) Then
            msg = "√Õœ «·√’‰«› ›Ì «·›« Ê—… ·„ Ì⁄œ „ÊÃÊœ«. «Õ–›Â „‰ «·›« Ê—…."
        ElseIf Not rs!IsActive Then
            msg = "«·„‰ Ã ´" & rs!ProductName & "ª €Ì— ‰‘ÿ. ›⁄¯·Â „‰ ‘«‘… «·„‰ Ã«  √Ê·«."
        ElseIf Not IsNull(rs!NewSellingPrice) Then
            If rs!NewSellingPrice < 0 Then
                msg = "”⁄— «·»Ì⁄ «·ÃœÌœ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»« (´" & rs!ProductName & "ª)."
            ElseIf rs!NewSellingPrice <> rs!SellingPrice Then
                priceChanges = priceChanges & " " & rs!ProductID & ":" & rs!SellingPrice & "->" & rs!NewSellingPrice
            End If
        End If
        If Len(msg) > 0 Then
            rs.Close
            PostPurchaseFromCart = msg
            Exit Function
        End If
        If ChargeVat And Nz(rs!VATCategory, "S") = "S" Then rate = vatRate Else rate = 0
        CalcAddLine Nz(rs!Quantity, 0), Nz(rs!UnitCost, 0), Nz(rs!LineDiscount, 0), rate
        ReDim Preserve productIDs(0 To n)
        ReDim Preserve newPrices(0 To n)
        productIDs(n) = rs!ProductID
        newPrices(n) = rs!NewSellingPrice
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    msg = CalcRun(InvoiceDiscount, False)
    If Len(msg) = 0 And Len(priceChanges) > 0 And Not HasPermission("PRODUCTS") Then
        msg = "·«  „·ﬂ ’·«ÕÌ…  ⁄œÌ· √”⁄«— «·»Ì⁄. «„”Õ ⁄„Êœ ´”⁄— «·»Ì⁄ «·ÃœÌœª."
    End If
    If Len(msg) > 0 Then
        PostPurchaseFromCart = msg
        Exit Function
    End If

    ' ---- 2. payment: cash = paid in full; credit = whatever was paid now (default 0)
    isCredit = (PaymentType = "CREDIT")
    If IsNull(PaidAmount) Then
        If isCredit Then tend = 0 Else tend = CalcTotal("TOTAL")
    Else
        tend = CCur(PaidAmount)
    End If
    msg = Settle(CalcTotal("TOTAL"), tend, isCredit, paid, remaining, change)
    If Len(msg) > 0 Then
        PostPurchaseFromCart = msg
        Exit Function
    End If

    ' ---- 3. post everything in one transaction
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    invNo = NextNumber("PURCHASE_INVOICE")
    Set rs = db.OpenRecordset("PurchaseInvoices", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!InvoiceNumber = invNo
    If Len(Trim$(SupplierInvoiceNo)) > 0 Then rs!SupplierInvoiceNo = Left$(Trim$(SupplierInvoiceNo), 30)
    rs!InvoiceDate = docDate
    rs!SupplierID = SupplierID
    rs!EmployeeID = CurrentUserID()
    rs!PaymentType = PaymentType
    rs!PaymentMethodID = PaymentMethodID
    rs!SubTotal = CalcTotal("SUBTOTAL")
    rs!Discount = CalcTotal("DISCOUNT")
    rs!TaxableAmount = CalcTotal("TAXABLE")
    rs!Tax = CalcTotal("TAX")
    rs!TotalAmount = CalcTotal("TOTAL")
    rs!PaidAmount = paid
    rs!RemainingAmount = remaining
    If Len(Notes) > 0 Then rs!Notes = Left$(Notes, 255)
    rs.Update
    rs.Bookmark = rs.LastModified
    invoiceID = rs!PurchaseInvoiceID
    rs.Close

    Set rs = db.OpenRecordset("PurchaseInvoiceDetails", dbOpenDynaset, dbAppendOnly)
    For i = 0 To n - 1
        rs.AddNew
        rs!PurchaseInvoiceID = invoiceID
        rs!LineNumber = i + 1
        rs!ProductID = productIDs(i)
        rs!Quantity = CalcLine(i, "QTY")
        rs!UnitCost = CalcLine(i, "UNIT")
        rs!Discount = CalcLine(i, "DISCOUNT")
        rs!NetAmount = CalcLine(i, "NET")
        rs!VATRate = CalcLine(i, "RATE")
        rs!Tax = CalcLine(i, "TAX")
        rs!LineTotal = CalcLine(i, "TOTAL")
        rs.Update
    Next
    rs.Close

    For i = 0 To n - 1
        unitCost = PurchaseUnitCost(CalcLine(i, "NET"), CalcLine(i, "TOTAL"), CalcLine(i, "QTY"), reclaim)
        ApplyStockMovement db, productIDs(i), CalcLine(i, "QTY"), TT_PURCHASE, unitCost, "PURCHASE", _
                           invoiceID, invNo, "", True
        UpdateProductPrices db, productIDs(i), CalcLine(i, "UNIT"), newPrices(i), SupplierID
    Next
    If remaining <> 0 Then AdjustBalance db, "Suppliers", "SupplierID", SupplierID, remaining

    ws.CommitTrans
    inTrans = False
    LogAction "PURCHASE", "PurchaseInvoices", invNo, "Total=" & CalcTotal("TOTAL")
    If Len(priceChanges) > 0 Then LogAction "PRICE_CHANGE", "Products", invNo, "Selling price:" & priceChanges
    NewInvoiceID = invoiceID
    Exit Function

EH:
    PostPurchaseFromCart = " ⁄–— Õ›Ÿ ›« Ê—… «·‘—«¡: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
End Function

Private Sub UpdateProductPrices(ByVal db As DAO.Database, ByVal ProductID As Long, ByVal LastCost As Currency, _
                                ByVal NewSellingPrice As Variant, ByVal SupplierID As Long)
    ' Last purchase price (excl. VAT), the default supplier when empty, and the new shelf price if typed.
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT ProductID, PurchasePrice, SellingPrice, SupplierID, UpdatedAt FROM Products " & _
                              "WHERE ProductID = " & ProductID, dbOpenDynaset)
    rs.Edit
    rs!PurchasePrice = LastCost
    If IsNull(rs!SupplierID) Then rs!SupplierID = SupplierID
    If Not IsNull(NewSellingPrice) Then rs!SellingPrice = NewSellingPrice
    rs!UpdatedAt = Now
    rs.Update
    rs.Close
End Sub

'------------------------------------------------------------------------------
' Purchase return
'------------------------------------------------------------------------------
Public Function PostPurchaseReturn(ByVal PurchaseInvoiceID As Long, ByVal Reason As String, _
                                   ByVal RefundType As String, ByVal PaymentMethodID As Variant, _
                                   ByRef NewReturnID As Long) As String
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean
    Dim n As Long, i As Long, msg As String, supplierID As Long, reclaim As Boolean
    Dim detailIDs() As Long, productIDs() As Long, qtys() As Currency, units() As Currency
    Dim nets() As Currency, taxes() As Currency, totals() As Currency, rates() As Currency
    Dim discs() As Currency, costs() As Currency
    Dim sumNet As Currency, sumTax As Currency, sumTotal As Currency, sumDisc As Currency
    Dim refunded As Currency, retID As Long, retNo As String

    On Error GoTo EH
    NewReturnID = 0
    Calendar = vbCalGreg
    Set db = CurrentDb
    If Not HasPermission("PURCHASE_RETURN") Then
        PostPurchaseReturn = "·«  „·ﬂ ’·«ÕÌ…  ”ÃÌ· „— Ã⁄«  «·„‘ —Ì« ."
        Exit Function
    End If
    If Len(Trim$(Reason)) = 0 Then
        PostPurchaseReturn = "”»» «·≈—Ã«⁄ „ÿ·Ê»."
        Exit Function
    End If
    If RefundType <> "CASH" And RefundType <> "CREDIT" Then RefundType = "CREDIT"
    supplierID = Nz(DbValue("SELECT SupplierID FROM PurchaseInvoices WHERE PurchaseInvoiceID = " & PurchaseInvoiceID), 0)
    If supplierID = 0 Then
        PostPurchaseReturn = "›« Ê—… «·‘—«¡ «·√’·Ì… €Ì— „ÊÃÊœ…."
        Exit Function
    End If
    reclaim = Len(Nz(SettingValue("VATNumber"), "")) > 0

    Set rs = db.OpenRecordset( _
        "SELECT r.PurchaseDetailID, r.ReturnQty, d.ProductID, d.Quantity, d.UnitCost, d.NetAmount, d.Tax, " & _
        "d.LineTotal, d.VATRate, d.PurchaseInvoiceID " & _
        "FROM tmpPurchaseReturnLines AS r INNER JOIN PurchaseInvoiceDetails AS d " & _
        "ON r.PurchaseDetailID = d.PurchaseDetailID WHERE r.ReturnQty > 0", dbOpenSnapshot)
    Do Until rs.EOF
        If rs!PurchaseInvoiceID <> PurchaseInvoiceID Then
            PostPurchaseReturn = "”ÿ— ·« Ì‰ „Ì ≈·Ï ›« Ê—… «·‘—«¡ «·√’·Ì…."
            rs.Close
            Exit Function
        End If
        ReDim Preserve detailIDs(0 To n): ReDim Preserve productIDs(0 To n)
        ReDim Preserve qtys(0 To n): ReDim Preserve units(0 To n)
        ReDim Preserve nets(0 To n): ReDim Preserve taxes(0 To n)
        ReDim Preserve totals(0 To n): ReDim Preserve rates(0 To n)
        ReDim Preserve discs(0 To n): ReDim Preserve costs(0 To n)
        detailIDs(n) = rs!PurchaseDetailID
        productIDs(n) = rs!ProductID
        qtys(n) = rs!ReturnQty
        units(n) = rs!UnitCost
        rates(n) = rs!VATRate
        ' the goods leave at the cost they came in with
        costs(n) = PurchaseUnitCost(rs!NetAmount, rs!LineTotal, rs!Quantity, reclaim)
        msg = ReturnAmounts(rs!Quantity, rs!LineTotal, rs!Tax, _
                            PrevPurchaseReturned(db, rs!PurchaseDetailID, "Quantity"), _
                            PrevPurchaseReturned(db, rs!PurchaseDetailID, "LineTotal"), _
                            PrevPurchaseReturned(db, rs!PurchaseDetailID, "Tax"), _
                            qtys(n), nets(n), taxes(n), totals(n))
        If Len(msg) > 0 Then
            PostPurchaseReturn = msg
            rs.Close
            Exit Function
        End If
        discs(n) = RoundMoney(CDec(qtys(n)) * CDec(units(n))) - nets(n)
        If discs(n) < 0 Then discs(n) = 0
        sumNet = sumNet + nets(n)
        sumTax = sumTax + taxes(n)
        sumTotal = sumTotal + totals(n)
        sumDisc = sumDisc + discs(n)
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    If n = 0 Then
        PostPurchaseReturn = "Õœœ «·ﬂ„Ì… «·„— Ã⁄… ·’‰› Ê«Õœ ⁄·Ï «·√ﬁ·."
        Exit Function
    End If
    msg = CheckStockAvailable(productIDs, qtys, n, "«·≈—Ã«⁄ ··„Ê—œ")
    If Len(msg) > 0 Then
        PostPurchaseReturn = msg
        Exit Function
    End If
    If RefundType = "CASH" Then refunded = sumTotal Else refunded = 0

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    retNo = NextNumber("PURCHASE_RETURN")
    Set rs = db.OpenRecordset("PurchaseReturns", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!ReturnNumber = retNo
    rs!ReturnDate = Now
    rs!PurchaseInvoiceID = PurchaseInvoiceID
    rs!SupplierID = supplierID
    rs!EmployeeID = CurrentUserID()
    rs!Reason = Left$(Reason, 255)
    rs!RefundType = RefundType
    rs!PaymentMethodID = PaymentMethodID
    rs!SubTotal = sumNet + sumDisc
    rs!Discount = sumDisc
    rs!TaxableAmount = sumNet
    rs!Tax = sumTax
    rs!TotalAmount = sumTotal
    rs!RefundedAmount = refunded
    rs.Update
    rs.Bookmark = rs.LastModified
    retID = rs!PurchaseReturnID
    rs.Close

    Set rs = db.OpenRecordset("PurchaseReturnDetails", dbOpenDynaset, dbAppendOnly)
    For i = 0 To n - 1
        rs.AddNew
        rs!PurchaseReturnID = retID
        rs!PurchaseDetailID = detailIDs(i)
        rs!ProductID = productIDs(i)
        rs!Quantity = qtys(i)
        rs!UnitCost = units(i)
        rs!Discount = discs(i)
        rs!NetAmount = nets(i)
        rs!VATRate = rates(i)
        rs!Tax = taxes(i)
        rs!LineTotal = totals(i)
        rs.Update
    Next
    rs.Close

    For i = 0 To n - 1
        ApplyStockMovement db, productIDs(i), -qtys(i), TT_PURCHASE_RETURN, costs(i), "PURCHASE_RETURN", _
                           retID, retNo, "", True
    Next
    If refunded - sumTotal <> 0 Then AdjustBalance db, "Suppliers", "SupplierID", supplierID, refunded - sumTotal

    ws.CommitTrans
    inTrans = False
    LogAction "PURCHASE_RETURN", "PurchaseReturns", retNo, "Invoice=" & PurchaseInvoiceID & " Total=" & sumTotal
    NewReturnID = retID
    Exit Function

EH:
    PostPurchaseReturn = " ⁄–— Õ›Ÿ „— Ã⁄ «·„‘ —Ì« : " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
End Function

Private Function PrevPurchaseReturned(ByVal db As DAO.Database, ByVal PurchaseDetailID As Long, _
                                      ByVal FieldName As String) As Currency
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Sum(" & FieldName & ") FROM PurchaseReturnDetails WHERE PurchaseDetailID = " & _
                              PurchaseDetailID, dbOpenSnapshot)
    PrevPurchaseReturned = Nz(rs(0), 0)
    rs.Close
End Function

'------------------------------------------------------------------------------
' Payment voucher
'------------------------------------------------------------------------------
Public Function PostSupplierPayment(ByVal SupplierID As Long, ByVal Amount As Currency, _
                                    ByVal PaymentMethodID As Long, ByVal Notes As String, _
                                    ByRef NewPaymentID As Long) As String
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean
    Dim owed As Currency, payNo As String, msg As String

    On Error GoTo EH
    NewPaymentID = 0
    Set db = CurrentDb
    If Not HasPermission("SUPPLIER_PAYMENTS") Then
        PostSupplierPayment = "·«  „·ﬂ ’·«ÕÌ…  ”ÃÌ· ”‰œ«  «·’—›."
        Exit Function
    End If
    If Amount <= 0 Then
        PostSupplierPayment = "«·„»·€ ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—."
        Exit Function
    End If
    msg = CheckSupplier(SupplierID)
    If Len(msg) > 0 Then
        PostSupplierPayment = msg
        Exit Function
    End If
    owed = Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierID = " & SupplierID), 0)
    If Amount > owed Then
        If Not AskYesNo("«·„»·€ (" & Format$(Amount, "#,##0.00") & ") √ﬂ»— „‰ «·„” Õﬁ ··„Ê—œ (" & _
                        Format$(owed, "#,##0.00") & ")° Ê”Ì’»Õ «·„Ê—œ „œÌ‰« ··„Õ·." & vbCrLf & _
                        "Â·  —Ìœ «·„ «»⁄…ø") Then
            PostSupplierPayment = " „ «·≈·€«¡."
            Exit Function
        End If
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    payNo = NextNumber("SUPPLIER_PAYMENT")
    Set rs = db.OpenRecordset("SupplierPayments", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!PaymentNumber = payNo
    rs!SupplierID = SupplierID
    rs!PaymentDate = Now
    rs!Amount = Amount
    rs!PaymentMethodID = PaymentMethodID
    rs!EmployeeID = CurrentUserID()
    If Len(Notes) > 0 Then rs!Notes = Left$(Notes, 255)
    rs.Update
    rs.Bookmark = rs.LastModified
    NewPaymentID = rs!PaymentID
    rs.Close
    AdjustBalance db, "Suppliers", "SupplierID", SupplierID, -Amount
    ws.CommitTrans
    inTrans = False
    LogAction "SUPPLIER_PAYMENT", "SupplierPayments", payNo, "Amount=" & Amount
    Exit Function

EH:
    PostSupplierPayment = " ⁄–— Õ›Ÿ ”‰œ «·’—›: " & Err.Description & " (" & Err.Number & ")"
    NewPaymentID = 0
    If inTrans Then ws.Rollback
End Function

'------------------------------------------------------------------------------
' Manual stock movements: OPENING (once per product), STOCK_IN, STOCK_OUT
'------------------------------------------------------------------------------
Public Function PostManualStock(ByVal ProductID As Long, ByVal TransactionTypeID As Long, _
                                ByVal Qty As Currency, ByVal UnitCost As Variant, ByVal Notes As String, _
                                ByRef RefNumber As String) As String
    Dim db As DAO.Database, ws As DAO.Workspace, inTrans As Boolean, msg As String
    Dim avgCost As Currency, cost As Currency, ids(0 To 0) As Long, qtys(0 To 0) As Currency

    On Error GoTo EH
    RefNumber = ""
    Set db = CurrentDb
    If Not HasPermission("INVENTORY_ADJUST") Then
        PostManualStock = "·«  „·ﬂ ’·«ÕÌ…  ⁄œÌ· «·„Œ“Ê‰ ÌœÊÌ«."
        Exit Function
    End If
    If IsNull(DbValue("SELECT ProductID FROM Products WHERE ProductID = " & ProductID)) Then
        PostManualStock = "«Œ — «·„‰ Ã."
        Exit Function
    End If
    If Qty <= 0 Then
        PostManualStock = "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ √ﬂ»— „‰ ’›—."
        Exit Function
    End If
    avgCost = Nz(DbValue("SELECT AverageCost FROM Products WHERE ProductID = " & ProductID), 0)
    If IsNull(UnitCost) Then cost = avgCost Else cost = CCur(UnitCost)
    If cost < 0 Then
        PostManualStock = "«· ﬂ·›… ·« Ì„ﬂ‰ √‰  ﬂÊ‰ ”«·»…."
        Exit Function
    End If

    Select Case TransactionTypeID
        Case TT_OPENING
            If Nz(DbValue("SELECT COUNT(*) FROM InventoryTransactions WHERE ProductID = " & ProductID), 0) > 0 Then
                msg = "«·—’Ìœ «·«›  «ÕÌ Ìı”ÃÛ¯· „—… Ê«Õœ… ﬁ»· √Ì Õ—ﬂ… ⁄·Ï «·„‰ Ã." & vbCrLf & _
                      "· ’ÕÌÕ «·ﬂ„Ì… «·¬‰ «” Œœ„ «·Ã—œ √Ê «·≈÷«›…/«·Œ’„."
            End If
        Case TT_STOCK_IN, TT_STOCK_OUT
            If Len(Trim$(Notes)) = 0 Then msg = "«ﬂ » ”»» «·Õ—ﬂ… ›Ì «·„·«ÕŸ«  („À«·:  «·›° ÂœÌ… „‰ «·„Ê—œ)."
            If Len(msg) = 0 And TransactionTypeID = TT_STOCK_OUT Then
                cost = avgCost                            ' goods leave at the average cost
                ids(0) = ProductID
                qtys(0) = Qty
                msg = CheckStockAvailable(ids, qtys, 1, "Œ’„ «·„Œ“Ê‰")
            End If
        Case Else
            msg = "‰Ê⁄ «·Õ—ﬂ… €Ì— „ «Õ ··≈œŒ«· «·ÌœÊÌ."
    End Select
    If Len(msg) > 0 Then
        PostManualStock = msg
        Exit Function
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    RefNumber = NextNumber("STOCK_ADJUST")
    If TransactionTypeID = TT_STOCK_OUT Then
        ApplyStockMovement db, ProductID, -Qty, TT_STOCK_OUT, cost, "MANUAL", 0, RefNumber, Notes
    Else
        ApplyStockMovement db, ProductID, Qty, TransactionTypeID, cost, "MANUAL", 0, RefNumber, Notes, True
    End If
    ws.CommitTrans
    inTrans = False
    LogAction "STOCK_MANUAL", "InventoryTransactions", RefNumber, _
              "Product=" & ProductID & " Type=" & TransactionTypeID & " Qty=" & Qty & " Cost=" & cost
    Exit Function

EH:
    PostManualStock = " ⁄–— Õ›Ÿ «·Õ—ﬂ…: " & Err.Description & " (" & Err.Number & ")"
    RefNumber = ""
    If inTrans Then ws.Rollback
End Function

'------------------------------------------------------------------------------
' Stocktake
'------------------------------------------------------------------------------
Public Function CreateStockCount(ByVal CategoryID As Variant, ByVal Notes As String, _
                                 ByRef NewCountID As Long) As String
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean
    Dim openNo As Variant, countNo As String, scope As String

    On Error GoTo EH
    NewCountID = 0
    Set db = CurrentDb
    If Not HasPermission("STOCK_COUNT") Then
        CreateStockCount = "·«  „·ﬂ ’·«ÕÌ… «·Ã—œ."
        Exit Function
    End If
    openNo = DbValue("SELECT CountNumber FROM StockCounts WHERE Status = 'OPEN'")
    If Not IsNull(openNo) Then
        CreateStockCount = "ÌÊÃœ Ã—œ „› ÊÕ (" & openNo & "). —Õ¯·Â √Ê √·€Â ﬁ»· »œ¡ Ã—œ ÃœÌœ."
        Exit Function
    End If
    scope = "IsActive = True"
    If Not IsNull(CategoryID) Then scope = scope & " AND CategoryID = " & CLng(CategoryID)
    If Nz(DbValue("SELECT COUNT(*) FROM Products WHERE " & scope), 0) = 0 Then
        CreateStockCount = "·«  ÊÃœ „‰ Ã«  ‰‘ÿ… ··Ã—œ ›Ì Â–« «·«Œ Ì«—."
        Exit Function
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    countNo = NextNumber("STOCK_COUNT")
    Set rs = db.OpenRecordset("StockCounts", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!CountNumber = countNo
    rs!CountDate = Now
    If Not IsNull(CategoryID) Then rs!CategoryID = CLng(CategoryID)
    rs!Status = "OPEN"
    rs!EmployeeID = CurrentUserID()
    If Len(Notes) > 0 Then rs!Notes = Left$(Notes, 255)
    rs.Update
    rs.Bookmark = rs.LastModified
    NewCountID = rs!StockCountID
    rs.Close
    db.Execute "INSERT INTO StockCountDetails (StockCountID, ProductID, SystemQuantity, ActualQuantity, " & _
               "Difference, UnitCost, DifferenceValue) SELECT " & NewCountID & ", ProductID, CurrentQuantity, " & _
               "Null, 0, AverageCost, 0 FROM Products WHERE " & scope, dbFailOnError
    ws.CommitTrans
    inTrans = False
    LogAction "STOCK_COUNT_OPEN", "StockCounts", countNo
    Exit Function

EH:
    CreateStockCount = " ⁄–— ≈‰‘«¡ «·Ã—œ: " & Err.Description & " (" & Err.Number & ")"
    NewCountID = 0
    If inTrans Then ws.Rollback
End Function

Public Function PostStockCount(ByVal StockCountID As Long, ByVal UncountedAsZero As Boolean, _
                               ByRef AdjustedLines As Long, ByRef NetValue As Currency) As String
    ' Differences are taken against the quantity at posting time, so sales made while
    ' counting do not appear as shortages. Uncounted lines stay unchanged unless
    ' UncountedAsZero (then the product is treated as not found).
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, p As DAO.Recordset, inTrans As Boolean
    Dim countNo As String, sysQty As Currency, actual As Currency, diff As Currency, cost As Currency

    On Error GoTo EH
    AdjustedLines = 0
    NetValue = 0
    Calendar = vbCalGreg
    Set db = CurrentDb
    If Not HasPermission("STOCK_COUNT") Then
        PostStockCount = "·«  „·ﬂ ’·«ÕÌ… «·Ã—œ."
        Exit Function
    End If
    If Nz(DbValue("SELECT Status FROM StockCounts WHERE StockCountID = " & StockCountID), "") <> "OPEN" Then
        PostStockCount = "«·Ã—œ €Ì— „ÊÃÊœ √Ê  „  —ÕÌ·Â √Ê ≈·€«ƒÂ."
        Exit Function
    End If
    countNo = DbValue("SELECT CountNumber FROM StockCounts WHERE StockCountID = " & StockCountID)

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    Set rs = db.OpenRecordset("SELECT * FROM StockCountDetails WHERE StockCountID = " & StockCountID & _
                              " ORDER BY StockCountDetailID", dbOpenDynaset)
    Do Until rs.EOF
        Set p = db.OpenRecordset("SELECT CurrentQuantity, AverageCost FROM Products WHERE ProductID = " & _
                                 rs!ProductID, dbOpenSnapshot)
        sysQty = p!CurrentQuantity
        cost = p!AverageCost
        p.Close
        If IsNull(rs!ActualQuantity) And Not UncountedAsZero Then
            actual = sysQty
        Else
            actual = Nz(rs!ActualQuantity, 0)
        End If
        diff = actual - sysQty
        rs.Edit
        rs!SystemQuantity = sysQty
        If Not IsNull(rs!ActualQuantity) Or UncountedAsZero Then rs!ActualQuantity = actual
        rs!Difference = diff
        rs!UnitCost = cost
        rs!DifferenceValue = RoundMoney(CDec(diff) * CDec(cost))
        rs.Update
        If diff <> 0 Then
            ApplyStockMovement db, rs!ProductID, diff, TT_ADJUSTMENT, cost, "STOCK_COUNT", StockCountID, countNo
            AdjustedLines = AdjustedLines + 1
            NetValue = NetValue + RoundMoney(CDec(diff) * CDec(cost))
        End If
        rs.MoveNext
    Loop
    rs.Close
    db.Execute "UPDATE StockCounts SET Status = 'POSTED', PostedAt = Now(), PostedByID = " & CurrentUserID() & _
               " WHERE StockCountID = " & StockCountID, dbFailOnError
    ws.CommitTrans
    inTrans = False
    LogAction "STOCK_COUNT_POST", "StockCounts", countNo, "Lines=" & AdjustedLines & " Value=" & NetValue
    Exit Function

EH:
    PostStockCount = " ⁄–—  —ÕÌ· «·Ã—œ: " & Err.Description & " (" & Err.Number & ")"
    AdjustedLines = 0
    NetValue = 0
    If inTrans Then ws.Rollback
End Function

Public Function CancelStockCount(ByVal StockCountID As Long) As String
    If Not HasPermission("STOCK_COUNT") Then
        CancelStockCount = "·«  „·ﬂ ’·«ÕÌ… «·Ã—œ."
        Exit Function
    End If
    If Nz(DbValue("SELECT Status FROM StockCounts WHERE StockCountID = " & StockCountID), "") <> "OPEN" Then
        CancelStockCount = "·« Ì„ﬂ‰ ≈·€«¡ Ã—œ „ı—Õ¯· √Ê „·€Ï."
        Exit Function
    End If
    CurrentDb.Execute "UPDATE StockCounts SET Status = 'CANCELLED' WHERE StockCountID = " & StockCountID, dbFailOnError
    LogAction "STOCK_COUNT_CANCEL", "StockCounts", CStr(StockCountID)
End Function

'------------------------------------------------------------------------------
' Checks
'------------------------------------------------------------------------------
Private Function CheckSupplier(ByVal SupplierID As Long) As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT SupplierName, IsActive FROM Suppliers WHERE SupplierID = " & SupplierID, _
                                     dbOpenSnapshot)
    If rs.EOF Then
        CheckSupplier = "«Œ — «·„Ê—œ."
    ElseIf Not rs!IsActive Then
        CheckSupplier = "«·„Ê—œ ´" & rs!SupplierName & "ª €Ì— ‰‘ÿ."
    End If
    rs.Close
End Function

Private Function PurchaseDocDate(ByVal Typed As Variant, ByRef DocDate As Date) As String
    ' Today -> the current time; an earlier day (invoice entered late) -> that day.
    If IsNull(Typed) Then
        DocDate = Now
    ElseIf Not IsDate(Typed) Then
        PurchaseDocDate = " «—ÌŒ «·›« Ê—… €Ì— ’ÕÌÕ."
    ElseIf DateValue(CDate(Typed)) > Date Then
        PurchaseDocDate = " «—ÌŒ «·›« Ê—… ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ›Ì «·„” ﬁ»·."
    ElseIf DateValue(CDate(Typed)) = Date Then
        DocDate = Now
    Else
        DocDate = DateValue(CDate(Typed))
    End If
End Function
