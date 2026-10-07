Attribute VB_Name = "modPurchases"
'==============================================================================
' modPurchases  -  Retail Store Management System (Phase 7)
'
' Posting of the purchase and inventory documents, each in ONE transaction:
'   PostPurchaseFromCart   purchase invoice from tmpPurchaseLines
'   PostPurchaseReturn     goods returned to the supplier from tmpPurchaseReturnLines
'   PostSupplierPayment    payment voucher (سند صرف)
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
        PostPurchaseFromCart = "لا تملك صلاحية تسجيل فواتير المشتريات."
        Exit Function
    End If
    msg = CheckSupplier(SupplierID)
    If Len(msg) = 0 Then msg = PurchaseDocDate(InvoiceDate, docDate)
    If Len(msg) = 0 And Len(Trim$(SupplierInvoiceNo)) > 0 Then
        If Not IsNull(DbValue("SELECT InvoiceNumber FROM PurchaseInvoices WHERE SupplierID = " & SupplierID & _
                              " AND SupplierInvoiceNo = " & SqlText(Trim$(SupplierInvoiceNo)))) Then
            msg = "فاتورة المورد رقم " & Trim$(SupplierInvoiceNo) & " مسجلة مسبقًا لنفس المورد."
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
            msg = "أحد الأصناف في الفاتورة لم يعد موجودًا. احذفه من الفاتورة."
        ElseIf Not rs!IsActive Then
            msg = "المنتج «" & rs!ProductName & "» غير نشط. فعّله من شاشة المنتجات أولًا."
        ElseIf Not IsNull(rs!NewSellingPrice) Then
            If rs!NewSellingPrice < 0 Then
                msg = "سعر البيع الجديد لا يمكن أن يكون سالبًا («" & rs!ProductName & "»)."
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
        msg = "لا تملك صلاحية تعديل أسعار البيع. امسح عمود «سعر البيع الجديد»."
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
    If remaining > 0 Then rs!DueDate = DueDateFor("S", SupplierID, docDate)       ' modAging
    rs!CashBoxID = CashBoxFor(Nz(PaymentMethodID, CASH_METHOD_ID), paid)      ' modCash
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
    PostPurchaseFromCart = "تعذر حفظ فاتورة الشراء: " & Err.Description & " (" & Err.Number & ")"
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
        PostPurchaseReturn = "لا تملك صلاحية تسجيل مرتجعات المشتريات."
        Exit Function
    End If
    If Len(Trim$(Reason)) = 0 Then
        PostPurchaseReturn = "سبب الإرجاع مطلوب."
        Exit Function
    End If
    If RefundType <> "CASH" And RefundType <> "CREDIT" Then RefundType = "CREDIT"
    supplierID = Nz(DbValue("SELECT SupplierID FROM PurchaseInvoices WHERE PurchaseInvoiceID = " & PurchaseInvoiceID), 0)
    If supplierID = 0 Then
        PostPurchaseReturn = "فاتورة الشراء الأصلية غير موجودة."
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
            PostPurchaseReturn = "سطر لا ينتمي إلى فاتورة الشراء الأصلية."
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
        PostPurchaseReturn = "حدد الكمية المرتجعة لصنف واحد على الأقل."
        Exit Function
    End If
    msg = CheckStockAvailable(productIDs, qtys, n, "الإرجاع للمورد")
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
    rs!CashBoxID = CashBoxFor(Nz(PaymentMethodID, CASH_METHOD_ID), refunded)
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
    PostPurchaseReturn = "تعذر حفظ مرتجع المشتريات: " & Err.Description & " (" & Err.Number & ")"
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
        PostSupplierPayment = "لا تملك صلاحية تسجيل سندات الصرف."
        Exit Function
    End If
    If Amount <= 0 Then
        PostSupplierPayment = "المبلغ يجب أن يكون أكبر من صفر."
        Exit Function
    End If
    msg = CheckSupplier(SupplierID)
    If Len(msg) > 0 Then
        PostSupplierPayment = msg
        Exit Function
    End If
    owed = Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierID = " & SupplierID), 0)
    If Amount > owed Then
        If Not AskYesNo("المبلغ (" & Format$(Amount, "#,##0.00") & ") أكبر من المستحق للمورد (" & _
                        Format$(owed, "#,##0.00") & ")، وسيصبح المورد مدينًا للمحل." & vbCrLf & _
                        "هل تريد المتابعة؟") Then
            PostSupplierPayment = "تم الإلغاء."
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
    rs!CashBoxID = CashBoxFor(PaymentMethodID, Amount)
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
    PostSupplierPayment = "تعذر حفظ سند الصرف: " & Err.Description & " (" & Err.Number & ")"
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
        PostManualStock = "لا تملك صلاحية تعديل المخزون يدويًا."
        Exit Function
    End If
    If IsNull(DbValue("SELECT ProductID FROM Products WHERE ProductID = " & ProductID)) Then
        PostManualStock = "اختر المنتج."
        Exit Function
    End If
    If Not ProductTracksStock(ProductID) Then
        PostManualStock = "هذا الصنف لا يتابع المخزون (يُحضَّر عند الطلب)، فلا تُسجَّل له حركات مخزون."
        Exit Function
    End If
    If Qty <= 0 Then
        PostManualStock = "الكمية يجب أن تكون أكبر من صفر."
        Exit Function
    End If
    avgCost = Nz(DbValue("SELECT AverageCost FROM Products WHERE ProductID = " & ProductID), 0)
    If IsNull(UnitCost) Then cost = avgCost Else cost = CCur(UnitCost)
    If cost < 0 Then
        PostManualStock = "التكلفة لا يمكن أن تكون سالبة."
        Exit Function
    End If

    Select Case TransactionTypeID
        Case TT_OPENING
            If Nz(DbValue("SELECT COUNT(*) FROM InventoryTransactions WHERE ProductID = " & ProductID), 0) > 0 Then
                msg = "الرصيد الافتتاحي يُسجَّل مرة واحدة قبل أي حركة على المنتج." & vbCrLf & _
                      "لتصحيح الكمية الآن استخدم الجرد أو الإضافة/الخصم."
            End If
        Case TT_STOCK_IN, TT_STOCK_OUT
            If Len(Trim$(Notes)) = 0 Then msg = "اكتب سبب الحركة في الملاحظات (مثال: تالف، هدية من المورد)."
            If Len(msg) = 0 And TransactionTypeID = TT_STOCK_OUT Then
                cost = avgCost                            ' goods leave at the average cost
                ids(0) = ProductID
                qtys(0) = Qty
                msg = CheckStockAvailable(ids, qtys, 1, "خصم المخزون")
            End If
        Case Else
            msg = "نوع الحركة غير متاح للإدخال اليدوي."
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
    PostManualStock = "تعذر حفظ الحركة: " & Err.Description & " (" & Err.Number & ")"
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
        CreateStockCount = "لا تملك صلاحية الجرد."
        Exit Function
    End If
    openNo = DbValue("SELECT CountNumber FROM StockCounts WHERE Status = 'OPEN'")
    If Not IsNull(openNo) Then
        CreateStockCount = "يوجد جرد مفتوح (" & openNo & "). رحّله أو ألغه قبل بدء جرد جديد."
        Exit Function
    End If
    scope = "IsActive = True AND TrackStock = True"
    If Not IsNull(CategoryID) Then scope = scope & " AND CategoryID = " & CLng(CategoryID)
    If Nz(DbValue("SELECT COUNT(*) FROM Products WHERE " & scope), 0) = 0 Then
        CreateStockCount = "لا توجد منتجات نشطة للجرد في هذا الاختيار."
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
    CreateStockCount = "تعذر إنشاء الجرد: " & Err.Description & " (" & Err.Number & ")"
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
        PostStockCount = "لا تملك صلاحية الجرد."
        Exit Function
    End If
    If Nz(DbValue("SELECT Status FROM StockCounts WHERE StockCountID = " & StockCountID), "") <> "OPEN" Then
        PostStockCount = "الجرد غير موجود أو تم ترحيله أو إلغاؤه."
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
    PostStockCount = "تعذر ترحيل الجرد: " & Err.Description & " (" & Err.Number & ")"
    AdjustedLines = 0
    NetValue = 0
    If inTrans Then ws.Rollback
End Function

Public Function CancelStockCount(ByVal StockCountID As Long) As String
    If Not HasPermission("STOCK_COUNT") Then
        CancelStockCount = "لا تملك صلاحية الجرد."
        Exit Function
    End If
    If Nz(DbValue("SELECT Status FROM StockCounts WHERE StockCountID = " & StockCountID), "") <> "OPEN" Then
        CancelStockCount = "لا يمكن إلغاء جرد مُرحّل أو ملغى."
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
        CheckSupplier = "اختر المورد."
    ElseIf Not rs!IsActive Then
        CheckSupplier = "المورد «" & rs!SupplierName & "» غير نشط."
    End If
    rs.Close
End Function

Private Function PurchaseDocDate(ByVal Typed As Variant, ByRef DocDate As Date) As String
    ' Today -> the current time; an earlier day (invoice entered late) -> that day.
    If IsNull(Typed) Then
        DocDate = Now
    ElseIf Not IsDate(Typed) Then
        PurchaseDocDate = "تاريخ الفاتورة غير صحيح."
    ElseIf DateValue(CDate(Typed)) > Date Then
        PurchaseDocDate = "تاريخ الفاتورة لا يمكن أن يكون في المستقبل."
    ElseIf DateValue(CDate(Typed)) = Date Then
        DocDate = Now
    Else
        DocDate = DateValue(CDate(Typed))
    End If
End Function
