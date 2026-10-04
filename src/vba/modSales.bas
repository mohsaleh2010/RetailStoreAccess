Attribute VB_Name = "modSales"
'==============================================================================
' modSales  -  Retail Store Management System (Phases 6-7)
'
' 1) Invoice calculation engine (also used by purchases in Phase 7):
'      CalcReset / CalcAddLine / CalcRun / CalcLine / CalcTotal
'    Same algorithm as tools/pricing.py; the tests replay its cases.
' 2) Posting, each in ONE transaction (all or nothing):
'      PostSaleFromCart     sale from the POS cart (tmpPOSLines)
'      PostSalesReturn      credit note from tmpReturnLines
'      PostCustomerPayment  receipt voucher
'      ApplyStockMovement   stock ledger + product quantity (+ average cost)
' Posting functions return "" on success or an Arabic error message.
'==============================================================================
Option Compare Database
Option Explicit

Public Const TT_PURCHASE As Long = 1
Public Const TT_SALE As Long = 2
Public Const TT_PURCHASE_RETURN As Long = 3
Public Const TT_SALES_RETURN As Long = 4
Public Const TT_STOCK_IN As Long = 5
Public Const TT_STOCK_OUT As Long = 6
Public Const TT_ADJUSTMENT As Long = 7
Public Const TT_OPENING As Long = 8

' ---- calculation state
Private m_n As Long
Private m_qty() As Currency
Private m_price() As Currency
Private m_disc() As Currency
Private m_rate() As Currency
Private m_unit() As Currency
Private m_lineDisc() As Currency
Private m_net() As Currency
Private m_tax() As Currency
Private m_total() As Currency
Private m_gross As Currency
Private m_subTotal As Currency
Private m_discTotal As Currency
Private m_taxable As Currency
Private m_taxTotal As Currency
Private m_grand As Currency

'==============================================================================
' 1) Calculation engine
'==============================================================================
Public Sub CalcReset()
    m_n = 0
    ReDim m_qty(0 To 15)
    ReDim m_price(0 To 15)
    ReDim m_disc(0 To 15)
    ReDim m_rate(0 To 15)
End Sub

Public Sub CalcAddLine(ByVal Qty As Currency, ByVal Price As Currency, ByVal Discount As Currency, _
                       ByVal Rate As Currency)
    If m_n > UBound(m_qty) Then
        ReDim Preserve m_qty(0 To 2 * m_n)
        ReDim Preserve m_price(0 To 2 * m_n)
        ReDim Preserve m_disc(0 To 2 * m_n)
        ReDim Preserve m_rate(0 To 2 * m_n)
    End If
    m_qty(m_n) = Qty
    m_price(m_n) = Price
    m_disc(m_n) = Discount
    m_rate(m_n) = Rate
    m_n = m_n + 1
End Sub

Public Function CalcRun(ByVal InvoiceDiscount As Currency, ByVal PricesIncludeVat As Boolean) As String
    ' Returns "" when the invoice is valid, otherwise the reason in Arabic.
    Dim i As Long, biggest As Long, gross() As Currency, share() As Currency, allocated As Currency
    Dim after As Currency, rate As Variant

    If m_n = 0 Then
        CalcRun = "لا توجد أصناف في الفاتورة."
        Exit Function
    End If
    ReDim gross(0 To m_n - 1)
    ReDim share(0 To m_n - 1)
    ReDim m_unit(0 To m_n - 1)
    ReDim m_lineDisc(0 To m_n - 1)
    ReDim m_net(0 To m_n - 1)
    ReDim m_tax(0 To m_n - 1)
    ReDim m_total(0 To m_n - 1)

    m_gross = 0
    For i = 0 To m_n - 1
        If m_qty(i) <= 0 Then
            CalcRun = "الكمية يجب أن تكون أكبر من صفر (السطر " & (i + 1) & ")."
            Exit Function
        End If
        If m_price(i) < 0 Or m_disc(i) < 0 Then
            CalcRun = "السعر والخصم لا يمكن أن يكونا سالبين (السطر " & (i + 1) & ")."
            Exit Function
        End If
        gross(i) = RoundMoney(CDec(m_qty(i)) * CDec(m_price(i))) - m_disc(i)
        If gross(i) < 0 Then
            CalcRun = "خصم السطر " & (i + 1) & " أكبر من قيمته."
            Exit Function
        End If
        m_gross = m_gross + gross(i)
    Next
    If InvoiceDiscount < 0 Or InvoiceDiscount > m_gross Then
        CalcRun = "خصم الفاتورة أكبر من قيمتها."
        Exit Function
    End If

    ' spread the invoice discount in proportion; the remainder goes to the largest line
    If InvoiceDiscount > 0 Then
        biggest = 0
        For i = 0 To m_n - 1
            share(i) = RoundMoney(CDec(InvoiceDiscount) * CDec(gross(i)) / CDec(m_gross))
            allocated = allocated + share(i)
            If gross(i) > gross(biggest) Then biggest = i
        Next
        share(biggest) = share(biggest) + (InvoiceDiscount - allocated)
    End If

    m_subTotal = 0: m_discTotal = 0: m_taxable = 0: m_taxTotal = 0: m_grand = 0
    For i = 0 To m_n - 1
        after = gross(i) - share(i)
        rate = CDec(m_rate(i))
        If PricesIncludeVat Then
            m_total(i) = after
            m_tax(i) = RoundMoney(CDec(after) * rate / (1 + rate))
            m_net(i) = m_total(i) - m_tax(i)
            m_unit(i) = RoundMoney(CDec(m_price(i)) / (1 + rate), 4)
        Else
            m_net(i) = after
            m_tax(i) = RoundMoney(CDec(after) * rate)
            m_total(i) = m_net(i) + m_tax(i)
            m_unit(i) = RoundMoney(m_price(i), 4)
        End If
        m_lineDisc(i) = RoundMoney(CDec(m_qty(i)) * CDec(m_unit(i))) - m_net(i)
        If m_lineDisc(i) < 0 Then
            ' rounding pushed the net above qty x price: lift the unit price instead
            m_unit(i) = RoundMoney(CDec(m_net(i)) / CDec(m_qty(i)), 4)
            m_lineDisc(i) = RoundMoney(CDec(m_qty(i)) * CDec(m_unit(i))) - m_net(i)
            If m_lineDisc(i) < 0 Then m_lineDisc(i) = 0
        End If
        m_subTotal = m_subTotal + m_net(i) + m_lineDisc(i)
        m_discTotal = m_discTotal + m_lineDisc(i)
        m_taxable = m_taxable + m_net(i)
        m_taxTotal = m_taxTotal + m_tax(i)
        m_grand = m_grand + m_total(i)
    Next
End Function

Public Function CalcLine(ByVal Index As Long, ByVal Part As String) As Currency
    Select Case Part
        Case "UNIT":     CalcLine = m_unit(Index)
        Case "DISCOUNT": CalcLine = m_lineDisc(Index)
        Case "NET":      CalcLine = m_net(Index)
        Case "TAX":      CalcLine = m_tax(Index)
        Case "TOTAL":    CalcLine = m_total(Index)
        Case "RATE":     CalcLine = m_rate(Index)
        Case "QTY":      CalcLine = m_qty(Index)
    End Select
End Function

Public Function CalcTotal(ByVal Part As String) As Currency
    Select Case Part
        Case "GROSS":    CalcTotal = m_gross
        Case "SUBTOTAL": CalcTotal = m_subTotal
        Case "DISCOUNT": CalcTotal = m_discTotal
        Case "TAXABLE":  CalcTotal = m_taxable
        Case "TAX":      CalcTotal = m_taxTotal
        Case "TOTAL":    CalcTotal = m_grand
    End Select
End Function

Public Function CalcCount() As Long
    CalcCount = m_n
End Function

Public Function Settle(ByVal Total As Currency, ByVal Tendered As Currency, ByVal IsCredit As Boolean, _
                       ByRef Paid As Currency, ByRef Remaining As Currency, ByRef Change As Currency) As String
    ' Cash sales must be paid in full; credit sales may be partly paid.
    If Tendered < 0 Then Tendered = 0
    If Not IsCredit And Tendered < Total Then
        Settle = "المبلغ المدفوع (" & Format$(Tendered, "#,##0.00") & ") أقل من إجمالي الفاتورة (" & _
                 Format$(Total, "#,##0.00") & ")."
        Exit Function
    End If
    If Tendered < Total Then Paid = Tendered Else Paid = Total
    Remaining = Total - Paid
    If Tendered > Total Then Change = Tendered - Total Else Change = 0
End Function

Public Function ReturnAmounts(ByVal SoldQty As Currency, ByVal SoldTotal As Currency, _
                              ByVal SoldTax As Currency, ByVal PrevQty As Currency, _
                              ByVal PrevTotal As Currency, ByVal PrevTax As Currency, _
                              ByVal Qty As Currency, ByRef NetOut As Currency, _
                              ByRef TaxOut As Currency, ByRef TotalOut As Currency) As String
    ' Partial returns are proportional; the return that completes the line takes what is left.
    If Qty <= 0 Or PrevQty + Qty > SoldQty Then
        ReturnAmounts = "الكمية المرتجعة (" & Qty & ") أكبر من المتبقي (" & (SoldQty - PrevQty) & ")."
        Exit Function
    End If
    If PrevQty + Qty = SoldQty Then
        TotalOut = SoldTotal - PrevTotal
        TaxOut = SoldTax - PrevTax
    Else
        TotalOut = RoundMoney(CDec(SoldTotal) * CDec(Qty) / CDec(SoldQty))
        TaxOut = RoundMoney(CDec(SoldTax) * CDec(Qty) / CDec(SoldQty))
    End If
    NetOut = TotalOut - TaxOut
End Function

Public Function WeightedAverage(ByVal CurQty As Currency, ByVal CurAvg As Currency, _
                                ByVal Qty As Currency, ByVal UnitCost As Currency) As Currency
    ' Average cost after Qty units (signed) at UnitCost. Same rule as tools/pricing.py.
    '   in : (CurQty x CurAvg + Qty x UnitCost) / (CurQty + Qty); no stock before -> UnitCost
    '   out at a known cost (purchase return): the same formula with a negative Qty, as long as
    '        stock remains and the result is not negative; otherwise the average is kept.
    Dim newQty As Currency, v As Currency
    newQty = CurQty + Qty
    WeightedAverage = CurAvg
    If Qty > 0 Then
        If CurQty <= 0 Then
            WeightedAverage = UnitCost
        Else
            WeightedAverage = RoundMoney((CDec(CurQty) * CDec(CurAvg) + CDec(Qty) * CDec(UnitCost)) / CDec(newQty), 4)
        End If
    ElseIf Qty < 0 And newQty > 0 And CurQty > 0 Then
        v = RoundMoney((CDec(CurQty) * CDec(CurAvg) + CDec(Qty) * CDec(UnitCost)) / CDec(newQty), 4)
        If v >= 0 Then WeightedAverage = v
    End If
End Function

Public Function PurchaseUnitCost(ByVal NetAmount As Currency, ByVal LineTotal As Currency, _
                                 ByVal Qty As Currency, ByVal VatRegistered As Boolean) As Currency
    ' Cost of one purchased unit after discounts. A VAT-registered store reclaims the input VAT,
    ' so its cost excludes VAT; an unregistered store bears the VAT as part of the cost.
    If Qty <= 0 Then Exit Function
    If VatRegistered Then
        PurchaseUnitCost = RoundMoney(CDec(NetAmount) / CDec(Qty), 4)
    Else
        PurchaseUnitCost = RoundMoney(CDec(LineTotal) / CDec(Qty), 4)
    End If
End Function

'==============================================================================
' 2) Posting
'==============================================================================
Public Function PostSaleFromCart(ByVal CustomerID As Long, ByVal PaymentType As String, _
                                 ByVal PaymentMethodID As Variant, ByVal InvoiceDiscount As Currency, _
                                 ByVal Tendered As Variant, ByVal Notes As String, _
                                 ByRef NewInvoiceID As Long) As String
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean
    Dim n As Long, i As Long, msg As String, vatRate As Currency, rate As Currency
    Dim productIDs() As Long, costs() As Currency, vatCats() As String
    Dim isCredit As Boolean, tend As Currency, paid As Currency, remaining As Currency, change As Currency
    Dim invoiceID As Long, invNo As String, invDate As Date, subType As String, priceChanged As Boolean

    On Error GoTo EH
    NewInvoiceID = 0
    Calendar = vbCalGreg
    Set db = CurrentDb
    vatRate = Nz(SettingValue("VATRate"), 0.15)

    ' ---- 1. cart -> calculation engine
    Set rs = db.OpenRecordset( _
        "SELECT t.ProductID, t.Quantity, t.UnitPrice, t.LineDiscount, p.ProductName, " & _
        "p.VATCategory, p.AverageCost, p.SellingPrice, p.IsActive " & _
        "FROM tmpPOSLines AS t LEFT JOIN Products AS p ON t.ProductID = p.ProductID " & _
        "ORDER BY t.LineNo", dbOpenSnapshot)
    CalcReset
    Do Until rs.EOF
        If IsNull(rs!ProductName) Then
            PostSaleFromCart = "أحد الأصناف في الفاتورة لم يعد موجودًا. احذفه من الفاتورة."
            rs.Close
            Exit Function
        End If
        If Not rs!IsActive Then
            PostSaleFromCart = "المنتج «" & rs!ProductName & "» غير نشط ولا يمكن بيعه."
            rs.Close
            Exit Function
        End If
        If rs!UnitPrice <> rs!SellingPrice Then priceChanged = True
        If Nz(rs!VATCategory, "S") = "S" Then rate = vatRate Else rate = 0
        CalcAddLine rs!Quantity, rs!UnitPrice, Nz(rs!LineDiscount, 0), rate
        ReDim Preserve productIDs(0 To n)
        ReDim Preserve costs(0 To n)
        ReDim Preserve vatCats(0 To n)
        productIDs(n) = rs!ProductID
        costs(n) = Nz(rs!AverageCost, 0)
        vatCats(n) = Nz(rs!VATCategory, "S")
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    msg = CalcRun(InvoiceDiscount, Nz(SettingValue("PricesIncludeVAT"), True))
    If Len(msg) > 0 Then
        PostSaleFromCart = msg
        Exit Function
    End If

    ' ---- 2. permissions: price changes and discounts
    If priceChanged And Not HasPermission("PRICE_OVERRIDE") Then
        PostSaleFromCart = "لا تملك صلاحية تعديل سعر البيع."
        Exit Function
    End If
    msg = CheckDiscountAllowed(InvoiceDiscount)
    If Len(msg) > 0 Then
        PostSaleFromCart = msg
        Exit Function
    End If

    ' ---- 3. customer and payment
    isCredit = (PaymentType = "CREDIT")
    If IsNull(Tendered) Then
        If isCredit Then tend = 0 Else tend = CalcTotal("TOTAL")
    Else
        tend = CCur(Tendered)
    End If
    msg = Settle(CalcTotal("TOTAL"), tend, isCredit, paid, remaining, change)
    If Len(msg) = 0 Then msg = CheckCustomer(CustomerID, isCredit, remaining)
    If Len(msg) = 0 Then msg = CheckStock(productIDs, n)
    If Len(msg) > 0 Then
        PostSaleFromCart = msg
        Exit Function
    End If

    ' ---- 4. post everything in one transaction
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    invNo = NextNumber("SALES_INVOICE")
    invDate = Now
    subType = InvoiceSubType(CustomerID)

    Set rs = db.OpenRecordset("SalesInvoices", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!InvoiceNumber = invNo
    rs!InvoiceDate = invDate
    rs!CustomerID = CustomerID
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
    rs!AmountTendered = tend
    rs!ChangeDue = change
    If Len(Notes) > 0 Then rs!Notes = Left$(Notes, 255)
    SetZatcaFields rs, subType, "388", invDate, CalcTotal("TOTAL"), CalcTotal("TAX")
    rs.Update
    rs.Bookmark = rs.LastModified
    invoiceID = rs!SalesInvoiceID
    rs.Close

    Set rs = db.OpenRecordset("SalesInvoiceDetails", dbOpenDynaset, dbAppendOnly)
    For i = 0 To n - 1
        rs.AddNew
        rs!SalesInvoiceID = invoiceID
        rs!LineNumber = i + 1
        rs!ProductID = productIDs(i)
        rs!Quantity = CalcLine(i, "QTY")
        rs!UnitPrice = CalcLine(i, "UNIT")
        rs!Discount = CalcLine(i, "DISCOUNT")
        rs!NetAmount = CalcLine(i, "NET")
        rs!VATCategory = vatCats(i)
        rs!VATRate = CalcLine(i, "RATE")
        rs!Tax = CalcLine(i, "TAX")
        rs!LineTotal = CalcLine(i, "TOTAL")
        rs!UnitCost = costs(i)
        rs.Update
    Next
    rs.Close

    For i = 0 To n - 1
        ApplyStockMovement db, productIDs(i), -CalcLine(i, "QTY"), TT_SALE, costs(i), "SALE", invoiceID, invNo
    Next
    If remaining <> 0 Then AdjustBalance db, "Customers", "CustomerID", CustomerID, remaining

    ws.CommitTrans
    inTrans = False
    LogAction "SALE", "SalesInvoices", invNo, "Total=" & CalcTotal("TOTAL")
    NewInvoiceID = invoiceID
    Exit Function

EH:
    PostSaleFromCart = "تعذر حفظ الفاتورة: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
End Function

Public Function PostSalesReturn(ByVal SalesInvoiceID As Long, ByVal Reason As String, _
                                ByVal RefundType As String, ByVal PaymentMethodID As Variant, _
                                ByRef NewReturnID As Long) As String
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean
    Dim inv As DAO.Recordset, n As Long, i As Long, msg As String
    Dim detailIDs() As Long, productIDs() As Long, qtys() As Currency, toStock() As Boolean
    Dim units() As Currency, nets() As Currency, taxes() As Currency, totals() As Currency
    Dim costs() As Currency, rates() As Currency, cats() As String, discs() As Currency
    Dim sumNet As Currency, sumTax As Currency, sumTotal As Currency, sumDisc As Currency
    Dim refunded As Currency, customerID As Long, retID As Long, retNo As String, retDate As Date
    Dim subType As String

    On Error GoTo EH
    NewReturnID = 0
    Calendar = vbCalGreg
    Set db = CurrentDb
    If Len(Trim$(Reason)) = 0 Then
        PostSalesReturn = "سبب الإرجاع مطلوب (إلزامي في الإشعار الدائن)."
        Exit Function
    End If
    Set inv = db.OpenRecordset("SELECT * FROM SalesInvoices WHERE SalesInvoiceID = " & SalesInvoiceID, _
                               dbOpenSnapshot)
    If inv.EOF Then
        inv.Close
        PostSalesReturn = "الفاتورة الأصلية غير موجودة."
        Exit Function
    End If
    customerID = inv!CustomerID
    subType = inv!InvoiceSubType
    inv.Close           ' before BeginTrans: closing it inside the transaction raises 3246
    If CustomerID = Nz(SettingValue("DefaultCustomerID"), 1) Then RefundType = "CASH"

    Set rs = db.OpenRecordset( _
        "SELECT r.SalesDetailID, r.ReturnQty, r.ReturnToStock, d.ProductID, d.Quantity, d.LineTotal, " & _
        "d.Tax, d.UnitPrice, d.UnitCost, d.VATRate, d.VATCategory, d.SalesInvoiceID " & _
        "FROM tmpReturnLines AS r INNER JOIN SalesInvoiceDetails AS d ON r.SalesDetailID = d.SalesDetailID " & _
        "WHERE r.ReturnQty > 0", dbOpenSnapshot)
    Do Until rs.EOF
        If rs!SalesInvoiceID <> SalesInvoiceID Then
            PostSalesReturn = "سطر لا ينتمي إلى الفاتورة الأصلية."
            rs.Close
            Exit Function
        End If
        ReDim Preserve detailIDs(0 To n): ReDim Preserve productIDs(0 To n)
        ReDim Preserve qtys(0 To n): ReDim Preserve toStock(0 To n)
        ReDim Preserve units(0 To n): ReDim Preserve nets(0 To n)
        ReDim Preserve taxes(0 To n): ReDim Preserve totals(0 To n)
        ReDim Preserve costs(0 To n): ReDim Preserve rates(0 To n)
        ReDim Preserve cats(0 To n): ReDim Preserve discs(0 To n)
        detailIDs(n) = rs!SalesDetailID
        productIDs(n) = rs!ProductID
        qtys(n) = rs!ReturnQty
        toStock(n) = Nz(rs!ReturnToStock, True)
        units(n) = rs!UnitPrice
        costs(n) = rs!UnitCost
        rates(n) = rs!VATRate
        cats(n) = rs!VATCategory
        msg = ReturnAmounts(rs!Quantity, rs!LineTotal, rs!Tax, _
                            PrevReturned(db, rs!SalesDetailID, "Quantity"), _
                            PrevReturned(db, rs!SalesDetailID, "LineTotal"), _
                            PrevReturned(db, rs!SalesDetailID, "Tax"), _
                            qtys(n), nets(n), taxes(n), totals(n))
        If Len(msg) > 0 Then
            PostSalesReturn = msg
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
        PostSalesReturn = "حدد الكمية المرتجعة لصنف واحد على الأقل."
        Exit Function
    End If
    If RefundType = "CASH" Then refunded = sumTotal Else refunded = 0

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    retNo = NextNumber("SALES_RETURN")
    retDate = Now
    Set rs = db.OpenRecordset("SalesReturns", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!ReturnNumber = retNo
    rs!ReturnDate = retDate
    rs!SalesInvoiceID = SalesInvoiceID
    rs!CustomerID = customerID
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
    SetZatcaFields rs, subType, "381", retDate, sumTotal, sumTax
    rs.Update
    rs.Bookmark = rs.LastModified
    retID = rs!SalesReturnID
    rs.Close

    Set rs = db.OpenRecordset("SalesReturnDetails", dbOpenDynaset, dbAppendOnly)
    For i = 0 To n - 1
        rs.AddNew
        rs!SalesReturnID = retID
        rs!SalesDetailID = detailIDs(i)
        rs!ProductID = productIDs(i)
        rs!Quantity = qtys(i)
        rs!UnitPrice = units(i)
        rs!Discount = discs(i)
        rs!NetAmount = nets(i)
        rs!VATCategory = cats(i)
        rs!VATRate = rates(i)
        rs!Tax = taxes(i)
        rs!LineTotal = totals(i)
        rs!UnitCost = costs(i)
        rs!ReturnToStock = toStock(i)
        rs.Update
    Next
    rs.Close

    For i = 0 To n - 1
        If toStock(i) Then
            ApplyStockMovement db, productIDs(i), qtys(i), TT_SALES_RETURN, costs(i), "SALES_RETURN", _
                               retID, retNo, "", True
        End If
    Next
    If refunded - sumTotal <> 0 Then AdjustBalance db, "Customers", "CustomerID", customerID, refunded - sumTotal

    ws.CommitTrans
    inTrans = False
    LogAction "SALES_RETURN", "SalesReturns", retNo, "Invoice=" & SalesInvoiceID & " Total=" & sumTotal
    NewReturnID = retID
    Exit Function

EH:
    PostSalesReturn = "تعذر حفظ المرتجع: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
End Function

Public Function PostCustomerPayment(ByVal CustomerID As Long, ByVal Amount As Currency, _
                                    ByVal PaymentMethodID As Long, ByVal Notes As String, _
                                    ByRef NewPaymentID As Long) As String
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean
    Dim balance As Currency, payNo As String

    On Error GoTo EH
    NewPaymentID = 0
    Set db = CurrentDb
    If Amount <= 0 Then
        PostCustomerPayment = "المبلغ يجب أن يكون أكبر من صفر."
        Exit Function
    End If
    If CustomerID = Nz(SettingValue("DefaultCustomerID"), 1) Then
        PostCustomerPayment = "لا تُسجَّل دفعات على العميل النقدي."
        Exit Function
    End If
    balance = Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerID = " & CustomerID), 0)
    If Amount > balance Then
        If Not AskYesNo("المبلغ (" & Format$(Amount, "#,##0.00") & ") أكبر من الرصيد المستحق (" & _
                        Format$(balance, "#,##0.00") & ")، وسيصبح للعميل رصيد دائن." & vbCrLf & _
                        "هل تريد المتابعة؟") Then
            PostCustomerPayment = "تم الإلغاء."
            Exit Function
        End If
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    payNo = NextNumber("CUSTOMER_PAYMENT")
    Set rs = db.OpenRecordset("CustomerPayments", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!PaymentNumber = payNo
    rs!CustomerID = CustomerID
    rs!PaymentDate = Now
    rs!Amount = Amount
    rs!PaymentMethodID = PaymentMethodID
    rs!EmployeeID = CurrentUserID()
    If Len(Notes) > 0 Then rs!Notes = Left$(Notes, 255)
    rs.Update
    rs.Bookmark = rs.LastModified
    NewPaymentID = rs!PaymentID
    rs.Close
    AdjustBalance db, "Customers", "CustomerID", CustomerID, -Amount
    ws.CommitTrans
    inTrans = False
    LogAction "CUSTOMER_PAYMENT", "CustomerPayments", payNo, "Amount=" & Amount
    Exit Function

EH:
    PostCustomerPayment = "تعذر حفظ سند القبض: " & Err.Description & " (" & Err.Number & ")"
    NewPaymentID = 0
    If inTrans Then ws.Rollback
End Function

'------------------------------------------------------------------------------
' Stock and balances (called inside the posting transactions)
'------------------------------------------------------------------------------
Public Sub ApplyStockMovement(ByVal db As DAO.Database, ByVal ProductID As Long, ByVal Qty As Currency, _
                              ByVal TransactionTypeID As Long, ByVal UnitCost As Currency, _
                              ByVal RefType As String, ByVal RefID As Long, ByVal RefNumber As String, _
                              Optional ByVal Notes As String = "", _
                              Optional ByVal RecalcAverage As Boolean = False)
    ' Qty is signed (+ in, - out). RecalcAverage: goods coming in (purchase, opening, stock in,
    ' sales return) or going back to the supplier at their purchase cost update the average cost.
    Dim rs As DAO.Recordset, curQty As Currency, newQty As Currency
    Set rs = db.OpenRecordset("SELECT ProductID, CurrentQuantity, AverageCost, PurchasePrice, UpdatedAt " & _
                              "FROM Products WHERE ProductID = " & ProductID, dbOpenDynaset)
    rs.Edit
    curQty = rs!CurrentQuantity
    newQty = curQty + Qty
    If RecalcAverage Then rs!AverageCost = WeightedAverage(curQty, rs!AverageCost, Qty, UnitCost)
    rs!CurrentQuantity = newQty
    rs!UpdatedAt = Now
    rs.Update
    rs.Close

    Set rs = db.OpenRecordset("InventoryTransactions", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!TransactionDate = Now
    rs!ProductID = ProductID
    rs!TransactionTypeID = TransactionTypeID
    rs!Quantity = Qty
    rs!UnitCost = UnitCost
    rs!QuantityAfter = newQty
    rs!ReferenceType = RefType
    If RefID > 0 Then rs!ReferenceID = RefID
    If Len(RefNumber) > 0 Then rs!ReferenceNumber = Left$(RefNumber, 20)
    rs!EmployeeID = CurrentUserID()
    If Len(Notes) > 0 Then rs!Notes = Left$(Notes, 255)
    rs.Update
    rs.Close
End Sub

Public Sub AdjustBalance(ByVal db As DAO.Database, ByVal TableName As String, ByVal KeyField As String, _
                         ByVal KeyValue As Long, ByVal Delta As Currency)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT " & KeyField & ", CurrentBalance FROM " & TableName & _
                              " WHERE " & KeyField & " = " & KeyValue, dbOpenDynaset)
    rs.Edit
    rs!CurrentBalance = rs!CurrentBalance + Delta
    rs.Update
    rs.Close
End Sub

'------------------------------------------------------------------------------
' Checks
'------------------------------------------------------------------------------
Private Function CheckCustomer(ByVal CustomerID As Long, ByVal IsCredit As Boolean, _
                               ByVal Remaining As Currency) As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Customers WHERE CustomerID = " & CustomerID, dbOpenSnapshot)
    If rs.EOF Then
        CheckCustomer = "العميل غير موجود."
    ElseIf Not rs!IsActive Then
        CheckCustomer = "العميل «" & rs!CustomerName & "» غير نشط."
    ElseIf IsCredit Or Remaining > 0 Then
        If Not rs!AllowCredit Then
            CheckCustomer = "العميل «" & rs!CustomerName & "» غير مسموح له بالبيع الآجل." & vbCrLf & _
                            "للبيع الآجل اختر عميلًا مسجلًا."
        ElseIf rs!CreditLimit > 0 And rs!CurrentBalance + Remaining > rs!CreditLimit Then
            CheckCustomer = "تجاوز حد الائتمان: الرصيد " & Format$(rs!CurrentBalance, "#,##0.00") & _
                            " + المتبقي " & Format$(Remaining, "#,##0.00") & " أكبر من الحد " & _
                            Format$(rs!CreditLimit, "#,##0.00") & "."
        End If
    End If
    rs.Close
End Function

Private Function CheckStock(ByRef ProductIDs() As Long, ByVal n As Long) As String
    Dim qtys() As Currency, i As Long
    ReDim qtys(0 To n - 1)
    For i = 0 To n - 1
        qtys(i) = CalcLine(i, "QTY")
    Next
    CheckStock = CheckStockAvailable(ProductIDs, qtys, n, "البيع")
End Function

Public Function CheckStockAvailable(ByRef ProductIDs() As Long, ByRef Qtys() As Currency, ByVal n As Long, _
                                    ByVal Operation As String) As String
    ' Sums the quantities per product and compares them with the stock. Taking out more than
    ' the stock needs the AllowNegativeStock setting or an administrator's approval.
    Dim i As Long, j As Long, need As Currency, available As Currency, shortList As String
    For i = 0 To n - 1
        need = 0
        For j = 0 To n - 1
            If ProductIDs(j) = ProductIDs(i) Then
                If j < i Then GoTo NextLine          ' already checked
                need = need + Qtys(j)
            End If
        Next
        available = Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & ProductIDs(i)), 0)
        If need > available Then
            shortList = shortList & vbCrLf & "- " & DbValue("SELECT ProductName FROM Products WHERE ProductID = " & _
                        ProductIDs(i)) & ": المطلوب " & need & " والمتوفر " & available
        End If
NextLine:
    Next
    If Len(shortList) = 0 Then Exit Function

    If Nz(SettingValue("AllowNegativeStock"), False) Then
        LogAction "NEGATIVE_STOCK", "Products", "", Operation & " - allowed by settings:" & shortList
    ElseIf HasPermission("ALLOW_NEGATIVE_STOCK") Then
        If AskYesNo("الكمية المتوفرة غير كافية:" & shortList & vbCrLf & vbCrLf & _
                    "هل تسمح بإتمام العملية (" & Operation & ") رغم ذلك؟ (بصلاحية مدير النظام)") Then
            LogAction "NEGATIVE_STOCK", "Products", "", Operation & " - approved by user:" & shortList
        Else
            CheckStockAvailable = "تم إلغاء الحفظ: الكمية غير كافية."
        End If
    Else
        CheckStockAvailable = "الكمية المتوفرة غير كافية:" & shortList & vbCrLf & _
                              Operation & " بأكثر من المتوفر يحتاج موافقة مدير النظام."
    End If
End Function

Private Function CheckDiscountAllowed(ByVal InvoiceDiscount As Currency) As String
    Dim maxPct As Currency, pct As Double, lineDisc As Currency, i As Long
    For i = 0 To m_n - 1
        lineDisc = lineDisc + m_disc(i)
    Next
    If CalcTotal("GROSS") + lineDisc <= 0 Then Exit Function      ' value before any discount
    pct = (lineDisc + InvoiceDiscount) / (CalcTotal("GROSS") + lineDisc)
    If pct <= 0 Then Exit Function
    maxPct = Nz(DbValue("SELECT MaxDiscountPercent FROM Employees WHERE EmployeeID = " & CurrentUserID()), 0)
    If pct > maxPct And Not HasPermission("DISCOUNT_OVERRIDE") Then
        CheckDiscountAllowed = "الخصم (" & Format$(pct, "0.0%") & ") أكبر من الحد المسموح لك (" & _
                               Format$(maxPct, "0.0%") & ")."
    End If
End Function

Private Function PrevReturned(ByVal db As DAO.Database, ByVal SalesDetailID As Long, _
                              ByVal FieldName As String) As Currency
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Sum(" & FieldName & ") FROM SalesReturnDetails WHERE SalesDetailID = " & _
                              SalesDetailID, dbOpenSnapshot)
    PrevReturned = Nz(rs(0), 0)
    rs.Close
End Function

Private Sub SetZatcaFields(ByVal rs As DAO.Recordset, ByVal SubType As String, ByVal TypeCode As String, _
                           ByVal DocDate As Date, ByVal Total As Currency, ByVal Tax As Currency)
    Dim vatNo As String
    rs!InvoiceSubType = SubType
    rs!InvoiceTypeCode = TypeCode
    rs!InvoiceUUID = NewUUID()
    rs!ICV = CLng(NextNumber("ZATCA_ICV"))
    rs!ZatcaStatus = ZatcaInitialStatus()
    vatNo = Nz(SettingValue("VATNumber"), "")
    If Len(vatNo) > 0 Then          ' a QR code is only required from VAT-registered sellers
        rs!QRCodeData = BuildZatcaQR(Nz(SettingValue("StoreName"), ""), vatNo, DocDate, Total, Tax)
    End If
End Sub
