Attribute VB_Name = "modTestSales"
'==============================================================================
' modTestSales  -  Retail Store Management System (Phase 6 tests)
'
' GENERATED FILE - do not edit by hand.  Source: tools/gen_test_sales.py
'
'   TestSales   1) calculation, ZATCA and QR vectors (same as the Python references)
'               2) the full sales cycle inside a transaction that is rolled back:
'                  opening stock 100 -> cash sale 10 -> stock 90 (spec), credit sale,
'                  receipt voucher, return, and the refusals (unpaid cash sale,
'                  credit for the cash customer, more than the stock, over-return)
'               3) the point-of-sale screen itself (scan, merge, discount, change)
'==============================================================================
Option Compare Database
Option Explicit

Private m_passed As Long
Private m_failed As Long
Private m_report As String

Public Function TestSales() As Boolean
    m_passed = 0: m_failed = 0: m_report = ""
    Debug.Print "=== TestSales  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    g_AutoAnswer = True
    EnsureLocalTables

    TestVectors
    TestPosting
    TestPOSScreen

    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & m_passed & " | ›‘·: " & m_failed
    If m_failed = 0 Then
        MsgBox "Ã„Ì⁄ «Œ »«—«  «·„»Ì⁄«  ‰«ÃÕ… (" & m_passed & " «Œ »«—«)." & vbCrLf & _
               "·„  ı —ﬂ √Ì »Ì«‰«  «Œ »«—.", vbInformation + MSG_RTL, "TestSales"
        TestSales = True
    Else
        MsgBox "‰ÃÕ " & m_passed & " Ê›‘· " & m_failed & ":" & vbCrLf & vbCrLf & Left$(m_report, 900), _
               vbExclamation + MSG_RTL, "TestSales"
    End If
End Function

'------------------------------------------------------------------------------
' 1) vectors
'------------------------------------------------------------------------------
Private Sub TestVectors()
    Dim e As String, t As String, msg As String
    e = "41:FED2B45E3FC1205B34D06E9CBAE9ABB754EBA9D5DBABE4D4E2EC17F5915907FAAAAAAAFE01EF8B0D008BD203A9FCFE825209AF3C9B92CC9F353759E7E89DBE0CB10D3D872486796E7AF1CFCB97A35EBE6C981B458ED5064AF"
    e = e & "9391CBFD6C8ABEE0A29AAE396F69FB2919AE4952C6710C8D36A93AE49E23A70DB7A4B6E0B0A885E5216D8273F268589DF972F1F6E86E99AFE887D04964D39F174C049D128C3F2715FCD49BAF6E5D3C1FD805C8295C63FB9724DA"
    e = e & "B30445D56B1BBAEB3A75FF5D230816E76E9CF36639304A84F6B2BFEA0E744C70"
    t = "AQxCb2JzIFJlY29yZHMCDzMxMDEyMjM5MzUwMDAwMwMUMjAyMi0wNC0yNVQxNTozMDowMFoEBzEwMDAuMDAFBjE1MC4wMA=="
    Call Record(QRChecksum(t) = e, "—„“ QR ·„À«· «·ÂÌ∆… «·—”„Ì")
    e = "45:FE29BB7F4BFC14FA006A906E937DF324BB74B0170035DBACB8FAEBAEC1009C620107FAAAAAAAAFE0026D120A00AA33CFFF8890234E820C0B2CE4CEB4AA1EB077E09ACF02B307AEF93CBA44298040B24D47A736ADE79A48420"
    e = e & "BC09EEBE33BF7D62CEAA8C48CB0B42A7FB70ED97B4260AD26FEAEF9BEFF8C4D9C50CC62EBA6EACDEA6B10531B4F135FBCBF9FDF81CB74500872A6F039E8F1EF4AD8E4217BA8A46FE5FE648ED4FE080E0868786AEAD9763875D1D"
    e = e & "0F2A2BC7E6FA2A545C4F6404082FA60B109DEF31920AADC39BE0AFFFEFD8041BC64845FF9ABABF12A7045CD1D2F11BAF4EFCB9FE5D3F82C804F6EA14DE25EFF0416B266002FEC2FBFFAB18"
    t = "ARnZhdiq2KzYsSDYp9mE2KfYrtiq2KjYp9ixAg8zMDAwMDAwMDAwMDAwMDMDFDIwMjYtMTAtMDRUMDk6MTU6MDBaBAcxMTUwLjAwBQYxNTAuMDA="
    Call Record(QRChecksum(t) = e, "—„“ QR ·«”„ „Õ· ⁄—»Ì")
    e = "77:FEC4D69AA92D5DDDC3FC1054DCA20E055555506EBE874FBBBB755542BB749B503AAAA5111115DBA44D0FAAAABFDDDBAEC1440D46222115555107FAAAAAAAAAAAAAAAAFE0048C712E2245555500A3734EFACAAFF777712942D"
    e = e & "5372BAA951111608FFDFB6AFEEF555556AA784AE6C2222455555AA73D36E62AAAF777776B9EBD4F52AAA95111160889DF16AEEEF555556AA484AC6CE222455555AA43DB6E74AAAF777776B9EB54713AAA95111160889AF02BEEE"
    e = e & "F555556AA4842EBC2222455555AAC3DB686AAAAF777776BBEB5459EAAA95111160B89AF08AEEEF555556BAC842ECD2222455555AACFDB6BFBAAAFF7777FB9C754746AAA911111C49AAAC02AEEEEAD5556AB5182AF12222455557"
    e = e & "1A4FDB68FAAAABF7777FBF6055502AAAA4111108B0A2C870EEEFD55552EBBB822F72222955556AA7BD3606ABAABD7776ABF6057902AA8A4111108BAA6D068EEFFD55552ED93926FB2222955556AA76DF004AAAABD7776ABF4C66"
    e = e & "910AAAA4111108BEE5D870EEEFD55552EBD3926FB6222955556AA66C38008ABABD7776ABF4C0E932A8AA4111108BEF65868EEEFD55552EBD37967BA222955556AA66840108AAABD7776ABB4E2EA78AAAA4111108BFEE78FEEEEF"
    e = e & "FD5553EB51FB771A2224555551A2AA410A8AAAEB7776ABB462AE448AAB1111144BFEE6CFEEEEFFD5553EB487B268A2225555548A264450B8AAAD577772BB492AEF8AAA81111168B1866AE2EEEF5D5555AB087B368A2225555548"
    e = e & "A22C440B8AAAD577772BB452AAF8AAA81111168B0A06FE2EEEF5D5555AB0B533E862225555548A2B4042FCAAAD577772BF012AAEAAAA81111168B1A06F6BEEEF5D5555AA0A533DC42225555548AEF4044FC2AAD577772B9652A8"
    e = e & "AA6AA8111116893C03E098EEF5D5555AA125171C62225555548A7B4040FCAAABF7777FB8052BAC6AAA911111C4BFA83E6BEEEFAD5556AB0451F31602E4555571ABA5828FCA2CBF7777FBDD16AEC6A6BA51111A82EAD3E29E8FEF"
    e = e & "55556AF0471B2AA22211555628FEC2298CAAAB57776BA8"
    t = "QQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQ"
    t = t & "QQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQ"
    t = t & "QQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQ"
    Call Record(QRChecksum(t) = e, "—„“ QR ÿÊÌ· («·≈’œ«— 15)")
    Call Record(BuildZatcaQR("Bobs Records", "310122393500003", DateSerial(2022, 4, 25) + TimeSerial(18, 30, 0), 1000, 150) = _
        "AQxCb2JzIFJlY29yZHMCDzMxMDEyMjM5MzUwMDAwMwMUMjAyMi0wNC0yNVQxNTozMDowMFoEBzEwMDAuMDAFBjE1MC4wMA==", _
        "Õ„Ê·… QR ·„À«· «·ÂÌ∆… «·—”„Ì")
    e = "ARnZhdiq2KzYsSDYp9mE2KfYrtiq2KjYp9ixAg8zMDAwMDAwMDAwMDAwMDMDFDIwMjYtMTAtMDRUMDk6MTU6MDBaBAcxMTUwLjAwBQYxNTAuMDA="
    Call Record(BuildZatcaQR("„ Ã— «·«Œ »«—", "300000000000003", DateSerial(2026, 10, 4) + TimeSerial(12, 15, 0), 1150, 150) = e, _
        "Õ„Ê·… QR »«·⁄—»Ì… (UTF-8)")
    CalcReset
    CalcAddLine 1, 11.50, 0, 0.15
    msg = CalcRun(0, True)
    Call Record(Len(msg) = 0 And CalcTotal("SUBTOTAL") = 10.00 And CalcTotal("DISCOUNT") = 0.00 And CalcTotal("TAX") = 1.50 And CalcTotal("TOTAL") = 11.50, "Õ”«»: ”⁄— ‘«„· 11.50 ◊ 1")
    Call Record(CalcLine(0, "UNIT") = 10.0000 And CalcLine(0, "TAX") = 1.50, "√”ÿ—: ”⁄— ‘«„· 11.50 ◊ 1")
    CalcReset
    CalcAddLine 3, 9.99, 0, 0.15
    msg = CalcRun(0, True)
    Call Record(Len(msg) = 0 And CalcTotal("SUBTOTAL") = 26.06 And CalcTotal("DISCOUNT") = 0.00 And CalcTotal("TAX") = 3.91 And CalcTotal("TOTAL") = 29.97, "Õ”«»: ”⁄— ‘«„· 9.99 ◊ 3")
    Call Record(CalcLine(0, "UNIT") = 8.6870 And CalcLine(0, "TAX") = 3.91, "√”ÿ—: ”⁄— ‘«„· 9.99 ◊ 3")
    CalcReset
    CalcAddLine 2, 25.00, 0, 0.15
    CalcAddLine 1, 7.95, 0, 0.15
    CalcAddLine 5, 3.45, 1.00, 0.15
    msg = CalcRun(10, True)
    Call Record(Len(msg) = 0 And CalcTotal("SUBTOTAL") = 65.39 And CalcTotal("DISCOUNT") = 9.56 And CalcTotal("TAX") = 8.37 And CalcTotal("TOTAL") = 64.20, "Õ”«»: À·«À… √”ÿ— ÊŒ’„ ›« Ê—… 10")
    Call Record(CalcLine(0, "UNIT") = 21.7391 And CalcLine(0, "TAX") = 5.64 And CalcLine(1, "UNIT") = 6.9130 And CalcLine(1, "TAX") = 0.90 And CalcLine(2, "UNIT") = 3.0000 And CalcLine(2, "TAX") = 1.83, "√”ÿ—: À·«À… √”ÿ— ÊŒ’„ ›« Ê—… 10")
    CalcReset
    CalcAddLine 0.375, 12.99, 0, 0.15
    msg = CalcRun(0, True)
    Call Record(Len(msg) = 0 And CalcTotal("SUBTOTAL") = 4.24 And CalcTotal("DISCOUNT") = 0.01 And CalcTotal("TAX") = 0.64 And CalcTotal("TOTAL") = 4.87, "Õ”«»: ﬂ„Ì… ﬂ”—Ì… 0.375 ﬂÌ·Ê ◊ 12.99")
    Call Record(CalcLine(0, "UNIT") = 11.2957 And CalcLine(0, "TAX") = 0.64, "√”ÿ—: ﬂ„Ì… ﬂ”—Ì… 0.375 ﬂÌ·Ê ◊ 12.99")
    CalcReset
    CalcAddLine 2, 10.00, 0, 0
    CalcAddLine 1, 23.00, 0, 0.15
    msg = CalcRun(0, True)
    Call Record(Len(msg) = 0 And CalcTotal("SUBTOTAL") = 40.00 And CalcTotal("DISCOUNT") = 0.00 And CalcTotal("TAX") = 3.00 And CalcTotal("TOTAL") = 43.00, "Õ”«»: ’‰› „⁄›Ï „⁄ ’‰› Œ«÷⁄")
    Call Record(CalcLine(0, "UNIT") = 10.0000 And CalcLine(0, "TAX") = 0.00 And CalcLine(1, "UNIT") = 20.0000 And CalcLine(1, "TAX") = 3.00, "√”ÿ—: ’‰› „⁄›Ï „⁄ ’‰› Œ«÷⁄")
    CalcReset
    CalcAddLine 2, 100.00, 5.00, 0.15
    msg = CalcRun(0, False)
    Call Record(Len(msg) = 0 And CalcTotal("SUBTOTAL") = 200.00 And CalcTotal("DISCOUNT") = 5.00 And CalcTotal("TAX") = 29.25 And CalcTotal("TOTAL") = 224.25, "Õ”«»: √”⁄«— €Ì— ‘«„·… 100 ◊ 2 Œ’„ 5")
    Call Record(CalcLine(0, "UNIT") = 100.0000 And CalcLine(0, "TAX") = 29.25, "√”ÿ—: √”⁄«— €Ì— ‘«„·… 100 ◊ 2 Œ’„ 5")
    CalcReset
    CalcAddLine 1, 19.99, 0, 0.15
    CalcAddLine 3, 4.10, 0, 0.15
    msg = CalcRun(3.33, False)
    Call Record(Len(msg) = 0 And CalcTotal("SUBTOTAL") = 32.29 And CalcTotal("DISCOUNT") = 3.33 And CalcTotal("TAX") = 4.34 And CalcTotal("TOTAL") = 33.30, "Õ”«»: €Ì— ‘«„·… „⁄ Œ’„ ›« Ê—… 3.33")
    Call Record(CalcLine(0, "UNIT") = 19.9900 And CalcLine(0, "TAX") = 2.69 And CalcLine(1, "UNIT") = 4.1000 And CalcLine(1, "TAX") = 1.65, "√”ÿ—: €Ì— ‘«„·… „⁄ Œ’„ ›« Ê—… 3.33")
    CalcReset
    CalcAddLine 10, 115.00, 0, 0.15
    msg = CalcRun(0, True)
    Call Record(Len(msg) = 0 And CalcTotal("SUBTOTAL") = 1000.00 And CalcTotal("DISCOUNT") = 0.00 And CalcTotal("TAX") = 150.00 And CalcTotal("TOTAL") = 1150.00, "Õ”«»: ›« Ê—… 1150 ·· Õﬁﬁ „‰ QR")
    Call Record(CalcLine(0, "UNIT") = 100.0000 And CalcLine(0, "TAX") = 150.00, "√”ÿ—: ›« Ê—… 1150 ·· Õﬁﬁ „‰ QR")
    CalcReset
    CalcAddLine 1, 5.00, 5.00, 0.15
    CalcAddLine 1, 2.30, 0, 0.15
    msg = CalcRun(0, True)
    Call Record(Len(msg) = 0 And CalcTotal("SUBTOTAL") = 6.35 And CalcTotal("DISCOUNT") = 4.35 And CalcTotal("TAX") = 0.30 And CalcTotal("TOTAL") = 2.30, "Õ”«»: Œ’„ ÌÃ⁄· «·”ÿ— ’›—«")
    Call Record(CalcLine(0, "UNIT") = 4.3478 And CalcLine(0, "TAX") = 0.00 And CalcLine(1, "UNIT") = 2.0000 And CalcLine(1, "TAX") = 0.30, "√”ÿ—: Œ’„ ÌÃ⁄· «·”ÿ— ’›—«")
    CalcReset
    CalcAddLine 7, 0.05, 0, 0.15
    msg = CalcRun(0, True)
    Call Record(Len(msg) = 0 And CalcTotal("SUBTOTAL") = 0.30 And CalcTotal("DISCOUNT") = 0.00 And CalcTotal("TAX") = 0.05 And CalcTotal("TOTAL") = 0.35, "Õ”«»:  ﬁ—Ì» Õœ¯Ì 0.05 ◊ 7")
    Call Record(CalcLine(0, "UNIT") = 0.0435 And CalcLine(0, "TAX") = 0.05, "√”ÿ—:  ﬁ—Ì» Õœ¯Ì 0.05 ◊ 7")
    Call Record(ZatcaAmount(1000) = "1000.00" And ZatcaAmount(0.5) = "0.50", "’Ì€… «·„»«·€ ›Ì QR")
    Call Record(ZatcaTimestamp(DateSerial(2026, 10, 4) + TimeSerial(1, 5, 9)) = "2026-10-03T22:05:09Z", "«·Êﬁ  ›Ì QR » ÊﬁÌ  UTC")
    Call Record(NewUUID() Like "????????-????-4???-????-????????????", " Ê·Ìœ UUID")
End Sub

'------------------------------------------------------------------------------
' 2) the sales cycle, rolled back at the end
'------------------------------------------------------------------------------
Private Sub TestPosting()
    Dim ws As DAO.Workspace, db As DAO.Database, inTrans As Boolean, integrityBefore As Long
    Dim pid As Long, cid As Long, msg As String, cashInv As Long, creditInv As Long, dummy As Long
    Dim retID As Long, payID As Long, invDate As Date
    On Error GoTo EH
    Set db = CurrentDb
    integrityBefore = DbValue("SELECT COUNT(*) FROM IntegrityCheckQuery")
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True

    db.Execute "UPDATE Settings SET StoreName = 'TEST „ Ã—', VATNumber = '300000000000003', " & _
               "VATRate = 0.15, PricesIncludeVAT = True, AllowNegativeStock = False", dbFailOnError
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    db.Execute "INSERT INTO Products (ProductCode, ProductName, Barcode, PurchasePrice, AverageCost, " & _
               "SellingPrice, CurrentQuantity, MinimumQuantity) VALUES ('TEST-POS-1', 'TEST „‰ Ã', " & _
               "'TESTPOS0001', 60, 60, 115, 0, 5)", dbFailOnError
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-POS-1'")
    ApplyStockMovement db, pid, 100, TT_OPENING, 60, "MANUAL", 0, "TEST-OPEN"
    Call Record(Stock(pid) = 100, "—’Ìœ «›  «ÕÌ 100 ﬁÿ⁄…")
    db.Execute "INSERT INTO Customers (CustomerName, AllowCredit, CreditLimit) VALUES " & _
               "('TEST ⁄„Ì· ¬Ã·', True, 1000)", dbFailOnError
    cid = DbValue("SELECT CustomerID FROM Customers WHERE CustomerName = 'TEST ⁄„Ì· ¬Ã·'")

    ' --- cash sale: 10 x 115 (incl. VAT)
    CartAdd pid, 10
    msg = PostSaleFromCart(1, "CASH", 1, 0, Null, "", cashInv)
    Call Record(Len(msg) = 0 And cashInv > 0, "Õ›Ÿ ›« Ê—… ‰ﬁœÌ… " & msg)
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    Call Record(Stock(pid) = 90, "»⁄œ »Ì⁄ 10 ﬁÿ⁄ √’»Õ «·„Œ“Ê‰ 90")
    Call Record(Inv(cashInv, "TotalAmount") = 1150 And Inv(cashInv, "Tax") = 150 And _
                Inv(cashInv, "TaxableAmount") = 1000, "«·≈Ã„«·Ì 1150 = 1000 + ÷—Ì»… 150")
    Call Record(Inv(cashInv, "PaidAmount") = 1150 And Inv(cashInv, "RemainingAmount") = 0, "„œ›Ê⁄… »«·ﬂ«„·")
    Call Record(Nz(DbValue("SELECT UnitCost FROM SalesInvoiceDetails WHERE SalesInvoiceID = " & cashInv), 0) = 60, _
                "Õ›Ÿ  ﬂ·›… «·’‰› ·ÕŸ… «·»Ì⁄")
    Call Record(Nz(DbValue("SELECT QuantityAfter FROM InventoryTransactions WHERE ReferenceType = 'SALE' " & _
                "AND ReferenceID = " & cashInv), 0) = 90, "Õ—ﬂ… „Œ“Ê‰ »Ì⁄ »«·—’Ìœ »⁄œÂ« 90")
    invDate = DbValue("SELECT InvoiceDate FROM SalesInvoices WHERE SalesInvoiceID = " & cashInv)
    Call Record(DbValue("SELECT QRCodeData FROM SalesInvoices WHERE SalesInvoiceID = " & cashInv) = _
                BuildZatcaQR("TEST „ Ã—", "300000000000003", invDate, 1150, 150), "—„“ QR „Õ›ÊŸ ›Ì «·›« Ê—…")
    Call Record(DbValue("SELECT InvoiceSubType FROM SalesInvoices WHERE SalesInvoiceID = " & cashInv) = "SIMPLIFIED" And _
                Nz(DbValue("SELECT ICV FROM SalesInvoices WHERE SalesInvoiceID = " & cashInv), 0) > 0 And _
                Len(Nz(DbValue("SELECT InvoiceUUID FROM SalesInvoices WHERE SalesInvoiceID = " & cashInv), "")) = 36, _
                "›« Ê—… „»”ÿ… »—ﬁ„ ICV Ê UUID")

    ' --- refusals
    CartAdd pid, 1
    msg = PostSaleFromCart(1, "CASH", 1, 0, 50, "", dummy)
    Call Record(Len(msg) > 0 And Stock(pid) = 90, "—›÷ »Ì⁄ ‰ﬁœÌ »„»·€ „œ›Ê⁄ √ﬁ· „‰ «·≈Ã„«·Ì")
    msg = PostSaleFromCart(1, "CREDIT", 1, 0, 0, "", dummy)
    Call Record(Len(msg) > 0, "—›÷ «·»Ì⁄ «·¬Ã· ··⁄„Ì· «·‰ﬁœÌ")
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartAdd pid, 500
    g_AutoAnswer = False
    msg = PostSaleFromCart(1, "CASH", 1, 0, Null, "", dummy)
    g_AutoAnswer = True
    Call Record(Len(msg) > 0 And Stock(pid) = 90, "„‰⁄ «·»Ì⁄ »√ﬂÀ— „‰ «·„ Ê›— »œÊ‰ „Ê«›ﬁ… «·„œÌ—")
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError

    ' --- credit sale: 2 x 115 = 230, paid 50
    CartAdd pid, 2
    msg = PostSaleFromCart(cid, "CREDIT", 1, 0, 50, "", creditInv)
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    Call Record(Len(msg) = 0 And Inv(creditInv, "RemainingAmount") = 180, "»Ì⁄ ¬Ã·: «·„ »ﬁÌ 180 " & msg)
    Call Record(Balance(cid) = 180, "—’Ìœ «·⁄„Ì· «·¬Ã· 180")
    Call Record(Stock(pid) = 88, "«·„Œ“Ê‰ 88")

    ' --- receipt voucher 100
    msg = PostCustomerPayment(cid, 100, 1, "TEST", payID)
    Call Record(Len(msg) = 0 And Balance(cid) = 80, "”‰œ ﬁ»÷ 100: «·—’Ìœ 80 " & msg)
    Call Record(Len(PostCustomerPayment(1, 10, 1, "", dummy)) > 0, "—›÷ ”‰œ ﬁ»÷ ⁄·Ï «·⁄„Ì· «·‰ﬁœÌ")

    ' --- return 1 unit of the credit sale to stock, credited to the account
    db.Execute "DELETE FROM tmpReturnLines", dbFailOnError
    db.Execute "INSERT INTO tmpReturnLines (SalesDetailID, ProductID, ReturnQty, ReturnToStock) " & _
               "SELECT SalesDetailID, ProductID, 1, True FROM SalesInvoiceDetails WHERE SalesInvoiceID = " & _
               creditInv, dbFailOnError
    Call Record(Len(PostSalesReturn(creditInv, "", "CREDIT", Null, dummy)) > 0, "—›÷ „— Ã⁄ »œÊ‰ ”»»")
    msg = PostSalesReturn(creditInv, "TEST ≈—Ã«⁄", "CREDIT", Null, retID)
    Call Record(Len(msg) = 0 And retID > 0, "Õ›Ÿ ≈‘⁄«— œ«∆‰ " & msg)
    Call Record(Nz(DbValue("SELECT TotalAmount FROM SalesReturns WHERE SalesReturnID = " & retID), 0) = 115 And _
                DbValue("SELECT InvoiceTypeCode FROM SalesReturns WHERE SalesReturnID = " & retID) = "381", _
                "ﬁÌ„… «·„— Ã⁄ 115 Ê‰Ê⁄Â ≈‘⁄«— œ«∆‰ 381")
    Call Record(Stock(pid) = 89 And Balance(cid) = -35, "«·„— Ã⁄ √⁄«œ ﬁÿ⁄… ··„Œ“Ê‰ (89) ÊŒ’„ 115 „‰ «·—’Ìœ (-35)")
    db.Execute "UPDATE tmpReturnLines SET ReturnQty = 5", dbFailOnError
    Call Record(Len(PostSalesReturn(creditInv, "TEST", "CREDIT", Null, dummy)) > 0, _
                "—›÷ ≈—Ã«⁄ √ﬂÀ— „‰ «·ﬂ„Ì… «·„ »ﬁÌ…")

    Call Record(DbValue("SELECT COUNT(*) FROM IntegrityCheckQuery") = integrityBefore, "›Õ’ «·”·«„… ·« ÌÃœ √Ì „‘ﬂ·… ÃœÌœ…")
    Call Record(CurrentDb.OpenRecordset("SELECT COUNT(*) FROM qrySalesDocPrint WHERE DocKind = 'SALE' " & _
                "AND DocID = " & cashInv)(0) = 1, "»Ì«‰«  ÿ»«⁄… «·›« Ê—…")

    ws.Rollback
    inTrans = False
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    db.Execute "DELETE FROM tmpReturnLines", dbFailOnError
    Call Record(DbValue("SELECT COUNT(*) FROM Products WHERE ProductCode = 'TEST-POS-1'") = 0, "«· —«Ã⁄ ⁄‰ ﬂ· »Ì«‰«  «·«Œ »«—")
    Exit Sub
EH:
    Fail "Œÿ√ €Ì— „ Êﬁ⁄ " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If inTrans Then ws.Rollback
End Sub

'------------------------------------------------------------------------------
' 3) the point-of-sale screen (nothing is posted)
'------------------------------------------------------------------------------
Private Sub TestPOSScreen()
    Dim frm As Access.Form, pid As Long
    On Error GoTo EH
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, Barcode, SellingPrice) VALUES " & _
                      "('TEST-POS-2', 'TEST „‰ Ã ‘«‘…', 'TESTPOS0002', 11.5)", dbFailOnError
    pid = DLookup("ProductID", "Products", "ProductCode = 'TEST-POS-2'")
    DoCmd.OpenForm "frmPOS", acNormal, , , , acHidden
    Set frm = Forms("frmPOS")
    Call Record(ScanCode(frm, "TESTPOS0002"), "„”Õ «·»«—ﬂÊœ Ì÷Ì› «·„‰ Ã")
    Call Record(ScanCode(frm, "TESTPOS0002") And DCount("*", "tmpPOSLines") = 1 And _
                DLookup("Quantity", "tmpPOSLines") = 2, "„”Õ ‰›” «·„‰ Ã Ì“Ìœ «·ﬂ„Ì… ›Ì ‰›” «·”ÿ—")
    Call Record(ScanCode(frm, "3*TEST-POS-2") And DLookup("Quantity", "tmpPOSLines") = 5, _
                "«·’Ì€… 3*«·ﬂÊœ  ÷Ì› 3 ﬁÿ⁄ (»«·ﬂÊœ √Ì÷«)")
    Call Record(frm!lblTotal.Caption = Format$(57.5, "#,##0.00"), "«·≈Ã„«·Ì 5 ◊ 11.50 = 57.50")
    Call Record(Not ScanCode(frm, "NO-SUCH-CODE"), "»«—ﬂÊœ €Ì— „ÊÃÊœ Ìı—›÷")
    frm!txtInvoiceDiscount.Value = 7.5
    RecalcPOS frm
    Call Record(frm!lblTotal.Caption = Format$(50, "#,##0.00") And frm!lblTax.Caption = Format$(6.52, "#,##0.00"), _
                "Œ’„ 7.50 ⁄·Ï «·›« Ê—…: «·≈Ã„«·Ì 50.00 Ê«·÷—Ì»… 6.52")
    frm!txtTendered.Value = 100
    RecalcPOS frm
    Call Record(InStr(frm!lblChange.Caption, Format$(50, "#,##0.00")) > 0, "«·»«ﬁÌ ··⁄„Ì· 50.00")
    RemoveCurrentLine frm!subLines.Form
    Call Record(DCount("*", "tmpPOSLines") = 0 And frm!lblTotal.Caption = Format$(0, "#,##0.00"), "Õ–› «·”ÿ—")
    DoCmd.Close acForm, "frmPOS", acSaveNo
    CurrentDb.Execute "DELETE FROM Products WHERE ProductID = " & pid, dbFailOnError
    Exit Sub
EH:
    Fail "‘«‘… «·»Ì⁄: Œÿ√ " & Err.Number & ": " & Err.Description
    On Error Resume Next
    DoCmd.Close acForm, "frmPOS", acSaveNo
    CurrentDb.Execute "DELETE FROM tmpPOSLines"
    CurrentDb.Execute "DELETE FROM Products WHERE ProductCode = 'TEST-POS-2'"
End Sub

'------------------------------------------------------------------------------
' helpers
'------------------------------------------------------------------------------
Private Sub CartAdd(ByVal ProductID As Long, ByVal Qty As Currency)
    CurrentDb.Execute "INSERT INTO tmpPOSLines (ProductID, ProductCode, ProductName, Quantity, UnitPrice, " & _
        "LineDiscount, Available) SELECT ProductID, ProductCode, ProductName, " & Trim$(Str$(Qty)) & _
        ", SellingPrice, 0, CurrentQuantity FROM Products WHERE ProductID = " & ProductID, dbFailOnError
End Sub

Private Function Stock(ByVal ProductID As Long) As Currency
    Stock = Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & ProductID), -1)
End Function

Private Function Balance(ByVal CustomerID As Long) As Currency
    Balance = Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerID = " & CustomerID), -9999)
End Function

Private Function Inv(ByVal InvoiceID As Long, ByVal FieldName As String) As Currency
    Inv = Nz(DbValue("SELECT " & FieldName & " FROM SalesInvoices WHERE SalesInvoiceID = " & InvoiceID), -1)
End Function

Private Sub Record(ByVal Passed As Boolean, ByVal Label As String)
    If Passed Then
        m_passed = m_passed + 1
        Debug.Print "[OK] " & Label
    Else
        Fail Label
    End If
End Sub

Private Sub Fail(ByVal Msg As String)
    m_failed = m_failed + 1
    m_report = m_report & "- " & Msg & vbCrLf
    Debug.Print "[X] " & Msg
End Sub
