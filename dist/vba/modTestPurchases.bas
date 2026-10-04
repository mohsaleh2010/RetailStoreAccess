Attribute VB_Name = "modTestPurchases"
'==============================================================================
' modTestPurchases  -  Retail Store Management System (Phase 7 tests)
'
' GENERATED FILE - do not edit by hand.  Source: tools/gen_test_purchases.py
'
'   TestPurchases  1) average-cost vectors (same as tools/pricing.py)
'                  2) the purchase and inventory cycle inside a transaction that
'                     is rolled back: new supplier -> buy 100 -> stock 100 ->
'                     sell 10 -> stock 90 (spec), second purchase at a new cost,
'                     purchase return, payment voucher, opening balance, stock-out,
'                     stocktake, and every refusal (duplicate supplier invoice,
'                     unpaid cash purchase, future date, over-return, ...)
'                  3) the purchase and inventory screens (nothing is posted)
'==============================================================================
Option Compare Database
Option Explicit

Private m_passed As Long
Private m_failed As Long
Private m_report As String

Public Function TestPurchases() As Boolean
    m_passed = 0: m_failed = 0: m_report = ""
    Debug.Print "=== TestPurchases  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    g_AutoAnswer = True
    EnsureLocalTables

    TestAverageVectors
    TestCycle
    TestPurchaseScreens

    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & m_passed & " | ›‘·: " & m_failed
    If m_failed = 0 Then
        MsgBox "Ã„Ì⁄ «Œ »«—«  «·„‘ —Ì«  Ê«·„Œ“Ê‰ ‰«ÃÕ… (" & m_passed & " «Œ »«—«)." & vbCrLf & _
               "·„  ı —ﬂ √Ì »Ì«‰«  «Œ »«—.", vbInformation + MSG_RTL, "TestPurchases"
        TestPurchases = True
    Else
        MsgBox "‰ÃÕ " & m_passed & " Ê›‘· " & m_failed & ":" & vbCrLf & vbCrLf & Left$(m_report, 900), _
               vbExclamation + MSG_RTL, "TestPurchases"
    End If
End Function

'------------------------------------------------------------------------------
' 1) vectors
'------------------------------------------------------------------------------
Private Sub TestAverageVectors()
    Call Record(WeightedAverage(0, 0, 100, 52.1739) = CCur(52.1739), "„ Ê”ÿ «· ﬂ·›…: √Ê· ‘—«¡ »œÊ‰ —’Ìœ")
    Call Record(WeightedAverage(90, 50, 10, 60) = CCur(51), "„ Ê”ÿ «· ﬂ·›…: ‘—«¡ À«‰Ú »”⁄— √⁄·Ï")
    Call Record(WeightedAverage(3, 10, 7, 10.3333) = CCur(10.2333), "„ Ê”ÿ «· ﬂ·›…: ﬂ”Ê—  Õ «Ã  ﬁ—Ì» 4 Œ«‰« ")
    Call Record(WeightedAverage(-5, 40, 20, 45) = CCur(45), "„ Ê”ÿ «· ﬂ·›…: ‘—«¡ »⁄œ —’Ìœ ”«·»")
    Call Record(WeightedAverage(100, 55, -10, 60) = CCur(54.4444), "„ Ê”ÿ «· ﬂ·›…: „— Ã⁄ ‘—«¡ » ﬂ·›… «·‘—«¡")
    Call Record(WeightedAverage(10, 55, -10, 60) = CCur(55), "„ Ê”ÿ «· ﬂ·›…: „— Ã⁄ Ìı›—€ «·„Œ“Ê‰ Ì»ﬁÌ «·„ Ê”ÿ")
    Call Record(WeightedAverage(2, 1, -1, 5) = CCur(1), "„ Ê”ÿ «· ﬂ·›…: „— Ã⁄ Ì‰ Ã „ Ê”ÿ« ”«·»« Ì»ﬁÌ «·„ Ê”ÿ")
    Call Record(WeightedAverage(10, 55, 0, 99) = CCur(55), "„ Ê”ÿ «· ﬂ·›…: Õ—ﬂ… ’›—Ì…")
    Call Record(PurchaseUnitCost(100, 115, 3, True) = CCur(33.3333), " ﬂ·›… «·ÊÕœ… 100/115 ˜ 3 „”Ã· »«·÷—Ì»…")
    Call Record(PurchaseUnitCost(100, 115, 3, False) = CCur(38.3333), " ﬂ·›… «·ÊÕœ… 100/115 ˜ 3 €Ì— „”Ã·")
    Call Record(PurchaseUnitCost(0, 0, 0, True) = CCur(0), " ﬂ·›… «·ÊÕœ… 0/0 ˜ 0 „”Ã· »«·÷—Ì»…")
End Sub

'------------------------------------------------------------------------------
' 2) the purchase and inventory cycle, rolled back at the end
'------------------------------------------------------------------------------
Private Sub TestCycle()
    Dim ws As DAO.Workspace, db As DAO.Database, inTrans As Boolean, integrityBefore As Long
    Dim sid As Long, cat As Long, pa As Long, pb As Long, msg As String, inv1 As Long, inv2 As Long
    Dim saleID As Long, dummy As Long, retID As Long, payID As Long, refNo As String
    Dim countID As Long, count2 As Long, adjusted As Long, netValue As Currency
    On Error GoTo EH
    Set db = CurrentDb
    integrityBefore = DbValue("SELECT COUNT(*) FROM IntegrityCheckQuery")
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True

    db.Execute "UPDATE Settings SET VATNumber = '300000000000003', VATRate = 0.15, PricesIncludeVAT = True, " & _
               "AllowNegativeStock = False", dbFailOnError
    db.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    db.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    ' an open stocktake of the user blocks a new one: closed here, restored by the rollback
    db.Execute "UPDATE StockCounts SET Status = 'CANCELLED' WHERE Status = 'OPEN'", dbFailOnError

    ' --- 1. new supplier and products
    db.Execute "INSERT INTO Suppliers (SupplierName, VATNumber) VALUES ('TEST „Ê—œ «·„‘ —Ì« ', '300000000000003')", _
               dbFailOnError
    sid = DbValue("SELECT SupplierID FROM Suppliers WHERE SupplierName = 'TEST „Ê—œ «·„‘ —Ì« '")
    db.Execute "INSERT INTO Categories (CategoryName) VALUES ('TEST  ’‰Ì› «·Ã—œ')", dbFailOnError
    cat = DbValue("SELECT CategoryID FROM Categories WHERE CategoryName = 'TEST  ’‰Ì› «·Ã—œ'")
    db.Execute "INSERT INTO Products (ProductCode, ProductName, Barcode, CategoryID, SellingPrice, MinimumQuantity) " & _
               "VALUES ('TEST-PUR-A', 'TEST „‰ Ã √', 'TESTPUR000A', " & cat & ", 115, 5)", dbFailOnError
    pa = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-PUR-A'")
    db.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, SellingPrice) " & _
               "VALUES ('TEST-PUR-B', 'TEST „‰ Ã »', " & cat & ", 23)", dbFailOnError
    pb = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-PUR-B'")
    Call Record(Stock(pa) = 0, "„‰ Ã ÃœÌœ »œÊ‰ —’Ìœ")

    ' --- 2. purchase 1: 100 x 50 + VAT, credit, 1000 paid now
    PurCartAdd pa, 100, 50
    msg = PostPurchaseFromCart(sid, "TEST-S-1", Date, "CREDIT", 1, True, 0, 1000, "", inv1)
    db.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    Call Record(Len(msg) = 0 And inv1 > 0, "Õ›Ÿ ›« Ê—… ‘—«¡ ¬Ã·… " & msg)
    Call Record(PurInv(inv1, "TotalAmount") = 5750 And PurInv(inv1, "Tax") = 750 And _
                PurInv(inv1, "RemainingAmount") = 4750, "«·≈Ã„«·Ì 5750 Ê«·÷—Ì»… 750 Ê«·„ »ﬁÌ 4750")
    Call Record(Stock(pa) = 100, "»⁄œ ‘—«¡ 100 √’»Õ «·„Œ“Ê‰ 100")
    Call Record(AvgCost(pa) = 50, "„ Ê”ÿ «· ﬂ·›… 50 (»œÊ‰ «·÷—Ì»… «·ﬁ«»·… ··«” —œ«œ)")
    Call Record(SupBalance(sid) = 4750, "—’Ìœ «·„Ê—œ = «·„ »ﬁÌ „‰ «·›« Ê—…")
    Call Record(DbValue("SELECT PurchasePrice FROM Products WHERE ProductID = " & pa) = 50 And _
                DbValue("SELECT SupplierID FROM Products WHERE ProductID = " & pa) = sid, "¬Œ— ”⁄— ‘—«¡ Ê«·„Ê—œ «·«› —«÷Ì")
    Call Record(Nz(DbValue("SELECT QuantityAfter FROM InventoryTransactions WHERE ReferenceType = 'PURCHASE' " & _
                "AND ReferenceID = " & inv1), 0) = 100, "Õ—ﬂ… „Œ“Ê‰ ‘—«¡ »«·—’Ìœ »⁄œÂ«")

    ' --- refusals (nothing may change)
    PurCartAdd pa, 1, 50
    Call Record(Len(PostPurchaseFromCart(sid, "TEST-S-1", Date, "CREDIT", 1, True, 0, 0, "", dummy)) > 0, _
                "—›÷  ﬂ—«— —ﬁ„ ›« Ê—… «·„Ê—œ ·‰›” «·„Ê—œ")
    Call Record(Len(PostPurchaseFromCart(sid, "", Date, "CASH", 1, True, 0, 10, "", dummy)) > 0, _
                "—›÷ ‘—«¡ ‰ﬁœÌ »„»·€ „œ›Ê⁄ √ﬁ· „‰ «·≈Ã„«·Ì")
    Call Record(Len(PostPurchaseFromCart(sid, "", Date + 1, "CASH", 1, True, 0, Null, "", dummy)) > 0, _
                "—›÷  «—ÌŒ ›« Ê—… ›Ì «·„” ﬁ»·")
    Call Record(Len(PostPurchaseFromCart(0, "", Date, "CASH", 1, True, 0, Null, "", dummy)) > 0, "—›÷ ›« Ê—… »œÊ‰ „Ê—œ")
    db.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    Call Record(Len(PostPurchaseFromCart(sid, "", Date, "CASH", 1, True, 0, Null, "", dummy)) > 0, _
                "—›÷ ›« Ê—… »œÊ‰ √’‰«›")
    Call Record(Stock(pa) = 100 And SupBalance(sid) = 4750, "«·—›÷ ·„ Ì€Ì— «·„Œ“Ê‰ Ê·« «·—’Ìœ")

    ' --- 3. the spec scenario: sell 10 -> 90
    CurrentDb.Execute "INSERT INTO tmpPOSLines (ProductID, ProductCode, ProductName, Quantity, UnitPrice, " & _
        "LineDiscount, Available) SELECT ProductID, ProductCode, ProductName, 10, SellingPrice, 0, " & _
        "CurrentQuantity FROM Products WHERE ProductID = " & pa, dbFailOnError
    msg = PostSaleFromCart(1, "CASH", 1, 0, Null, "", saleID)
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    Call Record(Len(msg) = 0 And Stock(pa) = 90, "»Ì⁄ 10: «·„Œ“Ê‰ 90 " & msg)
    Call Record(Nz(DbValue("SELECT UnitCost FROM SalesInvoiceDetails WHERE SalesInvoiceID = " & saleID), 0) = 50, _
                " ﬂ·›… «·»Ì⁄ = „ Ê”ÿ  ﬂ·›… «·‘—«¡")

    ' --- 4. purchase 2: 10 x 60, cash, new selling price
    PurCartAdd pa, 10, 60
    db.Execute "UPDATE tmpPurchaseLines SET NewSellingPrice = 120", dbFailOnError
    msg = PostPurchaseFromCart(sid, "TEST-S-2", Date, "CASH", 1, True, 0, Null, "", inv2)
    db.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    Call Record(Len(msg) = 0 And PurInv(inv2, "TotalAmount") = 690 And PurInv(inv2, "RemainingAmount") = 0, _
                "‘—«¡ ‰ﬁœÌ 690 „œ›Ê⁄ »«·ﬂ«„· " & msg)
    Call Record(Stock(pa) = 100 And AvgCost(pa) = 51, "«·„Œ“Ê‰ 100 Ê„ Ê”ÿ «· ﬂ·›… «·„—ÃÕ 51")
    Call Record(DbValue("SELECT SellingPrice FROM Products WHERE ProductID = " & pa) = 120, _
                " ÕœÌÀ ”⁄— «·»Ì⁄ „‰ ›« Ê—… «·‘—«¡")
    Call Record(SupBalance(sid) = 4750, "«·‘—«¡ «·‰ﬁœÌ ·« Ì€Ì— —’Ìœ «·„Ê—œ")

    ' --- 5. purchase return: 10 units of purchase 1, deducted from the account
    db.Execute "INSERT INTO tmpPurchaseReturnLines (PurchaseDetailID, ProductID, ReturnQty) SELECT PurchaseDetailID, " & _
               "ProductID, 10 FROM PurchaseInvoiceDetails WHERE PurchaseInvoiceID = " & inv1, dbFailOnError
    Call Record(Len(PostPurchaseReturn(inv1, "", "CREDIT", Null, dummy)) > 0, "—›÷ „— Ã⁄ »œÊ‰ ”»»")
    msg = PostPurchaseReturn(inv1, "TEST ⁄Ì» „’‰⁄Ì", "CREDIT", Null, retID)
    Call Record(Len(msg) = 0 And retID > 0, "Õ›Ÿ „— Ã⁄ «·„‘ —Ì«  " & msg)
    Call Record(Nz(DbValue("SELECT TotalAmount FROM PurchaseReturns WHERE PurchaseReturnID = " & retID), 0) = 575 And _
                Nz(DbValue("SELECT Tax FROM PurchaseReturns WHERE PurchaseReturnID = " & retID), 0) = 75, _
                "ﬁÌ„… «·„— Ã⁄ 575 „‰Â« ÷—Ì»… 75")
    Call Record(Stock(pa) = 90 And AvgCost(pa) = 51.1111, "«·„— Ã⁄ √Œ—Ã 10 » ﬂ·›… ‘—«∆Â«: «·„ Ê”ÿ 51.1111")
    Call Record(SupBalance(sid) = 4175, "«·„— Ã⁄ Œı’„ „‰ —’Ìœ «·„Ê—œ: 4175")
    db.Execute "UPDATE tmpPurchaseReturnLines SET ReturnQty = 95", dbFailOnError
    Call Record(Len(PostPurchaseReturn(inv1, "TEST", "CREDIT", Null, dummy)) > 0 And Stock(pa) = 90, _
                "—›÷ ≈—Ã«⁄ √ﬂÀ— „‰ «·„ »ﬁÌ „‰ «·›« Ê—…")
    db.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError

    ' --- 6. payment voucher
    Call Record(Len(PostSupplierPayment(sid, 0, 1, "", dummy)) > 0, "—›÷ ”‰œ ’—› »„»·€ ’›—")
    msg = PostSupplierPayment(sid, 1000, 1, "TEST", payID)
    Call Record(Len(msg) = 0 And SupBalance(sid) = 3175, "”‰œ ’—› 1000: «·—’Ìœ 3175 " & msg)

    ' --- 7. manual movements on product B
    Call Record(Len(PostManualStock(pa, TT_OPENING, 5, 10, "", refNo)) > 0, "—›÷ —’Ìœ «›  «ÕÌ ·„‰ Ã ·Â Õ—ﬂ« ")
    msg = PostManualStock(pb, TT_OPENING, 20, 10, "", refNo)
    Call Record(Len(msg) = 0 And Len(refNo) > 0 And Stock(pb) = 20 And AvgCost(pb) = 10, _
                "—’Ìœ «›  «ÕÌ 20 » ﬂ·›… 10 " & msg)
    Call Record(Len(PostManualStock(pb, TT_STOCK_IN, 1, 10, "", refNo)) > 0, "—›÷ ≈÷«›… „Œ“Ê‰ »œÊ‰ ”»»")
    Call Record(Len(PostManualStock(pb, TT_SALE, 1, Null, "TEST", refNo)) > 0, "—›÷ ‰Ê⁄ Õ—ﬂ… €Ì— ÌœÊÌ")
    msg = PostManualStock(pb, TT_STOCK_OUT, 2, Null, "TEST  «·›", refNo)
    Call Record(Len(msg) = 0 And Stock(pb) = 18, "Œ’„  «·› 2: «·„Œ“Ê‰ 18 " & msg)
    g_AutoAnswer = False
    Call Record(Len(PostManualStock(pb, TT_STOCK_OUT, 100, Null, "TEST", refNo)) > 0 And Stock(pb) = 18, _
                "„‰⁄ Œ’„ √ﬂÀ— „‰ «·„ Ê›— »œÊ‰ „Ê«›ﬁ… «·„œÌ—")
    g_AutoAnswer = True

    ' --- 8. stocktake of the test category
    msg = CreateStockCount(cat, "TEST", countID)
    Call Record(Len(msg) = 0 And countID > 0, "≈‰‘«¡ Ã—œ ·· ’‰Ì› " & msg)
    Call Record(DbValue("SELECT COUNT(*) FROM StockCountDetails WHERE StockCountID = " & countID) = 2 And _
                DbValue("SELECT SystemQuantity FROM StockCountDetails WHERE StockCountID = " & countID & _
                " AND ProductID = " & pa) = 90, "«·Ã—œ Ì÷„ „‰ ÃÌ «· ’‰Ì› »«·ﬂ„Ì… «·„”Ã·…")
    Call Record(Len(CreateStockCount(cat, "", dummy)) > 0, "„‰⁄ › Õ Ã—œÌ‰ ›Ì ‰›” «·Êﬁ ")
    db.Execute "UPDATE StockCountDetails SET ActualQuantity = 88 WHERE StockCountID = " & countID & _
               " AND ProductID = " & pa, dbFailOnError
    db.Execute "UPDATE StockCountDetails SET ActualQuantity = 18 WHERE StockCountID = " & countID & _
               " AND ProductID = " & pb, dbFailOnError
    msg = PostStockCount(countID, False, adjusted, netValue)
    Call Record(Len(msg) = 0 And adjusted = 1 And netValue = -102.22, _
                " —ÕÌ· «·Ã—œ:  ”ÊÌ… Ê«Õœ… »ﬁÌ„… -102.22 " & msg)
    Call Record(Stock(pa) = 88 And Stock(pb) = 18, "«·„Œ“Ê‰ »⁄œ «·Ã—œ 88 Ê 18")
    Call Record(Nz(DbValue("SELECT Quantity FROM InventoryTransactions WHERE ReferenceType = 'STOCK_COUNT' " & _
                "AND ReferenceID = " & countID), 0) = -2, "Õ—ﬂ…  ”ÊÌ… Ã—œ »«·›—ﬁ")
    Call Record(DbValue("SELECT Status FROM StockCounts WHERE StockCountID = " & countID) = "POSTED", "Õ«·… «·Ã—œ „ı—Õ¯·")
    Call Record(Len(PostStockCount(countID, False, adjusted, netValue)) > 0, "—›÷  —ÕÌ· Ã—œ „ı—Õ¯· „—… À«‰Ì…")
    Call Record(Len(CancelStockCount(countID)) > 0, "—›÷ ≈·€«¡ Ã—œ „ı—Õ¯·")
    msg = CreateStockCount(cat, "", count2)
    Call Record(Len(msg) = 0 And Len(CancelStockCount(count2)) = 0 And _
                DbValue("SELECT Status FROM StockCounts WHERE StockCountID = " & count2) = "CANCELLED" And _
                Stock(pa) = 88, "≈·€«¡ Ã—œ „› ÊÕ ·« Ì€Ì— «·„Œ“Ê‰")

    ' --- the saved queries agree with the posted documents
    Call Record(DbValue("SELECT Balance FROM SupplierBalanceQuery WHERE SupplierID = " & sid) = 3175, _
                "—’Ìœ «·„Ê—œ „‰ ﬂ‘› «·Õ”«» = 3175")
    Call Record(DbValue("SELECT COUNT(*) FROM StockBalanceQuery WHERE QuantityMismatch <> 0 AND ProductID IN (" & _
                pa & ", " & pb & ")") = 0, "«·ﬂ„Ì«   ÿ«»ﬁ œ› — Õ—ﬂ… «·„Œ“Ê‰")
    Call Record(DbValue("SELECT COUNT(*) FROM IntegrityCheckQuery") = integrityBefore, "›Õ’ «·”·«„… ·« ÌÃœ √Ì „‘ﬂ·… ÃœÌœ…")

    ws.Rollback
    inTrans = False
    db.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    db.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    Call Record(DbValue("SELECT COUNT(*) FROM Suppliers WHERE SupplierName = 'TEST „Ê—œ «·„‘ —Ì« '") = 0, _
                "«· —«Ã⁄ ⁄‰ ﬂ· »Ì«‰«  «·«Œ »«—")
    Exit Sub
EH:
    Fail "Œÿ√ €Ì— „ Êﬁ⁄ " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If inTrans Then ws.Rollback
End Sub

'------------------------------------------------------------------------------
' 3) the screens (nothing is posted)
'------------------------------------------------------------------------------
Private Sub TestPurchaseScreens()
    Dim frm As Access.Form, pid As Long
    On Error GoTo EH
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, Barcode, PurchasePrice, SellingPrice) VALUES " & _
                      "('TEST-PUR-C', 'TEST „‰ Ã ‘«‘… «·„‘ —Ì« ', 'TESTPUR000C', 40, 55)", dbFailOnError
    pid = DLookup("ProductID", "Products", "ProductCode = 'TEST-PUR-C'")
    DoCmd.OpenForm "frmPurchaseInvoice", acNormal, , , , acHidden
    Set frm = Forms("frmPurchaseInvoice")
    Call Record(PurScanCode(frm, "TESTPUR000C") And DLookup("UnitCost", "tmpPurchaseLines") = 40, _
                "„”Õ «·»«—ﬂÊœ Ì÷Ì› «·„‰ Ã »¬Œ— ”⁄— ‘—«¡")
    Call Record(PurScanCode(frm, "2*TEST-PUR-C") And DCount("*", "tmpPurchaseLines") = 1 And _
                DLookup("Quantity", "tmpPurchaseLines") = 3, "«·’Ì€… 2*«·ﬂÊœ  “Ìœ «·ﬂ„Ì… ›Ì ‰›” «·”ÿ—")
    Call Record(Not PurScanCode(frm, "NO-SUCH-CODE"), "»«—ﬂÊœ €Ì— „ÊÃÊœ Ìı—›÷")
    frm!chkChargeVAT.Value = True
    RecalcPurchase frm
    Call Record(frm!lblTotal.Caption = Format$(138, "#,##0.00") And frm!lblTax.Caption = Format$(18, "#,##0.00"), _
                "3 ◊ 40 + ÷—Ì»… 15% = 138.00")
    frm!chkChargeVAT.Value = False
    RecalcPurchase frm
    Call Record(frm!lblTotal.Caption = Format$(120, "#,##0.00"), "„Ê—œ €Ì— „”Ã· »«·÷—Ì»…: 120.00")
    PurRemoveLine frm!subLines.Form
    Call Record(DCount("*", "tmpPurchaseLines") = 0, "Õ–› «·”ÿ—")
    DoCmd.Close acForm, "frmPurchaseInvoice", acSaveNo

    DoCmd.OpenForm "frmInventory", acNormal, , , , acHidden
    Set frm = Forms("frmInventory")
    frm!txtSearch.Value = "TEST-PUR-C"
    InventoryRefresh frm
    Call Record(frm!lstProducts.ListCount = 2, "«·»ÕÀ ›Ì ‘«‘… «·„Œ“Ê‰ (⁄‰Ê«‰ + „‰ Ã)")
    DoCmd.Close acForm, "frmInventory", acSaveNo
    DoCmd.OpenForm "frmStockCount", acNormal, , , , acHidden
    Call Record(IsFormOpen("frmStockCount"), "› Õ ‘«‘… «·Ã—œ")
    DoCmd.Close acForm, "frmStockCount", acSaveNo
    CurrentDb.Execute "DELETE FROM Products WHERE ProductID = " & pid, dbFailOnError
    Exit Sub
EH:
    Fail "‘«‘«  «·„‘ —Ì« : Œÿ√ " & Err.Number & ": " & Err.Description
    On Error Resume Next
    DoCmd.Close acForm, "frmPurchaseInvoice", acSaveNo
    DoCmd.Close acForm, "frmInventory", acSaveNo
    DoCmd.Close acForm, "frmStockCount", acSaveNo
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines"
    CurrentDb.Execute "DELETE FROM Products WHERE ProductCode = 'TEST-PUR-C'"
End Sub

'------------------------------------------------------------------------------
' helpers
'------------------------------------------------------------------------------
Private Sub PurCartAdd(ByVal ProductID As Long, ByVal Qty As Currency, ByVal Cost As Currency)
    CurrentDb.Execute "INSERT INTO tmpPurchaseLines (ProductID, ProductCode, ProductName, Quantity, UnitCost, " & _
        "LineDiscount, SellingPrice) SELECT ProductID, ProductCode, ProductName, " & Trim$(Str$(Qty)) & ", " & _
        Trim$(Str$(Cost)) & ", 0, SellingPrice FROM Products WHERE ProductID = " & ProductID, dbFailOnError
End Sub

Private Function Stock(ByVal ProductID As Long) As Currency
    Stock = Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & ProductID), -1)
End Function

Private Function AvgCost(ByVal ProductID As Long) As Currency
    AvgCost = Nz(DbValue("SELECT AverageCost FROM Products WHERE ProductID = " & ProductID), -1)
End Function

Private Function SupBalance(ByVal SupplierID As Long) As Currency
    SupBalance = Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierID = " & SupplierID), -9999)
End Function

Private Function PurInv(ByVal InvoiceID As Long, ByVal FieldName As String) As Currency
    PurInv = Nz(DbValue("SELECT " & FieldName & " FROM PurchaseInvoices WHERE PurchaseInvoiceID = " & InvoiceID), -1)
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
