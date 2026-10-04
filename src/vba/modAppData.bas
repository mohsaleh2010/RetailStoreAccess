Attribute VB_Name = "modAppData"
'==============================================================================
' modAppData  -  Retail Store Management System (Phase 5)
'
' GENERATED FILE - do not edit by hand.  Source: tools/forms.py
' Runtime data used by modScreens: search SQL templates and the report catalogue.
' Search tokens: {LIKE} quoted Like pattern, {NUM} record number or -1,
'                {FROM} / {TO} date literals (TO is exclusive).
'==============================================================================
Option Compare Database
Option Explicit

Public Function SearchTemplate(ByVal Kind As String) As String
    Dim s As String
    Select Case Kind
        Case "PRODUCT"
            s = "SELECT p.ProductID, p.ProductCode AS [الكود], p.Barcode AS [الباركود], p.ProductName AS [المنتج], p.CurrentQuantity AS [الكمية], p.SellingPrice AS "
            s = s & "[السعر] FROM Products AS p WHERE p.ProductName Like {LIKE} OR p.ProductCode Like {LIKE} OR p.Barcode Like {LIKE} OR p.ProductNameEn Like {LIKE} OR "
            s = s & "p.ProductID = {NUM} ORDER BY p.ProductName"
        Case "CUSTOMER"
            s = "SELECT c.CustomerID, c.CustomerName AS [العميل], c.Mobile AS [الجوال], c.VATNumber AS [الرقم الضريبي], c.CurrentBalance AS [الرصيد] FROM Customers AS "
            s = s & "c WHERE c.CustomerName Like {LIKE} OR c.Mobile Like {LIKE} OR c.Phone Like {LIKE} OR c.VATNumber Like {LIKE} OR c.CustomerID = {NUM} ORDER BY "
            s = s & "c.CustomerName"
        Case "SUPPLIER"
            s = "SELECT s.SupplierID, s.SupplierName AS [المورد], s.ContactPerson AS [المسؤول], s.Mobile AS [الجوال], s.CurrentBalance AS [الرصيد] FROM Suppliers AS s "
            s = s & "WHERE s.SupplierName Like {LIKE} OR s.ContactPerson Like {LIKE} OR s.Mobile Like {LIKE} OR s.VATNumber Like {LIKE} OR s.SupplierID = {NUM} ORDER BY "
            s = s & "s.SupplierName"
        Case "SALE"
            s = "SELECT h.SalesInvoiceID, h.InvoiceNumber AS [رقم الفاتورة], h.InvoiceDate AS [التاريخ], c.CustomerName AS [العميل], h.TotalAmount AS [الإجمالي], "
            s = s & "h.RemainingAmount AS [المتبقي] FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID WHERE (h.InvoiceNumber Like {LIKE} OR "
            s = s & "c.CustomerName Like {LIKE} OR c.Mobile Like {LIKE}) AND h.InvoiceDate >= {FROM} AND h.InvoiceDate < {TO} ORDER BY h.InvoiceDate DESC"
        Case "PURCHASE"
            s = "SELECT h.PurchaseInvoiceID, h.InvoiceNumber AS [رقم الفاتورة], h.SupplierInvoiceNo AS [فاتورة المورد], h.InvoiceDate AS [التاريخ], s.SupplierName AS "
            s = s & "[المورد], h.TotalAmount AS [الإجمالي] FROM PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID WHERE (h.InvoiceNumber Like "
            s = s & "{LIKE} OR h.SupplierInvoiceNo Like {LIKE} OR s.SupplierName Like {LIKE}) AND h.InvoiceDate >= {FROM} AND h.InvoiceDate < {TO} ORDER BY h.InvoiceDate "
            s = s & "DESC"
    End Select
    SearchTemplate = s
End Function

Public Function SearchColumnCount(ByVal Kind As String) As Integer
    Select Case Kind
        Case "PRODUCT": SearchColumnCount = 6
        Case "CUSTOMER": SearchColumnCount = 5
        Case "SUPPLIER": SearchColumnCount = 5
        Case "SALE": SearchColumnCount = 6
        Case "PURCHASE": SearchColumnCount = 6
    End Select
End Function

Public Function SearchColumnWidths(ByVal Kind As String) As String
    Select Case Kind
        Case "PRODUCT": SearchColumnWidths = "0;1418;1984;5103;1418;1418"
        Case "CUSTOMER": SearchColumnWidths = "0;5103;1984;2552;1701"
        Case "SUPPLIER": SearchColumnWidths = "0;4536;2835;1984;1701"
        Case "SALE": SearchColumnWidths = "0;1984;2268;4536;1701;1701"
        Case "PURCHASE": SearchColumnWidths = "0;1984;1984;2268;3969;1701"
    End Select
End Function

Public Function ReportCount() As Long
    ReportCount = 21
End Function

Public Function ReportRow(ByVal Index As Long) As Variant
    ' Array(Key, Title, QueryName, ReportName, Needs, DateColumn)
    Select Case Index
        Case 1: ReportRow = Array("DAILY_SALES", "المبيعات اليومية", "DailySalesQuery", "rptDailySales", "D", "SaleDate")
        Case 2: ReportRow = Array("MONTHLY_SALES", "المبيعات الشهرية", "MonthlySalesQuery", "rptMonthlySales", "", "")
        Case 3: ReportRow = Array("SALES_PERIOD", "المبيعات حسب فترة", "SalesByPeriodQuery", "rptSalesByPeriod", "Pc", "")
        Case 4: ReportRow = Array("SALES_PRODUCT", "المبيعات حسب المنتج", "SalesByProductQuery", "rptSalesByProduct", "Pr", "")
        Case 5: ReportRow = Array("BEST_SELLING", "أفضل المنتجات مبيعًا", "BestSellingProductsQuery", "rptBestSelling", "P", "")
        Case 6: ReportRow = Array("LEAST_SELLING", "أقل المنتجات مبيعًا", "LeastSellingProductsQuery", "rptLeastSelling", "P", "")
        Case 7: ReportRow = Array("PURCHASES", "المشتريات", "PurchasesQuery", "rptPurchases", "Ps", "")
        Case 8: ReportRow = Array("STOCK", "المخزون الحالي", "StockBalanceQuery", "rptStockBalance", "", "")
        Case 9: ReportRow = Array("LOW_STOCK", "المنتجات منخفضة المخزون", "LowStockQuery", "rptLowStock", "", "")
        Case 10: ReportRow = Array("PRODUCT_MOVEMENT", "حركة منتج", "ProductMovementQuery", "rptProductMovement", "PR", "")
        Case 11: ReportRow = Array("CUSTOMER_STATEMENT", "كشف حساب عميل", "CustomerStatementQuery", "rptCustomerStatement", "PC", "")
        Case 12: ReportRow = Array("SUPPLIER_STATEMENT", "كشف حساب مورد", "SupplierStatementQuery", "rptSupplierStatement", "PS", "")
        Case 13: ReportRow = Array("EXPENSES", "المصروفات", "ExpensesQuery", "rptExpenses", "P", "")
        Case 14: ReportRow = Array("EXPENSES_BY_TYPE", "المصروفات حسب النوع", "ExpensesByTypeQuery", "rptExpensesByType", "P", "")
        Case 15: ReportRow = Array("PROFIT", "الأرباح", "ProfitQuery", "rptProfit", "P", "")
        Case 16: ReportRow = Array("SLOW_MOVING", "المنتجات غير المتحركة", "SlowMovingProductsQuery", "rptSlowMoving", "", "")
        Case 17: ReportRow = Array("STOCK_BY_CATEGORY", "المخزون حسب التصنيف", "StockByCategoryQuery", "rptStockByCategory", "", "")
        Case 18: ReportRow = Array("VAT_SUMMARY", "ملخص ضريبة القيمة المضافة", "VatSummaryQuery", "rptVatSummary", "P", "")
        Case 19: ReportRow = Array("CUSTOMER_BALANCES", "أرصدة العملاء", "CustomerBalanceQuery", "rptCustomerBalances", "", "")
        Case 20: ReportRow = Array("SUPPLIER_BALANCES", "أرصدة الموردين", "SupplierBalanceQuery", "rptSupplierBalances", "", "")
        Case 21: ReportRow = Array("INTEGRITY", "فحص سلامة البيانات", "IntegrityCheckQuery", "rptIntegrityCheck", "", "")
    End Select
End Function
