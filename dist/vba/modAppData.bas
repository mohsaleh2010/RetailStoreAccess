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
            s = "SELECT p.ProductID, p.ProductCode AS [ÇáßæÏ], p.Barcode AS [ÇáÈÇÑßæÏ], p.ProductName AS [ÇáãäÊÌ], p.CurrentQuantity AS [ÇáßãíÉ], p.SellingPrice AS "
            s = s & "[ÇáÓÚÑ] FROM Products AS p WHERE p.ProductName Like {LIKE} OR p.ProductCode Like {LIKE} OR p.Barcode Like {LIKE} OR p.ProductNameEn Like {LIKE} OR "
            s = s & "p.ProductID = {NUM} ORDER BY p.ProductName"
        Case "CUSTOMER"
            s = "SELECT c.CustomerID, c.CustomerName AS [ÇáÚãíá], c.Mobile AS [ÇáÌæÇá], c.VATNumber AS [ÇáÑÞã ÇáÖÑíÈí], c.CurrentBalance AS [ÇáÑÕíÏ] FROM Customers AS "
            s = s & "c WHERE c.CustomerName Like {LIKE} OR c.Mobile Like {LIKE} OR c.Phone Like {LIKE} OR c.VATNumber Like {LIKE} OR c.CustomerID = {NUM} ORDER BY "
            s = s & "c.CustomerName"
        Case "SUPPLIER"
            s = "SELECT s.SupplierID, s.SupplierName AS [ÇáãæÑÏ], s.ContactPerson AS [ÇáãÓÄæá], s.Mobile AS [ÇáÌæÇá], s.CurrentBalance AS [ÇáÑÕíÏ] FROM Suppliers AS s "
            s = s & "WHERE s.SupplierName Like {LIKE} OR s.ContactPerson Like {LIKE} OR s.Mobile Like {LIKE} OR s.VATNumber Like {LIKE} OR s.SupplierID = {NUM} ORDER BY "
            s = s & "s.SupplierName"
        Case "SALE"
            s = "SELECT h.SalesInvoiceID, h.InvoiceNumber AS [ÑÞã ÇáÝÇÊæÑÉ], h.InvoiceDate AS [ÇáÊÇÑíÎ], c.CustomerName AS [ÇáÚãíá], h.TotalAmount AS [ÇáÅÌãÇáí], "
            s = s & "h.RemainingAmount AS [ÇáãÊÈÞí] FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID WHERE (h.InvoiceNumber Like {LIKE} OR "
            s = s & "c.CustomerName Like {LIKE} OR c.Mobile Like {LIKE}) AND h.InvoiceDate >= {FROM} AND h.InvoiceDate < {TO} ORDER BY h.InvoiceDate DESC"
        Case "PURCHASE"
            s = "SELECT h.PurchaseInvoiceID, h.InvoiceNumber AS [ÑÞã ÇáÝÇÊæÑÉ], h.SupplierInvoiceNo AS [ÝÇÊæÑÉ ÇáãæÑÏ], h.InvoiceDate AS [ÇáÊÇÑíÎ], s.SupplierName AS "
            s = s & "[ÇáãæÑÏ], h.TotalAmount AS [ÇáÅÌãÇáí] FROM PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID WHERE (h.InvoiceNumber Like "
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
        Case 1: ReportRow = Array("DAILY_SALES", "ÇáãÈíÚÇÊ ÇáíæãíÉ", "DailySalesQuery", "rptDailySales", "D", "SaleDate")
        Case 2: ReportRow = Array("MONTHLY_SALES", "ÇáãÈíÚÇÊ ÇáÔåÑíÉ", "MonthlySalesQuery", "rptMonthlySales", "", "")
        Case 3: ReportRow = Array("SALES_PERIOD", "ÇáãÈíÚÇÊ ÍÓÈ ÝÊÑÉ", "SalesByPeriodQuery", "rptSalesByPeriod", "Pc", "")
        Case 4: ReportRow = Array("SALES_PRODUCT", "ÇáãÈíÚÇÊ ÍÓÈ ÇáãäÊÌ", "SalesByProductQuery", "rptSalesByProduct", "Pr", "")
        Case 5: ReportRow = Array("BEST_SELLING", "ÃÝÖá ÇáãäÊÌÇÊ ãÈíÚðÇ", "BestSellingProductsQuery", "rptBestSelling", "P", "")
        Case 6: ReportRow = Array("LEAST_SELLING", "ÃÞá ÇáãäÊÌÇÊ ãÈíÚðÇ", "LeastSellingProductsQuery", "rptLeastSelling", "P", "")
        Case 7: ReportRow = Array("PURCHASES", "ÇáãÔÊÑíÇÊ", "PurchasesQuery", "rptPurchases", "Ps", "")
        Case 8: ReportRow = Array("STOCK", "ÇáãÎÒæä ÇáÍÇáí", "StockBalanceQuery", "rptStockBalance", "", "")
        Case 9: ReportRow = Array("LOW_STOCK", "ÇáãäÊÌÇÊ ãäÎÝÖÉ ÇáãÎÒæä", "LowStockQuery", "rptLowStock", "", "")
        Case 10: ReportRow = Array("PRODUCT_MOVEMENT", "ÍÑßÉ ãäÊÌ", "ProductMovementQuery", "rptProductMovement", "PR", "")
        Case 11: ReportRow = Array("CUSTOMER_STATEMENT", "ßÔÝ ÍÓÇÈ Úãíá", "CustomerStatementQuery", "rptCustomerStatement", "PC", "")
        Case 12: ReportRow = Array("SUPPLIER_STATEMENT", "ßÔÝ ÍÓÇÈ ãæÑÏ", "SupplierStatementQuery", "rptSupplierStatement", "PS", "")
        Case 13: ReportRow = Array("EXPENSES", "ÇáãÕÑæÝÇÊ", "ExpensesQuery", "rptExpenses", "P", "")
        Case 14: ReportRow = Array("EXPENSES_BY_TYPE", "ÇáãÕÑæÝÇÊ ÍÓÈ ÇáäæÚ", "ExpensesByTypeQuery", "rptExpensesByType", "P", "")
        Case 15: ReportRow = Array("PROFIT", "ÇáÃÑÈÇÍ", "ProfitQuery", "rptProfit", "P", "")
        Case 16: ReportRow = Array("SLOW_MOVING", "ÇáãäÊÌÇÊ ÛíÑ ÇáãÊÍÑßÉ", "SlowMovingProductsQuery", "rptSlowMoving", "", "")
        Case 17: ReportRow = Array("STOCK_BY_CATEGORY", "ÇáãÎÒæä ÍÓÈ ÇáÊÕäíÝ", "StockByCategoryQuery", "rptStockByCategory", "", "")
        Case 18: ReportRow = Array("VAT_SUMMARY", "ãáÎÕ ÖÑíÈÉ ÇáÞíãÉ ÇáãÖÇÝÉ", "VatSummaryQuery", "rptVatSummary", "P", "")
        Case 19: ReportRow = Array("CUSTOMER_BALANCES", "ÃÑÕÏÉ ÇáÚãáÇÁ", "CustomerBalanceQuery", "rptCustomerBalances", "", "")
        Case 20: ReportRow = Array("SUPPLIER_BALANCES", "ÃÑÕÏÉ ÇáãæÑÏíä", "SupplierBalanceQuery", "rptSupplierBalances", "", "")
        Case 21: ReportRow = Array("INTEGRITY", "ÝÍÕ ÓáÇãÉ ÇáÈíÇäÇÊ", "IntegrityCheckQuery", "rptIntegrityCheck", "", "")
    End Select
End Function
