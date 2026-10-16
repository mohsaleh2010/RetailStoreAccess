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
            s = "SELECT p.ProductID, p.ProductCode AS [ÇáßæÏ], p.Barcode AS [ÑÞã ÇáÈÇÑßæÏ], p.ProductName AS [ÇáãäÊÌ], p.CurrentQuantity AS [ÇáßãíÉ], p.SellingPrice AS "
            s = s & "[ÇáÓÚÑ] FROM Products AS p WHERE p.ProductName Like {LIKE} OR p.ProductCode Like {LIKE} OR p.Barcode Like {LIKE} OR p.ProductNameEn Like {LIKE} OR "
            s = s & "p.ProductID = {NUM} ORDER BY p.ProductName"
        Case "CUSTOMER"
            s = "SELECT c.CustomerID, c.CustomerName AS [ÇáÚãíá], c.Mobile AS [ÑÞã ÇáÌæÇá], c.VATNumber AS [ÇáÑÞã ÇáÖÑíÈí], c.CurrentBalance AS [ÇáÑÕíÏ] FROM "
            s = s & "[@Customers] AS c WHERE c.CustomerName Like {LIKE} OR c.CustomerNameEn Like {LIKE} OR c.Mobile Like {LIKE} OR c.Phone Like {LIKE} OR c.VATNumber Like "
            s = s & "{LIKE} OR c.CustomerID = {NUM} ORDER BY c.CustomerName"
        Case "SUPPLIER"
            s = "SELECT s.SupplierID, s.SupplierName AS [ÇáãæÑÏ], s.ContactPerson AS [ÇáãÓÄæá], s.Mobile AS [ÑÞã ÇáÌæÇá], s.CurrentBalance AS [ÇáÑÕíÏ] FROM "
            s = s & "[@Suppliers] AS s WHERE s.SupplierName Like {LIKE} OR s.SupplierNameEn Like {LIKE} OR s.ContactPerson Like {LIKE} OR s.Mobile Like {LIKE} OR "
            s = s & "s.VATNumber Like {LIKE} OR s.SupplierID = {NUM} ORDER BY s.SupplierName"
        Case "SALE"
            s = "SELECT h.SalesInvoiceID, h.InvoiceNumber AS [ÑÞã ÇáÝÇÊæÑÉ], h.InvoiceDate AS [ÇáÊÇÑíÎ], c.CustomerName AS [ÇáÚãíá], h.TotalAmount AS [ÇáÅÌãÇáí], "
            s = s & "h.RemainingAmount AS [ÇáãÊÈÞí] FROM SalesInvoices AS h INNER JOIN [@Customers] AS c ON h.CustomerID = c.CustomerID WHERE (h.InvoiceNumber Like {LIKE} "
            s = s & "OR c.CustomerName Like {LIKE} OR c.CustomerNameEn Like {LIKE} OR c.Mobile Like {LIKE}) AND h.InvoiceDate >= {FROM} AND h.InvoiceDate < {TO} ORDER BY "
            s = s & "h.InvoiceDate DESC"
        Case "PURCHASE"
            s = "SELECT h.PurchaseInvoiceID, h.InvoiceNumber AS [ÑÞã ÇáÝÇÊæÑÉ], h.SupplierInvoiceNo AS [ÝÇÊæÑÉ ÇáãæÑÏ], h.InvoiceDate AS [ÇáÊÇÑíÎ], s.SupplierName AS "
            s = s & "[ÇáãæÑÏ], h.TotalAmount AS [ÇáÅÌãÇáí] FROM PurchaseInvoices AS h INNER JOIN [@Suppliers] AS s ON h.SupplierID = s.SupplierID WHERE (h.InvoiceNumber "
            s = s & "Like {LIKE} OR h.SupplierInvoiceNo Like {LIKE} OR s.SupplierName Like {LIKE} OR s.SupplierNameEn Like {LIKE}) AND h.InvoiceDate >= {FROM} AND "
            s = s & "h.InvoiceDate < {TO} ORDER BY h.InvoiceDate DESC"
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
    ReportCount = 57
End Function

Public Function ScreenPermission(ByVal FormName As String) As String
    ' Permission needed to open a screen ("" = any logged-in user). Source: forms.SCREEN_PERMISSIONS
    Select Case FormName
        Case "frmPOS": ScreenPermission = "SALES_POS"
        Case "frmSalesInvoice": ScreenPermission = "SALES_VIEW"
        Case "frmSalesReturn": ScreenPermission = "SALES_RETURN"
        Case "frmCustomerPayment": ScreenPermission = "CUSTOMER_PAYMENTS"
        Case "frmCustomers": ScreenPermission = "CUSTOMERS"
        Case "frmSuppliers": ScreenPermission = "SUPPLIERS"
        Case "frmSupplierPayment": ScreenPermission = "SUPPLIER_PAYMENTS"
        Case "frmPurchaseInvoice": ScreenPermission = "PURCHASES"
        Case "frmPurchaseView": ScreenPermission = "PURCHASES"
        Case "frmPurchaseReturn": ScreenPermission = "PURCHASE_RETURN"
        Case "frmProducts": ScreenPermission = "PRODUCTS"
        Case "frmCategories": ScreenPermission = "PRODUCTS"
        Case "frmUnits": ScreenPermission = "PRODUCTS"
        Case "frmInventory": ScreenPermission = "PRODUCTS"
        Case "frmStockCount": ScreenPermission = "STOCK_COUNT"
        Case "frmExpenses": ScreenPermission = "EXPENSES"
        Case "frmExpenseTypes": ScreenPermission = "EXPENSES"
        Case "frmReportCenter": ScreenPermission = "REPORTS"
        Case "frmSettings": ScreenPermission = "SETTINGS"
        Case "frmUsers": ScreenPermission = "USERS"
        Case "frmRoles": ScreenPermission = "USERS"
        Case "frmUserScreens": ScreenPermission = "USERS"
        Case "frmBackup": ScreenPermission = "BACKUP"
        Case "frmBarcodeLabels": ScreenPermission = "PRODUCTS"
        Case "frmLabelSettings": ScreenPermission = "PRODUCTS"
        Case "frmTouchPOS": ScreenPermission = "SALES_POS"
        Case "frmTouchPay": ScreenPermission = "SALES_POS"
        Case "frmCafePOS": ScreenPermission = "SALES_POS"
        Case "frmCafeItem": ScreenPermission = "SALES_POS"
        Case "frmTreasury": ScreenPermission = "CASH_CLOSING"
        Case "frmCashClosing": ScreenPermission = "CASH_CLOSING"
        Case "frmCashVoucher": ScreenPermission = "CASH_BOX"
        Case "frmCashBoxes": ScreenPermission = "CASH_BOX"
        Case "frmCashBoxUsers": ScreenPermission = "CASH_ALL_BOXES"
        Case "frmCashTransfers": ScreenPermission = "CASH_CLOSING"
        Case "frmWarehouses": ScreenPermission = "PRODUCTS"
        Case "frmStockTransfer": ScreenPermission = "STOCK_TRANSFER"
        Case "frmStockParties": ScreenPermission = "STOCK_DOCS"
        Case "frmStockDocument": ScreenPermission = "STOCK_DOCS"
        Case "frmStockNote": ScreenPermission = "STOCK_NOTES"
        Case "frmSalesOrder": ScreenPermission = "SALES_ORDERS"
        Case "frmRepDay": ScreenPermission = "REP_DAYS"
        Case "frmRepVisits": ScreenPermission = "REP_ROUTES"
        Case "frmRouteStops": ScreenPermission = "REP_ROUTES"
        Case "frmRepPrices": ScreenPermission = "REP_ROUTES"
        Case "frmJournal": ScreenPermission = "JOURNAL"
        Case "frmJournalEntry": ScreenPermission = "JOURNAL"
        Case "frmAccounts": ScreenPermission = "JOURNAL"
        Case "frmManualEntry": ScreenPermission = "MANUAL_ENTRY"
        Case "frmLedger": ScreenPermission = "JOURNAL"
        Case "frmFinancials": ScreenPermission = "REPORTS_PROFIT"
        Case "frmPeriodClosing": ScreenPermission = "PERIOD_CLOSE"
        Case "frmFiscalPeriods": ScreenPermission = "PERIOD_CLOSE"
        Case "frmVatReturn": ScreenPermission = "VAT_RETURN"
        Case "frmOpeningBalances": ScreenPermission = "OPENING_BALANCES"
        Case "frmOpeningCustomers": ScreenPermission = "OPENING_BALANCES"
        Case "frmOpeningSuppliers": ScreenPermission = "OPENING_BALANCES"
        Case "frmOpeningBoxes": ScreenPermission = "OPENING_BALANCES"
        Case "frmOpeningBanks": ScreenPermission = "OPENING_BALANCES"
        Case "frmOpeningStock": ScreenPermission = "OPENING_BALANCES"
        Case "frmAging": ScreenPermission = "REPORTS"
        Case "frmAllocation": ScreenPermission = "CUSTOMER_PAYMENTS"
        Case "frmBanks": ScreenPermission = "BANKS"
        Case "frmBankTx": ScreenPermission = "BANKS"
        Case "frmBankRecon": ScreenPermission = "BANKS"
        Case "frmCheques": ScreenPermission = "CHEQUES"
        Case "frmAssets": ScreenPermission = "FIXED_ASSETS"
        Case "frmDepreciation": ScreenPermission = "FIXED_ASSETS"
        Case "frmPayroll": ScreenPermission = "PAYROLL"
        Case "frmEmployeePay": ScreenPermission = "PAYROLL"
        Case "frmCostCenters": ScreenPermission = "JOURNAL"
        Case "frmBudget": ScreenPermission = "BUDGET"
        Case "frmRecurring": ScreenPermission = "EXPENSES"
        Case "frmAuditLog": ScreenPermission = "AUDIT_LOG"
        Case "frmCurrencies": ScreenPermission = "CURRENCIES"
        Case "frmCurrencyRates": ScreenPermission = "CURRENCIES"
        Case "frmSalesReps": ScreenPermission = "SALES_REPS"
        Case "frmRepTargets": ScreenPermission = "SALES_REPS"
        Case "frmCommissions": ScreenPermission = "SALES_REPS"
        Case "frmEnglishNames": ScreenPermission = "SETTINGS"
        Case "frmEInvoices": ScreenPermission = "EINVOICE"
        Case "frmZatcaSetup": ScreenPermission = "SETTINGS"
        Case "frmEtaSetup": ScreenPermission = "SETTINGS"
    End Select
End Function

Public Function ReportRow(ByVal Index As Long) As Variant
    ' Array(Key, Title, QueryName, ReportName, Needs, DateColumn)
    Select Case Index
        Case 1: ReportRow = Array("STATISTICS", "ÇáÅÍÕÇÆíÇÊ æÇáÑÓæã ÇáÈíÇäíÉ", "SalesByCategoryQuery", "rptStatistics", "P", "")
        Case 2: ReportRow = Array("DAILY_SALES", "ÇáãÈíÚÇÊ ÇáíæãíÉ", "DailySalesQuery", "rptDailySales", "D", "SaleDate")
        Case 3: ReportRow = Array("MONTHLY_SALES", "ÇáãÈíÚÇÊ ÇáÔåÑíÉ", "MonthlySalesQuery", "rptMonthlySales", "$", "")
        Case 4: ReportRow = Array("SALES_PERIOD", "ÇáãÈíÚÇÊ ÍÓÈ ÝÊÑÉ", "SalesByPeriodQuery", "rptSalesByPeriod", "Pc", "")
        Case 5: ReportRow = Array("SALES_PRODUCT", "ÇáãÈíÚÇÊ ÍÓÈ ÇáãäÊÌ", "SalesByProductQuery", "rptSalesByProduct", "Pr$", "")
        Case 6: ReportRow = Array("BEST_SELLING", "ÃÝÖá ÇáãäÊÌÇÊ ãÈíÚðÇ", "BestSellingProductsQuery", "rptBestSelling", "P$", "")
        Case 7: ReportRow = Array("LEAST_SELLING", "ÃÞá ÇáãäÊÌÇÊ ãÈíÚðÇ", "LeastSellingProductsQuery", "rptLeastSelling", "P", "")
        Case 8: ReportRow = Array("PURCHASES", "ÇáãÔÊÑíÇÊ", "PurchasesQuery", "rptPurchases", "Ps", "")
        Case 9: ReportRow = Array("STOCK", "ÇáãÎÒæä ÇáÍÇáí", "StockBalanceQuery", "rptStockBalance", "", "")
        Case 10: ReportRow = Array("STOCK_WAREHOUSE", "ÃÑÕÏÉ ÇáÃÕäÇÝ áßá ãÎÒä", "StockByWarehouseQuery", "rptStockByWarehouse", "w", "")
        Case 11: ReportRow = Array("STOCK_VALUE_WAREHOUSE", "ÊÞííã ÇáãÎÒæä áßá ãÎÒä", "StockValueByWarehouseQuery", "rptStockValueByWarehouse", "", "")
        Case 12: ReportRow = Array("TRANSFERS_IN_TRANSIT", "ÇáÊÍæíáÇÊ Èíä ÇáãÎÇÒä ÈÇáØÑíÞ", "TransfersInTransitQuery", "rptTransfersInTransit", "", "")
        Case 13: ReportRow = Array("NOTES_NOT_INVOICED", "ÃÐæäÇÊ ÇáÊÓáíã æÇáÇÓÊáÇã ÇáÊí áã ÊõÝæÊÑ", "NotesNotInvoicedQuery", "rptNotesNotInvoiced", "w", "")
        Case 14: ReportRow = Array("PENDING_ORDERS", "ØáÈíÇÊ ÇáÚãáÇÁ ÇáãÝÊæÍÉ: ÇáÈÇÞí ÈáÇ ÊÓáíã", "PendingOrdersQuery", "rptPendingOrders", "Dcw", "OrderDate")
        Case 15: ReportRow = Array("RESERVED_STOCK", "ÇáÃÕäÇÝ ÇáãÍÌæÒÉ ááØáÈíÇÊ æÇáãÊÇÍ ááÈíÚ", "ReservedStockQuery", "rptReservedStock", "w", "")
        Case 16: ReportRow = Array("REP_DAYS", "ÊÓæíÉ ÃíÇã ÇáãäÏæÈíä: ÇáÈÖÇÚÉ æÇáäÞÏíÉ", "RepDaySummaryQuery", "rptRepDays", "D", "DayDate")
        Case 17: ReportRow = Array("ROUTE_EXECUTION", "ÊäÝíÐ ÎØ ÇáÓíÑ: ÇáÒíÇÑÇÊ ÇáãÎØØÉ æÇáãäÝÐÉ æäÊÇÆÌåÇ", "RouteExecutionQuery", "rptRouteExecution", "D", "VisitDate")
        Case 18: ReportRow = Array("REP_VISITS", "ÒíÇÑÇÊ ÇáãäÏæÈíä ÈÇáÊÝÕíá", "RepVisitListQuery", "rptRepVisits", "Dc", "VisitDate")
        Case 19: ReportRow = Array("REP_VAN_STOCK", "ÑÕíÏ ÓíÇÑÇÊ ÇáãäÏæÈíä", "RepVanStockQuery", "rptRepVanStock", "w", "")
        Case 20: ReportRow = Array("STOCK_COUNT_DIFF", "ÝÑæÞÇÊ ÇáÌÑÏ ÇáãÑÍøáÉ áßá ãÎÒä", "StockCountDiffQuery", "rptStockCountDiff", "Dw", "CountDate")
        Case 21: ReportRow = Array("LOW_STOCK", "ÇáãäÊÌÇÊ ãäÎÝÖÉ ÇáãÎÒæä", "LowStockQuery", "rptLowStock", "", "")
        Case 22: ReportRow = Array("ITEM_CARD", "ßÇÑÊ ÇáÕäÝ (ÇáßãíÉ æÇáÞíãÉ áßá ãÎÒä)", "ItemCardQuery", "rptItemCard", "PRW", "")
        Case 23: ReportRow = Array("PRODUCT_MOVEMENT", "ÍÑßÉ ãäÊÌ", "ProductMovementQuery", "rptProductMovement", "PR", "")
        Case 24: ReportRow = Array("CUSTOMER_STATEMENT", "ßÔÝ ÍÓÇÈ Úãíá", "CustomerStatementQuery", "rptCustomerStatement", "PC", "")
        Case 25: ReportRow = Array("SUPPLIER_STATEMENT", "ßÔÝ ÍÓÇÈ ãæÑÏ", "SupplierStatementQuery", "rptSupplierStatement", "PS", "")
        Case 26: ReportRow = Array("EXPENSES", "ÇáãÕÑæÝÇÊ (ÊÝÕíáí)", "ExpensesQuery", "rptExpenses", "Pe", "")
        Case 27: ReportRow = Array("EXPENSES_BY_TYPE", "ÇáãÕÑæÝÇÊ (ÅÌãÇáí ÍÓÈ ÇáäæÚ)", "ExpensesByTypeQuery", "rptExpensesByType", "P", "")
        Case 28: ReportRow = Array("CASH_STATEMENT", "ÍÑßÉ ÇáÎÒíäÉ / ÇáÕäÏæÞ (ÊÝÕíáí)", "CashStatementQuery", "rptCashStatement", "Pb#", "")
        Case 29: ReportRow = Array("CASH_DAILY", "ÍÑßÉ ÇáÎÒíäÉ ÇáíæãíÉ (Ãæá Çáíæã æÂÎÑå)", "CashDailyQuery", "rptCashDaily", "Pb#", "")
        Case 30: ReportRow = Array("CASH_DAYBOOK", "íæãíÉ ÇáÕäÏæÞ ÇáÊÝÕíáíÉ (ÕÝÍÉ áßá ÕäÏæÞ æíæã)", "CashDayBookQuery", "rptCashDayBook", "Pb#", "")
        Case 31: ReportRow = Array("CASH_BALANCES", "ÃÑÕÏÉ ÇáÎÒíäÉ æÇáÕäÇÏíÞ", "CashBoxBalanceQuery", "rptCashBalances", "#", "")
        Case 32: ReportRow = Array("CASH_CLOSINGS", "ÊÕÝíÇÊ íæãíÉ ÇáßÇÔíÑ", "CashClosingsQuery", "rptCashClosings", "Pb#", "")
        Case 33: ReportRow = Array("PROFIT", "ÇáÃÑÈÇÍ", "ProfitQuery", "rptProfit", "P$", "")
        Case 34: ReportRow = Array("JOURNAL", "ÞíæÏ ÇáíæãíÉ", "JournalLinesQuery", "rptJournal", "PJ", "")
        Case 35: ReportRow = Array("TRIAL_BALANCE", "ãíÒÇä ÇáãÑÇÌÚÉ", "TrialBalanceQuery", "rptTrialBalance", "PJ", "")
        Case 36: ReportRow = Array("TRIAL_BALANCE_TREE", "ãíÒÇä ÇáãÑÇÌÚÉ ÈÇáãÓÊæíÇÊ", "TrialBalanceTreeQuery", "rptTrialBalanceTree", "PJ", "")
        Case 37: ReportRow = Array("ACCOUNT_TREE", "Ïáíá ÇáÍÓÇÈÇÊ (ÔÌÑÉ ÇáÍÓÇÈÇÊ)", "AccountTreeQuery", "rptAccountTree", "J", "")
        Case 38: ReportRow = Array("GENERAL_LEDGER", "ÏÝÊÑ ÇáÃÓÊÇÐ (ßá ÇáÍÓÇÈÇÊ)", "GeneralLedgerQuery", "rptGeneralLedger", "PJ", "")
        Case 39: ReportRow = Array("INCOME_STATEMENT", "ÞÇÆãÉ ÇáÏÎá", "IncomeStatementQuery", "rptIncomeStatement", "PJ$F", "")
        Case 40: ReportRow = Array("BALANCE_SHEET", "ÇáãíÒÇäíÉ ÇáÚãæãíÉ", "BalanceSheetQuery", "rptBalanceSheet", "PJ$F", "")
        Case 41: ReportRow = Array("OPENING_BALANCES", "ÃÑÕÏÉ Ãæá ÇáãÏÉ (Ýí ÈÏÇíÉ ÊÇÑíÎ ÇáÈÏÇíÉ)", "OpeningBalancesReportQuery", "rptOpeningBalances", "PJ", "")
        Case 42: ReportRow = Array("COST_CENTER_PROFIT", "ÞÇÆãÉ ÇáÏÎá ÍÓÈ ãÑßÒ ÇáÊßáÝÉ", "CostCenterProfitQuery", "rptCostCenterProfit", "PJ", "")
        Case 43: ReportRow = Array("COST_CENTER_ACCOUNTS", "ÅíÑÇÏÇÊ æãÕÑæÝÇÊ ßá ãÑßÒ ÊßáÝÉ", "CostCenterAccountsQuery", "rptCostCenterAccounts", "PJ", "")
        Case 44: ReportRow = Array("BUDGET_VS_ACTUAL", "ÇáãæÇÒäÉ ãÞÇÈá ÇáÝÚáí", "BudgetVsActualQuery", "rptBudgetVsActual", "PJ", "")
        Case 45: ReportRow = Array("FIXED_ASSETS", "ÓÌá ÇáÃÕæá ÇáËÇÈÊÉ", "FixedAssetsQuery", "rptFixedAssets", "J", "")
        Case 46: ReportRow = Array("AUDIT_TRAIL", "ÓÌá ÇáÊÏÞíÞ", "AuditTrailQuery", "rptAuditTrail", "DA", "LogDate")
        Case 47: ReportRow = Array("SLOW_MOVING", "ÇáãäÊÌÇÊ ÛíÑ ÇáãÊÍÑßÉ", "SlowMovingProductsQuery", "rptSlowMoving", "", "")
        Case 48: ReportRow = Array("STOCK_BY_CATEGORY", "ÇáãÎÒæä ÍÓÈ ÇáÊÕäíÝ", "StockByCategoryQuery", "rptStockByCategory", "", "")
        Case 49: ReportRow = Array("VAT_SUMMARY", "ãáÎÕ ÖÑíÈÉ ÇáÞíãÉ ÇáãÖÇÝÉ", "VatSummaryQuery", "rptVatSummary", "P$", "")
        Case 50: ReportRow = Array("EINVOICE_DOCS", "ãÓÊäÏÇÊ ÇáÝÇÊæÑÉ ÇáÅáßÊÑæäíÉ (ÇáãÞÈæáÉ æÇáãÑÝæÖÉ)", "EInvoiceReportQuery", "rptEInvoiceDocs", "D", "DocDate")
        Case 51: ReportRow = Array("CUSTOMER_BALANCES", "ÃÑÕÏÉ ÇáÚãáÇÁ", "CustomerBalanceQuery", "rptCustomerBalances", "", "")
        Case 52: ReportRow = Array("SUPPLIER_BALANCES", "ÃÑÕÏÉ ÇáãæÑÏíä", "SupplierBalanceQuery", "rptSupplierBalances", "", "")
        Case 53: ReportRow = Array("SUPPLIER_FX", "ÃÑÕÏÉ ÇáãæÑÏíä ÈÇáÚãáÇÊ æÝÑæÞ ÇáÚãáÉ", "SupplierFxBalanceQuery", "rptSupplierFx", "", "")
        Case 54: ReportRow = Array("REP_PERFORMANCE", "ÃÏÇÁ ÇáãäÏæÈíä: ÇáãÈíÚÇÊ æÇáÊÍÕíá æÇáåÏÝ æÇáÚãæáÉ", "RepPerformanceQuery", "rptRepPerformance", "P", "")
        Case 55: ReportRow = Array("REP_CUSTOMERS", "ÚãáÇÁ ÇáãäÏæÈíä æÃÑÕÏÊåã", "RepCustomersQuery", "rptRepCustomers", "", "")
        Case 56: ReportRow = Array("REP_COMMISSION_BALANCE", "ÚãæáÇÊ ÇáãäÏæÈíä ÇáãÓÊÍÞÉ", "RepCommissionBalanceQuery", "rptRepCommissionBalance", "", "")
        Case 57: ReportRow = Array("INTEGRITY", "ÝÍÕ ÓáÇãÉ ÇáÈíÇäÇÊ", "IntegrityCheckQuery", "rptIntegrityCheck", "", "")
    End Select
End Function
