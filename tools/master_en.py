"""English names of the system rows that BuildSchema seeds (docs/39-English-Master-Data.md).

They fill the optional English name fields of the tables in ENGLISH_NAMES (Accounts.AccountNameEn,
PaymentMethods.MethodNameEn, ..., Units.UnitNameEn). An English front-end shows
Nz(English name, Arabic name), so a row the user added without an English name keeps its Arabic one.
BuildSchema fills only the English names that are still empty (a name the user changed is kept)."""

ACCOUNT_NAMES_EN = {
    1: "Assets",
    11: "Current assets",
    1100: "Cash in treasury and boxes",
    110001: "Main treasury",
    110002: "Cashier box",
    1190: "Cash not allocated to a box",
    1200: "Mada and wallets under settlement (bank and network)",
    1210: "Bank accounts",
    1250: "Cheques under collection",
    1300: "Accounts receivable",
    1310: "Notes receivable",
    1350: "Allowance for doubtful debts",
    1400: "Inventory",
    1500: "VAT - input",
    1600: "Employee advances",
    1610: "Cash floats",
    1650: "Prepaid expenses",
    1660: "Accrued revenue",
    1690: "Deposits with others",
    12: "Non-current assets",
    1710: "Furniture and fixtures",
    1720: "Computers, devices and POS systems",
    1730: "Vehicles",
    1740: "Decorations and store improvements",
    1790: "Accumulated depreciation of fixed assets",
    1800: "Software and licenses",
    2: "Liabilities",
    21: "Current liabilities",
    2100: "Accounts payable",
    2110: "Notes payable (issued cheques)",
    2200: "VAT - output",
    2250: "VAT - settlement and payment",
    2300: "Accrued expenses",
    2310: "Salaries and wages payable",
    2320: "Social insurance payable",
    2330: "Sales rep commissions payable",
    2400: "Customer advances",
    2500: "Short-term loans",
    2600: "Zakat payable",
    22: "Non-current liabilities",
    2700: "Long-term loans",
    2800: "End-of-service benefits provision",
    3: "Equity",
    31: "Capital and owner current account",
    3200: "Capital",
    3100: "Owner current account",
    3900: "Opening balances",
    32: "Retained earnings and results",
    3300: "Retained earnings",
    3400: "Net profit (loss) of the year",
    4: "Revenue",
    41: "Operating revenue",
    4100: "Sales",
    4110: "Sales returns",
    4120: "Discounts allowed",
    42: "Other revenue",
    4200: "Miscellaneous revenue",
    4300: "Cash overages",
    4400: "Discounts received",
    4500: "Gain on sale of fixed assets",
    5: "Expenses",
    51: "Cost of sales",
    5100: "Cost of goods sold",
    5200: "Stock differences and adjustments",
    52: "Operating and administrative expenses",
    5300: "Expenses by type",
    5500: "Salaries and wages",
    5510: "Allowances and incentives",
    5520: "Social insurance",
    5530: "Sales rep commissions",
    5600: "Depreciation of fixed assets",
    5610: "Bank and POS fees",
    5620: "Government fees and licenses",
    5630: "Advertising",
    53: "Other expenses",
    5400: "Cash shortages",
    5650: "Loss on sale and disposal of fixed assets",
    5700: "Bad debts",
    5800: "Zakat",
    5900: "Miscellaneous expenses",
}

PAYMENT_METHODS_EN = {1: "Cash", 2: "Mada / card", 3: "Bank transfer", 4: "E-wallet"}

SOURCE_TYPES_EN = {
    "SALE": "Sales invoice", "SALES_RETURN": "Sales return", "PURCHASE": "Purchase invoice",
    "PURCHASE_RETURN": "Purchase return", "CUSTOMER_PAYMENT": "Customer receipt",
    "SUPPLIER_PAYMENT": "Supplier payment", "EXPENSE": "Expense", "CASH_VOUCHER": "Cash voucher",
    "STOCK_MOVE": "Manual stock move", "STOCK_COUNT": "Stock count adjustment",
    "BOX_OPENING": "Box opening balance", "CUSTOMER_OPENING": "Customer opening balance",
    "SUPPLIER_OPENING": "Supplier opening balance", "MANUAL": "Manual entry", "YEAR_CLOSE": "Year closing entry",
    "VAT_RETURN": "VAT return settlement", "VAT_PAYMENT": "VAT payment", "BANK_OPENING": "Bank opening balance",
    "BANK_TX": "Bank transaction", "CHEQUE": "Received or issued cheque",
    "CHEQUE_STATUS": "Cheque collection or bounce", "ASSET": "Fixed asset purchase",
    "ASSET_DISPOSAL": "Fixed asset sale or disposal", "DEPRECIATION": "Monthly depreciation entry",
    "PAYROLL": "Payroll entry", "PAYROLL_PAYMENT": "Salary payment", "COMMISSION": "Sales rep commissions entry",
}

TRANSACTION_TYPES_EN = {
    "PURCHASE": "Purchase", "SALE": "Sale", "PURCHASE_RETURN": "Purchase return", "SALES_RETURN": "Sales return",
    "STOCK_IN": "Stock addition", "STOCK_OUT": "Stock deduction", "ADJUSTMENT": "Count adjustment",
    "OPENING": "Opening balance",
}

ROLE_NAMES_EN = {1: "System administrator", 2: "Manager", 3: "Cashier"}

PERMISSION_NAMES_EN = {
    "SALES_POS": "Point of sale", "SALES_VIEW": "View and reprint invoices", "SALES_RETURN": "Sales returns",
    "PRICE_OVERRIDE": "Change the sale price in the invoice", "DISCOUNT_OVERRIDE": "Discount above the allowed limit",
    "ALLOW_NEGATIVE_STOCK": "Sell more than the available quantity", "CUSTOMERS": "Manage customers",
    "CUSTOMER_PAYMENTS": "Receipt vouchers", "PURCHASES": "Purchase invoices", "PURCHASE_RETURN": "Purchase returns",
    "SUPPLIERS": "Manage suppliers", "SUPPLIER_PAYMENTS": "Payment vouchers", "PRODUCTS": "Manage products and prices",
    "INVENTORY_ADJUST": "Manual stock addition and deduction", "STOCK_COUNT": "Stock count", "EXPENSES": "Expenses",
    "CASH_BOX": "Treasury: receipt, payment and transfer vouchers and boxes", "CASH_CLOSING": "Daily cashier closing",
    "JOURNAL": "Journal entries, chart of accounts and trial balance",
    "MANUAL_ENTRY": "Manual entries: add, edit and delete",
    "PERIOD_CLOSE": "Close and reopen periods and the fiscal year",
    "VAT_RETURN": "VAT return: approval and payment",
    "BANKS": "Banks: accounts, bank transactions and reconciliation",
    "CHEQUES": "Received and issued cheques: recording, collection and bounce",
    "FIXED_ASSETS": "Fixed assets and depreciation", "PAYROLL": "Payroll: preparation, posting and payment",
    "BUDGET": "Budget: preparation and comparison with actuals", "CURRENCIES": "Currencies and exchange rates",
    "SALES_REPS": "Sales reps: data, targets, commissions and reports", "REPORTS": "Operating reports",
    "REPORTS_PROFIT": "Profit and VAT reports", "DASHBOARD_FINANCIAL": "Financial figures on the dashboard",
    "SETTINGS": "Store settings", "USERS": "Users and permissions", "BACKUP": "Backup",
    "AUDIT_LOG": "Audit trail: who added, edited or deleted, with the values before and after",
    "EINVOICE": "E-invoicing: follow-up and resending",
}

SCREEN_TITLES_EN = {
    "frmPOS": "Point of sale (shops)", "frmTouchPOS": "Point of sale (restaurants)",
    "frmCafePOS": "Point of sale (cafes)", "frmSalesInvoice": "View and reprint invoices",
    "frmSalesReturn": "Sales returns", "frmCustomers": "Customers", "frmCustomerPayment": "Customer receipt vouchers",
    "frmPurchaseInvoice": "Purchase invoices", "frmPurchaseView": "View purchase invoices",
    "frmPurchaseReturn": "Purchase returns", "frmSuppliers": "Suppliers", "frmSupplierPayment": "Supplier payment vouchers",
    "frmProducts": "Products and prices", "frmCategories": "Categories", "frmUnits": "Units",
    "frmInventory": "Stock and manual moves", "frmStockCount": "Stock count", "frmBarcodeLabels": "Barcode labels",
    "frmLabelSettings": "Label settings", "frmExpenses": "Expenses", "frmExpenseTypes": "Expense types",
    "frmRecurring": "Recurring expenses", "frmTreasury": "Treasury", "frmCashVoucher": "Cash and transfer vouchers",
    "frmCashClosing": "Daily cashier closing", "frmCashBoxes": "Boxes", "frmAccounting": "Accounting and finance",
    "frmJournal": "Journal entries", "frmAccounts": "Chart of accounts", "frmManualEntry": "Manual entries",
    "frmLedger": "Account statement and general ledger", "frmFinancials": "Financial statements",
    "frmPeriodClosing": "Period and fiscal year closing", "frmVatReturn": "VAT return",
    "frmAging": "Aging (customers and suppliers)", "frmBanks": "Banks", "frmBankTx": "Bank transactions",
    "frmBankRecon": "Bank reconciliation", "frmCheques": "Received and issued cheques", "frmAssets": "Fixed assets",
    "frmDepreciation": "Monthly depreciation", "frmPayroll": "Payroll", "frmCostCenters": "Cost centers and branches",
    "frmBudget": "Budget", "frmCurrencies": "Currencies", "frmSalesReps": "Sales reps",
    "frmRepTargets": "Sales rep targets", "frmCommissions": "Sales rep commissions", "frmCurrencyRates": "Exchange rates",
    "frmAllocation": "Match payments to invoices", "frmReportCenter": "Reports", "frmSearch": "Search",
    "frmSettings": "Store settings", "frmUsers": "Users", "frmRoles": "Roles and permissions",
    "frmUserScreens": "Screen permissions of users", "frmAuditLog": "Audit trail", "frmBackup": "Backup",
    "frmEnglishNames": "English names", "frmEInvoices": "E-invoicing",
}

CATEGORY_NAMES_EN = {1: "General"}

UNIT_NAMES_EN = {1: "Piece", 2: "Box", 3: "Carton", 4: "Pack", 5: "Kilo", 6: "Litre", 7: "Metre", 8: "Set"}

EXPENSE_TYPE_NAMES_EN = {1: "Rent", 2: "Electricity", 3: "Water", 4: "Internet and telecom", 5: "Transport",
                         6: "Maintenance", 7: "Salaries", 8: "Supplies", 9: "Other expenses"}

CASH_BOX_NAMES_EN = {1: "Main treasury", 2: "Cashier box"}

CUSTOMER_NAMES_EN = {1: "Cash customer"}

MODULE_NAMES_EN = {"المبيعات": "Sales", "العملاء": "Customers", "المشتريات": "Purchases", "الموردون": "Suppliers",
                   "المخزون": "Inventory", "المصروفات": "Expenses", "الخزينة": "Treasury", "الحسابات": "Accounting",
                   "التقارير": "Reports", "النظام": "System"}

# table: (English field, Arabic field, key field, names by key)
ENGLISH_NAMES = {
    "Accounts": ("AccountNameEn", "AccountName", "AccountCode", ACCOUNT_NAMES_EN),
    "PaymentMethods": ("MethodNameEn", "MethodName", "PaymentMethodID", PAYMENT_METHODS_EN),
    "JournalSourceTypes": ("TypeNameEn", "TypeName", "SourceType", SOURCE_TYPES_EN),
    "TransactionTypes": ("TypeNameEn", "TypeName", "TypeCode", TRANSACTION_TYPES_EN),
    "Roles": ("RoleNameEn", "RoleName", "RoleID", ROLE_NAMES_EN),
    "Permissions": ("PermissionNameEn", "PermissionName", "PermissionKey", PERMISSION_NAMES_EN),
    "Screens": ("ScreenTitleEn", "ScreenTitle", "ScreenName", SCREEN_TITLES_EN),
    "Categories": ("CategoryNameEn", "CategoryName", "CategoryID", CATEGORY_NAMES_EN),
    "Units": ("UnitNameEn", "UnitName", "UnitID", UNIT_NAMES_EN),
    "ExpenseTypes": ("ExpenseTypeNameEn", "ExpenseTypeName", "ExpenseTypeID", EXPENSE_TYPE_NAMES_EN),
    # the names the user types (docs/40): no seeded rows but the two boxes and the cash customer
    "Customers": ("CustomerNameEn", "CustomerName", "CustomerID", CUSTOMER_NAMES_EN),
    "Suppliers": ("SupplierNameEn", "SupplierName", "SupplierID", {}),
    "CashBoxes": ("BoxNameEn", "BoxName", "CashBoxID", CASH_BOX_NAMES_EN),
    "Banks": ("BankNameEn", "BankName", "BankID", {}),
    "CostCenters": ("CenterNameEn", "CenterName", "CostCenterID", {}),
    "SalesReps": ("RepNameEn", "RepName", "SalesRepID", {}),
}

# More English names of a table of ENGLISH_NAMES, by the Arabic value (the group of a permission or screen in the
# permission screens, docs/41): table: [(English field, Arabic field, names by Arabic value)]
EXTRA_NAMES = {
    "Permissions": [("ModuleNameEn", "ModuleName", MODULE_NAMES_EN)],
    "Screens": [("ModuleNameEn", "ModuleName", MODULE_NAMES_EN)],
}

# The tables of the English names screen (frmEnglishNames, docs/42), the names the user types first.
# modEnglishNames.NAME_TABLES holds the same list (tests/test_english_names.py).
NAMES_SCREEN_TABLES = ["Customers", "Suppliers", "Products", "Categories", "Units", "CashBoxes", "Banks", "CostCenters",
                       "SalesReps", "ExpenseTypes", "Accounts", "PaymentMethods", "Roles"]


def names_screen_specs():
    """(table, key field, Arabic field, English field, English field size, table caption). Every key is a number."""
    from schema import table
    out = []
    for name in NAMES_SCREEN_TABLES:
        if name == "Products":
            en, ar, key = "ProductNameEn", "ProductName", "ProductID"
        else:
            en, ar, key, _ = ENGLISH_NAMES[name]
        t = table(name)
        fields = {f.name: f for f in t.fields}
        out.append((name, key, ar, en, fields[en].size, t.caption))
    return out


def names_screen_const() -> str:
    return ";".join(f"{t},{k},{a},{e},{size},{cap}" for t, k, a, e, size, cap in names_screen_specs())
