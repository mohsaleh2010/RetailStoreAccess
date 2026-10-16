# مرجع العلاقات (Relationships Reference)

> ملف مُولَّد تلقائيًا من `tools/schema.py` و`tools/relations.py` – لا تعدّله يدويًا.

عدد العلاقات: **210** – جميعها مع **Enforce Referential Integrity**. الحذف المتتالي: **25**، التحديث المتتالي: **12**.

حقول «سجّلها الموظف» التالية تشير إلى `Employees` بلا علاقة مفروضة: Access يسمح بـ32 فهرسًا لكل جدول، وكل علاقة تُحسب فهرسًا على طرفيها. البرنامج يكتب فيها المستخدم الحالي، والمستخدمون يُعطَّلون ولا يُحذفون: `AssetRevaluations.EmployeeID`، `AssetTransfers.EmployeeID`، `AuditLog.EmployeeID`، `BankReconciliations.EmployeeID`، `Budgets.EmployeeID`، `CashVouchers.ReceivedByID`، `CommissionRuns.EmployeeID`، `CustomerAllocations.EmployeeID`، `DepreciationRuns.EmployeeID`، `FiscalYearClosings.EmployeeID`، `PeriodClosings.EmployeeID`، `RepDays.AdvanceEmployeeID`، `RepDays.EmployeeID`، `RepVisits.EmployeeID`، `SalesOrders.EmployeeID`، `StockCounts.PostedByID`، `StockDocuments.EmployeeID`، `StockNotes.EmployeeID`، `StockTransfers.EmployeeID`، `StockTransfers.ReceivedByID`، `SupplierAllocations.EmployeeID`، `VatReturns.EmployeeID`، `Warehouses.KeeperID`.

## مخطط الكيانات والعلاقات

```mermaid
erDiagram
    Customers ||--o{ Settings : "DefaultCustomerID"
    Banks |o--o{ Settings : "DefaultBankID"
    Roles ||--o{ RolePermissions : "RoleID"
    Permissions ||--o{ RolePermissions : "PermissionKey"
    Roles ||--o{ Employees : "RoleID"
    CashBoxes |o--o{ Employees : "CashBoxID"
    CostCenters |o--o{ Employees : "CostCenterID"
    Warehouses |o--o{ Employees : "WarehouseID"
    Employees ||--o{ UserScreens : "EmployeeID"
    Screens ||--o{ UserScreens : "ScreenName"
    Employees |o--o{ Activations : "EmployeeID"
    Currencies ||--o{ CurrencyRates : "CurrencyCode"
    SalesReps |o--o{ CashBoxes : "SalesRepID"
    CashBoxes ||--o{ CashBoxUsers : "CashBoxID"
    Employees ||--o{ CashBoxUsers : "EmployeeID"
    CostCenters |o--o{ Warehouses : "CostCenterID"
    SalesReps |o--o{ Warehouses : "SalesRepID"
    Currencies |o--o{ Suppliers : "CurrencyCode"
    SalesReps |o--o{ Customers : "SalesRepID"
    Categories ||--o{ Products : "CategoryID"
    Units ||--o{ Products : "UnitID"
    Suppliers |o--o{ Products : "SupplierID"
    Customers ||--o{ SalesInvoices : "CustomerID"
    Employees ||--o{ SalesInvoices : "EmployeeID"
    PaymentMethods |o--o{ SalesInvoices : "PaymentMethodID"
    CashBoxes |o--o{ SalesInvoices : "CashBoxID"
    Banks |o--o{ SalesInvoices : "BankID"
    CostCenters |o--o{ SalesInvoices : "CostCenterID"
    SalesReps |o--o{ SalesInvoices : "SalesRepID"
    Warehouses ||--o{ SalesInvoices : "WarehouseID"
    SalesInvoices ||--o{ SalesInvoiceDetails : "SalesInvoiceID"
    Products ||--o{ SalesInvoiceDetails : "ProductID"
    SalesInvoices ||--o{ SalesReturns : "SalesInvoiceID"
    Customers ||--o{ SalesReturns : "CustomerID"
    Employees ||--o{ SalesReturns : "EmployeeID"
    PaymentMethods |o--o{ SalesReturns : "PaymentMethodID"
    CashBoxes |o--o{ SalesReturns : "CashBoxID"
    Banks |o--o{ SalesReturns : "BankID"
    CostCenters |o--o{ SalesReturns : "CostCenterID"
    SalesReps |o--o{ SalesReturns : "SalesRepID"
    Warehouses ||--o{ SalesReturns : "WarehouseID"
    SalesReturns ||--o{ SalesReturnDetails : "SalesReturnID"
    SalesInvoiceDetails ||--o{ SalesReturnDetails : "SalesDetailID"
    Products ||--o{ SalesReturnDetails : "ProductID"
    Suppliers ||--o{ PurchaseInvoices : "SupplierID"
    Employees ||--o{ PurchaseInvoices : "EmployeeID"
    PaymentMethods |o--o{ PurchaseInvoices : "PaymentMethodID"
    CashBoxes |o--o{ PurchaseInvoices : "CashBoxID"
    Banks |o--o{ PurchaseInvoices : "BankID"
    Currencies |o--o{ PurchaseInvoices : "CurrencyCode"
    Warehouses ||--o{ PurchaseInvoices : "WarehouseID"
    PurchaseInvoices ||--o{ PurchaseInvoiceDetails : "PurchaseInvoiceID"
    Products ||--o{ PurchaseInvoiceDetails : "ProductID"
    PurchaseInvoices ||--o{ PurchaseReturns : "PurchaseInvoiceID"
    Suppliers ||--o{ PurchaseReturns : "SupplierID"
    Employees ||--o{ PurchaseReturns : "EmployeeID"
    PaymentMethods |o--o{ PurchaseReturns : "PaymentMethodID"
    CashBoxes |o--o{ PurchaseReturns : "CashBoxID"
    Banks |o--o{ PurchaseReturns : "BankID"
    Currencies |o--o{ PurchaseReturns : "CurrencyCode"
    Warehouses ||--o{ PurchaseReturns : "WarehouseID"
    PurchaseReturns ||--o{ PurchaseReturnDetails : "PurchaseReturnID"
    PurchaseInvoiceDetails ||--o{ PurchaseReturnDetails : "PurchaseDetailID"
    Products ||--o{ PurchaseReturnDetails : "ProductID"
    Customers ||--o{ CustomerPayments : "CustomerID"
    PaymentMethods ||--o{ CustomerPayments : "PaymentMethodID"
    SalesInvoices |o--o{ CustomerPayments : "SalesInvoiceID"
    Employees ||--o{ CustomerPayments : "EmployeeID"
    CashBoxes |o--o{ CustomerPayments : "CashBoxID"
    Banks |o--o{ CustomerPayments : "BankID"
    Currencies |o--o{ CustomerPayments : "CurrencyCode"
    SalesReps |o--o{ CustomerPayments : "SalesRepID"
    Suppliers ||--o{ SupplierPayments : "SupplierID"
    PaymentMethods ||--o{ SupplierPayments : "PaymentMethodID"
    PurchaseInvoices |o--o{ SupplierPayments : "PurchaseInvoiceID"
    Employees ||--o{ SupplierPayments : "EmployeeID"
    CashBoxes |o--o{ SupplierPayments : "CashBoxID"
    Banks |o--o{ SupplierPayments : "BankID"
    Currencies |o--o{ SupplierPayments : "CurrencyCode"
    Banks ||--o{ BankTransactions : "BankID"
    Banks |o--o{ BankTransactions : "ToBankID"
    CashBoxes |o--o{ BankTransactions : "CashBoxID"
    Accounts |o--o{ BankTransactions : "CounterAccount"
    Employees ||--o{ BankTransactions : "EmployeeID"
    Customers |o--o{ Cheques : "CustomerID"
    Suppliers |o--o{ Cheques : "SupplierID"
    Banks |o--o{ Cheques : "BankID"
    Employees ||--o{ Cheques : "EmployeeID"
    Accounts ||--o{ FixedAssets : "AssetAccount"
    CostCenters |o--o{ FixedAssets : "CostCenterID"
    Banks |o--o{ FixedAssets : "BankID"
    CashBoxes |o--o{ FixedAssets : "CashBoxID"
    Accounts |o--o{ FixedAssets : "CounterAccount"
    Banks |o--o{ FixedAssets : "DisposalBankID"
    CashBoxes |o--o{ FixedAssets : "DisposalCashBoxID"
    Employees ||--o{ FixedAssets : "EmployeeID"
    DepreciationRuns ||--o{ AssetDepreciations : "RunID"
    FixedAssets ||--o{ AssetDepreciations : "AssetID"
    FixedAssets ||--o{ AssetRevaluations : "AssetID"
    FixedAssets ||--o{ AssetTransfers : "AssetID"
    CostCenters |o--o{ AssetTransfers : "FromCenterID"
    CostCenters |o--o{ AssetTransfers : "ToCenterID"
    Employees |o--o{ SalesReps : "EmployeeID"
    CostCenters |o--o{ SalesReps : "CostCenterID"
    SalesReps ||--o{ SalesRepTargets : "SalesRepID"
    CommissionRuns ||--o{ CommissionLines : "CommissionRunID"
    SalesReps ||--o{ CommissionLines : "SalesRepID"
    CostCenters |o--o{ CommissionLines : "CostCenterID"
    Budgets ||--o{ BudgetLines : "BudgetID"
    Accounts ||--o{ BudgetLines : "AccountCode"
    CostCenters |o--o{ BudgetLines : "CostCenterID"
    Banks |o--o{ PayrollRuns : "BankID"
    CashBoxes |o--o{ PayrollRuns : "CashBoxID"
    Employees ||--o{ PayrollRuns : "EmployeeID"
    PayrollRuns ||--o{ PayrollLines : "PayrollRunID"
    Employees ||--o{ PayrollLines : "EmployeeID"
    CostCenters |o--o{ PayrollLines : "CostCenterID"
    Banks ||--o{ BankReconciliations : "BankID"
    BankReconciliations ||--o{ BankClearings : "ReconciliationID"
    Banks ||--o{ BankClearings : "BankID"
    CustomerPayments ||--o{ CustomerAllocations : "PaymentID"
    SalesInvoices ||--o{ CustomerAllocations : "SalesInvoiceID"
    SupplierPayments ||--o{ SupplierAllocations : "PaymentID"
    PurchaseInvoices ||--o{ SupplierAllocations : "PurchaseInvoiceID"
    ExpenseTypes ||--o{ Expenses : "ExpenseTypeID"
    PaymentMethods |o--o{ Expenses : "PaymentMethodID"
    Employees ||--o{ Expenses : "EmployeeID"
    CashBoxes |o--o{ Expenses : "CashBoxID"
    Banks |o--o{ Expenses : "BankID"
    CostCenters |o--o{ Expenses : "CostCenterID"
    RecurringExpenses |o--o{ Expenses : "RecurringID"
    Currencies |o--o{ Expenses : "CurrencyCode"
    ExpenseTypes ||--o{ RecurringExpenses : "ExpenseTypeID"
    PaymentMethods ||--o{ RecurringExpenses : "PaymentMethodID"
    CashBoxes |o--o{ RecurringExpenses : "CashBoxID"
    Banks |o--o{ RecurringExpenses : "BankID"
    CostCenters |o--o{ RecurringExpenses : "CostCenterID"
    CashBoxes ||--o{ CashVouchers : "CashBoxID"
    CashBoxes |o--o{ CashVouchers : "ToCashBoxID"
    Expenses |o--o{ CashVouchers : "ExpenseID"
    CashClosings |o--o{ CashVouchers : "ClosingID"
    Employees |o--o{ CashVouchers : "AdvanceEmployeeID"
    CostCenters |o--o{ CashVouchers : "CostCenterID"
    Employees ||--o{ CashVouchers : "EmployeeID"
    SalesReps |o--o{ CashVouchers : "SalesRepID"
    CashBoxes ||--o{ CashClosings : "CashBoxID"
    Employees ||--o{ CashClosings : "EmployeeID"
    CashBoxes |o--o{ CashClosings : "ToCashBoxID"
    JournalSourceTypes ||--o{ JournalEntries : "SourceType"
    Currencies |o--o{ JournalEntries : "CurrencyCode"
    JournalEntries ||--o{ JournalLines : "EntryID"
    Accounts ||--o{ JournalLines : "AccountCode"
    CostCenters |o--o{ JournalLines : "CostCenterID"
    FiscalYearClosings ||--o{ FiscalYearClosingLines : "YearClosingID"
    Accounts ||--o{ FiscalYearClosingLines : "AccountCode"
    Accounts |o--o{ VatReturns : "PaidAccount"
    Employees ||--o{ ManualEntries : "EmployeeID"
    Currencies |o--o{ ManualEntries : "CurrencyCode"
    ManualEntries ||--o{ ManualEntryLines : "ManualEntryID"
    Accounts ||--o{ ManualEntryLines : "AccountCode"
    CostCenters |o--o{ ManualEntryLines : "CostCenterID"
    Accounts ||--o{ OpeningBalances : "AccountCode"
    Products ||--o{ OpeningStock : "ProductID"
    Warehouses ||--o{ OpeningStock : "WarehouseID"
    Products ||--o{ InventoryTransactions : "ProductID"
    TransactionTypes ||--o{ InventoryTransactions : "TransactionTypeID"
    Employees |o--o{ InventoryTransactions : "EmployeeID"
    Warehouses ||--o{ InventoryTransactions : "WarehouseID"
    Warehouses ||--o{ StockTransfers : "FromWarehouseID"
    Warehouses ||--o{ StockTransfers : "ToWarehouseID"
    StockTransfers ||--o{ StockTransferLines : "StockTransferID"
    Products ||--o{ StockTransferLines : "ProductID"
    Accounts |o--o{ StockParties : "AccountCode"
    CostCenters |o--o{ StockParties : "CostCenterID"
    Warehouses ||--o{ StockDocuments : "WarehouseID"
    StockParties ||--o{ StockDocuments : "StockPartyID"
    StockDocuments ||--o{ StockDocumentLines : "StockDocumentID"
    Products ||--o{ StockDocumentLines : "ProductID"
    Warehouses ||--o{ StockNotes : "WarehouseID"
    Customers |o--o{ StockNotes : "CustomerID"
    Suppliers |o--o{ StockNotes : "SupplierID"
    SalesOrders |o--o{ StockNotes : "SalesOrderID"
    SalesInvoices |o--o{ StockNotes : "SalesInvoiceID"
    PurchaseInvoices |o--o{ StockNotes : "PurchaseInvoiceID"
    StockNotes ||--o{ StockNoteLines : "StockNoteID"
    Products ||--o{ StockNoteLines : "ProductID"
    SalesOrderLines |o--o{ StockNoteLines : "SalesOrderLineID"
    Customers ||--o{ SalesOrders : "CustomerID"
    Warehouses ||--o{ SalesOrders : "WarehouseID"
    SalesReps |o--o{ SalesOrders : "SalesRepID"
    SalesReps ||--o{ RepDays : "SalesRepID"
    Warehouses ||--o{ RepDays : "WarehouseID"
    CashBoxes ||--o{ RepDays : "CashBoxID"
    Warehouses ||--o{ RepDays : "StoreID"
    StockCounts |o--o{ RepDays : "StockCountID"
    CashClosings |o--o{ RepDays : "ClosingID"
    SalesReps ||--o{ RouteStops : "SalesRepID"
    Customers ||--o{ RouteStops : "CustomerID"
    SalesReps ||--o{ RepVisits : "SalesRepID"
    Customers ||--o{ RepVisits : "CustomerID"
    SalesReps ||--o{ RepPrices : "SalesRepID"
    Products ||--o{ RepPrices : "ProductID"
    SalesOrders ||--o{ SalesOrderLines : "SalesOrderID"
    Products ||--o{ SalesOrderLines : "ProductID"
    Categories |o--o{ StockCounts : "CategoryID"
    Employees ||--o{ StockCounts : "EmployeeID"
    Warehouses ||--o{ StockCounts : "WarehouseID"
    StockCounts ||--o{ StockCountDetails : "StockCountID"
    Products ||--o{ StockCountDetails : "ProductID"
    AuditLog ||--o{ AuditChanges : "LogID"
```

`||--o{` = إلزامي (كل سجل في الجدول الفرعي يجب أن يرتبط بسجل في الأصلي)، `|o--o{` = اختياري (الحقل يمكن أن يكون فارغًا).

## قائمة العلاقات

| # | اسم العلاقة | الجدول الأصلي (1) | المفتاح | الجدول الفرعي (∞) | الحقل المرتبط | إلزامي | القاعدة |
|---|---|---|---|---|---|---|---|
| 1 | `FK_Settings_DefaultCustomerID` | Customers (العملاء) | `CustomerID` | Settings (إعدادات المحل) | `DefaultCustomerID` | ✔ | فرض التكامل |
| 2 | `FK_Settings_DefaultBankID` | Banks (البنوك) | `BankID` | Settings (إعدادات المحل) | `DefaultBankID` |  | فرض التكامل |
| 3 | `FK_RolePermissions_RoleID` | Roles (الأدوار) | `RoleID` | RolePermissions (صلاحيات الأدوار) | `RoleID` | ✔ | فرض التكامل + حذف متتالٍ |
| 4 | `FK_RolePermissions_PermissionKey` | Permissions (الصلاحيات) | `PermissionKey` | RolePermissions (صلاحيات الأدوار) | `PermissionKey` | ✔ | فرض التكامل + حذف متتالٍ + تحديث متتالٍ |
| 5 | `FK_Employees_RoleID` | Roles (الأدوار) | `RoleID` | Employees (الموظفون والمستخدمون) | `RoleID` | ✔ | فرض التكامل |
| 6 | `FK_Employees_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | Employees (الموظفون والمستخدمون) | `CashBoxID` |  | فرض التكامل |
| 7 | `FK_Employees_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | Employees (الموظفون والمستخدمون) | `CostCenterID` |  | فرض التكامل |
| 8 | `FK_Employees_WarehouseID` | Warehouses (المخازن) | `WarehouseID` | Employees (الموظفون والمستخدمون) | `WarehouseID` |  | فرض التكامل |
| 9 | `FK_UserScreens_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | UserScreens (صلاحيات الشاشات للمستخدم) | `EmployeeID` | ✔ | فرض التكامل |
| 10 | `FK_UserScreens_ScreenName` | Screens (الشاشات) | `ScreenName` | UserScreens (صلاحيات الشاشات للمستخدم) | `ScreenName` | ✔ | فرض التكامل + تحديث متتالٍ |
| 11 | `FK_Activations_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | Activations (تفعيل البرنامج) | `EmployeeID` |  | فرض التكامل |
| 12 | `FK_CurrencyRates_CurrencyCode` | Currencies (العملات) | `CurrencyCode` | CurrencyRates (أسعار العملات) | `CurrencyCode` | ✔ | فرض التكامل + تحديث متتالٍ |
| 13 | `FK_CashBoxes_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | CashBoxes (الخزينة والصناديق) | `SalesRepID` |  | فرض التكامل |
| 14 | `FK_CashBoxUsers_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | CashBoxUsers (أمناء الصناديق) | `CashBoxID` | ✔ | فرض التكامل |
| 15 | `FK_CashBoxUsers_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | CashBoxUsers (أمناء الصناديق) | `EmployeeID` | ✔ | فرض التكامل |
| 16 | `FK_Warehouses_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | Warehouses (المخازن) | `CostCenterID` |  | فرض التكامل |
| 17 | `FK_Warehouses_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | Warehouses (المخازن) | `SalesRepID` |  | فرض التكامل |
| 18 | `FK_Suppliers_CurrencyCode` | Currencies (العملات) | `CurrencyCode` | Suppliers (الموردون) | `CurrencyCode` |  | فرض التكامل + تحديث متتالٍ |
| 19 | `FK_Customers_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | Customers (العملاء) | `SalesRepID` |  | فرض التكامل |
| 20 | `FK_Products_CategoryID` | Categories (التصنيفات) | `CategoryID` | Products (المنتجات) | `CategoryID` | ✔ | فرض التكامل |
| 21 | `FK_Products_UnitID` | Units (وحدات القياس) | `UnitID` | Products (المنتجات) | `UnitID` | ✔ | فرض التكامل |
| 22 | `FK_Products_SupplierID` | Suppliers (الموردون) | `SupplierID` | Products (المنتجات) | `SupplierID` |  | فرض التكامل |
| 23 | `FK_SalesInvoices_CustomerID` | Customers (العملاء) | `CustomerID` | SalesInvoices (فواتير المبيعات) | `CustomerID` | ✔ | فرض التكامل |
| 24 | `FK_SalesInvoices_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | SalesInvoices (فواتير المبيعات) | `EmployeeID` | ✔ | فرض التكامل |
| 25 | `FK_SalesInvoices_PaymentMethodID` | PaymentMethods (طرق الدفع) | `PaymentMethodID` | SalesInvoices (فواتير المبيعات) | `PaymentMethodID` |  | فرض التكامل |
| 26 | `FK_SalesInvoices_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | SalesInvoices (فواتير المبيعات) | `CashBoxID` |  | فرض التكامل |
| 27 | `FK_SalesInvoices_BankID` | Banks (البنوك) | `BankID` | SalesInvoices (فواتير المبيعات) | `BankID` |  | فرض التكامل |
| 28 | `FK_SalesInvoices_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | SalesInvoices (فواتير المبيعات) | `CostCenterID` |  | فرض التكامل |
| 29 | `FK_SalesInvoices_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | SalesInvoices (فواتير المبيعات) | `SalesRepID` |  | فرض التكامل |
| 30 | `FK_SalesInvoices_WarehouseID` | Warehouses (المخازن) | `WarehouseID` | SalesInvoices (فواتير المبيعات) | `WarehouseID` | ✔ | فرض التكامل |
| 31 | `FK_SalesInvoiceDetails_SalesInvoiceID` | SalesInvoices (فواتير المبيعات) | `SalesInvoiceID` | SalesInvoiceDetails (تفاصيل فواتير المبيعات) | `SalesInvoiceID` | ✔ | فرض التكامل + حذف متتالٍ |
| 32 | `FK_SalesInvoiceDetails_ProductID` | Products (المنتجات) | `ProductID` | SalesInvoiceDetails (تفاصيل فواتير المبيعات) | `ProductID` | ✔ | فرض التكامل |
| 33 | `FK_SalesReturns_SalesInvoiceID` | SalesInvoices (فواتير المبيعات) | `SalesInvoiceID` | SalesReturns (مرتجعات المبيعات) | `SalesInvoiceID` | ✔ | فرض التكامل |
| 34 | `FK_SalesReturns_CustomerID` | Customers (العملاء) | `CustomerID` | SalesReturns (مرتجعات المبيعات) | `CustomerID` | ✔ | فرض التكامل |
| 35 | `FK_SalesReturns_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | SalesReturns (مرتجعات المبيعات) | `EmployeeID` | ✔ | فرض التكامل |
| 36 | `FK_SalesReturns_PaymentMethodID` | PaymentMethods (طرق الدفع) | `PaymentMethodID` | SalesReturns (مرتجعات المبيعات) | `PaymentMethodID` |  | فرض التكامل |
| 37 | `FK_SalesReturns_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | SalesReturns (مرتجعات المبيعات) | `CashBoxID` |  | فرض التكامل |
| 38 | `FK_SalesReturns_BankID` | Banks (البنوك) | `BankID` | SalesReturns (مرتجعات المبيعات) | `BankID` |  | فرض التكامل |
| 39 | `FK_SalesReturns_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | SalesReturns (مرتجعات المبيعات) | `CostCenterID` |  | فرض التكامل |
| 40 | `FK_SalesReturns_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | SalesReturns (مرتجعات المبيعات) | `SalesRepID` |  | فرض التكامل |
| 41 | `FK_SalesReturns_WarehouseID` | Warehouses (المخازن) | `WarehouseID` | SalesReturns (مرتجعات المبيعات) | `WarehouseID` | ✔ | فرض التكامل |
| 42 | `FK_SalesReturnDetails_SalesReturnID` | SalesReturns (مرتجعات المبيعات) | `SalesReturnID` | SalesReturnDetails (تفاصيل مرتجعات المبيعات) | `SalesReturnID` | ✔ | فرض التكامل + حذف متتالٍ |
| 43 | `FK_SalesReturnDetails_SalesDetailID` | SalesInvoiceDetails (تفاصيل فواتير المبيعات) | `SalesDetailID` | SalesReturnDetails (تفاصيل مرتجعات المبيعات) | `SalesDetailID` | ✔ | فرض التكامل |
| 44 | `FK_SalesReturnDetails_ProductID` | Products (المنتجات) | `ProductID` | SalesReturnDetails (تفاصيل مرتجعات المبيعات) | `ProductID` | ✔ | فرض التكامل |
| 45 | `FK_PurchaseInvoices_SupplierID` | Suppliers (الموردون) | `SupplierID` | PurchaseInvoices (فواتير المشتريات) | `SupplierID` | ✔ | فرض التكامل |
| 46 | `FK_PurchaseInvoices_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | PurchaseInvoices (فواتير المشتريات) | `EmployeeID` | ✔ | فرض التكامل |
| 47 | `FK_PurchaseInvoices_PaymentMethodID` | PaymentMethods (طرق الدفع) | `PaymentMethodID` | PurchaseInvoices (فواتير المشتريات) | `PaymentMethodID` |  | فرض التكامل |
| 48 | `FK_PurchaseInvoices_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | PurchaseInvoices (فواتير المشتريات) | `CashBoxID` |  | فرض التكامل |
| 49 | `FK_PurchaseInvoices_BankID` | Banks (البنوك) | `BankID` | PurchaseInvoices (فواتير المشتريات) | `BankID` |  | فرض التكامل |
| 50 | `FK_PurchaseInvoices_CurrencyCode` | Currencies (العملات) | `CurrencyCode` | PurchaseInvoices (فواتير المشتريات) | `CurrencyCode` |  | فرض التكامل + تحديث متتالٍ |
| 51 | `FK_PurchaseInvoices_WarehouseID` | Warehouses (المخازن) | `WarehouseID` | PurchaseInvoices (فواتير المشتريات) | `WarehouseID` | ✔ | فرض التكامل |
| 52 | `FK_PurchaseInvoiceDetails_PurchaseInvoiceID` | PurchaseInvoices (فواتير المشتريات) | `PurchaseInvoiceID` | PurchaseInvoiceDetails (تفاصيل فواتير المشتريات) | `PurchaseInvoiceID` | ✔ | فرض التكامل + حذف متتالٍ |
| 53 | `FK_PurchaseInvoiceDetails_ProductID` | Products (المنتجات) | `ProductID` | PurchaseInvoiceDetails (تفاصيل فواتير المشتريات) | `ProductID` | ✔ | فرض التكامل |
| 54 | `FK_PurchaseReturns_PurchaseInvoiceID` | PurchaseInvoices (فواتير المشتريات) | `PurchaseInvoiceID` | PurchaseReturns (مرتجعات المشتريات) | `PurchaseInvoiceID` | ✔ | فرض التكامل |
| 55 | `FK_PurchaseReturns_SupplierID` | Suppliers (الموردون) | `SupplierID` | PurchaseReturns (مرتجعات المشتريات) | `SupplierID` | ✔ | فرض التكامل |
| 56 | `FK_PurchaseReturns_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | PurchaseReturns (مرتجعات المشتريات) | `EmployeeID` | ✔ | فرض التكامل |
| 57 | `FK_PurchaseReturns_PaymentMethodID` | PaymentMethods (طرق الدفع) | `PaymentMethodID` | PurchaseReturns (مرتجعات المشتريات) | `PaymentMethodID` |  | فرض التكامل |
| 58 | `FK_PurchaseReturns_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | PurchaseReturns (مرتجعات المشتريات) | `CashBoxID` |  | فرض التكامل |
| 59 | `FK_PurchaseReturns_BankID` | Banks (البنوك) | `BankID` | PurchaseReturns (مرتجعات المشتريات) | `BankID` |  | فرض التكامل |
| 60 | `FK_PurchaseReturns_CurrencyCode` | Currencies (العملات) | `CurrencyCode` | PurchaseReturns (مرتجعات المشتريات) | `CurrencyCode` |  | فرض التكامل + تحديث متتالٍ |
| 61 | `FK_PurchaseReturns_WarehouseID` | Warehouses (المخازن) | `WarehouseID` | PurchaseReturns (مرتجعات المشتريات) | `WarehouseID` | ✔ | فرض التكامل |
| 62 | `FK_PurchaseReturnDetails_PurchaseReturnID` | PurchaseReturns (مرتجعات المشتريات) | `PurchaseReturnID` | PurchaseReturnDetails (تفاصيل مرتجعات المشتريات) | `PurchaseReturnID` | ✔ | فرض التكامل + حذف متتالٍ |
| 63 | `FK_PurchaseReturnDetails_PurchaseDetailID` | PurchaseInvoiceDetails (تفاصيل فواتير المشتريات) | `PurchaseDetailID` | PurchaseReturnDetails (تفاصيل مرتجعات المشتريات) | `PurchaseDetailID` | ✔ | فرض التكامل |
| 64 | `FK_PurchaseReturnDetails_ProductID` | Products (المنتجات) | `ProductID` | PurchaseReturnDetails (تفاصيل مرتجعات المشتريات) | `ProductID` | ✔ | فرض التكامل |
| 65 | `FK_CustomerPayments_CustomerID` | Customers (العملاء) | `CustomerID` | CustomerPayments (دفعات العملاء (سندات القبض)) | `CustomerID` | ✔ | فرض التكامل |
| 66 | `FK_CustomerPayments_PaymentMethodID` | PaymentMethods (طرق الدفع) | `PaymentMethodID` | CustomerPayments (دفعات العملاء (سندات القبض)) | `PaymentMethodID` | ✔ | فرض التكامل |
| 67 | `FK_CustomerPayments_SalesInvoiceID` | SalesInvoices (فواتير المبيعات) | `SalesInvoiceID` | CustomerPayments (دفعات العملاء (سندات القبض)) | `SalesInvoiceID` |  | فرض التكامل |
| 68 | `FK_CustomerPayments_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | CustomerPayments (دفعات العملاء (سندات القبض)) | `EmployeeID` | ✔ | فرض التكامل |
| 69 | `FK_CustomerPayments_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | CustomerPayments (دفعات العملاء (سندات القبض)) | `CashBoxID` |  | فرض التكامل |
| 70 | `FK_CustomerPayments_BankID` | Banks (البنوك) | `BankID` | CustomerPayments (دفعات العملاء (سندات القبض)) | `BankID` |  | فرض التكامل |
| 71 | `FK_CustomerPayments_CurrencyCode` | Currencies (العملات) | `CurrencyCode` | CustomerPayments (دفعات العملاء (سندات القبض)) | `CurrencyCode` |  | فرض التكامل + تحديث متتالٍ |
| 72 | `FK_CustomerPayments_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | CustomerPayments (دفعات العملاء (سندات القبض)) | `SalesRepID` |  | فرض التكامل |
| 73 | `FK_SupplierPayments_SupplierID` | Suppliers (الموردون) | `SupplierID` | SupplierPayments (دفعات الموردين (سندات الصرف)) | `SupplierID` | ✔ | فرض التكامل |
| 74 | `FK_SupplierPayments_PaymentMethodID` | PaymentMethods (طرق الدفع) | `PaymentMethodID` | SupplierPayments (دفعات الموردين (سندات الصرف)) | `PaymentMethodID` | ✔ | فرض التكامل |
| 75 | `FK_SupplierPayments_PurchaseInvoiceID` | PurchaseInvoices (فواتير المشتريات) | `PurchaseInvoiceID` | SupplierPayments (دفعات الموردين (سندات الصرف)) | `PurchaseInvoiceID` |  | فرض التكامل |
| 76 | `FK_SupplierPayments_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | SupplierPayments (دفعات الموردين (سندات الصرف)) | `EmployeeID` | ✔ | فرض التكامل |
| 77 | `FK_SupplierPayments_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | SupplierPayments (دفعات الموردين (سندات الصرف)) | `CashBoxID` |  | فرض التكامل |
| 78 | `FK_SupplierPayments_BankID` | Banks (البنوك) | `BankID` | SupplierPayments (دفعات الموردين (سندات الصرف)) | `BankID` |  | فرض التكامل |
| 79 | `FK_SupplierPayments_CurrencyCode` | Currencies (العملات) | `CurrencyCode` | SupplierPayments (دفعات الموردين (سندات الصرف)) | `CurrencyCode` |  | فرض التكامل + تحديث متتالٍ |
| 80 | `FK_BankTransactions_BankID` | Banks (البنوك) | `BankID` | BankTransactions (الحركات البنكية) | `BankID` | ✔ | فرض التكامل |
| 81 | `FK_BankTransactions_ToBankID` | Banks (البنوك) | `BankID` | BankTransactions (الحركات البنكية) | `ToBankID` |  | فرض التكامل |
| 82 | `FK_BankTransactions_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | BankTransactions (الحركات البنكية) | `CashBoxID` |  | فرض التكامل |
| 83 | `FK_BankTransactions_CounterAccount` | Accounts (دليل الحسابات (شجرة الحسابات)) | `AccountCode` | BankTransactions (الحركات البنكية) | `CounterAccount` |  | فرض التكامل |
| 84 | `FK_BankTransactions_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | BankTransactions (الحركات البنكية) | `EmployeeID` | ✔ | فرض التكامل |
| 85 | `FK_Cheques_CustomerID` | Customers (العملاء) | `CustomerID` | Cheques (الشيكات الواردة والصادرة) | `CustomerID` |  | فرض التكامل |
| 86 | `FK_Cheques_SupplierID` | Suppliers (الموردون) | `SupplierID` | Cheques (الشيكات الواردة والصادرة) | `SupplierID` |  | فرض التكامل |
| 87 | `FK_Cheques_BankID` | Banks (البنوك) | `BankID` | Cheques (الشيكات الواردة والصادرة) | `BankID` |  | فرض التكامل |
| 88 | `FK_Cheques_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | Cheques (الشيكات الواردة والصادرة) | `EmployeeID` | ✔ | فرض التكامل |
| 89 | `FK_FixedAssets_AssetAccount` | Accounts (دليل الحسابات (شجرة الحسابات)) | `AccountCode` | FixedAssets (الأصول الثابتة) | `AssetAccount` | ✔ | فرض التكامل |
| 90 | `FK_FixedAssets_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | FixedAssets (الأصول الثابتة) | `CostCenterID` |  | فرض التكامل |
| 91 | `FK_FixedAssets_BankID` | Banks (البنوك) | `BankID` | FixedAssets (الأصول الثابتة) | `BankID` |  | فرض التكامل |
| 92 | `FK_FixedAssets_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | FixedAssets (الأصول الثابتة) | `CashBoxID` |  | فرض التكامل |
| 93 | `FK_FixedAssets_CounterAccount` | Accounts (دليل الحسابات (شجرة الحسابات)) | `AccountCode` | FixedAssets (الأصول الثابتة) | `CounterAccount` |  | فرض التكامل |
| 94 | `FK_FixedAssets_DisposalBankID` | Banks (البنوك) | `BankID` | FixedAssets (الأصول الثابتة) | `DisposalBankID` |  | فرض التكامل |
| 95 | `FK_FixedAssets_DisposalCashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | FixedAssets (الأصول الثابتة) | `DisposalCashBoxID` |  | فرض التكامل |
| 96 | `FK_FixedAssets_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | FixedAssets (الأصول الثابتة) | `EmployeeID` | ✔ | فرض التكامل |
| 97 | `FK_AssetDepreciations_RunID` | DepreciationRuns (قيود الإهلاك الشهرية) | `RunID` | AssetDepreciations (إهلاك كل أصل في كل شهر) | `RunID` | ✔ | فرض التكامل + حذف متتالٍ |
| 98 | `FK_AssetDepreciations_AssetID` | FixedAssets (الأصول الثابتة) | `AssetID` | AssetDepreciations (إهلاك كل أصل في كل شهر) | `AssetID` | ✔ | فرض التكامل |
| 99 | `FK_AssetRevaluations_AssetID` | FixedAssets (الأصول الثابتة) | `AssetID` | AssetRevaluations (إعادة تقييم الأصول الثابتة) | `AssetID` | ✔ | فرض التكامل |
| 100 | `FK_AssetTransfers_AssetID` | FixedAssets (الأصول الثابتة) | `AssetID` | AssetTransfers (نقل الأصول الثابتة بين مراكز التكلفة) | `AssetID` | ✔ | فرض التكامل |
| 101 | `FK_AssetTransfers_FromCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | AssetTransfers (نقل الأصول الثابتة بين مراكز التكلفة) | `FromCenterID` |  | فرض التكامل |
| 102 | `FK_AssetTransfers_ToCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | AssetTransfers (نقل الأصول الثابتة بين مراكز التكلفة) | `ToCenterID` |  | فرض التكامل |
| 103 | `FK_SalesReps_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | SalesReps (المندوبين) | `EmployeeID` |  | فرض التكامل |
| 104 | `FK_SalesReps_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | SalesReps (المندوبين) | `CostCenterID` |  | فرض التكامل |
| 105 | `FK_SalesRepTargets_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | SalesRepTargets (أهداف المندوبين) | `SalesRepID` | ✔ | فرض التكامل + حذف متتالٍ |
| 106 | `FK_CommissionLines_CommissionRunID` | CommissionRuns (مسيرات العمولات) | `CommissionRunID` | CommissionLines (أسطر مسير العمولات) | `CommissionRunID` | ✔ | فرض التكامل + حذف متتالٍ |
| 107 | `FK_CommissionLines_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | CommissionLines (أسطر مسير العمولات) | `SalesRepID` | ✔ | فرض التكامل |
| 108 | `FK_CommissionLines_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | CommissionLines (أسطر مسير العمولات) | `CostCenterID` |  | فرض التكامل |
| 109 | `FK_BudgetLines_BudgetID` | Budgets (الموازنات التقديرية) | `BudgetID` | BudgetLines (أسطر الموازنة) | `BudgetID` | ✔ | فرض التكامل + حذف متتالٍ |
| 110 | `FK_BudgetLines_AccountCode` | Accounts (دليل الحسابات (شجرة الحسابات)) | `AccountCode` | BudgetLines (أسطر الموازنة) | `AccountCode` | ✔ | فرض التكامل |
| 111 | `FK_BudgetLines_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | BudgetLines (أسطر الموازنة) | `CostCenterID` |  | فرض التكامل |
| 112 | `FK_PayrollRuns_BankID` | Banks (البنوك) | `BankID` | PayrollRuns (مسيرات الرواتب) | `BankID` |  | فرض التكامل |
| 113 | `FK_PayrollRuns_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | PayrollRuns (مسيرات الرواتب) | `CashBoxID` |  | فرض التكامل |
| 114 | `FK_PayrollRuns_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | PayrollRuns (مسيرات الرواتب) | `EmployeeID` | ✔ | فرض التكامل |
| 115 | `FK_PayrollLines_PayrollRunID` | PayrollRuns (مسيرات الرواتب) | `PayrollRunID` | PayrollLines (أسطر مسير الرواتب) | `PayrollRunID` | ✔ | فرض التكامل + حذف متتالٍ |
| 116 | `FK_PayrollLines_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | PayrollLines (أسطر مسير الرواتب) | `EmployeeID` | ✔ | فرض التكامل |
| 117 | `FK_PayrollLines_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | PayrollLines (أسطر مسير الرواتب) | `CostCenterID` |  | فرض التكامل |
| 118 | `FK_BankReconciliations_BankID` | Banks (البنوك) | `BankID` | BankReconciliations (التسويات البنكية) | `BankID` | ✔ | فرض التكامل |
| 119 | `FK_BankClearings_ReconciliationID` | BankReconciliations (التسويات البنكية) | `ReconciliationID` | BankClearings (حركات الدفاتر المطابقة لكشف البنك) | `ReconciliationID` | ✔ | فرض التكامل + حذف متتالٍ |
| 120 | `FK_BankClearings_BankID` | Banks (البنوك) | `BankID` | BankClearings (حركات الدفاتر المطابقة لكشف البنك) | `BankID` | ✔ | فرض التكامل |
| 121 | `FK_CustomerAllocations_PaymentID` | CustomerPayments (دفعات العملاء (سندات القبض)) | `PaymentID` | CustomerAllocations (ربط سندات القبض بالفواتير) | `PaymentID` | ✔ | فرض التكامل + حذف متتالٍ |
| 122 | `FK_CustomerAllocations_SalesInvoiceID` | SalesInvoices (فواتير المبيعات) | `SalesInvoiceID` | CustomerAllocations (ربط سندات القبض بالفواتير) | `SalesInvoiceID` | ✔ | فرض التكامل |
| 123 | `FK_SupplierAllocations_PaymentID` | SupplierPayments (دفعات الموردين (سندات الصرف)) | `PaymentID` | SupplierAllocations (ربط سندات الصرف بفواتير الشراء) | `PaymentID` | ✔ | فرض التكامل + حذف متتالٍ |
| 124 | `FK_SupplierAllocations_PurchaseInvoiceID` | PurchaseInvoices (فواتير المشتريات) | `PurchaseInvoiceID` | SupplierAllocations (ربط سندات الصرف بفواتير الشراء) | `PurchaseInvoiceID` | ✔ | فرض التكامل |
| 125 | `FK_Expenses_ExpenseTypeID` | ExpenseTypes (أنواع المصروفات) | `ExpenseTypeID` | Expenses (المصروفات) | `ExpenseTypeID` | ✔ | فرض التكامل |
| 126 | `FK_Expenses_PaymentMethodID` | PaymentMethods (طرق الدفع) | `PaymentMethodID` | Expenses (المصروفات) | `PaymentMethodID` |  | فرض التكامل |
| 127 | `FK_Expenses_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | Expenses (المصروفات) | `EmployeeID` | ✔ | فرض التكامل |
| 128 | `FK_Expenses_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | Expenses (المصروفات) | `CashBoxID` |  | فرض التكامل |
| 129 | `FK_Expenses_BankID` | Banks (البنوك) | `BankID` | Expenses (المصروفات) | `BankID` |  | فرض التكامل |
| 130 | `FK_Expenses_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | Expenses (المصروفات) | `CostCenterID` |  | فرض التكامل |
| 131 | `FK_Expenses_RecurringID` | RecurringExpenses (المصروفات المتكررة) | `RecurringID` | Expenses (المصروفات) | `RecurringID` |  | فرض التكامل |
| 132 | `FK_Expenses_CurrencyCode` | Currencies (العملات) | `CurrencyCode` | Expenses (المصروفات) | `CurrencyCode` |  | فرض التكامل + تحديث متتالٍ |
| 133 | `FK_RecurringExpenses_ExpenseTypeID` | ExpenseTypes (أنواع المصروفات) | `ExpenseTypeID` | RecurringExpenses (المصروفات المتكررة) | `ExpenseTypeID` | ✔ | فرض التكامل |
| 134 | `FK_RecurringExpenses_PaymentMethodID` | PaymentMethods (طرق الدفع) | `PaymentMethodID` | RecurringExpenses (المصروفات المتكررة) | `PaymentMethodID` | ✔ | فرض التكامل |
| 135 | `FK_RecurringExpenses_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | RecurringExpenses (المصروفات المتكررة) | `CashBoxID` |  | فرض التكامل |
| 136 | `FK_RecurringExpenses_BankID` | Banks (البنوك) | `BankID` | RecurringExpenses (المصروفات المتكررة) | `BankID` |  | فرض التكامل |
| 137 | `FK_RecurringExpenses_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | RecurringExpenses (المصروفات المتكررة) | `CostCenterID` |  | فرض التكامل |
| 138 | `FK_CashVouchers_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | CashVouchers (سندات النقدية) | `CashBoxID` | ✔ | فرض التكامل |
| 139 | `FK_CashVouchers_ToCashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | CashVouchers (سندات النقدية) | `ToCashBoxID` |  | فرض التكامل |
| 140 | `FK_CashVouchers_ExpenseID` | Expenses (المصروفات) | `ExpenseID` | CashVouchers (سندات النقدية) | `ExpenseID` |  | فرض التكامل |
| 141 | `FK_CashVouchers_ClosingID` | CashClosings (تصفية يومية الكاشير) | `ClosingID` | CashVouchers (سندات النقدية) | `ClosingID` |  | فرض التكامل |
| 142 | `FK_CashVouchers_AdvanceEmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | CashVouchers (سندات النقدية) | `AdvanceEmployeeID` |  | فرض التكامل |
| 143 | `FK_CashVouchers_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | CashVouchers (سندات النقدية) | `CostCenterID` |  | فرض التكامل |
| 144 | `FK_CashVouchers_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | CashVouchers (سندات النقدية) | `EmployeeID` | ✔ | فرض التكامل |
| 145 | `FK_CashVouchers_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | CashVouchers (سندات النقدية) | `SalesRepID` |  | فرض التكامل |
| 146 | `FK_CashClosings_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | CashClosings (تصفية يومية الكاشير) | `CashBoxID` | ✔ | فرض التكامل |
| 147 | `FK_CashClosings_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | CashClosings (تصفية يومية الكاشير) | `EmployeeID` | ✔ | فرض التكامل |
| 148 | `FK_CashClosings_ToCashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | CashClosings (تصفية يومية الكاشير) | `ToCashBoxID` |  | فرض التكامل |
| 149 | `FK_JournalEntries_SourceType` | JournalSourceTypes (أنواع مصادر القيود) | `SourceType` | JournalEntries (قيود اليومية) | `SourceType` | ✔ | فرض التكامل + تحديث متتالٍ |
| 150 | `FK_JournalEntries_CurrencyCode` | Currencies (العملات) | `CurrencyCode` | JournalEntries (قيود اليومية) | `CurrencyCode` |  | فرض التكامل + تحديث متتالٍ |
| 151 | `FK_JournalLines_EntryID` | JournalEntries (قيود اليومية) | `EntryID` | JournalLines (أسطر القيود) | `EntryID` | ✔ | فرض التكامل + حذف متتالٍ |
| 152 | `FK_JournalLines_AccountCode` | Accounts (دليل الحسابات (شجرة الحسابات)) | `AccountCode` | JournalLines (أسطر القيود) | `AccountCode` | ✔ | فرض التكامل |
| 153 | `FK_JournalLines_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | JournalLines (أسطر القيود) | `CostCenterID` |  | فرض التكامل |
| 154 | `FK_FiscalYearClosingLines_YearClosingID` | FiscalYearClosings (إقفال السنوات المالية) | `YearClosingID` | FiscalYearClosingLines (أسطر قيود إقفال السنوات) | `YearClosingID` | ✔ | فرض التكامل + حذف متتالٍ |
| 155 | `FK_FiscalYearClosingLines_AccountCode` | Accounts (دليل الحسابات (شجرة الحسابات)) | `AccountCode` | FiscalYearClosingLines (أسطر قيود إقفال السنوات) | `AccountCode` | ✔ | فرض التكامل |
| 156 | `FK_VatReturns_PaidAccount` | Accounts (دليل الحسابات (شجرة الحسابات)) | `AccountCode` | VatReturns (إقرارات ضريبة القيمة المضافة) | `PaidAccount` |  | فرض التكامل |
| 157 | `FK_ManualEntries_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | ManualEntries (القيود اليدوية) | `EmployeeID` | ✔ | فرض التكامل |
| 158 | `FK_ManualEntries_CurrencyCode` | Currencies (العملات) | `CurrencyCode` | ManualEntries (القيود اليدوية) | `CurrencyCode` |  | فرض التكامل + تحديث متتالٍ |
| 159 | `FK_ManualEntryLines_ManualEntryID` | ManualEntries (القيود اليدوية) | `ManualEntryID` | ManualEntryLines (أسطر القيود اليدوية) | `ManualEntryID` | ✔ | فرض التكامل + حذف متتالٍ |
| 160 | `FK_ManualEntryLines_AccountCode` | Accounts (دليل الحسابات (شجرة الحسابات)) | `AccountCode` | ManualEntryLines (أسطر القيود اليدوية) | `AccountCode` | ✔ | فرض التكامل |
| 161 | `FK_ManualEntryLines_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | ManualEntryLines (أسطر القيود اليدوية) | `CostCenterID` |  | فرض التكامل |
| 162 | `FK_OpeningBalances_AccountCode` | Accounts (دليل الحسابات (شجرة الحسابات)) | `AccountCode` | OpeningBalances (الأرصدة الافتتاحية للحسابات) | `AccountCode` | ✔ | فرض التكامل |
| 163 | `FK_OpeningStock_ProductID` | Products (المنتجات) | `ProductID` | OpeningStock (الرصيد الافتتاحي للمخزون) | `ProductID` | ✔ | فرض التكامل |
| 164 | `FK_OpeningStock_WarehouseID` | Warehouses (المخازن) | `WarehouseID` | OpeningStock (الرصيد الافتتاحي للمخزون) | `WarehouseID` | ✔ | فرض التكامل |
| 165 | `FK_InventoryTransactions_ProductID` | Products (المنتجات) | `ProductID` | InventoryTransactions (حركة المخزون) | `ProductID` | ✔ | فرض التكامل |
| 166 | `FK_InventoryTransactions_TransactionTypeID` | TransactionTypes (أنواع حركات المخزون) | `TransactionTypeID` | InventoryTransactions (حركة المخزون) | `TransactionTypeID` | ✔ | فرض التكامل |
| 167 | `FK_InventoryTransactions_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | InventoryTransactions (حركة المخزون) | `EmployeeID` |  | فرض التكامل |
| 168 | `FK_InventoryTransactions_WarehouseID` | Warehouses (المخازن) | `WarehouseID` | InventoryTransactions (حركة المخزون) | `WarehouseID` | ✔ | فرض التكامل |
| 169 | `FK_StockTransfers_FromWarehouseID` | Warehouses (المخازن) | `WarehouseID` | StockTransfers (التحويل بين المخازن) | `FromWarehouseID` | ✔ | فرض التكامل |
| 170 | `FK_StockTransfers_ToWarehouseID` | Warehouses (المخازن) | `WarehouseID` | StockTransfers (التحويل بين المخازن) | `ToWarehouseID` | ✔ | فرض التكامل |
| 171 | `FK_StockTransferLines_StockTransferID` | StockTransfers (التحويل بين المخازن) | `StockTransferID` | StockTransferLines (أسطر التحويل بين المخازن) | `StockTransferID` | ✔ | فرض التكامل + حذف متتالٍ |
| 172 | `FK_StockTransferLines_ProductID` | Products (المنتجات) | `ProductID` | StockTransferLines (أسطر التحويل بين المخازن) | `ProductID` | ✔ | فرض التكامل |
| 173 | `FK_StockParties_AccountCode` | Accounts (دليل الحسابات (شجرة الحسابات)) | `AccountCode` | StockParties (جهات الصرف والاستلام المخزني) | `AccountCode` |  | فرض التكامل |
| 174 | `FK_StockParties_CostCenterID` | CostCenters (مراكز التكلفة والفروع) | `CostCenterID` | StockParties (جهات الصرف والاستلام المخزني) | `CostCenterID` |  | فرض التكامل |
| 175 | `FK_StockDocuments_WarehouseID` | Warehouses (المخازن) | `WarehouseID` | StockDocuments (سندات الصرف والاستلام المخزني) | `WarehouseID` | ✔ | فرض التكامل |
| 176 | `FK_StockDocuments_StockPartyID` | StockParties (جهات الصرف والاستلام المخزني) | `StockPartyID` | StockDocuments (سندات الصرف والاستلام المخزني) | `StockPartyID` | ✔ | فرض التكامل |
| 177 | `FK_StockDocumentLines_StockDocumentID` | StockDocuments (سندات الصرف والاستلام المخزني) | `StockDocumentID` | StockDocumentLines (أسطر سندات الصرف والاستلام المخزني) | `StockDocumentID` | ✔ | فرض التكامل + حذف متتالٍ |
| 178 | `FK_StockDocumentLines_ProductID` | Products (المنتجات) | `ProductID` | StockDocumentLines (أسطر سندات الصرف والاستلام المخزني) | `ProductID` | ✔ | فرض التكامل |
| 179 | `FK_StockNotes_WarehouseID` | Warehouses (المخازن) | `WarehouseID` | StockNotes (أذونات التسليم والاستلام) | `WarehouseID` | ✔ | فرض التكامل |
| 180 | `FK_StockNotes_CustomerID` | Customers (العملاء) | `CustomerID` | StockNotes (أذونات التسليم والاستلام) | `CustomerID` |  | فرض التكامل |
| 181 | `FK_StockNotes_SupplierID` | Suppliers (الموردون) | `SupplierID` | StockNotes (أذونات التسليم والاستلام) | `SupplierID` |  | فرض التكامل |
| 182 | `FK_StockNotes_SalesOrderID` | SalesOrders (طلبيات العملاء) | `SalesOrderID` | StockNotes (أذونات التسليم والاستلام) | `SalesOrderID` |  | فرض التكامل |
| 183 | `FK_StockNotes_SalesInvoiceID` | SalesInvoices (فواتير المبيعات) | `SalesInvoiceID` | StockNotes (أذونات التسليم والاستلام) | `SalesInvoiceID` |  | فرض التكامل |
| 184 | `FK_StockNotes_PurchaseInvoiceID` | PurchaseInvoices (فواتير المشتريات) | `PurchaseInvoiceID` | StockNotes (أذونات التسليم والاستلام) | `PurchaseInvoiceID` |  | فرض التكامل |
| 185 | `FK_StockNoteLines_StockNoteID` | StockNotes (أذونات التسليم والاستلام) | `StockNoteID` | StockNoteLines (أسطر أذونات التسليم والاستلام) | `StockNoteID` | ✔ | فرض التكامل + حذف متتالٍ |
| 186 | `FK_StockNoteLines_ProductID` | Products (المنتجات) | `ProductID` | StockNoteLines (أسطر أذونات التسليم والاستلام) | `ProductID` | ✔ | فرض التكامل |
| 187 | `FK_StockNoteLines_SalesOrderLineID` | SalesOrderLines (أسطر طلبيات العملاء) | `SalesOrderLineID` | StockNoteLines (أسطر أذونات التسليم والاستلام) | `SalesOrderLineID` |  | فرض التكامل |
| 188 | `FK_SalesOrders_CustomerID` | Customers (العملاء) | `CustomerID` | SalesOrders (طلبيات العملاء) | `CustomerID` | ✔ | فرض التكامل |
| 189 | `FK_SalesOrders_WarehouseID` | Warehouses (المخازن) | `WarehouseID` | SalesOrders (طلبيات العملاء) | `WarehouseID` | ✔ | فرض التكامل |
| 190 | `FK_SalesOrders_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | SalesOrders (طلبيات العملاء) | `SalesRepID` |  | فرض التكامل |
| 191 | `FK_RepDays_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | RepDays (أيام المندوبين) | `SalesRepID` | ✔ | فرض التكامل |
| 192 | `FK_RepDays_WarehouseID` | Warehouses (المخازن) | `WarehouseID` | RepDays (أيام المندوبين) | `WarehouseID` | ✔ | فرض التكامل |
| 193 | `FK_RepDays_CashBoxID` | CashBoxes (الخزينة والصناديق) | `CashBoxID` | RepDays (أيام المندوبين) | `CashBoxID` | ✔ | فرض التكامل |
| 194 | `FK_RepDays_StoreID` | Warehouses (المخازن) | `WarehouseID` | RepDays (أيام المندوبين) | `StoreID` | ✔ | فرض التكامل |
| 195 | `FK_RepDays_StockCountID` | StockCounts (جلسات الجرد) | `StockCountID` | RepDays (أيام المندوبين) | `StockCountID` |  | فرض التكامل |
| 196 | `FK_RepDays_ClosingID` | CashClosings (تصفية يومية الكاشير) | `ClosingID` | RepDays (أيام المندوبين) | `ClosingID` |  | فرض التكامل |
| 197 | `FK_RouteStops_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | RouteStops (خط سير المندوبين) | `SalesRepID` | ✔ | فرض التكامل + حذف متتالٍ |
| 198 | `FK_RouteStops_CustomerID` | Customers (العملاء) | `CustomerID` | RouteStops (خط سير المندوبين) | `CustomerID` | ✔ | فرض التكامل |
| 199 | `FK_RepVisits_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | RepVisits (زيارات المندوبين) | `SalesRepID` | ✔ | فرض التكامل |
| 200 | `FK_RepVisits_CustomerID` | Customers (العملاء) | `CustomerID` | RepVisits (زيارات المندوبين) | `CustomerID` | ✔ | فرض التكامل |
| 201 | `FK_RepPrices_SalesRepID` | SalesReps (المندوبين) | `SalesRepID` | RepPrices (أسعار المندوبين) | `SalesRepID` | ✔ | فرض التكامل + حذف متتالٍ |
| 202 | `FK_RepPrices_ProductID` | Products (المنتجات) | `ProductID` | RepPrices (أسعار المندوبين) | `ProductID` | ✔ | فرض التكامل |
| 203 | `FK_SalesOrderLines_SalesOrderID` | SalesOrders (طلبيات العملاء) | `SalesOrderID` | SalesOrderLines (أسطر طلبيات العملاء) | `SalesOrderID` | ✔ | فرض التكامل + حذف متتالٍ |
| 204 | `FK_SalesOrderLines_ProductID` | Products (المنتجات) | `ProductID` | SalesOrderLines (أسطر طلبيات العملاء) | `ProductID` | ✔ | فرض التكامل |
| 205 | `FK_StockCounts_CategoryID` | Categories (التصنيفات) | `CategoryID` | StockCounts (جلسات الجرد) | `CategoryID` |  | فرض التكامل |
| 206 | `FK_StockCounts_EmployeeID` | Employees (الموظفون والمستخدمون) | `EmployeeID` | StockCounts (جلسات الجرد) | `EmployeeID` | ✔ | فرض التكامل |
| 207 | `FK_StockCounts_WarehouseID` | Warehouses (المخازن) | `WarehouseID` | StockCounts (جلسات الجرد) | `WarehouseID` | ✔ | فرض التكامل |
| 208 | `FK_StockCountDetails_StockCountID` | StockCounts (جلسات الجرد) | `StockCountID` | StockCountDetails (تفاصيل الجرد) | `StockCountID` | ✔ | فرض التكامل + حذف متتالٍ |
| 209 | `FK_StockCountDetails_ProductID` | Products (المنتجات) | `ProductID` | StockCountDetails (تفاصيل الجرد) | `ProductID` | ✔ | فرض التكامل |
| 210 | `FK_AuditChanges_LogID` | AuditLog (سجل العمليات) | `LogID` | AuditChanges (تفاصيل سجل التدقيق) | `LogID` | ✔ | فرض التكامل + حذف متتالٍ |
