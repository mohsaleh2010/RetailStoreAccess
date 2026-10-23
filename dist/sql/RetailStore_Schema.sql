/*==============================================================================
  RetailStore - the back-end on SQL Server (plan 5A, docs/67-SQL-Server-Script.md)

  GENERATED FILE - do not edit by hand.
  Source of truth: tools/schema.py  ->  python3 tools/generate.py

  Creates a new database or upgrades an existing one (like BuildSchema in Access):
  every statement checks first, so the script can run again after each update.

  First time (SQL Server Management Studio, or sqlcmd):
      CREATE DATABASE RetailStore COLLATE Arabic_100_CI_AS;
  Then run this file in that database:
      sqlcmd -S .\SQLEXPRESS -d RetailStore -E -f 65001 -i RetailStore_Schema.sql

  97 tables, 1237 fields, 148 indexes, 231 relationships, 451 seed rows
==============================================================================*/
SET NOCOUNT ON;
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
SET XACT_ABORT ON;
GO

/*------------------------------------------------------------------------------
  1. Tables
------------------------------------------------------------------------------*/
IF OBJECT_ID(N'dbo.Settings', N'U') IS NULL
    CREATE TABLE [dbo].[Settings] (
        [SettingID] INT NOT NULL CONSTRAINT [DF_Settings_SettingID] DEFAULT (1),
        [StoreName] NVARCHAR(150) NOT NULL,
        [StoreNameEn] NVARCHAR(150) NULL,
        [VATNumber] NVARCHAR(15) NULL,
        [CRNumber] NVARCHAR(20) NULL,
        [BuildingNo] NVARCHAR(10) NULL,
        [StreetName] NVARCHAR(100) NULL,
        [District] NVARCHAR(100) NULL,
        [City] NVARCHAR(50) NULL,
        [PostalCode] NVARCHAR(10) NULL,
        [AdditionalNo] NVARCHAR(10) NULL,
        [CountryCode] NVARCHAR(2) NOT NULL CONSTRAINT [DF_Settings_CountryCode] DEFAULT (N'SA'),
        [Phone] NVARCHAR(20) NULL,
        [Email] NVARCHAR(100) NULL,
        [VATRate] MONEY NOT NULL CONSTRAINT [DF_Settings_VATRate] DEFAULT (0.15),
        [PricesIncludeVAT] BIT NOT NULL CONSTRAINT [DF_Settings_PricesIncludeVAT] DEFAULT (1),
        [AllowNegativeStock] BIT NOT NULL CONSTRAINT [DF_Settings_AllowNegativeStock] DEFAULT (0),
        [CurrencyCode] NVARCHAR(3) NOT NULL CONSTRAINT [DF_Settings_CurrencyCode] DEFAULT (N'SAR'),
        [DefaultCustomerID] INT NOT NULL CONSTRAINT [DF_Settings_DefaultCustomerID] DEFAULT (1),
        [ZatcaPhase] TINYINT NOT NULL CONSTRAINT [DF_Settings_ZatcaPhase] DEFAULT (1),
        [ZatcaEnvironment] NVARCHAR(20) NULL,
        [LastInvoiceHash] NVARCHAR(255) NULL,
        [BackupFolder] NVARCHAR(255) NULL,
        [BackupKeepCount] SMALLINT NOT NULL CONSTRAINT [DF_Settings_BackupKeepCount] DEFAULT (30),
        [SlowMovingDays] SMALLINT NOT NULL CONSTRAINT [DF_Settings_SlowMovingDays] DEFAULT (90),
        [ReceiptFooter] NVARCHAR(255) NULL,
        [LogoPath] NVARCHAR(255) NULL,
        [UpdatedAt] DATETIME NULL,
        [POSMode] NVARCHAR(10) NOT NULL CONSTRAINT [DF_Settings_POSMode] DEFAULT (N'RETAIL'),
        [ImagesFolder] NVARCHAR(255) NULL,
        [InvoicePrintMode] NVARCHAR(10) NOT NULL CONSTRAINT [DF_Settings_InvoicePrintMode] DEFAULT (N'PREVIEW'),
        [AllowAdminCompanyName] BIT NOT NULL CONSTRAINT [DF_Settings_AllowAdminCompanyName] DEFAULT (0),
        [ClosedThrough] DATETIME NULL,
        [OpeningDate] DATETIME NULL,
        [DefaultBankID] INT NULL,
        [GosiEmployeeRate] MONEY NOT NULL CONSTRAINT [DF_Settings_GosiEmployeeRate] DEFAULT (0.0975),
        [GosiEmployerRate] MONEY NOT NULL CONSTRAINT [DF_Settings_GosiEmployerRate] DEFAULT (0.1175),
        [GosiNonSaudiRate] MONEY NOT NULL CONSTRAINT [DF_Settings_GosiNonSaudiRate] DEFAULT (0.02),
        [GosiMaxWage] MONEY NOT NULL CONSTRAINT [DF_Settings_GosiMaxWage] DEFAULT (45000),
        [CreditBlockDays] SMALLINT NULL CONSTRAINT [DF_Settings_CreditBlockDays] DEFAULT (0),
        [EInvoiceEnabled] BIT NOT NULL CONSTRAINT [DF_Settings_EInvoiceEnabled] DEFAULT (0),
        [OpenSslPath] NVARCHAR(255) NULL,
        [ZatcaKeyFile] NVARCHAR(255) NULL,
        [ZatcaCertificate] NVARCHAR(MAX) NULL,
        [ZatcaDeviceSerial] NVARCHAR(36) NULL,
        [ZatcaBranchName] NVARCHAR(100) NULL,
        [ZatcaIndustry] NVARCHAR(100) NULL,
        [ZatcaCsr] NVARCHAR(MAX) NULL,
        [ZatcaComplianceToken] NVARCHAR(MAX) NULL,
        [ZatcaComplianceSecret] NVARCHAR(255) NULL,
        [ZatcaRequestId] NVARCHAR(30) NULL,
        [ZatcaProductionToken] NVARCHAR(MAX) NULL,
        [ZatcaProductionSecret] NVARCHAR(255) NULL,
        [ZatcaOnboardEnv] NVARCHAR(12) NULL,
        [ZatcaOnboardStage] NVARCHAR(20) NULL,
        [EtaClientId] NVARCHAR(100) NULL,
        [EtaClientSecret] NVARCHAR(255) NULL,
        [EtaPosSerial] NVARCHAR(100) NULL,
        [EtaPosOsVersion] NVARCHAR(50) NULL,
        [EtaPreSharedKey] NVARCHAR(255) NULL,
        [EtaBranchCode] NVARCHAR(10) NULL,
        [EtaActivityCode] NVARCHAR(8) NULL,
        [EtaGovernate] NVARCHAR(50) NULL,
        [EtaLastUUID] NVARCHAR(64) NULL,
        [EtaErpClientId] NVARCHAR(100) NULL,
        [EtaErpClientSecret] NVARCHAR(255) NULL,
        [EtaSignerPath] NVARCHAR(255) NULL,
        [EtaSignerArgs] NVARCHAR(255) NULL,
        [EtaTokenPin] NVARCHAR(50) NULL,
        [EInvoiceEnvironment] NVARCHAR(12) NOT NULL CONSTRAINT [DF_Settings_EInvoiceEnvironment] DEFAULT (N'TEST'),
        [ServerLiveSince] DATETIME NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Settings] PRIMARY KEY ([SettingID])
    );
IF OBJECT_ID(N'dbo.Sequences', N'U') IS NULL
    CREATE TABLE [dbo].[Sequences] (
        [SequenceName] NVARCHAR(30) NOT NULL,
        [Prefix] NVARCHAR(10) NULL,
        [NextValue] INT NOT NULL CONSTRAINT [DF_Sequences_NextValue] DEFAULT (1),
        [PadLength] TINYINT NOT NULL CONSTRAINT [DF_Sequences_PadLength] DEFAULT (6),
        [Description] NVARCHAR(100) NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Sequences] PRIMARY KEY ([SequenceName])
    );
IF OBJECT_ID(N'dbo.Roles', N'U') IS NULL
    CREATE TABLE [dbo].[Roles] (
        [RoleID] INT NOT NULL,
        [RoleCode] NVARCHAR(20) NOT NULL,
        [RoleName] NVARCHAR(50) NOT NULL,
        [RoleNameEn] NVARCHAR(50) NULL,
        [Description] NVARCHAR(255) NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Roles] PRIMARY KEY ([RoleID])
    );
IF OBJECT_ID(N'dbo.Permissions', N'U') IS NULL
    CREATE TABLE [dbo].[Permissions] (
        [PermissionKey] NVARCHAR(50) NOT NULL,
        [PermissionName] NVARCHAR(100) NOT NULL,
        [PermissionNameEn] NVARCHAR(100) NULL,
        [ModuleName] NVARCHAR(50) NULL,
        [ModuleNameEn] NVARCHAR(50) NULL,
        [SortOrder] SMALLINT NOT NULL CONSTRAINT [DF_Permissions_SortOrder] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Permissions] PRIMARY KEY ([PermissionKey])
    );
IF OBJECT_ID(N'dbo.RolePermissions', N'U') IS NULL
    CREATE TABLE [dbo].[RolePermissions] (
        [RoleID] INT NOT NULL,
        [PermissionKey] NVARCHAR(50) NOT NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_RolePermissions] PRIMARY KEY ([RoleID], [PermissionKey])
    );
IF OBJECT_ID(N'dbo.Employees', N'U') IS NULL
    CREATE TABLE [dbo].[Employees] (
        [EmployeeID] INT IDENTITY(1,1) NOT NULL,
        [EmployeeName] NVARCHAR(100) NOT NULL,
        [JobTitle] NVARCHAR(50) NULL,
        [Mobile] NVARCHAR(20) NULL,
        [Username] NVARCHAR(30) NOT NULL,
        [PasswordHash] NVARCHAR(64) NULL,
        [PasswordSalt] NVARCHAR(32) NULL,
        [RoleID] INT NOT NULL,
        [MaxDiscountPercent] MONEY NOT NULL CONSTRAINT [DF_Employees_MaxDiscountPercent] DEFAULT (0),
        [MustChangePassword] BIT NOT NULL CONSTRAINT [DF_Employees_MustChangePassword] DEFAULT (1),
        [FailedLoginCount] SMALLINT NOT NULL CONSTRAINT [DF_Employees_FailedLoginCount] DEFAULT (0),
        [LockedUntil] DATETIME NULL,
        [LastLoginAt] DATETIME NULL,
        [IsActive] BIT NOT NULL CONSTRAINT [DF_Employees_IsActive] DEFAULT (1),
        [Notes] NVARCHAR(MAX) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Employees_CreatedAt] DEFAULT (GETDATE()),
        [CashBoxID] INT NULL,
        [IsDeveloper] BIT NOT NULL CONSTRAINT [DF_Employees_IsDeveloper] DEFAULT (0),
        [CustomScreens] BIT NOT NULL CONSTRAINT [DF_Employees_CustomScreens] DEFAULT (0),
        [OnPayroll] BIT NOT NULL CONSTRAINT [DF_Employees_OnPayroll] DEFAULT (0),
        [BasicSalary] MONEY NOT NULL CONSTRAINT [DF_Employees_BasicSalary] DEFAULT (0),
        [HousingAllowance] MONEY NOT NULL CONSTRAINT [DF_Employees_HousingAllowance] DEFAULT (0),
        [TransportAllowance] MONEY NOT NULL CONSTRAINT [DF_Employees_TransportAllowance] DEFAULT (0),
        [OtherAllowance] MONEY NOT NULL CONSTRAINT [DF_Employees_OtherAllowance] DEFAULT (0),
        [IsSaudi] BIT NOT NULL CONSTRAINT [DF_Employees_IsSaudi] DEFAULT (0),
        [AdvanceInstallment] MONEY NOT NULL CONSTRAINT [DF_Employees_AdvanceInstallment] DEFAULT (0),
        [NationalID] NVARCHAR(15) NULL,
        [IBAN] NVARCHAR(34) NULL,
        [HireDate] DATETIME NULL,
        [CostCenterID] INT NULL,
        [WarehouseID] INT NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Employees] PRIMARY KEY ([EmployeeID])
    );
IF OBJECT_ID(N'dbo.Screens', N'U') IS NULL
    CREATE TABLE [dbo].[Screens] (
        [ScreenName] NVARCHAR(64) NOT NULL,
        [ScreenTitle] NVARCHAR(100) NOT NULL,
        [ScreenTitleEn] NVARCHAR(100) NULL,
        [ModuleName] NVARCHAR(50) NULL,
        [ModuleNameEn] NVARCHAR(50) NULL,
        [SortOrder] SMALLINT NOT NULL CONSTRAINT [DF_Screens_SortOrder] DEFAULT (0),
        [PermissionKey] NVARCHAR(50) NULL,
        [HasAdd] BIT NOT NULL CONSTRAINT [DF_Screens_HasAdd] DEFAULT (0),
        [HasEdit] BIT NOT NULL CONSTRAINT [DF_Screens_HasEdit] DEFAULT (0),
        [HasDelete] BIT NOT NULL CONSTRAINT [DF_Screens_HasDelete] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Screens] PRIMARY KEY ([ScreenName])
    );
IF OBJECT_ID(N'dbo.UserScreens', N'U') IS NULL
    CREATE TABLE [dbo].[UserScreens] (
        [EmployeeID] INT NOT NULL,
        [ScreenName] NVARCHAR(64) NOT NULL,
        [CanOpen] BIT NOT NULL CONSTRAINT [DF_UserScreens_CanOpen] DEFAULT (0),
        [CanAdd] BIT NOT NULL CONSTRAINT [DF_UserScreens_CanAdd] DEFAULT (0),
        [CanEdit] BIT NOT NULL CONSTRAINT [DF_UserScreens_CanEdit] DEFAULT (0),
        [CanDelete] BIT NOT NULL CONSTRAINT [DF_UserScreens_CanDelete] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_UserScreens] PRIMARY KEY ([EmployeeID], [ScreenName])
    );
IF OBJECT_ID(N'dbo.Activations', N'U') IS NULL
    CREATE TABLE [dbo].[Activations] (
        [ActivationID] INT IDENTITY(1,1) NOT NULL,
        [MachineID] NVARCHAR(24) NOT NULL,
        [ActivationCode] NVARCHAR(30) NOT NULL,
        [ComputerName] NVARCHAR(64) NULL,
        [ActivatedAt] DATETIME NULL CONSTRAINT [DF_Activations_ActivatedAt] DEFAULT (GETDATE()),
        [EmployeeID] INT NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Activations] PRIMARY KEY ([ActivationID])
    );
IF OBJECT_ID(N'dbo.Categories', N'U') IS NULL
    CREATE TABLE [dbo].[Categories] (
        [CategoryID] INT IDENTITY(1,1) NOT NULL,
        [CategoryName] NVARCHAR(100) NOT NULL,
        [CategoryNameEn] NVARCHAR(100) NULL,
        [Description] NVARCHAR(255) NULL,
        [IsActive] BIT NOT NULL CONSTRAINT [DF_Categories_IsActive] DEFAULT (1),
        [ImagePath] NVARCHAR(255) NULL,
        [TileColor] NVARCHAR(10) NOT NULL CONSTRAINT [DF_Categories_TileColor] DEFAULT (N'BLUE'),
        [SortOrder] SMALLINT NOT NULL CONSTRAINT [DF_Categories_SortOrder] DEFAULT (0),
        [IsAddOn] BIT NOT NULL CONSTRAINT [DF_Categories_IsAddOn] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Categories] PRIMARY KEY ([CategoryID])
    );
IF OBJECT_ID(N'dbo.Units', N'U') IS NULL
    CREATE TABLE [dbo].[Units] (
        [UnitID] INT IDENTITY(1,1) NOT NULL,
        [UnitName] NVARCHAR(30) NOT NULL,
        [UnitNameEn] NVARCHAR(30) NULL,
        [ZatcaUnitCode] NVARCHAR(10) NULL,
        [EtaUnitCode] NVARCHAR(10) NULL,
        [IsActive] BIT NOT NULL CONSTRAINT [DF_Units_IsActive] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Units] PRIMARY KEY ([UnitID])
    );
IF OBJECT_ID(N'dbo.PaymentMethods', N'U') IS NULL
    CREATE TABLE [dbo].[PaymentMethods] (
        [PaymentMethodID] INT NOT NULL,
        [MethodName] NVARCHAR(50) NOT NULL,
        [MethodNameEn] NVARCHAR(50) NULL,
        [ZatcaCode] NVARCHAR(5) NULL,
        [SortOrder] SMALLINT NOT NULL CONSTRAINT [DF_PaymentMethods_SortOrder] DEFAULT (0),
        [IsActive] BIT NOT NULL CONSTRAINT [DF_PaymentMethods_IsActive] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_PaymentMethods] PRIMARY KEY ([PaymentMethodID])
    );
IF OBJECT_ID(N'dbo.Currencies', N'U') IS NULL
    CREATE TABLE [dbo].[Currencies] (
        [CurrencyCode] NVARCHAR(3) NOT NULL,
        [CurrencyName] NVARCHAR(50) NOT NULL,
        [CurrencyNameEn] NVARCHAR(50) NULL,
        [Symbol] NVARCHAR(10) NULL,
        [DecimalPlaces] TINYINT NOT NULL CONSTRAINT [DF_Currencies_DecimalPlaces] DEFAULT (2),
        [SortOrder] SMALLINT NOT NULL CONSTRAINT [DF_Currencies_SortOrder] DEFAULT (0),
        [IsActive] BIT NOT NULL CONSTRAINT [DF_Currencies_IsActive] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Currencies] PRIMARY KEY ([CurrencyCode])
    );
IF OBJECT_ID(N'dbo.CurrencyRates', N'U') IS NULL
    CREATE TABLE [dbo].[CurrencyRates] (
        [CurrencyRateID] INT IDENTITY(1,1) NOT NULL,
        [CurrencyCode] NVARCHAR(3) NOT NULL,
        [RateDate] DATETIME NOT NULL CONSTRAINT [DF_CurrencyRates_RateDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [Rate] MONEY NOT NULL CONSTRAINT [DF_CurrencyRates_Rate] DEFAULT (1),
        [Notes] NVARCHAR(100) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CurrencyRates_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_CurrencyRates] PRIMARY KEY ([CurrencyRateID])
    );
IF OBJECT_ID(N'dbo.CashBoxes', N'U') IS NULL
    CREATE TABLE [dbo].[CashBoxes] (
        [CashBoxID] INT IDENTITY(1,1) NOT NULL,
        [BoxName] NVARCHAR(50) NOT NULL,
        [BoxNameEn] NVARCHAR(50) NULL,
        [BoxType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_CashBoxes_BoxType] DEFAULT (N'CASHIER'),
        [OpeningBalance] MONEY NOT NULL CONSTRAINT [DF_CashBoxes_OpeningBalance] DEFAULT (0),
        [OpeningDate] DATETIME NOT NULL CONSTRAINT [DF_CashBoxes_OpeningDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [IsActive] BIT NOT NULL CONSTRAINT [DF_CashBoxes_IsActive] DEFAULT (1),
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CashBoxes_CreatedAt] DEFAULT (GETDATE()),
        [SalesRepID] INT NULL,
        [ConfirmReceipt] BIT NOT NULL CONSTRAINT [DF_CashBoxes_ConfirmReceipt] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_CashBoxes] PRIMARY KEY ([CashBoxID])
    );
IF OBJECT_ID(N'dbo.CashBoxUsers', N'U') IS NULL
    CREATE TABLE [dbo].[CashBoxUsers] (
        [CashBoxUserID] INT IDENTITY(1,1) NOT NULL,
        [CashBoxID] INT NOT NULL,
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CashBoxUsers_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_CashBoxUsers] PRIMARY KEY ([CashBoxUserID])
    );
IF OBJECT_ID(N'dbo.Warehouses', N'U') IS NULL
    CREATE TABLE [dbo].[Warehouses] (
        [WarehouseID] INT IDENTITY(1,1) NOT NULL,
        [WarehouseCode] NVARCHAR(20) NULL,
        [WarehouseName] NVARCHAR(50) NOT NULL,
        [WarehouseNameEn] NVARCHAR(50) NULL,
        [WarehouseType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_Warehouses_WarehouseType] DEFAULT (N'MAIN'),
        [KeeperID] INT NULL,
        [CostCenterID] INT NULL,
        [SalesRepID] INT NULL,
        [ConfirmReceipt] BIT NOT NULL CONSTRAINT [DF_Warehouses_ConfirmReceipt] DEFAULT (0),
        [IsActive] BIT NOT NULL CONSTRAINT [DF_Warehouses_IsActive] DEFAULT (1),
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Warehouses_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Warehouses] PRIMARY KEY ([WarehouseID])
    );
IF OBJECT_ID(N'dbo.Suppliers', N'U') IS NULL
    CREATE TABLE [dbo].[Suppliers] (
        [SupplierID] INT IDENTITY(1,1) NOT NULL,
        [SupplierName] NVARCHAR(150) NOT NULL,
        [SupplierNameEn] NVARCHAR(150) NULL,
        [ContactPerson] NVARCHAR(100) NULL,
        [Mobile] NVARCHAR(20) NULL,
        [Phone] NVARCHAR(20) NULL,
        [Email] NVARCHAR(100) NULL,
        [VATNumber] NVARCHAR(15) NULL,
        [CRNumber] NVARCHAR(20) NULL,
        [Address] NVARCHAR(255) NULL,
        [City] NVARCHAR(50) NULL,
        [OpeningBalance] MONEY NOT NULL CONSTRAINT [DF_Suppliers_OpeningBalance] DEFAULT (0),
        [CurrentBalance] MONEY NOT NULL CONSTRAINT [DF_Suppliers_CurrentBalance] DEFAULT (0),
        [PaymentTermsDays] SMALLINT NULL CONSTRAINT [DF_Suppliers_PaymentTermsDays] DEFAULT (30),
        [IsActive] BIT NOT NULL CONSTRAINT [DF_Suppliers_IsActive] DEFAULT (1),
        [Notes] NVARCHAR(MAX) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Suppliers_CreatedAt] DEFAULT (GETDATE()),
        [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_Suppliers_CurrencyCode] DEFAULT (N'SAR'),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Suppliers] PRIMARY KEY ([SupplierID])
    );
IF OBJECT_ID(N'dbo.PriceLists', N'U') IS NULL
    CREATE TABLE [dbo].[PriceLists] (
        [PriceListID] INT IDENTITY(1,1) NOT NULL,
        [ListName] NVARCHAR(100) NOT NULL,
        [ListNameEn] NVARCHAR(100) NULL,
        [DiscountPercent] MONEY NOT NULL CONSTRAINT [DF_PriceLists_DiscountPercent] DEFAULT (0),
        [IsActive] BIT NOT NULL CONSTRAINT [DF_PriceLists_IsActive] DEFAULT (1),
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PriceLists_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_PriceLists] PRIMARY KEY ([PriceListID])
    );
IF OBJECT_ID(N'dbo.Customers', N'U') IS NULL
    CREATE TABLE [dbo].[Customers] (
        [CustomerID] INT IDENTITY(1,1) NOT NULL,
        [CustomerName] NVARCHAR(150) NOT NULL,
        [CustomerNameEn] NVARCHAR(150) NULL,
        [Mobile] NVARCHAR(20) NULL,
        [Phone] NVARCHAR(20) NULL,
        [Email] NVARCHAR(100) NULL,
        [VATNumber] NVARCHAR(15) NULL,
        [CRNumber] NVARCHAR(20) NULL,
        [BuildingNo] NVARCHAR(10) NULL,
        [StreetName] NVARCHAR(100) NULL,
        [District] NVARCHAR(100) NULL,
        [City] NVARCHAR(50) NULL,
        [PostalCode] NVARCHAR(10) NULL,
        [Address] NVARCHAR(255) NULL,
        [NationalID] NVARCHAR(20) NULL,
        [OpeningBalance] MONEY NOT NULL CONSTRAINT [DF_Customers_OpeningBalance] DEFAULT (0),
        [CurrentBalance] MONEY NOT NULL CONSTRAINT [DF_Customers_CurrentBalance] DEFAULT (0),
        [AllowCredit] BIT NOT NULL CONSTRAINT [DF_Customers_AllowCredit] DEFAULT (1),
        [CreditLimit] MONEY NOT NULL CONSTRAINT [DF_Customers_CreditLimit] DEFAULT (0),
        [PaymentTermsDays] SMALLINT NULL CONSTRAINT [DF_Customers_PaymentTermsDays] DEFAULT (30),
        [IsSystem] BIT NOT NULL CONSTRAINT [DF_Customers_IsSystem] DEFAULT (0),
        [IsActive] BIT NOT NULL CONSTRAINT [DF_Customers_IsActive] DEFAULT (1),
        [Notes] NVARCHAR(MAX) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Customers_CreatedAt] DEFAULT (GETDATE()),
        [SalesRepID] INT NULL,
        [PriceListID] INT NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Customers] PRIMARY KEY ([CustomerID])
    );
IF OBJECT_ID(N'dbo.Products', N'U') IS NULL
    CREATE TABLE [dbo].[Products] (
        [ProductID] INT IDENTITY(1,1) NOT NULL,
        [ProductCode] NVARCHAR(30) NOT NULL,
        [Barcode] NVARCHAR(50) NULL,
        [ProductName] NVARCHAR(150) NOT NULL,
        [ProductNameEn] NVARCHAR(150) NULL,
        [CategoryID] INT NOT NULL CONSTRAINT [DF_Products_CategoryID] DEFAULT (1),
        [UnitID] INT NOT NULL CONSTRAINT [DF_Products_UnitID] DEFAULT (1),
        [PurchasePrice] MONEY NOT NULL CONSTRAINT [DF_Products_PurchasePrice] DEFAULT (0),
        [AverageCost] MONEY NOT NULL CONSTRAINT [DF_Products_AverageCost] DEFAULT (0),
        [SellingPrice] MONEY NOT NULL CONSTRAINT [DF_Products_SellingPrice] DEFAULT (0),
        [VATCategory] NVARCHAR(1) NOT NULL CONSTRAINT [DF_Products_VATCategory] DEFAULT (N'S'),
        [CurrentQuantity] MONEY NOT NULL CONSTRAINT [DF_Products_CurrentQuantity] DEFAULT (0),
        [MinimumQuantity] MONEY NOT NULL CONSTRAINT [DF_Products_MinimumQuantity] DEFAULT (0),
        [SupplierID] INT NULL,
        [ProductLocation] NVARCHAR(50) NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [IsActive] BIT NOT NULL CONSTRAINT [DF_Products_IsActive] DEFAULT (1),
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Products_CreatedAt] DEFAULT (GETDATE()),
        [UpdatedAt] DATETIME NULL,
        [ImagePath] NVARCHAR(255) NULL,
        [TrackStock] BIT NOT NULL CONSTRAINT [DF_Products_TrackStock] DEFAULT (1),
        [SizePriceM] MONEY NULL,
        [SizePriceL] MONEY NULL,
        [EtaItemType] NVARCHAR(3) NULL,
        [EtaItemCode] NVARCHAR(100) NULL,
        [ReorderQuantity] MONEY NOT NULL CONSTRAINT [DF_Products_ReorderQuantity] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Products] PRIMARY KEY ([ProductID])
    );
IF OBJECT_ID(N'dbo.ProductUnits', N'U') IS NULL
    CREATE TABLE [dbo].[ProductUnits] (
        [ProductUnitID] INT IDENTITY(1,1) NOT NULL,
        [ProductID] INT NOT NULL,
        [UnitID] INT NOT NULL,
        [UnitFactor] MONEY NOT NULL,
        [Barcode] NVARCHAR(50) NULL,
        [SellingPrice] MONEY NOT NULL CONSTRAINT [DF_ProductUnits_SellingPrice] DEFAULT (0),
        [PurchasePrice] MONEY NOT NULL CONSTRAINT [DF_ProductUnits_PurchasePrice] DEFAULT (0),
        [IsActive] BIT NOT NULL CONSTRAINT [DF_ProductUnits_IsActive] DEFAULT (1),
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_ProductUnits_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_ProductUnits] PRIMARY KEY ([ProductUnitID])
    );
IF OBJECT_ID(N'dbo.PriceListItems', N'U') IS NULL
    CREATE TABLE [dbo].[PriceListItems] (
        [PriceListItemID] INT IDENTITY(1,1) NOT NULL,
        [PriceListID] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [UnitID] INT NULL,
        [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_PriceListItems_UnitPrice] DEFAULT (0),
        [ValidFrom] DATETIME NULL,
        [ValidTo] DATETIME NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PriceListItems_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_PriceListItems] PRIMARY KEY ([PriceListItemID])
    );
IF OBJECT_ID(N'dbo.SalesInvoices', N'U') IS NULL
    CREATE TABLE [dbo].[SalesInvoices] (
        [SalesInvoiceID] INT IDENTITY(1,1) NOT NULL,
        [InvoiceNumber] NVARCHAR(20) NOT NULL,
        [InvoiceDate] DATETIME NOT NULL CONSTRAINT [DF_SalesInvoices_InvoiceDate] DEFAULT (GETDATE()),
        [CustomerID] INT NOT NULL CONSTRAINT [DF_SalesInvoices_CustomerID] DEFAULT (1),
        [EmployeeID] INT NOT NULL,
        [PaymentType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_SalesInvoices_PaymentType] DEFAULT (N'CASH'),
        [PaymentMethodID] INT NULL,
        [SubTotal] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_SubTotal] DEFAULT (0),
        [Discount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_Discount] DEFAULT (0),
        [TaxableAmount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_TaxableAmount] DEFAULT (0),
        [Tax] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_Tax] DEFAULT (0),
        [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_TotalAmount] DEFAULT (0),
        [PaidAmount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_PaidAmount] DEFAULT (0),
        [RemainingAmount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_RemainingAmount] DEFAULT (0),
        [AmountTendered] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_AmountTendered] DEFAULT (0),
        [ChangeDue] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_ChangeDue] DEFAULT (0),
        [DueDate] DATETIME NULL,
        [Notes] NVARCHAR(255) NULL,
        [InvoiceSubType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_SalesInvoices_InvoiceSubType] DEFAULT (N'SIMPLIFIED'),
        [InvoiceTypeCode] NVARCHAR(3) NOT NULL CONSTRAINT [DF_SalesInvoices_InvoiceTypeCode] DEFAULT (N'388'),
        [InvoiceUUID] NVARCHAR(36) NULL,
        [ICV] INT NULL,
        [InvoiceHash] NVARCHAR(255) NULL,
        [PreviousInvoiceHash] NVARCHAR(255) NULL,
        [QRCodeData] NVARCHAR(MAX) NULL,
        [ZatcaStatus] NVARCHAR(20) NOT NULL CONSTRAINT [DF_SalesInvoices_ZatcaStatus] DEFAULT (N'NOT_SENT'),
        [ZatcaSubmittedAt] DATETIME NULL,
        [ZatcaResponse] NVARCHAR(MAX) NULL,
        [SignedXmlPath] NVARCHAR(255) NULL,
        [EInvoiceAttempts] SMALLINT NULL CONSTRAINT [DF_SalesInvoices_EInvoiceAttempts] DEFAULT (0),
        [EInvoiceXml] NVARCHAR(MAX) NULL,
        [EInvoiceError] NVARCHAR(255) NULL,
        [EtaUUID] NVARCHAR(64) NULL,
        [EtaPreviousUUID] NVARCHAR(64) NULL,
        [EtaSubmissionId] NVARCHAR(50) NULL,
        [EtaLongId] NVARCHAR(100) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SalesInvoices_CreatedAt] DEFAULT (GETDATE()),
        [OrderType] NVARCHAR(10) NULL,
        [TableNo] NVARCHAR(10) NULL,
        [DeliveryPhone] NVARCHAR(20) NULL,
        [DeliveryAddress] NVARCHAR(255) NULL,
        [OrderName] NVARCHAR(50) NULL,
        [CashBoxID] INT NULL,
        [BankID] INT NULL,
        [CostCenterID] INT NULL,
        [SalesRepID] INT NULL,
        [WarehouseID] INT NOT NULL CONSTRAINT [DF_SalesInvoices_WarehouseID] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_SalesInvoices] PRIMARY KEY ([SalesInvoiceID])
    );
IF OBJECT_ID(N'dbo.SalesInvoiceDetails', N'U') IS NULL
    CREATE TABLE [dbo].[SalesInvoiceDetails] (
        [SalesDetailID] INT IDENTITY(1,1) NOT NULL,
        [SalesInvoiceID] INT NOT NULL,
        [LineNumber] SMALLINT NOT NULL,
        [ProductID] INT NOT NULL,
        [Quantity] MONEY NOT NULL,
        [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_UnitPrice] DEFAULT (0),
        [Discount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_Discount] DEFAULT (0),
        [NetAmount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_NetAmount] DEFAULT (0),
        [VATCategory] NVARCHAR(1) NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_VATCategory] DEFAULT (N'S'),
        [VATRate] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_VATRate] DEFAULT (0.15),
        [Tax] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_Tax] DEFAULT (0),
        [LineTotal] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_LineTotal] DEFAULT (0),
        [UnitCost] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_UnitCost] DEFAULT (0),
        [LineNote] NVARCHAR(100) NULL,
        [UnitID] INT NULL,
        [UnitFactor] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_UnitFactor] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_SalesInvoiceDetails] PRIMARY KEY ([SalesDetailID])
    );
IF OBJECT_ID(N'dbo.SalesReturns', N'U') IS NULL
    CREATE TABLE [dbo].[SalesReturns] (
        [SalesReturnID] INT IDENTITY(1,1) NOT NULL,
        [ReturnNumber] NVARCHAR(20) NOT NULL,
        [ReturnDate] DATETIME NOT NULL CONSTRAINT [DF_SalesReturns_ReturnDate] DEFAULT (GETDATE()),
        [SalesInvoiceID] INT NOT NULL,
        [CustomerID] INT NOT NULL,
        [EmployeeID] INT NOT NULL,
        [Reason] NVARCHAR(255) NOT NULL,
        [RefundType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_SalesReturns_RefundType] DEFAULT (N'CASH'),
        [PaymentMethodID] INT NULL,
        [SubTotal] MONEY NOT NULL CONSTRAINT [DF_SalesReturns_SubTotal] DEFAULT (0),
        [Discount] MONEY NOT NULL CONSTRAINT [DF_SalesReturns_Discount] DEFAULT (0),
        [TaxableAmount] MONEY NOT NULL CONSTRAINT [DF_SalesReturns_TaxableAmount] DEFAULT (0),
        [Tax] MONEY NOT NULL CONSTRAINT [DF_SalesReturns_Tax] DEFAULT (0),
        [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_SalesReturns_TotalAmount] DEFAULT (0),
        [RefundedAmount] MONEY NOT NULL CONSTRAINT [DF_SalesReturns_RefundedAmount] DEFAULT (0),
        [Notes] NVARCHAR(255) NULL,
        [InvoiceSubType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_SalesReturns_InvoiceSubType] DEFAULT (N'SIMPLIFIED'),
        [InvoiceTypeCode] NVARCHAR(3) NOT NULL CONSTRAINT [DF_SalesReturns_InvoiceTypeCode] DEFAULT (N'381'),
        [InvoiceUUID] NVARCHAR(36) NULL,
        [ICV] INT NULL,
        [InvoiceHash] NVARCHAR(255) NULL,
        [PreviousInvoiceHash] NVARCHAR(255) NULL,
        [QRCodeData] NVARCHAR(MAX) NULL,
        [ZatcaStatus] NVARCHAR(20) NOT NULL CONSTRAINT [DF_SalesReturns_ZatcaStatus] DEFAULT (N'NOT_SENT'),
        [ZatcaSubmittedAt] DATETIME NULL,
        [ZatcaResponse] NVARCHAR(MAX) NULL,
        [SignedXmlPath] NVARCHAR(255) NULL,
        [EInvoiceAttempts] SMALLINT NULL CONSTRAINT [DF_SalesReturns_EInvoiceAttempts] DEFAULT (0),
        [EInvoiceXml] NVARCHAR(MAX) NULL,
        [EInvoiceError] NVARCHAR(255) NULL,
        [EtaUUID] NVARCHAR(64) NULL,
        [EtaPreviousUUID] NVARCHAR(64) NULL,
        [EtaSubmissionId] NVARCHAR(50) NULL,
        [EtaLongId] NVARCHAR(100) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SalesReturns_CreatedAt] DEFAULT (GETDATE()),
        [CashBoxID] INT NULL,
        [BankID] INT NULL,
        [CostCenterID] INT NULL,
        [SalesRepID] INT NULL,
        [WarehouseID] INT NOT NULL CONSTRAINT [DF_SalesReturns_WarehouseID] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_SalesReturns] PRIMARY KEY ([SalesReturnID])
    );
IF OBJECT_ID(N'dbo.SalesReturnDetails', N'U') IS NULL
    CREATE TABLE [dbo].[SalesReturnDetails] (
        [ReturnDetailID] INT IDENTITY(1,1) NOT NULL,
        [SalesReturnID] INT NOT NULL,
        [SalesDetailID] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [Quantity] MONEY NOT NULL,
        [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_UnitPrice] DEFAULT (0),
        [Discount] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_Discount] DEFAULT (0),
        [NetAmount] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_NetAmount] DEFAULT (0),
        [VATCategory] NVARCHAR(1) NOT NULL CONSTRAINT [DF_SalesReturnDetails_VATCategory] DEFAULT (N'S'),
        [VATRate] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_VATRate] DEFAULT (0.15),
        [Tax] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_Tax] DEFAULT (0),
        [LineTotal] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_LineTotal] DEFAULT (0),
        [UnitCost] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_UnitCost] DEFAULT (0),
        [ReturnToStock] BIT NOT NULL CONSTRAINT [DF_SalesReturnDetails_ReturnToStock] DEFAULT (1),
        [UnitID] INT NULL,
        [UnitFactor] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_UnitFactor] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_SalesReturnDetails] PRIMARY KEY ([ReturnDetailID])
    );
IF OBJECT_ID(N'dbo.PurchaseInvoices', N'U') IS NULL
    CREATE TABLE [dbo].[PurchaseInvoices] (
        [PurchaseInvoiceID] INT IDENTITY(1,1) NOT NULL,
        [InvoiceNumber] NVARCHAR(20) NOT NULL,
        [SupplierInvoiceNo] NVARCHAR(30) NULL,
        [InvoiceDate] DATETIME NOT NULL CONSTRAINT [DF_PurchaseInvoices_InvoiceDate] DEFAULT (GETDATE()),
        [SupplierID] INT NOT NULL,
        [EmployeeID] INT NOT NULL,
        [PaymentType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_PurchaseInvoices_PaymentType] DEFAULT (N'CASH'),
        [PaymentMethodID] INT NULL,
        [SubTotal] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_SubTotal] DEFAULT (0),
        [Discount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_Discount] DEFAULT (0),
        [TaxableAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_TaxableAmount] DEFAULT (0),
        [Tax] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_Tax] DEFAULT (0),
        [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_TotalAmount] DEFAULT (0),
        [PaidAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_PaidAmount] DEFAULT (0),
        [RemainingAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_RemainingAmount] DEFAULT (0),
        [DueDate] DATETIME NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PurchaseInvoices_CreatedAt] DEFAULT (GETDATE()),
        [CashBoxID] INT NULL,
        [BankID] INT NULL,
        [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_PurchaseInvoices_CurrencyCode] DEFAULT (N'SAR'),
        [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_ExchangeRate] DEFAULT (1),
        [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_ForeignAmount] DEFAULT (0),
        [WarehouseID] INT NOT NULL CONSTRAINT [DF_PurchaseInvoices_WarehouseID] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_PurchaseInvoices] PRIMARY KEY ([PurchaseInvoiceID])
    );
IF OBJECT_ID(N'dbo.PurchaseInvoiceDetails', N'U') IS NULL
    CREATE TABLE [dbo].[PurchaseInvoiceDetails] (
        [PurchaseDetailID] INT IDENTITY(1,1) NOT NULL,
        [PurchaseInvoiceID] INT NOT NULL,
        [LineNumber] SMALLINT NOT NULL,
        [ProductID] INT NOT NULL,
        [Quantity] MONEY NOT NULL,
        [UnitCost] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_UnitCost] DEFAULT (0),
        [Discount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_Discount] DEFAULT (0),
        [NetAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_NetAmount] DEFAULT (0),
        [VATRate] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_VATRate] DEFAULT (0.15),
        [Tax] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_Tax] DEFAULT (0),
        [LineTotal] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_LineTotal] DEFAULT (0),
        [UnitID] INT NULL,
        [UnitFactor] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_UnitFactor] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_PurchaseInvoiceDetails] PRIMARY KEY ([PurchaseDetailID])
    );
IF OBJECT_ID(N'dbo.PurchaseReturns', N'U') IS NULL
    CREATE TABLE [dbo].[PurchaseReturns] (
        [PurchaseReturnID] INT IDENTITY(1,1) NOT NULL,
        [ReturnNumber] NVARCHAR(20) NOT NULL,
        [ReturnDate] DATETIME NOT NULL CONSTRAINT [DF_PurchaseReturns_ReturnDate] DEFAULT (GETDATE()),
        [PurchaseInvoiceID] INT NOT NULL,
        [SupplierID] INT NOT NULL,
        [EmployeeID] INT NOT NULL,
        [Reason] NVARCHAR(255) NOT NULL,
        [RefundType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_PurchaseReturns_RefundType] DEFAULT (N'CREDIT'),
        [PaymentMethodID] INT NULL,
        [SubTotal] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_SubTotal] DEFAULT (0),
        [Discount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_Discount] DEFAULT (0),
        [TaxableAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_TaxableAmount] DEFAULT (0),
        [Tax] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_Tax] DEFAULT (0),
        [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_TotalAmount] DEFAULT (0),
        [RefundedAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_RefundedAmount] DEFAULT (0),
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PurchaseReturns_CreatedAt] DEFAULT (GETDATE()),
        [CashBoxID] INT NULL,
        [BankID] INT NULL,
        [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_PurchaseReturns_CurrencyCode] DEFAULT (N'SAR'),
        [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_ExchangeRate] DEFAULT (1),
        [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_ForeignAmount] DEFAULT (0),
        [WarehouseID] INT NOT NULL CONSTRAINT [DF_PurchaseReturns_WarehouseID] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_PurchaseReturns] PRIMARY KEY ([PurchaseReturnID])
    );
IF OBJECT_ID(N'dbo.PurchaseReturnDetails', N'U') IS NULL
    CREATE TABLE [dbo].[PurchaseReturnDetails] (
        [PurchaseReturnDetailID] INT IDENTITY(1,1) NOT NULL,
        [PurchaseReturnID] INT NOT NULL,
        [PurchaseDetailID] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [Quantity] MONEY NOT NULL,
        [UnitCost] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_UnitCost] DEFAULT (0),
        [Discount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_Discount] DEFAULT (0),
        [NetAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_NetAmount] DEFAULT (0),
        [VATRate] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_VATRate] DEFAULT (0.15),
        [Tax] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_Tax] DEFAULT (0),
        [LineTotal] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_LineTotal] DEFAULT (0),
        [UnitID] INT NULL,
        [UnitFactor] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_UnitFactor] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_PurchaseReturnDetails] PRIMARY KEY ([PurchaseReturnDetailID])
    );
IF OBJECT_ID(N'dbo.CustomerPayments', N'U') IS NULL
    CREATE TABLE [dbo].[CustomerPayments] (
        [PaymentID] INT IDENTITY(1,1) NOT NULL,
        [PaymentNumber] NVARCHAR(20) NOT NULL,
        [CustomerID] INT NOT NULL,
        [PaymentDate] DATETIME NOT NULL CONSTRAINT [DF_CustomerPayments_PaymentDate] DEFAULT (GETDATE()),
        [Amount] MONEY NOT NULL CONSTRAINT [DF_CustomerPayments_Amount] DEFAULT (0),
        [PaymentMethodID] INT NOT NULL CONSTRAINT [DF_CustomerPayments_PaymentMethodID] DEFAULT (1),
        [SalesInvoiceID] INT NULL,
        [EmployeeID] INT NOT NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CustomerPayments_CreatedAt] DEFAULT (GETDATE()),
        [CashBoxID] INT NULL,
        [BankID] INT NULL,
        [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_CustomerPayments_CurrencyCode] DEFAULT (N'SAR'),
        [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_CustomerPayments_ExchangeRate] DEFAULT (1),
        [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_CustomerPayments_ForeignAmount] DEFAULT (0),
        [SalesRepID] INT NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_CustomerPayments] PRIMARY KEY ([PaymentID])
    );
IF OBJECT_ID(N'dbo.SupplierPayments', N'U') IS NULL
    CREATE TABLE [dbo].[SupplierPayments] (
        [PaymentID] INT IDENTITY(1,1) NOT NULL,
        [PaymentNumber] NVARCHAR(20) NOT NULL,
        [SupplierID] INT NOT NULL,
        [PaymentDate] DATETIME NOT NULL CONSTRAINT [DF_SupplierPayments_PaymentDate] DEFAULT (GETDATE()),
        [Amount] MONEY NOT NULL CONSTRAINT [DF_SupplierPayments_Amount] DEFAULT (0),
        [PaymentMethodID] INT NOT NULL CONSTRAINT [DF_SupplierPayments_PaymentMethodID] DEFAULT (1),
        [PurchaseInvoiceID] INT NULL,
        [EmployeeID] INT NOT NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SupplierPayments_CreatedAt] DEFAULT (GETDATE()),
        [CashBoxID] INT NULL,
        [BankID] INT NULL,
        [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_SupplierPayments_CurrencyCode] DEFAULT (N'SAR'),
        [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_SupplierPayments_ExchangeRate] DEFAULT (1),
        [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_SupplierPayments_ForeignAmount] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_SupplierPayments] PRIMARY KEY ([PaymentID])
    );
IF OBJECT_ID(N'dbo.Banks', N'U') IS NULL
    CREATE TABLE [dbo].[Banks] (
        [BankID] INT IDENTITY(1,1) NOT NULL,
        [BankName] NVARCHAR(100) NOT NULL,
        [BankNameEn] NVARCHAR(100) NULL,
        [AccountNo] NVARCHAR(30) NULL,
        [IBAN] NVARCHAR(34) NULL,
        [OpeningBalance] MONEY NOT NULL CONSTRAINT [DF_Banks_OpeningBalance] DEFAULT (0),
        [OpeningDate] DATETIME NOT NULL CONSTRAINT [DF_Banks_OpeningDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [IsActive] BIT NOT NULL CONSTRAINT [DF_Banks_IsActive] DEFAULT (1),
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Banks_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Banks] PRIMARY KEY ([BankID])
    );
IF OBJECT_ID(N'dbo.BankTransactions', N'U') IS NULL
    CREATE TABLE [dbo].[BankTransactions] (
        [BankTxID] INT IDENTITY(1,1) NOT NULL,
        [TxNumber] NVARCHAR(20) NOT NULL,
        [TxDate] DATETIME NOT NULL CONSTRAINT [DF_BankTransactions_TxDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [TxType] NVARCHAR(12) NOT NULL,
        [BankID] INT NOT NULL,
        [ToBankID] INT NULL,
        [CashBoxID] INT NULL,
        [CounterAccount] INT NULL,
        [Amount] MONEY NOT NULL CONSTRAINT [DF_BankTransactions_Amount] DEFAULT (0),
        [FeeAmount] MONEY NOT NULL CONSTRAINT [DF_BankTransactions_FeeAmount] DEFAULT (0),
        [FeeVAT] MONEY NOT NULL CONSTRAINT [DF_BankTransactions_FeeVAT] DEFAULT (0),
        [Reference] NVARCHAR(40) NULL,
        [Description] NVARCHAR(255) NULL,
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_BankTransactions_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_BankTransactions] PRIMARY KEY ([BankTxID])
    );
IF OBJECT_ID(N'dbo.Cheques', N'U') IS NULL
    CREATE TABLE [dbo].[Cheques] (
        [ChequeID] INT IDENTITY(1,1) NOT NULL,
        [ChequeRef] NVARCHAR(20) NOT NULL,
        [Direction] NVARCHAR(3) NOT NULL,
        [CustomerID] INT NULL,
        [SupplierID] INT NULL,
        [ChequeNo] NVARCHAR(30) NOT NULL,
        [DrawerBank] NVARCHAR(100) NULL,
        [BankID] INT NULL,
        [IssueDate] DATETIME NOT NULL CONSTRAINT [DF_Cheques_IssueDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [DueDate] DATETIME NOT NULL CONSTRAINT [DF_Cheques_DueDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [Amount] MONEY NOT NULL CONSTRAINT [DF_Cheques_Amount] DEFAULT (0),
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_Cheques_Status] DEFAULT (N'PENDING'),
        [StatusDate] DATETIME NULL,
        [Notes] NVARCHAR(255) NULL,
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Cheques_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Cheques] PRIMARY KEY ([ChequeID])
    );
IF OBJECT_ID(N'dbo.FixedAssets', N'U') IS NULL
    CREATE TABLE [dbo].[FixedAssets] (
        [AssetID] INT IDENTITY(1,1) NOT NULL,
        [AssetCode] NVARCHAR(20) NOT NULL,
        [AssetName] NVARCHAR(150) NOT NULL,
        [AssetAccount] INT NOT NULL,
        [CostCenterID] INT NULL,
        [PurchaseDate] DATETIME NOT NULL CONSTRAINT [DF_FixedAssets_PurchaseDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [Cost] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_Cost] DEFAULT (0),
        [InputVAT] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_InputVAT] DEFAULT (0),
        [SalvageValue] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_SalvageValue] DEFAULT (0),
        [UsefulLifeMonths] SMALLINT NOT NULL,
        [DepStartDate] DATETIME NOT NULL CONSTRAINT [DF_FixedAssets_DepStartDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [DepMethod] NVARCHAR(2) NOT NULL CONSTRAINT [DF_FixedAssets_DepMethod] DEFAULT (N'SL'),
        [DepRate] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_DepRate] DEFAULT (0),
        [SourceType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_FixedAssets_SourceType] DEFAULT (N'BANK'),
        [BankID] INT NULL,
        [CashBoxID] INT NULL,
        [CounterAccount] INT NULL,
        [OpeningAccumDep] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_OpeningAccumDep] DEFAULT (0),
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_FixedAssets_Status] DEFAULT (N'ACTIVE'),
        [DisposalDate] DATETIME NULL,
        [DisposalProceeds] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_DisposalProceeds] DEFAULT (0),
        [DisposalTo] NVARCHAR(10) NULL,
        [DisposalBankID] INT NULL,
        [DisposalCashBoxID] INT NULL,
        [DisposalAccumDep] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_DisposalAccumDep] DEFAULT (0),
        [DisposalVAT] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_DisposalVAT] DEFAULT (0),
        [Notes] NVARCHAR(255) NULL,
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_FixedAssets_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_FixedAssets] PRIMARY KEY ([AssetID])
    );
IF OBJECT_ID(N'dbo.DepreciationRuns', N'U') IS NULL
    CREATE TABLE [dbo].[DepreciationRuns] (
        [RunID] INT IDENTITY(1,1) NOT NULL,
        [RunNumber] NVARCHAR(20) NOT NULL,
        [RunMonth] DATETIME NOT NULL CONSTRAINT [DF_DepreciationRuns_RunMonth] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_DepreciationRuns_TotalAmount] DEFAULT (0),
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_DepreciationRuns_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_DepreciationRuns] PRIMARY KEY ([RunID])
    );
IF OBJECT_ID(N'dbo.AssetDepreciations', N'U') IS NULL
    CREATE TABLE [dbo].[AssetDepreciations] (
        [DepreciationID] INT IDENTITY(1,1) NOT NULL,
        [RunID] INT NOT NULL,
        [LineNo] SMALLINT NOT NULL,
        [AssetID] INT NOT NULL,
        [Amount] MONEY NOT NULL CONSTRAINT [DF_AssetDepreciations_Amount] DEFAULT (0),
        [DepCenter] INT NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_AssetDepreciations] PRIMARY KEY ([DepreciationID])
    );
IF OBJECT_ID(N'dbo.AssetRevaluations', N'U') IS NULL
    CREATE TABLE [dbo].[AssetRevaluations] (
        [RevaluationID] INT IDENTITY(1,1) NOT NULL,
        [RevalNumber] NVARCHAR(20) NOT NULL,
        [AssetID] INT NOT NULL,
        [RevalDate] DATETIME NOT NULL CONSTRAINT [DF_AssetRevaluations_RevalDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [BookValueBefore] MONEY NOT NULL CONSTRAINT [DF_AssetRevaluations_BookValueBefore] DEFAULT (0),
        [NewValue] MONEY NOT NULL CONSTRAINT [DF_AssetRevaluations_NewValue] DEFAULT (0),
        [ValueChange] MONEY NOT NULL CONSTRAINT [DF_AssetRevaluations_ValueChange] DEFAULT (0),
        [SurplusPart] MONEY NOT NULL CONSTRAINT [DF_AssetRevaluations_SurplusPart] DEFAULT (0),
        [LossPart] MONEY NOT NULL CONSTRAINT [DF_AssetRevaluations_LossPart] DEFAULT (0),
        [MonthsLeft] SMALLINT NOT NULL,
        [RevalCenter] INT NULL,
        [Notes] NVARCHAR(255) NULL,
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_AssetRevaluations_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_AssetRevaluations] PRIMARY KEY ([RevaluationID])
    );
IF OBJECT_ID(N'dbo.AssetTransfers', N'U') IS NULL
    CREATE TABLE [dbo].[AssetTransfers] (
        [TransferID] INT IDENTITY(1,1) NOT NULL,
        [TransferNumber] NVARCHAR(20) NOT NULL,
        [AssetID] INT NOT NULL,
        [TransferDate] DATETIME NOT NULL CONSTRAINT [DF_AssetTransfers_TransferDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [FromCenterID] INT NULL,
        [ToCenterID] INT NULL,
        [AssetValue] MONEY NOT NULL CONSTRAINT [DF_AssetTransfers_AssetValue] DEFAULT (0),
        [AccumDep] MONEY NOT NULL CONSTRAINT [DF_AssetTransfers_AccumDep] DEFAULT (0),
        [Notes] NVARCHAR(255) NULL,
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_AssetTransfers_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_AssetTransfers] PRIMARY KEY ([TransferID])
    );
IF OBJECT_ID(N'dbo.CostCenters', N'U') IS NULL
    CREATE TABLE [dbo].[CostCenters] (
        [CostCenterID] INT IDENTITY(1,1) NOT NULL,
        [CenterCode] NVARCHAR(10) NOT NULL,
        [CenterName] NVARCHAR(100) NOT NULL,
        [CenterNameEn] NVARCHAR(100) NULL,
        [IsDefault] BIT NOT NULL CONSTRAINT [DF_CostCenters_IsDefault] DEFAULT (0),
        [IsActive] BIT NOT NULL CONSTRAINT [DF_CostCenters_IsActive] DEFAULT (1),
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CostCenters_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_CostCenters] PRIMARY KEY ([CostCenterID])
    );
IF OBJECT_ID(N'dbo.SalesReps', N'U') IS NULL
    CREATE TABLE [dbo].[SalesReps] (
        [SalesRepID] INT IDENTITY(1,1) NOT NULL,
        [RepCode] NVARCHAR(20) NOT NULL,
        [RepName] NVARCHAR(100) NOT NULL,
        [RepNameEn] NVARCHAR(100) NULL,
        [Mobile] NVARCHAR(20) NULL,
        [EmployeeID] INT NULL,
        [Region] NVARCHAR(50) NULL,
        [CostCenterID] INT NULL,
        [CommissionRate] MONEY NOT NULL CONSTRAINT [DF_SalesReps_CommissionRate] DEFAULT (0),
        [CommissionBase] NVARCHAR(10) NOT NULL CONSTRAINT [DF_SalesReps_CommissionBase] DEFAULT (N'SALES'),
        [CreditLimit] MONEY NOT NULL CONSTRAINT [DF_SalesReps_CreditLimit] DEFAULT (0),
        [MaxDiscountPercent] MONEY NOT NULL CONSTRAINT [DF_SalesReps_MaxDiscountPercent] DEFAULT (0),
        [IsActive] BIT NOT NULL CONSTRAINT [DF_SalesReps_IsActive] DEFAULT (1),
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SalesReps_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_SalesReps] PRIMARY KEY ([SalesRepID])
    );
IF OBJECT_ID(N'dbo.SalesRepTargets', N'U') IS NULL
    CREATE TABLE [dbo].[SalesRepTargets] (
        [TargetID] INT IDENTITY(1,1) NOT NULL,
        [SalesRepID] INT NOT NULL,
        [TargetYear] SMALLINT NOT NULL,
        [TargetMonth] TINYINT NOT NULL CONSTRAINT [DF_SalesRepTargets_TargetMonth] DEFAULT (1),
        [TargetAmount] MONEY NOT NULL CONSTRAINT [DF_SalesRepTargets_TargetAmount] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_SalesRepTargets] PRIMARY KEY ([TargetID])
    );
IF OBJECT_ID(N'dbo.CommissionRuns', N'U') IS NULL
    CREATE TABLE [dbo].[CommissionRuns] (
        [CommissionRunID] INT IDENTITY(1,1) NOT NULL,
        [RunNumber] NVARCHAR(20) NOT NULL,
        [RunMonth] DATETIME NOT NULL CONSTRAINT [DF_CommissionRuns_RunMonth] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_CommissionRuns_Status] DEFAULT (N'DRAFT'),
        [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_CommissionRuns_TotalAmount] DEFAULT (0),
        [EmployeeID] INT NOT NULL,
        [PostedAt] DATETIME NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CommissionRuns_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_CommissionRuns] PRIMARY KEY ([CommissionRunID])
    );
IF OBJECT_ID(N'dbo.CommissionLines', N'U') IS NULL
    CREATE TABLE [dbo].[CommissionLines] (
        [CommissionLineID] INT IDENTITY(1,1) NOT NULL,
        [CommissionRunID] INT NOT NULL,
        [SalesRepID] INT NOT NULL,
        [RepName] NVARCHAR(100) NULL,
        [CostCenterID] INT NULL,
        [NetSales] MONEY NOT NULL CONSTRAINT [DF_CommissionLines_NetSales] DEFAULT (0),
        [Collections] MONEY NOT NULL CONSTRAINT [DF_CommissionLines_Collections] DEFAULT (0),
        [CommissionBase] NVARCHAR(10) NOT NULL CONSTRAINT [DF_CommissionLines_CommissionBase] DEFAULT (N'SALES'),
        [BaseAmount] MONEY NOT NULL CONSTRAINT [DF_CommissionLines_BaseAmount] DEFAULT (0),
        [CommissionRate] MONEY NOT NULL CONSTRAINT [DF_CommissionLines_CommissionRate] DEFAULT (0),
        [Adjustment] MONEY NOT NULL CONSTRAINT [DF_CommissionLines_Adjustment] DEFAULT (0),
        [Commission] MONEY NOT NULL CONSTRAINT [DF_CommissionLines_Commission] DEFAULT (0),
        [Notes] NVARCHAR(150) NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_CommissionLines] PRIMARY KEY ([CommissionLineID])
    );
IF OBJECT_ID(N'dbo.Budgets', N'U') IS NULL
    CREATE TABLE [dbo].[Budgets] (
        [BudgetID] INT IDENTITY(1,1) NOT NULL,
        [BudgetYear] SMALLINT NOT NULL,
        [BudgetName] NVARCHAR(100) NOT NULL,
        [Notes] NVARCHAR(255) NULL,
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Budgets_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Budgets] PRIMARY KEY ([BudgetID])
    );
IF OBJECT_ID(N'dbo.BudgetLines', N'U') IS NULL
    CREATE TABLE [dbo].[BudgetLines] (
        [BudgetLineID] INT IDENTITY(1,1) NOT NULL,
        [BudgetID] INT NOT NULL,
        [AccountCode] INT NOT NULL,
        [CostCenterID] INT NULL,
        [M1] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M1] DEFAULT (0),
        [M2] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M2] DEFAULT (0),
        [M3] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M3] DEFAULT (0),
        [M4] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M4] DEFAULT (0),
        [M5] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M5] DEFAULT (0),
        [M6] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M6] DEFAULT (0),
        [M7] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M7] DEFAULT (0),
        [M8] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M8] DEFAULT (0),
        [M9] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M9] DEFAULT (0),
        [M10] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M10] DEFAULT (0),
        [M11] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M11] DEFAULT (0),
        [M12] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M12] DEFAULT (0),
        [Notes] NVARCHAR(150) NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_BudgetLines] PRIMARY KEY ([BudgetLineID])
    );
IF OBJECT_ID(N'dbo.PayrollRuns', N'U') IS NULL
    CREATE TABLE [dbo].[PayrollRuns] (
        [PayrollRunID] INT IDENTITY(1,1) NOT NULL,
        [RunNumber] NVARCHAR(20) NOT NULL,
        [PayMonth] DATETIME NOT NULL CONSTRAINT [DF_PayrollRuns_PayMonth] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_PayrollRuns_Status] DEFAULT (N'DRAFT'),
        [PaidDate] DATETIME NULL,
        [PaidFrom] NVARCHAR(10) NULL,
        [BankID] INT NULL,
        [CashBoxID] INT NULL,
        [PaidAmount] MONEY NOT NULL CONSTRAINT [DF_PayrollRuns_PaidAmount] DEFAULT (0),
        [Notes] NVARCHAR(255) NULL,
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PayrollRuns_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_PayrollRuns] PRIMARY KEY ([PayrollRunID])
    );
IF OBJECT_ID(N'dbo.PayrollLines', N'U') IS NULL
    CREATE TABLE [dbo].[PayrollLines] (
        [PayrollLineID] INT IDENTITY(1,1) NOT NULL,
        [PayrollRunID] INT NOT NULL,
        [EmployeeID] INT NOT NULL,
        [EmployeeName] NVARCHAR(100) NULL,
        [CostCenterID] INT NULL,
        [IsSaudi] BIT NOT NULL CONSTRAINT [DF_PayrollLines_IsSaudi] DEFAULT (0),
        [Basic] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_Basic] DEFAULT (0),
        [Housing] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_Housing] DEFAULT (0),
        [OtherAllow] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_OtherAllow] DEFAULT (0),
        [Overtime] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_Overtime] DEFAULT (0),
        [Additions] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_Additions] DEFAULT (0),
        [AbsenceDeduction] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_AbsenceDeduction] DEFAULT (0),
        [AdvanceDeduction] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_AdvanceDeduction] DEFAULT (0),
        [OtherDeduction] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_OtherDeduction] DEFAULT (0),
        [GosiWage] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_GosiWage] DEFAULT (0),
        [GosiEmployee] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_GosiEmployee] DEFAULT (0),
        [GosiEmployer] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_GosiEmployer] DEFAULT (0),
        [NetPay] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_NetPay] DEFAULT (0),
        [Notes] NVARCHAR(150) NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_PayrollLines] PRIMARY KEY ([PayrollLineID])
    );
IF OBJECT_ID(N'dbo.BankReconciliations', N'U') IS NULL
    CREATE TABLE [dbo].[BankReconciliations] (
        [ReconciliationID] INT IDENTITY(1,1) NOT NULL,
        [ReconNumber] NVARCHAR(20) NOT NULL,
        [BankID] INT NOT NULL,
        [StatementDate] DATETIME NOT NULL CONSTRAINT [DF_BankReconciliations_StatementDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [StatementBalance] MONEY NOT NULL CONSTRAINT [DF_BankReconciliations_StatementBalance] DEFAULT (0),
        [BookBalance] MONEY NOT NULL CONSTRAINT [DF_BankReconciliations_BookBalance] DEFAULT (0),
        [Outstanding] MONEY NOT NULL CONSTRAINT [DF_BankReconciliations_Outstanding] DEFAULT (0),
        [Difference] MONEY NOT NULL CONSTRAINT [DF_BankReconciliations_Difference] DEFAULT (0),
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_BankReconciliations_Status] DEFAULT (N'OPEN'),
        [Notes] NVARCHAR(255) NULL,
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_BankReconciliations_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_BankReconciliations] PRIMARY KEY ([ReconciliationID])
    );
IF OBJECT_ID(N'dbo.BankClearings', N'U') IS NULL
    CREATE TABLE [dbo].[BankClearings] (
        [ClearingID] INT IDENTITY(1,1) NOT NULL,
        [ReconciliationID] INT NOT NULL,
        [BankID] INT NOT NULL,
        [SourceType] NVARCHAR(20) NOT NULL,
        [SourceID] INT NOT NULL,
        [ClearedAmount] MONEY NOT NULL CONSTRAINT [DF_BankClearings_ClearedAmount] DEFAULT (0),
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_BankClearings_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_BankClearings] PRIMARY KEY ([ClearingID])
    );
IF OBJECT_ID(N'dbo.CustomerAllocations', N'U') IS NULL
    CREATE TABLE [dbo].[CustomerAllocations] (
        [AllocationID] INT IDENTITY(1,1) NOT NULL,
        [PaymentID] INT NOT NULL,
        [SalesInvoiceID] INT NOT NULL,
        [Amount] MONEY NOT NULL CONSTRAINT [DF_CustomerAllocations_Amount] DEFAULT (0),
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CustomerAllocations_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_CustomerAllocations] PRIMARY KEY ([AllocationID])
    );
IF OBJECT_ID(N'dbo.SupplierAllocations', N'U') IS NULL
    CREATE TABLE [dbo].[SupplierAllocations] (
        [AllocationID] INT IDENTITY(1,1) NOT NULL,
        [PaymentID] INT NOT NULL,
        [PurchaseInvoiceID] INT NOT NULL,
        [Amount] MONEY NOT NULL CONSTRAINT [DF_SupplierAllocations_Amount] DEFAULT (0),
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SupplierAllocations_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_SupplierAllocations] PRIMARY KEY ([AllocationID])
    );
IF OBJECT_ID(N'dbo.ExpenseTypes', N'U') IS NULL
    CREATE TABLE [dbo].[ExpenseTypes] (
        [ExpenseTypeID] INT IDENTITY(1,1) NOT NULL,
        [ExpenseTypeName] NVARCHAR(50) NOT NULL,
        [ExpenseTypeNameEn] NVARCHAR(50) NULL,
        [IsActive] BIT NOT NULL CONSTRAINT [DF_ExpenseTypes_IsActive] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_ExpenseTypes] PRIMARY KEY ([ExpenseTypeID])
    );
IF OBJECT_ID(N'dbo.Expenses', N'U') IS NULL
    CREATE TABLE [dbo].[Expenses] (
        [ExpenseID] INT IDENTITY(1,1) NOT NULL,
        [ExpenseNumber] NVARCHAR(20) NOT NULL,
        [ExpenseDate] DATETIME NOT NULL CONSTRAINT [DF_Expenses_ExpenseDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [ExpenseTypeID] INT NOT NULL,
        [Amount] MONEY NOT NULL CONSTRAINT [DF_Expenses_Amount] DEFAULT (0),
        [Tax] MONEY NOT NULL CONSTRAINT [DF_Expenses_Tax] DEFAULT (0),
        [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_Expenses_TotalAmount] DEFAULT (0),
        [PaymentMethodID] INT NULL,
        [SupplierInvoiceRef] NVARCHAR(30) NULL,
        [Description] NVARCHAR(255) NULL,
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Expenses_CreatedAt] DEFAULT (GETDATE()),
        [CashBoxID] INT NULL,
        [BankID] INT NULL,
        [CostCenterID] INT NULL,
        [RecurringID] INT NULL,
        [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_Expenses_CurrencyCode] DEFAULT (N'SAR'),
        [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_Expenses_ExchangeRate] DEFAULT (1),
        [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_Expenses_ForeignAmount] DEFAULT (0),
        [ForeignTax] MONEY NOT NULL CONSTRAINT [DF_Expenses_ForeignTax] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Expenses] PRIMARY KEY ([ExpenseID])
    );
IF OBJECT_ID(N'dbo.RecurringExpenses', N'U') IS NULL
    CREATE TABLE [dbo].[RecurringExpenses] (
        [RecurringID] INT IDENTITY(1,1) NOT NULL,
        [RecurringName] NVARCHAR(100) NOT NULL,
        [ExpenseTypeID] INT NOT NULL,
        [Amount] MONEY NOT NULL CONSTRAINT [DF_RecurringExpenses_Amount] DEFAULT (0),
        [Tax] MONEY NOT NULL CONSTRAINT [DF_RecurringExpenses_Tax] DEFAULT (0),
        [PaymentMethodID] INT NOT NULL,
        [CashBoxID] INT NULL,
        [BankID] INT NULL,
        [CostCenterID] INT NULL,
        [Frequency] NVARCHAR(10) NOT NULL CONSTRAINT [DF_RecurringExpenses_Frequency] DEFAULT (N'MONTHLY'),
        [DueDay] SMALLINT NOT NULL CONSTRAINT [DF_RecurringExpenses_DueDay] DEFAULT (1),
        [StartDate] DATETIME NOT NULL CONSTRAINT [DF_RecurringExpenses_StartDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [EndDate] DATETIME NULL,
        [NextDueDate] DATETIME NULL,
        [LastCreatedDate] DATETIME NULL,
        [Description] NVARCHAR(255) NULL,
        [IsActive] BIT NOT NULL CONSTRAINT [DF_RecurringExpenses_IsActive] DEFAULT (1),
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_RecurringExpenses_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_RecurringExpenses] PRIMARY KEY ([RecurringID])
    );
IF OBJECT_ID(N'dbo.CashVouchers', N'U') IS NULL
    CREATE TABLE [dbo].[CashVouchers] (
        [CashVoucherID] INT IDENTITY(1,1) NOT NULL,
        [VoucherNumber] NVARCHAR(20) NOT NULL,
        [VoucherDate] DATETIME NOT NULL CONSTRAINT [DF_CashVouchers_VoucherDate] DEFAULT (GETDATE()),
        [VoucherType] NVARCHAR(10) NOT NULL,
        [CashBoxID] INT NOT NULL,
        [ToCashBoxID] INT NULL,
        [Category] NVARCHAR(10) NOT NULL CONSTRAINT [DF_CashVouchers_Category] DEFAULT (N'OTHER'),
        [Amount] MONEY NOT NULL CONSTRAINT [DF_CashVouchers_Amount] DEFAULT (0),
        [PartyName] NVARCHAR(100) NULL,
        [Description] NVARCHAR(255) NULL,
        [ExpenseID] INT NULL,
        [ClosingID] INT NULL,
        [AdvanceEmployeeID] INT NULL,
        [CostCenterID] INT NULL,
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CashVouchers_CreatedAt] DEFAULT (GETDATE()),
        [SalesRepID] INT NULL,
        [InTransit] BIT NOT NULL CONSTRAINT [DF_CashVouchers_InTransit] DEFAULT (0),
        [ReceivedAt] DATETIME NULL,
        [ReceivedByID] INT NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_CashVouchers] PRIMARY KEY ([CashVoucherID])
    );
IF OBJECT_ID(N'dbo.CashClosings', N'U') IS NULL
    CREATE TABLE [dbo].[CashClosings] (
        [ClosingID] INT IDENTITY(1,1) NOT NULL,
        [ClosingNumber] NVARCHAR(20) NOT NULL,
        [ClosingDate] DATETIME NOT NULL CONSTRAINT [DF_CashClosings_ClosingDate] DEFAULT (GETDATE()),
        [CashBoxID] INT NOT NULL,
        [EmployeeID] INT NOT NULL,
        [PeriodStart] DATETIME NULL,
        [OpeningBalance] MONEY NOT NULL CONSTRAINT [DF_CashClosings_OpeningBalance] DEFAULT (0),
        [CashIn] MONEY NOT NULL CONSTRAINT [DF_CashClosings_CashIn] DEFAULT (0),
        [CashOut] MONEY NOT NULL CONSTRAINT [DF_CashClosings_CashOut] DEFAULT (0),
        [ExpectedBalance] MONEY NOT NULL CONSTRAINT [DF_CashClosings_ExpectedBalance] DEFAULT (0),
        [CountedAmount] MONEY NOT NULL CONSTRAINT [DF_CashClosings_CountedAmount] DEFAULT (0),
        [Difference] MONEY NOT NULL CONSTRAINT [DF_CashClosings_Difference] DEFAULT (0),
        [Destination] NVARCHAR(10) NOT NULL CONSTRAINT [DF_CashClosings_Destination] DEFAULT (N'MAIN'),
        [ToCashBoxID] INT NULL,
        [TransferAmount] MONEY NOT NULL CONSTRAINT [DF_CashClosings_TransferAmount] DEFAULT (0),
        [KeptAmount] MONEY NOT NULL CONSTRAINT [DF_CashClosings_KeptAmount] DEFAULT (0),
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CashClosings_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_CashClosings] PRIMARY KEY ([ClosingID])
    );
IF OBJECT_ID(N'dbo.Accounts', N'U') IS NULL
    CREATE TABLE [dbo].[Accounts] (
        [AccountCode] INT NOT NULL,
        [AccountName] NVARCHAR(100) NOT NULL,
        [AccountNameEn] NVARCHAR(100) NULL,
        [AccountType] NVARCHAR(10) NOT NULL,
        [ParentCode] INT NULL,
        [IsActive] BIT NOT NULL CONSTRAINT [DF_Accounts_IsActive] DEFAULT (1),
        [IsPosting] BIT NOT NULL CONSTRAINT [DF_Accounts_IsPosting] DEFAULT (1),
        [IsSystem] BIT NOT NULL CONSTRAINT [DF_Accounts_IsSystem] DEFAULT (0),
        [AccountLevel] TINYINT NOT NULL CONSTRAINT [DF_Accounts_AccountLevel] DEFAULT (1),
        [TreeKey] NVARCHAR(60) NULL,
        [Level1Code] INT NULL,
        [Level2Code] INT NULL,
        [Level3Code] INT NULL,
        [Level4Code] INT NULL,
        [Level5Code] INT NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Accounts] PRIMARY KEY ([AccountCode])
    );
IF OBJECT_ID(N'dbo.JournalSourceTypes', N'U') IS NULL
    CREATE TABLE [dbo].[JournalSourceTypes] (
        [SourceType] NVARCHAR(20) NOT NULL,
        [TypeName] NVARCHAR(50) NOT NULL,
        [TypeNameEn] NVARCHAR(50) NULL,
        [SortOrder] SMALLINT NOT NULL CONSTRAINT [DF_JournalSourceTypes_SortOrder] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_JournalSourceTypes] PRIMARY KEY ([SourceType])
    );
IF OBJECT_ID(N'dbo.JournalEntries', N'U') IS NULL
    CREATE TABLE [dbo].[JournalEntries] (
        [EntryID] INT IDENTITY(1,1) NOT NULL,
        [EntryNumber] NVARCHAR(20) NOT NULL,
        [EntryDate] DATETIME NOT NULL,
        [SourceType] NVARCHAR(20) NOT NULL,
        [SourceID] INT NOT NULL,
        [SourceNumber] NVARCHAR(20) NULL,
        [Description] NVARCHAR(255) NULL,
        [TotalDebit] MONEY NOT NULL CONSTRAINT [DF_JournalEntries_TotalDebit] DEFAULT (0),
        [TotalCredit] MONEY NOT NULL CONSTRAINT [DF_JournalEntries_TotalCredit] DEFAULT (0),
        [LineCount] SMALLINT NOT NULL CONSTRAINT [DF_JournalEntries_LineCount] DEFAULT (0),
        [Signature] MONEY NOT NULL CONSTRAINT [DF_JournalEntries_Signature] DEFAULT (0),
        [UpdatedAt] DATETIME NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_JournalEntries_CreatedAt] DEFAULT (GETDATE()),
        [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_JournalEntries_CurrencyCode] DEFAULT (N'SAR'),
        [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_JournalEntries_ExchangeRate] DEFAULT (1),
        [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_JournalEntries_ForeignAmount] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_JournalEntries] PRIMARY KEY ([EntryID])
    );
IF OBJECT_ID(N'dbo.JournalLines', N'U') IS NULL
    CREATE TABLE [dbo].[JournalLines] (
        [JournalLineID] INT IDENTITY(1,1) NOT NULL,
        [EntryID] INT NOT NULL,
        [LineNumber] SMALLINT NOT NULL,
        [AccountCode] INT NOT NULL,
        [Debit] MONEY NOT NULL CONSTRAINT [DF_JournalLines_Debit] DEFAULT (0),
        [Credit] MONEY NOT NULL CONSTRAINT [DF_JournalLines_Credit] DEFAULT (0),
        [LineText] NVARCHAR(255) NULL,
        [CostCenterID] INT NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_JournalLines] PRIMARY KEY ([JournalLineID])
    );
IF OBJECT_ID(N'dbo.PeriodClosings', N'U') IS NULL
    CREATE TABLE [dbo].[PeriodClosings] (
        [PeriodClosingID] INT IDENTITY(1,1) NOT NULL,
        [ActionType] NVARCHAR(12) NOT NULL,
        [ClosedThrough] DATETIME NULL,
        [PreviousThrough] DATETIME NULL,
        [FiscalYear] SMALLINT NULL,
        [Notes] NVARCHAR(255) NULL,
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PeriodClosings_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_PeriodClosings] PRIMARY KEY ([PeriodClosingID])
    );
IF OBJECT_ID(N'dbo.FiscalPeriods', N'U') IS NULL
    CREATE TABLE [dbo].[FiscalPeriods] (
        [FiscalPeriodID] INT IDENTITY(1,1) NOT NULL,
        [PeriodName] NVARCHAR(50) NOT NULL,
        [StartDate] DATETIME NOT NULL,
        [EndDate] DATETIME NOT NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_FiscalPeriods_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_FiscalPeriods] PRIMARY KEY ([FiscalPeriodID])
    );
IF OBJECT_ID(N'dbo.FiscalYearClosings', N'U') IS NULL
    CREATE TABLE [dbo].[FiscalYearClosings] (
        [YearClosingID] INT IDENTITY(1,1) NOT NULL,
        [FiscalYear] SMALLINT NOT NULL,
        [ClosingNumber] NVARCHAR(20) NOT NULL,
        [ClosingDate] DATETIME NOT NULL CONSTRAINT [DF_FiscalYearClosings_ClosingDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [NetProfit] MONEY NOT NULL CONSTRAINT [DF_FiscalYearClosings_NetProfit] DEFAULT (0),
        [EmployeeID] INT NOT NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_FiscalYearClosings_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_FiscalYearClosings] PRIMARY KEY ([YearClosingID])
    );
IF OBJECT_ID(N'dbo.FiscalYearClosingLines', N'U') IS NULL
    CREATE TABLE [dbo].[FiscalYearClosingLines] (
        [YearClosingLineID] INT IDENTITY(1,1) NOT NULL,
        [YearClosingID] INT NOT NULL,
        [LineNumber] SMALLINT NOT NULL,
        [AccountCode] INT NOT NULL,
        [Debit] MONEY NOT NULL CONSTRAINT [DF_FiscalYearClosingLines_Debit] DEFAULT (0),
        [Credit] MONEY NOT NULL CONSTRAINT [DF_FiscalYearClosingLines_Credit] DEFAULT (0),
        [LineText] NVARCHAR(150) NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_FiscalYearClosingLines] PRIMARY KEY ([YearClosingLineID])
    );
IF OBJECT_ID(N'dbo.VatReturns', N'U') IS NULL
    CREATE TABLE [dbo].[VatReturns] (
        [VatReturnID] INT IDENTITY(1,1) NOT NULL,
        [ReturnNumber] NVARCHAR(20) NOT NULL,
        [PeriodFrom] DATETIME NOT NULL,
        [PeriodTo] DATETIME NOT NULL,
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_VatReturns_Status] DEFAULT (N'DRAFT'),
        [SalesStdAmount] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesStdAmount] DEFAULT (0),
        [SalesStdAdjust] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesStdAdjust] DEFAULT (0),
        [SalesStdVAT] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesStdVAT] DEFAULT (0),
        [SalesZeroAmount] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesZeroAmount] DEFAULT (0),
        [SalesZeroAdjust] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesZeroAdjust] DEFAULT (0),
        [SalesExemptAmount] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesExemptAmount] DEFAULT (0),
        [SalesExemptAdjust] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesExemptAdjust] DEFAULT (0),
        [PurchStdAmount] MONEY NOT NULL CONSTRAINT [DF_VatReturns_PurchStdAmount] DEFAULT (0),
        [PurchStdAdjust] MONEY NOT NULL CONSTRAINT [DF_VatReturns_PurchStdAdjust] DEFAULT (0),
        [PurchStdVAT] MONEY NOT NULL CONSTRAINT [DF_VatReturns_PurchStdVAT] DEFAULT (0),
        [PurchZeroAmount] MONEY NOT NULL CONSTRAINT [DF_VatReturns_PurchZeroAmount] DEFAULT (0),
        [PurchZeroAdjust] MONEY NOT NULL CONSTRAINT [DF_VatReturns_PurchZeroAdjust] DEFAULT (0),
        [Corrections] MONEY NOT NULL CONSTRAINT [DF_VatReturns_Corrections] DEFAULT (0),
        [CarriedCredit] MONEY NOT NULL CONSTRAINT [DF_VatReturns_CarriedCredit] DEFAULT (0),
        [NetDue] MONEY NOT NULL CONSTRAINT [DF_VatReturns_NetDue] DEFAULT (0),
        [FiledDate] DATETIME NULL,
        [FilingRef] NVARCHAR(30) NULL,
        [PaidDate] DATETIME NULL,
        [PaidAmount] MONEY NOT NULL CONSTRAINT [DF_VatReturns_PaidAmount] DEFAULT (0),
        [PaidAccount] INT NULL,
        [Notes] NVARCHAR(255) NULL,
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_VatReturns_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_VatReturns] PRIMARY KEY ([VatReturnID])
    );
IF OBJECT_ID(N'dbo.ManualEntries', N'U') IS NULL
    CREATE TABLE [dbo].[ManualEntries] (
        [ManualEntryID] INT IDENTITY(1,1) NOT NULL,
        [EntryNumber] NVARCHAR(20) NOT NULL,
        [EntryDate] DATETIME NOT NULL CONSTRAINT [DF_ManualEntries_EntryDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [Description] NVARCHAR(255) NOT NULL,
        [Reference] NVARCHAR(50) NULL,
        [ReversalOfID] INT NULL,
        [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_ManualEntries_TotalAmount] DEFAULT (0),
        [EmployeeID] INT NOT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_ManualEntries_CreatedAt] DEFAULT (GETDATE()),
        [UpdatedAt] DATETIME NULL,
        [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_ManualEntries_CurrencyCode] DEFAULT (N'SAR'),
        [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_ManualEntries_ExchangeRate] DEFAULT (1),
        [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_ManualEntries_ForeignAmount] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_ManualEntries] PRIMARY KEY ([ManualEntryID])
    );
IF OBJECT_ID(N'dbo.ManualEntryLines', N'U') IS NULL
    CREATE TABLE [dbo].[ManualEntryLines] (
        [ManualLineID] INT IDENTITY(1,1) NOT NULL,
        [ManualEntryID] INT NOT NULL,
        [LineNumber] SMALLINT NOT NULL,
        [AccountCode] INT NOT NULL,
        [Debit] MONEY NOT NULL CONSTRAINT [DF_ManualEntryLines_Debit] DEFAULT (0),
        [Credit] MONEY NOT NULL CONSTRAINT [DF_ManualEntryLines_Credit] DEFAULT (0),
        [LineText] NVARCHAR(150) NULL,
        [CostCenterID] INT NULL,
        [ForeignDebit] MONEY NOT NULL CONSTRAINT [DF_ManualEntryLines_ForeignDebit] DEFAULT (0),
        [ForeignCredit] MONEY NOT NULL CONSTRAINT [DF_ManualEntryLines_ForeignCredit] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_ManualEntryLines] PRIMARY KEY ([ManualLineID])
    );
IF OBJECT_ID(N'dbo.OpeningBalances', N'U') IS NULL
    CREATE TABLE [dbo].[OpeningBalances] (
        [OpeningLineID] INT IDENTITY(1,1) NOT NULL,
        [AccountCode] INT NOT NULL,
        [Debit] MONEY NOT NULL CONSTRAINT [DF_OpeningBalances_Debit] DEFAULT (0),
        [Credit] MONEY NOT NULL CONSTRAINT [DF_OpeningBalances_Credit] DEFAULT (0),
        [LineText] NVARCHAR(150) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_OpeningBalances_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_OpeningBalances] PRIMARY KEY ([OpeningLineID])
    );
IF OBJECT_ID(N'dbo.OpeningStock', N'U') IS NULL
    CREATE TABLE [dbo].[OpeningStock] (
        [OpeningStockID] INT IDENTITY(1,1) NOT NULL,
        [ProductID] INT NOT NULL,
        [Quantity] MONEY NOT NULL CONSTRAINT [DF_OpeningStock_Quantity] DEFAULT (0),
        [UnitCost] MONEY NOT NULL CONSTRAINT [DF_OpeningStock_UnitCost] DEFAULT (0),
        [TransactionID] INT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_OpeningStock_CreatedAt] DEFAULT (GETDATE()),
        [WarehouseID] INT NOT NULL CONSTRAINT [DF_OpeningStock_WarehouseID] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_OpeningStock] PRIMARY KEY ([OpeningStockID])
    );
IF OBJECT_ID(N'dbo.TransactionTypes', N'U') IS NULL
    CREATE TABLE [dbo].[TransactionTypes] (
        [TransactionTypeID] INT NOT NULL,
        [TypeCode] NVARCHAR(20) NOT NULL,
        [TypeName] NVARCHAR(50) NOT NULL,
        [TypeNameEn] NVARCHAR(50) NULL,
        [Direction] SMALLINT NOT NULL,
        [IsManual] BIT NOT NULL CONSTRAINT [DF_TransactionTypes_IsManual] DEFAULT (0),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_TransactionTypes] PRIMARY KEY ([TransactionTypeID])
    );
IF OBJECT_ID(N'dbo.InventoryTransactions', N'U') IS NULL
    CREATE TABLE [dbo].[InventoryTransactions] (
        [TransactionID] INT IDENTITY(1,1) NOT NULL,
        [TransactionDate] DATETIME NOT NULL CONSTRAINT [DF_InventoryTransactions_TransactionDate] DEFAULT (GETDATE()),
        [ProductID] INT NOT NULL,
        [TransactionTypeID] INT NOT NULL,
        [Quantity] MONEY NOT NULL,
        [UnitCost] MONEY NOT NULL CONSTRAINT [DF_InventoryTransactions_UnitCost] DEFAULT (0),
        [QuantityAfter] MONEY NOT NULL CONSTRAINT [DF_InventoryTransactions_QuantityAfter] DEFAULT (0),
        [ReferenceType] NVARCHAR(20) NULL,
        [ReferenceID] INT NULL,
        [ReferenceNumber] NVARCHAR(20) NULL,
        [EmployeeID] INT NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_InventoryTransactions_CreatedAt] DEFAULT (GETDATE()),
        [WarehouseID] INT NOT NULL CONSTRAINT [DF_InventoryTransactions_WarehouseID] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_InventoryTransactions] PRIMARY KEY ([TransactionID])
    );
IF OBJECT_ID(N'dbo.StockTransfers', N'U') IS NULL
    CREATE TABLE [dbo].[StockTransfers] (
        [StockTransferID] INT IDENTITY(1,1) NOT NULL,
        [TransferNumber] NVARCHAR(20) NOT NULL,
        [TransferDate] DATETIME NOT NULL CONSTRAINT [DF_StockTransfers_TransferDate] DEFAULT (GETDATE()),
        [FromWarehouseID] INT NOT NULL,
        [ToWarehouseID] INT NOT NULL,
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_StockTransfers_Status] DEFAULT (N'DRAFT'),
        [InTransit] BIT NOT NULL CONSTRAINT [DF_StockTransfers_InTransit] DEFAULT (0),
        [ReceivedAt] DATETIME NULL,
        [ReceivedByID] INT NULL,
        [EmployeeID] INT NOT NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockTransfers_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_StockTransfers] PRIMARY KEY ([StockTransferID])
    );
IF OBJECT_ID(N'dbo.StockTransferLines', N'U') IS NULL
    CREATE TABLE [dbo].[StockTransferLines] (
        [StockTransferLineID] INT IDENTITY(1,1) NOT NULL,
        [StockTransferID] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [Quantity] MONEY NOT NULL,
        [UnitCost] MONEY NOT NULL CONSTRAINT [DF_StockTransferLines_UnitCost] DEFAULT (0),
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockTransferLines_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_StockTransferLines] PRIMARY KEY ([StockTransferLineID])
    );
IF OBJECT_ID(N'dbo.StockParties', N'U') IS NULL
    CREATE TABLE [dbo].[StockParties] (
        [StockPartyID] INT IDENTITY(1,1) NOT NULL,
        [PartyCode] NVARCHAR(20) NULL,
        [PartyName] NVARCHAR(50) NOT NULL,
        [PartyNameEn] NVARCHAR(50) NULL,
        [PartyType] NVARCHAR(12) NOT NULL CONSTRAINT [DF_StockParties_PartyType] DEFAULT (N'DEPARTMENT'),
        [AccountCode] INT NULL,
        [CostCenterID] INT NULL,
        [IsActive] BIT NOT NULL CONSTRAINT [DF_StockParties_IsActive] DEFAULT (1),
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockParties_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_StockParties] PRIMARY KEY ([StockPartyID])
    );
IF OBJECT_ID(N'dbo.StockDocuments', N'U') IS NULL
    CREATE TABLE [dbo].[StockDocuments] (
        [StockDocumentID] INT IDENTITY(1,1) NOT NULL,
        [DocNumber] NVARCHAR(20) NOT NULL,
        [DocType] NVARCHAR(10) NOT NULL,
        [DocDate] DATETIME NOT NULL CONSTRAINT [DF_StockDocuments_DocDate] DEFAULT (GETDATE()),
        [WarehouseID] INT NOT NULL CONSTRAINT [DF_StockDocuments_WarehouseID] DEFAULT (1),
        [StockPartyID] INT NOT NULL,
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_StockDocuments_Status] DEFAULT (N'DRAFT'),
        [EmployeeID] INT NOT NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockDocuments_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_StockDocuments] PRIMARY KEY ([StockDocumentID])
    );
IF OBJECT_ID(N'dbo.StockDocumentLines', N'U') IS NULL
    CREATE TABLE [dbo].[StockDocumentLines] (
        [StockDocumentLineID] INT IDENTITY(1,1) NOT NULL,
        [StockDocumentID] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [Quantity] MONEY NOT NULL,
        [UnitCost] MONEY NOT NULL CONSTRAINT [DF_StockDocumentLines_UnitCost] DEFAULT (0),
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockDocumentLines_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_StockDocumentLines] PRIMARY KEY ([StockDocumentLineID])
    );
IF OBJECT_ID(N'dbo.StockNotes', N'U') IS NULL
    CREATE TABLE [dbo].[StockNotes] (
        [StockNoteID] INT IDENTITY(1,1) NOT NULL,
        [NoteNumber] NVARCHAR(20) NOT NULL,
        [NoteType] NVARCHAR(10) NOT NULL,
        [NoteDate] DATETIME NOT NULL CONSTRAINT [DF_StockNotes_NoteDate] DEFAULT (GETDATE()),
        [WarehouseID] INT NOT NULL CONSTRAINT [DF_StockNotes_WarehouseID] DEFAULT (1),
        [CustomerID] INT NULL,
        [SupplierID] INT NULL,
        [SalesOrderID] INT NULL,
        [PurchaseOrderID] INT NULL,
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_StockNotes_Status] DEFAULT (N'DRAFT'),
        [SalesInvoiceID] INT NULL,
        [PurchaseInvoiceID] INT NULL,
        [EmployeeID] INT NOT NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockNotes_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_StockNotes] PRIMARY KEY ([StockNoteID])
    );
IF OBJECT_ID(N'dbo.StockNoteLines', N'U') IS NULL
    CREATE TABLE [dbo].[StockNoteLines] (
        [StockNoteLineID] INT IDENTITY(1,1) NOT NULL,
        [StockNoteID] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [Quantity] MONEY NOT NULL,
        [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_StockNoteLines_UnitPrice] DEFAULT (0),
        [UnitCost] MONEY NOT NULL CONSTRAINT [DF_StockNoteLines_UnitCost] DEFAULT (0),
        [SalesOrderLineID] INT NULL,
        [PurchaseOrderLineID] INT NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockNoteLines_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_StockNoteLines] PRIMARY KEY ([StockNoteLineID])
    );
IF OBJECT_ID(N'dbo.SalesOrders', N'U') IS NULL
    CREATE TABLE [dbo].[SalesOrders] (
        [SalesOrderID] INT IDENTITY(1,1) NOT NULL,
        [OrderNumber] NVARCHAR(20) NOT NULL,
        [OrderDate] DATETIME NOT NULL CONSTRAINT [DF_SalesOrders_OrderDate] DEFAULT (GETDATE()),
        [CustomerID] INT NOT NULL,
        [WarehouseID] INT NOT NULL CONSTRAINT [DF_SalesOrders_WarehouseID] DEFAULT (1),
        [SalesRepID] INT NULL,
        [RequiredDate] DATETIME NULL,
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_SalesOrders_Status] DEFAULT (N'DRAFT'),
        [ApprovedAt] DATETIME NULL,
        [PreparedAt] DATETIME NULL,
        [EmployeeID] INT NOT NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SalesOrders_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_SalesOrders] PRIMARY KEY ([SalesOrderID])
    );
IF OBJECT_ID(N'dbo.Quotations', N'U') IS NULL
    CREATE TABLE [dbo].[Quotations] (
        [QuotationID] INT IDENTITY(1,1) NOT NULL,
        [QuotationNumber] NVARCHAR(20) NOT NULL,
        [QuoteDate] DATETIME NOT NULL CONSTRAINT [DF_Quotations_QuoteDate] DEFAULT (GETDATE()),
        [CustomerID] INT NOT NULL,
        [ValidUntil] DATETIME NOT NULL CONSTRAINT [DF_Quotations_ValidUntil] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_Quotations_Status] DEFAULT (N'DRAFT'),
        [IssuedAt] DATETIME NULL,
        [SubTotal] MONEY NOT NULL CONSTRAINT [DF_Quotations_SubTotal] DEFAULT (0),
        [Tax] MONEY NOT NULL CONSTRAINT [DF_Quotations_Tax] DEFAULT (0),
        [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_Quotations_TotalAmount] DEFAULT (0),
        [SalesOrderID] INT NULL,
        [SalesInvoiceID] INT NULL,
        [EmployeeID] INT NOT NULL,
        [Terms] NVARCHAR(255) NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Quotations_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_Quotations] PRIMARY KEY ([QuotationID])
    );
IF OBJECT_ID(N'dbo.QuotationLines', N'U') IS NULL
    CREATE TABLE [dbo].[QuotationLines] (
        [QuotationLineID] INT IDENTITY(1,1) NOT NULL,
        [QuotationID] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [Quantity] MONEY NOT NULL,
        [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_QuotationLines_UnitPrice] DEFAULT (0),
        [NetAmount] MONEY NOT NULL CONSTRAINT [DF_QuotationLines_NetAmount] DEFAULT (0),
        [Tax] MONEY NOT NULL CONSTRAINT [DF_QuotationLines_Tax] DEFAULT (0),
        [LineTotal] MONEY NOT NULL CONSTRAINT [DF_QuotationLines_LineTotal] DEFAULT (0),
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_QuotationLines_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_QuotationLines] PRIMARY KEY ([QuotationLineID])
    );
IF OBJECT_ID(N'dbo.PurchaseOrders', N'U') IS NULL
    CREATE TABLE [dbo].[PurchaseOrders] (
        [PurchaseOrderID] INT IDENTITY(1,1) NOT NULL,
        [OrderNumber] NVARCHAR(20) NOT NULL,
        [OrderDate] DATETIME NOT NULL CONSTRAINT [DF_PurchaseOrders_OrderDate] DEFAULT (GETDATE()),
        [SupplierID] INT NOT NULL,
        [WarehouseID] INT NOT NULL CONSTRAINT [DF_PurchaseOrders_WarehouseID] DEFAULT (1),
        [ExpectedDate] DATETIME NULL,
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_PurchaseOrders_Status] DEFAULT (N'DRAFT'),
        [ApprovedAt] DATETIME NULL,
        [EmployeeID] INT NOT NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PurchaseOrders_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_PurchaseOrders] PRIMARY KEY ([PurchaseOrderID])
    );
IF OBJECT_ID(N'dbo.PurchaseOrderLines', N'U') IS NULL
    CREATE TABLE [dbo].[PurchaseOrderLines] (
        [PurchaseOrderLineID] INT IDENTITY(1,1) NOT NULL,
        [PurchaseOrderID] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [Quantity] MONEY NOT NULL,
        [UnitCost] MONEY NOT NULL CONSTRAINT [DF_PurchaseOrderLines_UnitCost] DEFAULT (0),
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PurchaseOrderLines_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_PurchaseOrderLines] PRIMARY KEY ([PurchaseOrderLineID])
    );
IF OBJECT_ID(N'dbo.RepDays', N'U') IS NULL
    CREATE TABLE [dbo].[RepDays] (
        [RepDayID] INT IDENTITY(1,1) NOT NULL,
        [DayNumber] NVARCHAR(20) NOT NULL,
        [DayDate] DATETIME NOT NULL CONSTRAINT [DF_RepDays_DayDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [SalesRepID] INT NOT NULL,
        [WarehouseID] INT NOT NULL,
        [CashBoxID] INT NOT NULL,
        [StoreID] INT NOT NULL CONSTRAINT [DF_RepDays_StoreID] DEFAULT (1),
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_RepDays_Status] DEFAULT (N'OPEN'),
        [OpenedAt] DATETIME NOT NULL CONSTRAINT [DF_RepDays_OpenedAt] DEFAULT (GETDATE()),
        [SettledAt] DATETIME NULL,
        [StockCountID] INT NULL,
        [ClosingID] INT NULL,
        [GoodsShortage] MONEY NOT NULL CONSTRAINT [DF_RepDays_GoodsShortage] DEFAULT (0),
        [AdvanceEmployeeID] INT NULL,
        [EmployeeID] INT NOT NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_RepDays_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_RepDays] PRIMARY KEY ([RepDayID])
    );
IF OBJECT_ID(N'dbo.RouteStops', N'U') IS NULL
    CREATE TABLE [dbo].[RouteStops] (
        [RouteStopID] INT IDENTITY(1,1) NOT NULL,
        [SalesRepID] INT NOT NULL,
        [CustomerID] INT NOT NULL,
        [VisitWeekday] TINYINT NOT NULL CONSTRAINT [DF_RouteStops_VisitWeekday] DEFAULT (1),
        [StopOrder] SMALLINT NOT NULL CONSTRAINT [DF_RouteStops_StopOrder] DEFAULT (0),
        [IsActive] BIT NOT NULL CONSTRAINT [DF_RouteStops_IsActive] DEFAULT (1),
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_RouteStops_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_RouteStops] PRIMARY KEY ([RouteStopID])
    );
IF OBJECT_ID(N'dbo.RepVisits', N'U') IS NULL
    CREATE TABLE [dbo].[RepVisits] (
        [RepVisitID] INT IDENTITY(1,1) NOT NULL,
        [VisitDate] DATETIME NOT NULL CONSTRAINT [DF_RepVisits_VisitDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE()))),
        [SalesRepID] INT NOT NULL,
        [CustomerID] INT NOT NULL,
        [IsPlanned] BIT NOT NULL CONSTRAINT [DF_RepVisits_IsPlanned] DEFAULT (0),
        [StopOrder] SMALLINT NOT NULL CONSTRAINT [DF_RepVisits_StopOrder] DEFAULT (0),
        [Result] NVARCHAR(10) NOT NULL CONSTRAINT [DF_RepVisits_Result] DEFAULT (N'PENDING'),
        [VisitedAt] DATETIME NULL,
        [EmployeeID] INT NULL,
        [Notes] NVARCHAR(255) NULL,
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_RepVisits_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_RepVisits] PRIMARY KEY ([RepVisitID])
    );
IF OBJECT_ID(N'dbo.RepPrices', N'U') IS NULL
    CREATE TABLE [dbo].[RepPrices] (
        [RepPriceID] INT IDENTITY(1,1) NOT NULL,
        [SalesRepID] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_RepPrices_UnitPrice] DEFAULT (0),
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_RepPrices_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_RepPrices] PRIMARY KEY ([RepPriceID])
    );
IF OBJECT_ID(N'dbo.SalesOrderLines', N'U') IS NULL
    CREATE TABLE [dbo].[SalesOrderLines] (
        [SalesOrderLineID] INT IDENTITY(1,1) NOT NULL,
        [SalesOrderID] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [Quantity] MONEY NOT NULL,
        [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_SalesOrderLines_UnitPrice] DEFAULT (0),
        [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SalesOrderLines_CreatedAt] DEFAULT (GETDATE()),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_SalesOrderLines] PRIMARY KEY ([SalesOrderLineID])
    );
IF OBJECT_ID(N'dbo.StockCounts', N'U') IS NULL
    CREATE TABLE [dbo].[StockCounts] (
        [StockCountID] INT IDENTITY(1,1) NOT NULL,
        [CountNumber] NVARCHAR(20) NOT NULL,
        [CountDate] DATETIME NOT NULL CONSTRAINT [DF_StockCounts_CountDate] DEFAULT (GETDATE()),
        [CategoryID] INT NULL,
        [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_StockCounts_Status] DEFAULT (N'OPEN'),
        [EmployeeID] INT NOT NULL,
        [PostedAt] DATETIME NULL,
        [PostedByID] INT NULL,
        [Notes] NVARCHAR(255) NULL,
        [WarehouseID] INT NOT NULL CONSTRAINT [DF_StockCounts_WarehouseID] DEFAULT (1),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_StockCounts] PRIMARY KEY ([StockCountID])
    );
IF OBJECT_ID(N'dbo.StockCountDetails', N'U') IS NULL
    CREATE TABLE [dbo].[StockCountDetails] (
        [StockCountDetailID] INT IDENTITY(1,1) NOT NULL,
        [StockCountID] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [SystemQuantity] MONEY NOT NULL CONSTRAINT [DF_StockCountDetails_SystemQuantity] DEFAULT (0),
        [ActualQuantity] MONEY NULL,
        [Difference] MONEY NOT NULL CONSTRAINT [DF_StockCountDetails_Difference] DEFAULT (0),
        [UnitCost] MONEY NOT NULL CONSTRAINT [DF_StockCountDetails_UnitCost] DEFAULT (0),
        [DifferenceValue] MONEY NOT NULL CONSTRAINT [DF_StockCountDetails_DifferenceValue] DEFAULT (0),
        [Notes] NVARCHAR(255) NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_StockCountDetails] PRIMARY KEY ([StockCountDetailID])
    );
IF OBJECT_ID(N'dbo.AuditLog', N'U') IS NULL
    CREATE TABLE [dbo].[AuditLog] (
        [LogID] INT IDENTITY(1,1) NOT NULL,
        [LogDate] DATETIME NOT NULL CONSTRAINT [DF_AuditLog_LogDate] DEFAULT (GETDATE()),
        [EmployeeID] INT NULL,
        [ActionType] NVARCHAR(30) NOT NULL,
        [ObjectName] NVARCHAR(50) NULL,
        [RecordID] NVARCHAR(30) NULL,
        [Details] NVARCHAR(MAX) NULL,
        [ComputerName] NVARCHAR(50) NULL,
        [RecordLabel] NVARCHAR(100) NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_AuditLog] PRIMARY KEY ([LogID])
    );
IF OBJECT_ID(N'dbo.AuditChanges', N'U') IS NULL
    CREATE TABLE [dbo].[AuditChanges] (
        [ChangeID] INT IDENTITY(1,1) NOT NULL,
        [LogID] INT NOT NULL,
        [LineNo] SMALLINT NOT NULL CONSTRAINT [DF_AuditChanges_LineNo] DEFAULT (0),
        [FieldName] NVARCHAR(64) NOT NULL,
        [FieldCaption] NVARCHAR(100) NULL,
        [OldValue] NVARCHAR(255) NULL,
        [NewValue] NVARCHAR(255) NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_AuditChanges] PRIMARY KEY ([ChangeID])
    );
IF OBJECT_ID(N'dbo.LabelSettings', N'U') IS NULL
    CREATE TABLE [dbo].[LabelSettings] (
        [LabelSettingID] INT NOT NULL CONSTRAINT [DF_LabelSettings_LabelSettingID] DEFAULT (1),
        [PrinterName] NVARCHAR(255) NULL,
        [LabelWidth] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_LabelWidth] DEFAULT (38),
        [LabelHeight] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_LabelHeight] DEFAULT (25),
        [LabelsAcross] TINYINT NOT NULL CONSTRAINT [DF_LabelSettings_LabelsAcross] DEFAULT (1),
        [ColumnGap] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_ColumnGap] DEFAULT (2),
        [RowGap] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_RowGap] DEFAULT (0),
        [MarginTop] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_MarginTop] DEFAULT (0),
        [MarginBottom] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_MarginBottom] DEFAULT (0),
        [MarginLeft] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_MarginLeft] DEFAULT (0),
        [MarginRight] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_MarginRight] DEFAULT (0),
        [BarHeight] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_BarHeight] DEFAULT (10),
        [BarWidth] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_BarWidth] DEFAULT (0.25),
        [TopLine1] NVARCHAR(10) NOT NULL CONSTRAINT [DF_LabelSettings_TopLine1] DEFAULT (N'STORE'),
        [TopLine2] NVARCHAR(10) NOT NULL CONSTRAINT [DF_LabelSettings_TopLine2] DEFAULT (N'NAME'),
        [BottomLine1] NVARCHAR(10) NOT NULL CONSTRAINT [DF_LabelSettings_BottomLine1] DEFAULT (N'BARCODE'),
        [BottomLine2] NVARCHAR(10) NOT NULL CONSTRAINT [DF_LabelSettings_BottomLine2] DEFAULT (N'PRICE'),
        [ShortName] NVARCHAR(30) NULL,
        [FontSize] TINYINT NOT NULL CONSTRAINT [DF_LabelSettings_FontSize] DEFAULT (7),
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_LabelSettings] PRIMARY KEY ([LabelSettingID])
    );
IF OBJECT_ID(N'dbo.EInvoiceLog', N'U') IS NULL
    CREATE TABLE [dbo].[EInvoiceLog] (
        [LogID] INT IDENTITY(1,1) NOT NULL,
        [LoggedAt] DATETIME NOT NULL CONSTRAINT [DF_EInvoiceLog_LoggedAt] DEFAULT (GETDATE()),
        [Country] NVARCHAR(2) NOT NULL,
        [DocKind] NVARCHAR(10) NULL,
        [DocID] INT NULL,
        [DocNumber] NVARCHAR(30) NULL,
        [Action] NVARCHAR(20) NOT NULL,
        [Endpoint] NVARCHAR(255) NULL,
        [HttpStatus] SMALLINT NULL CONSTRAINT [DF_EInvoiceLog_HttpStatus] DEFAULT (0),
        [Result] NVARCHAR(12) NOT NULL,
        [Message] NVARCHAR(255) NULL,
        [RequestBody] NVARCHAR(MAX) NULL,
        [ResponseBody] NVARCHAR(MAX) NULL,
        [DurationMs] INT NULL CONSTRAINT [DF_EInvoiceLog_DurationMs] DEFAULT (0),
        [EmployeeID] INT NULL,
        [RowVer] ROWVERSION NOT NULL,
        CONSTRAINT [PK_EInvoiceLog] PRIMARY KEY ([LogID])
    );
GO

/*------------------------------------------------------------------------------
  2. Columns an existing table is missing
------------------------------------------------------------------------------*/
IF COL_LENGTH(N'dbo.Settings', N'SettingID') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [SettingID] INT NOT NULL CONSTRAINT [DF_Settings_SettingID] DEFAULT (1);
IF COL_LENGTH(N'dbo.Settings', N'StoreName') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [StoreName] NVARCHAR(150) NOT NULL;
IF COL_LENGTH(N'dbo.Settings', N'StoreNameEn') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [StoreNameEn] NVARCHAR(150) NULL;
IF COL_LENGTH(N'dbo.Settings', N'VATNumber') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [VATNumber] NVARCHAR(15) NULL;
IF COL_LENGTH(N'dbo.Settings', N'CRNumber') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [CRNumber] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.Settings', N'BuildingNo') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [BuildingNo] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.Settings', N'StreetName') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [StreetName] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Settings', N'District') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [District] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Settings', N'City') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [City] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Settings', N'PostalCode') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [PostalCode] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.Settings', N'AdditionalNo') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [AdditionalNo] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.Settings', N'CountryCode') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [CountryCode] NVARCHAR(2) NOT NULL CONSTRAINT [DF_Settings_CountryCode] DEFAULT (N'SA');
IF COL_LENGTH(N'dbo.Settings', N'Phone') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [Phone] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.Settings', N'Email') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [Email] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Settings', N'VATRate') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [VATRate] MONEY NOT NULL CONSTRAINT [DF_Settings_VATRate] DEFAULT (0.15);
IF COL_LENGTH(N'dbo.Settings', N'PricesIncludeVAT') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [PricesIncludeVAT] BIT NOT NULL CONSTRAINT [DF_Settings_PricesIncludeVAT] DEFAULT (1);
IF COL_LENGTH(N'dbo.Settings', N'AllowNegativeStock') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [AllowNegativeStock] BIT NOT NULL CONSTRAINT [DF_Settings_AllowNegativeStock] DEFAULT (0);
IF COL_LENGTH(N'dbo.Settings', N'CurrencyCode') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [CurrencyCode] NVARCHAR(3) NOT NULL CONSTRAINT [DF_Settings_CurrencyCode] DEFAULT (N'SAR');
IF COL_LENGTH(N'dbo.Settings', N'DefaultCustomerID') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [DefaultCustomerID] INT NOT NULL CONSTRAINT [DF_Settings_DefaultCustomerID] DEFAULT (1);
IF COL_LENGTH(N'dbo.Settings', N'ZatcaPhase') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaPhase] TINYINT NOT NULL CONSTRAINT [DF_Settings_ZatcaPhase] DEFAULT (1);
IF COL_LENGTH(N'dbo.Settings', N'ZatcaEnvironment') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaEnvironment] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.Settings', N'LastInvoiceHash') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [LastInvoiceHash] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'BackupFolder') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [BackupFolder] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'BackupKeepCount') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [BackupKeepCount] SMALLINT NOT NULL CONSTRAINT [DF_Settings_BackupKeepCount] DEFAULT (30);
IF COL_LENGTH(N'dbo.Settings', N'SlowMovingDays') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [SlowMovingDays] SMALLINT NOT NULL CONSTRAINT [DF_Settings_SlowMovingDays] DEFAULT (90);
IF COL_LENGTH(N'dbo.Settings', N'ReceiptFooter') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ReceiptFooter] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'LogoPath') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [LogoPath] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'UpdatedAt') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [UpdatedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.Settings', N'POSMode') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [POSMode] NVARCHAR(10) NOT NULL CONSTRAINT [DF_Settings_POSMode] DEFAULT (N'RETAIL');
IF COL_LENGTH(N'dbo.Settings', N'ImagesFolder') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ImagesFolder] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'InvoicePrintMode') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [InvoicePrintMode] NVARCHAR(10) NOT NULL CONSTRAINT [DF_Settings_InvoicePrintMode] DEFAULT (N'PREVIEW');
IF COL_LENGTH(N'dbo.Settings', N'AllowAdminCompanyName') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [AllowAdminCompanyName] BIT NOT NULL CONSTRAINT [DF_Settings_AllowAdminCompanyName] DEFAULT (0);
IF COL_LENGTH(N'dbo.Settings', N'ClosedThrough') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ClosedThrough] DATETIME NULL;
IF COL_LENGTH(N'dbo.Settings', N'OpeningDate') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [OpeningDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.Settings', N'DefaultBankID') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [DefaultBankID] INT NULL;
IF COL_LENGTH(N'dbo.Settings', N'GosiEmployeeRate') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [GosiEmployeeRate] MONEY NOT NULL CONSTRAINT [DF_Settings_GosiEmployeeRate] DEFAULT (0.0975);
IF COL_LENGTH(N'dbo.Settings', N'GosiEmployerRate') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [GosiEmployerRate] MONEY NOT NULL CONSTRAINT [DF_Settings_GosiEmployerRate] DEFAULT (0.1175);
IF COL_LENGTH(N'dbo.Settings', N'GosiNonSaudiRate') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [GosiNonSaudiRate] MONEY NOT NULL CONSTRAINT [DF_Settings_GosiNonSaudiRate] DEFAULT (0.02);
IF COL_LENGTH(N'dbo.Settings', N'GosiMaxWage') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [GosiMaxWage] MONEY NOT NULL CONSTRAINT [DF_Settings_GosiMaxWage] DEFAULT (45000);
IF COL_LENGTH(N'dbo.Settings', N'CreditBlockDays') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [CreditBlockDays] SMALLINT NULL CONSTRAINT [DF_Settings_CreditBlockDays] DEFAULT (0);
IF COL_LENGTH(N'dbo.Settings', N'EInvoiceEnabled') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EInvoiceEnabled] BIT NOT NULL CONSTRAINT [DF_Settings_EInvoiceEnabled] DEFAULT (0);
IF COL_LENGTH(N'dbo.Settings', N'OpenSslPath') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [OpenSslPath] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'ZatcaKeyFile') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaKeyFile] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'ZatcaCertificate') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaCertificate] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.Settings', N'ZatcaDeviceSerial') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaDeviceSerial] NVARCHAR(36) NULL;
IF COL_LENGTH(N'dbo.Settings', N'ZatcaBranchName') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaBranchName] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Settings', N'ZatcaIndustry') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaIndustry] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Settings', N'ZatcaCsr') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaCsr] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.Settings', N'ZatcaComplianceToken') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaComplianceToken] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.Settings', N'ZatcaComplianceSecret') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaComplianceSecret] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'ZatcaRequestId') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaRequestId] NVARCHAR(30) NULL;
IF COL_LENGTH(N'dbo.Settings', N'ZatcaProductionToken') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaProductionToken] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.Settings', N'ZatcaProductionSecret') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaProductionSecret] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'ZatcaOnboardEnv') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaOnboardEnv] NVARCHAR(12) NULL;
IF COL_LENGTH(N'dbo.Settings', N'ZatcaOnboardStage') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ZatcaOnboardStage] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaClientId') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaClientId] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaClientSecret') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaClientSecret] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaPosSerial') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaPosSerial] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaPosOsVersion') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaPosOsVersion] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaPreSharedKey') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaPreSharedKey] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaBranchCode') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaBranchCode] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaActivityCode') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaActivityCode] NVARCHAR(8) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaGovernate') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaGovernate] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaLastUUID') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaLastUUID] NVARCHAR(64) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaErpClientId') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaErpClientId] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaErpClientSecret') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaErpClientSecret] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaSignerPath') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaSignerPath] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaSignerArgs') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaSignerArgs] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EtaTokenPin') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EtaTokenPin] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Settings', N'EInvoiceEnvironment') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [EInvoiceEnvironment] NVARCHAR(12) NOT NULL CONSTRAINT [DF_Settings_EInvoiceEnvironment] DEFAULT (N'TEST');
IF COL_LENGTH(N'dbo.Settings', N'ServerLiveSince') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [ServerLiveSince] DATETIME NULL;
IF COL_LENGTH(N'dbo.Settings', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Settings] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Sequences', N'SequenceName') IS NULL
    ALTER TABLE [dbo].[Sequences] ADD [SequenceName] NVARCHAR(30) NOT NULL;
IF COL_LENGTH(N'dbo.Sequences', N'Prefix') IS NULL
    ALTER TABLE [dbo].[Sequences] ADD [Prefix] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.Sequences', N'NextValue') IS NULL
    ALTER TABLE [dbo].[Sequences] ADD [NextValue] INT NOT NULL CONSTRAINT [DF_Sequences_NextValue] DEFAULT (1);
IF COL_LENGTH(N'dbo.Sequences', N'PadLength') IS NULL
    ALTER TABLE [dbo].[Sequences] ADD [PadLength] TINYINT NOT NULL CONSTRAINT [DF_Sequences_PadLength] DEFAULT (6);
IF COL_LENGTH(N'dbo.Sequences', N'Description') IS NULL
    ALTER TABLE [dbo].[Sequences] ADD [Description] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Sequences', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Sequences] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Roles', N'RoleID') IS NULL
    ALTER TABLE [dbo].[Roles] ADD [RoleID] INT NOT NULL;
IF COL_LENGTH(N'dbo.Roles', N'RoleCode') IS NULL
    ALTER TABLE [dbo].[Roles] ADD [RoleCode] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.Roles', N'RoleName') IS NULL
    ALTER TABLE [dbo].[Roles] ADD [RoleName] NVARCHAR(50) NOT NULL;
IF COL_LENGTH(N'dbo.Roles', N'RoleNameEn') IS NULL
    ALTER TABLE [dbo].[Roles] ADD [RoleNameEn] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Roles', N'Description') IS NULL
    ALTER TABLE [dbo].[Roles] ADD [Description] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Roles', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Roles] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Permissions', N'PermissionKey') IS NULL
    ALTER TABLE [dbo].[Permissions] ADD [PermissionKey] NVARCHAR(50) NOT NULL;
IF COL_LENGTH(N'dbo.Permissions', N'PermissionName') IS NULL
    ALTER TABLE [dbo].[Permissions] ADD [PermissionName] NVARCHAR(100) NOT NULL;
IF COL_LENGTH(N'dbo.Permissions', N'PermissionNameEn') IS NULL
    ALTER TABLE [dbo].[Permissions] ADD [PermissionNameEn] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Permissions', N'ModuleName') IS NULL
    ALTER TABLE [dbo].[Permissions] ADD [ModuleName] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Permissions', N'ModuleNameEn') IS NULL
    ALTER TABLE [dbo].[Permissions] ADD [ModuleNameEn] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Permissions', N'SortOrder') IS NULL
    ALTER TABLE [dbo].[Permissions] ADD [SortOrder] SMALLINT NOT NULL CONSTRAINT [DF_Permissions_SortOrder] DEFAULT (0);
IF COL_LENGTH(N'dbo.Permissions', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Permissions] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.RolePermissions', N'RoleID') IS NULL
    ALTER TABLE [dbo].[RolePermissions] ADD [RoleID] INT NOT NULL;
IF COL_LENGTH(N'dbo.RolePermissions', N'PermissionKey') IS NULL
    ALTER TABLE [dbo].[RolePermissions] ADD [PermissionKey] NVARCHAR(50) NOT NULL;
IF COL_LENGTH(N'dbo.RolePermissions', N'RowVer') IS NULL
    ALTER TABLE [dbo].[RolePermissions] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Employees', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [EmployeeID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.Employees', N'EmployeeName') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [EmployeeName] NVARCHAR(100) NOT NULL;
IF COL_LENGTH(N'dbo.Employees', N'JobTitle') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [JobTitle] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Employees', N'Mobile') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [Mobile] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.Employees', N'Username') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [Username] NVARCHAR(30) NOT NULL;
IF COL_LENGTH(N'dbo.Employees', N'PasswordHash') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [PasswordHash] NVARCHAR(64) NULL;
IF COL_LENGTH(N'dbo.Employees', N'PasswordSalt') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [PasswordSalt] NVARCHAR(32) NULL;
IF COL_LENGTH(N'dbo.Employees', N'RoleID') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [RoleID] INT NOT NULL;
IF COL_LENGTH(N'dbo.Employees', N'MaxDiscountPercent') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [MaxDiscountPercent] MONEY NOT NULL CONSTRAINT [DF_Employees_MaxDiscountPercent] DEFAULT (0);
IF COL_LENGTH(N'dbo.Employees', N'MustChangePassword') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [MustChangePassword] BIT NOT NULL CONSTRAINT [DF_Employees_MustChangePassword] DEFAULT (1);
IF COL_LENGTH(N'dbo.Employees', N'FailedLoginCount') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [FailedLoginCount] SMALLINT NOT NULL CONSTRAINT [DF_Employees_FailedLoginCount] DEFAULT (0);
IF COL_LENGTH(N'dbo.Employees', N'LockedUntil') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [LockedUntil] DATETIME NULL;
IF COL_LENGTH(N'dbo.Employees', N'LastLoginAt') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [LastLoginAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.Employees', N'IsActive') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_Employees_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.Employees', N'Notes') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [Notes] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.Employees', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Employees_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.Employees', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [CashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.Employees', N'IsDeveloper') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [IsDeveloper] BIT NOT NULL CONSTRAINT [DF_Employees_IsDeveloper] DEFAULT (0);
IF COL_LENGTH(N'dbo.Employees', N'CustomScreens') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [CustomScreens] BIT NOT NULL CONSTRAINT [DF_Employees_CustomScreens] DEFAULT (0);
IF COL_LENGTH(N'dbo.Employees', N'OnPayroll') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [OnPayroll] BIT NOT NULL CONSTRAINT [DF_Employees_OnPayroll] DEFAULT (0);
IF COL_LENGTH(N'dbo.Employees', N'BasicSalary') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [BasicSalary] MONEY NOT NULL CONSTRAINT [DF_Employees_BasicSalary] DEFAULT (0);
IF COL_LENGTH(N'dbo.Employees', N'HousingAllowance') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [HousingAllowance] MONEY NOT NULL CONSTRAINT [DF_Employees_HousingAllowance] DEFAULT (0);
IF COL_LENGTH(N'dbo.Employees', N'TransportAllowance') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [TransportAllowance] MONEY NOT NULL CONSTRAINT [DF_Employees_TransportAllowance] DEFAULT (0);
IF COL_LENGTH(N'dbo.Employees', N'OtherAllowance') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [OtherAllowance] MONEY NOT NULL CONSTRAINT [DF_Employees_OtherAllowance] DEFAULT (0);
IF COL_LENGTH(N'dbo.Employees', N'IsSaudi') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [IsSaudi] BIT NOT NULL CONSTRAINT [DF_Employees_IsSaudi] DEFAULT (0);
IF COL_LENGTH(N'dbo.Employees', N'AdvanceInstallment') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [AdvanceInstallment] MONEY NOT NULL CONSTRAINT [DF_Employees_AdvanceInstallment] DEFAULT (0);
IF COL_LENGTH(N'dbo.Employees', N'NationalID') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [NationalID] NVARCHAR(15) NULL;
IF COL_LENGTH(N'dbo.Employees', N'IBAN') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [IBAN] NVARCHAR(34) NULL;
IF COL_LENGTH(N'dbo.Employees', N'HireDate') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [HireDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.Employees', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.Employees', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [WarehouseID] INT NULL;
IF COL_LENGTH(N'dbo.Employees', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Employees] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Screens', N'ScreenName') IS NULL
    ALTER TABLE [dbo].[Screens] ADD [ScreenName] NVARCHAR(64) NOT NULL;
IF COL_LENGTH(N'dbo.Screens', N'ScreenTitle') IS NULL
    ALTER TABLE [dbo].[Screens] ADD [ScreenTitle] NVARCHAR(100) NOT NULL;
IF COL_LENGTH(N'dbo.Screens', N'ScreenTitleEn') IS NULL
    ALTER TABLE [dbo].[Screens] ADD [ScreenTitleEn] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Screens', N'ModuleName') IS NULL
    ALTER TABLE [dbo].[Screens] ADD [ModuleName] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Screens', N'ModuleNameEn') IS NULL
    ALTER TABLE [dbo].[Screens] ADD [ModuleNameEn] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Screens', N'SortOrder') IS NULL
    ALTER TABLE [dbo].[Screens] ADD [SortOrder] SMALLINT NOT NULL CONSTRAINT [DF_Screens_SortOrder] DEFAULT (0);
IF COL_LENGTH(N'dbo.Screens', N'PermissionKey') IS NULL
    ALTER TABLE [dbo].[Screens] ADD [PermissionKey] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Screens', N'HasAdd') IS NULL
    ALTER TABLE [dbo].[Screens] ADD [HasAdd] BIT NOT NULL CONSTRAINT [DF_Screens_HasAdd] DEFAULT (0);
IF COL_LENGTH(N'dbo.Screens', N'HasEdit') IS NULL
    ALTER TABLE [dbo].[Screens] ADD [HasEdit] BIT NOT NULL CONSTRAINT [DF_Screens_HasEdit] DEFAULT (0);
IF COL_LENGTH(N'dbo.Screens', N'HasDelete') IS NULL
    ALTER TABLE [dbo].[Screens] ADD [HasDelete] BIT NOT NULL CONSTRAINT [DF_Screens_HasDelete] DEFAULT (0);
IF COL_LENGTH(N'dbo.Screens', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Screens] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.UserScreens', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[UserScreens] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.UserScreens', N'ScreenName') IS NULL
    ALTER TABLE [dbo].[UserScreens] ADD [ScreenName] NVARCHAR(64) NOT NULL;
IF COL_LENGTH(N'dbo.UserScreens', N'CanOpen') IS NULL
    ALTER TABLE [dbo].[UserScreens] ADD [CanOpen] BIT NOT NULL CONSTRAINT [DF_UserScreens_CanOpen] DEFAULT (0);
IF COL_LENGTH(N'dbo.UserScreens', N'CanAdd') IS NULL
    ALTER TABLE [dbo].[UserScreens] ADD [CanAdd] BIT NOT NULL CONSTRAINT [DF_UserScreens_CanAdd] DEFAULT (0);
IF COL_LENGTH(N'dbo.UserScreens', N'CanEdit') IS NULL
    ALTER TABLE [dbo].[UserScreens] ADD [CanEdit] BIT NOT NULL CONSTRAINT [DF_UserScreens_CanEdit] DEFAULT (0);
IF COL_LENGTH(N'dbo.UserScreens', N'CanDelete') IS NULL
    ALTER TABLE [dbo].[UserScreens] ADD [CanDelete] BIT NOT NULL CONSTRAINT [DF_UserScreens_CanDelete] DEFAULT (0);
IF COL_LENGTH(N'dbo.UserScreens', N'RowVer') IS NULL
    ALTER TABLE [dbo].[UserScreens] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Activations', N'ActivationID') IS NULL
    ALTER TABLE [dbo].[Activations] ADD [ActivationID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.Activations', N'MachineID') IS NULL
    ALTER TABLE [dbo].[Activations] ADD [MachineID] NVARCHAR(24) NOT NULL;
IF COL_LENGTH(N'dbo.Activations', N'ActivationCode') IS NULL
    ALTER TABLE [dbo].[Activations] ADD [ActivationCode] NVARCHAR(30) NOT NULL;
IF COL_LENGTH(N'dbo.Activations', N'ComputerName') IS NULL
    ALTER TABLE [dbo].[Activations] ADD [ComputerName] NVARCHAR(64) NULL;
IF COL_LENGTH(N'dbo.Activations', N'ActivatedAt') IS NULL
    ALTER TABLE [dbo].[Activations] ADD [ActivatedAt] DATETIME NULL CONSTRAINT [DF_Activations_ActivatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.Activations', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[Activations] ADD [EmployeeID] INT NULL;
IF COL_LENGTH(N'dbo.Activations', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Activations] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Categories', N'CategoryID') IS NULL
    ALTER TABLE [dbo].[Categories] ADD [CategoryID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.Categories', N'CategoryName') IS NULL
    ALTER TABLE [dbo].[Categories] ADD [CategoryName] NVARCHAR(100) NOT NULL;
IF COL_LENGTH(N'dbo.Categories', N'CategoryNameEn') IS NULL
    ALTER TABLE [dbo].[Categories] ADD [CategoryNameEn] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Categories', N'Description') IS NULL
    ALTER TABLE [dbo].[Categories] ADD [Description] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Categories', N'IsActive') IS NULL
    ALTER TABLE [dbo].[Categories] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_Categories_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.Categories', N'ImagePath') IS NULL
    ALTER TABLE [dbo].[Categories] ADD [ImagePath] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Categories', N'TileColor') IS NULL
    ALTER TABLE [dbo].[Categories] ADD [TileColor] NVARCHAR(10) NOT NULL CONSTRAINT [DF_Categories_TileColor] DEFAULT (N'BLUE');
IF COL_LENGTH(N'dbo.Categories', N'SortOrder') IS NULL
    ALTER TABLE [dbo].[Categories] ADD [SortOrder] SMALLINT NOT NULL CONSTRAINT [DF_Categories_SortOrder] DEFAULT (0);
IF COL_LENGTH(N'dbo.Categories', N'IsAddOn') IS NULL
    ALTER TABLE [dbo].[Categories] ADD [IsAddOn] BIT NOT NULL CONSTRAINT [DF_Categories_IsAddOn] DEFAULT (0);
IF COL_LENGTH(N'dbo.Categories', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Categories] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Units', N'UnitID') IS NULL
    ALTER TABLE [dbo].[Units] ADD [UnitID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.Units', N'UnitName') IS NULL
    ALTER TABLE [dbo].[Units] ADD [UnitName] NVARCHAR(30) NOT NULL;
IF COL_LENGTH(N'dbo.Units', N'UnitNameEn') IS NULL
    ALTER TABLE [dbo].[Units] ADD [UnitNameEn] NVARCHAR(30) NULL;
IF COL_LENGTH(N'dbo.Units', N'ZatcaUnitCode') IS NULL
    ALTER TABLE [dbo].[Units] ADD [ZatcaUnitCode] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.Units', N'EtaUnitCode') IS NULL
    ALTER TABLE [dbo].[Units] ADD [EtaUnitCode] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.Units', N'IsActive') IS NULL
    ALTER TABLE [dbo].[Units] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_Units_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.Units', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Units] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.PaymentMethods', N'PaymentMethodID') IS NULL
    ALTER TABLE [dbo].[PaymentMethods] ADD [PaymentMethodID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PaymentMethods', N'MethodName') IS NULL
    ALTER TABLE [dbo].[PaymentMethods] ADD [MethodName] NVARCHAR(50) NOT NULL;
IF COL_LENGTH(N'dbo.PaymentMethods', N'MethodNameEn') IS NULL
    ALTER TABLE [dbo].[PaymentMethods] ADD [MethodNameEn] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.PaymentMethods', N'ZatcaCode') IS NULL
    ALTER TABLE [dbo].[PaymentMethods] ADD [ZatcaCode] NVARCHAR(5) NULL;
IF COL_LENGTH(N'dbo.PaymentMethods', N'SortOrder') IS NULL
    ALTER TABLE [dbo].[PaymentMethods] ADD [SortOrder] SMALLINT NOT NULL CONSTRAINT [DF_PaymentMethods_SortOrder] DEFAULT (0);
IF COL_LENGTH(N'dbo.PaymentMethods', N'IsActive') IS NULL
    ALTER TABLE [dbo].[PaymentMethods] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_PaymentMethods_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.PaymentMethods', N'RowVer') IS NULL
    ALTER TABLE [dbo].[PaymentMethods] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Currencies', N'CurrencyCode') IS NULL
    ALTER TABLE [dbo].[Currencies] ADD [CurrencyCode] NVARCHAR(3) NOT NULL;
IF COL_LENGTH(N'dbo.Currencies', N'CurrencyName') IS NULL
    ALTER TABLE [dbo].[Currencies] ADD [CurrencyName] NVARCHAR(50) NOT NULL;
IF COL_LENGTH(N'dbo.Currencies', N'CurrencyNameEn') IS NULL
    ALTER TABLE [dbo].[Currencies] ADD [CurrencyNameEn] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Currencies', N'Symbol') IS NULL
    ALTER TABLE [dbo].[Currencies] ADD [Symbol] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.Currencies', N'DecimalPlaces') IS NULL
    ALTER TABLE [dbo].[Currencies] ADD [DecimalPlaces] TINYINT NOT NULL CONSTRAINT [DF_Currencies_DecimalPlaces] DEFAULT (2);
IF COL_LENGTH(N'dbo.Currencies', N'SortOrder') IS NULL
    ALTER TABLE [dbo].[Currencies] ADD [SortOrder] SMALLINT NOT NULL CONSTRAINT [DF_Currencies_SortOrder] DEFAULT (0);
IF COL_LENGTH(N'dbo.Currencies', N'IsActive') IS NULL
    ALTER TABLE [dbo].[Currencies] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_Currencies_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.Currencies', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Currencies] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.CurrencyRates', N'CurrencyRateID') IS NULL
    ALTER TABLE [dbo].[CurrencyRates] ADD [CurrencyRateID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.CurrencyRates', N'CurrencyCode') IS NULL
    ALTER TABLE [dbo].[CurrencyRates] ADD [CurrencyCode] NVARCHAR(3) NOT NULL;
IF COL_LENGTH(N'dbo.CurrencyRates', N'RateDate') IS NULL
    ALTER TABLE [dbo].[CurrencyRates] ADD [RateDate] DATETIME NOT NULL CONSTRAINT [DF_CurrencyRates_RateDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.CurrencyRates', N'Rate') IS NULL
    ALTER TABLE [dbo].[CurrencyRates] ADD [Rate] MONEY NOT NULL CONSTRAINT [DF_CurrencyRates_Rate] DEFAULT (1);
IF COL_LENGTH(N'dbo.CurrencyRates', N'Notes') IS NULL
    ALTER TABLE [dbo].[CurrencyRates] ADD [Notes] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.CurrencyRates', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[CurrencyRates] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CurrencyRates_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.CurrencyRates', N'RowVer') IS NULL
    ALTER TABLE [dbo].[CurrencyRates] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.CashBoxes', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[CashBoxes] ADD [CashBoxID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.CashBoxes', N'BoxName') IS NULL
    ALTER TABLE [dbo].[CashBoxes] ADD [BoxName] NVARCHAR(50) NOT NULL;
IF COL_LENGTH(N'dbo.CashBoxes', N'BoxNameEn') IS NULL
    ALTER TABLE [dbo].[CashBoxes] ADD [BoxNameEn] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.CashBoxes', N'BoxType') IS NULL
    ALTER TABLE [dbo].[CashBoxes] ADD [BoxType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_CashBoxes_BoxType] DEFAULT (N'CASHIER');
IF COL_LENGTH(N'dbo.CashBoxes', N'OpeningBalance') IS NULL
    ALTER TABLE [dbo].[CashBoxes] ADD [OpeningBalance] MONEY NOT NULL CONSTRAINT [DF_CashBoxes_OpeningBalance] DEFAULT (0);
IF COL_LENGTH(N'dbo.CashBoxes', N'OpeningDate') IS NULL
    ALTER TABLE [dbo].[CashBoxes] ADD [OpeningDate] DATETIME NOT NULL CONSTRAINT [DF_CashBoxes_OpeningDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.CashBoxes', N'IsActive') IS NULL
    ALTER TABLE [dbo].[CashBoxes] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_CashBoxes_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.CashBoxes', N'Notes') IS NULL
    ALTER TABLE [dbo].[CashBoxes] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.CashBoxes', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[CashBoxes] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CashBoxes_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.CashBoxes', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[CashBoxes] ADD [SalesRepID] INT NULL;
IF COL_LENGTH(N'dbo.CashBoxes', N'ConfirmReceipt') IS NULL
    ALTER TABLE [dbo].[CashBoxes] ADD [ConfirmReceipt] BIT NOT NULL CONSTRAINT [DF_CashBoxes_ConfirmReceipt] DEFAULT (0);
IF COL_LENGTH(N'dbo.CashBoxes', N'RowVer') IS NULL
    ALTER TABLE [dbo].[CashBoxes] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.CashBoxUsers', N'CashBoxUserID') IS NULL
    ALTER TABLE [dbo].[CashBoxUsers] ADD [CashBoxUserID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.CashBoxUsers', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[CashBoxUsers] ADD [CashBoxID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CashBoxUsers', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[CashBoxUsers] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CashBoxUsers', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[CashBoxUsers] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CashBoxUsers_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.CashBoxUsers', N'RowVer') IS NULL
    ALTER TABLE [dbo].[CashBoxUsers] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Warehouses', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[Warehouses] ADD [WarehouseID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.Warehouses', N'WarehouseCode') IS NULL
    ALTER TABLE [dbo].[Warehouses] ADD [WarehouseCode] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.Warehouses', N'WarehouseName') IS NULL
    ALTER TABLE [dbo].[Warehouses] ADD [WarehouseName] NVARCHAR(50) NOT NULL;
IF COL_LENGTH(N'dbo.Warehouses', N'WarehouseNameEn') IS NULL
    ALTER TABLE [dbo].[Warehouses] ADD [WarehouseNameEn] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Warehouses', N'WarehouseType') IS NULL
    ALTER TABLE [dbo].[Warehouses] ADD [WarehouseType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_Warehouses_WarehouseType] DEFAULT (N'MAIN');
IF COL_LENGTH(N'dbo.Warehouses', N'KeeperID') IS NULL
    ALTER TABLE [dbo].[Warehouses] ADD [KeeperID] INT NULL;
IF COL_LENGTH(N'dbo.Warehouses', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[Warehouses] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.Warehouses', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[Warehouses] ADD [SalesRepID] INT NULL;
IF COL_LENGTH(N'dbo.Warehouses', N'ConfirmReceipt') IS NULL
    ALTER TABLE [dbo].[Warehouses] ADD [ConfirmReceipt] BIT NOT NULL CONSTRAINT [DF_Warehouses_ConfirmReceipt] DEFAULT (0);
IF COL_LENGTH(N'dbo.Warehouses', N'IsActive') IS NULL
    ALTER TABLE [dbo].[Warehouses] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_Warehouses_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.Warehouses', N'Notes') IS NULL
    ALTER TABLE [dbo].[Warehouses] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Warehouses', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[Warehouses] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Warehouses_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.Warehouses', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Warehouses] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Suppliers', N'SupplierID') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [SupplierID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.Suppliers', N'SupplierName') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [SupplierName] NVARCHAR(150) NOT NULL;
IF COL_LENGTH(N'dbo.Suppliers', N'SupplierNameEn') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [SupplierNameEn] NVARCHAR(150) NULL;
IF COL_LENGTH(N'dbo.Suppliers', N'ContactPerson') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [ContactPerson] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Suppliers', N'Mobile') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [Mobile] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.Suppliers', N'Phone') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [Phone] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.Suppliers', N'Email') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [Email] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Suppliers', N'VATNumber') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [VATNumber] NVARCHAR(15) NULL;
IF COL_LENGTH(N'dbo.Suppliers', N'CRNumber') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [CRNumber] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.Suppliers', N'Address') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [Address] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Suppliers', N'City') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [City] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Suppliers', N'OpeningBalance') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [OpeningBalance] MONEY NOT NULL CONSTRAINT [DF_Suppliers_OpeningBalance] DEFAULT (0);
IF COL_LENGTH(N'dbo.Suppliers', N'CurrentBalance') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [CurrentBalance] MONEY NOT NULL CONSTRAINT [DF_Suppliers_CurrentBalance] DEFAULT (0);
IF COL_LENGTH(N'dbo.Suppliers', N'PaymentTermsDays') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [PaymentTermsDays] SMALLINT NULL CONSTRAINT [DF_Suppliers_PaymentTermsDays] DEFAULT (30);
IF COL_LENGTH(N'dbo.Suppliers', N'IsActive') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_Suppliers_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.Suppliers', N'Notes') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [Notes] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.Suppliers', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Suppliers_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.Suppliers', N'CurrencyCode') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_Suppliers_CurrencyCode] DEFAULT (N'SAR');
IF COL_LENGTH(N'dbo.Suppliers', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Suppliers] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.PriceLists', N'PriceListID') IS NULL
    ALTER TABLE [dbo].[PriceLists] ADD [PriceListID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.PriceLists', N'ListName') IS NULL
    ALTER TABLE [dbo].[PriceLists] ADD [ListName] NVARCHAR(100) NOT NULL;
IF COL_LENGTH(N'dbo.PriceLists', N'ListNameEn') IS NULL
    ALTER TABLE [dbo].[PriceLists] ADD [ListNameEn] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.PriceLists', N'DiscountPercent') IS NULL
    ALTER TABLE [dbo].[PriceLists] ADD [DiscountPercent] MONEY NOT NULL CONSTRAINT [DF_PriceLists_DiscountPercent] DEFAULT (0);
IF COL_LENGTH(N'dbo.PriceLists', N'IsActive') IS NULL
    ALTER TABLE [dbo].[PriceLists] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_PriceLists_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.PriceLists', N'Notes') IS NULL
    ALTER TABLE [dbo].[PriceLists] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.PriceLists', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[PriceLists] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PriceLists_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.PriceLists', N'RowVer') IS NULL
    ALTER TABLE [dbo].[PriceLists] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Customers', N'CustomerID') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [CustomerID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.Customers', N'CustomerName') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [CustomerName] NVARCHAR(150) NOT NULL;
IF COL_LENGTH(N'dbo.Customers', N'CustomerNameEn') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [CustomerNameEn] NVARCHAR(150) NULL;
IF COL_LENGTH(N'dbo.Customers', N'Mobile') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [Mobile] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.Customers', N'Phone') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [Phone] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.Customers', N'Email') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [Email] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Customers', N'VATNumber') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [VATNumber] NVARCHAR(15) NULL;
IF COL_LENGTH(N'dbo.Customers', N'CRNumber') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [CRNumber] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.Customers', N'BuildingNo') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [BuildingNo] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.Customers', N'StreetName') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [StreetName] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Customers', N'District') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [District] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Customers', N'City') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [City] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Customers', N'PostalCode') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [PostalCode] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.Customers', N'Address') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [Address] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Customers', N'NationalID') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [NationalID] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.Customers', N'OpeningBalance') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [OpeningBalance] MONEY NOT NULL CONSTRAINT [DF_Customers_OpeningBalance] DEFAULT (0);
IF COL_LENGTH(N'dbo.Customers', N'CurrentBalance') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [CurrentBalance] MONEY NOT NULL CONSTRAINT [DF_Customers_CurrentBalance] DEFAULT (0);
IF COL_LENGTH(N'dbo.Customers', N'AllowCredit') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [AllowCredit] BIT NOT NULL CONSTRAINT [DF_Customers_AllowCredit] DEFAULT (1);
IF COL_LENGTH(N'dbo.Customers', N'CreditLimit') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [CreditLimit] MONEY NOT NULL CONSTRAINT [DF_Customers_CreditLimit] DEFAULT (0);
IF COL_LENGTH(N'dbo.Customers', N'PaymentTermsDays') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [PaymentTermsDays] SMALLINT NULL CONSTRAINT [DF_Customers_PaymentTermsDays] DEFAULT (30);
IF COL_LENGTH(N'dbo.Customers', N'IsSystem') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [IsSystem] BIT NOT NULL CONSTRAINT [DF_Customers_IsSystem] DEFAULT (0);
IF COL_LENGTH(N'dbo.Customers', N'IsActive') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_Customers_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.Customers', N'Notes') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [Notes] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.Customers', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Customers_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.Customers', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [SalesRepID] INT NULL;
IF COL_LENGTH(N'dbo.Customers', N'PriceListID') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [PriceListID] INT NULL;
IF COL_LENGTH(N'dbo.Customers', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Customers] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Products', N'ProductID') IS NULL
    ALTER TABLE [dbo].[Products] ADD [ProductID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.Products', N'ProductCode') IS NULL
    ALTER TABLE [dbo].[Products] ADD [ProductCode] NVARCHAR(30) NOT NULL;
IF COL_LENGTH(N'dbo.Products', N'Barcode') IS NULL
    ALTER TABLE [dbo].[Products] ADD [Barcode] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Products', N'ProductName') IS NULL
    ALTER TABLE [dbo].[Products] ADD [ProductName] NVARCHAR(150) NOT NULL;
IF COL_LENGTH(N'dbo.Products', N'ProductNameEn') IS NULL
    ALTER TABLE [dbo].[Products] ADD [ProductNameEn] NVARCHAR(150) NULL;
IF COL_LENGTH(N'dbo.Products', N'CategoryID') IS NULL
    ALTER TABLE [dbo].[Products] ADD [CategoryID] INT NOT NULL CONSTRAINT [DF_Products_CategoryID] DEFAULT (1);
IF COL_LENGTH(N'dbo.Products', N'UnitID') IS NULL
    ALTER TABLE [dbo].[Products] ADD [UnitID] INT NOT NULL CONSTRAINT [DF_Products_UnitID] DEFAULT (1);
IF COL_LENGTH(N'dbo.Products', N'PurchasePrice') IS NULL
    ALTER TABLE [dbo].[Products] ADD [PurchasePrice] MONEY NOT NULL CONSTRAINT [DF_Products_PurchasePrice] DEFAULT (0);
IF COL_LENGTH(N'dbo.Products', N'AverageCost') IS NULL
    ALTER TABLE [dbo].[Products] ADD [AverageCost] MONEY NOT NULL CONSTRAINT [DF_Products_AverageCost] DEFAULT (0);
IF COL_LENGTH(N'dbo.Products', N'SellingPrice') IS NULL
    ALTER TABLE [dbo].[Products] ADD [SellingPrice] MONEY NOT NULL CONSTRAINT [DF_Products_SellingPrice] DEFAULT (0);
IF COL_LENGTH(N'dbo.Products', N'VATCategory') IS NULL
    ALTER TABLE [dbo].[Products] ADD [VATCategory] NVARCHAR(1) NOT NULL CONSTRAINT [DF_Products_VATCategory] DEFAULT (N'S');
IF COL_LENGTH(N'dbo.Products', N'CurrentQuantity') IS NULL
    ALTER TABLE [dbo].[Products] ADD [CurrentQuantity] MONEY NOT NULL CONSTRAINT [DF_Products_CurrentQuantity] DEFAULT (0);
IF COL_LENGTH(N'dbo.Products', N'MinimumQuantity') IS NULL
    ALTER TABLE [dbo].[Products] ADD [MinimumQuantity] MONEY NOT NULL CONSTRAINT [DF_Products_MinimumQuantity] DEFAULT (0);
IF COL_LENGTH(N'dbo.Products', N'SupplierID') IS NULL
    ALTER TABLE [dbo].[Products] ADD [SupplierID] INT NULL;
IF COL_LENGTH(N'dbo.Products', N'ProductLocation') IS NULL
    ALTER TABLE [dbo].[Products] ADD [ProductLocation] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.Products', N'Notes') IS NULL
    ALTER TABLE [dbo].[Products] ADD [Notes] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.Products', N'IsActive') IS NULL
    ALTER TABLE [dbo].[Products] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_Products_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.Products', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[Products] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Products_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.Products', N'UpdatedAt') IS NULL
    ALTER TABLE [dbo].[Products] ADD [UpdatedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.Products', N'ImagePath') IS NULL
    ALTER TABLE [dbo].[Products] ADD [ImagePath] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Products', N'TrackStock') IS NULL
    ALTER TABLE [dbo].[Products] ADD [TrackStock] BIT NOT NULL CONSTRAINT [DF_Products_TrackStock] DEFAULT (1);
IF COL_LENGTH(N'dbo.Products', N'SizePriceM') IS NULL
    ALTER TABLE [dbo].[Products] ADD [SizePriceM] MONEY NULL;
IF COL_LENGTH(N'dbo.Products', N'SizePriceL') IS NULL
    ALTER TABLE [dbo].[Products] ADD [SizePriceL] MONEY NULL;
IF COL_LENGTH(N'dbo.Products', N'EtaItemType') IS NULL
    ALTER TABLE [dbo].[Products] ADD [EtaItemType] NVARCHAR(3) NULL;
IF COL_LENGTH(N'dbo.Products', N'EtaItemCode') IS NULL
    ALTER TABLE [dbo].[Products] ADD [EtaItemCode] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Products', N'ReorderQuantity') IS NULL
    ALTER TABLE [dbo].[Products] ADD [ReorderQuantity] MONEY NOT NULL CONSTRAINT [DF_Products_ReorderQuantity] DEFAULT (0);
IF COL_LENGTH(N'dbo.Products', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Products] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.ProductUnits', N'ProductUnitID') IS NULL
    ALTER TABLE [dbo].[ProductUnits] ADD [ProductUnitID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.ProductUnits', N'ProductID') IS NULL
    ALTER TABLE [dbo].[ProductUnits] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.ProductUnits', N'UnitID') IS NULL
    ALTER TABLE [dbo].[ProductUnits] ADD [UnitID] INT NOT NULL;
IF COL_LENGTH(N'dbo.ProductUnits', N'UnitFactor') IS NULL
    ALTER TABLE [dbo].[ProductUnits] ADD [UnitFactor] MONEY NOT NULL;
IF COL_LENGTH(N'dbo.ProductUnits', N'Barcode') IS NULL
    ALTER TABLE [dbo].[ProductUnits] ADD [Barcode] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.ProductUnits', N'SellingPrice') IS NULL
    ALTER TABLE [dbo].[ProductUnits] ADD [SellingPrice] MONEY NOT NULL CONSTRAINT [DF_ProductUnits_SellingPrice] DEFAULT (0);
IF COL_LENGTH(N'dbo.ProductUnits', N'PurchasePrice') IS NULL
    ALTER TABLE [dbo].[ProductUnits] ADD [PurchasePrice] MONEY NOT NULL CONSTRAINT [DF_ProductUnits_PurchasePrice] DEFAULT (0);
IF COL_LENGTH(N'dbo.ProductUnits', N'IsActive') IS NULL
    ALTER TABLE [dbo].[ProductUnits] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_ProductUnits_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.ProductUnits', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[ProductUnits] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_ProductUnits_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.ProductUnits', N'RowVer') IS NULL
    ALTER TABLE [dbo].[ProductUnits] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.PriceListItems', N'PriceListItemID') IS NULL
    ALTER TABLE [dbo].[PriceListItems] ADD [PriceListItemID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.PriceListItems', N'PriceListID') IS NULL
    ALTER TABLE [dbo].[PriceListItems] ADD [PriceListID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PriceListItems', N'ProductID') IS NULL
    ALTER TABLE [dbo].[PriceListItems] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PriceListItems', N'UnitID') IS NULL
    ALTER TABLE [dbo].[PriceListItems] ADD [UnitID] INT NULL;
IF COL_LENGTH(N'dbo.PriceListItems', N'UnitPrice') IS NULL
    ALTER TABLE [dbo].[PriceListItems] ADD [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_PriceListItems_UnitPrice] DEFAULT (0);
IF COL_LENGTH(N'dbo.PriceListItems', N'ValidFrom') IS NULL
    ALTER TABLE [dbo].[PriceListItems] ADD [ValidFrom] DATETIME NULL;
IF COL_LENGTH(N'dbo.PriceListItems', N'ValidTo') IS NULL
    ALTER TABLE [dbo].[PriceListItems] ADD [ValidTo] DATETIME NULL;
IF COL_LENGTH(N'dbo.PriceListItems', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[PriceListItems] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PriceListItems_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.PriceListItems', N'RowVer') IS NULL
    ALTER TABLE [dbo].[PriceListItems] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'SalesInvoiceID') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [SalesInvoiceID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'InvoiceNumber') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [InvoiceNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'InvoiceDate') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [InvoiceDate] DATETIME NOT NULL CONSTRAINT [DF_SalesInvoices_InvoiceDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.SalesInvoices', N'CustomerID') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [CustomerID] INT NOT NULL CONSTRAINT [DF_SalesInvoices_CustomerID] DEFAULT (1);
IF COL_LENGTH(N'dbo.SalesInvoices', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'PaymentType') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [PaymentType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_SalesInvoices_PaymentType] DEFAULT (N'CASH');
IF COL_LENGTH(N'dbo.SalesInvoices', N'PaymentMethodID') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [PaymentMethodID] INT NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'SubTotal') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [SubTotal] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_SubTotal] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoices', N'Discount') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [Discount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_Discount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoices', N'TaxableAmount') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [TaxableAmount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_TaxableAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoices', N'Tax') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [Tax] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_Tax] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoices', N'TotalAmount') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_TotalAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoices', N'PaidAmount') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [PaidAmount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_PaidAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoices', N'RemainingAmount') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [RemainingAmount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_RemainingAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoices', N'AmountTendered') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [AmountTendered] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_AmountTendered] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoices', N'ChangeDue') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [ChangeDue] MONEY NOT NULL CONSTRAINT [DF_SalesInvoices_ChangeDue] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoices', N'DueDate') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [DueDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'Notes') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'InvoiceSubType') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [InvoiceSubType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_SalesInvoices_InvoiceSubType] DEFAULT (N'SIMPLIFIED');
IF COL_LENGTH(N'dbo.SalesInvoices', N'InvoiceTypeCode') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [InvoiceTypeCode] NVARCHAR(3) NOT NULL CONSTRAINT [DF_SalesInvoices_InvoiceTypeCode] DEFAULT (N'388');
IF COL_LENGTH(N'dbo.SalesInvoices', N'InvoiceUUID') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [InvoiceUUID] NVARCHAR(36) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'ICV') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [ICV] INT NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'InvoiceHash') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [InvoiceHash] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'PreviousInvoiceHash') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [PreviousInvoiceHash] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'QRCodeData') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [QRCodeData] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'ZatcaStatus') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [ZatcaStatus] NVARCHAR(20) NOT NULL CONSTRAINT [DF_SalesInvoices_ZatcaStatus] DEFAULT (N'NOT_SENT');
IF COL_LENGTH(N'dbo.SalesInvoices', N'ZatcaSubmittedAt') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [ZatcaSubmittedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'ZatcaResponse') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [ZatcaResponse] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'SignedXmlPath') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [SignedXmlPath] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'EInvoiceAttempts') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [EInvoiceAttempts] SMALLINT NULL CONSTRAINT [DF_SalesInvoices_EInvoiceAttempts] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoices', N'EInvoiceXml') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [EInvoiceXml] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'EInvoiceError') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [EInvoiceError] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'EtaUUID') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [EtaUUID] NVARCHAR(64) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'EtaPreviousUUID') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [EtaPreviousUUID] NVARCHAR(64) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'EtaSubmissionId') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [EtaSubmissionId] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'EtaLongId') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [EtaLongId] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SalesInvoices_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.SalesInvoices', N'OrderType') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [OrderType] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'TableNo') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [TableNo] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'DeliveryPhone') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [DeliveryPhone] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'DeliveryAddress') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [DeliveryAddress] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'OrderName') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [OrderName] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [CashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'BankID') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [BankID] INT NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [SalesRepID] INT NULL;
IF COL_LENGTH(N'dbo.SalesInvoices', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [WarehouseID] INT NOT NULL CONSTRAINT [DF_SalesInvoices_WarehouseID] DEFAULT (1);
IF COL_LENGTH(N'dbo.SalesInvoices', N'RowVer') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'SalesDetailID') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [SalesDetailID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'SalesInvoiceID') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [SalesInvoiceID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'LineNumber') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [LineNumber] SMALLINT NOT NULL;
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'ProductID') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'Quantity') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [Quantity] MONEY NOT NULL;
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'UnitPrice') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_UnitPrice] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'Discount') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [Discount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_Discount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'NetAmount') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [NetAmount] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_NetAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'VATCategory') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [VATCategory] NVARCHAR(1) NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_VATCategory] DEFAULT (N'S');
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'VATRate') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [VATRate] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_VATRate] DEFAULT (0.15);
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'Tax') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [Tax] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_Tax] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'LineTotal') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [LineTotal] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_LineTotal] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'UnitCost') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [UnitCost] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_UnitCost] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'LineNote') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [LineNote] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'UnitID') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [UnitID] INT NULL;
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'UnitFactor') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [UnitFactor] MONEY NOT NULL CONSTRAINT [DF_SalesInvoiceDetails_UnitFactor] DEFAULT (1);
IF COL_LENGTH(N'dbo.SalesInvoiceDetails', N'RowVer') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'SalesReturnID') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [SalesReturnID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'ReturnNumber') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [ReturnNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'ReturnDate') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [ReturnDate] DATETIME NOT NULL CONSTRAINT [DF_SalesReturns_ReturnDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.SalesReturns', N'SalesInvoiceID') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [SalesInvoiceID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'CustomerID') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [CustomerID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'Reason') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [Reason] NVARCHAR(255) NOT NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'RefundType') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [RefundType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_SalesReturns_RefundType] DEFAULT (N'CASH');
IF COL_LENGTH(N'dbo.SalesReturns', N'PaymentMethodID') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [PaymentMethodID] INT NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'SubTotal') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [SubTotal] MONEY NOT NULL CONSTRAINT [DF_SalesReturns_SubTotal] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReturns', N'Discount') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [Discount] MONEY NOT NULL CONSTRAINT [DF_SalesReturns_Discount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReturns', N'TaxableAmount') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [TaxableAmount] MONEY NOT NULL CONSTRAINT [DF_SalesReturns_TaxableAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReturns', N'Tax') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [Tax] MONEY NOT NULL CONSTRAINT [DF_SalesReturns_Tax] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReturns', N'TotalAmount') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_SalesReturns_TotalAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReturns', N'RefundedAmount') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [RefundedAmount] MONEY NOT NULL CONSTRAINT [DF_SalesReturns_RefundedAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReturns', N'Notes') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'InvoiceSubType') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [InvoiceSubType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_SalesReturns_InvoiceSubType] DEFAULT (N'SIMPLIFIED');
IF COL_LENGTH(N'dbo.SalesReturns', N'InvoiceTypeCode') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [InvoiceTypeCode] NVARCHAR(3) NOT NULL CONSTRAINT [DF_SalesReturns_InvoiceTypeCode] DEFAULT (N'381');
IF COL_LENGTH(N'dbo.SalesReturns', N'InvoiceUUID') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [InvoiceUUID] NVARCHAR(36) NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'ICV') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [ICV] INT NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'InvoiceHash') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [InvoiceHash] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'PreviousInvoiceHash') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [PreviousInvoiceHash] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'QRCodeData') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [QRCodeData] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'ZatcaStatus') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [ZatcaStatus] NVARCHAR(20) NOT NULL CONSTRAINT [DF_SalesReturns_ZatcaStatus] DEFAULT (N'NOT_SENT');
IF COL_LENGTH(N'dbo.SalesReturns', N'ZatcaSubmittedAt') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [ZatcaSubmittedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'ZatcaResponse') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [ZatcaResponse] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'SignedXmlPath') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [SignedXmlPath] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'EInvoiceAttempts') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [EInvoiceAttempts] SMALLINT NULL CONSTRAINT [DF_SalesReturns_EInvoiceAttempts] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReturns', N'EInvoiceXml') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [EInvoiceXml] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'EInvoiceError') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [EInvoiceError] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'EtaUUID') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [EtaUUID] NVARCHAR(64) NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'EtaPreviousUUID') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [EtaPreviousUUID] NVARCHAR(64) NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'EtaSubmissionId') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [EtaSubmissionId] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'EtaLongId') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [EtaLongId] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SalesReturns_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.SalesReturns', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [CashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'BankID') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [BankID] INT NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [SalesRepID] INT NULL;
IF COL_LENGTH(N'dbo.SalesReturns', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [WarehouseID] INT NOT NULL CONSTRAINT [DF_SalesReturns_WarehouseID] DEFAULT (1);
IF COL_LENGTH(N'dbo.SalesReturns', N'RowVer') IS NULL
    ALTER TABLE [dbo].[SalesReturns] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'ReturnDetailID') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [ReturnDetailID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'SalesReturnID') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [SalesReturnID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'SalesDetailID') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [SalesDetailID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'ProductID') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'Quantity') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [Quantity] MONEY NOT NULL;
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'UnitPrice') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_UnitPrice] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'Discount') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [Discount] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_Discount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'NetAmount') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [NetAmount] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_NetAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'VATCategory') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [VATCategory] NVARCHAR(1) NOT NULL CONSTRAINT [DF_SalesReturnDetails_VATCategory] DEFAULT (N'S');
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'VATRate') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [VATRate] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_VATRate] DEFAULT (0.15);
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'Tax') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [Tax] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_Tax] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'LineTotal') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [LineTotal] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_LineTotal] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'UnitCost') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [UnitCost] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_UnitCost] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'ReturnToStock') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [ReturnToStock] BIT NOT NULL CONSTRAINT [DF_SalesReturnDetails_ReturnToStock] DEFAULT (1);
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'UnitID') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [UnitID] INT NULL;
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'UnitFactor') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [UnitFactor] MONEY NOT NULL CONSTRAINT [DF_SalesReturnDetails_UnitFactor] DEFAULT (1);
IF COL_LENGTH(N'dbo.SalesReturnDetails', N'RowVer') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'PurchaseInvoiceID') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [PurchaseInvoiceID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'InvoiceNumber') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [InvoiceNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'SupplierInvoiceNo') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [SupplierInvoiceNo] NVARCHAR(30) NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'InvoiceDate') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [InvoiceDate] DATETIME NOT NULL CONSTRAINT [DF_PurchaseInvoices_InvoiceDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'SupplierID') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [SupplierID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'PaymentType') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [PaymentType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_PurchaseInvoices_PaymentType] DEFAULT (N'CASH');
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'PaymentMethodID') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [PaymentMethodID] INT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'SubTotal') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [SubTotal] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_SubTotal] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'Discount') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [Discount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_Discount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'TaxableAmount') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [TaxableAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_TaxableAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'Tax') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [Tax] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_Tax] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'TotalAmount') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_TotalAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'PaidAmount') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [PaidAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_PaidAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'RemainingAmount') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [RemainingAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_RemainingAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'DueDate') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [DueDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'Notes') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PurchaseInvoices_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [CashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'BankID') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [BankID] INT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'CurrencyCode') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_PurchaseInvoices_CurrencyCode] DEFAULT (N'SAR');
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'ExchangeRate') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_ExchangeRate] DEFAULT (1);
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'ForeignAmount') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoices_ForeignAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [WarehouseID] INT NOT NULL CONSTRAINT [DF_PurchaseInvoices_WarehouseID] DEFAULT (1);
IF COL_LENGTH(N'dbo.PurchaseInvoices', N'RowVer') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'PurchaseDetailID') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [PurchaseDetailID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'PurchaseInvoiceID') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [PurchaseInvoiceID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'LineNumber') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [LineNumber] SMALLINT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'ProductID') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'Quantity') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [Quantity] MONEY NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'UnitCost') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [UnitCost] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_UnitCost] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'Discount') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [Discount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_Discount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'NetAmount') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [NetAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_NetAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'VATRate') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [VATRate] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_VATRate] DEFAULT (0.15);
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'Tax') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [Tax] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_Tax] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'LineTotal') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [LineTotal] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_LineTotal] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'UnitID') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [UnitID] INT NULL;
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'UnitFactor') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [UnitFactor] MONEY NOT NULL CONSTRAINT [DF_PurchaseInvoiceDetails_UnitFactor] DEFAULT (1);
IF COL_LENGTH(N'dbo.PurchaseInvoiceDetails', N'RowVer') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturns', N'PurchaseReturnID') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [PurchaseReturnID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturns', N'ReturnNumber') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [ReturnNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturns', N'ReturnDate') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [ReturnDate] DATETIME NOT NULL CONSTRAINT [DF_PurchaseReturns_ReturnDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.PurchaseReturns', N'PurchaseInvoiceID') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [PurchaseInvoiceID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturns', N'SupplierID') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [SupplierID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturns', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturns', N'Reason') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [Reason] NVARCHAR(255) NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturns', N'RefundType') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [RefundType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_PurchaseReturns_RefundType] DEFAULT (N'CREDIT');
IF COL_LENGTH(N'dbo.PurchaseReturns', N'PaymentMethodID') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [PaymentMethodID] INT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturns', N'SubTotal') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [SubTotal] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_SubTotal] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseReturns', N'Discount') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [Discount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_Discount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseReturns', N'TaxableAmount') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [TaxableAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_TaxableAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseReturns', N'Tax') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [Tax] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_Tax] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseReturns', N'TotalAmount') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_TotalAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseReturns', N'RefundedAmount') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [RefundedAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_RefundedAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseReturns', N'Notes') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.PurchaseReturns', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PurchaseReturns_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.PurchaseReturns', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [CashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturns', N'BankID') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [BankID] INT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturns', N'CurrencyCode') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_PurchaseReturns_CurrencyCode] DEFAULT (N'SAR');
IF COL_LENGTH(N'dbo.PurchaseReturns', N'ExchangeRate') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_ExchangeRate] DEFAULT (1);
IF COL_LENGTH(N'dbo.PurchaseReturns', N'ForeignAmount') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturns_ForeignAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseReturns', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [WarehouseID] INT NOT NULL CONSTRAINT [DF_PurchaseReturns_WarehouseID] DEFAULT (1);
IF COL_LENGTH(N'dbo.PurchaseReturns', N'RowVer') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'PurchaseReturnDetailID') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [PurchaseReturnDetailID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'PurchaseReturnID') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [PurchaseReturnID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'PurchaseDetailID') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [PurchaseDetailID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'ProductID') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'Quantity') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [Quantity] MONEY NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'UnitCost') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [UnitCost] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_UnitCost] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'Discount') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [Discount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_Discount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'NetAmount') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [NetAmount] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_NetAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'VATRate') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [VATRate] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_VATRate] DEFAULT (0.15);
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'Tax') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [Tax] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_Tax] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'LineTotal') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [LineTotal] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_LineTotal] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'UnitID') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [UnitID] INT NULL;
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'UnitFactor') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [UnitFactor] MONEY NOT NULL CONSTRAINT [DF_PurchaseReturnDetails_UnitFactor] DEFAULT (1);
IF COL_LENGTH(N'dbo.PurchaseReturnDetails', N'RowVer') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.CustomerPayments', N'PaymentID') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [PaymentID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.CustomerPayments', N'PaymentNumber') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [PaymentNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.CustomerPayments', N'CustomerID') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [CustomerID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CustomerPayments', N'PaymentDate') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [PaymentDate] DATETIME NOT NULL CONSTRAINT [DF_CustomerPayments_PaymentDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.CustomerPayments', N'Amount') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [Amount] MONEY NOT NULL CONSTRAINT [DF_CustomerPayments_Amount] DEFAULT (0);
IF COL_LENGTH(N'dbo.CustomerPayments', N'PaymentMethodID') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [PaymentMethodID] INT NOT NULL CONSTRAINT [DF_CustomerPayments_PaymentMethodID] DEFAULT (1);
IF COL_LENGTH(N'dbo.CustomerPayments', N'SalesInvoiceID') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [SalesInvoiceID] INT NULL;
IF COL_LENGTH(N'dbo.CustomerPayments', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CustomerPayments', N'Notes') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.CustomerPayments', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CustomerPayments_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.CustomerPayments', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [CashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.CustomerPayments', N'BankID') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [BankID] INT NULL;
IF COL_LENGTH(N'dbo.CustomerPayments', N'CurrencyCode') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_CustomerPayments_CurrencyCode] DEFAULT (N'SAR');
IF COL_LENGTH(N'dbo.CustomerPayments', N'ExchangeRate') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_CustomerPayments_ExchangeRate] DEFAULT (1);
IF COL_LENGTH(N'dbo.CustomerPayments', N'ForeignAmount') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_CustomerPayments_ForeignAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.CustomerPayments', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [SalesRepID] INT NULL;
IF COL_LENGTH(N'dbo.CustomerPayments', N'RowVer') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.SupplierPayments', N'PaymentID') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [PaymentID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.SupplierPayments', N'PaymentNumber') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [PaymentNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.SupplierPayments', N'SupplierID') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [SupplierID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SupplierPayments', N'PaymentDate') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [PaymentDate] DATETIME NOT NULL CONSTRAINT [DF_SupplierPayments_PaymentDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.SupplierPayments', N'Amount') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [Amount] MONEY NOT NULL CONSTRAINT [DF_SupplierPayments_Amount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SupplierPayments', N'PaymentMethodID') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [PaymentMethodID] INT NOT NULL CONSTRAINT [DF_SupplierPayments_PaymentMethodID] DEFAULT (1);
IF COL_LENGTH(N'dbo.SupplierPayments', N'PurchaseInvoiceID') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [PurchaseInvoiceID] INT NULL;
IF COL_LENGTH(N'dbo.SupplierPayments', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SupplierPayments', N'Notes') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SupplierPayments', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SupplierPayments_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.SupplierPayments', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [CashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.SupplierPayments', N'BankID') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [BankID] INT NULL;
IF COL_LENGTH(N'dbo.SupplierPayments', N'CurrencyCode') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_SupplierPayments_CurrencyCode] DEFAULT (N'SAR');
IF COL_LENGTH(N'dbo.SupplierPayments', N'ExchangeRate') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_SupplierPayments_ExchangeRate] DEFAULT (1);
IF COL_LENGTH(N'dbo.SupplierPayments', N'ForeignAmount') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_SupplierPayments_ForeignAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SupplierPayments', N'RowVer') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Banks', N'BankID') IS NULL
    ALTER TABLE [dbo].[Banks] ADD [BankID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.Banks', N'BankName') IS NULL
    ALTER TABLE [dbo].[Banks] ADD [BankName] NVARCHAR(100) NOT NULL;
IF COL_LENGTH(N'dbo.Banks', N'BankNameEn') IS NULL
    ALTER TABLE [dbo].[Banks] ADD [BankNameEn] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Banks', N'AccountNo') IS NULL
    ALTER TABLE [dbo].[Banks] ADD [AccountNo] NVARCHAR(30) NULL;
IF COL_LENGTH(N'dbo.Banks', N'IBAN') IS NULL
    ALTER TABLE [dbo].[Banks] ADD [IBAN] NVARCHAR(34) NULL;
IF COL_LENGTH(N'dbo.Banks', N'OpeningBalance') IS NULL
    ALTER TABLE [dbo].[Banks] ADD [OpeningBalance] MONEY NOT NULL CONSTRAINT [DF_Banks_OpeningBalance] DEFAULT (0);
IF COL_LENGTH(N'dbo.Banks', N'OpeningDate') IS NULL
    ALTER TABLE [dbo].[Banks] ADD [OpeningDate] DATETIME NOT NULL CONSTRAINT [DF_Banks_OpeningDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.Banks', N'IsActive') IS NULL
    ALTER TABLE [dbo].[Banks] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_Banks_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.Banks', N'Notes') IS NULL
    ALTER TABLE [dbo].[Banks] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Banks', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[Banks] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Banks_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.Banks', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Banks] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.BankTransactions', N'BankTxID') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [BankTxID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.BankTransactions', N'TxNumber') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [TxNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.BankTransactions', N'TxDate') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [TxDate] DATETIME NOT NULL CONSTRAINT [DF_BankTransactions_TxDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.BankTransactions', N'TxType') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [TxType] NVARCHAR(12) NOT NULL;
IF COL_LENGTH(N'dbo.BankTransactions', N'BankID') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [BankID] INT NOT NULL;
IF COL_LENGTH(N'dbo.BankTransactions', N'ToBankID') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [ToBankID] INT NULL;
IF COL_LENGTH(N'dbo.BankTransactions', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [CashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.BankTransactions', N'CounterAccount') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [CounterAccount] INT NULL;
IF COL_LENGTH(N'dbo.BankTransactions', N'Amount') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [Amount] MONEY NOT NULL CONSTRAINT [DF_BankTransactions_Amount] DEFAULT (0);
IF COL_LENGTH(N'dbo.BankTransactions', N'FeeAmount') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [FeeAmount] MONEY NOT NULL CONSTRAINT [DF_BankTransactions_FeeAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.BankTransactions', N'FeeVAT') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [FeeVAT] MONEY NOT NULL CONSTRAINT [DF_BankTransactions_FeeVAT] DEFAULT (0);
IF COL_LENGTH(N'dbo.BankTransactions', N'Reference') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [Reference] NVARCHAR(40) NULL;
IF COL_LENGTH(N'dbo.BankTransactions', N'Description') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [Description] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.BankTransactions', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.BankTransactions', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_BankTransactions_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.BankTransactions', N'RowVer') IS NULL
    ALTER TABLE [dbo].[BankTransactions] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Cheques', N'ChequeID') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [ChequeID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.Cheques', N'ChequeRef') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [ChequeRef] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.Cheques', N'Direction') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [Direction] NVARCHAR(3) NOT NULL;
IF COL_LENGTH(N'dbo.Cheques', N'CustomerID') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [CustomerID] INT NULL;
IF COL_LENGTH(N'dbo.Cheques', N'SupplierID') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [SupplierID] INT NULL;
IF COL_LENGTH(N'dbo.Cheques', N'ChequeNo') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [ChequeNo] NVARCHAR(30) NOT NULL;
IF COL_LENGTH(N'dbo.Cheques', N'DrawerBank') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [DrawerBank] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Cheques', N'BankID') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [BankID] INT NULL;
IF COL_LENGTH(N'dbo.Cheques', N'IssueDate') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [IssueDate] DATETIME NOT NULL CONSTRAINT [DF_Cheques_IssueDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.Cheques', N'DueDate') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [DueDate] DATETIME NOT NULL CONSTRAINT [DF_Cheques_DueDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.Cheques', N'Amount') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [Amount] MONEY NOT NULL CONSTRAINT [DF_Cheques_Amount] DEFAULT (0);
IF COL_LENGTH(N'dbo.Cheques', N'Status') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_Cheques_Status] DEFAULT (N'PENDING');
IF COL_LENGTH(N'dbo.Cheques', N'StatusDate') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [StatusDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.Cheques', N'Notes') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Cheques', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.Cheques', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Cheques_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.Cheques', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Cheques] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'AssetID') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [AssetID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'AssetCode') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [AssetCode] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'AssetName') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [AssetName] NVARCHAR(150) NOT NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'AssetAccount') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [AssetAccount] INT NOT NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'PurchaseDate') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [PurchaseDate] DATETIME NOT NULL CONSTRAINT [DF_FixedAssets_PurchaseDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.FixedAssets', N'Cost') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [Cost] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_Cost] DEFAULT (0);
IF COL_LENGTH(N'dbo.FixedAssets', N'InputVAT') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [InputVAT] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_InputVAT] DEFAULT (0);
IF COL_LENGTH(N'dbo.FixedAssets', N'SalvageValue') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [SalvageValue] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_SalvageValue] DEFAULT (0);
IF COL_LENGTH(N'dbo.FixedAssets', N'UsefulLifeMonths') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [UsefulLifeMonths] SMALLINT NOT NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'DepStartDate') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [DepStartDate] DATETIME NOT NULL CONSTRAINT [DF_FixedAssets_DepStartDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.FixedAssets', N'DepMethod') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [DepMethod] NVARCHAR(2) NOT NULL CONSTRAINT [DF_FixedAssets_DepMethod] DEFAULT (N'SL');
IF COL_LENGTH(N'dbo.FixedAssets', N'DepRate') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [DepRate] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_DepRate] DEFAULT (0);
IF COL_LENGTH(N'dbo.FixedAssets', N'SourceType') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [SourceType] NVARCHAR(10) NOT NULL CONSTRAINT [DF_FixedAssets_SourceType] DEFAULT (N'BANK');
IF COL_LENGTH(N'dbo.FixedAssets', N'BankID') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [BankID] INT NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [CashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'CounterAccount') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [CounterAccount] INT NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'OpeningAccumDep') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [OpeningAccumDep] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_OpeningAccumDep] DEFAULT (0);
IF COL_LENGTH(N'dbo.FixedAssets', N'Status') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_FixedAssets_Status] DEFAULT (N'ACTIVE');
IF COL_LENGTH(N'dbo.FixedAssets', N'DisposalDate') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [DisposalDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'DisposalProceeds') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [DisposalProceeds] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_DisposalProceeds] DEFAULT (0);
IF COL_LENGTH(N'dbo.FixedAssets', N'DisposalTo') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [DisposalTo] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'DisposalBankID') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [DisposalBankID] INT NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'DisposalCashBoxID') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [DisposalCashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'DisposalAccumDep') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [DisposalAccumDep] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_DisposalAccumDep] DEFAULT (0);
IF COL_LENGTH(N'dbo.FixedAssets', N'DisposalVAT') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [DisposalVAT] MONEY NOT NULL CONSTRAINT [DF_FixedAssets_DisposalVAT] DEFAULT (0);
IF COL_LENGTH(N'dbo.FixedAssets', N'Notes') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.FixedAssets', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_FixedAssets_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.FixedAssets', N'RowVer') IS NULL
    ALTER TABLE [dbo].[FixedAssets] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.DepreciationRuns', N'RunID') IS NULL
    ALTER TABLE [dbo].[DepreciationRuns] ADD [RunID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.DepreciationRuns', N'RunNumber') IS NULL
    ALTER TABLE [dbo].[DepreciationRuns] ADD [RunNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.DepreciationRuns', N'RunMonth') IS NULL
    ALTER TABLE [dbo].[DepreciationRuns] ADD [RunMonth] DATETIME NOT NULL CONSTRAINT [DF_DepreciationRuns_RunMonth] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.DepreciationRuns', N'TotalAmount') IS NULL
    ALTER TABLE [dbo].[DepreciationRuns] ADD [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_DepreciationRuns_TotalAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.DepreciationRuns', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[DepreciationRuns] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.DepreciationRuns', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[DepreciationRuns] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_DepreciationRuns_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.DepreciationRuns', N'RowVer') IS NULL
    ALTER TABLE [dbo].[DepreciationRuns] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.AssetDepreciations', N'DepreciationID') IS NULL
    ALTER TABLE [dbo].[AssetDepreciations] ADD [DepreciationID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.AssetDepreciations', N'RunID') IS NULL
    ALTER TABLE [dbo].[AssetDepreciations] ADD [RunID] INT NOT NULL;
IF COL_LENGTH(N'dbo.AssetDepreciations', N'LineNo') IS NULL
    ALTER TABLE [dbo].[AssetDepreciations] ADD [LineNo] SMALLINT NOT NULL;
IF COL_LENGTH(N'dbo.AssetDepreciations', N'AssetID') IS NULL
    ALTER TABLE [dbo].[AssetDepreciations] ADD [AssetID] INT NOT NULL;
IF COL_LENGTH(N'dbo.AssetDepreciations', N'Amount') IS NULL
    ALTER TABLE [dbo].[AssetDepreciations] ADD [Amount] MONEY NOT NULL CONSTRAINT [DF_AssetDepreciations_Amount] DEFAULT (0);
IF COL_LENGTH(N'dbo.AssetDepreciations', N'DepCenter') IS NULL
    ALTER TABLE [dbo].[AssetDepreciations] ADD [DepCenter] INT NULL;
IF COL_LENGTH(N'dbo.AssetDepreciations', N'RowVer') IS NULL
    ALTER TABLE [dbo].[AssetDepreciations] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.AssetRevaluations', N'RevaluationID') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [RevaluationID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.AssetRevaluations', N'RevalNumber') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [RevalNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.AssetRevaluations', N'AssetID') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [AssetID] INT NOT NULL;
IF COL_LENGTH(N'dbo.AssetRevaluations', N'RevalDate') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [RevalDate] DATETIME NOT NULL CONSTRAINT [DF_AssetRevaluations_RevalDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.AssetRevaluations', N'BookValueBefore') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [BookValueBefore] MONEY NOT NULL CONSTRAINT [DF_AssetRevaluations_BookValueBefore] DEFAULT (0);
IF COL_LENGTH(N'dbo.AssetRevaluations', N'NewValue') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [NewValue] MONEY NOT NULL CONSTRAINT [DF_AssetRevaluations_NewValue] DEFAULT (0);
IF COL_LENGTH(N'dbo.AssetRevaluations', N'ValueChange') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [ValueChange] MONEY NOT NULL CONSTRAINT [DF_AssetRevaluations_ValueChange] DEFAULT (0);
IF COL_LENGTH(N'dbo.AssetRevaluations', N'SurplusPart') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [SurplusPart] MONEY NOT NULL CONSTRAINT [DF_AssetRevaluations_SurplusPart] DEFAULT (0);
IF COL_LENGTH(N'dbo.AssetRevaluations', N'LossPart') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [LossPart] MONEY NOT NULL CONSTRAINT [DF_AssetRevaluations_LossPart] DEFAULT (0);
IF COL_LENGTH(N'dbo.AssetRevaluations', N'MonthsLeft') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [MonthsLeft] SMALLINT NOT NULL;
IF COL_LENGTH(N'dbo.AssetRevaluations', N'RevalCenter') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [RevalCenter] INT NULL;
IF COL_LENGTH(N'dbo.AssetRevaluations', N'Notes') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.AssetRevaluations', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.AssetRevaluations', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_AssetRevaluations_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.AssetRevaluations', N'RowVer') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.AssetTransfers', N'TransferID') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] ADD [TransferID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.AssetTransfers', N'TransferNumber') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] ADD [TransferNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.AssetTransfers', N'AssetID') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] ADD [AssetID] INT NOT NULL;
IF COL_LENGTH(N'dbo.AssetTransfers', N'TransferDate') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] ADD [TransferDate] DATETIME NOT NULL CONSTRAINT [DF_AssetTransfers_TransferDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.AssetTransfers', N'FromCenterID') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] ADD [FromCenterID] INT NULL;
IF COL_LENGTH(N'dbo.AssetTransfers', N'ToCenterID') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] ADD [ToCenterID] INT NULL;
IF COL_LENGTH(N'dbo.AssetTransfers', N'AssetValue') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] ADD [AssetValue] MONEY NOT NULL CONSTRAINT [DF_AssetTransfers_AssetValue] DEFAULT (0);
IF COL_LENGTH(N'dbo.AssetTransfers', N'AccumDep') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] ADD [AccumDep] MONEY NOT NULL CONSTRAINT [DF_AssetTransfers_AccumDep] DEFAULT (0);
IF COL_LENGTH(N'dbo.AssetTransfers', N'Notes') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.AssetTransfers', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.AssetTransfers', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_AssetTransfers_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.AssetTransfers', N'RowVer') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.CostCenters', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[CostCenters] ADD [CostCenterID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.CostCenters', N'CenterCode') IS NULL
    ALTER TABLE [dbo].[CostCenters] ADD [CenterCode] NVARCHAR(10) NOT NULL;
IF COL_LENGTH(N'dbo.CostCenters', N'CenterName') IS NULL
    ALTER TABLE [dbo].[CostCenters] ADD [CenterName] NVARCHAR(100) NOT NULL;
IF COL_LENGTH(N'dbo.CostCenters', N'CenterNameEn') IS NULL
    ALTER TABLE [dbo].[CostCenters] ADD [CenterNameEn] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.CostCenters', N'IsDefault') IS NULL
    ALTER TABLE [dbo].[CostCenters] ADD [IsDefault] BIT NOT NULL CONSTRAINT [DF_CostCenters_IsDefault] DEFAULT (0);
IF COL_LENGTH(N'dbo.CostCenters', N'IsActive') IS NULL
    ALTER TABLE [dbo].[CostCenters] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_CostCenters_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.CostCenters', N'Notes') IS NULL
    ALTER TABLE [dbo].[CostCenters] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.CostCenters', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[CostCenters] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CostCenters_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.CostCenters', N'RowVer') IS NULL
    ALTER TABLE [dbo].[CostCenters] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.SalesReps', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [SalesRepID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.SalesReps', N'RepCode') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [RepCode] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.SalesReps', N'RepName') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [RepName] NVARCHAR(100) NOT NULL;
IF COL_LENGTH(N'dbo.SalesReps', N'RepNameEn') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [RepNameEn] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.SalesReps', N'Mobile') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [Mobile] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.SalesReps', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [EmployeeID] INT NULL;
IF COL_LENGTH(N'dbo.SalesReps', N'Region') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [Region] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.SalesReps', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.SalesReps', N'CommissionRate') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [CommissionRate] MONEY NOT NULL CONSTRAINT [DF_SalesReps_CommissionRate] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReps', N'CommissionBase') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [CommissionBase] NVARCHAR(10) NOT NULL CONSTRAINT [DF_SalesReps_CommissionBase] DEFAULT (N'SALES');
IF COL_LENGTH(N'dbo.SalesReps', N'CreditLimit') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [CreditLimit] MONEY NOT NULL CONSTRAINT [DF_SalesReps_CreditLimit] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReps', N'MaxDiscountPercent') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [MaxDiscountPercent] MONEY NOT NULL CONSTRAINT [DF_SalesReps_MaxDiscountPercent] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesReps', N'IsActive') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_SalesReps_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.SalesReps', N'Notes') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SalesReps', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SalesReps_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.SalesReps', N'RowVer') IS NULL
    ALTER TABLE [dbo].[SalesReps] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.SalesRepTargets', N'TargetID') IS NULL
    ALTER TABLE [dbo].[SalesRepTargets] ADD [TargetID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.SalesRepTargets', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[SalesRepTargets] ADD [SalesRepID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesRepTargets', N'TargetYear') IS NULL
    ALTER TABLE [dbo].[SalesRepTargets] ADD [TargetYear] SMALLINT NOT NULL;
IF COL_LENGTH(N'dbo.SalesRepTargets', N'TargetMonth') IS NULL
    ALTER TABLE [dbo].[SalesRepTargets] ADD [TargetMonth] TINYINT NOT NULL CONSTRAINT [DF_SalesRepTargets_TargetMonth] DEFAULT (1);
IF COL_LENGTH(N'dbo.SalesRepTargets', N'TargetAmount') IS NULL
    ALTER TABLE [dbo].[SalesRepTargets] ADD [TargetAmount] MONEY NOT NULL CONSTRAINT [DF_SalesRepTargets_TargetAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesRepTargets', N'RowVer') IS NULL
    ALTER TABLE [dbo].[SalesRepTargets] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.CommissionRuns', N'CommissionRunID') IS NULL
    ALTER TABLE [dbo].[CommissionRuns] ADD [CommissionRunID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.CommissionRuns', N'RunNumber') IS NULL
    ALTER TABLE [dbo].[CommissionRuns] ADD [RunNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.CommissionRuns', N'RunMonth') IS NULL
    ALTER TABLE [dbo].[CommissionRuns] ADD [RunMonth] DATETIME NOT NULL CONSTRAINT [DF_CommissionRuns_RunMonth] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.CommissionRuns', N'Status') IS NULL
    ALTER TABLE [dbo].[CommissionRuns] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_CommissionRuns_Status] DEFAULT (N'DRAFT');
IF COL_LENGTH(N'dbo.CommissionRuns', N'TotalAmount') IS NULL
    ALTER TABLE [dbo].[CommissionRuns] ADD [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_CommissionRuns_TotalAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.CommissionRuns', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[CommissionRuns] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CommissionRuns', N'PostedAt') IS NULL
    ALTER TABLE [dbo].[CommissionRuns] ADD [PostedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.CommissionRuns', N'Notes') IS NULL
    ALTER TABLE [dbo].[CommissionRuns] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.CommissionRuns', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[CommissionRuns] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CommissionRuns_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.CommissionRuns', N'RowVer') IS NULL
    ALTER TABLE [dbo].[CommissionRuns] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.CommissionLines', N'CommissionLineID') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [CommissionLineID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.CommissionLines', N'CommissionRunID') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [CommissionRunID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CommissionLines', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [SalesRepID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CommissionLines', N'RepName') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [RepName] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.CommissionLines', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.CommissionLines', N'NetSales') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [NetSales] MONEY NOT NULL CONSTRAINT [DF_CommissionLines_NetSales] DEFAULT (0);
IF COL_LENGTH(N'dbo.CommissionLines', N'Collections') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [Collections] MONEY NOT NULL CONSTRAINT [DF_CommissionLines_Collections] DEFAULT (0);
IF COL_LENGTH(N'dbo.CommissionLines', N'CommissionBase') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [CommissionBase] NVARCHAR(10) NOT NULL CONSTRAINT [DF_CommissionLines_CommissionBase] DEFAULT (N'SALES');
IF COL_LENGTH(N'dbo.CommissionLines', N'BaseAmount') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [BaseAmount] MONEY NOT NULL CONSTRAINT [DF_CommissionLines_BaseAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.CommissionLines', N'CommissionRate') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [CommissionRate] MONEY NOT NULL CONSTRAINT [DF_CommissionLines_CommissionRate] DEFAULT (0);
IF COL_LENGTH(N'dbo.CommissionLines', N'Adjustment') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [Adjustment] MONEY NOT NULL CONSTRAINT [DF_CommissionLines_Adjustment] DEFAULT (0);
IF COL_LENGTH(N'dbo.CommissionLines', N'Commission') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [Commission] MONEY NOT NULL CONSTRAINT [DF_CommissionLines_Commission] DEFAULT (0);
IF COL_LENGTH(N'dbo.CommissionLines', N'Notes') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [Notes] NVARCHAR(150) NULL;
IF COL_LENGTH(N'dbo.CommissionLines', N'RowVer') IS NULL
    ALTER TABLE [dbo].[CommissionLines] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Budgets', N'BudgetID') IS NULL
    ALTER TABLE [dbo].[Budgets] ADD [BudgetID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.Budgets', N'BudgetYear') IS NULL
    ALTER TABLE [dbo].[Budgets] ADD [BudgetYear] SMALLINT NOT NULL;
IF COL_LENGTH(N'dbo.Budgets', N'BudgetName') IS NULL
    ALTER TABLE [dbo].[Budgets] ADD [BudgetName] NVARCHAR(100) NOT NULL;
IF COL_LENGTH(N'dbo.Budgets', N'Notes') IS NULL
    ALTER TABLE [dbo].[Budgets] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Budgets', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[Budgets] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.Budgets', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[Budgets] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Budgets_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.Budgets', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Budgets] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.BudgetLines', N'BudgetLineID') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [BudgetLineID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.BudgetLines', N'BudgetID') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [BudgetID] INT NOT NULL;
IF COL_LENGTH(N'dbo.BudgetLines', N'AccountCode') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [AccountCode] INT NOT NULL;
IF COL_LENGTH(N'dbo.BudgetLines', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.BudgetLines', N'M1') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [M1] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M1] DEFAULT (0);
IF COL_LENGTH(N'dbo.BudgetLines', N'M2') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [M2] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M2] DEFAULT (0);
IF COL_LENGTH(N'dbo.BudgetLines', N'M3') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [M3] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M3] DEFAULT (0);
IF COL_LENGTH(N'dbo.BudgetLines', N'M4') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [M4] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M4] DEFAULT (0);
IF COL_LENGTH(N'dbo.BudgetLines', N'M5') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [M5] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M5] DEFAULT (0);
IF COL_LENGTH(N'dbo.BudgetLines', N'M6') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [M6] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M6] DEFAULT (0);
IF COL_LENGTH(N'dbo.BudgetLines', N'M7') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [M7] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M7] DEFAULT (0);
IF COL_LENGTH(N'dbo.BudgetLines', N'M8') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [M8] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M8] DEFAULT (0);
IF COL_LENGTH(N'dbo.BudgetLines', N'M9') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [M9] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M9] DEFAULT (0);
IF COL_LENGTH(N'dbo.BudgetLines', N'M10') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [M10] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M10] DEFAULT (0);
IF COL_LENGTH(N'dbo.BudgetLines', N'M11') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [M11] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M11] DEFAULT (0);
IF COL_LENGTH(N'dbo.BudgetLines', N'M12') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [M12] MONEY NOT NULL CONSTRAINT [DF_BudgetLines_M12] DEFAULT (0);
IF COL_LENGTH(N'dbo.BudgetLines', N'Notes') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [Notes] NVARCHAR(150) NULL;
IF COL_LENGTH(N'dbo.BudgetLines', N'RowVer') IS NULL
    ALTER TABLE [dbo].[BudgetLines] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.PayrollRuns', N'PayrollRunID') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] ADD [PayrollRunID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.PayrollRuns', N'RunNumber') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] ADD [RunNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.PayrollRuns', N'PayMonth') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] ADD [PayMonth] DATETIME NOT NULL CONSTRAINT [DF_PayrollRuns_PayMonth] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.PayrollRuns', N'Status') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_PayrollRuns_Status] DEFAULT (N'DRAFT');
IF COL_LENGTH(N'dbo.PayrollRuns', N'PaidDate') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] ADD [PaidDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.PayrollRuns', N'PaidFrom') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] ADD [PaidFrom] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.PayrollRuns', N'BankID') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] ADD [BankID] INT NULL;
IF COL_LENGTH(N'dbo.PayrollRuns', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] ADD [CashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.PayrollRuns', N'PaidAmount') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] ADD [PaidAmount] MONEY NOT NULL CONSTRAINT [DF_PayrollRuns_PaidAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollRuns', N'Notes') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.PayrollRuns', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PayrollRuns', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PayrollRuns_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.PayrollRuns', N'RowVer') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.PayrollLines', N'PayrollLineID') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [PayrollLineID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.PayrollLines', N'PayrollRunID') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [PayrollRunID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PayrollLines', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PayrollLines', N'EmployeeName') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [EmployeeName] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.PayrollLines', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.PayrollLines', N'IsSaudi') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [IsSaudi] BIT NOT NULL CONSTRAINT [DF_PayrollLines_IsSaudi] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollLines', N'Basic') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [Basic] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_Basic] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollLines', N'Housing') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [Housing] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_Housing] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollLines', N'OtherAllow') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [OtherAllow] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_OtherAllow] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollLines', N'Overtime') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [Overtime] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_Overtime] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollLines', N'Additions') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [Additions] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_Additions] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollLines', N'AbsenceDeduction') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [AbsenceDeduction] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_AbsenceDeduction] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollLines', N'AdvanceDeduction') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [AdvanceDeduction] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_AdvanceDeduction] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollLines', N'OtherDeduction') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [OtherDeduction] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_OtherDeduction] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollLines', N'GosiWage') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [GosiWage] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_GosiWage] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollLines', N'GosiEmployee') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [GosiEmployee] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_GosiEmployee] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollLines', N'GosiEmployer') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [GosiEmployer] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_GosiEmployer] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollLines', N'NetPay') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [NetPay] MONEY NOT NULL CONSTRAINT [DF_PayrollLines_NetPay] DEFAULT (0);
IF COL_LENGTH(N'dbo.PayrollLines', N'Notes') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [Notes] NVARCHAR(150) NULL;
IF COL_LENGTH(N'dbo.PayrollLines', N'RowVer') IS NULL
    ALTER TABLE [dbo].[PayrollLines] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.BankReconciliations', N'ReconciliationID') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] ADD [ReconciliationID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.BankReconciliations', N'ReconNumber') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] ADD [ReconNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.BankReconciliations', N'BankID') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] ADD [BankID] INT NOT NULL;
IF COL_LENGTH(N'dbo.BankReconciliations', N'StatementDate') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] ADD [StatementDate] DATETIME NOT NULL CONSTRAINT [DF_BankReconciliations_StatementDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.BankReconciliations', N'StatementBalance') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] ADD [StatementBalance] MONEY NOT NULL CONSTRAINT [DF_BankReconciliations_StatementBalance] DEFAULT (0);
IF COL_LENGTH(N'dbo.BankReconciliations', N'BookBalance') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] ADD [BookBalance] MONEY NOT NULL CONSTRAINT [DF_BankReconciliations_BookBalance] DEFAULT (0);
IF COL_LENGTH(N'dbo.BankReconciliations', N'Outstanding') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] ADD [Outstanding] MONEY NOT NULL CONSTRAINT [DF_BankReconciliations_Outstanding] DEFAULT (0);
IF COL_LENGTH(N'dbo.BankReconciliations', N'Difference') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] ADD [Difference] MONEY NOT NULL CONSTRAINT [DF_BankReconciliations_Difference] DEFAULT (0);
IF COL_LENGTH(N'dbo.BankReconciliations', N'Status') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_BankReconciliations_Status] DEFAULT (N'OPEN');
IF COL_LENGTH(N'dbo.BankReconciliations', N'Notes') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.BankReconciliations', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.BankReconciliations', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_BankReconciliations_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.BankReconciliations', N'RowVer') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.BankClearings', N'ClearingID') IS NULL
    ALTER TABLE [dbo].[BankClearings] ADD [ClearingID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.BankClearings', N'ReconciliationID') IS NULL
    ALTER TABLE [dbo].[BankClearings] ADD [ReconciliationID] INT NOT NULL;
IF COL_LENGTH(N'dbo.BankClearings', N'BankID') IS NULL
    ALTER TABLE [dbo].[BankClearings] ADD [BankID] INT NOT NULL;
IF COL_LENGTH(N'dbo.BankClearings', N'SourceType') IS NULL
    ALTER TABLE [dbo].[BankClearings] ADD [SourceType] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.BankClearings', N'SourceID') IS NULL
    ALTER TABLE [dbo].[BankClearings] ADD [SourceID] INT NOT NULL;
IF COL_LENGTH(N'dbo.BankClearings', N'ClearedAmount') IS NULL
    ALTER TABLE [dbo].[BankClearings] ADD [ClearedAmount] MONEY NOT NULL CONSTRAINT [DF_BankClearings_ClearedAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.BankClearings', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[BankClearings] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_BankClearings_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.BankClearings', N'RowVer') IS NULL
    ALTER TABLE [dbo].[BankClearings] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.CustomerAllocations', N'AllocationID') IS NULL
    ALTER TABLE [dbo].[CustomerAllocations] ADD [AllocationID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.CustomerAllocations', N'PaymentID') IS NULL
    ALTER TABLE [dbo].[CustomerAllocations] ADD [PaymentID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CustomerAllocations', N'SalesInvoiceID') IS NULL
    ALTER TABLE [dbo].[CustomerAllocations] ADD [SalesInvoiceID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CustomerAllocations', N'Amount') IS NULL
    ALTER TABLE [dbo].[CustomerAllocations] ADD [Amount] MONEY NOT NULL CONSTRAINT [DF_CustomerAllocations_Amount] DEFAULT (0);
IF COL_LENGTH(N'dbo.CustomerAllocations', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[CustomerAllocations] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CustomerAllocations', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[CustomerAllocations] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CustomerAllocations_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.CustomerAllocations', N'RowVer') IS NULL
    ALTER TABLE [dbo].[CustomerAllocations] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.SupplierAllocations', N'AllocationID') IS NULL
    ALTER TABLE [dbo].[SupplierAllocations] ADD [AllocationID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.SupplierAllocations', N'PaymentID') IS NULL
    ALTER TABLE [dbo].[SupplierAllocations] ADD [PaymentID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SupplierAllocations', N'PurchaseInvoiceID') IS NULL
    ALTER TABLE [dbo].[SupplierAllocations] ADD [PurchaseInvoiceID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SupplierAllocations', N'Amount') IS NULL
    ALTER TABLE [dbo].[SupplierAllocations] ADD [Amount] MONEY NOT NULL CONSTRAINT [DF_SupplierAllocations_Amount] DEFAULT (0);
IF COL_LENGTH(N'dbo.SupplierAllocations', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[SupplierAllocations] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SupplierAllocations', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[SupplierAllocations] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SupplierAllocations_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.SupplierAllocations', N'RowVer') IS NULL
    ALTER TABLE [dbo].[SupplierAllocations] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.ExpenseTypes', N'ExpenseTypeID') IS NULL
    ALTER TABLE [dbo].[ExpenseTypes] ADD [ExpenseTypeID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.ExpenseTypes', N'ExpenseTypeName') IS NULL
    ALTER TABLE [dbo].[ExpenseTypes] ADD [ExpenseTypeName] NVARCHAR(50) NOT NULL;
IF COL_LENGTH(N'dbo.ExpenseTypes', N'ExpenseTypeNameEn') IS NULL
    ALTER TABLE [dbo].[ExpenseTypes] ADD [ExpenseTypeNameEn] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.ExpenseTypes', N'IsActive') IS NULL
    ALTER TABLE [dbo].[ExpenseTypes] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_ExpenseTypes_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.ExpenseTypes', N'RowVer') IS NULL
    ALTER TABLE [dbo].[ExpenseTypes] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Expenses', N'ExpenseID') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [ExpenseID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.Expenses', N'ExpenseNumber') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [ExpenseNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.Expenses', N'ExpenseDate') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [ExpenseDate] DATETIME NOT NULL CONSTRAINT [DF_Expenses_ExpenseDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.Expenses', N'ExpenseTypeID') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [ExpenseTypeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.Expenses', N'Amount') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [Amount] MONEY NOT NULL CONSTRAINT [DF_Expenses_Amount] DEFAULT (0);
IF COL_LENGTH(N'dbo.Expenses', N'Tax') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [Tax] MONEY NOT NULL CONSTRAINT [DF_Expenses_Tax] DEFAULT (0);
IF COL_LENGTH(N'dbo.Expenses', N'TotalAmount') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_Expenses_TotalAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.Expenses', N'PaymentMethodID') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [PaymentMethodID] INT NULL;
IF COL_LENGTH(N'dbo.Expenses', N'SupplierInvoiceRef') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [SupplierInvoiceRef] NVARCHAR(30) NULL;
IF COL_LENGTH(N'dbo.Expenses', N'Description') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [Description] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Expenses', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.Expenses', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Expenses_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.Expenses', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [CashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.Expenses', N'BankID') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [BankID] INT NULL;
IF COL_LENGTH(N'dbo.Expenses', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.Expenses', N'RecurringID') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [RecurringID] INT NULL;
IF COL_LENGTH(N'dbo.Expenses', N'CurrencyCode') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_Expenses_CurrencyCode] DEFAULT (N'SAR');
IF COL_LENGTH(N'dbo.Expenses', N'ExchangeRate') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_Expenses_ExchangeRate] DEFAULT (1);
IF COL_LENGTH(N'dbo.Expenses', N'ForeignAmount') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_Expenses_ForeignAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.Expenses', N'ForeignTax') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [ForeignTax] MONEY NOT NULL CONSTRAINT [DF_Expenses_ForeignTax] DEFAULT (0);
IF COL_LENGTH(N'dbo.Expenses', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Expenses] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.RecurringExpenses', N'RecurringID') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [RecurringID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.RecurringExpenses', N'RecurringName') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [RecurringName] NVARCHAR(100) NOT NULL;
IF COL_LENGTH(N'dbo.RecurringExpenses', N'ExpenseTypeID') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [ExpenseTypeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.RecurringExpenses', N'Amount') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [Amount] MONEY NOT NULL CONSTRAINT [DF_RecurringExpenses_Amount] DEFAULT (0);
IF COL_LENGTH(N'dbo.RecurringExpenses', N'Tax') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [Tax] MONEY NOT NULL CONSTRAINT [DF_RecurringExpenses_Tax] DEFAULT (0);
IF COL_LENGTH(N'dbo.RecurringExpenses', N'PaymentMethodID') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [PaymentMethodID] INT NOT NULL;
IF COL_LENGTH(N'dbo.RecurringExpenses', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [CashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.RecurringExpenses', N'BankID') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [BankID] INT NULL;
IF COL_LENGTH(N'dbo.RecurringExpenses', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.RecurringExpenses', N'Frequency') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [Frequency] NVARCHAR(10) NOT NULL CONSTRAINT [DF_RecurringExpenses_Frequency] DEFAULT (N'MONTHLY');
IF COL_LENGTH(N'dbo.RecurringExpenses', N'DueDay') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [DueDay] SMALLINT NOT NULL CONSTRAINT [DF_RecurringExpenses_DueDay] DEFAULT (1);
IF COL_LENGTH(N'dbo.RecurringExpenses', N'StartDate') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [StartDate] DATETIME NOT NULL CONSTRAINT [DF_RecurringExpenses_StartDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.RecurringExpenses', N'EndDate') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [EndDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.RecurringExpenses', N'NextDueDate') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [NextDueDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.RecurringExpenses', N'LastCreatedDate') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [LastCreatedDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.RecurringExpenses', N'Description') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [Description] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.RecurringExpenses', N'IsActive') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_RecurringExpenses_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.RecurringExpenses', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_RecurringExpenses_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.RecurringExpenses', N'RowVer') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'CashVoucherID') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [CashVoucherID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'VoucherNumber') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [VoucherNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'VoucherDate') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [VoucherDate] DATETIME NOT NULL CONSTRAINT [DF_CashVouchers_VoucherDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.CashVouchers', N'VoucherType') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [VoucherType] NVARCHAR(10) NOT NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [CashBoxID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'ToCashBoxID') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [ToCashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'Category') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [Category] NVARCHAR(10) NOT NULL CONSTRAINT [DF_CashVouchers_Category] DEFAULT (N'OTHER');
IF COL_LENGTH(N'dbo.CashVouchers', N'Amount') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [Amount] MONEY NOT NULL CONSTRAINT [DF_CashVouchers_Amount] DEFAULT (0);
IF COL_LENGTH(N'dbo.CashVouchers', N'PartyName') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [PartyName] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'Description') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [Description] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'ExpenseID') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [ExpenseID] INT NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'ClosingID') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [ClosingID] INT NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'AdvanceEmployeeID') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [AdvanceEmployeeID] INT NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CashVouchers_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.CashVouchers', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [SalesRepID] INT NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'InTransit') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [InTransit] BIT NOT NULL CONSTRAINT [DF_CashVouchers_InTransit] DEFAULT (0);
IF COL_LENGTH(N'dbo.CashVouchers', N'ReceivedAt') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [ReceivedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'ReceivedByID') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [ReceivedByID] INT NULL;
IF COL_LENGTH(N'dbo.CashVouchers', N'RowVer') IS NULL
    ALTER TABLE [dbo].[CashVouchers] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.CashClosings', N'ClosingID') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [ClosingID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.CashClosings', N'ClosingNumber') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [ClosingNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.CashClosings', N'ClosingDate') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [ClosingDate] DATETIME NOT NULL CONSTRAINT [DF_CashClosings_ClosingDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.CashClosings', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [CashBoxID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CashClosings', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.CashClosings', N'PeriodStart') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [PeriodStart] DATETIME NULL;
IF COL_LENGTH(N'dbo.CashClosings', N'OpeningBalance') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [OpeningBalance] MONEY NOT NULL CONSTRAINT [DF_CashClosings_OpeningBalance] DEFAULT (0);
IF COL_LENGTH(N'dbo.CashClosings', N'CashIn') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [CashIn] MONEY NOT NULL CONSTRAINT [DF_CashClosings_CashIn] DEFAULT (0);
IF COL_LENGTH(N'dbo.CashClosings', N'CashOut') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [CashOut] MONEY NOT NULL CONSTRAINT [DF_CashClosings_CashOut] DEFAULT (0);
IF COL_LENGTH(N'dbo.CashClosings', N'ExpectedBalance') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [ExpectedBalance] MONEY NOT NULL CONSTRAINT [DF_CashClosings_ExpectedBalance] DEFAULT (0);
IF COL_LENGTH(N'dbo.CashClosings', N'CountedAmount') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [CountedAmount] MONEY NOT NULL CONSTRAINT [DF_CashClosings_CountedAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.CashClosings', N'Difference') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [Difference] MONEY NOT NULL CONSTRAINT [DF_CashClosings_Difference] DEFAULT (0);
IF COL_LENGTH(N'dbo.CashClosings', N'Destination') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [Destination] NVARCHAR(10) NOT NULL CONSTRAINT [DF_CashClosings_Destination] DEFAULT (N'MAIN');
IF COL_LENGTH(N'dbo.CashClosings', N'ToCashBoxID') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [ToCashBoxID] INT NULL;
IF COL_LENGTH(N'dbo.CashClosings', N'TransferAmount') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [TransferAmount] MONEY NOT NULL CONSTRAINT [DF_CashClosings_TransferAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.CashClosings', N'KeptAmount') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [KeptAmount] MONEY NOT NULL CONSTRAINT [DF_CashClosings_KeptAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.CashClosings', N'Notes') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.CashClosings', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_CashClosings_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.CashClosings', N'RowVer') IS NULL
    ALTER TABLE [dbo].[CashClosings] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Accounts', N'AccountCode') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [AccountCode] INT NOT NULL;
IF COL_LENGTH(N'dbo.Accounts', N'AccountName') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [AccountName] NVARCHAR(100) NOT NULL;
IF COL_LENGTH(N'dbo.Accounts', N'AccountNameEn') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [AccountNameEn] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.Accounts', N'AccountType') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [AccountType] NVARCHAR(10) NOT NULL;
IF COL_LENGTH(N'dbo.Accounts', N'ParentCode') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [ParentCode] INT NULL;
IF COL_LENGTH(N'dbo.Accounts', N'IsActive') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_Accounts_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.Accounts', N'IsPosting') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [IsPosting] BIT NOT NULL CONSTRAINT [DF_Accounts_IsPosting] DEFAULT (1);
IF COL_LENGTH(N'dbo.Accounts', N'IsSystem') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [IsSystem] BIT NOT NULL CONSTRAINT [DF_Accounts_IsSystem] DEFAULT (0);
IF COL_LENGTH(N'dbo.Accounts', N'AccountLevel') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [AccountLevel] TINYINT NOT NULL CONSTRAINT [DF_Accounts_AccountLevel] DEFAULT (1);
IF COL_LENGTH(N'dbo.Accounts', N'TreeKey') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [TreeKey] NVARCHAR(60) NULL;
IF COL_LENGTH(N'dbo.Accounts', N'Level1Code') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [Level1Code] INT NULL;
IF COL_LENGTH(N'dbo.Accounts', N'Level2Code') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [Level2Code] INT NULL;
IF COL_LENGTH(N'dbo.Accounts', N'Level3Code') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [Level3Code] INT NULL;
IF COL_LENGTH(N'dbo.Accounts', N'Level4Code') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [Level4Code] INT NULL;
IF COL_LENGTH(N'dbo.Accounts', N'Level5Code') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [Level5Code] INT NULL;
IF COL_LENGTH(N'dbo.Accounts', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Accounts] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.JournalSourceTypes', N'SourceType') IS NULL
    ALTER TABLE [dbo].[JournalSourceTypes] ADD [SourceType] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.JournalSourceTypes', N'TypeName') IS NULL
    ALTER TABLE [dbo].[JournalSourceTypes] ADD [TypeName] NVARCHAR(50) NOT NULL;
IF COL_LENGTH(N'dbo.JournalSourceTypes', N'TypeNameEn') IS NULL
    ALTER TABLE [dbo].[JournalSourceTypes] ADD [TypeNameEn] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.JournalSourceTypes', N'SortOrder') IS NULL
    ALTER TABLE [dbo].[JournalSourceTypes] ADD [SortOrder] SMALLINT NOT NULL CONSTRAINT [DF_JournalSourceTypes_SortOrder] DEFAULT (0);
IF COL_LENGTH(N'dbo.JournalSourceTypes', N'RowVer') IS NULL
    ALTER TABLE [dbo].[JournalSourceTypes] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.JournalEntries', N'EntryID') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [EntryID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.JournalEntries', N'EntryNumber') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [EntryNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.JournalEntries', N'EntryDate') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [EntryDate] DATETIME NOT NULL;
IF COL_LENGTH(N'dbo.JournalEntries', N'SourceType') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [SourceType] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.JournalEntries', N'SourceID') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [SourceID] INT NOT NULL;
IF COL_LENGTH(N'dbo.JournalEntries', N'SourceNumber') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [SourceNumber] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.JournalEntries', N'Description') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [Description] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.JournalEntries', N'TotalDebit') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [TotalDebit] MONEY NOT NULL CONSTRAINT [DF_JournalEntries_TotalDebit] DEFAULT (0);
IF COL_LENGTH(N'dbo.JournalEntries', N'TotalCredit') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [TotalCredit] MONEY NOT NULL CONSTRAINT [DF_JournalEntries_TotalCredit] DEFAULT (0);
IF COL_LENGTH(N'dbo.JournalEntries', N'LineCount') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [LineCount] SMALLINT NOT NULL CONSTRAINT [DF_JournalEntries_LineCount] DEFAULT (0);
IF COL_LENGTH(N'dbo.JournalEntries', N'Signature') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [Signature] MONEY NOT NULL CONSTRAINT [DF_JournalEntries_Signature] DEFAULT (0);
IF COL_LENGTH(N'dbo.JournalEntries', N'UpdatedAt') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [UpdatedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.JournalEntries', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_JournalEntries_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.JournalEntries', N'CurrencyCode') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_JournalEntries_CurrencyCode] DEFAULT (N'SAR');
IF COL_LENGTH(N'dbo.JournalEntries', N'ExchangeRate') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_JournalEntries_ExchangeRate] DEFAULT (1);
IF COL_LENGTH(N'dbo.JournalEntries', N'ForeignAmount') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_JournalEntries_ForeignAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.JournalEntries', N'RowVer') IS NULL
    ALTER TABLE [dbo].[JournalEntries] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.JournalLines', N'JournalLineID') IS NULL
    ALTER TABLE [dbo].[JournalLines] ADD [JournalLineID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.JournalLines', N'EntryID') IS NULL
    ALTER TABLE [dbo].[JournalLines] ADD [EntryID] INT NOT NULL;
IF COL_LENGTH(N'dbo.JournalLines', N'LineNumber') IS NULL
    ALTER TABLE [dbo].[JournalLines] ADD [LineNumber] SMALLINT NOT NULL;
IF COL_LENGTH(N'dbo.JournalLines', N'AccountCode') IS NULL
    ALTER TABLE [dbo].[JournalLines] ADD [AccountCode] INT NOT NULL;
IF COL_LENGTH(N'dbo.JournalLines', N'Debit') IS NULL
    ALTER TABLE [dbo].[JournalLines] ADD [Debit] MONEY NOT NULL CONSTRAINT [DF_JournalLines_Debit] DEFAULT (0);
IF COL_LENGTH(N'dbo.JournalLines', N'Credit') IS NULL
    ALTER TABLE [dbo].[JournalLines] ADD [Credit] MONEY NOT NULL CONSTRAINT [DF_JournalLines_Credit] DEFAULT (0);
IF COL_LENGTH(N'dbo.JournalLines', N'LineText') IS NULL
    ALTER TABLE [dbo].[JournalLines] ADD [LineText] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.JournalLines', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[JournalLines] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.JournalLines', N'RowVer') IS NULL
    ALTER TABLE [dbo].[JournalLines] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.PeriodClosings', N'PeriodClosingID') IS NULL
    ALTER TABLE [dbo].[PeriodClosings] ADD [PeriodClosingID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.PeriodClosings', N'ActionType') IS NULL
    ALTER TABLE [dbo].[PeriodClosings] ADD [ActionType] NVARCHAR(12) NOT NULL;
IF COL_LENGTH(N'dbo.PeriodClosings', N'ClosedThrough') IS NULL
    ALTER TABLE [dbo].[PeriodClosings] ADD [ClosedThrough] DATETIME NULL;
IF COL_LENGTH(N'dbo.PeriodClosings', N'PreviousThrough') IS NULL
    ALTER TABLE [dbo].[PeriodClosings] ADD [PreviousThrough] DATETIME NULL;
IF COL_LENGTH(N'dbo.PeriodClosings', N'FiscalYear') IS NULL
    ALTER TABLE [dbo].[PeriodClosings] ADD [FiscalYear] SMALLINT NULL;
IF COL_LENGTH(N'dbo.PeriodClosings', N'Notes') IS NULL
    ALTER TABLE [dbo].[PeriodClosings] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.PeriodClosings', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[PeriodClosings] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PeriodClosings', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[PeriodClosings] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PeriodClosings_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.PeriodClosings', N'RowVer') IS NULL
    ALTER TABLE [dbo].[PeriodClosings] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.FiscalPeriods', N'FiscalPeriodID') IS NULL
    ALTER TABLE [dbo].[FiscalPeriods] ADD [FiscalPeriodID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.FiscalPeriods', N'PeriodName') IS NULL
    ALTER TABLE [dbo].[FiscalPeriods] ADD [PeriodName] NVARCHAR(50) NOT NULL;
IF COL_LENGTH(N'dbo.FiscalPeriods', N'StartDate') IS NULL
    ALTER TABLE [dbo].[FiscalPeriods] ADD [StartDate] DATETIME NOT NULL;
IF COL_LENGTH(N'dbo.FiscalPeriods', N'EndDate') IS NULL
    ALTER TABLE [dbo].[FiscalPeriods] ADD [EndDate] DATETIME NOT NULL;
IF COL_LENGTH(N'dbo.FiscalPeriods', N'Notes') IS NULL
    ALTER TABLE [dbo].[FiscalPeriods] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.FiscalPeriods', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[FiscalPeriods] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_FiscalPeriods_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.FiscalPeriods', N'RowVer') IS NULL
    ALTER TABLE [dbo].[FiscalPeriods] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.FiscalYearClosings', N'YearClosingID') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosings] ADD [YearClosingID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.FiscalYearClosings', N'FiscalYear') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosings] ADD [FiscalYear] SMALLINT NOT NULL;
IF COL_LENGTH(N'dbo.FiscalYearClosings', N'ClosingNumber') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosings] ADD [ClosingNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.FiscalYearClosings', N'ClosingDate') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosings] ADD [ClosingDate] DATETIME NOT NULL CONSTRAINT [DF_FiscalYearClosings_ClosingDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.FiscalYearClosings', N'NetProfit') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosings] ADD [NetProfit] MONEY NOT NULL CONSTRAINT [DF_FiscalYearClosings_NetProfit] DEFAULT (0);
IF COL_LENGTH(N'dbo.FiscalYearClosings', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosings] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.FiscalYearClosings', N'Notes') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosings] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.FiscalYearClosings', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosings] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_FiscalYearClosings_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.FiscalYearClosings', N'RowVer') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosings] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.FiscalYearClosingLines', N'YearClosingLineID') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosingLines] ADD [YearClosingLineID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.FiscalYearClosingLines', N'YearClosingID') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosingLines] ADD [YearClosingID] INT NOT NULL;
IF COL_LENGTH(N'dbo.FiscalYearClosingLines', N'LineNumber') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosingLines] ADD [LineNumber] SMALLINT NOT NULL;
IF COL_LENGTH(N'dbo.FiscalYearClosingLines', N'AccountCode') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosingLines] ADD [AccountCode] INT NOT NULL;
IF COL_LENGTH(N'dbo.FiscalYearClosingLines', N'Debit') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosingLines] ADD [Debit] MONEY NOT NULL CONSTRAINT [DF_FiscalYearClosingLines_Debit] DEFAULT (0);
IF COL_LENGTH(N'dbo.FiscalYearClosingLines', N'Credit') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosingLines] ADD [Credit] MONEY NOT NULL CONSTRAINT [DF_FiscalYearClosingLines_Credit] DEFAULT (0);
IF COL_LENGTH(N'dbo.FiscalYearClosingLines', N'LineText') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosingLines] ADD [LineText] NVARCHAR(150) NULL;
IF COL_LENGTH(N'dbo.FiscalYearClosingLines', N'RowVer') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosingLines] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.VatReturns', N'VatReturnID') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [VatReturnID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.VatReturns', N'ReturnNumber') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [ReturnNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.VatReturns', N'PeriodFrom') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [PeriodFrom] DATETIME NOT NULL;
IF COL_LENGTH(N'dbo.VatReturns', N'PeriodTo') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [PeriodTo] DATETIME NOT NULL;
IF COL_LENGTH(N'dbo.VatReturns', N'Status') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_VatReturns_Status] DEFAULT (N'DRAFT');
IF COL_LENGTH(N'dbo.VatReturns', N'SalesStdAmount') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [SalesStdAmount] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesStdAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'SalesStdAdjust') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [SalesStdAdjust] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesStdAdjust] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'SalesStdVAT') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [SalesStdVAT] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesStdVAT] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'SalesZeroAmount') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [SalesZeroAmount] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesZeroAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'SalesZeroAdjust') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [SalesZeroAdjust] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesZeroAdjust] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'SalesExemptAmount') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [SalesExemptAmount] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesExemptAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'SalesExemptAdjust') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [SalesExemptAdjust] MONEY NOT NULL CONSTRAINT [DF_VatReturns_SalesExemptAdjust] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'PurchStdAmount') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [PurchStdAmount] MONEY NOT NULL CONSTRAINT [DF_VatReturns_PurchStdAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'PurchStdAdjust') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [PurchStdAdjust] MONEY NOT NULL CONSTRAINT [DF_VatReturns_PurchStdAdjust] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'PurchStdVAT') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [PurchStdVAT] MONEY NOT NULL CONSTRAINT [DF_VatReturns_PurchStdVAT] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'PurchZeroAmount') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [PurchZeroAmount] MONEY NOT NULL CONSTRAINT [DF_VatReturns_PurchZeroAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'PurchZeroAdjust') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [PurchZeroAdjust] MONEY NOT NULL CONSTRAINT [DF_VatReturns_PurchZeroAdjust] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'Corrections') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [Corrections] MONEY NOT NULL CONSTRAINT [DF_VatReturns_Corrections] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'CarriedCredit') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [CarriedCredit] MONEY NOT NULL CONSTRAINT [DF_VatReturns_CarriedCredit] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'NetDue') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [NetDue] MONEY NOT NULL CONSTRAINT [DF_VatReturns_NetDue] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'FiledDate') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [FiledDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.VatReturns', N'FilingRef') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [FilingRef] NVARCHAR(30) NULL;
IF COL_LENGTH(N'dbo.VatReturns', N'PaidDate') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [PaidDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.VatReturns', N'PaidAmount') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [PaidAmount] MONEY NOT NULL CONSTRAINT [DF_VatReturns_PaidAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.VatReturns', N'PaidAccount') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [PaidAccount] INT NULL;
IF COL_LENGTH(N'dbo.VatReturns', N'Notes') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.VatReturns', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.VatReturns', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_VatReturns_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.VatReturns', N'RowVer') IS NULL
    ALTER TABLE [dbo].[VatReturns] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.ManualEntries', N'ManualEntryID') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [ManualEntryID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.ManualEntries', N'EntryNumber') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [EntryNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.ManualEntries', N'EntryDate') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [EntryDate] DATETIME NOT NULL CONSTRAINT [DF_ManualEntries_EntryDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.ManualEntries', N'Description') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [Description] NVARCHAR(255) NOT NULL;
IF COL_LENGTH(N'dbo.ManualEntries', N'Reference') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [Reference] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.ManualEntries', N'ReversalOfID') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [ReversalOfID] INT NULL;
IF COL_LENGTH(N'dbo.ManualEntries', N'TotalAmount') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_ManualEntries_TotalAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.ManualEntries', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.ManualEntries', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_ManualEntries_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.ManualEntries', N'UpdatedAt') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [UpdatedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.ManualEntries', N'CurrencyCode') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [CurrencyCode] NVARCHAR(3) NULL CONSTRAINT [DF_ManualEntries_CurrencyCode] DEFAULT (N'SAR');
IF COL_LENGTH(N'dbo.ManualEntries', N'ExchangeRate') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [ExchangeRate] MONEY NOT NULL CONSTRAINT [DF_ManualEntries_ExchangeRate] DEFAULT (1);
IF COL_LENGTH(N'dbo.ManualEntries', N'ForeignAmount') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [ForeignAmount] MONEY NOT NULL CONSTRAINT [DF_ManualEntries_ForeignAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.ManualEntries', N'RowVer') IS NULL
    ALTER TABLE [dbo].[ManualEntries] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.ManualEntryLines', N'ManualLineID') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] ADD [ManualLineID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.ManualEntryLines', N'ManualEntryID') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] ADD [ManualEntryID] INT NOT NULL;
IF COL_LENGTH(N'dbo.ManualEntryLines', N'LineNumber') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] ADD [LineNumber] SMALLINT NOT NULL;
IF COL_LENGTH(N'dbo.ManualEntryLines', N'AccountCode') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] ADD [AccountCode] INT NOT NULL;
IF COL_LENGTH(N'dbo.ManualEntryLines', N'Debit') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] ADD [Debit] MONEY NOT NULL CONSTRAINT [DF_ManualEntryLines_Debit] DEFAULT (0);
IF COL_LENGTH(N'dbo.ManualEntryLines', N'Credit') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] ADD [Credit] MONEY NOT NULL CONSTRAINT [DF_ManualEntryLines_Credit] DEFAULT (0);
IF COL_LENGTH(N'dbo.ManualEntryLines', N'LineText') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] ADD [LineText] NVARCHAR(150) NULL;
IF COL_LENGTH(N'dbo.ManualEntryLines', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.ManualEntryLines', N'ForeignDebit') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] ADD [ForeignDebit] MONEY NOT NULL CONSTRAINT [DF_ManualEntryLines_ForeignDebit] DEFAULT (0);
IF COL_LENGTH(N'dbo.ManualEntryLines', N'ForeignCredit') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] ADD [ForeignCredit] MONEY NOT NULL CONSTRAINT [DF_ManualEntryLines_ForeignCredit] DEFAULT (0);
IF COL_LENGTH(N'dbo.ManualEntryLines', N'RowVer') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.OpeningBalances', N'OpeningLineID') IS NULL
    ALTER TABLE [dbo].[OpeningBalances] ADD [OpeningLineID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.OpeningBalances', N'AccountCode') IS NULL
    ALTER TABLE [dbo].[OpeningBalances] ADD [AccountCode] INT NOT NULL;
IF COL_LENGTH(N'dbo.OpeningBalances', N'Debit') IS NULL
    ALTER TABLE [dbo].[OpeningBalances] ADD [Debit] MONEY NOT NULL CONSTRAINT [DF_OpeningBalances_Debit] DEFAULT (0);
IF COL_LENGTH(N'dbo.OpeningBalances', N'Credit') IS NULL
    ALTER TABLE [dbo].[OpeningBalances] ADD [Credit] MONEY NOT NULL CONSTRAINT [DF_OpeningBalances_Credit] DEFAULT (0);
IF COL_LENGTH(N'dbo.OpeningBalances', N'LineText') IS NULL
    ALTER TABLE [dbo].[OpeningBalances] ADD [LineText] NVARCHAR(150) NULL;
IF COL_LENGTH(N'dbo.OpeningBalances', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[OpeningBalances] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_OpeningBalances_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.OpeningBalances', N'RowVer') IS NULL
    ALTER TABLE [dbo].[OpeningBalances] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.OpeningStock', N'OpeningStockID') IS NULL
    ALTER TABLE [dbo].[OpeningStock] ADD [OpeningStockID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.OpeningStock', N'ProductID') IS NULL
    ALTER TABLE [dbo].[OpeningStock] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.OpeningStock', N'Quantity') IS NULL
    ALTER TABLE [dbo].[OpeningStock] ADD [Quantity] MONEY NOT NULL CONSTRAINT [DF_OpeningStock_Quantity] DEFAULT (0);
IF COL_LENGTH(N'dbo.OpeningStock', N'UnitCost') IS NULL
    ALTER TABLE [dbo].[OpeningStock] ADD [UnitCost] MONEY NOT NULL CONSTRAINT [DF_OpeningStock_UnitCost] DEFAULT (0);
IF COL_LENGTH(N'dbo.OpeningStock', N'TransactionID') IS NULL
    ALTER TABLE [dbo].[OpeningStock] ADD [TransactionID] INT NULL;
IF COL_LENGTH(N'dbo.OpeningStock', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[OpeningStock] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_OpeningStock_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.OpeningStock', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[OpeningStock] ADD [WarehouseID] INT NOT NULL CONSTRAINT [DF_OpeningStock_WarehouseID] DEFAULT (1);
IF COL_LENGTH(N'dbo.OpeningStock', N'RowVer') IS NULL
    ALTER TABLE [dbo].[OpeningStock] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.TransactionTypes', N'TransactionTypeID') IS NULL
    ALTER TABLE [dbo].[TransactionTypes] ADD [TransactionTypeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.TransactionTypes', N'TypeCode') IS NULL
    ALTER TABLE [dbo].[TransactionTypes] ADD [TypeCode] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.TransactionTypes', N'TypeName') IS NULL
    ALTER TABLE [dbo].[TransactionTypes] ADD [TypeName] NVARCHAR(50) NOT NULL;
IF COL_LENGTH(N'dbo.TransactionTypes', N'TypeNameEn') IS NULL
    ALTER TABLE [dbo].[TransactionTypes] ADD [TypeNameEn] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.TransactionTypes', N'Direction') IS NULL
    ALTER TABLE [dbo].[TransactionTypes] ADD [Direction] SMALLINT NOT NULL;
IF COL_LENGTH(N'dbo.TransactionTypes', N'IsManual') IS NULL
    ALTER TABLE [dbo].[TransactionTypes] ADD [IsManual] BIT NOT NULL CONSTRAINT [DF_TransactionTypes_IsManual] DEFAULT (0);
IF COL_LENGTH(N'dbo.TransactionTypes', N'RowVer') IS NULL
    ALTER TABLE [dbo].[TransactionTypes] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.InventoryTransactions', N'TransactionID') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [TransactionID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.InventoryTransactions', N'TransactionDate') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [TransactionDate] DATETIME NOT NULL CONSTRAINT [DF_InventoryTransactions_TransactionDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.InventoryTransactions', N'ProductID') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.InventoryTransactions', N'TransactionTypeID') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [TransactionTypeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.InventoryTransactions', N'Quantity') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [Quantity] MONEY NOT NULL;
IF COL_LENGTH(N'dbo.InventoryTransactions', N'UnitCost') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [UnitCost] MONEY NOT NULL CONSTRAINT [DF_InventoryTransactions_UnitCost] DEFAULT (0);
IF COL_LENGTH(N'dbo.InventoryTransactions', N'QuantityAfter') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [QuantityAfter] MONEY NOT NULL CONSTRAINT [DF_InventoryTransactions_QuantityAfter] DEFAULT (0);
IF COL_LENGTH(N'dbo.InventoryTransactions', N'ReferenceType') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [ReferenceType] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.InventoryTransactions', N'ReferenceID') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [ReferenceID] INT NULL;
IF COL_LENGTH(N'dbo.InventoryTransactions', N'ReferenceNumber') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [ReferenceNumber] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.InventoryTransactions', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [EmployeeID] INT NULL;
IF COL_LENGTH(N'dbo.InventoryTransactions', N'Notes') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.InventoryTransactions', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_InventoryTransactions_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.InventoryTransactions', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [WarehouseID] INT NOT NULL CONSTRAINT [DF_InventoryTransactions_WarehouseID] DEFAULT (1);
IF COL_LENGTH(N'dbo.InventoryTransactions', N'RowVer') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.StockTransfers', N'StockTransferID') IS NULL
    ALTER TABLE [dbo].[StockTransfers] ADD [StockTransferID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.StockTransfers', N'TransferNumber') IS NULL
    ALTER TABLE [dbo].[StockTransfers] ADD [TransferNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.StockTransfers', N'TransferDate') IS NULL
    ALTER TABLE [dbo].[StockTransfers] ADD [TransferDate] DATETIME NOT NULL CONSTRAINT [DF_StockTransfers_TransferDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.StockTransfers', N'FromWarehouseID') IS NULL
    ALTER TABLE [dbo].[StockTransfers] ADD [FromWarehouseID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockTransfers', N'ToWarehouseID') IS NULL
    ALTER TABLE [dbo].[StockTransfers] ADD [ToWarehouseID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockTransfers', N'Status') IS NULL
    ALTER TABLE [dbo].[StockTransfers] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_StockTransfers_Status] DEFAULT (N'DRAFT');
IF COL_LENGTH(N'dbo.StockTransfers', N'InTransit') IS NULL
    ALTER TABLE [dbo].[StockTransfers] ADD [InTransit] BIT NOT NULL CONSTRAINT [DF_StockTransfers_InTransit] DEFAULT (0);
IF COL_LENGTH(N'dbo.StockTransfers', N'ReceivedAt') IS NULL
    ALTER TABLE [dbo].[StockTransfers] ADD [ReceivedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.StockTransfers', N'ReceivedByID') IS NULL
    ALTER TABLE [dbo].[StockTransfers] ADD [ReceivedByID] INT NULL;
IF COL_LENGTH(N'dbo.StockTransfers', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[StockTransfers] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockTransfers', N'Notes') IS NULL
    ALTER TABLE [dbo].[StockTransfers] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.StockTransfers', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[StockTransfers] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockTransfers_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.StockTransfers', N'RowVer') IS NULL
    ALTER TABLE [dbo].[StockTransfers] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.StockTransferLines', N'StockTransferLineID') IS NULL
    ALTER TABLE [dbo].[StockTransferLines] ADD [StockTransferLineID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.StockTransferLines', N'StockTransferID') IS NULL
    ALTER TABLE [dbo].[StockTransferLines] ADD [StockTransferID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockTransferLines', N'ProductID') IS NULL
    ALTER TABLE [dbo].[StockTransferLines] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockTransferLines', N'Quantity') IS NULL
    ALTER TABLE [dbo].[StockTransferLines] ADD [Quantity] MONEY NOT NULL;
IF COL_LENGTH(N'dbo.StockTransferLines', N'UnitCost') IS NULL
    ALTER TABLE [dbo].[StockTransferLines] ADD [UnitCost] MONEY NOT NULL CONSTRAINT [DF_StockTransferLines_UnitCost] DEFAULT (0);
IF COL_LENGTH(N'dbo.StockTransferLines', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[StockTransferLines] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockTransferLines_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.StockTransferLines', N'RowVer') IS NULL
    ALTER TABLE [dbo].[StockTransferLines] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.StockParties', N'StockPartyID') IS NULL
    ALTER TABLE [dbo].[StockParties] ADD [StockPartyID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.StockParties', N'PartyCode') IS NULL
    ALTER TABLE [dbo].[StockParties] ADD [PartyCode] NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.StockParties', N'PartyName') IS NULL
    ALTER TABLE [dbo].[StockParties] ADD [PartyName] NVARCHAR(50) NOT NULL;
IF COL_LENGTH(N'dbo.StockParties', N'PartyNameEn') IS NULL
    ALTER TABLE [dbo].[StockParties] ADD [PartyNameEn] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.StockParties', N'PartyType') IS NULL
    ALTER TABLE [dbo].[StockParties] ADD [PartyType] NVARCHAR(12) NOT NULL CONSTRAINT [DF_StockParties_PartyType] DEFAULT (N'DEPARTMENT');
IF COL_LENGTH(N'dbo.StockParties', N'AccountCode') IS NULL
    ALTER TABLE [dbo].[StockParties] ADD [AccountCode] INT NULL;
IF COL_LENGTH(N'dbo.StockParties', N'CostCenterID') IS NULL
    ALTER TABLE [dbo].[StockParties] ADD [CostCenterID] INT NULL;
IF COL_LENGTH(N'dbo.StockParties', N'IsActive') IS NULL
    ALTER TABLE [dbo].[StockParties] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_StockParties_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.StockParties', N'Notes') IS NULL
    ALTER TABLE [dbo].[StockParties] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.StockParties', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[StockParties] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockParties_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.StockParties', N'RowVer') IS NULL
    ALTER TABLE [dbo].[StockParties] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.StockDocuments', N'StockDocumentID') IS NULL
    ALTER TABLE [dbo].[StockDocuments] ADD [StockDocumentID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.StockDocuments', N'DocNumber') IS NULL
    ALTER TABLE [dbo].[StockDocuments] ADD [DocNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.StockDocuments', N'DocType') IS NULL
    ALTER TABLE [dbo].[StockDocuments] ADD [DocType] NVARCHAR(10) NOT NULL;
IF COL_LENGTH(N'dbo.StockDocuments', N'DocDate') IS NULL
    ALTER TABLE [dbo].[StockDocuments] ADD [DocDate] DATETIME NOT NULL CONSTRAINT [DF_StockDocuments_DocDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.StockDocuments', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[StockDocuments] ADD [WarehouseID] INT NOT NULL CONSTRAINT [DF_StockDocuments_WarehouseID] DEFAULT (1);
IF COL_LENGTH(N'dbo.StockDocuments', N'StockPartyID') IS NULL
    ALTER TABLE [dbo].[StockDocuments] ADD [StockPartyID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockDocuments', N'Status') IS NULL
    ALTER TABLE [dbo].[StockDocuments] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_StockDocuments_Status] DEFAULT (N'DRAFT');
IF COL_LENGTH(N'dbo.StockDocuments', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[StockDocuments] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockDocuments', N'Notes') IS NULL
    ALTER TABLE [dbo].[StockDocuments] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.StockDocuments', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[StockDocuments] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockDocuments_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.StockDocuments', N'RowVer') IS NULL
    ALTER TABLE [dbo].[StockDocuments] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.StockDocumentLines', N'StockDocumentLineID') IS NULL
    ALTER TABLE [dbo].[StockDocumentLines] ADD [StockDocumentLineID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.StockDocumentLines', N'StockDocumentID') IS NULL
    ALTER TABLE [dbo].[StockDocumentLines] ADD [StockDocumentID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockDocumentLines', N'ProductID') IS NULL
    ALTER TABLE [dbo].[StockDocumentLines] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockDocumentLines', N'Quantity') IS NULL
    ALTER TABLE [dbo].[StockDocumentLines] ADD [Quantity] MONEY NOT NULL;
IF COL_LENGTH(N'dbo.StockDocumentLines', N'UnitCost') IS NULL
    ALTER TABLE [dbo].[StockDocumentLines] ADD [UnitCost] MONEY NOT NULL CONSTRAINT [DF_StockDocumentLines_UnitCost] DEFAULT (0);
IF COL_LENGTH(N'dbo.StockDocumentLines', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[StockDocumentLines] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockDocumentLines_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.StockDocumentLines', N'RowVer') IS NULL
    ALTER TABLE [dbo].[StockDocumentLines] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.StockNotes', N'StockNoteID') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [StockNoteID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.StockNotes', N'NoteNumber') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [NoteNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.StockNotes', N'NoteType') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [NoteType] NVARCHAR(10) NOT NULL;
IF COL_LENGTH(N'dbo.StockNotes', N'NoteDate') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [NoteDate] DATETIME NOT NULL CONSTRAINT [DF_StockNotes_NoteDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.StockNotes', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [WarehouseID] INT NOT NULL CONSTRAINT [DF_StockNotes_WarehouseID] DEFAULT (1);
IF COL_LENGTH(N'dbo.StockNotes', N'CustomerID') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [CustomerID] INT NULL;
IF COL_LENGTH(N'dbo.StockNotes', N'SupplierID') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [SupplierID] INT NULL;
IF COL_LENGTH(N'dbo.StockNotes', N'SalesOrderID') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [SalesOrderID] INT NULL;
IF COL_LENGTH(N'dbo.StockNotes', N'PurchaseOrderID') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [PurchaseOrderID] INT NULL;
IF COL_LENGTH(N'dbo.StockNotes', N'Status') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_StockNotes_Status] DEFAULT (N'DRAFT');
IF COL_LENGTH(N'dbo.StockNotes', N'SalesInvoiceID') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [SalesInvoiceID] INT NULL;
IF COL_LENGTH(N'dbo.StockNotes', N'PurchaseInvoiceID') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [PurchaseInvoiceID] INT NULL;
IF COL_LENGTH(N'dbo.StockNotes', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockNotes', N'Notes') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.StockNotes', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockNotes_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.StockNotes', N'RowVer') IS NULL
    ALTER TABLE [dbo].[StockNotes] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.StockNoteLines', N'StockNoteLineID') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] ADD [StockNoteLineID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.StockNoteLines', N'StockNoteID') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] ADD [StockNoteID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockNoteLines', N'ProductID') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockNoteLines', N'Quantity') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] ADD [Quantity] MONEY NOT NULL;
IF COL_LENGTH(N'dbo.StockNoteLines', N'UnitPrice') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] ADD [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_StockNoteLines_UnitPrice] DEFAULT (0);
IF COL_LENGTH(N'dbo.StockNoteLines', N'UnitCost') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] ADD [UnitCost] MONEY NOT NULL CONSTRAINT [DF_StockNoteLines_UnitCost] DEFAULT (0);
IF COL_LENGTH(N'dbo.StockNoteLines', N'SalesOrderLineID') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] ADD [SalesOrderLineID] INT NULL;
IF COL_LENGTH(N'dbo.StockNoteLines', N'PurchaseOrderLineID') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] ADD [PurchaseOrderLineID] INT NULL;
IF COL_LENGTH(N'dbo.StockNoteLines', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_StockNoteLines_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.StockNoteLines', N'RowVer') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.SalesOrders', N'SalesOrderID') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [SalesOrderID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.SalesOrders', N'OrderNumber') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [OrderNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.SalesOrders', N'OrderDate') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [OrderDate] DATETIME NOT NULL CONSTRAINT [DF_SalesOrders_OrderDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.SalesOrders', N'CustomerID') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [CustomerID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesOrders', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [WarehouseID] INT NOT NULL CONSTRAINT [DF_SalesOrders_WarehouseID] DEFAULT (1);
IF COL_LENGTH(N'dbo.SalesOrders', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [SalesRepID] INT NULL;
IF COL_LENGTH(N'dbo.SalesOrders', N'RequiredDate') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [RequiredDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.SalesOrders', N'Status') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_SalesOrders_Status] DEFAULT (N'DRAFT');
IF COL_LENGTH(N'dbo.SalesOrders', N'ApprovedAt') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [ApprovedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.SalesOrders', N'PreparedAt') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [PreparedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.SalesOrders', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesOrders', N'Notes') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.SalesOrders', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SalesOrders_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.SalesOrders', N'RowVer') IS NULL
    ALTER TABLE [dbo].[SalesOrders] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.Quotations', N'QuotationID') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [QuotationID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.Quotations', N'QuotationNumber') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [QuotationNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.Quotations', N'QuoteDate') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [QuoteDate] DATETIME NOT NULL CONSTRAINT [DF_Quotations_QuoteDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.Quotations', N'CustomerID') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [CustomerID] INT NOT NULL;
IF COL_LENGTH(N'dbo.Quotations', N'ValidUntil') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [ValidUntil] DATETIME NOT NULL CONSTRAINT [DF_Quotations_ValidUntil] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.Quotations', N'Status') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_Quotations_Status] DEFAULT (N'DRAFT');
IF COL_LENGTH(N'dbo.Quotations', N'IssuedAt') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [IssuedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.Quotations', N'SubTotal') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [SubTotal] MONEY NOT NULL CONSTRAINT [DF_Quotations_SubTotal] DEFAULT (0);
IF COL_LENGTH(N'dbo.Quotations', N'Tax') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [Tax] MONEY NOT NULL CONSTRAINT [DF_Quotations_Tax] DEFAULT (0);
IF COL_LENGTH(N'dbo.Quotations', N'TotalAmount') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [TotalAmount] MONEY NOT NULL CONSTRAINT [DF_Quotations_TotalAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.Quotations', N'SalesOrderID') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [SalesOrderID] INT NULL;
IF COL_LENGTH(N'dbo.Quotations', N'SalesInvoiceID') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [SalesInvoiceID] INT NULL;
IF COL_LENGTH(N'dbo.Quotations', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.Quotations', N'Terms') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [Terms] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Quotations', N'Notes') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.Quotations', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_Quotations_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.Quotations', N'RowVer') IS NULL
    ALTER TABLE [dbo].[Quotations] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.QuotationLines', N'QuotationLineID') IS NULL
    ALTER TABLE [dbo].[QuotationLines] ADD [QuotationLineID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.QuotationLines', N'QuotationID') IS NULL
    ALTER TABLE [dbo].[QuotationLines] ADD [QuotationID] INT NOT NULL;
IF COL_LENGTH(N'dbo.QuotationLines', N'ProductID') IS NULL
    ALTER TABLE [dbo].[QuotationLines] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.QuotationLines', N'Quantity') IS NULL
    ALTER TABLE [dbo].[QuotationLines] ADD [Quantity] MONEY NOT NULL;
IF COL_LENGTH(N'dbo.QuotationLines', N'UnitPrice') IS NULL
    ALTER TABLE [dbo].[QuotationLines] ADD [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_QuotationLines_UnitPrice] DEFAULT (0);
IF COL_LENGTH(N'dbo.QuotationLines', N'NetAmount') IS NULL
    ALTER TABLE [dbo].[QuotationLines] ADD [NetAmount] MONEY NOT NULL CONSTRAINT [DF_QuotationLines_NetAmount] DEFAULT (0);
IF COL_LENGTH(N'dbo.QuotationLines', N'Tax') IS NULL
    ALTER TABLE [dbo].[QuotationLines] ADD [Tax] MONEY NOT NULL CONSTRAINT [DF_QuotationLines_Tax] DEFAULT (0);
IF COL_LENGTH(N'dbo.QuotationLines', N'LineTotal') IS NULL
    ALTER TABLE [dbo].[QuotationLines] ADD [LineTotal] MONEY NOT NULL CONSTRAINT [DF_QuotationLines_LineTotal] DEFAULT (0);
IF COL_LENGTH(N'dbo.QuotationLines', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[QuotationLines] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_QuotationLines_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.QuotationLines', N'RowVer') IS NULL
    ALTER TABLE [dbo].[QuotationLines] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseOrders', N'PurchaseOrderID') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] ADD [PurchaseOrderID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseOrders', N'OrderNumber') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] ADD [OrderNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseOrders', N'OrderDate') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] ADD [OrderDate] DATETIME NOT NULL CONSTRAINT [DF_PurchaseOrders_OrderDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.PurchaseOrders', N'SupplierID') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] ADD [SupplierID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseOrders', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] ADD [WarehouseID] INT NOT NULL CONSTRAINT [DF_PurchaseOrders_WarehouseID] DEFAULT (1);
IF COL_LENGTH(N'dbo.PurchaseOrders', N'ExpectedDate') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] ADD [ExpectedDate] DATETIME NULL;
IF COL_LENGTH(N'dbo.PurchaseOrders', N'Status') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_PurchaseOrders_Status] DEFAULT (N'DRAFT');
IF COL_LENGTH(N'dbo.PurchaseOrders', N'ApprovedAt') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] ADD [ApprovedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.PurchaseOrders', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseOrders', N'Notes') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.PurchaseOrders', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PurchaseOrders_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.PurchaseOrders', N'RowVer') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseOrderLines', N'PurchaseOrderLineID') IS NULL
    ALTER TABLE [dbo].[PurchaseOrderLines] ADD [PurchaseOrderLineID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseOrderLines', N'PurchaseOrderID') IS NULL
    ALTER TABLE [dbo].[PurchaseOrderLines] ADD [PurchaseOrderID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseOrderLines', N'ProductID') IS NULL
    ALTER TABLE [dbo].[PurchaseOrderLines] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseOrderLines', N'Quantity') IS NULL
    ALTER TABLE [dbo].[PurchaseOrderLines] ADD [Quantity] MONEY NOT NULL;
IF COL_LENGTH(N'dbo.PurchaseOrderLines', N'UnitCost') IS NULL
    ALTER TABLE [dbo].[PurchaseOrderLines] ADD [UnitCost] MONEY NOT NULL CONSTRAINT [DF_PurchaseOrderLines_UnitCost] DEFAULT (0);
IF COL_LENGTH(N'dbo.PurchaseOrderLines', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[PurchaseOrderLines] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_PurchaseOrderLines_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.PurchaseOrderLines', N'RowVer') IS NULL
    ALTER TABLE [dbo].[PurchaseOrderLines] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.RepDays', N'RepDayID') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [RepDayID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.RepDays', N'DayNumber') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [DayNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.RepDays', N'DayDate') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [DayDate] DATETIME NOT NULL CONSTRAINT [DF_RepDays_DayDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.RepDays', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [SalesRepID] INT NOT NULL;
IF COL_LENGTH(N'dbo.RepDays', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [WarehouseID] INT NOT NULL;
IF COL_LENGTH(N'dbo.RepDays', N'CashBoxID') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [CashBoxID] INT NOT NULL;
IF COL_LENGTH(N'dbo.RepDays', N'StoreID') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [StoreID] INT NOT NULL CONSTRAINT [DF_RepDays_StoreID] DEFAULT (1);
IF COL_LENGTH(N'dbo.RepDays', N'Status') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_RepDays_Status] DEFAULT (N'OPEN');
IF COL_LENGTH(N'dbo.RepDays', N'OpenedAt') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [OpenedAt] DATETIME NOT NULL CONSTRAINT [DF_RepDays_OpenedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.RepDays', N'SettledAt') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [SettledAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.RepDays', N'StockCountID') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [StockCountID] INT NULL;
IF COL_LENGTH(N'dbo.RepDays', N'ClosingID') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [ClosingID] INT NULL;
IF COL_LENGTH(N'dbo.RepDays', N'GoodsShortage') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [GoodsShortage] MONEY NOT NULL CONSTRAINT [DF_RepDays_GoodsShortage] DEFAULT (0);
IF COL_LENGTH(N'dbo.RepDays', N'AdvanceEmployeeID') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [AdvanceEmployeeID] INT NULL;
IF COL_LENGTH(N'dbo.RepDays', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.RepDays', N'Notes') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.RepDays', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_RepDays_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.RepDays', N'RowVer') IS NULL
    ALTER TABLE [dbo].[RepDays] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.RouteStops', N'RouteStopID') IS NULL
    ALTER TABLE [dbo].[RouteStops] ADD [RouteStopID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.RouteStops', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[RouteStops] ADD [SalesRepID] INT NOT NULL;
IF COL_LENGTH(N'dbo.RouteStops', N'CustomerID') IS NULL
    ALTER TABLE [dbo].[RouteStops] ADD [CustomerID] INT NOT NULL;
IF COL_LENGTH(N'dbo.RouteStops', N'VisitWeekday') IS NULL
    ALTER TABLE [dbo].[RouteStops] ADD [VisitWeekday] TINYINT NOT NULL CONSTRAINT [DF_RouteStops_VisitWeekday] DEFAULT (1);
IF COL_LENGTH(N'dbo.RouteStops', N'StopOrder') IS NULL
    ALTER TABLE [dbo].[RouteStops] ADD [StopOrder] SMALLINT NOT NULL CONSTRAINT [DF_RouteStops_StopOrder] DEFAULT (0);
IF COL_LENGTH(N'dbo.RouteStops', N'IsActive') IS NULL
    ALTER TABLE [dbo].[RouteStops] ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_RouteStops_IsActive] DEFAULT (1);
IF COL_LENGTH(N'dbo.RouteStops', N'Notes') IS NULL
    ALTER TABLE [dbo].[RouteStops] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.RouteStops', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[RouteStops] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_RouteStops_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.RouteStops', N'RowVer') IS NULL
    ALTER TABLE [dbo].[RouteStops] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.RepVisits', N'RepVisitID') IS NULL
    ALTER TABLE [dbo].[RepVisits] ADD [RepVisitID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.RepVisits', N'VisitDate') IS NULL
    ALTER TABLE [dbo].[RepVisits] ADD [VisitDate] DATETIME NOT NULL CONSTRAINT [DF_RepVisits_VisitDate] DEFAULT (CONVERT(DATETIME, CONVERT(DATE, GETDATE())));
IF COL_LENGTH(N'dbo.RepVisits', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[RepVisits] ADD [SalesRepID] INT NOT NULL;
IF COL_LENGTH(N'dbo.RepVisits', N'CustomerID') IS NULL
    ALTER TABLE [dbo].[RepVisits] ADD [CustomerID] INT NOT NULL;
IF COL_LENGTH(N'dbo.RepVisits', N'IsPlanned') IS NULL
    ALTER TABLE [dbo].[RepVisits] ADD [IsPlanned] BIT NOT NULL CONSTRAINT [DF_RepVisits_IsPlanned] DEFAULT (0);
IF COL_LENGTH(N'dbo.RepVisits', N'StopOrder') IS NULL
    ALTER TABLE [dbo].[RepVisits] ADD [StopOrder] SMALLINT NOT NULL CONSTRAINT [DF_RepVisits_StopOrder] DEFAULT (0);
IF COL_LENGTH(N'dbo.RepVisits', N'Result') IS NULL
    ALTER TABLE [dbo].[RepVisits] ADD [Result] NVARCHAR(10) NOT NULL CONSTRAINT [DF_RepVisits_Result] DEFAULT (N'PENDING');
IF COL_LENGTH(N'dbo.RepVisits', N'VisitedAt') IS NULL
    ALTER TABLE [dbo].[RepVisits] ADD [VisitedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.RepVisits', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[RepVisits] ADD [EmployeeID] INT NULL;
IF COL_LENGTH(N'dbo.RepVisits', N'Notes') IS NULL
    ALTER TABLE [dbo].[RepVisits] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.RepVisits', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[RepVisits] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_RepVisits_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.RepVisits', N'RowVer') IS NULL
    ALTER TABLE [dbo].[RepVisits] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.RepPrices', N'RepPriceID') IS NULL
    ALTER TABLE [dbo].[RepPrices] ADD [RepPriceID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.RepPrices', N'SalesRepID') IS NULL
    ALTER TABLE [dbo].[RepPrices] ADD [SalesRepID] INT NOT NULL;
IF COL_LENGTH(N'dbo.RepPrices', N'ProductID') IS NULL
    ALTER TABLE [dbo].[RepPrices] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.RepPrices', N'UnitPrice') IS NULL
    ALTER TABLE [dbo].[RepPrices] ADD [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_RepPrices_UnitPrice] DEFAULT (0);
IF COL_LENGTH(N'dbo.RepPrices', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[RepPrices] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_RepPrices_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.RepPrices', N'RowVer') IS NULL
    ALTER TABLE [dbo].[RepPrices] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.SalesOrderLines', N'SalesOrderLineID') IS NULL
    ALTER TABLE [dbo].[SalesOrderLines] ADD [SalesOrderLineID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.SalesOrderLines', N'SalesOrderID') IS NULL
    ALTER TABLE [dbo].[SalesOrderLines] ADD [SalesOrderID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesOrderLines', N'ProductID') IS NULL
    ALTER TABLE [dbo].[SalesOrderLines] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.SalesOrderLines', N'Quantity') IS NULL
    ALTER TABLE [dbo].[SalesOrderLines] ADD [Quantity] MONEY NOT NULL;
IF COL_LENGTH(N'dbo.SalesOrderLines', N'UnitPrice') IS NULL
    ALTER TABLE [dbo].[SalesOrderLines] ADD [UnitPrice] MONEY NOT NULL CONSTRAINT [DF_SalesOrderLines_UnitPrice] DEFAULT (0);
IF COL_LENGTH(N'dbo.SalesOrderLines', N'CreatedAt') IS NULL
    ALTER TABLE [dbo].[SalesOrderLines] ADD [CreatedAt] DATETIME NOT NULL CONSTRAINT [DF_SalesOrderLines_CreatedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.SalesOrderLines', N'RowVer') IS NULL
    ALTER TABLE [dbo].[SalesOrderLines] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.StockCounts', N'StockCountID') IS NULL
    ALTER TABLE [dbo].[StockCounts] ADD [StockCountID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.StockCounts', N'CountNumber') IS NULL
    ALTER TABLE [dbo].[StockCounts] ADD [CountNumber] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.StockCounts', N'CountDate') IS NULL
    ALTER TABLE [dbo].[StockCounts] ADD [CountDate] DATETIME NOT NULL CONSTRAINT [DF_StockCounts_CountDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.StockCounts', N'CategoryID') IS NULL
    ALTER TABLE [dbo].[StockCounts] ADD [CategoryID] INT NULL;
IF COL_LENGTH(N'dbo.StockCounts', N'Status') IS NULL
    ALTER TABLE [dbo].[StockCounts] ADD [Status] NVARCHAR(10) NOT NULL CONSTRAINT [DF_StockCounts_Status] DEFAULT (N'OPEN');
IF COL_LENGTH(N'dbo.StockCounts', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[StockCounts] ADD [EmployeeID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockCounts', N'PostedAt') IS NULL
    ALTER TABLE [dbo].[StockCounts] ADD [PostedAt] DATETIME NULL;
IF COL_LENGTH(N'dbo.StockCounts', N'PostedByID') IS NULL
    ALTER TABLE [dbo].[StockCounts] ADD [PostedByID] INT NULL;
IF COL_LENGTH(N'dbo.StockCounts', N'Notes') IS NULL
    ALTER TABLE [dbo].[StockCounts] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.StockCounts', N'WarehouseID') IS NULL
    ALTER TABLE [dbo].[StockCounts] ADD [WarehouseID] INT NOT NULL CONSTRAINT [DF_StockCounts_WarehouseID] DEFAULT (1);
IF COL_LENGTH(N'dbo.StockCounts', N'RowVer') IS NULL
    ALTER TABLE [dbo].[StockCounts] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.StockCountDetails', N'StockCountDetailID') IS NULL
    ALTER TABLE [dbo].[StockCountDetails] ADD [StockCountDetailID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.StockCountDetails', N'StockCountID') IS NULL
    ALTER TABLE [dbo].[StockCountDetails] ADD [StockCountID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockCountDetails', N'ProductID') IS NULL
    ALTER TABLE [dbo].[StockCountDetails] ADD [ProductID] INT NOT NULL;
IF COL_LENGTH(N'dbo.StockCountDetails', N'SystemQuantity') IS NULL
    ALTER TABLE [dbo].[StockCountDetails] ADD [SystemQuantity] MONEY NOT NULL CONSTRAINT [DF_StockCountDetails_SystemQuantity] DEFAULT (0);
IF COL_LENGTH(N'dbo.StockCountDetails', N'ActualQuantity') IS NULL
    ALTER TABLE [dbo].[StockCountDetails] ADD [ActualQuantity] MONEY NULL;
IF COL_LENGTH(N'dbo.StockCountDetails', N'Difference') IS NULL
    ALTER TABLE [dbo].[StockCountDetails] ADD [Difference] MONEY NOT NULL CONSTRAINT [DF_StockCountDetails_Difference] DEFAULT (0);
IF COL_LENGTH(N'dbo.StockCountDetails', N'UnitCost') IS NULL
    ALTER TABLE [dbo].[StockCountDetails] ADD [UnitCost] MONEY NOT NULL CONSTRAINT [DF_StockCountDetails_UnitCost] DEFAULT (0);
IF COL_LENGTH(N'dbo.StockCountDetails', N'DifferenceValue') IS NULL
    ALTER TABLE [dbo].[StockCountDetails] ADD [DifferenceValue] MONEY NOT NULL CONSTRAINT [DF_StockCountDetails_DifferenceValue] DEFAULT (0);
IF COL_LENGTH(N'dbo.StockCountDetails', N'Notes') IS NULL
    ALTER TABLE [dbo].[StockCountDetails] ADD [Notes] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.StockCountDetails', N'RowVer') IS NULL
    ALTER TABLE [dbo].[StockCountDetails] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.AuditLog', N'LogID') IS NULL
    ALTER TABLE [dbo].[AuditLog] ADD [LogID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.AuditLog', N'LogDate') IS NULL
    ALTER TABLE [dbo].[AuditLog] ADD [LogDate] DATETIME NOT NULL CONSTRAINT [DF_AuditLog_LogDate] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.AuditLog', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[AuditLog] ADD [EmployeeID] INT NULL;
IF COL_LENGTH(N'dbo.AuditLog', N'ActionType') IS NULL
    ALTER TABLE [dbo].[AuditLog] ADD [ActionType] NVARCHAR(30) NOT NULL;
IF COL_LENGTH(N'dbo.AuditLog', N'ObjectName') IS NULL
    ALTER TABLE [dbo].[AuditLog] ADD [ObjectName] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.AuditLog', N'RecordID') IS NULL
    ALTER TABLE [dbo].[AuditLog] ADD [RecordID] NVARCHAR(30) NULL;
IF COL_LENGTH(N'dbo.AuditLog', N'Details') IS NULL
    ALTER TABLE [dbo].[AuditLog] ADD [Details] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.AuditLog', N'ComputerName') IS NULL
    ALTER TABLE [dbo].[AuditLog] ADD [ComputerName] NVARCHAR(50) NULL;
IF COL_LENGTH(N'dbo.AuditLog', N'RecordLabel') IS NULL
    ALTER TABLE [dbo].[AuditLog] ADD [RecordLabel] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.AuditLog', N'RowVer') IS NULL
    ALTER TABLE [dbo].[AuditLog] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.AuditChanges', N'ChangeID') IS NULL
    ALTER TABLE [dbo].[AuditChanges] ADD [ChangeID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.AuditChanges', N'LogID') IS NULL
    ALTER TABLE [dbo].[AuditChanges] ADD [LogID] INT NOT NULL;
IF COL_LENGTH(N'dbo.AuditChanges', N'LineNo') IS NULL
    ALTER TABLE [dbo].[AuditChanges] ADD [LineNo] SMALLINT NOT NULL CONSTRAINT [DF_AuditChanges_LineNo] DEFAULT (0);
IF COL_LENGTH(N'dbo.AuditChanges', N'FieldName') IS NULL
    ALTER TABLE [dbo].[AuditChanges] ADD [FieldName] NVARCHAR(64) NOT NULL;
IF COL_LENGTH(N'dbo.AuditChanges', N'FieldCaption') IS NULL
    ALTER TABLE [dbo].[AuditChanges] ADD [FieldCaption] NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.AuditChanges', N'OldValue') IS NULL
    ALTER TABLE [dbo].[AuditChanges] ADD [OldValue] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.AuditChanges', N'NewValue') IS NULL
    ALTER TABLE [dbo].[AuditChanges] ADD [NewValue] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.AuditChanges', N'RowVer') IS NULL
    ALTER TABLE [dbo].[AuditChanges] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.LabelSettings', N'LabelSettingID') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [LabelSettingID] INT NOT NULL CONSTRAINT [DF_LabelSettings_LabelSettingID] DEFAULT (1);
IF COL_LENGTH(N'dbo.LabelSettings', N'PrinterName') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [PrinterName] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.LabelSettings', N'LabelWidth') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [LabelWidth] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_LabelWidth] DEFAULT (38);
IF COL_LENGTH(N'dbo.LabelSettings', N'LabelHeight') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [LabelHeight] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_LabelHeight] DEFAULT (25);
IF COL_LENGTH(N'dbo.LabelSettings', N'LabelsAcross') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [LabelsAcross] TINYINT NOT NULL CONSTRAINT [DF_LabelSettings_LabelsAcross] DEFAULT (1);
IF COL_LENGTH(N'dbo.LabelSettings', N'ColumnGap') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [ColumnGap] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_ColumnGap] DEFAULT (2);
IF COL_LENGTH(N'dbo.LabelSettings', N'RowGap') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [RowGap] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_RowGap] DEFAULT (0);
IF COL_LENGTH(N'dbo.LabelSettings', N'MarginTop') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [MarginTop] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_MarginTop] DEFAULT (0);
IF COL_LENGTH(N'dbo.LabelSettings', N'MarginBottom') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [MarginBottom] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_MarginBottom] DEFAULT (0);
IF COL_LENGTH(N'dbo.LabelSettings', N'MarginLeft') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [MarginLeft] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_MarginLeft] DEFAULT (0);
IF COL_LENGTH(N'dbo.LabelSettings', N'MarginRight') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [MarginRight] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_MarginRight] DEFAULT (0);
IF COL_LENGTH(N'dbo.LabelSettings', N'BarHeight') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [BarHeight] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_BarHeight] DEFAULT (10);
IF COL_LENGTH(N'dbo.LabelSettings', N'BarWidth') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [BarWidth] MONEY NOT NULL CONSTRAINT [DF_LabelSettings_BarWidth] DEFAULT (0.25);
IF COL_LENGTH(N'dbo.LabelSettings', N'TopLine1') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [TopLine1] NVARCHAR(10) NOT NULL CONSTRAINT [DF_LabelSettings_TopLine1] DEFAULT (N'STORE');
IF COL_LENGTH(N'dbo.LabelSettings', N'TopLine2') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [TopLine2] NVARCHAR(10) NOT NULL CONSTRAINT [DF_LabelSettings_TopLine2] DEFAULT (N'NAME');
IF COL_LENGTH(N'dbo.LabelSettings', N'BottomLine1') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [BottomLine1] NVARCHAR(10) NOT NULL CONSTRAINT [DF_LabelSettings_BottomLine1] DEFAULT (N'BARCODE');
IF COL_LENGTH(N'dbo.LabelSettings', N'BottomLine2') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [BottomLine2] NVARCHAR(10) NOT NULL CONSTRAINT [DF_LabelSettings_BottomLine2] DEFAULT (N'PRICE');
IF COL_LENGTH(N'dbo.LabelSettings', N'ShortName') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [ShortName] NVARCHAR(30) NULL;
IF COL_LENGTH(N'dbo.LabelSettings', N'FontSize') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [FontSize] TINYINT NOT NULL CONSTRAINT [DF_LabelSettings_FontSize] DEFAULT (7);
IF COL_LENGTH(N'dbo.LabelSettings', N'RowVer') IS NULL
    ALTER TABLE [dbo].[LabelSettings] ADD [RowVer] ROWVERSION NOT NULL;
IF COL_LENGTH(N'dbo.EInvoiceLog', N'LogID') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [LogID] INT IDENTITY(1,1) NOT NULL;
IF COL_LENGTH(N'dbo.EInvoiceLog', N'LoggedAt') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [LoggedAt] DATETIME NOT NULL CONSTRAINT [DF_EInvoiceLog_LoggedAt] DEFAULT (GETDATE());
IF COL_LENGTH(N'dbo.EInvoiceLog', N'Country') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [Country] NVARCHAR(2) NOT NULL;
IF COL_LENGTH(N'dbo.EInvoiceLog', N'DocKind') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [DocKind] NVARCHAR(10) NULL;
IF COL_LENGTH(N'dbo.EInvoiceLog', N'DocID') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [DocID] INT NULL;
IF COL_LENGTH(N'dbo.EInvoiceLog', N'DocNumber') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [DocNumber] NVARCHAR(30) NULL;
IF COL_LENGTH(N'dbo.EInvoiceLog', N'Action') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [Action] NVARCHAR(20) NOT NULL;
IF COL_LENGTH(N'dbo.EInvoiceLog', N'Endpoint') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [Endpoint] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.EInvoiceLog', N'HttpStatus') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [HttpStatus] SMALLINT NULL CONSTRAINT [DF_EInvoiceLog_HttpStatus] DEFAULT (0);
IF COL_LENGTH(N'dbo.EInvoiceLog', N'Result') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [Result] NVARCHAR(12) NOT NULL;
IF COL_LENGTH(N'dbo.EInvoiceLog', N'Message') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [Message] NVARCHAR(255) NULL;
IF COL_LENGTH(N'dbo.EInvoiceLog', N'RequestBody') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [RequestBody] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.EInvoiceLog', N'ResponseBody') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [ResponseBody] NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.EInvoiceLog', N'DurationMs') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [DurationMs] INT NULL CONSTRAINT [DF_EInvoiceLog_DurationMs] DEFAULT (0);
IF COL_LENGTH(N'dbo.EInvoiceLog', N'EmployeeID') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [EmployeeID] INT NULL;
IF COL_LENGTH(N'dbo.EInvoiceLog', N'RowVer') IS NULL
    ALTER TABLE [dbo].[EInvoiceLog] ADD [RowVer] ROWVERSION NOT NULL;
GO

/*------------------------------------------------------------------------------
  3. Rules (CHECK), created again on every run
------------------------------------------------------------------------------*/
IF OBJECT_ID(N'dbo.CK_Settings_SettingID', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_SettingID];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_SettingID] CHECK ([SettingID] =1);
IF OBJECT_ID(N'dbo.CK_Settings_VATNumber', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_VATNumber];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_VATNumber] CHECK (([VATNumber] IS NULL) OR ([VATNumber] LIKE N'3[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]3') OR ([VATNumber] LIKE N'[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]'));
IF OBJECT_ID(N'dbo.CK_Settings_CountryCode', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_CountryCode];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_CountryCode] CHECK ([CountryCode] IN (N'SA',N'EG'));
IF OBJECT_ID(N'dbo.CK_Settings_VATRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_VATRate];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_VATRate] CHECK (([VATRate] >=0 AND [VATRate] <1));
IF OBJECT_ID(N'dbo.CK_Settings_ZatcaPhase', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_ZatcaPhase];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_ZatcaPhase] CHECK ([ZatcaPhase] IN (1,2));
IF OBJECT_ID(N'dbo.CK_Settings_BackupKeepCount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_BackupKeepCount];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_BackupKeepCount] CHECK ([BackupKeepCount] >=1);
IF OBJECT_ID(N'dbo.CK_Settings_SlowMovingDays', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_SlowMovingDays];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_SlowMovingDays] CHECK ([SlowMovingDays] >=1);
IF OBJECT_ID(N'dbo.CK_Settings_POSMode', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_POSMode];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_POSMode] CHECK ([POSMode] IN (N'RETAIL',N'RESTAURANT',N'CAFE'));
IF OBJECT_ID(N'dbo.CK_Settings_InvoicePrintMode', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_InvoicePrintMode];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_InvoicePrintMode] CHECK ([InvoicePrintMode] IN (N'DIRECT',N'PREVIEW',N'NONE'));
IF OBJECT_ID(N'dbo.CK_Settings_GosiEmployeeRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_GosiEmployeeRate];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_GosiEmployeeRate] CHECK (([GosiEmployeeRate] >=0 AND [GosiEmployeeRate] <1));
IF OBJECT_ID(N'dbo.CK_Settings_GosiEmployerRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_GosiEmployerRate];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_GosiEmployerRate] CHECK (([GosiEmployerRate] >=0 AND [GosiEmployerRate] <1));
IF OBJECT_ID(N'dbo.CK_Settings_GosiNonSaudiRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_GosiNonSaudiRate];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_GosiNonSaudiRate] CHECK (([GosiNonSaudiRate] >=0 AND [GosiNonSaudiRate] <1));
IF OBJECT_ID(N'dbo.CK_Settings_GosiMaxWage', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_GosiMaxWage];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_GosiMaxWage] CHECK ([GosiMaxWage] >=0);
IF OBJECT_ID(N'dbo.CK_Settings_CreditBlockDays', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_CreditBlockDays];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_CreditBlockDays] CHECK ([CreditBlockDays] >=0);
IF OBJECT_ID(N'dbo.CK_Settings_EInvoiceEnvironment', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Settings] DROP CONSTRAINT [CK_Settings_EInvoiceEnvironment];
ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [CK_Settings_EInvoiceEnvironment] CHECK ([EInvoiceEnvironment] IN (N'TEST',N'SIMULATION',N'PRODUCTION'));
IF OBJECT_ID(N'dbo.CK_Sequences_NextValue', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Sequences] DROP CONSTRAINT [CK_Sequences_NextValue];
ALTER TABLE [dbo].[Sequences] WITH CHECK ADD CONSTRAINT [CK_Sequences_NextValue] CHECK ([NextValue] >=1);
IF OBJECT_ID(N'dbo.CK_Sequences_PadLength', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Sequences] DROP CONSTRAINT [CK_Sequences_PadLength];
ALTER TABLE [dbo].[Sequences] WITH CHECK ADD CONSTRAINT [CK_Sequences_PadLength] CHECK ([PadLength] BETWEEN 0 AND 12);
IF OBJECT_ID(N'dbo.CK_Employees_MaxDiscountPercent', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Employees] DROP CONSTRAINT [CK_Employees_MaxDiscountPercent];
ALTER TABLE [dbo].[Employees] WITH CHECK ADD CONSTRAINT [CK_Employees_MaxDiscountPercent] CHECK (([MaxDiscountPercent] >=0 AND [MaxDiscountPercent] <=1));
IF OBJECT_ID(N'dbo.CK_Employees_BasicSalary', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Employees] DROP CONSTRAINT [CK_Employees_BasicSalary];
ALTER TABLE [dbo].[Employees] WITH CHECK ADD CONSTRAINT [CK_Employees_BasicSalary] CHECK ([BasicSalary] >=0);
IF OBJECT_ID(N'dbo.CK_Employees_HousingAllowance', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Employees] DROP CONSTRAINT [CK_Employees_HousingAllowance];
ALTER TABLE [dbo].[Employees] WITH CHECK ADD CONSTRAINT [CK_Employees_HousingAllowance] CHECK ([HousingAllowance] >=0);
IF OBJECT_ID(N'dbo.CK_Employees_TransportAllowance', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Employees] DROP CONSTRAINT [CK_Employees_TransportAllowance];
ALTER TABLE [dbo].[Employees] WITH CHECK ADD CONSTRAINT [CK_Employees_TransportAllowance] CHECK ([TransportAllowance] >=0);
IF OBJECT_ID(N'dbo.CK_Employees_OtherAllowance', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Employees] DROP CONSTRAINT [CK_Employees_OtherAllowance];
ALTER TABLE [dbo].[Employees] WITH CHECK ADD CONSTRAINT [CK_Employees_OtherAllowance] CHECK ([OtherAllowance] >=0);
IF OBJECT_ID(N'dbo.CK_Employees_AdvanceInstallment', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Employees] DROP CONSTRAINT [CK_Employees_AdvanceInstallment];
ALTER TABLE [dbo].[Employees] WITH CHECK ADD CONSTRAINT [CK_Employees_AdvanceInstallment] CHECK ([AdvanceInstallment] >=0);
IF OBJECT_ID(N'dbo.CK_Categories_TileColor', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Categories] DROP CONSTRAINT [CK_Categories_TileColor];
ALTER TABLE [dbo].[Categories] WITH CHECK ADD CONSTRAINT [CK_Categories_TileColor] CHECK ([TileColor] IN (N'BLUE',N'GREEN',N'ORANGE',N'PURPLE',N'RED',N'INDIGO',N'TEAL',N'PINK',N'BROWN',N'GREY'));
IF OBJECT_ID(N'dbo.CK_Currencies_CurrencyCode', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Currencies] DROP CONSTRAINT [CK_Currencies_CurrencyCode];
ALTER TABLE [dbo].[Currencies] WITH CHECK ADD CONSTRAINT [CK_Currencies_CurrencyCode] CHECK (LEN([CurrencyCode])=3);
IF OBJECT_ID(N'dbo.CK_Currencies_DecimalPlaces', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Currencies] DROP CONSTRAINT [CK_Currencies_DecimalPlaces];
ALTER TABLE [dbo].[Currencies] WITH CHECK ADD CONSTRAINT [CK_Currencies_DecimalPlaces] CHECK ([DecimalPlaces] BETWEEN 0 AND 3);
IF OBJECT_ID(N'dbo.CK_CurrencyRates_Rate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CurrencyRates] DROP CONSTRAINT [CK_CurrencyRates_Rate];
ALTER TABLE [dbo].[CurrencyRates] WITH CHECK ADD CONSTRAINT [CK_CurrencyRates_Rate] CHECK ([Rate] >0);
IF OBJECT_ID(N'dbo.CK_CashBoxes_BoxType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CashBoxes] DROP CONSTRAINT [CK_CashBoxes_BoxType];
ALTER TABLE [dbo].[CashBoxes] WITH CHECK ADD CONSTRAINT [CK_CashBoxes_BoxType] CHECK ([BoxType] IN (N'MAIN',N'CASHIER',N'CUSTODY'));
IF OBJECT_ID(N'dbo.CK_CashBoxes_OpeningBalance', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CashBoxes] DROP CONSTRAINT [CK_CashBoxes_OpeningBalance];
ALTER TABLE [dbo].[CashBoxes] WITH CHECK ADD CONSTRAINT [CK_CashBoxes_OpeningBalance] CHECK ([OpeningBalance] >=0);
IF OBJECT_ID(N'dbo.CK_Warehouses_WarehouseType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Warehouses] DROP CONSTRAINT [CK_Warehouses_WarehouseType];
ALTER TABLE [dbo].[Warehouses] WITH CHECK ADD CONSTRAINT [CK_Warehouses_WarehouseType] CHECK ([WarehouseType] IN (N'MAIN',N'BRANCH',N'VAN',N'WIP'));
IF OBJECT_ID(N'dbo.CK_Suppliers_VATNumber', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Suppliers] DROP CONSTRAINT [CK_Suppliers_VATNumber];
ALTER TABLE [dbo].[Suppliers] WITH CHECK ADD CONSTRAINT [CK_Suppliers_VATNumber] CHECK (([VATNumber] IS NULL) OR ([VATNumber] LIKE N'3[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]3') OR ([VATNumber] LIKE N'[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]'));
IF OBJECT_ID(N'dbo.CK_Suppliers_PaymentTermsDays', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Suppliers] DROP CONSTRAINT [CK_Suppliers_PaymentTermsDays];
ALTER TABLE [dbo].[Suppliers] WITH CHECK ADD CONSTRAINT [CK_Suppliers_PaymentTermsDays] CHECK ([PaymentTermsDays] >=0);
IF OBJECT_ID(N'dbo.CK_PriceLists_DiscountPercent', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PriceLists] DROP CONSTRAINT [CK_PriceLists_DiscountPercent];
ALTER TABLE [dbo].[PriceLists] WITH CHECK ADD CONSTRAINT [CK_PriceLists_DiscountPercent] CHECK (([DiscountPercent] >=0 AND [DiscountPercent] <1));
IF OBJECT_ID(N'dbo.CK_Customers_VATNumber', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Customers] DROP CONSTRAINT [CK_Customers_VATNumber];
ALTER TABLE [dbo].[Customers] WITH CHECK ADD CONSTRAINT [CK_Customers_VATNumber] CHECK (([VATNumber] IS NULL) OR ([VATNumber] LIKE N'3[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]3') OR ([VATNumber] LIKE N'[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]'));
IF OBJECT_ID(N'dbo.CK_Customers_CreditLimit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Customers] DROP CONSTRAINT [CK_Customers_CreditLimit];
ALTER TABLE [dbo].[Customers] WITH CHECK ADD CONSTRAINT [CK_Customers_CreditLimit] CHECK ([CreditLimit] >=0);
IF OBJECT_ID(N'dbo.CK_Customers_PaymentTermsDays', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Customers] DROP CONSTRAINT [CK_Customers_PaymentTermsDays];
ALTER TABLE [dbo].[Customers] WITH CHECK ADD CONSTRAINT [CK_Customers_PaymentTermsDays] CHECK ([PaymentTermsDays] >=0);
IF OBJECT_ID(N'dbo.CK_Products_PurchasePrice', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Products] DROP CONSTRAINT [CK_Products_PurchasePrice];
ALTER TABLE [dbo].[Products] WITH CHECK ADD CONSTRAINT [CK_Products_PurchasePrice] CHECK ([PurchasePrice] >=0);
IF OBJECT_ID(N'dbo.CK_Products_AverageCost', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Products] DROP CONSTRAINT [CK_Products_AverageCost];
ALTER TABLE [dbo].[Products] WITH CHECK ADD CONSTRAINT [CK_Products_AverageCost] CHECK ([AverageCost] >=0);
IF OBJECT_ID(N'dbo.CK_Products_SellingPrice', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Products] DROP CONSTRAINT [CK_Products_SellingPrice];
ALTER TABLE [dbo].[Products] WITH CHECK ADD CONSTRAINT [CK_Products_SellingPrice] CHECK ([SellingPrice] >=0);
IF OBJECT_ID(N'dbo.CK_Products_VATCategory', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Products] DROP CONSTRAINT [CK_Products_VATCategory];
ALTER TABLE [dbo].[Products] WITH CHECK ADD CONSTRAINT [CK_Products_VATCategory] CHECK ([VATCategory] IN (N'S',N'Z',N'E'));
IF OBJECT_ID(N'dbo.CK_Products_MinimumQuantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Products] DROP CONSTRAINT [CK_Products_MinimumQuantity];
ALTER TABLE [dbo].[Products] WITH CHECK ADD CONSTRAINT [CK_Products_MinimumQuantity] CHECK ([MinimumQuantity] >=0);
IF OBJECT_ID(N'dbo.CK_Products_SizePriceM', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Products] DROP CONSTRAINT [CK_Products_SizePriceM];
ALTER TABLE [dbo].[Products] WITH CHECK ADD CONSTRAINT [CK_Products_SizePriceM] CHECK ([SizePriceM] >=0);
IF OBJECT_ID(N'dbo.CK_Products_SizePriceL', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Products] DROP CONSTRAINT [CK_Products_SizePriceL];
ALTER TABLE [dbo].[Products] WITH CHECK ADD CONSTRAINT [CK_Products_SizePriceL] CHECK ([SizePriceL] >=0);
IF OBJECT_ID(N'dbo.CK_Products_ReorderQuantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Products] DROP CONSTRAINT [CK_Products_ReorderQuantity];
ALTER TABLE [dbo].[Products] WITH CHECK ADD CONSTRAINT [CK_Products_ReorderQuantity] CHECK ([ReorderQuantity] >=0);
IF OBJECT_ID(N'dbo.CK_ProductUnits_UnitFactor', N'C') IS NOT NULL
    ALTER TABLE [dbo].[ProductUnits] DROP CONSTRAINT [CK_ProductUnits_UnitFactor];
ALTER TABLE [dbo].[ProductUnits] WITH CHECK ADD CONSTRAINT [CK_ProductUnits_UnitFactor] CHECK (([UnitFactor] >0 AND [UnitFactor] <>1));
IF OBJECT_ID(N'dbo.CK_ProductUnits_SellingPrice', N'C') IS NOT NULL
    ALTER TABLE [dbo].[ProductUnits] DROP CONSTRAINT [CK_ProductUnits_SellingPrice];
ALTER TABLE [dbo].[ProductUnits] WITH CHECK ADD CONSTRAINT [CK_ProductUnits_SellingPrice] CHECK ([SellingPrice] >=0);
IF OBJECT_ID(N'dbo.CK_ProductUnits_PurchasePrice', N'C') IS NOT NULL
    ALTER TABLE [dbo].[ProductUnits] DROP CONSTRAINT [CK_ProductUnits_PurchasePrice];
ALTER TABLE [dbo].[ProductUnits] WITH CHECK ADD CONSTRAINT [CK_ProductUnits_PurchasePrice] CHECK ([PurchasePrice] >=0);
IF OBJECT_ID(N'dbo.CK_PriceListItems_UnitPrice', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PriceListItems] DROP CONSTRAINT [CK_PriceListItems_UnitPrice];
ALTER TABLE [dbo].[PriceListItems] WITH CHECK ADD CONSTRAINT [CK_PriceListItems_UnitPrice] CHECK ([UnitPrice] >0);
IF OBJECT_ID(N'dbo.CK_SalesInvoices_PaymentType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_PaymentType];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_PaymentType] CHECK ([PaymentType] IN (N'CASH',N'CREDIT'));
IF OBJECT_ID(N'dbo.CK_SalesInvoices_SubTotal', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_SubTotal];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_SubTotal] CHECK ([SubTotal] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoices_Discount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_Discount];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_Discount] CHECK ([Discount] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoices_TaxableAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_TaxableAmount];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_TaxableAmount] CHECK ([TaxableAmount] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoices_Tax', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_Tax];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_Tax] CHECK ([Tax] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoices_TotalAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_TotalAmount];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_TotalAmount] CHECK ([TotalAmount] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoices_PaidAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_PaidAmount];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_PaidAmount] CHECK ([PaidAmount] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoices_RemainingAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_RemainingAmount];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_RemainingAmount] CHECK ([RemainingAmount] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoices_AmountTendered', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_AmountTendered];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_AmountTendered] CHECK ([AmountTendered] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoices_ChangeDue', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_ChangeDue];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_ChangeDue] CHECK ([ChangeDue] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoices_InvoiceSubType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_InvoiceSubType];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_InvoiceSubType] CHECK ([InvoiceSubType] IN (N'SIMPLIFIED',N'STANDARD'));
IF OBJECT_ID(N'dbo.CK_SalesInvoices_InvoiceTypeCode', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_InvoiceTypeCode];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_InvoiceTypeCode] CHECK ([InvoiceTypeCode] IN (N'388',N'381',N'383'));
IF OBJECT_ID(N'dbo.CK_SalesInvoices_ZatcaStatus', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_ZatcaStatus];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_ZatcaStatus] CHECK ([ZatcaStatus] IN (N'NOT_SENT',N'PENDING',N'REPORTED',N'CLEARED',N'WARNING',N'REJECTED',N'SUBMITTED',N'VALID',N'INVALID',N'CANCELLED'));
IF OBJECT_ID(N'dbo.CK_SalesInvoices_OrderType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices_OrderType];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices_OrderType] CHECK (([OrderType] IS NULL) OR ([OrderType] IN (N'DINE_IN',N'TAKEAWAY',N'DELIVERY')));
IF OBJECT_ID(N'dbo.CK_SalesInvoices', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoices] DROP CONSTRAINT [CK_SalesInvoices];
ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [CK_SalesInvoices] CHECK ([PaidAmount]<=[TotalAmount] AND [RemainingAmount]=[TotalAmount]-[PaidAmount]);
IF OBJECT_ID(N'dbo.CK_SalesInvoiceDetails_Quantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] DROP CONSTRAINT [CK_SalesInvoiceDetails_Quantity];
ALTER TABLE [dbo].[SalesInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_SalesInvoiceDetails_Quantity] CHECK ([Quantity] >0);
IF OBJECT_ID(N'dbo.CK_SalesInvoiceDetails_UnitPrice', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] DROP CONSTRAINT [CK_SalesInvoiceDetails_UnitPrice];
ALTER TABLE [dbo].[SalesInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_SalesInvoiceDetails_UnitPrice] CHECK ([UnitPrice] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoiceDetails_Discount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] DROP CONSTRAINT [CK_SalesInvoiceDetails_Discount];
ALTER TABLE [dbo].[SalesInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_SalesInvoiceDetails_Discount] CHECK ([Discount] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoiceDetails_NetAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] DROP CONSTRAINT [CK_SalesInvoiceDetails_NetAmount];
ALTER TABLE [dbo].[SalesInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_SalesInvoiceDetails_NetAmount] CHECK ([NetAmount] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoiceDetails_VATCategory', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] DROP CONSTRAINT [CK_SalesInvoiceDetails_VATCategory];
ALTER TABLE [dbo].[SalesInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_SalesInvoiceDetails_VATCategory] CHECK ([VATCategory] IN (N'S',N'Z',N'E'));
IF OBJECT_ID(N'dbo.CK_SalesInvoiceDetails_VATRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] DROP CONSTRAINT [CK_SalesInvoiceDetails_VATRate];
ALTER TABLE [dbo].[SalesInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_SalesInvoiceDetails_VATRate] CHECK (([VATRate] >=0 AND [VATRate] <1));
IF OBJECT_ID(N'dbo.CK_SalesInvoiceDetails_Tax', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] DROP CONSTRAINT [CK_SalesInvoiceDetails_Tax];
ALTER TABLE [dbo].[SalesInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_SalesInvoiceDetails_Tax] CHECK ([Tax] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoiceDetails_LineTotal', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] DROP CONSTRAINT [CK_SalesInvoiceDetails_LineTotal];
ALTER TABLE [dbo].[SalesInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_SalesInvoiceDetails_LineTotal] CHECK ([LineTotal] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoiceDetails_UnitCost', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] DROP CONSTRAINT [CK_SalesInvoiceDetails_UnitCost];
ALTER TABLE [dbo].[SalesInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_SalesInvoiceDetails_UnitCost] CHECK ([UnitCost] >=0);
IF OBJECT_ID(N'dbo.CK_SalesInvoiceDetails_UnitFactor', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] DROP CONSTRAINT [CK_SalesInvoiceDetails_UnitFactor];
ALTER TABLE [dbo].[SalesInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_SalesInvoiceDetails_UnitFactor] CHECK ([UnitFactor] >0);
IF OBJECT_ID(N'dbo.CK_SalesReturns_RefundType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturns] DROP CONSTRAINT [CK_SalesReturns_RefundType];
ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [CK_SalesReturns_RefundType] CHECK ([RefundType] IN (N'CASH',N'CREDIT'));
IF OBJECT_ID(N'dbo.CK_SalesReturns_SubTotal', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturns] DROP CONSTRAINT [CK_SalesReturns_SubTotal];
ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [CK_SalesReturns_SubTotal] CHECK ([SubTotal] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReturns_Discount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturns] DROP CONSTRAINT [CK_SalesReturns_Discount];
ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [CK_SalesReturns_Discount] CHECK ([Discount] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReturns_TaxableAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturns] DROP CONSTRAINT [CK_SalesReturns_TaxableAmount];
ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [CK_SalesReturns_TaxableAmount] CHECK ([TaxableAmount] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReturns_Tax', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturns] DROP CONSTRAINT [CK_SalesReturns_Tax];
ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [CK_SalesReturns_Tax] CHECK ([Tax] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReturns_TotalAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturns] DROP CONSTRAINT [CK_SalesReturns_TotalAmount];
ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [CK_SalesReturns_TotalAmount] CHECK ([TotalAmount] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReturns_RefundedAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturns] DROP CONSTRAINT [CK_SalesReturns_RefundedAmount];
ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [CK_SalesReturns_RefundedAmount] CHECK ([RefundedAmount] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReturns_InvoiceSubType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturns] DROP CONSTRAINT [CK_SalesReturns_InvoiceSubType];
ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [CK_SalesReturns_InvoiceSubType] CHECK ([InvoiceSubType] IN (N'SIMPLIFIED',N'STANDARD'));
IF OBJECT_ID(N'dbo.CK_SalesReturns_InvoiceTypeCode', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturns] DROP CONSTRAINT [CK_SalesReturns_InvoiceTypeCode];
ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [CK_SalesReturns_InvoiceTypeCode] CHECK ([InvoiceTypeCode] IN (N'388',N'381',N'383'));
IF OBJECT_ID(N'dbo.CK_SalesReturns_ZatcaStatus', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturns] DROP CONSTRAINT [CK_SalesReturns_ZatcaStatus];
ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [CK_SalesReturns_ZatcaStatus] CHECK ([ZatcaStatus] IN (N'NOT_SENT',N'PENDING',N'REPORTED',N'CLEARED',N'WARNING',N'REJECTED',N'SUBMITTED',N'VALID',N'INVALID',N'CANCELLED'));
IF OBJECT_ID(N'dbo.CK_SalesReturns', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturns] DROP CONSTRAINT [CK_SalesReturns];
ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [CK_SalesReturns] CHECK ([RefundedAmount]<=[TotalAmount]);
IF OBJECT_ID(N'dbo.CK_SalesReturnDetails_Quantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturnDetails] DROP CONSTRAINT [CK_SalesReturnDetails_Quantity];
ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [CK_SalesReturnDetails_Quantity] CHECK ([Quantity] >0);
IF OBJECT_ID(N'dbo.CK_SalesReturnDetails_UnitPrice', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturnDetails] DROP CONSTRAINT [CK_SalesReturnDetails_UnitPrice];
ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [CK_SalesReturnDetails_UnitPrice] CHECK ([UnitPrice] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReturnDetails_Discount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturnDetails] DROP CONSTRAINT [CK_SalesReturnDetails_Discount];
ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [CK_SalesReturnDetails_Discount] CHECK ([Discount] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReturnDetails_NetAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturnDetails] DROP CONSTRAINT [CK_SalesReturnDetails_NetAmount];
ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [CK_SalesReturnDetails_NetAmount] CHECK ([NetAmount] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReturnDetails_VATCategory', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturnDetails] DROP CONSTRAINT [CK_SalesReturnDetails_VATCategory];
ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [CK_SalesReturnDetails_VATCategory] CHECK ([VATCategory] IN (N'S',N'Z',N'E'));
IF OBJECT_ID(N'dbo.CK_SalesReturnDetails_VATRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturnDetails] DROP CONSTRAINT [CK_SalesReturnDetails_VATRate];
ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [CK_SalesReturnDetails_VATRate] CHECK (([VATRate] >=0 AND [VATRate] <1));
IF OBJECT_ID(N'dbo.CK_SalesReturnDetails_Tax', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturnDetails] DROP CONSTRAINT [CK_SalesReturnDetails_Tax];
ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [CK_SalesReturnDetails_Tax] CHECK ([Tax] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReturnDetails_LineTotal', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturnDetails] DROP CONSTRAINT [CK_SalesReturnDetails_LineTotal];
ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [CK_SalesReturnDetails_LineTotal] CHECK ([LineTotal] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReturnDetails_UnitCost', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturnDetails] DROP CONSTRAINT [CK_SalesReturnDetails_UnitCost];
ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [CK_SalesReturnDetails_UnitCost] CHECK ([UnitCost] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReturnDetails_UnitFactor', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReturnDetails] DROP CONSTRAINT [CK_SalesReturnDetails_UnitFactor];
ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [CK_SalesReturnDetails_UnitFactor] CHECK ([UnitFactor] >0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoices_PaymentType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoices] DROP CONSTRAINT [CK_PurchaseInvoices_PaymentType];
ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoices_PaymentType] CHECK ([PaymentType] IN (N'CASH',N'CREDIT'));
IF OBJECT_ID(N'dbo.CK_PurchaseInvoices_SubTotal', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoices] DROP CONSTRAINT [CK_PurchaseInvoices_SubTotal];
ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoices_SubTotal] CHECK ([SubTotal] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoices_Discount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoices] DROP CONSTRAINT [CK_PurchaseInvoices_Discount];
ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoices_Discount] CHECK ([Discount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoices_TaxableAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoices] DROP CONSTRAINT [CK_PurchaseInvoices_TaxableAmount];
ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoices_TaxableAmount] CHECK ([TaxableAmount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoices_Tax', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoices] DROP CONSTRAINT [CK_PurchaseInvoices_Tax];
ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoices_Tax] CHECK ([Tax] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoices_TotalAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoices] DROP CONSTRAINT [CK_PurchaseInvoices_TotalAmount];
ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoices_TotalAmount] CHECK ([TotalAmount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoices_PaidAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoices] DROP CONSTRAINT [CK_PurchaseInvoices_PaidAmount];
ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoices_PaidAmount] CHECK ([PaidAmount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoices_RemainingAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoices] DROP CONSTRAINT [CK_PurchaseInvoices_RemainingAmount];
ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoices_RemainingAmount] CHECK ([RemainingAmount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoices_ExchangeRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoices] DROP CONSTRAINT [CK_PurchaseInvoices_ExchangeRate];
ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoices_ExchangeRate] CHECK ([ExchangeRate] >0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoices_ForeignAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoices] DROP CONSTRAINT [CK_PurchaseInvoices_ForeignAmount];
ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoices_ForeignAmount] CHECK ([ForeignAmount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoices', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoices] DROP CONSTRAINT [CK_PurchaseInvoices];
ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoices] CHECK ([PaidAmount]<=[TotalAmount] AND [RemainingAmount]=[TotalAmount]-[PaidAmount]);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoiceDetails_Quantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] DROP CONSTRAINT [CK_PurchaseInvoiceDetails_Quantity];
ALTER TABLE [dbo].[PurchaseInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoiceDetails_Quantity] CHECK ([Quantity] >0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoiceDetails_UnitCost', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] DROP CONSTRAINT [CK_PurchaseInvoiceDetails_UnitCost];
ALTER TABLE [dbo].[PurchaseInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoiceDetails_UnitCost] CHECK ([UnitCost] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoiceDetails_Discount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] DROP CONSTRAINT [CK_PurchaseInvoiceDetails_Discount];
ALTER TABLE [dbo].[PurchaseInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoiceDetails_Discount] CHECK ([Discount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoiceDetails_NetAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] DROP CONSTRAINT [CK_PurchaseInvoiceDetails_NetAmount];
ALTER TABLE [dbo].[PurchaseInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoiceDetails_NetAmount] CHECK ([NetAmount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoiceDetails_VATRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] DROP CONSTRAINT [CK_PurchaseInvoiceDetails_VATRate];
ALTER TABLE [dbo].[PurchaseInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoiceDetails_VATRate] CHECK (([VATRate] >=0 AND [VATRate] <1));
IF OBJECT_ID(N'dbo.CK_PurchaseInvoiceDetails_Tax', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] DROP CONSTRAINT [CK_PurchaseInvoiceDetails_Tax];
ALTER TABLE [dbo].[PurchaseInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoiceDetails_Tax] CHECK ([Tax] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoiceDetails_LineTotal', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] DROP CONSTRAINT [CK_PurchaseInvoiceDetails_LineTotal];
ALTER TABLE [dbo].[PurchaseInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoiceDetails_LineTotal] CHECK ([LineTotal] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseInvoiceDetails_UnitFactor', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] DROP CONSTRAINT [CK_PurchaseInvoiceDetails_UnitFactor];
ALTER TABLE [dbo].[PurchaseInvoiceDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseInvoiceDetails_UnitFactor] CHECK ([UnitFactor] >0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturns_RefundType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturns] DROP CONSTRAINT [CK_PurchaseReturns_RefundType];
ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturns_RefundType] CHECK ([RefundType] IN (N'CASH',N'CREDIT'));
IF OBJECT_ID(N'dbo.CK_PurchaseReturns_SubTotal', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturns] DROP CONSTRAINT [CK_PurchaseReturns_SubTotal];
ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturns_SubTotal] CHECK ([SubTotal] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturns_Discount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturns] DROP CONSTRAINT [CK_PurchaseReturns_Discount];
ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturns_Discount] CHECK ([Discount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturns_TaxableAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturns] DROP CONSTRAINT [CK_PurchaseReturns_TaxableAmount];
ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturns_TaxableAmount] CHECK ([TaxableAmount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturns_Tax', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturns] DROP CONSTRAINT [CK_PurchaseReturns_Tax];
ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturns_Tax] CHECK ([Tax] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturns_TotalAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturns] DROP CONSTRAINT [CK_PurchaseReturns_TotalAmount];
ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturns_TotalAmount] CHECK ([TotalAmount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturns_RefundedAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturns] DROP CONSTRAINT [CK_PurchaseReturns_RefundedAmount];
ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturns_RefundedAmount] CHECK ([RefundedAmount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturns_ExchangeRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturns] DROP CONSTRAINT [CK_PurchaseReturns_ExchangeRate];
ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturns_ExchangeRate] CHECK ([ExchangeRate] >0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturns_ForeignAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturns] DROP CONSTRAINT [CK_PurchaseReturns_ForeignAmount];
ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturns_ForeignAmount] CHECK ([ForeignAmount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturns', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturns] DROP CONSTRAINT [CK_PurchaseReturns];
ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturns] CHECK ([RefundedAmount]<=[TotalAmount]);
IF OBJECT_ID(N'dbo.CK_PurchaseReturnDetails_Quantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] DROP CONSTRAINT [CK_PurchaseReturnDetails_Quantity];
ALTER TABLE [dbo].[PurchaseReturnDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturnDetails_Quantity] CHECK ([Quantity] >0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturnDetails_UnitCost', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] DROP CONSTRAINT [CK_PurchaseReturnDetails_UnitCost];
ALTER TABLE [dbo].[PurchaseReturnDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturnDetails_UnitCost] CHECK ([UnitCost] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturnDetails_Discount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] DROP CONSTRAINT [CK_PurchaseReturnDetails_Discount];
ALTER TABLE [dbo].[PurchaseReturnDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturnDetails_Discount] CHECK ([Discount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturnDetails_NetAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] DROP CONSTRAINT [CK_PurchaseReturnDetails_NetAmount];
ALTER TABLE [dbo].[PurchaseReturnDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturnDetails_NetAmount] CHECK ([NetAmount] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturnDetails_VATRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] DROP CONSTRAINT [CK_PurchaseReturnDetails_VATRate];
ALTER TABLE [dbo].[PurchaseReturnDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturnDetails_VATRate] CHECK (([VATRate] >=0 AND [VATRate] <1));
IF OBJECT_ID(N'dbo.CK_PurchaseReturnDetails_Tax', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] DROP CONSTRAINT [CK_PurchaseReturnDetails_Tax];
ALTER TABLE [dbo].[PurchaseReturnDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturnDetails_Tax] CHECK ([Tax] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturnDetails_LineTotal', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] DROP CONSTRAINT [CK_PurchaseReturnDetails_LineTotal];
ALTER TABLE [dbo].[PurchaseReturnDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturnDetails_LineTotal] CHECK ([LineTotal] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseReturnDetails_UnitFactor', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] DROP CONSTRAINT [CK_PurchaseReturnDetails_UnitFactor];
ALTER TABLE [dbo].[PurchaseReturnDetails] WITH CHECK ADD CONSTRAINT [CK_PurchaseReturnDetails_UnitFactor] CHECK ([UnitFactor] >0);
IF OBJECT_ID(N'dbo.CK_CustomerPayments_Amount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CustomerPayments] DROP CONSTRAINT [CK_CustomerPayments_Amount];
ALTER TABLE [dbo].[CustomerPayments] WITH CHECK ADD CONSTRAINT [CK_CustomerPayments_Amount] CHECK ([Amount] >0);
IF OBJECT_ID(N'dbo.CK_CustomerPayments_ExchangeRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CustomerPayments] DROP CONSTRAINT [CK_CustomerPayments_ExchangeRate];
ALTER TABLE [dbo].[CustomerPayments] WITH CHECK ADD CONSTRAINT [CK_CustomerPayments_ExchangeRate] CHECK ([ExchangeRate] >0);
IF OBJECT_ID(N'dbo.CK_CustomerPayments_ForeignAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CustomerPayments] DROP CONSTRAINT [CK_CustomerPayments_ForeignAmount];
ALTER TABLE [dbo].[CustomerPayments] WITH CHECK ADD CONSTRAINT [CK_CustomerPayments_ForeignAmount] CHECK ([ForeignAmount] >=0);
IF OBJECT_ID(N'dbo.CK_SupplierPayments_Amount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SupplierPayments] DROP CONSTRAINT [CK_SupplierPayments_Amount];
ALTER TABLE [dbo].[SupplierPayments] WITH CHECK ADD CONSTRAINT [CK_SupplierPayments_Amount] CHECK ([Amount] >0);
IF OBJECT_ID(N'dbo.CK_SupplierPayments_ExchangeRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SupplierPayments] DROP CONSTRAINT [CK_SupplierPayments_ExchangeRate];
ALTER TABLE [dbo].[SupplierPayments] WITH CHECK ADD CONSTRAINT [CK_SupplierPayments_ExchangeRate] CHECK ([ExchangeRate] >0);
IF OBJECT_ID(N'dbo.CK_SupplierPayments_ForeignAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SupplierPayments] DROP CONSTRAINT [CK_SupplierPayments_ForeignAmount];
ALTER TABLE [dbo].[SupplierPayments] WITH CHECK ADD CONSTRAINT [CK_SupplierPayments_ForeignAmount] CHECK ([ForeignAmount] >=0);
IF OBJECT_ID(N'dbo.CK_BankTransactions_TxType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BankTransactions] DROP CONSTRAINT [CK_BankTransactions_TxType];
ALTER TABLE [dbo].[BankTransactions] WITH CHECK ADD CONSTRAINT [CK_BankTransactions_TxType] CHECK ([TxType] IN (N'DEPOSIT',N'WITHDRAW',N'SETTLEMENT',N'TRANSFER',N'OTHER_IN',N'OTHER_OUT'));
IF OBJECT_ID(N'dbo.CK_BankTransactions_Amount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BankTransactions] DROP CONSTRAINT [CK_BankTransactions_Amount];
ALTER TABLE [dbo].[BankTransactions] WITH CHECK ADD CONSTRAINT [CK_BankTransactions_Amount] CHECK ([Amount] >0);
IF OBJECT_ID(N'dbo.CK_BankTransactions_FeeAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BankTransactions] DROP CONSTRAINT [CK_BankTransactions_FeeAmount];
ALTER TABLE [dbo].[BankTransactions] WITH CHECK ADD CONSTRAINT [CK_BankTransactions_FeeAmount] CHECK ([FeeAmount] >=0);
IF OBJECT_ID(N'dbo.CK_BankTransactions_FeeVAT', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BankTransactions] DROP CONSTRAINT [CK_BankTransactions_FeeVAT];
ALTER TABLE [dbo].[BankTransactions] WITH CHECK ADD CONSTRAINT [CK_BankTransactions_FeeVAT] CHECK ([FeeVAT] >=0);
IF OBJECT_ID(N'dbo.CK_BankTransactions', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BankTransactions] DROP CONSTRAINT [CK_BankTransactions];
ALTER TABLE [dbo].[BankTransactions] WITH CHECK ADD CONSTRAINT [CK_BankTransactions] CHECK ([FeeAmount]+[FeeVAT]<[Amount] OR [TxType]<>N'SETTLEMENT');
IF OBJECT_ID(N'dbo.CK_Cheques_Direction', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Cheques] DROP CONSTRAINT [CK_Cheques_Direction];
ALTER TABLE [dbo].[Cheques] WITH CHECK ADD CONSTRAINT [CK_Cheques_Direction] CHECK ([Direction] IN (N'IN',N'OUT'));
IF OBJECT_ID(N'dbo.CK_Cheques_Amount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Cheques] DROP CONSTRAINT [CK_Cheques_Amount];
ALTER TABLE [dbo].[Cheques] WITH CHECK ADD CONSTRAINT [CK_Cheques_Amount] CHECK ([Amount] >0);
IF OBJECT_ID(N'dbo.CK_Cheques_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Cheques] DROP CONSTRAINT [CK_Cheques_Status];
ALTER TABLE [dbo].[Cheques] WITH CHECK ADD CONSTRAINT [CK_Cheques_Status] CHECK ([Status] IN (N'PENDING',N'COLLECTED',N'BOUNCED'));
IF OBJECT_ID(N'dbo.CK_Cheques', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Cheques] DROP CONSTRAINT [CK_Cheques];
ALTER TABLE [dbo].[Cheques] WITH CHECK ADD CONSTRAINT [CK_Cheques] CHECK (([Direction]=N'IN' AND [CustomerID] IS NOT NULL AND [SupplierID] IS NULL) OR ([Direction]=N'OUT' AND [SupplierID] IS NOT NULL AND [CustomerID] IS NULL));
IF OBJECT_ID(N'dbo.CK_FixedAssets_Cost', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets_Cost];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets_Cost] CHECK ([Cost] >0);
IF OBJECT_ID(N'dbo.CK_FixedAssets_InputVAT', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets_InputVAT];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets_InputVAT] CHECK ([InputVAT] >=0);
IF OBJECT_ID(N'dbo.CK_FixedAssets_SalvageValue', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets_SalvageValue];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets_SalvageValue] CHECK ([SalvageValue] >=0);
IF OBJECT_ID(N'dbo.CK_FixedAssets_UsefulLifeMonths', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets_UsefulLifeMonths];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets_UsefulLifeMonths] CHECK ([UsefulLifeMonths] >0);
IF OBJECT_ID(N'dbo.CK_FixedAssets_DepMethod', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets_DepMethod];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets_DepMethod] CHECK ([DepMethod] IN (N'SL',N'DB'));
IF OBJECT_ID(N'dbo.CK_FixedAssets_DepRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets_DepRate];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets_DepRate] CHECK (([DepRate] >=0 AND [DepRate] <=1));
IF OBJECT_ID(N'dbo.CK_FixedAssets_SourceType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets_SourceType];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets_SourceType] CHECK ([SourceType] IN (N'OPENING',N'BANK',N'CASHBOX',N'ACCOUNT'));
IF OBJECT_ID(N'dbo.CK_FixedAssets_OpeningAccumDep', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets_OpeningAccumDep];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets_OpeningAccumDep] CHECK ([OpeningAccumDep] >=0);
IF OBJECT_ID(N'dbo.CK_FixedAssets_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets_Status];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets_Status] CHECK ([Status] IN (N'ACTIVE',N'DISPOSED'));
IF OBJECT_ID(N'dbo.CK_FixedAssets_DisposalProceeds', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets_DisposalProceeds];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets_DisposalProceeds] CHECK ([DisposalProceeds] >=0);
IF OBJECT_ID(N'dbo.CK_FixedAssets_DisposalTo', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets_DisposalTo];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets_DisposalTo] CHECK (([DisposalTo] IS NULL) OR ([DisposalTo] IN (N'NONE',N'BANK',N'CASHBOX')));
IF OBJECT_ID(N'dbo.CK_FixedAssets_DisposalAccumDep', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets_DisposalAccumDep];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets_DisposalAccumDep] CHECK ([DisposalAccumDep] >=0);
IF OBJECT_ID(N'dbo.CK_FixedAssets_DisposalVAT', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets_DisposalVAT];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets_DisposalVAT] CHECK ([DisposalVAT] >=0);
IF OBJECT_ID(N'dbo.CK_FixedAssets', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FixedAssets] DROP CONSTRAINT [CK_FixedAssets];
ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [CK_FixedAssets] CHECK ([SalvageValue]+[OpeningAccumDep]<=[Cost]);
IF OBJECT_ID(N'dbo.CK_DepreciationRuns_TotalAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[DepreciationRuns] DROP CONSTRAINT [CK_DepreciationRuns_TotalAmount];
ALTER TABLE [dbo].[DepreciationRuns] WITH CHECK ADD CONSTRAINT [CK_DepreciationRuns_TotalAmount] CHECK ([TotalAmount] >=0);
IF OBJECT_ID(N'dbo.CK_AssetDepreciations_Amount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[AssetDepreciations] DROP CONSTRAINT [CK_AssetDepreciations_Amount];
ALTER TABLE [dbo].[AssetDepreciations] WITH CHECK ADD CONSTRAINT [CK_AssetDepreciations_Amount] CHECK ([Amount] >0);
IF OBJECT_ID(N'dbo.CK_AssetRevaluations_BookValueBefore', N'C') IS NOT NULL
    ALTER TABLE [dbo].[AssetRevaluations] DROP CONSTRAINT [CK_AssetRevaluations_BookValueBefore];
ALTER TABLE [dbo].[AssetRevaluations] WITH CHECK ADD CONSTRAINT [CK_AssetRevaluations_BookValueBefore] CHECK ([BookValueBefore] >=0);
IF OBJECT_ID(N'dbo.CK_AssetRevaluations_NewValue', N'C') IS NOT NULL
    ALTER TABLE [dbo].[AssetRevaluations] DROP CONSTRAINT [CK_AssetRevaluations_NewValue];
ALTER TABLE [dbo].[AssetRevaluations] WITH CHECK ADD CONSTRAINT [CK_AssetRevaluations_NewValue] CHECK ([NewValue] >=0);
IF OBJECT_ID(N'dbo.CK_AssetRevaluations_MonthsLeft', N'C') IS NOT NULL
    ALTER TABLE [dbo].[AssetRevaluations] DROP CONSTRAINT [CK_AssetRevaluations_MonthsLeft];
ALTER TABLE [dbo].[AssetRevaluations] WITH CHECK ADD CONSTRAINT [CK_AssetRevaluations_MonthsLeft] CHECK ([MonthsLeft] >0);
IF OBJECT_ID(N'dbo.CK_AssetTransfers_AssetValue', N'C') IS NOT NULL
    ALTER TABLE [dbo].[AssetTransfers] DROP CONSTRAINT [CK_AssetTransfers_AssetValue];
ALTER TABLE [dbo].[AssetTransfers] WITH CHECK ADD CONSTRAINT [CK_AssetTransfers_AssetValue] CHECK ([AssetValue] >=0);
IF OBJECT_ID(N'dbo.CK_AssetTransfers_AccumDep', N'C') IS NOT NULL
    ALTER TABLE [dbo].[AssetTransfers] DROP CONSTRAINT [CK_AssetTransfers_AccumDep];
ALTER TABLE [dbo].[AssetTransfers] WITH CHECK ADD CONSTRAINT [CK_AssetTransfers_AccumDep] CHECK ([AccumDep] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReps_CommissionRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReps] DROP CONSTRAINT [CK_SalesReps_CommissionRate];
ALTER TABLE [dbo].[SalesReps] WITH CHECK ADD CONSTRAINT [CK_SalesReps_CommissionRate] CHECK (([CommissionRate] >=0 AND [CommissionRate] <1));
IF OBJECT_ID(N'dbo.CK_SalesReps_CommissionBase', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReps] DROP CONSTRAINT [CK_SalesReps_CommissionBase];
ALTER TABLE [dbo].[SalesReps] WITH CHECK ADD CONSTRAINT [CK_SalesReps_CommissionBase] CHECK ([CommissionBase] IN (N'SALES',N'COLLECTION'));
IF OBJECT_ID(N'dbo.CK_SalesReps_CreditLimit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReps] DROP CONSTRAINT [CK_SalesReps_CreditLimit];
ALTER TABLE [dbo].[SalesReps] WITH CHECK ADD CONSTRAINT [CK_SalesReps_CreditLimit] CHECK ([CreditLimit] >=0);
IF OBJECT_ID(N'dbo.CK_SalesReps_MaxDiscountPercent', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesReps] DROP CONSTRAINT [CK_SalesReps_MaxDiscountPercent];
ALTER TABLE [dbo].[SalesReps] WITH CHECK ADD CONSTRAINT [CK_SalesReps_MaxDiscountPercent] CHECK (([MaxDiscountPercent] >=0 AND [MaxDiscountPercent] <=1));
IF OBJECT_ID(N'dbo.CK_SalesRepTargets_TargetYear', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesRepTargets] DROP CONSTRAINT [CK_SalesRepTargets_TargetYear];
ALTER TABLE [dbo].[SalesRepTargets] WITH CHECK ADD CONSTRAINT [CK_SalesRepTargets_TargetYear] CHECK ([TargetYear] BETWEEN 2000 AND 2100);
IF OBJECT_ID(N'dbo.CK_SalesRepTargets_TargetMonth', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesRepTargets] DROP CONSTRAINT [CK_SalesRepTargets_TargetMonth];
ALTER TABLE [dbo].[SalesRepTargets] WITH CHECK ADD CONSTRAINT [CK_SalesRepTargets_TargetMonth] CHECK ([TargetMonth] BETWEEN 1 AND 12);
IF OBJECT_ID(N'dbo.CK_SalesRepTargets_TargetAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesRepTargets] DROP CONSTRAINT [CK_SalesRepTargets_TargetAmount];
ALTER TABLE [dbo].[SalesRepTargets] WITH CHECK ADD CONSTRAINT [CK_SalesRepTargets_TargetAmount] CHECK ([TargetAmount] >=0);
IF OBJECT_ID(N'dbo.CK_CommissionRuns_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CommissionRuns] DROP CONSTRAINT [CK_CommissionRuns_Status];
ALTER TABLE [dbo].[CommissionRuns] WITH CHECK ADD CONSTRAINT [CK_CommissionRuns_Status] CHECK ([Status] IN (N'DRAFT',N'POSTED'));
IF OBJECT_ID(N'dbo.CK_CommissionRuns_TotalAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CommissionRuns] DROP CONSTRAINT [CK_CommissionRuns_TotalAmount];
ALTER TABLE [dbo].[CommissionRuns] WITH CHECK ADD CONSTRAINT [CK_CommissionRuns_TotalAmount] CHECK ([TotalAmount] >=0);
IF OBJECT_ID(N'dbo.CK_CommissionLines_CommissionRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CommissionLines] DROP CONSTRAINT [CK_CommissionLines_CommissionRate];
ALTER TABLE [dbo].[CommissionLines] WITH CHECK ADD CONSTRAINT [CK_CommissionLines_CommissionRate] CHECK (([CommissionRate] >=0 AND [CommissionRate] <1));
IF OBJECT_ID(N'dbo.CK_CommissionLines_Commission', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CommissionLines] DROP CONSTRAINT [CK_CommissionLines_Commission];
ALTER TABLE [dbo].[CommissionLines] WITH CHECK ADD CONSTRAINT [CK_CommissionLines_Commission] CHECK ([Commission] >=0);
IF OBJECT_ID(N'dbo.CK_BudgetLines_M1', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BudgetLines] DROP CONSTRAINT [CK_BudgetLines_M1];
ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [CK_BudgetLines_M1] CHECK ([M1] >=0);
IF OBJECT_ID(N'dbo.CK_BudgetLines_M2', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BudgetLines] DROP CONSTRAINT [CK_BudgetLines_M2];
ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [CK_BudgetLines_M2] CHECK ([M2] >=0);
IF OBJECT_ID(N'dbo.CK_BudgetLines_M3', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BudgetLines] DROP CONSTRAINT [CK_BudgetLines_M3];
ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [CK_BudgetLines_M3] CHECK ([M3] >=0);
IF OBJECT_ID(N'dbo.CK_BudgetLines_M4', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BudgetLines] DROP CONSTRAINT [CK_BudgetLines_M4];
ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [CK_BudgetLines_M4] CHECK ([M4] >=0);
IF OBJECT_ID(N'dbo.CK_BudgetLines_M5', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BudgetLines] DROP CONSTRAINT [CK_BudgetLines_M5];
ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [CK_BudgetLines_M5] CHECK ([M5] >=0);
IF OBJECT_ID(N'dbo.CK_BudgetLines_M6', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BudgetLines] DROP CONSTRAINT [CK_BudgetLines_M6];
ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [CK_BudgetLines_M6] CHECK ([M6] >=0);
IF OBJECT_ID(N'dbo.CK_BudgetLines_M7', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BudgetLines] DROP CONSTRAINT [CK_BudgetLines_M7];
ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [CK_BudgetLines_M7] CHECK ([M7] >=0);
IF OBJECT_ID(N'dbo.CK_BudgetLines_M8', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BudgetLines] DROP CONSTRAINT [CK_BudgetLines_M8];
ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [CK_BudgetLines_M8] CHECK ([M8] >=0);
IF OBJECT_ID(N'dbo.CK_BudgetLines_M9', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BudgetLines] DROP CONSTRAINT [CK_BudgetLines_M9];
ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [CK_BudgetLines_M9] CHECK ([M9] >=0);
IF OBJECT_ID(N'dbo.CK_BudgetLines_M10', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BudgetLines] DROP CONSTRAINT [CK_BudgetLines_M10];
ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [CK_BudgetLines_M10] CHECK ([M10] >=0);
IF OBJECT_ID(N'dbo.CK_BudgetLines_M11', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BudgetLines] DROP CONSTRAINT [CK_BudgetLines_M11];
ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [CK_BudgetLines_M11] CHECK ([M11] >=0);
IF OBJECT_ID(N'dbo.CK_BudgetLines_M12', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BudgetLines] DROP CONSTRAINT [CK_BudgetLines_M12];
ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [CK_BudgetLines_M12] CHECK ([M12] >=0);
IF OBJECT_ID(N'dbo.CK_PayrollRuns_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollRuns] DROP CONSTRAINT [CK_PayrollRuns_Status];
ALTER TABLE [dbo].[PayrollRuns] WITH CHECK ADD CONSTRAINT [CK_PayrollRuns_Status] CHECK ([Status] IN (N'DRAFT',N'POSTED'));
IF OBJECT_ID(N'dbo.CK_PayrollRuns_PaidFrom', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollRuns] DROP CONSTRAINT [CK_PayrollRuns_PaidFrom];
ALTER TABLE [dbo].[PayrollRuns] WITH CHECK ADD CONSTRAINT [CK_PayrollRuns_PaidFrom] CHECK (([PaidFrom] IS NULL) OR ([PaidFrom] IN (N'BANK',N'CASHBOX')));
IF OBJECT_ID(N'dbo.CK_PayrollRuns_PaidAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollRuns] DROP CONSTRAINT [CK_PayrollRuns_PaidAmount];
ALTER TABLE [dbo].[PayrollRuns] WITH CHECK ADD CONSTRAINT [CK_PayrollRuns_PaidAmount] CHECK ([PaidAmount] >=0);
IF OBJECT_ID(N'dbo.CK_PayrollLines_Basic', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollLines] DROP CONSTRAINT [CK_PayrollLines_Basic];
ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [CK_PayrollLines_Basic] CHECK ([Basic] >=0);
IF OBJECT_ID(N'dbo.CK_PayrollLines_Housing', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollLines] DROP CONSTRAINT [CK_PayrollLines_Housing];
ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [CK_PayrollLines_Housing] CHECK ([Housing] >=0);
IF OBJECT_ID(N'dbo.CK_PayrollLines_OtherAllow', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollLines] DROP CONSTRAINT [CK_PayrollLines_OtherAllow];
ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [CK_PayrollLines_OtherAllow] CHECK ([OtherAllow] >=0);
IF OBJECT_ID(N'dbo.CK_PayrollLines_Overtime', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollLines] DROP CONSTRAINT [CK_PayrollLines_Overtime];
ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [CK_PayrollLines_Overtime] CHECK ([Overtime] >=0);
IF OBJECT_ID(N'dbo.CK_PayrollLines_Additions', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollLines] DROP CONSTRAINT [CK_PayrollLines_Additions];
ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [CK_PayrollLines_Additions] CHECK ([Additions] >=0);
IF OBJECT_ID(N'dbo.CK_PayrollLines_AbsenceDeduction', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollLines] DROP CONSTRAINT [CK_PayrollLines_AbsenceDeduction];
ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [CK_PayrollLines_AbsenceDeduction] CHECK ([AbsenceDeduction] >=0);
IF OBJECT_ID(N'dbo.CK_PayrollLines_AdvanceDeduction', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollLines] DROP CONSTRAINT [CK_PayrollLines_AdvanceDeduction];
ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [CK_PayrollLines_AdvanceDeduction] CHECK ([AdvanceDeduction] >=0);
IF OBJECT_ID(N'dbo.CK_PayrollLines_OtherDeduction', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollLines] DROP CONSTRAINT [CK_PayrollLines_OtherDeduction];
ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [CK_PayrollLines_OtherDeduction] CHECK ([OtherDeduction] >=0);
IF OBJECT_ID(N'dbo.CK_PayrollLines_GosiWage', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollLines] DROP CONSTRAINT [CK_PayrollLines_GosiWage];
ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [CK_PayrollLines_GosiWage] CHECK ([GosiWage] >=0);
IF OBJECT_ID(N'dbo.CK_PayrollLines_GosiEmployee', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollLines] DROP CONSTRAINT [CK_PayrollLines_GosiEmployee];
ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [CK_PayrollLines_GosiEmployee] CHECK ([GosiEmployee] >=0);
IF OBJECT_ID(N'dbo.CK_PayrollLines_GosiEmployer', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PayrollLines] DROP CONSTRAINT [CK_PayrollLines_GosiEmployer];
ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [CK_PayrollLines_GosiEmployer] CHECK ([GosiEmployer] >=0);
IF OBJECT_ID(N'dbo.CK_BankReconciliations_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[BankReconciliations] DROP CONSTRAINT [CK_BankReconciliations_Status];
ALTER TABLE [dbo].[BankReconciliations] WITH CHECK ADD CONSTRAINT [CK_BankReconciliations_Status] CHECK ([Status] IN (N'OPEN',N'DONE'));
IF OBJECT_ID(N'dbo.CK_CustomerAllocations_Amount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CustomerAllocations] DROP CONSTRAINT [CK_CustomerAllocations_Amount];
ALTER TABLE [dbo].[CustomerAllocations] WITH CHECK ADD CONSTRAINT [CK_CustomerAllocations_Amount] CHECK ([Amount] >0);
IF OBJECT_ID(N'dbo.CK_SupplierAllocations_Amount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SupplierAllocations] DROP CONSTRAINT [CK_SupplierAllocations_Amount];
ALTER TABLE [dbo].[SupplierAllocations] WITH CHECK ADD CONSTRAINT [CK_SupplierAllocations_Amount] CHECK ([Amount] >0);
IF OBJECT_ID(N'dbo.CK_Expenses_Amount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Expenses] DROP CONSTRAINT [CK_Expenses_Amount];
ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [CK_Expenses_Amount] CHECK ([Amount] >0);
IF OBJECT_ID(N'dbo.CK_Expenses_Tax', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Expenses] DROP CONSTRAINT [CK_Expenses_Tax];
ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [CK_Expenses_Tax] CHECK ([Tax] >=0);
IF OBJECT_ID(N'dbo.CK_Expenses_TotalAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Expenses] DROP CONSTRAINT [CK_Expenses_TotalAmount];
ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [CK_Expenses_TotalAmount] CHECK ([TotalAmount] >=0);
IF OBJECT_ID(N'dbo.CK_Expenses_ExchangeRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Expenses] DROP CONSTRAINT [CK_Expenses_ExchangeRate];
ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [CK_Expenses_ExchangeRate] CHECK ([ExchangeRate] >0);
IF OBJECT_ID(N'dbo.CK_Expenses_ForeignAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Expenses] DROP CONSTRAINT [CK_Expenses_ForeignAmount];
ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [CK_Expenses_ForeignAmount] CHECK ([ForeignAmount] >=0);
IF OBJECT_ID(N'dbo.CK_Expenses_ForeignTax', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Expenses] DROP CONSTRAINT [CK_Expenses_ForeignTax];
ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [CK_Expenses_ForeignTax] CHECK ([ForeignTax] >=0);
IF OBJECT_ID(N'dbo.CK_Expenses', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Expenses] DROP CONSTRAINT [CK_Expenses];
ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [CK_Expenses] CHECK ([TotalAmount]=[Amount]+[Tax]);
IF OBJECT_ID(N'dbo.CK_RecurringExpenses_Amount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[RecurringExpenses] DROP CONSTRAINT [CK_RecurringExpenses_Amount];
ALTER TABLE [dbo].[RecurringExpenses] WITH CHECK ADD CONSTRAINT [CK_RecurringExpenses_Amount] CHECK ([Amount] >0);
IF OBJECT_ID(N'dbo.CK_RecurringExpenses_Tax', N'C') IS NOT NULL
    ALTER TABLE [dbo].[RecurringExpenses] DROP CONSTRAINT [CK_RecurringExpenses_Tax];
ALTER TABLE [dbo].[RecurringExpenses] WITH CHECK ADD CONSTRAINT [CK_RecurringExpenses_Tax] CHECK ([Tax] >=0);
IF OBJECT_ID(N'dbo.CK_RecurringExpenses_Frequency', N'C') IS NOT NULL
    ALTER TABLE [dbo].[RecurringExpenses] DROP CONSTRAINT [CK_RecurringExpenses_Frequency];
ALTER TABLE [dbo].[RecurringExpenses] WITH CHECK ADD CONSTRAINT [CK_RecurringExpenses_Frequency] CHECK ([Frequency] IN (N'MONTHLY',N'QUARTERLY',N'YEARLY'));
IF OBJECT_ID(N'dbo.CK_RecurringExpenses_DueDay', N'C') IS NOT NULL
    ALTER TABLE [dbo].[RecurringExpenses] DROP CONSTRAINT [CK_RecurringExpenses_DueDay];
ALTER TABLE [dbo].[RecurringExpenses] WITH CHECK ADD CONSTRAINT [CK_RecurringExpenses_DueDay] CHECK ([DueDay] BETWEEN 1 AND 28);
IF OBJECT_ID(N'dbo.CK_CashVouchers_VoucherType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CashVouchers] DROP CONSTRAINT [CK_CashVouchers_VoucherType];
ALTER TABLE [dbo].[CashVouchers] WITH CHECK ADD CONSTRAINT [CK_CashVouchers_VoucherType] CHECK ([VoucherType] IN (N'IN',N'OUT',N'TRANSFER'));
IF OBJECT_ID(N'dbo.CK_CashVouchers_Category', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CashVouchers] DROP CONSTRAINT [CK_CashVouchers_Category];
ALTER TABLE [dbo].[CashVouchers] WITH CHECK ADD CONSTRAINT [CK_CashVouchers_Category] CHECK ([Category] IN (N'OTHER',N'OWNER',N'EXPENSE',N'ADVANCE',N'SHORTAGE',N'OVERAGE',N'TRANSFER',N'COMMISSION'));
IF OBJECT_ID(N'dbo.CK_CashVouchers_Amount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CashVouchers] DROP CONSTRAINT [CK_CashVouchers_Amount];
ALTER TABLE [dbo].[CashVouchers] WITH CHECK ADD CONSTRAINT [CK_CashVouchers_Amount] CHECK ([Amount] >0);
IF OBJECT_ID(N'dbo.CK_CashVouchers', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CashVouchers] DROP CONSTRAINT [CK_CashVouchers];
ALTER TABLE [dbo].[CashVouchers] WITH CHECK ADD CONSTRAINT [CK_CashVouchers] CHECK ([VoucherType]<>N'TRANSFER' OR ([ToCashBoxID] IS NOT NULL AND [ToCashBoxID]<>[CashBoxID]));
IF OBJECT_ID(N'dbo.CK_CashClosings_CashIn', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CashClosings] DROP CONSTRAINT [CK_CashClosings_CashIn];
ALTER TABLE [dbo].[CashClosings] WITH CHECK ADD CONSTRAINT [CK_CashClosings_CashIn] CHECK ([CashIn] >=0);
IF OBJECT_ID(N'dbo.CK_CashClosings_CashOut', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CashClosings] DROP CONSTRAINT [CK_CashClosings_CashOut];
ALTER TABLE [dbo].[CashClosings] WITH CHECK ADD CONSTRAINT [CK_CashClosings_CashOut] CHECK ([CashOut] >=0);
IF OBJECT_ID(N'dbo.CK_CashClosings_CountedAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CashClosings] DROP CONSTRAINT [CK_CashClosings_CountedAmount];
ALTER TABLE [dbo].[CashClosings] WITH CHECK ADD CONSTRAINT [CK_CashClosings_CountedAmount] CHECK ([CountedAmount] >=0);
IF OBJECT_ID(N'dbo.CK_CashClosings_Destination', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CashClosings] DROP CONSTRAINT [CK_CashClosings_Destination];
ALTER TABLE [dbo].[CashClosings] WITH CHECK ADD CONSTRAINT [CK_CashClosings_Destination] CHECK ([Destination] IN (N'MAIN',N'OWNER',N'KEEP'));
IF OBJECT_ID(N'dbo.CK_CashClosings_TransferAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CashClosings] DROP CONSTRAINT [CK_CashClosings_TransferAmount];
ALTER TABLE [dbo].[CashClosings] WITH CHECK ADD CONSTRAINT [CK_CashClosings_TransferAmount] CHECK ([TransferAmount] >=0);
IF OBJECT_ID(N'dbo.CK_CashClosings_KeptAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CashClosings] DROP CONSTRAINT [CK_CashClosings_KeptAmount];
ALTER TABLE [dbo].[CashClosings] WITH CHECK ADD CONSTRAINT [CK_CashClosings_KeptAmount] CHECK ([KeptAmount] >=0);
IF OBJECT_ID(N'dbo.CK_CashClosings', N'C') IS NOT NULL
    ALTER TABLE [dbo].[CashClosings] DROP CONSTRAINT [CK_CashClosings];
ALTER TABLE [dbo].[CashClosings] WITH CHECK ADD CONSTRAINT [CK_CashClosings] CHECK ([TransferAmount]+[KeptAmount]=[CountedAmount]);
IF OBJECT_ID(N'dbo.CK_Accounts_AccountCode', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Accounts] DROP CONSTRAINT [CK_Accounts_AccountCode];
ALTER TABLE [dbo].[Accounts] WITH CHECK ADD CONSTRAINT [CK_Accounts_AccountCode] CHECK ([AccountCode] >0);
IF OBJECT_ID(N'dbo.CK_Accounts_AccountType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Accounts] DROP CONSTRAINT [CK_Accounts_AccountType];
ALTER TABLE [dbo].[Accounts] WITH CHECK ADD CONSTRAINT [CK_Accounts_AccountType] CHECK ([AccountType] IN (N'ASSET',N'LIABILITY',N'EQUITY',N'REVENUE',N'EXPENSE'));
IF OBJECT_ID(N'dbo.CK_JournalEntries_TotalDebit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[JournalEntries] DROP CONSTRAINT [CK_JournalEntries_TotalDebit];
ALTER TABLE [dbo].[JournalEntries] WITH CHECK ADD CONSTRAINT [CK_JournalEntries_TotalDebit] CHECK ([TotalDebit] >=0);
IF OBJECT_ID(N'dbo.CK_JournalEntries_TotalCredit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[JournalEntries] DROP CONSTRAINT [CK_JournalEntries_TotalCredit];
ALTER TABLE [dbo].[JournalEntries] WITH CHECK ADD CONSTRAINT [CK_JournalEntries_TotalCredit] CHECK ([TotalCredit] >=0);
IF OBJECT_ID(N'dbo.CK_JournalEntries_Signature', N'C') IS NOT NULL
    ALTER TABLE [dbo].[JournalEntries] DROP CONSTRAINT [CK_JournalEntries_Signature];
ALTER TABLE [dbo].[JournalEntries] WITH CHECK ADD CONSTRAINT [CK_JournalEntries_Signature] CHECK ([Signature] >=0);
IF OBJECT_ID(N'dbo.CK_JournalEntries_ExchangeRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[JournalEntries] DROP CONSTRAINT [CK_JournalEntries_ExchangeRate];
ALTER TABLE [dbo].[JournalEntries] WITH CHECK ADD CONSTRAINT [CK_JournalEntries_ExchangeRate] CHECK ([ExchangeRate] >0);
IF OBJECT_ID(N'dbo.CK_JournalEntries_ForeignAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[JournalEntries] DROP CONSTRAINT [CK_JournalEntries_ForeignAmount];
ALTER TABLE [dbo].[JournalEntries] WITH CHECK ADD CONSTRAINT [CK_JournalEntries_ForeignAmount] CHECK ([ForeignAmount] >=0);
IF OBJECT_ID(N'dbo.CK_JournalEntries', N'C') IS NOT NULL
    ALTER TABLE [dbo].[JournalEntries] DROP CONSTRAINT [CK_JournalEntries];
ALTER TABLE [dbo].[JournalEntries] WITH CHECK ADD CONSTRAINT [CK_JournalEntries] CHECK ([TotalDebit]=[TotalCredit]);
IF OBJECT_ID(N'dbo.CK_JournalLines_Debit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[JournalLines] DROP CONSTRAINT [CK_JournalLines_Debit];
ALTER TABLE [dbo].[JournalLines] WITH CHECK ADD CONSTRAINT [CK_JournalLines_Debit] CHECK ([Debit] >=0);
IF OBJECT_ID(N'dbo.CK_JournalLines_Credit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[JournalLines] DROP CONSTRAINT [CK_JournalLines_Credit];
ALTER TABLE [dbo].[JournalLines] WITH CHECK ADD CONSTRAINT [CK_JournalLines_Credit] CHECK ([Credit] >=0);
IF OBJECT_ID(N'dbo.CK_PeriodClosings_ActionType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PeriodClosings] DROP CONSTRAINT [CK_PeriodClosings_ActionType];
ALTER TABLE [dbo].[PeriodClosings] WITH CHECK ADD CONSTRAINT [CK_PeriodClosings_ActionType] CHECK ([ActionType] IN (N'CLOSE',N'REOPEN',N'YEAR_CLOSE',N'YEAR_OPEN'));
IF OBJECT_ID(N'dbo.CK_FiscalPeriods', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FiscalPeriods] DROP CONSTRAINT [CK_FiscalPeriods];
ALTER TABLE [dbo].[FiscalPeriods] WITH CHECK ADD CONSTRAINT [CK_FiscalPeriods] CHECK ([EndDate]>=[StartDate]);
IF OBJECT_ID(N'dbo.CK_FiscalYearClosingLines_Debit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FiscalYearClosingLines] DROP CONSTRAINT [CK_FiscalYearClosingLines_Debit];
ALTER TABLE [dbo].[FiscalYearClosingLines] WITH CHECK ADD CONSTRAINT [CK_FiscalYearClosingLines_Debit] CHECK ([Debit] >=0);
IF OBJECT_ID(N'dbo.CK_FiscalYearClosingLines_Credit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[FiscalYearClosingLines] DROP CONSTRAINT [CK_FiscalYearClosingLines_Credit];
ALTER TABLE [dbo].[FiscalYearClosingLines] WITH CHECK ADD CONSTRAINT [CK_FiscalYearClosingLines_Credit] CHECK ([Credit] >=0);
IF OBJECT_ID(N'dbo.CK_VatReturns_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[VatReturns] DROP CONSTRAINT [CK_VatReturns_Status];
ALTER TABLE [dbo].[VatReturns] WITH CHECK ADD CONSTRAINT [CK_VatReturns_Status] CHECK ([Status] IN (N'DRAFT',N'FILED'));
IF OBJECT_ID(N'dbo.CK_VatReturns_SalesStdAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[VatReturns] DROP CONSTRAINT [CK_VatReturns_SalesStdAmount];
ALTER TABLE [dbo].[VatReturns] WITH CHECK ADD CONSTRAINT [CK_VatReturns_SalesStdAmount] CHECK ([SalesStdAmount] >=0);
IF OBJECT_ID(N'dbo.CK_VatReturns_SalesZeroAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[VatReturns] DROP CONSTRAINT [CK_VatReturns_SalesZeroAmount];
ALTER TABLE [dbo].[VatReturns] WITH CHECK ADD CONSTRAINT [CK_VatReturns_SalesZeroAmount] CHECK ([SalesZeroAmount] >=0);
IF OBJECT_ID(N'dbo.CK_VatReturns_SalesExemptAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[VatReturns] DROP CONSTRAINT [CK_VatReturns_SalesExemptAmount];
ALTER TABLE [dbo].[VatReturns] WITH CHECK ADD CONSTRAINT [CK_VatReturns_SalesExemptAmount] CHECK ([SalesExemptAmount] >=0);
IF OBJECT_ID(N'dbo.CK_VatReturns_PurchStdAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[VatReturns] DROP CONSTRAINT [CK_VatReturns_PurchStdAmount];
ALTER TABLE [dbo].[VatReturns] WITH CHECK ADD CONSTRAINT [CK_VatReturns_PurchStdAmount] CHECK ([PurchStdAmount] >=0);
IF OBJECT_ID(N'dbo.CK_VatReturns_PurchZeroAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[VatReturns] DROP CONSTRAINT [CK_VatReturns_PurchZeroAmount];
ALTER TABLE [dbo].[VatReturns] WITH CHECK ADD CONSTRAINT [CK_VatReturns_PurchZeroAmount] CHECK ([PurchZeroAmount] >=0);
IF OBJECT_ID(N'dbo.CK_VatReturns_CarriedCredit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[VatReturns] DROP CONSTRAINT [CK_VatReturns_CarriedCredit];
ALTER TABLE [dbo].[VatReturns] WITH CHECK ADD CONSTRAINT [CK_VatReturns_CarriedCredit] CHECK ([CarriedCredit] >=0);
IF OBJECT_ID(N'dbo.CK_VatReturns_PaidAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[VatReturns] DROP CONSTRAINT [CK_VatReturns_PaidAmount];
ALTER TABLE [dbo].[VatReturns] WITH CHECK ADD CONSTRAINT [CK_VatReturns_PaidAmount] CHECK ([PaidAmount] >=0);
IF OBJECT_ID(N'dbo.CK_VatReturns', N'C') IS NOT NULL
    ALTER TABLE [dbo].[VatReturns] DROP CONSTRAINT [CK_VatReturns];
ALTER TABLE [dbo].[VatReturns] WITH CHECK ADD CONSTRAINT [CK_VatReturns] CHECK ([PeriodTo]>=[PeriodFrom] AND ([PaidAmount]=0 OR [PaidAmount]<=[NetDue]));
IF OBJECT_ID(N'dbo.CK_ManualEntries_TotalAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[ManualEntries] DROP CONSTRAINT [CK_ManualEntries_TotalAmount];
ALTER TABLE [dbo].[ManualEntries] WITH CHECK ADD CONSTRAINT [CK_ManualEntries_TotalAmount] CHECK ([TotalAmount] >=0);
IF OBJECT_ID(N'dbo.CK_ManualEntries_ExchangeRate', N'C') IS NOT NULL
    ALTER TABLE [dbo].[ManualEntries] DROP CONSTRAINT [CK_ManualEntries_ExchangeRate];
ALTER TABLE [dbo].[ManualEntries] WITH CHECK ADD CONSTRAINT [CK_ManualEntries_ExchangeRate] CHECK ([ExchangeRate] >0);
IF OBJECT_ID(N'dbo.CK_ManualEntries_ForeignAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[ManualEntries] DROP CONSTRAINT [CK_ManualEntries_ForeignAmount];
ALTER TABLE [dbo].[ManualEntries] WITH CHECK ADD CONSTRAINT [CK_ManualEntries_ForeignAmount] CHECK ([ForeignAmount] >=0);
IF OBJECT_ID(N'dbo.CK_ManualEntryLines_Debit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[ManualEntryLines] DROP CONSTRAINT [CK_ManualEntryLines_Debit];
ALTER TABLE [dbo].[ManualEntryLines] WITH CHECK ADD CONSTRAINT [CK_ManualEntryLines_Debit] CHECK ([Debit] >=0);
IF OBJECT_ID(N'dbo.CK_ManualEntryLines_Credit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[ManualEntryLines] DROP CONSTRAINT [CK_ManualEntryLines_Credit];
ALTER TABLE [dbo].[ManualEntryLines] WITH CHECK ADD CONSTRAINT [CK_ManualEntryLines_Credit] CHECK ([Credit] >=0);
IF OBJECT_ID(N'dbo.CK_ManualEntryLines_ForeignDebit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[ManualEntryLines] DROP CONSTRAINT [CK_ManualEntryLines_ForeignDebit];
ALTER TABLE [dbo].[ManualEntryLines] WITH CHECK ADD CONSTRAINT [CK_ManualEntryLines_ForeignDebit] CHECK ([ForeignDebit] >=0);
IF OBJECT_ID(N'dbo.CK_ManualEntryLines_ForeignCredit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[ManualEntryLines] DROP CONSTRAINT [CK_ManualEntryLines_ForeignCredit];
ALTER TABLE [dbo].[ManualEntryLines] WITH CHECK ADD CONSTRAINT [CK_ManualEntryLines_ForeignCredit] CHECK ([ForeignCredit] >=0);
IF OBJECT_ID(N'dbo.CK_ManualEntryLines', N'C') IS NOT NULL
    ALTER TABLE [dbo].[ManualEntryLines] DROP CONSTRAINT [CK_ManualEntryLines];
ALTER TABLE [dbo].[ManualEntryLines] WITH CHECK ADD CONSTRAINT [CK_ManualEntryLines] CHECK (([Debit]=0 OR [Credit]=0) AND [Debit]+[Credit]>0);
IF OBJECT_ID(N'dbo.CK_OpeningBalances_Debit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[OpeningBalances] DROP CONSTRAINT [CK_OpeningBalances_Debit];
ALTER TABLE [dbo].[OpeningBalances] WITH CHECK ADD CONSTRAINT [CK_OpeningBalances_Debit] CHECK ([Debit] >=0);
IF OBJECT_ID(N'dbo.CK_OpeningBalances_Credit', N'C') IS NOT NULL
    ALTER TABLE [dbo].[OpeningBalances] DROP CONSTRAINT [CK_OpeningBalances_Credit];
ALTER TABLE [dbo].[OpeningBalances] WITH CHECK ADD CONSTRAINT [CK_OpeningBalances_Credit] CHECK ([Credit] >=0);
IF OBJECT_ID(N'dbo.CK_OpeningBalances', N'C') IS NOT NULL
    ALTER TABLE [dbo].[OpeningBalances] DROP CONSTRAINT [CK_OpeningBalances];
ALTER TABLE [dbo].[OpeningBalances] WITH CHECK ADD CONSTRAINT [CK_OpeningBalances] CHECK (([Debit]=0 OR [Credit]=0) AND [Debit]+[Credit]>0);
IF OBJECT_ID(N'dbo.CK_OpeningStock_Quantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[OpeningStock] DROP CONSTRAINT [CK_OpeningStock_Quantity];
ALTER TABLE [dbo].[OpeningStock] WITH CHECK ADD CONSTRAINT [CK_OpeningStock_Quantity] CHECK ([Quantity] >=0);
IF OBJECT_ID(N'dbo.CK_OpeningStock_UnitCost', N'C') IS NOT NULL
    ALTER TABLE [dbo].[OpeningStock] DROP CONSTRAINT [CK_OpeningStock_UnitCost];
ALTER TABLE [dbo].[OpeningStock] WITH CHECK ADD CONSTRAINT [CK_OpeningStock_UnitCost] CHECK ([UnitCost] >=0);
IF OBJECT_ID(N'dbo.CK_TransactionTypes_Direction', N'C') IS NOT NULL
    ALTER TABLE [dbo].[TransactionTypes] DROP CONSTRAINT [CK_TransactionTypes_Direction];
ALTER TABLE [dbo].[TransactionTypes] WITH CHECK ADD CONSTRAINT [CK_TransactionTypes_Direction] CHECK ([Direction] IN (-1,0,1));
IF OBJECT_ID(N'dbo.CK_InventoryTransactions_Quantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[InventoryTransactions] DROP CONSTRAINT [CK_InventoryTransactions_Quantity];
ALTER TABLE [dbo].[InventoryTransactions] WITH CHECK ADD CONSTRAINT [CK_InventoryTransactions_Quantity] CHECK ([Quantity] <>0);
IF OBJECT_ID(N'dbo.CK_InventoryTransactions_UnitCost', N'C') IS NOT NULL
    ALTER TABLE [dbo].[InventoryTransactions] DROP CONSTRAINT [CK_InventoryTransactions_UnitCost];
ALTER TABLE [dbo].[InventoryTransactions] WITH CHECK ADD CONSTRAINT [CK_InventoryTransactions_UnitCost] CHECK ([UnitCost] >=0);
IF OBJECT_ID(N'dbo.CK_StockTransfers_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockTransfers] DROP CONSTRAINT [CK_StockTransfers_Status];
ALTER TABLE [dbo].[StockTransfers] WITH CHECK ADD CONSTRAINT [CK_StockTransfers_Status] CHECK ([Status] IN (N'DRAFT',N'SENT',N'RECEIVED'));
IF OBJECT_ID(N'dbo.CK_StockTransfers', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockTransfers] DROP CONSTRAINT [CK_StockTransfers];
ALTER TABLE [dbo].[StockTransfers] WITH CHECK ADD CONSTRAINT [CK_StockTransfers] CHECK ([ToWarehouseID]<>[FromWarehouseID]);
IF OBJECT_ID(N'dbo.CK_StockTransferLines_Quantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockTransferLines] DROP CONSTRAINT [CK_StockTransferLines_Quantity];
ALTER TABLE [dbo].[StockTransferLines] WITH CHECK ADD CONSTRAINT [CK_StockTransferLines_Quantity] CHECK ([Quantity] >0);
IF OBJECT_ID(N'dbo.CK_StockTransferLines_UnitCost', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockTransferLines] DROP CONSTRAINT [CK_StockTransferLines_UnitCost];
ALTER TABLE [dbo].[StockTransferLines] WITH CHECK ADD CONSTRAINT [CK_StockTransferLines_UnitCost] CHECK ([UnitCost] >=0);
IF OBJECT_ID(N'dbo.CK_StockParties_PartyType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockParties] DROP CONSTRAINT [CK_StockParties_PartyType];
ALTER TABLE [dbo].[StockParties] WITH CHECK ADD CONSTRAINT [CK_StockParties_PartyType] CHECK ([PartyType] IN (N'DEPARTMENT',N'PRODUCTION',N'WASTE',N'OTHER'));
IF OBJECT_ID(N'dbo.CK_StockDocuments_DocType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockDocuments] DROP CONSTRAINT [CK_StockDocuments_DocType];
ALTER TABLE [dbo].[StockDocuments] WITH CHECK ADD CONSTRAINT [CK_StockDocuments_DocType] CHECK ([DocType] IN (N'ISSUE',N'RECEIPT'));
IF OBJECT_ID(N'dbo.CK_StockDocuments_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockDocuments] DROP CONSTRAINT [CK_StockDocuments_Status];
ALTER TABLE [dbo].[StockDocuments] WITH CHECK ADD CONSTRAINT [CK_StockDocuments_Status] CHECK ([Status] IN (N'DRAFT',N'POSTED'));
IF OBJECT_ID(N'dbo.CK_StockDocumentLines_Quantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockDocumentLines] DROP CONSTRAINT [CK_StockDocumentLines_Quantity];
ALTER TABLE [dbo].[StockDocumentLines] WITH CHECK ADD CONSTRAINT [CK_StockDocumentLines_Quantity] CHECK ([Quantity] >0);
IF OBJECT_ID(N'dbo.CK_StockDocumentLines_UnitCost', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockDocumentLines] DROP CONSTRAINT [CK_StockDocumentLines_UnitCost];
ALTER TABLE [dbo].[StockDocumentLines] WITH CHECK ADD CONSTRAINT [CK_StockDocumentLines_UnitCost] CHECK ([UnitCost] >=0);
IF OBJECT_ID(N'dbo.CK_StockNotes_NoteType', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockNotes] DROP CONSTRAINT [CK_StockNotes_NoteType];
ALTER TABLE [dbo].[StockNotes] WITH CHECK ADD CONSTRAINT [CK_StockNotes_NoteType] CHECK ([NoteType] IN (N'DELIVERY',N'RECEIPT'));
IF OBJECT_ID(N'dbo.CK_StockNotes_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockNotes] DROP CONSTRAINT [CK_StockNotes_Status];
ALTER TABLE [dbo].[StockNotes] WITH CHECK ADD CONSTRAINT [CK_StockNotes_Status] CHECK ([Status] IN (N'DRAFT',N'POSTED',N'INVOICED'));
IF OBJECT_ID(N'dbo.CK_StockNoteLines_Quantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockNoteLines] DROP CONSTRAINT [CK_StockNoteLines_Quantity];
ALTER TABLE [dbo].[StockNoteLines] WITH CHECK ADD CONSTRAINT [CK_StockNoteLines_Quantity] CHECK ([Quantity] >0);
IF OBJECT_ID(N'dbo.CK_StockNoteLines_UnitPrice', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockNoteLines] DROP CONSTRAINT [CK_StockNoteLines_UnitPrice];
ALTER TABLE [dbo].[StockNoteLines] WITH CHECK ADD CONSTRAINT [CK_StockNoteLines_UnitPrice] CHECK ([UnitPrice] >=0);
IF OBJECT_ID(N'dbo.CK_StockNoteLines_UnitCost', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockNoteLines] DROP CONSTRAINT [CK_StockNoteLines_UnitCost];
ALTER TABLE [dbo].[StockNoteLines] WITH CHECK ADD CONSTRAINT [CK_StockNoteLines_UnitCost] CHECK ([UnitCost] >=0);
IF OBJECT_ID(N'dbo.CK_SalesOrders_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesOrders] DROP CONSTRAINT [CK_SalesOrders_Status];
ALTER TABLE [dbo].[SalesOrders] WITH CHECK ADD CONSTRAINT [CK_SalesOrders_Status] CHECK ([Status] IN (N'DRAFT',N'APPROVED',N'PREPARING',N'PARTIAL',N'DELIVERED',N'INVOICED',N'CANCELLED',N'CLOSED'));
IF OBJECT_ID(N'dbo.CK_Quotations_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Quotations] DROP CONSTRAINT [CK_Quotations_Status];
ALTER TABLE [dbo].[Quotations] WITH CHECK ADD CONSTRAINT [CK_Quotations_Status] CHECK ([Status] IN (N'DRAFT',N'ISSUED',N'ORDERED',N'INVOICED',N'REJECTED'));
IF OBJECT_ID(N'dbo.CK_Quotations_SubTotal', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Quotations] DROP CONSTRAINT [CK_Quotations_SubTotal];
ALTER TABLE [dbo].[Quotations] WITH CHECK ADD CONSTRAINT [CK_Quotations_SubTotal] CHECK ([SubTotal] >=0);
IF OBJECT_ID(N'dbo.CK_Quotations_Tax', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Quotations] DROP CONSTRAINT [CK_Quotations_Tax];
ALTER TABLE [dbo].[Quotations] WITH CHECK ADD CONSTRAINT [CK_Quotations_Tax] CHECK ([Tax] >=0);
IF OBJECT_ID(N'dbo.CK_Quotations_TotalAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[Quotations] DROP CONSTRAINT [CK_Quotations_TotalAmount];
ALTER TABLE [dbo].[Quotations] WITH CHECK ADD CONSTRAINT [CK_Quotations_TotalAmount] CHECK ([TotalAmount] >=0);
IF OBJECT_ID(N'dbo.CK_QuotationLines_Quantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[QuotationLines] DROP CONSTRAINT [CK_QuotationLines_Quantity];
ALTER TABLE [dbo].[QuotationLines] WITH CHECK ADD CONSTRAINT [CK_QuotationLines_Quantity] CHECK ([Quantity] >0);
IF OBJECT_ID(N'dbo.CK_QuotationLines_UnitPrice', N'C') IS NOT NULL
    ALTER TABLE [dbo].[QuotationLines] DROP CONSTRAINT [CK_QuotationLines_UnitPrice];
ALTER TABLE [dbo].[QuotationLines] WITH CHECK ADD CONSTRAINT [CK_QuotationLines_UnitPrice] CHECK ([UnitPrice] >=0);
IF OBJECT_ID(N'dbo.CK_QuotationLines_NetAmount', N'C') IS NOT NULL
    ALTER TABLE [dbo].[QuotationLines] DROP CONSTRAINT [CK_QuotationLines_NetAmount];
ALTER TABLE [dbo].[QuotationLines] WITH CHECK ADD CONSTRAINT [CK_QuotationLines_NetAmount] CHECK ([NetAmount] >=0);
IF OBJECT_ID(N'dbo.CK_QuotationLines_Tax', N'C') IS NOT NULL
    ALTER TABLE [dbo].[QuotationLines] DROP CONSTRAINT [CK_QuotationLines_Tax];
ALTER TABLE [dbo].[QuotationLines] WITH CHECK ADD CONSTRAINT [CK_QuotationLines_Tax] CHECK ([Tax] >=0);
IF OBJECT_ID(N'dbo.CK_QuotationLines_LineTotal', N'C') IS NOT NULL
    ALTER TABLE [dbo].[QuotationLines] DROP CONSTRAINT [CK_QuotationLines_LineTotal];
ALTER TABLE [dbo].[QuotationLines] WITH CHECK ADD CONSTRAINT [CK_QuotationLines_LineTotal] CHECK ([LineTotal] >=0);
IF OBJECT_ID(N'dbo.CK_PurchaseOrders_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseOrders] DROP CONSTRAINT [CK_PurchaseOrders_Status];
ALTER TABLE [dbo].[PurchaseOrders] WITH CHECK ADD CONSTRAINT [CK_PurchaseOrders_Status] CHECK ([Status] IN (N'DRAFT',N'APPROVED',N'PARTIAL',N'RECEIVED',N'INVOICED',N'CANCELLED',N'CLOSED'));
IF OBJECT_ID(N'dbo.CK_PurchaseOrderLines_Quantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseOrderLines] DROP CONSTRAINT [CK_PurchaseOrderLines_Quantity];
ALTER TABLE [dbo].[PurchaseOrderLines] WITH CHECK ADD CONSTRAINT [CK_PurchaseOrderLines_Quantity] CHECK ([Quantity] >0);
IF OBJECT_ID(N'dbo.CK_PurchaseOrderLines_UnitCost', N'C') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseOrderLines] DROP CONSTRAINT [CK_PurchaseOrderLines_UnitCost];
ALTER TABLE [dbo].[PurchaseOrderLines] WITH CHECK ADD CONSTRAINT [CK_PurchaseOrderLines_UnitCost] CHECK ([UnitCost] >=0);
IF OBJECT_ID(N'dbo.CK_RepDays_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[RepDays] DROP CONSTRAINT [CK_RepDays_Status];
ALTER TABLE [dbo].[RepDays] WITH CHECK ADD CONSTRAINT [CK_RepDays_Status] CHECK ([Status] IN (N'OPEN',N'SETTLED'));
IF OBJECT_ID(N'dbo.CK_RepDays_GoodsShortage', N'C') IS NOT NULL
    ALTER TABLE [dbo].[RepDays] DROP CONSTRAINT [CK_RepDays_GoodsShortage];
ALTER TABLE [dbo].[RepDays] WITH CHECK ADD CONSTRAINT [CK_RepDays_GoodsShortage] CHECK ([GoodsShortage] >=0);
IF OBJECT_ID(N'dbo.CK_RouteStops_VisitWeekday', N'C') IS NOT NULL
    ALTER TABLE [dbo].[RouteStops] DROP CONSTRAINT [CK_RouteStops_VisitWeekday];
ALTER TABLE [dbo].[RouteStops] WITH CHECK ADD CONSTRAINT [CK_RouteStops_VisitWeekday] CHECK ([VisitWeekday] BETWEEN 1 AND 7);
IF OBJECT_ID(N'dbo.CK_RepVisits_Result', N'C') IS NOT NULL
    ALTER TABLE [dbo].[RepVisits] DROP CONSTRAINT [CK_RepVisits_Result];
ALTER TABLE [dbo].[RepVisits] WITH CHECK ADD CONSTRAINT [CK_RepVisits_Result] CHECK ([Result] IN (N'PENDING',N'ORDER',N'SALE',N'PAYMENT',N'NO_SALE',N'CLOSED'));
IF OBJECT_ID(N'dbo.CK_RepPrices_UnitPrice', N'C') IS NOT NULL
    ALTER TABLE [dbo].[RepPrices] DROP CONSTRAINT [CK_RepPrices_UnitPrice];
ALTER TABLE [dbo].[RepPrices] WITH CHECK ADD CONSTRAINT [CK_RepPrices_UnitPrice] CHECK ([UnitPrice] >=0);
IF OBJECT_ID(N'dbo.CK_SalesOrderLines_Quantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesOrderLines] DROP CONSTRAINT [CK_SalesOrderLines_Quantity];
ALTER TABLE [dbo].[SalesOrderLines] WITH CHECK ADD CONSTRAINT [CK_SalesOrderLines_Quantity] CHECK ([Quantity] >0);
IF OBJECT_ID(N'dbo.CK_SalesOrderLines_UnitPrice', N'C') IS NOT NULL
    ALTER TABLE [dbo].[SalesOrderLines] DROP CONSTRAINT [CK_SalesOrderLines_UnitPrice];
ALTER TABLE [dbo].[SalesOrderLines] WITH CHECK ADD CONSTRAINT [CK_SalesOrderLines_UnitPrice] CHECK ([UnitPrice] >=0);
IF OBJECT_ID(N'dbo.CK_StockCounts_Status', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockCounts] DROP CONSTRAINT [CK_StockCounts_Status];
ALTER TABLE [dbo].[StockCounts] WITH CHECK ADD CONSTRAINT [CK_StockCounts_Status] CHECK ([Status] IN (N'OPEN',N'POSTED',N'CANCELLED'));
IF OBJECT_ID(N'dbo.CK_StockCountDetails_ActualQuantity', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockCountDetails] DROP CONSTRAINT [CK_StockCountDetails_ActualQuantity];
ALTER TABLE [dbo].[StockCountDetails] WITH CHECK ADD CONSTRAINT [CK_StockCountDetails_ActualQuantity] CHECK (([ActualQuantity] IS NULL) OR ([ActualQuantity] >=0));
IF OBJECT_ID(N'dbo.CK_StockCountDetails_UnitCost', N'C') IS NOT NULL
    ALTER TABLE [dbo].[StockCountDetails] DROP CONSTRAINT [CK_StockCountDetails_UnitCost];
ALTER TABLE [dbo].[StockCountDetails] WITH CHECK ADD CONSTRAINT [CK_StockCountDetails_UnitCost] CHECK ([UnitCost] >=0);
IF OBJECT_ID(N'dbo.CK_LabelSettings_LabelSettingID', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_LabelSettingID];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_LabelSettingID] CHECK ([LabelSettingID] =1);
IF OBJECT_ID(N'dbo.CK_LabelSettings_LabelWidth', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_LabelWidth];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_LabelWidth] CHECK ([LabelWidth] BETWEEN 15 AND 210);
IF OBJECT_ID(N'dbo.CK_LabelSettings_LabelHeight', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_LabelHeight];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_LabelHeight] CHECK ([LabelHeight] BETWEEN 10 AND 297);
IF OBJECT_ID(N'dbo.CK_LabelSettings_LabelsAcross', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_LabelsAcross];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_LabelsAcross] CHECK ([LabelsAcross] BETWEEN 1 AND 10);
IF OBJECT_ID(N'dbo.CK_LabelSettings_ColumnGap', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_ColumnGap];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_ColumnGap] CHECK ([ColumnGap] BETWEEN 0 AND 50);
IF OBJECT_ID(N'dbo.CK_LabelSettings_RowGap', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_RowGap];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_RowGap] CHECK ([RowGap] BETWEEN 0 AND 50);
IF OBJECT_ID(N'dbo.CK_LabelSettings_MarginTop', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_MarginTop];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_MarginTop] CHECK ([MarginTop] BETWEEN 0 AND 50);
IF OBJECT_ID(N'dbo.CK_LabelSettings_MarginBottom', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_MarginBottom];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_MarginBottom] CHECK ([MarginBottom] BETWEEN 0 AND 50);
IF OBJECT_ID(N'dbo.CK_LabelSettings_MarginLeft', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_MarginLeft];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_MarginLeft] CHECK ([MarginLeft] BETWEEN 0 AND 50);
IF OBJECT_ID(N'dbo.CK_LabelSettings_MarginRight', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_MarginRight];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_MarginRight] CHECK ([MarginRight] BETWEEN 0 AND 50);
IF OBJECT_ID(N'dbo.CK_LabelSettings_BarHeight', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_BarHeight];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_BarHeight] CHECK ([BarHeight] BETWEEN 3 AND 100);
IF OBJECT_ID(N'dbo.CK_LabelSettings_BarWidth', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_BarWidth];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_BarWidth] CHECK ([BarWidth] BETWEEN 0.1 AND 1);
IF OBJECT_ID(N'dbo.CK_LabelSettings_TopLine1', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_TopLine1];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_TopLine1] CHECK ([TopLine1] IN (N'NONE',N'STORE',N'NAME',N'PRICE',N'CODE',N'BARCODE'));
IF OBJECT_ID(N'dbo.CK_LabelSettings_TopLine2', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_TopLine2];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_TopLine2] CHECK ([TopLine2] IN (N'NONE',N'STORE',N'NAME',N'PRICE',N'CODE',N'BARCODE'));
IF OBJECT_ID(N'dbo.CK_LabelSettings_BottomLine1', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_BottomLine1];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_BottomLine1] CHECK ([BottomLine1] IN (N'NONE',N'STORE',N'NAME',N'PRICE',N'CODE',N'BARCODE'));
IF OBJECT_ID(N'dbo.CK_LabelSettings_BottomLine2', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_BottomLine2];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_BottomLine2] CHECK ([BottomLine2] IN (N'NONE',N'STORE',N'NAME',N'PRICE',N'CODE',N'BARCODE'));
IF OBJECT_ID(N'dbo.CK_LabelSettings_FontSize', N'C') IS NOT NULL
    ALTER TABLE [dbo].[LabelSettings] DROP CONSTRAINT [CK_LabelSettings_FontSize];
ALTER TABLE [dbo].[LabelSettings] WITH CHECK ADD CONSTRAINT [CK_LabelSettings_FontSize] CHECK ([FontSize] BETWEEN 5 AND 16);
IF OBJECT_ID(N'dbo.CK_EInvoiceLog_Result', N'C') IS NOT NULL
    ALTER TABLE [dbo].[EInvoiceLog] DROP CONSTRAINT [CK_EInvoiceLog_Result];
ALTER TABLE [dbo].[EInvoiceLog] WITH CHECK ADD CONSTRAINT [CK_EInvoiceLog_Result] CHECK ([Result] IN (N'OK',N'WARNING',N'REJECTED',N'ERROR',N'NETWORK'));
GO

/*------------------------------------------------------------------------------
  4. Indexes
------------------------------------------------------------------------------*/
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_RoleCode' AND object_id = OBJECT_ID(N'dbo.Roles'))
    CREATE UNIQUE INDEX [UX_RoleCode] ON [dbo].[Roles] ([RoleCode]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_Username' AND object_id = OBJECT_ID(N'dbo.Employees'))
    CREATE UNIQUE INDEX [UX_Username] ON [dbo].[Employees] ([Username]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_MachineID' AND object_id = OBJECT_ID(N'dbo.Activations'))
    CREATE UNIQUE INDEX [UX_MachineID] ON [dbo].[Activations] ([MachineID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_CategoryName' AND object_id = OBJECT_ID(N'dbo.Categories'))
    CREATE UNIQUE INDEX [UX_CategoryName] ON [dbo].[Categories] ([CategoryName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_UnitName' AND object_id = OBJECT_ID(N'dbo.Units'))
    CREATE UNIQUE INDEX [UX_UnitName] ON [dbo].[Units] ([UnitName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_MethodName' AND object_id = OBJECT_ID(N'dbo.PaymentMethods'))
    CREATE UNIQUE INDEX [UX_MethodName] ON [dbo].[PaymentMethods] ([MethodName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_CurrencyName' AND object_id = OBJECT_ID(N'dbo.Currencies'))
    CREATE UNIQUE INDEX [UX_CurrencyName] ON [dbo].[Currencies] ([CurrencyName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_CurrencyCode_RateDate' AND object_id = OBJECT_ID(N'dbo.CurrencyRates'))
    CREATE UNIQUE INDEX [UX_CurrencyCode_RateDate] ON [dbo].[CurrencyRates] ([CurrencyCode], [RateDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_BoxName' AND object_id = OBJECT_ID(N'dbo.CashBoxes'))
    CREATE UNIQUE INDEX [UX_BoxName] ON [dbo].[CashBoxes] ([BoxName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_CashBoxID_EmployeeID' AND object_id = OBJECT_ID(N'dbo.CashBoxUsers'))
    CREATE UNIQUE INDEX [UX_CashBoxID_EmployeeID] ON [dbo].[CashBoxUsers] ([CashBoxID], [EmployeeID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_WarehouseName' AND object_id = OBJECT_ID(N'dbo.Warehouses'))
    CREATE UNIQUE INDEX [UX_WarehouseName] ON [dbo].[Warehouses] ([WarehouseName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_SupplierName' AND object_id = OBJECT_ID(N'dbo.Suppliers'))
    CREATE INDEX [IX_SupplierName] ON [dbo].[Suppliers] ([SupplierName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Mobile' AND object_id = OBJECT_ID(N'dbo.Suppliers'))
    CREATE INDEX [IX_Mobile] ON [dbo].[Suppliers] ([Mobile]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ListName' AND object_id = OBJECT_ID(N'dbo.PriceLists'))
    CREATE UNIQUE INDEX [UX_ListName] ON [dbo].[PriceLists] ([ListName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_CustomerName' AND object_id = OBJECT_ID(N'dbo.Customers'))
    CREATE INDEX [IX_CustomerName] ON [dbo].[Customers] ([CustomerName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Mobile' AND object_id = OBJECT_ID(N'dbo.Customers'))
    CREATE INDEX [IX_Mobile] ON [dbo].[Customers] ([Mobile]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ProductCode' AND object_id = OBJECT_ID(N'dbo.Products'))
    CREATE UNIQUE INDEX [UX_ProductCode] ON [dbo].[Products] ([ProductCode]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_Barcode' AND object_id = OBJECT_ID(N'dbo.Products'))
    CREATE UNIQUE INDEX [UX_Barcode] ON [dbo].[Products] ([Barcode]) WHERE [Barcode] IS NOT NULL;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ProductName' AND object_id = OBJECT_ID(N'dbo.Products'))
    CREATE INDEX [IX_ProductName] ON [dbo].[Products] ([ProductName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ProductID_UnitID' AND object_id = OBJECT_ID(N'dbo.ProductUnits'))
    CREATE UNIQUE INDEX [UX_ProductID_UnitID] ON [dbo].[ProductUnits] ([ProductID], [UnitID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_Barcode' AND object_id = OBJECT_ID(N'dbo.ProductUnits'))
    CREATE UNIQUE INDEX [UX_Barcode] ON [dbo].[ProductUnits] ([Barcode]) WHERE [Barcode] IS NOT NULL;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_PriceListID_ProductID' AND object_id = OBJECT_ID(N'dbo.PriceListItems'))
    CREATE INDEX [IX_PriceListID_ProductID] ON [dbo].[PriceListItems] ([PriceListID], [ProductID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ProductID' AND object_id = OBJECT_ID(N'dbo.PriceListItems'))
    CREATE INDEX [IX_ProductID] ON [dbo].[PriceListItems] ([ProductID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_InvoiceNumber' AND object_id = OBJECT_ID(N'dbo.SalesInvoices'))
    CREATE UNIQUE INDEX [UX_InvoiceNumber] ON [dbo].[SalesInvoices] ([InvoiceNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_InvoiceDate' AND object_id = OBJECT_ID(N'dbo.SalesInvoices'))
    CREATE INDEX [IX_InvoiceDate] ON [dbo].[SalesInvoices] ([InvoiceDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_InvoiceUUID' AND object_id = OBJECT_ID(N'dbo.SalesInvoices'))
    CREATE UNIQUE INDEX [UX_InvoiceUUID] ON [dbo].[SalesInvoices] ([InvoiceUUID]) WHERE [InvoiceUUID] IS NOT NULL;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ZatcaStatus' AND object_id = OBJECT_ID(N'dbo.SalesInvoices'))
    CREATE INDEX [IX_ZatcaStatus] ON [dbo].[SalesInvoices] ([ZatcaStatus]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_SalesInvoiceID_LineNumber' AND object_id = OBJECT_ID(N'dbo.SalesInvoiceDetails'))
    CREATE UNIQUE INDEX [UX_SalesInvoiceID_LineNumber] ON [dbo].[SalesInvoiceDetails] ([SalesInvoiceID], [LineNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ReturnNumber' AND object_id = OBJECT_ID(N'dbo.SalesReturns'))
    CREATE UNIQUE INDEX [UX_ReturnNumber] ON [dbo].[SalesReturns] ([ReturnNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ReturnDate' AND object_id = OBJECT_ID(N'dbo.SalesReturns'))
    CREATE INDEX [IX_ReturnDate] ON [dbo].[SalesReturns] ([ReturnDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_InvoiceUUID' AND object_id = OBJECT_ID(N'dbo.SalesReturns'))
    CREATE UNIQUE INDEX [UX_InvoiceUUID] ON [dbo].[SalesReturns] ([InvoiceUUID]) WHERE [InvoiceUUID] IS NOT NULL;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_InvoiceNumber' AND object_id = OBJECT_ID(N'dbo.PurchaseInvoices'))
    CREATE UNIQUE INDEX [UX_InvoiceNumber] ON [dbo].[PurchaseInvoices] ([InvoiceNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_InvoiceDate' AND object_id = OBJECT_ID(N'dbo.PurchaseInvoices'))
    CREATE INDEX [IX_InvoiceDate] ON [dbo].[PurchaseInvoices] ([InvoiceDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_SupplierInvoiceNo' AND object_id = OBJECT_ID(N'dbo.PurchaseInvoices'))
    CREATE INDEX [IX_SupplierInvoiceNo] ON [dbo].[PurchaseInvoices] ([SupplierInvoiceNo]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_PurchaseInvoiceID_LineNumber' AND object_id = OBJECT_ID(N'dbo.PurchaseInvoiceDetails'))
    CREATE UNIQUE INDEX [UX_PurchaseInvoiceID_LineNumber] ON [dbo].[PurchaseInvoiceDetails] ([PurchaseInvoiceID], [LineNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ReturnNumber' AND object_id = OBJECT_ID(N'dbo.PurchaseReturns'))
    CREATE UNIQUE INDEX [UX_ReturnNumber] ON [dbo].[PurchaseReturns] ([ReturnNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ReturnDate' AND object_id = OBJECT_ID(N'dbo.PurchaseReturns'))
    CREATE INDEX [IX_ReturnDate] ON [dbo].[PurchaseReturns] ([ReturnDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_PaymentNumber' AND object_id = OBJECT_ID(N'dbo.CustomerPayments'))
    CREATE UNIQUE INDEX [UX_PaymentNumber] ON [dbo].[CustomerPayments] ([PaymentNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_PaymentDate' AND object_id = OBJECT_ID(N'dbo.CustomerPayments'))
    CREATE INDEX [IX_PaymentDate] ON [dbo].[CustomerPayments] ([PaymentDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_PaymentNumber' AND object_id = OBJECT_ID(N'dbo.SupplierPayments'))
    CREATE UNIQUE INDEX [UX_PaymentNumber] ON [dbo].[SupplierPayments] ([PaymentNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_PaymentDate' AND object_id = OBJECT_ID(N'dbo.SupplierPayments'))
    CREATE INDEX [IX_PaymentDate] ON [dbo].[SupplierPayments] ([PaymentDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_BankName' AND object_id = OBJECT_ID(N'dbo.Banks'))
    CREATE UNIQUE INDEX [UX_BankName] ON [dbo].[Banks] ([BankName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_TxNumber' AND object_id = OBJECT_ID(N'dbo.BankTransactions'))
    CREATE UNIQUE INDEX [UX_TxNumber] ON [dbo].[BankTransactions] ([TxNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_TxDate' AND object_id = OBJECT_ID(N'dbo.BankTransactions'))
    CREATE INDEX [IX_TxDate] ON [dbo].[BankTransactions] ([TxDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ChequeRef' AND object_id = OBJECT_ID(N'dbo.Cheques'))
    CREATE UNIQUE INDEX [UX_ChequeRef] ON [dbo].[Cheques] ([ChequeRef]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_DueDate' AND object_id = OBJECT_ID(N'dbo.Cheques'))
    CREATE INDEX [IX_DueDate] ON [dbo].[Cheques] ([DueDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ChequeNo' AND object_id = OBJECT_ID(N'dbo.Cheques'))
    CREATE INDEX [IX_ChequeNo] ON [dbo].[Cheques] ([ChequeNo]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_AssetCode' AND object_id = OBJECT_ID(N'dbo.FixedAssets'))
    CREATE UNIQUE INDEX [UX_AssetCode] ON [dbo].[FixedAssets] ([AssetCode]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AssetAccount' AND object_id = OBJECT_ID(N'dbo.FixedAssets'))
    CREATE INDEX [IX_AssetAccount] ON [dbo].[FixedAssets] ([AssetAccount]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_RunNumber' AND object_id = OBJECT_ID(N'dbo.DepreciationRuns'))
    CREATE UNIQUE INDEX [UX_RunNumber] ON [dbo].[DepreciationRuns] ([RunNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_RunMonth' AND object_id = OBJECT_ID(N'dbo.DepreciationRuns'))
    CREATE UNIQUE INDEX [UX_RunMonth] ON [dbo].[DepreciationRuns] ([RunMonth]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_RunID_AssetID' AND object_id = OBJECT_ID(N'dbo.AssetDepreciations'))
    CREATE UNIQUE INDEX [UX_RunID_AssetID] ON [dbo].[AssetDepreciations] ([RunID], [AssetID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AssetID' AND object_id = OBJECT_ID(N'dbo.AssetDepreciations'))
    CREATE INDEX [IX_AssetID] ON [dbo].[AssetDepreciations] ([AssetID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_RevalNumber' AND object_id = OBJECT_ID(N'dbo.AssetRevaluations'))
    CREATE UNIQUE INDEX [UX_RevalNumber] ON [dbo].[AssetRevaluations] ([RevalNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AssetID' AND object_id = OBJECT_ID(N'dbo.AssetRevaluations'))
    CREATE INDEX [IX_AssetID] ON [dbo].[AssetRevaluations] ([AssetID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_TransferNumber' AND object_id = OBJECT_ID(N'dbo.AssetTransfers'))
    CREATE UNIQUE INDEX [UX_TransferNumber] ON [dbo].[AssetTransfers] ([TransferNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AssetID' AND object_id = OBJECT_ID(N'dbo.AssetTransfers'))
    CREATE INDEX [IX_AssetID] ON [dbo].[AssetTransfers] ([AssetID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_CenterCode' AND object_id = OBJECT_ID(N'dbo.CostCenters'))
    CREATE UNIQUE INDEX [UX_CenterCode] ON [dbo].[CostCenters] ([CenterCode]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_CenterName' AND object_id = OBJECT_ID(N'dbo.CostCenters'))
    CREATE UNIQUE INDEX [UX_CenterName] ON [dbo].[CostCenters] ([CenterName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_RepCode' AND object_id = OBJECT_ID(N'dbo.SalesReps'))
    CREATE UNIQUE INDEX [UX_RepCode] ON [dbo].[SalesReps] ([RepCode]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_RepName' AND object_id = OBJECT_ID(N'dbo.SalesReps'))
    CREATE UNIQUE INDEX [UX_RepName] ON [dbo].[SalesReps] ([RepName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_SalesRepID_TargetYear_TargetMonth' AND object_id = OBJECT_ID(N'dbo.SalesRepTargets'))
    CREATE UNIQUE INDEX [UX_SalesRepID_TargetYear_TargetMonth] ON [dbo].[SalesRepTargets] ([SalesRepID], [TargetYear], [TargetMonth]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_RunNumber' AND object_id = OBJECT_ID(N'dbo.CommissionRuns'))
    CREATE UNIQUE INDEX [UX_RunNumber] ON [dbo].[CommissionRuns] ([RunNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_RunMonth' AND object_id = OBJECT_ID(N'dbo.CommissionRuns'))
    CREATE UNIQUE INDEX [UX_RunMonth] ON [dbo].[CommissionRuns] ([RunMonth]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_CommissionRunID_SalesRepID' AND object_id = OBJECT_ID(N'dbo.CommissionLines'))
    CREATE UNIQUE INDEX [UX_CommissionRunID_SalesRepID] ON [dbo].[CommissionLines] ([CommissionRunID], [SalesRepID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_BudgetYear' AND object_id = OBJECT_ID(N'dbo.Budgets'))
    CREATE UNIQUE INDEX [UX_BudgetYear] ON [dbo].[Budgets] ([BudgetYear]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_BudgetID' AND object_id = OBJECT_ID(N'dbo.BudgetLines'))
    CREATE INDEX [IX_BudgetID] ON [dbo].[BudgetLines] ([BudgetID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AccountCode' AND object_id = OBJECT_ID(N'dbo.BudgetLines'))
    CREATE INDEX [IX_AccountCode] ON [dbo].[BudgetLines] ([AccountCode]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_RunNumber' AND object_id = OBJECT_ID(N'dbo.PayrollRuns'))
    CREATE UNIQUE INDEX [UX_RunNumber] ON [dbo].[PayrollRuns] ([RunNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_PayMonth' AND object_id = OBJECT_ID(N'dbo.PayrollRuns'))
    CREATE UNIQUE INDEX [UX_PayMonth] ON [dbo].[PayrollRuns] ([PayMonth]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_PayrollRunID_EmployeeID' AND object_id = OBJECT_ID(N'dbo.PayrollLines'))
    CREATE UNIQUE INDEX [UX_PayrollRunID_EmployeeID] ON [dbo].[PayrollLines] ([PayrollRunID], [EmployeeID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ReconNumber' AND object_id = OBJECT_ID(N'dbo.BankReconciliations'))
    CREATE UNIQUE INDEX [UX_ReconNumber] ON [dbo].[BankReconciliations] ([ReconNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_BankID' AND object_id = OBJECT_ID(N'dbo.BankReconciliations'))
    CREATE INDEX [IX_BankID] ON [dbo].[BankReconciliations] ([BankID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_BankID_SourceType_SourceID' AND object_id = OBJECT_ID(N'dbo.BankClearings'))
    CREATE UNIQUE INDEX [UX_BankID_SourceType_SourceID] ON [dbo].[BankClearings] ([BankID], [SourceType], [SourceID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_PaymentID_SalesInvoiceID' AND object_id = OBJECT_ID(N'dbo.CustomerAllocations'))
    CREATE UNIQUE INDEX [UX_PaymentID_SalesInvoiceID] ON [dbo].[CustomerAllocations] ([PaymentID], [SalesInvoiceID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_SalesInvoiceID' AND object_id = OBJECT_ID(N'dbo.CustomerAllocations'))
    CREATE INDEX [IX_SalesInvoiceID] ON [dbo].[CustomerAllocations] ([SalesInvoiceID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_PaymentID_PurchaseInvoiceID' AND object_id = OBJECT_ID(N'dbo.SupplierAllocations'))
    CREATE UNIQUE INDEX [UX_PaymentID_PurchaseInvoiceID] ON [dbo].[SupplierAllocations] ([PaymentID], [PurchaseInvoiceID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_PurchaseInvoiceID' AND object_id = OBJECT_ID(N'dbo.SupplierAllocations'))
    CREATE INDEX [IX_PurchaseInvoiceID] ON [dbo].[SupplierAllocations] ([PurchaseInvoiceID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ExpenseTypeName' AND object_id = OBJECT_ID(N'dbo.ExpenseTypes'))
    CREATE UNIQUE INDEX [UX_ExpenseTypeName] ON [dbo].[ExpenseTypes] ([ExpenseTypeName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ExpenseNumber' AND object_id = OBJECT_ID(N'dbo.Expenses'))
    CREATE UNIQUE INDEX [UX_ExpenseNumber] ON [dbo].[Expenses] ([ExpenseNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ExpenseDate' AND object_id = OBJECT_ID(N'dbo.Expenses'))
    CREATE INDEX [IX_ExpenseDate] ON [dbo].[Expenses] ([ExpenseDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_RecurringName' AND object_id = OBJECT_ID(N'dbo.RecurringExpenses'))
    CREATE UNIQUE INDEX [UX_RecurringName] ON [dbo].[RecurringExpenses] ([RecurringName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_NextDueDate' AND object_id = OBJECT_ID(N'dbo.RecurringExpenses'))
    CREATE INDEX [IX_NextDueDate] ON [dbo].[RecurringExpenses] ([NextDueDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_VoucherNumber' AND object_id = OBJECT_ID(N'dbo.CashVouchers'))
    CREATE UNIQUE INDEX [UX_VoucherNumber] ON [dbo].[CashVouchers] ([VoucherNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_VoucherDate' AND object_id = OBJECT_ID(N'dbo.CashVouchers'))
    CREATE INDEX [IX_VoucherDate] ON [dbo].[CashVouchers] ([VoucherDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_CashBoxID' AND object_id = OBJECT_ID(N'dbo.CashVouchers'))
    CREATE INDEX [IX_CashBoxID] ON [dbo].[CashVouchers] ([CashBoxID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ClosingNumber' AND object_id = OBJECT_ID(N'dbo.CashClosings'))
    CREATE UNIQUE INDEX [UX_ClosingNumber] ON [dbo].[CashClosings] ([ClosingNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_CashBoxID_ClosingDate' AND object_id = OBJECT_ID(N'dbo.CashClosings'))
    CREATE INDEX [IX_CashBoxID_ClosingDate] ON [dbo].[CashClosings] ([CashBoxID], [ClosingDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ParentCode' AND object_id = OBJECT_ID(N'dbo.Accounts'))
    CREATE INDEX [IX_ParentCode] ON [dbo].[Accounts] ([ParentCode]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_TreeKey' AND object_id = OBJECT_ID(N'dbo.Accounts'))
    CREATE INDEX [IX_TreeKey] ON [dbo].[Accounts] ([TreeKey]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_EntryNumber' AND object_id = OBJECT_ID(N'dbo.JournalEntries'))
    CREATE UNIQUE INDEX [UX_EntryNumber] ON [dbo].[JournalEntries] ([EntryNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_SourceType_SourceID' AND object_id = OBJECT_ID(N'dbo.JournalEntries'))
    CREATE UNIQUE INDEX [UX_SourceType_SourceID] ON [dbo].[JournalEntries] ([SourceType], [SourceID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_EntryDate' AND object_id = OBJECT_ID(N'dbo.JournalEntries'))
    CREATE INDEX [IX_EntryDate] ON [dbo].[JournalEntries] ([EntryDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_EntryID_LineNumber' AND object_id = OBJECT_ID(N'dbo.JournalLines'))
    CREATE UNIQUE INDEX [UX_EntryID_LineNumber] ON [dbo].[JournalLines] ([EntryID], [LineNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AccountCode' AND object_id = OBJECT_ID(N'dbo.JournalLines'))
    CREATE INDEX [IX_AccountCode] ON [dbo].[JournalLines] ([AccountCode]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_PeriodName' AND object_id = OBJECT_ID(N'dbo.FiscalPeriods'))
    CREATE UNIQUE INDEX [UX_PeriodName] ON [dbo].[FiscalPeriods] ([PeriodName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_StartDate' AND object_id = OBJECT_ID(N'dbo.FiscalPeriods'))
    CREATE UNIQUE INDEX [UX_StartDate] ON [dbo].[FiscalPeriods] ([StartDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_EndDate' AND object_id = OBJECT_ID(N'dbo.FiscalPeriods'))
    CREATE UNIQUE INDEX [UX_EndDate] ON [dbo].[FiscalPeriods] ([EndDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_FiscalYear' AND object_id = OBJECT_ID(N'dbo.FiscalYearClosings'))
    CREATE UNIQUE INDEX [UX_FiscalYear] ON [dbo].[FiscalYearClosings] ([FiscalYear]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ClosingNumber' AND object_id = OBJECT_ID(N'dbo.FiscalYearClosings'))
    CREATE UNIQUE INDEX [UX_ClosingNumber] ON [dbo].[FiscalYearClosings] ([ClosingNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_YearClosingID_LineNumber' AND object_id = OBJECT_ID(N'dbo.FiscalYearClosingLines'))
    CREATE UNIQUE INDEX [UX_YearClosingID_LineNumber] ON [dbo].[FiscalYearClosingLines] ([YearClosingID], [LineNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ReturnNumber' AND object_id = OBJECT_ID(N'dbo.VatReturns'))
    CREATE UNIQUE INDEX [UX_ReturnNumber] ON [dbo].[VatReturns] ([ReturnNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_PeriodFrom' AND object_id = OBJECT_ID(N'dbo.VatReturns'))
    CREATE INDEX [IX_PeriodFrom] ON [dbo].[VatReturns] ([PeriodFrom]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_EntryNumber' AND object_id = OBJECT_ID(N'dbo.ManualEntries'))
    CREATE UNIQUE INDEX [UX_EntryNumber] ON [dbo].[ManualEntries] ([EntryNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_EntryDate' AND object_id = OBJECT_ID(N'dbo.ManualEntries'))
    CREATE INDEX [IX_EntryDate] ON [dbo].[ManualEntries] ([EntryDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ManualEntryID_LineNumber' AND object_id = OBJECT_ID(N'dbo.ManualEntryLines'))
    CREATE UNIQUE INDEX [UX_ManualEntryID_LineNumber] ON [dbo].[ManualEntryLines] ([ManualEntryID], [LineNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AccountCode' AND object_id = OBJECT_ID(N'dbo.ManualEntryLines'))
    CREATE INDEX [IX_AccountCode] ON [dbo].[ManualEntryLines] ([AccountCode]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AccountCode' AND object_id = OBJECT_ID(N'dbo.OpeningBalances'))
    CREATE INDEX [IX_AccountCode] ON [dbo].[OpeningBalances] ([AccountCode]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ProductID_WarehouseID' AND object_id = OBJECT_ID(N'dbo.OpeningStock'))
    CREATE UNIQUE INDEX [UX_ProductID_WarehouseID] ON [dbo].[OpeningStock] ([ProductID], [WarehouseID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_TypeCode' AND object_id = OBJECT_ID(N'dbo.TransactionTypes'))
    CREATE UNIQUE INDEX [UX_TypeCode] ON [dbo].[TransactionTypes] ([TypeCode]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_TransactionDate' AND object_id = OBJECT_ID(N'dbo.InventoryTransactions'))
    CREATE INDEX [IX_TransactionDate] ON [dbo].[InventoryTransactions] ([TransactionDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ReferenceType_ReferenceID' AND object_id = OBJECT_ID(N'dbo.InventoryTransactions'))
    CREATE INDEX [IX_ReferenceType_ReferenceID] ON [dbo].[InventoryTransactions] ([ReferenceType], [ReferenceID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_TransferNumber' AND object_id = OBJECT_ID(N'dbo.StockTransfers'))
    CREATE UNIQUE INDEX [UX_TransferNumber] ON [dbo].[StockTransfers] ([TransferNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_TransferDate' AND object_id = OBJECT_ID(N'dbo.StockTransfers'))
    CREATE INDEX [IX_TransferDate] ON [dbo].[StockTransfers] ([TransferDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_StockTransferID' AND object_id = OBJECT_ID(N'dbo.StockTransferLines'))
    CREATE INDEX [IX_StockTransferID] ON [dbo].[StockTransferLines] ([StockTransferID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_PartyName' AND object_id = OBJECT_ID(N'dbo.StockParties'))
    CREATE UNIQUE INDEX [UX_PartyName] ON [dbo].[StockParties] ([PartyName]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_DocNumber' AND object_id = OBJECT_ID(N'dbo.StockDocuments'))
    CREATE UNIQUE INDEX [UX_DocNumber] ON [dbo].[StockDocuments] ([DocNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_DocDate' AND object_id = OBJECT_ID(N'dbo.StockDocuments'))
    CREATE INDEX [IX_DocDate] ON [dbo].[StockDocuments] ([DocDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_StockDocumentID' AND object_id = OBJECT_ID(N'dbo.StockDocumentLines'))
    CREATE INDEX [IX_StockDocumentID] ON [dbo].[StockDocumentLines] ([StockDocumentID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_NoteNumber' AND object_id = OBJECT_ID(N'dbo.StockNotes'))
    CREATE UNIQUE INDEX [UX_NoteNumber] ON [dbo].[StockNotes] ([NoteNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_NoteDate' AND object_id = OBJECT_ID(N'dbo.StockNotes'))
    CREATE INDEX [IX_NoteDate] ON [dbo].[StockNotes] ([NoteDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_StockNoteID' AND object_id = OBJECT_ID(N'dbo.StockNoteLines'))
    CREATE INDEX [IX_StockNoteID] ON [dbo].[StockNoteLines] ([StockNoteID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_OrderNumber' AND object_id = OBJECT_ID(N'dbo.SalesOrders'))
    CREATE UNIQUE INDEX [UX_OrderNumber] ON [dbo].[SalesOrders] ([OrderNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_OrderDate' AND object_id = OBJECT_ID(N'dbo.SalesOrders'))
    CREATE INDEX [IX_OrderDate] ON [dbo].[SalesOrders] ([OrderDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Status' AND object_id = OBJECT_ID(N'dbo.SalesOrders'))
    CREATE INDEX [IX_Status] ON [dbo].[SalesOrders] ([Status]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_QuotationNumber' AND object_id = OBJECT_ID(N'dbo.Quotations'))
    CREATE UNIQUE INDEX [UX_QuotationNumber] ON [dbo].[Quotations] ([QuotationNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_QuoteDate' AND object_id = OBJECT_ID(N'dbo.Quotations'))
    CREATE INDEX [IX_QuoteDate] ON [dbo].[Quotations] ([QuoteDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Status' AND object_id = OBJECT_ID(N'dbo.Quotations'))
    CREATE INDEX [IX_Status] ON [dbo].[Quotations] ([Status]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_QuotationID' AND object_id = OBJECT_ID(N'dbo.QuotationLines'))
    CREATE INDEX [IX_QuotationID] ON [dbo].[QuotationLines] ([QuotationID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_OrderNumber' AND object_id = OBJECT_ID(N'dbo.PurchaseOrders'))
    CREATE UNIQUE INDEX [UX_OrderNumber] ON [dbo].[PurchaseOrders] ([OrderNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_OrderDate' AND object_id = OBJECT_ID(N'dbo.PurchaseOrders'))
    CREATE INDEX [IX_OrderDate] ON [dbo].[PurchaseOrders] ([OrderDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Status' AND object_id = OBJECT_ID(N'dbo.PurchaseOrders'))
    CREATE INDEX [IX_Status] ON [dbo].[PurchaseOrders] ([Status]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_PurchaseOrderID' AND object_id = OBJECT_ID(N'dbo.PurchaseOrderLines'))
    CREATE INDEX [IX_PurchaseOrderID] ON [dbo].[PurchaseOrderLines] ([PurchaseOrderID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_DayNumber' AND object_id = OBJECT_ID(N'dbo.RepDays'))
    CREATE UNIQUE INDEX [UX_DayNumber] ON [dbo].[RepDays] ([DayNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_DayDate' AND object_id = OBJECT_ID(N'dbo.RepDays'))
    CREATE INDEX [IX_DayDate] ON [dbo].[RepDays] ([DayDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Status' AND object_id = OBJECT_ID(N'dbo.RepDays'))
    CREATE INDEX [IX_Status] ON [dbo].[RepDays] ([Status]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_SalesRepID_CustomerID_VisitWeekday' AND object_id = OBJECT_ID(N'dbo.RouteStops'))
    CREATE UNIQUE INDEX [UX_SalesRepID_CustomerID_VisitWeekday] ON [dbo].[RouteStops] ([SalesRepID], [CustomerID], [VisitWeekday]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_VisitDate_SalesRepID_CustomerID' AND object_id = OBJECT_ID(N'dbo.RepVisits'))
    CREATE UNIQUE INDEX [UX_VisitDate_SalesRepID_CustomerID] ON [dbo].[RepVisits] ([VisitDate], [SalesRepID], [CustomerID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_SalesRepID' AND object_id = OBJECT_ID(N'dbo.RepVisits'))
    CREATE INDEX [IX_SalesRepID] ON [dbo].[RepVisits] ([SalesRepID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_SalesRepID_ProductID' AND object_id = OBJECT_ID(N'dbo.RepPrices'))
    CREATE UNIQUE INDEX [UX_SalesRepID_ProductID] ON [dbo].[RepPrices] ([SalesRepID], [ProductID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_SalesOrderID' AND object_id = OBJECT_ID(N'dbo.SalesOrderLines'))
    CREATE INDEX [IX_SalesOrderID] ON [dbo].[SalesOrderLines] ([SalesOrderID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_CountNumber' AND object_id = OBJECT_ID(N'dbo.StockCounts'))
    CREATE UNIQUE INDEX [UX_CountNumber] ON [dbo].[StockCounts] ([CountNumber]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_StockCountID_ProductID' AND object_id = OBJECT_ID(N'dbo.StockCountDetails'))
    CREATE UNIQUE INDEX [UX_StockCountID_ProductID] ON [dbo].[StockCountDetails] ([StockCountID], [ProductID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_LogDate' AND object_id = OBJECT_ID(N'dbo.AuditLog'))
    CREATE INDEX [IX_LogDate] ON [dbo].[AuditLog] ([LogDate]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ActionType' AND object_id = OBJECT_ID(N'dbo.AuditLog'))
    CREATE INDEX [IX_ActionType] ON [dbo].[AuditLog] ([ActionType]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_LogID' AND object_id = OBJECT_ID(N'dbo.AuditChanges'))
    CREATE INDEX [IX_LogID] ON [dbo].[AuditChanges] ([LogID]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_LoggedAt' AND object_id = OBJECT_ID(N'dbo.EInvoiceLog'))
    CREATE INDEX [IX_LoggedAt] ON [dbo].[EInvoiceLog] ([LoggedAt]);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_DocKind_DocID' AND object_id = OBJECT_ID(N'dbo.EInvoiceLog'))
    CREATE INDEX [IX_DocKind_DocID] ON [dbo].[EInvoiceLog] ([DocKind], [DocID]);
IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UX_ProductID' AND object_id = OBJECT_ID(N'dbo.OpeningStock'))
    DROP INDEX [UX_ProductID] ON [dbo].[OpeningStock];
GO

/*------------------------------------------------------------------------------
  5. Seed data
------------------------------------------------------------------------------*/
-- Settings: 1 rows (an empty table)
IF NOT EXISTS (SELECT 1 FROM [dbo].[Settings])
BEGIN
    INSERT INTO [dbo].[Settings] ([SettingID], [StoreName], [CountryCode], [VATRate], [PricesIncludeVAT], [AllowNegativeStock], [CurrencyCode], [DefaultCustomerID], [ZatcaPhase], [LastInvoiceHash], [BackupKeepCount], [SlowMovingDays], [ReceiptFooter]) VALUES (1, N'اسم المحل', N'SA', 0.15, 1, 0, N'SAR', 1, 1, N'NWZlY2ViNjZmZmM4NmYzOGQ5NTI3ODZjNmQ2OTZjNzljMmRiYzIzOWRkNGU5MWI0NjcyOWQ3M2EyN2ZiNTdlOQ==', 30, 90, N'شكرًا لزيارتكم');
END;
-- Sequences: 36 rows (the missing ones)
BEGIN
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'SALES_INVOICE')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'SALES_INVOICE', N'INV-', 1, 6, N'فواتير المبيعات');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'SALES_RETURN')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'SALES_RETURN', N'CRN-', 1, 6, N'مرتجعات المبيعات (إشعار دائن)');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'PURCHASE_INVOICE')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'PURCHASE_INVOICE', N'PUR-', 1, 6, N'فواتير المشتريات');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'PURCHASE_RETURN')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'PURCHASE_RETURN', N'PRT-', 1, 6, N'مرتجعات المشتريات');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'CUSTOMER_PAYMENT')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'CUSTOMER_PAYMENT', N'RCV-', 1, 6, N'سندات القبض من العملاء');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'SUPPLIER_PAYMENT')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'SUPPLIER_PAYMENT', N'PAY-', 1, 6, N'سندات الصرف للموردين');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'EXPENSE')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'EXPENSE', N'EXP-', 1, 6, N'المصروفات');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'SALES_REP')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'SALES_REP', N'REP-', 1, 3, N'أكواد المندوبين');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'COMMISSION_RUN')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'COMMISSION_RUN', N'COM-', 1, 5, N'مسيرات العمولات');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'STOCK_COUNT')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'STOCK_COUNT', N'CNT-', 1, 5, N'جلسات الجرد');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'STOCK_ADJUST')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'STOCK_ADJUST', N'ADJ-', 1, 6, N'حركات المخزون اليدوية');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'STOCK_TRANSFER')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'STOCK_TRANSFER', N'STT-', 1, 6, N'التحويل بين المخازن');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'STOCK_ISSUE')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'STOCK_ISSUE', N'SIS-', 1, 6, N'سندات الصرف المخزني');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'STOCK_RECEIPT')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'STOCK_RECEIPT', N'SRC-', 1, 6, N'سندات الاستلام المخزني');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'DELIVERY_NOTE')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'DELIVERY_NOTE', N'DLV-', 1, 6, N'أذونات التسليم للعملاء');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'RECEIPT_NOTE')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'RECEIPT_NOTE', N'GRN-', 1, 6, N'أذونات الاستلام من الموردين');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'SALES_ORDER')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'SALES_ORDER', N'SO-', 1, 6, N'طلبيات العملاء');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'PURCHASE_ORDER')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'PURCHASE_ORDER', N'PO-', 1, 6, N'أوامر الشراء');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'QUOTATION')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'QUOTATION', N'QT-', 1, 6, N'عروض الأسعار للعملاء');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'REP_DAY')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'REP_DAY', N'RD-', 1, 6, N'أيام المندوبين (البيع من السيارة)');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'PRODUCT_CODE')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'PRODUCT_CODE', N'P', 1, 5, N'أكواد المنتجات');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'ZATCA_ICV')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'ZATCA_ICV', NULL, 1, 0, N'عدّاد ICV لمستندات فاتورة');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'CASH_IN')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'CASH_IN', N'CIN-', 1, 6, N'سندات قبض النقدية (الخزينة)');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'CASH_OUT')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'CASH_OUT', N'COT-', 1, 6, N'سندات صرف النقدية (الخزينة)');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'CASH_TRANSFER')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'CASH_TRANSFER', N'TRF-', 1, 6, N'التحويل بين الصناديق');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'CASH_CLOSING')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'CASH_CLOSING', N'CLS-', 1, 6, N'تصفية يومية الكاشير');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'JOURNAL')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'JOURNAL', N'JV-', 1, 6, N'قيود اليومية');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'MANUAL_ENTRY')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'MANUAL_ENTRY', N'MJ-', 1, 6, N'القيود اليدوية');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'BANK_TX')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'BANK_TX', N'BNK-', 1, 6, N'الحركات البنكية');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'BANK_RECON')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'BANK_RECON', N'REC-', 1, 6, N'التسويات البنكية');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'CHEQUE')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'CHEQUE', N'CHQ-', 1, 6, N'الشيكات الواردة والصادرة');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'FIXED_ASSET')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'FIXED_ASSET', N'FA-', 1, 5, N'الأصول الثابتة');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'PAYROLL')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'PAYROLL', N'PAY-', 1, 5, N'مسيرات الرواتب');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'DEPRECIATION')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'DEPRECIATION', N'DEP-', 1, 6, N'قيود الإهلاك الشهرية');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'ASSET_REVAL')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'ASSET_REVAL', N'AR-', 1, 5, N'إعادة تقييم الأصول الثابتة');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Sequences] WHERE [SequenceName] = N'ASSET_TRANSFER')
        INSERT INTO [dbo].[Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES (N'ASSET_TRANSFER', N'AT-', 1, 5, N'نقل الأصول الثابتة بين مراكز التكلفة');
END;
-- Roles: 3 rows (an empty table)
IF NOT EXISTS (SELECT 1 FROM [dbo].[Roles])
BEGIN
    INSERT INTO [dbo].[Roles] ([RoleID], [RoleCode], [RoleName], [Description]) VALUES (1, N'ADMIN', N'مدير النظام', N'جميع الصلاحيات');
    INSERT INTO [dbo].[Roles] ([RoleID], [RoleCode], [RoleName], [Description]) VALUES (2, N'MANAGER', N'مدير', N'المبيعات والمشتريات والمخزون والتقارير');
    INSERT INTO [dbo].[Roles] ([RoleID], [RoleCode], [RoleName], [Description]) VALUES (3, N'CASHIER', N'كاشير', N'المبيعات والعملاء فقط');
END;
-- Permissions: 48 rows (the missing ones)
BEGIN
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'SALES_POS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'SALES_POS', N'نقطة البيع', N'المبيعات', 10);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SALES_POS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SALES_POS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, N'SALES_POS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'SALES_VIEW')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'SALES_VIEW', N'عرض وإعادة طباعة الفواتير', N'المبيعات', 11);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SALES_VIEW');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SALES_VIEW');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, N'SALES_VIEW');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'SALES_RETURN')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'SALES_RETURN', N'مرتجعات المبيعات', N'المبيعات', 12);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SALES_RETURN');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SALES_RETURN');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'PRICE_OVERRIDE')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'PRICE_OVERRIDE', N'تعديل سعر البيع في الفاتورة', N'المبيعات', 13);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PRICE_OVERRIDE');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'PRICE_OVERRIDE');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'DISCOUNT_OVERRIDE')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'DISCOUNT_OVERRIDE', N'خصم أعلى من الحد المسموح', N'المبيعات', 14);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'DISCOUNT_OVERRIDE');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'DISCOUNT_OVERRIDE');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'ALLOW_NEGATIVE_STOCK')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'ALLOW_NEGATIVE_STOCK', N'البيع بكمية أكبر من المتوفر', N'المبيعات', 15);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'ALLOW_NEGATIVE_STOCK');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'SALES_ORDERS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'SALES_ORDERS', N'طلبيات العملاء: التسجيل والاعتماد والتجهيز والإلغاء', N'المبيعات', 16);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SALES_ORDERS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SALES_ORDERS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'REP_DAYS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'REP_DAYS', N'يوم المندوب: تحميل السيارة والبيع باسمه وتسويته', N'المبيعات', 17);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'REP_DAYS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'REP_DAYS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'REP_ROUTES')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'REP_ROUTES', N'خط السير وزيارات المندوبين وأسعارهم', N'المبيعات', 18);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'REP_ROUTES');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'REP_ROUTES');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'QUOTATIONS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'QUOTATIONS', N'عروض الأسعار للعملاء: التسجيل والإصدار والتحويل', N'المبيعات', 19);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'QUOTATIONS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'QUOTATIONS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'CUSTOMERS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'CUSTOMERS', N'إدارة العملاء', N'العملاء', 20);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CUSTOMERS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CUSTOMERS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, N'CUSTOMERS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'CUSTOMER_PAYMENTS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'CUSTOMER_PAYMENTS', N'سندات القبض', N'العملاء', 21);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CUSTOMER_PAYMENTS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CUSTOMER_PAYMENTS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, N'CUSTOMER_PAYMENTS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'PURCHASES')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'PURCHASES', N'فواتير المشتريات', N'المشتريات', 30);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PURCHASES');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'PURCHASES');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'PURCHASE_RETURN')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'PURCHASE_RETURN', N'مرتجعات المشتريات', N'المشتريات', 31);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PURCHASE_RETURN');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'PURCHASE_RETURN');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'PURCHASE_ORDERS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'PURCHASE_ORDERS', N'أوامر الشراء واقتراح الشراء', N'المشتريات', 32);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PURCHASE_ORDERS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'PURCHASE_ORDERS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'SUPPLIERS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'SUPPLIERS', N'إدارة الموردين', N'الموردون', 40);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SUPPLIERS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SUPPLIERS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'SUPPLIER_PAYMENTS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'SUPPLIER_PAYMENTS', N'سندات الصرف', N'الموردون', 41);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SUPPLIER_PAYMENTS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SUPPLIER_PAYMENTS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'PRODUCTS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'PRODUCTS', N'إدارة المنتجات والأسعار', N'المخزون', 50);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PRODUCTS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'PRODUCTS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'INVENTORY_ADJUST')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'INVENTORY_ADJUST', N'إضافة وخصم مخزون يدوي', N'المخزون', 51);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'INVENTORY_ADJUST');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'INVENTORY_ADJUST');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'STOCK_COUNT')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'STOCK_COUNT', N'الجرد', N'المخزون', 52);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'STOCK_COUNT');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'STOCK_COUNT');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'STOCK_TRANSFER')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'STOCK_TRANSFER', N'التحويل بين المخازن: الإرسال والاستلام', N'المخزون', 53);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'STOCK_TRANSFER');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'STOCK_TRANSFER');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'STOCK_DOCS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'STOCK_DOCS', N'سندات الصرف والاستلام المخزني وجهاتها', N'المخزون', 54);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'STOCK_DOCS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'STOCK_DOCS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'STOCK_NOTES')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'STOCK_NOTES', N'أذونات التسليم للعملاء والاستلام من الموردين', N'المخزون', 55);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'STOCK_NOTES');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'STOCK_NOTES');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'EXPENSES')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'EXPENSES', N'المصروفات', N'المصروفات', 60);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'EXPENSES');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'EXPENSES');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'CASH_BOX')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'CASH_BOX', N'الخزينة: سندات القبض والصرف والتحويل والصناديق', N'الخزينة', 65);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CASH_BOX');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CASH_BOX');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'CASH_CLOSING')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'CASH_CLOSING', N'تصفية يومية الكاشير', N'الخزينة', 66);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CASH_CLOSING');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CASH_CLOSING');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, N'CASH_CLOSING');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'CASH_ALL_BOXES')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'CASH_ALL_BOXES', N'الصرف والتحويل من كل الصناديق (لا من صناديقه فقط)', N'الخزينة', 69);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CASH_ALL_BOXES');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CASH_ALL_BOXES');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'JOURNAL')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'JOURNAL', N'قيود اليومية ودليل الحسابات وميزان المراجعة', N'الحسابات', 75);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'JOURNAL');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'JOURNAL');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'MANUAL_ENTRY')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'MANUAL_ENTRY', N'القيود اليدوية: إضافة وتعديل وحذف', N'الحسابات', 76);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'MANUAL_ENTRY');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'MANUAL_ENTRY');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'PERIOD_CLOSE')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'PERIOD_CLOSE', N'إقفال الفترات والسنة المالية وإعادة فتحها', N'الحسابات', 77);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PERIOD_CLOSE');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'VAT_RETURN')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'VAT_RETURN', N'إقرار ضريبة القيمة المضافة: الاعتماد والسداد', N'الحسابات', 78);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'VAT_RETURN');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'VAT_RETURN');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'BANKS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'BANKS', N'البنوك: الحسابات والحركات البنكية والتسوية البنكية', N'الخزينة', 67);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'BANKS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'BANKS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'CHEQUES')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'CHEQUES', N'الشيكات الواردة والصادرة: التسجيل والتحصيل والارتداد', N'الخزينة', 68);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CHEQUES');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CHEQUES');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'FIXED_ASSETS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'FIXED_ASSETS', N'الأصول الثابتة والإهلاك', N'الحسابات', 79);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'FIXED_ASSETS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'FIXED_ASSETS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'PAYROLL')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'PAYROLL', N'مسير الرواتب: الإعداد والترحيل والصرف', N'الحسابات', 80);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PAYROLL');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'PAYROLL');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'BUDGET')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'BUDGET', N'الموازنة التقديرية: الإعداد والمقارنة بالفعلي', N'الحسابات', 81);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'BUDGET');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'BUDGET');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'CURRENCIES')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'CURRENCIES', N'العملات وأسعارها', N'الحسابات', 84);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CURRENCIES');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CURRENCIES');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'OPENING_BALANCES')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'OPENING_BALANCES', N'أرصدة بداية التشغيل: تاريخها وأرصدة الحسابات والعملاء والموردين والصناديق والبنوك والمخزون', N'الحسابات', 85);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'OPENING_BALANCES');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'OPENING_BALANCES');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'SALES_REPS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'SALES_REPS', N'المندوبين: البيانات والأهداف والعمولات وتقاريرهم', N'المبيعات', 15);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SALES_REPS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SALES_REPS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'REPORTS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'REPORTS', N'التقارير التشغيلية', N'التقارير', 70);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'REPORTS');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'REPORTS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'REPORTS_PROFIT')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'REPORTS_PROFIT', N'تقارير الأرباح والضريبة', N'التقارير', 71);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'REPORTS_PROFIT');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'REPORTS_PROFIT');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'DASHBOARD_FINANCIAL')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'DASHBOARD_FINANCIAL', N'الأرقام المالية في لوحة التحكم', N'التقارير', 72);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'DASHBOARD_FINANCIAL');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'DASHBOARD_FINANCIAL');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'SETTINGS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'SETTINGS', N'إعدادات المحل', N'النظام', 80);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SETTINGS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'USERS')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'USERS', N'المستخدمون والصلاحيات', N'النظام', 81);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'USERS');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'BACKUP')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'BACKUP', N'النسخ الاحتياطي', N'النظام', 82);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'BACKUP');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'SQL_SERVER')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'SQL_SERVER', N'قاعدة SQL Server: الاتصال ونقل البيانات وتشغيل البرنامج على الخادم', N'النظام', 84);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SQL_SERVER');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'AUDIT_LOG')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'AUDIT_LOG', N'سجل التدقيق: من أضاف أو عدّل أو حذف، والقيم قبل وبعد', N'النظام', 83);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'AUDIT_LOG');
        END
    END
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Permissions] WHERE [PermissionKey] = N'EINVOICE')
    BEGIN
        INSERT INTO [dbo].[Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES (N'EINVOICE', N'الفاتورة الإلكترونية: المتابعة وإعادة الإرسال', N'المبيعات', 16);
        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
        BEGIN
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'EINVOICE');
            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'EINVOICE');
        END
    END
END;
-- RolePermissions: 94 rows (an empty table)
IF NOT EXISTS (SELECT 1 FROM [dbo].[RolePermissions])
BEGIN
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SALES_POS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SALES_VIEW');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SALES_RETURN');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PRICE_OVERRIDE');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'DISCOUNT_OVERRIDE');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'ALLOW_NEGATIVE_STOCK');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SALES_ORDERS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'REP_DAYS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'REP_ROUTES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'QUOTATIONS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CUSTOMERS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CUSTOMER_PAYMENTS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PURCHASES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PURCHASE_RETURN');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PURCHASE_ORDERS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SUPPLIERS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SUPPLIER_PAYMENTS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PRODUCTS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'INVENTORY_ADJUST');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'STOCK_COUNT');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'STOCK_TRANSFER');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'STOCK_DOCS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'STOCK_NOTES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'EXPENSES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CASH_BOX');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CASH_CLOSING');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CASH_ALL_BOXES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'JOURNAL');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'MANUAL_ENTRY');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PERIOD_CLOSE');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'VAT_RETURN');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'BANKS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CHEQUES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'FIXED_ASSETS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'PAYROLL');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'BUDGET');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'CURRENCIES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'OPENING_BALANCES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SALES_REPS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'REPORTS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'REPORTS_PROFIT');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'DASHBOARD_FINANCIAL');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SETTINGS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'USERS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'BACKUP');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'SQL_SERVER');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'AUDIT_LOG');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, N'EINVOICE');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SALES_POS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SALES_VIEW');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SALES_RETURN');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'PRICE_OVERRIDE');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'DISCOUNT_OVERRIDE');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SALES_ORDERS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'REP_DAYS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'REP_ROUTES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'QUOTATIONS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CUSTOMERS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CUSTOMER_PAYMENTS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'PURCHASES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'PURCHASE_RETURN');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'PURCHASE_ORDERS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SUPPLIERS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SUPPLIER_PAYMENTS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'PRODUCTS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'INVENTORY_ADJUST');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'STOCK_COUNT');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'STOCK_TRANSFER');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'STOCK_DOCS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'STOCK_NOTES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'EXPENSES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CASH_BOX');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CASH_CLOSING');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CASH_ALL_BOXES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'JOURNAL');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'MANUAL_ENTRY');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'VAT_RETURN');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'BANKS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CHEQUES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'FIXED_ASSETS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'PAYROLL');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'BUDGET');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'CURRENCIES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'OPENING_BALANCES');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'SALES_REPS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'REPORTS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'REPORTS_PROFIT');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'DASHBOARD_FINANCIAL');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, N'EINVOICE');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, N'SALES_POS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, N'SALES_VIEW');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, N'CUSTOMERS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, N'CUSTOMER_PAYMENTS');
    INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, N'CASH_CLOSING');
END;
-- Employees: 1 rows (an empty table)
SET IDENTITY_INSERT [dbo].[Employees] ON;
IF NOT EXISTS (SELECT 1 FROM [dbo].[Employees])
BEGIN
    INSERT INTO [dbo].[Employees] ([EmployeeID], [EmployeeName], [JobTitle], [Username], [RoleID], [MaxDiscountPercent], [MustChangePassword], [IsActive]) VALUES (1, N'مدير النظام', N'مدير النظام', N'admin', 1, 1, 1, 1);
END;
SET IDENTITY_INSERT [dbo].[Employees] OFF;
-- Screens: 83 rows (the missing ones)
BEGIN
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmPOS')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmPOS', N'نقطة البيع (المحلات)', N'المبيعات', 10, N'SALES_POS', 1, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmTouchPOS')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmTouchPOS', N'نقطة البيع (المطاعم)', N'المبيعات', 20, N'SALES_POS', 1, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCafePOS')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCafePOS', N'نقطة البيع (الكافيهات)', N'المبيعات', 30, N'SALES_POS', 1, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmSalesInvoice')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmSalesInvoice', N'عرض الفواتير وإعادة طباعتها', N'المبيعات', 40, N'SALES_VIEW', 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmSalesReturn')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmSalesReturn', N'مرتجعات المبيعات', N'المبيعات', 50, N'SALES_RETURN', 1, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCustomers')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCustomers', N'العملاء', N'العملاء', 60, N'CUSTOMERS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCustomerPayment')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCustomerPayment', N'سندات القبض من العملاء', N'العملاء', 70, N'CUSTOMER_PAYMENTS', 1, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmPurchaseInvoice')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmPurchaseInvoice', N'فواتير المشتريات', N'المشتريات', 80, N'PURCHASES', 1, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmPurchaseView')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmPurchaseView', N'عرض فواتير المشتريات', N'المشتريات', 90, N'PURCHASES', 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmPurchaseReturn')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmPurchaseReturn', N'مرتجعات المشتريات', N'المشتريات', 100, N'PURCHASE_RETURN', 1, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmSuppliers')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmSuppliers', N'الموردون', N'الموردون', 110, N'SUPPLIERS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmSupplierPayment')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmSupplierPayment', N'سندات الصرف للموردين', N'الموردون', 120, N'SUPPLIER_PAYMENTS', 1, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmProducts')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmProducts', N'المنتجات والأسعار', N'المخزون', 130, N'PRODUCTS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCategories')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCategories', N'التصنيفات', N'المخزون', 140, N'PRODUCTS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmUnits')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmUnits', N'الوحدات', N'المخزون', 150, N'PRODUCTS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmProductUnits')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmProductUnits', N'وحدات الأصناف', N'المخزون', 160, N'PRODUCTS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmPriceLists')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmPriceLists', N'قوائم الأسعار', N'المبيعات', 170, N'PRODUCTS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmPriceListItems')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmPriceListItems', N'أسعار قوائم الأسعار', N'المبيعات', 180, N'PRODUCTS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmInventory')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmInventory', N'المخزون والحركات اليدوية', N'المخزون', 190, N'PRODUCTS', 1, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmStockCount')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmStockCount', N'الجرد', N'المخزون', 200, N'STOCK_COUNT', 1, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmBarcodeLabels')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmBarcodeLabels', N'ملصقات الباركود', N'المخزون', 210, N'PRODUCTS', 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmLabelSettings')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmLabelSettings', N'إعدادات الملصقات', N'المخزون', 220, N'PRODUCTS', 0, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmExpenses')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmExpenses', N'المصروفات', N'المصروفات', 230, N'EXPENSES', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmExpenseTypes')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmExpenseTypes', N'أنواع المصروفات', N'المصروفات', 240, N'EXPENSES', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmRecurring')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmRecurring', N'المصروفات المتكررة', N'المصروفات', 250, N'EXPENSES', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmTreasury')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmTreasury', N'الخزينة', N'الخزينة', 260, N'CASH_CLOSING', 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCashVoucher')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCashVoucher', N'سندات النقدية والتحويل', N'الخزينة', 270, N'CASH_BOX', 1, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCashClosing')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCashClosing', N'تصفية يومية الكاشير', N'الخزينة', 280, N'CASH_CLOSING', 1, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCashBoxes')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCashBoxes', N'الصناديق', N'الخزينة', 290, N'CASH_BOX', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmWarehouses')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmWarehouses', N'المخازن', N'المخزون', 300, N'PRODUCTS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmStockTransfer')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmStockTransfer', N'التحويل بين المخازن', N'المخزون', 310, N'STOCK_TRANSFER', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmStockParties')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmStockParties', N'جهات الصرف والاستلام المخزني', N'المخزون', 320, N'STOCK_DOCS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmStockDocument')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmStockDocument', N'سندات الصرف والاستلام المخزني', N'المخزون', 330, N'STOCK_DOCS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmStockNote')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmStockNote', N'أذونات التسليم والاستلام', N'المخزون', 340, N'STOCK_NOTES', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmSalesOrder')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmSalesOrder', N'طلبيات العملاء', N'المبيعات', 350, N'SALES_ORDERS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmPurchaseOrder')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmPurchaseOrder', N'أوامر الشراء', N'المشتريات', 360, N'PURCHASE_ORDERS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmQuotation')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmQuotation', N'عروض الأسعار للعملاء', N'المبيعات', 370, N'QUOTATIONS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmRepDay')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmRepDay', N'يوم المندوب والبيع من السيارة', N'المبيعات', 380, N'REP_DAYS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmRepVisits')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmRepVisits', N'زيارات المندوبين وخطة اليوم', N'المبيعات', 390, N'REP_ROUTES', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmRouteStops')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmRouteStops', N'خط سير المندوبين', N'المبيعات', 400, N'REP_ROUTES', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmRepPrices')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmRepPrices', N'أسعار المندوبين', N'المبيعات', 410, N'REP_ROUTES', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCashBoxUsers')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCashBoxUsers', N'أمناء الصناديق', N'الخزينة', 420, N'CASH_ALL_BOXES', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCashTransfers')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCashTransfers', N'استلام التحويلات بين الصناديق', N'الخزينة', 430, N'CASH_CLOSING', 0, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmAccounting')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmAccounting', N'المحاسبة والمالية', N'الحسابات', 440, NULL, 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmJournal')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmJournal', N'قيود اليومية', N'الحسابات', 450, N'JOURNAL', 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmAccounts')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmAccounts', N'دليل الحسابات', N'الحسابات', 460, N'JOURNAL', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmManualEntry')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmManualEntry', N'القيود اليدوية', N'الحسابات', 470, N'MANUAL_ENTRY', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmLedger')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmLedger', N'كشف حساب ودفتر الأستاذ', N'الحسابات', 480, N'JOURNAL', 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmFinancials')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmFinancials', N'القوائم المالية', N'الحسابات', 490, N'REPORTS_PROFIT', 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmPeriodClosing')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmPeriodClosing', N'إقفال الفترات والسنة المالية', N'الحسابات', 500, N'PERIOD_CLOSE', 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmFiscalPeriods')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmFiscalPeriods', N'الفترات المالية', N'الحسابات', 510, N'PERIOD_CLOSE', 1, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmVatReturn')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmVatReturn', N'إقرار ضريبة القيمة المضافة', N'الحسابات', 520, N'VAT_RETURN', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmOpeningBalances')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmOpeningBalances', N'أرصدة بداية التشغيل', N'الحسابات', 530, N'OPENING_BALANCES', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmAging')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmAging', N'أعمار الديون (العملاء والموردون)', N'التقارير', 540, N'REPORTS', 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmBanks')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmBanks', N'البنوك', N'الخزينة', 550, N'BANKS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmBankTx')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmBankTx', N'الحركات البنكية', N'الخزينة', 560, N'BANKS', 1, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmBankRecon')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmBankRecon', N'التسوية البنكية', N'الخزينة', 570, N'BANKS', 1, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCheques')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCheques', N'الشيكات الواردة والصادرة', N'الخزينة', 580, N'CHEQUES', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmAssets')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmAssets', N'الأصول الثابتة', N'الحسابات', 590, N'FIXED_ASSETS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmDepreciation')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmDepreciation', N'الإهلاك الشهري', N'الحسابات', 600, N'FIXED_ASSETS', 1, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmAssetChanges')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmAssetChanges', N'إعادة تقييم الأصول ونقلها', N'الحسابات', 610, N'FIXED_ASSETS', 1, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmPayroll')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmPayroll', N'مسير الرواتب', N'الحسابات', 620, N'PAYROLL', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCostCenters')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCostCenters', N'مراكز التكلفة والفروع', N'الحسابات', 630, N'JOURNAL', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmBudget')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmBudget', N'الموازنة التقديرية', N'الحسابات', 640, N'BUDGET', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCurrencies')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCurrencies', N'العملات', N'الحسابات', 650, N'CURRENCIES', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmSalesReps')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmSalesReps', N'المندوبين', N'المبيعات', 660, N'SALES_REPS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmRepTargets')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmRepTargets', N'أهداف المندوبين', N'المبيعات', 670, N'SALES_REPS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCommissions')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCommissions', N'عمولات المندوبين', N'المبيعات', 680, N'SALES_REPS', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmCurrencyRates')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmCurrencyRates', N'أسعار العملات', N'الحسابات', 690, N'CURRENCIES', 1, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmAllocation')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmAllocation', N'ربط السداد بالفواتير', N'العملاء', 700, N'CUSTOMER_PAYMENTS', 1, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmReportCenter')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmReportCenter', N'التقارير', N'التقارير', 710, N'REPORTS', 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmSearch')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmSearch', N'البحث', N'النظام', 720, NULL, 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmSettings')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmSettings', N'إعدادات المحل', N'النظام', 730, N'SETTINGS', 0, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmUsers')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmUsers', N'المستخدمون', N'النظام', 740, N'USERS', 1, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmRoles')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmRoles', N'الأدوار والصلاحيات', N'النظام', 750, N'USERS', 0, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmUserScreens')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmUserScreens', N'صلاحيات الشاشات للمستخدمين', N'النظام', 760, N'USERS', 0, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmAuditLog')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmAuditLog', N'سجل التدقيق', N'النظام', 770, N'AUDIT_LOG', 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmBackup')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmBackup', N'النسخ الاحتياطي', N'النظام', 780, N'BACKUP', 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmSqlServer')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmSqlServer', N'قاعدة البيانات على SQL Server', N'النظام', 790, N'SQL_SERVER', 0, 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmEnglishNames')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmEnglishNames', N'الأسماء الإنجليزية', N'النظام', 800, N'SETTINGS', 0, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmEInvoices')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmEInvoices', N'الفاتورة الإلكترونية', N'المبيعات', 810, N'EINVOICE', 0, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmZatcaSetup')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmZatcaSetup', N'إعداد ربط منصة فاتورة', N'النظام', 820, N'SETTINGS', 0, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Screens] WHERE [ScreenName] = N'frmEtaSetup')
        INSERT INTO [dbo].[Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES (N'frmEtaSetup', N'إعداد ربط منظومة الإيصال الإلكتروني', N'النظام', 830, N'SETTINGS', 0, 1, 0);
END;
-- Categories: 1 rows (an empty table)
SET IDENTITY_INSERT [dbo].[Categories] ON;
IF NOT EXISTS (SELECT 1 FROM [dbo].[Categories])
BEGIN
    INSERT INTO [dbo].[Categories] ([CategoryID], [CategoryName], [Description]) VALUES (1, N'عام', N'تصنيف افتراضي');
END;
SET IDENTITY_INSERT [dbo].[Categories] OFF;
-- Units: 8 rows (an empty table)
SET IDENTITY_INSERT [dbo].[Units] ON;
IF NOT EXISTS (SELECT 1 FROM [dbo].[Units])
BEGIN
    INSERT INTO [dbo].[Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (1, N'حبة', N'PCE');
    INSERT INTO [dbo].[Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (2, N'علبة', N'BX');
    INSERT INTO [dbo].[Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (3, N'كرتون', N'CT');
    INSERT INTO [dbo].[Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (4, N'باكيت', N'PK');
    INSERT INTO [dbo].[Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (5, N'كيلو', N'KGM');
    INSERT INTO [dbo].[Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (6, N'لتر', N'LTR');
    INSERT INTO [dbo].[Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (7, N'متر', N'MTR');
    INSERT INTO [dbo].[Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (8, N'طقم', N'SET');
END;
SET IDENTITY_INSERT [dbo].[Units] OFF;
-- PaymentMethods: 4 rows (an empty table)
IF NOT EXISTS (SELECT 1 FROM [dbo].[PaymentMethods])
BEGIN
    INSERT INTO [dbo].[PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (1, N'نقدي', N'10', 1);
    INSERT INTO [dbo].[PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (2, N'مدى / بطاقة', N'48', 2);
    INSERT INTO [dbo].[PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (3, N'تحويل بنكي', N'42', 3);
    INSERT INTO [dbo].[PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (4, N'محفظة إلكترونية', N'1', 4);
END;
-- Currencies: 11 rows (the missing ones)
BEGIN
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Currencies] WHERE [CurrencyCode] = N'SAR')
        INSERT INTO [dbo].[Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES (N'SAR', N'ريال سعودي', N'Saudi Riyal', N'ر.س', 2, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Currencies] WHERE [CurrencyCode] = N'USD')
        INSERT INTO [dbo].[Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES (N'USD', N'دولار أمريكي', N'US Dollar', N'$', 2, 2);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Currencies] WHERE [CurrencyCode] = N'EUR')
        INSERT INTO [dbo].[Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES (N'EUR', N'يورو', N'Euro', N'€', 2, 3);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Currencies] WHERE [CurrencyCode] = N'AED')
        INSERT INTO [dbo].[Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES (N'AED', N'درهم إماراتي', N'UAE Dirham', N'د.إ', 2, 4);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Currencies] WHERE [CurrencyCode] = N'KWD')
        INSERT INTO [dbo].[Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES (N'KWD', N'دينار كويتي', N'Kuwaiti Dinar', N'د.ك', 3, 5);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Currencies] WHERE [CurrencyCode] = N'BHD')
        INSERT INTO [dbo].[Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES (N'BHD', N'دينار بحريني', N'Bahraini Dinar', N'د.ب', 3, 6);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Currencies] WHERE [CurrencyCode] = N'QAR')
        INSERT INTO [dbo].[Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES (N'QAR', N'ريال قطري', N'Qatari Riyal', N'ر.ق', 2, 7);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Currencies] WHERE [CurrencyCode] = N'OMR')
        INSERT INTO [dbo].[Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES (N'OMR', N'ريال عماني', N'Omani Rial', N'ر.ع', 3, 8);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Currencies] WHERE [CurrencyCode] = N'EGP')
        INSERT INTO [dbo].[Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES (N'EGP', N'جنيه مصري', N'Egyptian Pound', N'ج.م', 2, 9);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Currencies] WHERE [CurrencyCode] = N'GBP')
        INSERT INTO [dbo].[Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES (N'GBP', N'جنيه إسترليني', N'Pound Sterling', N'£', 2, 10);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Currencies] WHERE [CurrencyCode] = N'CNY')
        INSERT INTO [dbo].[Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES (N'CNY', N'يوان صيني', N'Chinese Yuan', N'¥', 2, 11);
END;
-- CurrencyRates: 5 rows (the missing ones)
SET IDENTITY_INSERT [dbo].[CurrencyRates] ON;
BEGIN
    IF NOT EXISTS (SELECT 1 FROM [dbo].[CurrencyRates] WHERE [CurrencyRateID] = 1)
        INSERT INTO [dbo].[CurrencyRates] ([CurrencyRateID], [CurrencyCode], [RateDate], [Rate], [Notes]) VALUES (1, N'USD', '20000101', 3.75, N'سعر الربط الرسمي');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[CurrencyRates] WHERE [CurrencyRateID] = 2)
        INSERT INTO [dbo].[CurrencyRates] ([CurrencyRateID], [CurrencyCode], [RateDate], [Rate], [Notes]) VALUES (2, N'AED', '20000101', 1.0211, N'مرتبط بالدولار');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[CurrencyRates] WHERE [CurrencyRateID] = 3)
        INSERT INTO [dbo].[CurrencyRates] ([CurrencyRateID], [CurrencyCode], [RateDate], [Rate], [Notes]) VALUES (3, N'BHD', '20000101', 9.9734, N'مرتبط بالدولار');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[CurrencyRates] WHERE [CurrencyRateID] = 4)
        INSERT INTO [dbo].[CurrencyRates] ([CurrencyRateID], [CurrencyCode], [RateDate], [Rate], [Notes]) VALUES (4, N'QAR', '20000101', 1.0302, N'مرتبط بالدولار');
    IF NOT EXISTS (SELECT 1 FROM [dbo].[CurrencyRates] WHERE [CurrencyRateID] = 5)
        INSERT INTO [dbo].[CurrencyRates] ([CurrencyRateID], [CurrencyCode], [RateDate], [Rate], [Notes]) VALUES (5, N'OMR', '20000101', 9.7529, N'مرتبط بالدولار');
END;
SET IDENTITY_INSERT [dbo].[CurrencyRates] OFF;
-- CashBoxes: 2 rows (an empty table)
SET IDENTITY_INSERT [dbo].[CashBoxes] ON;
IF NOT EXISTS (SELECT 1 FROM [dbo].[CashBoxes])
BEGIN
    INSERT INTO [dbo].[CashBoxes] ([CashBoxID], [BoxName], [BoxType]) VALUES (1, N'الخزينة الرئيسية', N'MAIN');
    INSERT INTO [dbo].[CashBoxes] ([CashBoxID], [BoxName], [BoxType]) VALUES (2, N'صندوق الكاشير', N'CASHIER');
END;
SET IDENTITY_INSERT [dbo].[CashBoxes] OFF;
-- Warehouses: 1 rows (an empty table)
SET IDENTITY_INSERT [dbo].[Warehouses] ON;
IF NOT EXISTS (SELECT 1 FROM [dbo].[Warehouses])
BEGIN
    INSERT INTO [dbo].[Warehouses] ([WarehouseID], [WarehouseCode], [WarehouseName], [WarehouseType]) VALUES (1, N'WH-1', N'المخزن الرئيسي', N'MAIN');
END;
SET IDENTITY_INSERT [dbo].[Warehouses] OFF;
-- Customers: 1 rows (an empty table)
SET IDENTITY_INSERT [dbo].[Customers] ON;
IF NOT EXISTS (SELECT 1 FROM [dbo].[Customers])
BEGIN
    INSERT INTO [dbo].[Customers] ([CustomerID], [CustomerName], [AllowCredit], [IsSystem]) VALUES (1, N'عميل نقدي', 0, 1);
END;
SET IDENTITY_INSERT [dbo].[Customers] OFF;
-- ExpenseTypes: 9 rows (an empty table)
SET IDENTITY_INSERT [dbo].[ExpenseTypes] ON;
IF NOT EXISTS (SELECT 1 FROM [dbo].[ExpenseTypes])
BEGIN
    INSERT INTO [dbo].[ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (1, N'الإيجار');
    INSERT INTO [dbo].[ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (2, N'الكهرباء');
    INSERT INTO [dbo].[ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (3, N'المياه');
    INSERT INTO [dbo].[ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (4, N'الإنترنت والاتصالات');
    INSERT INTO [dbo].[ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (5, N'النقل');
    INSERT INTO [dbo].[ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (6, N'الصيانة');
    INSERT INTO [dbo].[ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (7, N'الرواتب');
    INSERT INTO [dbo].[ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (8, N'المستلزمات');
    INSERT INTO [dbo].[ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (9, N'مصروفات أخرى');
END;
SET IDENTITY_INSERT [dbo].[ExpenseTypes] OFF;
-- Accounts: 89 rows (the missing ones)
BEGIN
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1, N'الأصول', N'ASSET', NULL, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 11)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (11, N'الأصول المتداولة', N'ASSET', 1, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1100)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1100, N'النقدية بالخزينة والصناديق', N'ASSET', 11, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 110001)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (110001, N'الخزينة الرئيسية', N'ASSET', 1100, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 110002)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (110002, N'صندوق الكاشير', N'ASSET', 1100, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1190)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1190, N'نقدية غير موزعة على صندوق', N'ASSET', 11, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1195)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1195, N'نقدية بالطريق بين الصناديق', N'ASSET', 11, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1200)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1200, N'مدى والمحافظ تحت التسوية (البنك والشبكة)', N'ASSET', 11, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1210)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1210, N'الحسابات البنكية', N'ASSET', 11, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1250)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1250, N'شيكات تحت التحصيل', N'ASSET', 11, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1300)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1300, N'ذمم العملاء', N'ASSET', 11, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1310)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1310, N'أوراق القبض', N'ASSET', 11, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1350)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1350, N'مخصص الديون المشكوك في تحصيلها', N'ASSET', 11, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1400)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1400, N'المخزون', N'ASSET', 11, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1410)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1410, N'بضاعة بالطريق بين المخازن', N'ASSET', 11, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1420)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1420, N'إنتاج تحت التشغيل', N'ASSET', 11, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1430)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1430, N'بضاعة مسلّمة للعملاء لم تُفوتر', N'ASSET', 11, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1500)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1500, N'ضريبة القيمة المضافة - مدخلات', N'ASSET', 11, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1600)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1600, N'سلف الموظفين', N'ASSET', 11, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1610)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1610, N'العهد النقدية', N'ASSET', 11, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1650)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1650, N'مصروفات مدفوعة مقدمًا', N'ASSET', 11, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1660)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1660, N'إيرادات مستحقة', N'ASSET', 11, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1690)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1690, N'تأمينات لدى الغير', N'ASSET', 11, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 12)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (12, N'الأصول غير المتداولة', N'ASSET', 1, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1710)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1710, N'الأثاث والتجهيزات', N'ASSET', 12, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1720)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1720, N'الأجهزة والحاسبات وأنظمة نقاط البيع', N'ASSET', 12, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1730)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1730, N'السيارات ووسائل النقل', N'ASSET', 12, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1740)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1740, N'الديكورات وتحسينات المحل', N'ASSET', 12, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1790)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1790, N'مجمع إهلاك الأصول الثابتة', N'ASSET', 12, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 1800)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1800, N'البرامج والتراخيص', N'ASSET', 12, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2, N'الخصوم', N'LIABILITY', NULL, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 21)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (21, N'الخصوم المتداولة', N'LIABILITY', 2, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2100)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2100, N'ذمم الموردين', N'LIABILITY', 21, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2110)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2110, N'أوراق الدفع (شيكات صادرة)', N'LIABILITY', 21, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2150)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2150, N'بضاعة مستلمة من الموردين لم تُفوتر', N'LIABILITY', 21, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2200)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2200, N'ضريبة القيمة المضافة - مخرجات', N'LIABILITY', 21, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2250)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2250, N'ضريبة القيمة المضافة - التسوية والسداد', N'LIABILITY', 21, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2300)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2300, N'مصروفات مستحقة', N'LIABILITY', 21, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2310)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2310, N'رواتب وأجور مستحقة', N'LIABILITY', 21, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2320)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2320, N'التأمينات الاجتماعية المستحقة', N'LIABILITY', 21, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2330)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2330, N'عمولات مستحقة للمندوبين', N'LIABILITY', 21, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2400)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2400, N'دفعات مقدمة من العملاء', N'LIABILITY', 21, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2500)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2500, N'قروض قصيرة الأجل', N'LIABILITY', 21, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2600)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2600, N'الزكاة المستحقة', N'LIABILITY', 21, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 22)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (22, N'الخصوم غير المتداولة', N'LIABILITY', 2, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2700)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2700, N'قروض طويلة الأجل', N'LIABILITY', 22, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 2800)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2800, N'مخصص مكافأة نهاية الخدمة', N'LIABILITY', 22, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 3)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3, N'حقوق الملكية', N'EQUITY', NULL, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 31)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (31, N'رأس المال وجاري المالك', N'EQUITY', 3, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 3200)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3200, N'رأس المال', N'EQUITY', 31, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 3100)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3100, N'جاري المالك', N'EQUITY', 31, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 3900)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3900, N'إجمالي الأرصدة الافتتاحية', N'EQUITY', 31, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 32)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (32, N'الأرباح المحتجزة ونتائج الأعمال', N'EQUITY', 3, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 3300)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3300, N'الأرباح المحتجزة', N'EQUITY', 32, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 3400)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3400, N'صافي ربح (خسارة) العام', N'EQUITY', 32, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 3500)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3500, N'فائض إعادة تقييم الأصول الثابتة', N'EQUITY', 32, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 4)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4, N'الإيرادات', N'REVENUE', NULL, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 41)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (41, N'إيرادات النشاط', N'REVENUE', 4, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 4100)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4100, N'المبيعات', N'REVENUE', 41, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 4110)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4110, N'مردودات المبيعات', N'REVENUE', 41, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 4120)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4120, N'الخصم المسموح به', N'REVENUE', 41, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 42)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (42, N'إيرادات أخرى', N'REVENUE', 4, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 4200)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4200, N'إيرادات متنوعة', N'REVENUE', 42, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 4300)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4300, N'زيادة الصناديق', N'REVENUE', 42, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 4400)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4400, N'الخصم المكتسب', N'REVENUE', 42, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 4500)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4500, N'أرباح بيع الأصول الثابتة', N'REVENUE', 42, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5, N'المصروفات', N'EXPENSE', NULL, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 51)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (51, N'تكلفة المبيعات', N'EXPENSE', 5, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5100)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5100, N'تكلفة البضاعة المباعة', N'EXPENSE', 51, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5200)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5200, N'فروقات وتسويات المخزون', N'EXPENSE', 51, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5250)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5250, N'هالك المخزون', N'EXPENSE', 51, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 52)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (52, N'المصروفات التشغيلية والإدارية', N'EXPENSE', 5, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5300)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5300, N'المصروفات حسب النوع', N'EXPENSE', 52, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5500)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5500, N'الرواتب والأجور', N'EXPENSE', 52, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5510)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5510, N'البدلات والحوافز', N'EXPENSE', 52, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5520)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5520, N'التأمينات الاجتماعية', N'EXPENSE', 52, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5530)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5530, N'عمولات المندوبين', N'EXPENSE', 52, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5600)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5600, N'إهلاك الأصول الثابتة', N'EXPENSE', 52, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5610)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5610, N'عمولات البنوك ونقاط البيع', N'EXPENSE', 52, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5620)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5620, N'الرسوم الحكومية والتراخيص', N'EXPENSE', 52, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5630)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5630, N'الدعاية والإعلان', N'EXPENSE', 52, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5640)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5640, N'مستلزمات مصروفة من المخزون للأقسام', N'EXPENSE', 52, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 53)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (53, N'مصروفات أخرى', N'EXPENSE', 5, 0, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5400)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5400, N'عجز الصناديق', N'EXPENSE', 53, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5650)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5650, N'خسائر بيع واستبعاد الأصول الثابتة', N'EXPENSE', 53, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5660)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5660, N'خسائر انخفاض قيمة الأصول الثابتة', N'EXPENSE', 53, 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5700)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5700, N'الديون المعدومة', N'EXPENSE', 53, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5800)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5800, N'الزكاة', N'EXPENSE', 53, 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Accounts] WHERE [AccountCode] = 5900)
        INSERT INTO [dbo].[Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5900, N'مصروفات متنوعة', N'EXPENSE', 53, 1, 1);
END;
-- JournalSourceTypes: 38 rows (the missing ones)
BEGIN
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'SALE')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'SALE', N'فاتورة بيع', 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'SALES_RETURN')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'SALES_RETURN', N'مرتجع بيع', 2);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'PURCHASE')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'PURCHASE', N'فاتورة شراء', 3);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'PURCHASE_RETURN')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'PURCHASE_RETURN', N'مرتجع شراء', 4);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'CUSTOMER_PAYMENT')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'CUSTOMER_PAYMENT', N'سند قبض من عميل', 5);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'SUPPLIER_PAYMENT')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'SUPPLIER_PAYMENT', N'سند صرف لمورد', 6);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'EXPENSE')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'EXPENSE', N'مصروف', 7);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'CASH_VOUCHER')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'CASH_VOUCHER', N'سند نقدية', 8);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'STOCK_MOVE')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'STOCK_MOVE', N'حركة مخزون يدوية', 9);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'STOCK_COUNT')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'STOCK_COUNT', N'تسوية جرد', 10);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'BOX_OPENING')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'BOX_OPENING', N'رصيد افتتاحي لصندوق', 11);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'CUSTOMER_OPENING')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'CUSTOMER_OPENING', N'رصيد افتتاحي لعميل', 12);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'SUPPLIER_OPENING')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'SUPPLIER_OPENING', N'رصيد افتتاحي لمورد', 13);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'MANUAL')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'MANUAL', N'قيد يدوي', 14);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'YEAR_CLOSE')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'YEAR_CLOSE', N'قيد إقفال السنة', 15);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'VAT_RETURN')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'VAT_RETURN', N'تسوية إقرار ضريبة القيمة المضافة', 16);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'VAT_PAYMENT')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'VAT_PAYMENT', N'سداد ضريبة القيمة المضافة', 17);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'BANK_OPENING')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'BANK_OPENING', N'رصيد افتتاحي لبنك', 18);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'BANK_TX')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'BANK_TX', N'حركة بنكية', 19);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'CHEQUE')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'CHEQUE', N'شيك وارد أو صادر', 20);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'CHEQUE_STATUS')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'CHEQUE_STATUS', N'تحصيل أو ارتداد شيك', 21);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'ASSET')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'ASSET', N'اقتناء أصل ثابت', 22);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'ASSET_DISPOSAL')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'ASSET_DISPOSAL', N'بيع أو استبعاد أصل ثابت', 23);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'DEPRECIATION')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'DEPRECIATION', N'قيد الإهلاك الشهري', 24);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'PAYROLL')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'PAYROLL', N'قيد مسير الرواتب', 25);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'PAYROLL_PAYMENT')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'PAYROLL_PAYMENT', N'صرف الرواتب', 26);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'COMMISSION')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'COMMISSION', N'قيد عمولات المندوبين', 27);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'ACCOUNT_OPENING')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'ACCOUNT_OPENING', N'رصيد افتتاحي لحساب', 28);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'CASH_RECEIVE')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'CASH_RECEIVE', N'استلام تحويل نقدية', 29);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'STOCK_TRANSFER')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'STOCK_TRANSFER', N'تحويل بين المخازن', 30);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'STOCK_RECEIVE')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'STOCK_RECEIVE', N'استلام تحويل مخزني', 31);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'STOCK_ISSUE')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'STOCK_ISSUE', N'صرف مخزني لجهة', 32);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'STOCK_RECEIPT')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'STOCK_RECEIPT', N'استلام مخزني من جهة', 33);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'DELIVERY_NOTE')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'DELIVERY_NOTE', N'إذن تسليم لعميل', 34);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'RECEIPT_NOTE')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'RECEIPT_NOTE', N'إذن استلام من مورد', 35);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'ASSET_REVAL')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'ASSET_REVAL', N'إعادة تقييم أصل ثابت', 36);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'ASSET_TRANSFER')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'ASSET_TRANSFER', N'نقل أصل ثابت بين مراكز التكلفة', 37);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[JournalSourceTypes] WHERE [SourceType] = N'REP_SHORTAGE')
        INSERT INTO [dbo].[JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES (N'REP_SHORTAGE', N'عجز بضاعة المندوب سلفة عليه', 38);
END;
-- TransactionTypes: 14 rows (the missing ones)
BEGIN
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 1)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (1, N'PURCHASE', N'شراء', 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 2)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (2, N'SALE', N'بيع', -1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 3)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (3, N'PURCHASE_RETURN', N'مرتجع شراء', -1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 4)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (4, N'SALES_RETURN', N'مرتجع بيع', 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 5)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (5, N'STOCK_IN', N'إضافة مخزون', 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 6)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (6, N'STOCK_OUT', N'خصم مخزون', -1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 7)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (7, N'ADJUSTMENT', N'تسوية جرد', 0, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 8)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (8, N'OPENING', N'رصيد افتتاحي', 1, 1);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 9)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (9, N'TRANSFER_OUT', N'تحويل صادر لمخزن آخر', -1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 10)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (10, N'TRANSFER_IN', N'تحويل وارد من مخزن آخر', 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 11)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (11, N'PARTY_ISSUE', N'صرف لجهة (سند صرف)', -1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 12)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (12, N'PARTY_RECEIPT', N'استلام من جهة (سند استلام)', 1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 13)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (13, N'DELIVERY_NOTE', N'تسليم لعميل (إذن تسليم)', -1, 0);
    IF NOT EXISTS (SELECT 1 FROM [dbo].[TransactionTypes] WHERE [TransactionTypeID] = 14)
        INSERT INTO [dbo].[TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (14, N'RECEIPT_NOTE', N'استلام من مورد (إذن استلام)', 1, 0);
END;
-- StockParties: 1 rows (an empty table)
SET IDENTITY_INSERT [dbo].[StockParties] ON;
IF NOT EXISTS (SELECT 1 FROM [dbo].[StockParties])
BEGIN
    INSERT INTO [dbo].[StockParties] ([StockPartyID], [PartyCode], [PartyName], [PartyType]) VALUES (1, N'SP-1', N'الهالك والتالف', N'WASTE');
END;
SET IDENTITY_INSERT [dbo].[StockParties] OFF;
-- LabelSettings: 1 rows (an empty table)
IF NOT EXISTS (SELECT 1 FROM [dbo].[LabelSettings])
BEGIN
    INSERT INTO [dbo].[LabelSettings] ([LabelSettingID]) VALUES (1);
END;
GO

/*------------------------------------------------------------------------------
  6. English names of the system rows (only the empty ones)
------------------------------------------------------------------------------*/
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Assets' WHERE [AccountCode] = 1 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Current assets' WHERE [AccountCode] = 11 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Cash in treasury and boxes' WHERE [AccountCode] = 1100 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Main treasury' WHERE [AccountCode] = 110001 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Cashier box' WHERE [AccountCode] = 110002 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Cash not allocated to a box' WHERE [AccountCode] = 1190 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Cash in transit between boxes' WHERE [AccountCode] = 1195 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Goods in transit between warehouses' WHERE [AccountCode] = 1410 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Production work in progress' WHERE [AccountCode] = 1420 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Goods delivered to customers not invoiced' WHERE [AccountCode] = 1430 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Mada and wallets under settlement (bank and network)' WHERE [AccountCode] = 1200 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Bank accounts' WHERE [AccountCode] = 1210 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Cheques under collection' WHERE [AccountCode] = 1250 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Accounts receivable' WHERE [AccountCode] = 1300 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Notes receivable' WHERE [AccountCode] = 1310 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Allowance for doubtful debts' WHERE [AccountCode] = 1350 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Inventory' WHERE [AccountCode] = 1400 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'VAT - input' WHERE [AccountCode] = 1500 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Employee advances' WHERE [AccountCode] = 1600 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Cash floats' WHERE [AccountCode] = 1610 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Prepaid expenses' WHERE [AccountCode] = 1650 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Accrued revenue' WHERE [AccountCode] = 1660 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Deposits with others' WHERE [AccountCode] = 1690 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Non-current assets' WHERE [AccountCode] = 12 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Furniture and fixtures' WHERE [AccountCode] = 1710 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Computers, devices and POS systems' WHERE [AccountCode] = 1720 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Vehicles' WHERE [AccountCode] = 1730 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Decorations and store improvements' WHERE [AccountCode] = 1740 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Accumulated depreciation of fixed assets' WHERE [AccountCode] = 1790 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Software and licenses' WHERE [AccountCode] = 1800 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Liabilities' WHERE [AccountCode] = 2 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Current liabilities' WHERE [AccountCode] = 21 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Accounts payable' WHERE [AccountCode] = 2100 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Notes payable (issued cheques)' WHERE [AccountCode] = 2110 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Goods received from suppliers not invoiced' WHERE [AccountCode] = 2150 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'VAT - output' WHERE [AccountCode] = 2200 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'VAT - settlement and payment' WHERE [AccountCode] = 2250 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Accrued expenses' WHERE [AccountCode] = 2300 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Salaries and wages payable' WHERE [AccountCode] = 2310 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Social insurance payable' WHERE [AccountCode] = 2320 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Sales rep commissions payable' WHERE [AccountCode] = 2330 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Customer advances' WHERE [AccountCode] = 2400 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Short-term loans' WHERE [AccountCode] = 2500 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Zakat payable' WHERE [AccountCode] = 2600 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Non-current liabilities' WHERE [AccountCode] = 22 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Long-term loans' WHERE [AccountCode] = 2700 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'End-of-service benefits provision' WHERE [AccountCode] = 2800 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Equity' WHERE [AccountCode] = 3 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Capital and owner current account' WHERE [AccountCode] = 31 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Capital' WHERE [AccountCode] = 3200 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Owner current account' WHERE [AccountCode] = 3100 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Total opening balances' WHERE [AccountCode] = 3900 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Retained earnings and results' WHERE [AccountCode] = 32 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Retained earnings' WHERE [AccountCode] = 3300 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Net profit (loss) of the year' WHERE [AccountCode] = 3400 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Fixed asset revaluation surplus' WHERE [AccountCode] = 3500 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Revenue' WHERE [AccountCode] = 4 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Operating revenue' WHERE [AccountCode] = 41 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Sales' WHERE [AccountCode] = 4100 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Sales returns' WHERE [AccountCode] = 4110 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Discounts allowed' WHERE [AccountCode] = 4120 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Other revenue' WHERE [AccountCode] = 42 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Miscellaneous revenue' WHERE [AccountCode] = 4200 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Cash overages' WHERE [AccountCode] = 4300 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Discounts received' WHERE [AccountCode] = 4400 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Gain on sale of fixed assets' WHERE [AccountCode] = 4500 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Expenses' WHERE [AccountCode] = 5 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Cost of sales' WHERE [AccountCode] = 51 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Cost of goods sold' WHERE [AccountCode] = 5100 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Stock differences and adjustments' WHERE [AccountCode] = 5200 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Stock waste' WHERE [AccountCode] = 5250 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Operating and administrative expenses' WHERE [AccountCode] = 52 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Expenses by type' WHERE [AccountCode] = 5300 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Salaries and wages' WHERE [AccountCode] = 5500 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Allowances and incentives' WHERE [AccountCode] = 5510 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Social insurance' WHERE [AccountCode] = 5520 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Sales rep commissions' WHERE [AccountCode] = 5530 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Depreciation of fixed assets' WHERE [AccountCode] = 5600 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Bank and POS fees' WHERE [AccountCode] = 5610 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Government fees and licenses' WHERE [AccountCode] = 5620 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Advertising' WHERE [AccountCode] = 5630 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Supplies issued from stock to departments' WHERE [AccountCode] = 5640 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Other expenses' WHERE [AccountCode] = 53 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Cash shortages' WHERE [AccountCode] = 5400 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Loss on sale and disposal of fixed assets' WHERE [AccountCode] = 5650 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Impairment loss on fixed assets' WHERE [AccountCode] = 5660 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Bad debts' WHERE [AccountCode] = 5700 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Zakat' WHERE [AccountCode] = 5800 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[Accounts] SET [AccountNameEn] = N'Miscellaneous expenses' WHERE [AccountCode] = 5900 AND [AccountNameEn] IS NULL;
UPDATE [dbo].[PaymentMethods] SET [MethodNameEn] = N'Cash' WHERE [PaymentMethodID] = 1 AND [MethodNameEn] IS NULL;
UPDATE [dbo].[PaymentMethods] SET [MethodNameEn] = N'Mada / card' WHERE [PaymentMethodID] = 2 AND [MethodNameEn] IS NULL;
UPDATE [dbo].[PaymentMethods] SET [MethodNameEn] = N'Bank transfer' WHERE [PaymentMethodID] = 3 AND [MethodNameEn] IS NULL;
UPDATE [dbo].[PaymentMethods] SET [MethodNameEn] = N'E-wallet' WHERE [PaymentMethodID] = 4 AND [MethodNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Sales invoice' WHERE [SourceType] = N'SALE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Sales return' WHERE [SourceType] = N'SALES_RETURN' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Purchase invoice' WHERE [SourceType] = N'PURCHASE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Purchase return' WHERE [SourceType] = N'PURCHASE_RETURN' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Customer receipt' WHERE [SourceType] = N'CUSTOMER_PAYMENT' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Supplier payment' WHERE [SourceType] = N'SUPPLIER_PAYMENT' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Expense' WHERE [SourceType] = N'EXPENSE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Cash voucher' WHERE [SourceType] = N'CASH_VOUCHER' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Manual stock move' WHERE [SourceType] = N'STOCK_MOVE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Stock count adjustment' WHERE [SourceType] = N'STOCK_COUNT' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Box opening balance' WHERE [SourceType] = N'BOX_OPENING' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Customer opening balance' WHERE [SourceType] = N'CUSTOMER_OPENING' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Supplier opening balance' WHERE [SourceType] = N'SUPPLIER_OPENING' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Manual entry' WHERE [SourceType] = N'MANUAL' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Year closing entry' WHERE [SourceType] = N'YEAR_CLOSE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'VAT return settlement' WHERE [SourceType] = N'VAT_RETURN' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'VAT payment' WHERE [SourceType] = N'VAT_PAYMENT' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Bank opening balance' WHERE [SourceType] = N'BANK_OPENING' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Bank transaction' WHERE [SourceType] = N'BANK_TX' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Received or issued cheque' WHERE [SourceType] = N'CHEQUE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Cheque collection or bounce' WHERE [SourceType] = N'CHEQUE_STATUS' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Fixed asset purchase' WHERE [SourceType] = N'ASSET' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Fixed asset sale or disposal' WHERE [SourceType] = N'ASSET_DISPOSAL' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Monthly depreciation entry' WHERE [SourceType] = N'DEPRECIATION' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Payroll entry' WHERE [SourceType] = N'PAYROLL' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Salary payment' WHERE [SourceType] = N'PAYROLL_PAYMENT' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Sales rep commissions entry' WHERE [SourceType] = N'COMMISSION' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Account opening balance' WHERE [SourceType] = N'ACCOUNT_OPENING' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Cash transfer received' WHERE [SourceType] = N'CASH_RECEIVE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Transfer between warehouses' WHERE [SourceType] = N'STOCK_TRANSFER' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Stock transfer received' WHERE [SourceType] = N'STOCK_RECEIVE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Stock issue to a party' WHERE [SourceType] = N'STOCK_ISSUE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Stock receipt from a party' WHERE [SourceType] = N'STOCK_RECEIPT' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Delivery note to a customer' WHERE [SourceType] = N'DELIVERY_NOTE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Goods receipt note from a supplier' WHERE [SourceType] = N'RECEIPT_NOTE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Fixed asset revaluation' WHERE [SourceType] = N'ASSET_REVAL' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Fixed asset transfer between cost centres' WHERE [SourceType] = N'ASSET_TRANSFER' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[JournalSourceTypes] SET [TypeNameEn] = N'Rep goods shortage charged as an advance' WHERE [SourceType] = N'REP_SHORTAGE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Purchase' WHERE [TypeCode] = N'PURCHASE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Sale' WHERE [TypeCode] = N'SALE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Purchase return' WHERE [TypeCode] = N'PURCHASE_RETURN' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Sales return' WHERE [TypeCode] = N'SALES_RETURN' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Stock addition' WHERE [TypeCode] = N'STOCK_IN' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Stock deduction' WHERE [TypeCode] = N'STOCK_OUT' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Count adjustment' WHERE [TypeCode] = N'ADJUSTMENT' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Opening balance' WHERE [TypeCode] = N'OPENING' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Transfer out to another warehouse' WHERE [TypeCode] = N'TRANSFER_OUT' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Transfer in from another warehouse' WHERE [TypeCode] = N'TRANSFER_IN' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Issue to a party (issue voucher)' WHERE [TypeCode] = N'PARTY_ISSUE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Receipt from a party (receipt voucher)' WHERE [TypeCode] = N'PARTY_RECEIPT' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Delivery to a customer (delivery note)' WHERE [TypeCode] = N'DELIVERY_NOTE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[TransactionTypes] SET [TypeNameEn] = N'Receipt from a supplier (receipt note)' WHERE [TypeCode] = N'RECEIPT_NOTE' AND [TypeNameEn] IS NULL;
UPDATE [dbo].[Roles] SET [RoleNameEn] = N'System administrator' WHERE [RoleID] = 1 AND [RoleNameEn] IS NULL;
UPDATE [dbo].[Roles] SET [RoleNameEn] = N'Manager' WHERE [RoleID] = 2 AND [RoleNameEn] IS NULL;
UPDATE [dbo].[Roles] SET [RoleNameEn] = N'Cashier' WHERE [RoleID] = 3 AND [RoleNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Point of sale' WHERE [PermissionKey] = N'SALES_POS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'View and reprint invoices' WHERE [PermissionKey] = N'SALES_VIEW' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Sales returns' WHERE [PermissionKey] = N'SALES_RETURN' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Change the sale price in the invoice' WHERE [PermissionKey] = N'PRICE_OVERRIDE' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Discount above the allowed limit' WHERE [PermissionKey] = N'DISCOUNT_OVERRIDE' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Sell more than the available quantity' WHERE [PermissionKey] = N'ALLOW_NEGATIVE_STOCK' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Manage customers' WHERE [PermissionKey] = N'CUSTOMERS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Receipt vouchers' WHERE [PermissionKey] = N'CUSTOMER_PAYMENTS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Purchase invoices' WHERE [PermissionKey] = N'PURCHASES' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Purchase returns' WHERE [PermissionKey] = N'PURCHASE_RETURN' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Purchase orders and the purchase suggestion' WHERE [PermissionKey] = N'PURCHASE_ORDERS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Manage suppliers' WHERE [PermissionKey] = N'SUPPLIERS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Payment vouchers' WHERE [PermissionKey] = N'SUPPLIER_PAYMENTS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Manage products and prices' WHERE [PermissionKey] = N'PRODUCTS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Manual stock addition and deduction' WHERE [PermissionKey] = N'INVENTORY_ADJUST' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Stock count' WHERE [PermissionKey] = N'STOCK_COUNT' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Expenses' WHERE [PermissionKey] = N'EXPENSES' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Treasury: receipt, payment and transfer vouchers and boxes' WHERE [PermissionKey] = N'CASH_BOX' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Daily cashier closing' WHERE [PermissionKey] = N'CASH_CLOSING' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Pay and transfer from every box, not only own boxes' WHERE [PermissionKey] = N'CASH_ALL_BOXES' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Transfers between warehouses: sending and receiving' WHERE [PermissionKey] = N'STOCK_TRANSFER' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Stock issue and receipt vouchers and their parties' WHERE [PermissionKey] = N'STOCK_DOCS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Delivery notes to customers and goods receipt notes from suppliers' WHERE [PermissionKey] = N'STOCK_NOTES' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Customer orders: entry, approval, preparation and cancellation' WHERE [PermissionKey] = N'SALES_ORDERS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Customer quotations: entry, issue and conversion' WHERE [PermissionKey] = N'QUOTATIONS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Rep day: van loading, selling in the rep name and settlement' WHERE [PermissionKey] = N'REP_DAYS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Rep routes, visits and rep price lists' WHERE [PermissionKey] = N'REP_ROUTES' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Journal entries, chart of accounts and trial balance' WHERE [PermissionKey] = N'JOURNAL' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Manual entries: add, edit and delete' WHERE [PermissionKey] = N'MANUAL_ENTRY' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Close and reopen periods and the fiscal year' WHERE [PermissionKey] = N'PERIOD_CLOSE' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'VAT return: approval and payment' WHERE [PermissionKey] = N'VAT_RETURN' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Banks: accounts, bank transactions and reconciliation' WHERE [PermissionKey] = N'BANKS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Received and issued cheques: recording, collection and bounce' WHERE [PermissionKey] = N'CHEQUES' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Fixed assets and depreciation' WHERE [PermissionKey] = N'FIXED_ASSETS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Payroll: preparation, posting and payment' WHERE [PermissionKey] = N'PAYROLL' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Budget: preparation and comparison with actuals' WHERE [PermissionKey] = N'BUDGET' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Currencies and exchange rates' WHERE [PermissionKey] = N'CURRENCIES' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Start-up opening balances: date, accounts, customers, suppliers, cash boxes, banks and stock' WHERE [PermissionKey] = N'OPENING_BALANCES' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Sales reps: data, targets, commissions and reports' WHERE [PermissionKey] = N'SALES_REPS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Operating reports' WHERE [PermissionKey] = N'REPORTS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Profit and VAT reports' WHERE [PermissionKey] = N'REPORTS_PROFIT' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Financial figures on the dashboard' WHERE [PermissionKey] = N'DASHBOARD_FINANCIAL' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Store settings' WHERE [PermissionKey] = N'SETTINGS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Users and permissions' WHERE [PermissionKey] = N'USERS' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Backup' WHERE [PermissionKey] = N'BACKUP' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'SQL Server database: connection, data transfer and running on the server' WHERE [PermissionKey] = N'SQL_SERVER' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'Audit trail: who added, edited or deleted, with the values before and after' WHERE [PermissionKey] = N'AUDIT_LOG' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [PermissionNameEn] = N'E-invoicing: follow-up and resending' WHERE [PermissionKey] = N'EINVOICE' AND [PermissionNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [ModuleNameEn] = N'Sales' WHERE [ModuleName] = N'المبيعات' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [ModuleNameEn] = N'Customers' WHERE [ModuleName] = N'العملاء' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [ModuleNameEn] = N'Purchases' WHERE [ModuleName] = N'المشتريات' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [ModuleNameEn] = N'Suppliers' WHERE [ModuleName] = N'الموردون' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [ModuleNameEn] = N'Inventory' WHERE [ModuleName] = N'المخزون' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [ModuleNameEn] = N'Expenses' WHERE [ModuleName] = N'المصروفات' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [ModuleNameEn] = N'Treasury' WHERE [ModuleName] = N'الخزينة' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [ModuleNameEn] = N'Accounting' WHERE [ModuleName] = N'الحسابات' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [ModuleNameEn] = N'Reports' WHERE [ModuleName] = N'التقارير' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Permissions] SET [ModuleNameEn] = N'System' WHERE [ModuleName] = N'النظام' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Point of sale (shops)' WHERE [ScreenName] = N'frmPOS' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Point of sale (restaurants)' WHERE [ScreenName] = N'frmTouchPOS' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Point of sale (cafes)' WHERE [ScreenName] = N'frmCafePOS' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'View and reprint invoices' WHERE [ScreenName] = N'frmSalesInvoice' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Sales returns' WHERE [ScreenName] = N'frmSalesReturn' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Customers' WHERE [ScreenName] = N'frmCustomers' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Customer receipt vouchers' WHERE [ScreenName] = N'frmCustomerPayment' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Purchase invoices' WHERE [ScreenName] = N'frmPurchaseInvoice' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'View purchase invoices' WHERE [ScreenName] = N'frmPurchaseView' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Purchase returns' WHERE [ScreenName] = N'frmPurchaseReturn' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Suppliers' WHERE [ScreenName] = N'frmSuppliers' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Supplier payment vouchers' WHERE [ScreenName] = N'frmSupplierPayment' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Products and prices' WHERE [ScreenName] = N'frmProducts' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Categories' WHERE [ScreenName] = N'frmCategories' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Units' WHERE [ScreenName] = N'frmUnits' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Product units' WHERE [ScreenName] = N'frmProductUnits' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Price lists' WHERE [ScreenName] = N'frmPriceLists' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Purchase orders' WHERE [ScreenName] = N'frmPurchaseOrder' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Customer quotations' WHERE [ScreenName] = N'frmQuotation' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Price list prices' WHERE [ScreenName] = N'frmPriceListItems' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Stock and manual moves' WHERE [ScreenName] = N'frmInventory' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Stock count' WHERE [ScreenName] = N'frmStockCount' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Barcode labels' WHERE [ScreenName] = N'frmBarcodeLabels' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Label settings' WHERE [ScreenName] = N'frmLabelSettings' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Expenses' WHERE [ScreenName] = N'frmExpenses' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Expense types' WHERE [ScreenName] = N'frmExpenseTypes' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Recurring expenses' WHERE [ScreenName] = N'frmRecurring' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Treasury' WHERE [ScreenName] = N'frmTreasury' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Cash and transfer vouchers' WHERE [ScreenName] = N'frmCashVoucher' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Daily cashier closing' WHERE [ScreenName] = N'frmCashClosing' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Boxes' WHERE [ScreenName] = N'frmCashBoxes' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Box custodians' WHERE [ScreenName] = N'frmCashBoxUsers' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Receipt of transfers between boxes' WHERE [ScreenName] = N'frmCashTransfers' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Warehouses' WHERE [ScreenName] = N'frmWarehouses' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Transfers between warehouses' WHERE [ScreenName] = N'frmStockTransfer' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Stock issue and receipt parties' WHERE [ScreenName] = N'frmStockParties' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Stock issue and receipt vouchers' WHERE [ScreenName] = N'frmStockDocument' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Delivery and goods receipt notes' WHERE [ScreenName] = N'frmStockNote' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Customer orders' WHERE [ScreenName] = N'frmSalesOrder' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Rep day and van sales' WHERE [ScreenName] = N'frmRepDay' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Rep visits and the day plan' WHERE [ScreenName] = N'frmRepVisits' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Rep routes' WHERE [ScreenName] = N'frmRouteStops' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Rep price lists' WHERE [ScreenName] = N'frmRepPrices' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Fixed asset revaluation and transfer' WHERE [ScreenName] = N'frmAssetChanges' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Accounting and finance' WHERE [ScreenName] = N'frmAccounting' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Journal entries' WHERE [ScreenName] = N'frmJournal' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Chart of accounts' WHERE [ScreenName] = N'frmAccounts' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Manual entries' WHERE [ScreenName] = N'frmManualEntry' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Account statement and general ledger' WHERE [ScreenName] = N'frmLedger' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Financial statements' WHERE [ScreenName] = N'frmFinancials' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Period and fiscal year closing' WHERE [ScreenName] = N'frmPeriodClosing' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Fiscal periods' WHERE [ScreenName] = N'frmFiscalPeriods' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'VAT return' WHERE [ScreenName] = N'frmVatReturn' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Start-up opening balances' WHERE [ScreenName] = N'frmOpeningBalances' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Aging (customers and suppliers)' WHERE [ScreenName] = N'frmAging' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Banks' WHERE [ScreenName] = N'frmBanks' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Bank transactions' WHERE [ScreenName] = N'frmBankTx' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Bank reconciliation' WHERE [ScreenName] = N'frmBankRecon' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Received and issued cheques' WHERE [ScreenName] = N'frmCheques' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Fixed assets' WHERE [ScreenName] = N'frmAssets' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Monthly depreciation' WHERE [ScreenName] = N'frmDepreciation' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Payroll' WHERE [ScreenName] = N'frmPayroll' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Cost centers and branches' WHERE [ScreenName] = N'frmCostCenters' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Budget' WHERE [ScreenName] = N'frmBudget' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Currencies' WHERE [ScreenName] = N'frmCurrencies' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Sales reps' WHERE [ScreenName] = N'frmSalesReps' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Sales rep targets' WHERE [ScreenName] = N'frmRepTargets' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Sales rep commissions' WHERE [ScreenName] = N'frmCommissions' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Exchange rates' WHERE [ScreenName] = N'frmCurrencyRates' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Match payments to invoices' WHERE [ScreenName] = N'frmAllocation' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Reports' WHERE [ScreenName] = N'frmReportCenter' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Search' WHERE [ScreenName] = N'frmSearch' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Store settings' WHERE [ScreenName] = N'frmSettings' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Users' WHERE [ScreenName] = N'frmUsers' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Roles and permissions' WHERE [ScreenName] = N'frmRoles' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Screen permissions of users' WHERE [ScreenName] = N'frmUserScreens' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Audit trail' WHERE [ScreenName] = N'frmAuditLog' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Backup' WHERE [ScreenName] = N'frmBackup' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'SQL Server database' WHERE [ScreenName] = N'frmSqlServer' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'English names' WHERE [ScreenName] = N'frmEnglishNames' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'E-invoicing' WHERE [ScreenName] = N'frmEInvoices' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'Fatoora connection setup' WHERE [ScreenName] = N'frmZatcaSetup' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ScreenTitleEn] = N'E-receipt connection setup' WHERE [ScreenName] = N'frmEtaSetup' AND [ScreenTitleEn] IS NULL;
UPDATE [dbo].[Screens] SET [ModuleNameEn] = N'Sales' WHERE [ModuleName] = N'المبيعات' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Screens] SET [ModuleNameEn] = N'Customers' WHERE [ModuleName] = N'العملاء' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Screens] SET [ModuleNameEn] = N'Purchases' WHERE [ModuleName] = N'المشتريات' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Screens] SET [ModuleNameEn] = N'Suppliers' WHERE [ModuleName] = N'الموردون' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Screens] SET [ModuleNameEn] = N'Inventory' WHERE [ModuleName] = N'المخزون' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Screens] SET [ModuleNameEn] = N'Expenses' WHERE [ModuleName] = N'المصروفات' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Screens] SET [ModuleNameEn] = N'Treasury' WHERE [ModuleName] = N'الخزينة' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Screens] SET [ModuleNameEn] = N'Accounting' WHERE [ModuleName] = N'الحسابات' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Screens] SET [ModuleNameEn] = N'Reports' WHERE [ModuleName] = N'التقارير' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Screens] SET [ModuleNameEn] = N'System' WHERE [ModuleName] = N'النظام' AND [ModuleNameEn] IS NULL;
UPDATE [dbo].[Categories] SET [CategoryNameEn] = N'General' WHERE [CategoryID] = 1 AND [CategoryNameEn] IS NULL;
UPDATE [dbo].[Units] SET [UnitNameEn] = N'Piece' WHERE [UnitID] = 1 AND [UnitNameEn] IS NULL;
UPDATE [dbo].[Units] SET [UnitNameEn] = N'Box' WHERE [UnitID] = 2 AND [UnitNameEn] IS NULL;
UPDATE [dbo].[Units] SET [UnitNameEn] = N'Carton' WHERE [UnitID] = 3 AND [UnitNameEn] IS NULL;
UPDATE [dbo].[Units] SET [UnitNameEn] = N'Pack' WHERE [UnitID] = 4 AND [UnitNameEn] IS NULL;
UPDATE [dbo].[Units] SET [UnitNameEn] = N'Kilo' WHERE [UnitID] = 5 AND [UnitNameEn] IS NULL;
UPDATE [dbo].[Units] SET [UnitNameEn] = N'Litre' WHERE [UnitID] = 6 AND [UnitNameEn] IS NULL;
UPDATE [dbo].[Units] SET [UnitNameEn] = N'Metre' WHERE [UnitID] = 7 AND [UnitNameEn] IS NULL;
UPDATE [dbo].[Units] SET [UnitNameEn] = N'Set' WHERE [UnitID] = 8 AND [UnitNameEn] IS NULL;
UPDATE [dbo].[ExpenseTypes] SET [ExpenseTypeNameEn] = N'Rent' WHERE [ExpenseTypeID] = 1 AND [ExpenseTypeNameEn] IS NULL;
UPDATE [dbo].[ExpenseTypes] SET [ExpenseTypeNameEn] = N'Electricity' WHERE [ExpenseTypeID] = 2 AND [ExpenseTypeNameEn] IS NULL;
UPDATE [dbo].[ExpenseTypes] SET [ExpenseTypeNameEn] = N'Water' WHERE [ExpenseTypeID] = 3 AND [ExpenseTypeNameEn] IS NULL;
UPDATE [dbo].[ExpenseTypes] SET [ExpenseTypeNameEn] = N'Internet and telecom' WHERE [ExpenseTypeID] = 4 AND [ExpenseTypeNameEn] IS NULL;
UPDATE [dbo].[ExpenseTypes] SET [ExpenseTypeNameEn] = N'Transport' WHERE [ExpenseTypeID] = 5 AND [ExpenseTypeNameEn] IS NULL;
UPDATE [dbo].[ExpenseTypes] SET [ExpenseTypeNameEn] = N'Maintenance' WHERE [ExpenseTypeID] = 6 AND [ExpenseTypeNameEn] IS NULL;
UPDATE [dbo].[ExpenseTypes] SET [ExpenseTypeNameEn] = N'Salaries' WHERE [ExpenseTypeID] = 7 AND [ExpenseTypeNameEn] IS NULL;
UPDATE [dbo].[ExpenseTypes] SET [ExpenseTypeNameEn] = N'Supplies' WHERE [ExpenseTypeID] = 8 AND [ExpenseTypeNameEn] IS NULL;
UPDATE [dbo].[ExpenseTypes] SET [ExpenseTypeNameEn] = N'Other expenses' WHERE [ExpenseTypeID] = 9 AND [ExpenseTypeNameEn] IS NULL;
UPDATE [dbo].[Customers] SET [CustomerNameEn] = N'Cash customer' WHERE [CustomerID] = 1 AND [CustomerNameEn] IS NULL;
UPDATE [dbo].[CashBoxes] SET [BoxNameEn] = N'Main treasury' WHERE [CashBoxID] = 1 AND [BoxNameEn] IS NULL;
UPDATE [dbo].[CashBoxes] SET [BoxNameEn] = N'Cashier box' WHERE [CashBoxID] = 2 AND [BoxNameEn] IS NULL;
UPDATE [dbo].[Warehouses] SET [WarehouseNameEn] = N'Main warehouse' WHERE [WarehouseID] = 1 AND [WarehouseNameEn] IS NULL;
UPDATE [dbo].[StockParties] SET [PartyNameEn] = N'Waste and damage' WHERE [StockPartyID] = 1 AND [PartyNameEn] IS NULL;
GO

/*------------------------------------------------------------------------------
  7. Relationships
------------------------------------------------------------------------------*/
IF OBJECT_ID(N'dbo.FK_Settings_DefaultCustomerID', N'F') IS NULL
    ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [FK_Settings_DefaultCustomerID] FOREIGN KEY ([DefaultCustomerID]) REFERENCES [dbo].[Customers] ([CustomerID]);
IF OBJECT_ID(N'dbo.FK_Settings_DefaultBankID', N'F') IS NULL
    ALTER TABLE [dbo].[Settings] WITH CHECK ADD CONSTRAINT [FK_Settings_DefaultBankID] FOREIGN KEY ([DefaultBankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_RolePermissions_RoleID', N'F') IS NULL
    ALTER TABLE [dbo].[RolePermissions] WITH CHECK ADD CONSTRAINT [FK_RolePermissions_RoleID] FOREIGN KEY ([RoleID]) REFERENCES [dbo].[Roles] ([RoleID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_RolePermissions_PermissionKey', N'F') IS NULL
    ALTER TABLE [dbo].[RolePermissions] WITH CHECK ADD CONSTRAINT [FK_RolePermissions_PermissionKey] FOREIGN KEY ([PermissionKey]) REFERENCES [dbo].[Permissions] ([PermissionKey]) ON DELETE CASCADE ON UPDATE CASCADE;
IF OBJECT_ID(N'dbo.FK_Employees_RoleID', N'F') IS NULL
    ALTER TABLE [dbo].[Employees] WITH CHECK ADD CONSTRAINT [FK_Employees_RoleID] FOREIGN KEY ([RoleID]) REFERENCES [dbo].[Roles] ([RoleID]);
IF OBJECT_ID(N'dbo.FK_Employees_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[Employees] WITH CHECK ADD CONSTRAINT [FK_Employees_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_Employees_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[Employees] WITH CHECK ADD CONSTRAINT [FK_Employees_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_Employees_WarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[Employees] WITH CHECK ADD CONSTRAINT [FK_Employees_WarehouseID] FOREIGN KEY ([WarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_UserScreens_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[UserScreens] WITH CHECK ADD CONSTRAINT [FK_UserScreens_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_UserScreens_ScreenName', N'F') IS NULL
    ALTER TABLE [dbo].[UserScreens] WITH CHECK ADD CONSTRAINT [FK_UserScreens_ScreenName] FOREIGN KEY ([ScreenName]) REFERENCES [dbo].[Screens] ([ScreenName]) ON UPDATE CASCADE;
IF OBJECT_ID(N'dbo.FK_Activations_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[Activations] WITH CHECK ADD CONSTRAINT [FK_Activations_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_CurrencyRates_CurrencyCode', N'F') IS NULL
    ALTER TABLE [dbo].[CurrencyRates] WITH CHECK ADD CONSTRAINT [FK_CurrencyRates_CurrencyCode] FOREIGN KEY ([CurrencyCode]) REFERENCES [dbo].[Currencies] ([CurrencyCode]) ON UPDATE CASCADE;
IF OBJECT_ID(N'dbo.FK_CashBoxes_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[CashBoxes] WITH CHECK ADD CONSTRAINT [FK_CashBoxes_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]);
IF OBJECT_ID(N'dbo.FK_CashBoxUsers_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[CashBoxUsers] WITH CHECK ADD CONSTRAINT [FK_CashBoxUsers_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_CashBoxUsers_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[CashBoxUsers] WITH CHECK ADD CONSTRAINT [FK_CashBoxUsers_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_Warehouses_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[Warehouses] WITH CHECK ADD CONSTRAINT [FK_Warehouses_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_Warehouses_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[Warehouses] WITH CHECK ADD CONSTRAINT [FK_Warehouses_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]);
IF OBJECT_ID(N'dbo.FK_Suppliers_CurrencyCode', N'F') IS NULL
    ALTER TABLE [dbo].[Suppliers] WITH CHECK ADD CONSTRAINT [FK_Suppliers_CurrencyCode] FOREIGN KEY ([CurrencyCode]) REFERENCES [dbo].[Currencies] ([CurrencyCode]) ON UPDATE CASCADE;
IF OBJECT_ID(N'dbo.FK_Customers_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[Customers] WITH CHECK ADD CONSTRAINT [FK_Customers_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]);
IF OBJECT_ID(N'dbo.FK_Customers_PriceListID', N'F') IS NULL
    ALTER TABLE [dbo].[Customers] WITH CHECK ADD CONSTRAINT [FK_Customers_PriceListID] FOREIGN KEY ([PriceListID]) REFERENCES [dbo].[PriceLists] ([PriceListID]);
IF OBJECT_ID(N'dbo.FK_Products_CategoryID', N'F') IS NULL
    ALTER TABLE [dbo].[Products] WITH CHECK ADD CONSTRAINT [FK_Products_CategoryID] FOREIGN KEY ([CategoryID]) REFERENCES [dbo].[Categories] ([CategoryID]);
IF OBJECT_ID(N'dbo.FK_Products_UnitID', N'F') IS NULL
    ALTER TABLE [dbo].[Products] WITH CHECK ADD CONSTRAINT [FK_Products_UnitID] FOREIGN KEY ([UnitID]) REFERENCES [dbo].[Units] ([UnitID]);
IF OBJECT_ID(N'dbo.FK_Products_SupplierID', N'F') IS NULL
    ALTER TABLE [dbo].[Products] WITH CHECK ADD CONSTRAINT [FK_Products_SupplierID] FOREIGN KEY ([SupplierID]) REFERENCES [dbo].[Suppliers] ([SupplierID]);
IF OBJECT_ID(N'dbo.FK_ProductUnits_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[ProductUnits] WITH CHECK ADD CONSTRAINT [FK_ProductUnits_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_ProductUnits_UnitID', N'F') IS NULL
    ALTER TABLE [dbo].[ProductUnits] WITH CHECK ADD CONSTRAINT [FK_ProductUnits_UnitID] FOREIGN KEY ([UnitID]) REFERENCES [dbo].[Units] ([UnitID]);
IF OBJECT_ID(N'dbo.FK_PriceListItems_PriceListID', N'F') IS NULL
    ALTER TABLE [dbo].[PriceListItems] WITH CHECK ADD CONSTRAINT [FK_PriceListItems_PriceListID] FOREIGN KEY ([PriceListID]) REFERENCES [dbo].[PriceLists] ([PriceListID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_PriceListItems_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[PriceListItems] WITH CHECK ADD CONSTRAINT [FK_PriceListItems_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_PriceListItems_UnitID', N'F') IS NULL
    ALTER TABLE [dbo].[PriceListItems] WITH CHECK ADD CONSTRAINT [FK_PriceListItems_UnitID] FOREIGN KEY ([UnitID]) REFERENCES [dbo].[Units] ([UnitID]);
IF OBJECT_ID(N'dbo.FK_SalesInvoices_CustomerID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [FK_SalesInvoices_CustomerID] FOREIGN KEY ([CustomerID]) REFERENCES [dbo].[Customers] ([CustomerID]);
IF OBJECT_ID(N'dbo.FK_SalesInvoices_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [FK_SalesInvoices_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_SalesInvoices_PaymentMethodID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [FK_SalesInvoices_PaymentMethodID] FOREIGN KEY ([PaymentMethodID]) REFERENCES [dbo].[PaymentMethods] ([PaymentMethodID]);
IF OBJECT_ID(N'dbo.FK_SalesInvoices_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [FK_SalesInvoices_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_SalesInvoices_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [FK_SalesInvoices_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_SalesInvoices_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [FK_SalesInvoices_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_SalesInvoices_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [FK_SalesInvoices_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]);
IF OBJECT_ID(N'dbo.FK_SalesInvoices_WarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesInvoices] WITH CHECK ADD CONSTRAINT [FK_SalesInvoices_WarehouseID] FOREIGN KEY ([WarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_SalesInvoiceDetails_SalesInvoiceID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] WITH CHECK ADD CONSTRAINT [FK_SalesInvoiceDetails_SalesInvoiceID] FOREIGN KEY ([SalesInvoiceID]) REFERENCES [dbo].[SalesInvoices] ([SalesInvoiceID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_SalesInvoiceDetails_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] WITH CHECK ADD CONSTRAINT [FK_SalesInvoiceDetails_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_SalesInvoiceDetails_UnitID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesInvoiceDetails] WITH CHECK ADD CONSTRAINT [FK_SalesInvoiceDetails_UnitID] FOREIGN KEY ([UnitID]) REFERENCES [dbo].[Units] ([UnitID]);
IF OBJECT_ID(N'dbo.FK_SalesReturns_SalesInvoiceID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [FK_SalesReturns_SalesInvoiceID] FOREIGN KEY ([SalesInvoiceID]) REFERENCES [dbo].[SalesInvoices] ([SalesInvoiceID]);
IF OBJECT_ID(N'dbo.FK_SalesReturns_CustomerID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [FK_SalesReturns_CustomerID] FOREIGN KEY ([CustomerID]) REFERENCES [dbo].[Customers] ([CustomerID]);
IF OBJECT_ID(N'dbo.FK_SalesReturns_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [FK_SalesReturns_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_SalesReturns_PaymentMethodID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [FK_SalesReturns_PaymentMethodID] FOREIGN KEY ([PaymentMethodID]) REFERENCES [dbo].[PaymentMethods] ([PaymentMethodID]);
IF OBJECT_ID(N'dbo.FK_SalesReturns_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [FK_SalesReturns_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_SalesReturns_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [FK_SalesReturns_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_SalesReturns_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [FK_SalesReturns_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_SalesReturns_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [FK_SalesReturns_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]);
IF OBJECT_ID(N'dbo.FK_SalesReturns_WarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReturns] WITH CHECK ADD CONSTRAINT [FK_SalesReturns_WarehouseID] FOREIGN KEY ([WarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_SalesReturnDetails_SalesReturnID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [FK_SalesReturnDetails_SalesReturnID] FOREIGN KEY ([SalesReturnID]) REFERENCES [dbo].[SalesReturns] ([SalesReturnID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_SalesReturnDetails_SalesDetailID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [FK_SalesReturnDetails_SalesDetailID] FOREIGN KEY ([SalesDetailID]) REFERENCES [dbo].[SalesInvoiceDetails] ([SalesDetailID]);
IF OBJECT_ID(N'dbo.FK_SalesReturnDetails_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [FK_SalesReturnDetails_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_SalesReturnDetails_UnitID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReturnDetails] WITH CHECK ADD CONSTRAINT [FK_SalesReturnDetails_UnitID] FOREIGN KEY ([UnitID]) REFERENCES [dbo].[Units] ([UnitID]);
IF OBJECT_ID(N'dbo.FK_PurchaseInvoices_SupplierID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [FK_PurchaseInvoices_SupplierID] FOREIGN KEY ([SupplierID]) REFERENCES [dbo].[Suppliers] ([SupplierID]);
IF OBJECT_ID(N'dbo.FK_PurchaseInvoices_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [FK_PurchaseInvoices_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_PurchaseInvoices_PaymentMethodID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [FK_PurchaseInvoices_PaymentMethodID] FOREIGN KEY ([PaymentMethodID]) REFERENCES [dbo].[PaymentMethods] ([PaymentMethodID]);
IF OBJECT_ID(N'dbo.FK_PurchaseInvoices_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [FK_PurchaseInvoices_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_PurchaseInvoices_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [FK_PurchaseInvoices_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_PurchaseInvoices_CurrencyCode', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [FK_PurchaseInvoices_CurrencyCode] FOREIGN KEY ([CurrencyCode]) REFERENCES [dbo].[Currencies] ([CurrencyCode]) ON UPDATE CASCADE;
IF OBJECT_ID(N'dbo.FK_PurchaseInvoices_WarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoices] WITH CHECK ADD CONSTRAINT [FK_PurchaseInvoices_WarehouseID] FOREIGN KEY ([WarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_PurchaseInvoiceDetails_PurchaseInvoiceID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] WITH CHECK ADD CONSTRAINT [FK_PurchaseInvoiceDetails_PurchaseInvoiceID] FOREIGN KEY ([PurchaseInvoiceID]) REFERENCES [dbo].[PurchaseInvoices] ([PurchaseInvoiceID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_PurchaseInvoiceDetails_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] WITH CHECK ADD CONSTRAINT [FK_PurchaseInvoiceDetails_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_PurchaseInvoiceDetails_UnitID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseInvoiceDetails] WITH CHECK ADD CONSTRAINT [FK_PurchaseInvoiceDetails_UnitID] FOREIGN KEY ([UnitID]) REFERENCES [dbo].[Units] ([UnitID]);
IF OBJECT_ID(N'dbo.FK_PurchaseReturns_PurchaseInvoiceID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [FK_PurchaseReturns_PurchaseInvoiceID] FOREIGN KEY ([PurchaseInvoiceID]) REFERENCES [dbo].[PurchaseInvoices] ([PurchaseInvoiceID]);
IF OBJECT_ID(N'dbo.FK_PurchaseReturns_SupplierID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [FK_PurchaseReturns_SupplierID] FOREIGN KEY ([SupplierID]) REFERENCES [dbo].[Suppliers] ([SupplierID]);
IF OBJECT_ID(N'dbo.FK_PurchaseReturns_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [FK_PurchaseReturns_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_PurchaseReturns_PaymentMethodID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [FK_PurchaseReturns_PaymentMethodID] FOREIGN KEY ([PaymentMethodID]) REFERENCES [dbo].[PaymentMethods] ([PaymentMethodID]);
IF OBJECT_ID(N'dbo.FK_PurchaseReturns_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [FK_PurchaseReturns_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_PurchaseReturns_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [FK_PurchaseReturns_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_PurchaseReturns_CurrencyCode', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [FK_PurchaseReturns_CurrencyCode] FOREIGN KEY ([CurrencyCode]) REFERENCES [dbo].[Currencies] ([CurrencyCode]) ON UPDATE CASCADE;
IF OBJECT_ID(N'dbo.FK_PurchaseReturns_WarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseReturns] WITH CHECK ADD CONSTRAINT [FK_PurchaseReturns_WarehouseID] FOREIGN KEY ([WarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_PurchaseReturnDetails_PurchaseReturnID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] WITH CHECK ADD CONSTRAINT [FK_PurchaseReturnDetails_PurchaseReturnID] FOREIGN KEY ([PurchaseReturnID]) REFERENCES [dbo].[PurchaseReturns] ([PurchaseReturnID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_PurchaseReturnDetails_PurchaseDetailID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] WITH CHECK ADD CONSTRAINT [FK_PurchaseReturnDetails_PurchaseDetailID] FOREIGN KEY ([PurchaseDetailID]) REFERENCES [dbo].[PurchaseInvoiceDetails] ([PurchaseDetailID]);
IF OBJECT_ID(N'dbo.FK_PurchaseReturnDetails_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] WITH CHECK ADD CONSTRAINT [FK_PurchaseReturnDetails_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_PurchaseReturnDetails_UnitID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseReturnDetails] WITH CHECK ADD CONSTRAINT [FK_PurchaseReturnDetails_UnitID] FOREIGN KEY ([UnitID]) REFERENCES [dbo].[Units] ([UnitID]);
IF OBJECT_ID(N'dbo.FK_CustomerPayments_CustomerID', N'F') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] WITH CHECK ADD CONSTRAINT [FK_CustomerPayments_CustomerID] FOREIGN KEY ([CustomerID]) REFERENCES [dbo].[Customers] ([CustomerID]);
IF OBJECT_ID(N'dbo.FK_CustomerPayments_PaymentMethodID', N'F') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] WITH CHECK ADD CONSTRAINT [FK_CustomerPayments_PaymentMethodID] FOREIGN KEY ([PaymentMethodID]) REFERENCES [dbo].[PaymentMethods] ([PaymentMethodID]);
IF OBJECT_ID(N'dbo.FK_CustomerPayments_SalesInvoiceID', N'F') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] WITH CHECK ADD CONSTRAINT [FK_CustomerPayments_SalesInvoiceID] FOREIGN KEY ([SalesInvoiceID]) REFERENCES [dbo].[SalesInvoices] ([SalesInvoiceID]);
IF OBJECT_ID(N'dbo.FK_CustomerPayments_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] WITH CHECK ADD CONSTRAINT [FK_CustomerPayments_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_CustomerPayments_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] WITH CHECK ADD CONSTRAINT [FK_CustomerPayments_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_CustomerPayments_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] WITH CHECK ADD CONSTRAINT [FK_CustomerPayments_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_CustomerPayments_CurrencyCode', N'F') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] WITH CHECK ADD CONSTRAINT [FK_CustomerPayments_CurrencyCode] FOREIGN KEY ([CurrencyCode]) REFERENCES [dbo].[Currencies] ([CurrencyCode]) ON UPDATE CASCADE;
IF OBJECT_ID(N'dbo.FK_CustomerPayments_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[CustomerPayments] WITH CHECK ADD CONSTRAINT [FK_CustomerPayments_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]);
IF OBJECT_ID(N'dbo.FK_SupplierPayments_SupplierID', N'F') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] WITH CHECK ADD CONSTRAINT [FK_SupplierPayments_SupplierID] FOREIGN KEY ([SupplierID]) REFERENCES [dbo].[Suppliers] ([SupplierID]);
IF OBJECT_ID(N'dbo.FK_SupplierPayments_PaymentMethodID', N'F') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] WITH CHECK ADD CONSTRAINT [FK_SupplierPayments_PaymentMethodID] FOREIGN KEY ([PaymentMethodID]) REFERENCES [dbo].[PaymentMethods] ([PaymentMethodID]);
IF OBJECT_ID(N'dbo.FK_SupplierPayments_PurchaseInvoiceID', N'F') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] WITH CHECK ADD CONSTRAINT [FK_SupplierPayments_PurchaseInvoiceID] FOREIGN KEY ([PurchaseInvoiceID]) REFERENCES [dbo].[PurchaseInvoices] ([PurchaseInvoiceID]);
IF OBJECT_ID(N'dbo.FK_SupplierPayments_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] WITH CHECK ADD CONSTRAINT [FK_SupplierPayments_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_SupplierPayments_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] WITH CHECK ADD CONSTRAINT [FK_SupplierPayments_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_SupplierPayments_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] WITH CHECK ADD CONSTRAINT [FK_SupplierPayments_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_SupplierPayments_CurrencyCode', N'F') IS NULL
    ALTER TABLE [dbo].[SupplierPayments] WITH CHECK ADD CONSTRAINT [FK_SupplierPayments_CurrencyCode] FOREIGN KEY ([CurrencyCode]) REFERENCES [dbo].[Currencies] ([CurrencyCode]) ON UPDATE CASCADE;
IF OBJECT_ID(N'dbo.FK_BankTransactions_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[BankTransactions] WITH CHECK ADD CONSTRAINT [FK_BankTransactions_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_BankTransactions_ToBankID', N'F') IS NULL
    ALTER TABLE [dbo].[BankTransactions] WITH CHECK ADD CONSTRAINT [FK_BankTransactions_ToBankID] FOREIGN KEY ([ToBankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_BankTransactions_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[BankTransactions] WITH CHECK ADD CONSTRAINT [FK_BankTransactions_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_BankTransactions_CounterAccount', N'F') IS NULL
    ALTER TABLE [dbo].[BankTransactions] WITH CHECK ADD CONSTRAINT [FK_BankTransactions_CounterAccount] FOREIGN KEY ([CounterAccount]) REFERENCES [dbo].[Accounts] ([AccountCode]);
IF OBJECT_ID(N'dbo.FK_BankTransactions_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[BankTransactions] WITH CHECK ADD CONSTRAINT [FK_BankTransactions_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_Cheques_CustomerID', N'F') IS NULL
    ALTER TABLE [dbo].[Cheques] WITH CHECK ADD CONSTRAINT [FK_Cheques_CustomerID] FOREIGN KEY ([CustomerID]) REFERENCES [dbo].[Customers] ([CustomerID]);
IF OBJECT_ID(N'dbo.FK_Cheques_SupplierID', N'F') IS NULL
    ALTER TABLE [dbo].[Cheques] WITH CHECK ADD CONSTRAINT [FK_Cheques_SupplierID] FOREIGN KEY ([SupplierID]) REFERENCES [dbo].[Suppliers] ([SupplierID]);
IF OBJECT_ID(N'dbo.FK_Cheques_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[Cheques] WITH CHECK ADD CONSTRAINT [FK_Cheques_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_Cheques_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[Cheques] WITH CHECK ADD CONSTRAINT [FK_Cheques_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_FixedAssets_AssetAccount', N'F') IS NULL
    ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [FK_FixedAssets_AssetAccount] FOREIGN KEY ([AssetAccount]) REFERENCES [dbo].[Accounts] ([AccountCode]);
IF OBJECT_ID(N'dbo.FK_FixedAssets_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [FK_FixedAssets_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_FixedAssets_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [FK_FixedAssets_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_FixedAssets_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [FK_FixedAssets_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_FixedAssets_CounterAccount', N'F') IS NULL
    ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [FK_FixedAssets_CounterAccount] FOREIGN KEY ([CounterAccount]) REFERENCES [dbo].[Accounts] ([AccountCode]);
IF OBJECT_ID(N'dbo.FK_FixedAssets_DisposalBankID', N'F') IS NULL
    ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [FK_FixedAssets_DisposalBankID] FOREIGN KEY ([DisposalBankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_FixedAssets_DisposalCashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [FK_FixedAssets_DisposalCashBoxID] FOREIGN KEY ([DisposalCashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_FixedAssets_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[FixedAssets] WITH CHECK ADD CONSTRAINT [FK_FixedAssets_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_AssetDepreciations_RunID', N'F') IS NULL
    ALTER TABLE [dbo].[AssetDepreciations] WITH CHECK ADD CONSTRAINT [FK_AssetDepreciations_RunID] FOREIGN KEY ([RunID]) REFERENCES [dbo].[DepreciationRuns] ([RunID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_AssetDepreciations_AssetID', N'F') IS NULL
    ALTER TABLE [dbo].[AssetDepreciations] WITH CHECK ADD CONSTRAINT [FK_AssetDepreciations_AssetID] FOREIGN KEY ([AssetID]) REFERENCES [dbo].[FixedAssets] ([AssetID]);
IF OBJECT_ID(N'dbo.FK_AssetRevaluations_AssetID', N'F') IS NULL
    ALTER TABLE [dbo].[AssetRevaluations] WITH CHECK ADD CONSTRAINT [FK_AssetRevaluations_AssetID] FOREIGN KEY ([AssetID]) REFERENCES [dbo].[FixedAssets] ([AssetID]);
IF OBJECT_ID(N'dbo.FK_AssetTransfers_AssetID', N'F') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] WITH CHECK ADD CONSTRAINT [FK_AssetTransfers_AssetID] FOREIGN KEY ([AssetID]) REFERENCES [dbo].[FixedAssets] ([AssetID]);
IF OBJECT_ID(N'dbo.FK_AssetTransfers_FromCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] WITH CHECK ADD CONSTRAINT [FK_AssetTransfers_FromCenterID] FOREIGN KEY ([FromCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_AssetTransfers_ToCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[AssetTransfers] WITH CHECK ADD CONSTRAINT [FK_AssetTransfers_ToCenterID] FOREIGN KEY ([ToCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_SalesReps_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReps] WITH CHECK ADD CONSTRAINT [FK_SalesReps_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_SalesReps_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesReps] WITH CHECK ADD CONSTRAINT [FK_SalesReps_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_SalesRepTargets_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesRepTargets] WITH CHECK ADD CONSTRAINT [FK_SalesRepTargets_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_CommissionLines_CommissionRunID', N'F') IS NULL
    ALTER TABLE [dbo].[CommissionLines] WITH CHECK ADD CONSTRAINT [FK_CommissionLines_CommissionRunID] FOREIGN KEY ([CommissionRunID]) REFERENCES [dbo].[CommissionRuns] ([CommissionRunID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_CommissionLines_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[CommissionLines] WITH CHECK ADD CONSTRAINT [FK_CommissionLines_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]);
IF OBJECT_ID(N'dbo.FK_CommissionLines_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[CommissionLines] WITH CHECK ADD CONSTRAINT [FK_CommissionLines_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_BudgetLines_BudgetID', N'F') IS NULL
    ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [FK_BudgetLines_BudgetID] FOREIGN KEY ([BudgetID]) REFERENCES [dbo].[Budgets] ([BudgetID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_BudgetLines_AccountCode', N'F') IS NULL
    ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [FK_BudgetLines_AccountCode] FOREIGN KEY ([AccountCode]) REFERENCES [dbo].[Accounts] ([AccountCode]);
IF OBJECT_ID(N'dbo.FK_BudgetLines_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[BudgetLines] WITH CHECK ADD CONSTRAINT [FK_BudgetLines_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_PayrollRuns_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] WITH CHECK ADD CONSTRAINT [FK_PayrollRuns_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_PayrollRuns_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] WITH CHECK ADD CONSTRAINT [FK_PayrollRuns_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_PayrollRuns_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[PayrollRuns] WITH CHECK ADD CONSTRAINT [FK_PayrollRuns_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_PayrollLines_PayrollRunID', N'F') IS NULL
    ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [FK_PayrollLines_PayrollRunID] FOREIGN KEY ([PayrollRunID]) REFERENCES [dbo].[PayrollRuns] ([PayrollRunID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_PayrollLines_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [FK_PayrollLines_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_PayrollLines_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[PayrollLines] WITH CHECK ADD CONSTRAINT [FK_PayrollLines_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_BankReconciliations_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[BankReconciliations] WITH CHECK ADD CONSTRAINT [FK_BankReconciliations_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_BankClearings_ReconciliationID', N'F') IS NULL
    ALTER TABLE [dbo].[BankClearings] WITH CHECK ADD CONSTRAINT [FK_BankClearings_ReconciliationID] FOREIGN KEY ([ReconciliationID]) REFERENCES [dbo].[BankReconciliations] ([ReconciliationID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_BankClearings_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[BankClearings] WITH CHECK ADD CONSTRAINT [FK_BankClearings_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_CustomerAllocations_PaymentID', N'F') IS NULL
    ALTER TABLE [dbo].[CustomerAllocations] WITH CHECK ADD CONSTRAINT [FK_CustomerAllocations_PaymentID] FOREIGN KEY ([PaymentID]) REFERENCES [dbo].[CustomerPayments] ([PaymentID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_CustomerAllocations_SalesInvoiceID', N'F') IS NULL
    ALTER TABLE [dbo].[CustomerAllocations] WITH CHECK ADD CONSTRAINT [FK_CustomerAllocations_SalesInvoiceID] FOREIGN KEY ([SalesInvoiceID]) REFERENCES [dbo].[SalesInvoices] ([SalesInvoiceID]);
IF OBJECT_ID(N'dbo.FK_SupplierAllocations_PaymentID', N'F') IS NULL
    ALTER TABLE [dbo].[SupplierAllocations] WITH CHECK ADD CONSTRAINT [FK_SupplierAllocations_PaymentID] FOREIGN KEY ([PaymentID]) REFERENCES [dbo].[SupplierPayments] ([PaymentID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_SupplierAllocations_PurchaseInvoiceID', N'F') IS NULL
    ALTER TABLE [dbo].[SupplierAllocations] WITH CHECK ADD CONSTRAINT [FK_SupplierAllocations_PurchaseInvoiceID] FOREIGN KEY ([PurchaseInvoiceID]) REFERENCES [dbo].[PurchaseInvoices] ([PurchaseInvoiceID]);
IF OBJECT_ID(N'dbo.FK_Expenses_ExpenseTypeID', N'F') IS NULL
    ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [FK_Expenses_ExpenseTypeID] FOREIGN KEY ([ExpenseTypeID]) REFERENCES [dbo].[ExpenseTypes] ([ExpenseTypeID]);
IF OBJECT_ID(N'dbo.FK_Expenses_PaymentMethodID', N'F') IS NULL
    ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [FK_Expenses_PaymentMethodID] FOREIGN KEY ([PaymentMethodID]) REFERENCES [dbo].[PaymentMethods] ([PaymentMethodID]);
IF OBJECT_ID(N'dbo.FK_Expenses_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [FK_Expenses_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_Expenses_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [FK_Expenses_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_Expenses_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [FK_Expenses_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_Expenses_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [FK_Expenses_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_Expenses_RecurringID', N'F') IS NULL
    ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [FK_Expenses_RecurringID] FOREIGN KEY ([RecurringID]) REFERENCES [dbo].[RecurringExpenses] ([RecurringID]);
IF OBJECT_ID(N'dbo.FK_Expenses_CurrencyCode', N'F') IS NULL
    ALTER TABLE [dbo].[Expenses] WITH CHECK ADD CONSTRAINT [FK_Expenses_CurrencyCode] FOREIGN KEY ([CurrencyCode]) REFERENCES [dbo].[Currencies] ([CurrencyCode]) ON UPDATE CASCADE;
IF OBJECT_ID(N'dbo.FK_RecurringExpenses_ExpenseTypeID', N'F') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] WITH CHECK ADD CONSTRAINT [FK_RecurringExpenses_ExpenseTypeID] FOREIGN KEY ([ExpenseTypeID]) REFERENCES [dbo].[ExpenseTypes] ([ExpenseTypeID]);
IF OBJECT_ID(N'dbo.FK_RecurringExpenses_PaymentMethodID', N'F') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] WITH CHECK ADD CONSTRAINT [FK_RecurringExpenses_PaymentMethodID] FOREIGN KEY ([PaymentMethodID]) REFERENCES [dbo].[PaymentMethods] ([PaymentMethodID]);
IF OBJECT_ID(N'dbo.FK_RecurringExpenses_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] WITH CHECK ADD CONSTRAINT [FK_RecurringExpenses_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_RecurringExpenses_BankID', N'F') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] WITH CHECK ADD CONSTRAINT [FK_RecurringExpenses_BankID] FOREIGN KEY ([BankID]) REFERENCES [dbo].[Banks] ([BankID]);
IF OBJECT_ID(N'dbo.FK_RecurringExpenses_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[RecurringExpenses] WITH CHECK ADD CONSTRAINT [FK_RecurringExpenses_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_CashVouchers_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[CashVouchers] WITH CHECK ADD CONSTRAINT [FK_CashVouchers_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_CashVouchers_ToCashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[CashVouchers] WITH CHECK ADD CONSTRAINT [FK_CashVouchers_ToCashBoxID] FOREIGN KEY ([ToCashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_CashVouchers_ExpenseID', N'F') IS NULL
    ALTER TABLE [dbo].[CashVouchers] WITH CHECK ADD CONSTRAINT [FK_CashVouchers_ExpenseID] FOREIGN KEY ([ExpenseID]) REFERENCES [dbo].[Expenses] ([ExpenseID]);
IF OBJECT_ID(N'dbo.FK_CashVouchers_ClosingID', N'F') IS NULL
    ALTER TABLE [dbo].[CashVouchers] WITH CHECK ADD CONSTRAINT [FK_CashVouchers_ClosingID] FOREIGN KEY ([ClosingID]) REFERENCES [dbo].[CashClosings] ([ClosingID]);
IF OBJECT_ID(N'dbo.FK_CashVouchers_AdvanceEmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[CashVouchers] WITH CHECK ADD CONSTRAINT [FK_CashVouchers_AdvanceEmployeeID] FOREIGN KEY ([AdvanceEmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_CashVouchers_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[CashVouchers] WITH CHECK ADD CONSTRAINT [FK_CashVouchers_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_CashVouchers_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[CashVouchers] WITH CHECK ADD CONSTRAINT [FK_CashVouchers_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_CashVouchers_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[CashVouchers] WITH CHECK ADD CONSTRAINT [FK_CashVouchers_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]);
IF OBJECT_ID(N'dbo.FK_CashClosings_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[CashClosings] WITH CHECK ADD CONSTRAINT [FK_CashClosings_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_CashClosings_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[CashClosings] WITH CHECK ADD CONSTRAINT [FK_CashClosings_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_CashClosings_ToCashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[CashClosings] WITH CHECK ADD CONSTRAINT [FK_CashClosings_ToCashBoxID] FOREIGN KEY ([ToCashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_JournalEntries_SourceType', N'F') IS NULL
    ALTER TABLE [dbo].[JournalEntries] WITH CHECK ADD CONSTRAINT [FK_JournalEntries_SourceType] FOREIGN KEY ([SourceType]) REFERENCES [dbo].[JournalSourceTypes] ([SourceType]) ON UPDATE CASCADE;
IF OBJECT_ID(N'dbo.FK_JournalEntries_CurrencyCode', N'F') IS NULL
    ALTER TABLE [dbo].[JournalEntries] WITH CHECK ADD CONSTRAINT [FK_JournalEntries_CurrencyCode] FOREIGN KEY ([CurrencyCode]) REFERENCES [dbo].[Currencies] ([CurrencyCode]) ON UPDATE CASCADE;
IF OBJECT_ID(N'dbo.FK_JournalLines_EntryID', N'F') IS NULL
    ALTER TABLE [dbo].[JournalLines] WITH CHECK ADD CONSTRAINT [FK_JournalLines_EntryID] FOREIGN KEY ([EntryID]) REFERENCES [dbo].[JournalEntries] ([EntryID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_JournalLines_AccountCode', N'F') IS NULL
    ALTER TABLE [dbo].[JournalLines] WITH CHECK ADD CONSTRAINT [FK_JournalLines_AccountCode] FOREIGN KEY ([AccountCode]) REFERENCES [dbo].[Accounts] ([AccountCode]);
IF OBJECT_ID(N'dbo.FK_JournalLines_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[JournalLines] WITH CHECK ADD CONSTRAINT [FK_JournalLines_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_FiscalYearClosingLines_YearClosingID', N'F') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosingLines] WITH CHECK ADD CONSTRAINT [FK_FiscalYearClosingLines_YearClosingID] FOREIGN KEY ([YearClosingID]) REFERENCES [dbo].[FiscalYearClosings] ([YearClosingID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_FiscalYearClosingLines_AccountCode', N'F') IS NULL
    ALTER TABLE [dbo].[FiscalYearClosingLines] WITH CHECK ADD CONSTRAINT [FK_FiscalYearClosingLines_AccountCode] FOREIGN KEY ([AccountCode]) REFERENCES [dbo].[Accounts] ([AccountCode]);
IF OBJECT_ID(N'dbo.FK_VatReturns_PaidAccount', N'F') IS NULL
    ALTER TABLE [dbo].[VatReturns] WITH CHECK ADD CONSTRAINT [FK_VatReturns_PaidAccount] FOREIGN KEY ([PaidAccount]) REFERENCES [dbo].[Accounts] ([AccountCode]);
IF OBJECT_ID(N'dbo.FK_ManualEntries_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[ManualEntries] WITH CHECK ADD CONSTRAINT [FK_ManualEntries_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_ManualEntries_CurrencyCode', N'F') IS NULL
    ALTER TABLE [dbo].[ManualEntries] WITH CHECK ADD CONSTRAINT [FK_ManualEntries_CurrencyCode] FOREIGN KEY ([CurrencyCode]) REFERENCES [dbo].[Currencies] ([CurrencyCode]) ON UPDATE CASCADE;
IF OBJECT_ID(N'dbo.FK_ManualEntryLines_ManualEntryID', N'F') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] WITH CHECK ADD CONSTRAINT [FK_ManualEntryLines_ManualEntryID] FOREIGN KEY ([ManualEntryID]) REFERENCES [dbo].[ManualEntries] ([ManualEntryID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_ManualEntryLines_AccountCode', N'F') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] WITH CHECK ADD CONSTRAINT [FK_ManualEntryLines_AccountCode] FOREIGN KEY ([AccountCode]) REFERENCES [dbo].[Accounts] ([AccountCode]);
IF OBJECT_ID(N'dbo.FK_ManualEntryLines_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[ManualEntryLines] WITH CHECK ADD CONSTRAINT [FK_ManualEntryLines_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_OpeningBalances_AccountCode', N'F') IS NULL
    ALTER TABLE [dbo].[OpeningBalances] WITH CHECK ADD CONSTRAINT [FK_OpeningBalances_AccountCode] FOREIGN KEY ([AccountCode]) REFERENCES [dbo].[Accounts] ([AccountCode]);
IF OBJECT_ID(N'dbo.FK_OpeningStock_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[OpeningStock] WITH CHECK ADD CONSTRAINT [FK_OpeningStock_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_OpeningStock_WarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[OpeningStock] WITH CHECK ADD CONSTRAINT [FK_OpeningStock_WarehouseID] FOREIGN KEY ([WarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_InventoryTransactions_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] WITH CHECK ADD CONSTRAINT [FK_InventoryTransactions_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_InventoryTransactions_TransactionTypeID', N'F') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] WITH CHECK ADD CONSTRAINT [FK_InventoryTransactions_TransactionTypeID] FOREIGN KEY ([TransactionTypeID]) REFERENCES [dbo].[TransactionTypes] ([TransactionTypeID]);
IF OBJECT_ID(N'dbo.FK_InventoryTransactions_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] WITH CHECK ADD CONSTRAINT [FK_InventoryTransactions_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_InventoryTransactions_WarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[InventoryTransactions] WITH CHECK ADD CONSTRAINT [FK_InventoryTransactions_WarehouseID] FOREIGN KEY ([WarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_StockTransfers_FromWarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[StockTransfers] WITH CHECK ADD CONSTRAINT [FK_StockTransfers_FromWarehouseID] FOREIGN KEY ([FromWarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_StockTransfers_ToWarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[StockTransfers] WITH CHECK ADD CONSTRAINT [FK_StockTransfers_ToWarehouseID] FOREIGN KEY ([ToWarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_StockTransferLines_StockTransferID', N'F') IS NULL
    ALTER TABLE [dbo].[StockTransferLines] WITH CHECK ADD CONSTRAINT [FK_StockTransferLines_StockTransferID] FOREIGN KEY ([StockTransferID]) REFERENCES [dbo].[StockTransfers] ([StockTransferID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_StockTransferLines_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[StockTransferLines] WITH CHECK ADD CONSTRAINT [FK_StockTransferLines_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_StockParties_AccountCode', N'F') IS NULL
    ALTER TABLE [dbo].[StockParties] WITH CHECK ADD CONSTRAINT [FK_StockParties_AccountCode] FOREIGN KEY ([AccountCode]) REFERENCES [dbo].[Accounts] ([AccountCode]);
IF OBJECT_ID(N'dbo.FK_StockParties_CostCenterID', N'F') IS NULL
    ALTER TABLE [dbo].[StockParties] WITH CHECK ADD CONSTRAINT [FK_StockParties_CostCenterID] FOREIGN KEY ([CostCenterID]) REFERENCES [dbo].[CostCenters] ([CostCenterID]);
IF OBJECT_ID(N'dbo.FK_StockDocuments_WarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[StockDocuments] WITH CHECK ADD CONSTRAINT [FK_StockDocuments_WarehouseID] FOREIGN KEY ([WarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_StockDocuments_StockPartyID', N'F') IS NULL
    ALTER TABLE [dbo].[StockDocuments] WITH CHECK ADD CONSTRAINT [FK_StockDocuments_StockPartyID] FOREIGN KEY ([StockPartyID]) REFERENCES [dbo].[StockParties] ([StockPartyID]);
IF OBJECT_ID(N'dbo.FK_StockDocumentLines_StockDocumentID', N'F') IS NULL
    ALTER TABLE [dbo].[StockDocumentLines] WITH CHECK ADD CONSTRAINT [FK_StockDocumentLines_StockDocumentID] FOREIGN KEY ([StockDocumentID]) REFERENCES [dbo].[StockDocuments] ([StockDocumentID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_StockDocumentLines_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[StockDocumentLines] WITH CHECK ADD CONSTRAINT [FK_StockDocumentLines_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_StockNotes_WarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[StockNotes] WITH CHECK ADD CONSTRAINT [FK_StockNotes_WarehouseID] FOREIGN KEY ([WarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_StockNotes_CustomerID', N'F') IS NULL
    ALTER TABLE [dbo].[StockNotes] WITH CHECK ADD CONSTRAINT [FK_StockNotes_CustomerID] FOREIGN KEY ([CustomerID]) REFERENCES [dbo].[Customers] ([CustomerID]);
IF OBJECT_ID(N'dbo.FK_StockNotes_SupplierID', N'F') IS NULL
    ALTER TABLE [dbo].[StockNotes] WITH CHECK ADD CONSTRAINT [FK_StockNotes_SupplierID] FOREIGN KEY ([SupplierID]) REFERENCES [dbo].[Suppliers] ([SupplierID]);
IF OBJECT_ID(N'dbo.FK_StockNotes_SalesOrderID', N'F') IS NULL
    ALTER TABLE [dbo].[StockNotes] WITH CHECK ADD CONSTRAINT [FK_StockNotes_SalesOrderID] FOREIGN KEY ([SalesOrderID]) REFERENCES [dbo].[SalesOrders] ([SalesOrderID]);
IF OBJECT_ID(N'dbo.FK_StockNotes_PurchaseOrderID', N'F') IS NULL
    ALTER TABLE [dbo].[StockNotes] WITH CHECK ADD CONSTRAINT [FK_StockNotes_PurchaseOrderID] FOREIGN KEY ([PurchaseOrderID]) REFERENCES [dbo].[PurchaseOrders] ([PurchaseOrderID]);
IF OBJECT_ID(N'dbo.FK_StockNotes_SalesInvoiceID', N'F') IS NULL
    ALTER TABLE [dbo].[StockNotes] WITH CHECK ADD CONSTRAINT [FK_StockNotes_SalesInvoiceID] FOREIGN KEY ([SalesInvoiceID]) REFERENCES [dbo].[SalesInvoices] ([SalesInvoiceID]);
IF OBJECT_ID(N'dbo.FK_StockNotes_PurchaseInvoiceID', N'F') IS NULL
    ALTER TABLE [dbo].[StockNotes] WITH CHECK ADD CONSTRAINT [FK_StockNotes_PurchaseInvoiceID] FOREIGN KEY ([PurchaseInvoiceID]) REFERENCES [dbo].[PurchaseInvoices] ([PurchaseInvoiceID]);
IF OBJECT_ID(N'dbo.FK_StockNoteLines_StockNoteID', N'F') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] WITH CHECK ADD CONSTRAINT [FK_StockNoteLines_StockNoteID] FOREIGN KEY ([StockNoteID]) REFERENCES [dbo].[StockNotes] ([StockNoteID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_StockNoteLines_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] WITH CHECK ADD CONSTRAINT [FK_StockNoteLines_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_StockNoteLines_SalesOrderLineID', N'F') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] WITH CHECK ADD CONSTRAINT [FK_StockNoteLines_SalesOrderLineID] FOREIGN KEY ([SalesOrderLineID]) REFERENCES [dbo].[SalesOrderLines] ([SalesOrderLineID]);
IF OBJECT_ID(N'dbo.FK_StockNoteLines_PurchaseOrderLineID', N'F') IS NULL
    ALTER TABLE [dbo].[StockNoteLines] WITH CHECK ADD CONSTRAINT [FK_StockNoteLines_PurchaseOrderLineID] FOREIGN KEY ([PurchaseOrderLineID]) REFERENCES [dbo].[PurchaseOrderLines] ([PurchaseOrderLineID]);
IF OBJECT_ID(N'dbo.FK_SalesOrders_CustomerID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesOrders] WITH CHECK ADD CONSTRAINT [FK_SalesOrders_CustomerID] FOREIGN KEY ([CustomerID]) REFERENCES [dbo].[Customers] ([CustomerID]);
IF OBJECT_ID(N'dbo.FK_SalesOrders_WarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesOrders] WITH CHECK ADD CONSTRAINT [FK_SalesOrders_WarehouseID] FOREIGN KEY ([WarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_SalesOrders_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesOrders] WITH CHECK ADD CONSTRAINT [FK_SalesOrders_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]);
IF OBJECT_ID(N'dbo.FK_Quotations_CustomerID', N'F') IS NULL
    ALTER TABLE [dbo].[Quotations] WITH CHECK ADD CONSTRAINT [FK_Quotations_CustomerID] FOREIGN KEY ([CustomerID]) REFERENCES [dbo].[Customers] ([CustomerID]);
IF OBJECT_ID(N'dbo.FK_Quotations_SalesOrderID', N'F') IS NULL
    ALTER TABLE [dbo].[Quotations] WITH CHECK ADD CONSTRAINT [FK_Quotations_SalesOrderID] FOREIGN KEY ([SalesOrderID]) REFERENCES [dbo].[SalesOrders] ([SalesOrderID]);
IF OBJECT_ID(N'dbo.FK_Quotations_SalesInvoiceID', N'F') IS NULL
    ALTER TABLE [dbo].[Quotations] WITH CHECK ADD CONSTRAINT [FK_Quotations_SalesInvoiceID] FOREIGN KEY ([SalesInvoiceID]) REFERENCES [dbo].[SalesInvoices] ([SalesInvoiceID]);
IF OBJECT_ID(N'dbo.FK_QuotationLines_QuotationID', N'F') IS NULL
    ALTER TABLE [dbo].[QuotationLines] WITH CHECK ADD CONSTRAINT [FK_QuotationLines_QuotationID] FOREIGN KEY ([QuotationID]) REFERENCES [dbo].[Quotations] ([QuotationID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_QuotationLines_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[QuotationLines] WITH CHECK ADD CONSTRAINT [FK_QuotationLines_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_PurchaseOrders_SupplierID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] WITH CHECK ADD CONSTRAINT [FK_PurchaseOrders_SupplierID] FOREIGN KEY ([SupplierID]) REFERENCES [dbo].[Suppliers] ([SupplierID]);
IF OBJECT_ID(N'dbo.FK_PurchaseOrders_WarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseOrders] WITH CHECK ADD CONSTRAINT [FK_PurchaseOrders_WarehouseID] FOREIGN KEY ([WarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_PurchaseOrderLines_PurchaseOrderID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseOrderLines] WITH CHECK ADD CONSTRAINT [FK_PurchaseOrderLines_PurchaseOrderID] FOREIGN KEY ([PurchaseOrderID]) REFERENCES [dbo].[PurchaseOrders] ([PurchaseOrderID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_PurchaseOrderLines_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[PurchaseOrderLines] WITH CHECK ADD CONSTRAINT [FK_PurchaseOrderLines_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_RepDays_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[RepDays] WITH CHECK ADD CONSTRAINT [FK_RepDays_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]);
IF OBJECT_ID(N'dbo.FK_RepDays_WarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[RepDays] WITH CHECK ADD CONSTRAINT [FK_RepDays_WarehouseID] FOREIGN KEY ([WarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_RepDays_CashBoxID', N'F') IS NULL
    ALTER TABLE [dbo].[RepDays] WITH CHECK ADD CONSTRAINT [FK_RepDays_CashBoxID] FOREIGN KEY ([CashBoxID]) REFERENCES [dbo].[CashBoxes] ([CashBoxID]);
IF OBJECT_ID(N'dbo.FK_RepDays_StoreID', N'F') IS NULL
    ALTER TABLE [dbo].[RepDays] WITH CHECK ADD CONSTRAINT [FK_RepDays_StoreID] FOREIGN KEY ([StoreID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_RepDays_StockCountID', N'F') IS NULL
    ALTER TABLE [dbo].[RepDays] WITH CHECK ADD CONSTRAINT [FK_RepDays_StockCountID] FOREIGN KEY ([StockCountID]) REFERENCES [dbo].[StockCounts] ([StockCountID]);
IF OBJECT_ID(N'dbo.FK_RepDays_ClosingID', N'F') IS NULL
    ALTER TABLE [dbo].[RepDays] WITH CHECK ADD CONSTRAINT [FK_RepDays_ClosingID] FOREIGN KEY ([ClosingID]) REFERENCES [dbo].[CashClosings] ([ClosingID]);
IF OBJECT_ID(N'dbo.FK_RouteStops_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[RouteStops] WITH CHECK ADD CONSTRAINT [FK_RouteStops_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_RouteStops_CustomerID', N'F') IS NULL
    ALTER TABLE [dbo].[RouteStops] WITH CHECK ADD CONSTRAINT [FK_RouteStops_CustomerID] FOREIGN KEY ([CustomerID]) REFERENCES [dbo].[Customers] ([CustomerID]);
IF OBJECT_ID(N'dbo.FK_RepVisits_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[RepVisits] WITH CHECK ADD CONSTRAINT [FK_RepVisits_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]);
IF OBJECT_ID(N'dbo.FK_RepVisits_CustomerID', N'F') IS NULL
    ALTER TABLE [dbo].[RepVisits] WITH CHECK ADD CONSTRAINT [FK_RepVisits_CustomerID] FOREIGN KEY ([CustomerID]) REFERENCES [dbo].[Customers] ([CustomerID]);
IF OBJECT_ID(N'dbo.FK_RepPrices_SalesRepID', N'F') IS NULL
    ALTER TABLE [dbo].[RepPrices] WITH CHECK ADD CONSTRAINT [FK_RepPrices_SalesRepID] FOREIGN KEY ([SalesRepID]) REFERENCES [dbo].[SalesReps] ([SalesRepID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_RepPrices_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[RepPrices] WITH CHECK ADD CONSTRAINT [FK_RepPrices_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_SalesOrderLines_SalesOrderID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesOrderLines] WITH CHECK ADD CONSTRAINT [FK_SalesOrderLines_SalesOrderID] FOREIGN KEY ([SalesOrderID]) REFERENCES [dbo].[SalesOrders] ([SalesOrderID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_SalesOrderLines_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[SalesOrderLines] WITH CHECK ADD CONSTRAINT [FK_SalesOrderLines_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_StockCounts_CategoryID', N'F') IS NULL
    ALTER TABLE [dbo].[StockCounts] WITH CHECK ADD CONSTRAINT [FK_StockCounts_CategoryID] FOREIGN KEY ([CategoryID]) REFERENCES [dbo].[Categories] ([CategoryID]);
IF OBJECT_ID(N'dbo.FK_StockCounts_EmployeeID', N'F') IS NULL
    ALTER TABLE [dbo].[StockCounts] WITH CHECK ADD CONSTRAINT [FK_StockCounts_EmployeeID] FOREIGN KEY ([EmployeeID]) REFERENCES [dbo].[Employees] ([EmployeeID]);
IF OBJECT_ID(N'dbo.FK_StockCounts_WarehouseID', N'F') IS NULL
    ALTER TABLE [dbo].[StockCounts] WITH CHECK ADD CONSTRAINT [FK_StockCounts_WarehouseID] FOREIGN KEY ([WarehouseID]) REFERENCES [dbo].[Warehouses] ([WarehouseID]);
IF OBJECT_ID(N'dbo.FK_StockCountDetails_StockCountID', N'F') IS NULL
    ALTER TABLE [dbo].[StockCountDetails] WITH CHECK ADD CONSTRAINT [FK_StockCountDetails_StockCountID] FOREIGN KEY ([StockCountID]) REFERENCES [dbo].[StockCounts] ([StockCountID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_StockCountDetails_ProductID', N'F') IS NULL
    ALTER TABLE [dbo].[StockCountDetails] WITH CHECK ADD CONSTRAINT [FK_StockCountDetails_ProductID] FOREIGN KEY ([ProductID]) REFERENCES [dbo].[Products] ([ProductID]);
IF OBJECT_ID(N'dbo.FK_AuditChanges_LogID', N'F') IS NULL
    ALTER TABLE [dbo].[AuditChanges] WITH CHECK ADD CONSTRAINT [FK_AuditChanges_LogID] FOREIGN KEY ([LogID]) REFERENCES [dbo].[AuditLog] ([LogID]) ON DELETE CASCADE;
IF OBJECT_ID(N'dbo.FK_AssetRevaluations_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[AssetRevaluations] DROP CONSTRAINT [FK_AssetRevaluations_EmployeeID];
IF OBJECT_ID(N'dbo.FK_AssetTransfers_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[AssetTransfers] DROP CONSTRAINT [FK_AssetTransfers_EmployeeID];
IF OBJECT_ID(N'dbo.FK_AuditLog_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[AuditLog] DROP CONSTRAINT [FK_AuditLog_EmployeeID];
IF OBJECT_ID(N'dbo.FK_BankReconciliations_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[BankReconciliations] DROP CONSTRAINT [FK_BankReconciliations_EmployeeID];
IF OBJECT_ID(N'dbo.FK_Budgets_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[Budgets] DROP CONSTRAINT [FK_Budgets_EmployeeID];
IF OBJECT_ID(N'dbo.FK_CashVouchers_ReceivedByID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[CashVouchers] DROP CONSTRAINT [FK_CashVouchers_ReceivedByID];
IF OBJECT_ID(N'dbo.FK_CommissionRuns_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[CommissionRuns] DROP CONSTRAINT [FK_CommissionRuns_EmployeeID];
IF OBJECT_ID(N'dbo.FK_CustomerAllocations_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[CustomerAllocations] DROP CONSTRAINT [FK_CustomerAllocations_EmployeeID];
IF OBJECT_ID(N'dbo.FK_DepreciationRuns_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[DepreciationRuns] DROP CONSTRAINT [FK_DepreciationRuns_EmployeeID];
IF OBJECT_ID(N'dbo.FK_FiscalYearClosings_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[FiscalYearClosings] DROP CONSTRAINT [FK_FiscalYearClosings_EmployeeID];
IF OBJECT_ID(N'dbo.FK_PeriodClosings_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[PeriodClosings] DROP CONSTRAINT [FK_PeriodClosings_EmployeeID];
IF OBJECT_ID(N'dbo.FK_PurchaseOrders_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[PurchaseOrders] DROP CONSTRAINT [FK_PurchaseOrders_EmployeeID];
IF OBJECT_ID(N'dbo.FK_Quotations_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[Quotations] DROP CONSTRAINT [FK_Quotations_EmployeeID];
IF OBJECT_ID(N'dbo.FK_RepDays_AdvanceEmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[RepDays] DROP CONSTRAINT [FK_RepDays_AdvanceEmployeeID];
IF OBJECT_ID(N'dbo.FK_RepDays_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[RepDays] DROP CONSTRAINT [FK_RepDays_EmployeeID];
IF OBJECT_ID(N'dbo.FK_RepVisits_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[RepVisits] DROP CONSTRAINT [FK_RepVisits_EmployeeID];
IF OBJECT_ID(N'dbo.FK_SalesOrders_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[SalesOrders] DROP CONSTRAINT [FK_SalesOrders_EmployeeID];
IF OBJECT_ID(N'dbo.FK_StockCounts_PostedByID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[StockCounts] DROP CONSTRAINT [FK_StockCounts_PostedByID];
IF OBJECT_ID(N'dbo.FK_StockDocuments_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[StockDocuments] DROP CONSTRAINT [FK_StockDocuments_EmployeeID];
IF OBJECT_ID(N'dbo.FK_StockNotes_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[StockNotes] DROP CONSTRAINT [FK_StockNotes_EmployeeID];
IF OBJECT_ID(N'dbo.FK_StockTransfers_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[StockTransfers] DROP CONSTRAINT [FK_StockTransfers_EmployeeID];
IF OBJECT_ID(N'dbo.FK_StockTransfers_ReceivedByID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[StockTransfers] DROP CONSTRAINT [FK_StockTransfers_ReceivedByID];
IF OBJECT_ID(N'dbo.FK_SupplierAllocations_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[SupplierAllocations] DROP CONSTRAINT [FK_SupplierAllocations_EmployeeID];
IF OBJECT_ID(N'dbo.FK_VatReturns_EmployeeID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[VatReturns] DROP CONSTRAINT [FK_VatReturns_EmployeeID];
IF OBJECT_ID(N'dbo.FK_Warehouses_KeeperID', N'F') IS NOT NULL
    ALTER TABLE [dbo].[Warehouses] DROP CONSTRAINT [FK_Warehouses_KeeperID];
GO

/*------------------------------------------------------------------------------
  8. Database options (plan 5C, docs/70)
------------------------------------------------------------------------------*/
IF EXISTS (SELECT 1 FROM sys.databases WHERE database_id = DB_ID() AND is_read_committed_snapshot_on = 0)
    ALTER DATABASE CURRENT SET READ_COMMITTED_SNAPSHOT ON WITH ROLLBACK IMMEDIATE;
GO
PRINT N'RetailStore schema: done (97 tables, 1237 fields, 148 indexes, 231 relationships, 451 seed rows).';
GO
