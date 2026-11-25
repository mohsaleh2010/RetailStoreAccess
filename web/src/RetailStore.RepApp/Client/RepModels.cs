namespace RetailStore.RepApp.Client;

// The JSON of the API (docs/77, docs/78) as the phone reads it. The server's records are in RetailStore.Data, which the
// phone never references (no database on the phone but its own SQLite file); the tests check both read the same JSON.

public sealed record TokenResponse(string AccessToken, string TokenType, DateTime ExpiresAt, int EmployeeID, string EmployeeName,
                                   bool MustChangePassword);

public sealed record CreatedId(int Id, string Message);

public sealed record RepProfile(int SalesRepID, string RepCode, string RepName, string? Region, int? VanWarehouseID,
                                string? VanWarehouseName, int? CustodyBoxID, string? CustodyBoxName, int OrderWarehouseID,
                                decimal CreditLimit, decimal MaxDiscountPercent, decimal CustomersBalance);

public sealed record RepCustomer(int CustomerID, string CustomerName, string? Mobile, string? Phone, string? City,
                                 string? District, string? Address, string? VATNumber, decimal CurrentBalance,
                                 bool AllowCredit, decimal CreditLimit, int? PriceListID, bool Assigned, int[] RouteWeekdays);

public sealed record RepProduct(int ProductID, string ProductCode, string? Barcode, string ProductName, int UnitID,
                                string VATCategory, decimal SellingPrice, decimal Price, bool RepPrice, bool TrackStock,
                                decimal VanQuantity);

public sealed record RepVisit(int RepVisitID, DateTime VisitDate, int CustomerID, string CustomerName, bool IsPlanned,
                              short StopOrder, string Result, DateTime? VisitedAt, string? Notes);

public sealed record OrderLineView(int SalesOrderLineID, int ProductID, string ProductName, decimal Quantity, decimal UnitPrice);

public sealed record OrderView(int SalesOrderID, string OrderNumber, DateTime OrderDate, int CustomerID, string CustomerName,
                               int WarehouseID, int? SalesRepID, DateTime? RequiredDate, string Status, string? Notes,
                               decimal Total, OrderLineView[] Lines, string? ClientRef);

/// <summary>GET /api/v1/rep/sync: everything the phone keeps to work offline.</summary>
public sealed record RepSnapshot(DateTime ServerTime, string StoreName, RepProfile Profile, List<RepCustomer> Customers, List<RepProduct> Products,
                                 Dictionary<string, Dictionary<int, decimal>> PriceGroups, Dictionary<int, string> CustomerPriceGroups,
                                 List<RepVisit> Visits, List<OrderView> Orders, bool CanVisits, bool DayOpen = false);

public sealed record OrderLineInput(int ProductID, decimal Quantity, decimal? UnitPrice);

public sealed record OrderRequest(int CustomerID, int? WarehouseID, DateTime? RequiredDate, string? Notes, OrderLineInput[] Lines,
                                  string? ClientRef, DateTime? OrderDate);

// ---- the invoice and the receipt of the rep (docs/87): made on the program at once, each with its number on the phone

public sealed record CartLine(int ProductID, decimal Quantity, decimal UnitPrice);

public sealed record RepSaleRequest(string ClientRef, int CustomerID, string PaymentType, int? PaymentMethodID, decimal InvoiceDiscount,
                                    string? Notes, CartLine[] Lines, bool ConfirmStock = false);

public sealed record RepPaymentRequest(string ClientRef, int CustomerID, decimal Amount, int PaymentMethodID, string? Notes,
                                       bool ConfirmOverBalance = false);

public sealed record DocLine(string Text, decimal Quantity, decimal Amount);

public sealed record DocRow(string Label, decimal Amount);

/// <summary>The document the program made (PosDocument on the server): SALE or PAYMENT, to print.</summary>
public sealed record RepDocument(string Kind, int Id, string Number, DateTime Date, string Title, string CashierName, string PartyName,
                                 DocLine[] Lines, DocRow[] Rows, string? Note, string? Qr);
