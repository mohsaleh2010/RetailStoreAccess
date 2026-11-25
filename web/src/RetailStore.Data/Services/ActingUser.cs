namespace RetailStore.Data.Services;

/// <summary>The open rep day a user works in (modVanSales.StartActingAsRep): the van is the warehouse, the custody box is the
/// box and the rep of the day is the rep who sells (the rep app, docs/87).</summary>
public sealed record RepDayContext(int RepDayID, int SalesRepID, int WarehouseID, int CashBoxID);

/// <summary>Who does an operation (TempVars UserID in Access) and from where (AuditLog.ComputerName); RepDay = acting as the rep
/// of an open rep day (null = the user's own warehouse, box and rep).</summary>
public sealed record ActingUser(int EmployeeID, string Computer)
{
    public static readonly ActingUser Nobody = new(0, "WEB");

    public RepDayContext? RepDay { get; init; }
}
