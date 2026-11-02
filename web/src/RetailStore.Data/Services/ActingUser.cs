namespace RetailStore.Data.Services;

/// <summary>Who does an operation (TempVars UserID in Access) and from where (AuditLog.ComputerName).</summary>
public sealed record ActingUser(int EmployeeID, string Computer)
{
    public static readonly ActingUser Nobody = new(0, "WEB");
}
