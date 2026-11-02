using System.Security.Claims;
using RetailStore.Data.Services;

namespace RetailStore.UI.Services;

/// <summary>The claims of the signed-in user (set at login, RetailStore.Web).</summary>
public static class UserInfo
{
    public const string EmployeeIdClaim = ClaimTypes.NameIdentifier;
    public const string ComputerClaim = "computer";
    public const string MustChangeClaim = "must_change_password";

    public static int EmployeeId(ClaimsPrincipal? user) =>
        int.TryParse(user?.FindFirst(EmployeeIdClaim)?.Value, out int id) ? id : 0;

    public static string Name(ClaimsPrincipal? user) => user?.Identity?.Name ?? "";

    public static bool MustChangePassword(ClaimsPrincipal? user) => user?.FindFirst(MustChangeClaim)?.Value == "1";

    public static ActingUser Acting(ClaimsPrincipal? user) =>
        new(EmployeeId(user), user?.FindFirst(ComputerClaim)?.Value ?? "WEB");
}
