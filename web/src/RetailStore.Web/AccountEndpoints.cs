using System.Security.Claims;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using RetailStore.UI.Services;

namespace RetailStore.Web;

/// <summary>The cookie of a signed-in user, the logout and the language switch.</summary>
public static class AccountEndpoints
{
    public static async Task SignInAsync(HttpContext http, int employeeId, string name, string computer, bool mustChange)
    {
        var claims = new List<Claim>
        {
            new(UserInfo.EmployeeIdClaim, employeeId.ToString()),
            new(ClaimTypes.Name, name),
            new(UserInfo.ComputerClaim, computer),
        };
        if (mustChange)
        {
            claims.Add(new Claim(UserInfo.MustChangeClaim, "1"));
        }
        var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
        await http.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(identity));
    }

    /// <summary>AuditLog.ComputerName of a web operation: "WEB " + the address of the browser.</summary>
    public static string ComputerOf(HttpContext http)
    {
        string computer = "WEB " + (http.Connection.RemoteIpAddress?.ToString() ?? "");
        return computer.Length <= 50 ? computer : computer[..50];
    }

    /// <summary>Only a local path: never send the user to another site after the login.</summary>
    public static string SafeReturn(string? returnUrl) =>
        !string.IsNullOrEmpty(returnUrl) && returnUrl.StartsWith('/') && !returnUrl.StartsWith("//") && !returnUrl.StartsWith("/\\")
            ? returnUrl : "/";

    public static void MapAccountEndpoints(this WebApplication app)
    {
        app.MapPost("/account/logout", async (HttpContext http) =>
        {
            await http.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
            return Results.LocalRedirect("~/account/login");
        });
        app.MapGet("/lang/{code}", (string code, string? returnUrl, HttpContext http) =>
        {
            http.Response.Cookies.Append(UiLanguage.CookieName, code == "en" ? "en" : "ar",
                new CookieOptions { Expires = DateTimeOffset.Now.AddYears(1), IsEssential = true, HttpOnly = true, SameSite = SameSiteMode.Lax });
            return Results.LocalRedirect("~" + SafeReturn(returnUrl));
        });
    }
}
