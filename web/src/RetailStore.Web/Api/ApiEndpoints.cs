using System.ComponentModel;
using System.Text.Json;
using Microsoft.AspNetCore.Authorization;
using Microsoft.EntityFrameworkCore;
using RetailStore.Data;
using RetailStore.Data.DataScreens;
using RetailStore.Data.Services;
using RetailStore.Domain;
using RetailStore.UI.Services;

namespace RetailStore.Web.Api;

public sealed record TokenRequest(
    [property: Description("The user name of the program")] string Username,
    [property: Description("The password; empty only for the first login of admin")] string? Password,
    [property: Description("The name of the device or app (written in the audit trail)")] string? Device);

public sealed record TokenResponse(string AccessToken, string TokenType, DateTime ExpiresAt, int EmployeeID, string EmployeeName,
                                   bool MustChangePassword);

public sealed record PasswordRequest(string? OldPassword, string NewPassword, string ConfirmPassword);

public sealed record MeResponse(int EmployeeID, string EmployeeName, bool MustChangePassword, string[] Permissions,
                                string[] Screens, RepProfile? Rep);

public sealed record PlanRequest(DateTime? Date);

public sealed record VisitRequest(int CustomerID, DateTime? Date);

public sealed record VisitResultRequest(
    [property: Description("ORDER, SALE, PAYMENT, NO_SALE or CLOSED; empty = read from the customer's documents that day")] string? Result,
    string? Notes);

public sealed record OrderRequest(int CustomerID, int? WarehouseID, DateTime? RequiredDate, string? Notes, OrderLineInput[] Lines,
                                  string? ClientRef = null, DateTime? OrderDate = null);

public sealed record CreatedId(int Id, string Message);

/// <summary>The REST API of the system (docs/77), version 1 under /api/v1: tokens, the data screens and the work of the
/// sales rep. Every operation goes through the same services as the web screens, so the same checks, messages,
/// permissions and audit trail as Access.</summary>
public static class ApiEndpoints
{
    public const string Prefix = "/api/v1";

    public static RouteHandlerBuilder Needs(this RouteHandlerBuilder b, string key, string? screen = null, ScreenAction? action = null) =>
        b.WithMetadata(new ApiPermission(key, screen, action));

    private static ActingUser Acting(HttpContext http) => UserInfo.Acting(http.User);

    /// <summary>AuditLog.ComputerName of an API call: "API " + the device the app names + its address.</summary>
    public static string ComputerOf(HttpContext http, string? device)
    {
        string name = new string((device ?? "").Where(c => !char.IsControl(c)).ToArray()).Trim();
        string computer = ("API " + (name.Length > 0 ? name + " " : "") + (http.Connection.RemoteIpAddress?.ToString() ?? "")).Trim();
        return computer.Length <= 50 ? computer : computer[..50];
    }

    public static void MapRetailApi(this WebApplication app)
    {
        var api = app.MapGroup(Prefix).WithTags("RetailStore");

        // ---------------------------------------------------------------- tokens
        api.MapPost("/auth/token", async (TokenRequest body, HttpContext http, SecurityService security,
                                          IDbContextFactory<RetailDbContext> factory, ApiTokens tokens) =>
        {
            string computer = ComputerOf(http, body.Device);
            var r = await security.LoginAsync(body.Username ?? "", body.Password ?? "", computer, channel: "api");
            if (!r.Ok)
            {
                return ApiResults.Problem(http, StatusCodes.Status401Unauthorized, "login_failed", r.Message);
            }
            var (token, expires) = tokens.Create(r.EmployeeID, r.EmployeeName, computer, r.MustChangePassword,
                                                 await StampAsync(factory, r.EmployeeID));
            return Results.Ok(new TokenResponse(token, "Bearer", expires, r.EmployeeID, r.EmployeeName, r.MustChangePassword));
        })
        .AllowAnonymous().RequireRateLimiting(ApiSetup.TokenLimit)
        .WithName("CreateToken").WithSummary("Log in: a token for the other operations")
        .WithDescription("modSecurity.LoginUser: the same users and passwords as Access, 5 wrong passwords lock the account " +
                         "for 15 minutes, every login is in the audit trail with the device. Send the token in the header " +
                         "Authorization: Bearer followed by the token. A user who must change his password can only call /auth/password and /me.")
        .Produces<TokenResponse>().ProducesProblem(StatusCodes.Status401Unauthorized).ProducesProblem(StatusCodes.Status429TooManyRequests);

        var secured = api.MapGroup("").RequireAuthorization(new AuthorizeAttribute { AuthenticationSchemes = ApiTokens.Scheme })
            .AddEndpointFilter<ApiUserFilter>().AddEndpointFilter<PermissionFilter>()
            .ProducesProblem(StatusCodes.Status401Unauthorized).ProducesProblem(StatusCodes.Status403Forbidden);

        secured.MapPost("/auth/password", async (PasswordRequest body, HttpContext http, SecurityService security,
                                                 IDbContextFactory<RetailDbContext> factory, ApiTokens tokens) =>
        {
            var user = Acting(http);
            string problem = await security.ChangeOwnPasswordAsync(user, body.OldPassword ?? "", body.NewPassword ?? "",
                                                                   body.ConfirmPassword ?? "");
            if (problem.Length > 0)
            {
                return ApiResults.Refused(http, problem);
            }
            var (token, expires) = tokens.Create(user.EmployeeID, UserInfo.Name(http.User), user.Computer, false,
                                                 await StampAsync(factory, user.EmployeeID));
            return Results.Ok(new TokenResponse(token, "Bearer", expires, user.EmployeeID, UserInfo.Name(http.User), false));
        })
        .WithMetadata(new AllowWithPasswordChangeAttribute())
        .WithName("ChangePassword").WithSummary("Change the user's own password: a new token")
        .WithDescription("modSecurity.ChangeOwnPassword and the password rules of the program. The tokens given before stop working.")
        .Produces<TokenResponse>().ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        secured.MapGet("/me", async (HttpContext http, IDbContextFactory<RetailDbContext> factory, PermissionService permissions,
                                     RepService reps) =>
        {
            int id = UserInfo.EmployeeId(http.User);
            await using var db = await factory.CreateDbContextAsync();
            var e = await db.Employees.AsNoTracking().FirstAsync(x => x.EmployeeID == id);
            var keys = new List<string>();
            foreach (var key in await db.Permissions.AsNoTracking().OrderBy(p => p.PermissionKey).Select(p => p.PermissionKey).ToListAsync())
            {
                if (await permissions.HasPermissionAsync(id, key))
                {
                    keys.Add(key);
                }
            }
            var screens = new List<string>();
            foreach (var s in ScreenSpecs.All)
            {
                if (await permissions.CanOpenScreenAsync(id, s.Name))
                {
                    screens.Add(s.Route);
                }
            }
            return Results.Ok(new MeResponse(id, e.EmployeeName, e.MustChangePassword || e.PasswordHash is null, keys.ToArray(),
                                             screens.ToArray(), await reps.ProfileAsync(id, ApiResults.English(http))));
        })
        .WithMetadata(new AllowWithPasswordChangeAttribute())
        .WithName("Me").WithSummary("The user: his permissions, the screens he may open, and his rep")
        .Produces<MeResponse>();

        // ---------------------------------------------------------------- data screens
        MapScreen(secured, ScreenSpecs.Products, "PRODUCTS");
        MapScreen(secured, ScreenSpecs.Customers, "CUSTOMERS");
        MapScreen(secured, ScreenSpecs.Suppliers, "SUPPLIERS");
        MapScreen(secured, ScreenSpecs.Warehouses, "PRODUCTS");

        // ---------------------------------------------------------------- the sales rep
        RepApi.Map(secured.MapGroup("/rep").AddEndpointFilter<RepFilter>().WithTags("Sales rep")
                          .ProducesProblem(StatusCodes.Status403Forbidden));
    }

    private static async Task<string> StampAsync(IDbContextFactory<RetailDbContext> factory, int id)
    {
        await using var db = await factory.CreateDbContextAsync();
        var e = await db.Employees.AsNoTracking().Where(x => x.EmployeeID == id).Select(x => new { x.PasswordHash, x.PasswordSalt }).FirstAsync();
        return SecurityService.PasswordStamp(e.PasswordHash, e.PasswordSalt);
    }

    /// <summary>The operations of a data screen: list (search, inactive, pages), one record, look-ups, add, change,
    /// delete or deactivate. The permission of the screen (SCREEN_PERMISSIONS) and the actions of a user with his own
    /// screens (frmUserScreens) apply to each.</summary>
    private static void MapScreen<T>(RouteGroupBuilder secured, ScreenSpec<T> spec, string permission) where T : class, new()
    {
        var g = secured.MapGroup("/" + spec.Route).WithTags(spec.Route);
        string what = spec.Table;

        g.MapGet("", async (HttpContext http, DataScreenService screens, string? search, bool? includeInactive, int? skip, int? take) =>
        {
            var (total, rows) = await screens.RecordsAsync(spec, search, includeInactive == true, skip ?? 0, take ?? 100);
            return Results.Ok(new RecordPage(total, Math.Max(skip ?? 0, 0), rows.Count, rows.Select(r => RecordJson.ToJson(spec, r)).ToList()));
        })
        .Needs(permission, spec.Name).WithName("List" + what)
        .WithSummary("The records of " + what)
        .WithDescription("The same filter and search as the list of the screen " + spec.Name + " (search in " +
                         string.Join(", ", spec.Search) + "), in its order; take at most " + DataScreenService.MaxListRows + ".")
        .Produces<RecordPage>();

        g.MapGet("/{id:int}", async (int id, HttpContext http, DataScreenService screens) =>
        {
            var e = await screens.LoadAsync(spec, id);
            return e is null ? ApiResults.NotFound(http) : Results.Ok(RecordJson.ToJson(spec, e));
        })
        .Needs(permission, spec.Name).WithName("Get" + what).WithSummary("One record of " + what)
        .Produces<Dictionary<string, object?>>().ProducesProblem(StatusCodes.Status404NotFound);

        g.MapGet("/lookups", async (HttpContext http, DataScreenService screens) =>
        {
            var lists = new Dictionary<string, List<LookupItem>>();
            foreach (var f in spec.Fields.Where(f => f.Lookup is not null))
            {
                lists[f.Name] = await screens.LookupAsync(f.Lookup!, ApiResults.English(http));
            }
            foreach (var f in spec.Fields.Where(f => f.Options is not null))
            {
                lists[f.Name] = f.Options!.Select(o => new LookupItem(o.Value, ApiResults.Text(http, o.Text), true)).ToList();
            }
            return Results.Ok(lists);
        })
        .Needs(permission, spec.Name).WithName("Lookups" + what)
        .WithSummary("The choices of the fields of " + what + " (look-ups and value lists)")
        .Produces<Dictionary<string, List<LookupItem>>>();

        g.MapPost("", async (Dictionary<string, JsonElement> body, bool? confirm, HttpContext http, DataScreenService screens) =>
        {
            var e = await screens.NewAsync(spec);
            var (message, field) = RecordJson.Apply(spec, e, body);
            if (message.Length > 0)
            {
                return ApiResults.Refused(http, message, field);
            }
            var r = await screens.SaveAsync(spec, e, true, Acting(http), confirm == true);
            return await SavedAsync(http, screens, spec, r, created: true);
        })
        .Needs(permission, spec.Name, ScreenAction.Add).WithName("Add" + what).WithSummary("Add a record to " + what)
        .WithDescription(SaveText(spec)).Produces<SavedRecord>(StatusCodes.Status201Created)
        .ProducesProblem(StatusCodes.Status409Conflict).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        g.MapPut("/{id:int}", async (int id, Dictionary<string, JsonElement> body, bool? confirm, HttpContext http, DataScreenService screens) =>
        {
            var e = await screens.LoadAsync(spec, id);
            if (e is null)
            {
                return ApiResults.NotFound(http);
            }
            var (message, field) = RecordJson.Apply(spec, e, body);
            if (message.Length > 0)
            {
                return ApiResults.Refused(http, message, field);
            }
            var r = await screens.SaveAsync(spec, e, false, Acting(http), confirm == true);
            return await SavedAsync(http, screens, spec, r, created: false);
        })
        .Needs(permission, spec.Name, ScreenAction.Edit).WithName("Change" + what).WithSummary("Change a record of " + what)
        .WithDescription("Only the fields sent change; the others keep what the database holds. " + SaveText(spec))
        .Produces<SavedRecord>().ProducesProblem(StatusCodes.Status404NotFound)
        .ProducesProblem(StatusCodes.Status409Conflict).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        g.MapDelete("/{id:int}", async (int id, bool? deactivate, HttpContext http, DataScreenService screens) =>
        {
            if (await screens.LoadAsync(spec, id) is null)
            {
                return ApiResults.NotFound(http);
            }
            var r = await screens.DeleteAsync(spec, id, Acting(http), deactivate == true);
            return ApiResults.Of(http, r, () => Results.Ok(new SavedRecord(id, ApiResults.Text(http, r.Message), null)));
        })
        .Needs(permission, spec.Name, ScreenAction.Delete).WithName("Delete" + what).WithSummary("Delete a record of " + what)
        .WithDescription("modForms.DeleteRecord: a system row is never deleted; a record used by operations answers 409 " +
                         "(code confirm): send ?deactivate=true to deactivate it instead.")
        .Produces<SavedRecord>().ProducesProblem(StatusCodes.Status404NotFound)
        .ProducesProblem(StatusCodes.Status409Conflict).ProducesProblem(StatusCodes.Status422UnprocessableEntity);
    }

    private static string SaveText<T>(ScreenSpec<T> spec) where T : class, new() =>
        "The fields are the data fields of the screen " + spec.Name + " by their database names; the key and the locked " +
        "fields (" + string.Join(", ", spec.Fields.Where(f => f.Locked).Select(f => f.Name)) + ") are ignored. The checks " +
        "and their order are those of modForms.FormBeforeUpdate. A question (409, code confirm) is answered by sending " +
        "the same request with ?confirm=true. Audited (AuditLog / AuditChanges).";

    private static async Task<IResult> SavedAsync<T>(HttpContext http, DataScreenService screens, ScreenSpec<T> spec, SaveResult r,
                                                     bool created) where T : class, new()
    {
        if (r.Status != SaveStatus.Saved)
        {
            return ApiResults.Of(http, r, () => Results.Ok());
        }
        var now = await screens.LoadAsync(spec, r.Key);
        var body = new SavedRecord(r.Key, ApiResults.Text(http, r.Message), now is null ? null : RecordJson.ToJson(spec, now));
        return created ? Results.Created(Prefix + "/" + spec.Route + "/" + r.Key, body) : Results.Ok(body);
    }
}
