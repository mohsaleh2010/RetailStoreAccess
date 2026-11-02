using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using RetailStore.Data;
using RetailStore.Data.Services;
using RetailStore.UI.Services;

namespace RetailStore.Web.Api;

/// <summary>The API tokens (JWT, docs/77): signed with HMAC-SHA256 by the key of appsettings.json (Api:JwtKey), else by
/// a random key kept in App_Data/api-jwt.key (made on the first run, never in the repository).</summary>
public sealed class ApiTokens
{
    public const string Issuer = "RetailStore";
    public const string Audience = "RetailStore.Api";
    public const string StampClaim = "pwd_stamp";
    public const string Scheme = "Bearer";

    public SymmetricSecurityKey Key { get; }
    public TimeSpan Lifetime { get; }

    public ApiTokens(IConfiguration config, IWebHostEnvironment env)
    {
        Key = new SymmetricSecurityKey(LoadKey(config["Api:JwtKey"], Path.Combine(env.ContentRootPath, "App_Data", "api-jwt.key")));
        Lifetime = TimeSpan.FromHours(double.TryParse(config["Api:TokenHours"], out var h) && h > 0 ? h : 12);
    }

    /// <summary>The configured key (at least 32 bytes as text), else the key file, made when missing.</summary>
    public static byte[] LoadKey(string? configured, string file)
    {
        if (!string.IsNullOrWhiteSpace(configured))
        {
            var bytes = Encoding.UTF8.GetBytes(configured);
            if (bytes.Length < 32)
            {
                throw new InvalidOperationException("Api:JwtKey must be at least 32 characters");
            }
            return bytes;
        }
        if (File.Exists(file))
        {
            return Convert.FromBase64String(File.ReadAllText(file).Trim());
        }
        Directory.CreateDirectory(Path.GetDirectoryName(file)!);
        var key = RandomNumberGenerator.GetBytes(64);
        File.WriteAllText(file, Convert.ToBase64String(key));
        return key;
    }

    public TokenValidationParameters Validation() => new()
    {
        ValidateIssuer = true, ValidIssuer = Issuer,
        ValidateAudience = true, ValidAudience = Audience,
        ValidateIssuerSigningKey = true, IssuerSigningKey = Key,
        ValidateLifetime = true, ClockSkew = TimeSpan.FromMinutes(1),
        NameClaimType = ClaimTypes.Name,
    };

    public (string Token, DateTime ExpiresAt) Create(int employeeId, string name, string computer, bool mustChange, string stamp)
    {
        var expires = DateTime.UtcNow.Add(Lifetime);
        var claims = new List<Claim>
        {
            new(UserInfo.EmployeeIdClaim, employeeId.ToString()),
            new(ClaimTypes.Name, name),
            new(UserInfo.ComputerClaim, computer),
            new(StampClaim, stamp),
            new(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString("N")),
        };
        if (mustChange)
        {
            claims.Add(new Claim(UserInfo.MustChangeClaim, "1"));
        }
        var token = new JwtSecurityToken(Issuer, Audience, claims, DateTime.UtcNow, expires,
                                         new SigningCredentials(Key, SecurityAlgorithms.HmacSha256));
        return (new JwtSecurityTokenHandler().WriteToken(token), expires);
    }
}

/// <summary>Marks the endpoints a user who must change his password may still call (the password change, who am I).</summary>
public sealed class AllowWithPasswordChangeAttribute : Attribute;

/// <summary>Runs on every API endpoint that needs a token: the user still exists, is active and has the password the
/// token was given for (a new password, a disabled user: the token stops working at once); a user who must change his
/// password can only change it.</summary>
public sealed class ApiUserFilter(IDbContextFactory<RetailDbContext> factory) : IEndpointFilter
{
    public async ValueTask<object?> InvokeAsync(EndpointFilterInvocationContext context, EndpointFilterDelegate next)
    {
        var http = context.HttpContext;
        int id = UserInfo.EmployeeId(http.User);
        await using (var db = await factory.CreateDbContextAsync())
        {
            var e = await db.Employees.AsNoTracking().Where(x => x.EmployeeID == id)
                .Select(x => new { x.IsActive, x.PasswordHash, x.PasswordSalt, x.MustChangePassword }).FirstOrDefaultAsync();
            if (e is null || !e.IsActive ||
                SecurityService.PasswordStamp(e.PasswordHash, e.PasswordSalt) != http.User.FindFirst(ApiTokens.StampClaim)?.Value)
            {
                return ApiResults.Problem(http, StatusCodes.Status401Unauthorized, "unauthorized", "انتهت صلاحية الدخول. سجّل الدخول من جديد.");
            }
            bool mustChange = e.MustChangePassword || e.PasswordHash is null;
            if (mustChange && http.GetEndpoint()?.Metadata.GetMetadata<AllowWithPasswordChangeAttribute>() is null)
            {
                return ApiResults.Problem(http, StatusCodes.Status403Forbidden, "password_change_required",
                                          "يجب تغيير كلمة المرور قبل المتابعة.");
            }
        }
        return await next(context);
    }
}

/// <summary>The permission an API operation needs (docs/77): a permission key (modCommon.HasPermission), or a screen
/// with its action (modSecurity.CanOpenScreen / CanScreenAction). Shown in the OpenAPI document of each operation.</summary>
public sealed record ApiPermission(string? Key, string? Screen = null, ScreenAction? Action = null)
{
    public string Text => Screen is null ? Key! : Action is null ? Key + " (" + Screen + ")" : Key + " (" + Screen + ": " + Action + ")";
}

public sealed class PermissionFilter(PermissionService permissions) : IEndpointFilter
{
    public async ValueTask<object?> InvokeAsync(EndpointFilterInvocationContext context, EndpointFilterDelegate next)
    {
        var http = context.HttpContext;
        int id = UserInfo.EmployeeId(http.User);
        foreach (var p in http.GetEndpoint()?.Metadata.GetOrderedMetadata<ApiPermission>() ?? [])
        {
            string problem;
            if (p.Screen is not null)
            {
                problem = await permissions.OpenProblemAsync(id, p.Screen);
                if (problem.Length == 0 && p.Action is { } action)
                {
                    problem = await permissions.ActionProblemAsync(id, p.Screen, action);
                }
            }
            else
            {
                problem = await permissions.HasPermissionAsync(id, p.Key!) ? "" : "لا تملك صلاحية هذه العملية. راجع مدير النظام.";
            }
            if (problem.Length > 0)
            {
                return ApiResults.Problem(http, StatusCodes.Status403Forbidden, "forbidden", problem);
            }
        }
        return await next(context);
    }
}

/// <summary>The rep endpoints work for the rep linked to the user (SalesReps.EmployeeID): his id goes into the request.</summary>
public sealed class RepFilter(RepService reps) : IEndpointFilter
{
    public const string RepIdItem = "SalesRepID";

    public async ValueTask<object?> InvokeAsync(EndpointFilterInvocationContext context, EndpointFilterDelegate next)
    {
        var http = context.HttpContext;
        int rep = await reps.RepIdOfAsync(UserInfo.EmployeeId(http.User));
        if (rep == 0)
        {
            return ApiResults.Problem(http, StatusCodes.Status403Forbidden, "not_a_rep", RepService.NotARep);
        }
        http.Items[RepIdItem] = rep;
        return await next(context);
    }

    public static int RepOf(HttpContext http) => (int)http.Items[RepIdItem]!;
}
