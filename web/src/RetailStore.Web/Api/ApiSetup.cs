using System.Text.Json.Nodes;
using System.Threading.RateLimiting;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.OpenApi;
using RetailStore.Data.Services;
using Scalar.AspNetCore;

namespace RetailStore.Web.Api;

/// <summary>The services and the pipeline of the API (docs/77): JWT bearer tokens next to the cookie of the screens, the
/// rate limit of the login, problem details, and the OpenAPI document (/api/openapi/v1.json) with its reference page
/// (/api/docs), where every operation shows the permission it needs.</summary>
public static class ApiSetup
{
    public const string TokenLimit = "api-token";
    public const string DocumentPath = "/api/openapi/{documentName}.json";
    public const string DocsPath = "/api/docs";

    public static bool IsApi(HttpContext http) => http.Request.Path.StartsWithSegments("/api");

    public static void AddRetailApi(this WebApplicationBuilder builder)
    {
        builder.Services.AddSingleton<ApiTokens>();
        builder.Services.AddScoped<RepService>();
        builder.Services.AddProblemDetails();
        builder.Services.AddAuthentication()
            .AddJwtBearer(ApiTokens.Scheme, _ => { });
        builder.Services.AddOptions<JwtBearerOptions>(ApiTokens.Scheme)
            .Configure<ApiTokens>((options, tokens) =>
            {
                options.MapInboundClaims = false;
                options.TokenValidationParameters = tokens.Validation();
                options.Events = new JwtBearerEvents
                {
                    OnChallenge = async context =>
                    {
                        context.HandleResponse();
                        await ApiResults.Problem(context.HttpContext, StatusCodes.Status401Unauthorized, "unauthorized",
                                                 "سجّل الدخول أولًا.").ExecuteAsync(context.HttpContext);
                    },
                    OnForbidden = async context =>
                        await ApiResults.Problem(context.HttpContext, StatusCodes.Status403Forbidden, "forbidden",
                                                 "لا تملك صلاحية هذه العملية. راجع مدير النظام.").ExecuteAsync(context.HttpContext),
                };
            });
        builder.Services.AddRateLimiter(options =>
        {
            options.RejectionStatusCode = StatusCodes.Status429TooManyRequests;
            options.AddPolicy(TokenLimit, http => RateLimitPartition.GetFixedWindowLimiter(
                http.Connection.RemoteIpAddress?.ToString() ?? "",
                _ => new FixedWindowRateLimiterOptions
                {
                    PermitLimit = int.TryParse(builder.Configuration["Api:LoginsPerMinute"], out var n) && n > 0 ? n : 10,
                    Window = TimeSpan.FromMinutes(1), QueueLimit = 0,
                }));
        });
        builder.Services.AddOpenApi("v1", options =>
        {
            options.ShouldInclude = d => d.RelativePath?.StartsWith("api/", StringComparison.Ordinal) == true;
            options.AddDocumentTransformer((document, _, _) =>
            {
                document.Info = new OpenApiInfo
                {
                    Title = "RetailStore API", Version = "v1",
                    Description = "The REST API of the retail store system (docs/77-Web-API.md): the same users, permissions, " +
                                  "checks, messages and audit trail as the Access program, on the same database. Log in with " +
                                  "POST /api/v1/auth/token and send the header Authorization: Bearer followed by the token. Errors are problem details " +
                                  "(title = the message, code, field). Messages are Arabic; send Accept-Language: en for English.",
                };
                document.Components ??= new OpenApiComponents();
                document.Components.SecuritySchemes ??= new Dictionary<string, IOpenApiSecurityScheme>();
                document.Components.SecuritySchemes[ApiTokens.Scheme] = new OpenApiSecurityScheme
                {
                    Type = SecuritySchemeType.Http, Scheme = "bearer", BearerFormat = "JWT",
                    Description = "The token of POST /api/v1/auth/token",
                };
                return Task.CompletedTask;
            });
            options.AddOperationTransformer((operation, context, _) =>
            {
                var metadata = context.Description.ActionDescriptor.EndpointMetadata;
                var permissions = metadata.OfType<ApiPermission>().Select(p => p.Text).ToList();
                if (permissions.Count > 0)
                {
                    operation.Description = ((operation.Description ?? "") + " Permission: " + string.Join(", ", permissions) + ".").Trim();
                    operation.Extensions ??= new Dictionary<string, IOpenApiExtension>();
                    operation.Extensions["x-permission"] = new JsonNodeExtension(JsonValue.Create(string.Join(", ", permissions)));
                }
                if (metadata.OfType<AuthorizeAttribute>().Any() && !metadata.OfType<AllowAnonymousAttribute>().Any())
                {
                    operation.Security ??= [];
                    operation.Security.Add(new OpenApiSecurityRequirement
                    {
                        [new OpenApiSecuritySchemeReference(ApiTokens.Scheme, context.Document)] = [],
                    });
                }
                return Task.CompletedTask;
            });
        });
    }

    public static void UseRetailApi(this WebApplication app)
    {
        app.UseRateLimiter();
    }

    public static void MapRetailApiDocs(this WebApplication app)
    {
        app.MapRetailApi();
        if (!string.Equals(app.Configuration["Api:Docs"], "false", StringComparison.OrdinalIgnoreCase))
        {
            app.MapOpenApi(DocumentPath).AllowAnonymous();
            app.MapScalarApiReference(DocsPath, options => options
                .WithTitle("RetailStore API")
                .WithOpenApiRoutePattern(DocumentPath)
                .AddPreferredSecuritySchemes(ApiTokens.Scheme)
                .DisableAgent().DisableMcp().DisableTelemetry().DisableDefaultFonts()   // nothing leaves the shop
                .HideDeveloperTools()).AllowAnonymous();
        }
    }
}
