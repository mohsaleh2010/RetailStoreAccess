using Microsoft.EntityFrameworkCore;
using RetailStore.Data;

namespace RetailStore.Web;

/// <summary>GET /health (docs/88): the check of a new installation. It answers 200 {"status":"ok"} when the site runs
/// and reads its database (the Settings row), else 503 {"status":"database"}. It says nothing else (no server name,
/// no error text): the reason is in the log of the site.</summary>
public static class Health
{
    public static void MapHealth(this IEndpointRouteBuilder app)
    {
        app.MapGet("/health", async (IDbContextFactory<RetailDbContext> factory, ILoggerFactory logs, CancellationToken ct) =>
        {
            try
            {
                await using var db = await factory.CreateDbContextAsync(ct);
                bool settings = await db.Settings.AsNoTracking().AnyAsync(ct);
                return settings
                    ? Results.Json(new { status = "ok" })
                    : Results.Json(new { status = "database" }, statusCode: StatusCodes.Status503ServiceUnavailable);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                logs.CreateLogger("RetailStore.Health").LogError(ex, "Health check: the database does not answer");
                return Results.Json(new { status = "database" }, statusCode: StatusCodes.Status503ServiceUnavailable);
            }
        }).AllowAnonymous().ExcludeFromDescription();
    }
}
