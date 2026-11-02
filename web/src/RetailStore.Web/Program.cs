using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.EntityFrameworkCore;
using RetailStore.Data;
using RetailStore.UI;
using RetailStore.Web;
using RetailStore.Web.Components;

var builder = WebApplication.CreateBuilder(args);

// The data: the SQL Server database of plan 5 (the same one Access links to), or a SQLite file for development
// (created with the seed rows on first run). appsettings.json: Database:Provider = SqlServer | Sqlite.
string provider = builder.Configuration["Database:Provider"] ?? "Sqlite";
builder.Services.AddDbContextFactory<RetailDbContext>(options =>
{
    if (provider.Equals("SqlServer", StringComparison.OrdinalIgnoreCase))
    {
        options.UseSqlServer(builder.Configuration.GetConnectionString("SqlServer")
                             ?? throw new InvalidOperationException("ConnectionStrings:SqlServer is missing"));
    }
    else
    {
        options.UseSqlite(builder.Configuration.GetConnectionString("Sqlite") ?? "Data Source=App_Data/retailstore-dev.db");
    }
});

builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options =>
    {
        options.LoginPath = "/account/login";
        options.LogoutPath = "/account/logout";
        options.AccessDeniedPath = "/account/login";
        options.Cookie.Name = "RetailStore.Auth";
        options.Cookie.HttpOnly = true;
        options.Cookie.SameSite = SameSiteMode.Strict;
        options.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest;
        options.ExpireTimeSpan = TimeSpan.FromHours(10);
        options.SlidingExpiration = true;
    });
builder.Services.AddAuthorization();
builder.Services.AddCascadingAuthenticationState();
builder.Services.AddRetailStoreUi();

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

var app = builder.Build();

if (!provider.Equals("SqlServer", StringComparison.OrdinalIgnoreCase))
{
    Directory.CreateDirectory(Path.Combine(app.Environment.ContentRootPath, "App_Data"));
    using var scope = app.Services.CreateScope();
    using var db = scope.ServiceProvider.GetRequiredService<IDbContextFactory<RetailDbContext>>().CreateDbContext();
    DbSetup.EnsureSqliteDatabase(db);
}

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}
app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseHttpsRedirection();

app.UseAuthentication();
app.UseAuthorization();
app.UseAntiforgery();

app.MapStaticAssets();
app.MapAccountEndpoints();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode()
    .AddAdditionalAssemblies(typeof(RetailStore.UI.Layout.MainLayout).Assembly);

app.Run();

/// <summary>The entry point, visible to the tests.</summary>
public partial class Program;
