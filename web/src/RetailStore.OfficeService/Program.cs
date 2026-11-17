using Microsoft.EntityFrameworkCore;
using RetailStore.Data;
using RetailStore.Data.Journal;
using RetailStore.Data.Services;
using RetailStore.OfficeService;

// The office tasks as a Windows service (plan 7A, docs/85): the journal and the e-invoices of the documents the web, the Windows POS,
// the rep app and Access make, with no Access open. It works on the same database as the web (appsettings.json: Database:Provider
// SqlServer | Sqlite and the connection string) and runs when Settings.OfficeSyncComputer names it (OfficeService:Name, default
// "SERVICE-" + the computer name). Install: sc.exe create (docs/85).
var builder = Host.CreateApplicationBuilder(args);
builder.Services.AddWindowsService(options => options.ServiceName = "RetailStore Office Service");

string provider = builder.Configuration["Database:Provider"] ?? "SqlServer";
builder.Services.AddDbContextFactory<RetailDbContext>(options =>
{
    if (provider.Equals("Sqlite", StringComparison.OrdinalIgnoreCase))
    {
        options.UseSqlite(builder.Configuration.GetConnectionString("Sqlite") is { Length: > 0 } cs ? cs
                          : throw new InvalidOperationException("ConnectionStrings:Sqlite is missing"));
    }
    else
    {
        options.UseSqlServer(builder.Configuration.GetConnectionString("SqlServer") is { Length: > 0 } cs ? cs
                             : throw new InvalidOperationException("ConnectionStrings:SqlServer is missing"));
    }
});
builder.Services.AddSingleton<IJournalSourceReader, ServerViewJournalReader>();
builder.Services.AddSingleton<JournalService>();
builder.Services.AddSingleton(new HttpClient(new SocketsHttpHandler { PooledConnectionLifetime = TimeSpan.FromMinutes(10) }));
builder.Services.AddSingleton<IEtaSigner, ProcessEtaSigner>();
builder.Services.AddSingleton(sp => new EInvoiceService(sp.GetRequiredService<IDbContextFactory<RetailDbContext>>(),
                                                        sp.GetRequiredService<HttpClient>(), sp.GetRequiredService<IEtaSigner>()));
builder.Services.AddSingleton(sp => new OfficeSyncService(sp.GetRequiredService<IDbContextFactory<RetailDbContext>>(),
                                                          sp.GetRequiredService<JournalService>(),
                                                          sp.GetRequiredService<EInvoiceService>()));
builder.Services.AddHostedService<OfficeWorker>();

var host = builder.Build();
if (provider.Equals("Sqlite", StringComparison.OrdinalIgnoreCase))
{
    using var db = host.Services.GetRequiredService<IDbContextFactory<RetailDbContext>>().CreateDbContext();
    DbSetup.EnsureSqliteDatabase(db);
}
// By hand, before installing (docs/85): --check = is the platform set up and does the key sign here; --run-now = one run, printed.
if (args.Contains("--check") || args.Contains("--run-now"))
{
    var office = host.Services.GetRequiredService<OfficeSyncService>();
    var einvoices = host.Services.GetRequiredService<EInvoiceService>();
    string name = OfficeSyncService.ServiceName(host.Services.GetRequiredService<IConfiguration>()["OfficeService:Name"], Environment.MachineName);
    Console.OutputEncoding = System.Text.Encoding.UTF8;
    Console.WriteLine("Service name (Settings.OfficeSyncComputer): " + name);
    if (args.Contains("--check"))
    {
        string problem = await einvoices.ReadyProblemAsync();
        Console.WriteLine(problem.Length == 0 ? "E-invoicing: ready." : "E-invoicing: " + problem);
        Console.WriteLine("Signing with the device key: " + (await einvoices.CheckSigningAsync() is { Length: > 0 } p ? p : "OK"));
        Console.WriteLine("Pending e-invoices: " + await office.PendingCountAsync());
        Console.WriteLine("Due now: " + (await office.IsDueAsync(name) ? "yes" : "no"));
    }
    if (args.Contains("--run-now"))
    {
        var run = await office.RunAsync("MANUAL", name, new ActingUser(0, name));
        Console.WriteLine(run.Summary);
        Console.WriteLine(run.Problem.Length == 0 ? "No problem." : "Problem: " + run.Problem);
    }
    return;
}
host.Run();
