using Microsoft.Data.SqlClient;
using Microsoft.Data.Sqlite;
using Microsoft.EntityFrameworkCore;
using RetailStore.Data;
using RetailStore.Data.Services;
using RetailStore.Domain.Localization;

namespace RetailStore.PosApp.Connection;

/// <summary>A local database (docs/83): SQL Server of the store or a SQLite trial file, used directly with the services of
/// the web (SecurityService, PosService, SalesService): the same checks, rows and audit trail as Access.</summary>
public sealed class DirectBackend : IPosBackend, IDbContextFactory<RetailDbContext>
{
    private readonly DbContextOptions<RetailDbContext> _options;
    private readonly bool _sqlite;
    private bool _created;

    public DirectBackend(PosConnection connection, string? password)
    {
        _sqlite = connection.IsSqlite;
        var b = new DbContextOptionsBuilder<RetailDbContext>();
        string cs = connection.ConnectionString(password);
        if (_sqlite)
        {
            b.UseSqlite(cs);
        }
        else
        {
            b.UseSqlServer(cs);
        }
        _options = b.Options;
    }

    /// <summary>For the tests: a database given by its factory.</summary>
    public DirectBackend(IDbContextFactory<RetailDbContext> factory)
    {
        Factory = factory;
        _options = new DbContextOptionsBuilder<RetailDbContext>().Options;
        _created = true;
    }

    private IDbContextFactory<RetailDbContext>? Factory { get; }

    public RetailDbContext CreateDbContext() => Factory?.CreateDbContext() ?? new RetailDbContext(_options);

    private PermissionService Permissions => new(this);

    private SalesService Sales => new(this, Permissions);

    private PosService Pos => new(this, Permissions, Sales);

    private TreasuryService Treasury => new(this, Permissions);

    /// <summary>A SQLite trial file is created with the seed rows the first time (DbSetup); a SQL Server database never
    /// is: it comes from the script of docs/67.</summary>
    private void EnsureCreated()
    {
        if (_created || !_sqlite)
        {
            return;
        }
        using var db = CreateDbContext();
        DbSetup.EnsureSqliteDatabase(db);
        _created = true;
    }

    /// <summary>The database not reached = 0 (the work waits for it); any other failure = 500 with its message, so a sale
    /// is never lost nor marked refused by a fault of the program.</summary>
    private async Task<BackendResult<T>> RunAsync<T>(Func<Task<BackendResult<T>>> work)
    {
        try
        {
            EnsureCreated();
            await using (var db = CreateDbContext())
            {
                if (!await db.Database.CanConnectAsync())
                {
                    return BackendResult<T>.Unreachable("تعذر الاتصال بقاعدة البيانات.");
                }
            }
        }
        catch (Exception e) when (e is SqlException or SqliteException or InvalidOperationException or TimeoutException or IOException)
        {
            return BackendResult<T>.Unreachable(e.GetBaseException().Message);
        }
        try
        {
            return await work();
        }
        catch (Exception e) when (e is SqlException or SqliteException or DbUpdateException or InvalidOperationException or TimeoutException)
        {
            return new BackendResult<T>(default, 500, "error", e.GetBaseException().Message);
        }
    }

    public async Task<string> TestAsync(CancellationToken ct = default)
    {
        var r = await RunAsync(async () =>
        {
            await using var db = CreateDbContext();
            bool ok = await db.Settings.AnyAsync(ct);
            return ok ? BackendResult<bool>.Of(true) : BackendResult<bool>.Refused("قاعدة البيانات لا تحتوي إعدادات البرنامج.");
        });
        return r.Ok ? "" : r.Message;
    }

    public Task<BackendResult<BackendLogin>> LoginAsync(string username, string password, string device, CancellationToken ct = default) =>
        RunAsync(async () =>
        {
            var r = await new SecurityService(this).LoginAsync(username, password, "POS " + device, channel: "pos");
            return r.EmployeeID == 0
                ? BackendResult<BackendLogin>.Refused(r.Message, 401, "login_failed")
                : BackendResult<BackendLogin>.Of(new BackendLogin(r.EmployeeID, r.EmployeeName, r.MustChangePassword, null));
        });

    public Task<BackendResult<PosSnapshot>> SnapshotAsync(PosCredential who, bool english, CancellationToken ct = default) =>
        RunAsync(async () =>
        {
            string open = await Permissions.OpenProblemAsync(who.EmployeeID, "frmPOS");
            if (open.Length > 0)
            {
                return BackendResult<PosSnapshot>.Refused(Tr(open, english), 403, "forbidden");
            }
            var s = await Pos.SnapshotAsync(who.EmployeeID, english);
            return s is null ? BackendResult<PosSnapshot>.Refused(Tr("سجّل الدخول أولًا.", english), 401, "unauthorized")
                             : BackendResult<PosSnapshot>.Of(s);
        });

    public Task<BackendResult<DeviceSaleResult>> SendSaleAsync(PosCredential who, DeviceSaleRequest sale, bool english,
                                                               CancellationToken ct = default) =>
        RunAsync(async () =>
        {
            var r = await Pos.PostDeviceSaleAsync(new ActingUser(who.EmployeeID, "POS"), sale);
            return r.Status switch
            {
                SaveStatus.Saved => new BackendResult<DeviceSaleResult>(r, r.Duplicate ? 200 : 201, "", ""),
                SaveStatus.NeedsConfirmation => BackendResult<DeviceSaleResult>.Refused(Tr(r.Message, english), 409, "confirm"),
                _ => BackendResult<DeviceSaleResult>.Refused(Tr(r.Message, english)),
            };
        });

    /// <summary>The screen of an operation: refused (403) as the API's PermissionFilter would refuse it.</summary>
    private async Task<string> ScreenProblemAsync(int employeeId, string form, ScreenAction? action = null)
    {
        string open = await Permissions.OpenProblemAsync(employeeId, form);
        return open.Length > 0 || action is null ? open : await Permissions.ActionProblemAsync(employeeId, form, action.Value);
    }

    private static BackendResult<PosDocument> Answer(DeviceDocResult r, bool english) => r.Status switch
    {
        SaveStatus.Saved => new BackendResult<PosDocument>(r.Document, r.Duplicate ? 200 : 201, "", ""),
        SaveStatus.NeedsConfirmation => BackendResult<PosDocument>.Refused(Tr(r.Message, english), 409, "confirm"),
        _ => BackendResult<PosDocument>.Refused(Tr(r.Message, english)),
    };

    public Task<BackendResult<ReturnableInvoice>> ReturnableAsync(PosCredential who, string invoiceNumber, bool english,
                                                                 CancellationToken ct = default) =>
        RunAsync(async () =>
        {
            string problem = await ScreenProblemAsync(who.EmployeeID, "frmSalesReturn");
            if (problem.Length > 0)
            {
                return BackendResult<ReturnableInvoice>.Refused(Tr(problem, english), 403, "forbidden");
            }
            var inv = await Sales.ReturnableAsync(invoiceNumber, english);
            return inv is null ? BackendResult<ReturnableInvoice>.Refused(Tr("لا توجد فاتورة بهذا الرقم.", english), 404, "not_found")
                               : BackendResult<ReturnableInvoice>.Of(inv);
        });

    public Task<BackendResult<PosDocument>> SendReturnAsync(PosCredential who, DeviceReturnRequest request, bool english,
                                                           CancellationToken ct = default) =>
        RunAsync(async () => Answer(await Pos.PostDeviceReturnAsync(new ActingUser(who.EmployeeID, "POS"), request, english), english));

    public Task<BackendResult<PosDocument>> SendPaymentAsync(PosCredential who, DevicePaymentRequest request, bool english,
                                                            CancellationToken ct = default) =>
        RunAsync(async () => Answer(await Pos.PostDevicePaymentAsync(new ActingUser(who.EmployeeID, "POS"), request, english), english));

    public Task<BackendResult<PosClosingBoxes>> ClosingBoxesAsync(PosCredential who, bool english, CancellationToken ct = default) =>
        RunAsync(async () =>
        {
            string problem = await ScreenProblemAsync(who.EmployeeID, TreasuryService.ClosingForm);
            return problem.Length > 0 ? BackendResult<PosClosingBoxes>.Refused(Tr(problem, english), 403, "forbidden")
                                      : BackendResult<PosClosingBoxes>.Of(await Pos.ClosingBoxesAsync(who.EmployeeID, english));
        });

    public Task<BackendResult<ClosingPreview>> ClosingPreviewAsync(PosCredential who, int boxId, bool english, CancellationToken ct = default) =>
        RunAsync(async () =>
        {
            string problem = await ScreenProblemAsync(who.EmployeeID, TreasuryService.ClosingForm);
            if (problem.Length > 0)
            {
                return BackendResult<ClosingPreview>.Refused(Tr(problem, english), 403, "forbidden");
            }
            var p = await Treasury.ClosingPreviewAsync(boxId, english);
            return p is null ? BackendResult<ClosingPreview>.Refused(Tr("السجل غير موجود.", english), 404, "not_found")
                             : BackendResult<ClosingPreview>.Of(p);
        });

    public Task<BackendResult<PosDocument>> SendClosingAsync(PosCredential who, DeviceClosingRequest request, bool english,
                                                            CancellationToken ct = default) =>
        RunAsync(async () => Answer(await Pos.PostDeviceClosingAsync(new ActingUser(who.EmployeeID, "POS"), request, english), english));

    private static string Tr(string arabic, bool english) => english ? Translator.ToEnglish(arabic) : arabic;

    public void Dispose()
    {
        if (_sqlite)
        {
            SqliteConnection.ClearAllPools();
        }
    }
}
