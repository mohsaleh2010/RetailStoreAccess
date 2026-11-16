using RetailStore.Data.Services;
using RetailStore.Domain.Localization;
using RetailStore.Domain.Printing;
using RetailStore.PosApp.Connection;
using RetailStore.PosApp.Printing;

namespace RetailStore.PosApp.Offline;

/// <summary>The POS on the PC (docs/83): the connection to the data chosen by an administrator, the login (on the data,
/// else offline for a user who logged in here before), the snapshot kept for selling offline, the sales of the PC and
/// their sync, and printing. One instance for the life of the program; the screens listen to Changed.</summary>
public sealed class PosSession(PosStore store, IPosPlatform platform, Func<PosConnection, string?, IPosBackend>? backendFactory = null)
{
    public const string PasswordKey = "pos_db_password";
    public const string TokenPrefix = "pos_token_";
    public static readonly TimeSpan OfflineLoginDays = TimeSpan.FromDays(30);
    public static readonly TimeSpan SnapshotAge = TimeSpan.FromMinutes(10);

    private readonly Func<PosConnection, string?, IPosBackend> _factory = backendFactory ?? DefaultBackend;
    private IPosBackend? _backend;
    private bool _loaded;

    public event Action? Changed;

    public Func<DateTime> Clock { get; set; } = () => DateTime.Now;
    public PosConnection? Connection { get; private set; }
    public PosSnapshot? Snapshot { get; private set; }
    public DateTime? SnapshotAt { get; private set; }
    public CachedUser? User { get; private set; }
    public bool LoggedIn => User is not null;
    public bool WorkingOffline { get; private set; }
    public bool Syncing { get; private set; }
    public PosSyncReport? LastSync { get; private set; }
    public bool English { get; private set; }
    public string Printer => store.Get("Printer") ?? "";
    public int PaperMillimetres => int.TryParse(store.Get("Paper"), out var p) ? p : 80;
    public bool OpenDrawer => store.Get("OpenDrawer") != "0";
    public IPosPlatform Platform => platform;
    public string Dir => English ? "ltr" : "rtl";

    /// <summary>Tr: the Arabic text, or its English (the dictionary of the English interface, docs/38).</summary>
    public string T(string? arabic) => English ? Translator.ToEnglish(arabic) : arabic ?? "";

    private void Notify() => Changed?.Invoke();

    public static IPosBackend DefaultBackend(PosConnection c, string? password) =>
        c.IsLocal ? new DirectBackend(c, password) : new ApiBackend(new HttpClient { Timeout = TimeSpan.FromSeconds(20) }, c.ServerUrl);

    /// <summary>The settings of the database may be changed by an administrator only, or by anyone while none is set (the
    /// first run of the program on the PC).</summary>
    public bool CanEditDatabase => Connection is null || User?.IsAdmin == true;

    public async Task LoadAsync()
    {
        if (_loaded)
        {
            return;
        }
        _loaded = true;
        English = store.Get("Language") == "en";
        Connection = store.LoadConnection();
        if (Connection is not null)
        {
            _backend = _factory(Connection, await platform.GetSecretAsync(PasswordKey));
        }
        Notify();
    }

    public void SetLanguage(bool english)
    {
        English = english;
        store.Set("Language", english ? "en" : "ar");
        Notify();
    }

    public void SetPrinter(string printer, int paperMillimetres, bool openDrawer)
    {
        store.Set("Printer", printer);
        store.Set("Paper", paperMillimetres.ToString());
        store.Set("OpenDrawer", openDrawer ? "1" : "0");
        Notify();
    }

    // ------------------------------------------------------------------ the database settings (administrator)

    /// <summary>Try a connection before it is saved: "" or why it does not answer.</summary>
    public async Task<string> TestConnectionAsync(PosConnection connection, string? password)
    {
        string problem = connection.Problem(password);
        if (problem.Length > 0)
        {
            return T(problem);
        }
        using var backend = _factory(connection, password);
        string r = await backend.TestAsync();
        return r.Length == 0 ? "" : T("تعذر الاتصال بقاعدة البيانات:") + " " + T(r);
    }

    /// <summary>Save the connection (an administrator, or the first run): it must answer; another database is refused
    /// while sales of this PC wait to be sent to the old one, and then what was kept of the old one is removed and the
    /// user logs in again.</summary>
    public async Task<string> SaveConnectionAsync(PosConnection connection, string? password)
    {
        if (!CanEditDatabase)
        {
            return T("إعدادات قاعدة البيانات لمدير النظام فقط.");
        }
        string test = await TestConnectionAsync(connection, password);
        if (test.Length > 0)
        {
            return test;
        }
        bool otherData = Connection is not null && (Connection.Mode != connection.Mode || Connection.Describe() != connection.Describe());
        if (otherData && store.Sales(SaleStates.Pending, SaleStates.Rejected).Count > 0)
        {
            return T("على الجهاز فواتير لم تصل إلى قاعدة البيانات الحالية. أرسلها أولًا، ثم غيّر قاعدة البيانات.");
        }
        var saved = connection with
        {
            ServerUrl = PosConnection.NormalizeUrl(connection.ServerUrl), DeviceCode = connection.DeviceCode.Trim().ToUpperInvariant(),
        };
        store.SaveConnection(saved);
        await platform.SetSecretAsync(PasswordKey, saved.IsLocal && !saved.WindowsLogin && !saved.IsSqlite ? password : null);
        if (otherData)
        {
            foreach (var u in store.Users())
            {
                await platform.SetSecretAsync(TokenPrefix + u.EmployeeID, null);
            }
            store.ClearData();
            Logout();
        }
        _backend?.Dispose();
        Connection = saved;
        _backend = _factory(saved, password);
        Notify();
        return "";
    }

    // ------------------------------------------------------------------ login

    /// <summary>Log in: on the data (his password checked by the program, then kept here as a PBKDF2 digest), else, when the
    /// data does not answer, offline for a user who logged in on this PC in the last 30 days. "" or the message.</summary>
    public async Task<string> LoginAsync(string username, string password)
    {
        string name = (username ?? "").Trim();
        if (Connection is null || _backend is null)
        {
            return T("اضبط إعدادات قاعدة البيانات أولًا.");
        }
        if (name.Length == 0)
        {
            return T("اكتب اسم المستخدم.");
        }
        var r = await _backend.LoginAsync(name, password ?? "", platform.DeviceName);
        if (r.Ok && r.Value is { } login)
        {
            if (login.MustChangePassword)
            {
                return T("يجب تغيير كلمة المرور أولًا من البرنامج أو الموقع، ثم الدخول هنا.");
            }
            await platform.SetSecretAsync(TokenPrefix + login.EmployeeID, login.Token);
            var who = new PosCredential(login.EmployeeID, login.Token);
            var snap = await _backend.SnapshotAsync(who, English);
            if (!snap.Ok || snap.Value is null)
            {
                return snap.Retry ? T("تعذر الاتصال بقاعدة البيانات.") + " " + snap.Message : snap.Message;
            }
            var at = Clock();
            store.RememberUser(login.EmployeeID, name, login.EmployeeName, snap.Value.User.IsAdmin, password ?? "", at);
            store.SaveSnapshot(login.EmployeeID, snap.Value, at);
            Snapshot = snap.Value;
            SnapshotAt = at;
            WorkingOffline = false;
            User = new CachedUser(login.EmployeeID, name, login.EmployeeName, snap.Value.User.IsAdmin, at);
            Notify();
            return "";
        }
        if (!r.Retry)
        {
            store.ForgetUser(name);
            return r.Message;
        }
        var cached = store.CheckOffline(name, password ?? "");
        if (cached is null)
        {
            return T("تعذر الاتصال بقاعدة البيانات، ولم يدخل هذا المستخدم بكلمة المرور هذه على هذا الجهاز من قبل.");
        }
        if (Clock() - cached.LastOnline > OfflineLoginDays)
        {
            return T("تعذر الاتصال بقاعدة البيانات، وآخر دخول لهذا المستخدم على البرنامج منذ أكثر من 30 يومًا.");
        }
        var (snapshot, savedAt) = store.LoadSnapshot(cached.EmployeeID);
        if (snapshot is null)
        {
            return T("تعذر الاتصال بقاعدة البيانات، ولا توجد بيانات محفوظة لهذا المستخدم على الجهاز.");
        }
        Snapshot = snapshot;
        SnapshotAt = savedAt;
        WorkingOffline = true;
        User = cached;
        Notify();
        return "";
    }

    /// <summary>Log out: the sales and the tokens stay, so the sales of each cashier are sent while another one works.</summary>
    public void Logout()
    {
        User = null;
        Snapshot = null;
        SnapshotAt = null;
        Notify();
    }

    // ------------------------------------------------------------------ sync

    public int Waiting => store.Sales(SaleStates.Pending).Count;

    public int Refused => store.Sales(SaleStates.Rejected).Count;

    private async Task<PosCredential?> CredentialOfAsync(int employeeId)
    {
        if (Connection?.IsLocal == true)
        {
            return new PosCredential(employeeId, null);
        }
        string? token = await platform.GetSecretAsync(TokenPrefix + employeeId);
        return token is { Length: > 0 } ? new PosCredential(employeeId, token) : null;
    }

    /// <summary>Send the sales of the PC, then read the snapshot of the user again when it is older than ten minutes (or
    /// always with refresh).</summary>
    public async Task<PosSyncReport> SyncAsync(bool refresh = false)
    {
        if (_backend is null || Syncing)
        {
            return LastSync ?? new PosSyncReport(false, 0, 0, Waiting, "", null);
        }
        Syncing = true;
        Notify();
        try
        {
            var (sent, rejected, status, message) = await PosSync.SendSalesAsync(
                store, _backend, CredentialOfAsync, id => platform.SetSecretAsync(TokenPrefix + id, null), English, Clock);
            if (status == 0 || status >= 500 || status == 429)
            {
                WorkingOffline = true;
                return LastSync = new PosSyncReport(status != 0, sent, rejected, Waiting,
                                                    T("تعذر الاتصال بقاعدة البيانات. ستُرسل الفواتير عند عودة الاتصال."), null);
            }
            DateTime? snapshotAt = null;
            string note = Waiting > 0 ? T("فواتير بانتظار دخول الكاشير الذي باعها.") : "";
            if (User is not null && (refresh || SnapshotAt is null || Clock() - SnapshotAt > SnapshotAge))
            {
                var who = await CredentialOfAsync(User.EmployeeID);
                var snap = who is null ? null : await _backend.SnapshotAsync(who, English);
                if (snap is { Ok: true, Value: { } s })
                {
                    var at = Clock();
                    store.SaveSnapshot(User.EmployeeID, s, at);
                    store.UpdateUserFacts(User.EmployeeID, s.User.EmployeeName, s.User.IsAdmin);
                    Snapshot = s;
                    SnapshotAt = snapshotAt = at;
                    User = User with { IsAdmin = s.User.IsAdmin, EmployeeName = s.User.EmployeeName };
                }
                else if (snap is null || snap.Status == 401)
                {
                    note = T("انتهى الدخول على الخادم: اخرج وادخل من جديد لتحديث الأسعار والأرصدة.");
                }
                else if (!snap.Retry)
                {
                    note = snap.Message;
                }
            }
            WorkingOffline = false;
            store.ForgetSent(Clock().AddDays(-7));
            return LastSync = new PosSyncReport(true, sent, rejected, Waiting, note, snapshotAt);
        }
        finally
        {
            Syncing = false;
            Notify();
        }
    }

    // ------------------------------------------------------------------ selling

    public PosCart? NewCart() => Snapshot is null ? null : new PosCart(Snapshot);

    /// <summary>What the snapshot does not show yet: the sales of the PC waiting or refused, and those sent after the
    /// snapshot was read (the pieces of each product and the credit of each customer), so the next sale sees them.</summary>
    public PendingEffects Pending()
    {
        var pieces = new Dictionary<int, decimal>();
        var credit = new Dictionary<int, decimal>();
        // the times of the file are kept to the second: a sale sent in the second of the snapshot is counted (at worst twice)
        DateTime? since = SnapshotAt is { } at ? at.AddTicks(-(at.Ticks % TimeSpan.TicksPerSecond)) : null;
        foreach (var s in store.Sales().Where(s => s.State != SaleStates.Sent || since is null || s.SentAt >= since))
        {
            for (int i = 0; i < s.Request.Lines.Length; i++)
            {
                var l = s.Request.Lines[i];
                var p = Snapshot?.Products.FirstOrDefault(x => x.ProductID == l.ProductID);
                decimal factor = p?.Units.FirstOrDefault(u => u.UnitID == l.UnitID)?.Factor ?? 1;
                pieces[l.ProductID] = pieces.GetValueOrDefault(l.ProductID) + s.Receipt.Lines[i].Quantity * factor;
            }
            if (s.Receipt.Remaining != 0)
            {
                credit[s.Request.CustomerID] = credit.GetValueOrDefault(s.Request.CustomerID) + s.Receipt.Remaining;
            }
        }
        return new PendingEffects(pieces, credit);
    }

    /// <summary>Keep the sale on the PC with the next number of the PC, then it is sent: Saved and the sale, or the question
    /// (NeedsConfirmation) or the refusal with its message.</summary>
    public (SaveStatus Status, string Message, StoredSale? Sale) SaveSale(PosCart cart, bool confirmStock = false)
    {
        if (User is null || Connection is null)
        {
            return (SaveStatus.Refused, T("سجّل الدخول أولًا."), null);
        }
        var pending = Pending();
        var (status, message, _, _) = cart.Check(pending);
        if (status == SaveStatus.Refused || (status == SaveStatus.NeedsConfirmation && !confirmStock))
        {
            return (status, T(message), null);
        }
        var at = Clock();
        var (_, _, receipt) = cart.ToReceipt(store.NextInvoiceNumber(Connection.DeviceCode), at, confirmStock, pending);
        long id = store.AddSale(receipt!, User.EmployeeID, User.EmployeeName, at);
        Notify();
        return (SaveStatus.Saved, "", store.Sales().First(s => s.Id == id));
    }

    /// <summary>The sales of the PC, newest first.</summary>
    public List<StoredSale> Sales() => store.Sales().OrderByDescending(s => s.Id).ToList();

    /// <summary>Send a refused sale again; with the answer yes to the stock question: an administrator only.</summary>
    public string Retry(long id, bool confirmStock)
    {
        if (confirmStock && User?.IsAdmin != true)
        {
            return T("هذا الإجراء لمدير النظام فقط.");
        }
        store.Retry(id, confirmStock);
        Notify();
        return "";
    }

    /// <summary>Remove a refused sale from the PC (the office records it itself): an administrator only.</summary>
    public string Delete(long id)
    {
        if (User?.IsAdmin != true)
        {
            return T("هذا الإجراء لمدير النظام فقط.");
        }
        store.Delete(id);
        Notify();
        return "";
    }

    // ------------------------------------------------------------------ printing

    /// <summary>Print the receipt on the chosen printer (the drawer opens for a new cash sale): "" or the reason.</summary>
    public async Task<string> PrintAsync(StoredSale sale, bool reprint = false)
    {
        if (Printer.Length == 0)
        {
            return T("اختر الطابعة في الإعدادات.");
        }
        int width = EscPos.DotsFor(PaperMillimetres);
        var text = platform.DrawSlip(ReceiptSlip.Build(sale.Receipt, sale.State == SaleStates.Sent, T), width, !English);
        var qr = sale.Receipt.Qr is { Length: > 0 } q ? ReceiptSlip.Qr(q, PaperMillimetres >= 80 ? 5 : 4) : null;
        bool drawer = !reprint && OpenDrawer && sale.Request.PaymentType == "CASH" && sale.Request.PaymentMethodID == SalesService.CashMethodId;
        string problem = await platform.PrintAsync(Printer, ReceiptSlip.Job(ReceiptSlip.Stack(text, qr), drawer));
        return problem.Length == 0 ? "" : T("تعذرت الطباعة:") + " " + problem;
    }
}
