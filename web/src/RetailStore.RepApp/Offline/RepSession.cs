using RetailStore.Domain.Localization;
using RetailStore.RepApp.Client;
using RetailStore.RepApp.Printing;

namespace RetailStore.RepApp.Offline;

/// <summary>An order in the list of the phone: taken on the phone (waiting, sent or refused) or read from the server.</summary>
public sealed record OrderRow(long OutboxId, string State, string? Message, PendingOrder Order, string? OrderNumber, string? Status);

/// <summary>The rep app on the phone (docs/78): the settings, the login, the data kept for offline work, the outbox and
/// the sync, and printing. One instance for the life of the app; the screens listen to Changed.</summary>
public sealed class RepSession(LocalStore store, RepApiClient api, IRepPlatform platform)
{
    public const string TokenKey = "rep_token";
    private readonly SyncService _sync = new(store, api);
    private bool _loaded;

    public event Action? Changed;

    public Func<DateTime> Clock { get; set; } = () => DateTime.Now;
    public RepSnapshot? Snapshot { get; private set; }
    public DateTime? SnapshotAt { get; private set; }
    public bool English { get; private set; }
    public bool LoggedIn { get; private set; }
    public bool MustChangePassword { get; private set; }
    public bool Syncing { get; private set; }
    public SyncReport? LastSync { get; private set; }
    public string ServerUrl => store.Get("ServerUrl") ?? "";
    public string Username => store.Get("Username") ?? "";
    public string EmployeeName => store.Get("EmployeeName") ?? "";
    public string Printer => store.Get("Printer") ?? "";
    public int PaperMillimetres => int.TryParse(store.Get("Paper"), out var p) ? p : 58;
    public IRepPlatform Platform => platform;

    public string Dir => English ? "ltr" : "rtl";

    /// <summary>Tr: the Arabic text, or its English (the dictionary of the English interface, docs/38).</summary>
    public string T(string? arabic) => English ? Translator.ToEnglish(arabic) : arabic ?? "";

    private void Notify() => Changed?.Invoke();

    /// <summary>RepVisits.Result as the screens show it (AUTO = recorded on the phone, read from the day's documents).</summary>
    public string VisitResultText(string result) => T(result switch
    {
        "ORDER" => "طلبية",
        "SALE" => "بيع",
        "PAYMENT" => "تحصيل",
        "NO_SALE" => "لا شيء",
        "CLOSED" => "المحل مغلق",
        "AUTO" => "من مستندات اليوم",
        _ => "لم تُسجَّل",
    });

    /// <summary>SalesOrders.Status (rule_text of the field in tools/schema.py).</summary>
    public string OrderStatusText(string? status) => T(status switch
    {
        "DRAFT" => "مسودة",
        "APPROVED" => "معتمدة",
        "PREPARING" => "قيد التجهيز",
        "PARTIAL" => "مُسلَّمة جزئيًا",
        "DELIVERED" => "مُسلَّمة",
        "INVOICED" => "مفوترة",
        "CANCELLED" => "ملغاة",
        "CLOSED" => "مغلقة",
        _ => status ?? "",
    });

    /// <summary>Read the settings, the token and the last snapshot when the app starts.</summary>
    public async Task LoadAsync()
    {
        if (_loaded)
        {
            return;
        }
        _loaded = true;
        English = store.Get("Language") == "en";
        api.English = English;
        api.ServerUrl = ServerUrl;
        api.Token = await platform.GetSecretAsync(TokenKey);
        LoggedIn = api.Token is { Length: > 0 };
        MustChangePassword = store.Get("MustChangePassword") == "1";
        (Snapshot, SnapshotAt) = store.LoadSnapshot();
        Notify();
    }

    public void SetLanguage(bool english)
    {
        English = english;
        api.English = english;
        store.Set("Language", english ? "en" : "ar");
        Notify();
    }

    public void SetPrinter(string printer, int paperMillimetres)
    {
        store.Set("Printer", printer);
        store.Set("Paper", paperMillimetres.ToString());
        Notify();
    }

    // ------------------------------------------------------------------ login

    /// <summary>Log in on the server: "" or the message. Another user may log in only when nothing of the last one waits
    /// to be sent; what the last one kept is then removed.</summary>
    public async Task<string> LoginAsync(string server, string username, string password)
    {
        string url = RepApiClient.NormalizeServer(server);
        string user = (username ?? "").Trim();
        if (url.Length == 0)
        {
            return T("اكتب عنوان الخادم.");
        }
        if (user.Length == 0)
        {
            return T("اكتب اسم المستخدم.");
        }
        bool otherUser = Username.Length > 0 && !string.Equals(Username, user, StringComparison.OrdinalIgnoreCase);
        if (otherUser && store.Outbox(OutboxKinds.Pending, OutboxKinds.Rejected).Count > 0)
        {
            return T("على الجهاز عمليات لم تُرسل للمستخدم") + " " + Username + ". " + T("ادخل باسمه وزامن أولًا.");
        }
        api.ServerUrl = url;
        api.Token = null;
        var r = await api.LoginAsync(user, password ?? "", platform.DeviceName);
        if (!r.Ok || r.Value is null)
        {
            return r.Status == 0 ? T("تعذر الوصول إلى الخادم. تأكد من الإنترنت ومن عنوان الخادم.") : r.Message;
        }
        if (otherUser)
        {
            store.ClearData();
            Snapshot = null;
            SnapshotAt = null;
        }
        store.Set("ServerUrl", url);
        store.Set("Username", user);
        store.Set("EmployeeName", r.Value.EmployeeName);
        await KeepTokenAsync(r.Value);
        Notify();
        return "";
    }

    private async Task KeepTokenAsync(TokenResponse t)
    {
        api.Token = t.AccessToken;
        await platform.SetSecretAsync(TokenKey, t.AccessToken);
        LoggedIn = true;
        MustChangePassword = t.MustChangePassword;
        store.Set("MustChangePassword", t.MustChangePassword ? "1" : "0");
    }

    public async Task<string> ChangePasswordAsync(string oldPassword, string newPassword, string confirmPassword)
    {
        var r = await api.ChangePasswordAsync(oldPassword ?? "", newPassword ?? "", confirmPassword ?? "");
        if (!r.Ok || r.Value is null)
        {
            return r.Status == 0 ? T("تعذر الوصول إلى الخادم. تأكد من الإنترنت ومن عنوان الخادم.") : r.Message;
        }
        await KeepTokenAsync(r.Value);
        Notify();
        return "";
    }

    /// <summary>Log out: the token is removed; the data and the outbox stay for the next login of the same user.</summary>
    public async Task LogoutAsync()
    {
        api.Token = null;
        await platform.SetSecretAsync(TokenKey, null);
        LoggedIn = false;
        Notify();
    }

    // ------------------------------------------------------------------ sync

    public int Waiting => store.Outbox(OutboxKinds.Pending).Count;

    public int Refused => store.Outbox(OutboxKinds.Rejected).Count;

    public async Task<SyncReport> SyncAsync()
    {
        if (Syncing)
        {
            return LastSync ?? new SyncReport(false, false, 0, 0, Waiting, "", null);
        }
        if (!LoggedIn || MustChangePassword)
        {
            return LastSync = new SyncReport(false, !LoggedIn, 0, 0, Waiting, T("سجّل الدخول أولًا."), null);
        }
        if (!platform.HasNetwork)
        {
            return LastSync = new SyncReport(false, false, 0, 0, Waiting, T("لا اتصال بالإنترنت. ستُرسل العمليات عند عودته."), null);
        }
        Syncing = true;
        Notify();
        try
        {
            var report = await _sync.SyncAsync(Clock);
            if (report.NeedsLogin)
            {
                await LogoutAsync();
            }
            if (report.SnapshotAt is not null)
            {
                (Snapshot, SnapshotAt) = store.LoadSnapshot();
            }
            if (!report.Reached && report.Message.Length > 0)
            {
                report = report with { Message = T("تعذر الوصول إلى الخادم. ستُرسل العمليات عند عودة الاتصال.") };
            }
            return LastSync = report;
        }
        finally
        {
            Syncing = false;
            Notify();
        }
    }

    // ------------------------------------------------------------------ work offline

    public RepCustomer? Customer(int id) => Snapshot?.Customers.FirstOrDefault(c => c.CustomerID == id);

    public OrderDraft? NewOrder(int customerId) =>
        Snapshot is not null && Customer(customerId) is { } c ? new OrderDraft(Snapshot, c) : null;

    /// <summary>Keep the order in the outbox: "" and the order, or the reason.</summary>
    public (string Problem, PendingOrder? Order) SaveOrder(OrderDraft draft)
    {
        var (problem, order) = draft.ToPending(Clock());
        if (order is null)
        {
            return (T(problem), null);
        }
        store.Enqueue(OutboxKinds.Order, order.ClientRef, order, Clock());
        Notify();
        return ("", order);
    }

    public List<BoardVisit> Visits(DateTime day) => VisitBoard.Build(Snapshot, store.Outbox(), day);

    public void PlanDay(DateTime day)
    {
        store.Enqueue(OutboxKinds.VisitPlan, Guid.NewGuid().ToString("N"), new VisitDay(day.Date), Clock());
        Notify();
    }

    public string AddVisit(int customerId, DateTime day)
    {
        if (Visits(day).Any(v => v.CustomerID == customerId))
        {
            return T("العميل في زيارات هذا اليوم.");
        }
        store.Enqueue(OutboxKinds.VisitAdd, Guid.NewGuid().ToString("N"), new VisitAdd(customerId, day.Date), Clock());
        Notify();
        return "";
    }

    /// <summary>The result of a visit: "" = read on the server from the customer's documents that day.</summary>
    public void RecordVisit(BoardVisit visit, DateTime day, string result, string? notes)
    {
        string text = (notes ?? "").Trim();
        store.Enqueue(OutboxKinds.VisitResult, Guid.NewGuid().ToString("N"),
                      new VisitResultEntry(visit.RepVisitID, visit.CustomerID, day.Date, result, text.Length == 0 ? null : text, Clock()),
                      Clock());
        Notify();
    }

    /// <summary>The orders of the phone and of the server, newest first.</summary>
    public List<OrderRow> Orders()
    {
        var rows = new List<OrderRow>();
        var refs = new HashSet<string>();
        foreach (var o in Snapshot?.Orders ?? [])
        {
            if (o.ClientRef is not null)
            {
                refs.Add(o.ClientRef);
            }
            rows.Add(new OrderRow(0, OutboxKinds.Sent, null, FromServer(o), o.OrderNumber, o.Status));
        }
        foreach (var item in store.Outbox().Where(i => i.Kind == OutboxKinds.Order && !refs.Contains(i.ClientRef)))
        {
            rows.Add(new OrderRow(item.Id, item.State, item.Message, item.Read<PendingOrder>(), null, null));
        }
        return rows.OrderByDescending(r => r.Order.TakenAt).ToList();
    }

    public static PendingOrder FromServer(OrderView o) =>
        new(o.ClientRef ?? o.OrderNumber, o.CustomerID, o.CustomerName, o.OrderDate, o.RequiredDate, o.Notes,
            o.Lines.Select(l => new PendingOrderLine(l.ProductID, l.ProductName, l.Quantity, l.UnitPrice)).ToArray());

    public void RetryItem(long id)
    {
        store.Retry(id);
        Notify();
    }

    public void DeleteItem(long id)
    {
        store.Delete(id);
        Notify();
    }

    public List<OutboxItem> Problems() => store.Outbox(OutboxKinds.Rejected);

    // ------------------------------------------------------------------ printing

    /// <summary>Print the customer's copy on the chosen Bluetooth printer: "" or the reason.</summary>
    public async Task<string> PrintAsync(PendingOrder order, string? orderNumber)
    {
        if (Printer.Length == 0)
        {
            return T("اختر الطابعة في الإعدادات.");
        }
        var lines = OrderSlip.Build(order, orderNumber, Snapshot?.Profile.RepName ?? EmployeeName, Snapshot?.StoreName ?? "", T);
        var dots = platform.DrawSlip(lines, EscPos.DotsFor(PaperMillimetres), !English);
        string problem = await platform.PrintAsync(Printer, EscPos.Picture(dots));
        return problem.Length == 0 ? "" : T("تعذرت الطباعة:") + " " + problem;
    }
}
