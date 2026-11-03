using System.Globalization;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using RetailStore.RepApp.Client;

namespace RetailStore.RepApp.Offline;

/// <summary>What the rep recorded offline, waiting to be sent (docs/78): an order, a visit plan, a visit outside the
/// route, or the result of a visit. Sent in order; State PENDING -> SENT (ServerId) or REJECTED (Message, the reason the
/// program gave: the rep fixes it or deletes it).</summary>
public sealed record OutboxItem(long Id, string Kind, string ClientRef, string Payload, DateTime CreatedAt, string State,
                                string? Message, int? ServerId)
{
    public T Read<T>() => JsonSerializer.Deserialize<T>(Payload, RepApiClient.Json)!;
}

public static class OutboxKinds
{
    public const string Order = "ORDER";
    public const string VisitPlan = "VISIT_PLAN";
    public const string VisitAdd = "VISIT_ADD";
    public const string VisitResult = "VISIT_RESULT";

    public const string Pending = "PENDING";
    public const string Sent = "SENT";
    public const string Rejected = "REJECTED";
}

/// <summary>An order taken on the phone: the prices the customer was shown, the time it was taken, and its number on
/// the phone (ClientRef) that keeps the server from adding it twice.</summary>
public sealed record PendingOrderLine(int ProductID, string ProductName, decimal Quantity, decimal UnitPrice);

public sealed record PendingOrder(string ClientRef, int CustomerID, string CustomerName, DateTime TakenAt, DateTime? RequiredDate,
                                  string? Notes, PendingOrderLine[] Lines)
{
    public decimal Total => OrderDraft.RoundMoney(Lines.Sum(l => l.Quantity * l.UnitPrice));

    public OrderRequest ToRequest() =>
        new(CustomerID, null, RequiredDate, Notes, Lines.Select(l => new OrderLineInput(l.ProductID, l.Quantity, l.UnitPrice)).ToArray(),
            ClientRef, TakenAt);
}

public sealed record VisitDay(DateTime Date);

public sealed record VisitAdd(int CustomerID, DateTime Date);

/// <summary>The result of a visit: VisitID 0 = a visit made on the phone (plan or outside the route), found on the
/// server by its customer and day when it is sent. Result "" = read from the customer's documents that day.</summary>
public sealed record VisitResultEntry(int VisitID, int CustomerID, DateTime Date, string Result, string? Notes, DateTime RecordedAt);

/// <summary>The SQLite file of the rep app on the phone: the settings, the last snapshot of the server, and the outbox.
/// The token is not here: it is in the secure storage of the phone (IRepPlatform).</summary>
public sealed class LocalStore
{
    private readonly string _connection;

    public LocalStore(string path)
    {
        _connection = new SqliteConnectionStringBuilder { DataSource = path, Pooling = false }.ToString();
        Execute("""
            CREATE TABLE IF NOT EXISTS Settings (Key TEXT PRIMARY KEY, Value TEXT);
            CREATE TABLE IF NOT EXISTS Snapshot (Id INTEGER PRIMARY KEY CHECK (Id = 1), Json TEXT NOT NULL, SavedAt TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS Outbox (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, Kind TEXT NOT NULL, ClientRef TEXT NOT NULL, Payload TEXT NOT NULL,
                CreatedAt TEXT NOT NULL, State TEXT NOT NULL DEFAULT 'PENDING', Message TEXT, ServerId INTEGER);
            """);
    }

    private SqliteConnection Open()
    {
        var c = new SqliteConnection(_connection);
        c.Open();
        return c;
    }

    private void Execute(string sql, params (string Name, object? Value)[] args)
    {
        using var c = Open();
        using var cmd = c.CreateCommand();
        cmd.CommandText = sql;
        foreach (var (name, value) in args)
        {
            cmd.Parameters.AddWithValue(name, value ?? DBNull.Value);
        }
        cmd.ExecuteNonQuery();
    }

    private static string Stamp(DateTime d) => d.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

    private static DateTime ParseStamp(string s) => DateTime.ParseExact(s, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

    // ------------------------------------------------------------------ settings

    public string? Get(string key)
    {
        using var c = Open();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT Value FROM Settings WHERE Key = $k";
        cmd.Parameters.AddWithValue("$k", key);
        return cmd.ExecuteScalar() as string;
    }

    public void Set(string key, string? value) =>
        Execute("INSERT INTO Settings (Key, Value) VALUES ($k, $v) ON CONFLICT(Key) DO UPDATE SET Value = $v", ("$k", key), ("$v", value));

    // ------------------------------------------------------------------ the snapshot of the server

    public (RepSnapshot? Snapshot, DateTime? SavedAt) LoadSnapshot()
    {
        using var c = Open();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT Json, SavedAt FROM Snapshot WHERE Id = 1";
        using var r = cmd.ExecuteReader();
        if (!r.Read())
        {
            return (null, null);
        }
        return (JsonSerializer.Deserialize<RepSnapshot>(r.GetString(0), RepApiClient.Json), ParseStamp(r.GetString(1)));
    }

    public void SaveSnapshot(RepSnapshot snapshot, DateTime savedAt) =>
        Execute("INSERT INTO Snapshot (Id, Json, SavedAt) VALUES (1, $j, $s) ON CONFLICT(Id) DO UPDATE SET Json = $j, SavedAt = $s",
                ("$j", JsonSerializer.Serialize(snapshot, RepApiClient.Json)), ("$s", Stamp(savedAt)));

    /// <summary>Another user logs in on the phone: what the last one kept is removed (his outbox is sent first).</summary>
    public void ClearData() => Execute("DELETE FROM Snapshot; DELETE FROM Outbox;");

    // ------------------------------------------------------------------ the outbox

    public long Enqueue(string kind, string clientRef, object payload, DateTime at)
    {
        using var c = Open();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "INSERT INTO Outbox (Kind, ClientRef, Payload, CreatedAt) VALUES ($k, $r, $p, $c); SELECT last_insert_rowid();";
        cmd.Parameters.AddWithValue("$k", kind);
        cmd.Parameters.AddWithValue("$r", clientRef);
        cmd.Parameters.AddWithValue("$p", JsonSerializer.Serialize(payload, payload.GetType(), RepApiClient.Json));
        cmd.Parameters.AddWithValue("$c", Stamp(at));
        return (long)cmd.ExecuteScalar()!;
    }

    public List<OutboxItem> Outbox(params string[] states)
    {
        using var c = Open();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT Id, Kind, ClientRef, Payload, CreatedAt, State, Message, ServerId FROM Outbox ORDER BY Id";
        var list = new List<OutboxItem>();
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var item = new OutboxItem(r.GetInt64(0), r.GetString(1), r.GetString(2), r.GetString(3), ParseStamp(r.GetString(4)),
                                      r.GetString(5), r.IsDBNull(6) ? null : r.GetString(6), r.IsDBNull(7) ? null : r.GetInt32(7));
            if (states.Length == 0 || states.Contains(item.State))
            {
                list.Add(item);
            }
        }
        return list;
    }

    public void MarkSent(long id, int? serverId) =>
        Execute("UPDATE Outbox SET State = 'SENT', Message = NULL, ServerId = $s WHERE Id = $i", ("$s", serverId), ("$i", id));

    public void MarkRejected(long id, string message) =>
        Execute("UPDATE Outbox SET State = 'REJECTED', Message = $m WHERE Id = $i", ("$m", message), ("$i", id));

    /// <summary>Send a rejected item again (after the office fixed the reason, e.g. activated the customer).</summary>
    public void Retry(long id) => Execute("UPDATE Outbox SET State = 'PENDING', Message = NULL WHERE Id = $i", ("$i", id));

    public void Delete(long id) => Execute("DELETE FROM Outbox WHERE Id = $i AND State <> 'SENT'", ("$i", id));

    /// <summary>Sent orders are kept until the snapshot shows them (the server's orders carry their ClientRef), or until
    /// they are older than the given time; the other sent items go once the snapshot is read.</summary>
    public void ForgetSent(IEnumerable<string> onServer, DateTime olderThan)
    {
        foreach (var r in onServer)
        {
            Execute("DELETE FROM Outbox WHERE State = 'SENT' AND ClientRef = $r", ("$r", r));
        }
        Execute("DELETE FROM Outbox WHERE State = 'SENT' AND (Kind <> 'ORDER' OR CreatedAt < $o)", ("$o", Stamp(olderThan)));
    }
}
