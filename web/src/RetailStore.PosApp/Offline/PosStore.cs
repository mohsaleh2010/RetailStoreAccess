using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using RetailStore.Data.Services;
using RetailStore.PosApp.Connection;

namespace RetailStore.PosApp.Offline;

public static class SaleStates
{
    public const string Pending = "PENDING";
    public const string Sent = "SENT";
    public const string Rejected = "REJECTED";
}

/// <summary>A sale made on this PC: what was sent (Request), when, by whom, and what the program answered: PENDING ->
/// SENT (ServerId, SentAt) or REJECTED (Message: the office fixes the reason, then it is sent again).</summary>
public sealed record StoredSale(long Id, string ClientRef, string InvoiceNumber, int EmployeeID, string EmployeeName, DateTime CreatedAt,
                                string State, string? Message, int? ServerId, DateTime? SentAt, PosReceipt Receipt)
{
    public DeviceSaleRequest Request => Receipt.Request;
}

/// <summary>A user who logged in on this PC while the data answered: he may log in again without it (his password
/// checked against a PBKDF2 digest kept here, never the program's digest), with what the program last said of him.</summary>
public sealed record CachedUser(int EmployeeID, string Username, string EmployeeName, bool IsAdmin, DateTime LastOnline);

/// <summary>The SQLite file of the POS on the PC (docs/83): the settings, the connection to the data (without its
/// password), the users for the offline login, the last snapshot of each user, the counter of the invoice numbers and the
/// sales. The secrets (the SQL password, the API tokens) are in the protected storage of Windows (IPosPlatform).</summary>
public sealed class PosStore
{
    private readonly string _connection;
    public const int Iterations = 100_000;

    public PosStore(string path)
    {
        _connection = new SqliteConnectionStringBuilder { DataSource = path, Pooling = false }.ToString();
        Execute("""
            CREATE TABLE IF NOT EXISTS Settings (Key TEXT PRIMARY KEY, Value TEXT);
            CREATE TABLE IF NOT EXISTS Users (
                EmployeeID INTEGER PRIMARY KEY, Username TEXT NOT NULL, EmployeeName TEXT NOT NULL, IsAdmin INTEGER NOT NULL,
                Salt TEXT NOT NULL, Digest TEXT NOT NULL, LastOnline TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS Snapshots (EmployeeID INTEGER PRIMARY KEY, Json TEXT NOT NULL, SavedAt TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS Sales (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, ClientRef TEXT NOT NULL UNIQUE, InvoiceNumber TEXT NOT NULL UNIQUE,
                EmployeeID INTEGER NOT NULL, EmployeeName TEXT NOT NULL, Payload TEXT NOT NULL, CreatedAt TEXT NOT NULL,
                State TEXT NOT NULL DEFAULT 'PENDING', Message TEXT, ServerId INTEGER, SentAt TEXT);
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

    // ------------------------------------------------------------------ settings and the connection

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

    public PosConnection? LoadConnection() =>
        Get("Connection") is { } json ? JsonSerializer.Deserialize<PosConnection>(json, ApiBackend.Json) : null;

    public void SaveConnection(PosConnection connection) => Set("Connection", JsonSerializer.Serialize(connection, ApiBackend.Json));

    /// <summary>The next invoice number of this PC (its code and a counter that never goes back).</summary>
    public string NextInvoiceNumber(string deviceCode)
    {
        long n = long.TryParse(Get("NextNumber"), out var v) ? v : 1;
        Set("NextNumber", (n + 1).ToString(CultureInfo.InvariantCulture));
        return PosService.DeviceNumber(deviceCode, n);
    }

    /// <summary>The counter is moved past a number the program already has (a new PC with the code of an old one).</summary>
    public void SkipNumbersTo(long next)
    {
        long n = long.TryParse(Get("NextNumber"), out var v) ? v : 1;
        if (next > n)
        {
            Set("NextNumber", next.ToString(CultureInfo.InvariantCulture));
        }
    }

    // ------------------------------------------------------------------ users for the offline login

    public static string Digest(string password, string salt) =>
        Convert.ToBase64String(Rfc2898DeriveBytes.Pbkdf2(password ?? "", Convert.FromBase64String(salt), Iterations,
                                                         HashAlgorithmName.SHA256, 32));

    public void RememberUser(int employeeId, string username, string name, bool isAdmin, string password, DateTime at)
    {
        string salt = Convert.ToBase64String(RandomNumberGenerator.GetBytes(16));
        Execute("""
            INSERT INTO Users (EmployeeID, Username, EmployeeName, IsAdmin, Salt, Digest, LastOnline) VALUES ($i, $u, $n, $a, $s, $d, $t)
            ON CONFLICT(EmployeeID) DO UPDATE SET Username = $u, EmployeeName = $n, IsAdmin = $a, Salt = $s, Digest = $d, LastOnline = $t
            """, ("$i", employeeId), ("$u", username.Trim()), ("$n", name), ("$a", isAdmin ? 1 : 0), ("$s", salt),
                ("$d", Digest(password, salt)), ("$t", Stamp(at)));
    }

    public void UpdateUserFacts(int employeeId, string name, bool isAdmin) =>
        Execute("UPDATE Users SET EmployeeName = $n, IsAdmin = $a WHERE EmployeeID = $i", ("$n", name), ("$a", isAdmin ? 1 : 0),
                ("$i", employeeId));

    /// <summary>The program refused the login (a new password, a disabled user): he may no longer log in offline.</summary>
    public void ForgetUser(string username) => Execute("DELETE FROM Users WHERE lower(Username) = lower($u)", ("$u", username.Trim()));

    /// <summary>The user whose password matches the digest kept on this PC, else null.</summary>
    public CachedUser? CheckOffline(string username, string password)
    {
        using var c = Open();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT EmployeeID, Username, EmployeeName, IsAdmin, Salt, Digest, LastOnline FROM Users WHERE lower(Username) = lower($u)";
        cmd.Parameters.AddWithValue("$u", (username ?? "").Trim());
        using var r = cmd.ExecuteReader();
        if (!r.Read())
        {
            return null;
        }
        byte[] want = Convert.FromBase64String(r.GetString(5));
        byte[] got = Convert.FromBase64String(Digest(password, r.GetString(4)));
        return CryptographicOperations.FixedTimeEquals(want, got)
            ? new CachedUser(r.GetInt32(0), r.GetString(1), r.GetString(2), r.GetInt32(3) == 1, ParseStamp(r.GetString(6)))
            : null;
    }

    public List<CachedUser> Users()
    {
        using var c = Open();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT EmployeeID, Username, EmployeeName, IsAdmin, LastOnline FROM Users ORDER BY Username";
        using var r = cmd.ExecuteReader();
        var list = new List<CachedUser>();
        while (r.Read())
        {
            list.Add(new CachedUser(r.GetInt32(0), r.GetString(1), r.GetString(2), r.GetInt32(3) == 1, ParseStamp(r.GetString(4))));
        }
        return list;
    }

    // ------------------------------------------------------------------ the snapshot of each user

    public (PosSnapshot? Snapshot, DateTime? SavedAt) LoadSnapshot(int employeeId)
    {
        using var c = Open();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT Json, SavedAt FROM Snapshots WHERE EmployeeID = $i";
        cmd.Parameters.AddWithValue("$i", employeeId);
        using var r = cmd.ExecuteReader();
        if (!r.Read())
        {
            return (null, null);
        }
        return (JsonSerializer.Deserialize<PosSnapshot>(r.GetString(0), ApiBackend.Json), ParseStamp(r.GetString(1)));
    }

    public void SaveSnapshot(int employeeId, PosSnapshot snapshot, DateTime savedAt) =>
        Execute("INSERT INTO Snapshots (EmployeeID, Json, SavedAt) VALUES ($i, $j, $s) ON CONFLICT(EmployeeID) DO UPDATE SET Json = $j, SavedAt = $s",
                ("$i", employeeId), ("$j", JsonSerializer.Serialize(snapshot, ApiBackend.Json)), ("$s", Stamp(savedAt)));

    /// <summary>Another database: what was kept of the last one is removed (its sales must have been sent first).</summary>
    public void ClearData() => Execute("DELETE FROM Snapshots; DELETE FROM Users; DELETE FROM Sales WHERE State = 'SENT';");

    // ------------------------------------------------------------------ the sales of this PC

    public long AddSale(PosReceipt receipt, int employeeId, string employeeName, DateTime at)
    {
        using var c = Open();
        using var cmd = c.CreateCommand();
        cmd.CommandText = """
            INSERT INTO Sales (ClientRef, InvoiceNumber, EmployeeID, EmployeeName, Payload, CreatedAt) VALUES ($r, $n, $e, $m, $p, $c);
            SELECT last_insert_rowid();
            """;
        cmd.Parameters.AddWithValue("$r", receipt.Request.ClientRef);
        cmd.Parameters.AddWithValue("$n", receipt.Request.InvoiceNumber);
        cmd.Parameters.AddWithValue("$e", employeeId);
        cmd.Parameters.AddWithValue("$m", employeeName);
        cmd.Parameters.AddWithValue("$p", JsonSerializer.Serialize(receipt, ApiBackend.Json));
        cmd.Parameters.AddWithValue("$c", Stamp(at));
        return (long)cmd.ExecuteScalar()!;
    }

    public List<StoredSale> Sales(params string[] states)
    {
        using var c = Open();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT Id, ClientRef, InvoiceNumber, EmployeeID, EmployeeName, CreatedAt, State, Message, ServerId, Payload, SentAt " +
                          "FROM Sales ORDER BY Id";
        var list = new List<StoredSale>();
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var s = new StoredSale(r.GetInt64(0), r.GetString(1), r.GetString(2), r.GetInt32(3), r.GetString(4), ParseStamp(r.GetString(5)),
                                   r.GetString(6), r.IsDBNull(7) ? null : r.GetString(7), r.IsDBNull(8) ? null : r.GetInt32(8),
                                   r.IsDBNull(10) ? null : ParseStamp(r.GetString(10)),
                                   JsonSerializer.Deserialize<PosReceipt>(r.GetString(9), ApiBackend.Json)!);
            if (states.Length == 0 || states.Contains(s.State))
            {
                list.Add(s);
            }
        }
        return list;
    }

    public void MarkSent(long id, int serverId, DateTime at) =>
        Execute("UPDATE Sales SET State = 'SENT', Message = NULL, ServerId = $s, SentAt = $t WHERE Id = $i", ("$s", serverId),
                ("$t", Stamp(at)), ("$i", id));

    public void MarkRejected(long id, string message) =>
        Execute("UPDATE Sales SET State = 'REJECTED', Message = $m WHERE Id = $i", ("$m", message), ("$i", id));

    /// <summary>Send a refused sale again (after the office fixed the reason); confirmStock = with the answer yes to the
    /// question on the stock.</summary>
    public void Retry(long id, bool confirmStock)
    {
        var sale = Sales().FirstOrDefault(s => s.Id == id);
        if (sale is null || sale.State != SaleStates.Rejected)
        {
            return;
        }
        var receipt = confirmStock ? sale.Receipt with { Request = sale.Request with { ConfirmStock = true } } : sale.Receipt;
        Execute("UPDATE Sales SET State = 'PENDING', Message = NULL, Payload = $p WHERE Id = $i",
                ("$p", JsonSerializer.Serialize(receipt, ApiBackend.Json)), ("$i", id));
    }

    /// <summary>A refused sale removed from the PC (the office records it itself): never a sent one.</summary>
    public void Delete(long id) => Execute("DELETE FROM Sales WHERE Id = $i AND State = 'REJECTED'", ("$i", id));

    /// <summary>Sent sales older than the given time are forgotten (the program has them).</summary>
    public void ForgetSent(DateTime olderThan) =>
        Execute("DELETE FROM Sales WHERE State = 'SENT' AND CreatedAt < $o", ("$o", Stamp(olderThan)));
}
