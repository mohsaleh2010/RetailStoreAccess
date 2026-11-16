using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage;

namespace RetailStore.Data.Reports;

/// <summary>A condition on a column of a saved query (modScreens.PrepareReport: the period of a "D" report, the c / s / r / e / w
/// filters): Op is >=, &lt; or =.</summary>
public sealed record QueryFilter(string Field, string Op, object Value);

/// <summary>Runs a saved query of Access with its parameters (modQueryParams.QDate / QLong) and filters.
/// SQL Server: the view of the same name (docs/71; [en] in English), its parameters written to dbo.QueryParams for this
/// connection (HOST_NAME() + HOST_ID()) inside a transaction that is rolled back, so two pages of the same web process never see
/// each other's parameters. SQLite (development and tests): the very SQL of Access as views (SavedQueries.Views), with the
/// functions of tests/access_sqlite.py registered on the connection.</summary>
public static partial class SavedQueryRunner
{
    /// <summary>The parameters the report centre sets (SetPeriod, RequireChoice, SetQueryParam); each run sets all of them,
    /// those it does not use to Null (= not set: QDate gives 1899-12-30, QLong 0).</summary>
    public static readonly string[] WebParams =
        ["PeriodStart", "PeriodEnd", "CustomerID", "SupplierID", "ProductID", "WarehouseID", "CashBoxID", "AccountCode"];

    private const string ZeroDate = "1899-12-30 00:00:00";
    private static readonly HashSet<string> Names = SavedQueries.Views.Select(v => v.Name).ToHashSet(StringComparer.Ordinal);

    [GeneratedRegex(@"^[A-Za-z][A-Za-z0-9_]*$")]
    private static partial Regex Identifier();

    public static bool IsSqlite(RetailDbContext db) => db.Database.ProviderName == "Microsoft.EntityFrameworkCore.Sqlite";

    public static async Task<List<ReportRow>> RunAsync(RetailDbContext db, string query, IReadOnlyDictionary<string, object?> parameters,
                                                       IReadOnlyList<QueryFilter> filters, bool english)
    {
        if (!Names.Contains(query))
        {
            throw new ArgumentException("Not a saved query: " + query);
        }
        foreach (var p in parameters.Keys)
        {
            if (!WebParams.Contains(p))
            {
                throw new ArgumentException("Not a parameter of the report centre: " + p);
            }
        }
        foreach (var f in filters)
        {
            if (!Identifier().IsMatch(f.Field) || f.Op is not (">=" or "<" or "="))
            {
                throw new ArgumentException("Not a filter: " + f.Field + " " + f.Op);
            }
        }
        await db.Database.OpenConnectionAsync();
        try
        {
            return IsSqlite(db) ? await RunSqliteAsync(db, query, parameters, filters)
                                : await RunServerAsync(db, query, parameters, filters, english);
        }
        finally
        {
            await db.Database.CloseConnectionAsync();
        }
    }

    private static string Where(IReadOnlyList<QueryFilter> filters, DbCommand cmd, bool sqlite)
    {
        var parts = new List<string>();
        for (int i = 0; i < filters.Count; i++)
        {
            var f = filters[i];
            var p = cmd.CreateParameter();
            p.ParameterName = "@f" + i;
            p.Value = f.Value is DateTime d && sqlite ? d.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) : f.Value;
            cmd.Parameters.Add(p);
            parts.Add($"[{f.Field}] {f.Op} @f{i}");
        }
        return parts.Count == 0 ? "" : " WHERE " + string.Join(" AND ", parts);
    }

    private static async Task<List<ReportRow>> ReadAsync(DbCommand cmd)
    {
        var rows = new List<ReportRow>();
        await using var r = await cmd.ExecuteReaderAsync();
        while (await r.ReadAsync())
        {
            var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < r.FieldCount; i++)
            {
                values[r.GetName(i)] = r.IsDBNull(i) ? null : r.GetValue(i);
            }
            rows.Add(new ReportRow(values));
        }
        return rows;
    }

    // ------------------------------------------------------------------ SQL Server
    private static async Task<List<ReportRow>> RunServerAsync(RetailDbContext db, string query, IReadOnlyDictionary<string, object?> parameters,
                                                              IReadOnlyList<QueryFilter> filters, bool english)
    {
        await using var tx = await db.Database.BeginTransactionAsync();
        var cn = db.Database.GetDbConnection();
        await using (var merge = cn.CreateCommand())
        {
            merge.Transaction = tx.GetDbTransaction();
            var sets = string.Join(", ", WebParams.Select(p => $"[{p}] = @{p}"));
            var cols = string.Join(", ", WebParams.Select(p => $"[{p}]"));
            var vals = string.Join(", ", WebParams.Select(p => "@" + p));
            merge.CommandText =
                "MERGE [dbo].[QueryParams] AS t\n" +
                "USING (SELECT HOST_NAME() AS HostName, HOST_ID() AS HostProcess) AS s\n" +
                "ON t.HostName = s.HostName AND t.HostProcess = s.HostProcess\n" +
                $"WHEN MATCHED THEN UPDATE SET {sets}, UpdatedAt = GETDATE()\n" +
                $"WHEN NOT MATCHED THEN INSERT (HostName, HostProcess, {cols}) VALUES (s.HostName, s.HostProcess, {vals});";
            foreach (var p in WebParams)
            {
                var prm = merge.CreateParameter();
                prm.ParameterName = "@" + p;
                prm.DbType = p.StartsWith("Period", StringComparison.Ordinal) ? DbType.DateTime : DbType.Int32;
                prm.Value = parameters.TryGetValue(p, out var v) && v is not null ? v : DBNull.Value;
                merge.Parameters.Add(prm);
            }
            await merge.ExecuteNonQueryAsync();
        }
        await using var cmd = cn.CreateCommand();
        cmd.Transaction = tx.GetDbTransaction();
        cmd.CommandText = $"SELECT * FROM [{(english ? "en" : "dbo")}].[{query}]" + Where(filters, cmd, false);
        var rows = await ReadAsync(cmd);
        await tx.RollbackAsync();          // the parameters were for this run only
        return rows;
    }

    // ------------------------------------------------------------------ SQLite
    private static async Task<List<ReportRow>> RunSqliteAsync(RetailDbContext db, string query, IReadOnlyDictionary<string, object?> parameters,
                                                              IReadOnlyList<QueryFilter> filters)
    {
        var cn = (SqliteConnection)db.Database.GetDbConnection();
        RegisterFunctions(cn, parameters);
        await using var cmd = cn.CreateCommand();
        cmd.Transaction = db.Database.CurrentTransaction?.GetDbTransaction() as SqliteTransaction;
        cmd.CommandText = $"SELECT * FROM [{query}]" + Where(filters, cmd, true);
        return await ReadAsync(cmd);
    }

    /// <summary>The views of a SQLite database: created again when the saved queries changed (PRAGMA user_version keeps their
    /// fingerprint).</summary>
    public static void EnsureSqliteViews(RetailDbContext db)
    {
        db.Database.OpenConnection();
        try
        {
            var cn = (SqliteConnection)db.Database.GetDbConnection();
            int stamp = Fingerprint();
            using (var check = cn.CreateCommand())
            {
                check.CommandText = "PRAGMA user_version";
                if (Convert.ToInt32(check.ExecuteScalar(), CultureInfo.InvariantCulture) == stamp)
                {
                    return;
                }
            }
            var old = new List<string>();
            using (var list = cn.CreateCommand())
            {
                list.CommandText = "SELECT name FROM sqlite_master WHERE type = 'view'";
                using var r = list.ExecuteReader();
                while (r.Read())
                {
                    old.Add(r.GetString(0));
                }
            }
            using var tx = cn.BeginTransaction();
            foreach (var name in old)
            {
                using var drop = cn.CreateCommand();
                drop.Transaction = tx;
                drop.CommandText = $"DROP VIEW IF EXISTS [{name}]";
                drop.ExecuteNonQuery();
            }
            foreach (var (name, sql) in SavedQueries.Views)
            {
                using var create = cn.CreateCommand();
                create.Transaction = tx;
                create.CommandText = $"CREATE VIEW [{name}] AS {sql}";
                create.ExecuteNonQuery();
            }
            using (var set = cn.CreateCommand())
            {
                set.Transaction = tx;
                set.CommandText = "PRAGMA user_version = " + stamp.ToString(CultureInfo.InvariantCulture);
                set.ExecuteNonQuery();
            }
            tx.Commit();
        }
        finally
        {
            db.Database.CloseConnection();
        }
    }

    private static int Fingerprint()
    {
        unchecked
        {
            int h = 17;
            foreach (var (name, sql) in SavedQueries.Views)
            {
                foreach (char c in name + sql)
                {
                    h = h * 31 + c;
                }
            }
            return h == 0 ? 1 : h & 0x7FFFFFFF;
        }
    }

    private static object? N(object? v) => v is DBNull ? null : v;

    private static string Iso(DateTime d) => d.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

    private static DateTime ParseDate(object v) =>
        v is DateTime d ? d : DateTime.Parse(Convert.ToString(v, CultureInfo.InvariantCulture)!, CultureInfo.InvariantCulture);

    /// <summary>The functions of tests/access_sqlite.py (the Access / VBA functions of the saved queries) and the parameters of
    /// this run (QDate / QLong).</summary>
    public static void RegisterFunctions(SqliteConnection cn, IReadOnlyDictionary<string, object?> parameters)
    {
        cn.CreateFunction<object?, object?, object?>("Nz", (v, d) => N(v) ?? N(d));
        cn.CreateFunction<object?, string>("Space", n => new string(' ', (int)Math.Max(0, ReportRow.ToNumber(N(n)))));
        cn.CreateFunction<object?, object?>("CCur", v => N(v) is { } x ? Math.Round(Convert.ToDouble(x, CultureInfo.InvariantCulture), 4) : null);
        cn.CreateFunction<object?, object?>("CStr", v => N(v) switch
        {
            null => null,
            double d when d == Math.Floor(d) && Math.Abs(d) < 1e15 => ((long)d).ToString(CultureInfo.InvariantCulture),
            var x => Convert.ToString(x, CultureInfo.InvariantCulture),
        });
        cn.CreateFunction<object?, object?>("CLng", v => N(v) is { } x
            ? (long)Math.Round(Convert.ToDouble(x, CultureInfo.InvariantCulture), MidpointRounding.ToEven) : null);
        cn.CreateFunction<object?, object?>("DateValue", v => N(v) is { } x ? Iso(ParseDate(x).Date) : null);
        cn.CreateFunction<object?, object?>("Year", v => N(v) is { } x ? ParseDate(x).Year : null);
        cn.CreateFunction<object?, object?>("Month", v => N(v) is { } x ? ParseDate(x).Month : null);
        cn.CreateFunction("Date", () => Iso(DateTime.Today));
        cn.CreateFunction<string, object?, object?, object?>("DateAdd", (interval, n, v) =>
            N(v) is { } x ? Iso(ParseDate(x).AddDays(Convert.ToDouble(N(n) ?? 0, CultureInfo.InvariantCulture))) : null);
        cn.CreateFunction<string, object?, object?, object?>("DateDiff", (interval, a, b) =>
            N(a) is { } x && N(b) is { } y ? (long)(ParseDate(y).Date - ParseDate(x).Date).TotalDays : null);
        cn.CreateFunction<string, string>("QDate", name =>
            parameters.TryGetValue(name, out var v) && v is DateTime d ? Iso(d) : ZeroDate, isDeterministic: false);
        cn.CreateFunction<string, long>("QLong", name =>
            parameters.TryGetValue(name, out var v) && v is not null ? Convert.ToInt64(v, CultureInfo.InvariantCulture) : 0L,
            isDeterministic: false);
    }
}
