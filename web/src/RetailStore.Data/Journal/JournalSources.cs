using System.Data;
using System.Globalization;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage;
using RetailStore.Domain;
using RetailStore.Domain.Journal;

namespace RetailStore.Data.Journal;

/// <summary>Where the lines of the journal source queries come from.</summary>
public interface IJournalSourceReader
{
    Task<List<JournalSourceLine>> ReadAsync(RetailDbContext db, string view);
}

/// <summary>SQL Server: each journal source query is a view of the same name (docs/71, dist/sql/RetailStore_Queries.sql),
/// the very rules Access runs; the web does not write them again.</summary>
public sealed class ServerViewJournalReader : IJournalSourceReader
{
    public async Task<List<JournalSourceLine>> ReadAsync(RetailDbContext db, string view)
    {
        if (!AppData.JournalSourceViews.Contains(view))
        {
            throw new ArgumentException("Not a journal source view: " + view);
        }
        var cn = db.Database.GetDbConnection();
        if (cn.State != ConnectionState.Open)
        {
            await db.Database.OpenConnectionAsync();
        }
        await using var cmd = cn.CreateCommand();
        cmd.Transaction = db.Database.CurrentTransaction?.GetDbTransaction();
        cmd.CommandText = "SELECT SourceType, SourceID, SourceNumber, SourceDate, Party, LineOrder, AccountCode, Debit, Credit, " +
                          "LineText, CostCenter FROM [dbo].[" + view + "]";
        var list = new List<JournalSourceLine>();
        await using var r = await cmd.ExecuteReaderAsync();
        while (await r.ReadAsync())
        {
            list.Add(new JournalSourceLine(
                Convert.ToString(r[0], CultureInfo.InvariantCulture)!, Convert.ToInt32(r[1], CultureInfo.InvariantCulture),
                r.IsDBNull(2) ? null : Convert.ToString(r[2], CultureInfo.InvariantCulture),
                r.IsDBNull(3) ? null : Convert.ToDateTime(r[3], CultureInfo.InvariantCulture),
                r.IsDBNull(4) ? null : Convert.ToString(r[4], CultureInfo.InvariantCulture),
                Convert.ToInt32(r[5], CultureInfo.InvariantCulture), Convert.ToInt32(r[6], CultureInfo.InvariantCulture),
                r.IsDBNull(7) ? 0m : Convert.ToDecimal(r[7], CultureInfo.InvariantCulture),
                r.IsDBNull(8) ? 0m : Convert.ToDecimal(r[8], CultureInfo.InvariantCulture),
                r.IsDBNull(9) ? null : Convert.ToString(r[9], CultureInfo.InvariantCulture),
                r.IsDBNull(10) ? 0 : Convert.ToInt32(r[10], CultureInfo.InvariantCulture)));
        }
        return list;
    }
}

/// <summary>Fixed lines (tests): view name -> its lines.</summary>
public sealed class FixedJournalReader(IReadOnlyDictionary<string, List<JournalSourceLine>> lines) : IJournalSourceReader
{
    public Task<List<JournalSourceLine>> ReadAsync(RetailDbContext db, string view) =>
        Task.FromResult(lines.TryGetValue(view, out var l) ? l : []);
}
