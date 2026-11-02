using RetailStore.Domain.Common;

namespace RetailStore.Domain.Journal;

/// <summary>One line of a journal source view (qryJournal*, queries.jline): every operation gives its lines.</summary>
public sealed record JournalSourceLine(
    string SourceType, int SourceID, string? SourceNumber, DateTime? SourceDate, string? Party,
    int LineOrder, int AccountCode, decimal Debit, decimal Credit, string? LineText, int CostCenter);

/// <summary>A journal entry already in JournalEntries, as SyncJournal reads it.</summary>
public sealed record ExistingEntry(int EntryID, string SourceType, int SourceID, decimal Signature, DateTime EntryDate,
                                   decimal TotalDebit);

/// <summary>The header of an entry to write, from the lines of one operation (modJournal.FillHeader).</summary>
public sealed record PlannedHeader(
    string SourceType, int SourceID, string SourceNumber, DateTime EntryDate, string Description,
    decimal TotalDebit, decimal TotalCredit, decimal Signature, IReadOnlyList<JournalSourceLine> Lines);

public sealed class JournalPlan
{
    public List<PlannedHeader> Added { get; } = [];
    public List<(int EntryID, PlannedHeader Header)> Updated { get; } = [];
    public List<int> Removed { get; } = [];
    /// <summary>Operations whose debit and credit differ: no entry (shown to the user).</summary>
    public List<string> Unbalanced { get; } = [];
    /// <summary>Operations in a closed period that would change: left as they are (shown to the user).</summary>
    public List<string> Locked { get; } = [];
}

/// <summary>The decisions of modJournal.SyncJournal (tools/sim.py sync_journal), without the database:
/// a new operation gets an entry, a changed one (signature, date or total) is rebuilt under the same number,
/// a deleted one loses its entry; nothing changes in the closed period (Settings.ClosedThrough) but the year
/// closing; an unbalanced operation gets no entry.</summary>
public static class JournalPlanner
{
    public static readonly DateTime NoDate = new(2000, 1, 1);

    /// <summary>The signature of an entry: Sum(AccountCode x (2 Debit + Credit)) + Sum(CostCenter x (2 Debit + Credit) x 7),
    /// as Currency (4 decimals).</summary>
    public static decimal Signature(IEnumerable<JournalSourceLine> lines) =>
        MoneyMath.Currency(lines.Sum(l => l.AccountCode * (l.Debit + l.Debit + l.Credit))
                           + lines.Sum(l => l.CostCenter * (l.Debit + l.Debit + l.Credit) * 7));

    public static bool InClosedPeriod(string sourceType, DateTime? when, DateTime? closedThrough) =>
        closedThrough is not null && sourceType != "YEAR_CLOSE" && when is not null && when.Value.Date <= closedThrough.Value.Date;

    /// <param name="sources">The lines of each source view, in the order of AppData.JournalSourceViews.</param>
    /// <param name="kinds">JournalSourceTypes: SourceType -> its name (the description of the entry).</param>
    public static JournalPlan Plan(IEnumerable<IEnumerable<JournalSourceLine>> sources, IEnumerable<ExistingEntry> existing,
                                   IReadOnlyDictionary<string, string> kinds, DateTime? closedThrough)
    {
        var plan = new JournalPlan();
        var old = existing.ToDictionary(e => (e.SourceType, e.SourceID));
        var seen = new HashSet<(string, int)>();
        foreach (var view in sources)
        {
            foreach (var op in view.GroupBy(l => (l.SourceType, l.SourceID)))
            {
                seen.Add(op.Key);
                var lines = op.OrderBy(l => l.LineOrder).ToList();
                decimal debit = MoneyMath.Currency(lines.Sum(l => l.Debit));
                decimal credit = MoneyMath.Currency(lines.Sum(l => l.Credit));
                string kind = kinds.TryGetValue(op.Key.SourceType, out var k) ? k : op.Key.SourceType;
                string? number = lines.Select(l => l.SourceNumber).Where(n => n is not null).Max(StringComparer.Ordinal);
                if (debit != credit)
                {
                    plan.Unbalanced.Add(kind + " " + (number ?? ""));
                    continue;
                }
                var header = Header(op.Key.SourceType, op.Key.SourceID, lines, kind, debit, credit);
                if (!old.TryGetValue(op.Key, out var entry))
                {
                    if (InClosedPeriod(header.SourceType, header.EntryDate, closedThrough))
                    {
                        plan.Locked.Add(kind + " " + (number ?? ""));
                    }
                    else
                    {
                        plan.Added.Add(header);
                    }
                }
                else if (entry.Signature == header.Signature && entry.EntryDate == header.EntryDate
                         && entry.TotalDebit == header.TotalDebit)
                {
                    // unchanged
                }
                else if (InClosedPeriod(header.SourceType, entry.EntryDate, closedThrough)
                         || InClosedPeriod(header.SourceType, header.EntryDate, closedThrough))
                {
                    plan.Locked.Add(kind + " " + (number ?? ""));
                }
                else
                {
                    plan.Updated.Add((entry.EntryID, header));
                }
            }
        }
        foreach (var (key, entry) in old)
        {
            if (seen.Contains(key))
            {
                continue;
            }
            if (InClosedPeriod(key.SourceType, entry.EntryDate, closedThrough))
            {
                string kind = kinds.TryGetValue(key.SourceType, out var k) ? k : key.SourceType;
                plan.Locked.Add(kind + " (" + entry.EntryDate.ToString("yyyy/MM/dd") + ")");
            }
            else
            {
                plan.Removed.Add(entry.EntryID);
            }
        }
        return plan;
    }

    private static PlannedHeader Header(string sourceType, int sourceID, List<JournalSourceLine> lines, string kind,
                                        decimal debit, decimal credit)
    {
        string? number = lines.Select(l => l.SourceNumber).Where(n => n is not null).Max(StringComparer.Ordinal);
        string docNo = Left(number ?? "-", 20);
        string? party = lines.Select(l => l.Party).Where(p => p is not null).Max(StringComparer.Ordinal);
        string text = kind + " " + docNo;
        if (!string.IsNullOrEmpty(party))
        {
            text += " - " + party;
        }
        DateTime? when = lines.Select(l => l.SourceDate).Where(d => d is not null).Max();
        return new PlannedHeader(sourceType, sourceID, docNo, when ?? NoDate, Left(text, 255), debit, credit,
                                 Signature(lines), lines);
    }

    public static string Left(string s, int n) => s.Length <= n ? s : s[..n];
}
