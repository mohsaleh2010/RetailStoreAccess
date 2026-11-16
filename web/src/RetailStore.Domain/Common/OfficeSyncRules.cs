namespace RetailStore.Domain.Common;

/// <summary>Access kept for the office (plan 6H, docs/84): modOfficeSync. One Access computer brings the journal up to date and
/// sends the pending e-invoices of the documents the web, the Windows POS and the rep app made; the web only reads its state.</summary>
public static class OfficeSyncRules
{
    /// <summary>modOfficeSync.OfficeSyncDue: never ran, or the minutes have passed since the last run; 0 minutes = never.</summary>
    public static bool IsDue(DateTime? lastAt, int minutes, DateTime now) =>
        minutes > 0 && (lastAt is null || WholeMinutes(lastAt.Value, now) >= minutes);

    /// <summary>modOfficeSync.OfficeSyncStale: more than twice the period and 10 minutes without a run (the web warns).</summary>
    public static bool IsStale(DateTime? lastAt, int minutes, DateTime now) =>
        minutes > 0 && (lastAt is null || WholeMinutes(lastAt.Value, now) > 2 * minutes + 10);

    /// <summary>VBA DateDiff("n", a, b): the minute boundaries crossed, not the elapsed time rounded.</summary>
    public static long WholeMinutes(DateTime from, DateTime to)
    {
        static long Ticks(DateTime d) => new DateTime(d.Year, d.Month, d.Day, d.Hour, d.Minute, 0).Ticks / TimeSpan.TicksPerMinute;
        return Ticks(to) - Ticks(from);
    }
}
