using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Data.Journal;
using RetailStore.Domain.Common;

namespace RetailStore.Data.Services;

/// <summary>One run of the office tasks: what was done (OfficeSyncLog) and the problem ("" = none).</summary>
public sealed record OfficeSyncRun(int JournalAdded, int JournalUpdated, int JournalRemoved, int Sent, int NotSent, int Checked,
                                   string Problem, string Summary);

/// <summary>The office tasks in C# (plan 7A, docs/85): modOfficeSync.RunOfficeSync without Access. SyncJournal (JournalService),
/// then the pending e-invoices sent and, in Egypt, the submitted ones followed (EInvoiceService); each run is a row of
/// OfficeSyncLog and Settings.OfficeSyncLastAt / OfficeSyncLastResult, as Access writes them. The Windows service
/// (RetailStore.OfficeService) runs it when Settings.OfficeSyncComputer names the service and the time has come.</summary>
public sealed class OfficeSyncService(IDbContextFactory<RetailDbContext> factory, JournalService journal, EInvoiceService einvoices,
                                      Func<DateTime>? clock = null)
{
    private readonly SemaphoreSlim _running = new(1, 1);
    private readonly Func<DateTime> _now = clock ?? (() => DateTime.Now);

    /// <summary>The name of the Windows service in Settings.OfficeSyncComputer: OfficeService:Name, else "SERVICE-" + the computer
    /// name. It never equals the name of a computer, so the Access of that computer does not run the tasks a second time.</summary>
    public static string ServiceName(string? configured, string machine)
    {
        string name = string.IsNullOrWhiteSpace(configured) ? "SERVICE-" + machine : configured.Trim();
        return name.Length <= 64 ? name : name[..64];
    }

    /// <summary>modOfficeSync.IsOfficeSyncComputer: the settings name this runner (not case sensitive).</summary>
    public static bool IsNamed(string? named, string runner) =>
        !string.IsNullOrWhiteSpace(named) && string.Equals(named.Trim(), runner, StringComparison.OrdinalIgnoreCase);

    /// <summary>The automatic run is due for this runner (OfficeSyncTick): it is the one named and the minutes have passed.</summary>
    public async Task<bool> IsDueAsync(string runner)
    {
        await using var db = await factory.CreateDbContextAsync();
        var s = await SettingsReader.ReadAsync(db);
        return IsNamed(s.OfficeSyncComputer, runner) && OfficeSyncRules.IsDue(s.OfficeSyncLastAt, s.OfficeSyncMinutes ?? 0, _now());
    }

    private static string JoinProblem(string first, string second) =>
        first.Length == 0 ? second : second.Length == 0 ? first : first + " - " + second;

    /// <summary>modOfficeSync.RunOfficeSync. runKind AUTO / MANUAL; computer = the name written in OfficeSyncLog.</summary>
    public async Task<OfficeSyncRun> RunAsync(string runKind, string computer, ActingUser user)
    {
        if (!await _running.WaitAsync(0))
        {
            return new OfficeSyncRun(0, 0, 0, 0, 0, 0, "مهام المكتب تعمل الآن.", "");
        }
        int added = 0, updated = 0, removed = 0, sent = 0, notSent = 0, done = 0;
        string problem = "";
        try
        {
            try
            {
                var j = await journal.SyncAsync(user);
                (added, updated, removed, problem) = (j.Added, j.Updated, j.Removed, j.Message);
                bool enabled;
                await using (var db = await factory.CreateDbContextAsync())
                {
                    enabled = (await SettingsReader.ReadAsync(db)).EInvoiceEnabled;
                }
                if (enabled)
                {
                    string msg = await einvoices.ReadyProblemAsync();
                    if (msg.Length > 0)
                    {
                        if (await PendingCountAsync() > 0)
                        {
                            problem = JoinProblem(problem, msg);
                        }
                    }
                    else
                    {
                        (sent, notSent, msg) = await einvoices.SendPendingAsync();
                        problem = JoinProblem(problem, msg);
                        (done, _, msg) = await einvoices.RefreshSubmittedAsync();
                        problem = JoinProblem(problem, msg);
                    }
                }
            }
            catch (Exception e) when (e is not OutOfMemoryException)
            {
                problem = JoinProblem(problem, "تعذّر إكمال مهام المكتب: " + e.Message);
            }
            string summary = "القيود: " + added + " جديد، " + updated + " معدّل، " + removed + " محذوف. " +
                             "الفواتير الإلكترونية: أُرسلت " + sent + "، لم تُرسل " + notSent + "، تمت متابعة " + done + ".";
            await WriteLogAsync(runKind, computer, user, added, updated, removed, sent, notSent, done, problem);
            return new OfficeSyncRun(added, updated, removed, sent, notSent, done, problem, summary);
        }
        finally
        {
            _running.Release();
        }
    }

    /// <summary>modOfficeSync.PendingEInvoiceCount: the documents waiting for the platform.</summary>
    public async Task<int> PendingCountAsync()
    {
        await using var db = await factory.CreateDbContextAsync();
        if (!(await SettingsReader.ReadAsync(db)).EInvoiceEnabled)
        {
            return 0;
        }
        return await db.SalesInvoices.CountAsync(h => h.ZatcaStatus == "PENDING") +
               await db.SalesReturns.CountAsync(r => r.ZatcaStatus == "PENDING");
    }

    private async Task WriteLogAsync(string runKind, string computer, ActingUser user, int added, int updated, int removed, int sent,
                                     int notSent, int done, string problem)
    {
        await using var db = await factory.CreateDbContextAsync();
        DateTime stamp = _now();
        string? text = problem.Length == 0 ? null : problem.Length <= 255 ? problem : problem[..255];
        db.OfficeSyncLog.Add(new OfficeSyncLog
        {
            RunAt = stamp,
            ComputerName = computer.Length <= 64 ? computer : computer[..64],
            RunKind = runKind,
            JournalAdded = added,
            JournalUpdated = updated,
            JournalRemoved = removed,
            EInvoicesSent = sent,
            EInvoicesFailed = notSent,
            EInvoicesChecked = done,
            Problem = text,
            EmployeeID = user.EmployeeID == 0 ? null : user.EmployeeID,
        });
        await db.SaveChangesAsync();
        await db.Settings.Where(s => s.SettingID == 1).ExecuteUpdateAsync(u => u
            .SetProperty(s => s.OfficeSyncLastAt, stamp).SetProperty(s => s.OfficeSyncLastResult, text));
    }
}
