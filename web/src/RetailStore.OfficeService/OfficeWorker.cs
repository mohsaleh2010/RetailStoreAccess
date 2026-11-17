using RetailStore.Data.Services;

namespace RetailStore.OfficeService;

/// <summary>modOfficeSync.OfficeSyncTick without Access: every CheckSeconds the service looks whether it is the runner named in
/// Settings.OfficeSyncComputer and whether OfficeSyncMinutes have passed since the last run (OfficeSyncRules.IsDue), then runs the
/// office tasks (OfficeSyncService.RunAsync "AUTO"). Never stops on a problem: it is in OfficeSyncLog and the event log.</summary>
public sealed class OfficeWorker(OfficeSyncService office, IConfiguration config, ILogger<OfficeWorker> log) : BackgroundService
{
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        string name = OfficeSyncService.ServiceName(config["OfficeService:Name"], Environment.MachineName);
        int seconds = Math.Clamp(config.GetValue("OfficeService:CheckSeconds", 60), 15, 3600);
        log.LogInformation("Office tasks service {Name}: it runs when Settings.OfficeSyncComputer = {Name}.", name, name);
        var user = new ActingUser(0, name);
        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                if (await office.IsDueAsync(name))
                {
                    var run = await office.RunAsync("AUTO", name, user);
                    if (run.Problem.Length > 0)
                    {
                        log.LogWarning("Office tasks: {Summary} {Problem}", run.Summary, run.Problem);
                    }
                    else
                    {
                        log.LogInformation("Office tasks: {Summary}", run.Summary);
                    }
                }
            }
            catch (Exception e) when (e is not OperationCanceledException)
            {
                log.LogError(e, "Office tasks: the database could not be read.");
            }
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(seconds), stoppingToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
        }
    }
}
