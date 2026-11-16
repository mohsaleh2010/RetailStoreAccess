using System.Net;
using System.Net.Http.Json;
using RetailStore.Data.Entities;
using RetailStore.Data.Services;
using RetailStore.Domain.Common;

namespace RetailStore.Tests;

/// <summary>Access kept for the office (plan 6H, docs/84): the web reads the state of the office tasks of Access (modOfficeSync)
/// and warns who may run them when they stopped, failed, or nobody runs them while e-invoices wait.</summary>
public class OfficeSyncTests
{
    private static readonly DateTime T0 = new(2026, 11, 16, 10, 0, 0);

    [Fact]
    public void The_rules_follow_modOfficeSync()
    {
        // the same cases as TestOfficeSync in Access
        Assert.True(OfficeSyncRules.IsDue(null, 15, T0));
        Assert.False(OfficeSyncRules.IsDue(null, 0, T0));
        Assert.False(OfficeSyncRules.IsDue(T0.AddMinutes(-14), 15, T0));
        Assert.True(OfficeSyncRules.IsDue(T0.AddMinutes(-15), 15, T0));
        Assert.False(OfficeSyncRules.IsStale(T0.AddMinutes(-40), 15, T0));
        Assert.True(OfficeSyncRules.IsStale(T0.AddMinutes(-41), 15, T0));
        Assert.False(OfficeSyncRules.IsStale(null, 0, T0));
        // DateDiff("n"): minute boundaries crossed
        Assert.Equal(1, OfficeSyncRules.WholeMinutes(new DateTime(2026, 1, 1, 9, 59, 59), new DateTime(2026, 1, 1, 10, 0, 1)));
    }

    [Fact]
    public async Task The_dashboard_reads_the_office_tasks_and_warns()
    {
        using var tdb = new TestDb();
        var service = new DashboardService(tdb, tdb.Permissions);
        var s = await service.OfficeSyncAsync(T0);
        Assert.Equal((null, 15, false, false), (s.Computer, s.Minutes, s.Stale, s.NeedsAttention));     // nothing set, nothing waits

        using (var db = tdb.CreateDbContext())
        {
            var st = db.Settings.Single(x => x.SettingID == 1);
            st.OfficeSyncComputer = " OFFICE-PC ";
            st.OfficeSyncLastAt = T0.AddMinutes(-41);
            db.SaveChanges();
        }
        s = await service.OfficeSyncAsync(T0);
        Assert.Equal(("OFFICE-PC", true, true), (s.Computer, s.Stale, s.NeedsAttention));

        using (var db = tdb.CreateDbContext())
        {
            var st = db.Settings.Single(x => x.SettingID == 1);
            st.OfficeSyncLastAt = T0.AddMinutes(-5);
            st.OfficeSyncLastResult = "لا يرد الخادم.";
            db.SaveChanges();
        }
        s = await service.OfficeSyncAsync(T0);
        Assert.Equal((false, true), (s.Stale, s.NeedsAttention));                                       // the last run failed

        using (var db = tdb.CreateDbContext())
        {
            var st = db.Settings.Single(x => x.SettingID == 1);
            st.OfficeSyncLastResult = null;
            st.OfficeSyncComputer = null;
            st.EInvoiceEnabled = true;
            var c = db.Customers.First();
            db.SalesInvoices.Add(new SalesInvoices { InvoiceNumber = "WEB-1", InvoiceDate = T0, CustomerID = c.CustomerID, EmployeeID = 1,
                                                     PaymentMethodID = 1, ZatcaStatus = "PENDING" });
            db.SaveChanges();
        }
        s = await service.OfficeSyncAsync(T0);
        Assert.Equal((1, false, true), (s.PendingEInvoices, s.Stale, s.NeedsAttention));             // e-invoices wait, nobody sends
    }

    [Fact]
    public async Task The_api_gives_the_office_tasks_to_who_may_run_them()
    {
        await using var host = new ApiHost();
        var client = await host.AdminAsync();
        var r = await client.GetAsync("/api/v1/dashboard/office-sync");
        Assert.Equal(HttpStatusCode.OK, r.StatusCode);
        var body = await r.Content.ReadFromJsonAsync<OfficeSyncStatus>();
        Assert.Equal(15, body!.Minutes);
    }
}
