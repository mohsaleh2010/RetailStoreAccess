using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.Extensions.DependencyInjection;

namespace RetailStore.Tests;

/// <summary>The hosting of the site (docs/88): the health check of an installation and the kept keys.</summary>
public class HostingTests
{
    [Fact]
    public async Task HealthAnswersOkWithTheDatabase()
    {
        using var host = new ApiHost();
        var r = await host.CreateClient().GetAsync("/health");
        Assert.Equal(HttpStatusCode.OK, r.StatusCode);
        var body = await r.Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal("ok", body.GetProperty("status").GetString());
        Assert.Single(body.EnumerateObject());
    }

    [Fact]
    public async Task HealthSaysDatabaseWhenTheSettingsRowIsMissing()
    {
        using var host = new ApiHost();
        using (var db = host.Db.CreateDbContext())
        {
            db.Settings.RemoveRange(db.Settings);
            await db.SaveChangesAsync();
        }
        var r = await host.CreateClient().GetAsync("/health");
        Assert.Equal(HttpStatusCode.ServiceUnavailable, r.StatusCode);
        Assert.Equal("database", (await r.Content.ReadFromJsonAsync<JsonElement>()).GetProperty("status").GetString());
    }

    [Fact]
    public void ProtectedDataSurvivesARestartOfTheSite()
    {
        string protectedText;
        using (var first = new ApiHost())
        {
            _ = first.CreateClient();
            protectedText = first.Services.GetRequiredService<IDataProtectionProvider>().CreateProtector("docs88").Protect("cookie");
        }
        using var second = new ApiHost();
        _ = second.CreateClient();
        var protector = second.Services.GetRequiredService<IDataProtectionProvider>().CreateProtector("docs88");
        Assert.Equal("cookie", protector.Unprotect(protectedText));
    }
}
