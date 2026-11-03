using Microsoft.Extensions.DependencyInjection;
using RetailStore.RepApp.Client;
using RetailStore.RepApp.Offline;

namespace RetailStore.RepApp;

public static class RepAppServices
{
    /// <summary>The services of the rep app (docs/78): the SQLite file of the phone at databasePath, one HttpClient, one session.</summary>
    public static IServiceCollection AddRepApp(this IServiceCollection services, string databasePath)
    {
        services.AddSingleton(_ => new LocalStore(databasePath));
        services.AddSingleton(_ => new RepApiClient(new HttpClient { Timeout = TimeSpan.FromSeconds(30) }));
        services.AddSingleton<RepSession>();
        return services;
    }
}
