using Microsoft.Extensions.DependencyInjection;
using RetailStore.PosApp.Offline;

namespace RetailStore.PosApp;

public static class PosAppServices
{
    /// <summary>The services of the POS (docs/83): the SQLite file of the PC at databasePath and one session.</summary>
    public static IServiceCollection AddPosApp(this IServiceCollection services, string databasePath)
    {
        services.AddSingleton(_ => new PosStore(databasePath));
        services.AddSingleton(sp => new PosSession(sp.GetRequiredService<PosStore>(), sp.GetRequiredService<IPosPlatform>()));
        return services;
    }
}
