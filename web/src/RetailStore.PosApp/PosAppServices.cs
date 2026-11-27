using Microsoft.Extensions.DependencyInjection;
using RetailStore.PosApp.Offline;
using RetailStore.UI;
using RetailStore.UI.Services;

namespace RetailStore.PosApp;

public static class PosAppServices
{
    /// <summary>The services of the POS (docs/83): the SQLite file of the PC at databasePath and one session, and the
    /// services of the shared components (RetailStore.UI.Kit); the host calls AddMudServices after this.</summary>
    public static IServiceCollection AddPosApp(this IServiceCollection services, string databasePath)
    {
        services.AddSingleton(_ => new PosStore(databasePath));
        services.AddSingleton(sp => new PosSession(sp.GetRequiredService<PosStore>(), sp.GetRequiredService<IPosPlatform>()));
        services.AddErpUiKit();
        services.AddSingleton<ICurrencySource, PosCurrencySource>();
        return services;
    }
}

/// <summary>The currency of the shared components (MoneyText): that of the snapshot of the user, "" before the login
/// (Routes gives UiCurrency the snapshot's currency each time the session changes).</summary>
public sealed class PosCurrencySource(PosSession session) : ICurrencySource
{
    public Task<string> CurrencyCodeAsync() => Task.FromResult(session.Snapshot?.Settings.Currency ?? "");
}
