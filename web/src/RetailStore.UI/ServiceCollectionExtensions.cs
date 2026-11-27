using Microsoft.Extensions.DependencyInjection;
using RetailStore.Data.Journal;
using RetailStore.Data.Services;
using RetailStore.UI.Services;

namespace RetailStore.UI;

public static class ServiceCollectionExtensions
{
    /// <summary>The services of the screens (the database itself is registered by the host: web or MAUI).</summary>
    public static IServiceCollection AddRetailStoreUi(this IServiceCollection services)
    {
        services.AddErpUiKit();   // UiLanguage, UiTheme, UiCurrency, ShortcutService, ErpMudLocalizer (RetailStore.UI.Kit)
        services.AddScoped<IStoreInfo, StoreInfo>();
        services.AddScoped<ICurrencySource>(sp => sp.GetRequiredService<IStoreInfo>());   // the currency of the settings
        services.AddScoped<PermissionService>();
        services.AddScoped<DataScreenService>();
        services.AddScoped<SalesService>();
        services.AddScoped<PosService>();
        services.AddScoped<PurchaseService>();
        services.AddScoped<StockService>();
        services.AddScoped<SecurityService>();
        services.AddScoped<IJournalSourceReader, ServerViewJournalReader>();
        services.AddScoped<JournalService>();
        services.AddScoped<TreasuryService>();
        services.AddScoped<BankReconService>();
        services.AddScoped<ManualEntryService>();
        services.AddScoped<RepDocService>();
        services.AddScoped<AccountingService>();
        services.AddScoped<DashboardService>();
        services.AddScoped<IndicatorsService>();
        services.AddScoped<ReportService>();
        services.AddScoped<HomeReports>();
        return services;
    }
}
