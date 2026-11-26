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
        services.AddScoped<UiLanguage>();
        services.AddScoped<UiTheme>();
        services.AddScoped<UiCurrency>();
        services.AddScoped<ShortcutService>();
        services.AddTransient<MudBlazor.MudLocalizer, ErpMudLocalizer>();   // before AddMudServices of the host
        services.AddScoped<IStoreInfo, StoreInfo>();
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
        return services;
    }
}
