using Microsoft.Extensions.DependencyInjection;
using RetailStore.UI.Services;

namespace RetailStore.UI;

public static class KitServiceCollectionExtensions
{
    /// <summary>The services of the shared components: language, light / dark mode, currency, keyboard shortcuts and
    /// the Arabic texts of MudBlazor. The host registers an ICurrencySource and calls AddMudServices after this.</summary>
    public static IServiceCollection AddErpUiKit(this IServiceCollection services)
    {
        services.AddScoped<UiLanguage>();
        services.AddScoped<UiTheme>();
        services.AddScoped<UiCurrency>();
        services.AddScoped<ShortcutService>();
        services.AddTransient<MudBlazor.MudLocalizer, ErpMudLocalizer>();   // before AddMudServices of the host
        return services;
    }
}
