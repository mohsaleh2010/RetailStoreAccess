using Microsoft.Extensions.DependencyInjection;
using Microsoft.JSInterop;
using MudBlazor;
using RetailStore.PosApp;
using RetailStore.RepApp;
using RetailStore.UI.Services;

namespace RetailStore.Tests;

/// <summary>Phase 7, batch 7A of the redesign (docs/ui-inventory.md section 17): the theme, the shared components and
/// their services live in RetailStore.UI.Kit, which never reaches the database, and the Windows POS and the rep app get
/// the same services as the web (language, currency, shortcuts, the Arabic texts of MudBlazor).</summary>
public class KitTests
{
    private static string[] ReferencesOf(Type type) =>
        type.Assembly.GetReferencedAssemblies().Select(a => a.Name ?? "").ToArray();

    [Fact]
    public void The_kit_and_the_rep_app_never_reference_the_data_layer()
    {
        Assert.Equal("RetailStore.UI.Kit", typeof(UiLanguage).Assembly.GetName().Name);
        Assert.Equal("RetailStore.UI.Kit", typeof(RetailStore.UI.Theme.ErpTheme).Assembly.GetName().Name);
        Assert.Equal("RetailStore.UI.Kit", typeof(RetailStore.UI.Components.Shared.PageHeader).Assembly.GetName().Name);
        Assert.DoesNotContain("RetailStore.Data", ReferencesOf(typeof(UiLanguage)));
        Assert.DoesNotContain("RetailStore.UI", ReferencesOf(typeof(UiLanguage)));
        Assert.DoesNotContain("RetailStore.Data", ReferencesOf(typeof(RetailStore.RepApp.Routes)));
        Assert.DoesNotContain("RetailStore.UI", ReferencesOf(typeof(RetailStore.RepApp.Routes)));
        Assert.DoesNotContain("RetailStore.UI", ReferencesOf(typeof(RetailStore.PosApp.Routes)));
    }

    private static ServiceProvider Build(Action<IServiceCollection> add)
    {
        var services = new ServiceCollection();
        services.AddSingleton<IJSRuntime>(new NoJs());
        add(services);
        return services.BuildServiceProvider();
    }

    [Fact]
    public async Task The_pos_gets_the_services_of_the_shared_components()
    {
        var path = Path.Combine(Path.GetTempPath(), $"kit-pos-{Guid.NewGuid():N}.db");
        await using var sp = Build(s =>
        {
            s.AddSingleton<IPosPlatform>(new BasicPosPlatform());
            s.AddPosApp(path);
        });
        await using var scope = sp.CreateAsyncScope();
        Assert.NotNull(scope.ServiceProvider.GetRequiredService<UiLanguage>());
        Assert.NotNull(scope.ServiceProvider.GetRequiredService<UiTheme>());
        Assert.NotNull(scope.ServiceProvider.GetRequiredService<ShortcutService>());
        Assert.IsType<ErpMudLocalizer>(scope.ServiceProvider.GetRequiredService<MudLocalizer>());
        Assert.IsType<PosCurrencySource>(scope.ServiceProvider.GetRequiredService<ICurrencySource>());
        var currency = scope.ServiceProvider.GetRequiredService<UiCurrency>();
        await currency.EnsureAsync();
        Assert.Equal("", currency.Code);   // before the login: no snapshot, no currency
    }

    [Fact]
    public async Task The_rep_app_gets_them_without_a_currency()
    {
        var path = Path.Combine(Path.GetTempPath(), $"kit-rep-{Guid.NewGuid():N}.db");
        await using var sp = Build(s =>
        {
            s.AddSingleton<IRepPlatform>(new BasicPlatform());
            s.AddRepApp(path);
        });
        await using var scope = sp.CreateAsyncScope();
        Assert.IsType<ErpMudLocalizer>(scope.ServiceProvider.GetRequiredService<MudLocalizer>());
        Assert.IsType<FixedCurrencySource>(scope.ServiceProvider.GetRequiredService<ICurrencySource>());
        var currency = scope.ServiceProvider.GetRequiredService<UiCurrency>();
        await currency.EnsureAsync();
        Assert.Equal("", currency.Code);
        Assert.Equal("", currency.Word);
    }

    [Fact]
    public async Task A_currency_given_by_the_host_replaces_what_was_read()
    {
        var lang = new UiLanguage();
        var currency = new UiCurrency(new FixedCurrencySource("SAR"), lang);
        await currency.EnsureAsync();
        Assert.Equal("SAR", currency.Code);
        Assert.Equal("ريال", currency.Word);
        currency.Use(" egp ");
        await currency.EnsureAsync();
        Assert.Equal("EGP", currency.Code);
        lang.Set("en");
        Assert.Equal("EGP", currency.Code);
        currency.Use(null);
        Assert.Equal("", currency.Code);
    }

    private sealed class NoJs : IJSRuntime
    {
        public ValueTask<TValue> InvokeAsync<TValue>(string identifier, object?[]? args) => ValueTask.FromResult(default(TValue)!);

        public ValueTask<TValue> InvokeAsync<TValue>(string identifier, CancellationToken cancellationToken, object?[]? args) =>
            ValueTask.FromResult(default(TValue)!);
    }
}
