namespace RetailStore.UI.Services;

/// <summary>The currency of the program for the screens (Settings.CurrencyCode, else the operating country: SA riyal,
/// EG pound), read once per session. MoneyText shows its word next to an amount; nothing is converted here.</summary>
public sealed class UiCurrency(IStoreInfo store, UiLanguage lang)
{
    private Task? loading;

    /// <summary>SAR, EGP or the code of the settings; "" until loaded.</summary>
    public string Code { get; private set; } = "";

    public Task EnsureAsync() => loading ??= LoadAsync();

    private async Task LoadAsync() => Code = await store.CurrencyCodeAsync();

    /// <summary>The word shown after an amount, in the interface language.</summary>
    public string Word => Code switch
    {
        "SAR" => lang.T("ريال"),
        "EGP" => lang.T("جنيه"),
        _ => Code,
    };
}
