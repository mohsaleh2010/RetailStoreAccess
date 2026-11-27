namespace RetailStore.UI.Services;

/// <summary>The currency of the program for the screens (Settings.CurrencyCode, else the operating country: SA riyal,
/// EG pound), read once per session. MoneyText shows its word next to an amount; nothing is converted here.</summary>
public sealed class UiCurrency(ICurrencySource store, UiLanguage lang)
{
    private Task? loading;

    /// <summary>SAR, EGP or the code of the settings; "" until loaded.</summary>
    public string Code { get; private set; } = "";

    public Task EnsureAsync() => loading ??= LoadAsync();

    private async Task LoadAsync() => Code = await store.CurrencyCodeAsync();

    /// <summary>The host already knows the currency (the snapshot of the Windows POS after its login): it replaces what
    /// was read, and nothing is read any more.</summary>
    public void Use(string? code)
    {
        Code = (code ?? "").Trim().ToUpperInvariant();
        loading = Task.CompletedTask;
    }

    /// <summary>The word shown after an amount, in the interface language.</summary>
    public string Word => Code switch
    {
        "SAR" => lang.T("ريال"),
        "EGP" => lang.T("جنيه"),
        _ => Code,
    };
}
