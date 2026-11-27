namespace RetailStore.UI.Services;

/// <summary>Where the currency of the program comes from: the settings of the database on the web and the Windows POS
/// (IStoreInfo), the snapshot of the API in the rep app. The kit never reads a database itself.</summary>
public interface ICurrencySource
{
    /// <summary>The currency of the program: Settings.CurrencyCode, else that of the operating country (SAR / EGP).</summary>
    Task<string> CurrencyCodeAsync();
}

/// <summary>A currency already known (a preview host, a test, an app that read it from its snapshot).</summary>
public sealed class FixedCurrencySource(string code) : ICurrencySource
{
    public Task<string> CurrencyCodeAsync() => Task.FromResult(code);
}
