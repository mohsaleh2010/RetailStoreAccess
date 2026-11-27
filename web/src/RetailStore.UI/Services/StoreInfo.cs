using Microsoft.EntityFrameworkCore;
using RetailStore.Data;
using RetailStore.Data.Services;
using RetailStore.Domain.Common;

namespace RetailStore.UI.Services;

/// <summary>Small reads of the settings for the layout.</summary>
public interface IStoreInfo : ICurrencySource
{
    Task<string> StoreNameAsync(bool english);
}

public sealed class StoreInfo(IDbContextFactory<RetailDbContext> factory) : IStoreInfo
{
    public async Task<string> StoreNameAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var s = await SettingsReader.ReadAsync(db);
        return english && !string.IsNullOrWhiteSpace(s.StoreNameEn) ? s.StoreNameEn! : (s.StoreName ?? "");
    }

    public async Task<string> CurrencyCodeAsync()
    {
        await using var db = await factory.CreateDbContextAsync();
        var s = await SettingsReader.ReadAsync(db);
        if (!string.IsNullOrWhiteSpace(s.CurrencyCode))
        {
            return s.CurrencyCode.Trim().ToUpperInvariant();
        }
        return CountryRules.Normalize(s.CountryCode) == "EG" ? "EGP" : "SAR";
    }
}
