using Microsoft.EntityFrameworkCore;
using RetailStore.Data;
using RetailStore.Data.Services;

namespace RetailStore.UI.Services;

/// <summary>Small reads of the settings for the layout.</summary>
public interface IStoreInfo
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
}
