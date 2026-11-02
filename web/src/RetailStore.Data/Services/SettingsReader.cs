using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Domain.Common;

namespace RetailStore.Data.Services;

public static class SettingsReader
{
    /// <summary>The settings row (Settings.SettingID = 1).</summary>
    public static async Task<Settings> ReadAsync(RetailDbContext db) =>
        await db.Settings.AsNoTracking().FirstOrDefaultAsync(s => s.SettingID == 1) ?? new Settings();

    /// <summary>modCountry.AppCountry.</summary>
    public static async Task<string> CountryAsync(RetailDbContext db) => CountryRules.Normalize((await ReadAsync(db)).CountryCode);
}
