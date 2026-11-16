using Microsoft.EntityFrameworkCore;

namespace RetailStore.Data;

public static class DbSetup
{
    /// <summary>A SQLite development or test database: created from the model with the seed rows and the levels of
    /// the account tree when it is new.
    /// A SQL Server database is never created here: it comes from dist/sql/RetailStore_Schema.sql (docs/67).</summary>
    public static void EnsureSqliteDatabase(RetailDbContext db)
    {
        if (db.Database.ProviderName != "Microsoft.EntityFrameworkCore.Sqlite")
        {
            throw new InvalidOperationException("EnsureSqliteDatabase is for SQLite only.");
        }
        if (db.Database.EnsureCreated())
        {
            SeedData.Add(db);
            db.SaveChanges();
            // the levels of the account tree, as BuildSchema does (modAccounts.RebuildAccountTree)
            Journal.JournalService.EnsureAccountsAsync(db).GetAwaiter().GetResult();
            db.ChangeTracker.Clear();
        }
    }
}
