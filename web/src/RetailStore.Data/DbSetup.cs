using Microsoft.EntityFrameworkCore;

namespace RetailStore.Data;

public static class DbSetup
{
    /// <summary>A SQLite development or test database: created from the model with the seed rows when it is new.
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
            db.ChangeTracker.Clear();
        }
    }
}
