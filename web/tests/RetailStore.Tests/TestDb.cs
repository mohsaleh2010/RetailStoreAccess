using Microsoft.Data.Sqlite;
using Microsoft.EntityFrameworkCore;
using RetailStore.Data;
using RetailStore.Data.Services;

namespace RetailStore.Tests;

/// <summary>A new SQLite database in memory, created from the generated model with the seed rows (DbSetup),
/// kept while the test runs.</summary>
public sealed class TestDb : IDbContextFactory<RetailDbContext>, IDisposable
{
    private readonly SqliteConnection _connection = new("Data Source=:memory:;Foreign Keys=True");

    public TestDb()
    {
        _connection.Open();
        using var db = CreateDbContext();
        DbSetup.EnsureSqliteDatabase(db);
    }

    public RetailDbContext CreateDbContext() =>
        new(new DbContextOptionsBuilder<RetailDbContext>().UseSqlite(_connection).Options);

    public PermissionService Permissions => new(this);
    public DataScreenService Screens => new(this, Permissions);
    public SecurityService Security => new(this);

    public static readonly ActingUser Admin = new(1, "TEST");

    public void Dispose() => _connection.Dispose();
}
