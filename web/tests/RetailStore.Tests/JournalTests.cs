using System.Globalization;
using System.Text.Json;
using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Journal;
using RetailStore.Domain;
using RetailStore.Domain.Journal;

namespace RetailStore.Tests;

public class JournalPlannerTests
{
    private static readonly Dictionary<string, string> Kinds = new() { ["SALE"] = "فاتورة بيع", ["EXPENSE"] = "مصروف" };

    private static JournalSourceLine L(string type, int id, int order, int account, decimal debit, decimal credit,
                                       DateTime? when = null, int center = 0, string? party = "عميل") =>
        new(type, id, type + "-" + id, when ?? new DateTime(2026, 10, 5), party, order, account, debit, credit, "x", center);

    private static List<JournalSourceLine> Sale(int id, decimal amount, DateTime? when = null) =>
        [L("SALE", id, 1, 110001, amount, 0, when), L("SALE", id, 2, 4100, 0, amount, when)];

    [Fact]
    public void Signature_is_the_one_of_SyncJournal()
    {
        var lines = Sale(1, 100);
        // 110001 x (200 + 0) + 4100 x (0 + 100)
        Assert.Equal(110001m * 200 + 4100m * 100, JournalPlanner.Signature(lines));
        var centred = new List<JournalSourceLine> { L("SALE", 1, 1, 110001, 100, 0, center: 3), L("SALE", 1, 2, 4100, 0, 100, center: 3) };
        Assert.Equal(110001m * 200 + 4100m * 100 + 3m * 300 * 7, JournalPlanner.Signature(centred));
    }

    [Fact]
    public void New_changed_deleted_unbalanced_and_closed()
    {
        var plan = JournalPlanner.Plan([Sale(1, 100), [L("EXPENSE", 9, 1, 530001, 5, 0), L("EXPENSE", 9, 2, 110001, 0, 4)]],
                                       [], Kinds, null);
        var h = Assert.Single(plan.Added);
        Assert.Equal("فاتورة بيع SALE-1 - عميل", h.Description);
        Assert.Single(plan.Unbalanced);
        var existing = new List<ExistingEntry> { new(7, "SALE", 1, h.Signature, h.EntryDate, h.TotalDebit), new(8, "SALE", 2, 1, h.EntryDate, 1) };
        var again = JournalPlanner.Plan([Sale(1, 100)], existing, Kinds, null);
        Assert.Empty(again.Added);
        Assert.Empty(again.Updated);
        Assert.Equal([8], again.Removed);
        var changed = JournalPlanner.Plan([Sale(1, 120)], existing.Take(1), Kinds, null);
        Assert.Equal(7, Assert.Single(changed.Updated).EntryID);
        var closed = JournalPlanner.Plan([Sale(1, 120)], existing.Take(1), Kinds, new DateTime(2026, 10, 31));
        Assert.Empty(closed.Updated);
        Assert.Single(closed.Locked);
        Assert.True(JournalPlanner.InClosedPeriod("SALE", new DateTime(2026, 10, 31, 23, 0, 0), new DateTime(2026, 10, 31)));
        Assert.False(JournalPlanner.InClosedPeriod("YEAR_CLOSE", new DateTime(2026, 1, 1), new DateTime(2026, 10, 31)));
    }

    [Fact]
    public void The_account_tree_has_levels_and_keys()
    {
        var (places, problems) = AccountTree.Build(new Dictionary<int, int> { [1] = 0, [11] = 1, [1100] = 11, [110001] = 1100, [5] = 6, [6] = 5 });
        Assert.Equal(4, places[110001].Level);
        Assert.Equal("0000000001000000001100000011000000110001", places[110001].TreeKey);
        Assert.Equal([1, 11, 1100, 110001, null], places[110001].LevelCodes);
        Assert.Equal([5, 6], problems.Order());
    }
}

public class JournalServiceTests
{
    [Fact]
    public async Task Sync_numbers_new_entries_in_date_order_and_a_second_sync_changes_nothing()
    {
        using var t = new TestDb();
        var lines = new Dictionary<string, List<JournalSourceLine>>
        {
            ["qryJournalSale"] =
            [
                new("SALE", 2, "INV-2", new DateTime(2026, 10, 9), "عميل نقدي", 1, 110002, 50, 0, "نقدًا", 0),
                new("SALE", 2, "INV-2", new DateTime(2026, 10, 9), "عميل نقدي", 2, 4100, 0, 50, "المبيعات", 0),
                new("SALE", 1, "INV-1", new DateTime(2026, 10, 1), "عميل نقدي", 1, 110002, 30, 0, "نقدًا", 0),
                new("SALE", 1, "INV-1", new DateTime(2026, 10, 1), "عميل نقدي", 2, 4100, 0, 30, "المبيعات", 0),
            ],
        };
        var service = new JournalService(t, new FixedJournalReader(lines));
        var first = await service.SyncAsync(TestDb.Admin);
        Assert.Equal((2, 0, 0, ""), (first.Added, first.Updated, first.Removed, first.Message));
        await using (var db = t.CreateDbContext())
        {
            Assert.Equal("JV-000001", (await db.JournalEntries.SingleAsync(e => e.SourceID == 1)).EntryNumber);
            Assert.Equal("JV-000002", (await db.JournalEntries.SingleAsync(e => e.SourceID == 2)).EntryNumber);
            Assert.Equal(4, await db.JournalLines.CountAsync());
            Assert.Equal(4, (await db.Accounts.SingleAsync(a => a.AccountCode == 110002)).AccountLevel);
        }
        var second = await service.SyncAsync(TestDb.Admin);
        Assert.Equal((0, 0, 0), (second.Added, second.Updated, second.Removed));
        lines["qryJournalSale"].RemoveRange(0, 2);
        var third = await service.SyncAsync(TestDb.Admin);
        Assert.Equal((0, 0, 1), (third.Added, third.Updated, third.Removed));
    }
}

/// <summary>The C# planner against the Python mirror of SyncJournal (tools/sim.py) on the fixture business of the
/// Python tests: tests/test_web.py writes the source lines and the expected entries to a file and sets
/// RETAIL_JOURNAL_VECTORS; without it the test has nothing to compare.</summary>
public class JournalVectorTests
{
    [Fact]
    public void Planner_gives_the_entries_of_the_Python_mirror()
    {
        string? path = Environment.GetEnvironmentVariable("RETAIL_JOURNAL_VECTORS");
        if (string.IsNullOrEmpty(path))
        {
            return;
        }
        using var doc = JsonDocument.Parse(File.ReadAllText(path));
        var root = doc.RootElement;
        var kinds = root.GetProperty("kinds").EnumerateObject().ToDictionary(p => p.Name, p => p.Value.GetString()!);
        var sources = AppData.JournalSourceViews.Select(v => root.GetProperty("lines").TryGetProperty(v, out var arr)
            ? arr.EnumerateArray().Select(Line).ToList() : []).ToList();
        var plan = JournalPlanner.Plan(sources, [], kinds, null);
        var expected = root.GetProperty("entries").EnumerateArray().ToList();
        Assert.Equal(expected.Count, plan.Added.Count);
        var byKey = plan.Added.ToDictionary(h => (h.SourceType, h.SourceID));
        foreach (var e in expected)
        {
            var h = byKey[(e.GetProperty("SourceType").GetString()!, e.GetProperty("SourceID").GetInt32())];
            Assert.Equal(e.GetProperty("Description").GetString(), h.Description);
            Assert.Equal(e.GetProperty("TotalDebit").GetDecimal(), h.TotalDebit);
            Assert.Equal(e.GetProperty("Signature").GetDecimal(), h.Signature);
            Assert.Equal(e.GetProperty("LineCount").GetInt32(), h.Lines.Count);
            Assert.Equal(DateTime.Parse(e.GetProperty("EntryDate").GetString()!, CultureInfo.InvariantCulture), h.EntryDate);
        }
        Assert.Equal(root.GetProperty("unbalanced").GetInt32(), plan.Unbalanced.Count);
    }

    private static JournalSourceLine Line(JsonElement a) => new(
        a[0].GetString()!, a[1].GetInt32(), a[2].ValueKind == JsonValueKind.Null ? null : a[2].ToString(),
        a[3].ValueKind == JsonValueKind.Null ? null : DateTime.Parse(a[3].GetString()!, CultureInfo.InvariantCulture),
        a[4].ValueKind == JsonValueKind.Null ? null : a[4].GetString(), a[5].GetInt32(), a[6].GetInt32(),
        a[7].ValueKind == JsonValueKind.Null ? 0 : a[7].GetDecimal(), a[8].ValueKind == JsonValueKind.Null ? 0 : a[8].GetDecimal(),
        a[9].ValueKind == JsonValueKind.Null ? null : a[9].GetString(), a[10].ValueKind == JsonValueKind.Null ? 0 : a[10].GetInt32());
}
