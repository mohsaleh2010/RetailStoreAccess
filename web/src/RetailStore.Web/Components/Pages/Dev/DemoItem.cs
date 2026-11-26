namespace RetailStore.Web.Components.Pages.Dev;

/// <summary>A row of the components gallery (/dev/components): sample data of the page itself, never the database.</summary>
public sealed class DemoItem
{
    public string Code { get; set; } = "";
    public string Name { get; set; } = "";
    public string Unit { get; set; } = "";
    public decimal Price { get; set; }
    public decimal Balance { get; set; }
    public string Status { get; set; } = "DRAFT";
    public DateTime Date { get; set; }

    /// <summary>The message of the gallery dialog when the price is below zero (as a service would answer).</summary>
    public const string NegativePrice = "السعر لا يكون سالبًا";

    public static List<DemoItem> Samples() =>
    [
        new() { Code = "1001", Name = "سكر أبيض 1 كجم", Unit = "كيس", Price = 6.5m, Balance = 240, Status = "POSTED", Date = new(2026, 11, 2) },
        new() { Code = "1002", Name = "أرز بسمتي 5 كجم", Unit = "كيس", Price = 42.75m, Balance = 36, Status = "POSTED", Date = new(2026, 11, 9) },
        new() { Code = "1003", Name = "زيت دوار الشمس 1.5 لتر", Unit = "حبة", Price = 18.95m, Balance = -4, Status = "DRAFT", Date = new(2026, 11, 15) },
        new() { Code = "2001", Name = "شاي أسود 100 كيس", Unit = "علبة", Price = 14m, Balance = 12.5m, Status = "CANCELLED", Date = new(2026, 11, 20) },
        new() { Code = "2002", Name = "حليب طويل الأجل 1 لتر", Unit = "كرتون", Price = 64.2m, Balance = 0, Status = "DRAFT", Date = new(2026, 11, 24) },
    ];
}
