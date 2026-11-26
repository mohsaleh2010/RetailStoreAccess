using MudBlazor;

namespace RetailStore.UI.Components.Shared;

/// <summary>The caption (Arabic, shown through L.T), color and icon of a document status for StatusChip.</summary>
public sealed record StatusStyle(string Caption, Color Color, string Icon);

public static class StatusStyles
{
    private static readonly Dictionary<string, StatusStyle> Known = new(StringComparer.OrdinalIgnoreCase)
    {
        ["DRAFT"] = new("مسودة", Color.Default, Icons.Material.Outlined.EditNote),
        ["POSTED"] = new("مُرحَّل", Color.Success, Icons.Material.Outlined.TaskAlt),
        ["APPROVED"] = new("معتمد", Color.Info, Icons.Material.Outlined.Verified),
        ["PREPARING"] = new("قيد التجهيز", Color.Info, Icons.Material.Outlined.Inventory2),
        ["PARTIAL"] = new("جزئي", Color.Warning, Icons.Material.Outlined.Timelapse),
        ["PARTIAL_PAID"] = new("مدفوع جزئيًا", Color.Warning, Icons.Material.Outlined.Timelapse),
        ["PAID"] = new("مدفوع", Color.Success, Icons.Material.Outlined.Paid),
        ["UNPAID"] = new("غير مدفوع", Color.Error, Icons.Material.Outlined.MoneyOff),
        ["DELIVERED"] = new("تم التسليم", Color.Success, Icons.Material.Outlined.LocalShipping),
        ["RECEIVED"] = new("مستلم", Color.Success, Icons.Material.Outlined.MoveToInbox),
        ["INVOICED"] = new("مفوتر", Color.Success, Icons.Material.Outlined.ReceiptLong),
        ["IN_TRANSIT"] = new("في الطريق", Color.Warning, Icons.Material.Outlined.LocalShipping),
        ["INTRANSIT"] = new("في الطريق", Color.Warning, Icons.Material.Outlined.LocalShipping),
        ["PENDING"] = new("بانتظار الإرسال", Color.Warning, Icons.Material.Outlined.Schedule),
        ["OPEN"] = new("مفتوح", Color.Info, Icons.Material.Outlined.LockOpen),
        ["CLOSED"] = new("مغلق", Color.Default, Icons.Material.Outlined.Lock),
        ["EXPIRED"] = new("منتهي", Color.Warning, Icons.Material.Outlined.EventBusy),
        ["REJECTED"] = new("مرفوض", Color.Error, Icons.Material.Outlined.Block),
        ["CANCELLED"] = new("ملغى", Color.Error, Icons.Material.Outlined.Cancel),
    };

    /// <summary>The style of a status code (DRAFT, POSTED, CANCELLED, PARTIAL_PAID ...); an unknown code shows itself.</summary>
    public static StatusStyle For(string? status) =>
        Known.TryGetValue((status ?? "").Trim(), out var style)
            ? style
            : new StatusStyle(status ?? "", Color.Default, Icons.Material.Outlined.Circle);

    public static IReadOnlyCollection<string> Codes => Known.Keys;
}
