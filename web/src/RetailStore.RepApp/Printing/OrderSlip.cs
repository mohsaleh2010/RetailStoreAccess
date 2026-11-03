using System.Globalization;
using RetailStore.RepApp.Offline;

namespace RetailStore.RepApp.Printing;

public enum SlipAlign { Start, Center, End }

/// <summary>A line of a printed slip: the phone draws it (Arabic shaped, right to left) and the printer gets a picture.</summary>
public sealed record SlipLine(string Text, SlipAlign Align = SlipAlign.Start, bool Bold = false, bool Large = false);

/// <summary>The customer's copy of an order taken on the phone (docs/78), printed on the rep's Bluetooth printer: the
/// number on the phone until the server gives the SO- number, the customer, the rep, each line, the total. Prices include
/// VAT as on the order lines. T = the language of the app (modLang.Tr).</summary>
public static class OrderSlip
{
    public static List<SlipLine> Build(PendingOrder order, string? serverNumber, string repName, string storeName,
                                       Func<string, string> T)
    {
        var inv = CultureInfo.InvariantCulture;
        var lines = new List<SlipLine>
        {
            new(storeName, SlipAlign.Center, Bold: true, Large: true),
            new(T("طلبية عميل"), SlipAlign.Center, Bold: true),
            new(T("رقم الطلبية") + ": " + (serverNumber ?? order.ClientRef[..8].ToUpperInvariant())),
            new(T("التاريخ") + ": " + order.TakenAt.ToString("yyyy-MM-dd HH:mm", inv)),
            new(T("العميل") + ": " + order.CustomerName, Bold: true),
            new(T("المندوب") + ": " + repName),
            new(Rule),
        };
        foreach (var l in order.Lines)
        {
            lines.Add(new SlipLine(l.ProductName));
            lines.Add(new SlipLine(Quantity(l.Quantity) + " x " + Amount(l.UnitPrice) + " = " +
                                   Amount(OrderDraft.RoundMoney(l.Quantity * l.UnitPrice)), SlipAlign.End));
        }
        lines.Add(new SlipLine(Rule));
        lines.Add(new SlipLine(T("الإجمالي") + ": " + Amount(order.Total), SlipAlign.Start, Bold: true, Large: true));
        if (order.Notes is { Length: > 0 } notes)
        {
            lines.Add(new SlipLine(T("ملاحظات") + ": " + notes));
        }
        if (serverNumber is null)
        {
            lines.Add(new SlipLine(T("لم تُرسل إلى المكتب بعد"), SlipAlign.Center));
        }
        lines.Add(new SlipLine(T("طلبية وليست فاتورة ضريبية"), SlipAlign.Center));
        return lines;
    }

    public const string Rule = "--------------------------------";

    /// <summary>Money with two decimals.</summary>
    public static string Amount(decimal v) => v.ToString("0.00", CultureInfo.InvariantCulture);

    /// <summary>A quantity without the zeros it does not need.</summary>
    public static string Quantity(decimal v) => v.ToString("0.###", CultureInfo.InvariantCulture);
}
