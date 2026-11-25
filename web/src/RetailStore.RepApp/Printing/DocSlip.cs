using System.Globalization;
using RetailStore.Domain.Printing;
using RetailStore.RepApp.Client;

namespace RetailStore.RepApp.Printing;

/// <summary>The customer's copy of an invoice or a receipt the program made for the rep (docs/87): the store, the title (tax
/// invoice or simplified tax invoice, receipt voucher), the number the program gave, the time, the rep, the customer, the lines
/// and the totals; the ZATCA QR of the invoice is drawn under the text (Saudi Arabia). T = the language of the app.</summary>
public static class DocSlip
{
    public static List<SlipLine> Build(RepDocument d, string storeName, string repName, Func<string, string> T)
    {
        var inv = CultureInfo.InvariantCulture;
        var lines = new List<SlipLine>
        {
            new(storeName, SlipAlign.Center, Bold: true, Large: true),
            new(T(d.Title), SlipAlign.Center, Bold: true),
            new(T("الرقم") + ": " + d.Number, Bold: true),
            new(T("التاريخ") + ": " + d.Date.ToString("yyyy-MM-dd HH:mm:ss", inv)),
            new(T("المندوب") + ": " + repName),
            new(T("العميل") + ": " + d.PartyName, Bold: true),
        };
        if (d.Note is { Length: > 0 } note)
        {
            lines.Add(new SlipLine(T("طريقة الدفع") + ": " + T(note)));
        }
        lines.Add(new SlipLine(OrderSlip.Rule));
        foreach (var l in d.Lines)
        {
            lines.Add(new SlipLine(l.Text));
            lines.Add(new SlipLine(OrderSlip.Quantity(l.Quantity) + " = " + OrderSlip.Amount(l.Amount), SlipAlign.End));
        }
        if (d.Lines.Length > 0)
        {
            lines.Add(new SlipLine(OrderSlip.Rule));
        }
        for (int i = 0; i < d.Rows.Length; i++)
        {
            var r = d.Rows[i];
            bool main = d.Kind == "PAYMENT" ? i == 0 : r.Label == "الإجمالي شامل الضريبة";
            lines.Add(new SlipLine(T(r.Label) + ": " + OrderSlip.Amount(r.Amount), Bold: main, Large: main));
        }
        return lines;
    }
}
