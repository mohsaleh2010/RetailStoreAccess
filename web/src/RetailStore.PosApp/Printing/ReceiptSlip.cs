using System.Globalization;
using QRCoder;
using RetailStore.Domain.Printing;
using RetailStore.PosApp.Offline;

namespace RetailStore.PosApp.Printing;

/// <summary>The receipt of a sale made on the POS (docs/83): a simplified tax invoice as the program prints it
/// (qrySalesDocPrint): the store and its VAT number, the number the PC gave the invoice, the time, the cashier, the
/// customer, each line in its unit, the totals, the payment and the change, and the ZATCA QR (Saudi Arabia). The text is
/// drawn by Windows (Arabic shaped, right to left) and sent as a picture with the QR under it; a cash sale opens the
/// drawer. T = the language of the POS (modLang.Tr).</summary>
public static class ReceiptSlip
{
    public const string Rule = "--------------------------------";

    public static string Amount(decimal v) => v.ToString("#,##0.00", CultureInfo.InvariantCulture);

    public static string Quantity(decimal v) => v.ToString("0.###", CultureInfo.InvariantCulture);

    public static List<SlipLine> Build(PosReceipt r, bool sent, Func<string, string> T)
    {
        var inv = CultureInfo.InvariantCulture;
        var lines = new List<SlipLine> { new(r.StoreName, SlipAlign.Center, Bold: true, Large: true) };
        if (r.StoreVAT is { Length: > 0 } vat)
        {
            lines.Add(new SlipLine(T("الرقم الضريبي") + ": " + vat, SlipAlign.Center));
        }
        lines.Add(new SlipLine(T(string.IsNullOrEmpty(r.CustomerVAT) ? "فاتورة ضريبية مبسطة" : "فاتورة ضريبية"), SlipAlign.Center, Bold: true));
        lines.Add(new SlipLine(T("رقم الفاتورة") + ": " + r.Request.InvoiceNumber, Bold: true));
        lines.Add(new SlipLine(T("التاريخ") + ": " + r.Request.TakenAt.ToString("yyyy-MM-dd HH:mm:ss", inv)));
        lines.Add(new SlipLine(T("الكاشير") + ": " + r.CashierName));
        lines.Add(new SlipLine(T("العميل") + ": " + r.CustomerName));
        if (r.CustomerVAT is { Length: > 0 } cv)
        {
            lines.Add(new SlipLine(T("الرقم الضريبي للعميل") + ": " + cv));
        }
        lines.Add(new SlipLine(Rule));
        foreach (var l in r.Lines)
        {
            lines.Add(new SlipLine(l.ProductName + (l.UnitName.Length > 0 ? " (" + l.UnitName + ")" : "")));
            lines.Add(new SlipLine(Quantity(l.Quantity) + " x " + Amount(l.UnitPrice) + " = " + Amount(l.Total), SlipAlign.End));
            if (l.Discount > 0)
            {
                lines.Add(new SlipLine(T("الخصم") + ": " + Amount(l.Discount), SlipAlign.End));
            }
        }
        lines.Add(new SlipLine(Rule));
        lines.Add(new SlipLine(T("الإجمالي قبل الضريبة") + ": " + Amount(r.Taxable)));
        if (r.Discount > 0)
        {
            lines.Add(new SlipLine(T("الخصم") + ": " + Amount(r.Discount)));
        }
        lines.Add(new SlipLine(T("ضريبة القيمة المضافة") + ": " + Amount(r.Tax)));
        lines.Add(new SlipLine(T("الإجمالي شامل الضريبة") + ": " + Amount(r.Total) + " " + r.Currency, Bold: true, Large: true));
        if (r.Request.PaymentType == "CREDIT")
        {
            lines.Add(new SlipLine(T("المدفوع") + ": " + Amount(r.Paid)));
            lines.Add(new SlipLine(T("المتبقي") + ": " + Amount(r.Remaining), Bold: true));
        }
        else
        {
            lines.Add(new SlipLine(T("المبلغ المستلم") + ": " + Amount(r.Tendered)));
            lines.Add(new SlipLine(T("الباقي للعميل") + ": " + Amount(r.Change), Bold: true));
        }
        if (r.MethodName is { Length: > 0 } m)
        {
            lines.Add(new SlipLine(T("طريقة الدفع") + ": " + m));
        }
        if (!sent)
        {
            lines.Add(new SlipLine(T("لم تُرسل إلى البرنامج بعد"), SlipAlign.Center));
        }
        if (r.Footer is { Length: > 0 } footer)
        {
            lines.Add(new SlipLine(footer, SlipAlign.Center));
        }
        return lines;
    }

    /// <summary>The QR as dots: each module a square of the given size, with its quiet zone.</summary>
    public static bool[,] Qr(string text, int moduleDots)
    {
        using var gen = new QRCodeGenerator();
        using var data = gen.CreateQrCode(text, QRCodeGenerator.ECCLevel.M);
        var m = data.ModuleMatrix;
        int n = m.Count;
        var dots = new bool[n * moduleDots, n * moduleDots];
        for (int y = 0; y < n; y++)
        {
            for (int x = 0; x < n; x++)
            {
                if (!m[y][x])
                {
                    continue;
                }
                for (int dy = 0; dy < moduleDots; dy++)
                {
                    for (int dx = 0; dx < moduleDots; dx++)
                    {
                        dots[y * moduleDots + dy, x * moduleDots + dx] = true;
                    }
                }
            }
        }
        return dots;
    }

    /// <summary>The text picture and the QR under it, centred, in one picture of the width of the text.</summary>
    public static bool[,] Stack(bool[,] text, bool[,]? qr)
    {
        int width = text.GetLength(1);
        int textRows = text.GetLength(0);
        int qrRows = qr?.GetLength(0) ?? 0;
        int qrCols = qr?.GetLength(1) ?? 0;
        var all = new bool[textRows + qrRows, width];
        for (int y = 0; y < textRows; y++)
        {
            for (int x = 0; x < width; x++)
            {
                all[y, x] = text[y, x];
            }
        }
        int left = Math.Max(0, (width - qrCols) / 2);
        for (int y = 0; y < qrRows; y++)
        {
            for (int x = 0; x < qrCols && left + x < width; x++)
            {
                all[textRows + y, left + x] = qr![y, x];
            }
        }
        return all;
    }

    /// <summary>The whole job: the drawer pulse for a cash sale, then the picture (EscPos.Picture: bands, feed, cut).</summary>
    public static byte[] Job(bool[,] picture, bool openDrawer)
    {
        var job = new List<byte>();
        if (openDrawer)
        {
            job.AddRange(EscPos.Initialize);
            job.AddRange(EscPos.OpenDrawer);
        }
        job.AddRange(EscPos.Picture(picture));
        return job.ToArray();
    }
}
