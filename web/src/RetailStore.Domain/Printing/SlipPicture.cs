using QRCoder;

namespace RetailStore.Domain.Printing;

/// <summary>The picture of a printed slip with its QR (the POS on Windows, docs/83; the rep app, docs/87).</summary>
public static class SlipPicture
{
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
}
