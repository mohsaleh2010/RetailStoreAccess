namespace RetailStore.RepApp.Printing;

/// <summary>ESC/POS commands of a Bluetooth receipt printer (docs/78). The slip is sent as a picture (GS v 0), never as
/// text: the printers do not join Arabic letters nor write right to left, the phone does. A picture of many rows is
/// sent in bands, which the small printers' buffers accept.</summary>
public static class EscPos
{
    /// <summary>Dots of a line: 384 on 58 mm paper, 576 on 80 mm.</summary>
    public static int DotsFor(int paperMillimetres) => paperMillimetres >= 80 ? 576 : 384;

    public const int BandRows = 128;

    public static readonly byte[] Initialize = [0x1B, 0x40];

    public static byte[] Feed(int lines) => [0x1B, 0x64, (byte)Math.Clamp(lines, 0, 255)];

    /// <summary>GS V 66 0: feed to the cutter and cut (printers without a cutter ignore it).</summary>
    public static readonly byte[] Cut = [0x1D, 0x56, 0x42, 0x00];

    /// <summary>GS v 0 of a 1-bit picture (true = black), its width padded to whole bytes, the most significant bit first.</summary>
    public static byte[] Raster(bool[,] dots, int firstRow, int rows)
    {
        int width = dots.GetLength(1);
        int bytesPerRow = (width + 7) / 8;
        var data = new byte[8 + bytesPerRow * rows];
        data[0] = 0x1D; data[1] = 0x76; data[2] = 0x30; data[3] = 0x00;
        data[4] = (byte)(bytesPerRow & 0xFF); data[5] = (byte)(bytesPerRow >> 8);
        data[6] = (byte)(rows & 0xFF); data[7] = (byte)(rows >> 8);
        for (int y = 0; y < rows; y++)
        {
            for (int x = 0; x < width; x++)
            {
                if (dots[firstRow + y, x])
                {
                    data[8 + y * bytesPerRow + x / 8] |= (byte)(0x80 >> (x % 8));
                }
            }
        }
        return data;
    }

    /// <summary>The whole job: initialize, the picture in bands, feed, cut.</summary>
    public static byte[] Picture(bool[,] dots, int feedLines = 4)
    {
        var job = new List<byte>(Initialize);
        int height = dots.GetLength(0);
        for (int y = 0; y < height; y += BandRows)
        {
            job.AddRange(Raster(dots, y, Math.Min(BandRows, height - y)));
        }
        job.AddRange(Feed(feedLines));
        job.AddRange(Cut);
        return job.ToArray();
    }

    /// <summary>A grey picture (0 = black .. 255 = white, rows x columns) to dots: darker than the threshold is black.</summary>
    public static bool[,] Threshold(byte[] grey, int width, int height, byte threshold = 160)
    {
        var dots = new bool[height, width];
        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                dots[y, x] = grey[y * width + x] < threshold;
            }
        }
        return dots;
    }
}
