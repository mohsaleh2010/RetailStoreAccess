using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using RetailStore.Domain.Printing;
using RetailStore.PosApp;

namespace RetailStore.Pos;

/// <summary>Windows for the POS (docs/83): the secrets (the SQL password, the API tokens) encrypted for the Windows user
/// with DPAPI in a file of the program's folder; the printers of Windows; the receipt drawn by GDI+ (which joins the
/// Arabic letters and writes right to left) and sent RAW to the printer as ESC/POS (EscPos.Picture).</summary>
public sealed class WindowsPlatform : IPosPlatform
{
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("RetailStore.Pos");
    private readonly string _secretsFile = Path.Combine(FileSystem.AppDataDirectory, "secrets.dat");
    private readonly SemaphoreSlim _lock = new(1, 1);

    public string DeviceName => Environment.MachineName;

    private Dictionary<string, string> ReadSecrets()
    {
        if (!File.Exists(_secretsFile))
        {
            return [];
        }
        try
        {
            byte[] plain = ProtectedData.Unprotect(File.ReadAllBytes(_secretsFile), Entropy, DataProtectionScope.CurrentUser);
            return JsonSerializer.Deserialize<Dictionary<string, string>>(plain) ?? [];
        }
        catch (CryptographicException)
        {
            return [];      // another Windows user: his secrets cannot be read, the administrator sets them again
        }
    }

    public async Task<string?> GetSecretAsync(string key)
    {
        await _lock.WaitAsync();
        try
        {
            return ReadSecrets().GetValueOrDefault(key);
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task SetSecretAsync(string key, string? value)
    {
        await _lock.WaitAsync();
        try
        {
            var all = ReadSecrets();
            if (value is null)
            {
                all.Remove(key);
            }
            else
            {
                all[key] = value;
            }
            byte[] plain = JsonSerializer.SerializeToUtf8Bytes(all);
            await File.WriteAllBytesAsync(_secretsFile, ProtectedData.Protect(plain, Entropy, DataProtectionScope.CurrentUser));
        }
        finally
        {
            _lock.Release();
        }
    }

    public Task<IReadOnlyList<string>> PrintersAsync() =>
        Task.FromResult<IReadOnlyList<string>>(PrinterSettings.InstalledPrinters.Cast<string>().OrderBy(p => p).ToList());

    /// <summary>The receipt text as a picture: each line wrapped to the width, aligned, bold or large as asked.</summary>
    public bool[,] DrawSlip(IReadOnlyList<SlipLine> lines, int widthDots, bool rightToLeft)
    {
        using var normal = new Font("Tahoma", 22, FontStyle.Regular, GraphicsUnit.Pixel);
        using var bold = new Font("Tahoma", 22, FontStyle.Bold, GraphicsUnit.Pixel);
        using var large = new Font("Tahoma", 30, FontStyle.Bold, GraphicsUnit.Pixel);
        using var measure = new Bitmap(1, 1);
        using var mg = Graphics.FromImage(measure);
        var layout = new List<(SlipLine Line, Font Font, StringFormat Format, int Height)>();
        int total = 0;
        foreach (var l in lines)
        {
            var font = l.Large ? large : l.Bold ? bold : normal;
            var format = new StringFormat(rightToLeft ? StringFormatFlags.DirectionRightToLeft : 0)
            {
                Alignment = l.Align switch { SlipAlign.Center => StringAlignment.Center, SlipAlign.End => StringAlignment.Far, _ => StringAlignment.Near },
            };
            int h = (int)Math.Ceiling(mg.MeasureString(l.Text.Length == 0 ? " " : l.Text, font, widthDots, format).Height) + 2;
            layout.Add((l, font, format, h));
            total += h;
        }
        using var bmp = new Bitmap(widthDots, Math.Max(total, 1), PixelFormat.Format24bppRgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.White);
            g.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;
            g.SmoothingMode = SmoothingMode.None;
            int y = 0;
            foreach (var (l, font, format, h) in layout)
            {
                g.DrawString(l.Text, font, Brushes.Black, new RectangleF(0, y, widthDots, h), format);
                format.Dispose();
                y += h;
            }
        }
        var grey = new byte[widthDots * bmp.Height];
        var data = bmp.LockBits(new Rectangle(0, 0, widthDots, bmp.Height), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
        try
        {
            var row = new byte[data.Stride];
            for (int yy = 0; yy < bmp.Height; yy++)
            {
                Marshal.Copy(data.Scan0 + yy * data.Stride, row, 0, data.Stride);
                for (int x = 0; x < widthDots; x++)
                {
                    int b = row[x * 3], gr = row[x * 3 + 1], r = row[x * 3 + 2];
                    grey[yy * widthDots + x] = (byte)((r * 30 + gr * 59 + b * 11) / 100);
                }
            }
        }
        finally
        {
            bmp.UnlockBits(data);
        }
        return EscPos.Threshold(grey, widthDots, bmp.Height);
    }

    public Task<string> PrintAsync(string printer, byte[] job) => Task.Run(() => RawPrinter.Send(printer, job));
}

/// <summary>A RAW job to a printer of Windows (winspool): the ESC/POS bytes go to the printer as they are.</summary>
internal static class RawPrinter
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private sealed class DocInfo
    {
        public string DocName = "RetailStore receipt";
        public string? OutputFile;
        public string DataType = "RAW";
    }

    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool OpenPrinter(string name, out IntPtr handle, IntPtr defaults);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool ClosePrinter(IntPtr handle);

    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int StartDocPrinter(IntPtr handle, int level, [In] DocInfo info);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool EndDocPrinter(IntPtr handle);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool StartPagePrinter(IntPtr handle);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool EndPagePrinter(IntPtr handle);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool WritePrinter(IntPtr handle, byte[] bytes, int count, out int written);

    /// <summary>"" or the Windows error.</summary>
    public static string Send(string printer, byte[] job)
    {
        if (!OpenPrinter(printer, out var h, IntPtr.Zero))
        {
            return new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error()).Message;
        }
        try
        {
            if (StartDocPrinter(h, 1, new DocInfo()) == 0)
            {
                return new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error()).Message;
            }
            try
            {
                StartPagePrinter(h);
                bool ok = WritePrinter(h, job, job.Length, out int written);
                EndPagePrinter(h);
                return ok && written == job.Length ? "" : new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error()).Message;
            }
            finally
            {
                EndDocPrinter(h);
            }
        }
        finally
        {
            ClosePrinter(h);
        }
    }
}
