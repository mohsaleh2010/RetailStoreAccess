using Android.Bluetooth;
using Android.Content;
using Android.Graphics;
using Android.Text;
using Java.Util;
using RetailStore.RepApp;
using RetailStore.RepApp.Printing;
using AColor = Android.Graphics.Color;
using ALayout = Android.Text.Layout;

namespace RetailStore.App;

/// <summary>The phone's services for the rep app (docs/78): the token in the Android keystore (SecureStorage), the
/// network state, the paired Bluetooth printers (serial port profile) and the slip drawn by Android, which joins the
/// Arabic letters and writes right to left, then sent to the printer as a picture (EscPos.Picture).</summary>
public sealed class PhonePlatform : IRepPlatform
{
    private static readonly UUID SerialPort = UUID.FromString("00001101-0000-1000-8000-00805F9B34FB")!;

    public string DeviceName => (DeviceInfo.Current.Manufacturer + " " + DeviceInfo.Current.Model).Trim();

    public bool HasNetwork => Connectivity.Current.NetworkAccess != NetworkAccess.None;

    public async Task<string?> GetSecretAsync(string key)
    {
        try
        {
            return await SecureStorage.Default.GetAsync(key);
        }
        catch (Exception)
        {
            return null;       // the keystore was reset (a backup restored on another phone): log in again
        }
    }

    public async Task SetSecretAsync(string key, string? value)
    {
        if (value is null)
        {
            SecureStorage.Default.Remove(key);
        }
        else
        {
            await SecureStorage.Default.SetAsync(key, value);
        }
    }

    private static BluetoothAdapter? Adapter() =>
        (Platform.AppContext.GetSystemService(Context.BluetoothService) as BluetoothManager)?.Adapter;

    private static async Task<bool> AllowedAsync()
    {
        var status = await MainThread.InvokeOnMainThreadAsync(Permissions.RequestAsync<Permissions.Bluetooth>);
        return status == PermissionStatus.Granted;
    }

    public async Task<IReadOnlyList<string>> PrintersAsync()
    {
        if (!await AllowedAsync() || Adapter() is not { IsEnabled: true } adapter)
        {
            return [];
        }
        return adapter.BondedDevices?.Select(d => d.Name ?? d.Address ?? "").Where(n => n.Length > 0).Order().ToList() ?? [];
    }

    public async Task<string> PrintAsync(string printer, byte[] job)
    {
        if (!await AllowedAsync())
        {
            return "Bluetooth permission";
        }
        if (Adapter() is not { IsEnabled: true } adapter)
        {
            return "Bluetooth is off";
        }
        var device = adapter.BondedDevices?.FirstOrDefault(d => d.Name == printer || d.Address == printer);
        if (device is null)
        {
            return printer;
        }
        return await Task.Run(async () =>
        {
            try
            {
                using var socket = device.CreateRfcommSocketToServiceRecord(SerialPort)!;
                await socket.ConnectAsync();
                await socket.OutputStream!.WriteAsync(job);
                await socket.OutputStream.FlushAsync();
                await Task.Delay(500);       // let the printer read its buffer before the link closes
                socket.Close();
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        });
    }

    public bool[,] DrawSlip(IReadOnlyList<SlipLine> lines, int widthDots, bool rightToLeft)
    {
        var layouts = new List<StaticLayout>();
        foreach (var line in lines)
        {
            var paint = new TextPaint(PaintFlags.AntiAlias) { Color = AColor.Black, TextSize = line.Large ? 30 : 23 };
            paint.SetTypeface(Typeface.Create(Typeface.Default, line.Bold ? TypefaceStyle.Bold : TypefaceStyle.Normal));
            var align = line.Align switch
            {
                SlipAlign.Center => ALayout.Alignment.AlignCenter!,
                SlipAlign.End => ALayout.Alignment.AlignOpposite!,
                _ => ALayout.Alignment.AlignNormal!,
            };
            string text = line.Text.Length == 0 ? " " : line.Text;
            layouts.Add(StaticLayout.Builder.Obtain(text, 0, text.Length, paint, widthDots)!
                .SetAlignment(align)!
                .SetTextDirection(rightToLeft ? TextDirectionHeuristics.Rtl! : TextDirectionHeuristics.Ltr!)!
                .Build()!);
        }
        int height = Math.Max(1, layouts.Sum(l => l.Height) + 8);
        using var bitmap = Bitmap.CreateBitmap(widthDots, height, Bitmap.Config.Argb8888!)!;
        using (var canvas = new Canvas(bitmap))
        {
            canvas.DrawColor(AColor.White);
            foreach (var l in layouts)
            {
                l.Draw(canvas);
                canvas.Translate(0, l.Height);
            }
        }
        var pixels = new int[widthDots * height];
        bitmap.GetPixels(pixels, 0, widthDots, 0, 0, widthDots, height);
        var grey = new byte[pixels.Length];
        for (int i = 0; i < pixels.Length; i++)
        {
            int p = pixels[i];
            grey[i] = (byte)((((p >> 16) & 0xFF) + ((p >> 8) & 0xFF) + (p & 0xFF)) / 3);
        }
        return EscPos.Threshold(grey, widthDots, height);
    }
}
