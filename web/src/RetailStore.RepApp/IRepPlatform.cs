using RetailStore.RepApp.Printing;

namespace RetailStore.RepApp;

/// <summary>What the rep app needs from the phone (docs/78). The Android app (RetailStore.App) gives the real ones: the
/// secure storage of the token, the network state, the Bluetooth printers and the drawing of a slip with Arabic shaping;
/// the tests give their own.</summary>
public interface IRepPlatform
{
    /// <summary>The name written in the audit trail of the login ("API" + this + the address).</summary>
    string DeviceName { get; }

    /// <summary>False when the phone knows it has no network: the sync is not even tried.</summary>
    bool HasNetwork { get; }

    Task<string?> GetSecretAsync(string key);

    Task SetSecretAsync(string key, string? value);

    /// <summary>The paired Bluetooth printers (their names).</summary>
    Task<IReadOnlyList<string>> PrintersAsync();

    /// <summary>Draw the slip as a picture of the given width in dots (true = black).</summary>
    bool[,] DrawSlip(IReadOnlyList<SlipLine> lines, int widthDots, bool rightToLeft);

    /// <summary>Send the bytes to the printer: "" or the reason it failed.</summary>
    Task<string> PrintAsync(string printer, byte[] job);
}

/// <summary>A phone with no secure storage, printer or network knowledge (tests): secrets in memory.</summary>
public sealed class BasicPlatform : IRepPlatform
{
    private readonly Dictionary<string, string?> _secrets = [];

    public string DeviceName { get; set; } = "Rep app";
    public bool HasNetwork { get; set; } = true;
    public List<(string Printer, byte[] Job)> Printed { get; } = [];

    public Task<string?> GetSecretAsync(string key) => Task.FromResult(_secrets.GetValueOrDefault(key));

    public Task SetSecretAsync(string key, string? value)
    {
        _secrets[key] = value;
        return Task.CompletedTask;
    }

    public Task<IReadOnlyList<string>> PrintersAsync() => Task.FromResult<IReadOnlyList<string>>(["Test printer"]);

    /// <summary>One black row per line, as wide as its text, so the tests can count them.</summary>
    public bool[,] DrawSlip(IReadOnlyList<SlipLine> lines, int widthDots, bool rightToLeft)
    {
        var dots = new bool[lines.Count, widthDots];
        for (int y = 0; y < lines.Count; y++)
        {
            for (int x = 0; x < Math.Min(widthDots, lines[y].Text.Length); x++)
            {
                dots[y, rightToLeft ? widthDots - 1 - x : x] = true;
            }
        }
        return dots;
    }

    public Task<string> PrintAsync(string printer, byte[] job)
    {
        Printed.Add((printer, job));
        return Task.FromResult("");
    }
}
