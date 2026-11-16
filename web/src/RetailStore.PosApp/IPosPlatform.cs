using RetailStore.Domain.Printing;

namespace RetailStore.PosApp;

/// <summary>What the POS needs from Windows (docs/83). The Windows program (RetailStore.Pos) gives the real ones: the
/// protected storage of the secrets (DPAPI of the Windows user), the printers of Windows and the drawing of a receipt with
/// Arabic shaping; the tests give their own.</summary>
public interface IPosPlatform
{
    /// <summary>The name of the PC, written in the audit trail of the login.</summary>
    string DeviceName { get; }

    Task<string?> GetSecretAsync(string key);

    Task SetSecretAsync(string key, string? value);

    /// <summary>The printers of Windows (their names).</summary>
    Task<IReadOnlyList<string>> PrintersAsync();

    /// <summary>Draw the receipt text as a picture of the given width in dots (true = black).</summary>
    bool[,] DrawSlip(IReadOnlyList<SlipLine> lines, int widthDots, bool rightToLeft);

    /// <summary>Send the bytes to the printer as they are (RAW): "" or the reason it failed.</summary>
    Task<string> PrintAsync(string printer, byte[] job);
}

/// <summary>A PC with no protected storage nor printer (tests): secrets in memory, the jobs kept.</summary>
public sealed class BasicPosPlatform : IPosPlatform
{
    private readonly Dictionary<string, string?> _secrets = [];

    public string DeviceName { get; set; } = "POS PC";
    public List<(string Printer, byte[] Job)> Printed { get; } = [];
    public IReadOnlyDictionary<string, string?> Secrets => _secrets;

    public Task<string?> GetSecretAsync(string key) => Task.FromResult(_secrets.GetValueOrDefault(key));

    public Task SetSecretAsync(string key, string? value)
    {
        _secrets[key] = value;
        return Task.CompletedTask;
    }

    public Task<IReadOnlyList<string>> PrintersAsync() => Task.FromResult<IReadOnlyList<string>>(["Receipt printer"]);

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
