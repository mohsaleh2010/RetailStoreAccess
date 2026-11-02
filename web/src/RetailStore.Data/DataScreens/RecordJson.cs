using System.Globalization;
using System.Text.Json;

namespace RetailStore.Data.DataScreens;

/// <summary>A record of a data screen as JSON for the API (docs/77): its key and the data fields of the screen, by their
/// database names; and the fields a client sends, written on the record with the conversions of the screen.</summary>
public static class RecordJson
{
    public static Dictionary<string, object?> ToJson<T>(ScreenSpec<T> spec, T e) where T : class, new()
    {
        var d = new Dictionary<string, object?> { [spec.Key] = spec.KeyOf(e) };
        foreach (var f in spec.Fields.Where(f => f.IsData))
        {
            d[f.Name] = f.Get(e);
        }
        return d;
    }

    /// <summary>Writes the sent fields on the record (names in any case: ProductName or productName). The key and the locked fields (filled by the program: quantity,
    /// cost, balance) are ignored, so a record read from the API can be sent back as it is. Returns "" or the message
    /// and the field: an unknown field, or a value of the wrong kind.</summary>
    public static (string Message, string Field) Apply<T>(ScreenSpec<T> spec, T e, IReadOnlyDictionary<string, JsonElement> values)
        where T : class, new()
    {
        foreach (var (name, json) in values)
        {
            if (string.Equals(name, spec.Key, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }
            var f = spec.Fields.FirstOrDefault(x => x.IsData && string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase));
            if (f is null)
            {
                return ("حقل غير معروف: " + name, name);
            }
            if (f.Locked)
            {
                continue;
            }
            try
            {
                f.Set(e, Value(f.Kind, json));
            }
            catch (Exception ex) when (ex is FormatException or InvalidCastException or OverflowException or InvalidOperationException)
            {
                return ("قيمة غير مقبولة في حقل «" + f.Label + "».", f.Name);
            }
        }
        return ("", "");
    }

    private static object? Value(FieldKind kind, JsonElement json) => json.ValueKind switch
    {
        JsonValueKind.Null or JsonValueKind.Undefined => null,
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Number when kind is FieldKind.Text or FieldKind.Memo or FieldKind.Choice => json.GetRawText(),
        JsonValueKind.Number => json.GetDecimal(),
        JsonValueKind.String when kind is FieldKind.Date or FieldKind.DateTime =>
            DateTime.Parse(json.GetString()!, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind),
        JsonValueKind.String when kind is FieldKind.Bool => bool.Parse(json.GetString()!),
        JsonValueKind.String when kind is FieldKind.Money or FieldKind.Quantity or FieldKind.Rate or FieldKind.Integer =>
            json.GetString()!.Trim().Length == 0 ? null : decimal.Parse(json.GetString()!, NumberStyles.Number, CultureInfo.InvariantCulture),
        JsonValueKind.String => json.GetString(),
        _ => throw new FormatException("Not a value"),
    };
}
