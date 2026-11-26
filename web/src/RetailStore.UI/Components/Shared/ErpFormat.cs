using System.Globalization;

namespace RetailStore.UI.Components.Shared;

/// <summary>How the screens write numbers and dates: Latin digits and the invariant culture everywhere (the server
/// culture may be ar-SA, whose calendar is Hijri), amounts with two decimals and thousands separators.</summary>
public static class ErpFormat
{
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

    /// <summary>1,234.50 (two decimals unless asked); "-" for no value.</summary>
    public static string Money(decimal? value, int decimals = 2) =>
        value is { } v ? v.ToString("#,##0." + new string('0', Math.Clamp(decimals, 0, 4)), Invariant).TrimEnd('.') : "-";

    /// <summary>A quantity: up to three decimals, none when whole (12, 1.25, 0.125).</summary>
    public static string Qty(decimal? value) => value is { } v ? v.ToString("#,##0.###", Invariant) : "-";

    /// <summary>A rate kept as a fraction (0.15) written as a percent (15%).</summary>
    public static string Percent(decimal? fraction, int decimals = 1) =>
        fraction is { } f ? (f * 100).ToString("0." + new string('#', Math.Clamp(decimals, 0, 4)), Invariant).TrimEnd('.') + "%" : "-";

    public static string Date(DateTime? value) => value?.ToString("yyyy-MM-dd", Invariant) ?? "";

    /// <summary>The culture of the date pickers: Gregorian, Arabic or English month names.</summary>
    public static CultureInfo DateCulture(bool english)
    {
        try
        {
            var culture = (CultureInfo)CultureInfo.GetCultureInfo(english ? "en-GB" : "ar-EG").Clone();
            if (culture.Calendar is not GregorianCalendar)
            {
                culture.DateTimeFormat.Calendar = new GregorianCalendar();
            }
            return culture;
        }
        catch (CultureNotFoundException)
        {
            return Invariant;
        }
    }
}
