namespace RetailStore.Domain.Sales;

/// <summary>The scale barcode settings (LabelSettings.ScalePrefixes, ScaleItemDigits, ScaleValueKind; docs/86).</summary>
public sealed record ScaleSettings(string? Prefixes, int ItemDigits, string ValueKind)
{
    public static readonly ScaleSettings None = new(null, 5, "WEIGHT");

    public bool Enabled => ScaleBarcode.PrefixesOf(Prefixes).Count > 0;
}

/// <summary>A scale barcode read: the item code and the value (grams, or cents / piasters).</summary>
public sealed record ScaleCode(string ItemCode, long Value);

/// <summary>The scale barcodes of modScale (docs/86): an EAN-13 of 13 digits = a prefix of the settings, the item code
/// (the code of the product), the weight in grams or the price in cents, and the check digit. The same rule as VBA
/// (modScale.ScaleCodeParts / ScaleLine) and its Python mirror (tools/scale_reference.py).</summary>
public static class ScaleBarcode
{
    /// <summary>The EAN-13 check digit of 12 digits: weights 1, 3, 1, 3 ... from the left.</summary>
    public static int CheckDigit(string digits12)
    {
        int total = 0;
        for (int i = 0; i < 12; i++)
        {
            total += (digits12[i] - '0') * (i % 2 == 1 ? 3 : 1);
        }
        return (10 - total % 10) % 10;
    }

    private static bool IsDigits(string s) => s.Length > 0 && s.All(c => c is >= '0' and <= '9');

    /// <summary>The prefixes of the settings: each of 1 to 3 digits; anything else is ignored.</summary>
    public static List<string> PrefixesOf(string? setting) =>
        (setting ?? "").Split(',').Select(p => p.Trim()).Where(p => p.Length is >= 1 and <= 3 && IsDigits(p)).ToList();

    /// <summary>The item code and the value when the code is a scale barcode of the settings with a right check digit.</summary>
    public static ScaleCode? Parts(string? code, ScaleSettings settings)
    {
        string c = (code ?? "").Trim();
        int digits = settings.ItemDigits;
        if (c.Length != 13 || !IsDigits(c) || digits is < 4 or > 6)
        {
            return null;
        }
        foreach (var p in PrefixesOf(settings.Prefixes))
        {
            if (c.StartsWith(p, StringComparison.Ordinal) && p.Length + digits <= 11)
            {
                if (CheckDigit(c[..12]) != c[12] - '0')
                {
                    return null;
                }
                return new ScaleCode(c.Substring(p.Length, digits), long.Parse(c.Substring(p.Length + digits, 12 - p.Length - digits)));
            }
        }
        return null;
    }

    /// <summary>The product codes that match an item code: as printed, then without its leading zeros.</summary>
    public static string[] ItemCodes(string itemCode)
    {
        string stripped = itemCode.TrimStart('0');
        return stripped.Length == 0 || stripped == itemCode ? [itemCode] : [itemCode, stripped];
    }

    /// <summary>The quantity of the line: the weight in kilograms, or the price / the unit price (3 decimals); 0 when a
    /// price label meets a product without a price.</summary>
    public static decimal Quantity(long value, string valueKind, decimal unitPrice)
    {
        if (valueKind == "PRICE")
        {
            return unitPrice <= 0 ? 0 : Math.Round(value / 100m / unitPrice, 3, MidpointRounding.AwayFromZero);
        }
        return value / 1000m;
    }

    public static string NoProduct(string code, string itemCode) => "باركود الميزان " + code + ": لا يوجد صنف نشط بالكود " + itemCode + ".";

    public static string NoPrice(string code) => "باركود الميزان " + code + ": سعر الصنف صفر، فلا تُحسب الكمية من السعر.";

    public static string ZeroQuantity(string code) => "باركود الميزان " + code + ": الكمية صفر.";
}
