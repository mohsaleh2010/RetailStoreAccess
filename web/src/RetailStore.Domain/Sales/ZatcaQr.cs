using System.Globalization;
using System.Text;
using RetailStore.Domain.Common;

namespace RetailStore.Domain.Sales;

/// <summary>The ZATCA (Fatoora) phase-1 QR of modZatca.BuildZatcaQR and tools/zatca_reference.py:
/// Base64( TLV(1 seller) TLV(2 VAT number) TLV(3 UTC time) TLV(4 total with VAT) TLV(5 VAT) ).</summary>
public static class ZatcaQr
{
    public const int KsaUtcOffsetHours = 3;

    public static string Build(string seller, string vatNumber, DateTime invoiceDate, decimal totalWithVat, decimal vatTotal)
    {
        var data = new List<byte>();
        Append(data, 1, FitTo255Bytes(seller));
        Append(data, 2, vatNumber);
        Append(data, 3, Timestamp(invoiceDate));
        Append(data, 4, Amount(totalWithVat));
        Append(data, 5, Amount(vatTotal));
        return Convert.ToBase64String(data.ToArray());
    }

    /// <summary>modZatca.ZatcaAmount: two decimals, "." and ASCII digits.</summary>
    public static string Amount(decimal v) => MoneyMath.RoundMoney(v).ToString("0.00", CultureInfo.InvariantCulture);

    /// <summary>modZatca.ZatcaTimestamp: Saudi local time -> ISO 8601 UTC.</summary>
    public static string Timestamp(DateTime local) =>
        local.AddHours(-KsaUtcOffsetHours).ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);

    /// <summary>A TLV length is one byte: a very long name is shortened on a character boundary.</summary>
    private static string FitTo255Bytes(string s)
    {
        while (Encoding.UTF8.GetByteCount(s) > 255)
        {
            s = char.IsLowSurrogate(s[^1]) && s.Length > 1 ? s[..^2] : s[..^1];
        }
        return s;
    }

    private static void Append(List<byte> data, byte tag, string value)
    {
        var b = Encoding.UTF8.GetBytes(value);
        if (b.Length > 255)
        {
            throw new ArgumentException("TLV value too long (tag " + tag + ")");
        }
        data.Add(tag);
        data.Add((byte)b.Length);
        data.AddRange(b);
    }

    /// <summary>tools/zatca_reference.decode_payload: tag -> text.</summary>
    public static Dictionary<int, string> Decode(string base64)
    {
        var raw = Convert.FromBase64String(base64);
        var result = new Dictionary<int, string>();
        for (int i = 0; i + 1 < raw.Length; i += 2 + raw[i + 1])
        {
            result[raw[i]] = Encoding.UTF8.GetString(raw, i + 2, raw[i + 1]);
        }
        return result;
    }
}
