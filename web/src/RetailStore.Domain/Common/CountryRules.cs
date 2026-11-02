using System.Text.RegularExpressions;

namespace RetailStore.Domain.Common;

/// <summary>The operating country (modCountry, tools/country_reference.py): Settings.CountryCode SA or EG.</summary>
public static class CountryRules
{
    public static string Normalize(string? code) => string.Equals(code?.Trim(), "EG", StringComparison.OrdinalIgnoreCase) ? "EG" : "SA";

    /// <summary>"" when the tax number fits the country (an empty number is allowed: not VAT-registered).</summary>
    public static string TaxNumberProblem(string? value, string country)
    {
        string s = (value ?? "").Trim();
        if (s.Length == 0)
        {
            return "";
        }
        if (Normalize(country) == "EG")
        {
            return Regex.IsMatch(s, "^[0-9]{9}$") ? "" : "رقم التسجيل الضريبي في مصر 9 أرقام.";
        }
        return Regex.IsMatch(s, "^3[0-9]{13}3$") ? "" : "الرقم الضريبي في السعودية 15 رقمًا ويبدأ وينتهي بالرقم 3.";
    }

    /// <summary>A mobile number of the country: 05xxxxxxxx (Saudi Arabia), 01xxxxxxxxx (Egypt), or with the country code.</summary>
    public static bool MobileFits(string mobile, string country) =>
        Normalize(country) == "EG"
            ? Regex.IsMatch(mobile, @"^(01|\+201|201)[0-9]{9}$")
            : Regex.IsMatch(mobile, @"^(05|\+9665|9665)[0-9]{8}$");

    public static string MobileExample(string country) => Normalize(country) == "EG" ? "01xxxxxxxxx" : "05xxxxxxxx";

    public static string CountryName(string country) =>
        Normalize(country) == "EG" ? "جمهورية مصر العربية" : "المملكة العربية السعودية";

    /// <summary>The Access check of an e-mail address: Like "*?@?*.?*".</summary>
    public static bool EmailFits(string email) => Regex.IsMatch(email, @"^.+@.+\..+$", RegexOptions.Singleline);
}
