using System.Text;
using System.Text.RegularExpressions;

namespace RetailStore.Domain.Localization;

/// <summary>modLang.Tr of an English front-end (tools/i18n.py translate): every key of the dictionary found in the
/// text, the longest first, is replaced where it is not glued to other Arabic letters; then the Arabic punctuation.
/// The Arabic interface returns the text as it is.</summary>
public static partial class Translator
{
    private static readonly string[] OrderedKeys =
        EnglishTexts.Map.Keys.OrderBy(k => -k.Length).ThenBy(k => k, StringComparer.Ordinal).ToArray();

    [GeneratedRegex("[\u0621-\u064A\u066E-\u066F\u0671-\u06D3\u06FA-\u06FC\u064B-\u065F\u0670\u0640]")]
    private static partial Regex ArabicChar();

    public static bool HasArabic(string? s) => s is not null && ArabicChar().IsMatch(s);

    public static string ToEnglish(string? text)
    {
        if (string.IsNullOrEmpty(text) || !HasArabic(text))
        {
            return text ?? "";
        }
        string output = text;
        foreach (string key in OrderedKeys)
        {
            if (output.Contains(key, StringComparison.Ordinal))
            {
                output = ReplaceWord(output, key, EnglishTexts.Map[key]);
                if (!HasArabic(output))
                {
                    break;
                }
            }
        }
        foreach (var (ar, en) in EnglishTexts.Punctuation)
        {
            output = output.Replace(ar, en, StringComparison.Ordinal);
        }
        return output;
    }

    private static bool Glued(string s, int i) => i >= 0 && i < s.Length && ArabicChar().IsMatch(s[i].ToString());

    /// <summary>modLang.ReplaceWord.</summary>
    public static string ReplaceWord(string s, string key, string value)
    {
        var sb = new StringBuilder();
        int pos = 0;
        while (true)
        {
            int i = s.IndexOf(key, pos, StringComparison.Ordinal);
            if (i < 0)
            {
                break;
            }
            if (Glued(s, i - 1) || Glued(s, i + key.Length))
            {
                sb.Append(s, pos, i + 1 - pos);
                pos = i + 1;
                continue;
            }
            sb.Append(s, pos, i - pos).Append(value);
            pos = i + key.Length;
        }
        sb.Append(s, pos, s.Length - pos);
        return sb.ToString();
    }
}
