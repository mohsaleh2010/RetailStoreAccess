namespace RetailStore.UI.Components.Shared;

/// <summary>The search of LookupField over items in memory: a code typed exactly comes first, then codes that start
/// with the text, then names that hold it (case and Arabic letter forms ignored: أ إ آ = ا, ة = ه, ى = ي).</summary>
public static class LookupMatch
{
    public static IEnumerable<T> Filter<T>(IEnumerable<T> items, string? text, Func<T, string?>? codeOf, Func<T, string?> nameOf,
                                           int max = 20)
    {
        string t = Fold(text);
        if (t.Length == 0)
        {
            return items.Take(max);
        }
        return items
            .Select(item => (item, rank: Rank(Fold(codeOf?.Invoke(item)), Fold(nameOf(item)), t)))
            .Where(x => x.rank < 4)
            .OrderBy(x => x.rank)
            .Select(x => x.item)
            .Take(max);
    }

    private static int Rank(string code, string name, string t)
    {
        if (code.Length > 0 && code == t)
        {
            return 0;
        }
        if (code.Length > 0 && code.StartsWith(t, StringComparison.Ordinal))
        {
            return 1;
        }
        if (name.StartsWith(t, StringComparison.Ordinal))
        {
            return 2;
        }
        return name.Contains(t, StringComparison.Ordinal) || (code.Length > 0 && code.Contains(t, StringComparison.Ordinal)) ? 3 : 4;
    }

    /// <summary>The form compared: trimmed, lower case, the Arabic letter forms of a search folded together.</summary>
    public static string Fold(string? s)
    {
        if (string.IsNullOrWhiteSpace(s))
        {
            return "";
        }
        var chars = s.Trim().ToLowerInvariant().ToCharArray();
        for (int i = 0; i < chars.Length; i++)
        {
            chars[i] = chars[i] switch
            {
                'أ' or 'إ' or 'آ' => 'ا',
                'ة' => 'ه',
                'ى' => 'ي',
                _ => chars[i],
            };
        }
        return new string(chars);
    }
}
