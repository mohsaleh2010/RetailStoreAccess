using RetailStore.Domain.Localization;

namespace RetailStore.UI.Services;

/// <summary>The language of the interface for one user (cookie "lang"): Arabic right-to-left, or English left-to-right
/// through the same dictionary as the English front-end of Access (modLang.Tr).</summary>
public sealed class UiLanguage
{
    public const string CookieName = "lang";

    public bool English { get; private set; }

    public void Set(string? code) => English = string.Equals(code, "en", StringComparison.OrdinalIgnoreCase);

    public string Code => English ? "en" : "ar";
    public string Dir => English ? "ltr" : "rtl";

    /// <summary>Tr: the Arabic text, or its English.</summary>
    public string T(string? arabic) => English ? Translator.ToEnglish(arabic) : arabic ?? "";
}
