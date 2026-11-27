namespace RetailStore.UI.Services;

/// <summary>The light or dark mode of the interface for one user (cookie "theme"), read by App.razor and kept by the
/// interactive session like the language.</summary>
public sealed class UiTheme
{
    public const string CookieName = "theme";

    public bool Dark { get; private set; }

    public void Set(string? value) => Dark = string.Equals(value, "dark", StringComparison.OrdinalIgnoreCase);

    public void Set(bool dark) => Dark = dark;

    public string CookieValue => Dark ? "dark" : "light";
}
