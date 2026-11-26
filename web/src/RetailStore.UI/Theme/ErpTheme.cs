using MudBlazor;

namespace RetailStore.UI.Theme;

/// <summary>The one MudBlazor theme of the web (docs/ERP-Redesign-Prompt.md, phase 2): navy and teal, the Cairo font,
/// light and dark palettes. Colors are tuned for contrast (WCAG AA on white / on the dark surface); pages never use hex
/// colors of their own.</summary>
public static class ErpTheme
{
    private static readonly string[] Font = ["Cairo", "Segoe UI", "Tahoma", "sans-serif"];

    public static readonly MudTheme Theme = new()
    {
        PaletteLight = new PaletteLight
        {
            Primary = "#1E3A5F",
            Secondary = "#0B7A6B",
            Tertiary = "#F2A541",
            TertiaryContrastText = "#1F2937",
            Success = "#2E7D32",
            Error = "#C62828",
            Warning = "#B45309",
            Info = "#1D4ED8",
            Background = "#F5F7FA",
            BackgroundGray = "#EEF1F5",
            Surface = "#FFFFFF",
            AppbarBackground = "#1E3A5F",
            AppbarText = "#FFFFFF",
            DrawerBackground = "#FFFFFF",
            DrawerText = "#1F2937",
            DrawerIcon = "#475569",
            TextPrimary = "#1F2937",
            TextSecondary = "#4B5563",
            ActionDefault = "#4B5563",
            LinesDefault = "#E2E8F0",
            TableLines = "#E2E8F0",
            TableStriped = "rgba(15,23,42,0.025)",
            TableHover = "rgba(30,58,95,0.06)",
            Divider = "#E2E8F0",
        },
        PaletteDark = new PaletteDark
        {
            Primary = "#60A5FA",
            PrimaryContrastText = "#0F172A",
            Secondary = "#2DD4BF",
            SecondaryContrastText = "#0F172A",
            Tertiary = "#F2A541",
            TertiaryContrastText = "#1F2937",
            Success = "#4ADE80",
            Error = "#F87171",
            Warning = "#FBBF24",
            Info = "#93C5FD",
            Background = "#0F172A",
            BackgroundGray = "#111C33",
            Surface = "#1E293B",
            AppbarBackground = "#111827",
            AppbarText = "#E5E7EB",
            DrawerBackground = "#1E293B",
            DrawerText = "#E5E7EB",
            DrawerIcon = "#94A3B8",
            TextPrimary = "#E5E7EB",
            TextSecondary = "#94A3B8",
            ActionDefault = "#CBD5E1",
            LinesDefault = "#334155",
            TableLines = "#334155",
            TableStriped = "rgba(255,255,255,0.03)",
            TableHover = "rgba(96,165,250,0.08)",
            Divider = "#334155",
        },
        Typography = new Typography
        {
            Default = new DefaultTypography { FontFamily = Font, FontSize = ".9375rem", LineHeight = "1.6" },
            H4 = new H4Typography { FontFamily = Font, FontSize = "1.6rem", FontWeight = "700", LineHeight = "1.35" },
            H5 = new H5Typography { FontFamily = Font, FontSize = "1.3rem", FontWeight = "700", LineHeight = "1.4" },
            H6 = new H6Typography { FontFamily = Font, FontSize = "1.05rem", FontWeight = "600", LineHeight = "1.5" },
            Button = new ButtonTypography { FontFamily = Font, FontWeight = "600", TextTransform = "none" },
        },
        LayoutProperties = new LayoutProperties
        {
            DefaultBorderRadius = "8px",
            DrawerWidthRight = "260px",
            DrawerWidthLeft = "260px",
            AppbarHeight = "56px",
        },
    };

    /// <summary>The colors of the chart series (MudChart): primary, secondary, tertiary, info, success, warning of the palette of the mode.</summary>
    public static string[] ChartPalette(bool dark)
    {
        Palette p = dark ? Theme.PaletteDark : Theme.PaletteLight;
        static string Hex(MudBlazor.Utilities.MudColor c) => c.ToString(MudBlazor.Utilities.MudColorOutputFormats.Hex);
        return [Hex(p.Primary), Hex(p.Secondary), Hex(p.Tertiary), Hex(p.Info), Hex(p.Success), Hex(p.Warning)];
    }
}
