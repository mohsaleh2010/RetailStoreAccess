using RetailStore.PosApp;

namespace RetailStore.Pos;

/// <summary>The point of sale on Windows (docs/83): the screens of RetailStore.PosApp in a BlazorWebView, the SQLite file
/// of the PC in the program's own folder of the Windows user, and the services of Windows (protected storage, printers).</summary>
public static class MauiProgram
{
    public static MauiApp CreateMauiApp()
    {
        var builder = MauiApp.CreateBuilder();
        builder.UseMauiApp<App>();
        builder.Services.AddMauiBlazorWebView();
#if DEBUG
        builder.Services.AddBlazorWebViewDeveloperTools();
#endif
        builder.Services.AddPosApp(Path.Combine(FileSystem.AppDataDirectory, "pos.db"));
        builder.Services.AddSingleton<IPosPlatform, WindowsPlatform>();
        return builder.Build();
    }
}
