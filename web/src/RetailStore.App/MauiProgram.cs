using RetailStore.RepApp;

namespace RetailStore.App;

/// <summary>The rep app on Android (docs/78): the screens of RetailStore.RepApp in a BlazorWebView, the SQLite file in
/// the app's own folder, and the phone's services (secure storage, network, Bluetooth printer).</summary>
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
        builder.Services.AddRepApp(Path.Combine(FileSystem.AppDataDirectory, "rep.db"));
        builder.Services.AddSingleton<IRepPlatform, PhonePlatform>();
        return builder.Build();
    }
}
