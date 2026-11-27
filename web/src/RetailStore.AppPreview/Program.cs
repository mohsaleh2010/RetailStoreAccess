using MudBlazor.Services;
using RetailStore.AppPreview;
using RetailStore.AppPreview.Components;
using RetailStore.PosApp;
using RetailStore.RepApp;

// For development only (docs/ui-inventory.md section 17): the POS or the rep app in a browser. The services are those
// of the MAUI programs (AddPosApp / AddRepApp, then AddMudServices) with the platforms of the tests.
var builder = WebApplication.CreateBuilder(args);
var app = PreviewApp.FromArgs(builder.Configuration);
var data = Path.Combine(builder.Environment.ContentRootPath, "App_Data");
Directory.CreateDirectory(data);

builder.Services.AddSingleton(app);
if (app.Rep)
{
    builder.Services.AddRepApp(Path.Combine(data, "preview-rep.db"));
    builder.Services.AddSingleton<IRepPlatform>(new BasicPlatform { DeviceName = "Preview" });
}
else
{
    builder.Services.AddPosApp(Path.Combine(data, "preview-pos.db"));
    builder.Services.AddSingleton<IPosPlatform>(new BasicPosPlatform { DeviceName = "PREVIEW-PC" });
}
builder.Services.AddMudServices();
builder.Services.AddRazorComponents().AddInteractiveServerComponents();
if (string.IsNullOrEmpty(builder.Configuration["urls"]) && string.IsNullOrEmpty(builder.Configuration["ASPNETCORE_URLS"]))
{
    builder.WebHost.UseUrls(app.Rep ? "http://localhost:5151" : "http://localhost:5150");
}

var web = builder.Build();
web.UseAntiforgery();
web.MapStaticAssets();
web.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode()
    .AddAdditionalAssemblies(app.Routes.Assembly);
web.Run();
