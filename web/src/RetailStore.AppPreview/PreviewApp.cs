namespace RetailStore.AppPreview;

/// <summary>Which app the preview shows, chosen on the command line (--app pos | rep): both have their own "/" and
/// "/login", so one process shows one of them.</summary>
public sealed record PreviewApp(bool Rep)
{
    public string Name => Rep ? "rep" : "pos";

    /// <summary>The router of the app (Routes of RetailStore.PosApp or of RetailStore.RepApp).</summary>
    public Type Routes => Rep ? typeof(RetailStore.RepApp.Routes) : typeof(RetailStore.PosApp.Routes);

    /// <summary>The CSS of the app after the kit's, as in its index.html.</summary>
    public string Css => Rep ? "_content/RetailStore.RepApp/repapp.css" : "_content/RetailStore.PosApp/posapp.css";

    public string Title => Rep ? "RetailStore rep app (preview)" : "RetailStore POS (preview)";

    public static PreviewApp FromArgs(IConfiguration config) =>
        new(string.Equals(config["app"], "rep", StringComparison.OrdinalIgnoreCase));
}
