namespace RetailStore.Pos.WinUI;

/// <summary>The WinUI entry of the POS program.</summary>
public partial class App : MauiWinUIApplication
{
    public App()
    {
        InitializeComponent();
    }

    protected override MauiApp CreateMauiApp() => MauiProgram.CreateMauiApp();
}
