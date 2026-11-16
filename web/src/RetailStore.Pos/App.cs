using Microsoft.AspNetCore.Components.WebView.Maui;
using RetailStore.PosApp;

namespace RetailStore.Pos;

public sealed class App : Application
{
    protected override Window CreateWindow(IActivationState? activationState) =>
        new(new MainPage()) { Title = "RetailStore POS", Width = 1280, Height = 800 };
}

/// <summary>One page: the Blazor screens of the POS (Routes of RetailStore.PosApp).</summary>
public sealed class MainPage : ContentPage
{
    public MainPage()
    {
        var view = new BlazorWebView { HostPage = "wwwroot/index.html" };
        view.RootComponents.Add(new RootComponent { Selector = "#app", ComponentType = typeof(Routes) });
        Content = view;
    }
}
