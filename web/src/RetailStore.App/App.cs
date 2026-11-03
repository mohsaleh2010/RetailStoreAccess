using Microsoft.AspNetCore.Components.WebView.Maui;
using RetailStore.RepApp;

namespace RetailStore.App;

public sealed class App : Application
{
    protected override Window CreateWindow(IActivationState? activationState) => new(new MainPage()) { Title = "RetailStore" };
}

/// <summary>One page: the Blazor screens of the rep app (Routes of RetailStore.RepApp).</summary>
public sealed class MainPage : ContentPage
{
    public MainPage()
    {
        var view = new BlazorWebView { HostPage = "wwwroot/index.html" };
        view.RootComponents.Add(new RootComponent { Selector = "#app", ComponentType = typeof(Routes) });
        Content = view;
    }
}
