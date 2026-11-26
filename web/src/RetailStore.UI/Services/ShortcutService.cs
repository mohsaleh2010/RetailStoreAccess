using Microsoft.JSInterop;

namespace RetailStore.UI.Services;

/// <summary>The keyboard shortcuts of every screen (docs/ERP-Redesign-Prompt.md, phase 3).</summary>
public enum ErpKey
{
    /// <summary>F2: save.</summary>
    Save,
    /// <summary>F3: search (without a handler: the search box of the page, data-erp-search).</summary>
    Search,
    /// <summary>F4: new.</summary>
    New,
    /// <summary>Esc: close (an open dialog closes itself first: MudDialogProvider CloseOnEscapeKey).</summary>
    Close,
    /// <summary>Ctrl+P: print (without a handler: the print of the browser).</summary>
    Print,
}

/// <summary>One listener of the browser per session (erp.js, erp.shortcuts) and a stack of handlers: the last one
/// registered for a key runs, so an open dialog takes F2 from the page under it and gives it back when it closes.
/// Pages use the KeyboardShortcuts component; a component may call RegisterAsync itself (after its first render).</summary>
public sealed class ShortcutService(IJSRuntime js) : IAsyncDisposable
{
    private readonly List<Registration> stack = [];
    private DotNetObjectReference<ShortcutService>? self;

    /// <summary>The name erp.js gives a key.</summary>
    public static string KeyName(ErpKey key) => key switch
    {
        ErpKey.Save => "F2",
        ErpKey.Search => "F3",
        ErpKey.New => "F4",
        ErpKey.Close => "Escape",
        _ => "Ctrl+P",
    };

    /// <summary>Registers a handler for a key; disposing the result removes it.</summary>
    public async Task<IAsyncDisposable> RegisterAsync(ErpKey key, Func<Task> run)
    {
        var registration = new Registration(this, key, run);
        stack.Add(registration);
        await SyncAsync();
        return registration;
    }

    /// <summary>The handler of the key, or null: the last one registered.</summary>
    public Func<Task>? Handler(ErpKey key) => stack.LastOrDefault(r => r.Key == key)?.Run;

    /// <summary>Called by erp.js for a key that has a handler.</summary>
    [JSInvokable]
    public async Task OnKey(string name)
    {
        foreach (ErpKey key in Enum.GetValues<ErpKey>())
        {
            if (KeyName(key) == name && Handler(key) is { } run)
            {
                await run();
                return;
            }
        }
    }

    private async Task SyncAsync()
    {
        try
        {
            self ??= DotNetObjectReference.Create(this);
            var keys = stack.Select(r => KeyName(r.Key)).Distinct().ToArray();
            await js.InvokeVoidAsync("erp.shortcuts.attach", self, keys);
        }
        catch (JSDisconnectedException)
        {
            // the browser left: nothing to tell it
        }
        catch (InvalidOperationException)
        {
            // prerendering: no browser yet (components register after their first render)
        }
    }

    private async Task RemoveAsync(Registration registration)
    {
        if (stack.Remove(registration))
        {
            await SyncAsync();
        }
    }

    public async ValueTask DisposeAsync()
    {
        stack.Clear();
        if (self is not null)
        {
            try
            {
                await js.InvokeVoidAsync("erp.shortcuts.detach");
            }
            catch (JSDisconnectedException)
            {
            }
            catch (InvalidOperationException)
            {
            }
            self.Dispose();
        }
    }

    private sealed class Registration(ShortcutService owner, ErpKey key, Func<Task> run) : IAsyncDisposable
    {
        public ErpKey Key { get; } = key;
        public Func<Task> Run { get; } = run;

        public ValueTask DisposeAsync() => new(owner.RemoveAsync(this));
    }
}
