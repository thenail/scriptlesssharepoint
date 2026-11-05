using System.Diagnostics;
using System.IO;
using System.Text.Json.Nodes;
using System.Windows;
using System.Windows.Interop;

using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;

using ScriptlessSharePoint.Bridge;
using ScriptlessSharePoint.Connections;
using ScriptlessSharePoint.Diagnostics;

namespace ScriptlessSharePoint.App;

/// <summary>
/// Hosts the web UI (web/, served from the package folder on a virtual https
/// host) in WebView2 and answers its requests through <see cref="AppBridge"/>.
/// Only the app's own pages get the bridge; other links open in the browser.
/// </summary>
public partial class MainWindow : Window, IHostUi
{
    /// <summary>The web UI's origin: a virtual host mapped to the web\ folder.</summary>
    private const string Host = "app.scriptless-sharepoint.local";

    private readonly AppBridge bridge;
    /// <summary>
    /// The window's handle, got once on the UI thread: MSAL asks for it from its
    /// own threads (sign-in and silent refresh), where WPF objects can't be touched.
    /// </summary>
    private readonly IntPtr handle;
    private Uri origin = new($"https://{Host}/");

    public MainWindow()
    {
        InitializeComponent();
        Log.Init("app");
        Log.Write("app", $"{AppInfo.Name} {AppInfo.Version} starting from {AppContext.BaseDirectory}");
        handle = new WindowInteropHelper(this).EnsureHandle();
        bridge = new AppBridge(new ConnectionStore(), this);
        Loaded += async (_, _) => await StartAsync();
    }

    private async Task StartAsync()
    {
        try
        {
            // WebView2 keeps its data out of the (read-only) package folder
            var env = await CoreWebView2Environment.CreateAsync(null, Path.Combine(AppInfo.DataDir, "WebView2"));
            await Web.EnsureCoreWebView2Async(env);
        }
        catch (WebView2RuntimeNotFoundException)
        {
            if (MessageBox.Show(this,
                    "Scriptless SharePoint needs the Microsoft Edge WebView2 Runtime, which isn't installed. Open its download page?",
                    AppInfo.Name, MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                OpenUrl("https://developer.microsoft.com/microsoft-edge/webview2/");
            }
            Close();
            return;
        }

        var core = Web.CoreWebView2;
        var debug = Debugger.IsAttached || Environment.GetEnvironmentVariable("SSP_DEVTOOLS") == "1";
#if DEBUG
        debug = true;
#endif
        core.Settings.AreDevToolsEnabled = debug;
        core.Settings.AreDefaultContextMenusEnabled = debug;
        core.Settings.AreBrowserAcceleratorKeysEnabled = debug;
        core.Settings.IsStatusBarEnabled = false;
        core.Settings.IsGeneralAutofillEnabled = false;
        core.Settings.IsPasswordAutosaveEnabled = false;

        core.WebMessageReceived += OnWebMessage;
        core.NewWindowRequested += (_, e) =>
        {
            e.Handled = true;
            OpenExternal(e.Uri);
        };
        core.NavigationStarting += (_, e) =>
        {
            if (!IsOwn(e.Uri) && !e.Uri.StartsWith("data:", StringComparison.Ordinal))
            {
                Log.Write("webview", $"navigation to another site blocked and opened in the browser: {new Uri(e.Uri).GetLeftPart(UriPartial.Authority)}");
                e.Cancel = true;
                OpenExternal(e.Uri);
            }
            else if (!e.Uri.StartsWith("data:", StringComparison.Ordinal))
            {
                Log.Write("webview", $"navigating to {e.Uri}");
            }
        };
        core.NavigationCompleted += (_, e) =>
        {
            if (!e.IsSuccess)
            {
                Log.Write("webview", $"navigation failed: {e.WebErrorStatus}");
            }
        };

        // `yarn dev` in web/ serves the UI live: SSP_WEB_URL=http://localhost:5279
        var dev = Environment.GetEnvironmentVariable("SSP_WEB_URL");
        if (Uri.TryCreate(dev, UriKind.Absolute, out var devUri) && devUri.IsLoopback)
        {
            origin = new Uri(devUri.GetLeftPart(UriPartial.Authority) + "/");
            Web.Source = devUri;
            return;
        }
        var folder = Path.Combine(AppContext.BaseDirectory, "web");
        if (!File.Exists(Path.Combine(folder, "index.html")))
        {
            Web.NavigateToString(
                "<body style='font:14px Segoe UI;padding:2em'><h2>The web UI isn't built</h2>" +
                "<p>Run <code>yarn build</code> in <code>web/</code> and build again.</p></body>");
            return;
        }
        core.SetVirtualHostNameToFolderMapping(Host, folder, CoreWebView2HostResourceAccessKind.Deny);
        Web.Source = new Uri(origin, "index.html");
    }

    private async void OnWebMessage(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        if (!IsOwn(e.Source))
        {
            return;
        }
        string? message;
        try
        {
            message = e.TryGetWebMessageAsString();
        }
        catch (ArgumentException)
        {
            return; // not a string: the UI posts JSON text
        }
        var reply = await bridge.HandleAsync(message);
        if (reply is not null)
        {
            Web.CoreWebView2?.PostWebMessageAsString(reply);
        }
    }

    private bool IsOwn(string uri) =>
        Uri.TryCreate(uri, UriKind.Absolute, out var u) &&
        (u.GetLeftPart(UriPartial.Authority) + "/").Equals(origin.AbsoluteUri, StringComparison.OrdinalIgnoreCase);

    private void OpenExternal(string uri)
    {
        if (Uri.TryCreate(uri, UriKind.Absolute, out var u) && u.Scheme == Uri.UriSchemeHttps)
        {
            OpenUrl(u.AbsoluteUri);
        }
    }

    // ---- IHostUi ----

    public IntPtr WindowHandle => handle;

    public string? StartSite => App.StartSite;

    /// <summary>The execution alias when installed from the package, else the server's path next to the app.</summary>
    public string McpCommand
    {
        get
        {
            var dir = AppContext.BaseDirectory;
            if (dir.Contains(@"\WindowsApps\", StringComparison.OrdinalIgnoreCase))
            {
                return AppInfo.McpAlias;
            }
            var sibling = Path.Combine(dir, AppInfo.McpExe);
            return File.Exists(sibling) ? sibling : AppInfo.McpAlias;
        }
    }

    public string? PickFile(string title, string filter) => OnUi(() =>
    {
        var dialog = new OpenFileDialog { Title = title, Filter = filter };
        return dialog.ShowDialog(this) == true ? dialog.FileName : null;
    });

    public string? PickSaveFile(string title, string filter, string fileName) => OnUi(() =>
    {
        var dialog = new SaveFileDialog { Title = title, Filter = filter, FileName = fileName };
        return dialog.ShowDialog(this) == true ? dialog.FileName : null;
    });

    public void OpenUrl(string url)
    {
        try
        {
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true })?.Dispose();
        }
        catch (Exception e) when (e is System.ComponentModel.Win32Exception or InvalidOperationException)
        {
            OnUi(() => MessageBox.Show(this, $"Couldn't open {url}: {e.Message}", AppInfo.Name));
        }
    }

    /// <summary>Runs on the UI thread, whichever thread asks: WPF objects belong to it.</summary>
    private T OnUi<T>(Func<T> work) => Dispatcher.CheckAccess() ? work() : Dispatcher.Invoke(work);

    public void Notify(JsonObject message) =>
        Dispatcher.BeginInvoke(() => Web.CoreWebView2?.PostWebMessageAsString(message.ToJsonString()));
}
