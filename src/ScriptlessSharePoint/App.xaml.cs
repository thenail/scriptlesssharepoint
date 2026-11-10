using Microsoft.UI.Xaml;

using ScriptlessSharePoint.Diagnostics;

namespace ScriptlessSharePoint.App;

public partial class App : Application
{
    private Window? window;

    public App()
    {
        InitializeComponent();
        UnhandledException += (_, e) => Log.Write("app", $"unhandled: {e.Exception}");
    }

    /// <summary>A site to show at start (--site URL), e.g. from the MCP server's open_app.</summary>
    public static string? StartSite { get; private set; }

    /// <summary>A plan to show at start (--plan ID), from the MCP server's open_app.</summary>
    public static string? StartPlan { get; private set; }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        var argv = Environment.GetCommandLineArgs();
        for (var i = 0; i < argv.Length - 1; i++)
        {
            if (argv[i] == "--site")
            {
                StartSite = argv[i + 1];
            }
            else if (argv[i] == "--plan")
            {
                StartPlan = argv[i + 1];
            }
        }
        window = new MainWindow();
        window.Activate();
    }
}
