using System.Windows;

namespace ScriptlessSharePoint.App;

public partial class App : Application
{
    /// <summary>A site to show at start (--site URL), e.g. from the MCP server's open_app.</summary>
    public static string? StartSite { get; private set; }

    protected override void OnStartup(StartupEventArgs e)
    {
        for (var i = 0; i < e.Args.Length - 1; i++)
        {
            if (e.Args[i] == "--site")
            {
                StartSite = e.Args[i + 1];
            }
        }
        base.OnStartup(e);
    }
}
