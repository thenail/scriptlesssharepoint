using System.Reflection;

namespace ScriptlessSharePoint;

/// <summary>Names, version and folders shared by the app and the MCP server.</summary>
public static class AppInfo
{
    public const string Name = "Scriptless SharePoint";

    /// <summary>The app's executable in the package, and its execution alias.</summary>
    public const string AppExe = "ScriptlessSharePoint.exe";
    public const string AppAlias = "scriptless-sharepoint";

    /// <summary>The MCP server's executable in the package, and its execution alias.</summary>
    public const string McpExe = "scriptless-sharepoint-mcp.exe";
    public const string McpAlias = "scriptless-sharepoint-mcp";

    public static string Version { get; } =
        (typeof(AppInfo).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "0.0.0")
            .Split('+')[0];

    /// <summary>
    /// The built-in Entra app (client) ID, set at build time (SspClientId); null when
    /// the build has none and every connection brings its own.
    /// </summary>
    public static string? BuiltInClientId { get; } = ReadClientId();

    /// <summary>
    /// Where connections and the token cache live. In the MSIX package Windows
    /// redirects this to the package's own LocalCache, the same for both executables.
    /// </summary>
    public static string DataDir { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ScriptlessSharePoint");

    private static string? ReadClientId()
    {
        var value = typeof(AppInfo).Assembly
            .GetCustomAttributes<AssemblyMetadataAttribute>()
            .FirstOrDefault(a => a.Key == "SspClientId")?.Value;
        return Guid.TryParse(value, out var id) ? id.ToString() : null;
    }
}
