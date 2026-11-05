using System.Text.Json.Nodes;

namespace ScriptlessSharePoint.Bridge;

/// <summary>What the bridge needs from the window hosting the web UI.</summary>
public interface IHostUi
{
    /// <summary>
    /// The window sign-in dialogs belong to. MSAL reads it from its own threads,
    /// so it must be safe to read from any thread (a cached handle).
    /// </summary>
    IntPtr WindowHandle { get; }

    /// <summary>A site to show at start (--site URL), or null.</summary>
    string? StartSite { get; }

    /// <summary>How AI apps start the MCP server: the alias when installed, else the exe's path.</summary>
    string McpCommand { get; }

    /// <summary>A file to open, or null if the user cancelled.</summary>
    string? PickFile(string title, string filter);

    /// <summary>Where to save a file, or null if the user cancelled.</summary>
    string? PickSaveFile(string title, string filter, string fileName);

    /// <summary>Opens an https:// URL in the default browser.</summary>
    void OpenUrl(string url);

    /// <summary>Sends a notification (a JSON-RPC message without an ID) to the web UI.</summary>
    void Notify(JsonObject message);
}
