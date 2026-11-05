using System.Text;

using Microsoft.Identity.Client;

namespace ScriptlessSharePoint.Diagnostics;

/// <summary>
/// Diagnostics in <see cref="AppInfo.DataDir"/>\logs: sign-in (MSAL, without
/// personal data) and the app's operations (names, times and errors; never
/// arguments, tokens or SharePoint content). One file per process kind, kept small.
/// </summary>
public static class Log
{
    private const long MaxBytes = 2 * 1024 * 1024;
    private static readonly object Gate = new();
    private static string? name;

    public static string Folder { get; } = Path.Combine(AppInfo.DataDir, "logs");

    public static string FilePath => Path.Combine(Folder, $"{name ?? "app"}.log");

    /// <summary>Turns logging on for this process ("app" or "mcp"); until then (e.g. in tests) nothing is written.</summary>
    public static void Init(string processName) => name = processName;

    public static void Write(string area, string message)
    {
        if (name is null)
        {
            return;
        }
        try
        {
            lock (Gate)
            {
                Directory.CreateDirectory(Folder);
                var file = new FileInfo(FilePath);
                if (file.Exists && file.Length > MaxBytes)
                {
                    File.Move(file.FullName, file.FullName + ".old", overwrite: true);
                }
                File.AppendAllText(
                    FilePath,
                    $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss.fff} [{Environment.ProcessId}] {area}: {message}{Environment.NewLine}",
                    Encoding.UTF8);
            }
        }
        catch (IOException)
        {
            // diagnostics never break the app
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    /// <summary>MSAL's log callback (PII logging is off, so no user names or tokens).</summary>
    public static void Msal(LogLevel level, string message, bool containsPii)
    {
        if (!containsPii)
        {
            Write($"msal {level}", message);
        }
    }
}
