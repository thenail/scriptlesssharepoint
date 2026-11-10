using System.Text.Json;

namespace ScriptlessSharePoint.Settings;

/// <summary>The app's own preferences.</summary>
public sealed record AppSettings
{
    /// <summary>
    /// Ask before every command that changes SharePoint, not only those that
    /// delete or run on many targets (which always ask).
    /// </summary>
    public bool ConfirmChanges { get; init; }
}

/// <summary>
/// The preferences in settings.json in <see cref="AppInfo.DataDir"/>. A
/// missing or unreadable file gives the defaults.
/// </summary>
public sealed class SettingsStore(string? folder = null)
{
    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    public string FilePath { get; } = Path.Combine(folder ?? AppInfo.DataDir, "settings.json");

    public AppSettings Load()
    {
        try
        {
            return File.Exists(FilePath) ? JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(FilePath), Json) ?? new() : new();
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or JsonException)
        {
            return new();
        }
    }

    public void Save(AppSettings settings)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        var temp = FilePath + ".tmp";
        File.WriteAllText(temp, JsonSerializer.Serialize(settings, Json));
        File.Move(temp, FilePath, overwrite: true);
    }
}
