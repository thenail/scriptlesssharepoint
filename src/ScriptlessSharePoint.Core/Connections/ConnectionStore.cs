using System.Text.Json;
using System.Text.Json.Serialization;

namespace ScriptlessSharePoint.Connections;

/// <summary>The saved connections and which one is active.</summary>
public sealed class ConnectionData
{
    public string? ActiveId { get; set; }

    public List<ConnectionProfile> Connections { get; set; } = [];

    [JsonIgnore]
    public ConnectionProfile? Active =>
        Connections.FirstOrDefault(c => c.Id == ActiveId) ?? Connections.FirstOrDefault();

    /// <summary>A connection by name or ID (case-insensitive); null finds the active one.</summary>
    public ConnectionProfile? Find(string? nameOrId) =>
        string.IsNullOrWhiteSpace(nameOrId)
            ? Active
            : Connections.FirstOrDefault(c =>
                string.Equals(c.Id, nameOrId, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(c.Name, nameOrId.Trim(), StringComparison.OrdinalIgnoreCase));
}

/// <summary>
/// connections.json in <see cref="AppInfo.DataDir"/>. The app writes it; the MCP
/// server reads it on every call, so a connection made in the app is there for AI apps.
/// Secrets in it are DPAPI-encrypted (<see cref="Secrets"/>).
/// </summary>
public sealed class ConnectionStore
{
    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    public ConnectionStore(string? directory = null)
    {
        FilePath = Path.Combine(directory ?? AppInfo.DataDir, "connections.json");
    }

    public string FilePath { get; }

    public ConnectionData Load()
    {
        try
        {
            if (!File.Exists(FilePath))
            {
                return new ConnectionData();
            }
            using var stream = new FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return JsonSerializer.Deserialize<ConnectionData>(stream, Json) ?? new ConnectionData();
        }
        catch (JsonException)
        {
            return new ConnectionData();
        }
    }

    public void Save(ConnectionData data)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        var temp = FilePath + ".tmp";
        File.WriteAllText(temp, JsonSerializer.Serialize(data, Json));
        File.Move(temp, FilePath, overwrite: true);
    }

    /// <summary>Loads, changes and saves in one go.</summary>
    public ConnectionData Update(Action<ConnectionData> change)
    {
        var data = Load();
        change(data);
        Save(data);
        return data;
    }
}
