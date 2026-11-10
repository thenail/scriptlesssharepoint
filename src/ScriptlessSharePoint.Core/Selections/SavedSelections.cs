using System.Text.Json;
using System.Text.Json.Serialization;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Selections;

/// <summary>
/// Ticked things saved under a name, to come back to: for one connection, the
/// items, the site collections they're in, and the row each was ticked under
/// (so the table can show them again without loading what's around them).
/// <see cref="File"/> is the file it's kept in, which also says which one it is.
/// </summary>
public sealed record SavedSelection
{
    [JsonIgnore]
    public string File { get; init; } = "";

    public required string Name { get; init; }
    public required string ConnectionId { get; init; }
    public DateTimeOffset Saved { get; init; }
    public IReadOnlyList<SpNode> Items { get; init; } = [];

    /// <summary>The site collections the items are in, as the table showed them.</summary>
    public IReadOnlyList<SpNode> Sites { get; init; } = [];

    /// <summary>The subsite, list or folder each item was ticked in, shown above it.</summary>
    public IReadOnlyList<SpNode> Parents { get; init; } = [];
}

/// <summary>
/// The saved selections: one JSON file each in <see cref="Folder"/>
/// (%LOCALAPPDATA%\ScriptlessSharePoint\Selections). Removing one removes
/// the list, never anything in SharePoint.
/// </summary>
public sealed class SavedSelectionStore(string? folder = null)
{
    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() },
    };

    public string Folder { get; } = folder ?? Path.Combine(AppInfo.DataDir, "Selections");

    /// <summary>Every saved selection, the oldest first; files that can't be read are left out.</summary>
    public IReadOnlyList<SavedSelection> List()
    {
        if (!Directory.Exists(Folder))
        {
            return [];
        }
        return Directory.EnumerateFiles(Folder, "*.json")
            .Select(Load)
            .OfType<SavedSelection>()
            .OrderBy(s => s.Saved)
            .ThenBy(s => s.Name, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    public SavedSelection? Load(string file)
    {
        try
        {
            using var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return JsonSerializer.Deserialize<SavedSelection>(stream, Json) is { } s ? s with { File = file } : null;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or JsonException)
        {
            return null;
        }
    }

    /// <summary>A new saved selection, in a file of its own.</summary>
    public SavedSelection Save(string name, string connectionId, IEnumerable<SpNode> items, IEnumerable<SpNode> sites, IEnumerable<SpNode>? parents = null)
    {
        var saved = new SavedSelection
        {
            Name = name.Trim(),
            ConnectionId = connectionId,
            Saved = DateTimeOffset.Now,
            Items = items.DistinctBy(n => n.Key).ToList(),
            Sites = sites.DistinctBy(n => n.Key).ToList(),
            Parents = (parents ?? []).DistinctBy(n => n.Key).ToList(),
        };
        var file = Path.Combine(Folder, $"{saved.Saved:yyyyMMdd-HHmmss}-{Guid.NewGuid().ToString("n")[..8]}.json");
        Write(saved with { File = file });
        return saved with { File = file };
    }

    /// <summary>
    /// Adds items to a saved selection, with the site collections and parents
    /// they're shown under; what's in it already stays as it is. Null when the
    /// selection is gone.
    /// </summary>
    public SavedSelection? Add(string file, IEnumerable<SpNode> items, IEnumerable<SpNode> sites, IEnumerable<SpNode> parents)
    {
        if (Load(file) is not { } saved)
        {
            return null;
        }
        var added = saved with
        {
            Items = saved.Items.Concat(items).DistinctBy(n => n.Key).ToList(),
            Sites = saved.Sites.Concat(sites).DistinctBy(n => n.Key).ToList(),
            Parents = saved.Parents.Concat(parents).DistinctBy(n => n.Key).ToList(),
        };
        Write(added);
        return added;
    }

    /// <summary>
    /// Replaces what a saved selection holds (its name and place in the list
    /// stay). Null when the selection is gone.
    /// </summary>
    public SavedSelection? Update(string file, IEnumerable<SpNode> items, IEnumerable<SpNode> sites, IEnumerable<SpNode> parents)
    {
        if (Load(file) is not { } saved)
        {
            return null;
        }
        var updated = saved with
        {
            Items = items.DistinctBy(n => n.Key).ToList(),
            Sites = sites.DistinctBy(n => n.Key).ToList(),
            Parents = parents.DistinctBy(n => n.Key).ToList(),
        };
        Write(updated);
        return updated;
    }

    public SavedSelection? Rename(string file, string name)
    {
        if (Load(file) is not { } saved)
        {
            return null;
        }
        var renamed = saved with { Name = name.Trim() };
        Write(renamed);
        return renamed;
    }

    public void Delete(string file)
    {
        if (System.IO.File.Exists(file) && Path.GetDirectoryName(Path.GetFullPath(file)) == Path.GetFullPath(Folder).TrimEnd('\\'))
        {
            System.IO.File.Delete(file);
        }
    }

    /// <summary>Removes a connection's saved selections, when the connection goes.</summary>
    public int DeleteFor(string connectionId)
    {
        var mine = List().Where(s => s.ConnectionId == connectionId).ToList();
        foreach (var s in mine)
        {
            Delete(s.File);
        }
        return mine.Count;
    }

    /// <summary>Through a temporary file, so a reader never sees half a list.</summary>
    private void Write(SavedSelection saved)
    {
        Directory.CreateDirectory(Folder);
        var temp = saved.File + ".tmp";
        System.IO.File.WriteAllText(temp, JsonSerializer.Serialize(saved, Json));
        System.IO.File.Move(temp, saved.File, overwrite: true);
    }
}
