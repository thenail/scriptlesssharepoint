using System.Text.Json.Serialization;

namespace ScriptlessSharePoint.SharePoint;

/// <summary>What a row in the table is, and so which commands apply to it.</summary>
public enum NodeKind
{
    /// <summary>The tenant itself: the connection's admin centre (no row; the pane's Tenant commands).</summary>
    Tenant,

    /// <summary>A site collection, with its root web.</summary>
    SiteCollection,

    /// <summary>A web below a site collection's root.</summary>
    Subsite,

    /// <summary>A list that isn't a document library.</summary>
    List,

    /// <summary>A document library.</summary>
    Library,

    /// <summary>A folder in a list or library.</summary>
    Folder,

    /// <summary>A file in a library.</summary>
    File,

    /// <summary>An item in a list.</summary>
    Item,
}

/// <summary>
/// Something in SharePoint the table shows and commands act on. <see cref="Url"/>
/// is absolute; <see cref="WebUrl"/> is the web it lives in (itself for site
/// collections and subsites), and lists, folders, files and items carry their
/// list's ID and their server-relative URL (decoded, as SharePoint's FileRef).
/// </summary>
public sealed record SpNode
{
    public required NodeKind Kind { get; init; }
    public required string Title { get; init; }
    public required string Url { get; init; }
    public required string WebUrl { get; init; }

    /// <summary>The site collection's URL, when known.</summary>
    public string? SiteUrl { get; init; }

    public Guid? ListId { get; init; }

    /// <summary>The list item's ID: folders, files and items.</summary>
    public int? ItemId { get; init; }

    /// <summary>Lists and libraries: their root folder; folders and files: themselves.</summary>
    public string? ServerRelativeUrl { get; init; }

    /// <summary>Folders and items: whether their list is a document library.</summary>
    public bool InLibrary { get; init; }

    /// <summary>The web template ("STS#3") or the list's template number.</summary>
    public string? Template { get; init; }

    /// <summary>Items in a list, or in a folder when SharePoint says.</summary>
    public int? ItemCount { get; init; }

    /// <summary>A file's size in bytes.</summary>
    public long? Size { get; init; }

    public DateTimeOffset? Modified { get; init; }

    /// <summary>Who has a file checked out, if anyone.</summary>
    public string? CheckedOutTo { get; init; }

    /// <summary>Whether the row can be opened to show what's in it.</summary>
    [JsonIgnore]
    public bool CanExpand => Kind is NodeKind.SiteCollection or NodeKind.Subsite or NodeKind.List or NodeKind.Library or NodeKind.Folder;

    /// <summary>The same thing wherever it shows up: selections are kept by it.</summary>
    [JsonIgnore]
    public string Key => Kind is NodeKind.Item ? $"{Kind}:{ListId}:{ItemId}" : $"{Kind}:{Url.TrimEnd('/').ToLowerInvariant()}";

    /// <summary>The file or folder name: the last part of the server-relative URL.</summary>
    [JsonIgnore]
    public string LeafName => ServerRelativeUrl is { } u ? u.TrimEnd('/')[(u.TrimEnd('/').LastIndexOf('/') + 1)..] : Title;

    /// <summary>The folder a folder or file is in (server-relative).</summary>
    [JsonIgnore]
    public string? ParentFolder => ServerRelativeUrl is { } u && u.TrimEnd('/').LastIndexOf('/') is > 0 and var i ? u[..i] : null;

    /// <summary>
    /// Whether this row is on the way down to <paramref name="target"/>: the
    /// target itself, or the subsite, list or folder it's in. Opening a saved
    /// selection follows these down from its site collection.
    /// </summary>
    public bool Leads(SpNode target) => Key == target.Key || Kind switch
    {
        NodeKind.SiteCollection => SameUrl(Url, target.SiteUrl),
        NodeKind.Subsite => Within(target.WebUrl, WebUrl),
        NodeKind.List or NodeKind.Library => ListId is { } id && id == target.ListId,
        NodeKind.Folder => ListId is { } id && id == target.ListId && ServerRelativeUrl is { } folder &&
                           target.ServerRelativeUrl?.StartsWith(folder.TrimEnd('/') + "/", StringComparison.OrdinalIgnoreCase) == true,
        _ => false,
    };

    /// <summary>A site collection's row for a URL, when only the URL is known.</summary>
    public static SpNode SiteFromUrl(string url)
    {
        var u = url.TrimEnd('/');
        var path = Uri.TryCreate(u, UriKind.Absolute, out var uri) ? uri.AbsolutePath.TrimEnd('/') : "";
        return new SpNode { Kind = NodeKind.SiteCollection, Title = path.Length == 0 ? u : path, Url = u, WebUrl = u, SiteUrl = u };
    }

    private static bool SameUrl(string? a, string? b) =>
        a is not null && b is not null && string.Equals(a.TrimEnd('/'), b.TrimEnd('/'), StringComparison.OrdinalIgnoreCase);

    private static bool Within(string? url, string? web) =>
        SameUrl(url, web) || (url is not null && web is not null && url.StartsWith(web.TrimEnd('/') + "/", StringComparison.OrdinalIgnoreCase));

    public static SpNode FromSite(SiteSummary s) => new()
    {
        Kind = NodeKind.SiteCollection,
        Title = s.Title,
        Url = s.Url,
        WebUrl = s.Url,
        SiteUrl = s.Url,
        Template = s.Template,
        Modified = s.LastModified,
    };

    public static SpNode Tenant(string adminUrl, string title) => new()
    {
        Kind = NodeKind.Tenant,
        Title = title,
        Url = adminUrl,
        WebUrl = adminUrl,
    };
}

/// <summary>A page of a folder's contents; <see cref="Next"/> is SharePoint's paging info for the next, or null.</summary>
public sealed record NodePage(IReadOnlyList<SpNode> Nodes, string? Next);

public static class NodeKinds
{
    /// <summary>"site collection", "libraries"…: for sentences like "Run on 3 libraries".</summary>
    public static string Name(NodeKind kind, int count = 1) => (kind, count == 1) switch
    {
        (NodeKind.Tenant, _) => "tenant",
        (NodeKind.SiteCollection, true) => "site collection",
        (NodeKind.SiteCollection, false) => "site collections",
        (NodeKind.Subsite, true) => "subsite",
        (NodeKind.Subsite, false) => "subsites",
        (NodeKind.List, true) => "list",
        (NodeKind.List, false) => "lists",
        (NodeKind.Library, true) => "library",
        (NodeKind.Library, false) => "libraries",
        (NodeKind.Folder, true) => "folder",
        (NodeKind.Folder, false) => "folders",
        (NodeKind.File, true) => "file",
        (NodeKind.File, false) => "files",
        (NodeKind.Item, true) => "list item",
        (NodeKind.Item, false) => "list items",
        _ => kind.ToString().ToLowerInvariant(),
    };

    /// <summary>"Site collections", "Libraries": headings.</summary>
    public static string Heading(NodeKind kind)
    {
        var n = Name(kind, 2);
        return char.ToUpperInvariant(n[0]) + n[1..];
    }

    /// <summary>"3 libraries", "1 file".</summary>
    public static string Count(NodeKind kind, int count) => $"{count:N0} {Name(kind, count)}";
}
