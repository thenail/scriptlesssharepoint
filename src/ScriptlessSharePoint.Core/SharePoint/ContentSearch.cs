namespace ScriptlessSharePoint.SharePoint;

/// <summary>One thing SharePoint search found: a file, a page, a list item, or a list or library.</summary>
public sealed record SearchHit(
    string Path,
    string Title,
    string SiteUrl,
    string WebUrl)
{
    public string? SiteTitle { get; init; }
    public string? FileName { get; init; }
    public long? Size { get; init; }
    public DateTimeOffset? Modified { get; init; }
    public Guid? ListId { get; init; }
    public int? ItemId { get; init; }
    public bool IsDocument { get; init; }
    public string? ContentClass { get; init; }
}

/// <summary>A search hit as the table shows it: its row, the library or list row above it, its subsite's row, and its site collection's.</summary>
public sealed record FoundHit(SpNode Node, SpNode? Parent, SpNode? Web, SpNode Site);

/// <summary>A ready-made search for files, pages or lists across the tenant.</summary>
public sealed record ContentPreset(string Id, string Title, string Description, string Glyph, Func<DateTimeOffset, string> Query)
{
    /// <summary>The managed property results are sorted by, when any.</summary>
    public string? SortBy { get; init; }
    public bool Descending { get; init; }

    /// <summary>The query as it goes to search: SharePoint sites only, OneDrives left out.</summary>
    public string QueryAt(DateTimeOffset now) => $"({Query(now)}) -WebTemplate:SPSPERS";
}

/// <summary>The predefined searches, in the order the navigation pane lists them.</summary>
public static class ContentPresets
{
    /// <summary>How many hits a search ticks at most.</summary>
    public const int Limit = 5_000;

    public static readonly IReadOnlyList<ContentPreset> All =
    [
        new("files-100mb", "Files over 100 MB", "Big files, the biggest first.", "", _ => "IsDocument:1 Size>104857600")
            { SortBy = "Size", Descending = true },
        new("files-1gb", "Files over 1 GB", "Very big files, the biggest first.", "", _ => "IsDocument:1 Size>1073741824")
            { SortBy = "Size", Descending = true },
        new("files-old", "Files unchanged for 3 years", "Files nobody has changed in 3 years or more, the oldest first.", "",
            now => $"IsDocument:1 LastModifiedTime<{now.AddYears(-3):yyyy-MM-dd}") { SortBy = "LastModifiedTime" },
        new("files-guests", "Files guests can open", "Files people outside the organisation can open: shared with them, or in sites they belong to.", "",
            _ => "IsDocument:1 ViewableByExternalUsers:true"),
        new("files-pst", "Mailbox files (.pst)", "Outlook data files, which belong in Exchange rather than SharePoint.", "",
            _ => "FileExtension:pst") { SortBy = "Size", Descending = true },
        new("files-video", "Videos", "Video files, the biggest first.", "",
            _ => "IsDocument:1 (FileExtension:mp4 OR FileExtension:mov OR FileExtension:wmv OR FileExtension:avi OR FileExtension:mkv)")
            { SortBy = "Size", Descending = true },
        new("files-programs", "Programs and scripts", "Files that run code: .exe, .msi, .bat, .cmd, .ps1, .vbs.", "",
            _ => "FileExtension:exe OR FileExtension:msi OR FileExtension:bat OR FileExtension:cmd OR FileExtension:ps1 OR FileExtension:vbs"),
        new("news", "News posts", "Pages posted as news, the newest first.", "", _ => "PromotedState:2")
            { SortBy = "LastModifiedTime", Descending = true },
        new("pages-old", "Pages unchanged for a year", "Site pages nobody has changed in a year or more, the oldest first.", "",
            now => $"contentclass:STS_ListItem_WebPageLibrary LastModifiedTime<{now.AddYears(-1):yyyy-MM-dd}") { SortBy = "LastModifiedTime" },
    ];

    public static ContentPreset? Find(string id) => All.FirstOrDefault(p => p.Id == id);

    /// <summary>
    /// A hit as rows: files, pages and list items under their library or list
    /// (worked out from the address: its first folder below the site), lists
    /// and libraries under their site. Sites and other things search finds: null.
    /// </summary>
    public static FoundHit? Rows(SearchHit h)
    {
        if (!Uri.TryCreate(h.Path, UriKind.Absolute, out var uri) || !Uri.TryCreate(h.WebUrl, UriKind.Absolute, out var webUri))
        {
            return null;
        }
        var siteUrl = h.SiteUrl.TrimEnd('/');
        var webUrl = h.WebUrl.TrimEnd('/');
        var origin = TenantUrls.Resource(webUrl);
        var site = SpNode.FromSite(new SiteSummary(siteUrl, h.SiteTitle ?? SpNode.SiteFromUrl(siteUrl).Title, null, null, null));
        var web = string.Equals(siteUrl, webUrl, StringComparison.OrdinalIgnoreCase)
            ? null
            : new SpNode { Kind = NodeKind.Subsite, Title = webUri.AbsolutePath.TrimEnd('/').Split('/')[^1], Url = webUrl, WebUrl = webUrl, SiteUrl = siteUrl };
        var path = Uri.UnescapeDataString(uri.AbsolutePath);
        var webPath = Uri.UnescapeDataString(webUri.AbsolutePath).TrimEnd('/');
        var cc = h.ContentClass ?? "";

        if (cc.StartsWith("STS_List_", StringComparison.OrdinalIgnoreCase))
        {
            if (h.ListId is null)
            {
                return null;
            }
            var library = cc.EndsWith("Library", StringComparison.OrdinalIgnoreCase);
            var forms = path.IndexOf("/Forms/", StringComparison.OrdinalIgnoreCase);
            var root = forms > 0 ? path[..forms] : path.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase) ? path[..path.LastIndexOf('/')] : path;
            var list = new SpNode
            {
                Kind = library ? NodeKind.Library : NodeKind.List,
                Title = h.Title,
                Url = origin + root,
                WebUrl = webUrl,
                SiteUrl = siteUrl,
                ListId = h.ListId,
                ServerRelativeUrl = root,
            };
            return new FoundHit(list, web, web, site);
        }

        var file = h.IsDocument || cc is "STS_ListItem_DocumentLibrary" or "STS_ListItem_WebPageLibrary" || cc.EndsWith("Library", StringComparison.OrdinalIgnoreCase) && cc.StartsWith("STS_ListItem_", StringComparison.OrdinalIgnoreCase);
        var item = !file && cc.StartsWith("STS_ListItem_", StringComparison.OrdinalIgnoreCase);
        if ((!file && !item) || h.ListId is null || h.ItemId is null)
        {
            return null;
        }
        var parent = ListAbove(path, webPath, origin, webUrl, siteUrl, h.ListId.Value, file) ?? web;
        var node = file
            ? new SpNode
            {
                Kind = NodeKind.File,
                Title = h.FileName ?? path[(path.LastIndexOf('/') + 1)..],
                Url = origin + path,
                WebUrl = webUrl,
                SiteUrl = siteUrl,
                ListId = h.ListId,
                ItemId = h.ItemId,
                ServerRelativeUrl = path,
                InLibrary = true,
                Size = h.Size,
                Modified = h.Modified,
            }
            : new SpNode
            {
                Kind = NodeKind.Item,
                Title = h.Title,
                Url = h.Path,
                WebUrl = webUrl,
                SiteUrl = siteUrl,
                ListId = h.ListId,
                ItemId = h.ItemId,
                Modified = h.Modified,
            };
        return new FoundHit(node, parent, web, site);
    }

    /// <summary>The library or list a file or item is in, from its address: "Shared Documents", or "Lists/Tasks".</summary>
    private static SpNode? ListAbove(string path, string webPath, string origin, string webUrl, string siteUrl, Guid listId, bool library)
    {
        if (!path.StartsWith(webPath + "/", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }
        var parts = path[(webPath.Length + 1)..].Split('/');
        if (parts.Length < 2)
        {
            return null;
        }
        var leaf = string.Equals(parts[0], "Lists", StringComparison.OrdinalIgnoreCase) && parts.Length > 2 ? $"{parts[0]}/{parts[1]}" : parts[0];
        var root = $"{webPath}/{leaf}";
        return new SpNode
        {
            Kind = library ? NodeKind.Library : NodeKind.List,
            Title = leaf[(leaf.LastIndexOf('/') + 1)..],
            Url = origin + root,
            WebUrl = webUrl,
            SiteUrl = siteUrl,
            ListId = listId,
            ServerRelativeUrl = root,
            InLibrary = library,
        };
    }
}
