using System.Globalization;

using Microsoft.SharePoint.Client;

namespace ScriptlessSharePoint.SharePoint;

/// <summary>What's in a site, list or folder, for the table's rows.</summary>
public sealed partial class SharePointClient
{
    /// <summary>How many rows of a list or folder a page brings.</summary>
    public const int FolderPageSize = 500;

    /// <summary>
    /// What a site collection or subsite holds: its subsites, then its libraries
    /// and lists (not the hidden ones), each by title.
    /// </summary>
    public Task<IReadOnlyList<SpNode>> GetWebContentsAsync(SpNode web, CancellationToken ct = default) => RetryAsync<IReadOnlyList<SpNode>>(async () =>
    {
        using var ctx = await ContextAsync(web.WebUrl, ct).ConfigureAwait(false);
        var w = ctx.Web;
        ctx.Load(ctx.Site, s => s.Url);
        ctx.Load(w, x => x.Url);
        var lists = ctx.LoadQuery(w.Lists.Include(
            l => l.Id, l => l.Title, l => l.BaseType, l => l.BaseTemplate, l => l.Hidden,
            l => l.ItemCount, l => l.LastItemModifiedDate, l => l.RootFolder.ServerRelativeUrl));
        await ctx.ExecuteQueryAsync().ConfigureAwait(false);

        var siteUrl = ctx.Site.Url.TrimEnd('/');
        var webUrl = w.Url.TrimEnd('/');
        var resource = TenantUrls.Resource(webUrl);
        var nodes = new List<SpNode>();

        // subsites: all of them for administrators and apps, else the ones this user can open
        var subwebs = await LoadSubwebsAsync(ctx, w).ConfigureAwait(false);
        nodes.AddRange(subwebs
            .Select(s => new SpNode
            {
                Kind = NodeKind.Subsite,
                Title = string.IsNullOrWhiteSpace(s.Title) ? s.Url : s.Title,
                Url = s.Url.TrimEnd('/'),
                WebUrl = s.Url.TrimEnd('/'),
                SiteUrl = siteUrl,
                Template = $"{s.WebTemplate}#{s.Configuration}",
                Modified = Utc(s.LastItemModifiedDate),
            })
            .OrderBy(n => n.Title, StringComparer.CurrentCultureIgnoreCase));

        var visible = lists.Where(l => !l.Hidden).Select(l => ListNode(l, siteUrl, webUrl, resource));
        nodes.AddRange(visible
            .OrderBy(n => n.Kind == NodeKind.Library ? 0 : 1)
            .ThenBy(n => n.Title, StringComparer.CurrentCultureIgnoreCase));
        return nodes;
    }, ct);

    /// <summary>
    /// A site's home page (the web's welcome page) and the library it's in, or
    /// null when it has none that's a file in a library (a classic site's
    /// default.aspx at the web's root).
    /// </summary>
    public Task<FoundNode?> GetHomePageAsync(SpNode web, CancellationToken ct = default) => RetryAsync<FoundNode?>(async () =>
    {
        using var ctx = await ContextAsync(web.WebUrl, ct).ConfigureAwait(false);
        ctx.Load(ctx.Site, s => s.Url);
        ctx.Load(ctx.Web, x => x.Url, x => x.ServerRelativeUrl, x => x.RootFolder.WelcomePage);
        await ctx.ExecuteQueryAsync().ConfigureAwait(false);
        var welcome = ctx.Web.RootFolder.WelcomePage?.Split('?', '#')[0].Trim('/');
        return string.IsNullOrEmpty(welcome) ? null : await LocateAsync(ctx, welcome).ConfigureAwait(false);
    }, ct);

    /// <summary>
    /// What's at an address in a site, from the site's own address
    /// ("Shared Documents/Policies", "SitePages/Home.aspx", "Lists/Contacts"):
    /// a file, a folder, or a list or library, with the row it's shown under.
    /// Null when there's nothing there (or nothing in a list).
    /// </summary>
    public Task<FoundNode?> GetAtPathAsync(SpNode web, string path, CancellationToken ct = default) => RetryAsync<FoundNode?>(async () =>
    {
        using var ctx = await ContextAsync(web.WebUrl, ct).ConfigureAwait(false);
        ctx.Load(ctx.Site, s => s.Url);
        ctx.Load(ctx.Web, x => x.Url, x => x.ServerRelativeUrl);
        await ctx.ExecuteQueryAsync().ConfigureAwait(false);
        return await LocateAsync(ctx, path).ConfigureAwait(false);
    }, ct);

    /// <summary><see cref="GetAtPathAsync"/> in a context whose site URL and web URLs are loaded.</summary>
    private static async Task<FoundNode?> LocateAsync(ClientContext ctx, string path)
    {
        var w = ctx.Web;
        var siteUrl = ctx.Site.Url.TrimEnd('/');
        var webUrl = w.Url.TrimEnd('/');
        var resource = TenantUrls.Resource(webUrl);
        var full = $"{w.ServerRelativeUrl.TrimEnd('/')}/{Uri.UnescapeDataString(path).Replace('\\', '/').Trim('/')}";

        // a file
        var file = w.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(full));
        var fileItem = file.ListItemAllFields;
        var fileList = fileItem.ParentList;
        ctx.Load(file, f => f.Name, f => f.Title, f => f.Length, f => f.TimeLastModified, f => f.ServerRelativeUrl);
        ctx.Load(fileItem, i => i.Id);
        LoadList(ctx, fileList);
        if (await TryAsync(ctx).ConfigureAwait(false) && fileItem.ServerObjectIsNull != true && fileList.ServerObjectIsNull != true)
        {
            var list = ListNode(fileList, siteUrl, webUrl, resource);
            return new FoundNode(new SpNode
            {
                Kind = NodeKind.File,
                Title = string.IsNullOrWhiteSpace(file.Title) ? file.Name : file.Title,
                Url = resource + file.ServerRelativeUrl,
                WebUrl = webUrl,
                SiteUrl = siteUrl,
                ListId = list.ListId,
                ItemId = fileItem.Id,
                ServerRelativeUrl = file.ServerRelativeUrl,
                InLibrary = list.InLibrary,
                Size = file.Length,
                Modified = Utc(file.TimeLastModified),
            }, list);
        }

        // a folder in a list
        var folder = w.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(full));
        var folderItem = folder.ListItemAllFields;
        ctx.Load(folder, f => f.Name, f => f.ServerRelativeUrl, f => f.ItemCount, f => f.TimeLastModified);
        ctx.Load(folderItem, i => i.Id);
        if (await TryAsync(ctx).ConfigureAwait(false) && folderItem.ServerObjectIsNull != true)
        {
            var folderList = folderItem.ParentList;
            LoadList(ctx, folderList);
            if (await TryAsync(ctx).ConfigureAwait(false))
            {
                var list = ListNode(folderList, siteUrl, webUrl, resource);
                return new FoundNode(new SpNode
                {
                    Kind = NodeKind.Folder,
                    Title = folder.Name,
                    Url = resource + folder.ServerRelativeUrl,
                    WebUrl = webUrl,
                    SiteUrl = siteUrl,
                    ListId = list.ListId,
                    ItemId = folderItem.Id,
                    ServerRelativeUrl = folder.ServerRelativeUrl,
                    InLibrary = list.InLibrary,
                    ItemCount = folder.ItemCount,
                    Modified = Utc(folder.TimeLastModified),
                }, list);
            }
        }

        // a list or library itself
        var whole = w.GetList(full);
        LoadList(ctx, whole);
        return await TryAsync(ctx).ConfigureAwait(false) && whole.ServerObjectIsNull != true
            ? new FoundNode(ListNode(whole, siteUrl, webUrl, resource), null)
            : null;
    }

    /// <summary>Runs what's loaded; false when SharePoint says it isn't there.</summary>
    private static async Task<bool> TryAsync(ClientContext ctx)
    {
        try
        {
            await ctx.ExecuteQueryAsync().ConfigureAwait(false);
            return true;
        }
        catch (ServerException)
        {
            return false;
        }
    }

    private static void LoadList(ClientContext ctx, List l) => ctx.Load(l,
        x => x.Id, x => x.Title, x => x.BaseType, x => x.BaseTemplate, x => x.ItemCount, x => x.LastItemModifiedDate, x => x.RootFolder.ServerRelativeUrl);

    /// <summary>A list or library's row, as <see cref="GetWebContentsAsync"/> makes them.</summary>
    private static SpNode ListNode(List l, string siteUrl, string webUrl, string resource) => new()
    {
        Kind = l.BaseType == BaseType.DocumentLibrary ? NodeKind.Library : NodeKind.List,
        Title = l.Title,
        Url = resource + l.RootFolder.ServerRelativeUrl,
        WebUrl = webUrl,
        SiteUrl = siteUrl,
        ListId = l.Id,
        ServerRelativeUrl = l.RootFolder.ServerRelativeUrl,
        InLibrary = l.BaseType == BaseType.DocumentLibrary,
        Template = l.BaseTemplate.ToString(CultureInfo.InvariantCulture),
        ItemCount = l.ItemCount,
        Modified = Utc(l.LastItemModifiedDate),
    };

    private static async Task<IReadOnlyList<Web>> LoadSubwebsAsync(ClientContext ctx, Web web)
    {
        try
        {
            var all = ctx.LoadQuery(web.Webs.Include(x => x.Title, x => x.Url, x => x.WebTemplate, x => x.Configuration, x => x.LastItemModifiedDate));
            await ctx.ExecuteQueryAsync().ConfigureAwait(false);
            return all.ToList();
        }
        catch (Exception e) when (IsAccessDenied(e) || e is ServerException)
        {
            try
            {
                var mine = ctx.LoadQuery(web.GetSubwebsForCurrentUser(new SubwebQuery())
                    .Include(x => x.Title, x => x.Url, x => x.WebTemplate, x => x.Configuration, x => x.LastItemModifiedDate));
                await ctx.ExecuteQueryAsync().ConfigureAwait(false);
                return mine.ToList();
            }
            catch (Exception e2) when (IsAccessDenied(e2) || e2 is ServerException)
            {
                return [];
            }
        }
    }

    /// <summary>
    /// What's in a list, library or folder, a page at a time (<paramref name="cursor"/>
    /// is the previous page's <see cref="NodePage.Next"/>): folders first, then
    /// files or items, by name. Pages go by ID, which lists over the 5,000-item
    /// threshold allow; each page is sorted on its own. <paramref name="recursive"/>
    /// takes in everything in its subfolders too.
    /// </summary>
    public Task<NodePage> GetFolderContentsAsync(SpNode parent, string? cursor = null, CancellationToken ct = default) =>
        GetFolderContentsAsync(parent, cursor, recursive: false, ct);

    public Task<NodePage> GetFolderContentsAsync(SpNode parent, string? cursor, bool recursive, CancellationToken ct = default) => RetryAsync(async () =>
    {
        if (parent.ListId is not { } listId)
        {
            throw new ArgumentException("Only lists, libraries and folders have contents.");
        }
        using var ctx = await ContextAsync(parent.WebUrl, ct).ConfigureAwait(false);
        var list = ctx.Web.Lists.GetById(listId);
        // a file's size and who has it checked out are library columns: asking a list for them fails
        string[] fields = parent.InLibrary
            ? ["FileLeafRef", "FileRef", "Title", "Modified", "ItemChildCount", "FolderChildCount", "File_x0020_Size", "CheckoutUser"]
            : ["FileLeafRef", "FileRef", "Title", "Modified", "ItemChildCount", "FolderChildCount"];
        var query = new CamlQuery
        {
            ViewXml =
                $"<View{(recursive ? " Scope='RecursiveAll'" : "")}><Query><OrderBy Override='TRUE'><FieldRef Name='ID'/></OrderBy></Query>" +
                $"<ViewFields>{string.Concat(fields.Select(f => $"<FieldRef Name='{f}'/>"))}</ViewFields>" +
                $"<RowLimit Paged='TRUE'>{FolderPageSize}</RowLimit></View>",
            FolderServerRelativeUrl = parent.Kind == NodeKind.Folder ? parent.ServerRelativeUrl : null,
            ListItemCollectionPosition = string.IsNullOrEmpty(cursor) ? null : new ListItemCollectionPosition { PagingInfo = cursor },
        };
        var items = list.GetItems(query);
        System.Linq.Expressions.Expression<Func<ListItem, object>>[] load =
        [
            x => x.Id, x => x.FileSystemObjectType,
            .. fields.Select(f => (System.Linq.Expressions.Expression<Func<ListItem, object>>)(x => x[f])),
        ];
        ctx.Load(items, i => i.ListItemCollectionPosition, i => i.Include(load));
        ctx.Load(list, l => l.RootFolder.ServerRelativeUrl);
        await ctx.ExecuteQueryAsync().ConfigureAwait(false);

        var resource = TenantUrls.Resource(parent.WebUrl);
        var root = list.RootFolder.ServerRelativeUrl;
        var nodes = new List<SpNode>();
        foreach (var item in items)
        {
            var fileRef = item["FileRef"] as string ?? "";
            var leaf = item["FileLeafRef"] as string ?? fileRef.Split('/').Last();
            var folder = item.FileSystemObjectType == FileSystemObjectType.Folder;
            var kind = folder ? NodeKind.Folder : parent.InLibrary ? NodeKind.File : NodeKind.Item;
            var title = kind == NodeKind.Item && item["Title"] is string t && !string.IsNullOrWhiteSpace(t) ? t : leaf;
            nodes.Add(new SpNode
            {
                Kind = kind,
                Title = title,
                Url = kind == NodeKind.Item ? $"{resource}{root}/DispForm.aspx?ID={item.Id}" : resource + fileRef,
                WebUrl = parent.WebUrl,
                SiteUrl = parent.SiteUrl,
                ListId = listId,
                ItemId = item.Id,
                ServerRelativeUrl = fileRef,
                InLibrary = parent.InLibrary,
                ItemCount = folder ? ChildCount(item["ItemChildCount"]) + ChildCount(item["FolderChildCount"]) : null,
                Size = kind == NodeKind.File && long.TryParse(item["File_x0020_Size"]?.ToString(), out var size) ? size : null,
                Modified = item["Modified"] is DateTime m ? Utc(m) : null,
                CheckedOutTo = parent.InLibrary ? (item["CheckoutUser"] as FieldUserValue)?.LookupValue : null,
            });
        }
        nodes.Sort((a, b) =>
            a.Kind == NodeKind.Folder != (b.Kind == NodeKind.Folder)
                ? (a.Kind == NodeKind.Folder ? -1 : 1)
                : StringComparer.CurrentCultureIgnoreCase.Compare(a.Title, b.Title));
        return new NodePage(nodes, items.ListItemCollectionPosition?.PagingInfo);
    }, ct);

    private static int ChildCount(object? value) => value switch
    {
        FieldLookupValue l => int.TryParse(l.LookupValue, out var n) ? n : 0,
        string s => int.TryParse(s, out var n) ? n : 0,
        int n => n,
        _ => 0,
    };

    private static DateTimeOffset? Utc(DateTime d) =>
        d == default || d.Year < 1990 ? null : new DateTimeOffset(DateTime.SpecifyKind(d, DateTimeKind.Utc));
}
