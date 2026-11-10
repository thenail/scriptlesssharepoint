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

        var visible = lists.Where(l => !l.Hidden).Select(l => new SpNode
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
        });
        nodes.AddRange(visible
            .OrderBy(n => n.Kind == NodeKind.Library ? 0 : 1)
            .ThenBy(n => n.Title, StringComparer.CurrentCultureIgnoreCase));
        return nodes;
    }, ct);

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
    /// threshold allow; each page is sorted on its own.
    /// </summary>
    public Task<NodePage> GetFolderContentsAsync(SpNode parent, string? cursor = null, CancellationToken ct = default) => RetryAsync(async () =>
    {
        if (parent.ListId is not { } listId)
        {
            throw new ArgumentException("Only lists, libraries and folders have contents.");
        }
        using var ctx = await ContextAsync(parent.WebUrl, ct).ConfigureAwait(false);
        var list = ctx.Web.Lists.GetById(listId);
        var query = new CamlQuery
        {
            ViewXml =
                "<View><Query><OrderBy Override='TRUE'><FieldRef Name='ID'/></OrderBy></Query>" +
                "<ViewFields><FieldRef Name='FileLeafRef'/><FieldRef Name='FileRef'/><FieldRef Name='Title'/><FieldRef Name='Modified'/>" +
                "<FieldRef Name='File_x0020_Size'/><FieldRef Name='ItemChildCount'/><FieldRef Name='FolderChildCount'/><FieldRef Name='CheckoutUser'/></ViewFields>" +
                $"<RowLimit Paged='TRUE'>{FolderPageSize}</RowLimit></View>",
            FolderServerRelativeUrl = parent.Kind == NodeKind.Folder ? parent.ServerRelativeUrl : null,
            ListItemCollectionPosition = string.IsNullOrEmpty(cursor) ? null : new ListItemCollectionPosition { PagingInfo = cursor },
        };
        var items = list.GetItems(query);
        ctx.Load(items, i => i.ListItemCollectionPosition, i => i.Include(
            x => x.Id, x => x.FileSystemObjectType, x => x["FileLeafRef"], x => x["FileRef"], x => x["Title"], x => x["Modified"],
            x => x["File_x0020_Size"], x => x["ItemChildCount"], x => x["FolderChildCount"], x => x["CheckoutUser"]));
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
                CheckedOutTo = (item["CheckoutUser"] as FieldUserValue)?.LookupValue,
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
