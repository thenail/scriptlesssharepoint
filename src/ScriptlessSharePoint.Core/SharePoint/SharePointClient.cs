using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;

using ScriptlessSharePoint.Auth;
using ScriptlessSharePoint.Connections;

namespace ScriptlessSharePoint.SharePoint;

/// <summary>
/// The SharePoint operations, through CSOM, for one connection. Each operation
/// makes its own ClientContext with a fresh token, so expiry and throttling
/// retries stay simple.
/// </summary>
public sealed partial class SharePointClient
{
    /// <summary>List templates of a site collection's "Apps for SharePoint" library (336) and the tenant's (330).</summary>
    private const int SiteAppCatalogTemplate = 336;
    private const int TenantAppCatalogTemplate = 330;

    private static readonly TimeSpan AdminListLifetime = TimeSpan.FromMinutes(5);

    private readonly ITokenProvider tokens;
    private (DateTimeOffset At, List<SiteSummary> Sites)? adminList;

    public SharePointClient(ConnectionProfile profile, ITokenProvider tokens, bool interactive)
    {
        Profile = profile;
        this.tokens = tokens;
        Interactive = interactive;
        TenantUrl = TenantUrls.NormalizeTenantUrl(profile.TenantUrl);
    }

    public ConnectionProfile Profile { get; }

    /// <summary>May sign-in show UI (the app) or not (the MCP server)?</summary>
    public bool Interactive { get; }

    public string TenantUrl { get; }

    public string AdminUrl => Profile.EffectiveAdminUrl;

    /// <summary>Who the connection acts as, and the tenant root's title: a connection test.</summary>
    public Task<Identity> WhoAmIAsync(CancellationToken ct = default) => RetryAsync(async () =>
    {
        using var ctx = await ContextAsync(TenantUrl, ct).ConfigureAwait(false);
        var user = ctx.Web.CurrentUser;
        ctx.Load(ctx.Web, w => w.Title);
        ctx.Load(user, u => u.Title, u => u.Email, u => u.LoginName);
        await ctx.ExecuteQueryAsync().ConfigureAwait(false);
        var isApp = Profile.Kind == ConnectionKind.Certificate ||
                    user.LoginName.Contains("app@sharepoint", StringComparison.OrdinalIgnoreCase);
        return new Identity(user.Title, NullIfEmpty(user.Email), user.LoginName, isApp, ctx.Web.Title);
    }, ct);

    /// <summary>
    /// Site collections the connection can see, through SharePoint search (any
    /// user; results follow the search index, which can lag a few minutes), a page
    /// at a time: <paramref name="cursor"/> is the previous page's <see cref="SitePage.Next"/>.
    /// </summary>
    public Task<SitePage> SearchSitesAsync(string? text, int limit = 50, string? cursor = null, CancellationToken ct = default) => RetryAsync(async () =>
    {
        var start = int.TryParse(cursor, out var s) && s > 0 ? s : 0;
        using var ctx = await ContextAsync(TenantUrl, ct).ConfigureAwait(false);
        var query = new KeywordQuery(ctx)
        {
            QueryText = TenantUrls.SiteSearchQuery(text),
            RowLimit = Math.Clamp(limit, 1, 500),
            StartRow = start,
            TrimDuplicates = false,
            EnableQueryRules = false,
            ClientType = "ScriptlessSharePoint",
        };
        foreach (var p in new[] { "Title", "Path", "SiteId", "WebTemplate", "LastModifiedTime" })
        {
            query.SelectProperties.Add(p);
        }
        var results = new SearchExecutor(ctx).ExecuteQuery(query);
        await ctx.ExecuteQueryAsync().ConfigureAwait(false);

        var table = results.Value.FirstOrDefault(t => t.TableType == KnownTableTypes.RelevantResults);
        if (table is null)
        {
            return new SitePage([], null, 0, false, "search");
        }
        var sites = table.ResultRows.Select(row =>
        {
            var url = Get(row, "Path") ?? "";
            return new SiteSummary(
                url.TrimEnd('/'),
                Get(row, "Title") ?? url,
                Get(row, "WebTemplate"),
                DateTimeOffset.TryParse(Get(row, "LastModifiedTime"), out var m) ? m : null,
                Guid.TryParse(Get(row, "SiteId"), out var id) ? id : null);
        }).Where(x => x.Url.Length > 0).ToList();
        var end = start + table.RowCount;
        return new SitePage(sites, table.RowCount > 0 && end < table.TotalRows ? end.ToString() : null, table.TotalRows, false, "search");
    }, ct);

    /// <summary>
    /// Files, pages, items or lists SharePoint search finds for a KQL query, up
    /// to <paramref name="limit"/>, a page of 500 at a time. Search trims to what
    /// the connection can open, and runs a few minutes behind.
    /// </summary>
    public async Task<IReadOnlyList<SearchHit>> SearchContentAsync(string kql, string? sortBy = null, bool descending = false, int limit = ContentPresets.Limit,
        IProgress<int>? progress = null, CancellationToken ct = default)
    {
        var hits = new List<SearchHit>();
        using var ctx = await ContextAsync(TenantUrl, ct).ConfigureAwait(false);
        while (hits.Count < limit)
        {
            ct.ThrowIfCancellationRequested();
            var query = new KeywordQuery(ctx)
            {
                QueryText = kql,
                RowLimit = Math.Min(500, limit - hits.Count),
                StartRow = hits.Count,
                TrimDuplicates = false,
                EnableQueryRules = false,
                ClientType = "ScriptlessSharePoint",
            };
            foreach (var p in new[] { "Title", "Path", "Filename", "Size", "LastModifiedTime", "SPSiteURL", "SPWebUrl", "SiteTitle", "ListID", "ListItemID", "IsDocument", "contentclass" })
            {
                query.SelectProperties.Add(p);
            }
            if (sortBy is not null)
            {
                query.SortList.Add(sortBy, descending ? SortDirection.Descending : SortDirection.Ascending);
            }
            var results = new SearchExecutor(ctx).ExecuteQuery(query);
            await RetryAsync(async () =>
            {
                await ctx.ExecuteQueryAsync().ConfigureAwait(false);
                return true;
            }, ct).ConfigureAwait(false);
            var table = results.Value.FirstOrDefault(t => t.TableType == KnownTableTypes.RelevantResults);
            if (table is null || table.RowCount == 0)
            {
                break;
            }
            hits.AddRange(table.ResultRows.Select(Hit).OfType<SearchHit>());
            progress?.Report(hits.Count);
            if (hits.Count >= table.TotalRows || table.RowCount < query.RowLimit)
            {
                break;
            }
        }
        return hits;

        static SearchHit? Hit(IDictionary<string, object> row)
        {
            var r = new Dictionary<string, object>(row, StringComparer.OrdinalIgnoreCase);
            string? Text(string key) => r.TryGetValue(key, out var v) && v is not null ? NullIfEmpty(v.ToString()) : null;
            var path = Text("Path");
            var site = Text("SPSiteURL");
            var web = Text("SPWebUrl") ?? site;
            if (path is null || site is null || web is null)
            {
                return null;
            }
            return new SearchHit(path, Text("Title") ?? Text("Filename") ?? path, site, web)
            {
                SiteTitle = Text("SiteTitle"),
                FileName = Text("Filename"),
                Size = long.TryParse(Text("Size"), out var size) ? size : null,
                Modified = DateTimeOffset.TryParse(Text("LastModifiedTime"), System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AssumeUniversal, out var m) ? m : null,
                ListId = Guid.TryParse(Text("ListID"), out var list) ? list : null,
                ItemId = int.TryParse(Text("ListItemID"), out var item) ? item : null,
                IsDocument = string.Equals(Text("IsDocument"), "true", StringComparison.OrdinalIgnoreCase) || Text("IsDocument") == "1",
                ContentClass = Text("contentclass"),
            };
        }
    }

    /// <summary>The admin centre's own list of every site (its Active sites page reads it).</summary>
    private const string AdminSitesList = "DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS";

    /// <summary>Set when the tenant has no such list: the admin listing uses the tenant API instead.</summary>
    private bool adminSitesListMissing;

    /// <summary>
    /// Every site collection in the tenant, from the admin centre (needs a SharePoint
    /// administrator), a page at a time and searched in SharePoint: without
    /// <paramref name="text"/> all sites by title; with it, the sites whose title
    /// has all its words, then those whose URL has them. OneDrive, redirect and
    /// deleted sites are left out.
    /// </summary>
    /// <remarks>
    /// The admin centre's site list is indexed but over the list view threshold
    /// (5,000), so each query uses one index-friendly condition: a title or URL
    /// search runs as two queries (SharePoint refuses "title or URL" in one), one
    /// after the other in the cursor. A word matching thousands of sites is
    /// refused too: <see cref="SharePointOperationException"/> asks for more of it.
    /// Tenants without the list fall back to the tenant API (all sites at once).
    /// </remarks>
    public async Task<SitePage> ListSitesAsAdminAsync(string? text, int limit = 50, string? cursor = null, CancellationToken ct = default)
    {
        var words = TenantUrls.SearchWords(text);
        if (!adminSitesListMissing && cursor?.StartsWith("o:", StringComparison.Ordinal) != true)
        {
            try
            {
                return await RetryAsync(() => QueryAdminSitesListAsync(words, Math.Clamp(limit, 1, 500), cursor, ct), ct).ConfigureAwait(false);
            }
            catch (ServerException e) when (e.Message.Contains("does not exist", StringComparison.OrdinalIgnoreCase) && e.Message.Contains("list", StringComparison.OrdinalIgnoreCase))
            {
                adminSitesListMissing = true;
            }
            catch (ServerException e) when (IsThrottled(e))
            {
                throw new SharePointOperationException(
                    $"Too many sites match \"{string.Join(' ', words)}\" for SharePoint to search them. Type more of the site's name or URL.", e);
            }
        }
        return await ListSitesFromTenantApiAsync(words, limit, cursor, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// One page from the admin centre's site list. The cursor is "a:" (all sites),
    /// "t:" (title matches) or "u:" (URL matches) and SharePoint's paging info.
    /// A page is topped up from the next query when one runs out.
    /// </summary>
    private async Task<SitePage> QueryAdminSitesListAsync(IReadOnlyList<string> words, int limit, string? cursor, CancellationToken ct)
    {
        using var ctx = await ContextAsync(AdminUrl, ct).ConfigureAwait(false);
        var list = ctx.Web.Lists.GetByTitle(AdminSitesList);
        var sites = new List<SiteSummary>();
        var state = string.IsNullOrEmpty(cursor) ? (words.Count == 0 ? "a:" : "t:") : cursor;
        int? total = null;
        // a few rounds at most: skipped rows (deleted, OneDrive, already shown) can leave a page short
        for (var round = 0; round < 6 && state is not null && sites.Count < limit; round++)
        {
            var kind = state[0];
            var paging = state.Length > 2 ? state[2..] : null;
            var where = kind switch
            {
                't' => AllContain("Title", words),
                'u' => AllContain("SiteUrl", words),
                _ => null,
            };
            // all sites: by the title index; searches: a Where clause, sorted by title
            var orderBy = where is null
                ? "<OrderBy Override='TRUE' UseIndexForOrderBy='TRUE'><FieldRef Name='Title'/></OrderBy>"
                : "<OrderBy><FieldRef Name='Title'/></OrderBy>";
            var view =
                $"<View><Query>{(where is null ? "" : $"<Where>{where}</Where>")}{orderBy}</Query>" +
                "<ViewFields><FieldRef Name='Title'/><FieldRef Name='SiteUrl'/><FieldRef Name='TemplateName'/>" +
                "<FieldRef Name='LastActivityOn'/><FieldRef Name='TimeCreated'/><FieldRef Name='SiteId'/><FieldRef Name='TimeDeleted'/></ViewFields>" +
                $"<RowLimit Paged='TRUE'>{limit - sites.Count}</RowLimit></View>";
            var items = list.GetItems(new CamlQuery
            {
                ViewXml = view,
                ListItemCollectionPosition = paging is null ? null : new ListItemCollectionPosition { PagingInfo = paging },
            });
            ctx.Load(items, i => i.ListItemCollectionPosition, i => i.Include(
                x => x["Title"], x => x["SiteUrl"], x => x["TemplateName"], x => x["LastActivityOn"],
                x => x["TimeCreated"], x => x["SiteId"], x => x["TimeDeleted"]));
            if (kind == 'a' && total is null)
            {
                ctx.Load(list, l => l.ItemCount);
            }
            await ctx.ExecuteQueryAsync().ConfigureAwait(false);
            if (kind == 'a')
            {
                total ??= list.ItemCount;
            }

            foreach (var item in items)
            {
                var url = (item["SiteUrl"] as string)?.TrimEnd('/');
                var title = item["Title"] as string;
                var template = item["TemplateName"] as string;
                if (string.IsNullOrEmpty(url) || item["TimeDeleted"] is DateTime ||
                    template?.StartsWith("SPSPERS", StringComparison.OrdinalIgnoreCase) == true ||
                    template?.StartsWith("RedirectSite", StringComparison.OrdinalIgnoreCase) == true)
                {
                    continue;
                }
                // URL matches whose title matches too were in the title query already
                if (kind == 'u' && words.All(w => title?.Contains(w, StringComparison.OrdinalIgnoreCase) == true))
                {
                    continue;
                }
                var activity = item["LastActivityOn"] as DateTime? ?? item["TimeCreated"] as DateTime?;
                sites.Add(new SiteSummary(
                    url,
                    string.IsNullOrWhiteSpace(title) ? url : title,
                    template,
                    activity is { } a ? new DateTimeOffset(DateTime.SpecifyKind(a, DateTimeKind.Utc)) : null,
                    item["SiteId"] is Guid id ? id : null));
            }

            var next = items.ListItemCollectionPosition?.PagingInfo;
            state = next is not null ? $"{kind}:{next}" : kind == 't' ? "u:" : null;
        }
        return new SitePage(sites, state, total, total is not null, "admin");
    }

    private static string AllContain(string field, IReadOnlyList<string> words)
    {
        static string Contains(string f, string w) =>
            $"<Contains><FieldRef Name='{f}'/><Value Type='Text'>{System.Security.SecurityElement.Escape(w)}</Value></Contains>";
        var clause = Contains(field, words[0]);
        foreach (var w in words.Skip(1))
        {
            clause = $"<And>{clause}{Contains(field, w)}</And>";
        }
        return clause;
    }

    private static bool IsThrottled(ServerException e) =>
        e.ServerErrorTypeName?.Contains("Throttled", StringComparison.OrdinalIgnoreCase) == true ||
        e.Message.Contains("list view threshold", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// The fallback: every site from the tenant API (it can't search titles), kept
    /// for a few minutes and paged here; the cursor is "o:" and an offset.
    /// </summary>
    private async Task<SitePage> ListSitesFromTenantApiAsync(IReadOnlyList<string> words, int limit, string? cursor, CancellationToken ct)
    {
        var start = cursor?.StartsWith("o:", StringComparison.Ordinal) == true && int.TryParse(cursor[2..], out var s) ? s : 0;
        if (start == 0 || adminList is null || DateTimeOffset.UtcNow - adminList.Value.At > AdminListLifetime)
        {
            adminList = (DateTimeOffset.UtcNow, await LoadAllSitesAsync(ct).ConfigureAwait(false));
        }
        var matches = adminList.Value.Sites
            .Where(x => words.All(w =>
                x.Title.Contains(w, StringComparison.OrdinalIgnoreCase) ||
                x.Url.Contains(w, StringComparison.OrdinalIgnoreCase)))
            .ToList();
        var page = matches.Skip(start).Take(Math.Clamp(limit, 1, 500)).ToList();
        var end = start + page.Count;
        return new SitePage(page, end < matches.Count ? $"o:{end}" : null, matches.Count, false, "admin");
    }

    private const int MaxAdminSites = 20_000;

    private Task<List<SiteSummary>> LoadAllSitesAsync(CancellationToken ct) => RetryAsync(async () =>
    {
        using var ctx = await ContextAsync(AdminUrl, ct).ConfigureAwait(false);
        var tenant = new Tenant(ctx);
        var all = new List<SiteSummary>();
        string? next = null;
        do
        {
            ct.ThrowIfCancellationRequested();
            var page = tenant.GetSitePropertiesFromSharePointByFilters(new SPOSitePropertiesEnumerableFilter
            {
                IncludePersonalSite = PersonalSiteFilter.Exclude,
                IncludeDetail = false,
                StartIndex = next,
            });
            ctx.Load(page);
            await ctx.ExecuteQueryAsync().ConfigureAwait(false);
            foreach (var p in page)
            {
                if (p.Template?.StartsWith("RedirectSite", StringComparison.OrdinalIgnoreCase) == true)
                {
                    continue;
                }
                all.Add(new SiteSummary(
                    p.Url.TrimEnd('/'),
                    string.IsNullOrWhiteSpace(p.Title) ? p.Url : p.Title,
                    p.Template,
                    p.LastContentModifiedDate == default ? null : new DateTimeOffset(DateTime.SpecifyKind(p.LastContentModifiedDate, DateTimeKind.Utc)),
                    p.SiteId == Guid.Empty ? null : p.SiteId));
            }
            next = page.NextStartIndexFromSharePoint;
        }
        while (!string.IsNullOrEmpty(next) && all.Count < MaxAdminSites);
        all.Sort((a, b) => string.Compare(a.Title, b.Title, StringComparison.CurrentCultureIgnoreCase));
        return all;
    }, ct);

    private (DateTimeOffset At, IReadOnlyList<SiteFacts> Sites)? siteFacts;
    private (DateTimeOffset At, IReadOnlyList<SiteFacts> Sites)? oneDriveFacts;

    /// <summary>
    /// What the admin centre knows about every site collection (OneDrives and
    /// redirects left out), for the predefined selections: read a few hundred
    /// at a time, and kept for a few minutes unless <paramref name="refresh"/>.
    /// <paramref name="progress"/> hears how many have come. With
    /// <paramref name="oneDrives"/>, the OneDrives instead. Needs a SharePoint administrator.
    /// </summary>
    public async Task<IReadOnlyList<SiteFacts>> SiteFactsAsync(bool refresh = false, IProgress<int>? progress = null, CancellationToken ct = default, bool oneDrives = false)
    {
        if (!refresh && (oneDrives ? oneDriveFacts : siteFacts) is { } cached && DateTimeOffset.UtcNow - cached.At < AdminListLifetime)
        {
            return cached.Sites;
        }
        var sites = await RetryAsync<IReadOnlyList<SiteFacts>>(async () =>
        {
            using var ctx = await ContextAsync(AdminUrl, ct).ConfigureAwait(false);
            var tenant = new Tenant(ctx);
            var all = new List<SiteFacts>();
            string? next = null;
            do
            {
                ct.ThrowIfCancellationRequested();
                var page = tenant.GetSitePropertiesFromSharePointByFilters(new SPOSitePropertiesEnumerableFilter
                {
                    IncludePersonalSite = oneDrives ? PersonalSiteFilter.Include : PersonalSiteFilter.Exclude,
                    Template = oneDrives ? "SPSPERS" : null,
                    IncludeDetail = false,
                    StartIndex = next,
                });
                ctx.Load(page);
                await ctx.ExecuteQueryAsync().ConfigureAwait(false);
                all.AddRange(page.ToList()
                    .Where(p => oneDrives
                        ? p.Template?.StartsWith("SPSPERS", StringComparison.OrdinalIgnoreCase) == true
                        : p.Template?.StartsWith("RedirectSite", StringComparison.OrdinalIgnoreCase) != true)
                    .Select(FactsOf));
                progress?.Report(all.Count);
                next = page.NextStartIndexFromSharePoint;
            }
            while (!string.IsNullOrEmpty(next) && all.Count < MaxAdminSites);
            all.Sort((a, b) => string.Compare(a.Title, b.Title, StringComparison.CurrentCultureIgnoreCase));
            return all;
        }, ct).ConfigureAwait(false);
        if (oneDrives)
        {
            oneDriveFacts = (DateTimeOffset.UtcNow, sites);
        }
        else
        {
            siteFacts = (DateTimeOffset.UtcNow, sites);
        }
        return sites;
    }

    /// <summary>A site's facts; what SharePoint didn't send is left at its default.</summary>
    private static SiteFacts FactsOf(SiteProperties p)
    {
        static T Get<T>(Func<T> read, T fallback)
        {
            try
            {
                return read();
            }
            catch (PropertyOrFieldNotInitializedException)
            {
                return fallback;
            }
        }
        static DateTimeOffset? When(DateTime d) => d == default || d.Year < 1990 ? null : new DateTimeOffset(DateTime.SpecifyKind(d, DateTimeKind.Utc));

        var url = p.Url.TrimEnd('/');
        return new SiteFacts(
            url,
            Get(() => p.Title, null) is { Length: > 0 } title ? title : url,
            Get(() => p.Template, null),
            When(Get(() => p.LastContentModifiedDate, default(DateTime))),
            When(Get(() => p.CreatedTime, default(DateTime))),
            Get(() => p.SiteId, Guid.Empty) is var id && id != Guid.Empty ? id : null)
        {
            HubSiteId = Get(() => p.HubSiteId, Guid.Empty),
            IsHubSite = Get(() => p.IsHubSite, false),
            GroupId = Get(() => p.GroupId, Guid.Empty),
            IsTeamsConnected = Get(() => p.IsTeamsConnected, false),
            Sharing = Get(() => p.SharingCapability.ToString(), null),
            LockState = Get(() => p.LockState, null),
            StorageUsedMb = Get(() => p.StorageUsage, 0L),
            StorageLimitMb = Get(() => p.StorageMaximumLevel, 0L),
            SensitivityLabel = Get(() => p.SensitivityLabel, Guid.Empty),
            ArchiveStatus = Get(() => p.ArchiveStatus, null),
        };
    }

    /// <summary>
    /// A site collection (for any URL in it) and whether it has its own app catalog:
    /// from the tenant's list of site collection app catalogs when the connection
    /// may read it, otherwise from the site's Apps for SharePoint library.
    /// </summary>
    public Task<SiteDetails> GetSiteAsync(string url, CancellationToken ct = default) => RetryAsync(async () =>
    {
        var requested = TenantUrls.NormalizeSiteUrl(url, TenantUrl);
        using var ctx = await ContextAsync(requested, ct).ConfigureAwait(false);
        var site = ctx.Site;
        var root = site.RootWeb;
        ctx.Load(site, s => s.Id, s => s.Url);
        ctx.Load(root, w => w.Title, w => w.WebTemplate, w => w.Configuration);
        var lists = ctx.LoadQuery(root.Lists.Include(
            l => l.Title,
            l => l.BaseTemplate,
            l => l.ItemCount,
            l => l.RootFolder.Name,
            l => l.RootFolder.ServerRelativeUrl));
        await ctx.ExecuteQueryAsync().ConfigureAwait(false);

        var siteUrl = site.Url.TrimEnd('/');
        var library = lists.FirstOrDefault(l => l.BaseTemplate == SiteAppCatalogTemplate)
            ?? lists.FirstOrDefault(l =>
                l.BaseTemplate != TenantAppCatalogTemplate &&
                string.Equals(l.RootFolder.Name, "AppCatalog", StringComparison.OrdinalIgnoreCase));
        var libraryUrl = library is null ? null : $"{TenantUrls.Resource(siteUrl)}{library.RootFolder.ServerRelativeUrl}";

        var state = library is null ? AppCatalogState.NotEnabled : AppCatalogState.Enabled;
        var source = "site";
        try
        {
            var registered = await ListSiteAppCatalogsAsync(ct).ConfigureAwait(false);
            state = registered.Any(c => c.SiteId == site.Id || string.Equals(c.Url.TrimEnd('/'), siteUrl, StringComparison.OrdinalIgnoreCase))
                ? AppCatalogState.Enabled
                : AppCatalogState.NotEnabled;
            source = "tenant";
        }
        catch (Exception e) when (IsAccessDenied(e))
        {
            // not an admin: the library is what there is to go on
        }

        return new SiteDetails(
            siteUrl,
            root.Title,
            $"{root.WebTemplate}#{root.Configuration}",
            site.Id,
            state,
            source,
            state == AppCatalogState.Enabled ? libraryUrl : null,
            state == AppCatalogState.Enabled ? library?.ItemCount : null,
            string.Equals(requested, siteUrl, StringComparison.OrdinalIgnoreCase) ? null : requested);
    }, ct);

    /// <summary>The tenant's site collection app catalogs (needs a SharePoint administrator).</summary>
    public Task<IReadOnlyList<SiteAppCatalog>> ListSiteAppCatalogsAsync(CancellationToken ct = default) => RetryAsync<IReadOnlyList<SiteAppCatalog>>(async () =>
    {
        using var ctx = await ContextAsync(AdminUrl, ct).ConfigureAwait(false);
        var catalogs = ctx.Web.TenantAppCatalog.SiteCollectionAppCatalogsSites;
        ctx.Load(catalogs, c => c.Include(i => i.AbsoluteUrl, i => i.SiteID, i => i.ErrorMessage));
        await ctx.ExecuteQueryAsync().ConfigureAwait(false);
        return catalogs
            .Select(c => new SiteAppCatalog(c.AbsoluteUrl.TrimEnd('/'), c.SiteID, NullIfEmpty(c.ErrorMessage)))
            .OrderBy(c => c.Url, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }, ct);

    /// <summary>
    /// Gives a site collection its own app catalog (an Apps for SharePoint library),
    /// as Add-SPOSiteCollectionAppCatalog does. Needs a SharePoint administrator.
    /// </summary>
    public async Task<AppCatalogChange> EnableSiteAppCatalogAsync(string url, CancellationToken ct = default)
    {
        var site = await GetSiteAsync(url, ct).ConfigureAwait(false);
        if (site.AppCatalog == AppCatalogState.Enabled && site.AppCatalogSource == "tenant")
        {
            return new AppCatalogChange(site.Url, false, $"{site.Url} already has a site collection app catalog.");
        }
        return await RetryAsync(async () =>
        {
            using var ctx = await ContextAsync(AdminUrl, ct).ConfigureAwait(false);
            var item = ctx.Web.TenantAppCatalog.SiteCollectionAppCatalogsSites.Add(site.Url);
            ctx.Load(item, i => i.ErrorMessage);
            try
            {
                await ctx.ExecuteQueryAsync().ConfigureAwait(false);
            }
            catch (ServerException e) when (e.Message.Contains("app catalog", StringComparison.OrdinalIgnoreCase) &&
                                            !e.Message.Contains("site collection", StringComparison.OrdinalIgnoreCase))
            {
                throw new SharePointOperationException(
                    $"SharePoint: {e.Message} The tenant needs an app catalog first: in the SharePoint admin centre, open More features › Apps.", e);
            }
            if (!string.IsNullOrEmpty(item.ErrorMessage))
            {
                throw new SharePointOperationException($"SharePoint couldn't add the app catalog: {item.ErrorMessage}");
            }
            return new AppCatalogChange(
                site.Url,
                true,
                $"Enabled the site collection app catalog on {site.Url}. SharePoint adds its Apps for SharePoint library within a few minutes.");
        }, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Stops a site collection's app catalog being one (Remove-SPOSiteCollectionAppCatalog).
    /// The library and its files stay in the site. Needs a SharePoint administrator.
    /// </summary>
    public async Task<AppCatalogChange> DisableSiteAppCatalogAsync(string url, CancellationToken ct = default)
    {
        var site = await GetSiteAsync(url, ct).ConfigureAwait(false);
        if (site.AppCatalog == AppCatalogState.NotEnabled && site.AppCatalogSource == "tenant")
        {
            return new AppCatalogChange(site.Url, false, $"{site.Url} doesn't have a site collection app catalog.");
        }
        return await RetryAsync(async () =>
        {
            using var ctx = await ContextAsync(AdminUrl, ct).ConfigureAwait(false);
            ctx.Web.TenantAppCatalog.SiteCollectionAppCatalogsSites.Remove(site.Url);
            await ctx.ExecuteQueryAsync().ConfigureAwait(false);
            return new AppCatalogChange(
                site.Url,
                true,
                $"Removed the site collection app catalog from {site.Url}. Its Apps for SharePoint library and files stay in the site.");
        }, ct).ConfigureAwait(false);
    }

    /// <summary>A ClientContext for a URL, with a bearer token for its host.</summary>
    internal async Task<ClientContext> ContextAsync(string url, CancellationToken ct)
    {
        var token = await tokens.GetTokenAsync(TenantUrls.Resource(url), Interactive, ct).ConfigureAwait(false);
        var ctx = new ClientContext(url) { ClientTag = $"ScriptlessSharePoint/{AppInfo.Version}" };
        ctx.ExecutingWebRequest += (_, e) =>
        {
            e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
            // decorated traffic, as Microsoft's throttling guidance asks
            e.WebRequestExecutor.WebRequest.UserAgent = $"NONISV|ScriptlessSharePoint|ScriptlessSharePoint/{AppInfo.Version}";
        };
        return ctx;
    }

    /// <summary>The longest a request keeps trying while SharePoint throttles it.</summary>
    public static readonly TimeSpan Patience = TimeSpan.FromMinutes(10);

    /// <summary>The tenant's manners: SharePoint's pauses, and how many run at once (shared by every job and client on it).</summary>
    public Throttle Throttle => Throttle.For(TenantUrl);

    /// <summary>
    /// Runs an operation, waiting out SharePoint's pause first. When it's
    /// throttled (429) or SharePoint is briefly down (503), everything on this
    /// connection pauses for as long as SharePoint asked, then it tries again,
    /// for up to <see cref="Patience"/> in all.
    /// </summary>
    internal async Task<T> RetryAsync<T>(Func<Task<T>> operation, CancellationToken ct)
    {
        var started = DateTimeOffset.UtcNow;
        for (var attempt = 0; ; attempt++)
        {
            await Throttle.WhilePausedAsync(ct).ConfigureAwait(false);
            try
            {
                var result = await operation().ConfigureAwait(false);
                Throttle.Succeeded();
                return result;
            }
            catch (Exception e) when (SharePointErrors.RetryAfter(e, attempt) is { } wait && DateTimeOffset.UtcNow - started + wait < Patience)
            {
                Throttle.Throttled(wait);
                Diagnostics.Log.Write("throttle", $"{TenantUrl}: told to wait {wait.TotalSeconds:0} s (attempt {attempt + 1}); {Throttle.Limit} at once from now");
            }
        }
    }

    public static bool IsAccessDenied(Exception e) =>
        e is ServerUnauthorizedAccessException ||
        (e is System.Net.WebException { Response: System.Net.HttpWebResponse r } && (int)r.StatusCode is 401 or 403) ||
        (e is HttpRequestException { StatusCode: System.Net.HttpStatusCode.Unauthorized or System.Net.HttpStatusCode.Forbidden });

    private static string? Get(IDictionary<string, object> row, string key) =>
        row.TryGetValue(key, out var v) && v is not null ? NullIfEmpty(v.ToString()) : null;

    private static string? NullIfEmpty(string? s) => string.IsNullOrWhiteSpace(s) ? null : s;
}
