using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>
/// Made-up site collections, lists, folders and files for trying the table and
/// the pane without a tenant: SSP_DEMO=1. Commands still go to SharePoint, so
/// they fail here; everything up to Run can be tried.
/// </summary>
internal static class Demo
{
    public static bool On { get; } = Environment.GetEnvironmentVariable("SSP_DEMO") == "1";

    private const string Tenant = "https://contoso.sharepoint.com";

    public static readonly Identity Identity = new("Ada Lovelace", "ada@contoso.com", "i:0#.f|membership|ada@contoso.com", false, "Contoso");

    private static readonly (string Path, string Title, string Template)[] Sites =
    [
        ("", "Contoso", "SITEPAGEPUBLISHING#0"),
        ("/sites/hr", "Human Resources", "GROUP#0"),
        ("/sites/finance", "Finance", "GROUP#0"),
        ("/sites/intranet", "Intranet", "SITEPAGEPUBLISHING#0"),
        ("/sites/projects", "Projects", "STS#3"),
        ("/sites/legal", "Legal", "GROUP#0"),
        ("/sites/marketing", "Marketing", "SITEPAGEPUBLISHING#0"),
        ("/sites/appcatalog", "Apps", "APPCATALOG#0"),
        ("/teams/engineering", "Engineering", "GROUP#0"),
        ("/teams/engineering-platform", "Engineering - Platform", "TEAMCHANNEL#1"),
    ];

    public static Task<SitePage> SitesAsync(string? query)
    {
        var sites = Sites
            .Where(s => query is null || s.Title.Contains(query, StringComparison.OrdinalIgnoreCase) || s.Path.Contains(query, StringComparison.OrdinalIgnoreCase))
            .Select((s, i) => new SiteSummary(Tenant + s.Path, s.Title, s.Template, DateTimeOffset.Now.AddDays(-i * 3).AddHours(-i), Guid.NewGuid()))
            .ToList();
        return Delay(new SitePage(sites, null, sites.Count, false, "admin"));
    }

    public static Task<IReadOnlyList<SpNode>> WebAsync(SpNode web)
    {
        var lists = new List<SpNode>();
        if (web.Kind == NodeKind.SiteCollection && web.Url.EndsWith("/projects", StringComparison.Ordinal))
        {
            lists.Add(new SpNode { Kind = NodeKind.Subsite, Title = "Apollo", Url = web.Url + "/apollo", WebUrl = web.Url + "/apollo", SiteUrl = web.Url, Template = "STS#3", Modified = DateTimeOffset.Now.AddDays(-2) });
            lists.Add(new SpNode { Kind = NodeKind.Subsite, Title = "Gemini", Url = web.Url + "/gemini", WebUrl = web.Url + "/gemini", SiteUrl = web.Url, Template = "STS#3", Modified = DateTimeOffset.Now.AddDays(-40) });
        }
        var path = new Uri(web.Url).AbsolutePath.TrimEnd('/');
        SpNode List(NodeKind kind, string title, string leaf, string template, int count) => new()
        {
            Kind = kind,
            Title = title,
            Url = $"{Tenant}{path}/{leaf}",
            WebUrl = web.WebUrl,
            SiteUrl = web.SiteUrl,
            ListId = StableId($"{path}/{leaf}"),
            ServerRelativeUrl = $"{path}/{leaf}",
            InLibrary = kind == NodeKind.Library,
            Template = template,
            ItemCount = count,
            Modified = DateTimeOffset.Now.AddHours(-count),
        };
        lists.Add(List(NodeKind.Library, "Documents", "Shared Documents", "101", 1248));
        lists.Add(List(NodeKind.Library, "Site Pages", "SitePages", "119", 12));
        lists.Add(List(NodeKind.Library, "Site Assets", "SiteAssets", "101", 31));
        lists.Add(List(NodeKind.List, "Tasks", "Lists/Tasks", "171", 57));
        lists.Add(List(NodeKind.List, "Contacts", "Lists/Contacts", "105", 8));
        return Delay<IReadOnlyList<SpNode>>(lists);
    }

    public static Task<NodePage> FolderAsync(SpNode parent)
    {
        var nodes = new List<SpNode>();
        var resource = Tenant;
        SpNode Child(NodeKind kind, string name, int id, long? size = null) => new()
        {
            Kind = kind,
            Title = name,
            Url = kind == NodeKind.Item ? $"{parent.Url}/DispForm.aspx?ID={id}" : $"{resource}{parent.ServerRelativeUrl}/{name}",
            WebUrl = parent.WebUrl,
            SiteUrl = parent.SiteUrl,
            ListId = parent.ListId,
            ItemId = id,
            ServerRelativeUrl = $"{parent.ServerRelativeUrl}/{name}",
            InLibrary = parent.InLibrary,
            Size = size,
            ItemCount = kind == NodeKind.Folder ? id * 3 : null,
            Modified = DateTimeOffset.Now.AddDays(-id),
        };
        if (parent.InLibrary)
        {
            nodes.Add(Child(NodeKind.Folder, "2024", 1));
            nodes.Add(Child(NodeKind.Folder, "2025", 2));
            nodes.Add(Child(NodeKind.Folder, "Templates", 3));
            nodes.Add(Child(NodeKind.File, "Budget.xlsx", 4, 284_113));
            nodes.Add(Child(NodeKind.File, "Board minutes.docx", 5, 61_420));
            nodes.Add(Child(NodeKind.File, "Roadmap.pptx", 6, 9_814_022));
            nodes.Add(Child(NodeKind.File, "Logo.png", 7, 18_002));
        }
        else
        {
            for (var i = 1; i <= 6; i++)
            {
                nodes.Add(Child(NodeKind.Item, $"Item {i}", i));
            }
        }
        return Delay(new NodePage(nodes, null));
    }

    /// <summary>The same made-up list ID each time a site opens, so saved selections find their lists again.</summary>
    private static Guid StableId(string text) =>
        new(System.Security.Cryptography.MD5.HashData(System.Text.Encoding.UTF8.GetBytes(text.ToLowerInvariant())));

    private static async Task<T> Delay<T>(T value)
    {
        await Task.Delay(250);
        return value;
    }
}
