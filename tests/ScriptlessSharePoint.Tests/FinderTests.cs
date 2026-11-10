using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class FinderTests
{
    private const string Web = "https://contoso.sharepoint.com/sites/hr";

    private static SpNode List(string path, string title, string template, NodeKind kind = NodeKind.Library) => new()
    {
        Kind = kind,
        Title = title,
        Url = "https://contoso.sharepoint.com" + path,
        WebUrl = Web,
        SiteUrl = Web,
        ListId = Guid.NewGuid(),
        ServerRelativeUrl = path,
        Template = template,
        InLibrary = kind == NodeKind.Library,
    };

    // a Danish site: the titles are translated, the templates and addresses aren't
    private static readonly SpNode SitePages = List("/sites/hr/SitePages", "Webstedssider", "119");
    private static readonly SpNode Documents = List("/sites/hr/Shared Documents", "Dokumenter", "101");
    private static readonly SpNode Assets = List("/sites/hr/SiteAssets", "Webstedsaktiver", "101");
    private static readonly SpNode Contracts = List("/sites/hr/Contracts", "Kontrakter", "101");
    private static readonly SpNode Contacts = List("/sites/hr/Lists/Contacts", "Kontakter", "105", NodeKind.List);
    private static readonly SpNode[] All = [SitePages, Documents, Assets, Contracts, Contacts];

    private static SpNode[] Find(string which, string? name = null) =>
        All.Where(l => CommandCatalog.ListMatches(l, which, name)).ToArray();

    [Fact]
    public void Finds_lists_by_template_or_address_never_by_title()
    {
        Assert.Equal([SitePages], Find("sitepages"));
        Assert.Equal([Documents], Find("documents"));
        Assert.Equal([Assets], Find("siteassets"));
        Assert.Equal([SitePages, Documents, Assets, Contracts], Find("libraries"));
        Assert.Equal([Contacts], Find("lists"));
        Assert.Empty(Find("pages"));
    }

    [Fact]
    public void Finds_by_the_name_in_the_address()
    {
        Assert.Equal([Contracts], Find("name", "contracts"));
        Assert.Equal([Contacts], Find("name", "/Contacts/"));
        Assert.Equal([Contacts], Find("name", "Lists/Contacts"));
        Assert.Empty(Find("name", "Kontrakter"));
        Assert.Empty(Find("name", " "));
        Assert.Empty(Find("name"));
    }

    [Fact]
    public void Finds_lists_by_kind()
    {
        Assert.Equal([Contacts], Find("t:105"));
        Assert.Equal([SitePages, Documents, Assets, Contracts, Contacts], Find("all"));
        var tasks = List("/sites/hr/Lists/Tasks", "Opgaver", "171", NodeKind.List);
        Assert.True(CommandCatalog.ListMatches(tasks, "t:107,171", null));
        Assert.False(CommandCatalog.ListMatches(tasks, "t:100", null));
    }

    private static readonly SpNode Docx = Documents with
    {
        Kind = NodeKind.File, Title = "Budget 2025.docx", Url = Documents.Url + "/Budget 2025.docx",
        ServerRelativeUrl = Documents.ServerRelativeUrl + "/Budget 2025.docx", ItemId = 4,
    };

    [Theory]
    [InlineData(null, true)]
    [InlineData(" ", true)]
    [InlineData("budget", true)]
    [InlineData("2025", true)]
    [InlineData("*.docx", true)]
    [InlineData("*.DOCX", true)]
    [InlineData("budget 202?.docx", true)]
    [InlineData("*.xlsx", false)]
    [InlineData("Budget", true)]
    [InlineData("Budget*", true)]
    [InlineData("*Budget", false)]
    [InlineData("forecast", false)]
    [InlineData("(", false)]
    public void Matches_names(string? pattern, bool matches) =>
        Assert.Equal(matches, CommandCatalog.NameMatches(Docx, pattern));

    [Fact]
    public void Matches_content_by_kind_and_check_out()
    {
        var folder = Docx with { Kind = NodeKind.Folder, Title = "Budgets" };
        var item = Contacts with { Kind = NodeKind.Item, Title = "Budget contact", ItemId = 2 };
        Assert.True(CommandCatalog.ContentMatches(Docx, "files", "budget", false));
        Assert.False(CommandCatalog.ContentMatches(folder, "files", null, false));
        Assert.True(CommandCatalog.ContentMatches(folder, "folders", "budget", false));
        Assert.True(CommandCatalog.ContentMatches(item, "items", "budget", false));
        Assert.True(CommandCatalog.ContentMatches(item, "all", null, false));
        Assert.False(CommandCatalog.ContentMatches(Documents, "all", null, false));
        Assert.False(CommandCatalog.ContentMatches(Docx, "files", null, true));
        Assert.True(CommandCatalog.ContentMatches(Docx with { CheckedOutTo = "Alex" }, "files", null, true));
    }

    [Fact]
    public void Finders_tick_and_only_read()
    {
        var finders = CommandCatalog.All.Where(c => c.After == After.Tick).ToList();
        Assert.Equal(["find.subsites", "find.lists", "find.home", "find.path", "find.content", "find.list-state", "find.search"], finders.Select(c => c.Id));
        Assert.All(finders, c =>
        {
            Assert.True(c.ReadOnly, c.Id);
            Assert.Equal("Tick inside", c.Group);
            Assert.True(c.AppliesTo(NodeKind.SiteCollection) && c.AppliesTo(NodeKind.Subsite), c.Id);
        });
        // inside lists, libraries and folders: their files, folders and items
        foreach (var kind in new[] { NodeKind.List, NodeKind.Library, NodeKind.Folder })
        {
            Assert.Equal(["find.content"], finders.Where(c => c.AppliesTo(kind)).Select(c => c.Id));
        }
        Assert.Equal("find.subsites", CommandCatalog.For(NodeKind.SiteCollection).SkipWhile(c => c.Group == "Open").First().Id);
    }

    [Fact]
    public void Asks_for_the_list_name_only_with_that_choice()
    {
        var lists = CommandCatalog.Find("find.lists")!;
        var name = lists.Parameters.Single(p => p.Name == "list");
        var args = lists.Defaults();
        Assert.False(name.IsShown(args));
        Assert.Empty(lists.Check(args));

        args["which"] = "name";
        Assert.True(name.IsShown(args));
        Assert.Single(lists.Check(args));
        args["list"] = "Contracts";
        Assert.Empty(lists.Check(args));
    }

    [Fact]
    public void Asks_which_lists_only_when_sites_are_ticked()
    {
        var content = CommandCatalog.Find("find.content")!;
        var which = content.Parameters.Single(p => p.Name == "which");
        var args = content.Defaults();
        Assert.True(which.IsShown(args, NodeKind.SiteCollection));
        Assert.False(which.IsShown(args, NodeKind.Library));

        // run on libraries, the hidden inputs aren't sent and aren't missed
        args["which"] = null;
        Assert.Empty(content.Check(args, [NodeKind.Library]));
        Assert.NotEmpty(content.Check(args, [NodeKind.Library, NodeKind.Subsite]));
    }
}
