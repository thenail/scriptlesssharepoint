using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class ContentSearchTests
{
    private const string Site = "https://contoso.sharepoint.com/sites/hr";
    private static readonly Guid Docs = Guid.Parse("8f0f3a3c-2f4b-4c43-9a9e-3a0c5b4a9d01");

    [Fact]
    public void A_file_goes_under_its_library_and_site()
    {
        var hit = new SearchHit($"{Site}/Shared%20Documents/2024/Budget.xlsx", "Budget", Site, Site)
        {
            SiteTitle = "HR",
            FileName = "Budget.xlsx",
            ListId = Docs,
            ItemId = 12,
            IsDocument = true,
            ContentClass = "STS_ListItem_DocumentLibrary",
            Size = 2048,
        };
        var rows = ContentPresets.Rows(hit)!;
        Assert.Equal(NodeKind.File, rows.Node.Kind);
        Assert.Equal("Budget.xlsx", rows.Node.Title);
        Assert.Equal("/sites/hr/Shared Documents/2024/Budget.xlsx", rows.Node.ServerRelativeUrl);
        Assert.Equal((Docs, 12), (rows.Node.ListId!.Value, rows.Node.ItemId!.Value));
        Assert.Equal(NodeKind.Library, rows.Parent!.Kind);
        Assert.Equal("Shared Documents", rows.Parent.Title);
        Assert.Equal("/sites/hr/Shared Documents", rows.Parent.ServerRelativeUrl);
        Assert.Null(rows.Web);
        Assert.Equal("HR", rows.Site.Title);
        // the library leads to the file, and the site to the library: the table can stack them
        Assert.True(rows.Parent.Leads(rows.Node));
        Assert.Same(rows.Parent, rows.Node.NearestAbove([rows.Site, rows.Parent]));
    }

    [Fact]
    public void A_list_item_in_a_subsite_goes_under_its_list_and_subsite()
    {
        var web = Site + "/projects";
        var hit = new SearchHit($"{web}/Lists/Tasks/DispForm.aspx?ID=3", "Plan the move", Site, web)
        {
            ListId = Guid.NewGuid(),
            ItemId = 3,
            ContentClass = "STS_ListItem_Tasks",
        };
        var rows = ContentPresets.Rows(hit)!;
        Assert.Equal(NodeKind.Item, rows.Node.Kind);
        Assert.Equal(NodeKind.List, rows.Parent!.Kind);
        Assert.Equal("Tasks", rows.Parent.Title);
        Assert.Equal(NodeKind.Subsite, rows.Web!.Kind);
        Assert.Same(rows.Parent, rows.Node.NearestAbove([rows.Site, rows.Web, rows.Parent]));
        Assert.Same(rows.Web, rows.Parent.NearestAbove([rows.Site, rows.Web]));
    }

    [Fact]
    public void A_library_is_its_own_row_and_other_hits_are_none()
    {
        var library = ContentPresets.Rows(new SearchHit($"{Site}/Contracts/Forms/AllItems.aspx", "Contracts", Site, Site)
        {
            ListId = Docs,
            ContentClass = "STS_List_DocumentLibrary",
        })!;
        Assert.Equal(NodeKind.Library, library.Node.Kind);
        Assert.Equal("/sites/hr/Contracts", library.Node.ServerRelativeUrl);

        Assert.Null(ContentPresets.Rows(new SearchHit(Site, "HR", Site, Site) { ContentClass = "STS_Site" }));
        Assert.Null(ContentPresets.Rows(new SearchHit($"{Site}/Shared Documents/a.docx", "a", Site, Site) { IsDocument = true }));
    }

    [Fact]
    public void Predefined_searches_leave_out_onedrives_and_dates_move_with_today()
    {
        var old = ContentPresets.Find("files-old")!;
        var q = old.QueryAt(new DateTimeOffset(2026, 11, 10, 0, 0, 0, TimeSpan.Zero));
        Assert.Contains("LastModifiedTime<2023-11-10", q);
        Assert.EndsWith("-WebTemplate:SPSPERS", q);
        Assert.Equal("Size", ContentPresets.Find("files-100mb")!.SortBy);
    }

    [Fact]
    public void Every_predefined_selection_has_its_own_id()
    {
        var ids = SitePresets.All.Select(p => p.Id).Concat(ContentPresets.All.Select(p => p.Id)).ToList();
        Assert.Equal(ids.Count, ids.Distinct().Count());
        Assert.Contains(SitePresets.All, p => p.OneDrive);
        Assert.Contains(SitePresets.All, p => !p.OneDrive);
    }

    [Fact]
    public void Lists_by_their_state()
    {
        var library = new SpNode { Kind = NodeKind.Library, Title = "Docs", Url = Site + "/Docs", WebUrl = Site };
        var list = library with { Kind = NodeKind.List };
        var plain = new CommandCatalog.ListFacts(false, true, 10, false, false);
        Assert.True(CommandCatalog.ListInState(library, plain with { Unique = true }, "unique"));
        Assert.True(CommandCatalog.ListInState(library, plain with { Versioning = false }, "noversions"));
        Assert.False(CommandCatalog.ListInState(list, plain with { Versioning = false }, "noversions"));
        Assert.True(CommandCatalog.ListInState(list, plain with { Items = 6000 }, "large"));
        Assert.True(CommandCatalog.ListInState(list, plain with { Items = 0 }, "empty"));
        Assert.False(CommandCatalog.ListInState(list, plain, "empty"));
    }
}
