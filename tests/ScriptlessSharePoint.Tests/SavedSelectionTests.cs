using ScriptlessSharePoint.Selections;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class SavedSelectionTests : IDisposable
{
    private const string Site = "https://contoso.sharepoint.com/sites/hr";
    private static readonly Guid Docs = Guid.NewGuid();

    private readonly string dir = Path.Combine(Path.GetTempPath(), "ssp-saved-" + Guid.NewGuid().ToString("n"));
    private readonly SavedSelectionStore store;

    public SavedSelectionTests() => store = new SavedSelectionStore(dir);

    public void Dispose()
    {
        if (Directory.Exists(dir))
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    private static readonly SpNode SiteNode = SpNode.FromSite(new SiteSummary(Site, "Human Resources", "GROUP#0", null, null));

    private static readonly SpNode Subsite = new()
    {
        Kind = NodeKind.Subsite, Title = "Policies", Url = Site + "/policies", WebUrl = Site + "/policies", SiteUrl = Site,
    };

    private static readonly SpNode Library = new()
    {
        Kind = NodeKind.Library, Title = "Documents", Url = Site + "/policies/Shared Documents", WebUrl = Site + "/policies", SiteUrl = Site,
        ListId = Docs, ServerRelativeUrl = "/sites/hr/policies/Shared Documents", InLibrary = true,
    };

    private static readonly SpNode Folder = Library with
    {
        Kind = NodeKind.Folder, Title = "2025", Url = Library.Url + "/2025", ServerRelativeUrl = Library.ServerRelativeUrl + "/2025", ItemId = 3,
    };

    private static readonly SpNode File = Library with
    {
        Kind = NodeKind.File, Title = "Leave.docx", Url = Folder.Url + "/Leave.docx", ServerRelativeUrl = Folder.ServerRelativeUrl + "/Leave.docx", ItemId = 9, Size = 1234,
    };

    [Fact]
    public void Saves_lists_renames_and_removes()
    {
        var saved = store.Save(" Leave policy ", "conn1", [File, File, Library], [SiteNode]);
        Assert.Equal("Leave policy", saved.Name);
        Assert.Equal(2, saved.Items.Count);

        var listed = Assert.Single(store.List());
        Assert.Equal(saved.File, listed.File);
        Assert.Equal("conn1", listed.ConnectionId);
        Assert.Equal(File, listed.Items[0]);
        Assert.Equal(SiteNode, Assert.Single(listed.Sites));

        Assert.Equal("Renamed", store.Rename(saved.File, "Renamed")!.Name);
        Assert.Equal("Renamed", store.List()[0].Name);

        store.Delete(saved.File);
        Assert.Empty(store.List());
    }

    [Fact]
    public void A_connections_selections_go_with_it()
    {
        store.Save("a", "conn1", [File], [SiteNode]);
        store.Save("b", "conn2", [File], [SiteNode]);
        Assert.Equal(1, store.DeleteFor("conn1"));
        Assert.Equal("b", Assert.Single(store.List()).Name);
    }

    [Fact]
    public void Files_that_cant_be_read_are_left_out()
    {
        store.Save("good", "conn1", [File], [SiteNode]);
        System.IO.File.WriteAllText(Path.Combine(dir, "broken.json"), "{ nope");
        Assert.Equal("good", Assert.Single(store.List()).Name);
    }

    [Fact]
    public void Removing_never_touches_files_outside_the_folder()
    {
        var outside = Path.Combine(Path.GetTempPath(), $"ssp-outside-{Guid.NewGuid():n}.json");
        System.IO.File.WriteAllText(outside, "{}");
        try
        {
            store.Delete(outside);
            Assert.True(System.IO.File.Exists(outside));
        }
        finally
        {
            System.IO.File.Delete(outside);
        }
    }

    [Fact]
    public void The_way_down_to_a_file_goes_through_its_site_subsite_library_and_folders()
    {
        Assert.True(SiteNode.Leads(File));
        Assert.True(Subsite.Leads(File));
        Assert.True(Library.Leads(File));
        Assert.True(Folder.Leads(File));
        Assert.True(File.Leads(File));

        var otherFolder = Folder with { Url = Library.Url + "/2024", ServerRelativeUrl = Library.ServerRelativeUrl + "/2024" };
        Assert.False(otherFolder.Leads(File));
        // a folder whose name starts the same isn't the way
        var lookalike = Folder with { Url = Library.Url + "/20", ServerRelativeUrl = Library.ServerRelativeUrl + "/20" };
        Assert.False(lookalike.Leads(File));
        Assert.False((Library with { ListId = Guid.NewGuid() }).Leads(File));
        Assert.False((Subsite with { Url = Site + "/pol", WebUrl = Site + "/pol" }).Leads(File));
        Assert.False(SpNode.FromSite(new SiteSummary(Site + "2", "HR 2", null, null, null)).Leads(File));
    }

    [Fact]
    public void A_site_known_only_by_its_address_is_named_by_its_path()
    {
        var site = SpNode.SiteFromUrl(Site + "/");
        Assert.Equal("/sites/hr", site.Title);
        Assert.Equal(SiteNode.Key, site.Key);
    }
}
