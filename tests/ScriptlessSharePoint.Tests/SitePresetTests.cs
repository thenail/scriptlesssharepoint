using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class SitePresetTests
{
    private static readonly DateTimeOffset Now = new(2026, 11, 10, 12, 0, 0, TimeSpan.Zero);

    private static SiteFacts Site(string template = "STS#3", int daysQuiet = 1, int age = 400) =>
        new("https://contoso.sharepoint.com/sites/x", "X", template, Now.AddDays(-daysQuiet), Now.AddDays(-age), Guid.NewGuid())
        {
            Sharing = "Disabled",
            LockState = "Unlock",
            ArchiveStatus = "NotArchived",
            SensitivityLabel = Guid.NewGuid(),
        };

    private static bool Picks(string id, SiteFacts site) => SitePresets.Find(id)!.Matches(site, Now);

    [Fact]
    public void Ids_are_unique_and_each_has_a_title_and_an_icon()
    {
        Assert.Equal(SitePresets.All.Count, SitePresets.All.Select(p => p.Id).Distinct().Count());
        Assert.All(SitePresets.All, p => Assert.True(p.Title.Length > 0 && p.Description.Length > 0 && p.Glyph.Length == 1, p.Id));
    }

    [Fact]
    public void Inactive_sites_by_last_change()
    {
        Assert.True(Picks("inactive-90", Site(daysQuiet: 120)));
        Assert.False(Picks("inactive-90", Site(daysQuiet: 30)));
        Assert.False(Picks("inactive-365", Site(daysQuiet: 120)));
        Assert.True(Picks("inactive-365", Site(daysQuiet: 400)));
        Assert.False(Picks("inactive-90", Site() with { LastModified = null }));
        Assert.True(Picks("new-30", Site(age: 3)));
        Assert.False(Picks("new-30", Site(age: 60)));
    }

    [Fact]
    public void Hubs_and_what_belongs_to_them()
    {
        var hub = Site() with { IsHubSite = true, HubSiteId = Guid.NewGuid() };
        var member = Site() with { HubSiteId = Guid.NewGuid() };
        var alone = Site();
        Assert.Equal([true, false, false], new[] { hub, member, alone }.Select(s => Picks("hubs", s)));
        Assert.Equal([false, true, false], new[] { hub, member, alone }.Select(s => Picks("hub-connected", s)));
        Assert.Equal([false, false, true], new[] { hub, member, alone }.Select(s => Picks("no-hub", s)));
    }

    [Fact]
    public void Kinds_of_site()
    {
        Assert.True(Picks("communication", Site("SITEPAGEPUBLISHING#0")));
        Assert.True(Picks("classic", Site("STS#0")));
        Assert.False(Picks("classic", Site("STS#3")));
        Assert.True(Picks("groups", Site("GROUP#0") with { GroupId = Guid.NewGuid() }));
        Assert.True(Picks("teams", Site("GROUP#0") with { IsTeamsConnected = true }));
    }

    [Fact]
    public void Sharing_lock_storage_label_and_archive()
    {
        Assert.True(Picks("anyone", Site() with { Sharing = "ExternalUserAndGuestSharing" }));
        Assert.True(Picks("guests", Site() with { Sharing = "ExistingExternalUserSharingOnly" }));
        Assert.False(Picks("guests", Site()));
        Assert.True(Picks("locked", Site() with { LockState = "ReadOnly" }));
        Assert.False(Picks("locked", Site()));
        Assert.True(Picks("storage", Site() with { StorageUsedMb = 950, StorageLimitMb = 1000 }));
        Assert.False(Picks("storage", Site() with { StorageUsedMb = 500, StorageLimitMb = 1000 }));
        Assert.False(Picks("storage", Site() with { StorageUsedMb = 500, StorageLimitMb = 0 }));
        Assert.True(Picks("large", Site() with { StorageUsedMb = 30 * 1024 }));
        Assert.True(Picks("no-label", Site() with { SensitivityLabel = Guid.Empty }));
        Assert.False(Picks("no-label", Site()));
        Assert.True(Picks("archived", Site() with { ArchiveStatus = "RecentlyArchived" }));
        Assert.False(Picks("archived", Site()));
    }
}
