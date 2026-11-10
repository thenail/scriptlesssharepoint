namespace ScriptlessSharePoint.SharePoint;

/// <summary>
/// What the admin centre knows about a site collection, for picking sites by
/// it: activity, hub, group and Teams, sharing, lock, storage and label.
/// Storage is in MB; <see cref="Sharing"/> and <see cref="LockState"/> are
/// SharePoint's own names ("ExternalUserAndGuestSharing", "ReadOnly").
/// </summary>
public sealed record SiteFacts(
    string Url,
    string Title,
    string? Template,
    DateTimeOffset? LastModified,
    DateTimeOffset? Created,
    Guid? SiteId)
{
    public Guid HubSiteId { get; init; }
    public bool IsHubSite { get; init; }
    public Guid GroupId { get; init; }
    public bool IsTeamsConnected { get; init; }
    public string? Sharing { get; init; }
    public string? LockState { get; init; }
    public long StorageUsedMb { get; init; }
    public long StorageLimitMb { get; init; }
    public Guid SensitivityLabel { get; init; }
    public string? ArchiveStatus { get; init; }

    public SiteSummary Summary => new(Url, Title, Template, LastModified, SiteId);
}

/// <summary>A ready-made selection: the site collections (or, with <see cref="OneDrive"/>, the OneDrives) a rule picks.</summary>
public sealed record SitePreset(string Id, string Title, string Description, string Glyph, Func<SiteFacts, DateTimeOffset, bool> Matches)
{
    public bool OneDrive { get; init; }
}

/// <summary>The predefined selections, in the order the navigation pane lists them.</summary>
public static class SitePresets
{
    /// <summary>Site templates from before modern SharePoint.</summary>
    private static readonly HashSet<string> ClassicTemplates = new(StringComparer.OrdinalIgnoreCase)
    {
        "STS#0", "STS#1", "STS#2", "BLOG#0", "BDR#0", "PROJECTSITE#0", "ENTERWIKI#0", "BLANKINTERNET#0",
        "BLANKINTERNETCONTAINER#0", "CMSPUBLISHING#0", "COMMUNITY#0", "COMMUNITYPORTAL#0", "OFFILE#1", "DEV#0", "EHS#1", "VISPRUS#0",
    };

    public static readonly IReadOnlyList<SitePreset> All =
    [
        new("inactive-90", "Inactive for 90 days", "Nothing in them has changed for 90 days or more.", "",
            (s, now) => s.LastModified is { } m && m < now.AddDays(-90)),
        new("inactive-365", "Inactive for a year", "Nothing in them has changed for a year or more.", "",
            (s, now) => s.LastModified is { } m && m < now.AddDays(-365)),
        new("new-30", "Created in the last 30 days", "Site collections made in the last 30 days.", "",
            (s, now) => s.Created is { } c && c >= now.AddDays(-30)),
        new("hubs", "Hub sites", "Site collections that are hubs.", "",
            (s, _) => s.IsHubSite),
        new("hub-connected", "Connected to a hub", "Site collections that belong to a hub (hubs themselves left out).", "",
            (s, _) => s.HubSiteId != Guid.Empty && !s.IsHubSite),
        new("no-hub", "Not in a hub", "Site collections that neither are nor belong to a hub.", "",
            (s, _) => s.HubSiteId == Guid.Empty && !s.IsHubSite),
        new("groups", "Microsoft 365 group sites", "Team sites that come with a Microsoft 365 group.", "",
            (s, _) => s.GroupId != Guid.Empty),
        new("teams", "Connected to Teams", "Group sites whose group has a team in Microsoft Teams.", "",
            (s, _) => s.IsTeamsConnected),
        new("communication", "Communication sites", "Sites for news and information many people read.", "",
            (s, _) => s.Template?.StartsWith("SITEPAGEPUBLISHING", StringComparison.OrdinalIgnoreCase) == true),
        new("classic", "Classic sites", "Sites made from classic templates (team, blog, publishing, wiki…), candidates for modernising.", "",
            (s, _) => s.Template is { } t && ClassicTemplates.Contains(t)),
        new("anyone", "\"Anyone\" links allowed", "Site collections whose sharing allows links that need no sign-in.", "",
            (s, _) => s.Sharing == "ExternalUserAndGuestSharing"),
        new("guests", "Guests allowed", "Site collections that can be shared with people outside the organisation.", "",
            (s, _) => s.Sharing is { } x && x != "Disabled"),
        new("locked", "Locked or read-only", "Site collections that are read-only, or shut to everyone.", "",
            (s, _) => s.LockState is { } l && !string.Equals(l, "Unlock", StringComparison.OrdinalIgnoreCase)),
        new("storage", "Near the storage limit", "Site collections using 90 % or more of their storage limit.", "",
            (s, _) => s.StorageLimitMb > 0 && s.StorageUsedMb >= s.StorageLimitMb * 9 / 10),
        new("large", "Over 25 GB", "Site collections storing more than 25 GB.", "",
            (s, _) => s.StorageUsedMb > 25 * 1024),
        new("no-label", "Without a sensitivity label", "Site collections with no sensitivity label.", "",
            (s, _) => s.SensitivityLabel == Guid.Empty),
        new("archived", "Archived", "Site collections in Microsoft 365 Archive, or on their way in or out.", "",
            (s, _) => s.ArchiveStatus is { Length: > 0 } a && !string.Equals(a, "NotArchived", StringComparison.OrdinalIgnoreCase)),

        // ---- OneDrives ----

        new("od-inactive-90", "Inactive for 90 days", "OneDrives where nothing has changed for 90 days or more.", "\uE823",
            (s, now) => s.LastModified is { } m && m < now.AddDays(-90)) { OneDrive = true },
        new("od-inactive-365", "Inactive for a year", "OneDrives where nothing has changed for a year or more: often people who have left.", "\uE823",
            (s, now) => s.LastModified is { } m && m < now.AddDays(-365)) { OneDrive = true },
        new("od-large", "Over 100 GB", "OneDrives storing more than 100 GB.", "\uEDA2",
            (s, _) => s.StorageUsedMb > 100 * 1024) { OneDrive = true },
        new("od-storage", "Near the storage limit", "OneDrives using 90 % or more of their storage limit.", "\uEDA2",
            (s, _) => s.StorageLimitMb > 0 && s.StorageUsedMb >= s.StorageLimitMb * 9 / 10) { OneDrive = true },
        new("od-anyone", "\"Anyone\" links allowed", "OneDrives whose sharing allows links that need no sign-in.", "\uE774",
            (s, _) => s.Sharing == "ExternalUserAndGuestSharing") { OneDrive = true },
    ];

    public static SitePreset? Find(string id) => All.FirstOrDefault(p => p.Id == id);
}
