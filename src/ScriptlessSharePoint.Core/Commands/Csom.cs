using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;

using ScriptlessSharePoint.SharePoint;

using SpFile = Microsoft.SharePoint.Client.File;

namespace ScriptlessSharePoint.Commands;

/// <summary>The CSOM pieces the commands share: contexts, the objects a row stands for, users and roles.</summary>
internal static class Csom
{
    /// <summary>The longest a site operation (create, delete, restore) is waited for.</summary>
    private static readonly TimeSpan OperationLimit = TimeSpan.FromMinutes(10);

    public static Task<ClientContext> WebAsync(this CommandContext c, SpNode n) => c.Client.ContextAsync(n.WebUrl, c.Cancellation);

    /// <summary>A context for the target's site collection (its root web).</summary>
    public static Task<ClientContext> SiteAsync(this CommandContext c, SpNode n) => c.Client.ContextAsync(n.SiteUrl ?? n.WebUrl, c.Cancellation);

    public static Task<ClientContext> AdminAsync(this CommandContext c) => c.Client.ContextAsync(c.Client.AdminUrl, c.Cancellation);

    public static Task RunAsync(this ClientContext ctx) => ctx.ExecuteQueryAsync();

    public static List List(ClientContext ctx, SpNode n) =>
        ctx.Web.Lists.GetById(n.ListId ?? throw new ArgumentException($"{n.Title} isn't in a list."));

    public static ListItem Item(ClientContext ctx, SpNode n) =>
        List(ctx, n).GetItemById(n.ItemId ?? throw new ArgumentException($"{n.Title} isn't a list item."));

    public static SpFile File(ClientContext ctx, SpNode n) =>
        ctx.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(n.ServerRelativeUrl ?? throw new ArgumentException($"{n.Title} isn't a file.")));

    /// <summary>A folder row's folder; for a list or library, its root folder.</summary>
    public static Folder Folder(ClientContext ctx, SpNode n) => n.Kind is NodeKind.List or NodeKind.Library
        ? List(ctx, n).RootFolder
        : ctx.Web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(n.ServerRelativeUrl ?? throw new ArgumentException($"{n.Title} isn't a folder.")));

    /// <summary>What a row's permissions belong to: the web, the list, or the list item.</summary>
    public static SecurableObject Securable(ClientContext ctx, SpNode n) => n.Kind switch
    {
        NodeKind.SiteCollection or NodeKind.Subsite => ctx.Web,
        NodeKind.List or NodeKind.Library => List(ctx, n),
        NodeKind.Folder or NodeKind.File or NodeKind.Item => Item(ctx, n),
        _ => throw new ArgumentException($"{NodeKinds.Name(n.Kind)} has no permissions of its own."),
    };

    /// <summary>The permission levels to grant, as SharePoint names them on its sharing pages.</summary>
    public static readonly IReadOnlyList<Choice> Roles =
    [
        new("Reader", "Read"),
        new("Contributor", "Contribute"),
        new("Editor", "Edit"),
        new("WebDesigner", "Design"),
        new("Administrator", "Full control"),
    ];

    public static RoleType Role(string value) =>
        Enum.TryParse<RoleType>(value, out var r) && r is RoleType.Reader or RoleType.Contributor or RoleType.Editor or RoleType.WebDesigner or RoleType.Administrator
            ? r
            : throw new ArgumentException($"{value} isn't a permission level.");

    /// <summary>The web's own groups, as Add-PnPGroupMember -Group names them.</summary>
    public static readonly IReadOnlyList<Choice> SiteGroups =
    [
        new("Owners", "Owners"),
        new("Members", "Members"),
        new("Visitors", "Visitors"),
    ];

    public static Group SiteGroup(Web web, string which) => which switch
    {
        "Owners" => web.AssociatedOwnerGroup,
        "Members" => web.AssociatedMemberGroup,
        "Visitors" => web.AssociatedVisitorGroup,
        _ => throw new ArgumentException($"{which} isn't one of the site's groups."),
    };

    /// <summary>
    /// The site collection's properties, as the admin centre has them. Reading
    /// them with detail is slower: only for showing details.
    /// </summary>
    public static async Task<(ClientContext Ctx, Tenant Tenant, SiteProperties Props)> SitePropertiesAsync(this CommandContext c, SpNode n, bool detail = false)
    {
        var ctx = await c.AdminAsync().ConfigureAwait(false);
        var tenant = new Tenant(ctx);
        var props = tenant.GetSitePropertiesByUrl(n.SiteUrl ?? n.Url, detail);
        ctx.Load(props);
        await ctx.ExecuteQueryAsync().ConfigureAwait(false);
        return (ctx, tenant, props);
    }

    /// <summary>Changes a site collection's properties (Set-SPOSite) and waits for SharePoint to apply them.</summary>
    public static async Task<CommandResult> UpdateSiteAsync(this CommandContext c, SpNode n, Action<SiteProperties> change, string done)
    {
        var (ctx, _, props) = await c.SitePropertiesAsync(n).ConfigureAwait(false);
        using (ctx)
        {
            change(props);
            var op = props.Update();
            ctx.Load(op, o => o.IsComplete, o => o.PollingInterval);
            await ctx.ExecuteQueryAsync().ConfigureAwait(false);
            await WaitAsync(ctx, op, c.Cancellation).ConfigureAwait(false);
            return CommandResult.Done(done);
        }
    }

    /// <summary>
    /// Changes a property bag (a web's, or a list's root folder's). SharePoint
    /// refuses that where custom scripts are blocked, as they are on modern
    /// sites; then they're allowed for the moment, the change is made, and
    /// they're blocked again (needs a SharePoint administrator). Returns
    /// whether that was needed.
    /// </summary>
    public static async Task<bool> WithScriptsAsync(this CommandContext c, SpNode n, Func<Task> change)
    {
        try
        {
            await change().ConfigureAwait(false);
            return false;
        }
        catch (ServerUnauthorizedAccessException)
        {
        }

        ClientContext ctx;
        SiteProperties props;
        try
        {
            (ctx, _, props) = await c.SitePropertiesAsync(n).ConfigureAwait(false);
        }
        catch (Exception e) when (SharePointClient.IsAccessDenied(e))
        {
            throw new SharePointOperationException(
                "Access denied. Custom scripts are blocked on this site, so its settings for search and its property bag can't change. " +
                "A SharePoint administrator can do this (the app allows custom scripts for the moment and blocks them again).", e);
        }
        using (ctx)
        {
            if (props.DenyAddAndCustomizePages != DenyAddAndCustomizePagesStatus.Enabled)
            {
                throw new SharePointOperationException("Access denied. This needs the site's owner or a SharePoint administrator.");
            }
            await SetScriptsAsync(ctx, props, DenyAddAndCustomizePagesStatus.Disabled, c.Cancellation).ConfigureAwait(false);
            try
            {
                // allowing them can take a moment to reach the site
                for (var attempt = 1; ; attempt++)
                {
                    try
                    {
                        await change().ConfigureAwait(false);
                        return true;
                    }
                    catch (ServerUnauthorizedAccessException) when (attempt < 4)
                    {
                        await Task.Delay(TimeSpan.FromSeconds(5 * attempt), c.Cancellation).ConfigureAwait(false);
                    }
                }
            }
            finally
            {
                await SetScriptsAsync(ctx, props, DenyAddAndCustomizePagesStatus.Enabled, CancellationToken.None).ConfigureAwait(false);
            }
        }
    }

    private static async Task SetScriptsAsync(ClientContext ctx, SiteProperties props, DenyAddAndCustomizePagesStatus status, CancellationToken ct)
    {
        props.DenyAddAndCustomizePages = status;
        var op = props.Update();
        ctx.Load(op, o => o.IsComplete, o => o.PollingInterval);
        await ctx.ExecuteQueryAsync().ConfigureAwait(false);
        await WaitAsync(ctx, op, ct).ConfigureAwait(false);
    }

    /// <summary>"Search will crawl it again." and, when custom scripts were allowed for it, that they're blocked again.</summary>
    public static string ScriptsNote(string done, bool allowed) =>
        allowed ? $"{done} (Custom scripts were allowed for it, and blocked again.)" : done;

    /// <summary>Waits for a tenant operation (an SpoOperation) to finish, as the SPO cmdlets do unless -NoWait.</summary>
    public static async Task WaitAsync(ClientContext ctx, SpoOperation op, CancellationToken ct)
    {
        var until = DateTime.UtcNow + OperationLimit;
        if (!op.IsPropertyAvailable("IsComplete"))
        {
            ctx.Load(op, o => o.IsComplete, o => o.PollingInterval);
            await ctx.ExecuteQueryAsync().ConfigureAwait(false);
        }
        while (!op.IsComplete)
        {
            if (DateTime.UtcNow > until)
            {
                throw new SharePointOperationException("SharePoint is still working on it. Check again in a few minutes.");
            }
            await Task.Delay(Math.Clamp(op.PollingInterval, 1000, 15000), ct).ConfigureAwait(false);
            op.RefreshLoad();
            await ctx.ExecuteQueryAsync().ConfigureAwait(false);
        }
    }

    /// <summary>Finds a person or group in the web, adding them to its users as SharePoint needs.</summary>
    public static User EnsureUser(ClientContext ctx, string who) => ctx.Web.EnsureUser(who.Trim());

    /// <summary>"/sites/hr/Shared Documents/a.docx" with a new last part.</summary>
    public static string WithLeaf(string serverRelativeUrl, string leaf)
    {
        var url = serverRelativeUrl.TrimEnd('/');
        return url[..(url.LastIndexOf('/') + 1)] + leaf;
    }

    /// <summary>An absolute folder URL (as pasted from the browser) and a leaf name: where a copy or move goes.</summary>
    public static string Destination(string folderUrl, string leaf)
    {
        var url = Uri.UnescapeDataString(folderUrl.Trim()).TrimEnd('/');
        // a library's view page, as copied from the address bar: its folder
        var forms = url.IndexOf("/Forms/", StringComparison.OrdinalIgnoreCase);
        if (forms > 0)
        {
            url = url[..forms];
        }
        return $"{url}/{leaf}";
    }

    public static string Bool(bool b) => b ? "on" : "off";

    public static string Size(long bytes)
    {
        string[] units = ["bytes", "KB", "MB", "GB", "TB"];
        double v = bytes;
        var i = 0;
        while (v >= 1024 && i < units.Length - 1)
        {
            v /= 1024;
            i++;
        }
        return i == 0 ? $"{bytes:N0} bytes" : $"{v:0.#} {units[i]}";
    }
}
