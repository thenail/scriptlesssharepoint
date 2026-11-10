using System.Text;

using Microsoft.SharePoint.Client;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>Opening things, and reading what they are.</summary>
public static partial class CommandCatalog
{
    private static IEnumerable<Command> Common()
    {
        yield return new Command
        {
            Id = "open",
            Title = "Open in the browser",
            Group = "Open",
            Description = "Opens each one in your browser.",
            Targets = Everything,
            Glyph = "\uE8A7",
            ReadOnly = true,
            OneAtATime = true,
            ConfirmAbove = 10,
            Run = (c, n) =>
            {
                c.Host.OpenUrl(n.Url);
                return Done("Opened.");
            },
        };

        yield return new Command
        {
            Id = "details",
            Title = "Show details",
            Group = "Details",
            Description = "What SharePoint says about each one: owner, storage, settings, dates.",
            PowerShell = "Get-SPOSite -Detailed, Get-PnPWeb, Get-PnPList, Get-PnPFile",
            Targets = Everything,
            Glyph = "\uE946",
            ReadOnly = true,
            Run = DetailsAsync,
        };
    }

    private static async Task<CommandResult> DetailsAsync(CommandContext c, SpNode n)
    {
        var lines = new List<(string Name, string? Value)>();
        switch (n.Kind)
        {
            case NodeKind.SiteCollection:
                try
                {
                    var (ctx, _, p) = await c.SitePropertiesAsync(n, detail: true).ConfigureAwait(false);
                    using (ctx)
                    {
                        lines.AddRange(
                        [
                            ("Title", p.Title),
                            ("URL", p.Url),
                            ("Template", p.Template),
                            ("Owner", p.OwnerEmail is { Length: > 0 } ? p.OwnerEmail : p.OwnerLoginName),
                            ("Storage used", $"{Csom.Size(p.StorageUsage * 1024 * 1024)} of {Csom.Size(p.StorageMaximumLevel * 1024 * 1024)}"),
                            ("Sharing", p.SharingCapability.ToString()),
                            ("Lock state", p.LockState),
                            ("Custom scripts", p.DenyAddAndCustomizePages == Microsoft.Online.SharePoint.TenantAdministration.DenyAddAndCustomizePagesStatus.Disabled ? "allowed" : "blocked"),
                            ("Hub site", p.IsHubSite ? "yes" : p.HubSiteId != Guid.Empty ? $"joined to {p.HubSiteId}" : "no"),
                            ("Microsoft 365 group", p.GroupId != Guid.Empty ? p.GroupId.ToString() : "none"),
                            ("Subsites", p.WebsCount.ToString("N0")),
                            ("Created", p.CreatedTime.ToString("u")),
                            ("Last content change", p.LastContentModifiedDate.ToString("u")),
                        ]);
                    }
                }
                catch (Exception e) when (SharePointClient.IsAccessDenied(e))
                {
                    // not an administrator: what the site itself says
                    await WebDetailsAsync(c, n, lines).ConfigureAwait(false);
                }
                break;
            case NodeKind.Subsite:
                await WebDetailsAsync(c, n, lines).ConfigureAwait(false);
                break;
            case NodeKind.List or NodeKind.Library:
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var l = Csom.List(ctx, n);
                ctx.Load(l, x => x.Title, x => x.ItemCount, x => x.Created, x => x.LastItemModifiedDate, x => x.BaseTemplate,
                    x => x.EnableVersioning, x => x.MajorVersionLimit, x => x.EnableMinorVersions, x => x.ForceCheckout,
                    x => x.EnableModeration, x => x.Hidden, x => x.NoCrawl, x => x.OnQuickLaunch, x => x.ContentTypesEnabled,
                    x => x.HasUniqueRoleAssignments, x => x.RootFolder.ServerRelativeUrl);
                await ctx.RunAsync().ConfigureAwait(false);
                lines.AddRange(
                [
                    ("Title", l.Title),
                    ("Address", l.RootFolder.ServerRelativeUrl),
                    ("Template", l.BaseTemplate.ToString()),
                    ("Items", l.ItemCount.ToString("N0")),
                    ("Versioning", l.EnableVersioning ? $"on, {l.MajorVersionLimit:N0} major versions{(l.EnableMinorVersions ? ", with minor versions" : "")}" : "off"),
                    ("Require check-out", Csom.Bool(l.ForceCheckout)),
                    ("Content approval", Csom.Bool(l.EnableModeration)),
                    ("Content types", Csom.Bool(l.ContentTypesEnabled)),
                    ("In search results", Csom.Bool(!l.NoCrawl)),
                    ("Hidden", Csom.Bool(l.Hidden)),
                    ("On the navigation", Csom.Bool(l.OnQuickLaunch)),
                    ("Permissions", l.HasUniqueRoleAssignments ? "its own" : "inherited"),
                    ("Created", l.Created.ToString("u")),
                    ("Last change", l.LastItemModifiedDate.ToString("u")),
                ]);
                break;
            }
            case NodeKind.File:
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var f = Csom.File(ctx, n);
                ctx.Load(f, x => x.Name, x => x.Length, x => x.TimeCreated, x => x.TimeLastModified, x => x.UIVersionLabel,
                    x => x.Level, x => x.CheckOutType, x => x.Author.Title, x => x.ModifiedBy.Title, x => x.ListItemAllFields.HasUniqueRoleAssignments);
                var versions = f.Versions;
                ctx.Load(versions, v => v.Include(x => x.Size));
                await ctx.RunAsync().ConfigureAwait(false);
                lines.AddRange(
                [
                    ("Name", f.Name),
                    ("Address", n.ServerRelativeUrl),
                    ("Size", Csom.Size(f.Length)),
                    ("Version", $"{f.UIVersionLabel} ({f.Level.ToString().ToLowerInvariant()})"),
                    ("Previous versions", $"{versions.Count:N0}, {Csom.Size(versions.Sum(v => (long)v.Size))}"),
                    ("Checked out", f.CheckOutType == CheckOutType.None ? "no" : n.CheckedOutTo ?? "yes"),
                    ("Created", $"{f.TimeCreated:u} by {f.Author.Title}"),
                    ("Modified", $"{f.TimeLastModified:u} by {f.ModifiedBy.Title}"),
                    ("Permissions", f.ListItemAllFields.HasUniqueRoleAssignments ? "its own" : "inherited"),
                ]);
                break;
            }
            case NodeKind.Folder or NodeKind.Item:
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var item = Csom.Item(ctx, n);
                ctx.Load(item, x => x.Id, x => x.HasUniqueRoleAssignments, x => x["Created"], x => x["Modified"], x => x["Author"], x => x["Editor"], x => x["Title"]);
                await ctx.RunAsync().ConfigureAwait(false);
                lines.AddRange(
                [
                    ("Title", n.Title),
                    ("ID", item.Id.ToString()),
                    ("Address", n.Kind == NodeKind.Folder ? n.ServerRelativeUrl : n.Url),
                    ("Items", n.Kind == NodeKind.Folder ? n.ItemCount?.ToString("N0") : null),
                    ("Created", $"{item["Created"]:u} by {(item["Author"] as FieldUserValue)?.LookupValue}"),
                    ("Modified", $"{item["Modified"]:u} by {(item["Editor"] as FieldUserValue)?.LookupValue}"),
                    ("Permissions", item.HasUniqueRoleAssignments ? "its own" : "inherited"),
                ]);
                break;
            }
        }
        var text = new StringBuilder();
        foreach (var (name, value) in lines.Where(l => !string.IsNullOrWhiteSpace(l.Value)))
        {
            text.Append(name).Append(": ").AppendLine(value);
        }
        return CommandResult.Done("Details below.", text.ToString().TrimEnd());
    }

    private static async Task WebDetailsAsync(CommandContext c, SpNode n, List<(string, string?)> lines)
    {
        using var ctx = await c.WebAsync(n).ConfigureAwait(false);
        var w = ctx.Web;
        ctx.Load(w, x => x.Title, x => x.Url, x => x.Description, x => x.WebTemplate, x => x.Configuration, x => x.Created,
            x => x.LastItemModifiedDate, x => x.Language, x => x.HasUniqueRoleAssignments, x => x.NoCrawl);
        if (n.Kind == NodeKind.SiteCollection)
        {
            ctx.Load(ctx.Site, s => s.Usage, s => s.Owner.Email, s => s.Owner.Title);
        }
        await ctx.RunAsync().ConfigureAwait(false);
        lines.AddRange(
        [
            ("Title", w.Title),
            ("URL", w.Url),
            ("Description", w.Description),
            ("Template", $"{w.WebTemplate}#{w.Configuration}"),
            ("Language", w.Language.ToString()),
        ]);
        if (n.Kind == NodeKind.SiteCollection)
        {
            lines.Add(("Owner", ctx.Site.Owner.Email is { Length: > 0 } e ? e : ctx.Site.Owner.Title));
            lines.Add(("Storage used", Csom.Size(ctx.Site.Usage.Storage)));
        }
        lines.AddRange(
        [
            ("Permissions", w.HasUniqueRoleAssignments ? "its own" : "inherited"),
            ("In search results", Csom.Bool(!w.NoCrawl)),
            ("Created", w.Created.ToString("u")),
            ("Last change", w.LastItemModifiedDate.ToString("u")),
        ]);
    }
}
