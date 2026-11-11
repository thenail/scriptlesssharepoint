using System.Globalization;

using Microsoft.SharePoint.ApplicationPages.ClientPickerQuery;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>
/// Sharing as people do it in SharePoint: links to files, folders and items,
/// and invitations (by email) to them and to sites.
/// </summary>
public static partial class CommandCatalog
{
    /// <summary>The links Share makes, by SharePoint's SharingLinkKind number.</summary>
    private static readonly IReadOnlyList<Choice> LinkKinds =
    [
        new("2", "People in the organisation, who can view"),
        new("3", "People in the organisation, who can edit"),
        new("4", "Anyone with the link, who can view"),
        new("5", "Anyone with the link, who can edit"),
    ];

    /// <summary>Roles for ShareObject: Read and Contribute, as SharePoint's sharing dialog sends them.</summary>
    private const string ViewRole = "role:1073741826";
    private const string EditRole = "role:1073741827";

    private static IEnumerable<Command> SharingCommands()
    {
        yield return new Command
        {
            Id = "content.share-link",
            Title = "Make a sharing link",
            Group = "Sharing",
            Description = "Makes a link to each file, folder or item (or gives the one it has of that kind). Copy takes every link at once.",
            PowerShell = "Add-PnPFileSharingLink, Add-PnPFolderSharingLink",
            Targets = Content,
            Glyph = "",
            Parameters =
            [
                new("kind", "Link for", ParameterKind.Choice) { Choices = LinkKinds, Default = "2" },
                new("days", "Expires after (days)", ParameterKind.Number)
                {
                    Required = false,
                    Min = 1,
                    Max = 730,
                    Hint = "Empty: as the site's and tenant's settings say.",
                    ShowIf = new ShowIf("kind", "4", "5"),
                },
            ],
            Run = async (c, n) =>
            {
                var kind = int.Parse(c.Args.Text("kind"), CultureInfo.InvariantCulture);
                DateTimeOffset? expires = c.Args["days"] is null ? null : DateTimeOffset.UtcNow.Date.AddDays(c.Args.Number("days") + 1).AddSeconds(-1);
                var url = await c.Client.CreateSharingLinkAsync(n.WebUrl, ListOf(n), ItemOf(n), kind, expires, c.Cancellation).ConfigureAwait(false);
                var what = LinkKind((SharingLinkKind)kind);
                return CommandResult.Done(char.ToUpperInvariant(what[0]) + what[1..] + (expires is { } e ? $", until {e.LocalDateTime:d}." : "."), url);
            },
        };

        yield return new Command
        {
            Id = "content.links-remove",
            Title = "Remove sharing links",
            Group = "Sharing",
            Description = "Removes each one's \"Anyone\" links, its organisation links, or every sharing link. People given access in other ways keep it.",
            PowerShell = "Remove-PnPFileSharingLink, Remove-PnPFolderSharingLink",
            Targets = Content,
            Destructive = true,
            Glyph = "",
            Parameters =
            [
                new("which", "Remove", ParameterKind.Choice)
                {
                    Default = "anyone",
                    Choices = [new("anyone", "\"Anyone\" links"), new("organisation", "Organisation links"), new("all", "Every sharing link")],
                },
            ],
            Run = async (c, n) =>
            {
                List<(int Kind, string Id)> links;
                using (var ctx = await c.WebAsync(n).ConfigureAwait(false))
                {
                    var info = ObjectSharingInformation.GetListItemSharingInformation(ctx, ListOf(n), ItemOf(n), false, true, false, true, false, false);
                    ctx.Load(info, i => i.SharingLinks);
                    await ctx.RunAsync().ConfigureAwait(false);
                    var which = c.Args.Text("which");
                    links = (info.SharingLinks ?? [])
                        .Where(l => l.ShareId != Guid.Empty && !string.IsNullOrEmpty(l.Url) && which switch
                        {
                            "anyone" => l.LinkKind is SharingLinkKind.AnonymousView or SharingLinkKind.AnonymousEdit,
                            "organisation" => l.LinkKind is SharingLinkKind.OrganizationView or SharingLinkKind.OrganizationEdit,
                            _ => true,
                        })
                        .Select(l => ((int)l.LinkKind, l.ShareId.ToString()))
                        .ToList();
                }
                if (links.Count == 0)
                {
                    return CommandResult.Skipped("It has no such links.");
                }
                foreach (var (kind, id) in links)
                {
                    await c.Client.RemoveSharingLinkAsync(n.WebUrl, ListOf(n), ItemOf(n), kind, id, c.Cancellation).ConfigureAwait(false);
                }
                return CommandResult.Done($"Removed {Plural(links.Count, "link", "links")}.");
            },
        };

        yield return new Command
        {
            Id = "content.share-invite",
            Title = "Share with people",
            Group = "Sharing",
            Description = "Shares each file or folder with people, as Share does: they get access (guests too, where sharing allows) and, if asked, an email with a link.",
            PowerShell = "Add-PnPFileUserSharingLink, Add-PnPFolderUserSharingLink",
            Targets = FilesAndFolders,
            Glyph = "",
            Parameters =
            [
                People(),
                new("role", "They can", ParameterKind.Choice) { Default = "view", Choices = [new("view", "View"), new("edit", "Edit")] },
                Toggle("email", "Send them an email", true),
                new("message", "Message", ParameterKind.LongText) { Required = false, ShowIf = new ShowIf("email", "true") },
            ],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var picker = await PeoplePickerAsync(ctx, c.Args.Text("people")).ConfigureAwait(false);
                var result = Microsoft.SharePoint.Client.Web.ShareObject(ctx, TenantUrls.Resource(n.WebUrl) + n.ServerRelativeUrl, picker,
                    c.Args.Text("role") == "edit" ? EditRole : ViewRole, 0, false, c.Args.Toggle("email"), false, "", c.Args["message"] ?? "", true);
                return await SharedAsync(ctx, result, c.Args.Text("role") == "edit" ? "Shared: they can edit." : "Shared: they can view.").ConfigureAwait(false);
            },
        };

        yield return new Command
        {
            Id = "web.share-invite",
            Title = "Invite people to the site",
            Group = "Sharing",
            Description = "Adds people to one of each site's groups, as Share site does, and if asked emails them a link. Guests too, where sharing allows.",
            PowerShell = "Add-PnPGroupMember (with an invitation)",
            Targets = Webs,
            Glyph = "",
            Parameters =
            [
                People(),
                new("group", "As", ParameterKind.Choice) { Choices = Csom.SiteGroups, Default = "Members" },
                Toggle("email", "Send them an email", true),
                new("message", "Message", ParameterKind.LongText) { Required = false, ShowIf = new ShowIf("email", "true") },
            ],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var group = Csom.SiteGroup(ctx.Web, c.Args.Text("group"));
                ctx.Load(group, g => g.Id, g => g.Title);
                await ctx.RunAsync().ConfigureAwait(false);
                if (group.ServerObjectIsNull == true)
                {
                    return CommandResult.Failed($"It has no {c.Args.Text("group").ToLowerInvariant()} group (it may inherit its permissions).");
                }
                var picker = await PeoplePickerAsync(ctx, c.Args.Text("people")).ConfigureAwait(false);
                var result = Microsoft.SharePoint.Client.Web.ShareObject(ctx, n.WebUrl, picker, $"group:{group.Id}", group.Id, false,
                    c.Args.Toggle("email"), false, "", c.Args["message"] ?? "", true);
                return await SharedAsync(ctx, result, $"Invited: they're in {group.Title}.").ConfigureAwait(false);
            },
        };
    }

    private static CommandParameter People() =>
        new("people", "People", ParameterKind.LongText) { Placeholder = "alex@contoso.com", Hint = "Email addresses or user names: one per line, or separated by commas." };

    private static Guid ListOf(SpNode n) => n.ListId ?? throw new ArgumentException($"{n.Title} isn't in a list.");

    private static int ItemOf(SpNode n) => n.ItemId ?? throw new ArgumentException($"{n.Title} isn't a list item.");

    /// <summary>People as typed: one a line, or separated by commas or semicolons.</summary>
    internal static IReadOnlyList<string> PeopleOf(string typed) =>
        typed.Split(['\n', '\r', ',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

    /// <summary>
    /// The people as SharePoint's people picker resolves them (guests by email
    /// included), in the form ShareObject takes: a JSON array of picker entities.
    /// </summary>
    private static async Task<string> PeoplePickerAsync(ClientContext ctx, string typed)
    {
        var people = PeopleOf(typed);
        if (people.Count == 0)
        {
            throw new ArgumentException("Give at least one person.");
        }
        var resolved = people.Select(p => ClientPeoplePickerWebServiceInterface.ClientPeoplePickerResolveUser(ctx, new ClientPeoplePickerQueryParameters
        {
            AllowEmailAddresses = true,
            AllowMultipleEntities = false,
            MaximumEntitySuggestions = 30,
            PrincipalSource = PrincipalSource.All,
            PrincipalType = PrincipalType.User | PrincipalType.SecurityGroup | PrincipalType.SharePointGroup,
            QueryString = p,
        })).ToList();
        await ctx.RunAsync().ConfigureAwait(false);
        return "[" + string.Join(",", resolved.Select(r => r.Value)) + "]";
    }

    private static async Task<CommandResult> SharedAsync(ClientContext ctx, SharingResult result, string done)
    {
        ctx.Load(result, r => r.StatusCode, r => r.ErrorMessage);
        await ctx.RunAsync().ConfigureAwait(false);
        return result.StatusCode == SharingOperationStatusCode.CompletedSuccessfully
            ? CommandResult.Done(done)
            : CommandResult.Failed(string.IsNullOrEmpty(result.ErrorMessage) ? $"SharePoint said {result.StatusCode}." : result.ErrorMessage);
    }
}
