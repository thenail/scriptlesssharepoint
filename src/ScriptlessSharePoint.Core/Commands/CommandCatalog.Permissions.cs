using System.Globalization;
using System.Text;

using Microsoft.SharePoint.Client;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>Who has access to what: on sites, lists, folders, files and items alike.</summary>
public static partial class CommandCatalog
{
    /// <summary>What a sharing link is, in words.</summary>
    private static string LinkKind(SharingLinkKind kind) => kind switch
    {
        SharingLinkKind.AnonymousView => "anyone with the link can view",
        SharingLinkKind.AnonymousEdit => "anyone with the link can edit",
        SharingLinkKind.OrganizationView => "people in the organisation can view",
        SharingLinkKind.OrganizationEdit => "people in the organisation can edit",
        SharingLinkKind.Direct or SharingLinkKind.Flexible => "specific people",
        _ => kind.ToString(),
    };

    private static IEnumerable<Command> Permissions()
    {
        yield return new Command
        {
            Id = "perm.show",
            Title = "Show permissions",
            Group = "Permissions",
            Description = "Who has access to each one and how, and whether it has its own permissions or inherits them.",
            PowerShell = "Get-PnPProperty RoleAssignments",
            Targets = Securables,
            ReadOnly = true,
            Glyph = "\uE8D7",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var o = Csom.Securable(ctx, n);
                ctx.Load(o, x => x.HasUniqueRoleAssignments);
                ctx.Load(o.RoleAssignments, r => r.Include(a => a.Member.Title, a => a.Member.LoginName, a => a.Member.PrincipalType,
                    a => a.RoleDefinitionBindings.Include(d => d.Name, d => d.Hidden)));
                await ctx.RunAsync().ConfigureAwait(false);
                var text = new StringBuilder();
                text.AppendLine(o.HasUniqueRoleAssignments ? "Its own permissions:" : "Inherited permissions:");
                foreach (var a in o.RoleAssignments.OrderBy(a => a.Member.Title, StringComparer.CurrentCultureIgnoreCase))
                {
                    var roles = string.Join(", ", a.RoleDefinitionBindings.Where(d => !d.Hidden).Select(d => d.Name));
                    if (roles.Length > 0)
                    {
                        text.AppendLine($"  {a.Member.Title}  ·  {roles}");
                    }
                }
                return CommandResult.Done(o.HasUniqueRoleAssignments ? "Has its own permissions." : "Inherits its permissions.", Join(text));
            },
        };

        yield return new Command
        {
            Id = "perm.check",
            Title = "Check someone's access",
            Group = "Permissions",
            Description = "What a person can do with each one, whichever way they got it (groups included): full control, edit, view or nothing.",
            Targets = Securables,
            ReadOnly = true,
            Glyph = "\uE77B",
            Parameters = [Person("Person", "Their email address or user name.")],
            Run = async (c, n) =>
            {
                var who = c.Args.Text("who").Trim();
                var login = who.Contains('|') || !who.Contains('@') ? who : $"i:0#.f|membership|{who}";
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var permissions = n.Kind switch
                {
                    NodeKind.SiteCollection or NodeKind.Subsite => ctx.Web.GetUserEffectivePermissions(login),
                    NodeKind.List or NodeKind.Library => Csom.List(ctx, n).GetUserEffectivePermissions(login),
                    _ => Csom.Item(ctx, n).GetUserEffectivePermissions(login),
                };
                await ctx.RunAsync().ConfigureAwait(false);
                var p = permissions.Value;
                var level = p.Has(PermissionKind.ManagePermissions) ? "full control"
                    : p.Has(PermissionKind.EditListItems) ? "can edit"
                    : p.Has(PermissionKind.ViewListItems) || p.Has(PermissionKind.ViewPages) ? "can view"
                    : "no access";
                return CommandResult.Done($"{who}: {level}.");
            },
        };

        yield return new Command
        {
            Id = "perm.shared-with",
            Title = "Who it's shared with",
            Group = "Sharing",
            Description = "The people each one is shared with, and its sharing links: who they work for, what they allow, and when they expire.",
            PowerShell = "Get-PnPFileSharingLink, Get-PnPFolderSharingLink",
            Targets = Securables,
            ReadOnly = true,
            Glyph = "\uE72D",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var info = n.Kind is NodeKind.SiteCollection or NodeKind.Subsite
                    ? ObjectSharingInformation.GetWebSharingInformation(ctx, false, true, false, true, true, false)
                    : ObjectSharingInformation.GetObjectSharingInformation(ctx, Csom.Securable(ctx, n), false, true, false, true, true, false, false);
                ctx.Load(info, i => i.SharedWithUsersCollection, i => i.SharingLinks);
                await ctx.RunAsync().ConfigureAwait(false);
                var text = new StringBuilder();
                var people = info.SharedWithUsersCollection?.ToList() ?? [];
                var links = info.SharingLinks?.Where(l => !string.IsNullOrEmpty(l.Url)).ToList() ?? [];
                foreach (var u in people.OrderBy(u => u.Name, StringComparer.CurrentCultureIgnoreCase))
                {
                    var can = u.HasEditPermission ? "can edit" : u.HasReviewPermission ? "can review" : "can view";
                    text.AppendLine($"{u.Name}{(string.IsNullOrEmpty(u.Email) ? "" : $" ({u.Email})")}  ·  {can}{(u.IsExternalUser ? "  ·  guest" : "")}");
                }
                foreach (var l in links)
                {
                    var expires = Convert.ToString(l.Expiration, CultureInfo.CurrentCulture);
                    text.AppendLine($"Link: {LinkKind(l.LinkKind)}{(string.IsNullOrWhiteSpace(expires) ? "" : $", expires {expires}")}  ·  {l.Url}");
                }
                return CommandResult.Done($"{Plural(people.Count, "person", "people")}, {Plural(links.Count, "link", "links")}.", Join(text));
            },
        };

        yield return new Command
        {
            Id = "perm.grant",
            Title = "Give access",
            Group = "Permissions",
            Description = "Gives a person or group a permission level. Where something inherits its permissions, it gets its own first (a copy of the inherited ones).",
            PowerShell = "Set-PnPWebPermission / Set-PnPListPermission / Set-PnPListItemPermission -AddRole",
            Targets = Securables,
            Glyph = "\uE8FA",
            Parameters =
            [
                Person(),
                new("role", "Permission level", ParameterKind.Choice) { Choices = Csom.Roles, Default = "Contributor" },
            ],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var o = Csom.Securable(ctx, n);
                ctx.Load(o, x => x.HasUniqueRoleAssignments);
                var user = Csom.EnsureUser(ctx, c.Args.Text("who"));
                await ctx.RunAsync().ConfigureAwait(false);
                if (!o.HasUniqueRoleAssignments && n.Kind != NodeKind.SiteCollection)
                {
                    o.BreakRoleInheritance(copyRoleAssignments: true, clearSubscopes: false);
                }
                var bindings = new RoleDefinitionBindingCollection(ctx) { ctx.Web.RoleDefinitions.GetByType(Csom.Role(c.Args.Text("role"))) };
                o.RoleAssignments.Add(user, bindings);
                await ctx.RunAsync().ConfigureAwait(false);
                var level = Csom.Roles.First(r => r.Value == c.Args.Text("role")).Label;
                return CommandResult.Done($"Gave {c.Args.Text("who")} {level.ToLowerInvariant()}.");
            },
        };

        yield return new Command
        {
            Id = "perm.revoke",
            Title = "Remove access",
            Group = "Permissions",
            Description = "Takes away the access a person or group was given directly. Where something inherits its permissions, there's nothing to take away there.",
            PowerShell = "Set-PnPListPermission -RemoveRole, Remove-PnPListItemPermission",
            Targets = Securables,
            Glyph = "\uE8AB",
            Parameters = [Person()],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var o = Csom.Securable(ctx, n);
                ctx.Load(o, x => x.HasUniqueRoleAssignments);
                var user = Csom.EnsureUser(ctx, c.Args.Text("who"));
                ctx.Load(user, u => u.Id);
                await ctx.RunAsync().ConfigureAwait(false);
                if (!o.HasUniqueRoleAssignments)
                {
                    return CommandResult.Skipped("It inherits its permissions: change them where they come from.");
                }
                var assignments = o.RoleAssignments;
                ctx.Load(assignments, r => r.Include(a => a.PrincipalId));
                await ctx.RunAsync().ConfigureAwait(false);
                var mine = assignments.FirstOrDefault(a => a.PrincipalId == user.Id);
                if (mine is null)
                {
                    return CommandResult.Skipped($"{c.Args.Text("who")} had no access of their own here.");
                }
                mine.DeleteObject();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Removed {c.Args.Text("who")}'s access.");
            },
        };

        yield return new Command
        {
            Id = "perm.break",
            Title = "Stop inheriting permissions",
            Group = "Permissions",
            Description = "Gives each one its own permissions, starting from a copy of the inherited ones (or from none).",
            PowerShell = "Set-PnPList -BreakRoleInheritance, Set-PnPListItemPermission",
            Targets = Inheriting,
            Glyph = "\uE785",
            Parameters = [Toggle("copy", "Start from the inherited permissions", true, "Off: only you keep access to start with.")],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var o = Csom.Securable(ctx, n);
                ctx.Load(o, x => x.HasUniqueRoleAssignments);
                await ctx.RunAsync().ConfigureAwait(false);
                if (o.HasUniqueRoleAssignments)
                {
                    return CommandResult.Skipped("It has its own permissions already.");
                }
                o.BreakRoleInheritance(c.Args.Toggle("copy"), clearSubscopes: false);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Has its own permissions now.");
            },
        };

        yield return new Command
        {
            Id = "perm.reset",
            Title = "Inherit permissions again",
            Group = "Permissions",
            Description = "Drops each one's own permissions, so it has its parent's again. Access given only here is lost.",
            PowerShell = "Set-PnPList -ResetRoleInheritance, Restore-PnPListItemPermission",
            Targets = Inheriting,
            Destructive = true,
            Glyph = "\uE72E",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var o = Csom.Securable(ctx, n);
                ctx.Load(o, x => x.HasUniqueRoleAssignments);
                await ctx.RunAsync().ConfigureAwait(false);
                if (!o.HasUniqueRoleAssignments)
                {
                    return CommandResult.Skipped("It inherits its permissions already.");
                }
                o.ResetRoleInheritance();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Inherits its permissions again.");
            },
        };
    }
}
