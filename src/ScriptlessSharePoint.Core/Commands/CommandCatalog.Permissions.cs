using System.Text;

using Microsoft.SharePoint.Client;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>Who has access to what: on sites, lists, folders, files and items alike.</summary>
public static partial class CommandCatalog
{
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
