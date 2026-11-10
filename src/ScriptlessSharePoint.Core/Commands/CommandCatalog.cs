using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>
/// Every operation the app has, by what it acts on: the tenant, site collections,
/// subsites, lists and libraries, folders, files and list items. Together they
/// cover what the SharePoint Online Management Shell and PnP PowerShell are
/// used for on those things, through CSOM. Each says which cmdlet it stands in for.
/// </summary>
public static partial class CommandCatalog
{
    private static readonly NodeKind[] TenantOnly = [NodeKind.Tenant];
    private static readonly NodeKind[] Sites = [NodeKind.SiteCollection];
    private static readonly NodeKind[] Webs = [NodeKind.SiteCollection, NodeKind.Subsite];
    private static readonly NodeKind[] SubsitesOnly = [NodeKind.Subsite];
    private static readonly NodeKind[] Lists = [NodeKind.List, NodeKind.Library];
    private static readonly NodeKind[] Libraries = [NodeKind.Library];
    private static readonly NodeKind[] ListsOnly = [NodeKind.List];
    private static readonly NodeKind[] Containers = [NodeKind.List, NodeKind.Library, NodeKind.Folder];
    private static readonly NodeKind[] FileHolders = [NodeKind.Library, NodeKind.Folder];
    private static readonly NodeKind[] Content = [NodeKind.Folder, NodeKind.File, NodeKind.Item];
    private static readonly NodeKind[] FilesAndFolders = [NodeKind.Folder, NodeKind.File];
    private static readonly NodeKind[] Files = [NodeKind.File];
    private static readonly NodeKind[] FilesAndItems = [NodeKind.File, NodeKind.Item];
    private static readonly NodeKind[] Securables =
        [NodeKind.SiteCollection, NodeKind.Subsite, NodeKind.List, NodeKind.Library, NodeKind.Folder, NodeKind.File, NodeKind.Item];
    private static readonly NodeKind[] Inheriting =
        [NodeKind.Subsite, NodeKind.List, NodeKind.Library, NodeKind.Folder, NodeKind.File, NodeKind.Item];
    private static readonly NodeKind[] Everything = Securables;

    /// <summary>The order groups are listed in, in the pane.</summary>
    public static readonly IReadOnlyList<string> GroupOrder =
    [
        "Open",
        "Tick inside",
        "Details",
        "Create",
        "Content",
        "Versions",
        "Settings",
        "Sharing",
        "Permissions",
        "Administrators",
        "Groups",
        "Hub sites",
        "Apps and features",
        "Search",
        "Recycle bin",
        "Delete",
    ];

    /// <summary>
    /// Built on first use: the parts' static fields (choices, target sets) are in
    /// other files, whose initialisation order C# doesn't fix.
    /// </summary>
    private static readonly Lazy<IReadOnlyList<Command>> all = new(() =>
    [
        .. Common(),
        .. FindCommands(),
        .. Tenant(),
        .. SiteCollection(),
        .. Web(),
        .. Permissions(),
        .. ListCommands(),
        .. ContentCommands(),
    ]);

    public static IReadOnlyList<Command> All => all.Value;

    public static Command? Find(string id) => All.FirstOrDefault(c => c.Id == id);

    /// <summary>The commands for one kind of target, in the pane's order.</summary>
    public static IReadOnlyList<Command> For(NodeKind kind) =>
        All.Where(c => c.AppliesTo(kind))
            .OrderBy(c => GroupOrder.IndexOf(c.Group) is var i and >= 0 ? i : int.MaxValue)
            .ToList();

    private static int IndexOf(this IReadOnlyList<string> list, string value)
    {
        for (var i = 0; i < list.Count; i++)
        {
            if (list[i] == value)
            {
                return i;
            }
        }
        return -1;
    }

    // ---- shared parameter shapes ----

    private static CommandParameter Person(string label = "Person or group", string? hint = null) =>
        new("who", label, ParameterKind.Person)
        {
            Hint = hint ?? "An email address or user name (alex@contoso.com), or a group's claims name.",
            Placeholder = "alex@contoso.com",
        };

    private static CommandParameter Toggle(string name, string label, bool on, string? hint = null) =>
        new(name, label, ParameterKind.Toggle) { Default = on ? "true" : "false", Hint = hint };

    private static CommandParameter Comment(bool required = false) =>
        new("comment", "Comment", ParameterKind.LongText) { Required = required };

    /// <summary>"1 version", "3 versions".</summary>
    private static string Plural(int count, string one, string many) => $"{count:N0} {(count == 1 ? one : many)}";

    private static Task<CommandResult> Done(string message, string? detail = null) =>
        Task.FromResult(CommandResult.Done(message, detail));
}
