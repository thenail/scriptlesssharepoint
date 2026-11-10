using Microsoft.SharePoint.Client;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>What's in lists and libraries: folders, files and items.</summary>
public static partial class CommandCatalog
{
    /// <summary>The largest file uploaded in one request; bigger ones need an upload session.</summary>
    private const long MaxUpload = 250L * 1024 * 1024;

    private static IEnumerable<Command> ContentCommands()
    {
        // ---- create ----

        yield return new Command
        {
            Id = "content.new-folder",
            Title = "New folder",
            Group = "Create",
            Description = "Creates a folder in each list, library or folder.",
            PowerShell = "Add-PnPFolder",
            Targets = Containers,
            After = After.ReloadChildren,
            Glyph = "\uE8F4",
            Parameters = [new("name", "Name", ParameterKind.Text)],
            Run = async (c, n) =>
            {
                var name = c.Args.Text("name").Trim('/');
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                Csom.Folder(ctx, n).Folders.AddUsingPath(
                    ResourcePath.FromDecodedUrl($"{n.ServerRelativeUrl!.TrimEnd('/')}/{name}"),
                    new FolderCollectionAddParameters { Overwrite = false });
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Created {name}.");
            },
        };

        yield return new Command
        {
            Id = "content.upload",
            Title = "Upload a file",
            Group = "Create",
            Description = "Puts a file from this PC into each library or folder (up to 250 MB).",
            PowerShell = "Add-PnPFile",
            Targets = FileHolders,
            After = After.ReloadChildren,
            Glyph = "\uE898",
            Parameters =
            [
                new("file", "File", ParameterKind.LocalFile),
                Toggle("overwrite", "Replace a file with the same name", false),
            ],
            Run = async (c, n) =>
            {
                if (!n.InLibrary)
                {
                    return CommandResult.Skipped("Files go in libraries; this folder is in a list.");
                }
                var path = c.Args.Text("file");
                var info = new FileInfo(path);
                if (info.Length > MaxUpload)
                {
                    return CommandResult.Failed($"{info.Name} is over 250 MB: upload it in the browser.");
                }
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                await using var stream = System.IO.File.OpenRead(path);
                Csom.Folder(ctx, n).Files.Add(new FileCreationInformation
                {
                    ContentStream = stream,
                    Url = info.Name,
                    Overwrite = c.Args.Toggle("overwrite"),
                });
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Uploaded {info.Name}.");
            },
        };

        // ---- content ----

        yield return new Command
        {
            Id = "content.rename",
            Title = "Rename",
            Group = "Content",
            Description = "Gives each one a new name. {name} stands for its name as it is (files keep their extension when the new name has none).",
            PowerShell = "Rename-PnPFile, Rename-PnPFolder, Set-PnPListItem -Values @{Title}",
            Targets = Content,
            After = After.Update,
            Glyph = "\uE8AC",
            Parameters = [new("name", "New name", ParameterKind.Text) { Placeholder = "Archive - {name}" }],
            Run = async (c, n) =>
            {
                var current = n.Kind == NodeKind.File ? Path.GetFileNameWithoutExtension(n.LeafName) : n.Kind == NodeKind.Item ? n.Title : n.LeafName;
                var name = c.Args.Text("name").Replace("{name}", current, StringComparison.OrdinalIgnoreCase).Trim();
                if (n.Kind == NodeKind.File && Path.GetExtension(name).Length == 0)
                {
                    name += Path.GetExtension(n.LeafName);
                }
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var item = Csom.Item(ctx, n);
                if (n.Kind == NodeKind.Item)
                {
                    item["Title"] = name;
                    item.Update();
                    await ctx.RunAsync().ConfigureAwait(false);
                    return CommandResult.Done($"Renamed to {name}.") with { Updated = n with { Title = name } };
                }
                item["FileLeafRef"] = name;
                item.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                var path = Csom.WithLeaf(n.ServerRelativeUrl!, name);
                return CommandResult.Done($"Renamed to {name}.") with
                {
                    Updated = n with { Title = name, ServerRelativeUrl = path, Url = TenantUrls.Resource(n.WebUrl) + path },
                };
            },
        };

        yield return new Command
        {
            Id = "content.set-field",
            Title = "Set a column's value",
            Group = "Content",
            Description = "Sets one column of each item, file or folder. SharePoint reads the value as that column's type: a number, a date, a person's email address, a choice.",
            PowerShell = "Set-PnPListItem -Values",
            Targets = Content,
            Glyph = "\uE70F",
            Parameters =
            [
                new("field", "Column", ParameterKind.Text) { Hint = "Its internal name (as in the column's settings page address), e.g. Title or ProjectCode." },
                new("value", "Value", ParameterKind.Text) { Required = false, Hint = "Empty clears it." },
            ],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var item = Csom.Item(ctx, n);
                item.ParseAndSetFieldValue(c.Args.Text("field"), c.Args["value"] ?? "");
                item.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Set {c.Args.Text("field")}.");
            },
        };

        yield return MoveOrCopy(move: false);
        yield return MoveOrCopy(move: true);

        yield return new Command
        {
            Id = "content.download",
            Title = "Download",
            Group = "Content",
            Description = "Saves each file to a folder on this PC. A file of the same name there is kept, and the new one numbered.",
            PowerShell = "Get-PnPFile -AsFile",
            Targets = Files,
            Glyph = "\uE896",
            Parameters = [new("folder", "Save in", ParameterKind.LocalFolder)],
            Run = async (c, n) =>
            {
                var folder = c.Args.Text("folder");
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var data = Csom.File(ctx, n).OpenBinaryStream();
                await ctx.RunAsync().ConfigureAwait(false);
                var target = FreeName(folder, n.LeafName);
                await using (var output = new FileStream(target, FileMode.CreateNew, FileAccess.Write))
                {
                    await data.Value.CopyToAsync(output, c.Cancellation).ConfigureAwait(false);
                }
                return CommandResult.Done($"Saved {Path.GetFileName(target)}.");
            },
        };

        // ---- versions ----

        yield return new Command
        {
            Id = "file.check-out",
            Title = "Check out",
            Group = "Versions",
            Description = "Checks each file out to you, so nobody else can change it.",
            PowerShell = "Set-PnPFileCheckedOut",
            Targets = Files,
            Glyph = "\uE8A5",
            Run = (c, n) => FileAsync(c, n, f => f.CheckOut(), "Checked out."),
        };

        yield return new Command
        {
            Id = "file.check-in",
            Title = "Check in",
            Group = "Versions",
            Description = "Checks each file in, as a minor (draft) or major (published) version.",
            PowerShell = "Set-PnPFileCheckedIn",
            Targets = Files,
            Glyph = "\uE8FB",
            Parameters =
            [
                Comment(),
                new("type", "As", ParameterKind.Choice)
                {
                    Default = nameof(CheckinType.MajorCheckIn),
                    Choices =
                    [
                        new(nameof(CheckinType.MajorCheckIn), "A major version"),
                        new(nameof(CheckinType.MinorCheckIn), "A minor version"),
                        new(nameof(CheckinType.OverwriteCheckIn), "The same version (overwrite)"),
                    ],
                },
            ],
            Run = (c, n) => FileAsync(c, n, f => f.CheckIn(c.Args["comment"] ?? "", c.Args.Enum<CheckinType>("type")), "Checked in."),
        };

        yield return new Command
        {
            Id = "file.discard-check-out",
            Title = "Discard check-out",
            Group = "Versions",
            Description = "Undoes each file's check-out. Changes made while it was checked out are lost.",
            PowerShell = "Undo-PnPFileCheckedOut",
            Targets = Files,
            Destructive = true,
            Glyph = "\uE7A7",
            Run = (c, n) => FileAsync(c, n, f => f.UndoCheckOut(), "Check-out discarded."),
        };

        yield return new Command
        {
            Id = "file.publish",
            Title = "Publish a major version",
            Group = "Versions",
            Description = "Publishes each file's draft as a major version, for libraries that keep minor versions.",
            PowerShell = "Publish-PnPFile",
            Targets = Files,
            Glyph = "\uE898",
            Parameters = [Comment()],
            Run = (c, n) => FileAsync(c, n, f => f.Publish(c.Args["comment"] ?? ""), "Published."),
        };

        yield return new Command
        {
            Id = "file.unpublish",
            Title = "Unpublish",
            Group = "Versions",
            Description = "Takes each file's published version back to a draft.",
            PowerShell = "Unpublish-PnPFile",
            Targets = Files,
            Glyph = "\uE896",
            Parameters = [Comment()],
            Run = (c, n) => FileAsync(c, n, f => f.UnPublish(c.Args["comment"] ?? ""), "Unpublished."),
        };

        yield return new Command
        {
            Id = "content.approve",
            Title = "Approve",
            Group = "Versions",
            Description = "Approves each item or file, in lists and libraries that require approval.",
            PowerShell = "Set-PnPListItem -Values @{_ModerationStatus=0}",
            Targets = FilesAndItems,
            Glyph = "\uE8FB",
            Parameters = [Comment()],
            Run = (c, n) => ModerateAsync(c, n, approve: true),
        };

        yield return new Command
        {
            Id = "content.reject",
            Title = "Reject",
            Group = "Versions",
            Description = "Rejects each item or file, in lists and libraries that require approval.",
            PowerShell = "Set-PnPListItem -Values @{_ModerationStatus=1}",
            Targets = FilesAndItems,
            Glyph = "\uE711",
            Parameters = [Comment()],
            Run = (c, n) => ModerateAsync(c, n, approve: false),
        };

        yield return new Command
        {
            Id = "file.delete-versions",
            Title = "Delete previous versions",
            Group = "Versions",
            Description = "Deletes every earlier version of each file, keeping the current one: frees the space they take.",
            PowerShell = "Remove-PnPFileVersion -All",
            Targets = Files,
            Destructive = true,
            Glyph = "\uE74D",
            Run = (c, n) => FileAsync(c, n, f => f.Versions.DeleteAll(), "Deleted the previous versions."),
        };

        yield return new Command
        {
            Id = "file.versions",
            Title = "List versions",
            Group = "Versions",
            Description = "Each file's earlier versions: number, when, by whom, size and comment.",
            PowerShell = "Get-PnPFileVersion",
            Targets = Files,
            ReadOnly = true,
            Glyph = "\uE81C",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var file = Csom.File(ctx, n);
                ctx.Load(file, f => f.UIVersionLabel);
                ctx.Load(file.Versions, vs => vs.Include(v => v.VersionLabel, v => v.Created, v => v.CreatedBy.Title, v => v.Size, v => v.CheckInComment));
                await ctx.RunAsync().ConfigureAwait(false);
                var lines = file.Versions.ToList().AsEnumerable().Reverse().Select(v =>
                    $"{v.VersionLabel}  ·  {v.Created.ToLocalTime():g}  ·  {v.CreatedBy.Title}  ·  {Csom.Size(v.Size)}{(string.IsNullOrWhiteSpace(v.CheckInComment) ? "" : $"  ·  {v.CheckInComment}")}");
                return CommandResult.Done($"Version {file.UIVersionLabel}, {Plural(file.Versions.Count, "earlier one", "earlier ones")}.", string.Join(Environment.NewLine, lines));
            },
        };

        yield return new Command
        {
            Id = "file.restore-version",
            Title = "Restore an earlier version",
            Group = "Versions",
            Description = "Makes an earlier version of each file the current one, as a new version: nothing is lost.",
            PowerShell = "Restore-PnPFileVersion",
            Targets = Files,
            Glyph = "\uE777",
            Parameters = [new("label", "Version", ParameterKind.Text) { Required = false, Placeholder = "3.0", Hint = "Its number, as in the version history. Empty: the one before the current." }],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var file = Csom.File(ctx, n);
                ctx.Load(file.Versions, vs => vs.Include(v => v.VersionLabel, v => v.ID));
                await ctx.RunAsync().ConfigureAwait(false);
                var wanted = c.Args["label"]?.Trim();
                var version = wanted is null
                    ? file.Versions.ToList().MaxBy(v => v.ID)
                    : file.Versions.ToList().FirstOrDefault(v => v.VersionLabel == wanted);
                if (version is null)
                {
                    return CommandResult.Skipped(wanted is null ? "It has no earlier versions." : $"It has no version {wanted}.");
                }
                file.Versions.RestoreByLabel(version.VersionLabel);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Restored version {version.VersionLabel}.");
            },
        };

        yield return new Command
        {
            Id = "file.delete-old-versions",
            Title = "Delete older versions",
            Group = "Versions",
            Description = "Moves each file's earlier versions, beyond the newest ones to keep, to the recycle bin: frees their space. The current version stays.",
            PowerShell = "Remove-PnPFileVersion -Identity",
            Targets = Files,
            Destructive = true,
            Glyph = "\uE74D",
            Parameters = [new("keep", "Earlier versions to keep", ParameterKind.Number) { Default = "10", Min = 1, Max = 50000 }],
            Run = async (c, n) =>
            {
                var keep = (int)c.Args.Number("keep");
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var file = Csom.File(ctx, n);
                ctx.Load(file.Versions, vs => vs.Include(v => v.ID));
                await ctx.RunAsync().ConfigureAwait(false);
                var old = file.Versions.ToList().OrderByDescending(v => v.ID).Skip(keep).ToList();
                if (old.Count == 0)
                {
                    return CommandResult.Skipped($"It has {Plural(file.Versions.Count, "earlier version", "earlier versions")}.");
                }
                foreach (var v in old)
                {
                    file.Versions.RecycleByID(v.ID);
                }
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Moved {Plural(old.Count, "version", "versions")} to the recycle bin.");
            },
        };

        yield return new Command
        {
            Id = "content.stop-sharing",
            Title = "Stop sharing",
            Group = "Sharing",
            Description = "Removes each file's or folder's sharing links, and the access people were given by sharing it. Permissions given otherwise stay.",
            PowerShell = "Remove-PnPFileSharingLink, Remove-PnPFolderSharingLink",
            Targets = FilesAndFolders,
            Destructive = true,
            Glyph = "\uE711",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var result = Microsoft.SharePoint.Client.Web.UnshareObject(ctx, TenantUrls.Resource(n.WebUrl) + n.ServerRelativeUrl);
                ctx.Load(result, r => r.StatusCode, r => r.ErrorMessage);
                await ctx.RunAsync().ConfigureAwait(false);
                return result.StatusCode.ToString() == "CompletedSuccessfully"
                    ? CommandResult.Done("Stopped sharing it.")
                    : CommandResult.Failed(string.IsNullOrEmpty(result.ErrorMessage) ? $"SharePoint said {result.StatusCode}." : result.ErrorMessage);
            },
        };

        // ---- delete ----

        yield return new Command
        {
            Id = "content.delete",
            Title = "Delete",
            Group = "Delete",
            Description = "Moves each one, with what's in it, to the site's recycle bin.",
            PowerShell = "Remove-PnPFile / Remove-PnPFolder / Remove-PnPListItem -Recycle",
            Targets = Content,
            Destructive = true,
            After = After.Remove,
            Glyph = "\uE74D",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                Csom.Item(ctx, n).Recycle();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Deleted (in the recycle bin).");
            },
        };

        yield return new Command
        {
            Id = "content.delete-permanently",
            Title = "Delete for good",
            Group = "Delete",
            Description = "Deletes each one, with what's in it, without the recycle bin.",
            PowerShell = "Remove-PnPFile / Remove-PnPFolder / Remove-PnPListItem",
            Targets = Content,
            Destructive = true,
            After = After.Remove,
            Glyph = "\uE74D",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                Csom.Item(ctx, n).DeleteObject();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Deleted for good.");
            },
        };
    }

    private static Command MoveOrCopy(bool move) => new()
    {
        Id = move ? "content.move" : "content.copy",
        Title = move ? "Move to" : "Copy to",
        Group = "Content",
        Description = (move ? "Moves" : "Copies") + " each file or folder into another folder or library, in this site or another one (with versions and metadata, as SharePoint's own Move to and Copy to).",
        PowerShell = move ? "Move-PnPFile, Move-PnPFolder" : "Copy-PnPFile, Copy-PnPFolder",
        Targets = FilesAndFolders,
        After = move ? After.Remove : After.Nothing,
        Glyph = move ? "\uE8DE" : "\uE8C8",
        Parameters =
        [
            new("to", "Destination folder", ParameterKind.Url)
            {
                Placeholder = "https://contoso.sharepoint.com/sites/archive/Shared Documents/2024",
                Hint = "The folder's or library's address, as the browser shows it when it's open.",
            },
            Toggle("overwrite", "Replace files of the same name", false, "Off: a file of the same name there is kept, and the new one numbered."),
        ],
        Run = async (c, n) =>
        {
            using var ctx = await c.WebAsync(n).ConfigureAwait(false);
            var source = ResourcePath.FromDecodedUrl(TenantUrls.Resource(n.WebUrl) + n.ServerRelativeUrl);
            var target = ResourcePath.FromDecodedUrl(Csom.Destination(c.Args.Text("to"), n.LeafName));
            var overwrite = c.Args.Toggle("overwrite");
            var options = new MoveCopyOptions { KeepBoth = !overwrite, RetainEditorAndModifiedOnMove = true };
            switch (n.Kind, move)
            {
                case (NodeKind.File, true):
                    MoveCopyUtil.MoveFileByPath(ctx, source, target, overwrite, options);
                    break;
                case (NodeKind.File, false):
                    MoveCopyUtil.CopyFileByPath(ctx, source, target, overwrite, options);
                    break;
                case (_, true):
                    MoveCopyUtil.MoveFolderByPath(ctx, source, target, options);
                    break;
                default:
                    MoveCopyUtil.CopyFolderByPath(ctx, source, target, options);
                    break;
            }
            await ctx.RunAsync().ConfigureAwait(false);
            return CommandResult.Done(move ? "Moved." : "Copied.");
        },
    };

    private static async Task<CommandResult> FileAsync(CommandContext c, SpNode n, Action<Microsoft.SharePoint.Client.File> act, string done)
    {
        using var ctx = await c.WebAsync(n).ConfigureAwait(false);
        act(Csom.File(ctx, n));
        await ctx.RunAsync().ConfigureAwait(false);
        return CommandResult.Done(done);
    }

    private static async Task<CommandResult> ModerateAsync(CommandContext c, SpNode n, bool approve)
    {
        using var ctx = await c.WebAsync(n).ConfigureAwait(false);
        var item = Csom.Item(ctx, n);
        item["_ModerationStatus"] = approve ? 0 : 1;
        if (c.Args["comment"] is { } comment)
        {
            item["_ModerationComments"] = comment;
        }
        item.Update();
        await ctx.RunAsync().ConfigureAwait(false);
        return CommandResult.Done(approve ? "Approved." : "Rejected.");
    }

    /// <summary>A file name in a folder that isn't taken: "a.docx", else "a (2).docx"…</summary>
    internal static string FreeName(string folder, string name)
    {
        var path = Path.Combine(folder, name);
        var stem = Path.GetFileNameWithoutExtension(name);
        var ext = Path.GetExtension(name);
        for (var i = 2; System.IO.File.Exists(path); i++)
        {
            path = Path.Combine(folder, $"{stem} ({i}){ext}");
        }
        return path;
    }
}
