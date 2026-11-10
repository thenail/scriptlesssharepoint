using System.Globalization;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>What a command's input is, and so what the pane shows for it.</summary>
public enum ParameterKind
{
    Text,
    /// <summary>Several lines, e.g. a comment.</summary>
    LongText,
    /// <summary>A whole number, within <see cref="CommandParameter.Min"/> and <see cref="CommandParameter.Max"/>.</summary>
    Number,
    /// <summary>On or off: "true" or "false".</summary>
    Toggle,
    /// <summary>One of <see cref="CommandParameter.Choices"/>.</summary>
    Choice,
    /// <summary>A person or group: an email address, a user name or a claims login.</summary>
    Person,
    /// <summary>An https:// SharePoint URL.</summary>
    Url,
    /// <summary>A file on this PC, picked in the file dialog.</summary>
    LocalFile,
    /// <summary>A folder on this PC, picked in the folder dialog.</summary>
    LocalFolder,
}

public sealed record Choice(string Value, string Label);

/// <summary>An input a command asks for in the pane before it runs.</summary>
public sealed record CommandParameter(string Name, string Label, ParameterKind Kind)
{
    public bool Required { get; init; } = true;
    public string? Hint { get; init; }
    public string? Default { get; init; }
    public string? Placeholder { get; init; }
    public IReadOnlyList<Choice> Choices { get; init; } = [];
    public long? Min { get; init; }
    public long? Max { get; init; }

    /// <summary>What's wrong with a value, or null when it will do.</summary>
    public string? Check(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return Required && Kind != ParameterKind.Toggle ? $"{Label}: required." : null;
        }
        var v = value.Trim();
        switch (Kind)
        {
            case ParameterKind.Number:
                if (!long.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) &&
                    !long.TryParse(v, NumberStyles.Integer, CultureInfo.CurrentCulture, out n))
                {
                    return $"{Label}: enter a whole number.";
                }
                if (n < Min || n > Max)
                {
                    return $"{Label}: enter a number from {Min:N0} to {Max:N0}.";
                }
                break;
            case ParameterKind.Toggle when v is not ("true" or "false"):
                return $"{Label}: on or off.";
            case ParameterKind.Choice when Choices.All(c => c.Value != v):
                return $"{Label}: choose one of the options.";
            case ParameterKind.Url when !Uri.TryCreate(v, UriKind.Absolute, out var u) || u.Scheme != Uri.UriSchemeHttps:
                return $"{Label}: enter an https:// address.";
            case ParameterKind.LocalFile when !File.Exists(v):
                return $"{Label}: that file doesn't exist.";
            case ParameterKind.LocalFolder when !Directory.Exists(v):
                return $"{Label}: that folder doesn't exist.";
        }
        return null;
    }
}

/// <summary>The values given for a command's parameters, by name.</summary>
public sealed class CommandArgs
{
    private readonly Dictionary<string, string?> values;

    public CommandArgs(IReadOnlyDictionary<string, string?>? values = null) =>
        this.values = values is null ? new(StringComparer.OrdinalIgnoreCase) : new(values, StringComparer.OrdinalIgnoreCase);

    public string? this[string name]
    {
        get => values.TryGetValue(name, out var v) && !string.IsNullOrWhiteSpace(v) ? v.Trim() : null;
        set => values[name] = value;
    }

    public string Text(string name) => this[name] ?? throw new ArgumentException($"{name}: required.");

    public string? OptionalText(string name) => this[name];

    public long Number(string name) =>
        long.TryParse(Text(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ||
        long.TryParse(Text(name), NumberStyles.Integer, CultureInfo.CurrentCulture, out n)
            ? n
            : throw new ArgumentException($"{name}: not a number.");

    public bool Toggle(string name) => this[name] == "true";

    public T Enum<T>(string name) where T : struct, System.Enum =>
        System.Enum.TryParse<T>(Text(name), ignoreCase: true, out var v) ? v : throw new ArgumentException($"{name}: not one of the options.");

    public IReadOnlyDictionary<string, string?> ToDictionary() => values;
}

/// <summary>How one target went.</summary>
public sealed record CommandResult(bool Ok, string Message, string? Detail = null)
{
    public static CommandResult Done(string message, string? detail = null) => new(true, message, detail);
    public static CommandResult Skipped(string message) => new(true, message) { WasSkipped = true };
    public static CommandResult Failed(string message) => new(false, message);

    /// <summary>Nothing needed doing (already so).</summary>
    public bool WasSkipped { get; init; }

    /// <summary>The target as it is now (renamed): the table shows it so.</summary>
    public SpNode? Updated { get; init; }
}

/// <summary>What the table does with a row after a command ran on it.</summary>
public enum After
{
    /// <summary>Nothing changed that the table shows (reads, settings).</summary>
    Nothing,
    /// <summary>The row's own values changed (a new title): the result's <see cref="CommandResult.Updated"/> replaces it.</summary>
    Update,
    /// <summary>Something was added inside the row (a subsite, list, folder, file): load its contents again.</summary>
    ReloadChildren,
    /// <summary>The row is gone (deleted, moved): take it out of the table and the selection.</summary>
    Remove,
    /// <summary>The site collection list changed (a new or restored site): load it again.</summary>
    ReloadSites,
}

/// <summary>What commands may ask of the app.</summary>
public interface ICommandHost
{
    void OpenUrl(string url);
    void CopyText(string text);
}

/// <summary>Everything a command's run needs for one target.</summary>
public sealed class CommandContext(SharePointClient client, CommandArgs args, ICommandHost host, CancellationToken cancellation)
{
    public SharePointClient Client { get; } = client;
    public CommandArgs Args { get; } = args;
    public ICommandHost Host { get; } = host;
    public CancellationToken Cancellation { get; } = cancellation;
}

/// <summary>
/// One operation, as a PowerShell cmdlet (or a parameter of one) does it: what it
/// applies to, the inputs it asks for, and how it runs on one target. The runner
/// runs it on every selected target of those kinds.
/// </summary>
public sealed class Command
{
    public required string Id { get; init; }
    public required string Title { get; init; }

    /// <summary>The section it's listed under in the pane: "Sharing", "Permissions"…</summary>
    public required string Group { get; init; }

    public required string Description { get; init; }

    /// <summary>The PowerShell it stands in for, e.g. "Set-SPOSite -LockState" (SPO or PnP).</summary>
    public string? PowerShell { get; init; }

    public required IReadOnlyList<NodeKind> Targets { get; init; }

    public IReadOnlyList<CommandParameter> Parameters { get; init; } = [];

    /// <summary>Deletes or overwrites something: the app asks before running it.</summary>
    public bool Destructive { get; init; }

    /// <summary>Needs a SharePoint administrator (or an app with Sites.FullControl.All).</summary>
    public bool NeedsAdmin { get; init; }

    /// <summary>Only reads: says what it found in the results.</summary>
    public bool ReadOnly { get; init; }

    public After After { get; init; } = After.Nothing;

    /// <summary>Runs one target at a time, e.g. when targets could clash or the host is involved.</summary>
    public bool OneAtATime { get; init; }

    /// <summary>Asks before running on more targets than this (e.g. opening many browser tabs).</summary>
    public int? ConfirmAbove { get; init; }

    /// <summary>The Fluent icon glyph (Segoe Fluent Icons) shown beside it.</summary>
    public string Glyph { get; init; } = "\uE756";

    public required Func<CommandContext, SpNode, Task<CommandResult>> Run { get; init; }

    public bool AppliesTo(NodeKind kind) => Targets.Contains(kind);

    /// <summary>The problems with a set of values, one per parameter that has one.</summary>
    public IReadOnlyList<string> Check(CommandArgs args) =>
        Parameters.Select(p => p.Check(args[p.Name])).OfType<string>().ToList();

    /// <summary>Default values for the parameters that have them.</summary>
    public CommandArgs Defaults()
    {
        var args = new CommandArgs();
        foreach (var p in Parameters)
        {
            args[p.Name] = p.Default ?? (p.Kind == ParameterKind.Toggle ? "false" : null);
        }
        return args;
    }

    public override string ToString() => Title;
}
