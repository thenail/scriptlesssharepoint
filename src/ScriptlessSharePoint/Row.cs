using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;

using Microsoft.UI.Xaml;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>What a row of the table is: something in SharePoint, "show more", or a note (empty, an error).</summary>
public enum RowKind
{
    Node,
    More,
    Message,
}

/// <summary>What the table asks of the window: opening rows and ticking them.</summary>
public interface IRowOwner
{
    void Toggle(Row row);
    void Checked(Row row);
    void LoadMore(Row row);
}

/// <summary>
/// One row of the table: a site collection, subsite, list, library, folder,
/// file or item, indented under what it's in. Rows below an open row follow it
/// in the list, as File Size opens folders in place.
/// </summary>
public sealed class Row : INotifyPropertyChanged
{
    private const double IndentStep = 22;

    private readonly IRowOwner owner;
    private SpNode? node;
    private bool isExpanded;
    private bool isLoading;
    private bool isChecked;
    private string message;

    private Row(IRowOwner owner, RowKind kind, int depth, Row? parent, SpNode? node, string message)
    {
        this.owner = owner;
        Kind = kind;
        Depth = depth;
        Parent = parent;
        this.node = node;
        this.message = message;
    }

    public static Row ForNode(IRowOwner owner, SpNode node, Row? parent, bool isChecked) =>
        new(owner, RowKind.Node, parent is null ? 0 : parent.Depth + 1, parent, node, "") { isChecked = isChecked };

    public static Row More(IRowOwner owner, Row? parent, string cursor, string text) =>
        new(owner, RowKind.More, parent is null ? 0 : parent.Depth + 1, parent, null, text) { Cursor = cursor };

    public static Row Note(IRowOwner owner, Row? parent, string text) =>
        new(owner, RowKind.Message, parent is null ? 0 : parent.Depth + 1, parent, null, text);

    public event PropertyChangedEventHandler? PropertyChanged;

    public RowKind Kind { get; }
    public int Depth { get; }
    public Row? Parent { get; }

    /// <summary>"Show more": where the next page starts.</summary>
    public string? Cursor { get; set; }

    public SpNode? Node
    {
        get => node;
        set
        {
            node = value;
            Changed(nameof(Title));
            Changed(nameof(OpenName));
            Changed(nameof(TickName));
            Changed(nameof(TypeText));
            Changed(nameof(ItemsText));
            Changed(nameof(SizeText));
            Changed(nameof(ModifiedText));
            Changed(nameof(ToolTip));
        }
    }

    public bool IsNode => Kind == RowKind.Node;
    public bool IsMore => Kind == RowKind.More;
    public bool IsMessage => Kind == RowKind.Message;

    public Thickness Indent => new(Depth * IndentStep, 0, 0, 0);

    public bool CanExpand => node?.CanExpand == true;

    public bool IsExpanded
    {
        get => isExpanded;
        set
        {
            if (Set(ref isExpanded, value))
            {
                Changed(nameof(ChevronGlyph));
                Changed(nameof(OpenName));
            }
        }
    }

    public bool IsLoading
    {
        get => isLoading;
        set
        {
            if (Set(ref isLoading, value))
            {
                Changed(nameof(ShowChevron));
            }
        }
    }

    public bool ShowChevron => CanExpand && !isLoading;

    /// <summary>Keeps the names lined up whether or not a row can open.</summary>
    public double ChevronOpacity => CanExpand ? 1 : 0;

    public string ChevronGlyph => isExpanded ? "" : "";

    /// <summary>Ticked: in the selection the pane acts on.</summary>
    public bool IsChecked
    {
        get => isChecked;
        set
        {
            if (Set(ref isChecked, value))
            {
                owner.Checked(this);
            }
        }
    }

    /// <summary>Sets the tick from the selection, without telling the window back.</summary>
    public void SetChecked(bool value)
    {
        if (isChecked != value)
        {
            isChecked = value;
            Changed(nameof(IsChecked));
        }
    }

    public string Message
    {
        get => message;
        set => Set(ref message, value);
    }

    public string Title => node?.Title ?? message;

    /// <summary>What screen readers call the chevron and the tick box.</summary>
    public string OpenName => $"{(isExpanded ? "Close" : "Open")} {Title}";
    public string TickName => $"Tick {Title}";

    public string ToolTip => node is null ? message : $"{node.Title}\n{node.Url}";

    public string Glyph => Glyphs.For(node?.Kind);

    public string TypeText => node is null ? "" : Describe.Type(node);

    public string ItemsText => node?.ItemCount is { } n && node.Kind != NodeKind.File ? n.ToString("N0", CultureInfo.CurrentCulture) : "";

    public string SizeText => node?.Size is { } s ? Describe.Size(s) : "";

    public string ModifiedText => node?.Modified is { } m ? m.ToLocalTime().ToString("g", CultureInfo.CurrentCulture) : "";

    // ---- what the template's events call ----

    public void OnChevron() => owner.Toggle(this);

    public void OnDoubleTapped()
    {
        if (CanExpand)
        {
            owner.Toggle(this);
        }
    }

    public void OnMore() => owner.LoadMore(this);

    private bool Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }
        field = value;
        Changed(name!);
        return true;
    }

    private void Changed(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

/// <summary>The icon of each kind of thing (Segoe Fluent Icons).</summary>
public static class Glyphs
{
    public static string For(NodeKind? kind) => kind switch
    {
        NodeKind.Tenant => "",
        NodeKind.SiteCollection => "",
        NodeKind.Subsite => "",
        NodeKind.Library => "",
        NodeKind.List => "",
        NodeKind.Folder => "",
        NodeKind.File => "",
        NodeKind.Item => "",
        _ => "",
    };
}

/// <summary>Words for what SharePoint calls things by number or code.</summary>
public static class Describe
{
    public static string Type(SpNode n) => n.Kind switch
    {
        NodeKind.SiteCollection or NodeKind.Subsite => WebTemplate(n.Template),
        NodeKind.Library or NodeKind.List => ListTemplate(n.Template, n.Kind),
        NodeKind.Folder => "Folder",
        NodeKind.File => Path.GetExtension(n.LeafName) is { Length: > 1 } ext ? $"{ext[1..].ToUpperInvariant()} file" : "File",
        NodeKind.Item => "Item",
        _ => "",
    };

    public static string WebTemplate(string? template) => template?.ToUpperInvariant() switch
    {
        null or "" => "Site",
        "GROUP#0" => "Team site (group)",
        "SITEPAGEPUBLISHING#0" => "Communication site",
        "STS#3" => "Team site",
        "STS#0" => "Classic team site",
        "STS#1" => "Blank site",
        "TEAMCHANNEL#0" or "TEAMCHANNEL#1" => "Channel site",
        "APPCATALOG#0" => "App catalog",
        "SRCHCEN#0" => "Search centre",
        "BLANKINTERNET#0" or "BLANKINTERNETCONTAINER#0" => "Publishing site",
        "EHS#1" => "Team site (classic)",
        "POINTPUBLISHINGHUB#0" or "POINTPUBLISHINGTOPIC#0" => "Video portal",
        "REDIRECTSITE#0" => "Redirect",
        var t => t,
    };

    public static string ListTemplate(string? template, NodeKind kind) => template switch
    {
        "101" => "Document library",
        "109" => "Picture library",
        "119" => "Site pages",
        "850" => "Pages",
        "115" => "Form library",
        "700" => "My site documents",
        "336" or "330" => "App catalog",
        "100" => "List",
        "107" or "171" => "Tasks",
        "106" => "Calendar",
        "104" => "Announcements",
        "105" => "Contacts",
        "103" => "Links",
        "1100" => "Issues",
        "102" => "Survey",
        "108" => "Discussion board",
        "160" => "Access requests",
        "544" => "Persistent storage",
        "1230" => "Microsoft List",
        _ => kind == NodeKind.Library ? "Library" : "List",
    };

    public static string Size(long bytes)
    {
        if (bytes < 1024)
        {
            return $"{bytes:N0} bytes";
        }
        string[] units = ["KB", "MB", "GB", "TB"];
        double v = bytes / 1024.0;
        var i = 0;
        while (v >= 1024 && i < units.Length - 1)
        {
            v /= 1024;
            i++;
        }
        return v >= 100 ? $"{v:N0} {units[i]}" : $"{v:0.#} {units[i]}";
    }
}
