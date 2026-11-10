using System.ComponentModel;

using Microsoft.UI.Xaml;

namespace ScriptlessSharePoint.App;

/// <summary>
/// The widths of the table's columns after Name, shared by the headings and
/// every row. On a narrow table the last columns step aside, as in File Size,
/// so Name keeps room to read.
/// </summary>
public sealed class Columns : INotifyPropertyChanged
{
    /// <summary>What Name keeps before a column steps aside (it also has the indent, chevron and tick).</summary>
    private const double NameRoom = 300;

    private static readonly (string Name, double Width)[] Order =
    [
        (nameof(Type), 150),
        (nameof(Items), 80),
        (nameof(Size), 90),
        (nameof(Modified), 150),
    ];

    private readonly Dictionary<string, bool> shown = Order.ToDictionary(c => c.Name, _ => true);

    public static Columns Current { get; } = new();

    public event PropertyChangedEventHandler? PropertyChanged;

    public GridLength Type => Width(nameof(Type));
    public GridLength Items => Width(nameof(Items));
    public GridLength Size => Width(nameof(Size));
    public GridLength Modified => Width(nameof(Modified));

    /// <summary>Shows as many columns, in order, as leave Name its room.</summary>
    public void Fit(double tableWidth)
    {
        var left = tableWidth - NameRoom;
        foreach (var (name, width) in Order)
        {
            var show = left >= width;
            if (show)
            {
                left -= width;
            }
            if (shown[name] != show)
            {
                shown[name] = show;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
            }
        }
    }

    private GridLength Width(string name) =>
        new(shown[name] ? Order.First(c => c.Name == name).Width : 0);
}
