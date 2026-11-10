using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

namespace ScriptlessSharePoint.App;

/// <summary>
/// The gap between the table and the pane, as a handle: the cursor shows it
/// can be dragged, a grip shows on hover, and dragging reports how far it's
/// moved (the window sizes the pane). A double-click asks for the usual width.
/// </summary>
public sealed partial class Splitter : Grid
{
    private readonly Border grip;
    private double startX;
    private bool dragging;

    public Splitter()
    {
        ProtectedCursor = InputSystemCursor.Create(InputSystemCursorShape.SizeWestEast);
        Background = new SolidColorBrush(Microsoft.UI.Colors.Transparent);
        grip = new Border
        {
            Width = 4,
            Height = 40,
            CornerRadius = new CornerRadius(2),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Background = (Brush)Application.Current.Resources["ControlStrongFillColorDefaultBrush"],
            Opacity = 0,
        };
        Children.Add(grip);
        PointerEntered += (_, _) => grip.Opacity = 1;
        PointerExited += (_, _) => grip.Opacity = dragging ? 1 : 0;
        PointerPressed += (_, e) =>
        {
            dragging = CapturePointer(e.Pointer);
            startX = e.GetCurrentPoint(null).Position.X;
            Started?.Invoke();
            e.Handled = true;
        };
        PointerMoved += (_, e) =>
        {
            if (dragging)
            {
                Dragged?.Invoke(e.GetCurrentPoint(null).Position.X - startX);
                e.Handled = true;
            }
        };
        PointerReleased += (_, e) =>
        {
            if (dragging)
            {
                dragging = false;
                ReleasePointerCapture(e.Pointer);
                Released?.Invoke();
            }
        };
        PointerCaptureLost += (_, _) =>
        {
            if (dragging)
            {
                dragging = false;
                Released?.Invoke();
            }
            grip.Opacity = 0;
        };
        DoubleTapped += (_, _) => Reset?.Invoke();
    }

    public event Action? Started;

    /// <summary>How far it's been dragged since it was pressed (right is positive).</summary>
    public event Action<double>? Dragged;

    public event Action? Released;

    public event Action? Reset;
}
