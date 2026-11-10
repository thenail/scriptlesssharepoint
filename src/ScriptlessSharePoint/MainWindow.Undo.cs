using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

using ScriptlessSharePoint.Selections;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>
/// Undo for the ticks: whatever replaces what's ticked wholesale (a saved or
/// predefined selection, a finder's finds, Clear) keeps it first, and Undo
/// (in the toast, beside Clear, or Ctrl+Z) ticks it again, shown as a saved
/// selection shows its items. When what's replaced isn't kept anywhere else
/// (ticked by hand, or a saved selection changed since), a toast says so.
/// </summary>
public sealed partial class MainWindow
{
    /// <summary>How many replaced selections Undo goes back through.</summary>
    private const int UndoDepth = 20;

    /// <summary>What was ticked before, the newest last, with the predefined selection it came from.</summary>
    private readonly List<(SavedSelection Ticked, string? Preset)> ticksBefore = [];

    /// <summary>While Undo puts ticks back: that doesn't count as replacing them.</summary>
    private bool restoringTicks;

    /// <summary>
    /// Keeps what's ticked before <paramref name="what"/> replaces it. With
    /// <paramref name="always"/>, or when it's kept nowhere else, a toast
    /// offers Undo.
    /// </summary>
    private void KeepTicksBefore(string what, bool always)
    {
        if (restoringTicks || selection.Count == 0 || active is null)
        {
            return;
        }
        var items = selection.Values.ToList();
        var kept = ShownUnchanged();
        var ticked = (shownSaved ?? new SavedSelection { Name = "What was ticked", ConnectionId = active.Id, Saved = DateTimeOffset.Now }) with
        {
            ConnectionId = active.Id,
            Items = items,
            Sites = items.Select(n => siteOf.TryGetValue(n.Key, out var site) ? site : SpNode.SiteFromUrl(n.SiteUrl ?? n.Url)).DistinctBy(n => n.Key).ToList(),
            Parents = items.Select(n => parentOf.GetValueOrDefault(n.Key)).OfType<SpNode>().DistinctBy(n => n.Key).ToList(),
        };
        ticksBefore.RemoveAll(t => t.Ticked.ConnectionId != active.Id);
        ticksBefore.Add((ticked, shownPreset));
        if (ticksBefore.Count > UndoDepth)
        {
            ticksBefore.RemoveAt(0);
        }
        UpdateUndo();
        if (always || !kept)
        {
            ShowToast(what, $"In place of {Summary(items)}.", ("Undo", UndoTicks), success: true);
        }
    }

    /// <summary>Ticks again what was ticked before the last replacement, and keeps a plan on show.</summary>
    private void UndoTicks()
    {
        if (active is null || ticksBefore.Count == 0)
        {
            return;
        }
        var (ticked, preset) = ticksBefore[^1];
        ticksBefore.RemoveAt(ticksBefore.Count - 1);
        UpdateUndo();
        if (ticked.ConnectionId != active.Id || identity is null)
        {
            return;
        }
        var plan = PlanView.Visibility == Visibility.Visible ? shownPlan : null;
        restoringTicks = true;
        try
        {
            ShowSaved(ticked);
        }
        finally
        {
            restoringTicks = false;
        }
        shownRule = null;
        if (preset is not null)
        {
            shownPreset = preset;
            SelectNavQuietly(PresetTag + preset);
        }
        if (plan is not null)
        {
            shownPlan = plan;
            ShowPlan(plan);
        }
    }

    private void OnUndoTicks(object sender, RoutedEventArgs e) => UndoTicks();

    /// <summary>Whether what's ticked is what the saved (or predefined) selection on show holds, unchanged.</summary>
    private bool ShownUnchanged() =>
        shownSaved is { } s && s.Items.Count == selection.Count && s.Items.All(n => selection.ContainsKey(n.Key));

    /// <summary>The Undo button beside Clear: there when there's something to go back to.</summary>
    private void UpdateUndo()
    {
        var last = active is null ? null : ticksBefore.LastOrDefault(t => t.Ticked.ConnectionId == active.Id).Ticked;
        UndoButton.Visibility = last is null ? Visibility.Collapsed : Visibility.Visible;
        ToolTipService.SetToolTip(UndoButton, last is null ? null : $"Tick again what was ticked before: {Summary(last.Items)} (Ctrl+Z)");
    }

    /// <summary>Ctrl+Z: Undo for the ticks, unless a text box has the focus (then it's the text's).</summary>
    private bool TypingInText() =>
        Microsoft.UI.Xaml.Input.FocusManager.GetFocusedElement(Root.XamlRoot) is TextBox or PasswordBox or RichEditBox;
}
