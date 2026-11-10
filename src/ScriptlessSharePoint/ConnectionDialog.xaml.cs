using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

using ScriptlessSharePoint.Auth;
using ScriptlessSharePoint.Connections;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>Adds or edits a connection. Test signs in with it as it stands; Save keeps it.</summary>
public sealed partial class ConnectionDialog : ContentDialog
{
    private enum Mode
    {
        Windows,
        User,
        Certificate,
        Token,
    }

    private readonly Workspace workspace;
    private readonly ConnectionProfile? existing;
    private readonly MainWindow host;
    private string? id;
    private string? pfxPath;

    public ConnectionDialog(Workspace workspace, ConnectionProfile? existing, MainWindow host)
    {
        InitializeComponent();
        this.workspace = workspace;
        this.existing = existing;
        this.host = host;
        id = existing?.Id;

        Title = existing is null ? "New connection" : $"Edit {existing.Name}";
        var d = existing is null ? new ConnectionDraft() : ConnectionDraft.From(existing);
        TenantBox.Text = d.TenantUrl;
        NameBox.Text = d.Name;
        AdminUrlBox.Text = d.AdminUrl ?? "";
        AdvancedExpander.IsExpanded = !string.IsNullOrEmpty(d.AdminUrl);
        var mode = d.Kind switch
        {
            ConnectionKind.Certificate => Mode.Certificate,
            ConnectionKind.AccessToken => Mode.Token,
            _ => d.UseWindowsAccount ? Mode.Windows : Mode.User,
        };
        if (d.Kind == ConnectionKind.User)
        {
            UserClientIdBox.Text = d.ClientId ?? "";
            UserTenantBox.Text = d.TenantId ?? "";
            BrowserCheck.IsChecked = d.UseBrowser;
        }
        else if (d.Kind == ConnectionKind.Certificate)
        {
            CertClientIdBox.Text = d.ClientId ?? "";
            CertTenantBox.Text = d.TenantId ?? "";
            ThumbprintBox.Text = d.CertificateThumbprint ?? "";
        }
        if (d.HasToken)
        {
            TokenBox.PlaceholderText = "A token is saved. Paste a new one to replace it.";
        }
        AccountText.Text = d.AccountName is { } a ? $"Signed in as {a}." : "";
        var builtIn = AppInfo.BuiltInClientId;
        UserClientIdBox.PlaceholderText = builtIn is null ? "00000000-0000-0000-0000-000000000000" : "Empty: Scriptless SharePoint's own app";
        UserClientIdBox.Description = builtIn is null
            ? "The ID of your organisation's Entra app registration for this app."
            : $"Leave it empty to use Scriptless SharePoint's own app ({builtIn}).";

        ModeButtons.SelectedIndex = (int)mode;
        ShowMode();

        PrimaryButtonClick += OnSave;
        SecondaryButtonClick += OnTest;
        Opened += (_, _) => (existing is null ? TenantBox : NameBox).Focus(FocusState.Programmatic);
    }

    /// <summary>The connection as saved, when Save worked.</summary>
    public ConnectionProfile? Saved { get; private set; }

    private Mode CurrentMode => (Mode)Math.Max(0, ModeButtons.SelectedIndex);

    private void OnModeChanged(object sender, SelectionChangedEventArgs e) => ShowMode();

    private void ShowMode()
    {
        var mode = CurrentMode;
        UserPanel.Visibility = mode is Mode.Windows or Mode.User ? Visibility.Visible : Visibility.Collapsed;
        BrowserCheck.Visibility = mode == Mode.User ? Visibility.Visible : Visibility.Collapsed;
        CertPanel.Visibility = mode == Mode.Certificate ? Visibility.Visible : Visibility.Collapsed;
        TokenPanel.Visibility = mode == Mode.Token ? Visibility.Visible : Visibility.Collapsed;
    }

    private ConnectionDraft Draft()
    {
        var mode = CurrentMode;
        var user = mode is Mode.Windows or Mode.User;
        return new ConnectionDraft
        {
            Id = id,
            Name = NameBox.Text.Trim(),
            Kind = user ? ConnectionKind.User : mode == Mode.Certificate ? ConnectionKind.Certificate : ConnectionKind.AccessToken,
            TenantUrl = TenantBox.Text.Trim(),
            AdminUrl = Blank(AdminUrlBox.Text),
            ClientId = Blank(user ? UserClientIdBox.Text : mode == Mode.Certificate ? CertClientIdBox.Text : ""),
            TenantId = Blank(user ? UserTenantBox.Text : mode == Mode.Certificate ? CertTenantBox.Text : ""),
            UseWindowsAccount = mode == Mode.Windows,
            UseBrowser = mode == Mode.User && BrowserCheck.IsChecked == true,
            CertificateThumbprint = mode == Mode.Certificate ? Blank(ThumbprintBox.Text) : null,
        };
    }

    private string? Token() => CurrentMode == Mode.Token ? Blank(TokenBox.Text) : null;

    private static string? Blank(string? s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

    private void OnSave(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        try
        {
            Saved = workspace.Save(Draft(), Token());
        }
        catch (Exception e)
        {
            args.Cancel = true;
            Show(false, e is ConnectionException ? e.Message : SharePointErrors.Describe(e));
        }
    }

    private async void OnTest(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        args.Cancel = true;
        var deferral = args.GetDeferral();
        IsSecondaryButtonEnabled = false;
        IsPrimaryButtonEnabled = false;
        Show(null, "Signing in…");
        try
        {
            var r = await workspace.TestAsync(Draft(), Token());
            id = r.Id;
            AccountText.Text = r.AccountName is { } a ? $"Signed in as {a}." : "";
            var who = r.Identity.IsApp
                ? $"an app ({r.Identity.DisplayName})"
                : $"{r.Identity.DisplayName} ({r.Identity.Email ?? r.Identity.LoginName})";
            Show(true, $"Connected to {r.Identity.TenantTitle} as {who}.");
        }
        catch (Exception e)
        {
            Show(false, e is ConnectionException ? e.Message : SharePointErrors.Describe(e));
        }
        finally
        {
            IsSecondaryButtonEnabled = true;
            IsPrimaryButtonEnabled = true;
            deferral.Complete();
        }
    }

    private void Show(bool? ok, string text)
    {
        ResultBar.Severity = ok switch { true => InfoBarSeverity.Success, false => InfoBarSeverity.Error, null => InfoBarSeverity.Informational };
        ResultBar.Message = text;
        ResultBar.IsOpen = true;
    }

    /// <summary>The tenant as typed → https://contoso.sharepoint.com, and a name from it.</summary>
    private void OnTenantLostFocus(object sender, RoutedEventArgs e)
    {
        if (!TenantUrls.TryNormalizeTenantUrl(TenantBox.Text, out var url))
        {
            return;
        }
        TenantBox.Text = url;
        if (string.IsNullOrWhiteSpace(NameBox.Text))
        {
            var label = new Uri(url).Host.Split('.')[0];
            NameBox.Text = char.ToUpperInvariant(label[0]) + label[1..];
        }
    }

    private void OnThumbprintChanged(object sender, TextChangedEventArgs e)
    {
        var tp = ThumbprintBox.Text.Trim();
        if (tp.Length == 0)
        {
            ThumbprintInfo.Visibility = Visibility.Collapsed;
            return;
        }
        bool ok;
        try
        {
            ok = Certificates.Find(tp) is not null;
        }
        catch (Exception)
        {
            ok = false;
        }
        ThumbprintInfo.Text = (ok ? "✓ " : "⚠ ") + Certificates.Describe(tp);
        ThumbprintInfo.Visibility = Visibility.Visible;
    }

    private void OnTokenChanged(object sender, TextChangedEventArgs e)
    {
        var text = TokenBox.Text;
        if (string.IsNullOrWhiteSpace(text))
        {
            TokenInfo.Text = "";
            return;
        }
        var info = Auth.TokenInfo.TryParse(ConnectionDraft.CleanToken(text));
        if (info is null)
        {
            TokenInfo.Text = "⚠ That doesn't look like an access token (a JWT: three parts separated by dots).";
            return;
        }
        var expired = info.Expires < DateTimeOffset.UtcNow;
        TokenInfo.Text = string.Join("\n", new[]
        {
            $"For: {info.Audience ?? "?"}{(info.IsForSharePoint ? "" : "  ⚠ not a SharePoint token")}",
            info.Subject is { } s ? $"As: {s}" : null,
            info.Expires is { } x ? $"{(expired ? "⚠ Expired" : "Expires")} {x.ToLocalTime():g}" : null,
            info.Permissions is { } p ? $"Permissions: {p}" : null,
        }.OfType<string>());
        if (info.IsForSharePoint && info.Audience?.Contains(".sharepoint.", StringComparison.OrdinalIgnoreCase) == true && string.IsNullOrWhiteSpace(TenantBox.Text))
        {
            TenantBox.Text = info.Audience;
            OnTenantLostFocus(TenantBox, new RoutedEventArgs());
        }
    }

    private async void OnImportPfx(object sender, RoutedEventArgs e)
    {
        pfxPath = await host.PickFileAsync(".pfx", ".p12");
        if (pfxPath is not null)
        {
            PfxPassword.Password = "";
            PfxPanel.Visibility = Visibility.Visible;
            PfxPassword.Focus(FocusState.Programmatic);
        }
    }

    private void OnFinishImport(object sender, RoutedEventArgs e)
    {
        if (pfxPath is null)
        {
            return;
        }
        try
        {
            ThumbprintBox.Text = Certificates.ImportPfx(pfxPath, Blank(PfxPassword.Password));
            Show(true, $"Imported {Path.GetFileName(pfxPath)} into your certificate store.");
            PfxPanel.Visibility = Visibility.Collapsed;
            pfxPath = null;
        }
        catch (Exception ex)
        {
            Show(false, ex.Message);
        }
    }

    private async void OnCreateCertificate(object sender, RoutedEventArgs e)
    {
        var name = string.IsNullOrWhiteSpace(NameBox.Text) ? AppInfo.Name : $"{AppInfo.Name} {NameBox.Text.Trim()}";
        var path = await host.PickSaveFileAsync(name.Replace(' ', '-'), "Certificate", ".cer");
        if (path is null)
        {
            return;
        }
        try
        {
            ThumbprintBox.Text = Certificates.CreateSelfSigned(name, path);
            Show(true, $"Made a certificate valid for two years and saved its public key to {path}. Upload that file to the app registration (Certificates & secrets › Certificates). The private key stays in your certificate store.");
        }
        catch (Exception ex)
        {
            Show(false, ex.Message);
        }
    }
}
