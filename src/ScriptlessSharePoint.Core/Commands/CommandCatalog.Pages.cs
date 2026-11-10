using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

using Microsoft.SharePoint.Client;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>
/// Modern pages: those in a site's Site Pages library. They're list items whose
/// fields hold the page (its canvas in CanvasContent1, its header in
/// LayoutWebpartsContent), so they're made and changed through CSOM as PnP's
/// own library does: a page file from SharePoint's template, then its fields.
/// Page commands run on files and skip those that aren't modern pages.
/// </summary>
public static partial class CommandCatalog
{
    /// <summary>The Site Page content type.</summary>
    private const string SitePageContentType = "0x0101009D1CB255DA76424F860D91F20E6C4118";

    /// <summary>What marks a list item as a modern page: the page app's ID.</summary>
    private const string PageApplicationId = "b6917cb1-93a0-4b97-a84d-7cf49975d4ec";

    private const int SitePagesTemplate = 119;

    private static IEnumerable<Command> PageCommands()
    {
        yield return new Command
        {
            Id = "page.list",
            Title = "List pages",
            Group = "Pages",
            Description = "Each site's pages: published or draft, news, and which is the home page.",
            PowerShell = "Get-PnPPage",
            Targets = Webs,
            ReadOnly = true,
            Glyph = "",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                if (await SitePagesAsync(ctx).ConfigureAwait(false) is not { } library)
                {
                    return CommandResult.Skipped("It has no Site Pages library.");
                }
                var items = library.GetItems(new CamlQuery
                {
                    ViewXml = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where>" +
                              "<OrderBy><FieldRef Name='Modified' Ascending='FALSE'/></OrderBy></Query><RowLimit>2000</RowLimit></View>",
                });
                ctx.Load(items, xs => xs.Include(i => i["Title"], i => i["FileLeafRef"], i => i["FileRef"], i => i["PromotedState"], i => i["_UIVersionString"], i => i["Modified"]));
                ctx.Load(ctx.Web.RootFolder, f => f.WelcomePage);
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                await ctx.RunAsync().ConfigureAwait(false);
                var home = $"{ctx.Web.ServerRelativeUrl.TrimEnd('/')}/{ctx.Web.RootFolder.WelcomePage}";
                var rows = items.ToList().Select(i =>
                {
                    var file = Convert.ToString(i["FileLeafRef"], CultureInfo.InvariantCulture) ?? "";
                    var title = Convert.ToString(i["Title"], CultureInfo.InvariantCulture);
                    var version = Convert.ToString(i["_UIVersionString"], CultureInfo.InvariantCulture) ?? "";
                    var news = i["PromotedState"] is { } p && Convert.ToDouble(p, CultureInfo.InvariantCulture) >= 2;
                    var modified = i["Modified"] is DateTime m ? m.ToLocalTime().ToString("g", CultureInfo.CurrentCulture) : "";
                    return new[]
                    {
                        string.IsNullOrWhiteSpace(title) ? file : title,
                        file,
                        version.EndsWith(".0", StringComparison.Ordinal) ? "published" : "draft",
                        news ? "news" : "",
                        string.Equals(Convert.ToString(i["FileRef"], CultureInfo.InvariantCulture), home, StringComparison.OrdinalIgnoreCase) ? "home page" : "",
                        modified,
                    };
                }).ToList();
                return CommandResult.Report($"{Plural(rows.Count, "page", "pages")}.", ReportTable.Of(["Page", "File", "State", "News", "Home page", "Modified"], rows));
            },
        };

        yield return new Command
        {
            Id = "page.new",
            Title = "New page",
            Group = "Pages",
            Description = "Creates a modern page in each site's Site Pages library, with some text if given, as a page or as news.",
            PowerShell = "Add-PnPPage",
            Targets = Webs,
            Glyph = "",
            Parameters =
            [
                new("name", "Name", ParameterKind.Text) { Placeholder = "Holiday policy", Hint = "Its file name: spaces become dashes, .aspx is added." },
                new("title", "Title", ParameterKind.Text) { Required = false, Hint = "Empty: the name." },
                new("layout", "Layout", ParameterKind.Choice) { Default = "Article", Choices = [new("Article", "Article (with a title area)"), new("Home", "Home (no title area)")] },
                new("text", "Text", ParameterKind.LongText) { Required = false, Hint = "Optional. A blank line starts a new paragraph." },
                Toggle("news", "Post it as news", false),
                Toggle("publish", "Publish it", true, "Off: it stays a draft only editors see."),
            ],
            Run = async (c, n) =>
            {
                var name = PageFileName(c.Args.Text("name"));
                var title = c.Args["title"]?.Trim() ?? Path.GetFileNameWithoutExtension(name).Replace('-', ' ');
                var news = c.Args.Toggle("news");
                var canvas = c.Args["text"] is { } text ? WithText(null, text) : "<div></div>";
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                return await CreatePageAsync(ctx, name, overwrite: false, c.Args.Toggle("publish"), item =>
                {
                    item["Title"] = title;
                    item["PageLayoutType"] = news ? "Article" : c.Args.Text("layout");
                    item["CanvasContent1"] = canvas;
                    item["PromotedState"] = news ? 2 : 0;
                    if (news)
                    {
                        item["FirstPublishedDate"] = DateTime.UtcNow;
                    }
                }).ConfigureAwait(false);
            },
        };

        yield return new Command
        {
            Id = "page.copy-in",
            Title = "Copy a page into each site",
            Group = "Pages",
            Description = "Copies a page, with its text, web parts and header, into each site's Site Pages library: one page made once, put in many sites. Images and links still point where they did.",
            PowerShell = "Export-PnPPage, Invoke-PnPSiteTemplate",
            Targets = Webs,
            Glyph = "",
            Parameters =
            [
                new("source", "Page to copy", ParameterKind.Url) { Placeholder = "https://contoso.sharepoint.com/sites/templates/SitePages/Welcome.aspx", Hint = "The page's address, as the browser shows it." },
                new("name", "Name", ParameterKind.Text) { Required = false, Hint = "Empty: the page's own name." },
                Toggle("overwrite", "Replace a page of the same name", false),
                Toggle("publish", "Publish it", true),
            ],
            Run = async (c, n) =>
            {
                var source = c.Args.Text("source").Trim();
                var at = source.IndexOf("/SitePages/", StringComparison.OrdinalIgnoreCase);
                if (at < 0 || !Uri.TryCreate(source, UriKind.Absolute, out var uri))
                {
                    return CommandResult.Failed("That isn't a page's address: it should have /SitePages/ in it.");
                }
                Dictionary<string, object?> fields;
                using (var from = await c.Client.ContextAsync(source[..at], c.Cancellation).ConfigureAwait(false))
                {
                    var item = from.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(Uri.UnescapeDataString(uri.AbsolutePath))).ListItemAllFields;
                    from.Load(item);
                    await from.RunAsync().ConfigureAwait(false);
                    if (!IsPage(item))
                    {
                        return CommandResult.Failed("That isn't a modern page.");
                    }
                    fields = new[] { "Title", "PageLayoutType", "CanvasContent1", "LayoutWebpartsContent", "BannerImageUrl" }
                        .Where(item.FieldValues.ContainsKey)
                        .ToDictionary(f => f, f => (object?)item.FieldValues[f]);
                }
                var name = c.Args["name"] is { } typed ? PageFileName(typed) : Uri.UnescapeDataString(uri.Segments[^1]);
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                return await CreatePageAsync(ctx, name, c.Args.Toggle("overwrite"), c.Args.Toggle("publish"), item =>
                {
                    foreach (var (field, value) in fields)
                    {
                        item[field] = value;
                    }
                }).ConfigureAwait(false);
            },
        };

        yield return new Command
        {
            Id = "page.title",
            Title = "Page title",
            Group = "Pages",
            Description = "Changes each page's title (its address stays).",
            PowerShell = "Set-PnPPage -Title",
            Targets = Files,
            After = After.Update,
            Glyph = "",
            Parameters = [new("title", "Title", ParameterKind.Text), Toggle("publish", "Publish the change", true)],
            Run = (c, n) => ChangePageAsync(c, n, item => item["Title"] = c.Args.Text("title").Trim(), "Retitled."),
        };

        yield return new Command
        {
            Id = "page.news",
            Title = "News or page",
            Group = "Pages",
            Description = "Posts each page as news (it shows in news web parts and the SharePoint app), or makes it a plain page again.",
            PowerShell = "Set-PnPPage -PromoteAs",
            Targets = Files,
            Glyph = "",
            Parameters =
            [
                new("as", "Make it", ParameterKind.Choice) { Default = "news", Choices = [new("news", "News"), new("page", "A plain page")] },
                Toggle("publish", "Publish the change", true),
            ],
            Run = (c, n) => ChangePageAsync(c, n, item =>
            {
                var news = c.Args.Text("as") == "news";
                item["PromotedState"] = news ? 2 : 0;
                if (news && item.FieldValues.GetValueOrDefault("FirstPublishedDate") is null)
                {
                    item["FirstPublishedDate"] = DateTime.UtcNow;
                }
            }, c.Args.Text("as") == "news" ? "Posted as news." : "A plain page now."),
        };

        yield return new Command
        {
            Id = "page.add-text",
            Title = "Add text to a page",
            Group = "Pages",
            Description = "Adds a full-width section with text at the end of each page.",
            PowerShell = "Add-PnPPageTextPart",
            Targets = Files,
            Glyph = "",
            Parameters =
            [
                new("text", "Text", ParameterKind.LongText) { Hint = "A blank line starts a new paragraph." },
                Toggle("publish", "Publish the change", true),
            ],
            Run = (c, n) => ChangePageAsync(c, n, item =>
                item["CanvasContent1"] = WithText(Convert.ToString(item.FieldValues.GetValueOrDefault("CanvasContent1"), CultureInfo.InvariantCulture), c.Args.Text("text")),
                "Added the text."),
        };

        yield return new Command
        {
            Id = "page.home",
            Title = "Make it the home page",
            Group = "Pages",
            Description = "Makes the page the one its site opens on.",
            PowerShell = "Set-PnPHomePage",
            Targets = Files,
            Glyph = "",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var item = Csom.Item(ctx, n);
                ctx.Load(item);
                await ctx.RunAsync().ConfigureAwait(false);
                if (!IsPage(item))
                {
                    return CommandResult.Skipped("It isn't a modern page.");
                }
                ctx.Web.RootFolder.WelcomePage = WithinWeb(n.WebUrl, n.ServerRelativeUrl ?? "");
                ctx.Web.RootFolder.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Its site opens on it now.");
            },
        };

        yield return new Command
        {
            Id = "page.comments",
            Title = "Comments on the page",
            Group = "Pages",
            Description = "Whether people can comment on each page. Where the site turns comments off, they stay off.",
            PowerShell = "Set-PnPPage -CommentsEnabled",
            Targets = Files,
            Glyph = "",
            Parameters = [Toggle("allow", "Allow comments", true)],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var item = Csom.Item(ctx, n);
                ctx.Load(item);
                await ctx.RunAsync().ConfigureAwait(false);
                if (!IsPage(item))
                {
                    return CommandResult.Skipped("It isn't a modern page.");
                }
                item.SetCommentsDisabled(!c.Args.Toggle("allow"));
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done(c.Args.Toggle("allow") ? "Comments allowed." : "Comments turned off.");
            },
        };
    }

    // ---- making and changing pages ----

    /// <summary>The site's Site Pages library, with its folder and versioning; null when it has none.</summary>
    private static async Task<List?> SitePagesAsync(ClientContext ctx)
    {
        var lists = ctx.LoadQuery(ctx.Web.Lists.Where(l => l.BaseTemplate == SitePagesTemplate)
            .Include(l => l.Id, l => l.RootFolder.ServerRelativeUrl, l => l.EnableMinorVersions, l => l.EnableModeration));
        await ctx.RunAsync().ConfigureAwait(false);
        return lists.FirstOrDefault();
    }

    /// <summary>Whether a list item is a modern page.</summary>
    private static bool IsPage(ListItem item) =>
        (Convert.ToString(item.FieldValues.GetValueOrDefault("ContentTypeId"), CultureInfo.InvariantCulture) ?? "").StartsWith(SitePageContentType, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(Convert.ToString(item.FieldValues.GetValueOrDefault("ClientSideApplicationId"), CultureInfo.InvariantCulture), PageApplicationId, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// A page in the site's Site Pages library: made from SharePoint's page
    /// template (or, with <paramref name="overwrite"/>, the one there already),
    /// its fields set, then published when asked.
    /// </summary>
    private static async Task<CommandResult> CreatePageAsync(ClientContext ctx, string name, bool overwrite, bool publish, Action<ListItem> fill)
    {
        if (await SitePagesAsync(ctx).ConfigureAwait(false) is not { } library)
        {
            return CommandResult.Skipped("It has no Site Pages library.");
        }
        var url = $"{library.RootFolder.ServerRelativeUrl.TrimEnd('/')}/{name}";
        var existing = ctx.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(url));
        ctx.Load(existing, f => f.Exists);
        await ctx.RunAsync().ConfigureAwait(false);
        if (existing.Exists && !overwrite)
        {
            return CommandResult.Skipped($"It has a page called {name} already.");
        }
        var file = existing.Exists ? existing : library.RootFolder.Files.AddTemplateFile(url, TemplateFileType.ClientSidePage);
        var item = file.ListItemAllFields;
        item["ContentTypeId"] = SitePageContentType;
        item["ClientSideApplicationId"] = PageApplicationId;
        fill(item);
        item.Update();
        await ctx.RunAsync().ConfigureAwait(false);
        await PublishPageAsync(ctx, file, library, publish).ConfigureAwait(false);
        return CommandResult.Done($"{(existing.Exists ? "Replaced" : "Created")} {name}{(publish ? ", published" : ", as a draft")}.");
    }

    /// <summary>Changes a page's fields, then publishes it when the "publish" input says so.</summary>
    private static async Task<CommandResult> ChangePageAsync(CommandContext c, SpNode n, Action<ListItem> change, string done)
    {
        using var ctx = await c.WebAsync(n).ConfigureAwait(false);
        var item = Csom.Item(ctx, n);
        var list = item.ParentList;
        ctx.Load(item);
        ctx.Load(list, l => l.EnableMinorVersions, l => l.EnableModeration);
        await ctx.RunAsync().ConfigureAwait(false);
        if (!IsPage(item))
        {
            return CommandResult.Skipped("It isn't a modern page.");
        }
        change(item);
        item.Update();
        await ctx.RunAsync().ConfigureAwait(false);
        var publish = c.Args["publish"] is null || c.Args.Toggle("publish");
        await PublishPageAsync(ctx, item.File, list, publish).ConfigureAwait(false);
        return CommandResult.Done(publish ? done : $"{done} (a draft)");
    }

    /// <summary>Checks a page in when it's out, and publishes it as a major version where the library keeps drafts.</summary>
    private static async Task PublishPageAsync(ClientContext ctx, Microsoft.SharePoint.Client.File file, List library, bool publish)
    {
        ctx.Load(file, f => f.Level, f => f.CheckOutType);
        await ctx.RunAsync().ConfigureAwait(false);
        if (file.CheckOutType != CheckOutType.None)
        {
            file.CheckIn("", publish && !library.EnableMinorVersions ? CheckinType.MajorCheckIn : CheckinType.MinorCheckIn);
            await ctx.RunAsync().ConfigureAwait(false);
        }
        if (publish && library.EnableMinorVersions)
        {
            file.Publish("");
            await ctx.RunAsync().ConfigureAwait(false);
        }
    }

    // ---- names and canvas ----

    /// <summary>"Holiday policy" → "Holiday-policy.aspx", as SharePoint names new pages.</summary>
    internal static string PageFileName(string typed)
    {
        var name = typed.Trim();
        if (name.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
        {
            name = name[..^5];
        }
        name = Regex.Replace(name, "[\"*:<>?/\\\\|#%~&{}]", "");
        name = Regex.Replace(name, @"\s+", "-").Trim('-', '.');
        return (name.Length == 0 ? "Page" : name) + ".aspx";
    }

    /// <summary>Text as a text web part's HTML: paragraphs at blank lines, line breaks within them.</summary>
    internal static string TextHtml(string text) => string.Concat(
        Regex.Split(text.Replace("\r\n", "\n").Trim(), @"\n\s*\n")
            .Where(p => p.Trim().Length > 0)
            .Select(p => $"<p>{string.Join("<br>", p.Trim().Split('\n').Select(l => WebUtility.HtmlEncode(l.TrimEnd())))}</p>"));

    /// <summary>
    /// A page's canvas with a full-width text section added at the end. The
    /// canvas is the HTML SharePoint keeps in CanvasContent1 (or its JSON
    /// form); the new section goes after the last one.
    /// </summary>
    internal static string WithText(string? canvas, string text)
    {
        var id = Guid.NewGuid().ToString();
        var html = TextHtml(text);
        var trimmed = canvas?.Trim() ?? "";
        if (trimmed.StartsWith('['))
        {
            var controls = JsonNode.Parse(trimmed)!.AsArray();
            var last = controls.Select(x => x?["position"]?["zoneIndex"]?.GetValue<double>() ?? 0).DefaultIfEmpty(0).Max();
            var control = TextControlData(id, Math.Floor(last) + 1);
            control["innerHTML"] = html;
            // the page settings (controlType 0) stay last
            var at = controls.Count > 0 && controls[^1]?["controlType"]?.GetValue<int>() == 0 ? controls.Count - 1 : controls.Count;
            controls.Insert(at, control);
            return controls.ToJsonString();
        }
        var zones = Regex.Matches(WebUtility.HtmlDecode(trimmed), "\"zoneIndex\"\\s*:\\s*([0-9.]+)")
            .Select(m => double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture))
            .DefaultIfEmpty(0)
            .Max();
        var part = $"<div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"{WebUtility.HtmlEncode(TextControlData(id, Math.Floor(zones) + 1).ToJsonString())}\">" +
                   $"<div data-sp-rte=\"\">{html}</div></div>";
        if (trimmed.StartsWith("<div>", StringComparison.OrdinalIgnoreCase) && trimmed.EndsWith("</div>", StringComparison.OrdinalIgnoreCase))
        {
            return trimmed[..^6] + part + "</div>";
        }
        return $"<div>{trimmed}{part}</div>";
    }

    private static JsonObject TextControlData(string id, double zone) => new()
    {
        ["controlType"] = 4,
        ["displayMode"] = 2,
        ["id"] = id,
        ["position"] = new JsonObject
        {
            ["zoneIndex"] = zone,
            ["sectionIndex"] = 1,
            ["controlIndex"] = 1,
            ["sectionFactor"] = 12,
            ["layoutIndex"] = 1,
        },
        ["emphasis"] = new JsonObject(),
        ["zoneGroupMetadata"] = new JsonObject { ["type"] = 0 },
        ["editorType"] = "CKEditor",
    };
}
