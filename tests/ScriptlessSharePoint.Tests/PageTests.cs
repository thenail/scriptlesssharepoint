using System.Net;
using System.Text.Json.Nodes;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class PageTests
{
    [Theory]
    [InlineData("Holiday policy", "Holiday-policy.aspx")]
    [InlineData("Holiday policy.aspx", "Holiday-policy.aspx")]
    [InlineData("  Q3: results / plans? ", "Q3-results-plans.aspx")]
    [InlineData("###", "Page.aspx")]
    public void Page_names_are_file_names(string typed, string name) =>
        Assert.Equal(name, CommandCatalog.PageFileName(typed));

    [Fact]
    public void Text_becomes_paragraphs_and_is_escaped()
    {
        Assert.Equal("<p>One<br>two</p><p>&lt;b&gt; &amp; three</p>", CommandCatalog.TextHtml("One\r\ntwo\r\n\r\n<b> & three\n"));
    }

    [Fact]
    public void Text_goes_on_an_empty_page()
    {
        var canvas = CommandCatalog.WithText(null, "Hello");
        Assert.StartsWith("<div><div data-sp-canvascontrol=", canvas);
        Assert.EndsWith("</div>", canvas);
        Assert.Contains("<div data-sp-rte=\"\"><p>Hello</p></div>", canvas);
        Assert.Equal(1, Zone(canvas, 0));
    }

    [Fact]
    public void Text_goes_after_the_last_section()
    {
        var page = CommandCatalog.WithText(CommandCatalog.WithText(null, "First"), "Second");
        Assert.Equal(2, Zones(page).Count);
        Assert.Equal(2, Zone(page, 1));
        Assert.True(page.IndexOf("First", StringComparison.Ordinal) < page.IndexOf("Second", StringComparison.Ordinal));
        Assert.StartsWith("<div><div", page);
        Assert.EndsWith("</div></div></div>", page);
    }

    [Fact]
    public void Text_goes_into_the_json_canvas_before_the_page_settings()
    {
        var canvas = """[{"controlType":4,"position":{"zoneIndex":1,"sectionIndex":1},"innerHTML":"<p>a</p>"},{"controlType":0,"pageSettingsSlice":{}}]""";
        var controls = JsonNode.Parse(CommandCatalog.WithText(canvas, "b"))!.AsArray();
        Assert.Equal(3, controls.Count);
        Assert.Equal("<p>b</p>", controls[1]!["innerHTML"]!.GetValue<string>());
        Assert.Equal(2, controls[1]!["position"]!["zoneIndex"]!.GetValue<double>());
        Assert.Equal(0, controls[2]!["controlType"]!.GetValue<int>());
    }

    [Fact]
    public void Page_commands_run_on_sites_and_files()
    {
        var pages = CommandCatalog.All.Where(c => c.Group == "Pages").ToList();
        Assert.Contains(pages, c => c.Id == "page.new" && c.AppliesTo(NodeKind.Subsite));
        Assert.All(pages, c => Assert.True(c.AppliesTo(NodeKind.File) || c.AppliesTo(NodeKind.SiteCollection), c.Id));
    }

    private static List<double> Zones(string canvas) =>
        System.Text.RegularExpressions.Regex.Matches(WebUtility.HtmlDecode(canvas), "\"zoneIndex\":([0-9.]+)")
            .Select(m => double.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture))
            .ToList();

    private static double Zone(string canvas, int i) => Zones(canvas)[i];
}
