using System.Text.Json.Nodes;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class SharingTests
{
    [Theory]
    [InlineData("""{"sharingLinkInfo":{"Url":"https://contoso.sharepoint.com/:w:/s/hr/EaBc","LinkKind":2}}""")]
    [InlineData("""{"d":{"ShareLink":{"sharingLinkInfo":{"Url":"https://contoso.sharepoint.com/:w:/s/hr/EaBc"}}}}""")]
    public void Reads_the_new_links_address(string answer) =>
        Assert.Equal("https://contoso.sharepoint.com/:w:/s/hr/EaBc", SharePointClient.SharingLinkUrl(JsonNode.Parse(answer)));

    [Fact]
    public void A_link_without_an_address_is_none()
    {
        Assert.Null(SharePointClient.SharingLinkUrl(JsonNode.Parse("""{"sharingLinkInfo":{"Url":""}}""")));
        Assert.Null(SharePointClient.SharingLinkUrl(null));
    }

    [Fact]
    public void People_are_one_a_line_or_separated_by_commas()
    {
        Assert.Equal(["alex@contoso.com", "sam@fabrikam.com", "HR Team"],
            CommandCatalog.PeopleOf(" alex@contoso.com,\r\nsam@fabrikam.com; HR Team\n\nALEX@contoso.com "));
        Assert.Empty(CommandCatalog.PeopleOf(" , ;\n"));
    }

    [Fact]
    public void Anyone_links_ask_when_they_expire()
    {
        var link = CommandCatalog.Find("content.share-link")!;
        var days = link.Parameters.Single(p => p.Name == "days");
        var args = link.Defaults();
        Assert.False(days.IsShown(args));
        args["kind"] = "4";
        Assert.True(days.IsShown(args));
        Assert.True(CommandCatalog.Find("content.links-remove")!.Destructive);
    }
}
