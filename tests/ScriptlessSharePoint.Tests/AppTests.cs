using System.Text.Json.Nodes;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class AppTests
{
    private const string Answer = """
        {"value":[
          {"ID":"6a1d4a6b-8f2c-4f2e-9d55-1b2f0c4c1a01","Title":"News ticker","AppCatalogVersion":"1.2.0.0","InstalledVersion":"1.1.0.0","Deployed":true,"CanUpgrade":true,"IsClientSideSolution":true},
          {"ID":"6a1d4a6b-8f2c-4f2e-9d55-1b2f0c4c1a02","Title":"Forms","AppCatalogVersion":"2.0.0.0","InstalledVersion":"","Deployed":false,"CanUpgrade":false,"IsClientSideSolution":true},
          {"ID":"6a1d4a6b-8f2c-4f2e-9d55-1b2f0c4c1a03","Title":"forms","AppCatalogVersion":"1.0.0.0","Deployed":true},
          {"Title":"No ID"}
        ]}
        """;

    private static readonly IReadOnlyList<CatalogApp> Apps = SharePointClient.ParseApps(JsonNode.Parse(Answer));

    [Fact]
    public void Reads_the_catalogs_answer()
    {
        Assert.Equal(3, Apps.Count);
        var ticker = Apps[0];
        Assert.Equal("News ticker", ticker.Title);
        Assert.True(ticker.Installed && ticker.Deployed && ticker.CanUpgrade && ticker.ClientSide);
        Assert.Equal("1.2.0.0", ticker.CatalogVersion);
        Assert.False(Apps[1].Installed);
        Assert.False(Apps[2].Installed);
    }

    [Fact]
    public void Finds_apps_by_id_or_a_title_only_one_has()
    {
        Assert.Same(Apps[0], CommandCatalog.FindApp(Apps, " news TICKER "));
        Assert.Same(Apps[2], CommandCatalog.FindApp(Apps, "6a1d4a6b-8f2c-4f2e-9d55-1b2f0c4c1a03"));
        Assert.Null(CommandCatalog.FindApp(Apps, "Forms"));
        Assert.Null(CommandCatalog.FindApp(Apps, "Weather"));
    }

    [Theory]
    [InlineData("""{"odata.error":{"code":"-1, System.ArgumentException","message":{"lang":"en-US","value":"The app is in use."}}}""", "The app is in use.")]
    [InlineData("""{"error":{"code":"x","message":{"value":"Not deployed."}}}""", "Not deployed.")]
    [InlineData("<html>Bad gateway</html>", null)]
    [InlineData("", null)]
    public void Reads_sharepoints_rest_errors(string body, string? message) =>
        Assert.Equal(message, SharePointClient.RestError(body));
}
