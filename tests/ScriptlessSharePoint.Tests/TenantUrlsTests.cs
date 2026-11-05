using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public class TenantUrlsTests
{
    [Theory]
    [InlineData("contoso", "https://contoso.sharepoint.com")]
    [InlineData("Contoso", "https://contoso.sharepoint.com")]
    [InlineData("contoso.sharepoint.com", "https://contoso.sharepoint.com")]
    [InlineData("https://contoso.sharepoint.com/sites/hr/SitePages/Home.aspx", "https://contoso.sharepoint.com")]
    [InlineData("https://contoso-admin.sharepoint.com/", "https://contoso.sharepoint.com")]
    [InlineData("https://contoso-my.sharepoint.com/personal/x", "https://contoso.sharepoint.com")]
    [InlineData("https://fabrikam.sharepoint.us", "https://fabrikam.sharepoint.us")]
    public void Normalizes_tenant_urls(string input, string expected)
    {
        Assert.True(TenantUrls.TryNormalizeTenantUrl(input, out var url));
        Assert.Equal(expected, url);
    }

    [Theory]
    [InlineData("")]
    [InlineData("http://contoso.sharepoint.com")]
    [InlineData("not a url")]
    public void Rejects_what_isnt_a_tenant(string input) =>
        Assert.False(TenantUrls.TryNormalizeTenantUrl(input, out _));

    [Fact]
    public void Admin_url_adds_admin_to_the_first_label() =>
        Assert.Equal("https://contoso-admin.sharepoint.com", TenantUrls.AdminUrl("https://contoso.sharepoint.com"));

    [Theory]
    [InlineData("https://contoso.sharepoint.com/sites/hr", "https://contoso.sharepoint.com/sites/hr")]
    [InlineData("https://contoso.sharepoint.com/sites/hr/", "https://contoso.sharepoint.com/sites/hr")]
    [InlineData("/sites/hr", "https://contoso.sharepoint.com/sites/hr")]
    [InlineData("https://contoso.sharepoint.com/sites/hr/SitePages/Home.aspx?x=1#y", "https://contoso.sharepoint.com/sites/hr")]
    [InlineData("https://contoso.sharepoint.com/sites/hr/Shared%20Documents/Forms/AllItems.aspx", "https://contoso.sharepoint.com/sites/hr")]
    [InlineData("https://contoso.sharepoint.com/sites/hr/Lists/Tasks/AllItems.aspx", "https://contoso.sharepoint.com/sites/hr")]
    [InlineData("https://contoso.sharepoint.com/sites/hr/_layouts/15/settings.aspx", "https://contoso.sharepoint.com/sites/hr")]
    [InlineData("https://contoso.sharepoint.com/sites/team a", "https://contoso.sharepoint.com/sites/team%20a")]
    [InlineData("https://contoso.sharepoint.com", "https://contoso.sharepoint.com")]
    public void Site_urls_lose_pages_and_libraries(string input, string expected) =>
        Assert.Equal(expected, TenantUrls.NormalizeSiteUrl(input, "https://contoso.sharepoint.com"));

    [Fact]
    public void Site_urls_must_be_in_the_tenant()
    {
        var e = Assert.Throws<ArgumentException>(() =>
            TenantUrls.NormalizeSiteUrl("https://fabrikam.sharepoint.com/sites/hr", "https://contoso.sharepoint.com"));
        Assert.Contains("contoso.sharepoint.com", e.Message);
    }

    [Theory]
    [InlineData(null, "contentclass:STS_Site -WebTemplate:SPSPERS")]
    [InlineData("  ", "contentclass:STS_Site -WebTemplate:SPSPERS")]
    [InlineData("HR team", "hr* team* contentclass:STS_Site -WebTemplate:SPSPERS")]
    // operators, quotes and property restrictions can't get into the query
    [InlineData("a OR b\" Path:x", "a* or* b* path* x* contentclass:STS_Site -WebTemplate:SPSPERS")]
    public void Site_search_query(string? text, string expected) =>
        Assert.Equal(expected, TenantUrls.SiteSearchQuery(text));
}
