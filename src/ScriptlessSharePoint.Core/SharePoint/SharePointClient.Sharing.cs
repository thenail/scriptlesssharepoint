using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;

namespace ScriptlessSharePoint.SharePoint;

/// <summary>
/// Sharing links on a list item (a file, folder or item), through SharePoint's
/// REST API: CSOM can list them but not make or remove one. Kinds are
/// SharePoint's SharingLinkKind numbers (2/3 organisation view/edit, 4/5 anyone view/edit).
/// </summary>
public sealed partial class SharePointClient
{
    /// <summary>Makes a sharing link (or gets the one there is of that kind); returns its address.</summary>
    public async Task<string> CreateSharingLinkAsync(string webUrl, Guid listId, int itemId, int linkKind, DateTimeOffset? expires, CancellationToken ct = default)
    {
        var anyone = linkKind is 4 or 5;
        var settings = new JsonObject
        {
            ["linkKind"] = linkKind,
            ["allowAnonymousAccess"] = anyone,
            ["role"] = linkKind is 3 or 5 ? 2 : 1,
            ["scope"] = anyone ? 0 : 1,
            ["restrictShareMembership"] = false,
            ["expiration"] = expires?.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", System.Globalization.CultureInfo.InvariantCulture),
        };
        var body = new JsonObject { ["request"] = new JsonObject { ["createLink"] = true, ["settings"] = settings } };
        var answer = await RestAsync(webUrl, HttpMethod.Post, $"_api/web/lists('{listId}')/items({itemId})/ShareLink", Json(body), ct).ConfigureAwait(false);
        return SharingLinkUrl(answer) ?? throw new SharePointOperationException("SharePoint didn't give the link back.");
    }

    /// <summary>Removes one sharing link, by its kind and ID.</summary>
    public Task RemoveSharingLinkAsync(string webUrl, Guid listId, int itemId, int linkKind, string shareId, CancellationToken ct = default) =>
        RestAsync(webUrl, HttpMethod.Post, $"_api/web/lists('{listId}')/items({itemId})/UnshareLink",
            Json(new JsonObject { ["linkKind"] = linkKind, ["shareId"] = shareId }), ct);

    /// <summary>The link's address in ShareLink's answer (with or without OData's "d" wrapper).</summary>
    internal static string? SharingLinkUrl(JsonNode? answer)
    {
        var info = answer?["sharingLinkInfo"] ?? answer?["d"]?["ShareLink"]?["sharingLinkInfo"] ?? answer?["ShareLink"]?["sharingLinkInfo"];
        return info?["Url"]?.GetValue<string>() is { Length: > 0 } url ? url : null;
    }

    private static StringContent Json(JsonObject body)
    {
        var content = new StringContent(body.ToJsonString(), Encoding.UTF8);
        content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=nometadata");
        return content;
    }
}
