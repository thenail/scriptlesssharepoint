// The MCP bundle manifest (MCPB, manifest_version 0.3) that registers the
// server with Windows' on-device agent registry when the Store package is
// installed (package/AppxManifest.xml, com.microsoft.windows.ai.mcpServer).
// Windows answers initialize and tools/list from _meta without starting the
// server, so they're built from the same code the server answers with.
using System.Text.Json.Nodes;

namespace ScriptlessSharePoint.Mcp;

public static class Manifest
{
    public static JsonObject Bundle(string version) => new()
    {
        ["manifest_version"] = "0.3",
        ["name"] = McpServer.ServerName,
        ["display_name"] = AppInfo.Name,
        ["version"] = version,
        ["description"] = "Find SharePoint Online sites and enable site collection app catalogs, with the connections of the Scriptless SharePoint app.",
        ["long_description"] =
            "Lets AI apps search the SharePoint Online sites of your tenant, check whether a site collection has its own app catalog, and enable or remove one, without PowerShell. It signs in with the connections you set up in the Scriptless SharePoint app: your Windows account, another account, an app with a certificate, or an access token.",
        ["author"] = new JsonObject { ["name"] = AppInfo.Name },
        ["repository"] = new JsonObject { ["type"] = "git", ["url"] = "https://github.com/thenail/scriptlesssharepoint" },
        ["license"] = "MIT",
        ["icon"] = "icon.png",
        ["keywords"] = new JsonArray("sharepoint", "microsoft 365", "app catalog", "spfx", "admin"),
        ["server"] = new JsonObject
        {
            ["type"] = "binary",
            ["entry_point"] = AppInfo.McpExe,
            ["mcp_config"] = new JsonObject { ["command"] = AppInfo.McpExe, ["args"] = new JsonArray() },
        },
        ["tools"] = new JsonArray([.. SharePointTools.All.Select(t => (JsonNode)new JsonObject
        {
            ["name"] = t.Name,
            ["description"] = t.Description,
        })]),
        ["tools_generated"] = false,
        ["compatibility"] = new JsonObject { ["platforms"] = new JsonArray("win32") },
        ["_meta"] = new JsonObject
        {
            ["com.microsoft.windows"] = new JsonObject
            {
                ["static_responses"] = new JsonObject
                {
                    ["initialize"] = McpServer.InitializeResult(version),
                    ["tools/list"] = new JsonObject { ["tools"] = SharePointTools.ToolListJson() },
                },
            },
        },
    };
}
