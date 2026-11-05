# Scriptless SharePoint

Manage SharePoint Online without PowerShell. A Windows app (Microsoft Store, MSIX) with an MCP server, so AI apps can do the same work with the same connections.

The first feature is **site collection app catalogs**: find a site, see whether it has its own app catalog, and enable or remove one, as `Add-SPOSiteCollectionAppCatalog` and `Remove-SPOSiteCollectionAppCatalog` do.

## Signing in

Each connection is one SharePoint Online tenant plus a way of signing in:

| Sign in with | What it needs |
|---|---|
| The account signed in to Windows | An app (client) ID. Signs in silently through Windows (WAM). |
| Another account | An app ID. Shows Windows' account picker, or the browser if *Sign in in the browser* is on. |
| An app with a certificate (app-only) | App ID, tenant ID or domain, and a certificate. The app can import a `.pfx` or create a certificate. SharePoint rejects client secrets for app-only access. |
| An access token | A token for `https://<tenant>.sharepoint.com`, used as-is until it expires. |

Connections are saved in `%LOCALAPPDATA%\ScriptlessSharePoint\connections.json` (inside the package's own folder when installed from the MSIX). Pasted tokens are DPAPI-encrypted, and certificates stay in the Windows certificate store. The MSAL token cache sits in the same folder, so a sign-in done in the app also works for the MCP server.

### The Entra app registration

Register this once, then either set its ID as the built-in one (see below) or enter it in each connection:

1. **Entra admin centre › App registrations › New registration.** Choose *Accounts in any organizational directory* (multi-tenant), so every customer tenant can use it.
2. **Authentication › Add a platform › Mobile and desktop applications.** Add these redirect URIs:
   - `ms-appx-web://microsoft.aad.brokerplugin/<client id>` for Windows sign-in (WAM)
   - `http://localhost` for browser sign-in
3. **API permissions › SharePoint › Delegated › `AllSites.FullControl`.** This needs admin consent in each tenant.
4. For **app-only** connections, which use a customer's own registration: add **SharePoint › Application › `Sites.FullControl.All`** (with admin consent), and upload the certificate's `.cer` under *Certificates & secrets*.

Enabling or removing app catalogs, and listing *All sites*, also require the signed-in account to be a SharePoint administrator. Without admin rights, search still works (it shows the sites the account can see), and so does reading a site.

**Built-in app ID:** `17316cc6-b7c2-40a2-90c7-a7b76e7f2758`, set as `SspClientId` in `src/ScriptlessSharePoint.Core/ScriptlessSharePoint.Core.csproj`. Connections whose app ID is left empty use it. To build with a different one, use `package/identity.json` (`clientId`), `SSP_CLIENT_ID`, or `dotnet build -p:SspClientId=<guid>`.

## MCP server

`scriptless-sharepoint-mcp.exe` is a stdio MCP server that ships in the package. It follows PureSVG's design: a hand-written JSON-RPC 2.0 protocol, its own execution alias, and registration with Windows' on-device agent registry (`com.microsoft.windows.ai.mcpServer`, plus `mcp/manifest.json` with static `initialize` and `tools/list` responses).

| Tool | What it does |
|---|---|
| `list_connections` | The saved connections and which one is active |
| `get_connection` | Signs in and says who the connection acts as |
| `search_sites` | Finds sites through search (`scope: "search"`) or the admin centre (`"admin"`) |
| `get_site` | A site collection and whether it has its own app catalog |
| `list_site_app_catalogs` | Every site collection app catalog in the tenant |
| `enable_site_app_catalog` | Gives a site collection its own app catalog |
| `disable_site_app_catalog` | Removes one (destructive; asks the user through MCP elicitation when the client supports it) |
| `open_app` | Opens the app, for example to sign in |

The server only signs in silently. When a connection needs interactive sign-in, the tool says so, and the user signs in in the app. Values read from SharePoint come back inside `<sharepoint_data>` blocks, and the server's instructions tell the agent to treat them as data, not instructions.

In the app, **⋯ › Connect an AI app** shows the settings for Claude Desktop, Claude Code and VS Code, for example:

```sh
claude mcp add scriptless-sharepoint -- scriptless-sharepoint-mcp
```

## How it fits together

```
web/ (React + Fluent UI v9)  ──postMessage, JSON-RPC──▶  AppBridge (Core)  ──▶  SharePointClient (CSOM)
        hosted in WebView2 by src/ScriptlessSharePoint          ▲
                                                                 │ same connections, sign-in and operations
AI apps ──stdio, MCP──▶ scriptless-sharepoint-mcp ──────────────┘
```

The app is a small WPF window that hosts the web UI in WebView2. The UI's built files are served from the package folder on a virtual `https://app.scriptless-sharepoint.local/` origin. Only pages from that origin can talk to the bridge, and any other link opens in the browser. The bridge (`AppBridge`) answers JSON-RPC calls from the page, and the host supplies the window for Windows sign-in, file dialogs, and opening links.

## Development

Requires Windows, the .NET 10 SDK, Node with Yarn 1, and the Windows SDK (for packaging).

```powershell
dotnet build                      # everything; builds web/ the first time
dotnet test                       # .NET unit tests (no SharePoint needed)
dotnet run --project src/ScriptlessSharePoint
./package/build.ps1               # MSIX in package/dist/ (always rebuilds web/)
./package/build.ps1 -Register     # install the package folder for this user (Developer Mode)
./package/build.ps1 -Sign         # sign the .msix with a local test certificate
./package/icons.ps1 -Ico src/ScriptlessSharePoint/Assets/app.ico   # redraw the icon
```

Working on the UI:

```powershell
cd web
yarn dev                          # http://localhost:5279, with made-up data in a browser
yarn test; yarn typecheck; yarn build
# the real app with the live UI (Debug builds have DevTools: F12)
$env:SSP_WEB_URL = "http://localhost:5279"; dotnet run --project ../src/ScriptlessSharePoint
# ...or the live UI with made-up data in the app
$env:SSP_WEB_URL = "http://localhost:5279/?mock"; dotnet run --project ../src/ScriptlessSharePoint
```

A normal `dotnet build` builds `web/dist` only when it's missing. After changing the UI, run `yarn build` (or use `yarn dev`).

| Folder | What's in it |
|---|---|
| `web/` | The UI: React 19, Fluent UI v9, Vite. `src/bridge.ts` is the typed client for the bridge, and `src/mock.ts` supplies the made-up data for development |
| `src/ScriptlessSharePoint.Core` | Connections, sign-in (MSAL, WAM, certificates, tokens), SharePoint operations (CSOM), and the UI's bridge (`Bridge/`) |
| `src/ScriptlessSharePoint` | The Windows app: the WebView2 host |
| `src/ScriptlessSharePoint.Mcp` | The MCP server |
| `tests/ScriptlessSharePoint.Tests` | xUnit tests: URLs, tokens and connections, the bridge, and the MCP protocol against a fake backend |
| `package/` | The AppxManifest template, the Store identity, the icon generator and `build.ps1` |

The app and the MCP server are published self-contained into one folder and share the .NET runtime. For that to work, the MCP project uses the same WPF runtime profile as the app; `build.ps1` fails if the two would ship different copies of a file. The app needs the WebView2 runtime, which is part of Windows 11 and installed on almost every Windows 10 PC. If it's missing, the app offers its download page.

### Before submitting to the Store

- Put the Partner Center identity in `package/identity.json`.
- Review the licences of the bundled packages (MSAL is MIT; check the `Microsoft.SharePointOnline.CSOM` licence terms for redistribution) and add a notices file.
- Add a privacy policy URL. The app sends data only to Microsoft Entra ID and the user's SharePoint tenant.

## License

[MIT](LICENSE)
