# Scriptless SharePoint

Manage SharePoint Online without PowerShell. A Windows app (Microsoft Store, MSIX) with an MCP server, so AI apps can work with the same connections.

## What it does

- **Every site collection in one table.** Each connection's site collections, from the admin centre (or from search, for accounts that aren't SharePoint administrators). More site collections load as you scroll to the end. Rows open in place: a site to its subsites, lists and libraries, those to their folders, files and items (500 at a time, with a link for the next). Search them from the title bar.
- **Tick, then act.** Tick anything, at any level and in any site. The pane beside the table lists the commands for what's ticked (or for the tenant, when nothing is), grouped as *Create*, *Settings*, *Sharing*, *Permissions* and so on. Each command shows the PowerShell it stands in for.
- **Saved selections.** Save what's ticked under a name (the save button in the pane, or Ctrl+S). Saved selections are listed in the left pane below the connections. Picking one ticks its items again and shows them in the table, with the sites, lists and folders down to each opened. Anything that has gone since is left unticked. Rename or remove one from its right-click menu; removing one only removes the list.
- **One command, many targets.** A command opens its inputs in the pane. *Run* runs it on every ticked thing of that kind, a few at a time, and lists how each went. Anything that deletes or can't easily be undone asks first. *Tick the failed ones* sets up a retry.
- **What the shell does.** About 80 operations, through CSOM, covering what the SharePoint Online Management Shell and PnP PowerShell are used for on these things:

| On | For example |
|---|---|
| The tenant | New site collection, deleted sites (list, restore, purge), tenant sharing and default link |
| Site collections | Storage limit, lock state, custom scripts, unmanaged devices, page comments, sharing, owner and site collection administrators, hub sites, app catalogs, features, recycle bin, delete |
| Sites and subsites | Rename, new subsite, library or list, owners/members/visitors, search visibility and reindex, members can share, property bag, features |
| Lists and libraries | Rename, versioning, check-out, approval, content types, folders, attachments, navigation, hide, search, reindex, add a column, delete |
| Folders, files, items | New folder, upload, rename (with `{name}`), set a column, copy or move (also between sites), download, check in and out, publish, approve or reject, delete previous versions, delete |
| Everything | Open in the browser, show details, show, give, remove, break and reset permissions |

The catalog is in `src/ScriptlessSharePoint.Core/Commands/`, one file per kind of target.

## Signing in

Each connection is one SharePoint Online tenant plus a way of signing in:

| Sign in with | What it needs |
|---|---|
| The account signed in to Windows | An app (client) ID. Signs in silently through Windows (WAM). |
| Another account | An app ID. Shows Windows' account picker, or the browser if *Sign in in the browser* is on. |
| An app with a certificate (app-only) | App ID, tenant ID or domain, and a certificate. The app can import a `.pfx` or create a certificate. SharePoint rejects client secrets for app-only access. |
| An access token | A token for `https://<tenant>.sharepoint.com`, used as-is until it expires. |

Connections are saved in `%LOCALAPPDATA%\ScriptlessSharePoint\connections.json`, and saved selections in `Selections\` beside it, one file each (inside the package's own folder when installed from the MSIX). Pasted tokens are DPAPI-encrypted, and certificates stay in the Windows certificate store. The MSAL token cache sits in the same folder, so a sign-in done in the app also works for the MCP server.

### The Entra app registration

Register this once, then either set its ID as the built-in one (see below) or enter it in each connection:

1. **Entra admin centre › App registrations › New registration.** Choose *Accounts in any organizational directory* (multi-tenant), so every customer tenant can use it.
2. **Authentication › Add a platform › Mobile and desktop applications.** Add these redirect URIs:
   - `ms-appx-web://microsoft.aad.brokerplugin/<client id>` for Windows sign-in (WAM)
   - `http://localhost` for browser sign-in
3. **API permissions › SharePoint › Delegated › `AllSites.FullControl`.** This needs admin consent in each tenant.
4. For **app-only** connections, which use a customer's own registration: add **SharePoint › Application › `Sites.FullControl.All`** (with admin consent), and upload the certificate's `.cer` under *Certificates & secrets*.

Tenant commands, most site collection settings and listing every site need the signed-in account to be a SharePoint administrator; the pane marks those commands. Without admin rights the table lists the sites the account can see (through search), and site, list and file commands work wherever the account has the permissions.

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

In the app, **Settings › AI apps** shows the settings for Claude Desktop, Claude Code and VS Code, for example:

```sh
claude mcp add scriptless-sharepoint -- scriptless-sharepoint-mcp
```

## How it fits together

```
src/ScriptlessSharePoint (WinUI 3)  ──▶  Workspace, CommandCatalog, SharePointClient (Core, CSOM)
                                                ▲
                                                │ same connections, sign-in and operations
AI apps ──stdio, MCP──▶ scriptless-sharepoint-mcp ┘
```

The app is WinUI 3 (Windows App SDK, self-contained), laid out as File Size and Windows Settings are: connections down the left, Mica behind it all. `Workspace` holds the connections and the session of the one in use; `SharePointClient` reads the tree (`SharePointClient.Browse.cs`); `CommandCatalog` has the operations and `CommandRunner` runs one on many targets.

## Development

Requires Windows, the .NET 10 SDK, and the Windows SDK (for packaging).

```powershell
dotnet build                      # everything
dotnet test                       # unit tests (no SharePoint needed)
dotnet run --project src/ScriptlessSharePoint
./package/build.ps1               # MSIX in package/dist/
./package/build.ps1 -Register     # install the package folder for this user (Developer Mode)
./package/build.ps1 -Sign         # sign the .msix with a local test certificate
./package/icons.ps1 -Ico src/ScriptlessSharePoint/Assets/app.ico   # redraw the icon
```

Trying the UI without a tenant, or without touching your real connections:

```powershell
$env:SSP_DEMO = "1"                       # made-up sites, lists and files in the table (commands still go to SharePoint)
$env:SSP_DATA_DIR = "$env:TEMP\ssp-test"  # connections and sign-ins in another folder
dotnet run --project src/ScriptlessSharePoint
```

| Folder | What's in it |
|---|---|
| `src/ScriptlessSharePoint.Core` | Connections and the session (`Workspace`), sign-in (MSAL, WAM, certificates, tokens), the tree and the operations (CSOM), the command catalog (`Commands/`) |
| `src/ScriptlessSharePoint` | The Windows app: WinUI 3. `MainWindow.Table.cs` is the table, `MainWindow.Pane.cs` the pane |
| `src/ScriptlessSharePoint.Mcp` | The MCP server |
| `tests/ScriptlessSharePoint.Tests` | xUnit tests: URLs, tokens and connections, the command catalog and runner, and the MCP protocol against a fake backend |
| `package/` | The AppxManifest template, the Store identity, the icon generator and `build.ps1` |

The app and the MCP server are published self-contained into one folder and share the .NET runtime; `build.ps1` fails if the two would ship different copies of a file. It merges the app's and WinUI's `.pri` files into the package's `resources.pri`, which a packaged app reads its XAML and WinUI's resources from.

### Before submitting to the Store

- Put the Partner Center identity in `package/identity.json`.
- Review the licences of the bundled packages (MSAL is MIT; check the `Microsoft.SharePointOnline.CSOM` licence terms for redistribution) and add a notices file.
- Add a privacy policy URL. The app sends data only to Microsoft Entra ID and the user's SharePoint tenant.

## License

[MIT](LICENSE)
