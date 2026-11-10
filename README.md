# Scriptless SharePoint

Manage SharePoint Online without PowerShell. A Windows app (Microsoft Store, MSIX) with an MCP server, so AI apps can work with the same connections.

## What it does

- **Every site collection in one table.** Each connection's site collections, from the admin centre (or from search, for accounts that aren't SharePoint administrators). More site collections load as you scroll to the end. Rows open in place: a site to its subsites, lists and libraries, those to their folders, files and items (500 at a time, with a link for the next). Search them from the title bar. Right-click a row (or Shift+F10) to open it in the browser, copy its address, or tick everything beside it or in it; on the keyboard, Space ticks the row, Right and Left open and close it.
- **Tick, then act.** Tick anything, at any level and in any site. The pane beside the table lists the commands for what's ticked (or for the tenant, when nothing is), with a tab for each kind ticked (*Site collections 2 · Libraries 1 · Tenant*), grouped as *Create*, *Settings*, *Sharing*, *Permissions* and so on. Each command shows the PowerShell it stands in for. Whatever replaces what's ticked (a saved or predefined selection, *Tick inside*, Clear) can be undone: *Undo* in the message, beside Clear, or Ctrl+Z ticks again what was ticked before.
- **Predefined selections.** Below the connections, ready-made rules in three parts. *Site collections* and *OneDrives* are picked from the admin centre's properties (needs a SharePoint administrator): inactive for 90 days or a year, new, hub sites, connected to a hub or not, Microsoft 365 group and Teams sites, communication and classic sites, "Anyone" links or guests allowed, locked, near the storage limit or large, without a sensitivity label, archived. *Files and pages* come from SharePoint search (any account; a few minutes behind, and only what the account can open): files over 100 MB or 1 GB, unchanged for 3 years, that guests can open, .pst mailboxes, videos, programs and scripts, news posts, pages unchanged for a year. Picking one shows what it picks ticked, each file under its library and site; F5 runs it again, and Save keeps it as a saved selection. Inside ticked sites, *Tick inside* also finds lists by their settings (own permissions, no versions, over 5,000 items…) and runs the same searches, or your own KQL.
- **Saved selections.** Save what's ticked under a name (the save button in the pane, or Ctrl+S). Saved selections are listed in the left pane below the connections. Picking one ticks its items again and shows each in the table under the subsite, list or folder it was ticked in, without loading the rest of that parent. Close and open the parent to see everything in it. With one open, Save keeps what's ticked now in it (ticked items added, unticked ones removed); *Save as a new selection…* (the arrow beside Save, or Ctrl+Shift+S) makes another. Drag rows from the table onto a saved selection to add them (a ticked row brings everything ticked with it). Rename or remove one from its right-click menu; removing one only removes the list. Saving what a predefined selection picked can keep the rule instead of the list: opening it then picks again (a refresh icon marks these, and rows can't be dropped onto them).
- **One command, many targets.** A command opens its inputs in the pane. *Run* runs it on every ticked thing of that kind, a few at a time, each one's row showing it waiting, running, then how it went. Anything that deletes or can't easily be undone asks first; with *Confirm every change* on in Settings, every command that changes SharePoint asks about each item in its row instead (Run or Skip, or Run all). *Tick the failed ones* sets up a retry.
- **Jobs.** Every command or plan runs as a job in the background: the pane follows it, and *Close* leaves it running. The Jobs page lists the jobs running or paused, then the finished ones (the last 500, on this PC: the command, the connection, the inputs and how each target went). A job can be paused (the targets already running finish, the rest wait), carried on and cancelled; jobs running when the app closed come back paused, and carry on from where they were (a plan from its step). Jobs follow Microsoft's throttling guidance together: when SharePoint answers 429 or 503, every job on that tenant pauses for as long as SharePoint asks (*SharePoint asked to slow down: carrying on in 30 s*), fewer run at once for a while (6 at most, 2 after throttling, rising again as requests succeed), and throttled items keep trying for up to 10 minutes before they fail. From a finished job, *Tick these* or *Tick the failed ones* shows its targets ticked again, and *Run again…* opens the command with the same inputs.
- **Tick inside what's ticked.** *Tick inside* finds the same things in every ticked site, list, library or folder and ticks them in its place, ready for one command on all of them: subsites (one level or all), libraries and lists (Site Pages, Documents, Site Assets, by type, every one, or by the name in the address, never by the translated title), the home page, what's at an address (`Shared Documents/Policies`), and files, folders or list items by name (`*.docx`, `Budget`) or check-out, in subfolders too. The table opens each site to just what was found, so it can be dragged onto a saved selection; closing and opening a row shows everything in it again.
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
| `open_app` | Opens the app, for example to sign in, or showing a site or a plan |
| `list_commands` | The app's commands a plan can use, with their parameters and what they run on |
| `list_presets` | The predefined selections a plan can run on |
| `list_plans`, `get_plan` | The saved plans; a plan in full, with how each step went the last time it ran |
| `create_plan`, `update_plan` | Saves a plan (checked step by step first; one with problems isn't saved) |
| `prepare_run` | Makes a run of a plan ready: what to run it on, and its variables' values (for all, or per site) |
| `delete_plan` | Deletes a plan (nothing in SharePoint changes) |

**Plans** are how an AI app gets anything else done. A plan is steps: commands from the catalog, with their inputs, run in order on whatever is ticked when it runs. It isn't tied to items, so the same plan runs on 3 test sites and then on 100. A finder step ("Tick inside") ticks what it finds for the steps after it, so "the Site Pages library of each site, then 100 versions in it" is two steps.

Inputs can hold **variables**. Built-in ones come from each item (`{title}`, `{url}`, `{name}`, `{site}`, `{web}`, `{today}`), so *Rename* to `{title} (archived)` gives each site its own name. A plan's own variables (`{owner}`, `{quota}`) get their values when it runs: one value for all, or one per item from a table pasted from Excel or loaded from a CSV (the address first, then a column per variable). A row covers its site and everything in it, so libraries and files a finder step found take their site's values. Before running, the app checks that every item has its values and that they fit (quota must be a number); a value that's still wrong at run time fails only that item.

Plans change nothing by themselves: they appear in the app (on the Plans page, and first in the pane's command list for what's ticked) as soon as they're saved, with a message offering to open a new one (or a run an AI app has just prepared), the user opens one on what's ticked, gives the values, and runs it there. An AI app can prepare a run (`prepare_run`: the sites and a value table) that the user then only looks over and starts. Each step's run goes in the history, which is where `get_plan` reads the results. Plans can also be made in the app: *Add to plan* beside a command's Run keeps the command and its inputs (with any `{variables}` typed in them) as a step, and in an open plan *Add a step* offers the commands that fit where the plan has got to (after a finder, what it finds), so nothing needs ticking or running to build one. Each step's inputs can be changed, and steps moved or removed; a plan's menu (on the Plans page) renames, duplicates or deletes it, and *New plan* there starts one from nothing. A plan stays open in the pane while you pick other selections on the left: the left pane says what to work on, the pane what to do with it. *Run it on…* ticks the plan's test selection, a saved selection or a predefined one without leaving the plan.

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
