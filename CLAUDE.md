# CLAUDE.md

## Project Structure

Scriptless SharePoint manages SharePoint Online without PowerShell. It's a Microsoft Store app (MSIX), WinUI 3 in C# laid out like C:\Dev\FileSize (Settings-style navigation, a tree table with tick boxes, a pane beside it), plus an MCP server for AI apps. README.md covers the sign-in modes, the app registration, the commands and the tools.

- **`src/ScriptlessSharePoint.Core/`** - Shared by the app and the MCP server
  - `Connections/`: saved connections (`connections.json` in `AppInfo.DataDir`, secrets encrypted with DPAPI), `ConnectionDraft` (what the dialog edits) and `Workspace` (saving, testing, the active connection's `Session`, sign-in)
  - `Auth/`: one `ITokenProvider` per sign-in kind
  - `SharePoint/`: `SharePointClient` (CSOM; `.Browse.cs` reads the tree), `SpNode`/`NodeKind` (a row), `TenantUrls`, and the user-facing errors in `SharePointErrors`
  - `Selections/`: saved selections (`SavedSelectionStore`, one JSON file each in `Selections\` in `AppInfo.DataDir`), each for one connection, with the site collections its items are in
  - `History/`: the run history (`RunHistoryStore`, one JSON file per run in `History\` in `AppInfo.DataDir`, the newest 500 kept), saved as a run starts and again as it ends
  - `Settings/`: the app's preferences (`settings.json` in `AppInfo.DataDir`), such as confirming every change
  - `Commands/`: `Command` (targets, parameters, `Run` per target), `CommandCatalog` (one partial file per kind of target; `.Find.cs` has the finders, whose `After.Tick` results become the selection), `CommandRunner` (runs on many targets, 4 at a time), `Csom` (shared helpers)
- **`src/ScriptlessSharePoint/`** - The WinUI 3 app (Windows App SDK 2.x component packages, self-contained, `WindowsPackageType=None`; `build.ps1` packages it). `MainWindow.xaml.cs` has the chrome, connections and sign-in, `MainWindow.Table.cs` the rows (`Row`, opened in place), `MainWindow.Pane.cs` the commands, forms and results, `MainWindow.History.cs` the History page (ticking a run's targets again goes through `ShowSaved` with an unsaved selection), `MainWindow.Saved.cs` the saved selections (opening one loads nothing: each item shows under its saved parent, or the nearest row that `SpNode.Leads` to it, and those rows show only what was saved until reopened). `SSP_DEMO=1` fills the table with made-up data; `SSP_DATA_DIR` moves the data folder
- **`src/ScriptlessSharePoint.Mcp/`** - `scriptless-sharepoint-mcp.exe`, the stdio MCP server, ported from PureSVG's `packages/mcp`. It has a hand-written protocol (`McpServer`), tools (`SharePointTools`) behind `IToolBackend`, and the MCPB manifest (`--manifest`). It signs in silently only
- **`tests/ScriptlessSharePoint.Tests/`** - xUnit tests for URLs, connections, tokens, the workspace, the command catalog and runner, and MCP (against a fake backend)
- **`package/`** - The MSIX: `AppxManifest.xml` template, `identity.json`, `icons.ps1`, and `build.ps1` (publishes both exes self-contained into one folder, then makepri and makeappx)

## Commands

```powershell
dotnet build; dotnet test
dotnet run --project src/ScriptlessSharePoint
./package/build.ps1 [-Register | -Sign]
```

## Notes

- The version lives in `Directory.Build.props` only. The package version is that version plus `.0`
- The app and the MCP server share a package folder, and `build.ps1` fails on differing files
- `build.ps1` merges every `.pri` in the package folder (the app's, WinUI's) into `resources.pri`, one makepri index each, and indexes only `Assets\` as files. A packaged WinUI app reads its XAML from `resources.pri`; without the merge it crashes at start. `EnableMsixTooling` must stay on, or publish leaves out the app's `.pri`
- `Workspace` sessions are made on the UI thread, so `DelegatedTokenProvider` (which captures the UI `SynchronizationContext`) shows WAM there. Call `workspace` from UI code without `ConfigureAwait(false)`
- Adding an operation: add a `Command` to the `CommandCatalog` file for its targets (id, group from `GroupOrder`, description, the cmdlet it stands in for, parameters, `After`, `Destructive`/`NeedsAdmin`/`ReadOnly`). The pane builds its form and runs it; `CommandCatalogTests` checks the catalog. For AI apps, add the tool to `SharePointTools.All`, the `RunAsync` switch and `IToolBackend`
- Modern pages (`CommandCatalog.Pages.cs`) are made and changed through CSOM as PnP Framework does: `AddTemplateFile(…, ClientSidePage)`, then the item's fields (`CanvasContent1` holds the canvas HTML, `LayoutWebpartsContent` the header, `PromotedState` 2 for news). CSOM has no page API
- Making and removing sharing links (`SharePointClient.Sharing.cs`: `items(id)/ShareLink`, `UnshareLink`) is REST too; listing them and inviting people (`Web.ShareObject` with the people picker's entities) is CSOM
- Apps (`CommandCatalog.Apps.cs`) go through the app management REST API (`_api/web/tenantappcatalog`, `sitecollectionappcatalog`), which CSOM lacks: `SharePointClient.RestAsync` sends the connection's bearer token (no request digest needed) and turns 401/403/429/503 into `HttpRequestException` so they're retried and described as CSOM's are
- LINQ on a loaded CSOM collection (`fields.Where(…)`) binds to `IQueryable` and goes to CSOM's query provider: call `.ToList()` first
- `CommandCatalog.All` is built lazily: its static choice lists live in other partial files, whose initialisation order C# doesn't fix
- Values from SharePoint reach agents wrapped in `<sharepoint_data>`
- "All sites" (`SharePointClient.ListSitesAsAdminAsync`) reads the admin centre's own site list, `DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS`. It's undocumented, indexed, and over the 5,000-item threshold. What was checked against a 6,113-site tenant:
  - These work: unfiltered paging sorted by the Title index; one `Contains` (or `And` of `Contains`) on Title or SiteUrl
  - These are refused: `Or`; `Where` combined with `UseIndexForOrderBy`; a term matching thousands of rows
  - So a search is two queries in one cursor (`t:` titles, then `u:` URLs), and tenants without the list fall back to the tenant API (`o:` offsets)
- Sites paging is cursor-based (`SitePage.Next`), passed back as is by the UI and by MCP's `search_sites`
- CSOM calls, WAM and certificate sign-in can't be tested without a tenant. The unit tests cover everything around them
