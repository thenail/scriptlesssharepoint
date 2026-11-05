# CLAUDE.md

## Project Structure

Scriptless SharePoint manages SharePoint Online without PowerShell. It's a Microsoft Store app (MSIX) with a React and Fluent UI interface hosted in WebView2, plus an MCP server for AI apps. README.md covers the sign-in modes, the app registration, the tools and the architecture.

- **`web/`** - The UI (React 19, Fluent UI v9, Vite, TypeScript; Yarn 1)
  - `src/bridge.ts` is the typed JSON-RPC client for the host, and `src/types.ts` mirrors the C# records (camelCase; enums by name)
  - `src/mock.ts` stands in for the host in a browser (`yarn dev`) or with `?mock`
  - Components are in `src/components/`; `App.tsx` holds the state
- **`src/ScriptlessSharePoint.Core/`** - Shared by the app and the MCP server
  - `Connections/`: saved connections (`connections.json` in `AppInfo.DataDir`, secrets encrypted with DPAPI)
  - `Auth/`: one `ITokenProvider` per sign-in kind
  - `SharePoint/`: `SharePointClient` (CSOM), `TenantUrls`, and the user-facing errors in `SharePointErrors`
  - `Bridge/`: `AppBridge`, the UI's JSON-RPC backend (one method per operation), and `IHostUi`, which is what it needs from the window
- **`src/ScriptlessSharePoint/`** - The WebView2 host. `MainWindow` maps `web\` (copied from `web/dist`) to `https://app.scriptless-sharepoint.local/`, passes messages from that origin only to `AppBridge`, and opens every other link in the browser. `SSP_WEB_URL` (loopback only) loads `yarn dev` instead
- **`src/ScriptlessSharePoint.Mcp/`** - `scriptless-sharepoint-mcp.exe`, the stdio MCP server, ported from PureSVG's `packages/mcp`. It has a hand-written protocol (`McpServer`), tools (`SharePointTools`) behind `IToolBackend`, and the MCPB manifest (`--manifest`). It signs in silently only
- **`tests/ScriptlessSharePoint.Tests/`** - xUnit tests for URLs, connections, tokens, the bridge, and MCP (against a fake backend)
- **`package/`** - The MSIX: `AppxManifest.xml` template, `identity.json`, `icons.ps1`, and `build.ps1` (yarn build, publishes both exes self-contained into one folder, then makepri and makeappx)

## Commands

```powershell
dotnet build; dotnet test
cd web; yarn dev; yarn test; yarn typecheck; yarn build
./package/build.ps1 [-Register | -Sign]
```

## Notes

- The version lives in `Directory.Build.props` only. The package version is that version plus `.0`
- The MCP project sets `UseWPF` only to get the app's runtime profile. The two share a package folder, and `build.ps1` fails on differing files. Because of `UseWPF`, `System.IO` isn't an implicit using there
- Bridge requests run on the UI thread, so `DelegatedTokenProvider` (which captures the UI `SynchronizationContext`) and the file dialogs open there
- Adding an operation:
  - Add a method to `AppBridge` and to `api()` in `web/src/bridge.ts`, plus `mock.ts`
  - For AI apps, add the tool to `SharePointTools.All`, the `RunAsync` switch and `IToolBackend`. The Windows registry manifest picks it up
- Secrets never go to the page. Tokens come in through `connections.save`/`connections.test` and only `hasToken` goes back. Values from SharePoint reach agents wrapped in `<sharepoint_data>`
- "All sites" (`SharePointClient.ListSitesAsAdminAsync`) reads the admin centre's own site list, `DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS`. It's undocumented, indexed, and over the 5,000-item threshold. What was checked against a 6,113-site tenant:
  - These work: unfiltered paging sorted by the Title index; one `Contains` (or `And` of `Contains`) on Title or SiteUrl
  - These are refused: `Or`; `Where` combined with `UseIndexForOrderBy`; a term matching thousands of rows
  - So a search is two queries in one cursor (`t:` titles, then `u:` URLs), and tenants without the list fall back to the tenant API (`o:` offsets)
- Sites paging is cursor-based (`SitePage.Next`), passed back as is by the UI and by MCP's `search_sites`
- The dev server port is 5279, because 5173 is often taken by other projects on this machine
- CSOM calls, WAM and certificate sign-in can't be tested without a tenant. The unit tests cover everything around them
