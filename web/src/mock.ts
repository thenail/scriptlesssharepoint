// A stand-in for the host when the UI runs in a browser (`yarn dev` without
// the app): made-up connections and sites, so the UI can be worked on alone.
import type { Transport } from "./bridge";
import type { Connection, SiteDetails, SiteSummary } from "./types";

const sites: SiteSummary[] = [
  { url: "https://contoso.sharepoint.com", title: "Contoso intranet", template: "SITEPAGEPUBLISHING#0", lastModified: "2026-10-30T09:12:00Z" },
  { url: "https://contoso.sharepoint.com/sites/hr", title: "Human resources", template: "STS#3", lastModified: "2026-11-02T14:03:00Z" },
  { url: "https://contoso.sharepoint.com/sites/finance", title: "Finance", template: "GROUP#0", lastModified: "2026-10-12T08:00:00Z" },
  { url: "https://contoso.sharepoint.com/sites/spfx-dev", title: "SPFx development", template: "STS#3", lastModified: "2026-11-03T16:45:00Z" },
  { url: "https://contoso.sharepoint.com/sites/marketing", title: "Marketing", template: "SITEPAGEPUBLISHING#0", lastModified: "2026-09-21T11:30:00Z" },
];
// enough made-up sites to scroll through
const words = ["Project", "Team", "Sales", "Legal", "Design", "Support", "Research", "Events", "Board", "Partners"];
for (let i = 1; i <= 320; i++) {
  const name = `${words[i % words.length]} ${words[(i * 7) % words.length]} ${i}`;
  sites.push({ url: `https://contoso.sharepoint.com/sites/${name.replace(/ /g, "").toLowerCase()}`, title: name, template: i % 3 ? "STS#3" : "GROUP#0", lastModified: new Date(Date.UTC(2026, i % 12, (i % 27) + 1)).toISOString() });
}
const catalogs = new Set(["https://contoso.sharepoint.com/sites/spfx-dev"]);
let connections: Connection[] = [
  { id: "demo", name: "Contoso", kind: "User", tenantUrl: "https://contoso.sharepoint.com", useWindowsAccount: true, useBrowser: false, accountName: "ada@contoso.com", kindLabel: "Windows account" },
];
let activeId: string | undefined = "demo";
const identity = { displayName: "Ada Lovelace", email: "ada@contoso.com", isApp: false, tenantTitle: "Contoso" };

const details = (url: string): SiteDetails => {
  const s = sites.find((x) => x.url === url) ?? sites[0];
  const on = catalogs.has(s.url);
  return {
    url: s.url, title: s.title, webTemplate: s.template ?? "STS#3", siteId: "6b0e…", appCatalog: on ? "Enabled" : "NotEnabled",
    appCatalogSource: "tenant", appCatalogUrl: on ? `${s.url}/AppCatalog` : undefined, appCount: on ? 3 : undefined,
  };
};

const handlers: Record<string, (p: Record<string, unknown>) => unknown> = {
  "app.info": () => ({ name: "Scriptless SharePoint", version: "dev", dataDir: "%LOCALAPPDATA%\\ScriptlessSharePoint", mcpCommand: "scriptless-sharepoint-mcp" }),
  "connections.list": () => ({ activeId, connections }),
  "connections.save": (p) => {
    const c = { ...(p.connection as Connection) };
    c.id ??= crypto.randomUUID();
    connections = [...connections.filter((x) => x.id !== c.id), c];
    activeId = c.id;
    return { connection: c, activeId };
  },
  "connections.delete": (p) => {
    connections = connections.filter((x) => x.id !== p.id);
    activeId = connections[0]?.id;
    return { activeId, connections };
  },
  "connections.setActive": (p) => ((activeId = p.id as string), { activeId }),
  "connections.test": (p) => ({ id: (p.connection as Connection).id ?? "new", identity, accountName: identity.email }),
  "session.status": () => ({ connectionId: activeId, identity, accountName: identity.email }),
  "session.signIn": () => ({ identity, accountName: identity.email }),
  "session.signOut": () => null,
  "sites.search": (p) => {
    const q = String(p.query ?? "").toLowerCase().split(/\s+/).filter(Boolean);
    if (q.length === 1 && q[0].length === 1) {
      throw new Error(`Too many sites match "${q[0]}" for SharePoint to search them. Type more of the site's name or URL.`);
    }
    const found = sites
      .filter((s) => q.every((w) => s.title.toLowerCase().includes(w) || s.url.includes(w)))
      .sort((a, b) => a.title.localeCompare(b.title));
    const start = Number(p.cursor ?? 0);
    const end = start + Number(p.limit ?? 50);
    return {
      sites: found.slice(start, end),
      next: end < found.length ? String(end) : undefined,
      total: p.scope === "admin" && q.length ? undefined : found.length,
      totalIsEstimate: p.scope === "admin",
      source: p.scope,
    };
  },
  "sites.get": (p) => details(p.url as string),
  "catalogs.list": () => [...catalogs].map((url) => ({ url, siteId: "6b0e…" })),
  "catalogs.enable": (p) => (catalogs.add(p.url as string), { siteUrl: p.url, changed: true, message: `Enabled the site collection app catalog on ${p.url}.` }),
  "catalogs.disable": (p) => (catalogs.delete(p.url as string), { siteUrl: p.url, changed: true, message: `Removed the site collection app catalog from ${p.url}.` }),
  "certificates.pickPfx": () => ({ path: "C:\\certs\\app.pfx", fileName: "app.pfx" }),
  "certificates.importPfx": () => ({ thumbprint: "A1B2C3D4E5F60718293A4B5C6D7E8F9012345678", description: "CN=Contoso app, valid until 11/4/2028" }),
  "certificates.create": () => ({ thumbprint: "0F1E2D3C4B5A69788796A5B4C3D2E1F00F1E2D3C", description: "CN=Scriptless SharePoint, valid until 11/4/2028", cerPath: "C:\\certs\\Scriptless-SharePoint.cer" }),
  "certificates.describe": () => ({ ok: true, description: "CN=Contoso app, valid until 11/4/2028" }),
  "tokens.inspect": () => ({ audience: "https://contoso.sharepoint.com", expires: new Date(Date.now() + 3600e3).toISOString(), subject: "ada@contoso.com", isForSharePoint: true, expired: false }),
  "tenant.normalize": (p) => {
    const name = String(p.input ?? "").trim().replace(/^https?:\/\//, "").split(/[./]/)[0].replace(/-(admin|my)$/, "");
    return name ? { tenantUrl: `https://${name}.sharepoint.com`, adminUrl: `https://${name}-admin.sharepoint.com` } : null;
  },
  "shell.openUrl": (p) => (window.open(p.url as string, "_blank"), null),
};

export function mockTransport(): Transport {
  let listener: (m: string) => void = () => {};
  return {
    listen: (h) => (listener = h),
    post: (text) => {
      const { id, method, params } = JSON.parse(text);
      setTimeout(() => {
        const run = handlers[method];
        let reply: object;
        try {
          reply = run
            ? { jsonrpc: "2.0", id, result: run(params) }
            : { jsonrpc: "2.0", id, error: { code: -32601, message: `Method not found: ${method}` } };
        } catch (e) {
          reply = { jsonrpc: "2.0", id, error: { code: -32000, message: (e as Error).message } };
        }
        listener(JSON.stringify(reply));
      }, 250);
    },
  };
}
