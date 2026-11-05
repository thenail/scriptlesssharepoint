// JSON-RPC 2.0 with the host (AppBridge) over WebView2's postMessage. In a
// browser (`yarn dev` without the app) a mock answers instead.
import type {
  AppCatalogChange,
  AppInfo,
  CertificateResult,
  Connection,
  Connections,
  Identity,
  SessionStatus,
  SiteAppCatalog,
  SiteDetails,
  SitePage,
  TokenDetails,
} from "./types";

/** A failed call: the message is meant for the user. */
export class BridgeError extends Error {
  constructor(
    message: string,
    readonly signInRequired = false,
  ) {
    super(message);
  }
}

export interface Transport {
  post(message: string): void;
  listen(handler: (message: string) => void): void;
}

interface WebView {
  postMessage(message: string): void;
  addEventListener(type: "message", handler: (e: { data: unknown }) => void): void;
}

const webview = (globalThis as { chrome?: { webview?: WebView } }).chrome?.webview;

export function hostTransport(): Transport | null {
  if (!webview) {
    return null;
  }
  return {
    post: (m) => webview.postMessage(m),
    listen: (h) => webview.addEventListener("message", (e) => typeof e.data === "string" && h(e.data)),
  };
}

type Pending = { resolve(v: unknown): void; reject(e: Error): void };

export class Bridge {
  private nextId = 1;
  private readonly pending = new Map<number, Pending>();
  private readonly listeners = new Map<string, Set<(params: unknown) => void>>();

  constructor(private readonly transport: Transport) {
    transport.listen((text) => this.receive(text));
  }

  call<T>(method: string, params?: Record<string, unknown>): Promise<T> {
    const id = this.nextId++;
    return new Promise<T>((resolve, reject) => {
      this.pending.set(id, { resolve: resolve as (v: unknown) => void, reject });
      this.transport.post(JSON.stringify({ jsonrpc: "2.0", id, method, params: params ?? {} }));
    });
  }

  /** Listens for a notification from the host; returns the unsubscribe. */
  on(method: string, handler: (params: unknown) => void): () => void {
    const set = this.listeners.get(method) ?? new Set();
    set.add(handler);
    this.listeners.set(method, set);
    return () => set.delete(handler);
  }

  private receive(text: string) {
    let msg: { id?: number; method?: string; params?: unknown; result?: unknown; error?: { message: string; data?: { signInRequired?: boolean } } };
    try {
      msg = JSON.parse(text);
    } catch {
      return;
    }
    if (msg.method) {
      this.listeners.get(msg.method)?.forEach((h) => h(msg.params));
      return;
    }
    const waiter = msg.id === undefined ? undefined : this.pending.get(msg.id);
    if (!waiter) {
      return;
    }
    this.pending.delete(msg.id!);
    if (msg.error) {
      waiter.reject(new BridgeError(msg.error.message, !!msg.error.data?.signInRequired));
    } else {
      waiter.resolve(msg.result ?? null);
    }
  }
}

/** The host's operations, typed. */
export function api(bridge: Bridge) {
  const c = bridge.call.bind(bridge);
  return {
    info: () => c<AppInfo>("app.info"),
    connections: {
      list: () => c<Connections>("connections.list"),
      save: (connection: Connection, token?: string) =>
        c<{ connection: Connection; activeId: string }>("connections.save", { connection, token }),
      remove: (id: string) => c<Connections>("connections.delete", { id }),
      setActive: (id: string) => c<{ activeId: string }>("connections.setActive", { id }),
      test: (connection: Connection, token?: string) =>
        c<{ id: string; identity: Identity; accountName?: string }>("connections.test", { connection, token }),
    },
    session: {
      status: () => c<SessionStatus>("session.status"),
      signIn: () => c<{ identity: Identity; accountName?: string }>("session.signIn"),
      signOut: () => c<null>("session.signOut"),
    },
    sites: {
      search: (p: { query?: string; scope: "search" | "admin"; cursor?: string; limit?: number }) => c<SitePage>("sites.search", p),
      get: (url: string) => c<SiteDetails>("sites.get", { url }),
    },
    catalogs: {
      list: () => c<SiteAppCatalog[]>("catalogs.list"),
      enable: (url: string) => c<AppCatalogChange>("catalogs.enable", { url }),
      disable: (url: string) => c<AppCatalogChange>("catalogs.disable", { url }),
    },
    certificates: {
      pickPfx: () => c<{ path: string; fileName: string } | null>("certificates.pickPfx"),
      importPfx: (path: string, password?: string) => c<CertificateResult>("certificates.importPfx", { path, password }),
      create: (name: string) => c<CertificateResult | null>("certificates.create", { name }),
      describe: (thumbprint: string) => c<{ ok: boolean; description: string }>("certificates.describe", { thumbprint }),
    },
    tokens: { inspect: (token: string) => c<TokenDetails | null>("tokens.inspect", { token }) },
    tenant: { normalize: (input: string) => c<{ tenantUrl: string; adminUrl: string } | null>("tenant.normalize", { input }) },
    openUrl: (url: string) => c<null>("shell.openUrl", { url }),
    onAccountChanged: (h: (p: { connectionId: string; accountName?: string }) => void) =>
      bridge.on("session.accountChanged", h as (p: unknown) => void),
  };
}

export type Api = ReturnType<typeof api>;
