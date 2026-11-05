// What the host's AppBridge (src/ScriptlessSharePoint.Core/Bridge) sends and
// takes: C# records in camelCase, enums as their names.

export type ConnectionKind = "User" | "Certificate" | "AccessToken";

export interface Connection {
  id?: string;
  name: string;
  kind: ConnectionKind;
  tenantUrl: string;
  adminUrl?: string;
  clientId?: string;
  tenantId?: string;
  useWindowsAccount: boolean;
  useBrowser: boolean;
  accountName?: string;
  certificateThumbprint?: string;
  hasToken?: boolean;
  kindLabel?: string;
}

export interface Connections {
  activeId?: string;
  connections: Connection[];
}

export interface AppInfo {
  name: string;
  version: string;
  builtInClientId?: string;
  dataDir: string;
  mcpCommand: string;
  startSite?: string;
}

export interface Identity {
  displayName: string;
  email?: string;
  loginName?: string;
  isApp: boolean;
  tenantTitle: string;
}

export interface SessionStatus {
  connectionId: string;
  identity?: Identity;
  accountName?: string;
}

export interface SiteSummary {
  url: string;
  title: string;
  template?: string;
  lastModified?: string;
  siteId?: string;
  /** For app catalog rows: an error SharePoint reports for the catalog. */
  error?: string;
}

export interface SitePage {
  sites: SiteSummary[];
  /** The cursor for the next page; absent at the end. */
  next?: string;
  /** Every match, when the source knows it. */
  total?: number;
  totalIsEstimate: boolean;
  source: "search" | "admin";
}

export type AppCatalogState = "Enabled" | "NotEnabled" | "Unknown";

export interface SiteDetails {
  url: string;
  title: string;
  webTemplate: string;
  siteId: string;
  appCatalog: AppCatalogState;
  appCatalogSource: "tenant" | "site";
  appCatalogUrl?: string;
  appCount?: number;
  requestedUrl?: string;
}

export interface SiteAppCatalog {
  url: string;
  siteId: string;
  error?: string;
}

export interface AppCatalogChange {
  siteUrl: string;
  changed: boolean;
  message: string;
}

export interface TokenDetails {
  audience?: string;
  expires?: string;
  subject?: string;
  tenantId?: string;
  permissions?: string;
  isForSharePoint: boolean;
  expired: boolean;
}

export interface CertificateResult {
  thumbprint: string;
  description: string;
  cerPath?: string;
}
