import {
  Body1,
  Button,
  Caption1,
  makeStyles,
  MessageBar,
  MessageBarBody,
  SearchBox,
  Spinner,
  Tab,
  Table,
  TableBody,
  TableCell,
  TableCellLayout,
  TableHeader,
  TableHeaderCell,
  TableRow,
  TabList,
  tokens,
  useArrowNavigationGroup,
} from "@fluentui/react-components";
import { AppFolderRegular, ArrowClockwiseRegular, BuildingRegular, ErrorCircleRegular, GlobeRegular, SearchRegular } from "@fluentui/react-icons";
import { useCallback, useEffect, useRef, useState, type UIEvent } from "react";

import { BridgeError } from "../bridge";
import { useApp } from "../context";
import { sitePath } from "../snippets";
import type { SiteSummary } from "../types";

export type Scope = "search" | "admin" | "catalogs";

/** What the list shows: one scope's query, the pages loaded so far, and where the next starts. */
export interface SiteResults {
  scope: Scope;
  query: string;
  rows: SiteSummary[];
  next?: string;
  total?: number;
  totalIsEstimate?: boolean;
}

const PAGE = 50;
/** How long typing pauses before the search goes to SharePoint. */
const TYPING_PAUSE_MS = 350;
/** How close to the bottom (px) scrolling loads the next page. */
const NEAR_BOTTOM = 240;

const useStyles = makeStyles({
  pane: {
    display: "flex",
    flexDirection: "column",
    minHeight: 0,
    minWidth: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusLarge,
    boxShadow: tokens.shadow4,
    padding: "12px 16px 12px",
    gap: "10px",
  },
  hint: { color: tokens.colorNeutralForeground3 },
  searchRow: { display: "flex", gap: "8px", alignItems: "center" },
  search: { flexGrow: 1, maxWidth: "unset" },
  tableWrap: { flexGrow: 1, minHeight: 0, overflowY: "auto" },
  table: { tableLayout: "fixed" },
  row: { cursor: "pointer" },
  empty: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: "8px",
    padding: "48px 16px",
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
  },
  emptyIcon: { fontSize: "40px" },
  more: { display: "flex", justifyContent: "center", padding: "12px" },
  footer: { display: "flex", alignItems: "center", justifyContent: "space-between", minHeight: "24px" },
  url: { color: tokens.colorNeutralForeground3 },
});

const HINTS: Record<Scope, string> = {
  search: "SharePoint search: the sites this account can see, by words starting with what you type. New sites can take a few minutes to show up.",
  admin: "Every site in the tenant, from the SharePoint admin centre: titles containing what you type, then URLs. Needs a SharePoint administrator.",
  catalogs: "The site collections that have their own app catalog. Needs a SharePoint administrator.",
};

interface Props {
  results: SiteResults | null;
  selectedUrl: string | null;
  /** Whether the connection is signed in: until then nothing loads by itself (it would prompt). */
  signedIn: boolean;
  onResults(results: SiteResults | null): void;
  onSelect(url: string): void;
  /** Shows an error the pane can't, e.g. one that needs signing in. */
  onError(error: unknown): void;
  afterSignIn(): void;
}

export function SitesPane({ results, selectedUrl, signedIn, onResults, onSelect, onError, afterSignIn }: Props) {
  const { api } = useApp();
  const s = useStyles();
  const [scope, setScope] = useState<Scope>(results?.scope ?? "search");
  const [query, setQuery] = useState(results?.query ?? "");
  const [loading, setLoading] = useState<"first" | "more" | null>(null);
  const [problem, setProblem] = useState<string | null>(null);
  const navigation = useArrowNavigationGroup({ axis: "vertical" });
  const wrap = useRef<HTMLDivElement>(null);
  /** Each load's number: answers to older ones (the user kept typing) are dropped. */
  const request = useRef(0);
  const current = useRef(results);
  current.current = results;

  const fail = useCallback(
    (e: unknown) => {
      if (e instanceof BridgeError && !e.signInRequired) {
        setProblem(e.message);
      } else {
        onError(e);
      }
    },
    [onError],
  );

  /** The first page for a scope and query (replacing the list), or the next page of the list. */
  const load = useCallback(
    async (target: Scope, q: string, more: boolean) => {
      const prev = current.current;
      if (more && (!prev?.next || prev.scope !== target)) {
        return;
      }
      const id = ++request.current;
      setLoading(more ? "more" : "first");
      setProblem(null);
      try {
        if (target === "catalogs") {
          const list = await api.catalogs.list();
          if (id === request.current) {
            onResults({
              scope: "catalogs",
              query: "",
              rows: list.map((c) => ({ url: c.url, title: sitePath(c.url), siteId: c.siteId, error: c.error, template: "App catalog" })),
              total: list.length,
            });
          }
        } else {
          const page = await api.sites.search({ query: q || undefined, scope: target, cursor: more ? prev!.next : undefined, limit: PAGE });
          if (id === request.current) {
            onResults({
              scope: target,
              query: q,
              rows: more ? [...prev!.rows, ...page.sites.filter((x) => !prev!.rows.some((r) => r.url === x.url))] : page.sites,
              next: page.next,
              total: page.total,
              totalIsEstimate: page.totalIsEstimate,
            });
          }
        }
        afterSignIn();
      } catch (e) {
        if (id === request.current) {
          fail(e);
          if (!more) {
            onResults({ scope: target, query: q, rows: [] });
          }
        }
      } finally {
        if (id === request.current) {
          setLoading(null);
        }
      }
    },
    [api, onResults, afterSignIn, fail],
  );

  // typing: search the source once typing pauses (and straight away on a new tab or sign-in);
  // only the query, tab and sign-in start a search, not a new `load` from a re-render
  const loadRef = useRef(load);
  loadRef.current = load;
  const first = useRef(true);
  useEffect(() => {
    if (!signedIn) {
      return;
    }
    const q = query.trim();
    if (current.current?.scope === scope && current.current.query === q && current.current.rows.length) {
      return; // already showing it (e.g. back from another connection's sign-in)
    }
    const immediate = first.current || scope === "catalogs";
    first.current = false;
    const t = setTimeout(() => void loadRef.current(scope, q, false), immediate ? 0 : TYPING_PAUSE_MS);
    return () => clearTimeout(t);
  }, [query, scope, signedIn]);

  // infinite scroll: the next page when scrolled near the bottom, or when the list doesn't fill the pane
  const loadMoreIfNear = useCallback(() => {
    const el = wrap.current;
    const r = current.current;
    if (!el || !r?.next || loading || r.scope === "catalogs") {
      return;
    }
    if (el.scrollTop + el.clientHeight >= el.scrollHeight - NEAR_BOTTOM) {
      void load(r.scope, r.query, true);
    }
  }, [load, loading]);
  useEffect(() => {
    loadMoreIfNear();
  }, [results, loadMoreIfNear]);

  const changeScope = (next: Scope) => {
    setScope(next);
    setProblem(null);
    request.current++; // drop answers for the old tab
    setLoading(null);
    onResults(null);
    if (wrap.current) {
      wrap.current.scrollTop = 0;
    }
  };

  const rows = results?.scope === scope ? results.rows : null;
  const count = rows?.length ?? 0;
  const totalText =
    results?.total !== undefined && results.total !== null && scope !== "catalogs"
      ? ` of ${results.totalIsEstimate ? "about " : ""}${results.total.toLocaleString()}`
      : "";

  return (
    <section className={s.pane} aria-label="Sites">
      <TabList selectedValue={scope} onTabSelect={(_, d) => changeScope(d.value as Scope)} size="small">
        <Tab value="search" icon={<SearchRegular />}>
          Search
        </Tab>
        <Tab value="admin" icon={<BuildingRegular />}>
          All sites
        </Tab>
        <Tab value="catalogs" icon={<AppFolderRegular />}>
          Site app catalogs
        </Tab>
      </TabList>
      <Caption1 className={s.hint}>{HINTS[scope]}</Caption1>

      <div className={s.searchRow}>
        {scope === "catalogs" ? (
          <Button icon={<ArrowClockwiseRegular />} onClick={() => void load("catalogs", "", false)} disabled={!!loading}>
            Refresh
          </Button>
        ) : (
          <SearchBox
            className={s.search}
            placeholder={scope === "admin" ? "Search titles and URLs; empty for all sites" : "Search sites; empty for all sites"}
            aria-label="Search sites"
            value={query}
            onChange={(_, d) => setQuery(d.value)}
            onKeyDown={(e) => e.key === "Enter" && void load(scope, query.trim(), false)}
          />
        )}
        {loading === "first" && <Spinner size="tiny" aria-label="Loading" />}
      </div>

      {problem && (
        <MessageBar intent="warning">
          <MessageBarBody>{problem}</MessageBarBody>
        </MessageBar>
      )}

      <div className={s.tableWrap} ref={wrap} onScroll={(_: UIEvent) => loadMoreIfNear()}>
        {!signedIn && rows === null ? (
          <div className={s.empty}>
            <GlobeRegular className={s.emptyIcon} />
            <Body1>Sign in to see the sites.</Body1>
          </div>
        ) : rows === null ? (
          <div className={s.empty}>
            <Spinner label={scope === "catalogs" ? "Loading the site collection app catalogs…" : "Loading sites…"} />
          </div>
        ) : rows.length === 0 && !loading ? (
          <div className={s.empty}>
            <SearchRegular className={s.emptyIcon} />
            <Body1>{scope === "catalogs" ? "No site collection has its own app catalog yet." : "No sites found."}</Body1>
          </div>
        ) : (
          <>
            <Table className={s.table} size="small" aria-label="Sites" {...navigation}>
              <TableHeader>
                <TableRow>
                  <TableHeaderCell style={{ width: "38%" }}>Title</TableHeaderCell>
                  <TableHeaderCell>URL</TableHeaderCell>
                  <TableHeaderCell style={{ width: "18%" }}>{scope === "catalogs" ? "Status" : "Template"}</TableHeaderCell>
                  {scope !== "catalogs" && <TableHeaderCell style={{ width: "14%" }}>{scope === "admin" ? "Activity" : "Modified"}</TableHeaderCell>}
                </TableRow>
              </TableHeader>
              <TableBody>
                {rows.map((site) => (
                  <TableRow
                    key={site.url}
                    className={s.row}
                    tabIndex={0}
                    appearance={site.url === selectedUrl ? "brand" : "none"}
                    aria-selected={site.url === selectedUrl}
                    onClick={() => onSelect(site.url)}
                    onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && (e.preventDefault(), onSelect(site.url))}
                  >
                    <TableCell>
                      <TableCellLayout truncate title={site.title}>
                        {site.title}
                      </TableCellLayout>
                    </TableCell>
                    <TableCell>
                      <TableCellLayout truncate title={site.url} className={s.url}>
                        {sitePath(site.url)}
                      </TableCellLayout>
                    </TableCell>
                    <TableCell>
                      <TableCellLayout truncate media={site.error ? <ErrorCircleRegular color={tokens.colorPaletteRedForeground1} /> : undefined} title={site.error}>
                        {site.error ? "Error" : site.template}
                      </TableCellLayout>
                    </TableCell>
                    {scope !== "catalogs" && (
                      <TableCell>
                        <TableCellLayout truncate>{site.lastModified ? new Date(site.lastModified).toLocaleDateString() : ""}</TableCellLayout>
                      </TableCell>
                    )}
                  </TableRow>
                ))}
              </TableBody>
            </Table>
            {results?.next && (
              <div className={s.more}>
                {loading === "more" ? (
                  <Spinner size="tiny" label="Loading more…" />
                ) : (
                  <Button appearance="subtle" onClick={() => void load(scope, results.query, true)}>
                    Load more
                  </Button>
                )}
              </div>
            )}
          </>
        )}
      </div>

      <div className={s.footer}>
        <Caption1 className={s.hint}>
          {count > 0 && (scope === "catalogs" ? `${count} app catalog${count === 1 ? "" : "s"}` : `${count.toLocaleString()}${totalText} sites${results?.next ? " · scroll for more" : ""}`)}
        </Caption1>
      </div>
    </section>
  );
}
