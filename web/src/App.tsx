import {
  Button,
  Caption1,
  Dropdown,
  FluentProvider,
  makeStyles,
  Menu,
  MenuDivider,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  Option,
  Persona,
  ProgressBar,
  Spinner,
  Subtitle1,
  Toast,
  Toaster,
  ToastTitle,
  tokens,
  Tooltip,
  useId,
  useToastController,
} from "@fluentui/react-components";
import {
  AddRegular,
  BotSparkleRegular,
  DeleteRegular,
  DismissRegular,
  EditRegular,
  InfoRegular,
  MoreHorizontalRegular,
  PersonRegular,
  SignOutRegular,
} from "@fluentui/react-icons";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";

import { BridgeError, type Api } from "./bridge";
import { AboutDialog } from "./components/AboutDialog";
import { ConfirmDialog, type ConfirmRequest } from "./components/ConfirmDialog";
import { ConnectAiDialog } from "./components/ConnectAiDialog";
import { ConnectionDialog } from "./components/ConnectionDialog";
import { SitePane } from "./components/SitePane";
import { SitesPane, type SiteResults } from "./components/SitesPane";
import { Welcome } from "./components/Welcome";
import { AppContext, useSystemTheme, type AppContextValue } from "./context";
import type { AppInfo, Connection, Connections, SessionStatus, SiteDetails } from "./types";

const useStyles = makeStyles({
  shell: {
    height: "100vh",
    display: "flex",
    flexDirection: "column",
    backgroundColor: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground1,
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "12px",
    padding: "10px 16px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    minHeight: "40px",
  },
  brand: { display: "flex", alignItems: "center", gap: "10px", flexGrow: 1, minWidth: 0 },
  logo: { width: "28px", height: "28px" },
  account: { minWidth: 0, maxWidth: "260px" },
  connection: { minWidth: "220px", maxWidth: "280px" },
  progress: { height: "2px" },
  progressSpace: { height: "2px", flexShrink: 0 },
  body: {
    flexGrow: 1,
    minHeight: 0,
    display: "grid",
    gridTemplateColumns: "minmax(420px, 3fr) minmax(320px, 2fr)",
    gap: "16px",
    padding: "16px",
  },
  error: { margin: "12px 16px 0" },
});

type Dialog =
  | { type: "connection"; connection?: Connection }
  | { type: "connectAi" }
  | { type: "about" }
  | { type: "confirm"; request: ConfirmRequest }
  | null;

export function App({ api }: { api: Api }) {
  const theme = useSystemTheme();
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);
  const s = useStyles();

  const [info, setInfo] = useState<AppInfo | null>(null);
  const [conns, setConns] = useState<Connections | null>(null);
  const [status, setStatus] = useState<SessionStatus | null>(null);
  const [checking, setChecking] = useState(false);
  const [busy, setBusy] = useState(0);
  const [error, setError] = useState<{ message: string; signIn: boolean } | null>(null);
  const [results, setResults] = useState<SiteResults | null>(null);
  const [selectedUrl, setSelectedUrl] = useState<string | null>(null);
  const [site, setSite] = useState<SiteDetails | null>(null);
  const [loadingSite, setLoadingSite] = useState(false);
  const [dialog, setDialog] = useState<Dialog>(null);
  const siteRequest = useRef(0);

  const active = conns?.connections.find((c) => c.id === conns.activeId) ?? conns?.connections[0];

  const notify = useCallback<AppContextValue["notify"]>(
    (message, intent = "success") =>
      dispatchToast(
        <Toast>
          <ToastTitle>{message}</ToastTitle>
        </Toast>,
        { intent, timeout: 5000 },
      ),
    [dispatchToast],
  );

  /** Shows an error at the top, with Sign in when signing in would fix it. */
  const report = useCallback((e: unknown) => {
    setError({
      message: e instanceof Error ? e.message : String(e),
      signIn: e instanceof BridgeError && e.signInRequired,
    });
  }, []);

  /** Runs an operation, counting it as busy and showing its error. */
  const run = useCallback(
    async <T,>(work: () => Promise<T>): Promise<T | undefined> => {
      setBusy((b) => b + 1);
      setError(null);
      try {
        return await work();
      } catch (e) {
        report(e);
        return undefined;
      } finally {
        setBusy((b) => b - 1);
      }
    },
    [report],
  );

  /** Silently finds out who the active connection is signed in as. */
  const checkStatus = useCallback(async () => {
    setChecking(true);
    try {
      setStatus(await api.session.status());
    } catch {
      setStatus(null);
    } finally {
      setChecking(false);
    }
  }, [api]);

  const resetWork = () => {
    setResults(null);
    setSelectedUrl(null);
    setSite(null);
    setError(null);
  };

  const loadConnections = useCallback(async () => {
    const list = await api.connections.list();
    setConns(list);
    if (list.connections.length) {
      await checkStatus();
    } else {
      setStatus(null);
    }
    return list;
  }, [api, checkStatus]);

  const showSite = useCallback(
    async (url: string) => {
      const request = ++siteRequest.current;
      setSelectedUrl(url);
      setLoadingSite(true);
      const details = await run(() => api.sites.get(url));
      if (request === siteRequest.current) {
        setSite(details ?? null);
        setLoadingSite(false);
      }
    },
    [api, run],
  );

  // start: the app's info, the connections, who's signed in, and a site to show
  useEffect(() => {
    void (async () => {
      const i = await api.info();
      setInfo(i);
      const list = await loadConnections();
      if (i.startSite && list.connections.length) {
        void showSite(i.startSite);
      }
    })();
    return api.onAccountChanged(() => {
      void api.connections.list().then(setConns);
    });
  }, [api, loadConnections, showSite]);

  // an operation that signed in shows the account in the header
  const afterSignIn = useCallback(() => {
    if (!status?.identity) {
      void checkStatus();
    }
  }, [status, checkStatus]);

  const signIn = () =>
    run(async () => {
      const r = await api.session.signIn();
      setStatus((st) => ({ connectionId: active?.id ?? "", ...st, identity: r.identity, accountName: r.accountName }));
      notify(`Signed in to ${r.identity.tenantTitle}`);
    });

  const signOut = () =>
    run(async () => {
      await api.session.signOut();
      setStatus((st) => (st ? { ...st, identity: undefined } : st));
      notify("Signed out", "info");
    });

  const switchTo = (id: string) =>
    run(async () => {
      await api.connections.setActive(id);
      resetWork();
      setStatus(null);
      setConns((c) => (c ? { ...c, activeId: id } : c));
      await checkStatus();
    });

  const deleteActive = () => {
    if (!active?.id) {
      return;
    }
    setDialog({
      type: "confirm",
      request: {
        title: `Delete “${active.name}”?`,
        body: "The connection and its sign-in are removed from this PC. AI apps can't use it any more either.",
        confirm: "Delete",
        danger: true,
        onConfirm: () =>
          run(async () => {
            const list = await api.connections.remove(active.id!);
            setConns(list);
            resetWork();
            setStatus(null);
            if (list.connections.length) {
              await checkStatus();
            }
          }),
      },
    });
  };

  const onSaved = async (saved: Connection) => {
    setDialog(null);
    resetWork();
    await loadConnections();
    notify(`Saved “${saved.name}”`);
  };

  const context = useMemo<AppContextValue>(() => ({ api, info, notify }), [api, info, notify]);
  const identity = status?.identity;
  const hasConnections = !!conns?.connections.length;

  return (
    <FluentProvider theme={theme}>
      <AppContext.Provider value={context}>
        <div className={s.shell}>
          <header className={s.header}>
            <div className={s.brand}>
              <img src="./logo.png" alt="" className={s.logo} />
              <Subtitle1 as="h1" style={{ margin: 0 }}>
                Scriptless SharePoint
              </Subtitle1>
            </div>
            {hasConnections && (
              <>
                <div className={s.account}>
                  {checking ? (
                    <Spinner size="tiny" label="Checking sign-in…" />
                  ) : identity ? (
                    <Tooltip content={`${identity.displayName}${identity.loginName ? `\n${identity.loginName}` : ""}`} relationship="description">
                      <Persona
                        size="small"
                        name={identity.isApp ? "App-only" : identity.displayName}
                        secondaryText={identity.tenantTitle}
                        avatar={identity.isApp ? { icon: <BotSparkleRegular /> } : undefined}
                      />
                    </Tooltip>
                  ) : (
                    <Caption1>{status?.accountName ? `${status.accountName} · not signed in` : "Not signed in"}</Caption1>
                  )}
                </div>
                <Dropdown
                  className={s.connection}
                  aria-label="Connection"
                  value={active?.name ?? ""}
                  selectedOptions={active?.id ? [active.id] : []}
                  onOptionSelect={(_, d) => d.optionValue && d.optionValue !== active?.id && void switchTo(d.optionValue)}
                  disabled={busy > 0}
                >
                  {conns!.connections.map((c) => (
                    <Option key={c.id} value={c.id!} text={c.name}>
                      <div>
                        <div>{c.name}</div>
                        <Caption1>
                          {c.kindLabel} · {c.tenantUrl.replace("https://", "")}
                        </Caption1>
                      </div>
                    </Option>
                  ))}
                </Dropdown>
                {!identity && !checking && (
                  <Button appearance="primary" icon={<PersonRegular />} onClick={() => void signIn()} disabled={busy > 0}>
                    Sign in
                  </Button>
                )}
              </>
            )}
            <Menu positioning="below-end">
              <MenuTrigger disableButtonEnhancement>
                <Tooltip content="More" relationship="label">
                  <Button appearance="subtle" icon={<MoreHorizontalRegular />} />
                </Tooltip>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem icon={<AddRegular />} onClick={() => setDialog({ type: "connection" })}>
                    New connection…
                  </MenuItem>
                  <MenuItem icon={<EditRegular />} disabled={!active} onClick={() => setDialog({ type: "connection", connection: active })}>
                    Edit connection…
                  </MenuItem>
                  <MenuItem icon={<DeleteRegular />} disabled={!active} onClick={deleteActive}>
                    Delete connection
                  </MenuItem>
                  <MenuItem icon={<SignOutRegular />} disabled={!identity || active?.kind !== "User"} onClick={() => void signOut()}>
                    Sign out
                  </MenuItem>
                  <MenuDivider />
                  <MenuItem icon={<BotSparkleRegular />} onClick={() => setDialog({ type: "connectAi" })}>
                    Connect an AI app…
                  </MenuItem>
                  <MenuItem icon={<InfoRegular />} onClick={() => setDialog({ type: "about" })}>
                    About Scriptless SharePoint
                  </MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </header>
          {busy > 0 ? <ProgressBar className={s.progress} thickness="medium" /> : <div className={s.progressSpace} />}

          {error && (
            <MessageBar intent="error" className={s.error}>
              <MessageBarBody>{error.message}</MessageBarBody>
              <MessageBarActions
                containerAction={<Button appearance="transparent" icon={<DismissRegular />} aria-label="Dismiss" onClick={() => setError(null)} />}
              >
                {error.signIn && <Button onClick={() => void signIn()}>Sign in</Button>}
              </MessageBarActions>
            </MessageBar>
          )}

          {conns && !hasConnections ? (
            <Welcome onAdd={() => setDialog({ type: "connection" })} />
          ) : (
            <main className={s.body}>
              <SitesPane
                key={active?.id}
                results={results}
                selectedUrl={selectedUrl}
                signedIn={!!identity}
                onResults={setResults}
                onSelect={(url) => void showSite(url)}
                onError={report}
                afterSignIn={afterSignIn}
              />
              <SitePane
                site={site}
                loading={loadingSite}
                busy={busy > 0}
                run={run}
                onChanged={(details) => {
                  setSite(details);
                  afterSignIn();
                }}
                onRefresh={() => site && void showSite(site.url)}
                confirm={(request) => setDialog({ type: "confirm", request })}
              />
            </main>
          )}
        </div>

        {dialog?.type === "connection" && (
          <ConnectionDialog connection={dialog.connection} onClose={() => setDialog(null)} onSaved={onSaved} />
        )}
        {dialog?.type === "connectAi" && <ConnectAiDialog onClose={() => setDialog(null)} />}
        {dialog?.type === "about" && <AboutDialog onClose={() => setDialog(null)} />}
        {dialog?.type === "confirm" && <ConfirmDialog request={dialog.request} onClose={() => setDialog(null)} />}
        <Toaster toasterId={toasterId} position="bottom-end" />
      </AppContext.Provider>
    </FluentProvider>
  );
}
