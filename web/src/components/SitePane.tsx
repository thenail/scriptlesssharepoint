import {
  Badge,
  Body1,
  Button,
  Caption1,
  Card,
  CardHeader,
  Link,
  makeStyles,
  MessageBar,
  MessageBarBody,
  Spinner,
  Subtitle2,
  Text,
  Title3,
  tokens,
} from "@fluentui/react-components";
import { AppFolderRegular, ArrowClockwiseRegular, OpenRegular } from "@fluentui/react-icons";

import { useApp } from "../context";
import type { SiteDetails } from "../types";
import type { ConfirmRequest } from "./ConfirmDialog";

const useStyles = makeStyles({
  pane: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    minHeight: 0,
    minWidth: 0,
    overflowY: "auto",
  },
  placeholder: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: "8px",
    height: "100%",
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
    padding: "24px",
  },
  icon: { fontSize: "48px" },
  head: { display: "flex", flexDirection: "column", gap: "4px", padding: "4px 4px 0" },
  url: { wordBreak: "break-all" },
  meta: { color: tokens.colorNeutralForeground3 },
  status: { display: "flex", alignItems: "center", gap: "8px" },
  actions: { display: "flex", flexWrap: "wrap", gap: "8px" },
  help: { color: tokens.colorNeutralForeground3, padding: "0 4px" },
});

interface Props {
  site: SiteDetails | null;
  loading: boolean;
  busy: boolean;
  run<T>(work: () => Promise<T>): Promise<T | undefined>;
  onChanged(site: SiteDetails): void;
  onRefresh(): void;
  confirm(request: ConfirmRequest): void;
}

export function SitePane({ site, loading, busy, run, onChanged, onRefresh, confirm }: Props) {
  const { api, notify } = useApp();
  const s = useStyles();

  if (loading && !site) {
    return (
      <section className={s.pane}>
        <div className={s.placeholder}>
          <Spinner label="Reading the site…" />
        </div>
      </section>
    );
  }
  if (!site) {
    return (
      <section className={s.pane}>
        <div className={s.placeholder}>
          <AppFolderRegular className={s.icon} />
          <Subtitle2>Site collection app catalogs</Subtitle2>
          <Body1>Choose a site to see whether it has its own app catalog, and to enable or remove one.</Body1>
        </div>
      </section>
    );
  }

  const enabled = site.appCatalog === "Enabled";

  const change = (enable: boolean) =>
    run(async () => {
      const result = enable ? await api.catalogs.enable(site.url) : await api.catalogs.disable(site.url);
      onChanged(await api.sites.get(result.siteUrl));
      notify(result.message, result.changed ? "success" : "info");
    });

  const askEnable = () =>
    confirm({
      title: "Enable a site collection app catalog?",
      body: (
        <>
          <Text block>{site.url}</Text>
          <Text block style={{ marginTop: 8 }}>
            The site's owners can then add SharePoint Framework solutions to this site only. SharePoint adds an Apps for SharePoint library within a few minutes.
          </Text>
        </>
      ),
      confirm: "Enable",
      onConfirm: () => change(true),
    });

  const askRemove = () =>
    confirm({
      title: "Remove the site collection app catalog?",
      body: (
        <>
          <Text block>{site.url}</Text>
          <Text block style={{ marginTop: 8 }}>
            Its Apps for SharePoint library and files stay in the site, but it stops being an app catalog, and solutions added from it may stop working on the site.
          </Text>
        </>
      ),
      confirm: "Remove",
      danger: true,
      onConfirm: () => change(false),
    });

  return (
    <section className={s.pane} aria-label="Site" aria-busy={loading}>
      <div className={s.head}>
        <Title3 as="h2">{site.title}</Title3>
        <Link className={s.url} onClick={() => void api.openUrl(site.url)}>
          {site.url}
        </Link>
        <Caption1 className={s.meta}>
          Template {site.webTemplate} · Site ID {site.siteId}
          {site.requestedUrl && (
            <>
              <br />
              Showing the site collection of {site.requestedUrl}
            </>
          )}
        </Caption1>
      </div>

      <Card>
        <CardHeader
          image={<AppFolderRegular fontSize={24} />}
          header={<Subtitle2>Site collection app catalog</Subtitle2>}
          action={loading ? <Spinner size="tiny" /> : undefined}
        />
        <div className={s.status}>
          <Badge appearance="filled" color={enabled ? "success" : "informative"} shape="rounded">
            {enabled ? "Enabled" : "Not enabled"}
          </Badge>
          <Body1>
            {enabled
              ? site.appCount !== undefined && site.appCount !== null
                ? `Its Apps for SharePoint library has ${site.appCount} item${site.appCount === 1 ? "" : "s"}.`
                : "SharePoint may take a few minutes to add its Apps for SharePoint library."
              : "The site uses the tenant app catalog only."}
          </Body1>
        </div>
        {site.appCatalogSource === "site" && (
          <MessageBar intent="info">
            <MessageBarBody>
              Judged from the site's libraries: this connection can't read the tenant's list of app catalogs, which needs a SharePoint administrator. Enabling or removing one needs an administrator too.
            </MessageBarBody>
          </MessageBar>
        )}
        <div className={s.actions}>
          {enabled ? (
            <>
              {site.appCatalogUrl && (
                <Button icon={<OpenRegular />} onClick={() => void api.openUrl(site.appCatalogUrl!)}>
                  Open app catalog
                </Button>
              )}
              <Button onClick={askRemove} disabled={busy}>
                Remove app catalog
              </Button>
            </>
          ) : (
            <Button appearance="primary" onClick={askEnable} disabled={busy}>
              Enable app catalog
            </Button>
          )}
        </div>
      </Card>

      <div className={s.actions}>
        <Button icon={<OpenRegular />} onClick={() => void api.openUrl(site.url)}>
          Open site
        </Button>
        <Button icon={<ArrowClockwiseRegular />} onClick={onRefresh} disabled={busy}>
          Refresh
        </Button>
      </div>

      <Caption1 className={s.help}>
        A site collection app catalog lets a site's owners add SharePoint Framework (SPFx) solutions to that site only, without the tenant app catalog.
      </Caption1>
    </section>
  );
}
