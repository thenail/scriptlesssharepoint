import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Button,
  Caption1,
  Checkbox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  Field,
  Input,
  makeStyles,
  MessageBar,
  MessageBarBody,
  Radio,
  RadioGroup,
  Spinner,
  Textarea,
  tokens,
} from "@fluentui/react-components";
import { CertificateRegular, DocumentArrowUpRegular } from "@fluentui/react-icons";
import { useEffect, useState, type ReactNode } from "react";

import { BridgeError } from "../bridge";
import { useApp } from "../context";
import type { Connection, TokenDetails } from "../types";

type Mode = "windows" | "user" | "certificate" | "token";

interface Form {
  name: string;
  tenantUrl: string;
  adminUrl: string;
  mode: Mode;
  userClientId: string;
  userTenantId: string;
  useBrowser: boolean;
  certClientId: string;
  certTenantId: string;
  thumbprint: string;
  token: string;
}

const useStyles = makeStyles({
  surface: { maxWidth: "640px", width: "100%" },
  content: { display: "flex", flexDirection: "column", gap: "14px", maxHeight: "68vh", overflowY: "auto", paddingRight: "8px" },
  group: { display: "flex", flexDirection: "column", gap: "12px", paddingLeft: "4px" },
  note: {
    padding: "10px 12px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
  },
  buttons: { display: "flex", gap: "8px", flexWrap: "wrap" },
  mono: { fontFamily: tokens.fontFamilyMonospace },
  dim: { color: tokens.colorNeutralForeground3, whiteSpace: "pre-line" },
  actions: { justifyContent: "space-between", width: "100%" },
  right: { display: "flex", gap: "8px" },
});

const blank = (s: string) => (s.trim() ? s.trim() : undefined);

function initial(c?: Connection): Form {
  return {
    name: c?.name ?? "",
    tenantUrl: c?.tenantUrl ?? "",
    adminUrl: c?.adminUrl ?? "",
    mode: !c ? "windows" : c.kind === "Certificate" ? "certificate" : c.kind === "AccessToken" ? "token" : c.useWindowsAccount ? "windows" : "user",
    userClientId: c?.kind === "User" ? c.clientId ?? "" : "",
    userTenantId: c?.kind === "User" ? c.tenantId ?? "" : "",
    useBrowser: c?.useBrowser ?? false,
    certClientId: c?.kind === "Certificate" ? c.clientId ?? "" : "",
    certTenantId: c?.kind === "Certificate" ? c.tenantId ?? "" : "",
    thumbprint: c?.certificateThumbprint ?? "",
    token: "",
  };
}

function toConnection(f: Form, id: string | undefined): Connection {
  const user = f.mode === "windows" || f.mode === "user";
  return {
    id,
    name: f.name.trim(),
    kind: user ? "User" : f.mode === "certificate" ? "Certificate" : "AccessToken",
    tenantUrl: f.tenantUrl.trim(),
    adminUrl: blank(f.adminUrl),
    clientId: blank(user ? f.userClientId : f.mode === "certificate" ? f.certClientId : ""),
    tenantId: blank(user ? f.userTenantId : f.mode === "certificate" ? f.certTenantId : ""),
    useWindowsAccount: f.mode === "windows",
    useBrowser: f.mode === "user" && f.useBrowser,
    certificateThumbprint: f.mode === "certificate" ? blank(f.thumbprint) : undefined,
  };
}

/** Waits until a value has stopped changing for a moment. */
function useSettled<T>(value: T, ms = 400): T {
  const [settled, setSettled] = useState(value);
  useEffect(() => {
    const t = setTimeout(() => setSettled(value), ms);
    return () => clearTimeout(t);
  }, [value, ms]);
  return settled;
}

interface Props {
  connection?: Connection;
  onClose(): void;
  onSaved(connection: Connection): void;
}

export function ConnectionDialog({ connection, onClose, onSaved }: Props) {
  const { api, info, notify } = useApp();
  const s = useStyles();
  const [form, setForm] = useState<Form>(() => initial(connection));
  const [id, setId] = useState(connection?.id);
  const [working, setWorking] = useState<"test" | "save" | null>(null);
  const [result, setResult] = useState<{ ok: boolean; text: string } | null>(null);
  const [account, setAccount] = useState(connection?.accountName);
  const [certInfo, setCertInfo] = useState<{ ok: boolean; description: string } | null>(null);
  const [tokenInfo, setTokenInfo] = useState<TokenDetails | null | "invalid">(null);
  const [pfx, setPfx] = useState<{ path: string; fileName: string } | null>(null);
  const [pfxPassword, setPfxPassword] = useState("");
  const set = <K extends keyof Form>(key: K, value: Form[K]) => setForm((f) => ({ ...f, [key]: value }));

  const thumbprint = useSettled(form.thumbprint);
  useEffect(() => {
    if (!thumbprint.trim()) {
      setCertInfo(null);
      return;
    }
    let live = true;
    void api.certificates.describe(thumbprint).then((r) => live && setCertInfo(r), () => undefined);
    return () => {
      live = false;
    };
  }, [api, thumbprint]);

  const token = useSettled(form.token);
  useEffect(() => {
    if (!token.trim()) {
      setTokenInfo(null);
      return;
    }
    let live = true;
    void api.tokens.inspect(token).then((r) => {
      if (!live) {
        return;
      }
      setTokenInfo(r ?? "invalid");
      if (r?.isForSharePoint && r.audience?.includes(".sharepoint.")) {
        setForm((f) => (f.tenantUrl ? f : { ...f, tenantUrl: r.audience! }));
      }
    }, () => undefined);
    return () => {
      live = false;
    };
  }, [api, token]);

  /** The tenant as typed → https://contoso.sharepoint.com, and a name from it. */
  const normalizeTenant = async () => {
    const r = await api.tenant.normalize(form.tenantUrl).catch(() => null);
    if (r) {
      setForm((f) => {
        const label = new URL(r.tenantUrl).hostname.split(".")[0];
        return { ...f, tenantUrl: r.tenantUrl, name: f.name || label[0].toUpperCase() + label.slice(1) };
      });
    }
  };

  const failed = (e: unknown) => setResult({ ok: false, text: e instanceof BridgeError || e instanceof Error ? e.message : String(e) });

  const test = async () => {
    setWorking("test");
    setResult(null);
    try {
      const r = await api.connections.test(toConnection(form, id), blank(form.token));
      setId(r.id);
      setAccount(r.accountName);
      const who = r.identity.isApp ? `an app (${r.identity.displayName})` : `${r.identity.displayName} (${r.identity.email ?? r.identity.loginName})`;
      setResult({ ok: true, text: `Connected to ${r.identity.tenantTitle} as ${who}.` });
    } catch (e) {
      failed(e);
    } finally {
      setWorking(null);
    }
  };

  const save = async () => {
    setWorking("save");
    setResult(null);
    try {
      const r = await api.connections.save(toConnection(form, id), blank(form.token));
      onSaved(r.connection);
    } catch (e) {
      failed(e);
      setWorking(null);
    }
  };

  const importPfx = async () => {
    try {
      const picked = await api.certificates.pickPfx();
      if (picked) {
        setPfxPassword("");
        setPfx(picked);
      }
    } catch (e) {
      failed(e);
    }
  };

  const finishImport = async () => {
    const picked = pfx!;
    setPfx(null);
    try {
      const r = await api.certificates.importPfx(picked.path, blank(pfxPassword));
      set("thumbprint", r.thumbprint);
      notify(`Imported ${picked.fileName}`);
    } catch (e) {
      failed(e);
    }
  };

  const createCertificate = async () => {
    try {
      const r = await api.certificates.create(form.name.trim() ? `Scriptless SharePoint ${form.name.trim()}` : "Scriptless SharePoint");
      if (r) {
        set("thumbprint", r.thumbprint);
        setResult({ ok: true, text: `Made a certificate valid for two years and saved its public key to ${r.cerPath}. Upload that file to the app registration (Certificates & secrets › Certificates). The private key stays in your certificate store.` });
      }
    } catch (e) {
      failed(e);
    }
  };

  const user = form.mode === "windows" || form.mode === "user";
  const builtIn = info?.builtInClientId;
  const tokenText = (t: TokenDetails): ReactNode =>
    [
      `For: ${t.audience ?? "?"}${t.isForSharePoint ? "" : "  ⚠ not a SharePoint token"}`,
      t.subject && `As: ${t.subject}`,
      t.expires && `${t.expired ? "Expired" : "Expires"} ${new Date(t.expires).toLocaleString()}`,
      t.permissions && `Permissions: ${t.permissions}`,
    ]
      .filter(Boolean)
      .join("\n");

  return (
    <>
      <Dialog open onOpenChange={(_, d) => !d.open && onClose()}>
        <DialogSurface className={s.surface}>
          <form
            onSubmit={(e) => {
              e.preventDefault();
              void save();
            }}
          >
            <DialogBody>
              <DialogTitle>{connection ? `Edit ${connection.name}` : "New connection"}</DialogTitle>
              <DialogContent className={s.content}>
                <Field label="SharePoint tenant" required hint="The tenant's name or any SharePoint URL in it, e.g. contoso or https://contoso.sharepoint.com.">
                  <Input value={form.tenantUrl} onChange={(_, d) => set("tenantUrl", d.value)} onBlur={() => void normalizeTenant()} autoFocus={!connection} />
                </Field>
                <Field label="Name" required>
                  <Input value={form.name} onChange={(_, d) => set("name", d.value)} />
                </Field>

                <Field label="Sign in with">
                  <RadioGroup value={form.mode} onChange={(_, d) => set("mode", d.value as Mode)}>
                    <Radio value="windows" label="The account signed in to Windows" />
                    <Radio value="user" label="Another account (choose it when signing in)" />
                    <Radio value="certificate" label="An app with a certificate (app-only, no user)" />
                    <Radio value="token" label="An access token" />
                  </RadioGroup>
                </Field>

                {user && (
                  <div className={s.group}>
                    <Field
                      label="App (client) ID"
                      required={!builtIn}
                      hint={builtIn ? `Leave it empty to use Scriptless SharePoint's own app (${builtIn}).` : "The ID of your organisation's Entra app registration for this app."}
                    >
                      <Input className={s.mono} value={form.userClientId} onChange={(_, d) => set("userClientId", d.value)} placeholder="00000000-0000-0000-0000-000000000000" />
                    </Field>
                    <Field label="Tenant ID or domain" hint="Only for single-tenant app registrations, e.g. contoso.onmicrosoft.com.">
                      <Input value={form.userTenantId} onChange={(_, d) => set("userTenantId", d.value)} />
                    </Field>
                    {form.mode === "user" && (
                      <Checkbox
                        checked={form.useBrowser}
                        onChange={(_, d) => set("useBrowser", !!d.checked)}
                        label="Sign in in the browser instead of with Windows"
                      />
                    )}
                    {account && <Caption1 className={s.dim}>Signed in as {account}.</Caption1>}
                    <Caption1 className={s.note}>
                      The app registration needs, under Authentication › Mobile and desktop applications, the redirect URI
                      ms-appx-web://microsoft.aad.brokerplugin/&lt;client ID&gt; (Windows sign-in) or http://localhost (browser sign-in), and under API
                      permissions SharePoint › Delegated › AllSites.FullControl, with admin consent. Enabling app catalogs also needs a SharePoint administrator.
                    </Caption1>
                  </div>
                )}

                {form.mode === "certificate" && (
                  <div className={s.group}>
                    <Field label="App (client) ID" required>
                      <Input className={s.mono} value={form.certClientId} onChange={(_, d) => set("certClientId", d.value)} placeholder="00000000-0000-0000-0000-000000000000" />
                    </Field>
                    <Field label="Tenant ID or domain" required hint="e.g. contoso.onmicrosoft.com or the tenant's GUID.">
                      <Input value={form.certTenantId} onChange={(_, d) => set("certTenantId", d.value)} />
                    </Field>
                    <Field
                      label="Certificate thumbprint"
                      required
                      validationState={certInfo ? (certInfo.ok ? "success" : "error") : "none"}
                      validationMessage={certInfo?.description}
                    >
                      <Input className={s.mono} value={form.thumbprint} onChange={(_, d) => set("thumbprint", d.value)} />
                    </Field>
                    <div className={s.buttons}>
                      <Button icon={<DocumentArrowUpRegular />} onClick={() => void importPfx()}>
                        Import .pfx…
                      </Button>
                      <Button icon={<CertificateRegular />} onClick={() => void createCertificate()}>
                        Create a certificate…
                      </Button>
                    </div>
                    <Caption1 className={s.note}>
                      The app registration needs, under API permissions, SharePoint › Application › Sites.FullControl.All with admin consent, and this
                      certificate's .cer under Certificates &amp; secrets. SharePoint doesn't accept client secrets for app-only access, only certificates.
                    </Caption1>
                  </div>
                )}

                {form.mode === "token" && (
                  <div className={s.group}>
                    <Field
                      label="Access token"
                      required={!connection?.hasToken}
                      hint={connection?.hasToken && !form.token ? "A token is saved. Paste a new one to replace it." : undefined}
                      validationState={tokenInfo === "invalid" || (tokenInfo && (!tokenInfo.isForSharePoint || tokenInfo.expired)) ? "warning" : "none"}
                      validationMessage={tokenInfo === "invalid" ? "That doesn't look like an access token (a JWT: three parts separated by dots)." : undefined}
                    >
                      <Textarea className={s.mono} value={form.token} onChange={(_, d) => set("token", d.value)} rows={4} resize="vertical" />
                    </Field>
                    {tokenInfo && tokenInfo !== "invalid" && <Caption1 className={s.dim}>{tokenText(tokenInfo)}</Caption1>}
                    <Caption1 className={s.note}>
                      A token for https://&lt;tenant&gt;.sharepoint.com (not Microsoft Graph). Tokens can't be refreshed and last about an hour. It's kept
                      encrypted for your Windows account.
                    </Caption1>
                  </div>
                )}

                <Accordion collapsible defaultOpenItems={form.adminUrl ? ["advanced"] : []}>
                  <AccordionItem value="advanced">
                    <AccordionHeader size="small">Advanced</AccordionHeader>
                    <AccordionPanel>
                      <Field label="Admin centre URL" hint="Leave it empty for https://<tenant>-admin.sharepoint.com.">
                        <Input value={form.adminUrl} onChange={(_, d) => set("adminUrl", d.value)} />
                      </Field>
                    </AccordionPanel>
                  </AccordionItem>
                </Accordion>

                {result && (
                  <MessageBar intent={result.ok ? "success" : "error"}>
                    <MessageBarBody>{result.text}</MessageBarBody>
                  </MessageBar>
                )}
              </DialogContent>
              <DialogActions fluid className={s.actions}>
                <Button onClick={() => void test()} disabled={!!working} icon={working === "test" ? <Spinner size="tiny" /> : undefined}>
                  Test connection
                </Button>
                <div className={s.right}>
                  <DialogTrigger disableButtonEnhancement>
                    <Button appearance="secondary">Cancel</Button>
                  </DialogTrigger>
                  <Button appearance="primary" type="submit" disabled={!!working}>
                    Save
                  </Button>
                </div>
              </DialogActions>
            </DialogBody>
          </form>
        </DialogSurface>
      </Dialog>

      {pfx && (
        <Dialog open onOpenChange={(_, d) => !d.open && setPfx(null)}>
          <DialogSurface>
            <form
              onSubmit={(e) => {
                e.preventDefault();
                void finishImport();
              }}
            >
              <DialogBody>
                <DialogTitle>Certificate password</DialogTitle>
                <DialogContent>
                  <Field label={`The password of ${pfx.fileName}`} hint="Leave it empty if it has none.">
                    <Input type="password" value={pfxPassword} onChange={(_, d) => setPfxPassword(d.value)} autoFocus />
                  </Field>
                </DialogContent>
                <DialogActions>
                  <DialogTrigger disableButtonEnhancement>
                    <Button appearance="secondary">Cancel</Button>
                  </DialogTrigger>
                  <Button appearance="primary" type="submit">
                    Import
                  </Button>
                </DialogActions>
              </DialogBody>
            </form>
          </DialogSurface>
        </Dialog>
      )}
    </>
  );
}
