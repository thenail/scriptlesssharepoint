import {
  Body1,
  Button,
  Caption1,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  makeStyles,
  Subtitle2,
  tokens,
} from "@fluentui/react-components";
import { CheckmarkRegular, CopyRegular } from "@fluentui/react-icons";
import { useState } from "react";

import { useApp } from "../context";
import { connectSnippets } from "../snippets";

const useStyles = makeStyles({
  surface: { maxWidth: "680px" },
  content: { display: "flex", flexDirection: "column", gap: "18px" },
  section: { display: "flex", flexDirection: "column", gap: "4px" },
  dim: { color: tokens.colorNeutralForeground3 },
  row: { display: "flex", gap: "8px", alignItems: "flex-start" },
  code: {
    flexGrow: 1,
    minWidth: 0,
    margin: 0,
    padding: "8px 10px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    whiteSpace: "pre-wrap",
    overflowWrap: "anywhere",
    userSelect: "text",
  },
});

/**
 * How to add the MCP server that ships with the app to Claude Desktop, Claude
 * Code and VS Code. Windows' own agents find it without setup.
 */
export function ConnectAiDialog({ onClose }: { onClose(): void }) {
  const { info } = useApp();
  const s = useStyles();
  const [copied, setCopied] = useState<string | null>(null);
  const command = info?.mcpCommand ?? "scriptless-sharepoint-mcp";

  const copy = async (key: string, text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopied(key);
    } catch {
      // the text can still be selected and copied by hand
    }
  };

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className={s.surface}>
        <DialogBody>
          <DialogTitle>Connect an AI app</DialogTitle>
          <DialogContent className={s.content}>
            <Body1>
              Scriptless SharePoint comes with an MCP server that lets AI apps search your sites and enable site collection app catalogs, with the connections you set up here.
            </Body1>
            {connectSnippets(command).map((snippet) => (
              <div key={snippet.key} className={s.section}>
                <Subtitle2>{snippet.title}</Subtitle2>
                <Caption1 className={s.dim}>{snippet.how}</Caption1>
                <div className={s.row}>
                  <pre className={s.code}>{snippet.code}</pre>
                  <Button
                    icon={copied === snippet.key ? <CheckmarkRegular /> : <CopyRegular />}
                    onClick={() => void copy(snippet.key, snippet.code)}
                    aria-label={`Copy the ${snippet.title} settings`}
                  >
                    {copied === snippet.key ? "Copied" : "Copy"}
                  </Button>
                </div>
              </div>
            ))}
            <div className={s.section}>
              <Subtitle2>Windows</Subtitle2>
              <Body1>In the Microsoft Store version, Windows' own AI agents find the server without setup.</Body1>
            </div>
            <Caption1 className={s.dim}>
              The server runs on this PC and talks only to your SharePoint tenant. AI apps see the site names and settings its tools return; sign-ins and secrets stay in Scriptless SharePoint.
            </Caption1>
          </DialogContent>
          <DialogActions>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="primary">Close</Button>
            </DialogTrigger>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
