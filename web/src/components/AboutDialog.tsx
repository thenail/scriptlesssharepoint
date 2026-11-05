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
  tokens,
} from "@fluentui/react-components";

import { useApp } from "../context";

const useStyles = makeStyles({
  content: { display: "flex", gap: "16px", alignItems: "flex-start" },
  logo: { width: "56px", height: "56px" },
  text: { display: "flex", flexDirection: "column", gap: "8px" },
  dim: { color: tokens.colorNeutralForeground3, wordBreak: "break-all" },
});

export function AboutDialog({ onClose }: { onClose(): void }) {
  const { info } = useApp();
  const s = useStyles();
  return (
    <Dialog open onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Scriptless SharePoint</DialogTitle>
          <DialogContent className={s.content}>
            <img src="./logo.png" alt="" className={s.logo} />
            <div className={s.text}>
              <Body1>Version {info?.version}</Body1>
              <Body1>
                Manage SharePoint Online without PowerShell, here or from AI apps through its MCP server.
              </Body1>
              <Caption1 className={s.dim}>Connections and sign-ins are kept on this PC, in {info?.dataDir}.</Caption1>
            </div>
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
