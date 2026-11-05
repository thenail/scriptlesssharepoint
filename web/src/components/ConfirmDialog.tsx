import {
  Button,
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
import type { ReactNode } from "react";

export interface ConfirmRequest {
  title: string;
  body: ReactNode;
  confirm: string;
  danger?: boolean;
  onConfirm(): unknown;
}

const useStyles = makeStyles({
  danger: {
    backgroundColor: tokens.colorPaletteRedBackground3,
    color: tokens.colorNeutralForegroundOnBrand,
    ":hover": { backgroundColor: tokens.colorPaletteRedForeground1, color: tokens.colorNeutralForegroundOnBrand },
  },
});

export function ConfirmDialog({ request, onClose }: { request: ConfirmRequest; onClose(): void }) {
  const s = useStyles();
  return (
    <Dialog open modalType="alert" onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>{request.title}</DialogTitle>
          <DialogContent>{request.body}</DialogContent>
          <DialogActions>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="secondary">Cancel</Button>
            </DialogTrigger>
            <Button
              appearance="primary"
              className={request.danger ? s.danger : undefined}
              onClick={() => {
                onClose();
                void request.onConfirm();
              }}
            >
              {request.confirm}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
