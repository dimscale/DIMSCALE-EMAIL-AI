import * as React from "react";
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Switch,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import type { AutoSummary } from "./autoSummary";
import { strings } from "./strings";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: "4px", minWidth: 0 },
  note: {
    margin: 0,
    overflowWrap: "anywhere",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: "18px",
  },
});

/* Auto-mode switch. Off to on always goes through the consent dialog; only "Ligar" turns it on. */
const AutoSummarySwitch: React.FC<{ auto: AutoSummary }> = ({ auto }) => {
  const s = useStyles();
  const [open, setOpen] = React.useState(false);
  return (
    <div className={s.root}>
      <Switch
        label={strings.autoSummaryLabel}
        checked={auto.on}
        onChange={(_, d) => (d.checked ? setOpen(true) : void auto.change(false))}
      />
      <Dialog modalType="modal" open={open} onOpenChange={(_, d) => setOpen(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{strings.autoSummaryDialogTitle}</DialogTitle>
            <DialogContent>{strings.autoSummaryDialogBody}</DialogContent>
            <DialogActions>
              <Button
                appearance="primary"
                onClick={() => {
                  setOpen(false);
                  void auto.change(true);
                }}
              >
                {strings.autoSummaryConfirm}
              </Button>
              <Button appearance="secondary" onClick={() => setOpen(false)}>
                {strings.autoSummaryCancel}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
      {auto.on && <p className={s.note}>{strings.autoSummaryOnNote}</p>}
      {auto.saveFailed && (
        <p className={s.note} role="status">
          {strings.autoSummarySaveFailed}
        </p>
      )}
    </div>
  );
};

export default AutoSummarySwitch;
