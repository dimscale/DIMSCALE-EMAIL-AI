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
  useRestoreFocusTarget,
} from "@fluentui/react-components";
import type { AnalyzeAction } from "./analyzeClient";
import type { AutoMode } from "./autoMode";
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

const TEXT: Record<AnalyzeAction, { aria: string; title: string; body: string; note: string }> = {
  summarize: {
    aria: strings.autoSummaryAria,
    title: strings.autoSummaryDialogTitle,
    body: strings.autoSummaryDialogBody,
    note: strings.autoSummaryOnNote,
  },
  reply: {
    aria: strings.autoReplyAria,
    title: strings.autoReplyDialogTitle,
    body: strings.autoReplyDialogBody,
    note: strings.autoReplyOnNote,
  },
  translate: {
    aria: strings.autoTranslateAria,
    title: strings.autoTranslateDialogTitle,
    body: strings.autoTranslateDialogBody,
    note: strings.autoTranslateOnNote,
  },
};

/* Per-action auto-mode switch. The consent dialog shows only until that action's consent is remembered;
   only "Ligar" (or a remembered consent) turns the switch on. */
const AutoSwitch: React.FC<{ auto: AutoMode; action: AnalyzeAction; alertShown?: boolean }> = ({
  auto,
  action,
  alertShown,
}) => {
  const s = useStyles();
  const [open, setOpen] = React.useState(false);
  const restoreFocus = useRestoreFocusTarget(); // focus returns to the switch when the dialog closes
  const t = TEXT[action];
  const on = auto.on[action];
  const pausedHere = on && auto.paused && !alertShown ? auto.paused : null; // the alert already says it
  return (
    <div className={s.root}>
      <Switch
        {...restoreFocus}
        label={strings.autoSwitchLabel}
        aria-label={t.aria}
        checked={on}
        onChange={(_, d) => {
          if (!d.checked) void auto.change(action, false);
          else if (auto.consent[action]) void auto.change(action, true);
          else setOpen(true);
        }}
      />
      <Dialog modalType="modal" open={open} onOpenChange={(_, d) => setOpen(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{t.title}</DialogTitle>
            <DialogContent>{t.body}</DialogContent>
            <DialogActions>
              <Button
                appearance="primary"
                onClick={() => {
                  setOpen(false);
                  void auto.change(action, true);
                }}
              >
                {strings.autoConfirm}
              </Button>
              <Button appearance="secondary" onClick={() => setOpen(false)}>
                {strings.autoCancel}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
      {on && <p className={s.note}>{t.note}</p>}
      {pausedHere && (
        <p className={s.note} role="status">
          {pausedHere.code === "cap" ? strings.autoPausedNoteCap : strings.autoPausedNoteOther}
        </p>
      )}
      {auto.saveFailed[action] && (
        <p className={s.note} role="status">
          {strings.autoSaveFailed}
        </p>
      )}
    </div>
  );
};

export default AutoSwitch;
