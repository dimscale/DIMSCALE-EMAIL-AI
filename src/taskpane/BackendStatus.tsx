import * as React from "react";
import { createPortal } from "react-dom";
import { getTokenSilent, isNaaSupported, signIn, TokenResult } from "./auth";
import { callWhoami } from "./backend";
import { isConfigured } from "./config";
import { strings } from "./strings";
import { Button, MessageBar, MessageBarActions, MessageBarBody, makeStyles, tokens } from "@fluentui/react-components";
import { CheckmarkCircleFilled, ClockRegular, ErrorCircleRegular, WarningRegular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  line: { display: "flex", alignItems: "center", gap: "6px", minWidth: 0, overflowWrap: "anywhere", fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 },
  icon: { display: "flex", fontSize: "16px", flex: "none" },
});

type State =
  | { kind: "working" }
  | { kind: "notConfigured" }
  | { kind: "unsupported" }
  | { kind: "needSignIn" }
  | { kind: "connected"; name: string | null }
  | { kind: "error"; message: string; detail?: string };

// noticeHost: where the MessageBar goes (above the tab content). undefined renders it inline; null waits for the host.
const BackendStatus: React.FC<{ noticeHost?: HTMLElement | null }> = ({ noticeHost }) => {
  const [state, setState] = React.useState<State>({ kind: "working" });
  const alive = React.useRef(true);
  React.useEffect(() => {
    alive.current = true;
    return () => {
      alive.current = false;
    };
  }, []);
  const set = (s: State) => alive.current && setState(s);

  const withToken = async (t: TokenResult, fromPopup: boolean) => {
    if (t.kind === "interaction") return set({ kind: "needSignIn" });
    if (t.kind === "error") return set({ kind: "error", message: fromPopup ? strings.signInFailed : strings.signInError, detail: t.code });
    const r = await callWhoami(t.token);
    if (r.kind === "ok") return set({ kind: "connected", name: r.name });
    if (r.kind === "rejected") return set({ kind: "error", message: strings.rejected, detail: "401" });
    set({ kind: "error", message: strings.unreachable, detail: r.status ? String(r.status) : undefined });
  };

  const run = React.useCallback(async () => {
    setState({ kind: "working" });
    if (!isNaaSupported()) return set({ kind: "unsupported" });
    if (!isConfigured()) return set({ kind: "notConfigured" });
    await withToken(await getTokenSilent(), false);
  }, []);

  React.useEffect(() => {
    void run();
  }, [run]);

  const onSignIn = async () => {
    setState({ kind: "working" });
    await withToken(await signIn(), true);
  };

  const s = useStyles();
  // With a MessageBar showing the same information the line is hidden from AT to avoid a double announcement.
  const line = (icon: React.ReactNode, text: string, notice = false) => (
    <div className={s.line} role={notice ? undefined : "status"} aria-hidden={notice ? "true" : undefined}>
      <span aria-hidden="true" className={s.icon}>
        {icon}
      </span>
      <span>{text}</span>
    </div>
  );
  const notice = (bar: React.ReactNode) =>
    noticeHost === undefined ? bar : noticeHost ? createPortal(bar, noticeHost) : null;
  switch (state.kind) {
    case "working":
      return line(<ClockRegular />, strings.backendWorking);
    case "notConfigured":
      return (
        <div>
          {notice(<MessageBar intent="info">
            <MessageBarBody>{strings.notConfigured}</MessageBarBody>
          </MessageBar>)}
          {line(<WarningRegular />, strings.shortNotConfigured, true)}
        </div>
      );
    case "unsupported":
      return (
        <div>
          {notice(<MessageBar intent="info">
            <MessageBarBody>{strings.unsupported}</MessageBarBody>
          </MessageBar>)}
          {line(<WarningRegular />, strings.shortUnsupported, true)}
        </div>
      );
    case "needSignIn":
      return (
        <div>
          {notice(<MessageBar intent="info">
            <MessageBarBody>{strings.needSignIn}</MessageBarBody>
            <MessageBarActions>
              <Button onClick={onSignIn}>{strings.signIn}</Button>
            </MessageBarActions>
          </MessageBar>)}
          {line(<WarningRegular />, strings.shortNeedSignIn, true)}
        </div>
      );
    case "connected":
      return line(<CheckmarkCircleFilled />, state.name ? strings.connectedAs(state.name) : strings.connected);
    default:
      return (
        <div>
          {notice(<MessageBar intent="error" role="alert">
            <MessageBarBody>
              {state.message}
              {state.detail ? ` (${state.detail})` : ""}
            </MessageBarBody>
            <MessageBarActions>
              <Button onClick={() => void run()}>{strings.retry}</Button>
            </MessageBarActions>
          </MessageBar>)}
          {line(<ErrorCircleRegular />, strings.shortError, true)}
        </div>
      );
  }
};

export default BackendStatus;
