import * as React from "react";
import { getTokenSilent, isNaaSupported, signIn, TokenResult } from "./auth";
import { callWhoami } from "./backend";
import { isConfigured } from "./config";
import { strings } from "./strings";

type State =
  | { kind: "working" }
  | { kind: "notConfigured" }
  | { kind: "unsupported" }
  | { kind: "needSignIn" }
  | { kind: "connected"; name: string | null }
  | { kind: "error"; message: string; detail?: string };

const BackendStatus: React.FC = () => {
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

  switch (state.kind) {
    case "working":
      return <p>{strings.backendWorking}</p>;
    case "notConfigured":
      return <p>{strings.notConfigured}</p>;
    case "unsupported":
      return <p>{strings.unsupported}</p>;
    case "needSignIn":
      return (
        <div>
          <p>{strings.needSignIn}</p>
          <button onClick={onSignIn}>{strings.signIn}</button>
        </div>
      );
    case "connected":
      return <p>{state.name ? strings.connectedAs(state.name) : strings.connected}</p>;
    default:
      return (
        <div>
          <p role="alert">
            {state.message}
            {state.detail ? ` (${state.detail})` : ""}
          </p>
          <button onClick={() => void run()}>{strings.retry}</button>
        </div>
      );
  }
};

export default BackendStatus;
