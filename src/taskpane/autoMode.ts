import * as React from "react";
import type { AnalyzeAction, AnalyzeErrorCode } from "./analyzeClient";
import { AUTO_MODE_AVAILABLE } from "./config";
import { readAuto, readConsent, saveAuto } from "./prefs";

export const AUTO_DEBOUNCE_MS = 800;

type PerAction<T> = Record<AnalyzeAction, T>;
const all = <T>(f: (a: AnalyzeAction) => T): PerAction<T> => ({
  summarize: f("summarize"),
  reply: f("reply"),
  translate: f("translate"),
});

/* Pane-session auto-mode state for the three switches (lives in App so it survives the per-item remount of AiTools).
   `change(action, true)` must only be called after the consent dialog's "Ligar" or a remembered consent. */
export function useAutoMode() {
  // roamingSettings is not even read when the flag is off.
  const [on, setOn] = React.useState(() => all((a) => AUTO_MODE_AVAILABLE && readAuto(a)));
  const [consent, setConsent] = React.useState(() =>
    all((a) => AUTO_MODE_AVAILABLE && readConsent(a))
  );
  const [saveFailed, setSaveFailed] = React.useState(() => all(() => false));
  // One pause for the three switches: the cap and the backend faults are shared.
  const [paused, setPaused] = React.useState<{
    code: AnalyzeErrorCode;
    action: AnalyzeAction;
  } | null>(null);
  const change = React.useCallback(async (action: AnalyzeAction, value: boolean) => {
    setOn((o) => ({ ...o, [action]: value }));
    if (value) setConsent((c) => ({ ...c, [action]: true }));
    const ok = await saveAuto(action, value);
    setSaveFailed((f) => ({ ...f, [action]: !ok }));
  }, []);
  // "action:email id" pairs that already had a request in this pane session (one request per email per tab). Memory only.
  const attempted = React.useRef(new Set<string>()).current;
  const pause = React.useCallback(
    (code: AnalyzeErrorCode, action: AnalyzeAction) => setPaused({ code, action }),
    []
  );
  const gated = all((a) => AUTO_MODE_AVAILABLE && on[a]);
  return {
    available: AUTO_MODE_AVAILABLE,
    on: gated,
    consent,
    active: (a: AnalyzeAction) => gated[a] && !paused,
    saveFailed,
    paused,
    change,
    pause,
    attempted,
  };
}
export type AutoMode = ReturnType<typeof useAutoMode>;
