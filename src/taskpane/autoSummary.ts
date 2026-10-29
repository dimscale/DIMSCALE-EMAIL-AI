import * as React from "react";
import { AUTO_SUMMARY_AVAILABLE } from "./config";
import { readAutoSummary, saveAutoSummary } from "./prefs";

export const AUTO_DEBOUNCE_MS = 800;

/* Pane-session auto-mode state (lives in App so it survives the per-item remount of AiTools).
   `change(true)` must only be called after the consent dialog's "Ligar". */
export function useAutoSummary() {
  const [on, setOn] = React.useState(() => AUTO_SUMMARY_AVAILABLE && readAutoSummary());
  const [paused, setPaused] = React.useState(false);
  const [saveFailed, setSaveFailed] = React.useState(false);
  const change = React.useCallback(async (value: boolean) => {
    setOn(value);
    setSaveFailed(!(await saveAutoSummary(value)));
  }, []);
  // Items that already had a summary request in this pane session (one request per item). Memory only.
  const attempted = React.useRef(new Set<string>()).current;
  const pause = React.useCallback(() => setPaused(true), []);
  return {
    available: AUTO_SUMMARY_AVAILABLE,
    on: AUTO_SUMMARY_AVAILABLE && on,
    active: AUTO_SUMMARY_AVAILABLE && on && !paused,
    saveFailed,
    change,
    pause,
    attempted,
  };
}
export type AutoSummary = ReturnType<typeof useAutoSummary>;
