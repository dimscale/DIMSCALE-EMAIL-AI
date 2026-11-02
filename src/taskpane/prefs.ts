/* global Office */

import type { AnalyzeAction } from "./analyzeClient";

// F-003 per-action auto-mode preferences, stored in the user's mailbox roaming settings (not browser storage).
// "autoSummary" is the pre-existing key and keeps its meaning.
const KEY: Record<AnalyzeAction, string> = {
  summarize: "autoSummary",
  reply: "autoReply",
  translate: "autoTranslate",
};
// Consent given in the first-time dialog, remembered per action (amendment (i), Q-U18).
const CONSENT: Record<AnalyzeAction, string> = {
  summarize: "autoSummaryConsent",
  reply: "autoReplyConsent",
  translate: "autoTranslateConsent",
};

/* True only for a stored boolean true; anything else (absent, "true", 1, exception, no API) is off. */
const flag = (key: string): boolean => {
  try {
    return Office.context.roamingSettings?.get(key) === true;
  } catch {
    return false;
  }
};

export const readAuto = (action: AnalyzeAction): boolean => flag(KEY[action]);
// A stored "on" (the old autoSummary = true) counts as given consent (Q-U22).
export const readConsent = (action: AnalyzeAction): boolean =>
  flag(CONSENT[action]) || flag(KEY[action]);

/* Resolves true when saved. Never rejects: a failure only means the choice lasts for this pane session.
   Turning on also records the consent (the caller only turns on after the dialog or a remembered consent). */
export function saveAuto(action: AnalyzeAction, value: boolean): Promise<boolean> {
  return new Promise((resolve) => {
    try {
      const rs = Office.context.roamingSettings;
      rs.set(KEY[action], value);
      if (value) rs.set(CONSENT[action], true);
      rs.saveAsync((r: Office.AsyncResult<void>) =>
        resolve(r?.status === Office.AsyncResultStatus.Succeeded)
      );
    } catch {
      resolve(false);
    }
  });
}
