/* global Office */

// F-003 auto-summary preference, stored in the user's mailbox roaming settings (not browser storage).
const KEY = "autoSummary";

/* True only for a stored boolean true; anything else (absent, "true", 1, exception, no API) is off. */
export function readAutoSummary(): boolean {
  try {
    return Office.context.roamingSettings?.get(KEY) === true;
  } catch {
    return false;
  }
}

/* Resolves true when saved. Never rejects: a failure only means the choice lasts for this pane session. */
export function saveAutoSummary(value: boolean): Promise<boolean> {
  return new Promise((resolve) => {
    try {
      const rs = Office.context.roamingSettings;
      rs.set(KEY, value);
      rs.saveAsync((r: Office.AsyncResult<void>) =>
        resolve(r?.status === Office.AsyncResultStatus.Succeeded)
      );
    } catch {
      resolve(false);
    }
  });
}
