/* global Office */

import { toReplyHtml } from "./messageView";

/* F-010: the only write call in the add-in. One prependAsync at the top of the draft, never a whole-body replace.
   Resolves true on success. Errors and exceptions are never logged: they may carry content. */
export function insertIntoDraft(item: Office.MessageCompose, text: string): Promise<boolean> {
  return new Promise((resolve) => {
    try {
      item.body.getTypeAsync((t: Office.AsyncResult<Office.CoercionType>) => {
        if (t.status !== Office.AsyncResultStatus.Succeeded) return resolve(false);
        try {
          const html = t.value === Office.CoercionType.Html;
          // HTML: escaped text with <br>; text: plain text. Either way a blank line separates it from the old content.
          const data = html ? `${toReplyHtml(text)}<br><br>` : `${text}\n\n`;
          item.body.prependAsync(data, { coercionType: t.value }, (r: Office.AsyncResult<void>) =>
            resolve(r.status === Office.AsyncResultStatus.Succeeded)
          );
        } catch {
          resolve(false);
        }
      });
    } catch {
      resolve(false);
    }
  });
}
