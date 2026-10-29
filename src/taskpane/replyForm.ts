/* global Office */

/* Opens Outlook's reply form (the user edits and sends). Resolves true when the form was requested. */
export function openReplyForm(
  item: Office.MessageRead | undefined | null,
  html: string
): Promise<boolean> {
  return new Promise((resolve) => {
    try {
      if (!item) return resolve(false);
      if (Office.context.requirements.isSetSupported("Mailbox", "1.9")) {
        item.displayReplyFormAsync(html, (r: Office.AsyncResult<void>) =>
          resolve(r.status === Office.AsyncResultStatus.Succeeded)
        );
      } else {
        item.displayReplyForm(html); // Mailbox 1.1, no callback
        resolve(true);
      }
    } catch {
      resolve(false); // never log the exception: it may carry content
    }
  });
}
