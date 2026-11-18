/* global Office */

import { AnalyzePayload, ReplyLanguage, ReplyTone } from "./analyzeClient";
import { AI_TEXT_LIMIT, cut } from "./messageView";

/* F-010: compose item helpers. Compose items expose subject as an object with getAsync; read items as a string. */
export const isCompose = (item: unknown): item is Office.MessageCompose =>
  !!item && typeof (item as { subject?: { getAsync?: unknown } }).subject?.getAsync === "function";

export type Draft = { subject: string; body: string };

// Resolves null on any failure or exception. The error is never logged: it may carry content.
const read = (call: (cb: (r: Office.AsyncResult<string>) => void) => void) =>
  new Promise<string | null>((resolve) => {
    try {
      call((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null));
    } catch {
      resolve(null);
    }
  });

/* Only the subject and the body text are read. Recipients, sender and ids are never touched. */
export async function readDraft(item: Office.MessageCompose): Promise<Draft | null> {
  const [subject, body] = await Promise.all([
    read((cb) => item.subject.getAsync(cb)),
    read((cb) => item.body.getAsync(Office.CoercionType.Text, cb)),
  ]);
  return subject === null || body === null
    ? null
    : { subject: cut(subject, 200), body: cut(body.trim(), AI_TEXT_LIMIT) };
}

export const composePayload = (
  d: Draft,
  language: ReplyLanguage,
  tone: ReplyTone,
  refine?: { instruction: string; draft: string }
): AnalyzePayload => ({
  action: "reply",
  subject: d.subject,
  fromName: "",
  fromDomain: "",
  body: d.body,
  language,
  tone,
  ...(refine ?? {}),
});
