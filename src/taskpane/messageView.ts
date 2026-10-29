/* global Office */

import { strings } from "./strings";

export const BODY_LIMIT = 2000;
export const AI_TEXT_LIMIT = 4000;
export const REPLY_FORM_LIMIT = 32 * 1024;

export interface MessageView {
  subject: string;
  senderName: string;
  senderEmail: string;
  date: string; // Office.js MessageRead exposes only dateTimeCreated, not a received date
  body: string; // already truncated; empty string when the message has no text
  truncated: boolean;
  bodyLength: number;
  // F-003: what the AI request carries, cut from the full body (not from the preview)
  aiSubject: string;
  aiFromName: string;
  fromDomain: string;
  aiText: string;
}

// cut to n UTF-16 units without leaving a lone high surrogate (e.g. an emoji cut in half)
const cut = (t: string, n: number) => t.slice(0, n).replace(/[\uD800-\uDBFF]$/, "");

/* The part after the last "@", lower-cased; "" without "@". Never the local part. */
export const fromDomain = (address: string | undefined) => {
  const a = address ?? "";
  const i = a.lastIndexOf("@");
  return i < 0
    ? ""
    : a
        .slice(i + 1)
        .trim()
        .toLowerCase();
};

const REPLY_ESCAPES: Record<string, string> = {
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  '"': "&quot;",
  "'": "&#39;",
};

/* Reply-form body: displayReplyForm* reads the string as HTML, so the text is escaped and newlines become <br>. Capped at 32 KB. */
export function toReplyHtml(text: string): string {
  let out = "";
  for (const ch of text.replace(/\r\n|\r/g, "\n")) {
    const piece = ch === "\n" ? "<br>" : (REPLY_ESCAPES[ch] ?? ch);
    if (out.length + piece.length > REPLY_FORM_LIMIT) break;
    out += piece;
  }
  return out;
}

/* Pure: maps the Office.js message item and its plain-text body to what the panel shows. */
export function toMessageView(item: Office.MessageRead, bodyText: string): MessageView {
  const text = bodyText ?? "";
  const truncated = text.length > BODY_LIMIT;
  return {
    subject: item.subject?.trim() ? item.subject : strings.noSubject,
    senderName: item.from?.displayName ?? "",
    senderEmail: item.from?.emailAddress ?? "",
    date: item.dateTimeCreated ? item.dateTimeCreated.toLocaleString() : "",
    body: truncated ? cut(text, BODY_LIMIT) : text,
    truncated,
    bodyLength: text.length,
    aiSubject: cut(item.subject ?? "", 200),
    aiFromName: cut(item.from?.displayName ?? "", 100),
    fromDomain: fromDomain(item.from?.emailAddress),
    aiText: cut(text.trim(), AI_TEXT_LIMIT),
  };
}
