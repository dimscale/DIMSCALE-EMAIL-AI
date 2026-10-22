/* global Office */

import { strings } from "./strings";

export const BODY_LIMIT = 2000;

export interface MessageView {
  subject: string;
  senderName: string;
  senderEmail: string;
  date: string; // Office.js MessageRead exposes only dateTimeCreated, not a received date
  body: string; // already truncated; empty string when the message has no text
  truncated: boolean;
  bodyLength: number;
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
    // drop a trailing high surrogate left by cutting an emoji in half
    body: truncated ? text.slice(0, BODY_LIMIT).replace(/[\uD800-\uDBFF]$/, "") : text,
    truncated,
    bodyLength: text.length,
  };
}
