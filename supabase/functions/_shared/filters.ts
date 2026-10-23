import type { Msg } from "./graph.ts";
import type { Settings } from "./config.ts";
import { normalizeAddress, sha256Hex } from "./safe.ts";

/** Filter 2b: exact match on the lower-cased, trimmed address (never by domain). */
export const isAllowedSender = (from: string | null | undefined, allowList: string[]): boolean => {
  if (!from) return false;
  const f = normalizeAddress(from);
  return f !== "" && allowList.some((a) => normalizeAddress(a) === f);
};

const NOREPLY = /^(noreply|no-reply|donotreply|do-not-reply|mailer-daemon|postmaster)$/i;
export const isNoReplyLocalPart = (address: string) => NOREPLY.test(address.split("@")[0] ?? "");
export const domainOf = (address: string) => normalizeAddress(address).split("@")[1] ?? "";

export type FilterCtx = {
  mailSha256: string; // SHA-256 of the owner's lower-cased address
  inboxFolderId: string;
  nowMs: number;
  countsToday: { user: number; global: number }; // rows that reached the AI step: drafted, plus failed with an ai_* reason
};
export type FilterResult = { pass: true } | { skip: string };

/** Filters 3 to 17 in the spec's order; the first hit wins. Pure apart from hashing. */
export async function evaluate(m: Msg, c: FilterCtx, s: Settings): Promise<FilterResult> {
  if (m.isDraft || !c.inboxFolderId || m.parentFolderId !== c.inboxFolderId) return { skip: "not_inbox_message" };
  if (m.odataType.startsWith("#microsoft.graph.eventMessage")) return { skip: "event_message" };
  if (await sha256Hex(normalizeAddress(m.fromAddress)) === c.mailSha256) return { skip: "from_self" };
  const toHashes = await Promise.all(m.toAddresses.map((a) => sha256Hex(normalizeAddress(a))));
  if (!toHashes.includes(c.mailSha256)) return { skip: "not_direct_recipient" };
  const h = m.headers;
  const autoSubmitted = h["auto-submitted"];
  if (autoSubmitted !== undefined && autoSubmitted.trim().toLowerCase() !== "no") return { skip: "auto_submitted" };
  if (s.filterAutoResponseSuppress && h["x-auto-response-suppress"] !== undefined) return { skip: "auto_response_suppress" };
  if (["bulk", "list", "junk"].includes((h["precedence"] ?? "").trim().toLowerCase())) return { skip: "bulk_precedence" };
  if (h["list-id"] !== undefined || h["list-unsubscribe"] !== undefined) return { skip: "mailing_list" };
  if (isNoReplyLocalPart(m.fromAddress)) return { skip: "noreply_sender" };
  if (s.filterFocusedOther && m.inference.toLowerCase() === "other") return { skip: "focused_other" };
  const received = Date.parse(m.receivedDateTime);
  if (!Number.isFinite(received) || c.nowMs - received > s.maxAgeHours * 3_600_000) return { skip: "too_old" };
  if (m.body.trim() === "") return { skip: "empty_body" };
  const dom = domainOf(m.fromAddress);
  if (s.skipSenderDomains.includes(dom)) return { skip: "sender_domain_skipped" };
  if (s.internalSenders === "skip" && s.internalDomains.includes(dom)) return { skip: "internal_sender" };
  if (c.countsToday.user >= s.dailyCapPerUser || c.countsToday.global >= s.dailyCapGlobal) return { skip: "cap_reached" };
  return { pass: true };
}
