import type { Msg } from "./graph.ts";
import type { Settings } from "./config.ts";
import { normalizeAddress, sha256Hex } from "./safe.ts";

/** Filter 2b: exact match on the lower-cased, trimmed address (never by domain). */
export const isAllowedSender = (from: string | null | undefined, allowList: string[]): boolean => {
  if (!from) return false;
  const f = normalizeAddress(from);
  return f !== "" && allowList.some((a) => normalizeAddress(a) === f);
};

// F-013/AC-49 + amendment 2026-11-03: NFKC, drop ignorable/format characters, lower-case, trim, cut at the first "+". Mirrored in src/taskpane/noReply.ts.
const NOREPLY = /^(noreply|no-reply|donotreply|do-not-reply|mailer-daemon|postmaster)$/;
export const isNoReplyLocalPart = (address: string) => {
  const local = (address.split("@")[0] ?? "").normalize("NFKC").replace(/[\p{Default_Ignorable_Code_Point}\p{Cf}]/gu, "").toLowerCase().trim();
  return NOREPLY.test(local.split("+")[0]);
};
export const domainOf = (address: string) => normalizeAddress(address).split("@")[1] ?? "";

/** Filter 18: true only on an explicit dmarc=fail or compauth=fail token; anything missing or odd passes (fail-open). */
const AUTH_FAIL = /(?:^|[\s;])(?:dmarc|compauth)=fail(?=[\s;(]|$)/i;
// Parenthesised comments (nested, one linear pass) are removed first; an unclosed one drops the rest (fewer skips, fail-open).
const stripComments = (v: string) => {
  let depth = 0, out = "";
  for (const ch of v) {
    if (ch === "(") depth++;
    else if (ch === ")" && depth > 0) { depth--; if (depth === 0) out += " "; } else if (depth === 0) out += ch;
  }
  return out;
};
export const authFailed = (value: string | undefined) => value !== undefined && AUTH_FAIL.test(stripComments(value));

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
  if (s.filterAuthFailed && authFailed(h["authentication-results"])) return { skip: "auth_failed" };
  const received = Date.parse(m.receivedDateTime);
  if (!Number.isFinite(received) || c.nowMs - received > s.maxAgeHours * 3_600_000) return { skip: "too_old" };
  if (m.body.trim() === "") return { skip: "empty_body" };
  const dom = domainOf(m.fromAddress);
  if (s.skipSenderDomains.includes(dom)) return { skip: "sender_domain_skipped" };
  if (s.internalSenders === "skip" && s.internalDomains.includes(dom)) return { skip: "internal_sender" };
  if (c.countsToday.user >= s.dailyCapPerUser || c.countsToday.global >= s.dailyCapGlobal) return { skip: "cap_reached" };
  return { pass: true };
}
