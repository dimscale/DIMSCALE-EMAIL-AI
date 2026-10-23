import { type AiConfig, type AiResult, type DraftInput, generateDraft, MARKER_LINE } from "./ai.ts";
import { autoDraftsEnabled, type Get, type Settings } from "./config.ts";
import type { Claimed, Db } from "./db.ts";
import { domainOf, evaluate, isAllowedSender } from "./filters.ts";
import type { Graph, GResult } from "./graph.ts";
import { log } from "./log.ts";

export type ProcDeps = {
  fn: string;
  db: Db;
  graph: Graph;
  settings: Settings;
  get: Get;
  ai?: (input: DraftInput, cfg: AiConfig) => Promise<AiResult>;
  nowMs?: () => number;
  budgetMs?: number; // stop claiming new rows after this (Free plan wall clock is 150 s)
};

const MAX_ATTEMPTS = 5;

export async function processClaimed(d: ProcDeps, userId?: string): Promise<number> {
  if (!autoDraftsEnabled(d.get)) return 0; // kill switch: claim nothing, read nothing
  const now = d.nowMs ?? Date.now;
  const start = now();
  const inboxIds = new Map<string, string>(); // per run, per mailbox
  let n = 0;
  while (now() - start < (d.budgetMs ?? 100_000)) {
    const rows = await d.db.claim(5, userId); // at most one row per mailbox, so one request chain per mailbox
    if (rows.length === 0) break;
    await Promise.all(rows.map(async (row) => {
      const t0 = now();
      try {
        const reason = await processRow(d, row, inboxIds);
        log({ fn: d.fn, user_id: row.userId, reason, attempt: row.attempts, ms: now() - t0 });
      } catch {
        // Unexpected (for example a DB error): leave the row; its lease expires and it is retried.
        log({ fn: d.fn, user_id: row.userId, reason: "internal_error", attempt: row.attempts, ms: now() - t0 });
      }
    }));
    n += rows.length;
  }
  return n;
}

async function processRow(d: ProcDeps, row: Claimed, inboxIds: Map<string, string>): Promise<string> {
  const now = d.nowMs ?? Date.now;
  const done = async (status: "drafted" | "skipped" | "failed", reason: string) => {
    await d.db.finish(row.id, status, reason);
    return `${status}:${reason}`;
  };
  const retryOrFail = async (reason: string, hintSec?: number) => {
    const delaySec = hintSec ?? Math.min(3600, 30 * 2 ** (row.attempts - 1));
    const tooLate = now() + delaySec * 1000 > Date.parse(row.createdAt) + d.settings.maxAgeHours * 3_600_000;
    if (row.attempts >= MAX_ATTEMPTS || tooLate) return await done("failed", reason);
    await d.db.retry(row.id, reason, new Date(now() + delaySec * 1000).toISOString());
    return `retry:${reason}`;
  };
  const graphFail = async (r: Extract<GResult<unknown>, { ok: false }>, userId: string) => {
    if (r.status === 404) return await done("skipped", "message_gone");
    if (r.status === 403) {
      await d.db.setUserStatus(userId, "not_in_scope");
      return await done("failed", "graph_forbidden");
    }
    if (r.reason === "token_unavailable") return await retryOrFail("token_unavailable");
    if (r.status === 429 || r.status === 503) return await retryOrFail("graph_throttled", r.retryAfterSec);
    if (r.status === 0 || r.status === 401 || r.status >= 500) return await retryOrFail("graph_unavailable");
    return await done("failed", "graph_error");
  };

  const user = await d.db.userCtx(row.userId);
  if (!user || !user.enabled || user.contacts.length === 0) return await done("skipped", "user_disabled");

  // Filter 2b: the sender only, before any other content is read.
  const from = await d.graph.readFrom(user.oid, row.messageId);
  if (!from.ok) return await graphFail(from, row.userId);
  if (!isAllowedSender(from.data.from, user.contacts)) return await done("skipped", "sender_not_allowed");

  const msg = await d.graph.readMessage(user.oid, row.messageId);
  if (!msg.ok) return await graphFail(msg, row.userId);

  let inboxFolderId = inboxIds.get(user.oid);
  if (!inboxFolderId) {
    const f = await d.graph.inboxFolderId(user.oid);
    if (!f.ok) return await graphFail(f, row.userId);
    inboxFolderId = f.data;
    inboxIds.set(user.oid, inboxFolderId);
  }
  const verdict = await evaluate(msg.data, {
    mailSha256: user.mailSha256,
    inboxFolderId,
    nowMs: now(),
    countsToday: await d.db.countsToday(row.userId),
  }, d.settings);
  if ("skip" in verdict) return await done("skipped", verdict.skip);

  const ai = await (d.ai ?? ((i, c) => generateDraft(i, c)))(
    { subject: msg.data.subject, fromName: msg.data.fromName, fromDomain: domainOf(msg.data.fromAddress), body: msg.data.body },
    { apiKey: d.get("OPENAI_API_KEY") ?? "", model: d.settings.openaiModel, maxOutputTokens: d.settings.openaiMaxOutputTokens },
  );
  if (!ai.ok) return ai.retryable ? await retryOrFail(ai.reason) : await done("failed", ai.reason);

  // Exactly one draft-creating call. The reply text goes in as the comment, unmodified, after the marker line.
  const draft = await d.graph.createReply(user.oid, row.messageId, `${MARKER_LINE}\n\n${ai.reply}`);
  if (!draft.ok) return await graphFail(draft, row.userId);
  return await done("drafted", "ok");
}
