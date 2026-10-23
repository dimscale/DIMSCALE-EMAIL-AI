import { assert, assertEquals, assertFalse } from "@std/assert";
import type { AiResult } from "./ai.ts";
import { readSettings } from "./config.ts";
import type { Graph, GResult, Msg } from "./graph.ts";
import { processClaimed, type ProcDeps } from "./processor.ts";
import { sha256Hex } from "./safe.ts";
import { catchUp, handleLifecycle, syncSubscriptions } from "./subscriptions.ts";
import { fakeDb } from "./testing.ts";

const OWN = "owner@dimscale.com";
const NOW0 = Date.parse("2026-10-23T12:00:00Z");
const ownSha = await sha256Hex(OWN);
const okMsg: Msg = {
  odataType: "#microsoft.graph.message", subject: "Quote", fromAddress: "Ana@Exemplo.pt", fromName: "Ana", toAddresses: [OWN], isDraft: false,
  inference: "focused", receivedDateTime: "2026-10-23T11:30:00Z", headers: {}, body: "Please send the quote.", parentFolderId: "INBOX",
};
const fail = (status: number, extra: Partial<Extract<GResult<never>, { ok: false }>> = {}): GResult<never> => ({ ok: false, status, ...extra });

type G = { from?: GResult<{ from: string }>; msg?: GResult<Msg>; inbox?: GResult<string>; reply?: GResult<{ id: string }> };
function mockGraph(o: G = {}) {
  const calls: string[] = [];
  const graph = {
    readFrom: (_o: string, id: string) => (calls.push(`readFrom:${id}`), Promise.resolve(o.from ?? { ok: true, data: { from: " ANA@exemplo.pt " } })),
    readMessage: (_o: string, id: string) => (calls.push(`readMessage:${id}`), Promise.resolve(o.msg ?? { ok: true, data: okMsg })),
    inboxFolderId: () => (calls.push("inboxFolderId"), Promise.resolve(o.inbox ?? { ok: true, data: "INBOX" })),
    createReply: (_o: string, id: string, comment: string) => (calls.push(`createReply:${id}:${comment}`), Promise.resolve(o.reply ?? { ok: true, data: { id: "D" } })),
  } as unknown as Graph;
  return { graph, calls };
}

function setup(g: G = {}, ai?: (i: unknown) => Promise<AiResult>, env: Record<string, string> = {}) {
  let now = NOW0;
  const f = fakeDb(() => now);
  const user = f.addUser("33333333-3333-3333-3333-333333333333", ["ana@exemplo.pt"], ownSha);
  const mg = mockGraph(g);
  const aiCalls: unknown[] = [];
  const environment: Record<string, string> = { AUTO_DRAFTS_ENABLED: "true", OPENAI_API_KEY: "K", OPENAI_MODEL: "m", ...env };
  const deps: ProcDeps = {
    fn: "t", db: f.db, graph: mg.graph, get: (k) => environment[k], nowMs: () => now,
    settings: readSettings((k) => environment[k]).settings,
    ai: (i, c) => {
      aiCalls.push({ i, c });
      return ai ? ai(i) : Promise.resolve({ ok: true, reply: "Thanks, will send the quote." });
    },
  };
  return { f, user, deps, calls: mg.calls, aiCalls, advance: (ms: number) => (now += ms), now: () => now };
}
const enqueue = (s: ReturnType<typeof setup>, ...ids: string[]) => s.f.db.enqueue(s.user.id, ids);
const row = (s: ReturnType<typeof setup>, i = 0) => s.f.queue[i];
let silence: typeof console.log;
const quiet = () => { silence = console.log; console.log = () => {}; };
const loud = () => { console.log = silence; };

Deno.test("happy path: reads from, then the message, one AI call, one createReply with marker + exact reply, row drafted", async () => {
  const s = setup();
  await enqueue(s, "m1");
  assertEquals(await processClaimed(s.deps), 1);
  assertEquals(s.calls[0], "readFrom:m1");
  assertEquals(s.calls[1], "readMessage:m1");
  assertEquals(s.calls.length, 4);
  assertEquals(s.calls[3], "createReply:m1:[DIMSCALE AI draft - review and edit before sending]\n\nThanks, will send the quote.");
  assertEquals([row(s).status, row(s).reason], ["drafted", "ok"]);
  assertEquals(s.aiCalls.length, 1);
  const a = s.aiCalls[0] as { i: Record<string, string>; c: Record<string, unknown> };
  assertEquals(a.i, { subject: "Quote", fromName: "Ana", fromDomain: "exemplo.pt", body: "Please send the quote." }); // domain only, never the address
  assertEquals(a.c, { apiKey: "K", model: "m", maxOutputTokens: 800 });
});

Deno.test("AC-9: kill switch off claims nothing and calls nothing", async () => {
  for (const v of [undefined, "false", "TRUE"]) {
    const s = setup();
    const deps = { ...s.deps, get: (k: string) => (k === "AUTO_DRAFTS_ENABLED" ? v : "x") };
    await enqueue(s, "m1");
    assertEquals(await processClaimed(deps), 0);
    assertEquals([s.calls.length, s.aiCalls.length, row(s).status], [0, 0, "pending"]);
  }
});

Deno.test("AC-10: user disabled or without contacts after enqueueing: skipped, no Graph read", async () => {
  const s = setup();
  await enqueue(s, "m1", "m2");
  s.user.enabled = false;
  await processClaimed(s.deps);
  s.user.enabled = true;
  s.user.contacts = [];
  await processClaimed(s.deps);
  assertEquals(s.f.queue.map((r) => [r.status, r.reason]), [["skipped", "user_disabled"], ["skipped", "user_disabled"]]);
  assertEquals(s.calls.length, 0);
});

Deno.test("AC-12 filter 2b: not on the list ends sender_not_allowed after the from read only", async () => {
  for (const from of ["rui@exemplo.pt", "ana@exemplo.pt.evil.com", ""]) {
    const s = setup({ from: { ok: true, data: { from } } });
    await enqueue(s, "m1");
    await processClaimed(s.deps);
    assertEquals([row(s).status, row(s).reason], ["skipped", "sender_not_allowed"]);
    assertEquals(s.calls, ["readFrom:m1"]);
    assertEquals(s.aiCalls.length, 0);
  }
});

Deno.test("AC-12: a filtered listed sender ends skipped with its code, no AI, no createReply", async () => {
  const s = setup({ msg: { ok: true, data: { ...okMsg, headers: { "list-unsubscribe": "<x>" } } } });
  await enqueue(s, "m1");
  await processClaimed(s.deps);
  assertEquals([row(s).status, row(s).reason], ["skipped", "mailing_list"]);
  assertEquals(s.aiCalls.length, 0);
  assertFalse(s.calls.some((c) => c.startsWith("createReply")));
  // CC-only: the owner is not among the To recipients
  const t = setup({ msg: { ok: true, data: { ...okMsg, toAddresses: ["x@y.com"] } } });
  await enqueue(t, "m1");
  await processClaimed(t.deps);
  assertEquals(row(t).reason, "not_direct_recipient");
});

Deno.test("AC-14: daily caps end cap_reached without an AI call", async () => {
  const s = setup({}, undefined, { DAILY_CAP_PER_USER: "1" });
  await enqueue(s, "m1", "m2");
  await processClaimed(s.deps);
  assertEquals(s.f.queue.map((r) => [r.status, r.reason]), [["drafted", "ok"], ["skipped", "cap_reached"]]);
  assertEquals(s.aiCalls.length, 1);
});

Deno.test("AC-4: duplicates collapse, claims are exclusive, a live lease blocks, an expired lease is retaken", async () => {
  const s = setup();
  assertEquals(await enqueue(s, "m1"), 1);
  assertEquals(await enqueue(s, "m1"), 0);
  assertEquals(s.f.queue.length, 1);
  const [a, b] = await Promise.all([s.f.db.claim(5), s.f.db.claim(5)]);
  assertEquals(a.length + b.length, 1);
  assertEquals((await s.f.db.claim(5)).length, 0); // lease live
  s.advance(301_000);
  const again = await s.f.db.claim(5);
  assertEquals([again.length, again[0].attempts], [1, 2]);
});

Deno.test("one row per mailbox per claim, different mailboxes in parallel", async () => {
  const s = setup();
  const other = s.f.addUser("44444444-4444-4444-4444-444444444444", ["x@y.com"], ownSha);
  await enqueue(s, "m1", "m2");
  await s.f.db.enqueue(other.id, ["m3"]);
  const rows = await s.f.db.claim(5);
  assertEquals(rows.map((r) => r.messageId).sort(), ["m1", "m3"]);
});

Deno.test("AC-17: bad AI output fails the row, no createReply, no fallback text", async () => {
  const s = setup({}, () => Promise.resolve({ ok: false, reason: "ai_output_invalid", retryable: false }));
  await enqueue(s, "m1");
  await processClaimed(s.deps);
  assertEquals([row(s).status, row(s).reason], ["failed", "ai_output_invalid"]);
  assertFalse(s.calls.some((c) => c.startsWith("createReply")));
});

Deno.test("AC-22: AI unavailable is retried with backoff, then failed after 5 attempts", async () => {
  quiet();
  try {
    const s = setup({}, () => Promise.resolve({ ok: false, reason: "ai_unavailable", retryable: true }));
    await enqueue(s, "m1");
    for (let attempt = 1; attempt <= 5; attempt++) {
      assertEquals(await processClaimed(s.deps), 1);
      if (attempt < 5) {
        assertEquals([row(s).status, row(s).reason, row(s).nextAttemptAt], ["pending", "ai_unavailable", s.now() + 30_000 * 2 ** (attempt - 1)]);
        s.advance(30_000 * 2 ** (attempt - 1));
      }
    }
    assertEquals([row(s).status, row(s).reason, row(s).attempts], ["failed", "ai_unavailable", 5]);
    assertFalse(s.calls.some((c) => c.startsWith("createReply")));
  } finally {
    loud();
  }
});

Deno.test("AC-21: Graph 404, 403, 429 Retry-After, 5xx, token errors, max age", async () => {
  quiet();
  try {
    let s = setup({ from: fail(404) });
    await enqueue(s, "m1");
    await processClaimed(s.deps);
    assertEquals([row(s).status, row(s).reason], ["skipped", "message_gone"]);

    s = setup({ msg: fail(403) });
    await enqueue(s, "m1");
    await processClaimed(s.deps);
    assertEquals([row(s).status, row(s).reason, s.user.status], ["failed", "graph_forbidden", "not_in_scope"]);

    s = setup({ from: fail(429, { retryAfterSec: 120 }) });
    await enqueue(s, "m1");
    await processClaimed(s.deps);
    assertEquals([row(s).status, row(s).reason], ["pending", "graph_throttled"]);
    assert(row(s).nextAttemptAt >= s.now() + 120_000);

    s = setup({ reply: fail(503, { retryAfterSec: 7 }) });
    await enqueue(s, "m1");
    await processClaimed(s.deps);
    assertEquals([row(s).status, row(s).nextAttemptAt], ["pending", s.now() + 7000]);

    s = setup({ from: fail(500) });
    await enqueue(s, "m1");
    await processClaimed(s.deps);
    assertEquals([row(s).status, row(s).reason, row(s).nextAttemptAt], ["pending", "graph_unavailable", s.now() + 30_000]);

    s = setup({ from: fail(0, { reason: "token_unavailable" }) });
    await enqueue(s, "m1");
    await processClaimed(s.deps);
    assertEquals([row(s).status, row(s).reason], ["pending", "token_unavailable"]);

    s = setup({ from: fail(400) });
    await enqueue(s, "m1");
    await processClaimed(s.deps);
    assertEquals([row(s).status, row(s).reason], ["failed", "graph_error"]);

    // a retry that would land beyond MAX_AGE_HOURS fails instead
    s = setup({ from: fail(429, { retryAfterSec: 25 * 3600 }) });
    await enqueue(s, "m1");
    await processClaimed(s.deps);
    assertEquals([row(s).status, row(s).reason], ["failed", "graph_throttled"]);
  } finally {
    loud();
  }
});

Deno.test("AC-26: processor logs contain no ids, addresses or content", async () => {
  const out: string[] = [];
  const orig = console.log;
  console.log = (x: string) => out.push(x);
  try {
    const s = setup({}, () => Promise.resolve({ ok: true, reply: "SECRET-REPLY" }));
    await enqueue(s, "MSG-ID-123");
    await processClaimed(s.deps);
  } finally {
    console.log = orig;
  }
  assert(out.length > 0);
  for (const line of out) for (const bad of ["MSG-ID-123", "ana@", "Quote", "SECRET-REPLY", "33333333", "exemplo"]) assertFalse(line.includes(bad), `${bad} in ${line}`);
});

// ---------- subscriptions ----------
type SG = Partial<Record<"createSubscription" | "renewSubscription" | "deleteSubscription" | "listInboxIds", (...a: unknown[]) => Promise<GResult<unknown>>>>;
function subSetup(o: SG = {}) {
  const now = NOW0;
  const f = fakeDb(() => now);
  const calls: string[] = [];
  const dflt = {
    createSubscription: (oid: unknown) => (calls.push(`create:${oid}`), Promise.resolve({ ok: true, data: { id: `S-${oid}`, expiresAt: "2026-11-01T00:00:00Z" } })),
    renewSubscription: (id: unknown) => (calls.push(`patch:${id}`), Promise.resolve({ ok: true, data: { expiresAt: "2026-11-02T00:00:00Z" } })),
    deleteSubscription: (id: unknown) => (calls.push(`delete:${id}`), Promise.resolve({ ok: true, data: null })),
    listInboxIds: (oid: unknown, since: unknown) => (calls.push(`list:${oid}:${since}`), Promise.resolve({ ok: true, data: { ids: ["a", "b"], truncated: false } })),
  };
  const graph = { ...dflt, ...o } as unknown as Graph;
  return { f, calls, d: { fn: "t", db: f.db, graph, nowMs: () => now } };
}
const iso = (ms: number) => new Date(ms).toISOString();

Deno.test("AC-7: renewal creates, patches near expiry, deletes for disabled, honours the 10 minute guard", async () => {
  quiet();
  try {
    const s = subSetup();
    s.f.addUser("A", ["a@x.com"], "", true, "pending");
    const b = s.f.addUser("B", ["a@x.com"]);
    const c = s.f.addUser("C", ["a@x.com"]);
    const d = s.f.addUser("D", [], "", false, "off");
    const e = s.f.addUser("E", ["a@x.com"]);
    await s.f.db.setSubscription(b.id, "SB", iso(NOW0 + 24 * 3600_000));
    await s.f.db.setSubscription(c.id, "SC", iso(NOW0 + 24 * 3600_000));
    await s.f.db.setSubscription(d.id, "SD", iso(NOW0 + 24 * 3600_000));
    await s.f.db.setSubscription(e.id, "SE", iso(NOW0 + 5 * 24 * 3600_000));
    s.f.subs.get(b.id)!.lastRenewedAt = iso(NOW0 - 3600_000);
    s.f.subs.get(c.id)!.lastRenewedAt = iso(NOW0 - 120_000); // patched 2 min ago: skipped
    await syncSubscriptions(s.d, false);
    assertEquals(s.calls.sort(), ["create:A", "delete:SD", "patch:SB"]);
    assertEquals(s.f.subs.has(d.id), false);
    assertEquals(s.f.users.get("A")!.status, "active");
  } finally {
    loud();
  }
});

Deno.test("AC-40: 403 on creation sets not_in_scope; only the renew run retries it", async () => {
  quiet();
  try {
    const s = subSetup({ createSubscription: () => Promise.resolve({ ok: false, status: 403 }) });
    const u = s.f.addUser("A", ["a@x.com"], "", true, "pending");
    await syncSubscriptions(s.d, false);
    assertEquals(u.status, "not_in_scope");
    const calls = s.calls.length;
    await syncSubscriptions(s.d, false);
    assertEquals(s.calls.length, calls); // process run: not retried
    let created = 0;
    const s2 = subSetup({ createSubscription: () => (created++, Promise.resolve({ ok: false, status: 403 })) });
    s2.f.addUser("A", ["a@x.com"], "", true, "not_in_scope");
    await syncSubscriptions(s2.d, true);
    assertEquals(created, 1);
    const s3 = subSetup({ createSubscription: () => Promise.resolve({ ok: false, status: 500 }) });
    const u3 = s3.f.addUser("A", ["a@x.com"], "", true, "pending");
    await syncSubscriptions(s3.d, false);
    assertEquals(u3.status, "error");
  } finally {
    loud();
  }
});

Deno.test("AC-8: catch-up window is max(last catch-up, now - MAX_AGE), ids enqueued idempotently", async () => {
  const s = subSetup();
  const u = s.f.addUser("A", ["a@x.com"]);
  await s.f.db.setSubscription(u.id, "SA", iso(NOW0 + 9 * 24 * 3600_000));
  await catchUp(s.d, { id: u.id, oid: "A", lastCatchupAt: null }, 24);
  assertEquals(s.calls[0], `list:A:${iso(NOW0 - 24 * 3600_000)}`);
  assertEquals(s.f.queue.length, 2);
  assertEquals(s.f.subs.get(u.id)!.lastCatchupAt, iso(NOW0));
  await catchUp(s.d, { id: u.id, oid: "A", lastCatchupAt: iso(NOW0 - 3600_000) }, 24);
  assertEquals(s.calls[1], `list:A:${iso(NOW0 - 3600_000 - 300_000)}`); // 5 minute overlap (L5)
  assertEquals(s.f.queue.length, 2); // same ids: no new rows
  await catchUp(s.d, { id: u.id, oid: "A", lastCatchupAt: iso(NOW0 - 72 * 3600_000) }, 24);
  assertEquals(s.calls[2], `list:A:${iso(NOW0 - 24 * 3600_000)}`);
});

Deno.test("AC-5: lifecycle events", async () => {
  quiet();
  try {
    const s = subSetup();
    const u = s.f.addUser("A", ["a@x.com"]);
    await s.f.db.setSubscription(u.id, "SA", iso(NOW0 + 3600_000));
    await handleLifecycle(s.d, u.id, "SA", "reauthorizationRequired", 24);
    assertEquals(s.calls, ["patch:SA"]); // one PATCH, no separate reauthorize call
    s.calls.length = 0;
    await handleLifecycle(s.d, u.id, "SA", "subscriptionRemoved", 24);
    assertEquals(s.calls.map((c) => c.split(":")[0]), ["create", "list"]);
    assertEquals(s.f.subs.get(u.id)!.subId, "S-A");
    s.calls.length = 0;
    await handleLifecycle(s.d, u.id, "S-A", "missed", 24);
    assertEquals(s.calls.map((c) => c.split(":")[0]), ["list"]);
    s.calls.length = 0;
    await handleLifecycle(s.d, u.id, "S-A", "somethingElse", 24);
    assertEquals(s.calls, []);
  } finally {
    loud();
  }
});

// ---------- security review fixes ----------
Deno.test("M1: a display name that is or contains an address never reaches the AI", async () => {
  for (const fromName of ["ana@exemplo.pt", "ANA@Exemplo.PT", "Ana <ana@exemplo.pt>", "x@y.com"]) {
    const s = setup({ msg: { ok: true, data: { ...okMsg, fromName } } });
    await enqueue(s, "m1");
    await processClaimed(s.deps);
    assertEquals((s.aiCalls[0] as { i: { fromName: string } }).i.fromName, "");
  }
  const s = setup({ msg: { ok: true, data: { ...okMsg, fromName: "Ana Silva" } } });
  await enqueue(s, "m1");
  await processClaimed(s.deps);
  assertEquals((s.aiCalls[0] as { i: { fromName: string } }).i.fromName, "Ana Silva");
});

Deno.test("M2: an expired lease after 5 attempts fails the row (lease_expired) instead of reclaiming it; a row over the cap is failed before any call", async () => {
  const s = setup();
  await enqueue(s, "m1");
  for (let i = 0; i < 5; i++) {
    assertEquals((await s.f.db.claim(5)).length, 1);
    s.advance(301_000);
  }
  assertEquals(row(s).attempts, 5);
  assertEquals((await s.f.db.claim(5)).length, 0);
  assertEquals([row(s).status, row(s).reason], ["failed", "lease_expired"]);
  // processor guard: a claimed row with attempts > MAX_ATTEMPTS
  const t = setup();
  await enqueue(t, "m1");
  row(t).attempts = 5; // the next claim makes it 6
  assertEquals(await processClaimed(t.deps), 1);
  assertEquals([row(t).status, row(t).reason, t.calls.length, t.aiCalls.length], ["failed", "lease_expired", 0, 0]);
});

Deno.test("M3: a successful draft sets the user status to active; a successful renewal does too (enabled users only)", async () => {
  const s = setup();
  s.user.status = "error";
  await enqueue(s, "m1");
  await processClaimed(s.deps);
  assertEquals(s.user.status, "active");
  quiet();
  try {
    const r = subSetup();
    const u = r.f.addUser("A", ["a@x.com"], "", true, "error");
    await r.f.db.setSubscription(u.id, "SA", iso(NOW0 + 3600_000));
    r.f.subs.get(u.id)!.lastRenewedAt = iso(NOW0 - 3600_000); // outside the 10 minute guard
    await syncSubscriptions(r.d, false);
    assertEquals([r.calls, u.status], [["patch:SA"], "active"]);
    const l = subSetup();
    const off = l.f.addUser("B", [], "", false, "off");
    await l.f.db.setSubscription(off.id, "SB", iso(NOW0 + 3600_000));
    await handleLifecycle(l.d, off.id, "SB", "reauthorizationRequired", 24);
    assertEquals(off.status, "off"); // a disabled user is not resurrected
  } finally {
    loud();
  }
});

Deno.test("M5: daily caps count drafted rows and rows that failed at the AI step, not other failures", async () => {
  const s = setup({}, () => Promise.resolve({ ok: false, reason: "ai_output_invalid", retryable: false }), { DAILY_CAP_PER_USER: "2" });
  await enqueue(s, "m1", "m2", "m3");
  await processClaimed(s.deps);
  assertEquals(s.f.queue.map((r) => [r.status, r.reason]), [["failed", "ai_output_invalid"], ["failed", "ai_output_invalid"], ["skipped", "cap_reached"]]);
  assertEquals(s.aiCalls.length, 2);
  const g = setup({ from: fail(400) }); // graph_error is not an AI-step failure
  await enqueue(g, "m1");
  await processClaimed(g.deps);
  assertEquals(await g.f.db.countsToday(g.user.id), { user: 0, global: 0 });
});

Deno.test("L4: missing AI config fails config_missing after the sender check and before the full read, and does not use the daily caps", async () => {
  for (const env of [{ OPENAI_API_KEY: "" }, { OPENAI_MODEL: "" }] as Record<string, string>[]) {
    const s = setup({}, undefined, env);
    await enqueue(s, "m1");
    await processClaimed(s.deps);
    assertEquals([row(s).status, row(s).reason], ["failed", "config_missing"]);
    assertEquals(s.calls, ["readFrom:m1"]);
    assertEquals(await s.f.db.countsToday(s.user.id), { user: 0, global: 0 });
  }
});

Deno.test("L5: a truncated catch-up enqueues what it found but does not advance last_catchup_at", async () => {
  const s = subSetup({ listInboxIds: () => Promise.resolve({ ok: true, data: { ids: ["a"], truncated: true } }) });
  const u = s.f.addUser("A", ["a@x.com"]);
  await s.f.db.setSubscription(u.id, "SA", iso(NOW0 + 9 * 24 * 3600_000));
  quiet();
  try {
    await catchUp(s.d, { id: u.id, oid: "A", lastCatchupAt: null }, 24);
  } finally {
    loud();
  }
  assertEquals([s.f.queue.length, s.f.subs.get(u.id)!.lastCatchupAt], [1, null]);
});

Deno.test("L-2: a disabled user is never set active; an enabled user is", async () => {
  const s = setup();
  s.user.enabled = false;
  s.user.status = "error";
  await s.f.db.setUserStatus(s.user.id, "active");
  assertEquals(s.user.status, "error");
  await s.f.db.setUserStatus(s.user.id, "not_in_scope");
  assertEquals(s.user.status, "not_in_scope");
  s.user.enabled = true;
  await s.f.db.setUserStatus(s.user.id, "active");
  assertEquals(s.user.status, "active");
});
