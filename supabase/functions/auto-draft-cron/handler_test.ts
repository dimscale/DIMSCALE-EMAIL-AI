import { assertEquals, assertFalse } from "@std/assert";
import { readSettings } from "../_shared/config.ts";
import type { Graph } from "../_shared/graph.ts";
import type { ProcDeps } from "../_shared/processor.ts";
import { sha256Hex } from "../_shared/safe.ts";
import { fakeDb } from "../_shared/testing.ts";
import { type CronDeps, handle } from "./handler.ts";

const OID = "33333333-3333-3333-3333-333333333333";
const OWN = "owner@dimscale.com";

async function setup(env: Record<string, string> = { AUTO_DRAFTS_ENABLED: "true", CRON_SHARED_SECRET: "SECRET", OPENAI_API_KEY: "K", OPENAI_MODEL: "m" }) {
  const f = fakeDb();
  const u = f.addUser(OID, ["ana@exemplo.pt"], await sha256Hex(OWN));
  const calls: string[] = [];
  const graph = {
    createSubscription: () => (calls.push("create"), Promise.resolve({ ok: true, data: { id: "S1", expiresAt: "2026-11-01T00:00:00Z" } })),
    renewSubscription: () => (calls.push("patch"), Promise.resolve({ ok: true, data: { expiresAt: "2026-11-01T00:00:00Z" } })),
    deleteSubscription: () => (calls.push("delete"), Promise.resolve({ ok: true, data: null })),
    listInboxIds: () => (calls.push("list"), Promise.resolve({ ok: true, data: ["m1"] })),
    readFrom: () => (calls.push("readFrom"), Promise.resolve({ ok: true, data: { from: "ana@exemplo.pt" } })),
    readMessage: () => (calls.push("readMessage"), Promise.resolve({ ok: true, data: { odataType: "", subject: "S", fromAddress: "ana@exemplo.pt", fromName: "Ana", toAddresses: [OWN], isDraft: false, inference: "focused", receivedDateTime: new Date().toISOString(), headers: {}, body: "Hi", parentFolderId: "INBOX" } })),
    inboxFolderId: () => Promise.resolve({ ok: true, data: "INBOX" }),
    createReply: () => (calls.push("createReply"), Promise.resolve({ ok: true, data: { id: "D" } })),
  } as unknown as Graph;
  let procDepsCalled = 0;
  const proc: ProcDeps = { fn: "t", db: f.db, graph, get: (k) => env[k], settings: readSettings((k) => env[k]).settings, ai: () => (calls.push("openai"), Promise.resolve({ ok: true, reply: "Ok." })) };
  const d: CronDeps = { get: (k) => env[k], db: f.db, procDeps: () => (procDepsCalled++, proc) };
  return { f, u, calls, d, procDepsCalled: () => procDepsCalled };
}
const req = (op: unknown, secret: string | null = "SECRET", method = "POST") =>
  new Request("https://f.test/auto-draft-cron", { method, headers: secret ? { "x-cron-secret": secret } : {}, body: method === "POST" ? JSON.stringify({ op }) : undefined });
const orig = console.log;
const quiet = () => { console.log = () => {}; };
const loud = () => { console.log = orig; };

Deno.test("AC-24: 401 without the right x-cron-secret, and nothing is done", async () => {
  const s = await setup();
  quiet();
  try {
    for (const secret of [null, "", "WRONG", "SECRET2", "SECRE"]) {
      assertEquals((await handle(req("process", secret), s.d)).status, 401);
    }
    const noSecretConfigured = await setup({ AUTO_DRAFTS_ENABLED: "true" });
    assertEquals((await handle(req("process", ""), noSecretConfigured.d)).status, 401);
    assertEquals((await handle(req("process", "undefined"), noSecretConfigured.d)).status, 401);
    assertEquals([s.procDepsCalled(), s.calls.length, s.f.calls.length], [0, 0, 0]);
  } finally {
    loud();
  }
});

Deno.test("bad op and wrong method", async () => {
  const s = await setup();
  quiet();
  try {
    assertEquals((await handle(req("nope"), s.d)).status, 400);
    assertEquals((await handle(req(undefined), s.d)).status, 400);
    assertEquals((await handle(req("process", "SECRET", "GET"), s.d)).status, 405);
  } finally {
    loud();
  }
});

Deno.test("AC-9: kill switch off: process and renew do nothing (no Graph, no OpenAI); retention still runs", async () => {
  const s = await setup({ CRON_SHARED_SECRET: "SECRET", AUTO_DRAFTS_ENABLED: "false" });
  quiet();
  try {
    for (const op of ["process", "renew"]) {
      const res = await handle(req(op), s.d);
      assertEquals([res.status, (await res.json()).skipped], [200, "kill_switch"]);
    }
    assertEquals([s.calls.length, s.procDepsCalled(), s.f.queue.length], [0, 0, 0]);
    assertEquals((await handle(req("retention"), s.d)).status, 200);
    assertEquals(s.f.calls, ["retention"]);
  } finally {
    loud();
  }
});

Deno.test("AC-5/7/8/40: process creates the subscription, catches up, then processes the queue to a draft", async () => {
  const s = await setup();
  quiet();
  try {
    s.u.status = "pending";
    const res = await handle(req("process"), s.d);
    assertEquals(res.status, 200);
    assertEquals(s.calls, ["create", "list", "readFrom", "readMessage", "openai", "createReply"]);
    assertEquals(s.u.status, "active");
    assertEquals(s.f.queue.map((r) => [r.messageId, r.status]), [["m1", "drafted"]]);
    // a second run: the id is found again but nothing is drafted twice
    s.calls.length = 0;
    await handle(req("process"), s.d);
    assertFalse(s.calls.includes("createReply"));
    assertEquals(s.f.queue.length, 1);
  } finally {
    loud();
  }
});

Deno.test("renew: disabled users lose their subscription", async () => {
  const s = await setup();
  quiet();
  try {
    await s.f.db.setSubscription(s.u.id, "S1", "2026-12-01T00:00:00Z");
    s.u.enabled = false;
    assertEquals((await handle(req("renew"), s.d)).status, 200);
    assertEquals(s.calls, ["delete"]);
    assertEquals(s.f.subs.size, 0);
  } finally {
    loud();
  }
});
