import { assert, assertEquals, assertFalse } from "@std/assert";
import { fakeDb } from "../_shared/testing.ts";
import { handle, type WebhookDeps } from "./handler.ts";

const OID = "33333333-3333-3333-3333-333333333333";
function setup(env: Record<string, string> = { AUTO_DRAFTS_ENABLED: "true", GRAPH_CLIENT_STATE: "CS" }) {
  const f = fakeDb();
  const u = f.addUser(OID, ["a@x.com"]);
  void f.db.setSubscription(u.id, "SUB1", "2026-11-01T00:00:00Z");
  const waits: Promise<unknown>[] = [];
  const events: string[] = [];
  let pending: () => void = () => {};
  const d: WebhookDeps = {
    get: (k) => env[k],
    db: f.db,
    waitUntil: (p) => waits.push(p),
    runProcess: () => (events.push("process"), new Promise((r) => (pending = () => r(null)))), // stays pending until released
    runLifecycle: (_userId, subId, kind) => (events.push(`lifecycle:${kind}:${subId}`), Promise.resolve()),
  };
  return { f, u, d, waits, events, release: () => pending() };
}
const post = (body: unknown, qs = "") => new Request(`https://f.test/graph-webhook${qs}`, { method: "POST", body: typeof body === "string" ? body : JSON.stringify(body) });
const n = (id: string, over: Record<string, unknown> = {}) => ({ subscriptionId: "SUB1", clientState: "CS", changeType: "created", resourceData: { id }, ...over });
let orig: typeof console.log;
const quiet = (out: string[] = []) => { orig = console.log; console.log = (s: string) => out.push(s); return out; };
const loud = () => { console.log = orig; };

Deno.test("AC-1: validation handshake echoes the decoded token as text/plain and does nothing else", async () => {
  const s = setup({});
  s.d.db = new Proxy({}, { get: () => () => { throw new Error("db must not be touched"); } }) as never;
  quiet();
  try {
    const res = await handle(post("", "?validationToken=abc%20def%2B%3D%2F"), s.d);
    assertEquals(res.status, 200);
    assertEquals(res.headers.get("Content-Type"), "text/plain");
    assertEquals(await res.text(), "abc def+=/");
    assertEquals(s.waits.length, 0);
  } finally {
    loud();
  }
});

Deno.test("AC-2: wrong clientState or unknown subscription is ignored, still 202, nothing inserted", async () => {
  const s = setup();
  const out = quiet();
  try {
    const res = await handle(post({ value: [n("m1", { clientState: "bad" }), n("m2", { subscriptionId: "NOPE" }), n("m3", { clientState: undefined }), "junk", n("m4", { subscriptionId: undefined })] }), s.d);
    assertEquals(res.status, 202);
    assertEquals(s.f.queue.length, 0);
    assertEquals(s.waits.length, 0);
    const reasons = out.map((l) => JSON.parse(l).reason);
    assertEquals(reasons.filter((r) => r === "client_state_mismatch").length, 3);
    assertEquals(reasons.filter((r) => r === "unknown_subscription").length, 2);
  } finally {
    loud();
  }
});

Deno.test("AC-3 and AC-4: valid notifications insert one row per message, 202 before processing finishes, duplicates collapse", async () => {
  const s = setup();
  quiet();
  try {
    const res = await handle(post({ value: [n("m1"), n("m1"), n("m2")] }), s.d); // resolves although runProcess is still pending
    assertEquals(res.status, 202);
    assertEquals(s.f.queue.map((r) => r.messageId), ["m1", "m2"]);
    assertEquals(s.waits.length, 1);
    await Promise.resolve();
    assertEquals(s.events, ["process"]);
    await handle(post({ value: [n("m1")] }), s.d); // again: still two rows
    assertEquals(s.f.queue.length, 2);
    s.release();
  } finally {
    loud();
  }
});

Deno.test("AC-3: an insert failure returns 503 so Graph retries, and no processing starts", async () => {
  const s = setup();
  s.f.failEnqueue(true);
  quiet();
  try {
    const res = await handle(post({ value: [n("m1")] }), s.d);
    assertEquals(res.status, 503);
    assertEquals(s.waits.length, 0);
  } finally {
    loud();
  }
});

Deno.test("AC-3: only created changes with a resource id are queued", async () => {
  const s = setup();
  quiet();
  try {
    await handle(post({ value: [n("m1", { changeType: "updated" }), n("m2", { resourceData: {} }), n("m3", { resourceData: undefined })] }), s.d);
    assertEquals(s.f.queue.length, 0);
  } finally {
    loud();
  }
});

Deno.test("AC-9: kill switch off acknowledges and inserts nothing; validation still works", async () => {
  for (const v of [undefined, "false", "TRUE"]) {
    const s = setup({ GRAPH_CLIENT_STATE: "CS", ...(v ? { AUTO_DRAFTS_ENABLED: v } : {}) });
    quiet();
    try {
      assertEquals((await handle(post({ value: [n("m1")] }), s.d)).status, 202);
      assertEquals([s.f.queue.length, s.waits.length, s.f.calls.length], [0, 0, 0]);
      assertEquals(await (await handle(post("", "?validationToken=t"), s.d)).text(), "t");
    } finally {
      loud();
    }
  }
});

Deno.test("AC-5: lifecycle notifications get 202 and are routed; bad clientState is ignored", async () => {
  const s = setup();
  quiet();
  try {
    const res = await handle(post({ value: [
      { subscriptionId: "SUB1", clientState: "CS", lifecycleEvent: "reauthorizationRequired" },
      { subscriptionId: "SUB1", clientState: "CS", lifecycleEvent: "subscriptionRemoved" },
      { subscriptionId: "SUB1", clientState: "CS", lifecycleEvent: "missed" },
      { subscriptionId: "SUB1", clientState: "bad", lifecycleEvent: "missed" },
    ] }), s.d);
    assertEquals(res.status, 202);
    assertEquals(s.f.queue.length, 0);
    await Promise.resolve();
    await new Promise((r) => setTimeout(r, 0));
    assertEquals(s.events.filter((e) => e.startsWith("lifecycle")), ["lifecycle:reauthorizationRequired:SUB1", "lifecycle:subscriptionRemoved:SUB1", "lifecycle:missed:SUB1"]);
    s.release();
  } finally {
    loud();
  }
});

Deno.test("method and body errors; logs hold only fixed fields", async () => {
  const s = setup();
  const out = quiet();
  try {
    assertEquals((await handle(new Request("https://f.test/graph-webhook"), s.d)).status, 405);
    assertEquals((await handle(post("{not json"), s.d)).status, 400);
    await handle(post({ value: [n("SECRET-MSG-ID")] }), s.d);
    s.release();
  } finally {
    loud();
  }
  for (const l of out) {
    assertFalse(l.includes("SECRET-MSG-ID") || l.includes("CS") && l.includes("clientState"));
    assert(Object.keys(JSON.parse(l)).every((k) => ["fn", "id", "method", "status", "reason", "count", "ms"].includes(k)));
  }
});

Deno.test("AC-24: no database configured gives 503 for notifications (Graph retries)", async () => {
  const s = setup();
  s.d.db = null;
  quiet();
  try {
    assertEquals((await handle(post({ value: [n("m1")] }), s.d)).status, 503);
  } finally {
    loud();
  }
});
