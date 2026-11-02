// QA (adversarial) for F-013 amendment AC-12 (sender before content), AC-45 (auth_failed ordering), AC-47, AC-50.
// Mocked Graph, mocked OpenAI, in-memory fakeDb. Helpers copied from processor_test.ts (they are not exported there).
import { assertEquals, assertFalse } from "@std/assert";
import type { AiResult } from "./ai.ts";
import { readSettings } from "./config.ts";
import type { Graph, GResult, Msg } from "./graph.ts";
import { processClaimed, type ProcDeps } from "./processor.ts";
import { sha256Hex } from "./safe.ts";
import { fakeDb } from "./testing.ts";

const OWN = "owner@dimscale.com";
const NOW0 = Date.parse("2026-10-23T12:00:00Z");
const ownSha = await sha256Hex(OWN);
const okMsg: Msg = {
  odataType: "#microsoft.graph.message", subject: "Quote", fromAddress: "Ana@Exemplo.pt", fromName: "Ana", toAddresses: [OWN], isDraft: false,
  inference: "focused", receivedDateTime: "2026-10-23T11:30:00Z", headers: {}, body: "Please send the quote.", parentFolderId: "INBOX",
};

type G = { from?: GResult<{ from: string }>; msg?: GResult<Msg> };
function mockGraph(o: G = {}) {
  const calls: string[] = [];
  const graph = {
    readFrom: (_o: string, id: string) => (calls.push(`readFrom:${id}`), Promise.resolve(o.from ?? { ok: true, data: { from: "ANA@exemplo.pt" } })),
    readMessage: (_o: string, id: string) => (calls.push(`readMessage:${id}`), Promise.resolve(o.msg ?? { ok: true, data: okMsg })),
    inboxFolderId: () => (calls.push("inboxFolderId"), Promise.resolve({ ok: true, data: "INBOX" })),
    createReply: (_o: string, id: string) => (calls.push(`createReply:${id}`), Promise.resolve({ ok: true, data: { id: "D" } })),
  } as unknown as Graph;
  return { graph, calls };
}

function setup(g: G = {}, env: Record<string, string> = {}) {
  const now = NOW0;
  const f = fakeDb(() => now);
  const user = f.addUser("33333333-3333-3333-3333-333333333333", ["ana@exemplo.pt"], ownSha);
  const mg = mockGraph(g);
  const aiCalls: unknown[] = [];
  const environment: Record<string, string> = { AUTO_DRAFTS_ENABLED: "true", OPENAI_API_KEY: "K", OPENAI_MODEL: "m", ...env };
  const deps: ProcDeps = {
    fn: "t", db: f.db, graph: mg.graph, get: (k) => environment[k], nowMs: () => now,
    settings: readSettings((k) => environment[k]).settings,
    ai: (i): Promise<AiResult> => (aiCalls.push(i), Promise.resolve({ ok: true, reply: "Thanks." })),
  };
  return { f, user, deps, calls: mg.calls, aiCalls };
}
const enqueue = (s: ReturnType<typeof setup>, ...ids: string[]) => s.f.db.enqueue(s.user.id, ids);
const row = (s: ReturnType<typeof setup>, i = 0) => s.f.queue[i];
const withHeaders = (headers: Record<string, string>, over: Partial<Msg> = {}): G => ({ msg: { ok: true, data: { ...okMsg, headers, ...over } } });

Deno.test("ADV F-013/AC-12: a sender that is not allowed ends sender_not_allowed after the from-only read, even with dmarc=fail present", async () => {
  const s = setup({ from: { ok: true, data: { from: "impostor@exemplo.pt" } }, ...withHeaders({ "authentication-results": "dmarc=fail" }) });
  await enqueue(s, "m1");
  await processClaimed(s.deps);
  assertEquals([row(s).status, row(s).reason], ["skipped", "sender_not_allowed"]);
  assertEquals(s.calls, ["readFrom:m1"], "no full read, no inbox lookup");
  assertEquals(s.aiCalls.length, 0);
});

Deno.test("ADV F-013/AC-45: a second Authentication-Results header that fails skips with auth_failed; no AI, no draft", async () => {
  const s = setup(withHeaders({ "authentication-results": "dmarc=pass\ncompauth=fail reason=000" }));
  await enqueue(s, "m1");
  await processClaimed(s.deps);
  assertEquals([row(s).status, row(s).reason], ["skipped", "auth_failed"]);
  assertEquals(s.aiCalls.length, 0);
  assertFalse(s.calls.some((c) => c.startsWith("createReply")));
});

Deno.test("ADV F-013/AC-45: auth_failed comes before too_old and empty_body (a message that is also old and empty)", async () => {
  const s = setup(withHeaders({ "authentication-results": "dmarc=fail" }, { receivedDateTime: "2026-10-20T00:00:00Z", body: "  " }));
  await enqueue(s, "m1");
  await processClaimed(s.deps);
  assertEquals([row(s).status, row(s).reason], ["skipped", "auth_failed"]);
});

Deno.test("ADV F-013/AC-44/45: an earlier filter keeps its own code (auto-submitted beats auth_failed)", async () => {
  const s = setup(withHeaders({ "authentication-results": "dmarc=fail", "auto-submitted": "auto-replied" }));
  await enqueue(s, "m1");
  await processClaimed(s.deps);
  assertEquals([row(s).status, row(s).reason], ["skipped", "auto_submitted"]);
});

Deno.test("ADV F-013/AC-45 config: FILTER_AUTH_FAILED=false lets the same message reach the AI and a draft", async () => {
  const s = setup(withHeaders({ "authentication-results": "dmarc=fail" }), { FILTER_AUTH_FAILED: "false" });
  await enqueue(s, "m1");
  await processClaimed(s.deps);
  assertEquals([row(s).status, row(s).reason], ["drafted", "ok"]);
  assertEquals(s.aiCalls.length, 1);
});

Deno.test("ADV F-013/AC-46: look-alike and partial tokens and spf/dkim failures reach the AI (fail open)", async () => {
  for (const v of ["dmarc=failx", "xdmarc=fail", "spf=fail", "dkim=fail", "dmarc=pass; spf=fail", "dmarc=none", "compauth=softpass"]) {
    const s = setup(withHeaders({ "authentication-results": v }));
    await enqueue(s, "m1");
    await processClaimed(s.deps);
    assertEquals([row(s).status, row(s).reason], ["drafted", "ok"], v);
  }
});
