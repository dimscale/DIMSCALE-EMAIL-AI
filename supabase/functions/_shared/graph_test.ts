import { assert, assertEquals, assertRejects } from "@std/assert";
import { makeGraph } from "./graph.ts";

const OID = "33333333-3333-3333-3333-333333333333";
const WEBHOOK = "https://x.supabase.co/functions/v1/graph-webhook";
const NOW = Date.parse("2026-10-23T12:00:00Z");

type Call = { method: string; url: string; headers: Headers; body: unknown };
function setup(responses: (Response | Error)[] | ((c: Call) => Response)) {
  const calls: Call[] = [];
  const fetchFn = ((url: string, init: RequestInit) => {
    const c: Call = { method: init.method!, url, headers: init.headers as Headers, body: init.body ? JSON.parse(init.body as string) : undefined };
    calls.push(c);
    if (typeof responses === "function") return Promise.resolve(responses(c));
    const r = responses.shift()!;
    return r instanceof Error ? Promise.reject(r) : Promise.resolve(r);
  }) as unknown as typeof fetch;
  let invalidated = 0;
  const graph = makeGraph({ getToken: () => Promise.resolve("TOK"), invalidateToken: () => invalidated++, fetchFn, webhookUrl: WEBHOOK, clientState: "STATE", nowMs: () => NOW });
  return { graph, calls, invalidated: () => invalidated };
}
const j = (b: unknown, status = 200, headers: Record<string, string> = {}) => new Response(JSON.stringify(b), { status, headers });

Deno.test("AC-11: readFrom is exactly $select=from, then readMessage uses the exact select list and both Prefer values", async () => {
  const { graph, calls } = setup([
    j({ from: { emailAddress: { address: "ana@exemplo.pt", name: "Ana" } } }),
    j({
      "@odata.type": "#microsoft.graph.message", subject: "S", from: { emailAddress: { address: "ana@exemplo.pt", name: "Ana" } },
      toRecipients: [{ emailAddress: { address: "o@d.com" } }], isDraft: false, inferenceClassification: "focused",
      receivedDateTime: "2026-10-23T11:00:00Z", parentFolderId: "F", uniqueBody: { content: "body" },
      internetMessageHeaders: [{ name: "List-Id", value: "x" }, { name: "list-id", value: "dup" }],
    }),
  ]);
  const id = "AAMk/+=id";
  const f = await graph.readFrom(OID, id);
  assertEquals(f, { ok: true, data: { from: "ana@exemplo.pt" } });
  assertEquals(calls[0].method, "GET");
  assertEquals(calls[0].url, `https://graph.microsoft.com/v1.0/users/${OID}/messages/${encodeURIComponent(id)}?$select=from`);
  assertEquals(calls[0].headers.get("Authorization"), "Bearer TOK");
  const m = await graph.readMessage(OID, id);
  assertEquals(calls[1].url, `https://graph.microsoft.com/v1.0/users/${OID}/messages/${encodeURIComponent(id)}?$select=subject,from,toRecipients,isDraft,inferenceClassification,receivedDateTime,internetMessageHeaders,uniqueBody,parentFolderId`);
  const prefer = calls[1].headers.get("Prefer")!;
  assert(prefer.includes('outlook.body-content-type="text"') && prefer.includes('IdType="ImmutableId"'));
  assert(!calls[1].url.includes("body,") && !calls[1].url.includes("attachments") && !calls[1].url.includes("$value"));
  assert(m.ok);
  assertEquals(m.data.headers, { "list-id": "x" });
  assertEquals([m.data.body, m.data.toAddresses, m.data.isDraft, m.data.fromName], ["body", ["o@d.com"], false, "Ana"]);
});

Deno.test("readMessage fails closed: missing isDraft means draft", async () => {
  const { graph } = setup([j({})]);
  const m = await graph.readMessage(OID, "i");
  assert(m.ok && m.data.isDraft && m.data.body === "");
});

Deno.test("AC-18: createReply is one POST with comment and category, no recipients", async () => {
  const { graph, calls } = setup([j({ id: "DRAFT" }, 201)]);
  const r = await graph.createReply(OID, "m1", "[marker]\n\ntext <b>x</b>");
  assertEquals(r, { ok: true, data: { id: "DRAFT" } });
  assertEquals(calls.length, 1);
  assertEquals(calls[0].method, "POST");
  assertEquals(calls[0].url, `https://graph.microsoft.com/v1.0/users/${OID}/messages/m1/createReply`);
  assertEquals(calls[0].body, { comment: "[marker]\n\ntext <b>x</b>", message: { categories: ["DIMSCALE AI draft"] } });
});

Deno.test("AC-6: subscription creation body and header", async () => {
  const { graph, calls } = setup([j({ id: "S1", expirationDateTime: "2026-10-30T10:40:00Z" }, 201)]);
  const r = await graph.createSubscription(OID);
  assertEquals(r, { ok: true, data: { id: "S1", expiresAt: "2026-10-30T10:40:00Z" } });
  const b = calls[0].body as Record<string, unknown>;
  assertEquals(calls[0].method, "POST");
  assertEquals(calls[0].url, "https://graph.microsoft.com/v1.0/subscriptions");
  assertEquals(Object.keys(b).sort(), ["changeType", "clientState", "expirationDateTime", "lifecycleNotificationUrl", "notificationUrl", "resource"]);
  assertEquals([b.changeType, b.resource, b.notificationUrl, b.lifecycleNotificationUrl, b.clientState], ["created", `users/${OID}/mailFolders('inbox')/messages`, WEBHOOK, WEBHOOK, "STATE"]);
  assert(Date.parse(b.expirationDateTime as string) - NOW <= 10_080 * 60_000);
  assertEquals(calls[0].headers.get("Prefer"), 'IdType="ImmutableId"');
});

Deno.test("AC-6: 409 adopts the existing subscription", async () => {
  const { graph, calls } = setup([
    new Response("{}", { status: 409 }),
    j({ value: [{ id: "OTHER", resource: "users/x/mailFolders('inbox')/messages", notificationUrl: WEBHOOK }, { id: "MINE", resource: `Users/${OID}/mailFolders('inbox')/messages`, notificationUrl: WEBHOOK, expirationDateTime: "2026-10-28T00:00:00Z" }] }),
  ]);
  assertEquals(await graph.createSubscription(OID), { ok: true, data: { id: "MINE", expiresAt: "2026-10-28T00:00:00Z" } });
  assertEquals([calls[1].method, calls[1].url], ["GET", "https://graph.microsoft.com/v1.0/subscriptions"]);
});

Deno.test("subscription renew (PATCH) and delete (404 is fine)", async () => {
  const { graph, calls } = setup([j({ expirationDateTime: "2026-11-01T00:00:00Z" }), new Response(null, { status: 204 }), new Response("{}", { status: 404 }), new Response("{}", { status: 500 })]);
  assertEquals(await graph.renewSubscription("S1"), { ok: true, data: { expiresAt: "2026-11-01T00:00:00Z" } });
  assertEquals([calls[0].method, calls[0].url], ["PATCH", "https://graph.microsoft.com/v1.0/subscriptions/S1"]);
  assertEquals(Object.keys(calls[0].body as object), ["expirationDateTime"]);
  assertEquals((await graph.deleteSubscription("S1")).ok, true);
  assertEquals(calls[1].method, "DELETE");
  assertEquals((await graph.deleteSubscription("S2")).ok, true);
  assertEquals((await graph.deleteSubscription("S3")).ok, false);
});

Deno.test("AC-8: catch-up uses $select=id, pages, never asks for bodies", async () => {
  const { graph, calls } = setup([
    j({ value: [{ id: "a" }, { id: "b" }], "@odata.nextLink": `https://graph.microsoft.com/v1.0/users/${OID}/mailFolders('inbox')/messages?$skiptoken=zz` }),
    j({ value: [{ id: "c" }] }),
  ]);
  assertEquals(await graph.listInboxIds(OID, "2026-10-23T00:00:00.000Z"), { ok: true, data: ["a", "b", "c"] });
  assertEquals(calls[0].url, `https://graph.microsoft.com/v1.0/users/${OID}/mailFolders('inbox')/messages?$filter=receivedDateTime%20ge%202026-10-23T00%3A00%3A00.000Z&$select=id&$top=50`);
  assertEquals(calls.length, 2);
  assert(calls.every((c) => !c.url.includes("body")));
});

Deno.test("a nextLink pointing elsewhere is refused (no request is made)", async () => {
  const { graph, calls } = setup([j({ value: [], "@odata.nextLink": "https://evil.example/v1.0/x" })]);
  await assertRejects(() => graph.listInboxIds(OID, "2026-10-23T00:00:00.000Z"), Error, "graph_url_not_allowed");
  assertEquals(calls.length, 1);
});

Deno.test("AC-21: 429 and 503 carry Retry-After, 401 drops the cached token, network and token errors have reasons", async () => {
  const { graph, invalidated } = setup([j({}, 429, { "Retry-After": "17" }), j({}, 503, { "Retry-After": "5" }), j({}, 500), j({}, 401), new Error("net")]);
  assertEquals(await graph.readFrom(OID, "i"), { ok: false, status: 429, retryAfterSec: 17 });
  assertEquals(await graph.readFrom(OID, "i"), { ok: false, status: 503, retryAfterSec: 5 });
  assertEquals(await graph.readFrom(OID, "i"), { ok: false, status: 500, retryAfterSec: undefined });
  await graph.readFrom(OID, "i");
  assertEquals(invalidated(), 1);
  assertEquals(await graph.readFrom(OID, "i"), { ok: false, status: 0, reason: "network" });
  const g2 = makeGraph({ getToken: () => Promise.reject(new Error("token_unavailable")), webhookUrl: WEBHOOK, clientState: "S" });
  assertEquals(await g2.readFrom(OID, "i"), { ok: false, status: 0, reason: "token_unavailable" });
});

// ---------- AC-19: the client cannot call anything but its allow-list ----------
Deno.test("AC-19: source contains no send/reply/forward/move/copy/delete-message path", async () => {
  const root = new URL("..", import.meta.url);
  const files: URL[] = [];
  async function walk(dir: URL) {
    for await (const e of Deno.readDir(dir)) {
      const u = new URL(e.name + (e.isDirectory ? "/" : ""), dir);
      if (e.isDirectory) await walk(u);
      else if (e.name.endsWith(".ts") && !e.name.endsWith("_test.ts")) files.push(u);
    }
  }
  await walk(root);
  assert(files.length > 10);
  const forbidden = [/\/send/i, /\/reply\b/i, /replyall/i, /forward/i, /sendmail/i, /\/move\b/i, /\/copy\b/i, /Mail\.Send/i];
  for (const f of files) {
    const text = await Deno.readTextFile(f);
    for (const re of forbidden) assert(!re.test(text), `${f.pathname} matches ${re}`);
    // DELETE and PATCH: graph.ts (subscriptions only) and the settings function (its own HTTP verbs, no Graph).
    if (!f.pathname.endsWith("/graph.ts") && !f.pathname.endsWith("/auto-reply-settings/handler.ts")) assert(!/["']DELETE["']|["']PATCH["']/.test(text), `${f.pathname} uses DELETE/PATCH`);
  }
});

Deno.test("AC-19: the client refuses forbidden calls even if asked (no fetch happens)", async () => {
  // Reach the private call() through a graph built on a fetch that fails the test.
  const g = makeGraph({ getToken: () => Promise.resolve("T"), fetchFn: (() => { throw new Error("fetch must not run"); }) as unknown as typeof fetch, webhookUrl: WEBHOOK, clientState: "S" });
  // Public methods only build allow-listed URLs; a malformed oid or id cannot escape the path patterns.
  await assertRejects(() => g.readFrom("not-a-guid", "x"), Error, "graph_call_not_allowed");
  await assertRejects(() => g.readFrom(`${OID}/../..`, "x"), Error, "graph_call_not_allowed");
  // An id containing a slash is URL-encoded, so it stays one path segment (here the stubbed fetch throws, reported as a network error).
  assertEquals(await g.createReply(OID, "x/send", "c"), { ok: false, status: 0, reason: "network" });
});
