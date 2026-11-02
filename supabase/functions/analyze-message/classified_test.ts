import { assert, assertEquals, assertFalse } from "@std/assert";
import { buildInput } from "../_shared/ai.ts";
import { CLASSIFIED_INSTRUCTIONS, CLASSIFIED_SCHEMA, generateClassified, parseClassified } from "./classified.ts";
import { SUMMARY_INSTRUCTIONS } from "./summary.ts";

const DI = { subject: "Fatura", fromName: "Ana", fromDomain: "exemplo.pt", body: "Pague em https://exemplo.pt/pay ou ana@exemplo.pt" };
const INPUT = buildInput(DI);
const good = { points: ["a", "b", "c"], category: "action_required", priority: "high", tone: "neutral", reason: "O remetente pede um pagamento.", confidence: "medium", deadline: "sexta", requested_action: null };
const wrap = (t: string) => ({ status: "completed", output: [{ type: "message", content: [{ type: "output_text", text: t }] }] });
const out = (o: unknown) => wrap(JSON.stringify(o));
const invalid = { ok: false, reason: "ai_output_invalid", retryable: false } as const;
const run = (o: unknown) => parseClassified(o, INPUT);

Deno.test("F-007/AC-4: classified schema snapshot (enums fail the test when changed)", () => {
  assertEquals(JSON.parse(JSON.stringify(CLASSIFIED_SCHEMA)), {
    type: "object", additionalProperties: false,
    required: ["points", "category", "priority", "tone", "reason", "confidence", "deadline", "requested_action"],
    properties: {
      points: { type: "array", minItems: 3, maxItems: 3, items: { type: "string" } },
      category: { type: "string", enum: ["action_required", "information", "meeting", "invoice_document", "marketing", "automated_notice", "other"] },
      priority: { type: "string", enum: ["high", "normal", "low"] },
      tone: { type: "string", enum: ["positive", "neutral", "negative"] },
      reason: { type: "string" },
      confidence: { type: "string", enum: ["high", "medium", "low"] },
      deadline: { type: ["string", "null"] },
      requested_action: { type: ["string", "null"] },
    },
  });
});

Deno.test("F-007/AC-3: classified instructions replace only the tags sentence", () => {
  assertFalse(CLASSIFIED_INSTRUCTIONS.includes("Choose type, priority and tone"));
  assert(CLASSIFIED_INSTRUCTIONS.includes("exactly one category from the enum"));
  assert(CLASSIFIED_INSTRUCTIONS.includes("untrusted data"));
  assert(CLASSIFIED_INSTRUCTIONS.includes("Do not use the characters < or >."));
  assert(CLASSIFIED_INSTRUCTIONS.includes("Never state or imply that the email or the sender is safe, legitimate, verified or trustworthy; the reason describes the content only."));
  assertEquals(CLASSIFIED_INSTRUCTIONS.replace(/Classify the email into .* the reason describes the content only\./, "Choose type, priority and tone from the allowed values."), SUMMARY_INSTRUCTIONS);
});

Deno.test("F-007/AC-3: classified request is exact and input equals the default input", async () => {
  let seen: { url: string; init: RequestInit } | undefined;
  const f = ((url: string, init: RequestInit) => (seen = { url, init }, Promise.resolve(new Response(JSON.stringify(out(good)))))) as unknown as typeof fetch;
  const cfg = { apiKey: "KEY", model: "gpt-6-luna", maxOutputTokens: 2000, reasoningEffort: "low" };
  const di = { ...DI, fromName: "Ana Silva", body: "Olá. Pague em https://exemplo.pt/pay" };
  const r = await generateClassified({ ...di, recipients: ["x@y.pt"], oid: "OID-1" } as typeof di, cfg, f);
  assert(r.ok);
  assert(seen!.init.signal);
  const raw = seen!.init.body as string;
  const b = JSON.parse(raw);
  assertEquals(Object.keys(b).sort(), ["input", "instructions", "max_output_tokens", "model", "reasoning", "store", "text"]);
  assertFalse("tools" in b);
  assertEquals([b.model, b.store, b.max_output_tokens, b.reasoning], ["gpt-6-luna", false, 2000, { effort: "low" }]);
  assertEquals(b.instructions, CLASSIFIED_INSTRUCTIONS);
  assertEquals(b.input, buildInput(di));
  assertEquals(b.text.format, { type: "json_schema", name: "email_classification", strict: true, schema: JSON.parse(JSON.stringify(CLASSIFIED_SCHEMA)) });
  for (const v of ["x@y.pt", "OID-1", "KEY"]) assertFalse(raw.includes(v), v);
});

Deno.test("F-007/AC-5: valid fixture maps requested_action and trims reason", () => {
  assertEquals(run(out(good)), { ok: true, summary: { points: ["a", "b", "c"], category: "action_required", priority: "high", tone: "neutral", reason: "O remetente pede um pagamento.", confidence: "medium", deadline: "sexta", requestedAction: null } });
  const r = run(out({ ...good, reason: "  ok  ", deadline: null, requested_action: "Pagar" }));
  assert(r.ok && r.summary.reason === "ok" && r.summary.requestedAction === "Pagar");
  assert(run(out({ ...good, reason: "x".repeat(160) })).ok);
  for (const c of ["action_required", "information", "meeting", "invoice_document", "marketing", "automated_notice", "other"]) assert(run(out({ ...good, category: c })).ok, c);
});

Deno.test("F-007/AC-5: every failure is ai_output_invalid", () => {
  const msg = (content: unknown[]) => ({ status: "completed", output: [{ type: "message", content }] });
  const ot = { type: "output_text", text: JSON.stringify(good) };
  const { category: _c, ...noCat } = good;
  const cases: unknown[] = [
    { ...out(good), status: "incomplete" },
    msg([{ type: "refusal", refusal: "no" }]), msg([ot, ot]), msg([]),
    wrap("not json"), wrap("[]"), wrap("null"),
    out({ ...noCat, type: "request" }), out(noCat), out({ ...good, type: "request" }), out({ ...good, extra: 1 }),
    out({ ...good, reason: undefined }), out({ ...good, confidence: undefined }),
    out({ ...good, points: ["a", "b"] }), out({ ...good, points: ["a", "b", "c", "d"] }),
    out({ ...good, points: ["a", " ", "c"] }), out({ ...good, points: ["a", "b", "x".repeat(301)] }), out({ ...good, points: ["a", "b", 3] }),
    ...["urgent", "High", "", "request", 1, null].map((v) => out({ ...good, category: v })),
    ...["urgent", "High", ""].flatMap((v) => [out({ ...good, priority: v }), out({ ...good, tone: v }), out({ ...good, confidence: v })]),
    out({ ...good, confidence: "none" }),
    out({ ...good, reason: "" }), out({ ...good, reason: "   " }), out({ ...good, reason: "x".repeat(161) }), out({ ...good, reason: 5 }),
    out({ ...good, reason: "linha\nnova" }), out({ ...good, reason: "linha\r" }), out({ ...good, reason: "a\tb" }), out({ ...good, reason: "a\u0000b" }), out({ ...good, reason: "\nok" }),
    out({ ...good, reason: "<b>x</b>" }), out({ ...good, reason: "a > b" }),
    out({ ...good, reason: "Ver https://evil.example/x" }), out({ ...good, reason: "Ver www.evil.example" }),
    out({ ...good, reason: "Escreva para boss@evil.example" }),
    out({ ...good, reason: "Pague em https://exemplo.pt/pay" }), // a URL that IS in the email
    out({ ...good, reason: "Escreva a ana@exemplo.pt" }), // an address that IS in the email
    out({ ...good, deadline: "" }), out({ ...good, deadline: "x".repeat(201) }), out({ ...good, requested_action: "" }), out({ ...good, requested_action: "x".repeat(201) }),
    out({ ...good, points: ["<b>x</b>", "b", "c"] }), out({ ...good, deadline: "a<b" }), out({ ...good, requested_action: "a>b" }),
    out({ ...good, points: ["a", "b", "http://evil.example"] }), out({ ...good, requested_action: "ver //evil.example/x" }),
    null, "x",
  ];
  for (const c of cases) assertEquals(run(c), invalid);
});

Deno.test("F-007/AC-5: transport failures map like F-003", async () => {
  const st = (s: number) => generateClassified(DI, { apiKey: "K", model: "m", maxOutputTokens: 1 }, (() => Promise.resolve(new Response("{}", { status: s }))) as unknown as typeof fetch);
  assertEquals(await st(503), { ok: false, reason: "ai_unavailable", retryable: true });
  assertEquals(await st(401), { ok: false, reason: "ai_unavailable", retryable: false });
});
