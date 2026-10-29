import { assert, assertEquals, assertFalse } from "@std/assert";
import { buildInput } from "../_shared/ai.ts";
import { generateTranslation, parseTranslation, TRANSLATE_INSTRUCTIONS, TRANSLATE_SCHEMA } from "./translate.ts";

const DI = { subject: "Fatura", fromName: "", fromDomain: "", body: "Pague em https://exemplo.pt/pay ou ana@exemplo.pt" };
const good = { subject: "Invoice", body: "Pay at [link] or [address]" };
const wrap = (t: string) => ({ status: "completed", output: [{ type: "message", content: [{ type: "output_text", text: t }] }] });
const out = (o: unknown) => wrap(JSON.stringify(o));
const invalid = { ok: false, reason: "ai_output_invalid", retryable: false } as const;

Deno.test("AC-37: exact translate request, no sender data", async () => {
  let init: RequestInit | undefined;
  const f = ((_u: string, i: RequestInit) => (init = i, Promise.resolve(new Response(JSON.stringify(out(good)))))) as unknown as typeof fetch;
  const r = await generateTranslation(DI, "en", { apiKey: "sk-test", model: "m", maxOutputTokens: 4000, reasoningEffort: "low" }, f);
  assertEquals(r, { ok: true, translation: good });
  const b = JSON.parse(init!.body as string);
  assertEquals(b, {
    model: "m", instructions: TRANSLATE_INSTRUCTIONS("en"), input: buildInput(DI), store: false, max_output_tokens: 4000, reasoning: { effort: "low" },
    text: { format: { type: "json_schema", name: "email_translation", strict: true, schema: TRANSLATE_SCHEMA } },
  });
  assertFalse("tools" in b);
  assert(b.instructions.includes("into English") && TRANSLATE_INSTRUCTIONS("pt").includes("into European Portuguese (pt-PT)"));
  assert(b.input.includes("From name: \n") && init!.signal);
});

Deno.test("AC-38: valid, and every failure is ai_output_invalid", () => {
  assertEquals(parseTranslation(out(good)), { ok: true, translation: good });
  assert(parseTranslation(out({ subject: "", body: "x" })).ok);
  const ot = { type: "output_text", text: JSON.stringify(good) };
  const msg = (content: unknown[]) => ({ status: "completed", output: [{ type: "message", content }] });
  const cases: unknown[] = [
    { ...out(good), status: "incomplete" }, msg([{ type: "refusal", refusal: "no" }]), msg([]), msg([ot, ot]), wrap("{"), wrap("[]"),
    out({ subject: "a" }), out({ body: "a" }), out({ ...good, extra: 1 }), out({ subject: 1, body: "x" }), out({ subject: "a", body: null }),
    out({ ...good, body: "  " }), out({ ...good, body: "" }), out({ ...good, body: "x".repeat(6001) }), out({ ...good, subject: "x".repeat(301) }),
    out({ ...good, body: "a <b>" }), out({ ...good, subject: "a>b" }),
    out({ ...good, body: "see https://exemplo.pt/pay" }), out({ ...good, body: "ana@exemplo.pt" }), out({ ...good, subject: "www.exemplo.pt" }), out({ ...good, body: "go //exemplo.pt" }),
  ];
  for (const c of cases) assertEquals(parseTranslation(c), invalid);
  assert(parseTranslation(out({ subject: "x".repeat(300), body: "x".repeat(6000) })).ok);
});
