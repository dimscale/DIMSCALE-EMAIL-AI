import { assert, assertEquals } from "@std/assert";
import { buildInput, generateDraft, INSTRUCTIONS, SCHEMA } from "../_shared/ai.ts";
import { generateReply, LANGUAGE_SENTENCES, LANGUAGES, TONE_OPTIONS, TONE_SENTENCES } from "./replyOptions.ts";

const DI = { subject: "Fatura", fromName: "Ana", fromDomain: "exemplo.pt", body: "Olá, pode pagar?" };
const CFG = { apiKey: "sk-test", model: "m", maxOutputTokens: 2000, reasoningEffort: "low" };
const reply = new Response(JSON.stringify({ status: "completed", output: [{ type: "message", content: [{ type: "output_text", text: '{"reply":"Ok."}' }] }] }));
const capture = async (run: (f: typeof fetch) => Promise<unknown>) => {
  let body = "";
  await run(((_u: string, init: RequestInit) => (body = init.body as string, Promise.resolve(reply.clone()))) as unknown as typeof fetch);
  return body;
};

Deno.test("AC-36: match + none is byte-identical to generateDraft", async () => {
  const a = await capture((f) => generateDraft(DI, CFG, f));
  assertEquals(await capture((f) => generateReply(DI, { language: "match", tone: "none" }, CFG, f)), a);
});

Deno.test("AC-36: every option adds only the fixed sentences, language before tone; input and schema unchanged", async () => {
  const base = JSON.parse(await capture((f) => generateDraft(DI, CFG, f)));
  for (const language of LANGUAGES) {
    for (const tone of TONE_OPTIONS) {
      if (language === "match" && tone === "none") continue;
      const b = JSON.parse(await capture((f) => generateReply(DI, { language, tone }, CFG, f)));
      assertEquals(b.instructions, [INSTRUCTIONS, (LANGUAGE_SENTENCES as Record<string, string>)[language], (TONE_SENTENCES as Record<string, string>)[tone]].filter(Boolean).join(" "));
      assertEquals([b.input, b.text, b.max_output_tokens, b.reasoning], [buildInput(DI), base.text, 2000, { effort: "low" }]);
      assertEquals(b.text.format.schema, SCHEMA);
    }
  }
  assertEquals(LANGUAGE_SENTENCES.pt, "Write the reply in European Portuguese (pt-PT), whatever the language of the email. This replaces the language rule above.");
  assertEquals(LANGUAGE_SENTENCES.en, "Write the reply in English, whatever the language of the email. This replaces the language rule above.");
  assertEquals(TONE_SENTENCES.concise, "Make the reply concise: at most three short sentences besides the greeting and the closing.");
  assertEquals(TONE_SENTENCES.detailed, "Make the reply detailed instead of short: answer each point or question the sender raised.");
  assertEquals(TONE_SENTENCES.formal, "Use a formal, courteous register.");
  assertEquals(TONE_SENTENCES.friendly, "Use a warm, friendly register while staying professional.");
  assert(Object.keys(TONE_SENTENCES).length === 4 && Object.keys(LANGUAGE_SENTENCES).length === 2);
});
