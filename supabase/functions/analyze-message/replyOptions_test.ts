import { assert, assertEquals, assertFalse } from "@std/assert";
import { buildInput, generateDraft, INSTRUCTIONS, SCHEMA } from "../_shared/ai.ts";
import { baseInstructions, generateReply, LANGUAGE_PHRASE, LANGUAGE_REMINDERS, LANGUAGES, NO_DOMAINS_SENTENCE, NO_SIGNOFF_SENTENCE, PARAGRAPH_SENTENCE, REFINE_NO_SIGNOFF_SENTENCE, REVISE_SENTENCE, TONE_OPTIONS, TONE_SENTENCES } from "./replyOptions.ts";
import { hasLink } from "../_shared/ai.ts";

const DI = { subject: "Fatura", fromName: "Ana", fromDomain: "exemplo.pt", body: "Olá, pode pagar?" };
const CFG = { apiKey: "sk-test", model: "m", maxOutputTokens: 2000, reasoningEffort: "low" };
const reply = new Response(JSON.stringify({ status: "completed", output: [{ type: "message", content: [{ type: "output_text", text: '{"reply":"Ok."}' }] }] }));
const REM = LANGUAGE_REMINDERS as Record<string, string>;
const tail = (language: string) => REM[language] ? [REM[language]] : [];
const capture = async (run: (f: typeof fetch) => Promise<unknown>) => {
  let body = "";
  await run(((_u: string, init: RequestInit) => (body = init.body as string, Promise.resolve(reply.clone()))) as unknown as typeof fetch);
  return body;
};

Deno.test("AC-36: match + none differs from generateDraft (F-013) only by the manual-reply sentences", async () => {
  const a = JSON.parse(await capture((f) => generateDraft(DI, CFG, f)));
  assertEquals(a.instructions, INSTRUCTIONS); // F-013 request unchanged
  assertFalse(a.instructions.includes(NO_DOMAINS_SENTENCE));
  const b = JSON.parse(await capture((f) => generateReply(DI, { language: "match", tone: "none" }, CFG, f)));
  assertEquals(b, { ...a, instructions: `${INSTRUCTIONS} ${NO_DOMAINS_SENTENCE} ${NO_SIGNOFF_SENTENCE} ${PARAGRAPH_SENTENCE}` });
});

Deno.test("F-003 (g): no sign-off and one-idea-per-paragraph sentences are in every manual reply variant, never in generateDraft", async () => {
  assertEquals(NO_SIGNOFF_SENTENCE, "Do not add a closing line, sign-off or signature and never write the sender's or the user's name at the end; the user's email client adds the signature. Start with a short greeting.");
  assertEquals(PARAGRAPH_SENTENCE, "Put each idea or request in its own short paragraph, separated by a blank line.");
  const UNLESS = "unless the user's instruction explicitly asks for a closing";
  assertEquals(REFINE_NO_SIGNOFF_SENTENCE.startsWith("Do not add a closing line, sign-off or signature " + UNLESS + ", and never write the sender's or the user's name at the end;"), true);
  assert(REFINE_NO_SIGNOFF_SENTENCE.endsWith("the user's email client adds the signature. Start with a short greeting."));
  const draft = JSON.parse(await capture((f) => generateDraft(DI, CFG, f))).instructions as string;
  for (const s of [NO_SIGNOFF_SENTENCE, REFINE_NO_SIGNOFF_SENTENCE, PARAGRAPH_SENTENCE, "sign-off", "short greeting", "blank line"]) assertFalse(draft.includes(s), s);
  for (const language of LANGUAGES) {
    for (const tone of TONE_OPTIONS) {
      const ins = JSON.parse(await capture((f) => generateReply(DI, { language, tone }, CFG, f))).instructions as string;
      assert(ins.includes(NO_SIGNOFF_SENTENCE) && ins.includes(PARAGRAPH_SENTENCE), `${language}/${tone}`);
      assertFalse(ins.includes(UNLESS));
      assertEquals(ins.indexOf(NO_DOMAINS_SENTENCE) < ins.indexOf(NO_SIGNOFF_SENTENCE) && ins.indexOf(NO_SIGNOFF_SENTENCE) < ins.indexOf(PARAGRAPH_SENTENCE), true);
      const ref = JSON.parse(await capture((f) => generateReply(DI, { language, tone }, CFG, f, { instruction: "mais curto", draft: "Olá." }))).instructions as string;
      assert(ref.includes(REFINE_NO_SIGNOFF_SENTENCE) && ref.includes(PARAGRAPH_SENTENCE) && ref.includes(UNLESS), `refine ${language}/${tone}`);
      assertFalse(ref.includes(NO_SIGNOFF_SENTENCE)); // the refine variant replaces, not adds
      assert(ref.endsWith(language === "match" ? REVISE_SENTENCE : REM[language]));
      assertEquals(ref, [baseInstructions(language), (TONE_SENTENCES as Record<string, string>)[tone], NO_DOMAINS_SENTENCE, REFINE_NO_SIGNOFF_SENTENCE, PARAGRAPH_SENTENCE, REVISE_SENTENCE, ...tail(language)].filter(Boolean).join(" "));
    }
  }
});

Deno.test("AC-36: every option adds only the fixed sentences, language before tone; input and schema unchanged", async () => {
  const base = JSON.parse(await capture((f) => generateDraft(DI, CFG, f)));
  for (const language of LANGUAGES) {
    for (const tone of TONE_OPTIONS) {
      const b = JSON.parse(await capture((f) => generateReply(DI, { language, tone }, CFG, f)));
      assertEquals(b.instructions, [baseInstructions(language), (TONE_SENTENCES as Record<string, string>)[tone], NO_DOMAINS_SENTENCE, NO_SIGNOFF_SENTENCE, PARAGRAPH_SENTENCE, ...tail(language)].filter(Boolean).join(" "));
      assertEquals([b.input, b.text, b.max_output_tokens, b.reasoning], [buildInput(DI), base.text, 2000, { effort: "low" }]);
      assertEquals(b.text.format.schema, SCHEMA);
    }
  }
  assertEquals(LANGUAGE_REMINDERS.pt, "The whole reply must be in European Portuguese (pt-PT).");
  assertEquals(LANGUAGE_REMINDERS.en, "The whole reply must be in English.");
  assertEquals(TONE_SENTENCES.concise, "Make the reply concise: at most three short sentences besides the greeting.");
  assertEquals(/closing/i.test(TONE_SENTENCES.concise), false); // no clash with NO_SIGNOFF_SENTENCE (user decision 1)
  assertEquals(TONE_SENTENCES.detailed, "Make the reply detailed instead of short: answer each point or question the sender raised.");
  assertEquals(TONE_SENTENCES.formal, "Use a formal, courteous register.");
  assertEquals(TONE_SENTENCES.friendly, "Use a warm, friendly register while staying professional.");
  assert(Object.keys(TONE_SENTENCES).length === 4 && Object.keys(LANGUAGE_REMINDERS).length === 2);
});

Deno.test("realistic MyEdenred notification reply passes parseOutput and hasLink", async () => {
  const email = { subject: "MyEdenred - Confirme o seu email", fromName: "MyEdenred", fromDomain: "appmyedenred.pt", body: "Confirme o seu email para ativar a conta." };
  const txt = "Boa tarde,\n\nObrigado pelo aviso. Vou confirmar o meu email.\n\nCom os melhores cumprimentos,\nThiago";
  const f = (() => Promise.resolve(new Response(JSON.stringify({ status: "completed", output: [{ type: "message", content: [{ type: "output_text", text: JSON.stringify({ reply: txt }) }] }] })))) as unknown as typeof fetch;
  const r = await generateReply(email, { language: "match", tone: "none" }, CFG, f);
  assertEquals(r, { ok: true, reply: txt });
  assertFalse(hasLink(txt));
});

Deno.test("F-003 (h): en and pt rewrite the base language rule, never contain the old phrase, and end with the reminder; match is unchanged; F-013 draft unchanged", async () => {
  assertEquals(LANGUAGE_PHRASE, "in the language of that email");
  assertEquals(INSTRUCTIONS.split(LANGUAGE_PHRASE).length, 2); // exactly once
  assertEquals(baseInstructions("match"), INSTRUCTIONS);
  assertEquals(baseInstructions("en"), INSTRUCTIONS.replace(LANGUAGE_PHRASE, "in English"));
  assertEquals(baseInstructions("pt"), INSTRUCTIONS.replace(LANGUAGE_PHRASE, "in European Portuguese (pt-PT)"));
  const draft = JSON.parse(await capture((f) => generateDraft(DI, CFG, f))).instructions as string;
  assertEquals(draft, INSTRUCTIONS);
  for (const tone of TONE_OPTIONS) {
    for (const revise of [undefined, { instruction: "mais curto", draft: "Olá." }]) {
      const get = async (language: "match" | "pt" | "en") => JSON.parse(await capture((f) => generateReply(DI, { language, tone }, CFG, f, revise))).instructions as string;
      for (const [language, name] of [["en", "English"], ["pt", "European Portuguese (pt-PT)"]] as const) {
        const ins = await get(language);
        assertFalse(ins.includes(LANGUAGE_PHRASE), `${language}/${tone}`);
        assertFalse(ins.includes("whatever the language"));
        assert(ins.startsWith(INSTRUCTIONS.split(LANGUAGE_PHRASE)[0] + `in ${name}`));
        assert(ins.endsWith(` The whole reply must be in ${name}.`));
      }
      const m = await get("match");
      assert(m.includes(LANGUAGE_PHRASE) && !m.includes("The whole reply must be"));
    }
  }
});
