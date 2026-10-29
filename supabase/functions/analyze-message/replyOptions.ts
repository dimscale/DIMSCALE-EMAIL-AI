// F-003 reply options (decision 13): enums map to fixed English sentences; nothing from the client is free text.
import { type AiConfig, type AiResult, buildInput, type DraftInput, generateDraft, INSTRUCTIONS, parseOutput, requestResponses, SCHEMA } from "../_shared/ai.ts";

export const LANGUAGES = ["match", "pt", "en"] as const;
export const TONE_OPTIONS = ["none", "concise", "detailed", "formal", "friendly"] as const;
export type ReplyOpts = { language: (typeof LANGUAGES)[number]; tone: (typeof TONE_OPTIONS)[number] };

export const LANGUAGE_SENTENCES: Record<Exclude<ReplyOpts["language"], "match">, string> = {
  pt: "Write the reply in European Portuguese (pt-PT), whatever the language of the email. This replaces the language rule above.",
  en: "Write the reply in English, whatever the language of the email. This replaces the language rule above.",
};
export const TONE_SENTENCES: Record<Exclude<ReplyOpts["tone"], "none">, string> = {
  concise: "Make the reply concise: at most three short sentences besides the greeting and the closing.",
  detailed: "Make the reply detailed instead of short: answer each point or question the sender raised.",
  formal: "Use a formal, courteous register.",
  friendly: "Use a warm, friendly register while staying professional.",
};

export async function generateReply(input: DraftInput, opts: ReplyOpts, cfg: AiConfig, fetchFn: typeof fetch = fetch): Promise<AiResult> {
  const extra = [
    opts.language === "match" ? "" : LANGUAGE_SENTENCES[opts.language],
    opts.tone === "none" ? "" : TONE_SENTENCES[opts.tone],
  ].filter(Boolean);
  if (extra.length === 0) return generateDraft(input, cfg, fetchFn); // default options: byte-identical to F-013
  const text = buildInput(input);
  const r = await requestResponses(text, cfg, [INSTRUCTIONS, ...extra].join(" "), "draft_reply", SCHEMA, fetchFn);
  return r.ok ? parseOutput(r.json, text) : r;
}
