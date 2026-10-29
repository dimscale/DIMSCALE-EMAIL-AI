/* global fetch, AbortController, AbortSignal, setTimeout, clearTimeout */
import { ANALYZE_URL } from "./config";

export type AnalyzeAction = "summarize" | "reply" | "translate";
// Options are enums only. The optional instruction/draft refinement fields travel in the input, never in the function's instructions (F-003).
export type ReplyLanguage = "match" | "pt" | "en";
export type ReplyTone = "none" | "concise" | "detailed" | "formal" | "friendly";
// "auto": Portuguese -> English, English -> Portuguese (F-003 amendment 2026-10-29 (b)).
export type TargetLanguage = "auto" | "pt" | "en";

type Base = { subject: string; fromName: string; fromDomain: string; body: string };
export type AnalyzePayload =
  | (Base & { action: "summarize" })
  | (Base & {
      action: "reply";
      language: ReplyLanguage;
      tone: ReplyTone;
      // Refinement (F-003 amendment (d)): both or none; instruction 1..300 trimmed, draft 1..2000.
      instruction?: string;
      draft?: string;
    })
  | (Base & { action: "translate"; targetLanguage: TargetLanguage });

export type Translation = { subject: string; body: string };

export type Summary = {
  points: string[];
  type: string;
  priority: string;
  tone: string;
  deadline: string | null;
  requestedAction: string | null;
};

export type AnalyzeErrorCode =
  | "unauthorized"
  | "payload"
  | "cap"
  | "disabled"
  | "not_configured"
  | "unavailable"
  | "invalid"
  | "other";

export type AnalyzeResult =
  | { kind: "summary"; summary: Summary }
  | { kind: "reply"; reply: string }
  | { kind: "translation"; translation: Translation }
  | { kind: "aborted" } // the caller aborted (item changed, auto mode off): the UI ignores it
  | { kind: "error"; code: AnalyzeErrorCode };

export const RETRYABLE: AnalyzeErrorCode[] = ["unavailable", "invalid", "other"];

const str = (v: unknown) => (typeof v === "string" ? v : null);

export async function analyze(
  token: string,
  payload: AnalyzePayload,
  fetchFn: typeof fetch = fetch,
  signal?: AbortSignal
): Promise<AnalyzeResult> {
  const ctrl = new AbortController();
  let aborted = false;
  const onAbort = () => {
    aborted = true;
    ctrl.abort();
  };
  if (signal?.aborted) return { kind: "aborted" };
  signal?.addEventListener("abort", onAbort);
  const timer = setTimeout(() => ctrl.abort(), 40000);
  try {
    const res = await fetchFn(ANALYZE_URL, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      signal: ctrl.signal,
    });
    const body = await res.json().catch(() => null);
    const error = body?.error;
    if (!res.ok) {
      if (res.status === 401) return { kind: "error", code: "unauthorized" };
      if (res.status === 422) return { kind: "error", code: "payload" };
      if (res.status === 429 && error === "cap_reached") return { kind: "error", code: "cap" };
      if (res.status === 503 && error === "disabled") return { kind: "error", code: "disabled" };
      if (res.status === 503 && error === "ai_not_configured")
        return { kind: "error", code: "not_configured" };
      if (res.status === 503 && error === "ai_unavailable")
        return { kind: "error", code: body?.retryable === true ? "unavailable" : "invalid" };
      if (res.status === 502 && error === "ai_output_invalid")
        return { kind: "error", code: "invalid" };
      return { kind: "error", code: "other" };
    }
    if (body?.ok === true && body.action === "reply" && str(body.reply) !== null)
      return { kind: "reply", reply: body.reply };
    const tr = body?.translation;
    if (
      body?.ok === true &&
      body.action === "translate" &&
      str(tr?.subject) !== null &&
      str(tr?.body) !== null
    )
      return { kind: "translation", translation: { subject: tr.subject, body: tr.body } };
    const s = body?.summary;
    if (
      body?.ok === true &&
      body.action === "summarize" &&
      Array.isArray(s?.points) &&
      s.points.every((p: unknown) => typeof p === "string")
    )
      return {
        kind: "summary",
        summary: {
          points: s.points,
          type: str(s.type) ?? "",
          priority: str(s.priority) ?? "",
          tone: str(s.tone) ?? "",
          deadline: str(s.deadline),
          requestedAction: str(s.requestedAction),
        },
      };
    return { kind: "error", code: "other" };
  } catch {
    if (aborted) return { kind: "aborted" };
    return { kind: "error", code: "unavailable" }; // network error or timeout
  } finally {
    signal?.removeEventListener("abort", onAbort);
    clearTimeout(timer);
  }
}
