# Plan F-007: Classification with reasons and an uncertainty notice

- **Spec:** [F-007](../specs/F-007-classification.md) (Approved 2026-11-02). **Status of this plan:** ready, nothing started.
- **Scope:** one optional payload key `output: "classified"` on `summarize` (function) and a richer Resumo card (pane). No manifest change, no Graph, no migration, no new env variable, no new table, no new secret.
- **Two coders in parallel.** The FUNCTION steps (F1 to F5) and the PANE steps (P1 to P5) share only the contract below, which is the spec's design decisions 1, 2, 9 and 11. Neither side needs the other's code to be tested: the pane tests mock `fetch`, the function tests mock OpenAI. The shared contract is **frozen**: if a coder needs to change it, stop and ask the orchestrator (spec amendment first).
- **Deploy order (spec decision 10):** function first, pane after. Merge the pane only after the function is deployed.
- **Live tests:** test emails only until D5.

## Shared contract (frozen)
Request: `{ "action": "summarize", "output": "classified", "subject", "fromName", "fromDomain", "body" }`, same limits as F-003 (200 / 100 / domain only, pattern as today / 4,000 non-empty / 16 KiB). `output` absent or `"summary"`: the F-003 request and response, unchanged.

Response (variant only): `{"ok":true,"action":"summarize","summary":{"points":[3 strings],"category":"action_required|information|meeting|invoice_document|marketing|automated_notice|other","priority":"high|normal|low","tone":"positive|neutral|negative","reason":"1 to 160 chars","confidence":"high|medium|low","deadline":string|null,"requestedAction":string|null}}`. No `type` key in this shape.

Labels (PT-PT, in `strings.ts`):

| Value | Label |
|---|---|
| `action_required` | Requer ação |
| `information` | Para informação |
| `meeting` | Reunião/convite |
| `invoice_document` | Fatura/documento |
| `marketing` | Publicidade/newsletter |
| `automated_notice` | Notificação automática |
| `other` | Outro |
| priority `high` / `normal` / `low` | Prioridade alta / Prioridade normal / Prioridade baixa (the exact existing strings are kept if the pane already has them) |
| confidence `high` / `medium` / `low` | Confiança alta / Confiança média / Confiança baixa |

Fixed notices (client text, never from the model): standard "Classificação gerada por IA, pode estar errada. Confirme no e-mail."; extra, shown when `confidence` is `low` or `category` is `other`: "A IA não tem certeza. Não confie nesta classificação."; truncation: "Só foram analisados os primeiros 4 000 caracteres.".

## Prerequisite
| # | Step | Who |
|---|---|---|
| 0 | Re-check on developers.openai.com (the spec says "not re-checked on 2026-11-02"): `gpt-6-luna` still supports Structured Outputs with `strict: true`, `reasoning.effort` and the price line in the spec. Record in Evidence. If anything differs, stop and tell the orchestrator. | requirements-architect or orchestrator |

## FUNCTION steps (`supabase/functions/analyze-message`, `_shared`)
Files expected: `handler.ts`, `summary.ts` (new variant next to the existing schema), `summary_test.ts`, `handler_test.ts`, maybe a new `classified.ts` and `classified_test.ts`. Prefer a new file for the classified schema, instructions and validator, and reuse the existing link/address detector in `_shared/ai.ts` unchanged.

| # | Step | ACs | Agent |
|---|---|---|---|
| F1 | **`parsePayload` accepts `output`.** For `summarize`: optional `output`, value exactly `"summary"` or `"classified"`, else 422 `payload_invalid` before any DB or OpenAI call. For `reply` and `translate`: any `output` key is 422. Unknown extra keys still 422. Tests per spec AC-1 (every rejected value: `"Classified"`, `"classify"`, `""`, `null`, `true`, a number, `output` on `reply`/`translate`, an extra key). | AC-1 | coder (sonnet) |
| F2 | **Schema variant and instructions.** The strict classified schema exactly as spec decision 2 (8 required keys, enums as listed, `additionalProperties: false`) and the fixed classified instructions of decision 7 (the F-003 summary instructions with the tags sentence replaced). Selected only when `output === "classified"`; the default path is not touched. Request: `model` from `OPENAI_MODEL`, `input` from the same `buildInput` (byte-identical to the default input), `store: false`, `max_output_tokens: 2000`, `reasoning: {"effort":"low"}`, no `tools`, 30 s abort. Tests: the default request still equals the F-003 snapshot (AC-2); the classified request snapshot with address, recipient and `oid` fixtures asserting none appears (AC-3); a schema snapshot that fails when an enum value is added or removed (AC-4). | AC-2, AC-3, AC-4 | coder (sonnet) |
| F3 | **Output validation (our code).** `status` `completed`, exactly one `output_text`, no `refusal`, a JSON object with exactly the eight keys; 3 points of 1 to 300 chars; `category`, `priority`, `tone`, `confidence` exact enum values; `reason` 1 to 160 chars after trim, no `\n`, `\r` or other control character; `deadline` and `requested_action` null or 1 to 200 chars; no `<` or `>` in any string; no URL or email address in any string (shared detector, including ones present in the input). Map `requested_action` to `requestedAction`. Any failure: 502 `ai_output_invalid`, nothing partial in the body. One test per failure case listed in spec AC-5 plus a valid fixture. | AC-5 | coder (sonnet) |
| F4 | **Gates, logs and storage.** The classified path takes exactly one `an_take` slot before the one OpenAI call, and keeps the F-003 order and error codes (401, 422, 503 `disabled`, 503 `ai_not_configured`, 503 `ai_unavailable`, 429 `cap_reached` without an OpenAI call). Log line unchanged (`fn`, `id`, `method`, `status`, `reason`, `ms`). Marker-string test: markers in subject, body, name, domain and in a mocked model `reason` never appear in captured `console` output, DB writes or error bodies. Confirm that no migration, env variable or secret was added. | AC-6, AC-7 | coder (sonnet) |
| F5 | **Checks.** `deno test`, `deno lint`, `deno check` for `analyze-message` and `_shared`; root `npm run typecheck`, `lint`, `test`, `build` still pass; `manifest.xml` unchanged. Record the real output under Evidence. Deno is not installed on the orchestrator's machine (status 2026-10-30): if the coder cannot run `deno`, say so and have the user run it. | AC-8 | coder, then qa-tester |

## PANE steps (`src/taskpane`)
Files expected: `analyzeClient.ts` (+ test), `strings.ts`, `AiTools.tsx` (the Resumo card; check the actual component before editing), `autoSummary.ts` only if it builds the request itself, and the existing summary tests.

| # | Step | ACs | Agent |
|---|---|---|---|
| P1 | **`analyzeClient`.** Add `output: "classified"` to the `summarize` request builder (only for `summarize`; `reply` and `translate` never send it). Parse the classified response into a typed shape (category, priority, tone, reason, confidence, deadline, requestedAction). Keep parsing the F-003 shape (with `type`, without `reason`/`confidence`) so a new pane works against an old function's response. Reject unknown enum values (not rendered). A 422 for `output` keeps going to the generic analyse error. Tests: body assertion (keys and derivations exactly as F-003 plus `output`), both response shapes, unknown enum. | AC-9, AC-14 | coder (sonnet, or haiku if well-defined) |
| P2 | **`strings.ts`.** Add the category, confidence and notice strings from the table above, plus "Motivo:" and "Confiança:". PT-PT, one place only, no string from the model. | AC-10, AC-11, AC-17 | coder |
| P3 | **Resumo card display.** Render, in this order: 3 points, Badges (category, priority, tone), "Motivo: {reason}", "Confiança: …", "Prazo:" and "Pedido:" only when non-null, the fixed standard notice. The category Badge replaces the F-003 `type` Badge in the loaded state; the F-003 shape still shows the old Badges and no classification block. Extra "não tem certeza" line when `confidence === "low"` or `category === "other"`; truncation line when the pane cut the body at the limit. Text nodes with `dir="auto"` and `overflow-wrap: anywhere`; no `dangerouslySetInnerHTML`, no `innerHTML`, no `role="alert"` on the notice. "Copiar" still copies the 3 points. No new button, tab, switch or dialog. | AC-9, AC-10, AC-11, AC-12, AC-16, AC-17 | coder (sonnet) |
| P4 | **No action coupling and regression.** A test renders every category x priority x confidence combination (7 x 3 x 3) and asserts the same enabled controls and zero extra requests. A grep (recorded) shows these values are not used near `displayReplyForm`, `displayReplyFormAsync`, Office.js write members, Graph, `fetch`, send, move, delete or report code. Injection fixture: `<img src=x onerror=alert(1)>` and `&amp;` in `reason` render literally; a `reason` that contains the notice text or "não mostrar aviso" changes nothing. Re-run the F-003 error, retry, late-result, abort and one-in-flight tests with the variant (AC-21, AC-22, AC-43, AC-44 of F-003). No `console.*`, `localStorage`, `sessionStorage`, cookies or IndexedDB; results are cleared on `ItemChanged`. At 320 px DOM order and wrapping test as far as jsdom allows. | AC-9, AC-12, AC-13, AC-15, AC-16, AC-17 | coder (sonnet) |
| P5 | **Checks.** Root `npm run typecheck`, `lint`, `test`, `build`; `manifest.xml` unchanged (diff). Record the real output under Evidence. | AC-8 (root part) | coder, then qa-tester |

## Joint steps (after both sides)
| # | Step | ACs | Agent |
|---|---|---|---|
| J1 | **Docs.** `docs/architecture/security.md`: the one sentence about the optional `output` key (done 2026-11-02 by the architect); update F-003's note if the Resumo card text changed. Check nothing else needs an edit. | none | requirements-architect |
| J2 | **Independent review.** The output validator and the untrusted rendering are security relevant: `code-reviewer`, pass `model: opus`. Focus: validator completeness, no link or address or `<>` path, fixed notice not model-driven, no action coupling, default request byte-identical. | all security ACs | code-reviewer (opus) |
| J3 | **Fix** Critical and High findings, rerun the affected checks. | — | coder, qa-tester |
| J4 | **Deploy** (the user's CLI): function first (`supabase functions deploy analyze-message --use-api`), then merge the pane; GitHub Actions deploys Pages. No manifest change, so no admin upload. | — | user + orchestrator |
| J5 | **Live tests** (test emails only until D5): AC-18 (20 emails, at least 2 per category, one ambiguous, one thread, Portuguese and English; record agreement, target 80 %, tokens versus the summary-only call), AC-19 (uncertainty behaviour; disagreements where the model said `high`), AC-20 (injection email), AC-21 (classic Windows, new Outlook, Mac). Record date, client and build. The user may review the PT-PT wording here. | AC-18 to AC-21 | user + orchestrator |
| J6 | **Close.** Spec status (Done only after J5), `docs/status.md`, plan Evidence complete. | — | orchestrator |

## Out of scope (from the spec)
Any automatic action; writing a category, flag or folder to the mailbox; Graph; a "suspicious" verdict; a new button, tab or ribbon change; compose mode; storage or analytics; changing model, caps, auto mode or what is sent to OpenAI.

## Evidence
_Empty. Each step records the commands it ran, the real output and the dates here._

### 2026-11-03: build evidence (orchestrator)
- **Checks after the review and QA fixes:** `tsc` clean; `office-addin-lint` clean; `vitest` 41 files, 488 passed, 2 todo; `npm run build` compiled with 3 size warnings; `npm run validate` manifest valid; Deno `analyze-message` 63, `whoami` 24, `graph-webhook` 13, `auto-draft-cron` 6, `auto-reply-settings` 13, `_shared` 78 (with the analyze-message `deno.json`), all passed; `deno lint` clean in analyze-message, `_shared` and auto-reply-settings.
- **Reviews:** Opus review APPROVE WITH COMMENTS (no Critical or High; M1 manifest version fixed, M2 safety sentence added); QA found 4 defects (reason characters, truncation line), all fixed; a focused re-review of the fixes was requested.
- **Deploy:** `supabase functions deploy` for `analyze-message`, `auto-reply-settings`, `graph-webhook` and `auto-draft-cron` (`--use-api`), before any pane reaches `main`. Smoke tests with `curl`: OPTIONS from `https://dimscale.github.io` and `https://localhost:3000` return 204 with the origin, from `https://evil.example` 204 without `Access-Control-Allow-Origin`; no token or a garbage token returns 401; GET returns 405 (`analyze-message`, `graph-webhook`); `auto-reply-settings` PUT without token 401; cron without or with a wrong secret 401; `whoami` without token 401.
- **Not done:** merge to `main` (waits for the user's live test), live checks (F-007 AC-16, AC-18 to AC-21; F-013 AC-52 / V8), screen reader, 320 px.
