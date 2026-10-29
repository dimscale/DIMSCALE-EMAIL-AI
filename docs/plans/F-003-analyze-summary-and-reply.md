# Plan F-003: AI in the Analisar tab (summary and suggested reply)

- **Status:** In progress (2026-10-29). Spec approved; backend and pane built and reviewed; deploy and live tests pending.
- **Status:** Draft (2026-10-29). Nothing past step 0 starts until the user sets the spec to `Approved`.
- **Blocked by:** the user's approval of the spec; the user's answers to Q-U1 (model) and Q-U4 (real mail before D5). D5 (DPO) blocks use on real mail, not the build or the test-email live checks. No tenant-admin dependency.
- **Amendment 2026-10-29 (Kutools-style pane):** steps 19 to 36 below implement spec AC-35 to AC-51 together with the F-012 plan steps 17 to 24. Nothing in steps 20 to 36 starts until the user confirms the amended specs (the architect proposals Q-U10 to Q-U15 and F-012 Q-K1 to Q-K6); the orchestrator records the answers here.
- **Concurrency note:** other agents edit `src/` and tests at the same time. Each coder step re-reads the files it touches right before editing and keeps its diff to the files listed.

## Configuration values
| Value | Default | Secret? | Where |
|---|---|---|---|
| `OPENAI_API_KEY` | — | **Yes** | Supabase secret (already set 2026-10-23) |
| `OPENAI_MODEL` | none (fail closed) | No | Supabase secret, **set at deploy** (Q-U1: `gpt-6.1-terra` or `gpt-6-luna`). Shared with F-013 |
| `ANALYZE_DISABLED` | unset = enabled | No | Kill switch: `true` → 503 `disabled` |
| `ANALYZE_DAILY_CAP_PER_USER` | 50 | No | Optional |
| `ANALYZE_DAILY_CAP_GLOBAL` | 1,500 | No | Optional. Both caps: a value above 1,000,000 or invalid → default |
| `ENTRA_TENANT_ID`, `ENTRA_CLIENT_ID`, `ALLOWED_ORIGINS` | existing | No | Already set (F-002/F-013) |
| `SUPABASE_URL`, secret keys | platform | Yes (platform) | Injected by Supabase |
| `reasoning.effort` `low`, `max_output_tokens` 2,000 | constants in the function | No | Code (`shortcut:` comment; env only if the user asks) |

## Steps
| # | Step | Files | ACs | Agent |
|---|---|---|---|---|
| 0 | Decisions: the user approves the spec and answers Q-U1 to Q-U9 (at least Q-U1, Q-U4); the orchestrator records them here and in `docs/status.md`. DPO questions D5 1 to 4 sent. | — | prerequisite | user, orchestrator |
| 1 | Docs before code (SDD step 3): `security.md` (new user-initiated flow pane → `analyze-message` → OpenAI with the exact fields, quoted history, `store: false`, the pseudonymised-`oid` counter and its purge inside `an_take`, reply-form escaping), `overview.md` (diagram line for `analyze-message`, capability rows unchanged), `status.md` (F-003/F-004 merged row, F-005 partly covered, D5 now also blocks F-003 real use), F-012 spec note that the two "Em breve" AI cards are superseded by F-003. | docs only | — | requirements-architect |
| 2 | Migration `supabase/migrations/<ts>_f003_analyze_usage.sql`: table `analyze_usage(user_key text, day date, count int, pk(user_key, day))`, RLS on, no policies, revoke from `anon`/`authenticated`; `an_take(p_user text, p_cap int, p_global int) returns boolean` (`set search_path = ''`): delete rows with `day < current_date - 2`, then atomically increment today's row only if user count < cap and today's sum < global (lock the day, e.g. `pg_advisory_xact_lock` on the date, so the global check is not racy), return whether a slot was taken; `execute` only to `service_role`. Do not touch the F-013 migration. | migration | AC-5 (atomicity proven only by the local two-parallel-calls check), AC-6 | coder (sonnet) |
| 3 | `_shared/ai.ts` refactor, **no F-013 behaviour change**: extract (a) `requestResponses(text, cfg, instructions, name, schema, fetchFn)` returning the parsed JSON or the existing `ai_unavailable`/`ai_output_invalid` failures, (b) `singleOutputText(res)` (the `completed`/refusal/exactly-one-`output_text` part of `parseOutput`). `generateDraft` and `parseOutput` keep their signatures and results. Add optional `reasoningEffort?: string` to `AiConfig`: the request gets `reasoning: {effort}` only when set. Add a snapshot test of the F-013 request body without `reasoningEffort`. Run every existing `_shared` test unchanged. | `_shared/ai.ts`, new `_shared/ai_test.ts` (snapshot only) | AC-14 | coder (sonnet) |
| 4 | `supabase/functions/analyze-message/summary.ts`: `SUMMARY_INSTRUCTIONS`, `SUMMARY_SCHEMA` (exactly the spec's), `parseSummary(res, input)` (spec validation list; reuses `singleOutputText`; rejects any URL or email address, even one in the input), `generateSummary(input, cfg, fetchFn)`. Tests: one fixture per AC-9 failure, the valid case, a URL present in the email rejected, the exact request body. | `analyze-message/summary.ts`, `summary_test.ts` | AC-7, AC-9 | coder (sonnet) |
| 5 | Function `analyze-message`: `handler.ts` (pattern of `auto-reply-settings/handler.ts`: CORS, `validateStateful`, strict payload, kill switch, OpenAI config, `an_take` via `sha256Hex(oid)`, dispatch to `generateSummary` / `generateDraft` with `{model, apiKey, maxOutputTokens: 2000, reasoningEffort: "low"}`, status mapping, one `log()` line), `index.ts` (as `auto-reply-settings/index.ts`), `deno.json` (same pinned `jose`). `Db` gets one method `analyzeTake(userKey, cap, global)` → `an_take`; update the in-memory fakes in `_shared/testing.ts`. `supabase/config.toml`: `[functions.analyze-message] verify_jwt = false` with a one-line comment. Tests: AC-1 to AC-5, AC-8, AC-10 to AC-12 (marker-string log test). | `analyze-message/{handler,index}.ts`, `deno.json`, `handler_test.ts`, `_shared/db.ts`, `_shared/testing.ts`, `supabase/config.toml` | AC-1 to AC-5, AC-8, AC-10 to AC-13 | coder (sonnet) |
| 6 | Backend checks: `deno test`, `deno lint`, `deno check` on `supabase/functions` (all functions, so F-013 regressions show). Grep the new function for `console.` (only via `log`), `graph.microsoft.com`, `Mail.Send`. Record output below. | — | AC-14, AC-15 | qa-tester (haiku) |
| 7 | **Independent review of steps 2 to 5** (auth, cap race, prompt, output validation, logging, F-013 non-regression). Report only. | — | all function ACs | code-reviewer (**opus**) |
| 8 | Fix Critical/High from step 7 (or escalate), rerun step 6. | as step 2–5 | — | coder (sonnet), qa-tester |
| 9 | Pane client: `config.ts` gets `ANALYZE_URL = "https://qepzpkseyzclxnqguqta.supabase.co/functions/v1/analyze-message"`; new `src/taskpane/analyzeClient.ts` with `analyze(token, payload, fetchFn)` (POST, 40 s abort, maps the responses to a typed result per the spec's error table). Unit tests for each status. | `config.ts`, `analyzeClient.ts`, `analyzeClient.test.ts` | AC-18 (headers), AC-21 (mapping) | coder (**haiku**) |
| 10 | Pane payload and reply-form helpers (pure): in `messageView.ts` add `aiText` (trimmed body cut to 4,000, no lone surrogate) and `fromDomain` (after the last `@`, lower-cased, `""` without `@`) without changing existing fields; new `replyForm.ts` with `toReplyHtml(text)` (escape `& < > " '`, `\r?\n` → `<br>`) and `openReplyForm(item, html)` (1.9 → `displayReplyFormAsync`, else `displayReplyForm` in `try/catch`; returns a promise of ok/failed). Tests for every escape, the CRLF case, both branches and both failure paths. | `messageView.ts`, `replyForm.ts`, tests | AC-18, AC-24 | coder (**haiku**) |
| 11 | Pane UI: new `AiTools.tsx` rendered by `MessagePanel` in the `ok` state below the accordion (gets the view and an item key); buttons, privacy notice, pending state, summary and reply cards, Badges via the enum label map, "Copiar" with fallback, "Abrir resposta para revisão", error MessageBars with retry per the table, discard of late results on item change, empty-body disabled state, `getTokenSilent` at click. `App.tsx`: remove the "Resumo do email" and "Resposta rápida" `AiToolCard`s, keep "Melhorar texto e tom". `strings.ts`: the strings listed below. Tests for AC-16, AC-17, AC-19 to AC-23, AC-25; update existing App/F-012 tests that asserted the removed cards (assertion replaced, not weakened). | `AiTools.tsx`, `MessagePanel.tsx`, `App.tsx`, `strings.ts`, tests | AC-16, AC-17, AC-19 to AC-23, AC-25 | coder (sonnet) |
| 12 | Pane checks: `npm run typecheck`, `lint`, `test`, `build`, `build:dev`, `validate`; vitest count not lower than before step 9; grep `src/` for `dangerouslySetInnerHTML`, `innerHTML`, `console.`, `localStorage`, `sessionStorage`, `displayReplyAllForm`, `Mail.Send`; `git diff --stat manifest.xml` empty (run by the orchestrator, not by subagents told not to use git). Record output below. | — | AC-25, AC-26 | qa-tester (haiku) |
| 13 | Review of steps 9 to 11 (rendering, escaping, no auto-run, data sent). | — | pane ACs | code-reviewer (sonnet) |
| 14 | Fix Critical/High, rerun step 12. | — | — | coder, qa-tester |
| 15 | **Deploy backend** (the user's CLI, project `qepzpkseyzclxnqguqta`): `supabase db push` (adds `analyze_usage`/`an_take`); `supabase secrets set OPENAI_MODEL=<Q-U1 choice>` (the only new secret; `OPENAI_API_KEY` already set); `supabase functions deploy analyze-message --use-api`; `supabase secrets list` shows `OPENAI_API_KEY`, `OPENAI_MODEL`, `ENTRA_TENANT_ID`, `ENTRA_CLIENT_ID`, `ALLOWED_ORIGINS` (names only, never values in reports). Smoke: `OPTIONS` from `https://dimscale.github.io` → 204 with the origin; `POST` without a token → 401. Confirm `AUTO_DRAFTS_ENABLED` is still not `true` (setting `OPENAI_MODEL` must not start F-013). | — | AC-1, AC-2 (live smoke) | user + orchestrator |
| 16 | **Deploy pane:** merge to `main`; the F-011 workflow publishes to Pages. No manifest change, so no admin re-upload. | — | AC-26 | orchestrator |
| 17 | **Live tests, test emails only until D5:** AC-27 (thread summary, time to result), AC-28 (reply form, Sent Items empty, original unchanged, signature/quote noted), AC-29 (new Outlook, Mac and its UI), AC-30 (clipboard), AC-31 (logs, cap 1, kill switch, restore), AC-32 (injection email). Record real `usage` tokens from the OpenAI dashboard for the cost table. | — | AC-27 to AC-32 | user + orchestrator |
| 18 | Close: spec status, `status.md` (F-003 Done; F-004 merged; F-005 partly delivered), evidence complete, cost table updated with measured usage. | docs | — | orchestrator |

### Amendment 2026-10-29: Kutools-style pane (auto mode, translate, reply options)
Lanes that can run in parallel because their files don't overlap: **A** function (`supabase/functions/analyze-message/`), **B** pane client (`analyzeClient.ts`), **C** manifest, icons and build (F-012 plan step 18). Lane **D** (pane UI) is sequential because every step touches `AiTools.tsx`/`App.tsx`/`strings.ts`: F-012 step 17 → step 27 → step 28 → F-012 step 19.

| # | Step | Files | ACs | Agent | Lane / after |
|---|---|---|---|---|---|
| 19 | Docs: amend F-003 and F-012 specs and plans, `security.md`, `status.md`. | docs only | — | requirements-architect | done 2026-10-29 |
| 20 | Function, payload and reply options: `parsePayload` per action (spec table: base keys, optional `language`/`tone` for `reply`, required `targetLanguage` and empty `fromName`/`fromDomain` for `translate`); new `analyze-message/reply.ts` with the option-sentence constants and `generateReply(input, opts, cfg, fetchFn)`: `match`+`none` → `generateDraft` unchanged; otherwise `requestResponses` with `INSTRUCTIONS + " " + sentences`, F-013 `SCHEMA`, `parseOutput`; then the existing no-links check. Tests: AC-35 (every 422 case, no DB/OpenAI calls), AC-36 (snapshot equality for defaults, exact instructions for every option, `input` unchanged, injection in `tone` → 422). | `analyze-message/handler.ts`, `reply.ts`, `reply_test.ts`, `handler_test.ts` | AC-35, AC-36 | coder (sonnet) | A |
| 21 | Translate module: `analyze-message/translate.ts` with `TRANSLATE_INSTRUCTIONS(target)`, `TRANSLATE_SCHEMA`, `parseTranslation(res, input)` (spec validation; reuses `singleOutputText` and the no-URL/no-address check used by `summary.ts`), `generateTranslation(input, target, cfg, fetchFn)` with `maxOutputTokens: 4000`. Tests: exact request body (AC-37), one fixture per AC-38 failure, valid case. New files only, so it runs in parallel with step 20. | `translate.ts`, `translate_test.ts` | AC-37, AC-38 | coder (sonnet) | A, parallel with 20 |
| 22 | Wire `translate` into `handler.ts` (dispatch, response shape `translation`), cap slot taken before every action, marker-string log test extended to `translate` and reply options. | `handler.ts`, `handler_test.ts` | AC-38 (response), AC-39 | coder (haiku) | A, after 20 and 21 |
| 23 | Backend checks: `deno test`, `deno lint`, `deno check` on all of `supabase/functions` (F-013 regressions show); grep `analyze-message` for `console.` outside `log`, `graph.microsoft.com`, `Mail.Send`. Record output. | — | AC-14, AC-15, AC-39 | qa-tester (haiku) | A |
| 24 | **Independent review of steps 20 to 22** (strict payload, prompt construction from enums only, no free text in instructions, translate minimisation, output validation, F-013 non-regression, logging). Report only. | — | AC-35 to AC-39 | code-reviewer (**opus**) | A |
| 25 | Fix Critical/High from step 24 (or escalate), rerun step 23. | as 20–22 | — | coder (sonnet), qa-tester | A |
| 26 | Pane client: `analyzeClient.ts` types for `language`/`tone`/`targetLanguage`, translate result mapping, optional `AbortSignal` with an "aborted" result; `messageView.ts` gets a `translatePayload` helper (empty `fromName`/`fromDomain`). Unit tests. | `analyzeClient.ts`, `analyzeClient.test.ts`, `messageView.ts`, `messageView.ai.test.ts` | AC-41 (payload), AC-46 | coder (haiku) | B, parallel with A and C |
| 27 | Pane cards: inside the F-012 `ToolCard` shell (F-012 step 17 first): Resposta `Select`s "Idioma"/"Tom" sent with every reply; Traduzir card (`Select` "Traduzir para", "Traduzir", result with "Assunto:" line and body, Copiar = subject + blank line + body); strings below; remove the "DIMSCALE AI Assistant" heading and the "Melhorar texto e tom" card (if Q-K3 confirmed). Update the existing AiTools/App tests (assertions replaced, not weakened). | `AiTools.tsx` (or split per card), `strings.ts`, tests | AC-40, AC-41, AC-17 (amended) | coder (sonnet) | D, after F-012 step 17 and step 26 |
| 28 | Auto mode: `src/taskpane/prefs.ts` (`readAutoSummary()` true only for `=== true`, `saveAutoSummary(v)` wrapping `roamingSettings.set` + `saveAsync` in a promise, never throws); `AUTO_SUMMARY_AVAILABLE` in `config.ts` (value per Q-U13); in the Resumo card the `Switch`, the confirmation `Dialog`, the "ligado" line, the save-failed notice; the scheduler (800 ms debounce, one request per item per session, `AbortController` for the in-flight auto request, pause on 429/401/503 disabled/not configured, no retry). Keep the scheduler in a small hook or pure module so it is tested with fake timers. Tests: AC-42, AC-43 (a) to (i), AC-44, AC-45. | `prefs.ts`, `config.ts`, `AiTools.tsx`, `strings.ts`, new `autoSummary.ts`, tests | AC-42 to AC-45 | coder (sonnet) | D, after 27 |
| 29 | Pane checks: `npm run typecheck`, `lint`, `test`, `build`, `build:dev`, `validate`; vitest count not lower than before step 26; grep `src/` for `localStorage`, `sessionStorage`, `indexedDB`, `document.cookie`, `location.search`, `location.hash`, `dangerouslySetInnerHTML`, `innerHTML`, `console.`, `displayReplyAllForm`, `Mail.Send`. Record output. | — | AC-25 (amended), AC-26 (amended), AC-42 | qa-tester (haiku) | D, after 28 and F-012 step 19 |
| 30 | **Independent review of steps 26 to 28 and F-012 step 19** (auto-mode consent path: no request without "Ligar"; abort and late-result discard; pause logic; entry-page start; data sent per action; rendering). Report only. | — | AC-40 to AC-46 | code-reviewer (**opus**) | D |
| 31 | Fix Critical/High from step 30, rerun step 29. | as 26–28 | — | coder, qa-tester | D |
| 32 | **Deploy the function first** (user's CLI): `supabase functions deploy analyze-message --use-api`. Smoke: the old pane's `reply` without options still works (test email); `OPTIONS`/401 as step 15. No new secret, no migration. | — | AC-35 (live smoke) | user + orchestrator | after 25 |
| 33 | **Deploy the pane:** merge to `main`; F-011 publishes `taskpane.html` and `resumo.html`. The installed manifest (1.0.1.0) keeps working (F-012 edge case). | — | — | orchestrator | after 31 and 32 |
| 34 | **Live, test emails only (D5):** AC-47 (options, translate), AC-48 (auto mode on a test mailbox, with the flag switched on in a dev build if Q-U13 keeps production off), AC-49 (roaming), AC-50 (reload, with F-012 AC-52), AC-51 (translate injection). | — | AC-47 to AC-51 | user + orchestrator | after 33, with F-012 step 23 |
| 35 | If AC-50 shows the stale "on" after a reload: spec amendment for the `sessionStorage` "off" guard, user approval, coder, review, then recheck. Otherwise skip. | — | AC-50 | requirements-architect, coder | conditional |
| 36 | Close the amendment: spec status lines, `status.md`, evidence complete; the admin upload itself is F-012 step 24. | docs | — | orchestrator | last |

### Strings to add for the 2026-10-29 amendment (PT-PT, for the user's review)
| Key (suggested) | Text |
|---|---|
| `replyLanguageLabel` | Idioma |
| `replyLanguageMatch` / `replyLanguagePt` / `replyLanguageEn` | Idioma do e-mail / Português / Inglês |
| `replyToneLabel` | Tom |
| `replyToneNone` / `Concise` / `Detailed` / `Formal` / `Friendly` | Normal / Conciso / Detalhado / Formal / Amigável |
| `translateTargetLabel` | Traduzir para |
| `translatePt` / `translateEn` | Português / Inglês |
| `aiTranslate` / `aiTranslating` | Traduzir / A traduzir… |
| `aiTranslationTitle` | Tradução |
| `aiTranslatedSubject` | Assunto: {x} |
| `autoSummaryLabel` | Resumo automático |
| `autoSummaryDialogTitle` | Ligar o resumo automático? |
| `autoSummaryDialogBody` | Com o resumo automático, sempre que abrir um e-mail com este painel no separador Resumo, o assunto, o nome e domínio do remetente e o texto do e-mail são enviados à OpenAI para gerar o resumo, sem clicar em Resumir. Cada resumo conta para o seu limite diário de 50 pedidos de IA. Pode desligar a qualquer momento. |
| `autoSummaryConfirm` / `autoSummaryCancel` | Ligar / Cancelar |
| `autoSummaryOnNote` | Resumo automático ligado: cada e-mail que abrir com o painel neste separador é enviado à OpenAI. |
| `autoSummarySaveFailed` | Não foi possível guardar a preferência. Aplica-se apenas até fechar o painel. |
| `autoPausedCap` | Atingiu o limite diário de pedidos de IA. O resumo automático fica em pausa até amanhã. |
| `autoPausedSuffix` | O resumo automático fica em pausa até reabrir o painel. |

### Strings to add to `strings.ts` (PT-PT, for the user's review, Q-U9)
| Key (suggested) | Text |
|---|---|
| `aiSummarize` | Resumir |
| `aiSuggestReply` | Sugerir resposta |
| `aiSummarizing` | A resumir… |
| `aiReplying` | A preparar resposta… |
| `aiPrivacyNotice` | Ao usar a IA, o assunto, o nome e domínio do remetente e o texto desta mensagem são enviados à OpenAI. |
| `aiNoText` | Esta mensagem não tem texto para analisar. |
| `aiNeedSignIn` | Inicie sessão (em baixo) para usar a IA. |
| `aiSummaryTitle` (reuse) | Resumo do email |
| `aiReplyTitle` (reuse the key, text changes from "Resposta rápida") | Resposta sugerida |
| `aiAdvisory` | Gerado por IA, pode estar errado. |
| `aiDeadline` | Prazo: {x} |
| `aiRequestedAction` | Pedido: {x} |
| `aiCopy` / `aiCopied` | Copiar / Copiado |
| `aiCopyFailed` | Não foi possível copiar. Selecione o texto e use Ctrl+C. |
| `aiOpenReply` | Abrir resposta para revisão |
| `aiOpenReplyFailed` | Não foi possível abrir a resposta. Use Copiar. |
| errors | the seven texts of the spec's error mapping table |
| type labels | request "Pedido", information "Informação", meeting "Reunião", invoice_payment "Fatura/pagamento", marketing "Publicidade", notification "Notificação automática", other "Outro" |
| priority labels | high "Prioridade alta", normal "Prioridade normal", low "Prioridade baixa" |
| tone labels | positive "Tom positivo", neutral "Tom neutro", negative "Tom negativo" |

## Stop and rollback
- Immediate stop: `supabase secrets set ANALYZE_DISABLED=true` (no redeploy; the pane shows "A IA está temporariamente desligada.").
- Cheaper model: `supabase secrets set OPENAI_MODEL=gpt-6-luna`.
- Stop auto mode only (2026-10-29): set `AUTO_SUMMARY_AVAILABLE = false` in `config.ts` and merge (Pages deploy, no admin); users' saved preference is then ignored. The kill switch above stops everything at once.
- Full removal: `supabase functions delete analyze-message`; revert the pane commit (Pages redeploys). The `analyze_usage` rows are purged by `an_take` itself, so they stay after the last call; drop the table in a new migration (tiny, pseudonymised).
- Nothing in any mailbox needs cleaning: the feature never writes mail.

## Evidence
_Empty. Each step records the commands it ran, the real output, dates and live observations here._

### 2026-10-29: backend deploy (step 15) and smoke tests
- **Checks before deploy (orchestrator, real runs):** Deno suites `analyze-message` 72, `whoami` 24, `graph-webhook` 68, `auto-draft-cron` 61, `auto-reply-settings` 64, all passed; `deno lint` clean on `analyze-message`; root typecheck and lint clean, vitest 251 passed (maxWorkers 4), build passes.
- **Review:** the backend security review (Opus) found no Critical or High. The two Medium items were fixed: `security.md` updated, and AC-32 resolved by the decision that outputs contain no link or address at all.
- **Deploy (user's CLI):**
  - `supabase db push`: only `20261029120000_f003_analyze_usage.sql` was pending (dry run first) and was applied;
  - `supabase secrets set OPENAI_MODEL=gpt-6-luna`;
  - `supabase functions deploy analyze-message --use-api`.
- **Smoke tests with `curl`:**

| Request | Result |
|---|---|
| OPTIONS from `https://dimscale.github.io` | 204, `Access-Control-Allow-Origin: https://dimscale.github.io`, methods `POST, OPTIONS` |
| OPTIONS from `https://evil.example` | 204 with no `Access-Control-Allow-Origin` |
| POST with no token | 401 `unauthorized` |
| POST with a garbage token | 401 `unauthorized` |
| GET | 405 `method_not_allowed` |

- **Not done yet:** local Postgres checks (AC-5/AC-6 two parallel `an_take` calls), the live tests in Outlook (AC-27 to AC-32), merge to `main` so the pane reaches the published site.

