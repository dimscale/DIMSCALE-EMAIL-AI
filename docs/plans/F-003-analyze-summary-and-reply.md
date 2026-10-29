# Plan F-003: AI in the Analisar tab (summary and suggested reply)

- **Status:** In progress (2026-10-29). Spec approved; backend and pane built and reviewed; deploy and live tests pending.
- **Status:** Draft (2026-10-29). Nothing past step 0 starts until the user sets the spec to `Approved`.
- **Blocked by:** the user's approval of the spec; the user's answers to Q-U1 (model) and Q-U4 (real mail before D5). D5 (DPO) blocks use on real mail, not the build or the test-email live checks. No tenant-admin dependency.
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

