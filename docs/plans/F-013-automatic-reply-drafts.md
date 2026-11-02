# Plan F-013: Automatic reply drafts (server side)

- **Spec:** [F-013](../specs/F-013-automatic-reply-drafts.md). **ADR:** [ADR-0005](../decisions/0005-server-side-mail-processing.md).
- **Status:** Draft, amended 2026-10-23 for the contact allow-list (v1 trigger). Nothing in here may start until the spec is Approved and ADR-0005 Accepted by the user. Exceptions: step 0 (decisions) and the admin checklist (P1) can be sent in advance.
- **Priority:** before F-015 and F-016 (user direction 2026-10-23). The `ai` module built in step 7 is the shared draft engine that F-015 and F-004 reuse.
- **Amendment 2026-11-02:** Q-C1, Q-C3 and Q-C4 were decided by the user (recommendations of the research note applied). Steps 6a, 6b, 6c and 10d below implement them (spec "Amendment 2026-11-02", AC-43 to AC-52). Steps 6a and 6b change only pure code and tests and can be done in parallel with step 5 (they need no Graph client change); step 6c touches step 10a and 10b code, so it follows them. Q-C1, Q-C3 and Q-C4 no longer block anything.
- **Blocked by:** ADR-0005 acceptance, D5 (the DPO, OpenAI account and model), admin checklist P1, the user's answers to the spec's open questions (1 to 9 and Q-C1 to Q-C4), DPO D-3 (contact list). The pane steps (10b, 10c) also need F-012's tab shell merged.

## Configuration values

| Value | Example / format | Secret? | Where it lives |
|---|---|---|---|
| Tenant ID | GUID (already known, `c4333a40-…`) | No | Function env `ENTRA_TENANT_ID` (exists) |
| Drafts app client ID | GUID of "DIMSCALE Mail AI Drafts" | No | Function env `GRAPH_APP_CLIENT_ID` |
| Certificate SHA-256 thumbprint | base64url (`x5t#S256`) | No | Function env `GRAPH_CERT_X5T_S256` |
| Certificate private key | base64 of PKCS#8 DER, single line (V7) | **Yes** | Supabase function secret `GRAPH_CERT_PRIVATE_KEY` only |
| Webhook URL | `https://qepzpkseyzclxnqguqta.supabase.co/functions/v1/graph-webhook` | No | Function env `GRAPH_WEBHOOK_URL` |
| Graph `clientState` | 64 random characters | **Yes** | Function secret `GRAPH_CLIENT_STATE` |
| Cron shared secret | 64 random characters | **Yes** | Function secret `CRON_SHARED_SECRET`, and Vault secret `auto_draft_cron_secret` (read by `pg_cron`) |
| OpenAI API key | — | **Yes** | Function secret `OPENAI_API_KEY` (D5) |
| OpenAI model, `max_output_tokens` | from D5, `800` | No | Env `OPENAI_MODEL`, `OPENAI_MAX_OUTPUT_TOKENS` |
| Kill switch | `true` / anything else | No | Env `AUTO_DRAFTS_ENABLED` (default off) |
| Filters and caps | see spec | No | Env `MAX_AGE_HOURS`, `SKIP_SENDER_DOMAINS`, `INTERNAL_SENDERS`, `INTERNAL_DOMAINS`, `DAILY_CAP_PER_USER`, `DAILY_CAP_GLOBAL`, filter toggles |

The private key, `clientState`, cron secret and OpenAI key never go into git, `config.toml`, logs, chat transcripts or reports. The local `supabase/functions/.env` stays git-ignored.

## Manual prerequisites

### P0. Certificate (the user, on the dev machine, outside the repo)
Sources: [certificate credentials](https://learn.microsoft.com/entra/identity-platform/certificate-credentials).
Run this in **Git Bash** (it provides `openssl`, `base64` and `tr`), in a folder **outside** the repository (for example `~/mailai-cert/`). The validity is set by the admin's policy (admin question 3). 365 days is shown.
```
openssl req -x509 -newkey rsa:3072 -nodes -keyout drafts.key -out drafts.crt -days 365 -subj "/CN=DIMSCALE Mail AI Drafts"
openssl x509 -in drafts.crt -outform der -out drafts.cer
openssl pkcs8 -topk8 -nocrypt -in drafts.key -outform der | base64 -w0 > drafts.key.b64
openssl x509 -in drafts.crt -outform der | openssl dgst -sha256 -binary | base64 | tr '+/' '-_' | tr -d '='
```
1. Send **only `drafts.cer`** to the admin (P1.B).
2. Run `supabase secrets set GRAPH_CERT_PRIVATE_KEY="$(cat drafts.key.b64)"` and record the thumbprint (last command) as `GRAPH_CERT_X5T_S256`.
3. Then delete `drafts.key` and `drafts.key.b64`, or move them to the company password manager. Note the expiry date in `docs/status.md` (rotation reminder).

### P1. Tenant admin checklist
Moved to one consolidated PT-PT checklist for MMSI: [docs/admin/2026-10-23-mmsi-request.md](../admin/2026-10-23-mmsi-request.md). It covers the add-in app clean-up (revoke old grants, `access_as_user` and `Mail.ReadBasic` consent, Assignment required), the new "DIMSCALE Mail AI Drafts" app with certificate, the Exchange RBAC for Applications assignments scoped to the mail-enabled group (commands, verification, never `Mail.Send`), the emergency stop, and the later centralized deployment. The values the admin returns are recorded under Evidence (non-secret only).

### P2. Supabase (the user)
- Confirm the plan (Free: 150 s wall clock) (user question 8).
- Make sure the `pg_cron` and `pg_net` extensions are available in the project (Dashboard > Database > Extensions).

## Steps

| # | Step | ACs | Agent |
|---|---|---|---|
| 0 | Decisions: the user accepts ADR-0005 and answers spec questions 1 to 9; the DPO answers D5; the admin completes P1 and returns the values. The orchestrator records the non-secret values (IDs, thumbprint, expiry, group name) in this plan. | prerequisite | user, DPO, admin, orchestrator |
| 1 | Docs before code: `security.md` (app-only exception, RBAC scope, certificate secret, webhook `clientState`, cron secret, stored metadata and retention, the OpenAI terms from the spec), `overview.md` (diagram plus a "Mailbox (server, app-only)" capability row; the backend is no longer stateless), `requirements.md` (R8, if user question 7 is yes), `status.md` (F-013 row, new decisions). | — (SDD step 3) | requirements-architect |
| 2 | **Live spike** (after P0, P1 and the cache delay). A throwaway Deno script in the scratchpad, **not committed**, against the pilot mailbox. Its token logic is the same as step 4. It runs **one** `GET` message, **one** `createReply` with `comment` containing `<b>x</b>` and `message.categories`, a `GET` of the draft, `POST /subscriptions` to a temporary `graph-webhook` echo (or just the validation handler from step 9 deployed with the kill switch off), and a `GET` on a mailbox outside the group. Record the answers to V1 to V7 and the AC-28 live `403`. The script then `DELETE`s the test **subscription** (not mail). The user deletes the test draft by hand. Spec fallback paths are chosen from the results. If V3 fails, **stop and escalate** (R7). | AC-28 (part), V1 to V7 | coder (script) + user (runs it, deletes the draft) + orchestrator (records) |
| 3 | Migration in `supabase/migrations/`: the four tables (`auto_draft_users` with `status`, `auto_reply_contacts` unique `(user_id, address)`, `graph_subscriptions`, `auto_draft_queue` unique `(user_id, message_id)`), RLS enabled with no policies; `pg_cron` jobs that call `auto-draft-cron` with `pg_net` (every 2 min "process" including creation of missing subscriptions, every 6 h "renew", daily "retention") with the `x-cron-secret` read from Vault; the retention SQL (queue 30 days; user/contact rule). | AC-25, AC-40 | coder |
| 3b | Move the F-002 token validation into `supabase/functions/_shared/` unchanged; `whoami` imports it; its 24 tests pass untouched. Add the `oid`/`tid` requirement as a separate helper used only by stateful functions. | AC-36 | coder |
| 4 | Shared modules: `config` (env parsing, defaults, kill switch), `log` (fixed fields only), `appToken` (PS256 assertion with `jose` `SignJWT` and `importPKCS8`, using the same pinned `jose` version as `whoami`; token cache). Unit tests with a key pair generated in the test. | AC-9, AC-13, AC-23, AC-26 | coder |
| 5 | `graph` client: create, patch, delete and list subscriptions (`Prefer: IdType="ImmutableId"`, 409 handling); read a message (exact `$select` and both `Prefer` headers); list Inbox IDs since a time (paging); `createReply` (path chosen in step 2); `429` and `Retry-After` handling that returns a retry hint instead of sleeping. An allow-list of the only HTTP methods and paths the client may call (makes AC-19 enforceable in tests). | AC-6, AC-7, AC-8, AC-11, AC-18, AC-19, AC-21 | coder |
| 6 | `filters`: a pure `isAllowedSender(from, allowList)` (lower-case, trim) run on the `$select=from` read, then a pure function `(message, user, config, countsToday) → {pass} or {skip, reason}`. One test per filter, the allow-list match/non-match cases, the pass case and the caps. | AC-11, AC-12, AC-13, AC-14 | coder |
| 6a | **Amendment 2026-11-02, filters (pure code, in `filters`).** (1) Filter 12 default off: change the default of its on/off configuration value to off, keep the code and the `focused_other` reason. (2) Filter 6: implement exactly "owner's SHA-256 is among the SHA-256 of the `toRecipients` addresses" (lower-cased, same hashing as filter 5); CC-only, BCC-only or empty `toRecipients` skip with `not_direct_recipient`. (3) Filter 18 `auth_failed`: a pure `hasAuthFailure(internetMessageHeaders)` that scans every header named `Authentication-Results` (case-insensitive) for the tokens `dmarc=fail` or `compauth=fail` with the boundary rule of the spec amendment; any other input (undefined, empty, unparsable, other results) returns false (fail-open). Wire it after filters 3 to 10 and before 13 to 17, behind its on/off configuration value (default on, invalid value falls back with `config_default`). Add the new `$select` item only if missing: `internetMessageHeaders` is already in the list, so **no new Graph call**. One test per fixture of AC-45 and AC-46, the filter 12 and 6 cases of AC-44, the config cases of AC-48. | AC-44, AC-45, AC-46, AC-48, AC-50 | coder |
| 6b | **Amendment 2026-11-02, privacy and shape tests.** A test with marker strings in header values and in `Authentication-Results`: assert they are absent from the mocked OpenAI request body, captured `console` output, error bodies and DB writes (only the reason code is recorded). Assert the second read URL and both `Prefer` headers are unchanged. Re-run the grep of step 11 for the forbidden Graph methods after this change. | AC-47, AC-50, AC-51 | coder |
| 6c | **Amendment 2026-11-02, manual entry.** Reject the no-reply local parts (list in filter 11) in the pane's manual entry (step 10b) and in the `PUT` validation (step 10a), reusing one shared list constant per side. Tests: pane (text message) and function (422 `payload_invalid`, nothing written). Add a grep check that `Contacts.Read` and `/me/contacts` do not appear in `src/`, the manifests or code. | AC-43, AC-49 | coder |
| 7 | `ai`: the input builder (truncation, marker stripping, domain only), the Responses API call (`store: false`, no tools, strict schema, `max_output_tokens`, 30 s abort), and output validation. Tests inspect the exact request body. | AC-15, AC-16, AC-17, AC-22 | coder |
| 8 | `processor`: atomic claim with lease, one at a time per mailbox, state transitions and reason codes, retries and `next_attempt_at`, a time budget per invocation below the wall-clock limit. Tests include two concurrent claims (local Postgres). | AC-4, AC-10, AC-21 | coder |
| 9 | Function `graph-webhook`: validation handshake, `clientState` and subscription checks, change versus lifecycle routing, insert and `202` (`503` on insert failure), `EdgeRuntime.waitUntil(processor)`. `[functions.graph-webhook] verify_jwt = false`. | AC-1, AC-2, AC-3, AC-5 | coder |
| 10 | Function `auto-draft-cron`: `x-cron-secret` check; operations "process", "renew" (create, patch, delete, 10-min guard) plus a per-user catch-up, and "retention". `[functions.auto-draft-cron] verify_jwt = false`. | AC-5, AC-7, AC-8, AC-24 | coder |
| 10a | Function `auto-reply-settings`: shared auth (3b), CORS, `GET`/`PUT`/`DELETE` as in the spec, strict validation, atomic list replace, `mail_sha256`, logs with counts only. `[functions.auto-reply-settings] verify_jwt = false`. | AC-36, AC-37, AC-38, AC-39 | coder |
| 10b | Pane, Responder tab (needs F-012's tab shell merged): `Mail.ReadBasic` consent flow, candidate list from recent senders, manual entry, checkboxes, counter, switch, "Guardar", "Apagar tudo", status text, PT-PT strings in `strings.ts`. The Graph call lives in its own module; the backend request reuses the F-002 token request unchanged. | AC-32, AC-33, AC-34, AC-35 | coder |
| 10c | Amend F-002/AC-3 note ("no Graph scope in `src/`" becomes "only `Mail.ReadBasic`, in the Graph module") and add `Mail.ReadBasic` to the overview capability table. | — | requirements-architect |
| 10d | **Amendment 2026-11-02, spike addition (V8) for step 2.** In the same live spike, one extra `GET /users/{oid}/messages/{id}?$select=internetMessageHeaders` (one external test mail, one internal test mail). Record only: whether headers come back, the number of `Authentication-Results` headers, and the `dmarc` / `compauth` tokens. Never record addresses or content. The result goes under Evidence and may change nothing (fail-open stays). If the user can produce a forged-sender test, record the outcome. | AC-52, V8 | user + orchestrator (records) |
| 11 | Checks: `deno test`, `deno lint` and `deno check` for the new functions; root `npm run typecheck`, `lint`, `test` and `build`. Greps for `/send`, `/reply`, `replyAll`, `forward`, `sendMail`, `/move`, `/copy`, `DELETE` on messages, `Mail.Send`, and `console.` with content fields. Secret scan. Record the output below. | AC-19, AC-26, AC-27, AC-31 | qa-tester |
| 12 | Independent review: auth, tokens, webhook, prompt, data flow (pass `model: opus`). | all security ACs | code-reviewer |
| 13 | Fix Critical and High findings (or escalate). Rerun step 11. | — | coder, qa-tester |
| 14 | Deploy (the user's CLI): `supabase db push`; `supabase secrets set` for the config table values (kill switch **off**); `supabase functions deploy graph-webhook auto-draft-cron auto-reply-settings --use-api`; Vault secrets (SQL editor, never committed): `select vault.create_secret('<cron secret>','auto_draft_cron_secret');` and `select vault.create_secret('https://qepzpkseyzclxnqguqta.supabase.co','project_url');`. `ENTRA_TENANT_ID` must be set for `graph-webhook` (it fails closed without it). The pilot user enables the feature and picks contacts in the Responder tab (no hand-written SQL). Confirm that nothing happens with the kill switch off (AC-9 live), then set `AUTO_DRAFTS_ENABLED=true`. | AC-9 | user + orchestrator |
| 15 | Live tests: AC-41 (Responder tab, persistence); then with all Outlook clients closed: AC-29 (listed external sender → draft within 5 min, times recorded); AC-20; AC-30 (unlisted sender, newsletter header, out-of-office reply, CC-only, switch off, kill switch off); AC-6 (`GET /subscriptions`); AC-26 (one day of logs); webhook p95 (AC-3). Record the date, client, results and the open-question answers. | AC-3, AC-6, AC-18, AC-20, AC-26, AC-29, AC-30, AC-41 | user + orchestrator |
| 16 | Close: spec status, `status.md`, ADR-0005 status, cert expiry reminder, evidence complete. | — | orchestrator |

## Stop and rollback
- Feature off: `supabase secrets set AUTO_DRAFTS_ENABLED=false` (immediate). One user off: the user switches off in the Responder tab (or the operator runs `update auto_draft_users set enabled = false where …`). The next cron run deletes that user's subscription.
- Full removal: run "renew" with everyone disabled (deletes the subscriptions), unschedule the cron jobs, delete the two functions, and ask the admin to remove the role assignments and the certificate (P1.E).
- Drafts already created stay in users' Drafts. We never delete mail.

## Evidence
_Empty. Each step records the commands it ran, the real output, the dates and the live observations here._


### 2026-10-29: prerequisites prepared by the orchestrator (nothing sent to the tenant admin yet)
- **Certificate:** self-signed RSA 2048 `CN=DIMSCALE Mail AI Drafts`, valid until 2028-10-28, created on the user's PC outside the repo and OneDrive, in the folder `.dimscale-mailai-certs` of the user's home directory (it holds `drafts.crt` public, `drafts.key.pem` private and `cron-secret.txt`). Only `drafts.crt` goes to the admin. Renew before October 2028. The user should keep a backup of the private key.
- **Supabase secrets set (names only):** `GRAPH_CERT_PRIVATE_KEY`, `GRAPH_CERT_X5T_S256`, `GRAPH_CLIENT_STATE`, `CRON_SHARED_SECRET`, `GRAPH_WEBHOOK_URL`. **Still missing:** `GRAPH_APP_CLIENT_ID` (from the admin). `AUTO_DRAFTS_ENABLED` is not `true`.
- **Vault (pg):** `auto_draft_cron_secret` and `project_url` created.
- **Remote state check (read-only SQL):** the 16 `ad_*` functions exist with `search_path=""`, and the review-fix markers are present in `ad_claim`, `ad_counts_today`, `ad_user_status` and `ad_retention`. The 3 cron jobs exist and read the URL and the secret from Vault. EXECUTE is held only by `postgres` and `service_role`. The 5 tables have RLS on and no access for `anon` or `authenticated`. So the in-place migration edits were already in the applied version, and no follow-up migration is needed.
- **Decision:** the request to the tenant admin (MMSI) is sent only when the add-in is ready (user decision).

### 2026-11-03: build evidence (orchestrator)
- **Checks after the review and QA fixes:** `tsc` clean; `office-addin-lint` clean; `vitest` 41 files, 488 passed, 2 todo; `npm run build` compiled with 3 size warnings; `npm run validate` manifest valid; Deno `analyze-message` 63, `whoami` 24, `graph-webhook` 13, `auto-draft-cron` 6, `auto-reply-settings` 13, `_shared` 78 (with the analyze-message `deno.json`), all passed; `deno lint` clean in analyze-message, `_shared` and auto-reply-settings.
- **Reviews:** Opus review APPROVE WITH COMMENTS (no Critical or High; M1 manifest version fixed, M2 safety sentence added); QA found 4 defects (reason characters, truncation line), all fixed; a focused re-review of the fixes was requested.
- **Deploy:** `supabase functions deploy` for `analyze-message`, `auto-reply-settings`, `graph-webhook` and `auto-draft-cron` (`--use-api`), before any pane reaches `main`. Smoke tests with `curl`: OPTIONS from `https://dimscale.github.io` and `https://localhost:3000` return 204 with the origin, from `https://evil.example` 204 without `Access-Control-Allow-Origin`; no token or a garbage token returns 401; GET returns 405 (`analyze-message`, `graph-webhook`); `auto-reply-settings` PUT without token 401; cron without or with a wrong secret 401; `whoami` without token 401.
- **Not done:** merge to `main` (waits for the user's live test), live checks (F-007 AC-16, AC-18 to AC-21; F-013 AC-52 / V8), screen reader, 320 px.
