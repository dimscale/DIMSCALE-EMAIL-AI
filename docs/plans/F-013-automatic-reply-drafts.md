# Plan F-013: Automatic reply drafts (server side)

- **Spec:** [F-013](../specs/F-013-automatic-reply-drafts.md). **ADR:** [ADR-0005](../decisions/0005-server-side-mail-processing.md).
- **Status:** Draft. Nothing in here may start until the user approves the spec and accepts ADR-0005. Exceptions: step 0 (decisions) and the admin checklist (P1) can be sent in advance if the user wants.
- **Blocked by:** ADR-0005 acceptance, D5 (the DPO, OpenAI account and model), admin checklist P1, the user's answers to the spec's open questions 1 to 9.

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

### P1. Tenant admin checklist (ready to send)

> **Subject: DIMSCALE Mail AI: app-only access for automatic reply drafts (no send permission)**
>
> Hello. For DIMSCALE Mail AI we need a server-side app that creates **reply drafts** in the Drafts folder of a small pilot group. It must **never send** email. Please do the steps below and send us the values in **[return]**. Microsoft references: https://learn.microsoft.com/exchange/permissions-exo/application-rbac and https://learn.microsoft.com/entra/identity-platform/certificate-credentials
>
> **A. App registration (Entra admin center > App registrations > New registration)**
> 1. Name `DIMSCALE Mail AI Drafts`. Supported account types: **this organizational directory only (single tenant)**. **No redirect URI.**
> 2. API permissions: **remove** the default Microsoft Graph `User.Read`. **Do not add or consent any Microsoft Graph application permission** (no `Mail.Read`, `Mail.ReadWrite`, `Mail.Send` in Entra). Mail access is granted in Exchange (step C), scoped to the group. An Entra consent would be tenant-wide and would cancel the scoping.
> 3. Certificates & secrets > Certificates > **Upload certificate** > the `drafts.cer` file we sent. Do **not** create a client secret.
> 4. **[return]** Application (client) ID; the certificate thumbprint and expiry shown after upload.
> 5. **[return]** From **Enterprise applications** > `DIMSCALE Mail AI Drafts`: the **Object ID** of the service principal. This is not the Object ID on the App registrations page.
>
> **B. Group**
> 6. A **mail-enabled** group (a Microsoft 365 group, mail-enabled security group or distribution list) named for example `DIMSCALE Mail AI Auto-Drafts`, with the pilot users as **direct** members. Nested groups are not counted. **[return]** the group name and the pilot users.
>
> **C. Exchange Online PowerShell (member of Organization Management; Exchange Administrator in Entra)**
> ```powershell
> Connect-ExchangeOnline
> New-ServicePrincipal -AppId <Application (client) ID> -ObjectId <Enterprise app Object ID> -DisplayName "DIMSCALE Mail AI Drafts"
> $dn = (Get-Group "DIMSCALE Mail AI Auto-Drafts").DistinguishedName
> New-ManagementScope -Name "DIMSCALE Mail AI Auto-Drafts users" -RecipientRestrictionFilter "MemberOfGroup -eq '$dn'"
> New-ManagementRoleAssignment -App <Enterprise app Object ID> -Role "Application Mail.Read" -CustomResourceScope "DIMSCALE Mail AI Auto-Drafts users"
> New-ManagementRoleAssignment -App <Enterprise app Object ID> -Role "Application Mail.ReadWrite" -CustomResourceScope "DIMSCALE Mail AI Auto-Drafts users"
> ```
> 7. **Never assign** `Application Mail.Send`, `Application Mail Full Access`, `Application Exchange Full Access`, `Application EWS.AccessAsApp` or `Application SMTP.SendAsApp`.
> 8. Verify, then **[return]** the output of all three:
> ```powershell
> Test-ServicePrincipalAuthorization -Identity <Enterprise app Object ID> -Resource <pilot user mailbox>      # expect InScope = True for both roles
> Test-ServicePrincipalAuthorization -Identity <Enterprise app Object ID> -Resource <a mailbox NOT in the group> # expect InScope = False
> Get-ManagementRoleAssignment -RoleAssignee <Enterprise app Object ID> | Format-Table Name,Role,CustomResourceScope
> ```
> 9. Changes take **30 minutes to 2 hours** to apply (Exchange cache).
>
> **D. Still pending from before (unchanged)**
> 10. On the add-in app "DIMSCALE AI Email Assistant": revoke the old tenant-wide delegated grants `Mail.ReadWrite`, `offline_access` and `User.Read`. Grant admin consent for `access_as_user`. Set Assignment required with the group (F-002 AC-21).
>
> **E. Emergency stop (keep for reference)**
> - Fastest on our side: we switch the feature off in the backend (takes effect at once).
> - Your side: delete the certificate from the app registration, so no new tokens can be issued (existing tokens expire within about 1 hour). And/or `Remove-ManagementRoleAssignment` for both assignments (takes effect within 30 minutes to 2 hours).
>
> **Questions:** can you use certificates (validity and rotation policy)? Is there any Conditional Access for workload identities? Is the F-002 sign-in group mail-enabled and could it be reused? Who is the emergency contact?

### P2. Supabase (the user)
- Confirm the plan (Free: 150 s wall clock) (user question 8).
- Make sure the `pg_cron` and `pg_net` extensions are available in the project (Dashboard > Database > Extensions).

## Steps

| # | Step | ACs | Agent |
|---|---|---|---|
| 0 | Decisions: the user accepts ADR-0005 and answers spec questions 1 to 9; the DPO answers D5; the admin completes P1 and returns the values. The orchestrator records the non-secret values (IDs, thumbprint, expiry, group name) in this plan. | prerequisite | user, DPO, admin, orchestrator |
| 1 | Docs before code: `security.md` (app-only exception, RBAC scope, certificate secret, webhook `clientState`, cron secret, stored metadata and retention, the OpenAI terms from the spec), `overview.md` (diagram plus a "Mailbox (server, app-only)" capability row; the backend is no longer stateless), `requirements.md` (R8, if user question 7 is yes), `status.md` (F-013 row, new decisions). | — (SDD step 3) | requirements-architect |
| 2 | **Live spike** (after P0, P1 and the cache delay). A throwaway Deno script in the scratchpad, **not committed**, against the pilot mailbox. Its token logic is the same as step 4. It runs **one** `GET` message, **one** `createReply` with `comment` containing `<b>x</b>` and `message.categories`, a `GET` of the draft, `POST /subscriptions` to a temporary `graph-webhook` echo (or just the validation handler from step 9 deployed with the kill switch off), and a `GET` on a mailbox outside the group. Record the answers to V1 to V7 and the AC-28 live `403`. The script then `DELETE`s the test **subscription** (not mail). The user deletes the test draft by hand. Spec fallback paths are chosen from the results. If V3 fails, **stop and escalate** (R7). | AC-28 (part), V1 to V7 | coder (script) + user (runs it, deletes the draft) + orchestrator (records) |
| 3 | Migration in `supabase/migrations/`: the three tables, unique `(user_id, message_id)`, RLS enabled with no policies; `pg_cron` jobs that call `auto-draft-cron` with `pg_net` (every 2 min "process", every 6 h "renew", daily "retention") with the `x-cron-secret` read from Vault; the retention SQL. | AC-25 | coder |
| 4 | Shared modules: `config` (env parsing, defaults, kill switch), `log` (fixed fields only), `appToken` (PS256 assertion with `jose` `SignJWT` and `importPKCS8`, using the same pinned `jose` version as `whoami`; token cache). Unit tests with a key pair generated in the test. | AC-9, AC-13, AC-23, AC-26 | coder |
| 5 | `graph` client: create, patch, delete and list subscriptions (`Prefer: IdType="ImmutableId"`, 409 handling); read a message (exact `$select` and both `Prefer` headers); list Inbox IDs since a time (paging); `createReply` (path chosen in step 2); `429` and `Retry-After` handling that returns a retry hint instead of sleeping. An allow-list of the only HTTP methods and paths the client may call (makes AC-19 enforceable in tests). | AC-6, AC-7, AC-8, AC-11, AC-18, AC-19, AC-21 | coder |
| 6 | `filters`: a pure function `(message, user, config, countsToday) → {pass} or {skip, reason}`. One test per filter, plus the pass case and the caps. | AC-12, AC-13, AC-14 | coder |
| 7 | `ai`: the input builder (truncation, marker stripping, domain only), the Responses API call (`store: false`, no tools, strict schema, `max_output_tokens`, 30 s abort), and output validation. Tests inspect the exact request body. | AC-15, AC-16, AC-17, AC-22 | coder |
| 8 | `processor`: atomic claim with lease, one at a time per mailbox, state transitions and reason codes, retries and `next_attempt_at`, a time budget per invocation below the wall-clock limit. Tests include two concurrent claims (local Postgres). | AC-4, AC-10, AC-21 | coder |
| 9 | Function `graph-webhook`: validation handshake, `clientState` and subscription checks, change versus lifecycle routing, insert and `202` (`503` on insert failure), `EdgeRuntime.waitUntil(processor)`. `[functions.graph-webhook] verify_jwt = false`. | AC-1, AC-2, AC-3, AC-5 | coder |
| 10 | Function `auto-draft-cron`: `x-cron-secret` check; operations "process", "renew" (create, patch, delete, 10-min guard) plus a per-user catch-up, and "retention". `[functions.auto-draft-cron] verify_jwt = false`. | AC-5, AC-7, AC-8, AC-24 | coder |
| 11 | Checks: `deno test`, `deno lint` and `deno check` for the new functions; root `npm run typecheck`, `lint`, `test` and `build`. Greps for `/send`, `/reply`, `replyAll`, `forward`, `sendMail`, `/move`, `/copy`, `DELETE` on messages, `Mail.Send`, and `console.` with content fields. Secret scan. Record the output below. | AC-19, AC-26, AC-27, AC-31 | qa-tester |
| 12 | Independent review: auth, tokens, webhook, prompt, data flow (pass `model: opus`). | all security ACs | code-reviewer |
| 13 | Fix Critical and High findings (or escalate). Rerun step 11. | — | coder, qa-tester |
| 14 | Deploy (the user's CLI): `supabase db push`; `supabase secrets set` for the config table values (kill switch **off**); `supabase functions deploy graph-webhook auto-draft-cron --use-api`; Vault secret `auto_draft_cron_secret`; insert the pilot user row (`entra_oid`, `mail_sha256`, `enabled = true`). Confirm that nothing happens with the kill switch off (AC-9 live), then set `AUTO_DRAFTS_ENABLED=true`. | AC-9 | user + orchestrator |
| 15 | Live tests with all Outlook clients closed: AC-29 (external email → draft within 5 min, with times recorded); AC-20; AC-30 (newsletter header, out-of-office reply, CC-only, kill switch off); AC-6 (`GET /subscriptions`); AC-26 (inspect one day of logs); webhook p95 (AC-3). Record the date, client, results and the open-question answers. | AC-3, AC-6, AC-18, AC-20, AC-26, AC-29, AC-30 | user + orchestrator |
| 16 | Close: spec status, `status.md`, ADR-0005 status, cert expiry reminder, evidence complete. | — | orchestrator |

## Stop and rollback
- Feature off: `supabase secrets set AUTO_DRAFTS_ENABLED=false` (immediate). One user off: `update auto_draft_users set enabled = false where …`. The next renewal deletes that user's subscription.
- Full removal: run "renew" with everyone disabled (deletes the subscriptions), unschedule the cron jobs, delete the two functions, and ask the admin to remove the role assignments and the certificate (P1.E).
- Drafts already created stay in users' Drafts. We never delete mail.

## Evidence
_Empty. Each step records the commands it ran, the real output, the dates and the live observations here._
