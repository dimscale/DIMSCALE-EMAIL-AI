# Security audit: "imagine someone tries to steal these emails"

- **Date:** 2026-11-03
- **Branch / commit:** `dev` at `9eaa7a1` (working tree also has an uncommitted `prototype:ui` script and `src/taskpane/prototype/`, neither in a production build)
- **Type:** read-only, attacker-centric threat model. Nothing was edited, deployed or opened that holds a secret. No email content, token or key value appears here.
- **Reviewer:** independent code-reviewer agent (Opus)

## Legend
- **VERIFIED**: seen in code, config, git history, or a public endpoint queried read-only (file:line or command given).
- **ASSUMED / UNKNOWN**: a cloud setting I cannot see (Supabase dashboard, Entra, Exchange, Cloudflare, GitHub org settings, the owner's PC encryption). Stated as a question for the owner or MMSI, never as a fact.

## Bottom line
1. **The code itself is in good shape against mail theft.** No code path stores or logs email content, the backend validates Entra tokens strictly, Graph calls are allow-listed, AI output cannot carry links or markup, and no secret is in the public repo or its history.
2. **The real theft paths are around the code, not in it.** Four things decide who can steal mail:
   - the Drafts certificate private key and the PC that holds a copy;
   - who can push code that runs inside every user's Outlook (unprotected `main` and `dev` on a public repo);
   - the old delegated `Mail.ReadWrite` grant that may still sit on the add-in's app registration;
   - the Supabase account, which holds every secret.
3. **No Critical finding is open today**, mainly because the Drafts certificate and RBAC are not live yet (D14) and no manifest has gone to MMSI. Two High findings **become Critical**:
   - H1, the moment MMSI uploads the certificate and assigns the RBAC roles;
   - H2 together with H3, if the legacy `Mail.ReadWrite` consent exists when the pane is deployed to all users.

## 1. Who can read mail content today, and how

| Who / what | What it can see | How | Kept where, how long | Evidence |
|---|---|---|---|---|
| **Task pane** (in the user's Outlook only) | Subject, sender name and address, date and plain-text body of the **open** item. With a pinned pane that is every email the user reads. | Office.js `ReadItem` (`item.body.getAsync` Text) | In memory only. The only things persisted are six booleans in `roamingSettings`, with no content. | VERIFIED `manifest.xml:32`, `src/taskpane/MessagePanel.tsx:110`, `src/taskpane/messageView.ts:62-80`; grep found no `localStorage`/`sessionStorage`/`console.` in non-test `src/` |
| Pane, "Lista" tab | `from` of up to 300 Inbox messages from the last 30 days (no body) | Delegated Graph `Mail.ReadBasic` token, sent only to `graph.microsoft.com`; `nextLink` host-checked | In memory | VERIFIED `src/taskpane/graph.ts:5,80-137` |
| **analyze-message** function | Subject ≤ 200, sender display name ≤ 100, sender **domain** (never the address), body ≤ 4,000 incl. quoted history; for refine also instruction ≤ 300 and the draft ≤ 2,000 | `POST` from the pane after Entra token validation | **Nothing stored.** The DB holds only a counter keyed by an unsalted SHA-256 of `oid`. Logs hold `fn,id,method,status,reason,ms` only. | VERIFIED `supabase/functions/analyze-message/handler.ts:40-76,143,172`, `_shared/log.ts:1-19`, migration `20261029120000_f003_analyze_usage.sql:4-9` |
| **OpenAI** | Exactly what analyze-message or F-013 forwards | Responses API, `store: false`, no tools, strict JSON schema, `api.openai.com` (not the EU endpoint) | Not stored as application state. OpenAI abuse-monitoring logs up to 30 days (as recorded in `security.md:35`, not re-checked today). | VERIFIED `_shared/ai.ts:149-160` |
| Auto mode (F-003 (i)) | Same fields, for **every email opened on a tab whose "Automático" switch is on**, no click | Pane flag `AUTO_MODE_AVAILABLE = true` | As above | VERIFIED `src/taskpane/config.ts:19`; D5 still open (`status.md:77`) |
| **Supabase Postgres** | **No subjects, bodies, display names, AI output or tokens.** Stored: `auto_draft_users` (`entra_oid`, `mail_sha256` = unsalted SHA-256 of the user's own address, flags, status, timestamps); `auto_reply_contacts` (the chosen contact **addresses in plain text**, third-party personal data); `graph_subscriptions` (ids, expiry); `auto_draft_queue` (Graph **immutable message IDs**, status, reason, attempts); `analyze_usage` (hash, day, count). RLS on, no policies, `revoke all` from `anon`/`authenticated`, RPCs granted to `service_role` only. | Functions use the secret key | Queue 30 days; contacts until removed or 90/30 days inactive. Backups not verified. | VERIFIED migration `20261023120000_f013_auto_reply_drafts.sql:7-56,201-223`; `20261029120000_f003_analyze_usage.sql:11-13,32-33` |
| **F-013 server path** (not live: waits for MMSI D14; kill switch `AUTO_DRAFTS_ENABLED` must be exactly `"true"`) | **What the code reads:** (1) `from` only; (2) for allow-listed senders only: `subject,from,toRecipients,isDraft,inferenceClassification,receivedDateTime,internetMessageHeaders,uniqueBody` (text),`parentFolderId`. **What the token could read:** everything `Application Mail.Read` + `Application Mail.ReadWrite` allow in every mailbox in the RBAC scope (all folders, full bodies, attachments, move and delete; not send). | App-only token, PS256 certificate assertion | In memory only; message ID + status in the queue 30 days; draft created in the user's own Drafts via `createReply` only | VERIFIED `_shared/graph.ts:8-17,90,94-129,157-163`, `_shared/processor.ts:77-116`, `_shared/config.ts:18`, `_shared/appToken.ts:5-53` |
| GitHub, Cloudflare | No mail content | n/a | n/a | They **control the code** that reads mail (H2, M2) |

The Graph allow-list in `_shared/graph.ts:8-17` protects only against our own bugs. Anyone who holds the certificate key calls Graph directly and is limited only by the Exchange RBAC scope.

## 2. Credentials: what a thief gains, how to notice, how to revoke

| Credential | Where it lives | A thief gains | Which emails | Detect | Revoke (fastest first) |
|---|---|---|---|---|---|
| **Drafts app certificate private key** | Supabase secret `GRAPH_CERT_PRIVATE_KEY` (ASSUMED set) **and a file `drafts.key.pem` in `C:\Users\thiago.santos\.dimscale-mailai-certs`** (VERIFIED name, size and ACL only; not opened) | App-only Graph token for "DIMSCALE Mail AI Drafts" from anywhere on the internet, about 1 h per token, renewable as long as the cert is registered | **Every message, every folder, attachments, in every mailbox that is a direct member of the RBAC group**, plus modify, move and delete (`Application Mail.ReadWrite`). No send. Today: none, if MMSI has not uploaded the cert or assigned the roles (UNKNOWN, D14 open). | Entra sign-in logs > **Service principal sign-ins** for the Drafts app from IPs that are not Supabase/AWS eu-west-1; Purview audit `MailItemsAccessed` by the app (licence-dependent, verify); `GET /subscriptions` showing unknown `notificationUrl`s | MMSI: (1) disable the Drafts enterprise app (stops new tokens); (2) delete the certificate from the app registration; (3) `Remove-ManagementRoleAssignment` for both roles (30 min to 2 h). Existing tokens live up to about 1 h. Then set `AUTO_DRAFTS_ENABLED=false`. |
| `CRON_SHARED_SECRET` | Supabase secret, Supabase Vault, **and `cron-secret.txt` on the owner's PC** (VERIFIED name only) | Can trigger `process`/`renew`/`retention` runs. Responses are only `{ok:true}`. | None (no data returned) | Function logs: unexpected `auto-draft-cron` rate | New value in the secret and in Vault |
| `GRAPH_CLIENT_STATE` | Supabase secret | Combined with a subscription ID and the public tenant ID: forged notifications that enqueue message IDs for an existing user. Every filter still applies, and drafts go only into that user's own mailbox. | None returned | Logs `client_state_mismatch`, `unknown_subscription`, `tenant_mismatch` | Rotate and recreate subscriptions (clientState is set at creation) |
| `OPENAI_API_KEY` | Supabase secret | Paid API use. **No access to past content** (`store:false` on every call). | None | OpenAI usage dashboard, budget alerts | Revoke in OpenAI, set a new secret (read per request, no redeploy) |
| Supabase **secret / service-role key** | Supabase (auto-injected into functions) | DB access: contact lists (third-party addresses), oids, hashed own addresses, message IDs. No content. | None directly (IDs are useless without a Graph token) | Supabase API logs | Rotate the secret keys |
| **Supabase account / CLI access token** | Owner's account; CLI token in Windows Credential Manager (VERIFIED entry `Supabase CLI:supabase`) | **Everything above at once**: deploy a function that dumps env (cert key, OpenAI key), or replace `analyze-message` to copy every body users send. | All mail sent to the AI, plus everything the cert reaches | Unexpected function deploy versions; org audit (plan-dependent) | Revoke CLI tokens, rotate the password and MFA, then rotate **all** function secrets and the cert |
| **GitHub credential** with push to `main`/`dev` | Owner's account; Git credential in Windows Credential Manager (VERIFIED entry `git:https://github.com`) | Ship pane code to every Outlook that loads the add-in (H2) | Every email opened with the pane, plus mailbox metadata (`Mail.ReadBasic`), plus full mailbox if H3's legacy grant is consented | GitHub push notifications, Cloudflare/Pages deploy notifications | Revoke sessions and PATs, revert, redeploy, revoke the SPA's Graph consent (H3), revoke user sessions |
| User Entra tokens in the pane | Obtained by MSAL NAA. MSAL default cache is `sessionStorage` (VERIFIED `node_modules/@azure/msal-browser/dist/config/Configuration.mjs:59`, no override in `src/taskpane/auth.ts:15-21`). No refresh token in our code: `SCOPES` has no `offline_access` (`config.ts:10`), no OBO, nothing stored server-side. | `access_as_user` token (bearer, about 1 h): call our functions as that user, i.e. read or replace their contact list, enable auto-drafts, and up to 50 analyze calls per day. `Mail.ReadBasic` token: metadata of their whole mailbox. | That user's own mailbox metadata | 401 reason spikes | Revoke the user's sessions (`Revoke-MgUserSignInSession`) |

**Not secrets** (fine to be public; VERIFIED in `src/taskpane/config.ts:2-7`): tenant ID, SPA client ID, Supabase project ref and function URLs, add-in IDs. Possessing them grants nothing. Every function demands an Entra token for this exact client (`_shared/entra.ts:39-49`, `_shared/stateful.ts:12-13`) or a shared secret.

## 3. Findings (ranked)

### H1. The Drafts certificate private key and the cron secret sit in plain files on the owner's PC, next to GitHub and Supabase credentials. High; becomes Critical once MMSI activates the cert and RBAC.
- **Attack scenario.** Any of these copies `drafts.key.pem`:
  - an infostealer;
  - a malicious npm install script run during development;
  - a prompt-injected AI coding agent with shell access;
  - any local administrator of this managed PC (`BUILTIN\Administradores` has Full control).

  From any machine they then request app-only Graph tokens and silently read every mailbox in the RBAC group until the cert is removed. The same PC holds the Git credential (pane code, H2) and the Supabase CLI token (all secrets, H4), so one machine compromise gives everything.
- **Evidence (VERIFIED, metadata only).**
  - `ls -la ~/.dimscale-mailai-certs`: `drafts.key.pem` (1,732 bytes), `drafts.crt`, `cron-secret.txt`, dated 2026-10-29.
  - `icacls`: SYSTEM, Administradores, the user.
  - `docs/plans/F-013-automatic-reply-drafts.md:33` generates the key with `-nodes` (no passphrase), and `:41` says "Then delete `drafts.key` and `drafts.key.b64`, or move them to the company password manager". That was not done.
  - The folder is outside the OneDrive-synced folder (good).
  - BitLocker status could not be read without admin (UNKNOWN).
- **Impact.** All mail (read, modify, delete; not send) of every RBAC-group member.
- **Likelihood.** Low to medium. Developer PCs are the most attacked asset, and this one runs AI agents and dev tooling with install scripts.
- **Mitigation (S).**
  1. Confirm the Supabase secret is set and works (live spike, or `supabase secrets list` shows the name).
  2. Delete `drafts.key.pem` and `cron-secret.txt`, or move them into the company password manager. Keep `drafts.crt` (public).
  3. At rotation, generate the key, set it in Supabase and delete it in the same session.
  4. Optionally (M, needs Workload Identities Premium): Conditional Access for the Drafts service principal by location, after confirming Supabase egress IPs are stable.

### H2. Unprotected `main` and `dev` in a public repo decide what code runs inside users' Outlook. High.
- **Attack scenario.** Someone with push rights, or with a stolen GitHub session or PAT, pushes a small change that sends `item.body.getAsync` output, plus a silently acquired Graph token, to an outside server.
  - `main` deploys automatically to GitHub Pages (today's production pane) and Cloudflare production.
  - `dev` deploys to `https://dev.dimscale-mail-ai.pages.dev/`, a registered NAA redirect origin (ADR-0007 step 6), so that code can obtain tokens for the SPA client.
  - No review, status check or second person is needed.
- **Evidence (VERIFIED).**
  - Public GitHub API: `branches/main` and `branches/dev` both `"protected": false`; `rules/branches/main` returns an empty list.
  - The `github-pages` environment allows `main` only, but anyone who can push to `main` passes that.
  - `.github/workflows/pages.yml:3-6,34-37` deploys on push to `main`.
  - `webpack.config.js:10-20` builds `dev` as the test add-in.
  - `curl -I` on the three hosts: **no Content-Security-Policy**.
  - One public org member and only the owner's commits on `dev`: today the practical attacker is whoever takes over the owner's GitHub account.
- **Impact.** Every email any user opens while the pane is open, plus subjects, senders and recipients of their whole mailbox (`Mail.ReadBasic`, if admin-consented). With H3, full mailbox read and delete. Today: the owner's mailbox (sideloaded, no MMSI deployment yet, ADR-0007 Q-9). After the MMSI upload: every assigned user.
- **Likelihood.** Low to medium (account takeover or a supply-chain push).
- **Mitigation (S).**
  1. GitHub ruleset on `main` **and** `dev`: require a pull request and the `build` check, block force-push and deletion, restrict who can push.
  2. Require 2FA for the `dimscale` org (UNKNOWN today).
  3. Push and deploy notifications.
  4. Optionally (M), signed commits.
  5. Make the repo private (ADR-0007 step 14).
  6. A CSP (L4) does not stop a malicious deploy, but it blocks injected-script exfiltration.

### H3. The add-in's app registration may still carry delegated `Mail.ReadWrite`, `offline_access` and `User.Read`, and "Assignment required" is not on yet. High (cloud setting, UNKNOWN).
- **Attack scenario.** If those legacy grants are consented, any code running on a registered redirect origin can silently get a Graph token that reads **and deletes the entire mailbox** of whoever opens it. The registered origins are:
  - `dimscale.github.io`, shared by every Pages site of the `dimscale` org;
  - the `dev` Cloudflare alias;
  - `localhost:3000`.

  Without "Assignment required", every account in the tenant can get tokens for the SPA (B2B guests probably too, not verified).
- **Evidence.**
  - `docs/admin/2026-10-23-mmsi-request.md:16,29,32` (A1: remove `Mail.ReadWrite`, `User.Read`, `offline_access`; A4: Assignment required), still unchecked.
  - `docs/status.md:47,86` (D14 pending), `docs/decisions/0005-server-side-mail-processing.md:81`.
  - VERIFIED that the code needs none of these grants: `src/taskpane/config.ts:10` (`access_as_user`, `profile`), `src/taskpane/graph.ts:5` (`Mail.ReadBasic`). Removing them breaks nothing.
- **Impact.** Full mailbox of every pane user, combined with H2 or M3.
- **Mitigation (S, MMSI).** Do checklist items A1 to A4 now, ahead of the rest of D14. Verify under Enterprise apps > Permissions, **both** the Admin consent and the User consent tabs, that only `access_as_user`, `openid`, `profile` and `Mail.ReadBasic` remain. Remove `brk-multihub://localhost:3000` from the production registration, or move dev to a separate registration (M).

### H4. The Supabase account is the master key. High (cloud setting, UNKNOWN).
- **Attack scenario.** Taking over the Supabase account, an org member, or the CLI token on the PC lets the attacker deploy a new `analyze-message` that copies every body users send (auto mode makes this continuous), or a function that reads `GRAPH_CERT_PRIVATE_KEY` from env (H1 impact without touching the PC).
- **Evidence.**
  - VERIFIED: secrets are read from env at runtime (`_shared/appToken.ts:55-58`, `analyze-message/index.ts:6,13`).
  - VERIFIED: the CLI credential exists in Windows Credential Manager.
  - UNKNOWN: MFA, org members, SSO, audit logs on the current plan (ADR-0007 Q-8).
- **Impact.** All mail sent to the AI, plus all RBAC mailboxes once F-013 is live.
- **Likelihood.** Low.
- **Mitigation (S).** MFA (TOTP or hardware key) on the Supabase account, a single owner, review org members, `supabase logout` when not deploying, a DIMSCALE mailbox as the account email.

### M1. The RBAC group defines the certificate's reach, and its membership is an insider lever. Medium (UNKNOWN).
- **Scenario.**
  - Whoever can add a direct member to the mail-enabled group (group owners, Exchange/Groups admins at MMSI, and **anyone at all if it is a public Microsoft 365 group**) silently extends what the Drafts cert, and whoever holds it, can read.
  - An Exchange admin could also widen the management-scope filter or assign a broader role.
- **Evidence.** `docs/admin/2026-10-23-mmsi-request.md:40,56-58,60`.
- **Mitigation (S).**
  - Use a mail-enabled security group or a closed distribution list (`MemberJoinRestriction Closed`), with owners limited to MMSI plus the owner.
  - Alert on membership changes (Purview audit "Add member to group").
  - Re-run `Test-ServicePrincipalAuthorization` and `Get-ManagementRoleAssignment` monthly.
  - Keep the pilot small.

### M2. Cloudflare Pages build and account. Medium (mostly UNKNOWN).
- **Scenario.** A tampered build artifact means pane code in Outlook (H2 impact). Ways it can happen:
  - Cloudflare account takeover.
  - Over-broad permissions of the Cloudflare GitHub App.
  - A dev-dependency install script running in Cloudflare's build. Cloudflare's build image installs dependencies itself before the build command, so `--ignore-scripts` in the command may not cover that install (verify; Cloudflare documents `SKIP_DEPENDENCY_INSTALL`).
  - A build that skips the gates.
  - Previews are public and built for every branch unless restricted (ADR-0007 table, line 48).
- **Evidence.** `webpack.config.js:9-20` (the closed host list is good). `curl -I` shows `Access-Control-Allow-Origin: *`, which is harmless for static files.
- **Mitigation (S).**
  - Preview branches: `dev` only.
  - Set `SKIP_DEPENDENCY_INSTALL=1` and use the ADR-0007 step 3 build command (`npm ci --ignore-scripts` plus all gates).
  - MFA on the Cloudflare account, which should be owned by a DIMSCALE mailbox.
  - GitHub App on "Only select repositories".
  - Deploy notifications.

### M3. The shared `dimscale.github.io` origin is still a token-capable origin. Medium.
- **Scenario.** Any org member can publish another Pages site on the same origin. A user who is tricked into sideloading a manifest for it (if user sideloading is allowed, UNKNOWN) runs code that gets NAA tokens for our SPA client.
- **Evidence.** `docs/architecture/security.md:67`, ADR-0007 plan step 14, `webpack.config.js:17` (the default is still github.io).
- **Mitigation.** Finish the ADR-0007 cut-over and remove the `dimscale.github.io` redirect and origin (M, planned). In the meantime, MMSI blocks user sideloading of custom add-ins (S).

### M4. Auto mode sends real mail to OpenAI without a click while D5 is open. Medium (privacy/compliance, not a theft path).
- **Evidence.** VERIFIED `src/taskpane/config.ts:19`. Minimisation is verified: domain only, no recipients or IDs (`analyze-message/handler.ts:47-51`), and `store:false` (`_shared/ai.ts:156`).
- **Mitigation (S).**
  - Keep per-user opt-in.
  - Use a project-scoped OpenAI key with a budget limit and MFA on the OpenAI org.
  - Consider the EU endpoint with ZDR once D5 allows it.

### M5. A spoofed listed contact plus `Reply-To` can produce a draft addressed to the attacker. Medium (needs a user send).
- **Scenario.** A forged `From` of an allow-listed contact passes filter 18 when there is no explicit `dmarc=fail`/`compauth=fail` (fail-open by design). A `Reply-To` header points to the attacker. `createReply` (F-013) and `displayReplyFormAsync` (pane) follow Outlook reply semantics, so the draft is probably addressed to `Reply-To` (NOT VERIFIED for Graph `createReply`). If the user adds details and sends, they go to the attacker.
- **Evidence.** `_shared/filters.ts:20-36,61`, `_shared/processor.ts:113`, `_shared/graph.ts:90,157-163` (`replyTo` not selected or checked), `docs/specs/F-013-automatic-reply-drafts.md:247` (R10).
- **Mitigation (S).** Add `replyTo` to the existing `$select` and skip when it differs from `from` (a new reason code, no data-flow change). Verify the behaviour in the live spike.

### L1. A stolen `access_as_user` token can alter the user's own auto-draft settings. Low.
- `auto-reply-settings` returns and replaces the user's contact list (third-party addresses) and can enable drafting for an attacker-chosen contact. It still reaches only that user's own mailbox.
- Evidence: `auto-reply-settings/handler.ts:69-83`.
- The token is short-lived and goes only to Supabase. Nothing to do beyond H3 (Assignment required).

### L2. Forged webhook notifications (requires `GRAPH_CLIENT_STATE`). Low.
- Defence in depth is VERIFIED: the clientState is compared in constant time first, then the tenant, then a known subscription ID (`graph-webhook/handler.ts:53-68`, `_shared/safe.ts:9-16`). Duplicates collapse (`ad_enqueue` unique key, migration lines 140-147). Nothing is returned to the caller.
- The `validationToken` echo is `text/plain` + `nosniff` (`graph-webhook/handler.ts:26-29`).

### L3. Webhook body is read fully before the size check. Low (availability only).
- `graph-webhook/handler.ts:32-35` reads the whole request (`req.text()`) before checking 512 KiB; analyze-message streams with a cap (`handler.ts:20-35`).
- There is no app-level rate limit on public functions; Supabase platform limits apply.
- Fix: reuse the capped reader (S).

### L4. No CSP or security headers on the pane hosts. Low (defence in depth).
- VERIFIED with `curl -I` on github.io, `dimscale-mail-ai.pages.dev`, `dev.dimscale-mail-ai.pages.dev`.
- Add a Cloudflare `_headers` with `connect-src` limited to `login.microsoftonline.com`, the Supabase project host and `graph.microsoft.com`, and `script-src 'self' appsforoffice.microsoft.com`. Test live in Outlook first (Office.js and NAA have their own needs). S to M.

### L5. `.gitignore` gaps. Low.
- It covers `.env*`, `*.pem`, `*.crt`, `*.key` (`.gitignore:1-12`) but not `*.pfx`, `*.p12`, `*.cer`, `*.b64` or `*secret*.txt`.
- The untracked `image.png` in the repo root should be checked by the owner before any `git add -A` (not opened here).
- Add the patterns (S).

### L6. Platform logs not under our control. Low (UNKNOWN).
- Our code logs fixed fields only (VERIFIED `_shared/log.ts:13-18`, `whoami/handler.ts:46`, `analyze-message/handler.ts:172`).
- Supabase's own gateway logs keep request metadata. Confirm in the dashboard that the `Authorization` header is not retained (S).

### L7. Public-repo reconnaissance. Low.
- Architecture, group names, the MMSI checklist and a real third-party sender domain from a live bug (in the F-003 spec and tests) are public.
- Going private (ADR-0007 step 14) helps going forward. Existing clones keep the history.

### L8. Pseudonymisation is reversible. Low (already documented).
- `analyze_usage.user_key` and `auto_draft_users.mail_sha256` are unsalted SHA-256 hashes, reversible by guessing tenant oids or addresses (`security.md:58`).

### L9. Supabase Auth sign-ups. Low (UNKNOWN).
- The local `supabase/config.toml` has `enable_signup = true` (local only). If the remote project allows sign-ups, anyone can mint an `authenticated` JWT. All tables and RPCs are revoked from `authenticated` (VERIFIED), so they get no data.
- Disable sign-ups in the remote project anyway (S).

## 4. Controls checked and found sound (VERIFIED)
- **Token validation.**
  - RS256 only; `iss` exact tenant v2.0; `aud` is a string and equals the client ID; `ver` 2.0; `azp` equals the client ID, which rejects tokens minted for other clients; `scp` contains `access_as_user`, which rejects app-only tokens; `exp` required, 60 s skew.
  - Stateful functions also require `tid` to equal the configured tenant and a GUID `oid`, and key rows by the token's `oid`, never the body.
  - `_shared/entra.ts:37-56`, `_shared/stateful.ts:8-15`, `auto-reply-settings/handler.ts:69`.
- **CORS.** Exact allow-list that fails closed on `*`, `null` and unparsable entries (`_shared/entra.ts:13-27`). CORS is not an auth boundary: the boundary is the Entra redirect list plus the token. Adding the `dev` alias to `ALLOWED_ORIGINS` adds no theft path by itself. The redirect URI does (H2, H3).
- **No cross-user access.** The user is resolved from the token's `oid` (pane) or from our own subscription table (webhook). A user cannot make the backend read another mailbox (`_shared/processor.ts:74-87`, `ad_user_ctx`).
- **SSRF / token leakage.**
  - Backend Graph calls must start with `https://graph.microsoft.com/v1.0/` and match a method and path allow-list, including `nextLink` (`_shared/graph.ts:52-57,141-152`).
  - The pane's `nextLink` is host-checked (`src/taskpane/graph.ts:118-136`).
  - The OpenAI key goes only to `api.openai.com` (`_shared/ai.ts:149-151`).
- **Prompt-injection exfiltration.**
  - Summary, reply, translate and classify outputs, manual or automatic, are rejected if they contain any link, address, bare domain, IPv4 or `<`/`>`, including look-alike dots and zero-width tricks (`_shared/ai.ts:51-74`, `summary.ts:70`, `translate.ts:55`, `classified.ts:73`, `handler.ts:166-169`).
  - There is no markdown renderer, so no image-beacon channel.
  - Output is rendered as React text (no `innerHTML`/`dangerouslySetInnerHTML` in `src/`).
  - The reply form gets HTML-escaped text (`src/taskpane/messageView.ts:42-59`).
  - The model has no tools.
  - F-013 drafts allow only links already in the email (`_shared/ai.ts:131`) and go only into a reply to that sender.
  - The model sees one email per call, so it cannot leak a different email.
- **No send, move or delete in code** (`_shared/graph.ts:8-17`).
- **Logs.** Fixed field list; errors are swallowed without detail (`_shared/log.ts`, `_shared/db.ts:47-50`, `_shared/appToken.ts:48-50`, `src/taskpane/replyForm.ts:19-21`).
- **Secrets in git.** Full history scan (94 commits, all refs) for OpenAI/Supabase/GitHub/AWS/Slack token formats, JWTs, PEM private keys, base64 DER and `KEY=` assignments. The only hits are code literals (`_shared/db.ts:43` comment, `_shared/appToken.ts:13` PEM wrapper) and doc placeholders (F-013 plan line 39, F-015 plan line 20). No `.env`, `.pem`, `.pfx`, `.key`, `.cer`, mail export or HAR file was ever committed. No `.env` exists in the working tree.
- **Dependencies.** `jose` pinned (`npm:jose@6.2.13`) with lock files. `@azure/msal-browser` pinned exactly. CI uses `npm ci --ignore-scripts`, SHA-pinned actions, `contents: read`.

## 5. Cloud settings to confirm (I could not see these)
1. Entra, SPA app: which delegated grants are consented (admin **and** user consent); Assignment required; redirect URIs (is `localhost:3000` in prod?); whether B2B guests can get tokens.
2. Entra, Drafts app: is the cert uploaded; any Graph *application* permission consented (must be none); the owners of the app registration (an owner can add their own cert or secret, which is another theft path); the cert expiry date.
3. Exchange: role assignments and scope (`Get-ManagementRoleAssignment`); group type, join restriction and owners; mailbox auditing / `MailItemsAccessed` availability.
4. Tenant: is user sideloading of custom add-ins allowed; Conditional Access and MFA for the owner and admins.
5. GitHub org: 2FA required; members, outside collaborators and roles; Cloudflare GitHub App repository scope; Actions settings for fork PRs.
6. Cloudflare: account MFA and members; preview branch control; build command and `SKIP_DEPENDENCY_INSTALL`; whether fork PRs are built.
7. Supabase: account MFA, org members, plan and audit logs, remote Auth sign-ups, which secrets are set (names only), backups and retention.
8. OpenAI: org members and MFA, project-scoped key, budget limit, data controls (ZDR/EU per D5).
9. Owner's PC: BitLocker on, no other local admins beyond IT, EDR present.

## 6. Prioritised action list

### Code changes (small, each its own reviewed change)
1. **M5:** select `replyTo` and skip (new reason `reply_to_mismatch`) when it differs from `from`, in F-013. Show a visible note in the pane when `Reply-To` differs. S.
2. **L3:** bounded body read in `graph-webhook`. S.
3. **L4:** `_headers` with CSP and `Referrer-Policy`, tested live in Outlook first. S to M.
4. **L5:** extend `.gitignore` (`*.pfx`, `*.p12`, `*.cer`, `*.b64`, `*secret*.txt`). S.

### Configuration (owner / MMSI)
1. **Owner, today (H1):** confirm the Supabase secret works, then delete `drafts.key.pem` and `cron-secret.txt` from `C:\Users\thiago.santos\.dimscale-mailai-certs`, or move them to the company password manager. S.
2. **Owner, today (H2):** GitHub ruleset on `main` and `dev` (PR required, `build` check required, no force-push or deletion, restricted pushers). Org-wide 2FA. S.
3. **MMSI, before anything else in D14 (H3):** remove and revoke `Mail.ReadWrite`, `offline_access` and `User.Read` on the SPA app; Assignment required = Yes with the group. S.
4. **Owner (H4, M2):** MFA on the Supabase, Cloudflare, GitHub and OpenAI accounts, all owned by DIMSCALE mailboxes. Cloudflare previews for `dev` only, `SKIP_DEPENDENCY_INSTALL=1`, gated build command, GitHub App on this repo only. `supabase logout` when not deploying. S.
5. **MMSI (M1, M3):** closed-membership group for the RBAC scope, minimal owners, alert on membership change. Block user sideloading of custom add-ins. S.
6. **Owner (M3, L7):** finish the ADR-0007 cut-over, then make the repo private and remove the `dimscale.github.io` redirect and origin. M.
7. **Owner (M4):** OpenAI project key with a monthly budget cap. S.

### Monitoring and response

**How to notice a theft:**
- Entra sign-in logs for the **Drafts service principal** from unexpected IPs or at unexpected times.
- Purview audit `MailItemsAccessed` by the Drafts app on pilot mailboxes.
- Membership changes on the RBAC group.
- New credentials (certificate or secret) added to either app registration (audit "Update application - Certificates and secrets management").
- Unknown Graph subscriptions.
- Pushes and deploys you did not make (GitHub and Cloudflare notifications).
- New Supabase function versions you did not deploy.
- OpenAI usage spikes.
- In our own logs: bursts of `client_state_mismatch`, `tenant_mismatch` or `unknown_subscription`, and 401 reason spikes.

**What to revoke first** (in this order, by suspected source):
1. **Cert suspected:**
   1. MMSI disables the Drafts enterprise app.
   2. MMSI deletes the cert.
   3. MMSI removes both RBAC role assignments.
   4. Owner sets `AUTO_DRAFTS_ENABLED=false`.
   5. Generate a new cert later.
2. **Pane code suspected:**
   1. Roll back the Cloudflare deployment / revert `main` (and `dev`).
   2. MMSI revokes the SPA's Graph consent (stops new Graph tokens).
   3. MMSI revokes sign-in sessions of affected users.
   4. Set `AUTO_MODE_AVAILABLE=false` only after the revert is trusted.
3. **Supabase account suspected:**
   1. Owner rotates the password and MFA and revokes CLI tokens.
   2. Treat the cert as stolen (step 1 above).
   3. Rotate the OpenAI key, Supabase secret keys, `GRAPH_CLIENT_STATE` (recreate subscriptions) and `CRON_SHARED_SECRET` (secret and Vault).
   4. Redeploy functions from a reviewed commit.
4. **OpenAI key only:** revoke in OpenAI and set a new secret. No mail exposure.

## Not checked
- No live calls to Graph, Entra, Supabase or OpenAI.
- `deno test` not run (Deno not installed); `npm test` not run (no code changed).
- Cloudflare's dependency-install behaviour and Graph `createReply` `Reply-To` handling are stated from memory and marked "verify".
- The certificate file was not opened. "Unencrypted" rests on the plan's `-nodes` command, not on reading the file.
- Line numbers refer to commit `9eaa7a1`. During this audit another session left uncommitted edits:
  - `_shared/log.ts` adds the fields `action` (enum) and `rt` (reasoning-token count);
  - `_shared/ai.ts` adds an optional `service_tier`;
  - preflight `Access-Control-Max-Age` was added in `analyze-message` and `auto-reply-settings`;
  - pane edits in `AiTools.tsx` and `autoMode.ts`.

  I reviewed the backend diffs. They add no content to logs and change no finding here, but shift some `analyze-message/handler.ts` line numbers by about 5. The pane diffs were not reviewed.
