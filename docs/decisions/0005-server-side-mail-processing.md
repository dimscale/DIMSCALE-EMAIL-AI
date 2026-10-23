# ADR-0005: Server-side mail processing (automatic reply drafts)

- **Status:** Accepted (2026-10-23) by the user, with the approval of F-013. The tenant-admin setup (MMSI) and the DPO items are still pending.
- **Date:** 2026-10-22, amended 2026-10-23
- **Feature:** [F-013 automatic reply drafts](../specs/F-013-automatic-reply-drafts.md)

## Context
New user requirement (2026-10-22): "When an email arrives, even with Outlook closed, the system runs, creates the reply emails and leaves them in Drafts." The user still reviews and sends every reply by hand. Nothing is ever sent automatically.

**Amendment (user direction 2026-10-23):** the trigger is a per-user **contact allow-list**. In the pane's "Responder" tab the user ticks the contacts the system answers; only mail from those senders gets a draft. The user still requires it to work **with Outlook closed**, so the server-side architecture below is unchanged; the allow-list only narrows what it reads and sends to OpenAI.

This is the first feature that runs **without the user and without Outlook**. It changes three things that the current architecture rules out:
- The backend is no longer stateless. It must remember subscriptions and which messages it already handled.
- The backend calls Microsoft Graph itself. Until now Graph was called only from the task pane with a delegated token, and [security.md](../architecture/security.md) says "No application (app-only) permissions".
- Email content reaches OpenAI without a click from the user, including mail from third parties.

CLAUDE.md rule 8 requires an ADR for this.

### Verified platform facts
| Fact | Source |
|---|---|
| An Outlook add-in can't do this. Event-based activation has no "message received" event, its events are compose and send oriented, handlers time out after about 300 s, and the add-in runs only while Outlook is open. | https://learn.microsoft.com/office/dev/add-ins/develop/event-based-activation |
| Graph change notifications support Outlook `message` with delegated **and application** permissions. With delegated permissions you can subscribe only to the signed-in user's mailbox. | https://learn.microsoft.com/graph/outlook-change-notifications-overview |
| Message subscriptions need `Mail.Read` (application). "Microsoft Graph subscriptions don't support write access permissions when only read access permissions are needed." | https://learn.microsoft.com/graph/api/subscription-post-subscriptions |
| Maximum lifetime for an Outlook message subscription is 10,080 minutes (under 7 days), or 1,440 minutes with resource data. Message notification latency averages under 1 minute, with a maximum of 3 minutes. | https://learn.microsoft.com/graph/api/resources/subscription |
| The webhook must answer with a 2xx within 3 seconds. Otherwise Graph retries for up to 4 hours. Slow endpoints are throttled, and notifications can be dropped "in a way that they can't be recovered". | https://learn.microsoft.com/graph/change-notifications-delivery-webhooks |
| Lifecycle notifications: `reauthorizationRequired`, `subscriptionRemoved`, `missed`. After a removal or a miss, the app must resync, for example with a delta query. | https://learn.microsoft.com/graph/change-notifications-lifecycle-events |
| `createReply` creates a draft and needs `Mail.ReadWrite` (delegated or application). Sending is a separate call. `reply` (which sends) needs `Mail.Send`. | https://learn.microsoft.com/graph/api/message-createreply, https://learn.microsoft.com/graph/api/message-reply |
| Application `Mail.ReadWrite` is "Read and write mail in all mailboxes … Does not include permission to send mail". It needs admin consent. | https://learn.microsoft.com/graph/permissions-reference |
| **RBAC for Applications** in Exchange Online grants an app a role (for example `Application Mail.ReadWrite`) limited to a resource scope (a management scope or an administrative unit). It "replaces Application Access Policies". Grants are a **union** with Entra consents, so the unscoped Entra grant must be removed or never given. Changes take 30 minutes to 2 hours to apply because of caching. A `MemberOfGroup` filter counts only direct members. | https://learn.microsoft.com/exchange/permissions-exo/application-rbac |
| Client credentials flow: a shared secret, a certificate (`private_key_jwt`) or a federated credential. A certificate or a federated credential gives "a higher level of assurance". No refresh tokens are issued. | https://learn.microsoft.com/entra/identity-platform/v2-oauth2-client-creds-grant-flow, https://learn.microsoft.com/entra/identity-platform/certificate-credentials |
| Workload identity federation needs a token from an external OIDC identity provider (GitHub Actions, Kubernetes, Google, AWS, SPIFFE, and others). Supabase Edge Functions are not listed, and Supabase documents no workload identity token for functions. | https://learn.microsoft.com/entra/workload-id/workload-identity-federation |
| OBO needs a confidential client (secret or certificate), and only `offline_access` returns a refresh token. Refresh tokens last 90 days by default (24 hours for SPA redirects) and can be revoked by admin and user actions. | https://learn.microsoft.com/entra/identity-platform/v2-oauth2-on-behalf-of-flow, https://learn.microsoft.com/entra/identity-platform/refresh-tokens |
| Supabase Edge Functions: wall clock 150 s on Free and 400 s on paid plans, 2 s of CPU per request, `EdgeRuntime.waitUntil` for work after the response. Supabase Cron (`pg_cron` plus `pg_net`) can call functions on a schedule. Vault stores encrypted secrets. RLS with no policies blocks Data API access, and secret keys bypass RLS. | https://supabase.com/docs/guides/functions/limits, https://supabase.com/docs/guides/functions/background-tasks, https://supabase.com/docs/guides/functions/schedule-functions, https://supabase.com/docs/guides/database/vault, https://supabase.com/docs/guides/database/postgres/row-level-security |
| Power Automate's Office 365 Outlook connector: the "When a new email arrives (V3)" trigger polls (the webhook trigger is deprecated), "there could be rare circumstances when the trigger's delay to fire may take up to one hour", and "some emails could be missed". Premium connectors (such as HTTP) need a Power Automate Premium or Process license. The Free license covers standard connectors only. | https://learn.microsoft.com/connectors/office365/, https://learn.microsoft.com/power-platform/admin/power-automate-licensing/types |

## Options

| Option | Summary | Pros | Cons | Verdict |
|---|---|---|---|---|
| **A. Add-in only** | Event-based activation | No backend change | Has no "message received" event and doesn't run with Outlook closed. Can't meet the requirement. | **Rejected** |
| **B. Graph webhooks + Supabase, app-only (application permissions, scoped by Exchange RBAC for Applications)** | A separate confidential app registration with a certificate credential. Exchange roles `Application Mail.Read` and `Application Mail.ReadWrite`, scoped to a DIMSCALE mail-enabled group. Supabase Edge Functions receive the notifications, and Supabase Cron renews subscriptions. | Works with Outlook closed and without any user sign-in. No per-user token storage. Access is limited by Exchange to the group's mailboxes. No `Mail.Send`. No new vendor (ADR-0003). Microsoft documents RBAC for Applications as the replacement for Application Access Policies. | One credential can read and write mail in **every mailbox in scope**, so a leak of the certificate private key is serious. Admin work in Exchange PowerShell. RBAC changes take 30 to 2 hours to apply. A public webhook endpoint. | **Recommended** |
| **C. Graph webhooks + Supabase, delegated `Mail.ReadWrite` + `offline_access`, refresh tokens stored per user** | The task pane sends its NAA token for `access_as_user` to a function. The function runs OBO to get Graph tokens plus a refresh token and stores it encrypted (Vault or AES with a key in secrets). | Access follows each user's own consent and mailbox. No tenant-wide app permission. | The SPA registration (a public client today) becomes confidential, or a second app with a credential is needed anyway, so there's still a secret. Each user must open the add-in once to bootstrap, and again after the token is revoked or expires (90 days, password reset by an admin, revocation). When that happens drafts stop silently. Needs a table of long-lived, high-value refresh tokens (bound to the user and client, usable against any resource the client has consent for). More code: OBO, refresh rotation, encryption, re-consent UX. It contradicts the pending admin request to revoke the old delegated `Mail.ReadWrite`/`offline_access` grants. | Rejected (higher risk and fragility for the same blast radius per user) |
| **D. Power Automate flow per user** | "When a new email arrives (V3)" → HTTP to OpenAI or to our function → create a draft | Microsoft-native. Microsoft manages the connection tokens. | A polling trigger with delays of up to 1 h and documented missed emails. HTTP is a premium connector (Premium or Process license per user or flow). The OpenAI key, or a credential for our function, would live in every flow. One flow per user to maintain outside source control, with no tests and no review (breaks the SDD lifecycle). The connector lists "Draft an email message" and "Reply to email (V3)", and the docs don't say whether Reply sends at once, so the threading of a reply draft is unverified. | Rejected for now. It's an option only if DIMSCALE already has Premium licenses and accepts no code review. |
| E. n8n, VPS, Azure Functions with Event Hubs | Other hosts | — | n8n and a VPS are excluded by CLAUDE.md rule 8 and ADR-0003 option C. Azure needs a subscription (ADR-0004: an Azure account required a payment card). | Excluded |

### Credential for option B
| Credential | Verdict |
|---|---|
| **Certificate (recommended).** We generate the key pair locally. Only the public `.cer` goes to the admin. The private key (base64 PKCS#8) is stored only as a Supabase function secret. Assertions are signed PS256 with `jose` (already pinned for F-002). | The private key never passes through the admin or email. This is the higher-assurance option per Microsoft. |
| Client secret | Simpler code, but the admin creates it and has to send it to us (a transfer risk). Fallback only, if the admin can't upload a certificate. |
| Federated identity credential | Not available: Supabase issues no workload OIDC token for functions (see the facts table). |

## Decision (proposed)
**Option B**, with the following parts:
1. **A new, separate app registration** "DIMSCALE Mail AI Drafts" (single tenant, no redirect URI, certificate credential). Keeping it separate from the add-in's SPA registration means the public SPA client never gets a credential or application permissions, and either app can be disabled on its own.
2. **No Microsoft Graph application permission is consented in Entra.** The permissions come only from Exchange **RBAC for Applications**: `Application Mail.Read` (needed for subscriptions) and `Application Mail.ReadWrite` (for `createReply`), both with a custom resource scope `MemberOfGroup -eq '<DN of a mail-enabled DIMSCALE group>'`. `Mail.Send` and any "Full Access" role are never assigned. Since `/reply`, `/send` and `sendMail` need `Mail.Send`, the platform itself blocks sending as a second line of defense.
3. **Two Supabase Edge Functions** (`verify_jwt = false`, each authenticates its caller itself):
   - `graph-webhook`, called by Graph. It handles the validation handshake, checks `clientState`, stores a minimal queue row, answers `202` within 3 s, and starts processing with `EdgeRuntime.waitUntil`.
   - `auto-draft-cron`, called by Supabase Cron with a shared secret. It retries queued work, renews subscriptions, catches up after lifecycle events, and deletes old metadata.
4. **Supabase Postgres (eu-west-1)** stores metadata only: enabled users, subscriptions, and processed message IDs with their status and reason code. RLS is enabled with no policies. No subject, body, address or AI output is stored.
5. **OpenAI** gets only what [F-013](../specs/F-013-automatic-reply-drafts.md) lists, through the Responses API with `store: false`, no tools and a strict JSON schema. The output becomes a reply draft through `createReply`, with a visible "AI draft" marker. The user sends it, or doesn't.
6. **Contact allow-list (amendment 2026-10-23).**
   - The pane lists candidate contacts from the user's recent Inbox senders with **delegated `Mail.ReadBasic`** (the Graph token stays in the pane) plus manual entry, and saves the chosen addresses and an on/off switch through a third function, `auto-reply-settings`, authenticated with the F-002 token validation. The client never writes the database directly.
   - Postgres gains `auto_reply_contacts` (lower-cased addresses keyed by the user's Entra `oid`): third-party personal data, deletable by the user, with a retention rule (DPO).
   - For every notified message the processor first reads only `from`; a sender not on the list ends the row with no further read and no OpenAI call. Subscriptions exist only for enabled users with at least one contact.
   - This does not change the permission model: the app-only roles still allow reading any message of an in-scope mailbox, so the allow-list is an application control, not a security boundary. The RBAC group remains the boundary.
7. **Kill switches**, from fastest to slowest:
   - the `AUTO_DRAFTS_ENABLED` function secret (read on every call, no redeploy needed);
   - per-user switch off in the Responder tab (or `enabled = false`);
   - removing the certificate from the app registration (no new tokens);
   - removing the Exchange role assignments (takes effect in 30 minutes to 2 hours).

## Consequences
- **Docs to update when this ADR is accepted** (the F-013 plan, step 1):
  - `security.md`: today it says "No application (app-only) permissions", "Stores no email content" and that the backend "never receives Graph tokens". New sections are needed for app-only Graph, webhook authentication, stored metadata and the third-party (sender) data flow.
  - `overview.md`: the diagram, a capability boundary row for "Mailbox (server, app-only)", and "Stateless".
  - `requirements.md`: "Automatic rules or filters" is listed as out of scope. Add R8 "Automatic reply drafts", if the user agrees.
- A new secret class appears: the certificate private key, the `clientState`, the cron shared secret and the OpenAI key, all in Supabase secrets or Vault. The private key needs rotation before the certificate expires.
- The tenant admin has to run Exchange Online PowerShell (Organization Management) and create or reuse a **mail-enabled** group with **direct** members.
- The pending admin request to **revoke the old delegated `Mail.ReadWrite`, `offline_access` and `User.Read` grants** on the add-in's SPA registration is still correct and still needed. Option B doesn't use delegated grants, and leaving them would widen the SPA's reach.
- Email content of incoming mail from the **senders the user chose**, including external ones, is processed by OpenAI without a per-message user action. This **needs the DPO's decision (D5)** before going live. The allow-list makes the flow narrower and easier to justify than the 2026-10-22 "every Inbox message" design.
- Supabase stores each user's chosen contact addresses (personal data of third parties) and the delegated `Mail.ReadBasic` scope is added to the add-in app (admin consent is required in this tenant because Assignment required is on).
- The public webhook URL can be called by anyone. Requests without a valid `clientState` are dropped. Anything they could cost is limited to Supabase invocations, since they never reach Graph or OpenAI.
- Graph may mark a slow endpoint "slow" or "drop" and lose notifications. Mitigations: the fast ack, catch-up on `missed`, and a periodic catch-up window. A switch to Event Hubs would need Azure (excluded for now).

## Open
- User: accept option B with the contact allow-list, or ask for C or D.
- DPO: D5 for automated processing of mail from chosen contacts by OpenAI; storage and retention of the contact list; the DPAs with OpenAI and Supabase.
- Tenant admin (MMSI): the checklist in [docs/admin/2026-10-23-mmsi-request.md](../admin/2026-10-23-mmsi-request.md): RBAC for Applications without any Entra application consent, the mail-enabled group, the certificate, and delegated `Mail.ReadBasic` consent on the add-in app.
- To verify in the F-013 spike: whether a subscription and `createReply` work with RBAC-only grants (Microsoft documents Graph as a supported protocol, but doesn't say anything specific about subscriptions).
