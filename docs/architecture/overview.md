# Architecture Overview

## Assessment at project start (2026-10-22)
_Historical. For the current state see [status](../status.md)._
- Repo `dimscale/DIMSCALE-EMAIL-AI`, branch `main`, 1 commit ("Initial commit") containing only an empty `README.md`.
- **Nothing exists yet:** no framework, package manager, manifest, auth, Graph integration or backend. There is no working code to preserve.
- Local toolchain: Node 24.22.0, npm 11.19.1. pnpm and the GitHub CLI are not installed.
- Note: the parent folder `C:\Users\thiago.santos` is itself a git repo (branch `master`). This project has its own `.git`, so it is unaffected, but don't run git commands from the home folder.

## Target architecture

```
 GitHub Actions build ─► GitHub Pages (https://dimscale.github.io/DIMSCALE-EMAIL-AI/)
                              │ static files + public production manifest
                              ▼
 Outlook (web / new Windows / classic Windows / Mac)
 └─ Task pane (React + TypeScript, Office.js)            ← served over HTTPS (localhost:3000 in dev)
      ├─ Office.js: read the active item, open the reply form   (no Graph, ReadItem)
      ├─ MSAL.js NAA ─► Microsoft Graph, delegated `Mail.ReadBasic` only (F-013 candidate senders; token stays in the pane)
      └─ MSAL.js NAA ─► Entra token for api://<client-id>/access_as_user
           └─ fetch + Bearer token ─► Supabase Edge Functions (verify_jwt = false, each validates its caller)
                                     ├─ F-002 `whoami`: validates the token itself (JWKS, issuer, audience, azp, scp, expiry)
                                     ├─ F-013 `auto-reply-settings`: same validation + oid/tid; stores the user's switch and chosen contacts
                                     └─ later (F-003+): minimizes and delimits the email text for OpenAI

 Outlook closed, server side (F-013, app-only):
 Exchange mailbox ─ Graph change notification (new Inbox message) ─► `graph-webhook` (clientState check, 202, queue row)
 `auto-draft-cron` (pg_cron, x-cron-secret) ─► subscriptions, catch-up, retention
 worker ─► Entra (client credentials, certificate) ─► Graph (read `from`; only for allow-listed senders the minimal body; `createReply` draft)
        ─► OpenAI API (key in Supabase secrets, `store: false`) ─► draft in the user's Drafts folder (never sent)
        ─► Supabase Postgres (RLS locked: contacts, subscriptions, queue metadata)
```

Sign-in is limited to one DIMSCALE Entra group (enterprise app, Assignment required). Details: [security](security.md#authentication-and-authorization).

Production hosting ([F-011](../specs/F-011-github-pages-hosting.md)): the tenant admin uploads the production manifest once in the Microsoft 365 admin center, and later code changes reach users through a normal deploy. Entra has the SPA redirects `brk-multihub://localhost:3000` and `brk-multihub://dimscale.github.io`, and the function's `ALLOWED_ORIGINS` is `https://localhost:3000,https://dimscale.github.io`. The Pages origin is shared org-wide, see [security](security.md#hosting-github-pages-adr-0004-f-011).

Decisions: [ADR-0001 stack](../decisions/0001-addin-stack.md), [ADR-0002 auth](../decisions/0002-auth-nested-app-auth.md), [ADR-0003 AI backend](../decisions/0003-ai-backend.md), [ADR-0004 task pane hosting](../decisions/0004-taskpane-hosting.md) (GitHub Pages), [ADR-0005 server-side mail processing](../decisions/0005-server-side-mail-processing.md) (app-only, F-013). All five are Accepted. Features: [F-002](../specs/F-002-authenticated-backend.md), [F-013](../specs/F-013-automatic-reply-drafts.md).

## Capability boundaries

| Tier | What | Mechanism | Permission |
|---|---|---|---|
| **Active item** | Read the open message (subject, sender, body) | Office.js `Office.context.mailbox.item` (`body.getAsync`, Mailbox 1.1) | Manifest `ReadItem` |
| **Active item** | Open a prefilled reply for the user to edit and send | `item.displayReplyFormAsync` (Mailbox 1.9), which opens a form and never sends | Manifest `ReadItem` |
| **Mailbox (pane)** | List recent senders for the contact picker (`from` only, no body) | Graph `GET /me/mailFolders/inbox/messages?$select=from…` | Delegated `Mail.ReadBasic` (admin consent, F-013) |
| **Mailbox (pane)** | List and select other messages (F-015) | Graph `GET /me/messages` | Delegated, least privileged for the field set (to be decided in F-015) |
| **Mailbox (server, app-only)** | Read new Inbox mail of allow-listed senders and create a reply draft in Drafts with Outlook closed (F-013) | Graph change notifications, `GET` message, `POST /users/{id}/messages/{id}/createReply` | Exchange RBAC for Applications `Application Mail.Read` + `Application Mail.ReadWrite`, scoped to a mail-enabled group; no Entra application consent |
| **Never** | Send, delete, move, report | none | `Mail.Send` is never requested or assigned |

Active-item features (R1 to R6) need **no Graph permission**. Graph arrives with F-013: delegated `Mail.ReadBasic` in the pane, and the one app-only exception above on the server (a decision recorded in ADR-0005).

Sources:
- NAA: https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in
- MessageRead (body, displayReplyFormAsync): https://learn.microsoft.com/javascript/api/outlook/office.messageread
- createReply: https://learn.microsoft.com/graph/api/message-createreply (creates a draft only, sending is a separate call)
- Manifests: https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests

## Boundaries and responsibilities
- **Task pane**: UI and Office.js calls. Holds no secrets. Renders all email and AI text as plain text.
- **AI backend** (Supabase Edge Functions): every function authenticates its caller itself, since `verify_jwt = false` (Entra token for `whoami` and `auto-reply-settings`, `clientState` for `graph-webhook`, `x-cron-secret` for `auto-draft-cron`). The only component that holds the OpenAI key and the Drafts app's certificate. **Not stateless since F-013**: Postgres (RLS locked) keeps the user's switch, chosen contact addresses, Graph subscriptions and a message-ID queue with reason codes. It stores no subject, body, display name, AI output or token. Logs contain metadata only.
- **Microsoft Graph**: the pane calls it with a delegated token (`Mail.ReadBasic`); the backend never receives that token. The backend calls Graph app-only, within the RBAC scope, for F-013.
