# Architecture Overview

## Assessment at project start (2026-10-22)
_Historical. For the current state see [status](../status.md)._
- Repo `dimscale/DIMSCALE-EMAIL-AI`, branch `main`, 1 commit ("Initial commit") containing only an empty `README.md`.
- **Nothing exists yet:** no framework, package manager, manifest, auth, Graph integration or backend. There is no working code to preserve.
- Local toolchain: Node 24.22.0, npm 11.19.1. pnpm and the GitHub CLI are not installed.
- Note: the parent folder `C:\Users\thiago.santos` is itself a git repo (branch `master`). This project has its own `.git`, so it is unaffected, but don't run git commands from the home folder.

## Target architecture

```
 Outlook (web / new Windows / classic Windows / Mac)
 └─ Task pane (React + TypeScript, Office.js)            ← served over HTTPS
      ├─ Office.js: read the active item, open the reply form   (no Graph, ReadItem)
      ├─ MSAL.js with nested app auth (NAA) ─► Microsoft Graph  (only for mailbox-tier features)
      └─ MSAL.js NAA ─► Entra token for api://<client-id>/access_as_user
           └─ fetch + Bearer token ─► Supabase Edge Function (verify_jwt = false)
                                     ├─ F-002 (`whoami`): validates the token itself (JWKS, issuer, audience, azp, scp, expiry), returns the user's name
                                     ├─ later (F-003+): minimizes and delimits the email text
                                     └─ later: OpenAI API (key held in the Supabase secret store)
```

Sign-in is limited to one DIMSCALE Entra group (enterprise app, Assignment required). Details: [security](security.md#authentication-and-authorization).

Decisions: [ADR-0001 stack](../decisions/0001-addin-stack.md), [ADR-0002 auth](../decisions/0002-auth-nested-app-auth.md), [ADR-0003 AI backend](../decisions/0003-ai-backend.md), [ADR-0004 task pane hosting](../decisions/0004-taskpane-hosting.md) (GitHub Pages). All three are Accepted (ADR-0003: Supabase Edge Functions). Feature: [F-002](../specs/F-002-authenticated-backend.md).

## Capability boundaries

| Tier | What | Mechanism | Permission |
|---|---|---|---|
| **Active item** | Read the open message (subject, sender, body) | Office.js `Office.context.mailbox.item` (`body.getAsync`, Mailbox 1.1) | Manifest `ReadItem` |
| **Active item** | Open a prefilled reply for the user to edit and send | `item.displayReplyFormAsync` (Mailbox 1.9), which opens a form and never sends | Manifest `ReadItem` |
| **Mailbox** | List and select other messages | Graph `GET /me/messages` | Delegated `Mail.Read` (incremental consent) |
| **Mailbox** | Create or update a reply draft in Drafts | Graph `POST /me/messages/{id}/createReply`, then `PATCH` the draft | Delegated `Mail.ReadWrite` |
| **Never** | Send, delete, move, report | none | `Mail.Send` is never requested |

The MVP (R1 to R6) fits entirely in the **active item** tier, so the first releases need **no Graph permission and no admin consent**. Graph arrives with R7 and optional server-side drafts.

Sources:
- NAA: https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in
- MessageRead (body, displayReplyFormAsync): https://learn.microsoft.com/javascript/api/outlook/office.messageread
- createReply: https://learn.microsoft.com/graph/api/message-createreply (creates a draft only, sending is a separate call)
- Manifests: https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests

## Boundaries and responsibilities
- **Task pane**: UI and Office.js calls. Holds no secrets. Renders all email and AI text as plain text.
- **AI backend** (Supabase Edge Function): authenticates every request itself, since `verify_jwt = false`. The only component that will hold the OpenAI key (not used in F-002). Stateless. Stores no email content, and F-002 receives none. Its logs contain metadata only (request id, status, reason code, duration).
- **Microsoft Graph**: called directly from the task pane with a delegated token. The backend never receives Graph tokens.
