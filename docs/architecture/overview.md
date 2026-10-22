# Architecture Overview

## Assessment at project start (2026-10-22)
_Historical. For the current state see [status](../status.md)._
- Repo `dimscale/DIMSCALE-EMAIL-AI`, branch `main`, 1 commit ("Initial commit") containing only an empty `README.md`.
- **Nothing exists yet:** no framework, package manager, manifest, auth, Graph integration or backend. There is no working code to preserve.
- Local toolchain: Node 24.22.0, npm 11.19.1. pnpm and the GitHub CLI are not installed.
- Note: the parent folder `C:\Users\thiago.santos` is itself a git repo (branch `master`). This project has its own `.git`, so it is unaffected, but don't run git commands from the home folder.

## Target architecture (proposed)

```
 Outlook (web / new Windows / classic Windows / Mac)
 └─ Task pane (React + TypeScript, Office.js)            ← served over HTTPS
      ├─ Office.js: read the active item, open the reply form   (no Graph, ReadItem)
      ├─ MSAL.js with nested app auth (NAA) ─► Microsoft Graph  (only for mailbox-tier features)
      └─ fetch + Entra ID token ─► AI backend (proposed: Supabase Edge Function)
                                     ├─ validates the token (issuer, audience, expiry)
                                     ├─ minimizes and delimits the email text
                                     └─ OpenAI API (key held in the backend secret store)
```

Decisions: [ADR-0001 stack](../decisions/0001-addin-stack.md), [ADR-0002 auth](../decisions/0002-auth-nested-app-auth.md), [ADR-0003 AI backend](../decisions/0003-ai-backend.md). ADR-0001 and ADR-0003 are Accepted. ADR-0002 is Proposed.

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
- **AI backend**: the only component holding the OpenAI key. Stateless. Stores no email content. Its logs contain metadata only (request id, latency, status).
- **Microsoft Graph**: called directly from the task pane with a delegated token. The backend never receives Graph tokens.
