# F-001: Task pane shell and active message context

- **Status:** Done (2026-10-22)
- **Tier:** Active item only. No Graph, no AI, no backend.
- **Depends on:** ADR-0001 approved.
- **Plan:** [docs/plans/F-001-active-message-context.md](../plans/F-001-active-message-context.md)

## Goal
Scaffold the add-in and show the context of the open message in the task pane. This is the base every AI feature builds on, and the step that proves the dev and sideload loop works.

## Behavior
When the user opens the add-in on a message in read mode, the panel shows:
- the subject
- the sender's display name and email address
- the item date and time (local format), from `item.dateTimeCreated`. Office.js `MessageRead` has no received-date property, so the UI labels it "Date". It is verified only for received mail in a live test.
- the plain-text body, truncated to the first 2,000 characters, with a "(truncated)" marker when cut
- the body length in characters

## Acceptance criteria
| ID | Criterion | Verify by |
|---|---|---|
| F-001/AC-1 | The project is scaffolded per ADR-0001. `npm run build:dev` succeeds and the manifest passes `npm run validate`. The production `npm run build` fails on purpose until the hosting URL is set (status D7). | Command output |
| F-001/AC-2 | The manifest requests only `ReadItem` permission and activates only on the message read surface. | Manifest inspection |
| F-001/AC-3 | The panel shows the subject, sender name and address, and the item date (`dateTimeCreated`) from `Office.context.mailbox.item`. | Unit test (mocked Office.js) and manual sideload |
| F-001/AC-4 | The body is read with `item.body.getAsync(Office.CoercionType.Text)`. Text longer than 2,000 characters is cut with a "(truncated)" marker, and the full length is shown. | Unit test |
| F-001/AC-5 | An empty body shows "This message has no text content". An empty subject shows "(no subject)". | Unit test |
| F-001/AC-6 | If `body.getAsync` fails, the panel shows a user-readable error and a Retry button, and logs no email content. | Unit test |
| F-001/AC-7 | If there is no item (for example a read surface with no selection), the panel shows "Open a message to use DIMSCALE Mail AI". | Unit test |
| F-001/AC-8 | All email text is rendered as text nodes (no `dangerouslySetInnerHTML` or `innerHTML`). An HTML or script string in the body appears literally. | Unit test and review |
| F-001/AC-9 | No network requests other than those made by Office.js and the add-in's own assets. No `console.*` call outputs subject, body or addresses. | Review and grep |
| F-001/AC-10 | Typecheck, lint, unit tests and build pass. | Command output |

## Edge cases
- A very large body (1 MB or more): truncated, the UI stays responsive.
- A body with HTML, scripts or emoji: shown literally and safely.
- A sender with no display name: show the address only.
- Office.js not ready yet (`Office.onReady` not resolved): render nothing until it is ready.

## Out of scope
AI calls, authentication or MSAL, Graph, compose mode, pinning the pane or `ItemChanged` handling, styling beyond basic readable defaults (Fluent UI only if the generator includes it), localization.

## References
- https://learn.microsoft.com/office/dev/add-ins/quickstarts/outlook-quickstart-yo
- https://learn.microsoft.com/javascript/api/outlook/office.messageread (body: Mailbox 1.1; `dateTimeCreated`: "Gets the date and time that an item was created"; no received-date property exists)

## Open questions
- Should the UI language be Portuguese or English? (Default: English, with strings kept in one place.)
