# Product Requirements: DIMSCALE Mail AI

## Goal
Help DIMSCALE users understand, triage and answer email faster from inside Outlook, while keeping the user in control of every email action.

## In scope (initial product)
| ID | Capability | Tier |
|---|---|---|
| R1 | Outlook task pane in message read mode | Active item |
| R2 | Summarize the open message | Active item + AI |
| R3 | Suggest a professional reply | Active item + AI |
| R4 | Refine a draft ("shorter", "more diplomatic", "more formal", free text) | Active item + AI |
| R5 | Classify the message as important, normal, advertising or potentially suspicious, with an explanation and a "may be wrong" notice | Active item + AI |
| R6 | Put the suggested reply into a reply the user reviews and sends themselves | Active item (Office.js reply form) or Mailbox (Graph draft) |
| R7 | View and select other messages, beyond the open one | Mailbox (Graph) |
| R8 | Automatic reply drafts for contacts the user selects: when one of them writes, a reply draft appears in Drafts, even with Outlook closed. Drafts only, the user reviews and sends ([F-013](../specs/F-013-automatic-reply-drafts.md)) | Mailbox (server, app-only) + AI + pane |

Tiers are defined in [architecture overview](../architecture/overview.md#capability-boundaries).

## Hard constraints
- Never send, delete, move or report email automatically. Automatic drafts (R8) are the only automatic action: a draft is created, never sent, and only for senders the user chose.
- The AI output is advisory. Classification is never presented as certain.
- The OpenAI API key stays server-side.
- Send the AI as little email content as possible.
- Email content can never instruct the system.

## Out of scope (for now)
Automatic rules or filters other than R8's contact allow-list (no rule that sends, moves, deletes, flags or categorizes mail, no all-senders mode), bulk actions (F-015 is a separate, manual feature), attachment analysis, calendar, mobile-specific UX, multi-language UI, analytics or telemetry on email content, and Marketplace (AppSource) publication.

## Users
DIMSCALE staff on Microsoft 365 work accounts (tenant to be confirmed, see [status](../status.md#open-decisions)).

## Success signals (qualitative, to refine)
- A summary is available in a few seconds after the user clicks.
- The user edits a suggested reply less than they would write one from scratch.
- No incident of mail sent, deleted or leaked by the add-in.
