# Project Status

_Last updated: 2026-10-22_

## Current state
- Milestone 0 (SDD setup) is done: the rules, agents, architecture, security, ADRs, test strategy and the F-001 spec and plan are written.
- F-001 is implemented, QA'd and reviewed (31 tests, all checks green). The evidence is in the [plan](plans/F-001-active-message-context.md#evidence).

## Next step
The user does the live sideload of F-001 (plan step 7). Then close F-001 and decide D3 and D4 to unblock F-002.

## Backlog (prioritized)
| ID | Feature | Tier | Depends on | Status |
|---|---|---|---|---|
| F-001 | Task pane shell and active message context | Active item | ADR-0001 | Awaiting live sideload |
| F-002 | Authenticated AI backend skeleton (token validation, a health operation, no OpenAI yet) | Backend | ADR-0002, ADR-0003 | Not started |
| F-003 | Summarize the open message | Active item + AI | F-001, F-002 | Not started |
| F-004 | Suggest a reply (shown in the panel, copyable) | Active item + AI | F-003 | Not started |
| F-005 | Open a prefilled reply form with the suggestion (`displayReplyFormAsync`; the user sends) | Active item | F-004 | Not started |
| F-006 | Refine the suggestion (shorter, diplomatic, formal, custom) | Active item + AI | F-004 | Not started |
| F-007 | Classification (important, normal, advertising, potentially suspicious) with reasons and an uncertainty notice | Active item + AI | F-003 | Not started |
| F-008 | NAA sign-in and Graph `Mail.Read`: list and select recent messages | Mailbox | ADR-0002 | Not started |
| F-009 | Graph reply draft (`createReply` plus `PATCH`, needs `Mail.ReadWrite`) as an alternative to F-005 | Mailbox | F-008 | Not started |
| F-010 | Compose-mode support (refine the text being written) | Active item + AI | F-006 | Needs research |

## Open decisions
| # | Decision | Owner | Blocks |
|---|---|---|---|
| D1 | ~~Approve ADR-0001~~ (accepted 2026-10-22) | User | — |
| D2 | Which Outlook clients DIMSCALE uses | User | Manifest and testing |
| D3 | Backend vendor: Supabase (recommended) or Azure Functions (ADR-0003) | User | F-002 |
| D4 | Entra tenant and who creates the app registration and grants consent | User / IT | F-002, F-008 |
| D5 | The OpenAI account, model, and GDPR acceptance of sending email content to OpenAI | User / DPO | F-003 |
| D6 | UI language (PT or EN). EN is used for now, with strings in `src/taskpane/strings.ts` | User | — |
| D7 | Production hosting URL (replaces the `contoso.com` placeholder; the production build fails until it is set) | User | Any deployment |
| D8 | Keep the IE11 polyfills (`core-js`, `es6-promise`, browserslist `ie 11`)? Depends on D2 (classic Outlook webview) | User | Bundle cleanup |
| D9 | `package.json` license: MIT (generator default) or proprietary/UNLICENSED | User | — |

## Blockers
None for F-001. F-002 is blocked by D3 and D4.
