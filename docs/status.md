# Project Status

_Last updated: 2026-10-22_

## Current state
- Milestone 0 (SDD setup) is done: the rules, agents, architecture, security, ADRs, test strategy and the F-001 spec and plan are written.
- F-001 is Done: implemented, QA'd, reviewed and live-tested in classic Outlook on Windows. The evidence is in the [plan](plans/F-001-active-message-context.md#evidence).

## Next step
F-002 code is done and verified offline (see the plan's Evidence section). Waiting for: the tenant admin to do P1 (app registration, group assignment, then send the client ID and tenant ID), and the user to do P2 (Supabase CLI login and link, project ref) and to accept the "not configured" state. Then: deploy (step 7) and the live tests (step 10).

## Backlog (prioritized)
| ID | Feature | Tier | Depends on | Status |
|---|---|---|---|---|
| F-001 | Task pane shell and active message context | Active item | ADR-0001 | Done (live-tested in classic Outlook on Windows) |
| F-002 | Authenticated AI backend skeleton (token validation, a health operation, no OpenAI yet) | Backend | ADR-0002, ADR-0003 | Code done and verified offline; live steps pending |
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
| D3 | ~~Backend vendor~~: Supabase, accepted 2026-10-22 (ADR-0003) | User | — |
| D4 | The Entra app registration and consent are done by the DIMSCALE tenant admin (agreed 2026-10-22). Waiting for the client ID and tenant ID | Tenant admin | F-002 live steps 7 and 10 |
| D10 | Access limited to one Entra group, via Assignment required (needs Entra ID P1 or P2). Which group, and is P1 or P2 licensed? | User / admin | F-002 AC-21 |
| D11 | ~~Supabase project~~: ref `qepzpkseyzclxnqguqta`, region `eu-west-1` (Ireland, EU). Linked with the CLI on 2026-10-22 | User | — |
| D5 | The OpenAI account, model, and GDPR acceptance of sending email content to OpenAI | User / DPO | F-003 |
| D6 | UI language (PT or EN). EN is used for now, with strings in `src/taskpane/strings.ts` | User | — |
| D7 | Production hosting URL (replaces the `contoso.com` placeholder; the production build fails until it is set) | User | Any deployment |
| D8 | Keep the IE11 polyfills (`core-js`, `es6-promise`, browserslist `ie 11`)? Depends on D2 (classic Outlook webview) | User | Bundle cleanup |
| D9 | `package.json` license: MIT (generator default) or proprietary/UNLICENSED | User | — |

## Blockers
F-002 code: none. F-002 deploy and live tests are blocked by D4 and D11.
