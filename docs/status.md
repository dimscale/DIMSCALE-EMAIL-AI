# Project Status

_Last updated: 2026-10-22_

## Current state
- Milestone 0 (SDD setup) is done: the rules, agents, architecture, security, ADRs, test strategy and the F-001 spec and plan are written.
- F-002 is Done: NAA sign-in to the Supabase `whoami` function, live-tested in Outlook (AC-21 pending the admin).
- F-001 is Done: implemented, QA'd, reviewed and live-tested in classic Outlook on Windows. The evidence is in the [plan](plans/F-001-active-message-context.md#evidence).

## Next step
F-011 (GitHub Pages hosting): the spec is being written, then the user approves it. In parallel, the tenant admin:
- revokes the old Graph grants (`Mail.ReadWrite`, `offline_access`, `User.Read`);
- grants admin consent for `access_as_user`;
- sets Assignment required with the group (F-002 AC-21).

## Backlog (prioritized)
| ID | Feature | Tier | Depends on | Status |
|---|---|---|---|---|
| F-001 | Task pane shell and active message context | Active item | ADR-0001 | Done (live-tested in classic Outlook on Windows) |
| F-002 | Authenticated AI backend skeleton (token validation, a health operation, no OpenAI yet) | Backend | ADR-0002, ADR-0003 | Done (live-tested 2026-10-22). AC-21 is pending the admin |
| F-003 | Summarize the open message | Active item + AI | F-001, F-002 | Not started |
| F-004 | Suggest a reply (shown in the panel, copyable) | Active item + AI | F-003 | Not started |
| F-005 | Open a prefilled reply form with the suggestion (`displayReplyFormAsync`; the user sends) | Active item | F-004 | Not started |
| F-006 | Refine the suggestion (shorter, diplomatic, formal, custom) | Active item + AI | F-004 | Not started |
| F-007 | Classification (important, normal, advertising, potentially suspicious) with reasons and an uncertainty notice | Active item + AI | F-003 | Not started |
| F-008 | NAA sign-in and Graph `Mail.Read`: list and select recent messages | Mailbox | ADR-0002 | Not started |
| F-009 | Graph reply draft (`createReply` plus `PATCH`, needs `Mail.ReadWrite`) as an alternative to F-005 | Mailbox | F-008 | Not started |
| F-010 | Compose-mode support (refine the text being written) | Active item + AI | F-006 | Needs research |
| F-011 | Production hosting on GitHub Pages: deploy Action, hashed assets, production manifest, Entra redirect, `ALLOWED_ORIGINS`, centralized deployment by the admin | Platform | ADR-0004, F-002 | Not started |

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
| D7 | Production hosting: ADR-0004 (GitHub Pages) is on hold because the repo is private. Choose between: public repo; GitHub Team; Azure Static Web Apps; a separate public build repo | User | F-011 deploy |
| D8 | Keep the IE11 polyfills (`core-js`, `es6-promise`, browserslist `ie 11`)? Depends on D2 (classic Outlook webview) | User | Bundle cleanup |
| D9 | `package.json` license: MIT (generator default) or proprietary/UNLICENSED | User | — |

## Blockers
F-002 code: none. F-002 deploy and live tests are blocked by D4 and D11.

## Working agreements
- Tell the user before starting any UI or visual design work (requested 2026-10-22).
