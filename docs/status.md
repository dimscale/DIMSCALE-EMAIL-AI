# Project Status

_Last updated: 2026-10-22 (afternoon)_

## Current state
- F-001, F-002 and F-011 are Done. The add-in is live at `https://dimscale.github.io/DIMSCALE-EMAIL-AI/`, deployed by GitHub Actions from `main`, and tested in classic Outlook from the production manifest (sideloaded by the user).
- The repo `dimscale/DIMSCALE-EMAIL-AI` is public (ADR-0004). The Supabase `whoami` function is deployed in eu-west-1.
- In progress: F-012 (task pane design, spec being written from the user's reference mockup) and F-013 (server-side automatic reply drafts; ADR-0005 and the spec being written).

## Next step
1. The user reviews the F-012 spec and picks a visual direction from the prototype.
2. The user reviews ADR-0005 and the F-013 spec, then takes the questions to the tenant admin (MMSI) and the DPO.
3. Centralized deployment by the admin comes after the design (user decision).

Pending with the tenant admin:
- revoke the old delegated Graph grants (`Mail.ReadWrite`, `offline_access`, `User.Read`);
- grant admin consent for `access_as_user`;
- set Assignment required with the group (F-002 AC-21).

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
| F-011 | Production hosting on GitHub Pages: deploy Action, hashed assets, production manifest, Entra redirect, `ALLOWED_ORIGINS`, centralized deployment by the admin | Platform | ADR-0004, F-002 | Done 2026-10-22 (admin deployment deferred until after design) |
| F-012 | Task pane UI/UX design (DIMSCALE brand, Fluent UI, tabs Analisar/Responder/Escrever per the user's reference) | Active item | F-001, F-002 | Spec being written |
| F-013 | Server-side automatic reply drafts (Graph change notifications, Supabase, never sends) | Backend + Mailbox | ADR-0005 | ADR and spec being written |

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
| D6 | UI language: PT-PT proposed, following the user's reference mockup. To be confirmed with F-012. Strings are in `src/taskpane/strings.ts` | User | F-012 |
| D12 | Automatic drafts for incoming mail with Outlook closed (requested 2026-10-22): permission model, filters, and the GDPR/DPO decision (see ADR-0005) | User / admin / DPO | F-013 |
| D7 | ~~Production hosting~~: GitHub Pages with the repo made public, decided 2026-10-22 (ADR-0004) | User | — |
| D8 | Keep the IE11 polyfills (`core-js`, `es6-promise`, browserslist `ie 11`)? Depends on D2 (classic Outlook webview) | User | Bundle cleanup |
| D9 | `package.json` license: MIT (generator default) or proprietary/UNLICENSED | User | — |

## Blockers
F-002 code: none. F-002 deploy and live tests are blocked by D4 and D11.

## Working agreements
- Tell the user before starting any UI or visual design work (requested 2026-10-22).
