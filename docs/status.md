# Project Status

_Last updated: 2026-10-29 (pane merged to main, refine reply, single ribbon button)_

## Current state
- **Update 2026-10-29 (merged to `main`, live on GitHub Pages):** the Kutools-style pane is built and live-tested by the user in Outlook: tabs Resumo | Resposta | Traduzir | Automático, cards, translate (automatic PT/EN), reply language and tone, and **"Pedir alterações"** (F-003 amendment d: refine the suggestion, "Versão anterior", reviewed by Opus and QA'd, function `analyze-message` redeployed). Manifest 1.0.2.0: `DisplayName` "Assistente IA", **one** ribbon button "DIMSCALE Mail" (the Resumo button was removed on the user's request; `resumo.html` is still built but not linked). Auto mode stays hidden (`AUTO_SUMMARY_AVAILABLE=false`) until D5. Production manifest and pane verified after deploy (no `localhost`, taskpane bundle contains the refine UI).
- **Not yet done:** live checks at 320 px and 200 % zoom, screen reader, injection with a live model; redeploy of `graph-webhook` and `auto-draft-cron` (shared `stripMarkers` changed, no behaviour change); the custom-tab experiment (F-012 AC-37 to AC-42); the MMSI request (F-013 live) once the user says the add-in is ready; DPO D5 and D15.
- **Older notes below are partly superseded by this update** (for example "not yet built" for the pane and the two-button manifest).
- **Kutools-style pane (user decisions 2026-10-29, specs amended, not yet built):** tabs **Resumo | Resposta | Traduzir | Automático** (Automático = the former Responder tab), each tool a card with Copiar/Expandir; **auto mode** for the summary (per-user switch, **off by default**, dialog before it turns on, preference in `roamingSettings`, counts against the 50/1,500 caps, 800 ms debounce, pauses on the cap); new **`translate`** action (pt | en); **reply options** language match/pt/en and tone (part of F-006); ribbon group "DIMSCALE MAIL" with **"IA"** and **"Resumo"** buttons (manifest 1.0.2.0). See [F-003](specs/F-003-analyze-summary-and-reply.md) AC-35 to AC-51, [F-012](specs/F-012-taskpane-design.md) AC-43 to AC-54, plans F-003 steps 19 to 36 and F-012 steps 17 to 24. **The architect's proposals (F-003 Q-U10 to Q-U15, F-012 Q-K1 to Q-K6) await the user's confirmation before the coder starts.**
- **Manifest order:** the ribbon change (two buttons, second page `resumo.html`) is built, deployed to Pages and **checked live by sideload (F-012 AC-52) before anything is sent to the tenant admin**; the admin then uploads the final manifest once (F-012 AC-53). Microsoft does not document how Outlook handles two task pane buttons (`TaskpaneId` is not supported in Outlook, no shared runtime), hence the live check.
- **D5 and auto mode:** auto mode sends every email the user opens on the Resumo tab, so it can't stay "test emails only" on a real mailbox. Proposed: `AUTO_SUMMARY_AVAILABLE = false` in production until the DPO answers D5 (new DPO question 5 in F-003).
- **F-003 (merges F-003, F-004 and most of F-005; Approved 2026-10-29):** backend (`analyze-message`, migration `analyze_usage`/`an_take`) and pane built, independent review done. **Pending: security-review fixes verified, deploy (`supabase db push`, `OPENAI_MODEL`, function deploy), pane deploy, live tests (AC-27 to AC-32).** Decisions: model `gpt-6-luna`, caps 50 per user and 1,500 global per UTC day, reply mirrors the email's language, **test emails only until D5 (DPO)**. Outputs contain no link or email address at all. Data flow in [security](architecture/security.md#third-party-processing-openai).
- F-001, F-002 and F-011 are Done. The add-in is live at `https://dimscale.github.io/DIMSCALE-EMAIL-AI/`, deployed by GitHub Actions from `main`, and tested in classic Outlook from the production manifest (sideloaded by the user).
- The repo `dimscale/DIMSCALE-EMAIL-AI` is public (ADR-0004). The Supabase `whoami` function is deployed in eu-west-1.
- **F-012 Approved (2026-10-23), in progress:** direction D "Outlook nativo" (Fluent stock themes + Office theme colours), 320 px compact, tabs Analisar | Responder, "DIMSCALE MAIL" group on the default ribbon tab (no `CustomTab`), pinnable pane with `ItemChanged`. Spec amended 2026-10-23 to match.
- **F-013 is the priority (user direction 2026-10-23):** automatic reply drafts only for **contacts the user selects** in the Responder tab, working with Outlook closed. Spec, plan and ADR-0005 amended 2026-10-23. Their status lines are set by the user (Approved / Accepted); the orchestrator reports the user approved them on 2026-10-23 and records it.

- **User answers 2026-10-23:** ribbon button label "DIMSCALE MAIL"; wants an own "DIMSCALE MAIL" tab (to be proven by an experiment, F-012 AC-37 to AC-40); Outlook clients in use are classic Windows, new Outlook and Mac; contacts cap 50 confirmed; a pinned pane stays open across messages until closed, reopened by the "DIMSCALE MAIL" button.
- **F-013 backend in progress** (Supabase functions, other agent). The **OpenAI key is configured in Supabase as `OPENAI_API_KEY` (2026-10-23)**; D5 (account/model/GDPR acceptance) still open. `security.md`, `overview.md` and `requirements.md` (R8) updated for F-013 (plan step 1 and 10c done).
- F-012 spec amended: message excerpt collapsed by default in a "Mensagem" accordion (orchestrator decision).

## Next step
0. **Kutools-style pane (2026-10-29):** the user confirms the amended F-003/F-012 specs (Q-U10 to Q-U15, Q-K1 to Q-K6). Then, in parallel: function changes (F-003 plan steps 20 to 25, review with opus), pane client (step 26), manifest/icons/`resumo.html` (F-012 step 18); then the pane UI in sequence (F-012 step 17 → F-003 steps 27, 28 → F-012 step 19, review with opus). Deploy the function **before** merging the pane; then the live ribbon check (F-012 step 23) and only then the admin upload (F-012 step 24). This replaces item 2 below for the manifest content.
1. **F-012 fixes** (coder): button label "DIMSCALE MAIL", collapsed "Mensagem" accordion, PT-PT literals, 40 px header, contrast guard, version bump, SVG icon master. Then QA and review. Merge to `main`; the F-011 workflow deploys.
1b. **Custom-tab experiment** (F-012 AC-37 to AC-42): the coder prepares `manifest.customtab-test.xml` (own `<Id>`, never deployed); the user sideloads it in classic Windows, new Outlook and Mac and reports; production adopts it only if all three behave. Mac live checks (NAA, pinning from 16.13, theme) are in AC-42 and F-013 AC-42.
2. **Production manifest for centralized deployment:** generated by CI after F-012 is merged and deployed (`https://dimscale.github.io/DIMSCALE-EMAIL-AI/manifest.xml`, with the ribbon entry and pinning). The admin uploads it once (Integrated apps). Later code changes deploy without the admin; only manifest changes need a re-upload.
3. **MMSI checklist pending:** the user sends [docs/admin/2026-10-23-mmsi-request.md](admin/2026-10-23-mmsi-request.md) and collects the returned values.
4. DPO: D5 (OpenAI) and D15 (contact list storage).
5. F-013 implementation per its plan once the admin values and D5 are in (step 2 live spike first). Then F-015, F-016.

Pending with the tenant admin (all in the checklist above):
- add-in app: remove `Mail.ReadWrite`, `User.Read`, `offline_access`; add delegated `Mail.ReadBasic`; admin consent (`access_as_user`, `openid`, `profile`, `Mail.ReadBasic`); Assignment required with the group (F-002 AC-21);
- F-013: mail-enabled group with direct members, app "DIMSCALE Mail AI Drafts" with certificate, Exchange RBAC for Applications (`Application Mail.Read`, `Application Mail.ReadWrite`) scoped to the group, no Entra application consent, never `Mail.Send`;
- later: centralized deployment of the final manifest (**1.0.2.0 with the "IA" and "Resumo" buttons, only after F-012 AC-52/AC-53**; do not send an earlier manifest).

## Backlog (prioritized)
| ID | Feature | Tier | Depends on | Status |
|---|---|---|---|---|
| F-001 | Task pane shell and active message context | Active item | ADR-0001 | Done (live-tested in classic Outlook on Windows) |
| F-002 | Authenticated AI backend skeleton (token validation, a health operation, no OpenAI yet) | Backend | ADR-0002, ADR-0003 | Done (live-tested 2026-10-22). AC-21 is pending the admin |
| F-011 | Production hosting on GitHub Pages | Platform | ADR-0004, F-002 | Done 2026-10-22 (admin deployment after F-012) |
| F-012 | Task pane design: direction D, 320 px compact, tabs Resumo/Resposta/Traduzir/Automático (amended 2026-10-29, formerly Analisar/Responder), "DIMSCALE MAIL" ribbon group with "IA" and "Resumo" buttons, pinnable | Active item + manifest | F-001, F-002 | Approved 2026-10-23, in progress; amendment 2026-10-29 awaits the user's confirmation of Q-K1 to Q-K6 |
| F-013 | Automatic reply drafts for **selected contacts** (Responder tab allow-list, delegated `Mail.ReadBasic` for candidates, `auto-reply-settings` function, Graph change notifications, app-only RBAC-scoped, never sends) | Backend + Mailbox + pane | ADR-0005, F-012 (tab shell), D5, D15, admin checklist | Spec amended 2026-10-23; next after F-012 |
| F-015 | Manual batch reply drafts (tick up to 10 messages) | Mailbox + Backend | F-013 (draft engine, shared auth), F-012 | Spec draft; after F-013 |
| F-016 | Per-user style profile from ticked sent replies | Mailbox + Backend | F-015 | Spec draft; after F-015 |
| F-003 | Summarize the open message (plus F-004 and the reply-form part of F-005, F-012 "Resumo da conversa"); amended 2026-10-29: auto mode, `translate`, reply language/tone | Active item + AI | F-001, F-002 | Approved 2026-10-29; built, reviewed, backend deployed; live tests pending. Amendment 2026-10-29 awaits the user's confirmation of Q-U10 to Q-U15 |
| F-004 | Suggest a reply (shown in the panel, copyable) | Active item + AI | F-003, F-013 engine | Merged into F-003 |
| F-005 | Open a prefilled reply form (`displayReplyFormAsync`; the user sends) | Active item | F-004 | Delivered by F-003 (reply-to-sender only) |
| F-006 | Refine the suggestion | Active item + AI | F-004 | Fixed tone and language options moved into F-003 (2026-10-29); free-text refinement not started |
| F-007 | Classification with reasons and an uncertainty notice | Active item + AI | F-003 | Not started |
| F-008 | ~~Graph list of messages~~ | — | — | Superseded by F-015 (list) and F-013 (sender list) |
| F-009 | ~~Graph reply draft~~ | — | — | Superseded by F-013 / F-015 |
| F-010 | Compose-mode support | Active item + AI | F-006 | Needs research |

## Open decisions
| # | Decision | Owner | Blocks |
|---|---|---|---|
| D1 | ~~Approve ADR-0001~~ (accepted 2026-10-22) | User | — |
| D2 | ~~Outlook clients~~: **classic Windows, new Outlook and Mac** (2026-10-23). Live checks in F-012 AC-33/36/41/42 and F-013 AC-41/42; web is not a target | User | — |
| D3 | ~~Backend vendor~~: Supabase (ADR-0003) | User | — |
| D4 | ~~Entra app registration~~ (done; client ID `d157a20c-…`) | Tenant admin | — |
| D5 | The OpenAI account, model, and GDPR acceptance of sending mail to OpenAI (F-013 chosen contacts; F-003 the open message on click, incl. quoted history, and the pseudonymised-`oid` usage counter; **since 2026-10-29 also F-003 auto mode, every email opened on the Resumo tab when the user opted in, and `translate`**). **Model chosen 2026-10-29: `gpt-6-luna`; DPO answer pending** (F-003 uses test emails only until then) | User / DPO | F-013, F-003 real-mail use |
| D6 | ~~UI language~~: PT-PT (2026-10-22) | User | — |
| D7 | ~~Production hosting~~: GitHub Pages (ADR-0004) | User | — |
| D8 | Keep the IE11 polyfills? Depends on D2 | User | Bundle cleanup |
| D9 | `package.json` license: MIT or UNLICENSED | User | — |
| D10 | Access limited to one Entra group via Assignment required (needs Entra ID P1/P2). Which group, and is P1/P2 licensed? | User / admin | F-002 AC-21 |
| D11 | ~~Supabase project~~: `qepzpkseyzclxnqguqta`, eu-west-1 | User | — |
| D12 | ~~Automatic drafts trigger~~: selected contacts (allow-list), 2026-10-23. Permission model per ADR-0005 | User | — |
| D13 | ADR for the stateful backend (needed by F-015/F-016; F-013's tables are covered by ADR-0005) | User | F-015, F-016 |
| D14 | MMSI checklist returned (add-in app clean-up, `Mail.ReadBasic` consent, Assignment required, Drafts app, RBAC, group) | Tenant admin | F-013 spike and live |
| D15 | DPO: storing each user's chosen contact addresses in Supabase (retention, lawful basis, notice) | DPO | F-013 live |
| D16 | F-012 Q-R2 answered: label "DIMSCALE MAIL". Q-R1 answered: user wants a custom tab; experiment pending (AC-37 to AC-40), result decides the production manifest | User | Custom tab only |
| D18 | Auto mode before D5: keep `AUTO_SUMMARY_AVAILABLE = false` in production until the DPO answers (proposed), or enable now (F-003 Q-U13) | User | Auto mode in production |
| D19 | Confirm the architect proposals of the 2026-10-29 amendment: preference in `roamingSettings`, auto only on the Resumo tab, translate limits, cap sharing, default target (F-003 Q-U10 to Q-U15); "DIMSCALE MAIL" button becomes "IA", in-pane expand, drop "Melhorar texto e tom", tab label fallback, options not remembered, which custom-tab variant was tested (F-012 Q-K1 to Q-K6) | User | Amendment implementation |
| D17 | F-013 Q-C2 answered (50 contacts). Still open: Q-C1 (candidate source), Q-C3 (filters 6/12 for listed contacts), Q-C4 (forged senders) | User | F-013 implementation |

## Blockers
F-012: none. F-003 deploy and live tests: none (no tenant admin needed); use on real mail: D5. Kutools-style amendment: D19 (user confirmation) before coding; auto mode in production: D18/D5; admin upload: F-012 AC-52/AC-53. F-013 implementation: D5, D14, D15 (and the user's approval of the amended spec/ADR recorded by the orchestrator).

## Working agreements
- Tell the user before starting any UI or visual design work (requested 2026-10-22).
