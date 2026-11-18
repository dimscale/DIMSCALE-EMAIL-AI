# Plan: F-010 compose mode

- **Spec:** [docs/specs/F-010-compose-mode.md](../specs/F-010-compose-mode.md). **No implementation before the user approves the spec (Status: Approved).**
- **Scope:** pane code and `manifest.xml` only. No change under `supabase/`. One MMSI upload at the end.
- **Order constraint:** none with the function. The pane can reach `main` before the manifest is uploaded (the compose view is unreachable until manifest 1.1.0.0 is installed; the read pane does not change). The manifest goes to MMSI **only after** the live checks (steps 10 and 11).
- **Review:** step 1 (permission change) and step 4 (the only write) get an **Opus** `code-reviewer` pass.
- **Evidence:** every step records the real command output in "Evidence" below when it is done.

## Steps
| # | Step | Owner | ACs | Notes |
|---|---|---|---|---|
| 1 | Edit `manifest.xml` per the spec snippet: version 1.1.0.0, `ReadWriteItem`, `MessageComposeCommandSurface` button, `IaCompose.Tooltip`. Update `src/taskpane/manifest.test.ts` (version assertion, replace the "only ReadItem" assertions with `ReadWriteItem` only and the forbidden list, add the compose extension point checks, keep the read extension point assertions). Do **not** touch the dropped test manifest. | coder | AC-1, AC-2, AC-3 | Production build test must cover both extension points and no `localhost`. |
| 2 | Compose read helpers (new small module, for example `composeItem.ts`): `isCompose(item)`, promise wrappers for `subject.getAsync` and `body.getAsync(Text)` (never log errors), `buildComposeRequest`. Reuse the existing 200/4,000 cut and no-lone-surrogate helper from `messageView.ts`. Mode switch at the pane root (`App.tsx`/`MessagePanel.tsx`): compose view vs today's read pane; re-detect on `ItemChanged`. | coder | AC-4, AC-6, AC-7 | The read path stays byte-identical. The mock item throws on `to`, `cc`, `bcc`, `from`. |
| 3 | Compose card component: reuse the Resposta card pieces (`Select`s, tones, error table, refine, "Versão anterior", copy, spinner) from `AiTools.tsx`/`ToolCard.tsx`; no tabs, no auto switch, no auto request; snapshot for refine; abort and clear on `ItemChanged`. | coder | AC-5, AC-8, AC-9, AC-10 | Prefer extracting the shared reply card over copying. Do not edit the `-DIM27` conflict copies. |
| 4 | Insert helper (for example `composeInsert.ts`): `getTypeAsync` then one `prependAsync` with the escaping of `replyForm.ts`/`messageView.ts` (HTML: escaped text, `<br>`, trailing `<br><br>`; Text: text plus blank line); one-click guard; status and alert texts; no fallback to any other write method. | coder | AC-11, AC-12, AC-13 | This is the only write call in `src/` apart from the existing reply-form display. |
| 5 | `strings.ts` PT-PT texts from the spec (notice, empty draft, read error, insert success and error, help line, button labels); 320 px layout check in jsdom guards (`layoutGuard.test.ts` pattern). | coder | AC-14 | The user reviews wording. |
| 6 | Function contract check: add a Deno handler test case only if the empty `fromName`/`fromDomain` reply payload with each tone and with `instruction`+`draft` is not covered; extend `contractParity.test.ts` if it compares client and function payloads. No product change under `supabase/`. | coder | AC-15 | `deno test` needs Deno (not installed on this machine, as recorded in status). If it cannot run, say so; do not claim it. |
| 7 | Docs: add the manifest and permission wording to `docs/admin/2026-10-23-mmsi-request.md` as a PT-PT paragraph for the update (new version 1.1.0.0, permission `ReadItem` to `ReadWriteItem`, compose button, "never sends", ask O-2); note in F-003 (AC-26) and F-012 (permission wording) that the manifest is 1.1.0.0 / `ReadWriteItem` from this feature; `security.md` and `status.md` are already updated by the spec author. | requirements-architect or coder | AC-16 | Documents only. |
| 8 | Verify: `npm run typecheck`, `npm run lint`, `npm test` (vitest), `npm run build`, `npm run build:dev`, `npm run validate`; grep checks of AC-13; record counts (vitest file and test totals not lower than before). | coder | AC-1 to AC-16 | Paste the real output below. |
| 9 | Independent `qa-tester` on the ACs, then `code-reviewer` with `model: opus` (permission change, only write call, untrusted text). Fix Critical and High findings and rerun the affected checks. | qa-tester, code-reviewer | all | |
| 10 | **Sideload checks by the user** with manifest 1.1.0.0 (dev build, as done for earlier manifests, F-012 AC-52 runbook) in **classic Windows, new Outlook, Mac**: new message, reply, forward, opened F-013 draft, inline reply; pin in compose; suggestion, insert, copy, refine; injection draft; read-pane regression. Record date, client, build. | user | AC-17 to AC-22 | If O-1 shows the top-level `Rule`/`FormSettings` must change, amend the spec, change the manifest, repeat from step 8. Nothing goes to MMSI before this passes. |
| 11 | **One MMSI upload** (user): send the final production manifest 1.1.0.0 (file from the CI build, not the dev one) with the paragraph from step 7; MMSI presses **Update**; users restart Outlook; record any extra approval and the time until the button appears. | user + MMSI | AC-23 | Tell MMSI the Entra app, scopes and the Drafts app are unchanged. |
| 12 | Close: spec status (set by the user), `status.md` (backlog F-010, D20), plan evidence, security.md verified against the final code. | orchestrator | all | |

## Evidence (fill in when run)
- Step 8 commands and output: _pending_
- Step 9 reviews: _pending_
- Step 10 live checks (date, client, build, results): _pending_
- Step 11 MMSI answer: _pending_
