# Plan F-011: Production hosting on GitHub Pages

- **Spec:** [F-011](../specs/F-011-github-pages-hosting.md). **Status:** In progress.
- **Blocked by:** spec approval; F-002 code merged to `main` (the workflow runs from `main`; `src/taskpane/config.ts` already has the real, non-secret IDs, so the production placeholder guard passes).
- **Branch:** implement on a new branch `f-011-github-pages-hosting` created from the F-002 branch tip or from `main` once F-002 is merged. Do not implement on `f-002-ai-backend`.

## Configuration values

| Value | Value / format | Secret? | Where it lives |
|---|---|---|---|
| Production origin | `https://dimscale.github.io` | No | Entra redirect, Supabase `ALLOWED_ORIGINS` |
| Production base URL (`urlProd`) | `https://dimscale.github.io/DIMSCALE-EMAIL-AI/` | No | `webpack.config.js` |
| Entra SPA redirect (new) | `brk-multihub://dimscale.github.io` | No | Entra app registration |
| `ALLOWED_ORIGINS` (new value) | `https://localhost:3000,https://dimscale.github.io` | No | Supabase function secrets (env) |
| GitHub token for the deploy | automatic `GITHUB_TOKEN` and OIDC | n/a | Created per run by GitHub. No repository secret is added. |

## Manual prerequisites

| # | Who | When | What | Exact steps | Evidence |
|---|---|---|---|---|---|
| M1 | Repo admin of `dimscale/DIMSCALE-EMAIL-AI` | After steps 2 and 3 are merged to `main`, before the first deploy | Enable Pages with Source "GitHub Actions" | 1. Open the repo on GitHub, **Settings**. 2. Sidebar, "Code, planning, and automation", **Pages**. 3. Under "Build and deployment", set **Source** to **GitHub Actions**. 4. Do not pick a suggested workflow template (the repo has its own). 5. Optionally in **Settings > Environments > github-pages**, confirm a deployment branch rule allowing only `main`. Source: [publishing source](https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site). If the setting is missing or greyed out, check the organization's Pages and Actions policies (spec open question 8). | Screenshot or text of the Pages settings |
| M2 | App registration owner (the user) | Before the live Outlook test (step 9) | Add the production SPA redirect | Entra admin center > App registrations > "DIMSCALE AI Email Assistant" (client ID `d157a20c-8820-43ad-8dbf-174a5cade5ea`) > Authentication > Single-page application > Add URI `brk-multihub://dimscale.github.io` (origin only, no path, no trailing slash) > Save. Keep `brk-multihub://localhost:3000`. Do not add any other platform or secret. Source: [NAA setup](https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in). | List of redirect URIs (text) |
| M3 | The user (Supabase CLI is linked to project `qepzpkseyzclxnqguqta`) | Before the live Outlook test (step 9) | Allow the production origin in CORS | From the repo root: `supabase secrets set ALLOWED_ORIGINS="https://localhost:3000,https://dimscale.github.io"`. This replaces the value, so both origins are in the one command. Then run the AC-15 `curl` checks. Whether the running function needs a redeploy to see the new value is not confirmed: if the OPTIONS check does not show the new origin, run `supabase functions deploy whoami --use-api` (as in F-002 step 7) and test again. Source: [Supabase secrets](https://supabase.com/docs/guides/functions/secrets). | curl output (headers only) |
| M4 | Tenant admin (Exchange admin or equivalent; see spec) | After step 8 (site live) and M2, M3 | Centralized deployment | Runbook below. | Dates, who, scope, result |
| M5 | Tenant admin | Before M4 reaches the group | Entra consent and group assignment from F-002 (P1 pending items, AC-21) | As in the F-002 plan. Not repeated here. | F-002 evidence |

### Runbook for M4 (admin)
Source: [Deploy Office Add-ins](https://learn.microsoft.com/microsoft-365/admin/manage/manage-deployment-of-add-ins), [requirements](https://learn.microsoft.com/microsoft-365/admin/manage/centralized-deployment-of-add-ins).
1. Get the manifest: `https://dimscale.github.io/DIMSCALE-EMAIL-AI/manifest.xml` (save the file), or `dist/manifest.xml` from a local `npm run build`. Check it contains no `localhost`.
2. Microsoft 365 admin center > **Settings** > **Integrated apps** > **Add-ins** link near the top > **Deploy Add-in** > **Next** > choose to upload custom apps > **Choose File** > the manifest > **Upload**. (The Microsoft page describes the add-in portal path and says the Integrated apps portal is the recommended one. The exact wording of the custom-upload option in the Integrated apps portal is not verified. Use whichever the admin center shows for an add-in-only XML manifest.)
3. Assign **Just me** first. Deploy. Wait (up to 24 to 72 hours per Microsoft), restart Outlook, and run the live test (step 9).
4. Then change the assignment to the DIMSCALE group (spec open question 9) and announce it to users.
5. Later manifest changes: upload the new manifest (bump `<Version>`), using the Update button of the add-in.

## Steps

| # | Step | ACs | Agent |
|---|---|---|---|
| 0 | The user approves the spec and answers the open questions (at least 1, 4, 5, 6). | prerequisite | user |
| 1 | Update `docs/architecture/security.md` and `overview.md`: Pages data flow, the shared-origin risk and the rule from open question 3, the new redirect and `ALLOWED_ORIGINS` values, the public manifest. | AC-20 | requirements-architect |
| 2 | `webpack.config.js` only: set `urlProd`; add the guard on `https://` and the trailing `/`; in production mode only, set `output.filename` and `output.chunkFilename` to `[name].[contenthash].js`; apply the source-map decision (production only); leave `output.publicPath` unset (default `auto`, to be confirmed on webpack.js.org, and confirmed by AC-5); keep icons and `manifest*.xml` unhashed and the dev settings unchanged. Do not touch `manifest.xml`, `src/` or `supabase/`. | AC-1, AC-2, AC-5, AC-6, AC-7, AC-9, AC-19 | coder |
| 3 | `.github/workflows/<name>.yml` (for example `pages.yml`):<br>- `on: push` (branch `main`), `pull_request`, `workflow_dispatch`.<br>- Top-level `permissions: contents: read`.<br>- Job `build`: checkout, setup-node (Node per open question 11, `cache: npm`), `npm ci --ignore-scripts`, typecheck, lint, test, `npm run build`, `npx office-addin-manifest validate dist/manifest.xml`, then `configure-pages` and `upload-pages-artifact` with `path: dist`, both only when the event is not `pull_request`.<br>- Job `deploy`: `needs: build`, `if: github.event_name != 'pull_request'`, `permissions: pages: write, id-token: write`, `environment: name: github-pages, url: ${{ steps.deployment.outputs.page_url }}`, step `deploy-pages` with `id: deployment`.<br>- `concurrency: group: pages, cancel-in-progress: false`.<br>- Re-check the current major versions on the releases pages (seen 2026-10-22: `deploy-pages` v5.0.1, `upload-pages-artifact` v5.0.0, `configure-pages` v6.0.0, `setup-node` v7.1.0; `checkout` to confirm), then pin each action to the full commit SHA of that release, taken from the action's own repository, with the version in a trailing comment. Check upload-pages-artifact's dotfile exclusion does not drop anything needed (nothing in `dist/` should start with a dot). No `secrets.*`. No `pull_request_target`. | AC-8, AC-10, AC-11 | coder |
| 4 | Local verification of the build: run typecheck, lint, test, `npm run build`, `npm run build:dev`, `npm run validate`, `npx office-addin-manifest validate dist/manifest.xml`; the grep, diff and listing checks for AC-2, AC-4, AC-6, AC-8, AC-9; serve a copy of `dist/` under `/DIMSCALE-EMAIL-AI/` and request every referenced asset (AC-5); build twice and compare; `deno test` in `supabase/functions/whoami` to confirm no regression; `git diff --stat` shows no change in `src/`, `supabase/`, `manifest.xml`. Record real output in the Evidence section. | AC-1 to AC-9, AC-19 | qa-tester |
| 5 | Review of the workflow and build config (security-sensitive CI and supply chain: pass `model: opus`): permissions, pinning, triggers, no secrets, no deploy from PR, `--ignore-scripts` behavior. | AC-8, AC-10 | code-reviewer |
| 6 | Fix Critical and High findings, rerun step 4 checks. | as affected | coder |
| 7 | The user merges to `main` (F-002 first if still unmerged), M1 is done by the repo admin, then the first workflow run happens. If Pages was not enabled before the run, record what `configure-pages` did (open question 7). Record the run URL and result. | AC-12 | user + repo admin + orchestrator |
| 8 | Live checks of the site: `curl -I` for `taskpane.html`, one hashed JS file, `assets/icon-64.png` and `manifest.xml` (status 200, cache headers recorded); compare the served hashed filenames with `dist/` from the same commit; confirm the served manifest has no `localhost`. Then M2 and M3, and the AC-14 and AC-15 evidence. | AC-12, AC-13, AC-14, AC-15 | orchestrator + user |
| 9 | M4 with scope "Just me" first, then live in classic Outlook on Windows: the add-in appears (record the delay), opens on a message, shows the F-001 context and the F-002 status line, console shows no CORS error and no email content. Then widen to the group. | AC-16 | tenant admin + user |
| 10 | Code-only change test: a one-string change on `main` (a branch and PR for it, merged by the user), wait for the deploy, then in Outlook record what is needed to see it (reopen the pane, Ctrl+F5, restart Outlook). Revert the string afterwards through the normal flow. | AC-17 | user + orchestrator |
| 11 | Failure gate: on a throwaway branch add a deliberately failing test, open a PR (and, if wanted, `workflow_dispatch` on that branch). Confirm the `build` job fails, no deploy runs, and the site is unchanged. Delete the branch. | AC-11 | user + orchestrator |
| 12 | Rollback drill: re-run the workflow for the previous good commit (Actions > the run > Re-run all jobs, (a `workflow_dispatch` on another ref builds but never deploys, because the deploy needs `refs/heads/main`)) or revert the last commit on `main`. Confirm the served hashed filenames changed back. Write the chosen procedure (and which one worked) in the Evidence section. A re-run uses the original run's commit; the `github-pages` environment rules still apply. | AC-18 | user + orchestrator |
| 13 | Close: spec status (the user decides Approved and Done), `docs/status.md` (F-011 Done, D7 resolved), commit message `F-011: ...`. | AC-20 | orchestrator |

## Evidence

### 2026-10-22: steps 1 to 4 (re-run by the orchestrator)

| Check | Result |
|---|---|
| typecheck, lint | exit 0 |
| test | 59 passed |
| `deno test` (deno 2.9.6) | 24 passed |
| `npm run build` | passes (3 bundle-size warnings). Hashed `polyfill.*.js`, `react.*.js` and `taskpane.*.js`; unhashed `taskpane.html`, `manifest.xml` and `assets/` |
| `dist/` | no `.map` files |
| `dist/manifest.xml` | 0 `localhost`, `ReadItem` only, validates |

Coder-reported checks:
- AC-5 subpath served locally under `/DIMSCALE-EMAIL-AI/`: all script refs return 200.
- AC-6 hash stability: only the changed bundle got a new name.
- `npm ci --ignore-scripts` build works in a clean copy (Windows).
- Actions pinned to SHAs: checkout v7.0.1, setup-node v7.1.0, configure-pages v6.0.0, upload-pages-artifact v5.0.0, deploy-pages v5.0.1.

Fixed: the config-placeholder guard matched `PROJECT-REF` inside `isConfigured()`. It now matches `//PROJECT-REF.`.

**D7 resolved:** the repo is now public (ADR-0004). Function `ALLOWED_ORIGINS` set to `https://localhost:3000,https://dimscale.github.io`. Live preflight results:
- `dimscale.github.io`: allowed
- `localhost:3000`: allowed
- `evil.example`: no `Access-Control-Allow-Origin`

It took effect without redeploying the function.

### 2026-10-22: step 5 review (code-reviewer, opus), APPROVE WITH COMMENTS
No Critical or High findings. A fork pull request can't deploy, can't get the Pages or OIDC tokens, can't read secrets and can't poison `main`'s npm cache. All 5 action SHAs were verified with `git ls-remote`.

Fixed:
- **Medium:** the workflow-level `concurrency` let PR runs cancel or delay a pending `main` deploy. It is now on the deploy job only.
- **Low:** checkout uses `persist-credentials: false`.
- **Low:** `configure-pages` was dropped (its outputs were unused and it needed the `pages` permission).
- **Low:** the rollback wording in step 12.
- Added `.github/dependabot.yml` (github-actions and npm, weekly).

Repo settings recommended to the user:
- fork PR approval for external contributors;
- `github-pages` environment limited to `main`;
- protection on `main`;
- Dependabot and secret scanning.

Still open: the first Linux CI run is the evidence for `npm ci --ignore-scripts` on Ubuntu.
_To be filled during implementation. Record the real command output for each AC, the run URLs, the `curl -I` headers, and the dates of the live checks. A live Outlook test is claimed only if it was actually performed._
