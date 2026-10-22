# F-011: Production hosting on GitHub Pages

- **Status:** Draft (awaiting user review)
- **Tier:** Platform (hosting, CI/CD, deployment). No Office.js change, no Graph, no AI, no new backend.
- **Depends on:** [ADR-0004](../decisions/0004-taskpane-hosting.md) (Accepted), F-002 code merged to `main` (the workflow only runs from `main`).
- **Plan:** [docs/plans/F-011-github-pages-hosting.md](../plans/F-011-github-pages-hosting.md)

## Goal
DIMSCALE users get the add-in with zero configuration. A push to `main` runs the checks, builds the task pane and publishes it to `https://dimscale.github.io/DIMSCALE-EMAIL-AI/`. The same build produces a production manifest that the tenant admin uploads once in the Microsoft 365 admin center (Integrated apps). After that, code changes reach users through a normal deploy, with no manifest change.

## Behavior
1. **Build.** `npm run build` (production mode) writes `dist/` with:
   - `taskpane.html` (unhashed name);
   - content-hashed JS (and CSS, if any) files, referenced from `taskpane.html` with relative URLs, so the `/DIMSCALE-EMAIL-AI/` subpath works;
   - `assets/icon-*.png` (unhashed, because the manifest points to them by URL);
   - `manifest.xml`, which is the source `manifest.xml` with every `https://localhost:3000/` replaced by `https://dimscale.github.io/DIMSCALE-EMAIL-AI/`.
2. **Pipeline.** A GitHub Actions workflow runs typecheck, lint, tests, the production build and the production manifest validation. Only if all pass, and only on `main`, it publishes `dist/` with the official Pages actions.
3. **Hand-off.** The production manifest is served at `https://dimscale.github.io/DIMSCALE-EMAIL-AI/manifest.xml` and is also in `dist/`. The admin uploads it (or gives the admin center that URL, see open question 6).
4. **Dev loop unchanged.** `npm start` and `npm run build:dev` still use `https://localhost:3000/` and the source manifest.

## Where each piece lives (tiers)
- Active item (Office.js, `ReadItem`): unchanged. The manifest permission stays `ReadItem`.
- Mailbox (Graph): not used.
- Hosting and deployment: GitHub Pages (static files, public HTTPS), GitHub Actions (build and deploy), Microsoft 365 admin center (centralized deployment of the manifest), Entra (one more SPA redirect), Supabase (one more allowed origin).

## Acceptance criteria
"Mocks" means verifiable offline on the developer machine or in CI without Outlook. "Live" needs GitHub, Entra, Supabase or Outlook.

| ID | Criterion | Verify by |
|---|---|---|
| F-011/AC-1 | `urlProd` in `webpack.config.js` is `https://dimscale.github.io/DIMSCALE-EMAIL-AI/` (trailing slash). `npm run build` succeeds, the existing contoso and config-placeholder guards still exist, and a new guard fails the build if `urlProd` does not start with `https://` or does not end with `/`. A build with `urlProd` reset to `https://www.contoso.com/` fails with the guard message. | Mocks (command output) |
| F-011/AC-2 | In `dist/manifest.xml`, the 7 URLs that were `https://localhost:3000/...` (`IconUrl`, `HighResolutionIconUrl`, the form `SourceLocation`, 3 `bt:Image`, 1 `bt:Url`) start with `https://dimscale.github.io/DIMSCALE-EMAIL-AI/`. The text `localhost` does not appear in `dist/manifest.xml`. A diff between the source `manifest.xml` and `dist/manifest.xml` shows only those URL attribute changes. | Mocks (grep and diff) |
| F-011/AC-3 | `npx office-addin-manifest validate dist/manifest.xml` passes ("The manifest is valid."). Record whether the command needs network access. | Mocks (command output) |
| F-011/AC-4 | `dist/manifest.xml` still has `<Permissions>ReadItem</Permissions>`, only the `MessageReadCommandSurface` extension point, no `Mail.Send`, no `WebApplicationInfo`, no `AppDomains`. The source `manifest.xml` is byte-identical to its state before F-011. | Mocks (diff, grep) |
| F-011/AC-5 | Subpath works: serving `dist/` under the path `/DIMSCALE-EMAIL-AI/` on a local static server, every `src` and `href` in `taskpane.html` resolves to HTTP 200 and none starts with `/` or `http://localhost`. | Mocks (script and curl) |
| F-011/AC-6 | Production JS (and CSS) filenames contain a content hash. `taskpane.html`, `manifest.xml` and `assets/icon-*.png` have no hash. Two builds of the same commit give identical filenames. After changing only a file under `src/taskpane/`, the vendor chunk (`react`, `polyfill`) keeps its filename and the changed chunk gets a new one. | Mocks (build twice, compare listings) |
| F-011/AC-7 | Dev mode is unchanged: `npm run build:dev` produces the previous (unhashed) filenames and a manifest identical to the source one, and `npm run validate` and `npm start` still work with `https://localhost:3000/`. | Mocks for the build and validate. Live for `npm start` (Outlook). |
| F-011/AC-8 | No secret in the bundle or the workflow. A scan of `dist/` and `.github/workflows/` finds no OpenAI key pattern (`sk-`), no `Authorization` literal, no Supabase secret/service-role key, no client secret. The only config values in `dist/` are the non-secret ones in [security.md](../architecture/security.md) (client ID, tenant ID, function URL). The workflow references no `secrets.*` (the automatic `GITHUB_TOKEN` is used only by the Pages actions). | Mocks (scan, review) |
| F-011/AC-9 | `dist/` contains only the build output: no `.env*`, no `supabase/`, no `docs/`, no `node_modules`. The source-map decision (open question 4) is applied and recorded. | Mocks (listing) |
| F-011/AC-10 | The workflow `.github/workflows/<name>.yml` has these properties:<br>- triggers: `push` to `main`, `pull_request` (checks only), `workflow_dispatch`;<br>- top-level `permissions: contents: read`, and the deploy job adds only `pages: write` and `id-token: write`;<br>- the Pages actions are `actions/configure-pages`, `actions/upload-pages-artifact`, `actions/deploy-pages`, and every third-party-or-GitHub action is pinned to a full commit SHA with the version in a comment;<br>- the deploy job declares `environment: github-pages` and `needs` the check/build job, and does not run on `pull_request`;<br>- `concurrency` group `pages` with `cancel-in-progress: false`;<br>- installs with `npm ci --ignore-scripts`. | Mocks (review) |
| F-011/AC-11 | The check/build job runs, in this order, `npm run typecheck`, `npm run lint`, `npm run test`, `npm run build` and `office-addin-manifest validate dist/manifest.xml`. If any step exits non-zero, the deploy job does not run and nothing is published. | Live (GitHub Actions): a throwaway branch with a failing test, run through `workflow_dispatch` or a PR, shows the deploy job skipped or not triggered, and the Pages site unchanged |
| F-011/AC-12 | Pages is enabled with Source "GitHub Actions". A push to `main` produces a successful run whose deploy step reports the URL. `https://dimscale.github.io/DIMSCALE-EMAIL-AI/taskpane.html`, `/manifest.xml` and `/assets/icon-64.png` return 200 over HTTPS, and the hashed JS filenames in the served `taskpane.html` equal those in the `dist/` of that commit. | Live (GitHub, curl) |
| F-011/AC-13 | Cache evidence recorded: `curl -I` of `taskpane.html`, a hashed JS file and an icon, noting the `Cache-Control`, `ETag` and `Last-Modified` headers GitHub Pages actually returns. No claim about a custom header is made, since Pages offers none. | Live (curl, record only) |
| F-011/AC-14 | The Entra app registration has the SPA redirect `brk-multihub://dimscale.github.io` (origin only, no path, platform Single-page application), and still has `brk-multihub://localhost:3000`. | Live (user or admin; screenshot or text from the Entra admin center) |
| F-011/AC-15 | The Supabase function's `ALLOWED_ORIGINS` is exactly `https://localhost:3000,https://dimscale.github.io`. A `curl` OPTIONS from `Origin: https://dimscale.github.io` returns `Access-Control-Allow-Origin: https://dimscale.github.io`. The same from `https://localhost:3000` still works, and from `https://evil.example` returns no `Access-Control-Allow-Origin`. | Live (curl) |
| F-011/AC-16 | After the admin uploads `dist/manifest.xml` in Integrated apps and assigns it (first to the admin only, then to the DIMSCALE group), classic Outlook on Windows shows the add-in on a message in read mode. Opening it shows the F-001 message context and the F-002 status line "signed in" with no CORS error in the console. Record the dates, the Outlook build and how long it took to appear. | Live (admin + user, classic Outlook on Windows) |
| F-011/AC-17 | A code-only change (a visible string) merged to `main` reaches the add-in with no manifest re-upload. Record what the user had to do (reopen the pane, restart Outlook) and how long it took. | Live (Outlook) |
| F-011/AC-18 | Rollback works and is documented: re-running the workflow on the previous good commit (or reverting the commit on `main`) makes the previous hashed bundle the served one. Exercised once. | Live (GitHub) |
| F-011/AC-19 | F-001 and F-002 behavior is unchanged. `git diff` for F-011 touches nothing under `src/` or `supabase/` and does not touch `manifest.xml`. The vitest count (59 at F-002) and the `deno test` count (24) do not drop. | Mocks (diff, command output) |
| F-011/AC-20 | Docs updated: [security.md](../architecture/security.md) covers the Pages data flow, the shared-origin risk and the `ALLOWED_ORIGINS` / redirect values. A short runbook for the admin and the repo admin steps is in the plan. `docs/status.md` is updated at close. | Review |

## Edge cases
- **Shared origin (ADR-0004).** Anything else published under `https://dimscale.github.io/` shares the browser origin: same `localStorage`/`sessionStorage`, same CORS allow-listing. Accepted by ADR-0004: only org members can publish Pages, and a token still needs the user's Entra sign-in. Consequence for this feature: the origin, not the path, is what Entra and Supabase trust. See open question 3 about MSAL's cache location.
- **Pages not enabled.** The workflow must fail visibly (the `configure-pages` step) and publish nothing. Whether `configure-pages` fails, or enables Pages itself, is not stated in the page consulted. Open question 7, verified on the first run.
- **Pages enabled but source is "Deploy from a branch".** Same: a visible failure, no site change.
- **Admin has not deployed the manifest yet.** The site is live but nobody sees the add-in. Not a defect. Developers keep using the sideloaded dev manifest.
- **Manifest changed later** (URLs, permissions, surfaces, buttons, icon file names): the admin must upload the new manifest. Microsoft documents this for line-of-business add-ins. Bump `<Version>` in the manifest for each re-upload (not currently automated, see out of scope).
- **Dev and production manifest share one `<Id>`.** A user who has both the sideloaded dev manifest and the centrally deployed one may see two entries or a conflict. Open question 5.
- **Propagation delay.** Microsoft says up to 24 hours (centralized deployment page) and 24-72 hours (deploy page, ribbon). Test with the admin first.
- **Stale code in a long-lived pane.** Hashed filenames protect against a stale script for a given `taskpane.html`. `taskpane.html` itself can be cached; the cache behavior of Outlook webviews for it is not documented (open question 2).
- **Failed deploy after a good one.** The previous deployment keeps being served (the artifact is only swapped by a successful `deploy-pages`).
- **Concurrent pushes.** The `pages` concurrency group serializes them without cancelling a deploy in progress.
- **Pages limits.** Site at most 1 GB, soft bandwidth limit 100 GB per month, deployment timeout 10 minutes. The bundle is far below these.
- **Free-plan constraint.** The repo must stay public (ADR-0004). A private repo with Pages needs a paid plan, which would change this feature.
- **Org policy.** Organization settings may restrict which actions can run or whether Pages can be published. Open question 8.
- **Admin center group limits.** Centralized deployment supports mail-enabled top-level groups, not nested groups or non-mail-enabled security groups. Open question 9.

## Out of scope
- Custom domain `mailai.dimscale.com` (option B in ADR-0004; a later change of URL, redirect, `ALLOWED_ORIGINS` and a manifest re-upload).
- CSP and other security headers (Pages cannot set custom headers; a `<meta http-equiv>` CSP is a separate decision).
- Any AI feature (F-003 onwards), and any change to F-001 or F-002 code, tests, or specs.
- Changing `manifest.xml` (including a separate Id or version bump for dev, and automatic version bumping). Proposed only as open questions.
- A staging environment or preview deployments for pull requests.
- Unified manifest migration.
- Automatic Entra or Supabase configuration from CI. These stay manual and no credential for them is put in GitHub.

## Security notes
- The deploy uses GitHub's OIDC-based Pages deployment (`id-token: write`) and the automatic `GITHUB_TOKEN`. No repository secret is needed or added.
- Everything under the Pages URL is public, including the manifest. That is acceptable: it holds only non-secret identifiers. Authorization is by Entra token validated in the function, not by hiding the page.
- The workflow runs untrusted pull request code only with `contents: read` and without deploy permissions; it never uses `pull_request_target`.
- Supply chain: actions pinned to commit SHAs, `npm ci --ignore-scripts`, exact versions from the lockfile.

## References
- Custom Pages workflows (permissions `pages: write`, `id-token: write`, `environment: github-pages`, example action versions `configure-pages@v5`, `upload-pages-artifact@v4`, `deploy-pages@v4`): https://docs.github.com/en/pages/getting-started-with-github-pages/using-custom-workflows-with-github-pages
- Publishing source "GitHub Actions" (Settings > Pages > Build and deployment > Source): https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site
- Pages limits (1 GB site, 100 GB/month soft, 10 min deploy timeout; commercial use note): https://docs.github.com/en/pages/getting-started-with-github-pages/github-pages-limits
- Secure use of Actions (pin to full SHA, least-privilege `GITHUB_TOKEN`, avoid `pull_request_target`): https://docs.github.com/en/actions/reference/security/secure-use
- Latest releases seen on 2026-10-22 (the docs page examples are older): `actions/deploy-pages` v5.0.1 (https://github.com/actions/deploy-pages/releases), `actions/upload-pages-artifact` v5.0.0, with a v4 change that dotfiles are excluded and a v5 `include-hidden-files` input (https://github.com/actions/upload-pages-artifact/releases), `actions/configure-pages` v6.0.0 (https://github.com/actions/configure-pages/releases), `actions/setup-node` v7.1.0 (https://github.com/actions/setup-node/releases). `actions/checkout` latest not confirmed (the setup-node notes mention 7.0.1). The coder re-checks and resolves SHAs.
- Deploy Office Add-ins in the Microsoft 365 admin center (Settings > Integrated apps > Add-ins > Deploy Add-in > upload manifest or provide URL; Everyone, Specific users/groups or Just me; 24-72 hours to appear; LOB manifest updates are uploaded by the admin with the Update button; add-in only manifests are supported in the add-in portal): https://learn.microsoft.com/microsoft-365/admin/manage/manage-deployment-of-add-ins
- Requirements for centralized deployment (Exchange Online mailboxes, Exchange admin role, up to 24 hours, nested and non-mail-enabled security groups not supported, `AppsForOfficeEnabled`): https://learn.microsoft.com/microsoft-365/admin/manage/centralized-deployment-of-add-ins
- NAA redirect `brk-multihub://<origin>`, origin only, no subpath, platform SPA: https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in
- Office cache (web cache and Wef cache exist; reload a task pane with Ctrl+F5; clear the cache after manifest changes such as icon file names): https://learn.microsoft.com/office/dev/add-ins/testing/clear-cache
- Webpack `output.publicPath` default `auto` and `[contenthash]`: to be confirmed by the coder on https://webpack.js.org/guides/caching/ and https://webpack.js.org/configuration/output/#outputpublicpath (not fetched in this session).

## Open questions
1. **Repo admin.** Who has admin rights on `dimscale/DIMSCALE-EMAIL-AI` to set Pages Source to "GitHub Actions"? The consulted page lists no required role. Repository admin is the usual requirement but is not confirmed.
2. **Outlook webview cache.** Microsoft's pages say a web cache exists and that Ctrl+F5 reloads a task pane, but give no rule for HTTP cache headers or how long `taskpane.html` stays cached in classic Outlook, new Outlook or Outlook on the web. GitHub Pages' own `Cache-Control` for the site is also not stated in the docs consulted (third-party sources suggest a short fixed lifetime, not verified). The spec therefore relies on hashed asset names plus measurement (AC-13, AC-17) and makes no promise about `taskpane.html` freshness.
3. **MSAL cache location on a shared origin.** F-002 open question 3 (cache location) was left at the MSAL default. With another Pages site on `dimscale.github.io`, browser storage is shared. Does the user accept the ADR-0004 risk as is, or should F-011 add a requirement (for example, "no other Pages site in the org without review")? Proposed default: accept, and record a rule that no other Pages site is published under `dimscale.github.io` without review.
4. **Source maps in production.** The config uses `devtool: "source-map"` for both modes. Proposed default: no source maps in the production artifact (smaller, and nothing extra published), keeping them for dev. The source is public anyway. Confirm.
5. **Dev and production manifest `<Id>`.** Both use `05c2e1c9-3e1d-406e-9a91-e9ac64854143`. Keep it (and have developers remove the sideloaded copy when testing the central one), or give the dev manifest its own Id and name? The second changes `manifest.xml`, which this spec excludes, so it needs a separate decision.
6. **How the admin gets the manifest.** Upload the file from `dist/` or from a downloaded copy, or give the admin center the URL `https://dimscale.github.io/DIMSCALE-EMAIL-AI/manifest.xml`? The admin page offers both. Whether a URL-based upload re-fetches automatically later is not stated. Proposed default: upload the file, and re-upload on manifest changes.
7. **Behavior of `configure-pages` when Pages is not enabled.** Not stated in the page consulted. To be verified on the first run and recorded.
8. **Org policies.** Is there an org-level restriction on allowed GitHub Actions, or on publishing Pages, for `dimscale`? GitHub's limits page also says Pages is not meant for commercial SaaS. This is an internal add-in, but DIMSCALE should confirm the use fits GitHub's terms.
9. **Which group gets the add-in in the admin center.** Centralized deployment does not support nested or non-mail-enabled security groups. Is the Entra group chosen for F-002/AC-21 (D10) a mail-enabled, top-level group? If not, the admin center assignment needs another group, and the two lists must be kept in sync (Entra enforces who can sign in; the admin center only decides who sees the button).
10. **Licenses and Exchange.** Are all target users on a license listed by Microsoft and on Exchange Online mailboxes with `AppsForOfficeEnabled` true (D2 also asks which Outlook clients are used)?
11. **Node version in CI.** Local is Node 24. Proposed default: `node-version: 24`. Confirm it matches the project's intent (`package.json` has no `engines`).
12. **Does `office-addin-manifest validate` call a remote service?** Not confirmed. If it does, CI validation depends on network availability, and a flaky result should not block a deploy without a decision from the user.

## Risks
- Shared origin (see edge cases) and Pages `Cache-Control` unknown (open question 2).
- The first deployment depends on manual steps by three different people (repo admin, user, tenant admin). The plan orders them.
- `office-addin-dev-certs` is `require`d at the top of `webpack.config.js`. If `npm ci --ignore-scripts` leaves it unable to load in CI (for example, `keytar`), the build fails in CI only. The coder checks this on the first run and reports instead of removing `--ignore-scripts`.
- Pinning to SHAs stops automatic security updates for the actions (GitHub notes Dependabot does not alert on SHA-pinned actions). Review the pins when touching the workflow.
