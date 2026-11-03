# ADR-0007: Hosting for the task pane (private repo, stable domain, internal use)

- **Status:** Accepted (2026-11-03) by the user ("aprovo adr-0007"). Earlier: Proposed (2026-11-03). Supersedes ADR-0004 once accepted (ADR-0004 stays as history).
- **Date:** 2026-11-03
- **Deciders:** the user. Inputs: orchestrator, tenant admin (MMSI) for the DNS record and the manifest upload.

## Context
ADR-0004 chose GitHub Pages from a **public** repo (`dimscale/DIMSCALE-EMAIL-AI`) because free Pages needs a public repo and Azure needed a card. The user's goals now (2026-11-03):
1. ~~Make the product commercial.~~ **Corrected by the user (2026-11-03): internal use at DIMSCALE only, not sold to other companies.** The commercial-use terms below are kept as background; the decisive reasons are now goals 2 to 4 (private repo, free, no dependency on the user's PC, no new MMSI upload for updates). No multi-tenant work, AppSource listing or customer DPA is in scope.
2. The repo must **not** be public.
3. Nothing may depend on the user's PC (`npm start` is for development only).
4. After the tenant admin (MMSI) installs the manifest once, **updates must not require sending anything to MMSI again**.

Constraints: no Azure (card), no new platform without an accepted ADR (CLAUDE.md rule 8), no access yet to the `dimscale.com` DNS (the user gets it later today, or asks MMSI).

Current coupling to the hosting origin (all must change when the origin changes): manifest URLs (`SourceLocation`, icons, `bt:Url`), `urlProd` in `webpack.config.js`, the Supabase secret `ALLOWED_ORIGINS` (`https://localhost:3000,https://dimscale.github.io`), and the Entra SPA redirect URIs `brk-multihub://localhost:3000` and `brk-multihub://dimscale.github.io`.

## Key design rule: the manifest points to a stable domain we own
The manifest must point to `https://mailai.dimscale.com/` (proposed hostname, a subdomain of DIMSCALE's own domain), never to a provider hostname (`*.github.io`, `*.pages.dev`, ...). The provider is then reached only through one DNS `CNAME`. Changing provider later is a DNS change plus our own settings, with **no new manifest upload**, because the manifest, the Entra redirect `brk-multihub://mailai.dimscale.com` and `ALLOWED_ORIGINS` keep the same value.

Migrating *to* this domain is itself a manifest change (the URLs change), so there is exactly **one** last upload to MMSI. **Do not send MMSI a manifest that points to `dimscale.github.io`**; send only the final one with `mailai.dimscale.com` (see the open question on whether one was already sent).

### What does and does not need a new manifest upload
Sources: Microsoft Learn says that for an add-in an admin uploaded, "the admin must upload the new manifest file to support metadata changes", that "The web application where the code and logic for the add-in runs can change at any time", that "Whenever you make a change to the manifest, you must raise the version number", and that changes to permissions, `Scopes` or `Events` require the admin to consent again ("Users remain on the existing version of the add-in until the admin consents to the updates").

| Change | New manifest to MMSI? |
|---|---|
| Task pane code: HTML, JS, CSS, React, i18n text, new tabs, new features that need no new permission | **No** (served from the host; "can change at any time") |
| Supabase functions, prompts, caps, migrations, `ALLOWED_ORIGINS` for an unchanged origin | **No** |
| Moving the site to another provider while `mailai.dimscale.com` keeps resolving to it (DNS `CNAME` change, same HTTPS hostname, same paths) | **No** |
| Any manifest field: `DisplayName`, descriptions, ribbon group/button labels, `Version`, add-in commands, `SourceLocation`/icon/`bt:Url` URLs, `AppDomains`, a different path or hostname | **Yes** (and a `<Version>` bump) |
| Permission change (`ReadItem` to something else), `Scopes`, `Events` | **Yes**, plus admin consent again |
| New Graph scope or app role in Entra (not a manifest field) | No manifest, but an Entra admin consent by the tenant admin |
| Icon image **files** replaced at the same URL | Not documented whether Outlook re-fetches them; treat as unknown (Open question Q-7) |

Source pages: https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins (section "Security and permission considerations for admins", "Updates for add-ins happen as follows"), https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish ("Deploy updates"), https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests ("Hosting requirements": HTTPS; images must be cacheable, no `no-cache`/`no-store`).

## Options
Facts below were read on the cited pages on 2026-11-03. "Not verified" means the page did not say; it is an open question, not an assumption.

| | **A. Cloudflare Pages + private GitHub repo** (recommended) | B. GitHub Pages, private repo, paid plan | C. Netlify | D. Vercel | E. Azure Static Web Apps | F. Keep GitHub Pages, public repo |
|---|---|---|---|---|---|---|
| Private repo | Yes: "Both private and public repositories are supported" (GitHub); org repos via the "Cloudflare Workers and Pages" GitHub App | Yes, on Team ("Pages and wikis: Included" for Team; Free lists public repos only) | Org-owned private repos are listed under **Pro** only | Not verified | Not evaluated (rejected) | No (goal 2 fails) |
| Commercial use | No blanket ban found in the Self-Serve Subscription Agreement; it bars collecting card data on web properties receiving Free Services (s.2.2.1(h)) and reselling access (s.2.2.1(a)). The Pages/Workers service-specific terms were not found (Q-2) | **Risky**, see quote below; clause targets "free" hosting and "commercial software as a service (SaaS)", our product is SaaS-like; not resolved by paying | Netlify staff say free tier may be used commercially (forum, 2021 to May 2025); the Self-Serve Subscription Agreement was not read (Q-2) | **Hobby: "restricts users to non-commercial, personal use only"** so Pro is mandatory | Not evaluated | Same quote as B, and it applies today |
| Cost | $0 (Free: 500 builds/month, 20,000 files, 25 MiB per file, unlimited previews) | Team $4 per user per month (every org member who needs private repos; seat count not verified) | Pro $20 per month (unlimited members), 3,000 credits; Free $0, 300 credits | Pro: Developer seat $20 per user per month | Not evaluated | $0 |
| Custom subdomain **without moving the DNS zone**, record | `CNAME mailai -> <project>.pages.dev`. Docs: "it is not necessary for your site to be a Cloudflare zone"; apex would need a Cloudflare zone, a subdomain does not. The domain must first be added in the dashboard, else the CNAME gives 522 | `CNAME mailai -> dimscale.github.io` ("excluding the repository name") | `CNAME mailai -> <project>.netlify.app` | `CNAME mailai -> <unique value shown per project>` (example in docs: `d1d4fc829fe7bc7c.vercel-dns-017.com`); a `TXT` at `_vercel` only if the domain is used in another Vercel account | `CNAME mailai -> <name>.azurestaticapps.net` | `CNAME mailai -> dimscale.github.io` (ADR-0004 option B) |
| HTTPS certificate | Docs on known issues: verification fails if Cloudflare cannot validate the certificate request at `http://<domain>/.well-known/acme-challenge/*`; no statement that the certificate is free (Q-3) | Optional "Enforce HTTPS", "can take up to 24 hours before this option is available" | "SSL certificates may require additional verification steps" | Not stated on the page read | "Free SSL/TLS certificates are automatically created" | As B |
| Preview deployments per branch | Yes: every non-production branch is built; branch alias `<branch>.<project>.pages.dev` (e.g. `fix-api.<project>.pages.dev`); unlimited; **public by default**; optional access policy restricts to the Cloudflare account; preview branches configurable (all, none, include/exclude) | None documented in the pages read; one Pages site per publishing source (not verified) | "Unlimited deploy previews" (Free) | Not verified | Preview environments exist; "Adding a custom domain to a preview environment is not supported" | None |
| Build from GitHub | Yes, Git integration (production branch + previews), or build inside the project; default Node 22.23.4, set `NODE_VERSION` or `.node-version` | Yes, GitHub Actions (today's `pages.yml`) | Yes | Yes | Yes | Yes (today) |
| Meets all four goals | Yes, pending Q-1 to Q-5 | Goals 1 and 4 at risk (SaaS clause), no per-branch previews | Only with Pro ($20 per month) | Only with Pro | Rejected (card) | Fails goals 1 and 2 |

### GitHub Pages prohibited-uses clause (quoted, as returned from the page)
"GitHub Pages is not intended for or allowed to be used as a free web-hosting service to run your online business" and the restriction also covers any site "primarily directed at either facilitating commercial transactions or providing commercial software as a service (SaaS)." Source: https://docs.github.com/en/pages/getting-started-with-github-pages/github-pages-limits ("Usage limits"). The same page lists soft limits: 1 GB site, 100 GB per month bandwidth. The quote was extracted by a fetch tool, so read the clause on the page before accepting this ADR.

### Vercel Hobby terms (quoted)
"the Hobby plan restricts users to non-commercial, personal use only." Source: https://vercel.com/docs/plans/hobby

### Remaining facts
- GitHub Pages from a private repo: "If the account that owns the repository uses GitHub Free or GitHub Free for organizations, the repository must be public." https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site . Pages **access control** (a site only the org can read) needs GitHub Enterprise Cloud: https://docs.github.com/en/pages/getting-started-with-github-pages/changing-the-visibility-of-your-github-pages-site . That would also block Outlook from loading the pane, so it is not useful here. On Team the site is published publicly from a private repo (the page does not state the default; Q-6).
- Cloudflare Pages limits: https://developers.cloudflare.com/pages/platform/limits/ . Custom domains: https://developers.cloudflare.com/pages/configuration/custom-domains/ . Previews: https://developers.cloudflare.com/pages/configuration/preview-deployments/ . Branch build controls: https://developers.cloudflare.com/pages/configuration/branch-build-controls/ . GitHub integration: https://developers.cloudflare.com/pages/configuration/git-integration/github-integration/ and https://developers.cloudflare.com/pages/get-started/git-integration/ . Known issues (acme-challenge, no Access policy on the domain, no wildcard): https://developers.cloudflare.com/pages/platform/known-issues/ . Build image: https://developers.cloudflare.com/pages/configuration/build-image/ . Terms: https://www.cloudflare.com/terms/
- Cloudflare direction of travel: the migration page says "Unlike Pages, Workers does not support any domain whose nameservers are not managed by Cloudflare." Pages is therefore the only Cloudflare product that fits our "CNAME only, zone stays with IT" constraint. I found no official Pages sunset notice (a third-party wiki claims one; unverified). https://developers.cloudflare.com/workers/static-assets/migration-guides/migrate-from-pages/
- Netlify: https://www.netlify.com/pricing/ ; commercial-use forum thread https://answers.netlify.com/t/can-we-use-netlify-free-plan-for-commercial-purposes/41545 ; external DNS https://docs.netlify.com/domains-https/custom-domains/configure-external-dns/
- Vercel domains: https://vercel.com/docs/domains/working-with-domains/add-a-domain . Changing only the A or CNAME at the current DNS provider "does not require moving the rest of your DNS records."
- Azure SWA custom domains: https://learn.microsoft.com/en-us/azure/static-web-apps/custom-domain
- GitHub plans and price: https://docs.github.com/en/get-started/learning-about-github/githubs-plans and https://github.com/pricing . Custom domain CNAME: https://docs.github.com/en/pages/configuring-a-custom-domain-for-your-github-pages-site/managing-a-custom-domain-for-your-github-pages-site

## Decision (proposed)
**Option A: Cloudflare Pages with the private GitHub repo, served at `https://mailai.dimscale.com/`.**
- Free, no card claimed by any page read (signup requirements not verified, Q-1).
- Private repo and org repo supported; preview per branch at a stable alias, so `dev` can be tested from Outlook without the user's PC; the user's `npm start` is then only for development.
- Subdomain by one `CNAME`; DIMSCALE's DNS zone stays where it is.
- The stable domain makes the provider replaceable by DNS alone (key design rule). If Cloudflare Pages is ever withdrawn, options C, D, B or a future provider fit the same record type.

Not chosen: B (the SaaS clause and no previews), D (Hobby bars commercial use; Pro is $20 per seat), C (works free only if the org-owned private repo condition is met; it appears under Pro; informal commercial answer), E (card), F (fails goals 1 and 2).

### Rules that come with the decision
1. The manifest, Entra redirect and `ALLOWED_ORIGINS` use `mailai.dimscale.com` only. Provider hostnames are used only for the test/preview manifest.
2. The preview/dev manifest is a **separate** add-in: its own `<Id>`, `DisplayName` "DIMSCALE Mail (teste)", and URLs on the dev branch alias (`https://dev.<project>.pages.dev/`). The user sideloads it; it is never sent to MMSI. Its host needs its own Entra redirect `brk-multihub://dev.<project>.pages.dev` and an `ALLOWED_ORIGINS` entry (Q-4: whether Entra accepts a wildcard here is not verified, so register one fixed alias).
3. Nothing sensitive is in the bundle (already true: no secrets on the client). Preview URLs are public by default; the static code is public to anyone who has the URL. Do not put real data in previews.
4. The Cloudflare build must run the same gates as `pages.yml` (typecheck, lint, tests, `office-addin-manifest validate`) inside the build command, so a broken `main` cannot publish.
5. Make the repo private and switch GitHub Pages off **only after** the cut-over is verified (plan step order).

## Related operational decision (not part of the hosting choice): Supabase plan
Today the project is on the Free plan (not verified in the dashboard, Q-8). On https://supabase.com/pricing: Free is "from $0/month" with "Free projects are paused after 1 week of inactivity" and no automatic backups; Pro is "from $25/month" with "Projects on paid plans are not paused for inactivity" and "Daily backups stored for 7 days". A paused project makes the add-in, the F-013 webhooks and the cron stop, so a commercial product with real users should not run on Free. Moving to Pro needs a payment card (the Azure objection may apply; Q-9) and a DPO/billing owner. The pricing page says nothing about commercial use on Free (Q-2). Separate decision, tracked in D21 of `status.md`; it does not block the hosting move.

## Consequences
- After the single final manifest upload, front-end releases are a `git push` (no MMSI). Only manifest changes (table above) need MMSI.
- One DNS record at DIMSCALE's DNS is needed once; the zone is not moved.
- ADR-0004's "shared `dimscale.github.io` origin" risk disappears (own origin).
- GitHub Pages stops being the host; the add-in breaks the moment Pages stops if the manifest still points there, hence the order in the plan.
- Docs to update when accepted (not done now): `docs/architecture/overview.md`, `docs/architecture/security.md` (hosting section), `docs/specs/F-011-github-pages-hosting.md` status note, ADR-0004 status line (Superseded by 0007), `docs/admin/2026-10-23-mmsi-request.md` (URL).
- New risks: Cloudflare Pages vs Workers direction (above); builds are no longer gated by GitHub Actions unless the checks are in the build command; Outlook caches the pane (asset names stay content-hashed, ADR-0004).

## Open questions (for the user, nothing assumed)
- **Q-1** Does Cloudflare account creation and the Pages Free plan need a payment card? Not stated on the pages read. The user rejected Azure for the card, so check before step 1.
- **Q-2** Commercial-use text: Cloudflare's Pages/Workers service-specific terms were not found, only the Self-Serve Subscription Agreement (no blanket ban). Netlify's agreement and Supabase's Free terms were not read. Someone accountable (user or legal) should read them before accepting.
- **Q-3** Is the Cloudflare certificate for a CNAME-onboarded custom domain free and automatic? The pages read describe ACME validation but not the price. Also: does `dimscale.com` have `CAA` records that would block issuance? Ask whoever manages DNS (not looked up).
- **Q-4** Does Entra accept a wildcard SPA redirect URI for NAA (`brk-multihub://*.pages.dev`)? Not verified; the plan registers one fixed alias.
- **Q-5** Node 24 (used in CI) on the Cloudflare build image: the page lists default 22.23.4 and `NODE_VERSION`; support for 24 was not verified.
- **Q-6** Is the GitHub Pages site public by default when published from a private repo on Team? The page does not say (only matters if option B is revisited).
- **Q-7** Does Outlook re-fetch icon images at the same URL after they change? Not documented in the pages read.
- **Q-8** Which Supabase plan is the project on today, and who pays Pro?
- **Q-9** Was a manifest already sent to MMSI or uploaded? If it points to `dimscale.github.io`, the cut-over needs one more upload (version bump, "Update" in the admin center) and Pages must stay on until the update has reached users.
- **Q-10** Hostname: `mailai.dimscale.com` is the name given in the request. Any naming rule at DIMSCALE IT?
- **Q-11** Who owns the Entra app registration edits (add the redirect URI): MMSI or the user?
- **Q-12** Other places that hold the GitHub Pages hostname: searched the code, only `webpack.config.js` has it. Supabase secrets and the Entra registration are not visible from the repo and must be checked in the plan steps.

## Out of scope
Changing the backend (Supabase stays, ADR-0003); the manifest content (separate feature work); Cloudflare Access/WAF; moving the DNS zone; custom headers/CSP (can be added later with `_headers`, https://developers.cloudflare.com/pages/configuration/serving-pages/); n8n, VPS or Azure.
