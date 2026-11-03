# Plan: ADR-0007 hosting migration (GitHub Pages public to Cloudflare Pages + private repo)

Status: **Draft, blocked until the user accepts [ADR-0007](../decisions/0007-commercial-hosting.md).** Nothing below is done. Owners: **user**, **orchestrator** (main session and its coders), **IT-MMSI** (tenant admin and whoever manages `dimscale.com` DNS).

## What can start before the DNS access
Steps 1 to 8 need **no** DNS access and **no** MMSI. DNS is first needed at step 9, and MMSI at step 12. The user can start today with the account (step 1) and answer the open questions.

## Steps (in order)

| # | Step | Owner | Notes / evidence to record |
|---|---|---|---|
| 0 | **Done 2026-11-03: ADR-0007 accepted by the user.** Q-9 answered: no manifest has been sent to MMSI yet (user, 2026-11-02). Accept ADR-0007 (or choose another option). Answer Q-1 (card?), Q-2 (terms read), Q-9 (was any manifest already sent to MMSI?). | user | Without acceptance nothing else starts (CLAUDE.md rule 8). |
| 1 | Create the Cloudflare account (no network or DNS change). Stop if a card is demanded (Q-1) and report. | user | Account owner and recovery email should be a DIMSCALE mailbox, not a personal one. |
| 2 | Create a Pages project from the private GitHub repo: install the "Cloudflare Workers and Pages" GitHub App on the `dimscale` org for this repo only ("Only select repositories"). Needs org owner or GitHub Apps Manager role. Proposed project name `dimscale-mail-ai` (so `dimscale-mail-ai.pages.dev`; availability not verified). The repo may be made private first or later (step 14); the connection works for both. | user (org owner) | https://developers.cloudflare.com/pages/configuration/git-integration/github-integration/ |
| 3 | Build settings: production branch `main`; build command runs the same gates as `pages.yml`: `npm ci --ignore-scripts && npm run typecheck && npm run lint && npm run test && npm run build && npx office-addin-manifest validate dist/manifest.xml`; output directory `dist`; env var `NODE_VERSION` (Q-5: 24 if supported, else 22 and run the checks). Preview branches: include `dev`, exclude `dependabot/*`. | user (dashboard) with orchestrator giving the exact values | https://developers.cloudflare.com/pages/configuration/build-image/ , https://developers.cloudflare.com/pages/configuration/branch-build-controls/ |
| 4 | Code change in `webpack.config.js`: `urlProd` is no longer hard-coded to GitHub Pages; the manifest base URL comes from an explicit env var and the existing guard stays (https, trailing `/`, not contoso) plus an **allow-list of two hosts** (`mailai.dimscale.com`, the dev alias). Two build targets: production manifest (`mailai.dimscale.com`, `<Id>` 05c2e1c9-..., current DisplayName) and test manifest (dev alias, **new `<Id>`**, DisplayName "DIMSCALE Mail (teste)"). The publicPath stays `auto`, so serving from `/` works. Tests for the guard. Do not touch the `main` pipeline until step 10. | orchestrator (coder, then review) | Docs-only today; this is the first code step. Record typecheck/lint/test/build output in this plan. |
| 5 | Supabase: add the dev alias origin (`https://dev.<project>.pages.dev`) and `https://mailai.dimscale.com` to the secret `ALLOWED_ORIGINS` **in addition** to the existing values (do not remove `https://dimscale.github.io` yet). | user or orchestrator (Supabase CLI/dashboard), user approves | `supabase secrets set ALLOWED_ORIGINS=...`; check the function still answers from the old origin. No log of secrets. |
| 6 | Entra: add SPA redirect URIs `brk-multihub://dev.<project>.pages.dev` and `brk-multihub://mailai.dimscale.com` to the app registration (client ID `d157a20c-...`). Keep `brk-multihub://dimscale.github.io` and `brk-multihub://localhost:3000`. | IT-MMSI (or the user if he owns the registration, Q-11) | NAA uses `brk-multihub://<host>`; wildcard acceptance unverified (Q-4). |
| 7 | Test via the preview URL: push `dev`; Cloudflare builds `https://dev.<project>.pages.dev/`; the user sideloads the **test manifest** (separate `<Id>`; the production add-in and test add-in coexist) and checks sign-in (NAA), a summary on a test email, reply form, tab Lista. No real mail until D5. | user | Evidence: build log OK, manifest validate OK, screenshot of the working pane, no CORS or redirect errors. |
| 8 | Ask whoever manages `dimscale.com` DNS for the record (text below). Before sending, in the Cloudflare project: **Custom domains, add `mailai.dimscale.com`** first (otherwise the CNAME returns 522). | user (sends); user does the dashboard step | https://developers.cloudflare.com/pages/configuration/custom-domains/ |
| 9 | IT creates the record. Verify with `nslookup -type=CNAME mailai.dimscale.com` and that Cloudflare shows the domain as active and the certificate as issued; open `https://mailai.dimscale.com/taskpane.html`. If it fails: check `CAA` records and the `/.well-known/acme-challenge/` path (known issues page, Q-3). | IT-MMSI creates; user verifies | https://developers.cloudflare.com/pages/platform/known-issues/ |
| 10 | Merge to `main` so the production build points to `mailai.dimscale.com`; Cloudflare deploys `main`. Keep `.github/workflows/pages.yml` and GitHub Pages **running** (do not delete yet). Check: production `dist/manifest.xml` has no `localhost`, no `dimscale.github.io`, no `pages.dev`; same checks as F-011. | orchestrator | The Pages site at the old URL stays online, so any manifest already uploaded (Q-9) still works. |
| 11 | Manifest version bump (`<Version>` greater than the last one sent; the production `<Id>` unchanged) and review. | orchestrator (coder, reviewer) | Rule from Microsoft: a manifest change requires a higher version. https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish |
| 12 | **The one MMSI upload**: the final manifest (Integrated apps; "Update" if one already exists). Combine with the pending checklist in `docs/admin/2026-10-23-mmsi-request.md` (update the URL there first). Do not send a manifest pointing to `github.io`. | user sends, IT-MMSI uploads | Microsoft: LOB add-in updates "the next time a user starts the relevant Microsoft 365 productivity application"; appearance on the ribbon "can take 24-72 hours". https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins |
| 13 | Cut-over verification with the production add-in (users on the new manifest): pane loads from `mailai.dimscale.com`, sign-in, summary on a test email. Wait a grace period (proposed 14 days, user decides) and confirm no client still uses the old origin (Cloudflare analytics or Supabase function metadata only; no email content logged). | user | Live test claimed only if performed. |
| 14 | **Only then:** (a) delete `pages.yml`; (b) disable GitHub Pages in the repo settings; (c) make the repo private; (d) remove `https://dimscale.github.io` from `ALLOWED_ORIGINS` and `brk-multihub://dimscale.github.io` from Entra. | user (GitHub, Supabase) / IT-MMSI (Entra) | Order (a) to (d) matters, see the reason below. |
| 15 | Docs: accept ADR-0007 in `status.md`, mark ADR-0004 Superseded, update `overview.md`, `security.md` (hosting section), F-011 note, MMSI request. | orchestrator | Record D7 as superseded. |
| 16 | Separate: Supabase plan decision (tracked in D21): check the plan, move to Pro if the user agrees, before real users. | user | Not on the critical path of the migration. |

### Why the repo goes private and Pages is turned off last
GitHub Pages on a free organization needs a public repo ("If the account that owns the repository uses GitHub Free or GitHub Free for organizations, the repository must be public"). Making the repo private first would stop the Pages deployment and could take the old URL down. If MMSI already uploaded a manifest that points to the old URL (Q-9), the add-in would **break at that moment** for every user, and until the new manifest has propagated (24 to 72 hours or more). Doing the cut-over first, waiting the grace period, and only then privatizing means there is never a moment without a working host. It also keeps the rollback below possible.

### Rollback
- **Before step 14:** if `mailai.dimscale.com` or Cloudflare misbehaves, change the DNS `CNAME` target back to `dimscale.github.io` (Pages still on, repo still public). No manifest change if `mailai.dimscale.com` is the manifest host, because GitHub Pages can also serve a custom domain (needs the domain set in the repo Pages settings; ADR-0004 option B; https://docs.github.com/en/pages/configuring-a-custom-domain-for-your-github-pages-site/managing-a-custom-domain-for-your-github-pages-site). Note the old project-site path `/DIMSCALE-EMAIL-AI/` would become `/` under a custom domain, so that rollback needs a quick test before it is relied on.
- **After step 14:** the old host no longer exists; rollback means pointing the CNAME at another provider from ADR-0007 (Netlify Pro, Vercel Pro). DNS-only, still no new manifest, thanks to the stable domain. Keep the Cloudflare project until then.
- **Function side:** `ALLOWED_ORIGINS` and the Entra redirect are only added in steps 5 and 6 and removed in step 14, so every earlier state accepts both origins.

## Text for the DNS administrator (PT-PT, ready to send)

Assunto: Pedido de registo DNS (CNAME) para o subdomínio mailai.dimscale.com

Olá,

Preciso que seja criado **um único registo CNAME** na zona DNS `dimscale.com`, para o complemento do Outlook "DIMSCALE Mail" ficar acessível num endereço da DIMSCALE:

- **Tipo:** CNAME
- **Nome / Host:** `mailai` (isto é, `mailai.dimscale.com`)
- **Destino / Valor:** `dimscale-mail-ai.pages.dev` (confirmar o valor exato comigo antes de criar, se eu ainda não o tiver confirmado)
- **TTL:** 3600 (ou o valor por omissão)
- **Proxy:** não aplicável (DNS simples, sem alterar nameservers)

Notas:
1. **Não é preciso mover nem delegar a zona `dimscale.com`**; os nameservers e todos os outros registos (incluindo e-mail/MX) ficam como estão. É só este CNAME.
2. Se existirem registos **CAA** na zona, peço que me avisem (podem impedir a emissão do certificado HTTPS). Não é preciso criar nenhum registo TXT.
3. Evitar registos wildcard (`*.dimscale.com`) para este fim.
4. Depois de criado, peço que confirmem com `nslookup -type=CNAME mailai.dimscale.com`.

O endereço fica definitivo: se mudarmos de fornecedor de alojamento no futuro, será só alterar o destino deste CNAME, sem novo pedido ao MMSI para o manifesto do complemento.

Obrigado.

(Nota interna, apagar antes de enviar: o nome `dimscale-mail-ai` é proposto no passo 2; só enviar depois de o projeto existir e de o domínio ter sido adicionado no painel do Cloudflare, passo 8. Se o nome do projeto for diferente, trocar o valor.)

## Order of the user's next actions given the DNS situation (2026-11-03)
1. Today, without DNS: accept the ADR (step 0), create the Cloudflare account (1), connect the repo (2), settings (3). The coder can do step 4 in parallel; steps 5 to 7 give a working dev test at a `pages.dev` address.
2. When the user gets DNS access (later today), the user can add the record directly; **if the user must ask MMSI, send the text above** (it asks for one record only). If the same person at MMSI will also receive the manifest, send the manifest request (step 12) together, but only after step 9 passes.

## Sources
Same as ADR-0007, plus: Microsoft Learn on updates https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins and https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish ; GitHub Pages creation rule https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site ; Cloudflare pages above.
