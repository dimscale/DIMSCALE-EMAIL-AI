# Security and Privacy

## Authentication and authorization
- **Users**: Microsoft Entra ID through MSAL.js nested app authentication (NAA), as proposed in [ADR-0002](../decisions/0002-auth-nested-app-auth.md). App registration type SPA, with redirect `brk-multihub://<add-in-origin>` (origin only, no path). Registered redirects: `brk-multihub://localhost:3000` (dev) and `brk-multihub://dimscale.github.io` (production, F-011). Check `Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")` and show a clear "unsupported client" state when it returns false.
- **Graph**: delegated scopes only, requested incrementally per feature. `Mail.Send` is never requested. No application (app-only) permissions.
- **AI backend** ([F-002](../specs/F-002-authenticated-backend.md)): without token validation the backend would be an open, paid OpenAI proxy, so every request is authenticated by the function itself.
  - **Flow**: task pane → MSAL.js NAA → Entra access token for `api://<client-id>/access_as_user` (scopes `access_as_user` and `profile`, no Graph scope) → `Authorization: Bearer` on a `fetch` to the Supabase Edge Function. One single-tenant app registration is both SPA client and API, with `requestedAccessTokenVersion` = 2.
  - **Validation (in the function, before any other logic)**: signature against the tenant's v2.0 JWKS (selected by `kid`, remote key set with cache and refetch on unknown `kid`); `alg` RS256 only; `iss` exactly `https://login.microsoftonline.com/<tenant-id>/v2.0`; `ver` = `2.0`; `aud` = the client ID; `azp` = the client ID; `scp` contains `access_as_user`; `exp`/`nbf` with 60 s clock tolerance. Anything else gets 401.
  - **`verify_jwt = false`** for the function (`supabase/config.toml`): Supabase's platform check is for Supabase-issued JWTs and keys, not Entra tokens. With it off anyone can invoke the function, so the handler's own validation is the only gate.
  - **CORS**: exact allow-list from `ALLOWED_ORIGINS`, never `*`. Production value: `https://localhost:3000,https://dimscale.github.io` (localhost for dev, the Pages origin for production, origin only, no path). `Vary: Origin`, allowed header `authorization`, methods `GET, OPTIONS`. Non-allowed origins get no `Access-Control-Allow-Origin`.
  - **Responses and logs**: success returns only `{"ok":true,"name":<name or null>}`. Failure returns a generic 401 that does not say which check failed. Logs hold request id, status, a fixed reason code (for example `aud_mismatch`) and duration, never tokens, claims or the name.
  - **Who can sign in (F-002/AC-21)**: access is limited to one DIMSCALE Entra group through the enterprise app's **Assignment required = Yes** with only that group assigned (needs Entra ID P1 or P2). Users outside it get no token. This is enforced by Entra, not by function code. Source: https://learn.microsoft.com/entra/identity/enterprise-apps/assign-user-or-group-access-portal
  - **Data**: F-002 sends no email content. Supabase receives only the access token, which carries user identifiers (`oid`, `tid`, `name`, `preferred_username`), so it is a processor of identity data (confirm project region for GDPR, decision D5).
- **Tokens**: MSAL manages its own cache. Our code never logs tokens, never sends Graph tokens to the backend, and never stores tokens itself.

## Secrets
- The OpenAI API key lives only in the backend secret store (for example Supabase function secrets).
- Nothing secret goes in the frontend bundle, the manifest, browser storage or git. `.env*` files are git-ignored. A code reviewer runs a secret scan on every review.
- **Not secret (F-002)**: Entra client ID, tenant ID, API scope, function URL, Supabase project ref and `ALLOWED_ORIGINS`. They may be committed or set with `supabase secrets set` (Supabase's env mechanism), and are configured per environment. The Supabase publishable key is browser-safe and not used.
- **Secret**: the Supabase CLI access token (developer machine only), the database password and secret/service-role keys (unused in F-002), and later the OpenAI key. No client secret or certificate exists on the app registration. `supabase/functions/.env` is git-ignored.

## Email content handling
- **Untrusted input**: email subject and body are data. The backend places them in a clearly delimited user message, and the system prompt states that the content must not be followed as instructions.
- **No actions from model output**: the model returns structured JSON (validated against a schema). The UI shows it and only the user can act on it. The model has no tools.
- **Rendering**: AI and email text is rendered as React text nodes. `dangerouslySetInnerHTML` and `innerHTML` are forbidden for this content. Reply HTML passed to `displayReplyFormAsync` is built from escaped text.
- **Minimization**: send only what the feature needs: plain-text body (not HTML), subject, sender display name and domain, truncated to the limit in each spec. No attachments, no recipient lists unless a spec justifies it.
- **Logging**: no email content, addresses or tokens in browser console logs, backend logs or error reports. Log request ids, durations and status codes only.

## Third-party processing (OpenAI)
- Email content leaves Microsoft 365 when it is sent to the OpenAI API. Before production, **verify OpenAI's current API data-usage and retention policy** and document it here. Also confirm with DIMSCALE that this is acceptable under its GDPR obligations (a DPA with OpenAI, and the region).
- Show users a one-line notice in the panel that content is processed by an external AI service.

## Hosting (GitHub Pages, [ADR-0004](../decisions/0004-taskpane-hosting.md), [F-011](../specs/F-011-github-pages-hosting.md))
- **Data flow**: push to `main` → GitHub Actions (typecheck, lint, tests, production build, manifest validation) → GitHub Pages → Outlook loads the task pane from `https://dimscale.github.io/DIMSCALE-EMAIL-AI/`. The page then calls Entra (MSAL NAA) and the Supabase function as described above. Pages serves static files only and sees no email content.
- **Public artifacts**: everything under the Pages URL is public, including the production `manifest.xml`. It holds only non-secret identifiers (add-in Id, URLs, `ReadItem`). Authorization is the Entra token validated by the function, not hiding the page.
- **No production source maps** are published. Dev builds keep them.
- **Shared origin risk (accepted, ADR-0004)**: `dimscale.github.io` is one browser origin for every Pages site in the `dimscale` org. They share `localStorage`/`sessionStorage` (including the MSAL cache) and are all covered by the Entra redirect and `ALLOWED_ORIGINS` entries. Mitigation: only org members can publish Pages, a token still needs the user's Entra sign-in, and the rule **no other Pages site under the `dimscale` org origin without review**. Moving to a custom domain (ADR-0004 option B) would remove the sharing.
- **CI supply chain (deploy workflow)**: every action pinned to a full commit SHA (version in a comment); top-level `permissions: contents: read`, and only the deploy job adds `pages: write` and `id-token: write`; the deploy job runs only on `main` (never on pull requests) and needs the check job; no `pull_request_target`; no repository secrets (the Pages actions use the automatic `GITHUB_TOKEN`); `npm ci --ignore-scripts`. SHA pins get no automatic Dependabot alerts, so review them when touching the workflow.

## Supply chain
- Dev-only tooling (`office-addin-debugging`, `office-addin-dev-certs`, and others) pulls packages with install scripts: `keytar`, `@azure/msal-node-extensions`, `@microsoft/m365agentstoolkit-cli`, `core-js`. None of them ship in the bundle. CI uses `npm ci --ignore-scripts` (see Hosting).

## Classification safety
- The "potentially suspicious" label is advisory, shown with its reasons and a "may be wrong" notice. The add-in never reports phishing, moves mail to junk or blocks senders on its own.
