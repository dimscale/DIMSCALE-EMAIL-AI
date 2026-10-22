# Plan F-002: Authenticated AI backend skeleton

- **Spec:** [F-002](../specs/F-002-authenticated-backend.md). **Status:** In progress (spec approved 2026-10-22).
- **Blocked by:** D4 (app registration), ADR-0002 acceptance, spec approval. Steps 0 and 1 can start once the spec is approved; step 0 needs IT.

## Configuration values

| Value | Example / format | Secret? | Where it lives |
|---|---|---|---|
| Entra tenant ID | GUID | **No** | Task pane config file (committed), function env `ENTRA_TENANT_ID` |
| Entra application (client) ID | GUID | **No** | Task pane config file (committed), function env `ENTRA_CLIENT_ID` |
| API scope | `api://<client-id>/access_as_user` | **No** | Derived from the client ID in the task pane config |
| Function URL | `https://<project-ref>.supabase.co/functions/v1/whoami` | **No** | Task pane config file (committed) |
| Supabase project ref / URL | `<project-ref>`, `https://<project-ref>.supabase.co` | **No** | Task pane config, `supabase link` |
| Allowed origins | `https://localhost:3000` | **No** | Function env `ALLOWED_ORIGINS` |
| Supabase publishable (anon) key | `sb_publishable_...` | **No** (browser-safe per Supabase), but **not used** unless open question 2 forces it | — |
| Supabase CLI access token (from `supabase login`) | — | **Yes** | Only on the developer's machine, managed by the CLI. Never in the repo. |
| Supabase database password, secret / service-role keys | — | **Yes** | Not used by F-002. Never in the repo or the function env we set. |
| OpenAI API key | — | **Yes** | Not part of F-002. |

F-002 has **no application secret**. The function env values are set with `supabase secrets set` only because that is Supabase's env mechanism ([Secrets](https://supabase.com/docs/guides/functions/secrets)); they are not secret. The local file `supabase/functions/.env` is still git-ignored (the root `.gitignore` already matches `.env` at any depth).

## Manual prerequisites (user / IT)

### P1. Entra app registration (a DIMSCALE admin, at least Cloud Application Administrator)
Sources: [Enable NAA](https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in), [Expose a web API](https://learn.microsoft.com/entra/identity-platform/quickstart-configure-app-expose-web-apis), [API permissions](https://learn.microsoft.com/entra/identity-platform/quickstart-configure-app-access-web-apis), [Access tokens](https://learn.microsoft.com/entra/identity-platform/access-tokens).
1. Entra admin center > App registrations > New registration.
   - Name: `DIMSCALE Mail AI (dev)`.
   - Supported account types: **Accounts in this organizational directory only (DIMSCALE only - Single tenant)**. (The Microsoft walkthrough uses multitenant + personal; we deliberately don't.)
   - Redirect URI: platform **Single-page application (SPA)**, value `brk-multihub://localhost:3000`.
2. Copy the **Application (client) ID** and **Directory (tenant) ID** to the orchestrator (not secret).
3. **Manifest** (Microsoft Graph App Manifest editor): set `api.requestedAccessTokenVersion` to `2`. Save.
4. **Expose an API**: Application ID URI > Add > accept `api://<client-id>` > Save. Add a scope:
   - Scope name `access_as_user`; Who can consent **Admins and users**;
   - Admin consent display name "Access DIMSCALE Mail AI backend as the signed-in user"; description "Lets the DIMSCALE Mail AI add-in call its backend on behalf of the signed-in user.";
   - User consent display name and description: same text; State **Enabled**.
5. Expose an API > **Authorized client applications** > Add a client application > the same client ID > tick `access_as_user`.
6. **API permissions**: remove the auto-added Microsoft Graph `User.Read` (F-002 needs no Graph). Add a permission > **My APIs** > `DIMSCALE Mail AI (dev)` > Delegated > `access_as_user`. Then **Grant admin consent for DIMSCALE**.
7. **Restrict to the group** (user decision 2026-10-22): Enterprise apps > `DIMSCALE Mail AI (dev)` > Properties > **Assignment required? = Yes**. Then Users and groups > Add user/group > the chosen group > Assign. Group assignment needs Entra ID **P1 or P2**; without it, tell the orchestrator, because a code-side group or role check would be needed instead. Apps that require assignment need admin consent ([source](https://learn.microsoft.com/entra/identity/enterprise-apps/assign-user-or-group-access-portal)).
8. Do **not** create a client secret or certificate. Do not add Graph `Mail.*` permissions.
9. Open `https://login.microsoftonline.com/<tenant-id>/v2.0/.well-known/openid-configuration` in a browser and send the orchestrator the `issuer` and `jwks_uri` values (expected `https://login.microsoftonline.com/<tenant-id>/v2.0` and `https://login.microsoftonline.com/<tenant-id>/discovery/v2.0/keys`). They are public.

### P2. Supabase CLI and project link (the user)
Sources: [Quickstart](https://supabase.com/docs/guides/functions/quickstart), [Deploy](https://supabase.com/docs/guides/functions/deploy).
1. Install the Supabase CLI (see the Supabase CLI install guide) and **Deno** (needed for `deno test`). For `supabase functions serve`, Docker or a Docker-compatible runtime must run.
2. `supabase login` (stores the CLI token locally; secret, never committed).
3. From the repo root, after step 2 below has created `supabase/`: `supabase link --project-ref <project-ref>`.
4. Tell the orchestrator the project ref and region (spec open question 5).

### P3. Outlook build
**Done 2026-10-22:** classic Outlook Version 2609 (Build 20430.20182, Current Channel), which meets the requirement. Check: File > Office Account > About Outlook. It must be Version 2409 (Build 18025.20000) or later ([requirement set](https://learn.microsoft.com/javascript/api/requirement-sets/common/nested-app-auth-requirement-sets)). Record it in Evidence.

## Steps

| # | Step | ACs | Agent |
|---|---|---|---|
| 0 | User/IT complete P1, P2, P3 and send the non-secret values. Orchestrator records them (IDs only) in this plan. | prerequisite | user / IT + orchestrator |
| 1 | Update `docs/architecture/security.md` and `overview.md` with the F-002 data flow (spec "Security notes"). | — (step 3 of the SDD lifecycle) | requirements-architect |
| 2 | `supabase init` at the repo root (creates `supabase/config.toml`). Add `[functions.whoami] verify_jwt = false`. Create `supabase/functions/whoami/` with `index.ts` and `deno.json` (pin `npm:jose@<exact version>`). Exclude `supabase/` from the root `tsconfig.json`, the lint config and vitest's include, so the Node toolchain ignores Deno code. | AC-15, AC-20 | coder |
| 3 | Function: a pure `validateToken(token, keySet, config)` using `jose.jwtVerify` with `algorithms: ["RS256"]`, exact `issuer`, exact `audience`, `clockTolerance: 60`, then checks `ver === "2.0"`, `azp === clientId`, `scp` contains `access_as_user`. A CORS helper driven by `ALLOWED_ORIGINS`. The handler: `OPTIONS` → CORS; non-GET → 405; else parse `Authorization`, validate, return 200/401. The JWKS (`createRemoteJWKSet` on the P1.8 `jwks_uri`) is created once at module scope. Logging: request id, status, reason code, duration only. Run once with `supabase functions serve whoami` and `curl` without any `apikey` header to answer spec open question 2; record the result. | AC-8 to AC-16 | coder |
| 4 | `deno test` for step 3, with a test RSA key pair generated in the test and a local JWKS (`createLocalJWKSet`). One case per rejection in AC-8 to AC-12, the success shape in AC-13, the CORS cases in AC-14. No network. | AC-8 to AC-14 | coder |
| 5 | Task pane: add `@azure/msal-browser` (exact version, record it). A config module with tenant ID, client ID and function URL. An auth module: NAA check, `createNestablePublicClientApplication` with the single-tenant authority and no `cacheLocation` override (pending open question 3), `acquireTokenSilent` then a `signIn()` that calls `acquireTokenPopup`. A backend client: `fetch` GET with the bearer header and a 10 s abort. | AC-1 to AC-6 | coder |
| 6 | `BackendStatus` component rendered under the F-001 `MessagePanel`, with the states from the spec "Behavior" section, strings in `strings.ts`. Vitest with mocked `Office.context.requirements`, mocked MSAL and mocked `fetch`. Confirm the F-001 tests still pass unchanged. | AC-1, AC-4, AC-5, AC-7 | coder |
| 7 | Set function env with `supabase secrets set ENTRA_TENANT_ID=... ENTRA_CLIENT_ID=... ALLOWED_ORIGINS=https://localhost:3000` and `supabase functions deploy whoami`. Confirm `verify_jwt = false` took effect (a request with no header reaches the handler and gets the function's own 401 JSON). | AC-15, AC-19 (no-token case) | user (CLI is logged in as the user) + orchestrator |
| 8 | Run typecheck, lint, vitest, `npm run build:dev`, `npm run validate`, `deno test`; grep `src/` and `supabase/` for `console.` with token/claim data, `innerHTML`, Graph URLs, `Mail.`; secret scan. Record output below. | AC-3, AC-16, AC-17, AC-20 | qa-tester |
| 9 | Independent review (security-sensitive: pass `model: opus`). | AC-2, AC-6, AC-14, AC-15, AC-16 | code-reviewer |
| 10 | Live: sideload in classic Outlook (P3 version), open a message, record the status line and the console (no CORS error). Then the negative calls in AC-19: no token, a Graph token (from Graph Explorer), an expired token. Record dates, client, results. | AC-18, AC-19 | user + orchestrator |
| 11 | Close: spec status, `docs/status.md` (F-002 Done, D4 resolved), ADR-0002 status if the user accepts it. | — | orchestrator |

## Evidence
_Empty until implementation._
