# F-002: Authenticated AI backend skeleton

- **Status:** Draft (awaiting user approval)
- **Tier:** Backend + user authentication. No Graph scope, no email content, no OpenAI.
- **Depends on:** F-001 (Done), ADR-0002 (Proposed, must be accepted before implementation), ADR-0003 (Accepted), decision D4 (Entra tenant and app registration).
- **Plan:** [docs/plans/F-002-authenticated-backend.md](../plans/F-002-authenticated-backend.md)

## Goal
Prove the secure path from the task pane to the AI backend end to end, before any email content or OpenAI key exists:
the task pane gets an Entra ID access token for the add-in's own API through MSAL.js nested app authentication (NAA), calls one Supabase Edge Function with it, and the function validates the token itself and answers with minimal, non-sensitive information.

## Verified platform facts (what the design relies on)
| Fact | Source |
|---|---|
| NAA uses `createNestablePublicClientApplication` from `@azure/msal-browser`. Pattern: `acquireTokenSilent`, and on `InteractionRequiredAuthError` call `acquireTokenPopup`. `acquireTokenRedirect`, `loginRedirect` and `logout*` throw under NAA. | [Enable NAA](https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in) |
| The app registration needs an SPA redirect `brk-multihub://<origin>` (origin only, e.g. `brk-multihub://localhost:3000`). Outlook does not need the extra task pane redirect (that one is for Word, Excel and PowerPoint on the web). | same |
| A token request must include at least one scope other than `openid`, `profile`, `email`, `offline_access`. | same |
| NAA works for a custom API scope, not only Graph: the official OfficeDev sample requests `api://<clientId>/Todolist.Read` with NAA and sends the token to its own web API. One app registration is both client and API: Application ID URI `api://{clientId}`, the client's own ID added under "Authorized client applications", `requestedAccessTokenVersion` = 2. | [Outlook-Add-in-SSO-NAA-Identity README](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA-Identity), [authConfig.ts](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/auth/Outlook-Add-in-SSO-NAA-Identity/SPA/src/taskpane/authConfig.ts) |
| Check NAA at runtime with `Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")`. It cannot be declared in an Outlook manifest. Minimum Outlook builds: web supported; Windows M365 Version 2409 (Build 18025.20000); retail perpetual 2501; LTSC 2408; Mac 16.89. Not supported in Outlook.com or Gmail mailboxes. | [NestedAppAuth requirement set](https://learn.microsoft.com/javascript/api/requirement-sets/common/nested-app-auth-requirement-sets) |
| Single-tenant MSAL authority: `https://login.microsoftonline.com/<tenant-id>/` (it must match the registration's supported account types). | [MSAL client configuration](https://learn.microsoft.com/entra/identity-platform/msal-client-application-configuration) |
| Token version is chosen by the **resource** through `requestedAccessTokenVersion` (`null` or `1` gives v1.0, `2` gives v2.0), whatever endpoint the client uses. Entra-only apps default to v1.0. | [Access tokens](https://learn.microsoft.com/entra/identity-platform/access-tokens) |
| v2.0 token: `aud` is always the API's client ID (GUID); `iss` ends in `/v2.0`, i.e. `https://login.microsoftonline.com/<tenant-id>/v2.0`; `ver` = `"2.0"`; client in `azp`; scopes in `scp`; `oid`, `tid`, `name` need the `profile` scope; `name` is display only. | [Access token claims](https://learn.microsoft.com/entra/identity-platform/access-token-claims-reference) |
| Validation: signature against the keys from the OIDC metadata `jwks_uri` (select by `kid`, keys rotate, refresh about every 24 h), exact issuer match for a tenant-specific metadata endpoint, `aud` must match, lifetime. A v2.0 token must be validated with the v2.0 metadata `https://login.microsoftonline.com/<tenant-id>/v2.0/.well-known/openid-configuration`. Default token lifetime 60 to 90 minutes. | [Access tokens, Validate tokens](https://learn.microsoft.com/entra/identity-platform/access-tokens#validate-tokens) |
| App ID URI `api://<appId>` is a secure default pattern. Apps with `requestedAccessTokenVersion` = 2 are exempt from the identifier-URI restriction policy. | [Identifier URI restrictions](https://learn.microsoft.com/entra/identity-platform/identifier-uri-restrictions) |
| Expose an API: scope full string = App ID URI + `/` + scope name; "Who can consent: Admins and users"; pre-authorizing a client suppresses user consent prompts. | [Expose a web API](https://learn.microsoft.com/entra/identity-platform/quickstart-configure-app-expose-web-apis) |
| "Grant admin consent for {tenant}" removes per-user consent prompts. Portal-registered apps get Graph `User.Read` added automatically. | [Configure API permissions](https://learn.microsoft.com/entra/identity-platform/quickstart-configure-app-access-web-apis) |
| Supabase's platform JWT check (`verify_jwt`, default `true`) accepts Supabase-issued JWTs and Supabase keys; it does not mention third-party tokens. Turn it off per function with `[functions.<name>] verify_jwt = false` in `supabase/config.toml`; then "anyone" can invoke the function, so the handler must authenticate. | [Function configuration](https://supabase.com/docs/guides/functions/function-configuration), [Authorization headers](https://supabase.com/docs/guides/functions/auth-headers), [Securing functions](https://supabase.com/docs/guides/functions/auth) |
| Browser calls need the function to answer the CORS preflight (`OPTIONS`) itself; Supabase's examples use `Access-Control-Allow-Origin: *`. | [Supabase CORS](https://supabase.com/docs/guides/functions/cors) |
| Function URL: `https://<project-ref>.supabase.co/functions/v1/<name>`. Secrets/env: `Deno.env.get`, `supabase secrets set`, local `supabase/functions/.env`; names cannot start with `SUPABASE_`. | [Deploy](https://supabase.com/docs/guides/functions/deploy), [Secrets](https://supabase.com/docs/guides/functions/secrets) |
| Supabase's own docs verify JWTs with the `jose` library (`createRemoteJWKSet` + `jwtVerify`). npm packages are imported with `npm:` specifiers from a per-function `deno.json`. | [Supabase JWTs](https://supabase.com/docs/guides/auth/jwts), [Dependencies](https://supabase.com/docs/guides/functions/dependencies) |
| `AppDomains` lists domains whose **pages** the add-in loads and trusts (IFrames, navigation). It is not documented as a requirement for `fetch`/XHR; cross-origin calls are governed by the webview's same-origin policy, i.e. CORS on the server. | [AppDomains](https://learn.microsoft.com/javascript/api/manifest/appdomains), [Same-origin policy in add-ins](https://learn.microsoft.com/office/dev/add-ins/develop/addressing-same-origin-policy-limitations) |

## Design decisions in this spec
- **One app registration**, single tenant (DIMSCALE), acting as both SPA client and API, as in the OfficeDev sample. Application ID URI `api://<client-id>`, one delegated scope `access_as_user`, `requestedAccessTokenVersion` = 2.
- **Token request scopes:** `["api://<client-id>/access_as_user", "profile"]`. `profile` is listed explicitly because Entra documents it as the condition for `oid`, `tid` and `name`; the NAA page allows it next to a resource scope. No Graph scope.
- **Validation on v2.0 tokens only:** issuer `https://login.microsoftonline.com/<tenant-id>/v2.0` (exact), audience `<client-id>` (exact), `ver` = `2.0`, `azp` = `<client-id>`, `scp` contains `access_as_user`, `alg` RS256 only, `exp`/`nbf` with at most 60 s clock tolerance. JWKS from the tenant's v2.0 metadata (`https://login.microsoftonline.com/<tenant-id>/discovery/v2.0/keys`, to be confirmed against the metadata document in the plan).
- **Library: `jose`**, imported as `npm:jose@<exact version>` in the function's `deno.json`. Why: Supabase's docs use it for exactly this (remote JWKS + verify); it runs on Deno with zero dependencies; `createRemoteJWKSet` caches keys and refetches on an unknown `kid`, which covers Entra key rotation; `jwtVerify` checks signature, `iss`, `aud`, `exp`, `nbf` and the algorithm allow-list in one call. The alternative (hand-written RS256 verification with WebCrypto) is more code and more risk.
- **Function `whoami`**, `GET` only. Response on success: `200 {"ok": true, "name": "<name claim or null>"}`. On any token failure: `401 {"ok": false, "error": "unauthorized"}`, without saying which check failed. The reason is logged server side as a fixed code (e.g. `aud_mismatch`), never with token or claim values.
- **CORS:** an allow-list from the `ALLOWED_ORIGINS` env value (exact match, no wildcard), `Vary: Origin`, allowed headers `authorization` only, methods `GET, OPTIONS`.
- **UI:** silent token on panel load. If interaction is needed, show a "Sign in" button and call `acquireTokenPopup` only from that click (a user gesture, so the popup is less likely to be blocked). The F-001 message context keeps working whatever the backend state is.

## Behavior
Below the F-001 message context, the panel shows one backend status line:
- while working: "Connecting to DIMSCALE backend…"
- success: "Backend connected as {name}", or "Backend connected" when the function returns no name
- unsupported client: "This Outlook version doesn't support sign-in for DIMSCALE Mail AI. Update Outlook or use Outlook on the web."
- sign-in needed: "Sign in to connect to the DIMSCALE backend" with a "Sign in" button
- error: a short message for the error class (sign-in failed, backend rejected the sign-in, backend unreachable), the MSAL `errorCode` or HTTP status when there is one, and a "Retry" button

## Acceptance criteria
| ID | Criterion | Verify by |
|---|---|---|
| F-002/AC-1 | When `Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")` is false, the panel shows the unsupported-client line, MSAL is not initialized, no backend call is made, and the F-001 message context still renders. | Unit test (mocked Office.js) |
| F-002/AC-2 | MSAL is created with `createNestablePublicClientApplication`, `clientId` = the configured client ID and `authority` = `https://login.microsoftonline.com/<tenant-id>` (never `common` or `organizations`). | Unit test (mocked MSAL) and review |
| F-002/AC-3 | The token request uses exactly the scopes `api://<client-id>/access_as_user` and `profile`. No Graph scope appears anywhere in `src/`. | Unit test and grep |
| F-002/AC-4 | `acquireTokenSilent` is tried first. On `InteractionRequiredAuthError` the panel shows the "Sign in" button; `acquireTokenPopup` is called only from its click. On any other silent error the panel shows the error line with the MSAL `errorCode`. | Unit test (mocked MSAL) |
| F-002/AC-5 | If `acquireTokenPopup` fails (popup blocked, user closed it, consent denied or admin approval required), the panel shows "Sign-in failed" with the MSAL `errorCode` and a Retry button, and makes no backend call. | Unit test (mocked MSAL); one live case if it can be provoked |
| F-002/AC-6 | With a token, the task pane calls `GET <function-url>` with exactly one auth header, `Authorization: Bearer <token>`, no body, and no email field (subject, sender, body, item id) in the URL or headers. The request is aborted after 10 s. | Unit test (mocked `fetch`) and review |
| F-002/AC-7 | Response mapping: 200 with `ok: true` → "Backend connected as {name}" (or without name); 401 → "Backend rejected the sign-in (401)"; other status, network error, CORS failure or timeout → "Backend unreachable" with Retry. The name is rendered as a React text node. | Unit test |
| F-002/AC-8 | The function rejects with 401 `{"ok":false,"error":"unauthorized"}` when the `Authorization` header is missing, is not `Bearer <token>`, or the token is malformed. | Local function test (`deno test`) |
| F-002/AC-9 | The function rejects with 401 a token whose signature does not verify against the JWKS, whose `alg` is not RS256 (including `none` and HS256), or whose `kid` is unknown after one JWKS refresh. | Local function test with a test key pair and a local JWKS |
| F-002/AC-10 | The function rejects with 401 a token whose `iss` is not exactly `https://login.microsoftonline.com/<tenant-id>/v2.0` (other tenant, v1 `https://sts.windows.net/<tid>/`, trailing-slash variants), or whose `ver` is not `2.0`. | Local function test |
| F-002/AC-11 | The function rejects with 401 a token whose `aud` is not exactly the client ID (including the App ID URI form, the Graph audience and another app's ID), whose `azp` is not the client ID, or whose `scp` does not contain `access_as_user`. | Local function test |
| F-002/AC-12 | The function rejects with 401 a token that is expired or not yet valid beyond a 60 s tolerance, and accepts one inside it. | Local function test |
| F-002/AC-13 | A valid token returns 200 `{"ok":true,"name":...}`. The body contains no other claim (no `oid`, `tid`, `preferred_username`, email or token). | Local function test, then live |
| F-002/AC-14 | CORS: `OPTIONS` from an allowed origin returns 204 with `Access-Control-Allow-Origin: <that origin>` (never `*`), `Access-Control-Allow-Headers: authorization`, `Access-Control-Allow-Methods: GET, OPTIONS`, `Vary: Origin`, without needing a token. Responses to a non-allowed origin carry no `Access-Control-Allow-Origin`. 200 and 401 responses to allowed origins carry the CORS headers so the browser can read them. Methods other than `GET`/`OPTIONS` return 405. | Local function test; live preflight via the task pane |
| F-002/AC-15 | `supabase/config.toml` sets `verify_jwt = false` for `whoami` only, and the function's own validation runs before any other logic on every non-`OPTIONS` request. | File inspection and review |
| F-002/AC-16 | No secret in the repo or bundle: client ID, tenant ID, function URL and allowed origins are the only new config values, all non-secret. `supabase/functions/.env` is git-ignored. No MSAL token, `Authorization` header, claim value or name is passed to `console.*` in the task pane or the function. The function logs only a request id, status, fixed reason code and duration. | Grep, secret scan, review |
| F-002/AC-17 | The manifest is unchanged: `ReadItem` only, no new permission, no `AppDomains` entry (unless AC-18 shows one is required, which then needs a spec change). | Manifest inspection, `npm run validate` |
| F-002/AC-18 | Live, classic Outlook on Windows (Version 2409 or later): opening the pane on a message shows "Backend connected as {the signed-in user's name}" against the deployed function, with no console CORS error. | Live sideload, recorded in the plan |
| F-002/AC-19 | Live negative: calling the deployed function with no token, and with a Microsoft Graph token or an expired token, returns 401. | Live `curl`/script against the deployed function, recorded in the plan |
| F-002/AC-20 | Typecheck, lint, unit tests, build and `npm run validate` pass for the add-in; `deno test` passes for the function; the root `tsconfig`, lint and vitest do not try to compile the Deno function code. | Command output |

## Edge cases
- **Expired token at the backend:** 401 (AC-12). The panel shows "Backend rejected the sign-in" with Retry; Retry calls `acquireTokenSilent` again, which returns a fresh token. No automatic retry loop.
- **Wrong tenant** (guest or another directory signs in): the single-tenant authority should prevent it; if a token still arrives, the issuer check rejects it (AC-10).
- **Wrong audience** (a Graph token or a token for another app): rejected (AC-11). Graph tokens are not validatable per Microsoft, which is another reason never to accept them.
- **v1.0 token** (the registration was left at `requestedAccessTokenVersion` = null): rejected by the `ver`/issuer checks (AC-10); the plan's prerequisite sets version 2.
- **Missing or malformed `Authorization` header:** 401 (AC-8).
- **NAA unsupported** (old Outlook build, Outlook.com/Gmail mailbox): unsupported line (AC-1).
- **Popup blocked or closed by the user:** "Sign-in failed" with the code (AC-5). No Office dialog fallback in F-002 (ADR-0002).
- **Consent required:** with admin consent granted (plan prerequisite) users never see a prompt. Without it, the silent call raises `InteractionRequiredAuthError` and the popup shows consent or an "admin approval required" page, after which AC-5 applies.
- **Backend down, 5xx, DNS failure, timeout:** "Backend unreachable" with Retry (AC-7).
- **CORS preflight:** the `Authorization` header makes the browser send `OPTIONS` first; it must succeed without a token (AC-14). A disallowed origin fails in the browser and the panel shows "Backend unreachable".
- **Signing key rotation:** an unknown `kid` makes `jose` refetch the JWKS once; a still-unknown `kid` is 401 (AC-9).
- **No `name` claim:** "Backend connected" without a name (AC-7, AC-13).
- **Item changes while the pane is open:** out of scope (F-001 has no `ItemChanged` handling); the backend line does not depend on the item.

## Out of scope
OpenAI or any AI call; any email content sent to the backend; Microsoft Graph and Graph scopes (the auto-added `User.Read` is removed in the plan); sign-out UI; the Office dialog API fallback for non-NAA clients; per-user rate limiting and cost caps (ADR-0003 open item, needed before F-003); production hosting and the production manifest (D7); a custom domain for the function; storing anything in Supabase (database, storage, Supabase Auth); app roles or group-based authorization (any authenticated DIMSCALE user with consent may call `whoami`); `ItemChanged` handling; localization.

## Security notes (data flow change)
New outbound flows: task pane → `login.microsoftonline.com` (through the Outlook NAA broker) and task pane → `<project-ref>.supabase.co`. The only data sent to Supabase is the access token, which contains user identifiers (`oid`, `tid`, `name`, `preferred_username`) but no email content. `docs/architecture/security.md` needs a short update (plan step 1): the exact validation rules above, CORS allow-list, `verify_jwt = false` and why, and that Supabase (a processor in the chosen region) now receives user identity tokens.

## References
- https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in
- https://learn.microsoft.com/javascript/api/requirement-sets/common/nested-app-auth-requirement-sets
- https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA-Identity
- https://learn.microsoft.com/entra/identity-platform/msal-client-application-configuration
- https://learn.microsoft.com/entra/identity-platform/access-tokens
- https://learn.microsoft.com/entra/identity-platform/access-token-claims-reference
- https://learn.microsoft.com/entra/identity-platform/quickstart-configure-app-expose-web-apis
- https://learn.microsoft.com/entra/identity-platform/quickstart-configure-app-access-web-apis
- https://learn.microsoft.com/entra/identity-platform/identifier-uri-restrictions
- https://learn.microsoft.com/javascript/api/manifest/appdomains
- https://learn.microsoft.com/office/dev/add-ins/develop/addressing-same-origin-policy-limitations
- https://supabase.com/docs/guides/functions/function-configuration
- https://supabase.com/docs/guides/functions/auth-headers
- https://supabase.com/docs/guides/functions/auth
- https://supabase.com/docs/guides/functions/cors
- https://supabase.com/docs/guides/functions/secrets
- https://supabase.com/docs/guides/functions/deploy
- https://supabase.com/docs/guides/functions/quickstart
- https://supabase.com/docs/guides/functions/dependencies
- https://supabase.com/docs/guides/functions/unit-test
- https://supabase.com/docs/guides/auth/jwts

## Open questions
1. **D4, still open:** who creates the app registration and grants admin consent, and the tenant ID. Implementation is blocked until then.
2. **Gateway header with `verify_jwt = false`:** Supabase's docs don't say whether the platform still requires an `apikey` header when JWT verification is off. Plan step 3 tests it. If it is required, the publishable key (non-secret) would be sent too, which needs a spec change.
3. **MSAL cache location:** Microsoft's NAA sample sets `cacheLocation: "localStorage"`; the docs don't state the NAA default. Proposal: leave MSAL's default and set nothing. Confirm, or accept `localStorage` (tokens persisted on disk by the webview).
4. **Who can call:** should `whoami` (and later AI functions) be limited to an Entra group or app role, or is "any DIMSCALE user who consents" enough? This spec assumes the latter; a role check is a small later addition.
5. **Supabase region and project ref:** the function region (EU?) matters for GDPR (D5). Confirm the project's region.
6. **Outlook build on the test machine:** classic Outlook must be Version 2409 or later for NAA. F-001 was live-tested without recording the version.
7. **Production origin (D7):** `ALLOWED_ORIGINS` and the `brk-multihub://` redirect need the production origin later; F-002 configures `https://localhost:3000` only.
