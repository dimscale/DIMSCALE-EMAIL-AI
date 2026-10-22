# ADR-0002: User authentication with nested app authentication (NAA)

- **Status:** Proposed (needed by F-002; F-001 needs no auth)
- **Date:** 2026-10-22

## Context
Mailbox-tier features need delegated Graph tokens, and the AI backend needs to know who is calling. Microsoft documents MSAL.js nested app authentication (`createNestablePublicClientApplication`) as the SSO method for Office add-ins. It needs an SPA app registration with the redirect `brk-multihub://<origin>`, and requirement set `NestedAppAuth 1.1`.
Source: https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in

## Decision (proposed)
- Use `@azure/msal-browser` with NAA. Get tokens with `acquireTokenSilent`, falling back to `acquireTokenPopup`.
- One Entra app registration (single tenant: DIMSCALE) that exposes an API scope for the AI backend. Graph scopes are added incrementally.
- When NAA is unsupported, show an "unsupported client" message. No custom fallback login for the MVP.

## Consequences
- Needs a DIMSCALE tenant admin to create the app registration.
- Clients without NAA can't use the AI features.

## Open
- Confirm the tenant and who can create the app registration and grant consent.
- ~~Verify the NAA flow for a custom API scope~~ Verified 2026-10-22 for F-002: NAA can acquire a token for the add-in's own API. The OfficeDev sample [Outlook-Add-in-SSO-NAA-Identity](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA-Identity) requests `api://<clientId>/<scope>` through `createNestablePublicClientApplication` and sends it to its own web API. Setup: one registration as both client and API, Application ID URI `api://<clientId>` (a secure default per [identifier URI restrictions](https://learn.microsoft.com/entra/identity-platform/identifier-uri-restrictions)), a delegated scope (F-002 uses `access_as_user`), the app's own client ID under "Authorized client applications", SPA redirect `brk-multihub://<origin>`, and `requestedAccessTokenVersion` = 2 so the API receives v2.0 tokens (`iss` = `https://login.microsoftonline.com/<tenant-id>/v2.0`, `aud` = client ID; [access tokens](https://learn.microsoft.com/entra/identity-platform/access-tokens)). Single tenant: authority `https://login.microsoftonline.com/<tenant-id>` ([MSAL configuration](https://learn.microsoft.com/entra/identity-platform/msal-client-application-configuration)). Details in [F-002](../specs/F-002-authenticated-backend.md).
