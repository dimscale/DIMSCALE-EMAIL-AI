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
- Verify, in the Microsoft docs, the exact NAA flow for a token to a custom API scope (not only Graph) before F-002.
