# ADR-0003: Backend for OpenAI calls

- **Status:** Accepted (2026-10-22). Option A, Supabase Edge Functions. The user created the Supabase project.
- **Date:** 2026-10-22

## Context
The OpenAI API key must never reach the client, so a server component is required. No backend exists yet, and the brief forbids introducing one without an explicit decision.

## Options
| Option | Pros | Cons |
|---|---|---|
| **A. Supabase Edge Functions** (recommended) | Managed, secret store, no server to run, TypeScript (Deno), EU region available | New vendor account; Entra token validation must be written by us (JWKS) |
| B. Azure Functions | Same cloud as Microsoft 365, native Entra integration (Easy Auth) | Azure subscription and ops overhead |
| C. n8n or a VPS | — | Excluded by the brief without strong justification; more ops and attack surface |

## Decision
Option A, one stateless function per AI operation (or one function with an operation field):
1. Validate the Entra ID token (issuer = DIMSCALE tenant, audience = add-in API, expiry).
2. Validate and truncate the input.
3. Call OpenAI with a fixed system prompt and the email in a delimited user message, asking for structured JSON output.
4. Validate the output schema and return it.
5. Store nothing and log metadata only.

## Open
- ~~Vendor~~: Supabase (accepted 2026-10-22).
- The DIMSCALE OpenAI account and model choice (verify the current models and data-usage terms on the OpenAI docs at implementation time).
- Per-user rate limiting and a cost cap.
