# Security and Privacy

## Authentication and authorization
- **Users**: Microsoft Entra ID through MSAL.js nested app authentication (NAA), as proposed in [ADR-0002](../decisions/0002-auth-nested-app-auth.md). App registration type SPA, with redirect `brk-multihub://<add-in-origin>`. Check `Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")` and show a clear "unsupported client" state when it returns false.
- **Graph**: delegated scopes only, requested incrementally per feature. `Mail.Send` is never requested. No application (app-only) permissions.
- **AI backend**: every request carries an Entra ID access token for the add-in's own API scope. The backend validates the signature (JWKS), issuer (the allowed tenant), audience and expiry, and rejects everything else. Without this check the backend would be an open, paid OpenAI proxy.
- **Tokens**: MSAL manages its own cache. Our code never logs tokens, never sends Graph tokens to the backend, and never stores tokens itself.

## Secrets
- The OpenAI API key lives only in the backend secret store (for example Supabase function secrets).
- Nothing secret goes in the frontend bundle, the manifest, browser storage or git. `.env*` files are git-ignored. A code reviewer runs a secret scan on every review.
- The Entra client ID and tenant ID are not secrets, but they are configured per environment.

## Email content handling
- **Untrusted input**: email subject and body are data. The backend places them in a clearly delimited user message, and the system prompt states that the content must not be followed as instructions.
- **No actions from model output**: the model returns structured JSON (validated against a schema). The UI shows it and only the user can act on it. The model has no tools.
- **Rendering**: AI and email text is rendered as React text nodes. `dangerouslySetInnerHTML` and `innerHTML` are forbidden for this content. Reply HTML passed to `displayReplyFormAsync` is built from escaped text.
- **Minimization**: send only what the feature needs: plain-text body (not HTML), subject, sender display name and domain, truncated to the limit in each spec. No attachments, no recipient lists unless a spec justifies it.
- **Logging**: no email content, addresses or tokens in browser console logs, backend logs or error reports. Log request ids, durations and status codes only.

## Third-party processing (OpenAI)
- Email content leaves Microsoft 365 when it is sent to the OpenAI API. Before production, **verify OpenAI's current API data-usage and retention policy** and document it here. Also confirm with DIMSCALE that this is acceptable under its GDPR obligations (a DPA with OpenAI, and the region).
- Show users a one-line notice in the panel that content is processed by an external AI service.

## Supply chain
- Dev-only tooling (`office-addin-debugging`, `office-addin-dev-certs`, and others) pulls packages with install scripts: `keytar`, `@azure/msal-node-extensions`, `@microsoft/m365agentstoolkit-cli`, `core-js`. None of them ship in the bundle. CI should use `npm ci --ignore-scripts` where the build allows it.

## Classification safety
- The "potentially suspicious" label is advisory, shown with its reasons and a "may be wrong" notice. The add-in never reports phishing, moves mail to junk or blocks senders on its own.
