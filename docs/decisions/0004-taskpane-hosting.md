# ADR-0004: Task pane hosting

- **Status:** Accepted (2026-10-22). The repo turned out to be private, and free GitHub Pages needs a public repo. The user decided to make `dimscale/DIMSCALE-EMAIL-AI` **public**. Azure Static Web Apps was dropped because creating an Azure account required a payment card. Before the switch, the full history (20 commits, all branches) was scanned: no secrets, no sensitive files, and no personal data beyond the commit author identity.
- **Date:** 2026-10-22

## Context
The task pane (static HTML, JS and icons) is served from `https://localhost:3000` during development. For DIMSCALE users it needs a public HTTPS origin. Supabase stays the backend (ADR-0003), but on `*.supabase.co` both Edge Functions and Storage serve HTML as `text/plain`. HTML is allowed only with a paid custom domain.
Sources:
- https://supabase.com/docs/guides/functions/limits
- https://supabase.com/changelog/29633-xhtml-responses-are-only-allowed-with-a-custom-domain-enabled

## Options
| Option | Cost | Notes |
|---|---|---|
| **A. GitHub Pages, default domain** `https://dimscale.github.io/DIMSCALE-EMAIL-AI/` (chosen) | Free | No new vendor, deployed by a GitHub Action, HTTPS. No custom headers. The repo must stay public on the free plan. The origin is shared with any other Pages site in the `dimscale` org. |
| B. GitHub Pages with `mailai.dimscale.com` | Free + one DNS CNAME | Own origin. The user may request the DNS record later. |
| C. Azure Static Web Apps | Free tier | Needs an Azure subscription. Allows custom headers. |
| D. Supabase custom domain | Paid plan + add-on | Single vendor. |

## Decision
Option A for now. Moving to option B later only changes our side: the manifest URLs, the Entra SPA redirect `brk-multihub://<origin>`, the function's `ALLOWED_ORIGINS`, and a manifest re-upload by the admin.

## Consequences
- Users configure nothing. The admin deploys the manifest centrally (Microsoft 365 admin center → Integrated apps) and grants admin consent.
- Task pane code updates reach users on deploy, with no manifest change. A manifest re-upload is needed only when the manifest itself changes (URLs, permissions, surfaces, buttons, icons).
- Asset filenames must be content-hashed so Outlook's webview cache doesn't serve stale code.
- The shared `dimscale.github.io` origin is an accepted risk: only org members can publish Pages, and tokens still need the user's Entra sign-in in Outlook.
