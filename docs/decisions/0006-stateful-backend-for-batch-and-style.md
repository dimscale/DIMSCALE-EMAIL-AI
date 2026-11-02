# ADR-0006: Stateful backend for the manual batch (F-015) and the style profile (F-016)

- **Status:** **Accepted (2026-11-02), option B staged**, by the user's delegation ("pode implementar essas decisões", 2026-11-02; recorded by the orchestrator). Decision D13 is resolved. Staging stays: F-015 first with no new table; F-016's `style_profiles` only after the DPO answers D15. F-015 and F-016 remain Draft (not approved for implementation). The ADR-0003 pointer was added 2026-11-02. `overview.md` "Stateless" wording is to be aligned when F-015 starts.
- **Date:** 2026-11-02
- **Features:** [F-015 manual batch drafts](../specs/F-015-manual-batch-drafts.md), [F-016 style profile](../specs/F-016-style-profile.md)
- **Related:** [ADR-0003](0003-ai-backend.md) (says "stateless", "store nothing"), [ADR-0005](0005-server-side-mail-processing.md) (F-013 tables, already accepted)

## Context
ADR-0003 describes one stateless function per AI operation. That is no longer the state of the repo:
- `supabase/migrations/20261029120000_f003_analyze_usage.sql` adds the per-user usage counter (`analyze_usage`, function `an_take`) for F-003.
- `supabase/migrations/20261023120000_f013_auto_reply_drafts.sql` adds four F-013 tables (ADR-0005).
- Both follow one convention: table in `public`, RLS enabled with no policy, `revoke all ... from anon, authenticated`, access only by Edge Functions with the service role (RPC functions granted to `service_role`). Checked by grep of the migrations on 2026-11-02.

F-015 and F-016 were drafted on 2026-10-22, before this. They propose a `private` schema reached through `SUPABASE_DB_URL` with a new Postgres client. That is a different pattern from the one now in the repo.

## What would be stored
| Data | Feature | New? | Content |
|---|---|---|---|
| Per-user request counter for `draft-reply` and `style-profile` builds | F-015, F-016 | No. Reuse the F-003 counter pattern (`an_take`) with new keys; new buckets only if the function cannot take a key | `oid` (pseudonymised as in F-003), bucket, count. No content. |
| Style profile | F-016 | **Yes**, one table | `oid`, `tid`, `profile` (jsonb, at most 800 characters, fixed schema, user-reviewed), `schema_version`, `created_at`, `updated_at`, `last_used_at`. No email text, subject, address or message ID. |

F-015 alone (without a style profile) needs **no new table**: the batch reads and writes the mailbox through Graph in the pane, and the only server state is the counter.

## Facts checked
| Fact | Source |
|---|---|
| RLS enabled with no policy: nothing is accessible through the Data API with a publishable key. Secret keys bypass RLS. | https://supabase.com/docs/guides/database/postgres/row-level-security |
| `roamingSettings`: 32 KB per add-in, stored on the server only after `saveAsync`, "not secure storage" (readable through EWS and Extended MAPI), classic Windows needs a reload to see changes from another client, and web, Mac and new Outlook need a refresh or restart. | https://learn.microsoft.com/javascript/api/outlook/office.roamingsettings (via search summary; re-check the exact wording before relying on it) |
| Backups: Pro 7 daily, Team 14. Whether deleted rows persist in backups until expiry, and the backup region, are **not stated**. | https://supabase.com/docs/guides/platform/backups (from the F-016 spec; not re-read) |
| F-013 cannot read `roamingSettings`: it runs server side with Outlook closed. | F-013 spec, "Goal" |

## Options
| Option | Summary | Pros | Cons |
|---|---|---|---|
| **A. Stay stateless** | The profile lives in `roamingSettings`. The pane sends it with each `draft-reply` call. The counter reuses the F-003 table. | No new personal-data store; nothing for the DPO to accept beyond D5; smallest change. | Not secure storage (readable by EWS and MAPI). Other clients need a reload or restart to see changes. **F-013 (Outlook closed) can never apply the style.** The server must treat the profile as untrusted input on every call (it already must). 32 KB is ample for an 800-character profile. |
| **B. Supabase tables, existing convention** | One new table `style_profiles` in `public`, RLS on, no policy, `revoke` from `anon` and `authenticated`, read and written only through `service_role` RPC functions (as in the two existing migrations). Counter reused. | One source of truth across all clients. F-013 and F-004 can apply the style server side. Same infrastructure, no new vendor, no new Postgres client. Delete and export are single statements. | The server now stores a derived personal profile (DPO item). Backup persistence of deleted rows is undocumented. A leaver needs an owner for deletion. |
| **C. Hybrid** | `roamingSettings` is the source now; a server copy is added only when F-013 needs the style. | Defers the DPO question. | Two stores to keep in sync; the server copy needs B's work later anyway. |
| **D. `private` schema through `SUPABASE_DB_URL`** (as drafted in F-015 and F-016) | Direct Postgres client. | Matches the 2026-10-22 drafts. | A second access pattern, a new client dependency (CLAUDE.md rule 8 asks for an ADR for a heavy one), and a database URL as a new secret. No benefit over B for these two small tables. |

## Recommendation (proposed, the user decides)
**Option B, staged:**
1. **F-015 first, with no new table.** Reuse the F-003 counter for the `draft-reply` rate limit. F-015 works without a profile (F-015/AC-20 "absent" cases).
2. **F-016 adds one table `style_profiles`** only after the DPO answers D15 for the profile. The spec text changes from `private.style_profiles` reached by `SUPABASE_DB_URL` to a `public.style_profiles` table with RLS, `revoke` and `service_role` RPC functions (the existing convention). Option D is dropped.
3. If the DPO refuses the stored profile, fall back to **A** for F-015/F-016 and accept that F-013 uses the default style.

Reasons for B over A: the user's own direction is that F-013 works with Outlook closed on every device, and A can never give the style to that path. Reasons for staging: F-015 needs nothing new, and the only new personal data (the profile) is held until the DPO decides.

## Retention, RLS and deletion (for B)
- **Region:** Supabase project `qepzpkseyzclxnqguqta`, eu-west-1 (D11).
- **RLS:** enabled, no policy, no grant to `anon` or `authenticated`. The Data API cannot reach the table. Every function statement is keyed by the `oid` from the validated token (never from the body) and covered by a two-user test (F-016/AC-13).
- **Retention:** the profile is kept until the user deletes it or it is unused for 365 days (`last_used_at`, else `updated_at`). Counter buckets are deleted after 7 days. The purge uses Supabase Cron; if it can't be enabled, a documented monthly admin SQL (F-016 Q9).
- **User controls:** view, edit, "Copiar como JSON" (export), "Apagar perfil".
- **Leavers:** an admin deletes the row by `oid`. The owner is still open (F-016 Q8).
- **Backups:** deleted rows may stay in backups until they expire. This is a DPO item.
- **Never stored:** email bodies, subjects, addresses, message IDs, recipient names, AI replies or tokens.

## DPO link (D15)
D15 today covers the F-013 contact list. Proposed extension: add "the derived style profile (no email text), 365-day unused retention, user deletion, offboarding, backup persistence". The build step (samples sent once to OpenAI) stays under D5. The notice wording for both is the DPO's.

## Consequences
- ADR-0003's "store nothing" is superseded for these operations once this ADR is accepted. Update ADR-0003 with a one-line pointer to ADR-0005 and ADR-0006.
- `security.md` needs a "Stored data" section covering `style_profiles` (when F-016 is approved), and the F-015 and F-016 specs need their storage text aligned with the existing migration convention.
- Docs to update at acceptance: F-015 D-2 and D-6, F-016 D-4, `overview.md` ("Stateless" wording).

## Open
- **User (D13):** accept B staged, choose A, or choose C.
- **DPO (D15):** the style profile as stored data; retention; backup persistence.
- **Admin or orchestrator:** the owner of leaver deletion; whether Supabase Cron can be enabled.
- Verify at implementation: the exact `an_take` contract (key and window) before reusing it for `draft-reply`.
