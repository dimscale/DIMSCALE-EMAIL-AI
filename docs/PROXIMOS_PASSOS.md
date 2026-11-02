# Next steps (summary, 2026-11-02)

Source of truth: [docs/status.md](status.md). This page is a short reading guide.

## What works
- The pane is live on GitHub Pages and tested by the user in Outlook: tabs Resumo, Resposta, Traduzir and Automático; summary, suggested reply (language and tone, "Pedir alterações"), translate.
- The manifest is 1.0.3.0: "Assistente IA", one ribbon button "DIMSCALE Mail", `ReadItem` only.
- The Automático tab accepts contacts by hand; the F-013 backend is built. Nothing is sent automatically.
- Auto summary mode stays hidden until the DPO answers D5.

## What needs the user
- Live checks: 320 px and 200 % zoom, screen reader, injection with a live model.
- Answers to F-013 Q-C1, Q-C3, Q-C4 (see [research 2026-11-02](research/2026-11-02-f013-open-questions.md)).
- Decide D13 ([ADR-0006](decisions/0006-stateful-backend-for-batch-and-style.md), proposed) before F-015 and F-016.
- Confirm the pending amendments (D19).
- Tell the architect when the add-in is ready, then send the MMSI request.

## What needs MMSI or the DPO
- MMSI: [docs/admin/2026-10-23-mmsi-request.md](admin/2026-10-23-mmsi-request.md). A reply "Permissões atualizadas" arrived on 2026-10-29, but what was done is unconfirmed; the request has a box to tick the done items first.
- DPO: D5 (OpenAI) and D15 (stored contact list and, if accepted, style profile).
- F-013 live needs both, plus the returned values (`GRAPH_APP_CLIENT_ID` and the rest).
