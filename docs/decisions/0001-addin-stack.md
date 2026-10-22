# ADR-0001: Add-in stack and scaffold

- **Status:** Accepted (2026-10-22)
- **Date:** 2026-10-22

## Context
The repo is empty. Microsoft's quickstart scaffolds Outlook add-ins with the Yeoman generator (`yo office`), which offers both manifest types.
Source: https://learn.microsoft.com/office/dev/add-ins/quickstarts/outlook-quickstart-yo

## Decision
- Scaffold with `yo office`: Outlook task pane, **TypeScript**, React if the generator offers React for Outlook (to verify when scaffolding; fall back to plain TypeScript).
- Package manager: **npm** (already installed, used by the generator).
- Manifest: **add-in only (XML)** for the MVP. It is production-supported for Outlook. Move to the unified manifest only after verifying its client support on the Outlook clients DIMSCALE uses.
  Source: https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests
- Keep whatever lint, test and build tooling the generator provides. Add a unit test runner only if none is included.

## Consequences
- Standard dev loop (`npm start` to sideload, `npm run validate` to check the manifest).
- The XML manifest cannot combine the add-in with Teams apps. That is not needed.

## Open
- Which Outlook clients DIMSCALE uses (classic Windows, new Outlook, web, Mac).
