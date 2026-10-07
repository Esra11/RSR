# Changelog

## v.1 — 2026-10-08

First versioned RSR deployment.

### Recording

- Captures keyboard, mouse, scrolling, menus, dialogs, charts, and workflow intent while preserving input order.
- Contains failures from inaccessible or unresolved controls instead of guessing actions.
- Handles access denied responses from elevated applications without crashing the recorder.

### Planning

- Generates execution plans and review PDFs through Microsoft Foundry.
- Validates intent generated actions and rejects incomplete or unsupported plans.
- Converts supported Excel interactions into verified filters, replacements, dynamic ranges, formulas, value pastes, and chart operations.

### Replay

- Uses live UI Automation with verified native fallbacks for Windows applications.
- Supports Excel, Edge, private browser windows, Windows Search, and elevated DebugDiag workflows tested during development.
- Replays dynamic filtered rows and populated ranges instead of relying on fixed recorded row numbers.
- Handles optional dialogs and UAC application startup without accepting unrelated windows.
- Uses semantic browser commands when toolbar accessibility identities change.

### Setup and maintenance

- Documents Microsoft Foundry resource creation and repository-root `.env` configuration.
- Keeps secrets, recordings, intermediate build output, and local runtime files out of source control.
- Includes focused regression checks for planning, recording, browser behavior, Excel workflows, filters, and window activation.
