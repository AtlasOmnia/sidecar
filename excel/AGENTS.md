# AGENTS.md — Sidecar for Excel

Desktop Excel sidebar for a user's own Hermes agent. The Hermes plugin (`plugins/atlas-sidecar-excel/`) is the only
thing a user installs; it runs the loopback bridge, the approval flow and the chat forwarding itself. The sidebar
(`src/`) is a static bundle built with `npm run build:static -- --base <https site>/`.

## Verified scope and operating limits

Latest live run: 2026-11-11 about 00:26 in desktop Excel for Mac, plugin and sidebar built from
`a7b3fd7`. Connect and Allow succeeded; a 3-row table with header in A1:B4 and `=SUM(B2:B4)` in B5
were read back from the workbook. Revoke dropped the sidebar to “needs your OK”. The bridge answered
200 to the configured origin and 403 to another origin. The sidebar used a private test site via `addin_base`;
the CA was trusted in the login keychain, and this Mac also has it in the System keychain.

Earlier the same night at `35ef09e` (behaviour unchanged since): reversal on request by clearing the range
(no real undo), plugin reload with the sidebar open (it reconnected by itself and edited again), approval
remembered after restarting Excel, and the next Connect after revoke asked for approval again.
Windows desktop Excel is untested. Also not tested: dark theme, keyboard-only use, the public Pages origin,
TLS-only trust, Excel for the web, or full tool/WorkbookPlan coverage. Reopen the sidebar if it does not reconnect.
Live-tested with a private sidebar test site; public-origin acceptance is pending publication.

Trust the plugin's local CA once, manually, using the login-keychain command printed by the bridge
(or explicitly request `/sidecar-excel trust`). The Excel test Mac also has it in the System keychain.
System-keychain trust was not required in our Word and PowerPoint runs (login-keychain trust only). Removal: Keychain Access > login > Certificates > “Sidecar local CA ...” > Delete;
remove any matching System entry if you installed one, then delete only the plugin's `tls` folder.
Uninstall/revoke do not remove OS trust. No tool should change keychains.

The default `addin_base` is `https://atlasomnia.github.io/sidecar-excel/` (public live test pending publication).
For another site set `plugins.entries.atlas-sidecar-excel.settings.addin_base` to that HTTPS base.
Match the built manifest's base URL to that setting. See README for the build/config commands.
The bridge/approval is local; Hermes may send workbook/chat data to the configured model provider.

## Layout

- `plugins/atlas-sidecar-excel/__init__.py` — registers the 28 `sidecar_excel_*` tools, `/sidecar-excel`, `hermes sidecar-excel`.
- `plugins/atlas-sidecar-excel/desk_bridge.py` — app-neutral code intended to be shared with Word and PowerPoint;
  specifics live in `desk_app.py`. This checkout has lifecycle fixes not yet propagated to siblings: do not claim
  byte-for-byte parity without checking each repo, and do not edit sibling repos as part of an Excel-only task.
- `plugins/atlas-sidecar-excel/companion/` — the sidebar ticket protocol (session, heartbeat, pending, result,
  invoke); preserve the Excel WorkbookPlan transaction contract and trusted pre/post hooks.
- `src/deskConnect.ts` — bridge discovery and the Connect-to-Hermes flow; `src/officeBridge*.ts` — ticket client;
  `src/agent.ts` — chat loop; `src/taskpane/` — UI. `manifest.template.xml` + `scripts/build-static.mjs`.
- `docs/desk-bridge.md` — protocol and security design.

## Rules

- Tool names keep the `sidecar_excel_` prefix (Workbench owns `excel_*`); ports stay in 47640-47644.
- No API key files, no API key in the sidebar, no `.env` reads: `API_SERVER_KEY` only through
  `agent.secret_scope.get_secret`. Per-install TLS private keys are the explicit exception: owner-only in plugin data,
  never logged or committed; CA trust is manual, never performed by a tool. Never log bodies, grants, keys or workbook text.
- No tool may approve, connect or revoke a sidebar.
- Tests never touch the real Hermes home (`scripts/tests/conftest.py` sets a temporary `HERMES_HOME`).

## Commands

```sh
npm ci
npm test                       # vitest + pytest scripts/tests
npm run generate:plugin-schemas # after changing tool schemas in src/
npm run build:static -- --base https://example.org/excel/
hermes plugins validate plugins/atlas-sidecar-excel
```
