# AGENTS.md — Sidecar for PowerPoint

Desktop PowerPoint sidebar for a user's own Hermes agent. The Hermes plugin (`plugins/atlas-sidecar-powerpoint/`) is the only
thing a user installs; it runs the loopback bridge, the approval flow and the chat forwarding itself. The sidebar
(`src/`) is a static bundle built with `npm run build:static -- --base <https site>/`.

## Layout

- `plugins/atlas-sidecar-powerpoint/__init__.py` — registers the 15 `sidecar_powerpoint_*` tools, `/sidecar-powerpoint`, `hermes sidecar-powerpoint`.
- `plugins/atlas-sidecar-powerpoint/desk_bridge.py` — the same design as the Word and Excel plugins, not identical files; app
  specifics live in `desk_app.py`. Preserve PowerPoint-specific lifecycle protections.
- `plugins/atlas-sidecar-powerpoint/companion/` — the sidebar ticket protocol (session, heartbeat, pending, result,
  invoke), including request boundaries, response deadlines and bounded overload cleanup.
- `src/deskConnect.ts` — bridge discovery and the Connect-to-Hermes flow; `src/officeBridge*.ts` — ticket client;
  `src/agent.ts` — chat loop; `src/taskpane/` — UI. `manifest.template.xml` + `scripts/build-static.mjs`.
- `docs/desk-bridge.md` — protocol and security design.

## Rules

- Tool names keep the `sidecar_powerpoint_` prefix (Workbench owns `powerpoint_*`); ports stay in 47650-47654.
- No API key files, no API key in the sidebar, no `.env` reads: `API_SERVER_KEY` only through
  `agent.secret_scope.get_secret`. Per-install TLS private keys are the explicit exception: owner-only in plugin data,
  never logged or committed; CA trust is manual, never performed by a tool. Never log bodies, grants, keys or document text.
- No tool may approve, connect or revoke a sidebar.
- Tests never touch the real Hermes home (`scripts/tests/conftest.py` sets a temporary `HERMES_HOME`).

## Commands

```sh
npm ci
npm test                       # vitest + pytest scripts/tests
npm run generate:plugin-schemas # after changing tool schemas in src/
npm run build:static -- --base https://example.org/powerpoint/
hermes plugins validate plugins/atlas-sidecar-powerpoint
```
