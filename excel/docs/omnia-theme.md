# Omnia Voice theme

The Excel pane uses the shared, unmodified Omnia Voice rev-3 token sheet, followed
by `taskpane.css`. The synchronous `omniaTheme` entry resolves preferences before
styles and body paint; it does not wait for the Office.js network script.

- Default: Teal + Auto.
- Stored key: `omnia.theme`, JSON `{ "palette": "teal", "mode": "auto" }`.
- Palettes: teal, blue, gold, red. Modes: auto, dark, light.
- Auto: Office `isDarkTheme`, then Office body-background relative luminance,
  then `prefers-color-scheme` when Office's theme is unavailable.
- Before Office initialization, Auto uses the available OS theme; `Office.onReady`
  resolves the host theme. Focus/visibility/OS changes and a lightweight 1-second
  theme snapshot refresh keep Auto current across Excel hosts without depending
  on an unsupported Office theme-change event.
- Palette/appearance controls use native labelled radios (keyboard arrow keys)
  and apply immediately. The existing Save/Close semantics for connection settings
  are unchanged; theme preferences are independent of that form.
- Existing manifest identity, version, locations, hosts, permissions and requirement
  sets are unchanged. Both ribbon override versions use the full label
  `Sidecar for Excel`.

## Verification

Use Node 22. Temporary files must stay in this worktree's `.tmp` directory.
Do not run `npm start`, a sideload command or any live service for theme QA.

    npm test
    node node_modules/typescript/bin/tsc --noEmit
    npm run build
    node scripts/validate-omnia-theme.mjs <optional-path-to-approved-kit>
    node scripts/omnia-theme-capture.mjs --after <output-folder>

The offline validator uses the installed Office manifest parser, XML parsing,
resource resolution, icon dimensions, the pinned token SHA-256 and pre-theme
manifest invariants. It never calls Microsoft's validation service. The existing
`npm run validate` uploads XML to that service, so it is not an offline-safe gate.
This local check is not a substitute for Microsoft service/schema certification.

When running from an agent process that inherits `HERMES_OFFICE_UPSTREAM_KEY`,
`HERMES_OFFICE_UPSTREAM_HOST` or `HERMES_OFFICE_UPSTREAM_PORT`, remove those three
variables from the test subprocess environment only. A Windows lifecycle test
deliberately asserts that these variables are initially absent. Do not edit any
credential file or production environment to accommodate the test.

The capture harness uses installed headless Edge, an isolated user profile under
`.tmp` and the first free loopback port from 47140. It renders the actual production
bundle from `dist`, stubs Office.js and intercepts every non-harness request. Its
messages, activity and bridge display are synthetic fixtures, not evidence of a
model response, workbook mutation or a live Office/companion session.

`--before` captures the currently built bundle as `shot-excel-before.png`; build the
pre-theme revision in an isolated location to regenerate a true baseline.

Expected after-shots: dark teal, light teal, dark settings and dark gold. The JSON
receipt also checks scheme changes, persistence and width 320/380/450 layouts.

Approved token SHA-256:
`9d4a6a258b587d6c65c9773521ef84e3442915027e052d1f2e82725242cfc3d5`

## Change notes

- Added Omnia Voice Calm pane styling and mark PNG assets.
- Added Office-aware, persisted palette and appearance preferences.
- Rebranded both ribbon/manifest versions without changing add-in routing.
- No repository CHANGELOG existed at the time of this visual-only change.
