# Changelog

## Unreleased

### Changed
- Product and display name: Sidecar for Word, part of Sidecar by Atlas Omnia.
- Plugin id and settings namespace: `atlas-sidecar-word`; tools use `sidecar_word_*`, and commands use
  `/sidecar-word` and `hermes sidecar-word`. Approvals from the previous plugin id are not carried over,
  so the sidebar asks for approval again after migration.
- Public install defaults and publication instructions are prepared; the public GitHub Pages origin is not live-tested.
- Documentation records the 2026-11-11 Mac run at `7eaf1fc`, including formatting lost after a failed edit
  and repair. Earlier Heading 1 success does not guarantee formatting survives a failed edit and repair.

## 1.1.0 — plugin-hosted sidebar

### Changed
- The loopback bridge runs inside the Hermes plugin; no separate companion service, proxy or dev server
  is needed. The sidebar is a static HTTPS build.
- The bridge uses verified localhost TLS with a per-install local CA requiring manual user trust.
- Sidebar approval is remembered as a hashed grant and can be revoked. It is not per-edit approval.
- Chat is forwarded to the user's Hermes API server. The API key is read through Hermes and never reaches
  the sidebar.
- Added Apache-2.0 LICENSE and NOTICE files and a draft catalog entry.

## 2026-10-14 — sidebar appearance

### Changed
- Updated task pane spacing, typography, composer, conversation and settings styling.
- Added Teal/Blue/Gold/Red palettes, Auto/Dark/Light appearance and keyboard-operated settings swatches.
  Headless theme checks are not native Word acceptance; dark theme and keyboard-only use remain untested live.

## 2026-08-13 — security hardening

### Changed
- Updated dependency overrides and lock metadata.
- Added a bounded JSON nesting guard and hardening tests for companion requests.

## 0.2.0 — 2026-07-10

### Added
- WordApi 1.2 manifest requirement and mocked Office.js execution-path tests.
- Non-stream OpenAI-compatible JSON response handling.
- Whole-word find/replace, bounded document structure reads, named styles, OOXML insertion and comments.

### Changed
- Guarded comment insertion with a WordApi 1.4 runtime requirement-set check.
- Corrected paragraph counts and explicit selection-formatting properties.

Historical setup and internal acceptance notes for the retired standalone service are omitted; use the README
for current installation instructions, live-test evidence and known issues.
