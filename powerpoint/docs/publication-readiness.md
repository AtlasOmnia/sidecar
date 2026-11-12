<!-- SPDX-License-Identifier: Apache-2.0 -->
# Publication readiness — local review fixes

Nothing was published during this pass. No installed plugin, served sidebar, profile,
gateway or keychain was changed. Existing private history was not rewritten.

## Observed verification

- Failing first, before the implementation changes: `test_publish_readiness.py`
  reported **7 failed**. The revoked-approval case failed with `DID NOT RAISE BridgeError`;
  all three inherited-gateway CLI cases failed with `CLI started a listener`.
  Plain status also incorrectly included `add-trusted-cert`, the installed license files
  were absent, and the Allow question omitted the full agent authority warning.
- A separate failing-first new-CA-name test reported **1 failed, 7 deselected**:
  the common name still began `Sidecar local CA`, not `Sidecar for PowerPoint local CA`.
- Final `npm test`: **38 Vitest files, 366 tests passed; 204 pytest tests passed**,
  no failures or skips. Three Python warnings: anyio's pytest API deprecation and the
  intentional TLSv1/TLSv1.1 rejection probes. Existing security tests remain present.
- Explicit private-base `npm run build:static -- --base <private-test-site>/` succeeded:
  webpack 5.108.4, two performance warnings (351 KiB `.icns` asset and code-splitting
  recommendation). This is a local build, not a deployment or a fresh Office live test.
- Desktop PowerPoint for Mac was live-tested at `315f1f0` on 2026-11-11 about 00:32,
  including one-request editing/readback, Try again recovery after plugin reload and revoke.
  The earlier `3682e5f` run at about 00:01 also covered a second slide, a request open for
  about 24 seconds, reversal and approval after restarting PowerPoint. Both used a private
  test site via `addin_base` and login-keychain CA trust only. The plugin-served same-origin
  default in 1.2.0 has not yet been live-tested in Office; TLS-only trust remains untested.

## Publication gate / runbook

Only after publication is authorized:

1. Prepare the public repository `AtlasOmnia/sidecar-powerpoint` from a fresh public
   root commit in a separate publishing checkout, using a noreply author identity.
   Keep this checkout's private history private; do not push this branch unchanged.
2. Fill the catalog's `PUBLISH_TIME_COMMIT_SHA` placeholder with the public pinned
   40-hex commit. It is deliberately not a valid submission pin yet.
3. Optional, only for an external hosted sidebar: enable GitHub Pages with GitHub Actions as its source. The prepared
   `.github/workflows/pages.yml` builds `dist/` for
   `https://atlasomnia.github.io/sidecar-powerpoint/` and deploys it. The default plugin-served localhost sidebar needs no Pages deployment.
4. If publishing that optional external site, verify the repository and Pages manifest, taskpane and SHA256SUMS URLs.
   Retain webpack's bundled license files with the site.
5. In an isolated install, exercise the plugin-served localhost default; for an external override, use that exact hosted origin too.
   Connect, matching-code Allow, one prompt/edit/readback, one upstream request and real answer, restart/revoke.
   Approval is origin-wide, not path-wide; only configure a hosted origin you trust.

TLS-only trust (`-p ssl`) and a server-auth EKU on the CA remain planned hardening;
the trust command and CA extension policy were intentionally not changed.
The PowerPoint icon redesign and a full SPDX-header sweep remain non-blocking follow-ups.

## Reuse lessons

- Fence inherited-gateway CLI processes before registration/autostart, not only inside
  their command handler; the listener can already have started by handler dispatch.
- Check both session membership and current approval immediately before issuing a tool
  ticket; a connected session or outstanding pane long-poll is not continuing authority.
- Pin live evidence to the rebuilt sidebar and reloaded plugin commit. A backend test or
  a private-origin test does not establish public-origin acceptance.
