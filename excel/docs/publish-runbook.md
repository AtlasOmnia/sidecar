<!-- SPDX-License-Identifier: Apache-2.0 -->
# Publication handoff — not executed

Nothing in this remediation run publishes a repository or a site. The current development history remains
private. Public names are decided:

- Repository: `https://github.com/AtlasOmnia/sidecar-excel`
- Optional hosted-sidebar base: `https://atlasomnia.github.io/sidecar-excel/` (the plugin's bundled localhost sidebar is the default)
- Optional hosted manifest: `https://atlasomnia.github.io/sidecar-excel/manifest.xml`; for default sideloading run `/sidecar-excel manifest` or `hermes sidecar-excel manifest`

## Owner's later publication run

1. Review/export the approved tracked tree only, excluding `.git`, runtime data, dependencies and build output.
   Scan the export for personal paths, private host names, credentials and internal documents. Do not push this
   development branch or its history. Preserve it privately.
2. Initialize a separate publication repository and create one fresh root commit using the owner's approved
   GitHub noreply author identity. Do not amend or rewrite this development repository's history.
3. The catalog `sha: TODO_PUBLISH_COMMIT` is deliberately not an installable catalog pin. Record the exported
   release's 40-hex commit SHA in the catalog submission; bump version and pin together on later releases.
   The pin must name the exported plugin commit, not an unrelated development-history commit. Submit the
   catalog separately after publication, so a self-referential commit hash is not required.
4. Create `AtlasOmnia/sidecar-excel`, push only the fresh publication history, and enable Pages with source
   GitHub Actions. The prepared `.github/workflows/pages.yml` builds the static `dist` artifact and deploys it.
5. Verify the repository and (if Pages is deployed) hosted manifest URLs, `taskpane.html`, `commands.html`, assets
   and `SHA256SUMS` from the public origin. Preserve webpack's `taskpane.*.js.LICENSE.txt`. Verify checksums
   against the release build.
6. In a fresh disposable Hermes profile, qualify the documented plugin install and default localhost same-origin
   Connect, Allow, read/write/readback and revoke. Test the optional Pages override separately if maintained.
   Trust prompts require the operator; do not automate keychain changes. Record the exact tested publication SHA.

When using GitHub Pages, its origin is account-wide, not path-isolated. Keep all content on that origin trusted;
consider moving to a dedicated origin later. Do not change the CA shape or trust command without a separately approved
live retest. Planned TLS hardening remains `-p ssl` and a server-auth EKU on the CA.

## Latest live evidence and remaining scope

On 2026-11-11 about 00:26, desktop Excel for Mac ran the plugin and sidebar built from `a7b3fd7`.
Connect and Allow succeeded; a 3-row table with header in A1:B4 and `=SUM(B2:B4)` in B5 were read
back from the workbook. Revoke dropped the sidebar to “needs your OK”. The bridge answered 200 to
the configured origin and 403 to another origin. The sidebar was served from a private test site via
`addin_base`, with the CA trusted in the login keychain; this Mac also has it in the System keychain.

Earlier the same night at `35ef09e` (behaviour unchanged since), reversal on request cleared the range
(no real undo); plugin reload with the sidebar open reconnected by itself and edited again; approval
survived an Excel restart; the next Connect after revoke asked for approval again. See the README for trust
and removal instructions. These private-site runs are not public-origin acceptance.

Still untested: Windows desktop Excel, dark theme, keyboard-only use, the optional public GitHub Pages override,
TLS-only trust, runtime accessibility, or full live tool/WorkbookPlan coverage. Publication/install/Pages
execution and CA server-auth EKU hardening are not proven by these runs.
