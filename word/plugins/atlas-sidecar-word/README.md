<!-- SPDX-License-Identifier: Apache-2.0 -->
# Sidecar for Word (1.2.0)

## Judge quickstart — Word only (Mac)

The default sidebar is bundled with the plugin and served from its HTTPS loopback bridge at `https://localhost:<actual-port>/sidecar-word/`; GitHub Pages or other hosting is not required.
1. Install `hermes plugins install https://github.com/AtlasOmnia/sidecar-word#plugins/atlas-sidecar-word --enable`; enable the Hermes API server and start the gateway as described in **Setup** below.
2. Review and run this install's printed local-CA trust command. In Hermes chat run `/sidecar-word manifest`, or run `hermes sidecar-word manifest` in a terminal; follow the printed copy instruction to sideload the generated manifest into Word.
3. Open a blank synthetic file, click **Connect to Hermes**, compare the code in Hermes, then **Allow**. Allow grants the configured origin access to your **Hermes agent and all its tools until revoked**, not per-edit approval.
4. Ask: “Write Quarterly Summary as Heading 1, then one short paragraph about fruit prices.” Then request an independent document readback. Use native Office Undo where supported; Stop does not roll back edits. Revoke with `/sidecar-word revoke` when done; removing CA trust is a separate manual step.
Windows desktop Office is untested. A self-hosted site can be selected through `addin_base`; approval then covers that configured origin.


**Sidecar by Atlas Omnia.** Your Hermes agent, in the sidebar of Microsoft Excel, Word, and PowerPoint.

Sidecar puts your Hermes agent in a side panel next to the file you're working on. Ask it to draft, edit, analyze, or build without leaving your document, spreadsheet, or deck.

Sidecar is a Hermes plugin plus an Office add-in that opens your agent in the sidebar of Microsoft Excel, Word, and PowerPoint. It works with the file you have open, so you can ask for a formula, a rewrite, or a slide outline and get it in place.

In Word, Sidecar can:

- read the selection, the document text, a document overview and its paragraph structure;
- insert text at the start, the end or the selection, replace the selection, find and replace;
- apply an exact-match edit plan, optionally highlighting what changed;
- format the selection and apply named Word styles;
- add comments on the selection;
- insert trusted WordprocessingML (OOXML) content.

After one manual local-CA trust step and sidebar approval in Hermes, the agent can read and edit the document
from that sidebar. This is not per-edit approval. Hermes may use remote models and tools according to your
configuration; the bridge is local. No API keys stored by the plugin (TLS private keys live in plugin data).

Your agent: it runs on your own Hermes setup, with the models and tools you already use.

Chat with your own Hermes agent in a sidebar inside **desktop Word for Mac** while Hermes reads,
reviews and edits the open document.

The plugin serves the bundled static sidebar from its HTTPS loopback bridge at
`https://localhost:<actual-port>/sidecar-word/`. It forwards the sidebar's chat to your Hermes API server and gives
Hermes twelve `sidecar_word_*` tools that act on the open document through the sidebar. No separate service, proxy
or GitHub Pages site is required; HTTPS loopback still requires one explicit user trust step for a per-install local CA.

Status — accepted repair at `d4d0de942d47f4e27bf1cb8f45e6e6ce7d21444d`: 429 JavaScript and 213 Python tests passed, plus types, generated schemas, plugin validation and static build. Native Word 16.114.4 on macOS 27.2 verified exact Heading1/Normal paragraph boundaries and styles, targeted style repair with another paragraph selected, both search-tool boundaries and literal caret, supported short-edit native Undo, preserved bold/table/image fixtures, and revoke. Final-bundle heading smoke passed. These are historical private-origin results for the earlier build; they do not qualify this changed loopback-serving path, the public origin or every Word operation.

Historical pre-repair failure (superseded by accepted repair above): Tested live in desktop Word for Mac on 2026-11-11 at about 00:29, with the plugin and sidebar built from `7eaf1fc`. Connect and Allow worked. The prompt “Write the line Quarterly Summary as a Heading 1, then one short paragraph about fruit prices” wrote both paragraphs, which were read back from the document, but the first line ended as Normal, not Heading 1. OOXML merged the heading/body before an `apply_edit_plan` search failed with `SearchStringInvalidOrTooLong` on a retained 312-character ASCII find. Later recovery restored separate text paragraphs, but final style readback was Normal; the precise style-loss operation is unisolated. Revoke dropped the sidebar to “needs your OK”.

Earlier historical run at `2387154`: heading text and a bulleted list written with readback; Heading 1 style applied when asked; reversal on request removed the text but left empty paragraphs (no real undo). Plugin reload with the sidebar open reconnected by itself; approval was remembered after restarting Word; the next Connect after revoke asked for approval again.

Previous native acceptance used a private test site via `addin_base`; the CA was trusted in the login keychain only. This new bundled loopback route has backend tests but has not been exercised in live Word. Windows desktop Word is not tested. Dark theme, keyboard-only use, Word for the web, the public GitHub Pages origin, and TLS-only trust are not tested.

## Setup

You need Hermes (0.23.0 or newer) and desktop Word (Microsoft 365) on the same computer.
Only desktop Word for Mac has live evidence; Windows instructions below are unverified setup guidance.

1. **Install the plugin** (from the catalog, or by URL):
   `hermes plugins install https://github.com/AtlasOmnia/sidecar-word#plugins/atlas-sidecar-word --enable`
   It asks for `API_SERVER_KEY` if you have none yet: paste a long random value (for example the output of
   `openssl rand -hex 32`). Hermes stores it in its own `.env`; the plugin never stores or shows it.
2. **Enable the Hermes API server**: add `API_SERVER_ENABLED=true` to `~/.hermes/.env` (the profile you use), then
   start or restart the gateway: `hermes gateway`. The plugin's bridge starts with it.
3. **Trust the local bridge CA once**: review and run the single platform-specific command printed when the bridge
   starts (also shown by `/sidecar-word trust` or `hermes sidecar-word trust`). See [Local TLS trust](#local-tls-trust).
4. **Load the sidebar into Word**: after the bridge has started, run `/sidecar-word manifest` in Hermes or `hermes sidecar-word manifest` in a terminal. The plugin writes a manifest for its bundled sidebar using the actual bound port and prints the Mac sideload copy instruction; no hosted download is needed. To use a self-hosted site, set `addin_base` first (see [Settings](#settings)). Then in Word:
   Home > Add-ins > Sidecar for Word.
5. **One click**: press **Connect to Hermes** in the sidebar. The sidebar shows a code. In any Hermes chat run
   `/sidecar-word` and choose **Allow** only if the code matches the one in your sidebar. Done; ask away.

Hermes remembers the approval for this sidebar (it survives Word and Hermes restarts). Take it back any time with
`/sidecar-word revoke` (or `hermes sidecar-word revoke` in a terminal); the sidebar then has to ask again.

## Local TLS trust

A WKWebView probe of the HTTPS sidebar failed against both HTTP loopback names. The plugin instead serves only
`https://localhost:47630-47634`, bound to IPv4 `127.0.0.1`. There is no HTTP/numeric-host fallback and certificate
verification remains enabled. Python dependency: **`cryptography>=43`**, declared in `plugin.yaml`; ensure it is
installed in the Python environment running Hermes if your host does not install declared plugin dependencies.
No external OpenSSL command or certificate download is needed.

On first bridge start, a unique local CA and localhost leaf certificate are generated in
`<owning Hermes home>/plugin-data/atlas-sidecar-word/tls/`. The plugin prints **one command for your platform**
with the actual CA path. Review it; **the plugin never executes it**, requests elevation, or changes trust itself.
For example (replace the placeholder with the actual printed path):

* macOS, your login keychain, **not** the system keychain; no `sudo`:
  `security add-trusted-cert -r trustRoot -k "$HOME/Library/Keychains/login.keychain-db" '<plugin-data>/tls/ca.pem'`
* Windows, PowerShell, your current user's root store, **not** Local Machine:
  `certutil -user -addstore Root '<plugin-data>\tls\ca.pem'`

**Disclosure:** For TLS it can only vouch for localhost. The trust command does not limit the certificate to TLS, so protect the key file and remove the trust when you stop using Sidecar. Its critical name constraints permit the DNS `localhost`
subtree and exactly IP `127.0.0.1/32` (no IPv6). X.509 DNS subtree semantics also permit names below
`localhost`; the server leaf has only DNS `localhost`. It cannot vouch for unrelated sites such as `example.com`
in conforming verifiers. Protect the plugin data directory and retained CA key:
TLS keys are created with mode 0600 and the TLS directory with 0700 on POSIX; on Windows protect the owning
user's directory with its user-only ACL. Never share or commit `ca-key.pem` or `localhost-key.pem`. Only the public
`ca.pem` is used for the manual trust command; no API credential is included. TLS trust does **not** approve a
sidebar: the code/Allow flow is still required.

The CA survives restarts (at most 825 days, path length 0, critical certificate/CRL-signing usage only);
the server-auth, non-CA localhost leaf lasts at most 397 days (currently one year) and is renewed on bridge start
within 30 days of expiry using the same CA, without another trust step. Missing/corrupt identities fail closed.
Legacy unconstrained CAs, incorrect constraints, or overlong CA validity are refused without replacement.
A one-time instruction per runtime says to remove the old CA trust first, then delete the TLS directory;
the plugin never silently reuses or rotates that CA. For recovery, stop this plugin's bridge, remove its matching CA from your login
keychain (Keychain Access) or Current User > Trusted Root Certification Authorities (Windows certificate manager),
then remove only this plugin's `tls` directory and start it again. Review/trust the newly printed CA. Uninstalling
the plugin or `/sidecar-word revoke` does **not** remove OS certificate trust; remove the matching CA manually.
On macOS remove trust in **Keychain Access > login > Certificates > Sidecar for Word local CA … > Delete**
(older installs: **Sidecar local CA …**). Compare with `openssl x509 -noout -fingerprint -in <tls>/ca.pem`,
then delete only this plugin's `tls` folder (path above). Re-trust the printed replacement if starting again.
Plugin reload with the sidebar open reconnected by itself in the Mac test; reopen the sidebar if it does not reconnect. Reload Word after changing trust. Until trusted, a browser may report only `TypeError: Load failed`; do not
disable verification or revert to HTTP. Managed devices may require your administrator's approval of user CAs.

## Load the manifest (sideload)

The plugin bundles the default manifest and writes a copy rendered for the current `addin_base` and actual bound port into its plugin data directory. Run `/sidecar-word manifest` in Hermes or `hermes sidecar-word manifest` in a terminal after the bridge is running; it prints the manifest path and Mac copy instruction. For an optional self-hosted build, use `npm run build:static -- --base <site>/` to create its manifest.

**macOS** (Word, Excel and PowerPoint read add-in manifests from their own `wef` folder):

```sh
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
cp /path/printed/by/the/manifest/command/manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/atlas-sidecar-word.xml
```

Quit and reopen Word, then Home > Add-ins (or Insert > Add-ins > My Add-ins) > Developer Add-ins > Sidecar for
Word. To remove it, delete that file.

**Windows (not tested)** — intended per-user developer sideload:

```bat
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v AtlasSidecarWord /t REG_SZ /d "C:\path\to\manifest.xml" /f
```

Quit and reopen Word, then Home > Add-ins > Sidecar for Word. To remove it:
`reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v AtlasSidecarWord /f`.
Alternative: put `manifest.xml` in a shared folder and add it under File > Options > Trust Center > Trust Center
Settings > Trusted Add-in Catalogs (Show in Menu), then Home > Add-ins > More Add-ins > Shared Folder.

## What Hermes can do in Word

`sidecar_word_get_document_overview`, `sidecar_word_read_selection`, `sidecar_word_read_document_text`,
`sidecar_word_get_document_structure` (reads); `sidecar_word_apply_edit_plan`, `sidecar_word_insert_text`,
`sidecar_word_replace_selection`, `sidecar_word_find_replace`, `sidecar_word_format_selection`, `sidecar_word_apply_style`,
`sidecar_word_insert_ooxml`, `sidecar_word_insert_comment` (edits). Details: [docs/word-tools.md](https://github.com/AtlasOmnia/sidecar-word/blob/main/docs/word-tools.md).

Approving the pane authorizes all twelve known tools, including new headings, paragraphs, lists, rewrites,
formatting and comments. Every call needs a valid, bound sidebar session with active approval; unknown tools,
missing/invalid sessions and revoked approval fail closed. Model arguments cannot supply that authority.
There is no trustworthy turn id at the handler seam, so there is no one-comment-per-turn limit.
The tools act only on the document whose sidebar started the chat; another chat is refused (`ForeignSession`).

There is **no real undo** tool provided by Sidecar; use native Office Undo (verified for a supported short edit only). In contrast, asking Hermes to reverse an edit makes another edit, not an undo transaction.
Keep a backup or saved version before important edits; do not assume automatic rollback.

## Settings

Plugin settings (`plugins.entries.atlas-sidecar-word.settings` in `config.yaml`, or the Desktop Plugins tab):

| Setting | Default | Meaning |
| --- | --- | --- |
| `addin_base` | `https://localhost:47630/sidecar-word/` | HTTPS base URL for the sidebar. The plugin serves its bundled build here by default (using the actual bound port); set a self-hosted HTTPS build to override it. The bridge accepts only this configured origin. The installer can override it. |
| `api_server_url` | `http://127.0.0.1:8642` | Your Hermes API server (loopback only). Chat automatically uses this plugin's own profile: `/p/<profile>/v1/...` for a named profile, `/v1/...` for default. An explicit `/p/<profile>` URL path is accepted only for that same named profile; other paths are rejected. |
| `autostart` | `true` | Start the bridge when the gateway loads the plugin; off: on the first `/sidecar-word` or `sidecar_word_*` call. |

## Use your own sidebar site

Build with `npm ci && npm run build:static -- --base https://sidebar.example.org/word/` and serve the contents of `dist/` on that HTTPS site. Set the owning Hermes profile's exact config key:

```sh
hermes config set plugins.entries.atlas-sidecar-word.settings.addin_base https://sidebar.example.org/word/
```

Or set `addin_base` under `plugins.entries.atlas-sidecar-word.settings` in that profile's `config.yaml` / Desktop Plugins tab. Restart the owning gateway after the change and sideload the new site's `manifest.xml`; the manifest build base and this setting must agree. The plugin reads `ctx.get_config("addin_base", ...)` in `DeskRuntime.from_ctx` (`plugins/atlas-sidecar-word/desk_bridge.py`). A profile-selecting invocation can use `hermes --profile <profile> config set ...` instead.

## Safety, in short

* The HTTPS bridge binds `127.0.0.1` only, on ports 47630-47634 (first free), requires Host `localhost:<actual-port>`, and
  answers only the sidebar site's exact origin.
* A sidebar gets a session only after you allow it in Hermes. Hermes stores a SHA-256 hash of that sidebar's
  grant (never the grant) in `<hermes home>/plugin-data/atlas-sidecar-word/approvals.json` (owner-only); the grant
  stays in the sidebar's browser storage. `/sidecar-word revoke` deletes it and closes open sidebars at their next
  request.
* No shared API key file. The sidebar never sees an API key. The plugin reads `API_SERVER_KEY` through
  Hermes (`agent.secret_scope.get_secret`) when it forwards a chat request, keeps it in memory, and never writes
  or logs it.
* Allowing a sidebar lets the site at `<origin>` use your Hermes agent and all its tools until you revoke. Use a site you control or trust; open only documents you trust. Here `<origin>` is the whole configured HTTPS origin (default `https://localhost:<actual-port>`), not just `/sidecar-word/`; any other page on that same origin shares this authority.
* No Sidecar tool can approve. Approval is the `/sidecar-word` command or `hermes sidecar-word allow`; an agent that can run shell commands could run the latter, subject to Hermes' own command approval.
* Logs carry metadata only (route, status, time), never document text, prompts, grants or keys.

## Disclosures

* **Background listener:** an HTTPS listener at `https://localhost:47630-47634`, bound to `127.0.0.1`, in the process that runs the Hermes
  gateway (or the plugin host), from when the gateway loads the plugin (or the first `/sidecar-word` / tool call when
  `autostart` is off) until Hermes exits. It does not overlap Workbench (47615-47619) or the other Sidecar
  plugins (Excel 47640-47644, PowerPoint 47650-47654).
* **Network:** forwards the sidebar's chat requests to your Hermes API server on loopback (`api_server_url`), with
  `API_SERVER_KEY` read through Hermes. No other network calls, no telemetry, no self-updater.
* **Files:** `plugin-data/atlas-sidecar-word/approvals.json` and `pending.json` (hashes and metadata), plus
  `tls/ca.pem`, `tls/localhost.pem` (public certificates), `tls/ca-key.pem` and `tls/localhost-key.pem` (private keys).
* **The sidebar page** is bundled with the plugin and served by its HTTPS loopback bridge at `https://localhost:<actual-port>/sidecar-word/`; it loads Microsoft's
  Office.js from `appsforoffice.microsoft.com` and sends document data only to the HTTPS localhost bridge. The build writes `SHA256SUMS`. A custom `addin_base` may still select a self-hosted site, whose build is outside this plugin's catalog pin.
* **Same author, separate lineage:** Atlas Omnia also publishes Workbench (Office for the web inside Hermes
  Desktop). This plugin is the desktop-Office sidebar; its tools use their own `sidecar_word_` prefix so both can be
  installed together.

## Limits

* Retained evidence identifies the failed search as 312 ASCII characters, beyond Word’s 255-character search bound. OOXML had already merged the heading/body before that failure; final Heading 1 was Normal, but the exact operation losing style remains unisolated. The bounded search/style repair at d4d0de9 passed the native checks described above: source and encoded queries are capped at 255 UTF-16 units, replacement capacity stays 5000, actual styles are read back, and failed searches must stop destructive recovery. Guidance is not a hard execution fence; use native Office Undo for recovery after uncertain mutations.
* Picture-only paragraph reuse shares its first inserted text with the picture. Arbitrary rich formatting, fields/content controls, comprehensive native Unicode/encoded-boundary coverage and 5000-character native replacement remain unqualified. Upstream Hermes schema rejection may echo arguments. Native short-edit Undo is not whole-turn grouping or proof for all operations.
* Reversing an insert can leave empty paragraphs behind.
* Stop ends the turn; an edit already requested may still land, and nothing is rolled back.
* Planned TLS hardening: TLS-only trust (`-p ssl`) and a server-auth EKU on the CA. Neither is applied by the current trust command.
* Approval protects against web pages, not other programs on the same computer. The public Pages origin is account-wide; keep other pages on it free of untrusted scripts/content.
* One profile's bridge per port; with several Hermes profiles running this plugin, the sidebar connects to the
  first one that answers.
* The sidebar must reach `https://localhost` from the HTTPS page, with the local CA manually trusted for the
  current user. Local network policies may still block it; successful backend TLS tests are not live Office acceptance.
* Large documents: reads are bounded (see docs/word-tools.md); a turn has a time limit and a 30-second checkpoint.

## Publication checklist (not executed)

Nothing is published by this checkout. At publication, keep the development history private: prepare a fresh-root public snapshot with a GitHub noreply author identity, rather than pushing this branch's personal-data-bearing history. Do not rewrite this private branch's history. Review the snapshot, create `AtlasOmnia/sidecar-word`, and fill the catalog `sha` with its reviewed 40-hex public commit (the publish-time placeholder is not an installable pin).

Enable GitHub Pages with GitHub Actions and let `.github/workflows/pages.yml` build/deploy the public site. Verify `https://atlasomnia.github.io/sidecar-word/manifest.xml`, `taskpane.html`, and `SHA256SUMS`, then run a clean-profile public-origin Connect, Allow, edit/readback and revoke check before claiming that origin works. These publication actions and live checks are deliberately not run tonight.

## Tests

```sh
npm ci
npm test            # vitest (sidebar) + pytest (plugin: scripts/tests)
hermes plugins validate plugins/atlas-sidecar-word
```

## Layout

* `plugins/atlas-sidecar-word/` — the Hermes plugin (what the catalog installs): `__init__.py` (tools, commands),
  `desk_bridge.py` (bridge, approvals, chat forwarding; same design across the three Sidecar plugins, not identical files), `desk_app.py`,
  `companion/` (sidebar ticket protocol), `bridge_client.py`, `review_policy.py`, `tool_schemas.json`.
* `src/` — the sidebar (TypeScript); `manifest.template.xml`; `scripts/build-static.mjs`.
* `catalog/atlas-sidecar-word.yaml` — draft catalog entry (not submitted).

License: Apache-2.0 (see LICENSE and NOTICE). Microsoft, Office and Word are trademarks of Microsoft; no
affiliation or endorsement is implied. Hermes is a Nous Research project.

## Trademarks

Sidecar is developed by Atlas Omnia. It is not affiliated with or endorsed by Microsoft or Nous Research. Microsoft, Excel, Word, and PowerPoint are trademarks of the Microsoft group of companies.
