<!-- SPDX-License-Identifier: Apache-2.0 -->
# Sidecar for PowerPoint (1.1.0)

## Judge quickstart — PowerPoint only (Mac)

Publication candidate: public installation and public-origin smoke are still pending.
1. After publication, install `hermes plugins install https://github.com/AtlasOmnia/sidecar-powerpoint#plugins/atlas-sidecar-powerpoint --enable`; enable the Hermes API server and start the gateway as described in **Setup** below.
2. Review and run this install's printed local-CA trust command, download `https://atlasomnia.github.io/sidecar-powerpoint/manifest.xml`, and sideload it into desktop PowerPoint using the instructions below.
3. Open a blank synthetic file, click **Connect to Hermes**, compare the code in Hermes, then **Allow**. Allow grants the configured origin access to your **Hermes agent and all its tools until revoked**, not per-edit approval.
4. Ask: “Set slide 1 title to Sidecar Demo and subtitle to Synthetic quarterly update.” Then request an independent document readback. Use native Office Undo where supported; Stop does not roll back edits. Revoke with `/sidecar-powerpoint revoke` when done; removing CA trust is a separate manual step.
Windows desktop Office is untested. GitHub Pages authority covers the whole configured origin, not just this app's path.


**Sidecar by Atlas Omnia.** Your Hermes agent, in the sidebar of Microsoft Excel, Word, and PowerPoint.

Sidecar puts your Hermes agent in a side panel next to the file you're working on. Ask it to draft, edit, analyze, or build without leaving your document, spreadsheet, or deck.

Sidecar is a Hermes plugin plus an Office add-in that opens your agent in the sidebar of Microsoft Excel, Word, and PowerPoint. It works with the file you have open, so you can ask for a formula, a rewrite, or a slide outline and get it in place.

In PowerPoint, Sidecar can:

- read the presentation overview, the selected slides, all slides, or the text on one slide;
- add slides, including a title-and-bullets slide with a safe default layout;
- add text boxes and shapes, replace the selected text, find and replace text on a slide;
- format text in a shape (bold, italic, color, size, font);
- reorder slides, delete a slide (never the last one), list layouts and apply a layout to a slide.

After one manual local-CA trust step, click Connect to Hermes and approve the sidebar once in Hermes.
That approval authorizes the configured origin to use your Hermes agent and all its tools until revoked; it is not a confirmation for each edit.
The bridge runs locally; your configured Hermes model/provider may receive presentation context.
No API keys are stored by the plugin (TLS private keys are protected in plugin data).

Your agent: it runs on your own Hermes setup, with the models and tools you already use.

Chat with your own Hermes agent in a sidebar inside **desktop PowerPoint for Mac** while Hermes reads,
reviews and edits the open presentation.

The sidebar is a small static web page that PowerPoint loads. Everything else runs inside the Hermes plugin: it opens a
loopback bridge on your computer, forwards the sidebar's chat to your Hermes API server, and gives Hermes fifteen
`sidecar_powerpoint_*` tools that act on the open presentation through the sidebar. No separate service, no proxy to start,
no Microsoft app registration. HTTPS loopback requires one explicit user trust step for a per-install local CA.

## Status

Live-tested in desktop PowerPoint for Mac at `315f1f0`, 2026-11-11 about 00:32:
Connect and Allow; one prompt changed slide 1's title and subtitle, verified by readback
from the deck, and produced exactly one request to Hermes. After a plugin reload with the
sidebar open, it showed “Lost the connection to Hermes on this computer.” with a Try again
button; Try again reconnected. Revoke dropped the sidebar to “needs your OK”.

In the earlier run at `3682e5f`, about 00:01 the same night (behaviour unchanged since),
one prompt changed slide 1's title and subtitle and added a second slide with three bullets,
verified by readback. The request stayed open for about 24 seconds with no timeout or repeat. After a plugin reload in that run the sidebar showed “Not connected” and did not reconnect within 10 seconds; reopening it connected.
Reversal on request restored the deck (no real undo). Approval was remembered after restarting
PowerPoint; the next Connect after revoke asked for approval again.

PowerPoint needs Try again (or reopen the sidebar) after a plugin reload; Excel and Word
reconnect by themselves. Both runs used a private test site via the `addin_base` setting,
with the CA trusted in the login keychain only.
Windows desktop PowerPoint is untested. Dark theme, keyboard-only use, PowerPoint for the web
and TLS-only trust were not tested. The public GitHub Pages origin is first exercised at publication.

## Setup

You need Hermes (0.23.0 or newer) and desktop PowerPoint (Microsoft 365 for Mac; Windows is an untested target) on the same computer.

1. **Install the plugin** (from the catalog, or by URL):
   `hermes plugins install https://github.com/AtlasOmnia/sidecar-powerpoint#plugins/atlas-sidecar-powerpoint --enable`
   It asks for `API_SERVER_KEY` if you have none yet: paste a long random value (for example the output of
   `openssl rand -hex 32`). Hermes stores it in its own `.env`; the plugin never stores or shows it.
2. **Enable the Hermes API server**: add `API_SERVER_ENABLED=true` to `~/.hermes/.env` (the profile you use), then
   start or restart the gateway: `hermes gateway`. The plugin's bridge starts with it.
3. **Trust the local bridge CA once**: review and run the single platform-specific command printed when the bridge
   starts (also shown by `/sidecar-powerpoint trust` or `hermes sidecar-powerpoint trust`; plain status does not repeat it). See [Local TLS trust](#local-tls-trust).
4. **Load the sidebar into PowerPoint**: download [manifest.xml](https://atlasomnia.github.io/sidecar-powerpoint/manifest.xml) from the sidebar site (the plugin setting `addin_base`,
   see [Settings](#settings)) and sideload it, see [Load the manifest](#load-the-manifest-sideload). Then in PowerPoint:
   Home > Add-ins > Sidecar for PowerPoint.
5. **One click**: press **Connect to Hermes** in the sidebar. The sidebar shows a code. In any Hermes chat run
   `/sidecar-powerpoint` and choose **Allow** (Hermes shows the same code). Allow only if the code matches the one in your sidebar. Then ask away.

Hermes remembers the approval for this sidebar (tested after restarting PowerPoint; Hermes restart persistence is implemented but not live-tested here). Take it back any time with
`/sidecar-powerpoint revoke` (or `hermes sidecar-powerpoint revoke` in a terminal); the sidebar then has to ask again.

## Local TLS trust

A WKWebView probe of the HTTPS sidebar failed against both HTTP loopback names. The plugin instead serves only
`https://localhost:47650-47654`, bound to IPv4 `127.0.0.1`. There is no HTTP/numeric-host fallback and certificate
verification remains enabled. Python dependency: **`cryptography>=43`**, declared in `plugin.yaml`; ensure it is
installed in the Python environment running Hermes if your host does not install declared plugin dependencies.
No external OpenSSL command or certificate download is needed.

On first bridge start, a unique local CA and localhost leaf certificate are generated in
`<owning Hermes home>/plugin-data/atlas-sidecar-powerpoint/tls/`. The plugin prints **one command for your platform**
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
the plugin or `/sidecar-powerpoint revoke` does **not** remove OS certificate trust; remove the matching CA manually.
To remove trust on Mac: Keychain Access > login > Certificates > “Sidecar for PowerPoint local CA ...” (older installs: “Sidecar local CA ...”) > Delete,
then delete only this plugin's `tls` folder after stopping the bridge. Uninstall/revoke alone does not remove trust.
Compare with `openssl x509 -noout -fingerprint -in <plugin-data>/tls/ca.pem` before removing trust.
TLS-only trust (`-p ssl`) and a server-auth EKU on the CA are planned hardening; neither is implemented yet.
Reload PowerPoint after changing trust. Until trusted, a browser may report only `TypeError: Load failed`; do not
disable verification or revert to HTTP. Managed devices may require your administrator's approval of user CAs.

## Load the manifest (sideload)

The manifest is built for one sidebar site (`npm run build:static -- --base <site>/` writes `dist/manifest.xml`).

**macOS** (Word, Excel and PowerPoint read add-in manifests from their own `wef` folder):

```sh
mkdir -p ~/Library/Containers/com.microsoft.PowerPoint/Data/Documents/wef
cp manifest.xml ~/Library/Containers/com.microsoft.PowerPoint/Data/Documents/wef/atlas-sidecar-powerpoint.xml
```

Quit and reopen PowerPoint, then Home > Add-ins (or Insert > Add-ins > My Add-ins) > Developer Add-ins > Sidecar for
PowerPoint. To remove it, delete that file.

**Windows — untested guidance** (per-user developer sideload, no admin rights):

```bat
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v AtlasSidecarPowerPoint /t REG_SZ /d "C:\path\to\manifest.xml" /f
```

Quit and reopen PowerPoint, then Home > Add-ins > Sidecar for PowerPoint. To remove it:
`reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v AtlasSidecarPowerPoint /f`.
Alternative: put `manifest.xml` in a shared folder and add it under File > Options > Trust Center > Trust Center
Settings > Trusted Add-in Catalogs (Show in Menu), then Home > Add-ins > More Add-ins > Shared Folder.

## What Hermes can do in PowerPoint

`sidecar_powerpoint_get_presentation_overview`, `sidecar_powerpoint_read_selected_slides`, `sidecar_powerpoint_read_all_slides`, `sidecar_powerpoint_create_slide`, `sidecar_powerpoint_create_text_slide`, `sidecar_powerpoint_add_text_box`, `sidecar_powerpoint_update_slide_text`, `sidecar_powerpoint_add_shape`, `sidecar_powerpoint_read_slide_content`, `sidecar_powerpoint_delete_slide`, `sidecar_powerpoint_reorder_slide`, `sidecar_powerpoint_find_and_replace_text`, `sidecar_powerpoint_list_slide_layouts`, `sidecar_powerpoint_set_slide_layout`, `sidecar_powerpoint_format_text`. Details: [tool reference](https://github.com/AtlasOmnia/sidecar-powerpoint/blob/main/docs/powerpoint-tools.md).

The tools act only on the presentation whose sidebar started the chat; another chat is refused (`ForeignSession`). Speaker notes, slide duplication and image insertion are not advertised: stable Office.js cannot execute those operations here.

## Settings

Plugin settings (`plugins.entries.atlas-sidecar-powerpoint.settings` in `config.yaml`, or the Desktop Plugins tab):

| Setting | Default | Meaning |
| --- | --- | --- |
| `addin_base` | `https://atlasomnia.github.io/sidecar-powerpoint/` | HTTPS sidebar site; build the manifest for the same base. The bridge accepts only its origin (the whole host, not its path). Public origin first exercised at publication. |
| `api_server_url` | `http://127.0.0.1:8642` | Your Hermes API server (loopback only). Chat uses `/p/<profile>/v1/chat/completions` for the owning named profile, or `/v1/chat/completions` for default. An explicit `/p/<profile>` base path is allowed only for the same profile. |
| `autostart` | `true` | Start the bridge when the gateway loads the plugin; off: on the first `/sidecar-powerpoint` or `sidecar_powerpoint_*` call. |

## Use your own sidebar site

Set the active profile's setting to an HTTPS URL ending in `/`:

```sh
hermes config set plugins.entries.atlas-sidecar-powerpoint.settings.addin_base https://sidebar.example.org/powerpoint/
npm ci
npm run build:static -- --base https://sidebar.example.org/powerpoint/
```

Host `dist/` unchanged at that URL, including `manifest.xml`, `SHA256SUMS` and bundled license files.
Reload the plugin/gateway, sideload that site's manifest and reopen the sidebar. The plugin reads
`addin_base` through `ctx.get_config("addin_base", ...)`, not an environment variable or slash command.
Use a dedicated origin you control: approval covers the whole host. The default GitHub Pages origin
is shared with other AtlasOmnia Pages sites, not isolated by the `/sidecar-powerpoint/` path.

## Safety, in short

Allowing a sidebar lets the site at the configured origin use your Hermes agent and all its tools until you revoke.
Use a site you control or trust; open only documents you trust. The Allow question names that origin.

* The HTTPS bridge binds `127.0.0.1` only, on ports 47650-47654 (first free), requires Host `localhost:<actual-port>`, and
  answers only the sidebar site's exact origin.
* A sidebar gets a session only after you allow it in Hermes. Hermes stores a SHA-256 hash of that sidebar's
  grant (never the grant) in `<hermes home>/plugin-data/atlas-sidecar-powerpoint/approvals.json` (owner-only); the grant
  stays in the sidebar's browser storage. `/sidecar-powerpoint revoke` deletes it and closes open sidebars at their next
  request.
* No shared API key file. The sidebar never sees an API key. The plugin reads `API_SERVER_KEY` through
  Hermes (`agent.secret_scope.get_secret`) when it forwards a chat request, keeps it in memory, and never writes
  or logs it.
* No Sidecar tool can approve. Approval is the `/sidecar-powerpoint` command or `hermes sidecar-powerpoint allow`; an agent that can run shell commands could run the latter, subject to Hermes' own command approval.
* Logs carry metadata only (route, status, time), never document text, prompts, grants or keys.

## Disclosures

* **Background listener:** an HTTPS listener at `https://localhost:47650-47654`, bound to `127.0.0.1`, in the process that runs the Hermes
  gateway (or the plugin host), from when the gateway loads the plugin (or the first `/sidecar-powerpoint` / tool call when
  `autostart` is off) until the plugin stops or Hermes exits. It does not overlap Workbench (47615-47619) or the other Sidecar
  plugins (Word 47630-47634, Excel 47640-47644).
* **Network:** forwards the sidebar's chat requests to your Hermes API server on loopback (`api_server_url`), with
  `API_SERVER_KEY` read through Hermes. The bridge has no other outbound service calls, telemetry or self-updater.
  Hermes itself uses your configured model/provider and tools; local bridge transport does not imply local-only inference.
* **Files:** `plugin-data/atlas-sidecar-powerpoint/approvals.json` and `pending.json` (hashes and metadata), plus
  `tls/ca.pem`, `tls/localhost.pem` (public certificates), `tls/ca-key.pem` and `tls/localhost-key.pem` (private keys).
* **The sidebar page** is served from `addin_base`, outside this plugin's catalog pin; it loads Microsoft's
  Office.js from `appsforoffice.microsoft.com` and sends presentation data only to the HTTPS localhost bridge. Build it yourself at
  the pinned commit with `npm ci && npm run build:static -- --base <your site>/`; the build writes `SHA256SUMS`.
* **Same author, separate lineage:** Atlas Omnia also publishes Workbench (Office for the web inside Hermes
  Desktop). This plugin is the desktop-Office sidebar; its tools use their own `sidecar_powerpoint_` prefix so both can be
  installed together.

## Limits

* After a plugin reload, use Try again or reopen the sidebar to reconnect. After an update, close and reopen the sidebar to load the current pane.
* There is no real undo: asking Hermes to reverse an edit makes it edit again, not restore an Office undo transaction.
  Save a copy before important changes; stopping a turn does not roll back edits already made.
* Native Hermes turns have a 300-second budget, with progress every 5 seconds and a live 30-second Stop / Continue
  checkpoint. Continue resets the remaining budget to 300 seconds from that decision; an unanswered checkpoint
  remains bounded by the current deadline. Native failures are not automatically retried.

* Tested live only in desktop PowerPoint for Mac. Windows desktop PowerPoint is untested; PowerPoint for the web is untested.
* One profile's bridge per port; with several Hermes profiles running this plugin, the sidebar connects to the
  first one that answers.
* The sidebar must reach `https://localhost` from the HTTPS page, with the local CA manually trusted for the
  current user. Local network policies may still block it; successful backend TLS tests are not live Office acceptance.
* Large presentations: reads are bounded (see docs/powerpoint-tools.md); a turn has a time limit and a 30-second checkpoint.

## Tests

```sh
npm ci
npm test            # vitest (sidebar) + pytest (plugin: scripts/tests)
hermes plugins validate plugins/atlas-sidecar-powerpoint
```

## Layout

* `plugins/atlas-sidecar-powerpoint/` — the Hermes plugin (what the catalog installs): `__init__.py` (tools, commands),
  `desk_bridge.py` (bridge, approvals, chat forwarding; shared by the three Sidecar plugins), `desk_app.py`,
  `companion/` (sidebar ticket protocol), `bridge_client.py`, `tool_schemas.json`.
* `src/` — the sidebar (TypeScript); `manifest.template.xml`; `scripts/build-static.mjs`.
* `catalog/atlas-sidecar-powerpoint.yaml` — draft catalog entry (not submitted).

License: Apache-2.0 (see LICENSE and NOTICE). Microsoft, Office and PowerPoint are trademarks of Microsoft; no
affiliation or endorsement is implied. Hermes is a Nous Research project.

## Trademarks

Sidecar is developed by Atlas Omnia. It is not affiliated with or endorsed by Microsoft or Nous Research. Microsoft, Excel, Word, and PowerPoint are trademarks of the Microsoft group of companies.
