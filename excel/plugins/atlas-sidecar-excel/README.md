<!-- SPDX-License-Identifier: Apache-2.0 -->
# Sidecar for Excel (1.1.0)

## Judge quickstart — Excel only (Mac)

Publication candidate: public installation and public-origin smoke are still pending.
1. After publication, install `hermes plugins install https://github.com/AtlasOmnia/sidecar-excel#plugins/atlas-sidecar-excel --enable`; enable the Hermes API server and start the gateway as described in **Setup** below.
2. Review and run this install's printed local-CA trust command, then run `/sidecar-excel manifest` (or `hermes sidecar-excel manifest`) and follow the printed Mac sideload instruction. No hosted sidebar is required.
3. Open a blank synthetic file, click **Connect to Hermes**, compare the code in Hermes, then **Allow**. Allow grants the configured origin access to your **Hermes agent and all its tools until revoked**, not per-edit approval.
4. Ask: “Write Fruit/Price and Apple/2 in A1:B2, then put =SUM(B2) in B3.” Then request an independent document readback. Use native Office Undo where supported; Stop does not roll back edits. Revoke with `/sidecar-excel revoke` when done; removing CA trust is a separate manual step.
Windows desktop Office is untested. A custom HTTPS override grants the whole configured origin, not just this app's path.


**Sidecar by Atlas Omnia.** Your Hermes agent, in the sidebar of Microsoft Excel, Word, and PowerPoint.

Sidecar puts your Hermes agent in a side panel next to the file you're working on. Ask it to draft, edit, analyze, or build without leaving your document, spreadsheet, or deck.

Sidecar is a Hermes plugin plus an Office add-in that opens your agent in the sidebar of Microsoft Excel, Word, and PowerPoint. It works with the file you have open, so you can ask for a formula, a rewrite, or a slide outline and get it in place.

In Excel, Sidecar can:

- read ranges (values, formulas, number formats) and get an overview of sheets, selection and used ranges;
- write values and formulas, fill formulas down or across, import CSV data, find and replace, clear ranges;
- format ranges, add conditional formatting and data validation, freeze panes;
- create, rename, copy, activate, delete, protect and unprotect sheets;
- create, sort and filter tables, add charts, and create pivot tables where your Excel supports it through Office.js;
- compute range statistics, recalculate the workbook, add or clear cell comments (notes);
- inspect or run one bounded multi-step workbook plan at a time.

After one manual local-CA trust step, click Connect to Hermes and allow the sidebar once in Hermes.
That approval permits the configured origin to use your Hermes agent and all its tools until revoked; it is not a confirmation prompt for every edit. The bridge runs on your
machine; Hermes may send chat and workbook data to your configured model provider. No API keys are stored by the
plugin (TLS private keys are protected in plugin data).

Your agent: it runs on your own Hermes setup, with the models and tools you already use.

Chat with your own Hermes agent in a sidebar inside **desktop Excel for Mac** while Hermes reads,
reviews and edits the open workbook.

The sidebar is a small static web page served by the plugin's HTTPS loopback bridge at
`https://localhost:<actual-port>/sidecar-excel/`. Everything else runs inside the Hermes plugin: it opens a
loopback bridge on your computer, forwards the sidebar's chat to your Hermes API server, and gives Hermes 28
`sidecar_excel_*` tools that act on the open workbook through the sidebar. No separate service, no proxy to start,
no Microsoft app registration. HTTPS loopback requires one explicit user trust step for a per-install local CA.

Status — latest live run: 2026-11-11 about 00:26 in desktop Excel for Mac, plugin and sidebar built from
`a7b3fd7`. Connect and Allow succeeded; a 3-row table with header in A1:B4 and `=SUM(B2:B4)` in B5
were read back from the workbook. Revoke dropped the sidebar to “needs your OK”. The bridge answered
200 to the configured origin and 403 to another origin. The sidebar used a private test site via `addin_base`;
the CA was trusted in the login keychain, and this Mac also has it in the System keychain.

Earlier the same night at `35ef09e` (behaviour unchanged since): reversal on request by clearing the range
(no real undo), plugin reload with the sidebar open (it reconnected by itself and edited again), approval
remembered after restarting Excel, and the next Connect after revoke asked for approval again.
Windows desktop Excel is untested. Also not tested: dark theme, keyboard-only use, the optional public Pages override,
TLS-only trust, Excel for the web, or every tool/workbook plan. Reopen the sidebar if it does not reconnect.
Live-tested with the sidebar served from a private test site before the bundled same-origin default; the optional Pages override has not been exercised.

## Setup

Publication is pending; the owner’s [publish runbook](https://github.com/AtlasOmnia/sidecar-excel/blob/main/docs/publish-runbook.md) covers the fresh-history export,
optional Pages deployment and public-origin acceptance run. Pages hosting is not required for the default local sidebar.

You need Hermes (0.23.0 or newer) and desktop Excel on the same computer. Only desktop Excel for Mac has been live-tested.

1. **Install the plugin** (the repository is prepared for publication, not yet published):
   `hermes plugins install https://github.com/AtlasOmnia/sidecar-excel#plugins/atlas-sidecar-excel --enable`
   It asks for `API_SERVER_KEY` if you have none yet: paste a long random value (for example the output of
   `openssl rand -hex 32`). Hermes stores it in its own `.env`; the plugin never stores or shows it.
2. **Enable the Hermes API server**: add `API_SERVER_ENABLED=true` to `~/.hermes/.env` (the profile you use), then
   start or restart the gateway: `hermes gateway`. The plugin's bridge starts with it.
3. **Trust the local bridge CA once**: review and run the single platform-specific command printed when the bridge
   starts (request it again with `/sidecar-excel trust` or `hermes sidecar-excel trust`; ordinary status is quiet). See [Local TLS trust](#local-tls-trust).
4. **Load the sidebar into Excel**: run `/sidecar-excel manifest` in Hermes or `hermes sidecar-excel manifest` in a terminal.
   It writes a manifest for the actual bridge port into plugin data and prints the Mac copy command; no hosting is required.
   For your own site, build and sideload that site's manifest instead, see [Load the manifest](#load-the-manifest-sideload).
   Then in Excel:
   Home > Add-ins > Sidecar for Excel.
5. **One click**: press **Connect to Hermes** in the sidebar. The sidebar shows a code. In any Hermes chat run
   `/sidecar-excel` and choose **Allow**. Allow only if the code matches the one in your sidebar. Done; ask away.

Hermes remembers the approval for this sidebar (it survives Excel and Hermes restarts). Take it back any time with
`/sidecar-excel revoke` (or `hermes sidecar-excel revoke` in a terminal); the sidebar then has to ask again.

On the tested plugin reload the open sidebar reconnected by itself and edited again. Reopen the sidebar if it
does not reconnect; reopen after a site update to load the current client.
A transport failure shows “Lost the connection to Hermes on this computer.” and **Try again**; it is not itself
evidence that approval was lost. Only missing/revoked approval requires another Allow.

There is **no real undo** supplied by Sidecar. Hermes can reverse its own edit by making another edit, such as
clearing a range; that is not Excel Undo and cannot reliably restore overwritten values or formatting.

## Local TLS trust

A WKWebView probe of the HTTPS sidebar failed against both HTTP loopback names. The plugin instead serves only
`https://localhost:47640-47644`, bound to IPv4 `127.0.0.1`. There is no HTTP/numeric-host fallback and certificate
verification remains enabled. Python dependency: **`cryptography>=43`**, declared in `plugin.yaml`; ensure it is
installed in the Python environment running Hermes if your host does not install declared plugin dependencies.
No external OpenSSL command or certificate download is needed.

On first bridge start, a unique local CA and localhost leaf certificate are generated in
`<owning Hermes home>/plugin-data/atlas-sidecar-excel/tls/`. The plugin prints **one command for your platform**
with the actual CA path. Review it; **the plugin never executes it**, requests elevation, or changes trust itself.
For example (replace the placeholder with the actual printed path):

* macOS, the printed command targets your login keychain; no `sudo`:
  `security add-trusted-cert -r trustRoot -k "$HOME/Library/Keychains/login.keychain-db" '<plugin-data>/tls/ca.pem'`
* Windows (untested), PowerShell, your current user's root store, **not** Local Machine:
  `certutil -user -addstore Root '<plugin-data>\tls\ca.pem'`
  The printed Windows text also shows this CA's SHA-1 thumbprint and the command that later removes exactly
  this trust from the same store (stop the bridge first): `certutil -user -delstore Root '<thumbprint>'`

The Excel test Mac also has the CA in the System keychain; that entry was not removed or retested here.
System-keychain trust was not required in our Word and PowerPoint runs, which connected with login-keychain
trust only. Do not infer a need for system-wide trust from the Excel test.

**Disclosure:** For TLS it can only vouch for localhost. The trust command does not limit the certificate to TLS,
so protect the key file and remove the trust when you stop using Sidecar. Critical name constraints permit the
DNS `localhost` subtree and exactly IP `127.0.0.1/32` (no IPv6); the server leaf has only DNS `localhost`.
Conforming TLS verifiers reject unrelated DNS/IP names, but this is not a restriction on other trust policies.
Planned hardening: TLS-only trust (`-p ssl`) and a server-auth EKU on the CA; neither is implemented yet.
Protect the plugin data directory and retained CA key:
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
keychain: **Keychain Access > login > Certificates > “Sidecar for Excel local CA ...” > Delete**
(existing installs may say “Sidecar local CA ...”). Compare with
`openssl x509 -noout -fingerprint -in '<plugin-data>/tls/ca.pem'` before deleting,
or Current User > Trusted Root Certification Authorities (Windows certificate manager, untested),
then remove only this plugin's `tls` directory and start it again. Review/trust the newly printed CA. Uninstalling
the plugin or `/sidecar-excel revoke` does **not** remove OS certificate trust; remove the matching CA manually.
If the CA was also added to the System keychain, remove that matching entry too.
Delete the plugin's `tls` folder only after removing trust. Reload Excel after changing trust. Until trusted, a browser may report only `TypeError: Load failed`; do not
disable verification or revert to HTTP. Managed devices may require your administrator's approval of user CAs.

## Load the manifest (sideload)

By default, run `/sidecar-excel manifest` or `hermes sidecar-excel manifest` while the gateway is running. The plugin writes
`manifest.xml` for its actual localhost port into plugin data and prints the Mac sideload command. For a custom site,
build a site-specific manifest with `npm run build:static -- --base <site>/` (writes `dist/manifest.xml`).

**macOS** (Word, Excel and PowerPoint read add-in manifests from their own `wef` folder):

```sh
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/atlas-sidecar-excel.xml
```

Quit and reopen Excel, then Home > Add-ins (or Insert > Add-ins > My Add-ins) > Developer Add-ins > Sidecar for
Excel. To remove it, delete that file.

**Windows — untested instructions** (per-user developer sideload, no admin rights):

```bat
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v AtlasSidecarExcel /t REG_SZ /d "C:\path\to\manifest.xml" /f
```

Quit and reopen Excel, then Home > Add-ins > Sidecar for Excel. To remove it:
`reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v AtlasSidecarExcel /f`.
Alternative: put `manifest.xml` in a shared folder and add it under File > Options > Trust Center > Trust Center
Settings > Trusted Add-in Catalogs (Show in Menu), then Home > Add-ins > More Add-ins > Shared Folder.

## What Hermes can do in Excel

`sidecar_excel_get_workbook_overview`, `sidecar_excel_read_range`, `sidecar_excel_summarize_range_stats`, `sidecar_excel_write_range`, `sidecar_excel_format_range`, `sidecar_excel_conditional_format_range`, `sidecar_excel_freeze_panes`, `sidecar_excel_set_data_validation`, `sidecar_excel_create_sheet`, `sidecar_excel_delete_sheet`, `sidecar_excel_activate_sheet`, `sidecar_excel_insert_table`, `sidecar_excel_sort_table`, `sidecar_excel_filter_table`, `sidecar_excel_create_chart`, `sidecar_excel_find_replace`, `sidecar_excel_clear_range`, `sidecar_excel_autofill_formula`, `sidecar_excel_rename_sheet`, `sidecar_excel_copy_sheet`, `sidecar_excel_protect_sheet`, `sidecar_excel_unprotect_sheet`, `sidecar_excel_create_pivot_table`, `sidecar_excel_run_calculation`, `sidecar_excel_add_comment`, `sidecar_excel_clear_comments`, `sidecar_excel_import_csv_to_sheet`, `sidecar_excel_execute_plan`.

The bounded `sidecar_excel_execute_plan` path preserves trusted turn metadata, inspect/execute phase fences,
transaction identity, replay protection, pane confirmation and execution-uncertainty handling (never automatically retry
an uncertain plan). WorkbookPlan rollout remains opt-in in the sidebar settings. Primitive tools retain their authority gates.
Tools target only the workbook whose sidebar started the chat; another chat is refused (`ForeignSession`).
See [docs/excel-tools.md](https://github.com/AtlasOmnia/sidecar-excel/blob/main/docs/excel-tools.md) and [docs/workbook-plan-runtime.md](https://github.com/AtlasOmnia/sidecar-excel/blob/main/docs/workbook-plan-runtime.md).

## Settings

Plugin settings (`plugins.entries.atlas-sidecar-excel.settings` in `config.yaml`, or the Desktop Plugins tab):

| Setting | Default | Meaning |
| --- | --- | --- |
| `addin_base` | `https://localhost:47640/sidecar-excel/` (runtime uses the actual bound port) | Bundled sidebar served by the plugin over the same-origin HTTPS bridge. Optional: set an HTTPS base built and hosted by you; the bridge then accepts exactly that site's origin. |
| `api_server_url` | `http://127.0.0.1:8642` | Your Hermes API server (loopback only). Chat automatically uses `/p/<profile>/v1/chat/completions` for the plugin's own named profile, or `/v1/chat/completions` for default. An explicit `/p/<profile>` base path is allowed only for that same named profile; other paths and profiles are rejected. |
| `autostart` | `true` | Start the bridge when the gateway loads the plugin; off: on the first `/sidecar-excel` or `sidecar_excel_*` call. |

## Use your own sidebar site

Build and host the static output on an HTTPS site you control:

```sh
npm ci
npm run build:static -- --base https://sidebar.example.org/excel/
hermes config set plugins.entries.atlas-sidecar-excel.settings.addin_base https://sidebar.example.org/excel/
```

The exact config key is `plugins.entries.atlas-sidecar-excel.settings.addin_base` in the owning profile's
`config.yaml` (or set `addin_base` in the Desktop Plugins tab). Restart the gateway after changing it;
download that site's `manifest.xml` and sideload it instead of the default loopback manifest. The build base and setting
must match. The bridge accepts exactly one origin, not one URL path: all pages on that host share authority.
The default is served on the plugin's localhost origin. For custom deployments, use a dedicated origin and keep it
free of untrusted scripts/content; the bridge grants authority to that entire origin.

## Safety, in short

* Allowing lets the sidebar at the configured origin (default `https://localhost:<actual-port>`) use your Hermes agent
  and all its tools until you revoke. Use a site you control or trust; open only documents you trust.

* The HTTPS bridge binds `127.0.0.1` only, on ports 47640-47644 (first free), requires Host `localhost:<actual-port>`, and
  answers only the sidebar site's exact origin.
* A sidebar gets a session only after you allow it in Hermes. Hermes stores a SHA-256 hash of that sidebar's
  grant (never the grant) in `<hermes home>/plugin-data/atlas-sidecar-excel/approvals.json` (owner-only); the grant
  stays in the sidebar's browser storage. `/sidecar-excel revoke` deletes it and closes open sidebars at their next
  request.
* No shared API key file. The sidebar never sees an API key. The plugin reads `API_SERVER_KEY` through
  Hermes (`agent.secret_scope.get_secret`) when it forwards a chat request, keeps it in memory, and never writes
  or logs it.
* No Sidecar tool can approve. Approval uses `/sidecar-excel` or `hermes sidecar-excel allow <code>`;
  an agent with a shell tool could run the latter, subject to Hermes' own command approval.
* Logs carry metadata only (route, status, time), never workbook text, prompts, grants or keys.

## Disclosures

* **Background listener:** an HTTPS listener at `https://localhost:47640-47644`, bound to `127.0.0.1`, in the process that runs the Hermes
  gateway (or the plugin host), from when the gateway loads the plugin (or the first `/sidecar-excel` / tool call when
  `autostart` is off) until plugin stop/unload or Hermes exit. CLI status/allow/revoke never start a listener. It does not overlap Workbench (47615-47619) or the other Sidecar
  plugins (Word 47630-47634, PowerPoint 47650-47654).
* **Network:** forwards the sidebar's chat requests to your Hermes API server on loopback (`api_server_url`), with
  `API_SERVER_KEY` read through Hermes. No other network calls, no telemetry, no self-updater.
* **Files:** `plugin-data/atlas-sidecar-excel/approvals.json` and `pending.json` (hashes and metadata), plus
  `tls/ca.pem`, `tls/localhost.pem` (public certificates), `tls/ca-key.pem` and `tls/localhost-key.pem` (private keys).
* **The sidebar page** is served from `addin_base`, outside this plugin's catalog pin; it loads Microsoft's
  Office.js from `appsforoffice.microsoft.com` and sends workbook data only to the HTTPS localhost bridge. Build it yourself at
  the pinned commit with `npm ci && npm run build:static -- --base <your site>/`; the build writes `SHA256SUMS`.
* **Same author, separate lineage:** Atlas Omnia also publishes Workbench (Office for the web inside Hermes
  Desktop). This plugin is the desktop-Office sidebar; its tools use their own `sidecar_excel_` prefix so both can be
  installed together.

## Limits

* Stop ends the turn; an edit already requested may still land, and nothing is rolled back.

* Live-tested only in desktop Excel for Mac for the limited flow listed above. Windows desktop Excel is untested; Excel for the web is untested.
* One profile's bridge per port; with several Hermes profiles running this plugin, the sidebar connects to the
  first one that answers.
* The sidebar must reach `https://localhost` from the HTTPS page, with the local CA manually trusted for the
  current user. Local network policies may still block it; successful backend TLS tests are not live Office acceptance.
* Large workbooks: reads are bounded (see docs/excel-tools.md); a turn has a time limit and a 30-second checkpoint.

## Tests

```sh
npm ci
npm test            # vitest (sidebar) + pytest (plugin: scripts/tests)
hermes plugins validate plugins/atlas-sidecar-excel
```

## Layout

* `plugins/atlas-sidecar-excel/` — the Hermes plugin (what the catalog installs): `__init__.py` (tools, commands),
  `desk_bridge.py` (bridge, approvals, chat forwarding; shared by the three Sidecar plugins), `desk_app.py`,
  `companion/` (sidebar ticket protocol), `bridge_client.py`, `tool_schemas.json`.
* `src/` — the sidebar (TypeScript); `manifest.template.xml`; `scripts/build-static.mjs`.
* `catalog/atlas-sidecar-excel.yaml` — draft catalog entry (not submitted).

License: Apache-2.0 (see LICENSE and NOTICE). Microsoft, Office and Excel are trademarks of Microsoft; no
affiliation or endorsement is implied. Hermes is a Nous Research project.

## Trademarks

Sidecar is developed by Atlas Omnia. It is not affiliated with or endorsed by Microsoft or Nous Research. Microsoft, Excel, Word, and PowerPoint are trademarks of the Microsoft group of companies.
