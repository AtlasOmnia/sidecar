<!-- SPDX-License-Identifier: Apache-2.0 -->
# Sidecar bridge: how the sidebar reaches Hermes

App-neutral bridge design for Sidecar for Word, Excel and PowerPoint (`desk_app.py` holds the app's name,
port range and default site). This Word port uses verified localhost TLS; sibling synchronization is separate.

```
desktop Office sidebar (static page served by the plugin bridge by default; an `addin_base` override may select a self-hosted HTTPS site)
   |  https://localhost:<app port>/<sidecar-command>/   exact Origin, Host localhost:<actual-port>
   v
desk bridge (thread inside the Hermes plugin, 127.0.0.1 only)
   |-- GET/HEAD /<sidecar-command>/*      allowlisted static bundle from sidebar/SHA256SUMS
   |-- GET  /<app>/hello                 identify (service hermes-desk-<app>)
   |-- POST /<app>/connect               file a connect request -> code, requestId, pollSecret
   |-- POST /<app>/connect/status        pending | denied | expired | approved (+ grant, once)
   |-- POST /<app>/session               needs X-Hermes-Desk-Grant -> sessionId, sessionKey
   |-- POST /<app>/heartbeat, GET /<app>/pending, POST /<app>/result, DELETE /<app>/session
   |-- POST /v1/chat/completions         session auth -> Hermes API server (API_SERVER_KEY added here)
   `-- POST /<app>/invoke                sidecar_<app>_* tool calls of this process (in-memory key)
Hermes API server (127.0.0.1:8642) -> agent loop -> sidecar_<app>_* tools -> ticket -> sidebar -> Office.js
```

Historical pre-repair failure (superseded by accepted d4d0de9): tested live in desktop Word for Mac on 2026-11-11 at about 00:29, with the plugin and sidebar built from `7eaf1fc`. Connect and Allow worked. The prompt “Write the line Quarterly Summary as a Heading 1, then one short paragraph about fruit prices” wrote both paragraphs, which were read back from the document, but the first line ended as Normal, not Heading 1. OOXML merged the heading/body before an `apply_edit_plan` search failed with `SearchStringInvalidOrTooLong` on a retained 312-character ASCII find. Later recovery restored separate text paragraphs, but final style readback was Normal; the precise style-loss operation is unisolated. Revoke dropped the sidebar to “needs your OK”.

Earlier the same night at `2387154` (historical run): heading text and a bulleted list written with readback; Heading 1 style applied when asked; reversal on request removed the text but left empty paragraphs (no real undo). Plugin reload with the sidebar open reconnected by itself; approval was remembered after restarting Word; the next Connect after revoke asked for approval again.

The sidebar was served from a private test site via `addin_base`; the CA was trusted in the login keychain only. Windows desktop Word is not tested. Dark theme, keyboard-only use, Word for the web, the public GitHub Pages origin, and TLS-only trust are not tested.

Retained evidence identifies the failed search as 312 ASCII characters, beyond Word’s 255-character search bound. OOXML had already merged the heading/body before that failure; final Heading 1 was Normal, but the exact operation losing style remains unisolated. The accepted d4d0de9 repair passed 429 JavaScript and 213 Python tests and bounded native Word 16.114.4/macOS 27.2 acceptance (exact Heading1/Normal paragraphs, targeted style repair, search boundaries, short-edit native Office Undo, bold/table/image preservation and revoke). Picture-only paragraph sharing, arbitrary rich formatting, fields/content controls, comprehensive native Unicode/encoded-boundary and 5000-character replacement remain unqualified. The repaired bounds are: source and encoded queries are capped at 255 UTF-16 units, replacement capacity stays 5000, actual styles are read back, and failed searches must stop destructive recovery. Guidance is not a hard execution fence; use native Office Undo for recovery after uncertain mutations.

The public GitHub Pages origin remains optional and is first exercised at publication.
The default Word sidebar is bundled with the plugin and served at `https://localhost:<actual-port>/sidecar-word/`; an `addin_base` override may select a self-hosted HTTPS site.
After a reload, reopen the sidebar if it does not reconnect.
There is no real undo tool: Hermes reverses an edit by making another edit.

## Ports

Word 47630-47634, Excel 47640-47644, PowerPoint 47650-47654; the bridge takes the first free port of its range and
the sidebar probes the range with `/<app>/hello`. Workbench uses 47615-47619; nothing overlaps.

## TLS and manual user trust

The bridge binds IPv4 `127.0.0.1` only, serves HTTPS only and accepts exactly one Host header with value
`localhost:<actual-port>`. Numeric Host, another port, missing/duplicate Host, and plaintext HTTP are refused.
There is no insecure retry. The WKWebView diagnostic reproduced `TypeError: Load failed` on both HTTP
loopback names and HTTP 200 on verified HTTPS with fixture-only CA trust. This avoids the observed insecure
transport failure; the opaque fetch error alone does not distinguish mixed-content from platform transport policy.
Neither backend TLS success nor the harness is a claim of live Office acceptance.

`desk_tls.py` creates a unique RSA-2048/SHA-256 CA (at most 825 days, CA path length 0, critical certificate/CRL-signing usage only) and a one-year server-auth leaf
with DNS SAN **localhost only** on first bridge start, not on validation or runtime construction. Dependency:
`cryptography>=43`, declared in `plugin.yaml`. Complete identities are published atomically from a private staging
directory under plugin data; a competing first-run process reuses the winning identity. Existing invalid/incomplete
identities fail closed without rotating the CA. On start the leaf renews within 30 days of expiry under the same CA.

Storage in the owning profile's plugin data directory:

| Path | Contents |
| --- | --- |
| `tls/ca.pem` | Public CA, the only file used for manual trust |
| `tls/ca-key.pem` | CA private key, never log/share/commit |
| `tls/localhost.pem` | Public localhost certificate |
| `tls/localhost-key.pem` | Server private key, never log/share/commit |

All files are created 0600, directory 0700 on POSIX; Windows relies on user-only directory ACLs. Runtime-generated
files are not source artifacts. TLS >=1.2 is required. Accepted sockets handshake in bounded handler threads,
not the accept loop, with the existing request timeout and handler limit; one stalled handshake cannot block discovery.
The plugin's own tool client trusts this install's CA explicitly and still verifies the localhost hostname.
The standalone companion test transport can still use HTTP; the production desk runtime never starts that transport
without TLS and never falls back to it. Upstream Hermes API server HTTP on loopback is unchanged.

On bridge start (also with `/sidecar-word trust` or `hermes sidecar-word trust`; plain status is quiet), the plugin prints one manual command for the host platform. **It never runs the
command, calls an OS trust API, elevates privileges or installs system-wide trust.** Examples with placeholder paths:

* macOS login keychain: `security add-trusted-cert -r trustRoot -k "$HOME/Library/Keychains/login.keychain-db" '<plugin-data>/tls/ca.pem'`
* Windows PowerShell user root: `certutil -user -addstore Root '<plugin-data>\tls\ca.pem'`
* Windows removal of exactly this CA later (the printed text includes its SHA-1 thumbprint; stop the bridge first):
  `certutil -user -delstore Root '<thumbprint>'`

The operator must review and explicitly trust the public CA, then reload Word. For TLS it can only vouch for localhost. The trust command does not limit the certificate to TLS, so protect the key file and remove the trust when you stop using Sidecar. For TLS,
critical name constraints permit DNS `localhost` and IP `127.0.0.1/32`, with no IPv6. DNS subtree semantics also
permit names below `localhost`; the generated non-CA server-auth leaf has only DNS `localhost` and lasts at most
397 days. Conforming verifiers reject unrelated DNS/IP certificates signed by this CA. Protect the keys.
Legacy unconstrained CAs and incorrect constraints/overlong CA validity refuse service without rotation, printing
one recovery instruction per runtime: remove old trust manually before deleting this plugin's TLS directory.
Only then restart and review the newly generated CA trust command.
Trust does not replace the approval protocol below. Uninstall/revoke do not remove keychain/root-store entries.
Revoke prints the exact removal steps for this install's CA (stop the bridge, remove it by SHA-1 fingerprint, then
delete only this plugin's `tls` directory); the plugin still never runs them.
For expired CA or corrupt identity, stop this plugin, manually remove its matching CA from the login keychain or
Windows Current User root store, remove only this plugin's `tls` directory, restart and review/trust the replacement.
On macOS: Keychain Access > login > Certificates > “Sidecar local CA …” > Delete; then delete only
this plugin's `tls` folder. Re-trust the printed replacement if restarting.
No production certificate verification bypass is supported. See [README](../README.md#local-tls-trust).

## Approval

1. The sidebar posts `/<app>/connect`. The bridge stores a pending request (app, origin, a short code, sha256 of the
   request's poll secret) in `plugin-data/<plugin>/pending.json` and answers with the code.
2. The user runs `/sidecar-<app>` in any Hermes chat; Hermes asks Allow or Deny (`ctx.ask_choice`) naming the code.
   `/sidecar-<app> allow <code>`, `/sidecar-<app> deny <code>` and `hermes sidecar-<app> allow|deny <code>` do the same
   without the question.
3. On the sidebar's next status poll an allowed request turns into a 256-bit grant, returned once. Hermes stores
   only sha256(grant), the origin and the time in `approvals.json` (0600). The sidebar keeps the grant in its
   origin's localStorage, so the approval survives Office and Hermes restarts.
4. `POST /<app>/session` requires a stored grant for the configured origin; every later pane request re-checks
   that the grant still exists (the store is re-read when the file changes), so `/sidecar-<app> revoke` or
   `hermes sidecar-<app> revoke` closes open sidebars at their next request.

Deny is one-shot; requests expire after 10 minutes; at most 3 wait per app; 20 connect requests a minute.
No Sidecar tool can approve, connect or revoke. An agent with shell access could run the CLI subject to Hermes command approval.
Allowing lets the whole configured origin use your Hermes agent and all its tools until revoke. Use a site you control or trust; open only documents you trust.
Planned hardening: TLS-only trust (`-p ssl`) and a server-auth EKU on the CA; the current command and CA shape are unchanged.

## Chat forwarding

Only `POST /v1/chat/completions`, only with a valid sidebar session. The bridge drops any Authorization header from
the page, sets `Authorization: Bearer <API_SERVER_KEY>` (read with `agent.secret_scope.get_secret` and kept in
memory) and forces `X-Hermes-Session-Id` to the sidebar's own session id, so tool calls from that turn reach that
sidebar and no other. The response (SSE) is streamed back unchanged; nothing is retried.

## Tool calls

`sidecar_<app>_*` handlers run in the same process as the bridge. They refuse a chat that is not a connected sidebar's
session (`ForeignSession`), then post a ticket over verified HTTPS localhost with the bridge's random per-process
plugin key (never on disk). The sidebar long-polls `/<app>/pending`, runs the Office.js operation and posts the
result. The handler receives trusted `session_id`, `task_id` and `user_task` (None on this path),
not `turn_id`. Approval plus a valid bound bridge session authorizes every known tool, not prompt
classification. Unknown tools, missing/invalid sessions and revoked grants fail closed. No comment
reservation is made without a trustworthy turn id. Stop fences accepted sockets, shuts them down and
closes them so old TLS keep-alive connections cannot keep serving a reloaded handler.

## What is never stored or logged

API_SERVER_KEY, grants, session keys, prompts, document text are never stored in plugin files or logged.
TLS private keys are the explicit on-disk exception, protected as described above and never logged.
Logs are route template, status and elapsed time; manual trust output contains only the public certificate path.
