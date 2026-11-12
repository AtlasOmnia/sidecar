<!-- SPDX-License-Identifier: Apache-2.0 -->
# Sidecar bridge: how the sidebar reaches Hermes

App-neutral bridge design for Sidecar for Word, Excel and PowerPoint (`desk_app.py` holds the app's name,
port range and default site). This PowerPoint bridge uses the verified-localhost-TLS implementation ported from Excel.

```
desktop Office sidebar (bundled at /sidecar-powerpoint/, same-origin HTTPS from the plugin bridge by default)
   |  https://localhost:<actual app port>   exact Origin when present, Host localhost:<actual-port>
   v
desk bridge (thread inside the Hermes plugin, 127.0.0.1 only)
   |-- GET/HEAD /sidecar-powerpoint/<file>  checksummed bundled sidebar assets
   |-- POST/PUT/DELETE /sidecar-powerpoint/* -> 405
   |-- GET  /<app>/hello                 identify (service hermes-desk-<app>)
   |-- POST /<app>/connect               file a connect request -> code, requestId, pollSecret
   |-- POST /<app>/connect/status        pending | denied | expired | approved (+ grant, once)
   |-- POST /<app>/session               needs X-Hermes-Desk-Grant -> sessionId, sessionKey
   |-- POST /<app>/heartbeat, GET /<app>/pending, POST /<app>/result, DELETE /<app>/session
   |-- POST /v1/chat/completions         session auth -> Hermes API server (API_SERVER_KEY added here)
   `-- POST /<app>/invoke                sidecar_<app>_* tool calls of this process (in-memory key)
Hermes API server (127.0.0.1:8642) -> agent loop -> sidecar_<app>_* tools -> ticket -> sidebar -> Office.js
```

## Operational boundaries

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
and TLS-only trust were not tested. The plugin-served same-origin default in 1.2.0 has not yet been live-tested in Office.
After a plugin reload, use Try again or reopen the sidebar to reconnect. After an update, close and reopen the sidebar to load the current pane. Stopping the bridge shuts down and closes
accepted sockets, including TLS keep-alive connections; the existing overload cleanup and request deadlines remain.
There is no real undo: Hermes reverses its own edit by editing again, without an Office rollback transaction.
Default `addin_base`: `https://localhost:47650/sidecar-powerpoint/`; the bridge resolves the first free port in 47650-47654 and serves its bundled sidebar at `https://localhost:<actual-port>/sidecar-powerpoint/`. Run `/sidecar-powerpoint manifest` to write a manifest for that actual port. External HTTPS hosting remains optional through `addin_base` and requires an exactly matching build.

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

On bridge start (also with `/sidecar-powerpoint trust` or `hermes sidecar-powerpoint trust`, not plain status), the plugin prints one manual command for the host platform. **It never runs the
command, calls an OS trust API, elevates privileges or installs system-wide trust.** Examples with placeholder paths:

* macOS login keychain: `security add-trusted-cert -r trustRoot -k "$HOME/Library/Keychains/login.keychain-db" '<plugin-data>/tls/ca.pem'`
* Windows PowerShell user root: `certutil -user -addstore Root '<plugin-data>\tls\ca.pem'`

The operator must review and explicitly trust the public CA, then reload PowerPoint. For TLS it can only vouch for localhost. The trust command does not limit the certificate to TLS, so protect the key file and remove the trust when you stop using Sidecar.
For TLS,
critical name constraints permit DNS `localhost` and IP `127.0.0.1/32`, with no IPv6. DNS subtree semantics also
permit names below `localhost`; the generated non-CA server-auth leaf has only DNS `localhost` and lasts at most
397 days. Conforming verifiers reject unrelated DNS/IP certificates signed by this CA. Protect the keys.
Legacy unconstrained CAs and incorrect constraints/overlong CA validity refuse service without rotation, printing
one recovery instruction per runtime: remove old trust manually before deleting this plugin's TLS directory.
Only then restart and review the newly generated CA trust command.
Trust does not replace the approval protocol below. Uninstall/revoke do not remove keychain/root-store entries.
For expired CA or corrupt identity, stop this plugin, manually remove its matching CA from the login keychain or
Windows Current User root store, remove only this plugin's `tls` directory, restart and review/trust the replacement.
On Mac remove trust in Keychain Access > login > Certificates > “Sidecar for PowerPoint local CA ...” (older installs: “Sidecar local CA ...”) > Delete,
then delete this plugin's `tls` folder while the bridge is stopped.
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
No Sidecar tool can approve, connect or revoke. An agent with shell access could run the CLI approval command, subject to Hermes command approval.
Allowing lets the site at the configured origin use your Hermes agent and all its tools until you revoke.
Use a site you control or trust; open only documents you trust.
TLS-only trust (`-p ssl`) and a server-auth EKU on the CA are planned hardening; neither is implemented yet.

## Chat forwarding

Only `POST /v1/chat/completions`, only with a valid sidebar session. The bridge drops any Authorization header from
the page, sets `Authorization: Bearer <API_SERVER_KEY>` (read with `agent.secret_scope.get_secret` and kept in
memory) and forces `X-Hermes-Session-Id` to the sidebar's own session id, so tool calls from that turn reach that
sidebar and no other. The response (SSE) is streamed back unchanged; nothing is retried.

The native pane request is one whole Hermes agent turn, not one model completion: its timeout is 300 seconds,
with progress status every 5 seconds. The 30-second checkpoint runs while the call is in flight; Continue resets
the turn deadline to 300 seconds from the decision, Stop aborts it, and no decision remains deadline-bounded.
Native failures, including timeout, are not retried automatically. Cancellation does not undo edits already performed.

## Tool calls

`sidecar_<app>_*` handlers run in the same process as the bridge. They refuse a chat that is not a connected sidebar's
session (`ForeignSession`), then post a ticket over verified HTTPS localhost with the bridge's random per-process
plugin key (never on disk). The sidebar long-polls `/<app>/pending`, runs the Office.js operation and posts the
result.

## What is never stored or logged

The bridge does not persist or log API_SERVER_KEY, plaintext grants, session keys, prompts or document text.
The Hermes agent/server has its own conversation retention policy; this statement does not cover that history.
TLS private keys are the explicit on-disk exception, protected as described above and never logged.
Logs are route template, status and elapsed time; manual trust output contains only the public certificate path.
