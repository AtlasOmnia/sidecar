<!-- SPDX-License-Identifier: Apache-2.0 -->
# Sidecar bridge: how the sidebar reaches Hermes

App-neutral bridge design for Sidecar for Word, Excel and PowerPoint (`desk_app.py` holds the app's name,
port range and default site). This Excel change introduces TLS; sibling repository synchronization is separate.

```
desktop Office sidebar (static page from addin_base, HTTPS)
   |  https://localhost:<app port>   exact Origin, Host localhost:<actual-port>
   v
desk bridge (thread inside the Hermes plugin, 127.0.0.1 only)
   |-- GET  /<app>/hello                 identify (service hermes-desk-<app>)
   |-- POST /<app>/connect               file a connect request -> code, requestId, pollSecret
   |-- POST /<app>/connect/status        pending | denied | expired | approved (+ grant, once)
   |-- POST /<app>/session               needs X-Hermes-Desk-Grant -> sessionId, sessionKey
   |-- POST /<app>/heartbeat, GET /<app>/pending, POST /<app>/result, DELETE /<app>/session
   |-- POST /v1/chat/completions         session auth -> Hermes API server (API_SERVER_KEY added here)
   `-- POST /<app>/invoke                sidecar_<app>_* tool calls of this process (in-memory key)
Hermes API server (127.0.0.1:8642) -> agent loop -> sidecar_<app>_* tools -> ticket -> sidebar -> Office.js
```

## Verified scope and operating limits

Latest live run: 2026-11-11 about 00:26 in desktop Excel for Mac, plugin and sidebar built from
`a7b3fd7`. Connect and Allow succeeded; a 3-row table with header in A1:B4 and `=SUM(B2:B4)` in B5
were read back from the workbook. Revoke dropped the sidebar to “needs your OK”. The bridge answered
200 to the configured origin and 403 to another origin. The sidebar used a private test site via `addin_base`;
the CA was trusted in the login keychain, and this Mac also has it in the System keychain.

Earlier the same night at `35ef09e` (behaviour unchanged since): reversal on request by clearing the range
(no real undo), plugin reload with the sidebar open (it reconnected by itself and edited again), approval
remembered after restarting Excel, and the next Connect after revoke asked for approval again.
Windows desktop Excel is untested. Also not tested: dark theme, keyboard-only use, the public Pages origin,
TLS-only trust, Excel for the web, or full tool/WorkbookPlan coverage. Reopen the sidebar if it does not reconnect.
Live-tested with a private sidebar test site; public-origin acceptance is pending publication.

Trust the plugin's local CA once, manually, using the login-keychain command printed by the bridge
(or explicitly request `/sidecar-excel trust`). The Excel test Mac also has it in the System keychain.
System-keychain trust was not required in our Word and PowerPoint runs (login-keychain trust only). Removal: Keychain Access > login > Certificates > “Sidecar local CA ...” > Delete;
remove any matching System entry if you installed one, then delete only the plugin's `tls` folder.
Uninstall/revoke do not remove OS trust. No tool should change keychains.

The default `addin_base` is `https://atlasomnia.github.io/sidecar-excel/` (public live test pending publication).
For another site set `plugins.entries.atlas-sidecar-excel.settings.addin_base` to that HTTPS base.
Match the built manifest's base URL to that setting. See README for the build/config commands.
The bridge/approval is local; Hermes may send workbook/chat data to the configured model provider.

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

On bridge start (or explicitly via `/sidecar-excel trust` / `hermes sidecar-excel trust`, not ordinary status), the plugin prints one manual command for the host platform. **It never runs the
command, calls an OS trust API, elevates privileges or installs system-wide trust.** Examples with placeholder paths:

* macOS login keychain: `security add-trusted-cert -r trustRoot -k "$HOME/Library/Keychains/login.keychain-db" '<plugin-data>/tls/ca.pem'`
* Windows PowerShell user root: `certutil -user -addstore Root '<plugin-data>\tls\ca.pem'`

The operator must review and explicitly trust the public CA, then reload Excel. For TLS it can only vouch for localhost. The trust command does not limit the certificate to TLS, so protect the key file and remove the trust when you stop using Sidecar. Planned hardening: TLS-only trust (`-p ssl`) and a server-auth EKU on the CA. TLS constraints:
critical name constraints permit DNS `localhost` and IP `127.0.0.1/32`, with no IPv6. DNS subtree semantics also
permit names below `localhost`; the generated non-CA server-auth leaf has only DNS `localhost` and lasts at most
397 days. Conforming verifiers reject unrelated DNS/IP certificates signed by this CA. Protect the keys.
Legacy unconstrained CAs and incorrect constraints/overlong CA validity refuse service without rotation, printing
one recovery instruction per runtime: remove old trust manually before deleting this plugin's TLS directory.
Only then restart and review the newly generated CA trust command.
Trust does not replace the approval protocol below. Uninstall/revoke do not remove keychain/root-store entries.
For expired CA or corrupt identity, stop this plugin, manually remove its matching CA from the login keychain or
Windows Current User root store, remove only this plugin's `tls` directory, restart and review/trust the replacement.
No production certificate verification bypass is supported. See [README](../README.md#local-tls-trust).

## Shutdown and CLI process ownership

`_BoundedThreadingHTTPServer.get_request()` registers accepted sockets (after optional TLS wrapping) in a
lock-protected set. `close_request()` removes normal, rejected and failed-handshake clients.
`close_clients()` sets a stopping fence under the same lock, snapshots/clears the set, then calls
`shutdown(socket.SHUT_RDWR)` and `close()` on each client. Shutdown interrupts handler reads even when a
handler's `makefile()` still retains the descriptor. Concurrent accepts see the fence and are closed rather
than dispatched. `StoppableHandler.parse_request()` also fences new requests on existing connections.
`ExcelCompanionServer.stop()` calls `begin_shutdown()`, cancels sessions/tickets, and allows already-admitted
requests up to one second (`drain_requests()`) to flush their cancellation/uncertainty receipts before closing
all clients, stopping the accept loop and closing the listening socket. Idle keep-alive reads do not extend
this grace period; stalled bodies/writes are forcibly closed at its end. `ThreadingHTTPServer.server_close()` alone is insufficient:
daemon keep-alive handlers otherwise remain attached to old code after plugin reload.

`_in_gateway()` checks the gateway marker but rejects Sidecar CLI argv even if `_HERMES_GATEWAY=1` was inherited.
Plugin registration invokes `maybe_autostart()` before `cli_run()`, so the guard must run at registration time.
CLI status/allow/deny/revoke only read/change the approval store, never start a listener. Ordinary status omits
TLS trust instructions; the explicit `trust` action displays the existing identity's command without starting
or creating a bridge. The shutdown and inherited-marker regression tests use disposable local listeners,
not an installed gateway, keychain or real Hermes CLI.

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
No tool can approve, connect or revoke.

## Chat forwarding

Only `POST /v1/chat/completions`, only with a valid sidebar session. The bridge drops any Authorization header from
the page, sets `Authorization: Bearer <API_SERVER_KEY>` (read with `agent.secret_scope.get_secret` and kept in
memory) and forces `X-Hermes-Session-Id` to the sidebar's own session id, so tool calls from that turn reach that
sidebar and no other. The response (SSE) is streamed back unchanged; nothing is retried.

## Tool calls

`sidecar_<app>_*` handlers run in the same process as the bridge. They refuse a chat that is not a connected sidebar's
session (`ForeignSession`), then post a ticket over verified HTTPS localhost with the bridge's random per-process
plugin key (never on disk). The sidebar long-polls `/<app>/pending`, runs the Office.js operation and posts the
result.

## What is never stored or logged

API_SERVER_KEY, grants, session keys, prompts, document text are never stored in plugin files or logged.
TLS private keys are the explicit on-disk exception, protected as described above and never logged.
Logs are route template, status and elapsed time; manual trust output contains only the public certificate path.
