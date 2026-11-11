# Native chat request lifetime regression

## Root cause

The composed path is pane `HermesAgent.oneModelCall` → TLS Desk handler →
`ChatProxy.forward` → Hermes API SSE response → pane SSE parser.

PowerPoint's `ResponseDeadlineMixin` starts a one-second absolute write deadline
at `send_response_only`. `ChatProxy.forward` then blocks in `upstream.read1`
between sending headers and writing SSE bytes. The first chunk arriving after
one second triggers `ResponseWriteAborted` before any body reaches the pane.
The proxy catches that `OSError` and closes its upstream connection. The base
handler's final flush can then log `REQUEST_PARSE error=ClientDisconnect`.
That log label does not establish that a malformed request or pane cancellation
caused the disconnect.

The pane's existing retry-loop guard already prevents native network retries,
but `oneModelCall` had a separate unconditional stream-to-JSON fallback. An empty
or absent stream body dispatched another chat POST with `stream: false`, replaying
a native turn that could already have edited the presentation.

The fix ends the header write budget after headers are flushed and ends each
chunk's budget after write+flush, **before waiting for more upstream data**.
Downstream writes still have the existing one-second bound; ordinary companion
responses keep their existing whole-response bound. The upstream timeout remains
600 seconds, and the native pane turn remains bounded at 300 seconds (with the
existing explicit checkpoint extension). Native empty responses now fail closed
with a deck-review warning, never invoking the fallback. Direct-provider fallback
is unchanged.

## Evidence and ruled-out hypotheses

- Read-only gateway log correlation for 2026-11-10 23:45:01, 23:46:35 and
  23:48:04 showed PowerPoint `ClientDisconnect` alongside 200/zero-byte upstream
  responses, followed by nonempty responses. No credentials or prompt content are
  needed to establish that ordering.
- Excel and Word use the equivalent proxy relay but do not wrap their handler
  with `ResponseDeadlineMixin`. Both have header deadlines. PowerPoint cancels
  its header timer after successful framing, before chat dispatch.
- The real composed TLS regression completes a pending/result tool roundtrip
  while the first chat request is open. Its fake upstream sends headers, waits
  three seconds, sends one SSE chunk, waits another three seconds, then completes.
  Before the fix, the client got an empty body (`None` in the test helper), not
  the expected full SSE payload. After the fix, the entire payload is delivered
  and the upstream records exactly one request. This distinguishes an elapsed
  response-write budget from tool-traffic-induced socket shutdown.
- The server accepted-socket closure runs on server stop, not when a tool ticket
  is claimed. Overload cleanup applies only to connections denied a handler slot;
  the composed regression is below capacity and exercises concurrent connections.
- Client fake-fetch coverage survives six seconds including simulated tool work,
  without abort or additional fetch. Existing native budget/checkpoint tests
  cover 25-second, 300-second and explicitly extended turns. The five-second
  progress callback and 30-second checkpoint are not the four-second trigger.
- The new empty/done-only/no-body client regressions failed before the fix by
  returning the second response, `Already done`, instead of rejecting (three
  failures; the delayed-fetch case already passed). The native network retry
  guard alone therefore did not prevent replay.

An older test explicitly required native stream-false fallback. It now asserts
one dispatch, protocol failure, unchanged history, no final answer and no local
execution instead. This replaces an unsafe expectation rather than dropping it.

## Scope

Source and automated tests only. No installed plugin, gateway, profile, keychain,
or running PowerPoint instance was modified in that fix pass. The owner subsequently rebuilt
and reloaded `3682e5f` and exercised the desktop Mac flow: one prompt, exactly one upstream
request open for about 24 seconds, successful edits/readback and the real sidebar answer,
without timeout or repeat. See [current live-test boundaries](../README.md#status).
