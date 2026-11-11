# PowerPoint localhost TLS port — verification

Source: Excel commits `51bff9b`, `f141b4b`, `62826e4`, `6c4ef8e`, `15ee898`.
Historical receipt for the earlier TLS port only (not the current live-test status).
Subsequently, title/subtitle editing was tested live in desktop PowerPoint for Mac; Windows remains untested.

Target: PowerPoint, branch `originals-compliance`. Local source/build verification only; no installation, deployment, OS trust changes or live Office acceptance.

## Test-first evidence

Ported `test_desk_tls.py`, security transport tests and connector tests before implementation.

- Initial collection exposed the absent implementation: `ModuleNotFoundError: No module named 'desk_plugin.desk_tls'`.
- Temporarily deferred that import to allow the baseline behavior tests to execute, then restored the original import once the module existed. Baseline run: **7 failed, 19 deselected**.
  - `test_d4_tls_first_run_and_persistent_identity`: `AssertionError: assert 'http://127.0.0.1:63878' == 'https://localhost:63878'`.
  - `test_s5_bridge_serves_with_truststore_injected`: `FileNotFoundError` loading the absent `tls/ca.pem`; truststore fixture ran, not skipped.
  - `test_s5_failed_handshake_is_logged_not_silent`: `ssl.SSLError: [SSL: WRONG_VERSION_NUMBER] wrong version number` against the plaintext bridge.
  - All four configured-origin preflight cases: `AttributeError: 'DeskBridge' object has no attribute 'tls'`.
- After initial TLS plumbing, before the truststore/CORS fixes: **4 failed, 2 passed, 20 deselected**.
  - Truststore test: `ConnectionResetError: [Errno 54] Connection reset by peer` after the handshake.
  - Heartbeat, pending and result preflight tests: `AssertionError: assert None == 'https://sidebar.example.test'` at `headers.get("access-control-allow-origin") == ORIGIN`; response was 204 without CORS. Session was already handled by the desk bridge.
- With TLS but without handshake logging: **1 failed, 25 deselected**. Exact assertion: `AssertionError: assert 'route=TLS_HANDSHAKE error=' in ''`.
- Full-suite integration exposed a PowerPoint-specific HTTP-only native-mode predicate: **4 failed, 340 passed** in Vitest. The HTTPS fixture lost `X-Hermes-Session-Id`, advertised local tools, and no longer failed closed on missing session identity. Added HTTPS localhost recognition without changing the direct-model path or adding a discovery fallback.

## Final observed commands

Environment:

```sh
export TMPDIR=<private-scratch-directory>
export PYTHON=<test-venv>/bin/python
export PYTHONPATH=<hermes-source>:<hermes-site-packages>
export PYTHONDONTWRITEBYTECODE=1
```

- Focused: `(cd scripts/tests && $PYTHON -m pytest -q -rs -p no:cacheprovider test_desk_tls.py test_desk_bridge_security.py)` — exit 0, **68 passed, 3 warnings in 33.18s**, no skips; truststore test executed.
- `npm test` — exit 0:

```text
 Test Files  32 passed (32)
      Tests  344 passed (344)
194 passed, 3 warnings in 69.76s (0:01:09)
```

  Pytest: zero failed, zero skipped. The launcher explicitly honors `$PYTHON` and disables the cache provider.
- `npm run build:static -- --base https://sidebar.example.test/powerpoint/` — exit 0; `webpack 5.108.4 compiled with 2 warnings`; `Built dist for https://sidebar.example.test/powerpoint/`. Warnings concern the 351 KiB `.icns` asset and webpack performance recommendations. Build output stayed in the target repo.
- `git diff --check` — exit 0 after whitespace cleanup.

## Exact pane-preflight coverage

Each receives OPTIONS with session headers from the configured add-in origin, asserts 204 plus exact Allow-Origin and all four requested header names, and checks that an unrelated origin receives no Allow-Origin:

- `POST /powerpoint/heartbeat`
- `GET /powerpoint/pending`
- `POST /powerpoint/result`
- `POST /powerpoint/session`

The test asserts that this route set equals PowerPoint's complete `OPTIONS_ROUTES`. PowerPoint has no `/execute-start` route. Existing security tests continue covering the desk connect preflight and origin rejection.

## Adaptation and preserved behavior

- Copied app-neutral `desk_tls.py` exactly from the final Excel reference: constrained localhost CA, safe legacy-CA refusal, renewal, owner-only identity files, stdlib server SSLContext despite truststore injection, and C-level TLS minimum-version setter.
- Retained PowerPoint's `OverloadCleanupMixin`, concurrency validation, bounded response-write implementation and request deadlines. Added TLS handshakes in bounded request threads, not the accept loop; logger remains `powerpoint_companion` with exception-class-only handshake metadata.
- PowerPoint's `_handle_options` uses `_emit_response`, unlike Excel's direct response-writing path. Retained that writer and replaced only its hard-coded origin decision with `self._check_origin()`. Other direct constant comparisons are in `_check_origin` itself and the two CORS helpers; the desk bridge overrides all three, so no other production hard-coded origin gate remains.
- PowerPoint uses `http.client`, not Excel's urllib/plan client. Added HTTPSConnection with the runtime-supplied verified CA context while preserving its bounded same-request-ID retry and response validation. Standalone companion HTTP fixtures remain supported; production DeskRuntime and sidebar discovery always use HTTPS.
- Kept the 15 PowerPoint tools, ports 47650–47654, configured add-in base, profile-aware upstream API routing and PowerPoint protocol intact. No Excel plan fixtures or plan routes were copied. Equivalent frontend wiring is covered by connector, manifest and native-agent fixtures.

## Reuse lesson and limits

When porting a bridge URL scheme, audit every URL-sensitive consumer, not only discovery: an HTTP-only native-mode predicate can silently turn a server-owned agent turn into a direct-model/local-tool turn. Keep an HTTPS native-mode fixture that verifies session headers and the absence of client-owned tool execution.

No live desktop PowerPoint test or OS-keychain trust test was performed. No Hermes CLI was started, no live plugin/profile/gateway/launchd/release files were changed, no trust command was executed, and nothing was pushed. Test CA keys stayed in disposable scratch directories and were never printed or committed. Skills/profile memory were not edited because this task permits writes only in the target repository.
