"""R2d-3: companion and plugin legs of the shared plan wire contract.

The fixture ``src/__tests__/fixtures/companion-wire/plan-wire-contract.v1.json``
is the single recorded wire contract. This leg replays its pane, plugin and
cancel requests against the REAL ``ExcelCompanionServer`` (in process,
ephemeral loopback port, injected clock) and asserts every companion response
exactly (ordered keys, no subset matching). The plugin leg sends the recorded
invoke through the REAL ``bridge_client.invoke_bridge`` and asserts the
client-side outcome plus the model-visible result of the real plan handler.
``src/__tests__/companion-wire-contract.test.ts`` replays the same fixture
against the real TypeScript bridge.
"""
from __future__ import annotations

import hashlib
import http.client
import json
import re
import socket
import threading
import time
from pathlib import Path
from typing import Any
from unittest.mock import patch

import pytest

from desk_plugin.companion import protocol
from desk_plugin.companion.server import ExcelCompanionServer
from scripts.tests.test_companion_plan_execution import PLUGIN_KEY
from scripts.tests.test_office_plugin_plan_contract import (
    _R2D2G2Server,
    _r2d2_plan_handler_result,
    load_plugin,
    valid_execute_plan,
    valid_plan,
)

REPO = Path(__file__).resolve().parents[2]
FIXTURE_PATH = REPO / "src" / "__tests__" / "fixtures" / "companion-wire" / "plan-wire-contract.v2.json"
FIXTURE_BYTES = FIXTURE_PATH.read_bytes()
FIXTURE: dict[str, Any] = json.loads(FIXTURE_BYTES)
FIXTURE_SHA256 = hashlib.sha256(FIXTURE_BYTES).hexdigest()
CASES: list[dict[str, Any]] = FIXTURE["cases"]
COMPANION_CASES = [case for case in CASES if "companion" in case["sides"]]
PLUGIN_CASES = [case for case in CASES if "plugin" in case["sides"]]
# Pinned independently of the fixture (R2d-3 G2 / A1): the fixture's own
# ``requiredTags`` and per-case ``sides`` are metadata, so a fixture-only edit
# could otherwise drop a tag or a leg without failing anything. Rule: a new or
# changed case (e.g. R3a's C23) must be added deliberately to BOTH this table
# and ``EXPECTED_CASE_SIDES`` in src/__tests__/companion-wire-contract.test.ts,
# and the pinned per-leg counts updated in both files.
REQUIRED_TAGS: tuple[str, ...] = (
    "session", "pending", "execute_start", "completed", "not_executed", "duplicate", "cancelled",
    "executing_uncertain", "expired_claimed", "expired_unclaimed", "4xx", "5xx", "truncation",
    "mixed_version", "primitive", "invalid_ticket", "ack_timeout", "bridge_throw", "route_payload",
    "phase_inspect", "phase_fence",
)
EXPECTED_CASE_SIDES: dict[str, tuple[str, ...]] = {
    "C01-session":                                         ("companion", "bridge"),
    "C02-completed":                                       ("companion", "plugin", "bridge"),
    "C03-not-executed":                                    ("companion", "plugin", "bridge"),
    "C04-duplicate":                                       ("companion",),
    "C05-claimed-executed-true":                           ("companion",),
    "C06-lost-ack-response":                               ("companion", "plugin", "bridge"),
    "C07-double-ack":                                      ("companion", "bridge"),
    "C07-double-ack/late_ack":                             ("bridge",),
    "C08-cancel-pending":                                  ("companion", "plugin", "bridge"),
    "C09-cancel-claimed":                                  ("companion", "plugin", "bridge"),
    "C10-cancel-executing":                                ("companion", "plugin", "bridge"),
    "C11-claimed-lease-expiry":                            ("companion", "plugin", "bridge"),
    "C12-executing-window":                                ("companion",),
    "C13-unclaimed-expiry":                                ("companion", "plugin"),
    "C14-4xx-set/result-without-executed":                 ("companion",),
    "C14-4xx-set/execute-start-ticket":                    ("companion",),
    "C14-4xx-set/pane-auth":                               ("companion",),
    "C14-4xx-set/execute-start-401":                       ("companion", "bridge"),
    "C14-4xx-set/second-invoke-same-turn":                 ("companion", "plugin"),
    "C14-4xx-set/invoke-unknown-session":                  ("companion", "plugin"),
    "C15-5xx-transport/504-no-category":                   ("plugin",),
    "C15-5xx-transport/503-html":                          ("plugin",),
    "C15-5xx-transport/500-internal-json":                 ("plugin",),
    "C16-truncation/L":                                    ("plugin",),
    "C16-truncation/L-short-of-content-length":            ("plugin",),
    "C16-truncation/B1-pending":                           ("bridge",),
    "C16-truncation/B2-execute-start-truncated":           ("bridge",),
    "C16-truncation/B2b-execute-start-text-rejects":       ("bridge",),
    "C17-primitive":                                       ("companion", "bridge"),
    "C17-primitive/B-trusted":                             ("bridge",),
    "C18-invalid-plan-ticket/version-2":                   ("bridge",),
    "C18-invalid-plan-ticket/arguments-array":             ("bridge",),
    "C18-invalid-plan-ticket/trusted-missing":             ("bridge",),
    "C18-invalid-plan-ticket/trusted-extra-key":           ("bridge",),
    "C18-invalid-plan-ticket/trusted-foreign-fingerprint": ("bridge",),
    "C18-invalid-plan-ticket/trusted-turn-not-uuid":       ("bridge",),
    "C18-invalid-plan-ticket/trusted-v1-two-keys":         ("bridge",),
    "C18-invalid-plan-ticket/trusted-phase-unknown":       ("bridge",),
    "C18-invalid-plan-ticket/trusted-turnkey-equals-op":   ("bridge",),
    "C19a-throw-before-ack":                               ("companion", "bridge"),
    "C19b-throw-after-ack":                                ("companion", "bridge"),
    "C20-ack-timeout":                                     ("bridge",),
    "C20-ack-timeout/aborted-signal":                      ("bridge",),
    "MV1-old-pane-no-executed":                            ("companion", "plugin"),
    "MV2-old-pane-no-fingerprint":                         ("companion", "plugin"),
    "MV2-B-non-negotiated-plan-ticket":                    ("bridge",),
    "MV3-new-pane-r2c-companion":                          ("bridge",),
    "C21a-rejected":                                       ("companion", "plugin", "bridge"),
    "C21b-refuse":                                         ("companion", "plugin", "bridge"),
    "C21c-cancelled":                                      ("companion", "plugin", "bridge"),
    "C21d-route-uncertain":                                ("companion", "plugin", "bridge"),
    "C21e-verification-failed":                            ("companion", "plugin", "bridge"),
    "C21-R/R1-rollout-disabled":                           ("bridge",),
    "C21-R/R2-no-bridge-session":                          ("bridge",),
    "C22-not-started":                                     ("companion", "plugin"),
    "C23-plan-fence-capacity":                             ("companion", "plugin"),
    "C24-inspect-then-execute":                            ("companion", "plugin"),
    "C25-inspect-twice":                                   ("companion", "plugin"),
    "C26-inspect-after-execute":                           ("companion", "plugin"),
    "C27-inspect-ticket-cannot-execute":                   ("companion",),
    "C28-phase-outcome-mismatch":                          ("companion",),
    "MV4-v1-plugin-envelope":                              ("companion",),
    "MV5-v2-plugin-to-v1-companion":                       ("plugin",),
    "C29-phase-outcome-mismatch-ticket":                   ("bridge",),
    "C30-inspect-ticket-gate-on":                          ("bridge",),
}
EXPECTED_COMPANION_CASE_COUNT = 37
EXPECTED_PLUGIN_CASE_COUNT = 28
STEP_KINDS = {
    "pane", "pane_transport", "pane_hang", "plugin_invoke_begin", "plugin_invoke_end",
    "plugin_cancel", "clock_advance", "invoke_transport",
}
SIDES = {"companion", "plugin", "bridge"}
UUID_RE = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$", re.I)
HEX64_RE = re.compile(r"^[0-9a-f]{64}$")
JOIN_SECONDS = 5.0


# --------------------------------------------------------------------- tokens
def _initial_bindings() -> dict[str, Any]:
    bindings = {f"${name}": value for name, value in FIXTURE["ids"].items()}
    bindings["$fingerprint"] = FIXTURE["workbookPlanFingerprint"]
    return bindings


def _is_token(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("$")


def resolve(value: Any, bindings: dict[str, Any]) -> Any:
    """Replace every token; an unbound token is a fixture error."""
    if _is_token(value):
        assert value in bindings, f"unbound token {value}"
        return bindings[value]
    if isinstance(value, dict):
        return {key: resolve(item, bindings) for key, item in value.items()}
    if isinstance(value, list):
        return [resolve(item, bindings) for item in value]
    return value


def match(expected: Any, actual: Any, bindings: dict[str, Any], where: str) -> None:
    """Exact match with ordered keys; ``$name`` binds on first sight."""
    if expected == "$any:uuid":
        assert isinstance(actual, str) and UUID_RE.match(actual), (where, actual)
        return
    if expected == "$any:hex64":
        assert isinstance(actual, str) and HEX64_RE.match(actual), (where, actual)
        return
    if _is_token(expected):
        if expected in bindings:
            assert actual == bindings[expected] and type(actual) is type(bindings[expected]), (
                where, expected, bindings[expected], actual)
        else:
            assert isinstance(actual, str) and actual, (where, expected, actual)
            bindings[expected] = actual
        return
    if isinstance(expected, dict):
        assert isinstance(actual, dict), (where, expected, actual)
        assert list(actual) == list(expected), (where, "keys", list(expected), list(actual))
        for key, item in expected.items():
            match(item, actual[key], bindings, f"{where}.{key}")
        return
    if isinstance(expected, list):
        assert isinstance(actual, list) and len(actual) == len(expected), (where, expected, actual)
        for index, item in enumerate(expected):
            match(item, actual[index], bindings, f"{where}[{index}]")
        return
    assert type(actual) is type(expected) and actual == expected, (where, expected, actual)


def _loads_ordered(raw: bytes) -> Any:
    return json.loads(raw.decode("utf-8")) if raw else None


# ------------------------------------------------------------------- transport
def _send(port: int, method: str, path: str, body: Any, headers: dict[str, str]) -> tuple[int, bytes]:
    conn = http.client.HTTPConnection("127.0.0.1", port, timeout=JOIN_SECONDS)
    try:
        raw = None if body is None else json.dumps(body, separators=(",", ":")).encode("utf-8")
        conn.request(method, path, body=raw, headers=headers)
        response = conn.getresponse()
        return response.status, response.read()
    finally:
        conn.close()


def _assert_response(expected: dict[str, Any], status: int, raw: bytes, bindings: dict[str, Any],
                     where: str) -> None:
    assert "responseText" not in expected, f"{where}: responseText is replayed by the bridge leg only"
    assert status == expected["status"], (where, expected["status"], status, raw[:300])
    if expected.get("empty") is True:
        assert raw == b"", (where, raw)
        return
    match(expected["body"], _loads_ordered(raw), bindings, f"{where}.body")


class _ForwardingProxy:
    """Loopback proxy that forwards one invoke to the real companion, then
    returns the real status line and headers with the body cut to N bytes
    (``None``: all of it) and Content-Length raised by ``content_length_delta``."""

    def __init__(self, upstream_port: int, body_bytes: int | None, content_length_delta: int) -> None:
        self.upstream_port = upstream_port
        self.body_bytes = body_bytes
        self.content_length_delta = content_length_delta
        self.connections = 0
        self.upstream_response: bytes = b""
        self._sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self._sock.bind(("127.0.0.1", 0))
        self._sock.listen(4)
        self._sock.settimeout(0.2)
        self.port = self._sock.getsockname()[1]
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._serve, daemon=True)

    def __enter__(self) -> "_ForwardingProxy":
        self._thread.start()
        return self

    def __exit__(self, *_args: Any) -> None:
        self._stop.set()
        self._thread.join(timeout=JOIN_SECONDS)
        self._sock.close()

    def _serve(self) -> None:
        while not self._stop.is_set():
            try:
                conn, _addr = self._sock.accept()
            except (TimeoutError, socket.timeout):
                continue
            except OSError:
                return
            self.connections += 1
            with conn:
                conn.settimeout(JOIN_SECONDS * 2)
                request = _R2D2G2Server._read_request(conn)
                with socket.create_connection(("127.0.0.1", self.upstream_port), timeout=JOIN_SECONDS * 2) as up:
                    up.sendall(request)
                    chunks = []
                    while True:
                        chunk = up.recv(65536)
                        if not chunk:
                            break
                        chunks.append(chunk)
                self.upstream_response = b"".join(chunks)
                head, _sep, body = self.upstream_response.partition(b"\r\n\r\n")
                lines = head.split(b"\r\n")
                for index, line in enumerate(lines):
                    name, _colon, value = line.partition(b":")
                    if name.strip().lower() == b"content-length":
                        declared = int(value.strip()) + self.content_length_delta
                        lines[index] = b"Content-Length: " + str(declared).encode("ascii")
                sent = body if self.body_bytes is None else body[: self.body_bytes]
                try:
                    conn.sendall(b"\r\n".join(lines) + b"\r\n\r\n" + sent)
                    conn.shutdown(socket.SHUT_WR)
                except OSError:
                    pass


def _raw_http_response(step: dict[str, Any]) -> bytes:
    body = (json.dumps(step["body"]) if "body" in step else step["bodyText"]).encode("utf-8")
    head = (f"HTTP/1.1 {step['status']}\r\nContent-Type: application/json\r\nConnection: close\r\n"
            f"Content-Length: {len(body)}\r\n\r\n")
    return head.encode("ascii") + body


# ---------------------------------------------------------------------- runner
class _Run:
    """Replays one case. ``mode`` is ``companion`` (raw invoke) or ``plugin``."""

    def __init__(self, case: dict[str, Any], mode: str) -> None:
        self.case = case
        self.mode = mode
        self.bindings = _initial_bindings()
        self.clock = [5000.0]
        self.server = ExcelCompanionServer(
            host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
            clock=lambda: self.clock[0],
            # The real wait never ends a case: only the injected clock (via
            # sweep) decides expiry; every invoke thread is joined in <= 5 s.
            invoke_timeout_seconds=FIXTURE["constants"]["invokeTimeoutSeconds"],
        )
        self.port = 0
        self.invoke_thread: threading.Thread | None = None
        self.invoke_outcome: dict[str, Any] = {}
        self.transport: dict[str, Any] | None = None
        self.transport_server: Any = None
        self.plugin = load_plugin() if mode == "plugin" else None

    # -- helpers
    def _session(self):
        session_id = self.bindings.get("$sessionId")
        return None if session_id is None else self.server.sessions.get(session_id)

    def _pane_headers(self, step: dict[str, Any]) -> dict[str, str]:
        headers = dict(resolve(step["headers"], self.bindings))
        override = step.get("companionRequest", {})
        headers.update(resolve(override.get("headers", {}), self.bindings))
        if not override.get("omitOrigin"):
            headers["Origin"] = FIXTURE["origin"]
        return headers

    def run(self) -> None:
        self.server.start()
        self.port = self.server.server_address()[1]
        try:
            for index, step in enumerate(self.case["steps"]):
                kind = step["kind"]
                assert kind in STEP_KINDS, f"unknown step kind {kind}"
                if step.get("only") == "bridge" or kind == "pane_hang":
                    continue
                getattr(self, f"_step_{kind}")(step, f"{self.case['id']}[{index}]")
            assert self.invoke_thread is None, "an invoke was begun but never ended"
        finally:
            if self.transport_server is not None:
                self.transport_server.__exit__(None, None, None)
            self.server.stop()

    # -- step kinds
    def _step_pane(self, step: dict[str, Any], where: str) -> None:
        body = resolve(step["body"], self.bindings) if "body" in step else None
        status, raw = _send(self.port, step["method"], step["path"], body, self._pane_headers(step))
        _assert_response(step["response"], status, raw, self.bindings, where)

    def _step_pane_transport(self, step: dict[str, Any], where: str) -> None:
        # The bridge leg sees a transport fault; the companion processed it.
        self._step_pane(step, where)

    def _step_clock_advance(self, step: dict[str, Any], _where: str) -> None:
        self.clock[0] += step["seconds"]
        self.server.sessions.sweep()

    def _step_plugin_cancel(self, step: dict[str, Any], where: str) -> None:
        status, raw = _send(self.port, "POST", "/excel/cancel", resolve(step["body"], self.bindings),
                            {"Content-Type": "application/json", "X-Hermes-Plugin-Key": PLUGIN_KEY})
        _assert_response(step["response"], status, raw, self.bindings, where)

    def _step_invoke_transport(self, step: dict[str, Any], _where: str) -> None:
        assert self.mode == "plugin", "invoke_transport is a plugin-leg mutation"
        self.transport = step
        if step["mode"] == "respond":
            self.transport_server = _R2D2G2Server(_raw_http_response(step)).__enter__()
        else:
            assert step["mode"] == "truncate"
            self.transport_server = _ForwardingProxy(
                self.port, step["bodyBytes"], step.get("contentLengthDelta", 0)).__enter__()

    def _step_plugin_invoke_begin(self, step: dict[str, Any], _where: str) -> None:
        assert self.invoke_thread is None, "overlapping invokes"
        request = resolve(step["request"], self.bindings)
        session = self._session()
        setup = self.case.get("companionSetup")
        if setup is not None:
            assert session is not None and set(setup) == {"planFenceLimit"}, (self.case["id"], setup)
            session._plan_turn_history_limit = setup["planFenceLimit"]
        before = session.get_ticket() if session is not None else None
        self.invoke_outcome = {"request": request}
        target = self._raw_invoke if self.mode == "companion" else self._plugin_invoke
        self.invoke_thread = threading.Thread(target=target, args=(request,), daemon=True)
        self.invoke_thread.start()
        # Wait until the companion minted a new pending ticket, or the invoke
        # already answered (4xx before a ticket, or an exact reattach).
        deadline = time.monotonic() + JOIN_SECONDS
        while time.monotonic() < deadline and self.invoke_thread.is_alive():
            current = session.get_ticket() if session is not None else None
            if current is not None and current is not before and current.state == "pending":
                break
            time.sleep(0.002)

    def _raw_invoke(self, request: dict[str, Any]) -> None:
        port = self.transport_server.port if self.transport_server is not None else self.port
        try:
            self.invoke_outcome["http"] = _send(port, "POST", "/excel/invoke", request, {
                "Content-Type": "application/json", "X-Hermes-Plugin-Key": PLUGIN_KEY})
        except Exception as exc:  # pragma: no cover - surfaced by the end step
            self.invoke_outcome["error"] = exc

    def _plugin_invoke(self, request: dict[str, Any]) -> None:
        bridge_client = self.plugin.bridge_client
        port = self.transport_server.port if self.transport_server is not None else self.port
        sent: list[bytes] = []
        real_urlopen = bridge_client.urllib.request.urlopen

        def recording_urlopen(req: Any, timeout: Any = None):
            sent.append(req.data)
            return real_urlopen(req, timeout=timeout)

        trusted = None
        if "trusted" in request:
            trusted = {
                "turn_operation_id": request["trusted"]["turnOperationId"],
                "request_id": request["requestId"],
                "workbook_plan_fingerprint": request["trusted"]["workbookPlanFingerprint"],
                "phase": request["trusted"]["phase"],
                "turn_key": request["trusted"]["turnKey"],
            }
        self.invoke_outcome["sent"] = sent
        with (
            patch.object(bridge_client, "_connection",
                         return_value=(f"http://127.0.0.1:{port}", Path("/never-read"))),
            patch.object(bridge_client, "_validate_key", return_value=PLUGIN_KEY),
            patch.object(bridge_client.urllib.request, "urlopen", recording_urlopen),
        ):
            try:
                self.invoke_outcome["returns"] = bridge_client.invoke_bridge(
                    request["sessionId"], request["tool"], request["arguments"], trusted=trusted)
            except bridge_client.BridgeError as exc:
                self.invoke_outcome["raises"] = exc
            except Exception as exc:  # pragma: no cover - surfaced by the end step
                self.invoke_outcome["error"] = exc

    def _step_plugin_invoke_end(self, step: dict[str, Any], where: str) -> None:
        thread = self.invoke_thread
        assert thread is not None, f"{where}: no invoke in flight"
        thread.join(JOIN_SECONDS)
        assert not thread.is_alive(), f"{where}: invoke did not finish"
        self.invoke_thread = None
        outcome = self.invoke_outcome
        assert "error" not in outcome, (where, outcome.get("error"))
        if self.mode == "companion":
            status, raw = outcome["http"]
            _assert_response(step["response"], status, raw, self.bindings, where)
            return
        self._assert_plugin_outcome(step, where)

    def _assert_plugin_outcome(self, step: dict[str, Any], where: str) -> None:
        outcome = self.invoke_outcome
        expected = step["pluginOutcome"]
        bridge_client = self.plugin.bridge_client
        # The plugin sent exactly the recorded invoke body, once.
        sent = outcome["sent"]
        assert len(sent) == 1, (where, "urlopen calls", len(sent))
        match(outcome["request"], json.loads(sent[0]), self.bindings, f"{where}.sentInvoke")
        if self.transport is not None:
            assert self.transport_server.connections == expected.get("connections", 1), where
            if self.transport["mode"] == "truncate":
                # The proxy forwarded the real companion's complete answer.
                head, _sep, body = self.transport_server.upstream_response.partition(b"\r\n\r\n")
                assert head.split(b"\r\n", 1)[0] == f"HTTP/1.1 {step['response']['status']} OK".encode(), head
                match(step["response"]["body"], _loads_ordered(body), self.bindings, f"{where}.upstream")
        if "returns" in expected:
            assert "returns" in outcome, (where, "raised", outcome.get("raises"))
            assert expected["returns"] == "$response.body"
            match(step["response"]["body"], outcome["returns"], self.bindings, f"{where}.returns")
            view_module = load_plugin()
            fake = lambda *_a, **_kw: outcome["returns"]  # noqa: E731
        else:
            raised = outcome.get("raises")
            assert raised is not None, (where, "returned", outcome.get("returns"))
            assert type(raised) is bridge_client.BridgeError
            assert raised.category == expected["raises"]["category"], (where, raised.category, str(raised))
            assert str(raised) == expected["raises"]["message"], (where, str(raised))
            view_module = load_plugin()

            def fake(*_a: Any, **_kw: Any) -> Any:
                raise view_module.bridge_client.BridgeError(raised.category, str(raised))
        if "modelView" in expected:
            # A fresh plugin module per view: its one-shot turn registry must
            # not see the previous invoke of a multi-invoke case.
            # R4g-2c-r1 rule W1: the model view runs under the plan this invoke actually sent (the plugin's
            # executed:false filter is phase-aware); before r1 every view used the inspect plan valid_plan().
            view = _model_view(view_module, fake, outcome["request"]["arguments"])
            match(resolve(expected["modelView"], self.bindings), view, self.bindings, f"{where}.modelView")


class _Context:
    def __init__(self) -> None:
        self.tools: list[dict[str, Any]] = []
        self.hooks: dict[str, Any] = {}

    def register_tool(self, **kwargs: Any) -> None:
        self.tools.append(kwargs)

    def register_hook(self, name: str, callback: Any) -> None:
        self.hooks[name] = callback


def _model_view(module: Any, bridge: Any, plan: dict[str, Any] | None = None) -> Any:
    """What the model sees: the real plan handler unwrapping the real outcome (under `plan`, default valid_plan())."""
    context = _Context()
    module.register(context)
    return _r2d2_plan_handler_result(module, context, bridge, plan)


# ----------------------------------------------------------------------- tests
def test_fixture_header_constants_and_sha256() -> None:
    print(f"plan-wire-contract.v2.json sha256={FIXTURE_SHA256} cases={len(CASES)}")
    assert FIXTURE["contractVersion"] == 2
    assert FIXTURE["protocolVersion"] == protocol.PROTOCOL_VERSION
    assert "/" + FIXTURE["hostPrefix"] == protocol.HOST_PREFIX
    assert FIXTURE["origin"] == protocol.REQUIRED_ORIGIN
    constants = FIXTURE["constants"]
    assert constants["claimLeaseSeconds"] == protocol.CLAIM_LEASE_SECONDS
    assert constants["planExecutingLeaseSeconds"] == protocol.PLAN_EXECUTING_LEASE_SECONDS
    assert constants["invokeTimeoutSeconds"] == protocol.INVOKE_TIMEOUT_SECONDS
    assert constants["sessionTtlSeconds"] == protocol.SESSION_TTL_SECONDS
    assert isinstance(constants["executeStartAckTimeoutMs"], int)  # asserted by the bridge leg
    generated = (REPO / "src" / "generated" / "workbookPlanSchema.ts").read_text(encoding="utf-8")
    assert f'export const workbookPlanFingerprint = "{FIXTURE["workbookPlanFingerprint"]}" as const;' in generated
    plugin = load_plugin()
    assert plugin._PLAN_FINGERPRINT == FIXTURE["workbookPlanFingerprint"]
    assert plugin._PLAN_BRIDGE_TOOL == FIXTURE["pluginPlanTool"] == plugin.bridge_client.PLAN_TOOL
    assert plugin._PLAN_TOOL == "sidecar_" + FIXTURE["planTool"]
    assert FIXTURE["messages"]["pluginPlanUncertain"] == plugin.bridge_client.PLAN_UNCERTAIN_MESSAGE


def test_fixture_coverage_guard() -> None:
    ids = [case["id"] for case in CASES]
    assert len(ids) == len(set(ids))
    tags = {tag for case in CASES for tag in case["tags"]}
    missing = set(FIXTURE["requiredTags"]) - tags
    assert not missing, f"required tags missing from the fixture: {sorted(missing)}"
    assert set(FIXTURE["stepKinds"]) == STEP_KINDS
    for case in CASES:
        assert case["sides"] and set(case["sides"]) <= SIDES, case["id"]
        version = case["companionVersion"]
        assert version in ("current", "synthetic", "r2c-cd5d021"), case["id"]
        if version != "current":
            # The Python leg must never claim a synthetic or historical case.
            assert not {"companion", "plugin"} & set(case["sides"]), case["id"]
        for step in case["steps"]:
            assert step["kind"] in STEP_KINDS, (case["id"], step["kind"])
            if "responseText" in step.get("response", {}):
                assert "companion" not in case["sides"] and "plugin" not in case["sides"], case["id"]
    # Every leg replays exactly the cases the independent pinned table assigns it.
    pinned_companion = [cid for cid, sides in EXPECTED_CASE_SIDES.items() if "companion" in sides]
    pinned_plugin = [cid for cid, sides in EXPECTED_CASE_SIDES.items() if "plugin" in sides]
    assert len(pinned_companion) == EXPECTED_COMPANION_CASE_COUNT
    assert len(pinned_plugin) == EXPECTED_PLUGIN_CASE_COUNT
    assert [case["id"] for case in COMPANION_CASES] == pinned_companion
    assert [case["id"] for case in PLUGIN_CASES] == pinned_plugin
    assert len(COMPANION_CASES) == EXPECTED_COMPANION_CASE_COUNT
    assert len(PLUGIN_CASES) == EXPECTED_PLUGIN_CASE_COUNT


def test_fixture_pins_required_tags_and_case_sides_independently() -> None:
    assert len(REQUIRED_TAGS) == 21 and len(set(REQUIRED_TAGS)) == len(REQUIRED_TAGS)
    assert FIXTURE["requiredTags"] == list(REQUIRED_TAGS)
    tags = {tag for case in CASES for tag in case["tags"]}
    missing = [tag for tag in REQUIRED_TAGS if tag not in tags]
    assert not missing, f"pinned required tags missing from the cases: {missing}"
    ids = [case["id"] for case in CASES]
    unknown = [cid for cid in ids if cid not in EXPECTED_CASE_SIDES]
    assert not unknown, f"cases not in EXPECTED_CASE_SIDES (add them deliberately to both legs): {unknown}"
    absent = [cid for cid in EXPECTED_CASE_SIDES if cid not in ids]
    assert not absent, f"pinned cases missing from the fixture: {absent}"
    for case in CASES:
        assert sorted(case["sides"]) == sorted(EXPECTED_CASE_SIDES[case["id"]]), case["id"]


def _responses(case: dict[str, Any]) -> list[tuple[str, str, Any]]:
    return [(step["method"], step["path"], step["response"]) for step in case["steps"]
            if step["kind"] == "pane" and "response" in step]


def test_bridge_only_cases_reuse_only_companion_proven_responses() -> None:
    """A bridge-only case may reuse a companion response only if it is
    byte-identical (in the fixture) to one the Python leg proves elsewhere."""
    by_id = {case["id"]: case for case in CASES}
    checked = 0
    for case in CASES:
        if case["sides"] != ["bridge"] or case["companionVersion"] == "r2c-cd5d021":
            continue
        proven_by = case.get("responseProvenBy")
        assert proven_by, f"{case['id']} needs responseProvenBy"
        proven: list[tuple[str, str, Any]] = []
        for proven_id in proven_by:
            source = by_id[proven_id]
            assert "companion" in source["sides"] and source["companionVersion"] == "current", proven_id
            proven.extend(_responses(source))
        for step in case["steps"]:
            if step["kind"] != "pane" or step.get("syntheticResponse"):
                continue
            entry = (step["method"], step["path"], step["response"])
            assert entry in proven, f"{case['id']}: {entry[0]} {entry[1]} response is not companion-proven"
            checked += 1
        if case["id"] == "C16-truncation/B1-pending":
            source = next(s for s in by_id["C02-completed"]["steps"]
                          if s["kind"] == "pane" and s["path"] == "/excel/pending")
            sample = {f"${k}": v for k, v in FIXTURE["samples"].items()} | _initial_bindings()
            text = case["steps"][1]["response"]["responseText"]
            assert json.dumps(resolve(source["response"]["body"], sample)).startswith(text)
            assert json.dumps(resolve(source["response"]["body"], sample), separators=(",", ":")).startswith(text)
    assert checked > 0


@pytest.mark.parametrize("case", COMPANION_CASES, ids=[case["id"] for case in COMPANION_CASES])
def test_companion_replays_contract_case(case: dict[str, Any]) -> None:
    _Run(case, "companion").run()


@pytest.mark.parametrize("case", PLUGIN_CASES, ids=[case["id"] for case in PLUGIN_CASES])
def test_plugin_replays_contract_case(case: dict[str, Any]) -> None:
    _Run(case, "plugin").run()


# --------------------------------------------------------------------- R4h
# dev3-r4h-release §5.2: the bridge and route non-executing plan payload shapes
# as ONE consumer contract (D-R4-7: both kept), mirrored independently of the TS
# leg. Literals are test-local; nothing here starts a companion or touches the
# network. The plugin consumer must wrap on the envelope ``executed: False``
# only; ``summary``, ``error``, ``category`` and ``result.finalSummary`` are
# display-only, and ``executed: False`` with ``reason: "already_executed"``
# means "not executed AGAIN" (the workbook was changed by an earlier ticket).
import copy  # noqa: E402  (R4h append; stdlib only)

R4H_BRIDGE_NOT_EXECUTED_ERRORS: dict[str, str] = {
    "invalid": "Workbook plan ticket was invalid and was not executed.",
    "rejected": "Workbook plan was not executed.",
}
R4H_ROUTE_NOT_EXECUTED_SUMMARIES: dict[str, str] = {
    "ask": "The workbook plan asked for clarification; nothing was executed.",
    "refuse": "The workbook plan was refused; nothing was executed.",
    "rejected": "The workbook plan was rejected before execution; nothing was executed.",
    "cancelled": "The workbook plan was cancelled before execution; nothing was executed.",
    "not_started": "Workbook plan execution did not start; nothing was executed.",
    "invalid": "The workbook plan ticket was invalid; nothing was executed.",
    "already_executed": "This workbook operation was already applied earlier and was not executed again.",
}
R4H_NON_EXECUTING_REASONS: tuple[str, ...] = (
    "ask", "refuse", "rejected", "cancelled", "not_started", "invalid", "already_executed",
)
R4H_UNCERTAIN_ERROR = "Workbook plan execution ended without a confirmed result."
# (case id, shape, reason) of every fixture plugin_invoke_end with envelope executed False.
R4H_PLUGIN_NOT_EXECUTED = sorted([
    ("C03-not-executed", "route", "ask"), ("C04-duplicate", "route", "ask"),
    ("C05-claimed-executed-true", "route", "not_started"),
    ("C14-4xx-set/result-without-executed", "route", "ask"), ("C14-4xx-set/pane-auth", "route", "ask"),
    ("C14-4xx-set/execute-start-401", "route", "not_started"),
    ("C14-4xx-set/second-invoke-same-turn", "route", "ask"), ("C14-4xx-set/second-invoke-same-turn", "route", "ask"),
    ("C19a-throw-before-ack", "bridge", "rejected"),
    ("C21a-rejected", "route", "rejected"), ("C21b-refuse", "route", "refuse"),
    ("C21c-cancelled", "route", "cancelled"), ("C22-not-started", "route", "not_started"),
    # R4g-2a (F6): the T8 inspection body (no reason) as the companion records it.
    ("C24-inspect-then-execute", "inspect", None), ("C25-inspect-twice", "inspect", None),
    ("C27-inspect-ticket-cannot-execute", "inspect", None),
])
R4H_PLUGIN_UNCERTAIN = sorted([("C19b-throw-after-ack", "bridge"), ("C21d-route-uncertain", "route")])
# Labelled synthetic bodies (P2 only; never sent to the companion).
R4H_SYNTHETIC_BODIES: tuple[tuple[str, dict[str, Any]], ...] = (
    ("synthetic:bridge-invalid", {
        "ok": False, "action": "workbook_plan", "reason": "invalid",
        "error": "Workbook plan ticket was invalid and was not executed.",
    }),
    # R4g-2c-r1 rule W2: the synthetic already_executed body carries a real executor report (the shape of
    # workbookPlanExecutor.ts WorkbookPlanExecutionReport); {"synthetic": true} is no longer a closed body.
    ("synthetic:route-already-executed", {
        "ok": False, "action": "workbook_plan", "reason": "already_executed",
        "summary": "This workbook operation was already applied earlier and was not executed again.",
        "result": {"execution": {
            "outcome": "committed", "transactionId": "5f0c2a8e-3b1d-5e47-9a6c-1d2e3f405162", "operationCount": 1,
            "operations": [{"operationId": "7a1b2c3d-4e5f-5061-8273-a4b5c6d7e8f9", "operationIndex": 0,
                            "effect": "write", "groupIndex": 0, "state": "committed", "status": "applied"}],
            "verification": "not_attempted", "rollback": "not_attempted"}},
    }),
)


def _r4h_bridge_shape(body: Any) -> list[str]:
    """T1 problems (empty list = the bridge non-executing shape)."""
    if not isinstance(body, dict):
        return ["not an object"]
    problems: list[str] = []
    if body.get("ok") is not False:
        problems.append("ok")
    if body.get("action") != "workbook_plan":
        problems.append("action")
    reason = body.get("reason")
    if reason == "invalid":
        if list(body) != ["ok", "action", "reason", "error"]:
            problems.append(f"keys {list(body)}")
    elif reason == "rejected":
        if list(body) != ["ok", "action", "category", "reason", "error"]:
            problems.append(f"keys {list(body)}")
        if body.get("category") != "Execution":
            problems.append("category")
    else:
        return problems + [f"reason {reason!r}"]
    if body.get("error") != R4H_BRIDGE_NOT_EXECUTED_ERRORS[reason]:
        problems.append("error literal")
    return problems


def _r4h_route_shape(body: Any) -> list[str]:
    """T2 problems (empty list = the route non-executing shape, seven reasons)."""
    if not isinstance(body, dict):
        return ["not an object"]
    problems: list[str] = []
    if body.get("ok") is not False:
        problems.append("ok")
    if body.get("action") != "workbook_plan":
        problems.append("action")
    reason = body.get("reason")
    if reason not in R4H_NON_EXECUTING_REASONS:
        return problems + [f"reason {reason!r}"]
    with_result = reason in ("ask", "refuse", "already_executed")
    expected = ["ok", "action", "reason", "summary"] + (["result"] if with_result else [])
    if list(body) != expected:
        problems.append(f"keys {list(body)}")
    if body.get("summary") != R4H_ROUTE_NOT_EXECUTED_SUMMARIES[reason]:
        problems.append("summary literal")
    if with_result:
        result = body.get("result")
        if not isinstance(result, dict):
            problems.append("result")
        elif reason == "already_executed":
            if list(result) != ["execution"] or not isinstance(result.get("execution"), dict):
                problems.append("result.execution")
        else:
            if list(result) != ["outcome", "finalSummary"] or result.get("outcome") != reason:
                problems.append("result keys/outcome")
            final = result.get("finalSummary")
            if not isinstance(final, str) or len(final) > 2048:
                problems.append("result.finalSummary")
    return problems


def _r4h_inspect_shape(body: Any) -> list[str]:
    """R4g-2a T8 problems (empty list = the R4g-1 inspection body; F6, B-A4 per-read shape)."""
    if not isinstance(body, dict):
        return ["not an object"]
    problems: list[str] = []
    if list(body) != ["ok", "action", "phase", "result"]:
        problems.append(f"keys {list(body)}")
    if body.get("ok") is not True or body.get("action") != "workbook_plan" or body.get("phase") != "inspect":
        problems.append("ok/action/phase")
    result = body.get("result")
    if not isinstance(result, dict) or list(result) != ["outcome", "inspectionCount", "reads"]:
        return problems + ["result keys"]
    if result.get("outcome") != "inspected" or result.get("inspectionCount") != 1:
        problems.append("outcome/inspectionCount")
    reads = result.get("reads")
    if not isinstance(reads, list) or not all(isinstance(read, dict) and isinstance(read.get("ok"), bool) for read in reads):
        problems.append("reads")
    return problems


def _r4h_shape(body: Any, where: str) -> str:
    t1, t2, t8 = _r4h_bridge_shape(body), _r4h_route_shape(body), _r4h_inspect_shape(body)
    assert [not t1, not t2, not t8].count(True) == 1, f"{where}: T1 {t1} / T2 {t2} / T8 {t8} (must match exactly one)"
    assert not ("summary" in body and "error" in body), f"{where}: summary and error co-occur"
    return "bridge" if not t1 else "route" if not t2 else "inspect"


def _r4h_bindings() -> dict[str, Any]:
    return {f"${k}": v for k, v in FIXTURE["samples"].items()} | _initial_bindings()


def _r4h_plugin_envelopes_with_plans() -> list[tuple[str, dict[str, Any], dict[str, Any]]]:
    """R4g-2c-r1 rule W1: each recorded envelope with the plan (arguments) of the invoke that produced it."""
    bindings = _r4h_bindings()
    out: list[tuple[str, dict[str, Any], dict[str, Any]]] = []
    for case in CASES:
        plan: Any = None
        for step in case["steps"]:
            if step["kind"] == "plugin_invoke_begin":
                plan = resolve(step["request"], bindings)["arguments"]
            if step["kind"] != "plugin_invoke_end":
                continue
            body = resolve(step.get("response", {}).get("body"), bindings)
            if isinstance(body, dict) and body.get("ok") is True and "data" in body:
                out.append((case["id"], body, plan))
    return out


def _r4h_plugin_envelopes() -> list[tuple[str, dict[str, Any]]]:
    return [(case_id, body) for case_id, body, _plan in _r4h_plugin_envelopes_with_plans()]


def _r4h_view(envelope: dict[str, Any], plan: dict[str, Any] | None = None) -> Any:
    """The real plan handler over a fixed companion envelope (fresh plugin module per view)."""
    frozen = copy.deepcopy(envelope)
    return _model_view(load_plugin(), lambda *_a, **_kw: copy.deepcopy(frozen), plan)


def test_r4h_plugin_wraps_every_non_executing_body_unchanged() -> None:
    seen: list[tuple[str, str, str]] = []
    uncertain: list[tuple[str, str]] = []
    for case_id, envelope, plan in _r4h_plugin_envelopes_with_plans():  # R4g-2c-r1 rule W1: the recorded plan
        data = envelope["data"]
        if envelope.get("executed") is False:
            shape = _r4h_shape(data, case_id)
            view = _r4h_view(envelope, plan)
            assert view == {"executed": False, "result": data}, case_id
            assert list(view) == ["executed", "result"] and list(view["result"]) == list(data), case_id
            seen.append((case_id, shape, data.get("reason")))
        elif envelope.get("executed") is True and isinstance(data, dict) and data.get("state") == "uncertain":
            view = _r4h_view(envelope, plan)
            assert view == data and list(view) == list(data), case_id
            assert "reason" not in view and "summary" not in view, case_id
            assert view.get("error") == R4H_UNCERTAIN_ERROR, case_id
            uncertain.append((case_id, "bridge" if "category" in data else "route"))
    assert sorted(seen) == R4H_PLUGIN_NOT_EXECUTED
    assert ("C19a-throw-before-ack", "bridge", "rejected") in seen  # gap g1: the bridge shape reaches the plugin
    assert sorted(uncertain) == R4H_PLUGIN_UNCERTAIN


_R4G2B_NOT_EXECUTED = {"ok": False, "action": "workbook_plan", "category": "Execution", "reason": "rejected",
                       "error": "Workbook plan was not executed."}  # R4g-2b rule V: the plugin's fixed body


_R4G2C_R1_UUID = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}")
_R4G2C_R1_PAIRS = {("committed", "applied"), ("committed", "verified"), ("committed", "rolled_back"),
                   ("not_committed", "not_started"), ("not_committed", "failed"),
                   ("unknown_within_batch", "unknown_within_batch"), ("unknown_within_batch", "rolled_back")}


def _r4g2c_r1_report(value: Any) -> bool:
    """R4g-2c-r1 rule W2 oracle (independent of the plugin): workbookPlanExecutor.ts WorkbookPlanExecutionReport as
    the route replays it (isReport + validReportInvariants, 1..30 operations, some operation not not_committed)."""
    def count(v: Any) -> bool:
        return type(v) is int and 0 <= v <= 2 ** 53 - 1
    if not isinstance(value, dict):
        return False
    keys = set(value)
    required = {"outcome", "transactionId", "operationCount", "operations", "verification", "rollback"}
    if not required <= keys <= required | {"firstFailedGroup", "firstFailedOperationIndex"}:
        return False
    ops = value["operations"]
    if value["outcome"] not in ("committed", "partial", "stale", "cancelled", "replay", "conflict", "restart_rejected",
                                "rejected", "persistence_unavailable", "persistence_unavailable_after_mutation"):
        return False
    if not (isinstance(value["transactionId"], str) and _R4G2C_R1_UUID.fullmatch(value["transactionId"])):
        return False
    if not (count(value["operationCount"]) and 1 <= value["operationCount"] <= 30 and isinstance(ops, list)
            and len(ops) == value["operationCount"]):
        return False
    if value["verification"] not in ("not_attempted", "reconciled", "reconciliation_failed"):
        return False
    if value["rollback"] not in ("not_attempted", "unavailable", "succeeded", "failed"):
        return False
    if not all(count(value[key]) for key in ("firstFailedGroup", "firstFailedOperationIndex") if key in value):
        return False
    for op in ops:
        if not isinstance(op, dict) or set(op) != {"operationId", "operationIndex", "effect", "groupIndex", "state",
                                                     "status"}:
            return False
        if not (isinstance(op["operationId"], str) and _R4G2C_R1_UUID.fullmatch(op["operationId"])
                and count(op["operationIndex"]) and count(op["groupIndex"])
                and op["effect"] in ("write", "format", "structure", "filter", "calculate")
                and (op["state"], op["status"]) in _R4G2C_R1_PAIRS):
            return False
    states = {op["state"] for op in ops}
    if value["outcome"] == "committed" and states != {"committed"}:
        return False
    if value["outcome"] in ("stale", "cancelled") and states != {"not_committed"}:
        return False
    return states != {"not_committed"}


def _r4g2c_closed_shape(body: Any, phase: str = "execute") -> bool:
    """R4g-2c rule W oracle (independent of the plugin): a closed R4h T1/T2 refusal shape by key set, reason and
    value types (display text is not pinned by the plugin; sizes are far below its caps here).
    R4g-2c-r1 rule W2: result-bearing refusals only in the execute phase; already_executed needs a real report."""
    if not isinstance(body, dict) or body.get("ok") is not False or body.get("action") != "workbook_plan":
        return False
    if "result" in body and phase != "execute":
        return False
    shapes = {
        frozenset({"ok", "action", "reason", "error"}): {"invalid"},
        frozenset({"ok", "action", "category", "reason", "error"}): {"rejected"},
        frozenset({"ok", "action", "reason", "summary"}): {"rejected", "cancelled", "not_started", "invalid"},
        frozenset({"ok", "action", "reason", "summary", "result"}): {"ask", "refuse", "already_executed"},
    }
    reason = body.get("reason")
    if reason not in shapes.get(frozenset(body), set()):
        return False
    if not all(isinstance(body[key], str) for key in ("error", "category", "summary") if key in body):
        return False
    if "result" not in body:
        return True
    result = body["result"]
    if reason == "already_executed":
        return isinstance(result, dict) and list(result) == ["execution"] and _r4g2c_r1_report(result["execution"])
    return (isinstance(result, dict) and sorted(result) == ["finalSummary", "outcome"] and result["outcome"] == reason
            and isinstance(result["finalSummary"], str))


def test_r4h_plugin_discriminates_on_envelope_not_display_text() -> None:
    # R4g-2c-r1 rule W1: each body is viewed under the plan of the invoke that produced it (synthetic: execute).
    bodies = [(case_id, envelope["data"], plan) for case_id, envelope, plan in _r4h_plugin_envelopes_with_plans()
              if envelope.get("executed") is False]
    assert len(bodies) == len(R4H_PLUGIN_NOT_EXECUTED)
    bodies += [(label, copy.deepcopy(body), valid_execute_plan()) for label, body in R4H_SYNTHETIC_BODIES]
    for label, body in R4H_SYNTHETIC_BODIES:
        _r4h_shape(body, label)  # the synthetic bodies are themselves contract-valid
    checked = 0
    refused_t8 = 0
    refused_refusal = 0  # R4g-2c rule W
    for label, body, plan in bodies:
        phase = "inspect" if plan.get("outcome") == "inspect" else "execute"
        variants: list[tuple[str, dict[str, Any]]] = []
        for key in ("summary", "error", "category", "result"):
            if key in body:
                dropped = copy.deepcopy(body)
                del dropped[key]
                variants.append((f"drop {key}", dropped))
                blanked = copy.deepcopy(body)
                blanked[key] = ""
                variants.append((f"blank {key}", blanked))
        for reason in R4H_NON_EXECUTING_REASONS:
            swapped = copy.deepcopy(body)
            swapped["reason"] = reason
            variants.append((f"reason {reason}", swapped))
        without_reason = copy.deepcopy(body)
        without_reason.pop("reason", None)  # envelope-only probe; not a valid contract body (T8 has none)
        variants.append(("drop reason", without_reason))
        for name, variant in variants:
            where = f"{label} / {name}"
            envelope = {"ok": True, "ticketId": "r4h-ticket", "requestId": "r4h-request", "data": variant, "executed": False}
            # R4g-2b rule V: the plugin re-verifies inspection bodies (T8); a T8 variant that no longer conforms is
            # replaced by the fixed not-executed body. Every refusal variant (ok False) is still wrapped unchanged.
            # R4g-2c rule W: a refusal variant (ok False) is wrapped unchanged only while it keeps a closed R4h T1/T2
            # shape (_r4g2c_closed_shape); any other refusal variant is now the fixed body as well (counted).
            if variant == body or _r4g2c_closed_shape(variant, phase):
                expected = variant
            elif variant.get("ok") is False:
                expected = _R4G2B_NOT_EXECUTED
                refused_refusal += 1
            else:
                expected = _R4G2B_NOT_EXECUTED
                refused_t8 += 1
            assert _r4h_view(envelope, plan) == {"executed": False, "result": expected}, where
            executed = dict(envelope, executed=True)
            assert _r4h_view(executed, plan) == variant, where
            checked += 1
    assert checked > 15 * 7
    assert refused_t8 > 0
    assert refused_refusal > 0  # R4g-2c rule W


def test_r4g2c_r1_wire_recorded_refusals_follow_the_phase_of_the_invoke() -> None:
    """R4g-2c-r1 (B1 phase restriction): every recorded executed:false body passes verbatim under the plan that
    produced it (pinned above); under the other phase only the bodies without a result (T1 bridge bodies and the
    summary-only T2 reasons) still pass, and the result-bearing ones (ask/refuse; the synthetic already_executed) and
    the T8 inspection bodies are the fixed body."""
    seen = {"kept": 0, "fixed": 0}
    bodies = [(case_id, envelope, plan) for case_id, envelope, plan in _r4h_plugin_envelopes_with_plans()
              if envelope.get("executed") is False]
    bodies += [(label, {"ok": True, "ticketId": "r4h-ticket", "requestId": "r4h-request", "data": copy.deepcopy(body),
                        "executed": False}, valid_execute_plan()) for label, body in R4H_SYNTHETIC_BODIES]
    for label, envelope, plan in bodies:
        assert _r4h_view(envelope, plan) == {"executed": False, "result": envelope["data"]}, label
        other = valid_execute_plan() if plan.get("outcome") == "inspect" else valid_plan()
        data = envelope["data"]
        keeps = data.get("ok") is False and "result" not in data
        expected = data if keeps else _R4G2B_NOT_EXECUTED
        assert _r4h_view(envelope, other) == {"executed": False, "result": expected}, label
        seen["kept" if keeps else "fixed"] += 1
    assert seen["kept"] > 0 and seen["fixed"] > 0, seen


def test_r4h_non_executing_contract_literals_match_fixture_and_closed_sets() -> None:
    bindings = _r4h_bindings()
    bodies: list[tuple[str, dict[str, Any]]] = []
    for case in CASES:
        for step in case["steps"]:
            if step["kind"] == "pane" and step.get("path") == "/excel/result" and step.get("only") != "python":
                posted = resolve(step.get("body"), bindings)
                if isinstance(posted, dict) and posted.get("executed") is False:
                    bodies.append((case["id"], posted["result"]))
    bodies += [(case_id, envelope["data"]) for case_id, envelope in _r4h_plugin_envelopes()
               if envelope.get("executed") is False]
    route_reasons: set[str] = set()
    bridge_reasons: set[str] = set()
    for case_id, body in bodies:
        shape = _r4h_shape(body, case_id)
        if shape == "route":
            assert body["summary"] == R4H_ROUTE_NOT_EXECUTED_SUMMARIES[body["reason"]], case_id
            route_reasons.add(body["reason"])
        elif shape == "inspect":
            continue  # T8 (R4g-2a, F6): no reason; pinned by test_r4g2_w7
        else:
            assert body["error"] == R4H_BRIDGE_NOT_EXECUTED_ERRORS[body["reason"]], case_id
            bridge_reasons.add(body["reason"])
    assert route_reasons <= set(R4H_NON_EXECUTING_REASONS)
    assert bridge_reasons == {"invalid", "rejected"}
    assert route_reasons == {"ask", "refuse", "rejected", "cancelled", "not_started"}
    # g2 (route invalid, TS H4 drives it live) and g4 (already_executed, owned by the R4e tests).
    assert set(R4H_NON_EXECUTING_REASONS) - route_reasons == {"invalid", "already_executed"}
    assert sorted(R4H_ROUTE_NOT_EXECUTED_SUMMARIES) == sorted(R4H_NON_EXECUTING_REASONS)
    assert len(R4H_NON_EXECUTING_REASONS) == len(set(R4H_NON_EXECUTING_REASONS)) == 7


# --------------------------------------------------------------------- R4g-2a
# dev3-r4g2-release.md §5.1 W1/W5/W7/W8 (F6, F7): fixture-wide pins, independent of the fixture's own metadata.
R4G2_V1_PATH = REPO / "src" / "__tests__" / "fixtures" / "companion-wire" / "plan-wire-contract.v1.json"
R4G2_V1_SHA256 = "a1d9f433b2922b302ae6cf309153acaddb2a9b61dc9cc28dba53f06f5daf5662"
R4G2_TRUSTED_KEYS = ["turnOperationId", "workbookPlanFingerprint", "phase", "turnKey"]
# Deliberately malformed or v1 trusted objects (bridge-only synthetic tickets and the v1 plugin invoke).
R4G2_NON_V2_TRUSTED = {
    "C18-invalid-plan-ticket/trusted-extra-key": ["turnOperationId", "workbookPlanFingerprint", "phase", "turnKey", "extra"],
    "C18-invalid-plan-ticket/trusted-v1-two-keys": ["turnOperationId", "workbookPlanFingerprint"],
    "MV4-v1-plugin-envelope": ["turnOperationId", "workbookPlanFingerprint"],
}
R4G2_INSPECT_CASES = ["C24-inspect-then-execute", "C25-inspect-twice", "C27-inspect-ticket-cannot-execute",
                      "C30-inspect-ticket-gate-on"]
R4G2_J2 = {"session": "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa", "turn": "r2a-real-uuid5",
           "turnKey": "5a83b37f-b6e2-5031-8d37-df04bbf71b8a",
           "inspect": "cfed8ba1-049f-58dd-ac4d-dad4417a0ca2", "execute": "7244df56-2215-500f-98d1-2798e9313f70"}


def _r4g2_trusted_objects() -> list[tuple[str, dict[str, Any]]]:
    out: list[tuple[str, dict[str, Any]]] = []

    def walk(case_id: str, node: Any) -> None:
        if isinstance(node, dict):
            for key, value in node.items():
                if key == "trusted" and isinstance(value, dict):
                    out.append((case_id, value))
                else:
                    walk(case_id, value)
        elif isinstance(node, list):
            for value in node:
                walk(case_id, value)

    for case in CASES:
        walk(case["id"], case)
    return out


def test_r4g2_w1_v1_fixture_is_kept_byte_identical_for_provenance() -> None:
    raw = R4G2_V1_PATH.read_bytes()
    assert hashlib.sha256(raw).hexdigest() == R4G2_V1_SHA256
    assert json.loads(raw)["contractVersion"] == 1
    assert FIXTURE_PATH.name == "plan-wire-contract.v2.json" and FIXTURE_SHA256 != R4G2_V1_SHA256


def test_r4g2_w5_every_trusted_object_carries_the_four_v2_keys_in_order() -> None:
    seen_exceptions: dict[str, list[str]] = {}
    count = 0
    for case_id, trusted in _r4g2_trusted_objects():
        count += 1
        if case_id in R4G2_NON_V2_TRUSTED:
            seen_exceptions[case_id] = list(trusted)
            continue
        assert list(trusted) == R4G2_TRUSTED_KEYS, (case_id, list(trusted))
        assert trusted["phase"] in ("inspect", "execute") or case_id == "C18-invalid-plan-ticket/trusted-phase-unknown", case_id
        assert trusted["turnKey"] != trusted["turnOperationId"] or case_id == "C18-invalid-plan-ticket/trusted-turnkey-equals-op", case_id
    assert seen_exceptions == R4G2_NON_V2_TRUSTED
    assert count == 78 + 21  # 78 migrated v1 objects + 21 in the new cases (recorded)


def test_r4g2_w7_inspection_body_class_t8_and_t4_prime_closure() -> None:
    bindings = _r4h_bindings()
    inspect_cases: list[str] = []
    for case in CASES:
        for step in case["steps"]:
            bodies: list[Any] = []
            if step["kind"] == "pane" and step.get("path") == "/excel/result":
                posted = resolve(step.get("body"), bindings)
                if isinstance(posted, dict) and posted.get("executed") is False:
                    bodies.append(posted["result"])
            if step["kind"] == "plugin_invoke_end":
                body = resolve(step.get("response", {}).get("body"), bindings)
                if isinstance(body, dict) and body.get("executed") is False and "data" in body:
                    bodies.append(body["data"])
            for body in bodies:
                shape = _r4h_shape(body, case["id"])  # T4': exactly one of T1/T2/T8
                if shape == "inspect":
                    assert not {"reason", "summary", "state", "error", "execution"} & set(body), case["id"]
                    assert not {"reason", "summary", "state", "error", "execution"} & set(body["result"]), case["id"]
                    if case["id"] not in inspect_cases:
                        inspect_cases.append(case["id"])
    assert inspect_cases == R4G2_INSPECT_CASES
    # g2m14: a T8 body that gains `reason` (or loses a key) is no longer T8.
    t8 = next(resolve(step["body"], bindings)["result"] for step in CASES[[c["id"] for c in CASES].index("C24-inspect-then-execute")]["steps"]
              if step["kind"] == "pane" and step.get("path") == "/excel/result")
    assert _r4h_inspect_shape(t8) == []
    assert _r4h_inspect_shape(dict(t8, reason="rejected")) != []
    assert _r4h_inspect_shape({k: v for k, v in t8.items() if k != "phase"}) != []
    assert _r4h_inspect_shape(dict(t8, result=dict(t8["result"], reads=[[1]]))) != []


def test_r4g2_w8_j2_vectors_through_the_real_plugin_path_of_c24() -> None:
    by_id = {case["id"]: case for case in CASES}
    c24 = [step["request"] for step in by_id["C24-inspect-then-execute"]["steps"] if step["kind"] == "plugin_invoke_begin"]
    module = load_plugin()
    context = _Context()
    module.register(context)
    hook = context.hooks["pre_tool_call"]
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == module._PLAN_TOOL)
    sent: list[dict[str, Any]] = []

    from scripts.tests.test_office_plugin_plan_contract import _FakeBridgeResponse

    def recording_urlopen(request: Any, timeout: Any = None):
        sent.append(json.loads(request.data))
        return _FakeBridgeResponse(b'{"ok":true,"data":{"accepted":true}}')

    with (
        patch.object(module.bridge_client, "_connection", return_value=("http://127.0.0.1:9", Path("/never-read"))),
        patch.object(module.bridge_client, "_validate_key", return_value=PLUGIN_KEY),
        patch.object(module.bridge_client.urllib.request, "urlopen", recording_urlopen),
    ):
        for request in c24:
            plan = request["arguments"]
            assert hook(tool_name=module._PLAN_TOOL, args=plan, session_id=R4G2_J2["session"], turn_id=R4G2_J2["turn"],
                        tool_call_id="call-w8", api_request_id="request-w8", task_id="task-w8") is None
            handler(plan, session_id=R4G2_J2["session"])
    assert [list(body["trusted"]) for body in sent] == [R4G2_TRUSTED_KEYS, R4G2_TRUSTED_KEYS]
    assert [list(body["trusted"]) for body in sent] == [list(request["trusted"]) for request in c24]
    assert [body["trusted"]["phase"] for body in sent] == [request["trusted"]["phase"] for request in c24] == ["inspect", "execute"]
    assert [body["requestId"] for body in sent] == [R4G2_J2["inspect"], R4G2_J2["execute"]]
    assert [body["trusted"]["turnOperationId"] for body in sent] == [R4G2_J2["inspect"], R4G2_J2["execute"]]
    assert {body["trusted"]["turnKey"] for body in sent} == {R4G2_J2["turnKey"]}
