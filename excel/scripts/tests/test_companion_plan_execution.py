"""Focused Task 12 companion transport tests."""
from __future__ import annotations

import json
import threading
import time
import urllib.error
import urllib.request

from desk_plugin.companion.server import ExcelCompanionServer

import uuid as _r4g2_uuid  # noqa: E402  (R4g-2a rule M)
_R4G2_TURN_KEY_NS = _r4g2_uuid.UUID("4b4b0000-0000-4000-8000-0000000000a2")


def _r4g2_turn_key(operation_id: str) -> str:
    """R4g-2a rule M: a per-operation turnKey (never equal to the operation id) for a migrated v1 envelope."""
    return str(_r4g2_uuid.uuid5(_R4G2_TURN_KEY_NS, str(operation_id)))

PLUGIN_KEY = "task12-plugin-key-" + ("x" * 32)
FINGERPRINT = "b1f1b6353c9fe49d11d9dcdca7e41a97c06d7535cab04bb29d87fdc2d3d3d8ac"
OTHER_FINGERPRINT = "0" * 64
PLAN = {
    "version": 1,
    "outcome": "ask",
    "message": "No mutation in this transport fixture",
}


def _request(
    server: ExcelCompanionServer,
    method: str,
    path: str,
    body: dict | None = None,
    headers: dict[str, str] | None = None,
) -> tuple[int, dict]:
    raw = None if body is None else json.dumps(body, separators=(",", ":")).encode()
    request = urllib.request.Request(
        f"http://{server.server_address()[0]}:{server.server_address()[1]}{path}",
        method=method,
        data=raw,
        headers={
            **({"Content-Type": "application/json"} if raw is not None else {}),
            **(headers or {}),
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=3) as response:
            data = response.read()
            return response.status, json.loads(data) if data else {}
    except urllib.error.HTTPError as error:
        data = error.read()
        return error.code, json.loads(data) if data else {}


def _session(server: ExcelCompanionServer, fingerprint: str = FINGERPRINT) -> dict:
    status, body = _request(
        server,
        "POST",
        "/excel/session",
        {"protocolVersion": "1.0", "documentId": "doc", "paneId": "pane", "workbookPlanFingerprint": fingerprint},
        {"Origin": "https://localhost:3000"},
    )
    assert status == 201, body
    return body


def _pane_headers(session: dict) -> dict[str, str]:
    return {
        "Origin": "https://localhost:3000",
        "X-Hermes-Session-Key": session["sessionKey"],
        "X-Hermes-Session-Id": session["sessionId"],
        "X-Hermes-Pane-Id": "pane",
    }


def _invoke_async(server: ExcelCompanionServer, session: dict, request_id: str, trusted: dict | None):
    result: dict[str, object] = {}

    def run() -> None:
        status, body = _request(
            server,
            "POST",
            "/excel/invoke",
            {
                "protocolVersion": "1.0",
                "sessionId": session["sessionId"],
                "requestId": request_id,
                "tool": "execute_plan",
                "arguments": PLAN,
                **({"trusted": trusted} if trusted is not None else {}),
            },
            {"X-Hermes-Plugin-Key": PLUGIN_KEY},
        )
        result.update(status=status, body=body)

    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    return thread, result


def _wait_for_ticket(server: ExcelCompanionServer, session_id: str) -> None:
    deadline = time.monotonic() + 1
    session = server.sessions.get(session_id)
    assert session is not None
    while not session.has_ticket() and time.monotonic() < deadline:
        time.sleep(0.005)
    assert session.has_ticket()


def test_plan_transport_maps_internal_route_and_preserves_private_trust() -> None:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=2)
    server.start()
    try:
        session = _session(server)
        trusted = {"turnOperationId": "11111111-1111-4111-8111-111111111111", "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key("11111111-1111-4111-8111-111111111111")}
        thread, result = _invoke_async(server, session, "11111111-1111-4111-8111-111111111111", trusted)
        _wait_for_ticket(server, session["sessionId"])
        status, pending = _request(server, "GET", "/excel/pending", headers=_pane_headers(session))
        assert status == 200
        assert pending["tool"] == "excel_execute_plan"
        assert pending["arguments"] == PLAN
        assert pending["trusted"] == trusted
        # R2d-2 contract change: a plan result never acknowledged as executing
        # must carry executed:false; an undisposed or "executed" result from
        # claimed is refused without changing state.
        assert _request(
            server, "POST", "/excel/result",
            {"ticketId": pending["ticketId"], "result": {"ok": True}},
            _pane_headers(session),
        )[0] == 400
        assert _request(
            server, "POST", "/excel/result",
            {"ticketId": pending["ticketId"], "executed": True, "result": {"ok": True}},
            _pane_headers(session),
        )[0] == 409
        assert server.sessions.get(session["sessionId"]).get_ticket().state == "claimed"
        result_status, _ = _request(
            server,
            "POST",
            "/excel/result",
            {"ticketId": pending["ticketId"], "executed": False, "result": {"ok": True}},
            _pane_headers(session),
        )
        assert result_status == 200
        thread.join(2)
        assert not thread.is_alive()
        assert result["status"] == 200
        assert result["body"]["executed"] is False
        assert result["body"]["data"] == {"ok": True}
    finally:
        server.stop()


def test_second_plan_ticket_in_same_trusted_turn_is_rejected_before_pending() -> None:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=2)
    server.start()
    try:
        session = _session(server)
        trusted = {"turnOperationId": "33333333-3333-4333-8333-333333333333", "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key("33333333-3333-4333-8333-333333333333")}
        first, first_result = _invoke_async(server, session, "33333333-3333-4333-8333-333333333333", trusted)
        _wait_for_ticket(server, session["sessionId"])
        status, pending = _request(server, "GET", "/excel/pending", headers=_pane_headers(session))
        assert status == 200
        assert _request(server, "POST", "/excel/result", {"ticketId": pending["ticketId"], "executed": False, "result": {"ok": True}}, _pane_headers(session))[0] == 200
        first.join(2)
        assert first_result["status"] == 200

        second, second_result = _invoke_async(server, session, "55555555-5555-4555-8555-555555555555", trusted)
        second.join(2)
        assert not second.is_alive()
        assert second_result["status"] == 400
        assert not server.sessions.get(session["sessionId"]).has_ticket()
    finally:
        server.stop()


def test_plan_fingerprint_and_trust_metadata_fail_before_ticket_creation() -> None:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=0.2)
    server.start()
    try:
        session = _session(server)
        for trusted in (
            {"turnOperationId": "66666666-6666-4666-8666-666666666666", "workbookPlanFingerprint": OTHER_FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key("66666666-6666-4666-8666-666666666666")},
            None,
        ):
            thread, result = _invoke_async(server, session, "66666666-6666-4666-8666-666666666666", trusted)
            thread.join(2)
            assert not thread.is_alive()
            assert result["status"] == 400 if trusted is None else 409
            assert not server.sessions.get(session["sessionId"]).has_ticket()
    finally:
        server.stop()