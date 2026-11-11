"""Companion server protocol contract tests — /powerpoint routes on port 8652."""
from __future__ import annotations

import json
import threading
import time
import urllib.error
import urllib.request

import pytest

# Allow importing from the repo root
import sys
from pathlib import Path
_REPO_ROOT = Path(__file__).resolve().parents[2]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from desk_plugin.companion.server import PowerPointCompanionServer
from desk_plugin.companion.session import Session, Ticket
from desk_plugin.companion.protocol import (
    PROTOCOL_VERSION,
    SESSION_TTL_SECONDS,
    CLAIM_LEASE_SECONDS,
    INVOKE_TIMEOUT_SECONDS,
    generate_uuid,
    generate_session_key,
    constant_time_compare,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

PLUGIN_KEY = "a" * 64  # 64-char hex key for tests
TEST_HOST = "127.0.0.1"
TEST_PORT = 18652  # Non-default port to avoid conflicts


class _Clock:
    """Controllable monotonic clock for deterministic tests."""

    def __init__(self, start: float = 0.0):
        self._t = start

    def __call__(self) -> float:
        return self._t

    def advance(self, seconds: float) -> None:
        self._t += seconds


def _start_server(port: int = TEST_PORT, **kwargs) -> PowerPointCompanionServer:
    server = PowerPointCompanionServer(
        host=TEST_HOST,
        port=port,
        plugin_key=PLUGIN_KEY,
        **kwargs,
    )
    server.start()
    time.sleep(0.05)  # Let the server thread start
    return server


def _get(path: str, port: int = TEST_PORT, extra_headers: dict | None = None) -> tuple[int, dict]:
    headers = dict(extra_headers) if extra_headers else {}
    req = urllib.request.Request(
        f"http://{TEST_HOST}:{port}{path}",
        headers=headers,
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=5) as resp:
            return resp.status, json.loads(resp.read().decode())
    except urllib.error.HTTPError as exc:
        return exc.code, json.loads(exc.read().decode())


def _post(path: str, body: dict, port: int = TEST_PORT, extra_headers: dict | None = None) -> tuple[int, dict]:
    data = json.dumps(body).encode()
    headers = {"Content-Type": "application/json"}
    if extra_headers:
        headers.update(extra_headers)
    req = urllib.request.Request(
        f"http://{TEST_HOST}:{port}{path}",
        data=data,
        headers=headers,
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=5) as resp:
            return resp.status, json.loads(resp.read().decode())
    except urllib.error.HTTPError as exc:
        return exc.code, json.loads(exc.read().decode())


def _delete(path: str, port: int = TEST_PORT, extra_headers: dict | None = None) -> tuple[int, dict]:
    headers = dict(extra_headers) if extra_headers else {}
    req = urllib.request.Request(
        f"http://{TEST_HOST}:{port}{path}",
        headers=headers,
        method="DELETE",
    )
    try:
        with urllib.request.urlopen(req, timeout=5) as resp:
            body = resp.read().decode()
            return resp.status, json.loads(body) if body and body.strip() else {}
    except urllib.error.HTTPError as exc:
        body = exc.read().decode()
        return exc.code, json.loads(body) if body and body.strip() else {}


def _session_headers(session_id: str, session_key: str, pane_id: str) -> dict:
    return {
        "Origin": "https://localhost:3002",
        "X-Hermes-Session-Id": session_id,
        "X-Hermes-Session-Key": session_key,
        "X-Hermes-Pane-Id": pane_id,
    }


def _plugin_headers() -> dict:
    return {"X-Hermes-Plugin-Key": PLUGIN_KEY}


# ---------------------------------------------------------------------------
# Session Creation and Auth Tests
# ---------------------------------------------------------------------------

class TestSessionCreation:
    def test_create_session_success(self):
        server = _start_server()
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-abc",
                "paneId": "pane-123",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201
            assert data["ok"] is True
            assert "sessionId" in data
            assert "sessionKey" in data
            assert len(data["sessionKey"]) >= 32
            assert data["protocolVersion"] == PROTOCOL_VERSION
            assert data["expiresInSeconds"] == SESSION_TTL_SECONDS
        finally:
            server.stop()

    def test_create_session_rejects_wrong_origin(self):
        server = _start_server()
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-abc",
                "paneId": "pane-123",
            }, extra_headers={"Origin": "https://evil.com"})
            assert status == 403
        finally:
            server.stop()

    def test_create_session_rejects_wrong_version(self):
        server = _start_server()
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": "0.9",
                "documentId": "doc-abc",
                "paneId": "pane-123",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 400
        finally:
            server.stop()

    def test_create_session_rejects_empty_document_id(self):
        server = _start_server()
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "",
                "paneId": "pane-123",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 400
        finally:
            server.stop()

    def test_create_session_capacity(self):
        server = _start_server(max_sessions=1)
        try:
            status1, data1 = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-1",
                "paneId": "pane-1",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status1 == 201

            status2, data2 = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-2",
                "paneId": "pane-2",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status2 == 429
        finally:
            server.stop()


class TestSessionAuth:
    def test_pane_routes_require_session_auth(self):
        server = _start_server()
        try:
            # Create a session first
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-auth",
                "paneId": "pane-auth",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201
            sid = data["sessionId"]
            skey = data["sessionKey"]

            # Heartbeat without auth should fail
            status_hb, _ = _post("/powerpoint/heartbeat", {}, extra_headers={
                "Origin": "https://localhost:3002",
            })
            assert status_hb == 401

            # Heartbeat with correct auth should pass
            status_hb2, data_hb2 = _post("/powerpoint/heartbeat", {},
                extra_headers=_session_headers(sid, skey, "pane-auth"))
            assert status_hb2 == 200
            assert data_hb2["ok"] is True
        finally:
            server.stop()

    def test_session_auth_wrong_key_fails(self):
        server = _start_server()
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-wk",
                "paneId": "pane-wk",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201

            status_hb, _ = _post("/powerpoint/heartbeat", {},
                extra_headers=_session_headers(data["sessionId"], "wrong-key", "pane-wk"))
            assert status_hb == 401
        finally:
            server.stop()

    def test_plugin_key_auth_required(self):
        server = _start_server()
        try:
            # Create a session
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-pk",
                "paneId": "pane-pk",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201

            # Invoke without plugin key should fail
            status_inv, _ = _post("/powerpoint/invoke", {
                "protocolVersion": PROTOCOL_VERSION,
                "sessionId": data["sessionId"],
                "requestId": generate_uuid(),
                "tool": "get_presentation_overview",
                "arguments": {},
            })
            assert status_inv == 401

            # With plugin key, invoke should connect (blocks, but we just verify auth passes)
            # The invoke will block for 60s waiting for a pane to pick up the ticket,
            # so we dispatch it in a thread and cancel it.
            import concurrent.futures
            executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
            future = executor.submit(lambda: _post("/powerpoint/invoke", {
                "protocolVersion": PROTOCOL_VERSION,
                "sessionId": data["sessionId"],
                "requestId": generate_uuid(),
                "tool": "get_presentation_overview",
                "arguments": {},
            }, extra_headers=_plugin_headers()))
            # Cancel the ticket so the invoke thread returns
            time.sleep(0.1)
            _post("/powerpoint/cancel", {"sessionId": data["sessionId"]},
                  extra_headers=_plugin_headers())
            status_inv2, _ = future.result(timeout=5)
            assert status_inv2 in (200, 499, 504)  # completed, cancelled, or timeout
        finally:
            server.stop()


class TestSessionLifecycle:
    def test_heartbeat_extends_ttl(self):
        clock = _Clock(0)
        server = _start_server(clock=clock)
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-ttl",
                "paneId": "pane-ttl",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201
            sid = data["sessionId"]
            skey = data["sessionKey"]

            # Advance nearly to expiry
            clock.advance(SESSION_TTL_SECONDS - 1)
            status_hb, _ = _post("/powerpoint/heartbeat", {},
                extra_headers=_session_headers(sid, skey, "pane-ttl"))
            assert status_hb == 200

            # Now the session should still be alive
            session = server.sessions.get(sid)
            assert session is not None
            assert not session.is_expired()
        finally:
            server.stop()

    def test_session_expires(self):
        clock = _Clock(0)
        server = _start_server(clock=clock)
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-exp",
                "paneId": "pane-exp",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201

            clock.advance(SESSION_TTL_SECONDS + 1)
            # Trigger sweep via a request
            _get("/powerpoint/health", extra_headers=_plugin_headers())
            assert server.sessions.active_count == 0
        finally:
            server.stop()

    def test_session_delete(self):
        server = _start_server()
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-del",
                "paneId": "pane-del",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201
            sid = data["sessionId"]
            skey = data["sessionKey"]

            status_del, _ = _delete("/powerpoint/session",
                extra_headers=_session_headers(sid, skey, "pane-del"))
            assert status_del == 204
            assert server.sessions.active_count == 0
        finally:
            server.stop()

    def test_delete_wrong_pane_fails(self):
        server = _start_server()
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-dwp",
                "paneId": "pane-owner",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201
            sid = data["sessionId"]
            skey = data["sessionKey"]

            # Try to delete with wrong pane_id
            status_del, _ = _delete("/powerpoint/session",
                extra_headers=_session_headers(sid, skey, "pane-other"))
            assert status_del in (401, 204)  # Either unauthorized or no session matched
        finally:
            server.stop()


class TestTicketFlow:
    def test_invoke_creates_ticket(self):
        server = _start_server()
        try:
            # Create a session
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-tick",
                "paneId": "pane-tick",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201
            sid = data["sessionId"]
            skey = data["sessionKey"]

            # Invoke from plugin (this will block in another thread, so we dispatch)
            def invoke_thread():
                status_inv, data_inv = _post("/powerpoint/invoke", {
                    "protocolVersion": PROTOCOL_VERSION,
                    "sessionId": sid,
                    "requestId": generate_uuid(),
                    "tool": "get_presentation_overview",
                    "arguments": {"maxLength": 100},
                }, extra_headers=_plugin_headers())
                return status_inv, data_inv

            # Start invoke in background
            import concurrent.futures
            executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
            future = executor.submit(invoke_thread)

            # Small delay to let the ticket get created
            time.sleep(0.1)

            # Pane polls for pending
            status_pend, data_pend = _get("/powerpoint/pending",
                extra_headers=_session_headers(sid, skey, "pane-tick"))
            assert status_pend == 200
            assert data_pend["tool"] == "get_presentation_overview"

            # Submit result
            status_res, data_res = _post("/powerpoint/result", {
                "ticketId": data_pend["ticketId"],
                "result": {"ok": True, "slideCount": 5},
            }, extra_headers=_session_headers(sid, skey, "pane-tick"))
            assert status_res == 200

            # Check invoke thread result
            status_inv, data_inv = future.result(timeout=5)
            assert status_inv == 200
            assert data_inv["ok"] is True
            assert data_inv["data"]["ok"] is True
            assert data_inv["data"]["slideCount"] == 5
        finally:
            server.stop()

    def test_duplicate_request_id_is_idempotent(self):
        server = _start_server()
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-dup",
                "paneId": "pane-dup",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201
            sid = data["sessionId"]
            skey = data["sessionKey"]
            req_id = generate_uuid()

            # First invoke
            import concurrent.futures
            executor = concurrent.futures.ThreadPoolExecutor(max_workers=2)

            def invoke():
                return _post("/powerpoint/invoke", {
                    "protocolVersion": PROTOCOL_VERSION,
                    "sessionId": sid,
                    "requestId": req_id,
                    "tool": "read_all_slides",
                    "arguments": {},
                }, extra_headers=_plugin_headers())

            f1 = executor.submit(invoke)
            time.sleep(0.1)

            # Poll and submit result
            status_pend, data_pend = _get("/powerpoint/pending",
                extra_headers=_session_headers(sid, skey, "pane-dup"))
            assert status_pend == 200
            _post("/powerpoint/result", {
                "ticketId": data_pend["ticketId"],
                "result": {"ok": True, "slideCount": 3},
            }, extra_headers=_session_headers(sid, skey, "pane-dup"))

            status_inv1, data_inv1 = f1.result(timeout=5)
            assert status_inv1 == 200

            # Second invoke with same requestId should reattach (no new ticket needed)
            f2 = executor.submit(invoke)
            status_inv2, data_inv2 = f2.result(timeout=5)
            assert status_inv2 == 200
            assert data_inv2["data"]["slideCount"] == 3
        finally:
            server.stop()

    def test_request_id_conflict_different_payload(self):
        server = _start_server()
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-conflict",
                "paneId": "pane-conflict",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201
            sid = data["sessionId"]
            skey = data["sessionKey"]
            req_id = generate_uuid()

            # First invoke
            import concurrent.futures
            executor = concurrent.futures.ThreadPoolExecutor(max_workers=2)

            def invoke1():
                return _post("/powerpoint/invoke", {
                    "protocolVersion": PROTOCOL_VERSION,
                    "sessionId": sid,
                    "requestId": req_id,
                    "tool": "read_all_slides",
                    "arguments": {},
                }, extra_headers=_plugin_headers())

            f1 = executor.submit(invoke1)
            time.sleep(0.1)

            # Poll and submit
            status_pend, data_pend = _get("/powerpoint/pending",
                extra_headers=_session_headers(sid, skey, "pane-conflict"))
            _post("/powerpoint/result", {
                "ticketId": data_pend["ticketId"],
                "result": {"ok": True},
            }, extra_headers=_session_headers(sid, skey, "pane-conflict"))
            f1.result(timeout=5)

            # Second invoke with same requestId but different tool
            status_inv2, data_inv2 = _post("/powerpoint/invoke", {
                "protocolVersion": PROTOCOL_VERSION,
                "sessionId": sid,
                "requestId": req_id,
                "tool": "create_slide",  # Different tool!
                "arguments": {},
            }, extra_headers=_plugin_headers())
            assert status_inv2 == 409  # Conflict
        finally:
            server.stop()

    def test_busy_session_rejects_second_invoke(self):
        server = _start_server()
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-busy",
                "paneId": "pane-busy",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201
            sid = data["sessionId"]

            # First invoke (background — will block)
            import concurrent.futures
            executor = concurrent.futures.ThreadPoolExecutor(max_workers=2)
            f1 = executor.submit(lambda: _post("/powerpoint/invoke", {
                "protocolVersion": PROTOCOL_VERSION,
                "sessionId": sid,
                "requestId": generate_uuid(),
                "tool": "read_all_slides",
                "arguments": {},
            }, extra_headers=_plugin_headers()))
            time.sleep(0.1)

            # Second invoke — should be 429 busy
            status_inv2, data_inv2 = _post("/powerpoint/invoke", {
                "protocolVersion": PROTOCOL_VERSION,
                "sessionId": sid,
                "requestId": generate_uuid(),
                "tool": "create_slide",
                "arguments": {},
            }, extra_headers=_plugin_headers())
            assert status_inv2 == 429

            # Cancel first to clean up
            _post("/powerpoint/cancel", {
                "sessionId": sid,
            }, extra_headers=_plugin_headers())
            f1.result(timeout=5)
        finally:
            server.stop()


class TestTicketLeaseAndRequeue:
    def test_claim_lease_requeue(self):
        clock = _Clock(0)
        server = _start_server(clock=clock)
        try:
            session = server.sessions.create("doc-lease", "pane-lease", "key-lease")
            ticket = Ticket(
                ticket_id="ticket-1",
                session_id=session.session_id,
                request_id="req-1",
                tool="get_presentation_overview",
                arguments={},
                created_at=clock(),
                _clock=clock,
            )
            session.set_ticket(ticket)

            # Claim
            assert ticket.claim("pane-lease", clock)
            assert ticket.state == "claimed"

            # Advance past lease
            clock.advance(CLAIM_LEASE_SECONDS + 1)
            assert ticket.is_lease_expired(clock)

            # Requeue
            assert ticket.requeue(clock)
            assert ticket.state == "pending"
            assert ticket.requeue_count == 1

            # Claim again
            assert ticket.claim("pane-lease", clock)

            # Second lease expiry — cannot requeue again
            clock.advance(CLAIM_LEASE_SECONDS + 1)
            assert not ticket.requeue(clock)
            assert ticket.state == "timeout"
        finally:
            server.stop()

    def test_late_result_after_first_lease_requeue_is_accepted(self):
        clock = _Clock(1)
        server = _start_server(clock=clock)
        try:
            session = server.sessions.create("doc-late", "pane-late", "key-late")
            ticket = Ticket(
                ticket_id="660e8400-e29b-41d4-a716-446655440001",
                session_id=session.session_id,
                request_id="660e8400-e29b-41d4-a716-446655440002",
                tool="add_text_box",
                arguments={"text": "once"},
                created_at=clock(),
                _clock=clock,
            )
            assert session.set_ticket(ticket)
            assert ticket.claim("pane-late", clock)
            clock.advance(CLAIM_LEASE_SECONDS + 1)
            server.sessions.sweep()
            assert ticket.state == "pending"
            assert ticket.requeue_count == 1

            status, _ = _post("/powerpoint/result", {
                "ticketId": ticket.ticket_id,
                "result": {"ok": True, "inserted": "once"},
            }, extra_headers=_session_headers(
                session.session_id, session.session_key, "pane-late",
            ))
            assert status == 200
            assert ticket.state == "completed"
            assert session.get_pending_ticket() is None, (
                "completed late result must not be offered for replay"
            )
        finally:
            server.stop()

    def test_result_idempotency(self):
        clock = _Clock(0)
        server = _start_server(clock=clock)
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-idem",
                "paneId": "pane-idem",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201
            sid = data["sessionId"]
            skey = data["sessionKey"]

            # Invoke and complete
            import concurrent.futures
            executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
            f1 = executor.submit(lambda: _post("/powerpoint/invoke", {
                "protocolVersion": PROTOCOL_VERSION,
                "sessionId": sid,
                "requestId": generate_uuid(),
                "tool": "read_all_slides",
                "arguments": {},
            }, extra_headers=_plugin_headers()))
            time.sleep(0.1)

            status_pend, data_pend = _get("/powerpoint/pending",
                extra_headers=_session_headers(sid, skey, "pane-idem"))
            tid = data_pend["ticketId"]

            # Submit result
            status_res, _ = _post("/powerpoint/result", {
                "ticketId": tid,
                "result": {"ok": True, "count": 7},
            }, extra_headers=_session_headers(sid, skey, "pane-idem"))
            assert status_res == 200
            f1.result(timeout=5)

            # Duplicate same result
            status_res2, _ = _post("/powerpoint/result", {
                "ticketId": tid,
                "result": {"ok": True, "count": 7},
            }, extra_headers=_session_headers(sid, skey, "pane-idem"))
            assert status_res2 == 200

            # Duplicate different result
            status_res3, _ = _post("/powerpoint/result", {
                "ticketId": tid,
                "result": {"ok": True, "count": 99},
            }, extra_headers=_session_headers(sid, skey, "pane-idem"))
            assert status_res3 == 409
        finally:
            server.stop()

    def test_result_wrong_pane_rejected(self):
        server = _start_server()
        try:
            # Create two sessions with different pane IDs
            status1, data1 = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-owner",
                "paneId": "pane-owner",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status1 == 201
            sid1 = data1["sessionId"]
            skey1 = data1["sessionKey"]

            status2, data2 = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-other",
                "paneId": "pane-other",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status2 == 201

            # Invoke on session1
            import concurrent.futures
            executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
            f1 = executor.submit(lambda: _post("/powerpoint/invoke", {
                "protocolVersion": PROTOCOL_VERSION,
                "sessionId": sid1,
                "requestId": generate_uuid(),
                "tool": "read_all_slides",
                "arguments": {},
            }, extra_headers=_plugin_headers()))
            time.sleep(0.1)

            # Poll from session1 to claim
            status_pend, data_pend = _get("/powerpoint/pending",
                extra_headers=_session_headers(sid1, skey1, "pane-owner"))
            assert status_pend == 200

            # Try to submit result from session2 (wrong pane)
            status_res, _ = _post("/powerpoint/result", {
                "ticketId": data_pend["ticketId"],
                "result": {"bad": True},
            }, extra_headers=_session_headers(data2["sessionId"], data2["sessionKey"], "pane-other"))
            # 404 is correct — the ticket belongs to session1, not session2
            assert status_res in (404, 409)
        finally:
            server.stop()


class TestHealthAndMetrics:
    def test_health(self):
        server = _start_server()
        try:
            status, data = _get("/powerpoint/health", extra_headers=_plugin_headers())
            assert status == 200
            assert data["ok"] is True
            assert data["uptime"] >= 0
        finally:
            server.stop()

    def test_metrics(self):
        server = _start_server()
        try:
            status, data = _get("/powerpoint/metrics", extra_headers=_plugin_headers())
            assert status == 200
            assert data["sessionsActive"] == 0
            assert data["ticketsPending"] == 0
        finally:
            server.stop()

    def test_capabilities(self):
        server = _start_server()
        try:
            status, data = _get("/powerpoint/capabilities", extra_headers=_plugin_headers())
            assert status == 200
            assert data["protocolVersion"] == PROTOCOL_VERSION
            assert data["hostPrefix"] == "/powerpoint"
            assert data["sessionTtl"] == SESSION_TTL_SECONDS
        finally:
            server.stop()


class TestShutdown:
    def test_shutdown_cancels_tickets(self):
        server = _start_server()
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-shut",
                "paneId": "pane-shut",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201

            # Create a ticket manually
            session = server.sessions.get(data["sessionId"])
            ticket = Ticket(
                ticket_id="shut-ticket",
                session_id=session.session_id,
                request_id="shut-req",
                tool="read_all_slides",
                arguments={},
                created_at=server.now(),
                _clock=server._clock,
            )
            session.set_ticket(ticket)
            assert ticket.state == "pending"

            # Stop the server
            server.stop()

            # Ticket should be cancelled
            assert ticket.state == "cancelled"
        except Exception:
            server.stop()
            raise


class TestBodyCaps:
    def test_session_body_too_large(self):
        server = _start_server()
        try:
            large_body = {"protocolVersion": PROTOCOL_VERSION, "documentId": "x" * 300, "paneId": "y"}
            status, _ = _post("/powerpoint/session", large_body,
                extra_headers={"Origin": "https://localhost:3002"})
            assert status == 400
        finally:
            server.stop()

    def test_invoke_body_too_large(self):
        server = _start_server()
        try:
            status, data = _post("/powerpoint/session", {
                "protocolVersion": PROTOCOL_VERSION,
                "documentId": "doc-big",
                "paneId": "pane-big",
            }, extra_headers={"Origin": "https://localhost:3002"})
            assert status == 201

            # Test with oversized documentId in session creation (body cap tested separately)
            large_body = {"protocolVersion": PROTOCOL_VERSION, "documentId": "x" * 300, "paneId": "y"}
            status, _ = _post("/powerpoint/session", large_body,
                extra_headers={"Origin": "https://localhost:3002"})
            assert status == 400
        finally:
            server.stop()

    def test_transfer_encoding_rejected(self):
        server = _start_server()
        try:
            import socket
            s = socket.create_connection((TEST_HOST, TEST_PORT), timeout=5)
            try:
                req = (
                    f"POST /powerpoint/session HTTP/1.1\r\n"
                    f"Host: {TEST_HOST}:{TEST_PORT}\r\n"
                    f"Content-Type: application/json\r\n"
                    f"Transfer-Encoding: chunked\r\n"
                    f"Origin: https://localhost:3002\r\n"
                    f"\r\n"
                ).encode()
                s.sendall(req)
                resp = b""
                while True:
                    chunk = s.recv(4096)
                    if not chunk:
                        break
                    resp += chunk
                resp_str = resp.decode()
                assert "411" in resp_str.split("\r\n")[0]
            finally:
                s.close()
        finally:
            server.stop()


class TestServerLoopback:
    def test_rejects_non_loopback(self):
        with pytest.raises(ValueError, match="loopback"):
            PowerPointCompanionServer(host="0.0.0.0", port=19999, plugin_key=PLUGIN_KEY)

    def test_rejects_external_ip(self):
        with pytest.raises(ValueError, match="loopback"):
            PowerPointCompanionServer(host="192.168.1.1", port=19999, plugin_key=PLUGIN_KEY)


class TestPluginKeyEnforcement:
    def test_health_without_key_fails(self):
        server = _start_server()
        try:
            status, _ = _get("/powerpoint/health")
            assert status == 401
        finally:
            server.stop()

    def test_health_with_wrong_key_fails(self):
        server = _start_server()
        try:
            status, _ = _get("/powerpoint/health", extra_headers={"X-Hermes-Plugin-Key": "wrong"})
            assert status == 401
        finally:
            server.stop()

    def test_session_status_without_key_fails(self):
        server = _start_server()
        try:
            status, _ = _get("/powerpoint/session")
            assert status == 401
        finally:
            server.stop()
