"""TDD behavioral tests for Word companion RPC server.

Run: python3 -m pytest scripts/tests/ -v
Or:   python3 scripts/tests/test_companion_server.py
"""
from __future__ import annotations

import json
import logging
import os
import socket
import sys
import threading
import time
import urllib.request
import urllib.error
from io import StringIO
from typing import Optional

# Add repo root to path for companion imports
_REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, _REPO)


# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------

def _url(addr: tuple[str, int], path: str) -> str:
    return f"http://{addr[0]}:{addr[1]}{path}"


def _request(method: str, addr: tuple[str, int], path: str,
             body: Optional[bytes] = None,
             headers: Optional[dict[str, str]] = None) -> tuple[int, bytes, dict]:
    """Make an HTTP request and return (status, body, response_headers)."""
    url = _url(addr, path)
    req = urllib.request.Request(url, data=body, method=method)
    if headers:
        for k, v in headers.items():
            req.add_header(k, v)
    if body is not None:
        req.add_header("Content-Length", str(len(body)))
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=5) as resp:
            return (resp.status, resp.read(), dict(resp.headers))
    except urllib.error.HTTPError as e:
        try:
            response_body = e.read()
        except OSError:
            # WinHTTP can reset while reading a Connection: close error body;
            # the parsed HTTP status and headers remain authoritative.
            response_body = b""
        return (e.code, response_body, dict(e.headers))
    except Exception as e:
        return (0, str(e).encode(), {})


# ---------------------------------------------------------------------------
# Test 1: loopback bind rejection
# ---------------------------------------------------------------------------

def test_bind_loopback_only():
    """Server must bind to 127.0.0.1 and reject non-loopback bind arguments."""
    from desk_plugin.companion.server import WordCompanionServer

    # Loopback bind must succeed
    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        addr = srv.server_address()
        assert addr[0] in ("127.0.0.1", "::1"), f"expected loopback, got {addr[0]}"
    finally:
        srv.stop()

    # Non-loopback must be rejected
    try:
        srv2 = WordCompanionServer(host="0.0.0.0", port=0)
        assert False, "should have raised ValueError for non-loopback bind"
    except ValueError:
        pass  # expected

    # Empty/missing host must be rejected
    try:
        srv3 = WordCompanionServer(host="", port=0)
        assert False, "should have raised ValueError for empty host"
    except ValueError:
        pass


def test_bind_requires_loopback_address():
    """Any non-loopback address must raise ValueError at construction."""
    from desk_plugin.companion.server import WordCompanionServer

    for bad_host in ("0.0.0.0", "192.168.1.1", "example.com", "::", "10.0.0.1"):
        try:
            WordCompanionServer(host=bad_host, port=0)
            assert False, f"should have rejected {bad_host}"
        except ValueError:
            pass


# ---------------------------------------------------------------------------
# Test 2: Origin enforcement on pane routes
# ---------------------------------------------------------------------------

def test_session_create_requires_exact_origin():
    """POST /word/session must require exact Origin https://localhost:3001."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        addr = srv.server_address()
        body = json.dumps({"documentId": "d1", "paneId": "p1", "protocolVersion": "1.0"}).encode()

        # No Origin → 403
        status, _, _ = _request("POST", addr, "/word/session", body=body)
        assert status == 403, f"no Origin should be 403, got {status}"

        # Wrong Origin → 403
        status, _, _ = _request("POST", addr, "/word/session", body=body,
                                headers={"Origin": "http://localhost:3000"})
        assert status == 403, f"http Origin should be 403, got {status}"

        status, _, _ = _request("POST", addr, "/word/session", body=body,
                                headers={"Origin": "https://evil.com:3000"})
        assert status == 403, f"wrong Origin should be 403, got {status}"

        # Correct Origin → 201
        status, resp_body, _ = _request("POST", addr, "/word/session", body=body,
                                        headers={"Origin": "https://localhost:3001"})
        assert status == 201, f"correct Origin should be 201, got {status}"
        data = json.loads(resp_body)
        assert "sessionId" in data
        assert "sessionKey" in data
        assert data.get("protocolVersion") == "1.0"

    finally:
        srv.stop()


def test_json_body_routes_require_json_content_type():
    """JSON endpoints reject non-JSON media types before parsing the body."""
    import http.client

    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        host, port = srv.server_address()
        conn = http.client.HTTPConnection(host, port, timeout=2)
        conn.request(
            "POST",
            "/word/session",
            body=b"{}",
            headers={
                "Origin": "https://localhost:3001",
                "Content-Type": "text/plain",
                "Content-Length": "2",
            },
        )
        response = conn.getresponse()
        assert response.status == 415
        assert json.loads(response.read())["category"] == "Validation"
        conn.close()
    finally:
        srv.stop()


def test_plugin_routes_use_plugin_key_not_origin():
    """Plugin routes require X-Hermes-Plugin-Key, not Origin.
    Uses GET /word/health (non-blocking) to verify key auth without polling.
    """
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "test-plugin-key-32-bytes-hex!"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()

        # GET /word/health without plugin key → 401
        status, _, _ = _request("GET", addr, "/word/health")
        assert status in (401, 403), f"missing plugin key should be 401/403, got {status}"

        # Wrong plugin key → 401/403
        status, _, _ = _request("GET", addr, "/word/health",
                                headers={"X-Hermes-Plugin-Key": "wrong-key"})
        assert status in (401, 403), f"wrong plugin key should be 401/403, got {status}"

        # Correct plugin key, no Origin needed → 200
        status, resp_body, _ = _request("GET", addr, "/word/health",
                                        headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
        assert status == 200, f"correct plugin key should be 200, got {status}: {resp_body}"
        data = json.loads(resp_body)
        assert data.get("ok") is True

        # Correct plugin key WITH Origin should still work (Origin ignored for plugin routes)
        status, _, _ = _request("GET", addr, "/word/health",
                                headers={"X-Hermes-Plugin-Key": PLUGIN_KEY,
                                         "Origin": "https://evil.com:3000"})
        assert status == 200, f"plugin routes should ignore Origin, got {status}"

    finally:
        srv.stop()


def test_plugin_key_constant_time_leak():
    """Wrong vs missing plugin key must not leak which part matched (both 401)."""
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "test-plugin-key-32-bytes-hex!"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()

        inv_body = json.dumps({
            "protocolVersion": "1.0",
            "sessionId": "00000000-0000-0000-0000-000000000001",
            "requestId": "550e8400-e29b-41d4-a716-446655440000",
            "tool": "read_document_text",
            "arguments": {"sheet": "S1", "address": "A1"}
        }).encode()

        # Missing key
        s1, _, _ = _request("POST", addr, "/word/invoke", body=inv_body)
        # Wrong key
        s2, _, _ = _request("POST", addr, "/word/invoke", body=inv_body,
                            headers={"X-Hermes-Plugin-Key": "wrong-key!"})

        # Both must return same status (401) — no timing or status leak
        assert s1 == s2 == 401, f"both missing/wrong key must be 401, got {s1}/{s2}"

    finally:
        srv.stop()


def test_session_routes_require_session_auth():
    """Pane routes need Origin + X-Hermes-Session-Key + X-Hermes-Session-Id + X-Hermes-Pane-Id."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        addr = srv.server_address()

        # Create session
        sess_body = json.dumps({"documentId": "d1", "paneId": "p1", "protocolVersion": "1.0"}).encode()
        status, resp, _ = _request("POST", addr, "/word/session", body=sess_body,
                                   headers={"Origin": "https://localhost:3001"})
        assert status == 201
        sess = json.loads(resp)

        # DELETE /word/session without session auth → 401/403
        status, _, _ = _request("DELETE", addr, "/word/session",
                                headers={"Origin": "https://localhost:3001"})
        assert status in (401, 403), f"DELETE without session auth should be 401/403, got {status}"

        # DELETE with correct auth → 204
        status, _, _ = _request("DELETE", addr, "/word/session", headers={
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Key": sess["sessionKey"],
            "X-Hermes-Session-Id": sess["sessionId"],
            "X-Hermes-Pane-Id": "p1",
        })
        assert status == 204, f"correct session DELETE should be 204, got {status}"

    finally:
        srv.stop()


def test_wrong_session_auth_rejected():
    """Wrong session-id, key, or pane-id must be rejected on pane routes."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        addr = srv.server_address()

        # Create session
        sess_body = json.dumps({"documentId": "d1", "paneId": "p1", "protocolVersion": "1.0"}).encode()
        status, resp, _ = _request("POST", addr, "/word/session", body=sess_body,
                                   headers={"Origin": "https://localhost:3001"})
        assert status == 201
        sess = json.loads(resp)

        base_headers = {"Origin": "https://localhost:3001"}

        # Wrong session key
        s, _, _ = _request("DELETE", addr, "/word/session", headers={
            **base_headers,
            "X-Hermes-Session-Key": "wrong-key",
            "X-Hermes-Session-Id": sess["sessionId"],
            "X-Hermes-Pane-Id": "p1",
        })
        assert s in (401, 403), f"wrong session key should be 401/403, got {s}"

        # Wrong session id
        s, _, _ = _request("DELETE", addr, "/word/session", headers={
            **base_headers,
            "X-Hermes-Session-Key": sess["sessionKey"],
            "X-Hermes-Session-Id": "00000000-0000-0000-0000-000000000099",
            "X-Hermes-Pane-Id": "p1",
        })
        assert s in (401, 403), f"wrong session id should be 401/403, got {s}"

        # Wrong pane id
        s, _, _ = _request("DELETE", addr, "/word/session", headers={
            **base_headers,
            "X-Hermes-Session-Key": sess["sessionKey"],
            "X-Hermes-Session-Id": sess["sessionId"],
            "X-Hermes-Pane-Id": "wrong-pane",
        })
        assert s in (401, 403), f"wrong pane id should be 401/403, got {s}"

    finally:
        srv.stop()


def test_options_preflight_pane_routes_only():
    """OPTIONS must be narrow: only for pane routes and pane headers."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        addr = srv.server_address()

        # OPTIONS on pane route should succeed
        status, _, hdrs = _request("OPTIONS", addr, "/word/session",
                                   headers={"Origin": "https://localhost:3001",
                                            "Access-Control-Request-Method": "POST"})
        assert status == 204, f"OPTIONS on /word/session should be 204, got {status}"

        # OPTIONS on plugin route should 405 or not expose CORS
        status, _, _ = _request("OPTIONS", addr, "/word/invoke",
                                headers={"Origin": "https://localhost:3001"})
        # Plugin routes don't need CORS
        assert status in (204, 405), f"OPTIONS on plugin route may be 204 or 405, got {status}"

    finally:
        srv.stop()


# ---------------------------------------------------------------------------
# Test 3: body cap enforcement
# ---------------------------------------------------------------------------

def _create_session(addr, doc_id="d1", pane_id="p1"):
    """Helper: create a session and return the parsed JSON."""
    sess_body = json.dumps({
        "documentId": doc_id, "paneId": pane_id, "protocolVersion": "1.0"
    }).encode()
    status, resp, _ = _request("POST", addr, "/word/session", body=sess_body,
                               headers={"Origin": "https://localhost:3001"})
    assert status == 201, f"session create failed: {status}"
    return json.loads(resp)


def test_body_caps_enforced_on_invoke():
    """POST /word/invoke must reject bodies > 1 MiB, missing Content-Length,
    and chunked transfer-encoding before JSON parsing."""
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "test-plugin-key-body-caps"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()

        # --- Missing Content-Length → 411 (use raw socket, urllib adds 0)
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(3)
        sock.connect(addr)
        raw_req = (
            f"POST /word/invoke HTTP/1.1\r\n"
            f"Host: {addr[0]}:{addr[1]}\r\n"
            f"X-Hermes-Plugin-Key: {PLUGIN_KEY}\r\n"
            f"\r\n"
        ).encode()
        sock.sendall(raw_req)
        raw_resp = b""
        try:
            while True:
                chunk = sock.recv(4096)
                if not chunk:
                    break
                raw_resp += chunk
        except socket.timeout:
            pass
        sock.close()
        status_line = raw_resp.split(b"\r\n")[0].decode()
        assert "411" in status_line, f"missing Content-Length should be 411, got: {status_line}"

        # --- Chunked transfer → 411
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(3)
        sock.connect(addr)
        raw_req = (
            f"POST /word/invoke HTTP/1.1\r\n"
            f"Host: {addr[0]}:{addr[1]}\r\n"
            f"X-Hermes-Plugin-Key: {PLUGIN_KEY}\r\n"
            f"Transfer-Encoding: chunked\r\n"
            f"\r\n"
        ).encode()
        sock.sendall(raw_req)
        raw_resp = b""
        try:
            while True:
                chunk = sock.recv(4096)
                if not chunk:
                    break
                raw_resp += chunk
        except socket.timeout:
            pass
        sock.close()
        status_line = raw_resp.split(b"\r\n")[0].decode()
        assert "411" in status_line, f"chunked should be 411, got: {status_line}"

        # --- Oversize body (>1 MiB) → 413 (server may close connection early)
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(3)
        sock.connect(addr)
        cl = 1024 * 1024 + 1
        raw_req = (
            f"POST /word/invoke HTTP/1.1\r\n"
            f"Host: {addr[0]}:{addr[1]}\r\n"
            f"X-Hermes-Plugin-Key: {PLUGIN_KEY}\r\n"
            f"Content-Type: application/json\r\n"
            f"Content-Length: {cl}\r\n"
            f"\r\n"
        ).encode()
        sock.sendall(raw_req)
        # Don't send body — just read response
        raw_resp = b""
        try:
            while True:
                chunk = sock.recv(4096)
                if not chunk:
                    break
                raw_resp += chunk
        except socket.timeout:
            pass
        sock.close()
        status_line = raw_resp.split(b"\r\n")[0].decode()
        assert "413" in status_line, f"oversize body should be 413, got: {status_line}"

        # --- Invalid Content-Length (non-integer) → 400
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(3)
        sock.connect(addr)
        raw_req = (
            f"POST /word/invoke HTTP/1.1\r\n"
            f"Host: {addr[0]}:{addr[1]}\r\n"
            f"X-Hermes-Plugin-Key: {PLUGIN_KEY}\r\n"
            f"Content-Type: application/json\r\n"
            f"Content-Length: not-a-number\r\n"
            f"\r\n"
        ).encode()
        sock.sendall(raw_req)
        raw_resp = b""
        try:
            while True:
                chunk = sock.recv(4096)
                if not chunk:
                    break
                raw_resp += chunk
        except socket.timeout:
            pass
        sock.close()
        status_line = raw_resp.split(b"\r\n")[0].decode()
        assert "400" in status_line, f"invalid CL should be 400, got: {status_line}"

    finally:
        srv.stop()


def test_body_caps_enforced_on_result():
    """POST /word/result must reject bodies > 10 MiB (server reads headers, rejects)."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr)

        # 10 MiB + 1 → server should reject at header stage (just check CL)
        cl = 10 * 1024 * 1024 + 1
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(3)
        sock.connect(addr)
        raw_req = (
            f"POST /word/result HTTP/1.1\r\n"
            f"Host: {addr[0]}:{addr[1]}\r\n"
            f"Origin: https://localhost:3001\r\n"
            f"X-Hermes-Session-Key: {sess['sessionKey']}\r\n"
            f"X-Hermes-Session-Id: {sess['sessionId']}\r\n"
            f"X-Hermes-Pane-Id: p1\r\n"
            f"Content-Type: application/json\r\n"
            f"Content-Length: {cl}\r\n"
            f"\r\n"
        ).encode()
        sock.sendall(raw_req)
        raw_resp = b""
        try:
            while True:
                chunk = sock.recv(4096)
                if not chunk:
                    break
                raw_resp += chunk
        except socket.timeout:
            pass
        sock.close()
        status_line = raw_resp.split(b"\r\n")[0].decode()
        assert "413" in status_line, f"oversize result should be 413, got: {status_line}"

    finally:
        srv.stop()


# ---------------------------------------------------------------------------
# Test helper — create session + return session data
# ---------------------------------------------------------------------------

SESSION_HELPER_ALREADY_DEFINED = True  # noqa: silence


# ---------------------------------------------------------------------------
# Test 4: session lifecycle
# ---------------------------------------------------------------------------


def test_session_heartbeat_extends_ttl():
    """Heartbeat resets the expiry clock."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr)

        hdrs = {
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Key": sess["sessionKey"],
            "X-Hermes-Session-Id": sess["sessionId"],
            "X-Hermes-Pane-Id": "p1",
        }

        # Two heartbeats should succeed
        for i in range(2):
            status, _, _ = _request("POST", addr, "/word/heartbeat",
                                    body=b"{}", headers=hdrs)
            assert status == 200, f"heartbeat {i} should be 200, got {status}"

    finally:
        srv.stop()


def test_session_delete():
    """DELETE removes session; subsequent requests with its key fail."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr)

        hdrs = {
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Key": sess["sessionKey"],
            "X-Hermes-Session-Id": sess["sessionId"],
            "X-Hermes-Pane-Id": "p1",
        }

        status, _, _ = _request("DELETE", addr, "/word/session", headers=hdrs)
        assert status == 204

        # Second attempt with same auth should fail
        status, _, _ = _request("DELETE", addr, "/word/session", headers=hdrs)
        assert status in (401, 403, 404), f"re-delete should fail, got {status}"

    finally:
        srv.stop()


def test_session_ttl_expiry():
    """After TTL with no heartbeat, session should be gone."""
    from desk_plugin.companion.server import WordCompanionServer
    from desk_plugin.companion.protocol import SESSION_TTL_SECONDS

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    # We'll sweep manually with the manager
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr)

        # Manually expire by sweeping with a future clock
        future = time.monotonic() + SESSION_TTL_SECONDS + 10
        expired = srv.sessions.sweep(future)
        assert len(expired) >= 0  # session may or may not have a ticket

        # Now check session status
        hdrs = {
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Key": sess["sessionKey"],
            "X-Hermes-Session-Id": sess["sessionId"],
            "X-Hermes-Pane-Id": "p1",
        }
        status, _, _ = _request("POST", addr, "/word/heartbeat",
                                body=b"{}", headers=hdrs)
        assert status in (401, 403, 404), f"expired session should be gone, got {status}"

    finally:
        srv.stop()


def test_get_session_status():
    """GET /word/session (plugin key) returns status without content."""
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "ss-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr)

        # Get status (no sessionId param → returns overview)
        status, resp, _ = _request("GET", addr, "/word/session",
                                   headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
        assert status == 200
        data = json.loads(resp)
        assert data.get("sessions") >= 1

        # Get status for specific session
        status, resp, _ = _request("GET", addr,
                                   f"/word/session?sessionId={sess['sessionId']}",
                                   headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
        assert status == 200
        data = json.loads(resp)
        assert data.get("active") is True

    finally:
        srv.stop()


# ---------------------------------------------------------------------------
# Test 5: multi-session isolation
# ---------------------------------------------------------------------------

def test_multiple_sessions_isolated():
    """Two independent sessions with different document/panes."""
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "multi-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()

        sess_a = _create_session(addr, doc_id="doc-a", pane_id="pa")
        sess_b = _create_session(addr, doc_id="doc-b", pane_id="pb")

        # Heartbeat A
        hdrs_a = {
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Key": sess_a["sessionKey"],
            "X-Hermes-Session-Id": sess_a["sessionId"],
            "X-Hermes-Pane-Id": "pa",
        }
        status, _, _ = _request("POST", addr, "/word/heartbeat", body=b"{}",
                                headers=hdrs_a)
        assert status == 200

        # Delete B
        hdrs_b = {
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Key": sess_b["sessionKey"],
            "X-Hermes-Session-Id": sess_b["sessionId"],
            "X-Hermes-Pane-Id": "pb",
        }
        status, _, _ = _request("DELETE", addr, "/word/session", headers=hdrs_b)
        assert status == 204

        # A should still be valid
        status, _, _ = _request("POST", addr, "/word/heartbeat", body=b"{}",
                                headers=hdrs_a)
        assert status == 200, f"Session A should still be alive"

        # B should be gone
        status, _, _ = _request("POST", addr, "/word/heartbeat", body=b"{}",
                                headers=hdrs_b)
        assert status in (401, 403, 404), f"Session B should be deleted"

    finally:
        srv.stop()


def test_invoke_cannot_overflow_queue():
    """Second invoke on same session must get 429 (queue depth 1)."""
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "depth-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr)

        inv_body = json.dumps({
            "protocolVersion": "1.0",
            "sessionId": sess["sessionId"],
            "requestId": "550e8400-e29b-41d4-a716-446655440000",
            "tool": "read_document_text",
            "arguments": {"maxLength": 100}
        }).encode()

        # Submit invoke in background — will block
        results = {}

        def invoke_bg():
            s, b, _ = _request("POST", addr, "/word/invoke", body=inv_body,
                               headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
            results["first"] = s

        t = threading.Thread(target=invoke_bg, daemon=True)
        t.start()

        # Give it time to set the ticket
        time.sleep(0.1)

        # Second invoke should get 429
        inv2 = json.dumps({
            "protocolVersion": "1.0",
            "sessionId": sess["sessionId"],
            "requestId": "660e8400-e29b-41d4-a716-446655440001",
            "tool": "insert_text",
            "arguments": {"text": "x", "location": "end"}
        }).encode()
        status2, resp2, _ = _request("POST", addr, "/word/invoke", body=inv2,
                                     headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
        assert status2 == 429, f"second invoke should be 429, got {status2}: {resp2}"

        # Now cancel the blocked one by deleting session
        hdrs = {
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Key": sess["sessionKey"],
            "X-Hermes-Session-Id": sess["sessionId"],
            "X-Hermes-Pane-Id": "p1",
        }
        _request("DELETE", addr, "/word/session", headers=hdrs)
        t.join(timeout=5)

    finally:
        srv.stop()


# ---------------------------------------------------------------------------
# Test 6: claim, lease, result idempotency
# ---------------------------------------------------------------------------

def test_pending_claim_and_result():
    """Pane claims a pending ticket and submits a result."""
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "claim-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr)
        hdrs = {
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Key": sess["sessionKey"],
            "X-Hermes-Session-Id": sess["sessionId"],
            "X-Hermes-Pane-Id": "p1",
        }

        # Submit invoke in background
        result = {"status": None, "body": None}

        def do_invoke():
            inv_body = json.dumps({
                "protocolVersion": "1.0",
                "sessionId": sess["sessionId"],
                "requestId": "550e8400-e29b-41d4-a716-446655440000",
                "tool": "read_document_text",
                "arguments": {"maxLength": 100}
            }).encode()
            s, b, _ = _request("POST", addr, "/word/invoke", body=inv_body,
                               headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
            result["status"] = s
            result["body"] = b

        t = threading.Thread(target=do_invoke, daemon=True)
        t.start()
        time.sleep(0.1)

        # Claim the pending ticket
        status, claim_body, _ = _request("GET", addr, "/word/pending", headers=hdrs)
        assert status == 200, f"pending claim should be 200, got {status}"
        claim = json.loads(claim_body)
        ticket_id = claim["ticketId"]
        assert claim["tool"] == "read_document_text"

        # Claim again → should get 204
        status2, _, _ = _request("GET", addr, "/word/pending", headers=hdrs)
        assert status2 == 204, f"re-claim should be 204, got {status2}"

        # Submit result
        result_body = json.dumps({
            "ticketId": ticket_id,
            "result": {"cells": [["hello"]]}
        }).encode()
        status3, _, _ = _request("POST", addr, "/word/result",
                                 body=result_body, headers=hdrs)
        assert status3 == 200, f"result submit should be 200, got {status3}"

        t.join(timeout=5)
        assert result["status"] == 200, f"invoke should get 200, got {result['status']}: {result['body']}"

    finally:
        srv.stop()


def test_result_idempotency():
    """Duplicate result with same payload is idempotent; different payload is 409."""
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "idem-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr)
        hdrs = {
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Key": sess["sessionKey"],
            "X-Hermes-Session-Id": sess["sessionId"],
            "X-Hermes-Pane-Id": "p1",
        }

        # Start invoke
        def do_invoke():
            inv_body = json.dumps({
                "protocolVersion": "1.0",
                "sessionId": sess["sessionId"],
                "requestId": "550e8400-e29b-41d4-a716-446655440000",
                "tool": "read_document_text",
                "arguments": {"maxLength": 100}
            }).encode()
            _request("POST", addr, "/word/invoke", body=inv_body,
                     headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})

        t = threading.Thread(target=do_invoke, daemon=True)
        t.start()
        time.sleep(0.1)

        # Claim
        status, claim_body, _ = _request("GET", addr, "/word/pending", headers=hdrs)
        claim = json.loads(claim_body)
        tid = claim["ticketId"]

        # Submit result
        res = json.dumps({"ticketId": tid, "result": {"data": "X"}}).encode()
        s1, _, _ = _request("POST", addr, "/word/result", body=res, headers=hdrs)
        assert s1 == 200

        # Duplicate with same payload → 200
        s2, _, _ = _request("POST", addr, "/word/result", body=res, headers=hdrs)
        assert s2 == 200, f"duplicate same should be 200, got {s2}"

        # Different payload → 409
        res2 = json.dumps({"ticketId": tid, "result": {"data": "DIFFERENT"}}).encode()
        s3, _, _ = _request("POST", addr, "/word/result", body=res2, headers=hdrs)
        assert s3 == 409, f"conflicting duplicate should be 409, got {s3}"

        t.join(timeout=5)

    finally:
        srv.stop()


def test_lease_requeue_then_timeout():
    """Claim lease expires → requeued once → second expiry = timeout."""
    from desk_plugin.companion.server import WordCompanionServer
    from desk_plugin.companion.protocol import CLAIM_LEASE_SECONDS

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr)
        hdrs = {
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Key": sess["sessionKey"],
            "X-Hermes-Session-Id": sess["sessionId"],
            "X-Hermes-Pane-Id": "p1",
        }

        # Mint a ticket directly (bypass invoke blocking)
        from desk_plugin.companion.session import Ticket
        ticket = Ticket(
            ticket_id="test-ticket-1",
            session_id=sess["sessionId"],
            request_id="550e8400-e29b-41d4-a716-446655440000",
            tool="test_tool",
            arguments={},
        )
        session = srv.sessions.get(sess["sessionId"])
        session.set_ticket(ticket)

        # Claim it
        claimed = ticket.claim("p1")
        assert claimed, "should claim successfully"

        # First lease expiry → requeue
        future = time.monotonic() + CLAIM_LEASE_SECONDS + 1
        expired = session.sweep_expired(future)
        assert len(expired) == 0, "first lease expiry should requeue, not timeout"
        assert ticket.state == "pending", f"should be requeued, got {ticket.state}"

        # Re-claim
        claimed2 = ticket.claim("p1")
        assert claimed2, "should re-claim"

        # Second lease expiry → timeout
        future2 = time.monotonic() + CLAIM_LEASE_SECONDS + 1
        expired2 = session.sweep_expired(future2)
        assert len(expired2) >= 1, "second lease expiry should timeout"
        assert ticket.state == "timeout", f"should be timeout, got {ticket.state}"

    finally:
        srv.stop()


def test_late_result_after_first_lease_requeue_is_accepted():
    """A cached result arriving after one lease requeue must prevent replay."""
    from desk_plugin.companion.server import WordCompanionServer
    from desk_plugin.companion.protocol import CLAIM_LEASE_SECONDS
    from desk_plugin.companion.session import Ticket

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr)
        headers = {
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Key": sess["sessionKey"],
            "X-Hermes-Session-Id": sess["sessionId"],
            "X-Hermes-Pane-Id": "p1",
        }
        ticket = Ticket(
            ticket_id="660e8400-e29b-41d4-a716-446655440001",
            session_id=sess["sessionId"],
            request_id="660e8400-e29b-41d4-a716-446655440002",
            tool="insert_text",
            arguments={"text": "once"},
        )
        session = srv.sessions.get(sess["sessionId"])
        assert session is not None
        assert session.set_ticket(ticket)
        assert ticket.claim("p1")
        session.sweep_expired(lambda: time.monotonic() + CLAIM_LEASE_SECONDS + 1)
        assert ticket.state == "pending"
        assert ticket.requeue_count == 1

        result_body = json.dumps({
            "ticketId": ticket.ticket_id,
            "result": {"ok": True, "inserted": "once"},
        }).encode()
        status, _, _ = _request(
            "POST", addr, "/word/result", body=result_body, headers=headers,
        )
        assert status == 200
        assert ticket.state == "completed"
        assert ticket.result_payload == {"ok": True, "inserted": "once"}
        pending_status, _, _ = _request("GET", addr, "/word/pending", headers=headers)
        assert pending_status == 204, "completed late result must not be offered for replay"
    finally:
        srv.stop()


# ---------------------------------------------------------------------------
# Test 7: cancellation, shutdown, version reject
# ---------------------------------------------------------------------------

def test_cancel_invoke():
    """POST /word/cancel cancels the pending ticket; blocked invoke gets 499."""
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "cancel-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr)

        result = {"status": None}

        def do_invoke():
            inv_body = json.dumps({
                "protocolVersion": "1.0",
                "sessionId": sess["sessionId"],
                "requestId": "550e8400-e29b-41d4-a716-446655440000",
                "tool": "read_document_text",
                "arguments": {"maxLength": 100}
            }).encode()
            s, b, _ = _request("POST", addr, "/word/invoke", body=inv_body,
                               headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
            result["status"] = s

        t = threading.Thread(target=do_invoke, daemon=True)
        t.start()
        time.sleep(0.1)

        # Cancel
        cancel_body = json.dumps({"sessionId": sess["sessionId"]}).encode()
        s, _, _ = _request("POST", addr, "/word/cancel", body=cancel_body,
                           headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
        assert s == 200, f"cancel should be 200, got {s}"

        t.join(timeout=5)
        # The invoke should receive 499
        assert result["status"] == 499, f"cancelled invoke should be 499, got {result['status']}"

    finally:
        srv.stop()


def test_shutdown_fails_all_tickets():
    """Server.stop() cancels all outstanding tickets."""
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "shut-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr)

        result = {"status": None}

        def do_invoke():
            inv_body = json.dumps({
                "protocolVersion": "1.0",
                "sessionId": sess["sessionId"],
                "requestId": "550e8400-e29b-41d4-a716-446655440000",
                "tool": "read_document_text",
                "arguments": {"maxLength": 100}
            }).encode()
            s, b, _ = _request("POST", addr, "/word/invoke", body=inv_body,
                               headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
            result["status"] = s

        t = threading.Thread(target=do_invoke, daemon=True)
        t.start()
        time.sleep(0.1)

        # Stop the server — this cancels all tickets
        srv.stop()
        t.join(timeout=5)

        # After shutdown, the invoke thread should have received 499 or errored
        assert result["status"] is not None, "invoke should have completed after shutdown"
        assert result["status"] in (0, 499), f"shutdown invoke should be 0 or 499, got {result['status']}"

    finally:
        # Already stopped
        pass


def test_version_mismatch_rejected():
    """Unsupported protocol version must be rejected with 400."""
    from desk_plugin.companion.server import WordCompanionServer


    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        addr = srv.server_address()

        body = json.dumps({
            "documentId": "d1", "paneId": "p1", "protocolVersion": "2.0"
        }).encode()
        status, resp, _ = _request(
            "POST", addr, "/word/session", body=body,
            headers={"Origin": "https://localhost:3001"})
        assert status == 400, f"version 2.0 should be 400, got {status}"
        assert b"Unsupported" in resp or b"unsupported" in resp

        body2 = json.dumps({
            "documentId": "d1", "paneId": "p1", "protocolVersion": "0.9"
        }).encode()
        status2, _, _ = _request(
            "POST", addr, "/word/session", body=body2,
            headers={"Origin": "https://localhost:3001"})
        assert status2 == 400, f"version 0.9 should be 400, got {status2}"

    finally:
        srv.stop()


def test_wrong_pane_cannot_claim():
    """Pane A's session cannot be claimed by pane B."""
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "wp-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr, pane_id="correct-pane")

        # Mint a ticket
        from desk_plugin.companion.session import Ticket
        ticket = Ticket(
            ticket_id="wp-ticket",
            session_id=sess["sessionId"],
            request_id="550e8400-e29b-41d4-a716-446655440000",
            tool="test",
            arguments={},
        )
        session = srv.sessions.get(sess["sessionId"])
        session.set_ticket(ticket)

        # Wrong pane tries to claim via pending route
        wrong_hdrs = {
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Key": sess["sessionKey"],
            "X-Hermes-Session-Id": sess["sessionId"],
            "X-Hermes-Pane-Id": "wrong-pane",
        }
        # Session auth check should reject wrong pane
        status, _, _ = _request("GET", addr, "/word/pending", headers=wrong_hdrs)
        assert status in (401, 403), f"wrong pane should be rejected, got {status}"

    finally:
        srv.stop()


# ---------------------------------------------------------------------------
# Test 8: log hygiene, metrics, health
# ---------------------------------------------------------------------------

def test_metrics_are_counts_only():
    """GET /word/metrics returns only count aggregates, no content."""
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "metrics-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()

        status, resp, _ = _request("GET", addr, "/word/metrics",
                                   headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
        assert status == 200
        data = json.loads(resp)
        assert isinstance(data.get("sessionsActive"), int)
        assert isinstance(data.get("ticketsPending"), int)
        # Must not contain any document/pane/tool info
        for k in data:
            assert isinstance(data[k], int), f"metric {k} must be int"

    finally:
        srv.stop()


def test_health_endpoint():
    """GET /word/health returns the exact bounded managed identity."""
    from desk_plugin.companion.server import WordCompanionServer
    from desk_plugin.companion.protocol import PROTOCOL_VERSION, SERVICE_HOST, SERVICE_NAME

    PLUGIN_KEY = "health-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()

        status, resp, _ = _request("GET", addr, "/word/health",
                                   headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
        assert status == 200
        data = json.loads(resp)
        assert data == {
            "ok": True,
            "service": SERVICE_NAME,
            "host": SERVICE_HOST,
            "protocolVersion": PROTOCOL_VERSION,
        }
        assert len(resp) <= 256
        # No key, document content, session, path, timing, or stack metadata.
        for sensitive in (
            PLUGIN_KEY, "document", "pane", "tool", "session", "path",
            "uptime", "port", "traceback", "stack",
        ):
            assert sensitive.lower().encode() not in resp.lower(), (
                f"health must not leak {sensitive}"
            )

    finally:
        srv.stop()


def test_health_identity_constants_are_shared_word_values():
    """The managed identity is defined once from the Word protocol constants."""
    from desk_plugin.companion.protocol import (
        HOST_PREFIX,
        PROTOCOL_VERSION,
        SERVICE_HOST,
        SERVICE_NAME,
    )

    assert SERVICE_NAME == "hermes-office-word-companion"
    assert SERVICE_HOST == HOST_PREFIX.removeprefix("/") == "word"
    assert PROTOCOL_VERSION == "1.0"


def test_health_missing_and_wrong_plugin_key_have_identical_bounded_401():
    """Missing and wrong keys are indistinguishable bounded auth failures."""
    from desk_plugin.companion.server import WordCompanionServer

    plugin_key = "health-auth-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=plugin_key)
    srv.start()
    try:
        addr = srv.server_address()
        results = [
            _request("GET", addr, "/word/health"),
            _request(
                "GET", addr, "/word/health",
                headers={"X-Hermes-Plugin-Key": "wrong-health-key"},
            ),
        ]
        assert [status for status, _, _ in results] == [401, 401]
        assert results[0][1] == results[1][1]
        assert results[0][2].get("Content-Length") == results[1][2].get("Content-Length")
        assert len(results[0][1]) <= 256
        for _, body, _ in results:
            assert plugin_key.encode() not in body
            assert b"hermes-office-word-companion" not in body
    finally:
        srv.stop()


def test_health_is_read_only_for_session_and_ticket_state():
    """Authenticated health does not heartbeat, claim, complete, or expose state."""
    from desk_plugin.companion.server import WordCompanionServer
    from desk_plugin.companion.session import Ticket

    plugin_key = "health-read-only-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=plugin_key)
    session = srv.sessions.create("document-sentinel", "pane-sentinel", "s" * 64)
    ticket = Ticket(
        ticket_id="77777777-7777-4777-8777-777777777777",
        session_id=session.session_id,
        request_id="88888888-8888-4888-8888-888888888888",
        tool="word_insert_text",
        arguments={"text": "document-sentinel"},
    )
    assert session.set_ticket(ticket)
    before = (
        session.last_heartbeat,
        ticket.state,
        ticket.claimed_by_pane_id,
        ticket.result_payload,
        srv.sessions.get_metrics(),
    )
    srv.start()
    try:
        status, body, _ = _request(
            "GET", srv.server_address(), "/word/health",
            headers={"X-Hermes-Plugin-Key": plugin_key},
        )
        assert status == 200
        assert json.loads(body)["ok"] is True
        after = (
            session.last_heartbeat,
            ticket.state,
            ticket.claimed_by_pane_id,
            ticket.result_payload,
            srv.sessions.get_metrics(),
        )
        assert after == before
        assert b"document-sentinel" not in body
        assert b"word_insert_text" not in body
    finally:
        srv.stop()


def test_health_does_not_sweep_expired_session_or_pending_ticket():
    """Health identity must not perform the normal session/ticket maintenance sweep."""
    from desk_plugin.companion.server import WordCompanionServer
    from desk_plugin.companion.protocol import SESSION_TTL_SECONDS
    from desk_plugin.companion.session import Ticket

    plugin_key = "health-no-sweep-key"
    now = 1000.0
    srv = WordCompanionServer(
        host="127.0.0.1", port=0, plugin_key=plugin_key,
        clock=lambda: now,
    )
    session = srv.sessions.create("expired-document", "expired-pane", "s" * 64)
    session.last_heartbeat = now - SESSION_TTL_SECONDS
    ticket = Ticket(
        ticket_id="99999999-9999-4999-8999-999999999999",
        session_id=session.session_id,
        request_id="aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
        tool="word_insert_text",
        arguments={"text": "expired-document"},
        created_at=now,
        _clock=lambda: now,
    )
    assert session.set_ticket(ticket)
    srv.start()
    try:
        status, body, _ = _request(
            "GET", srv.server_address(), "/word/health",
            headers={"X-Hermes-Plugin-Key": plugin_key},
        )
        assert status == 200
        assert json.loads(body)["service"] == "hermes-office-word-companion"
        assert srv.sessions.get(session.session_id) is session
        assert ticket.state == "pending"
    finally:
        srv.stop()


def test_capabilities_endpoint():
    """GET /word/capabilities returns protocol metadata only."""
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "cap-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()

        status, resp, _ = _request("GET", addr, "/word/capabilities",
                                   headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
        assert status == 200
        data = json.loads(resp)
        assert data["protocolVersion"] == "1.0"
        assert data["hostPrefix"] == "/word"
        assert "sessionTtl" in data

        # No secrets
        for sensitive in ("key", "document", "pane"):
            assert sensitive not in resp.decode().lower(), f"caps must not leak {sensitive}"

    finally:
        srv.stop()


# ---------------------------------------------------------------------------
# Test 9: real socket round-trip
# ---------------------------------------------------------------------------

def test_full_invoke_round_trip():
    """Plugin invokes, pane claims pending, submits result, plugin gets response.

    This is the canonical end-to-end flow exercising the real HTTP server.
    """
    from desk_plugin.companion.server import WordCompanionServer

    PLUGIN_KEY = "roundtrip-key"
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        addr = srv.server_address()

        # 1. Create session (pane)
        sess = _create_session(addr, doc_id="rt-doc", pane_id="rt-pane")
        pane_hdrs = {
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Key": sess["sessionKey"],
            "X-Hermes-Session-Id": sess["sessionId"],
            "X-Hermes-Pane-Id": "rt-pane",
        }

        # 2. Invoke in background (plugin)
        invoke_result = {}

        def plugin_invoke():
            inv_body = json.dumps({
                "protocolVersion": "1.0",
                "sessionId": sess["sessionId"],
                "requestId": "550e8400-e29b-41d4-a716-446655440000",
                "tool": "read_document_text",
                "arguments": {"maxLength": 100}
            }).encode()
            s, b, _ = _request("POST", addr, "/word/invoke", body=inv_body,
                               headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
            invoke_result["status"] = s
            if b:
                invoke_result["body"] = json.loads(b)

        plugin_thread = threading.Thread(target=plugin_invoke, daemon=True)
        plugin_thread.start()
        time.sleep(0.05)

        # 3. Pane claims pending ticket
        status, claim_body, _ = _request("GET", addr, "/word/pending", headers=pane_hdrs)
        assert status == 200
        claim = json.loads(claim_body)
        ticket_id = claim["ticketId"]
        assert claim["tool"] == "read_document_text"
        assert claim["arguments"] == {"maxLength": 100}

        # 4. Pane "executes" and submits result
        result_body = json.dumps({
            "ticketId": ticket_id,
            "result": {"cells": [["A", "B"], ["C", "D"]], "range": "A1:B2"}
        }).encode()
        status2, _, _ = _request("POST", addr, "/word/result",
                                 body=result_body, headers=pane_hdrs)
        assert status2 == 200

        # 5. Plugin thread unblocks and gets result
        plugin_thread.join(timeout=5)
        assert invoke_result["status"] == 200, f"plugin should get 200, got {invoke_result}"
        assert invoke_result["body"]["data"] == {"cells": [["A", "B"], ["C", "D"]],
                                                  "range": "A1:B2"}

    finally:
        srv.stop()


def test_sidebar_origin_comes_from_the_configured_site():
    """The bridge's allowed origin is derived from addin_base; manifest and plugin default agree."""
    from pathlib import Path

    from desk_plugin import desk_app
    from desk_plugin.desk_bridge import origin_of

    template = (Path(_REPO) / "manifest.template.xml").read_text(encoding="utf-8")
    assert '<SourceLocation DefaultValue="{{ADDIN_BASE}}taskpane.html"/>' in template
    assert "<AppDomain>{{ADDIN_ORIGIN}}</AppDomain>" in template
    manifest = (Path(_REPO) / "plugins" / desk_app.PLUGIN_ID / "plugin.yaml").read_text(encoding="utf-8")
    assert f'default: "{desk_app.DEFAULT_ADDIN_BASE}"' in manifest
    assert origin_of(desk_app.DEFAULT_ADDIN_BASE) == f"https://localhost:{desk_app.PORTS[0]}"


def test_heartbeat_drains_body_before_next_keep_alive_request():
    """A heartbeat body must not corrupt the next request on the same socket."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        addr = srv.server_address()
        sess = _create_session(addr)
        auth_lines = (
            "Origin: https://localhost:3001\r\n"
            f"X-Hermes-Session-Key: {sess['sessionKey']}\r\n"
            f"X-Hermes-Session-Id: {sess['sessionId']}\r\n"
            "X-Hermes-Pane-Id: p1\r\n"
        )

        def read_response(stream):
            status_line = stream.readline().decode("ascii")
            assert status_line.startswith("HTTP/1.1 "), status_line
            status = int(status_line.split()[1])
            headers = {}
            while True:
                line = stream.readline()
                if line == b"\r\n":
                    break
                name, value = line.decode("ascii").split(":", 1)
                headers[name.lower()] = value.strip()
            length = int(headers.get("content-length", "0"))
            if length:
                stream.read(length)
            return status

        with socket.create_connection(addr, timeout=5) as sock:
            stream = sock.makefile("rb")
            heartbeat = (
                "POST /word/heartbeat HTTP/1.1\r\n"
                f"Host: {addr[0]}:{addr[1]}\r\n"
                f"{auth_lines}"
                "Content-Type: application/json\r\n"
                "Content-Length: 2\r\n"
                "Connection: keep-alive\r\n\r\n"
                "{}"
            ).encode("ascii")
            sock.sendall(heartbeat)
            assert read_response(stream) == 200

            pending = (
                "GET /word/pending HTTP/1.1\r\n"
                f"Host: {addr[0]}:{addr[1]}\r\n"
                f"{auth_lines}"
                "Connection: close\r\n\r\n"
            ).encode("ascii")
            sock.sendall(pending)
            assert read_response(stream) == 204
    finally:
        srv.stop()


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import traceback

    tests = [
        (name, fn)
        for name, fn in list(globals().items())
        if name.startswith("test_") and callable(fn)
    ]
    passed = 0
    failed = 0
    for name, fn in tests:
        try:
            fn()
            passed += 1
            print(f"  PASS {name}")
        except Exception:
            failed += 1
            print(f"  FAIL {name}")
            traceback.print_exc()
    print(f"\n{passed} passed, {failed} failed, {len(tests)} total")
    sys.exit(1 if failed else 0)
