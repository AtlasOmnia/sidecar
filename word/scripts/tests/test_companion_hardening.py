"""Focused hardening tests for the Word companion RPC bridge."""
from __future__ import annotations

import http.client
import io
import json
import logging
import os
import socket
import sys
import threading
import time
from typing import Optional, cast
from unittest.mock import patch

import pytest

_REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, _REPO)

from scripts.tests.test_companion_server import _create_session, _request  # noqa: E402


PLUGIN_KEY = "hardening-test-plugin-key"


class FakeClock:
    def __init__(self, value: float = 1000.0) -> None:
        self.value = value

    def __call__(self) -> float:
        return self.value

    def advance(self, seconds: float) -> None:
        self.value += seconds


def _pane_headers(session: dict, pane_id: str = "p1") -> dict[str, str]:
    return {
        "Origin": "https://localhost:3001",
        "X-Hermes-Session-Key": session["sessionKey"],
        "X-Hermes-Session-Id": session["sessionId"],
        "X-Hermes-Pane-Id": pane_id,
    }


def _invoke_body(session_id: str, request_id: str, tool: str = "read_document_text",
                 arguments: Optional[dict] = None) -> bytes:
    return json.dumps({
        "protocolVersion": "1.0",
        "sessionId": session_id,
        "requestId": request_id,
        "tool": tool,
        "arguments": arguments if arguments is not None else {"maxLength": 100},
    }).encode()


def _wait_until(predicate, timeout: float = 2.0) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if predicate():
            return
        threading.Event().wait(0.005)
    raise AssertionError("condition was not reached before timeout")


def test_invoke_timeout_is_504_not_500() -> None:
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
                               invoke_timeout_seconds=0.03)
    srv.start()
    try:
        session = _create_session(srv.server_address())
        status, body, _ = _request(
            "POST", srv.server_address(), "/word/invoke",
            body=_invoke_body(session["sessionId"], "11111111-1111-4111-8111-111111111111"),
            headers={"X-Hermes-Plugin-Key": PLUGIN_KEY},
        )
        assert status == 504, (status, body)
        assert json.loads(body)["category"] == "Timeout"
    finally:
        srv.stop()


def test_request_id_replay_is_at_most_once_and_conflicts_fail_closed() -> None:
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
                               invoke_timeout_seconds=2)
    srv.start()
    try:
        addr = srv.server_address()
        session = _create_session(addr)
        headers = _pane_headers(session)
        request_id = "22222222-2222-4222-8222-222222222222"
        body = _invoke_body(session["sessionId"], request_id)
        responses: list[tuple[int, bytes]] = []

        def invoke() -> None:
            status, response, _ = _request(
                "POST", addr, "/word/invoke", body=body,
                headers={"X-Hermes-Plugin-Key": PLUGIN_KEY},
            )
            responses.append((status, response))

        first = threading.Thread(target=invoke, daemon=True)
        first.start()
        _wait_until(lambda: bool(srv.sessions.get(session["sessionId"]).has_ticket()))
        authoritative = srv.sessions.get(session["sessionId"]).get_ticket()

        replay = threading.Thread(target=invoke, daemon=True)
        replay.start()
        _wait_until(lambda: replay.is_alive())
        assert srv.sessions.get(session["sessionId"]).get_ticket() is authoritative

        status, claim_body, _ = _request("GET", addr, "/word/pending", headers=headers)
        assert status == 200
        ticket_id = json.loads(claim_body)["ticketId"]
        result = json.dumps({"ticketId": ticket_id, "result": {"ok": "once"}}).encode()
        assert _request("POST", addr, "/word/result", body=result, headers=headers)[0] == 200
        first.join(2)
        replay.join(2)
        assert [status for status, _ in responses] == [200, 200]
        assert len({json.loads(response)["ticketId"] for _, response in responses}) == 1

        # Completed replay returns the same ticket/result without a second pane claim.
        replay_status, replay_body, _ = _request(
            "POST", addr, "/word/invoke", body=body,
            headers={"X-Hermes-Plugin-Key": PLUGIN_KEY},
        )
        assert replay_status == 200
        assert json.loads(replay_body)["ticketId"] == ticket_id
        assert _request("GET", addr, "/word/pending", headers=headers)[0] == 204

        # Same requestId with a different payload must fail closed.
        conflict = _invoke_body(session["sessionId"], request_id,
                                tool="insert_text", arguments={"text": "x", "location": "end"})
        status, _, _ = _request(
            "POST", addr, "/word/invoke", body=conflict,
            headers={"X-Hermes-Plugin-Key": PLUGIN_KEY},
        )
        assert status == 409
    finally:
        srv.stop()


def test_automatic_sweep_requeues_once_then_times_out() -> None:
    from desk_plugin.companion.server import WordCompanionServer

    clock = FakeClock()
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
                               clock=clock, invoke_timeout_seconds=120)
    srv.start()
    try:
        addr = srv.server_address()
        session = _create_session(addr)
        headers = _pane_headers(session)
        body = _invoke_body(session["sessionId"], "33333333-3333-4333-8333-333333333333")
        response: dict[str, int] = {}

        def invoke() -> None:
            response["status"] = _request(
                "POST", addr, "/word/invoke", body=body,
                headers={"X-Hermes-Plugin-Key": PLUGIN_KEY},
            )[0]

        thread = threading.Thread(target=invoke, daemon=True)
        thread.start()
        _wait_until(lambda: bool(srv.sessions.get(session["sessionId"]).has_ticket()))
        assert _request("GET", addr, "/word/pending", headers=headers)[0] == 200

        clock.advance(21)
        # Dispatch sweep requeues, then this poll claims the same ticket again.
        assert _request("GET", addr, "/word/pending", headers=headers)[0] == 200
        clock.advance(21)
        # Second lease expiry is terminal and wakes the plugin request.
        assert _request("GET", addr, "/word/pending", headers=headers)[0] == 204
        thread.join(2)
        assert response["status"] == 504
    finally:
        srv.stop()


def test_capacity_limits_sessions_and_handler_threads() -> None:
    from desk_plugin.companion.server import WordCompanionServer

    session_srv = WordCompanionServer(host="127.0.0.1", port=0, max_sessions=1)
    session_srv.start()
    try:
        _create_session(session_srv.server_address(), doc_id="one")
        body = json.dumps({
            "protocolVersion": "1.0", "documentId": "two", "paneId": "p2"
        }).encode()
        status, _, _ = _request(
            "POST", session_srv.server_address(), "/word/session", body=body,
            headers={"Origin": "https://localhost:3001"},
        )
        assert status == 429
    finally:
        session_srv.stop()

    thread_srv = WordCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
        max_handler_threads=1, invoke_timeout_seconds=2,
    )
    thread_srv.start()
    try:
        addr = thread_srv.server_address()
        session = _create_session(addr)
        body = _invoke_body(session["sessionId"], "44444444-4444-4444-8444-444444444444")
        holder = threading.Thread(
            target=lambda: _request("POST", addr, "/word/invoke", body=body,
                                    headers={"X-Hermes-Plugin-Key": PLUGIN_KEY}),
            daemon=True,
        )
        holder.start()
        _wait_until(lambda: bool(thread_srv.sessions.get(session["sessionId"]).has_ticket()))
        status, response, _ = _request(
            "GET", addr, "/word/health",
            headers={"X-Hermes-Plugin-Key": PLUGIN_KEY},
        )
        assert status == 503, (status, response)
        assert json.loads(response) == {"ok": False, "error": "Server busy"}
        for marker in (
            PLUGIN_KEY, "hermes-office-word-companion", "protocolVersion",
            "uptime", "session", "document", "path",
        ):
            assert marker.encode() not in response
    finally:
        thread_srv.stop()


def test_slow_partial_body_is_bounded() -> None:
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
        request_timeout_seconds=0.05,
    )
    srv.start()
    try:
        sock = socket.create_connection(srv.server_address(), timeout=1)
        try:
            request = (
                "POST /word/invoke HTTP/1.1\r\n"
                f"Host: {srv.server_address()[0]}\r\n"
                f"X-Hermes-Plugin-Key: {PLUGIN_KEY}\r\n"
                "Content-Type: application/json\r\n"
                "Content-Length: 100\r\n\r\n"
                "{"
            ).encode("ascii")
            started = time.monotonic()
            sock.sendall(request)
            response = sock.recv(4096)
            assert time.monotonic() - started < 1
            assert b" 408 " in response
            assert b"Connection: close" in response
        finally:
            sock.close()
    finally:
        srv.stop()


def _read_raw_response(sock: socket.socket) -> bytes:
    chunks: list[bytes] = []
    while True:
        try:
            chunk = sock.recv(65536)
        except socket.timeout:
            break
        if not chunk:
            break
        chunks.append(chunk)
    return b"".join(chunks)


def test_rejected_heartbeats_send_one_error_and_do_not_extend_ttl() -> None:
    """Missing and oversized heartbeat bodies fail closed on the raw socket."""
    from desk_plugin.companion.server import WordCompanionServer

    clock = FakeClock()
    srv = WordCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, clock=clock,
    )
    session = srv.sessions.create("document", "pane", "s" * 64)
    before = session.last_heartbeat
    srv.start()
    try:
        host, port = srv.server_address()
        for content_length, expected_status in (
            (None, 411),
            (str(64 * 1024 + 1), 413),
        ):
            sock = socket.create_connection((host, port), timeout=1)
            sock.settimeout(1)
            try:
                lines = [
                    "POST /word/heartbeat HTTP/1.1",
                    f"Host: {host}:{port}",
                    "Origin: https://localhost:3001",
                    f"X-Hermes-Session-Id: {session.session_id}",
                    f"X-Hermes-Session-Key: {session.session_key}",
                    "X-Hermes-Pane-Id: pane",
                    "Connection: close",
                ]
                if content_length is not None:
                    lines.append(f"Content-Length: {content_length}")
                sock.sendall(("\r\n".join(lines) + "\r\n\r\n").encode("ascii"))
                raw = _read_raw_response(sock)
                status_line = raw.split(b"\r\n", 1)[0]
                assert f"HTTP/1.1 {expected_status} ".encode("ascii") in status_line
                assert raw.count(b"HTTP/1.1") == 1
                assert b"200 OK" not in raw
                assert session.last_heartbeat == before
            finally:
                sock.close()
    finally:
        srv.stop()


def test_header_trickle_releases_workers_and_server_recovers() -> None:
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
        max_handler_threads=2, request_timeout_seconds=0.2,
    )
    srv.start()
    host, port = srv.server_address()
    blockers = [socket.create_connection((host, port), timeout=1) for _ in range(2)]
    try:
        for sock in blockers:
            sock.sendall(b"GET /word/health HTTP/1.1\r\nHost: 127.0.0.1\r\nX-Trickle: ")
        for _ in range(8):
            for sock in blockers:
                try:
                    sock.sendall(b"x")
                except OSError:
                    pass
            time.sleep(0.05)

        conn = http.client.HTTPConnection(host, port, timeout=2)
        conn.request("GET", "/word/health", headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})
        response = conn.getresponse()
        assert response.status == 200
        response.read()
        conn.close()
    finally:
        for sock in blockers:
            sock.close()
        srv.stop()


def test_body_trickle_hits_one_absolute_408_deadline() -> None:
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
        request_timeout_seconds=0.2,
    )
    session = srv.sessions.create("document", "pane", "s" * 64)
    srv.start()
    host, port = srv.server_address()
    sock = socket.create_connection((host, port), timeout=1)
    sock.settimeout(1)
    try:
        headers = (
            "POST /word/heartbeat HTTP/1.1\r\n"
            f"Host: {host}:{port}\r\n"
            "Origin: https://localhost:3001\r\n"
            f"X-Hermes-Session-Id: {session.session_id}\r\n"
            f"X-Hermes-Session-Key: {session.session_key}\r\n"
            "X-Hermes-Pane-Id: pane\r\n"
            "Content-Type: application/json\r\n"
            "Content-Length: 100\r\n"
            "Connection: close\r\n\r\n"
        ).encode("ascii")
        sock.sendall(headers)
        for _ in range(8):
            try:
                sock.sendall(b"x")
            except OSError:
                break
            time.sleep(0.05)
        raw = _read_raw_response(sock)
        assert b"408 Request Timeout" in raw.split(b"\r\n", 1)[0]
        assert raw.count(b"HTTP/1.1") == 1
    finally:
        sock.close()
        srv.stop()


def test_header_watchdog_is_cancelled_before_long_invoke_wait() -> None:
    """A parsed /invoke may outlive the HTTP-header deadline while Office executes."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
        request_timeout_seconds=0.05, invoke_timeout_seconds=1,
    )
    srv.start()
    try:
        addr = srv.server_address()
        session = _create_session(addr)
        pane_headers = _pane_headers(session)
        result: dict[str, object] = {}

        def invoke() -> None:
            status, body, _ = _request(
                "POST", addr, "/word/invoke",
                body=_invoke_body(session["sessionId"], "55555555-5555-4555-8555-555555555555"),
                headers={"X-Hermes-Plugin-Key": PLUGIN_KEY},
            )
            result.update(status=status, body=body)

        thread = threading.Thread(target=invoke, daemon=True)
        thread.start()
        _wait_until(lambda: bool(srv.sessions.get(session["sessionId"]).has_ticket()))
        status, pending, _ = _request("GET", addr, "/word/pending", headers=pane_headers)
        assert status == 200
        time.sleep(0.1)
        ticket_id = json.loads(pending)["ticketId"]
        result_body = json.dumps({"ticketId": ticket_id, "result": {"ok": True}}).encode()
        assert _request("POST", addr, "/word/result", body=result_body, headers=pane_headers)[0] == 200
        thread.join(2)
        assert result["status"] == 200
    finally:
        srv.stop()


def _history_ticket(session_id: str, index: int, result: object | None = None):
    from desk_plugin.companion.session import Ticket

    ticket = Ticket(
        ticket_id=f"{index:08x}-0000-4000-8000-000000000001",
        session_id=session_id,
        request_id=f"{index:08x}-0000-4000-8000-000000000002",
        tool="read_document_text",
        arguments={"index": index},
    )
    if result is not None:
        ticket.set_result(result)
    return ticket


def test_protocol_preserves_body_caps_and_sets_operational_memory_limits() -> None:
    from desk_plugin.companion.protocol import (
        MAX_INVOKE_BODY,
        MAX_REQUEST_HISTORY,
        MAX_RESULT_BODY,
        MAX_RETAINED_HISTORY_BYTES,
        MAX_SESSIONS,
    )

    assert MAX_INVOKE_BODY == 1 * 1024 * 1024
    assert MAX_RESULT_BODY == 10 * 1024 * 1024
    assert MAX_REQUEST_HISTORY == 4
    assert MAX_RETAINED_HISTORY_BYTES == 12 * 1024 * 1024
    assert MAX_SESSIONS == 16


def test_request_history_keeps_four_and_never_evicts_active_ticket() -> None:
    from desk_plugin.companion.session import Session

    session = Session("session", "document", "pane", "key")
    for index in range(1, 5):
        ticket = _history_ticket(session.session_id, index)
        assert session.begin_ticket(ticket)[0] == "created"
        ticket.set_result({"ok": True, "index": index})

    active = _history_ticket(session.session_id, 5)
    assert session.begin_ticket(active)[0] == "created"
    assert len(session._request_history) == 4
    assert session._request_history.get(active.request_id) is active
    assert active.state == "pending"
    assert session.find_ticket("00000001-0000-4000-8000-000000000001") is None


def test_terminal_payload_enforces_twelve_mib_history_budget() -> None:
    from desk_plugin.companion.protocol import MAX_RETAINED_HISTORY_BYTES
    from desk_plugin.companion.session import Session

    session = Session("session", "document", "pane", "key")
    tickets = []
    for index in range(1, 4):
        ticket = _history_ticket(session.session_id, index)
        assert session.begin_ticket(ticket)[0] == "created"
        ticket.set_result({"data": "x" * (5 * 1024 * 1024)})
        tickets.append(ticket)

    retained = list(session._request_history.values())
    assert tickets[0] not in retained
    assert tickets[1:] == retained
    assert sum(ticket.retained_size_bytes() for ticket in retained) <= MAX_RETAINED_HISTORY_BYTES


def test_default_session_capacity_is_sixteen() -> None:
    from desk_plugin.companion.session import SessionManager

    manager = SessionManager()
    for index in range(16):
        manager.create(f"document-{index}", f"pane-{index}", "s" * 64)
    assert manager.active_count == 16
    with pytest.raises(RuntimeError, match="capacity"):
        manager.create("document-overflow", "pane-overflow", "s" * 64)


def test_no_content_responses_have_zero_length() -> None:
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        session = _create_session(srv.server_address())
        status, body, headers = _request(
            "GET", srv.server_address(), "/word/pending",
            headers=_pane_headers(session),
        )
        assert status == 204
        assert body == b""
        assert headers.get("Content-Length") == "0"
    finally:
        srv.stop()


def test_cors_is_exact_and_logs_exclude_content() -> None:
    from desk_plugin.companion.server import WordCompanionServer

    stream = io.StringIO()
    capture = logging.StreamHandler(stream)
    logger = logging.getLogger("word_companion")
    logger.addHandler(capture)
    secret_markers = [
        "DOC-SENTINEL", "PANE-SENTINEL", "TOOL-SENTINEL",
        "ARG-SENTINEL", "RESULT-SENTINEL", "KEY-SENTINEL",
    ]
    key = "KEY-SENTINEL-" + ("k" * 32)
    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=key,
                               invoke_timeout_seconds=1)
    srv.start()
    try:
        addr = srv.server_address()
        body = json.dumps({
            "protocolVersion": "1.0",
            "documentId": "DOC-SENTINEL",
            "paneId": "PANE-SENTINEL",
        }).encode()
        status, _, headers = _request(
            "POST", addr, "/word/session", body=body,
            headers={"Origin": "https://localhost:3001"},
        )
        assert status == 201
        assert headers.get("Access-Control-Allow-Origin") == "https://localhost:3001"
        assert "Origin" in headers.get("Vary", "")

        bad_status, _, bad_headers = _request(
            "POST", addr, "/word/session", body=body,
            headers={"Origin": "https://evil.invalid"},
        )
        assert bad_status == 403
        assert "Access-Control-Allow-Origin" not in bad_headers
        assert _request(
            "OPTIONS", addr, "/word/invoke",
            headers={"Origin": "https://localhost:3001"},
        )[0] == 405
    finally:
        srv.stop()
        logger.removeHandler(capture)

    log_text = stream.getvalue()
    for marker in secret_markers:
        assert marker not in log_text, f"metadata log leaked {marker}"


def test_request_parser_disconnect_is_suppressed_without_second_response() -> None:
    """Connection resets raised by the base request parser must not escape."""
    from http.server import BaseHTTPRequestHandler
    from desk_plugin.companion.handler import CompanionHandler
    from desk_plugin.companion.server import WordCompanionServer

    server = WordCompanionServer(host="127.0.0.1", port=0, request_timeout_seconds=60)
    handler = object.__new__(CompanionHandler)
    handler._server_ref = server
    handler.close_connection = False
    original = BaseHTTPRequestHandler.handle_one_request

    def raise_reset(self: BaseHTTPRequestHandler) -> None:
        raise ConnectionResetError("peer reset during request parse")

    setattr(BaseHTTPRequestHandler, "handle_one_request", raise_reset)
    try:
        handler.handle_one_request()
    finally:
        setattr(BaseHTTPRequestHandler, "handle_one_request", original)

    assert handler.close_connection is True


def test_session_auth_validates_all_three_values_without_short_circuiting() -> None:
    """Pane auth must compare session id, key, and pane for every tuple."""
    from desk_plugin.companion import handler as handler_module
    from desk_plugin.companion.handler import constant_time_compare as real_compare
    from desk_plugin.companion.server import WordCompanionServer

    server = WordCompanionServer(host="127.0.0.1", port=0, plugin_key="plugin-key")
    server.start()
    try:
        address = server.server_address()
        session = _create_session(address)
        session_id = session["sessionId"]
        session_key = session["sessionKey"]
        pane_id = "p1"

        def probe(headers: dict[str, str]) -> tuple[int, list[tuple[str, str]]]:
            comparisons: list[tuple[str, str]] = []

            def recording_compare(left: str, right: str) -> bool:
                comparisons.append((left, right))
                return real_compare(left, right)

            with patch.object(handler_module, "constant_time_compare", side_effect=recording_compare):
                status, _, _ = _request("GET", address, "/word/pending", headers=headers)
            return status, comparisons

        def assert_auth_comparisons(
            comparisons: list[tuple[str, str]],
            expected: list[tuple[str, str]],
        ) -> None:
            observed = [pair for pair in comparisons if pair in expected]
            assert observed[:3] == expected, (comparisons, expected)

        valid = {
            "Origin": "https://localhost:3001",
            "X-Hermes-Session-Id": session_id,
            "X-Hermes-Session-Key": session_key,
            "X-Hermes-Pane-Id": pane_id,
        }
        status, comparisons = probe(valid)
        assert status == 204
        assert_auth_comparisons(comparisons, [
            (session_id, session_id),
            (session_key, session_key),
            (pane_id, pane_id),
        ])

        wrong_key = {**valid, "X-Hermes-Session-Key": "wrong-key"}
        status, comparisons = probe(wrong_key)
        assert status == 401
        assert_auth_comparisons(comparisons, [
            (session_id, session_id),
            ("wrong-key", session_key),
            (pane_id, pane_id),
        ])

        wrong_pane = {**valid, "X-Hermes-Pane-Id": "wrong-pane"}
        status, comparisons = probe(wrong_pane)
        assert status == 401
        assert_auth_comparisons(comparisons, [
            (session_id, session_id),
            (session_key, session_key),
            ("wrong-pane", pane_id),
        ])

        unknown_session = {**valid, "X-Hermes-Session-Id": "unknown-session"}
        status, comparisons = probe(unknown_session)
        assert status == 401
        assert_auth_comparisons(comparisons, [
            ("unknown-session", "__dummy__"),
            (session_key, "__dummy__"),
            (pane_id, "__dummy__"),
        ])
    finally:
        server.stop()


@pytest.mark.parametrize("mode", ["eof", "bytes", "deadline", "timeout"])
def test_overload_drain_follows_response_and_is_absolutely_bounded(mode) -> None:
    from desk_plugin.companion import server as server_module

    server = server_module._BoundedThreadingHTTPServer.__new__(server_module._BoundedThreadingHTTPServer)
    server._handler_semaphore = threading.BoundedSemaphore(1)
    server._handler_semaphore.acquire()
    server._request_timeout_seconds = 10.0
    actions = []
    received = []

    class FakeSocket:
        chunks = [b"request headers", b"body", b""]

        def settimeout(self, value):
            actions.append(("timeout", value))

        def sendall(self, value):
            actions.append(("send", value))

        def shutdown(self, how):
            actions.append(("shutdown", how))

        def recv(self, size):
            actions.append(("recv", size))
            received.append(size)
            if mode == "timeout":
                raise socket.timeout("synthetic peer stall")
            return self.chunks.pop(0) if mode == "eof" else b"x" * size

    fake = FakeSocket()
    server.close_request = lambda request: actions.append(("close", request))
    clock = (lambda: 0.0) if mode != "deadline" else iter([0.0, 0.1, 0.3]).__next__
    with patch.object(server_module.time, "monotonic", side_effect=clock):
        server.process_request(cast(socket.socket, fake), ("127.0.0.1", 1))
    assert actions[-1] == ("close", fake)
    assert actions[1][0] == "send"
    assert actions[1][1].startswith(b"HTTP/1.1 503 ")
    assert actions.index(("shutdown", socket.SHUT_WR)) < next(
        index for index, action in enumerate(actions) if action[0] == "recv"
    )
    assert 0 < sum(received) <= 64 * 1024
    assert all(0 < action[1] <= 0.2 for action in actions if action[0] == "timeout")
    if mode == "eof":
        assert len(received) == 3
    if mode == "deadline":
        assert len(received) == 1


def _json_body_with_depth(depth: int, *, session_id: str | None = None) -> bytes:
    """Build a valid protocol body whose maximum container depth is exact."""
    assert depth >= 1
    nested_arrays = b"[" * (depth - 1) + b"0" + b"]" * (depth - 1)
    if session_id is None:
        return (
            b'{"protocolVersion":"1.0","documentId":"depth-doc",'
            b'"paneId":"depth-pane","nested":' + nested_arrays + b"}"
        )
    return (
        b'{"sessionId":"' + session_id.encode() + b'","nested":'
        + nested_arrays + b"}"
    )


def test_json_nesting_depth_64_is_accepted_through_authenticated_endpoint() -> None:
    """The contract boundary accepts exactly 64 structural container levels."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        status, body, _ = _request(
            "POST",
            srv.server_address(),
            "/word/session",
            body=_json_body_with_depth(64),
            headers={"Origin": "https://localhost:3001"},
        )
        assert status == 201, (status, body)
        assert srv.sessions.active_count == 1
    finally:
        srv.stop()


def test_json_nesting_depth_65_is_rejected_before_parse_and_route_mutation() -> None:
    """Depth 65 is a fixed close-delimited 400 before endpoint mutation/parse."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        with patch(
            "desk_plugin.companion.handler.json.loads",
            side_effect=AssertionError("depth guard must run before json.loads"),
        ):
            status, body, headers = _request(
                "POST",
                srv.server_address(),
                "/word/session",
                body=_json_body_with_depth(65),
                headers={"Origin": "https://localhost:3001"},
            )
        assert status == 400, (status, body)
        assert headers.get("Connection") == "close"
        assert body == b'{"ok": false, "error": "Bad Request: JSON nesting too deep", "category": "Validation"}'
        assert srv.sessions.active_count == 0
    finally:
        srv.stop()


def test_json_nesting_ignores_delimiters_inside_ordinary_and_escaped_strings() -> None:
    """String contents cannot increase structural depth, including escapes."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    srv.start()
    try:
        body = json.dumps({
            "protocolVersion": "1.0",
            "documentId": "depth-doc",
            "paneId": "depth-pane",
            "note": '{["\\"]} ordinary braces and escaped quote',
        }).encode()
        status, response, _ = _request(
            "POST",
            srv.server_address(),
            "/word/session",
            body=body,
            headers={"Origin": "https://localhost:3001"},
        )
        assert status == 201, (status, response)
        assert srv.sessions.active_count == 1
    finally:
        srv.stop()


def test_json_nesting_malformed_within_depth_uses_fixed_invalid_json_response() -> None:
    """Malformed JSON that stays within the cap follows the existing 400 path."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        status, body, headers = _request(
            "POST",
            srv.server_address(),
            "/word/session",
            body=b'{"protocolVersion":"1.0","documentId":"d1",',
            headers={"Origin": "https://localhost:3001"},
        )
        assert status == 400, (status, body)
        assert headers.get("Connection") == "close"
        assert body == b'{"ok": false, "error": "Bad Request: invalid JSON", "category": "Validation"}'
        assert srv.sessions.active_count == 0
    finally:
        srv.stop()


def test_json_nesting_recursion_error_from_decode_is_bounded_without_leakage() -> None:
    """A parser recursion failure becomes the fixed invalid-JSON response."""
    from desk_plugin.companion.server import WordCompanionServer

    srv = WordCompanionServer(host="127.0.0.1", port=0)
    srv.start()
    try:
        with patch(
            "desk_plugin.companion.handler.json.loads",
            side_effect=RecursionError("secret parser internals"),
        ):
            status, body, headers = _request(
                "POST",
                srv.server_address(),
                "/word/session",
                body=b'{"protocolVersion":"1.0","documentId":"d1","paneId":"p1"}',
                headers={"Origin": "https://localhost:3001"},
            )
        assert status == 400, (status, body)
        assert headers.get("Connection") == "close"
        assert body == b'{"ok": false, "error": "Bad Request: invalid JSON", "category": "Validation"}'
        assert b"RecursionError" not in body
        assert b"secret parser internals" not in body
    finally:
        srv.stop()


if __name__ == "__main__":
    import traceback

    tests = [(name, value) for name, value in list(globals().items())
             if name.startswith("test_") and callable(value)]
    passed = 0
    failed = 0
    for name, test in tests:
        try:
            test()
            passed += 1
            print(f"  PASS {name}")
        except Exception:
            failed += 1
            print(f"  FAIL {name}")
            traceback.print_exc()
    print(f"\n{passed} passed, {failed} failed, {len(tests)} total")
    raise SystemExit(1 if failed else 0)
