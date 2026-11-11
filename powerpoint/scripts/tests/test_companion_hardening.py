"""Companion hardening tests — edge cases, limits, and error handling."""
from __future__ import annotations

import http.client
import socket
import sys
import time
from pathlib import Path
from unittest.mock import patch

_REPO_ROOT = Path(__file__).resolve().parents[2]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

import pytest

from desk_plugin.companion.session import SessionManager, Session, Ticket
from desk_plugin.companion.protocol import (
    MAX_REQUEST_HISTORY,
    MAX_RETAINED_HISTORY_BYTES,
    MAX_SESSIONS,
    generate_uuid,
)
from desk_plugin.companion.server import PowerPointCompanionServer


class _Clock:
    def __init__(self, start: float = 0.0):
        self._t = start

    def __call__(self) -> float:
        return self._t

    def advance(self, seconds: float) -> None:
        self._t += seconds


class TestSessionManagerLimits:
    def test_session_capacity(self):
        clock = _Clock()
        sm = SessionManager(clock=clock, max_sessions=3)
        assert sm.active_count == 0

        sm.create("doc-1", "pane-1", "key-1")
        sm.create("doc-2", "pane-2", "key-2")
        sm.create("doc-3", "pane-3", "key-3")
        assert sm.active_count == 3

        with pytest.raises(RuntimeError, match="capacity"):
            sm.create("doc-4", "pane-4", "key-4")

    def test_ticket_requeue_max_once(self):
        clock = _Clock()
        ticket = Ticket(
            ticket_id="t1", session_id="s1", request_id="r1",
            tool="read_all_slides", arguments={},
            created_at=clock(), _clock=clock,
        )
        assert ticket.state == "pending"
        assert ticket.claim("p1", clock)
        assert ticket.state == "claimed"

        # Requeue once
        clock.advance(25)  # Past lease
        assert ticket.requeue(clock)
        assert ticket.state == "pending"
        assert ticket.requeue_count == 1

        # Claim again
        assert ticket.claim("p1", clock)
        clock.advance(25)

        # Second requeue should fail
        assert not ticket.requeue(clock)
        assert ticket.state == "timeout"

    def test_ticket_invoke_timeout(self):
        clock = _Clock()
        ticket = Ticket(
            ticket_id="t2", session_id="s2", request_id="r2",
            tool="read_all_slides", arguments={},
            created_at=clock(), _clock=clock,
            invoke_timeout_seconds=60,
        )
        clock.advance(59)
        assert not ticket.is_invoke_expired(clock)
        clock.advance(2)
        assert ticket.is_invoke_expired(clock)

    def test_session_expiry(self):
        clock = _Clock()
        sm = SessionManager(clock=clock)
        session = sm.create("doc-exp", "pane-exp", "key-exp")
        assert not session.is_expired(clock)

        clock.advance(61)
        assert session.is_expired(clock)

    def test_shutdown_cancels_all(self):
        clock = _Clock()
        sm = SessionManager(clock=clock)
        s1 = sm.create("d1", "p1", "k1")
        s2 = sm.create("d2", "p2", "k2")

        t1 = Ticket("t1", s1.session_id, "r1", "tool", {}, created_at=clock(), _clock=clock)
        t2 = Ticket("t2", s2.session_id, "r2", "tool", {}, created_at=clock(), _clock=clock)
        s1.set_ticket(t1)
        s2.set_ticket(t2)

        cancelled = sm.shutdown()
        assert len(cancelled) == 2
        assert t1.state == "cancelled"
        assert t2.state == "cancelled"
        assert sm.active_count == 0

    def test_request_history_bounded_by_count(self):
        clock = _Clock()
        session = Session("sid", "doc", "pane", "key", _clock=clock)
        for i in range(MAX_REQUEST_HISTORY * 2):
            ticket = Ticket(f"t{i}", "sid", f"req-{i}", "tool", {},
                          created_at=clock(), _clock=clock)
            ticket.set_result("ok")
            session.set_ticket(ticket)
            session.clear_ticket()
        assert session.find_ticket("t0") is None
        assert session.find_ticket(f"t{MAX_REQUEST_HISTORY * 2 - 1}") is not None

    def test_request_history_bounded_by_serialized_bytes(self):
        clock = _Clock()
        session = Session("sid", "doc", "pane", "key", _clock=clock)
        one_megabyte = "x" * (1024 * 1024)
        for i in range(10):
            ticket = Ticket(f"large-{i}", "sid", f"large-req-{i}", "tool",
                          {"slide": i}, created_at=clock(), _clock=clock)
            ticket.set_result({"text": one_megabyte})
            session.set_ticket(ticket)
            session.clear_ticket()
            session.trim_history()
        assert session.retained_history_bytes <= MAX_RETAINED_HISTORY_BYTES
        assert session.find_ticket("large-0") is None
        assert session.find_ticket("large-9") is not None

    def test_sweep_expired_sessions(self):
        clock = _Clock()
        sm = SessionManager(clock=clock)
        s1 = sm.create("d1", "p1", "k1")
        s2 = sm.create("d2", "p2", "k2")

        t1 = Ticket("t1", s1.session_id, "r1", "tool", {}, created_at=clock(), _clock=clock)
        s1.set_ticket(t1)

        # Advance past TTL
        clock.advance(61)
        failed = sm.sweep(clock)
        assert len(failed) >= 1
        # s1 and s2 should both be expired
        assert sm.active_count == 0

    def test_pane_id_enforcement(self):
        sm = SessionManager()
        session = sm.create("doc", "pane-real", "key")
        assert sm.delete_for_pane(session.session_id, "pane-real") is True

        session2 = sm.create("doc2", "pane-owner", "key2")
        assert sm.delete_for_pane(session2.session_id, "pane-wrong") is False

    def test_delete_cancels_ticket(self):
        sm = SessionManager()
        session = sm.create("doc", "pane", "key")
        ticket = Ticket("t-del", session.session_id, "r-del", "tool", {})
        session.set_ticket(ticket)
        assert ticket.state == "pending"

        sm.delete(session.session_id)
        assert ticket.state == "cancelled"
        assert sm.active_count == 0

    def test_metrics_counts(self):
        sm = SessionManager()
        s1 = sm.create("d1", "p1", "k1")
        t1 = Ticket("t1", s1.session_id, "r1", "tool", {})
        s1.set_ticket(t1)

        metrics = sm.get_metrics()
        assert metrics["sessionsActive"] == 1
        assert metrics["ticketsPending"] == 1
        assert metrics["ticketsCompleted"] == 0

        t1.set_result("done")
        metrics = sm.get_metrics()
        assert metrics["ticketsCompleted"] == 1


class TestCompanionServerConstructor:
    def test_loopback_only(self):
        with pytest.raises(ValueError, match="loopback"):
            PowerPointCompanionServer(host="10.0.0.1", plugin_key="x" * 32)

    def test_positive_timeout_required(self):
        with pytest.raises(ValueError, match="positive"):
            PowerPointCompanionServer(invoke_timeout_seconds=0, plugin_key="x" * 32)

    def test_request_timeout_required(self):
        with pytest.raises(ValueError, match="positive"):
            PowerPointCompanionServer(request_timeout_seconds=-1, plugin_key="x" * 32)

    def test_default_port_is_8652(self):
        server = PowerPointCompanionServer(plugin_key="x" * 32)
        assert server.DEFAULT_PORT == 8652

    def test_empty_host_rejected(self):
        with pytest.raises(ValueError, match="empty"):
            PowerPointCompanionServer(host="", plugin_key="x" * 32)


class TestHeartbeatKeepAlive:
    """Rejected heartbeat bodies must not mutate state or double-write responses."""

    @staticmethod
    def _request(content_length: str | None) -> tuple[bytes, float]:
        clock = _Clock(100.0)
        server = PowerPointCompanionServer(
            host="127.0.0.1", port=0, plugin_key="p" * 64, clock=clock,
        )
        session = server.sessions.create("deck", "pane", "s" * 64)
        before = session.last_heartbeat
        server.start()
        host, port = server.server_address()
        sock = socket.create_connection((host, port), timeout=2)
        sock.settimeout(1)
        try:
            lines = [
                "POST /powerpoint/heartbeat HTTP/1.1",
                f"Host: {host}:{port}",
                "Origin: https://localhost:3002",
                f"X-Hermes-Session-Id: {session.session_id}",
                f"X-Hermes-Session-Key: {session.session_key}",
                "X-Hermes-Pane-Id: pane",
                "Connection: close",
            ]
            if content_length is not None:
                lines.append(f"Content-Length: {content_length}")
            sock.sendall(("\r\n".join(lines) + "\r\n\r\n").encode("ascii"))
            chunks: list[bytes] = []
            while True:
                try:
                    chunk = sock.recv(65536)
                except socket.timeout:
                    break
                if not chunk:
                    break
                chunks.append(chunk)
            return b"".join(chunks), session.last_heartbeat - before
        finally:
            sock.close()
            server.stop()

    @pytest.mark.parametrize(("content_length", "expected_status"), [
        (None, 411),
        (str(64 * 1024 + 1), 413),
    ])
    def test_rejected_heartbeat_is_fail_closed(self, content_length, expected_status):
        raw, heartbeat_delta = self._request(content_length)
        status_line = raw.split(b"\r\n", 1)[0]
        assert f"HTTP/1.1 {expected_status} ".encode("ascii") in status_line
        assert raw.count(b"HTTP/1.1") == 1
        assert b"200 OK" not in raw
        assert heartbeat_delta == 0.0


class TestAbsoluteRequestDeadlines:
    def test_header_trickle_cannot_hold_all_worker_slots(self):
        server = PowerPointCompanionServer(
            host="127.0.0.1", port=0, plugin_key="p" * 64,
            max_handler_threads=2, request_timeout_seconds=0.2,
        )
        server.start()
        host, port = server.server_address()
        blockers = [socket.create_connection((host, port), timeout=1) for _ in range(2)]
        try:
            for sock in blockers:
                sock.sendall(b"GET /powerpoint/health HTTP/1.1\r\nHost: 127.0.0.1\r\nX-Trickle: ")
            for _ in range(8):
                for sock in blockers:
                    try:
                        sock.sendall(b"x")
                    except OSError:
                        pass
                time.sleep(0.05)

            conn = http.client.HTTPConnection(host, port, timeout=2)
            conn.request("GET", "/powerpoint/health", headers={"X-Hermes-Plugin-Key": "p" * 64})
            response = conn.getresponse()
            assert response.status == 200
            response.read()
            conn.close()
        finally:
            for sock in blockers:
                sock.close()
            server.stop()

    def test_body_trickle_hits_absolute_deadline_once(self):
        server = PowerPointCompanionServer(
            host="127.0.0.1", port=0, plugin_key="p" * 64,
            request_timeout_seconds=0.2,
        )
        session = server.sessions.create("deck", "pane", "s" * 64)
        server.start()
        host, port = server.server_address()
        sock = socket.create_connection((host, port), timeout=1)
        sock.settimeout(1)
        try:
            headers = (
                "POST /powerpoint/heartbeat HTTP/1.1\r\n"
                f"Host: {host}:{port}\r\n"
                "Origin: https://localhost:3002\r\n"
                f"X-Hermes-Session-Id: {session.session_id}\r\n"
                f"X-Hermes-Session-Key: {session.session_key}\r\n"
                "X-Hermes-Pane-Id: pane\r\n"
                "Content-Length: 100\r\n"
                "Content-Type: application/json\r\n"
                "Connection: close\r\n\r\n"
            ).encode("ascii")
            sock.sendall(headers)
            for _ in range(8):
                try:
                    sock.sendall(b"x")
                except OSError:
                    break
                time.sleep(0.05)
            chunks: list[bytes] = []
            while True:
                try:
                    chunk = sock.recv(65536)
                except socket.timeout:
                    break
                if not chunk:
                    break
                chunks.append(chunk)
            raw = b"".join(chunks)
            assert b"408 Request Timeout" in raw.split(b"\r\n", 1)[0]
            assert raw.count(b"HTTP/1.1") == 1
        finally:
            sock.close()
            server.stop()


def test_request_parser_disconnect_is_suppressed_without_second_response() -> None:
    """Connection resets raised by the base request parser must not escape."""
    from http.server import BaseHTTPRequestHandler
    from desk_plugin.companion.handler import CompanionHandler
    from desk_plugin.companion.server import PowerPointCompanionServer

    server = PowerPointCompanionServer(
        host="127.0.0.1", port=0, request_timeout_seconds=60,
    )
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
    import http.client

    from desk_plugin.companion import handler as handler_module
    from desk_plugin.companion.handler import constant_time_compare as real_compare

    server = PowerPointCompanionServer(host="127.0.0.1", port=0, plugin_key="p" * 64)
    server.start()
    try:
        host, port = server.server_address()
        session = server.sessions.create("deck", "p1", "s" * 64)
        session_id = session.session_id
        session_key = session.session_key
        pane_id = "p1"

        def probe(headers: dict[str, str]) -> tuple[int, list[tuple[str, str]]]:
            comparisons: list[tuple[str, str]] = []

            def recording_compare(left: str, right: str) -> bool:
                comparisons.append((left, right))
                return real_compare(left, right)

            with patch.object(handler_module, "constant_time_compare", side_effect=recording_compare):
                connection = http.client.HTTPConnection(host, port, timeout=5)
                try:
                    connection.request("GET", "/powerpoint/pending", headers=headers)
                    response = connection.getresponse()
                    status = response.status
                    response.read()
                finally:
                    connection.close()
            return status, comparisons

        def assert_auth_comparisons(
            comparisons: list[tuple[str, str]],
            expected: list[tuple[str, str]],
        ) -> None:
            observed = [pair for pair in comparisons if pair in expected]
            assert observed[:3] == expected, (comparisons, expected)

        valid = {
            "Origin": "https://localhost:3002",
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
