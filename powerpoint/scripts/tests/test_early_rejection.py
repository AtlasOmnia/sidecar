"""Early rejection must not leave a bounded request body on the connection."""
from __future__ import annotations

import io
import socket
from email.message import Message
from types import SimpleNamespace
from typing import Any

import pytest

from desk_plugin.companion.handler import CompanionHandler
from desk_plugin.companion.server import PowerPointCompanionServer
from scripts.tests.test_companion_server import _post, _session_headers


@pytest.mark.parametrize(("path", "headers", "status"), [
    ("/powerpoint/heartbeat", {"Origin": "https://localhost:3002"}, 401),
    ("/powerpoint/session", {"Origin": "https://invalid.example"}, 403),
    ("/powerpoint/invoke", {}, 401),
    ("/unknown", {}, 404),
])
def test_early_rejection_drains_bounded_body_and_closes(path, headers, status):
    handler: Any = object.__new__(CompanionHandler)
    handler.path = path
    handler.command = "POST"
    handler.close_connection = False
    handler.headers = Message()
    for name, value in {"Content-Length": "2", **headers}.items():
        handler.headers[name] = value
    handler.rfile = io.BytesIO(b"{}next-request")
    handler.connection = SimpleNamespace(settimeout=lambda timeout: None)
    observed = []
    handler._send_error = lambda code, message, **kwargs: observed.append(
        (code, handler.rfile.tell(), handler.close_connection, kwargs.get("extra_headers")))
    handler._dispatch("POST")
    assert observed == [(status, 2, True, {"Connection": "close"})]
    assert handler.rfile.read() == b"next-request"


@pytest.mark.parametrize("length", ["65537", "bad", "-1"])
def test_early_rejection_does_not_drain_unbounded_or_invalid_body(length):
    handler: Any = object.__new__(CompanionHandler)
    handler.path = "/powerpoint/invoke"
    handler.command = "POST"
    handler.close_connection = False
    handler.headers = Message()
    handler.headers["Content-Length"] = length
    handler.rfile = io.BytesIO(b"do-not-read")
    handler.connection = SimpleNamespace(settimeout=lambda timeout: None)
    observed = []
    handler._send_error = lambda code, message, **kwargs: observed.append(code)
    handler._dispatch("POST")
    assert observed == [401]
    assert handler.rfile.tell() == 0
    assert handler.close_connection is True


def test_early_rejection_read_timeout_preserves_one_auth_error():
    handler: Any = object.__new__(CompanionHandler)
    handler.path = "/powerpoint/invoke"
    handler.command = "POST"
    handler.close_connection = False
    handler.headers = Message()
    handler.headers["Content-Length"] = "2"
    timeouts = []
    handler.connection = SimpleNamespace(settimeout=timeouts.append)

    def timeout(size):
        raise socket.timeout("synthetic body timeout")

    handler.rfile = SimpleNamespace(read1=timeout)
    observed = []
    handler._send_error = lambda code, message, **kwargs: observed.append(code)
    handler._dispatch("POST")
    assert observed == [401]
    assert handler.close_connection is True
    assert timeouts
    assert all(0 < value <= 0.25 for value in timeouts)


def test_wrong_key_responses_are_complete_on_real_loopback():
    server = PowerPointCompanionServer(host="127.0.0.1", port=0, plugin_key="synthetic-" + "p" * 64)
    session = server.sessions.create("fixture-deck", "fixture-pane", "s" * 64)
    server.start()
    try:
        _, port = server.server_address()
        for _ in range(20):
            status, data = _post("/powerpoint/heartbeat", {}, port=port,
                                extra_headers=_session_headers(session.session_id, "wrong-key", "fixture-pane"))
            assert status == 401
            assert data == {"ok": False, "error": "Unauthorized", "category": "Auth"}
    finally:
        server.stop()
