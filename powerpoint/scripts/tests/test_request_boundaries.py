"""Synthetic request-boundary regressions; no Office or Hermes process needed."""
from __future__ import annotations

import http.client
import importlib.util
import io
import json
import logging
from pathlib import Path
import socket
import sys
import threading
import time
from email.message import Message
from unittest.mock import Mock

import pytest

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from desk_plugin.companion.handler import CompanionHandler
from desk_plugin.companion.server import PowerPointCompanionServer



def headers(*pairs):
    result = Message()
    for key, value in pairs:
        result[key] = value
    return result


def raw_reply(address, request):
    with socket.create_connection(address, timeout=2) as sock:
        sock.settimeout(2)
        sock.sendall(request)
        # Half-close ensures rejected unread bytes do not reset Windows replies.
        sock.shutdown(socket.SHUT_WR)
        data = bytearray()
        while True:
            chunk = sock.recv(4096)
            if not chunk:
                return bytes(data)
            data.extend(chunk)


@pytest.mark.parametrize("length", ["+2", "02", "2, 2", "-1", "\u0662"])
def test_companion_rejects_noncanonical_length_without_read(length):
    handler = object.__new__(CompanionHandler)
    handler.headers = headers(("Content-Length", length))
    handler.connection = Mock()
    handler.rfile = Mock(wraps=io.BytesIO(b"{}"))
    handler._send_error = Mock()
    assert handler._read_body(1024, time.monotonic()) is None
    handler._send_error.assert_called_once()
    assert handler._send_error.call_args.args[0] == 400
    handler.rfile.read.assert_not_called()


@pytest.mark.parametrize("values", [("2", "2"), ("2", "3")])
def test_companion_rejects_duplicate_length(values):
    handler = object.__new__(CompanionHandler)
    handler.headers = headers(*[("Content-Length", value) for value in values])
    handler.connection = Mock()
    handler.rfile = Mock(wraps=io.BytesIO(b"{}"))
    handler._send_error = Mock()
    assert handler._read_body(1024, time.monotonic()) is None
    assert handler._send_error.call_args.args[0] == 400
    handler.rfile.read.assert_not_called()


def test_companion_depth_limit_precedes_recursive_parser():
    handler = object.__new__(CompanionHandler)
    handler._send_error = Mock()
    assert handler._parse_json(b"[" * 65 + b"0" + b"]" * 65, time.monotonic()) is None
    assert handler._send_error.call_args.args[0] == 400
    body = json.dumps({"text": '[\\\"' * 100}).encode()
    assert handler._parse_json(body, time.monotonic()) == json.loads(body)


@pytest.mark.parametrize("content_type", [None, "text/plain", "application/jsonp", "application/json;", "application/json; charset=utf-8; CHARSET=x", "application/json; note=\"unterminated"])
def test_companion_media_type_gate(content_type):
    handler = object.__new__(CompanionHandler)
    pairs = [] if content_type is None else [("Content-Type", content_type)]
    handler.headers = headers(*pairs)
    handler._reject_before_body = Mock()
    assert not handler._require_json_content_type(time.monotonic())
    assert handler._reject_before_body.call_args.args[0] == 415


@pytest.mark.parametrize("content_type", ["application/json", 'Application/JSON; charset="utf-8"', 'application/json; note="a;b"'])
def test_companion_json_parameters_allowed(content_type):
    handler = object.__new__(CompanionHandler)
    handler.headers = headers(("Content-Type", content_type))
    assert handler._require_json_content_type(time.monotonic())












def test_companion_framing_rejects_before_session_creation():
    server = PowerPointCompanionServer(port=0, plugin_key="synthetic-key")
    server.start()
    try:
        reply = raw_reply(server.server_address(), b"POST /powerpoint/session HTTP/1.1\r\nHost: localhost\r\nOrigin: https://localhost:3002\r\nContent-Type: application/json\r\nContent-Length: 2\r\nContent-Length: 2\r\n\r\n{}")
        assert reply.startswith(b"HTTP/1.1 400 ")
        assert server.sessions.active_count == 0
    finally:
        server.stop()


def test_companion_unknown_target_is_not_log_content():
    handler = object.__new__(CompanionHandler)
    handler.path = "/powerpoint/target-marker?secret=query-marker"
    assert "target-marker" not in handler._route_template()
    assert "query-marker" not in handler._route_template()


@pytest.mark.parametrize("transport", ["companion"])
@pytest.mark.parametrize("case,status", [("duplicate", 400), ("oversized", 413), ("auth", 401)])
def test_expect_rejections_never_invite_a_body(transport, case, status, monkeypatch):
    if transport == "proxy":
        server = proxy.create_server("127.0.0.1", 0)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        address = server.server_address
        target = proxy.ROUTE
        if case == "auth":
            status = 403
        auth = "Origin: https://evil.example\r\n" if case == "auth" else ""
    else:
        server = PowerPointCompanionServer(port=0, plugin_key="synthetic-key")
        server.start()
        address = server.server_address()
        target = "/powerpoint/invoke"
        auth = "X-Hermes-Plugin-Key: synthetic-key\r\n" if case != "auth" else ""
    length = "2\r\nContent-Length: 2" if case == "duplicate" else ("9999999" if case == "oversized" else "2")
    try:
        reply = raw_reply(address, (f"POST {target} HTTP/1.1\r\nHost: localhost\r\nExpect: 100-continue\r\nContent-Type: application/json\r\nContent-Length: {length}\r\n" + auth + "\r\n").encode())
        assert reply.startswith(f"HTTP/1.1 {status} ".encode())
        assert b"100 Continue" not in reply
    finally:
        if transport == "proxy":
            server.shutdown()
            server.server_close()
            thread.join(2)
        else:
            server.stop()


def test_accepted_expect_sends_one_continue_at_body_boundary():
    server = PowerPointCompanionServer(port=0, plugin_key="synthetic-key")
    server.start()
    body = b'{"protocolVersion":"1.0","documentId":"deck","paneId":"pane","capabilities":[]}'
    try:
        with socket.create_connection(server.server_address(), timeout=2) as sock:
            sock.settimeout(2)
            sock.sendall((f"POST /powerpoint/session HTTP/1.1\r\nHost: localhost\r\nOrigin: https://localhost:3002\r\nExpect: 100-continue\r\nContent-Type: application/json\r\nContent-Length: {len(body)}\r\nConnection: close\r\n\r\n").encode())
            stream = sock.makefile("rb")
            assert stream.readline() == b"HTTP/1.1 100 Continue\r\n"
            assert stream.readline() == b"\r\n"
            sock.sendall(body)
            reply = stream.read()
            assert reply.startswith(b"HTTP/1.1 201 ")
            assert b"100 Continue" not in reply
            assert server.sessions.active_count == 1
            stream.close()
    finally:
        server.stop()
