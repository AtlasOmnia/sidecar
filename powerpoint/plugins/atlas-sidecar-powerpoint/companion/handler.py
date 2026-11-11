"""HTTP request handler for the PowerPoint companion RPC server.

Auth: dual system
  - Pane routes: exact Origin + X-Hermes-Session-Key + X-Hermes-Session-Id + X-Hermes-Pane-Id
  - Plugin routes: X-Hermes-Plugin-Key (constant-time comparison), no Origin needed

Body caps enforced before JSON parse. No secrets or content logged.
"""
from __future__ import annotations

import json
import logging
import socket
import threading
import time
import urllib.parse
from http.server import BaseHTTPRequestHandler
from typing import TYPE_CHECKING, Any

from .protocol import (
    CORS_HEADERS,
    CORS_METHODS,
    HOST_PREFIX,
    INVOKE_TIMEOUT_SECONDS,
    MAX_DOCUMENT_ID_LEN,
    MAX_INVOKE_BODY,
    MAX_JSON_DEPTH,
    MAX_PANE_ID_LEN,
    MAX_RESULT_BODY,
    MAX_SESSION_BODY,
    MAX_TOOL_NAME_LEN,
    OPTIONS_ROUTES,
    ORIGIN_ONLY_ROUTES,
    PANE_AUTH_ROUTES,
    PLUGIN_KEY_ROUTES,
    PROTOCOL_VERSION,
    REQUIRED_ORIGIN,
    SERVICE_HOST,
    SERVICE_NAME,
    SESSION_TTL_SECONDS,
    constant_time_compare,
    generate_uuid,
    validate_bounded_string,
    validate_uuid,
)
from .session import Ticket
from .request_boundaries import canonical_content_length, is_valid_json_content_type, json_depth_exceeds_limit
from .response_deadline import ResponseDeadlineMixin, ResponseWriteAborted

if TYPE_CHECKING:
    from .server import PowerPointCompanionServer


# Metadata-only logger (no bodies, keys, tool names, arguments, document content)
_md_log = logging.getLogger("powerpoint_companion")
_REQUEST_BODY_CAPS = {
    ("POST", "/powerpoint/session"): MAX_SESSION_BODY,
    ("POST", "/powerpoint/heartbeat"): MAX_SESSION_BODY,
    ("POST", "/powerpoint/invoke"): MAX_INVOKE_BODY,
    ("POST", "/powerpoint/result"): MAX_RESULT_BODY,
    ("POST", "/powerpoint/cancel"): MAX_SESSION_BODY,
}
_KNOWN_METHODS = frozenset({"DELETE", "GET", "HEAD", "OPTIONS", "PATCH", "POST", "PUT"})


def _metadata_log(method: str, route: str, status: int, elapsed_ms: float,
                  ticket_id: str = "", session_id: str = "") -> None:
    """Log only metadata; trusted bridge session IDs never enter logs."""
    _md_log.info(
        "method=%s route=%s status=%s elapsed=%.1f ticket=%s",
        method if method in _KNOWN_METHODS else "OTHER", route, status, elapsed_ms, ticket_id,
    )


class CompanionHandler(ResponseDeadlineMixin, BaseHTTPRequestHandler):
    """Handler wired to a specific PowerPointCompanionServer instance."""

    protocol_version = "HTTP/1.1"

    _server_ref: "PowerPointCompanionServer | None" = None

    @classmethod
    def with_server(cls, srv: "PowerPointCompanionServer") -> type:
        """Create a handler subclass bound to a server instance."""
        ref = srv

        class BoundHandler(cls):
            _server_ref = ref

        return BoundHandler

    # ------------------------------------------------------------------
    # Logging — metadata only
    # ------------------------------------------------------------------

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A002
        """Suppress default stderr logging — use metadata-only logger."""

    # ------------------------------------------------------------------
    # Absolute header deadline
    # ------------------------------------------------------------------

    _header_deadline_timer: threading.Timer | None = None

    def parse_request(self) -> bool:
        self._request_framing_ready = False
        self._expect_continue = False
        self._parsed_path = None
        self._parsed_query = {}
        self._auth_session = None
        if not super().parse_request():
            return False
        path, query = self._parse_url(self.path)
        self._parsed_path, self._parsed_query = path, query
        route_key = (self.command, path)
        # The shared Desk layer dispatches these routes after this companion's
        # canonical framing gate. Keep its bounded bodies and route recognition
        # here so parse_request does not reject them before Desk authorization.
        caps = dict(_REQUEST_BODY_CAPS)
        desk = getattr(self._server_ref, "desk", None)
        desk_routes = set()
        if desk is not None:
            from ..desk_bridge import MAX_CHAT_BODY, MAX_CONNECT_BODY
            desk_routes = {
                ("GET", f"/{desk.app}/hello"),
                ("POST", f"/{desk.app}/connect"),
                ("POST", f"/{desk.app}/connect/status"),
                ("POST", "/v1/chat/completions"),
            }
            caps.update({("POST", f"/{desk.app}/connect"): MAX_CONNECT_BODY,
                         ("POST", f"/{desk.app}/connect/status"): MAX_CONNECT_BODY,
                         ("POST", "/v1/chat/completions"): MAX_CHAT_BODY})
        cap = caps.get(route_key)
        if self.headers.get_all("Transfer-Encoding"):
            return self._reject_request_framing(411, "Transfer-Encoding not supported")
        length, error = canonical_content_length(self.headers.get_all("Content-Length"), cap)
        if error == "missing":
            if cap is not None:
                return self._reject_request_framing(411, "Length Required")
            length = 0
        elif error == "oversized":
            return self._reject_request_framing(413, "Payload Too Large")
        elif error is not None or length is None:
            return self._reject_request_framing(400, "Bad Request: invalid Content-Length")
        self._request_content_length = length
        if length and self.command in {"DELETE", "GET", "OPTIONS"}:
            return self._reject_request_framing(400, "Bad Request: unexpected request body")
        known = (route_key in desk_routes or route_key in ORIGIN_ONLY_ROUTES or route_key in PANE_AUTH_ROUTES
                 or route_key in PLUGIN_KEY_ROUTES
                 or (self.command == "OPTIONS" and path in OPTIONS_ROUTES))
        if length and not known:
            if self.command not in {"DELETE", "GET", "OPTIONS", "POST"}:
                self.close_connection = True
            else:
                return self._reject_request_framing(404, "Not Found")
        self._request_framing_ready = True
        self._cancel_header_deadline()
        return True

    def _reject_request_framing(self, status: int, message: str) -> bool:
        self._cancel_header_deadline()
        self.close_connection = True
        self._send_error(status, message, start=time.monotonic(), extra_headers={"Connection": "close"})
        return False

    def handle_expect_100(self) -> bool:
        # Defer until canonical framing, route, auth and media gates pass.
        self._expect_continue = True
        return True

    def _accept_expect_continue(self) -> None:
        if getattr(self, "_expect_continue", False):
            self._expect_continue = False
            self.send_response_only(100)
            self.end_headers()
            self._flush_response()
            self._finish_response_write()  # Interim reply has its own budget.

    def handle_one_request(self) -> None:
        """Bound total header-read time, including trickled bytes."""
        self._parsed_path = None
        self._parsed_query = {}
        self._auth_session = None
        srv = self._server_ref
        timeout = srv.request_timeout_seconds if srv is not None else 10.0
        timer = threading.Timer(timeout, self._expire_header_read)
        timer.daemon = True
        self._header_deadline_timer = timer
        timer.start()
        self._handling_request = True
        self._response_committed = False
        try:
            try:
                super().handle_one_request()
            except (ResponseWriteAborted, OSError):
                self._abort_response_write()
                _md_log.info("method=UNKNOWN route=REQUEST_PARSE error=ClientDisconnect")
        finally:
            self._finish_response_write()
            self._handling_request = False
            timer.cancel()
            if self._header_deadline_timer is timer:
                self._header_deadline_timer = None

    def _cancel_header_deadline(self) -> None:
        timer = self._header_deadline_timer
        if timer is not None:
            timer.cancel()
            self._header_deadline_timer = None

    def _expire_header_read(self) -> None:
        self.close_connection = True
        try:
            self.connection.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass

    # ------------------------------------------------------------------
    # HTTP method dispatchers
    # ------------------------------------------------------------------

    def do_OPTIONS(self) -> None:  # noqa: N802
        self._cancel_header_deadline()
        self._handle_options()

    def do_GET(self) -> None:  # noqa: N802
        self._cancel_header_deadline()
        self._dispatch("GET")

    def do_POST(self) -> None:  # noqa: N802
        self._cancel_header_deadline()
        self._dispatch("POST")

    def do_DELETE(self) -> None:  # noqa: N802
        self._cancel_header_deadline()
        self._dispatch("DELETE")

    # ------------------------------------------------------------------
    # URL parsing
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_url(raw_path: str) -> tuple[str, dict[str, list[str]]]:
        """Parse path+query via urlsplit. Returns (path, query_dict)."""
        try:
            parsed = urllib.parse.urlsplit(raw_path)
        except ValueError:
            return "/", {}
        path = parsed.path.rstrip("/") or "/"
        query = urllib.parse.parse_qs(parsed.query)
        return path, query

    # ------------------------------------------------------------------
    # Dispatch + auth
    # ------------------------------------------------------------------

    def _dispatch(self, method: str) -> None:
        """Route request through auth layer then to handler."""
        start = time.monotonic()

        # Parse path and query
        path, query = self._parse_url(self.path)
        self._parsed_path = path
        self._parsed_query = query

        # Check if it's under our host prefix
        if not path.startswith(HOST_PREFIX):
            self._reject_before_body(404, "Not Found", start)
            return

        # Sweep stale sessions, leases, and invoke deadlines on every request.
        srv = self._server_ref
        if srv is not None:
            srv.sessions.sweep()

        route_key = (method, path)

        # Determine auth type
        if route_key in ORIGIN_ONLY_ROUTES:
            if not self._check_origin():
                self._reject_before_body(403, "Forbidden", start)
                return
        elif route_key in PANE_AUTH_ROUTES:
            if not self._check_origin():
                self._reject_before_body(403, "Forbidden", start)
                return
            session = self._check_session_auth()
            if session is None:
                self._reject_before_body(401, "Unauthorized", start)
                return
            self._auth_session = session
        elif route_key in PLUGIN_KEY_ROUTES:
            if not self._check_plugin_key():
                self._reject_before_body(401, "Unauthorized", start)
                return
        else:
            self._reject_before_body(404, "Not Found", start)
            return

        # Retain crop/W0 auth verdict precedence, then enforce JSON media types.
        if route_key in _REQUEST_BODY_CAPS and not self._require_json_content_type(start):
            return

        # Route to handler
        try:
            if route_key == ("POST", "/powerpoint/session"):
                self._handle_session_create(start)
            elif route_key == ("DELETE", "/powerpoint/session"):
                self._handle_session_delete(start)
            elif route_key == ("GET", "/powerpoint/session"):
                self._handle_session_status(start)
            elif route_key == ("POST", "/powerpoint/heartbeat"):
                self._handle_heartbeat(start)
            elif route_key == ("GET", "/powerpoint/pending"):
                self._handle_pending(start)
            elif route_key == ("POST", "/powerpoint/invoke"):
                self._handle_invoke(start)
            elif route_key == ("POST", "/powerpoint/result"):
                self._handle_result(start)
            elif route_key == ("POST", "/powerpoint/cancel"):
                self._handle_cancel(start)
            elif route_key == ("GET", "/powerpoint/capabilities"):
                self._handle_capabilities(start)
            elif route_key == ("GET", "/powerpoint/metrics"):
                self._handle_metrics(start)
            elif route_key == ("GET", "/powerpoint/health"):
                self._handle_health(start)
            else:
                self._send_error(404, "Not Found", start=start)
        except (ResponseWriteAborted, BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            self._abort_response_write()
            _md_log.info(
                "method=%s route=%s error=ClientDisconnect",
                method, self._route_template(),
            )
        except Exception as exc:
            # Log only exception class + fixed metadata — never traceback, body, or args
            _md_log.error(
                "method=%s route=%s error=%s",
                method, self._route_template(), type(exc).__name__,
            )
            if getattr(self, "_response_committed", False):
                self._abort_response_write()
                return
            try:
                self._send_error(500, "Internal Server Error", start=start)
            except OSError:
                self.close_connection = True

    # ------------------------------------------------------------------
    # Auth checks
    # ------------------------------------------------------------------

    def _reject_before_body(self, status: int, message: str, start: float) -> None:
        """Discard only safely framed, bounded bytes before an early rejection.

        Closing with an unread body can reset a Windows client before it receives
        the error. Never parse unauthenticated JSON or wait indefinitely, and
        never reuse a rejected connection with leftover body bytes.
        """
        self.close_connection = True
        srv = self._server_ref
        budget = min(srv.request_timeout_seconds if srv is not None else 10.0, 0.25)
        deadline = time.monotonic() + budget
        lengths = self.headers.get_all("Content-Length", [])
        try:
            self.connection.settimeout(budget)
            if (self.headers.get("Transfer-Encoding") is None and len(lengths) == 1
                    and lengths[0].isascii() and lengths[0].isdigit()):
                length = int(lengths[0])
                if length <= 64 * 1024:
                    while length:
                        remaining_time = deadline - time.monotonic()
                        if remaining_time <= 0:
                            break
                        self.connection.settimeout(remaining_time)
                        chunk = self.rfile.read1(min(length, 4096))
                        if not chunk:
                            break
                        length -= len(chunk)
            self.connection.settimeout(budget)
        except (OSError, ValueError):
            # Preserve the original auth/route verdict on truncation or timeout.
            pass
        self._send_error(status, message, start=start, extra_headers={"Connection": "close"})

    def _drain_rejected_body(self, max_bytes: int, start: float) -> bool:
        """Best-effort bounded drain for a request rejected before parsing (desk bridge contract)."""
        if self.headers.get_all("Transfer-Encoding"):
            return False
        length, error = canonical_content_length(self.headers.get_all("Content-Length"), max_bytes)
        if error == "missing":
            return True
        if error is not None or length is None:
            return False
        srv = self._server_ref
        timeout_seconds = srv.request_timeout_seconds if srv is not None else 10.0
        deadline = start + timeout_seconds
        remaining = length
        try:
            while remaining:
                remaining_time = deadline - time.monotonic()
                if remaining_time <= 0:
                    return False
                self.connection.settimeout(remaining_time)
                chunk = self.rfile.read(min(64 * 1024, remaining))
                if not chunk:
                    return False
                remaining -= len(chunk)
            self.connection.settimeout(timeout_seconds)
            return True
        except (OSError, socket.timeout):
            return False

    def _check_origin(self) -> bool:
        """Check Origin header matches exactly https://localhost:3002."""
        origin = self.headers.get("Origin", "")
        return constant_time_compare(origin, REQUIRED_ORIGIN)

    def _require_json_content_type(self, start: float) -> bool:
        values = self.headers.get_all("Content-Type", [])
        if len(values) == 1 and is_valid_json_content_type(values[0]):
            return True
        self._reject_before_body(415, "Unsupported Media Type", start)
        return False

    def _check_plugin_key(self) -> bool:
        """Check X-Hermes-Plugin-Key with constant-time comparison.

        Missing key and wrong key both return False — no timing/status leak.
        """
        srv = self._server_ref
        if srv is None or srv.plugin_key is None:
            return False
        provided = self.headers.get("X-Hermes-Plugin-Key", "")
        if not provided:
            constant_time_compare("__missing__", srv.plugin_key)
            return False
        return constant_time_compare(provided, srv.plugin_key)

    def _check_session_auth(self) -> Any:
        """Verify X-Hermes-Session-Id, X-Hermes-Session-Key, X-Hermes-Pane-Id.

        Returns the Session object if authenticated, None otherwise.
        All failure paths use constant-time comparisons to prevent leaks.
        """
        srv = self._server_ref
        if srv is None:
            return None

        session_id = self.headers.get("X-Hermes-Session-Id", "")
        session_key = self.headers.get("X-Hermes-Session-Key", "")
        pane_id = self.headers.get("X-Hermes-Pane-Id", "")

        session = srv.sessions.get(session_id)
        expected_session_id = session.session_id if session is not None else "__dummy__"
        expected_session_key = session.session_key if session is not None else "__dummy__"
        expected_pane_id = session.pane_id if session is not None else "__dummy__"

        session_id_ok = constant_time_compare(session_id, expected_session_id)
        session_key_ok = constant_time_compare(session_key, expected_session_key)
        pane_id_ok = constant_time_compare(pane_id, expected_pane_id)

        if session is None or not (session_id_ok and session_key_ok and pane_id_ok):
            return None

        return session

    # ------------------------------------------------------------------
    # Body reading with caps
    # ------------------------------------------------------------------

    def _read_body(self, max_bytes: int, start: float) -> bytes | None:
        """Read request body with size cap enforced before full read.

        Returns body bytes or None (error already sent).
        Sends Connection: close when rejecting to avoid desynchronization.
        """
        # Direct callers still validate every framing field; wire decisions
        # happen before route/auth side effects in parse_request.
        if self.headers.get_all("Transfer-Encoding"):
            self.close_connection = True
            self._send_error(411, "Transfer-Encoding not supported",
                             start=start, extra_headers={"Connection": "close"})
            return None

        length, error = canonical_content_length(self.headers.get_all("Content-Length"), max_bytes)
        if error == "missing":
            self.close_connection = True
            self._send_error(411, "Length Required", start=start,
                             extra_headers={"Connection": "close"})
            return None

        if error == "invalid" or (length is None and error != "oversized"):
            self.close_connection = True
            self._send_error(400, "Bad Request: invalid Content-Length", start=start,
                             extra_headers={"Connection": "close"})
            return None

        if error == "oversized":
            self.close_connection = True
            self._send_error(413, "Payload Too Large", start=start,
                             extra_headers={"Connection": "close"})
            return None

        self._accept_expect_continue()
        srv = self._server_ref
        timeout_seconds = srv.request_timeout_seconds if srv is not None else 10.0
        deadline = start + timeout_seconds
        chunks: list[bytes] = []
        remaining_bytes = length
        try:
            while remaining_bytes:
                remaining_time = deadline - time.monotonic()
                if remaining_time <= 0:
                    raise socket.timeout("absolute body deadline exceeded")
                self.connection.settimeout(remaining_time)
                chunk = self.rfile.read(min(64 * 1024, remaining_bytes))
                if not chunk:
                    self._send_error(400, "Bad Request: short body", start=start,
                                     extra_headers={"Connection": "close"})
                    return None
                chunks.append(chunk)
                remaining_bytes -= len(chunk)
            self.connection.settimeout(timeout_seconds)
            return b"".join(chunks)
        except socket.timeout:
            self.close_connection = True
            self._send_error(408, "Request Timeout: body deadline exceeded", start=start,
                             extra_headers={"Connection": "close"})
            return None
        except Exception:
            self.close_connection = True
            self._send_error(400, "Bad Request: failed to read body", start=start,
                             extra_headers={"Connection": "close"})
            return None

    def _parse_json(self, body: bytes, start: float) -> Any | None:
        """Parse JSON body. Returns parsed value or None (error already sent)."""
        try:
            if json_depth_exceeds_limit(body, MAX_JSON_DEPTH):
                raise ValueError("JSON depth exceeded")
            return json.loads(body.decode("utf-8"))
        except (ValueError, UnicodeDecodeError, RecursionError):
            self.close_connection = True
            self._send_error(400, "Bad Request: invalid JSON", start=start,
                             extra_headers={"Connection": "close"})
            return None

    # ------------------------------------------------------------------
    # CORS helpers
    # ------------------------------------------------------------------

    def _send_cors_headers(self) -> None:
        """Send CORS headers for pane routes with valid Origin."""
        origin = self.headers.get("Origin", "")
        if constant_time_compare(origin, REQUIRED_ORIGIN):
            self.send_header("Access-Control-Allow-Origin", REQUIRED_ORIGIN)
        self.send_header("Access-Control-Allow-Methods", CORS_METHODS)
        self.send_header("Access-Control-Allow-Headers", ", ".join(CORS_HEADERS))
        self.send_header("Access-Control-Max-Age", "600")
        self.send_header("Vary", "Origin")

    # ------------------------------------------------------------------
    # Unified response helpers
    # ------------------------------------------------------------------

    def _emit_response(self, status: int, body: bytes | None, start: float,
                       ticket_id: str = "", session_id: str = "",
                       extra_headers: dict[str, str] | None = None,
                       preflight_cors: bool | None = None) -> None:
        headers = {}
        tokens = set()
        for key, value in (extra_headers or {}).items():
            if key.lower() == "connection":
                tokens.update(token.strip().lower() for token in value.split(","))
            else:
                headers[key] = value
        path = getattr(self, "_parsed_path", None)
        if path is None:
            path, _ = self._parse_url(getattr(self, "path", "/"))
        if (getattr(self, "close_connection", False) or "close" in tokens
                or (getattr(self, "command", "") == "GET" and path == "/powerpoint/pending")):
            self.close_connection = True
            headers["Connection"] = "close"
        elif tokens:
            headers["Connection"] = ", ".join(sorted(tokens))
        self._begin_response_write()
        try:
            self.send_response(status)
            if body is not None:
                self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body) if body else 0))
            if preflight_cors is True:
                self._send_cors_headers()
            elif preflight_cors is None and self._is_pane_route():
                self._add_cors_response_headers()
            for key, value in headers.items():
                self.send_header(key, value)
            self.end_headers()
            if body and getattr(self, "command", "") != "HEAD":
                self._write_response(body)
            self._flush_response()
            _metadata_log(getattr(self, "command", "OTHER"), self._route_template(), status,
                          (time.monotonic() - start) * 1000, ticket_id, session_id)
        except ResponseWriteAborted:
            self._abort_response_write()
        finally:
            if not getattr(self, "_handling_request", False):
                self._finish_response_write()

    def _send_error(self, status: int, message: str = "",
                    start: float = 0.0, ticket_id: str = "",
                    session_id: str = "",
                    extra_headers: dict[str, str] | None = None) -> None:
        """Send an error response with metadata-only logging."""
        body = json.dumps({"ok": False, "error": message, "category": self._error_category(status)}).encode()
        self._emit_response(status, body, start, ticket_id, session_id, extra_headers)

    def _send_json(self, status: int, data: Any,
                   start: float = 0.0, ticket_id: str = "",
                   session_id: str = "") -> None:
        """Send a JSON response."""
        body = json.dumps(data).encode()
        self._emit_response(status, body, start, ticket_id, session_id)

    def _send_empty(self, status: int, start: float = 0.0,
                    ticket_id: str = "", session_id: str = "") -> None:
        """Send an empty response, including valid pane CORS metadata."""
        self._emit_response(status, None, start, ticket_id, session_id)

    def _is_pane_route(self) -> bool:
        """Check if current route is a pane route (needs CORS headers)."""
        path = getattr(self, "_parsed_path", None)
        if path is None:
            path, _ = self._parse_url(getattr(self, "path", "/"))
        route_key = (self.command, path)
        return route_key in PANE_AUTH_ROUTES or route_key in ORIGIN_ONLY_ROUTES

    def _add_cors_response_headers(self) -> None:
        """Add CORS response headers only when Origin is valid."""
        origin = self.headers.get("Origin", "")
        if constant_time_compare(origin, REQUIRED_ORIGIN):
            self.send_header("Access-Control-Allow-Origin", REQUIRED_ORIGIN)
        self.send_header("Vary", "Origin")

    def _route_template(self) -> str:
        """Return only allowlisted routes, never attacker-controlled path text."""
        path = getattr(self, "_parsed_path", None)
        if path is None:
            path, _ = self._parse_url(getattr(self, "path", "/"))
        known = {route for _, route in ORIGIN_ONLY_ROUTES | PANE_AUTH_ROUTES | PLUGIN_KEY_ROUTES}
        return path if path in known else "UNKNOWN"

    @staticmethod
    def _error_category(status: int) -> str:
        """Map HTTP status to error category."""
        if status == 400:
            return "Validation"
        if status == 401:
            return "Auth"
        if status == 403:
            return "Auth"
        if status == 404:
            return "NotFound"
        if status == 409:
            return "Conflict"
        if status in {413, 415}:
            return "Validation"
        if status == 429:
            return "Capacity"
        if status == 499:
            return "Cancellation"
        if status == 503:
            return "Capacity"
        if status == 504:
            return "Timeout"
        return "Internal"

    # ------------------------------------------------------------------
    # OPTIONS handler
    # ------------------------------------------------------------------

    def _handle_options(self, start: float = 0.0) -> None:
        """Handle CORS preflight for pane routes only.

        Plugin route OPTIONS returns exact 405.
        """
        path, _ = self._parse_url(self.path)

        # Check if this is a plugin route first
        route_key = ("OPTIONS", path)
        if route_key in PLUGIN_KEY_ROUTES or (
            path.startswith(HOST_PREFIX + "/") and path not in OPTIONS_ROUTES
        ):
            self._send_error(405, "Method Not Allowed", start=start)
            return

        if path not in OPTIONS_ROUTES:
            self._send_error(405, "Method Not Allowed", start=start)
            return

        # The desk bridge overrides this check with the configured add-in origin; the
        # companion's development-origin constant must not gate production preflights.
        if not self._check_origin():
            self._emit_response(204, None, start, preflight_cors=False)
            return

        self._emit_response(204, None, start, preflight_cors=True)

    # ------------------------------------------------------------------
    # Route handlers
    # ------------------------------------------------------------------

    def _handle_session_create(self, start: float) -> None:
        """POST /powerpoint/session — create a new session (pane, origin-gated)."""
        body = self._read_body(MAX_SESSION_BODY, start)
        if body is None:
            return

        data = self._parse_json(body, start)
        if data is None:
            return

        if not isinstance(data, dict):
            self._send_error(400, "Bad Request: expected JSON object", start=start)
            return

        # Canonical protocolVersion only — reject "version" legacy alias
        version = data.get("protocolVersion", "")
        if version != PROTOCOL_VERSION:
            self._send_error(400, f"Unsupported protocol version: {version}", start=start)
            return

        document_id = data.get("documentId", "")
        pane_id = data.get("paneId", "")

        # Validate bounded strings
        doc_err = validate_bounded_string(document_id, MAX_DOCUMENT_ID_LEN, "documentId")
        if doc_err:
            self._send_error(400, doc_err, start=start)
            return

        pane_err = validate_bounded_string(pane_id, MAX_PANE_ID_LEN, "paneId")
        if pane_err:
            self._send_error(400, pane_err, start=start)
            return

        srv = self._server_ref
        if srv is None:
            self._send_error(500, "Server not available", start=start)
            return

        from .protocol import generate_session_key
        session_key = generate_session_key()

        try:
            session = srv.sessions.create(document_id, pane_id, session_key)
        except RuntimeError:
            self._send_error(429, "Session capacity reached", start=start)
            return

        self._send_json(201, {
            "ok": True,
            "sessionId": session.session_id,
            "sessionKey": session_key,
            "protocolVersion": PROTOCOL_VERSION,
            "expiresInSeconds": SESSION_TTL_SECONDS,
        }, start=start, session_id=session.session_id)

    def _handle_session_delete(self, start: float) -> None:
        """DELETE /powerpoint/session — unregister session (pane auth)."""
        session = getattr(self, "_auth_session", None)
        if session is None:
            self._send_error(401, "Unauthorized", start=start)
            return

        srv = self._server_ref
        if srv:
            srv.sessions.delete_for_pane(session.session_id, session.pane_id)

        self._send_empty(204, start=start, session_id=session.session_id)

    def _handle_session_status(self, start: float) -> None:
        """GET /powerpoint/session — get session status (plugin-key auth).

        Returns status without document content or pane identity.
        """
        srv = self._server_ref
        if srv is None:
            self._send_error(500, "Server not available", start=start)
            return

        # Get session_id from parsed query or header
        session_id = (self._get_query_param("sessionId")
                      or self.headers.get("X-Hermes-Session-Id", ""))

        if session_id:
            session = srv.sessions.get(session_id)
            if session is None:
                self._send_error(404, "Session not found", start=start)
                return

            pending_count = 1 if session.has_ticket() else 0
            self._send_json(200, {
                "active": True,
                "sessionId": session.session_id,
                "pending": pending_count,
            }, start=start, session_id=session_id)
        else:
            self._send_json(200, {
                "active": srv.sessions.active_count > 0,
                "sessions": srv.sessions.active_count,
            }, start=start)

    def _handle_heartbeat(self, start: float) -> None:
        """POST /powerpoint/heartbeat — extend session TTL (pane auth).

        Consumes the request body even though heartbeat semantics are stateless —
        required for same-socket keep-alive so the next request does not
        encounter a stale body on the connection.
        """
        session = getattr(self, "_auth_session", None)
        if session is None:
            self._send_error(401, "Unauthorized", start=start)
            return

        # Consume request body for keep-alive safety. Fail closed when framing,
        # size, timeout, or truncation checks reject the body.
        body = self._read_body(MAX_SESSION_BODY, start)
        if body is None:
            return

        session.heartbeat()
        self._send_json(200, {"ok": True}, start=start, session_id=session.session_id)

    def _handle_pending(self, start: float) -> None:
        """GET /powerpoint/pending — claim oldest pending ticket (pane auth)."""
        session = getattr(self, "_auth_session", None)
        if session is None:
            self._send_error(401, "Unauthorized", start=start)
            return

        ticket = session.get_pending_ticket()
        if ticket is None:
            self._send_empty(204, start=start, session_id=session.session_id)
            return

        if ticket.claim(session.pane_id):
            self._send_json(200, {
                "ticketId": ticket.ticket_id,
                "requestId": ticket.request_id,
                "tool": ticket.tool,
                "arguments": ticket.arguments,
            }, start=start, ticket_id=ticket.ticket_id, session_id=session.session_id)
        else:
            self._send_empty(204, start=start, session_id=session.session_id)

    def _handle_invoke(self, start: float) -> None:
        """POST /powerpoint/invoke — mint a ticket and block for result (plugin-key auth)."""
        body = self._read_body(MAX_INVOKE_BODY, start)
        if body is None:
            return

        data = self._parse_json(body, start)
        if data is None:
            return

        if not isinstance(data, dict):
            self._send_error(400, "Bad Request: expected JSON object", start=start)
            return

        # Version check
        version = data.get("protocolVersion", "")
        if version != PROTOCOL_VERSION:
            self._send_error(400, f"Unsupported protocol version: {version}", start=start)
            return

        session_id = data.get("sessionId", "")
        request_id = data.get("requestId", "")
        tool = data.get("tool", "")
        arguments = data.get("arguments", {})

        if not session_id or not validate_uuid(session_id):
            self._send_error(400, "Bad Request: valid sessionId UUID required", start=start)
            return

        if not request_id or not validate_uuid(request_id):
            self._send_error(400, "Bad Request: valid requestId UUID required", start=start)
            return

        if not tool or not isinstance(tool, str):
            self._send_error(400, "Bad Request: tool name required", start=start)
            return

        tool_err = validate_bounded_string(tool, MAX_TOOL_NAME_LEN, "tool")
        if tool_err:
            self._send_error(400, tool_err, start=start)
            return

        if not isinstance(arguments, dict):
            self._send_error(400, "Bad Request: arguments must be object", start=start)
            return

        srv = self._server_ref
        if srv is None:
            self._send_error(500, "Server not available", start=start)
            return

        session = srv.sessions.get(session_id)
        if session is None:
            self._send_error(404, "Session not found", start=start,
                             session_id=session_id)
            return

        candidate = Ticket(
            ticket_id=generate_uuid(),
            session_id=session_id,
            request_id=request_id,
            tool=tool,
            arguments=arguments,
            created_at=srv.now(),
            invoke_timeout_seconds=srv.invoke_timeout_seconds,
            _clock=srv._clock,
        )

        disposition, ticket = session.begin_ticket(candidate)
        if disposition == "conflict":
            self._send_error(409, "requestId reused with different payload",
                             start=start, session_id=session_id)
            return
        if disposition == "busy":
            self._send_error(429, "A ticket is already pending for this session",
                             start=start, session_id=session_id)
            return

        # A duplicate request reattaches to the original ticket. It never
        # creates a second Office mutation.
        signaled = ticket.wait_for_result(timeout=srv.invoke_timeout_seconds)
        if not signaled and not ticket.is_terminal():
            ticket.timeout()

        if ticket.state == "completed":
            self._send_json(200, {
                "ok": True,
                "ticketId": ticket.ticket_id,
                "requestId": request_id,
                "data": ticket.result_payload,
            }, start=start, ticket_id=ticket.ticket_id, session_id=session_id)
        elif ticket.state == "timeout":
            self._send_error(504, "Invoke timed out", start=start,
                             ticket_id=ticket.ticket_id, session_id=session_id)
        elif ticket.state == "cancelled":
            self._send_error(499, "Cancelled", start=start,
                             ticket_id=ticket.ticket_id, session_id=session_id)
        else:
            self._send_error(500, "Unknown ticket state", start=start,
                             ticket_id=ticket.ticket_id, session_id=session_id)

    def _handle_result(self, start: float) -> None:
        """POST /powerpoint/result — submit result for a claimed ticket (pane auth).

        Preserves any JSON value (object/array/string/number/bool/null).
        Duplicate same canonical payload is idempotent (200); conflicting is 409.
        """
        session = getattr(self, "_auth_session", None)
        if session is None:
            self._send_error(401, "Unauthorized", start=start)
            return

        body = self._read_body(MAX_RESULT_BODY, start)
        if body is None:
            return

        data = self._parse_json(body, start)
        if data is None:
            return

        if not isinstance(data, dict):
            self._send_error(400, "Bad Request: expected JSON object", start=start)
            return

        ticket_id = data.get("ticketId", "")
        if not ticket_id or not validate_uuid(ticket_id):
            self._send_error(400, "Bad Request: valid ticketId UUID required",
                             start=start, session_id=session.session_id)
            return

        ticket = session.find_ticket(ticket_id)
        if ticket is None:
            self._send_error(404, "No ticket found", start=start,
                             session_id=session.session_id)
            return

        if ticket.ticket_id != ticket_id:
            self._send_error(409, "Ticket ID mismatch", start=start,
                             session_id=session.session_id)
            return

        if ticket.state == "completed":
            # Idempotent: same completed payload → 200; different → 409
            result_payload = data.get("result")
            # Canonical comparison via JSON round-trip for structural equality
            if _json_canonical_equal(result_payload, ticket.result_payload):
                self._send_json(200, {"ok": True, "ticketId": ticket_id},
                                start=start, ticket_id=ticket_id,
                                session_id=session.session_id)
            else:
                self._send_error(409, "Duplicate result submission with conflicting data",
                                 start=start, ticket_id=ticket_id,
                                 session_id=session.session_id)
            return

        result_payload = data.get("result")
        disposition = ticket.submit_result(result_payload, session.pane_id)
        if disposition == "wrong_pane":
            self._send_error(409, "Ticket claimed by different pane",
                             start=start, ticket_id=ticket_id,
                             session_id=session.session_id)
            return
        if disposition != "accepted":
            state = disposition.removeprefix("invalid:")
            self._send_error(409, f"Ticket in state {state}, cannot accept result",
                             start=start, ticket_id=ticket_id,
                             session_id=session.session_id)
            return

        session.trim_history()
        self._send_json(200, {"ok": True, "ticketId": ticket_id},
                        start=start, ticket_id=ticket_id,
                        session_id=session.session_id)

    def _handle_cancel(self, start: float) -> None:
        """POST /powerpoint/cancel — cancel a ticket (plugin-key auth)."""
        srv = self._server_ref
        if srv is None:
            self._send_error(500, "Server not available", start=start)
            return

        body = self._read_body(MAX_SESSION_BODY, start)
        if body is None:
            return

        data = self._parse_json(body, start)
        if data is None:
            return

        if not isinstance(data, dict):
            self._send_error(400, "Bad Request: expected JSON object", start=start)
            return

        session_id = data.get("sessionId", "")
        ticket_id = data.get("ticketId", "")

        if not session_id or not validate_uuid(session_id):
            self._send_error(400, "Bad Request: valid sessionId UUID required", start=start)
            return

        session = srv.sessions.get(session_id)
        if session is None:
            self._send_error(404, "Session not found", start=start)
            return

        if ticket_id:
            if not validate_uuid(ticket_id):
                self._send_error(400, "Bad Request: valid ticketId UUID required", start=start)
                return
            ticket = session.find_ticket(ticket_id)
        else:
            ticket = session.get_ticket()

        if ticket is None:
            self._send_error(404, "No ticket found", start=start,
                             session_id=session_id)
            return

        if ticket_id and ticket.ticket_id != ticket_id:
            self._send_error(409, "Ticket ID mismatch", start=start,
                             session_id=session_id)
            return

        ticket.cancel()
        self._send_json(200, {"ok": True, "ticketId": ticket.ticket_id},
                        start=start, ticket_id=ticket.ticket_id,
                        session_id=session_id)

    def _handle_capabilities(self, start: float) -> None:
        """GET /powerpoint/capabilities — session runtime capability metadata (plugin-key auth).

        No document content or identity revealed.
        """
        self._send_json(200, {
            "protocolVersion": PROTOCOL_VERSION,
            "hostPrefix": HOST_PREFIX,
            "sessionTtl": SESSION_TTL_SECONDS,
            "heartbeatInterval": 15,
            "claimLease": 20,
            "invokeTimeout": INVOKE_TIMEOUT_SECONDS,
            "maxInvokeBody": MAX_INVOKE_BODY,
            "maxResultBody": MAX_RESULT_BODY,
        }, start=start)

    def _handle_metrics(self, start: float) -> None:
        """GET /powerpoint/metrics — counts only, no content (plugin-key auth)."""
        srv = self._server_ref
        if srv is None:
            self._send_error(500, "Server not available", start=start)
            return

        metrics = srv.sessions.get_metrics()
        self._send_json(200, metrics, start=start)

    def _handle_health(self, start: float) -> None:
        """GET /powerpoint/health — basic health check (plugin-key auth)."""
        srv = self._server_ref
        self._send_json(200, {
            "ok": True,
            "service": SERVICE_NAME,
            "host": SERVICE_HOST,
            "protocolVersion": PROTOCOL_VERSION,
            "uptime": srv.uptime if srv else 0,
        }, start=start)

    def _get_query_param(self, name: str) -> str | None:
        """Extract a query parameter from the parsed query dict."""
        query = getattr(self, "_parsed_query", None)
        if query is None:
            _, query = self._parse_url(self.path)
        values = query.get(name)
        if values:
            return values[0]
        return None


def _json_canonical_equal(a: Any, b: Any) -> bool:
    """Compare two arbitrary JSON values for structural equality via canonical serialization."""
    return json.dumps(a, sort_keys=True, separators=(",", ":")) == \
           json.dumps(b, sort_keys=True, separators=(",", ":"))
