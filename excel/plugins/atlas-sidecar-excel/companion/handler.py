"""HTTP request handler for the Excel companion RPC server.

Auth: dual system
  - Pane routes: exact Origin + X-Hermes-Session-Key + X-Hermes-Session-Id + X-Hermes-Pane-Id
  - Plugin routes: X-Hermes-Plugin-Key (constant-time comparison), no Origin needed

Body caps enforced before JSON parse. No secrets or content logged.
"""
from __future__ import annotations

import json
import logging
import re
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
    PLAN_PHASES,
    PLAN_TRUSTED_KEYS,
    PLUGIN_KEY_ROUTES,
    PROTOCOL_VERSION,
    REQUIRED_ORIGIN,
    WORKBOOK_PLAN_VERSION,
    SESSION_TTL_SECONDS,
    constant_time_compare,
    generate_uuid,
    validate_bounded_string,
    validate_uuid,
)
from .session import Ticket

if TYPE_CHECKING:
    from .server import ExcelCompanionServer


# Metadata-only logger (no bodies, keys, tool names, arguments, document content)
_md_log = logging.getLogger("excel_companion")
_PLAN_PLUGIN_TOOL = "execute_plan"
_PLAN_PANE_TOOL = "excel_execute_plan"
_FINGERPRINT_RE = re.compile(r"^[0-9a-f]{64}$")
# R4g-2b (R4g-2a lane-B A2): plan turnKey / op ids are canonical lowercase hyphenated UUIDs (versions 1-5, RFC
# variant; so never nil), the same spelling the pane parser accepts. Use with fullmatch.
_CANONICAL_UUID_RE = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}")
# Content-free plan invoke outcomes (R2d-2). Never include workbook content.
_PLAN_UNCERTAIN_ERROR = "Workbook plan execution outcome is uncertain"
_PLAN_UNCERTAIN_CATEGORY = "ExecutionUncertain"
# Content-free rejection of a new plan op once the session's plan fences are full (R3a).
_PLAN_FENCE_CAPACITY_ERROR = "Workbook plan capacity reached for this session"


def _metadata_log(method: str, route: str, status: int, elapsed_ms: float,
                  ticket_id: str = "", session_id: str = "") -> None:
    """Log only metadata: timestamp, method, route template, status, elapsed, IDs."""
    _md_log.info(
        "method=%s route=%s status=%s elapsed=%.1f ticket=%s session=%s",
        method, route, status, elapsed_ms, ticket_id, session_id,
    )


class CompanionHandler(BaseHTTPRequestHandler):
    """Handler wired to a specific ExcelCompanionServer instance."""

    protocol_version = "HTTP/1.1"

    _server_ref: "ExcelCompanionServer | None" = None

    @classmethod
    def with_server(cls, srv: "ExcelCompanionServer") -> type:
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

    def handle_one_request(self) -> None:
        """Bound total header-read time, including trickled bytes."""
        srv = self._server_ref
        timeout = srv.request_timeout_seconds if srv is not None else 10.0
        timer = threading.Timer(timeout, self._expire_header_read)
        timer.daemon = True
        self._header_deadline_timer = timer
        timer.start()
        try:
            try:
                super().handle_one_request()
            except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
                self.close_connection = True
                _md_log.info("method=UNKNOWN route=REQUEST_PARSE error=ClientDisconnect")
        finally:
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
        parsed = urllib.parse.urlsplit(raw_path)
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
            self._send_error(404, "Not Found", start=start)
            return

        # Sweep stale sessions, leases, and invoke deadlines on every request.
        srv = self._server_ref
        if srv is not None:
            srv.sessions.sweep()

        route_key = (method, path)

        # Determine auth type
        if route_key in ORIGIN_ONLY_ROUTES:
            if not self._check_origin():
                self._send_error(403, "Forbidden", start=start)
                return
        elif route_key in PANE_AUTH_ROUTES:
            if not self._check_origin():
                self._send_error(403, "Forbidden", start=start)
                return
            session = self._check_session_auth()
            if session is None:
                self._send_error(401, "Unauthorized", start=start)
                return
            self._auth_session = session
        elif route_key in PLUGIN_KEY_ROUTES:
            if not self._check_plugin_key():
                self._send_error(401, "Unauthorized", start=start)
                return
        else:
            self._send_error(404, "Not Found", start=start)
            return

        # Route to handler
        try:
            if route_key == ("POST", "/excel/session"):
                self._handle_session_create(start)
            elif route_key == ("DELETE", "/excel/session"):
                self._handle_session_delete(start)
            elif route_key == ("GET", "/excel/session"):
                self._handle_session_status(start)
            elif route_key == ("POST", "/excel/heartbeat"):
                self._handle_heartbeat(start)
            elif route_key == ("GET", "/excel/pending"):
                self._handle_pending(start)
            elif route_key == ("POST", "/excel/invoke"):
                self._handle_invoke(start)
            elif route_key == ("POST", "/excel/result"):
                self._handle_result(start)
            elif route_key == ("POST", "/excel/execute-start"):
                self._handle_execute_start(start)
            elif route_key == ("POST", "/excel/cancel"):
                self._handle_cancel(start)
            elif route_key == ("GET", "/excel/capabilities"):
                self._handle_capabilities(start)
            elif route_key == ("GET", "/excel/metrics"):
                self._handle_metrics(start)
            elif route_key == ("GET", "/excel/health"):
                self._handle_health(start)
            else:
                self._send_error(404, "Not Found", start=start)
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            self.close_connection = True
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
            try:
                self._send_error(500, "Internal Server Error", start=start)
            except OSError:
                self.close_connection = True

    # ------------------------------------------------------------------
    # Auth checks
    # ------------------------------------------------------------------

    def _check_origin(self) -> bool:
        """Check Origin header matches the configured sidebar origin exactly."""
        origin = self.headers.get("Origin", "")
        return constant_time_compare(origin, REQUIRED_ORIGIN)

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

    def _drain_rejected_body(self, max_bytes: int, start: float) -> bool:
        """Best-effort bounded drain for a request rejected before parsing."""
        if self.headers.get("Transfer-Encoding") is not None:
            return False
        raw_length = self.headers.get("Content-Length")
        if raw_length is None:
            return True
        try:
            length = int(raw_length)
        except ValueError:
            return False
        if length < 0 or length > max_bytes:
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

    def _read_body(self, max_bytes: int, start: float) -> bytes | None:
        """Read request body with size cap enforced before full read.

        Returns body bytes or None (error already sent).
        Sends Connection: close when rejecting to avoid desynchronization.
        """
        # Reject any Transfer-Encoding header (any value, not just "chunked")
        te = self.headers.get("Transfer-Encoding")
        if te is not None:
            self._send_error(411, "Transfer-Encoding not supported",
                             start=start, extra_headers={"Connection": "close"})
            return None

        # Check Content-Length header
        cl = self.headers.get("Content-Length")
        if cl is None:
            self._send_error(411, "Length Required", start=start,
                             extra_headers={"Connection": "close"})
            return None

        try:
            length = int(cl)
        except ValueError:
            self._send_error(400, "Bad Request: invalid Content-Length", start=start,
                             extra_headers={"Connection": "close"})
            return None

        if length < 0:
            self._send_error(400, "Bad Request: negative Content-Length", start=start,
                             extra_headers={"Connection": "close"})
            return None

        if length > max_bytes:
            self._send_error(413, "Payload Too Large", start=start,
                             extra_headers={"Connection": "close"})
            return None

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
        if _json_depth_exceeds_limit(body, MAX_JSON_DEPTH):
            self._send_error(400, "Bad Request: invalid JSON", start=start,
                             extra_headers={"Connection": "close"})
            return None
        try:
            return json.loads(body.decode("utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError, RecursionError):
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

    def _send_error(self, status: int, message: str = "",
                    start: float = 0.0, ticket_id: str = "",
                    session_id: str = "",
                    extra_headers: dict[str, str] | None = None,
                    category: str | None = None,
                    extra_body: dict[str, str] | None = None) -> None:
        """Send an error response with metadata-only logging."""
        body = json.dumps({
            "ok": False,
            "error": message,
            "category": category or self._error_category(status),
            **(extra_body or {}),
        }).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))

        # CORS for pane routes that already passed auth
        if self._is_pane_route():
            self._add_cors_response_headers()

        # Extra headers (e.g., Connection: close for body rejections)
        if extra_headers:
            for k, v in extra_headers.items():
                self.send_header(k, v)

        self.end_headers()
        self.wfile.write(body)
        elapsed = (time.monotonic() - start) * 1000
        route = self._route_template()
        _metadata_log(self.command, route, status, elapsed, ticket_id, session_id)

    def _send_json(self, status: int, data: Any,
                   start: float = 0.0, ticket_id: str = "",
                   session_id: str = "") -> None:
        """Send a JSON response."""
        body = json.dumps(data).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))

        # CORS for pane routes
        if self._is_pane_route():
            self._add_cors_response_headers()

        self.end_headers()
        self.wfile.write(body)
        elapsed = (time.monotonic() - start) * 1000
        route = self._route_template()
        _metadata_log(self.command, route, status, elapsed, ticket_id, session_id)

    def _send_empty(self, status: int, start: float = 0.0,
                    ticket_id: str = "", session_id: str = "") -> None:
        """Send an empty response, including valid pane CORS metadata."""
        self.send_response(status)
        self.send_header("Content-Length", "0")
        if self._is_pane_route():
            self._add_cors_response_headers()
        self.end_headers()
        elapsed = (time.monotonic() - start) * 1000
        _metadata_log(self.command, self._route_template(), status, elapsed,
                      ticket_id, session_id)

    def _is_pane_route(self) -> bool:
        """Check if current route is a pane route (needs CORS headers)."""
        path = getattr(self, "_parsed_path", self.path.rstrip("/") or "/")
        route_key = (self.command, path)
        return route_key in PANE_AUTH_ROUTES or route_key in ORIGIN_ONLY_ROUTES

    def _add_cors_response_headers(self) -> None:
        """Add CORS response headers only when Origin is valid."""
        origin = self.headers.get("Origin", "")
        if constant_time_compare(origin, REQUIRED_ORIGIN):
            self.send_header("Access-Control-Allow-Origin", REQUIRED_ORIGIN)
        self.send_header("Vary", "Origin")

    def _route_template(self) -> str:
        """Return a parameterized route template (no concrete IDs or query values)."""
        path = getattr(self, "_parsed_path", None)
        if path is None:
            path, _ = self._parse_url(self.path)
        parts = path.split("/")
        result_parts = []
        for p in parts:
            if validate_uuid(p):
                result_parts.append("{id}")
            else:
                result_parts.append(p)
        return "/".join(result_parts)

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
        if status == 408:
            return "Timeout"
        if status == 409:
            return "Conflict"
        if status == 413:
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

        # Only respond to valid origins. Ask the handler, not the constant: the desk bridge overrides
        # _check_origin with the configured add-in origin, and the CORS headers below follow it.
        if not self._check_origin():
            self.send_response(204)
            self.send_header("Content-Length", "0")
            self.end_headers()
            return

        self.send_response(204)
        self._send_cors_headers()
        self.send_header("Content-Length", "0")
        self.end_headers()
        elapsed = (time.monotonic() - start) * 1000
        _metadata_log("OPTIONS", self._route_template(), 204, elapsed)

    # ------------------------------------------------------------------
    # Route handlers
    # ------------------------------------------------------------------

    def _handle_session_create(self, start: float) -> None:
        """POST /excel/session — create a new session (pane, origin-gated)."""
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
        workbook_plan_fingerprint = data.get("workbookPlanFingerprint")
        if workbook_plan_fingerprint is not None and (
            not isinstance(workbook_plan_fingerprint, str)
            or not _FINGERPRINT_RE.fullmatch(workbook_plan_fingerprint)
        ):
            self._send_error(400, "workbookPlanFingerprint must be a lowercase SHA-256 fingerprint", start=start)
            return

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

        session_key = generate_uuid()  # Using UUID for session keys too (consistent)
        # Actually generate_session_key() for session key entropy
        from .protocol import generate_session_key
        session_key = generate_session_key()

        try:
            session = srv.sessions.create(
                document_id,
                pane_id,
                session_key,
                workbook_plan_fingerprint=workbook_plan_fingerprint,
            )
        except RuntimeError:
            self._send_error(429, "Session capacity reached", start=start)
            return

        response = {
            "ok": True,
            "sessionId": session.session_id,
            "sessionKey": session_key,
            "protocolVersion": PROTOCOL_VERSION,
            "expiresInSeconds": SESSION_TTL_SECONDS,
        }
        if workbook_plan_fingerprint is not None:
            response["workbookPlanFingerprint"] = workbook_plan_fingerprint
        self._send_json(201, response, start=start, session_id=session.session_id)

    def _handle_session_delete(self, start: float) -> None:
        """DELETE /excel/session — unregister session (pane auth)."""
        session = getattr(self, "_auth_session", None)
        if session is None:
            self._send_error(401, "Unauthorized", start=start)
            return

        srv = self._server_ref
        if srv:
            srv.sessions.delete_for_pane(session.session_id, session.pane_id)

        self._send_empty(204, start=start, session_id=session.session_id)

    def _handle_session_status(self, start: float) -> None:
        """GET /excel/session — get session status (plugin-key auth).

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
        """POST /excel/heartbeat — extend session TTL (pane auth).

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
        """GET /excel/pending — claim oldest pending ticket (pane auth)."""
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
                **({"trusted": ticket.trusted} if ticket.trusted is not None else {}),
            }, start=start, ticket_id=ticket.ticket_id, session_id=session.session_id)
        else:
            self._send_empty(204, start=start, session_id=session.session_id)

    def _handle_invoke(self, start: float) -> None:
        """POST /excel/invoke — mint a ticket and block for result (plugin-key auth)."""
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
        trusted = data.get("trusted")

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

        if tool == _PLAN_PLUGIN_TOOL and arguments.get("version") != WORKBOOK_PLAN_VERSION:
            self._send_error(400, "Workbook plan schema version is invalid", start=start, session_id=session_id)
            return
        if tool == _PLAN_PLUGIN_TOOL and any(
            field in data for field in (
                "documentId", "paneId", "sessionKey", "turnOperationId", "workbookPlanFingerprint",
            )
        ):
            self._send_error(400, "Workbook plan routing identity is server-owned", start=start, session_id=session_id)
            return

        if tool == _PLAN_PLUGIN_TOOL:
            if not isinstance(trusted, dict) or set(trusted) != PLAN_TRUSTED_KEYS:
                self._send_error(400, "Trusted workbook plan metadata is required", start=start, session_id=session_id)
                return
            turn_operation_id = trusted.get("turnOperationId")
            fingerprint = trusted.get("workbookPlanFingerprint")
            phase = trusted.get("phase")
            turn_key = trusted.get("turnKey")
            if (
                not isinstance(turn_operation_id, str)
                or not _CANONICAL_UUID_RE.fullmatch(turn_operation_id)
                or request_id != turn_operation_id
                or not isinstance(fingerprint, str)
                or not _FINGERPRINT_RE.fullmatch(fingerprint)
                or phase not in PLAN_PHASES
                or not isinstance(turn_key, str)
                or not _CANONICAL_UUID_RE.fullmatch(turn_key)
                or turn_key == turn_operation_id
                # The phase is derived from the validated outcome (R4g-2a).
                or (phase == "inspect") != (arguments.get("outcome") == "inspect")
            ):
                self._send_error(400, "Trusted workbook plan metadata is invalid", start=start, session_id=session_id)
                return
        elif trusted is not None:
            self._send_error(400, "Trusted metadata is only valid for the workbook plan route", start=start, session_id=session_id)
            return

        trusted_plan: dict[str, Any] = trusted if tool == _PLAN_PLUGIN_TOOL and isinstance(trusted, dict) else {}
        srv = self._server_ref
        if srv is None:
            self._send_error(500, "Server not available", start=start)
            return

        session = srv.sessions.get(session_id)
        if session is None:
            self._send_error(404, "Session not found", start=start,
                             session_id=session_id)
            return

        if tool == _PLAN_PLUGIN_TOOL:
            if session.workbook_plan_fingerprint is None:
                self._send_error(409, "Workbook plan fingerprint was not negotiated", start=start, session_id=session_id)
                return
            if trusted_plan["workbookPlanFingerprint"] != session.workbook_plan_fingerprint:
                self._send_error(409, "Workbook plan fingerprint negotiation failed", start=start, session_id=session_id)
                return
            transport_tool = _PLAN_PANE_TOOL
            ticket_trusted = {
                "turnOperationId": trusted_plan["turnOperationId"],
                "workbookPlanFingerprint": trusted_plan["workbookPlanFingerprint"],
                "phase": trusted_plan["phase"],
                "turnKey": trusted_plan["turnKey"],
            }
        else:
            transport_tool = tool
            ticket_trusted = None

        candidate = Ticket(
            ticket_id=generate_uuid(),
            session_id=session_id,
            request_id=request_id,
            tool=transport_tool,
            arguments=arguments,
            trusted=ticket_trusted,
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
        if disposition == "plan_duplicate":
            self._send_error(409, "A second workbook plan call for this trusted turn was rejected",
                             start=start, session_id=session_id)
            return
        if disposition == "replay_evicted":
            self._send_error(409, "Workbook plan replay identity is no longer executable",
                             start=start, session_id=session_id)
            return
        if disposition == "plan_fence_capacity":
            self._send_error(429, _PLAN_FENCE_CAPACITY_ERROR, start=start, session_id=session_id,
                             extra_body={"state": "plan_fence_capacity"})
            return

        # A duplicate request reattaches to the original ticket. It never
        # creates a second Office mutation.
        signaled = ticket.wait_for_result(timeout=srv.invoke_timeout_seconds)
        if not signaled and not ticket.is_terminal():
            ticket.timeout()

        if ticket.is_plan:
            self._send_plan_invoke_outcome(ticket, request_id, start, session_id)
            return

        if ticket.state == "completed" or (
            ticket.state == "terminal" and ticket.terminal_reason == "completed"
        ):
            self._send_json(200, {
                "ok": True,
                "ticketId": ticket.ticket_id,
                "requestId": request_id,
                "data": ticket.result_payload,
            }, start=start, ticket_id=ticket.ticket_id, session_id=session_id)
        elif ticket.state == "cancelled" or (
            ticket.state == "terminal" and (ticket.terminal_reason or "").startswith("cancelled")
        ):
            self._send_error(499, "Cancelled", start=start,
                             ticket_id=ticket.ticket_id, session_id=session_id)
        elif ticket.state == "timeout" or ticket.state == "terminal":
            self._send_error(504, "Invoke timed out", start=start,
                             ticket_id=ticket.ticket_id, session_id=session_id)
        else:
            self._send_error(500, "Unknown ticket state", start=start,
                             ticket_id=ticket.ticket_id, session_id=session_id)

    def _send_plan_invoke_outcome(self, ticket: Ticket, request_id: str,
                                  start: float, session_id: str) -> None:
        """Map a plan ticket's terminal reason to a distinct, content-free state."""
        reason = ticket.terminal_reason if ticket.state == "terminal" else None
        if reason in ("completed", "not_executed"):
            self._send_json(200, {
                "ok": True,
                "ticketId": ticket.ticket_id,
                "requestId": request_id,
                "data": ticket.result_payload,
                "executed": reason == "completed",
            }, start=start, ticket_id=ticket.ticket_id, session_id=session_id)
        elif reason is not None and reason.startswith("cancelled"):
            self._send_error(499, "Cancelled", start=start,
                             ticket_id=ticket.ticket_id, session_id=session_id)
        elif reason == "executing_uncertain":
            self._send_error(504, _PLAN_UNCERTAIN_ERROR, start=start,
                             ticket_id=ticket.ticket_id, session_id=session_id,
                             category=_PLAN_UNCERTAIN_CATEGORY,
                             extra_body={"state": "executing_uncertain"})
        elif reason in ("expired_claimed", "expired_unclaimed"):
            self._send_error(504, "Invoke timed out", start=start,
                             ticket_id=ticket.ticket_id, session_id=session_id,
                             extra_body={"state": "not_started"})
        else:
            self._send_error(500, "Unknown ticket state", start=start,
                             ticket_id=ticket.ticket_id, session_id=session_id)

    def _handle_plan_result(self, session: Any, ticket: Ticket, ticket_id: str,
                            data: dict[str, Any], start: float) -> None:
        """Accept a plan result only with a valid one-way executed disposition."""
        executed = data.get("executed")
        if type(executed) is not bool:
            self._send_error(400, "Bad Request: workbook plan result requires a boolean executed disposition",
                             start=start, ticket_id=ticket_id,
                             session_id=session.session_id)
            return
        disposition = ticket.submit_result(data.get("result"), session.pane_id, executed=executed)
        if disposition in ("accepted", "duplicate"):
            if disposition == "accepted":
                session.trim_history()
            self._send_json(200, {"ok": True, "ticketId": ticket_id},
                            start=start, ticket_id=ticket_id,
                            session_id=session.session_id)
            return
        if disposition == "wrong_pane":
            message = "Ticket claimed by different pane"
        elif disposition == "conflict":
            message = "Duplicate result submission with conflicting data"
        elif disposition == "terminal":
            message = "Ticket reached a non-replayable terminal state"
        else:
            state = disposition.removeprefix("invalid:")
            message = f"Ticket in state {state}, cannot accept this result disposition"
        self._send_error(409, message, start=start, ticket_id=ticket_id,
                         session_id=session.session_id)

    def _handle_result(self, start: float) -> None:
        """POST /excel/result — submit result for a claimed ticket (pane auth).

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

        if ticket.is_plan:
            self._handle_plan_result(session, ticket, ticket_id, data, start)
            return

        if ticket.state == "completed" or (
            ticket.state == "terminal" and ticket.terminal_reason == "completed"
        ):
            # Exact terminal retries are idempotent; changed payloads conflict.
            result_payload = data.get("result")
            if _json_canonical_equal(result_payload, ticket.result_payload):
                self._send_json(200, {"ok": True, "ticketId": ticket_id},
                                start=start, ticket_id=ticket_id,
                                session_id=session.session_id)
            else:
                self._send_error(409, "Duplicate result submission with conflicting data",
                                 start=start, ticket_id=ticket_id,
                                 session_id=session.session_id)
            return

        if ticket.state == "terminal":
            self._send_error(409, "Ticket reached a non-replayable terminal state",
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

    def _handle_execute_start(self, start: float) -> None:
        """POST /excel/execute-start — atomically cross a plan's execution boundary."""
        session = getattr(self, "_auth_session", None)
        if session is None:
            self._send_error(401, "Unauthorized", start=start)
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
        ticket_id = data.get("ticketId", "")
        if not isinstance(ticket_id, str) or not validate_uuid(ticket_id):
            self._send_error(400, "Bad Request: valid ticketId UUID required", start=start,
                             session_id=session.session_id)
            return
        ticket = session.find_ticket(ticket_id)
        if ticket is None:
            self._send_error(404, "No ticket found", start=start, session_id=session.session_id)
            return
        disposition = ticket.mark_executing(session.pane_id)
        if disposition == "accepted":
            self._send_json(200, {"ok": True, "ticketId": ticket_id},
                            start=start, ticket_id=ticket_id, session_id=session.session_id)
            return
        if disposition == "wrong_pane":
            self._send_error(409, "Ticket claimed by different pane", start=start,
                             ticket_id=ticket_id, session_id=session.session_id)
            return
        self._send_error(409, "Ticket execution boundary is not available", start=start,
                         ticket_id=ticket_id, session_id=session.session_id)

    def _handle_cancel(self, start: float) -> None:
        """POST /excel/cancel — cancel a ticket (plugin-key auth)."""
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
        """GET /excel/capabilities — session runtime capability metadata (plugin-key auth).

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
        """GET /excel/metrics — counts only, no content (plugin-key auth)."""
        srv = self._server_ref
        if srv is None:
            self._send_error(500, "Server not available", start=start)
            return

        metrics = srv.sessions.get_metrics()
        self._send_json(200, metrics, start=start)

    def _handle_health(self, start: float) -> None:
        """GET /excel/health — content-free managed identity (plugin-key auth)."""
        srv = self._server_ref
        self._send_json(200, {
            "ok": True,
            "service": "hermes-office-excel-companion",
            "host": "excel",
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


def _json_depth_exceeds_limit(body: bytes, max_depth: int) -> bool:
    """Return whether structural JSON container depth exceeds ``max_depth``.

    This is deliberately a byte-level, single-pass scanner. It tracks only
    string/escape state and the current structural depth, so it remains safe
    for payloads that are too deeply nested for recursive Python traversal.
    JSON syntax validity remains the parser's responsibility.
    """
    depth = 0
    in_string = False
    escaped = False

    for byte in body:
        if in_string:
            if escaped:
                escaped = False
            elif byte == 0x5C:  # backslash
                escaped = True
            elif byte == 0x22:  # quote
                in_string = False
            continue

        if byte == 0x22:  # quote
            in_string = True
        elif byte == 0x7B or byte == 0x5B:  # { or [
            depth += 1
            if depth > max_depth:
                return True
        elif byte == 0x7D or byte == 0x5D:  # } or ]
            if depth:
                depth -= 1

    return False
