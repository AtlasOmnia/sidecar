# SPDX-License-Identifier: Apache-2.0
# Copyright 2026 Atlas Omnia
"""Desk bridge: the loopback companion, approval flow and chat proxy that run INSIDE this Hermes plugin.

This file is identical in atlas-sidecar-word, atlas-sidecar-excel and atlas-sidecar-powerpoint; the app
specifics (app name, port range, companion classes, default add-in site) come from ``desk_app.py``.

Security model (README "Safety"):
- serves HTTPS localhost, binds IPv4 127.0.0.1 only on the app's port range (never Workbench's range);
  Host must be exactly localhost:<actual-port>; no plaintext or numeric-host fallback;
- every pane route requires Origin == the configured add-in site's origin (exact match);
- a pane gets a bridge session only with a per-pane grant that a human approved in Hermes
  (``/sidecar-<app>`` -> Allow). The approval is remembered as sha256(grant) in this profile's plugin data
  folder and is revocable (``/sidecar-<app> revoke``); the grant itself lives only in the pane;
- chat requests are forwarded to the Hermes API server with API_SERVER_KEY read through
  ``agent.secret_scope.get_secret`` at request time. The key is never written, logged or sent to the pane;
- tool calls reach the pane through the companion ticket queue with a per-process random plugin key held
  in memory only;
- logs carry metadata only (route template, status, elapsed); never bodies, grants, keys or document text.
  TLS private keys (not API keys) live owner-only in plugin data; user trust is always manual.
"""
from __future__ import annotations

import hashlib
import hmac
import http.client
import json
import logging
import os
import secrets
import shlex
import sys
import threading
import time
import urllib.parse
import uuid
from pathlib import Path
from typing import Any, Callable, Optional

from .desk_tls import LocalTLS

log = logging.getLogger(__name__)

DESK_PROTOCOL = "desk-1"
GRANT_HEADER = "X-Hermes-Desk-Grant"
PENDING_TTL_SECONDS = 600
MAX_PENDING = 3
MAX_GRANTS = 8
MAX_CONNECTS_PER_MINUTE = 20
MAX_CONNECT_BODY = 4096
MAX_CHAT_BODY = 10_000_000  # the Hermes API server's own request cap
UPSTREAM_TIMEOUT_SECONDS = 600.0
DEFAULT_API_SERVER_URL = "http://127.0.0.1:8642"
CODE_ALPHABET = "23456789ABCDEFGHJKMNPQRSTVWXYZ"  # no 0/O/1/I/L/U
CODE_LENGTH = 8
STREAM_CHUNK = 64 * 1024
# Response headers relayed from the API server to the pane (never Set-Cookie, never auth material).
RELAYED_RESPONSE_HEADERS = ("Content-Type", "X-Hermes-Session-Id")


class DeskError(RuntimeError):
    """A bounded, category-labelled failure that is safe to show to the user."""

    def __init__(self, category: str, message: str) -> None:
        super().__init__(message)
        self.category = category


# --------------------------------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------------------------------

def check_addin_base(base: Any) -> str:
    """The add-in site's base URL: https, a host, a path ending in /, no query/fragment/userinfo."""
    parts = urllib.parse.urlsplit(base) if isinstance(base, str) else None
    if (parts is None or parts.scheme != "https" or not parts.hostname or "@" in parts.netloc
            or not parts.path.endswith("/") or parts.query or parts.fragment or base.endswith("//")
            or any(ch.isspace() for ch in base)):
        raise DeskError("Configuration", "addin_base must be an https URL ending in / (no query, fragment or user)")
    return base


def origin_of(base: str) -> str:
    """scheme://host[:port] of the add-in site: the only Origin the bridge accepts."""
    parts = urllib.parse.urlsplit(check_addin_base(base))
    return f"{parts.scheme}://{parts.netloc.lower()}"


def _own_api_prefix() -> str:
    """Capture the owning profile while plugin load is still in its Hermes scope."""
    from hermes_constants import get_hermes_home, profile_name_for_home
    profile = profile_name_for_home(get_hermes_home())
    return "" if profile in (None, "default") else f"/p/{urllib.parse.quote(profile, safe='')}"


def check_api_server_url(url: Any, *, profile_prefix: Optional[str] = None) -> tuple[str, int]:
    """(host, port) of the loopback Hermes API; only this profile's optional path is allowed."""
    if profile_prefix is None:
        profile_prefix = _own_api_prefix()
    parts = urllib.parse.urlsplit(url) if isinstance(url, str) else None
    try:
        port = parts.port if parts else None
    except ValueError:
        port = None
    if (parts is None or parts.scheme != "http" or parts.hostname not in ("127.0.0.1", "localhost", "::1")
            or parts.username or parts.password or parts.path not in ("", "/", profile_prefix) or parts.query
            or parts.fragment or port is None):
        raise DeskError("Configuration", "api_server_url must be a loopback http URL such as http://127.0.0.1:8642")
    return parts.hostname, port


def normalize_code(raw: Any) -> str:
    if not isinstance(raw, str):
        return ""
    return "".join(ch for ch in raw.upper() if ch.isalnum())[:32]


def format_code(code: str) -> str:
    return f"{code[:4]}-{code[4:]}"


def _hash(secret: str) -> str:
    return hashlib.sha256(secret.encode("utf-8")).hexdigest()


def _clean_client(client: Any) -> dict:
    client = client if isinstance(client, dict) else {}
    out = {}
    for key in ("host", "platform", "version"):
        value = client.get(key)
        if isinstance(value, (str, int, float)) and not isinstance(value, bool):
            text = "".join(ch for ch in str(value) if ch.isprintable())[:64]
            if text:
                out[key] = text
    return out


# --------------------------------------------------------------------------------------------------
# Approval store (persisted: approvals survive Office and Hermes restarts; revocable)
# --------------------------------------------------------------------------------------------------

def _atomic_write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    fd = os.open(str(tmp), os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(data, handle, separators=(",", ":"), sort_keys=True)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp, path)
    finally:
        if tmp.exists():
            tmp.unlink()
    if os.name != "nt":
        os.chmod(path, 0o600)


def _read_json(path: Path) -> dict:
    try:
        with open(path, "r", encoding="utf-8") as handle:
            value = json.load(handle)
    except FileNotFoundError:
        return {}
    except (OSError, ValueError):
        log.warning("desk bridge: unreadable state file %s ignored", path.name)
        return {}
    return value if isinstance(value, dict) else {}


class ApprovalStore:
    """Per-app approvals (hashed grants) and pending connect requests, in the plugin data folder.

    Files: approvals.json (grants) and pending.json (connect requests waiting for Allow/Deny). Neither
    holds a usable secret: only sha256 digests of the pane's grant and of a request's poll secret.
    Both are re-read on every decision so an Allow given in another Hermes process (the CLI) counts.
    """

    def __init__(self, data_dir: Path, app: str, clock: Callable[[], float] = time.time) -> None:
        self.data_dir = Path(data_dir)
        self.app = app
        self._clock = clock
        self._lock = threading.RLock()
        self._connect_times: list[float] = []
        self._grant_cache: tuple[Optional[tuple[int, int]], set[str]] = (None, set())

    @property
    def approvals_path(self) -> Path:
        return self.data_dir / "approvals.json"

    @property
    def pending_path(self) -> Path:
        return self.data_dir / "pending.json"

    # -- raw state --
    def _grants(self) -> list:
        apps = _read_json(self.approvals_path).get("apps")
        entry = apps.get(self.app) if isinstance(apps, dict) else None
        grants = entry.get("grants") if isinstance(entry, dict) else None
        return [g for g in grants if isinstance(g, dict)] if isinstance(grants, list) else []

    def _save_grants(self, grants: list) -> None:
        apps = _read_json(self.approvals_path).get("apps")
        if not isinstance(apps, dict):
            apps = {}
        if grants:
            apps[self.app] = {"grants": grants}
        else:
            apps.pop(self.app, None)
        _atomic_write_json(self.approvals_path, {"version": 1, "apps": apps})

    def _requests(self) -> list:
        now = self._clock()
        items = _read_json(self.pending_path).get("requests")
        items = items if isinstance(items, list) else []
        return [r for r in items if isinstance(r, dict) and r.get("app") == self.app
                and isinstance(r.get("expiresAt"), (int, float)) and r["expiresAt"] > now]

    def _save_requests(self, requests: list) -> None:
        items = _read_json(self.pending_path).get("requests")
        items = items if isinstance(items, list) else []
        others = [r for r in items if isinstance(r, dict) and r.get("app") != self.app]
        _atomic_write_json(self.pending_path, {"version": 1, "requests": others + requests})

    # -- connect flow --
    def create_request(self, origin: str, client: Any = None) -> dict:
        with self._lock:
            now = self._clock()
            self._connect_times = [t for t in self._connect_times if now - t < 60]
            if len(self._connect_times) >= MAX_CONNECTS_PER_MINUTE:
                raise DeskError("RateLimited", "Too many connect requests; wait a minute")
            self._connect_times.append(now)
            requests = [r for r in self._requests() if r.get("status") == "pending"]
            requests.sort(key=lambda r: r.get("createdAt", 0))
            while len(requests) >= MAX_PENDING:
                requests.pop(0)
            code = "".join(secrets.choice(CODE_ALPHABET) for _ in range(CODE_LENGTH))
            poll_secret = secrets.token_urlsafe(32)
            request = {
                "id": str(uuid.uuid4()), "app": self.app, "code": code, "origin": origin,
                "client": _clean_client(client), "status": "pending", "createdAt": now,
                "expiresAt": now + PENDING_TTL_SECONDS, "pollHash": _hash(poll_secret),
            }
            requests.append(request)
            self._save_requests(requests)
            return {"requestId": request["id"], "code": format_code(code), "pollSecret": poll_secret,
                    "expiresInSeconds": PENDING_TTL_SECONDS}

    def pending(self) -> list[dict]:
        """Requests waiting for a decision, newest first (public fields only)."""
        now = self._clock()
        out = []
        for r in sorted(self._requests(), key=lambda r: r.get("createdAt", 0), reverse=True):
            if r.get("status") == "pending":
                out.append({"requestId": r["id"], "code": format_code(str(r.get("code", ""))),
                            "origin": r.get("origin", ""), "client": r.get("client", {}),
                            "ageSeconds": int(max(0, now - r.get("createdAt", now)))})
        return out

    def decide(self, allow: bool, code: Optional[str] = None) -> dict:
        """Allow or deny one pending request (by code, or the only one waiting)."""
        with self._lock:
            requests = self._requests()
            waiting = [r for r in requests if r.get("status") == "pending"]
            wanted = normalize_code(code) if code else ""
            if wanted:
                match = [r for r in waiting if hmac.compare_digest(str(r.get("code", "")), wanted)]
            else:
                match = waiting
            if not match:
                raise DeskError("NotFound", "No sidebar connect request is waiting" + (" with that code" if wanted else ""))
            if len(match) > 1:
                raise DeskError("Ambiguous", "More than one sidebar is waiting; name its code")
            chosen = match[0]
            chosen["status"] = "approved" if allow else "denied"
            chosen["decidedAt"] = self._clock()
            self._save_requests(requests)
            return {"requestId": chosen["id"], "code": format_code(chosen["code"]), "status": chosen["status"],
                    "origin": chosen.get("origin", "")}

    def poll(self, request_id: Any, poll_secret: Any, origin: str) -> dict:
        """The pane asks how its request went. On approval the grant is minted and returned ONCE."""
        if not isinstance(request_id, str) or not isinstance(poll_secret, str) or len(poll_secret) > 200:
            raise DeskError("Validation", "requestId and pollSecret are required")
        with self._lock:
            requests = self._requests()
            entry = next((r for r in requests if r.get("id") == request_id), None)
            if entry is None:
                return {"status": "expired"}
            if not hmac.compare_digest(str(entry.get("pollHash", "")), _hash(poll_secret)) \
                    or entry.get("origin") != origin:
                raise DeskError("Authentication", "Unknown connect request")
            status = entry.get("status")
            if status == "pending":
                return {"status": "pending", "code": format_code(entry["code"])}
            requests.remove(entry)
            self._save_requests(requests)
            if status != "approved":
                return {"status": "denied"}
            grant = secrets.token_urlsafe(32)
            grants = self._grants()
            grants.append({"id": str(uuid.uuid4()), "hash": _hash(grant), "origin": origin,
                           "client": entry.get("client", {}), "createdAt": self._clock()})
            grants.sort(key=lambda g: g.get("createdAt", 0))
            self._save_grants(grants[-MAX_GRANTS:])
            return {"status": "approved", "grant": grant}

    # -- grants --
    def check_grant(self, grant: Any, origin: str) -> Optional[str]:
        """The grant's id when it is an approved grant for this origin, else None."""
        if not isinstance(grant, str) or not grant or len(grant) > 200:
            return None
        digest = _hash(grant)
        found = None
        for g in self._grants():
            if hmac.compare_digest(str(g.get("hash", "")), digest) and g.get("origin") == origin:
                found = g.get("id")
        return found if isinstance(found, str) else None

    def grant_active(self, grant_id: Optional[str]) -> bool:
        if not grant_id:
            return False
        try:
            st = self.approvals_path.stat()
            stamp: Optional[tuple[int, int]] = (st.st_mtime_ns, st.st_size)
        except OSError:
            return False
        with self._lock:
            cached_stamp, ids = self._grant_cache
            if cached_stamp != stamp:
                ids = {g.get("id") for g in self._grants() if isinstance(g.get("id"), str)}
                self._grant_cache = (stamp, ids)
            return grant_id in ids

    def revoke(self) -> int:
        """Forget every approval for this app and every waiting request. Returns grants removed."""
        with self._lock:
            removed = len(self._grants())
            self._save_grants([])
            self._save_requests([])
            self._grant_cache = (None, set())
            return removed

    def status(self) -> dict:
        grants = self._grants()
        return {"approved": bool(grants), "grants": len(grants), "pending": len(self.pending())}


# --------------------------------------------------------------------------------------------------
# API server key (Hermes' secret route; memory only)
# --------------------------------------------------------------------------------------------------

class KeyProvider:
    """API_SERVER_KEY through ``agent.secret_scope.get_secret`` (rule 11). The value is kept in memory
    only, so bridge threads (which Hermes does not scope) can reuse the value read in a scoped moment
    (plugin load, tool call, slash command). Never written, never logged, never returned to the pane."""

    def __init__(self, reader: Optional[Callable[[], Optional[str]]] = None) -> None:
        self._reader = reader or _read_api_server_key
        self._lock = threading.Lock()
        self._value: Optional[str] = None

    def refresh(self) -> None:
        try:
            value = self._reader()
        except Exception:  # noqa: BLE001 - an unscoped read under multiplexing raises; keep the last value
            return
        if isinstance(value, str) and value.strip():
            with self._lock:
                self._value = value.strip()

    def __call__(self) -> Optional[str]:
        self.refresh()
        with self._lock:
            return self._value

    def __repr__(self) -> str:  # never show the value
        return "KeyProvider(<redacted>)"


def _read_api_server_key() -> Optional[str]:
    from agent.secret_scope import get_secret  # documented plugin secret route
    return get_secret("API_SERVER_KEY")


# --------------------------------------------------------------------------------------------------
# Chat proxy (folded into the bridge; replaces the old key-injecting proxy process)
# --------------------------------------------------------------------------------------------------

class ChatProxy:
    def __init__(self, api_server_url: str, key_provider: Callable[[], Optional[str]],
                 timeout: float = UPSTREAM_TIMEOUT_SECONDS) -> None:
        profile_prefix = _own_api_prefix()
        self.host, self.port = check_api_server_url(api_server_url, profile_prefix=profile_prefix)
        self.chat_path = f"{profile_prefix}/v1/chat/completions"
        self.url = api_server_url.rstrip("/")
        self.key_provider = key_provider
        self.timeout = float(timeout)

    def forward(self, handler: Any, body: bytes, session_id: str, start: float) -> None:
        key = self.key_provider()
        if not key:
            handler._send_error(503, "Hermes API server key is not available (set API_SERVER_KEY and "
                                "enable the API server)", start=start)
            return
        headers = {
            "Content-Type": "application/json",
            "Content-Length": str(len(body)),
            "Authorization": f"Bearer {key}",
            "X-Hermes-Session-Id": session_id,
            "Accept": handler.headers.get("Accept", "") or "text/event-stream, application/json",
            "Connection": "close",
        }
        conn = http.client.HTTPConnection(self.host, self.port, timeout=self.timeout)
        try:
            try:
                conn.request("POST", self.chat_path, body=body, headers=headers)
                upstream = conn.getresponse()
            except (OSError, http.client.HTTPException):
                handler._send_error(502, f"The Hermes API server is not reachable at {self.url}; enable it "
                                    "(API server) and start the Hermes gateway", start=start)
                return
            handler.close_connection = True
            handler.send_response(upstream.status)
            for name in RELAYED_RESPONSE_HEADERS:
                value = upstream.getheader(name)
                if value:
                    handler.send_header(name, value)
            handler.send_header("Cache-Control", "no-store")
            handler.send_header("Connection", "close")
            handler._add_cors_response_headers()
            handler.end_headers()
            while True:
                try:
                    chunk = upstream.read1(STREAM_CHUNK)
                except (OSError, http.client.HTTPException):
                    break
                if not chunk:
                    break
                try:
                    handler.wfile.write(chunk)
                    handler.wfile.flush()
                except OSError:
                    break  # the pane went away (Stop or closed); the upstream turn is abandoned
        finally:
            conn.close()


# --------------------------------------------------------------------------------------------------
# HTTP layer: the app's companion handler, wrapped
# --------------------------------------------------------------------------------------------------

STATIC_MIME_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".ico": "image/x-icon",
    ".json": "application/json",
    ".xml": "application/xml; charset=utf-8",
    ".txt": "text/plain; charset=utf-8",
}
STATIC_CSP = (
    "default-src 'none'; base-uri 'self'; object-src 'none'; "
    "script-src 'self' https://appsforoffice.microsoft.com; "
    "style-src 'self' 'unsafe-inline'; "
    "connect-src 'self' https://appsforoffice.microsoft.com; "
    "img-src 'self' data:; font-src 'self' data:; "
    "frame-ancestors 'self' https://*.office.com https://*.office365.com "
    "https://*.officeapps.live.com https://*.office.net https://*.microsoft365.com"
)


def _sidebar_checksums(sidebar_dir: Path) -> dict[str, str]:
    """Read the bundled allowlist of static sidebar assets."""
    try:
        lines = (sidebar_dir / "SHA256SUMS").read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeDecodeError):
        return {}
    allowed: dict[str, str] = {}
    for line in lines:
        digest, separator, name = line.partition("  ")
        if (not separator or len(digest) != 64 or any(c not in "0123456789abcdef" for c in digest)
                or not name or "\\" in name or "\x00" in name or ".." in name):
            continue
        parts = name.split("/")
        if any(part in ("", ".", "..") for part in parts):
            continue
        allowed[name] = digest
    return allowed


def _read_sidebar_asset(sidebar_dir: Path, name: str) -> Optional[bytes]:
    """Read a listed regular file without following symlinks or leaving the bundle."""
    expected = _sidebar_checksums(sidebar_dir).get(name)
    if expected is None:
        return None
    try:
        root = sidebar_dir.resolve(strict=True)
        candidate = sidebar_dir
        for part in name.split("/"):
            candidate = candidate / part
            if candidate.is_symlink():
                return None
        resolved = candidate.resolve(strict=True)
        if not resolved.is_relative_to(root) or not resolved.is_file():
            return None
        with open(resolved, "rb") as handle:
            content = handle.read()
    except (OSError, RuntimeError):
        return None
    if not hmac.compare_digest(hashlib.sha256(content).hexdigest(), expected):
        return None
    return content


def make_desk_handler(base: type, app: str, command: str) -> type:
    """Subclass the app's companion handler: static files, Host check, origin, approval and chat proxy."""
    prefix = f"/{app}"
    static_prefix = f"/{command}/"
    sidebar_dir = Path(__file__).with_name("sidebar")
    hello, connect, connect_status = f"{prefix}/hello", f"{prefix}/connect", f"{prefix}/connect/status"
    session_route = ("POST", f"{prefix}/session")
    chat = "/v1/chat/completions"
    desk_paths = {hello, connect, connect_status, chat}

    class DeskHandler(base):  # type: ignore[misc, valid-type]
        server_version = "HermesDeskBridge/1"
        sys_version = ""
        _desk_grant_id: Optional[str] = None

        def _desk(self) -> "DeskBridge":
            return self._server_ref.desk  # type: ignore[attr-defined]

        # -- gates --
        def _host_ok(self) -> bool:
            srv = self._server_ref
            if srv is None:
                return False
            port = srv.server_address()[1]
            return self.headers.get_all("Host", []) == [f"localhost:{port}"]

        def _origin_value(self) -> str:
            return self.headers.get("Origin", "")

        def _check_origin(self) -> bool:
            allowed = self._desk().allowed_origin
            return hmac.compare_digest(self._origin_value().encode("utf-8"), allowed.encode("utf-8"))

        def _send_cors_headers(self) -> None:
            if self._check_origin():
                self.send_header("Access-Control-Allow-Origin", self._desk().allowed_origin)
            self.send_header("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", ", ".join(self._desk().cors_headers))
            self.send_header("Access-Control-Max-Age", "600")
            if self.headers.get("Access-Control-Request-Private-Network") == "true" and self._check_origin():
                self.send_header("Access-Control-Allow-Private-Network", "true")
            self.send_header("Vary", "Origin")

        def _add_cors_response_headers(self) -> None:
            if self._check_origin():
                self.send_header("Access-Control-Allow-Origin", self._desk().allowed_origin)
            self.send_header("Vary", "Origin")

        def _is_pane_route(self) -> bool:
            path = getattr(self, "_parsed_path", None)
            if path in desk_paths:
                return True
            return super()._is_pane_route()

        def _check_session_auth(self) -> Any:
            session = super()._check_session_auth()
            if session is None:
                return None
            if not self._desk().session_allowed(session.session_id):
                self._desk().drop_session(session)  # approval revoked: this pane must connect again
                return None
            return session

        def _send_json(self, status: int, data: Any, *args: Any, **kwargs: Any) -> None:
            if (status == 201 and getattr(self, "_desk_creating_session", False) and isinstance(data, dict)
                    and isinstance(data.get("sessionId"), str)):
                self._desk().bind_session(data["sessionId"], self._desk_grant_id)
            super()._send_json(status, data, *args, **kwargs)

        def _refuse(self, status: int, message: str, start: float, extra: Optional[dict] = None,
                    drain: bool = True) -> None:
            if drain:  # a body that was never read must not be parsed as the next request
                self._drain_rejected_body(MAX_CONNECT_BODY, start)
            self.close_connection = True
            body = {"ok": False, "error": message, "category": self._error_category(status)}
            if extra:
                body.update(extra)
            raw = json.dumps(body).encode()
            self.send_response(status)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(raw)))
            self.send_header("Connection", "close")
            self._add_cors_response_headers()
            self.end_headers()
            self.wfile.write(raw)

        # -- routing --
        def _send_static_error(self, status: int) -> None:
            self.close_connection = True
            self.send_response(status)
            self.send_header("Content-Length", "0")
            self.send_header("Connection", "close")
            self.end_headers()

        def _serve_static(self, method: str, raw_path: str) -> None:
            origins = self.headers.get_all("Origin", [])
            if origins and (len(origins) != 1 or not hmac.compare_digest(
                    origins[0].encode("utf-8"), self._desk().allowed_origin.encode("utf-8"))):
                return self._send_static_error(403)

            encoded_name = raw_path[len(static_prefix):]
            if not encoded_name or "%2f" in encoded_name.lower():
                return self._send_static_error(404)
            try:
                name = urllib.parse.unquote(encoded_name, encoding="utf-8", errors="strict")
            except UnicodeDecodeError:
                return self._send_static_error(404)
            parts = name.split("/")
            if (not name or ".." in name or "\\" in name or "\x00" in name or name.startswith("/")
                    or "//" in name or any(part in ("", ".", "..") for part in parts)):
                return self._send_static_error(404)
            content = _read_sidebar_asset(sidebar_dir, name)
            if content is None:
                return self._send_static_error(404)

            self.send_response(200)
            self.send_header("Content-Type", STATIC_MIME_TYPES.get(Path(name).suffix.lower(), "application/octet-stream"))
            self.send_header("Content-Length", str(len(content)))
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("Referrer-Policy", "no-referrer")
            self.send_header("Cache-Control", "no-cache")
            self.send_header("Content-Security-Policy", STATIC_CSP)
            self.end_headers()
            if method == "GET":
                self.wfile.write(content)

        # -- routing --
        def do_HEAD(self) -> None:  # noqa: N802
            self._cancel_header_deadline()
            self._dispatch("HEAD")

        def do_PUT(self) -> None:  # noqa: N802
            self._cancel_header_deadline()
            self._dispatch("PUT")

        def do_OPTIONS(self) -> None:  # noqa: N802
            self._cancel_header_deadline()
            start = time.monotonic()
            raw_path = urllib.parse.urlsplit(self.path).path
            path, _ = self._parse_url(self.path)
            self._parsed_path = path
            if raw_path.startswith(static_prefix):
                if not self._host_ok():
                    return self._refuse(421, "Unexpected Host", start)
                return self._send_static_error(405)
            if not self._host_ok() or not self._check_origin():
                self.close_connection = True
                self.send_response(403)
                self.send_header("Content-Length", "0")
                self.send_header("Connection", "close")
                self.end_headers()
                return
            if path in desk_paths or path == session_route[1]:
                self.send_response(204)
                self._send_cors_headers()
                self.send_header("Content-Length", "0")
                self.end_headers()
                return
            self._handle_options(start)

        def _dispatch(self, method: str) -> None:
            start = time.monotonic()
            raw_path = urllib.parse.urlsplit(self.path).path
            path, query = self._parse_url(self.path)
            self._parsed_path, self._parsed_query = path, query
            if not self._host_ok():
                return self._refuse(421, "Unexpected Host", start)
            if raw_path.startswith(static_prefix):
                if method not in ("GET", "HEAD"):
                    return self._send_static_error(405)
                return self._serve_static(method, raw_path)
            route = (method, path)
            if path in desk_paths:
                if not self._check_origin():
                    return self._refuse(403, "Origin not allowed", start)
                try:
                    if route == ("GET", hello):
                        return self._send_json(200, self._desk().hello(), start=start)
                    if route == ("POST", connect):
                        return self._desk_connect(start)
                    if route == ("POST", connect_status):
                        return self._desk_connect_status(start)
                    if route == ("POST", chat):
                        return self._desk_chat(start)
                except DeskError as exc:
                    status = {"RateLimited": 429, "Authentication": 401, "Validation": 400}.get(exc.category, 400)
                    return self._refuse(status, str(exc), start, drain=False)  # body already read
                return self._refuse(405, "Method Not Allowed", start)
            if route == session_route:
                if not self._check_origin():
                    return self._refuse(403, "Origin not allowed", start)
                grant_id = self._desk().store.check_grant(self.headers.get(GRANT_HEADER, ""),
                                                          self._desk().allowed_origin)
                if grant_id is None:
                    return self._refuse(403, f"Hermes approval required: click Connect to Hermes, then "
                                        f"allow it in Hermes with /{self._desk().command}", start,
                                        {"approval": "required"})
                self._desk_grant_id = grant_id
                self._desk_creating_session = True
                try:
                    return super()._dispatch(method)
                finally:
                    self._desk_creating_session = False
            return super()._dispatch(method)

        def _read_object(self, max_bytes: int, start: float) -> Optional[dict]:
            body = self._read_body(max_bytes, start)
            if body is None:
                return None
            data = self._parse_json(body, start)
            if data is None:
                return None
            if not isinstance(data, dict):
                self._send_error(400, "Bad Request: expected JSON object", start=start)
                return None
            return data

        def _desk_connect(self, start: float) -> None:
            data = self._read_object(MAX_CONNECT_BODY, start)
            if data is None:
                return
            out = self._desk().store.create_request(self._desk().allowed_origin, data.get("client"))
            out.update({"ok": True, "command": f"/{self._desk().command}"})
            self._send_json(201, out, start=start)

        def _desk_connect_status(self, start: float) -> None:
            data = self._read_object(MAX_CONNECT_BODY, start)
            if data is None:
                return
            out = self._desk().store.poll(data.get("requestId"), data.get("pollSecret"), self._desk().allowed_origin)
            out["ok"] = True
            self._send_json(200, out, start=start)

        def _desk_chat(self, start: float) -> None:
            session = self._check_session_auth()
            if session is None:
                return self._refuse(401, "Unauthorized", start)
            body = self._read_body(MAX_CHAT_BODY, start)
            if body is None:
                return
            self._desk().proxy.forward(self, body, session.session_id, start)

    DeskHandler.__name__ = f"Desk{base.__name__}"
    return DeskHandler


class DeskBridge:
    """One profile's bridge for one Office app: companion server + approval store + chat proxy."""

    def __init__(self, *, app: str, command: str, service: str, server_cls: type, handler_cls: type,
                 store: ApprovalStore, allowed_origin: str, proxy: ChatProxy, ports: tuple,
                 dynamic_loopback_origin: bool = False, extra_cors_headers: tuple = (),
                 server_kwargs: Optional[dict] = None) -> None:
        self.app = app
        self.command = command
        self.service = service
        self.store = store
        self.tls = LocalTLS(store.data_dir)
        self._allowed_origin = allowed_origin
        self._dynamic_loopback_origin = dynamic_loopback_origin
        self.proxy = proxy
        self.ports = tuple(ports)
        self.instance = secrets.token_hex(8)
        self.plugin_key = secrets.token_hex(32)  # in memory only: authorizes this process's own tool calls
        self.cors_headers = tuple(dict.fromkeys(("Content-Type", "X-Hermes-Session-Key", "X-Hermes-Session-Id",
                                                 "X-Hermes-Pane-Id", GRANT_HEADER, "Accept")
                                                + tuple(extra_cors_headers)))
        self._server_kwargs = dict(server_kwargs or {})
        self._lock = threading.Lock()
        self._server: Any = None
        self._session_grants: dict[str, str] = {}
        desk_handler = make_desk_handler(handler_cls, app, command)
        bridge = self

        class DeskServer(server_cls):  # type: ignore[misc, valid-type]
            desk = bridge

            def _handler_base(self) -> type:
                return desk_handler

        self._server_cls = DeskServer

    def __repr__(self) -> str:
        return f"DeskBridge(app={self.app!r}, port={self.port!r})"

    @property
    def port(self) -> Optional[int]:
        srv = self._server
        return srv.server_address()[1] if srv is not None else None

    @property
    def base_url(self) -> Optional[str]:
        port = self.port
        return f"https://localhost:{port}" if port else None

    @property
    def allowed_origin(self) -> str:
        port = self.port
        if self._dynamic_loopback_origin and port:
            return f"https://localhost:{port}"
        return self._allowed_origin

    @property
    def sessions(self) -> Any:
        return self._server.sessions if self._server is not None else None

    def ensure_started(self) -> int:
        with self._lock:
            if self._server is not None:
                return self.port  # type: ignore[return-value]
            try:
                tls_context = self.tls.ensure()
            except Exception as exc:
                # Never expose parser exceptions that might include key material.
                raise DeskError("TLS", "Local TLS identity could not be loaded. Install cryptography>=43; "
                                "check plugin data permissions and the documented certificate recovery steps "
                                f"({type(exc).__name__}).") from None
            last: Optional[Exception] = None
            for port in self.ports:
                server = self._server_cls(host="127.0.0.1", port=port, plugin_key=self.plugin_key,
                                          ssl_context=tls_context,
                                          **self._server_kwargs)
                try:
                    self._server = server
                    server.start()
                except OSError as exc:
                    self._server = None
                    last = exc
                    continue
                log.info("desk bridge: %s listening on https://localhost:%s (IPv4 loopback)", self.app, port)
                print(self.tls.trust_instructions())
                return port
            span = f"127.0.0.1:{self.ports[0]}-{self.ports[-1]}"
            raise DeskError("PortsBusy", f"Every Sidecar {self.app} bridge port ({span}) is in use "
                                         f"({type(last).__name__ if last else 'no port'}).")

    def stop(self) -> None:
        with self._lock:
            server, self._server = self._server, None
        if server is not None:
            server.stop()

    def hello(self) -> dict:
        return {"ok": True, "service": self.service, "app": self.app, "protocol": DESK_PROTOCOL,
                "instance": self.instance, "command": f"/{self.command}"}

    # -- sessions minted from grants --
    def bind_session(self, session_id: str, grant_id: Optional[str]) -> None:
        if grant_id:
            with self._lock:
                self._session_grants[session_id] = grant_id

    def session_allowed(self, session_id: str) -> bool:
        with self._lock:
            grant_id = self._session_grants.get(session_id)
        return self.store.grant_active(grant_id)

    def drop_session(self, session: Any) -> None:
        with self._lock:
            self._session_grants.pop(session.session_id, None)
        sessions = self.sessions
        if sessions is not None:
            try:
                sessions.delete(session.session_id)
            except Exception:  # noqa: BLE001 - already gone
                pass

    def drop_all_sessions(self) -> int:
        with self._lock:
            ids = list(self._session_grants)
            self._session_grants.clear()
        sessions = self.sessions
        if sessions is not None:
            for session_id in ids:
                try:
                    sessions.delete(session_id)
                except Exception:  # noqa: BLE001
                    pass
        return len(ids)

    def has_session(self, session_id: str) -> bool:
        with self._lock:
            return session_id in self._session_grants


# --------------------------------------------------------------------------------------------------
# Plugin runtime: configuration, start policy, the /sidecar-<app> command and `hermes sidecar-<app>`
# --------------------------------------------------------------------------------------------------

def _in_gateway() -> bool:
    """True in the process that serves the Hermes API server (the gateway sets this before it loads
    plugins; a plugin host started by the gateway inherits it)."""
    # CLI children inherit the marker; registration happens before cli_run().
    commands = {"sidecar-word", "sidecar-excel", "sidecar-powerpoint"}
    return os.environ.get("_HERMES_GATEWAY") == "1" and not commands.intersection(sys.argv[1:])


def _data_dir_for(plugin_id: str) -> Path:
    try:
        from plugins.plugin_storage import plugin_data_dir  # documented plugin storage API
        return Path(plugin_data_dir(plugin_id))
    except ImportError:
        from hermes_constants import get_hermes_home
        root = Path(get_hermes_home()) / "plugin-data" / plugin_id
        root.mkdir(parents=True, exist_ok=True)
        return root


class DeskRuntime:
    """Everything one loaded copy of the plugin needs: settings, approvals, bridge, commands."""

    def __init__(self, *, appcfg: Any, data_dir: Path, addin_base: str, api_server_url: str,
                 key_provider: Callable[[], Optional[str]], autostart: bool = True,
                 ask_choice: Optional[Callable[[str, list], Optional[str]]] = None,
                 in_gateway: Callable[[], bool] = _in_gateway) -> None:
        self.cfg = appcfg
        self.addin_base = check_addin_base(addin_base)
        self._default_addin_base = self.addin_base == appcfg.DEFAULT_ADDIN_BASE
        self.allowed_origin = origin_of(self.addin_base)
        self.api_server_url = api_server_url
        self.store = ApprovalStore(Path(data_dir), appcfg.APP)
        self.key_provider = key_provider
        self.proxy = ChatProxy(api_server_url, key_provider)
        self.bridge = DeskBridge(
            app=appcfg.APP, command=appcfg.COMMAND, service=appcfg.SERVICE,
            server_cls=appcfg.CompanionServer, handler_cls=appcfg.CompanionHandler, store=self.store,
            allowed_origin=self.allowed_origin, proxy=self.proxy, ports=appcfg.PORTS,
            dynamic_loopback_origin=self._default_addin_base,
            extra_cors_headers=getattr(appcfg, "EXTRA_CORS_HEADERS", ()),
            server_kwargs=getattr(appcfg, "SERVER_KWARGS", None))
        self.autostart = bool(autostart)
        self._ask_choice = ask_choice
        self._in_gateway = in_gateway

    def __repr__(self) -> str:
        return f"DeskRuntime(app={self.cfg.APP!r}, origin={self.allowed_origin!r})"

    @classmethod
    def from_ctx(cls, ctx: Any, appcfg: Any) -> "DeskRuntime":
        def setting(key: str, default: Any) -> Any:
            try:
                value = ctx.get_config(key, default)
            except Exception:  # noqa: BLE001 - no config available (tests, validate): use the default
                return default
            return default if value in (None, "") else value

        autostart = setting("autostart", True)
        provider = KeyProvider()
        provider.refresh()  # plugin load runs in the owning profile's scope
        return cls(appcfg=appcfg, data_dir=_data_dir_for(appcfg.PLUGIN_ID),
                   addin_base=setting("addin_base", appcfg.DEFAULT_ADDIN_BASE),
                   api_server_url=setting("api_server_url", DEFAULT_API_SERVER_URL),
                   key_provider=provider,
                   autostart=autostart if isinstance(autostart, bool) else str(autostart).lower() != "false",
                   ask_choice=getattr(ctx, "ask_choice", None))

    # -- lifecycle --
    def maybe_autostart(self) -> Optional[int]:
        if not (self.autostart and self._in_gateway()):
            return None
        try:
            return self.ensure_started()
        except DeskError as exc:
            log.warning("desk bridge: %s not started: %s", self.cfg.APP, exc)
            return None

    def ensure_started(self) -> int:
        refresh = getattr(self.key_provider, "refresh", None)
        if callable(refresh):
            refresh()
        port = self.bridge.ensure_started()
        if self._default_addin_base:
            self.addin_base = f"https://localhost:{port}/{self.cfg.COMMAND}/"
            self.allowed_origin = origin_of(self.addin_base)
        return port

    def stop(self) -> None:
        self.bridge.stop()

    # -- human-facing text --
    def status_text(self, *, include_trust: bool = False) -> str:
        st = self.store.status()
        port = self.bridge.port
        lines = [f"Sidecar for {self.cfg.APP_TITLE}:",
                 f"- bridge: {'listening on https://localhost:' + str(port) if port else 'not running in this process'}",
                 f"- add-in site: {self.addin_base} (only origin {self.allowed_origin} is accepted)",
                 f"- approval: {'allowed (' + str(st['grants']) + ' sidebar grant(s))' if st['approved'] else 'not allowed yet'}"]
        for p in self.store.pending():
            lines.append(f"- waiting: sidebar request {p['code']} ({p['ageSeconds']} s ago)")
        if include_trust and self.bridge.tls.ca_path.is_file():
            lines.append(self.bridge.tls.trust_instructions())
        return "\n".join(lines)

    def _decided(self, out: dict) -> str:
        if out["status"] == "approved":
            return (f"Allowed: the {self.cfg.APP_TITLE} sidebar (code {out['code']}) is connecting. Hermes remembers "
                    f"this; run /{self.cfg.COMMAND} revoke to take it back.")
        return f"Denied: the {self.cfg.APP_TITLE} sidebar request {out['code']} was refused."

    def revoke_text(self) -> str:
        removed = self.store.revoke()
        dropped = self.bridge.drop_all_sessions()
        return (f"Revoked Sidecar for {self.cfg.APP_TITLE}: {removed} remembered approval(s) removed, "
                f"{dropped} open sidebar session(s) closed. The sidebar must ask again.")

    def write_manifest_text(self) -> str:
        """Render the bundled manifest for this runtime's current add-in base and write it to plugin data."""
        bundled_path = Path(__file__).with_name("sidebar") / "manifest.xml"
        bundled_base = self.cfg.DEFAULT_ADDIN_BASE
        bundled_origin = origin_of(bundled_base)
        try:
            manifest = bundled_path.read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError):
            raise DeskError("Manifest", "The bundled sidebar manifest is unavailable; rebuild the Word sidebar bundle") from None
        domain = f"<AppDomain>{bundled_origin}</AppDomain>"
        if bundled_base not in manifest or domain not in manifest:
            raise DeskError("Manifest", "The bundled sidebar manifest does not match this plugin's default")
        manifest = manifest.replace(bundled_base, self.addin_base)
        manifest = manifest.replace(domain, f"<AppDomain>{self.allowed_origin}</AppDomain>")
        if "{{" in manifest:
            raise DeskError("Manifest", "The bundled sidebar manifest contains unresolved placeholders")

        path = self.store.data_dir / "manifest.xml"
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(manifest, encoding="utf-8")
        except OSError:
            raise DeskError("Manifest", "The Word manifest could not be written to plugin data") from None
        wef_dir = f"~/Library/Containers/com.microsoft.{self.cfg.APP_TITLE}/Data/Documents/wef"
        destination = f"{wef_dir}/{self.cfg.PLUGIN_ID}.xml"
        instruction = f"mkdir -p {wef_dir} && cp {shlex.quote(str(path))} {destination}"
        return f"Manifest written to {path}\nSideload on Mac: {instruction}"

    # -- /sidecar-<app> --
    def slash(self, raw_args: str = "") -> str:
        words = (raw_args or "").split()
        verb = words[0].lower() if words else ""
        code = words[1] if len(words) > 1 else None
        try:
            if self._in_gateway():
                self.maybe_start_quietly()
            if verb == "manifest":
                self.ensure_started()
                return self.write_manifest_text()
            if verb in ("status", "trust"):
                return self.status_text(include_trust=verb == "trust")
            if verb in ("revoke", "forget"):
                return self.revoke_text()
            if verb in ("allow", "approve", "deny", "reject"):
                return self._decided(self.store.decide(verb in ("allow", "approve"), code))
            if verb in ("", "connect"):
                waiting = self.store.pending()
                if not waiting:
                    return (self.status_text() + f"\n\nNo {self.cfg.APP_TITLE} sidebar is waiting. Open Sidecar in "
                            f"{self.cfg.APP_TITLE} and click Connect to Hermes, then run /{self.cfg.COMMAND} again.")
                newest = waiting[0]
                question = (f"Allow the {self.cfg.APP_TITLE} sidebar (code {newest['code']}, from {newest['origin']}) "
                            f"to use Hermes? Allowing a sidebar lets the site at {newest['origin']} use your "
                            f"Hermes agent and all its tools until you revoke. Use a site you control or trust; "
                            f"open only documents you trust. Revoke with /{self.cfg.COMMAND} revoke.")
                choice = self._ask_choice(question, [
                    {"value": "allow", "label": "Allow", "description": f"connect this {self.cfg.APP_TITLE} sidebar"},
                    {"value": "deny", "label": "Deny", "description": "refuse this request"},
                ]) if self._ask_choice else None
                if choice not in ("allow", "deny"):
                    return (f"Sidebar request {newest['code']} is waiting. Reply with /{self.cfg.COMMAND} allow "
                            f"{newest['code']} or /{self.cfg.COMMAND} deny {newest['code']}.")
                return self._decided(self.store.decide(choice == "allow", newest["code"]))
            return (f"Usage: /{self.cfg.COMMAND} [manifest|allow|deny <code>|revoke|status|trust]. With no argument Hermes asks "
                    f"you to Allow or Deny the waiting {self.cfg.APP_TITLE} sidebar.")
        except DeskError as exc:
            return f"Sidecar for {self.cfg.APP_TITLE}: {exc}"

    def maybe_start_quietly(self) -> None:
        try:
            self.ensure_started()
        except DeskError as exc:
            log.warning("desk bridge: %s not started: %s", self.cfg.APP, exc)

    # -- `hermes sidecar-<app>` --
    def cli_setup(self, subparser: Any) -> None:
        subparser.add_argument("action", choices=["manifest", "status", "allow", "deny", "revoke", "trust"],
                               help="write the sideload manifest, show status or trust instructions, allow/deny, or revoke approval")
        subparser.add_argument("code", nargs="?", default=None, help="the code the sidebar shows (allow/deny)")
        subparser.set_defaults(func=self.cli_run)

    def cli_run(self, args: Any) -> int:
        action = getattr(args, "action", "status")
        try:
            if action in ("status", "trust"):
                print(self.status_text(include_trust=action == "trust"))
            elif action == "manifest":
                self.ensure_started()
                print(self.write_manifest_text())
            elif action == "revoke":
                removed = self.store.revoke()
                print(f"Revoked Sidecar for {self.cfg.APP_TITLE}: {removed} remembered approval(s) removed. "
                      f"An open sidebar is disconnected at its next request.")
            else:
                print(self._decided(self.store.decide(action == "allow", getattr(args, "code", None))))
        except DeskError as exc:
            print(f"Sidecar for {self.cfg.APP_TITLE}: {exc}")
            return 1
        return 0
