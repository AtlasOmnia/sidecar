# SPDX-License-Identifier: Apache-2.0
# Copyright 2026 Atlas Omnia
"""Bounded stdlib HTTP client from a Word tool call to this process's own desk bridge.

The bridge runs inside this plugin (desk_bridge.py); the key is the bridge's per-process random
plugin key, held in memory. Nothing is read from or written to disk."""
from __future__ import annotations

import http.client
import ipaddress
import json
import socket
import ssl
import time
import urllib.parse
import uuid
from typing import Any

PROTOCOL_VERSION = "1.0"
# The canonical apply_edit_plan envelope is 1,032,976 ASCII bytes at its
# advertised 100 x (find=5000, replace=5000, reason=250) limits. JSON control
# characters can expand to six bytes when escaped, so 8 MiB leaves bounded
# envelope/UTF-8 headroom while remaining a finite transport budget.
MAX_REQUEST_BYTES = 8 * 1024 * 1024
MAX_RESPONSE_BYTES = 8 * 1024 * 1024
INVOKE_HTTP_TIMEOUT_SECONDS = 65.0


class BridgeError(RuntimeError):
    """Safe, category-labelled bridge failure with no secret/body details."""

    def __init__(self, category: str, message: str) -> None:
        super().__init__(message)
        self.category = category


def _validate_base_url(value: str) -> str:
    if not isinstance(value, str):
        raise BridgeError("Unavailable", "The Sidecar Word bridge is not running")
    try:
        parsed = urllib.parse.urlsplit(value)
        hostname = parsed.hostname
        port = parsed.port
        if (parsed.scheme == "https" and hostname == "localhost" and port is not None
                and 1 <= port <= 65535 and not parsed.username and not parsed.password
                and parsed.path in ("", "/") and not parsed.query and not parsed.fragment):
            return f"https://localhost:{port}"
        address = ipaddress.ip_address(hostname or "")
    except (ValueError, TypeError):
        raise BridgeError("Configuration", "Word bridge URL is invalid") from None
    if (parsed.scheme != "http" or hostname != "127.0.0.1"
            or parsed.username or parsed.password
            or parsed.path not in ("", "/") or parsed.query or parsed.fragment):
        raise BridgeError("Configuration", "Word bridge URL is invalid")
    if not address.is_loopback or port is None or not 1 <= port <= 65535:
        raise BridgeError("Configuration", "Word bridge URL is invalid")
    return f"http://127.0.0.1:{port}"


def _read_bounded(response: Any) -> bytes:
    declared = response.headers.get("Content-Length")
    if declared is not None:
        try:
            size = int(declared)
        except ValueError:
            raise BridgeError("Protocol", "Companion response is invalid") from None
        if size < 0 or size > MAX_RESPONSE_BYTES:
            raise BridgeError("Capacity", "Companion response exceeds the allowed size")
    data = response.read(MAX_RESPONSE_BYTES + 1)
    if len(data) > MAX_RESPONSE_BYTES:
        raise BridgeError("Capacity", "Companion response exceeds the allowed size")
    return data


def _decode_response(data: bytes) -> dict[str, Any]:
    try:
        value = json.loads(data.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError):
        raise BridgeError("Protocol", "Companion response is invalid") from None
    if not isinstance(value, dict):
        raise BridgeError("Protocol", "Companion response is invalid")
    return value


def _validate_success_response(response: dict[str, Any], request_id: str) -> dict[str, Any]:
    """Validate the exact success envelope emitted by `/word/invoke`."""
    if set(response) != {"ok", "ticketId", "requestId", "data"}:
        raise BridgeError("Protocol", "Word companion response is invalid")
    if response["ok"] is not True:
        raise BridgeError("Protocol", "Word companion response is invalid")
    ticket_id = response["ticketId"]
    if not isinstance(ticket_id, str):
        raise BridgeError("Protocol", "Word companion response is invalid")
    try:
        uuid.UUID(ticket_id)
    except (ValueError, AttributeError, TypeError):
        raise BridgeError("Protocol", "Word companion response is invalid") from None
    if response["requestId"] != request_id:
        raise BridgeError("Protocol", "Word companion response is invalid")
    return response


def _status_error(status: int) -> BridgeError:
    if status in (401, 403):
        return BridgeError("Authentication", "Word companion authentication failed")
    if status == 400:
        return BridgeError("Validation", "Word companion rejected the request")
    if status == 404:
        return BridgeError("ForeignSession", "Word bridge session is unavailable")
    if status == 409:
        return BridgeError("Protocol", "Word companion rejected a conflicting request")
    if status in (411, 415):
        return BridgeError("Validation", "Word bridge requires a JSON request body")
    if status == 413:
        return BridgeError("Capacity", "Word bridge payload exceeds the allowed size")
    if status == 408:
        return BridgeError("Timeout", "Word bridge request timed out")
    if status == 429:
        return BridgeError("Unavailable", "Word bridge session is busy")
    if status == 499:
        return BridgeError("Cancellation", "Word bridge request was cancelled")
    if status == 504:
        return BridgeError("Timeout", "Word bridge request timed out")
    if status in (502, 503):
        return BridgeError("Unavailable", "Word companion is unavailable")
    return BridgeError("Protocol", "Word companion returned an unexpected response")


def invoke_bridge(
    session_id: str,
    tool: str,
    arguments: dict[str, Any],
    request_id: str | None = None,
    *,
    base_url: str | None,
    key: str,
    ssl_context: ssl.SSLContext | None = None,
) -> dict[str, Any]:
    """Invoke a canonical Word tool on this process's own desk bridge (in-memory key), with one bounded
    same-requestId retry."""
    try:
        uuid.UUID(session_id)
        request_id = request_id or str(uuid.uuid4())
        uuid.UUID(request_id)
    except (ValueError, AttributeError, TypeError):
        raise BridgeError("Validation", "Trusted Word bridge session is missing or malformed") from None
    if not isinstance(tool, str) or not tool or len(tool) > 256:
        raise BridgeError("Validation", "Word bridge tool name is invalid")
    if not isinstance(arguments, dict):
        raise BridgeError("Validation", "Word bridge arguments are invalid")
    try:
        body = json.dumps({
            "protocolVersion": PROTOCOL_VERSION,
            "sessionId": session_id,
            "requestId": request_id,
            "tool": tool,
            "arguments": arguments,
        }, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    except (TypeError, ValueError, UnicodeEncodeError):
        raise BridgeError("Validation", "Word bridge arguments are not JSON-safe") from None
    if len(body) > MAX_REQUEST_BYTES:
        raise BridgeError("Capacity", "Word bridge request exceeds the allowed size")

    base_url = _validate_base_url(base_url)  # type: ignore[arg-type]
    if not isinstance(key, str) or len(key) < 32:
        raise BridgeError("Unavailable", "The Sidecar Word bridge is not running")
    deadline = time.monotonic() + INVOKE_HTTP_TIMEOUT_SECONDS
    last_transport_error: BridgeError | None = None
    for attempt in range(2):
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            break
        parsed_base_url = urllib.parse.urlsplit(base_url)
        host = parsed_base_url.hostname
        port = parsed_base_url.port
        if host is None or port is None:
            raise BridgeError("Configuration", "Word bridge URL is invalid")
        if parsed_base_url.scheme == "https":
            connection = http.client.HTTPSConnection(host, port, timeout=remaining, context=ssl_context)
        else:
            # Standalone companion tests retain their HTTP transport; DeskRuntime always uses TLS.
            connection = http.client.HTTPConnection(host, port, timeout=remaining)
        try:
            connection.request(
                "POST",
                "/word/invoke",
                body=body,
                headers={
                    "Content-Type": "application/json",
                    "Content-Length": str(len(body)),
                    "X-Hermes-Plugin-Key": key,
                    "Connection": "close",
                },
            )
            response = connection.getresponse()
            if response.status != 200:
                raise _status_error(response.status)
            content_type = response.headers.get("Content-Type", "")
            if content_type.split(";", 1)[0].strip().lower() != "application/json":
                raise BridgeError("Protocol", "Word companion response is invalid")
            return _validate_success_response(
                _decode_response(_read_bounded(response)),
                request_id,
            )
        except (socket.timeout, TimeoutError):
            last_transport_error = BridgeError("Timeout", "Word bridge request timed out")
        except OSError:
            last_transport_error = BridgeError("Unavailable", "Word companion is unavailable")
        except http.client.HTTPException:
            raise BridgeError("Protocol", "Word companion response is invalid") from None
        finally:
            connection.close()
        if attempt == 0:
            continue
    raise last_transport_error or BridgeError("Timeout", "Word bridge request timed out")
