"""Bounded stdlib HTTP client for the Excel companion bridge."""
from __future__ import annotations

import http.client
import ipaddress
import json
import ssl
import urllib.error
import urllib.parse
import urllib.request
import uuid
from typing import Any

PROTOCOL_VERSION = "1.0"
DEFAULT_BASE_URL = "http://127.0.0.1:8650"
INVOKE_HTTP_TIMEOUT_SECONDS = 65
MAX_CONFIG_BYTES = 64 * 1024
MAX_RESPONSE_BYTES = 10 * 1024 * 1024 + 64 * 1024
PLAN_TOOL = "execute_plan"
PLAN_UNCERTAIN_CATEGORY = "ExecutionUncertain"
# Content-free and final: an uncertain plan is never retried by this client.
PLAN_UNCERTAIN_MESSAGE = (
    "Workbook plan execution outcome is uncertain; it was not retried "
    "and must not be re-run automatically"
)


class BridgeError(RuntimeError):
    def __init__(self, category: str, message: str) -> None:
        super().__init__(message)
        self.category = category


def _validate_base_url(value: str) -> str:
    parsed = urllib.parse.urlsplit(value)
    if (parsed.scheme == "https" and parsed.hostname == "localhost" and parsed.port is not None
            and not parsed.username and not parsed.password and parsed.path in ("", "/")
            and not parsed.query and not parsed.fragment):
        return value.rstrip("/")
    if parsed.scheme != "http" or parsed.username or parsed.password:
        raise BridgeError("Configuration", "Excel bridge base URL must be unauthenticated loopback HTTP")
    if parsed.path not in ("", "/") or parsed.query or parsed.fragment or parsed.port is None:
        raise BridgeError("Configuration", "Excel bridge base URL must contain only loopback host and port")
    try:
        address = ipaddress.ip_address(parsed.hostname or "")
    except ValueError as exc:
        raise BridgeError("Configuration", "Excel bridge base URL must use a loopback IP address") from exc
    if not address.is_loopback or parsed.hostname != "127.0.0.1":
        raise BridgeError("Configuration", "Excel bridge base URL must use a loopback IP address")
    return value.rstrip("/")


def _connection() -> tuple[str, str]:
    raise BridgeError("Unavailable", "The Sidecar Excel bridge is not running")


def _validate_key(key: str) -> str:
    if not isinstance(key, str) or len(key) < 32:
        raise BridgeError("Unavailable", "The Sidecar Excel bridge is not running")
    return key


def _read_bounded(response: Any) -> bytes:
    content_length = response.headers.get("Content-Length")
    if content_length is not None:
        try:
            declared = int(content_length)
        except ValueError as exc:
            raise BridgeError("Protocol", "Companion returned invalid Content-Length") from exc
        if declared < 0 or declared > MAX_RESPONSE_BYTES:
            raise BridgeError("Capacity", "Companion response exceeds the allowed size")
    data = response.read(MAX_RESPONSE_BYTES + 1)
    if len(data) > MAX_RESPONSE_BYTES:
        raise BridgeError("Capacity", "Companion response exceeds the allowed size")
    return data


def _decode_json(data: bytes) -> dict[str, Any]:
    try:
        value = json.loads(data.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise BridgeError("Protocol", "Companion returned invalid JSON") from exc
    if not isinstance(value, dict):
        raise BridgeError("Protocol", "Companion returned a non-object response")
    return value


def invoke_bridge(
    session_id: str,
    tool: str,
    arguments: dict[str, Any],
    *,
    trusted: Any = None,
    base_url: str | None = None,
    key: str | None = None,
    ssl_context: ssl.SSLContext | None = None,
) -> dict[str, Any]:
    """Invoke one canonical Excel tool through an exact companion session."""
    if type(session_id) is not str:
        raise BridgeError("Validation", "bridgeSessionId must be a UUID")
    if len(session_id) > 256:
        raise BridgeError("Capacity", "bridgeSessionId exceeds the allowed size")
    try:
        uuid.UUID(session_id)
    except (ValueError, AttributeError, TypeError):
        session_valid = False
    else:
        session_valid = True
    if not session_valid:
        raise BridgeError("Validation", "bridgeSessionId must be a UUID")
    if type(tool) is not str:
        raise BridgeError("Validation", "Excel bridge tool name is invalid")
    if len(tool) == 0:
        raise BridgeError("Validation", "Excel bridge tool name is invalid")
    if len(tool) > 256:
        raise BridgeError("Capacity", "Excel bridge tool name exceeds the allowed size")
    if type(arguments) is not dict:
        raise BridgeError("Validation", "Excel bridge arguments must be an object")

    request_id = str(uuid.uuid4())
    wire_trusted: dict[str, str] | None = None
    if tool == "execute_plan":
        if trusted is None or type(trusted) is not dict:
            raise BridgeError("Validation", "Trusted workbook plan metadata is invalid")
        expected_keys = {
            "turn_operation_id",
            "request_id",
            "workbook_plan_fingerprint",
            "phase",
            "turn_key",
        }
        if (
            len(trusted) != 5
            or any(type(key) is not str for key in trusted)
            or set(trusted) != expected_keys
        ):
            raise BridgeError("Validation", "Trusted workbook plan metadata is invalid")

        turn_operation_id = trusted["turn_operation_id"]
        trusted_request_id = trusted["request_id"]
        workbook_plan_fingerprint = trusted["workbook_plan_fingerprint"]
        phase = trusted["phase"]
        turn_key = trusted["turn_key"]
        if not all(
            type(value) is str
            for value in (
                turn_operation_id,
                trusted_request_id,
                workbook_plan_fingerprint,
                phase,
                turn_key,
            )
        ):
            raise BridgeError("Validation", "Trusted workbook plan metadata is invalid")

        bounded_values = (
            (turn_operation_id, 256),
            (trusted_request_id, 256),
            (workbook_plan_fingerprint, 64),
            (phase, 256),
            (turn_key, 256),
        )
        for value, max_bytes in bounded_values:
            if len(value) > max_bytes:
                raise BridgeError(
                    "Capacity",
                    "Trusted workbook plan metadata exceeds the allowed size",
                )
            try:
                encoded = value.encode("utf-8")
            except UnicodeEncodeError:
                encoded = None
            if encoded is None:
                raise BridgeError(
                    "Validation",
                    "Trusted workbook plan metadata is invalid",
                )
            if len(encoded) > max_bytes:
                raise BridgeError(
                    "Capacity",
                    "Trusted workbook plan metadata exceeds the allowed size",
                )

        try:
            uuid.UUID(turn_operation_id)
        except (ValueError, AttributeError, TypeError):
            turn_valid = False
        else:
            turn_valid = True
        try:
            uuid.UUID(trusted_request_id)
        except (ValueError, AttributeError, TypeError):
            request_valid = False
        else:
            request_valid = True
        fingerprint_valid = (
            len(workbook_plan_fingerprint) == 64
            and all(char in "0123456789abcdef" for char in workbook_plan_fingerprint)
        )
        try:
            turn_key_valid = str(uuid.UUID(turn_key)) == turn_key
        except (ValueError, AttributeError, TypeError):
            turn_key_valid = False
        if (
            not turn_valid
            or not request_valid
            or turn_operation_id != trusted_request_id
            or not fingerprint_valid
            or phase not in ("inspect", "execute")
            or not turn_key_valid
            or turn_key == turn_operation_id
        ):
            raise BridgeError("Validation", "Trusted workbook plan metadata is invalid")

        envelope = json.dumps(
            {
                "turnOperationId": turn_operation_id,
                "workbookPlanFingerprint": workbook_plan_fingerprint,
                "phase": phase,
                "turnKey": turn_key,
            },
            separators=(",", ":"),
            ensure_ascii=False,
        ).encode("utf-8")
        if len(envelope) > 256:
            raise BridgeError(
                "Capacity",
                "Trusted workbook plan metadata exceeds the allowed size",
            )
        request_id = trusted_request_id
        wire_trusted = {
            "turnOperationId": turn_operation_id,
            "workbookPlanFingerprint": workbook_plan_fingerprint,
            "phase": phase,
            "turnKey": turn_key,
        }
    elif trusted is not None:
        raise BridgeError("Validation", "Trusted metadata is only valid for the workbook plan route")

    if base_url is None and key is None:
        base_url, key = _connection()
    key = _validate_key(key)
    if not isinstance(base_url, str):
        raise BridgeError("Unavailable", "The Sidecar Excel bridge is not running")
    base_url = _validate_base_url(base_url)
    payload: dict[str, Any] = {
        "protocolVersion": PROTOCOL_VERSION,
        "sessionId": session_id,
        "requestId": request_id,
        "tool": tool,
        "arguments": arguments,
    }
    if wire_trusted is not None:
        payload["trusted"] = wire_trusted
    body = json.dumps(payload, separators=(",", ":")).encode("utf-8")
    request = urllib.request.Request(
        f"{base_url}/excel/invoke",
        data=body,
        method="POST",
        headers={
            "Content-Type": "application/json",
            "Content-Length": str(len(body)),
            "X-Hermes-Plugin-Key": key,
        },
    )
    if tool == PLAN_TOOL:
        return _invoke_plan_once(request, ssl_context=ssl_context)
    try:
        with urllib.request.urlopen(request, timeout=INVOKE_HTTP_TIMEOUT_SECONDS,
                                    **({"context": ssl_context} if ssl_context is not None else {})) as response:
            return _decode_json(_read_bounded(response))
    except urllib.error.HTTPError as exc:
        data = _read_bounded(exc)
        try:
            detail = _decode_json(data)
        except BridgeError:
            detail = {}
        category = str(detail.get("category") or ("Timeout" if exc.code == 504 else "Bridge"))
        message = str(detail.get("error") or f"Excel companion returned HTTP {exc.code}")
        raise BridgeError(category, message) from None
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise BridgeError("Unavailable", "Excel companion is unavailable") from exc


def _read_plan_detail(response: Any, *, declared_oversize_is_capacity: bool) -> dict[str, Any] | None:
    """Return a delivered plan response body as a JSON object, or ``None``.

    ``None`` means the body is not provably complete and well formed: a read
    error, a short read against Content-Length, an unparseable Content-Length,
    a body that runs past the bound after the headers were delivered, invalid
    UTF-8/JSON, or a non-object. The companion may already have executed in
    every one of those cases. Only a declared oversize length on a success
    response is rejected as ``Capacity`` before anything is read.
    """
    content_length = response.headers.get("Content-Length")
    declared: int | None = None
    if content_length is not None:
        try:
            declared = int(content_length)
        except ValueError:
            return None
        if declared < 0:
            return None
        if declared > MAX_RESPONSE_BYTES:
            if declared_oversize_is_capacity:
                raise BridgeError("Capacity", "Companion response exceeds the allowed size")
            return None
    try:
        data = response.read(MAX_RESPONSE_BYTES + 1)
    except (TimeoutError, OSError, ValueError, http.client.HTTPException):
        return None
    if len(data) > MAX_RESPONSE_BYTES or (declared is not None and len(data) != declared):
        return None
    try:
        value = json.loads(data.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError):
        return None
    return value if isinstance(value, dict) else None


def _invoke_plan_once(request: urllib.request.Request, *, ssl_context: ssl.SSLContext | None = None) -> dict[str, Any]:
    """Send one workbook plan exactly once and preserve uncertainty.

    urllib wraps only connect/send failures in ``URLError``; those mean the
    companion never received a complete request, so the plan never started.
    Any loss after the request was sent (reset, timeout, truncated response)
    is reported as ``ExecutionUncertain`` and is never retried. That includes
    a delivered 2xx or 5xx response whose body cannot be fully read or
    decoded. Every complete, decodable 5xx is ``ExecutionUncertain`` too (a
    generic companion 500 can follow an execute-start), except the
    companion's own never-started shape: 504 with category ``Timeout`` and
    state ``not_started``. A 4xx keeps its status-derived category: the
    companion issues plan 4xx responses only before a ticket is created or
    reattached, or as 499 for a ticket cancelled before execute-start.
    """
    try:
        with urllib.request.urlopen(request, timeout=INVOKE_HTTP_TIMEOUT_SECONDS,
                                    **({"context": ssl_context} if ssl_context is not None else {})) as response:
            detail = _read_plan_detail(response, declared_oversize_is_capacity=True)
            if detail is None:
                raise BridgeError(PLAN_UNCERTAIN_CATEGORY, PLAN_UNCERTAIN_MESSAGE) from None
            return detail
    except urllib.error.HTTPError as exc:
        detail = _read_plan_detail(exc, declared_oversize_is_capacity=False)
        if detail is None:
            if exc.code >= 500:
                raise BridgeError(PLAN_UNCERTAIN_CATEGORY, PLAN_UNCERTAIN_MESSAGE) from None
            category = "Cancellation" if exc.code == 499 else "Bridge"
            raise BridgeError(category, f"Excel companion returned HTTP {exc.code}") from None
        category = str(detail.get("category") or "")
        if exc.code >= 500 and not (
            exc.code == 504 and detail.get("category") == "Timeout" and detail.get("state") == "not_started"
        ):
            raise BridgeError(PLAN_UNCERTAIN_CATEGORY, PLAN_UNCERTAIN_MESSAGE) from None
        if category == PLAN_UNCERTAIN_CATEGORY or (exc.code == 504 and not category):
            raise BridgeError(PLAN_UNCERTAIN_CATEGORY, PLAN_UNCERTAIN_MESSAGE) from None
        category = category or ("Timeout" if exc.code == 504 else "Bridge")
        message = str(detail.get("error") or f"Excel companion returned HTTP {exc.code}")
        raise BridgeError(category, message) from None
    except urllib.error.URLError as exc:
        raise BridgeError("Unavailable", "Excel companion is unavailable") from exc
    except (TimeoutError, OSError, http.client.HTTPException):
        raise BridgeError(PLAN_UNCERTAIN_CATEGORY, PLAN_UNCERTAIN_MESSAGE) from None
