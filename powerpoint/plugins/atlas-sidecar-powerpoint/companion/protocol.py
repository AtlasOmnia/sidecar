"""Protocol constants, version, auth helpers, and validation for the PowerPoint companion RPC."""
from __future__ import annotations

import hmac
import uuid

PROTOCOL_VERSION = "1.0"
HOST_PREFIX = "/powerpoint"
SERVICE_NAME = "hermes-office-powerpoint-companion"
SERVICE_HOST = "powerpoint"

# Pane origin gate
REQUIRED_ORIGIN = "https://localhost:3002"

# Body size caps (PowerPoint tools exchange text/metadata, never binary slide assets)
MAX_INVOKE_BODY = 256 * 1024             # 256 KiB
MAX_RESULT_BODY = 2 * 1024 * 1024        # 2 MiB
MAX_SESSION_BODY = 64 * 1024             # 64 KiB for session/heartbeat/etc
MAX_JSON_DEPTH = 64

# Retained replay history is bounded by both entries and serialized bytes.
MAX_REQUEST_HISTORY = 8
MAX_RETAINED_HISTORY_BYTES = 4 * 1024 * 1024  # 4 MiB per session

# Session TTL and heartbeat
SESSION_TTL_SECONDS = 60
HEARTBEAT_INTERVAL = 15
CLAIM_LEASE_SECONDS = 20
INVOKE_TIMEOUT_SECONDS = 60

# Bounded identifiers
MAX_DOCUMENT_ID_LEN = 256
MAX_PANE_ID_LEN = 256
MAX_TOOL_NAME_LEN = 256

# Max sessions (bounded concurrency and retained-memory envelope)
MAX_SESSIONS = 16

# Handler concurrency
MAX_HANDLER_THREADS = 20

# CORS headers for pane routes
CORS_HEADERS = [
    "Origin",
    "Content-Type",
    "X-Hermes-Session-Key",
    "X-Hermes-Session-Id",
    "X-Hermes-Pane-Id",
]
CORS_METHODS = "GET, POST, DELETE, OPTIONS"


def constant_time_compare(a: str, b: str) -> bool:
    """Constant-time string comparison to prevent timing leaks."""
    return hmac.compare_digest(a.encode("utf-8"), b.encode("utf-8"))


def validate_uuid(value: str) -> bool:
    """Return True if value is a valid UUID string."""
    try:
        uuid.UUID(value)
        return True
    except (ValueError, AttributeError):
        return False


def generate_session_key() -> str:
    """Generate a random hex session key (64 hex chars = 32 bytes)."""
    import os
    return os.urandom(32).hex()


def generate_uuid() -> str:
    """Generate a random UUID v4 string."""
    return str(uuid.uuid4())


def validate_bounded_string(value: str, max_len: int, name: str) -> str | None:
    """Validate a bounded identifier string. Returns error message or None."""
    if not value or not isinstance(value, str):
        return f"{name} must be a non-empty string"
    if len(value) > max_len:
        return f"{name} exceeds maximum length {max_len}"
    # Reject control characters
    if any(ord(c) < 32 for c in value):
        return f"{name} contains control characters"
    return None


# Route classification
PANE_AUTH_ROUTES = {
    ("DELETE", "/powerpoint/session"),
    ("POST", "/powerpoint/heartbeat"),
    ("GET", "/powerpoint/pending"),
    ("POST", "/powerpoint/result"),
}

ORIGIN_ONLY_ROUTES = {
    ("POST", "/powerpoint/session"),
}

PLUGIN_KEY_ROUTES = {
    ("POST", "/powerpoint/invoke"),
    ("POST", "/powerpoint/cancel"),
    ("GET", "/powerpoint/session"),
    ("GET", "/powerpoint/capabilities"),
    ("GET", "/powerpoint/metrics"),
    ("GET", "/powerpoint/health"),
}

OPTIONS_ROUTES = {
    "/powerpoint/session",
    "/powerpoint/heartbeat",
    "/powerpoint/pending",
    "/powerpoint/result",
}
