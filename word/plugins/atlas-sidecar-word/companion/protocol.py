"""Protocol constants, version, auth helpers, and validation for the Word companion RPC."""
from __future__ import annotations

from typing import Optional

import hmac
import uuid

PROTOCOL_VERSION = "1.0"
HOST_PREFIX = "/word"
SERVICE_NAME = "hermes-office-word-companion"
SERVICE_HOST = HOST_PREFIX.removeprefix("/")

# Pane origin gate
REQUIRED_ORIGIN = "https://localhost:3001"

# Body size caps
MAX_INVOKE_BODY = 1 * 1024 * 1024      # 1 MiB
MAX_RESULT_BODY = 10 * 1024 * 1024     # 10 MiB
MAX_SESSION_BODY = 64 * 1024           # 64 KiB for session/heartbeat/etc

# Maximum structural JSON container nesting accepted by the companion.
MAX_JSON_DEPTH = 64

# Session TTL and heartbeat
SESSION_TTL_SECONDS = 60
HEARTBEAT_INTERVAL = 15
CLAIM_LEASE_SECONDS = 20
INVOKE_TIMEOUT_SECONDS = 60

# Bounded identifiers
MAX_DOCUMENT_ID_LEN = 256
MAX_PANE_ID_LEN = 256
MAX_TOOL_NAME_LEN = 256

# Bounded in-memory state
MAX_SESSIONS = 16
MAX_REQUEST_HISTORY = 4
MAX_RETAINED_HISTORY_BYTES = 12 * 1024 * 1024

# Handler concurrency and request deadlines
MAX_HANDLER_THREADS = 20
WATCHDOG_HEADER_TIMEOUT_SECONDS = 10.0

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
    except (ValueError, AttributeError, TypeError):
        return False


def generate_session_key() -> str:
    """Generate a random hex session key (64 hex chars = 32 bytes)."""
    import os
    return os.urandom(32).hex()


def generate_uuid() -> str:
    """Generate a random UUID v4 string."""
    return str(uuid.uuid4())


def validate_bounded_string(value: str, max_len: int, name: str) -> Optional[str]:
    """Validate a bounded identifier string. Returns error message or None."""
    if not isinstance(value, str) or not value:
        return f"{name} must be a non-empty string"
    if len(value) > max_len:
        return f"{name} exceeds maximum length {max_len}"
    # Reject control characters
    if any(ord(c) < 32 for c in value):
        return f"{name} contains control characters"
    return None


# Route classification
PANE_AUTH_ROUTES = {
    ("DELETE", "/word/session"),
    ("POST", "/word/heartbeat"),
    ("GET", "/word/pending"),
    ("POST", "/word/result"),
}

ORIGIN_ONLY_ROUTES = {
    ("POST", "/word/session"),
}

PLUGIN_KEY_ROUTES = {
    ("POST", "/word/invoke"),
    ("POST", "/word/cancel"),
    ("GET", "/word/session"),
    ("GET", "/word/capabilities"),
    ("GET", "/word/metrics"),
    ("GET", "/word/health"),
}

OPTIONS_ROUTES = {
    "/word/session",
    "/word/heartbeat",
    "/word/pending",
    "/word/result",
}
