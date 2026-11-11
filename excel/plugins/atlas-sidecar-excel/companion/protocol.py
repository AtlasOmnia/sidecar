"""Protocol constants, version, auth helpers, and validation for the Excel companion RPC."""
from __future__ import annotations

import hmac
import uuid

PROTOCOL_VERSION = "1.0"
WORKBOOK_PLAN_VERSION = 1
HOST_PREFIX = "/excel"

# Pane origin gate
REQUIRED_ORIGIN = "https://localhost:3000"

# Body size caps
MAX_INVOKE_BODY = 1 * 1024 * 1024      # 1 MiB
MAX_RESULT_BODY = 10 * 1024 * 1024     # 10 MiB
MAX_SESSION_BODY = 64 * 1024           # 64 KiB for session/heartbeat/etc

# Session TTL and heartbeat
SESSION_TTL_SECONDS = 60
HEARTBEAT_INTERVAL = 15
CLAIM_LEASE_SECONDS = 20
INVOKE_TIMEOUT_SECONDS = 60
# A plan that crossed execute-start gets its own executing deadline, measured
# from executing_at (not claimed_at). The acknowledgement must arrive inside the
# 20 s claim lease, so claim lease + executing deadline (40 s) stays inside the
# 60 s invoke timeout for a promptly claimed ticket; the invoke deadline remains
# the outer bound either way. Expiry is terminal executing_uncertain and never
# authorizes a new attempt.
PLAN_EXECUTING_LEASE_SECONDS = 20

# Bounded identifiers
MAX_DOCUMENT_ID_LEN = 256
MAX_PANE_ID_LEN = 256
MAX_TOOL_NAME_LEN = 256

# Retained replay history is bounded by both entries and serialized bytes.
MAX_REQUEST_HISTORY = 4
MAX_RETAINED_HISTORY_BYTES = 12 * 1024 * 1024  # 12 MiB per session
# Default per-session bound for BOTH plan fences (turn history and replay
# fences). They are never evicted; a new plan op beyond it is rejected (R3a).
MAX_PLAN_REPLAY_FENCES = 1024

# Maximum structural JSON container nesting accepted by the companion.
MAX_JSON_DEPTH = 128

# R4g-2a trusted workbook plan envelope v2 (wire contract version 2). The
# protocol version above stays "1.0"; mixed plan-envelope versions are refused
# by these exact key sets on every hop.
PLAN_TRUSTED_KEYS = frozenset({"turnOperationId", "workbookPlanFingerprint", "phase", "turnKey"})
PLAN_PHASES = ("inspect", "execute")
PLAN_WIRE_CONTRACT_VERSION = 2

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
    return hmac.compare_digest(a.encode("utf-8", "surrogatepass"), b.encode("utf-8", "surrogatepass"))


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
    ("DELETE", "/excel/session"),
    ("POST", "/excel/heartbeat"),
    ("GET", "/excel/pending"),
    ("POST", "/excel/result"),
    ("POST", "/excel/execute-start"),
}

ORIGIN_ONLY_ROUTES = {
    ("POST", "/excel/session"),
}

PLUGIN_KEY_ROUTES = {
    ("POST", "/excel/invoke"),
    ("POST", "/excel/cancel"),
    ("GET", "/excel/session"),
    ("GET", "/excel/capabilities"),
    ("GET", "/excel/metrics"),
    ("GET", "/excel/health"),
}

OPTIONS_ROUTES = {
    "/excel/session",
    "/excel/heartbeat",
    "/excel/pending",
    "/excel/result",
    "/excel/execute-start",
}
