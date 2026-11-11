"""Bounded, non-recursive request validation shared by local HTTP transports."""
from __future__ import annotations

MAX_CONTENT_TYPE_HEADER_LENGTH = 4096
_HTTP_TOKEN_CHARS = frozenset(
    "!#$%&'*+-.^_`|~0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
)


def canonical_content_length(values: list[str] | None, cap: int | None) -> tuple[int | None, str | None]:
    """Accept one ASCII canonical decimal; compare to the cap before conversion.

    Non-body routes only need a body-bearing bit, never an unbounded integer.
    """
    values = values or []
    if len(values) > 1:
        return None, "invalid"
    if not values:
        return None, "missing"
    raw = values[0]
    if (not isinstance(raw, str) or not raw
            or any(char < "0" or char > "9" for char in raw)
            or (len(raw) > 1 and raw[0] == "0")):
        return None, "invalid"
    if raw == "0":
        return 0, None
    if cap is None:
        return 1, None
    cap_text = str(cap)
    if len(raw) > len(cap_text) or (len(raw) == len(cap_text) and raw > cap_text):
        return None, "oversized"
    return int(raw, 10), None


def json_depth_exceeds_limit(body: bytes, max_depth: int) -> bool:
    """Bound container depth before JSON parsing, ignoring escaped string bytes."""
    depth = 0
    in_string = False
    escaped = False
    for byte in body:
        if in_string:
            if escaped:
                escaped = False
            elif byte == 92:
                escaped = True
            elif byte == 34:
                in_string = False
        elif byte == 34:
            in_string = True
        elif byte in (91, 123):
            depth += 1
            if depth > max_depth:
                return True
        elif byte in (93, 125):
            depth = max(0, depth - 1)
    return False


def _is_http_token(value: str) -> bool:
    return bool(value) and all(char in _HTTP_TOKEN_CHARS for char in value)


def _is_valid_quoted_parameter(value: str) -> bool:
    if len(value) < 2 or value[0] != '"' or value[-1] != '"':
        return False
    escaped = False
    for char in value[1:-1]:
        if escaped:
            if ord(char) < 32 or ord(char) == 127:
                return False
            escaped = False
        elif char == "\\":
            escaped = True
        elif char == '"' or ord(char) < 32 or ord(char) == 127:
            return False
    return not escaped


def is_valid_json_content_type(value: str) -> bool:
    """Require application/json with bounded, unambiguous valid parameters."""
    if len(value) > MAX_CONTENT_TYPE_HEADER_LENGTH:
        return False
    parts = []
    start = 0
    in_quotes = False
    escaped = False
    for index, char in enumerate(value):
        if in_quotes:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_quotes = False
        elif char == '"':
            in_quotes = True
        elif char == ";":
            parts.append(value[start:index])
            start = index + 1
    if in_quotes or escaped:
        return False
    parts.append(value[start:])
    if parts[0].strip(" \t").lower() != "application/json":
        return False
    seen = set()
    for raw in parts[1:]:
        name, separator, parameter = raw.strip(" \t").partition("=")
        name = name.strip(" \t")
        parameter = parameter.strip(" \t")
        if not separator or not _is_http_token(name) or name.lower() in seen:
            return False
        seen.add(name.lower())
        if not (_is_http_token(parameter) or _is_valid_quoted_parameter(parameter)):
            return False
    return True
