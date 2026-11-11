"""Word tool allowlist; approved, bound pane sessions authorize reads and writes.

Authorization is checked by invoke_bridge against the live bridge and approval
store, not inferred from model arguments or unavailable user-task/turn kwargs.
"""
from __future__ import annotations

_ALLOWED_TOOLS = frozenset({
    "get_document_overview", "read_selection", "read_document_text", "get_document_structure",
    "apply_edit_plan", "insert_text", "replace_selection", "find_replace", "format_selection",
    "apply_style", "insert_ooxml", "insert_comment",
})


class PolicyError(ValueError):
    """An unknown Word tool cannot be executed."""


def enforce_tool_policy(tool: str) -> None:
    """Fail closed on unknown tools; never pretend to have trusted turn intent."""
    if tool not in _ALLOWED_TOOLS:
        raise PolicyError(f"Unknown Word tool '{tool}' is denied")
