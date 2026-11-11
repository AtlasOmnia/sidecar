# SPDX-License-Identifier: Apache-2.0
# Copyright 2026 Atlas Omnia
"""Sidecar for Word: Word tools for Hermes through the desktop Word sidebar.

The sidebar (a static HTTPS page loaded into desktop Word) talks to the desk bridge that this plugin runs
on 127.0.0.1 (desk_bridge.py). Tools are named sidecar_word_* so they never collide with other Office plugins.
"""
from __future__ import annotations

import json
import logging
import uuid
from pathlib import Path
from typing import Any, Callable, Optional

from . import desk_app
from .bridge_client import BridgeError, invoke_bridge as _invoke_bridge_http
from .desk_bridge import DeskError, DeskRuntime
from .review_policy import PolicyError, enforce_tool_policy

log = logging.getLogger(__name__)

_TOOLSET = desk_app.TOOLSET
_PREFIX = desk_app.TOOL_PREFIX
_SCHEMA_PATH = Path(__file__).with_name("tool_schemas.json")
_RUNTIME: Optional[DeskRuntime] = None
_RUNTIME_ERROR: Optional[str] = None


def _load_schemas() -> list[dict[str, Any]]:
    try:
        value = json.loads(_SCHEMA_PATH.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise RuntimeError("Word plugin schema artifact is unavailable") from exc
    if not isinstance(value, list):
        raise RuntimeError("Word plugin schema artifact must be an array")
    names: set[str] = set()
    for schema in value:
        if not isinstance(schema, dict):
            raise RuntimeError("Word plugin contains an invalid tool schema")
        name = schema.get("name")
        if not isinstance(name, str) or not name.startswith(_PREFIX) or name in names:
            raise RuntimeError("Word plugin tool names must be unique and namespaced")
        if not isinstance(schema.get("description"), str) or not isinstance(schema.get("parameters"), dict):
            raise RuntimeError("Word plugin schemas must be complete function specs")
        if "bridgeSessionId" in schema["parameters"].get("properties", {}):
            raise RuntimeError("Word plugin schemas must not expose bridgeSessionId")
        names.add(name)
    return value


def _error(exc: BridgeError) -> str:
    return json.dumps({"ok": False, "error": {"category": exc.category, "message": str(exc)}}, separators=(",", ":"))


def invoke_bridge(session_id: str, tool: str, arguments: dict[str, Any]) -> dict[str, Any]:
    """Send one canonical Word tool to the sidebar paired with this chat (this process's bridge)."""
    runtime = _RUNTIME
    if runtime is None:
        raise BridgeError("Configuration", _RUNTIME_ERROR or "Sidecar for Word is not configured")
    runtime.key_provider()  # a tool call runs in the profile's scope: refresh the in-memory API key
    if runtime.bridge.port is None:
        try:
            runtime.ensure_started()  # first use
        except DeskError as exc:
            raise BridgeError(exc.category, str(exc)) from None
    if not runtime.bridge.has_session(session_id) or not runtime.bridge.session_allowed(session_id):
        raise BridgeError("ForeignSession", "No Sidecar for Word sidebar is connected to this chat. Open the "
                          "Sidecar sidebar in Word, connect it, and chat from there.")
    return _invoke_bridge_http(session_id, tool, arguments, base_url=runtime.bridge.base_url,
                               key=runtime.bridge.plugin_key,
                               ssl_context=runtime.bridge.tls.client_context())


def _handler_for(plugin_name: str) -> Callable[..., str]:
    canonical_name = plugin_name.removeprefix(_PREFIX)

    def handler(args: dict[str, Any] | None, **kwargs: Any) -> str:
        if not isinstance(args, dict):
            return _error(BridgeError("Validation", "Word tool arguments must be an object"))
        if "bridgeSessionId" in args:
            return _error(BridgeError("Validation", "bridgeSessionId is trusted request context and cannot be supplied by the model"))
        session_id = kwargs.get("session_id")
        try:
            if not isinstance(session_id, str):
                raise ValueError
            uuid.UUID(session_id)
        except (ValueError, AttributeError, TypeError):
            return _error(BridgeError("Validation", "Trusted Word bridge session is missing or malformed"))
        try:
            # Hermes forwards task_id, session_id and user_task (None on this
            # path), but no turn_id. Pane approval is the write authorization.
            enforce_tool_policy(canonical_name)
        except PolicyError as exc:
            return _error(BridgeError("Policy", str(exc)))
        try:
            response = invoke_bridge(session_id, canonical_name, dict(args))
            if response.get("ok") is not True or "data" not in response:
                raise BridgeError("Protocol", "Word companion response is invalid")
            encoded = json.dumps(response["data"], separators=(",", ":"), ensure_ascii=False)
            if len(encoded.encode("utf-8")) > 8 * 1024 * 1024:
                raise BridgeError("Capacity", "Word tool result exceeds the allowed size")
            return encoded
        except BridgeError as exc:
            return _error(exc)
        except (TypeError, ValueError, UnicodeEncodeError):
            return _error(BridgeError("Protocol", "Word companion response is invalid"))
        except Exception:
            return _error(BridgeError("Protocol", "Word tool execution failed"))

    return handler


def register(ctx: Any) -> None:
    # Native Python plugins require explicit read-only skill registration.
    if hasattr(ctx, "register_skill"):
        ctx.register_skill("office-guidance", Path(__file__).parent / "skills/office-guidance/SKILL.md",
                           description='Use when drafting, editing or styling a menu, brief or document in native Sidecar Word.')
    global _RUNTIME, _RUNTIME_ERROR
    try:
        _RUNTIME = DeskRuntime.from_ctx(ctx, desk_app)
        _RUNTIME_ERROR = None
    except DeskError as exc:
        _RUNTIME, _RUNTIME_ERROR = None, f"Sidecar for Word settings: {exc}"
        log.warning("%s", _RUNTIME_ERROR)
    for schema in _load_schemas():
        name = schema["name"]
        ctx.register_tool(
            name=name,
            toolset=_TOOLSET,
            schema=schema,
            handler=_handler_for(name),
            check_fn=None,
            override=False,
            description=schema["description"],
        )
    runtime = _RUNTIME
    if runtime is None:
        return
    if hasattr(ctx, "register_command"):
        ctx.register_command(
            desk_app.COMMAND, runtime.slash,
            description="Allow, deny or revoke the Sidecar for Word sidebar; status",
            args_hint="[allow|deny <code>|revoke|status]",
        )
    if hasattr(ctx, "register_cli_command"):
        ctx.register_cli_command(
            name=desk_app.COMMAND, help="Sidecar for Word: status, allow/deny a sidebar, revoke",
            setup_fn=runtime.cli_setup, handler_fn=runtime.cli_run,
        )
    if hasattr(ctx, "on_unload"):
        ctx.on_unload(runtime.stop)
    runtime.maybe_autostart()
