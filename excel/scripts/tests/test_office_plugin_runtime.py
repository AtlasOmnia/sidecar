"""Runtime registration and dispatch tests for the Excel Hermes plugin."""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from test_desk_bridge_security import plugin_env

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
PLUGIN_DIR = REPO / "plugins" / "atlas-sidecar-excel"


class FakeContext:
    def __init__(self) -> None:
        self.tools: list[dict] = []

    def register_tool(self, **kwargs) -> None:
        self.tools.append(kwargs)


def load_plugin():
    import desk_loader, uuid
    return desk_loader.fresh("excel_runtime_" + uuid.uuid4().hex)


def test_registers_every_generated_schema_without_core_override() -> None:
    module = load_plugin()
    context = FakeContext()
    module.register(context)
    schemas = json.loads((PLUGIN_DIR / "tool_schemas.json").read_text(encoding="utf-8"))
    assert len(context.tools) == len(schemas) == 28
    assert {tool["name"] for tool in context.tools} == {schema["name"] for schema in schemas}
    assert all(tool["toolset"] == "sidecar-excel" for tool in context.tools)
    assert all(tool.get("override", False) is False for tool in context.tools)
    assert all(tool.get("check_fn") is None for tool in context.tools)


def test_handler_uses_trusted_session_kwarg_and_returns_companion_json() -> None:
    module = load_plugin()
    calls: list[dict] = []

    def fake_invoke(session_id: str, tool: str, arguments: dict) -> dict:
        calls.append({"session_id": session_id, "tool": tool, "arguments": arguments})
        return {"ok": True, "data": {"cells": [[1]]}}

    module.invoke_bridge = fake_invoke
    context = FakeContext()
    module.register(context)
    read_tool = next(tool for tool in context.tools if tool["name"] == "sidecar_excel_read_range")
    result = json.loads(read_tool["handler"]({
        "sheet": "Sheet1",
        "address": "A1",
    }, session_id="aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"))
    assert result == {"cells": [[1]]}
    assert calls == [{
        "session_id": "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
        "tool": "read_range",
        "arguments": {"sheet": "Sheet1", "address": "A1"},
    }]


def test_handler_rejects_model_supplied_bridge_target_before_invoke() -> None:
    module = load_plugin()
    calls: list[dict] = []

    def fake_invoke(session_id: str, tool: str, arguments: dict) -> dict:
        calls.append({"session_id": session_id, "tool": tool, "arguments": arguments})
        return {"ok": True}

    module.invoke_bridge = fake_invoke
    context = FakeContext()
    module.register(context)
    read_tool = next(tool for tool in context.tools if tool["name"] == "sidecar_excel_read_range")
    result = json.loads(read_tool["handler"]({
        "bridgeSessionId": "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
        "sheet": "Sheet1",
        "address": "A1",
    }, session_id="aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"))

    assert result["ok"] is False
    assert result["error"]["category"] == "Validation"
    assert calls == []


def test_handler_rejects_invalid_trusted_session_before_invoke() -> None:
    module = load_plugin()
    calls: list[dict] = []

    def fake_invoke(session_id: str, tool: str, arguments: dict) -> dict:
        calls.append({"session_id": session_id, "tool": tool, "arguments": arguments})
        return {"ok": True}

    module.invoke_bridge = fake_invoke
    context = FakeContext()
    module.register(context)
    read_tool = next(tool for tool in context.tools if tool["name"] == "sidecar_excel_read_range")

    for kwargs in ({}, {"session_id": "not-a-uuid"}):
        result = json.loads(read_tool["handler"]({
            "sheet": "Sheet1",
            "address": "A1",
        }, **kwargs))
        assert result["ok"] is False
        assert result["error"]["category"] == "Validation"
    assert calls == []


def test_two_live_sessions_enforce_trusted_handler_binding(plugin_env, monkeypatch):
    from test_desk_bridge_security import test_register_starts_bridge_in_gateway_and_tool_reaches_pane
    test_register_starts_bridge_in_gateway_and_tool_reaches_pane(plugin_env, monkeypatch)
