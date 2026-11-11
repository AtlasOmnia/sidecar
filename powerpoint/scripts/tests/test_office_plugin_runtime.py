"""Plugin runtime tests — verify PowerPoint plugin registration simulation."""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[2]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

import pytest

PLUGIN_DIR = _REPO_ROOT / "plugins" / "atlas-sidecar-powerpoint"


def _load_plugin_module():
    import desk_loader
    return desk_loader.fresh("sidecar_powerpoint_runtime_test")


def _load_bridge_module():
    return _load_plugin_module().bridge_client


class MockPluginContext:
    """Simulate the Hermes plugin registration context."""

    def __init__(self):
        self.registered_tools: list[dict] = []

    def register_tool(self, name, toolset, schema, handler, check_fn=None,
                      override=False, description=""):
        self.registered_tools.append({
            "name": name,
            "toolset": toolset,
            "schema": schema,
            "handler": handler,
            "check_fn": check_fn,
            "override": override,
            "description": description,
        })


class TestPluginRegistration:
    def test_register_all_tools(self):
        plugin = _load_plugin_module()
        schemas = plugin._load_schemas()
        assert len(schemas) == 15

        ctx = MockPluginContext()
        plugin.register(ctx)

        assert len(ctx.registered_tools) == 15
        for tool in ctx.registered_tools:
            assert tool["name"].startswith("sidecar_powerpoint_")
            assert tool["toolset"] == "sidecar-powerpoint"
            assert tool["override"] is False

    def test_handler_removes_prefix(self):
        plugin = _load_plugin_module()
        handler = plugin._handler_for("sidecar_powerpoint_get_presentation_overview")
        assert callable(handler)

    def test_handler_routes_only_from_trusted_session_kwarg(self, monkeypatch):
        plugin = _load_plugin_module()
        calls = []
        monkeypatch.setattr(plugin, "invoke_bridge", lambda session_id, tool, arguments: calls.append((session_id, tool, arguments)) or {"ok": True, "data": {"ok": True}})
        handler = plugin._handler_for("sidecar_powerpoint_get_presentation_overview")
        result = json.loads(handler({}, session_id="550e8400-e29b-41d4-a716-446655440000"))
        assert result == {"ok": True}
        assert calls == [("550e8400-e29b-41d4-a716-446655440000", "get_presentation_overview", {})]

    def test_handler_rejects_missing_or_caller_supplied_session_identity(self, monkeypatch):
        plugin = _load_plugin_module()
        monkeypatch.setattr(plugin, "invoke_bridge", lambda *args: pytest.fail("untrusted identity reached bridge"))
        handler = plugin._handler_for("sidecar_powerpoint_get_presentation_overview")
        assert json.loads(handler({})) ["ok"] is False
        assert json.loads(handler({"bridgeSessionId": "550e8400-e29b-41d4-a716-446655440000"})) ["ok"] is False

    def test_load_schemas_rejects_duplicate_names(self):
        plugin = _load_plugin_module()
        schemas = plugin._load_schemas()
        names = [s["name"] for s in schemas]
        assert len(names) == len(set(names)), "Duplicate tool names in schemas"

    def test_load_schemas_rejects_invalid_prefix(self):
        plugin = _load_plugin_module()
        schemas = plugin._load_schemas()
        for schema in schemas:
            assert schema["name"].startswith("sidecar_powerpoint_"), \
                f"Invalid name: {schema['name']}"


class TestBridgeClient:
    def test_url_requires_this_process_bridge(self):
        bc = _load_bridge_module()
        with pytest.raises(bc.BridgeError):
            bc._validate_base_url("https://example.com")

    def test_protocol_version(self):
        bc = _load_bridge_module()
        assert bc.PROTOCOL_VERSION == "1.0"

    def test_bridge_error_str(self):
        bc = _load_bridge_module()
        err = bc.BridgeError("Validation", "test error")
        assert str(err) == "test error"
        assert err.category == "Validation"
