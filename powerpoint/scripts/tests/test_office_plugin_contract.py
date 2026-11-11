"""Plugin contract tests — verify PowerPoint plugin schema and naming."""
from __future__ import annotations

import json
import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[2]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

import pytest

PLUGIN_DIR = _REPO_ROOT / "plugins" / "atlas-sidecar-powerpoint"


class TestPluginSchemaContract:
    def test_schema_file_exists(self):
        schema_path = PLUGIN_DIR / "tool_schemas.json"
        assert schema_path.is_file(), f"Missing {schema_path}"

    def test_schema_is_valid_json(self):
        schema_path = PLUGIN_DIR / "tool_schemas.json"
        data = json.loads(schema_path.read_text(encoding="utf-8"))
        assert isinstance(data, list)

    def test_all_tools_have_sidecar_powerpoint_prefix(self):
        schema_path = PLUGIN_DIR / "tool_schemas.json"
        data = json.loads(schema_path.read_text(encoding="utf-8"))
        for schema in data:
            assert schema["name"].startswith("sidecar_powerpoint_"), \
                f"Tool {schema['name']} must have sidecar_powerpoint_ prefix"

    def test_schemas_never_expose_bridge_session_id(self):
        schema_path = PLUGIN_DIR / "tool_schemas.json"
        data = json.loads(schema_path.read_text(encoding="utf-8"))
        for schema in data:
            props = schema.get("parameters", {}).get("properties", {})
            assert "bridgeSessionId" not in props, \
                f"Tool {schema['name']} must not expose bridgeSessionId"

    def test_tool_count_is_15(self):
        schema_path = PLUGIN_DIR / "tool_schemas.json"
        data = json.loads(schema_path.read_text(encoding="utf-8"))
        assert len(data) == 15, f"Expected 15 tools, got {len(data)}"

    def test_plugin_yaml_exists(self):
        yaml_path = PLUGIN_DIR / "plugin.yaml"
        assert yaml_path.is_file()

    def test_plugin_init_exists(self):
        init_path = PLUGIN_DIR / "__init__.py"
        assert init_path.is_file()

    def test_bridge_client_exists(self):
        client_path = PLUGIN_DIR / "bridge_client.py"
        assert client_path.is_file()

    def test_plugin_yaml_lists_15_tools(self):
        yaml_path = PLUGIN_DIR / "plugin.yaml"
        content = yaml_path.read_text(encoding="utf-8")
        tool_lines = [l.strip() for l in content.split("\n") if l.strip().startswith("- sidecar_powerpoint_")]
        assert len(tool_lines) == 15, f"plugin.yaml should list 15 tools, found {len(tool_lines)}"

    def test_plugin_yaml_order_matches_generated_schema_order(self):
        schema_names = [
            schema["name"]
            for schema in json.loads((PLUGIN_DIR / "tool_schemas.json").read_text(encoding="utf-8"))
        ]
        yaml_names = [
            line.strip()[2:].strip()
            for line in (PLUGIN_DIR / "plugin.yaml").read_text(encoding="utf-8").splitlines()
            if line.strip().startswith("- sidecar_powerpoint_")
        ]
        assert len(schema_names) == 15
        assert len(yaml_names) == 15
        assert len(set(yaml_names)) == 15
        assert "bridgeSessionId" not in schema_names
        assert yaml_names == schema_names


def test_unsupported_stubs_are_not_plugin_tools():
    schemas = json.loads((PLUGIN_DIR / "tool_schemas.json").read_text())
    names = {s["name"] for s in schemas}
    assert names.isdisjoint({"sidecar_powerpoint_add_speaker_notes", "sidecar_powerpoint_duplicate_slide", "sidecar_powerpoint_add_image"})
