"""Contract tests for generated Word Hermes plugin schemas."""
from __future__ import annotations

import json
import subprocess
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SCHEMAS = REPO / "plugins" / "atlas-sidecar-word" / "tool_schemas.json"
EXPECTED_TOOLS = 12


def test_generated_schema_artifact_is_current() -> None:
    completed = subprocess.run(
        ["node", "scripts/generate-office-plugin-schemas.mjs", "--check"],
        cwd=REPO,
        text=True,
        capture_output=True,
        check=False,
    )
    assert completed.returncode == 0, completed.stdout + completed.stderr


def test_plugin_schemas_are_namespaced_and_do_not_expose_bridge_session() -> None:
    schemas = json.loads(SCHEMAS.read_text(encoding="utf-8"))
    assert len(schemas) == EXPECTED_TOOLS
    names = [schema["name"] for schema in schemas]
    assert len(names) == len(set(names))
    assert all(name.startswith("sidecar_word_") for name in names)
    for schema in schemas:
        parameters = schema["parameters"]
        assert parameters["additionalProperties"] is False
        assert "bridgeSessionId" not in parameters.get("properties", {})
        assert "bridgeSessionId" not in parameters.get("required", [])


def test_plugin_manifest_matches_generated_tool_names() -> None:
    schemas = json.loads(SCHEMAS.read_text(encoding="utf-8"))
    manifest_lines = (SCHEMAS.parent / "plugin.yaml").read_text(encoding="utf-8").splitlines()
    declared = [line.strip()[2:] for line in manifest_lines if line.startswith("  - sidecar_")]
    assert declared == [schema["name"] for schema in schemas]


if __name__ == "__main__":
    tests = [(name, value) for name, value in list(globals().items())
             if name.startswith("test_") and callable(value)]
    passed = 0
    for name, test in tests:
        test()
        passed += 1
        print(f"  PASS {name}")
    print(f"\n{passed} passed, 0 failed, {len(tests)} total")
