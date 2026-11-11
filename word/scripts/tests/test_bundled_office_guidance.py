# SPDX-License-Identifier: Apache-2.0
"""Guidance discovery contracts; not live model or native-host acceptance."""
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[2]
APP = ROOT.name if ROOT.name in ("excel", "word", "powerpoint") else next((p.name.removeprefix("atlas-sidecar-") for p in (ROOT / "plugins").iterdir() if p.is_dir() and p.name.startswith("atlas-sidecar-")))
PLUGIN = ROOT / "plugins" / ("atlas-sidecar-" + APP)

def test_real_host_load_discovers_entry_and_portable_references(tmp_path):
    target = tmp_path / "relocated" / PLUGIN.name
    shutil.copytree(PLUGIN, target, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    code = r'''import json, sys
from pathlib import Path
from hermes_cli.plugins import PluginManager
from hermes_cli.plugins_manifest import parse_manifest_file
from tools.skills_tool_plugin import _plugin_skill_linked_files, _serve_skill_file
p = Path(sys.argv[1]); app = sys.argv[2]
m = PluginManager()
manifest = parse_manifest_file(p / "plugin.yaml", p, "user", "")
m._load_plugin(manifest)
q = manifest.name + ":office-guidance"
entry = m.find_plugin_skill(q)
assert entry is not None, (q, m.list_plugin_skill_metadata())
metadata = next(x for x in m.list_plugin_skill_metadata() if x["name"] == q)
expected = {"excel": ["spreadsheet", "budget", "chart"], "word": ["menu", "brief", "document"], "powerpoint": ["deck", "slides", "presentation"]}[app]
for word in expected: assert word in metadata["description"].lower(), metadata
links = _plugin_skill_linked_files(entry.parent)
refs = links["references"]
assert len(refs) >= 3, refs
text = entry.read_text()
for ref in refs:
    assert "](" + ref + ")" in text, ref
    reply = json.loads(_serve_skill_file(entry.parent, ref, q))
    assert reply["success"], reply
    assert reply["content"] == (entry.parent / ref).read_text()
print(json.dumps({"qualified": q, "metadata": metadata, "references": refs}, default=str))
m.unload()
'''
    env = dict(os.environ, HERMES_HOME=str(tmp_path / "home"), KANBAN_HOME=str(tmp_path / "kanban"), GATEWAY_LOCK_DIR=str(tmp_path / "locks"))
    env.pop("_HERMES_GATEWAY", None)
    result = subprocess.run([sys.executable, "-c", code, str(target), APP], env=env, text=True, capture_output=True, timeout=90)
    assert result.returncode == 0, result.stdout + result.stderr

def test_guidance_contract_privacy_and_registered_names():
    skill = PLUGIN / "skills/office-guidance"
    texts = {p.name: p.read_text() for p in skill.rglob("*.md")}
    assert {"SKILL.md", "design-playbook.md", "provenance.md", APP + "-workflow.md"} <= texts.keys()
    text = "\n".join(texts.values())
    assert not re.search(r"/Users/|[A-Z]:\\|[\w.-]+\.ts\.net|BEGIN .*PRIVATE KEY", text)
    names = {x["name"] for x in json.loads((PLUGIN / "tool_schemas.json").read_text())}
    assert set(re.findall(r"\bsidecar_" + APP + r"_[a-z_]+\b", text)) <= names
    for phrase in ["typography", "grid", "spacing", "palette", "contrast", "hierarchy", "readback", "external workflow", "not live"]:
        assert phrase in text, phrase
    if APP == "excel":
        for phrase in ["formulas", "currency", "assumptions", "reconciliation", "source", "trendlines", "axes"]: assert phrase in text, phrase
    if APP == "word":
        for phrase in ["targetText", "paragraphs", "Flat OPC", "GeneralException", "FIRST failed OOXML", "nothing changed", "HTML", "styles.xml", "escape"]: assert phrase in text, phrase
    if APP == "powerpoint":
        for phrase in ["shapeId", "zero-based", "font", "layout", "clipping", "Unsupported"]: assert phrase in text, phrase
