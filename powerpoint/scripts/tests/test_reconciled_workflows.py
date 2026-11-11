# SPDX-License-Identifier: Apache-2.0
from pathlib import Path
ROOT = next(p for p in Path(__file__).resolve().parents if (p / "plugin.yaml").exists() or (p / "plugins").exists())
def test_research_reconciliation_contract():
    refs = ROOT / 'plugins/atlas-sidecar-powerpoint/skills/office-guidance/references'
    design = (refs / "design-playbook.md").read_text()
    for phrase in ["claim → evidence", "keyboard", "blocker", "microtype", "CSS pixels"]: assert phrase in design
    text = (refs / "powerpoint-workflow.md").read_text()
    assert "15 registered" in text and "duplication" in text
