"""Public documentation must distinguish live evidence from intended support."""
from pathlib import Path
import pytest

ROOT = Path(__file__).resolve().parents[2]


@pytest.mark.parametrize("relative", ["README.md", "plugins/atlas-sidecar-word/README.md"])
def test_readme_disclosures(relative):
    text = (ROOT / relative).read_text()
    assert "Windows desktop Word is not tested" in text
    assert "heading text and a bulleted list written with readback" in text
    assert "no real undo" in text.lower()
    assert "reopen the sidebar" in text
    assert "Keychain Access > login > Certificates" in text
    assert "Sidecar local CA" in text
    assert ".ts.net" not in text
    assert "installer" in text and "addin_base" in text


def test_tool_docs_do_not_claim_undo_or_obsolete_local_agent_modes():
    text = (ROOT / "docs/word-tools.md").read_text()
    assert "no real undo" in text.lower()
    assert "Routine iteration caps" not in text
    assert "scripts/word-no-key-proxy.py" not in text


def test_catalog_bounds_platform_evidence():
    text = (ROOT / "catalog/atlas-sidecar-word.yaml").read_text()
    assert "Windows desktop Word is not tested" in text
    assert "reopen the sidebar" in text
    assert "no real undo" in text.lower()
    assert "installer" in text and "addin_base" in text


def test_docs_name_the_plugin_served_default_and_keep_pages_optional():
    catalog = (ROOT / "catalog/atlas-sidecar-word.yaml").read_text(encoding="utf-8")
    assert "The public default is https://atlasomnia.github.io" not in catalog
    assert "deploy Pages and" not in catalog
    assert "The default is the plugin-served sidebar at https://localhost:<actual-port>/sidecar-word/" in catalog
    for relative in ("README.md", "plugins/atlas-sidecar-word/README.md"):
        text = (ROOT / relative).read_text(encoding="utf-8")
        assert "Optional, only for the hosted origin (the default sidebar is served by the plugin): enable GitHub Pages" in text
