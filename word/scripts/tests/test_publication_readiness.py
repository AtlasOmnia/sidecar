"""Publication regressions; no external services or OS trust changes."""
from pathlib import Path
from types import SimpleNamespace
import pytest
from test_desk_bridge_security import (plugin_env, FakeCtx, ADDIN_BASE, ORIGIN, _free_ports, make_runtime)

ROOT = Path(__file__).resolve().parents[2]


def test_allow_question_discloses_full_agent_authority(tmp_path):
    rt = make_runtime(tmp_path)
    rt.store.create_request(ORIGIN)
    seen = []
    rt._ask_choice = lambda question, choices: seen.append(question)
    rt.slash()
    assert f"site at {ORIGIN}" in seen[0]
    assert "all its tools until you revoke" in seen[0]
    assert "Use a site you control or trust; open only documents you trust." in seen[0]


@pytest.mark.parametrize("platform", ["darwin", "win32"])
def test_trust_disclosure_describes_policy_breadth(tmp_path, platform):
    rt = make_runtime(tmp_path)
    assert "For TLS it can only vouch for localhost. The trust command does not limit the certificate to TLS, so protect the key file and remove the trust when you stop using Sidecar." in rt.bridge.tls.trust_instructions(platform)


def test_loopback_default_and_optional_publication_workflow_are_explicit():
    from desk_plugin import desk_app
    assert desk_app.DEFAULT_ADDIN_BASE == f"https://localhost:{desk_app.PORTS[0]}/{desk_app.COMMAND}/"
    manifest = (ROOT / "plugins/atlas-sidecar-word/plugin.yaml").read_text()
    catalog = (ROOT / "catalog/atlas-sidecar-word.yaml").read_text()
    workflow = (ROOT / ".github/workflows/pages.yml").read_text()
    assert f'default: "{desk_app.DEFAULT_ADDIN_BASE}"' in manifest
    assert "plugin serves its bundled sidebar" in manifest
    assert "self-hosted" in manifest and "addin_base" in manifest
    for text in (manifest, catalog):
        assert "platforms: [macos]  # Windows desktop Word untested" in text
    assert "repo: https://github.com/AtlasOmnia/sidecar-word" in catalog
    assert "filled by the publish runbook" in catalog
    assert "npm ci && npm run build:static -- --base https://atlasomnia.github.io/sidecar-word/" in workflow
    assert "actions/deploy-pages@" in workflow


@pytest.mark.parametrize("command", ["sidecar-word", "sidecar-excel", "sidecar-powerpoint"])
def test_cli_inherited_gateway_never_listens(plugin_env, monkeypatch, command):
    import sys
    monkeypatch.setenv("_HERMES_GATEWAY", "1")
    monkeypatch.setattr(sys, "argv", ["hermes", command, "status"])
    monkeypatch.setattr(plugin_env.desk_app, "PORTS", _free_ports(2))
    plugin_env.register(FakeCtx({"addin_base": ADDIN_BASE}))
    assert plugin_env._RUNTIME.bridge.port is None
    assert not plugin_env._RUNTIME.bridge.tls.ca_path.exists()


def test_status_quiet_explicit_trust_available(tmp_path, capsys):
    rt = make_runtime(tmp_path)
    rt.bridge.tls.ensure()
    assert rt.cli_run(SimpleNamespace(action="status")) == 0
    assert "add-trusted-cert" not in capsys.readouterr().out
    assert "add-trusted-cert" not in rt.slash("status")
    assert rt.cli_run(SimpleNamespace(action="trust")) == 0
    assert "add-trusted-cert" in capsys.readouterr().out
    assert "add-trusted-cert" in rt.slash("trust")


@pytest.mark.parametrize("name", ["LICENSE", "NOTICE", "README.md"])
def test_installed_copies_match_root(name):
    assert (ROOT / name).read_bytes() == (ROOT / "plugins/atlas-sidecar-word" / name).read_bytes()


def test_new_word_ca_name_and_legacy_name_load(tmp_path):
    import re
    from datetime import datetime, timezone
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.x509.oid import NameOID
    from desk_plugin.desk_tls import LocalTLS
    tls = LocalTLS(tmp_path)
    tls.ensure()
    ca = x509.load_pem_x509_certificate(tls.ca_path.read_bytes())
    assert re.fullmatch(r"Sidecar for Word local CA [0-9a-f]{16}", ca.subject.get_attributes_for_oid(NameOID.COMMON_NAME)[0].value)
    key = serialization.load_pem_private_key(tls.ca_key_path.read_bytes(), password=None)
    leaf_key = serialization.load_pem_private_key(tls.key_path.read_bytes(), password=None)
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "Sidecar local CA 0123456789abcdef")])
    builder = (x509.CertificateBuilder().subject_name(name).issuer_name(name).public_key(key.public_key())
        .serial_number(ca.serial_number).not_valid_before(ca.not_valid_before_utc).not_valid_after(ca.not_valid_after_utc))
    for ext in ca.extensions:
        builder = builder.add_extension(ext.value, ext.critical)
    old = builder.sign(key, hashes.SHA256())
    old_bytes = old.public_bytes(serialization.Encoding.PEM)
    tls.ca_path.write_bytes(old_bytes)
    tls.cert_path.write_bytes(tls._leaf(old, key, leaf_key, datetime.now(timezone.utc)).public_bytes(serialization.Encoding.PEM))
    tls.ensure()
    assert tls.ca_path.read_bytes() == old_bytes
