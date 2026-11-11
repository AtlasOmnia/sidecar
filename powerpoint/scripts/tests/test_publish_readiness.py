# SPDX-License-Identifier: Apache-2.0
"""Publication regressions; never touches a real keychain or installed plugin."""
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
import desk_loader
from test_desk_bridge_security import make_runtime

ROOT = Path(__file__).resolve().parents[2]


def test_revoked_approval_blocks_ticket_immediately(monkeypatch):
    plugin = desk_loader.fresh("revoked_ticket_test")
    allowed = [True]
    bridge = SimpleNamespace(port=12345, has_session=lambda sid: True,
                             session_allowed=lambda sid: allowed[0], base_url="https://localhost:12345",
                             plugin_key="test-only", tls=SimpleNamespace(client_context=lambda: None))
    plugin._RUNTIME = SimpleNamespace(bridge=bridge, key_provider=lambda: None)
    calls = []
    monkeypatch.setattr(plugin, "_invoke_bridge_http", lambda *a, **kw: calls.append(a) or {"ok": True})
    plugin.invoke_bridge("session", "read_all_slides", {})
    assert len(calls) == 1
    allowed[0] = False  # the grant was revoked while the turn was running
    with pytest.raises(plugin.BridgeError, match="sidebar"):
        plugin.invoke_bridge("session", "read_all_slides", {})
    assert len(calls) == 1


@pytest.mark.parametrize("command", ["sidecar-powerpoint", "sidecar-word", "sidecar-excel"])
def test_inherited_gateway_cli_never_starts_listener(tmp_path, monkeypatch, command):
    from desk_plugin.desk_bridge import _in_gateway
    monkeypatch.setenv("_HERMES_GATEWAY", "1")
    monkeypatch.setattr(sys, "argv", ["hermes", command, "status"])
    rt = make_runtime(tmp_path)
    rt._in_gateway = _in_gateway
    monkeypatch.setattr(rt, "ensure_started", lambda: pytest.fail("CLI started a listener"))
    assert rt.maybe_autostart() is None
    rt.slash("status")
    assert not _in_gateway()
    monkeypatch.setattr(sys, "argv", ["hermes", "gateway"])
    assert _in_gateway()


def test_plain_status_does_not_repeat_trust(tmp_path):
    rt = make_runtime(tmp_path)
    rt.bridge.tls.ensure()
    assert "add-trusted-cert" not in rt.status_text()
    assert "add-trusted-cert" in rt.slash("trust")


def test_installed_license_and_notice_match_root():
    for name in ("LICENSE", "NOTICE"):
        assert (ROOT / name).read_bytes() == (ROOT / "plugins/atlas-sidecar-powerpoint" / name).read_bytes()


def test_allow_question_discloses_agent_authority(tmp_path):
    questions = []
    rt = make_runtime(tmp_path)
    rt._ask_choice = lambda question, choices: questions.append(question) or "deny"
    rt.store.create_request(rt.allowed_origin, {})
    rt.slash()
    assert f"site at {rt.allowed_origin}" in questions[0]
    assert "all its tools until you revoke" in questions[0]
    assert "open only documents you trust" in questions[0]


def test_new_ca_name_and_legacy_ca_still_loads(tmp_path, monkeypatch):
    from cryptography import x509
    from cryptography.x509.oid import NameOID
    from desk_plugin.desk_tls import LocalTLS
    tls = LocalTLS(tmp_path / "new")
    tls.ensure()
    ca = x509.load_pem_x509_certificate(tls.ca_path.read_bytes())
    assert ca.subject.get_attributes_for_oid(NameOID.COMMON_NAME)[0].value.startswith("Sidecar for PowerPoint local CA ")
    tls.ensure()
    legacy = LocalTLS(tmp_path / "legacy")
    original = x509.NameAttribute.__init__
    with monkeypatch.context() as patch:
        def old_name(self, oid, value, *args, **kwargs):
            original(self, oid, value.replace("Sidecar for PowerPoint local CA ", "Sidecar local CA "), *args, **kwargs)
        patch.setattr(x509.NameAttribute, "__init__", old_name)
        legacy.ensure()
    before = legacy.ca_path.read_bytes()
    assert x509.load_pem_x509_certificate(before).subject.get_attributes_for_oid(NameOID.COMMON_NAME)[0].value.startswith("Sidecar local CA ")
    legacy.ensure()
    assert legacy.ca_path.read_bytes() == before
