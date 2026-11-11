"""Authorization uses approved bridge sessions, not invented review context."""
import json
from types import SimpleNamespace

import pytest
import desk_loader

SESSION = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"
TOOLS = ["get_document_overview", "read_selection", "read_document_text", "get_document_structure",
         "apply_edit_plan", "insert_text", "replace_selection", "find_replace", "format_selection",
         "apply_style", "insert_ooxml", "insert_comment"]


def setup_handler(monkeypatch, tool, *, bound=True, approved=True):
    module = desk_loader.fresh("desk_plugin_policy_test")
    calls = []
    bridge = SimpleNamespace(port=12345, base_url="https://localhost:12345", plugin_key="test-only",
                             has_session=lambda sid: bound and sid == SESSION,
                             session_allowed=lambda sid: approved and sid == SESSION,
                             tls=SimpleNamespace(client_context=lambda: None))
    module._RUNTIME = SimpleNamespace(bridge=bridge, key_provider=lambda: None)
    def fake_http(session_id, name, arguments, **kwargs):
        calls.append((session_id, name, arguments))
        return {"ok": True, "data": {"executed": True}}
    monkeypatch.setattr(module, "_invoke_bridge_http", fake_http)
    return module._handler_for("sidecar_word_" + tool), calls


@pytest.mark.parametrize("tool", TOOLS)
def test_actual_hermes_handler_kwargs_allow_approved_bound_tools(monkeypatch, tool):
    handler, calls = setup_handler(monkeypatch, tool)
    # model_tools.py:840-852 forwards these fields only, user_task defaults to None.
    result = json.loads(handler({}, task_id="task-1", session_id=SESSION, user_task=None))
    assert result == {"executed": True}
    assert calls == [(SESSION, tool, {})]


@pytest.mark.parametrize("bound,approved", [(False, True), (True, False), (False, False)])
@pytest.mark.parametrize("tool", TOOLS)
def test_missing_binding_or_revoked_approval_never_reaches_pane(monkeypatch, bound, approved, tool):
    handler, calls = setup_handler(monkeypatch, tool, bound=bound, approved=approved)
    result = json.loads(handler({}, task_id="task-1", session_id=SESSION, user_task=None))
    assert result["ok"] is False
    assert calls == []


@pytest.mark.parametrize("sid", [None, "", "bad", 123])
def test_invalid_trusted_session_fails_closed(monkeypatch, sid):
    handler, calls = setup_handler(monkeypatch, "insert_text")
    assert json.loads(handler({}, task_id="task-1", session_id=sid, user_task=None))["ok"] is False
    assert calls == []


def test_unknown_tool_fails_closed(monkeypatch):
    handler, calls = setup_handler(monkeypatch, "future_write")
    assert json.loads(handler({}, session_id=SESSION))["error"]["category"] == "Policy"
    assert calls == []


@pytest.mark.parametrize("field", ["bridgeSessionId", "session_id", "user_task", "turn_id"])
def test_model_cannot_supply_authority(monkeypatch, field):
    handler, calls = setup_handler(monkeypatch, "insert_text", bound=False)
    result = json.loads(handler({field: SESSION}, task_id="task-1", session_id=SESSION, user_task=None))
    assert result["ok"] is False
    assert calls == []


def test_comments_are_not_falsely_limited_to_one_per_session(monkeypatch):
    handler, calls = setup_handler(monkeypatch, "insert_comment")
    for text in ["first", "second"]:
        assert json.loads(handler({"text": text}, task_id="task-1", session_id=SESSION, user_task=None)) == {"executed": True}
    assert len(calls) == 2
