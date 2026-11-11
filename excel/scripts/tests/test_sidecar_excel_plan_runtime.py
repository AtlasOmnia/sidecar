"""Real in-plugin Desk runtime -> renamed trusted plan hook -> companion -> pane protocol."""
import json
import threading
import time
import uuid

from test_desk_bridge_security import (FakeCtx, ADDIN_BASE, _free_ports, connect, request, session_headers, plugin_env)


def test_sidecar_excel_plan_hooks_reach_the_approved_pane_without_key_files(plugin_env, monkeypatch):
    monkeypatch.setenv("_HERMES_GATEWAY", "1")
    monkeypatch.setattr(plugin_env.desk_app, "PORTS", _free_ports(2))

    class Ctx(FakeCtx):
        hooks: dict | None = None
        def register_hook(self, name, callback):
            if self.hooks is None:
                self.hooks = {}
            self.hooks[name] = callback

    ctx = Ctx({"addin_base": ADDIN_BASE})
    plugin_env.register(ctx)
    assert isinstance(ctx.hooks, dict)
    rt = plugin_env._RUNTIME
    grant = connect(rt)["grant"]
    pane = str(uuid.uuid4())
    status, _, session = request(rt, "POST", "/excel/session", {
        "protocolVersion": "1.0", "documentId": str(uuid.uuid4()), "paneId": pane,
        "workbookPlanFingerprint": plugin_env._PLAN_FINGERPRINT,
    }, headers={"X-Hermes-Desk-Grant": grant})
    assert status == 201 and isinstance(session, dict)
    headers = session_headers(session, pane)
    plan = {
        "version": 1, "outcome": "execute", "summary": "Write one cell",
        "operations": [{"kind": "write_values", "sheet": "Sheet1", "address": "A1", "values": [["ok"]]}],
        "verification": [{"kind": "range_values_equal", "sheet": "Sheet1", "address": "A1", "expectedValues": [["ok"]]}],
    }
    metadata = dict(session_id=session["sessionId"], turn_id="desk-plan-turn", task_id="task",
                    tool_call_id="call", api_request_id="request")
    assert ctx.hooks["pre_tool_call"](tool_name="excel_execute_plan", args=plan, **metadata) is None
    assert plugin_env._trusted_plan_context.get() is None  # old plugin name cannot seed trust
    assert ctx.hooks["pre_tool_call"](tool_name="sidecar_excel_execute_plan", args=plan, **metadata) is None
    seen = {}
    errors = []

    def pane_loop():
        try:
            deadline = time.monotonic() + 5
            while time.monotonic() < deadline:
                status, _, ticket = request(rt, "GET", "/excel/pending", headers=headers)
                if status == 200:
                    assert isinstance(ticket, dict)
                    seen.update(ticket)
                    assert ticket["tool"] == "excel_execute_plan"  # canonical pane protocol is unchanged
                    assert ticket["trusted"]["phase"] == "execute"
                    assert ticket["trusted"]["workbookPlanFingerprint"] == plugin_env._PLAN_FINGERPRINT
                    assert request(rt, "POST", "/excel/execute-start", {"ticketId": ticket["ticketId"]}, headers=headers)[0] == 200
                    assert request(rt, "POST", "/excel/result", {
                        "ticketId": ticket["ticketId"], "executed": True, "result": {"accepted": True},
                    }, headers=headers)[0] == 200
                    return
                time.sleep(0.01)
            raise AssertionError("No plan ticket arrived")
        except BaseException as exc:
            errors.append(exc)

    thread = threading.Thread(target=pane_loop, daemon=True)
    thread.start()
    output = ctx.tools["sidecar_excel_execute_plan"]["handler"](plan, **metadata)
    thread.join(6)
    assert not thread.is_alive() and not errors, errors
    assert seen and json.loads(output) == {"accepted": True}
    ctx.hooks["post_tool_call"](tool_name="sidecar_excel_execute_plan")
    assert plugin_env._trusted_plan_context.get() is None
    replay = json.loads(ctx.tools["sidecar_excel_execute_plan"]["handler"](plan, **metadata))
    assert replay["error"]["category"] == "Validation"
