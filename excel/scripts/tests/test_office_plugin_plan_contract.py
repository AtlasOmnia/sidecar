"""Task 12 contract and trusted plan-route tests."""
from __future__ import annotations

import copy
import importlib.util
import json
import re
import shutil
import subprocess
import sys
import traceback
import uuid

import pytest
from concurrent.futures import ThreadPoolExecutor
from contextvars import copy_context
from pathlib import Path
from typing import Any, Callable, cast
from unittest.mock import patch

REPO = Path(__file__).resolve().parents[2]
PLUGIN_DIR = REPO / "plugins" / "atlas-sidecar-excel"
SCHEMAS = PLUGIN_DIR / "tool_schemas.json"
PLAN_TOOL = "sidecar_excel_execute_plan"
SESSION = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"
OTHER_SESSION = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb"
TRUSTED_ENVELOPE_CANARY = "g3-trusted-envelope-canary"
# R4g-2a rule M: the turnKey of migrated internal envelopes (canonical, never equal to an op id here).
_R4G2_TURN_KEY = "4b4b0000-0000-4000-8000-0000000000a1"

# Task 12's historical generated fingerprint was
# b1f1b6353c9fe49d11d9dcdca7e41a97c06d7535cab04bb29d87fdc2d3d3d8ac.
# The authored contract has since changed; current parity is derived from the
# authored generator instead of freezing either historical or current output.
HISTORICAL_TASK12_FINGERPRINT = "b1f1b6353c9fe49d11d9dcdca7e41a97c06d7535cab04bb29d87fdc2d3d3d8ac"


def load_plugin():
    name = "hermes_plugins.hermes_office_excel_plan_test"
    spec = importlib.util.spec_from_file_location(
        name, PLUGIN_DIR / "__init__.py", submodule_search_locations=[str(PLUGIN_DIR)]
    )
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


class FakeContext:
    def __init__(self) -> None:
        self.tools: list[dict] = []
        self.hooks: dict[str, object] = {}

    def register_tool(self, **kwargs) -> None:
        self.tools.append(kwargs)

    def register_hook(self, name: str, callback) -> None:
        self.hooks[name] = callback


def plan_schema() -> dict:
    return next(schema for schema in json.loads(SCHEMAS.read_text()) if schema["name"] == PLAN_TOOL)


def generated_schema_text() -> str:
    return (REPO / "src/generated/workbookPlanSchema.ts").read_text()


def generated_schema() -> dict:
    return json.loads(
        generated_schema_text().split("export const workbookPlanSchema = ", 1)[1].split(
            " as const;", 1
        )[0]
    )


def generated_fingerprint() -> str:
    match = re.search(
        r'export const workbookPlanFingerprint = "([0-9a-f]{64})" as const;',
        generated_schema_text(),
    )
    assert match is not None
    return match.group(1)


def assert_authored_generated_consistency() -> None:
    node = shutil.which("node")
    assert node is not None, "Node is required to validate the authored WorkbookPlan generator"
    result = subprocess.run(
        [node, str(REPO / "scripts/generate-workbook-plan-contract.mjs"), "--check"],
        cwd=REPO,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def assert_generated_plugin_consistency(
    generated: dict, fingerprint: str, plugin: dict
) -> None:
    assert plugin["parameters"] == generated
    assert plugin["x-workbook-plan-fingerprint"] == fingerprint


def valid_plan() -> dict:
    return {
        "version": 1,
        "outcome": "inspect",
        "reason": "Need workbook facts before proposing a change",
        "reads": [{"kind": "read_range", "sheet": "Sheet1", "address": "A1:B2"}],
    }


def registered():
    module = load_plugin()
    context = FakeContext()
    module.register(context)
    return module, context


def test_authored_generated_and_plugin_plan_schema_are_consistent() -> None:
    assert_authored_generated_consistency()
    assert_generated_plugin_consistency(
        generated_schema(), generated_fingerprint(), plan_schema()
    )


def test_schema_drift_is_rejected_in_an_isolated_fixture() -> None:
    generated = generated_schema()
    drifted_plugin = copy.deepcopy(plan_schema())
    drifted_plugin["parameters"]["properties"]["outcome"]["enum"] = [
        *drifted_plugin["parameters"]["properties"]["outcome"]["enum"],
        "intentional_drift",
    ]
    try:
        assert_generated_plugin_consistency(
            generated, generated_fingerprint(), drifted_plugin
        )
    except AssertionError:
        return
    raise AssertionError("intentional schema drift was accepted")


def test_plan_schema_is_closed_and_has_no_private_or_invented_fields() -> None:
    schema = plan_schema()
    model_text = json.dumps(schema["parameters"], sort_keys=True)
    for forbidden in (
        "bridgeSessionId", "paneId", "documentId", "sessionId", "sessionKey",
        "turn_operation_id", "turnOperationId", "requestId", "idempotency",
        "policy", "limits", "fingerprint", "routing", "replace",
    ):
        assert forbidden not in model_text
    assert schema["parameters"]["additionalProperties"] is False
    assert all(
        node.get("additionalProperties") is False
        for node in _objects(schema["parameters"])
    )


def _assert_recursively_bounded(schema: dict[str, Any]) -> None:
    for path, node in _schema_nodes(schema, "$parameters"):
        if not isinstance(node, dict):
            continue
        if node.get("type") == "object":
            assert node.get("additionalProperties") is False, path
            properties = node.get("properties", {})
            assert isinstance(properties, dict), path
            if "maxProperties" in node:
                max_properties = node["maxProperties"]
                assert (
                    isinstance(max_properties, int)
                    and not isinstance(max_properties, bool)
                    and max_properties >= 0
                    and max_properties == len(properties)
                ), path
        elif node.get("type") == "array":
            max_items = node.get("maxItems")
            assert (
                isinstance(max_items, int)
                and not isinstance(max_items, bool)
                and max_items >= 0
            ), path
        elif node.get("type") == "string":
            if "maxLength" in node:
                max_length = node["maxLength"]
                assert (
                    isinstance(max_length, int)
                    and not isinstance(max_length, bool)
                    and max_length >= 0
                ), path
            else:
                enum = node.get("enum")
                finite_enum = isinstance(enum, list) and all(
                    isinstance(item, str) for item in enum
                )
                finite_const = isinstance(node.get("const"), str)
                assert finite_enum or finite_const, path


def _schema_nodes(value: Any, path: str):
    yield path, value
    if isinstance(value, dict):
        for key, child in value.items():
            yield from _schema_nodes(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            yield from _schema_nodes(child, f"{path}[{index}]")


def test_plan_schema_is_recursively_bounded() -> None:
    _assert_recursively_bounded(plan_schema()["parameters"])


def test_boundedness_checker_accepts_explicitly_bounded_string_fixture() -> None:
    fixture = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "max_length": {"type": "string", "maxLength": 8},
            "enum": {"type": "string", "enum": ["red", "blue"]},
            "const": {"type": "string", "const": "fixed"},
        },
    }
    _assert_recursively_bounded(fixture)


def test_boundedness_checker_rejects_unbounded_string_fixture() -> None:
    fixture = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "nested": {
                "oneOf": [
                    {"type": "integer"},
                    {"type": "object", "additionalProperties": False, "properties": {
                        "arbitrary": {"type": "string"},
                    }},
                ],
            },
        },
    }
    try:
        _assert_recursively_bounded(fixture)
    except AssertionError:
        return
    raise AssertionError("unbounded string fixture was accepted")


def _objects(value):
    if isinstance(value, dict):
        if value.get("type") == "object":
            yield value
        for child in value.values():
            yield from _objects(child)
    elif isinstance(value, list):
        for child in value:
            yield from _objects(child)


def test_registration_is_28_tools_manifest_aligned_and_hooked() -> None:
    _, context = registered()
    schemas = json.loads(SCHEMAS.read_text())
    assert len(context.tools) == len(schemas) == 28
    assert context.hooks.keys() == {"pre_tool_call", "post_tool_call"}
    assert [tool["name"] for tool in context.tools] == [schema["name"] for schema in schemas]
    assert sum(tool["name"] != PLAN_TOOL for tool in context.tools) == 27


def test_hook_handler_bridge_end_to_end_private_envelope_and_model_cannot_override() -> None:
    module, context = registered()
    calls = []

    def bridge(session_id, tool, arguments, *, trusted=None):
        calls.append((session_id, tool, arguments, trusted))
        return {"ok": True, "data": {"accepted": True}}

    plan_handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    hook = context.hooks["pre_tool_call"]
    with patch.object(module, "invoke_bridge", bridge):
        hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="turn-1",
             tool_call_id="call-1", api_request_id="request-1", task_id="task-1")
        rejected = json.loads(plan_handler({**valid_plan(), "turnOperationId": "model", "fingerprint": "model"}, session_id=SESSION))
        assert rejected["error"]["category"] == "Validation"
        hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="turn-2",
             tool_call_id="call-2", api_request_id="request-2", task_id="task-1")
        result = json.loads(plan_handler(valid_plan(), session_id=SESSION))
    assert result == {"accepted": True}
    assert calls[0][1] == "execute_plan"
    assert calls[0][2] == valid_plan()
    assert calls[0][3]["turn_operation_id"] == calls[0][3]["request_id"]
    assert calls[0][3]["workbook_plan_fingerprint"] == generated_fingerprint()
    assert calls[0][3]["turn_operation_id"] != "model"


def test_plan_requires_fresh_matching_trusted_context_and_consumes_it() -> None:
    module, context = registered()
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    hook = context.hooks["pre_tool_call"]
    calls = []

    def bridge(*args, **kwargs):
        calls.append((args, kwargs))
        return {"ok": True}

    with patch.object(module, "invoke_bridge", bridge):
        for kwargs in (
            {},
            {"session_id": OTHER_SESSION},
            {"session_id": SESSION},
        ):
            assert json.loads(handler(valid_plan(), **kwargs))["error"]["category"] == "Validation"
        hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="turn-1",
             tool_call_id="call-1", api_request_id="request-1", task_id="task-1")
        assert json.loads(handler(valid_plan(), session_id=OTHER_SESSION))["error"]["category"] == "Validation"
        assert json.loads(handler(valid_plan(), session_id=SESSION))["ok"] is True
        assert json.loads(handler(valid_plan(), session_id=SESSION))["error"]["category"] == "Validation"
    assert len(calls) == 1


def test_invalid_plan_hook_clears_stale_context_and_primitive_is_unchanged() -> None:
    module, context = registered()
    hook = context.hooks["pre_tool_call"]
    plan_handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    primitive_handler = next(tool["handler"] for tool in context.tools if tool["name"] == "sidecar_excel_read_range")
    with patch.object(module, "invoke_bridge", return_value={"ok": True, "data": {}}) as bridge:
        hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="turn-1",
             tool_call_id="call-1", api_request_id="request-1", task_id="task-1")
        hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="",
             tool_call_id="call-2", api_request_id="request-2", task_id="task-2")
        assert json.loads(plan_handler(valid_plan(), session_id=SESSION))["error"]["category"] == "Validation"
        primitive_handler({"sheet": "Sheet1", "address": "A1"}, session_id=SESSION)
    assert bridge.call_count == 1
    assert bridge.call_args.args == (SESSION, "read_range", {"sheet": "Sheet1", "address": "A1"})
    assert bridge.call_args.kwargs == {}


def test_same_turn_identity_is_deterministic_and_later_turn_differs() -> None:
    module, context = registered()
    hook = context.hooks["pre_tool_call"]
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    envelopes = []
    with patch.object(module, "invoke_bridge", side_effect=lambda *a, **kw: envelopes.append(kw["trusted"]) or {"ok": True}):
        hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="same",
             tool_call_id="call-a", api_request_id="request-a", task_id="task")
        handler(valid_plan(), session_id=SESSION)
        blocked = hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="same",
                       tool_call_id="call-b", api_request_id="request-b", task_id="task")
        assert blocked["action"] == "block"
        assert json.loads(handler(valid_plan(), session_id=SESSION))["error"]["category"] == "Validation"
        hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="later",
             tool_call_id="call-c", api_request_id="request-c", task_id="task")
        handler(valid_plan(), session_id=SESSION)
    assert envelopes[0]["turn_operation_id"] != envelopes[1]["turn_operation_id"]


def test_trusted_context_is_cleared_after_handler_and_post_hook() -> None:
    module, context = registered()
    hook = context.hooks["pre_tool_call"]
    post = cast(Callable[..., Any], context.hooks["post_tool_call"])
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    with patch.object(module, "invoke_bridge", return_value={"ok": True}):
        hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="cleanup",
             tool_call_id="call-cleanup", api_request_id="request-cleanup", task_id="task")
        assert json.loads(handler(valid_plan(), session_id=SESSION))["ok"] is True
        post(tool_name=PLAN_TOOL, args=valid_plan(), result="{}", session_id=SESSION,
             turn_id="cleanup", tool_call_id="call-cleanup", api_request_id="request-cleanup", task_id="task")
        assert json.loads(handler(valid_plan(), session_id=SESSION))["error"]["category"] == "Validation"


def test_bridge_response_bounds_and_content_free_errors() -> None:
    module = load_plugin()
    bridge_client = module.bridge_client

    class OversizedResponse:
        headers = {"Content-Length": str(bridge_client.MAX_RESPONSE_BYTES + 1)}

        def read(self, _limit: int) -> bytes:
            return b""

    try:
        bridge_client._read_bounded(OversizedResponse())
    except bridge_client.BridgeError as exc:
        assert exc.category == "Capacity"
        assert "secret" not in str(exc)
    else:
        raise AssertionError("oversized companion response was accepted")


def _assert_real_bridge_rejects_trusted_metadata(
    bridge_client: Any, trusted: dict[str, Any], expected_category: str
) -> None:
    seam_calls: list[str] = []

    def denied_seam(name: str):
        def call(*_args: Any, **_kwargs: Any) -> None:
            seam_calls.append(name)
            raise AssertionError(f"{name} seam was called for rejected trusted metadata")

        return call

    with (
        patch.object(bridge_client, "_connection", side_effect=denied_seam("config")),
        patch.object(bridge_client, "_validate_key", side_effect=denied_seam("key")),
        patch.object(
            bridge_client.urllib.request,
            "urlopen",
            side_effect=denied_seam("network"),
        ),
    ):
        try:
            bridge_client.invoke_bridge(
                SESSION,
                "execute_plan",
                valid_plan(),
                trusted=trusted,
            )
        except bridge_client.BridgeError as exc:
            assert exc.category == expected_category
            assert TRUSTED_ENVELOPE_CANARY not in str(exc)
        except TypeError as exc:
            assert seam_calls == []
            assert TRUSTED_ENVELOPE_CANARY not in str(exc)
            raise AssertionError(
                "real bridge receiver rejected trusted metadata with unexpected "
                "TypeError; expected structured BridgeError"
            ) from exc
        else:
            raise AssertionError("rejected trusted metadata was accepted")

    assert seam_calls == []


def test_real_bridge_rejects_malformed_trusted_envelope_before_side_effects() -> None:
    module = load_plugin()
    _assert_real_bridge_rejects_trusted_metadata(
        module.bridge_client,
        {
            "turn_operation_id": TRUSTED_ENVELOPE_CANARY,
            "request_id": "",
            "workbook_plan_fingerprint": generated_fingerprint(),
            "phase": "inspect",
            "turn_key": _R4G2_TURN_KEY,
        },
        "Validation",
    )


def test_real_bridge_rejects_oversized_trusted_envelope_before_side_effects() -> None:
    module = load_plugin()
    _assert_real_bridge_rejects_trusted_metadata(
        module.bridge_client,
        {
            "turn_operation_id": SESSION,
            "request_id": TRUSTED_ENVELOPE_CANARY + ("x" * 1024),
            "workbook_plan_fingerprint": generated_fingerprint(),
            "phase": "inspect",
            "turn_key": _R4G2_TURN_KEY,
        },
        "Capacity",
    )


def test_version_and_pane_fingerprint_mismatch_fail_before_bridge_dispatch() -> None:
    module, context = registered()
    hook = context.hooks["pre_tool_call"]
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    calls = []
    with patch.object(module, "invoke_bridge", side_effect=lambda *args, **kwargs: calls.append(args)):
        invalid_version = valid_plan()
        invalid_version["version"] = 2
        hook(tool_name=PLAN_TOOL, args=invalid_version, session_id=SESSION, turn_id="version",
             tool_call_id="call-version", api_request_id="request-version", task_id="task")
        assert json.loads(handler(invalid_version, session_id=SESSION))["error"]["category"] == "Validation"
        hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="fingerprint",
             tool_call_id="call-fingerprint", api_request_id="request-fingerprint", task_id="task",
             pane_plan_fingerprint="0" * 64)
        assert json.loads(handler(valid_plan(), session_id=SESSION))["error"]["category"] == "Validation"
    assert calls == []


def test_trusted_context_cross_context_worker_cleanup_is_one_shot() -> None:
    module, context = registered()
    hook = context.hooks["pre_tool_call"]
    post = cast(Callable[..., Any], context.hooks["post_tool_call"])
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    calls = []

    def bridge(*args, **kwargs):
        calls.append((args, kwargs))
        return {"ok": True, "data": {"accepted": True}}

    def current_trusted_context():
        holder = module._trusted_plan_context.get()
        if holder is None:
            return None
        consume = getattr(holder, "consume", None)
        return consume() if callable(consume) else holder

    with patch.object(module, "invoke_bridge", bridge):
        hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="copied",
             tool_call_id="call-copied", api_request_id="request-copied", task_id="task")
        worker = copy_context()

        def complete_in_worker():
            result = handler(valid_plan(), session_id=SESSION)
            post(tool_name=PLAN_TOOL, args=valid_plan(), result="{}", session_id=SESSION,
                 turn_id="copied", tool_call_id="call-copied", api_request_id="request-copied",
                 task_id="task")
            return result, current_trusted_context()

        result, worker_context = worker.run(complete_in_worker)
        parent_context = current_trusted_context()

    assert json.loads(result) == {"accepted": True}
    assert worker_context is None
    assert parent_context is None
    assert len(calls) == 1


def test_trusted_plan_copied_contexts_isolate_concurrent_sessions_and_turns() -> None:
    module, context = registered()
    hook = context.hooks["pre_tool_call"]
    post = cast(Callable[..., Any], context.hooks["post_tool_call"])
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    calls = []

    def bridge(*args, **kwargs):
        calls.append((args, kwargs))
        return {"ok": True, "data": {"accepted": kwargs["trusted"]["request_id"]}}

    def start(parent, session, turn):
        parent.run(
            hook,
            tool_name=PLAN_TOOL,
            args=valid_plan(),
            session_id=session,
            turn_id=turn,
            tool_call_id=f"call-{turn}",
            api_request_id=f"request-{turn}",
            task_id=f"task-{turn}",
        )
        return parent.copy()

    def current_trusted_context():
        holder = module._trusted_plan_context.get()
        if holder is None:
            return None
        consume = getattr(holder, "consume", None)
        return consume() if callable(consume) else holder

    parent_a = copy_context()
    parent_b = copy_context()
    worker_a = start(parent_a, SESSION, "concurrent-a")
    worker_b = start(parent_b, OTHER_SESSION, "concurrent-b")

    def complete(worker, session, turn):
        def complete_in_worker():
            result = handler(valid_plan(), session_id=session)
            post(tool_name=PLAN_TOOL, args=valid_plan(), result="{}", session_id=session,
                 turn_id=turn, tool_call_id=f"call-{turn}", api_request_id=f"request-{turn}",
                 task_id=f"task-{turn}")
            return result, current_trusted_context()

        return worker.run(complete_in_worker)

    with patch.object(module, "invoke_bridge", bridge):
        with ThreadPoolExecutor(max_workers=2) as executor:
            futures = [
                executor.submit(complete, worker_a, SESSION, "concurrent-a"),
                executor.submit(complete, worker_b, OTHER_SESSION, "concurrent-b"),
            ]
            results = [future.result() for future in futures]

    accepted = [json.loads(result[0])["accepted"] for result in results]
    assert set(accepted) == {call[1]["trusted"]["request_id"] for call in calls}
    assert all(result[1] is None for result in results)
    assert current_trusted_context() is None
    assert parent_a.run(current_trusted_context) is None
    assert parent_b.run(current_trusted_context) is None
    assert len(calls) == 2
    assert {call[0][0] for call in calls} == {SESSION, OTHER_SESSION}
    assert len({call[1]["trusted"]["request_id"] for call in calls}) == 2
    assert all(
        call[1]["trusted"]["turn_operation_id"] == call[1]["trusted"]["request_id"]
        for call in calls
    )


class _TrustedStringSubclass(str):
    pass


class _TrustedDictSubclass(dict):
    def __iter__(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)


class _TrustedConversionHook:
    def __str__(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def __bytes__(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def encode(self, *_args, **_kwargs):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def __len__(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)


class _FakeBridgeResponse:
    def __init__(self, payload: bytes) -> None:
        self._payload = payload
        self.headers = {"Content-Length": str(len(payload))}

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def read(self, _limit: int) -> bytes:
        return self._payload


def _trusted_metadata(
    *,
    turn_operation_id: Any = SESSION,
    request_id: Any = SESSION,
    workbook_plan_fingerprint: Any = None,
    phase: Any = "inspect",
    turn_key: Any = _R4G2_TURN_KEY,
) -> dict[str, Any]:
    return {
        "turn_operation_id": turn_operation_id,
        "request_id": request_id,
        "workbook_plan_fingerprint": (
            generated_fingerprint()
            if workbook_plan_fingerprint is None
            else workbook_plan_fingerprint
        ),
        "phase": phase,
        "turn_key": turn_key,
    }


def _assert_content_free_bridge_error(exc: BaseException) -> None:
    pending = [exc]
    seen: set[int] = set()
    while pending:
        current = pending.pop()
        if id(current) in seen:
            continue
        seen.add(id(current))
        assert TRUSTED_ENVELOPE_CANARY not in str(current)
        assert TRUSTED_ENVELOPE_CANARY not in "".join(traceback.format_exception(current))
        if current.__cause__ is not None:
            pending.append(current.__cause__)
        if current.__context__ is not None:
            pending.append(current.__context__)


def _invoke_bridge_assertion_level(bridge_client: Any, *args: Any, **kwargs: Any) -> Any:
    try:
        return bridge_client.invoke_bridge(*args, **kwargs)
    except TypeError as exc:
        raise AssertionError(
            "trusted receiver raised unexpected TypeError instead of completing the call"
        ) from exc


def _call_handler_assertion_level(handler: Any, *args: Any, **kwargs: Any) -> Any:
    try:
        return handler(*args, **kwargs)
    except TypeError as exc:
        raise AssertionError(
            "registered plan handler raised unexpected TypeError instead of completing the call"
        ) from exc


def _assert_trusted_receiver_rejects(
    bridge_client: Any,
    trusted: Any,
    expected_category: str,
    *,
    tool: str = "execute_plan",
    arguments: Any = None,
    session_id: str = SESSION,
) -> None:
    seam_calls: list[str] = []
    if arguments is None:
        arguments = valid_plan()

    def denied_seam(name: str):
        def call(*_args: Any, **_kwargs: Any) -> None:
            seam_calls.append(name)
            raise AssertionError(f"{name} seam was called for rejected trusted metadata")

        return call

    caught: BaseException | None = None
    try:
        with (
            patch.object(bridge_client, "_connection", side_effect=denied_seam("config")),
            patch.object(bridge_client, "_validate_key", side_effect=denied_seam("key")),
            patch.object(
                bridge_client.urllib.request,
                "urlopen",
                side_effect=denied_seam("network"),
            ),
        ):
            bridge_client.invoke_bridge(
                session_id,
                tool,
                arguments,
                trusted=trusted,
            )
    except bridge_client.BridgeError as exc:
        caught = exc
    except TypeError as exc:
        assert seam_calls == []
        assert TRUSTED_ENVELOPE_CANARY not in str(exc)
        raise AssertionError(
            "trusted receiver raised unexpected TypeError instead of BridgeError"
        ) from exc
    except Exception as exc:
        raise AssertionError(
            f"trusted receiver raised {type(exc).__name__} instead of BridgeError"
        ) from exc

    assert caught is not None
    assert isinstance(caught, bridge_client.BridgeError)
    assert caught.category == expected_category
    _assert_content_free_bridge_error(caught)
    assert seam_calls == []


@pytest.mark.parametrize(
    "case_name,trusted,expected_category",
    [
        ("absent", None, "Validation"),
        ("list_container", [], "Validation"),
        ("string_container", TRUSTED_ENVELOPE_CANARY, "Validation"),
        ("scalar_container", 7, "Validation"),
        ("dict_subclass", _TrustedDictSubclass(_trusted_metadata()), "Validation"),
        ("missing_turn", {"request_id": SESSION, "workbook_plan_fingerprint": generated_fingerprint()}, "Validation"),
        ("missing_request", {"turn_operation_id": SESSION, "workbook_plan_fingerprint": generated_fingerprint()}, "Validation"),
        ("missing_fingerprint", {"turn_operation_id": SESSION, "request_id": SESSION}, "Validation"),
        ("camel_case_keys", {"turnOperationId": SESSION, "workbookPlanFingerprint": generated_fingerprint()}, "Validation"),
        (
            "extra_key",
            {
                **_trusted_metadata(),
                "extra": TRUSTED_ENVELOPE_CANARY,
            },
            "Validation",
        ),
    ],
)
def test_trusted_receiver_rejects_absent_wrong_container_and_wrong_key_set(
    case_name: str, trusted: Any, expected_category: str
) -> None:
    del case_name
    module = load_plugin()
    _assert_trusted_receiver_rejects(module.bridge_client, trusted, expected_category)


@pytest.mark.parametrize(
    "field,bad_value,case_name",
    [
        ("turn_operation_id", TRUSTED_ENVELOPE_CANARY.encode(), "turn_bytes"),
        ("request_id", TRUSTED_ENVELOPE_CANARY.encode(), "request_bytes"),
        ("workbook_plan_fingerprint", TRUSTED_ENVELOPE_CANARY.encode(), "fingerprint_bytes"),
        ("turn_operation_id", uuid.UUID(SESSION), "turn_uuid"),
        ("request_id", uuid.UUID(SESSION), "request_uuid"),
        ("workbook_plan_fingerprint", uuid.UUID(SESSION), "fingerprint_uuid"),
        ("turn_operation_id", True, "turn_bool"),
        ("request_id", False, "request_bool"),
        ("workbook_plan_fingerprint", True, "fingerprint_bool"),
        ("turn_operation_id", [SESSION], "turn_list"),
        ("request_id", (SESSION,), "request_tuple"),
        ("workbook_plan_fingerprint", {"value": SESSION}, "fingerprint_dict"),
        ("turn_operation_id", _TrustedStringSubclass(SESSION), "turn_str_subclass"),
        ("request_id", _TrustedStringSubclass(SESSION), "request_str_subclass"),
        (
            "workbook_plan_fingerprint",
            _TrustedStringSubclass(generated_fingerprint()),
            "fingerprint_str_subclass",
        ),
        ("turn_operation_id", _TrustedDictSubclass(), "turn_dict_subclass"),
        ("request_id", _TrustedDictSubclass(), "request_dict_subclass"),
        ("workbook_plan_fingerprint", _TrustedDictSubclass(), "fingerprint_dict_subclass"),
        ("turn_operation_id", _TrustedConversionHook(), "turn_conversion_hook"),
        ("request_id", _TrustedConversionHook(), "request_conversion_hook"),
        ("workbook_plan_fingerprint", _TrustedConversionHook(), "fingerprint_conversion_hook"),
    ],
)
def test_trusted_receiver_rejects_every_metadata_field_wrong_type(
    field: str, bad_value: Any, case_name: str
) -> None:
    del case_name
    trusted = _trusted_metadata()
    trusted[field] = bad_value
    module = load_plugin()
    _assert_trusted_receiver_rejects(module.bridge_client, trusted, "Validation")


@pytest.mark.parametrize(
    "case_name,trusted,expected_category",
    [
        (
            "turn_capacity",
            _trusted_metadata(turn_operation_id="x" * 257),
            "Capacity",
        ),
        (
            "request_capacity",
            _trusted_metadata(request_id="x" * 257),
            "Capacity",
        ),
        (
            "fingerprint_capacity",
            _trusted_metadata(workbook_plan_fingerprint="a" * 65),
            "Capacity",
        ),
        (
            "turn_multibyte_capacity",
            _trusted_metadata(turn_operation_id="é" * 129),
            "Capacity",
        ),
        (
            "request_multibyte_capacity",
            _trusted_metadata(request_id="é" * 129),
            "Capacity",
        ),
        (
            "fingerprint_multibyte_capacity",
            _trusted_metadata(workbook_plan_fingerprint="é" * 33),
            "Capacity",
        ),
        (
            "turn_lone_surrogate",
            _trusted_metadata(turn_operation_id="\ud800"),
            "Validation",
        ),
        (
            "request_lone_surrogate",
            _trusted_metadata(request_id="\ud800"),
            "Validation",
        ),
        (
            "fingerprint_lone_surrogate",
            _trusted_metadata(workbook_plan_fingerprint="\ud800"),
            "Validation",
        ),
        (
            "turn_malformed_canary",
            _trusted_metadata(turn_operation_id=TRUSTED_ENVELOPE_CANARY),
            "Validation",
        ),
        (
            "request_malformed_canary",
            _trusted_metadata(request_id=TRUSTED_ENVELOPE_CANARY),
            "Validation",
        ),
        (
            "unequal_valid_uuid",
            _trusted_metadata(request_id=OTHER_SESSION),
            "Validation",
        ),
        (
            "fingerprint_uppercase",
            _trusted_metadata(workbook_plan_fingerprint="A" + generated_fingerprint()[1:]),
            "Validation",
        ),
        (
            "fingerprint_nonhex",
            _trusted_metadata(workbook_plan_fingerprint="g" * 64),
            "Validation",
        ),
        (
            "fingerprint_short",
            _trusted_metadata(workbook_plan_fingerprint="a" * 63),
            "Validation",
        ),
    ],
)
def test_trusted_receiver_checks_capacity_before_uuid_and_fingerprint_semantics(
    case_name: str, trusted: dict[str, Any], expected_category: str
) -> None:
    del case_name
    module = load_plugin()
    _assert_trusted_receiver_rejects(module.bridge_client, trusted, expected_category)


def test_trusted_receiver_rejects_plan_without_host_trust_even_when_arguments_spoof_metadata() -> None:
    spoofed_arguments = {
        **valid_plan(),
        "turnOperationId": TRUSTED_ENVELOPE_CANARY,
        "workbookPlanFingerprint": TRUSTED_ENVELOPE_CANARY,
    }
    module = load_plugin()
    _assert_trusted_receiver_rejects(
        module.bridge_client,
        None,
        "Validation",
        arguments=spoofed_arguments,
    )


@pytest.mark.parametrize(
    "trusted,case_name",
    [
        ({}, "empty_dict"),
        (_trusted_metadata(), "valid_shape"),
        ([], "list"),
        (TRUSTED_ENVELOPE_CANARY, "string"),
        (0, "integer"),
    ],
)
def test_trusted_receiver_rejects_trust_on_primitive_tool(
    trusted: Any, case_name: str
) -> None:
    del case_name
    module = load_plugin()
    _assert_trusted_receiver_rejects(
        module.bridge_client,
        trusted,
        "Validation",
        tool="read_range",
        arguments={"sheet": "Sheet1", "address": "A1"},
    )


def test_trusted_receiver_preserves_valid_plan_wire_metadata_and_host_text() -> None:
    module = load_plugin()
    bridge_client = module.bridge_client
    calls: list[tuple[Any, int]] = []
    response_payload = json.dumps(
        {"ok": True, "data": {"accepted": True}}, separators=(",", ":")
    ).encode("utf-8")
    host_uuid = uuid.uuid5(uuid.UUID("6f37c2f0-29df-5c72-b6bb-e89f2f18f0c5"), "r2a-host")
    host_text = "{" + str(host_uuid).upper() + "}"
    operation_uuid = uuid.uuid5(uuid.UUID("6f37c2f0-29df-5c72-b6bb-e89f2f18f0c5"), "r2a-operation")
    operation_text = "urn:uuid:" + str(operation_uuid)
    unknown_fingerprint = "0" * 64
    if unknown_fingerprint == generated_fingerprint():
        unknown_fingerprint = "f" * 64
    arguments = {"payload": "é", "emoji": "🦄"}

    def fake_urlopen(request: Any, timeout: int):
        calls.append((request, timeout))
        return _FakeBridgeResponse(response_payload)

    with (
        patch.object(
            bridge_client,
            "_connection",
            return_value=("http://127.0.0.1:8650", "k" * 32),
        ),
        patch.object(bridge_client, "_validate_key", return_value="k" * 32),
        patch.object(bridge_client.urllib.request, "urlopen", side_effect=fake_urlopen),
    ):
        result = _invoke_bridge_assertion_level(
            bridge_client,
            host_text,
            "execute_plan",
            arguments,
            trusted={
                "turn_operation_id": operation_text,
                "request_id": operation_text,
                "workbook_plan_fingerprint": unknown_fingerprint,
                "phase": "execute",
                "turn_key": _R4G2_TURN_KEY,
            },
        )

    assert result == {"ok": True, "data": {"accepted": True}}
    assert len(calls) == 1
    request, timeout = calls[0]
    assert request.full_url == "http://127.0.0.1:8650/excel/invoke"
    assert request.get_method() == "POST"
    assert timeout == 65
    assert request.get_header("Content-type") == "application/json"
    assert request.get_header("X-hermes-plugin-key") == "k" * 32
    body_bytes = request.data
    assert isinstance(body_bytes, bytes)
    assert request.get_header("Content-length") == str(len(body_bytes))
    body = json.loads(body_bytes.decode("utf-8"))
    assert list(body) == [
        "protocolVersion",
        "sessionId",
        "requestId",
        "tool",
        "arguments",
        "trusted",
    ]
    assert set(body) == {
        "protocolVersion",
        "sessionId",
        "requestId",
        "tool",
        "arguments",
        "trusted",
    }
    assert body["protocolVersion"] == "1.0"
    assert body["sessionId"] == host_text
    assert body["requestId"] == operation_text
    assert body["tool"] == "execute_plan"
    assert body["arguments"] == arguments
    assert list(body["trusted"]) == ["turnOperationId", "workbookPlanFingerprint", "phase", "turnKey"]
    assert body["trusted"] == {
        "turnOperationId": operation_text,
        "workbookPlanFingerprint": unknown_fingerprint,
        "phase": "execute",
        "turnKey": _R4G2_TURN_KEY,
    }
    assert unknown_fingerprint != generated_fingerprint()


def test_trusted_receiver_primitive_none_generates_request_id_without_trusted_wire_member() -> None:
    module = load_plugin()
    bridge_client = module.bridge_client
    calls: list[tuple[Any, int]] = []
    response_payload = b'{"ok":true,"data":{"accepted":true}}'
    arguments = {"sheet": "Sheet1", "address": "A1", "text": "é"}

    def fake_urlopen(request: Any, timeout: int):
        calls.append((request, timeout))
        return _FakeBridgeResponse(response_payload)

    with (
        patch.object(
            bridge_client,
            "_connection",
            return_value=("http://127.0.0.1:8650", "k" * 32),
        ),
        patch.object(bridge_client, "_validate_key", return_value="k" * 32),
        patch.object(bridge_client.urllib.request, "urlopen", side_effect=fake_urlopen),
    ):
        result = _invoke_bridge_assertion_level(
            bridge_client,
            SESSION,
            "read_range",
            arguments,
            trusted=None,
        )

    assert result == {"ok": True, "data": {"accepted": True}}
    assert len(calls) == 1
    request, timeout = calls[0]
    assert timeout == 65
    body_bytes = request.data
    assert request.get_header("Content-length") == str(len(body_bytes))
    body = json.loads(body_bytes.decode("utf-8"))
    assert list(body) == ["protocolVersion", "sessionId", "requestId", "tool", "arguments"]
    assert "trusted" not in body
    assert body["sessionId"] == SESSION
    assert body["tool"] == "read_range"
    assert body["arguments"] == arguments
    generated_request_id = uuid.UUID(body["requestId"])
    assert str(generated_request_id) == body["requestId"]


def test_trusted_receiver_real_hook_registered_handler_and_receiver_preserve_uuid5_and_fingerprint() -> None:
    module, context = registered()
    bridge_client = module.bridge_client
    hook = context.hooks["pre_tool_call"]
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    plan = valid_plan()
    turn_id = "r2a-real-uuid5"
    calls: list[tuple[Any, int]] = []
    response_payload = b'{"ok":true,"data":{"accepted":true}}'

    def fake_urlopen(request: Any, timeout: int):
        calls.append((request, timeout))
        return _FakeBridgeResponse(response_payload)

    with (
        patch.object(
            bridge_client,
            "_connection",
            return_value=("http://127.0.0.1:8650", "k" * 32),
        ),
        patch.object(bridge_client, "_validate_key", return_value="k" * 32),
        patch.object(bridge_client.urllib.request, "urlopen", side_effect=fake_urlopen),
    ):
        assert (
            hook(
                tool_name=PLAN_TOOL,
                args=plan,
                session_id=SESSION,
                turn_id=turn_id,
                tool_call_id="call-r2a",
                api_request_id="request-r2a",
                task_id="task-r2a",
            )
            is None
        )
        result = json.loads(
            _call_handler_assertion_level(handler, plan, session_id=SESSION)
        )

    assert result == {"accepted": True}
    assert len(calls) == 1
    request, timeout = calls[0]
    assert timeout == 65
    body = json.loads(request.data.decode("utf-8"))
    expected_turn_key = str(
        uuid.uuid5(module._OPERATION_NAMESPACE, f"{SESSION}\\0{turn_id}")
    )
    assert uuid.UUID(expected_turn_key).version == 5
    # R4g-2a J2 literal (valid_plan() is an inspect plan): op(inspect) for (SESSION, "r2a-real-uuid5").
    expected_operation_id = "cfed8ba1-049f-58dd-ac4d-dad4417a0ca2"
    assert list(body) == [
        "protocolVersion",
        "sessionId",
        "requestId",
        "tool",
        "arguments",
        "trusted",
    ]
    assert body["sessionId"] == SESSION
    assert body["requestId"] == expected_operation_id
    assert body["tool"] == "execute_plan"
    assert body["arguments"] == plan
    assert body["trusted"] == {
        "turnOperationId": expected_operation_id,
        "workbookPlanFingerprint": generated_fingerprint(),
        "phase": "inspect",
        "turnKey": expected_turn_key,
    }
    assert request.full_url == "http://127.0.0.1:8650/excel/invoke"
    assert request.get_method() == "POST"
    assert request.get_header("Content-length") == str(len(request.data))


def test_trusted_receiver_model_override_is_rejected_before_config_key_or_network() -> None:
    module, context = registered()
    bridge_client = module.bridge_client
    hook = context.hooks["pre_tool_call"]
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    seam_calls: list[str] = []

    def denied_seam(name: str):
        def call(*_args: Any, **_kwargs: Any) -> None:
            seam_calls.append(name)
            raise AssertionError(f"{name} seam was called for model override")

        return call

    model_override = {
        **valid_plan(),
        "turnOperationId": TRUSTED_ENVELOPE_CANARY,
    }
    with (
        patch.object(bridge_client, "_connection", side_effect=denied_seam("config")),
        patch.object(bridge_client, "_validate_key", side_effect=denied_seam("key")),
        patch.object(
            bridge_client.urllib.request,
            "urlopen",
            side_effect=denied_seam("network"),
        ),
    ):
        hook(
            tool_name=PLAN_TOOL,
            args=valid_plan(),
            session_id=SESSION,
            turn_id="override-turn",
            tool_call_id="override-call",
            api_request_id="override-request",
            task_id="override-task",
        )
        result = json.loads(handler(model_override, session_id=SESSION))

    assert result["ok"] is False
    assert result["error"]["category"] == "Validation"
    assert TRUSTED_ENVELOPE_CANARY not in json.dumps(result)
    assert seam_calls == []


class _R2aReceiverHookObject:
    def __bool__(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def __len__(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def __str__(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def __repr__(self):
        return "<r2a-hook-object>"

    def __eq__(self, _other):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def replace(self, *_args, **_kwargs):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def encode(self, *_args, **_kwargs):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)


class _R2aReceiverHookString(str):
    def __bool__(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def __len__(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def __str__(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def __eq__(self, _other):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def replace(self, *_args, **_kwargs):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def encode(self, *_args, **_kwargs):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)


class _R2aReceiverHookDict(dict):
    def __bool__(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def __len__(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def __iter__(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def keys(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def items(self):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)

    def __getitem__(self, _key):
        raise RuntimeError(TRUSTED_ENVELOPE_CANARY)


def _ordinary_arguments() -> dict[str, str]:
    return {"sheet": "Sheet1", "address": "A1"}


_ARGUMENTS_OMITTED = object()


def _assert_receiver_rejects_before_side_effects(
    bridge_client: Any,
    expected_category: str,
    *,
    session_id: Any = SESSION,
    tool: Any = "read_range",
    arguments: Any = _ARGUMENTS_OMITTED,
    trusted: Any = None,
) -> None:
    seam_calls: list[str] = []
    if arguments is _ARGUMENTS_OMITTED:
        arguments = _ordinary_arguments()

    def denied_seam(name: str):
        def call(*_args: Any, **_kwargs: Any) -> None:
            seam_calls.append(name)
            raise RuntimeError(f"{TRUSTED_ENVELOPE_CANARY}:{name}")

        return call

    try:
        with (
            patch.object(bridge_client, "_connection", side_effect=denied_seam("config")),
            patch.object(bridge_client, "_validate_key", side_effect=denied_seam("key")),
            patch.object(
                bridge_client.urllib.request,
                "urlopen",
                side_effect=denied_seam("network"),
            ),
        ):
            caught = bridge_client.invoke_bridge(
                session_id,
                tool,
                arguments,
                trusted=trusted,
            )
    except bridge_client.BridgeError as exc:
        caught = exc
    except BaseException as exc:
        raise AssertionError(
            f"ordinary receiver raised {type(exc).__name__} instead of BridgeError"
        ) from exc

    assert isinstance(caught, bridge_client.BridgeError)
    assert caught.category == expected_category
    _assert_content_free_bridge_error(caught)
    assert seam_calls == []


@pytest.mark.parametrize(
    "case_name,case_factory,expected_category",
    [
        ("session_none", lambda: (None, "read_range", _ordinary_arguments()), "Validation"),
        (
            "session_bytes",
            lambda: (SESSION.encode("ascii"), "read_range", _ordinary_arguments()),
            "Validation",
        ),
        (
            "session_uuid_object",
            lambda: (uuid.UUID(SESSION), "read_range", _ordinary_arguments()),
            "Validation",
        ),
        (
            "session_hook_object",
            lambda: (_R2aReceiverHookObject(), "read_range", _ordinary_arguments()),
            "Validation",
        ),
        (
            "session_string_subclass",
            lambda: (_R2aReceiverHookString(SESSION), "read_range", _ordinary_arguments()),
            "Validation",
        ),
        (
            "session_plain_canary",
            lambda: (TRUSTED_ENVELOPE_CANARY, "read_range", _ordinary_arguments()),
            "Validation",
        ),
        (
            "session_capacity",
            lambda: (SESSION + ("x" * 256), "read_range", _ordinary_arguments()),
            "Capacity",
        ),
        ("tool_none", lambda: (SESSION, None, _ordinary_arguments()), "Validation"),
        ("tool_integer", lambda: (SESSION, 7, _ordinary_arguments()), "Validation"),
        (
            "tool_hook_object",
            lambda: (SESSION, _R2aReceiverHookObject(), _ordinary_arguments()),
            "Validation",
        ),
        (
            "tool_string_subclass",
            lambda: (SESSION, _R2aReceiverHookString("read_range"), _ordinary_arguments()),
            "Validation",
        ),
        ("tool_empty", lambda: (SESSION, "", _ordinary_arguments()), "Validation"),
        (
            "tool_capacity",
            lambda: (SESSION, "r" * 257, _ordinary_arguments()),
            "Capacity",
        ),
        ("arguments_none", lambda: (SESSION, "read_range", None), "Validation"),
        ("arguments_list", lambda: (SESSION, "read_range", []), "Validation"),
        (
            "arguments_hook_object",
            lambda: (SESSION, "read_range", _R2aReceiverHookObject()),
            "Validation",
        ),
        (
            "arguments_dict_subclass",
            lambda: (SESSION, "read_range", _R2aReceiverHookDict()),
            "Validation",
        ),
    ],
)
def test_trusted_receiver_rejects_ordinary_call_shape_before_hooks_and_side_effects(
    case_name: str, case_factory: Callable[[], tuple[Any, Any, Any]], expected_category: str
) -> None:
    del case_name
    session_id, tool, arguments = case_factory()
    module = load_plugin()
    _assert_receiver_rejects_before_side_effects(
        module.bridge_client,
        expected_category,
        session_id=session_id,
        tool=tool,
        arguments=arguments,
    )


def test_trusted_receiver_rejects_host_trust_on_literal_plugin_and_noncanonical_routes() -> None:
    module = load_plugin()
    for tool in ("sidecar_excel_execute_plan", "read_range", "sidecar_excel_read_range"):
        _assert_receiver_rejects_before_side_effects(
            module.bridge_client,
            "Validation",
            tool=tool,
            arguments=_ordinary_arguments(),
            trusted=_trusted_metadata(),
        )


def test_trusted_receiver_checks_plain_key_count_before_materializing_extra_keys() -> None:
    module = load_plugin()
    trusted = _trusted_metadata()
    trusted.update(
        {
            f"extra_{index}": TRUSTED_ENVELOPE_CANARY
            for index in range(4096)
        }
    )
    _assert_receiver_rejects_before_side_effects(
        module.bridge_client,
        "Validation",
        trusted=trusted,
    )


def test_trusted_receiver_rejects_plain_32_character_malformed_session_without_context_value_or_seams() -> None:
    module = load_plugin()
    bridge_client = module.bridge_client
    session_id = 'privatecanary'.ljust(32,'z')
    assert len(session_id) == 32
    seam_calls: list[str] = []

    def denied_seam(name: str):
        def call(*_args: Any, **_kwargs: Any) -> None:
            seam_calls.append(name)
            raise AssertionError(f"{name} seam was called for malformed session")

        return call

    caught: Any = None
    with (
        patch.object(bridge_client, "_connection", side_effect=denied_seam("config")),
        patch.object(bridge_client, "_validate_key", side_effect=denied_seam("key")),
        patch.object(
            bridge_client.urllib.request,
            "urlopen",
            side_effect=denied_seam("network"),
        ),
    ):
        try:
            bridge_client.invoke_bridge(
                session_id,
                "read_range",
                _ordinary_arguments(),
            )
        except bridge_client.BridgeError as exc:
            caught = exc

    assert caught is not None
    assert isinstance(caught, bridge_client.BridgeError)
    assert caught.category == "Validation"
    assert str(caught) == "bridgeSessionId must be a UUID"

    pending = [caught]
    seen: set[int] = set()
    while pending:
        current = pending.pop()
        if id(current) in seen:
            continue
        seen.add(id(current))
        assert session_id not in str(current)
        assert session_id not in "".join(traceback.format_exception(current))
        if current.__cause__ is not None:
            pending.append(current.__cause__)
        if current.__context__ is not None:
            pending.append(current.__context__)

    assert seam_calls == []


# R2a G5 qualification is deliberately test-only.  These stdlib imports are
# resolved before any fail-closed guard is installed; no candidate module is
# imported by this preloading block.
import builtins as _g5_builtins
import contextlib as _g5_contextlib
import hashlib as _g5_hashlib
import http.client as _g5_http_client
import importlib.util as _g5_importlib_util
import io as _g5_io
import ipaddress as _g5_ipaddress
import os as _g5_os
import pathlib as _g5_pathlib
import socket as _g5_socket
import stat as _g5_stat
import subprocess as _g5_subprocess
import threading as _g5_threading
import typing as _g5_typing
import urllib.error as _g5_url_error
import urllib.parse as _g5_url_parse
import urllib.request as _g5_url_request
import uuid as _g5_uuid
from types import ModuleType as _g5_ModuleType

_G5_INIT_PATH = PLUGIN_DIR / "__init__.py"
_G5_BRIDGE_PATH = PLUGIN_DIR / "bridge_client.py"
_G5_PREFIX_LENGTH = 44157  # NAMES: +6 bytes from desk_excel_ -> sidecar_excel_ (x2) in the pinned prefix
_G5_PREFIX_SHA256 = "437c37366a782f8ff781eeef5e1d9e93eab988d4f743c3d3d4d741a7ebddcb77"
_G5_GUARD_CANARY = "r2a-g5-synthetic-guard-canary"
_G5_SCHEMA_BYTES = SCHEMAS.read_bytes()
_G5_LOAD_SEQUENCE = 0


class _G5GuardState:
    def __init__(self, allowed_paths: set[str]) -> None:
        self.allowed_paths = allowed_paths
        self.phase = "before"
        self.events: list[str] = []
        self.selftest_denials = {name: 0 for name in ("file", "network", "connection", "process")}
        self.import_denials = {name: 0 for name in ("file", "network", "connection", "process")}
        self.receiver_denials = {name: 0 for name in ("file", "network", "connection", "process")}
        self.protected_seams: list[str] = []
        self.forbidden_original_calls = 0
        self.allowed_reads = 0
        self.restored = False

    def deny(self, kind: str, detail: str) -> None:
        if self.phase == "selftest":
            bucket = self.selftest_denials
        elif self.phase == "receiver":
            bucket = self.receiver_denials
        else:
            bucket = self.import_denials
        bucket[kind] += 1
        self.events.append(f"denied:{self.phase}:{kind}:{detail}")
        # This path never calls the saved/original operation.
        raise AssertionError(f"{_G5_GUARD_CANARY}:{kind}")

    def snapshot(self) -> dict[str, Any]:
        return {
            "events": list(self.events),
            "selftest_denials": dict(self.selftest_denials),
            "import_denials": dict(self.import_denials),
            "receiver_denials": dict(self.receiver_denials),
            "protected_seams": list(self.protected_seams),
            "forbidden_original_calls": self.forbidden_original_calls,
            "allowed_reads": self.allowed_reads,
            "restored": self.restored,
        }


def _g5_path_key(value: Any) -> str | None:
    if isinstance(value, int):
        return None
    try:
        return _g5_os.path.normpath(_g5_os.fspath(value))
    except TypeError:
        return None


def _g5_read_mode(mode: Any) -> bool:
    if not isinstance(mode, str):
        return False
    return not any(flag in mode for flag in ("w", "a", "x", "+"))


def _g5_allowed_read(state: _G5GuardState, value: Any, mode: Any) -> bool:
    return _g5_read_mode(mode) and _g5_path_key(value) in state.allowed_paths


@_g5_contextlib.contextmanager
def _g5_guards(state: _G5GuardState):
    replacements: list[tuple[Any, str, Any]] = []

    def replace(target: Any, attribute: str, replacement: Any) -> None:
        original = getattr(target, attribute)
        replacements.append((target, attribute, original))
        setattr(target, attribute, replacement)

    def guarded_open(file: Any, mode: str = "r", *args: Any, **kwargs: Any):
        if _g5_allowed_read(state, file, mode):
            state.allowed_reads += 1
            return replacements[0][2](file, mode, *args, **kwargs)
        return state.deny("file", str(_g5_path_key(file) or "<non-path>"))

    def guarded_io_open(file: Any, mode: str = "r", *args: Any, **kwargs: Any):
        if _g5_allowed_read(state, file, mode):
            state.allowed_reads += 1
            # The io.open replacement is the second registered replacement.
            return replacements[1][2](file, mode, *args, **kwargs)
        return state.deny("file", str(_g5_path_key(file) or "<non-path>"))

    def guarded_stat(path: Any, *args: Any, **kwargs: Any):
        if _g5_path_key(path) in state.allowed_paths:
            state.allowed_reads += 1
            return _g5_original_stat(path, *args, **kwargs)
        return state.deny("file", str(_g5_path_key(path) or "<non-path>"))

    def guarded_lstat(path: Any, *args: Any, **kwargs: Any):
        if _g5_path_key(path) in state.allowed_paths:
            state.allowed_reads += 1
            return _g5_original_lstat(path, *args, **kwargs)
        return state.deny("file", str(_g5_path_key(path) or "<non-path>"))

    def guarded_os_open(path: Any, flags: int, *args: Any, **kwargs: Any):
        write_flags = (
            _g5_os.O_WRONLY
            | _g5_os.O_RDWR
            | getattr(_g5_os, "O_CREAT", 0)
            | getattr(_g5_os, "O_TRUNC", 0)
            | getattr(_g5_os, "O_APPEND", 0)
            | getattr(_g5_os, "O_EXCL", 0)
        )
        if _g5_path_key(path) in state.allowed_paths and not (flags & write_flags):
            state.allowed_reads += 1
            return _g5_original_os_open(path, flags, *args, **kwargs)
        return state.deny("file", str(_g5_path_key(path) or "<non-path>"))

    def deny_network(*_args: Any, **_kwargs: Any) -> None:
        state.deny("network", "network")

    def deny_connection(*_args: Any, **_kwargs: Any) -> None:
        state.deny("connection", "connection")

    def deny_process(*_args: Any, **_kwargs: Any) -> None:
        state.deny("process", "process")

    _g5_original_stat = _g5_os.stat
    _g5_original_lstat = _g5_os.lstat
    _g5_original_os_open = _g5_os.open
    replace(_g5_builtins, "open", guarded_open)
    replace(_g5_io, "open", guarded_io_open)
    replace(_g5_os, "stat", guarded_stat)
    replace(_g5_os, "lstat", guarded_lstat)
    replace(_g5_os, "open", guarded_os_open)
    # Python 3.14 on Windows uses native metadata fast paths that bypass
    # os.stat/os.lstat. Guard their public entry points as well.
    for attribute in ("exists", "lexists", "isfile", "isdir", "islink", "isjunction"):
        if not hasattr(_g5_os.path, attribute):
            continue
        original_metadata = getattr(_g5_os.path, attribute)

        def guarded_metadata(path: Any, *args: Any, _original=original_metadata, **kwargs: Any):
            if _g5_path_key(path) in state.allowed_paths:
                state.allowed_reads += 1
                return _original(path, *args, **kwargs)
            return state.deny("file", str(_g5_path_key(path) or "<non-path>"))

        replace(_g5_os.path, attribute, guarded_metadata)
    replace(_g5_url_request, "urlopen", deny_network)
    replace(_g5_socket, "create_connection", deny_connection)
    replace(_g5_socket.socket, "connect", deny_connection)
    if hasattr(_g5_socket.socket, "connect_ex"):
        replace(_g5_socket.socket, "connect_ex", deny_connection)
    replace(_g5_http_client.HTTPConnection, "connect", deny_connection)
    if hasattr(_g5_http_client, "HTTPSConnection"):
        replace(_g5_http_client.HTTPSConnection, "connect", deny_connection)
    for attribute in ("run", "Popen", "call", "check_call", "check_output"):
        replace(_g5_subprocess, attribute, deny_process)
    for attribute in (
        "system", "popen", "execv", "execve", "execl", "execlp", "execle", "execvp",
        "spawnv", "spawnve", "spawnl", "spawnle", "spawnlp", "spawnlpe",
    ):
        if hasattr(_g5_os, attribute):
            replace(_g5_os, attribute, deny_process)

    state.events.append("guards-installed")
    try:
        yield state
    finally:
        for target, attribute, original in reversed(replacements):
            setattr(target, attribute, original)
        state.restored = True
        state.events.append("guards-restored")


def _g5_expect_denial(state: _G5GuardState, kind: str, operation: Callable[[], Any]) -> None:
    try:
        operation()
    except AssertionError as exc:
        assert _G5_GUARD_CANARY in str(exc)
        assert state.forbidden_original_calls == 0
    else:
        raise AssertionError(f"synthetic {kind} guard did not deny")


def _g5_run_guard_selftests(state: _G5GuardState) -> None:
    state.phase = "selftest"
    state.events.append("guard-selftest-start")
    _g5_expect_denial(
        state,
        "config-read",
        lambda: Path("/r2a-g5-forbidden/excel.json").read_text(encoding="utf-8"),
    )
    _g5_expect_denial(
        state,
        "key-read",
        lambda: Path("/r2a-g5-forbidden/excel.key").read_bytes(),
    )
    _g5_expect_denial(
        state,
        "arbitrary-stat",
        lambda: Path("/r2a-g5-forbidden/anything").exists(),
    )
    _g5_expect_denial(
        state,
        "arbitrary-write",
        lambda: Path("/r2a-g5-forbidden/anything").write_text("blocked", encoding="utf-8"),
    )
    _g5_expect_denial(
        state,
        "network-request",
        lambda: _g5_url_request.urlopen("http://198.51.100.1:9"),
    )
    _g5_expect_denial(
        state,
        "network-connection",
        lambda: _g5_socket.create_connection(("198.51.100.1", 9), timeout=0.01),
    )
    _g5_expect_denial(
        state,
        "process-launch",
        lambda: _g5_subprocess.run(["/r2a-g5-forbidden/process"]),
    )
    assert all(value > 0 for value in state.selftest_denials.values())
    assert state.forbidden_original_calls == 0
    state.events.append("guard-selftest-complete")
    state.phase = "import"


def _g5_blocked_candidate_seam(state: _G5GuardState, name: str):
    def blocked(*_args: Any, **_kwargs: Any) -> None:
        state.protected_seams.append(name)
        raise AssertionError(f"{_G5_GUARD_CANARY}:{name}")

    return blocked


def _g5_exec_source(module: _g5_ModuleType, path: Path, source: bytes) -> None:
    code = compile(source, str(path), "exec")
    exec(code, module.__dict__)


def _g5_guarded_load_plugin():
    global _G5_LOAD_SEQUENCE
    sequence = _G5_LOAD_SEQUENCE
    _G5_LOAD_SEQUENCE += 1
    init_source = _G5_INIT_PATH.read_bytes()
    bridge_source = _G5_BRIDGE_PATH.read_bytes()
    schema_source = SCHEMAS.read_bytes()
    allowed_paths = {str(_G5_INIT_PATH), str(_G5_BRIDGE_PATH), str(SCHEMAS)}
    state = _G5GuardState(allowed_paths)
    package_name = f"hermes_office_excel_r2a_g5_{sequence}_{_g5_uuid.uuid4().hex}"
    bridge_name = f"{package_name}.bridge_client"
    package: _g5_ModuleType | None = None
    bridge: _g5_ModuleType | None = None
    bridge_patches: list[tuple[Any, str, Any]] = []
    receiver_error: BaseException | None = None

    def restore_bridge_patches() -> None:
        for target, attribute, original in reversed(bridge_patches):
            setattr(target, attribute, original)
        bridge_patches.clear()

    def cleanup_candidate_cache() -> None:
        _g5_sys_modules = sys.modules
        _g5_sys_modules.pop(bridge_name, None)
        _g5_sys_modules.pop(package_name, None)
        for suffix in ("desk_app", "desk_bridge"):
            _g5_sys_modules.pop(package_name + "." + suffix, None)
        state.events.append("candidate-cache-clean")

    import desk_plugin
    for submodule in ("desk_app", "desk_bridge"):
        sys.modules[package_name + "." + submodule] = getattr(desk_plugin, submodule)
    with _g5_guards(state):
        try:
            _g5_run_guard_selftests(state)
            state.events.append("bridge-exec-start")
            bridge_spec = _g5_importlib_util.spec_from_file_location(bridge_name, str(_G5_BRIDGE_PATH))
            assert bridge_spec is not None
            bridge = _g5_importlib_util.module_from_spec(bridge_spec)
            sys.modules[bridge_name] = bridge
            _g5_exec_source(bridge, _G5_BRIDGE_PATH, bridge_source)
            assert bridge.__file__ == str(_G5_BRIDGE_PATH)
            assert bridge.invoke_bridge.__code__.co_filename == str(_G5_BRIDGE_PATH)
            state.events.append("bridge-exec-complete")

            bridge_patches.extend(
                [
                    (bridge, "_connection", bridge._connection),
                    (bridge, "_validate_key", bridge._validate_key),
                    (bridge.urllib.request, "urlopen", bridge.urllib.request.urlopen),
                ]
            )
            setattr(bridge, "_connection", _g5_blocked_candidate_seam(state, "config"))
            setattr(bridge, "_validate_key", _g5_blocked_candidate_seam(state, "key"))
            # Rebind the real bridge module's request entry point before the
            # package executes; the installed guard remains the only network path.
            bridge.urllib.request.urlopen = _g5_url_request.urlopen
            state.events.append("receiver-protections-installed")

            package_spec = _g5_importlib_util.spec_from_file_location(
                package_name,
                str(_G5_INIT_PATH),
                submodule_search_locations=[str(PLUGIN_DIR)],
            )
            assert package_spec is not None
            package = _g5_importlib_util.module_from_spec(package_spec)
            sys.modules[package_name] = package
            state.events.append("package-exec-start")
            _g5_exec_source(package, _G5_INIT_PATH, init_source)
            state.events.append("package-exec-complete")
            assert package.__file__ == str(_G5_INIT_PATH)
            assert package.bridge_client is bridge
            assert package._SCHEMA_PATH == SCHEMAS
            assert package._SCHEMAS == json.loads(schema_source.decode("utf-8"))

            state.phase = "receiver"
            state.events.append("receiver-qualification-start")
            try:
                bridge.invoke_bridge(
                    None,
                    "read_range",
                    {"sheet": "Sheet1", "address": "A1"},
                    trusted=None,
                )
            except bridge.BridgeError as exc:
                assert exc.category == "Validation"
                assert str(exc) == "bridgeSessionId must be a UUID"
                receiver_error = exc
            else:
                raise AssertionError("guarded receiver accepted invalid session input")
            assert state.protected_seams == []
            assert state.receiver_denials == {
                "file": 0,
                "network": 0,
                "connection": 0,
                "process": 0,
            }
            state.events.append("receiver-qualification-complete")

            cleanup_candidate_cache()
            qualification = state.snapshot()
            qualification.update(
                {
                    "package_name": package_name,
                    "bridge_name": bridge_name,
                    "cache_clean": package_name not in sys.modules and bridge_name not in sys.modules,
                    "init_sha256": _g5_hashlib.sha256(init_source).hexdigest(),
                    "bridge_sha256": _g5_hashlib.sha256(bridge_source).hexdigest(),
                    "schema_sha256": _g5_hashlib.sha256(schema_source).hexdigest(),
                    "source_paths": [str(_G5_INIT_PATH), str(_G5_BRIDGE_PATH)],
                    "schema_path": str(SCHEMAS),
                }
            )
            setattr(package, "_r2a_g5_guard_qualification", qualification)
        finally:
            restore_bridge_patches()
            cleanup_candidate_cache()

    assert receiver_error is not None
    assert _G5_GUARD_CANARY not in str(receiver_error)
    assert _G5_GUARD_CANARY not in "".join(traceback.format_exception(receiver_error))
    # Transaction tests inject a memory-only endpoint, not a real Desk listener.
    # Real startup, approval, paired-session dispatch and revocation are exercised by the security suite.
    class TransportTLS:
        # This memory-only seam targets a standalone HTTP companion or a mocked urlopen.
        # Production DeskRuntime always supplies its per-install verified TLS context.
        @staticmethod
        def client_context():
            return None
    class TransportBridge:
        tls = TransportTLS()
        port = 1
        @property
        def base_url(self):
            return bridge._connection()[0]
        @property
        def plugin_key(self):
            return bridge._connection()[1]
        def has_session(self, session_id):
            return True
        def session_allowed(self, session_id):
            # This explicit transaction-only double represents an approved session.
            # Real persisted-grant checks and cross-process revoke run in
            # test_desk_bridge_security; no approval fixture is injected there.
            return self.has_session(session_id)
    class TransportRuntime:
        bridge = TransportBridge()
        key_provider = staticmethod(lambda: None)
        @classmethod
        def from_ctx(cls, ctx, app):
            return cls()
        def maybe_autostart(self):
            return None
    package.DeskRuntime = TransportRuntime
    return package


# Existing focused imports resolve this late-bound name at call time.  The
# historical prefix and every existing assertion/case therefore remain intact,
# while each candidate import receives a fresh guarded namespace.
load_plugin = _g5_guarded_load_plugin


def test_trusted_receiver_guard_qualification_precedes_candidate_execution_and_restores() -> None:
    original_open = _g5_builtins.open
    original_io_open = _g5_io.open
    original_urlopen = _g5_url_request.urlopen
    original_run = _g5_subprocess.run
    module = load_plugin()
    qualification = module._r2a_g5_guard_qualification
    events = qualification["events"]
    order = [
        "guards-installed",
        "guard-selftest-start",
        "guard-selftest-complete",
        "bridge-exec-start",
        "bridge-exec-complete",
        "receiver-protections-installed",
        "package-exec-start",
        "package-exec-complete",
        "receiver-qualification-start",
        "receiver-qualification-complete",
        "candidate-cache-clean",
    ]
    positions = [events.index(marker) for marker in order]
    assert positions == sorted(positions)
    assert qualification["selftest_denials"] == {
        "file": 4,
        "network": 1,
        "connection": 1,
        "process": 1,
    }
    assert qualification["import_denials"] == {
        "file": 0,
        "network": 0,
        "connection": 0,
        "process": 0,
    }
    assert qualification["receiver_denials"] == {
        "file": 0,
        "network": 0,
        "connection": 0,
        "process": 0,
    }
    assert qualification["protected_seams"] == []
    assert qualification["forbidden_original_calls"] == 0
    assert qualification["cache_clean"] is True
    assert qualification["restored"] is False
    assert _g5_builtins.open is original_open
    assert _g5_io.open is original_io_open
    assert _g5_url_request.urlopen is original_urlopen
    assert _g5_subprocess.run is original_run
    assert module.__file__ == str(_G5_INIT_PATH)
    assert module.bridge_client.__file__ == str(_G5_BRIDGE_PATH)
    assert module.bridge_client.invoke_bridge.__code__.co_filename == str(_G5_BRIDGE_PATH)
    assert module._SCHEMA_PATH == SCHEMAS
    assert module._SCHEMAS == json.loads(_G5_SCHEMA_BYTES.decode("utf-8"))


def test_trusted_receiver_remains_under_fail_closed_guards_after_import() -> None:
    module = load_plugin()
    bridge_client = module.bridge_client
    state = _G5GuardState({str(_G5_INIT_PATH), str(_G5_BRIDGE_PATH), str(SCHEMAS)})
    seam_calls: list[str] = []
    caught: BaseException | None = None

    def denied_seam(name: str):
        def call(*_args: Any, **_kwargs: Any) -> None:
            seam_calls.append(name)
            raise AssertionError(f"{_G5_GUARD_CANARY}:{name}")

        return call

    with _g5_guards(state):
        state.phase = "receiver"
        with (
            patch.object(bridge_client, "_connection", side_effect=denied_seam("config")),
            patch.object(bridge_client, "_validate_key", side_effect=denied_seam("key")),
        ):
            try:
                bridge_client.invoke_bridge(
                    None,
                    "read_range",
                    {"sheet": "Sheet1", "address": "A1"},
                    trusted=None,
                )
            except bridge_client.BridgeError as exc:
                assert exc.category == "Validation"
                assert str(exc) == "bridgeSessionId must be a UUID"
                caught = exc
            else:
                raise AssertionError("guarded trusted receiver accepted invalid input")

    assert caught is not None
    assert _G5_GUARD_CANARY not in str(caught)
    assert _G5_GUARD_CANARY not in "".join(traceback.format_exception(caught))
    assert seam_calls == []
    assert state.receiver_denials == {
        "file": 0,
        "network": 0,
        "connection": 0,
        "process": 0,
    }
    assert state.forbidden_original_calls == 0
    assert state.restored is True


def test_trusted_receiver_g5_g1_ported_prefix_is_pinned() -> None:
    current = Path(__file__).read_bytes()
    assert len(current) >= _G5_PREFIX_LENGTH
    assert _g5_hashlib.sha256(current[:_G5_PREFIX_LENGTH]).hexdigest() == _G5_PREFIX_SHA256



# ---------------------------------------------------------------------------
# R2d-2: machine-readable plan uncertainty through the plugin bridge client
# ---------------------------------------------------------------------------

import http.client as _r2d2_http_client
import io as _r2d2_io
import urllib.error as _r2d2_url_error

_R2D2_CANARY = "r2d2-workbook-content-canary"
_R2D2_OPERATION = str(uuid.uuid5(uuid.UUID("6f37c2f0-29df-5c72-b6bb-e89f2f18f0c5"), "r2d2-operation"))


def _r2d2_trusted() -> dict[str, str]:
    return {
        "turn_operation_id": _R2D2_OPERATION,
        "request_id": _R2D2_OPERATION,
        "workbook_plan_fingerprint": generated_fingerprint(),
        "phase": "inspect",
        "turn_key": _R4G2_TURN_KEY,
    }


def _r2d2_http_error(code: int, payload: bytes) -> _r2d2_url_error.HTTPError:
    return _r2d2_url_error.HTTPError(
        "http://127.0.0.1:8650/excel/invoke",
        code,
        "error",
        cast(Any, {"Content-Length": str(len(payload))}),
        _r2d2_io.BytesIO(payload),
    )


class _R2D2DroppedResponse:
    headers = {"Content-Length": "4096"}

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def read(self, _limit: int) -> bytes:
        raise ConnectionResetError(_R2D2_CANARY)


def _r2d2_invoke(bridge_client: Any, tool: str, side_effect: Any) -> tuple[Any, int]:
    calls: list[Any] = []

    def fake_urlopen(request: Any, timeout: int):
        calls.append(request)
        if isinstance(side_effect, BaseException):
            raise side_effect
        return side_effect

    plan = tool == "execute_plan"
    with (
        patch.object(
            bridge_client,
            "_connection",
            return_value=("http://127.0.0.1:8650", "k" * 32),
        ),
        patch.object(bridge_client, "_validate_key", return_value="k" * 32),
        patch.object(bridge_client.urllib.request, "urlopen", side_effect=fake_urlopen),
    ):
        try:
            outcome: Any = bridge_client.invoke_bridge(
                SESSION,
                tool,
                valid_plan() if plan else {"sheet": "Sheet1", "address": "A1"},
                trusted=_r2d2_trusted() if plan else None,
            )
        except bridge_client.BridgeError as exc:
            outcome = exc
    return outcome, len(calls)


def test_r2d2_plan_uncertain_companion_status_is_its_own_category_and_not_retried() -> None:
    bridge_client = load_plugin().bridge_client
    body = json.dumps({
        "ok": False,
        "error": "Workbook plan execution outcome is uncertain",
        "category": "ExecutionUncertain",
        "state": "executing_uncertain",
    }).encode("utf-8")
    # A plan 504 without a companion category is never downgraded to Timeout.
    outcome, calls = _r2d2_invoke(bridge_client, "execute_plan", _r2d2_http_error(504, b""))
    assert isinstance(outcome, bridge_client.BridgeError)
    assert outcome.category == "ExecutionUncertain"
    assert calls == 1

    outcome, calls = _r2d2_invoke(bridge_client, "execute_plan", _r2d2_http_error(504, body))
    assert isinstance(outcome, bridge_client.BridgeError)
    assert outcome.category == "ExecutionUncertain"
    assert str(outcome) == bridge_client.PLAN_UNCERTAIN_MESSAGE
    assert calls == 1


def test_r2d2_plan_not_started_timeout_stays_timeout() -> None:
    bridge_client = load_plugin().bridge_client
    body = json.dumps({
        "ok": False, "error": "Invoke timed out", "category": "Timeout", "state": "not_started",
    }).encode("utf-8")
    outcome, calls = _r2d2_invoke(bridge_client, "execute_plan", _r2d2_http_error(504, body))
    assert isinstance(outcome, bridge_client.BridgeError)
    assert outcome.category == "Timeout"
    assert calls == 1


def test_r2d2_plan_mid_flight_disconnect_is_uncertain_without_retry() -> None:
    bridge_client = load_plugin().bridge_client
    for failure in (
        _r2d2_http_client.RemoteDisconnected(_R2D2_CANARY),
        ConnectionResetError(_R2D2_CANARY),
        TimeoutError(_R2D2_CANARY),
        _R2D2DroppedResponse(),
        _r2d2_http_client.IncompleteRead(_R2D2_CANARY.encode("utf-8")),
    ):
        outcome, calls = _r2d2_invoke(bridge_client, "execute_plan", failure)
        assert isinstance(outcome, bridge_client.BridgeError), (failure, outcome)
        assert outcome.category == "ExecutionUncertain", (failure, outcome.category)
        assert _R2D2_CANARY not in str(outcome)
        assert calls == 1


def test_r2d2_plan_connect_failure_before_send_is_unavailable() -> None:
    bridge_client = load_plugin().bridge_client
    outcome, calls = _r2d2_invoke(
        bridge_client, "execute_plan", _r2d2_url_error.URLError(ConnectionRefusedError(_R2D2_CANARY)),
    )
    assert isinstance(outcome, bridge_client.BridgeError)
    assert outcome.category == "Unavailable"
    assert calls == 1


def test_r2d2_primitive_transport_categories_are_unchanged() -> None:
    bridge_client = load_plugin().bridge_client
    for failure in (
        _r2d2_http_client.RemoteDisconnected(_R2D2_CANARY),
        ConnectionResetError(_R2D2_CANARY),
        TimeoutError(_R2D2_CANARY),
        _R2D2DroppedResponse(),
        _r2d2_url_error.URLError(ConnectionRefusedError(_R2D2_CANARY)),
    ):
        outcome, calls = _r2d2_invoke(bridge_client, "read_range", failure)
        assert isinstance(outcome, bridge_client.BridgeError)
        assert outcome.category == "Unavailable"
        assert calls == 1
    outcome, calls = _r2d2_invoke(bridge_client, "read_range", _r2d2_http_error(504, b""))
    assert isinstance(outcome, bridge_client.BridgeError)
    assert outcome.category == "Timeout"
    assert calls == 1


def _r2d2_plan_handler_result(module: Any, context: Any, bridge: Any, plan: dict | None = None) -> Any:
    # R4g-2c-r1 rule W1: optional `plan` (default unchanged: valid_plan()) so the wire leg can view a recorded invoke
    # under the plan it actually sent (the plugin's executed:false filter is phase-aware since R4g-2c-r1).
    hook = context.hooks["pre_tool_call"]
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    plan = valid_plan() if plan is None else plan
    with patch.object(module, "invoke_bridge", bridge):
        hook(tool_name=PLAN_TOOL, args=copy.deepcopy(plan), session_id=SESSION, turn_id="r2d2-turn",
             tool_call_id="call-r2d2", api_request_id="request-r2d2", task_id="task-r2d2")
        return json.loads(handler(copy.deepcopy(plan), session_id=SESSION))


def test_r2d2_plan_handler_unwraps_not_executed_completed_and_uncertain() -> None:
    module, context = registered()
    # R4g-2b rule R: a not-executed report is a refusal (ok False) or a verified inspection body; the synthetic
    # report gains "ok": False so this row keeps pinning the unchanged unwrap (an ok-less body is now refused, P1).
    # R4g-2c rule R2: an ok-false report passes only in a closed R4h T1/T2 shape; the synthetic report becomes the
    # route's T2 `rejected` body so this row keeps pinning the unchanged unwrap (the R4g-2b form is refused, q1c).
    report = {"ok": False, "action": "workbook_plan", "reason": "rejected",
              "summary": "The workbook plan was rejected before execution; nothing was executed."}
    not_executed = _r2d2_plan_handler_result(
        module, context,
        lambda *_a, **_kw: {
            "ok": True, "ticketId": SESSION, "requestId": _R2D2_OPERATION,
            "data": report, "executed": False,
        },
    )
    assert not_executed == {"executed": False, "result": report}

    module, context = registered()
    completed = _r2d2_plan_handler_result(
        module, context,
        lambda *_a, **_kw: {
            "ok": True, "ticketId": SESSION, "requestId": _R2D2_OPERATION,
            "data": {"outcome": "committed"}, "executed": True,
        },
    )
    assert completed == {"outcome": "committed"}

    module, context = registered()
    bridge_client = module.bridge_client

    def uncertain(*_a: Any, **_kw: Any) -> Any:
        raise bridge_client.BridgeError("ExecutionUncertain", bridge_client.PLAN_UNCERTAIN_MESSAGE)

    result = _r2d2_plan_handler_result(module, context, uncertain)
    assert result == {
        "ok": False,
        "error": {"category": "ExecutionUncertain", "message": bridge_client.PLAN_UNCERTAIN_MESSAGE},
    }
    assert _R2D2_CANARY not in json.dumps(result)


def test_r2d2_plan_handler_real_bridge_uncertain_is_content_free() -> None:
    module, context = registered()
    bridge_client = module.bridge_client
    hook = context.hooks["pre_tool_call"]
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    calls: list[Any] = []

    def fake_urlopen(request: Any, timeout: int):
        calls.append(request)
        raise _r2d2_http_client.RemoteDisconnected(_R2D2_CANARY)

    with (
        patch.object(
            bridge_client,
            "_connection",
            return_value=("http://127.0.0.1:8650", "k" * 32),
        ),
        patch.object(bridge_client, "_validate_key", return_value="k" * 32),
        patch.object(bridge_client.urllib.request, "urlopen", side_effect=fake_urlopen),
    ):
        hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="r2d2-real",
             tool_call_id="call-r2d2-real", api_request_id="request-r2d2-real", task_id="task-r2d2")
        raw = handler(valid_plan(), session_id=SESSION)
    result = json.loads(raw)
    assert result["ok"] is False
    assert result["error"]["category"] == "ExecutionUncertain"
    assert _R2D2_CANARY not in raw
    assert len(calls) == 1



# ---------------------------------------------------------------------------
# R2d-2 G2: a delivered plan response that is truncated or undecodable is
# ExecutionUncertain. Every case below uses a real in-test loopback socket
# server; urlopen is never monkeypatched.
# ---------------------------------------------------------------------------

import socket as _r2d2g2_socket
import threading as _r2d2g2_threading

_R2D2G2_CANARY = "r2d2g2-workbook-content-canary"


class _R2D2G2Server:
    """Threaded loopback server that replies with exact raw bytes, then closes."""

    def __init__(self, raw_response: bytes) -> None:
        self.raw_response = raw_response
        self.connections = 0
        self.requests: list[bytes] = []
        self._sock = _r2d2g2_socket.socket(_r2d2g2_socket.AF_INET, _r2d2g2_socket.SOCK_STREAM)
        self._sock.setsockopt(_r2d2g2_socket.SOL_SOCKET, _r2d2g2_socket.SO_REUSEADDR, 1)
        self._sock.bind(("127.0.0.1", 0))
        self._sock.listen(8)
        self._sock.settimeout(0.2)
        self.port = self._sock.getsockname()[1]
        self._stop = _r2d2g2_threading.Event()
        self._thread = _r2d2g2_threading.Thread(target=self._serve, daemon=True)

    def __enter__(self) -> "_R2D2G2Server":
        self._thread.start()
        return self

    def __exit__(self, *_args: Any) -> None:
        self._stop.set()
        self._thread.join(timeout=5)
        self._sock.close()

    def _serve(self) -> None:
        while not self._stop.is_set():
            try:
                conn, _addr = self._sock.accept()
            except (TimeoutError, _r2d2g2_socket.timeout):
                continue
            except OSError:
                return
            self.connections += 1
            with conn:
                conn.settimeout(5)
                self.requests.append(self._read_request(conn))
                try:
                    conn.sendall(self.raw_response)
                    conn.shutdown(_r2d2g2_socket.SHUT_WR)
                except OSError:
                    # The client may stop reading an oversize body early.
                    pass

    @staticmethod
    def _read_request(conn: Any) -> bytes:
        # Consume the whole request so closing never turns into a reset.
        data = b""
        while b"\r\n\r\n" not in data:
            chunk = conn.recv(65536)
            if not chunk:
                return data
            data += chunk
        head, _sep, body = data.partition(b"\r\n\r\n")
        length = 0
        for line in head.split(b"\r\n")[1:]:
            name, _colon, value = line.partition(b":")
            if name.strip().lower() == b"content-length":
                length = int(value.strip())
        while len(body) < length:
            chunk = conn.recv(65536)
            if not chunk:
                break
            body += chunk
        return head + b"\r\n\r\n" + body


def _r2d2g2_raw(status: str, body: bytes, *, content_length: Any = "auto") -> bytes:
    lines = [f"HTTP/1.1 {status}", "Content-Type: application/json", "Connection: close"]
    if content_length == "auto":
        lines.append(f"Content-Length: {len(body)}")
    elif content_length is not None:
        lines.append(f"Content-Length: {content_length}")
    return ("\r\n".join(lines) + "\r\n\r\n").encode("ascii") + body


def _r2d2g2_invoke(raw_response: bytes, tool: str = "execute_plan") -> tuple[Any, int]:
    bridge_client = load_plugin().bridge_client
    plan = tool == "execute_plan"
    with _R2D2G2Server(raw_response) as server:
        with (
            patch.object(
                bridge_client,
                "_connection",
                return_value=(f"http://127.0.0.1:{server.port}", "k" * 32),
            ),
            patch.object(bridge_client, "_validate_key", return_value="k" * 32),
        ):
            try:
                outcome: Any = bridge_client.invoke_bridge(
                    SESSION,
                    tool,
                    valid_plan() if plan else {"sheet": "Sheet1", "address": "A1"},
                    trusted=_r2d2_trusted() if plan else None,
                )
            except bridge_client.BridgeError as exc:
                outcome = exc
        connections = server.connections
        assert server.requests and b"POST /excel/invoke " in server.requests[0]
    return outcome, connections


def _r2d2g2_assert_uncertain(outcome: Any, connections: int) -> None:
    bridge_client = load_plugin().bridge_client
    assert type(outcome).__name__ == "BridgeError", outcome
    assert outcome.category == "ExecutionUncertain", (outcome.category, str(outcome))
    assert str(outcome) == bridge_client.PLAN_UNCERTAIN_MESSAGE
    assert _R2D2G2_CANARY not in str(outcome)
    assert outcome.__cause__ is None
    assert connections == 1


_R2D2G2_COMPLETED = json.dumps({
    "ok": True,
    "ticketId": SESSION,
    "requestId": _R2D2_OPERATION,
    "data": {"outcome": "committed", "note": _R2D2G2_CANARY},
    "executed": True,
}).encode("utf-8")


def test_r2d2g2_plan_200_short_body_against_content_length_is_uncertain() -> None:
    # (a) Content-Length N, fewer than N bytes, then close.
    outcome, connections = _r2d2g2_invoke(
        _r2d2g2_raw("200 OK", _R2D2G2_COMPLETED[:20], content_length=len(_R2D2G2_COMPLETED)),
    )
    _r2d2g2_assert_uncertain(outcome, connections)


def test_r2d2g2_plan_200_truncated_json_without_content_length_is_uncertain() -> None:
    # (b) No Content-Length, truncated JSON body, then close.
    outcome, connections = _r2d2g2_invoke(
        _r2d2g2_raw("200 OK", _R2D2G2_COMPLETED[:-7], content_length=None),
    )
    _r2d2g2_assert_uncertain(outcome, connections)


def test_r2d2g2_plan_200_complete_but_undecodable_body_is_uncertain() -> None:
    # (c) Complete non-JSON, non-UTF-8 and non-object bodies.
    for body in (
        f"<html>{_R2D2G2_CANARY}</html>".encode("utf-8"),
        b"\xff\xfe" + _R2D2G2_CANARY.encode("utf-8"),
        json.dumps([_R2D2G2_CANARY]).encode("utf-8"),
        b"",
    ):
        outcome, connections = _r2d2g2_invoke(_r2d2g2_raw("200 OK", body))
        _r2d2g2_assert_uncertain(outcome, connections)


def test_r2d2g2_plan_200_unprovable_length_is_uncertain_but_declared_oversize_stays_capacity() -> None:
    bridge_client = load_plugin().bridge_client
    # A body that runs past the bound after headers was delivered: the
    # companion may already have executed, so it is uncertain, not Capacity.
    oversize = b"[" + b" " * bridge_client.MAX_RESPONSE_BYTES + b"]"
    outcome, connections = _r2d2g2_invoke(_r2d2g2_raw("200 OK", oversize, content_length=None))
    _r2d2g2_assert_uncertain(outcome, connections)
    # An unparseable Content-Length on a delivered 200 is uncertain too.
    outcome, connections = _r2d2g2_invoke(
        _r2d2g2_raw("200 OK", _R2D2G2_COMPLETED, content_length="12x"),
    )
    _r2d2g2_assert_uncertain(outcome, connections)
    # A declared oversize length is rejected before reading: still Capacity.
    outcome, connections = _r2d2g2_invoke(
        _r2d2g2_raw("200 OK", b"{}", content_length=bridge_client.MAX_RESPONSE_BYTES + 1),
    )
    assert type(outcome).__name__ == "BridgeError"
    assert outcome.category == "Capacity"
    assert connections == 1


def test_r2d2g2_plan_5xx_truncated_or_undecodable_body_is_uncertain() -> None:
    # (d) 504 with a truncated body, plus the A3 cases for 5xx error bodies.
    uncertain_504 = json.dumps({
        "ok": False,
        "error": "Workbook plan execution outcome is uncertain",
        "category": "ExecutionUncertain",
        "state": "executing_uncertain",
    }).encode("utf-8")
    internal = json.dumps({
        "ok": False, "error": "Internal Server Error", "category": "Internal",
    }).encode("utf-8")
    for raw in (
        _r2d2g2_raw("504 Gateway Timeout", uncertain_504[:15], content_length=len(uncertain_504)),
        _r2d2g2_raw("504 Gateway Timeout", uncertain_504[:-3], content_length=None),
        _r2d2g2_raw("504 Gateway Timeout", uncertain_504, content_length="nope"),
        _r2d2g2_raw("500 Internal Server Error", internal[:12], content_length=len(internal)),
        _r2d2g2_raw("500 Internal Server Error", f"<{_R2D2G2_CANARY}>".encode("utf-8")),
        _r2d2g2_raw("502 Bad Gateway", b"", content_length=None),
    ):
        outcome, connections = _r2d2g2_invoke(raw)
        _r2d2g2_assert_uncertain(outcome, connections)


def test_r2d2g2_plan_4xx_truncated_body_keeps_status_derived_category() -> None:
    # (e) A 4xx is issued before any execute-start (or for a cancelled
    # unclaimed/claimed ticket), so an unreadable body keeps its status category.
    conflict = json.dumps({
        "ok": False, "error": "requestId reused with different payload", "category": "Conflict",
    }).encode("utf-8")
    cancelled = json.dumps({
        "ok": False, "error": "Cancelled", "category": "Cancellation",
    }).encode("utf-8")
    cases = (
        (_r2d2g2_raw("409 Conflict", conflict[:10], content_length=len(conflict)), "Bridge"),
        (_r2d2g2_raw("409 Conflict", conflict[:-4], content_length=None), "Bridge"),
        (_r2d2g2_raw("400 Bad Request", f"<{_R2D2G2_CANARY}>".encode("utf-8")), "Bridge"),
        (_r2d2g2_raw("499 Client Closed Request", cancelled[:9], content_length=len(cancelled)),
         "Cancellation"),
    )
    for raw, category in cases:
        outcome, connections = _r2d2g2_invoke(raw)
        assert type(outcome).__name__ == "BridgeError", outcome
        assert outcome.category == category, (raw[:40], outcome.category)
        assert outcome.category != "ExecutionUncertain"
        assert _R2D2G2_CANARY not in str(outcome)
        assert connections == 1
    # A complete 4xx keeps the companion's own category and message.
    outcome, connections = _r2d2g2_invoke(_r2d2g2_raw("409 Conflict", conflict))
    assert outcome.category == "Conflict"
    assert str(outcome) == "requestId reused with different payload"
    assert connections == 1


def test_r2d2g2_plan_complete_valid_responses_keep_their_mapping() -> None:
    # (f) Regression: complete valid responses are unchanged.
    outcome, connections = _r2d2g2_invoke(_r2d2g2_raw("200 OK", _R2D2G2_COMPLETED))
    assert outcome == json.loads(_R2D2G2_COMPLETED)
    assert connections == 1
    outcome, connections = _r2d2g2_invoke(
        _r2d2g2_raw("200 OK", _R2D2G2_COMPLETED, content_length=None),
    )
    assert outcome == json.loads(_R2D2G2_COMPLETED)
    assert connections == 1
    not_started = json.dumps({
        "ok": False, "error": "Invoke timed out", "category": "Timeout", "state": "not_started",
    }).encode("utf-8")
    outcome, connections = _r2d2g2_invoke(_r2d2g2_raw("504 Gateway Timeout", not_started))
    assert outcome.category == "Timeout"
    assert connections == 1
    outcome, connections = _r2d2g2_invoke(_r2d2g2_raw("504 Gateway Timeout", b"{}"))
    _r2d2g2_assert_uncertain(outcome, connections)


def test_r2d2g2_primitive_truncated_200_is_unchanged_protocol() -> None:
    # (g) The primitive path is unchanged from the parent.
    body = json.dumps({"ok": True, "data": {"values": [[_R2D2G2_CANARY]]}}).encode("utf-8")
    outcome, connections = _r2d2g2_invoke(
        _r2d2g2_raw("200 OK", body[:10], content_length=len(body)), tool="read_range",
    )
    assert type(outcome).__name__ == "BridgeError"
    assert outcome.category == "Protocol"
    assert str(outcome) == "Companion returned invalid JSON"
    assert connections == 1
    outcome, connections = _r2d2g2_invoke(
        _r2d2g2_raw("500 Internal Server Error", body[:10], content_length=len(body)), tool="read_range",
    )
    assert outcome.category == "Bridge"
    assert connections == 1


# ---------------------------------------------------------------------------
# R3a (G8 / G2-A1, Q5): every complete, decodable plan 5xx is uncertain except
# the companion's own never-started 504 shape; the companion, not the plugin's
# advisory seen-turn set, is the single authority for plan-op acceptance.
# ---------------------------------------------------------------------------

import time as _r3a_time

from desk_plugin.companion.handler import CompanionHandler as _R3aHandler
from desk_plugin.companion.server import ExcelCompanionServer as _R3aServer
from scripts.tests.test_companion_plan_execution import (
    PLUGIN_KEY as _R3A_PLUGIN_KEY,
    _pane_headers as _r3a_pane_headers,
    _request as _r3a_request,
    _session as _r3a_session,
)

_R3A_CAPACITY_MESSAGE = "Workbook plan capacity reached for this session"


def _r3a_json(value: dict) -> bytes:
    return json.dumps(value).encode("utf-8")


def test_r3a_ra7_complete_plan_5xx_is_uncertain_unless_exact_not_started() -> None:
    for status, body in (
        ("500 Internal Server Error", {"ok": False, "error": "Internal Server Error", "category": "Internal"}),
        ("500 Internal Server Error", {"ok": False, "error": "Unknown ticket state", "category": "Internal"}),
        ("503 Service Unavailable", {"ok": False, "error": "x", "category": "Capacity"}),
        ("502 Bad Gateway", {"ok": False, "error": "x", "category": "Bridge"}),
        ("504 Gateway Timeout", {"ok": False, "error": "Invoke timed out", "category": "Timeout"}),
        ("500 Internal Server Error",
         {"ok": False, "error": "x", "category": "Timeout", "state": "not_started"}),
    ):
        outcome, connections = _r2d2g2_invoke(_r2d2g2_raw(status, _r3a_json(body)))
        _r2d2g2_assert_uncertain(outcome, connections)


def test_r3a_gp6_plan_not_started_and_4xx_and_primitive_5xx_keep_their_categories() -> None:
    outcome, connections = _r2d2g2_invoke(_r2d2g2_raw("504 Gateway Timeout", _r3a_json({
        "ok": False, "error": "Invoke timed out", "category": "Timeout", "state": "not_started",
    })))
    assert (outcome.category, str(outcome), connections) == ("Timeout", "Invoke timed out", 1)
    outcome, connections = _r2d2g2_invoke(_r2d2g2_raw("429 Too Many Requests", _r3a_json({
        "ok": False, "error": _R3A_CAPACITY_MESSAGE, "category": "Capacity", "state": "plan_fence_capacity",
    })))
    assert (outcome.category, str(outcome), connections) == ("Capacity", _R3A_CAPACITY_MESSAGE, 1)
    outcome, connections = _r2d2g2_invoke(_r2d2g2_raw("409 Conflict", _r3a_json({
        "ok": False, "error": "requestId reused with different payload", "category": "Conflict",
    })))
    assert (outcome.category, str(outcome), connections) == (
        "Conflict", "requestId reused with different payload", 1)
    outcome, connections = _r2d2g2_invoke(_r2d2g2_raw("500 Internal Server Error", _r3a_json({
        "ok": False, "error": "Internal Server Error", "category": "Internal",
    })), tool="read_range")
    assert (outcome.category, str(outcome), connections) == ("Internal", "Internal Server Error", 1)


class _R3aCompanion:
    """A real loopback companion plus a pane driver and a counting bridge."""

    def __init__(self) -> None:
        self.server = _R3aServer(host="127.0.0.1", port=0, plugin_key=_R3A_PLUGIN_KEY, invoke_timeout_seconds=3)

    def __enter__(self) -> "_R3aCompanion":
        self.server.start()
        self.session = _r3a_session(self.server, generated_fingerprint())
        self.session_obj = self.server.sessions.get(self.session["sessionId"])
        self.headers = _r3a_pane_headers(self.session)
        return self

    def __exit__(self, *_args: Any) -> None:
        self.server.stop()

    def bridge_patches(self, bridge_client: Any, calls: list[Any]):
        host, port = self.server.server_address()
        real_urlopen = bridge_client.urllib.request.urlopen

        def counting_urlopen(request: Any, timeout: Any = None):
            # urllib.request is process-global: count only plugin invokes,
            # not the pane driver's requests.
            if getattr(request, "full_url", "").endswith("/excel/invoke"):
                calls.append(request)
            return real_urlopen(request, timeout=timeout)

        return (
            patch.object(bridge_client, "_connection", return_value=(f"http://{host}:{port}", "k" * 32)),
            patch.object(bridge_client, "_validate_key", return_value=_R3A_PLUGIN_KEY),
            patch.object(bridge_client.urllib.request, "urlopen", counting_urlopen),
        )

    def serve_pending(self, result: Any) -> str:
        """Pane: claim the pending ticket, cross execute-start, report executed:true."""
        deadline = _r3a_time.monotonic() + 3
        while self.session_obj.get_pending_ticket() is None and _r3a_time.monotonic() < deadline:
            _r3a_time.sleep(0.005)
        status, pending = _r3a_request(self.server, "GET", "/excel/pending", headers=self.headers)
        assert status == 200, status
        ticket_id = pending["ticketId"]
        assert _r3a_request(self.server, "POST", "/excel/execute-start", {"ticketId": ticket_id}, self.headers)[0] == 200
        assert _r3a_request(self.server, "POST", "/excel/result",
                            {"ticketId": ticket_id, "executed": True, "result": result}, self.headers)[0] == 200
        return ticket_id

    def pending_status(self) -> int:
        return _r3a_request(self.server, "GET", "/excel/pending", headers=self.headers)[0]


def _r3a_in_thread(fn: Callable[[], Any]) -> tuple[Any, dict[str, Any]]:
    import threading as _threading
    box: dict[str, Any] = {}
    # A copied context, as Hermes' worker threads use, carries the pre-hook's
    # one-shot trusted plan holder into the handler thread.
    context = copy_context()

    def run() -> None:
        try:
            box["value"] = context.run(fn)
        except BaseException as exc:  # surfaced by the caller
            box["error"] = exc

    thread = _threading.Thread(target=run, daemon=True)
    thread.start()
    return thread, box


def _r3a_join(thread: Any, box: dict[str, Any]) -> Any:
    thread.join(5)
    assert not thread.is_alive(), "call did not finish"
    if "error" in box:
        raise box["error"]
    return box["value"]


def valid_execute_plan() -> dict:
    """R4g-2a rule M'': a schema-valid execute plan (the wire fixture's C02 plan) for the execute-path r3a tests."""
    return {
        "version": 1,
        "outcome": "execute",
        "summary": "Write one cell",
        "operations": [{"kind": "write_values", "sheet": "Sheet1", "address": "A1", "values": [["ok"]]}],
        "verification": [{"kind": "range_values_equal", "sheet": "Sheet1", "address": "A1", "expectedValues": [["ok"]]}],
    }


def test_r3a_ra8_plan_500_after_execute_start_is_uncertain_end_to_end() -> None:
    canary = {"outcome": "committed", "note": _R2D2G2_CANARY}
    boom = RuntimeError(_R2D2G2_CANARY)
    # (1) The real bridge client against a real companion whose plan outcome
    #     writer fails only after the pane executed: the dispatch catch-all 500.
    with _R3aCompanion() as companion, patch.object(_R3aHandler, "_send_plan_invoke_outcome", side_effect=boom):
        bridge_client = load_plugin().bridge_client
        calls: list[Any] = []
        op = str(uuid.uuid4())
        trusted = {"turn_operation_id": op, "request_id": op, "workbook_plan_fingerprint": generated_fingerprint(),
                   "phase": "execute", "turn_key": _R4G2_TURN_KEY}
        p1, p2, p3 = companion.bridge_patches(bridge_client, calls)
        with p1, p2, p3:
            def call() -> Any:
                try:
                    return bridge_client.invoke_bridge(companion.session["sessionId"], "execute_plan",
                                                       valid_execute_plan(), trusted=trusted)
                except bridge_client.BridgeError as exc:
                    return exc
            thread, box = _r3a_in_thread(call)
            companion.serve_pending(canary)
            outcome = _r3a_join(thread, box)
        assert companion.session_obj.get_ticket().terminal_reason == "completed"
        _r2d2g2_assert_uncertain(outcome, len(calls))

    # (2) The same through load_plugin() and the registered handler: the model
    #     sees ExecutionUncertain, never Internal.
    with _R3aCompanion() as companion, patch.object(_R3aHandler, "_send_plan_invoke_outcome", side_effect=boom):
        module, context = registered()
        hook = context.hooks["pre_tool_call"]
        handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
        calls = []
        p1, p2, p3 = companion.bridge_patches(module.bridge_client, calls)
        with p1, p2, p3:
            session_id = companion.session["sessionId"]
            assert hook(tool_name=PLAN_TOOL, args=valid_execute_plan(), session_id=session_id, turn_id="r3a-ra8",
                        tool_call_id="call-r3a-ra8", api_request_id="request-r3a-ra8", task_id="task-r3a") is None
            thread, box = _r3a_in_thread(lambda: handler(valid_execute_plan(), session_id=session_id))
            companion.serve_pending(canary)
            raw = _r3a_join(thread, box)
        assert json.loads(raw) == {
            "ok": False,
            "error": {"category": "ExecutionUncertain", "message": module.bridge_client.PLAN_UNCERTAIN_MESSAGE},
        }
        assert _R2D2G2_CANARY not in raw
        assert len(calls) == 1


def test_r3a_gp1_plugin_seen_set_is_advisory_and_companion_is_authoritative() -> None:
    with _R3aCompanion() as companion:
        module, context = registered()
        hook = context.hooks["pre_tool_call"]
        handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
        session_id = companion.session["sessionId"]
        responses: list[Any] = []
        real_invoke = module.bridge_client.invoke_bridge

        def recording_invoke(*args: Any, **kwargs: Any) -> Any:
            try:
                response = real_invoke(*args, **kwargs)
            except module.bridge_client.BridgeError as exc:
                responses.append(exc)
                raise
            responses.append(response)
            return response

        def pre_hook(plan: dict) -> Any:
            return hook(tool_name=PLAN_TOOL, args=plan, session_id=session_id, turn_id="r3a-gp1-turn",
                        tool_call_id="call-r3a-gp1", api_request_id="request-r3a-gp1", task_id="task-r3a")

        result = {"outcome": "inspected"}
        calls: list[Any] = []
        p1, p2, p3 = companion.bridge_patches(module.bridge_client, calls)
        with p1, p2, p3, patch.object(module, "invoke_bridge", recording_invoke):
            assert pre_hook(valid_execute_plan()) is None
            thread, box = _r3a_in_thread(lambda: handler(valid_execute_plan(), session_id=session_id))
            ticket_id = companion.serve_pending(result)
            assert json.loads(_r3a_join(thread, box)) == result
            first = responses[-1]
            assert first["ticketId"] == ticket_id and first["executed"] is True

            # The plugin forgets (session, turn): its seen-set is advisory.
            module._seen_plan_turns.clear()
            assert pre_hook(valid_execute_plan()) is None
            thread, box = _r3a_in_thread(lambda: handler(valid_execute_plan(), session_id=session_id))
            assert json.loads(_r3a_join(thread, box)) == result
            assert responses[-1] == first
            assert companion.pending_status() == 204
            assert list(companion.session_obj._request_history) == [first["requestId"]]
            assert companion.session_obj.get_ticket().ticket_id == ticket_id

            # A changed plan for the same (session, turn) is an atomic conflict.
            module._seen_plan_turns.clear()
            changed = dict(valid_execute_plan(), summary="A different plan for the same trusted turn")
            assert pre_hook(changed) is None
            thread, box = _r3a_in_thread(lambda: handler(changed, session_id=session_id))
            assert json.loads(_r3a_join(thread, box)) == {
                "ok": False, "error": {"category": "Conflict", "message": "requestId reused with different payload"},
            }
            assert companion.pending_status() == 204
            assert list(companion.session_obj._request_history) == [first["requestId"]]
            assert companion.session_obj.get_ticket().ticket_id == ticket_id
        assert len(calls) == 3


# ---------------------------------------------------------------------------
# R4g-2a (dev3-r4g2-release.md §5.1, F7, F8): trusted plan envelope v2 on the plugin side. P1-P6 and P4b
# drive the REAL registered pre-hook and plan handler (only the module-level invoke_bridge is recorded) and
# the REAL bridge_client. J2 vectors are literals computed once (finalizer, independent of this writer):
# NS 5f8f4f1c-...9248, session SESSION, turn "r2a-real-uuid5".
# ---------------------------------------------------------------------------
_R4G2_J2_TURN = "r2a-real-uuid5"
_R4G2_J2_TURN_KEY = "5a83b37f-b6e2-5031-8d37-df04bbf71b8a"
_R4G2_J2_OP_INSPECT = "cfed8ba1-049f-58dd-ac4d-dad4417a0ca2"
_R4G2_J2_OP_EXECUTE = "7244df56-2215-500f-98d1-2798e9313f70"
_R4G2_BLOCK = {"action": "block", "message": "Workbook plan trusted context rejected."}
_R4G2_INVALID_PLAN = {"version": 1, "outcome": "inspect", "reason": "no reads"}


def _r4g2_plugin():
    """A fresh registered module whose invoke_bridge records the private envelope (no network)."""
    module, context = registered()
    hook = context.hooks["pre_tool_call"]
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    calls: list[dict[str, Any]] = []

    def recording_invoke(session_id: Any, tool: Any, arguments: Any, trusted: Any = None, **_kwargs: Any) -> Any:
        calls.append({"session": session_id, "tool": tool, "arguments": json.loads(json.dumps(arguments)),
                      "trusted": None if trusted is None else dict(trusted), "keys": None if trusted is None else list(trusted)})
        return {"ok": True, "data": {"accepted": True}}

    def call(plan: Any, turn: str = _R4G2_J2_TURN, session: str = SESSION) -> Any:
        """Hermes' shape: pre-hook, then (if admitted) the handler, in one context."""
        blocked = hook(tool_name=PLAN_TOOL, args=plan, session_id=session, turn_id=turn,
                       tool_call_id=f"call-{len(calls)}", api_request_id="request-r4g2", task_id="task-r4g2")
        if blocked is not None:
            return blocked
        return json.loads(handler(plan, session_id=session))

    return module, hook, handler, calls, recording_invoke, call


def test_r4g2_p1_inspect_then_execute_same_turn_two_invokes_five_key_envelopes_and_j2_vectors() -> None:
    module, _hook, _handler, calls, recording_invoke, call = _r4g2_plugin()
    with patch.object(module, "invoke_bridge", recording_invoke):
        assert call(valid_plan()) == {"accepted": True}
        assert call(valid_execute_plan()) == {"accepted": True}
    assert len(calls) == 2
    keys = ["turn_operation_id", "request_id", "workbook_plan_fingerprint", "phase", "turn_key"]
    assert [c["keys"] for c in calls] == [keys, keys]
    inspect, execute = calls[0]["trusted"], calls[1]["trusted"]
    assert (inspect["phase"], execute["phase"]) == ("inspect", "execute")
    assert inspect["turn_key"] == execute["turn_key"] == _R4G2_J2_TURN_KEY
    assert inspect["turn_operation_id"] == inspect["request_id"] == _R4G2_J2_OP_INSPECT
    assert execute["turn_operation_id"] == execute["request_id"] == _R4G2_J2_OP_EXECUTE
    assert len({_R4G2_J2_TURN_KEY, _R4G2_J2_OP_INSPECT, _R4G2_J2_OP_EXECUTE}) == 3
    assert inspect["workbook_plan_fingerprint"] == execute["workbook_plan_fingerprint"] == generated_fingerprint()
    assert calls[0]["arguments"] == valid_plan() and calls[1]["arguments"] == valid_execute_plan()


def test_r4g2_p2_execute_then_inspect_same_turn_second_is_blocked() -> None:
    module, _hook, _handler, calls, recording_invoke, call = _r4g2_plugin()
    with patch.object(module, "invoke_bridge", recording_invoke):
        assert call(valid_execute_plan()) == {"accepted": True}
        assert call(valid_plan()) == _R4G2_BLOCK
        assert call(valid_execute_plan()) == _R4G2_BLOCK  # nothing after an execute
    assert len(calls) == 1


def test_r4g2_p3_inspect_then_inspect_same_turn_second_is_blocked_and_a_later_turn_is_fresh() -> None:
    module, _hook, _handler, calls, recording_invoke, call = _r4g2_plugin()
    with patch.object(module, "invoke_bridge", recording_invoke):
        assert call(valid_plan()) == {"accepted": True}
        assert call(valid_plan()) == _R4G2_BLOCK
        assert call(valid_plan(), turn="r4g2-later-turn") == {"accepted": True}
    assert len(calls) == 2
    assert calls[0]["trusted"]["turn_key"] != calls[1]["trusted"]["turn_key"]


def test_r4g2_p4_two_invalid_plans_exhaust_the_turn() -> None:
    module, _hook, _handler, calls, recording_invoke, call = _r4g2_plugin()
    with patch.object(module, "invoke_bridge", recording_invoke):
        assert call(_R4G2_INVALID_PLAN) == _R4G2_BLOCK  # the one correction opportunity
        assert call(_R4G2_INVALID_PLAN) == _R4G2_BLOCK  # the second invalid attempt exhausts the turn
        assert call(valid_plan()) == _R4G2_BLOCK
        assert call(valid_execute_plan()) == _R4G2_BLOCK
        assert call(valid_plan(), turn="r4g2-p4-later") == {"accepted": True}  # only that turn
    assert [c["trusted"]["phase"] for c in calls] == ["inspect"]


def test_r4g2_p4b_one_invalid_plan_then_a_valid_plan_is_admitted() -> None:
    module, _hook, _handler, calls, recording_invoke, call = _r4g2_plugin()
    with patch.object(module, "invoke_bridge", recording_invoke):
        assert call(_R4G2_INVALID_PLAN) == _R4G2_BLOCK
        assert call(valid_execute_plan()) == {"accepted": True}
    assert len(calls) == 1
    assert calls[0]["trusted"]["phase"] == "execute"
    assert calls[0]["trusted"]["request_id"] == _R4G2_J2_OP_EXECUTE


def test_r4g2_p5_handler_rederives_the_phase_and_refuses_a_disagreeing_context() -> None:
    # (a) handler args that differ from the hooked plan (other outcome): Validation, no invoke.
    module, hook, handler, calls, recording_invoke, _call = _r4g2_plugin()
    with patch.object(module, "invoke_bridge", recording_invoke):
        assert hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="r4g2-p5a",
                    tool_call_id="call-p5a", api_request_id="request-p5a", task_id="task-p5a") is None
        result = json.loads(handler(valid_execute_plan(), session_id=SESSION))
    assert result == {"ok": False, "error": {"category": "Validation", "message": "Trusted plan context does not match the plan"}}
    # (b) white box: the hooked context's phase disagrees with the plan's outcome (defence in depth).
    module, hook, handler, calls_b, recording_invoke, _call = _r4g2_plugin()
    with patch.object(module, "invoke_bridge", recording_invoke):
        assert hook(tool_name=PLAN_TOOL, args=valid_plan(), session_id=SESSION, turn_id="r4g2-p5b",
                    tool_call_id="call-p5b", api_request_id="request-p5b", task_id="task-p5b") is None
        context = module._trusted_plan_context.get().peek()
        assert context["phase"] == "inspect"
        context["phase"] = "execute"
        result = json.loads(handler(valid_plan(), session_id=SESSION))
    assert result == {"ok": False, "error": {"category": "Validation", "message": "Trusted plan context does not match the plan"}}
    assert calls == [] and calls_b == []


@pytest.mark.parametrize(
    ("label", "trusted"),
    [
        ("v1_three_keys", {"turn_operation_id": SESSION, "request_id": SESSION, "workbook_plan_fingerprint": None}),
        ("six_keys", {"turn_operation_id": SESSION, "request_id": SESSION, "workbook_plan_fingerprint": None,
                      "phase": "execute", "turn_key": _R4G2_J2_TURN_KEY, "extra": "x"}),
        ("phase_unknown", {"turn_operation_id": SESSION, "request_id": SESSION, "workbook_plan_fingerprint": None,
                           "phase": "commit", "turn_key": _R4G2_J2_TURN_KEY}),
        ("phase_uppercase", {"turn_operation_id": SESSION, "request_id": SESSION, "workbook_plan_fingerprint": None,
                             "phase": "EXECUTE", "turn_key": _R4G2_J2_TURN_KEY}),
        ("turn_key_equals_op", {"turn_operation_id": SESSION, "request_id": SESSION, "workbook_plan_fingerprint": None,
                                "phase": "execute", "turn_key": SESSION}),
        ("turn_key_not_uuid", {"turn_operation_id": SESSION, "request_id": SESSION, "workbook_plan_fingerprint": None,
                               "phase": "execute", "turn_key": "r4g2-not-a-uuid"}),
        ("turn_key_not_canonical", {"turn_operation_id": SESSION, "request_id": SESSION, "workbook_plan_fingerprint": None,
                                    "phase": "execute", "turn_key": _R4G2_J2_TURN_KEY.upper()}),
        ("turn_key_not_str", {"turn_operation_id": SESSION, "request_id": SESSION, "workbook_plan_fingerprint": None,
                              "phase": "execute", "turn_key": uuid.UUID(_R4G2_J2_TURN_KEY)}),
    ],
)
def test_r4g2_p6_bridge_client_refuses_every_malformed_v2_envelope_before_any_request(label: str, trusted: dict) -> None:
    bridge_client = load_plugin().bridge_client
    trusted = {key: (generated_fingerprint() if value is None else value) for key, value in trusted.items()}
    opened: list[Any] = []
    with (
        patch.object(bridge_client, "_connection", return_value=("http://127.0.0.1:8650", "k" * 32)),
        patch.object(bridge_client, "_validate_key", return_value="k" * 32),
        patch.object(bridge_client.urllib.request, "urlopen", side_effect=lambda *a, **k: opened.append(a)),
    ):
        with pytest.raises(bridge_client.BridgeError) as raised:
            bridge_client.invoke_bridge(SESSION, "execute_plan", valid_execute_plan(), trusted=trusted)
    assert raised.value.category == "Validation", label
    assert opened == [], label


def test_r4g2_p6b_bridge_client_sends_exactly_four_wire_keys_in_order() -> None:
    bridge_client = load_plugin().bridge_client
    calls: list[Any] = []

    def fake_urlopen(request: Any, timeout: int):
        calls.append(request)
        return _FakeBridgeResponse(b'{"ok":true,"data":{"accepted":true}}')

    trusted = {"turn_operation_id": _R4G2_J2_OP_EXECUTE, "request_id": _R4G2_J2_OP_EXECUTE,
               "workbook_plan_fingerprint": generated_fingerprint(), "phase": "execute", "turn_key": _R4G2_J2_TURN_KEY}
    with (
        patch.object(bridge_client, "_connection", return_value=("http://127.0.0.1:8650", "k" * 32)),
        patch.object(bridge_client, "_validate_key", return_value="k" * 32),
        patch.object(bridge_client.urllib.request, "urlopen", side_effect=fake_urlopen),
    ):
        bridge_client.invoke_bridge(SESSION, "execute_plan", valid_execute_plan(), trusted=trusted)
    body = json.loads(calls[0].data.decode("utf-8"))
    assert list(body["trusted"]) == ["turnOperationId", "workbookPlanFingerprint", "phase", "turnKey"]
    assert body["trusted"] == {"turnOperationId": _R4G2_J2_OP_EXECUTE, "workbookPlanFingerprint": generated_fingerprint(),
                               "phase": "execute", "turnKey": _R4G2_J2_TURN_KEY}
    assert body["requestId"] == _R4G2_J2_OP_EXECUTE


def test_r4g2_p7_the_turn_key_binds_the_session_so_another_pane_session_never_shares_a_turn_fence() -> None:
    module, _hook, _handler, calls, recording_invoke, call = _r4g2_plugin()
    with patch.object(module, "invoke_bridge", recording_invoke):
        assert call(valid_execute_plan()) == {"accepted": True}
        assert call(valid_execute_plan(), session=OTHER_SESSION) == {"accepted": True}
    assert calls[0]["trusted"]["turn_key"] == _R4G2_J2_TURN_KEY
    assert calls[1]["trusted"]["turn_key"] not in (_R4G2_J2_TURN_KEY, _R4G2_J2_OP_EXECUTE, _R4G2_J2_OP_INSPECT)
    assert calls[1]["trusted"]["request_id"] != _R4G2_J2_OP_EXECUTE


# ---------------------------------------------------------------------------
# R4g-2b (OMH release record 2026-10-07, scope (2); R4g-2a lane-B A3): the plugin re-verifies every
# executed:false body before the model sees it. A refusal (a JSON object with ok === false) passes unchanged; any
# other body must be a conforming T8 inspection body (exact keys; reads = plain objects with exactly {ok, data}, ok
# boolean; <= 1 MiB per read and <= 1.5 MiB per body as compact UTF-8 JSON) answering an inspect-phase call.
# Anything else is replaced by the fixed not-executed body (the bridge's own T1 `rejected` result).
# ---------------------------------------------------------------------------
_R4G2B_READ_CAP = 1_048_576
_R4G2B_BODY_CAP = 1_572_864
_R4G2B_FIXED = {"ok": False, "action": "workbook_plan", "category": "Execution", "reason": "rejected",
                "error": "Workbook plan was not executed."}


def _r4g2b_t8(reads: list) -> dict:
    return {"ok": True, "action": "workbook_plan", "phase": "inspect",
            "result": {"outcome": "inspected", "inspectionCount": 1, "reads": reads}}


def _r4g2b_bytes(value: Any) -> int:
    return len(json.dumps(value, separators=(",", ":"), ensure_ascii=False).encode("utf-8"))


def _r4g2b_model_view(data: Any, plan: dict | None = None) -> Any:
    module, context = registered()
    plan = valid_plan() if plan is None else plan
    hook = context.hooks["pre_tool_call"]
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)
    seen: list[Any] = []

    def bridge(*_a: Any, **kw: Any) -> Any:
        seen.append(kw.get("trusted", {}).get("phase"))
        return {"ok": True, "ticketId": SESSION, "requestId": "r", "data": copy.deepcopy(data), "executed": False}

    with patch.object(module, "invoke_bridge", bridge):
        hook(tool_name=PLAN_TOOL, args=plan, session_id=SESSION, turn_id="r4g2b-turn", tool_call_id="call-r4g2b",
             api_request_id="request-r4g2b", task_id="task-r4g2b")
        text = handler(plan, session_id=SESSION)
    assert seen == [plan["outcome"] if plan["outcome"] == "inspect" else "execute"]
    return json.loads(text)


def _r4g2b_read_of_bytes(size: int) -> dict:
    read = {"ok": True, "data": "x" * (size - _r4g2b_bytes({"ok": True, "data": ""}))}
    assert _r4g2b_bytes(read) == size
    return read


def test_r4g2b_p0_control_conforming_inspection_body_and_refusals_pass_verbatim() -> None:
    body = _r4g2b_t8([{"ok": True, "data": [["v"]]}, {"data": None, "ok": False}])
    assert _r4g2b_model_view(body) == {"executed": False, "result": body}
    t1 = {"ok": False, "action": "workbook_plan", "reason": "invalid", "error": "Workbook plan ticket was invalid and was not executed."}
    t2 = {"ok": False, "action": "workbook_plan", "reason": "rejected",
          "summary": "The workbook plan was rejected before execution; nothing was executed."}
    for refusal in (t1, t2):
        assert _r4g2b_model_view(refusal) == {"executed": False, "result": refusal}
        assert _r4g2b_model_view(refusal, valid_execute_plan()) == {"executed": False, "result": refusal}


@pytest.mark.parametrize("label,data", [
    ("string read (canary)", _r4g2b_t8(["canary-string-read"])),
    ("number and null reads", _r4g2b_t8([7, None])),
    ("array read", _r4g2b_t8([[{"ok": True, "data": 1}]])),
    ("read without ok", _r4g2b_t8([{"data": [["v"]]}])),
    ("read whose ok is a string", _r4g2b_t8([{"ok": "true", "data": 1}])),
    ("read whose ok is 1", _r4g2b_t8([{"ok": 1, "data": 1}])),
    ("read without data", _r4g2b_t8([{"ok": True}])),
    ("read with extra formulas", _r4g2b_t8([{"ok": True, "data": 1, "formulas": [["=SECRET()"]]}])),
    ("error-shaped read", _r4g2b_t8([{"ok": False, "error": "x"}])),
    ("raw read_range result", _r4g2b_t8([{"sheet": "Sheet1", "address": "A1", "values": [["v"]]}])),
    ("reads not a list", {**_r4g2b_t8([]), "result": {"outcome": "inspected", "inspectionCount": 1, "reads": {"a": 1}}}),
    ("inspectionCount True", {**_r4g2b_t8([]), "result": {"outcome": "inspected", "inspectionCount": True, "reads": []}}),
    ("inspectionCount 2", {**_r4g2b_t8([]), "result": {"outcome": "inspected", "inspectionCount": 2, "reads": []}}),
    ("outcome committed", {**_r4g2b_t8([]), "result": {"outcome": "committed", "inspectionCount": 1, "reads": []}}),
    ("extra result key", {**_r4g2b_t8([]), "result": {"outcome": "inspected", "inspectionCount": 1, "reads": [], "x": 1}}),
    ("T8 plus reason", {**_r4g2b_t8([]), "reason": "rejected"}),
    ("T8 without phase", {"ok": True, "action": "workbook_plan", "result": _r4g2b_t8([])["result"]}),
    ("phase execute", {**_r4g2b_t8([]), "phase": "execute"}),
    ("ok-less report", {"executed": False, "reason": "declined"}),
    ("ok truthy string", {"ok": "yes", "data": "x" * 64}),
    ("non-object body", "canary-string-body"),
    ("list body", [1, 2]),
])
def test_r4g2b_p1_nonconforming_bodies_are_replaced_by_the_fixed_not_executed_body(label: str, data: Any) -> None:
    view = _r4g2b_model_view(data)
    assert view == {"executed": False, "result": _R4G2B_FIXED}, label
    assert "canary" not in json.dumps(view) and "SECRET" not in json.dumps(view)


def test_r4g2b_p2_an_inspection_body_answering_an_execute_call_is_refused() -> None:
    body = _r4g2b_t8([{"ok": True, "data": [["v"]]}])
    assert _r4g2b_model_view(body, valid_execute_plan()) == {"executed": False, "result": _R4G2B_FIXED}


def test_r4g2b_p3_per_read_cap_is_utf8_bytes_of_compact_json() -> None:
    at_cap = _r4g2b_read_of_bytes(_R4G2B_READ_CAP)
    assert _r4g2b_model_view(_r4g2b_t8([at_cap])) == {"executed": False, "result": _r4g2b_t8([at_cap])}
    over = _r4g2b_read_of_bytes(_R4G2B_READ_CAP + 1)
    assert _r4g2b_model_view(_r4g2b_t8([over])) == {"executed": False, "result": _R4G2B_FIXED}
    base = _r4g2b_bytes({"ok": True, "data": ""})
    wide = {"ok": True, "data": "\u00e4" * ((_R4G2B_READ_CAP - base) // 2 + 1)}
    assert len(json.dumps(wide, separators=(",", ":"), ensure_ascii=False)) < _R4G2B_READ_CAP < _r4g2b_bytes(wide)
    assert _r4g2b_model_view(_r4g2b_t8([wide])) == {"executed": False, "result": _R4G2B_FIXED}


def test_r4g2b_p4_body_cap_is_utf8_bytes_of_the_compact_inspection_body() -> None:
    first = _r4g2b_read_of_bytes(_R4G2B_READ_CAP)
    overhead = _r4g2b_bytes(_r4g2b_t8([first, {"ok": True, "data": ""}]))
    second = {"ok": True, "data": "x" * (_R4G2B_BODY_CAP - overhead)}
    assert _r4g2b_bytes(second) <= _R4G2B_READ_CAP and _r4g2b_bytes(_r4g2b_t8([first, second])) == _R4G2B_BODY_CAP
    assert _r4g2b_model_view(_r4g2b_t8([first, second])) == {"executed": False, "result": _r4g2b_t8([first, second])}
    over = {"ok": True, "data": "x" * (_R4G2B_BODY_CAP - overhead + 1)}
    assert _r4g2b_bytes(_r4g2b_t8([first, over])) == _R4G2B_BODY_CAP + 1
    assert _r4g2b_model_view(_r4g2b_t8([first, over])) == {"executed": False, "result": _R4G2B_FIXED}


# ---------------------------------------------------------------------------
# R4g-2c (dev3-r4g2c-release.md items 1-4 and 6; R4g-2b lane-B A1-A4/A6, lane-A F1/F2/F4): pre-flip hardening of the
# plugin's executed:false re-verification.
# - item 1: an ok:false body passes only in one of the closed R4h shapes: T1 bridge invalid {ok, action, reason, error}
#   and rejected {ok, action, category, reason, error}; T2 route {ok, action, reason, summary} for rejected / cancelled /
#   not_started / invalid, + result {outcome, finalSummary} iff ask/refuse (outcome == reason), + result {execution}
#   iff already_executed. Display strings are strings of <= 2048 UTF-16 units and the body is <= 64 KiB (provisional).
#   Everything else is the fixed not-executed body (no new strings).
#   R4g-2c-r1 (lane A F1 / lane B B1, B2): display strings count code points (<= 2048, as R4h/R4j and the route's
#   boundedCodePointLength); result.execution must be the executor's closed report shape; result-bearing refusals
#   (ask/refuse/already_executed) pass only in the execute phase. Rule R1: the rows below that pinned the R4g-2c
#   behaviour are flipped to the stricter r1 behaviour (q0 phase split, q0b code-point controls, the q1 astral row at
#   2049 code points, the fixture-like already_executed body now a real report).
# - items 2/3: sizes are the UTF-8 bytes of the pane's JSON.stringify text: lone surrogates count as their 6-byte
#   \\uXXXX escape and numbers as JS writes them (1e-7, 0.000001); NaN/Infinity (never produced by JS) fail closed. A
#   measuring problem is never a tool error.
# - item 4: a T8 body with zero reads is refused. item 6 (mirror): a read nesting deeper than 64 levels is refused.
# JS texts below are written out by hand (the JSON.stringify form) so the expected byte counts never come from Python's
# own serializer.
# ---------------------------------------------------------------------------
_R4G2C_REFUSAL_CAP = 65_536
_R4G2C_TEXT_UNITS = 2048
_R4G2C_DEPTH_CAP = 64
_R4G2C_SUMMARY = {
    "ask": "The workbook plan asked for clarification; nothing was executed.",
    "refuse": "The workbook plan was refused; nothing was executed.",
    "rejected": "The workbook plan was rejected before execution; nothing was executed.",
    "cancelled": "The workbook plan was cancelled before execution; nothing was executed.",
    "not_started": "Workbook plan execution did not start; nothing was executed.",
    "invalid": "The workbook plan ticket was invalid; nothing was executed.",
    "already_executed": "This workbook operation was already applied earlier and was not executed again.",
}
_R4G2C_T1_INVALID = {"ok": False, "action": "workbook_plan", "reason": "invalid",
                     "error": "Workbook plan ticket was invalid and was not executed."}


def _r4g2c_t2(reason: str, **extra: Any) -> dict:
    return {"ok": False, "action": "workbook_plan", "reason": reason, "summary": _R4G2C_SUMMARY[reason], **extra}


def _r4g2c_closed_refusals() -> list[tuple[str, dict]]:
    out = [("T1 invalid", dict(_R4G2C_T1_INVALID)), ("T1 rejected", dict(_R4G2B_FIXED))]
    for reason in ("rejected", "cancelled", "not_started", "invalid"):
        out.append((f"T2 {reason}", _r4g2c_t2(reason)))
    for reason in ("ask", "refuse"):
        out.append((f"T2 {reason}", _r4g2c_t2(reason, result={"outcome": reason, "finalSummary": "Which sheet?"})))
    out.append(("T2 already_executed", _r4g2c_t2("already_executed", result={"execution": _r4g2c_r1_report()})))
    return out


# R4g-2c-r1: an execution report exactly as workbookPlanExecutor.ts produces it (WorkbookPlanExecutionReport: fixed
# keys, enum values, lowercase UUIDs, safe non-negative ints, closed 6-key operation records); the shape of lane B's
# 52 recorded real-route replays (committed 1 op / partial with firstFailed*).
_R4G2C_R1_TX = "0f1e2d3c-4b5a-5968-8776-a5b4c3d2e1f0"
_R4G2C_R1_OP = "1a2b3c4d-5e6f-5a7b-9c8d-0e1f2a3b4c5d"


def _r4g2c_r1_op(index: int = 0, **over: Any) -> dict:
    op = {"operationId": str(uuid.UUID(int=uuid.UUID(_R4G2C_R1_OP).int + index)), "operationIndex": index,
          "effect": "write", "groupIndex": index, "state": "committed", "status": "applied"}
    return {**op, **over}


def _r4g2c_r1_report(count: int = 1, **over: Any) -> dict:
    report = {"outcome": "committed", "transactionId": _R4G2C_R1_TX, "operationCount": count,
              "operations": [_r4g2c_r1_op(i) for i in range(count)], "verification": "not_attempted",
              "rollback": "not_attempted"}
    return {**report, **over}


def _r4g2c_view_text(data: Any, plan: dict | None = None) -> str:
    """The raw handler text (never an exception) for a companion envelope carrying `data` as executed:false."""
    module, context = registered()
    plan = valid_plan() if plan is None else plan
    hook = context.hooks["pre_tool_call"]
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == PLAN_TOOL)

    def bridge(*_a: Any, **_kw: Any) -> Any:
        return {"ok": True, "ticketId": SESSION, "requestId": "r", "data": copy.deepcopy(data), "executed": False}

    with patch.object(module, "invoke_bridge", bridge):
        hook(tool_name=PLAN_TOOL, args=plan, session_id=SESSION, turn_id="r4g2c-turn", tool_call_id="call-r4g2c",
             api_request_id="request-r4g2c", task_id="task-r4g2c")
        return handler(plan, session_id=SESSION)


def _r4g2c_js_read(items_js: list[str], size: int) -> tuple[dict, str]:
    """A read whose JSON.stringify text is exactly `size` UTF-8 bytes: {"ok":true,"data":[<items>,"<x padding>"]}."""
    head = '{"ok":true,"data":[' + "".join(item + "," for item in items_js) + '"'
    tail = '"]}'
    text = head + "x" * (size - len(head.encode("utf-8")) - len(tail)) + tail
    assert len(text.encode("utf-8")) == size
    return json.loads(text), text


def test_r4g2c_q0_control_every_closed_refusal_shape_passes_verbatim_in_both_phases() -> None:
    # R4g-2c-r1 rule R1: every closed shape passes in the execute phase; in the inspect phase only the bodies without
    # a result (T1 invalid/rejected, T2 rejected/cancelled/not_started/invalid) pass (see r4g2c_r1_p rows).
    for label, body in _r4g2c_closed_refusals():
        assert _r4g2b_model_view(body, valid_execute_plan()) == {"executed": False, "result": body}, label
        expected = body if "result" not in body else _R4G2B_FIXED
        assert _r4g2b_model_view(body, valid_plan()) == {"executed": False, "result": expected}, label
    # a lone surrogate inside a display string is measured as JS escapes it, never a tool error
    lone = _r4g2c_t2("ask", result={"outcome": "ask", "finalSummary": "Which sheet \ud800?"})
    assert _r4g2b_model_view(lone, valid_execute_plan()) == {"executed": False, "result": lone}


def test_r4g2c_q0b_control_refusal_boundaries_are_inclusive() -> None:
    # R4g-2c-r1 rule R1: 2048 code points (4096 UTF-16 units of U+1F600) is inside the cap; the byte-cap boundary row
    # is test_r4g2c_r1_cap_refusal_byte_cap_is_inclusive (a padded `execution` is no longer a report).
    at_cap = _r4g2c_t2("ask", result={"outcome": "ask", "finalSummary": "\U0001F600" * _R4G2C_TEXT_UNITS})
    assert _r4g2b_model_view(at_cap, valid_execute_plan()) == {"executed": False, "result": at_cap}


def _r4g2c_nest(levels: int) -> Any:
    value: Any = []
    for _ in range(levels - 1):
        value = [value]
    return value


def _r4g2c_open_refusals() -> list[tuple[str, Any]]:
    t8 = _r4g2b_t8([{"ok": True, "data": [["v"]]}])
    big = "x" * (3 * 1024 * 1024)
    ask = _r4g2c_t2("ask", result={"outcome": "ask", "finalSummary": "Which sheet?"})
    executed = _r4g2c_closed_refusals()[-1][1]
    return [
        ("T8-shaped body flipped to ok false, formula canary read",
         {**_r4g2b_t8([{"ok": True, "data": 1, "formulas": [["=WEBSERVICE(\"canary\")"]]}]), "ok": False}),
        ("raw read_range SECRET under ok false", {**t8, "ok": False, "result": {"values": [["SECRET"]]}}),
        ("T2 rejected plus a leak key", {**_r4g2c_t2("rejected"), "leak": "canary"}),
        ("3 MiB ok-false string", {"ok": False, "error": big}),
        ("ok false and error only", {"ok": False, "error": "canary"}),
        ("T1 invalid without action", {k: v for k, v in _R4G2C_T1_INVALID.items() if k != "action"}),
        ("T1 invalid with reason ask", {**_R4G2C_T1_INVALID, "reason": "ask"}),
        ("T1 rejected with reason invalid", {**_R4G2B_FIXED, "reason": "invalid"}),
        ("T1 error not a string", {**_R4G2C_T1_INVALID, "error": ["canary"]}),
        ("T1 category not a string", {**_R4G2B_FIXED, "category": {"canary": 1}}),
        ("T2 without summary", {k: v for k, v in _r4g2c_t2("cancelled").items() if k != "summary"}),
        ("T2 unknown reason", {**_r4g2c_t2("cancelled"), "reason": "declined"}),
        ("T2 summary not a string", {**_r4g2c_t2("not_started"), "summary": 7}),
        ("T2 summary over 2048 code points", {**_r4g2c_t2("not_started"), "summary": "s" * (_R4G2C_TEXT_UNITS + 1)}),
        ("T2 ask without result", _r4g2c_t2("ask")),
        ("T2 rejected with a result", _r4g2c_t2("rejected", result={"outcome": "rejected", "finalSummary": "x"})),
        ("T2 ask with outcome refuse", {**ask, "result": {"outcome": "refuse", "finalSummary": "x"}}),
        ("T2 ask with an extra result key", {**ask, "result": {"outcome": "ask", "finalSummary": "x", "reads": [1]}}),
        ("T2 ask finalSummary over 2048 code points (astral)",
         {**ask, "result": {"outcome": "ask", "finalSummary": "\U0001F600" * (_R4G2C_TEXT_UNITS + 1)}}),
        ("T2 ask finalSummary not a string", {**ask, "result": {"outcome": "ask", "finalSummary": None}}),
        ("T2 already_executed with execution null", {**executed, "result": {"execution": None}}),
        ("T2 already_executed with execution a list", {**executed, "result": {"execution": [1]}}),
        ("T2 already_executed over the refusal cap", {**executed, "result": {"execution": {"pad": "x" * _R4G2C_REFUSAL_CAP}}}),
        ("T2 already_executed nesting 70 levels", {**executed, "result": {"execution": {"n": _r4g2c_nest(68)}}}),
        ("ok false with a NaN", {**executed, "result": {"execution": {"n": float("nan")}}}),
    ]


@pytest.mark.parametrize("label,data", _r4g2c_open_refusals())
def test_r4g2c_q1_ok_false_bodies_outside_the_closed_shapes_are_replaced_by_the_fixed_body(label: str, data: Any) -> None:
    for plan in (valid_plan(), valid_execute_plan()):
        text = _r4g2c_view_text(data, plan)
        assert json.loads(text) == {"executed": False, "result": _R4G2B_FIXED}, label
        assert "canary" not in text and "SECRET" not in text and len(text) < 1024, label


def test_r4g2c_q1b_refusal_cap_plus_one_byte_is_refused() -> None:
    head = '{"ok":false,"action":"workbook_plan","reason":"already_executed","summary":"' + \
        _R4G2C_SUMMARY["already_executed"] + '","result":{"execution":{"pad":"'
    tail = '"}}}'
    text = head + "x" * (_R4G2C_REFUSAL_CAP + 1 - len(head) - len(tail)) + tail
    assert len(text.encode("utf-8")) == _R4G2C_REFUSAL_CAP + 1
    assert _r4g2b_model_view(json.loads(text), valid_execute_plan()) == {"executed": False, "result": _R4G2B_FIXED}


def test_r4g2c_q1c_the_old_rule_r_report_without_action_is_now_refused() -> None:
    assert _r4g2b_model_view({"ok": False, "executed": False, "reason": "declined"}) == \
        {"executed": False, "result": _R4G2B_FIXED}


def test_r4g2c_q2_lone_surrogate_reads_are_measured_as_js_escapes_never_a_tool_error() -> None:
    body = _r4g2b_t8([{"ok": True, "data": [["a\ud800b", "\udc00"]]}])
    text = _r4g2c_view_text(body)
    assert json.loads(text) == {"executed": False, "result": body}
    # exactly the per-read cap in JS bytes (each lone surrogate is a 6-byte \\uXXXX escape) passes; one more byte fails
    at_cap, js = _r4g2c_js_read(['"\\ud800"'] * 1000, _R4G2B_READ_CAP)
    assert js.count("\\ud800") == 1000
    assert _r4g2b_model_view(_r4g2b_t8([at_cap])) == {"executed": False, "result": _r4g2b_t8([at_cap])}
    over, _ = _r4g2c_js_read(['"\\ud800"'] * 1000, _R4G2B_READ_CAP + 1)
    assert _r4g2b_model_view(_r4g2b_t8([over])) == {"executed": False, "result": _R4G2B_FIXED}
    # the same body answering an execute call is the fixed body, not an exception
    assert json.loads(_r4g2c_view_text(body, valid_execute_plan())) == {"executed": False, "result": _R4G2B_FIXED}


def test_r4g2c_q3_numbers_are_measured_as_javascript_writes_them() -> None:
    # 1e-7: JS "1e-7" (Python "1e-07" is one byte longer). Exactly the cap in JS bytes passes.
    dense, js = _r4g2c_js_read(["1e-7"] * 100_000, _R4G2B_READ_CAP)
    assert len(json.dumps(dense, separators=(",", ":"), ensure_ascii=False)) > _R4G2B_READ_CAP
    assert _r4g2b_model_view(_r4g2b_t8([dense])) == {"executed": False, "result": _r4g2b_t8([dense])}
    # 0.000001: JS "0.000001" (Python "1e-06" is three bytes shorter). One byte over the cap in JS bytes fails.
    small, js = _r4g2c_js_read(["0.000001"] * 50_000, _R4G2B_READ_CAP + 1)
    assert len(json.dumps(small, separators=(",", ":"), ensure_ascii=False)) <= _R4G2B_READ_CAP
    assert _r4g2b_model_view(_r4g2b_t8([small])) == {"executed": False, "result": _R4G2B_FIXED}
    # the body cap uses the same measure: two JS-exact reads summing to exactly the body cap pass
    first, _ = _r4g2c_js_read(["1e-7"] * 100_000, _R4G2B_READ_CAP)
    overhead = len(json.dumps(_r4g2b_t8([first, {"ok": True, "data": [""]}]), separators=(",", ":"))) + \
        100_000 * (len("1e-7") - len("1e-07"))
    second, _ = _r4g2c_js_read([], _R4G2B_BODY_CAP - overhead + len('{"ok":true,"data":[""]}'))
    assert _r4g2b_model_view(_r4g2b_t8([first, second])) == {"executed": False, "result": _r4g2b_t8([first, second])}
    third, _ = _r4g2c_js_read([], _R4G2B_BODY_CAP - overhead + len('{"ok":true,"data":[""]}') + 1)
    assert _r4g2b_model_view(_r4g2b_t8([first, third])) == {"executed": False, "result": _R4G2B_FIXED}


@pytest.mark.parametrize("label,value", [("NaN", float("nan")), ("Infinity", float("inf")), ("-Infinity", float("-inf"))])
def test_r4g2c_q3b_non_finite_numbers_are_never_measured_or_shown(label: str, value: float) -> None:
    text = _r4g2c_view_text(_r4g2b_t8([{"ok": True, "data": [[value]]}]))
    assert json.loads(text) == {"executed": False, "result": _R4G2B_FIXED}, label
    assert "NaN" not in text and "Infinity" not in text


def test_r4g2c_q4_an_inspection_body_with_zero_reads_is_refused() -> None:
    assert _r4g2b_model_view(_r4g2b_t8([])) == {"executed": False, "result": _R4G2B_FIXED}
    one = _r4g2b_t8([{"ok": False, "data": None}])
    assert _r4g2b_model_view(one) == {"executed": False, "result": one}


def test_r4g2c_q4b_the_r4g2b_rows_reasons_still_refuse_with_a_read_present() -> None:
    read = [{"ok": True, "data": 1}]
    for label, result in [
        ("inspectionCount True", {"outcome": "inspected", "inspectionCount": True, "reads": read}),
        ("inspectionCount 2", {"outcome": "inspected", "inspectionCount": 2, "reads": read}),
        ("outcome committed", {"outcome": "committed", "inspectionCount": 1, "reads": read}),
        ("extra result key", {"outcome": "inspected", "inspectionCount": 1, "reads": read, "x": 1}),
    ]:
        assert _r4g2b_model_view({**_r4g2b_t8(read), "result": result}) == {"executed": False, "result": _R4G2B_FIXED}, label
    for label, body in [("T8 plus reason", {**_r4g2b_t8(read), "reason": "rejected"}),
                        ("phase execute", {**_r4g2b_t8(read), "phase": "execute"})]:
        assert _r4g2b_model_view(body) == {"executed": False, "result": _R4G2B_FIXED}, label


def test_r4g2c_q5_reads_nest_at_most_64_levels() -> None:
    at_cap = {"ok": True, "data": _r4g2c_nest(_R4G2C_DEPTH_CAP - 1)}
    assert _r4g2b_model_view(_r4g2b_t8([at_cap])) == {"executed": False, "result": _r4g2b_t8([at_cap])}
    deeper = {"ok": True, "data": _r4g2c_nest(_R4G2C_DEPTH_CAP)}
    assert _r4g2b_model_view(_r4g2b_t8([deeper])) == {"executed": False, "result": _R4G2B_FIXED}
    # far beyond the cap: still the fixed body, never a RecursionError tool failure
    very = {"ok": True, "data": _r4g2c_nest(300)}
    assert json.loads(_r4g2c_view_text(_r4g2b_t8([very]))) == {"executed": False, "result": _R4G2B_FIXED}


# ---------------------------------------------------------------------------
# R4g-2c-r1 (OMH repair release on 8fdb5ad: lane A F1-F3, lane B B1/B2).
# - item 1 (F1/B2): display strings are bounded in code points (Python len, as the route's boundedCodePointLength and
#   the R4h/R4j wire legs), cap 2048: 2048 x U+1F600 / CJK / U+20000 pass verbatim, 2049 do not.
# - item 2 (B1): already_executed carries result.execution only in the executor's WorkbookPlanExecutionReport shape
#   (git show 8fdb5ad:src/workbookPlanExecutor.ts isReport/isOperationReport/validStateStatus/validReportInvariants,
#   one primitive per plan operation so 1..30 operations, and the route replays only a report with some operation not
#   "not_committed"); anything else is the fixed body. Result-bearing refusals (ask/refuse/already_executed) pass only
#   in the execute phase; the inspect phase accepts the T1 bridge bodies and the summary-only T2 reasons.
# - item 3 (F2/F3): the T1 invalid key set is exact; _js_number_text at exactly 21 integer digits.
# ---------------------------------------------------------------------------
_R4G2C_R1_CANARY_T8 = {"outcome": "inspected", "inspectionCount": 1,
                       "reads": [{"ok": True, "data": [["SECRET"]], "formulas": [["=WEBSERVICE(\"http://evil\")"]],
                                  "sheet": "Secret"}]}


def _r4g2c_r1_ae(execution: Any) -> dict:
    return _r4g2c_t2("already_executed", result={"execution": execution})


def _r4g2c_r1_text_bodies() -> list[tuple[str, dict]]:
    out = []
    for label, text in [("2048 x U+1F600", "\U0001F600" * 2048), ("2048 CJK", "\u6f22" * 2048),
                        ("1500 x U+20000 + 548 ascii", "\U00020000" * 1500 + "a" * 548),
                        ("2048 lone surrogates", "\ud800" * 2048)]:
        assert len(text) == _R4G2C_TEXT_UNITS
        for reason in ("ask", "refuse"):
            out.append((f"{reason} {label}", _r4g2c_t2(reason, result={"outcome": reason, "finalSummary": text})))
    return out


@pytest.mark.parametrize("label,body", _r4g2c_r1_text_bodies())
def test_r4g2c_r1_t_display_text_is_bounded_in_code_points(label: str, body: dict) -> None:
    assert _r4g2b_model_view(body, valid_execute_plan()) == {"executed": False, "result": body}, label
    over = copy.deepcopy(body)
    over["result"]["finalSummary"] += "\U0001F600"
    assert _r4g2b_model_view(over, valid_execute_plan()) == {"executed": False, "result": _R4G2B_FIXED}, label


def test_r4g2c_r1_t2_every_display_string_uses_the_same_code_point_cap() -> None:
    for body in (_r4g2c_t2("rejected"), dict(_R4G2C_T1_INVALID), dict(_R4G2B_FIXED)):
        key = "summary" if "summary" in body else "error"
        at_cap = {**body, key: "\U0001F600" * _R4G2C_TEXT_UNITS}
        over = {**body, key: "\U0001F600" * (_R4G2C_TEXT_UNITS + 1)}
        for plan in (valid_plan(), valid_execute_plan()):
            assert _r4g2b_model_view(at_cap, plan) == {"executed": False, "result": at_cap}
            assert _r4g2b_model_view(over, plan) == {"executed": False, "result": _R4G2B_FIXED}
    at_cap = {**_R4G2B_FIXED, "category": "\u6f22" * _R4G2C_TEXT_UNITS}
    assert _r4g2b_model_view(at_cap) == {"executed": False, "result": at_cap}
    over = {**_R4G2B_FIXED, "category": "\u6f22" * (_R4G2C_TEXT_UNITS + 1)}
    assert _r4g2b_model_view(over) == {"executed": False, "result": _R4G2B_FIXED}


def _r4g2c_r1_real_reports() -> list[tuple[str, dict]]:
    partial = _r4g2c_r1_report(3, outcome="partial", verification="reconciled", rollback="unavailable",
                               firstFailedGroup=1, firstFailedOperationIndex=1,
                               operations=[_r4g2c_r1_op(0), _r4g2c_r1_op(1, state="unknown_within_batch",
                                                                         status="unknown_within_batch"),
                                           _r4g2c_r1_op(2, state="not_committed", status="failed")])
    return [
        ("committed, 1 operation", _r4g2c_r1_report()),
        ("committed, 30 operations (the plan schema maximum)", _r4g2c_r1_report(30)),
        ("partial with firstFailed*", partial),
        ("persistence_unavailable_after_mutation, verified", _r4g2c_r1_report(
            2, outcome="persistence_unavailable_after_mutation",
            operations=[_r4g2c_r1_op(0, status="verified"), _r4g2c_r1_op(1, status="rolled_back")])),
        ("rollback succeeded, every enum member used",
         _r4g2c_r1_report(5, outcome="partial", verification="reconciliation_failed", rollback="succeeded",
                          operations=[_r4g2c_r1_op(i, effect=e) for i, e in
                                      enumerate(("write", "format", "structure", "filter", "calculate"))])),
        ("groupIndex 0 for every operation", _r4g2c_r1_report(
            2, operations=[_r4g2c_r1_op(0, groupIndex=0), _r4g2c_r1_op(1, groupIndex=0)])),
    ]


@pytest.mark.parametrize("label,execution", _r4g2c_r1_real_reports())
def test_r4g2c_r1_r_executor_reports_pass_verbatim_in_the_execute_phase_only(label: str, execution: dict) -> None:
    body = _r4g2c_r1_ae(execution)
    assert _r4g2b_model_view(body, valid_execute_plan()) == {"executed": False, "result": body}, label
    assert _r4g2b_model_view(body, valid_plan()) == {"executed": False, "result": _R4G2B_FIXED}, label


def _r4g2c_r1_bad_reports() -> list[tuple[str, Any]]:
    one = _r4g2c_r1_report()
    op = one["operations"][0]

    def with_op(**over: Any) -> dict:
        return {**one, "operations": [{**op, **over}]}

    def without(key: str) -> dict:
        return {k: v for k, v in one.items() if k != key}

    return [
        ("inspection-shaped canary (lane B)", _R4G2C_R1_CANARY_T8),
        ("60 KiB arbitrary text (lane B)", {"leak": "SECRET" * 10000}),
        ("empty object", {}),
        ("synthetic marker", {"synthetic": True}),
        ("the R4g-2c fixture literal (transactionId t, verification passed)",
         {"outcome": "committed", "transactionId": "t", "operationCount": 1, "operations": [],
          "verification": "passed", "rollback": "not_needed"}),
        ("lane B attack report (verification passed, rollback not_needed)",
         {**_r4g2c_r1_report(2), "verification": "passed", "rollback": "not_needed"}),
        ("extra report key carrying a canary", {**one, "leak": "SECRET"}),
        ("extra operation key carrying a canary", with_op(leak="SECRET")),
        ("missing outcome", without("outcome")),
        ("missing transactionId", without("transactionId")),
        ("missing operationCount", without("operationCount")),
        ("missing operations", without("operations")),
        ("missing verification", without("verification")),
        ("missing rollback", without("rollback")),
        ("missing operation status", {**one, "operations": [{k: v for k, v in op.items() if k != "status"}]}),
        ("outcome outside the enum", {**one, "outcome": "SECRET"}),
        ("outcome not a string", {**one, "outcome": ["committed"]}),
        ("verification outside the enum", {**one, "verification": "passed"}),
        ("rollback outside the enum", {**one, "rollback": "not_needed"}),
        ("transactionId uppercase", {**one, "transactionId": _R4G2C_R1_TX.upper()}),
        ("transactionId version 0", {**one, "transactionId": "0f1e2d3c-4b5a-0968-8776-a5b4c3d2e1f0"}),
        ("transactionId variant c", {**one, "transactionId": "0f1e2d3c-4b5a-5968-c776-a5b4c3d2e1f0"}),
        ("transactionId with a trailing newline", {**one, "transactionId": _R4G2C_R1_TX + "\n"}),
        ("operationId not a UUID", with_op(operationId="SECRET")),
        ("operationCount 0 with no operations", {**one, "operationCount": 0, "operations": []}),
        # R4g-2c-r1 takeover: an empty report under an outcome without an operation-state invariant (the route never
        # replays one: no operation is committed or unknown); kills the operationCount >= 0 mutant.
        ("operationCount 0 with no operations (partial)",
         {**one, "outcome": "partial", "operationCount": 0, "operations": []}),
        ("operationCount 0 with no operations (replay)",
         {**one, "outcome": "replay", "operationCount": 0, "operations": []}),
        ("31 operations", _r4g2c_r1_report(31)),
        ("operationCount differs from operations", {**one, "operationCount": 2}),
        ("operationCount a bool", {**one, "operationCount": True}),
        ("operationCount a float", {**one, "operationCount": 1.0}),
        ("operationCount a string", {**one, "operationCount": "1"}),
        ("operations not a list", {**one, "operations": {"0": op}}),
        ("operation not an object", {**one, "operations": ["SECRET"]}),
        ("operationIndex negative", with_op(operationIndex=-1)),
        ("operationIndex above 2^53-1", with_op(operationIndex=2 ** 53)),
        ("groupIndex a float", with_op(groupIndex=0.5)),
        ("effect outside the enum", with_op(effect="SECRET")),
        ("state outside the enum", with_op(state="SECRET")),
        ("status outside the enum", with_op(status="SECRET")),
        ("state/status pair invalid (committed/failed)", with_op(status="failed")),
        ("state/status pair invalid (not_committed/applied)", {**one, "outcome": "partial",
                                                                "operations": [{**op, "state": "not_committed"}]}),
        ("committed outcome with an unknown operation",
         with_op(state="unknown_within_batch", status="unknown_within_batch")),
        ("stale outcome with a committed operation", {**one, "outcome": "stale"}),
        ("cancelled outcome with a committed operation", {**one, "outcome": "cancelled"}),
        ("no operation was committed (the route does not replay it)",
         {**one, "outcome": "partial", "operations": [{**op, "state": "not_committed", "status": "failed"}]}),
        ("firstFailedGroup negative", {**one, "firstFailedGroup": -1}),
        ("firstFailedOperationIndex a string", {**one, "firstFailedOperationIndex": "0"}),
        ("firstFailedGroup null", {**one, "firstFailedGroup": None}),
        ("execution a list", [one]),
        ("execution null", None),
    ]


@pytest.mark.parametrize("label,execution", _r4g2c_r1_bad_reports())
def test_r4g2c_r1_b_anything_but_an_executor_report_is_the_fixed_body_in_both_phases(label: str, execution: Any) -> None:
    for plan in (valid_plan(), valid_execute_plan()):
        text = _r4g2c_view_text(_r4g2c_r1_ae(execution), plan)
        assert json.loads(text) == {"executed": False, "result": _R4G2B_FIXED}, label
        assert "SECRET" not in text and "WEBSERVICE" not in text and len(text) < 1024, label


def test_r4g2c_r1_b2_the_result_wrapper_stays_closed_around_a_real_report() -> None:
    real = _r4g2c_r1_report()
    for label, body in [
        ("result extra key", _r4g2c_t2("already_executed", result={"execution": real, "leak": "SECRET"})),
        ("body extra key", {**_r4g2c_r1_ae(real), "leak": "SECRET"}),
        ("ask outcome around a report", _r4g2c_t2("ask", result={"execution": real})),
        ("summary not a string", {**_r4g2c_r1_ae(real), "summary": ["SECRET"]}),
    ]:
        assert _r4g2b_model_view(body, valid_execute_plan()) == {"executed": False, "result": _R4G2B_FIXED}, label


def test_r4g2c_r1_p_inspect_phase_accepts_only_the_summary_only_and_bridge_bodies() -> None:
    passing = [dict(_R4G2C_T1_INVALID), dict(_R4G2B_FIXED)] + \
        [_r4g2c_t2(reason) for reason in ("rejected", "cancelled", "not_started", "invalid")]
    for body in passing:
        assert _r4g2b_model_view(body, valid_plan()) == {"executed": False, "result": body}
        assert _r4g2b_model_view(body, valid_execute_plan()) == {"executed": False, "result": body}
    for reason in ("ask", "refuse"):
        body = _r4g2c_t2(reason, result={"outcome": reason, "finalSummary": "Which sheet?"})
        assert _r4g2b_model_view(body, valid_plan()) == {"executed": False, "result": _R4G2B_FIXED}
        assert _r4g2b_model_view(body, valid_execute_plan()) == {"executed": False, "result": body}


@pytest.mark.parametrize("label,body", [
    ("T1 invalid plus a leak key", {**_R4G2C_T1_INVALID, "leak": "canary"}),
    ("T1 invalid plus category under reason invalid", {**_R4G2C_T1_INVALID, "category": "Execution"}),
    ("T1 rejected plus a leak key", {**_R4G2B_FIXED, "leak": "canary"}),
])
def test_r4g2c_r1_k_the_t1_bridge_key_sets_are_exact(label: str, body: dict) -> None:
    for plan in (valid_plan(), valid_execute_plan()):
        text = _r4g2c_view_text(body, plan)
        assert json.loads(text) == {"executed": False, "result": _R4G2B_FIXED}, label
        assert "canary" not in text, label


def test_r4g2c_r1_n_js_number_text_at_exactly_21_integer_digits() -> None:
    module, _context = registered()
    # JS String(x) for |x| < 1e21 with 21 integer digits is the plain digit string; 1e21 switches to exponent form.
    assert module._js_number_text(1e20) == "100000000000000000000"
    assert module._js_number_text(-1e20) == "-100000000000000000000"
    assert module._js_number_text(1.2345678901234568e20) == "123456789012345680000"
    assert module._js_number_text(9.999999999999999e20) == "999999999999999900000"
    assert module._js_number_text(1e21) == "1e+21"
    assert module._js_number_text(1e19) == "10000000000000000000"


def test_r4g2c_r1_cap_refusal_byte_cap_is_inclusive() -> None:
    # No closed body reaches 64 KiB any more (the largest is a 30-operation report or 2048 escaped code points), so the
    # byte-cap comparison is pinned with the cap lowered to the exact JS byte size of a real body (all ASCII, written
    # out as JSON.stringify writes it).
    body = _r4g2c_r1_ae(_r4g2c_r1_report(2))
    js = ('{"ok":false,"action":"workbook_plan","reason":"already_executed","summary":"'
          + _R4G2C_SUMMARY["already_executed"] + '","result":{"execution":{"outcome":"committed","transactionId":"'
          + _R4G2C_R1_TX + '","operationCount":2,"operations":['
          + ",".join('{"operationId":"' + o["operationId"] + '","operationIndex":' + str(i) + ',"effect":"write",'
                     '"groupIndex":' + str(i) + ',"state":"committed","status":"applied"}'
                     for i, o in enumerate(body["result"]["execution"]["operations"]))
          + '],"verification":"not_attempted","rollback":"not_attempted"}}}')
    assert json.loads(js) == body
    module, _context = registered()
    with patch.object(module, "_REFUSAL_BODY_MAX_BYTES", len(js)):
        assert module._verified_not_executed_body(copy.deepcopy(body), "execute") == body
    with patch.object(module, "_REFUSAL_BODY_MAX_BYTES", len(js) - 1):
        assert module._verified_not_executed_body(copy.deepcopy(body), "execute") == _R4G2B_FIXED


# ---------------------------------------------------------------------------
# R4g-3 fold-in (OMH release record 2026-10-07; r4g2c/r1/tester-report-A.md G1/G2, lane-A survivors M2/M3): the
# already_executed report check, pinned row by row against the executor's own rules (workbookPlanExecutor.ts
# validStateStatus / validReportInvariants). Parity pins: production is unchanged and already refuses these.
# ---------------------------------------------------------------------------
_R4G3_STATES = ("committed", "not_committed", "unknown_within_batch")
_R4G3_STATUSES = ("applied", "verified", "rolled_back", "not_started", "failed", "unknown_within_batch")
# workbookPlanExecutor.ts:1262-1265 validStateStatus, written out independently.
_R4G3_VALID_PAIRS = frozenset({
    ("committed", "applied"), ("committed", "verified"), ("committed", "rolled_back"),
    ("not_committed", "not_started"), ("not_committed", "failed"),
    ("unknown_within_batch", "unknown_within_batch"), ("unknown_within_batch", "rolled_back"),
})


def _r4g3_two_op_report(outcome: str, state: str, status: str) -> dict:
    """Operation 0 committed/applied (so the route would replay it); operation 1 carries the pair under test."""
    report = _r4g2c_r1_report(2, outcome=outcome)
    report["operations"][1] = {**report["operations"][1], "state": state, "status": status}
    return report


@pytest.mark.parametrize("state,status", [
    ("unknown_within_batch", "unknown_within_batch"),
    ("unknown_within_batch", "rolled_back"),
    ("not_committed", "failed"),
    ("not_committed", "not_started"),
])
def test_r4g3_g1_a_committed_report_with_mixed_operation_states_is_refused(state: str, status: str) -> None:
    assert (state, status) in _R4G3_VALID_PAIRS  # each operation is valid on its own; only the mix is wrong
    mixed = _r4g2c_r1_ae(_r4g3_two_op_report("committed", state, status))
    for plan in (valid_plan(), valid_execute_plan()):
        text = _r4g2c_view_text(mixed, plan)
        assert json.loads(text) == {"executed": False, "result": _R4G2B_FIXED}, (state, status)
    # controls: the same operations under outcome partial (no state invariant) and an all-committed report pass
    for control in (_r4g2c_r1_ae(_r4g3_two_op_report("partial", state, status)),
                    _r4g2c_r1_ae(_r4g3_two_op_report("committed", "committed", "verified"))):
        assert _r4g2b_model_view(control, valid_execute_plan()) == {"executed": False, "result": control}, (state, status)


@pytest.mark.parametrize("state,status", [(state, status) for state in _R4G3_STATES for status in _R4G3_STATUSES])
def test_r4g3_g2_every_state_status_pair_is_pinned_one_by_one(state: str, status: str) -> None:
    assert len(_R4G3_VALID_PAIRS) == 7 and len(_R4G3_STATES) * len(_R4G3_STATUSES) - len(_R4G3_VALID_PAIRS) == 11
    body = _r4g2c_r1_ae(_r4g3_two_op_report("partial", state, status))
    expected = body if (state, status) in _R4G3_VALID_PAIRS else _R4G2B_FIXED
    assert _r4g2b_model_view(body, valid_execute_plan()) == {"executed": False, "result": expected}, (state, status)
    # the inspect phase never admits a result-bearing refusal, valid or not
    assert _r4g2b_model_view(body, valid_plan()) == {"executed": False, "result": _R4G2B_FIXED}, (state, status)


def test_r6a_generated_workbook_plan_skill_matches_schema_and_drift_gate() -> None:
    skill_path = PLUGIN_DIR / "skills" / "workbook-plans" / "SKILL.md"
    assert skill_path.is_file(), "Generated workbook plan skill is missing"
    text = skill_path.read_text(encoding="utf-8")
    assert f"Schema fingerprint: {generated_fingerprint()}." in text
    assert generated_fingerprint() == plan_schema()["x-workbook-plan-fingerprint"]
    node = shutil.which("node")
    assert node is not None
    result = subprocess.run(
        [node, str(REPO / "scripts/generate-office-plugin-schemas.mjs"), "--check"],
        cwd=REPO, capture_output=True, text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def test_r6a_r1_guidance_records_possible_change_and_terminal_call_stop() -> None:
    text = (PLUGIN_DIR / "skills/workbook-plans/SKILL.md").read_text(encoding="utf-8")
    assert "At most one plan call per response." in text
    assert "was already applied earlier" not in text
    assert "an earlier attempt of this same operation may already have changed the workbook" in text
    assert "it may be partial or unknown" in text
    assert "make no further plan calls in that turn; report the outcome instead" in text
    assert "2,048 Unicode code points" in text
    assert "ask / refuse / finalSummary text" in text


@pytest.mark.parametrize("method", ["exists", "lexists", "isfile", "isdir", "islink"])
def test_w0_metadata_guard_covers_path_fast_paths_and_restores(method: str) -> None:
    original = getattr(_g5_os.path, method)
    state = _G5GuardState({str(SCHEMAS)})
    forbidden = SCHEMAS.parent / "w0-forbidden-metadata"
    with _g5_guards(state):
        _g5_expect_denial(state, method, lambda: getattr(_g5_os.path, method)(forbidden))
        assert _g5_os.path.exists(SCHEMAS) is True
    assert getattr(_g5_os.path, method) is original
    assert state.forbidden_original_calls == 0
