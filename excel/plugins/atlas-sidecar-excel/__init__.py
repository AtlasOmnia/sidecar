"""Hermes plugin exposing Excel Office.js tools through the companion bridge."""
from __future__ import annotations

import copy
import json
import math
import re
import sys
import threading
import uuid
from collections import OrderedDict
from contextvars import ContextVar
from pathlib import Path
from typing import Any, Callable, Optional, Union

from . import desk_app
from .desk_bridge import DeskRuntime, DeskError
from .bridge_client import BridgeError, invoke_bridge as _invoke_bridge_http

bridge_client = sys.modules[__name__ + ".bridge_client"]

_RUNTIME: Optional[DeskRuntime] = None
_RUNTIME_ERROR: Optional[str] = None

def invoke_bridge(session_id: str, tool: str, arguments: dict[str, Any], *, trusted: Any = None) -> dict[str, Any]:
    """Send one canonical Excel tool to the sidebar paired with this chat (this process's bridge)."""
    runtime = _RUNTIME
    if runtime is None:
        raise BridgeError("Configuration", _RUNTIME_ERROR or "Sidecar for Excel is not configured")
    runtime.key_provider()  # a tool call runs in the profile's scope: refresh the in-memory API key
    if runtime.bridge.port is None:
        try:
            runtime.ensure_started()  # first use
        except DeskError as exc:
            raise BridgeError(exc.category, str(exc)) from None
    if not runtime.bridge.has_session(session_id) or not runtime.bridge.session_allowed(session_id):
        raise BridgeError("ForeignSession", "No Sidecar for Excel sidebar is connected to this chat. Open the "
                          "Sidecar sidebar in Excel, connect it, and chat from there.")
    return _invoke_bridge_http(session_id, tool, arguments, base_url=runtime.bridge.base_url,
                               key=runtime.bridge.plugin_key, trusted=trusted,
                               ssl_context=runtime.bridge.tls.client_context())



_TOOLSET = "sidecar-excel"
_SCHEMA_PATH = Path(__file__).with_name("tool_schemas.json")
_PLAN_TOOL = "sidecar_excel_execute_plan"
_PLAN_BRIDGE_TOOL = "execute_plan"
_OPERATION_NAMESPACE = uuid.UUID("5f8f4f1c-4c57-4a35-8b4f-c0b2aa3f9248")
_MAX_TRUSTED_METADATA_BYTES = 256

class _TrustedPlanHolder:
    """A bounded, shared one-shot lifecycle for copied execution contexts."""

    __slots__ = ("_context", "_lock")

    def __init__(self, context: dict[str, Any]) -> None:
        self._context: Optional[dict[str, Any]] = context
        self._lock = threading.Lock()

    def peek(self) -> Optional[dict[str, Any]]:
        with self._lock:
            return self._context

    def consume(self) -> Optional[dict[str, Any]]:
        with self._lock:
            context = self._context
            self._context = None
            return context

    def clear(self) -> None:
        with self._lock:
            self._context = None


# The plan handler is deliberately bound to metadata captured from Hermes'
# trusted dispatch kwargs. ContextVar keeps concurrent tool calls isolated;
# the holder is shared by copied worker contexts only as a one-shot lifecycle.
# The model's arguments never participate in identity construction.
_trusted_plan_context: ContextVar[Optional[_TrustedPlanHolder]] = ContextVar(
    "hermes_excel_trusted_plan_context", default=None
)
# Advisory only (R3a): the companion is the single authority for plan-op
# acceptance. This process-global set may forget (LRU) without ever allowing
# a second ticket: the same (session, turn) derives the same operation id, so
# the companion answers existing, conflict, replay_evicted, or 429 capacity.
# R4g-2a: each (session, turn) entry records the plan phases admitted for that
# Hermes turn (at most one inspect, then at most one execute; nothing after an
# execute) and the count of invalid plan attempts (the second exhausts the turn).
_seen_plan_turns: OrderedDict[tuple[str, str], dict[str, Any]] = OrderedDict()
_seen_plan_turns_lock = threading.Lock()
_MAX_SEEN_PLAN_TURNS = 1024
_TRUSTED_CONTEXT_REJECTED = "Workbook plan trusted context rejected."


def _load_schemas() -> list[dict[str, Any]]:
    value = json.loads(_SCHEMA_PATH.read_text(encoding="utf-8"))
    if not isinstance(value, list):
        raise RuntimeError("Excel plugin schema artifact must be an array")
    names: set[str] = set()
    for schema in value:
        if not isinstance(schema, dict) or not isinstance(schema.get("name"), str):
            raise RuntimeError("Excel plugin contains an invalid tool schema")
        name = schema["name"]
        if not name.startswith("sidecar_excel_") or name in names:
            raise RuntimeError("Excel plugin tool names must be unique and namespaced")
        names.add(name)
    return value


_SCHEMAS = _load_schemas()
_PLAN_SCHEMA = next((schema for schema in _SCHEMAS if schema.get("name") == _PLAN_TOOL), None)
if _PLAN_SCHEMA is None or not isinstance(_PLAN_SCHEMA.get("parameters"), dict):
    raise RuntimeError("Excel plugin is missing the generated WorkbookPlan schema")
_PLAN_PARAMETERS: dict[str, Any] = _PLAN_SCHEMA["parameters"]
_PLAN_FINGERPRINT = _PLAN_SCHEMA.get("x-workbook-plan-fingerprint")
if not isinstance(_PLAN_FINGERPRINT, str) or len(_PLAN_FINGERPRINT) != 64:
    raise RuntimeError("Excel plugin has an invalid WorkbookPlan fingerprint")


def _safe_metadata(value: Any) -> bool:
    return (
        isinstance(value, str)
        and bool(value)
        and len(value.encode("utf-8")) <= _MAX_TRUSTED_METADATA_BYTES
        and all(ord(char) >= 33 and ord(char) != 127 for char in value)
    )


def _canonical_session_id(value: Any) -> str:
    if not isinstance(value, str):
        raise BridgeError("Validation", "Trusted session_id must be a UUID")
    try:
        parsed = uuid.UUID(value)
    except (ValueError, AttributeError) as exc:
        raise BridgeError("Validation", "Trusted session_id must be a UUID") from exc
    if str(parsed) != value.lower():
        raise BridgeError("Validation", "Trusted session_id must be a UUID")
    return value


def _schema_matches(value: Any, schema: dict[str, Any], root: dict[str, Any]) -> bool:
    if "$ref" in schema:
        reference = schema.get("$ref")
        if not isinstance(reference, str) or not reference.startswith("#/$defs/"):
            return False
        definition = root.get("$defs", {}).get(reference.removeprefix("#/$defs/"))
        return isinstance(definition, dict) and _schema_matches(value, definition, root)
    if "const" in schema and value != schema["const"]:
        return False
    if "enum" in schema and value not in schema["enum"]:
        return False
    if "oneOf" in schema:
        branches = schema.get("oneOf")
        if not isinstance(branches, list):
            return False
        return sum(
            isinstance(branch, dict) and _schema_matches(value, branch, root)
            for branch in branches
        ) == 1

    schema_type = schema.get("type")
    if schema_type == "object":
        if not isinstance(value, dict):
            return False
        required = schema.get("required", [])
        properties = schema.get("properties", {})
        if not isinstance(required, list) or not isinstance(properties, dict):
            return False
        if any(name not in value for name in required):
            return False
        if schema.get("additionalProperties") is False and any(name not in properties for name in value):
            return False
        return all(
            name not in properties
            or not isinstance(properties[name], dict)
            or _schema_matches(item, properties[name], root)
            for name, item in value.items()
        )
    if schema_type == "array":
        if not isinstance(value, list):
            return False
        if len(value) < schema.get("minItems", 0) or len(value) > schema.get("maxItems", len(value)):
            return False
        items = schema.get("items")
        return isinstance(items, dict) and all(_schema_matches(item, items, root) for item in value)
    if schema_type == "string":
        return (
            isinstance(value, str)
            and len(value) >= schema.get("minLength", 0)
            and len(value) <= schema.get("maxLength", len(value))
        )
    if schema_type == "integer":
        return isinstance(value, int) and not isinstance(value, bool) and _numeric_bounds(value, schema)
    if schema_type == "number":
        return isinstance(value, (int, float)) and not isinstance(value, bool) and _numeric_bounds(value, schema)
    if schema_type == "boolean":
        return isinstance(value, bool)
    if schema_type == "null":
        return value is None
    return True


def _numeric_bounds(value: Union[int, float], schema: dict[str, Any]) -> bool:
    if "minimum" in schema and value < schema["minimum"]:
        return False
    if "maximum" in schema and value > schema["maximum"]:
        return False
    if "exclusiveMinimum" in schema and value <= schema["exclusiveMinimum"]:
        return False
    if "exclusiveMaximum" in schema and value >= schema["exclusiveMaximum"]:
        return False
    return True


def _valid_plan(value: Any) -> bool:
    return isinstance(value, dict) and value.get("version") == 1 and _schema_matches(
        value, _PLAN_PARAMETERS, _PLAN_PARAMETERS
    )


def _private_fingerprint_from_kwargs(kwargs: dict[str, Any]) -> Optional[str]:
    # The server may provide a pane's negotiated fingerprint as trusted turn
    # metadata.  When present it must match this static plugin artifact before
    # a ticket is even considered.  Older Hermes runtimes omit it; the private
    # envelope still carries this plugin fingerprint for the companion check.
    supplied = None
    for key in ("workbook_plan_fingerprint", "pane_plan_fingerprint", "plan_fingerprint"):
        if key in kwargs:
            supplied = kwargs[key]
            break
    if supplied is None:
        return _PLAN_FINGERPRINT
    if supplied != _PLAN_FINGERPRINT:
        return None
    return supplied


def _clear_trusted_plan_context() -> None:
    holder = _trusted_plan_context.get()
    if holder is not None:
        holder.clear()
    _trusted_plan_context.set(None)


def _plan_phase(args: dict[str, Any]) -> str:
    """R4g-2a: the trusted phase comes only from the schema-validated outcome."""
    return "inspect" if args.get("outcome") == "inspect" else "execute"


def _block_trusted_plan() -> dict[str, str]:
    return {"action": "block", "message": _TRUSTED_CONTEXT_REJECTED}


def _remember_trusted_plan_context(*, tool_name: str = "", args: Any = None, **kwargs: Any) -> Optional[dict[str, str]]:
    if tool_name != _PLAN_TOOL:
        return None
    _clear_trusted_plan_context()
    session_id = kwargs.get("session_id")
    turn_id = kwargs.get("turn_id")
    task_id = kwargs.get("task_id")
    tool_call_id = kwargs.get("tool_call_id")
    api_request_id = kwargs.get("api_request_id")
    try:
        session = _canonical_session_id(session_id)
    except BridgeError:
        return _block_trusted_plan()
    if not all(_safe_metadata(value) for value in (turn_id, task_id, tool_call_id, api_request_id)):
        return _block_trusted_plan()
    fingerprint = _private_fingerprint_from_kwargs(kwargs)
    turn = str(turn_id)
    turn_key = (session, turn)
    valid = fingerprint is not None and isinstance(args, dict) and _valid_plan(args)
    phase = _plan_phase(args) if valid else ""
    with _seen_plan_turns_lock:
        entry = _seen_plan_turns.get(turn_key)
        if entry is None:
            entry = {"phases": set(), "invalid": 0}
            _seen_plan_turns[turn_key] = entry
        _seen_plan_turns.move_to_end(turn_key)
        # LRU eviction here is safe; rejecting at capacity would disable the
        # plan tool for the whole process. The companion fence is authoritative.
        while len(_seen_plan_turns) > _MAX_SEEN_PLAN_TURNS:
            _seen_plan_turns.popitem(last=False)
        if entry["invalid"] >= 2:
            return _block_trusted_plan()
        if not valid:
            # One bounded correction opportunity per Hermes turn: the second
            # invalid attempt exhausts the turn.
            entry["invalid"] += 1
            return _block_trusted_plan()
        phases = entry["phases"]
        if "execute" in phases or (phase == "inspect" and phases):
            return _block_trusted_plan()
        phases.add(phase)
    turn_key_id = str(uuid.uuid5(_OPERATION_NAMESPACE, f"{session}\\0{turn}"))
    operation_id = str(uuid.uuid5(_OPERATION_NAMESPACE, f"{session}\\0{turn}\\0{phase}"))
    context = {
        "args": copy.deepcopy(args),
        "session_id": session,
        "turn_id": turn,
        "task_id": task_id,
        "tool_call_id": tool_call_id,
        "api_request_id": api_request_id,
        "operation_id": operation_id,
        "fingerprint": fingerprint,
        "phase": phase,
        "turn_key": turn_key_id,
    }
    _trusted_plan_context.set(_TrustedPlanHolder(context))
    return None


# R4g-2b (R4g-2a lane-B A3): the plugin re-verifies every executed:false body before the model sees it. A refusal
# (an object with ok False) passes unchanged; anything else must be the T8 inspection body of an inspect-phase call
# whose reads are plain objects with exactly {ok, data} (ok boolean), each <= 1 MiB and the body <= 1.5 MiB as compact
# UTF-8 JSON (the pane route's caps). Otherwise the model sees the bridge's fixed not-executed body.
# R4g-2c (R4g-2b lane-B A1-A4, lane-A F1/F2): a refusal passes only in one of the closed R4h T1/T2 shapes (bridge
# invalid/rejected, route reasons), display strings <= 2048 UTF-16 units and the body <= 64 KiB; a T8 body needs at
# least one read and each read at most 64 levels of nesting (the route's cap). Sizes are the UTF-8 bytes of the text the
# pane's JSON.stringify produced: the plugin only sees the companion's re-parsed value, so it re-derives JS number text
# and JS escapes (lone surrogates are \uXXXX); a value JS cannot have produced (NaN/Infinity, a non-JSON type) or any
# measuring error fails closed with the fixed body, never a tool error.
# R4g-2c-r1 (lane A F1, lane B B1/B2): display strings are bounded in code points (as the route's
# boundedCodePointLength and the R4h/R4j wire legs); already_executed needs result.execution in the executor's closed
# report shape (workbookPlanExecutor.ts isReport/isOperationReport/validStateStatus/validReportInvariants, 1..30
# operations: one primitive per plan operation, and the route replays only a report with an operation not
# "not_committed"); a refusal carrying a result (ask/refuse/already_executed) is accepted only in the execute phase.
_INSPECTION_READ_MAX_BYTES = 1_048_576
_INSPECTION_BODY_MAX_BYTES = 1_572_864
_INSPECTION_READ_MAX_DEPTH = 64
_REFUSAL_BODY_MAX_BYTES = 65_536
_REFUSAL_TEXT_MAX_CODE_POINTS = 2048
_ROUTE_REASONS_WITHOUT_RESULT = frozenset({"rejected", "cancelled", "not_started", "invalid"})
_REPORT_MAX_OPERATIONS = 30
_REPORT_MAX_SAFE_INTEGER = 2 ** 53 - 1
_REPORT_KEYS = frozenset({"outcome", "transactionId", "operationCount", "operations", "verification", "rollback"})
_REPORT_OPTIONAL_KEYS = frozenset({"firstFailedGroup", "firstFailedOperationIndex"})
_REPORT_OUTCOMES = frozenset({"committed", "partial", "stale", "cancelled", "replay", "conflict", "restart_rejected",
                              "rejected", "persistence_unavailable", "persistence_unavailable_after_mutation"})
_REPORT_VERIFICATIONS = frozenset({"not_attempted", "reconciled", "reconciliation_failed"})
_REPORT_ROLLBACKS = frozenset({"not_attempted", "unavailable", "succeeded", "failed"})
_REPORT_OPERATION_KEYS = frozenset({"operationId", "operationIndex", "effect", "groupIndex", "state", "status"})
_REPORT_EFFECTS = frozenset({"write", "format", "structure", "filter", "calculate"})
_REPORT_STATE_STATUSES = {
    "committed": frozenset({"applied", "verified", "rolled_back"}),
    "not_committed": frozenset({"not_started", "failed"}),
    "unknown_within_batch": frozenset({"unknown_within_batch", "rolled_back"}),
}
_REPORT_STATES = frozenset(_REPORT_STATE_STATUSES)
_REPORT_UUID = re.compile("[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}")
_SURROGATE = re.compile("[\ud800-\udfff]")
_NOT_EXECUTED_BODY = {
    "ok": False, "action": "workbook_plan", "category": "Execution", "reason": "rejected",
    "error": "Workbook plan was not executed.",
}


# ECMAScript Number::toString of a finite float (Python repr has the same shortest round-trip digits).
def _js_number_text(value: float) -> str:
    if value == 0:
        return "0"
    text = repr(value)
    sign = "-" if text.startswith("-") else ""
    mantissa, _, exponent = text.lstrip("-").partition("e")
    whole, _, fraction = mantissa.partition(".")
    digits = whole + fraction
    point = len(whole) + (int(exponent) if exponent else 0)
    stripped = digits.lstrip("0")
    point -= len(digits) - len(stripped)
    digits = stripped.rstrip("0")
    k = len(digits)
    if k <= point <= 21:
        return sign + digits + "0" * (point - k)
    if 0 < point <= 21:
        return sign + digits[:point] + "." + digits[point:]
    if -6 < point <= 0:
        return sign + "0." + "0" * (-point) + digits
    e = point - 1
    tail = ("e+" if e >= 0 else "e-") + str(abs(e))
    return sign + (digits if k == 1 else digits[0] + "." + digits[1:]) + tail


# UTF-8 bytes of JSON.stringify(value), or None when JS could not have produced it or it nests deeper than max_depth.
def _js_json_bytes(value: Any, max_depth: int) -> Optional[int]:
    adjust = 0
    stack: list[tuple[Any, int]] = [(value, 0)]
    while stack:
        item, depth = stack.pop()
        kind = type(item)
        if kind is dict or kind is list:
            if depth + 1 > max_depth:
                return None
            if kind is dict:
                if any(type(key) is not str for key in item):
                    return None
                stack.extend((child, depth + 1) for child in item.values())
            else:
                stack.extend((child, depth + 1) for child in item)
        elif kind is float:
            if not math.isfinite(item):
                return None
            adjust += len(_js_number_text(item)) - len(repr(item))
        elif not (item is None or kind is bool or kind is int or kind is str):
            return None
    try:
        text = json.dumps(value, separators=(",", ":"), ensure_ascii=False, allow_nan=False)
        size = len(text.encode("utf-8", "surrogatepass"))
    except (ValueError, TypeError, RecursionError, UnicodeError):
        return None
    return size + 3 * len(_SURROGATE.findall(text)) + adjust


def _report_count(value: Any) -> bool:
    return type(value) is int and 0 <= value <= _REPORT_MAX_SAFE_INTEGER


def _report_member(value: Any, allowed: frozenset) -> bool:
    return type(value) is str and value in allowed


def _report_uuid(value: Any) -> bool:
    return type(value) is str and _REPORT_UUID.fullmatch(value) is not None


def _replayed_execution_report(report: Any) -> bool:
    if type(report) is not dict or not _REPORT_KEYS <= set(report) <= _REPORT_KEYS | _REPORT_OPTIONAL_KEYS:
        return False
    operations, count = report["operations"], report["operationCount"]
    if not (_report_member(report["outcome"], _REPORT_OUTCOMES) and _report_uuid(report["transactionId"])
            and _report_count(count) and 1 <= count <= _REPORT_MAX_OPERATIONS
            and type(operations) is list and len(operations) == count
            and _report_member(report["verification"], _REPORT_VERIFICATIONS)
            and _report_member(report["rollback"], _REPORT_ROLLBACKS)
            and all(_report_count(report[key]) for key in _REPORT_OPTIONAL_KEYS if key in report)):
        return False
    states = set()
    for operation in operations:
        if type(operation) is not dict or set(operation) != _REPORT_OPERATION_KEYS:
            return False
        state = operation["state"]
        if not (_report_uuid(operation["operationId"]) and _report_count(operation["operationIndex"])
                and _report_member(operation["effect"], _REPORT_EFFECTS) and _report_count(operation["groupIndex"])
                and _report_member(state, _REPORT_STATES)
                and _report_member(operation["status"], _REPORT_STATE_STATUSES[state])):
            return False
        states.add(state)
    if report["outcome"] == "committed" and states != {"committed"}:
        return False
    if report["outcome"] in ("stale", "cancelled") and states != {"not_committed"}:
        return False
    return states != {"not_committed"}


def _closed_refusal_body(data: Any, phase: str) -> bool:
    if type(data) is not dict or data.get("ok") is not False or data.get("action") != "workbook_plan":
        return False
    keys, reason = set(data), data.get("reason")
    texts = [data[key] for key in ("error", "category", "summary") if key in data]
    if keys == {"ok", "action", "reason", "error"}:
        shaped = reason == "invalid"
    elif keys == {"ok", "action", "category", "reason", "error"}:
        shaped = reason == "rejected"
    elif keys == {"ok", "action", "reason", "summary"}:
        shaped = reason in _ROUTE_REASONS_WITHOUT_RESULT
    elif keys == {"ok", "action", "reason", "summary", "result"} and phase == "execute":
        result = data["result"]
        if reason in ("ask", "refuse"):
            shaped = type(result) is dict and set(result) == {"outcome", "finalSummary"} and result["outcome"] == reason
            if shaped:
                texts.append(result["finalSummary"])
        else:
            shaped = (reason == "already_executed" and type(result) is dict and set(result) == {"execution"}
                      and _replayed_execution_report(result["execution"]))
    else:
        shaped = False
    if not shaped or any(type(text) is not str or len(text) > _REFUSAL_TEXT_MAX_CODE_POINTS for text in texts):
        return False
    size = _js_json_bytes(data, _INSPECTION_READ_MAX_DEPTH)
    return size is not None and size <= _REFUSAL_BODY_MAX_BYTES


def _conforming_inspection_body(data: Any) -> bool:
    if type(data) is not dict or set(data) != {"ok", "action", "phase", "result"}:
        return False
    if data["ok"] is not True or data["action"] != "workbook_plan" or data["phase"] != "inspect":
        return False
    result = data["result"]
    if type(result) is not dict or set(result) != {"outcome", "inspectionCount", "reads"}:
        return False
    if result["outcome"] != "inspected" or type(result["inspectionCount"]) is not int or result["inspectionCount"] != 1:
        return False
    reads = result["reads"]
    if type(reads) is not list or not reads:
        return False
    for read in reads:
        if type(read) is not dict or set(read) != {"ok", "data"} or type(read["ok"]) is not bool:
            return False
        size = _js_json_bytes(read, _INSPECTION_READ_MAX_DEPTH)
        if size is None or size > _INSPECTION_READ_MAX_BYTES:
            return False
    size = _js_json_bytes(data, _INSPECTION_READ_MAX_DEPTH + 3)
    return size is not None and size <= _INSPECTION_BODY_MAX_BYTES


def _verified_not_executed_body(data: Any, phase: str) -> Any:
    try:
        if _closed_refusal_body(data, phase):
            return data
        if phase == "inspect" and _conforming_inspection_body(data):
            return data
    except (ValueError, TypeError, RecursionError, UnicodeError):
        pass
    return dict(_NOT_EXECUTED_BODY)


def _plan_handler(args: Any, **kwargs: Any) -> str:
    try:
        session_id = _canonical_session_id(kwargs.get("session_id"))
        holder = _trusted_plan_context.get()
        if holder is None:
            raise BridgeError("Validation", "Trusted plan context is absent or stale")
        context = holder.peek()
        if context is None:
            raise BridgeError("Validation", "Trusted plan context is absent or stale")
        if session_id != context["session_id"]:
            raise BridgeError("Validation", "Trusted plan context does not match the trusted session")
        context = holder.consume()
        _trusted_plan_context.set(None)
        if context is None:
            raise BridgeError("Validation", "Trusted plan context is absent or stale")
        for field in ("turn_id", "task_id", "tool_call_id", "api_request_id"):
            if field in kwargs and kwargs[field] != context[field]:
                raise BridgeError("Validation", "Trusted plan context does not match the trusted turn")
        for field in ("workbook_plan_fingerprint", "pane_plan_fingerprint", "plan_fingerprint"):
            if field in kwargs and kwargs[field] != context["fingerprint"]:
                raise BridgeError("Validation", "Trusted plan fingerprint does not match")
        if not isinstance(args, dict) or args != context["args"]:
            raise BridgeError("Validation", "Trusted plan context does not match the plan")
        if not _valid_plan(args):
            raise BridgeError("Validation", "WorkbookPlan version or shape is invalid")
        if _plan_phase(args) != context["phase"]:
            raise BridgeError("Validation", "Trusted plan context does not match the plan")
        trusted = {
            "turn_operation_id": context["operation_id"],
            "request_id": context["operation_id"],
            "workbook_plan_fingerprint": context["fingerprint"],
            "phase": context["phase"],
            "turn_key": context["turn_key"],
        }
        response = invoke_bridge(session_id, _PLAN_BRIDGE_TOOL, args, trusted=trusted)
        if response.get("ok") is True and response.get("executed") is False and "data" in response:
            # Never executed: the model sees an explicit bounded disposition.
            result = {"executed": False, "result": _verified_not_executed_body(response["data"], context["phase"])}
        else:
            result = response.get("data") if response.get("ok") is True and "data" in response else response
        return json.dumps(result, separators=(",", ":"))
    except BridgeError as exc:
        return json.dumps({
            "ok": False,
            "error": {"category": exc.category, "message": str(exc)},
        }, separators=(",", ":"))


def _clear_trusted_plan_context_after_call(*, tool_name: str = "", **_kwargs: Any) -> None:
    if tool_name == _PLAN_TOOL:
        _clear_trusted_plan_context()


def _handler_for(plugin_name: str) -> Callable[..., str]:
    canonical_name = plugin_name.removeprefix("sidecar_excel_")
    if plugin_name == _PLAN_TOOL:
        return _plan_handler

    def handler(args: dict[str, Any], **kwargs: Any) -> str:
        try:
            safe_args = dict(args or {})
            if "bridgeSessionId" in safe_args:
                raise BridgeError(
                    "Validation",
                    "bridgeSessionId is server-owned and must not be supplied in tool arguments",
                )
            session_id = _canonical_session_id(kwargs.get("session_id"))
            response = invoke_bridge(session_id, canonical_name, safe_args)
            result = response.get("data") if response.get("ok") is True and "data" in response else response
            return json.dumps(result, separators=(",", ":"))
        except BridgeError as exc:
            return json.dumps({
                "ok": False,
                "error": {
                    "category": exc.category,
                    "message": str(exc),
                },
            }, separators=(",", ":"))

    return handler


def register(ctx: Any) -> None:
    """Register tools and human-only approval commands; preserve plan hooks."""
    # Native Python plugins require explicit read-only skill registration.
    if hasattr(ctx, "register_skill"):
        ctx.register_skill("office-guidance", Path(__file__).parent / "skills/office-guidance/SKILL.md",
                           description='Use when building or analyzing a spreadsheet, budget, table or chart in native Sidecar Excel.')
        ctx.register_skill("workbook-plans", Path(__file__).parent / "skills/workbook-plans/SKILL.md",
                           description="Submit bounded WorkbookPlan inspection and execution in Sidecar Excel; plan mode only.")
    global _RUNTIME, _RUNTIME_ERROR
    try:
        _RUNTIME = DeskRuntime.from_ctx(ctx, desk_app)
        _RUNTIME_ERROR = None
    except DeskError as exc:
        _RUNTIME, _RUNTIME_ERROR = None, f"Sidecar for Excel settings: {exc}"

    for schema in _SCHEMAS:
        name = schema["name"]
        ctx.register_tool(
            name=name,
            toolset=_TOOLSET,
            schema=schema,
            handler=_handler_for(name),
            check_fn=None,
            override=False,
            description=schema.get("description", ""),
        )
    # Older local contract fakes do not expose hooks; real Hermes does.  The
    # server-owned plan path is fail-closed when this trusted hook is absent.
    register_hook = getattr(ctx, "register_hook", None)
    if callable(register_hook):
        register_hook("pre_tool_call", _remember_trusted_plan_context)
        register_hook("post_tool_call", _clear_trusted_plan_context_after_call)

    runtime = _RUNTIME
    if runtime is None:
        return
    if hasattr(ctx, "register_command"):
        ctx.register_command(
            desk_app.COMMAND, runtime.slash,
            description="Write a sideload manifest, allow, deny or revoke the Sidecar for Excel sidebar; status",
            args_hint="[manifest|allow|deny <code>|revoke|status|trust]",
        )
    if hasattr(ctx, "register_cli_command"):
        ctx.register_cli_command(
            name=desk_app.COMMAND, help="Sidecar for Excel: status, trust instructions, allow/deny a sidebar, revoke",
            setup_fn=runtime.cli_setup, handler_fn=runtime.cli_run,
        )
    if hasattr(ctx, "on_unload"):
        ctx.on_unload(runtime.stop)
    runtime.maybe_autostart()
