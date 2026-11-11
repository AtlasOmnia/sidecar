"""RED tests for the Task 13 plan-ticket lifecycle.

These tests intentionally exercise the real HTTP/session seams. They are added
before the production lifecycle repair so the baseline failure is meaningful.
"""
from __future__ import annotations

import io
import json
import logging
import threading

from desk_plugin.companion.protocol import CLAIM_LEASE_SECONDS
from desk_plugin.companion.server import ExcelCompanionServer
from scripts.tests.test_companion_plan_execution import (
    FINGERPRINT,
    OTHER_FINGERPRINT,
    PLAN,
    PLUGIN_KEY,
    _invoke_async,
    _pane_headers,
    _request,
    _session,
    _wait_for_ticket,
)

import uuid as _r4g2_uuid  # noqa: E402  (R4g-2a rule M)
_R4G2_TURN_KEY_NS = _r4g2_uuid.UUID("4b4b0000-0000-4000-8000-0000000000a2")


def _r4g2_turn_key(operation_id: str) -> str:
    """R4g-2a rule M: a per-operation turnKey (never equal to the operation id) for a migrated v1 envelope."""
    return str(_r4g2_uuid.uuid5(_R4G2_TURN_KEY_NS, str(operation_id)))


def test_plan_claim_ack_and_executing_timeout_are_one_way() -> None:
    clock_value = [1000.0]
    server = ExcelCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
        clock=lambda: clock_value[0], invoke_timeout_seconds=120,
    )
    server.start()
    try:
        session = _session(server)
        trusted = {
            "turnOperationId": "77777777-7777-4777-8777-777777777777",
            "workbookPlanFingerprint": FINGERPRINT,
            "phase": "execute",
            "turnKey": _r4g2_turn_key("77777777-7777-4777-8777-777777777777"),
        }
        invoke, invoke_result = _invoke_async(server, session, trusted["turnOperationId"], trusted)
        _wait_for_ticket(server, session["sessionId"])
        headers = _pane_headers(session)
        status, pending = _request(server, "GET", "/excel/pending", headers=headers)
        assert status == 200
        ticket = server.sessions.get(session["sessionId"]).get_ticket()
        assert ticket.state == "claimed"
        assert _request(server, "GET", "/excel/pending", headers=headers)[0] == 204

        start = {"ticketId": pending["ticketId"]}
        assert _request(server, "POST", "/excel/execute-start", start, headers)[0] == 200
        assert ticket.state == "executing"
        assert _request(server, "POST", "/excel/execute-start", start, headers)[0] == 409

        clock_value[0] += CLAIM_LEASE_SECONDS + 1
        server.sessions.sweep()
        assert ticket.state == "terminal"
        assert ticket.terminal_reason == "executing_uncertain"
        assert _request(server, "GET", "/excel/pending", headers=headers)[0] == 204
        invoke.join(2)
        assert not invoke.is_alive()
        # R2d-2: executing expiry is machine-readable uncertainty, not Timeout.
        assert invoke_result["status"] == 504
        assert invoke_result["body"] == {
            "ok": False,
            "error": "Workbook plan execution outcome is uncertain",
            "category": "ExecutionUncertain",
            "state": "executing_uncertain",
        }
    finally:
        server.stop()


def test_plan_cancel_and_lease_recovery_distinguish_execution_boundary() -> None:
    clock_value = [2000.0]
    server = ExcelCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
        clock=lambda: clock_value[0], invoke_timeout_seconds=120,
    )
    server.start()
    try:
        unclaimed = _session(server)
        trusted = {
            "turnOperationId": "88888888-8888-4888-8888-888888888888",
            "workbookPlanFingerprint": FINGERPRINT,
            "phase": "execute",
            "turnKey": _r4g2_turn_key("88888888-8888-4888-8888-888888888888"),
        }
        invoke, result = _invoke_async(server, unclaimed, trusted["turnOperationId"], trusted)
        _wait_for_ticket(server, unclaimed["sessionId"])
        ticket = server.sessions.get(unclaimed["sessionId"]).get_ticket()
        assert ticket.state == "pending"
        assert _request(
            server, "POST", "/excel/cancel",
            {"sessionId": unclaimed["sessionId"], "ticketId": ticket.ticket_id},
            {"X-Hermes-Plugin-Key": PLUGIN_KEY},
        )[0] == 200
        assert ticket.state == "terminal"
        assert ticket.terminal_reason == "cancelled_unclaimed"
        invoke.join(2)
        assert result["status"] == 499

        claimed = _session(server)
        trusted = {
            "turnOperationId": "99999999-9999-4999-8999-999999999999",
            "workbookPlanFingerprint": FINGERPRINT,
            "phase": "execute",
            "turnKey": _r4g2_turn_key("99999999-9999-4999-8999-999999999999"),
        }
        invoke, result = _invoke_async(server, claimed, trusted["turnOperationId"], trusted)
        _wait_for_ticket(server, claimed["sessionId"])
        headers = _pane_headers(claimed)
        assert _request(server, "GET", "/excel/pending", headers=headers)[0] == 200
        ticket = server.sessions.get(claimed["sessionId"]).get_ticket()
        # R2d-2 (M7): plans are never requeued. Stay inside the claim lease so
        # the claimed ticket is still cancellable; lease expiry is covered by
        # test_plan_claimed_lease_expiry_is_terminal_and_never_requeued.
        clock_value[0] += CLAIM_LEASE_SECONDS - 1
        server.sessions.sweep()
        assert ticket.state == "claimed"
        assert ticket.requeue_count == 0
        assert _request(server, "GET", "/excel/pending", headers=headers)[0] == 204
        assert _request(
            server, "POST", "/excel/cancel",
            {"sessionId": claimed["sessionId"], "ticketId": ticket.ticket_id},
            {"X-Hermes-Plugin-Key": PLUGIN_KEY},
        )[0] == 200
        assert ticket.state == "terminal"
        assert ticket.terminal_reason == "cancelled_claimed"
        invoke.join(2)
        assert result["status"] == 499
    finally:
        server.stop()


def test_plan_exact_duplicate_result_conflict_and_concurrent_replay() -> None:
    server = ExcelCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=2,
    )
    server.start()
    try:
        session = _session(server)
        trusted = {
            "turnOperationId": "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
            "workbookPlanFingerprint": FINGERPRINT,
            "phase": "execute",
            "turnKey": _r4g2_turn_key("aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"),
        }
        request_threads = [
            _invoke_async(server, session, trusted["turnOperationId"], trusted)
            for _ in range(2)
        ]
        for thread, _ in request_threads:
            _wait_for_ticket(server, session["sessionId"])
            assert thread.is_alive()
        headers = _pane_headers(session)
        status, pending = _request(server, "GET", "/excel/pending", headers=headers)
        assert status == 200
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
        payload = {"ok": True, "state": "verified"}
        body = {"ticketId": pending["ticketId"], "executed": True, "result": payload}
        assert _request(server, "POST", "/excel/result", body, headers)[0] == 200
        assert _request(server, "POST", "/excel/result", body, headers)[0] == 200
        assert _request(
            server, "POST", "/excel/result",
            {"ticketId": pending["ticketId"], "executed": True, "result": {"ok": False}}, headers,
        )[0] == 409
        assert _request(
            server, "POST", "/excel/result",
            {"ticketId": pending["ticketId"], "executed": False, "result": payload}, headers,
        )[0] == 409
        for thread, result in request_threads:
            thread.join(2)
            assert result["status"] == 200
    finally:
        server.stop()


def test_plan_identity_version_and_foreign_fields_fail_before_enqueue() -> None:
    server = ExcelCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=1,
    )
    server.start()
    try:
        session = _session(server)
        session_obj = server.sessions.get(session["sessionId"])
        request_id = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb"
        base = {
            "protocolVersion": "1.0", "sessionId": session["sessionId"],
            "requestId": request_id, "tool": "execute_plan", "arguments": PLAN,
        }
        for trusted in (
            None,
            {"turnOperationId": request_id},
            {"turnOperationId": request_id, "workbookPlanFingerprint": OTHER_FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key(request_id)},
            {"turnOperationId": "cccccccc-cccc-4ccc-8ccc-cccccccccccc", "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key("cccccccc-cccc-4ccc-8ccc-cccccccccccc")},
        ):
            body = dict(base)
            if trusted is not None:
                body["trusted"] = trusted
            status, _ = _request(server, "POST", "/excel/invoke", body, {"X-Hermes-Plugin-Key": PLUGIN_KEY})
            assert status in (400, 409)
            assert not session_obj.has_ticket()

        for version in (None, 0, 2):
            plan = dict(PLAN)
            if version is None:
                plan.pop("version")
            else:
                plan["version"] = version
            body = dict(base)
            body["arguments"] = plan
            body["trusted"] = {"turnOperationId": request_id, "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key(request_id)}
            status, _ = _request(server, "POST", "/excel/invoke", body, {"X-Hermes-Plugin-Key": PLUGIN_KEY})
            assert status == 400
            assert not session_obj.has_ticket()
    finally:
        server.stop()


def test_plan_later_trusted_turn_is_allowed() -> None:
    server = ExcelCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=2,
    )
    server.start()
    try:
        session = _session(server)
        headers = _pane_headers(session)
        for turn in (
            "13131313-1313-4131-8131-131313131313",
            "14141414-1414-4141-8141-141414141414",
        ):
            trusted = {"turnOperationId": turn, "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key(turn)}
            invoke, result = _invoke_async(server, session, turn, trusted)
            _wait_for_ticket(server, session["sessionId"])
            status, pending = _request(server, "GET", "/excel/pending", headers=headers)
            assert status == 200
            assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
            assert _request(server, "POST", "/excel/result", {"ticketId": pending["ticketId"], "executed": True, "result": {"ok": True}}, headers)[0] == 200
            invoke.join(2)
            assert result["status"] == 200
    finally:
        server.stop()


def test_plan_restart_rejects_old_session_claim_and_execute_start() -> None:
    first = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=2)
    first.start()
    try:
        session = _session(first)
        trusted = {"turnOperationId": "15151515-1515-4151-8151-151515151515", "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key("15151515-1515-4151-8151-151515151515")}
        invoke, result = _invoke_async(first, session, trusted["turnOperationId"], trusted)
        _wait_for_ticket(first, session["sessionId"])
        _, pending = _request(first, "GET", "/excel/pending", headers=_pane_headers(session))
        first.stop()
        invoke.join(2)
        assert result["status"] in (0, 499)
    finally:
        first.stop()

    second = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    second.start()
    try:
        headers = _pane_headers(session)
        assert _request(second, "GET", "/excel/pending", headers=headers)[0] == 401
        assert _request(second, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 401
    finally:
        second.stop()


def test_plan_evicted_terminal_identity_remains_non_executable() -> None:
    server = ExcelCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=2,
    )
    server.start()
    try:
        session = _session(server)
        session_obj = server.sessions.get(session["sessionId"])
        session_obj._history_limit = 1
        headers = _pane_headers(session)
        first_turn = "16161616-1616-4161-8161-161616161616"
        first_trusted = {"turnOperationId": first_turn, "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key(first_turn)}
        invoke, result = _invoke_async(server, session, first_turn, first_trusted)
        _wait_for_ticket(server, session["sessionId"])
        _, pending = _request(server, "GET", "/excel/pending", headers=headers)
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
        assert _request(server, "POST", "/excel/result", {"ticketId": pending["ticketId"], "executed": True, "result": {"ok": True}}, headers)[0] == 200
        invoke.join(2)
        assert result["status"] == 200

        second_turn = "17171717-1717-4171-8171-171717171717"
        second_trusted = {"turnOperationId": second_turn, "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key(second_turn)}
        invoke, result = _invoke_async(server, session, second_turn, second_trusted)
        _wait_for_ticket(server, session["sessionId"])
        _, pending = _request(server, "GET", "/excel/pending", headers=headers)
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
        assert _request(server, "POST", "/excel/result", {"ticketId": pending["ticketId"], "executed": True, "result": {"ok": True}}, headers)[0] == 200
        invoke.join(2)
        assert result["status"] == 200

        replay, result = _invoke_async(server, session, first_turn, first_trusted)
        replay.join(2)
        assert result["status"] == 409
    finally:
        server.stop()


def test_plan_failure_logs_and_errors_are_content_free() -> None:
    marker = "PLAN-CONTENT-SENTINEL"
    stream = io.StringIO()
    capture = logging.StreamHandler(stream)
    logger = logging.getLogger("excel_companion")
    logger.addHandler(capture)
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY)
    server.start()
    try:
        session = _session(server)
        plan = dict(PLAN)
        plan["message"] = marker
        body = {
            "protocolVersion": "1.0", "sessionId": session["sessionId"],
            "requestId": "18181818-1818-4181-8181-181818181818",
            "tool": "execute_plan", "arguments": plan,
            "trusted": {
                "turnOperationId": "18181818-1818-4181-8181-181818181818",
                "workbookPlanFingerprint": OTHER_FINGERPRINT,
                "phase": "execute",
                "turnKey": _r4g2_turn_key("18181818-1818-4181-8181-181818181818"),
            },
        }
        status, response = _request(server, "POST", "/excel/invoke", body, {"X-Hermes-Plugin-Key": PLUGIN_KEY})
        assert status == 409
        assert marker not in json.dumps(response)
    finally:
        server.stop()
        logger.removeHandler(capture)
    assert marker not in stream.getvalue()


def test_plan_changed_digest_reusing_request_identity_conflicts_before_queue() -> None:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=1)
    server.start()
    try:
        session = _session(server)
        turn = "19191919-1919-4191-8191-191919191919"
        trusted = {"turnOperationId": turn, "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key(turn)}
        first, result = _invoke_async(server, session, turn, trusted)
        _wait_for_ticket(server, session["sessionId"])
        headers = _pane_headers(session)
        _, pending = _request(server, "GET", "/excel/pending", headers=headers)
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
        assert _request(server, "POST", "/excel/result", {"ticketId": pending["ticketId"], "executed": True, "result": {"ok": True}}, headers)[0] == 200
        first.join(2)
        assert result["status"] == 200

        changed = dict(PLAN)
        changed["message"] = "changed"
        body = {
            "protocolVersion": "1.0", "sessionId": session["sessionId"],
            "requestId": turn, "tool": "execute_plan", "arguments": changed,
            "trusted": trusted,
        }
        status, _ = _request(server, "POST", "/excel/invoke", body, {"X-Hermes-Plugin-Key": PLUGIN_KEY})
        assert status == 409
        assert _request(server, "GET", "/excel/pending", headers=headers)[0] == 204
    finally:
        server.stop()


def test_plan_foreign_document_and_pane_identity_fail_before_execution() -> None:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=1)
    server.start()
    try:
        session_a = _session(server)
        session_b = _session(server)
        turn = "20202020-2020-4020-8020-202020202020"
        trusted = {"turnOperationId": turn, "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key(turn)}
        body = {
            "protocolVersion": "1.0", "sessionId": session_a["sessionId"],
            "requestId": turn, "tool": "execute_plan", "arguments": PLAN,
            "trusted": trusted, "documentId": "doc-b",
        }
        status, _ = _request(server, "POST", "/excel/invoke", body, {"X-Hermes-Plugin-Key": PLUGIN_KEY})
        assert status == 400
        assert not server.sessions.get(session_a["sessionId"]).has_ticket()
        assert _request(server, "GET", "/excel/pending", headers=_pane_headers(session_b))[0] == 204

        # The authenticated pane binding is exact; pane B cannot claim A.
        thread, result = _invoke_async(server, session_a, turn, trusted)
        _wait_for_ticket(server, session_a["sessionId"])
        foreign_headers = _pane_headers(session_a)
        foreign_headers["X-Hermes-Pane-Id"] = "foreign-pane"
        assert _request(server, "GET", "/excel/pending", headers=foreign_headers)[0] == 401
        _request(server, "POST", "/excel/cancel", {"sessionId": session_a["sessionId"]}, {"X-Hermes-Plugin-Key": PLUGIN_KEY})
        thread.join(2)
        assert result["status"] == 499
    finally:
        server.stop()


def test_plan_executing_cancel_is_uncertain_and_not_reclaimable() -> None:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=2)
    server.start()
    try:
        session = _session(server)
        turn = "21212121-2121-4121-8121-212121212121"
        trusted = {"turnOperationId": turn, "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key(turn)}
        invoke, result = _invoke_async(server, session, turn, trusted)
        _wait_for_ticket(server, session["sessionId"])
        headers = _pane_headers(session)
        _, pending = _request(server, "GET", "/excel/pending", headers=headers)
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
        ticket = server.sessions.get(session["sessionId"]).get_ticket()
        assert _request(
            server, "POST", "/excel/cancel",
            {"sessionId": session["sessionId"], "ticketId": ticket.ticket_id},
            {"X-Hermes-Plugin-Key": PLUGIN_KEY},
        )[0] == 200
        assert ticket.state == "terminal"
        assert ticket.terminal_reason == "executing_uncertain"
        assert _request(server, "GET", "/excel/pending", headers=headers)[0] == 204
        invoke.join(2)
        assert result["status"] == 504
        assert result["body"]["category"] == "ExecutionUncertain"
        assert result["body"]["state"] == "executing_uncertain"
    finally:
        server.stop()



# ---------------------------------------------------------------------------
# R2d-2: one-way plan terminal lifecycle and machine-readable uncertain state
# ---------------------------------------------------------------------------

_UNCERTAIN_BODY = {
    "ok": False,
    "error": "Workbook plan execution outcome is uncertain",
    "category": "ExecutionUncertain",
    "state": "executing_uncertain",
}
_NOT_STARTED_BODY = {
    "ok": False,
    "error": "Invoke timed out",
    "category": "Timeout",
    "state": "not_started",
}
_CANCELLED_BODY = {"ok": False, "error": "Cancelled", "category": "Cancellation"}


def _plugin_headers() -> dict[str, str]:
    return {"X-Hermes-Plugin-Key": PLUGIN_KEY}


def _claimed_plan(server: ExcelCompanionServer, turn: str):
    session = _session(server)
    trusted = {"turnOperationId": turn, "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key(turn)}
    invoke, result = _invoke_async(server, session, turn, trusted)
    _wait_for_ticket(server, session["sessionId"])
    headers = _pane_headers(session)
    status, pending = _request(server, "GET", "/excel/pending", headers=headers)
    assert status == 200
    ticket = server.sessions.get(session["sessionId"]).get_ticket()
    assert ticket.state == "claimed"
    return session, headers, pending, ticket, invoke, result


def test_plan_result_requires_boolean_executed_disposition_and_leaves_state() -> None:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=3)
    server.start()
    try:
        session, headers, pending, ticket, invoke, result = _claimed_plan(
            server, "22222222-2222-4222-8222-222222222201",
        )
        for disposition in ("missing", None, "true", 1, 0, [], {}):
            body = {"ticketId": pending["ticketId"], "result": {"ok": True}}
            if disposition != "missing":
                body["executed"] = disposition
            status, _ = _request(server, "POST", "/excel/result", body, headers)
            assert status in (400, 409), (disposition, status)
            assert ticket.state == "claimed"
            assert ticket.result_payload is None
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
        for disposition in ("missing", None, "false", 1):
            body = {"ticketId": pending["ticketId"], "result": {"ok": True}}
            if disposition != "missing":
                body["executed"] = disposition
            status, _ = _request(server, "POST", "/excel/result", body, headers)
            assert status in (400, 409), (disposition, status)
            assert ticket.state == "executing"
            assert ticket.result_payload is None
        assert invoke.is_alive()
        _request(server, "POST", "/excel/cancel", {"sessionId": session["sessionId"]}, _plugin_headers())
        invoke.join(2)
        assert result["body"] == _UNCERTAIN_BODY
    finally:
        server.stop()


def test_plan_result_before_execute_start_cannot_claim_execution() -> None:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=3)
    server.start()
    try:
        session, headers, pending, ticket, invoke, result = _claimed_plan(
            server, "22222222-2222-4222-8222-222222222202",
        )
        status, body = _request(
            server, "POST", "/excel/result",
            {"ticketId": pending["ticketId"], "executed": True, "result": {"ok": True}}, headers,
        )
        assert status == 409
        assert body["category"] == "Conflict"
        assert ticket.state == "claimed"
        assert ticket.terminal_reason is None
        assert ticket.result_payload is None
        assert invoke.is_alive()
        _request(server, "POST", "/excel/cancel", {"sessionId": session["sessionId"]}, _plugin_headers())
        invoke.join(2)
        assert result["status"] == 499
        assert result["body"] == _CANCELLED_BODY
    finally:
        server.stop()


def test_plan_not_executed_from_executing_is_rejected_and_completed_shape_is_pinned() -> None:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=3)
    server.start()
    try:
        turn = "22222222-2222-4222-8222-222222222203"
        session, headers, pending, ticket, invoke, result = _claimed_plan(server, turn)
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
        status, _ = _request(
            server, "POST", "/excel/result",
            {"ticketId": pending["ticketId"], "executed": False, "result": {"reason": "declined"}}, headers,
        )
        assert status == 409
        assert ticket.state == "executing"
        assert ticket.result_payload is None
        report = {"outcome": "committed", "operations": 1}
        assert _request(
            server, "POST", "/excel/result",
            {"ticketId": pending["ticketId"], "executed": True, "result": report}, headers,
        )[0] == 200
        assert ticket.terminal_reason == "completed"
        invoke.join(2)
        assert result["status"] == 200
        assert result["body"] == {
            "ok": True,
            "ticketId": pending["ticketId"],
            "requestId": turn,
            "data": report,
            "executed": True,
        }
    finally:
        server.stop()


def test_plan_not_executed_happy_path_duplicate_and_conflict() -> None:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=3)
    server.start()
    try:
        turn = "22222222-2222-4222-8222-222222222204"
        session, headers, pending, ticket, invoke, result = _claimed_plan(server, turn)
        report = {"executed": False, "reason": "declined"}
        body = {"ticketId": pending["ticketId"], "executed": False, "result": report}
        assert _request(server, "POST", "/excel/result", body, headers)[0] == 200
        assert ticket.state == "terminal"
        assert ticket.terminal_reason == "not_executed"
        assert ticket.result_payload == report
        # M12: exact duplicate is idempotent; conflicting duplicates conflict.
        assert _request(server, "POST", "/excel/result", body, headers)[0] == 200
        assert _request(
            server, "POST", "/excel/result",
            {"ticketId": pending["ticketId"], "executed": False, "result": {"reason": "cancelled"}}, headers,
        )[0] == 409
        assert _request(
            server, "POST", "/excel/result",
            {"ticketId": pending["ticketId"], "executed": True, "result": report}, headers,
        )[0] == 409
        assert ticket.terminal_reason == "not_executed"
        assert ticket.result_payload == report
        # A late acknowledgement after a non-executing terminal is refused.
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 409
        invoke.join(2)
        assert result["status"] == 200
        assert result["body"] == {
            "ok": True,
            "ticketId": pending["ticketId"],
            "requestId": turn,
            "data": report,
            "executed": False,
        }
    finally:
        server.stop()


def test_plan_concurrent_conflicting_results_exactly_one_wins() -> None:
    for executing in (False, True):
        server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=3)
        server.start()
        try:
            turn = "22222222-2222-4222-8222-22222222220" + ("6" if executing else "5")
            session, headers, pending, ticket, invoke, result = _claimed_plan(server, turn)
            if executing:
                assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
            count = 4  # stays within the server's listen backlog
            barrier = threading.Barrier(count)
            statuses: list[tuple[int, dict]] = [(-1, {})] * count

            def post(index: int) -> None:
                body = {
                    "ticketId": pending["ticketId"],
                    "executed": executing,
                    "result": {"writer": index},
                }
                barrier.wait()
                statuses[index] = (
                    _request(server, "POST", "/excel/result", body, headers)[0],
                    body["result"],
                )

            threads = [threading.Thread(target=post, args=(i,)) for i in range(count)]
            for thread in threads:
                thread.start()
            for thread in threads:
                thread.join(5)
            winners = [payload for status, payload in statuses if status == 200]
            assert len(winners) == 1, statuses
            assert sorted(status for status, _ in statuses) == [200] + [409] * (count - 1)
            assert ticket.result_payload == winners[0]
            assert ticket.terminal_reason == ("completed" if executing else "not_executed")
            invoke.join(2)
            assert result["status"] == 200
            assert result["body"]["data"] == winners[0]
            assert result["body"]["executed"] is executing
        finally:
            server.stop()


def test_plan_wrong_pane_result_is_rejected_without_state_change() -> None:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=3)
    server.start()
    try:
        session, headers, pending, ticket, invoke, result = _claimed_plan(
            server, "22222222-2222-4222-8222-222222222207",
        )
        ticket.claimed_by_pane_id = "another-pane"
        for executed in (False, True):
            status, body = _request(
                server, "POST", "/excel/result",
                {"ticketId": pending["ticketId"], "executed": executed, "result": {"ok": True}}, headers,
            )
            assert status == 409
            assert ticket.state == "claimed"
            assert ticket.result_payload is None
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 409
        assert ticket.state == "claimed"
        ticket.claimed_by_pane_id = "pane"
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
        ticket.claimed_by_pane_id = "another-pane"
        status, _ = _request(
            server, "POST", "/excel/result",
            {"ticketId": pending["ticketId"], "executed": True, "result": {"ok": True}}, headers,
        )
        assert status == 409
        assert ticket.state == "executing"
        assert ticket.result_payload is None
        _request(server, "POST", "/excel/cancel", {"sessionId": session["sessionId"]}, _plugin_headers())
        invoke.join(2)
        assert result["body"] == _UNCERTAIN_BODY
    finally:
        server.stop()


def test_plan_claimed_lease_expiry_is_terminal_and_never_requeued() -> None:
    clock_value = [3000.0]
    server = ExcelCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
        clock=lambda: clock_value[0], invoke_timeout_seconds=120,
    )
    server.start()
    try:
        session, headers, pending, ticket, invoke, result = _claimed_plan(
            server, "22222222-2222-4222-8222-222222222208",
        )
        clock_value[0] += CLAIM_LEASE_SECONDS + 1
        server.sessions.sweep()
        assert ticket.state == "terminal"
        assert ticket.terminal_reason == "expired_claimed"
        assert ticket.requeue_count == 0
        # Never redelivered to the pane.
        assert _request(server, "GET", "/excel/pending", headers=headers)[0] == 204
        invoke.join(2)
        assert not invoke.is_alive()
        assert result["status"] == 504
        assert result["body"] == _NOT_STARTED_BODY
        # A late acknowledgement and a late result are both refused.
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 409
        for executed in (False, True):
            assert _request(
                server, "POST", "/excel/result",
                {"ticketId": pending["ticketId"], "executed": executed, "result": {"ok": True}}, headers,
            )[0] == 409
        assert ticket.terminal_reason == "expired_claimed"
        assert ticket.result_payload is None
    finally:
        server.stop()


def test_plan_executing_deadline_is_measured_from_execute_start() -> None:
    from desk_plugin.companion import protocol

    clock_value = [4000.0]
    server = ExcelCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
        clock=lambda: clock_value[0], invoke_timeout_seconds=120,
    )
    server.start()
    try:
        session, headers, pending, ticket, invoke, result = _claimed_plan(
            server, "22222222-2222-4222-8222-222222222209",
        )
        clock_value[0] += 15  # e.g. confirmation inside the claim lease
        server.sessions.sweep()
        assert ticket.state == "claimed"
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
        clock_value[0] += 10  # 25 s after claim, 10 s after execute-start
        server.sessions.sweep()
        assert ticket.state == "executing"
        assert invoke.is_alive()
        lease = protocol.PLAN_EXECUTING_LEASE_SECONDS
        assert 10 < lease <= protocol.INVOKE_TIMEOUT_SECONDS - protocol.CLAIM_LEASE_SECONDS
        assert ticket.executing_at == 4015.0
        clock_value[0] = ticket.executing_at + lease
        server.sessions.sweep()
        assert ticket.state == "terminal"
        assert ticket.terminal_reason == "executing_uncertain"
        invoke.join(2)
        assert result["status"] == 504
        assert result["body"] == _UNCERTAIN_BODY
        # No new attempt is authorized after uncertainty.
        assert _request(server, "GET", "/excel/pending", headers=headers)[0] == 204
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 409
        assert _request(
            server, "POST", "/excel/result",
            {"ticketId": pending["ticketId"], "executed": True, "result": {"ok": True}}, headers,
        )[0] == 409
    finally:
        server.stop()


def test_plan_cancel_while_executing_is_uncertain_and_content_free() -> None:
    marker = "PLAN-UNCERTAIN-CONTENT-SENTINEL"
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=3)
    server.start()
    try:
        session, headers, pending, ticket, invoke, result = _claimed_plan(
            server, "22222222-2222-4222-8222-222222222210",
        )
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
        assert _request(
            server, "POST", "/excel/cancel",
            {"sessionId": session["sessionId"], "ticketId": ticket.ticket_id}, _plugin_headers(),
        )[0] == 200
        assert ticket.terminal_reason == "executing_uncertain"
        # A late pane report cannot convert uncertainty into completion.
        assert _request(
            server, "POST", "/excel/result",
            {"ticketId": pending["ticketId"], "executed": True, "result": {"cell": marker}}, headers,
        )[0] == 409
        invoke.join(2)
        assert result["status"] == 504
        assert result["body"] == _UNCERTAIN_BODY
        assert marker not in json.dumps(result["body"])
    finally:
        server.stop()


def test_plan_unclaimed_invoke_timeout_is_not_started() -> None:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=0.3)
    server.start()
    try:
        session = _session(server)
        turn = "22222222-2222-4222-8222-222222222211"
        trusted = {"turnOperationId": turn, "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key(turn)}
        invoke, result = _invoke_async(server, session, turn, trusted)
        _wait_for_ticket(server, session["sessionId"])
        ticket = server.sessions.get(session["sessionId"]).get_ticket()
        invoke.join(3)
        assert ticket.terminal_reason == "expired_unclaimed"
        assert result["status"] == 504
        assert result["body"] == _NOT_STARTED_BODY
    finally:
        server.stop()


def test_primitive_result_and_lease_requeue_are_unchanged() -> None:
    clock_value = [5000.0]
    server = ExcelCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
        clock=lambda: clock_value[0], invoke_timeout_seconds=120,
    )
    server.start()
    try:
        session = _session(server)
        request_id = "22222222-2222-4222-8222-222222222212"
        outcome: dict[str, object] = {}

        def run() -> None:
            status, body = _request(
                server, "POST", "/excel/invoke",
                {
                    "protocolVersion": "1.0", "sessionId": session["sessionId"],
                    "requestId": request_id, "tool": "read_range",
                    "arguments": {"sheet": "Sheet1", "address": "A1"},
                },
                _plugin_headers(),
            )
            outcome.update(status=status, body=body)

        thread = threading.Thread(target=run, daemon=True)
        thread.start()
        _wait_for_ticket(server, session["sessionId"])
        headers = _pane_headers(session)
        status, pending = _request(server, "GET", "/excel/pending", headers=headers)
        assert status == 200
        ticket = server.sessions.get(session["sessionId"]).get_ticket()
        clock_value[0] += CLAIM_LEASE_SECONDS + 1
        server.sessions.sweep()
        assert ticket.state == "pending"
        assert ticket.requeue_count == 1
        assert _request(server, "GET", "/excel/pending", headers=headers)[0] == 200
        assert _request(
            server, "POST", "/excel/result",
            {"ticketId": pending["ticketId"], "result": {"values": [[1]]}}, headers,
        )[0] == 200
        thread.join(2)
        assert outcome["status"] == 200
        assert outcome["body"] == {
            "ok": True,
            "ticketId": pending["ticketId"],
            "requestId": request_id,
            "data": {"values": [[1]]},
        }
    finally:
        server.stop()


# ---------------------------------------------------------------------------
# R3a: the companion is the single authority for plan-op acceptance (G2),
# restart safety rests on per-process session nonces (G6), the claim lease is
# rechecked at the execution boundary (G7/A1).
# ---------------------------------------------------------------------------

import ast as _r3a_ast
import time as _r3a_time
from pathlib import Path as _R3aPath
from unittest.mock import patch

import desk_plugin.companion.session as _r3a_session_module
from desk_plugin.companion.session import Session as _R3aSession
from desk_plugin.companion.session import Ticket as _R3aTicket

_R3A_CAPACITY_BODY = {
    "ok": False,
    "error": "Workbook plan capacity reached for this session",
    "category": "Capacity",
    "state": "plan_fence_capacity",
}
_R3A_OP = [f"{n:08x}-3a3a-4a3a-8a3a-3a3a3a3a3a{n:02x}" for n in range(0, 40)]
_R3A_RESULT = {"ok": True, "action": "workbook_plan"}


def _r3a_invoke_body(session: dict, op: str, plan: dict | None = None) -> dict:
    return {
        "protocolVersion": "1.0",
        "sessionId": session["sessionId"],
        "requestId": op,
        "tool": "execute_plan",
        "arguments": PLAN if plan is None else plan,
        "trusted": {"turnOperationId": op, "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key(op)},
    }


def _r3a_attempt(server: ExcelCompanionServer, session: dict, op: str,
                 plan: dict | None = None) -> tuple[int, dict, list[str]]:
    """Invoke one plan op. If the companion mints a ticket, the pane serves it
    to completion (executed:true). Returns (status, body, ticketIds served)."""
    session_obj = server.sessions.get(session["sessionId"])
    headers = _pane_headers(session)
    outcome: dict[str, object] = {}

    def run() -> None:
        outcome["response"] = _request(server, "POST", "/excel/invoke",
                                       _r3a_invoke_body(session, op, plan), _plugin_headers())

    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    served: list[str] = []
    deadline = _r3a_time.monotonic() + 2
    while thread.is_alive() and _r3a_time.monotonic() < deadline:
        if session_obj.get_pending_ticket() is not None:
            status, pending = _request(server, "GET", "/excel/pending", headers=headers)
            assert status == 200, status
            served.append(pending["ticketId"])
            assert _request(server, "POST", "/excel/execute-start",
                            {"ticketId": pending["ticketId"]}, headers)[0] == 200
            assert _request(server, "POST", "/excel/result",
                            {"ticketId": pending["ticketId"], "executed": True, "result": _R3A_RESULT},
                            headers)[0] == 200
            break
        _r3a_time.sleep(0.005)
    thread.join(3)
    assert not thread.is_alive(), "invoke did not finish"
    status, body = outcome["response"]  # type: ignore[misc]
    return status, body, served


def _r3a_primitive(server: ExcelCompanionServer, session: dict, request_id: str) -> None:
    outcome: dict[str, object] = {}

    def run() -> None:
        outcome["response"] = _request(server, "POST", "/excel/invoke", {
            "protocolVersion": "1.0", "sessionId": session["sessionId"], "requestId": request_id,
            "tool": "read_range", "arguments": {"sheet": "Sheet1", "address": "A1"},
        }, _plugin_headers())

    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    _wait_for_ticket(server, session["sessionId"])
    headers = _pane_headers(session)
    status, pending = _request(server, "GET", "/excel/pending", headers=headers)
    assert status == 200
    assert _request(server, "POST", "/excel/result",
                    {"ticketId": pending["ticketId"], "result": {"values": [[1]]}}, headers)[0] == 200
    thread.join(2)
    assert outcome["response"][0] == 200  # type: ignore[index]


def _r3a_server() -> ExcelCompanionServer:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=2)
    server.start()
    return server


def test_r3a_ra1_plan_fence_capacity_rejects_new_op_without_state_change() -> None:
    server = _r3a_server()
    try:
        session = _session(server)
        session_obj = server.sessions.get(session["sessionId"])
        session_obj._plan_turn_history_limit = 1
        status, first, served = _r3a_attempt(server, session, _R3A_OP[1])
        assert status == 200 and len(served) == 1
        fences_before = dict(session_obj._plan_replay_fences)
        history_before = list(session_obj._request_history)

        status, body, served = _r3a_attempt(server, session, _R3A_OP[2])
        assert (status, body) == (429, _R3A_CAPACITY_BODY), (status, body, served)
        assert served == []
        assert _request(server, "GET", "/excel/pending", headers=_pane_headers(session))[0] == 204
        assert list(session_obj._plan_turn_history) == [_R3A_OP[1]]
        assert dict(session_obj._plan_replay_fences) == fences_before
        assert list(session_obj._request_history) == history_before
        assert session_obj.get_ticket().ticket_id == first["ticketId"]
    finally:
        server.stop()


def test_r3a_ra2_capacity_never_evicts_so_redelivery_is_existing(monkeypatch) -> None:
    monkeypatch.setattr(_r3a_session_module, "MAX_PLAN_REPLAY_FENCES", 1)
    server = _r3a_server()
    try:
        session = _session(server)
        session_obj = server.sessions.get(session["sessionId"])
        session_obj._history_limit = 1
        session_obj._plan_turn_history_limit = 1
        headers = _pane_headers(session)
        status, first, served_first = _r3a_attempt(server, session, _R3A_OP[1])
        assert status == 200 and served_first == [first["ticketId"]]

        observed = []
        for op in (_R3A_OP[2], _R3A_OP[3]):
            status, body, served = _r3a_attempt(server, session, op)
            observed.append((status, body, served))
            assert _request(server, "GET", "/excel/pending", headers=headers)[0] == 204
        status, replay, served_replay = _r3a_attempt(server, session, _R3A_OP[1])
        observed.append((status, replay, served_replay))
        assert _request(server, "GET", "/excel/pending", headers=headers)[0] == 204

        assert observed[0][:2] == (429, _R3A_CAPACITY_BODY), observed
        assert observed[1][:2] == (429, _R3A_CAPACITY_BODY), observed
        # Exactly one ticket was ever created for op 1; the redelivery is the
        # cached original (existing), never a second pending ticket.
        assert served_replay == [], observed
        assert status == 200 and replay == first, observed
        assert session_obj._request_history[_R3A_OP[1]].ticket_id == first["ticketId"]
    finally:
        server.stop()


def test_r3a_ra3_capacity_check_runs_after_existing_conflict_and_replay_fences() -> None:
    server = _r3a_server()
    try:
        session = _session(server)
        session_obj = server.sessions.get(session["sessionId"])
        session_obj._plan_turn_history_limit = 1
        status, first, _ = _r3a_attempt(server, session, _R3A_OP[1])
        assert status == 200

        # Exact redelivery at capacity: cached result, not 429 (Q1 behaviour).
        status, body, served = _r3a_attempt(server, session, _R3A_OP[1])
        assert (status, body, served) == (200, first, [])

        # Changed plan under the same op: atomic 409, state unchanged.
        snapshot = (list(session_obj._request_history), list(session_obj._plan_turn_history),
                    dict(session_obj._plan_replay_fences), session_obj.get_ticket())
        changed = dict(PLAN, message="A different plan under the same operation")
        status, body, served = _r3a_attempt(server, session, _R3A_OP[1], changed)
        assert (status, body, served) == (409, {
            "ok": False, "error": "requestId reused with different payload", "category": "Conflict",
        }, [])
        assert snapshot == (list(session_obj._request_history), list(session_obj._plan_turn_history),
                            dict(session_obj._plan_replay_fences), session_obj.get_ticket())

        # A primitive pushes op 1 out of request history: replay fence, not 429.
        session_obj._history_limit = 1
        _r3a_primitive(server, session, "22222222-2222-4222-8222-2222222223a3")
        assert _R3A_OP[1] not in session_obj._request_history
        assert _R3A_OP[1] in session_obj._plan_replay_fences
        status, body, served = _r3a_attempt(server, session, _R3A_OP[1])
        assert (status, body, served) == (409, {
            "ok": False, "error": "Workbook plan replay identity is no longer executable", "category": "Conflict",
        }, [])

        # Only a genuinely new op is rejected for capacity.
        status, body, served = _r3a_attempt(server, session, _R3A_OP[2])
        assert (status, body, served) == (429, _R3A_CAPACITY_BODY, [])
    finally:
        server.stop()


def _r3a_unit_session() -> _R3aSession:
    return _R3aSession(session_id="s", document_id="doc", pane_id="pane", session_key="k",
                       workbook_plan_fingerprint=FINGERPRINT)


def _r3a_plan_ticket(request_id: str, turn: str | None, **kwargs) -> _R3aTicket:
    trusted = None if turn is None else {"turnOperationId": turn, "workbookPlanFingerprint": FINGERPRINT, "phase": "execute", "turnKey": _r4g2_turn_key(turn)}
    return _R3aTicket(ticket_id=f"t-{request_id}", session_id="s", request_id=request_id,
                      tool="excel_execute_plan", arguments=dict(PLAN), trusted=trusted, **kwargs)


def test_r3a_ra4_begin_ticket_refuses_plan_without_matching_trusted_turn() -> None:
    for turn in (_R3A_OP[5], None):
        session = _r3a_unit_session()
        disposition, _ = session.begin_ticket(_r3a_plan_ticket(_R3A_OP[4], turn))
        assert disposition == "conflict", (turn, disposition)
        assert session.get_ticket() is None
        assert not session._request_history
        assert not session._plan_turn_history
        assert not session._plan_replay_fences


def test_r3a_ra5_mark_executing_rechecks_claim_lease_under_lock() -> None:
    clock = [1000.0]
    ticket = _r3a_plan_ticket(_R3A_OP[6], _R3A_OP[6], _clock=lambda: clock[0])
    assert ticket.claim("pane") is True
    clock[0] = 1000.0 + CLAIM_LEASE_SECONDS
    assert ticket.mark_executing("pane") == "invalid:expired"
    assert ticket.state == "terminal"
    assert ticket.terminal_reason == "expired_claimed"
    assert ticket.executing_at == 0.0
    assert ticket.wait_for_result(0) is True
    assert ticket.mark_executing("pane") == "invalid:terminal"

    # Strictly inside the lease the boundary is still crossed exactly once.
    inside = _r3a_plan_ticket(_R3A_OP[7], _R3A_OP[7], _clock=lambda: clock[0])
    assert inside.claim("pane") is True
    clock[0] += CLAIM_LEASE_SECONDS - 0.001
    assert inside.mark_executing("pane") == "accepted"
    assert inside.state == "executing" and inside.executing_at == clock[0]
    assert inside.wait_for_result(0) is False


def test_r3a_ra6_execute_start_after_unswept_lease_expiry_is_refused() -> None:
    clock_value = [6000.0]
    server = ExcelCompanionServer(
        host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
        clock=lambda: clock_value[0], invoke_timeout_seconds=120,
    )
    server.sessions.sweep = lambda *_a, **_kw: []  # type: ignore[method-assign]
    server.start()
    try:
        session, headers, pending, ticket, invoke, result = _claimed_plan(server, _R3A_OP[8])
        clock_value[0] += CLAIM_LEASE_SECONDS
        status, body = _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)
        assert (status, body) == (409, {
            "ok": False, "error": "Ticket execution boundary is not available", "category": "Conflict",
        })
        assert ticket.state == "terminal" and ticket.terminal_reason == "expired_claimed"
        invoke.join(2)
        assert not invoke.is_alive()
        assert result["status"] == 504
        assert result["body"] == _NOT_STARTED_BODY
        assert _request(server, "POST", "/excel/result",
                        {"ticketId": pending["ticketId"], "executed": True, "result": {"ok": True}},
                        headers)[0] == 409
    finally:
        server.stop()


def test_r3a_gp2_restarted_companion_never_routes_an_old_session_or_op() -> None:
    from scripts.tests.test_office_plugin_plan_contract import load_plugin

    first = _r3a_server()
    try:
        session = _session(first)
        op = _R3A_OP[9]
        status, done, _ = _r3a_attempt(first, session, op)
        assert status == 200
    finally:
        first.stop()

    second = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=2)
    second.start()
    try:
        status, body = _request(second, "POST", "/excel/invoke", _r3a_invoke_body(session, op), _plugin_headers())
        assert (status, body) == (404, {"ok": False, "error": "Session not found", "category": "NotFound"})
        metrics = _request(second, "GET", "/excel/metrics", headers=_plugin_headers())[1]
        assert metrics["ticketsPending"] == 0 and metrics["sessionsActive"] == 0
        headers = _pane_headers(session)
        for method, path, payload in (
            ("GET", "/excel/pending", None),
            ("POST", "/excel/execute-start", {"ticketId": done["ticketId"]}),
            ("POST", "/excel/result", {"ticketId": done["ticketId"], "executed": True, "result": {}}),
            ("POST", "/excel/heartbeat", {}),
            ("DELETE", "/excel/session", None),
        ):
            assert _request(second, method, path, payload, headers)[0] == 401, (method, path)

        bridge_client = load_plugin().bridge_client
        calls: list[object] = []
        real_urlopen = bridge_client.urllib.request.urlopen

        def counting_urlopen(request, timeout=None):
            calls.append(request)
            return real_urlopen(request, timeout=timeout)

        host, port = second.server_address()
        outcome: object = None
        with (
            patch.object(bridge_client, "_connection", return_value=(f"http://{host}:{port}", _R3aPath("/never-read"))),
            patch.object(bridge_client, "_validate_key", return_value=PLUGIN_KEY),
            patch.object(bridge_client.urllib.request, "urlopen", counting_urlopen),
        ):
            try:
                bridge_client.invoke_bridge(session["sessionId"], "execute_plan", PLAN, trusted={
                    "turn_operation_id": op, "request_id": op, "workbook_plan_fingerprint": FINGERPRINT,
                    "phase": "execute", "turn_key": _r4g2_turn_key(op),
                })
            except bridge_client.BridgeError as exc:
                outcome = exc
        assert outcome.category == "NotFound" and str(outcome) == "Session not found"
        assert len(calls) == 1
        assert _request(second, "GET", "/excel/metrics", headers=_plugin_headers())[1]["ticketsPending"] == 0
    finally:
        second.stop()


def test_r3a_gp3_sessions_are_process_nonces_with_no_restore_path() -> None:
    root = _R3aPath(__file__).resolve().parents[2] / "plugins/atlas-sidecar-excel/companion"
    trees = {name: _r3a_ast.parse((root / f"{name}.py").read_text(encoding="utf-8"))
             for name in ("session", "handler", "server")}

    banned_modules = {"pickle", "shelve", "sqlite3", "dbm"}
    for node in _r3a_ast.walk(trees["session"]):
        if isinstance(node, _r3a_ast.Import):
            assert not {alias.name.split(".")[0] for alias in node.names} & banned_modules
        if isinstance(node, _r3a_ast.ImportFrom):
            assert (node.module or "").split(".")[0] not in banned_modules
        if isinstance(node, _r3a_ast.Call):
            func = node.func
            assert not (isinstance(func, _r3a_ast.Name) and func.id == "open")
            if isinstance(func, _r3a_ast.Attribute):
                assert func.attr not in ("write_text", "write_bytes", "open")
                assert not (func.attr == "dump" and isinstance(func.value, _r3a_ast.Name)
                            and func.value.id == "json")

    sites: list[tuple[str, str, object]] = []
    for module_name, tree in trees.items():
        for cls in [n for n in _r3a_ast.walk(tree) if isinstance(n, _r3a_ast.ClassDef)] + [tree]:
            body = cls.body
            for func in [n for n in body if isinstance(n, (_r3a_ast.FunctionDef, _r3a_ast.AsyncFunctionDef))]:
                for node in _r3a_ast.walk(func):
                    targets: list[object] = []
                    if isinstance(node, _r3a_ast.Assign):
                        targets = list(node.targets)
                    elif isinstance(node, (_r3a_ast.AugAssign, _r3a_ast.AnnAssign)):
                        targets = [node.target]
                    for target in targets:
                        if (isinstance(target, _r3a_ast.Subscript)
                                and isinstance(target.value, _r3a_ast.Attribute)
                                and target.value.attr == "_sessions"):
                            owner = cls.name if isinstance(cls, _r3a_ast.ClassDef) else "<module>"
                            sites.append((module_name, f"{owner}.{func.name}", target.slice))
                    if (isinstance(node, _r3a_ast.Call) and isinstance(node.func, _r3a_ast.Attribute)
                            and node.func.attr in ("update", "setdefault", "__setitem__")
                            and isinstance(node.func.value, _r3a_ast.Attribute)
                            and node.func.value.attr == "_sessions"):
                        sites.append((module_name, f"{func.name}:{node.func.attr}", None))
    assert len(sites) == 1, sites
    module_name, where, key = sites[0]
    assert (module_name, where) == ("session", "SessionManager.create")
    assert isinstance(key, _r3a_ast.Name)
    create = next(
        f for c in _r3a_ast.walk(trees["session"]) if isinstance(c, _r3a_ast.ClassDef) and c.name == "SessionManager"
        for f in c.body if isinstance(f, _r3a_ast.FunctionDef) and f.name == "create"
    )
    bindings = [
        n for n in _r3a_ast.walk(create)
        if isinstance(n, _r3a_ast.Assign) and any(isinstance(t, _r3a_ast.Name) and t.id == key.id for t in n.targets)
    ]
    assert len(bindings) == 1
    value = bindings[0].value
    assert isinstance(value, _r3a_ast.Call) and isinstance(value.func, _r3a_ast.Name)
    assert value.func.id == "generate_uuid" and not value.args

    constructions = [
        n for n in _r3a_ast.walk(trees["server"])
        if isinstance(n, _r3a_ast.Call) and isinstance(n.func, _r3a_ast.Name) and n.func.id == "SessionManager"
    ]
    assert len(constructions) == 1


def test_r3a_gp4_plan_fences_are_bounded_retained_and_primitives_still_work() -> None:
    limit = 3
    session = _r3a_unit_session()
    session._history_limit = 1
    session._plan_turn_history_limit = limit
    admitted: list[str] = []
    last_created: list[str] = []
    # P = new or repeated plan op (index), R = primitive (index)
    sequence = [("P", 1), ("R", 1), ("P", 2), ("P", 1), ("P", 2), ("R", 2), ("P", 3),
                ("P", 4), ("R", 3), ("P", 5), ("P", 3), ("P", 4), ("R", 4), ("P", 1),
                ("P", 6), ("P", 2), ("R", 5), ("P", 3), ("P", 7), ("R", 6)]
    assert len(sequence) == 20
    for step, (kind, n) in enumerate(sequence):
        if kind == "R":
            ticket = _R3aTicket(ticket_id=f"r{n}", session_id="s", request_id=f"prim-{n}",
                                tool="read_range", arguments={"sheet": "Sheet1", "address": f"A{n}"})
            expected = "created"
        else:
            op = _R3A_OP[20 + n]
            ticket = _r3a_plan_ticket(op, op)
            if op in admitted:
                expected = "existing" if last_created and last_created[-1] == op else "replay_evicted"
            else:
                expected = "created" if len(admitted) < limit else "plan_fence_capacity"
        disposition, returned = session.begin_ticket(ticket)
        assert disposition == expected, (step, kind, n, disposition, expected)
        if disposition == "created":
            last_created.append(ticket.request_id)
            if ticket.is_plan:
                admitted.append(ticket.request_id)
            returned.set_result({"ok": True})
            session.trim_history()
        fences = set(session._plan_replay_fences)
        turns = set(session._plan_turn_history)
        assert fences <= turns, step
        assert len(turns) <= limit, step
        assert set(admitted) <= set(session._request_history) | fences, step
    assert len(admitted) == limit


def test_r3a_gp5_concurrent_same_op_creates_exactly_one_ticket() -> None:
    server = _r3a_server()
    try:
        session = _session(server)
        session_obj = server.sessions.get(session["sessionId"])
        dispositions: list[str] = []
        original = session_obj.begin_ticket

        def recording(ticket):
            outcome = original(ticket)
            dispositions.append(outcome[0])
            return outcome

        session_obj.begin_ticket = recording  # type: ignore[method-assign]
        barrier = threading.Barrier(2)
        results: list[tuple[int, dict]] = []

        def run() -> None:
            barrier.wait(2)
            results.append(_request(server, "POST", "/excel/invoke",
                                    _r3a_invoke_body(session, _R3A_OP[10]), _plugin_headers()))

        threads = [threading.Thread(target=run, daemon=True) for _ in range(2)]
        for thread in threads:
            thread.start()
        _wait_for_ticket(server, session["sessionId"])
        deadline = _r3a_time.monotonic() + 2
        while len(dispositions) < 2 and _r3a_time.monotonic() < deadline:
            _r3a_time.sleep(0.005)
        headers = _pane_headers(session)
        status, pending = _request(server, "GET", "/excel/pending", headers=headers)
        assert status == 200
        assert _request(server, "GET", "/excel/pending", headers=headers)[0] == 204
        assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
        assert _request(server, "POST", "/excel/result",
                        {"ticketId": pending["ticketId"], "executed": True, "result": _R3A_RESULT}, headers)[0] == 200
        for thread in threads:
            thread.join(3)
            assert not thread.is_alive()
        assert sorted(dispositions) == ["created", "existing"]
        assert [status for status, _ in results] == [200, 200]
        assert {body["ticketId"] for _, body in results} == {pending["ticketId"]}
        assert list(session_obj._request_history) == [_R3A_OP[10]]
        assert _request(server, "GET", "/excel/pending", headers=headers)[0] == 204
    finally:
        server.stop()


# ---------------------------------------------------------------------------
# R3d-T (Python half, PX1-PX9): cross-layer restart/replay matrix. The real
# registered plugin (pre-hook + handler) -> the real bridge_client -> a real
# loopback ExcelCompanionServer -> a Python pane driver that mutates only after
# an accepted execute-start. "Mutation" = POST /excel/execute-start -> 200.
# Contract helpers are imported inside each test (the contract module is
# import-only); expected model views come from the frozen wire fixture.
# ---------------------------------------------------------------------------

import uuid as _r3d_uuid
from contextlib import nullcontext as _r3d_nullcontext

_R3D_FIXTURE = (_R3aPath(__file__).resolve().parents[2]
                / "src" / "__tests__" / "fixtures" / "companion-wire" / "plan-wire-contract.v2.json")
_R3D_FIXTURE_SHA256 = "4b27a29446334173ccae49911ebf2a29f04ebd69748db18af975ed715ae8c14d"
_R3D_NOT_FOUND = {"ok": False, "error": "Session not found", "category": "NotFound"}
_R3D_RESULT = {"ok": True, "action": "workbook_plan", "result": {"outcome": "committed", "note": "r3d"}}
_R3D_S_TURN = "r3d-hermes-turn-1"
_R3D_S_TURN_2 = "r3d-hermes-turn-2"
_R3D_ROUTE: dict = {}


def valid_execute_plan() -> dict:
    """R4g-2a rule M'': a schema-valid execute plan (the wire fixture's C02 plan) for the execute-path r3d tests."""
    return {
        "version": 1,
        "outcome": "execute",
        "summary": "Write one cell",
        "operations": [{"kind": "write_values", "sheet": "Sheet1", "address": "A1", "values": [["ok"]]}],
        "verification": [{"kind": "range_values_equal", "sheet": "Sheet1", "address": "A1", "expectedValues": [["ok"]]}],
    }


def _r3d_fixture_model_view(case_id: str) -> dict:
    import hashlib
    raw = _R3D_FIXTURE.read_bytes()
    assert hashlib.sha256(raw).hexdigest() == _R3D_FIXTURE_SHA256
    case = next(c for c in json.loads(raw)["cases"] if c["id"] == case_id)
    end = next(s for s in case["steps"] if s["kind"] == "plugin_invoke_end")
    return end["pluginOutcome"]["modelView"]


class _R3dCompanion:
    """A real loopback companion (optionally on an injected clock) plus a counting pane driver."""

    def __init__(self, fingerprint: str, clock=None, invoke_timeout: float = 3) -> None:
        kwargs = {} if clock is None else {"clock": clock}
        self.fingerprint = fingerprint
        self.server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY,
                                           invoke_timeout_seconds=invoke_timeout, **kwargs)
        self.acks = 0
        self.ack_statuses: list[int] = []

    def start(self) -> "_R3dCompanion":
        self.server.start()
        self.session = _session(self.server, self.fingerprint)
        self.session_obj = self.server.sessions.get(self.session["sessionId"])
        self.headers = _pane_headers(self.session)
        return self

    def new_session(self) -> dict:
        return _session(self.server, self.fingerprint)

    def stop(self) -> None:
        if not getattr(self, "_stopped", False):
            self._stopped = True
            self.server.stop()

    def patches(self, bridge_client, calls: list):
        host, port = self.server.server_address()
        real_urlopen = bridge_client.urllib.request.urlopen

        def counting_urlopen(request, timeout=None):
            if getattr(request, "full_url", "").endswith("/excel/invoke"):
                calls.append(request)
            return real_urlopen(request, timeout=timeout)

        # Every fresh registered() module loads its own bridge_client; each one
        # routes through _R3D_ROUTE (set here), and urllib is the shared stdlib.
        return (
            patch.dict(_R3D_ROUTE, {"url": f"http://{host}:{port}"}),
            _r3d_nullcontext(),
            patch.object(bridge_client.urllib.request, "urlopen", counting_urlopen),
        )

    def wait_pending(self, session_obj=None) -> None:
        target = session_obj or self.session_obj
        deadline = _r3a_time.monotonic() + 3
        while target.get_pending_ticket() is None and _r3a_time.monotonic() < deadline:
            _r3a_time.sleep(0.005)
        assert target.get_pending_ticket() is not None, "no pending ticket"

    def claim(self, headers=None, session_obj=None) -> str:
        self.wait_pending(session_obj)
        status, pending = _request(self.server, "GET", "/excel/pending", headers=headers or self.headers)
        assert status == 200, status
        return pending["ticketId"]

    def ack(self, ticket_id: str, headers=None) -> int:
        status = _request(self.server, "POST", "/excel/execute-start", {"ticketId": ticket_id}, headers or self.headers)[0]
        self.ack_statuses.append(status)
        if status == 200:
            self.acks += 1
        return status

    def report(self, ticket_id: str, result=None, headers=None) -> int:
        return _request(self.server, "POST", "/excel/result",
                        {"ticketId": ticket_id, "executed": True, "result": _R3D_RESULT if result is None else result},
                        headers or self.headers)[0]

    def serve(self, headers=None, session_obj=None, result=None) -> str:
        """Pane: claim, cross execute-start (the only mutation), report executed:true."""
        ticket_id = self.claim(headers, session_obj)
        assert self.ack(ticket_id, headers) == 200
        assert self.report(ticket_id, result, headers) == 200
        return ticket_id

    def pending_status(self, headers=None) -> int:
        return _request(self.server, "GET", "/excel/pending", headers=headers or self.headers)[0]

    def metrics(self) -> dict:
        return _request(self.server, "GET", "/excel/metrics", headers={"X-Hermes-Plugin-Key": PLUGIN_KEY})[1]


def _r3d_plugin(contract):
    """A fresh registered() plugin module: its advisory _seen_plan_turns starts empty."""
    module, context = contract.registered()
    # The module is disposable, so its own bridge_client is pointed at the
    # companion currently routed by _R3dCompanion.patches (never at real config).
    module.bridge_client._connection = lambda *_a, **_k: (_R3D_ROUTE["url"], _R3aPath("/never-read"))
    module.bridge_client._validate_key = lambda *_a, **_k: PLUGIN_KEY
    hook = context.hooks["pre_tool_call"]
    handler = next(tool["handler"] for tool in context.tools if tool["name"] == contract.PLAN_TOOL)
    return module, hook, handler


def _r3d_start(contract, plugin, session_id: str, turn: str, plan: dict, call_id: str = "call-r3d"):
    """Runs pre-hook + handler in one copied context on a worker thread (Hermes' shape)."""
    _module, hook, handler = plugin

    def run():
        blocked = hook(tool_name=contract.PLAN_TOOL, args=plan, session_id=session_id, turn_id=turn,
                       tool_call_id=call_id, api_request_id=f"request-{call_id}", task_id="task-r3d")
        if blocked is not None:
            return ("blocked", blocked)
        return ("result", handler(plan, session_id=session_id))

    return contract._r3a_in_thread(run)


def _r3d_op(session_id: str, turn: str, phase: str = "execute") -> str:
    return str(_r3d_uuid.uuid5(_r3d_uuid.UUID("5f8f4f1c-4c57-4a35-8b4f-c0b2aa3f9248"), f"{session_id}\\0{turn}\\0{phase}"))


def _r3d_px1_setup(contract, companion: _R3dCompanion, calls: list):
    """Instance 1 runs (S, H1) to a served, committed result R."""
    plugin = _r3d_plugin(contract)
    sid = companion.session["sessionId"]
    thread, box = _r3d_start(contract, plugin, sid, _R3D_S_TURN, valid_execute_plan())
    ticket_id = companion.serve()
    kind, raw = contract._r3a_join(thread, box)
    assert kind == "result"
    assert json.loads(raw) == _R3D_RESULT
    assert len(calls) == 1
    return ticket_id, raw


def test_r3d_px1_new_plugin_instance_replays_the_cached_terminal_meaning() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    companion = _R3dCompanion(contract.generated_fingerprint()).start()
    try:
        calls: list = []
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            ticket_id, first = _r3d_px1_setup(contract, companion, calls)
            fresh = _r3d_plugin(contract)
            assert not fresh[0]._seen_plan_turns
            thread, box = _r3d_start(contract, fresh, companion.session["sessionId"], _R3D_S_TURN,
                                     valid_execute_plan(), "call-r3d-px1-fresh")
            kind, replay = contract._r3a_join(thread, box)
        assert kind == "result"
        assert replay == first  # byte-equal cached terminal meaning
        op = _r3d_op(companion.session["sessionId"], _R3D_S_TURN)
        assert list(companion.session_obj._request_history) == [op]
        assert companion.session_obj.get_ticket().ticket_id == ticket_id
        assert companion.pending_status() == 204
        assert len(calls) == 2
        assert companion.acks == 1
    finally:
        companion.stop()


def test_r3d_px2_restart_new_turn_identical_plan_is_an_independent_later_write() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    companion = _R3dCompanion(contract.generated_fingerprint()).start()
    try:
        calls: list = []
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            first_ticket, first = _r3d_px1_setup(contract, companion, calls)
            sid = companion.session["sessionId"]
            second_result = {"ok": True, "action": "workbook_plan", "result": {"outcome": "committed", "note": "h2"}}
            thread, box = _r3d_start(contract, _r3d_plugin(contract), sid, _R3D_S_TURN_2,
                                     valid_execute_plan(), "call-r3d-px2")
            second_ticket = companion.serve(result=second_result)
            kind, second = contract._r3a_join(thread, box)
        assert kind == "result"
        op1, op2 = _r3d_op(sid, _R3D_S_TURN), _r3d_op(sid, _R3D_S_TURN_2)
        assert op1 != op2
        assert second_ticket != first_ticket
        assert list(companion.session_obj._request_history) == [op1, op2]
        assert json.loads(first) == _R3D_RESULT and json.loads(second) == second_result
        assert companion.acks == 2  # one per op
        assert companion.pending_status() == 204
    finally:
        companion.stop()


def _r3d_restart_case(executing: bool) -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    fingerprint = contract.generated_fingerprint()
    bridge_client = contract.load_plugin().bridge_client
    first = _R3dCompanion(fingerprint).start()
    sid = first.session["sessionId"]
    calls: list = []
    try:
        p1, p2, p3 = first.patches(bridge_client, calls)
        with p1, p2, p3:
            plan = valid_execute_plan() if executing else contract.valid_plan()
            thread, box = _r3d_start(contract, _r3d_plugin(contract), sid, _R3D_S_TURN, plan)
            ticket_id = first.claim()
            if executing:
                assert first.ack(ticket_id) == 200
            first.stop()  # companion restart: every session and ticket is dropped
            kind, raw = contract._r3a_join(thread, box)
    finally:
        first.stop()
    assert kind == "result"
    observed = json.loads(raw)
    category = observed["error"]["category"]
    assert observed.get("ok") is False and category != "Internal", observed
    if executing:
        assert observed == _r3d_fixture_model_view("C10-cancel-executing")
    else:
        # OQ6: recorded at the base: a claimed ticket cancelled by shutdown is Cancellation (C09 view).
        assert observed == _r3d_fixture_model_view("C09-cancel-claimed")
    assert first.acks == (1 if executing else 0)

    second = _R3dCompanion(fingerprint)
    second.server.start()
    try:
        replay_calls: list = []
        p1, p2, p3 = second.patches(bridge_client, replay_calls)
        with p1, p2, p3:
            thread, box = _r3d_start(contract, _r3d_plugin(contract), sid, _R3D_S_TURN,
                                     plan, "call-r3d-replay")
            kind, raw = contract._r3a_join(thread, box)
        assert kind == "result"
        assert json.loads(raw) == {"ok": False, "error": {"category": "NotFound", "message": "Session not found"}}
        assert len(replay_calls) == 1
        assert second.metrics()["ticketsPending"] == 0
        assert first.acks + second.acks == (1 if executing else 0)
    finally:
        second.stop()


def test_r3d_px3a_companion_restart_while_claimed_never_executes_and_old_session_is_gone() -> None:
    _r3d_restart_case(executing=False)


def test_r3d_px3b_companion_restart_while_executing_is_uncertain_and_never_retried() -> None:
    _r3d_restart_case(executing=True)


def _r3d_rotation_case(executing: bool) -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    companion = _R3dCompanion(contract.generated_fingerprint()).start()
    try:
        calls: list = []
        sid = companion.session["sessionId"]
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            thread, box = _r3d_start(contract, _r3d_plugin(contract), sid, _R3D_S_TURN, valid_execute_plan())
            ticket_id = companion.claim()
            if executing:
                assert companion.ack(ticket_id) == 200
            # The R3b rotation path: the pane DELETEs its session.
            assert _request(companion.server, "DELETE", "/excel/session", headers=companion.headers)[0] == 204
            kind, raw = contract._r3a_join(thread, box)
            assert kind == "result"
            assert json.loads(raw) == _r3d_fixture_model_view(
                "C10-cancel-executing" if executing else "C09-cancel-claimed")
            # A stale re-invoke for (S, H1) through a fresh plugin module.
            thread, box = _r3d_start(contract, _r3d_plugin(contract), sid, _R3D_S_TURN,
                                     valid_execute_plan(), "call-r3d-stale")
            kind, stale = contract._r3a_join(thread, box)
            assert kind == "result"
            assert json.loads(stale) == {"ok": False, "error": {"category": "NotFound", "message": "Session not found"}}
            assert companion.acks == (1 if executing else 0)
            if not executing:
                # A new session S2 with a new turn executes once.
                s2 = companion.new_session()
                s2_headers = _pane_headers(s2)
                s2_obj = companion.server.sessions.get(s2["sessionId"])
                thread, box = _r3d_start(contract, _r3d_plugin(contract), s2["sessionId"], _R3D_S_TURN_2,
                                         valid_execute_plan(), "call-r3d-s2")
                companion.serve(headers=s2_headers, session_obj=s2_obj)
                kind, raw = contract._r3a_join(thread, box)
                assert kind == "result" and json.loads(raw) == _R3D_RESULT
                assert companion.acks == 1  # S: 0, S2: 1
        assert len(calls) == (2 if executing else 3)
    finally:
        companion.stop()


def test_r3d_px4a_rotation_while_claimed_cancels_and_a_new_session_executes_once() -> None:
    _r3d_rotation_case(executing=False)


def test_r3d_px4b_rotation_while_executing_is_uncertain_and_stale_reinvoke_is_404() -> None:
    _r3d_rotation_case(executing=True)


def test_r3d_px5a_concurrent_duplicate_same_op_from_fresh_modules_creates_one_ticket() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    companion = _R3dCompanion(contract.generated_fingerprint()).start()
    try:
        dispositions: list[str] = []
        original = companion.session_obj.begin_ticket

        def recording(ticket):
            outcome = original(ticket)
            dispositions.append(outcome[0])
            return outcome

        companion.session_obj.begin_ticket = recording  # type: ignore[method-assign]
        calls: list = []
        sid = companion.session["sessionId"]
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            barrier = threading.Barrier(2)
            plugins = [_r3d_plugin(contract), _r3d_plugin(contract)]
            started = []
            for index, plugin in enumerate(plugins):
                _module, hook, handler = plugin

                def run(hook=hook, handler=handler, index=index):
                    call_id = f"call-r3d-px5a-{index}"
                    assert hook(tool_name=contract.PLAN_TOOL, args=valid_execute_plan(), session_id=sid,
                                turn_id=_R3D_S_TURN, tool_call_id=call_id, api_request_id=f"request-{call_id}",
                                task_id="task-r3d") is None
                    barrier.wait(5)
                    return handler(valid_execute_plan(), session_id=sid)

                started.append(contract._r3a_in_thread(run))
            deadline = _r3a_time.monotonic() + 3
            while len(dispositions) < 2 and _r3a_time.monotonic() < deadline:
                _r3a_time.sleep(0.005)
            ticket_id = companion.serve()
            raws = [contract._r3a_join(thread, box) for thread, box in started]
        assert sorted(dispositions) == ["created", "existing"]
        assert raws[0] == raws[1] and json.loads(raws[0]) == _R3D_RESULT
        assert list(companion.session_obj._request_history) == [_r3d_op(sid, _R3D_S_TURN)]
        assert companion.session_obj.get_ticket().ticket_id == ticket_id
        assert companion.pending_status() == 204
        assert len(calls) == 2
        assert companion.acks == 1
    finally:
        companion.stop()


def test_r3d_px5b_same_turn_second_tool_call_is_blocked_by_the_pre_hook() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    companion = _R3dCompanion(contract.generated_fingerprint()).start()
    try:
        calls: list = []
        sid = companion.session["sessionId"]
        plugin = _r3d_plugin(contract)
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            thread, box = _r3d_start(contract, plugin, sid, _R3D_S_TURN, valid_execute_plan(), "call-r3d-a")
            companion.serve()
            kind, raw = contract._r3a_join(thread, box)
            assert kind == "result" and json.loads(raw) == _R3D_RESULT
            thread, box = _r3d_start(contract, plugin, sid, _R3D_S_TURN, valid_execute_plan(), "call-r3d-b")
            kind, blocked = contract._r3a_join(thread, box)
        assert (kind, blocked) == ("blocked", {"action": "block", "message": "Workbook plan trusted context rejected."})
        assert len(calls) == 1
        assert companion.acks == 1
        assert companion.pending_status() == 204
    finally:
        companion.stop()


def test_r3d_px6_capacity_refuses_a_new_op_and_replays_the_admitted_one_byte_equal() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    companion = _R3dCompanion(contract.generated_fingerprint()).start()
    try:
        companion.session_obj._plan_turn_history_limit = 1
        calls: list = []
        sid = companion.session["sessionId"]
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            _ticket_id, first = _r3d_px1_setup(contract, companion, calls)
            fences = dict(companion.session_obj._plan_replay_fences)
            turns = list(companion.session_obj._plan_turn_history)
            thread, box = _r3d_start(contract, _r3d_plugin(contract), sid, _R3D_S_TURN_2,
                                     valid_execute_plan(), "call-r3d-px6-new")
            kind, refused = contract._r3a_join(thread, box)
            assert kind == "result"
            assert json.loads(refused) == {
                "ok": False, "error": {"category": "Capacity", "message": contract._R3A_CAPACITY_MESSAGE},
            }
            assert companion.pending_status() == 204
            thread, box = _r3d_start(contract, _r3d_plugin(contract), sid, _R3D_S_TURN,
                                     valid_execute_plan(), "call-r3d-px6-replay")
            kind, replay = contract._r3a_join(thread, box)
        assert kind == "result" and replay == first
        assert dict(companion.session_obj._plan_replay_fences) == fences
        assert list(companion.session_obj._plan_turn_history) == turns == [_r3d_op(sid, _R3D_S_TURN)]
        assert companion.acks == 1
        assert len(calls) == 3
    finally:
        companion.stop()


def _r3d_clocked(contract, claim: bool) -> None:
    clock = [7000.0]
    companion = _R3dCompanion(contract.generated_fingerprint(), clock=lambda: clock[0], invoke_timeout=30).start()
    try:
        calls: list = []
        sid = companion.session["sessionId"]
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            thread, box = _r3d_start(contract, _r3d_plugin(contract), sid, _R3D_S_TURN, contract.valid_plan())
            if claim:
                ticket_id = companion.claim()
                clock[0] += CLAIM_LEASE_SECONDS  # C11: lease ends; the late ack is refused
                assert companion.ack(ticket_id) == 409
                assert companion.session_obj.get_ticket().terminal_reason == "expired_claimed"
            else:
                companion.wait_pending()
                clock[0] += 30  # C13 pattern: past the unclaimed invoke window (session TTL 60 s not reached)
                assert companion.pending_status() == 204  # the request sweeps the ticket
                assert companion.session_obj.get_ticket().terminal_reason == "expired_unclaimed"
            kind, raw = contract._r3a_join(thread, box)
            assert kind == "result"
            expected = _r3d_fixture_model_view("C11-claimed-lease-expiry" if claim else "C13-unclaimed-expiry")
            assert json.loads(raw) == expected
            ticket_id = companion.session_obj.get_ticket().ticket_id
            thread, box = _r3d_start(contract, _r3d_plugin(contract), sid, _R3D_S_TURN,
                                     contract.valid_plan(), "call-r3d-replay")
            kind, replay = contract._r3a_join(thread, box)
        assert kind == "result" and replay == raw  # same terminal meaning, byte-equal
        assert companion.session_obj.get_ticket().ticket_id == ticket_id  # no new ticket
        assert list(companion.session_obj._request_history) == [_r3d_op(sid, _R3D_S_TURN, "inspect")]
        assert companion.pending_status() == 204
        assert companion.acks == 0
        assert len(calls) == 2
    finally:
        companion.stop()


def test_r3d_px7a_unclaimed_expiry_is_not_started_and_replay_keeps_that_meaning() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    _r3d_clocked(contract, claim=False)


def test_r3d_px7b_claimed_lease_expiry_refuses_the_ack_and_replay_keeps_that_meaning() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    _r3d_clocked(contract, claim=True)


def test_r3d_px8_changed_digest_from_a_fresh_module_is_an_atomic_conflict() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    companion = _R3dCompanion(contract.generated_fingerprint()).start()
    try:
        calls: list = []
        sid = companion.session["sessionId"]
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            ticket_id, _first = _r3d_px1_setup(contract, companion, calls)
            history = list(companion.session_obj._request_history)
            changed = dict(valid_execute_plan(), summary="A different plan for the same trusted turn (R3d)")
            thread, box = _r3d_start(contract, _r3d_plugin(contract), sid, _R3D_S_TURN, changed, "call-r3d-px8")
            kind, raw = contract._r3a_join(thread, box)
        assert kind == "result"
        assert json.loads(raw) == {
            "ok": False, "error": {"category": "Conflict", "message": "requestId reused with different payload"},
        }
        assert list(companion.session_obj._request_history) == history
        assert companion.session_obj.get_ticket().ticket_id == ticket_id
        assert companion.pending_status() == 204
        assert companion.acks == 1
    finally:
        companion.stop()


def test_r3d_px9_trust_anchor_operation_id_binds_the_session() -> None:
    from contextvars import copy_context as _r3d_copy_context

    from scripts.tests import test_office_plugin_plan_contract as contract

    s1 = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"
    s2 = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb"

    def derive(session_id: str, turn: str) -> str:
        module, hook, _handler = _r3d_plugin(contract)

        def run() -> str:
            assert hook(tool_name=contract.PLAN_TOOL, args=contract.valid_plan(), session_id=session_id,
                        turn_id=turn, tool_call_id="call-r3d-px9", api_request_id="request-r3d-px9",
                        task_id="task-r3d") is None
            return module._trusted_plan_context.get().peek()["operation_id"]

        return _r3d_copy_context().run(run)

    a = derive(s1, _R3D_S_TURN)
    b = derive(s2, _R3D_S_TURN)
    again = derive(s1, _R3D_S_TURN)
    assert a != b
    assert a == again == _r3d_op(s1, _R3D_S_TURN, "inspect")
    assert b == _r3d_op(s2, _R3D_S_TURN, "inspect")


# ---------------------------------------------------------------------------
# R4g-2a (dev3-r4g2-release.md §5.1, F9, F10; R4g-1 lane-B B-A3/B-A4): the companion's authoritative
# per-turnKey phase fence over the REAL loopback server (C1-C5, cross-turn row, size cap) and one
# white-box Ticket row (C1b). Plans/ids are test-local; nothing here edits a shared helper.
# ---------------------------------------------------------------------------
from desk_plugin.companion import protocol as _r4g2_protocol  # noqa: E402
from desk_plugin.companion.session import Ticket as _R4g2Ticket  # noqa: E402

_R4G2_INSPECT = {"version": 1, "outcome": "inspect", "reason": "Read one cell",
                 "reads": [{"kind": "read_range", "sheet": "Sheet1", "address": "A1"}]}
_R4G2_EXECUTE = {"version": 1, "outcome": "execute", "summary": "Write one cell",
                 "operations": [{"kind": "write_values", "sheet": "Sheet1", "address": "A1", "values": [["ok"]]}],
                 "verification": [{"kind": "range_values_equal", "sheet": "Sheet1", "address": "A1", "expectedValues": [["ok"]]}]}
_R4G2_INSPECTION = {"ok": True, "action": "workbook_plan", "phase": "inspect",
                    "result": {"outcome": "inspected", "inspectionCount": 1, "reads": [{"ok": True, "data": [["ok"]]}]}}
_R4G2_COMMITTED = {"ok": True, "action": "workbook_plan", "result": {"outcome": "committed"}}
_R4G2_OP = [f"{n:08x}-4a4a-4a4a-8a4a-4a4a4a4a4a{n:02x}" for n in range(0, 16)]
_R4G2_KEY = [f"{n:08x}-4b4b-4b4b-8b4b-4b4b4b4b4b{n:02x}" for n in range(0, 16)]
_R4G2_DUPLICATE = {"ok": False, "error": "A second workbook plan call for this trusted turn was rejected", "category": "Conflict"}
_R4G2_REQUIRED = {"ok": False, "error": "Trusted workbook plan metadata is required", "category": "Validation"}
_R4G2_INVALID = {"ok": False, "error": "Trusted workbook plan metadata is invalid", "category": "Validation"}


def _r4g2_trusted(op: str, phase: str, key: str) -> dict:
    return {"turnOperationId": op, "workbookPlanFingerprint": FINGERPRINT, "phase": phase, "turnKey": key}


def _r4g2_body(session: dict, op: str, plan: dict, trusted: dict | None) -> dict:
    body = {"protocolVersion": "1.0", "sessionId": session["sessionId"], "requestId": op, "tool": "execute_plan",
            "arguments": plan}
    if trusted is not None:
        body["trusted"] = trusted
    return body


def _r4g2_invoke(server: ExcelCompanionServer, session: dict, op: str, plan: dict, trusted: dict | None):
    box: dict = {}

    def run() -> None:
        box["status"], box["body"] = _request(server, "POST", "/excel/invoke", _r4g2_body(session, op, plan, trusted),
                                              {"X-Hermes-Plugin-Key": PLUGIN_KEY})

    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    return thread, box


def _r4g2_join(thread: threading.Thread, box: dict) -> tuple[int, dict]:
    thread.join(5)
    assert not thread.is_alive()
    return box["status"], box["body"]


def _r4g2_claim(server: ExcelCompanionServer, session: dict) -> dict:
    _wait_for_ticket(server, session["sessionId"])
    session_obj = server.sessions.get(session["sessionId"])
    import time as _time
    deadline = _time.monotonic() + 3
    while session_obj.get_pending_ticket() is None and _time.monotonic() < deadline:
        _time.sleep(0.005)
    status, pending = _request(server, "GET", "/excel/pending", headers=_pane_headers(session))
    assert status == 200, (status, pending)
    return pending


def _r4g2_inspect_done(server: ExcelCompanionServer, session: dict, op: str, key: str) -> dict:
    """Admit an inspect op and answer it executed:false with the inspection body."""
    thread, box = _r4g2_invoke(server, session, op, _R4G2_INSPECT, _r4g2_trusted(op, "inspect", key))
    pending = _r4g2_claim(server, session)
    assert pending["trusted"] == _r4g2_trusted(op, "inspect", key)
    assert list(pending["trusted"]) == ["turnOperationId", "workbookPlanFingerprint", "phase", "turnKey"]
    assert _request(server, "POST", "/excel/result", {"ticketId": pending["ticketId"], "executed": False,
                                                      "result": _R4G2_INSPECTION}, _pane_headers(session))[0] == 200
    status, body = _r4g2_join(thread, box)
    assert status == 200 and body["executed"] is False and body["data"] == _R4G2_INSPECTION, (status, body)
    return pending


def _r4g2_execute_done(server: ExcelCompanionServer, session: dict, op: str, key: str) -> dict:
    thread, box = _r4g2_invoke(server, session, op, _R4G2_EXECUTE, _r4g2_trusted(op, "execute", key))
    pending = _r4g2_claim(server, session)
    headers = _pane_headers(session)
    assert _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)[0] == 200
    assert _request(server, "POST", "/excel/result", {"ticketId": pending["ticketId"], "executed": True,
                                                      "result": _R4G2_COMMITTED}, headers)[0] == 200
    status, body = _r4g2_join(thread, box)
    assert status == 200 and body["executed"] is True, (status, body)
    return pending


def _r4g2_refused(server: ExcelCompanionServer, session: dict, op: str, plan: dict, trusted: dict | None) -> tuple[int, dict]:
    status, body = _request(server, "POST", "/excel/invoke", _r4g2_body(session, op, plan, trusted),
                            {"X-Hermes-Plugin-Key": PLUGIN_KEY})
    assert _request(server, "GET", "/excel/pending", headers=_pane_headers(session))[0] == 204
    return status, body


def _r4g2_state(session_obj) -> tuple:
    return (list(session_obj._request_history), list(session_obj._plan_turn_history),
            dict(session_obj._plan_replay_fences), {k: sorted(v) for k, v in session_obj._plan_turn_phases.items()})


def _r4g2_server(**kwargs) -> ExcelCompanionServer:
    server = ExcelCompanionServer(host="127.0.0.1", port=0, plugin_key=PLUGIN_KEY, invoke_timeout_seconds=3, **kwargs)
    server.start()
    return server


def test_r4g2_c1_inspect_then_execute_same_turnkey_inspect_ticket_never_crosses_execution() -> None:
    server = _r4g2_server()
    try:
        session = _session(server)
        headers = _pane_headers(session)
        op_i, op_e, key = _R4G2_OP[1], _R4G2_OP[2], _R4G2_KEY[1]
        thread, box = _r4g2_invoke(server, session, op_i, _R4G2_INSPECT, _r4g2_trusted(op_i, "inspect", key))
        pending = _r4g2_claim(server, session)
        ticket = server.sessions.get(session["sessionId"]).find_ticket(pending["ticketId"])
        assert ticket.phase == "inspect"
        start = _request(server, "POST", "/excel/execute-start", {"ticketId": pending["ticketId"]}, headers)
        assert start == (409, {"ok": False, "error": "Ticket execution boundary is not available", "category": "Conflict"})
        assert ticket.state == "claimed"
        executed = _request(server, "POST", "/excel/result",
                            {"ticketId": pending["ticketId"], "executed": True, "result": _R4G2_COMMITTED}, headers)
        assert executed == (409, {"ok": False, "error": "Ticket in state claimed, cannot accept this result disposition",
                                  "category": "Conflict"})
        assert ticket.state == "claimed" and ticket.result_executed is None
        assert _request(server, "POST", "/excel/result",
                        {"ticketId": pending["ticketId"], "executed": False, "result": _R4G2_INSPECTION}, headers)[0] == 200
        assert _r4g2_join(thread, box) == (200, {"ok": True, "ticketId": pending["ticketId"], "requestId": op_i,
                                                 "data": _R4G2_INSPECTION, "executed": False})
        assert ticket.terminal_reason == "not_executed"
        second = _r4g2_execute_done(server, session, op_e, key)
        assert second["trusted"] == _r4g2_trusted(op_e, "execute", key)
        session_obj = server.sessions.get(session["sessionId"])
        assert list(session_obj._plan_turn_history) == [op_i, op_e]
        assert {k: sorted(v) for k, v in session_obj._plan_turn_phases.items()} == {key: ["execute", "inspect"]}
    finally:
        server.stop()


def test_r4g2_c1b_white_box_an_inspect_ticket_refuses_executed_true_even_if_forced_to_executing() -> None:
    ticket = _R4g2Ticket(ticket_id=_R4G2_OP[0], session_id="s", request_id=_R4G2_OP[3], tool="excel_execute_plan",
                         arguments=dict(_R4G2_INSPECT), trusted=_r4g2_trusted(_R4G2_OP[3], "inspect", _R4G2_KEY[3]))
    ticket.state = "claimed"
    ticket.claimed_by_pane_id = "pane"
    ticket.claimed_at = ticket._clock()
    assert ticket.mark_executing("pane") == "invalid:phase"
    assert ticket.state == "claimed"
    ticket.state = "executing"  # not reachable over HTTP (C1); forced to probe the result boundary alone
    assert ticket.submit_result(_R4G2_COMMITTED, "pane", executed=True) == "invalid:executing"
    assert ticket.state == "executing" and ticket.result_executed is None
    execute = _R4g2Ticket(ticket_id=_R4G2_OP[0], session_id="s", request_id=_R4G2_OP[4], tool="excel_execute_plan",
                          arguments=dict(_R4G2_EXECUTE), trusted=_r4g2_trusted(_R4G2_OP[4], "execute", _R4G2_KEY[3]))
    assert execute.phase == "execute"
    primitive = _R4g2Ticket(ticket_id=_R4G2_OP[0], session_id="s", request_id=_R4G2_OP[5], tool="read_range",
                            arguments={"sheet": "Sheet1", "address": "A1"})
    assert primitive.phase is None


def test_r4g2_c2_fence_refuses_a_second_inspect_and_anything_after_an_execute_and_stores_nothing() -> None:
    server = _r4g2_server()
    try:
        session = _session(server)
        session_obj = server.sessions.get(session["sessionId"])
        key = _R4G2_KEY[2]
        _r4g2_inspect_done(server, session, _R4G2_OP[1], key)
        before = _r4g2_state(session_obj)
        assert _r4g2_refused(server, session, _R4G2_OP[2], _R4G2_INSPECT, _r4g2_trusted(_R4G2_OP[2], "inspect", key)) \
            == (409, _R4G2_DUPLICATE)  # inspect -> inspect
        assert _r4g2_state(session_obj) == before
        _r4g2_execute_done(server, session, _R4G2_OP[3], key)
        before = _r4g2_state(session_obj)
        assert _r4g2_refused(server, session, _R4G2_OP[4], _R4G2_INSPECT, _r4g2_trusted(_R4G2_OP[4], "inspect", key)) \
            == (409, _R4G2_DUPLICATE)  # execute -> inspect
        assert _r4g2_refused(server, session, _R4G2_OP[5], _R4G2_EXECUTE, _r4g2_trusted(_R4G2_OP[5], "execute", key)) \
            == (409, _R4G2_DUPLICATE)  # execute -> execute
        assert _r4g2_state(session_obj) == before
        # a fresh turnKey is an independent turn
        _r4g2_execute_done(server, session, _R4G2_OP[6], _R4G2_KEY[6])
    finally:
        server.stop()


def test_r4g2_c2b_execute_without_inspect_then_execute_on_the_same_turnkey_is_refused() -> None:
    server = _r4g2_server()
    try:
        session = _session(server)
        _r4g2_execute_done(server, session, _R4G2_OP[1], _R4G2_KEY[1])
        assert _r4g2_refused(server, session, _R4G2_OP[2], _R4G2_EXECUTE,
                             _r4g2_trusted(_R4G2_OP[2], "execute", _R4G2_KEY[1])) == (409, _R4G2_DUPLICATE)
    finally:
        server.stop()


def test_r4g2_c3_malformed_v2_envelopes_and_phase_outcome_mismatch_are_400_without_a_ticket() -> None:
    server = _r4g2_server()
    try:
        session = _session(server)
        session_obj = server.sessions.get(session["sessionId"])
        op, key = _R4G2_OP[1], _R4G2_KEY[1]
        cases = [
            ("phase inspect + execute plan", _R4G2_EXECUTE, _r4g2_trusted(op, "inspect", key), _R4G2_INVALID),
            ("phase execute + inspect plan", _R4G2_INSPECT, _r4g2_trusted(op, "execute", key), _R4G2_INVALID),
            ("v1 two keys", _R4G2_EXECUTE, {"turnOperationId": op, "workbookPlanFingerprint": FINGERPRINT}, _R4G2_REQUIRED),
            ("three keys", _R4G2_EXECUTE, {"turnOperationId": op, "workbookPlanFingerprint": FINGERPRINT, "phase": "execute"},
             _R4G2_REQUIRED),
            ("five keys", _R4G2_EXECUTE, dict(_r4g2_trusted(op, "execute", key), extra="x"), _R4G2_REQUIRED),
            ("turnKey == turnOperationId", _R4G2_EXECUTE, _r4g2_trusted(op, "execute", op), _R4G2_INVALID),
            ("phase unknown", _R4G2_EXECUTE, _r4g2_trusted(op, "commit", key), _R4G2_INVALID),
            ("phase uppercase", _R4G2_EXECUTE, _r4g2_trusted(op, "EXECUTE", key), _R4G2_INVALID),
            ("turnKey not a uuid", _R4G2_EXECUTE, _r4g2_trusted(op, "execute", "not-a-uuid"), _R4G2_INVALID),
            ("turnKey not a string", _R4G2_EXECUTE, dict(_r4g2_trusted(op, "execute", key), turnKey=7), _R4G2_INVALID),
            ("phase not a string", _R4G2_EXECUTE, dict(_r4g2_trusted(op, "execute", key), phase=["execute"]), _R4G2_INVALID),
        ]
        for label, plan, trusted, expected in cases:
            assert _r4g2_refused(server, session, op, plan, trusted) == (400, expected), label
            assert not session_obj.has_ticket(), label
            assert _r4g2_state(session_obj) == ([], [], {}, {}), label
    finally:
        server.stop()


def test_r4g2_c4_capacity_counts_the_phase_map_refuses_the_new_op_and_never_evicts() -> None:
    server = _r4g2_server()
    try:
        session = _session(server)
        session_obj = server.sessions.get(session["sessionId"])
        session_obj._plan_turn_history_limit = 3
        session_obj._history_limit = 1  # cached results are evicted; fences and the phase map are not
        _r4g2_inspect_done(server, session, _R4G2_OP[1], _R4G2_KEY[1])
        _r4g2_execute_done(server, session, _R4G2_OP[2], _R4G2_KEY[1])
        _r4g2_inspect_done(server, session, _R4G2_OP[3], _R4G2_KEY[2])
        before = _r4g2_state(session_obj)
        assert before[3] == {_R4G2_KEY[1]: ["execute", "inspect"], _R4G2_KEY[2]: ["inspect"]}
        status, body = _r4g2_refused(server, session, _R4G2_OP[4], _R4G2_EXECUTE,
                                     _r4g2_trusted(_R4G2_OP[4], "execute", _R4G2_KEY[2]))
        assert (status, body) == (429, {"ok": False, "error": "Workbook plan capacity reached for this session",
                                        "category": "Capacity", "state": "plan_fence_capacity"})
        assert _r4g2_state(session_obj) == before
        # the evicted turnKey K1 is still fenced (never evicted), and K2's admitted inspect still refuses a second one
        assert _r4g2_refused(server, session, _R4G2_OP[5], _R4G2_INSPECT,
                             _r4g2_trusted(_R4G2_OP[5], "inspect", _R4G2_KEY[1])) == (409, _R4G2_DUPLICATE)
        assert _r4g2_refused(server, session, _R4G2_OP[6], _R4G2_INSPECT,
                             _r4g2_trusted(_R4G2_OP[6], "inspect", _R4G2_KEY[2])) == (409, _R4G2_DUPLICATE)
        assert _r4g2_state(session_obj) == before
    finally:
        server.stop()


def test_r4g2_c5_replay_identity_includes_phase_and_turnkey() -> None:
    server = _r4g2_server()
    try:
        session = _session(server)
        session_obj = server.sessions.get(session["sessionId"])
        op, key = _R4G2_OP[1], _R4G2_KEY[1]
        first = _r4g2_inspect_done(server, session, op, key)
        # same identity: reattach to the cached terminal result (no second ticket)
        status, body = _request(server, "POST", "/excel/invoke", _r4g2_body(session, op, _R4G2_INSPECT,
                                                                             _r4g2_trusted(op, "inspect", key)),
                                {"X-Hermes-Plugin-Key": PLUGIN_KEY})
        assert (status, body["ticketId"], body["executed"]) == (200, first["ticketId"], False)
        # same op id and plan, other turnKey: conflict (request history)
        assert _r4g2_refused(server, session, op, _R4G2_INSPECT, _r4g2_trusted(op, "inspect", _R4G2_KEY[2])) == (
            409, {"ok": False, "error": "requestId reused with different payload", "category": "Conflict"})
        # evict the cached result: the replay fence keeps the 5-part identity
        session_obj._history_limit = 1
        _r4g2_execute_done(server, session, _R4G2_OP[2], _R4G2_KEY[2])
        assert op not in session_obj._request_history and op in session_obj._plan_replay_fences
        assert _r4g2_refused(server, session, op, _R4G2_INSPECT, _r4g2_trusted(op, "inspect", key)) == (
            409, {"ok": False, "error": "Workbook plan replay identity is no longer executable", "category": "Conflict"})
        assert _r4g2_refused(server, session, op, _R4G2_INSPECT, _r4g2_trusted(op, "inspect", _R4G2_KEY[3])) == (
            409, {"ok": False, "error": "requestId reused with different payload", "category": "Conflict"})
    finally:
        server.stop()


def test_r4g2_b_a3_turnkey_fence_survives_terminal_tickets_and_later_pane_claims() -> None:
    server = _r4g2_server()
    try:
        session = _session(server)
        session_obj = server.sessions.get(session["sessionId"])
        key = _R4G2_KEY[1]
        _r4g2_inspect_done(server, session, _R4G2_OP[1], key)
        _r4g2_execute_done(server, session, _R4G2_OP[2], key)
        # later pane turns: other turnKeys are served in between, then the old turnKey is retried
        for n in (3, 4):
            _r4g2_execute_done(server, session, _R4G2_OP[n], _R4G2_KEY[n])
        before = _r4g2_state(session_obj)
        for n, phase, plan in ((5, "inspect", _R4G2_INSPECT), (6, "execute", _R4G2_EXECUTE)):
            assert _r4g2_refused(server, session, _R4G2_OP[n], plan, _r4g2_trusted(_R4G2_OP[n], phase, key)) \
                == (409, _R4G2_DUPLICATE)
        assert _r4g2_state(session_obj) == before
        # the fence is per session: a second pane session has its own (the plugin derives the turnKey from the
        # session id, test_r4g2_p7), so the same literal key there is a different Hermes turn
        other = _session(server)
        _r4g2_execute_done(server, other, _R4G2_OP[7], key)
    finally:
        server.stop()


def test_r4g2_b_a4_inspect_result_size_cap_is_the_result_body_cap() -> None:
    server = _r4g2_server()
    try:
        session = _session(server)
        headers = _pane_headers(session)
        op, key = _R4G2_OP[1], _R4G2_KEY[1]
        thread, box = _r4g2_invoke(server, session, op, _R4G2_INSPECT, _r4g2_trusted(op, "inspect", key))
        pending = _r4g2_claim(server, session)
        assert _r4g2_protocol.MAX_RESULT_BODY == 10 * 1024 * 1024
        # Declared one byte over the cap: refused from the header alone, before any body byte is read.
        import http.client as _r4g2_http
        host, port = server.server_address()
        conn = _r4g2_http.HTTPConnection(host, port, timeout=5)
        try:
            conn.putrequest("POST", "/excel/result")
            for name, value in {**headers, "Content-Type": "application/json",
                                "Content-Length": str(_r4g2_protocol.MAX_RESULT_BODY + 1)}.items():
                conn.putheader(name, value)
            conn.endheaders()
            status = conn.getresponse().status
        finally:
            conn.close()
        assert status == 413
        ticket = server.sessions.get(session["sessionId"]).find_ticket(pending["ticketId"])
        assert ticket.state == "claimed" and ticket.result_payload is None
        assert _request(server, "POST", "/excel/result",
                        {"ticketId": pending["ticketId"], "executed": False, "result": _R4G2_INSPECTION}, headers)[0] == 200
        assert _r4g2_join(thread, box)[1]["data"] == _R4G2_INSPECTION
    finally:
        server.stop()


# ---------------------------------------------------------------------------
# R4g-2b (OMH release record 2026-10-07, scope (3); R4g-2a lane-B A2): the companion admits a plan turnKey and
# op id (turnOperationId == requestId) only as canonical lowercase hyphenated UUIDs. Every non-canonical
# spelling is the existing 400 "Trusted workbook plan metadata is invalid" with no ticket and no fence state.
# Ids carry hex letters so the case variants differ from the genuine id. Primitive requestIds are unchanged.
# ---------------------------------------------------------------------------
_R4G2B_OP = "5e7eabcd-ef01-4a2b-8c3d-4e5f6a7b8c9d"
_R4G2B_KEY = "4b4bcdef-0123-4abc-9def-0123456789ab"


def _r4g2b_noncanonical(value: str) -> list[tuple[str, str]]:
    return [
        ("uppercase", value.upper()),
        ("one uppercase letter", value[:-1] + value[-1].upper()),
        ("braces", "{" + value + "}"),
        ("urn:uuid: prefix", "urn:uuid:" + value),
        ("no dashes", value.replace("-", "")),
        ("nil", "00000000-0000-0000-0000-000000000000"),
    ]


def test_r4g2b_k0_control_canonical_lowercase_ids_are_admitted_for_both_phases() -> None:
    server = _r4g2_server()
    try:
        session = _session(server)
        assert _R4G2B_OP.upper() != _R4G2B_OP and _R4G2B_KEY[-1] in "abcdef"
        _r4g2_inspect_done(server, session, _R4G2B_OP, _R4G2B_KEY)
        other = "6f8aabcd-0c2e-4b6d-9f8a-0c1e3a5c7e9b"
        _r4g2_execute_done(server, session, other, _R4G2B_KEY)
    finally:
        server.stop()


def _r4g2b_expect_refused(op: str, key: str, label: str) -> None:
    """One fresh server per spelling (few requests per server; never near the handler-thread cap)."""
    server = _r4g2_server()
    try:
        session = _session(server)
        session_obj = server.sessions.get(session["sessionId"])
        for phase, plan in (("inspect", _R4G2_INSPECT), ("execute", _R4G2_EXECUTE)):
            # requestId == turnOperationId, as an honest client pairs them.
            outcome = _r4g2_refused(server, session, op, plan, _r4g2_trusted(op, phase, key))
            assert outcome == (400, _R4G2_INVALID), (label, phase, outcome)
            assert not session_obj.has_ticket(), (label, phase)
            assert _r4g2_state(session_obj) == ([], [], {}, {}), (label, phase)
    finally:
        server.stop()


def test_r4g2b_k1_noncanonical_turnkey_is_400_invalid_without_a_ticket() -> None:
    cases = list(_r4g2b_noncanonical(_R4G2B_KEY))
    cases.append(("turnKey == op uppercased (same UUID value)", _R4G2B_OP.upper()))
    for label, form in cases:
        _r4g2b_expect_refused(_R4G2B_OP, form, "turnKey " + label)


def test_r4g2b_k2_noncanonical_op_id_is_400_invalid_without_a_ticket() -> None:
    for label, form in _r4g2b_noncanonical(_R4G2B_OP):
        _r4g2b_expect_refused(form, _R4G2B_KEY, "op " + label)


def test_r4g2b_k3_primitive_request_ids_keep_their_existing_validation() -> None:
    server = _r4g2_server()
    try:
        session = _session(server)
        upper = "22222222-2222-4222-8222-2222222222ab".upper()
        box: dict = {}

        def run() -> None:
            box["status"], box["body"] = _request(server, "POST", "/excel/invoke", {
                "protocolVersion": "1.0", "sessionId": session["sessionId"], "requestId": upper,
                "tool": "read_range", "arguments": {"sheet": "Sheet1", "address": "A1"},
            }, {"X-Hermes-Plugin-Key": PLUGIN_KEY})

        thread = threading.Thread(target=run, daemon=True)
        thread.start()
        _wait_for_ticket(server, session["sessionId"])
        status, pending = _request(server, "GET", "/excel/pending", headers=_pane_headers(session))
        assert status == 200 and pending["requestId"] == upper and "trusted" not in pending, (status, pending)
        assert _request(server, "POST", "/excel/result", {"ticketId": pending["ticketId"], "result": {"ok": True}},
                        _pane_headers(session))[0] == 200
        thread.join(5)
        assert not thread.is_alive()
        assert box["status"] == 200, box
    finally:
        server.stop()


# ---------------------------------------------------------------------------
# R4g-2c (dev3-r4g2c-release.md item 6; R4g-2b lane-B A4): a result body nested beyond the JSON decoder's recursion
# limit is the companion's existing 400 "Bad Request: invalid JSON" (no new string), not a 500. The ticket stays
# claimed without a result, so a later conforming result still completes the invoke. The deep body is sent as raw
# bytes (Python's own encoder would refuse to build it).
# ---------------------------------------------------------------------------
def _r4g2c_post_raw(server: ExcelCompanionServer, path: str, raw: bytes, headers: dict) -> tuple[int, dict]:
    import http.client as _r4g2c_http
    host, port = server.server_address()
    conn = _r4g2c_http.HTTPConnection(host, port, timeout=5)
    try:
        conn.request("POST", path, body=raw, headers={**headers, "Content-Type": "application/json"})
        response = conn.getresponse()
        data = response.read()
        return response.status, json.loads(data) if data else {}
    finally:
        conn.close()


def test_r4g2c_k4_a_too_deep_result_body_is_400_invalid_json_and_the_ticket_survives() -> None:
    server = _r4g2_server()
    try:
        session = _session(server)
        headers = _pane_headers(session)
        op, key = _R4G2_OP[3], _R4G2_KEY[3]
        thread, box = _r4g2_invoke(server, session, op, _R4G2_INSPECT, _r4g2_trusted(op, "inspect", key))
        pending = _r4g2_claim(server, session)
        deep = "[" * 5000 + "]" * 5000
        raw = ('{"ticketId":"' + pending["ticketId"] + '","executed":false,"result":{"ok":true,"action":"workbook_plan",'
               '"phase":"inspect","result":{"outcome":"inspected","inspectionCount":1,"reads":[{"ok":true,"data":'
               + deep + '}]}}}').encode()
        status, body = _r4g2c_post_raw(server, "/excel/result", raw, headers)
        assert (status, body.get("ok"), body.get("error")) == (400, False, "Bad Request: invalid JSON"), (status, body)
        ticket = server.sessions.get(session["sessionId"]).find_ticket(pending["ticketId"])
        assert ticket.state == "claimed" and ticket.result_payload is None
        # control: a 64-level read is an ordinary result
        shallow = "[" * 63 + "]" * 63
        raw_ok = ('{"ticketId":"' + pending["ticketId"] + '","executed":false,"result":{"ok":true,"action":"workbook_plan",'
                  '"phase":"inspect","result":{"outcome":"inspected","inspectionCount":1,"reads":[{"ok":true,"data":'
                  + shallow + '}]}}}').encode()
        assert _r4g2c_post_raw(server, "/excel/result", raw_ok, headers)[0] == 200
        status, invoked = _r4g2_join(thread, box)
        assert status == 200 and invoked["executed"] is False, (status, invoked)
        assert invoked["data"]["result"]["reads"] == [{"ok": True, "data": json.loads(shallow)}]
    finally:
        server.stop()


# ---------------------------------------------------------------------------
# R4g-3 (dev3-r4g3-release.md §1.1 item 3, §3 K7, §5.1 E-2, §5.2 Python legs; RE-PIN R2/R4; OMH release record
# 2026-10-07): the REAL plugin (pre_tool_call hook + plan handler, a fresh registered() module) over the REAL
# loopback companion, with an HTTP pane double that claims tickets and posts results. Parity pins (E6): the plugin
# and companion already speak inspect -> execute, so every row is green at 3e3707b and at HEAD. The three body
# literals below are pinned independently in src/__tests__/server-inspect-e2e-r4g3.test.ts: K7 (the fixture's canned
# T8 body) and E-8/E-9 (the bytes the REAL taskpane server reader makes the REAL bridge post there). Each passes the
# plugin's closed T8 acceptance unchanged (not the fixed not-executed body).
# ---------------------------------------------------------------------------
_R4G3_K7 = ('{"ok":true,"action":"workbook_plan","phase":"inspect","result":{"outcome":"inspected","inspectionCount":1,'
            '"reads":[{"ok":true,"data":[["ok"]]}]}}')
_R4G3_E8 = ('{"ok":true,"action":"workbook_plan","phase":"inspect","result":{"outcome":"inspected","inspectionCount":1,'
            '"reads":[{"ok":true,"data":{"sheet":"Sheet1","address":"Sheet1!A1","rowCount":1,"columnCount":1,'
            '"values":[["ok"]],"formulas":[["ok"]],"numberFormat":[["General"]]}}]}}')
_R4G3_E9 = ('{"ok":true,"action":"workbook_plan","phase":"inspect","result":{"outcome":"inspected","inspectionCount":1,'
            '"reads":[{"ok":false,"data":null}]}}')
_R4G3_FIXED = {"ok": False, "action": "workbook_plan", "category": "Execution", "reason": "rejected",
               "error": "Workbook plan was not executed."}
_R4G3_BLOCKED = {"action": "block", "message": "Workbook plan trusted context rejected."}
_R4G3_COMMITTED = {"ok": True, "action": "workbook_plan", "result": {"outcome": "committed"}}
_R4G3_TURN = "r4g3-hermes-turn-1"


def _r4g3_dump(value) -> str:
    return json.dumps(value, separators=(",", ":"), ensure_ascii=False)


def _r4g3_claim(companion: _R3dCompanion) -> dict:
    companion.wait_pending()
    status, pending = _request(companion.server, "GET", "/excel/pending", headers=companion.headers)
    assert status == 200, (status, pending)
    return pending


def _r4g3_post(companion: _R3dCompanion, ticket_id: str, executed: bool, result) -> int:
    return _request(companion.server, "POST", "/excel/result",
                    {"ticketId": ticket_id, "executed": executed, "result": result}, companion.headers)[0]


def _r4g3_inspect(contract, companion: _R3dCompanion, plugin, turn: str, body_text: str, call_id: str):
    """Hermes-like inspect call; the pane double posts body_text (parsed, key order kept) executed:false."""
    thread, box = _r3d_start(contract, plugin, companion.session["sessionId"], turn, dict(_R4G2_INSPECT), call_id)
    pending = _r4g3_claim(companion)
    assert pending["arguments"] == _R4G2_INSPECT and pending["trusted"]["phase"] == "inspect", pending
    assert _r4g3_post(companion, pending["ticketId"], False, json.loads(body_text)) == 200
    kind, raw = contract._r3a_join(thread, box)
    assert kind == "result", (kind, raw)
    return pending, raw


def _r4g3_execute(contract, companion: _R3dCompanion, plugin, turn: str, call_id: str):
    thread, box = _r3d_start(contract, plugin, companion.session["sessionId"], turn, valid_execute_plan(), call_id)
    pending = _r4g3_claim(companion)
    assert pending["trusted"]["phase"] == "execute", pending
    assert companion.ack(pending["ticketId"]) == 200
    assert _r4g3_post(companion, pending["ticketId"], True, _R4G3_COMMITTED) == 200
    kind, raw = contract._r3a_join(thread, box)
    assert kind == "result", (kind, raw)
    return pending, raw


def _r4g3_blocked(contract, plugin, sid: str, turn: str, plan: dict, call_id: str) -> None:
    thread, box = _r3d_start(contract, plugin, sid, turn, plan, call_id)
    assert contract._r3a_join(thread, box) == ("blocked", _R4G3_BLOCKED), call_id


def test_r4g3_e2_real_plugin_and_companion_inspect_then_one_execute_in_one_hermes_turn() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    companion = _R3dCompanion(contract.generated_fingerprint()).start()
    try:
        calls: list = []
        sid = companion.session["sessionId"]
        plugin = _r3d_plugin(contract)
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            pending_i, raw_i = _r4g3_inspect(contract, companion, plugin, _R4G3_TURN, _R4G3_K7, "call-r4g3-inspect")
            assert companion.acks == 0 and companion.pending_status() == 204
            pending_e, raw_e = _r4g3_execute(contract, companion, plugin, _R4G3_TURN, "call-r4g3-execute")
        view_i = json.loads(raw_i)
        assert view_i == {"executed": False, "result": json.loads(_R4G3_K7)}
        assert list(view_i) == ["executed", "result"]
        assert _r4g3_dump(view_i["result"]) == _R4G3_K7  # K7 (d): byte identity, key order included
        assert view_i["result"] != _R4G3_FIXED  # RE-PIN R4: the plugin passed the T8 body unchanged
        assert json.loads(raw_e) == _R4G3_COMMITTED
        op_i, op_e = _r3d_op(sid, _R4G3_TURN, "inspect"), _r3d_op(sid, _R4G3_TURN, "execute")
        assert (pending_i["trusted"]["turnOperationId"], pending_e["trusted"]["turnOperationId"]) == (op_i, op_e)
        assert pending_i["trusted"]["turnKey"] == pending_e["trusted"]["turnKey"]
        assert list(pending_e["trusted"]) == ["turnOperationId", "workbookPlanFingerprint", "phase", "turnKey"]
        assert pending_i["ticketId"] != pending_e["ticketId"]
        assert list(companion.session_obj._request_history) == [op_i, op_e]
        assert {k: sorted(v) for k, v in companion.session_obj._plan_turn_phases.items()} == {
            pending_i["trusted"]["turnKey"]: ["execute", "inspect"]}
        assert companion.acks == 1 and companion.ack_statuses == [200]
        assert len(calls) == 2
        assert companion.pending_status() == 204
    finally:
        companion.stop()


def test_r4g3_k7_fixture_c24_and_c30_bodies_dump_to_the_k7_literal() -> None:
    import hashlib
    raw = _R3D_FIXTURE.read_bytes()
    assert hashlib.sha256(raw).hexdigest() == _R3D_FIXTURE_SHA256
    cases = {case["id"]: case for case in json.loads(raw)["cases"]}

    def posted_inspection(case_id: str):
        return next(step["body"]["result"] for step in cases[case_id]["steps"]
                    if step["kind"] == "pane" and step.get("path") == "/excel/result" and step["body"]["executed"] is False)

    assert _r4g3_dump(posted_inspection("C24-inspect-then-execute")) == _R4G3_K7
    assert _r4g3_dump(posted_inspection("C30-inspect-ticket-gate-on")) == _R4G3_K7
    assert _r4g3_dump(cases["C30-inspect-ticket-gate-on"]["bridge"]["paneReport"]) == _R4G3_K7
    assert "C30-inspect-ticket-gate-off" not in cases


def test_r4g3_e8_real_reader_bodies_pass_the_plugin_unchanged_through_the_companion() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    companion = _R3dCompanion(contract.generated_fingerprint()).start()
    try:
        calls: list = []
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            for n, literal in enumerate((_R4G3_E8, _R4G3_E9)):
                _pending, raw = _r4g3_inspect(contract, companion, _r3d_plugin(contract), f"r4g3-e8-turn-{n}",
                                              literal, f"call-r4g3-e8-{n}")
                view = json.loads(raw)
                assert view == {"executed": False, "result": json.loads(literal)}, n
                assert _r4g3_dump(view["result"]) == literal, n
        assert companion.acks == 0 and len(calls) == 2
    finally:
        companion.stop()


def test_r4g3_n1_n2_n3_plugin_blocks_a_second_plan_call_and_the_companion_refuses_it_if_forced() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    companion = _R3dCompanion(contract.generated_fingerprint()).start()
    try:
        calls: list = []
        sid = companion.session["sessionId"]
        plugin = _r3d_plugin(contract)
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            pending_i, _raw = _r4g3_inspect(contract, companion, plugin, _R4G3_TURN, _R4G3_K7, "call-r4g3-n-i")
            key, fp = pending_i["trusted"]["turnKey"], pending_i["trusted"]["workbookPlanFingerprint"]

            def forced(op: str, phase: str) -> dict:
                return dict(_r4g2_trusted(op, phase, key), workbookPlanFingerprint=fp)

            # N1 inspect -> inspect: blocked by the plugin; forced past it -> companion 409
            _r4g3_blocked(contract, plugin, sid, _R4G3_TURN, dict(_R4G2_INSPECT), "call-r4g3-n1")
            before = _r4g2_state(companion.session_obj)
            assert _r4g2_refused(companion.server, companion.session, _R4G2_OP[1], _R4G2_INSPECT,
                                 forced(_R4G2_OP[1], "inspect")) == (409, _R4G2_DUPLICATE)
            assert _r4g2_state(companion.session_obj) == before
            _r4g3_execute(contract, companion, plugin, _R4G3_TURN, "call-r4g3-n-e")
            # N2 execute -> inspect and N3 execute -> execute: blocked; forced -> 409; nothing stored
            _r4g3_blocked(contract, plugin, sid, _R4G3_TURN, dict(_R4G2_INSPECT), "call-r4g3-n2")
            _r4g3_blocked(contract, plugin, sid, _R4G3_TURN, valid_execute_plan(), "call-r4g3-n3")
            before = _r4g2_state(companion.session_obj)
            assert _r4g2_refused(companion.server, companion.session, _R4G2_OP[2], _R4G2_INSPECT,
                                 forced(_R4G2_OP[2], "inspect")) == (409, _R4G2_DUPLICATE)
            assert _r4g2_refused(companion.server, companion.session, _R4G2_OP[3], _R4G2_EXECUTE,
                                 forced(_R4G2_OP[3], "execute")) == (409, _R4G2_DUPLICATE)
            assert _r4g2_state(companion.session_obj) == before
        assert companion.acks == 1 and len(calls) == 2 + 3  # 2 plugin invokes; the 3 blocked calls never invoked; 3 forced
    finally:
        companion.stop()


def test_r4g3_n2b_execute_first_then_inspect_is_blocked_and_refused_if_forced() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    companion = _R3dCompanion(contract.generated_fingerprint()).start()
    try:
        calls: list = []
        sid = companion.session["sessionId"]
        plugin = _r3d_plugin(contract)
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            pending_e, _raw = _r4g3_execute(contract, companion, plugin, _R4G3_TURN, "call-r4g3-n2b-e")
            key, fp = pending_e["trusted"]["turnKey"], pending_e["trusted"]["workbookPlanFingerprint"]
            _r4g3_blocked(contract, plugin, sid, _R4G3_TURN, dict(_R4G2_INSPECT), "call-r4g3-n2b-i")
            before = _r4g2_state(companion.session_obj)
            forced = dict(_r4g2_trusted(_R4G2_OP[5], "inspect", key), workbookPlanFingerprint=fp)
            assert _r4g2_refused(companion.server, companion.session, _R4G2_OP[5], _R4G2_INSPECT, forced) == (409, _R4G2_DUPLICATE)
            assert _r4g2_state(companion.session_obj) == before
            assert {k: sorted(v) for k, v in companion.session_obj._plan_turn_phases.items()} == {key: ["execute"]}
        assert companion.acks == 1 and len(calls) == 1 + 1  # 1 plugin invoke; the blocked call never invoked; 1 forced
    finally:
        companion.stop()


def test_r4g3_n4_phase_outcome_mismatch_is_400_without_a_ticket() -> None:
    server = _r4g2_server()
    try:
        session = _session(server)
        session_obj = server.sessions.get(session["sessionId"])
        op, key = _R4G2_OP[4], _R4G2_KEY[4]
        for plan, phase in ((_R4G2_EXECUTE, "inspect"), (_R4G2_INSPECT, "execute")):
            assert _r4g2_refused(server, session, op, plan, _r4g2_trusted(op, phase, key)) == (400, _R4G2_INVALID)
            assert not session_obj.has_ticket() and _r4g2_state(session_obj) == ([], [], {}, {})
    finally:
        server.stop()


def test_r4g3_n7_n8_replays_reattach_to_the_cached_views_and_execute_nothing_again() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    companion = _R3dCompanion(contract.generated_fingerprint()).start()
    try:
        calls: list = []
        sid = companion.session["sessionId"]
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            _pi, raw_i = _r4g3_inspect(contract, companion, _r3d_plugin(contract), _R4G3_TURN, _R4G3_K7, "call-r4g3-r-i")
            _pe, raw_e = _r4g3_execute(contract, companion, _r3d_plugin(contract), _R4G3_TURN, "call-r4g3-r-e")
            history = list(companion.session_obj._request_history)
            # N7 / N8: a fresh plugin instance (empty seen set) re-invokes each op of the same Hermes turn
            for plan, first, call_id in ((dict(_R4G2_INSPECT), raw_i, "call-r4g3-r7"), (valid_execute_plan(), raw_e, "call-r4g3-r8")):
                thread, box = _r3d_start(contract, _r3d_plugin(contract), sid, _R4G3_TURN, plan, call_id)
                kind, replay = contract._r3a_join(thread, box)
                assert (kind, replay) == ("result", first), call_id  # byte-equal cached view
                assert companion.pending_status() == 204, call_id
        assert list(companion.session_obj._request_history) == history
        assert companion.acks == 1 and len(calls) == 4
        assert _r4g3_dump(json.loads(raw_i)["result"]) == _R4G3_K7
    finally:
        companion.stop()


def test_r4g3_n12_hostile_pane_cannot_execute_an_inspect_ticket_and_the_plugin_view_is_never_executed() -> None:
    from scripts.tests import test_office_plugin_plan_contract as contract

    companion = _R3dCompanion(contract.generated_fingerprint()).start()
    try:
        calls: list = []
        p1, p2, p3 = companion.patches(contract.load_plugin().bridge_client, calls)
        with p1, p2, p3:
            thread, box = _r3d_start(contract, _r3d_plugin(contract), companion.session["sessionId"], _R4G3_TURN,
                                     dict(_R4G2_INSPECT), "call-r4g3-n12")
            pending = _r4g3_claim(companion)
            assert companion.ack(pending["ticketId"]) == 409
            assert _r4g3_post(companion, pending["ticketId"], True, _R4G3_COMMITTED) == 409
            assert _r4g3_post(companion, pending["ticketId"], False, json.loads(_R4G3_K7)) == 200
            kind, raw = contract._r3a_join(thread, box)
        assert kind == "result"
        view = json.loads(raw)
        assert view["executed"] is False and _r4g3_dump(view["result"]) == _R4G3_K7
        assert companion.acks == 0 and companion.ack_statuses == [409]
    finally:
        companion.stop()
