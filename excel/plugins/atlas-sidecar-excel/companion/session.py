"""Session and ticket state machine for the Excel companion RPC server.

Session: binds documentId + paneId, holds TTL, heartbeat, outstanding ticket.
Ticket: one per session, queue depth 1, atomic claim, lease/requeue/exhaustion.

Clock injection: all time-dependent methods accept a `clock` callable
(default time.monotonic) for deterministic unit testing.

Ticket lifecycle:
  - pending -> claimed -> completed (normal flow via result)
  - pending -> cancelled (cancel before claim or during invoke)
  - pending -> timeout (invoke timeout before claim)
  - claimed -> pending (lease requeue, max 1 time)
  - claimed -> timeout (second lease expiry)
  - completed tickets are RETAINED briefly for duplicate idempotency;
    a new invoke after completion/cancel/timeout clears the old ticket.

Plan (excel_execute_plan) lifecycle is one-way and never requeued:
  - pending -> claimed -> executing -> terminal(completed)   [executed:true]
  - pending -> claimed -> terminal(not_executed)             [executed:false]
  - claimed lease expiry -> terminal(expired_claimed), including an expired,
    not-yet-swept claim found at execute-start (mark_executing rechecks the lease)
  - executing deadline (from executing_at) or cancel -> terminal(executing_uncertain)

Plan safety state (R3a): the companion is the single authority for plan-op
acceptance. A session's plan turn history and plan replay fences are never
evicted during its life; both are bounded by one per-session limit, and a new
plan op that would exceed it is rejected (plan_fence_capacity) without storing
anything. Sessions are per-process nonces and are never persisted or restored.
"""
from __future__ import annotations

import hashlib
import json
import threading
import time
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Any, Callable, Union

from .protocol import (
    CLAIM_LEASE_SECONDS,
    INVOKE_TIMEOUT_SECONDS,
    PLAN_EXECUTING_LEASE_SECONDS,
    MAX_REQUEST_HISTORY,
    MAX_RETAINED_HISTORY_BYTES,
    MAX_PLAN_REPLAY_FENCES,
    MAX_SESSIONS,
    SESSION_TTL_SECONDS,
    constant_time_compare,
    generate_uuid,
)

# Default clock for production use
_DEFAULT_CLOCK: Callable[[], float] = time.monotonic
ClockOrNow = Union[Callable[[], float], float, None]


def _resolve_now(value: ClockOrNow, default_clock: Callable[[], float]) -> float:
    """Resolve an injected clock callable, explicit timestamp, or default clock."""
    if callable(value):
        return value()
    if value is not None:
        return float(value)
    return default_clock()


def _pane_matches(expected: str | None, provided: str) -> bool:
    ok = constant_time_compare(expected if expected is not None else provided, provided)
    return expected is not None and ok


def _content_digest(value: Any) -> str:
    """Hash JSON for replay identity without retaining content in replay fences."""
    try:
        encoded = json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    except (TypeError, ValueError):
        return "invalid"
    return hashlib.sha256(encoded).hexdigest()


def _canonical_json(value: Any) -> str | None:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False)
    except (TypeError, ValueError):
        return None


def _is_terminal_state(state: str) -> bool:
    return state in ("completed", "cancelled", "timeout", "terminal")


@dataclass
class Ticket:
    """A single invoke ticket in a session's queue (depth 1)."""

    ticket_id: str
    session_id: str
    request_id: str
    tool: str
    arguments: dict[str, Any]
    trusted: dict[str, str] | None = None
    request_digest: str | None = None

    # Primitive lifecycle remains unchanged. Plan lifecycle is one-way:
    # pending (unclaimed) -> claimed -> executing -> terminal.
    state: str = "pending"
    claimed_by_pane_id: str | None = None
    claimed_at: float = 0.0
    executing_at: float = 0.0
    requeue_count: int = 0

    # Result — any JSON value (not forced to dict)
    result_payload: Any = None
    completed_at: float = 0.0
    terminal_reason: str | None = None
    # Plan-only bridge disposition of an accepted result (None until accepted).
    result_executed: bool | None = None

    # Blocked invoke waiter
    _result_event: threading.Event = field(default_factory=threading.Event)

    # Lock for state transitions
    _lock: threading.Lock = field(default_factory=threading.Lock)

    # Creation time (set by Session/Ticket factory)
    created_at: float = 0.0
    invoke_timeout_seconds: float = INVOKE_TIMEOUT_SECONDS
    _clock: Callable[[], float] = field(default=_DEFAULT_CLOCK, repr=False, compare=False)

    def __post_init__(self) -> None:
        if self.created_at == 0.0:
            self.created_at = self._clock()
        if self.request_digest is None:
            self.request_digest = _content_digest(self.arguments)

    @property
    def is_plan(self) -> bool:
        return self.tool == "excel_execute_plan"

    @property
    def phase(self) -> str | None:
        """R4g-2a: the trusted plan phase (None for primitives)."""
        return None if self.trusted is None else self.trusted.get("phase")

    def _now(self, clock: ClockOrNow = None) -> float:
        return _resolve_now(clock, self._clock)

    def wait_for_result(self, timeout: float | None = None) -> bool:
        return self._result_event.wait(timeout)

    def set_result(self, payload: Any, clock: Callable[[], float] | None = None) -> None:
        """Set a result for test/primitive compatibility paths."""
        with self._lock:
            if _is_terminal_state(self.state):
                return
            self.state = "terminal" if self.is_plan else "completed"
            self.result_payload = payload
            self.terminal_reason = "completed" if self.is_plan else None
            self.result_executed = True if self.is_plan else None
            self.completed_at = self._now(clock)
        self._result_event.set()

    def submit_result(self, payload: Any, pane_id: str,
                      clock: Callable[[], float] | None = None,
                      executed: bool | None = None) -> str:
        """Atomically accept a result at the only permitted lifecycle boundary.

        Plan tickets require the bridge-level ``executed`` disposition:
        ``True`` only from ``executing`` and ``False`` only from ``claimed``,
        both by the claiming pane. Exact duplicates of an accepted plan result
        return ``"duplicate"``; anything else that conflicts leaves state
        unchanged. Primitive tickets ignore ``executed``.
        """
        with self._lock:
            if self.is_plan:
                disposition = self._accept_plan_result_locked(payload, pane_id, executed, clock)
                if disposition != "accepted":
                    return disposition
            else:
                if _is_terminal_state(self.state):
                    return "completed" if self.terminal_reason == "completed" else "terminal"
                if self.state == "claimed":
                    if not _pane_matches(self.claimed_by_pane_id, pane_id):
                        return "wrong_pane"
                elif not (self.state == "pending" and self.requeue_count == 1):
                    return f"invalid:{self.state}"
                self.state = "completed"
                self.result_payload = payload
                self.terminal_reason = None
                self.completed_at = self._now(clock)
        self._result_event.set()
        return "accepted"

    def _accept_plan_result_locked(self, payload: Any, pane_id: str,
                                   executed: Any, clock: ClockOrNow) -> str:
        """One-way plan result transition; caller holds ``self._lock``."""
        if type(executed) is not bool:
            return "invalid_disposition"
        if _is_terminal_state(self.state):
            if self.terminal_reason not in ("completed", "not_executed"):
                return "terminal"
            if not _pane_matches(self.claimed_by_pane_id, pane_id):
                return "wrong_pane"
            incoming = _canonical_json(payload)
            if (
                self.result_executed is executed
                and incoming is not None
                and incoming == _canonical_json(self.result_payload)
            ):
                return "duplicate"
            return "conflict"
        if self.state not in ("claimed", "executing"):
            return f"invalid:{self.state}"
        if not _pane_matches(self.claimed_by_pane_id, pane_id):
            return "wrong_pane"
        if executed and self.state != "executing":
            return f"invalid:{self.state}"
        if executed and self.phase == "inspect":
            # R4g-2a: an inspect ticket never reports an execution.
            return f"invalid:{self.state}"
        if not executed and self.state != "claimed":
            return f"invalid:{self.state}"
        self.state = "terminal"
        self.result_payload = payload
        self.result_executed = executed
        self.terminal_reason = "completed" if executed else "not_executed"
        self.completed_at = self._now(clock)
        return "accepted"

    def cancel(self) -> None:
        """Cancel without ever claiming an executing mutation was undone."""
        with self._lock:
            if _is_terminal_state(self.state):
                return
            if self.is_plan:
                self.terminal_reason = {
                    "pending": "cancelled_unclaimed",
                    "claimed": "cancelled_claimed",
                    "executing": "executing_uncertain",
                }.get(self.state, "cancelled")
                self.state = "terminal"
            else:
                self.state = "cancelled"
        self._result_event.set()

    def timeout(self, clock: Callable[[], float] | None = None) -> None:
        """Expire safely; executing plans become uncertain terminal work."""
        with self._lock:
            if _is_terminal_state(self.state):
                return
            if self.is_plan:
                self.terminal_reason = (
                    "executing_uncertain" if self.state == "executing"
                    else "expired_claimed" if self.state == "claimed"
                    else "expired_unclaimed"
                )
                self.state = "terminal"
            else:
                self.state = "timeout"
        self._result_event.set()

    def claim(self, pane_id: str, clock: Callable[[], float] | None = None) -> bool:
        """Atomically claim this ticket for a pane."""
        with self._lock:
            if self.state != "pending":
                return False
            self.state = "claimed"
            self.claimed_by_pane_id = pane_id
            self.claimed_at = self._now(clock)
            return True

    def mark_executing(self, pane_id: str, clock: ClockOrNow = None) -> str:
        """Cross the irreversible Office execution boundary exactly once."""
        with self._lock:
            if not self.is_plan:
                return "not_plan"
            if self.phase == "inspect":
                # R4g-2a: an inspect ticket never crosses the execution boundary.
                return "invalid:phase"
            if self.state == "executing":
                return "already_executing"
            if self.state != "claimed":
                return f"invalid:{self.state}"
            if not _pane_matches(self.claimed_by_pane_id, pane_id):
                return "wrong_pane"
            now = self._now(clock)
            if now - self.claimed_at < CLAIM_LEASE_SECONDS:
                self.state = "executing"
                self.executing_at = now
                return "accepted"
            # The claim lease expired before an unswept execute-start: the
            # plan never started and is never requeued.
            self.state = "terminal"
            self.terminal_reason = "expired_claimed"
        self._result_event.set()
        return "invalid:expired"

    def requeue(self, clock: Callable[[], float] | None = None) -> bool:
        """Requeue only claimed primitive work; plan tickets are never requeued."""
        with self._lock:
            if self.state != "claimed":
                return False
            if self.is_plan or self.requeue_count >= 1:
                if self.is_plan:
                    self.state = "terminal"
                    self.terminal_reason = "expired_claimed"
                else:
                    self.state = "timeout"
                self._result_event.set()
                return False
            self.state = "pending"
            self.claimed_by_pane_id = None
            self.claimed_at = 0.0
            self.requeue_count += 1
            return True

    def is_lease_expired(self, clock: Callable[[], float] | None = None) -> bool:
        if self.state not in ("claimed", "executing"):
            return False
        if self.is_plan and self.state == "executing":
            return (self._now(clock) - self.executing_at) >= PLAN_EXECUTING_LEASE_SECONDS
        return (self._now(clock) - self.claimed_at) >= CLAIM_LEASE_SECONDS

    def is_invoke_expired(self, clock: Callable[[], float] | None = None) -> bool:
        if _is_terminal_state(self.state):
            return False
        return (self._now(clock) - self.created_at) >= self.invoke_timeout_seconds

    def is_terminal(self) -> bool:
        with self._lock:
            return self.state in ("completed", "cancelled", "timeout", "terminal")


@dataclass
class Session:
    """A single Excel document session."""

    session_id: str
    document_id: str
    pane_id: str
    session_key: str
    workbook_plan_fingerprint: str | None = None
    last_heartbeat: float = 0.0
    _clock: Callable[[], float] = field(default=_DEFAULT_CLOCK, repr=False, compare=False)

    # One ticket at a time; completed retained for bounded request replay
    _ticket: Ticket | None = None
    _lock: threading.Lock = field(default_factory=threading.Lock)
    _request_history: OrderedDict[str, Ticket] = field(default_factory=OrderedDict)
    _history_limit: int = MAX_REQUEST_HISTORY
    _history_byte_limit: int = MAX_RETAINED_HISTORY_BYTES
    _plan_turn_history: OrderedDict[str, None] = field(default_factory=OrderedDict)
    # One limit bounds both plan structures; they are never evicted. At the
    # limit a new plan op is rejected (plan_fence_capacity), never admitted.
    _plan_turn_history_limit: int = MAX_PLAN_REPLAY_FENCES
    # Content-free request fences survive cached-result eviction in this session.
    _plan_replay_fences: OrderedDict[str, tuple[str, str, str, str, str]] = field(default_factory=OrderedDict)
    # R4g-2a: turnKey -> phases admitted for that Hermes turn (inspect only
    # first, at most one execute, nothing after an execute). Never evicted;
    # bounded by the same limit as the two fences above.
    _plan_turn_phases: OrderedDict[str, set[str]] = field(default_factory=OrderedDict)

    def __post_init__(self) -> None:
        now = self._clock()
        if self.last_heartbeat == 0.0:
            self.last_heartbeat = now

    # ------------------------------------------------------------------
    # Clock-aware helpers
    # ------------------------------------------------------------------

    def _now(self, clock: ClockOrNow = None) -> float:
        return _resolve_now(clock, self._clock)

    def is_expired(self, clock: Callable[[], float] | None = None) -> bool:
        """Check if session TTL has elapsed since last heartbeat."""
        return (self._now(clock) - self.last_heartbeat) >= SESSION_TTL_SECONDS

    def heartbeat(self, clock: Callable[[], float] | None = None) -> None:
        """Update last heartbeat time."""
        self.last_heartbeat = self._now(clock)

    def has_ticket(self) -> bool:
        """Check if this session has an outstanding (active) ticket."""
        with self._lock:
            return self._ticket is not None and self._ticket.state in ("pending", "claimed", "executing")

    def get_pending_ticket(self) -> Ticket | None:
        """Get the pending ticket if any."""
        with self._lock:
            if self._ticket and self._ticket.state == "pending":
                return self._ticket
            return None

    def begin_ticket(self, ticket: Ticket) -> tuple[str, Ticket]:
        """Atomically create, reattach, or reject a request."""
        with self._lock:
            existing = self._request_history.get(ticket.request_id)
            if existing is not None:
                same_identity = (
                    existing.tool == ticket.tool
                    and existing.request_digest == ticket.request_digest
                    and existing.trusted == ticket.trusted
                ) if ticket.is_plan else (
                    existing.tool == ticket.tool
                    and existing.arguments == ticket.arguments
                    and existing.trusted == ticket.trusted
                )
                if not same_identity:
                    return "conflict", existing
                self._request_history.move_to_end(ticket.request_id)
                return "existing", existing

            if ticket.is_plan:
                # Every plan op is its own trusted turn, so every fence key is
                # also a turn-history key (unreachable over HTTP: handler 400).
                if ticket.trusted is None or ticket.trusted.get("turnOperationId") != ticket.request_id:
                    return "conflict", ticket
                fence = self._plan_replay_fences.get(ticket.request_id)
                if fence is not None:
                    identity = (
                        (ticket.trusted or {}).get("turnOperationId", ""),
                        (ticket.trusted or {}).get("workbookPlanFingerprint", ""),
                        ticket.request_digest or "invalid",
                        (ticket.trusted or {}).get("phase", ""),
                        (ticket.trusted or {}).get("turnKey", ""),
                    )
                    if fence == identity:
                        return "replay_evicted", ticket
                    return "conflict", ticket

            if self._ticket is not None and self._ticket.state in ("pending", "claimed", "executing"):
                return "busy", self._ticket

            if ticket.is_plan and ticket.trusted is not None:
                turn_id = ticket.trusted.get("turnOperationId", "")
                if turn_id in self._plan_turn_history:
                    return "plan_duplicate", self._ticket or ticket
                # R4g-2a phase fence per turnKey (no new disposition: plan_duplicate).
                admitted = self._plan_turn_phases.get(ticket.trusted.get("turnKey", ""))
                if admitted is not None and (
                    "execute" in admitted or ticket.trusted.get("phase") == "inspect"
                ):
                    return "plan_duplicate", self._ticket or ticket

            if ticket.is_plan and (
                len(self._plan_turn_history) >= self._plan_turn_history_limit
                or len(self._plan_replay_fences) >= self._plan_turn_history_limit
                or (
                    ticket.trusted is not None
                    and ticket.trusted.get("turnKey", "") not in self._plan_turn_phases
                    and len(self._plan_turn_phases) >= self._plan_turn_history_limit
                )
            ):
                # Admitting this new op would exceed the limit: reject it and
                # store nothing. Known ops were answered above.
                return "plan_fence_capacity", ticket

            self._ticket = ticket
            self._request_history[ticket.request_id] = ticket
            if ticket.is_plan and ticket.trusted is not None:
                turn_id = ticket.trusted.get("turnOperationId", "")
                self._plan_turn_history[turn_id] = None
                self._plan_turn_history.move_to_end(turn_id)
                turn_key = ticket.trusted.get("turnKey", "")
                self._plan_turn_phases.setdefault(turn_key, set()).add(ticket.trusted.get("phase", ""))
                self._plan_turn_phases.move_to_end(turn_key)
            self._trim_history_locked()
            return "created", ticket

    @staticmethod
    def _ticket_retained_bytes(ticket: Ticket) -> int:
        """Return a conservative UTF-8 size for payload retained for replay."""
        try:
            encoded = json.dumps(
                {
                    "ticketId": ticket.ticket_id,
                    "requestId": ticket.request_id,
                    "tool": ticket.tool,
                    "arguments": ticket.arguments,
                    "trusted": ticket.trusted,
                    "result": ticket.result_payload,
                    "state": ticket.state,
                },
                sort_keys=True,
                separators=(",", ":"),
                ensure_ascii=False,
            ).encode("utf-8")
            return len(encoded)
        except (TypeError, ValueError):
            # Protocol payloads should be JSON values. Fail closed for budgeting.
            return MAX_RETAINED_HISTORY_BYTES + 1

    def _retained_history_bytes_locked(self) -> int:
        return sum(self._ticket_retained_bytes(ticket) for ticket in self._request_history.values())

    @property
    def retained_history_bytes(self) -> int:
        with self._lock:
            return self._retained_history_bytes_locked()

    def trim_history(self) -> None:
        """Apply count and byte budgets after terminal payload state changes."""
        with self._lock:
            self._trim_history_locked()

    def _trim_history_locked(self) -> None:
        while (
            len(self._request_history) > self._history_limit
            or self._retained_history_bytes_locked() > self._history_byte_limit
        ):
            evicted = False
            for request_id, retained in list(self._request_history.items()):
                if retained is self._ticket and not retained.is_terminal():
                    continue
                self._request_history.pop(request_id, None)
                if retained.is_plan:
                    trusted = retained.trusted or {}
                    self._plan_replay_fences[request_id] = (
                        trusted.get("turnOperationId", ""),
                        trusted.get("workbookPlanFingerprint", ""),
                        retained.request_digest or "invalid",
                        trusted.get("phase", ""),
                        trusted.get("turnKey", ""),
                    )
                    self._plan_replay_fences.move_to_end(request_id)
                    # Never evicted: every fence key is in _plan_turn_history,
                    # which begin_ticket keeps within _plan_turn_history_limit.
                evicted = True
                break
            if not evicted:
                # The only remaining record is the active non-terminal ticket.
                break

    def set_ticket(self, ticket: Ticket) -> bool:
        """Backward-compatible wrapper for tests that create one ticket."""
        status, _ = self.begin_ticket(ticket)
        return status == "created"

    def clear_ticket(self) -> None:
        """Remove completed/cancelled ticket (used to force-clear after retention window)."""
        with self._lock:
            if self._ticket and self._ticket.is_terminal():
                self._ticket = None

    def get_ticket(self) -> Ticket | None:
        """Get current ticket regardless of state."""
        with self._lock:
            return self._ticket

    def find_ticket(self, ticket_id: str) -> Ticket | None:
        """Find a retained ticket by UUID for duplicate-result idempotency."""
        with self._lock:
            for ticket in self._request_history.values():
                if ticket.ticket_id == ticket_id:
                    return ticket
            return None

    def sweep_expired(self, clock: Callable[[], float] | None = None) -> list[Ticket]:
        """Sweep expired leases and invoke timeouts. Returns list of failed tickets."""
        now_val = self._now(clock)
        results: list[Ticket] = []
        with self._lock:
            ticket = self._ticket
            if ticket is None:
                return results

            # Check invoke timeout first
            if ticket.is_invoke_expired(clock):
                ticket.timeout(clock)
                results.append(ticket)
                return results

            # A plan that crossed execution is never requeued. Lease expiry
            # records uncertainty and wakes the waiting plugin request.
            if ticket.is_plan and ticket.state == "executing" and ticket.is_lease_expired(clock):
                ticket.timeout(clock)
                results.append(ticket)
                return results

            # Claimed-but-not-executing primitive work may be recovered once;
            # a claimed plan becomes terminal expired_claimed (never requeued).
            if ticket.is_lease_expired(clock):
                requeued = ticket.requeue(clock)
                if not requeued:
                    results.append(ticket)

        return results


class SessionManager:
    """Thread-safe session registry with bounded capacity."""

    def __init__(self, clock: Callable[[], float] | None = None,
                 max_sessions: int = MAX_SESSIONS) -> None:
        self._sessions: dict[str, Session] = {}
        self._lock = threading.Lock()
        self._clock = clock or _DEFAULT_CLOCK
        self._max_sessions = max_sessions

    def _now(self) -> float:
        return self._clock()

    @property
    def max_sessions(self) -> int:
        return self._max_sessions

    def create(
        self,
        document_id: str,
        pane_id: str,
        session_key: str,
        workbook_plan_fingerprint: str | None = None,
    ) -> Session:
        """Create a new session. Raises RuntimeError if at capacity."""
        session_id = generate_uuid()
        now = self._now()
        session = Session(
            session_id=session_id,
            document_id=document_id,
            pane_id=pane_id,
            session_key=session_key,
            workbook_plan_fingerprint=workbook_plan_fingerprint,
            last_heartbeat=now,
            _clock=self._clock,
        )
        with self._lock:
            if len(self._sessions) >= self._max_sessions:
                raise RuntimeError("Session capacity reached")
            self._sessions[session_id] = session
        return session

    def get(self, session_id: str) -> Session | None:
        """Get session by ID."""
        with self._lock:
            return self._sessions.get(session_id)

    def delete(self, session_id: str) -> bool:
        """Delete session. Returns True if it existed."""
        with self._lock:
            if session_id in self._sessions:
                sess = self._sessions.pop(session_id)
                ticket = sess.get_ticket()
                if ticket and ticket.state in ("pending", "claimed", "executing"):
                    ticket.cancel()
                return True
            return False

    def delete_for_pane(self, session_id: str, pane_id: str) -> bool:
        """Delete session only if pane_id matches."""
        with self._lock:
            sess = self._sessions.get(session_id)
            if sess and constant_time_compare(sess.pane_id, pane_id):
                del self._sessions[session_id]
                ticket = sess.get_ticket()
                if ticket and ticket.state in ("pending", "claimed", "executing"):
                    ticket.cancel()
                return True
            return False

    def sweep(self, clock: Callable[[], float] | None = None) -> list[Ticket]:
        """Sweep all sessions: expire stale sessions and lease/invoke timeouts."""
        clock_fn = clock or self._clock
        failed: list[Ticket] = []

        with self._lock:
            expired_sessions = [
                sid for sid, s in self._sessions.items()
                if s.is_expired(clock_fn)
            ]
            for sid in expired_sessions:
                sess = self._sessions.pop(sid)
                ticket = sess.get_ticket()
                if ticket and ticket.state in ("pending", "claimed", "executing"):
                    ticket.timeout(clock_fn)
                    failed.append(ticket)

            for session in list(self._sessions.values()):
                failed.extend(session.sweep_expired(clock_fn))

        return failed

    def shutdown(self) -> list[Ticket]:
        """Cancel all active sessions and tickets. Returns list of cancelled tickets."""
        failed: list[Ticket] = []
        with self._lock:
            for session in list(self._sessions.values()):
                ticket = session.get_ticket()
                if ticket and ticket.state in ("pending", "claimed", "executing"):
                    ticket.cancel()
                    failed.append(ticket)
            self._sessions.clear()
        return failed

    @property
    def active_count(self) -> int:
        """Number of active sessions."""
        with self._lock:
            return len(self._sessions)

    def get_metrics(self) -> dict[str, Any]:
        """Return metrics snapshot (counts only, no content)."""
        with self._lock:
            pending = 0
            claimed = 0
            executing = 0
            completed = 0
            terminal = 0
            cancelled = 0
            timed_out = 0
            for s in self._sessions.values():
                t = s.get_ticket()
                if t is None:
                    continue
                if t.state == "pending":
                    pending += 1
                elif t.state == "claimed":
                    claimed += 1
                elif t.state == "executing":
                    executing += 1
                elif t.state == "completed":
                    completed += 1
                elif t.state == "terminal":
                    terminal += 1
                elif t.state == "cancelled":
                    cancelled += 1
                elif t.state == "timeout":
                    timed_out += 1

        return {
            "sessionsActive": len(self._sessions),
            "ticketsPending": pending,
            "ticketsClaimed": claimed,
            "ticketsExecuting": executing,
            "ticketsCompleted": completed,
            "ticketsTerminal": terminal,
            "ticketsCancelled": cancelled,
            "ticketsTimedOut": timed_out,
        }
