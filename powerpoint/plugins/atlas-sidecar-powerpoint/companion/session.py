"""Session and ticket state machine for the PowerPoint companion RPC server.

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
"""
from __future__ import annotations

import json
import threading
import time
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Any, Callable, Union

from .protocol import (
    CLAIM_LEASE_SECONDS,
    INVOKE_TIMEOUT_SECONDS,
    MAX_REQUEST_HISTORY,
    MAX_RETAINED_HISTORY_BYTES,
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


@dataclass
class Ticket:
    """A single invoke ticket in a session's queue (depth 1)."""

    ticket_id: str
    session_id: str
    request_id: str
    tool: str
    arguments: dict[str, Any]

    # State machine
    state: str = "pending"  # pending | claimed | completed | cancelled | timeout
    claimed_by_pane_id: str | None = None
    claimed_at: float = 0.0
    requeue_count: int = 0

    # Result — any JSON value (not forced to dict)
    result_payload: Any = None
    completed_at: float = 0.0

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

    # ------------------------------------------------------------------
    # Clock-aware helpers
    # ------------------------------------------------------------------

    def _now(self, clock: ClockOrNow = None) -> float:
        return _resolve_now(clock, self._clock)

    def wait_for_result(self, timeout: float | None = None) -> bool:
        """Block until terminal state; return False when the wait deadline expires."""
        return self._result_event.wait(timeout)

    def set_result(self, payload: Any, clock: Callable[[], float] | None = None) -> None:
        """Set the result (any JSON value) and wake blocked waiter."""
        with self._lock:
            if self.state in ("completed", "cancelled", "timeout"):
                return
            self.state = "completed"
            self.result_payload = payload
            self.completed_at = self._now(clock)
        self._result_event.set()

    def submit_result(self, payload: Any, pane_id: str,
                      clock: Callable[[], float] | None = None) -> str:
        """Atomically accept a result for a claim or one-time requeued ticket."""
        with self._lock:
            if self.state == "completed":
                return "completed"
            if self.state == "claimed":
                if not constant_time_compare(self.claimed_by_pane_id or "__dummy__", pane_id):
                    return "wrong_pane"
            elif not (self.state == "pending" and self.requeue_count == 1):
                return f"invalid:{self.state}"
            self.state = "completed"
            self.result_payload = payload
            self.completed_at = self._now(clock)
        self._result_event.set()
        return "accepted"

    def cancel(self) -> None:
        """Cancel this ticket."""
        with self._lock:
            if self.state in ("completed", "cancelled", "timeout"):
                return
            self.state = "cancelled"
        self._result_event.set()

    def timeout(self, clock: Callable[[], float] | None = None) -> None:
        """Mark as timed out."""
        with self._lock:
            if self.state in ("completed", "cancelled", "timeout"):
                return
            self.state = "timeout"
        self._result_event.set()

    def claim(self, pane_id: str, clock: Callable[[], float] | None = None) -> bool:
        """Atomically claim this ticket for a pane. Returns True if claimed."""
        with self._lock:
            if self.state != "pending":
                return False
            self.state = "claimed"
            self.claimed_by_pane_id = pane_id
            self.claimed_at = self._now(clock)
            return True

    def requeue(self, clock: Callable[[], float] | None = None) -> bool:
        """Requeue claimed ticket back to pending. Returns True if requeued."""
        with self._lock:
            if self.state != "claimed":
                return False
            if self.requeue_count >= 1:
                self.state = "timeout"
                self._result_event.set()
                return False
            self.state = "pending"
            self.claimed_by_pane_id = None
            self.claimed_at = 0.0
            self.requeue_count += 1
            return True

    def is_lease_expired(self, clock: Callable[[], float] | None = None) -> bool:
        """Check if the claim lease has expired."""
        if self.state != "claimed":
            return False
        return (self._now(clock) - self.claimed_at) >= CLAIM_LEASE_SECONDS

    def is_invoke_expired(self, clock: Callable[[], float] | None = None) -> bool:
        """Check if the overall invoke has timed out."""
        if self.state in ("completed", "cancelled", "timeout"):
            return False
        return (self._now(clock) - self.created_at) >= self.invoke_timeout_seconds

    def is_terminal(self) -> bool:
        """Return True if ticket is in a terminal state."""
        with self._lock:
            return self.state in ("completed", "cancelled", "timeout")


@dataclass
class Session:
    """A single PowerPoint document session."""

    session_id: str
    document_id: str
    pane_id: str
    session_key: str
    last_heartbeat: float = 0.0
    _clock: Callable[[], float] = field(default=_DEFAULT_CLOCK, repr=False, compare=False)

    # One ticket at a time; completed retained for bounded request replay
    _ticket: Ticket | None = None
    _lock: threading.Lock = field(default_factory=threading.Lock)
    _request_history: OrderedDict[str, Ticket] = field(default_factory=OrderedDict)
    _history_limit: int = MAX_REQUEST_HISTORY
    _history_byte_limit: int = MAX_RETAINED_HISTORY_BYTES

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
            return self._ticket is not None and self._ticket.state in ("pending", "claimed")

    def get_pending_ticket(self) -> Ticket | None:
        """Get the pending ticket if any."""
        with self._lock:
            if self._ticket and self._ticket.state == "pending":
                return self._ticket
            return None

    def begin_ticket(self, ticket: Ticket) -> tuple[str, Ticket]:
        """Atomically create, reattach, or reject a request.

        Returns ``created``, ``existing``, ``busy``, or ``conflict`` with the
        authoritative ticket. Bounded history prevents retries from minting a
        second Office mutation.
        """
        with self._lock:
            existing = self._request_history.get(ticket.request_id)
            if existing is not None:
                if existing.tool != ticket.tool or existing.arguments != ticket.arguments:
                    return "conflict", existing
                self._request_history.move_to_end(ticket.request_id)
                return "existing", existing

            if self._ticket is not None and self._ticket.state in ("pending", "claimed"):
                return "busy", self._ticket

            self._ticket = ticket
            self._request_history[ticket.request_id] = ticket
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

            # Check lease expiry
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

    def create(self, document_id: str, pane_id: str, session_key: str) -> Session:
        """Create a new session. Raises RuntimeError if at capacity."""
        session_id = generate_uuid()
        now = self._now()
        session = Session(
            session_id=session_id,
            document_id=document_id,
            pane_id=pane_id,
            session_key=session_key,
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
                if ticket and ticket.state in ("pending", "claimed"):
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
                if ticket and ticket.state in ("pending", "claimed"):
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
                if ticket and ticket.state in ("pending", "claimed"):
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
                if ticket and ticket.state in ("pending", "claimed"):
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
            completed = 0
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
                elif t.state == "completed":
                    completed += 1
                elif t.state == "cancelled":
                    cancelled += 1
                elif t.state == "timeout":
                    timed_out += 1

        return {
            "sessionsActive": len(self._sessions),
            "ticketsPending": pending,
            "ticketsClaimed": claimed,
            "ticketsCompleted": completed,
            "ticketsCancelled": cancelled,
            "ticketsTimedOut": timed_out,
        }
