import {
  ServerOwnedTurnBindingStore,
  WorkbookPlanTurnActivationGuard,
  type WorkbookPlanConfirmationBinding,
  type ServerOwnedPaneTurn,
  type ServerOwnedPlanPhase,
  type ServerOwnedTurnBinding,
  type ServerOwnedTicketIdentity,
} from "../workbookPlanOrchestrator";
import type { TrustedConfirmationEvent, TrustedWorkbookRequest } from "../workbookPlanRequest";
import { createPrimitiveMutationAuthority, type PrimitiveMutationAuthority } from "../primitiveMutationAuthority";

const DOCUMENT_IDENTITY_UNAVAILABLE_REASON = "document-identity-unavailable";
const BRIDGE_RECONNECT_FAILED_REASON = "bridge-reconnect-failed";
/** Bound on outstanding Hermes-request marks; a full table stays tainted (fail-closed). */
const MAX_HERMES_REQUEST_MARKS = 64;

export interface TaskPaneTurnHandle {
  readonly generation: number;
  readonly turnId: string;
  readonly controller: AbortController;
  readonly trustedRequest: TrustedWorkbookRequest | null;
  readonly turnStart: number | null;
  readonly cancellationSignal: AbortSignal;
}

export interface TaskPaneTurnStartContext {
  readonly trustedRequest: TrustedWorkbookRequest;
  readonly turnStart: number;
  readonly cancellationSignal: AbortSignal;
  /** Rollout snapshot for this turn's primitive mutation authority (R4c). */
  readonly workbookPlanRolloutEnabled?: boolean;
}

export interface TaskPaneBridgeIdentity {
  readonly sessionId: string | null;
  readonly documentId: string | null;
}

export interface TaskPaneBridgeSession {
  readonly sessionId: string;
  readonly documentId: string;
}

/** Bridge state as published to the task pane (structurally `BridgePublicState`). */
export interface TaskPaneBridgeState {
  readonly connected: boolean;
  readonly sessionId: string | null;
  readonly documentId: string | null;
}

/** The task pane's active-turn identity as seen by `updateBridgeState`. */
export interface TaskPaneActiveTurnIdentity {
  readonly turnId: string;
  readonly sessionId: string | null;
  readonly documentId: string | null;
}

/**
 * One Hermes request issued on a bridge session. It is unclean until
 * `endClean()` proves the request resolved; an uncleaned request taints its
 * session so the next rollout-enabled turn rotates away from it.
 */
export interface TaskPaneHermesRequest {
  endClean(): void;
}

export interface TaskPaneTurnCoordinatorDependencies {
  readonly getOrCreateDocumentId: () => Promise<string>;
  readonly getBridgeIdentity: () => TaskPaneBridgeIdentity;
  readonly reconnectForDocumentChange: (expectedDocumentId: string) => Promise<TaskPaneBridgeSession>;
  readonly bindings: ServerOwnedTurnBindingStore;
  readonly paneId: string;
  readonly now?: () => number;
}

interface PendingWorkbookPlanConfirmation {
  readonly binding: WorkbookPlanConfirmationBinding;
  readonly resolve: (event: unknown) => void;
  readonly reject: (error: unknown) => void;
  readonly cleanup: () => void;
}

interface PaneConfirmationEventRecord {
  readonly binding: WorkbookPlanConfirmationBinding;
  readonly trusted: TrustedConfirmationEvent;
  readonly owner: symbol;
  valid: boolean;
}

const paneConfirmationEvents = new WeakMap<object, PaneConfirmationEventRecord>();

export type TaskPaneTurnPreparation =
  | { readonly status: "active"; readonly handle: TaskPaneTurnHandle; readonly sessionId: string | null; readonly documentId: string }
  | { readonly status: "stale"; readonly handle: TaskPaneTurnHandle }
  | { readonly status: "rejected"; readonly handle: TaskPaneTurnHandle; readonly reason: string };

export type TaskPaneServerAuthorization =
  | { readonly ok: false; readonly reason: string }
  | TaskPaneAuthorizedServerOperation;

export interface TaskPaneAuthorizedServerOperation {
  readonly ok: true;
  readonly identity: ServerOwnedTicketIdentity;
  readonly binding: ServerOwnedTurnBinding;
  readonly snapshot: TaskPaneAuthorizedServerSnapshot;
  readonly assertLive: () => void;
}

/** R4g-1: the trusted phase of one server ticket (absent on legacy, phase-less tickets). */
export interface TaskPaneServerPhase {
  readonly phase: ServerOwnedPlanPhase;
  readonly turnKey: string;
}

export interface TaskPaneAuthorizedServerSnapshot {
  readonly generation: number;
  readonly turnId: string;
  readonly controller: AbortController;
  readonly cancellationSignal: AbortSignal;
  readonly sessionId: string;
  readonly documentId: string;
  readonly paneId: string;
  readonly serverOperationId: string;
  readonly turnStart: number | null;
  readonly trustedRequest: TrustedWorkbookRequest | null;
  /** R4g-1: present only for a phase ticket. */
  readonly phase?: ServerOwnedPlanPhase;
}

export type TaskPaneTurnContinuationResult<T> =
  | TaskPaneTurnPreparation
  | ({ readonly status: "active"; readonly handle: TaskPaneTurnHandle; readonly sessionId: string | null; readonly documentId: string; readonly value: T });

export type TaskPaneServerAuthorizationResult<T> =
  | Extract<TaskPaneServerAuthorization, { readonly ok: false }>
  | (TaskPaneAuthorizedServerOperation & { readonly result: T });

interface PreparedTurn {
  readonly handle: TaskPaneTurnHandle;
  readonly sessionId: string | null;
  readonly documentId: string;
}

/**
 * Production task-pane/server coordination seam. Office identity preload and
 * server-ticket authorization share one generation-owned turn boundary so a
 * late preload cannot activate state for a newer turn.
 */
export class TaskPaneTurnCoordinator {
  private readonly activationGuard = new WorkbookPlanTurnActivationGuard();
  private current: {
    readonly handle: TaskPaneTurnHandle;
    prepared: PreparedTurn | null;
    readonly authority: PrimitiveMutationAuthority;
  } | null = null;
  private pendingConfirmation: PendingWorkbookPlanConfirmation | null = null;
  private activeConfirmationEvent: object | null = null;
  private activeServerOperationId: string | null = null;
  private readonly confirmationOwner = Symbol("pane-confirmation-owner");
  /** Outstanding (unclean) Hermes-request marks, keyed by mark, valued by bridge session. */
  private readonly hermesRequestMarks = new Map<symbol, string>();
  /** Sessions that issued a request while the mark table was full: tainted until rotated away from. */
  private readonly overflowTaintedSessions = new Set<string>();
  /** Set only while this coordinator's prepareTurn reconnects for this handle. */
  private rotatingHandle: TaskPaneTurnHandle | null = null;
  /** R4f (A4): the reconnect promise of the rotation owned by `rotatingHandle`. */
  private rotationInFlight: Promise<unknown> | null = null;
  /** R4f (A4): the current unprepared turn while its prepareTurn waits out another handle's rotation. */
  private awaitingRotationHandle: TaskPaneTurnHandle | null = null;
  /** R4f (A6, D-R4-8): a failed coordinator reconnect arms one attempt for the next sessionless prepare. */
  private reconnectRetryArmed = false;

  private readonly now: () => number;

  constructor(private readonly dependencies: TaskPaneTurnCoordinatorDependencies) {
    this.now = dependencies.now ?? (() => Date.now());
  }

  beginTurn(turnId: string, controller: AbortController, context?: TaskPaneTurnStartContext): TaskPaneTurnHandle {
    if (this.current !== null) this.stopCurrentTurn();
    const generation = this.activationGuard.begin(turnId, controller);
    const handle: TaskPaneTurnHandle = {
      generation,
      turnId,
      controller,
      trustedRequest: context === undefined ? null : freezeTrustedRequest(context.trustedRequest),
      turnStart: context?.turnStart ?? null,
      cancellationSignal: context?.cancellationSignal ?? controller.signal,
    };
    const authority = createPrimitiveMutationAuthority({
      trustedRequest: handle.trustedRequest,
      workbookPlanRolloutEnabled: context?.workbookPlanRolloutEnabled === true,
    });
    this.current = { handle, prepared: null, authority };
    return handle;
  }

  /** R4c: this turn's primitive mutation authority, or null once the turn was finished, stopped or replaced. */
  primitiveAuthorityFor(handle: TaskPaneTurnHandle): PrimitiveMutationAuthority | null {
    if (!this.isOwnedHandle(handle)) return null;
    const authority = this.current!.authority;
    return authority.isTerminal ? null : authority;
  }

  /** R4c: latch the turn's authority once a direct WorkbookPlan turn has been entered. */
  markWorkbookPlanEntered(handle: Pick<TaskPaneTurnHandle, "generation" | "turnId" | "controller">): void {
    if (!this.isOwnedHandle(handle as TaskPaneTurnHandle)) return;
    this.current!.authority.markWorkbookPlanEntered();
  }

  async prepareTurn(handle: TaskPaneTurnHandle): Promise<TaskPaneTurnPreparation> {
    let documentId: string;
    try {
      documentId = await this.dependencies.getOrCreateDocumentId();
    } catch {
      if (!this.isCurrentHandle(handle)) return { status: "stale", handle };
      return {
        status: "rejected",
        handle,
        reason: DOCUMENT_IDENTITY_UNAVAILABLE_REASON,
      };
    }
    if (!this.isCurrentHandle(handle)) return { status: "stale", handle };

    if (this.rotationInFlight !== null && this.rotatingHandle !== handle) {
      // R4f (A4): another handle's rotation is in flight. Wait it out, raced
      // against this turn's own signal, exempt from force-stop meanwhile, then
      // re-verify; never go active sessionless because of it.
      const rotation = this.rotationInFlight;
      const signal = handle.cancellationSignal;
      let release: () => void = () => undefined;
      const onAbort = (): void => release();
      this.awaitingRotationHandle = handle;
      signal.addEventListener("abort", onAbort, { once: true });
      try {
        await new Promise<void>((resolve) => {
          release = resolve;
          if (signal.aborted) resolve();
          rotation.then(() => resolve(), () => resolve());
        });
      } finally {
        signal.removeEventListener("abort", onAbort);
        if (this.awaitingRotationHandle === handle) this.awaitingRotationHandle = null;
      }
      if (!this.isCurrentHandle(handle)) return { status: "stale", handle };
    }

    let bridge = this.dependencies.getBridgeIdentity();
    this.pruneForeignSessionTaint(bridge.sessionId);
    let sessionId = bridge.sessionId;
    if (sessionId !== null) this.reconnectRetryArmed = false;
    const needsRotation = sessionId !== null && (bridge.documentId !== documentId || this.isSessionTainted(sessionId));
    const needsRetry = sessionId === null && this.reconnectRetryArmed;
    if (needsRotation || needsRetry) {
      const rotatedFromSessionId = sessionId;
      if (rotatedFromSessionId !== null) this.dependencies.bindings.invalidateSession(rotatedFromSessionId);
      if (needsRetry) this.reconnectRetryArmed = false;
      let session: TaskPaneBridgeSession;
      let reconnect: Promise<TaskPaneBridgeSession> | null = null;
      this.rotatingHandle = handle;
      try {
        reconnect = this.dependencies.reconnectForDocumentChange(documentId);
        this.rotationInFlight = reconnect;
        session = await reconnect;
      } catch {
        // Armed before the currency check: a rotation whose turn was stopped
        // still gives the next sessionless prepare its one attempt.
        this.reconnectRetryArmed = true;
        if (!this.isCurrentHandle(handle)) return { status: "stale", handle };
        return {
          status: "rejected",
          handle,
          reason: BRIDGE_RECONNECT_FAILED_REASON,
        };
      } finally {
        if (this.rotatingHandle === handle) this.rotatingHandle = null;
        if (reconnect !== null && this.rotationInFlight === reconnect) this.rotationInFlight = null;
      }
      // The old session was stopped and DELETEd and a new one started: a
      // stale Hermes turn on it can no longer reach this pane.
      if (rotatedFromSessionId !== null) this.clearSessionTaint(rotatedFromSessionId);
      if (!this.isCurrentHandle(handle)) return { status: "stale", handle };
      bridge = this.dependencies.getBridgeIdentity();
      if (
        session.documentId !== documentId ||
        bridge.sessionId !== session.sessionId ||
        bridge.documentId !== documentId
      ) {
        this.dependencies.bindings.invalidateSession(session.sessionId);
        return { status: "rejected", handle, reason: "Bridge session document identity mismatch." };
      }
      sessionId = session.sessionId;
    }

    if (!this.isCurrentHandle(handle)) return { status: "stale", handle };
    bridge = this.dependencies.getBridgeIdentity();
    if (sessionId !== null && (bridge.sessionId !== sessionId || bridge.documentId !== documentId)) {
      this.dependencies.bindings.invalidateSession(sessionId);
      return { status: "rejected", handle, reason: "Current document does not match the authenticated bridge session." };
    }

    const prepared = { handle, sessionId, documentId };
    if (sessionId !== null) {
      const turn: ServerOwnedPaneTurn = {
        sessionId,
        documentId,
        paneId: this.dependencies.paneId,
        turnId: handle.turnId,
      };
      this.dependencies.bindings.activateTurn(turn);
    }
    // isCurrentHandle(handle) holds synchronously here, so `current` is this turn: keep its authority.
    this.current = { handle, prepared, authority: this.current!.authority };
    return { status: "active", handle, sessionId, documentId };
  }

  /**
   * Marks a Hermes request on `sessionId` synchronously, before it is sent, in
   * every model mode. The mark is unclean until `endClean()` (resolved path
   * only); any other end (Stop, deadline, error, forced stop, throw) leaves the
   * session tainted. Only a successful rotation away from the session clears
   * its marks. A null session takes no mark.
   */
  beginHermesRequest(sessionId: string | null): TaskPaneHermesRequest {
    if (sessionId === null) return Object.freeze({ endClean: () => undefined });
    if (this.hermesRequestMarks.size >= MAX_HERMES_REQUEST_MARKS) {
      this.overflowTaintedSessions.add(sessionId);
      return Object.freeze({ endClean: () => undefined });
    }
    const mark = Symbol("hermes-request");
    this.hermesRequestMarks.set(mark, sessionId);
    return Object.freeze({ endClean: () => { this.hermesRequestMarks.delete(mark); } });
  }

  /** Count of outstanding Hermes-request marks (bounded at 64). */
  get hermesRequestMarkCount(): number {
    return this.hermesRequestMarks.size;
  }

  /**
   * Task-pane force-stop policy for a published bridge state. Reproduces the
   * task pane's disconnect / identity-change rule, except while this
   * coordinator's own prepareTurn is reconnecting for the current, still
   * unprepared turn: that rotation must not abort the turn it is preparing
   * (prepareTurn re-verifies the identity after the reconnect).
   */
  bridgeStateInvalidatesActiveTurn(
    state: TaskPaneBridgeState,
    active: TaskPaneActiveTurnIdentity | null,
    previous: TaskPaneBridgeIdentity,
  ): boolean {
    if (active === null) return false;
    const identityChanged = previous.sessionId !== state.sessionId || previous.documentId !== state.documentId;
    const invalidated = !state.connected ||
      active.sessionId !== state.sessionId ||
      (active.documentId !== null && active.documentId !== state.documentId) ||
      identityChanged;
    if (!invalidated) return false;
    return !this.isOwnRotationWindow(active.turnId);
  }

  authorizeServerOperation(serverOperationId: string, serverPhase?: TaskPaneServerPhase): TaskPaneServerAuthorization {
    const current = this.current;
    if (current === null || current.prepared === null || !this.isCurrentHandle(current.handle)) {
      return { ok: false, reason: "Trusted server turn identity is unavailable." };
    }
    const bridge = this.dependencies.getBridgeIdentity();
    if (
      current.prepared.sessionId === null ||
      bridge.sessionId !== current.prepared.sessionId ||
      bridge.documentId !== current.prepared.documentId
    ) {
      if (current.prepared.sessionId !== null) this.dependencies.bindings.invalidateSession(current.prepared.sessionId);
      current.prepared = null;
      return { ok: false, reason: "Trusted server document identity does not match the authenticated bridge session." };
    }
    const identity: ServerOwnedTicketIdentity = {
      sessionId: current.prepared.sessionId,
      documentId: current.prepared.documentId,
      paneId: this.dependencies.paneId,
      turnId: current.handle.turnId,
      serverOperationId,
      ...(serverPhase === undefined ? {} : { phase: serverPhase.phase, turnKey: serverPhase.turnKey }),
    };
    const bound = this.dependencies.bindings.bindServerTicket(identity);
    if (bound.ok !== true) return bound;
    this.activeServerOperationId = identity.serverOperationId;
    current.authority.markWorkbookPlanEntered();
    const snapshot = Object.freeze({
      generation: current.handle.generation,
      turnId: current.handle.turnId,
      controller: current.handle.controller,
      cancellationSignal: current.handle.cancellationSignal,
      sessionId: identity.sessionId,
      documentId: identity.documentId,
      paneId: identity.paneId,
      serverOperationId: identity.serverOperationId,
      turnStart: current.handle.turnStart,
      trustedRequest: current.handle.trustedRequest,
      ...(serverPhase === undefined ? {} : { phase: serverPhase.phase }),
    });
    const authorized: TaskPaneAuthorizedServerOperation = {
      ok: true,
      identity,
      binding: bound.binding,
      snapshot,
      assertLive: () => this.assertAuthorizedOperationLive(identity, bound.binding, snapshot),
    };
    return authorized;
  }

  waitForWorkbookPlanConfirmation(
    binding: WorkbookPlanConfirmationBinding,
    signal?: AbortSignal,
  ): Promise<unknown> {
    const current = this.current;
    const effectiveSignal = signal ?? current?.handle.cancellationSignal;
    if (binding.deadline <= this.now()) {
      return Promise.reject(new DOMException("The pending workbook action expired.", "TimeoutError"));
    }
    if (!this.isConfirmationBindingCurrent(binding) || effectiveSignal === undefined) {
      return Promise.reject(new DOMException("The pending workbook action is no longer active.", "AbortError"));
    }
    if (this.pendingConfirmation !== null) {
      return Promise.reject(new DOMException("A pending workbook action already exists.", "AbortError"));
    }
    const remaining = binding.deadline - this.now();
    if (remaining <= 0) {
      return Promise.reject(new DOMException("The pending workbook action expired.", "TimeoutError"));
    }

    return new Promise<unknown>((resolve, reject) => {
      let settled = false;
      const timer = setTimeout(() => {
        finish(() => reject(new DOMException("The pending workbook action expired.", "TimeoutError")));
      }, remaining);
      const cleanup = (): void => {
        clearTimeout(timer);
        effectiveSignal.removeEventListener("abort", onAbort);
      };
      const finish = (callback: () => void): void => {
        if (settled) return;
        settled = true;
        cleanup();
        if (this.pendingConfirmation?.binding === binding) this.pendingConfirmation = null;
        callback();
      };
      const onAbort = (): void => {
        finish(() => reject(
          new DOMException(
            "The pending workbook action was stopped.",
            this.now() >= binding.deadline ? "TimeoutError" : "AbortError",
          ),
        ));
      };
      this.pendingConfirmation = { binding, resolve: (event) => finish(() => resolve(event)), reject: (error) => finish(() => reject(error)), cleanup };
      effectiveSignal.addEventListener("abort", onAbort, { once: true });
      if (effectiveSignal.aborted) onAbort();
    });
  }

  continuePendingWorkbookPlan(): boolean {
    const pending = this.pendingConfirmation;
    if (pending === null) return false;
    if (pending.binding.deadline <= this.now()) {
      this.cancelPendingConfirmation(new DOMException("The pending workbook action expired.", "TimeoutError"));
      return false;
    }
    if (!this.isConfirmationBindingCurrent(pending.binding)) {
      this.cancelPendingConfirmation(new DOMException("The pending workbook action is no longer active.", "AbortError"));
      return false;
    }
    pending.cleanup();
    this.pendingConfirmation = null;
    const event = Object.freeze({
      eventId: crypto.randomUUID(),
      confirmed: true as const,
      binding: Object.freeze({ ...pending.binding }),
    });
    const record: PaneConfirmationEventRecord = {
      binding: pending.binding,
      trusted: Object.freeze({ eventId: event.eventId, confirmed: true as const, source: "user" as const }),
      owner: this.confirmationOwner,
      valid: true,
    };
    paneConfirmationEvents.set(event, record);
    this.activeConfirmationEvent = event;
    pending.resolve(event);
    return true;
  }

  consumeWorkbookPlanConfirmation(
    candidate: unknown,
    binding: WorkbookPlanConfirmationBinding,
  ): TrustedConfirmationEvent | undefined {
    if (candidate === null || typeof candidate !== "object") return undefined;
    const record = paneConfirmationEvents.get(candidate);
    if (record === undefined || record.owner !== this.confirmationOwner || !record.valid || !sameConfirmationBinding(record.binding, binding)) return undefined;
    if (!this.isConfirmationBindingCurrent(binding)) return undefined;
    record.valid = false;
    if (this.activeConfirmationEvent === candidate) this.activeConfirmationEvent = null;
    return record.trusted;
  }

  cancelPendingConfirmation(error = new DOMException("The pending workbook action was stopped.", "AbortError")): void {
    const pending = this.pendingConfirmation;
    if (pending !== null) {
      pending.cleanup();
      this.pendingConfirmation = null;
      pending.reject(error);
    }
    this.invalidateActiveConfirmationEvent();
  }

  async continueTurn<T>(handle: TaskPaneTurnHandle, continuation: (preparation: Extract<TaskPaneTurnPreparation, { readonly status: "active" }>) => Promise<T> | T): Promise<TaskPaneTurnContinuationResult<T>> {
    const preparation = await this.prepareTurn(handle);
    if (preparation.status !== "active") return preparation;
    if (!this.isCurrentHandle(handle)) return { status: "stale", handle };
    const value = await continuation(preparation);
    if (!this.isCurrentHandle(handle)) return { status: "stale", handle };
    return { ...preparation, value };
  }

  async authorizeAndRun<T>(serverOperationId: string, continuation: (authorized: Extract<TaskPaneServerAuthorization, { readonly ok: true }>) => Promise<T> | T, serverPhase?: TaskPaneServerPhase): Promise<TaskPaneServerAuthorizationResult<T>> {
    const authorized = serverPhase === undefined
      ? this.authorizeServerOperation(serverOperationId)
      : this.authorizeServerOperation(serverOperationId, serverPhase);
    if (authorized.ok !== true) return authorized;
    const result = await continuation(authorized);
    if (!this.isAuthorizedOperationCurrent(authorized.identity, authorized.snapshot)) {
      return { ok: false, reason: "Trusted server turn is no longer active." };
    }
    return { ...authorized, result };
  }

  stopCurrentTurn(): void {
    const current = this.current;
    if (current === null) return;
    this.cancelPendingConfirmation();
    if (current.prepared !== null && current.prepared.sessionId !== null) {
      this.dependencies.bindings.clearTurn({
        sessionId: current.prepared.sessionId,
        documentId: current.prepared.documentId,
        paneId: this.dependencies.paneId,
        turnId: current.handle.turnId,
      });
    }
    this.activationGuard.invalidate();
    this.activeServerOperationId = null;
    current.authority.terminate();
    current.handle.controller.abort();
    this.current = null;
  }

  finishTurn(handle: TaskPaneTurnHandle): void {
    if (!this.isOwnedHandle(handle)) return;
    const current = this.current!;
    if (current.prepared !== null && current.prepared.sessionId !== null) {
      this.dependencies.bindings.clearTurn({
        sessionId: current.prepared.sessionId,
        documentId: current.prepared.documentId,
        paneId: this.dependencies.paneId,
        turnId: current.handle.turnId,
      });
    }
    this.cancelPendingConfirmation();
    this.activationGuard.invalidate();
    this.activeServerOperationId = null;
    current.authority.terminate();
    this.current = null;
  }

  isCurrent(generation: number, turnId: string, controller: AbortController): boolean {
    const current = this.current;
    if (current === null ||
        !this.activationGuard.isCurrent(generation, turnId, controller) ||
        current.handle.generation !== generation ||
        current.handle.turnId !== turnId ||
        current.handle.controller !== controller ||
        current.handle.cancellationSignal.aborted) {
      return false;
    }
    if (current.prepared === null) return true;
    const bridge = this.dependencies.getBridgeIdentity();
    return bridge.sessionId === current.prepared.sessionId && bridge.documentId === current.prepared.documentId;
  }

  private assertAuthorizedOperationLive(
    identity: ServerOwnedTicketIdentity,
    binding: ServerOwnedTurnBinding,
    snapshot: TaskPaneAuthorizedServerSnapshot,
  ): void {
    if (!this.isAuthorizedOperationLive(identity, binding, snapshot)) {
      throw new DOMException("The running Hermes task was stopped.", "AbortError");
    }
  }

  private isAuthorizedOperationCurrent(
    identity: ServerOwnedTicketIdentity,
    snapshot: TaskPaneAuthorizedServerSnapshot,
  ): boolean {
    const current = this.current;
    if (current === null || current.prepared === null) return false;
    if (!this.isCurrentHandle(current.handle) ||
        current.handle.generation !== snapshot.generation ||
        current.handle.turnId !== snapshot.turnId ||
        current.handle.controller !== snapshot.controller) {
      return false;
    }
    const bridge = this.dependencies.getBridgeIdentity();
    return bridge.sessionId === identity.sessionId && bridge.documentId === identity.documentId;
  }

  private isAuthorizedOperationLive(
    identity: ServerOwnedTicketIdentity,
    binding: ServerOwnedTurnBinding,
    snapshot: TaskPaneAuthorizedServerSnapshot,
  ): boolean {
    return this.isAuthorizedOperationCurrent(identity, snapshot) &&
      this.dependencies.bindings.isServerTicketLive(identity, binding);
  }

  private isSessionTainted(sessionId: string): boolean {
    if (this.overflowTaintedSessions.has(sessionId)) return true;
    for (const markedSessionId of this.hermesRequestMarks.values()) {
      if (markedSessionId === sessionId) return true;
    }
    return false;
  }

  private clearSessionTaint(sessionId: string): void {
    for (const [mark, markedSessionId] of [...this.hermesRequestMarks]) {
      if (markedSessionId === sessionId) this.hermesRequestMarks.delete(mark);
    }
    this.overflowTaintedSessions.delete(sessionId);
  }

  /** R4f (A3): drops every mark and overflow taint whose session is not the current bridge session (all when null). */
  private pruneForeignSessionTaint(currentSessionId: string | null): void {
    for (const [mark, markedSessionId] of [...this.hermesRequestMarks]) {
      if (markedSessionId !== currentSessionId) this.hermesRequestMarks.delete(mark);
    }
    for (const taintedSessionId of [...this.overflowTaintedSessions]) {
      if (taintedSessionId !== currentSessionId) this.overflowTaintedSessions.delete(taintedSessionId);
    }
  }

  private isOwnRotationWindow(turnId: string): boolean {
    const current = this.current;
    const rotating = this.rotatingHandle;
    return rotating !== null &&
      current !== null &&
      current.prepared === null &&
      current.handle.turnId === turnId &&
      this.isCurrentHandle(current.handle) &&
      (current.handle === rotating || current.handle === this.awaitingRotationHandle);
  }

  /**
   * R4d: the live turn's primitive mutation authority for a server primitive
   * ticket on `bridgeSessionId`, or null. A rollout-on turn must be prepared on
   * that session and the current document; a rollout-off turn is never prepared.
   */
  serverPrimitiveAuthority(input: { readonly turnId: string; readonly bridgeSessionId: string }): PrimitiveMutationAuthority | null {
    const current = this.current;
    if (current === null || current.handle.turnId !== input.turnId) return null;
    const bridge = this.dependencies.getBridgeIdentity();
    const live = this.isCurrentHandle(current.handle) &&
      !current.authority.isTerminal &&
      bridge.sessionId === input.bridgeSessionId &&
      ((current.prepared !== null &&
        current.prepared.sessionId === input.bridgeSessionId &&
        bridge.documentId === current.prepared.documentId) ||
        (current.prepared === null && current.authority.rolloutEnabled === false));
    return live ? current.authority : null;
  }

  private isOwnedHandle(handle: TaskPaneTurnHandle): boolean {
    const current = this.current;
    return current !== null &&
      current.handle.generation === handle.generation &&
      current.handle.turnId === handle.turnId &&
      current.handle.controller === handle.controller;
  }

  private isCurrentHandle(handle: TaskPaneTurnHandle): boolean {
    return this.activationGuard.isCurrent(handle.generation, handle.turnId, handle.controller) &&
      !handle.cancellationSignal.aborted;
  }

  private isConfirmationBindingCurrent(binding: WorkbookPlanConfirmationBinding): boolean {
    const current = this.current;
    if (current === null || current.prepared === null || !this.isCurrentHandle(current.handle)) return false;
    if (!sameConfirmationBindingIdentity(binding, current.handle, current.prepared.documentId)) return false;
    if (binding.route === "direct") {
      if (binding.sessionId !== this.dependencies.paneId || binding.serverOperationId !== undefined) return false;
    } else if (
      binding.sessionId !== current.prepared.sessionId ||
      binding.sessionId === null ||
      binding.serverOperationId === undefined ||
      binding.serverOperationId !== this.activeServerOperationId
    ) {
      return false;
    }
    return binding.deadline > this.now();
  }

  private invalidateActiveConfirmationEvent(): void {
    if (this.activeConfirmationEvent === null) return;
    const record = paneConfirmationEvents.get(this.activeConfirmationEvent);
    if (record !== undefined) record.valid = false;
    this.activeConfirmationEvent = null;
  }
}

function sameConfirmationBindingIdentity(
  binding: WorkbookPlanConfirmationBinding,
  handle: TaskPaneTurnHandle,
  documentId: string,
): boolean {
  return binding.documentId === documentId &&
    binding.turnId === handle.turnId &&
    binding.generation === handle.generation &&
    /^[0-9a-f]{64}$/i.test(binding.planDigest) &&
    Number.isFinite(binding.deadline);
}

function sameConfirmationBinding(
  left: WorkbookPlanConfirmationBinding,
  right: WorkbookPlanConfirmationBinding,
): boolean {
  return left.planDigest === right.planDigest &&
    left.sessionId === right.sessionId &&
    left.documentId === right.documentId &&
    left.turnId === right.turnId &&
    left.generation === right.generation &&
    left.route === right.route &&
    left.deadline === right.deadline &&
    left.serverOperationId === right.serverOperationId;
}

function freezeTrustedRequest(request: TrustedWorkbookRequest): TrustedWorkbookRequest {
  const confirmation = request.confirmation === undefined
    ? undefined
    : Object.freeze({ ...request.confirmation });
  return Object.freeze({
    ...request,
    ...(confirmation === undefined ? {} : { confirmation }),
  });
}

export type { ServerOwnedTurnBinding };
