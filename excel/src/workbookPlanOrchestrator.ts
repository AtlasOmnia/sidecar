import type { ToolCall } from "./types";
import type { WorkbookPlan, InspectionOperation, VerificationRule } from "./generated/workbookPlanSchema";
import {
  normalizeWorkbookPlanResponse,
  type NormalizedWorkbookPlan,
  type WorkbookPlanProtocolResult,
} from "./workbookPlanProtocol";
import { canonicalizeWorkbookPlan } from "./workbookPlanCanonicalization";
import {
  executeAcceptedWorkbookPlan,
  getAcceptedWorkbookPlanPostconditionsForVerification,
  prepareWorkbookPlanExecution,
  type AcceptedWorkbookPlanTransaction,
  type BrowserIdempotencyStore,
  type TrustedOperationBoundary,
  type WorkbookPlanExecutionAdapter,
  type WorkbookPlanExecutionReport,
} from "./workbookPlanExecutor";
import type { VerifiedWorkbookCompileContext } from "./workbookPlanCompiler";
import type { WorkbookPlanPolicyInput } from "./workbookPlanPolicy";
import type { TrustedConfirmationEvent } from "./workbookPlanRequest";
import {
  prepareWorkbookPlanVerificationRules,
  verifyWorkbookPlan,
  type WorkbookPlanVerificationReport,
} from "./workbookPlanVerification";

export const WORKBOOK_PLAN_TURN_TIMEOUT_MS = 90_000;
export const WORKBOOK_PLAN_MAX_CORRECTIONS = 1;

export interface WorkbookPlanModelResponse {
  readonly content: string;
  readonly toolCalls: readonly ToolCall[];
}

export interface WorkbookPlanTurnIdentity {
  readonly sessionId: string;
  readonly documentId: string;
  readonly turnId: string;
}

export type WorkbookPlanTurnRoute = "direct" | "server-owned";

export interface WorkbookPlanConfirmationBinding {
  readonly planDigest: string;
  readonly sessionId: string;
  readonly documentId: string;
  readonly turnId: string;
  readonly generation: number;
  readonly route: WorkbookPlanTurnRoute;
  readonly deadline: number;
  readonly serverOperationId?: string;
}

export type WorkbookPlanTurnStatePhase = "none" | "inspected" | "pending_confirmation" | "executing" | "terminal";

/** Shared trusted state for direct continuation and server-owned tickets. */
export class WorkbookPlanTurnState {
  private currentPhase: WorkbookPlanTurnStatePhase = "none";

  get phase(): WorkbookPlanTurnStatePhase {
    return this.currentPhase;
  }

  beginInspection(): boolean {
    if (this.currentPhase !== "none") return false;
    this.currentPhase = "inspected";
    return true;
  }

  beginExecution(): boolean {
    if (this.currentPhase !== "none" && this.currentPhase !== "inspected" && this.currentPhase !== "pending_confirmation") return false;
    this.currentPhase = "executing";
    return true;
  }

  beginPendingConfirmation(): boolean {
    if (this.currentPhase !== "none" && this.currentPhase !== "inspected") return false;
    this.currentPhase = "pending_confirmation";
    return true;
  }

  resumePendingConfirmation(): boolean {
    if (this.currentPhase !== "pending_confirmation") return false;
    this.currentPhase = "none";
    return true;
  }

  finishTerminal(): void {
    this.currentPhase = "terminal";
  }
}

export interface ServerOwnedPaneTurn {
  readonly sessionId: string;
  readonly documentId: string;
  readonly paneId: string;
  readonly turnId: string;
}

/** R4g-1: trusted server ticket phase (D-R4-5 = I1). Absent on legacy (phase-less) tickets. */
export type ServerOwnedPlanPhase = "inspect" | "execute";

export interface ServerOwnedTicketIdentity extends ServerOwnedPaneTurn {
  readonly serverOperationId: string;
  /** R4g-1: present only on phase tickets, always together with `turnKey`. */
  readonly phase?: ServerOwnedPlanPhase;
  readonly turnKey?: string;
}

export interface ServerOwnedTurnBinding {
  readonly activeTurnId: string;
  readonly serverOperationId: string;
  readonly state: WorkbookPlanTurnState;
}

export type ServerOwnedTurnBindingResult =
  | { readonly ok: true; readonly reused: boolean; readonly binding: ServerOwnedTurnBinding }
  | { readonly ok: false; readonly reason: string };

export class WorkbookPlanTurnActivationGuard {
  private generation = 0;
  private active: { readonly generation: number; readonly turnId: string; readonly controller: AbortController } | null = null;

  begin(turnId: string, controller: AbortController): number {
    this.generation += 1;
    this.active = { generation: this.generation, turnId, controller };
    return this.generation;
  }

  invalidate(): void {
    this.generation += 1;
    this.active = null;
  }

  isCurrent(generation: number, turnId: string, controller: AbortController): boolean {
    return this.active?.generation === generation &&
      this.active.turnId === turnId &&
      this.active.controller === controller &&
      !controller.signal.aborted;
  }
}

const DEFAULT_SERVER_OWNED_REPLAY_MAX_BYTES = 64 * 1024;
const SERVER_PHASE_REFUSED = "Trusted server turn phase is not permitted.";

interface ServerOwnedReplayTombstone {
  readonly serverOperationId: string;
  readonly identityKey: string;
  readonly expiresAt: number;
  readonly bytes: number;
}

export class ServerOwnedTurnBindingStore {
  private readonly entries = new Map<string, {
    readonly key: string;
    readonly activeKey: string;
    readonly serverOperationId: string;
    readonly identity: ServerOwnedTicketIdentity;
    readonly binding: ServerOwnedTurnBinding;
    lastTouched: number;
    terminal: boolean;
  }>();
  private readonly activeTurns = new Map<string, ServerOwnedPaneTurn>();
  private readonly activeBindings = new Map<string, string>();
  private readonly replayTombstones = new Map<string, ServerOwnedReplayTombstone>();
  private readonly now: () => number;
  private readonly ttlMs: number;
  private readonly maxEntries: number;
  private readonly maxBytes: number;
  private replayBytes = 0;

  constructor(options: {
    readonly now: () => number;
    readonly ttlMs: number;
    readonly maxEntries: number;
    readonly maxBytes?: number;
  }) {
    this.now = options.now;
    this.ttlMs = Math.max(1, options.ttlMs);
    this.maxEntries = Math.max(1, Math.floor(options.maxEntries));
    this.maxBytes = Math.max(1, Math.floor(options.maxBytes ?? DEFAULT_SERVER_OWNED_REPLAY_MAX_BYTES));
  }

  activateTurn(turn: ServerOwnedPaneTurn): void {
    this.prune();
    if (!isServerOwnedPaneTurn(turn)) return;
    const previous = this.activeTurns.get(turn.paneId);
    if (previous !== undefined && !samePaneTurn(previous, turn)) this.clearTurn(previous);
    this.activeTurns.set(turn.paneId, { ...turn });
  }

  bindServerTicket(ticket: ServerOwnedTicketIdentity): ServerOwnedTurnBindingResult {
    this.prune();
    if (!isServerOwnedTicketIdentity(ticket)) return { ok: false, reason: "Invalid trusted server turn identity." };
    const active = this.activeTurns.get(ticket.paneId);
    if (active === undefined || !samePaneTurn(active, ticket)) {
      return { ok: false, reason: "Trusted server turn identity is unavailable." };
    }

    const key = serverOwnedTicketKey(ticket);
    const existing = this.entries.get(key);
    if (existing !== undefined) {
      if (existing.identity.phase !== ticket.phase || existing.identity.turnKey !== ticket.turnKey) {
        return { ok: false, reason: SERVER_PHASE_REFUSED };
      }
      if (existing.terminal) return { ok: false, reason: "Trusted server turn is terminal." };
      existing.lastTouched = this.now();
      return { ok: true, reused: true, binding: existing.binding };
    }

    const activeKey = serverOwnedPaneTurnKey(active);
    if (ticket.phase === undefined && this.activeBindings.has(activeKey)) {
      return { ok: false, reason: "A different server turn is already bound." };
    }
    if (this.replayTombstones.has(ticket.serverOperationId)) {
      return { ok: false, reason: "Trusted server operation was replayed." };
    }
    const replayBytes = serverOwnedReplayBytes(ticket);
    if (this.replayTombstones.size >= this.maxEntries || this.replayBytes > this.maxBytes - replayBytes) {
      return { ok: false, reason: "Trusted server turn capacity is unavailable." };
    }
    // R4g-1: phase rules run after every legacy check above. A phase ticket gets
    // a fresh state, or (execute after a delivered inspection with the same
    // turnKey) the inspection's state object; anything else is refused.
    const phaseState = ticket.phase === undefined ? undefined : this.phaseBindingState(ticket, activeKey);
    if (phaseState === null) return { ok: false, reason: SERVER_PHASE_REFUSED };

    const now = this.now();
    const binding: ServerOwnedTurnBinding = {
      activeTurnId: ticket.turnId,
      serverOperationId: ticket.serverOperationId,
      state: phaseState ?? new WorkbookPlanTurnState(),
    };
    this.entries.set(key, {
      key,
      activeKey,
      serverOperationId: ticket.serverOperationId,
      identity: { ...ticket },
      binding,
      lastTouched: now,
      terminal: false,
    });
    this.activeBindings.set(activeKey, key);
    this.replayTombstones.set(ticket.serverOperationId, {
      serverOperationId: ticket.serverOperationId,
      identityKey: key,
      expiresAt: now + this.ttlMs,
      bytes: replayBytes,
    });
    this.replayBytes += replayBytes;
    return { ok: true, reused: false, binding };
  }

  finishTerminal(ticket: ServerOwnedTicketIdentity): void {
    this.prune();
    const entry = this.entries.get(serverOwnedTicketKey(ticket));
    if (entry === undefined) return;
    entry.terminal = true;
    entry.lastTouched = this.now();
    entry.binding.state.finishTerminal();
  }

  /**
   * R4g-1: ends a delivered server inspection. The inspect entry becomes
   * terminal; the shared state stays `inspected` (for the one execute op with
   * the same turnKey) only if it is exactly `inspected`, otherwise it is
   * terminated exactly as by finishTerminal.
   */
  finishInspection(ticket: ServerOwnedTicketIdentity): void {
    this.prune();
    const entry = this.entries.get(serverOwnedTicketKey(ticket));
    if (entry === undefined) return;
    entry.terminal = true;
    entry.lastTouched = this.now();
    if (entry.identity.phase !== "inspect" || entry.binding.state.phase !== "inspected") entry.binding.state.finishTerminal();
  }

  isServerTicketLive(ticket: ServerOwnedTicketIdentity, binding: ServerOwnedTurnBinding): boolean {
    this.prune();
    const key = serverOwnedTicketKey(ticket);
    const entry = this.entries.get(key);
    const active = this.activeTurns.get(ticket.paneId);
    return entry !== undefined &&
      entry.binding === binding &&
      !entry.terminal &&
      this.activeBindings.get(entry.activeKey) === key &&
      samePaneTurn(active, ticket);
  }

  clearTurn(turn: ServerOwnedPaneTurn): void {
    const activeKey = serverOwnedPaneTurnKey(turn);
    for (const entry of [...this.entries.values()]) {
      if (entry.activeKey === activeKey) this.removeEntry(entry);
    }
    if (samePaneTurn(this.activeTurns.get(turn.paneId), turn)) this.activeTurns.delete(turn.paneId);
  }

  invalidateSession(sessionId: string): void {
    this.invalidate((entry) => entry.identity.sessionId === sessionId, (turn) => turn.sessionId === sessionId);
  }

  invalidateDocument(documentId: string): void {
    this.invalidate((entry) => entry.identity.documentId === documentId, (turn) => turn.documentId === documentId);
  }

  stop(): void {
    this.entries.clear();
    this.activeTurns.clear();
    this.activeBindings.clear();
  }

  private invalidate(
    entryMatches: (entry: { readonly identity: ServerOwnedTicketIdentity }) => boolean,
    turnMatches: (turn: ServerOwnedPaneTurn) => boolean,
  ): void {
    for (const entry of [...this.entries.values()]) {
      if (entryMatches(entry)) this.removeEntry(entry);
    }
    for (const [paneId, turn] of this.activeTurns) {
      if (turnMatches(turn)) this.activeTurns.delete(paneId);
    }
  }

  private prune(): void {
    const now = this.now();
    for (const entry of [...this.entries.values()]) {
      const active = this.activeTurns.get(entry.identity.paneId);
      const isActive = active !== undefined && serverOwnedPaneTurnKey(active) === entry.activeKey;
      if (entry.terminal && now - entry.lastTouched >= this.ttlMs) {
        this.removeEntry(entry);
        if (isActive) {
          this.activeTurns.delete(entry.identity.paneId);
        }
      }
    }
    for (const tombstone of [...this.replayTombstones.values()]) {
      if (now >= tombstone.expiresAt && this.entriesByReplayTombstone(tombstone) === undefined) {
        this.removeReplayTombstone(tombstone);
      }
    }
  }

  /**
   * R4g-1 phase fence for one pane turn: at most one inspect op, then at most
   * one execute op; no op after an execute. Returns the state the new binding
   * uses, or null when the phase is refused.
   */
  private phaseBindingState(ticket: ServerOwnedTicketIdentity, activeKey: string): WorkbookPlanTurnState | null {
    const bound = [...this.entries.values()].filter((entry) => entry.activeKey === activeKey);
    if (bound.length === 0) return this.activeBindings.has(activeKey) ? null : new WorkbookPlanTurnState();
    if (ticket.phase !== "execute" || bound.length !== 1) return null;
    const inspection = bound[0]!;
    if (
      inspection.identity.phase !== "inspect" ||
      inspection.identity.turnKey !== ticket.turnKey ||
      !inspection.terminal ||
      inspection.binding.state.phase !== "inspected"
    ) {
      return null;
    }
    return inspection.binding.state;
  }

  private removeEntry(entry: { readonly key: string; readonly activeKey: string; readonly serverOperationId: string }): void {
    this.entries.delete(entry.key);
    if (this.activeBindings.get(entry.activeKey) === entry.key) this.activeBindings.delete(entry.activeKey);
  }

  private entriesByReplayTombstone(tombstone: ServerOwnedReplayTombstone): {
    readonly terminal: boolean;
  } | undefined {
    const entry = this.entries.get(tombstone.identityKey);
    if (entry === undefined || entry.serverOperationId !== tombstone.serverOperationId) return undefined;
    return entry;
  }

  private removeReplayTombstone(tombstone: ServerOwnedReplayTombstone): void {
    if (this.replayTombstones.delete(tombstone.serverOperationId)) {
      this.replayBytes -= tombstone.bytes;
    }
  }
}

export interface WorkbookPlanTurnDependencies {
  readonly identity: WorkbookPlanTurnIdentity;
  readonly trustedOperationBoundary: TrustedOperationBoundary;
  readonly compileContext?: VerifiedWorkbookCompileContext;
  readonly getCompileContext?: (
    signal?: AbortSignal,
  ) => Promise<VerifiedWorkbookCompileContext>;
  /** Plan-aware host capture required for the complete B1 TOCTOU manifest. */
  readonly getCompileContextForPlan?: (
    plan: Extract<WorkbookPlan, { readonly outcome: "execute" }>,
    signal?: AbortSignal,
  ) => Promise<VerifiedWorkbookCompileContext>;
  readonly policyInput:
    | WorkbookPlanPolicyInput
    | ((
      plan: Extract<WorkbookPlan, { readonly outcome: "execute" }>,
      confirmation?: TrustedConfirmationEvent,
    ) => WorkbookPlanPolicyInput);
  readonly store: BrowserIdempotencyStore;
  readonly model: (
    continuation?: string,
    signal?: AbortSignal,
  ) => Promise<WorkbookPlanModelResponse>;
  readonly inspect: (
    reads: readonly InspectionOperation[],
    signal?: AbortSignal,
  ) => Promise<string>;
  readonly adapter: WorkbookPlanExecutionAdapter & WorkbookPlanVerificationAdapter;
  readonly signal?: AbortSignal;
  readonly turnStart?: number;
  readonly rolloutEnabled?: boolean;
  readonly route?: WorkbookPlanTurnRoute;
  /** Required for server-owned tickets; never inferred from model content. */
  readonly trustedTurnOperationId?: string;
  /** Route-owned preflight runs before inspection or host context reads. */
  readonly resourcePreflight?: (
    plan: WorkbookPlan,
    signal?: AbortSignal,
  ) => { readonly ok: true } | { readonly ok: false; readonly reason: string } | Promise<{ readonly ok: true } | { readonly ok: false; readonly reason: string }>;
  /** Exact route-owned resource authority after plan-aware runtime capture. */
  readonly resourcePreflightAfterRuntime?: (
    plan: Extract<WorkbookPlan, { readonly outcome: "execute" }>,
    signal?: AbortSignal,
  ) => { readonly ok: true } | { readonly ok: false; readonly reason: string } | Promise<{ readonly ok: true } | { readonly ok: false; readonly reason: string }>;
  /** Server tickets provide the already-normalized model response. */
  readonly initialResponse?: WorkbookPlanModelResponse;
  /** Server inspection tickets terminate after their one read. */
  readonly continueAfterInspection?: boolean;
  /**
   * R4g-1: trusted server ticket phase. Only `route === "server-owned"` with
   * `serverPhase === "inspect"` runs the one bounded inspection; the direct
   * route ignores it.
   */
  readonly serverPhase?: ServerOwnedPlanPhase;
  readonly turnState?: WorkbookPlanTurnState;
  /** Production route liveness fence checked before and after async boundaries. */
  readonly assertLive?: () => void;
  readonly now?: () => number;
  /** Generation-owned pane authority for a pending exact plan. */
  readonly turnGeneration?: number;
  readonly awaitPaneConfirmation?: (
    binding: WorkbookPlanConfirmationBinding,
    signal?: AbortSignal,
  ) => Promise<unknown>;
  readonly consumePaneConfirmation?: (
    candidate: unknown,
    binding: WorkbookPlanConfirmationBinding,
  ) => TrustedConfirmationEvent | undefined;
  readonly cancelPaneConfirmation?: () => void;
  /**
   * Trusted server-route acknowledgement, awaited by the executor gate after
   * the fresh observation and immediately before the first Office mutation.
   * Absent on the direct route, which then passes exactly the bare signal.
   */
  readonly beforeFirstMutation?: () => Promise<void>;
}

export interface WorkbookPlanVerificationAdapter {
  readonly verifyPlan?: (
    rules: readonly VerificationRule[],
    signal?: AbortSignal,
  ) => Promise<WorkbookPlanVerificationReport | boolean>;
  /** Compatibility seam for deterministic unit/native readback adapters. */
  readonly readVerification?: () => string;
}

export type WorkbookPlanTurnOutcome =
  | "committed"
  | "ask"
  | "refuse"
  | "inspected"
  | "rejected"
  | "verification_failed"
  | "cancelled"
  | "partial";

export interface WorkbookPlanTurnResult {
  readonly outcome: WorkbookPlanTurnOutcome;
  readonly finalSummary: string;
  readonly inspectionCount: number;
  readonly correctionCount: number;
  readonly execution?: WorkbookPlanExecutionReport;
  readonly verification: "verified" | "not_attempted" | "failed";
  /** R4g-1: the raw `inspect` result; set only on a server-owned inspect-phase `inspected` result. */
  readonly inspectionReads?: string;
}

const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

function serverOwnedPaneTurnKey(turn: ServerOwnedPaneTurn): string {
  return [turn.sessionId, turn.documentId, turn.paneId, turn.turnId].join("\u0000");
}

function serverOwnedTicketKey(ticket: ServerOwnedTicketIdentity): string {
  return `${serverOwnedPaneTurnKey(ticket)}\u0000${ticket.serverOperationId}`;
}

function serverOwnedReplayBytes(ticket: ServerOwnedTicketIdentity): number {
  return 64 + ticket.sessionId.length + ticket.documentId.length + ticket.paneId.length + ticket.turnId.length + ticket.serverOperationId.length;
}

function samePaneTurn(left: ServerOwnedPaneTurn | undefined, right: ServerOwnedPaneTurn): boolean {
  return left !== undefined &&
    left.sessionId === right.sessionId &&
    left.documentId === right.documentId &&
    left.paneId === right.paneId &&
    left.turnId === right.turnId;
}

function isServerOwnedPaneTurn(value: unknown): value is ServerOwnedPaneTurn {
  if (value === null || typeof value !== "object") return false;
  const candidate = value as Partial<ServerOwnedPaneTurn>;
  return isUuid(candidate.sessionId) && isUuid(candidate.documentId) && isUuid(candidate.paneId) && isUuid(candidate.turnId);
}

function isServerOwnedTicketIdentity(value: unknown): value is ServerOwnedTicketIdentity {
  if (!isServerOwnedPaneTurn(value)) return false;
  const candidate = value as Partial<ServerOwnedTicketIdentity>;
  if (!isUuid(candidate.serverOperationId)) return false;
  if (candidate.phase === undefined && candidate.turnKey === undefined) return true;
  return (candidate.phase === "inspect" || candidate.phase === "execute") && isUuid(candidate.turnKey);
}

/**
 * Runs the bounded WorkbookPlan state machine. This module deliberately has
 * no provider, task-pane, or Office dependency: callers supply the model,
 * inspection, and native execution seams.
 */
export async function createWorkbookPlanTurn(
  dependencies: WorkbookPlanTurnDependencies,
): Promise<WorkbookPlanTurnResult> {
  const now = dependencies.now ?? (() => Date.now());
  const turnStart = dependencies.turnStart ?? now();
  const deadline = turnStart + WORKBOOK_PLAN_TURN_TIMEOUT_MS;
  const turnState = dependencies.turnState ?? new WorkbookPlanTurnState();
  const identity = dependencies.identity;
  if (!isUuid(identity.sessionId) || !isUuid(identity.documentId) || !isUuid(identity.turnId)) {
    return rejected("Invalid trusted workbook turn identity.", 0, 0);
  }
  if (dependencies.route === "server-owned" && !isUuid(dependencies.trustedTurnOperationId)) {
    return rejected("Trusted server turn identity is unavailable.", 0, 0);
  }
  if (dependencies.rolloutEnabled !== true) {
    return rejected("Workbook plan execution is disabled.", 0, 0);
  }

  dependencies.assertLive?.();

  if (isSignalAborted(dependencies.signal)) {
    return cancelledOrRejected(new DOMException("The running Hermes task was stopped.", "AbortError"), 0, 0);
  }
  if (now() >= deadline) {
    return cancelledOrRejected(new DOMException("Workbook plan turn exceeded 90 seconds.", "TimeoutError"), 0, 0);
  }

  let inspections = 0;
  let corrections = 0;
  let response: WorkbookPlanModelResponse;
  try {
    response = dependencies.initialResponse ?? await callWithTurnOwnership(
      () => dependencies.model(undefined, dependencies.signal),
      dependencies.signal,
      deadline,
      now,
      dependencies.assertLive,
    );
  } catch (error) {
    return cancelledOrRejected(error, inspections, corrections);
  }

  while (true) {
    let normalized: WorkbookPlanProtocolResult;
    try {
      dependencies.assertLive?.();
      normalized = await normalizeResponse(response);
      dependencies.assertLive?.();
    } catch (error) {
      return cancelledOrRejected(error, inspections, corrections);
    }
    if (normalized.ok === false) {
      if (corrections >= WORKBOOK_PLAN_MAX_CORRECTIONS) {
        return rejected("Invalid workbook plan protocol envelope.", inspections, corrections);
      }
      corrections += 1;
      try {
        response = await callWithTurnOwnership(
          () => dependencies.model(
            "The previous response was rejected. Return exactly one valid workbook plan tool call; do not mix primitive calls or prose.",
            dependencies.signal,
          ),
          dependencies.signal,
          deadline,
          now,
          dependencies.assertLive,
        );
        continue;
      } catch (error) {
        return cancelledOrRejected(error, inspections, corrections);
      }
    }

    if (normalized.kind === "none") {
      return rejected("No workbook plan was provided.", inspections, corrections);
    }

    const plan = normalized.plan;
    const serverInspection = dependencies.route === "server-owned" && dependencies.serverPhase === "inspect";
    if (serverInspection && plan.outcome !== "inspect") {
      turnState.finishTerminal();
      return rejected("A server inspection ticket requires an inspect plan.", inspections, corrections);
    }
    if (plan.outcome === "ask" || plan.outcome === "refuse") {
      if (turnState.phase === "executing" || turnState.phase === "terminal") {
        return rejected("A terminal workbook turn cannot accept another model outcome.", inspections, corrections);
      }
      turnState.finishTerminal();
      return {
        outcome: plan.outcome,
        finalSummary: plan.message,
        inspectionCount: inspections,
        correctionCount: corrections,
        verification: "not_attempted",
      };
    }

    if (plan.outcome === "inspect") {
      if (dependencies.route === "server-owned" && !serverInspection) {
        turnState.finishTerminal();
        return rejected("Server-owned workbook inspection is not permitted.", inspections, corrections);
      }
      if (inspections >= 1 || !turnState.beginInspection()) {
        return rejected("A workbook turn may perform at most one inspection cycle.", inspections, corrections);
      }
      if (dependencies.resourcePreflight !== undefined) {
        const resourceResult = await callWithTurnOwnership(
          () => Promise.resolve(dependencies.resourcePreflight!(plan, dependencies.signal)),
          dependencies.signal,
          deadline,
          now,
          dependencies.assertLive,
        );
        if (resourceResult.ok !== true) {
          turnState.finishTerminal();
          return rejected("Workbook inspection resource preflight failed.", inspections, corrections);
        }
      }
      inspections += 1;
      if (serverInspection) {
        // R4g-1: one bounded read; its result goes back to the route, never to a model.
        let inspectionReads: string;
        try {
          inspectionReads = await callWithTurnOwnership(
            () => dependencies.inspect(plan.reads, dependencies.signal),
            dependencies.signal,
            deadline,
            now,
            dependencies.assertLive,
          );
        } catch (error) {
          return cancelledOrRejected(error, inspections, corrections);
        }
        return {
          outcome: "inspected",
          finalSummary: "Workbook inspection completed; awaiting the terminal plan outcome.",
          inspectionCount: inspections,
          correctionCount: corrections,
          verification: "not_attempted",
          inspectionReads,
        };
      }
      if (dependencies.continueAfterInspection === false) {
        try {
          await callWithTurnOwnership(
            () => dependencies.inspect(plan.reads, dependencies.signal),
            dependencies.signal,
            deadline,
            now,
          );
        } catch (error) {
          return cancelledOrRejected(error, inspections, corrections);
        }
        return {
          outcome: "inspected",
          finalSummary: "Workbook inspection completed; awaiting the terminal plan outcome.",
          inspectionCount: inspections,
          correctionCount: corrections,
          verification: "not_attempted",
        };
      }
      try {
        const inspectionResult = await callWithTurnOwnership(
          () => dependencies.inspect(plan.reads, dependencies.signal),
          dependencies.signal,
          deadline,
          now,
          dependencies.assertLive,
        );
        response = await callWithTurnOwnership(
          () => dependencies.model(inspectionResult, dependencies.signal),
          dependencies.signal,
          deadline,
          now,
          dependencies.assertLive,
        );
        continue;
      } catch (error) {
        return cancelledOrRejected(error, inspections, corrections);
      }
    }

    if (plan.outcome !== "execute") {
      return rejected("Unsupported workbook plan outcome.", inspections, corrections);
    }
    try {
      return await executePlan(plan, dependencies, inspections, corrections, deadline, now, turnState);
    } catch (error) {
      return cancelledOrRejected(error, inspections, corrections);
    }
  }
}

async function executePlan(
  plan: Extract<WorkbookPlan, { readonly outcome: "execute" }>,
  dependencies: WorkbookPlanTurnDependencies,
  inspections: number,
  corrections: number,
  deadline: number,
  now: () => number,
  turnState: WorkbookPlanTurnState,
): Promise<WorkbookPlanTurnResult> {
  if (isSignalAborted(dependencies.signal)) {
    return { outcome: "cancelled", finalSummary: "Turn stopped before workbook mutation.", inspectionCount: inspections, correctionCount: corrections, verification: "not_attempted" };
  }

  const canonical = await callWithTurnOwnership(
    () => canonicalizeWorkbookPlan(plan),
    dependencies.signal,
    deadline,
    now,
    dependencies.assertLive,
  );
  if (!canonical.ok) {
    turnState.finishTerminal();
    return rejected("Workbook plan canonicalization failed before mutation.", inspections, corrections);
  }
  const immutablePlan = deepFreezeWorkbookPlan(canonical.value) as Extract<WorkbookPlan, { readonly outcome: "execute" }>;
  const route = dependencies.route ?? "direct";
  const confirmationBinding: WorkbookPlanConfirmationBinding = {
    planDigest: canonical.semanticDigest,
    sessionId: dependencies.identity.sessionId,
    documentId: dependencies.identity.documentId,
    turnId: dependencies.identity.turnId,
    generation: dependencies.turnGeneration ?? 0,
    route,
    deadline,
    ...(route === "server-owned" && dependencies.trustedTurnOperationId !== undefined
      ? { serverOperationId: dependencies.trustedTurnOperationId }
      : {}),
  };

  if (dependencies.resourcePreflight !== undefined) {
    const resourceResult = await callWithTurnOwnership(
      () => Promise.resolve(dependencies.resourcePreflight!(immutablePlan, dependencies.signal)),
      dependencies.signal,
      deadline,
      now,
      dependencies.assertLive,
    );
    if (resourceResult.ok !== true) {
      turnState.finishTerminal();
      return rejected("Workbook execution resource preflight failed.", inspections, corrections);
    }
  }

  const compileContext = dependencies.getCompileContextForPlan !== undefined
    ? await callWithTurnOwnership(() => dependencies.getCompileContextForPlan!(immutablePlan, dependencies.signal), dependencies.signal, deadline, now, dependencies.assertLive)
    : dependencies.compileContext ?? (
      dependencies.getCompileContext !== undefined
        ? await callWithTurnOwnership(() => dependencies.getCompileContext!(dependencies.signal), dependencies.signal, deadline, now, dependencies.assertLive)
        : undefined
    );
  if (compileContext === undefined) {
    turnState.finishTerminal();
    return rejected("Workbook runtime context is unavailable.", inspections, corrections);
  }

  if (
    (dependencies.route === "direct" || dependencies.route === "server-owned") &&
    dependencies.resourcePreflightAfterRuntime === undefined
  ) {
    turnState.finishTerminal();
    return rejected("Workbook execution live resource authority is unavailable.", inspections, corrections);
  }

  if (dependencies.resourcePreflightAfterRuntime !== undefined) {
    const resourceResult = await callWithTurnOwnership(
      () => Promise.resolve(dependencies.resourcePreflightAfterRuntime!(immutablePlan, dependencies.signal)),
      dependencies.signal,
      deadline,
      now,
      dependencies.assertLive,
    );
    if (resourceResult.ok !== true) {
      turnState.finishTerminal();
      return rejected("Workbook execution live resource preflight failed.", inspections, corrections);
    }
  }

  // Preparation canonicalizes, applies policy, compiles every operation, and
  // freezes all arguments before the adapter is given any mutation capability.
  const prepare = (confirmation?: TrustedConfirmationEvent) => callWithTurnOwnership(
    () => prepareWorkbookPlanExecution({
      plan: immutablePlan,
      trustedOperationBoundary: dependencies.trustedOperationBoundary,
      compileContext,
      policyInput: typeof dependencies.policyInput === "function"
        ? dependencies.policyInput(immutablePlan, confirmation)
        : dependencies.policyInput,
    }),
    dependencies.signal,
    deadline,
    now,
    dependencies.assertLive,
  );
  let prepared = await prepare();
  if (requiresPaneConfirmation(prepared)) {
    if (
      dependencies.awaitPaneConfirmation === undefined ||
      dependencies.consumePaneConfirmation === undefined ||
      !turnState.beginPendingConfirmation()
    ) {
      turnState.finishTerminal();
      return rejected("Workbook confirmation is unavailable for this turn.", inspections, corrections);
    }
    let candidate: unknown;
    try {
      candidate = await dependencies.awaitPaneConfirmation(confirmationBinding, dependencies.signal);
      dependencies.assertLive?.();
    } catch (error) {
      dependencies.cancelPaneConfirmation?.();
      turnState.finishTerminal();
      throw error;
    }
    const confirmation = dependencies.consumePaneConfirmation(candidate, confirmationBinding);
    if (confirmation === undefined) {
      dependencies.cancelPaneConfirmation?.();
      turnState.finishTerminal();
      return rejected("Workbook confirmation was not accepted.", inspections, corrections);
    }
    if (!turnState.resumePendingConfirmation()) {
      turnState.finishTerminal();
      return rejected("Workbook confirmation is no longer active.", inspections, corrections);
    }
    prepared = await prepare(confirmation);
  }
  if (!prepared.accepted) {
    turnState.finishTerminal();
    return rejected("Workbook plan validation failed before mutation.", inspections, corrections);
  }
  if (!turnState.beginExecution()) {
    return rejected("Workbook execution is not permitted after the turn terminal boundary.", inspections, corrections);
  }
  if (isSignalAborted(dependencies.signal) || now() >= deadline) {
    return { outcome: "cancelled", finalSummary: "Turn stopped before workbook mutation.", inspectionCount: inspections, correctionCount: corrections, verification: "not_attempted" };
  }
  dependencies.assertLive?.();
  const execution = await executeAcceptedWorkbookPlan(
    prepared.transaction,
    dependencies.trustedOperationBoundary,
    dependencies.adapter,
    dependencies.store,
    dependencies.beforeFirstMutation === undefined
      ? dependencies.signal
      : { signal: dependencies.signal, beforeFirstMutation: dependencies.beforeFirstMutation, deadlineAt: deadline },
  );
  turnState.finishTerminal();
  if (isSignalAborted(dependencies.signal) || now() >= deadline) {
    const mutationWasStarted = execution.operations.some((operation) => operation.state !== "not_committed");
    return {
      outcome: mutationWasStarted ? "partial" : "cancelled",
      finalSummary: "Turn stopped; any in-flight workbook mutation was not claimed to be undone.",
      inspectionCount: inspections,
      correctionCount: corrections,
      execution,
      verification: "not_attempted",
    };
  }
  if (execution.outcome !== "committed") {
    return {
      outcome: execution.outcome === "cancelled" ? "cancelled" : "partial",
      finalSummary: `Workbook execution ended with reported state: ${execution.outcome}.`,
      inspectionCount: inspections,
      correctionCount: corrections,
      execution,
      verification: "not_attempted",
    };
  }

  const verification = await callWithTurnOwnership(
    () => verifyPlan(
      dependencies.adapter.verifyPlan !== undefined
        ? (getAcceptedWorkbookPlanPostconditionsForVerification(prepared.transaction) ?? plan.verification)
        : immutablePlan.verification,
      dependencies.adapter,
      dependencies.signal,
    ),
    dependencies.signal,
    deadline,
    now,
    dependencies.assertLive,
  );
  if (!verification.ok) {
    return {
      outcome: "verification_failed",
      finalSummary: "Workbook execution completed but postcondition verification failed; the turn terminated.",
      inspectionCount: inspections,
      correctionCount: corrections,
      execution,
      verification: "failed",
    };
  }
  return {
    outcome: "committed",
    finalSummary: `Verified ${execution.operationCount} workbook operation${execution.operationCount === 1 ? "" : "s"} completed successfully.`,
    inspectionCount: inspections,
    correctionCount: corrections,
    execution,
    verification: "verified",
  };
}

async function normalizeResponse(response: WorkbookPlanModelResponse): Promise<WorkbookPlanProtocolResult> {
  return normalizeWorkbookPlanResponse(response.content, JSON.stringify(response.toolCalls));
}

async function verifyPlan(
  rules: readonly VerificationRule[],
  adapter: WorkbookPlanTurnDependencies["adapter"],
  signal?: AbortSignal,
): Promise<{ readonly ok: boolean }> {
  if (adapter.verifyPlan !== undefined) {
    const result = await adapter.verifyPlan(rules, signal);
    return typeof result === "boolean" ? { ok: result } : { ok: result.ok === true };
  }
  if (adapter.readVerification === undefined) return { ok: false };
  const prepared = prepareWorkbookPlanVerificationRules(JSON.stringify(rules));
  if (prepared === null) return { ok: false };
  const report = verifyWorkbookPlan(prepared, { read: () => adapter.readVerification!() });
  return { ok: report.ok };
}

function requiresPaneConfirmation(
  result: { readonly accepted: boolean; readonly reason?: string; readonly policyReason?: string },
): boolean {
  if (result.accepted !== false || result.reason !== "policy_rejected") return false;
  switch (result.policyReason) {
    case "ambiguous_mutation_intent":
    case "structural_intent_required":
    case "active_sheet_mismatch":
    case "selection_scope_mismatch":
    case "confirmation_required":
      return true;
    default:
      return false;
  }
}

function deepFreezeWorkbookPlan<T>(value: T, seen = new WeakSet<object>()): T {
  if (value === null || (typeof value !== "object" && typeof value !== "function")) return value;
  const object = value as object;
  if (seen.has(object)) return value;
  seen.add(object);
  for (const key of Reflect.ownKeys(object)) {
    const descriptor = Object.getOwnPropertyDescriptor(object, key);
    if (descriptor !== undefined && "value" in descriptor) {
      deepFreezeWorkbookPlan(descriptor.value, seen);
    }
  }
  return Object.freeze(value);
}

function rejected(message: string, inspections: number, corrections: number): WorkbookPlanTurnResult {
  return { outcome: "rejected", finalSummary: message, inspectionCount: inspections, correctionCount: corrections, verification: "not_attempted" };
}

function cancelledOrRejected(error: unknown, inspections: number, corrections: number): WorkbookPlanTurnResult {
  if (error instanceof DOMException && (error.name === "AbortError" || error.name === "TimeoutError")) {
    return { outcome: "cancelled", finalSummary: error.name === "TimeoutError"
      ? "The 90-second workbook turn deadline expired; any in-flight mutation was not claimed to be undone."
      : "Turn stopped before claiming workbook success.", inspectionCount: inspections, correctionCount: corrections, verification: "not_attempted" };
  }
  return rejected(error instanceof Error ? error.message : "Workbook plan turn failed.", inspections, corrections);
}

async function callWithTurnOwnership<T>(
  operation: () => Promise<T>,
  signal: AbortSignal | undefined,
  deadline: number,
  now: () => number,
  assertLive?: () => void,
): Promise<T> {
  const cancellationError = (): DOMException => now() >= deadline
    ? new DOMException("Workbook plan turn exceeded 90 seconds.", "TimeoutError")
    : new DOMException("The running Hermes task was stopped.", "AbortError");

  assertLive?.();
  if (signal !== undefined && signal.aborted) throw cancellationError();
  if (now() >= deadline) throw new DOMException("Workbook plan turn exceeded 90 seconds.", "TimeoutError");

  const operationPromise = operation();
  if (signal === undefined) {
    const result = await operationPromise;
    assertLive?.();
    if (now() >= deadline) throw new DOMException("Workbook plan turn exceeded 90 seconds.", "TimeoutError");
    return result;
  }

  return new Promise<T>((resolve, reject) => {
    let settled = false;
    const cleanup = () => {
      signal.removeEventListener("abort", onAbort);
    };
    const finish = (callback: () => void): void => {
      if (settled) return;
      settled = true;
      cleanup();
      callback();
    };
    const onAbort = () => {
      finish(() => reject(cancellationError()));
    };

    signal.addEventListener("abort", onAbort, { once: true });
    if (signal.aborted) {
      onAbort();
      return;
    }
    operationPromise.then(
      (result) => {
        finish(() => {
          try {
            assertLive?.();
            if (signal.aborted) {
              reject(cancellationError());
            } else if (now() >= deadline) {
              reject(new DOMException("Workbook plan turn exceeded 90 seconds.", "TimeoutError"));
            } else {
              resolve(result);
            }
          } catch (error) {
            reject(error);
          }
        });
      },
      (error) => {
        finish(() => reject(error));
      },
    );
  });
}

function isUuid(value: unknown): value is string {
  return typeof value === "string" && UUID.test(value);
}

function isSignalAborted(signal: AbortSignal | undefined): boolean {
  return signal !== undefined && signal.aborted === true;
}

export type { AcceptedWorkbookPlanTransaction };
export type { NormalizedWorkbookPlan };
