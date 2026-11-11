import { WORKBOOK_PLAN_FINGERPRINT, type WorkbookPlan } from "../generated/workbookPlanSchema";
import { EXECUTE_START_ACK_TIMEOUT_MS, type BridgePlanExecutionControl, type BridgeTrustedPlanEnvelope } from "../officeBridge";
import type { ToolCall } from "../types";
import {
  createWorkbookPlanTurn,
  type ServerOwnedPlanPhase,
  type ServerOwnedTicketIdentity,
  type WorkbookPlanConfirmationBinding,
  type WorkbookPlanModelResponse,
  type WorkbookPlanTurnDependencies,
  type WorkbookPlanTurnOutcome,
  type WorkbookPlanTurnResult,
} from "../workbookPlanOrchestrator";
import { createTrustedOperationBoundary, isReplayedWorkbookPlanReport, type BrowserIdempotencyStore, type WorkbookPlanExecutionReport } from "../workbookPlanExecutor";
import { normalizeWorkbookPlanResponse, WORKBOOK_PLAN_TOOL_NAME } from "../workbookPlanProtocol";
import type { TrustedConfirmationEvent, TrustedWorkbookRequest } from "../workbookPlanRequest";
import {
  buildWorkbookPlanPolicyInput,
  preflightWorkbookPlanResources,
  type WorkbookPlanRouteRuntimeContext,
} from "../workbookPlanRoute";
import type { WorkbookPlanResourceContext } from "../workbookPlanResources";
import type { TaskPaneServerAuthorizationResult, TaskPaneTurnCoordinator } from "./taskPaneTurnCoordinator";

/** Legacy private plugin-side name; the pane never routes or validates it. */
const LEGACY_PLAN_TOOL_NAME = "execute_plan";

const LEGACY_PLAN_TOOL_MESSAGE = "Workbook plan tool name is not accepted by the pane.";
const UNTRUSTED_PLAN_MESSAGE = "Server-owned workbook plan requires trusted metadata.";
const PRIMITIVE_TRUST_MESSAGE = "Primitive Excel tools never accept trusted plan metadata.";
const SERVER_INSPECTION_MESSAGE = "Server-owned workbook inspection is not permitted.";

/**
 * R4g-1 support gate for trusted server inspection phases. Stays false until
 * R4g-3; while false a phase-inspect ticket is answered `rejected` before any
 * binding, read or store access. Only tests override it (dependency below).
 */
export const SERVER_INSPECTION_SUPPORTED = true;

export type { ServerOwnedPlanPhase };

/**
 * R4g-1: a trusted plan envelope that may carry a server phase. The R4g-1
 * bridge parser still yields 2-key (phase-less) envelopes only; `phase` and
 * `turnKey` are both present or both absent.
 */
export type ServerOwnedPlanTicketEnvelope = BridgeTrustedPlanEnvelope & {
  readonly phase?: ServerOwnedPlanPhase;
  readonly turnKey?: string;
};

const PHASE_TURN_KEY = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;

type ExecutePlan = Extract<WorkbookPlan, { readonly outcome: "execute" }>;

/** Bounded, content-free reason for a server-owned plan that never acknowledged execute-start. */
export type ServerOwnedPlanNotExecutedReason =
  | "ask"
  | "refuse"
  | "rejected"
  | "cancelled"
  | "not_started"
  | "invalid"
  | "already_executed";

/**
 * Fixed, content-free summary per non-executing reason. Orchestrator
 * `finalSummary` text is never forwarded for these: for outcomes produced by
 * `cancelledOrRejected` it can carry raw lower-layer `error.message` text.
 */
const NOT_EXECUTED_SUMMARIES: Readonly<Record<ServerOwnedPlanNotExecutedReason, string>> = Object.freeze({
  ask: "The workbook plan asked for clarification; nothing was executed.",
  refuse: "The workbook plan was refused; nothing was executed.",
  rejected: "The workbook plan was rejected before execution; nothing was executed.",
  cancelled: "The workbook plan was cancelled before execution; nothing was executed.",
  not_started: "Workbook plan execution did not start; nothing was executed.",
  invalid: "The workbook plan ticket was invalid; nothing was executed.",
  already_executed: "This workbook operation may have been applied earlier and was not executed again.",
});

export const EMPTY_RESOURCE_CONTEXT: WorkbookPlanResourceContext = {
  compileContext: {
    sheetNames: new Set<string>(),
    tablesByName: new Map(),
    supportedPrimitiveNames: new Set<string>(),
    supportedCapabilities: new Set<string>(),
    chartSourceResolutionsBySheet: new Map(),
  },
};

export async function preflightInitialWorkbookPlanResponse(response: WorkbookPlanModelResponse): Promise<void> {
  const normalized = await normalizeWorkbookPlanResponse(response.content, JSON.stringify(response.toolCalls));
  if (normalized.ok !== true || normalized.kind === "none") return;
  if (normalized.plan.outcome !== "inspect" && normalized.plan.outcome !== "execute") return;
  const preflight = preflightWorkbookPlanResources(normalized.plan, EMPTY_RESOURCE_CONTEXT);
  if (preflight.ok !== true) throw new Error("Workbook plan resource preflight failed");
}

export function trustedRequestWithConfirmation(
  request: TrustedWorkbookRequest,
  confirmation?: TrustedConfirmationEvent,
): TrustedWorkbookRequest {
  if (confirmation === undefined) return request;
  return Object.freeze({
    ...request,
    confirmation: Object.freeze({ ...confirmation }),
  });
}

/** Snapshot of the pane's active trusted turn, read at ticket time. */
export interface ServerOwnedPlanActiveTurn {
  readonly turnOperationId: string | null;
  readonly trustedRequest: TrustedWorkbookRequest | null;
  readonly sessionId: string | null;
  readonly turnStart: number | null;
  readonly hasCancellation: boolean;
}

export interface ServerOwnedPlanRouteDependencies {
  /** Pane-owned rollout setting; checked before any session, document or runtime read. */
  readonly isRolloutEnabled: () => boolean;
  readonly getBridgeSessionId: () => string | null;
  readonly getActiveTurn: () => ServerOwnedPlanActiveTurn;
  readonly coordinator: Pick<
    TaskPaneTurnCoordinator,
    "authorizeAndRun" | "consumeWorkbookPlanConfirmation" | "cancelPendingConfirmation"
  >;
  readonly bindings: {
    readonly finishTerminal: (ticket: ServerOwnedTicketIdentity) => void;
    /** R4g-1: ends a delivered server inspection, keeping the shared state `inspected`. */
    readonly finishInspection: (ticket: ServerOwnedTicketIdentity) => void;
  };
  /** Plan-aware native runtime capture (Office reads). */
  readonly readRuntimeContext: (plan?: ExecutePlan) => Promise<WorkbookPlanRouteRuntimeContext>;
  readonly createAdapter: () => WorkbookPlanTurnDependencies["adapter"];
  readonly createStore: () => BrowserIdempotencyStore;
  readonly inspect: WorkbookPlanTurnDependencies["inspect"];
  readonly awaitPaneConfirmation: (
    binding: WorkbookPlanConfirmationBinding,
    signal?: AbortSignal,
  ) => Promise<unknown>;
  /** Existing primitive validator path; plan names never reach it. */
  readonly executePrimitive: (name: string, argumentsValue: Record<string, unknown>) => Promise<string>;
  /** Clock for the claim-lease confirmation cap; defaults to Date.now. */
  readonly now?: () => number;
  /** R4g-1 test-only override of SERVER_INSPECTION_SUPPORTED; taskpane.ts never passes it. Read only for phase-inspect tickets. */
  readonly serverInspectionSupported?: () => boolean;
}

/** Companion claim lease (protocol.py CLAIM_LEASE_SECONDS = 20; fixture claimLeaseSeconds). */
export const COMPANION_CLAIM_LEASE_MS = 20_000;
/** Margin kept between the confirmation cap and the latest execute-start acknowledgement. */
export const CONFIRMATION_CAP_MARGIN_MS = 2_000;

/**
 * Caps the server-route pane confirmation at ticket receipt + claim lease -
 * execute-start ack timeout - margin (13 s), so an unanswered confirmation is
 * answered executed:false before the companion's claim lease ends. Pure: no
 * session, document or runtime read. Await and consume see the same capped
 * binding values; an already-expired confirmation is rejected without showing
 * the pane bar. Coordinator methods are bound explicitly (it is a class).
 */
export function withClaimLeaseConfirmationCap(
  raw: ServerOwnedPlanRouteDependencies,
  claimObservedAt: number,
  now: () => number,
): ServerOwnedPlanRouteDependencies {
  const cap = claimObservedAt + COMPANION_CLAIM_LEASE_MS - EXECUTE_START_ACK_TIMEOUT_MS - CONFIRMATION_CAP_MARGIN_MS;
  const capBinding = (binding: WorkbookPlanConfirmationBinding): WorkbookPlanConfirmationBinding =>
    binding.deadline <= cap ? binding : Object.freeze({ ...binding, deadline: cap });
  return Object.freeze({
    ...raw,
    awaitPaneConfirmation: (binding: WorkbookPlanConfirmationBinding, signal?: AbortSignal): Promise<unknown> => {
      const capped = capBinding(binding);
      return capped.deadline <= now()
        ? Promise.reject(new DOMException("The pending workbook action expired.", "TimeoutError"))
        : raw.awaitPaneConfirmation(capped, signal);
    },
    coordinator: Object.freeze({
      authorizeAndRun: ((...args: Parameters<TaskPaneTurnCoordinator["authorizeAndRun"]>) =>
        raw.coordinator.authorizeAndRun(...args)) as TaskPaneTurnCoordinator["authorizeAndRun"],
      consumeWorkbookPlanConfirmation: (candidate: unknown, binding: WorkbookPlanConfirmationBinding) =>
        raw.coordinator.consumeWorkbookPlanConfirmation(candidate, capBinding(binding)),
      cancelPendingConfirmation: () => raw.coordinator.cancelPendingConfirmation(),
    }),
  });
}

export interface ServerOwnedPlanRoute {
  readonly executeCompanionTool: (
    name: string,
    argumentsValue: Record<string, unknown>,
    trusted?: BridgeTrustedPlanEnvelope,
    execution?: BridgePlanExecutionControl,
  ) => Promise<string>;
  readonly executeServerOwnedWorkbookPlan: (
    planArguments: Record<string, unknown>,
    trusted: ServerOwnedPlanTicketEnvelope,
    execution?: BridgePlanExecutionControl,
  ) => Promise<string>;
}

type AcknowledgementState = "unused" | "pending" | "resolved" | "rejected";

/**
 * Server-owned workbook plan route, extracted from taskpane.ts (R2d-1, M2) so
 * the production chain can be exercised with injected dependencies. The pane
 * keys on the canonical WORKBOOK_PLAN_TOOL_NAME and requires a trusted
 * envelope; the bridge's one-shot execute-start acknowledgement is handed to
 * the executor as `beforeFirstMutation`, so execution is acknowledged only at
 * the first-mutation gate.
 */
export function createServerOwnedPlanRoute(dependencies: ServerOwnedPlanRouteDependencies): ServerOwnedPlanRoute {
  async function executeCompanionTool(
    name: string,
    argumentsValue: Record<string, unknown>,
    trusted?: BridgeTrustedPlanEnvelope,
    execution?: BridgePlanExecutionControl,
  ): Promise<string> {
    // Name checks come first: no read, Office call or validator runs for a
    // legacy or untrusted plan name.
    if (name === LEGACY_PLAN_TOOL_NAME) throw new Error(LEGACY_PLAN_TOOL_MESSAGE);
    if (name !== WORKBOOK_PLAN_TOOL_NAME) {
      if (trusted !== undefined || execution !== undefined) throw new Error(PRIMITIVE_TRUST_MESSAGE);
      return dependencies.executePrimitive(name, argumentsValue);
    }
    if (trusted === undefined) throw new Error(UNTRUSTED_PLAN_MESSAGE);
    return executeServerOwnedWorkbookPlan(argumentsValue, trusted, execution);
  }

  async function executeServerOwnedWorkbookPlan(
    planArguments: Record<string, unknown>,
    trusted: ServerOwnedPlanTicketEnvelope,
    execution?: BridgePlanExecutionControl,
  ): Promise<string> {
    // This gate is intentionally before session, document, runtime, or plan
    // inspection reads. Task12 supplies the ticket; the pane owns the setting.
    if (dependencies.isRolloutEnabled() !== true) return notExecuted("not_started");
    const sessionId = dependencies.getBridgeSessionId();
    if (sessionId === null) return notExecuted("rejected");
    const active = dependencies.getActiveTurn();
    if (active.turnOperationId === null || active.trustedRequest === null || active.sessionId !== sessionId) {
      return notExecuted("rejected");
    }
    if (active.turnStart === null || active.hasCancellation !== true) return notExecuted("rejected");
    if (trusted.workbookPlanFingerprint !== WORKBOOK_PLAN_FINGERPRINT) return notExecuted("invalid");
    // R4g-1 phase fence, before any binding, read or store access. (a)/(b)
    // read only the envelope and the raw outcome string; (c) reads the
    // support gate for phase-inspect tickets only. Phase-less tickets skip it.
    const phase = trusted.phase;
    const turnKey = trusted.turnKey;
    if (phase !== undefined || turnKey !== undefined) {
      if ((phase !== "inspect" && phase !== "execute") || typeof turnKey !== "string" || !PHASE_TURN_KEY.test(turnKey) || turnKey === trusted.turnOperationId) {
        return notExecuted("invalid");
      }
      const outcome = planArguments.outcome;
      if (phase === "inspect" ? outcome !== "inspect" : outcome !== "execute" && outcome !== "ask" && outcome !== "refuse") {
        return notExecuted("invalid");
      }
      if (phase === "inspect") {
        const supported = dependencies.serverInspectionSupported;
        if ((supported === undefined ? SERVER_INSPECTION_SUPPORTED : supported() === true) !== true) return notExecuted("rejected");
      }
    }
    const phaseArguments: [] | [{ readonly phase: ServerOwnedPlanPhase; readonly turnKey: string }] =
      phase === undefined || turnKey === undefined ? [] : [{ phase, turnKey }];
    // Ticket receipt time: nothing above awaits, so this is the same tick as
    // the bridge's GET /pending body read (>= the companion's claimed_at).
    const claimObservedAt = (dependencies.now ?? Date.now)();
    const cappedDependencies = withClaimLeaseConfirmationCap(dependencies, claimObservedAt, dependencies.now ?? Date.now);

    // Mutated from inside the authorized continuation and the executor gate.
    const progress: {
      acknowledgement: AcknowledgementState;
      captured: WorkbookPlanTurnResult | undefined;
      inspection: string | undefined;
    } = {
      acknowledgement: "unused",
      captured: undefined,
      inspection: undefined,
    };
    let authorized: TaskPaneServerAuthorizationResult<WorkbookPlanTurnResult>;
    try {
      authorized = await dependencies.coordinator.authorizeAndRun(trusted.turnOperationId, async (authorizedOperation) => {
        const ticketIdentity = authorizedOperation.identity;
        const snapshot = authorizedOperation.snapshot;
        const trustedRequest = snapshot.trustedRequest;
        const turnStart = snapshot.turnStart;
        if (trustedRequest === null || turnStart === null) {
          throw new Error("Trusted server turn snapshot is unavailable");
        }
        authorizedOperation.assertLive();
        const planResponse: WorkbookPlanModelResponse = {
          content: "",
          toolCalls: [{
            id: ticketIdentity.serverOperationId,
            type: "function",
            function: { name: WORKBOOK_PLAN_TOOL_NAME, arguments: JSON.stringify(planArguments) },
          } as ToolCall],
        };
        try {
          authorizedOperation.assertLive();
          await preflightInitialWorkbookPlanResponse(planResponse);
          authorizedOperation.assertLive();
        } catch (error) {
          dependencies.bindings.finishTerminal(ticketIdentity);
          throw error;
        }
        const documentId = ticketIdentity.documentId;
        let runtime: WorkbookPlanRouteRuntimeContext | undefined;
        const loadRuntime = async (
          plan?: ExecutePlan,
          signal?: AbortSignal,
        ): Promise<WorkbookPlanRouteRuntimeContext> => {
          authorizedOperation.assertLive();
          if (signal !== undefined && signal.aborted) {
            throw new DOMException("The running Hermes task was stopped.", "AbortError");
          }
          runtime = plan === undefined
            ? (runtime ?? await dependencies.readRuntimeContext())
            : await dependencies.readRuntimeContext(plan);
          authorizedOperation.assertLive();
          if (signal !== undefined && signal.aborted) {
            throw new DOMException("The running Hermes task was stopped.", "AbortError");
          }
          return runtime;
        };
        const beforeFirstMutation = execution === undefined
          ? undefined
          : async (): Promise<void> => {
            progress.acknowledgement = "pending";
            try {
              await execution.acknowledgeExecuteStart(snapshot.cancellationSignal);
            } catch (error) {
              progress.acknowledgement = "rejected";
              throw error;
            }
            progress.acknowledgement = "resolved";
          };
        const turnState = authorizedOperation.binding.state;
        let result: WorkbookPlanTurnResult;
        try {
          const dependencies = cappedDependencies;
          result = await createWorkbookPlanTurn({
            identity: { sessionId: ticketIdentity.sessionId, documentId, turnId: ticketIdentity.turnId },
            trustedOperationBoundary: createTrustedOperationBoundary({ sessionId: ticketIdentity.sessionId, documentId, turnId: ticketIdentity.serverOperationId }),
            getCompileContextForPlan: async (plan, _signal) => (await loadRuntime(plan, _signal)).compileContext,
            policyInput: (plan, confirmation) => {
              authorizedOperation.assertLive();
              if (runtime === undefined) throw new Error("Trusted workbook runtime context is unavailable");
              return buildWorkbookPlanPolicyInput(plan, runtime, trustedRequestWithConfirmation(trustedRequest, confirmation));
            },
            resourcePreflight: (plan, _signal) => preflightWorkbookPlanResources(plan, EMPTY_RESOURCE_CONTEXT),
            resourcePreflightAfterRuntime: (plan) => runtime === undefined
              ? { ok: false, reason: "live_runtime_unavailable" }
              : preflightWorkbookPlanResources(plan, runtime),
            store: dependencies.createStore(),
            model: async () => { throw new Error("Server-owned ticket unexpectedly requested model continuation"); },
            initialResponse: planResponse,
            inspect: dependencies.inspect,
            adapter: dependencies.createAdapter(),
            signal: snapshot.cancellationSignal,
            turnStart,
            route: "server-owned",
            trustedTurnOperationId: ticketIdentity.serverOperationId,
            turnGeneration: snapshot.generation,
            awaitPaneConfirmation: dependencies.awaitPaneConfirmation,
            consumePaneConfirmation: (candidate, binding) => dependencies.coordinator.consumeWorkbookPlanConfirmation(candidate, binding),
            cancelPaneConfirmation: () => dependencies.coordinator.cancelPendingConfirmation(),
            turnState,
            continueAfterInspection: false,
            rolloutEnabled: dependencies.isRolloutEnabled() === true,
            assertLive: authorizedOperation.assertLive,
            ...(beforeFirstMutation === undefined ? {} : { beforeFirstMutation }),
            ...(phase === undefined ? {} : { serverPhase: phase }),
          });
        } catch (error) {
          dependencies.bindings.finishTerminal(ticketIdentity);
          throw error;
        }
        if (phase === "inspect") {
          // R4g-1: only a delivered, well-formed read keeps the shared state
          // `inspected`; every other end terminates it.
          const body = result.outcome === "inspected" ? inspectionBody(result) : undefined;
          if (body === undefined) {
            dependencies.bindings.finishTerminal(ticketIdentity);
          } else {
            dependencies.bindings.finishInspection(ticketIdentity);
            progress.inspection = body;
          }
          progress.captured = result;
          return result;
        }
        dependencies.bindings.finishTerminal(ticketIdentity);
        progress.captured = result;
        if (result.outcome === "inspected") {
          throw new Error(SERVER_INSPECTION_MESSAGE);
        }
        return result;
      }, ...phaseArguments);
    } catch (error) {
      if (progress.acknowledgement === "resolved") return executedWithoutConfirmedResult();
      if (progress.acknowledgement !== "unused") return notExecuted("not_started");
      return notExecuted(isCancellation(error) ? "cancelled" : "rejected");
    }

    // A Stop after the continuation finished makes the authorization stale,
    // but the captured report is still the truthful outcome (for example a
    // partial write that is never claimed as undone).
    const result = authorized.ok === true ? authorized.result : progress.captured;
    if (progress.acknowledgement === "resolved") {
      // Once execute-start was acknowledged the workbook may have changed, so
      // the model only ever sees an executed report or the uncertain shape:
      // never a rejected/cancelled outcome or a non-executing reason (which
      // could invite a retry that writes twice).
      return hasExecutedReport(result) ? executedResult(result) : executedWithoutConfirmedResult();
    }
    if (progress.acknowledgement !== "unused") return notExecuted("not_started", result);
    if (authorized.ok === true && progress.inspection !== undefined) return progress.inspection;
    if (result === undefined) return notExecuted("rejected");
    if (isReplayedWorkbookPlanReport(result.execution) && result.execution!.operations.some((operation) => operation.state !== "not_committed")) return alreadyExecuted(result.execution!);
    return notExecuted(reasonForOutcome(result.outcome), result);
  }

  return { executeCompanionTool, executeServerOwnedWorkbookPlan };
}

function executedResult(result: WorkbookPlanTurnResult): string {
  return JSON.stringify({ ok: result.outcome === "committed", action: "workbook_plan", result });
}

/**
 * R4g-1 inspection body (pane-internal until R4g-2 carries phases on the
 * wire): exactly ok, action, phase and the parsed `reads` of the one bounded
 * read. No reason, summary, state, error, execution or display text.
 * Undefined when the read result is not a JSON object with a `reads` array.
 * R4g-2b (R4g-2a lane-B A3): undefined (fail closed) unless every read is a
 * plain object with exactly the T8 read keys {ok, data} and a boolean `ok`,
 * each read is at most INSPECTION_READ_MAX_BYTES and the whole body at most
 * INSPECTION_BODY_MAX_BYTES (UTF-8 bytes of the compact JSON).
 * R4g-2c (R4g-2b lane-B A6, A4): also undefined for an empty `reads` list (an
 * inspection with no read never unlocks the execute phase) and for a read
 * nesting deeper than INSPECTION_READ_MAX_DEPTH levels.
 */
function inspectionBody(result: WorkbookPlanTurnResult): string | undefined {
  if (result.inspectionCount !== 1 || typeof result.inspectionReads !== "string") return undefined;
  let parsed: unknown;
  try {
    parsed = JSON.parse(result.inspectionReads);
  } catch {
    return undefined;
  }
  if (parsed === null || typeof parsed !== "object" || Array.isArray(parsed)) return undefined;
  const reads = (parsed as { readonly reads?: unknown }).reads;
  if (!Array.isArray(reads) || reads.length === 0) return undefined;
  if (!reads.every(isBoundedInspectionRead)) return undefined;
  const body = JSON.stringify({
    ok: true,
    action: "workbook_plan",
    phase: "inspect",
    result: { outcome: "inspected", inspectionCount: 1, reads },
  });
  return utf8Length(body) <= INSPECTION_BODY_MAX_BYTES ? body : undefined;
}

/**
 * R4g-2b provisional caps (r4g2b/case-state.md): one read may carry the whole
 * 5,000-cell inspection budget (values, formulas and number formats of every
 * cell) at about 200 bytes per cell; the body leaves room for that read plus
 * smaller ones, and stays far below the companion's 10 MiB result cap.
 */
const INSPECTION_READ_MAX_BYTES = 1_048_576;
const INSPECTION_BODY_MAX_BYTES = 1_572_864;
/** R4g-2c provisional cap: real read shapes nest at most 4 levels; the companion's JSON parser fails near 1,000. */
const INSPECTION_READ_MAX_DEPTH = 64;

function isBoundedInspectionRead(read: unknown): boolean {
  if (read === null || typeof read !== "object" || Array.isArray(read) || Object.getPrototypeOf(read) !== Object.prototype) return false;
  const keys = Object.keys(read);
  if (keys.length !== 2 || !keys.includes("ok") || !keys.includes("data")) return false;
  if (typeof (read as { readonly ok?: unknown }).ok !== "boolean") return false;
  if (!withinNestingDepth(read, INSPECTION_READ_MAX_DEPTH)) return false;
  return utf8Length(JSON.stringify(read)) <= INSPECTION_READ_MAX_BYTES;
}

function utf8Length(text: string): number {
  return new TextEncoder().encode(text).length;
}

/** True when `value` nests at most `limit` levels of arrays/objects (the read object itself is level 1). */
function withinNestingDepth(value: unknown, limit: number): boolean {
  if (value === null || typeof value !== "object") return true;
  if (limit <= 0) return false;
  const children: readonly unknown[] = Array.isArray(value) ? value : Object.values(value);
  return children.every((child) => withinNestingDepth(child, limit - 1));
}

function executedWithoutConfirmedResult(): string {
  return JSON.stringify({
    ok: false,
    action: "workbook_plan",
    state: "uncertain",
    error: "Workbook plan execution ended without a confirmed result.",
  });
}

/**
 * An executed report exists only for committed / verification_failed /
 * partial results that carry their execution report. After a write is
 * acknowledged, `rejected`, `cancelled`, `ask`, `refuse` or `inspected`, or a
 * missing execution report, means the outcome is uncertain.
 */
function hasExecutedReport(result: WorkbookPlanTurnResult | undefined): result is WorkbookPlanTurnResult {
  if (result === undefined || result.execution === undefined) return false;
  return result.outcome === "committed" || result.outcome === "verification_failed" || result.outcome === "partial";
}

/**
 * Non-executing payload: the bounded reason plus its fixed summary. The
 * orchestrator's `finalSummary` is forwarded only for ask/refuse, where
 * createWorkbookPlanTurn sets it to `plan.message` of the schema-validated
 * plan (bounded at 2048 characters) rather than to any lower-layer error.
 */
function notExecuted(reason: ServerOwnedPlanNotExecutedReason, result?: WorkbookPlanTurnResult): string {
  const planMessage = (reason === "ask" || reason === "refuse") && result !== undefined && result.outcome === reason
    ? { result: { outcome: result.outcome, finalSummary: result.finalSummary } }
    : {};
  return JSON.stringify({
    ok: false,
    action: "workbook_plan",
    reason,
    summary: NOT_EXECUTED_SUMMARIES[reason],
    ...planMessage,
  });
}

/**
 * R4e: this ticket never started a write, but its store claim returned an
 * earlier attempt's recorded report showing the workbook was (or may have
 * been) changed. The recorded, schema-validated report is forwarded; the
 * orchestrator's finalSummary is not.
 */
function alreadyExecuted(execution: WorkbookPlanExecutionReport): string {
  return JSON.stringify({
    ok: false,
    action: "workbook_plan",
    reason: "already_executed",
    summary: NOT_EXECUTED_SUMMARIES.already_executed,
    result: { execution },
  });
}

function reasonForOutcome(outcome: WorkbookPlanTurnOutcome): ServerOwnedPlanNotExecutedReason {
  switch (outcome) {
    case "ask":
    case "refuse":
    case "rejected":
    case "cancelled":
      return outcome;
    case "inspected":
      return "rejected";
    default:
      // committed / partial / verification_failed without an acknowledgement
      // cannot have mutated through this route.
      return "not_started";
  }
}

function isCancellation(error: unknown): boolean {
  return error instanceof DOMException && (error.name === "AbortError" || error.name === "TimeoutError");
}
