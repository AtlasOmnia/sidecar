import { afterEach, describe, expect, it, vi } from "vitest";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import type { BridgePublicState, OfficeCompanionBridge } from "../officeBridge";
import { createExcelCompanionBridge } from "../officeBridgeExcel";
import { OfficeBridgeLifecycleController } from "../officeBridgeLifecycle";
import {
  ServerOwnedTurnBindingStore,
  type WorkbookPlanConfirmationBinding,
  type WorkbookPlanVerificationAdapter,
} from "../workbookPlanOrchestrator";
import {
  BrowserIdempotencyStore,
  MemoryIdempotencyStorage,
  type WorkbookPlanExecutionAdapter,
} from "../workbookPlanExecutor";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import type { WorkbookPlanRouteRuntimeContext } from "../workbookPlanRoute";
import { classifyWorkbookRequest, type TrustedWorkbookRequest } from "../workbookPlanRequest";
import {
  TaskPaneTurnCoordinator,
  type TaskPaneTurnHandle,
  type TaskPaneTurnPreparation,
} from "../taskpane/taskPaneTurnCoordinator";
import * as routeModule from "../taskpane/serverOwnedPlanRoute";

/**
 * R4f real-chain rows (release dev3-r4f-release.md §5): A3 (marks/taints of
 * replaced sessions are pruned at prepare), A4 (a prepare waits out another
 * handle's rotation, raced with its own signal), A5 (a late background error
 * from a superseded session never tears down the new one), A6 (D-R4-8: one
 * bounded reconnect attempt per sessionless prepare after a failed rotation)
 * and LP11 (the force-stop policy equals the base decision when no rotation is
 * in flight). Real chain: createExcelCompanionBridge ->
 * OfficeBridgeLifecycleController -> TaskPaneTurnCoordinator ->
 * createServerOwnedPlanRoute. The SessionCompanion double and sessionHarness
 * are COPIED from task-pane-session-rotation-r3b.test.ts:105-525 (not
 * imported; that file stays byte-identical), with a hold-able POST /session and
 * GET /pending added.
 */

const DOC1 = "d0c00000-0000-4000-8000-000000000001";
const DOC2 = "d0c00000-0000-4000-8000-000000000002";
const PANE = "9a9e0000-0000-4000-8000-0000000000aa";
const T0 = "7a7a0000-0000-4000-8000-000000000000";
const T1 = "7a7a0000-0000-4000-8000-000000000001";
const T2 = "7a7a0000-0000-4000-8000-000000000002";
const T3 = "7a7a0000-0000-4000-8000-000000000003";
const T4 = "7a7a0000-0000-4000-8000-000000000004";
const T2_OP = "5e7e0000-0000-4000-8000-000000000002";
const EXPLICIT_PROMPT = "write ok into A1";
const NOT_FOUND_BODY = { ok: false, error: "Session not found", category: "NotFound" };
const COMPILE_CONTEXT: VerifiedWorkbookCompileContext = {
  sheetNames: new Set(["Sheet1"]),
  tablesByName: new Map(),
  supportedPrimitiveNames: new Set(["write_range"]),
  supportedCapabilities: new Set(["excel_workbook_access", "workbook_plan_v1_operation_audit"]),
  chartSourceResolutionsBySheet: new Map(),
};
const RUNTIME: WorkbookPlanRouteRuntimeContext = {
  activeSheetName: "Sheet1",
  selection: { sheetName: "Sheet1", address: "A1" },
  compileContext: COMPILE_CONTEXT,
};
const EXECUTE_PLAN = {
  version: 1,
  outcome: "execute",
  summary: "write one cell",
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }],
  verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["ok"]] }],
} as const;

interface ActiveTurnView {
  turnId: string;
  sessionId: string | null;
  documentId: string | null;
}

interface R3bCoordinatorApi {
  beginHermesRequest?: (sessionId: string | null) => { endClean(): void };
  bridgeStateInvalidatesActiveTurn?: (
    state: BridgePublicState,
    active: ActiveTurnView | null,
    previous: { readonly sessionId: string | null; readonly documentId: string | null },
  ) => boolean;
  readonly hermesRequestMarkCount?: number;
}

function r3b(coordinator: TaskPaneTurnCoordinator): R3bCoordinatorApi {
  return coordinator as unknown as R3bCoordinatorApi;
}

function json(status: number, body?: unknown): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
  });
}

function uuid(prefix: string, n: number): string {
  return `${prefix}-0000-4000-8000-${String(n).padStart(12, "0")}`;
}

type TicketState = "pending" | "claimed" | "executing" | "completed" | "not_executed" | "cancelled" | "uncertain";

interface SessionTicket {
  readonly ticketId: string;
  readonly requestId: string;
  readonly tool: string;
  readonly args: Record<string, unknown>;
  readonly trusted?: unknown;
  state: TicketState;
}

interface SessionRecord {
  readonly sessionId: string;
  readonly sessionKey: string;
  readonly documentId: string;
  readonly tickets: SessionTicket[];
}

type ResultScriptStep = number | "network";

/** Session-scoped companion double (tickets keyed by session). */
class SessionCompanion {
  readonly byKey = new Map<string, SessionRecord>();
  readonly byId = new Map<string, SessionRecord>();
  readonly deleted: string[] = [];
  readonly results: { sessionId: string; body: Record<string, unknown> }[] = [];
  readonly tickets = new Map<string, SessionTicket>();
  readonly resultScripts = new Map<string, ResultScriptStep[]>();
  sessionPosts = 0;
  executeStarts = 0;
  sessionFailure: number | null = null;
  /** R4f: holds for the next POST /session / GET /pending requests (FIFO); released by the test. */
  private readonly sessionHolds: Promise<number | null>[] = [];
  private readonly pendingHolds: Promise<"process" | "network">[] = [];
  heldSessionWaiting = 0;
  heldPendingWaiting = 0;
  private sessionCounter = 0;
  private ticketCounter = 0;

  constructor(private readonly events: string[]) {}

  /** Plugin-side invoke: 404 (fixture C14 invoke-unknown-session) for an unknown or deleted session. */
  invoke(
    sessionId: string,
    serverOperationId: string,
    plan: Record<string, unknown>,
  ): { status: number; body: Record<string, unknown>; ticketId?: string } {
    return this.enqueue(sessionId, "excel_execute_plan", { ...plan }, {
      turnOperationId: serverOperationId,
      workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
      phase: "execute",
      turnKey: "4b4b0000-0000-4000-8000-0000000000a3",
    });
  }

  invokePrimitive(sessionId: string, tool: string, args: Record<string, unknown>) {
    return this.enqueue(sessionId, tool, args, undefined);
  }

  private enqueue(sessionId: string, tool: string, args: Record<string, unknown>, trusted: unknown) {
    const session = this.byId.get(sessionId);
    if (session === undefined) return { status: 404, body: { ...NOT_FOUND_BODY } };
    this.ticketCounter += 1;
    const ticket: SessionTicket = {
      ticketId: uuid("71c70000", this.ticketCounter),
      requestId: uuid("4e4e0000", this.ticketCounter),
      tool,
      args,
      ...(trusted === undefined ? {} : { trusted }),
      state: "pending",
    };
    session.tickets.push(ticket);
    this.tickets.set(ticket.ticketId, ticket);
    return { status: 202, body: { ok: true }, ticketId: ticket.ticketId };
  }

  /** Holds the next POST /session; release(status) answers it with `status` (or normally when null). */
  holdNextSession(): { release: (status?: number | null) => void } {
    let release!: (status: number | null) => void;
    this.sessionHolds.push(new Promise<number | null>((done) => { release = done; }));
    return { release: (status = null) => release(status) };
  }

  /** Holds the next GET /pending before authentication; "process" answers it as the companion would at release time. */
  holdNextPending(): { release: (mode: "process" | "network") => void } {
    let release!: (mode: "process" | "network") => void;
    this.pendingHolds.push(new Promise<"process" | "network">((done) => { release = done; }));
    return { release: (mode) => release(mode) };
  }

  resultsFor(ticketId: string | undefined): Record<string, unknown>[] {
    return this.results.filter((entry) => entry.body.ticketId === ticketId).map((entry) => entry.body);
  }

  readonly fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const url = String(input);
    const route = url.replace(/^.*\/excel\//, "");
    const method = init?.method ?? "GET";
    const headers = (init?.headers ?? {}) as Record<string, string>;
    const body = typeof init?.body === "string" && init.body ? JSON.parse(init.body) as Record<string, unknown> : {};
    if (method === "POST" && route === "session") {
      this.sessionPosts += 1;
      const hold = this.sessionHolds.shift();
      if (hold !== undefined) {
        this.heldSessionWaiting += 1;
        const status = await hold;
        this.heldSessionWaiting -= 1;
        if (status !== null) return json(status, { error: "Companion unavailable" });
      }
      if (this.sessionFailure !== null) return json(this.sessionFailure, { error: "Companion unavailable" });
      this.sessionCounter += 1;
      const sessionId = uuid("aaaaaaaa", this.sessionCounter).replace(/^aaaaaaaa-0000-4000-8000/, "aaaaaaaa-aaaa-4aaa-8aaa");
      const sessionKey = `${"k".repeat(40)}${this.sessionCounter}`;
      const record: SessionRecord = { sessionId, sessionKey, documentId: String(body.documentId), tickets: [] };
      this.byKey.set(sessionKey, record);
      this.byId.set(sessionId, record);
      this.events.push(`session:${sessionId}`);
      return json(201, {
        ok: true, sessionId, sessionKey, protocolVersion: "1.0", expiresInSeconds: 60,
        workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
      });
    }
    if (method === "GET" && route === "pending") {
      const hold = this.pendingHolds.shift();
      if (hold !== undefined) {
        this.heldPendingWaiting += 1;
        const mode = await hold;
        this.heldPendingWaiting -= 1;
        if (mode === "network") throw new TypeError("scripted pending network reset");
      }
    }
    const session = this.byKey.get(headers["X-Hermes-Session-Key"] ?? "");
    if (session === undefined || headers["X-Hermes-Session-Id"] !== session.sessionId) {
      return json(401, { error: "Unauthorized" });
    }
    if (method === "DELETE" && route === "session") {
      this.byKey.delete(session.sessionKey);
      this.byId.delete(session.sessionId);
      for (const ticket of session.tickets) {
        if (ticket.state === "pending" || ticket.state === "claimed") ticket.state = "cancelled";
        else if (ticket.state === "executing") ticket.state = "uncertain";
      }
      this.deleted.push(session.sessionId);
      this.events.push(`delete:${session.sessionId}`);
      return json(200, { ok: true });
    }
    if (method === "POST" && route === "heartbeat") return json(200, { ok: true });
    if (method === "GET" && route === "pending") {
      const ticket = session.tickets.find((candidate) => candidate.state === "pending");
      if (ticket === undefined) return json(204);
      ticket.state = "claimed";
      this.events.push("companion-claim");
      return json(200, {
        ticketId: ticket.ticketId, requestId: ticket.requestId, tool: ticket.tool, arguments: ticket.args,
        ...(ticket.trusted === undefined ? {} : { trusted: ticket.trusted }),
      });
    }
    if (method === "POST" && route === "execute-start") {
      this.executeStarts += 1;
      const ticket = session.tickets.find((candidate) => candidate.ticketId === body.ticketId);
      if (ticket === undefined || ticket.state !== "claimed") return json(409, { error: "conflict" });
      ticket.state = "executing";
      this.events.push("execute-start");
      return json(200, { ok: true, ticketId: ticket.ticketId });
    }
    if (method === "POST" && route === "result") {
      this.events.push("result");
      this.results.push({ sessionId: session.sessionId, body });
      const script = this.resultScripts.get(String(body.ticketId));
      const step = script?.shift();
      if (step === "network") throw new TypeError("scripted result network reset");
      if (typeof step === "number" && step !== 200) return json(step, { error: `scripted result ${step}` });
      const ticket = session.tickets.find((candidate) => candidate.ticketId === body.ticketId);
      if (ticket === undefined) return json(404, { error: "No ticket found" });
      if (ticket.tool !== "excel_execute_plan") {
        if (ticket.state !== "claimed") return json(409, { error: "conflict" });
        ticket.state = "completed";
        return json(200, { ok: true, ticketId: ticket.ticketId });
      }
      if (typeof body.executed !== "boolean") return json(400, { error: "executed required" });
      if (body.executed === true && ticket.state !== "executing") return json(409, { error: "conflict" });
      if (body.executed === false && ticket.state !== "claimed") return json(409, { error: "conflict" });
      ticket.state = body.executed === true ? "completed" : "not_executed";
      return json(200, { ok: true, ticketId: ticket.ticketId });
    }
    return json(404, { error: "route" });
  };
}

function statefulAdapter(events: string[]) {
  const cells = new Map<string, unknown>();
  const adapter: WorkbookPlanExecutionAdapter & WorkbookPlanVerificationAdapter = {
    async executeFirstGroup(gate) {
      events.push("fresh");
      return gate.executeAfterFresh("fresh", async (group) => {
        events.push("mutation");
        for (const operation of group.operations) {
          cells.set(`${String(operation.args.sheet)}!${String(operation.args.address)}`, operation.args.values);
        }
        return { status: "committed" };
      });
    },
    async executeGroup(group) {
      events.push("mutation");
      for (const operation of group.operations) {
        cells.set(`${String(operation.args.sheet)}!${String(operation.args.address)}`, operation.args.values);
      }
      return { status: "committed" };
    },
    async reconcileFailedGroup(group) {
      return group.operations.map(() => "unknown_within_batch" as const);
    },
    async verifyPlan() {
      events.push("verify-read");
      return JSON.stringify(cells.get("Sheet1!A1")) === JSON.stringify([["ok"]]);
    },
  };
  return { adapter, cells };
}

interface Hooks {
  onRuntimeRead?: () => void;
}

interface SpyView {
  readonly mock: { readonly calls: readonly (readonly unknown[])[] };
}

type RouteDependencies = Parameters<typeof routeModule.createServerOwnedPlanRoute>[0];

interface SessionHarness {
  readonly events: string[];
  readonly companion: SessionCompanion;
  readonly lifecycle: OfficeBridgeLifecycleController;
  readonly bridge: () => OfficeCompanionBridge;
  readonly coordinator: TaskPaneTurnCoordinator;
  readonly bindings: ServerOwnedTurnBindingStore;
  readonly route: ReturnType<typeof routeModule.createServerOwnedPlanRoute>;
  readonly cells: Map<string, unknown>;
  readonly intervals: (() => void)[];
  readonly docs: { current: string };
  readonly hooks: Hooks;
  readonly state: {
    active: ActiveTurnView | null;
    activeRequest: TrustedWorkbookRequest | null;
    activeTurnStart: number | null;
    last: { sessionId: string | null; documentId: string | null };
    forcedStops: number;
    planExecuteToolCalls: number;
  };
  readonly awaitSpy: SpyView;
  readonly consumeSpy: SpyView;
  readonly routeNow: SpyView;
  beginTurn(turnId: string, prompt?: string): { handle: TaskPaneTurnHandle; controller: AbortController };
  prepare(handle: TaskPaneTurnHandle): Promise<TaskPaneTurnPreparation>;
  stop(): void;
  finish(handle: TaskPaneTurnHandle): void;
}

const cleanups: (() => Promise<void> | void)[] = [];

afterEach(async () => {
  while (cleanups.length > 0) {
    try { await cleanups.pop()!(); } catch { /* best-effort test teardown */ }
  }
  vi.useRealTimers();
});

async function sessionHarness(): Promise<SessionHarness> {
  const events: string[] = [];
  const companion = new SessionCompanion(events);
  const intervals: (() => void)[] = [];
  const docs = { current: DOC1 };
  const hooks: Hooks = {};
  const bindings = new ServerOwnedTurnBindingStore({ now: () => Date.now(), ttlMs: 90_000, maxEntries: 32, maxBytes: 64 * 1024 });
  const state: SessionHarness["state"] = {
    active: null,
    activeRequest: null,
    activeTurnStart: null,
    last: { sessionId: null, documentId: null },
    forcedStops: 0,
    planExecuteToolCalls: 0,
  };
  let route!: ReturnType<typeof routeModule.createServerOwnedPlanRoute>;
  let bridge: OfficeCompanionBridge | undefined;
  // eslint-disable-next-line prefer-const
  let coordinator: TaskPaneTurnCoordinator;
  /** Test copy of taskpane.ts updateBridgeState's turn/binding decision (copied from r3b, verbatim base fallback kept). */
  const onState = (next: BridgePublicState): void => {
    const previous = { sessionId: state.last.sessionId, documentId: state.last.documentId };
    const api = r3b(coordinator);
    let invalidated: boolean;
    if (typeof api.bridgeStateInvalidatesActiveTurn === "function") {
      invalidated = api.bridgeStateInvalidatesActiveTurn(next, state.active, previous);
    } else {
      // taskpane.ts:407-413 @39ab0a2, verbatim decision.
      const identityChanged = previous.sessionId !== next.sessionId || previous.documentId !== next.documentId;
      invalidated = state.active !== null &&
        (!next.connected ||
          state.active.sessionId !== next.sessionId ||
          (state.active.documentId !== null && state.active.documentId !== next.documentId) ||
          identityChanged);
    }
    if (invalidated) {
      state.forcedStops += 1;
      coordinator.stopCurrentTurn();
      state.active = null;
    }
    if (previous.sessionId !== null &&
        (previous.sessionId !== next.sessionId || previous.documentId !== next.documentId)) {
      bindings.invalidateSession(previous.sessionId);
    }
    if (previous.documentId !== null && previous.documentId !== next.documentId) {
      bindings.invalidateDocument(previous.documentId);
    }
    state.last = { sessionId: next.sessionId, documentId: next.documentId };
  };
  const lifecycle = new OfficeBridgeLifecycleController((emit) => {
    bridge = createExcelCompanionBridge(
      emit,
      (name, argumentsValue, trusted, control) => {
        if (name === "excel_execute_plan") state.planExecuteToolCalls += 1;
        return route.executeCompanionTool(name, argumentsValue, trusted, control);
      },
      {
        fetchImpl: companion.fetch as typeof fetch,
        documentId: async () => docs.current,
        paneId: PANE,
        setIntervalImpl: ((handler: () => void) => { intervals.push(handler); return intervals.length; }) as unknown as typeof setInterval,
        clearIntervalImpl: (() => undefined) as unknown as typeof clearInterval,
      },
    );
    return bridge;
  }, onState);
  coordinator = new TaskPaneTurnCoordinator({
    getOrCreateDocumentId: async () => docs.current,
    getBridgeIdentity: () => ({ sessionId: lifecycle.sessionId, documentId: lifecycle.documentId }),
    reconnectForDocumentChange: (expected) => lifecycle.reconnectForDocumentChange(expected),
    bindings,
    paneId: PANE,
  });
  const { adapter, cells } = statefulAdapter(events);
  const storage = new MemoryIdempotencyStorage();
  const awaitSpy = vi.fn((binding: WorkbookPlanConfirmationBinding, signal?: AbortSignal) =>
    coordinator.waitForWorkbookPlanConfirmation(binding, signal));
  const consumeSpy = vi.spyOn(coordinator, "consumeWorkbookPlanConfirmation");
  const routeNow = vi.fn(() => Date.now());
  let runtimeReadHookFired = false;
  const dependencies = {
    isRolloutEnabled: () => true,
    getBridgeSessionId: () => lifecycle.sessionId,
    getActiveTurn: () => ({
      turnOperationId: state.active?.turnId ?? null,
      trustedRequest: state.activeRequest,
      sessionId: state.active?.sessionId ?? null,
      turnStart: state.activeTurnStart,
      hasCancellation: state.active !== null,
    }),
    coordinator,
    bindings,
    readRuntimeContext: async () => {
      events.push("runtime-read");
      if (!runtimeReadHookFired && hooks.onRuntimeRead !== undefined) {
        runtimeReadHookFired = true;
        hooks.onRuntimeRead();
      }
      return RUNTIME;
    },
    createAdapter: () => adapter,
    createStore: () => new BrowserIdempotencyStore(storage),
    inspect: async () => { events.push("inspect-read"); return "{}"; },
    awaitPaneConfirmation: awaitSpy,
    executePrimitive: async (name: string) => {
      events.push(`primitive:${name}`);
      return JSON.stringify({ ok: true, data: [[1]] });
    },
    now: routeNow,
  };
  route = routeModule.createServerOwnedPlanRoute(dependencies as unknown as RouteDependencies);
  await lifecycle.start();
  const harness: SessionHarness = {
    events, companion, lifecycle, bridge: () => bridge!, coordinator, bindings, route, cells, intervals, docs, hooks,
    state, awaitSpy, consumeSpy, routeNow,
    beginTurn(turnId, prompt = EXPLICIT_PROMPT) {
      const controller = new AbortController();
      const trustedRequest = classifyWorkbookRequest(prompt);
      const turnStart = Date.now();
      // sendPrompt order: identity first (activeTurnSessionId = bridge session), then beginTurn.
      state.active = { turnId, sessionId: lifecycle.sessionId, documentId: null };
      state.activeRequest = trustedRequest;
      state.activeTurnStart = turnStart;
      const handle = coordinator.beginTurn(turnId, controller, { trustedRequest, turnStart, cancellationSignal: controller.signal });
      return { handle, controller };
    },
    async prepare(handle) {
      const preparation = await coordinator.prepareTurn(handle);
      if (preparation.status === "active" && state.active?.turnId === handle.turnId) {
        state.active.sessionId = preparation.sessionId;
        state.active.documentId = preparation.documentId;
      }
      return preparation;
    },
    stop() {
      coordinator.stopCurrentTurn();
      state.active = null;
    },
    finish(handle) {
      coordinator.finishTurn(handle);
      if (state.active?.turnId === handle.turnId) state.active = null;
    },
  };
  cleanups.push(async () => {
    coordinator.stopCurrentTurn();
    await lifecycle.stop();
  });
  return harness;
}

/** Flushes real macrotasks without touching faked setTimeout. */
async function flushUntil(condition: () => boolean): Promise<void> {
  for (let attempt = 0; attempt < 400 && !condition(); attempt += 1) {
    await new Promise<void>((done) => setImmediate(done));
  }
  expect(condition()).toBe(true);
}

function sessionOf(preparation: TaskPaneTurnPreparation): string | null | undefined {
  return preparation.status === "active" ? preparation.sessionId : undefined;
}

function diagnostics(h: SessionHarness): string {
  return `diag: deletes=${JSON.stringify(h.companion.deleted)} sessionPosts=${h.companion.sessionPosts} ` +
    `executeStarts=${h.companion.executeStarts} mutations=${h.events.filter((e) => e === "mutation").length} ` +
    `cells=${h.cells.size} forcedStops=${h.state.forcedStops} results=${JSON.stringify(h.companion.results.map((r) => r.body))}`;
}

/** T1 runs on S1 and issues a Hermes request; `end` decides how T1 ends. */
async function runT1WithHermesRequest(h: SessionHarness, end: (handle: TaskPaneTurnHandle) => void, clean = false) {
  const s1 = h.lifecycle.sessionId!;
  const t1 = h.beginTurn(T1);
  const prepared = await h.prepare(t1.handle);
  expect(prepared.status).toBe("active");
  const request = r3b(h.coordinator).beginHermesRequest?.(s1);
  if (clean) request?.endClean();
  end(t1.handle);
  return { s1, t1, request };
}
/** R4f: continueTurn with the harness's sendPrompt bookkeeping (the active turn adopts the prepared identity). */
async function continueT(h: SessionHarness, handle: TaskPaneTurnHandle, continuation: (preparation: TaskPaneTurnPreparation) => unknown = () => "ran") {
  return h.coordinator.continueTurn(handle, async (preparation) => {
    if (h.state.active?.turnId === handle.turnId) {
      h.state.active.sessionId = preparation.sessionId;
      h.state.active.documentId = preparation.documentId;
    }
    return continuation(preparation);
  });
}

async function flushMacrotasks(turns = 40): Promise<void> {
  for (let turn = 0; turn < turns; turn += 1) await new Promise<void>((done) => setImmediate(done));
}

/** One clean rollout-on turn: prepare -> mark -> endClean -> finish. Returns whether prepare rotated and the mark count after prepare. */
async function cleanTurn(h: SessionHarness, turnId: string): Promise<{ rotated: boolean; marksAfterPrepare: number; sessionId: string | null | undefined }> {
  const before = h.lifecycle.sessionId;
  const turn = h.beginTurn(turnId);
  const preparation = await h.prepare(turn.handle);
  expect(preparation.status, diagnostics(h)).toBe("active");
  const marksAfterPrepare = h.coordinator.hermesRequestMarkCount;
  h.coordinator.beginHermesRequest(sessionOf(preparation) ?? null).endClean();
  h.finish(turn.handle);
  return { rotated: sessionOf(preparation) !== before, marksAfterPrepare, sessionId: sessionOf(preparation) };
}

/** T1 leaves an unclean mark on S1; T2's prepare rotates and its POST /session is held. */
async function rotatingT2WithHeldSession(h: SessionHarness) {
  const { s1 } = await runT1WithHermesRequest(h, () => h.stop());
  const hold = h.companion.holdNextSession();
  const t2 = h.beginTurn(T2);
  const p2 = h.prepare(t2.handle);
  await flushUntil(() => h.companion.heldSessionWaiting === 1);
  expect(h.companion.deleted).toEqual([s1]);
  expect(h.lifecycle.sessionId).toBeNull();
  return { s1, hold, t2, p2 };
}

/** Verbatim base decision (r3b :374-381 copy of taskpane.ts updateBridgeState). */
function baseDecision(
  next: BridgePublicState,
  active: ActiveTurnView | null,
  previous: { sessionId: string | null; documentId: string | null },
): boolean {
  const identityChanged = previous.sessionId !== next.sessionId || previous.documentId !== next.documentId;
  return active !== null &&
    (!next.connected ||
      active.sessionId !== next.sessionId ||
      (active.documentId !== null && active.documentId !== next.documentId) ||
      identityChanged);
}

describe("R4f A3: marks and overflow taints of replaced sessions are pruned at prepare", () => {
  it("LA3 (P8) 64 unclean marks on sessions replaced by stop()/start(), then 5 clean turns: no rotation, marks 0 after the first prepare", async () => {
    const h = await sessionHarness();
    for (let round = 0; round < 4; round += 1) {
      const replaced = h.lifecycle.sessionId!;
      for (let mark = 0; mark < 16; mark += 1) h.coordinator.beginHermesRequest(replaced);
      await h.lifecycle.stop();
      await h.lifecycle.start();
    }
    expect(h.coordinator.hermesRequestMarkCount).toBe(64);
    const posts = h.companion.sessionPosts;
    const deletes = [...h.companion.deleted];
    expect(posts).toBe(5);
    expect(deletes).toHaveLength(4);
    const turns = [];
    for (let n = 0; n < 5; n += 1) turns.push(await cleanTurn(h, uuid("7a7a0000", 10 + n)));
    const diag = `turns=${JSON.stringify(turns)} ${diagnostics(h)}`;
    expect(turns.map((turn) => turn.rotated), diag).toEqual([false, false, false, false, false]);
    expect(turns[0].marksAfterPrepare, diag).toBe(0);
    expect(h.companion.sessionPosts, diag).toBe(posts);
    expect(h.companion.deleted, diag).toEqual(deletes);
  });

  it("LA3b a full table plus an overflow taint on S_old; S_old replaced by stop()/start(): the next turns do not rotate", async () => {
    const h = await sessionHarness();
    const old = h.lifecycle.sessionId!;
    for (let mark = 0; mark < 65; mark += 1) h.coordinator.beginHermesRequest(old);
    expect(h.coordinator.hermesRequestMarkCount).toBe(64);
    await h.lifecycle.stop();
    await h.lifecycle.start();
    const posts = h.companion.sessionPosts;
    const turns = [await cleanTurn(h, T2), await cleanTurn(h, T3)];
    const diag = `turns=${JSON.stringify(turns)} ${diagnostics(h)}`;
    expect(turns.map((turn) => turn.rotated), diag).toEqual([false, false]);
    expect(turns[0].marksAfterPrepare, diag).toBe(0);
    expect(h.companion.sessionPosts, diag).toBe(posts);
    expect(h.companion.deleted, diag).toEqual([old]);
  });

  it("LA3d an unclean mark on S1; S1 replaced by a current-session 401 reconnect (heartbeat handler): next turn does not rotate, marks 0", async () => {
    const h = await sessionHarness();
    const s1 = h.lifecycle.sessionId!;
    h.coordinator.beginHermesRequest(s1);
    const record = h.companion.byId.get(s1)!;
    h.companion.byKey.delete(record.sessionKey);
    const heartbeat = h.intervals[h.intervals.length - 2];
    heartbeat();
    await flushUntil(() => h.lifecycle.sessionId !== null && h.lifecycle.sessionId !== s1);
    const s2 = h.lifecycle.sessionId;
    const posts = h.companion.sessionPosts;
    const turn = await cleanTurn(h, T2);
    const diag = `turn=${JSON.stringify(turn)} ${diagnostics(h)}`;
    expect(turn.rotated, diag).toBe(false);
    expect(turn.sessionId, diag).toBe(s2);
    expect(turn.marksAfterPrepare, diag).toBe(0);
    expect(h.companion.sessionPosts, diag).toBe(posts);
  });

  it("LA3c (parity) an unclean mark on the current session still rotates the next turn (RB1 shape)", async () => {
    const h = await sessionHarness();
    const { s1 } = await runT1WithHermesRequest(h, () => h.stop());
    const t2 = h.beginTurn(T2);
    const p2 = await h.prepare(t2.handle);
    expect(p2.status).toBe("active");
    expect(sessionOf(p2)).not.toBe(s1);
    expect(sessionOf(p2)).toBe(h.lifecycle.sessionId);
    expect(h.companion.deleted).toEqual([s1]);
    expect(h.coordinator.hermesRequestMarkCount).toBe(0);
  });
});

describe("R4f A4: a prepare waits out another turn's rotation (raced with its own signal)", () => {
  it("LA4 (P1c) Stop during T2's rotation, then T3: T3 goes active on S2, is not force-stopped, and its continuation runs once", async () => {
    const h = await sessionHarness();
    const { s1, hold, t2, p2 } = await rotatingT2WithHeldSession(h);
    h.stop();
    const t3 = h.beginTurn(T3);
    const continuation = vi.fn(() => "ran");
    const c3 = continueT(h, t3.handle, continuation);
    await flushMacrotasks();
    hold.release();
    const r3 = await c3;
    const r2 = await p2;
    await flushMacrotasks();
    const diag = `r2=${r2.status} r3=${JSON.stringify({ status: r3.status, sessionId: sessionOf(r3) })} lifecycle=${String(h.lifecycle.sessionId)} ${diagnostics(h)}`;
    expect(r2.status, diag).toBe("stale");
    expect(r3.status, diag).toBe("active");
    expect(sessionOf(r3), diag).not.toBeNull();
    expect(sessionOf(r3), diag).not.toBe(s1);
    expect(sessionOf(r3), diag).toBe(h.lifecycle.sessionId);
    expect(continuation, diag).toHaveBeenCalledTimes(1);
    expect(t3.controller.signal.aborted, diag).toBe(false);
    expect(h.state.forcedStops, diag).toBe(0);
    expect(h.companion.deleted, diag).toEqual([s1]);
    expect(h.companion.sessionPosts, diag).toBe(2);
    expect(h.coordinator.isCurrent(t3.handle.generation, T3, t3.controller), diag).toBe(true);
    expect(h.coordinator.isCurrent(t2.handle.generation, T2, t2.controller)).toBe(false);
  });

  it("LA4s T3 is stopped while it awaits T2's rotation: its prepare resolves stale before the release, no continuation, no active turn after", async () => {
    const h = await sessionHarness();
    const { hold, t2, p2 } = await rotatingT2WithHeldSession(h);
    h.stop();
    const t3 = h.beginTurn(T3);
    const continuation = vi.fn(() => "ran");
    let settled: Awaited<ReturnType<typeof continueT>> | null = null;
    const c3 = continueT(h, t3.handle, continuation).then((result) => { settled = result; return result; });
    await flushMacrotasks();
    h.stop();
    await flushUntil(() => settled !== null);
    const beforeRelease = settled as Awaited<ReturnType<typeof continueT>> | null;
    hold.release();
    const r2 = await p2;
    await c3;
    await flushMacrotasks();
    const diag = `beforeRelease=${String(beforeRelease?.status)} r2=${r2.status} ${diagnostics(h)}`;
    expect(beforeRelease?.status, diag).toBe("stale");
    expect(continuation, diag).not.toHaveBeenCalled();
    expect(r2.status, diag).toBe("stale");
    expect(h.coordinator.isCurrent(t3.handle.generation, T3, t3.controller), diag).toBe(false);
    expect(h.coordinator.isCurrent(t2.handle.generation, T2, t2.controller), diag).toBe(false);
    expect(h.state.active, diag).toBeNull();
  });

  it("LA4f the held rotation fails 503 (companion healthy afterwards): T2 stale, T3 makes one reconnect attempt and goes active on S2", async () => {
    const h = await sessionHarness();
    const { s1, hold, p2 } = await rotatingT2WithHeldSession(h);
    h.stop();
    const t3 = h.beginTurn(T3);
    const continuation = vi.fn(() => "ran");
    const c3 = continueT(h, t3.handle, continuation);
    await flushMacrotasks();
    const postsBeforeRelease = h.companion.sessionPosts;
    hold.release(503);
    const r2 = await p2;
    const r3 = await c3;
    await flushMacrotasks();
    const diag = `r2=${r2.status} r3=${JSON.stringify({ status: r3.status, sessionId: sessionOf(r3) })} ${diagnostics(h)}`;
    expect(r2.status, diag).toBe("stale");
    expect(r3.status, diag).toBe("active");
    expect(sessionOf(r3), diag).not.toBeNull();
    expect(sessionOf(r3), diag).toBe(h.lifecycle.sessionId);
    expect(h.companion.sessionPosts - postsBeforeRelease, diag).toBe(1);
    expect(continuation, diag).toHaveBeenCalledTimes(1);
    expect(t3.controller.signal.aborted, diag).toBe(false);
    expect(h.state.forcedStops, diag).toBe(0);
    expect(h.companion.deleted, diag).toEqual([s1]);
  });

  it("LA4n (parity) an unprepared current turn that is not awaiting (prepareTurn never called) is force-stopped by the S2 emission as at the base", async () => {
    const h = await sessionHarness();
    const { hold, p2 } = await rotatingT2WithHeldSession(h);
    h.stop();
    const t3 = h.beginTurn(T3);
    hold.release();
    const r2 = await p2;
    await flushMacrotasks();
    expect(r2.status).toBe("stale");
    expect(h.state.forcedStops, diagnostics(h)).toBe(1);
    expect(t3.controller.signal.aborted).toBe(true);
    expect(h.coordinator.isCurrent(t3.handle.generation, T3, t3.controller)).toBe(false);
  });
});

describe("R4f A5: a late background error from the replaced session never tears down the new one", () => {
  async function rotationWithHeldS1Poll(h: SessionHarness) {
    const { s1 } = await runT1WithHermesRequest(h, () => h.stop());
    const hold = h.companion.holdNextPending();
    const pollS1 = h.intervals[h.intervals.length - 1];
    pollS1();
    await flushUntil(() => h.companion.heldPendingWaiting === 1);
    const t2 = h.beginTurn(T2);
    const p2 = await h.prepare(t2.handle);
    expect(p2.status, diagnostics(h)).toBe("active");
    const s2 = sessionOf(p2)!;
    expect(s2).not.toBe(s1);
    expect(h.companion.deleted).toEqual([s1]);
    return { s1, s2, hold, t2 };
  }

  it("LA5 (P5c) an S1 poll fired before the rotation is answered 401 after T2 is prepared on S2: S2 kept, T2 not aborted, its own plan commits once", async () => {
    const h = await sessionHarness();
    const { s2, hold, t2 } = await rotationWithHeldS1Poll(h);
    hold.release("process");
    await flushMacrotasks();
    const diag = `lifecycle=${String(h.lifecycle.sessionId)} ${diagnostics(h)}`;
    expect(h.lifecycle.sessionId, diag).toBe(s2);
    expect(t2.controller.signal.aborted, diag).toBe(false);
    expect(h.state.forcedStops, diag).toBe(0);
    expect(h.companion.sessionPosts, diag).toBe(2);
    const own = h.companion.invoke(s2, T2_OP, EXECUTE_PLAN);
    await expect(h.bridge().pollOnce()).resolves.toBe(true);
    expect(h.companion.executeStarts, diagnostics(h)).toBe(1);
    expect(h.events.filter((event) => event === "mutation")).toHaveLength(1);
    expect(h.cells.get("Sheet1!A1")).toEqual([["ok"]]);
    expect(h.companion.resultsFor(own.ticketId)[0]).toMatchObject({ executed: true, result: { ok: true } });
  });

  it("LA5n the held S1 poll fails with a network error instead: T2 is not aborted", async () => {
    const h = await sessionHarness();
    const { s2, hold, t2 } = await rotationWithHeldS1Poll(h);
    hold.release("network");
    await flushMacrotasks();
    const diag = `lifecycle=${String(h.lifecycle.sessionId)} ${diagnostics(h)}`;
    expect(t2.controller.signal.aborted, diag).toBe(false);
    expect(h.state.forcedStops, diag).toBe(0);
    expect(h.lifecycle.sessionId, diag).toBe(s2);
    expect(h.coordinator.isCurrent(t2.handle.generation, T2, t2.controller), diag).toBe(true);
  });
});

describe("R4f A6 (D-R4-8): one bounded reconnect attempt per sessionless prepare after a failed rotation", () => {
  async function failedRotation(h: SessionHarness) {
    const { s1 } = await runT1WithHermesRequest(h, () => h.stop());
    h.companion.sessionFailure = 503;
    const t2 = h.beginTurn(T2);
    const continuation = vi.fn();
    const r2 = await h.coordinator.continueTurn(t2.handle, continuation);
    expect(r2).toMatchObject({ status: "rejected", reason: "bridge-reconnect-failed" });
    expect(continuation).not.toHaveBeenCalled();
    expect(h.lifecycle.sessionId).toBeNull();
    h.finish(t2.handle);
    return { s1 };
  }

  it("LA6 (P3a) companion healthy again, no lifecycle.start(): T3's prepare makes exactly one POST /session and goes active on S2", async () => {
    const h = await sessionHarness();
    const { s1 } = await failedRotation(h);
    h.companion.sessionFailure = null;
    const posts = h.companion.sessionPosts;
    const t3 = h.beginTurn(T3);
    const continuation = vi.fn(() => "ran");
    const r3 = await continueT(h, t3.handle, continuation);
    const diag = `r3=${JSON.stringify({ status: r3.status, sessionId: sessionOf(r3) })} ${diagnostics(h)}`;
    expect(h.companion.sessionPosts - posts, diag).toBe(1);
    expect(r3.status, diag).toBe("active");
    expect(sessionOf(r3), diag).not.toBeNull();
    expect(sessionOf(r3), diag).not.toBe(s1);
    expect(sessionOf(r3), diag).toBe(h.lifecycle.sessionId);
    expect(continuation, diag).toHaveBeenCalledTimes(1);
    expect(h.state.forcedStops, diag).toBe(0);
  });

  it("LA6b still failing at T3: T3 rejected after exactly one attempt; healthy at T4: T4 makes one attempt and goes active (OQ-F1 re-arm)", async () => {
    const h = await sessionHarness();
    await failedRotation(h);
    const posts3 = h.companion.sessionPosts;
    const t3 = h.beginTurn(T3);
    const c3 = vi.fn();
    const r3 = await continueT(h, t3.handle, c3);
    const attempts3 = h.companion.sessionPosts - posts3;
    h.finish(t3.handle);
    h.companion.sessionFailure = null;
    const posts4 = h.companion.sessionPosts;
    const t4 = h.beginTurn(T4);
    const c4 = vi.fn(() => "ran");
    const r4 = await continueT(h, t4.handle, c4);
    const attempts4 = h.companion.sessionPosts - posts4;
    const diag = `r3=${JSON.stringify(r3)} attempts3=${attempts3} r4=${JSON.stringify({ status: r4.status, sessionId: sessionOf(r4) })} attempts4=${attempts4} ${diagnostics(h)}`;
    expect(r3, diag).toMatchObject({ status: "rejected", reason: "bridge-reconnect-failed" });
    expect(attempts3, diag).toBe(1);
    expect(c3, diag).not.toHaveBeenCalled();
    expect(attempts4, diag).toBe(1);
    expect(r4.status, diag).toBe("active");
    expect(sessionOf(r4), diag).toBe(h.lifecycle.sessionId);
    expect(sessionOf(r4), diag).not.toBeNull();
    expect(c4, diag).toHaveBeenCalledTimes(1);
  });

  it("LA6c (parity) after a failed rotation, lifecycle.start() (visibility) gives a session: the next prepare makes no attempt and goes active on it", async () => {
    const h = await sessionHarness();
    const { s1 } = await failedRotation(h);
    h.companion.sessionFailure = null;
    await h.lifecycle.start();
    const restarted = h.lifecycle.sessionId;
    const posts = h.companion.sessionPosts;
    const t3 = h.beginTurn(T3);
    const p3 = await h.prepare(t3.handle);
    expect(p3.status).toBe("active");
    expect(sessionOf(p3)).toBe(restarted);
    expect(sessionOf(p3)).not.toBe(s1);
    expect(h.companion.sessionPosts).toBe(posts);
    h.finish(t3.handle);
    const t4 = h.beginTurn(T4);
    const p4 = await h.prepare(t4.handle);
    expect(sessionOf(p4)).toBe(restarted);
    expect(h.companion.sessionPosts).toBe(posts);
  });
});

describe("R4f LP11: with no rotation in flight the force-stop policy equals the base decision", () => {
  it("LP11 all 1620 combinations match the verbatim base decision (no current turn, and a current unprepared turn)", async () => {
    const h = await sessionHarness();
    const S1 = "aaaaaaaa-aaaa-4aaa-8aaa-000000000001";
    const S2 = "aaaaaaaa-aaaa-4aaa-8aaa-000000000002";
    const sessions = [null, S1, S2];
    const documents = [null, DOC1, DOC2];
    const actives: (ActiveTurnView | null)[] = [null];
    for (const sessionId of sessions) for (const documentId of documents) actives.push({ turnId: T1, sessionId, documentId });
    const evaluate = (): { total: number; mismatches: string[] } => {
      let total = 0;
      const mismatches: string[] = [];
      for (const connected of [true, false]) {
        for (const stateSession of sessions) {
          for (const stateDocument of documents) {
            for (const active of actives) {
              for (const previousSession of sessions) {
                for (const previousDocument of documents) {
                  const next: BridgePublicState = { connected, sessionId: stateSession, documentId: stateDocument };
                  const previous = { sessionId: previousSession, documentId: previousDocument };
                  const got = h.coordinator.bridgeStateInvalidatesActiveTurn(next, active, previous);
                  const want = baseDecision(next, active, previous);
                  total += 1;
                  if (got !== want) mismatches.push(JSON.stringify({ next, active, previous, got, want }));
                }
              }
            }
          }
        }
      }
      return { total, mismatches };
    };
    const idle = evaluate();
    expect(idle.total).toBe(1620);
    expect(idle.mismatches.slice(0, 5)).toEqual([]);
    h.beginTurn(T1);
    const unprepared = evaluate();
    expect(unprepared.total).toBe(1620);
    expect(unprepared.mismatches.slice(0, 5)).toEqual([]);
  });
});
