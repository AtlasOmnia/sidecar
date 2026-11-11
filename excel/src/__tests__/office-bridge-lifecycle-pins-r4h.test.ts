import { afterEach, describe, expect, it, vi } from "vitest";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import { OfficeCompanionBridge, type BridgePublicState } from "../officeBridge";
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
 * R4h lifecycle pins (dev3-r4h-release OMH RELEASE RECORD 2026-10-07; r4f
 * tester-report-A F1-F3). Each row pins EXISTING R4f behaviour at ac981da (no
 * production change):
 * - F1 (OQ-F7) the bridge timer callback captures (session, generation) at the
 *   tick: a second current-session 401 after a 401 auto-reconnect reconnects
 *   again;
 * - F2 a prepare that finds a session disarms the one-shot reconnect retry;
 * - F3 the A4 waiter re-checks currency right after the rotation wait (abort
 *   plus a failed rotation): no reconnect attempt is made for a stopped turn.
 *
 * Helpers are COPIED as text, never imported (both source files stay
 * byte-identical):
 * - from office-bridge-lifecycle-r4f.test.ts at ac981da: lines 16-22, 41-54,
 *   72-87, 97-100, 109-120 and 137-142 (bridge constants, response,
 *   sessionResponse, routedFetch, count, timedBridge, waitUntil);
 * - from task-pane-lifecycle-r4f.test.ts at ac981da: lines 41-101, 103-319,
 *   321-517, 519-561 and 575-585 (constants, SessionCompanion, statefulAdapter,
 *   sessionHarness, flushUntil, sessionOf, diagnostics, runT1WithHermesRequest,
 *   continueT, flushMacrotasks, rotatingT2WithHeldSession) and the
 *   failedRotation body of :813-824. The one extension of the copy is the
 *   Hooks.onReconnect observer in sessionHarness (F3).
 */

// --- copied: office-bridge-lifecycle-r4f.test.ts:16-22 ---
const SESSION_ID = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const SESSION_KEY = "s".repeat(64);
const SESSION_ID_2 = "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee";
const SESSION_KEY_2 = "t".repeat(64);
const SESSION_ID_3 = "ffffffff-ffff-4fff-8fff-ffffffffffff";
const SESSION_KEY_3 = "u".repeat(64);
const PANE_ID = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
// --- copied: office-bridge-lifecycle-r4f.test.ts:41-54 ---
function response(status: number, body?: unknown): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
  });
}

function sessionResponse(sessionId = SESSION_ID, sessionKey = SESSION_KEY): Response {
  return response(201, {
    ok: true, sessionId, sessionKey,
    protocolVersion: "1.0", expiresInSeconds: 60,
    workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
  });
}
// --- copied: office-bridge-lifecycle-r4f.test.ts:72-87 ---
type FetchRoute = (init: RequestInit | undefined) => Promise<Response> | Response;

/** Route-keyed fetch double; unknown routes fail loudly (copied from office-bridge-task13.test.ts). */
function routedFetch(routes: Record<string, FetchRoute[]>) {
  const calls: { url: string; init?: RequestInit }[] = [];
  const fetchImpl = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    calls.push({ url, init });
    const key = `${init?.method ?? "GET"} ${url}`;
    const queue = routes[key];
    if (queue === undefined || queue.length === 0) throw new Error(`unexpected fetch ${key}`);
    const next = queue.length > 1 ? queue.shift()! : queue[0];
    return next(init);
  });
  return { fetchImpl: fetchImpl as unknown as typeof fetch, calls, urls: () => calls.map((call) => call.url) };
}
// --- copied: office-bridge-lifecycle-r4f.test.ts:97-100 ---
function count(urls: string[], url: string, calls?: { url: string; init?: RequestInit }[], method?: string): number {
  if (calls !== undefined && method !== undefined) return calls.filter((call) => call.url === url && (call.init?.method ?? "GET") === method).length;
  return urls.filter((entry) => entry === url).length;
}
// --- copied: office-bridge-lifecycle-r4f.test.ts:109-120 ---
function timedBridge(fetchImpl: typeof fetch, executeTool: ConstructorParameters<typeof OfficeCompanionBridge>[0]["executeTool"]) {
  const handlers: (() => void)[] = [];
  const states: BridgePublicState[] = [];
  const bridge = new OfficeCompanionBridge({
    hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
    executeTool, fetchImpl,
    onState: (state) => { states.push({ ...state }); },
    setIntervalImpl: ((handler: () => void) => { handlers.push(handler); return handlers.length; }) as unknown as typeof setInterval,
    clearIntervalImpl: (() => undefined) as unknown as typeof clearInterval,
  });
  return { bridge, handlers, states };
}
// --- copied: office-bridge-lifecycle-r4f.test.ts:137-142 ---
async function waitUntil(condition: () => boolean, diag: () => string = () => ""): Promise<void> {
  for (let attempt = 0; attempt < 400 && !condition(); attempt += 1) {
    await new Promise<void>((done) => setImmediate(done));
  }
  expect(condition(), diag()).toBe(true);
}
// --- copied: task-pane-lifecycle-r4f.test.ts:41-101 ---
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
// --- copied: task-pane-lifecycle-r4f.test.ts:103-319 ---
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
// --- copied: task-pane-lifecycle-r4f.test.ts:321-517 (+ Hooks.onReconnect, R4h) ---
interface Hooks {
  onRuntimeRead?: () => void;
  /** R4h extension of the copy (F3): sees each coordinator reconnect promise as it is created. */
  onReconnect?: (reconnect: Promise<unknown>) => void;
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
    reconnectForDocumentChange: (expected) => {
      const reconnect = lifecycle.reconnectForDocumentChange(expected);
      hooks.onReconnect?.(reconnect);
      return reconnect;
    },
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
// --- copied: task-pane-lifecycle-r4f.test.ts:519-561 ---
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
// --- copied: task-pane-lifecycle-r4f.test.ts:575-585 ---
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

// ---------------------------------------------------------------------------
// R4h F1 (r4f lane-A F1, OQ-F7)
// ---------------------------------------------------------------------------
describe("R4h F1 (OQ-F7): the bridge timer callback captures the session at the tick, not at timer creation", () => {
  it("F1 a second current-session 401 after a 401 auto-reconnect reconnects again (S1 -> S2 -> S3) through the same live poll timer", async () => {
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse(), () => sessionResponse(SESSION_ID_2, SESSION_KEY_2), () => sessionResponse(SESSION_ID_3, SESSION_KEY_3)],
      "GET /excel/pending": [() => response(401, { error: "Unauthorized" }), () => response(401, { error: "Unauthorized" }), () => response(204)],
      "POST /excel/heartbeat": [() => response(200, { ok: true })],
    });
    const { bridge, handlers } = timedBridge(routes.fetchImpl, vi.fn(async () => "{}"));
    await bridge.start();
    expect(bridge.sessionId).toBe(SESSION_ID);
    const timers = handlers.length;
    handlers[1]();
    await waitUntil(() => bridge.sessionId === SESSION_ID_2);
    await flushMacrotasks();
    const diag = () => `session=${String(bridge.sessionId)} timers=${handlers.length}/${timers} posts=${count(routes.urls(), "/excel/session", routes.calls, "POST")} urls=${JSON.stringify(routes.urls())}`;
    // The 401 auto-reconnect keeps the original interval timers: handler 1 is still the live poll timer.
    expect(handlers.length, diag()).toBe(timers);
    handlers[1]();
    await waitUntil(() => bridge.sessionId === SESSION_ID_3, diag);
    await flushMacrotasks();
    const pending = routes.calls.filter((call) => call.url === "/excel/pending");
    expect(pending.map((call) => (call.init?.headers as Record<string, string>)["X-Hermes-Session-Id"]), diag()).toEqual([SESSION_ID, SESSION_ID_2]);
    expect(count(routes.urls(), "/excel/session", routes.calls, "POST"), diag()).toBe(3);
    expect(bridge.sessionId, diag()).toBe(SESSION_ID_3);
  });
});

// ---------------------------------------------------------------------------
// R4h F2 / F3 (r4f lane-A F2, F3): A6 retry disarm and the A4 post-wait re-check
// ---------------------------------------------------------------------------
describe("R4h F2/F3: the one-shot reconnect retry is disarmed by a found session and never spent by a stopped waiter", () => {
  /** Copied from task-pane-lifecycle-r4f.test.ts:813-824 (failedRotation). */
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

  it("F2 failed rotation, then a visibility start gives a session and T3's prepare finds it; after the bridge loses that session by stop(), T4's sessionless prepare makes no reconnect attempt", async () => {
    const h = await sessionHarness();
    const { s1 } = await failedRotation(h);
    h.companion.sessionFailure = null;
    await h.lifecycle.start();
    const restarted = h.lifecycle.sessionId;
    expect(restarted).not.toBeNull();
    expect(restarted).not.toBe(s1);
    const t3 = h.beginTurn(T3);
    const p3 = await h.prepare(t3.handle);
    expect(p3.status, diagnostics(h)).toBe("active");
    expect(sessionOf(p3)).toBe(restarted);
    h.finish(t3.handle);
    await h.lifecycle.stop();
    expect(h.lifecycle.sessionId).toBeNull();
    const posts = h.companion.sessionPosts;
    const t4 = h.beginTurn(T4);
    const p4 = await h.prepare(t4.handle);
    const diag = `p4=${JSON.stringify({ status: p4.status, sessionId: sessionOf(p4) })} lifecycle=${String(h.lifecycle.sessionId)} ${diagnostics(h)}`;
    expect(h.companion.sessionPosts - posts, diag).toBe(0);
    expect(h.lifecycle.sessionId, diag).toBeNull();
    h.finish(t4.handle);
  });

  it("F3 T3 awaits T2's held rotation and is stopped in the same job in which that rotation fails: T3 is stale with no reconnect attempt of its own, and T4 still gets the one armed attempt", async () => {
    const h = await sessionHarness();
    const control: { stopWaiterOnFailure: boolean; stoppedInJob: number } = { stopWaiterOnFailure: false, stoppedInJob: 0 };
    h.hooks.onReconnect = (reconnect) => {
      // Only T2's rotation: the first reaction registered on its reconnect promise.
      h.hooks.onReconnect = undefined;
      reconnect.then(undefined, () => {
        if (!control.stopWaiterOnFailure) return;
        control.stoppedInJob += 1;
        h.stop();
      });
    };
    const { s1, hold, p2 } = await rotatingT2WithHeldSession(h);
    h.stop();
    const t3 = h.beginTurn(T3);
    const continuation = vi.fn(() => "ran");
    const c3 = continueT(h, t3.handle, continuation);
    await flushMacrotasks();
    expect(h.coordinator.isCurrent(t3.handle.generation, T3, t3.controller)).toBe(true);
    control.stopWaiterOnFailure = true;
    const postsBeforeRelease = h.companion.sessionPosts;
    hold.release(503);
    const r2 = await p2;
    const r3 = await c3;
    await flushMacrotasks();
    const diag = `r2=${r2.status} r3=${JSON.stringify({ status: r3.status, sessionId: sessionOf(r3) })} stopped=${control.stoppedInJob} lifecycle=${String(h.lifecycle.sessionId)} ${diagnostics(h)}`;
    expect(control.stoppedInJob, diag).toBe(1);
    expect(r2.status, diag).toBe("stale");
    expect(r3.status, diag).toBe("stale");
    expect(continuation, diag).not.toHaveBeenCalled();
    expect(t3.controller.signal.aborted, diag).toBe(true);
    expect(h.companion.sessionPosts - postsBeforeRelease, diag).toBe(0);
    expect(h.lifecycle.sessionId, diag).toBeNull();
    expect(h.companion.deleted, diag).toEqual([s1]);
    const posts4 = h.companion.sessionPosts;
    const t4 = h.beginTurn(T4);
    const c4 = vi.fn(() => "ran");
    const r4 = await continueT(h, t4.handle, c4);
    const diag4 = `r4=${JSON.stringify({ status: r4.status, sessionId: sessionOf(r4) })} ${diagnostics(h)}`;
    expect(h.companion.sessionPosts - posts4, diag4).toBe(1);
    expect(r4.status, diag4).toBe("active");
    expect(sessionOf(r4), diag4).toBe(h.lifecycle.sessionId);
    expect(c4, diag4).toHaveBeenCalledTimes(1);
  });
});
