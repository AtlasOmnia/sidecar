import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import * as officeBridge from "../officeBridge";
import type { BridgePublicState, OfficeCompanionBridge } from "../officeBridge";
import { createExcelCompanionBridge } from "../officeBridgeExcel";
import { OfficeBridgeLifecycleController } from "../officeBridgeLifecycle";
import {
  createWorkbookPlanTurn,
  ServerOwnedTurnBindingStore,
  WORKBOOK_PLAN_TURN_TIMEOUT_MS,
  type WorkbookPlanConfirmationBinding,
  type WorkbookPlanVerificationAdapter,
} from "../workbookPlanOrchestrator";
import {
  BrowserIdempotencyStore,
  createTrustedOperationBoundary,
  MemoryIdempotencyStorage,
  type WorkbookPlanExecutionAdapter,
} from "../workbookPlanExecutor";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import { buildWorkbookPlanPolicyInput, type WorkbookPlanRouteRuntimeContext } from "../workbookPlanRoute";
import { classifyWorkbookRequest, type TrustedWorkbookRequest } from "../workbookPlanRequest";
import {
  TaskPaneTurnCoordinator,
  type TaskPaneTurnHandle,
  type TaskPaneTurnPreparation,
} from "../taskpane/taskPaneTurnCoordinator";
import * as routeModule from "../taskpane/serverOwnedPlanRoute";

/**
 * R3b (G1/D1a, F-A, D2 cap, R2d-1 A2). Real chain: createExcelCompanionBridge
 * -> OfficeBridgeLifecycleController -> TaskPaneTurnCoordinator ->
 * createServerOwnedPlanRoute -> orchestrator -> executor, against a
 * session-scoped companion double whose DELETE pops the session and cancels its
 * tickets (session.py SessionManager.delete) and whose invoke on an unknown or
 * deleted session answers the fixture C14 404. New coordinator/route APIs are
 * reached through casts so a missing API is an assertion failure.
 */

const DOC1 = "d0c00000-0000-4000-8000-000000000001";
const DOC2 = "d0c00000-0000-4000-8000-000000000002";
const PANE = "9a9e0000-0000-4000-8000-0000000000aa";
const T0 = "7a7a0000-0000-4000-8000-000000000000";
const T1 = "7a7a0000-0000-4000-8000-000000000001";
const T2 = "7a7a0000-0000-4000-8000-000000000002";
const T3 = "7a7a0000-0000-4000-8000-000000000003";
const OP0 = "5e7e0000-0000-4000-8000-000000000000";
const H1_OP = "5e7e0000-0000-4000-8000-000000000001";
const T2_OP = "5e7e0000-0000-4000-8000-000000000002";
const EXPLICIT_PROMPT = "write ok into A1";
/** Classifies as ambiguous_mutation: the policy requires pane confirmation. */
const AMBIGUOUS_PROMPT = "do the usual thing";
const NOT_FOUND_BODY = { ok: false, error: "Session not found", category: "NotFound" };
const CANCELLED_SUMMARY = "The workbook plan was cancelled before execution; nothing was executed.";
const REJECTED_SUMMARY = "The workbook plan was rejected before execution; nothing was executed.";
const FIXTURE = JSON.parse(
  readFileSync(resolve(__dirname, "fixtures/companion-wire/plan-wire-contract.v2.json"), "utf8"),
) as { constants: { claimLeaseSeconds: number; executeStartAckTimeoutMs: number } };
const TASKPANE_SOURCE = readFileSync(resolve(__dirname, "../taskpane/taskpane.ts"), "utf8");
const ROUTE_SOURCE = readFileSync(resolve(__dirname, "../taskpane/serverOwnedPlanRoute.ts"), "utf8");
const BRIDGE_SOURCE = readFileSync(resolve(__dirname, "../officeBridge.ts"), "utf8");

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
  /** Test copy of taskpane.ts updateBridgeState's turn/binding decision. */
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

/** Shared RB1/RB2/RB3/RB3m expectation: T2 rotated before any Hermes request, H1 cannot reach it. */
async function expectStaleHermesTurnCannotReachT2(h: SessionHarness, s1: string): Promise<void> {
  const t2 = h.beginTurn(T2);
  const p2 = await h.prepare(t2.handle);
  const invoke = h.companion.invoke(s1, H1_OP, EXECUTE_PLAN);
  const polled = await h.bridge().pollOnce().catch((error: unknown) => `rejected:${String(error)}`);
  const diag = `${diagnostics(h)} polled=${String(polled)} invoke=${invoke.status}`;
  expect(p2.status, diag).toBe("active");
  expect(sessionOf(p2), `T2 must not run on the tainted session S1; ${diag}`).not.toBe(s1);
  expect(sessionOf(p2), diag).toBe(h.lifecycle.sessionId);
  expect(h.companion.deleted, diag).toEqual([s1]);
  expect(invoke.status, diag).toBe(404);
  expect(invoke.body).toEqual(NOT_FOUND_BODY);
  expect(polled, diag).toBe(false);
  expect(h.companion.executeStarts, diag).toBe(0);
  expect(h.events, diag).not.toContain("mutation");
  expect(h.cells.size, diag).toBe(0);
  expect(h.coordinator.isCurrent(t2.handle.generation, T2, t2.controller)).toBe(true);
}

describe("R3b RED: an unclean pane turn end rotates the bridge session before the next turn", () => {
  it("RB1 stale-after-Stop: H1's later invoke hits the deleted session and never binds to T2", async () => {
    const h = await sessionHarness();
    const { s1 } = await runT1WithHermesRequest(h, () => h.stop());
    await expectStaleHermesTurnCannotReachT2(h, s1);
  });

  it("RB2 stale-after-error: a provider/network throw leaves the mark uncleaned and T2 rotates", async () => {
    const h = await sessionHarness();
    const { s1 } = await runT1WithHermesRequest(h, (handle) => {
      try {
        throw new TypeError("simulated provider network failure");
      } catch {
        h.finish(handle);
      }
    });
    await expectStaleHermesTurnCannotReachT2(h, s1);
  });

  it("RB3 stale-after-deadline: a 90 s TimeoutError end taints the session and T2 rotates", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    const h = await sessionHarness();
    const { s1 } = await runT1WithHermesRequest(h, (handle) => {
      vi.setSystemTime(Date.now() + WORKBOOK_PLAN_TURN_TIMEOUT_MS + 1);
      const timeout = new DOMException("Turn exceeded the 90-second time limit.", "TimeoutError");
      expect(timeout.name).toBe("TimeoutError");
      h.finish(handle);
    });
    await expectStaleHermesTurnCannotReachT2(h, s1);
  });

  it("RB3m direct-mode taint (OQ3): the mark is taken without X-Hermes-Session-Id use and T2 still rotates", async () => {
    const h = await sessionHarness();
    // Direct model mode never sends X-Hermes-Session-Id, but the pane still
    // passes the bridge session to beginHermesRequest (no model-mode branch; SP1).
    const { s1 } = await runT1WithHermesRequest(h, () => h.stop());
    await expectStaleHermesTurnCannotReachT2(h, s1);
  });

  it("RB4a F-A: a document-switch rotation inside prepareTurn keeps its own unprepared turn (no prompt drop)", async () => {
    const h = await sessionHarness();
    const s1 = h.lifecycle.sessionId!;
    const invalidateSession = vi.spyOn(h.bindings, "invalidateSession");
    const t0 = h.beginTurn(T0);
    expect((await h.prepare(t0.handle)).status).toBe("active");
    const old = h.coordinator.authorizeServerOperation(OP0);
    expect(old.ok).toBe(true);
    h.docs.current = DOC2;
    const t1 = h.beginTurn(T1);
    const p1 = await h.prepare(t1.handle);
    const diag = diagnostics(h);
    expect(p1.status, `F-A prompt drop: ${diag}`).toBe("active");
    expect(h.state.forcedStops, diag).toBe(0);
    expect(t1.controller.signal.aborted).toBe(false);
    expect(h.coordinator.isCurrent(t1.handle.generation, T1, t1.controller)).toBe(true);
    expect(sessionOf(p1)).not.toBe(s1);
    expect(p1.status === "active" ? p1.documentId : null).toBe(DOC2);
    expect(h.companion.deleted).toEqual([s1]);
    expect(invalidateSession).toHaveBeenCalledWith(s1);
    if (old.ok === true) expect(h.bindings.isServerTicketLive(old.identity, old.binding)).toBe(false);
  });

  it("RB4b F-A: a taint rotation inside prepareTurn keeps its own unprepared turn", async () => {
    const h = await sessionHarness();
    const invalidateSession = vi.spyOn(h.bindings, "invalidateSession");
    const { s1 } = await runT1WithHermesRequest(h, (handle) => h.finish(handle));
    const t2 = h.beginTurn(T2);
    const p2 = await h.prepare(t2.handle);
    const diag = diagnostics(h);
    expect(p2.status, diag).toBe("active");
    expect(sessionOf(p2), diag).not.toBe(s1);
    expect(h.state.forcedStops, diag).toBe(0);
    expect(t2.controller.signal.aborted).toBe(false);
    expect(h.coordinator.isCurrent(t2.handle.generation, T2, t2.controller)).toBe(true);
    expect(invalidateSession).toHaveBeenCalledWith(s1);
  });
});

describe("R3b RED: server-route confirmation deadline is capped at the claim lease less ack timeout and margin", () => {
  async function confirmationTicket(options: { runtimeAdvanceMs?: number } = {}) {
    vi.useFakeTimers({ toFake: ["Date", "setTimeout", "clearTimeout"] });
    vi.setSystemTime(new Date("2026-10-06T12:00:00.000Z"));
    const h = await sessionHarness();
    const s1 = h.lifecycle.sessionId!;
    const t1 = h.beginTurn(T1, AMBIGUOUS_PROMPT);
    const turnStart = Date.now();
    expect((await h.prepare(t1.handle)).status).toBe("active");
    vi.setSystemTime(Date.now() + 30_000);
    if (options.runtimeAdvanceMs !== undefined) {
      const advance = options.runtimeAdvanceMs;
      h.hooks.onRuntimeRead = () => { vi.setSystemTime(Date.now() + advance); };
    }
    const t0 = Date.now();
    const invoke = h.companion.invoke(s1, H1_OP, EXECUTE_PLAN);
    const polled = h.bridge().pollOnce();
    cleanups.push(async () => {
      h.coordinator.stopCurrentTurn();
      await polled.catch(() => undefined);
    });
    await flushUntil(() => h.awaitSpy.mock.calls.length > 0 || h.companion.results.length > 0);
    return { h, t1, turnStart, t0, ticketId: invoke.ticketId!, polled };
  }

  it("RB5 the awaited binding deadline is ticket receipt + 13 s and expiry posts executed:false cancelled before the lease ends", async () => {
    const { h, turnStart, t0, ticketId, polled } = await confirmationTicket();
    const awaited = h.awaitSpy.mock.calls[0]?.[0] as WorkbookPlanConfirmationBinding | undefined;
    expect(awaited?.deadline, `uncapped turn deadline would be ${turnStart + WORKBOOK_PLAN_TURN_TIMEOUT_MS}`).toBe(t0 + 13_000);
    expect(awaited!.deadline).toBeLessThan(turnStart + WORKBOOK_PLAN_TURN_TIMEOUT_MS);
    await vi.advanceTimersByTimeAsync(12_999);
    expect(h.companion.resultsFor(ticketId)).toHaveLength(0);
    await vi.advanceTimersByTimeAsync(1);
    await expect(polled).resolves.toBe(true);
    expect(Date.now()).toBeLessThan(t0 + FIXTURE.constants.claimLeaseSeconds * 1000);
    expect(h.companion.resultsFor(ticketId)).toEqual([{
      ticketId,
      executed: false,
      result: { ok: false, action: "workbook_plan", reason: "cancelled", summary: CANCELLED_SUMMARY },
    }]);
    expect(h.companion.executeStarts).toBe(0);
    expect(h.events).not.toContain("mutation");
    expect(h.cells.size).toBe(0);
  });

  it("RB5b no dead confirmation: once the cap has passed the pane confirmation dependency is never called", async () => {
    const { h, ticketId, polled } = await confirmationTicket({ runtimeAdvanceMs: 13_001 });
    expect(h.awaitSpy, "a dead confirmation must not reach the pane bar").not.toHaveBeenCalled();
    await expect(polled).resolves.toBe(true);
    expect(h.companion.resultsFor(ticketId)).toEqual([{
      ticketId,
      executed: false,
      result: { ok: false, action: "workbook_plan", reason: "cancelled", summary: CANCELLED_SUMMARY },
    }]);
    expect(h.companion.executeStarts).toBe(0);
    expect(h.events).not.toContain("mutation");
    expect(h.cells.size).toBe(0);
  });

  it("RB5c consume receives the same capped binding values as the awaited binding", async () => {
    const { h, t0, ticketId, polled } = await confirmationTicket();
    await vi.advanceTimersByTimeAsync(5_000);
    expect(h.coordinator.continuePendingWorkbookPlan()).toBe(true);
    await expect(polled).resolves.toBe(true);
    const awaited = h.awaitSpy.mock.calls[0]?.[0] as WorkbookPlanConfirmationBinding;
    const consumed = h.consumeSpy.mock.calls[0]?.[1] as WorkbookPlanConfirmationBinding | undefined;
    expect(consumed?.deadline).toBe(t0 + 13_000);
    expect({ ...consumed }).toEqual({ ...awaited });
    expect(h.companion.resultsFor(ticketId)[0]).toMatchObject({ executed: true });
  });

  it("RB6 constant pins: lease = fixture claimLeaseSeconds, margin 2 s, cap 13 s", () => {
    const mod = routeModule as unknown as Record<string, unknown>;
    expect(mod.COMPANION_CLAIM_LEASE_MS).toBe(FIXTURE.constants.claimLeaseSeconds * 1000);
    expect(mod.COMPANION_CLAIM_LEASE_MS).toBe(20_000);
    expect(mod.CONFIRMATION_CAP_MARGIN_MS).toBe(2_000);
    expect(FIXTURE.constants.executeStartAckTimeoutMs).toBe(officeBridge.EXECUTE_START_ACK_TIMEOUT_MS);
    expect(Number(mod.COMPANION_CLAIM_LEASE_MS) - officeBridge.EXECUTE_START_ACK_TIMEOUT_MS - Number(mod.CONFIRMATION_CAP_MARGIN_MS)).toBe(13_000);
    expect(typeof mod.withClaimLeaseConfirmationCap).toBe("function");
  });
});

describe("R3b RED: result re-post is bounded at 3 attempts per ticket (real chain)", () => {
  it.each([
    ["400", 400, "scripted result 400"],
    ["500", 500, "scripted result 500"],
    ["network reject", "network", "scripted result network reset"],
  ] as const)("RB7 %s: exactly 3 identical posts, each poll rethrows the original error, then the next ticket is served", async (_label, step, message) => {
    const h = await sessionHarness();
    const s1 = h.lifecycle.sessionId!;
    const t1 = h.beginTurn(T1);
    expect((await h.prepare(t1.handle)).status).toBe("active");
    const first = h.companion.invoke(s1, H1_OP, EXECUTE_PLAN);
    h.companion.resultScripts.set(first.ticketId!, Array.from({ length: 10 }, () => step));
    const second = h.companion.invokePrimitive(s1, "read_range", { sheet: "Sheet1", address: "A1" });
    const outcomes: string[] = [];
    for (let poll = 0; poll < 10; poll += 1) {
      outcomes.push(await h.bridge().pollOnce().then(
        (value) => `resolved:${String(value)}`,
        (error: unknown) => `rejected:${error instanceof Error ? error.message : String(error)}`,
      ));
    }
    const firstBodies = h.companion.resultsFor(first.ticketId);
    expect(firstBodies, `outcomes=${JSON.stringify(outcomes)}`).toHaveLength(3);
    expect(firstBodies[1]).toEqual(firstBodies[0]);
    expect(firstBodies[2]).toEqual(firstBodies[0]);
    expect(firstBodies[0]).toMatchObject({ executed: true });
    expect(outcomes.slice(0, 3)).toEqual([`rejected:${message}`, `rejected:${message}`, `rejected:${message}`]);
    expect(outcomes[3]).toBe("resolved:true");
    expect(outcomes.slice(4)).toEqual(Array.from({ length: 6 }, () => "resolved:false"));
    expect(h.companion.resultsFor(second.ticketId)).toEqual([{ ticketId: second.ticketId, result: { ok: true, data: [[1]] } }]);
    expect(h.companion.tickets.get(second.ticketId!)?.state).toBe("completed");
    expect(h.state.planExecuteToolCalls).toBe(1);
    expect(h.companion.executeStarts).toBe(1);
    expect(h.events.filter((event) => event === "mutation")).toHaveLength(1);
  });
});

describe("R3b GREEN pins", () => {
  it("GB1 clean-no-churn: a cleaned Hermes request does not rotate the session", async () => {
    const h = await sessionHarness();
    const { s1 } = await runT1WithHermesRequest(h, (handle) => h.finish(handle), true);
    const posts = h.companion.sessionPosts;
    const t2 = h.beginTurn(T2);
    const p2 = await h.prepare(t2.handle);
    expect(p2.status).toBe("active");
    expect(sessionOf(p2)).toBe(s1);
    expect(h.companion.sessionPosts).toBe(posts);
    expect(h.companion.deleted).toEqual([]);
  });

  it("GB2 rotation failure fails closed before any Hermes request; taint persists; a later healthy session goes active", async () => {
    const h = await sessionHarness();
    const { s1 } = await runT1WithHermesRequest(h, () => h.stop());
    const marksBefore = r3b(h.coordinator).hermesRequestMarkCount;
    h.companion.sessionFailure = 503;
    const t2 = h.beginTurn(T2);
    const continuation = vi.fn();
    const gated = await h.coordinator.continueTurn(t2.handle, continuation);
    expect(gated).toMatchObject({ status: "rejected", reason: "bridge-reconnect-failed" });
    expect(continuation).not.toHaveBeenCalled();
    expect(h.events).not.toContain("mutation");
    expect(h.cells.size).toBe(0);
    expect(marksBefore).toBe(1);
    expect(r3b(h.coordinator).hermesRequestMarkCount).toBe(1);
    expect(h.companion.deleted).toEqual([s1]);
    h.finish(t2.handle);
    // Recovery: the companion is healthy again and the lifecycle restarts
    // (the failed rotation already stopped and DELETEd S1).
    h.companion.sessionFailure = null;
    await h.lifecycle.start();
    const t3 = h.beginTurn(T3);
    const p3 = await h.prepare(t3.handle);
    expect(p3.status).toBe("active");
    expect(sessionOf(p3)).not.toBe(s1);
    expect(sessionOf(p3)).toBe(h.lifecycle.sessionId);
    expect(h.companion.deleted).toEqual([s1]);
  });

  it("GB3 a stale ticket claimed before the rotation (T2 begun, not prepared) is rejected with zero mutation", async () => {
    const h = await sessionHarness();
    const { s1 } = await runT1WithHermesRequest(h, () => h.stop());
    const t2 = h.beginTurn(T2);
    const invoke = h.companion.invoke(s1, H1_OP, EXECUTE_PLAN);
    await expect(h.bridge().pollOnce()).resolves.toBe(true);
    expect(h.companion.resultsFor(invoke.ticketId)).toEqual([{
      ticketId: invoke.ticketId,
      executed: false,
      result: { ok: false, action: "workbook_plan", reason: "rejected", summary: REJECTED_SUMMARY },
    }]);
    expect(h.companion.executeStarts).toBe(0);
    expect(h.cells.size).toBe(0);
    const p2 = await h.prepare(t2.handle);
    expect(p2.status).toBe("active");
    expect(sessionOf(p2)).not.toBe(s1);
  });

  it("GB4 after rotation T2's own identical plan on the new session commits once", async () => {
    const h = await sessionHarness();
    const { s1 } = await runT1WithHermesRequest(h, () => h.stop());
    const t2 = h.beginTurn(T2);
    const p2 = await h.prepare(t2.handle);
    const s2 = sessionOf(p2)!;
    expect(s2).not.toBe(s1);
    expect(h.companion.invoke(s1, H1_OP, EXECUTE_PLAN).status).toBe(404);
    const own = h.companion.invoke(s2, T2_OP, EXECUTE_PLAN);
    await expect(h.bridge().pollOnce()).resolves.toBe(true);
    expect(h.companion.executeStarts).toBe(1);
    expect(h.events.filter((event) => event === "mutation")).toHaveLength(1);
    expect(h.cells.get("Sheet1!A1")).toEqual([["ok"]]);
    expect(h.companion.resultsFor(own.ticketId)[0]).toMatchObject({ executed: true, result: { ok: true } });
  });

  it("GB5 a Stop before prepareTurn with no Hermes request does not churn the session", async () => {
    const h = await sessionHarness();
    const s1 = h.lifecycle.sessionId!;
    h.beginTurn(T1);
    h.stop();
    const t2 = h.beginTurn(T2);
    const p2 = await h.prepare(t2.handle);
    expect(sessionOf(p2)).toBe(s1);
    expect(h.companion.deleted).toEqual([]);
    expect(h.companion.sessionPosts).toBe(1);
  });

  it("GB6 marks are bounded at 64, stay tainted when full, and one rotation clears them", async () => {
    const h = await sessionHarness();
    const s1 = h.lifecycle.sessionId!;
    const api = r3b(h.coordinator);
    const requests = Array.from({ length: 65 }, () => api.beginHermesRequest?.(s1));
    expect(api.hermesRequestMarkCount).toBe(64);
    requests[0]?.endClean();
    requests[0]?.endClean();
    expect(api.hermesRequestMarkCount).toBe(63);
    const t2 = h.beginTurn(T2);
    const p2 = await h.prepare(t2.handle);
    expect(sessionOf(p2)).not.toBe(s1);
    expect(api.hermesRequestMarkCount).toBe(0);
    for (const request of requests) request?.endClean();
    expect(api.hermesRequestMarkCount).toBe(0);
    h.finish(t2.handle);
    const t3 = h.beginTurn(T3);
    const p3 = await h.prepare(t3.handle);
    expect(sessionOf(p3)).toBe(sessionOf(p2));
    expect(h.companion.deleted).toEqual([s1]);
  });

  it("GB6b a full mark table with every listed mark cleaned still keeps the overflowed session tainted", async () => {
    const h = await sessionHarness();
    const s1 = h.lifecycle.sessionId!;
    const api = r3b(h.coordinator);
    const requests = Array.from({ length: 65 }, () => api.beginHermesRequest?.(s1));
    for (const request of requests) request?.endClean();
    expect(api.hermesRequestMarkCount).toBe(0);
    const t2 = h.beginTurn(T2);
    const p2 = await h.prepare(t2.handle);
    expect(sessionOf(p2)).not.toBe(s1);
    expect(h.companion.deleted).toEqual([s1]);
  });

  it("GB7 a confirmation accepted at t0 + 10 s commits once; the direct route keeps the 90 s turn deadline", async () => {
    vi.useFakeTimers({ toFake: ["Date", "setTimeout", "clearTimeout"] });
    vi.setSystemTime(new Date("2026-10-06T12:00:00.000Z"));
    const h = await sessionHarness();
    const s1 = h.lifecycle.sessionId!;
    const t1 = h.beginTurn(T1, AMBIGUOUS_PROMPT);
    expect((await h.prepare(t1.handle)).status).toBe("active");
    const invoke = h.companion.invoke(s1, H1_OP, EXECUTE_PLAN);
    const polled = h.bridge().pollOnce();
    cleanups.push(async () => { h.coordinator.stopCurrentTurn(); await polled.catch(() => undefined); });
    await flushUntil(() => h.awaitSpy.mock.calls.length > 0);
    await vi.advanceTimersByTimeAsync(10_000);
    expect(h.coordinator.continuePendingWorkbookPlan()).toBe(true);
    await expect(polled).resolves.toBe(true);
    expect(h.companion.executeStarts).toBe(1);
    expect(h.events.filter((event) => event === "mutation")).toHaveLength(1);
    expect(h.companion.resultsFor(invoke.ticketId)[0]).toMatchObject({ executed: true, result: { ok: true } });

    // Direct route: unchanged binding deadline = turnStart + 90 s.
    const directTurnStart = Date.now();
    const directBindings: WorkbookPlanConfirmationBinding[] = [];
    const identity = { sessionId: PANE, documentId: DOC1, turnId: T3 };
    await createWorkbookPlanTurn({
      identity,
      trustedOperationBoundary: createTrustedOperationBoundary(identity),
      compileContext: COMPILE_CONTEXT,
      policyInput: (plan, confirmation) => buildWorkbookPlanPolicyInput(plan, RUNTIME, {
        ...classifyWorkbookRequest(AMBIGUOUS_PROMPT),
        ...(confirmation === undefined ? {} : { confirmation }),
      }),
      resourcePreflightAfterRuntime: () => ({ ok: true as const }),
      store: new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      model: async () => ({ content: "", toolCalls: [] }),
      initialResponse: {
        content: "",
        toolCalls: [{ id: "direct", type: "function", function: { name: "excel_execute_plan", arguments: JSON.stringify(EXECUTE_PLAN) } }],
      },
      inspect: async () => "{}",
      adapter: statefulAdapter([]).adapter,
      rolloutEnabled: true,
      route: "direct",
      trustedTurnOperationId: T3,
      turnStart: directTurnStart,
      turnGeneration: 1,
      awaitPaneConfirmation: async (binding) => {
        directBindings.push(binding);
        throw new DOMException("stop", "AbortError");
      },
      consumePaneConfirmation: () => undefined,
    });
    expect(directBindings).toHaveLength(1);
    expect(directBindings[0].deadline).toBe(directTurnStart + WORKBOOK_PLAN_TURN_TIMEOUT_MS);
    expect(TASKPANE_SOURCE).toContain("awaitPaneConfirmation: waitForPendingPlanConfirmation,\n    consumePaneConfirmation: (candidate, binding) => turnCoordinator.consumeWorkbookPlanConfirmation(candidate, binding),");
  });

  it("GB8 F-A boundary: a 401 auto-reconnect during an active prepared turn still force-stops it", async () => {
    const h = await sessionHarness();
    const s1 = h.lifecycle.sessionId!;
    const t1 = h.beginTurn(T1);
    expect((await h.prepare(t1.handle)).status).toBe("active");
    const record = h.companion.byId.get(s1)!;
    h.companion.byKey.delete(record.sessionKey);
    const heartbeat = h.intervals[h.intervals.length - 2];
    heartbeat();
    await flushUntil(() => h.lifecycle.sessionId !== null && h.lifecycle.sessionId !== s1);
    expect(h.state.forcedStops).toBeGreaterThanOrEqual(1);
    expect(t1.controller.signal.aborted).toBe(true);
    expect(h.coordinator.isCurrent(t1.handle.generation, T1, t1.controller)).toBe(false);
  });

  it("GB8b the rotation-window exception is scoped to this coordinator's own unprepared turn", async () => {
    const h = await sessionHarness();
    const api = r3b(h.coordinator);
    expect(typeof api.bridgeStateInvalidatesActiveTurn).toBe("function");
    const s1 = h.lifecycle.sessionId!;
    const previous = { sessionId: s1, documentId: DOC1 };
    const disconnected: BridgePublicState = { connected: false, sessionId: null, documentId: null };
    expect(api.bridgeStateInvalidatesActiveTurn?.(disconnected, null, previous)).toBe(false);
    expect(api.bridgeStateInvalidatesActiveTurn?.(disconnected, { turnId: T1, sessionId: s1, documentId: null }, previous)).toBe(true);
    expect(api.bridgeStateInvalidatesActiveTurn?.(
      { connected: true, sessionId: s1, documentId: DOC1 }, { turnId: T1, sessionId: s1, documentId: DOC1 }, previous,
    )).toBe(false);
    expect(api.bridgeStateInvalidatesActiveTurn?.(
      { connected: true, sessionId: s1, documentId: DOC2 }, { turnId: T1, sessionId: s1, documentId: null }, previous,
    )).toBe(true);
  });
});

describe("R3b SP1 source pins (M3 style)", () => {
  const continuationStart = TASKPANE_SOURCE.indexOf("const runTurnContinuation = async");
  const continuationEnd = TASKPANE_SOURCE.indexOf("if (settings.workbookPlanRolloutEnabled === true) {", continuationStart);
  const continuation = TASKPANE_SOURCE.slice(continuationStart, continuationEnd);
  const bridgeStateStart = TASKPANE_SOURCE.indexOf("function updateBridgeState(state: BridgePublicState): void {");
  const bridgeState = TASKPANE_SOURCE.slice(bridgeStateStart, TASKPANE_SOURCE.indexOf("\n}\n", bridgeStateStart));

  it("taskpane.ts marks every Hermes request before agent.run with one captured session id, in every model mode", () => {
    expect(continuationStart).toBeGreaterThan(0);
    expect(continuationEnd).toBeGreaterThan(continuationStart);
    const begin = continuation.indexOf("const hermesRequest = turnCoordinator.beginHermesRequest(hermesSessionId);");
    const run = continuation.indexOf("agent.run(");
    expect(continuation).toContain("const hermesSessionId = bridgeLifecycle.sessionId;");
    expect(begin).toBeGreaterThan(0);
    expect(run).toBeGreaterThan(begin);
    const runCall = continuation.slice(run, continuation.indexOf("// Map stopReason", run));
    expect(runCall).toMatch(/turnSignal,\n\s+hermesSessionId,\n\s+\);/);
    expect(runCall).not.toContain("bridgeLifecycle.sessionId");
    const beforeBegin = continuation.slice(continuation.indexOf("const trace = startTrace"), begin);
    expect(beforeBegin).not.toMatch(/\bif\s*\(/);
    expect(continuation.slice(Math.max(0, begin - 400), begin)).not.toMatch(/settings\.model|hermes-agent|isHermesAgent/);
  });

  it("taskpane.ts proves a clean end only on the resolved path after agent.run", () => {
    const run = continuation.indexOf("agent.run(");
    const catchIndex = continuation.indexOf("} catch (error) {", run);
    const finallyIndex = continuation.indexOf("} finally {", run);
    expect(continuation.split(".endClean()")).toHaveLength(2);
    const endClean = continuation.indexOf("hermesRequest.endClean();");
    expect(endClean).toBeGreaterThan(continuation.indexOf("// Map stopReason", run));
    expect(endClean).toBeLessThan(catchIndex);
    expect(catchIndex).toBeLessThan(finallyIndex);
  });

  it("taskpane.ts updateBridgeState delegates the force-stop decision and still invalidates bindings unconditionally", () => {
    expect(bridgeState).toContain("turnCoordinator.bridgeStateInvalidatesActiveTurn(");
    expect(bridgeState).toContain("if (activeTurnInvalidated) stopRunningAgent(true);");
    expect(bridgeState).toMatch(/\n  if \(lastBridgeSessionId !== null &&\n      \(lastBridgeSessionId !== state\.sessionId \|\| lastBridgeDocumentId !== state\.documentId\)\) \{\n    serverOwnedTurnBindings\.invalidateSession\(lastBridgeSessionId\);\n  \}/);
    expect(bridgeState).toMatch(/\n  if \(lastBridgeDocumentId !== null && lastBridgeDocumentId !== state\.documentId\) \{\n    serverOwnedTurnBindings\.invalidateDocument\(lastBridgeDocumentId\);\n  \}/);
  });

  it("serverOwnedPlanRoute.ts applies the cap through the block-scoped shadow and keeps the task11-d2 literals", () => {
    const shadow = ROUTE_SOURCE.indexOf("const dependencies = cappedDependencies;");
    const call = ROUTE_SOURCE.indexOf("result = await createWorkbookPlanTurn({");
    expect(shadow).toBeGreaterThan(0);
    expect(call).toBeGreaterThan(shadow);
    expect(ROUTE_SOURCE.slice(shadow, call)).toMatch(/^const dependencies = cappedDependencies;\n\s+$/);
    // Ticket receipt time is read synchronously (no await since the route was
    // entered) and only after the rollout gate and the bridge-session check:
    // the frozen wire leg (C21-R) pins that rollout-off and no-session tickets
    // touch no other dependency.
    const entry = ROUTE_SOURCE.indexOf("  async function executeServerOwnedWorkbookPlan(");
    const gate = ROUTE_SOURCE.indexOf("if (dependencies.isRolloutEnabled() !== true) return notExecuted(\"not_started\");", entry);
    const sessionCheck = ROUTE_SOURCE.indexOf("if (sessionId === null) return notExecuted(\"rejected\");", entry);
    const observed = ROUTE_SOURCE.indexOf("const claimObservedAt = (dependencies.now ?? Date.now)();", entry);
    expect(entry).toBeGreaterThan(0);
    expect(observed).toBeGreaterThan(sessionCheck);
    expect(sessionCheck).toBeGreaterThan(gate);
    expect(observed).toBeLessThan(ROUTE_SOURCE.indexOf("dependencies.coordinator.authorizeAndRun(", entry));
    expect(ROUTE_SOURCE.slice(entry, observed)).not.toMatch(/\bawait\b/);
    expect(ROUTE_SOURCE).toContain("awaitPaneConfirmation: dependencies.awaitPaneConfirmation");
    expect(ROUTE_SOURCE).toContain("consumePaneConfirmation: (candidate, binding) => dependencies.coordinator.consumeWorkbookPlanConfirmation(candidate, binding)");
    expect(ROUTE_SOURCE).toContain("cancelPaneConfirmation: () => dependencies.coordinator.cancelPendingConfirmation()");
  });

  it("officeBridge.ts declares the A2 bound and the fixed oversize message", () => {
    expect(BRIDGE_SOURCE).toContain("const MAX_RESULT_POST_ATTEMPTS = 3;");
    expect(BRIDGE_SOURCE).toContain("const OVERSIZE_PRIMITIVE_RESULT_MESSAGE = \"Office tool result was too large to submit.\";");
  });
});
