import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import type { BridgePublicState, OfficeCompanionBridge } from "../officeBridge";
import { createExcelCompanionBridge } from "../officeBridgeExcel";
import { OfficeBridgeLifecycleController } from "../officeBridgeLifecycle";
import {
  ServerOwnedTurnBindingStore,
  type WorkbookPlanVerificationAdapter,
} from "../workbookPlanOrchestrator";
import {
  BrowserIdempotencyStore,
  DeterministicIdempotencyCoordinator,
  type IdempotencyStorage,
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
 * R3d-T (TS half, TX rows): cross-layer restart/replay matrix. Real chain per
 * pane instance: createExcelCompanionBridge -> OfficeBridgeLifecycleController
 * -> TaskPaneTurnCoordinator + ServerOwnedTurnBindingStore ->
 * createServerOwnedPlanRoute -> orchestrator -> executor -> a NEW
 * BrowserIdempotencyStore per ticket (taskpane.ts createStore shape) over ONE
 * shared storage and ONE shared DeterministicIdempotencyCoordinator. A "pane
 * instance" is a fresh bridge + lifecycle + coordinator + binding store +
 * route (+ a fresh store runtimeOwner per ticket): how a reload looks to every
 * layer. The companion double answers only with response bodies read from the
 * frozen wire fixture (sha256 pinned below); its delete semantics mirror
 * scripts/companion/session.py Ticket.cancel. Rows marked [synthetic] use a
 * redelivery the real companion never performs (M7, no requeue) as
 * defence-in-depth input only. "Mutation" = adapter mutation events plus the
 * double's accepted execute-starts. Harness adapted by copy from
 * task-pane-session-rotation-r3b.test.ts and workbook-plan-production-route.test.ts;
 * nothing is imported from other test files.
 */

type Json = null | boolean | number | string | Json[] | { [key: string]: Json };

interface FixtureStep {
  readonly kind: string;
  readonly method?: string;
  readonly path?: string;
  readonly request?: { readonly arguments?: Json };
  readonly response?: { readonly status: number; readonly body?: Json };
}

interface FixtureCase {
  readonly id: string;
  readonly steps: readonly FixtureStep[];
}

const FIXTURE_SHA256 = "4b27a29446334173ccae49911ebf2a29f04ebd69748db18af975ed715ae8c14d";
const FIXTURE_BYTES = readFileSync(resolve(__dirname, "fixtures/companion-wire/plan-wire-contract.v2.json"));
const FIXTURE = JSON.parse(FIXTURE_BYTES.toString("utf8")) as { readonly cases: readonly FixtureCase[] };

function fixtureStep(
  caseId: string,
  index: number,
  kind: string,
  method: string | undefined,
  path: string | undefined,
  status: number | null,
): FixtureStep {
  const found = FIXTURE.cases.find((candidate) => candidate.id === caseId);
  if (found === undefined) throw new Error(`fixture case ${caseId} is missing`);
  const step = found.steps[index];
  if (step === undefined || step.kind !== kind || step.method !== method || step.path !== path ||
      (status === null ? step.response !== undefined : step.response?.status !== status)) {
    throw new Error(`fixture step ${caseId}#${index} is not ${kind} ${String(method)} ${String(path)} ${status}`);
  }
  return step;
}

function fill(value: Json | undefined, vars: Record<string, Json>): Json {
  if (value === undefined) throw new Error("fixture body is missing");
  if (typeof value === "string" && value.startsWith("$")) {
    const name = value.slice(1);
    if (!(name in vars)) throw new Error(`unbound fixture variable ${value}`);
    return vars[name];
  }
  if (Array.isArray(value)) return value.map((entry) => fill(entry, vars));
  if (value !== null && typeof value === "object") {
    return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, fill(entry, vars)]));
  }
  return value;
}

/** §4.3 fidelity map: every non-trivial double response is read from these fixture steps. */
const FX = {
  session201: fixtureStep("C02-completed", 0, "pane", "POST", "/excel/session", 201),
  planRequest: fixtureStep("C02-completed", 1, "plugin_invoke_begin", undefined, undefined, null),
  pending200: fixtureStep("C02-completed", 2, "pane", "GET", "/excel/pending", 200),
  executeStart200: fixtureStep("C02-completed", 3, "pane", "POST", "/excel/execute-start", 200),
  result200: fixtureStep("C02-completed", 4, "pane", "POST", "/excel/result", 200),
  executeStart409: fixtureStep("C09-cancel-claimed", 5, "pane", "POST", "/excel/execute-start", 409),
  result409: fixtureStep("C09-cancel-claimed", 6, "pane", "POST", "/excel/result", 409),
  leaseExecuteStart409: fixtureStep("C11-claimed-lease-expiry", 5, "pane", "POST", "/excel/execute-start", 409),
  unauthorized401: fixtureStep("C14-4xx-set/pane-auth", 2, "pane", "GET", "/excel/pending", 401),
  invoke404: fixtureStep("C14-4xx-set/invoke-unknown-session", 1, "plugin_invoke_end", undefined, undefined, 404),
};

function stepRequestArguments(step: FixtureStep): Record<string, unknown> {
  const value = step.request?.arguments;
  if (value === undefined || value === null || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("fixture plan arguments are missing");
  }
  return value as Record<string, unknown>;
}

/** The fixture's own execute plan (C02 plugin_invoke_begin arguments). */
const PLAN = stepRequestArguments(FX.planRequest);
/** Same operations, changed summary: a different canonical digest for the same op. */
const CHANGED_PLAN = { ...PLAN, summary: "Write one cell (R3d changed digest)" };
/** The route's non-executing not_started payload; also the result body the fixture C09 pane posts. */
const A5_PAYLOAD = {
  ok: false,
  action: "workbook_plan",
  reason: "not_started",
  summary: "Workbook plan execution did not start; nothing was executed.",
};
const REPLAYED = "Trusted server operation was replayed.";
const CAPACITY = "Trusted server turn capacity is unavailable.";
const STORAGE_KEY = "hermes.excel.workbookPlan.idempotency.v1";
const H = 10 * 60 * 1000;
const DAY = 24 * 60 * 60 * 1000;
const T_BASE = Date.UTC(2026, 9, 6, 12, 0, 0);

const DOC = "d0c00000-0000-4000-8000-0000000000d3";
const PANE = "9a9e0000-0000-4000-8000-0000000000d3";
const EXPLICIT_PROMPT = "write ok into A1";

function uuid(prefix: string, n: number): string {
  return `${prefix}-0000-4000-8000-${String(n).padStart(12, "0")}`;
}
const turn = (n: number): string => uuid("7a7a0000", n);
const op = (n: number): string => uuid("5e7e0000", n);
const X = op(1);
const Y = op(2);

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

function json(status: number, body?: unknown): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
  });
}

type TicketState =
  | "pending" | "claimed" | "executing" | "completed" | "not_executed"
  | "cancelled_unclaimed" | "cancelled_claimed" | "executing_uncertain";

interface DoubleTicket {
  readonly ticketId: string;
  readonly op: string;
  readonly sessionId: string;
  readonly args: Record<string, unknown>;
  readonly maxDeliveries: number;
  deliveries: number;
  state: TicketState;
}

interface DoubleSession {
  readonly sessionId: string;
  readonly sessionKey: string;
  readonly documentId: string;
  readonly tickets: DoubleTicket[];
}

/** Session-scoped companion double; every 2xx/4xx body comes from the fixture (§4.3). */
class R3dCompanion {
  readonly byKey = new Map<string, DoubleSession>();
  readonly byId = new Map<string, DoubleSession>();
  readonly tickets = new Map<string, DoubleTicket>();
  readonly executeStarts: { ticketId: string; op: string; sessionId: string; status: number }[] = [];
  readonly results: { ticketId: string; op: string; status: number; body: Record<string, unknown> }[] = [];
  readonly deleted: string[] = [];
  unauthorized = 0;
  sessionPosts = 0;
  restarts = 0;
  private adopt: DoubleSession | null = null;
  private sessionCounter = 0;
  private ticketCounter = 0;

  /** [synthetic] The next pane registration re-attaches to an existing session (same id and key). */
  adoptNext(sessionId: string): void {
    const session = this.byId.get(sessionId);
    if (session === undefined) throw new Error("cannot adopt an unknown session");
    this.adopt = session;
  }

  /** Companion process restart: every session (and so every ticket) is gone (D3, session nonces). */
  restart(): void {
    this.byKey.clear();
    this.byId.clear();
    this.restarts += 1;
  }

  /** Plugin-side invoke; unknown session -> fixture C14 invoke-unknown-session 404. */
  invoke(sessionId: string, operationId: string, plan: Record<string, unknown>, maxDeliveries = 1):
    { status: number; body: Json; ticketId?: string } {
    const session = this.byId.get(sessionId);
    if (session === undefined) return { status: 404, body: fill(FX.invoke404.response!.body, {}) };
    this.ticketCounter += 1;
    const ticket: DoubleTicket = {
      ticketId: uuid("71c70000", this.ticketCounter), op: operationId, sessionId, args: { ...plan },
      maxDeliveries, deliveries: 0, state: "pending",
    };
    session.tickets.push(ticket);
    this.tickets.set(ticket.ticketId, ticket);
    return { status: 202, body: { ok: true }, ticketId: ticket.ticketId };
  }

  postedFor(ticketId: string): { status: number; body: Record<string, unknown> }[] {
    return this.results.filter((entry) => entry.ticketId === ticketId).map(({ status, body }) => ({ status, body }));
  }

  acceptedExecuteStarts(operationId?: string): number {
    return this.executeStarts.filter((entry) => entry.status === 200 && (operationId === undefined || entry.op === operationId)).length;
  }

  executeStartsFor(ticketId: string): number {
    return this.executeStarts.filter((entry) => entry.ticketId === ticketId).length;
  }

  readonly fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const url = String(input);
    const route = url.replace(/^.*\/excel\//, "");
    const method = init?.method ?? "GET";
    const headers = (init?.headers ?? {}) as Record<string, string>;
    const body = typeof init?.body === "string" && init.body ? JSON.parse(init.body) as Record<string, unknown> : {};
    if (method === "POST" && route === "session") {
      this.sessionPosts += 1;
      let session = this.adopt;
      this.adopt = null;
      if (session === null) {
        this.sessionCounter += 1;
        session = {
          sessionId: uuid("aaaaaaaa", this.sessionCounter).replace(/^aaaaaaaa-0000-4000-8000/, "aaaaaaaa-aaaa-4aaa-8aaa"),
          sessionKey: this.sessionCounter.toString(16).padStart(64, "0"),
          documentId: String(body.documentId),
          tickets: [],
        };
        this.byKey.set(session.sessionKey, session);
        this.byId.set(session.sessionId, session);
      }
      return json(201, fill(FX.session201.response!.body, {
        sessionId: session.sessionId, sessionKey: session.sessionKey, fingerprint: WORKBOOK_PLAN_FINGERPRINT,
      }));
    }
    const session = this.byKey.get(headers["X-Hermes-Session-Key"] ?? "");
    if (session === undefined || headers["X-Hermes-Session-Id"] !== session.sessionId) {
      this.unauthorized += 1;
      return json(401, fill(FX.unauthorized401.response!.body, {}));
    }
    if (method === "DELETE" && route === "session") {
      this.byKey.delete(session.sessionKey);
      this.byId.delete(session.sessionId);
      // session.py Ticket.cancel: pending -> cancelled_unclaimed, claimed -> cancelled_claimed,
      // executing -> executing_uncertain (never "undone").
      for (const ticket of session.tickets) {
        if (ticket.state === "pending") ticket.state = "cancelled_unclaimed";
        else if (ticket.state === "claimed") ticket.state = "cancelled_claimed";
        else if (ticket.state === "executing") ticket.state = "executing_uncertain";
      }
      this.deleted.push(session.sessionId);
      return json(204);
    }
    if (method === "POST" && route === "heartbeat") return json(200, { ok: true });
    if (method === "GET" && route === "pending") {
      const ticket = session.tickets.find((candidate) => candidate.state === "pending") ??
        // [synthetic] duplicate delivery of an already handed-out ticket.
        session.tickets.find((candidate) => candidate.deliveries > 0 && candidate.deliveries < candidate.maxDeliveries);
      if (ticket === undefined) return json(204);
      if (ticket.state === "pending") ticket.state = "claimed";
      ticket.deliveries += 1;
      const envelope = fill(FX.pending200.response!.body, {
        ticketId: ticket.ticketId, turnOperationId: ticket.op, fingerprint: WORKBOOK_PLAN_FINGERPRINT,
        turnKey: "4b4b0000-0000-4000-8000-0000000000a3",
      }) as Record<string, Json>;
      envelope.arguments = ticket.args as Json;
      return json(200, envelope);
    }
    if (method === "POST" && route === "execute-start") {
      const ticket = session.tickets.find((candidate) => candidate.ticketId === body.ticketId);
      const accepted = ticket !== undefined && ticket.state === "claimed";
      this.executeStarts.push({
        ticketId: String(body.ticketId), op: ticket?.op ?? "", sessionId: session.sessionId, status: accepted ? 200 : 409,
      });
      if (!accepted) return json(409, fill(FX.executeStart409.response!.body, {}));
      ticket.state = "executing";
      return json(200, fill(FX.executeStart200.response!.body, { ticketId: ticket.ticketId }));
    }
    if (method === "POST" && route === "result") {
      const ticket = session.tickets.find((candidate) => candidate.ticketId === body.ticketId);
      const ok = ticket !== undefined && typeof body.executed === "boolean" &&
        ((body.executed === true && ticket.state === "executing") || (body.executed === false && ticket.state === "claimed"));
      this.results.push({ ticketId: String(body.ticketId), op: ticket?.op ?? "", status: ok ? 200 : 409, body });
      if (!ok) return json(409, fill(FX.result409.response!.body, {}));
      ticket.state = body.executed === true ? "completed" : "not_executed";
      return json(200, fill(FX.result200.response!.body, { ticketId: ticket.ticketId }));
    }
    return json(404, { error: "route" });
  };
}

/** Shared browser storage with scripted faults (TX9). */
class ScriptedStorage implements IdempotencyStorage {
  readonly values = new Map<string, string>();
  getThrows = false;
  setFault: ((value: string) => boolean) | null = null;

  getItem(key: string): string | null {
    if (this.getThrows) throw new Error("storage read unavailable");
    return this.values.get(key) ?? null;
  }

  setItem(key: string, value: string): void {
    if (this.setFault?.(value) === true) throw new DOMException("The quota has been exceeded.", "QuotaExceededError");
    this.values.set(key, value);
  }

  removeItem(key: string): void {
    this.values.delete(key);
  }

  raw(): string | null {
    return this.values.get(STORAGE_KEY) ?? null;
  }

  records(): Record<string, unknown>[] {
    const raw = this.raw();
    return raw === null ? [] : JSON.parse(raw) as Record<string, unknown>[];
  }
}

interface Hold {
  readonly reached: () => boolean;
  readonly wait: () => Promise<void>;
  readonly release: () => void;
}

function hold(): Hold {
  let reached = false;
  let release!: () => void;
  const gate = new Promise<void>((done) => { release = done; });
  return { reached: () => reached, wait: async () => { reached = true; await gate; }, release: () => release() };
}

interface World {
  readonly companion: R3dCompanion;
  readonly storage: ScriptedStorage;
  readonly locks: DeterministicIdempotencyCoordinator;
  readonly cells: Map<string, unknown>;
  readonly mutations: string[];
  readonly claims: { pane: string; kind: string; report?: unknown }[];
  readonly finishes: { pane: string; ok: boolean; report: unknown }[];
  readonly holds: { beforeGate?: { pane: string; hold: Hold }; duringMutation?: { pane: string; hold: Hold }; beforeClaim?: { pane: string; hold: Hold } };
  storeOptions: { maxEntries?: number };
}

function world(): World {
  return {
    companion: new R3dCompanion(),
    storage: new ScriptedStorage(),
    locks: new DeterministicIdempotencyCoordinator(),
    cells: new Map(),
    mutations: [],
    claims: [],
    finishes: [],
    holds: {},
    storeOptions: {},
  };
}

/** taskpane.ts createStore shape: a new store object (new runtimeOwner) per ticket over the shared storage. */
function newStore(w: World, pane: string): BrowserIdempotencyStore {
  const store = new BrowserIdempotencyStore(w.storage, { clock: () => Date.now(), coordinator: w.locks, ...w.storeOptions });
  const claim = store.claim.bind(store);
  store.claim = async (...args: Parameters<typeof store.claim>) => {
    if (w.holds.beforeClaim?.pane === pane) await w.holds.beforeClaim.hold.wait();
    const result = await claim(...args);
    w.claims.push({ pane, kind: result.kind, ...("report" in result ? { report: result.report } : {}) });
    return result;
  };
  const finish = store.finish.bind(store);
  store.finish = async (...args: Parameters<typeof store.finish>) => {
    const ok = await finish(...args);
    w.finishes.push({ pane, ok, report: JSON.parse(JSON.stringify(args[3])) });
    return ok;
  };
  return store;
}

function adapterFor(w: World, pane: string): WorkbookPlanExecutionAdapter & WorkbookPlanVerificationAdapter {
  const write = (operations: readonly { readonly args: Record<string, unknown> }[]): void => {
    for (const operation of operations) {
      w.cells.set(`${String(operation.args.sheet)}!${String(operation.args.address)}`, operation.args.values);
    }
  };
  return {
    async executeFirstGroup(gate) {
      if (w.holds.beforeGate?.pane === pane) await w.holds.beforeGate.hold.wait();
      return gate.executeAfterFresh("fresh", async (group) => {
        w.mutations.push(pane);
        write(group.operations);
        if (w.holds.duringMutation?.pane === pane) await w.holds.duringMutation.hold.wait();
        return { status: "committed" };
      });
    },
    async executeGroup(group) {
      w.mutations.push(pane);
      write(group.operations);
      return { status: "committed" };
    },
    async reconcileFailedGroup(group) {
      return group.operations.map(() => "unknown_within_batch" as const);
    },
    async verifyPlan() {
      return JSON.stringify(w.cells.get("Sheet1!A1")) === JSON.stringify([["ok"]]);
    },
  };
}

interface ActiveTurnView {
  turnId: string;
  sessionId: string | null;
  documentId: string | null;
}

interface CoordinatorBridgeDecision {
  bridgeStateInvalidatesActiveTurn?: (
    state: BridgePublicState,
    active: ActiveTurnView | null,
    previous: { readonly sessionId: string | null; readonly documentId: string | null },
  ) => boolean;
}

type RouteDependencies = Parameters<typeof routeModule.createServerOwnedPlanRoute>[0];

interface Pane {
  readonly label: string;
  readonly lifecycle: OfficeBridgeLifecycleController;
  readonly bridge: () => OfficeCompanionBridge;
  readonly coordinator: TaskPaneTurnCoordinator;
  readonly authorizations: string[];
  readonly state: { planExecuteToolCalls: number; forcedStops: number };
  beginTurn(turnId: string): { handle: TaskPaneTurnHandle; controller: AbortController };
  prepare(handle: TaskPaneTurnHandle): Promise<TaskPaneTurnPreparation>;
  finish(handle: TaskPaneTurnHandle): void;
  backgroundPoll(): void;
}

const cleanups: (() => Promise<void> | void)[] = [];

beforeEach(() => {
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(T_BASE);
});

afterEach(async () => {
  while (cleanups.length > 0) {
    try { await cleanups.pop()!(); } catch { /* best-effort test teardown */ }
  }
  vi.useRealTimers();
});

async function paneInstance(w: World, label: string, adoptSessionId?: string): Promise<Pane> {
  if (adoptSessionId !== undefined) w.companion.adoptNext(adoptSessionId);
  const intervals: { handler: () => void; ms: number }[] = [];
  const bindings = new ServerOwnedTurnBindingStore({ now: () => Date.now(), ttlMs: 90_000, maxEntries: 32, maxBytes: 64 * 1024 });
  const state = {
    active: null as ActiveTurnView | null,
    activeRequest: null as TrustedWorkbookRequest | null,
    activeTurnStart: null as number | null,
    last: { sessionId: null as string | null, documentId: null as string | null },
    forcedStops: 0,
    planExecuteToolCalls: 0,
  };
  const authorizations: string[] = [];
  let route!: ReturnType<typeof routeModule.createServerOwnedPlanRoute>;
  let bridge: OfficeCompanionBridge | undefined;
  // eslint-disable-next-line prefer-const
  let coordinator: TaskPaneTurnCoordinator;
  /** Test copy of taskpane.ts updateBridgeState's turn/binding decision (r3b harness). */
  const onState = (next: BridgePublicState): void => {
    const previous = { sessionId: state.last.sessionId, documentId: state.last.documentId };
    const api = coordinator as unknown as CoordinatorBridgeDecision;
    let invalidated: boolean;
    if (typeof api.bridgeStateInvalidatesActiveTurn === "function") {
      invalidated = api.bridgeStateInvalidatesActiveTurn(next, state.active, previous);
    } else {
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
        fetchImpl: w.companion.fetch as typeof fetch,
        documentId: async () => DOC,
        paneId: PANE,
        setIntervalImpl: ((handler: () => void, ms: number) => {
          intervals.push({ handler, ms });
          return intervals.length;
        }) as unknown as typeof setInterval,
        clearIntervalImpl: (() => undefined) as unknown as typeof clearInterval,
      },
    );
    return bridge;
  }, onState);
  coordinator = new TaskPaneTurnCoordinator({
    getOrCreateDocumentId: async () => DOC,
    getBridgeIdentity: () => ({ sessionId: lifecycle.sessionId, documentId: lifecycle.documentId }),
    reconnectForDocumentChange: (expected) => lifecycle.reconnectForDocumentChange(expected),
    bindings,
    paneId: PANE,
  });
  // Records each binding decision (the route maps every refusal to one fixed payload).
  const authorizeAndRun = coordinator.authorizeAndRun.bind(coordinator);
  (coordinator as unknown as { authorizeAndRun: typeof authorizeAndRun }).authorizeAndRun =
    (async (serverOperationId: string, continuation: never) => {
      const result = await authorizeAndRun(serverOperationId, continuation);
      authorizations.push(result.ok === true ? "ok" : result.reason);
      return result;
    }) as typeof authorizeAndRun;
  const adapter = adapterFor(w, label);
  route = routeModule.createServerOwnedPlanRoute({
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
    readRuntimeContext: async () => RUNTIME,
    createAdapter: () => adapter,
    createStore: () => newStore(w, label),
    inspect: async () => "{}",
    awaitPaneConfirmation: (binding: Parameters<TaskPaneTurnCoordinator["waitForWorkbookPlanConfirmation"]>[0], signal?: AbortSignal) =>
      coordinator.waitForWorkbookPlanConfirmation(binding, signal),
    executePrimitive: async () => JSON.stringify({ ok: true }),
    now: () => Date.now(),
  } as unknown as RouteDependencies);
  await lifecycle.start();
  cleanups.push(async () => {
    coordinator.stopCurrentTurn();
    await lifecycle.stop();
  });
  return {
    label, lifecycle, bridge: () => bridge!, coordinator, authorizations, state,
    beginTurn(turnId) {
      const controller = new AbortController();
      const trustedRequest = classifyWorkbookRequest(EXPLICIT_PROMPT);
      const turnStart = Date.now();
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
    finish(handle) {
      coordinator.finishTurn(handle);
      if (state.active?.turnId === handle.turnId) state.active = null;
    },
    backgroundPoll() {
      const fastest = Math.min(...intervals.map((entry) => entry.ms));
      const poll = [...intervals].reverse().find((entry) => entry.ms === fastest);
      poll!.handler();
    },
  };
}

/** Flushes real macrotasks without touching the faked Date. */
async function flushUntil(condition: () => boolean, diag = ""): Promise<void> {
  for (let attempt = 0; attempt < 400 && !condition(); attempt += 1) {
    await new Promise<void>((done) => setImmediate(done));
  }
  expect(condition(), diag).toBe(true);
}

interface Delivery {
  readonly ticketId: string;
  readonly posted: { status: number; body: Record<string, unknown> }[];
  readonly handle: TaskPaneTurnHandle;
}

/** One pane turn: begin + prepare, the companion enqueues `operationId`, one bridge poll serves it. */
async function deliver(
  w: World,
  pane: Pane,
  turnId: string,
  operationId: string,
  plan: Record<string, unknown> = PLAN,
  options: { keepTurn?: boolean; maxDeliveries?: number } = {},
): Promise<Delivery> {
  const { handle } = pane.beginTurn(turnId);
  const prepared = await pane.prepare(handle);
  expect(prepared.status, `${pane.label} prepare`).toBe("active");
  const invoked = w.companion.invoke(pane.lifecycle.sessionId!, operationId, plan, options.maxDeliveries ?? 1);
  expect(invoked.status).toBe(202);
  await expect(pane.bridge().pollOnce()).resolves.toBe(true);
  if (options.keepTurn !== true) pane.finish(handle);
  return { ticketId: invoked.ticketId!, posted: w.companion.postedFor(invoked.ticketId!), handle };
}

function diag(w: World): string {
  return `diag: mutations=${JSON.stringify(w.mutations)} claims=${JSON.stringify(w.claims.map((c) => `${c.pane}:${c.kind}`))} ` +
    `executeStarts=${JSON.stringify(w.companion.executeStarts)} results=${JSON.stringify(w.companion.results)}`;
}

function expectCommitted(w: World, delivery: Delivery): Record<string, unknown> {
  expect(delivery.posted, diag(w)).toHaveLength(1);
  expect(delivery.posted[0].status).toBe(200);
  expect(delivery.posted[0].body).toMatchObject({ ticketId: delivery.ticketId, executed: true, result: { ok: true, action: "workbook_plan" } });
  const result = delivery.posted[0].body.result as { result?: { outcome?: string } };
  expect(result.result?.outcome).toBe("committed");
  return delivery.posted[0].body;
}

/** The exact A5 shape: executed:false + the fixed not_started payload, and no execute-start for the ticket. */
function expectA5(w: World, delivery: Delivery): void {
  expect(delivery.posted, diag(w)).toHaveLength(1);
  expect(delivery.posted[0].body, diag(w)).toEqual({ ticketId: delivery.ticketId, executed: false, result: A5_PAYLOAD });
  expect(w.companion.executeStartsFor(delivery.ticketId)).toBe(0);
}

/** R4e: a replayed already-executed op: executed:false + already_executed with the recorded report, and no execute-start for the ticket. */
function expectReplayedExecuted(w: World, delivery: Delivery, cached: unknown): void {
  expect(delivery.posted, diag(w)).toHaveLength(1);
  expect(delivery.posted[0].body, diag(w)).toEqual({
    ticketId: delivery.ticketId,
    executed: false,
    result: {
      ok: false,
      action: "workbook_plan",
      reason: "already_executed",
      summary: "This workbook operation may have been applied earlier and was not executed again.",
      result: { execution: cached },
    },
  });
  expect(w.companion.executeStartsFor(delivery.ticketId)).toBe(0);
}

function expectNotExecuted(w: World, delivery: Delivery): void {
  expect(delivery.posted, diag(w)).toHaveLength(1);
  expect(delivery.posted[0].body, diag(w)).toMatchObject({ ticketId: delivery.ticketId, executed: false, result: { ok: false } });
  expect(w.companion.executeStartsFor(delivery.ticketId)).toBe(0);
}

function lastClaim(w: World): { pane: string; kind: string; report?: unknown } {
  expect(w.claims.length, diag(w)).toBeGreaterThan(0);
  return w.claims[w.claims.length - 1];
}

/** A commits X on a fresh session at T_BASE under pane turn 1 (mutation 1). */
async function committedX(w: World): Promise<{ a: Pane; s: string; first: Delivery }> {
  const a = await paneInstance(w, "A");
  const s = a.lifecycle.sessionId!;
  const first = await deliver(w, a, turn(1), X);
  expectCommitted(w, first);
  expect(w.mutations).toEqual(["A"]);
  expect(w.companion.acceptedExecuteStarts(X)).toBe(1);
  expect(w.claims.map((claim) => claim.kind)).toEqual(["claimed"]);
  expect(w.finishes).toHaveLength(1);
  expect(w.finishes[0].ok).toBe(true);
  return { a, s, first };
}

describe("R3d TX fixture fidelity (§4.3)", () => {
  it("TX0 the frozen wire fixture is byte-pinned and supplies every double response body", () => {
    expect(createHash("sha256").update(FIXTURE_BYTES).digest("hex")).toBe(FIXTURE_SHA256);
    expect(fill(FX.unauthorized401.response!.body, {})).toEqual({ ok: false, error: "Unauthorized", category: "Auth" });
    expect(fill(FX.invoke404.response!.body, {})).toEqual({ ok: false, error: "Session not found", category: "NotFound" });
    expect(fill(FX.executeStart409.response!.body, {})).toEqual(fill(FX.leaseExecuteStart409.response!.body, {}));
    expect(Object.keys(FX.pending200.response!.body as Record<string, Json>).sort())
      .toEqual(["arguments", "requestId", "ticketId", "tool", "trusted"]);
    expect((FX.pending200.response!.body as Record<string, Json>).arguments).toEqual(PLAN);
    // The route's not_started payload is exactly the result body the fixture's C09 pane posts.
    expect((FX.result409 as { body?: { result?: unknown } }).body?.result).toEqual(A5_PAYLOAD);
  });
});

describe("R3d TX restart/replay matrix (pane bridge -> route -> store), ≤ 1 mutation per operation", () => {
  it("TX1 [synthetic] a new pane instance on the same session gets op X again: store terminal, already_executed, no second mutation", async () => {
    const w = world();
    const { s } = await committedX(w);
    const before = w.storage.raw();
    const b = await paneInstance(w, "B", s);
    expect(b.lifecycle.sessionId).toBe(s);
    const again = await deliver(w, b, turn(2), X);
    expect(lastClaim(w).pane).toBe("B");
    expect(lastClaim(w).kind, diag(w)).toBe("terminal");
    expect(lastClaim(w).report).toEqual(w.finishes[0].report);
    expect(w.storage.raw()).toBe(before);
    expectReplayedExecuted(w, again, w.finishes[0].report);
    expect(w.mutations).toEqual(["A"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(1);
  });

  it("TX2 reload while claimed: the deleted session cancels the claimed ticket, a stale invoke is 404, S2 serves a fresh op once", async () => {
    const w = world();
    const a = await paneInstance(w, "A");
    const s = a.lifecycle.sessionId!;
    const gate = hold();
    w.holds.beforeGate = { pane: "A", hold: gate };
    const { handle } = a.beginTurn(turn(1));
    expect((await a.prepare(handle)).status).toBe("active");
    const invoked = w.companion.invoke(s, X, PLAN);
    const polled = a.bridge().pollOnce().catch((error: unknown) => `rejected:${String(error)}`);
    await flushUntil(() => gate.reached(), diag(w));
    const ticket = w.companion.tickets.get(invoked.ticketId!)!;
    expect(ticket.state).toBe("claimed");
    expect(w.companion.acceptedExecuteStarts()).toBe(0);
    // A is abandoned by a reload: its unload deletes S.
    await a.lifecycle.stop();
    expect(w.companion.deleted).toEqual([s]);
    expect(ticket.state).toBe("cancelled_claimed");
    const b = await paneInstance(w, "B");
    const s2 = b.lifecycle.sessionId!;
    expect(s2).not.toBe(s);
    expect(w.companion.invoke(s, X, PLAN)).toEqual({ status: 404, body: { ok: false, error: "Session not found", category: "NotFound" } });
    const fresh = await deliver(w, b, turn(2), Y);
    expectCommitted(w, fresh);
    // The abandoned instance resumes: its acknowledgement is bound to the dead session, so it never mutates.
    gate.release();
    await polled;
    await flushUntil(() => w.claims.length >= 2);
    expect(w.mutations, diag(w)).toEqual(["B"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(0);
    expect(w.companion.acceptedExecuteStarts(Y)).toBe(1);
  });

  it("TX3a [synthetic] reload while executing, young orphan: restart_rejected below and at H, conflict for a changed digest, never a second mutation", async () => {
    const w = world();
    const a = await paneInstance(w, "A");
    const s = a.lifecycle.sessionId!;
    const gate = hold();
    w.holds.duringMutation = { pane: "A", hold: gate };
    const { handle } = a.beginTurn(turn(1));
    expect((await a.prepare(handle)).status).toBe("active");
    w.companion.invoke(s, X, PLAN);
    void a.bridge().pollOnce().catch(() => undefined);
    await flushUntil(() => gate.reached(), diag(w));
    const createdAt = Date.now();
    const raw = w.storage.raw();
    expect(w.storage.records()).toMatchObject([{ state: "in_flight", createdAt }]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(1);

    vi.setSystemTime(createdAt + 1_000);
    const b = await paneInstance(w, "B", s);
    const young = await deliver(w, b, turn(2), X);
    expect(lastClaim(w)).toMatchObject({ pane: "B", kind: "restart_rejected" });
    expectA5(w, young);
    expect(w.storage.raw()).toBe(raw);

    vi.setSystemTime(createdAt + H);
    const c = await paneInstance(w, "C", s);
    const boundary = await deliver(w, c, turn(3), X);
    expect(lastClaim(w)).toMatchObject({ pane: "C", kind: "restart_rejected" });
    expectA5(w, boundary);
    expect(w.storage.raw()).toBe(raw);

    const d = await paneInstance(w, "D", s);
    const changed = await deliver(w, d, turn(4), X, CHANGED_PLAN);
    expect(lastClaim(w)).toMatchObject({ pane: "D", kind: "conflict" });
    expectA5(w, changed);
    expect(w.storage.raw()).toBe(raw);
    expect(w.storage.records()).toMatchObject([{ state: "in_flight", createdAt }]);
    expect(w.mutations).toEqual(["A"]);
    expect(w.companion.acceptedExecuteStarts()).toBe(1);
  });

  it("TX3b [synthetic] reload while executing, aged orphan: the R3c marker, restart_rejected for any digest, one reconciliation", async () => {
    const w = world();
    const a = await paneInstance(w, "A");
    const s = a.lifecycle.sessionId!;
    const gate = hold();
    w.holds.duringMutation = { pane: "A", hold: gate };
    const { handle } = a.beginTurn(turn(1));
    expect((await a.prepare(handle)).status).toBe("active");
    w.companion.invoke(s, X, PLAN);
    void a.bridge().pollOnce().catch(() => undefined);
    await flushUntil(() => gate.reached(), diag(w));
    const createdAt = Date.now();
    const [orphan] = w.storage.records();
    expect(orphan).toMatchObject({ state: "in_flight", createdAt });

    const markedAt = createdAt + H + 1;
    vi.setSystemTime(markedAt);
    const b = await paneInstance(w, "B", s);
    const same = await deliver(w, b, turn(2), X);
    // Raw storage first: the record is now the marker stamped with the marking time.
    const records = w.storage.records();
    expect(records, diag(w)).toHaveLength(1);
    expect(records[0]).toMatchObject({
      boundary: orphan.boundary, transactionId: orphan.transactionId, digest: orphan.digest, state: "terminal", createdAt: markedAt,
    });
    expect(records[0].report).toMatchObject({ outcome: "restart_rejected", operations: [] });
    expect(lastClaim(w)).toMatchObject({ pane: "B", kind: "restart_rejected" });
    expectA5(w, same);
    const marker = w.storage.raw();

    const c = await paneInstance(w, "C", s);
    const changed = await deliver(w, c, turn(3), X, CHANGED_PLAN);
    expect(lastClaim(w)).toMatchObject({ pane: "C", kind: "restart_rejected" });
    expectA5(w, changed);
    expect(w.storage.raw()).toBe(marker);
    const counter = new BrowserIdempotencyStore(w.storage, { clock: () => Date.now(), coordinator: w.locks });
    expect(await counter.reconciliationRequiredCount()).toBe(1);
    expect(w.mutations).toEqual(["A"]);
    expect(w.companion.acceptedExecuteStarts()).toBe(1);
  });

  it("TX4 companion restart: the next poll is a fixture 401, auto-reconnect to S2, the old session invoke is 404, a new op commits once", async () => {
    const w = world();
    const { a, s } = await committedX(w);
    w.companion.restart();
    a.backgroundPoll();
    await flushUntil(() => a.lifecycle.sessionId !== null && a.lifecycle.sessionId !== s, diag(w));
    expect(w.companion.unauthorized).toBeGreaterThanOrEqual(1);
    const s2 = a.lifecycle.sessionId!;
    expect(w.companion.invoke(s, X, PLAN)).toEqual({ status: 404, body: { ok: false, error: "Session not found", category: "NotFound" } });
    const fresh = await deliver(w, a, turn(2), Y);
    expectCommitted(w, fresh);
    expect(w.companion.tickets.get(fresh.ticketId)!.sessionId).toBe(s2);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(1);
    expect(w.companion.acceptedExecuteStarts(Y)).toBe(1);
    expect(w.mutations).toEqual(["A", "A"]);
  });

  it("TX5c(i) [synthetic, out of model, residual] the same op id on a different session through a new pane instance is a new boundary and commits again", async () => {
    // Observed @8637337 and recorded as a residual (R6): the pane keys on (session, document, op);
    // only the plugin derivation op = uuid5(session, turn) (PX9) makes this input unreachable.
    const w = world();
    const { s } = await committedX(w);
    const b = await paneInstance(w, "B");
    expect(b.lifecycle.sessionId).not.toBe(s);
    const other = await deliver(w, b, turn(2), X);
    expectCommitted(w, other);
    expect(lastClaim(w)).toMatchObject({ pane: "B", kind: "claimed" });
    expect(w.mutations).toEqual(["A", "B"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(2);
  });

  it("TX5c(ii) [synthetic, out of model] the same op id on a different session in the same pane instance within 90 s is refused by the op-keyed tombstone", async () => {
    const w = world();
    const { a, s } = await committedX(w);
    w.companion.restart();
    a.backgroundPoll();
    await flushUntil(() => a.lifecycle.sessionId !== null && a.lifecycle.sessionId !== s, diag(w));
    vi.setSystemTime(T_BASE + 1_000);
    const other = await deliver(w, a, turn(2), X);
    expect(a.authorizations[a.authorizations.length - 1]).toBe(REPLAYED);
    expectNotExecuted(w, other);
    expect(w.claims).toHaveLength(1);
    expect(w.mutations).toEqual(["A"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(1);
  });

  it("TX6a binding-store capacity (production 90 s / 32 / 64 KiB): a 33rd op is refused with zero mutation, then binds and commits once after 90 s", async () => {
    const w = world();
    const a = await paneInstance(w, "A");
    for (let n = 1; n <= 32; n += 1) expectCommitted(w, await deliver(w, a, turn(n), op(100 + n)));
    expect(w.mutations).toHaveLength(32);
    const raw = w.storage.raw();
    expect(w.storage.records()).toHaveLength(32);
    const over = op(133);
    const refused = await deliver(w, a, turn(33), over);
    expect(a.authorizations[a.authorizations.length - 1]).toBe(CAPACITY);
    expectNotExecuted(w, refused);
    expect(w.companion.acceptedExecuteStarts(over)).toBe(0);
    expect(w.mutations).toHaveLength(32);
    expect(w.storage.raw()).toBe(raw);
    vi.setSystemTime(T_BASE + 90_000);
    // [synthetic] redelivery of the refused op.
    const later = await deliver(w, a, turn(34), over);
    expect(a.authorizations[a.authorizations.length - 1]).toBe("ok");
    expectCommitted(w, later);
    expect(w.companion.acceptedExecuteStarts(over)).toBe(1);
    expect(w.mutations).toHaveLength(33);
  });

  it("TX6b executor-store capacity with two young terminals: persistence_unavailable, zero mutation, both terminals byte-unchanged and still replay as terminal", async () => {
    const w = world();
    w.storeOptions = { maxEntries: 2 };
    const a = await paneInstance(w, "A");
    expectCommitted(w, await deliver(w, a, turn(1), op(11)));
    expectCommitted(w, await deliver(w, a, turn(2), op(12)));
    const raw = w.storage.raw();
    const third = await deliver(w, a, turn(3), op(13));
    expect(lastClaim(w).kind, diag(w)).toBe("persistence_unavailable");
    expectA5(w, third);
    expect(w.storage.raw()).toBe(raw);
    expect(w.mutations).toHaveLength(2);
    vi.setSystemTime(T_BASE + 91_000);
    for (const [n, replayed] of [[4, op(11)], [5, op(12)]] as const) {
      // [synthetic] redelivery of a committed op after the binding TTL.
      const again = await deliver(w, a, turn(n), replayed);
      expect(lastClaim(w).kind, diag(w)).toBe("terminal");
      expectReplayedExecuted(w, again, w.finishes[n - 4].report);
      expect(w.companion.acceptedExecuteStarts(replayed)).toBe(1);
    }
    expect(w.storage.raw()).toBe(raw);
    expect(w.mutations).toHaveLength(2);
  });

  it("TX7a [synthetic] the binding TTL passes inside the store window: refused at 89 999 ms, admitted at 90 000 ms, store terminal, already_executed, mutation stays 1", async () => {
    const w = world();
    const { a } = await committedX(w);
    const raw = w.storage.raw();
    vi.setSystemTime(T_BASE + 89_999);
    const early = await deliver(w, a, turn(2), X);
    expect(a.authorizations[a.authorizations.length - 1]).toBe(REPLAYED);
    expectNotExecuted(w, early);
    expect(w.claims).toHaveLength(1);
    vi.setSystemTime(T_BASE + 90_000);
    const admitted = await deliver(w, a, turn(3), X);
    expect(a.authorizations[a.authorizations.length - 1]).toBe("ok");
    expect(lastClaim(w)).toMatchObject({ pane: "A", kind: "terminal" });
    expect(lastClaim(w).report).toEqual(w.finishes[0].report);
    expectReplayedExecuted(w, admitted, w.finishes[0].report);
    expect(w.storage.raw()).toBe(raw);
    expect(w.mutations).toEqual(["A"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(1);
  });

  it("TX7b [synthetic] past H but under 24 h with no capacity pressure: store terminal, already_executed, mutation stays 1", async () => {
    const w = world();
    const { a } = await committedX(w);
    vi.setSystemTime(T_BASE + DAY - 60_000);
    const late = await deliver(w, a, turn(2), X);
    expect(lastClaim(w).kind, diag(w)).toBe("terminal");
    expectReplayedExecuted(w, late, w.finishes[0].report);
    expect(w.mutations).toEqual(["A"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(1);
  });

  it("TX7c [synthetic, residual] past the 24 h TTL the store has forgotten op X and the pane executes it again", async () => {
    // Observed @8637337 and recorded as the R3c post-window residual (R6), anchored by the
    // companion's no-requeue rule and per-process session nonces (R3a GP2/GP3).
    const w = world();
    const { a } = await committedX(w);
    vi.setSystemTime(T_BASE + DAY + 1);
    const forgotten = await deliver(w, a, turn(2), X);
    expect(lastClaim(w).kind, diag(w)).toBe("claimed");
    expectCommitted(w, forgotten);
    expect(w.mutations).toEqual(["A", "A"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(2);
  });

  it("TX8a [synthetic] concurrent duplicate delivery to two pane instances: one claim, one restart_rejected, one execute-start, one mutation", async () => {
    const w = world();
    const a = await paneInstance(w, "A");
    const s = a.lifecycle.sessionId!;
    const b = await paneInstance(w, "B", s);
    const ta = a.beginTurn(turn(1));
    const tb = b.beginTurn(turn(2));
    expect((await a.prepare(ta.handle)).status).toBe("active");
    expect((await b.prepare(tb.handle)).status).toBe("active");
    // Deterministic overlap (§4.4; R4e-2 TX8a fix): B is held before its store claim until A has
    // claimed and is parked inside its first mutation, so B always meets A's young in-flight record.
    const gate = hold();
    const bClaim = hold();
    w.holds.duringMutation = { pane: "A", hold: gate };
    w.holds.beforeClaim = { pane: "B", hold: bClaim };
    const invoked = w.companion.invoke(s, X, PLAN, 2);
    const polls = Promise.allSettled([a.bridge().pollOnce(), b.bridge().pollOnce()]);
    await flushUntil(() => bClaim.reached() && gate.reached(), diag(w));
    expect(w.claims.map((claim) => `${claim.pane}:${claim.kind}`), diag(w)).toEqual(["A:claimed"]);
    bClaim.release();
    await flushUntil(() => w.claims.length === 2, diag(w));
    gate.release();
    const settled = await polls;
    expect(settled.map((entry) => entry.status)).toEqual(["fulfilled", "fulfilled"]);
    expect(w.claims.map((claim) => `${claim.pane}:${claim.kind}`), diag(w)).toEqual(["A:claimed", "B:restart_rejected"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(1);
    expect(w.mutations).toEqual(["A"]);
    const posted = w.companion.postedFor(invoked.ticketId!);
    expect(posted.filter((entry) => entry.body.executed === true && entry.status === 200)).toHaveLength(1);
    expect(posted.filter((entry) => entry.body.executed === false)).toEqual([
      { status: 409, body: { ticketId: invoked.ticketId, executed: false, result: A5_PAYLOAD } },
    ]);
  });

  // R4f (R4-TX8b fixed): a ticket redelivered to the same bridge after its result post is dropped by the
  // bridge's per-session dispatched-ticket record before executeTool. The real companion never redelivers
  // a plan (M7); the primitive requeue path is covered in office-bridge-lifecycle-r4f.test.ts.
  it("TX8b [synthetic, R4-TX8b fixed] the same ticket polled twice by one bridge: executeTool once, one result post, one mutation, one execute-start", async () => {
    const w = world();
    const a = await paneInstance(w, "A");
    const first = await deliver(w, a, turn(1), X, PLAN, { keepTurn: true, maxDeliveries: 2 });
    expectCommitted(w, first);
    await expect(a.bridge().pollOnce()).resolves.toBe(false);
    const posted = w.companion.postedFor(first.ticketId);
    const observed = `executeTool=${a.state.planExecuteToolCalls} posts=${JSON.stringify(posted)} ${diag(w)}`;
    expect(w.companion.tickets.get(first.ticketId)!.deliveries, observed).toBe(2);
    expect(a.state.planExecuteToolCalls, observed).toBe(1);
    expect(w.mutations, observed).toEqual(["A"]);
    expect(w.companion.acceptedExecuteStarts(X), observed).toBe(1);
    expect(w.claims.map((claim) => claim.kind), observed).toEqual(["claimed"]);
    expect(posted.filter((entry) => entry.body.executed === true && entry.status === 200), observed).toHaveLength(1);
    expect(posted.filter((entry) => entry.body.executed === false), observed).toEqual([]);
    expect(posted, observed).toHaveLength(1);
  });

  it("TX9a a corrupt store fails closed: executed:false, no execute-start, zero mutation, stored bytes unchanged", async () => {
    const w = world();
    w.storage.values.set(STORAGE_KEY, "{not json");
    const a = await paneInstance(w, "A");
    const refused = await deliver(w, a, turn(1), X);
    expect(lastClaim(w).kind).toBe("persistence_unavailable");
    expectA5(w, refused);
    expect(w.storage.raw()).toBe("{not json");
    expect(w.mutations).toEqual([]);
    expect(w.companion.acceptedExecuteStarts()).toBe(0);
  });

  it("TX9b an unavailable store (getItem throws) fails closed with zero mutation and unchanged bytes", async () => {
    const w = world();
    const a = await paneInstance(w, "A");
    expectCommitted(w, await deliver(w, a, turn(1), Y));
    const raw = w.storage.raw();
    w.storage.getThrows = true;
    const refused = await deliver(w, a, turn(2), X);
    expect(lastClaim(w).kind).toBe("persistence_unavailable");
    expectA5(w, refused);
    w.storage.getThrows = false;
    expect(w.storage.raw()).toBe(raw);
    expect(w.mutations).toEqual(["A"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(0);
  });

  it("TX9b' an unavailable store (setItem throws a quota DOMException) fails closed with zero mutation and unchanged bytes", async () => {
    const w = world();
    const a = await paneInstance(w, "A");
    expectCommitted(w, await deliver(w, a, turn(1), Y));
    const raw = w.storage.raw();
    w.storage.setFault = () => true;
    const refused = await deliver(w, a, turn(2), X);
    expect(lastClaim(w).kind).toBe("persistence_unavailable");
    expectA5(w, refused);
    w.storage.setFault = null;
    expect(w.storage.raw()).toBe(raw);
    expect(w.mutations).toEqual(["A"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(0);
  });

  it("TX9c [synthetic] finish fails after the mutation: acknowledged-write payload, record stays in_flight, later tickets restart_rejected young and as a marker", async () => {
    const w = world();
    const a = await paneInstance(w, "A");
    const s = a.lifecycle.sessionId!;
    w.storage.setFault = (value) => (JSON.parse(value) as { state?: string }[]).some((record) => record.state === "terminal");
    const written = await deliver(w, a, turn(1), X);
    w.storage.setFault = null;
    expect(w.finishes.map((entry) => entry.ok)).toEqual([false]);
    expect(written.posted, diag(w)).toHaveLength(1);
    expect(written.posted[0].status).toBe(200);
    expect(written.posted[0].body.executed).toBe(true);
    const payload = written.posted[0].body.result as Record<string, unknown>;
    expect(payload).not.toHaveProperty("reason");
    if (payload.state === "uncertain") {
      expect(payload).toEqual({
        ok: false, action: "workbook_plan", state: "uncertain", error: "Workbook plan execution ended without a confirmed result.",
      });
    } else {
      const inner = payload.result as { outcome?: string; execution?: unknown };
      expect(payload.ok).toBe(false);
      expect(["partial", "verification_failed"]).toContain(inner.outcome);
      expect(inner.execution).toBeDefined();
    }
    const createdAt = Date.now();
    expect(w.storage.records()).toMatchObject([{ state: "in_flight", createdAt }]);
    const raw = w.storage.raw();

    vi.setSystemTime(createdAt + 1_000);
    const b = await paneInstance(w, "B", s);
    const young = await deliver(w, b, turn(2), X);
    expect(lastClaim(w)).toMatchObject({ pane: "B", kind: "restart_rejected" });
    expectA5(w, young);
    expect(w.storage.raw()).toBe(raw);

    vi.setSystemTime(createdAt + H + 1);
    const c = await paneInstance(w, "C", s);
    const aged = await deliver(w, c, turn(3), X);
    expect(w.storage.records()).toMatchObject([{ state: "terminal", createdAt: createdAt + H + 1, report: { outcome: "restart_rejected" } }]);
    expect(lastClaim(w)).toMatchObject({ pane: "C", kind: "restart_rejected" });
    expectA5(w, aged);
    expect(w.mutations).toEqual(["A"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(1);
  });

  it("TX10 [synthetic] changed digest for the same op at a new pane instance: store conflict, A5, record byte-unchanged, mutation stays 1", async () => {
    const w = world();
    const { s } = await committedX(w);
    const raw = w.storage.raw();
    const b = await paneInstance(w, "B", s);
    const changed = await deliver(w, b, turn(2), X, CHANGED_PLAN);
    expect(lastClaim(w)).toMatchObject({ pane: "B", kind: "conflict" });
    expectA5(w, changed);
    expect(w.storage.raw()).toBe(raw);
    expect(w.mutations).toEqual(["A"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(1);
  });

  it("TX11 an independent later identical write: a new op with the identical plan in a new pane turn commits once with its own boundary and transaction", async () => {
    const w = world();
    const { a } = await committedX(w);
    const later = await deliver(w, a, turn(2), Y);
    expectCommitted(w, later);
    expect(w.claims.map((claim) => claim.kind)).toEqual(["claimed", "claimed"]);
    const records = w.storage.records();
    expect(records).toHaveLength(2);
    expect(records.map((record) => record.state)).toEqual(["terminal", "terminal"]);
    expect(records[0].boundary).not.toBe(records[1].boundary);
    expect(records[0].transactionId).not.toBe(records[1].transactionId);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(1);
    expect(w.companion.acceptedExecuteStarts(Y)).toBe(1);
    expect(w.mutations).toEqual(["A", "A"]);
  });

  it("TX12a [synthetic, R4e: truthful replay, R4-A5 closed] a committed op re-presented under a new pane turn at a new pane instance is reported as already_executed with its recorded outcome", async () => {
    const w = world();
    const { s } = await committedX(w);
    const raw = w.storage.raw();
    const b = await paneInstance(w, "B", s);
    const replay = await deliver(w, b, turn(2), X);
    expect(lastClaim(w)).toMatchObject({ pane: "B", kind: "terminal" });
    expect(w.storage.raw()).toBe(raw);
    expectReplayedExecuted(w, replay, w.finishes[0].report);
    expect(w.mutations).toEqual(["A"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(1);
  });

  it("TX12b [synthetic, R4e: truthful replay, R4-A5 closed] the same pane instance after 90 s under a new pane turn is reported as already_executed with its recorded outcome", async () => {
    const w = world();
    const { a } = await committedX(w);
    const raw = w.storage.raw();
    vi.setSystemTime(T_BASE + 90_000);
    const replay = await deliver(w, a, turn(2), X);
    expect(lastClaim(w)).toMatchObject({ pane: "A", kind: "terminal" });
    expect(w.storage.raw()).toBe(raw);
    expectReplayedExecuted(w, replay, w.finishes[0].report);
    expect(w.mutations).toEqual(["A"]);
    expect(w.companion.acceptedExecuteStarts(X)).toBe(1);
  });
});
