import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import type { BridgeTrustedPlanEnvelopeV2, OfficeCompanionBridge } from "../officeBridge";
import { createExcelCompanionBridge } from "../officeBridgeExcel";
import {
  createWorkbookPlanTurn,
  ServerOwnedTurnBindingStore,
  type WorkbookPlanVerificationAdapter,
} from "../workbookPlanOrchestrator";
import {
  BrowserIdempotencyStore,
  DeterministicIdempotencyCoordinator,
  MemoryIdempotencyStorage,
  createTrustedOperationBoundary,
  type WorkbookPlanExecutionAdapter,
} from "../workbookPlanExecutor";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import {
  buildWorkbookPlanPolicyInput,
  preflightWorkbookPlanResources,
  type WorkbookPlanRouteRuntimeContext,
} from "../workbookPlanRoute";
import { classifyWorkbookRequest, type TrustedWorkbookRequest } from "../workbookPlanRequest";
import { TaskPaneTurnCoordinator, type TaskPaneTurnHandle } from "../taskpane/taskPaneTurnCoordinator";
import { createServerOwnedPlanRoute, EMPTY_RESOURCE_CONTEXT } from "../taskpane/serverOwnedPlanRoute";
import { createServerPrimitiveGate, PRIMITIVE_DENIAL_MESSAGES } from "../primitiveMutationAuthority";

/**
 * R4e integration (dev3-r4e-release §5, file 5): the REAL createExcelCompanionBridge ->
 * createServerOwnedPlanRoute -> createWorkbookPlanTurn -> executeAcceptedWorkbookPlan chain.
 * Harness copied from workbook-plan-production-route.test.ts:286-362 (nothing imported from other
 * test files), with createStore building a new store per ticket (taskpane.ts shape) over ONE shared
 * storage and lock coordinator, so a terminal replay can be seeded by an earlier turn.
 * - RT1/RT6: a replayed already-executed op is reported as already_executed with its recorded report.
 * - RT2-RT4: every other cached or refused outcome keeps its base payload byte-for-byte.
 * - RT5: GA3 end to end (deadline crossed inside the execute-start acknowledgement).
 * - RT7: the direct-route shape passes the bare signal and never meets a deadline refusal.
 * - W1/W2 (release record row a) and D1 (row b): test-only pins folded from the R4d reviews.
 */

const DOC = "d0c00000-0000-4000-8000-0000000000e4";
const DOC_2 = "d0c00000-0000-4000-8000-0000000000e5";
const PANE = "9a9e0000-0000-4000-8000-0000000000e4";
const REQUEST = "4e4e0000-0000-4000-8000-0000000000e4";
const DIRECT_SESSION = "aaaaaaaa-aaaa-4aaa-8aaa-0000000000e4";
const T_BASE = Date.UTC(2026, 9, 7, 12, 0, 0);
const STORAGE_KEY = "hermes.excel.workbookPlan.idempotency.v1";
const SESSION_1 = "aaaaaaaa-aaaa-4aaa-8aaa-000000000001";
const SESSION_2 = "aaaaaaaa-aaaa-4aaa-8aaa-000000000002";

const uuid = (prefix: string, n: number): string => `${prefix}-0000-4000-8000-${String(n).padStart(12, "0")}`;
const turn = (n: number): string => uuid("7a7a0000", n);
const op = (n: number): string => uuid("5e7e0000", n);
const X = op(1);

const COMPILE_CONTEXT: VerifiedWorkbookCompileContext = {
  sheetNames: new Set(["Sheet1"]),
  tablesByName: new Map(),
  supportedPrimitiveNames: new Set(["write_range", "run_calculation"]),
  supportedCapabilities: new Set(["excel_workbook_access", "workbook_plan_v1_operation_audit"]),
  chartSourceResolutionsBySheet: new Map(),
};
const RUNTIME: WorkbookPlanRouteRuntimeContext = {
  activeSheetName: "Sheet1",
  selection: { sheetName: "Sheet1", address: "A1" },
  compileContext: COMPILE_CONTEXT,
};
const PLAN = {
  version: 1,
  outcome: "execute",
  summary: "write one cell",
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }],
  verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["ok"]] }],
} as const;
const CHANGED_PLAN = { ...PLAN, summary: "write one cell (R4e changed digest)" };
/** Two effect groups (write, then calculate): a failing second group leaves a recorded partial. */
const TWO_GROUP_PLAN = {
  ...PLAN,
  summary: "write one cell then calculate",
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }, { kind: "calculate" }],
} as const;
const PROMPT = "write ok into A1";
const TRUSTED = (operationId: string): BridgeTrustedPlanEnvelopeV2 => ({
  turnOperationId: operationId,
  workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
  phase: "execute",
  turnKey: "4b4b0000-0000-4000-8000-0000000000a3",
});

const ALREADY_EXECUTED_SUMMARY = "This workbook operation may have been applied earlier and was not executed again.";
const NOT_STARTED_PAYLOAD = {
  ok: false,
  action: "workbook_plan",
  reason: "not_started",
  summary: "Workbook plan execution did not start; nothing was executed.",
};
const CANCELLED_PAYLOAD = {
  ok: false,
  action: "workbook_plan",
  reason: "cancelled",
  summary: "The workbook plan was cancelled before execution; nothing was executed.",
};
const UNCERTAIN_PAYLOAD = {
  ok: false,
  action: "workbook_plan",
  state: "uncertain",
  error: "Workbook plan execution ended without a confirmed result.",
};

function json(status: number, body?: unknown): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
  });
}

interface DoubleTicket {
  readonly ticketId: string;
  readonly tool: string;
  readonly args: Record<string, unknown>;
  readonly trusted: unknown;
  state: "pending" | "claimed" | "executing" | "terminal";
  terminalReason?: "completed" | "not_executed";
}

/** Stateful R2d-2 companion lifecycle double (copied); a new ticket id per enqueue. */
class CompanionDouble {
  readonly sessions = new Map<string, string>();
  ticket: DoubleTicket | null = null;
  readonly tickets: DoubleTicket[] = [];
  readonly results: Record<string, unknown>[] = [];
  readonly executeStarts: string[] = [];
  ackMode: "ok" | 409 = "ok";
  /** Runs when an accepted execute-start moved the ticket to executing, before the 200 answer. */
  onExecuteStart: (() => void) | undefined;
  private counter = 0;
  private ticketCounter = 0;

  enqueue(tool: string, args: Record<string, unknown>, trusted: unknown): DoubleTicket {
    this.ticketCounter += 1;
    const ticket: DoubleTicket = { ticketId: uuid("71c70000", this.ticketCounter), tool, args, trusted, state: "pending" };
    this.ticket = ticket;
    this.tickets.push(ticket);
    return ticket;
  }

  resultsFor(ticketId: string): Record<string, unknown>[] {
    return this.results.filter((body) => body.ticketId === ticketId);
  }

  executeStartsFor(ticketId: string): number {
    return this.executeStarts.filter((id) => id === ticketId).length;
  }

  readonly fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const url = String(input);
    const route = url.replace(/^.*\/excel\//, "");
    const method = init?.method ?? "GET";
    const headers = (init?.headers ?? {}) as Record<string, string>;
    const authed = this.sessions.has(headers["X-Hermes-Session-Key"] ?? "");
    const body = typeof init?.body === "string" && init.body ? JSON.parse(init.body) as Record<string, unknown> : {};
    if (method === "POST" && route === "session") {
      this.counter += 1;
      const sessionId = `aaaaaaaa-aaaa-4aaa-8aaa-${String(this.counter).padStart(12, "0")}`;
      const sessionKey = `${"k".repeat(40)}${this.counter}`;
      this.sessions.set(sessionKey, sessionId);
      return json(201, {
        ok: true, sessionId, sessionKey, protocolVersion: "1.0", expiresInSeconds: 60,
        workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
      });
    }
    if (!authed) return json(401, { error: "Unauthorized" });
    if (method === "DELETE" && route === "session") return json(200, { ok: true });
    if (method === "POST" && route === "heartbeat") return json(200, { ok: true });
    if (method === "GET" && route === "pending") {
      const ticket = this.ticket;
      if (ticket === null || ticket.state !== "pending") return json(204);
      ticket.state = "claimed";
      return json(200, { ticketId: ticket.ticketId, requestId: REQUEST, tool: ticket.tool, arguments: ticket.args, trusted: ticket.trusted });
    }
    if (method === "POST" && route === "execute-start") {
      const ticket = this.ticket;
      this.executeStarts.push(String(body.ticketId));
      if (this.ackMode === 409) return json(409, { error: "conflict" });
      if (ticket === null || ticket.ticketId !== body.ticketId || ticket.state !== "claimed") return json(409, { error: "conflict" });
      ticket.state = "executing";
      this.onExecuteStart?.();
      return json(200, { ok: true, ticketId: ticket.ticketId });
    }
    if (method === "POST" && route === "result") {
      this.results.push(body);
      const ticket = this.tickets.find((candidate) => candidate.ticketId === body.ticketId);
      if (ticket === undefined) return json(404, { error: "No ticket found" });
      if (typeof body.executed !== "boolean") return json(400, { error: "executed required" });
      if (body.executed === true && ticket.state !== "executing") return json(409, { error: "conflict" });
      if (body.executed === false && ticket.state !== "claimed") return json(409, { error: "conflict" });
      ticket.state = "terminal";
      ticket.terminalReason = body.executed === true ? "completed" : "not_executed";
      return json(200, { ok: true, ticketId: ticket.ticketId });
    }
    return json(404, { error: "route" });
  };
}

interface AdapterMode {
  observation: "fresh" | "stale";
  /** First group may have changed cells but its result cannot be reconciled. */
  unknownFirstGroup: boolean;
  /** The second effect group reports sync_failed (reconciled as not_committed). */
  failSecondGroup: boolean;
}

interface World {
  readonly companion: CompanionDouble;
  readonly storage: MemoryIdempotencyStorage;
  readonly locks: DeterministicIdempotencyCoordinator;
  readonly cells: Map<string, unknown>;
  readonly mode: AdapterMode;
  mutations: number;
  claims: string[];
  finishes: { ok: boolean; report: unknown }[];
}

function statefulAdapter(w: World): WorkbookPlanExecutionAdapter & WorkbookPlanVerificationAdapter {
  const write = (operations: readonly { readonly args: Record<string, unknown> }[]): void => {
    for (const operation of operations) {
      if (operation.args.address !== undefined) w.cells.set(`${String(operation.args.sheet)}!${String(operation.args.address)}`, operation.args.values);
    }
  };
  return {
    async executeFirstGroup(gate) {
      return gate.executeAfterFresh(w.mode.observation, async (group) => {
        w.mutations += 1;
        if (w.mode.unknownFirstGroup) return { status: "sync_failed", failureCode: "sync_failed" };
        write(group.operations);
        return { status: "committed" };
      });
    },
    async executeGroup(group) {
      w.mutations += 1;
      if (w.mode.failSecondGroup) return { status: "sync_failed", failureCode: "sync_failed" };
      write(group.operations);
      return { status: "committed" };
    },
    async reconcileFailedGroup(group) {
      return group.operations.map(() => w.mode.unknownFirstGroup ? "unknown_within_batch" as const : "not_committed" as const);
    },
    async verifyPlan() {
      return JSON.stringify(w.cells.get("Sheet1!A1")) === JSON.stringify([["ok"]]);
    },
  };
}

function world(): World {
  return {
    companion: new CompanionDouble(),
    storage: new MemoryIdempotencyStorage(),
    locks: new DeterministicIdempotencyCoordinator(),
    cells: new Map(),
    mode: { observation: "fresh", unknownFirstGroup: false, failSecondGroup: false },
    mutations: 0,
    claims: [],
    finishes: [],
  };
}

/** taskpane.ts createStore shape: a new store per ticket over the shared storage and lock coordinator. */
function newStore(w: World): BrowserIdempotencyStore {
  const store = new BrowserIdempotencyStore(w.storage, { clock: () => Date.now(), coordinator: w.locks });
  const claim = store.claim.bind(store);
  store.claim = async (...args: Parameters<typeof store.claim>) => {
    const result = await claim(...args);
    w.claims.push(result.kind);
    return result;
  };
  const finish = store.finish.bind(store);
  store.finish = async (...args: Parameters<typeof store.finish>) => {
    const ok = await finish(...args);
    w.finishes.push({ ok, report: JSON.parse(JSON.stringify(args[3])) });
    return ok;
  };
  return store;
}

interface Pane {
  readonly bridge: OfficeCompanionBridge;
  readonly coordinator: TaskPaneTurnCoordinator;
}

async function pane(w: World): Promise<Pane> {
  const bindings = new ServerOwnedTurnBindingStore({ now: () => Date.now(), ttlMs: 90_000, maxEntries: 32 });
  const state = {
    turnId: null as string | null,
    request: null as TrustedWorkbookRequest | null,
    sessionId: null as string | null,
    turnStart: null as number | null,
  };
  let route!: ReturnType<typeof createServerOwnedPlanRoute>;
  const bridge = createExcelCompanionBridge(
    undefined,
    (name, argumentsValue, trusted, control) => route.executeCompanionTool(name, argumentsValue, trusted, control),
    {
      fetchImpl: w.companion.fetch as typeof fetch,
      documentId: async () => DOC,
      paneId: PANE,
      setIntervalImpl: (() => 1) as unknown as typeof setInterval,
      clearIntervalImpl: (() => undefined) as unknown as typeof clearInterval,
    },
  );
  const coordinator = new TaskPaneTurnCoordinator({
    getOrCreateDocumentId: async () => DOC,
    getBridgeIdentity: () => ({ sessionId: bridge.sessionId, documentId: bridge.documentId }),
    reconnectForDocumentChange: async () => { throw new Error("unexpected document change"); },
    bindings,
    paneId: PANE,
  });
  await bridge.connect();
  const adapter = statefulAdapter(w);
  route = createServerOwnedPlanRoute({
    isRolloutEnabled: () => true,
    getBridgeSessionId: () => bridge.sessionId,
    getActiveTurn: () => ({
      turnOperationId: state.turnId,
      trustedRequest: state.request,
      sessionId: state.sessionId,
      turnStart: state.turnStart,
      hasCancellation: state.turnId !== null,
    }),
    coordinator,
    bindings,
    readRuntimeContext: async () => RUNTIME,
    createAdapter: () => adapter,
    createStore: () => newStore(w),
    inspect: async () => "{}",
    awaitPaneConfirmation: (binding, signal) => coordinator.waitForWorkbookPlanConfirmation(binding, signal),
    executePrimitive: async () => JSON.stringify({ ok: true }),
    now: () => Date.now(),
  });
  const current: Pane & { state: typeof state } = { bridge, coordinator, state };
  panes.set(current, state);
  return current;
}

const panes = new Map<Pane, { turnId: string | null; request: TrustedWorkbookRequest | null; sessionId: string | null; turnStart: number | null }>();

interface Delivery {
  readonly ticketId: string;
  readonly posted: Record<string, unknown>[];
  readonly handle: TaskPaneTurnHandle;
}

/** One pane turn: begin + prepare, the companion enqueues `operationId`, one bridge poll serves it, the turn finishes. */
async function deliver(w: World, p: Pane, turnId: string, operationId: string, plan: Record<string, unknown> = PLAN): Promise<Delivery> {
  const state = panes.get(p)!;
  const controller = new AbortController();
  const trustedRequest = classifyWorkbookRequest(PROMPT);
  const turnStart = Date.now();
  const handle = p.coordinator.beginTurn(turnId, controller, { trustedRequest, turnStart, cancellationSignal: controller.signal });
  const prepared = await p.coordinator.prepareTurn(handle);
  expect(prepared.status).toBe("active");
  state.turnId = turnId;
  state.request = trustedRequest;
  state.sessionId = p.bridge.sessionId;
  state.turnStart = turnStart;
  const ticket = w.companion.enqueue("excel_execute_plan", { ...plan }, TRUSTED(operationId));
  await expect(p.bridge.pollOnce()).resolves.toBe(true);
  p.coordinator.finishTurn(handle);
  state.turnId = null;
  return { ticketId: ticket.ticketId, posted: w.companion.resultsFor(ticket.ticketId), handle };
}

function diag(w: World): string {
  return `diag: mutations=${w.mutations} claims=${JSON.stringify(w.claims)} executeStarts=${JSON.stringify(w.companion.executeStarts)} results=${JSON.stringify(w.companion.results)}`;
}

function storedReports(w: World): unknown[] {
  const raw = w.storage.getItem(STORAGE_KEY);
  return raw === null ? [] : (JSON.parse(raw) as { report?: unknown }[]).map((record) => record.report);
}

function expectExecutedCommitted(w: World, delivery: Delivery): void {
  expect(delivery.posted, diag(w)).toHaveLength(1);
  expect(delivery.posted[0]).toMatchObject({ ticketId: delivery.ticketId, executed: true, result: { ok: true, action: "workbook_plan", result: { outcome: "committed" } } });
}

/** The exact R4e already_executed payload: executed:false, the recorded report, and no execute-start for this ticket. */
function expectAlreadyExecuted(w: World, delivery: Delivery, recorded: unknown): void {
  expect(delivery.posted, diag(w)).toHaveLength(1);
  expect(delivery.posted[0], diag(w)).toEqual({
    ticketId: delivery.ticketId,
    executed: false,
    result: { ok: false, action: "workbook_plan", reason: "already_executed", summary: ALREADY_EXECUTED_SUMMARY, result: { execution: recorded } },
  });
  expect(Object.keys(delivery.posted[0].result as object)).toEqual(["ok", "action", "reason", "summary", "result"]);
  expect(w.companion.executeStartsFor(delivery.ticketId)).toBe(0);
}

/** A base non-executing payload, byte-identical, with no execute-start for this ticket. */
function expectNotExecutedPayload(w: World, delivery: Delivery, payload: Record<string, unknown>): void {
  expect(delivery.posted, diag(w)).toHaveLength(1);
  expect(delivery.posted[0], diag(w)).toEqual({ ticketId: delivery.ticketId, executed: false, result: payload });
  expect(JSON.stringify(delivery.posted[0].result)).toBe(JSON.stringify(payload));
  expect(w.companion.executeStartsFor(delivery.ticketId)).toBe(0);
}

beforeEach(() => {
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(T_BASE);
});

afterEach(() => {
  panes.clear();
  vi.useRealTimers();
});

describe("R4e truthful replay of an already-executed server op (route -> orchestrator -> executor)", () => {
  it("RT1 a committed op re-presented under a new pane turn is reported as already_executed with the stored report; no execute-start, mutation stays 1", async () => {
    const w = world();
    const p = await pane(w);
    expectExecutedCommitted(w, await deliver(w, p, turn(1), X));
    const raw = w.storage.getItem(STORAGE_KEY);
    vi.setSystemTime(T_BASE + 90_000);
    const replay = await deliver(w, p, turn(2), X);
    expect(w.claims, diag(w)).toEqual(["claimed", "terminal"]);
    expect(w.storage.getItem(STORAGE_KEY)).toBe(raw);
    const [stored] = storedReports(w);
    expect(stored).toMatchObject({ outcome: "committed" });
    expect(stored).toEqual(w.finishes[0].report);
    expectAlreadyExecuted(w, replay, stored);
    expect(w.mutations).toBe(1);
    expect(w.companion.tickets[1]).toMatchObject({ state: "terminal", terminalReason: "not_executed" });
  });

  it("RT6 a recorded partial (first group committed, second sync_failed) replayed is already_executed with the partial report", async () => {
    const w = world();
    const p = await pane(w);
    w.mode.failSecondGroup = true;
    const first = await deliver(w, p, turn(1), X, TWO_GROUP_PLAN);
    expect(first.posted, diag(w)).toHaveLength(1);
    expect(first.posted[0]).toMatchObject({ executed: true, result: { ok: false, result: { outcome: "partial" } } });
    const mutationsAfterFirst = w.mutations;
    expect(mutationsAfterFirst).toBe(2);
    const [stored] = storedReports(w) as { outcome: string; operations: { state: string }[] }[];
    expect(stored.outcome).toBe("partial");
    expect(stored.operations.map((operation) => operation.state)).toEqual(["committed", "not_committed"]);
    vi.setSystemTime(T_BASE + 90_000);
    w.mode.failSecondGroup = false;
    const replay = await deliver(w, p, turn(2), X, TWO_GROUP_PLAN);
    expect(w.claims.at(-1)).toBe("terminal");
    expectAlreadyExecuted(w, replay, stored);
    expect(w.mutations).toBe(mutationsAfterFirst);
  });

  it("R6a-2 an unknown-only report replay says may have been applied, preserves the report and never starts or mutates again", async () => {
    const w = world();
    const p = await pane(w);
    w.mode.unknownFirstGroup = true;
    await deliver(w, p, turn(1), X);
    const [stored] = storedReports(w) as { outcome: string; operations: { state: string }[] }[];
    expect(stored.outcome).toBe("partial");
    expect(stored.operations.map((operation) => operation.state)).toEqual(["unknown_within_batch"]);
    expect(w.cells.size).toBe(0);
    const raw = w.storage.getItem(STORAGE_KEY);
    const attempts = w.mutations;
    w.mode.unknownFirstGroup = false;
    vi.setSystemTime(T_BASE + 90_000);
    const replay = await deliver(w, p, turn(2), X);
    expectAlreadyExecuted(w, replay, stored);
    const result = replay.posted[0].result as { summary: string };
    expect(result.summary).toContain("may have been applied");
    expect(result.summary).not.toContain("already applied");
    expect(w.mutations).toBe(attempts);
    expect(w.cells.size).toBe(0);
    expect(w.storage.getItem(STORAGE_KEY)).toBe(raw);
  });

  it("RT2 [pin] a recorded cancelled report (the acknowledgement failed on turn 1) replays with the base cancelled payload", async () => {
    const w = world();
    const p = await pane(w);
    w.companion.ackMode = 409;
    const first = await deliver(w, p, turn(1), X);
    expect(first.posted, diag(w)).toEqual([{ ticketId: first.ticketId, executed: false, result: NOT_STARTED_PAYLOAD }]);
    expect(w.mutations).toBe(0);
    expect(storedReports(w)).toMatchObject([{ outcome: "cancelled" }]);
    w.companion.ackMode = "ok";
    vi.setSystemTime(T_BASE + 90_000);
    const replay = await deliver(w, p, turn(2), X);
    expect(w.claims).toEqual(["claimed", "terminal"]);
    expectNotExecutedPayload(w, replay, CANCELLED_PAYLOAD);
    expect(w.mutations).toBe(0);
  });

  it("RT3 [pin] a recorded stale report replays with the base not_started payload", async () => {
    const w = world();
    const p = await pane(w);
    w.mode.observation = "stale";
    const first = await deliver(w, p, turn(1), X);
    expectNotExecutedPayload(w, first, NOT_STARTED_PAYLOAD);
    expect(storedReports(w)).toMatchObject([{ outcome: "stale" }]);
    w.mode.observation = "fresh";
    vi.setSystemTime(T_BASE + 90_000);
    const replay = await deliver(w, p, turn(2), X);
    expect(w.claims).toEqual(["claimed", "terminal"]);
    expectNotExecutedPayload(w, replay, NOT_STARTED_PAYLOAD);
    expect(w.mutations).toBe(0);
  });

  it("RT4 [pin] restart_rejected (young foreign orphan), conflict (changed digest) and persistence_unavailable keep the base not_started payload", async () => {
    // Young orphan: turn 1 commits, then the record is rewritten as a foreign in-flight claim (seeded).
    const orphanWorld = world();
    const orphanPane = await pane(orphanWorld);
    expectExecutedCommitted(orphanWorld, await deliver(orphanWorld, orphanPane, turn(1), X));
    const [record] = JSON.parse(orphanWorld.storage.getItem(STORAGE_KEY)!) as Record<string, unknown>[];
    orphanWorld.storage.setItem(STORAGE_KEY, JSON.stringify([{
      boundary: record.boundary, transactionId: record.transactionId, digest: record.digest,
      owner: "0aaaaaaa-0000-4000-8000-000000000001", fence: "0fffffff-0000-4000-8000-000000000001",
      state: "in_flight", createdAt: T_BASE + 60_000,
    }]));
    vi.setSystemTime(T_BASE + 90_000);
    const orphan = await deliver(orphanWorld, orphanPane, turn(2), X);
    expect(orphanWorld.claims.at(-1)).toBe("restart_rejected");
    expectNotExecutedPayload(orphanWorld, orphan, NOT_STARTED_PAYLOAD);
    expect(orphanWorld.mutations).toBe(1);

    vi.setSystemTime(T_BASE);
    const conflictWorld = world();
    const conflictPane = await pane(conflictWorld);
    expectExecutedCommitted(conflictWorld, await deliver(conflictWorld, conflictPane, turn(1), X));
    vi.setSystemTime(T_BASE + 90_000);
    const changed = await deliver(conflictWorld, conflictPane, turn(2), X, CHANGED_PLAN);
    expect(conflictWorld.claims.at(-1)).toBe("conflict");
    expectNotExecutedPayload(conflictWorld, changed, NOT_STARTED_PAYLOAD);
    expect(conflictWorld.mutations).toBe(1);

    vi.setSystemTime(T_BASE);
    const corruptWorld = world();
    corruptWorld.storage.setItem(STORAGE_KEY, "{not json");
    const corruptPane = await pane(corruptWorld);
    const refused = await deliver(corruptWorld, corruptPane, turn(1), X);
    expect(corruptWorld.claims).toEqual(["persistence_unavailable"]);
    expectNotExecutedPayload(corruptWorld, refused, NOT_STARTED_PAYLOAD);
    expect(corruptWorld.mutations).toBe(0);
  });
});

describe("R4e GA3 deadline and direct-route pins", () => {
  it("RT5 the 90 s turn deadline passes inside the execute-start acknowledgement: zero mutation, executed:true with the exact uncertain shape", async () => {
    const w = world();
    const p = await pane(w);
    w.companion.onExecuteStart = () => { vi.setSystemTime(T_BASE + 90_000); };
    const delivery = await deliver(w, p, turn(1), X);
    expect(w.companion.executeStartsFor(delivery.ticketId)).toBe(1);
    expect(w.mutations, diag(w)).toBe(0);
    expect(w.cells.size).toBe(0);
    expect(delivery.posted, diag(w)).toHaveLength(1);
    expect(delivery.posted[0], diag(w)).toEqual({ ticketId: delivery.ticketId, executed: true, result: UNCERTAIN_PAYLOAD });
    expect(w.companion.tickets[0]).toMatchObject({ state: "terminal", terminalReason: "completed" });
    expect(storedReports(w)).toMatchObject([{ outcome: "cancelled" }]);
  });

  it("RT7 [pin] the direct-route shape (no beforeFirstMutation) hands the executor the bare signal and never refuses on the deadline", async () => {
    const turnId = turn(7);
    const controller = new AbortController();
    const trustedRequest = classifyWorkbookRequest(PROMPT);
    const turnStart = Date.now();
    const signals: (AbortSignal | undefined)[] = [];
    let mutations = 0;
    const adapter: WorkbookPlanExecutionAdapter & WorkbookPlanVerificationAdapter = {
      async executeFirstGroup(gate, signal) {
        signals.push(signal);
        vi.setSystemTime(turnStart + 90_001);
        return gate.executeAfterFresh("fresh", async () => { mutations += 1; return { status: "committed" }; });
      },
      async executeGroup() { mutations += 1; return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch" as const); },
      async verifyPlan() { return true; },
    };
    const planResponse = {
      content: "",
      toolCalls: [{ id: "call_r4e_direct", type: "function" as const, function: { name: "excel_execute_plan", arguments: JSON.stringify(PLAN) } }],
    };
    const dependencies: Parameters<typeof createWorkbookPlanTurn>[0] = {
      identity: { sessionId: DIRECT_SESSION, documentId: DOC, turnId },
      trustedOperationBoundary: createTrustedOperationBoundary({ sessionId: DIRECT_SESSION, documentId: DOC, turnId }),
      getCompileContextForPlan: async () => RUNTIME.compileContext,
      policyInput: (plan) => buildWorkbookPlanPolicyInput(plan, RUNTIME, trustedRequest),
      resourcePreflight: (plan) => preflightWorkbookPlanResources(plan, EMPTY_RESOURCE_CONTEXT),
      resourcePreflightAfterRuntime: (plan) => preflightWorkbookPlanResources(plan, RUNTIME),
      store: new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      model: async () => planResponse,
      initialResponse: planResponse,
      inspect: async () => "{}",
      adapter,
      signal: controller.signal,
      turnStart,
      rolloutEnabled: true,
      route: "direct",
      trustedTurnOperationId: turnId,
    };
    const result = await createWorkbookPlanTurn(dependencies);
    expect(signals).toHaveLength(1);
    expect(signals[0]).toBe(controller.signal);
    expect(mutations).toBe(1);
    expect(result.execution?.outcome).toBe("committed");
  });
});

describe("R4e test-only rows folded from the R4d reviews (release record rows a and b)", () => {
  const taskpaneSource = (): string => readFileSync(resolve(__dirname, "../taskpane/taskpane.ts"), "utf8");

  it("W1 (row a, source pin) the server primitive gate reads the bridge lifecycle session and the turn's recorded session", () => {
    const source = taskpaneSource();
    const start = source.indexOf("const serverPrimitiveGate = createServerPrimitiveGate({");
    expect(start).toBeGreaterThan(-1);
    const block = source.slice(start, source.indexOf("});", start) + 3);
    expect(block.split("\n")).toEqual([
      "const serverPrimitiveGate = createServerPrimitiveGate({",
      "  getBridgeSessionId: () => bridgeLifecycle?.sessionId ?? null,",
      "  getActiveTurn: () => ({ turnOperationId: activeTurnOperationId, sessionId: activeTurnSessionId }),",
      "  resolveAuthority: (input) => turnCoordinator.serverPrimitiveAuthority(input),",
      "  executeWithAuthority: executeValidatedExcelBridgeToolWithAuthority,",
      "  executeReadOnly: executeValidatedExcelBridgeTool,",
      "});",
    ]);
    // The recorded turn session is written only at turn begin, at turn preparation, and cleared at turn end.
    const writes = source.split("\n").map((line) => line.trim()).filter((line) => /^activeTurnSessionId = /.test(line));
    expect(writes).toEqual([
      "activeTurnSessionId = bridgeLifecycle.sessionId;",
      "activeTurnSessionId = preparation.sessionId;",
      "activeTurnSessionId = null;",
      "activeTurnSessionId = null;",
    ]);
  });

  it("W2 (row a, behaviour) with the taskpane wiring a rollout-off turn begun on S1 is denied once the bridge moved to S2", async () => {
    const identity = { sessionId: SESSION_1 as string | null, documentId: DOC as string | null };
    const coordinator = new TaskPaneTurnCoordinator({
      getOrCreateDocumentId: async () => DOC,
      getBridgeIdentity: () => ({ ...identity }),
      reconnectForDocumentChange: async () => { throw new Error("unexpected document change"); },
      bindings: new ServerOwnedTurnBindingStore({ now: () => Date.now(), ttlMs: 90_000, maxEntries: 32 }),
      paneId: PANE,
    });
    const controller = new AbortController();
    const turnId = turn(21);
    coordinator.beginTurn(turnId, controller, {
      trustedRequest: classifyWorkbookRequest(PROMPT), turnStart: Date.now(), cancellationSignal: controller.signal, workbookPlanRolloutEnabled: false,
    });
    const activeTurnSessionId = identity.sessionId; // taskpane.ts: recorded when the turn begins
    const withAuthority: string[] = [];
    const readOnly: string[] = [];
    const wiring = (sessionOf: () => string | null) => createServerPrimitiveGate({
      getBridgeSessionId: () => identity.sessionId,
      getActiveTurn: () => ({ turnOperationId: turnId, sessionId: sessionOf() }),
      resolveAuthority: (input) => coordinator.serverPrimitiveAuthority(input),
      executeWithAuthority: async (name) => { withAuthority.push(name); return JSON.stringify({ ok: true }); },
      executeReadOnly: async (name) => { readOnly.push(name); return JSON.stringify({ ok: true }); },
    });
    const gate = wiring(() => activeTurnSessionId);
    const args = { sheet: "Sheet1", address: "A1", values: [["ok"]] };
    expect(await gate("write_range", args)).toBe(JSON.stringify({ ok: true }));
    expect(withAuthority).toEqual(["write_range"]);
    identity.sessionId = SESSION_2; // 401 auto-reconnect mid-turn
    const denied = JSON.parse(await gate("write_range", args)) as { ok: boolean; error: string };
    expect(denied.ok).toBe(false);
    expect(denied.error).toBe(PRIMITIVE_DENIAL_MESSAGES.TURN_INACTIVE);
    expect(withAuthority).toEqual(["write_range"]);
    expect(readOnly).toEqual([]);
    // Contrast (why W1 pins the wiring): reporting the live bridge session as the turn's session lets it through.
    const liveSessionWiring = wiring(() => identity.sessionId);
    await liveSessionWiring("write_range", args);
    expect(withAuthority).toEqual(["write_range", "write_range"]);
  });

  it("D1 (row b) serverPrimitiveAuthority denies a prepared rollout-on turn when the document drifts on the same session", async () => {
    const identity = { sessionId: SESSION_1 as string | null, documentId: DOC as string | null };
    const coordinator = new TaskPaneTurnCoordinator({
      getOrCreateDocumentId: async () => DOC,
      getBridgeIdentity: () => ({ ...identity }),
      reconnectForDocumentChange: async () => { throw new Error("unexpected document change"); },
      bindings: new ServerOwnedTurnBindingStore({ now: () => Date.now(), ttlMs: 90_000, maxEntries: 32 }),
      paneId: PANE,
    });
    const controller = new AbortController();
    const turnId = turn(22);
    const handle = coordinator.beginTurn(turnId, controller, {
      trustedRequest: classifyWorkbookRequest(PROMPT), turnStart: Date.now(), cancellationSignal: controller.signal, workbookPlanRolloutEnabled: true,
    });
    await expect(coordinator.prepareTurn(handle)).resolves.toMatchObject({ status: "active", sessionId: SESSION_1, documentId: DOC });
    const authority = coordinator.serverPrimitiveAuthority({ turnId, bridgeSessionId: SESSION_1 });
    expect(authority).not.toBeNull();
    identity.documentId = DOC_2;
    expect(coordinator.serverPrimitiveAuthority({ turnId, bridgeSessionId: SESSION_1 })).toBeNull();
    identity.documentId = DOC;
    expect(coordinator.serverPrimitiveAuthority({ turnId, bridgeSessionId: SESSION_1 })).toBe(authority);
  });
});
