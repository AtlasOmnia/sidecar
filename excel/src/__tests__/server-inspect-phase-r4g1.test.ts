import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { describe, expect, it, vi } from "vitest";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import { OfficeCompanionBridge, type BridgePlanExecutionControl } from "../officeBridge";
import { PRIMITIVE_DENIAL_MESSAGES } from "../primitiveMutationAuthority";
import * as routeModule from "../taskpane/serverOwnedPlanRoute";
import { TaskPaneTurnCoordinator } from "../taskpane/taskPaneTurnCoordinator";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import {
  BrowserIdempotencyStore,
  createTrustedOperationBoundary,
  MemoryIdempotencyStorage,
  type WorkbookPlanExecutionAdapter,
} from "../workbookPlanExecutor";
import {
  createWorkbookPlanTurn,
  ServerOwnedTurnBindingStore,
  WorkbookPlanTurnState,
  type ServerOwnedTicketIdentity,
  type ServerOwnedTurnBindingResult,
  type WorkbookPlanConfirmationBinding,
  type WorkbookPlanModelResponse,
  type WorkbookPlanTurnDependencies,
  type WorkbookPlanVerificationAdapter,
} from "../workbookPlanOrchestrator";
import { classifyWorkbookRequest } from "../workbookPlanRequest";
import type { WorkbookPlanRouteRuntimeContext } from "../workbookPlanRoute";

/**
 * R4g-1 (dev3-r4g1-release.md): the pane route, orchestrator, binding store and coordinator accept a
 * trusted server inspection phase (one bounded read, then one execute on the same pane turn state)
 * behind SERVER_INSPECTION_SUPPORTED, which stays false. Real chain: createServerOwnedPlanRoute ->
 * TaskPaneTurnCoordinator -> ServerOwnedTurnBindingStore -> createWorkbookPlanTurn ->
 * executeAcceptedWorkbookPlan; Office, adapter, store and inspect are counting doubles. The harness
 * shape is COPIED from workbook-plan-production-route.test.ts (blob 4dcc873297be) and
 * task-pane-lifecycle-r4f.test.ts (neither is imported nor edited); R15/R15b use the real
 * OfficeCompanionBridge. RED rows R1-R15b, parity rows L1-L3 (identical at base and HEAD), GREEN rows
 * S1-S14. No real-time waits, no fake timers; the store clock is a plain counter.
 */

type RouteDependencies = Parameters<typeof routeModule.createServerOwnedPlanRoute>[0];

const SESSION = "aaaaaaaa-aaaa-4aaa-8aaa-000000000001";
const DOC = "d0c00000-0000-4000-8000-000000000001";
const PANE = "9a9e0000-0000-4000-8000-000000000002";
const PANE_2 = "9a9e0000-0000-4000-8000-000000000003";
const TURN = "7a7a0000-0000-4000-8000-000000000004";
const OP1 = "5e7e0000-0000-4000-8000-000000000011";
const OP2 = "5e7e0000-0000-4000-8000-000000000012";
const OP3 = "5e7e0000-0000-4000-8000-000000000013";
const K = "4b4b0000-0000-4000-8000-000000000021";
const K2 = "4b4b0000-0000-4000-8000-000000000022";
const EXPLICIT_PROMPT = "write ok into A1";
/** Classifies as ambiguous_mutation: the policy requires pane confirmation (copied from task-pane-session-rotation-r3b). */
const AMBIGUOUS_PROMPT = "do the usual thing";
const TTL_MS = 90_000;
/** The binding store's fixed refusal for a phase-fence violation (R4g-1). */
const PHASE_REFUSAL = { ok: false, reason: "Trusted server turn phase is not permitted." };

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
const INSPECT_PLAN = { version: 1, outcome: "inspect", reason: "look", reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1" }] } as const;
/** 100 x 100 = 10,000 cells: over the 5,000-cell inspection read budget. */
const OVERSIZED_INSPECT_PLAN = { version: 1, outcome: "inspect", reason: "look", reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1:CV100" }] } as const;
const ASK_PLAN = { version: 1, outcome: "ask", message: "Which sheet?" } as const;
const REFUSE_PLAN = { version: 1, outcome: "refuse", message: "Not permitted." } as const;
const INSPECTION_READS = [{ ok: true, data: [["v"]] }];
const INSPECT_RESULT = JSON.stringify({ ok: true, reads: INSPECTION_READS });

const SUMMARY = {
  rejected: "The workbook plan was rejected before execution; nothing was executed.",
  cancelled: "The workbook plan was cancelled before execution; nothing was executed.",
  invalid: "The workbook plan ticket was invalid; nothing was executed.",
} as const;
const notExecutedBody = (reason: keyof typeof SUMMARY): string =>
  JSON.stringify({ ok: false, action: "workbook_plan", reason, summary: SUMMARY[reason] });
const INSPECTION_BODY = JSON.stringify({
  ok: true,
  action: "workbook_plan",
  phase: "inspect",
  result: { outcome: "inspected", inspectionCount: 1, reads: INSPECTION_READS },
});

type Phase = "inspect" | "execute";
interface Envelope {
  readonly turnOperationId: string;
  readonly workbookPlanFingerprint: string;
  readonly phase?: string;
  readonly turnKey?: string;
}
const envelope = (op: string, phase?: Phase | string, turnKey?: string): Envelope => ({
  turnOperationId: op,
  workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
  ...(phase === undefined ? {} : { phase }),
  ...(turnKey === undefined ? {} : { turnKey }),
});

/** Stateful Office adapter double (COPIED shape from workbook-plan-production-route.test.ts). */
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

interface HarnessOptions {
  /** true: the test-only override answers true; false: the override answers false; undefined: no override (constant). */
  readonly gate?: boolean;
  readonly inspect?: (reads: unknown, signal?: AbortSignal) => Promise<string>;
  readonly prompt?: string;
  readonly maxEntries?: number;
  readonly awaitPaneConfirmation?: (coordinator: TaskPaneTurnCoordinator) => RouteDependencies["awaitPaneConfirmation"];
  readonly createStore?: () => BrowserIdempotencyStore;
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

async function harness(options: HarnessOptions = {}) {
  const events: string[] = [];
  const clock = { t: 1_000_000 };
  const bindings = new ServerOwnedTurnBindingStore({ now: () => clock.t, ttlMs: TTL_MS, maxEntries: options.maxEntries ?? 32, maxBytes: 64 * 1024 });
  const bindSpy = vi.spyOn(bindings, "bindServerTicket");
  const coordinator = new TaskPaneTurnCoordinator({
    getOrCreateDocumentId: async () => DOC,
    getBridgeIdentity: () => ({ sessionId: SESSION, documentId: DOC }),
    reconnectForDocumentChange: async () => { throw new Error("unexpected document change"); },
    bindings,
    paneId: PANE,
  });
  const controller = new AbortController();
  const trustedRequest = classifyWorkbookRequest(options.prompt ?? EXPLICIT_PROMPT);
  const turnStart = Date.now();
  const handle = coordinator.beginTurn(TURN, controller, { trustedRequest, turnStart, cancellationSignal: controller.signal, workbookPlanRolloutEnabled: true });
  await expect(coordinator.prepareTurn(handle)).resolves.toMatchObject({ status: "active" });
  const { adapter, cells } = statefulAdapter(events);
  const store = new BrowserIdempotencyStore(new MemoryIdempotencyStorage());
  const claim = store.claim.bind(store);
  store.claim = async (...args: Parameters<typeof store.claim>) => {
    events.push("store-claim");
    return claim(...args);
  };
  const inspect = vi.fn(async (reads: unknown, signal?: AbortSignal) => {
    events.push("inspect-read");
    return options.inspect === undefined ? INSPECT_RESULT : options.inspect(reads, signal);
  });
  const dependencies: Record<string, unknown> = {
    isRolloutEnabled: () => true,
    getBridgeSessionId: () => SESSION,
    getActiveTurn: () => ({ turnOperationId: TURN, trustedRequest, sessionId: SESSION, turnStart, hasCancellation: true }),
    coordinator,
    bindings,
    readRuntimeContext: async () => { events.push("runtime-read"); return RUNTIME; },
    createAdapter: () => adapter,
    createStore: options.createStore ?? (() => store),
    inspect,
    awaitPaneConfirmation: options.awaitPaneConfirmation?.(coordinator)
      ?? ((binding: WorkbookPlanConfirmationBinding, signal?: AbortSignal) => coordinator.waitForWorkbookPlanConfirmation(binding, signal)),
    executePrimitive: async (name: string) => { events.push(`primitive:${name}`); return JSON.stringify({ ok: true }); },
    ...(options.gate === undefined ? {} : { serverInspectionSupported: () => options.gate }),
  };
  const route = routeModule.createServerOwnedPlanRoute(dependencies as unknown as RouteDependencies);
  const acks = { count: 0 };
  const control = (): BridgePlanExecutionControl => ({
    acknowledgeExecuteStart: async () => { acks.count += 1; events.push("ack"); },
  });
  const run = (plan: Record<string, unknown>, ticket: Envelope): Promise<string> =>
    route.executeServerOwnedWorkbookPlan({ ...plan }, ticket as never, control());
  const bound = (index: number): ServerOwnedTurnBindingResult => bindSpy.mock.results[index]?.value as ServerOwnedTurnBindingResult;
  const stateOf = (index: number): WorkbookPlanTurnState | undefined => {
    const result = bound(index);
    return result?.ok === true ? result.binding.state : undefined;
  };
  const count = (event: string): number => events.filter((entry) => entry === event).length;
  return { events, clock, bindings, bindSpy, coordinator, controller, route, cells, inspect, acks, run, bound, stateOf, count, store, dependencies };
}

const WORKBOOK_TOUCHES = ["runtime-read", "store-claim", "fresh", "mutation", "verify-read", "ack"];
function expectNoWriteAccess(h: Awaited<ReturnType<typeof harness>>): void {
  for (const touch of WORKBOOK_TOUCHES) expect(h.count(touch), touch).toBe(0);
  expect(h.cells.size).toBe(0);
}

// ---------------------------------------------------------------------------
// 5.1 RED rows (fail at the base 2c64ebd; gate override ON unless stated)
// ---------------------------------------------------------------------------
describe("R4g-1 RED rows: one bounded server inspection, then one execute on the shared turn state", () => {
  it("R1 inspect op (phase inspect, K) with a valid inspect plan: one inspect call, the §4 body byte-exact, state inspected, inspect entry terminal, no write access", async () => {
    const h = await harness({ gate: true });
    const body = await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K));
    expect(body).toBe(INSPECTION_BODY);
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(h.stateOf(0)?.phase).toBe("inspected");
    expectNoWriteAccess(h);
    const identity = h.bindSpy.mock.calls[0]![0];
    expect(h.bindings.bindServerTicket(identity)).toEqual({ ok: false, reason: "Trusted server turn is terminal." });
  });

  it("R2 R1 then execute op (new op id, phase execute, same K): commits once on the SAME state object, beginExecution from inspected, one ack", async () => {
    const h = await harness({ gate: true });
    expect(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(INSPECTION_BODY);
    const shared = h.stateOf(0)!;
    const seen: string[] = [];
    const begin = shared.beginExecution.bind(shared);
    shared.beginExecution = () => { seen.push(shared.phase); return begin(); };
    const payload = JSON.parse(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K))) as Record<string, unknown>;
    expect(h.bound(1)?.ok).toBe(true);
    expect(h.stateOf(1)).toBe(shared);
    expect(seen).toEqual(["inspected"]);
    expect(payload).toMatchObject({ ok: true, action: "workbook_plan", result: { outcome: "committed", verification: "verified" } });
    expect(payload).not.toHaveProperty("reason");
    expect(h.acks.count).toBe(1);
    expect(h.count("mutation")).toBe(1);
    expect(h.cells.get("Sheet1!A1")).toEqual([["ok"]]);
    expect(h.inspect).toHaveBeenCalledTimes(1);
  });

  it("R3 inspect then a second inspect op (new id, same K): the second is rejected, zero further reads", async () => {
    const h = await harness({ gate: true });
    expect(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(INSPECTION_BODY);
    expect(await h.run(INSPECT_PLAN, envelope(OP2, "inspect", K))).toBe(notExecutedBody("rejected"));
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(h.bound(1)).toEqual(PHASE_REFUSAL);
  });

  it("R4 execute op (K) then inspect op (K): the inspect is refused at bind, zero reads (no plan op after execute)", async () => {
    const h = await harness({ gate: true });
    expect(JSON.parse(await h.run(EXECUTE_PLAN, envelope(OP1, "execute", K)))).toMatchObject({ ok: true, result: { outcome: "committed" } });
    expect(await h.run(INSPECT_PLAN, envelope(OP2, "inspect", K))).toBe(notExecutedBody("rejected"));
    expect(h.inspect).not.toHaveBeenCalled();
    expect(h.bound(1)).toEqual(PHASE_REFUSAL);
  });

  it("R5 execute then execute (two ids, K): the second is rejected, zero reads, zero further acks", async () => {
    const h = await harness({ gate: true });
    expect(JSON.parse(await h.run(EXECUTE_PLAN, envelope(OP1, "execute", K)))).toMatchObject({ ok: true });
    const reads = h.count("runtime-read");
    expect(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K))).toBe(notExecutedBody("rejected"));
    expect(h.count("runtime-read")).toBe(reads);
    expect(h.acks.count).toBe(1);
    expect(h.count("mutation")).toBe(1);
    expect(h.bound(1)).toEqual(PHASE_REFUSAL);
  });

  it("R6 inspect (K) then execute (K2 != K): the execute is rejected, zero reads (the turnKey is bound to the pane turn)", async () => {
    const h = await harness({ gate: true });
    expect(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(INSPECTION_BODY);
    expect(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K2))).toBe(notExecutedBody("rejected"));
    expectNoWriteAccess(h);
    expect(h.bound(1)).toEqual(PHASE_REFUSAL);
    expect(h.stateOf(0)?.phase).toBe("inspected");
  });

  it("R7 Stop (abort the turn controller) during the inspect call: cancelled, the shared state terminal, no inspection body", async () => {
    const gate = deferred<string>();
    const reached = deferred<void>();
    const h = await harness({ gate: true, inspect: async () => { reached.resolve(); return gate.promise; } });
    const pending = h.run(INSPECT_PLAN, envelope(OP1, "inspect", K));
    await Promise.race([reached.promise, pending.then(() => undefined)]);
    expect(h.inspect).toHaveBeenCalledTimes(1);
    h.controller.abort();
    expect(await pending).toBe(notExecutedBody("cancelled"));
    gate.resolve(INSPECT_RESULT);
    expect(h.stateOf(0)?.phase).toBe("terminal");
    expect(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K))).toBe(notExecutedBody("rejected"));
    expectNoWriteAccess(h);
  });

  it("R8 inspect() throws: rejected, the shared state terminal, the following execute (same K) refused with zero reads", async () => {
    const h = await harness({ gate: true, inspect: async () => { throw new Error("INSPECT_SECRET_CANARY_R4G1"); } });
    const body = await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K));
    expect(body).toBe(notExecutedBody("rejected"));
    expect(body).not.toContain("CANARY");
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(h.stateOf(0)?.phase).toBe("terminal");
    expect(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K))).toBe(notExecutedBody("rejected"));
    expect(h.bound(1)).toEqual(PHASE_REFUSAL);
    expectNoWriteAccess(h);
  });

  it("R9 inspect resource preflight fails (10,000 cells > 5,000): rejected, zero inspect calls, the shared state terminal", async () => {
    const h = await harness({ gate: true });
    expect(await h.run(OVERSIZED_INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(notExecutedBody("rejected"));
    expect(h.inspect).not.toHaveBeenCalled();
    expect(h.stateOf(0)?.phase).toBe("terminal");
    expect(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K))).toBe(notExecutedBody("rejected"));
    expect(h.bound(1)).toEqual(PHASE_REFUSAL);
    expectNoWriteAccess(h);
  });

  it("R10 replay: the same inspect op id delivered twice: the second is refused at bind, inspect stays at one call", async () => {
    const h = await harness({ gate: true });
    expect(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(INSPECTION_BODY);
    expect(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(notExecutedBody("rejected"));
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(h.bound(1)).toEqual({ ok: false, reason: "Trusted server turn is terminal." });
  });

  it("R11 replay: the same execute op id after the commit is refused, zero further mutation", async () => {
    const h = await harness({ gate: true });
    expect(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(INSPECTION_BODY);
    expect(JSON.parse(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K)))).toMatchObject({ ok: true });
    expect(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K))).toBe(notExecutedBody("rejected"));
    expect(h.count("mutation")).toBe(1);
    expect(h.acks.count).toBe(1);
    expect(h.bound(2)).toEqual({ ok: false, reason: "Trusted server turn is terminal." });
  });

  it("R12 session invalidation between inspect and execute: the execute is rejected with zero reads", async () => {
    const h = await harness({ gate: true });
    expect(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(INSPECTION_BODY);
    h.bindings.invalidateSession(SESSION);
    expect(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K))).toBe(notExecutedBody("rejected"));
    expect(h.bound(1)).toEqual({ ok: false, reason: "Trusted server turn identity is unavailable." });
    expectNoWriteAccess(h);
  });

  it("R13 TTL prune between inspect and execute (store clock past ttlMs, pane turn active): the execute is refused (identity unavailable), never a fresh none state", async () => {
    const h = await harness({ gate: true });
    expect(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(INSPECTION_BODY);
    h.clock.t += TTL_MS;
    expect(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K))).toBe(notExecutedBody("rejected"));
    expect(h.bound(1)).toEqual({ ok: false, reason: "Trusted server turn identity is unavailable." });
    expectNoWriteAccess(h);
  });

  it("R13b timing control: just before ttlMs the execute succeeds on the shared inspected state", async () => {
    const h = await harness({ gate: true });
    expect(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(INSPECTION_BODY);
    h.clock.t += TTL_MS - 1;
    expect(JSON.parse(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K)))).toMatchObject({ ok: true, result: { outcome: "committed" } });
    expect(h.stateOf(1)).toBe(h.stateOf(0));
  });

  it("R13c a failed inspect, then a TTL prune, then execute: refused, zero reads", async () => {
    const h = await harness({ gate: true, inspect: async () => { throw new Error("boom"); } });
    expect(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(notExecutedBody("rejected"));
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(h.stateOf(0)?.phase).toBe("terminal");
    h.clock.t += TTL_MS;
    expect(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K))).toBe(notExecutedBody("rejected"));
    expect(h.bound(1)).toEqual({ ok: false, reason: "Trusted server turn identity is unavailable." });
    expectNoWriteAccess(h);
  });

  it.each([
    ["non-JSON", "not json"],
    ["no reads array", JSON.stringify({ ok: true })],
    ["reads not an array", JSON.stringify({ ok: true, reads: { a: 1 } })],
    ["JSON array", JSON.stringify([1])],
  ])("R14 inspect() returns %s: rejected, the shared state terminal, the following execute refused", async (_label, raw) => {
    const h = await harness({ gate: true, inspect: async () => raw });
    expect(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(notExecutedBody("rejected"));
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(h.stateOf(0)?.phase).toBe("terminal");
    expect(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K))).toBe(notExecutedBody("rejected"));
    expect(h.bound(1)).toEqual(PHASE_REFUSAL);
    expectNoWriteAccess(h);
  });
});

// ---------------------------------------------------------------------------
// R15/R15b: the real OfficeCompanionBridge posts the route's inspection body executed:false
// ---------------------------------------------------------------------------
const BRIDGE_SESSION = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const BRIDGE_SESSION_2 = "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee";
const BRIDGE_PANE = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
const BRIDGE_TICKET = "cccccccc-cccc-4ccc-8ccc-cccccccccccc";
const BRIDGE_REQUEST = "dddddddd-dddd-4ddd-8ddd-dddddddddddd";

function response(status: number, body?: unknown): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
  });
}

/** Session-aware fetch double (shape COPIED from office-bridge-session-scope-r4j.test.ts); unknown routes fail loudly. */
function bridgeFetch(resultStatuses: number[]) {
  const posts: { session: string; body: Record<string, unknown> }[] = [];
  let sessions = 0;
  let executeStarts = 0;
  let pendingServed = false;
  const fetchImpl = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    const method = init?.method ?? "GET";
    const session = (init?.headers as Record<string, string> | undefined)?.["X-Hermes-Session-Id"] ?? "-";
    if (method === "POST" && url === "/excel/session") {
      sessions += 1;
      return response(201, {
        ok: true,
        sessionId: sessions === 1 ? BRIDGE_SESSION : BRIDGE_SESSION_2,
        sessionKey: (sessions === 1 ? "s" : "t").repeat(64),
        protocolVersion: "1.0", expiresInSeconds: 60, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
      });
    }
    if (method === "DELETE" && url === "/excel/session") return response(200, { ok: true });
    if (method === "POST" && url === "/excel/heartbeat") return response(200, { ok: true });
    if (method === "GET" && url === "/excel/pending") {
      if (pendingServed) return response(204);
      pendingServed = true;
      return response(200, {
        ticketId: BRIDGE_TICKET, requestId: BRIDGE_REQUEST, tool: "excel_execute_plan", arguments: { ...INSPECT_PLAN },
        trusted: { turnOperationId: BRIDGE_REQUEST, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT, phase: "inspect", turnKey: K },
      });
    }
    if (method === "POST" && url === "/excel/execute-start") { executeStarts += 1; return response(200, { ok: true }); }
    if (method === "POST" && url === "/excel/result") {
      posts.push({ session, body: JSON.parse(String(init?.body)) as Record<string, unknown> });
      const status = resultStatuses.shift() ?? 200;
      return status === 200 ? response(200, { ok: true }) : response(status, { error: "too large" });
    }
    throw new Error(`unexpected fetch ${method} ${url}`);
  });
  return { fetchImpl: fetchImpl as unknown as typeof fetch, posts, executeStarts: () => executeStarts };
}

async function routeInspectionBody(): Promise<string> {
  const h = await harness({ gate: true });
  return h.run(INSPECT_PLAN, envelope(OP1, "inspect", K));
}

async function settle(promise: Promise<unknown>): Promise<string> {
  return promise.then((value) => `resolved:${String(value)}`, (error: unknown) => `rejected:${error instanceof Error ? error.name : "?"}`);
}

describe("R4g-1 R15/R15b: the real bridge carries the inspection body executed:false and never upgrades it", () => {
  it("R15 the inspection post is answered 413: the retry posts the fixed rejected substitute executed:false, zero acks", async () => {
    const body = await routeInspectionBody();
    const f = bridgeFetch([413, 200]);
    const executeTool = vi.fn(async () => body);
    const bridge = new OfficeCompanionBridge({ hostPrefix: "excel", documentId: async () => "doc-1", paneId: BRIDGE_PANE, executeTool, fetchImpl: f.fetchImpl });
    await bridge.connect();
    expect(await settle(bridge.pollOnce())).toBe("rejected:BridgeHttpError");
    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    expect(f.posts.map((post) => Object.keys(post.body))).toEqual([["ticketId", "executed", "result"], ["ticketId", "executed", "result"]]);
    expect(f.posts[0]!.body).toEqual({ ticketId: BRIDGE_TICKET, executed: false, result: JSON.parse(INSPECTION_BODY) });
    expect(f.posts[1]!.body.executed).toBe(false);
    expect(f.posts[1]!.body.result).toMatchObject({ ok: false, action: "workbook_plan", category: "Execution", reason: "rejected" });
    expect(f.posts[1]!.body.result).not.toHaveProperty("state");
    expect(f.executeStarts()).toBe(0);
    expect(executeTool).toHaveBeenCalledTimes(1);
  });

  it("R15b after the 413 the bridge session is replaced before the retry: the cached inspection entry is dropped unposted, never sent with the new session", async () => {
    const body = await routeInspectionBody();
    const f = bridgeFetch([413]);
    const executeTool = vi.fn(async () => body);
    const handlers: (() => void)[] = [];
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "excel", documentId: async () => "doc-1", paneId: BRIDGE_PANE, executeTool, fetchImpl: f.fetchImpl,
      setIntervalImpl: ((handler: () => void) => { handlers.push(handler); return handlers.length; }) as unknown as typeof setInterval,
      clearIntervalImpl: (() => undefined) as unknown as typeof clearInterval,
    });
    await bridge.start();
    expect(await settle(bridge.pollOnce())).toBe("rejected:BridgeHttpError");
    expect(f.posts[0]!.body).toEqual({ ticketId: BRIDGE_TICKET, executed: false, result: JSON.parse(INSPECTION_BODY) });
    await bridge.stop();
    await bridge.start();
    expect(bridge.sessionId).toBe(BRIDGE_SESSION_2);
    expect(await settle(bridge.pollOnce())).toBe("resolved:false");
    expect(f.posts).toHaveLength(1);
    expect(f.posts.filter((post) => post.session === BRIDGE_SESSION_2)).toEqual([]);
    expect(f.executeStarts()).toBe(0);
    await bridge.stop();
  });
});

// ---------------------------------------------------------------------------
// 5.2 parity rows (identical assertions at the base and at HEAD)
// ---------------------------------------------------------------------------
const UUID_V4 = /[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}/gi;
const KNOWN = new Map<string, string>([[SESSION, "<S>"], [DOC, "<DOC>"], [PANE, "<PANE>"], [TURN, "<TURN>"], [OP1, "<OP1>"], [OP2, "<OP2>"]]);
function normalize(text: string): string {
  return text.replace(UUID_V4, (match) => KNOWN.get(match.toLowerCase()) ?? "<uuid>").replace(/\b\d{10,}\b/g, "<n>").replace(/"[0-9a-f]{64}"/g, "\"<hex64>\"");
}

/** L1: legacy (no-phase) ticket through the real route; the log is printed for the base/HEAD diff and pinned literally. */
async function legacyLog(plan: Record<string, unknown>, gate?: boolean): Promise<string[]> {
  const h = await harness(gate === undefined ? {} : { gate });
  const payload = await h.run(plan, envelope(OP1));
  const identityKeys = h.bindSpy.mock.calls.map((call) => Object.keys(call[0]).join(","));
  const log = [
    `events=${h.events.join(",")}`,
    `inspect=${h.inspect.mock.calls.length} acks=${h.acks.count} binds=${h.bindSpy.mock.calls.length} bindOk=${h.bindSpy.mock.results.map((r) => (r.value as ServerOwnedTurnBindingResult).ok).join(",")}`,
    `identityKeys=${identityKeys.join("|")}`,
    `state=${h.stateOf(0)?.phase ?? "-"}`,
    `payload=${normalize(payload)}`,
  ];
  // eslint-disable-next-line no-console
  console.log(`L1-LOG ${String(plan.outcome)} ${JSON.stringify(log)}`);
  return log;
}

const L1_IDENTITY = "identityKeys=sessionId,documentId,paneId,turnId,serverOperationId";
/** Recorded from the base 2c64ebd run of this row (r4g1/red-phase.log L1-LOG execute); the ids are deterministic (UUIDv5). */
const L1_EXECUTE_PAYLOAD = '{"ok":true,"action":"workbook_plan","result":{"outcome":"committed","finalSummary":"Verified 1 workbook operation completed successfully.","inspectionCount":0,"correctionCount":0,"execution":{"outcome":"committed","transactionId":"af1553e6-b39b-5b5b-b377-4352e49b2c42","operationCount":1,"operations":[{"operationId":"d3524b30-5319-55e1-b219-80747743d80a","operationIndex":0,"effect":"write","groupIndex":0,"state":"committed","status":"applied"}],"verification":"not_attempted","rollback":"not_attempted"},"verification":"verified"}}';
describe("R4g-1 parity rows (base and HEAD): legacy no-phase tickets are unchanged", () => {
  it("L1 execute write: identical event order, store writes, identity key set and payload shape", async () => {
    const log = await legacyLog(EXECUTE_PLAN);
    expect(log).toEqual([
      "events=runtime-read,store-claim,fresh,ack,mutation,verify-read",
      "inspect=0 acks=1 binds=1 bindOk=true",
      L1_IDENTITY,
      "state=terminal",
      `payload=${L1_EXECUTE_PAYLOAD}`,
    ]);
  });

  it.each([["ask", ASK_PLAN], ["refuse", REFUSE_PLAN]] as const)("L1 %s: identical call log and the plan message payload", async (reason, plan) => {
    const log = await legacyLog(plan);
    expect(log).toEqual([
      "events=",
      "inspect=0 acks=0 binds=1 bindOk=true",
      L1_IDENTITY,
      "state=terminal",
      `payload=${JSON.stringify({ ok: false, action: "workbook_plan", reason, summary: reason === "ask" ? "The workbook plan asked for clarification; nothing was executed." : "The workbook plan was refused; nothing was executed.", result: { outcome: reason, finalSummary: plan.message } })}`,
    ]);
  });

  it("L1/L2 legacy inspect ticket: rejected T2 shape, zero inspect calls (gate absent, and with the test override on: S4)", async () => {
    for (const gate of [undefined, true]) {
      const log = await legacyLog(INSPECT_PLAN, gate);
      expect(log, String(gate)).toEqual([
        "events=",
        "inspect=0 acks=0 binds=1 bindOk=true",
        L1_IDENTITY,
        "state=terminal",
        `payload=${notExecutedBody("rejected")}`,
      ]);
    }
  });

  it("L3 a legacy ticket after a legacy ticket on the same pane turn is refused (A different server turn is already bound.)", async () => {
    const h = await harness();
    expect(await h.run(ASK_PLAN, envelope(OP1))).toContain("\"reason\":\"ask\"");
    expect(await h.run(ASK_PLAN, envelope(OP2))).toBe(notExecutedBody("rejected"));
    expect(h.bound(1)).toEqual({ ok: false, reason: "A different server turn is already bound." });
  });
});

// ---------------------------------------------------------------------------
// 5.3 GREEN rows (new behaviour)
// ---------------------------------------------------------------------------
const forbidden = () => vi.fn(() => { throw new Error("must not be called"); });
function forbiddenStubs() {
  const stubs = {
    coordinator: { authorizeAndRun: forbidden(), consumeWorkbookPlanConfirmation: forbidden(), cancelPendingConfirmation: forbidden() },
    bindings: { finishTerminal: forbidden(), finishInspection: forbidden() },
    readRuntimeContext: forbidden(),
    createAdapter: forbidden(),
    createStore: forbidden(),
    inspect: forbidden(),
    awaitPaneConfirmation: forbidden(),
    executePrimitive: forbidden(),
    now: forbidden(),
  };
  const calls = (): number => [
    stubs.coordinator.authorizeAndRun, stubs.coordinator.consumeWorkbookPlanConfirmation, stubs.coordinator.cancelPendingConfirmation,
    stubs.bindings.finishTerminal, stubs.bindings.finishInspection, stubs.readRuntimeContext, stubs.createAdapter, stubs.createStore,
    stubs.inspect, stubs.awaitPaneConfirmation, stubs.executePrimitive, stubs.now,
  ].reduce((sum, stub) => sum + stub.mock.calls.length, 0);
  return { stubs, calls };
}
const ACTIVE_TURN = {
  turnOperationId: "t",
  trustedRequest: { classification: "explicit_mutation", selectionScopedRequest: false },
  sessionId: "S",
  turnStart: 1,
  hasCancellation: true,
};
function recordingRoute(target: Record<string, unknown>) {
  const reads: string[] = [];
  const dependencies = new Proxy(target, {
    get(object, property) {
      reads.push(String(property));
      return Reflect.get(object, property);
    },
  }) as unknown as RouteDependencies;
  const route = routeModule.createServerOwnedPlanRoute(dependencies);
  const constructionReads = [...reads];
  return { route, reads, constructionReads };
}
const gates = (rollout = true, session: string | null = "S") => ({
  isRolloutEnabled: vi.fn(() => rollout),
  getBridgeSessionId: vi.fn(() => session),
  getActiveTurn: vi.fn(() => ACTIVE_TURN),
});

describe("R4g-1 GREEN rows", () => {
  it("S1 gate OFF (explicit override false): a phase-inspect ticket answers the T2 rejected body after reading only the 3 pre-read gates + serverInspectionSupported; zero stub calls", async () => {
    expect((routeModule as Record<string, unknown>).SERVER_INSPECTION_SUPPORTED).toBe(true);
    const { stubs, calls } = forbiddenStubs();
    const { route, reads, constructionReads } = recordingRoute({ ...stubs, ...gates(), serverInspectionSupported: () => false });
    expect(constructionReads).toEqual([]);
    const body = await route.executeServerOwnedWorkbookPlan({ ...INSPECT_PLAN }, { turnOperationId: OP1, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT, phase: "inspect", turnKey: K } as never);
    expect(body).toBe(notExecutedBody("rejected"));
    expect(reads).toEqual(["isRolloutEnabled", "getBridgeSessionId", "getActiveTurn", "serverInspectionSupported"]);
    expect(calls()).toBe(0);
  });

  it("S1b gate OFF on the real chain: no binding is created and a following execute on the pane turn behaves as today's first execute; an override answering false is off", async () => {
    for (const gate of [false]) {
      const h = await harness(gate === undefined ? {} : { gate });
      expect(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K)), String(gate)).toBe(notExecutedBody("rejected"));
      expect(h.bindSpy).not.toHaveBeenCalled();
      expect(h.inspect).not.toHaveBeenCalled();
      expect(JSON.parse(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K)))).toMatchObject({ ok: true, result: { outcome: "committed" } });
      expect(h.stateOf(0)?.phase).toBe("terminal");
      expect(h.acks.count).toBe(1);
    }
  });

  it.each([
    ["phase inspect + execute outcome", EXECUTE_PLAN, { phase: "inspect", turnKey: K }],
    ["phase execute + inspect outcome", INSPECT_PLAN, { phase: "execute", turnKey: K }],
    ["phase inspect + ask outcome", ASK_PLAN, { phase: "inspect", turnKey: K }],
    ["phase without turnKey", INSPECT_PLAN, { phase: "inspect" }],
    ["turnKey without phase", EXECUTE_PLAN, { turnKey: K }],
    ["turnKey not a UUID", EXECUTE_PLAN, { phase: "execute", turnKey: "not-a-uuid" }],
    ["turnKey equals turnOperationId", EXECUTE_PLAN, { phase: "execute", turnKey: OP1 }],
    ["unknown phase", EXECUTE_PLAN, { phase: "commit", turnKey: K }],
    ["non-string outcome", { ...EXECUTE_PLAN, outcome: 1 }, { phase: "execute", turnKey: K }],
  ] as const)("S2 %s -> invalid before authorizeAndRun; every non-gate dependency uncalled", async (_label, plan, extra) => {
    const { stubs, calls } = forbiddenStubs();
    const { route, reads } = recordingRoute({ ...stubs, ...gates(), serverInspectionSupported: forbidden() });
    const body = await route.executeServerOwnedWorkbookPlan({ ...plan } as Record<string, unknown>, { turnOperationId: OP1, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT, ...extra } as never);
    expect(body).toBe(notExecutedBody("invalid"));
    expect(calls()).toBe(0);
    expect(reads).toEqual(["isRolloutEnabled", "getBridgeSessionId", "getActiveTurn"]);
  });

  it("S3 the inspection body has exactly the §4 keys in order, no reason/summary/state/error/execution, and reads deep-equal the inspect() reads", async () => {
    const h = await harness({ gate: true });
    const body = JSON.parse(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K))) as Record<string, unknown>;
    expect(Object.keys(body)).toEqual(["ok", "action", "phase", "result"]);
    expect(Object.keys(body.result as object)).toEqual(["outcome", "inspectionCount", "reads"]);
    for (const key of ["reason", "summary", "state", "error", "execution", "finalSummary"]) {
      expect(body, key).not.toHaveProperty(key);
      expect(body.result as object, key).not.toHaveProperty(key);
    }
    expect((body.result as { reads: unknown }).reads).toEqual(INSPECTION_READS);
  });

  it("S5 orchestrator direct calls: serverPhase inspect + execute outcome -> rejected before any access; serverPhase execute + inspect -> today's rejection; inspect phase -> inspected with inspectionReads; direct route ignores serverPhase", async () => {
    const ids = { sessionId: SESSION, documentId: DOC, turnId: TURN };
    const call = (payload: unknown): WorkbookPlanModelResponse => ({ content: "", toolCalls: [{ id: "c", type: "function", function: { name: "excel_execute_plan", arguments: JSON.stringify(payload) } }] });
    const access: string[] = [];
    const base = (overrides: Record<string, unknown>): WorkbookPlanTurnDependencies => ({
      identity: ids,
      trustedOperationBoundary: createTrustedOperationBoundary({ sessionId: SESSION, documentId: DOC, turnId: OP1 }),
      getCompileContextForPlan: async () => { access.push("runtime"); return COMPILE_CONTEXT; },
      policyInput: () => { access.push("policy"); throw new Error("policy"); },
      store: new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      model: async () => { access.push("model"); throw new Error("model"); },
      inspect: async () => { access.push("inspect"); return INSPECT_RESULT; },
      resourcePreflightAfterRuntime: async () => ({ ok: true as const }),
      adapter: statefulAdapter(access).adapter,
      rolloutEnabled: true,
      ...overrides,
    } as WorkbookPlanTurnDependencies);
    const server = { route: "server-owned", trustedTurnOperationId: OP1, continueAfterInspection: false };

    const state1 = new WorkbookPlanTurnState();
    const r1 = await createWorkbookPlanTurn(base({ ...server, serverPhase: "inspect", turnState: state1, initialResponse: call(EXECUTE_PLAN) }));
    expect(r1.outcome).toBe("rejected");
    expect(access).toEqual([]);
    expect(state1.phase).toBe("terminal");

    const state2 = new WorkbookPlanTurnState();
    const r2 = await createWorkbookPlanTurn(base({ ...server, serverPhase: "execute", turnState: state2, initialResponse: call(INSPECT_PLAN) }));
    expect(r2).toMatchObject({ outcome: "rejected", finalSummary: "Server-owned workbook inspection is not permitted." });
    expect(access).toEqual([]);

    const state3 = new WorkbookPlanTurnState();
    const r3 = await createWorkbookPlanTurn(base({ ...server, serverPhase: "inspect", turnState: state3, initialResponse: call(INSPECT_PLAN) }));
    expect(r3).toMatchObject({ outcome: "inspected", inspectionCount: 1, inspectionReads: INSPECT_RESULT });
    expect(access).toEqual(["inspect"]);
    expect(state3.phase).toBe("inspected");

    access.length = 0;
    const direct = { route: "direct", continueAfterInspection: false, initialResponse: call(INSPECT_PLAN) };
    const withPhase = await createWorkbookPlanTurn(base({ ...direct, serverPhase: "inspect" }));
    const withoutPhase = await createWorkbookPlanTurn(base({ ...direct }));
    expect(withPhase).toEqual(withoutPhase);
    expect(withPhase).not.toHaveProperty("inspectionReads");
    expect(access).toEqual(["inspect", "inspect"]);
  });

  it("S6 store unit: finishInspection keeps an inspected state and terminals any other; the inspect entry is terminal either way", () => {
    for (const begin of [true, false]) {
      const store = new ServerOwnedTurnBindingStore({ now: () => 0, ttlMs: TTL_MS, maxEntries: 8 });
      const turn = { sessionId: SESSION, documentId: DOC, paneId: PANE, turnId: TURN };
      store.activateTurn(turn);
      const ticket = { ...turn, serverOperationId: OP1, phase: "inspect", turnKey: K } as ServerOwnedTicketIdentity;
      const bound = store.bindServerTicket(ticket);
      expect(bound.ok).toBe(true);
      const state = (bound as { binding: { state: WorkbookPlanTurnState } }).binding.state;
      if (begin) expect(state.beginInspection()).toBe(true);
      (store as unknown as { finishInspection: (t: ServerOwnedTicketIdentity) => void }).finishInspection(ticket);
      expect(state.phase, String(begin)).toBe(begin ? "inspected" : "terminal");
      expect(store.bindServerTicket(ticket)).toEqual({ ok: false, reason: "Trusted server turn is terminal." });
    }
  });

  it("S7 store unit: an inspect->execute turn takes two tombstones; capacity exhaustion refuses the NEW op and stores nothing", () => {
    const turn = { sessionId: SESSION, documentId: DOC, paneId: PANE, turnId: TURN };
    const inspect = { ...turn, serverOperationId: OP1, phase: "inspect", turnKey: K } as ServerOwnedTicketIdentity;
    const execute = { ...turn, serverOperationId: OP2, phase: "execute", turnKey: K } as ServerOwnedTicketIdentity;
    const finishInspection = (store: ServerOwnedTurnBindingStore, ticket: ServerOwnedTicketIdentity) =>
      (store as unknown as { finishInspection: (t: ServerOwnedTicketIdentity) => void }).finishInspection(ticket);

    const two = new ServerOwnedTurnBindingStore({ now: () => 0, ttlMs: TTL_MS, maxEntries: 2 });
    two.activateTurn(turn);
    const first = two.bindServerTicket(inspect);
    expect(first.ok).toBe(true);
    expect((first as { binding: { state: WorkbookPlanTurnState } }).binding.state.beginInspection()).toBe(true);
    finishInspection(two, inspect);
    expect(two.bindServerTicket(execute).ok).toBe(true);
    const other = { sessionId: SESSION, documentId: DOC, paneId: PANE_2, turnId: TURN };
    two.activateTurn(other);
    expect(two.bindServerTicket({ ...other, serverOperationId: OP3 })).toEqual({ ok: false, reason: "Trusted server turn capacity is unavailable." });

    const one = new ServerOwnedTurnBindingStore({ now: () => 0, ttlMs: TTL_MS, maxEntries: 1 });
    one.activateTurn(turn);
    const bound = one.bindServerTicket(inspect);
    const state = (bound as { binding: { state: WorkbookPlanTurnState } }).binding.state;
    expect(state.beginInspection()).toBe(true);
    finishInspection(one, inspect);
    expect(one.bindServerTicket(execute)).toEqual({ ok: false, reason: "Trusted server turn capacity is unavailable." });
    expect(state.phase).toBe("inspected");
    expect(one.bindServerTicket(execute)).toEqual({ ok: false, reason: "Trusted server turn capacity is unavailable." });
    expect(one.bindServerTicket(inspect)).toEqual({ ok: false, reason: "Trusted server turn is terminal." });
  });

  it("S8 confirmation after an inspect: the binding carries the EXECUTE op id; a confirmation for the inspect op id is not current", async () => {
    const seen: WorkbookPlanConfirmationBinding[] = [];
    const stale: string[] = [];
    const h = await harness({
      gate: true,
      prompt: AMBIGUOUS_PROMPT,
      awaitPaneConfirmation: (coordinator) => async (binding, signal) => {
        seen.push(binding);
        stale.push(await coordinator.waitForWorkbookPlanConfirmation({ ...binding, serverOperationId: OP1 }, signal)
          .then(() => "resolved", (error: unknown) => (error instanceof DOMException ? `${error.name}:${error.message}` : "other")));
        const waiting = coordinator.waitForWorkbookPlanConfirmation(binding, signal);
        expect(coordinator.continuePendingWorkbookPlan()).toBe(true);
        return waiting;
      },
    });
    expect(await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(INSPECTION_BODY);
    const payload = JSON.parse(await h.run(EXECUTE_PLAN, envelope(OP2, "execute", K))) as Record<string, unknown>;
    expect(seen).toHaveLength(1);
    expect(seen[0]!.serverOperationId).toBe(OP2);
    expect(stale).toEqual(["AbortError:The pending workbook action is no longer active."]);
    expect(payload).toMatchObject({ ok: true, result: { outcome: "committed" } });
    expect(h.acks.count).toBe(1);
  });

  it("S9/S10 static: constant true once; taskpane never passes the override; bridge blob unchanged; legacy texts present; no forbidden inspection tokens", () => {
    const read = (path: string): string => readFileSync(new URL(path, import.meta.url), "utf8");
    const route = read("../taskpane/serverOwnedPlanRoute.ts");
    const taskpane = read("../taskpane/taskpane.ts");
    const orchestrator = read("../workbookPlanOrchestrator.ts");
    expect(route.split("export const SERVER_INSPECTION_SUPPORTED = true;").length - 1).toBe(1);
    expect(route).not.toMatch(/SERVER_INSPECTION_SUPPORTED = false/);
    expect(taskpane).not.toContain("serverInspectionSupported");
    expect(taskpane).not.toContain("SERVER_INSPECTION_SUPPORTED");
    const bridge = readFileSync(new URL("../officeBridge.ts", import.meta.url));
    const blob = createHash("sha1").update(`blob ${bridge.length}\0`).update(bridge).digest("hex");
    expect(blob.slice(0, 12)).toBe("ed16694ddb9f");
    expect(route).toContain("const SERVER_INSPECTION_MESSAGE = \"Server-owned workbook inspection is not permitted.\";");
    expect(orchestrator).toContain("Server-owned workbook inspection is not permitted.");
    for (const source of [route, taskpane, orchestrator]) {
      expect(source).not.toContain("inspectionPayload");
      expect(source).not.toContain("workbook_plan_inspection");
      expect(source).not.toContain("let inspectionResult: string");
    }
  });

  it("S11 no-mix latch: a gate-on inspect bind latches plan-entered; a gate-off inspect latches nothing; a legacy inspect latches (base behaviour)", async () => {
    const decide = async (h: Awaited<ReturnType<typeof harness>>): Promise<string> => {
      const authority = h.coordinator.serverPrimitiveAuthority({ turnId: TURN, bridgeSessionId: SESSION });
      expect(authority).not.toBeNull();
      return (JSON.parse(await authority!.executePrimitive("delete_sheet", { name: "Old" })) as { error: string }).error;
    };
    const fresh = await harness();
    expect(await decide(fresh)).toBe(PRIMITIVE_DENIAL_MESSAGES.LEGACY_NO_PLAN_PATH);

    const gateOn = await harness({ gate: true });
    expect(await gateOn.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(INSPECTION_BODY);
    expect(await decide(gateOn)).toBe(PRIMITIVE_DENIAL_MESSAGES.PLAN_ENTERED);

    const gateOff = await harness({ gate: false });
    expect(await gateOff.run(INSPECT_PLAN, envelope(OP1, "inspect", K))).toBe(notExecutedBody("rejected"));
    expect(await decide(gateOff)).toBe(PRIMITIVE_DENIAL_MESSAGES.LEGACY_NO_PLAN_PATH);

    const legacy = await harness({ gate: true });
    expect(await legacy.run(INSPECT_PLAN, envelope(OP1))).toBe(notExecutedBody("rejected"));
    expect(await decide(legacy)).toBe(PRIMITIVE_DENIAL_MESSAGES.PLAN_ENTERED);
  });

  it("S12 an inspect-phase success never yields already_executed, even when the store double would replay a report", async () => {
    const claims: string[] = [];
    const replaying = new BrowserIdempotencyStore(new MemoryIdempotencyStorage());
    replaying.claim = (async () => { claims.push("claim"); throw new Error("store must not be claimed by an inspection"); }) as typeof replaying.claim;
    const h = await harness({ gate: true, createStore: () => replaying });
    const body = await h.run(INSPECT_PLAN, envelope(OP1, "inspect", K));
    expect(body).toBe(INSPECTION_BODY);
    expect(body).not.toContain("already_executed");
    expect(claims).toEqual([]);
  });

  it("S13 H4-style recording Proxy over phase-less tickets: the 3 pre-read gates read exactly H4's lists; construction reads nothing; a valid legacy ticket reads the base prefix and never serverInspectionSupported", async () => {
    const foreign = "0".repeat(64);
    const drives = [
      { rollout: false, session: "S", fingerprint: WORKBOOK_PLAN_FINGERPRINT, reason: "not_started", reads: ["isRolloutEnabled"] },
      { rollout: true, session: null, fingerprint: WORKBOOK_PLAN_FINGERPRINT, reason: "rejected", reads: ["isRolloutEnabled", "getBridgeSessionId"] },
      { rollout: true, session: "S", fingerprint: foreign, reason: "invalid", reads: ["isRolloutEnabled", "getBridgeSessionId", "getActiveTurn"] },
    ] as const;
    for (const drive of drives) {
      const { stubs, calls } = forbiddenStubs();
      const { route, reads, constructionReads } = recordingRoute({ ...stubs, ...gates(drive.rollout, drive.session) });
      expect(constructionReads).toEqual([]);
      const body = JSON.parse(await route.executeServerOwnedWorkbookPlan({ ...ASK_PLAN }, { turnOperationId: "t", workbookPlanFingerprint: drive.fingerprint } as never)) as { reason: string };
      expect(body.reason).toBe(drive.reason);
      expect(reads).toEqual(drive.reads);
      expect(calls()).toBe(0);
    }
    const { stubs } = forbiddenStubs();
    const authorizeAndRun = vi.fn(async () => ({ ok: false, reason: "x" }));
    const { route, reads } = recordingRoute({ ...stubs, ...gates(), now: () => 0, coordinator: { ...stubs.coordinator, authorizeAndRun } });
    expect(await route.executeServerOwnedWorkbookPlan({ ...ASK_PLAN }, { turnOperationId: "t", workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT } as never)).toBe(notExecutedBody("rejected"));
    expect(reads.slice(0, 4)).toEqual(["isRolloutEnabled", "getBridgeSessionId", "getActiveTurn", "now"]);
    expect(reads).not.toContain("serverInspectionSupported");
    expect(authorizeAndRun).toHaveBeenCalledTimes(1);
    expect(authorizeAndRun.mock.calls[0]).toHaveLength(2);
  });

  it("S14 identity: a legacy identity has exactly the base key set and 244 replay bytes; phase/turnKey are validated when present", async () => {
    const h = await harness();
    await h.run(ASK_PLAN, envelope(OP1));
    expect(Object.keys(h.bindSpy.mock.calls[0]![0]).sort()).toEqual(["documentId", "paneId", "serverOperationId", "sessionId", "turnId"]);

    const turn = { sessionId: SESSION, documentId: DOC, paneId: PANE, turnId: TURN };
    for (const [maxBytes, ok] of [[244, true], [243, false]] as const) {
      const store = new ServerOwnedTurnBindingStore({ now: () => 0, ttlMs: TTL_MS, maxEntries: 8, maxBytes });
      store.activateTurn(turn);
      expect(store.bindServerTicket({ ...turn, serverOperationId: OP1 }).ok, `legacy ${maxBytes}`).toBe(ok);
      const phased = new ServerOwnedTurnBindingStore({ now: () => 0, ttlMs: TTL_MS, maxEntries: 8, maxBytes });
      phased.activateTurn(turn);
      expect(phased.bindServerTicket({ ...turn, serverOperationId: OP1, phase: "inspect", turnKey: K } as ServerOwnedTicketIdentity).ok, `phase ${maxBytes}`).toBe(ok);
    }

    const invalid = { ok: false, reason: "Invalid trusted server turn identity." };
    for (const extra of [{ phase: "commit", turnKey: K }, { phase: "inspect", turnKey: "nope" }, { phase: "inspect" }, { turnKey: K }, { phase: "execute", turnKey: 7 }]) {
      const store = new ServerOwnedTurnBindingStore({ now: () => 0, ttlMs: TTL_MS, maxEntries: 8 });
      store.activateTurn(turn);
      expect(store.bindServerTicket({ ...turn, serverOperationId: OP1, ...extra } as unknown as ServerOwnedTicketIdentity), JSON.stringify(extra)).toEqual(invalid);
    }
  });
});
