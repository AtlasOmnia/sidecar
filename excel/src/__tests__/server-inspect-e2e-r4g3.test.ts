import { readFileSync } from "node:fs";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import * as officeBridgeModule from "../officeBridge";
import type { BridgePlanExecutionControl } from "../officeBridge";
import * as routeModuleStatic from "../taskpane/serverOwnedPlanRoute";
import * as coordinatorModule from "../taskpane/taskPaneTurnCoordinator";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import * as executorModule from "../workbookPlanExecutor";
import type { WorkbookPlanExecutionAdapter } from "../workbookPlanExecutor";
import * as orchestratorModule from "../workbookPlanOrchestrator";
import type {
  ServerOwnedTurnBindingResult,
  WorkbookPlanConfirmationBinding,
  WorkbookPlanTurnState,
  WorkbookPlanVerificationAdapter,
} from "../workbookPlanOrchestrator";
import * as requestModule from "../workbookPlanRequest";
import type { WorkbookPlanRouteRuntimeContext } from "../workbookPlanRoute";

/**
 * R4g-3 (dev3-r4g3-release.md §1.1 item 3, §5, RE-PIN R2/R4, OMH release record 2026-10-07): server inspect then one
 * execute, end to end on the pane side with SERVER_INSPECTION_SUPPORTED = true and NO test override (the constant
 * decides). The REAL OfficeCompanionBridge parses the companion's v2 ticket and hands the REAL route the ticket's own
 * envelope; the route drives the real coordinator/binding store, orchestrator, executor and store. The companion is
 * a stub fetch double answering v2-fixture-shaped tickets (no spawned Python; the Python half lives in
 * scripts/tests/test_companion_task13.py test_r4g3_*). The two halves are joined by the K7 / E-8 literals below,
 * pinned independently in both languages.
 *
 * Harness (harness, stubBridge, planTicket, v2) COPIED from server-plan-envelope-v2-r4g2.test.ts (not imported) and
 * parameterised by a module set, so the same code runs over the static module graph (gate on) and over a fresh
 * vi.doMock graph (E5 gate-off twins, OQ-G3-6; real taskpane reader rows E-8..E-11, N16).
 */

type RouteModule = typeof import("../taskpane/serverOwnedPlanRoute");
type RouteDependencies = Parameters<RouteModule["createServerOwnedPlanRoute"]>[0];
interface Modules {
  readonly route: RouteModule;
  readonly bridge: typeof import("../officeBridge");
  readonly coordinator: typeof import("../taskpane/taskPaneTurnCoordinator");
  readonly executor: typeof import("../workbookPlanExecutor");
  readonly orchestrator: typeof import("../workbookPlanOrchestrator");
  readonly request: typeof import("../workbookPlanRequest");
}
const STATIC_MODULES: Modules = {
  route: routeModuleStatic,
  bridge: officeBridgeModule,
  coordinator: coordinatorModule,
  executor: executorModule,
  orchestrator: orchestratorModule,
  request: requestModule,
};
/**
 * After vi.resetModules (+ doMock): every harness class from the same fresh graph (RE-PIN R2: mixed graphs reject).
 * The bridge is imported first: the route imports officeBridge, so importing the route first would let the
 * mock factory's importOriginal load officeBridge against the ORIGINAL route module (the gate would stay true).
 */
async function freshModules(): Promise<Modules> {
  const bridge = await import("../officeBridge");
  return {
    route: await import("../taskpane/serverOwnedPlanRoute"),
    bridge,
    coordinator: await import("../taskpane/taskPaneTurnCoordinator"),
    executor: await import("../workbookPlanExecutor"),
    orchestrator: await import("../workbookPlanOrchestrator"),
    request: await import("../workbookPlanRequest"),
  };
}

const SESSION = "aaaaaaaa-aaaa-4aaa-8aaa-000000000001";
const DOC = "d0c00000-0000-4000-8000-000000000001";
const PANE = "9a9e0000-0000-4000-8000-000000000002";
const TURN = "7a7a0000-0000-4000-8000-000000000004";
const OP1 = "5e7e0000-0000-4000-8000-000000000011";
const K = "4b4b0000-0000-4000-8000-000000000021";
const K2 = "4b4b0000-0000-4000-8000-000000000022";
const TTL_MS = 90_000;
const BRIDGE_SESSION = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const BRIDGE_SESSION_2 = "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee";
const BRIDGE_PANE = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
const TICKET = "cccccccc-cccc-4ccc-8ccc-cccccccccccc";
const TICKET_2 = "cccccccc-cccc-4ccc-8ccc-ccccccccccc2";
const TICKET_3 = "cccccccc-cccc-4ccc-8ccc-ccccccccccc3";
const REQUEST = "dddddddd-dddd-4ddd-8ddd-dddddddddddd";
const REQUEST_2 = "dddddddd-dddd-4ddd-8ddd-ddddddddddd2";

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
const ASK_PLAN = { version: 1, outcome: "ask", message: "Which sheet?" } as const;
/** The v2 fixture's canned T8 read (C24/C30): what the inspect double returns. */
const INSPECTION_READS = [{ ok: true, data: [["ok"]] }];
const INSPECT_RESULT = JSON.stringify({ ok: true, reads: INSPECTION_READS });
const SUMMARY = {
  ask: "The workbook plan asked for clarification; nothing was executed.",
  rejected: "The workbook plan was rejected before execution; nothing was executed.",
  cancelled: "The workbook plan was cancelled before execution; nothing was executed.",
  not_started: "Workbook plan execution did not start; nothing was executed.",
} as const;
const notExecutedBody = (reason: keyof typeof SUMMARY): string =>
  JSON.stringify({ ok: false, action: "workbook_plan", reason, summary: SUMMARY[reason] });

/** K7 (§3): the one inspection-body literal, pinned independently in scripts/tests/test_companion_task13.py. Key order is part of it. */
const K7 = "{\"ok\":true,\"action\":\"workbook_plan\",\"phase\":\"inspect\",\"result\":{\"outcome\":\"inspected\",\"inspectionCount\":1,\"reads\":[{\"ok\":true,\"data\":[[\"ok\"]]}]}}";
/** E-8 (RE-PIN R2): the real taskpane reader's T8 body for read_range Sheet1!A1; also pinned in the Python half. */
const E8_BODY = "{\"ok\":true,\"action\":\"workbook_plan\",\"phase\":\"inspect\",\"result\":{\"outcome\":\"inspected\",\"inspectionCount\":1,\"reads\":[{\"ok\":true,\"data\":{\"sheet\":\"Sheet1\",\"address\":\"Sheet1!A1\",\"rowCount\":1,\"columnCount\":1,\"values\":[[\"ok\"]],\"formulas\":[[\"ok\"]],\"numberFormat\":[[\"General\"]]}}]}}";
/** E-9: a failing real read is the fixed {ok:false, data:null}; also pinned in the Python half. */
const E9_BODY = "{\"ok\":true,\"action\":\"workbook_plan\",\"phase\":\"inspect\",\"result\":{\"outcome\":\"inspected\",\"inspectionCount\":1,\"reads\":[{\"ok\":false,\"data\":null}]}}";
/** The observed phase-bearing binding identity (R4g-3 rule B; B4c). */
const PHASE_IDENTITY_KEYS = ["sessionId", "documentId", "paneId", "turnId", "serverOperationId", "phase", "turnKey"];
const UNCERTAIN_BODY = { ok: false, action: "workbook_plan", state: "uncertain", error: "Workbook plan execution ended without a confirmed result." };
const OVERSIZE_NOT_EXECUTED = { ok: false, action: "workbook_plan", category: "Execution", reason: "rejected", error: "Workbook plan was not executed." };
const INVALID_TICKET_RESULT = { ok: false, action: "workbook_plan", reason: "invalid", error: "Workbook plan ticket was invalid and was not executed." };

const v2 = (op: string, phase: string, turnKey: string = K): Record<string, unknown> =>
  ({ turnOperationId: op, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT, phase, turnKey });

function response(status: number, body?: unknown): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
  });
}

interface Ticket { readonly ticketId: string; readonly requestId: string; readonly tool: string; readonly arguments: unknown; readonly trusted?: unknown }

/** Session-aware fetch double (shape COPIED from server-plan-envelope-v2-r4g2.test.ts bridgeFetch); unknown routes fail loudly. */
function bridgeFetch(tickets: Ticket[], resultStatuses: number[] = []) {
  const posts: { session: string; body: Record<string, unknown>; raw: string }[] = [];
  let sessions = 0;
  let executeStarts = 0;
  const queue = [...tickets];
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
      const next = queue.shift();
      return next === undefined ? response(204) : response(200, next);
    }
    if (method === "POST" && url === "/excel/execute-start") { executeStarts += 1; return response(200, { ok: true }); }
    if (method === "POST" && url === "/excel/result") {
      const raw = String(init?.body);
      posts.push({ session, body: JSON.parse(raw) as Record<string, unknown>, raw });
      const status = resultStatuses.shift() ?? 200;
      return status === 200 ? response(200, { ok: true }) : response(status, { error: "refused" });
    }
    throw new Error(`unexpected fetch ${method} ${url}`);
  });
  return { fetchImpl: fetchImpl as unknown as typeof fetch, posts, executeStarts: () => executeStarts };
}

function planTicket(argumentsValue: unknown, trusted: unknown, ticketId = TICKET, requestId = REQUEST): Ticket {
  return { ticketId, requestId, tool: "excel_execute_plan", arguments: argumentsValue, ...(trusted === undefined ? {} : { trusted }) };
}

function makeStubBridge(m: Modules) {
  return function stubBridge(tickets: Ticket[], executeTool: (...args: unknown[]) => Promise<string>, resultStatuses: number[] = []) {
    const f = bridgeFetch(tickets, resultStatuses);
    const handlers: (() => void)[] = [];
    const bridge = new m.bridge.OfficeCompanionBridge({
      hostPrefix: "excel", documentId: async () => "doc-1", paneId: BRIDGE_PANE,
      executeTool: executeTool as ConstructorParameters<typeof m.bridge.OfficeCompanionBridge>[0]["executeTool"],
      fetchImpl: f.fetchImpl,
      setIntervalImpl: ((handler: () => void) => { handlers.push(handler); return handlers.length; }) as unknown as typeof setInterval,
      clearIntervalImpl: (() => undefined) as unknown as typeof clearInterval,
    });
    return { bridge, f };
  };
}

// ---------------------------------------------------------------------------
// Route/coordinator harness (COPIED shape from server-plan-envelope-v2-r4g2.test.ts; not imported). Additions
// (defaults = the copied behaviour): rollout, activeTurn, onFirstGroup (N15), claim spy on the store.
// ---------------------------------------------------------------------------
function statefulAdapter(events: string[], onFirstGroup?: () => void) {
  const cells = new Map<string, unknown>();
  const adapter: WorkbookPlanExecutionAdapter & WorkbookPlanVerificationAdapter = {
    async executeFirstGroup(gate) {
      events.push("fresh");
      return gate.executeAfterFresh("fresh", async (group) => {
        events.push("mutation");
        for (const operation of group.operations) {
          cells.set(`${String(operation.args.sheet)}!${String(operation.args.address)}`, operation.args.values);
        }
        onFirstGroup?.();
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

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

interface HarnessOptions {
  readonly gate?: boolean;
  readonly inspect?: (reads: unknown, signal?: AbortSignal) => Promise<string>;
  readonly rollout?: boolean;
  readonly activeTurn?: boolean;
  readonly onFirstGroup?: (clock: { t: number }) => void;
}

function makeHarness(m: Modules) {
  return async function harness(options: HarnessOptions = {}) {
    const events: string[] = [];
    const clock = { t: 1_000_000 };
    const bindings = new m.orchestrator.ServerOwnedTurnBindingStore({ now: () => clock.t, ttlMs: TTL_MS, maxEntries: 32, maxBytes: 64 * 1024 });
    const bindSpy = vi.spyOn(bindings, "bindServerTicket");
    const coordinator = new m.coordinator.TaskPaneTurnCoordinator({
      getOrCreateDocumentId: async () => DOC,
      getBridgeIdentity: () => ({ sessionId: SESSION, documentId: DOC }),
      reconnectForDocumentChange: async () => { throw new Error("unexpected document change"); },
      bindings,
      paneId: PANE,
    });
    const controller = new AbortController();
    const trustedRequest = m.request.classifyWorkbookRequest("write ok into A1");
    const turnStart = Date.now();
    const handle = coordinator.beginTurn(TURN, controller, { trustedRequest, turnStart, cancellationSignal: controller.signal, workbookPlanRolloutEnabled: true });
    await expect(coordinator.prepareTurn(handle)).resolves.toMatchObject({ status: "active" });
    const { adapter, cells } = statefulAdapter(events, options.onFirstGroup === undefined ? undefined : () => options.onFirstGroup!(clock));
    const store = new m.executor.BrowserIdempotencyStore(new m.executor.MemoryIdempotencyStorage());
    const claimSpy = vi.spyOn(store, "claim");
    const inspect = vi.fn(async (reads: unknown, signal?: AbortSignal) => {
      events.push("inspect-read");
      return options.inspect === undefined ? INSPECT_RESULT : options.inspect(reads, signal);
    });
    const activeTurn = options.activeTurn === false
      ? { turnOperationId: null, trustedRequest: null, sessionId: null, turnStart: null, hasCancellation: false }
      : { turnOperationId: TURN, trustedRequest, sessionId: SESSION, turnStart, hasCancellation: true };
    const dependencies: Record<string, unknown> = {
      isRolloutEnabled: () => options.rollout ?? true,
      getBridgeSessionId: () => SESSION,
      getActiveTurn: () => activeTurn,
      coordinator,
      bindings,
      readRuntimeContext: async () => { events.push("runtime-read"); return RUNTIME; },
      createAdapter: () => adapter,
      createStore: () => store,
      inspect,
      awaitPaneConfirmation: (binding: WorkbookPlanConfirmationBinding, signal?: AbortSignal) => coordinator.waitForWorkbookPlanConfirmation(binding, signal),
      executePrimitive: async (name: string) => { events.push(`primitive:${name}`); return JSON.stringify({ ok: true }); },
      ...(options.gate === undefined ? {} : { serverInspectionSupported: () => options.gate }),
    };
    const route = m.route.createServerOwnedPlanRoute(dependencies as unknown as RouteDependencies);
    const acks = { count: 0 };
    const control = (): BridgePlanExecutionControl => ({
      acknowledgeExecuteStart: async () => { acks.count += 1; events.push("ack"); },
    });
    const run = (plan: Record<string, unknown>, ticket: Record<string, unknown>): Promise<string> =>
      route.executeServerOwnedWorkbookPlan({ ...plan }, ticket as never, control());
    const bound = (index: number): ServerOwnedTurnBindingResult => bindSpy.mock.results[index]?.value as ServerOwnedTurnBindingResult;
    const stateOf = (index: number): WorkbookPlanTurnState | undefined => {
      const result = bound(index);
      return result?.ok === true ? result.binding.state : undefined;
    };
    const count = (event: string): number => events.filter((entry) => entry === event).length;
    return { events, clock, bindings, bindSpy, coordinator, controller, route, cells, inspect, acks, run, bound, stateOf, count, store, claimSpy };
  };
}

const harness = makeHarness(STATIC_MODULES);
const stubBridge = makeStubBridge(STATIC_MODULES);
type Harness = Awaited<ReturnType<typeof harness>>;

const WORKBOOK_TOUCHES = ["runtime-read", "store-claim", "fresh", "mutation", "verify-read", "ack"];
function expectNoWriteAccess(h: Harness): void {
  for (const touch of WORKBOOK_TOUCHES) expect(h.count(touch), touch).toBe(0);
  expect(h.cells.size).toBe(0);
}

/** The real bridge wired to the harness route exactly as taskpane.ts:163 wires executeTool. */
function wired(h: Harness, tickets: Ticket[], resultStatuses: number[] = [], spy?: (args: unknown[]) => void) {
  const executeTool = (...args: unknown[]) => {
    spy?.(args);
    return h.route.executeCompanionTool(args[0] as string, args[1] as Record<string, unknown>, args[2] as never, args[3] as BridgePlanExecutionControl);
  };
  return stubBridge(tickets, executeTool, resultStatuses);
}
const posted = (raw: string): Record<string, unknown> => JSON.parse(raw) as Record<string, unknown>;

// ---------------------------------------------------------------------------
// E rows: inspect -> one execute through the REAL bridge -> REAL route, gate = the constant (true)
// ---------------------------------------------------------------------------
describe("R4g-3 E: inspect then one execute through the REAL bridge and REAL route (SERVER_INSPECTION_SUPPORTED, no override)", () => {
  it("E-1 inspect -> execute on one turnKey: one read, the posted inspection body is the K7 bytes, then one ack, one commit, one shared turn state; the store records only the execute op", async () => {
    expect(routeModuleStatic.SERVER_INSPECTION_SUPPORTED).toBe(true);
    const h = await harness();
    const { bridge, f } = wired(h, [
      planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect")),
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2),
    ]);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    // K7 (b): the bytes the real bridge POSTs, re-serialized from the recorded request body.
    expect(f.posts[0]!.raw).toBe(`{"ticketId":"${TICKET}","executed":false,"result":${K7}}`);
    expect(JSON.stringify(posted(f.posts[0]!.raw).result)).toBe(K7);
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(h.inspect.mock.calls[0]![0]).toEqual(INSPECT_PLAN.reads);
    expect(f.executeStarts()).toBe(0);
    expect(h.stateOf(0)?.phase).toBe("inspected");
    expect(h.claimSpy).not.toHaveBeenCalled();
    expectNoWriteAccess(h);

    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts).toHaveLength(2);
    expect(f.posts[1]!.body.executed).toBe(true);
    expect(f.posts[1]!.body.result).toMatchObject({ ok: true, action: "workbook_plan", result: { outcome: "committed" } });
    expect(f.executeStarts()).toBe(1);
    expect(h.acks.count).toBe(0); // the bridge's own control acknowledged, not the harness control
    expect(h.count("mutation")).toBe(1);
    expect(h.cells.get("Sheet1!A1")).toEqual([["ok"]]);
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(h.bound(1)?.ok).toBe(true);
    expect(h.stateOf(1)).toBe(h.stateOf(0)); // one WorkbookPlanTurnState: inspected -> executing -> terminal
    expect(h.stateOf(0)?.phase).toBe("terminal");
    expect(h.claimSpy).toHaveBeenCalledTimes(1);
    expect(JSON.stringify(h.claimSpy.mock.calls[0])).not.toContain(REQUEST);
    expect(h.bindSpy.mock.calls.map(([ticket]) => [ticket.serverOperationId, ticket.phase, ticket.turnKey])).toEqual([[REQUEST, "inspect", K], [REQUEST_2, "execute", K]]);
  });

  it("K7 (a)/(c): the real route's return string, fixture C24's posted result and C30's posted result and paneReport are the K7 bytes", async () => {
    const h = await harness();
    expect(await h.run(INSPECT_PLAN, v2(OP1, "inspect"))).toBe(K7);
    const fixture = JSON.parse(readFileSync(new URL("./fixtures/companion-wire/plan-wire-contract.v2.json", import.meta.url), "utf8")) as {
      cases: { id: string; bridge?: { paneReport?: unknown }; steps: { kind: string; path?: string; body?: { executed?: unknown; result?: unknown } }[] }[];
    };
    const inspectResult = (id: string): string => {
      const wireCase = fixture.cases.find((entry) => entry.id === id)!;
      const step = wireCase.steps.find((entry) => entry.kind === "pane" && entry.path === "/excel/result" && entry.body?.executed === false)!;
      return JSON.stringify(step.body!.result);
    };
    expect(inspectResult("C24-inspect-then-execute")).toBe(K7);
    expect(inspectResult("C30-inspect-ticket-gate-on")).toBe(K7);
    expect(JSON.stringify(fixture.cases.find((entry) => entry.id === "C30-inspect-ticket-gate-on")!.bridge!.paneReport)).toBe(K7);
  });

  it("E-3 inspect -> ask (execute phase, ask outcome): the T2 ask body, zero acks, the shared state terminal", async () => {
    const h = await harness();
    const { bridge, f } = wired(h, [
      planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect")),
      planTicket({ ...ASK_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2),
    ]);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(await bridge.pollOnce()).toBe(true);
    expect(JSON.stringify(f.posts[0]!.body.result)).toBe(K7);
    expect(f.posts[1]!.raw).toBe(JSON.stringify({
      ticketId: TICKET_2, executed: false,
      result: { ok: false, action: "workbook_plan", reason: "ask", summary: SUMMARY.ask, result: { outcome: "ask", finalSummary: "Which sheet?" } },
    }));
    expect(f.executeStarts()).toBe(0);
    expect(h.stateOf(1)).toBe(h.stateOf(0));
    expect(h.stateOf(0)?.phase).toBe("terminal");
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expectNoWriteAccess(h);
  });

  it("E-5 the 5,000-cell inspect budget through the real bridge: 5,000 cells are read once; 5,001 cells are rejected with zero reads", async () => {
    const at = { ...INSPECT_PLAN, reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1:A5000" }] };
    const over = { ...INSPECT_PLAN, reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1:A5001" }] };
    const ok = await harness();
    const passing = wired(ok, [planTicket(at, v2(REQUEST, "inspect"))]);
    await passing.bridge.connect();
    expect(await passing.bridge.pollOnce()).toBe(true);
    expect(JSON.stringify(passing.f.posts[0]!.body.result)).toBe(K7);
    expect(ok.inspect).toHaveBeenCalledTimes(1);
    expect(ok.inspect.mock.calls[0]![0]).toEqual(at.reads);

    const no = await harness();
    const refused = wired(no, [planTicket(over, v2(REQUEST, "inspect"))]);
    await refused.bridge.connect();
    expect(await refused.bridge.pollOnce()).toBe(true);
    expect(refused.f.posts[0]!.raw).toBe(JSON.stringify({ ticketId: TICKET, executed: false, result: JSON.parse(notExecutedBody("rejected")) }));
    expect(no.inspect).not.toHaveBeenCalled();
    expect(no.stateOf(0)?.phase).toBe("terminal");
    expectNoWriteAccess(no);
  });

  it("E-6 / K9 an execute-only turn (no inspect) on the phase path commits exactly once with the phase-bearing binding identity", async () => {
    const h = await harness();
    const { bridge, f } = wired(h, [planTicket({ ...EXECUTE_PLAN }, v2(REQUEST, "execute"))]);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.executeStarts()).toBe(1);
    expect(f.posts[0]!.body.executed).toBe(true);
    expect(f.posts[0]!.body.result).toMatchObject({ ok: true, action: "workbook_plan", result: { outcome: "committed" } });
    expect(h.count("mutation")).toBe(1);
    expect(h.inspect).not.toHaveBeenCalled();
    expect(Object.keys(h.bindSpy.mock.calls[0]![0])).toEqual(PHASE_IDENTITY_KEYS);
    expect(h.bindSpy.mock.calls[0]![0]).toMatchObject({ phase: "execute", turnKey: K, serverOperationId: REQUEST });
  });

  it("E-7 the handed envelope is the ticket's own v2 object, keys in order [turnOperationId, workbookPlanFingerprint, phase, turnKey], for inspect and execute", async () => {
    const h = await harness();
    const handed: unknown[] = [];
    const { bridge } = wired(h, [
      planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect")),
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2),
    ], [], (args) => handed.push(args[2]));
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(await bridge.pollOnce()).toBe(true);
    expect(handed).toEqual([v2(REQUEST, "inspect"), v2(REQUEST_2, "execute")]);
    for (const envelope of handed) expect(Object.keys(envelope as object)).toEqual(["turnOperationId", "workbookPlanFingerprint", "phase", "turnKey"]);
  });

  it("D1 / N14 rollout OFF: inspect and execute tickets both post the not_started T2 body before any read or binding (byte-identical to base)", async () => {
    const h = await harness({ rollout: false });
    const { bridge, f } = wired(h, [
      planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect")),
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2),
    ]);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts.map((post) => post.raw)).toEqual([
      `{"ticketId":"${TICKET}","executed":false,"result":{"ok":false,"action":"workbook_plan","reason":"not_started","summary":"Workbook plan execution did not start; nothing was executed."}}`,
      `{"ticketId":"${TICKET_2}","executed":false,"result":{"ok":false,"action":"workbook_plan","reason":"not_started","summary":"Workbook plan execution did not start; nothing was executed."}}`,
    ]);
    expect(h.inspect).not.toHaveBeenCalled();
    expect(h.bindSpy).not.toHaveBeenCalled();
    expect(f.executeStarts()).toBe(0);
    expectNoWriteAccess(h);
  });
});

// ---------------------------------------------------------------------------
// N rows: the negative matrix through the REAL bridge -> REAL route. Zero mutation and zero acks unless stated.
// (P) = parity pin green at base and HEAD; (R) = needs a delivered inspection, red at base.
// ---------------------------------------------------------------------------
async function inspected(h: Harness, more: Ticket[] = [], resultStatuses: number[] = []) {
  const wiring = wired(h, [planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect")), ...more], resultStatuses);
  await wiring.bridge.connect();
  expect(await wiring.bridge.pollOnce()).toBe(true);
  expect(JSON.stringify(wiring.f.posts[0]!.body.result)).toBe(K7);
  return wiring;
}
const rejectedPost = (ticketId: string) => JSON.stringify({ ticketId, executed: false, result: JSON.parse(notExecutedBody("rejected")) });

describe("R4g-3 N: negative matrix through the REAL bridge and REAL route (gate = the constant)", () => {
  it("N1 (R) inspect -> inspect on one turn: the second inspect ticket is refused; one read in total", async () => {
    const h = await harness();
    const { bridge, f } = await inspected(h, [planTicket({ ...INSPECT_PLAN }, v2(REQUEST_2, "inspect"), TICKET_2, REQUEST_2)]);
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts[1]!.raw).toBe(rejectedPost(TICKET_2));
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(h.bound(1)?.ok).toBe(false);
    expect(f.executeStarts()).toBe(0);
    expectNoWriteAccess(h);
  });

  it("N2 (P) execute -> inspect: the inspect ticket is refused with zero reads", async () => {
    const h = await harness();
    const { bridge, f } = wired(h, [
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST, "execute")),
      planTicket({ ...INSPECT_PLAN }, v2(REQUEST_2, "inspect"), TICKET_2, REQUEST_2),
    ]);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts[0]!.body.executed).toBe(true);
    expect(f.posts[1]!.raw).toBe(rejectedPost(TICKET_2));
    expect(h.inspect).not.toHaveBeenCalled();
    expect(f.executeStarts()).toBe(1);
    expect(h.count("mutation")).toBe(1);
  });

  it("N3 (P) execute -> execute: the second execute ticket is refused; one commit in total", async () => {
    const h = await harness();
    const { bridge, f } = wired(h, [
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST, "execute")),
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2),
    ]);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts.map((post) => post.body.executed)).toEqual([true, false]);
    expect(f.posts[1]!.raw).toBe(rejectedPost(TICKET_2));
    expect(f.executeStarts()).toBe(1);
    expect(h.count("mutation")).toBe(1);
  });

  it("N4 (P) phase/outcome mismatch: the bridge posts T1 invalid before executeTool (C29)", async () => {
    const h = await harness();
    const handed: unknown[] = [];
    const { bridge, f } = wired(h, [
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST, "inspect")),
      planTicket({ ...INSPECT_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2),
    ], [], (args) => handed.push(args));
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(await bridge.pollOnce()).toBe(true);
    expect(handed).toEqual([]);
    expect(f.posts.map((post) => post.body)).toEqual([
      { ticketId: TICKET, executed: false, result: INVALID_TICKET_RESULT },
      { ticketId: TICKET_2, executed: false, result: INVALID_TICKET_RESULT },
    ]);
    expect(h.bindSpy).not.toHaveBeenCalled();
    expectNoWriteAccess(h);
  });

  it("N5 (R) Stop between the phases: the execute ticket is cancelled/refused with zero further reads and zero acks", async () => {
    const h = await harness();
    const { bridge, f } = await inspected(h, [planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2)]);
    h.controller.abort();
    expect(await bridge.pollOnce()).toBe(true);
    const second = posted(f.posts[1]!.raw);
    expect(second.executed).toBe(false);
    expect([notExecutedBody("cancelled"), notExecutedBody("rejected")]).toContain(JSON.stringify(second.result));
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(f.executeStarts()).toBe(0);
    expectNoWriteAccess(h);
  });

  it("N6 (R) Stop during the inspect read: the inspection is cancelled with no inspection body and the execute is refused", async () => {
    let target: Harness | undefined;
    const h = await harness({ inspect: async () => { target!.controller.abort(); return INSPECT_RESULT; } });
    target = h;
    const { bridge, f } = wired(h, [
      planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect")),
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2),
    ]);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(f.posts[0]!.raw).toBe(JSON.stringify({ ticketId: TICKET, executed: false, result: JSON.parse(notExecutedBody("cancelled")) }));
    expect(f.posts[0]!.raw).not.toContain("inspected");
    expect(await bridge.pollOnce()).toBe(true);
    expect(posted(f.posts[1]!.raw).executed).toBe(false);
    expect([notExecutedBody("cancelled"), notExecutedBody("rejected")]).toContain(JSON.stringify(posted(f.posts[1]!.raw).result));
    expect(f.executeStarts()).toBe(0);
    expectNoWriteAccess(h);
  });

  it("N7 (R) redelivery of the inspect ticket and a replay of the inspect op: the read count stays 1 and no second inspection body is posted", async () => {
    const h = await harness();
    const { bridge, f } = await inspected(h, [
      planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect")),
      planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect"), TICKET_2, REQUEST),
    ]);
    expect(await bridge.pollOnce()).toBe(false); // same ticketId: never re-executed or re-posted (R4f)
    expect(await bridge.pollOnce()).toBe(true); // same op, new ticket: the route refuses the terminal op
    expect(f.posts).toHaveLength(2);
    expect(f.posts[1]!.raw).toBe(rejectedPost(TICKET_2));
    expect(f.posts.filter((post) => post.raw.includes("\"inspected\""))).toHaveLength(1);
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expectNoWriteAccess(h);
  });

  it("N8 (P) replay of the execute op after its commit: no second ack and no second commit", async () => {
    const h = await harness();
    const { bridge, f } = wired(h, [
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST, "execute")),
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST, "execute")),
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST, "execute"), TICKET_2, REQUEST),
    ]);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(await bridge.pollOnce()).toBe(false);
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts.map((post) => post.body.executed)).toEqual([true, false]);
    expect(f.posts[1]!.raw).toBe(rejectedPost(TICKET_2));
    expect(f.executeStarts()).toBe(1);
    expect(h.count("mutation")).toBe(1);
  });

  it("N9 (R) session rotation between the phases: the execute ticket is refused at the pane with zero reads", async () => {
    const h = await harness();
    const { bridge, f } = await inspected(h, [planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2)]);
    h.bindings.invalidateSession(SESSION);
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts[1]!.raw).toBe(rejectedPost(TICKET_2));
    expect(h.bound(1)).toEqual({ ok: false, reason: "Trusted server turn identity is unavailable." });
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(f.executeStarts()).toBe(0);
    expectNoWriteAccess(h);
  });

  it("N10 (R) pane reload between the phases (fresh coordinator and binding store, no active pane turn): the execute ticket is rejected with zero mutation", async () => {
    const before = await harness();
    await inspected(before);
    const reloaded = await harness({ activeTurn: false });
    const { bridge, f } = wired(reloaded, [planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2)]);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts[0]!.raw).toBe(rejectedPost(TICKET_2));
    expect(reloaded.bindSpy).not.toHaveBeenCalled();
    expect(f.executeStarts()).toBe(0);
    expectNoWriteAccess(reloaded);
    expectNoWriteAccess(before);
  });

  it("N11 (R) a foreign turnKey execute after an inspect is refused (R4g-1 R6 path)", async () => {
    const h = await harness();
    const { bridge, f } = await inspected(h, [planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute", K2), TICKET_2, REQUEST_2)]);
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts[1]!.raw).toBe(rejectedPost(TICKET_2));
    expect(h.bound(1)?.ok).toBe(false);
    expect(f.executeStarts()).toBe(0);
    expectNoWriteAccess(h);
  });

  it("N12 (R) the route's inspect path never acknowledges execute-start: zero execute-starts and executed:false for the delivered inspection", async () => {
    const h = await harness();
    const controls: BridgePlanExecutionControl[] = [];
    const { bridge, f } = wired(h, [planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect"))], [], (args) => controls.push(args[3] as BridgePlanExecutionControl));
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(controls).toHaveLength(1);
    expect(JSON.stringify(f.posts[0]!.body.result)).toBe(K7);
    expect(f.posts[0]!.body.executed).toBe(false);
    expect(f.executeStarts()).toBe(0);
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expectNoWriteAccess(h);
  });

  it("N13 (P) a 413 on the inspect result post: the bridge re-posts the fixed R4f T1 rejected substitute, executed:false", async () => {
    const h = await harness();
    const { bridge, f } = wired(h, [planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect"))], [413]);
    await bridge.connect();
    await expect(bridge.pollOnce()).rejects.toMatchObject({ status: 413 });
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts).toHaveLength(2);
    expect(f.posts[1]!.raw).toBe(JSON.stringify({ ticketId: TICKET, executed: false, result: OVERSIZE_NOT_EXECUTED }));
    expect(f.executeStarts()).toBe(0);
    expectNoWriteAccess(h);
  });

  it("N15 (R) R4g-1 B-A2: a TTL prune of the inspect sibling while the execute is in flight gives the honest uncertain body: one ack, one mutation; replay and redelivery execute nothing", async () => {
    const h = await harness({ onFirstGroup: (clock) => { clock.t += TTL_MS; } });
    const { bridge, f } = await inspected(h, [
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2),
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2),
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute"), TICKET_3, REQUEST_2),
    ]);
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts[1]!.raw).toBe(JSON.stringify({ ticketId: TICKET_2, executed: true, result: UNCERTAIN_BODY }));
    expect(f.executeStarts()).toBe(1);
    expect(h.count("mutation")).toBe(1);
    expect(await bridge.pollOnce()).toBe(false); // redelivery of the same ticket
    expect(await bridge.pollOnce()).toBe(true); // replay of the same op on a new ticket
    expect(posted(f.posts[2]!.raw).executed).toBe(false);
    expect(f.posts).toHaveLength(3);
    expect(f.executeStarts()).toBe(1);
    expect(h.count("mutation")).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// The REAL taskpane server inspection reader (RE-PIN R2: E-8..E-11, N16). taskpane.ts is imported under
// vi.doMock with bootstrap neutered; its createServerOwnedPlanRoute call is captured to obtain the REAL `inspect`
// dependency the pane passes (no export added). Excel.run is a stub, so the real executeValidatedExcelBridgeTool
// -> excelTools.readRange path runs. No override: the constant decides.
// ---------------------------------------------------------------------------
function stubExcel(onGetItem: (name: string) => void = () => undefined): void {
  const range = { load: () => undefined, rowCount: 1, columnCount: 1, values: [["ok"]], formulas: [["ok"]], numberFormat: [["General"]] };
  const context = {
    workbook: { worksheets: { getItem: (name: string) => {
      onGetItem(name);
      if (name !== "Sheet1") throw Object.assign(new Error(`ItemNotFound secret-sheet ${name}`), { code: "ItemNotFound" });
      return { getRange: () => range };
    } } },
    sync: async () => undefined,
  };
  (globalThis as Record<string, unknown>).Excel = { run: async (work: (c: unknown) => Promise<unknown>) => work(context) };
}

async function realReaderChain(onGetItem?: (name: string, h: { controller: AbortController | undefined }) => void) {
  vi.resetModules();
  (globalThis as Record<string, unknown>).localStorage = { getItem: () => null, setItem: () => undefined };
  vi.doMock("../taskpane/bootstrapCoordinator", () => ({ bootstrap: vi.fn() }));
  let captured: Record<string, unknown> | undefined;
  vi.doMock("../taskpane/serverOwnedPlanRoute", async (importOriginal) => {
    const original = await importOriginal<RouteModule>();
    return { ...original, createServerOwnedPlanRoute: (deps: RouteDependencies) => {
      captured = deps as unknown as Record<string, unknown>;
      return original.createServerOwnedPlanRoute(deps);
    } };
  });
  await import("../taskpane/taskpane");
  const m = await freshModules();
  expect(captured).toBeDefined();
  const realInspect = captured!.inspect as (reads: unknown, signal?: AbortSignal) => Promise<string>;
  const holder: { controller: AbortController | undefined } = { controller: undefined };
  stubExcel((name) => onGetItem?.(name, holder));
  const h = await makeHarness(m)({ inspect: (reads, signal) => realInspect(reads, signal) });
  holder.controller = h.controller;
  return { m, h, realInspect, stubBridge: makeStubBridge(m) };
}

describe("R4g-3 E-8..E-11, N16: the REAL taskpane server inspection reader through the REAL route (and the REAL bridge)", () => {
  afterEach(() => {
    vi.doUnmock("../taskpane/bootstrapCoordinator");
    vi.doUnmock("../taskpane/serverOwnedPlanRoute");
    vi.resetModules();
    delete (globalThis as Record<string, unknown>).Excel;
    delete (globalThis as Record<string, unknown>).localStorage;
  });

  it("E-8 a real read_range result is {ok:true, data:<the tool object>} inside the exact T8 body: route return and the real bridge's posted result; state inspected; zero acks", async () => {
    const direct = await realReaderChain();
    expect(await direct.h.run(INSPECT_PLAN, v2(OP1, "inspect"))).toBe(E8_BODY);
    expect(direct.h.inspect).toHaveBeenCalledTimes(1);
    expect(direct.h.stateOf(0)?.phase).toBe("inspected");
    expect(direct.h.acks.count).toBe(0);

    const viaBridge = await realReaderChain();
    const executeTool = (...args: unknown[]) => viaBridge.h.route.executeCompanionTool(
      args[0] as string, args[1] as Record<string, unknown>, args[2] as never, args[3] as BridgePlanExecutionControl,
    );
    const { bridge, f } = viaBridge.stubBridge([planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect"))], executeTool);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts[0]!.raw).toBe(`{"ticketId":"${TICKET}","executed":false,"result":${E8_BODY}}`);
    expect(JSON.stringify(posted(f.posts[0]!.raw).result)).toBe(E8_BODY);
    expect(f.executeStarts()).toBe(0);
    expect(viaBridge.h.stateOf(0)?.phase).toBe("inspected");
    expectNoWriteAccess(viaBridge.h);
  });

  it("E-8b real reader inspect then one execute on the same turn through the real bridge: one read, one ack, one commit", async () => {
    const chain = await realReaderChain();
    const executeTool = (...args: unknown[]) => chain.h.route.executeCompanionTool(
      args[0] as string, args[1] as Record<string, unknown>, args[2] as never, args[3] as BridgePlanExecutionControl,
    );
    const { bridge, f } = chain.stubBridge([
      planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect")),
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2),
    ], executeTool);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(await bridge.pollOnce()).toBe(true);
    expect(JSON.stringify(f.posts[0]!.body.result)).toBe(E8_BODY);
    expect(f.posts[1]!.body.executed).toBe(true);
    expect(f.posts[1]!.body.result).toMatchObject({ ok: true, result: { outcome: "committed" } });
    expect(f.executeStarts()).toBe(1);
    expect(chain.h.count("mutation")).toBe(1);
    expect(chain.h.inspect).toHaveBeenCalledTimes(1);
    expect(chain.h.stateOf(1)).toBe(chain.h.stateOf(0));
  });

  it("E-9 a failing real read (unknown sheet, ItemNotFound) becomes the fixed {ok:false, data:null}: no sheet name, error text, ItemNotFound or error key leaves the pane; state inspected", async () => {
    const h = (await realReaderChain()).h;
    const plan = { ...INSPECT_PLAN, reads: [{ kind: "read_range", sheet: "Hidden", address: "A1" }] };
    const body = await h.run(plan, v2(OP1, "inspect"));
    expect(body).toBe(E9_BODY);
    for (const leak of ["Hidden", "secret-sheet", "ItemNotFound", "error", "Sheet not found"]) expect(body).not.toContain(leak);
    expect(h.stateOf(0)?.phase).toBe("inspected");
  });

  it("E-10 the real reader's own output: keys [ok, reads]; every read exactly [ok, data] (an ok read and a failing read)", async () => {
    const { realInspect } = await realReaderChain();
    const out = JSON.parse(await realInspect([
      { kind: "read_range", sheet: "Sheet1", address: "A1" },
      { kind: "read_range", sheet: "Nope", address: "A1" },
    ])) as { ok: boolean; reads: Record<string, unknown>[] };
    expect(Object.keys(out)).toEqual(["ok", "reads"]);
    expect(out.ok).toBe(true);
    expect(out.reads.map((read) => Object.keys(read))).toEqual([["ok", "data"], ["ok", "data"]]);
    expect(out.reads[0]).toEqual({ ok: true, data: { sheet: "Sheet1", address: "Sheet1!A1", rowCount: 1, columnCount: 1, values: [["ok"]], formulas: [["ok"]], numberFormat: [["General"]] } });
    expect(out.reads[1]).toEqual({ ok: false, data: null });
  });

  it("N16 (R) Stop during the real reader (abort before the 2nd read): cancelled, no inspection body, the execute is refused", async () => {
    let reads = 0;
    const { h } = await realReaderChain((_name, holder) => { reads += 1; if (reads === 1) holder.controller?.abort(); });
    const plan = { ...INSPECT_PLAN, reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1" }, { kind: "read_range", sheet: "Sheet1", address: "B1" }] };
    const body = await h.run(plan, v2(OP1, "inspect"));
    expect(body).toBe(notExecutedBody("cancelled"));
    expect(reads).toBe(1);
    expect(await h.run(EXECUTE_PLAN, v2(REQUEST_2, "execute"))).not.toContain("\"ok\":true");
    expect(h.acks.count).toBe(0);
    expectNoWriteAccess(h);
  });
});

describe("R4g-3 E-11 static taskpane pin (K1′): the server-only reader is wired at the server route only", () => {
  it("E-11 hunk rules: one server-reader wiring inside createServerOwnedPlanRoute({...}); the direct route keeps inspectWorkbookPlanReads; the new functions sit after inspectWorkbookPlanReads; one Office read call site", () => {
    const taskpane = readFileSync(new URL("../taskpane/taskpane.ts", import.meta.url), "utf8");
    const count = (needle: string): number => taskpane.split(needle).length - 1;
    expect(count("inspect: inspectServerWorkbookPlanReads,")).toBe(1);
    expect(count("inspect: inspectWorkbookPlanReads,")).toBe(1);
    const serverBlock = taskpane.slice(taskpane.indexOf("const serverOwnedPlanRoute = createServerOwnedPlanRoute({"), taskpane.indexOf("\n});\n", taskpane.indexOf("const serverOwnedPlanRoute = createServerOwnedPlanRoute({")));
    expect(serverBlock).toContain("\n  inspect: inspectServerWorkbookPlanReads,\n");
    expect(serverBlock).not.toContain("inspectWorkbookPlanReads,");
    const direct = taskpane.indexOf("    inspect: inspectWorkbookPlanReads,\n");
    expect(taskpane.lastIndexOf("route: \"direct\"")).toBeGreaterThan(direct);
    expect(direct).toBeGreaterThan(taskpane.indexOf("const serverOwnedPlanRoute = createServerOwnedPlanRoute({"));
    const shared = taskpane.indexOf("async function inspectWorkbookPlanReads(");
    const server = taskpane.indexOf("async function inspectServerWorkbookPlanReads(");
    const mapper = taskpane.indexOf("function toServerInspectionRead(read: unknown)");
    const next = taskpane.indexOf("async function waitForPendingPlanConfirmation(");
    expect(shared).toBeGreaterThan(0);
    expect(server).toBeGreaterThan(shared);
    expect(mapper).toBeGreaterThan(server);
    expect(next).toBeGreaterThan(mapper);
    expect(count("await executeValidatedExcelBridgeTool(read.kind, { sheet: read.sheet, address: read.address })")).toBe(1);
    const reader = taskpane.slice(server, next);
    expect(reader).toContain("await inspectWorkbookPlanReads(reads, signal)");
    expect(reader).toContain("{ ok: false, data: null }");
    expect(reader).not.toMatch(/\berror\b\s*:/);
    expect(taskpane).not.toContain("serverInspectionSupported");
    expect(taskpane).not.toContain("SERVER_INSPECTION_SUPPORTED");
  });
});

// ---------------------------------------------------------------------------
// E5 / K6′ gate-off twins (OQ-G3-6): the bridge's gate-off downgrade pinned with vi.doMock of the route constant
// (false) over a fresh module graph. The it() bodies are the R4g-2a B1 x3, B4, B4b, B4c bodies at 3e3707b,
// byte-for-byte; `harness`, `stubBridge` and `routeModule` resolve to the gate-off graph below.
// ---------------------------------------------------------------------------
describe("R4g-3 E5 twins: the gate-off bridge downgrade (doMock SERVER_INSPECTION_SUPPORTED = false)", () => {
  /* eslint-disable @typescript-eslint/no-shadow */
  let harness: ReturnType<typeof makeHarness>;
  let stubBridge: ReturnType<typeof makeStubBridge>;
  let routeModule: RouteModule;
  beforeEach(async () => {
    vi.resetModules();
    vi.doMock("../taskpane/serverOwnedPlanRoute", async (importOriginal) => ({ ...(await importOriginal<RouteModule>()), SERVER_INSPECTION_SUPPORTED: false }));
    const m = await freshModules();
    harness = makeHarness(m);
    stubBridge = makeStubBridge(m);
    routeModule = m.route;
  });
  afterEach(() => {
    vi.doUnmock("../taskpane/serverOwnedPlanRoute");
    vi.resetModules();
  });

  it.each([
    ["execute", EXECUTE_PLAN],
    ["execute", ASK_PLAN],
    ["inspect", INSPECT_PLAN],
  ] as const)("B1 a %s-phase ticket (%j) reaches executeTool once with exactly {turnOperationId, workbookPlanFingerprint} and the frozen control-gate-off (doMock)", async (phase, plan) => {
    expect((routeModule as Record<string, unknown>).SERVER_INSPECTION_SUPPORTED).toBe(false);
    const executeTool = vi.fn(async (..._args: unknown[]) => JSON.stringify({ ok: false, action: "workbook_plan", reason: "rejected", summary: SUMMARY.rejected }));
    const { bridge, f } = stubBridge([planTicket({ ...plan }, v2(REQUEST, phase))], executeTool);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(executeTool).toHaveBeenCalledTimes(1);
    const args = executeTool.mock.calls[0]!;
    expect(args).toHaveLength(4);
    expect(args[0]).toBe("excel_execute_plan");
    expect(args[1]).toEqual(plan);
    expect(args[2]).toEqual({ turnOperationId: REQUEST, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT });
    expect(Object.keys(args[2] as object)).toEqual(["turnOperationId", "workbookPlanFingerprint"]);
    expect(args[2]).not.toHaveProperty("phase");
    expect(args[2]).not.toHaveProperty("turnKey");
    expect(Object.isFrozen(args[3])).toBe(true);
    expect(f.posts).toHaveLength(1);
    expect(f.posts[0]!.body.executed).toBe(false);
  });

  it("B4 a 4-key inspect ticket is answered by the route's legacy inspect refusal (T2 rejected) with zero reads and zero acks-gate-off (doMock)", async () => {
    const h = await harness();
    const executeTool = (...args: unknown[]) => h.route.executeCompanionTool(
      args[0] as string, args[1] as Record<string, unknown>, args[2] as never, args[3] as BridgePlanExecutionControl,
    );
    const { bridge, f } = stubBridge([planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect"))], executeTool);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts.map((post) => post.raw)).toEqual([JSON.stringify({ ticketId: TICKET, executed: false, result: JSON.parse(notExecutedBody("rejected")) })]);
    expect(h.inspect).not.toHaveBeenCalled();
    expect(f.executeStarts()).toBe(0);
    expectNoWriteAccess(h);
    // The bridge handed the route a phase-less ticket: the binding has the base (legacy) identity key set.
    expect(Object.keys(h.bindSpy.mock.calls[0]![0])).toEqual(["sessionId", "documentId", "paneId", "turnId", "serverOperationId"]);
  });

  it("B4b deviation D5 (recorded): after a gate-off inspect the same pane turn's execute ticket is refused at the pane with zero mutation (v1-equivalent: one plan op per pane turn)-gate-off (doMock)", async () => {
    const h = await harness();
    const executeTool = (...args: unknown[]) => h.route.executeCompanionTool(
      args[0] as string, args[1] as Record<string, unknown>, args[2] as never, args[3] as BridgePlanExecutionControl,
    );
    const { bridge, f } = stubBridge([
      planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect")),
      planTicket({ ...EXECUTE_PLAN }, v2(REQUEST_2, "execute"), TICKET_2, REQUEST_2),
    ], executeTool);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts.map((post) => post.body.executed)).toEqual([false, false]);
    expect(f.posts[1]!.raw).toBe(JSON.stringify({ ticketId: TICKET_2, executed: false, result: JSON.parse(notExecutedBody("rejected")) }));
    expect(h.bound(1)).toEqual({ ok: false, reason: "A different server turn is already bound." });
    expect(f.executeStarts()).toBe(0);
    expectNoWriteAccess(h);
  });

  it("B4c a 4-key execute ticket on a fresh pane turn commits once through the base (phase-less) route path-gate-off (doMock)", async () => {
    const h = await harness();
    const executeTool = (...args: unknown[]) => h.route.executeCompanionTool(
      args[0] as string, args[1] as Record<string, unknown>, args[2] as never, args[3] as BridgePlanExecutionControl,
    );
    const { bridge, f } = stubBridge([planTicket({ ...EXECUTE_PLAN }, v2(REQUEST, "execute"))], executeTool);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.executeStarts()).toBe(1);
    expect(f.posts[0]!.body.executed).toBe(true);
    expect(f.posts[0]!.body.result).toMatchObject({ ok: true, action: "workbook_plan", result: { outcome: "committed" } });
    expect(h.count("mutation")).toBe(1);
    expect(Object.keys(h.bindSpy.mock.calls[0]![0])).toEqual(["sessionId", "documentId", "paneId", "turnId", "serverOperationId"]);
  });
  /* eslint-enable @typescript-eslint/no-shadow */
});
