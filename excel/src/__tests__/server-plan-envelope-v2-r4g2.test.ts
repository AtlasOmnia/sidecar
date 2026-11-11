import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { describe, expect, it, vi } from "vitest";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import { OfficeCompanionBridge, type BridgePlanExecutionControl, type BridgeTrustedPlanEnvelopeV2 } from "../officeBridge";
import * as routeModule from "../taskpane/serverOwnedPlanRoute";
import { TaskPaneTurnCoordinator } from "../taskpane/taskPaneTurnCoordinator";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import {
  BrowserIdempotencyStore,
  MemoryIdempotencyStorage,
  type WorkbookPlanExecutionAdapter,
} from "../workbookPlanExecutor";
import {
  ServerOwnedTurnBindingStore,
  WorkbookPlanTurnState,
  type ServerOwnedTurnBindingResult,
  type WorkbookPlanConfirmationBinding,
  type WorkbookPlanVerificationAdapter,
} from "../workbookPlanOrchestrator";
import { classifyWorkbookRequest } from "../workbookPlanRequest";
import type { WorkbookPlanRouteRuntimeContext } from "../workbookPlanRoute";

/**
 * R4g-2a (dev3-r4g2-release.md §5.1 B1-B4/B3', F2/F4/F11; OMH correction 2026-10-07 item 2): the pane leg of
 * the trusted plan envelope v2. The REAL OfficeCompanionBridge parses the exact four-key ticket envelope; while
 * the route's SERVER_INSPECTION_SUPPORTED is false it hands executeTool only the legacy two-key projection
 * (no phase-bearing envelope reaches executeTool), refuses a phase/outcome mismatch as an invalid ticket, and
 * answers a primitive ticket carrying trusted metadata with the fixed primitive error (r2d-1 A3). B4 drives the
 * REAL route through the REAL bridge. B-A5 and A4 (R4g-1 lane-B / lane-A2 folds) drive the real route chain with
 * the test-only gate override on. The route/coordinator harness is COPIED from server-inspect-phase-r4g1.test.ts
 * (that file is neither imported nor edited for this); every double is local.
 */

type RouteDependencies = Parameters<typeof routeModule.createServerOwnedPlanRoute>[0];

const SESSION = "aaaaaaaa-aaaa-4aaa-8aaa-000000000001";
const DOC = "d0c00000-0000-4000-8000-000000000001";
const PANE = "9a9e0000-0000-4000-8000-000000000002";
const TURN = "7a7a0000-0000-4000-8000-000000000004";
const OP1 = "5e7e0000-0000-4000-8000-000000000011";
const OP2 = "5e7e0000-0000-4000-8000-000000000012";
const K = "4b4b0000-0000-4000-8000-000000000021";
const TTL_MS = 90_000;
const BRIDGE_SESSION = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const BRIDGE_SESSION_2 = "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee";
const BRIDGE_PANE = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
const TICKET = "cccccccc-cccc-4ccc-8ccc-cccccccccccc";
const TICKET_2 = "cccccccc-cccc-4ccc-8ccc-ccccccccccc2";
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
const INSPECTION_READS = [{ ok: true, data: [["v"]] }];
const INSPECT_RESULT = JSON.stringify({ ok: true, reads: INSPECTION_READS });
const SUMMARY = {
  rejected: "The workbook plan was rejected before execution; nothing was executed.",
  cancelled: "The workbook plan was cancelled before execution; nothing was executed.",
} as const;
const notExecutedBody = (reason: keyof typeof SUMMARY): string =>
  JSON.stringify({ ok: false, action: "workbook_plan", reason, summary: SUMMARY[reason] });
const INSPECTION_BODY = JSON.stringify({
  ok: true,
  action: "workbook_plan",
  phase: "inspect",
  result: { outcome: "inspected", inspectionCount: 1, reads: INSPECTION_READS },
});
const FIXED_PRIMITIVE_ERROR = { ok: false, category: "Execution", error: "Office tool execution failed" };
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

/** Session-aware fetch double (shape COPIED from server-inspect-phase-r4g1.test.ts bridgeFetch); unknown routes fail loudly. */
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

async function settle(promise: Promise<unknown>): Promise<string> {
  return promise.then((value) => `resolved:${String(value)}`, (error: unknown) => `rejected:${error instanceof Error ? error.name : "?"}`);
}

function stubBridge(tickets: Ticket[], executeTool: (...args: unknown[]) => Promise<string>, resultStatuses: number[] = []) {
  const f = bridgeFetch(tickets, resultStatuses);
  const handlers: (() => void)[] = [];
  const bridge = new OfficeCompanionBridge({
    hostPrefix: "excel", documentId: async () => "doc-1", paneId: BRIDGE_PANE,
    executeTool: executeTool as ConstructorParameters<typeof OfficeCompanionBridge>[0]["executeTool"],
    fetchImpl: f.fetchImpl,
    setIntervalImpl: ((handler: () => void) => { handlers.push(handler); return handlers.length; }) as unknown as typeof setInterval,
    clearIntervalImpl: (() => undefined) as unknown as typeof clearInterval,
  });
  return { bridge, f };
}

// ---------------------------------------------------------------------------
// Route/coordinator harness (COPIED shape from server-inspect-phase-r4g1.test.ts; not imported)
// ---------------------------------------------------------------------------
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

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

async function harness(options: { readonly gate?: boolean; readonly inspect?: (reads: unknown, signal?: AbortSignal) => Promise<string> } = {}) {
  const events: string[] = [];
  const clock = { t: 1_000_000 };
  const bindings = new ServerOwnedTurnBindingStore({ now: () => clock.t, ttlMs: TTL_MS, maxEntries: 32, maxBytes: 64 * 1024 });
  const bindSpy = vi.spyOn(bindings, "bindServerTicket");
  const coordinator = new TaskPaneTurnCoordinator({
    getOrCreateDocumentId: async () => DOC,
    getBridgeIdentity: () => ({ sessionId: SESSION, documentId: DOC }),
    reconnectForDocumentChange: async () => { throw new Error("unexpected document change"); },
    bindings,
    paneId: PANE,
  });
  const controller = new AbortController();
  const trustedRequest = classifyWorkbookRequest("write ok into A1");
  const turnStart = Date.now();
  const handle = coordinator.beginTurn(TURN, controller, { trustedRequest, turnStart, cancellationSignal: controller.signal, workbookPlanRolloutEnabled: true });
  await expect(coordinator.prepareTurn(handle)).resolves.toMatchObject({ status: "active" });
  const { adapter, cells } = statefulAdapter(events);
  const store = new BrowserIdempotencyStore(new MemoryIdempotencyStorage());
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
    createStore: () => store,
    inspect,
    awaitPaneConfirmation: (binding: WorkbookPlanConfirmationBinding, signal?: AbortSignal) => coordinator.waitForWorkbookPlanConfirmation(binding, signal),
    executePrimitive: async (name: string) => { events.push(`primitive:${name}`); return JSON.stringify({ ok: true }); },
    ...(options.gate === undefined ? {} : { serverInspectionSupported: () => options.gate }),
  };
  const route = routeModule.createServerOwnedPlanRoute(dependencies as unknown as RouteDependencies);
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
  return { events, clock, bindings, bindSpy, coordinator, controller, route, cells, inspect, acks, run, bound, stateOf, count, store };
}

const WORKBOOK_TOUCHES = ["runtime-read", "store-claim", "fresh", "mutation", "verify-read", "ack"];
function expectNoWriteAccess(h: Awaited<ReturnType<typeof harness>>): void {
  for (const touch of WORKBOOK_TOUCHES) expect(h.count(touch), touch).toBe(0);
  expect(h.cells.size).toBe(0);
}

// ---------------------------------------------------------------------------
// B1 / B-A1: no phase-bearing envelope reaches executeTool while the gate is off
// ---------------------------------------------------------------------------
describe("R4g-2a B1/B-A1: the bridge parses the v2 envelope and hands executeTool the ticket's own v2 envelope (gate on)", () => {
  it.each([
    ["execute", EXECUTE_PLAN],
    ["execute", ASK_PLAN],
    ["inspect", INSPECT_PLAN],
  ] as const)("B1 a %s-phase ticket (%j) reaches executeTool once with the ticket's own v2 envelope (gate on) and the frozen control", async (phase, plan) => {
    expect((routeModule as Record<string, unknown>).SERVER_INSPECTION_SUPPORTED).toBe(true);
    const executeTool = vi.fn(async (..._args: unknown[]) => JSON.stringify({ ok: false, action: "workbook_plan", reason: "rejected", summary: SUMMARY.rejected }));
    const { bridge, f } = stubBridge([planTicket({ ...plan }, v2(REQUEST, phase))], executeTool);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(executeTool).toHaveBeenCalledTimes(1);
    const args = executeTool.mock.calls[0]!;
    expect(args).toHaveLength(4);
    expect(args[0]).toBe("excel_execute_plan");
    expect(args[1]).toEqual(plan);
    expect(args[2]).toEqual(v2(REQUEST, phase));
    expect(Object.keys(args[2] as object)).toEqual(["turnOperationId", "workbookPlanFingerprint", "phase", "turnKey"]);
    expect(args[2]).toHaveProperty("phase", phase);
    expect(args[2]).toHaveProperty("turnKey", K);
    expect(Object.isFrozen(args[3])).toBe(true);
    expect(f.posts).toHaveLength(1);
    expect(f.posts[0]!.body.executed).toBe(false);
  });

  it("B1b the type: the parser's V2 shape is assignable to the unchanged BridgeTrustedPlanEnvelope (compile-time) and the source keeps the 2-key interface", () => {
    const sample: BridgeTrustedPlanEnvelopeV2 = { turnOperationId: OP1, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT, phase: "execute", turnKey: K };
    const legacy: { turnOperationId: string; workbookPlanFingerprint: string } = sample;
    expect(legacy.turnOperationId).toBe(OP1);
    const source = readFileSync(new URL("../officeBridge.ts", import.meta.url), "utf8");
    expect(source).toContain("export interface BridgeTrustedPlanEnvelope {\n  turnOperationId: string;\n  workbookPlanFingerprint: typeof WORKBOOK_PLAN_FINGERPRINT;\n}\n");
    expect(source).toContain("import { SERVER_INSPECTION_SUPPORTED } from \"./taskpane/serverOwnedPlanRoute\";");
  });
});

// ---------------------------------------------------------------------------
// B2: every malformed / mixed-version envelope is an invalid ticket, executeTool 0
// ---------------------------------------------------------------------------
describe("R4g-2a B2: malformed and mixed-version ticket envelopes post executed:false/invalid and never reach executeTool", () => {
  it.each([
    ["v1 two keys (pre-R4g-2 companion)", EXECUTE_PLAN, { turnOperationId: REQUEST, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT }],
    ["three keys", EXECUTE_PLAN, { turnOperationId: REQUEST, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT, phase: "execute" }],
    ["five keys", EXECUTE_PLAN, { ...v2(REQUEST, "execute"), extra: "x" }],
    ["unknown phase", EXECUTE_PLAN, v2(REQUEST, "commit")],
    ["uppercase phase", EXECUTE_PLAN, v2(REQUEST, "EXECUTE")],
    ["turnKey equals turnOperationId", EXECUTE_PLAN, v2(REQUEST, "execute", REQUEST)],
    ["turnKey not a UUID", EXECUTE_PLAN, v2(REQUEST, "execute", "not-a-uuid")],
    ["turnKey not a string", EXECUTE_PLAN, { ...v2(REQUEST, "execute"), turnKey: 7 }],
    ["turnKey version 0", EXECUTE_PLAN, v2(REQUEST, "execute", "4b4b0000-0000-0000-8000-000000000021")],
    ["phase inspect + execute plan", EXECUTE_PLAN, v2(REQUEST, "inspect")],
    ["phase inspect + ask plan", ASK_PLAN, v2(REQUEST, "inspect")],
    ["phase execute + inspect plan", INSPECT_PLAN, v2(REQUEST, "execute")],
    ["foreign fingerprint", EXECUTE_PLAN, { ...v2(REQUEST, "execute"), workbookPlanFingerprint: "0".repeat(64) }],
  ] as const)("B2 %s", async (_label, plan, trusted) => {
    const executeTool = vi.fn();
    const { bridge, f } = stubBridge([planTicket({ ...plan }, trusted)], executeTool);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(executeTool).not.toHaveBeenCalled();
    expect(f.executeStarts()).toBe(0);
    expect(f.posts.map((post) => post.body)).toEqual([{ ticketId: TICKET, executed: false, result: INVALID_TICKET_RESULT }]);
  });

  // R4g-2b rule U (R4g-2a lane-B A2): the pane parser accepts canonical lowercase UUIDs only, so the uppercase
  // spelling that B2b pinned as accepted is now an invalid ticket (stricter; same assertions as B2).
  it("B2b an uppercase-hex turnKey is an invalid ticket (R4g-2b: canonical lowercase UUIDs only), executeTool 0", async () => {
    expect(K.toUpperCase()).not.toBe(K);
    const executeTool = vi.fn();
    const { bridge, f } = stubBridge([planTicket({ ...ASK_PLAN }, v2(REQUEST, "execute", K.toUpperCase()))], executeTool);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(executeTool).not.toHaveBeenCalled();
    expect(f.executeStarts()).toBe(0);
    expect(f.posts.map((post) => post.body)).toEqual([{ ticketId: TICKET, executed: false, result: INVALID_TICKET_RESULT }]);
  });
});

// ---------------------------------------------------------------------------
// B3' (r2d-1 A3, F4, J11)
// ---------------------------------------------------------------------------
describe("R4g-2a B3': a primitive ticket that carries trusted metadata posts the fixed primitive error (r2d-1 A3)", () => {
  it("B3' one post of {ticketId, result: FIXED} per delivery, executeTool 0, no execute-start; a redelivery posts again (not recorded as dispatched)", async () => {
    const executeTool = vi.fn();
    const primitive = { ticketId: TICKET, requestId: REQUEST, tool: "write_range", arguments: { sheet: "Sheet1", address: "A1", values: [[1]] }, trusted: v2(REQUEST, "execute") };
    const { bridge, f } = stubBridge([primitive, primitive], executeTool);
    await bridge.connect();
    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    expect(await settle(bridge.pollOnce())).toBe("resolved:false");
    expect(executeTool).not.toHaveBeenCalled();
    expect(f.executeStarts()).toBe(0);
    const fixed = JSON.stringify({ ticketId: TICKET, result: FIXED_PRIMITIVE_ERROR });
    expect(f.posts.map((post) => post.raw)).toEqual([fixed, fixed]);
  });

  it("B3' owner-scoped (R4j): a refused post is cached for this session only; after a session replacement it is dropped unposted", async () => {
    const executeTool = vi.fn();
    const primitive = { ticketId: TICKET, requestId: REQUEST, tool: "read_range", arguments: { sheet: "Sheet1", address: "A1" }, trusted: { anything: true } };
    const { bridge, f } = stubBridge([primitive], executeTool, [500]);
    await bridge.start();
    expect(await settle(bridge.pollOnce())).toBe("rejected:BridgeHttpError");
    expect(f.posts.map((post) => post.raw)).toEqual([JSON.stringify({ ticketId: TICKET, result: FIXED_PRIMITIVE_ERROR })]);
    await bridge.stop();
    await bridge.start();
    expect(bridge.sessionId).toBe(BRIDGE_SESSION_2);
    expect(await settle(bridge.pollOnce())).toBe("resolved:false");
    expect(f.posts).toHaveLength(1);
    expect(executeTool).not.toHaveBeenCalled();
    await bridge.stop();
  });

  it("B3' the activeToolCalls guard is untouched: a timer tick right after an A3 post still polls", async () => {
    const executeTool = vi.fn();
    const primitive = { ticketId: TICKET, requestId: REQUEST, tool: "read_range", arguments: { sheet: "Sheet1", address: "A1" }, trusted: v2(REQUEST, "execute") };
    const { bridge, f } = stubBridge([primitive], executeTool);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect((bridge as unknown as { activeToolCalls: number }).activeToolCalls).toBe(0);
    expect(await bridge.pollOnce()).toBe(false);
    expect(f.posts).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// B4: the real route through the real bridge, gate off
// ---------------------------------------------------------------------------
describe("R4g-2a B4: the REAL route through the REAL bridge with the gate on (constant true, no override)", () => {
  it("B4 a 4-key inspect ticket is answered by the route's inspection body (T8) with one read and zero acks", async () => {
    const h = await harness();
    const executeTool = (...args: unknown[]) => h.route.executeCompanionTool(
      args[0] as string, args[1] as Record<string, unknown>, args[2] as never, args[3] as BridgePlanExecutionControl,
    );
    const { bridge, f } = stubBridge([planTicket({ ...INSPECT_PLAN }, v2(REQUEST, "inspect"))], executeTool);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(f.posts.map((post) => post.raw)).toEqual([JSON.stringify({ ticketId: TICKET, executed: false, result: JSON.parse(INSPECTION_BODY) })]);
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(f.executeStarts()).toBe(0);
    expectNoWriteAccess(h);
    // The bridge handed the route the ticket's own envelope: the binding has the phase-bearing identity key set.
    expect(Object.keys(h.bindSpy.mock.calls[0]![0])).toEqual(["sessionId", "documentId", "paneId", "turnId", "serverOperationId", "phase", "turnKey"]);
  });

  it("B4b inspect then execute on the same pane turn (gate on): one read, one ack, one commit", async () => {
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
    expect(f.posts.map((post) => post.body.executed)).toEqual([false, true]);
    expect(f.posts[1]!.body.result).toMatchObject({ ok: true, action: "workbook_plan", result: { outcome: "committed" } });
    expect(h.bound(1)?.ok).toBe(true);
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(f.executeStarts()).toBe(1);
    expect(h.count("mutation")).toBe(1);
  });

  it("B4c a 4-key execute ticket on a fresh pane turn commits once through the phase route path", async () => {
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
    expect(Object.keys(h.bindSpy.mock.calls[0]![0])).toEqual(["sessionId", "documentId", "paneId", "turnId", "serverOperationId", "phase", "turnKey"]);
  });
});

// ---------------------------------------------------------------------------
// F11 (R4j B-A1): the plan-path counter moved inside the try
// ---------------------------------------------------------------------------
describe("R4g-2a F11: the plan-path activeToolCalls increment sits inside the try that wraps executeTool", () => {
  it("F11 static order: acknowledgement before the try; increment and dispatch record are the first statements in it; executeTool follows", () => {
    const source = readFileSync(new URL("../officeBridge.ts", import.meta.url), "utf8");
    const at = source.indexOf("const raw = await this.deps.executeTool(tool, argumentsValue, trusted, acknowledgement.control);");
    expect(at).toBeGreaterThan(0);
    const before = source.slice(source.lastIndexOf("const acknowledgement = this.createExecuteStartAcknowledgement(session, generation, ticketId);", at), at);
    expect(before).toMatch(/^const acknowledgement = [^\n]*\n {4}let result: unknown;\n {4}try \{\n(?: {6}\/\/[^\n]*\n)* {6}this\.activeToolCalls \+= 1;\n {6}this\.recordDispatchedTicket\(session, ticketId\);\n {6}$/);
    expect(source).not.toMatch(/stop\(\)[\s\S]{0,400}activeToolCalls = 0/);
  });

  it("F11 a held plan executeTool keeps the timer guard closed until it settles", async () => {
    const gate = deferred<string>();
    const executeTool = vi.fn(async () => gate.promise);
    const { bridge, f } = stubBridge([planTicket({ ...ASK_PLAN }, v2(REQUEST, "execute"))], executeTool);
    await bridge.connect();
    const poll = bridge.pollOnce();
    for (let i = 0; i < 20 && executeTool.mock.calls.length === 0; i += 1) await new Promise<void>((done) => setImmediate(done));
    expect((bridge as unknown as { activeToolCalls: number }).activeToolCalls).toBe(1);
    gate.resolve(JSON.stringify({ ok: false, action: "workbook_plan", reason: "ask", summary: "s" }));
    expect(await poll).toBe(true);
    expect((bridge as unknown as { activeToolCalls: number }).activeToolCalls).toBe(0);
    expect(f.posts).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// R4g-1 folds on the real route chain (test-only override on)
// ---------------------------------------------------------------------------
describe("R4g-2a folds of R4g-1 review items (route chain, gate override on)", () => {
  it("B-A5 a stale authorization blocks the inspection body: the session is invalidated while the read is in flight -> no inspection body, cancelled (the stale-authorization abort), and the binding is gone so a following execute is refused", async () => {
    let target: Awaited<ReturnType<typeof harness>> | undefined;
    const h = await harness({ gate: true, inspect: async () => { target!.bindings.invalidateSession(SESSION); return INSPECT_RESULT; } });
    target = h;
    const body = await h.run(INSPECT_PLAN, v2(OP1, "inspect"));
    expect(body).not.toBe(INSPECTION_BODY);
    expect(body).toBe(notExecutedBody("cancelled"));
    expect(h.inspect).toHaveBeenCalledTimes(1);
    expect(await h.run(EXECUTE_PLAN, v2(OP2, "execute"))).toBe(notExecutedBody("rejected"));
    expect(h.bound(1)).toEqual({ ok: false, reason: "Trusted server turn identity is unavailable." });
    expectNoWriteAccess(h);
  });

  it("A4 / probe (e): an execute op for the same turnKey is refused while the inspect read is in flight; the inspection then completes and nothing was written", async () => {
    const gate = deferred<string>();
    const reached = deferred<void>();
    const h = await harness({ gate: true, inspect: async () => { reached.resolve(); return gate.promise; } });
    const inspecting = h.run(INSPECT_PLAN, v2(OP1, "inspect"));
    await reached.promise;
    expect(await h.run(EXECUTE_PLAN, v2(OP2, "execute"))).toBe(notExecutedBody("rejected"));
    expect(h.bound(1)?.ok).toBe(false);
    expect(h.acks.count).toBe(0);
    gate.resolve(INSPECT_RESULT);
    expect(await inspecting).toBe(INSPECTION_BODY);
    expectNoWriteAccess(h);
  });
});

// ---------------------------------------------------------------------------
// W8 (TS): the J2 literal vectors, recomputed by an independent uuid5 (RFC 4122 §4.3) implementation
// ---------------------------------------------------------------------------
function uuid5(namespace: string, name: string): string {
  const ns = Buffer.from(namespace.replace(/-/g, ""), "hex");
  const hash = createHash("sha1").update(ns).update(Buffer.from(name, "utf8")).digest();
  hash[6] = (hash[6]! & 0x0f) | 0x50;
  hash[8] = (hash[8]! & 0x3f) | 0x80;
  const hex = hash.subarray(0, 16).toString("hex");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}`;
}

describe("R4g-2a W8 (TS leg): J2 vectors", () => {
  it("W8 turnKey / op(inspect) / op(execute) for (aaaaaaaa-...-aaaaaaaaaaaa, r2a-real-uuid5) with the two-character \\0 separator", () => {
    const NS = "5f8f4f1c-4c57-4a35-8b4f-c0b2aa3f9248";
    const base = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa\\0r2a-real-uuid5";
    expect(uuid5(NS, base)).toBe("5a83b37f-b6e2-5031-8d37-df04bbf71b8a");
    expect(uuid5(NS, `${base}\\0inspect`)).toBe("cfed8ba1-049f-58dd-ac4d-dad4417a0ca2");
    expect(uuid5(NS, `${base}\\0execute`)).toBe("7244df56-2215-500f-98d1-2798e9313f70");
    const plugin = readFileSync(new URL("../../plugins/atlas-sidecar-excel/__init__.py", import.meta.url), "utf8");
    expect(plugin).toContain(`_OPERATION_NAMESPACE = uuid.UUID("${NS}")`);
  });
});

// ---------------------------------------------------------------------------
// W7 (TS leg; F6 T8 / T4'): every executed:false body recorded in the v2 fixture is exactly one of
// T1 (bridge), T2 (route) or T8 (inspection). With the gate off no bridge case can post a T8 body (the
// bridge hands only two keys to the route), so the T8 bodies live in companion/plugin-side cases.
// ---------------------------------------------------------------------------
type Json = null | boolean | number | string | Json[] | { [key: string]: Json };
const isObject = (value: Json | undefined): value is { [key: string]: Json } =>
  typeof value === "object" && value !== null && !Array.isArray(value);
function t1Problems(body: Json): string[] {
  if (!isObject(body)) return ["not an object"];
  return body.ok === false && body.action === "workbook_plan" && typeof body.reason === "string"
    && typeof body.error === "string" && !("summary" in body) ? [] : ["not T1"];
}
function t2Problems(body: Json): string[] {
  if (!isObject(body)) return ["not an object"];
  return body.ok === false && body.action === "workbook_plan" && typeof body.reason === "string"
    && typeof body.summary === "string" && !("error" in body) ? [] : ["not T2"];
}
function t8Problems(body: Json): string[] {
  if (!isObject(body)) return ["not an object"];
  const problems: string[] = [];
  if (JSON.stringify(Object.keys(body)) !== JSON.stringify(["ok", "action", "phase", "result"])) problems.push("keys");
  if (body.ok !== true || body.action !== "workbook_plan" || body.phase !== "inspect") problems.push("ok/action/phase");
  const result = body.result;
  if (!isObject(result) || JSON.stringify(Object.keys(result)) !== JSON.stringify(["outcome", "inspectionCount", "reads"])) return [...problems, "result keys"];
  if (result.outcome !== "inspected" || result.inspectionCount !== 1) problems.push("outcome/inspectionCount");
  if (!Array.isArray(result.reads) || !result.reads.every((read) => isObject(read) && typeof read.ok === "boolean")) problems.push("reads");
  return problems;
}

describe("R4g-2a W7 (TS leg): T8 inspection class and the T4' closure over the v2 fixture", () => {
  it("W7 every executed:false body is exactly one of T1/T2/T8; T8 bodies carry no reason/summary/state/error/execution; the T8 cases are pinned", () => {
    const fixture = JSON.parse(readFileSync(new URL("./fixtures/companion-wire/plan-wire-contract.v2.json", import.meta.url), "utf8")) as { cases: { id: string; steps: { kind: string; path?: string; body?: Json; response?: { body?: Json } }[] }[] };
    const inspectCases: string[] = [];
    let bodies = 0;
    for (const wireCase of fixture.cases) {
      for (const step of wireCase.steps) {
        const found: Json[] = [];
        if (step.kind === "pane" && step.path === "/excel/result" && isObject(step.body ?? null) && (step.body as { executed?: Json }).executed === false) {
          found.push((step.body as { result: Json }).result);
        }
        const answer = step.response?.body;
        if (step.kind === "plugin_invoke_end" && isObject(answer ?? null) && (answer as { executed?: Json }).executed === false && "data" in (answer as object)) {
          found.push((answer as { data: Json }).data);
        }
        for (const body of found) {
          bodies += 1;
          const matches = [t1Problems(body), t2Problems(body), t8Problems(body)].filter((problems) => problems.length === 0).length;
          expect(matches, `${wireCase.id}: ${JSON.stringify(body)}`).toBe(1);
          if (t8Problems(body).length === 0) {
            for (const key of ["reason", "summary", "state", "error", "execution"]) {
              expect(body as object, wireCase.id).not.toHaveProperty(key);
              expect((body as { result: object }).result, wireCase.id).not.toHaveProperty(key);
            }
            if (!inspectCases.includes(wireCase.id)) inspectCases.push(wireCase.id);
          }
        }
      }
    }
    expect(bodies).toBeGreaterThan(40);
    expect(inspectCases).toEqual(["C24-inspect-then-execute", "C25-inspect-twice", "C27-inspect-ticket-cannot-execute", "C30-inspect-ticket-gate-on"]);
    // g2m14: a T8 body that gains `reason` (or loses a key) is no longer T8.
    const t8 = JSON.parse(INSPECTION_BODY) as Json;
    expect(t8Problems(t8)).toEqual([]);
    expect(t8Problems({ ...(t8 as object), reason: "rejected" } as Json)).not.toEqual([]);
    expect(t8Problems({ ok: true, action: "workbook_plan", result: (t8 as { result: Json }).result } as Json)).not.toEqual([]);
  });
});
