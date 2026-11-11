import { readFileSync } from "node:fs";
import { describe, expect, it, vi } from "vitest";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import { OfficeCompanionBridge, type BridgePlanExecutionControl } from "../officeBridge";
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
 * R4g-2b extended scope (OMH release record 2026-10-07; R4g-2a lane-B A3 and A2):
 * - IR rows: the REAL route builds the T8 inspection body only from reads that are plain JSON objects with
 *   exactly the keys of the v2 fixture's T8 reads ({ok, data}; `ok` boolean), each at most READ_CAP UTF-8
 *   bytes as compact JSON, and the whole body at most BODY_CAP bytes. Anything else fails closed exactly
 *   like the R4g-1 R14 malformed-read path: the fixed `rejected` body, the shared state terminal, and the
 *   following execute refused. Driven through the real coordinator/binding store with the test-only gate
 *   override on (SERVER_INSPECTION_SUPPORTED itself stays false).
 * - UK rows: the REAL OfficeCompanionBridge's v2 parser accepts turnKey and turnOperationId only as
 *   canonical lowercase hyphenated UUIDs; uppercase, mixed case, braces, urn:, missing dashes and nil are
 *   invalid tickets (executed:false/invalid, executeTool 0, no execute-start).
 * The route/coordinator harness and the fetch double are COPIED in shape from
 * server-plan-envelope-v2-r4g2.test.ts (that file is neither imported nor edited here); every double is local.
 */

type RouteDependencies = Parameters<typeof routeModule.createServerOwnedPlanRoute>[0];

/** Provisional caps (r4g2b/case-state.md): 1 MiB per read, 1.5 MiB per inspection body. Independent literals. */
const READ_CAP = 1_048_576;
const BODY_CAP = 1_572_864;

const SESSION = "aaaaaaaa-aaaa-4aaa-8aaa-000000000001";
const DOC = "d0c00000-0000-4000-8000-000000000001";
const PANE = "9a9e0000-0000-4000-8000-000000000002";
const TURN = "7a7a0000-0000-4000-8000-000000000004";
const OP1 = "5e7e0000-0000-4000-8000-000000000011";
const OP2 = "5e7e0000-0000-4000-8000-000000000012";
const K = "4b4b0000-0000-4000-8000-000000000021";
const TTL_MS = 90_000;
const BRIDGE_SESSION = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const BRIDGE_PANE = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
const TICKET = "cccccccc-cccc-4ccc-8ccc-cccccccccccc";
/** Letter-bearing canonical ids (so case variants differ from the genuine id). */
const OP_L = "5e7eabcd-ef01-4a2b-8c3d-4e5f6a7b8c9d";
const KEY_L = "4b4bcdef-0123-4abc-9def-0123456789ab";

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
const SUMMARY = {
  rejected: "The workbook plan was rejected before execution; nothing was executed.",
} as const;
const REJECTED_BODY = JSON.stringify({ ok: false, action: "workbook_plan", reason: "rejected", summary: SUMMARY.rejected });
const PHASE_REFUSAL = { ok: false, reason: "Trusted server turn phase is not permitted." };
const INVALID_TICKET_RESULT = { ok: false, action: "workbook_plan", reason: "invalid", error: "Workbook plan ticket was invalid and was not executed." };

const v2 = (op: string, phase: string, turnKey: string = K): Record<string, unknown> =>
  ({ turnOperationId: op, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT, phase, turnKey });

const inspectionBody = (reads: unknown[]): string => JSON.stringify({
  ok: true,
  action: "workbook_plan",
  phase: "inspect",
  result: { outcome: "inspected", inspectionCount: 1, reads },
});
const bytes = (text: string): number => new TextEncoder().encode(text).length;

// ---------------------------------------------------------------------------
// Route/coordinator harness (COPIED shape; not imported)
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

async function harness(rawInspection: string) {
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
  const inspect = vi.fn(async (_reads: unknown, _signal?: AbortSignal) => {
    events.push("inspect-read");
    return rawInspection;
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
    serverInspectionSupported: () => true,
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
  return { events, cells, inspect, acks, run, bound, stateOf, count };
}

const WORKBOOK_TOUCHES = ["runtime-read", "store-claim", "fresh", "mutation", "verify-read", "ack"];
function expectNoWriteAccess(h: Awaited<ReturnType<typeof harness>>): void {
  for (const touch of WORKBOOK_TOUCHES) expect(h.count(touch), touch).toBe(0);
  expect(h.cells.size).toBe(0);
}

async function expectAccepted(raw: string, reads: unknown[]): Promise<void> {
  const h = await harness(raw);
  const body = await h.run(INSPECT_PLAN, v2(OP1, "inspect"));
  expect(body).toBe(inspectionBody(reads));
  expect(h.inspect).toHaveBeenCalledTimes(1);
  expect(h.stateOf(0)?.phase).toBe("inspected");
  expectNoWriteAccess(h);
}

async function expectFailClosed(raw: string): Promise<void> {
  const h = await harness(raw);
  expect(await h.run(INSPECT_PLAN, v2(OP1, "inspect"))).toBe(REJECTED_BODY);
  expect(h.inspect).toHaveBeenCalledTimes(1);
  expect(h.stateOf(0)?.phase).toBe("terminal");
  expect(await h.run(EXECUTE_PLAN, v2(OP2, "execute"))).toBe(REJECTED_BODY);
  expect(h.bound(1)).toEqual(PHASE_REFUSAL);
  expectNoWriteAccess(h);
}

const rawOf = (reads: unknown[]): string => JSON.stringify({ ok: true, reads });
/** A read whose compact JSON is exactly `size` UTF-8 bytes (ASCII padding). */
const readOfBytes = (size: number): { ok: boolean; data: string } => {
  const read = { ok: true, data: "x".repeat(size - bytes(JSON.stringify({ ok: true, data: "" }))) };
  expect(bytes(JSON.stringify(read))).toBe(size);
  return read;
};

// ---------------------------------------------------------------------------
// IR: per-read inspection shape and size at the route
// ---------------------------------------------------------------------------
describe("R4g-2b IR: the route builds the T8 inspection body only from conforming, bounded reads; anything else fails closed", () => {
  it("IR0 control: conforming reads ({ok, data}, ok boolean, in any key order) give the exact T8 body and keep the shared state inspected", async () => {
    const reads = [{ ok: true, data: [["v"]] }, { data: null, ok: false }, { ok: true, data: { values: [[1, "a"]], formulas: [["=A1"]] } }];
    await expectAccepted(rawOf(reads), reads);
  });

  it.each([
    ["a string read (canary)", rawOf(["canary-string-read"])],
    ["a number read", rawOf([7])],
    ["a null read", rawOf([null])],
    ["a boolean read", rawOf([true])],
    ["an array read", rawOf([[{ ok: true, data: 1 }]])],
    ["a read without ok", rawOf([{ data: [["v"]] }])],
    ["a read whose ok is a string", rawOf([{ ok: "true", data: [["v"]] }])],
    ["a read whose ok is a number", rawOf([{ ok: 1, data: [["v"]] }])],
    ["a read without data", rawOf([{ ok: true }])],
    ["a read with an extra formulas key", rawOf([{ ok: true, data: [["v"]], formulas: [["=SECRET()"]] }])],
    ["an error-shaped read {ok:false, error}", rawOf([{ ok: false, error: "Invalid inspection result" }])],
    ["a read with an own __proto__ key", "{\"ok\":true,\"reads\":[{\"ok\":true,\"data\":1,\"__proto__\":{\"x\":1}}]}"],
    ["a read with a constructor key", rawOf([{ ok: true, data: 1, constructor: "x" }])],
    ["one conforming and one non-conforming read", rawOf([{ ok: true, data: 1 }, { ok: true, data: 1, extra: 0 }])],
    ["a raw read_range tool result (no ok/data envelope)", rawOf([{ sheet: "Sheet1", address: "Sheet1!A1", rowCount: 1, columnCount: 1, values: [["v"]], formulas: [["v"]], numberFormat: [["General"]] }])],
  ])("IR1 %s: the fixed rejected body, the shared state terminal, the following execute refused", async (_label, raw) => {
    await expectFailClosed(raw);
  });

  it("IR2 per-read cap: a read of exactly READ_CAP bytes is accepted; one byte more fails closed", async () => {
    const atCap = readOfBytes(READ_CAP);
    await expectAccepted(rawOf([atCap]), [atCap]);
    await expectFailClosed(rawOf([readOfBytes(READ_CAP + 1)]));
  });

  it("IR2b the per-read cap counts UTF-8 bytes, not UTF-16 code units", async () => {
    const base = bytes(JSON.stringify({ ok: true, data: "" }));
    const read = { ok: true, data: "ä".repeat(Math.floor((READ_CAP - base) / 2) + 1) };
    expect(JSON.stringify(read).length).toBeLessThan(READ_CAP);
    expect(bytes(JSON.stringify(read))).toBeGreaterThan(READ_CAP);
    await expectFailClosed(rawOf([read]));
  });

  it("IR3 body cap: reads each within READ_CAP whose body is exactly BODY_CAP are accepted; one byte more fails closed", async () => {
    const first = readOfBytes(READ_CAP);
    const overhead = bytes(inspectionBody([first, { ok: true, data: "" }]));
    const second = { ok: true, data: "x".repeat(BODY_CAP - overhead) };
    expect(bytes(JSON.stringify(second))).toBeLessThanOrEqual(READ_CAP);
    expect(bytes(inspectionBody([first, second]))).toBe(BODY_CAP);
    await expectAccepted(rawOf([first, second]), [first, second]);
    const over = { ok: true, data: "x".repeat(BODY_CAP - overhead + 1) };
    expect(bytes(inspectionBody([first, over]))).toBe(BODY_CAP + 1);
    await expectFailClosed(rawOf([first, over]));
  });

  it("IR4 static: the caps sit well below the 10 MiB result cap and the gate constant is now true", () => {
    expect(BODY_CAP).toBeLessThan(10 * 1024 * 1024 / 4);
    expect(READ_CAP).toBeLessThan(BODY_CAP);
    expect((routeModule as Record<string, unknown>).SERVER_INSPECTION_SUPPORTED).toBe(true);
    const route = readFileSync(new URL("../taskpane/serverOwnedPlanRoute.ts", import.meta.url), "utf8");
    expect(route).toContain("1_048_576");
    expect(route).toContain("1_572_864");
  });
});

// ---------------------------------------------------------------------------
// UK: canonical UUIDs at the pane parser
// ---------------------------------------------------------------------------
function response(status: number, body?: unknown): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
  });
}

function stubBridge(ticket: Record<string, unknown>, executeTool: (...args: unknown[]) => Promise<string>) {
  const posts: Record<string, unknown>[] = [];
  let executeStarts = 0;
  let delivered = false;
  const fetchImpl = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    const method = init?.method ?? "GET";
    if (method === "POST" && url === "/excel/session") {
      return response(201, {
        ok: true, sessionId: BRIDGE_SESSION, sessionKey: "s".repeat(64),
        protocolVersion: "1.0", expiresInSeconds: 60, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
      });
    }
    if (method === "GET" && url === "/excel/pending") {
      if (delivered) return response(204);
      delivered = true;
      return response(200, ticket);
    }
    if (method === "POST" && url === "/excel/execute-start") { executeStarts += 1; return response(200, { ok: true }); }
    if (method === "POST" && url === "/excel/result") {
      posts.push(JSON.parse(String(init?.body)) as Record<string, unknown>);
      return response(200, { ok: true });
    }
    throw new Error(`unexpected fetch ${method} ${url}`);
  });
  const bridge = new OfficeCompanionBridge({
    hostPrefix: "excel", documentId: async () => "doc-1", paneId: BRIDGE_PANE,
    executeTool: executeTool as ConstructorParameters<typeof OfficeCompanionBridge>[0]["executeTool"],
    fetchImpl: fetchImpl as unknown as typeof fetch,
    setIntervalImpl: (() => 1) as unknown as typeof setInterval,
    clearIntervalImpl: (() => undefined) as unknown as typeof clearInterval,
  });
  return { bridge, posts, executeStarts: () => executeStarts };
}

const planTicket = (trusted: Record<string, unknown>, plan: Record<string, unknown> = { ...ASK_PLAN }): Record<string, unknown> =>
  ({ ticketId: TICKET, requestId: OP_L, tool: "excel_execute_plan", arguments: plan, trusted });

const NONCANONICAL = (value: string): [string, string][] => [
  ["uppercase", value.toUpperCase()],
  ["one uppercase letter", `${value.slice(0, -1)}${value.slice(-1).toUpperCase()}`],
  ["braces", `{${value}}`],
  ["urn:uuid: prefix", `urn:uuid:${value}`],
  ["no dashes", value.replace(/-/g, "")],
  ["nil", "00000000-0000-0000-0000-000000000000"],
];

describe("R4g-2b UK: the pane parser accepts turnKey and turnOperationId only as canonical lowercase hyphenated UUIDs", () => {
  it("UK0 control: canonical lowercase ids reach executeTool once (own v2 envelope, gate on)", async () => {
    expect(OP_L.toUpperCase()).not.toBe(OP_L);
    expect(KEY_L.slice(-1)).toMatch(/[a-f]/);
    const executeTool = vi.fn(async (..._args: unknown[]) => JSON.stringify({ ok: false, action: "workbook_plan", reason: "rejected", summary: SUMMARY.rejected }));
    const { bridge, posts } = stubBridge(planTicket(v2(OP_L, "execute", KEY_L)), executeTool);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(executeTool.mock.calls[0]![2]).toEqual(v2(OP_L, "execute", KEY_L));
    expect(posts).toHaveLength(1);
  });

  it.each([
    ...NONCANONICAL(KEY_L).map(([label, form]): [string, Record<string, unknown>] => [`turnKey ${label}`, v2(OP_L, "execute", form)]),
    ...NONCANONICAL(OP_L).map(([label, form]): [string, Record<string, unknown>] => [`turnOperationId ${label}`, v2(form, "execute", KEY_L)]),
    ["turnKey = the op id uppercased (same UUID value)", v2(OP_L, "execute", OP_L.toUpperCase())],
  ])("UK1 %s: an invalid ticket (executed:false/invalid), executeTool 0, no execute-start", async (_label, trusted) => {
    const executeTool = vi.fn();
    const { bridge, posts, executeStarts } = stubBridge(planTicket(trusted), executeTool);
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(executeTool).not.toHaveBeenCalled();
    expect(executeStarts()).toBe(0);
    expect(posts).toEqual([{ ticketId: TICKET, executed: false, result: INVALID_TICKET_RESULT }]);
  });
});
