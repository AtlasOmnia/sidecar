import { describe, expect, it, vi } from "vitest";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import type { BridgePlanExecutionControl } from "../officeBridge";
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
 * R4g-2c pre-flip inspection hardening (dev3-r4g2c-release.md items 4-6; R4g-2b lane-B A4/A6, lane-A F3):
 * - H4: the REAL route refuses an inspection whose `reads` list is empty, fail closed exactly like the
 *   R4g-1 R14 malformed-read path (fixed `rejected` body, shared state terminal, following execute refused).
 * - H5: the route's phase turnKey check is case-sensitive lowercase canonical (as the pane parser's
 *   isCanonicalUuid and the companion's regex); an uppercase or mixed-case turnKey is the fixed `invalid`
 *   body before any binding, read or write.
 * - H6: each read nests at most 64 levels (provisional cap; the read object is level 1). Deeper reads,
 *   including lane-B's 5,000-level read, fail closed.
 * Driven through the real coordinator/binding store with the test-only gate override on
 * (SERVER_INSPECTION_SUPPORTED itself stays false). The route/coordinator harness is COPIED from
 * server-inspection-bounds-r4g2b.test.ts (that file is neither imported nor edited here).
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
const SUMMARY = {
  rejected: "The workbook plan was rejected before execution; nothing was executed.",
} as const;
const REJECTED_BODY = JSON.stringify({ ok: false, action: "workbook_plan", reason: "rejected", summary: SUMMARY.rejected });
const PHASE_REFUSAL = { ok: false, reason: "Trusted server turn phase is not permitted." };

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

const DEPTH_CAP = 64;
const INVALID_BODY = JSON.stringify({ ok: false, action: "workbook_plan", reason: "invalid", summary: "The workbook plan ticket was invalid; nothing was executed." });
/** `levels` nested arrays (`[]` is one level). */
const nestedArrays = (levels: number): unknown => {
  let value: unknown = [];
  for (let level = 1; level < levels; level += 1) value = [value];
  return value;
};
/** `levels` nested objects (`{}` is one level). */
const nestedObjects = (levels: number): unknown => {
  let value: unknown = {};
  for (let level = 1; level < levels; level += 1) value = { a: value };
  return value;
};
/** Nesting depth with the read object itself as level 1 (independent of the route; iterative). */
const depthOf = (root: unknown): number => {
  let max = 0;
  const stack: [unknown, number][] = [[root, 1]];
  while (stack.length > 0) {
    const [value, level] = stack.pop()!;
    if (value === null || typeof value !== "object") continue;
    max = Math.max(max, level);
    for (const child of Array.isArray(value) ? value : Object.values(value as Record<string, unknown>)) stack.push([child, level + 1]);
  }
  return max;
};

// ---------------------------------------------------------------------------
// H4 (item 4, R4g-2b lane-B A6): zero reads never unlock the execute phase
// ---------------------------------------------------------------------------
describe("R4g-2c H4: an inspection with zero reads fails closed at the route", () => {
  it("H4a control: one conforming read is still the exact T8 body with the shared state inspected", async () => {
    const reads = [{ ok: true, data: [["v"]] }];
    await expectAccepted(rawOf(reads), reads);
  });

  it("H4b {ok:true, reads:[]}: the fixed rejected body, the shared state terminal, the following execute refused", async () => {
    await expectFailClosed(rawOf([]));
  });

  it("H4c an empty reads list next to an extra top-level key still fails closed", async () => {
    await expectFailClosed(JSON.stringify({ ok: true, reads: [], note: "x" }));
  });
});

// ---------------------------------------------------------------------------
// H5 (item 5, R4g-2b lane-A F3): the route's phase turnKey is canonical lowercase
// ---------------------------------------------------------------------------
describe("R4g-2c H5: the route accepts a phase turnKey only as a canonical lowercase UUID", () => {
  const KEY_UPPER = K.toUpperCase();
  const KEY_ONE_UPPER = `${K.slice(0, 1)}${K.slice(1, 2).toUpperCase()}${K.slice(2)}`;

  it("H5a control: the lowercase key admits the execute phase and commits once", async () => {
    expect(KEY_UPPER).not.toBe(K);
    expect(KEY_ONE_UPPER).not.toBe(K);
    const h = await harness(rawOf([{ ok: true, data: 1 }]));
    const body = JSON.parse(await h.run(EXECUTE_PLAN, v2(OP1, "execute"))) as { ok?: unknown };
    expect(body.ok).toBe(true);
    expect(h.count("mutation")).toBe(1);
  });

  it.each([
    ["execute, uppercase turnKey", EXECUTE_PLAN, "execute", KEY_UPPER],
    ["execute, one uppercase letter", EXECUTE_PLAN, "execute", KEY_ONE_UPPER],
    ["inspect, uppercase turnKey", INSPECT_PLAN, "inspect", KEY_UPPER],
    ["inspect, one uppercase letter", INSPECT_PLAN, "inspect", KEY_ONE_UPPER],
  ] as const)("H5b %s: the fixed invalid body before any binding, read or write", async (_label, plan, phase, key) => {
    const h = await harness(rawOf([{ ok: true, data: 1 }]));
    expect(await h.run(plan, v2(OP1, phase, key))).toBe(INVALID_BODY);
    expect(h.bound(0)).toBeUndefined();
    expect(h.inspect).not.toHaveBeenCalled();
    expectNoWriteAccess(h);
  });
});

// ---------------------------------------------------------------------------
// H6 (item 6, R4g-2b lane-B A4): per-read nesting depth
// ---------------------------------------------------------------------------
describe("R4g-2c H6: the route bounds each read's nesting depth (provisional cap 64) and fails closed beyond it", () => {
  it("H6a a read nested exactly DEPTH_CAP levels (arrays, objects) is accepted", async () => {
    for (const nest of [nestedArrays, nestedObjects]) {
      const read = { ok: true, data: nest(DEPTH_CAP - 1) };
      expect(depthOf(read)).toBe(DEPTH_CAP);
      await expectAccepted(rawOf([read]), [read]);
    }
  });

  it.each([
    ["arrays, DEPTH_CAP + 1", () => JSON.stringify({ ok: true, data: nestedArrays(DEPTH_CAP) }), DEPTH_CAP + 1],
    ["objects, DEPTH_CAP + 1", () => JSON.stringify({ ok: true, data: nestedObjects(DEPTH_CAP) }), DEPTH_CAP + 1],
    // Build deep wire JSON without recursive stringify: Node 22 on Windows can
    // exhaust its stack in the test before the production depth guard runs.
    ["arrays, 5,000 levels (lane-B A4)", () => `{"ok":true,"data":${"[".repeat(4_999)}${"]".repeat(4_999)}}`, 5_000],
    ["objects, 5,000 levels", () => `{"ok":true,"data":${'{"a":'.repeat(4_998)}{}${"}".repeat(4_998)}}`, 5_000],
  ] as const)("H6b %s: the fixed rejected body, the shared state terminal, the following execute refused", async (_label, makeRaw, depth) => {
    const rawRead = makeRaw();
    expect(depthOf(JSON.parse(rawRead))).toBe(depth);
    expect(bytes(rawRead)).toBeLessThan(READ_CAP);
    await expectFailClosed(`{"ok":true,"reads":[${rawRead}]}`);
  });

  it("H6c one shallow read next to one over-deep read fails closed (every read is checked)", async () => {
    await expectFailClosed(rawOf([{ ok: true, data: 1 }, { ok: false, data: nestedArrays(DEPTH_CAP) }]));
  });
});
