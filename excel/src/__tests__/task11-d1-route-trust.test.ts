import { describe, expect, it, vi } from "vitest";
import { readFileSync } from "node:fs";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import * as workbookPlanOrchestrator from "../workbookPlanOrchestrator";
import {
  createWorkbookPlanTurn,
  ServerOwnedTurnBindingStore,
  WorkbookPlanTurnState,
  type ServerOwnedPaneTurn,
  type ServerOwnedTicketIdentity,
  type WorkbookPlanTurnDependencies,
} from "../workbookPlanOrchestrator";
import { createTrustedOperationBoundary, BrowserIdempotencyStore, MemoryIdempotencyStorage, type WorkbookPlanExecutionAdapter } from "../workbookPlanExecutor";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import type { WorkbookPlanPolicyInput } from "../workbookPlanPolicy";
import type { TrustedWorkbookRequest } from "../workbookPlanRequest";
import { preflightWorkbookPlanResources, type WorkbookPlanRouteRuntimeContext } from "../workbookPlanRoute";
import type { WorkbookPlanResourceContext } from "../workbookPlanResources";
import { readNativeWorkbookPlanRuntimeContext } from "../workbookPlanOfficeAdapter";
import {
  TaskPaneTurnCoordinator,
  type TaskPaneServerAuthorization,
  type TaskPaneTurnCoordinatorDependencies,
} from "../taskpane/taskPaneTurnCoordinator";

const ACTIVE: ServerOwnedPaneTurn = {
  sessionId: "11111111-1111-4111-8111-111111111111",
  documentId: "22222222-2222-4222-8222-222222222222",
  paneId: "33333333-3333-4333-8333-333333333333",
  turnId: "44444444-4444-4444-8444-444444444444",
};
const SERVER_OPERATION_ID = "55555555-5555-4555-8555-555555555555";
const COMPILE_CONTEXT: VerifiedWorkbookCompileContext = {
  sheetNames: new Set(["Sheet1"]),
  tablesByName: new Map(),
  supportedPrimitiveNames: new Set(["write_range"]),
  supportedCapabilities: new Set(["excel_workbook_access", "workbook_plan_v1_operation_audit"]),
  chartSourceResolutionsBySheet: new Map(),
};
const EMPTY_RESOURCE_CONTEXT: WorkbookPlanResourceContext = {
  compileContext: {
    sheetNames: new Set(),
    tablesByName: new Map(),
    supportedPrimitiveNames: new Set(),
    supportedCapabilities: new Set(),
    chartSourceResolutionsBySheet: new Map(),
  },
};
const LIVE_RESOURCE_RUNTIME: WorkbookPlanRouteRuntimeContext = {
  activeSheetName: "Sheet1",
  selection: { sheetName: "Sheet1", address: "A1" },
  compileContext: {
    ...COMPILE_CONTEXT,
    tablesByName: new Map([["Orders", { canonicalName: "Orders", columnCount: 4, rowCount: 1_000_000 }]]),
  },
};
const POLICY: WorkbookPlanPolicyInput = {
  requestClassification: "explicit_mutation",
  activeSheetName: "Sheet1",
  selectionScopedRequest: false,
  operationSummaries: [],
  capabilitySummaries: ["excel_workbook_access", "workbook_plan_v1_operation_audit"],
  resourceUsage: { operationCount: 0, cellsRead: 1, cellsWritten: 0, formulaUtf8Bytes: 0, csvPayloadUtf8Bytes: 0, commentUtf8Bytes: 0, sheetsCreated: 0, chartsCreated: 0, tablesCreated: 0, verificationReads: 1 },
  resourceLimits: { operationCount: 10, cellsRead: 10, cellsWritten: 10, formulaUtf8Bytes: 1000, csvPayloadUtf8Bytes: 1000, commentUtf8Bytes: 1000, sheetsCreated: 1, chartsCreated: 1, tablesCreated: 1, verificationReads: 10 },
};
const INSPECT_RESPONSE = {
  version: 1,
  outcome: "inspect",
  reason: "need context",
  reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1" }],
} as const;

function ticket(overrides: Partial<ServerOwnedTicketIdentity> = {}): ServerOwnedTicketIdentity {
  return {
    ...ACTIVE,
    serverOperationId: SERVER_OPERATION_ID,
    ...overrides,
  };
}

type AuthorizedServerOperation = Extract<TaskPaneServerAuthorization, { readonly ok: true }>;

/** Reads a source file that may not exist yet (R2d-1 extraction); missing reads as "" so assertions fail. */
function readOptionalSource(relative: string): string {
  try {
    return readFileSync(new URL(relative, import.meta.url), "utf8");
  } catch {
    return "";
  }
}

function createDownstreamSpies() {
  const spies = {
    agentSend: vi.fn(),
    normalize: vi.fn(),
    runtime: vi.fn(),
    inspect: vi.fn(),
    trustedPreparation: vi.fn(),
    adapter: vi.fn(),
    office: vi.fn(),
  };
  return {
    spies,
    run: () => Object.values(spies).forEach((spy) => spy()),
  };
}

describe("Task 11 D1 trusted server ticket binding", () => {
  it("binds a first server UUIDv5 ticket to the active pane turn without comparing direct UUIDs", () => {
    const store = new ServerOwnedTurnBindingStore({ now: () => 1_000, ttlMs: 10_000, maxEntries: 4 });
    store.activateTurn(ACTIVE);

    const first = store.bindServerTicket(ticket());
    const exactReuse = store.bindServerTicket(ticket());

    expect(first).toMatchObject({ ok: true, reused: false });
    expect(exactReuse).toMatchObject({ ok: true, reused: true });
    expect(first.ok && first.binding.activeTurnId).toBe(ACTIVE.turnId);
    expect(first.ok && first.binding.serverOperationId).toBe(SERVER_OPERATION_ID);
    expect(ACTIVE.turnId).not.toBe(SERVER_OPERATION_ID);
  });

  it("rejects tickets before/after the active turn and all mismatched identities before document/runtime access", () => {
    const store = new ServerOwnedTurnBindingStore({ now: () => 1_000, ttlMs: 10_000, maxEntries: 4 });
    store.activateTurn(ACTIVE);

    const before = new ServerOwnedTurnBindingStore({ now: () => 1_000, ttlMs: 10_000, maxEntries: 4 });
    expect(before.bindServerTicket(ticket())).toMatchObject({ ok: false });

    for (const mismatch of [
      { sessionId: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa" },
      { documentId: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb" },
      { paneId: "cccccccc-cccc-4ccc-8ccc-cccccccccccc" },
      { turnId: "dddddddd-dddd-4ddd-8ddd-dddddddddddd" },
    ]) {
      expect(store.bindServerTicket(ticket(mismatch))).toMatchObject({ ok: false });
    }

    expect(store.bindServerTicket(ticket())).toMatchObject({ ok: true });
    expect(store.bindServerTicket(ticket({ serverOperationId: "66666666-6666-4666-8666-666666666666" }))).toMatchObject({ ok: false });
    store.finishTerminal(ticket());
    expect(store.bindServerTicket(ticket())).toMatchObject({ ok: false });

    store.clearTurn(ACTIVE);
    expect(store.bindServerTicket(ticket())).toMatchObject({ ok: false });
  });

  it("retains a content-free replay tombstone across finalization and Stop", () => {
    const turnB: ServerOwnedPaneTurn = {
      ...ACTIVE,
      turnId: "66666666-6666-4666-8666-666666666666",
    };
    const store = new ServerOwnedTurnBindingStore({ now: () => 1_000, ttlMs: 10_000, maxEntries: 4 });
    store.activateTurn(ACTIVE);
    expect(store.bindServerTicket(ticket())).toMatchObject({ ok: true });
    store.finishTerminal(ticket());
    store.clearTurn(ACTIVE);
    store.activateTurn(turnB);
    expect(store.bindServerTicket({ ...turnB, serverOperationId: SERVER_OPERATION_ID })).toMatchObject({ ok: false });

    const stopped = new ServerOwnedTurnBindingStore({ now: () => 1_000, ttlMs: 10_000, maxEntries: 4 });
    stopped.activateTurn(ACTIVE);
    expect(stopped.bindServerTicket(ticket())).toMatchObject({ ok: true });
    stopped.stop();
    stopped.activateTurn(turnB);
    expect(stopped.bindServerTicket({ ...turnB, serverOperationId: SERVER_OPERATION_ID })).toMatchObject({ ok: false });
  });

  it("retains replay protection through the horizon, expires it only afterward, and never evicts an active entry", () => {
    let now = 1_000;
    const store = new ServerOwnedTurnBindingStore({ now: () => now, ttlMs: 100, maxEntries: 1 });
    store.activateTurn(ACTIVE);
    expect(store.bindServerTicket(ticket())).toMatchObject({ ok: true });
    store.clearTurn(ACTIVE);

    const turnB: ServerOwnedPaneTurn = {
      ...ACTIVE,
      turnId: "77777777-7777-4777-8777-777777777777",
    };
    store.activateTurn(turnB);
    now = 1_099;
    expect(store.bindServerTicket({ ...turnB, serverOperationId: SERVER_OPERATION_ID })).toMatchObject({ ok: false });
    now = 1_100;
    expect(store.bindServerTicket({ ...turnB, serverOperationId: SERVER_OPERATION_ID })).toMatchObject({ ok: true });

    now = 2_000;
    const activeStore = new ServerOwnedTurnBindingStore({ now: () => now, ttlMs: 100, maxEntries: 1 });
    activeStore.activateTurn(ACTIVE);
    expect(activeStore.bindServerTicket(ticket())).toMatchObject({ ok: true });
    const otherPane: ServerOwnedPaneTurn = {
      ...ACTIVE,
      paneId: "88888888-8888-4888-8888-888888888888",
      turnId: "99999999-9999-4999-8999-999999999999",
    };
    activeStore.activateTurn(otherPane);
    expect(activeStore.bindServerTicket({ ...otherPane, serverOperationId: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa" })).toMatchObject({ ok: false });
  });

  it("bounds replay tombstones by bytes without evicting the existing operation", () => {
    const store = new ServerOwnedTurnBindingStore({
      now: () => 1_000,
      ttlMs: 10_000,
      maxEntries: 4,
      maxBytes: 1,
    } as ConstructorParameters<typeof ServerOwnedTurnBindingStore>[0]);
    store.activateTurn(ACTIVE);
    expect(store.bindServerTicket(ticket())).toMatchObject({ ok: false });
  });

  it("rejects a server-owned inspect outcome before primitive inspection, runtime capture, or adapter access", async () => {
    const calls: string[] = [];
    let documentReads = 0;
    const turnState = new WorkbookPlanTurnState();
    const dependencies: WorkbookPlanTurnDependencies = {
      identity: {
        sessionId: ACTIVE.sessionId,
        documentId: ACTIVE.documentId,
        turnId: ACTIVE.turnId,
      },
      trustedOperationBoundary: createTrustedOperationBoundary({
        sessionId: ACTIVE.sessionId,
        documentId: ACTIVE.documentId,
        turnId: ACTIVE.turnId,
      }),
      policyInput: POLICY,
      store: new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      model: async () => { calls.push("model"); throw new Error("server model continuation must not run"); },
      initialResponse: {
        content: "",
        toolCalls: [{ id: "call-1", type: "function", function: { name: "excel_execute_plan", arguments: JSON.stringify(INSPECT_RESPONSE) } }],
      },
      inspect: async () => { calls.push("inspect"); return "{}"; },
      getCompileContextForPlan: async () => { documentReads += 1; calls.push("runtime"); return COMPILE_CONTEXT; },
      resourcePreflight: async () => { calls.push("preflight"); return { ok: true }; },
      adapter: {
        executeFirstGroup: async () => { calls.push("adapter"); return { status: "committed" }; },
        executeGroup: async () => { calls.push("adapter-extra"); return { status: "committed" }; },
        reconcileFailedGroup: async () => [],
      },
      route: "server-owned",
      trustedTurnOperationId: ACTIVE.turnId,
      turnState,
      continueAfterInspection: false,
      rolloutEnabled: true,
    };

    const result = await createWorkbookPlanTurn(dependencies);

    expect(result.outcome).toBe("rejected");
    expect(calls).toEqual([]);
    expect(documentReads).toBe(0);
    expect(turnState.phase).toBe("terminal");
  });

  it.each(["direct", "server-owned"] as const)("R7-B2 %s rejects unknown-table then oversized range BEFORE native context capture", async (route) => {
    const payloadLoads: string[] = [];
    const range = (address: string) => ({
      address: `Sheet1!${address}`, rowCount: address === "A1:XFD1048576" ? 1_048_576 : 1,
      columnCount: address === "A1:XFD1048576" ? 16_384 : 1,
      values: [[null]], formulas: [[""]], numberFormat: [["General"]],
      load: (fields: string | readonly string[]) => {
        if (Array.isArray(fields) && fields.includes("values")) payloadLoads.push(address);
      },
    }); // Synthetic load recorder; never allocate a whole-sheet value matrix.
    const sheet = { name: "Sheet1", isNullObject: false, load: vi.fn(), getRange: range, getUsedRangeOrNullObject: () => ({ ...range("A1"), isNullObject: false }) };
    const table = { name: "Table1", isNullObject: false, load: vi.fn(), columns: { items: [{ name: "A" }], load: vi.fn() }, getRange: () => range("A1") };
    const context = {
      workbook: {
        worksheets: { items: [sheet], load: vi.fn(), getActiveWorksheet: () => sheet, getItem: () => sheet, getItemOrNullObject: () => sheet },
        tables: { items: [table], load: vi.fn(), getItemOrNullObject: () => table },
      }, sync: vi.fn(async () => undefined),
    };
    const officeRun = vi.fn(async (work: (value: unknown) => Promise<unknown>) => work(context));
    const previous = Object.getOwnPropertyDescriptor(globalThis, "Excel");
    Object.defineProperty(globalThis, "Excel", { configurable: true, writable: true, value: { run: officeRun } });
    const plan = {
      version: 1, outcome: "execute", summary: "Sort then format",
      operations: [
        { kind: "sort_table", tableName: "Table1", columnIndex: 0, ascending: true },
        { kind: "format_range", sheet: "Sheet1", address: "A1:XFD1048576", format: { bold: true } },
      ], verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "B1", expected: null }],
    } as const;
    let live: WorkbookPlanRouteRuntimeContext = LIVE_RESOURCE_RUNTIME;
    const mutation = vi.fn(async () => ({ status: "committed" as const }));
    try {
      const result = await createWorkbookPlanTurn({
        identity: { sessionId: ACTIVE.sessionId, documentId: ACTIVE.documentId, turnId: ACTIVE.turnId },
        trustedOperationBoundary: createTrustedOperationBoundary({ sessionId: ACTIVE.sessionId, documentId: ACTIVE.documentId, turnId: ACTIVE.turnId }),
        policyInput: POLICY, store: new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
        model: async () => { throw new Error("initial response should be consumed"); },
        initialResponse: { content: "", toolCalls: [{ id: "call-r7b2", type: "function", function: { name: "excel_execute_plan", arguments: JSON.stringify(plan) } }] },
        inspect: async () => "{}", resourcePreflight: async (candidate) => preflightWorkbookPlanResources(candidate, EMPTY_RESOURCE_CONTEXT),
        getCompileContextForPlan: async (candidate) => { live = await readNativeWorkbookPlanRuntimeContext(candidate); return live.compileContext; },
        resourcePreflightAfterRuntime: async (candidate) => preflightWorkbookPlanResources(candidate, live),
        adapter: { executeFirstGroup: mutation, executeGroup: mutation, reconcileFailedGroup: async () => [] },
        rolloutEnabled: true, route, trustedTurnOperationId: ACTIVE.turnId,
      });
      expect(result.outcome).toBe("rejected");
      expect(payloadLoads).toEqual([]); // RED on 0bd33bf queues the whole-sheet load.
      expect(officeRun).not.toHaveBeenCalled();
      expect(mutation).not.toHaveBeenCalled();
    } finally {
      if (previous) Object.defineProperty(globalThis, "Excel", previous); else delete (globalThis as Record<string, unknown>).Excel;
    }
  });

  it("recomputes exact resources after live runtime capture and stops before preparation or mutation", async () => {
    const calls: string[] = [];
    const resourceResults: boolean[] = [];
    const dependencies: WorkbookPlanTurnDependencies = {
      identity: {
        sessionId: ACTIVE.sessionId,
        documentId: ACTIVE.documentId,
        turnId: ACTIVE.turnId,
      },
      trustedOperationBoundary: createTrustedOperationBoundary({
        sessionId: ACTIVE.sessionId,
        documentId: ACTIVE.documentId,
        turnId: ACTIVE.turnId,
      }),
      policyInput: POLICY,
      store: new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      model: async () => ({
        content: "",
        toolCalls: [{ id: "call-1", type: "function", function: { name: "excel_execute_plan", arguments: JSON.stringify({
          version: 1,
          outcome: "execute",
          summary: "bounded write",
          operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }],
          verification: [{ kind: "table_row_order", tableName: "Orders", expectedOrder: ["1"] }],
        }) } }],
      }),
      inspect: async () => "{}",
      resourcePreflight: async (plan: Parameters<typeof preflightWorkbookPlanResources>[0]) => {
        calls.push("empty");
        const result = preflightWorkbookPlanResources(plan, EMPTY_RESOURCE_CONTEXT);
        resourceResults.push(result.ok);
        return result;
      },
      getCompileContextForPlan: async () => { calls.push("runtime"); return COMPILE_CONTEXT; },
      resourcePreflightAfterRuntime: async (plan: Parameters<typeof preflightWorkbookPlanResources>[0]) => {
        calls.push("live");
        const result = preflightWorkbookPlanResources(plan, LIVE_RESOURCE_RUNTIME);
        resourceResults.push(result.ok);
        return result;
      },
      adapter: {
        executeFirstGroup: async () => { calls.push("adapter"); return { status: "committed" as const }; },
        executeGroup: async () => { calls.push("adapter-extra"); return { status: "committed" as const }; },
        reconcileFailedGroup: async () => [],
        verifyPlan: async () => true,
      },
      rolloutEnabled: true,
    };

    const result = await createWorkbookPlanTurn(dependencies);

    expect(result.outcome).toBe("rejected");
    expect(calls).toEqual(["empty", "runtime", "live"]);
    expect(resourceResults).toEqual([true, false]);
  });

  it("rejects a direct executable route when live resource authority is omitted after runtime capture", async () => {
    const calls: string[] = [];
    let adapterCalls = 0;
    const plan = {
      version: 1,
      outcome: "execute",
      summary: "bounded write",
      operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "B1", expected: null }],
    } as const;
    const result = await createWorkbookPlanTurn({
      identity: { sessionId: ACTIVE.sessionId, documentId: ACTIVE.documentId, turnId: ACTIVE.turnId },
      trustedOperationBoundary: createTrustedOperationBoundary({
        sessionId: ACTIVE.sessionId,
        documentId: ACTIVE.documentId,
        turnId: ACTIVE.turnId,
      }),
      policyInput: {
        ...POLICY,
        operationSummaries: [{
          operationIndex: 0,
          effect: "write",
          target: { operationIndex: 0, sheetNames: ["Sheet1"], selectionRelationship: "inside" },
        }],
        resourceUsage: { ...POLICY.resourceUsage, operationCount: 1, cellsWritten: 1 },
      },
      store: new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      model: async () => ({
        content: "",
        toolCalls: [{ id: "call-1", type: "function", function: { name: "excel_execute_plan", arguments: JSON.stringify(plan) } }],
      }),
      inspect: async () => "{}",
      resourcePreflight: async () => { calls.push("empty"); return { ok: true as const }; },
      getCompileContextForPlan: async () => { calls.push("runtime"); return COMPILE_CONTEXT; },
      adapter: {
        executeFirstGroup: async () => { adapterCalls += 1; return { status: "committed" as const }; },
        executeGroup: async () => { adapterCalls += 1; return { status: "committed" as const }; },
        reconcileFailedGroup: async () => [],
        verifyPlan: async () => true,
      },
      route: "direct",
      trustedTurnOperationId: ACTIVE.turnId,
      rolloutEnabled: true,
    });

    expect(result.outcome).toBe("rejected");
    expect(calls).toEqual(["empty", "runtime"]);
    expect(adapterCalls).toBe(0);
  });

  it("invalidates a deferred preload continuation before a later turn can activate it", async () => {
    const Guard = (workbookPlanOrchestrator as typeof workbookPlanOrchestrator & {
      WorkbookPlanTurnActivationGuard?: new () => {
        begin(turnId: string, controller: AbortController): number;
        invalidate(): void;
        isCurrent(generation: number, turnId: string, controller: AbortController): boolean;
      };
    }).WorkbookPlanTurnActivationGuard;
    expect(typeof Guard).toBe("function");
    const guard = new Guard!();
    const controllerA = new AbortController();
    const generationA = guard.begin("turn-a", controllerA);
    let resolvePreload!: () => void;
    const preload = new Promise<void>((resolve) => { resolvePreload = resolve; });
    const lateActivation = preload.then(() => guard.isCurrent(generationA, "turn-a", controllerA));
    guard.invalidate();
    controllerA.abort();

    const controllerB = new AbortController();
    const generationB = guard.begin("turn-b", controllerB);
    resolvePreload();

    await expect(lateActivation).resolves.toBe(false);
    expect(guard.isCurrent(generationB, "turn-b", controllerB)).toBe(true);
  });

  it("couples the task-pane fingerprint guard to the generated fingerprint and rejects a guard mutation", () => {
    // R2d-1 (M2/M3): the server-owned route guard moved from taskpane.ts to the
    // extracted serverOwnedPlanRoute.ts; the same assertions now read that module.
    const source = readOptionalSource("../taskpane/serverOwnedPlanRoute.ts");
    expect(source).toMatch(/import \{ WORKBOOK_PLAN_FINGERPRINT, type WorkbookPlan \} from "\.\.\/generated\/workbookPlanSchema";/);
    expect(source).toContain("trusted.workbookPlanFingerprint !== WORKBOOK_PLAN_FINGERPRINT");
    expect(source).not.toContain(`trusted.workbookPlanFingerprint !== "${WORKBOOK_PLAN_FINGERPRINT}"`);
    const taskPaneSource = readFileSync(new URL("../taskpane/taskpane.ts", import.meta.url), "utf8");
    expect(taskPaneSource).not.toContain(`"${WORKBOOK_PLAN_FINGERPRINT}"`);

    const mutated = source.replace(
      "trusted.workbookPlanFingerprint !== WORKBOOK_PLAN_FINGERPRINT",
      "trusted.workbookPlanFingerprint !== MUTATED_WORKBOOK_PLAN_FINGERPRINT",
    );
    expect(mutated).not.toContain("trusted.workbookPlanFingerprint !== WORKBOOK_PLAN_FINGERPRINT");
  });

  it("removes the unused inspection payload while preserving the explicit server-owned fail-closed branch", () => {
    const source = readFileSync(new URL("../workbookPlanOrchestrator.ts", import.meta.url), "utf8");
    expect(source).not.toContain("let inspectionResult: string");
    expect(source).toContain("Server-owned workbook inspection is not permitted.");
    const taskPaneSource = readFileSync(new URL("../taskpane/taskpane.ts", import.meta.url), "utf8");
    expect(taskPaneSource).not.toContain("inspectionPayload");
    expect(taskPaneSource).not.toContain("workbook_plan_inspection");
    const routeSource = readOptionalSource("../taskpane/serverOwnedPlanRoute.ts");
    expect(routeSource).not.toBe("");
    expect(routeSource).not.toContain("inspectionPayload");
    expect(routeSource).not.toContain("workbook_plan_inspection");
  });

  it("fails closed before direct route host reads and rolls back to primitive tools when disabled", () => {
    const source = readFileSync(new URL("../taskpane/taskpane.ts", import.meta.url), "utf8");
    const directStart = source.indexOf("async function runDirectWorkbookPlanTurn");
    // R2d-1 (M3): executeCompanionTool moved to serverOwnedPlanRoute.ts; the
    // direct route now ends where the next taskpane function begins.
    const directEnd = source.indexOf("async function inspectWorkbookPlanReads", directStart);
    expect(directStart).toBeGreaterThanOrEqual(0);
    expect(directEnd).toBeGreaterThan(directStart);
    const directRoute = source.slice(directStart, directEnd);
    expect(directRoute.indexOf("assertWorkbookPlanRolloutEnabled")).toBeGreaterThanOrEqual(0);
    expect(directRoute.indexOf("assertWorkbookPlanRolloutEnabled")).toBeLessThan(directRoute.indexOf("preflightInitialWorkbookPlanResponse"));
    expect(directRoute.indexOf("assertWorkbookPlanRolloutEnabled")).toBeLessThan(directRoute.indexOf("getOrCreateExcelDocumentId"));

    const offProviderSource = readFileSync(new URL("../agent.ts", import.meta.url), "utf8");
    expect(offProviderSource).toContain("return settings.workbookPlanRolloutEnabled === true");
    expect(offProviderSource).toContain(": [...toolSchemas]");
  });

  it("gates the extracted server route on rollout before session, runtime or plan reads, keyed on the canonical plan tool name", () => {
    const routeSource = readOptionalSource("../taskpane/serverOwnedPlanRoute.ts");
    const serverStart = routeSource.indexOf("async function executeServerOwnedWorkbookPlan");
    expect(serverStart).toBeGreaterThanOrEqual(0);
    const serverRoute = routeSource.slice(serverStart);
    const rolloutGate = serverRoute.indexOf("isRolloutEnabled()");
    expect(rolloutGate).toBeGreaterThanOrEqual(0);
    expect(rolloutGate).toBeLessThan(serverRoute.indexOf("getBridgeSessionId()"));
    expect(rolloutGate).toBeLessThan(serverRoute.indexOf("readRuntimeContext("));
    expect(rolloutGate).toBeLessThan(serverRoute.indexOf("preflightInitialWorkbookPlanResponse("));

    expect(routeSource).toMatch(/import \{[^}]*\bWORKBOOK_PLAN_TOOL_NAME\b[^}]*\} from "\.\.\/workbookPlanProtocol";/);
    expect(routeSource).not.toContain("\"excel_execute_plan\"");
    const bridgeSource = readFileSync(new URL("../officeBridge.ts", import.meta.url), "utf8");
    expect(bridgeSource).toMatch(/import \{[^}]*\bWORKBOOK_PLAN_TOOL_NAME\b[^}]*\} from "\.\/workbookPlanProtocol";/);
    expect(bridgeSource).not.toContain("\"excel_execute_plan\"");
    const taskPaneSource = readFileSync(new URL("../taskpane/taskpane.ts", import.meta.url), "utf8");
    expect(taskPaneSource).not.toContain("name !== \"execute_plan\"");
    expect(taskPaneSource).toContain("createServerOwnedPlanRoute(");
  });

  it("rejects disabled or absent rollout before any direct or server route dependency is touched", async () => {
    for (const route of ["direct", "server-owned"] as const) {
      for (const rolloutEnabled of [false, undefined]) {
        const calls: string[] = [];
        const result = await createWorkbookPlanTurn({
          identity: { sessionId: ACTIVE.sessionId, documentId: ACTIVE.documentId, turnId: ACTIVE.turnId },
          trustedOperationBoundary: createTrustedOperationBoundary({ sessionId: ACTIVE.sessionId, documentId: ACTIVE.documentId, turnId: ACTIVE.turnId }),
          policyInput: POLICY,
          store: new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
          model: async () => { calls.push("model"); return { content: "", toolCalls: [] }; },
          inspect: async () => { calls.push("inspect"); return "{}"; },
          resourcePreflight: async () => { calls.push("preflight"); return { ok: true as const }; },
          getCompileContextForPlan: async () => { calls.push("runtime"); return COMPILE_CONTEXT; },
          adapter: {
            executeFirstGroup: async () => { calls.push("adapter"); return { status: "committed" as const }; },
            executeGroup: async () => { calls.push("adapter-extra"); return { status: "committed" as const }; },
            reconcileFailedGroup: async () => [],
          },
          route,
          trustedTurnOperationId: ACTIVE.turnId,
          initialResponse: { content: "", toolCalls: [] },
          rolloutEnabled,
        });
        expect(result.outcome).toBe("rejected");
        expect(calls).toEqual([]);
      }
    }
  });

  it("bounds TTL and count without evicting active state, then clears terminal, session, document, and stop state", () => {
    let now = 1_000;
    const store = new ServerOwnedTurnBindingStore({ now: () => now, ttlMs: 100, maxEntries: 1 });
    store.activateTurn(ACTIVE);
    expect(store.bindServerTicket(ticket())).toMatchObject({ ok: true });

    now = 1_100;
    expect(store.bindServerTicket(ticket())).toMatchObject({ ok: true, reused: true });

    const capacityStore = new ServerOwnedTurnBindingStore({ now: () => 1_000, ttlMs: 10_000, maxEntries: 1 });
    capacityStore.activateTurn(ACTIVE);
    expect(capacityStore.bindServerTicket(ticket())).toMatchObject({ ok: true });
    const otherPane: ServerOwnedPaneTurn = { ...ACTIVE, paneId: "66666666-6666-4666-8666-666666666666", turnId: "77777777-7777-4777-8777-777777777777" };
    capacityStore.activateTurn(otherPane);
    expect(capacityStore.bindServerTicket({ ...otherPane, serverOperationId: "88888888-8888-4888-8888-888888888888" })).toMatchObject({ ok: false });

    const cleanupStore = new ServerOwnedTurnBindingStore({ now: () => 1_000, ttlMs: 10_000, maxEntries: 4 });
    cleanupStore.activateTurn(ACTIVE);
    expect(cleanupStore.bindServerTicket(ticket())).toMatchObject({ ok: true });
    cleanupStore.finishTerminal(ticket());
    expect(cleanupStore.bindServerTicket(ticket())).toMatchObject({ ok: false });

    cleanupStore.clearTurn(ACTIVE);
    cleanupStore.activateTurn(ACTIVE);
    expect(cleanupStore.bindServerTicket(ticket())).toMatchObject({ ok: false });
    cleanupStore.invalidateSession(ACTIVE.sessionId);
    expect(cleanupStore.bindServerTicket(ticket())).toMatchObject({ ok: false });

    cleanupStore.activateTurn(ACTIVE);
    expect(cleanupStore.bindServerTicket(ticket())).toMatchObject({ ok: false });
    cleanupStore.invalidateDocument(ACTIVE.documentId);
    expect(cleanupStore.bindServerTicket(ticket())).toMatchObject({ ok: false });

    cleanupStore.activateTurn(ACTIVE);
    expect(cleanupStore.bindServerTicket(ticket())).toMatchObject({ ok: false });
    cleanupStore.stop();
    expect(cleanupStore.bindServerTicket(ticket())).toMatchObject({ ok: false });
  });

  it("blocks a stopped deferred preload from activating or reaching any plan/server seam while the next turn proceeds", async () => {
    let resolveDocumentA!: (documentId: string) => void;
    const documentA = new Promise<string>((resolve) => { resolveDocumentA = resolve; });
    const documentPreload = vi.fn()
      .mockReturnValueOnce(documentA)
      .mockResolvedValueOnce(ACTIVE.documentId);
    const bindings = new ServerOwnedTurnBindingStore({ now: () => 1_000, ttlMs: 10_000, maxEntries: 4 });
    const activateTurn = vi.spyOn(bindings, "activateTurn");
    const bindServerTicket = vi.spyOn(bindings, "bindServerTicket");
    const downstreamA = createDownstreamSpies();
    const downstreamB = createDownstreamSpies();
    const dependencies: TaskPaneTurnCoordinatorDependencies = {
      getOrCreateDocumentId: documentPreload,
      getBridgeIdentity: () => ({ sessionId: ACTIVE.sessionId, documentId: ACTIVE.documentId }),
      reconnectForDocumentChange: vi.fn(),
      bindings,
      paneId: ACTIVE.paneId,
    };
    const coordinator = new TaskPaneTurnCoordinator(dependencies);
    const controllerA = new AbortController();
    const turnA = coordinator.beginTurn("aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa", controllerA);
    const continuationA = vi.fn(async () => {
      downstreamA.run();
    });
    const continuationAPromise = coordinator.continueTurn(turnA, continuationA);

    coordinator.stopCurrentTurn();

    const controllerB = new AbortController();
    const turnB = coordinator.beginTurn("bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb", controllerB);
    const continuationB = vi.fn(async () => {
      downstreamB.run();
    });
    const continuationBPromise = coordinator.continueTurn(turnB, continuationB);
    resolveDocumentA("aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa");

    const [staleA, activeB] = await Promise.all([continuationAPromise, continuationBPromise]);
    expect(staleA.status).toBe("stale");
    expect(activeB).toMatchObject({ status: "active", documentId: ACTIVE.documentId });
    expect(continuationA).toHaveBeenCalledTimes(0);
    Object.values(downstreamA.spies).forEach((spy) => expect(spy).toHaveBeenCalledTimes(0));
    expect(continuationB).toHaveBeenCalledTimes(1);
    Object.values(downstreamB.spies).forEach((spy) => expect(spy).toHaveBeenCalledTimes(1));
    expect(activateTurn).toHaveBeenCalledTimes(1);

    const authorizedB = await coordinator.authorizeAndRun("cccccccc-cccc-4ccc-8ccc-cccccccccccc", async () => undefined);
    expect(authorizedB).toMatchObject({
      ok: true,
      identity: {
        sessionId: ACTIVE.sessionId,
        documentId: ACTIVE.documentId,
        paneId: ACTIVE.paneId,
        turnId: turnB.turnId,
        serverOperationId: "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
      },
    });
    expect(bindServerTicket).toHaveBeenCalledTimes(1);
    expect(coordinator.isCurrent(turnB.generation, turnB.turnId, controllerB)).toBe(true);
  });

  it("fails closed on an authenticated A/current B mismatch, then authorizes only a matching B turn", async () => {
    let bridgeIdentity = { sessionId: ACTIVE.sessionId, documentId: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa" };
    const reconnect = vi.fn(async () => ({
      sessionId: ACTIVE.sessionId,
      documentId: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
      protocolVersion: "1.0",
      expiresInSeconds: 60,
    }));
    const bindings = new ServerOwnedTurnBindingStore({ now: () => 1_000, ttlMs: 10_000, maxEntries: 4 });
    const activateTurn = vi.spyOn(bindings, "activateTurn");
    const bindServerTicket = vi.spyOn(bindings, "bindServerTicket");
    const downstream = createDownstreamSpies();
    const coordinator = new TaskPaneTurnCoordinator({
      getOrCreateDocumentId: vi.fn(async () => ACTIVE.documentId),
      getBridgeIdentity: () => bridgeIdentity,
      reconnectForDocumentChange: reconnect,
      bindings,
      paneId: ACTIVE.paneId,
    });
    const controllerA = new AbortController();
    const turnA = coordinator.beginTurn("dddddddd-dddd-4ddd-8ddd-dddddddddddd", controllerA);
    const mismatch = await coordinator.prepareTurn(turnA);
    expect(mismatch.status).toBe("rejected");
    expect(reconnect).toHaveBeenCalledTimes(1);
    expect(reconnect).toHaveBeenCalledWith(ACTIVE.documentId);
    expect(activateTurn).not.toHaveBeenCalled();

    const staleTicket = await coordinator.authorizeAndRun("eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee", async (authorized) => {
      downstream.run();
      return authorized.identity;
    });
    expect(staleTicket).toMatchObject({ ok: false });
    expect(bindServerTicket).not.toHaveBeenCalled();
    Object.values(downstream.spies).forEach((spy) => expect(spy).not.toHaveBeenCalled());

    bridgeIdentity = { sessionId: ACTIVE.sessionId, documentId: ACTIVE.documentId };
    const controllerB = new AbortController();
    const turnB = coordinator.beginTurn("ffffffff-ffff-4fff-8fff-ffffffffffff", controllerB);
    const matching = await coordinator.prepareTurn(turnB);
    expect(matching).toMatchObject({ status: "active", documentId: ACTIVE.documentId });
    const matchingDownstream = vi.fn(async (authorized: AuthorizedServerOperation) => {
      downstream.run();
      bindings.finishTerminal(authorized.identity);
      return authorized.identity;
    });
    const matchingTicket = await coordinator.authorizeAndRun("99999999-9999-4999-8999-999999999999", matchingDownstream);
    expect(matchingTicket).toMatchObject({
      ok: true,
      result: {
        sessionId: ACTIVE.sessionId,
        documentId: ACTIVE.documentId,
        paneId: ACTIVE.paneId,
        turnId: turnB.turnId,
        serverOperationId: "99999999-9999-4999-8999-999999999999",
      },
    });
    expect(matchingDownstream).toHaveBeenCalledTimes(1);
    expect(matchingDownstream).toHaveBeenCalledWith(expect.objectContaining({
      identity: expect.objectContaining({ turnId: turnB.turnId }),
    }));
    expect(bindServerTicket).toHaveBeenCalledTimes(1);
    Object.values(downstream.spies).forEach((spy) => expect(spy).toHaveBeenCalledTimes(1));
    expect(activateTurn).toHaveBeenCalledTimes(1);
  });

  it("sanitizes hostile document-preload errors at the trust seam", async () => {
    const hostileError = [
      'workbook-value="Payroll total $999,999.99"',
      'sheet="Hidden Payroll"',
      'address="XFD1048576"',
      'credential="Bearer sk-live-hostile"',
      'sentinel="D1_PRELOAD_HOSTILE_7F4C"',
    ].join(" ");
    const getOrCreateDocumentId = vi.fn(async () => {
      throw new Error(hostileError);
    });
    const reconnectForDocumentChange = vi.fn(async () => ({
      sessionId: ACTIVE.sessionId,
      documentId: ACTIVE.documentId,
    }));
    const coordinator = new TaskPaneTurnCoordinator({
      getOrCreateDocumentId,
      getBridgeIdentity: vi.fn(() => ({ sessionId: null, documentId: null })),
      reconnectForDocumentChange,
      bindings: new ServerOwnedTurnBindingStore({ now: () => 1_000, ttlMs: 10_000, maxEntries: 4 }),
      paneId: ACTIVE.paneId,
    });
    const handle = coordinator.beginTurn("aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa", new AbortController());

    const preparation = await coordinator.prepareTurn(handle);

    expect(preparation.status).toBe("rejected");
    if (preparation.status !== "rejected") throw new Error("expected rejected preparation");
    expect(preparation.reason).toBe("document-identity-unavailable");
    expect(preparation.reason).not.toContain(hostileError);
    for (const injectedValue of ["Payroll total $999,999.99", "Hidden Payroll", "XFD1048576", "Bearer sk-live-hostile", "D1_PRELOAD_HOSTILE_7F4C"]) {
      expect(preparation.reason).not.toContain(injectedValue);
    }
    expect(reconnectForDocumentChange).not.toHaveBeenCalled();
  });

  it("sanitizes hostile reconnect errors and prevents activation, binding, and downstream work", async () => {
    const hostileError = [
      'workbook-value="Current B confidential forecast"',
      'sheet="Authenticated A"',
      'address="B7:C9"',
      'credential="api_key=sk-reconnect-hostile"',
      'sentinel="D1_RECONNECT_HOSTILE_91BD"',
    ].join(" ");
    const bridgeIdentity = {
      sessionId: ACTIVE.sessionId,
      documentId: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
    };
    const reconnectForDocumentChange = vi.fn(async () => {
      throw new Error(hostileError);
    });
    const bindings = new ServerOwnedTurnBindingStore({ now: () => 1_000, ttlMs: 10_000, maxEntries: 4 });
    const activateTurn = vi.spyOn(bindings, "activateTurn");
    const bindServerTicket = vi.spyOn(bindings, "bindServerTicket");
    const downstream = createDownstreamSpies();
    const coordinator = new TaskPaneTurnCoordinator({
      getOrCreateDocumentId: vi.fn(async () => ACTIVE.documentId),
      getBridgeIdentity: () => bridgeIdentity,
      reconnectForDocumentChange,
      bindings,
      paneId: ACTIVE.paneId,
    });
    const handle = coordinator.beginTurn("bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb", new AbortController());

    const preparation = await coordinator.prepareTurn(handle);

    expect(preparation.status).toBe("rejected");
    if (preparation.status !== "rejected") throw new Error("expected rejected preparation");
    expect(preparation.reason).toBe("bridge-reconnect-failed");
    expect(preparation.reason).not.toContain(hostileError);
    for (const injectedValue of ["Current B confidential forecast", "Authenticated A", "B7:C9", "api_key=sk-reconnect-hostile", "D1_RECONNECT_HOSTILE_91BD"]) {
      expect(preparation.reason).not.toContain(injectedValue);
    }
    expect(reconnectForDocumentChange).toHaveBeenCalledTimes(1);
    expect(reconnectForDocumentChange).toHaveBeenCalledWith(ACTIVE.documentId);
    expect(activateTurn).not.toHaveBeenCalled();

    const authorization = await coordinator.authorizeAndRun("cccccccc-cccc-4ccc-8ccc-cccccccccccc", async () => {
      downstream.run();
      return "must-not-run";
    });
    expect(authorization).toMatchObject({ ok: false });
    expect(bindServerTicket).not.toHaveBeenCalled();
    Object.values(downstream.spies).forEach((spy) => expect(spy).not.toHaveBeenCalled());
  });

  it("fences an authorized A route after bridge identity replacement before runtime/preparation/Office mutation, while B remains attributable", async () => {
    let bridgeIdentity = { sessionId: ACTIVE.sessionId, documentId: ACTIVE.documentId };
    let resolveRuntimeA!: (context: VerifiedWorkbookCompileContext) => void;
    let runtimeAStarted = false;
    const runtimeA = new Promise<VerifiedWorkbookCompileContext>((resolve) => { resolveRuntimeA = resolve; });
    const bindings = new ServerOwnedTurnBindingStore({ now: () => 1_000, ttlMs: 10_000, maxEntries: 8 });
    const coordinator = new TaskPaneTurnCoordinator({
      getOrCreateDocumentId: vi.fn(async () => ACTIVE.documentId),
      getBridgeIdentity: () => bridgeIdentity,
      reconnectForDocumentChange: vi.fn(),
      bindings,
      paneId: ACTIVE.paneId,
    });
    const controllerA = new AbortController();
    const mutableRequestA: { classification: TrustedWorkbookRequest["classification"]; selectionScopedRequest: boolean } = {
      classification: "explicit_mutation",
      selectionScopedRequest: false,
    };
    const originalRequestA = Object.freeze({ ...mutableRequestA });
    const turnA = coordinator.beginTurn("aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa", controllerA, {
      trustedRequest: mutableRequestA,
      turnStart: 1_234,
      cancellationSignal: controllerA.signal,
    });
    mutableRequestA.classification = "read_only";
    mutableRequestA.selectionScopedRequest = true;
    await expect(coordinator.prepareTurn(turnA)).resolves.toMatchObject({ status: "active" });
    const plan = {
      version: 1,
      outcome: "execute",
      summary: "authorized A write",
      operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["A"]] }],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "B1", expected: null }],
    } as const;
    const planResponse = {
      content: "",
      toolCalls: [{
        id: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
        type: "function" as const,
        function: { name: "excel_execute_plan", arguments: JSON.stringify(plan) },
      }],
    };
    const executePolicy = {
      ...POLICY,
      operationSummaries: [{ operationIndex: 0, effect: "write" as const, target: { operationIndex: 0, sheetNames: ["Sheet1"], selectionRelationship: "inside" as const } }],
      resourceUsage: { ...POLICY.resourceUsage, operationCount: 1, cellsWritten: 1 },
    };
    let adapterMutationA = 0;
    const operationA = coordinator.authorizeAndRun("bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb", async (authorized) => {
      expect(authorized.snapshot).toMatchObject({
        generation: turnA.generation,
        turnId: turnA.turnId,
        sessionId: ACTIVE.sessionId,
        documentId: ACTIVE.documentId,
        paneId: ACTIVE.paneId,
        serverOperationId: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
        turnStart: 1_234,
        trustedRequest: originalRequestA,
      });
      expect(authorized.snapshot.trustedRequest).not.toBe(mutableRequestA);
      expect(authorized.snapshot.trustedRequest).toEqual(originalRequestA);
      expect(authorized.snapshot.trustedRequest).not.toEqual(mutableRequestA);
      const result = await createWorkbookPlanTurn({
        identity: {
          sessionId: authorized.identity.sessionId,
          documentId: authorized.identity.documentId,
          turnId: authorized.identity.turnId,
        },
        trustedOperationBoundary: createTrustedOperationBoundary({
          sessionId: authorized.identity.sessionId,
          documentId: authorized.identity.documentId,
          turnId: authorized.identity.turnId,
        }),
        policyInput: executePolicy,
        store: new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
        initialResponse: planResponse,
        model: async () => { throw new Error("A must not request model continuation"); },
        inspect: async () => { throw new Error("A must not inspect"); },
        getCompileContextForPlan: async () => {
          runtimeAStarted = true;
          return runtimeA;
        },
        resourcePreflightAfterRuntime: async () => ({ ok: true as const }),
        adapter: {
          executeFirstGroup: async (gate: Parameters<WorkbookPlanExecutionAdapter["executeFirstGroup"]>[0]) => {
            adapterMutationA += 1;
            return gate.executeAfterFresh("fresh", async () => ({ status: "committed" as const }));
          },
          executeGroup: async () => ({ status: "committed" as const }),
          reconcileFailedGroup: async (group: Parameters<WorkbookPlanExecutionAdapter["reconcileFailedGroup"]>[0]) => group.operations.map(() => "unknown_within_batch" as const),
          verifyPlan: async () => true,
        },
        signal: controllerA.signal,
        route: "server-owned",
        trustedTurnOperationId: authorized.identity.serverOperationId,
        rolloutEnabled: true,
        assertLive: authorized.assertLive,
      });
      return result;
    });

    await vi.waitFor(() => expect(runtimeAStarted).toBe(true));
    bridgeIdentity = {
      sessionId: "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
      documentId: "dddddddd-dddd-4ddd-8ddd-dddddddddddd",
    };
    const coordinatorB = new TaskPaneTurnCoordinator({
      getOrCreateDocumentId: vi.fn(async () => bridgeIdentity.documentId),
      getBridgeIdentity: () => bridgeIdentity,
      reconnectForDocumentChange: vi.fn(),
      bindings,
      paneId: ACTIVE.paneId,
    });
    const controllerB = new AbortController();
    const turnB = coordinatorB.beginTurn("eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee", controllerB);
    await expect(coordinatorB.prepareTurn(turnB)).resolves.toMatchObject({ status: "active" });
    const operationB = await coordinatorB.authorizeAndRun("ffffffff-ffff-4fff-8fff-ffffffffffff", async (authorized) => authorized.identity);
    expect(operationB).toMatchObject({ ok: true, result: { sessionId: bridgeIdentity.sessionId, documentId: bridgeIdentity.documentId, turnId: turnB.turnId } });

    resolveRuntimeA(COMPILE_CONTEXT);
    const staleA = await operationA;
    expect(staleA.ok).toBe(false);
    expect(adapterMutationA).toBe(0);
  });
});
