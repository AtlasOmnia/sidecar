import { describe, expect, it, vi } from "vitest";
import { createNativeWorkbookPlanAdapter } from "../workbookPlanOfficeAdapter";
import {
  BrowserIdempotencyStore,
  MemoryIdempotencyStorage,
  createTrustedOperationBoundary,
  executeAcceptedWorkbookPlan,
  prepareWorkbookPlanExecution,
  type WorkbookPlanExecutionAdapter,
} from "../workbookPlanExecutor";
import {
  createWorkbookPlanTurn,
  type WorkbookPlanModelResponse,
  type WorkbookPlanTurnDependencies,
} from "../workbookPlanOrchestrator";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import type { WorkbookPlanPolicyInput } from "../workbookPlanPolicy";
import type { WorkbookPlan } from "../generated/workbookPlanSchema";

const UUIDS = [
  "11111111-1111-4111-8111-111111111111",
  "22222222-2222-4222-8222-222222222222",
  "33333333-3333-4333-8333-333333333333",
  "44444444-4444-4444-8444-444444444444",
  "55555555-5555-4555-8555-555555555555",
  "66666666-6666-4666-8666-666666666666",
  "77777777-7777-4777-8777-777777777777",
  "88888888-8888-4888-8888-888888888888",
  "99999999-9999-4999-8999-999999999999",
];

const primitiveNames = ["write_range", "format_range"];
const compileContext: VerifiedWorkbookCompileContext = {
  sheetNames: new Set(["Sheet1"]),
  tablesByName: new Map(),
  supportedPrimitiveNames: new Set(primitiveNames),
  supportedCapabilities: new Set(["excel_workbook_access", "workbook_plan_v1_operation_audit"]),
  chartSourceResolutionsBySheet: new Map(),
};

function boundary(turnIndex: number) {
  return createTrustedOperationBoundary({
    sessionId: UUIDS[0],
    documentId: UUIDS[1],
    turnId: UUIDS[turnIndex],
  });
}

function operation(kind: "write_values" | "format_range", address: string): Record<string, unknown> {
  return kind === "write_values"
    ? { kind, sheet: "Sheet1", address, values: [[address]] }
    : { kind, sheet: "Sheet1", address, format: { numberFormat: "0.00" } };
}

function planFor(kinds: readonly ("write_values" | "format_range")[]): Extract<WorkbookPlan, { readonly outcome: "execute" }> {
  return {
    version: 1,
    outcome: "execute",
    summary: "bounded cancellation test",
    operations: kinds.map((kind, index) => operation(kind, `${String.fromCharCode(65 + index)}1`)) as never,
    verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "Z1", expected: null }],
  };
}

function policyFor(effects: readonly ("write" | "format")[]): WorkbookPlanPolicyInput {
  return {
    requestClassification: "explicit_mutation",
    activeSheetName: "Sheet1",
    selectionScopedRequest: false,
    operationSummaries: effects.map((effect, operationIndex) => ({
      operationIndex,
      effect,
      target: { operationIndex, sheetNames: ["Sheet1"], selectionRelationship: "inside" },
    })),
    capabilitySummaries: ["excel_workbook_access", "workbook_plan_v1_operation_audit"],
    resourceUsage: {
      operationCount: effects.length,
      cellsRead: 0,
      cellsWritten: effects.length,
      formulaUtf8Bytes: 0,
      csvPayloadUtf8Bytes: 0,
      commentUtf8Bytes: 0,
      sheetsCreated: 0,
      chartsCreated: 0,
      tablesCreated: 0,
      verificationReads: 1,
    },
    resourceLimits: {
      operationCount: 10,
      cellsRead: 10,
      cellsWritten: 10,
      formulaUtf8Bytes: 1000,
      csvPayloadUtf8Bytes: 1000,
      commentUtf8Bytes: 1000,
      sheetsCreated: 1,
      chartsCreated: 1,
      tablesCreated: 1,
      verificationReads: 10,
    },
  };
}

function uuidSource(): () => string {
  let index = 3;
  return () => UUIDS[index++] ?? UUIDS[8];
}

async function accepted(
  plan: WorkbookPlan,
  effects: readonly ("write" | "format")[],
  turnIndex: number,
  extra: Partial<Parameters<typeof prepareWorkbookPlanExecution>[0]> = {},
) {
  const result = await prepareWorkbookPlanExecution({
    plan,
    compileContext,
    policyInput: policyFor(effects),
    trustedOperationBoundary: extra.trustedOperationBoundary ?? boundary(turnIndex),
    uuidSource: uuidSource(),
    ...extra,
  });
  expect(result.accepted).toBe(true);
  if (!result.accepted) throw new Error("expected accepted plan");
  return result.transaction;
}

function modelCall(plan: WorkbookPlan): WorkbookPlanModelResponse {
  return {
    content: "",
    toolCalls: [{
      id: "call-1",
      type: "function",
      function: { name: "excel_execute_plan", arguments: JSON.stringify(plan) },
    }],
  };
}

function turnDependencies(
  plan: Extract<WorkbookPlan, { readonly outcome: "execute" }>,
  signal: AbortSignal,
  adapter: WorkbookPlanExecutionAdapter,
  turnIndex: number,
  overrides: Partial<WorkbookPlanTurnDependencies> = {},
): WorkbookPlanTurnDependencies {
  return {
    identity: { sessionId: UUIDS[0], documentId: UUIDS[1], turnId: UUIDS[turnIndex] },
    trustedOperationBoundary: boundary(turnIndex),
    compileContext,
    policyInput: policyFor(plan.operations.map((item) => item.kind === "format_range" ? "format" : "write")),
    store: new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
    model: async () => modelCall(plan),
    inspect: async () => "{}",
    resourcePreflightAfterRuntime: async () => ({ ok: true as const }),
    adapter,
    signal,
    rolloutEnabled: true,
    initialResponse: modelCall(plan),
    trustedTurnOperationId: UUIDS[turnIndex + 1],
    route: "direct",
    turnStart: Date.now(),
    ...overrides,
  };
}

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

function adapterWithFirst(
  first: WorkbookPlanExecutionAdapter["executeFirstGroup"],
  later: WorkbookPlanExecutionAdapter["executeGroup"] = async () => ({ status: "committed" }),
  reconcile: WorkbookPlanExecutionAdapter["reconcileFailedGroup"] = async (group) => group.operations.map(() => "unknown_within_batch" as const),
): WorkbookPlanExecutionAdapter {
  return { executeFirstGroup: first, executeGroup: later, reconcileFailedGroup: reconcile };
}

describe("Task 11 B2a2 in-flight cancellation production seams", () => {
  it("aborts before the first mutation handoff with zero adapter mutation", async () => {
    const controller = new AbortController();
    controller.abort();
    const trustedBoundary = boundary(2);
    const transaction = await accepted(planFor(["write_values", "format_range"]), ["write", "format"], 2, {
      trustedOperationBoundary: trustedBoundary,
    });
    let firstCalls = 0;
    let laterCalls = 0;

    const report = await executeAcceptedWorkbookPlan(
      transaction,
      trustedBoundary,
      adapterWithFirst(
        async () => { firstCalls += 1; return { status: "committed" }; },
        async () => { laterCalls += 1; return { status: "committed" }; },
      ),
      new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      { signal: controller.signal },
    );

    expect(report.outcome).toBe("cancelled");
    expect(firstCalls).toBe(0);
    expect(laterCalls).toBe(0);
    expect(report.operations.every((item) => item.state === "not_committed")).toBe(true);
  });

  it("aborts during the first whole-plan precondition read without queueing mutation", async () => {
    const controller = new AbortController();
    const precondition = deferred<"fresh" | "stale">();
    let mutationCalls = 0;
    const trustedBoundary = boundary(3);
    const transaction = await accepted(planFor(["write_values"]), ["write"], 3, {
      trustedOperationBoundary: trustedBoundary,
      preconditionCheck: () => precondition.promise,
    });

    const execution = executeAcceptedWorkbookPlan(
      transaction,
      trustedBoundary,
      adapterWithFirst(async (gate) => gate.executeAfterFresh("fresh", async () => {
        mutationCalls += 1;
        return { status: "committed" };
      })),
      new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      { signal: controller.signal },
    );
    await vi.waitFor(() => expect(mutationCalls).toBe(0));
    controller.abort();
    precondition.resolve("fresh");
    const report = await execution;

    expect(report.outcome).toBe("cancelled");
    expect(mutationCalls).toBe(0);
    expect(report.operations[0].state).toBe("not_committed");
  });

  it("keeps the orchestrator promise pending through a committed first Office batch and skips later work/verification", async () => {
    const controller = new AbortController();
    const batch = deferred<{ readonly status: "committed" }>();
    let handedOff = false;
    let laterCalls = 0;
    let verificationCalls = 0;
    const plan = planFor(["write_values", "format_range"]);
    const adapter = adapterWithFirst(
      async (gate) => gate.executeAfterFresh("fresh", async () => {
        handedOff = true;
        return batch.promise;
      }),
      async () => { laterCalls += 1; return { status: "committed" }; },
    ) as WorkbookPlanExecutionAdapter & { verifyPlan: () => Promise<boolean> };
    adapter.verifyPlan = async () => { verificationCalls += 1; return true; };

    const turn = createWorkbookPlanTurn(turnDependencies(plan, controller.signal, adapter, 4));
    await vi.waitFor(() => expect(handedOff).toBe(true));
    controller.abort();
    let settled = false;
    void turn.then(() => { settled = true; });
    await Promise.resolve();
    expect(settled).toBe(false);

    batch.resolve({ status: "committed" });
    const result = await turn;

    expect(result.outcome).toBe("partial");
    expect(result.execution?.outcome).toBe("partial");
    expect(result.execution?.operations[0].state).toBe("committed");
    expect(result.execution?.operations[1].state).toBe("not_committed");
    expect(laterCalls).toBe(0);
    expect(verificationCalls).toBe(0);
  });

  it("waits for a failed first batch, preserves unknown state, and never claims rollback", async () => {
    const controller = new AbortController();
    const batch = deferred<never>();
    let handedOff = false;
    let rollbackCalls = 0;
    let reconcileCalls = 0;
    const trustedBoundary = boundary(5);
    const transaction = await accepted(planFor(["write_values", "format_range"]), ["write", "format"], 5, {
      trustedOperationBoundary: trustedBoundary,
      inverse: {
        completeCapturedPreState: true,
        independentlyVerifiedSafe: true,
        rollback: async () => { rollbackCalls += 1; return true; },
      },
    });

    const execution = executeAcceptedWorkbookPlan(
      transaction,
      trustedBoundary,
      adapterWithFirst(
        async (gate) => gate.executeAfterFresh("fresh", async () => {
          handedOff = true;
          return batch.promise;
        }),
        async () => { throw new Error("later group must not start"); },
        async (group) => {
          reconcileCalls += 1;
          return group.operations.map(() => "unknown_within_batch" as const);
        },
      ),
      new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      { signal: controller.signal },
    );
    await vi.waitFor(() => expect(handedOff).toBe(true));
    controller.abort();
    let settled = false;
    void execution.then(() => { settled = true; });
    await Promise.resolve();
    expect(settled).toBe(false);

    batch.reject(new Error("Office sync failed"));
    const report = await execution;

    expect(report.outcome).toBe("partial");
    expect(report.operations[0].state).toBe("unknown_within_batch");
    expect(report.operations[0].status).toBe("unknown_within_batch");
    expect(report.operations[1].state).toBe("not_committed");
    expect(report.rollback).toBe("not_attempted");
    expect(rollbackCalls).toBe(0);
    expect(reconcileCalls).toBe(1);
  });

  it("waits for a later-group handoff, preserves prior committed groups, and does not start the next group", async () => {
    const controller = new AbortController();
    const batch = deferred<never>();
    let laterHandedOff = false;
    let laterCalls = 0;
    let finalGroupCalls = 0;
    const trustedBoundary = boundary(6);
    const transaction = await accepted(planFor(["write_values", "format_range", "write_values"]), ["write", "format", "write"], 6, {
      trustedOperationBoundary: trustedBoundary,
    });

    const execution = executeAcceptedWorkbookPlan(
      transaction,
      trustedBoundary,
      adapterWithFirst(
        async (gate) => gate.executeAfterFresh("fresh", async () => ({ status: "committed" })),
        async (_group, groupIndex) => {
          if (groupIndex === 2) { finalGroupCalls += 1; return { status: "committed" }; }
          laterCalls += 1;
          laterHandedOff = true;
          return batch.promise;
        },
        async (group) => group.operations.map(() => "unknown_within_batch" as const),
      ),
      new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      { signal: controller.signal },
    );
    await vi.waitFor(() => expect(laterHandedOff).toBe(true));
    controller.abort();
    let settled = false;
    void execution.then(() => { settled = true; });
    await Promise.resolve();
    expect(settled).toBe(false);

    batch.reject(new Error("Office sync failed"));
    const report = await execution;

    expect(report.outcome).toBe("partial");
    expect(report.operations.map((item) => item.state)).toEqual([
      "committed", "unknown_within_batch", "not_committed",
    ]);
    expect(report.operations[1].status).toBe("unknown_within_batch");
    expect(laterCalls).toBe(1);
    expect(finalGroupCalls).toBe(0);
  });
});

describe("Task 11 B2a2 native Office adapter handoff boundaries", () => {
  it("awaits an already-started context.sync after abort and reports its real commit", async () => {
    const controller = new AbortController();
    const mutationSync = deferred<void>();
    let syncCalls = 0;
    let valuesLoaded = false;
    const range: Record<string, unknown> = {
      format: {},
      load: () => { valuesLoaded = false; },
    };
    Object.defineProperty(range, "values", {
      configurable: true,
      get: () => {
        if (!valuesLoaded) throw new Error("read before load+sync: values");
        return [["before"]];
      },
      set: () => { /* mutation queue is intentionally observable only through sync */ },
    });
    const context = {
      workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
      sync: async () => {
        syncCalls += 1;
        if (syncCalls === 1) {
          valuesLoaded = true;
          return;
        }
        return mutationSync.promise;
      },
    };
    const previousExcel = (globalThis as Record<string, unknown>).Excel;
    (globalThis as Record<string, unknown>).Excel = {
      run: async (callback: (value: typeof context) => Promise<unknown>) => callback(context),
    };
    try {
      const adapter = createNativeWorkbookPlanAdapter();
      const execution = adapter.executeGroup({
        groupIndex: 0,
        effect: "write",
        operations: [{
          operationId: UUIDS[3],
          operationIndex: 0,
          kind: "write_values",
          primitiveName: "write_range",
          args: { sheet: "Sheet1", address: "A1", values: [["x"]] },
        }],
      }, 0, controller.signal);
      // R2 adds an exact pre-state read sync before mutation. Abort only after
      // the second sync has handed the queued mutation to Office.
      await vi.waitFor(() => expect(syncCalls).toBe(2));
      controller.abort();
      let settled = false;
      void execution.then(() => { settled = true; });
      await Promise.resolve();
      expect(settled).toBe(false);
      mutationSync.resolve();
      await expect(execution).resolves.toEqual({ status: "committed" });
    } finally {
      if (previousExcel === undefined) delete (globalThis as Record<string, unknown>).Excel;
      else (globalThis as Record<string, unknown>).Excel = previousExcel;
    }
  });

  it("checks cancellation after the whole-plan precondition read before queueing mutation", async () => {
    const controller = new AbortController();
    const firstSync = deferred<void>();
    let syncCalls = 0;
    let mutationQueued = false;
    const preconditionRange: Record<string, unknown> = {
      rowCount: 1,
      columnCount: 1,
      address: "A1",
      load: () => undefined,
    };
    Object.defineProperty(preconditionRange, "values", {
      configurable: true,
      get: () => [[null]],
      set: () => { mutationQueued = true; },
    });
    const sheet = { getRange: () => preconditionRange };
    const context = {
      workbook: {
        worksheets: {
          getItem: () => sheet,
          getItemOrNullObject: () => ({ isNullObject: false, load: () => undefined }),
        },
        tables: { getItemOrNullObject: () => ({ isNullObject: false, load: () => undefined }) },
      },
      sync: () => {
        syncCalls += 1;
        return syncCalls === 1 ? firstSync.promise : Promise.resolve();
      },
    };
    const previousExcel = (globalThis as Record<string, unknown>).Excel;
    (globalThis as Record<string, unknown>).Excel = {
      run: async (callback: (value: typeof context) => Promise<unknown>) => callback(context),
    };
    try {
      const adapter = createNativeWorkbookPlanAdapter();
      const execution = adapter.executeFirstGroup({
        preconditionManifest: {
          version: 1,
          sheets: [],
          tables: [],
          ranges: [{ sheet: "Sheet1", address: "A1", rowCount: 1, columnCount: 1 }],
          chartSources: [],
          charts: [],
          comments: [],
        },
        executeAfterFresh: async (_observation, mutation) => mutation({
          groupIndex: 0,
          effect: "write",
          operations: [{
            operationId: UUIDS[4],
            operationIndex: 0,
            kind: "write_values",
            primitiveName: "write_range",
            args: { sheet: "Sheet1", address: "A1", values: [["x"]] },
          }],
        }),
      }, controller.signal);
      await vi.waitFor(() => expect(syncCalls).toBe(1));
      controller.abort();
      firstSync.resolve();
      await expect(execution).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
      expect(mutationQueued).toBe(false);
      expect(syncCalls).toBe(1);
    } finally {
      if (previousExcel === undefined) delete (globalThis as Record<string, unknown>).Excel;
      else (globalThis as Record<string, unknown>).Excel = previousExcel;
    }
  });
});
