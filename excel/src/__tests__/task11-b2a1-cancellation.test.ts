import { describe, expect, it, vi } from "vitest";
import { createTurnCancellation } from "../agent";
import {
  createWorkbookPlanTurn,
  type WorkbookPlanModelResponse,
  type WorkbookPlanTurnDependencies,
  type WorkbookPlanVerificationAdapter,
} from "../workbookPlanOrchestrator";
import {
  BrowserIdempotencyStore,
  MemoryIdempotencyStorage,
  type WorkbookPlanExecutionAdapter,
  createTrustedOperationBoundary,
} from "../workbookPlanExecutor";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import type { WorkbookPlanPolicyInput } from "../workbookPlanPolicy";

const IDS = {
  sessionId: "11111111-1111-4111-8111-111111111111",
  documentId: "22222222-2222-4222-8222-222222222222",
  turnId: "33333333-3333-4333-8333-333333333333",
};
const boundary = createTrustedOperationBoundary(IDS);
const compileContext: VerifiedWorkbookCompileContext = {
  sheetNames: new Set(["Sheet1"]),
  tablesByName: new Map(),
  supportedPrimitiveNames: new Set(["write_range"]),
  supportedCapabilities: new Set(["excel_workbook_access", "workbook_plan_v1_operation_audit"]),
  chartSourceResolutionsBySheet: new Map(),
};
const policy: WorkbookPlanPolicyInput = {
  requestClassification: "explicit_mutation",
  activeSheetName: "Sheet1",
  selectionScopedRequest: false,
  operationSummaries: [{ operationIndex: 0, effect: "write", target: { operationIndex: 0, sheetNames: ["Sheet1"], selectionRelationship: "inside" } }],
  capabilitySummaries: ["excel_workbook_access", "workbook_plan_v1_operation_audit"],
  resourceUsage: { operationCount: 1, cellsRead: 0, cellsWritten: 1, formulaUtf8Bytes: 0, csvPayloadUtf8Bytes: 0, commentUtf8Bytes: 0, sheetsCreated: 0, chartsCreated: 0, tablesCreated: 0, verificationReads: 1 },
  resourceLimits: { operationCount: 10, cellsRead: 10, cellsWritten: 10, formulaUtf8Bytes: 1000, csvPayloadUtf8Bytes: 1000, commentUtf8Bytes: 1000, sheetsCreated: 1, chartsCreated: 1, tablesCreated: 1, verificationReads: 10 },
};
const executePlan = {
  version: 1,
  outcome: "execute",
  summary: "bounded write",
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }],
  verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "B1", expected: null }],
} as const;
const inspectPlan = {
  version: 1,
  outcome: "inspect",
  reason: "need context",
  reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1" }],
} as const;
const call = (payload: unknown): WorkbookPlanModelResponse => ({
  content: "",
  toolCalls: [{ id: "call-1", type: "function", function: { name: "excel_execute_plan", arguments: JSON.stringify(payload) } }],
});

function adapter(overrides: Partial<WorkbookPlanExecutionAdapter & WorkbookPlanVerificationAdapter> = {}): WorkbookPlanExecutionAdapter & WorkbookPlanVerificationAdapter {
  return {
    executeFirstGroup: async (gate) => gate.executeAfterFresh("fresh", async () => ({ status: "committed" })),
    executeGroup: async () => ({ status: "committed" }),
    reconcileFailedGroup: async (group) => group.operations.map(() => "unknown_within_batch" as const),
    ...overrides,
  };
}

function dependencies(overrides: Partial<WorkbookPlanTurnDependencies> = {}): WorkbookPlanTurnDependencies {
  return {
    identity: IDS,
    trustedOperationBoundary: boundary,
    compileContext,
    policyInput: policy,
    store: new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
    model: async () => call(executePlan),
    inspect: async () => "{}",
    resourcePreflightAfterRuntime: async () => ({ ok: true as const }),
    adapter: adapter(),
    rolloutEnabled: true,
    ...overrides,
  };
}

describe("Task 11 B2a1 pane-origin cancellation", () => {
  it("fails closed when earlier pane work has already consumed the absolute turn budget", async () => {
    let modelCalls = 0;
    let inspectCalls = 0;
    let runtimeCalls = 0;
    let mutationCalls = 0;
    let verificationCalls = 0;
    const result = await createWorkbookPlanTurn(dependencies({
      turnStart: 1_000,
      now: () => 91_001,
      model: async () => { modelCalls += 1; return call(executePlan); },
      inspect: async () => { inspectCalls += 1; return "{}"; },
      getCompileContextForPlan: async () => { runtimeCalls += 1; return compileContext; },
      adapter: adapter({
        executeFirstGroup: async () => { mutationCalls += 1; return { status: "committed" }; },
        verifyPlan: async () => { verificationCalls += 1; return true; },
      }),
    }));

    expect(result.outcome).toBe("cancelled");
    expect(modelCalls).toBe(0);
    expect(inspectCalls).toBe(0);
    expect(runtimeCalls).toBe(0);
    expect(mutationCalls).toBe(0);
    expect(verificationCalls).toBe(0);
  });

  it("threads the exact pane signal through inspection, runtime, preparation execution, and verification", async () => {
    const userController = new AbortController();
    const cancellation = createTurnCancellation(Date.now(), userController.signal);
    const seen: AbortSignal[] = [];
    const result = await createWorkbookPlanTurn(dependencies({
      signal: cancellation.signal,
      turnStart: Date.now(),
      inspect: async (_reads, signal) => { seen.push(signal!); return "{}"; },
      model: async () => call(executePlan),
      getCompileContextForPlan: async (_plan, signal) => { seen.push(signal!); return compileContext; },
      resourcePreflight: async (_plan, signal) => { seen.push(signal!); return { ok: true }; },
      adapter: adapter({
        executeFirstGroup: async (gate, signal) => { seen.push(signal!); return gate.executeAfterFresh("fresh", async () => ({ status: "committed" })); },
        verifyPlan: async (_rules, signal) => { seen.push(signal!); return true; },
      }),
    }));
    cancellation.cleanup();

    expect(result.outcome).toBe("committed");
    expect(result.verification).toBe("verified");
    expect(seen.length).toBeGreaterThanOrEqual(4);
    expect(seen.every((signal) => signal === cancellation.signal)).toBe(true);
  });

  it("actively aborts a blocked signal-aware inspection before the turn resolves", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);
    const cancellation = createTurnCancellation(0, new AbortController().signal);
    let inspectionSettled = false;
    let inspectionStarted = false;
    const turn = createWorkbookPlanTurn(dependencies({
      signal: cancellation.signal,
      turnStart: 0,
      now: () => Date.now(),
      model: async () => call(inspectPlan),
      inspect: async (_reads, signal) => {
        inspectionStarted = true;
        return new Promise<string>((resolve) => {
          signal!.addEventListener("abort", () => { inspectionSettled = true; resolve("{}"); }, { once: true });
        });
      },
    }));

    await vi.waitFor(() => {
      expect(inspectionStarted).toBe(true);
    }, { timeout: 1_000, interval: 1 });
    await vi.advanceTimersByTimeAsync(90_000);
    const result = await turn;
    cancellation.cleanup();
    vi.useRealTimers();

    expect(result.outcome).toBe("cancelled");
    expect(inspectionSettled).toBe(true);
  });

  it("uses pane-owned timing and signal for a server-owned ticket, not ticket payload timing", async () => {
    const userController = new AbortController();
    const cancellation = createTurnCancellation(Date.now(), userController.signal);
    const seen: AbortSignal[] = [];
    let modelCalls = 0;
    const result = await createWorkbookPlanTurn(dependencies({
      route: "server-owned",
      trustedTurnOperationId: "44444444-4444-4444-8444-444444444444",
      initialResponse: call(executePlan),
      signal: cancellation.signal,
      turnStart: Date.now(),
      model: async () => { modelCalls += 1; return call(executePlan); },
      getCompileContextForPlan: async (_plan, signal) => { seen.push(signal!); return compileContext; },
      resourcePreflight: async (_plan, signal) => { seen.push(signal!); return { ok: true }; },
      adapter: adapter({
        executeFirstGroup: async (gate, signal) => { seen.push(signal!); return gate.executeAfterFresh("fresh", async () => ({ status: "committed" })); },
        verifyPlan: async (_rules, signal) => { seen.push(signal!); return true; },
      }),
    }));
    cancellation.cleanup();

    expect(result.outcome).toBe("committed");
    expect(modelCalls).toBe(0);
    expect(seen.every((signal) => signal === cancellation.signal)).toBe(true);
  });
});
