import { describe, expect, it } from "vitest";
import {
  createWorkbookPlanTurn,
  WorkbookPlanTurnState,
  type WorkbookPlanModelResponse,
  type WorkbookPlanTurnDependencies,
} from "../workbookPlanOrchestrator";
import { createTrustedOperationBoundary, MemoryIdempotencyStorage, BrowserIdempotencyStore, type WorkbookPlanExecutionAdapter } from "../workbookPlanExecutor";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import type { WorkbookPlanPolicyInput } from "../workbookPlanPolicy";

const IDS = {
  sessionId: "11111111-1111-4111-8111-111111111111",
  documentId: "22222222-2222-4222-8222-222222222222",
  turnId: "33333333-3333-4333-8333-333333333333",
};
const boundary = createTrustedOperationBoundary(IDS);
const context: VerifiedWorkbookCompileContext = {
  sheetNames: new Set(["Sheet1"]),
  tablesByName: new Map(),
  supportedPrimitiveNames: new Set(["write_range", "run_calculation"]),
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
const executePlan = { version: 1, outcome: "execute", summary: "model claim must not be used", operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }], verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "B1", expected: null }] } as const;
const inspectPlan = { version: 1, outcome: "inspect", reason: "need context", reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1" }] } as const;
const askPlan = { version: 1, outcome: "ask", message: "Please clarify" } as const;
const call = (payload: unknown): WorkbookPlanModelResponse => ({ content: "", toolCalls: [{ id: "call-1", type: "function", function: { name: "excel_execute_plan", arguments: JSON.stringify(payload) } }] });

function adapter(readback: string): { adapter: WorkbookPlanExecutionAdapter; calls: string[]; reads: number } {
  const calls: string[] = [];
  let reads = 0;
  return {
    calls,
    get reads() { return reads; },
    adapter: {
      async executeFirstGroup(gate) { calls.push("execute"); return gate.executeAfterFresh("fresh", async () => ({ status: "committed" })); },
      async executeGroup() { calls.push("execute-extra"); return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch" as const); },
      readVerification() { reads += 1; return readback; },
    } as WorkbookPlanExecutionAdapter & { readVerification: () => string },
  };
}

function deps(responses: WorkbookPlanModelResponse[], overrides: Partial<WorkbookPlanTurnDependencies> = {}): WorkbookPlanTurnDependencies {
  let index = 0;
  const a = adapter(JSON.stringify({ status: "supported", snapshot: { kind: "sentinel", value: null } }));
  return {
    identity: IDS,
    trustedOperationBoundary: boundary,
    compileContext: context,
    policyInput: policy,
    store: new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
    model: async () => responses[index++] ?? responses[responses.length - 1]!,
    inspect: async () => { a.calls.push("inspect"); return JSON.stringify({ ok: true }); },
    resourcePreflightAfterRuntime: async () => ({ ok: true as const }),
    adapter: a.adapter,
    rolloutEnabled: true,
    ...overrides,
  };
}

describe("route-neutral workbook plan orchestrator", () => {
  it("performs at most one inspect cycle followed by one fully verified execution", async () => {
    const a = adapter(JSON.stringify({ status: "supported", snapshot: { kind: "sentinel", value: null } }));
    const result = await createWorkbookPlanTurn(deps([call(inspectPlan), call(executePlan)], { adapter: a.adapter, inspect: async () => { a.calls.push("inspect"); return "{}"; } }));
    expect(a.calls).toEqual(["inspect", "execute"]);
    expect(result.outcome).toBe("committed");
    expect(result.verification).toBe("verified");
    expect(result.finalSummary).not.toContain("model claim");
  });

  it("rejects a second inspection and never mutates", async () => {
    const a = adapter(JSON.stringify({ status: "supported", snapshot: { kind: "sentinel", value: null } }));
    const result = await createWorkbookPlanTurn(deps([call(inspectPlan), call(inspectPlan)], {
      adapter: a.adapter,
      inspect: async () => { a.calls.push("inspect"); return "{}"; },
    }));
    expect(result.outcome).toBe("rejected");
    expect(a.calls).toEqual(["inspect"]);
  });

  it("returns ask/refuse without executor access", async () => {
    const a = adapter("");
    const result = await createWorkbookPlanTurn(deps([call(askPlan)], { adapter: a.adapter }));
    expect(result.outcome).toBe("ask");
    expect(a.calls).toEqual([]);
  });

  it("offers one correction and then fails closed for invalid protocol", async () => {
    const a = adapter("");
    const result = await createWorkbookPlanTurn(deps([{ content: "", toolCalls: [] }, { content: "", toolCalls: [] }], { adapter: a.adapter }));
    expect(result.outcome).toBe("rejected");
    expect(a.calls).toEqual([]);
  });

  it("fails before any Office dependency is touched for malformed trusted identity", async () => {
    let touched = false;
    await expect(createWorkbookPlanTurn(deps([call(askPlan)], { identity: { ...IDS, documentId: "bad" }, inspect: async () => { touched = true; return ""; } }))).resolves.toMatchObject({ outcome: "rejected" });
    expect(touched).toBe(false);
  });

  it("fails closed when the rollout flag is absent or false", async () => {
    let modelCalls = 0;
    const result = await createWorkbookPlanTurn(deps([call(askPlan)], {
      rolloutEnabled: false,
      model: async () => { modelCalls += 1; return call(askPlan); },
    }));
    expect(result.outcome).toBe("rejected");
    expect(modelCalls).toBe(0);
  });

  it("requires a trusted server turn ID and durably replays an exact duplicate ticket", async () => {
    const missing = await createWorkbookPlanTurn(deps([call(executePlan)], {
      route: "server-owned",
    }));
    expect(missing.outcome).toBe("rejected");

    const trustedTurnOperationId = "44444444-4444-4444-8444-444444444444";
    const storage = new MemoryIdempotencyStorage();
    const first = await createWorkbookPlanTurn(deps([call(executePlan)], {
      route: "server-owned",
      trustedTurnOperationId,
      store: new BrowserIdempotencyStore(storage),
    }));
    expect(first.outcome).toBe("committed");
    const second = await createWorkbookPlanTurn(deps([call(executePlan)], {
      route: "server-owned",
      trustedTurnOperationId,
      store: new BrowserIdempotencyStore(storage),
    }));
    expect(second.outcome).toBe("committed");
  });

  it("durably replays a second direct execution in the same trusted turn", async () => {
    const trustedTurnOperationId = "55555555-5555-4555-8555-555555555555";
    const storage = new MemoryIdempotencyStorage();
    const firstAdapter = adapter(JSON.stringify({ status: "supported", snapshot: { kind: "sentinel", value: null } }));
    const first = await createWorkbookPlanTurn(deps([call(executePlan)], {
      route: "direct",
      trustedTurnOperationId,
      adapter: firstAdapter.adapter,
      store: new BrowserIdempotencyStore(storage),
    }));
    expect(first.outcome).toBe("committed");

    const secondAdapter = adapter(JSON.stringify({ status: "supported", snapshot: { kind: "sentinel", value: null } }));
    let replayVerificationCalls = 0;
    const second = await createWorkbookPlanTurn(deps([call(executePlan)], {
      route: "direct",
      trustedTurnOperationId,
      adapter: {
        ...secondAdapter.adapter,
        verifyPlan: async () => { replayVerificationCalls += 1; return true; },
      },
      store: new BrowserIdempotencyStore(storage),
    }));
    expect(second.outcome).toBe("committed");
    expect(second.verification).toBe("verified");
    expect(replayVerificationCalls).toBe(1);
    expect(secondAdapter.calls).toEqual([]);
  });

  it("rejects execution-first followed by inspection on the shared turn state", async () => {
    const state = new WorkbookPlanTurnState();
    const trustedTurnOperationId = "66666666-6666-4666-8666-666666666666";
    const firstAdapter = adapter(JSON.stringify({ status: "supported", snapshot: { kind: "sentinel", value: null } }));
    const first = await createWorkbookPlanTurn(deps([call(executePlan)], {
      route: "direct",
      trustedTurnOperationId,
      turnState: state,
      adapter: firstAdapter.adapter,
    }));
    expect(first.outcome).toBe("committed");

    const secondAdapter = adapter("");
    const second = await createWorkbookPlanTurn(deps([call(inspectPlan)], {
      route: "direct",
      trustedTurnOperationId,
      turnState: state,
      adapter: secondAdapter.adapter,
      inspect: async () => { secondAdapter.calls.push("inspect"); return "{}"; },
    }));
    expect(second.outcome).toBe("rejected");
    expect(secondAdapter.calls).toEqual([]);
  });

  it("passes a trusted beforeFirstMutation gate to the executor: called once after fresh and before the first mutation", async () => {
    const order: string[] = [];
    const gated: WorkbookPlanExecutionAdapter & { verifyPlan: () => Promise<boolean> } = {
      async executeFirstGroup(gate) {
        order.push("fresh");
        return gate.executeAfterFresh("fresh", async () => { order.push("mutation"); return { status: "committed" }; });
      },
      async executeGroup() { order.push("mutation-extra"); return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch" as const); },
      async verifyPlan() { return true; },
    };
    const result = await createWorkbookPlanTurn(deps([call(executePlan)], {
      adapter: gated,
      route: "server-owned",
      trustedTurnOperationId: "88888888-8888-4888-8888-888888888888",
      beforeFirstMutation: async () => { order.push("acknowledge"); },
    } as Partial<WorkbookPlanTurnDependencies>));
    expect(result.outcome).toBe("committed");
    expect(order).toEqual(["fresh", "acknowledge", "mutation"]);
  });

  it("performs zero mutation when the beforeFirstMutation gate rejects, and keeps the AbortSignal semantics", async () => {
    const order: string[] = [];
    const gated: WorkbookPlanExecutionAdapter & { verifyPlan: () => Promise<boolean> } = {
      async executeFirstGroup(gate, signal) {
        order.push(signal === controller.signal ? "fresh-with-turn-signal" : "fresh-without-turn-signal");
        return gate.executeAfterFresh("fresh", async () => { order.push("mutation"); return { status: "committed" }; });
      },
      async executeGroup() { order.push("mutation-extra"); return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch" as const); },
      async verifyPlan() { return true; },
    };
    const controller = new AbortController();
    const result = await createWorkbookPlanTurn(deps([call(executePlan)], {
      adapter: gated,
      signal: controller.signal,
      route: "server-owned",
      trustedTurnOperationId: "99999999-9999-4999-8999-999999999999",
      beforeFirstMutation: async () => { order.push("acknowledge"); throw new Error("execute-start rejected"); },
    } as Partial<WorkbookPlanTurnDependencies>));
    expect(order).toEqual(["fresh-with-turn-signal", "acknowledge"]);
    expect(result.outcome).toBe("cancelled");
    expect(result.execution?.operations.every((operation) => operation.state === "not_committed")).toBe(true);
  });

  it("rejects server-owned inspect before access and terminals the shared state", async () => {
    const state = new WorkbookPlanTurnState();
    const trustedTurnOperationId = "77777777-7777-4777-8777-777777777777";
    const inspectAdapter = adapter("");
    const inspected = await createWorkbookPlanTurn(deps([], {
      route: "server-owned",
      trustedTurnOperationId,
      turnState: state,
      initialResponse: call(inspectPlan),
      continueAfterInspection: false,
      adapter: inspectAdapter.adapter,
      inspect: async () => { inspectAdapter.calls.push("inspect"); return "{}"; },
    }));
    expect(inspected.outcome).toBe("rejected");
    expect(inspectAdapter.calls).toEqual([]);

    const executeAdapter = adapter(JSON.stringify({ status: "supported", snapshot: { kind: "sentinel", value: null } }));
    const executed = await createWorkbookPlanTurn(deps([], {
      route: "server-owned",
      trustedTurnOperationId,
      turnState: state,
      initialResponse: call(executePlan),
      adapter: executeAdapter.adapter,
    }));
    expect(executed.outcome).toBe("rejected");
    expect(executeAdapter.calls).toEqual([]);
  });
});
