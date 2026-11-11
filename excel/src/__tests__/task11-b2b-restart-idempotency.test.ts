import { describe, expect, it } from "vitest";
import {
  BrowserIdempotencyStore,
  MemoryIdempotencyStorage,
  createTrustedOperationBoundary,
  executeAcceptedWorkbookPlan,
  prepareWorkbookPlanExecution,
  type IdempotencyStorage,
  type WorkbookPlanExecutionAdapter,
} from "../workbookPlanExecutor";
import { createWorkbookPlanTurn, type WorkbookPlanModelResponse } from "../workbookPlanOrchestrator";
import type { WorkbookOperation, WorkbookPlan, VerificationRule } from "../generated/workbookPlanSchema";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import type { WorkbookPlanPolicyInput } from "../workbookPlanPolicy";
import { WORKBOOK_PLAN_CAPABILITY_NAMES } from "../workbookPlanPolicy";

const IDS = {
  sessionId: "11111111-1111-4111-8111-111111111111",
  documentId: "22222222-2222-4222-8222-222222222222",
  turnId: "33333333-3333-4333-8333-333333333333",
};

const operation = (value: string): WorkbookOperation => ({
  kind: "write_values",
  sheet: "Sheet1",
  address: "A1",
  values: [[value]],
});

const planFor = (value: string): WorkbookPlan => ({
  version: 1,
  outcome: "execute",
  summary: "bounded recovery test plan",
  operations: [operation(value)],
  verification: [{ kind: "calculation_completed" }],
});

const compileContext = (): VerifiedWorkbookCompileContext => ({
  sheetNames: new Set(["Sheet1"]),
  tablesByName: new Map(),
  supportedPrimitiveNames: new Set(["write_range"]),
  supportedCapabilities: new Set([
    "excel_workbook_access",
    "workbook_plan_v1_operation_audit",
  ]),
  chartSourceResolutionsBySheet: new Map(),
});

const policyInput = (): WorkbookPlanPolicyInput => ({
  requestClassification: "explicit_mutation",
  activeSheetName: "Sheet1",
  selectionScopedRequest: false,
  operationSummaries: [{
    operationIndex: 0,
    effect: "write",
    target: { operationIndex: 0, sheetNames: ["Sheet1"], selectionRelationship: "inside" },
  }],
  capabilitySummaries: [...WORKBOOK_PLAN_CAPABILITY_NAMES],
  resourceUsage: {
    operationCount: 1,
    cellsRead: 0,
    cellsWritten: 1,
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
});

async function prepare(plan: WorkbookPlan, trustedOperationBoundary = createTrustedOperationBoundary(IDS)) {
  const result = await prepareWorkbookPlanExecution({
    plan,
    trustedOperationBoundary,
    compileContext: compileContext(),
    policyInput: policyInput(),
  });
  expect(result.accepted).toBe(true);
  if (!result.accepted) throw new Error("expected accepted transaction");
  return result.transaction;
}

function adapterFor(mutate: () => void, calls: string[] = []): WorkbookPlanExecutionAdapter {
  return {
    async executeFirstGroup(context) {
      calls.push("first");
      return context.executeAfterFresh("fresh", async () => {
        mutate();
        return { status: "committed" };
      });
    },
    async executeGroup() {
      calls.push("group");
      return { status: "committed" };
    },
    async reconcileFailedGroup(group) {
      calls.push("reconcile");
      return group.operations.map(() => "unknown_within_batch" as const);
    },
  };
}

function modelResponse(plan: WorkbookPlan): WorkbookPlanModelResponse {
  return {
    content: "",
    toolCalls: [{
      id: "b2b-call",
      type: "function",
      function: { name: "excel_execute_plan", arguments: JSON.stringify(plan) },
    }],
  };
}

function orchestratorAdapter(
  calls: string[],
  mutate: () => void,
  verify: (rules: readonly VerificationRule[], signal?: AbortSignal) => Promise<boolean> = async () => true,
) {
  return {
    ...adapterFor(mutate, calls),
    verifyPlan: async (rules: readonly VerificationRule[], signal?: AbortSignal) => {
      calls.push("verification");
      return verify(rules, signal);
    },
  };
}

function orchestratorDeps(
  boundary: ReturnType<typeof createTrustedOperationBoundary>,
  storage: IdempotencyStorage,
  plan: WorkbookPlan,
  route: "direct" | "server-owned",
  calls: string[],
  mutate: () => void,
  verify: (rules: readonly VerificationRule[], signal?: AbortSignal) => Promise<boolean> = async () => true,
  overrides: Partial<Parameters<typeof createWorkbookPlanTurn>[0]> = {},
) {
  return {
    identity: IDS,
    trustedOperationBoundary: boundary,
    compileContext: compileContext(),
    policyInput: policyInput(),
    store: new BrowserIdempotencyStore(storage),
    model: async () => modelResponse(plan),
    inspect: async () => "{}",
    resourcePreflightAfterRuntime: async () => ({ ok: true as const }),
    adapter: orchestratorAdapter(calls, mutate, verify),
    rolloutEnabled: true,
    route,
    trustedTurnOperationId: route === "server-owned" ? "44444444-4444-4444-8444-444444444444" : undefined,
    ...overrides,
  };
}

class FailingStorage implements IdempotencyStorage {
  getItem(): string | null { throw new Error("storage read failed"); }
  setItem(): void { throw new Error("storage write failed"); }
  removeItem(): void { throw new Error("storage remove failed"); }
}

async function waitFor(predicate: () => boolean): Promise<void> {
  for (let attempt = 0; attempt < 50; attempt += 1) {
    if (predicate()) return;
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  throw new Error("condition did not become true");
}

describe("Task 11 B2b restart-durable idempotency", () => {
  it("reconstructs the same trusted boundary and replays the exact terminal report without a second mutation", async () => {
    const storage = new MemoryIdempotencyStorage();
    const firstBoundary = createTrustedOperationBoundary(IDS);
    const secondBoundary = createTrustedOperationBoundary({ ...IDS });
    const firstTransaction = await prepare(planFor("terminal-canary"), firstBoundary);
    const secondTransaction = await prepare(planFor("terminal-canary"), secondBoundary);

    expect(secondTransaction.transactionId).toBe(firstTransaction.transactionId);
    expect(secondTransaction.semanticDigest).toBe(firstTransaction.semanticDigest);

    let mutations = 0;
    const firstReport = await executeAcceptedWorkbookPlan(
      firstTransaction,
      firstBoundary,
      adapterFor(() => { mutations += 1; }),
      new BrowserIdempotencyStore(storage),
    );
    const replayCalls: string[] = [];
    const replayReport = await executeAcceptedWorkbookPlan(
      secondTransaction,
      secondBoundary,
      adapterFor(() => { mutations += 1; }, replayCalls),
      new BrowserIdempotencyStore(storage),
    );

    expect(replayReport).toEqual(firstReport);
    expect(replayReport.operations.map((item) => item.operationId)).toEqual(
      firstReport.operations.map((item) => item.operationId),
    );
    expect(mutations).toBe(1);
    expect(replayCalls).toEqual([]);
  });

  it("rejects a reconstructed runtime forever while the persisted record is in flight", async () => {
    let now = 0;
    const storage = new MemoryIdempotencyStorage();
    const boundary = createTrustedOperationBoundary(IDS);
    const reconstructedBoundary = createTrustedOperationBoundary({ ...IDS });
    const transaction = await prepare(planFor("in-flight-canary"), boundary);
    const reconstructed = await prepare(planFor("in-flight-canary"), reconstructedBoundary);
    let release!: () => void;
    const handoff = new Promise<void>((resolve) => { release = resolve; });
    let mutations = 0;
    const first = executeAcceptedWorkbookPlan(
      transaction,
      boundary,
      {
        ...adapterFor(() => { mutations += 1; }),
        async executeFirstGroup(context) {
          return context.executeAfterFresh("fresh", async () => {
            await handoff;
            mutations += 1;
            return { status: "committed" };
          });
        },
      },
      new BrowserIdempotencyStore(storage, { clock: () => now, ttlMs: 1 }),
    );

    await waitFor(() => storage.getItem("hermes.excel.workbookPlan.idempotency.v1")?.includes('"state":"in_flight"') === true);
    const restarted = new BrowserIdempotencyStore(storage, { clock: () => now, ttlMs: 1 });
    const restartCalls: string[] = [];
    const restartReport = await executeAcceptedWorkbookPlan(
      reconstructed,
      reconstructedBoundary,
      adapterFor(() => { mutations += 1; }, restartCalls),
      restarted,
    );
    expect(restartReport.outcome).toBe("restart_rejected");
    expect(restartCalls).toEqual([]);

    now = 100;
    expect((await restarted.claim(
      `${IDS.sessionId}.${IDS.documentId}.${IDS.turnId}`,
      reconstructed.transactionId,
      reconstructed.semanticDigest,
    )).kind).toBe("restart_rejected");
    expect(storage.getItem("hermes.excel.workbookPlan.idempotency.v1")).toContain('"state":"in_flight"');

    release();
    const firstReport = await first;
    expect(firstReport.outcome).toBe("committed");
    expect(mutations).toBe(1);
  });

  it("conflicts atomically on a changed semantic digest before adapter access", async () => {
    const storage = new MemoryIdempotencyStorage();
    const firstBoundary = createTrustedOperationBoundary(IDS);
    const changedBoundary = createTrustedOperationBoundary({ ...IDS });
    const first = await prepare(planFor("first-semantic-plan"), firstBoundary);
    const changed = await prepare(planFor("changed-semantic-plan"), changedBoundary);
    expect(changed.transactionId).not.toBe(first.transactionId);

    await expect(executeAcceptedWorkbookPlan(
      first,
      firstBoundary,
      adapterFor(() => undefined),
      new BrowserIdempotencyStore(storage),
    )).resolves.toMatchObject({ outcome: "committed" });

    let changedMutations = 0;
    const changedCalls: string[] = [];
    const report = await executeAcceptedWorkbookPlan(
      changed,
      changedBoundary,
      adapterFor(() => { changedMutations += 1; }, changedCalls),
      new BrowserIdempotencyStore(storage),
    );
    expect(report.outcome).toBe("conflict");
    expect(changedMutations).toBe(0);
    expect(changedCalls).toEqual([]);
  });

  it("serializes independently owned runtimes so exactly one can mutate", async () => {
    const storage = new MemoryIdempotencyStorage();
    const boundaryA = createTrustedOperationBoundary(IDS);
    const boundaryB = createTrustedOperationBoundary({ ...IDS });
    const transactionA = await prepare(planFor("race-canary"), boundaryA);
    const transactionB = await prepare(planFor("race-canary"), boundaryB);
    let mutations = 0;
    const reports = await Promise.all([
      executeAcceptedWorkbookPlan(transactionA, boundaryA, adapterFor(() => { mutations += 1; }), new BrowserIdempotencyStore(storage)),
      executeAcceptedWorkbookPlan(transactionB, boundaryB, adapterFor(() => { mutations += 1; }), new BrowserIdempotencyStore(storage)),
    ]);
    expect(mutations).toBe(1);
    expect(reports.filter((report) => report.outcome === "committed")).toHaveLength(1);
    expect(reports.some((report) => report.outcome === "restart_rejected" || report.outcome === "committed")).toBe(true);
  });

  it("keeps terminal bounds explicit, never evicts in-flight, and fails closed on malformed or unavailable storage", async () => {
    let now = 0;
    const storage = new MemoryIdempotencyStorage();
    const boundary = createTrustedOperationBoundary(IDS);
    const transaction = await prepare(planFor("bounded-content-canary"), boundary);
    const bounded = new BrowserIdempotencyStore(storage, { clock: () => now, ttlMs: 1, maxEntries: 1, maxBytes: 4096 });
    await expect(executeAcceptedWorkbookPlan(transaction, boundary, adapterFor(() => undefined), bounded)).resolves.toMatchObject({ outcome: "committed" });
    const persisted = storage.getItem("hermes.excel.workbookPlan.idempotency.v1")!;
    expect(persisted).not.toContain("bounded-content-canary");
    expect(persisted).not.toContain("values");

    now = 100;
    const secondBoundary = createTrustedOperationBoundary({ ...IDS, turnId: "55555555-5555-4555-8555-555555555555" });
    const second = await prepare(planFor("second-bounded-plan"), secondBoundary);
    expect((await bounded.claim(
      "11111111-1111-4111-8111-111111111111.22222222-2222-4222-8222-222222222222.55555555-5555-4555-8555-555555555555",
      second.transactionId,
      second.semanticDigest,
    )).kind).toBe("claimed");
    now = 200;
    expect((await bounded.claim(
      "11111111-1111-4111-8111-111111111111.22222222-2222-4222-8222-222222222222.55555555-5555-4555-8555-555555555555",
      second.transactionId,
      second.semanticDigest,
    )).kind).toBe("in_flight");
    expect(storage.getItem("hermes.excel.workbookPlan.idempotency.v1")).toContain('"state":"in_flight"');

    expect((await new BrowserIdempotencyStore(new FailingStorage()).claim(
      `${IDS.sessionId}.${IDS.documentId}.${IDS.turnId}`,
      transaction.transactionId,
      transaction.semanticDigest,
    )).kind).toBe("persistence_unavailable");
    expect((await new BrowserIdempotencyStore(new MemoryIdempotencyStorage()).claim(
      `${IDS.sessionId}.${IDS.documentId}.${IDS.turnId}`,
      transaction.transactionId,
      transaction.semanticDigest,
    )).kind).toBe("claimed");
  });

  it("governs reconstructed direct and server-owned requests with durable storage, not process-local duplicate sets", async () => {
    const directStorage = new MemoryIdempotencyStorage();
    const directPlan = planFor("direct-replay");
    let directMutations = 0;
    const firstDirectCalls: string[] = [];
    const firstDirect = await createWorkbookPlanTurn(orchestratorDeps(
      createTrustedOperationBoundary(IDS), directStorage, directPlan, "direct", firstDirectCalls, () => { directMutations += 1; },
    ));
    const secondDirectCalls: string[] = [];
    const secondDirect = await createWorkbookPlanTurn(orchestratorDeps(
      createTrustedOperationBoundary({ ...IDS }), directStorage, directPlan, "direct", secondDirectCalls, () => { directMutations += 1; },
    ));
    expect(firstDirect.outcome).toBe("committed");
    expect(secondDirect.outcome).toBe("committed");
    expect(directMutations).toBe(1);
    expect(secondDirectCalls).toEqual(["verification"]);

    const serverStorage = new MemoryIdempotencyStorage();
    const serverPlan = planFor("server-replay");
    let serverMutations = 0;
    const firstServer = await createWorkbookPlanTurn(orchestratorDeps(
      createTrustedOperationBoundary(IDS), serverStorage, serverPlan, "server-owned", [], () => { serverMutations += 1; },
    ));
    const secondServerCalls: string[] = [];
    const secondServer = await createWorkbookPlanTurn(orchestratorDeps(
      createTrustedOperationBoundary({ ...IDS }), serverStorage, serverPlan, "server-owned", secondServerCalls, () => { serverMutations += 1; },
    ));
    expect(firstServer.outcome).toBe("committed");
    expect(secondServer.outcome).toBe("committed");
    expect(serverMutations).toBe(1);
    expect(secondServerCalls).toEqual(["verification"]);
  });

  it("does not upgrade a verification failure when a reconstructed exact replay reuses the terminal executor report", async () => {
    const storage = new MemoryIdempotencyStorage();
    let mutations = 0;
    let firstVerificationCalls = 0;
    const firstCalls: string[] = [];
    const first = await createWorkbookPlanTurn(orchestratorDeps(
      createTrustedOperationBoundary(IDS),
      storage,
      planFor("verification-failed-replay"),
      "direct",
      firstCalls,
      () => { mutations += 1; },
      async () => { firstVerificationCalls += 1; return false; },
    ));
    expect(first.outcome).toBe("verification_failed");
    expect(first.verification).toBe("failed");
    expect(firstVerificationCalls).toBe(1);
    expect(mutations).toBe(1);

    let replayVerificationCalls = 0;
    const replayCalls: string[] = [];
    const replay = await createWorkbookPlanTurn(orchestratorDeps(
      createTrustedOperationBoundary({ ...IDS }),
      storage,
      planFor("verification-failed-replay"),
      "direct",
      replayCalls,
      () => { mutations += 1; },
      async () => { replayVerificationCalls += 1; return false; },
    ));

    expect(replay.outcome).toBe("verification_failed");
    expect(replay.verification).toBe("failed");
    expect(replay.execution?.outcome).toBe("committed");
    expect(replay.execution?.verification).toBe("not_attempted");
    expect(replayVerificationCalls).toBe(1);
    expect(mutations).toBe(1);
    expect(replayCalls).toEqual(["verification"]);
  });

  it("re-verifies an exact committed replay with compiler postconditions and its exact turn signal", async () => {
    const storage = new MemoryIdempotencyStorage();
    const plan = planFor("verified-replay");
    let mutations = 0;
    const firstController = new AbortController();
    const firstRules: VerificationRule[][] = [];
    const firstSignals: AbortSignal[] = [];
    const first = await createWorkbookPlanTurn(orchestratorDeps(
      createTrustedOperationBoundary(IDS),
      storage,
      plan,
      "direct",
      [],
      () => { mutations += 1; },
      async (rules, signal) => {
        firstRules.push([...rules]);
        firstSignals.push(signal!);
        return true;
      },
      { signal: firstController.signal },
    ));
    expect(first.outcome).toBe("committed");
    expect(first.verification).toBe("verified");
    expect(firstSignals).toEqual([firstController.signal]);

    const replayController = new AbortController();
    const replayRules: VerificationRule[][] = [];
    const replaySignals: AbortSignal[] = [];
    let replayVerificationCalls = 0;
    const replayCalls: string[] = [];
    const replay = await createWorkbookPlanTurn(orchestratorDeps(
      createTrustedOperationBoundary({ ...IDS }),
      storage,
      plan,
      "direct",
      replayCalls,
      () => { mutations += 1; },
      async (rules, signal) => {
        replayVerificationCalls += 1;
        replayRules.push([...rules]);
        replaySignals.push(signal!);
        return true;
      },
      { signal: replayController.signal },
    ));

    expect(replay.outcome).toBe("committed");
    expect(replay.verification).toBe("verified");
    expect(replayVerificationCalls).toBe(1);
    expect(replayRules).toEqual(firstRules);
    expect(replaySignals).toEqual([replayController.signal]);
    expect(mutations).toBe(1);
    expect(replayCalls).toEqual(["verification"]);
  });

  it("cannot upgrade an exact replay to success after its verifier aborts the linked turn signal", async () => {
    const storage = new MemoryIdempotencyStorage();
    const plan = planFor("cancelled-replay");
    let mutations = 0;
    const first = await createWorkbookPlanTurn(orchestratorDeps(
      createTrustedOperationBoundary(IDS),
      storage,
      plan,
      "direct",
      [],
      () => { mutations += 1; },
    ));
    expect(first.outcome).toBe("committed");

    const replayController = new AbortController();
    const replayCalls: string[] = [];
    let replayVerificationCalls = 0;
    const replay = await createWorkbookPlanTurn(orchestratorDeps(
      createTrustedOperationBoundary({ ...IDS }),
      storage,
      plan,
      "direct",
      replayCalls,
      () => { mutations += 1; },
      async (_rules, signal) => {
        replayVerificationCalls += 1;
        expect(signal).toBe(replayController.signal);
        replayController.abort();
        return true;
      },
      { signal: replayController.signal },
    ));

    expect(replay.outcome).toBe("cancelled");
    expect(replay.verification).toBe("not_attempted");
    expect(replayVerificationCalls).toBe(1);
    expect(mutations).toBe(1);
    expect(replayCalls).toEqual(["verification"]);
  });

  it("does not verify changed-digest conflicts or reconstructed in-flight rejections", async () => {
    const storage = new MemoryIdempotencyStorage();
    let mutations = 0;
    const first = await createWorkbookPlanTurn(orchestratorDeps(
      createTrustedOperationBoundary(IDS),
      storage,
      planFor("conflict-first-plan"),
      "direct",
      [],
      () => { mutations += 1; },
    ));
    expect(first.outcome).toBe("committed");

    let conflictVerificationCalls = 0;
    const conflictCalls: string[] = [];
    const conflict = await createWorkbookPlanTurn(orchestratorDeps(
      createTrustedOperationBoundary({ ...IDS }),
      storage,
      planFor("conflict-changed-plan"),
      "direct",
      conflictCalls,
      () => { mutations += 1; },
      async () => { conflictVerificationCalls += 1; return true; },
    ));
    expect(conflict.execution?.outcome).toBe("conflict");
    expect(conflict.verification).toBe("not_attempted");
    expect(conflictVerificationCalls).toBe(0);
    expect(conflictCalls).toEqual([]);
    expect(mutations).toBe(1);

    const inFlightStorage = new MemoryIdempotencyStorage();
    let release!: () => void;
    const handoff = new Promise<void>((resolve) => { release = resolve; });
    const firstInFlight = createWorkbookPlanTurn(orchestratorDeps(
      createTrustedOperationBoundary(IDS),
      inFlightStorage,
      planFor("in-flight-replay"),
      "direct",
      [],
      () => { mutations += 1; },
      async () => true,
      {
        adapter: {
          ...orchestratorAdapter([], () => undefined),
          async executeFirstGroup(context) {
            return context.executeAfterFresh("fresh", async () => {
              await handoff;
              mutations += 1;
              return { status: "committed" };
            });
          },
        },
      },
    ));
    await waitFor(() => inFlightStorage.getItem("hermes.excel.workbookPlan.idempotency.v1")?.includes('"state":"in_flight"') === true);

    let restartVerificationCalls = 0;
    const restartCalls: string[] = [];
    try {
      const restart = await createWorkbookPlanTurn(orchestratorDeps(
        createTrustedOperationBoundary({ ...IDS }),
        inFlightStorage,
        planFor("in-flight-replay"),
        "direct",
        restartCalls,
        () => { mutations += 1; },
        async () => { restartVerificationCalls += 1; return true; },
      ));
      expect(restart.execution?.outcome).toBe("restart_rejected");
      expect(restart.verification).toBe("not_attempted");
      expect(restartVerificationCalls).toBe(0);
      expect(restartCalls).toEqual([]);
      expect(mutations).toBe(1);
    } finally {
      release();
      expect((await firstInFlight).outcome).toBe("committed");
    }
  });
});
