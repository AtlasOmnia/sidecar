import { describe, expect, it } from "vitest";
import {
  BrowserIdempotencyStore,
  DeterministicIdempotencyCoordinator,
  MemoryIdempotencyStorage,
  createTrustedOperationBoundary,
  executeAcceptedWorkbookPlan,
  getAcceptedWorkbookPlanPostconditionsForVerification,
  prepareWorkbookPlanExecution,
  type IdempotencyStorage,
  type WorkbookPlanExecutionAdapter,
  type WorkbookPlanExecutionReport,
  type TrustedOperationBoundary,
} from "../workbookPlanExecutor";
import type { WorkbookPlanPolicyInput } from "../workbookPlanPolicy";
import { getExcelToolDefinition } from "../excelTools";
import type { WorkbookOperation, WorkbookPlan } from "../generated/workbookPlanSchema";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import { WORKBOOK_PLAN_CAPABILITY_NAMES } from "../workbookPlanPolicy";

const UUIDS = [
  "11111111-1111-4111-8111-111111111111",
  "22222222-2222-4222-8222-222222222222",
  "33333333-3333-4333-8333-333333333333",
  "44444444-4444-4444-8444-444444444444",
  "55555555-5555-4555-8555-555555555555",
  "66666666-6666-4666-8666-666666666666",
  "77777777-7777-4777-8777-777777777777",
  "88888888-8888-4888-8888-888888888888",
];

const BOUNDARY_A = createTrustedOperationBoundary({ sessionId: UUIDS[0], documentId: UUIDS[1], turnId: UUIDS[2] });
const BOUNDARY_B = createTrustedOperationBoundary({ sessionId: UUIDS[0], documentId: UUIDS[1], turnId: UUIDS[7] });
const BOUNDARY_C = createTrustedOperationBoundary({ sessionId: UUIDS[0], documentId: UUIDS[1], turnId: UUIDS[6] });
const TOKEN_A = "11111111-1111-4111-8111-111111111111.22222222-2222-4222-8222-222222222222.33333333-3333-4333-8333-333333333333";
const TOKEN_B = "11111111-1111-4111-8111-111111111111.22222222-2222-4222-8222-222222222222.88888888-8888-4888-8888-888888888888";const uuidSource = (): (() => string) => {
  let index = 0;
  return () => UUIDS[index++] ?? "99999999-9999-4999-8999-999999999999";
};

const operation = (
  address: string,
  value: string,
): WorkbookOperation => ({
  kind: "write_values",
  sheet: "Sheet1",
  address,
  values: [[value]],
});

const formatOperation = (address: string): WorkbookOperation => ({
  kind: "format_range",
  sheet: "Sheet1",
  address,
  format: { numberFormat: "0.00" },
});

const planFor = (operations: readonly WorkbookOperation[]): WorkbookPlan => ({
  version: 1,
  outcome: "execute",
  summary: "test plan",
  operations,
  verification: [{ kind: "calculation_completed" }],
});

const primitiveNames = [
  "write_range",
  "format_range",
  "autofill_formula",
  "insert_table",
  "sort_table",
  "filter_table",
  "create_chart",
  "set_data_validation",
  "conditional_format_range",
  "freeze_panes",
  "create_sheet",
  "rename_sheet",
  "copy_sheet",
  "add_comment",
  "run_calculation",
];

const compileContext = (): VerifiedWorkbookCompileContext => ({
  sheetNames: new Set(["Sheet1"]),
  tablesByName: new Map(),
  supportedPrimitiveNames: new Set(
    primitiveNames.filter((name) => getExcelToolDefinition(name) !== undefined),
  ),
  supportedCapabilities: new Set([
    "excel_workbook_access",
    "workbook_plan_v1_operation_audit",
  ]),
  chartSourceResolutionsBySheet: new Map(),
});

const policyInputFor = (
  effects: readonly ("write" | "format" | "structure" | "filter" | "calculate")[],
  operationCount = effects.length,
): WorkbookPlanPolicyInput => ({
  requestClassification: "explicit_mutation",
  activeSheetName: "Sheet1",
  selectionScopedRequest: false,
  operationSummaries: effects.map((effect, operationIndex) => ({
    operationIndex,
    effect,
    target: {
      operationIndex,
      sheetNames: effect === "calculate" ? [] : ["Sheet1"],
      selectionRelationship: "inside",
    },
  })),
  capabilitySummaries: [...WORKBOOK_PLAN_CAPABILITY_NAMES],
  resourceUsage: {
    operationCount,
    cellsRead: 0,
    cellsWritten: operationCount,
    formulaUtf8Bytes: 0,
    csvPayloadUtf8Bytes: 0,
    commentUtf8Bytes: 0,
    sheetsCreated: 0,
    chartsCreated: 0,
    tablesCreated: 0,
    verificationReads: 0,
  },
  resourceLimits: {
    operationCount: 100,
    cellsRead: 100,
    cellsWritten: 100,
    formulaUtf8Bytes: 10000,
    csvPayloadUtf8Bytes: 10000,
    commentUtf8Bytes: 10000,
    sheetsCreated: 100,
    chartsCreated: 100,
    tablesCreated: 100,
    verificationReads: 100,
  },
});

interface AdapterOptions {
  readonly failGroup?: number;
  readonly reconciliation?: readonly ("committed" | "not_committed" | "unknown_within_batch")[];
  readonly stale?: boolean;
  readonly onPrecondition?: () => Promise<"fresh" | "stale">;
}

const makeAdapter = (options: AdapterOptions = {}): {
  adapter: WorkbookPlanExecutionAdapter;
  calls: string[];
  groups: number[][];
} => {
  const calls: string[] = [];
  const groups: number[][] = [];
  const adapter: WorkbookPlanExecutionAdapter = {
    async executeFirstGroup(context) {
      calls.push("first");
      const precondition = options.onPrecondition === undefined
        ? options.stale ? "stale" : "fresh"
        : await options.onPrecondition();
      return context.executeAfterFresh(precondition, async (group) => {
        groups.push(group.operations.map((item) => item.operationIndex));
        if (options.failGroup === 0) return { status: "sync_failed", failureCode: "sync_failed" };
        return { status: "committed" };
      });
    },
    async executeGroup(group, groupIndex) {
      calls.push(`group:${groupIndex}`);
      groups.push(group.operations.map((item) => item.operationIndex));
      if (options.failGroup === groupIndex) return { status: "sync_failed", failureCode: "sync_failed" };
      return { status: "committed" };
    },
    async reconcileFailedGroup(group) {
      calls.push("reconcile");
      return options.reconciliation ?? group.operations.map(() => "unknown_within_batch");
    },
  };
  return { adapter, calls, groups };
};

const prepare = async (
  plan: unknown,
  effects: readonly ("write" | "format" | "structure" | "filter" | "calculate")[],
  extra: Partial<Parameters<typeof prepareWorkbookPlanExecution>[0]> = {},
) => prepareWorkbookPlanExecution({
  plan,
  compileContext: compileContext(),
  policyInput: policyInputFor(effects),
  trustedOperationBoundary: extra.trustedOperationBoundary ?? BOUNDARY_A,
  uuidSource: uuidSource(),
  ...extra,
});

const accepted = async (
  plan: WorkbookPlan,
  effects: readonly ("write" | "format" | "structure" | "filter" | "calculate")[] = ["write"],
  extra: Partial<Parameters<typeof prepareWorkbookPlanExecution>[0]> = {},
) => {
  const result = await prepare(plan, effects, extra);
  expect(result.accepted).toBe(true);
  if (!result.accepted) throw new Error("expected accepted plan");
  return result.transaction;
};

type TestExecutionOptions = {
  readonly signal?: AbortSignal;
  readonly beforeFirstMutation?: unknown;
};

const execute = (
  transaction: Awaited<ReturnType<typeof accepted>>,
  adapter: WorkbookPlanExecutionAdapter,
  store = new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
  boundary: TrustedOperationBoundary = BOUNDARY_A,
  options?: TestExecutionOptions,
): Promise<WorkbookPlanExecutionReport> =>
  executeAcceptedWorkbookPlan(
    transaction,
    boundary,
    adapter,
    store,
    options as unknown as Parameters<typeof executeAcceptedWorkbookPlan>[4],
  );

const deferred = <T>() => {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
};

class RawStorage implements IdempotencyStorage {
  constructor(private value: string | null = null, private readonly failGet = false, private readonly failSet = false) {}

  getItem(): string | null {
    if (this.failGet) throw new Error("read failed");
    return this.value;
  }

  setItem(_key: string, value: string): void {
    if (this.failSet) throw new Error("write failed");
    this.value = value;
  }

  removeItem(): void {}

  read(): string | null {
    return this.value;
  }
}

class EventStorage implements IdempotencyStorage {
  private value: string | null = null;

  constructor(private readonly events: string[]) {}

  getItem(): string | null {
    return this.value;
  }

  setItem(_key: string, value: string): void {
    this.events.push(this.value === null ? "durable-claim" : "durable-finish");
    this.value = value;
  }

  removeItem(): void {
    this.value = null;
  }

  read(): string | null {
    return this.value;
  }
}

describe("workbook plan executor entry", () => {
  it("validates the complete plan before mutation, including later operations", async () => {
    const malformed = {
      ...planFor([operation("A1", "safe")]),
      operations: [operation("A1", "safe"), { kind: "write_values", bogus: "bad" }],
    };
    const { adapter, calls } = makeAdapter();
    const result = await prepare(malformed, ["write", "write"]);
    expect(result.accepted).toBe(false);
    expect(await Promise.resolve(calls)).toEqual([]);
    expect(adapter).toBeDefined();
  });

  it("rejects an unapproved Task 7 policy without adapter calls", async () => {
    const result = await prepare(planFor([operation("A1", "safe")]), ["write"], {
      policyInput: policyInputFor(["write"], 2),
    });
    expect(result).toEqual({ accepted: false, reason: "policy_rejected", policyReason: "inconsistent_operation_counts" });
  });

  it("rechecks stale preconditions in the distinct first-group batch and mutates zero times", async () => {
    const transaction = await accepted(planFor([operation("A1", "safe")]), ["write"]);
    let mutations = 0;
    const { adapter, calls } = makeAdapter({ stale: true });
    const report = await execute(transaction, {
      ...adapter,
      async executeFirstGroup(context) {
        const result = await adapter.executeFirstGroup(context);
        return result;
      },
    });
    mutations += report.operations.filter((item) => item.state === "committed").length;
    expect(calls).toEqual(["first"]);
    expect(mutations).toBe(0);
    expect(report.outcome).toBe("stale");
    expect(report.operations[0].state).toBe("not_committed");
  });

  it("groups only consecutive effects and preserves primitive ordering", async () => {
    const plan = planFor([
      operation("A1", "a"), operation("B1", "b"), formatOperation("C1"),
      formatOperation("D1"), operation("E1", "e"),
    ]);
    const transaction = await accepted(plan, ["write", "write", "format", "format", "write"]);
    const { adapter, calls, groups } = makeAdapter();
    const report = await execute(transaction, adapter);
    expect(calls).toEqual(["first", "group:1", "group:2"]);
    expect(groups).toEqual([[0, 1], [2, 3], [4]]);
    expect(report.operations.map((item) => item.operationIndex)).toEqual([0, 1, 2, 3, 4]);
    expect(report.operations.every((item) => item.state === "committed")).toBe(true);
  });

  it("reconciles a failed group, reports mixed states, and never starts later groups", async () => {
    const plan = planFor([operation("A1", "a"), formatOperation("B1"), formatOperation("C1"), operation("D1", "d")]);
    const transaction = await accepted(plan, ["write", "format", "format", "write"]);
    const { adapter, calls } = makeAdapter({
      failGroup: 1,
      reconciliation: ["committed", "unknown_within_batch"],
    });
    const report = await execute(transaction, adapter);
    expect(calls).toEqual(["first", "group:1", "reconcile"]);
    expect(report.outcome).toBe("partial");
    expect(report.firstFailedGroup).toBe(1);
    expect(report.firstFailedOperationIndex).toBe(1);
    expect(report.operations.map((item) => item.state)).toEqual([
      "committed", "committed", "unknown_within_batch", "not_committed",
    ]);
    expect(report.operations[1].status).toBe("applied");
    expect(report.operations[2].status).toBe("unknown_within_batch");
    expect(report.verification).toBe("reconciled");
  });

  it("replays exact identity from the terminal cache and conflicts on a digest mismatch", async () => {
    const store = new BrowserIdempotencyStore(new MemoryIdempotencyStorage());
    const first = await accepted(planFor([operation("A1", "a")]), ["write"]);
    const second = await accepted(planFor([operation("A1", "b")]), ["write"]);
    const firstAdapter = makeAdapter();
    const firstReport = await execute(first, firstAdapter.adapter, store);
    const replayAdapter = makeAdapter();
    const replay = await execute(first, replayAdapter.adapter, store);
    const conflict = await execute(second, makeAdapter().adapter, store);
    expect(firstReport.outcome).toBe("committed");
    expect(replay.outcome).toBe("committed");
    expect(replay.operations).toEqual(firstReport.operations);
    expect(replayAdapter.calls).toEqual([]);
    expect(conflict.outcome).toBe("conflict");
  });

  it("allows a semantically identical write in a later trusted boundary", async () => {
    const store = new BrowserIdempotencyStore(new MemoryIdempotencyStorage());
    const first = await accepted(planFor([operation("A1", "a")]), ["write"], { trustedOperationBoundary: BOUNDARY_A });
    const later = await accepted(planFor([operation("A1", "a")]), ["write"], { trustedOperationBoundary: BOUNDARY_B });
    const firstAdapter = makeAdapter();
    await execute(first, firstAdapter.adapter, store, BOUNDARY_A);
    const laterAdapter = makeAdapter();
    const report = await execute(later, laterAdapter.adapter, store, BOUNDARY_B);
    expect(report.outcome).toBe("committed");
    expect(laterAdapter.calls).toEqual(["first"]);
  });

  it("persists terminal results and rejects restart of an in-flight claim", async () => {
    const storage = new MemoryIdempotencyStorage();
    const firstStore = new BrowserIdempotencyStore(storage);
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"]);
    const committed = await execute(transaction, makeAdapter().adapter, firstStore);
    expect((await firstStore.claim(TOKEN_A, transaction.transactionId, transaction.semanticDigest)).kind).toBe("terminal");
    const rehydrated = new BrowserIdempotencyStore(storage);
    const replayAdapter = makeAdapter();
    expect((await execute(transaction, replayAdapter.adapter, rehydrated)).outcome).toBe("committed");
    expect(replayAdapter.calls).toEqual([]);

    const secondTransaction = await accepted(planFor([operation("B1", "b")]), ["write"], { trustedOperationBoundary: BOUNDARY_B });
    const claimStore = new BrowserIdempotencyStore(storage);
    expect((await claimStore.claim(TOKEN_B, secondTransaction.transactionId, secondTransaction.semanticDigest)).kind).toBe("claimed");
    const restarted = new BrowserIdempotencyStore(storage);
    const restartReport = await execute(secondTransaction, makeAdapter().adapter, restarted, BOUNDARY_B);
    expect(restartReport.outcome).toBe("restart_rejected");
    expect(committed.operations[0].state).toBe("committed");
  });

  it("evicts bounded terminal entries by TTL/count/bytes but never reclaims in-flight", async () => {
    let now = 0;
    const storage = new MemoryIdempotencyStorage();
    const store = new BrowserIdempotencyStore(storage, {
      clock: () => now,
      ttlMs: 10,
      maxEntries: 1,
      maxBytes: 4096,
    });
    const first = await accepted(planFor([operation("A1", "a")]), ["write"]);
    const second = await accepted(planFor([operation("B1", "b")]), ["write"]);
    await execute(first, makeAdapter().adapter, store);
    now = 20;
    expect((await store.claim(TOKEN_A, first.transactionId, first.semanticDigest)).kind).toBe("claimed");
    const inFlight = await accepted(planFor([operation("C1", "c")]), ["write"], { trustedOperationBoundary: BOUNDARY_B });
    expect((await store.claim(TOKEN_B, inFlight.transactionId, inFlight.semanticDigest)).kind).toBe("persistence_unavailable");
    now = 30;
    await execute(second, makeAdapter().adapter, store);
    const restarted = new BrowserIdempotencyStore(storage, { clock: () => now, ttlMs: 10, maxEntries: 1, maxBytes: 4096 });
    expect((await restarted.claim(TOKEN_A, first.transactionId, first.semanticDigest)).kind).toBe("restart_rejected");
  });

  it("rolls back only with complete, independently verified inverse safety", async () => {
    const inverse = { completeCapturedPreState: true, independentlyVerifiedSafe: true, rollback: async () => true };
    const transaction = await accepted(planFor([operation("A1", "a"), formatOperation("B1")]), ["write", "format"], { inverse });
    const report = await execute(transaction, makeAdapter({ failGroup: 1, reconciliation: ["unknown_within_batch"] }).adapter);
    expect(report.rollback).toBe("succeeded");
    expect(report.operations[0].status).toBe("rolled_back");
    expect(report.outcome).toBe("partial");

    const ineligible = await accepted(planFor([operation("A1", "a"), formatOperation("B1")]), ["write", "format"], {
      inverse: { completeCapturedPreState: false, independentlyVerifiedSafe: true, rollback: async () => true },
    });
    const ineligibleReport = await execute(ineligible, makeAdapter({ failGroup: 1 }).adapter);
    expect(ineligibleReport.rollback).toBe("unavailable");
  });

  it("reports rollback failure without upgrading partial application", async () => {
    const transaction = await accepted(planFor([operation("A1", "a"), formatOperation("B1")]), ["write", "format"], {
      inverse: { completeCapturedPreState: true, independentlyVerifiedSafe: true, rollback: async () => false },
    });
    const report = await execute(transaction, makeAdapter({ failGroup: 1 }).adapter);
    expect(report.rollback).toBe("failed");
    expect(report.outcome).toBe("partial");
  });

  it("rejects a directly forged accepted object", async () => {
    const forged = Object.freeze({ transactionId: UUIDS[0], semanticDigest: "digest", operationCount: 1 });
    const report = await execute(forged, makeAdapter().adapter);
    expect(report.outcome).toBe("rejected");
  });

  it("uses an actual adapter mutation counter for a TOCTOU change between preflight and execution", async () => {
    const workbook = { version: 1, mutations: 0 };
    const transaction = await accepted(planFor([operation("A1", "safe")]), ["write"], {
      preconditionCheck: async () => workbook.version === 1 ? "fresh" : "stale",
    });
    const adapter: WorkbookPlanExecutionAdapter = {
      async executeFirstGroup(context) {
        workbook.version = 2;
        return context.executeAfterFresh("fresh", async (group) => {
          workbook.mutations += group.operations.length;
          return { status: "committed" };
        });
      },
      async executeGroup() { return { status: "committed" }; },
      async reconcileFailedGroup() { return ["not_committed"]; },
    };
    const report = await execute(transaction, adapter);
    expect(workbook.mutations).toBe(0);
    expect(report.outcome).toBe("stale");
  });

  it("derives mid-batch reconciliation from actual fake-workbook state", async () => {
    const workbook = new Set<number>();
    let mutationCount = 0;
    const transaction = await accepted(planFor([
      operation("A1", "a"), formatOperation("B1"), formatOperation("C1"), operation("D1", "d"),
    ]), ["write", "format", "format", "write"]);
    const adapter: WorkbookPlanExecutionAdapter = {
      async executeFirstGroup(context) {
        return context.executeAfterFresh("fresh", async (group) => {
          for (const item of group.operations) { workbook.add(item.operationIndex); mutationCount += 1; }
          return { status: "committed" };
        });
      },
      async executeGroup(group, groupIndex) {
        if (groupIndex === 1) {
          workbook.add(group.operations[0].operationIndex);
          mutationCount += 1;
          return { status: "sync_failed", failureCode: "sync_failed" };
        }
        for (const item of group.operations) { workbook.add(item.operationIndex); mutationCount += 1; }
        return { status: "committed" };
      },
      async reconcileFailedGroup(group) {
        return group.operations.map((item) => workbook.has(item.operationIndex) ? "committed" : "not_committed");
      },
    };
    const report = await execute(transaction, adapter);
    expect(mutationCount).toBe(2);
    expect(report.operations.map((item) => item.state)).toEqual([
      "committed", "committed", "not_committed", "not_committed",
    ]);
    expect(report.operations[2].status).toBe("failed");
    expect(report.operations[3].status).toBe("not_started");
  });

  it("binds the accepted boundary before store or adapter access", async () => {
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"], { trustedOperationBoundary: BOUNDARY_A });
    const storage = new MemoryIdempotencyStorage();
    let mutations = 0;
    const adapter = makeAdapter();
    const countingAdapter: WorkbookPlanExecutionAdapter = {
      ...adapter.adapter,
      async executeFirstGroup(context) { mutations += 1; return adapter.adapter.executeFirstGroup(context); },
    };
    const wrong = await execute(transaction, countingAdapter, new BrowserIdempotencyStore(storage), BOUNDARY_B);
    expect(wrong.outcome).toBe("rejected");
    expect(mutations).toBe(0);
    expect(storage.getItem("hermes.excel.workbookPlan.idempotency.v1")).toBeNull();
    const first = await execute(transaction, countingAdapter, new BrowserIdempotencyStore(storage), BOUNDARY_A);
    expect(first.outcome).toBe("committed");
    const swapped = await execute(transaction, countingAdapter, new BrowserIdempotencyStore(storage), BOUNDARY_B);
    expect(swapped.outcome).toBe("rejected");
    expect(mutations).toBe(1);
  });

  it("fails closed for storage reads, malformed payloads, invalid records, and absent browser storage", async () => {
    const cases: BrowserIdempotencyStore[] = [
      new BrowserIdempotencyStore(new RawStorage(null, true)),
      new BrowserIdempotencyStore(new RawStorage("{}")),
      new BrowserIdempotencyStore(new RawStorage("[{\"state\":\"corrupt\"}]")),
      new BrowserIdempotencyStore(new RawStorage(JSON.stringify([{
        boundary: TOKEN_A, owner: UUIDS[4], fence: UUIDS[5], transactionId: UUIDS[0], digest: "digest", state: "terminal", createdAt: 0,
        report: { outcome: "content-bearing", transactionId: UUIDS[0], operationCount: 0, operations: [], verification: "not_attempted", rollback: "not_attempted" },
      }]))),
      new BrowserIdempotencyStore(),
    ];
    for (const store of cases) expect((await store.claim(TOKEN_A, UUIDS[0], "digest")).kind).toBe("persistence_unavailable");
  });

  it("fails closed when TTL cleanup cannot be persisted", async () => {
    const storage = new RawStorage(JSON.stringify([{
      boundary: TOKEN_A, owner: UUIDS[4], fence: UUIDS[5], transactionId: UUIDS[0], digest: "digest", state: "terminal", createdAt: 0,
      report: { outcome: "committed", transactionId: UUIDS[0], operationCount: 0, operations: [], verification: "not_attempted", rollback: "not_attempted" },
    }]), false, true);
    const store = new BrowserIdempotencyStore(storage, { clock: () => 100, ttlMs: 1 });
    expect((await store.claim(TOKEN_B, UUIDS[1], "digest")).kind).toBe("persistence_unavailable");
  });

  it("does not report durable success after post-mutation finish failure and never retries mutation", async () => {
    const storage = new RawStorage();
    const store = new BrowserIdempotencyStore(storage);
    const transaction = await accepted(planFor([operation("A1", "content-canary")]), ["write"]);
    let mutations = 0;
    const adapter: WorkbookPlanExecutionAdapter = {
      async executeFirstGroup(context) {
        return context.executeAfterFresh("fresh", async (group) => {
          mutations += group.operations.length;
          return { status: "committed" };
        });
      },
      async executeGroup() { return { status: "committed" }; },
      async reconcileFailedGroup() { return ["unknown_within_batch"]; },
    };
    const originalSet = storage.setItem.bind(storage);
    let writes = 0;
    storage.setItem = (key, value) => { writes += 1; if (writes > 1) throw new Error("terminal write failed"); originalSet(key, value); };
    const first = await execute(transaction, adapter, store);
    const retry = await execute(transaction, adapter, store);
    expect(first.outcome).toBe("persistence_unavailable_after_mutation");
    expect(retry.outcome).toBe("persistence_unavailable_after_mutation");
    expect(mutations).toBe(1);
    expect(JSON.stringify(first)).not.toContain("content-canary");
  });

  it("keeps workbook content out of successful persisted records", async () => {
    const storage = new MemoryIdempotencyStorage();
    const transaction = await accepted(planFor([operation("A1", "content-canary")]), ["write"]);
    const report = await execute(transaction, makeAdapter().adapter, new BrowserIdempotencyStore(storage));
    expect(report.outcome).toBe("committed");
    const persisted = storage.getItem("hermes.excel.workbookPlan.idempotency.v1");
    expect(persisted).not.toContain("content-canary");
    expect(persisted).not.toContain("values");
  });

  it("rejects malformed reconciliation entries without leaking adapter data", async () => {
    const transaction = await accepted(planFor([operation("A1", "a"), formatOperation("B1")]), ["write", "format"]);
    const base = makeAdapter({ failGroup: 1 });
    const adapter: WorkbookPlanExecutionAdapter = {
      ...base.adapter,
      reconcileFailedGroup: async () => ["adapter-secret", "committed"] as never,
    };
    const report = await execute(transaction, adapter);
    expect(report.operations[1].state).toBe("unknown_within_batch");
    expect(report.operations[1].status).toBe("unknown_within_batch");
    expect(report.verification).toBe("reconciliation_failed");
    expect(JSON.stringify(report)).not.toContain("adapter-secret");
  });

  it("does not accept stale or unknown runtime results from a later group", async () => {
    const transaction = await accepted(planFor([operation("A1", "a"), formatOperation("B1")]), ["write", "format"]);
    const base = makeAdapter();
    const adapter: WorkbookPlanExecutionAdapter = {
      ...base.adapter,
      executeGroup: async () => ({ status: "stale_precondition" } as never),
      reconcileFailedGroup: async (group) => group.operations.map(() => "not_committed"),
    };
    const report = await execute(transaction, adapter);
    expect(report.operations[0].state).toBe("committed");
    expect(report.operations[0].status).toBe("applied");
    expect(report.operations[1].state).toBe("not_committed");
    expect(report.operations[1].status).toBe("failed");
    expect(report.outcome).toBe("partial");
  });

  it("requires the first-group adapter to call the precondition handshake", async () => {
    const transaction = await accepted(planFor([operation("A1", "a"), formatOperation("B1")]), ["write", "format"]);
    const base = makeAdapter();
    const adapter: WorkbookPlanExecutionAdapter = {
      ...base.adapter,
      executeFirstGroup: async () => ({ status: "committed" }),
      reconcileFailedGroup: async (group) => { base.calls.push("reconcile"); return group.operations.map(() => "not_committed"); },
    };
    const report = await execute(transaction, adapter);
    expect(base.calls).toEqual(["reconcile"]);
    expect(report.operations[0].status).toBe("failed");
    expect(report.operations[1].status).toBe("not_started");
    expect(report.outcome).toBe("partial");
  });

  it("retains frozen postconditions only behind the accepted token and excludes them from reports", async () => {
    const expectedValues = [["postcondition-canary"]];
    const executePlan = planFor([operation("A1", "a")]);
    if (executePlan.outcome !== "execute") throw new Error("expected execute plan");
    const transaction = await accepted({
      ...executePlan,
      verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues }],
    });
    const rules = getAcceptedWorkbookPlanPostconditionsForVerification(transaction);
    expect(rules).toBeDefined();
    expect(Object.isFrozen(rules)).toBe(true);
    expect(Object.isFrozen(rules?.[0])).toBe(true);
    expect(() => { (rules as unknown as Array<unknown>)[0] = {}; }).toThrow();
    expect(getAcceptedWorkbookPlanPostconditionsForVerification(Object.freeze({}) as never)).toBeUndefined();
    const report = await execute(transaction, makeAdapter().adapter);
    expect(JSON.stringify(report)).not.toContain("postcondition-canary");
  });
});

describe("workbook plan executor quality/security repairs", () => {
  const boundary = (turnId = UUIDS[2]) => createTrustedOperationBoundary({
    sessionId: UUIDS[0], documentId: UUIDS[1], turnId,
  });

  it("serializes two owners through one coordinator so exactly one can mutate", async () => {
    const storage = new MemoryIdempotencyStorage();
    const coordinator = new DeterministicIdempotencyCoordinator();
    const storeA = new BrowserIdempotencyStore(storage, { coordinator });
    const storeB = new BrowserIdempotencyStore(storage, { coordinator });
    const trustedBoundary = boundary();
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"], {
      trustedOperationBoundary: trustedBoundary,
    });
    let mutations = 0;
    const adapter: WorkbookPlanExecutionAdapter = {
      async executeFirstGroup(context) {
        return context.executeAfterFresh("fresh", async () => {
          mutations += 1;
          return { status: "committed" };
        });
      },
      async executeGroup() { return { status: "committed" }; },
      async reconcileFailedGroup() { return ["unknown_within_batch"]; },
    };
    const reports = await Promise.all([
      executeAcceptedWorkbookPlan(transaction, trustedBoundary, adapter, storeA),
      executeAcceptedWorkbookPlan(transaction, trustedBoundary, adapter, storeB),
    ]);
    expect(mutations).toBe(1);
    expect(reports.filter((report) => report.outcome === "committed")).toHaveLength(1);
    expect(reports.some((report) => report.outcome === "restart_rejected")).toBe(true);
  });

  it("fails closed on raw oversized data before JSON.parse", async () => {
    const oversized = new RawStorage("[\"" + "x".repeat(500) + "\"]");
    const store = new BrowserIdempotencyStore(oversized, { maxBytes: 64 });
    expect((await store.claim("11111111-1111-4111-8111-111111111111.22222222-2222-4222-8222-222222222222.33333333-3333-4333-8333-333333333333", UUIDS[3], "digest")).kind)
      .toBe("persistence_unavailable");
  });

  it("rejects a terminal report whose transaction identity is tampered", async () => {
    const storage = new RawStorage(JSON.stringify([{
      boundary: "11111111-1111-4111-8111-111111111111.22222222-2222-4222-8222-222222222222.33333333-3333-4333-8333-333333333333",
      transactionId: UUIDS[3], digest: "digest", owner: UUIDS[4], fence: UUIDS[5], state: "terminal", createdAt: 0,
      report: { outcome: "committed", transactionId: UUIDS[6], operationCount: 0, operations: [], verification: "not_attempted", rollback: "not_attempted" },
    }]));
    expect((await new BrowserIdempotencyStore(storage).claim(
      "11111111-1111-4111-8111-111111111111.22222222-2222-4222-8222-222222222222.33333333-3333-4333-8333-333333333333", UUIDS[3], "digest")).kind)
      .toBe("persistence_unavailable");
  });

  it("normalizes throwing proxies and malformed reconciliation without rejecting", async () => {
    const trustedBoundary = boundary();
    const transaction = await accepted(planFor([operation("A1", "a"), formatOperation("B1")]), ["write", "format"], {
      trustedOperationBoundary: trustedBoundary,
    });
    const badResult = new Proxy({}, { get() { throw new Error("status getter"); } });
    const badReconciliation = new Proxy([], { get() { throw new Error("length getter"); } });
    const adapter: WorkbookPlanExecutionAdapter = {
      async executeFirstGroup() { return badResult as never; },
      async executeGroup() { return { status: "sync_failed", failureCode: "sync_failed" }; },
      async reconcileFailedGroup() { return badReconciliation as never; },
    };
    const report = await executeAcceptedWorkbookPlan(transaction, trustedBoundary, adapter, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()));
    expect(report.outcome).toBe("partial");
    expect(report.verification).toBe("reconciliation_failed");
    expect(report.operations[0].state).toBe("unknown_within_batch");
  });

  it("does not hand the first group to the adapter until one exact fresh gate succeeds", async () => {
    const trustedBoundary = boundary();
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"], {
      trustedOperationBoundary: trustedBoundary,
      preconditionCheck: async () => "fresh",
    });
    let mutationCallbacks = 0;
    const report = await executeAcceptedWorkbookPlan(transaction, trustedBoundary, {
      async executeFirstGroup(context) {
        const first = await context.executeAfterFresh("malformed", async () => {
          mutationCallbacks += 1;
          return { status: "committed" };
        });
        expect(first.status).toBe("sync_failed");
        const second = await context.executeAfterFresh("fresh", async () => {
          mutationCallbacks += 1;
          return { status: "committed" };
        });
        expect(second.status).toBe("sync_failed");
        return first;
      },
      async executeGroup() { return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch"); },
    }, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()));
    expect(mutationCallbacks).toBe(0);
    expect(report.verification).toBe("reconciled");
  });

  it("cancels before first-group handoff and never starts later groups after an abort", async () => {
    const trustedBoundary = boundary(UUIDS[7]);
    const transaction = await accepted(planFor([operation("A1", "a"), formatOperation("B1")]), ["write", "format"], {
      trustedOperationBoundary: trustedBoundary,
    });
    const controller = new AbortController();
    controller.abort();
    let calls = 0;
    const report = await executeAcceptedWorkbookPlan(transaction, trustedBoundary, {
      async executeFirstGroup() { calls += 1; return { status: "committed" }; },
      async executeGroup() { calls += 1; return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch"); },
    }, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()), { signal: controller.signal });
    expect(report.outcome).toBe("cancelled");
    expect(calls).toBe(0);
    expect(report.operations.every((item) => item.state === "not_committed")).toBe(true);
  });

  it("waits for a handed-off mutation, reports its evidence, and stops before later groups", async () => {
    const trustedBoundary = createTrustedOperationBoundary({ sessionId: UUIDS[0], documentId: UUIDS[1], turnId: UUIDS[5] });
    const transaction = await accepted(planFor([operation("A1", "a"), formatOperation("B1")]), ["write", "format"], {
      trustedOperationBoundary: trustedBoundary,
    });
    const controller = new AbortController();
    let laterCalls = 0;
    const report = await executeAcceptedWorkbookPlan(transaction, trustedBoundary, {
      async executeFirstGroup(context) {
        return context.executeAfterFresh("fresh", async () => {
          controller.abort();
          return { status: "committed" };
        });
      },
      async executeGroup() { laterCalls += 1; return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch"); },
    }, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()), { signal: controller.signal });
    expect(report.operations[0].state).toBe("committed");
    expect(report.operations[1].state).toBe("not_committed");
    expect(laterCalls).toBe(0);
  });

  it("uses the exact gated result instead of an outer committed override", async () => {
    const trustedBoundary = boundary();
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"], {
      trustedOperationBoundary: trustedBoundary,
    });
    let mutations = 0;
    const report = await executeAcceptedWorkbookPlan(transaction, trustedBoundary, {
      async executeFirstGroup(context) {
        void context.executeAfterFresh("fresh", async () => {
          mutations += 1;
          return { status: "sync_failed", failureCode: "sync_failed" };
        });
        return { status: "committed" };
      },
      async executeGroup() { return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch"); },
    }, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()));
    expect(mutations).toBe(1);
    expect(report.outcome).toBe("partial");
    expect(report.operations[0].state).toBe("unknown_within_batch");
  });

  it("awaits a fire-and-forget gated mutation and ignores the outer failure", async () => {
    const trustedBoundary = boundary(UUIDS[4]);
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"], {
      trustedOperationBoundary: trustedBoundary,
    });
    let mutations = 0;
    const report = await executeAcceptedWorkbookPlan(transaction, trustedBoundary, {
      async executeFirstGroup(context) {
        void context.executeAfterFresh("fresh", async () => {
          mutations += 1;
          return { status: "committed" };
        });
        return { status: "sync_failed", failureCode: "sync_failed" };
      },
      async executeGroup() { return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch"); },
    }, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()));
    expect(mutations).toBe(1);
    expect(report.outcome).toBe("committed");
    expect(report.operations[0].state).toBe("committed");
  });

  it("revokes a delayed gate as soon as the adapter returns", async () => {
    const trustedBoundary = boundary(UUIDS[5]);
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"], {
      trustedOperationBoundary: trustedBoundary,
    });
    let mutations = 0;
    let delayedResult: string | undefined;
    const report = await executeAcceptedWorkbookPlan(transaction, trustedBoundary, {
      async executeFirstGroup(context) {
        setTimeout(() => {
          void context.executeAfterFresh("fresh", async () => {
            mutations += 1;
            return { status: "committed" };
          }).then((result) => { delayedResult = result.status; });
        }, 0);
        return { status: "sync_failed", failureCode: "sync_failed" };
      },
      async executeGroup() { return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch"); },
    }, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()));
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(delayedResult).toBe("sync_failed");
    expect(mutations).toBe(0);
    expect(report.outcome).toBe("partial");
  });

  it("treats a second gate call as misuse even after the first mutation committed", async () => {
    const trustedBoundary = boundary(UUIDS[6]);
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"], {
      trustedOperationBoundary: trustedBoundary,
    });
    let mutations = 0;
    const report = await executeAcceptedWorkbookPlan(transaction, trustedBoundary, {
      async executeFirstGroup(context) {
        const first = await context.executeAfterFresh("fresh", async () => {
          mutations += 1;
          return { status: "committed" };
        });
        const second = await context.executeAfterFresh("fresh", async () => {
          mutations += 1;
          return { status: "committed" };
        });
        expect(first.status).toBe("committed");
        expect(second.status).toBe("sync_failed");
        return first;
      },
      async executeGroup() { return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch"); },
    }, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()));
    expect(mutations).toBe(1);
    expect(report.outcome).toBe("partial");
    expect(report.operations[0].state).toBe("unknown_within_batch");
  });

  it("keeps the exact gated commit when the adapter throws after handing it off", async () => {
    const trustedBoundary = boundary(UUIDS[7]);
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"], {
      trustedOperationBoundary: trustedBoundary,
    });
    let mutations = 0;
    const report = await executeAcceptedWorkbookPlan(transaction, trustedBoundary, {
      async executeFirstGroup(context) {
        void context.executeAfterFresh("fresh", async () => {
          mutations += 1;
          return { status: "committed" };
        });
        throw new Error("outer adapter failure");
      },
      async executeGroup() { return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch"); },
    }, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()));
    expect(mutations).toBe(1);
    expect(report.outcome).toBe("committed");
  });

  it("requires an exact finish fence and never overwrites in-flight state on rejection", async () => {
    const storage = new MemoryIdempotencyStorage();
    const store = new BrowserIdempotencyStore(storage);
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"]);
    const claim = await store.claim(TOKEN_A, transaction.transactionId, transaction.semanticDigest);
    expect(claim.kind).toBe("claimed");
    if (claim.kind !== "claimed") throw new Error("expected claim");
    const report = await execute(transaction, makeAdapter().adapter);
    expect(await store.finish(TOKEN_A, transaction.transactionId, transaction.semanticDigest, report, undefined as never)).toBe(false);
    expect(await store.finish(TOKEN_A, transaction.transactionId, transaction.semanticDigest, report, "wrong-fence")).toBe(false);
    expect(storage.getItem("hermes.excel.workbookPlan.idempotency.v1")).toContain('"state":"in_flight"');
    expect(await store.finish(TOKEN_A, transaction.transactionId, transaction.semanticDigest, report, claim.fence)).toBe(true);
    expect(storage.getItem("hermes.excel.workbookPlan.idempotency.v1")).toContain('"state":"terminal"');
  });

  it("normalizes dense reconciliation through descriptors without invoking getters", async () => {
    const trustedBoundary = boundary();
    const transaction = await accepted(planFor([operation("A1", "a"), formatOperation("B1"), formatOperation("C1")]), ["write", "format", "format"], {
      trustedOperationBoundary: trustedBoundary,
    });
    let getterCalls = 0;
    const dense = ["committed", "not_committed"] as const;
    const reconciliation = new Proxy(dense, {
      get(target, property, receiver) {
        if (property === "length" || property === "0" || property === "1") {
          getterCalls += 1;
          throw new Error("reconciliation getter invoked");
        }
        return Reflect.get(target, property, receiver);
      },
    });
    const base = makeAdapter({ failGroup: 1 });
    const report = await executeAcceptedWorkbookPlan(transaction, trustedBoundary, {
      ...base.adapter,
      reconcileFailedGroup: async () => reconciliation as never,
    }, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()));
    expect(getterCalls).toBe(0);
    expect(report.verification).toBe("reconciled");
    expect(report.operations[2].state).toBe("not_committed");
  });
});

describe("workbook plan executor R2c trusted beforeFirstMutation seam", () => {
  const allNotStarted = (report: WorkbookPlanExecutionReport): boolean =>
    report.operations.every((operation) => operation.state === "not_committed" && operation.status === "not_started");

  const mutationAdapter = (onMutation: () => void): WorkbookPlanExecutionAdapter => ({
    async executeFirstGroup(context) {
      return context.executeAfterFresh("fresh", async () => {
        onMutation();
        return { status: "committed" };
      });
    },
    async executeGroup() {
      onMutation();
      return { status: "committed" };
    },
    async reconcileFailedGroup(group) {
      return group.operations.map(() => "unknown_within_batch" as const);
    },
  });

  it("orders durable claim, fresh/trusted checks, callback, and mutation, invoking the callback once", async () => {
    const events: string[] = [];
    const storage = new EventStorage(events);
    const store = new BrowserIdempotencyStore(storage);
    const transaction = await accepted(
      planFor([operation("A1", "a"), formatOperation("B1"), operation("C1", "c")]),
      ["write", "format", "write"],
      {
        preconditionCheck: async () => {
          events.push("trusted-precondition");
          return "fresh";
        },
      },
    );
    let callbackCalls = 0;
    const report = await execute(transaction, {
      async executeFirstGroup(context) {
        events.push("fresh-observation");
        return context.executeAfterFresh("fresh", async (group) => {
          events.push(`mutation:${group.groupIndex}`);
          return { status: "committed" };
        });
      },
      async executeGroup(group) {
        events.push(`mutation:${group.groupIndex}`);
        return { status: "committed" };
      },
      async reconcileFailedGroup(group) {
        return group.operations.map(() => "unknown_within_batch" as const);
      },
    }, store, BOUNDARY_A, {
      beforeFirstMutation: async () => {
        callbackCalls += 1;
        events.push("callback");
      },
    });

    expect(report.outcome).toBe("committed");
    expect(callbackCalls).toBe(1);
    expect(events).toEqual([
      "durable-claim",
      "fresh-observation",
      "trusted-precondition",
      "callback",
      "mutation:0",
      "mutation:1",
      "mutation:2",
      "durable-finish",
    ]);
  });

  it("fails closed on a throwing callback, hides its error, and terminally fences the claim", async () => {
    const canary = "private-before-first-mutation-canary";
    const storage = new MemoryIdempotencyStorage();
    const store = new BrowserIdempotencyStore(storage);
    const transaction = await accepted(planFor([operation("A1", "content")]), ["write"]);
    let callbackCalls = 0;
    let mutations = 0;
    let reconciliations = 0;
    const adapter = {
      ...mutationAdapter(() => { mutations += 1; }),
      reconcileFailedGroup: async (group: { readonly operations: readonly unknown[] }) => {
        reconciliations += 1;
        return group.operations.map(() => "unknown_within_batch" as const);
      },
    } as WorkbookPlanExecutionAdapter;
    const report = await execute(transaction, adapter, store, BOUNDARY_A, {
      beforeFirstMutation: async () => {
        callbackCalls += 1;
        throw new Error(canary);
      },
    });

    expect(report.outcome).toBe("cancelled");
    expect(callbackCalls).toBe(1);
    expect(mutations).toBe(0);
    expect(reconciliations).toBe(0);
    expect(allNotStarted(report)).toBe(true);
    expect(JSON.stringify(report)).not.toContain(canary);
    expect(storage.getItem("hermes.excel.workbookPlan.idempotency.v1")).not.toContain(canary);
    expect(storage.getItem("hermes.excel.workbookPlan.idempotency.v1")).toContain('"state":"terminal"');

    let replayCallbackCalls = 0;
    let replayMutations = 0;
    const replay = await execute(
      transaction,
      mutationAdapter(() => { replayMutations += 1; }),
      store,
      BOUNDARY_A,
      { beforeFirstMutation: async () => { replayCallbackCalls += 1; } },
    );
    expect(replay.outcome).toBe("cancelled");
    expect(replayCallbackCalls).toBe(0);
    expect(replayMutations).toBe(0);
  });

  it("aborts a deferred callback before handoff with zero mutation and non-committed accounting", async () => {
    const controller = new AbortController();
    const callbackGate = deferred<void>();
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"]);
    let callbackCalls = 0;
    let mutations = 0;
    const execution = execute(
      transaction,
      mutationAdapter(() => { mutations += 1; }),
      new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      BOUNDARY_A,
      {
        signal: controller.signal,
        beforeFirstMutation: async () => {
          callbackCalls += 1;
          return callbackGate.promise;
        },
      },
    );
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    controller.abort();
    callbackGate.resolve();
    const report = await execution;

    expect(report.outcome).toBe("cancelled");
    expect(callbackCalls).toBe(1);
    expect(mutations).toBe(0);
    expect(allNotStarted(report)).toBe(true);
  });

  it("revokes a first-group gate while its callback is pending and never mutates after resolution", async () => {
    const callbackGate = deferred<void>();
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"]);
    let callbackCalls = 0;
    let mutations = 0;
    let reconciliations = 0;
    let delayedResult: string | undefined;
    const execution = execute(
      transaction,
      {
        async executeFirstGroup(context) {
          void context.executeAfterFresh("fresh", async () => {
            mutations += 1;
            return { status: "committed" };
          }).then((result) => { delayedResult = result.status; });
          return { status: "sync_failed", failureCode: "sync_failed" };
        },
        async executeGroup() {
          throw new Error("later group must not start");
        },
        async reconcileFailedGroup(group) {
          reconciliations += 1;
          return group.operations.map(() => "unknown_within_batch" as const);
        },
      },
      new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      BOUNDARY_A,
      {
        beforeFirstMutation: async () => {
          callbackCalls += 1;
          return callbackGate.promise;
        },
      },
    );
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    callbackGate.resolve();
    const report = await execution;
    await Promise.resolve();

    expect(report.outcome).toBe("cancelled");
    expect(callbackCalls).toBe(1);
    expect(mutations).toBe(0);
    expect(reconciliations).toBe(0);
    expect(delayedResult).toBe("sync_failed");
    expect(allNotStarted(report)).toBe(true);
  });

  it("does not invoke the callback or mutate for stale observation or stale trusted preconditions", async () => {
    for (const mode of ["observation", "trusted"] as const) {
      let callbackCalls = 0;
      let mutations = 0;
      const transaction = await accepted(planFor([operation("A1", mode)]), ["write"], mode === "trusted"
        ? { preconditionCheck: async () => "stale" as const }
        : {});
      const report = await execute(transaction, {
        async executeFirstGroup(context) {
          return context.executeAfterFresh(mode === "observation" ? "stale" : "fresh", async () => {
            mutations += 1;
            return { status: "committed" };
          });
        },
        async executeGroup() { mutations += 1; return { status: "committed" }; },
        async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch" as const); },
      }, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()), BOUNDARY_A, {
        beforeFirstMutation: async () => { callbackCalls += 1; },
      });

      expect(report.outcome).toBe("stale");
      expect(callbackCalls).toBe(0);
      expect(mutations).toBe(0);
      expect(allNotStarted(report)).toBe(true);
    }
  });

  it("does not invoke the callback or mutate for a pre-cancelled signal", async () => {
    const controller = new AbortController();
    controller.abort();
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"]);
    let callbackCalls = 0;
    let adapterCalls = 0;
    const report = await execute(transaction, {
      async executeFirstGroup() { adapterCalls += 1; return { status: "committed" }; },
      async executeGroup() { adapterCalls += 1; return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch" as const); },
    }, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()), BOUNDARY_A, {
      signal: controller.signal,
      beforeFirstMutation: async () => { callbackCalls += 1; },
    });

    expect(report.outcome).toBe("cancelled");
    expect(adapterCalls).toBe(0);
    expect(callbackCalls).toBe(0);
    expect(allNotStarted(report)).toBe(true);
  });

  it("does not invoke the callback or mutate on terminal replay, conflict, in-flight, or persistence failure", async () => {
    const storage = new MemoryIdempotencyStorage();
    const store = new BrowserIdempotencyStore(storage);
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"]);
    let firstMutations = 0;
    expect((await execute(transaction, mutationAdapter(() => { firstMutations += 1; }), store)).outcome).toBe("committed");
    expect(firstMutations).toBe(1);

    let replayCallbackCalls = 0;
    let replayMutations = 0;
    const replay = await execute(transaction, mutationAdapter(() => { replayMutations += 1; }), store, BOUNDARY_A, {
      beforeFirstMutation: async () => { replayCallbackCalls += 1; },
    });
    expect(replay.outcome).toBe("committed");
    expect(replayCallbackCalls).toBe(0);
    expect(replayMutations).toBe(0);

    const conflictTransaction = await accepted(planFor([operation("A1", "changed")]), ["write"]);
    let conflictCallbackCalls = 0;
    let conflictMutations = 0;
    const conflict = await execute(conflictTransaction, mutationAdapter(() => { conflictMutations += 1; }), store, BOUNDARY_A, {
      beforeFirstMutation: async () => { conflictCallbackCalls += 1; },
    });
    expect(conflict.outcome).toBe("conflict");
    expect(conflictCallbackCalls).toBe(0);
    expect(conflictMutations).toBe(0);

    const inFlightStorage = new MemoryIdempotencyStorage();
    const inFlightTransaction = await accepted(planFor([operation("B1", "b")]), ["write"], {
      trustedOperationBoundary: BOUNDARY_B,
    });
    const holder = new BrowserIdempotencyStore(inFlightStorage);
    const claim = await holder.claim(TOKEN_B, inFlightTransaction.transactionId, inFlightTransaction.semanticDigest);
    expect(claim.kind).toBe("claimed");
    let inFlightCallbackCalls = 0;
    let inFlightMutations = 0;
    const inFlight = await execute(inFlightTransaction, mutationAdapter(() => { inFlightMutations += 1; }), new BrowserIdempotencyStore(inFlightStorage), BOUNDARY_B, {
      beforeFirstMutation: async () => { inFlightCallbackCalls += 1; },
    });
    expect(inFlight.outcome).toBe("restart_rejected");
    expect(inFlightCallbackCalls).toBe(0);
    expect(inFlightMutations).toBe(0);

    const failingTransaction = await accepted(planFor([operation("C1", "c")]), ["write"], {
      trustedOperationBoundary: BOUNDARY_C,
    });
    let failureCallbackCalls = 0;
    let failureMutations = 0;
    const failure = await execute(failingTransaction, mutationAdapter(() => { failureMutations += 1; }), new BrowserIdempotencyStore(new RawStorage(null, true)), BOUNDARY_C, {
      beforeFirstMutation: async () => { failureCallbackCalls += 1; },
    });
    expect(failure.outcome).toBe("persistence_unavailable");
    expect(failureCallbackCalls).toBe(0);
    expect(failureMutations).toBe(0);
  });

  it("fails closed for wrong or missing boundaries, malformed callbacks, and foreign accepted transactions", async () => {
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"]);
    let wrongCallbackCalls = 0;
    let wrongMutations = 0;
    const wrongBoundary = await execute(transaction, mutationAdapter(() => { wrongMutations += 1; }), new BrowserIdempotencyStore(new MemoryIdempotencyStorage()), BOUNDARY_B, {
      beforeFirstMutation: async () => { wrongCallbackCalls += 1; },
    });
    expect(wrongBoundary.outcome).toBe("rejected");
    expect(wrongCallbackCalls).toBe(0);
    expect(wrongMutations).toBe(0);

    let missingCallbackCalls = 0;
    let missingMutations = 0;
    const missingBoundary = await executeAcceptedWorkbookPlan(
      transaction,
      undefined as never,
      mutationAdapter(() => { missingMutations += 1; }),
      new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      { beforeFirstMutation: async () => { missingCallbackCalls += 1; } } as never,
    );
    expect(missingBoundary.outcome).toBe("rejected");
    expect(missingCallbackCalls).toBe(0);
    expect(missingMutations).toBe(0);

    let malformedMutations = 0;
    const malformed = await execute(transaction, mutationAdapter(() => { malformedMutations += 1; }), new BrowserIdempotencyStore(new MemoryIdempotencyStorage()), BOUNDARY_A, {
      beforeFirstMutation: "not-a-callback",
    });
    expect(malformed.outcome).toBe("cancelled");
    expect(malformedMutations).toBe(0);
    expect(allNotStarted(malformed)).toBe(true);

    const foreign = Object.freeze({
      transactionId: transaction.transactionId,
      semanticDigest: transaction.semanticDigest,
      operationCount: transaction.operationCount,
    });
    let foreignCallbackCalls = 0;
    let foreignMutations = 0;
    const foreignReport = await execute(foreign as never, mutationAdapter(() => { foreignMutations += 1; }), new BrowserIdempotencyStore(new MemoryIdempotencyStorage()), BOUNDARY_A, {
      beforeFirstMutation: async () => { foreignCallbackCalls += 1; },
    });
    expect(foreignReport.outcome).toBe("rejected");
    expect(foreignCallbackCalls).toBe(0);
    expect(foreignMutations).toBe(0);
  });

  it("preserves no-callback execution and the direct AbortSignal overload", async () => {
    const controller = new AbortController();
    const transaction = await accepted(planFor([operation("A1", "a")]), ["write"]);
    let mutations = 0;
    const report = await executeAcceptedWorkbookPlan(
      transaction,
      BOUNDARY_A,
      mutationAdapter(() => { mutations += 1; }),
      new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      controller.signal,
    );

    expect(report.outcome).toBe("committed");
    expect(mutations).toBe(1);
  });
});
