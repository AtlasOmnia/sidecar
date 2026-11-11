import { afterEach, describe, expect, it, vi } from "vitest";
import {
  BrowserIdempotencyStore,
  MemoryIdempotencyStorage,
  createTrustedOperationBoundary,
  executeAcceptedWorkbookPlan,
  prepareWorkbookPlanExecution,
  type IdempotencyStorage,
  type TrustedOperationBoundary,
  type WorkbookPlanExecutionAdapter,
  type WorkbookPlanExecutionOptions,
  type WorkbookPlanExecutionReport,
  type WorkbookPlanGroupExecutionResult,
} from "../workbookPlanExecutor";
import { getExcelToolDefinition } from "../excelTools";
import type { WorkbookOperation, WorkbookPlan } from "../generated/workbookPlanSchema";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import { WORKBOOK_PLAN_CAPABILITY_NAMES, type WorkbookPlanPolicyInput } from "../workbookPlanPolicy";

/**
 * R4e executor/store hardening (dev3-r4e-release §5.1/§5.2, file 4).
 * - E1-E6: R2c A2 gate label on a callback throw; R2d-1 GA3 optional deadlineAt rechecked after the
 *   beforeFirstMutation acknowledgement and immediately before the first mutation.
 * - S1-S6: R3c A1 clock guards in claim/finish; R3c B-A2 clone-first finish.
 * - R1-R3: replay marking (isReplayedWorkbookPlanReport) without changing the cached report bytes.
 * Helpers are copied from workbook-plan-executor.test.ts (nothing is imported from other test files).
 * deadlineAt is passed through `as unknown as WorkbookPlanExecutionOptions` and the predicate is
 * reached by a dynamic import, so this file type-checks at the base and RED is assertion-only.
 */

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
const STORAGE_KEY = "hermes.excel.workbookPlan.idempotency.v1";
const T = Date.UTC(2026, 9, 7, 12, 0, 0);
const DEADLINE = T + 90_000;
const SYNC_FAILED = { status: "sync_failed", failureCode: "sync_failed" };

const BOUNDARY_A = createTrustedOperationBoundary({ sessionId: UUIDS[0], documentId: UUIDS[1], turnId: UUIDS[2] });
const BOUNDARY_B = createTrustedOperationBoundary({ sessionId: UUIDS[0], documentId: UUIDS[1], turnId: UUIDS[7] });
const BOUNDARY_C = createTrustedOperationBoundary({ sessionId: UUIDS[0], documentId: UUIDS[1], turnId: UUIDS[6] });
const token = (turnIndex: number): string => `${UUIDS[0]}.${UUIDS[1]}.${UUIDS[turnIndex]}`;

const uuidSource = (): (() => string) => {
  let index = 0;
  return () => UUIDS[index++] ?? "99999999-9999-4999-8999-999999999999";
};

const operation = (address: string, value: string): WorkbookOperation => ({
  kind: "write_values",
  sheet: "Sheet1",
  address,
  values: [[value]],
});

const planFor = (operations: readonly WorkbookOperation[]): WorkbookPlan => ({
  version: 1,
  outcome: "execute",
  summary: "test plan",
  operations,
  verification: [{ kind: "calculation_completed" }],
});

const primitiveNames = ["write_range", "format_range", "run_calculation"];

const compileContext = (): VerifiedWorkbookCompileContext => ({
  sheetNames: new Set(["Sheet1"]),
  tablesByName: new Map(),
  supportedPrimitiveNames: new Set(primitiveNames.filter((name) => getExcelToolDefinition(name) !== undefined)),
  supportedCapabilities: new Set(["excel_workbook_access", "workbook_plan_v1_operation_audit"]),
  chartSourceResolutionsBySheet: new Map(),
});

const policyInputFor = (operationCount: number): WorkbookPlanPolicyInput => ({
  requestClassification: "explicit_mutation",
  activeSheetName: "Sheet1",
  selectionScopedRequest: false,
  operationSummaries: Array.from({ length: operationCount }, (_, operationIndex) => ({
    operationIndex,
    effect: "write" as const,
    target: { operationIndex, sheetNames: ["Sheet1"], selectionRelationship: "inside" as const },
  })),
  capabilitySummaries: [...WORKBOOK_PLAN_CAPABILITY_NAMES],
  resourceUsage: {
    operationCount, cellsRead: 0, cellsWritten: operationCount, formulaUtf8Bytes: 0, csvPayloadUtf8Bytes: 0,
    commentUtf8Bytes: 0, sheetsCreated: 0, chartsCreated: 0, tablesCreated: 0, verificationReads: 0,
  },
  resourceLimits: {
    operationCount: 100, cellsRead: 100, cellsWritten: 100, formulaUtf8Bytes: 10000, csvPayloadUtf8Bytes: 10000,
    commentUtf8Bytes: 10000, sheetsCreated: 100, chartsCreated: 100, tablesCreated: 100, verificationReads: 100,
  },
});

const accepted = async (
  value: string,
  boundary: TrustedOperationBoundary = BOUNDARY_A,
) => {
  const result = await prepareWorkbookPlanExecution({
    plan: planFor([operation("A1", value)]),
    compileContext: compileContext(),
    policyInput: policyInputFor(1),
    trustedOperationBoundary: boundary,
    uuidSource: uuidSource(),
  });
  expect(result.accepted).toBe(true);
  if (!result.accepted) throw new Error("expected accepted plan");
  return result.transaction;
};

type R4eOptions = {
  readonly signal?: AbortSignal;
  readonly beforeFirstMutation?: unknown;
  readonly deadlineAt?: unknown;
};

const execute = (
  transaction: Awaited<ReturnType<typeof accepted>>,
  adapter: WorkbookPlanExecutionAdapter,
  store: BrowserIdempotencyStore,
  options?: R4eOptions | AbortSignal,
  boundary: TrustedOperationBoundary = BOUNDARY_A,
): Promise<WorkbookPlanExecutionReport> =>
  executeAcceptedWorkbookPlan(transaction, boundary, adapter, store, options as unknown as WorkbookPlanExecutionOptions);

/** Records every gate value the adapter observes and every mutation the gate hands off. */
const gateAdapter = (mutation: () => void = () => undefined): {
  adapter: WorkbookPlanExecutionAdapter;
  gateValues: unknown[];
  mutations: () => number;
  reconciliations: () => number;
} => {
  const gateValues: unknown[] = [];
  let mutations = 0;
  let reconciliations = 0;
  const adapter: WorkbookPlanExecutionAdapter = {
    async executeFirstGroup(context) {
      const value = await context.executeAfterFresh("fresh", async (): Promise<WorkbookPlanGroupExecutionResult> => {
        mutations += 1;
        mutation();
        return { status: "committed" };
      });
      gateValues.push(value);
      return value;
    },
    async executeGroup() {
      mutations += 1;
      return { status: "committed" };
    },
    async reconcileFailedGroup(group) {
      reconciliations += 1;
      return group.operations.map(() => "unknown_within_batch" as const);
    },
  };
  return { adapter, gateValues, mutations: () => mutations, reconciliations: () => reconciliations };
};

class CountingStorage implements IdempotencyStorage {
  value: string | null = null;
  reads = 0;
  writes = 0;
  failGet = false;
  failTerminalSet = false;

  getItem(): string | null {
    this.reads += 1;
    if (this.failGet) throw new Error("read failed");
    return this.value;
  }

  setItem(_key: string, value: string): void {
    if (this.failTerminalSet && (JSON.parse(value) as { state?: string }[]).some((record) => record.state === "terminal")) {
      throw new Error("write failed");
    }
    this.writes += 1;
    this.value = value;
  }

  removeItem(): void {
    this.value = null;
  }

  records(): Record<string, unknown>[] {
    return this.value === null ? [] : JSON.parse(this.value) as Record<string, unknown>[];
  }
}

const allNotStarted = (report: WorkbookPlanExecutionReport): boolean =>
  report.operations.length > 0 &&
  report.operations.every((item) => item.state === "not_committed" && item.status === "not_started");

const committedReport = (transactionId: string): WorkbookPlanExecutionReport => ({
  outcome: "committed",
  transactionId,
  operationCount: 0,
  operations: [],
  verification: "not_attempted",
  rollback: "not_attempted",
});

type ReplayPredicate = (value: unknown) => boolean;

async function replayPredicate(): Promise<unknown> {
  return ((await import("../workbookPlanExecutor")) as Record<string, unknown>).isReplayedWorkbookPlanReport;
}

afterEach(() => {
  vi.useRealTimers();
});

describe("R4e E rows: first-group gate label and post-acknowledgement deadline", () => {
  it("E1 a throwing beforeFirstMutation resolves the gate to sync_failed; the report stays cancelled with nothing reconciled or leaked", async () => {
    const canary = "R4E_CALLBACK_CANARY_5D1";
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage);
    const transaction = await accepted("e1");
    const probe = gateAdapter();
    let callbackCalls = 0;
    const report = await execute(transaction, probe.adapter, store, {
      beforeFirstMutation: async () => { callbackCalls += 1; throw new Error(canary); },
    });
    expect(probe.gateValues).toEqual([SYNC_FAILED]);
    expect(report.outcome).toBe("cancelled");
    expect(allNotStarted(report)).toBe(true);
    expect(callbackCalls).toBe(1);
    expect(probe.mutations()).toBe(0);
    expect(probe.reconciliations()).toBe(0);
    expect(JSON.stringify(report)).not.toContain(canary);
    expect(storage.value).not.toContain(canary);
    expect(storage.records()).toMatchObject([{ state: "terminal", report: { outcome: "cancelled" } }]);
  });

  it("E2 the turn deadline passes while the acknowledgement is in flight: zero mutation, cancelled, stored terminal cancelled, gate sync_failed", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    vi.setSystemTime(T);
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage);
    const transaction = await accepted("e2");
    const probe = gateAdapter();
    const report = await execute(transaction, probe.adapter, store, {
      deadlineAt: DEADLINE,
      beforeFirstMutation: async () => { vi.setSystemTime(DEADLINE); },
    });
    expect(probe.mutations()).toBe(0);
    expect(report.outcome).toBe("cancelled");
    expect(allNotStarted(report)).toBe(true);
    expect(probe.gateValues).toEqual([SYNC_FAILED]);
    expect(storage.records()).toMatchObject([{ state: "terminal", report: { outcome: "cancelled" } }]);
  });

  it("E3 boundary: one millisecond before deadlineAt mutates once and commits; exactly deadlineAt mutates zero times", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    const outcomes: Record<string, { outcome: string; mutations: number }> = {};
    for (const [label, at] of [["before", DEADLINE - 1], ["equal", DEADLINE]] as const) {
      vi.setSystemTime(T);
      const transaction = await accepted(`e3-${label}`);
      const probe = gateAdapter();
      const report = await execute(transaction, probe.adapter, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()), {
        deadlineAt: DEADLINE,
        beforeFirstMutation: async () => { vi.setSystemTime(at); },
      });
      outcomes[label] = { outcome: report.outcome, mutations: probe.mutations() };
    }
    expect(outcomes).toEqual({
      before: { outcome: "committed", mutations: 1 },
      equal: { outcome: "cancelled", mutations: 0 },
    });
  });

  it("E4 a malformed deadlineAt (NaN, Infinity, string, object, throwing getter) is cancelled before the claim: no storage access, no callback, no mutation", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    vi.setSystemTime(T);
    const throwing: Record<string, unknown> = {};
    Object.defineProperty(throwing, "deadlineAt", { enumerable: true, get: () => { throw new Error("deadline getter"); } });
    const variants: [string, () => Record<string, unknown>][] = [
      ["NaN", () => ({ deadlineAt: Number.NaN })],
      ["Infinity", () => ({ deadlineAt: Number.POSITIVE_INFINITY })],
      ["string", () => ({ deadlineAt: "90000" })],
      ["object", () => ({ deadlineAt: {} })],
      ["throwing getter", () => throwing],
    ];
    const observed: Record<string, unknown> = {};
    for (const [label, make] of variants) {
      const storage = new CountingStorage();
      const transaction = await accepted(`e4-${label}`);
      const probe = gateAdapter();
      let callbackCalls = 0;
      const options = make();
      options.beforeFirstMutation = async () => { callbackCalls += 1; };
      const report = await execute(transaction, probe.adapter, new BrowserIdempotencyStore(storage), options);
      observed[label] = {
        outcome: report.outcome, notStarted: allNotStarted(report), reads: storage.reads, writes: storage.writes,
        callbackCalls, mutations: probe.mutations(),
      };
    }
    const expected = { outcome: "cancelled", notStarted: true, reads: 0, writes: 0, callbackCalls: 0, mutations: 0 };
    expect(observed).toEqual(Object.fromEntries(variants.map(([label]) => [label, expected])));
  });

  it("E5 a deadline already past when the gate opens (no callback) refuses the first write: zero mutation, cancelled", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    vi.setSystemTime(DEADLINE + 5);
    const storage = new CountingStorage();
    const transaction = await accepted("e5");
    const probe = gateAdapter();
    const report = await execute(transaction, probe.adapter, new BrowserIdempotencyStore(storage), { deadlineAt: DEADLINE });
    expect(probe.mutations()).toBe(0);
    expect(report.outcome).toBe("cancelled");
    expect(allNotStarted(report)).toBe(true);
    expect(probe.gateValues).toEqual([SYNC_FAILED]);
    expect(storage.records()).toMatchObject([{ state: "terminal", report: { outcome: "cancelled" } }]);
  });

  it("E6 [pin] without deadlineAt nothing changes: the AbortSignal overload and an options object commit even far past any 90 s window", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    vi.setSystemTime(T);
    const bare = gateAdapter(() => { vi.setSystemTime(T + 10 * 90_000); });
    const controller = new AbortController();
    const bareReport = await execute(await accepted("e6-bare"), bare.adapter, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()), controller.signal);
    expect(bareReport.outcome).toBe("committed");
    expect(bare.mutations()).toBe(1);
    expect(bare.gateValues).toEqual([{ status: "committed" }]);

    vi.setSystemTime(T);
    const withCallback = gateAdapter();
    let callbackCalls = 0;
    const optionsReport = await execute(await accepted("e6-options"), withCallback.adapter, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()), {
      signal: controller.signal,
      beforeFirstMutation: async () => { callbackCalls += 1; vi.setSystemTime(T + 10 * 90_000); },
    });
    expect(optionsReport.outcome).toBe("committed");
    expect(callbackCalls).toBe(1);
    expect(withCallback.mutations()).toBe(1);
  });
});

describe("R4e S rows: store clock guards and clone-first finish", () => {
  it("S1 a NaN or +Infinity clock at finish returns persistence_unavailable_after_mutation, writes no createdAt:null, and does not brick the origin", async () => {
    const observed: Record<string, unknown> = {};
    for (const bad of [Number.NaN, Number.POSITIVE_INFINITY]) {
      let now = 0;
      const storage = new CountingStorage();
      const store = new BrowserIdempotencyStore(storage, { clock: () => now });
      const probe = gateAdapter(() => { now = bad; });
      const report = await execute(await accepted(`s1-${String(bad)}`), probe.adapter, store);
      now = 1000;
      const next = await new BrowserIdempotencyStore(storage, { clock: () => now }).claim(token(4), UUIDS[3], "digest-s1");
      observed[String(bad)] = {
        outcome: report.outcome,
        mutations: probe.mutations(),
        createdAtNull: (storage.value ?? "").includes("\"createdAt\":null"),
        states: storage.records().filter((record) => record.boundary === token(2)).map((record) => record.state),
        next: next.kind,
      };
    }
    const expected = { outcome: "persistence_unavailable_after_mutation", mutations: 1, createdAtNull: false, states: ["in_flight"], next: "claimed" };
    expect(observed).toEqual({ NaN: expected, Infinity: expected });
  });

  it("S2 a negative clock at finish behaves as S1: no negative record, the record stays in_flight, a later claim succeeds", async () => {
    let now = 0;
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => now });
    const probe = gateAdapter(() => { now = -5; });
    const report = await execute(await accepted("s2"), probe.adapter, store);
    now = 1000;
    const next = await new BrowserIdempotencyStore(storage, { clock: () => now }).claim(token(4), UUIDS[3], "digest-s2");
    expect({
      outcome: report.outcome,
      mutations: probe.mutations(),
      negative: storage.records().some((record) => (record.createdAt as number) < 0),
      states: storage.records().map((record) => record.state),
      next: next.kind,
    }).toEqual({ outcome: "persistence_unavailable_after_mutation", mutations: 1, negative: false, states: ["in_flight", "in_flight"], next: "claimed" });
  });

  it("S3 a negative clock at claim is persistence_unavailable with zero reads and writes; a sane clock then claims", async () => {
    let now = -5;
    const storage = new CountingStorage();
    const first = await new BrowserIdempotencyStore(storage, { clock: () => now }).claim(token(2), UUIDS[3], "digest-s3");
    const touched = { reads: storage.reads, writes: storage.writes };
    now = 10;
    const second = await new BrowserIdempotencyStore(storage, { clock: () => now }).claim(token(4), UUIDS[4], "digest-s3b");
    expect({ first: first.kind, ...touched, second: second.kind }).toEqual({ first: "persistence_unavailable", reads: 0, writes: 0, second: "claimed" });
  });

  it("S4 getter TOCTOU sweep (k = 1..10): finish never persists a restart_rejected marker; k = 1 refuses, k >= 2 persists committed", async () => {
    const observed: string[] = [];
    const expected: string[] = [];
    for (let k = 1; k <= 10; k += 1) {
      const storage = new CountingStorage();
      const store = new BrowserIdempotencyStore(storage, { clock: () => 0 });
      const boundary = token(2);
      const transactionId = UUIDS[3];
      const claim = await store.claim(boundary, transactionId, "digest-s4");
      if (claim.kind !== "claimed") throw new Error("expected claimed");
      let reads = 0;
      const base = committedReport(transactionId) as unknown as Record<string, unknown>;
      const evil: Record<string, unknown> = {};
      for (const key of Object.keys(base)) if (key !== "outcome") evil[key] = base[key];
      Object.defineProperty(evil, "outcome", {
        enumerable: true,
        get: () => { reads += 1; return reads < k ? "committed" : "restart_rejected"; },
      });
      const ok = await store.finish(boundary, transactionId, "digest-s4", evil as unknown as WorkbookPlanExecutionReport, claim.fence);
      const record = storage.records()[0] as { state?: string; report?: { outcome?: string } };
      const count = await new BrowserIdempotencyStore(storage, { clock: () => 0 }).reconciliationRequiredCount();
      observed.push(`k${k}:${String(ok)}/${record.state}/${record.report?.outcome ?? "-"}/markers${String(count)}`);
      expected.push(k === 1 ? `k${k}:false/in_flight/-/markers0` : `k${k}:true/terminal/committed/markers0`);
    }
    expect(observed).toEqual(expected);
  });

  it("S5 [pin] a plain restart_rejected report is refused by finish and the record stays in_flight (LB6a)", async () => {
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => 0 });
    const claim = await store.claim(token(2), UUIDS[3], "digest-s5");
    if (claim.kind !== "claimed") throw new Error("expected claimed");
    const ok = await store.finish(token(2), UUIDS[3], "digest-s5", { ...committedReport(UUIDS[3]), outcome: "restart_rejected" }, claim.fence);
    expect(ok).toBe(false);
    expect(storage.records().map((record) => record.state)).toEqual(["in_flight"]);
  });

  it("S6 [pin] a throwing clock: claim is persistence_unavailable, finish is false, nothing is written (LB4e)", async () => {
    let throwNow = false;
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => { if (throwNow) throw new Error("clock"); return 0; } });
    const claim = await store.claim(token(2), UUIDS[3], "digest-s6");
    if (claim.kind !== "claimed") throw new Error("expected claimed");
    throwNow = true;
    const writes = storage.writes;
    const other = await store.claim(token(4), UUIDS[4], "digest-s6b");
    const finished = await store.finish(token(2), UUIDS[3], "digest-s6", committedReport(UUIDS[3]), claim.fence);
    expect({ other: other.kind, finished, writes: storage.writes - writes }).toEqual({ other: "persistence_unavailable", finished: false, writes: 0 });
  });
});

describe("R4e R rows: replayed-report marking", () => {
  it("R1 isReplayedWorkbookPlanReport is true only for a terminal-claim replay, whose bytes equal the stored report", async () => {
    const candidate = await replayPredicate();
    expect(typeof candidate, "workbookPlanExecutor must export isReplayedWorkbookPlanReport").toBe("function");
    const isReplayed = candidate as ReplayPredicate;
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage);
    const transaction = await accepted("r1");
    const first = await execute(transaction, gateAdapter().adapter, store);
    expect(first.outcome).toBe("committed");
    const replayProbe = gateAdapter();
    const replay = await execute(transaction, replayProbe.adapter, store);
    expect(replayProbe.mutations()).toBe(0);
    expect(isReplayed(replay)).toBe(true);
    expect(isReplayed(first)).toBe(false);
    expect(JSON.stringify(replay)).toBe(JSON.stringify(storage.records()[0].report));
    expect(JSON.stringify(replay)).toBe(JSON.stringify(first));

    const conflict = await execute(await accepted("r1-changed"), gateAdapter().adapter, store);
    expect(conflict.outcome).toBe("conflict");

    const foreignStorage = new MemoryIdempotencyStorage();
    const holderTransaction = await accepted("r1-held", BOUNDARY_B);
    expect((await new BrowserIdempotencyStore(foreignStorage).claim(token(7), holderTransaction.transactionId, holderTransaction.semanticDigest)).kind).toBe("claimed");
    const restart = await execute(holderTransaction, gateAdapter().adapter, new BrowserIdempotencyStore(foreignStorage), undefined, BOUNDARY_B);
    expect(restart.outcome).toBe("restart_rejected");

    const unavailableStorage = new CountingStorage();
    unavailableStorage.failGet = true;
    const unavailable = await execute(await accepted("r1-pu", BOUNDARY_C), gateAdapter().adapter, new BrowserIdempotencyStore(unavailableStorage), undefined, BOUNDARY_C);
    expect(unavailable.outcome).toBe("persistence_unavailable");

    const malformed = await execute(await accepted("r1-malformed"), gateAdapter().adapter, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()), {
      beforeFirstMutation: "not-a-callback",
    });
    expect(malformed.outcome).toBe("cancelled");

    const rejected = await execute(transaction, gateAdapter().adapter, store, undefined, BOUNDARY_B);
    expect(rejected.outcome).toBe("rejected");

    for (const value of [conflict, restart, unavailable, malformed, rejected, JSON.parse(JSON.stringify(replay)), { ...replay }, undefined, null, "committed", 1]) {
      expect(isReplayed(value)).toBe(false);
    }
  });

  it("R2 [pin] persistence_unavailable_after_mutation from the same store instance is not marked replayed (documented scope)", async () => {
    const storage = new CountingStorage();
    storage.failTerminalSet = true;
    const store = new BrowserIdempotencyStore(storage);
    const transaction = await accepted("r2");
    const first = await execute(transaction, gateAdapter().adapter, store);
    expect(first.outcome).toBe("persistence_unavailable_after_mutation");
    const probe = gateAdapter();
    const again = await execute(transaction, probe.adapter, store);
    expect(again.outcome).toBe("persistence_unavailable_after_mutation");
    expect(probe.mutations()).toBe(0);
    const candidate = await replayPredicate();
    expect(typeof candidate === "function" ? (candidate as ReplayPredicate)(again) : false).toBe(false);
  });

  it("R3 [pin] GC9: an exact replay returns the cached report byte-equal", async () => {
    const storage = new MemoryIdempotencyStorage();
    const transaction = await accepted("r3");
    const report = await execute(transaction, gateAdapter().adapter, new BrowserIdempotencyStore(storage));
    expect(report.outcome).toBe("committed");
    const probe = gateAdapter();
    const replay = await execute(transaction, probe.adapter, new BrowserIdempotencyStore(storage));
    expect(JSON.stringify(replay)).toBe(JSON.stringify(report));
    expect(probe.mutations()).toBe(0);
  });
});
