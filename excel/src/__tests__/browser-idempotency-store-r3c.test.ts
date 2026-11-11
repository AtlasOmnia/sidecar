import { afterEach, describe, expect, it, vi } from "vitest";
import * as executorModule from "../workbookPlanExecutor";
import {
  BrowserIdempotencyStore,
  DeterministicIdempotencyCoordinator,
  createTrustedOperationBoundary,
  executeAcceptedWorkbookPlan,
  prepareWorkbookPlanExecution,
  type AcceptedWorkbookPlanTransaction,
  type IdempotencyStorage,
  type TrustedOperationBoundary,
  type WorkbookPlanExecutionAdapter,
  type WorkbookPlanExecutionReport,
} from "../workbookPlanExecutor";
import { getExcelToolDefinition } from "../excelTools";
import type { WorkbookOperation, WorkbookPlan } from "../generated/workbookPlanSchema";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import { WORKBOOK_PLAN_CAPABILITY_NAMES, type WorkbookPlanPolicyInput } from "../workbookPlanPolicy";

/**
 * R3c: browser idempotency store.
 * - G3: capacity eviction removes only terminal records (markers included) older than the acceptance window.
 * - G4/D4: foreign in-flight records older than the window become terminal restart_rejected markers stamped
 *   with the marking time.
 * - A4 (R2c): a malformed beforeFirstMutation callback returns cancelled without touching the store.
 * - I6: non-finite clock fails closed in claim; finish refuses a restart_rejected report.
 * Self-contained: plan/compile helpers are copied here, nothing is imported from other test files.
 */

/** Test-local acceptance window (D5: 10 min). ACCEPTANCE_HORIZON_MS is asserted equal to this only in GC7. */
const H = 10 * 60 * 1000;
const STORAGE_KEY = "hermes.excel.workbookPlan.idempotency.v1";
const SESSION = "5e550000-0000-4000-8000-000000000001";
const DOCUMENT = "d0c00000-0000-4000-8000-000000000002";

const uuid = (n: number): string => `00000000-0000-4000-8000-${n.toString(16).padStart(12, "0")}`;
const turnUuid = (n: number): string => `7a7a0000-0000-4000-8000-${n.toString(16).padStart(12, "0")}`;
const tokenFor = (turn: number): string => `${SESSION}.${DOCUMENT}.${turnUuid(turn)}`;
const boundaryFor = (turn: number): TrustedOperationBoundary =>
  createTrustedOperationBoundary({ sessionId: SESSION, documentId: DOCUMENT, turnId: turnUuid(turn) });

/** Synthetic store identity (transaction UUID + digest) for direct store-level claims. */
const identity = (n: number): { readonly boundary: string; readonly tx: string; readonly digest: string } => ({
  boundary: tokenFor(n),
  tx: uuid(0x100000 + n),
  digest: `digest-${n}`,
});

interface RawRecord {
  readonly boundary: string;
  readonly transactionId: string;
  readonly digest: string;
  readonly owner: string;
  readonly fence: string;
  readonly state: "in_flight" | "terminal";
  readonly createdAt: number;
  readonly report?: WorkbookPlanExecutionReport;
}

class FakeLocalStorage implements IdempotencyStorage {
  private readonly values = new Map<string, string>();
  readonly writes: { readonly key: string; readonly value: string }[] = [];
  reads = 0;
  throwOnGet = false;
  throwOnSet = false;

  constructor(private readonly quotaBytes = Number.POSITIVE_INFINITY) {}

  getItem(key: string): string | null {
    this.reads += 1;
    if (this.throwOnGet) throw new Error("storage read failed");
    return this.values.get(key) ?? null;
  }

  setItem(key: string, value: string): void {
    if (this.throwOnSet) throw new Error("storage write failed");
    let used = 0;
    for (const [k, v] of this.values) if (k !== key) used += k.length + v.length;
    if (used + key.length + value.length > this.quotaBytes) {
      throw new DOMException("The quota has been exceeded.", "QuotaExceededError");
    }
    this.writes.push({ key, value });
    this.values.set(key, value);
  }

  removeItem(key: string): void {
    this.values.delete(key);
  }

  seed(value: string): void {
    this.values.set(STORAGE_KEY, value);
  }

  raw(): string | null {
    return this.values.get(STORAGE_KEY) ?? null;
  }

  records(): RawRecord[] {
    const raw = this.raw();
    return raw === null ? [] : JSON.parse(raw) as RawRecord[];
  }

  record(boundary: string): RawRecord | undefined {
    return this.records().find((record) => record.boundary === boundary);
  }

  recordJson(boundary: string): string | undefined {
    const record = this.record(boundary);
    return record === undefined ? undefined : JSON.stringify(record);
  }
}

const operation = (address: string, value: string): WorkbookOperation => ({
  kind: "write_values",
  sheet: "Sheet1",
  address,
  values: [[value]],
});

const planFor = (operations: readonly WorkbookOperation[]): WorkbookPlan => ({
  version: 1,
  outcome: "execute",
  summary: "r3c test plan",
  operations,
  verification: [{ kind: "calculation_completed" }],
});

const compileContext = (): VerifiedWorkbookCompileContext => ({
  sheetNames: new Set(["Sheet1"]),
  tablesByName: new Map(),
  supportedPrimitiveNames: new Set(
    ["write_range", "format_range", "run_calculation"].filter((name) => getExcelToolDefinition(name) !== undefined),
  ),
  supportedCapabilities: new Set(["excel_workbook_access", "workbook_plan_v1_operation_audit"]),
  chartSourceResolutionsBySheet: new Map(),
});

type Effect = "write" | "format";

const effectOf = (item: WorkbookOperation): Effect => (item.kind === "format_range" ? "format" : "write");

const policyInputFor = (effects: readonly Effect[], operationCount = effects.length): WorkbookPlanPolicyInput => ({
  requestClassification: "explicit_mutation",
  activeSheetName: "Sheet1",
  selectionScopedRequest: false,
  operationSummaries: effects.map((effect, operationIndex) => ({
    operationIndex,
    effect,
    target: { operationIndex, sheetNames: ["Sheet1"], selectionRelationship: "inside" as const },
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

interface Prepared {
  readonly transaction: AcceptedWorkbookPlanTransaction;
  readonly boundary: TrustedOperationBoundary;
  readonly token: string;
}

/** Real prepareWorkbookPlanExecution with deterministic (boundary-derived) transaction/operation ids. */
const prepared = async (turn: number, operations: readonly WorkbookOperation[] = [operation("A1", `v${turn}`)]): Promise<Prepared> => {
  const boundary = boundaryFor(turn);
  const result = await prepareWorkbookPlanExecution({
    plan: planFor(operations),
    trustedOperationBoundary: boundary,
    compileContext: compileContext(),
    policyInput: policyInputFor(operations.map(effectOf)),
  });
  expect(result.accepted).toBe(true);
  if (!result.accepted) throw new Error("expected accepted plan");
  return { transaction: result.transaction, boundary, token: tokenFor(turn) };
};

const mutationAdapter = (onMutation: () => void | Promise<void>): WorkbookPlanExecutionAdapter => ({
  async executeFirstGroup(context) {
    return context.executeAfterFresh("fresh", async () => {
      await onMutation();
      return { status: "committed" };
    });
  },
  async executeGroup() {
    await onMutation();
    return { status: "committed" };
  },
  async reconcileFailedGroup(group) {
    return group.operations.map(() => "unknown_within_batch" as const);
  },
});

const run = (
  plan: Prepared,
  store: BrowserIdempotencyStore,
  onMutation: () => void | Promise<void>,
  options?: { readonly beforeFirstMutation?: unknown },
): Promise<WorkbookPlanExecutionReport> =>
  executeAcceptedWorkbookPlan(
    plan.transaction,
    plan.boundary,
    mutationAdapter(onMutation),
    store,
    options as unknown as Parameters<typeof executeAcceptedWorkbookPlan>[4],
  );

const markerReport = (transactionId: string): WorkbookPlanExecutionReport => ({
  outcome: "restart_rejected",
  transactionId,
  operationCount: 0,
  operations: [],
  verification: "not_attempted",
  rollback: "not_attempted",
});

const expectedMarker = (original: RawRecord, markTime: number): RawRecord => ({
  boundary: original.boundary,
  transactionId: original.transactionId,
  digest: original.digest,
  owner: original.owner,
  fence: original.fence,
  state: "terminal",
  createdAt: markTime,
  report: markerReport(original.transactionId),
});

const terminalRecord = (n: number, createdAt: number, outcome: "committed" | "restart_rejected"): RawRecord => {
  const id = identity(n);
  return {
    boundary: id.boundary,
    transactionId: id.tx,
    digest: id.digest,
    owner: uuid(0x200000 + n),
    fence: uuid(0x300000 + n),
    state: "terminal",
    createdAt,
    report: { ...markerReport(id.tx), outcome },
  };
};

const inFlightRecord = (n: number, createdAt: number): RawRecord => {
  const id = identity(n);
  return {
    boundary: id.boundary,
    transactionId: id.tx,
    digest: id.digest,
    owner: uuid(0x200000 + n),
    fence: uuid(0x300000 + n),
    state: "in_flight",
    createdAt,
  };
};

const claimId = (store: BrowserIdempotencyStore, n: number) => {
  const id = identity(n);
  return store.claim(id.boundary, id.tx, id.digest);
};

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

describe("R3c RED: window-bounded eviction (G3)", () => {
  it("RC1 keeps a young terminal record under entry pressure; the replay never mutates twice", async () => {
    const now = 0;
    const storage = new FakeLocalStorage();
    const options = { clock: () => now, maxEntries: 2 };
    const a = await prepared(1);
    const c = await prepared(3);
    let mutationsA = 0;
    let mutationsC = 0;
    expect((await run(a, new BrowserIdempotencyStore(storage, options), () => { mutationsA += 1; })).outcome).toBe("committed");
    const holder = new BrowserIdempotencyStore(storage, options);
    expect((await claimId(holder, 2)).kind).toBe("claimed");

    const cReport = await run(c, new BrowserIdempotencyStore(storage, options), () => { mutationsC += 1; });
    let replayCallbacks = 0;
    const replay = await run(a, new BrowserIdempotencyStore(storage, options), () => { mutationsA += 1; }, {
      beforeFirstMutation: async () => { replayCallbacks += 1; },
    });
    expect({
      c: cReport.outcome,
      mutationsC,
      replay: replay.outcome,
      mutationsA,
      replayCallbacks,
      aState: storage.record(a.token)?.state,
    }).toEqual({
      c: "persistence_unavailable",
      mutationsC: 0,
      replay: "committed",
      mutationsA: 1,
      replayCallbacks: 0,
      aState: "terminal",
    });
  });

  it("RC2 keeps a young terminal record under byte pressure", async () => {
    const now = 0;
    const a = await prepared(1);
    const c = await prepared(3);
    // Probe the exact size of [A terminal, B in flight] with unbounded storage.
    const probe = new FakeLocalStorage();
    await run(a, new BrowserIdempotencyStore(probe, { clock: () => now }), () => undefined);
    await claimId(new BrowserIdempotencyStore(probe, { clock: () => now }), 2);
    const fit = new TextEncoder().encode(probe.raw()!).byteLength;

    const storage = new FakeLocalStorage();
    const options = { clock: () => now, maxEntries: 128, maxBytes: fit };
    let mutationsA = 0;
    let mutationsC = 0;
    expect((await run(a, new BrowserIdempotencyStore(storage, options), () => { mutationsA += 1; })).outcome).toBe("committed");
    expect((await claimId(new BrowserIdempotencyStore(storage, options), 2)).kind).toBe("claimed");
    const cReport = await run(c, new BrowserIdempotencyStore(storage, options), () => { mutationsC += 1; });
    const replay = await run(a, new BrowserIdempotencyStore(storage, options), () => { mutationsA += 1; });
    expect({
      c: cReport.outcome,
      mutationsC,
      replay: replay.outcome,
      mutationsA,
      aState: storage.record(a.token)?.state,
    }).toEqual({
      c: "persistence_unavailable",
      mutationsC: 0,
      replay: "committed",
      mutationsA: 1,
      aState: "terminal",
    });
  });

  it("RC3 evicts a terminal record only strictly after the window (now = H keeps, now = H + 1 evicts)", async () => {
    let now = 0;
    const storage = new FakeLocalStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => now, maxEntries: 2 });
    const a = await prepared(1);
    expect((await run(a, store, () => undefined)).outcome).toBe("committed");
    expect((await claimId(store, 2)).kind).toBe("claimed");

    now = H;
    const atWindow = (await claimId(store, 3)).kind;
    const keptAtWindow = storage.record(a.token)?.state;
    now = H + 1;
    const afterWindow = (await claimId(store, 3)).kind;
    expect({ atWindow, keptAtWindow, afterWindow, aAfter: storage.record(a.token) }).toEqual({
      atWindow: "persistence_unavailable",
      keptAtWindow: "terminal",
      afterWindow: "claimed",
      aAfter: undefined,
    });
  });

  it("RC4 finish under byte pressure with only a young terminal reports persistence_unavailable_after_mutation", async () => {
    const now = 0;
    const a = await prepared(1);
    const b = await prepared(2);
    const probe = new FakeLocalStorage();
    await run(a, new BrowserIdempotencyStore(probe, { clock: () => now }), () => undefined);
    await new BrowserIdempotencyStore(probe, { clock: () => now }).claim(b.token, b.transaction.transactionId, b.transaction.semanticDigest);
    const fitInFlight = new TextEncoder().encode(probe.raw()!).byteLength;

    const storage = new FakeLocalStorage();
    const options = { clock: () => now, maxBytes: fitInFlight };
    expect((await run(a, new BrowserIdempotencyStore(storage, options), () => undefined)).outcome).toBe("committed");
    const aBefore = storage.recordJson(a.token);
    let mutationsB = 0;
    const report = await run(b, new BrowserIdempotencyStore(storage, options), () => { mutationsB += 1; });
    expect({
      outcome: report.outcome,
      operations: report.operations.map((item) => item.state),
      mutationsB,
      aUnchanged: storage.recordJson(a.token) === aBefore,
      bState: storage.record(b.token)?.state,
    }).toEqual({
      outcome: "persistence_unavailable_after_mutation",
      operations: ["committed"],
      mutationsB: 1,
      aUnchanged: true,
      bState: "in_flight",
    });
    expect(report.transactionId).toBe(b.transaction.transactionId);
  });
});

describe("R3c RED: orphaned in-flight markers (G4/D4)", () => {
  it("RC5 marks a foreign in-flight record older than the window, stamped with the marking time", async () => {
    let now = 0;
    const storage = new FakeLocalStorage();
    const s1 = new BrowserIdempotencyStore(storage, { clock: () => now });
    const t = identity(1);
    expect((await claimId(s1, 1)).kind).toBe("claimed");
    const original = storage.record(t.boundary)!;
    expect(original.state).toBe("in_flight");

    now = H + 1;
    const s2 = new BrowserIdempotencyStore(storage, { clock: () => now });
    expect((await claimId(s2, 2)).kind).toBe("claimed");
    const marked = storage.record(t.boundary)!;
    expect(marked).toEqual(expectedMarker(original, H + 1));
    expect(marked.createdAt).toBe(H + 1);
    expect(Object.keys(marked).sort()).toEqual(["boundary", "createdAt", "digest", "fence", "owner", "report", "state", "transactionId"]);
    console.log(`[r3c-measure] RC5 marker JSON (${new TextEncoder().encode(JSON.stringify(marked)).byteLength} bytes): ${JSON.stringify(marked)}`);
  });

  it("RC6 128 orphans: a new claim succeeds only once the markers are older than the window from marking", async () => {
    let now = 0;
    const storage = new FakeLocalStorage();
    for (let n = 1; n <= 128; n += 1) {
      expect((await claimId(new BrowserIdempotencyStore(storage, { clock: () => now }), n)).kind).toBe("claimed");
    }
    const v = 1000;

    now = H + 1;
    const first = (await claimId(new BrowserIdempotencyStore(storage, { clock: () => now }), v)).kind;
    const afterFirst = storage.records();
    const rawAfterFirst = storage.raw();
    expect({
      first,
      count: afterFirst.length,
      markers: afterFirst.filter((record) => record.state === "terminal" && record.report?.outcome === "restart_rejected").length,
      inFlight: afterFirst.filter((record) => record.state === "in_flight").length,
      allStampedAtMarking: afterFirst.every((record) => record.createdAt === H + 1),
    }).toEqual({ first: "persistence_unavailable", count: 128, markers: 128, inFlight: 0, allStampedAtMarking: true });
    console.log(`[r3c-measure] RC6 storage bytes at 128 markers: ${new TextEncoder().encode(rawAfterFirst!).byteLength}`);

    now = 2 * H + 1;
    const writesBefore = storage.writes.length;
    expect((await claimId(new BrowserIdempotencyStore(storage, { clock: () => now }), v)).kind).toBe("persistence_unavailable");
    expect(storage.raw()).toBe(rawAfterFirst);
    expect(storage.writes.length).toBe(writesBefore);

    now = 2 * H + 2;
    const store = new BrowserIdempotencyStore(storage, { clock: () => now });
    expect((await claimId(store, v)).kind).toBe("claimed");
    const after = storage.records();
    expect(after.find((record) => record.boundary === identity(1).boundary)).toBeUndefined();
    expect(after.filter((record) => record.state === "terminal" && record.report?.outcome === "restart_rejected")).toHaveLength(127);
    expect(after.filter((record) => record.state === "in_flight").map((record) => record.boundary)).toEqual([identity(v).boundary]);
    expect(typeof store.reconciliationRequiredCount).toBe("function");
    expect(await store.reconciliationRequiredCount()).toBe(127);
  });

  it("RC7 a marker is non-reusable for any digest", async () => {
    let now = 0;
    const storage = new FakeLocalStorage();
    expect((await claimId(new BrowserIdempotencyStore(storage, { clock: () => now }), 1)).kind).toBe("claimed");
    now = H + 1;
    const s2 = new BrowserIdempotencyStore(storage, { clock: () => now });
    expect((await claimId(s2, 2)).kind).toBe("claimed");
    const t = identity(1);
    const same = (await s2.claim(t.boundary, t.tx, t.digest)).kind;
    const other = (await s2.claim(t.boundary, uuid(0x999), "other-digest")).kind;
    expect({ same, other }).toEqual({ same: "restart_rejected", other: "restart_rejected" });
  });

  it("RC8 exposes a content-free marker count that never writes and is undefined when storage is unavailable", async () => {
    let now = 0;
    const storage = new FakeLocalStorage();
    expect((await claimId(new BrowserIdempotencyStore(storage, { clock: () => now }), 1)).kind).toBe("claimed");
    now = H + 1;
    const s2 = new BrowserIdempotencyStore(storage, { clock: () => now });
    expect((await claimId(s2, 2)).kind).toBe("claimed");
    // One more orphan, below the window.
    expect((await claimId(new BrowserIdempotencyStore(storage, { clock: () => now }), 3)).kind).toBe("claimed");

    expect(typeof s2.reconciliationRequiredCount).toBe("function");
    const writesBefore = storage.writes.length;
    const rawBefore = storage.raw();
    expect(await s2.reconciliationRequiredCount()).toBe(1);
    expect(storage.writes.length).toBe(writesBefore);
    expect(storage.raw()).toBe(rawBefore);
    expect(storage.record(identity(3).boundary)?.state).toBe("in_flight");

    storage.throwOnGet = true;
    expect(await s2.reconciliationRequiredCount()).toBeUndefined();
  });

  it("RC14 a fresh marker blocks its own key and capacity for one window after marking", async () => {
    let now = 0;
    const storage = new FakeLocalStorage();
    const options = { clock: () => now, maxEntries: 2 };
    expect((await claimId(new BrowserIdempotencyStore(storage, options), 1)).kind).toBe("claimed");
    now = H + 1;
    const s2 = new BrowserIdempotencyStore(storage, options);
    expect((await claimId(s2, 2)).kind).toBe("claimed");
    const t = identity(1);
    const markerBefore = storage.recordJson(t.boundary);

    now = 2 * H + 1;
    const writesBefore = storage.writes.length;
    const other = (await s2.claim(t.boundary, uuid(0x999), "other-digest")).kind;
    const w = (await claimId(s2, 3)).kind;
    expect({
      other,
      w,
      markerUnchanged: storage.recordJson(t.boundary) === markerBefore,
      markerState: storage.record(t.boundary)?.report?.outcome,
      writes: storage.writes.length - writesBefore,
    }).toEqual({ other: "restart_rejected", w: "persistence_unavailable", markerUnchanged: true, markerState: "restart_rejected", writes: 0 });
  });
});

describe("R3c RED: A4 and fail-closed hardening", () => {
  it("RC9 a malformed callback returns cancelled without any store access; a corrected retry executes once", async () => {
    const storage = new FakeLocalStorage();
    const plan = await prepared(1);
    let mutations = 0;
    const malformed = await run(plan, new BrowserIdempotencyStore(storage), () => { mutations += 1; }, {
      beforeFirstMutation: "not-a-callback",
    });
    const afterRun1 = { outcome: malformed.outcome, mutations, raw: storage.raw(), reads: storage.reads, writes: storage.writes.length };

    let callbacks = 0;
    const corrected = await run(plan, new BrowserIdempotencyStore(storage), () => { mutations += 1; }, {
      beforeFirstMutation: async () => { callbacks += 1; },
    });
    expect({ afterRun1, corrected: corrected.outcome, callbacks, mutations }).toEqual({
      afterRun1: { outcome: "cancelled", mutations: 0, raw: null, reads: 0, writes: 0 },
      corrected: "committed",
      callbacks: 1,
      mutations: 1,
    });
    expect(malformed.operations.every((item) => item.state === "not_committed" && item.status === "not_started")).toBe(true);
  });

  it("RC12 finish refuses a restart_rejected report, so a marker cannot be forged through the API", async () => {
    const storage = new FakeLocalStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => 0 });
    const t = identity(1);
    const claim = await store.claim(t.boundary, t.tx, t.digest);
    expect(claim.kind).toBe("claimed");
    if (claim.kind !== "claimed") throw new Error("expected claim");
    const finished = await store.finish(t.boundary, t.tx, t.digest, markerReport(t.tx), claim.fence);
    expect({ finished, state: storage.record(t.boundary)?.state }).toEqual({ finished: false, state: "in_flight" });
  });

  it("RC13 a non-finite clock fails closed with no storage write", async () => {
    const storage = new FakeLocalStorage();
    storage.seed(JSON.stringify([terminalRecord(1, 0, "committed")]));
    const rawBefore = storage.raw();
    const store = new BrowserIdempotencyStore(storage, { clock: () => Number.NaN });
    const kind = (await claimId(store, 2)).kind;
    expect({ kind, writes: storage.writes.length, unchanged: storage.raw() === rawBefore }).toEqual({
      kind: "persistence_unavailable",
      writes: 0,
      unchanged: true,
    });
  });
});

describe("R3c GREEN pins", () => {
  it("GC1 leaves a foreign orphan untouched at exactly the window", async () => {
    let now = 0;
    const storage = new FakeLocalStorage();
    expect((await claimId(new BrowserIdempotencyStore(storage, { clock: () => now }), 1)).kind).toBe("claimed");
    now = H;
    const s2 = new BrowserIdempotencyStore(storage, { clock: () => now });
    expect((await claimId(s2, 2)).kind).toBe("claimed");
    const t = storage.record(identity(1).boundary)!;
    expect(t.state).toBe("in_flight");
    expect(t.createdAt).toBe(0);
    expect((await claimId(s2, 1)).kind).toBe("restart_rejected");
  });

  it("GC2 never marks the store's own in-flight record", async () => {
    let now = 0;
    const storage = new FakeLocalStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => now });
    const t = identity(1);
    const claim = await store.claim(t.boundary, t.tx, t.digest);
    if (claim.kind !== "claimed") throw new Error("expected claim");
    now = H + 1;
    expect((await claimId(store, 2)).kind).toBe("claimed");
    expect(storage.record(t.boundary)?.state).toBe("in_flight");
    expect((await store.claim(t.boundary, t.tx, t.digest)).kind).toBe("in_flight");
    const committed: WorkbookPlanExecutionReport = { ...markerReport(t.tx), outcome: "committed" };
    expect(await store.finish(t.boundary, t.tx, t.digest, committed, claim.fence)).toBe(true);
    expect(storage.record(t.boundary)?.state).toBe("terminal");
    expect(storage.record(t.boundary)?.report?.outcome).toBe("committed");
  });

  it("GC3 marks an orphan exactly once under concurrent claims and never rewrites it", async () => {
    let now = 0;
    const storage = new FakeLocalStorage();
    const coordinator = new DeterministicIdempotencyCoordinator();
    const make = () => new BrowserIdempotencyStore(storage, { clock: () => now, coordinator });
    expect((await claimId(make(), 1)).kind).toBe("claimed");
    const t = identity(1).boundary;
    now = H + 1;
    const results = await Promise.all([claimId(make(), 2), claimId(make(), 3), claimId(make(), 4)]);
    expect(results.map((result) => result.kind)).toEqual(["claimed", "claimed", "claimed"]);
    const markerJson = storage.recordJson(t);
    now = H + 5;
    expect((await claimId(make(), 5)).kind).toBe("claimed");
    expect(storage.recordJson(t)).toBe(markerJson);
    expect(storage.record(t)?.createdAt).toBe(H + 1);

    const statesPerWrite = storage.writes.map(({ value }) => {
      const record = (JSON.parse(value) as RawRecord[]).find((item) => item.boundary === t);
      return record === undefined ? "absent" : `${record.state}:${record.report?.outcome ?? "-"}:${record.createdAt}`;
    });
    const transitions = statesPerWrite.filter((state, index) =>
      index > 0 && statesPerWrite[index - 1].startsWith("in_flight") && state.startsWith("terminal:restart_rejected"));
    expect(transitions).toHaveLength(1);
    expect(new Set(statesPerWrite.filter((state) => state.startsWith("terminal")))).toEqual(new Set([`terminal:restart_rejected:${H + 1}`]));
  });

  it("GC4 a hung owner that resumes after marking keeps the marker and reports persistence_unavailable_after_mutation", async () => {
    let now = 0;
    const storage = new FakeLocalStorage();
    const plan = await prepared(1);
    const s1 = new BrowserIdempotencyStore(storage, { clock: () => now });
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    let entered!: () => void;
    const inMutation = new Promise<void>((resolve) => { entered = resolve; });
    const execution = run(plan, s1, async () => { entered(); await gate; });
    await inMutation;

    now = H + 1;
    expect((await claimId(new BrowserIdempotencyStore(storage, { clock: () => now }), 2)).kind).toBe("claimed");
    const markerJson = storage.recordJson(plan.token);
    expect(storage.record(plan.token)?.report?.outcome).toBe("restart_rejected");

    release();
    const report = await execution;
    expect(report.outcome).toBe("persistence_unavailable_after_mutation");
    expect(report.operations.map((item) => item.state)).toEqual(["committed"]);
    expect(storage.recordJson(plan.token)).toBe(markerJson);
    const again = await s1.claim(plan.token, plan.transaction.transactionId, plan.transaction.semanticDigest);
    expect(again.kind).toBe("persistence_unavailable_after_mutation");
  });

  it("GC5 a marked boundary executes nothing, even through a reconstructed boundary and transaction", async () => {
    let now = 0;
    const storage = new FakeLocalStorage();
    const original = await prepared(1);
    expect((await new BrowserIdempotencyStore(storage, { clock: () => now }).claim(
      original.token, original.transaction.transactionId, original.transaction.semanticDigest,
    )).kind).toBe("claimed");
    now = H + 1;
    expect((await claimId(new BrowserIdempotencyStore(storage, { clock: () => now }), 2)).kind).toBe("claimed");

    const reconstructed = await prepared(1);
    const writesBefore = storage.writes.length;
    let callbacks = 0;
    let mutations = 0;
    const report = await run(reconstructed, new BrowserIdempotencyStore(storage, { clock: () => now }), () => { mutations += 1; }, {
      beforeFirstMutation: async () => { callbacks += 1; },
    });
    expect(report.outcome).toBe("restart_rejected");
    expect(callbacks).toBe(0);
    expect(mutations).toBe(0);
    expect(storage.writes.length).toBe(writesBefore);
  });

  it("GC6 TTL prunes plain terminals at ttl + 1 and markers at markTime + ttl + 1, with no capacity pressure", async () => {
    let now = 0;
    const ttlMs = 2 * H;
    const storage = new FakeLocalStorage();
    const make = () => new BrowserIdempotencyStore(storage, { clock: () => now, ttlMs });
    const a = await prepared(1);
    expect((await run(a, make(), () => undefined)).outcome).toBe("committed");
    now = ttlMs;
    expect((await claimId(make(), 10)).kind).toBe("claimed");
    expect(storage.record(a.token)?.state).toBe("terminal");
    now = ttlMs + 1;
    expect((await claimId(make(), 11)).kind).toBe("claimed");
    expect(storage.record(a.token)).toBeUndefined();

    // Marker: the orphan is claimed at 0, marked at M, and lives until M + ttl.
    const markerStorage = new FakeLocalStorage();
    const makeM = () => new BrowserIdempotencyStore(markerStorage, { clock: () => now, ttlMs });
    now = 0;
    expect((await claimId(makeM(), 1)).kind).toBe("claimed");
    const M = H + 1;
    now = M;
    expect((await claimId(makeM(), 2)).kind).toBe("claimed");
    expect(markerStorage.record(identity(1).boundary)?.createdAt).toBe(M);
    now = M + ttlMs;
    expect((await claimId(makeM(), 3)).kind).toBe("claimed");
    expect(markerStorage.record(identity(1).boundary)?.report?.outcome).toBe("restart_rejected");
    now = M + ttlMs + 1;
    expect((await claimId(makeM(), 4)).kind).toBe("claimed");
    expect(markerStorage.record(identity(1).boundary)).toBeUndefined();
  });

  it("GC6b reproduces the existing terminal TTL / in-flight bound behaviour unchanged", async () => {
    let now = 0;
    const storage = new FakeLocalStorage();
    const options = { clock: () => now, ttlMs: 10, maxEntries: 1, maxBytes: 4096 };
    const store = new BrowserIdempotencyStore(storage, options);
    const first = await prepared(1);
    const second = await prepared(1, [operation("B1", "b")]);
    const inFlight = await prepared(2, [operation("C1", "c")]);
    await run(first, store, () => undefined);
    now = 20;
    expect((await store.claim(first.token, first.transaction.transactionId, first.transaction.semanticDigest)).kind).toBe("claimed");
    expect((await store.claim(inFlight.token, inFlight.transaction.transactionId, inFlight.transaction.semanticDigest)).kind).toBe("persistence_unavailable");
    now = 30;
    await run(second, store, () => undefined);
    const restarted = new BrowserIdempotencyStore(storage, options);
    expect((await restarted.claim(first.token, first.transaction.transactionId, first.transaction.semanticDigest)).kind).toBe("restart_rejected");
  });

  it("GC7 production defaults: 10 min window, 24 h retention, default localStorage + navigator.locks path", async () => {
    expect((executorModule as Record<string, unknown>).ACCEPTANCE_HORIZON_MS).toBe(H);
    expect(H).toBe(600_000);
    const fake = new FakeLocalStorage();
    const queues = new Map<string, Promise<unknown>>();
    const fakeLocks = {
      request(name: string, _options: unknown, callback: (lock: unknown) => Promise<unknown>): Promise<unknown> {
        const previous = queues.get(name) ?? Promise.resolve();
        const next = previous.then(() => callback({ name, mode: "exclusive" }), () => callback({ name, mode: "exclusive" }));
        queues.set(name, next.catch(() => undefined));
        return next;
      },
    };
    vi.stubGlobal("localStorage", fake);
    vi.stubGlobal("navigator", { locks: fakeLocks });
    vi.useFakeTimers({ toFake: ["Date"] });
    const t0 = Date.UTC(2026, 9, 6, 12, 0, 0);
    vi.setSystemTime(t0);

    // Window <= retention, observed through behaviour: a terminal at 23 h 59 m is still replayed.
    const a = await prepared(1);
    let mutations = 0;
    expect((await run(a, new BrowserIdempotencyStore(), () => { mutations += 1; })).outcome).toBe("committed");
    vi.setSystemTime(t0 + 23 * 60 * 60 * 1000 + 59 * 60 * 1000);
    expect((await run(a, new BrowserIdempotencyStore(), () => { mutations += 1; })).outcome).toBe("committed");
    expect(mutations).toBe(1);
    expect(fake.writes.every((write) => write.key === STORAGE_KEY)).toBe(true);

    const t1 = Date.now();
    expect((await claimId(new BrowserIdempotencyStore(), 2)).kind).toBe("claimed");
    vi.setSystemTime(t1 + 9 * 60 * 1000 + 59 * 1000);
    expect((await claimId(new BrowserIdempotencyStore(), 3)).kind).toBe("claimed");
    expect(fake.record(identity(2).boundary)?.state).toBe("in_flight");
    expect(fake.record(identity(2).boundary)?.createdAt).toBe(t1);

    const markTime = t1 + 10 * 60 * 1000 + 1;
    vi.setSystemTime(markTime);
    expect((await claimId(new BrowserIdempotencyStore(), 4)).kind).toBe("claimed");
    const marked = fake.record(identity(2).boundary)!;
    expect(marked.state).toBe("terminal");
    expect(marked.report?.outcome).toBe("restart_rejected");
    expect(marked.createdAt).toBe(markTime);
    expect(fake.record(identity(3).boundary)?.state).toBe("in_flight");
  });

  it("GC8 fails closed with zero mutation and unchanged bytes on corrupt storage and failed marking writes", async () => {
    const orphanJson = JSON.stringify([inFlightRecord(1, 0)]);
    const cases: { readonly name: string; readonly storage: FakeLocalStorage }[] = [];
    const corrupt = new FakeLocalStorage();
    corrupt.seed(`${orphanJson.slice(0, -1)},{`);
    cases.push({ name: "corrupt-json", storage: corrupt });
    const unparseable = new FakeLocalStorage();
    unparseable.seed(JSON.stringify([inFlightRecord(1, 0), { state: "corrupt" }]));
    cases.push({ name: "unparseable-record", storage: unparseable });
    const throwing = new FakeLocalStorage();
    throwing.seed(orphanJson);
    throwing.throwOnSet = true;
    cases.push({ name: "throw-on-set", storage: throwing });
    // Quota fits the in-flight orphan but not the larger marker.
    const quota = new FakeLocalStorage(STORAGE_KEY.length + orphanJson.length + 20);
    quota.seed(orphanJson);
    cases.push({ name: "quota-exceeded", storage: quota });

    const plan = await prepared(2);
    const outcomes: Record<string, unknown> = {};
    for (const { name, storage } of cases) {
      const rawBefore = storage.raw();
      let mutations = 0;
      const report = await run(plan, new BrowserIdempotencyStore(storage, { clock: () => H + 1 }), () => { mutations += 1; });
      outcomes[name] = { outcome: report.outcome, mutations, unchanged: storage.raw() === rawBefore };
    }
    const expected = { outcome: "persistence_unavailable", mutations: 0, unchanged: true };
    expect(outcomes).toEqual({
      "corrupt-json": expected,
      "unparseable-record": expected,
      "throw-on-set": expected,
      "quota-exceeded": expected,
    });
  });

  it("GC9 exact replay returns the cached report byte-equal; a changed digest on a normal terminal is a conflict", async () => {
    const storage = new FakeLocalStorage();
    const plan = await prepared(1);
    const report = await run(plan, new BrowserIdempotencyStore(storage), () => undefined);
    expect(report.outcome).toBe("committed");
    let mutations = 0;
    const replay = await run(plan, new BrowserIdempotencyStore(storage), () => { mutations += 1; });
    expect(JSON.stringify(replay)).toBe(JSON.stringify(report));
    expect(mutations).toBe(0);
    const changed = await prepared(1, [operation("A1", "changed")]);
    expect((await run(changed, new BrowserIdempotencyStore(storage), () => { mutations += 1; })).outcome).toBe("conflict");
    expect(mutations).toBe(0);
  });

  it("GC10 a 30-operation terminal record leaves room for at least 40 records in 256 KiB", async () => {
    // write_values plans exceed the validation work budget beyond 15 operations, so the schema-max (30)
    // plan uses format_range; per-operation report bytes do not depend on the primitive's arguments.
    const operations: WorkbookOperation[] = Array.from({ length: 30 }, (_, index) => ({
      kind: "format_range",
      sheet: "Sheet1",
      address: `A${index + 1}`,
      format: { numberFormat: "0.00" },
    }));
    const plan = await prepared(1, operations);
    const storage = new FakeLocalStorage();
    let mutations = 0;
    const report = await run(plan, new BrowserIdempotencyStore(storage), () => { mutations += 1; });
    expect(report.outcome).toBe("committed");
    expect(report.operationCount).toBe(30);
    const size = new TextEncoder().encode(storage.recordJson(plan.token)!).byteLength;
    const capacity = Math.floor((256 * 1024) / size);
    console.log(`[r3c-measure] GC10 30-operation terminal record bytes=${size} floor(256KiB/size)=${capacity}`);
    expect(capacity).toBeGreaterThanOrEqual(40);
  });

  it("GC11 keeps storage key v1 and loads stored markers through the unchanged record schema", async () => {
    const storage = new FakeLocalStorage();
    storage.seed(JSON.stringify([terminalRecord(1, 5, "restart_rejected"), inFlightRecord(2, 5)]));
    const store = new BrowserIdempotencyStore(storage, { clock: () => 10 });
    expect((await claimId(store, 3)).kind).toBe("claimed");
    expect(storage.writes.map((write) => write.key)).toEqual([STORAGE_KEY]);
    expect(storage.record(identity(1).boundary)).toEqual(terminalRecord(1, 5, "restart_rejected"));
  });

  it("GC13 an aged marker and an aged plain terminal evict alike, first in array order; a young terminal is kept", async () => {
    let now = 2 * H + 2;
    const storage = new FakeLocalStorage();
    const p = terminalRecord(1, 0, "committed");
    const m = terminalRecord(2, H + 1, "restart_rejected");
    const y = terminalRecord(3, 2 * H + 2, "committed");
    storage.seed(JSON.stringify([p, m, y]));
    const store = new BrowserIdempotencyStore(storage, { clock: () => now, maxEntries: 3 });
    expect((await claimId(store, 4)).kind).toBe("claimed");
    expect(storage.records().map((record) => record.boundary)).toEqual([m.boundary, y.boundary, identity(4).boundary]);
    expect((await claimId(store, 5)).kind).toBe("claimed");
    expect(storage.records().map((record) => record.boundary)).toEqual([y.boundary, identity(4).boundary, identity(5).boundary]);
    now = 3 * H + 2;
    expect((await claimId(store, 6)).kind).toBe("persistence_unavailable");
    expect(storage.record(y.boundary)).toEqual(y);
  });
});
