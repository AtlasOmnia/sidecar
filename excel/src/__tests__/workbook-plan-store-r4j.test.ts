import { describe, expect, it } from "vitest";
import {
  BrowserIdempotencyStore,
  type IdempotencyCoordinator,
  type IdempotencyStorage,
  type WorkbookPlanExecutionReport,
} from "../workbookPlanExecutor";

/**
 * R4j J3 (R4e-2 lane-A F1 / lane-B A3): a refused or failed finish is remembered under the
 * transactionId the claim recorded, never under the caller's (foreign, look-alike or malformed)
 * report id. Only when this runtime owns the claim; otherwise the base content is kept (OQ-J3).
 * - JS1-JS7: RED at the base (the remembered report echoes the caller's id).
 * - JS-p1..JS-p3: parity pins, identical assertions at the base and at HEAD.
 * UUIDS, token, CountingStorage and committedReport are COPIED as text from
 * workbook-plan-store-r4e2.test.ts:17-60 at f3e3c0a (that file is neither imported nor edited).
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
const token = (turnIndex: number): string => `${UUIDS[0]}.${UUIDS[1]}.${UUIDS[turnIndex]}`;

class CountingStorage implements IdempotencyStorage {
  value: string | null = null;
  reads = 0;
  writes = 0;

  getItem(): string | null {
    this.reads += 1;
    return this.value;
  }

  setItem(_key: string, value: string): void {
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

const committedReport = (transactionId: string): WorkbookPlanExecutionReport => ({
  outcome: "committed",
  transactionId,
  operationCount: 0,
  operations: [],
  verification: "not_attempted",
  rollback: "not_attempted",
});

// ---------------------------------------------------------------------------
// Local helpers (R4j)
// ---------------------------------------------------------------------------
const P_U_A_M = "persistence_unavailable_after_mutation";
const DIGEST = "digest-r4j";

/** The independent expectation for a remembered report: the caller's report, outcome p_u_a_m, id = `id`. */
const remembered = (report: WorkbookPlanExecutionReport, id: string): WorkbookPlanExecutionReport =>
  ({ ...JSON.parse(JSON.stringify(report)) as WorkbookPlanExecutionReport, outcome: P_U_A_M, transactionId: id });

async function claimed(store: BrowserIdempotencyStore, boundary = token(2), id = UUIDS[3]): Promise<string> {
  const claim = await store.claim(boundary, id, DIGEST);
  if (claim.kind !== "claimed") throw new Error(`expected claimed, got ${claim.kind}`);
  return claim.fence;
}

async function reclaim(store: BrowserIdempotencyStore, storage: CountingStorage, boundary = token(2), id = UUIDS[3]) {
  const again = await store.claim(boundary, id, DIGEST);
  return {
    kind: again.kind,
    report: "report" in again ? again.report : undefined,
    stored: storage.records().map((record) => `${String(record.transactionId)}:${String(record.state)}`),
    markers: await store.reconciliationRequiredCount(),
  };
}

describe("R4j J3: a refused finish is remembered under the claimed record's transactionId", () => {
  // The look-alike row claims an id with hex letters so that its uppercase spelling differs.
  const LETTERED = "abcdef01-2345-4678-9abc-def012345678";
  it.each([
    ["JS1 a foreign UUID", UUIDS[3], UUIDS[4]],
    ["JS2 an uppercase look-alike", LETTERED, LETTERED.toUpperCase()],
    ["JS3 a non-UUID", UUIDS[3], "not-a-uuid"],
  ] as const)("%s report id: the re-claim answers p_u_a_m with the record's id; the operations are the caller's; record in_flight; no marker", async (_label, recordId, foreign) => {
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => 1000 });
    const fence = await claimed(store, token(2), recordId);
    const report = committedReport(foreign);
    const finished = await store.finish(token(2), recordId, DIGEST, report, fence);
    const again = await reclaim(store, storage, token(2), recordId);
    expect({ finished, ...again }).toEqual({
      finished: false,
      kind: P_U_A_M,
      report: remembered(report, recordId),
      stored: [`${recordId}:in_flight`],
      markers: 0,
    });
    expect(again.report?.transactionId).toBe(recordId);
    expect(again.report?.operations).toEqual(report.operations);
  });

  it("JS4 the index<0 path: a wrong fence (correct id), then a finish whose argument AND report id are foreign: the record's id both times", async () => {
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => 1000 });
    const fence = await claimed(store);
    const wrongFence = UUIDS[7];
    expect(await store.finish(token(2), UUIDS[3], DIGEST, committedReport(UUIDS[3]), wrongFence)).toBe(false);
    const first = await reclaim(store, storage);
    // Foreign finish argument too (jm7: stamping the finish argument instead of the owned record fails here).
    expect(await store.finish(token(2), UUIDS[4], DIGEST, committedReport(UUIDS[4]), fence)).toBe(false);
    const second = await reclaim(store, storage);
    // And a foreign report id with the right argument (refusal path).
    expect(await store.finish(token(2), UUIDS[3], DIGEST, committedReport(UUIDS[5]), fence)).toBe(false);
    const third = await reclaim(store, storage);
    expect([first, second, third].map((row) => ({ kind: row.kind, id: row.report?.transactionId, stored: row.stored, markers: row.markers })))
      .toEqual([1, 2, 3].map(() => ({ kind: P_U_A_M, id: UUIDS[3], stored: [`${UUIDS[3]}:in_flight`], markers: 0 })));
  });

  it("JS5 an invalid finish clock (the first guard) with a foreign report id: the record's id", async () => {
    let now = 1000;
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => now });
    const fence = await claimed(store);
    now = Number.NaN;
    const report = committedReport(UUIDS[4]);
    expect(await store.finish(token(2), UUIDS[3], DIGEST, report, fence)).toBe(false);
    now = 2000;
    const again = await reclaim(store, storage);
    expect(again).toEqual({ kind: P_U_A_M, report: remembered(report, UUIDS[3]), stored: [`${UUIDS[3]}:in_flight`], markers: 0 });
  });

  it("JS6 the coordinator throws on finish (outer catch) with a foreign report id: the record's id", async () => {
    let calls = 0;
    const coordinator: IdempotencyCoordinator = {
      async runExclusive<T>(_name: string, task: () => Promise<T>): Promise<T> {
        calls += 1;
        if (calls === 2) throw new Error("lock unavailable");
        return task();
      },
    };
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => 1000, coordinator });
    const fence = await claimed(store);
    const report = committedReport(UUIDS[4]);
    expect(await store.finish(token(2), UUIDS[3], DIGEST, report, fence)).toBe(false);
    const again = await reclaim(store, storage);
    expect(calls).toBe(4);
    expect(again).toEqual({ kind: P_U_A_M, report: remembered(report, UUIDS[3]), stored: [`${UUIDS[3]}:in_flight`], markers: 0 });
  });

  it("JS7 a report whose transactionId getter throws (fallback branch): the fixed fallback report carries the record's id", async () => {
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => 1000 });
    const fence = await claimed(store);
    const hostile = { ...committedReport(UUIDS[3]) } as Record<string, unknown>;
    Object.defineProperty(hostile, "transactionId", { enumerable: true, get: () => { throw new Error("getter boom"); } });
    expect(await store.finish(token(2), UUIDS[3], DIGEST, hostile as unknown as WorkbookPlanExecutionReport, fence)).toBe(false);
    const again = await reclaim(store, storage);
    expect(again).toEqual({
      kind: P_U_A_M,
      report: {
        outcome: P_U_A_M, transactionId: UUIDS[3], operationCount: 0, operations: [],
        verification: "not_attempted", rollback: "not_attempted",
      },
      stored: [`${UUIDS[3]}:in_flight`],
      markers: 0,
    });
  });
});

describe("R4j J3 parity pins (identical at the base and at HEAD)", () => {
  it("JS-p1 a matching-id finish whose write fails: the remembered report deep-equals the caller's report with outcome p_u_a_m", async () => {
    class FailingWrites extends CountingStorage {
      failWrites = false;
      override setItem(key: string, value: string): void {
        if (this.failWrites) throw new Error("quota");
        super.setItem(key, value);
      }
    }
    const storage = new FailingWrites();
    const store = new BrowserIdempotencyStore(storage, { clock: () => 1000 });
    const fence = await claimed(store);
    storage.failWrites = true;
    const report = committedReport(UUIDS[3]);
    expect(await store.finish(token(2), UUIDS[3], DIGEST, report, fence)).toBe(false);
    storage.failWrites = false;
    const again = await reclaim(store, storage);
    expect(again).toEqual({
      kind: P_U_A_M,
      report: { ...report, outcome: P_U_A_M },
      stored: [`${UUIDS[3]}:in_flight`],
      markers: 0,
    });
  });

  it("JS-p2 no owned claim (a second store on the same storage finishes a boundary it never claimed): the base content, the caller's id", async () => {
    const storage = new CountingStorage();
    const owner = new BrowserIdempotencyStore(storage, { clock: () => 1000 });
    await claimed(owner);
    const other = new BrowserIdempotencyStore(storage, { clock: () => 1000 });
    const report = committedReport(UUIDS[4]);
    expect(await other.finish(token(2), UUIDS[3], DIGEST, report, UUIDS[7])).toBe(false);
    const again = await reclaim(other, storage);
    expect(again).toEqual({ kind: P_U_A_M, report: { ...report, outcome: P_U_A_M }, stored: [`${UUIDS[3]}:in_flight`], markers: 0 });
    expect(again.report?.transactionId).toBe(UUIDS[4]);
  });

  it("JS-p3 a successful finish clears the remembered failure: refused finish, then the right report, then the claim replays terminal", async () => {
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => 1000 });
    const fence = await claimed(store);
    expect(await store.finish(token(2), UUIDS[3], DIGEST, committedReport(UUIDS[4]), fence)).toBe(false);
    expect(await store.finish(token(2), UUIDS[3], DIGEST, committedReport(UUIDS[3]), fence)).toBe(true);
    const again = await reclaim(store, storage);
    expect(again).toEqual({ kind: "terminal", report: committedReport(UUIDS[3]), stored: [`${UUIDS[3]}:terminal`], markers: 0 });
  });
});
