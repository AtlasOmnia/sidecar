import { describe, expect, it } from "vitest";
import {
  BrowserIdempotencyStore,
  type IdempotencyStorage,
  type WorkbookPlanExecutionReport,
} from "../workbookPlanExecutor";

/**
 * R4e-2 folded store items (dev3-r4e2-release "OMH DECISIONS + RELEASE 2026-10-07").
 * - L-A: finish refuses an invalid clock with zero storage reads and writes (R4e lane A A-2: the
 *   guard-order pin that S1/S2 did not assert).
 * - L-B A3: finish refuses a report whose transactionId does not match the claimed record, fail
 *   closed; the record stays in_flight and the store (and the origin) stays usable.
 * Helpers are copied from workbook-plan-executor-r4e.test.ts (nothing is imported from other tests).
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

describe("R4e-2 L-A: finish clock guard runs before any storage access", () => {
  it("LA1 [pin] an invalid finish clock (NaN, +/-Infinity, -1, -1e-9) returns false with zero storage reads and zero writes; the record stays in_flight", async () => {
    const observed: Record<string, unknown> = {};
    const expected: Record<string, unknown> = {};
    for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, -1, -1e-9]) {
      let now = 0;
      const storage = new CountingStorage();
      const store = new BrowserIdempotencyStore(storage, { clock: () => now });
      const claim = await store.claim(token(2), UUIDS[3], "digest-la1");
      if (claim.kind !== "claimed") throw new Error("expected claimed");
      const before = { reads: storage.reads, writes: storage.writes };
      now = bad;
      const finished = await store.finish(token(2), UUIDS[3], "digest-la1", committedReport(UUIDS[3]), claim.fence);
      observed[String(bad)] = {
        finished,
        reads: storage.reads - before.reads,
        writes: storage.writes - before.writes,
        states: storage.records().map((record) => record.state),
      };
      expected[String(bad)] = { finished: false, reads: 0, writes: 0, states: ["in_flight"] };
    }
    expect(observed).toEqual(expected);
  });
});

describe("R4e-2 L-B A3: finish binds the report's transactionId to the claimed record", () => {
  it("LB3 a report with a foreign transactionId is refused: false, record stays in_flight, the op answers p_u_a_m, other ops and a fresh store still claim, no marker", async () => {
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => 1000 });
    const claim = await store.claim(token(2), UUIDS[3], "digest-lb3");
    if (claim.kind !== "claimed") throw new Error("expected claimed");
    const finished = await store.finish(token(2), UUIDS[3], "digest-lb3", committedReport(UUIDS[4]), claim.fence);
    const again = await store.claim(token(2), UUIDS[3], "digest-lb3");
    const other = await store.claim(token(5), UUIDS[5], "digest-lb3-other");
    const fresh = new BrowserIdempotencyStore(storage, { clock: () => 2000 });
    const freshClaim = await fresh.claim(token(6), UUIDS[6], "digest-lb3-fresh");
    const markers = await fresh.reconciliationRequiredCount();
    expect({
      finished,
      again: again.kind,
      other: other.kind,
      freshClaim: freshClaim.kind,
      markers,
      stored: storage.records().map((record) => `${String(record.transactionId)}:${String(record.state)}`),
    }).toEqual({
      finished: false,
      again: "persistence_unavailable_after_mutation",
      other: "claimed",
      freshClaim: "claimed",
      markers: 0,
      stored: [`${UUIDS[3]}:in_flight`, `${UUIDS[5]}:in_flight`, `${UUIDS[6]}:in_flight`],
    });
  });

  it("LB3b [pin] the matching transactionId still finishes: true, stored terminal, and a fresh store replays it as terminal", async () => {
    const storage = new CountingStorage();
    const store = new BrowserIdempotencyStore(storage, { clock: () => 1000 });
    const claim = await store.claim(token(2), UUIDS[3], "digest-lb3b");
    if (claim.kind !== "claimed") throw new Error("expected claimed");
    const finished = await store.finish(token(2), UUIDS[3], "digest-lb3b", committedReport(UUIDS[3]), claim.fence);
    const replay = await new BrowserIdempotencyStore(storage, { clock: () => 2000 }).claim(token(2), UUIDS[3], "digest-lb3b");
    expect({ finished, state: storage.records()[0]?.state, replay: replay.kind }).toEqual({ finished: true, state: "terminal", replay: "terminal" });
  });
});
