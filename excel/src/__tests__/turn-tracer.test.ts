import { describe, expect, it, beforeEach } from "vitest";
import {
  startTrace,
  recordModelCall,
  recordToolResult,
  finalizeTrace,
  loadTraces,
  clearTraces,
  TurnTrace,
} from "../turnTracer";
import { ToolErrorCategory } from "../toolErrors";

// Mock localStorage for the test environment
const store = new Map<string, string>();

beforeEach(() => {
  store.clear();
  // vitest jsdom provides localStorage; override with our mock store
  Object.defineProperty(globalThis, "localStorage", {
    value: {
      getItem: (key: string) => store.get(key) ?? null,
      setItem: (key: string, value: string) => { store.set(key, value); },
      removeItem: (key: string) => { store.delete(key); },
      clear: () => { store.clear(); },
      get length() { return store.size; },
      key: (index: number) => {
        const keys = Array.from(store.keys());
        return keys[index] ?? null;
      },
    },
    writable: true,
    configurable: true,
  });
});

describe("turnTracer", () => {
  it("starts a trace with a unique ID", () => {
    const trace = startTrace(15);
    expect(trace.turnId).toMatch(/^turn_\d+_[a-z0-9]{4}$/);
    expect(trace.startedAt).toBeTruthy();
    expect(trace.maxIterations).toBe(15);
    expect(trace.turnNumber).toBeGreaterThan(0);
  });

  it("increments turn numbers across traces", () => {
    const t1 = startTrace(15);
    const t2 = startTrace(10);
    expect(t2.turnNumber).toBe(t1.turnNumber + 1);
  });

  it("records model call metadata", () => {
    const trace = startTrace(15);
    recordModelCall(trace, 0, 1234, true, 200);
    recordModelCall(trace, 1, 567, false, 503);

    expect(trace.modelCallAttempts).toBe(2);
    expect(trace.modelCalls).toHaveLength(2);
    expect(trace.modelCalls[0]).toEqual({
      attempt: 0,
      durationMs: 1234,
      statusCode: 200,
      success: true,
    });
    expect(trace.modelCalls[1]).toEqual({
      attempt: 1,
      durationMs: 567,
      statusCode: 503,
      success: false,
    });
  });

  it("records tool result metadata", () => {
    const trace = startTrace(15);
    recordToolResult(trace, "read_range", 45, false);
    recordToolResult(trace, "write_range", 23, true, ToolErrorCategory.OfficeHost);
    recordToolResult(trace, "format_range", 12, false);

    expect(trace.toolCallCount).toBe(3);
    expect(trace.toolResults).toHaveLength(3);
    expect(trace.toolResults[0]).toEqual({
      name: "read_range",
      durationMs: 45,
      isError: false,
      errorCategory: undefined,
    });
    expect(trace.toolResults[1]).toEqual({
      name: "write_range",
      durationMs: 23,
      isError: true,
      errorCategory: ToolErrorCategory.OfficeHost,
    });
  });

  it("finalizes trace with total duration and persists", () => {
    const trace = startTrace(15);
    trace.outcome = "completed";
    finalizeTrace(trace);

    expect(trace.totalDurationMs).toBeGreaterThanOrEqual(0);

    const saved = loadTraces();
    expect(saved).toHaveLength(1);
    expect(saved[0].turnId).toBe(trace.turnId);
    expect(saved[0].outcome).toBe("completed");
  });

  it("enforces MAX_TRACES FIFO eviction", () => {
    // Create 55 traces (MAX_TRACES = 50)
    for (let i = 0; i < 55; i++) {
      const trace = startTrace(15);
      trace.outcome = "completed";
      finalizeTrace(trace);
    }

    const saved = loadTraces();
    expect(saved).toHaveLength(50);
    // First trace should have smaller turn number than last (FIFO eviction)
    expect(saved[0].turnNumber).toBeLessThan(saved[49].turnNumber);
    // 5 traces were evicted
    expect(saved[49].turnNumber - saved[0].turnNumber).toBe(49);
  });

  it("never stores sensitive content", () => {
    const trace = startTrace(15);
    finalizeTrace(trace);

    const saved = loadTraces()[0];
    const json = JSON.stringify(saved);

    // Verify no content-bearing fields exist
    expect(json).not.toContain("prompt");
    expect(json).not.toContain("workbook");
    expect(json).not.toContain("values");
    expect(json).not.toContain("apiKey");
    expect(json).not.toContain("api_key");
    expect(json).not.toContain("arguments");
    expect(json).not.toContain("response");
    expect(json).not.toContain("password");
  });

  it("clearTraces removes all persisted traces", () => {
    for (let i = 0; i < 3; i++) {
      const trace = startTrace(15);
      trace.outcome = "completed";
      finalizeTrace(trace);
    }
    expect(loadTraces()).toHaveLength(3);

    clearTraces();
    expect(loadTraces()).toHaveLength(0);
  });

  it("handles missing localStorage gracefully", () => {
    // @ts-expect-error - testing without localStorage
    globalThis.localStorage = undefined;
    const trace = startTrace(15);
    trace.outcome = "completed";
    // Should not throw
    expect(() => finalizeTrace(trace)).not.toThrow();
    expect(() => loadTraces()).not.toThrow();
    expect(loadTraces()).toEqual([]);
  });
});
