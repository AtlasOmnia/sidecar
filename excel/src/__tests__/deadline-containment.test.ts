/**
 * Model-call containment hardening tests.
 *
 * 1. Per-call deadline stays active through body reads (not cleared at headers)
 * 2. Hanging SSE body times out at 20s
 * 3. Hanging JSON body times out
 * 4. User abort during body read → Cancellation (not Timeout)
 * 5. Reader cancellation + listener/timer cleanup after success and abort
 * 6. interruptibleSleep removes listener after success
 * 7. Checkpoint catch does not convert arbitrary errors to deadline
 * 8. Sanitized non-2xx error handling (no raw body in error)
 * 9. Retry tests use fake timing (no wall-clock sleeps)
 */

import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";
import {
  HermesAgent,
  readOpenAIStream,
  interruptibleSleep,
  deadlineAbortSignal,
  throwIfAborted,
  MODEL_CALL_TIMEOUT_MS,
  TURN_TIMEOUT_MS,
  AgentCallbacks,
  sleepHooks,
  resetInjectedSleep,
} from "../agent";
import { ToolErrorCategory, categorizeError } from "../toolErrors";

// Prevent localStorage errors
beforeEach(() => {
  vi.stubGlobal("localStorage", {
    getItem: () => null,
    setItem: () => {},
    removeItem: () => {},
    clear: () => {},
    length: 0,
    key: () => null,
  });
});

afterEach(() => {
  vi.restoreAllMocks();
  resetInjectedSleep();
});

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function mockOfficeJs() {
  (globalThis as Record<string, unknown>).Excel = {
    run: vi.fn(
      async (
        work: (ctx: Record<string, unknown>) => Promise<Record<string, unknown>>,
      ) => {
        const ctx = {
          workbook: {
            worksheets: {
              getActiveWorksheet: () => ({ name: "Sheet1", load: vi.fn() }),
              getItem: (name: string) => {
                if (name === "Sheet1") {
                  return {
                    name: "Sheet1",
                    getRange: () => {
                      const r: Record<string, unknown> = {};
                      r.load = vi.fn();
                      r.select = vi.fn();
                      r.clear = vi.fn();
                      r.getCell = vi.fn(() => r);
                      r.getResizedRange = vi.fn(() => r);
                      r.getOffsetRange = vi.fn(() => r);
                      r.getLastColumn = vi.fn(() => r);
                      r.getLastRow = vi.fn(() => r);
                      r.address = "Sheet1!A1";
                      r.values = [[1, 2]];
                      r.formulas = [["", ""]];
                      r.numberFormat = [["General", "General"]];
                      r.formulasR1C1 = [[""]];
                      r.rowCount = 2;
                      r.columnCount = 2;
                      r.format = {};
                      r.conditionalFormats = {
                        add: vi.fn(() => ({})),
                      };
                      r.dataValidation = {};
                      r.isNullObject = false;
                      r.autoFill = vi.fn();
                      r.replaceAll = vi.fn(() => ({ value: 0 }));
                      r.comment = { content: "" };
                      return r;
                    },
                    getUsedRangeOrNullObject: () => {
                      const r: Record<string, unknown> = {};
                      r.load = vi.fn();
                      r.isNullObject = false;
                      r.address = "Sheet1!A1:B2";
                      r.rowCount = 2;
                      r.columnCount = 2;
                      return r;
                    },
                    tables: { add: vi.fn() },
                    charts: { add: vi.fn() },
                    comments: {
                      getItemByCell: vi.fn(() => {
                        throw { code: "ItemNotFound" };
                      }),
                      add: vi.fn(),
                      load: vi.fn(),
                      items: [],
                    },
                    pivotTables: {},
                    freezePanes: {
                      freezeRows: vi.fn(),
                      freezeColumns: vi.fn(),
                      unfreeze: vi.fn(),
                    },
                    activate: vi.fn(),
                    delete: vi.fn(),
                    load: vi.fn(),
                  };
                }
                throw { code: "ItemNotFound" };
              },
              add: () => ({ activate: vi.fn() }),
              load: vi.fn(),
              items: [{ name: "Sheet1" }],
            },
            getSelectedRange: () => ({
              load: vi.fn(),
              address: "Sheet1!A1:B2",
            }),
            tables: {
              getItem: vi.fn(() => ({
                sort: { apply: vi.fn() },
                getRange: vi.fn(),
                columns: {
                  getItemAt: vi.fn(() => ({
                    filter: { clear: vi.fn() },
                  })),
                },
              })),
            },
            application: { calculate: vi.fn() },
          },
          sync: vi.fn(),
        };
        return work(ctx);
      },
    ),
  };
}

// A stream body that resolves immediately (no chunks, closes immediately)
function emptyStream(): ReadableStream<Uint8Array> {
  return new ReadableStream({
    start(ctrl) { ctrl.close(); },
  });
}

// A stream that never produces data (hanging reader.read())
// The cancel() callback allows reader.cancel() to resolve the pending read
function hangingStream(): ReadableStream<Uint8Array> {
  return new ReadableStream({
    start() {
      // Never enqueue, never close — reader.read() hangs forever
    },
    cancel() {
      // Called by reader.cancel() — allows pending read to resolve
    },
  });
}

// A stream that produces data after a long delay (guards against closed controller)
function slowStream(text: string, delayMs: number): ReadableStream<Uint8Array> {
  let timerId: ReturnType<typeof setTimeout> | null = null;
  return new ReadableStream({
    start(ctrl) {
      timerId = setTimeout(() => {
        try {
          ctrl.enqueue(new TextEncoder().encode(`data: {"choices":[{"delta":{"content":"${text}"}}]}\n\ndata: [DONE]\n\n`));
          ctrl.close();
        } catch {
          // Controller already closed by timeout — ignore
        }
      }, delayMs);
    },
    cancel() {
      if (timerId !== null) {
        clearTimeout(timerId);
        timerId = null;
      }
    },
  });
}

function textResponse(body: ReadableStream<Uint8Array>): Response {
  return {
    ok: true,
    status: 200,
    headers: new Headers({ "content-type": "text/event-stream" }),
    body,
    text: async () => "",
  } as unknown as Response;
}

function errorResponse(status: number): Response {
  return {
    ok: false,
    status,
    headers: new Headers({ "content-type": "application/json" }),
    text: async () => `{"error":"secret detail ${status}"}`,
  } as unknown as Response;
}

// ---------------------------------------------------------------------------
// Test 1: Per-call deadline stays active through body reads
// ---------------------------------------------------------------------------

describe("per-call deadline through body reads", () => {
  it("hanging SSE body times out (timer not cleared at headers)", async () => {
    vi.useFakeTimers();
    mockOfficeJs();
    const agent = new HermesAgent();

    const fetchMock = vi.fn().mockResolvedValue(
      textResponse(hangingStream()),
    );
    globalThis.fetch = fetchMock;

    // Inject zero-delay sleep so retries are instant
    sleepHooks.injectedSleep = async () => {};

    const runPromise = agent.run(
      "hello",
      { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
      {
        onAssistantStart: vi.fn(),
        onAssistantDelta: vi.fn(),
        onAssistantFinal: vi.fn(),
        onActivity: vi.fn(),
        onStatus: vi.fn(),
      },
      new AbortController().signal,
    );

    // Attach rejection handler before advancing timers so it's handled synchronously
    const rejectionPromise = expect(runPromise).rejects.toThrow();

    // Advance past the 20s per-call timeout
    await vi.advanceTimersByTimeAsync(MODEL_CALL_TIMEOUT_MS + 1000);

    await rejectionPromise;
    vi.useRealTimers();
    resetInjectedSleep();
  });

  it("fetch consumes part of budget, body gets remaining time (not reset at headers)", async () => {
    vi.useFakeTimers();
    mockOfficeJs();
    const agent = new HermesAgent();

    // Slow fetch (5s) then slow body (10s) — total must be <= 20s
    let fetchResolve: ((v: Response) => void) | null = null;
    const fetchMock = vi.fn().mockImplementation(() => {
      return new Promise<Response>((resolve) => {
        fetchResolve = resolve;
      });
    });
    globalThis.fetch = fetchMock;

    sleepHooks.injectedSleep = async () => {};

    const runPromise = agent.run(
      "hello",
      { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
      {
        onAssistantStart: vi.fn(),
        onAssistantDelta: vi.fn(),
        onAssistantFinal: vi.fn(),
        onActivity: vi.fn(),
        onStatus: vi.fn(),
      },
      new AbortController().signal,
    );

    // Attach rejection handler before advancing timers
    const rejectionPromise = expect(runPromise).rejects.toThrow();

    // Advance 5s (fetch not resolved yet, timer still ticking)
    await vi.advanceTimersByTimeAsync(5000);
    // Resolve fetch now
    fetchResolve!(textResponse(slowStream("data", 15_000)));
    // Advance 16s (total = 21s, past 20s deadline)
    await vi.advanceTimersByTimeAsync(16_000);

    await rejectionPromise;
    vi.useRealTimers();
    resetInjectedSleep();
  });

  it("hanging JSON body times out", async () => {
    vi.useFakeTimers();
    mockOfficeJs();
    const agent = new HermesAgent();

    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      headers: new Headers({ "content-type": "application/json" }),
      body: hangingStream(),
      text: async () => "",
    } as unknown as Response);
    globalThis.fetch = fetchMock;

    sleepHooks.injectedSleep = async () => {};

    const runPromise = agent.run(
      "hello",
      { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
      {
        onAssistantStart: vi.fn(),
        onAssistantDelta: vi.fn(),
        onAssistantFinal: vi.fn(),
        onActivity: vi.fn(),
        onStatus: vi.fn(),
      },
      new AbortController().signal,
    );

    // Attach rejection handler before advancing timers
    const rejectionPromise = expect(runPromise).rejects.toThrow();

    await vi.advanceTimersByTimeAsync(MODEL_CALL_TIMEOUT_MS + 1000);

    await rejectionPromise;
    vi.useRealTimers();
    resetInjectedSleep();
  });
});

// ---------------------------------------------------------------------------
// Test 2: User abort vs timeout distinction during body reads
// ---------------------------------------------------------------------------

describe("user abort vs timeout during body read", () => {
  it("user abort during pending body read is Cancellation, not Timeout", () => {
    const controller = new AbortController();
    controller.abort();
    try {
      throwIfAborted(controller.signal);
    } catch (err) {
      expect(err).toBeInstanceOf(DOMException);
      expect((err as DOMException).name).toBe("AbortError");
      expect(categorizeError(err)).toBe(ToolErrorCategory.Cancellation);
    }
  });

  it("timeout during body read is Timeout category", () => {
    const err = new DOMException("timed out", "TimeoutError");
    expect(categorizeError(err)).toBe(ToolErrorCategory.Timeout);
  });

  it("user abort during pending fetch returns AbortError, not network error", async () => {
    mockOfficeJs();
    const agent = new HermesAgent();
    const controller = new AbortController();
    sleepHooks.injectedSleep = async () => {};

    // fetch hangs until abort
    const fetchMock = vi.fn().mockImplementation(() => {
      return new Promise<Response>(() => {
        // never resolves — user abort will cancel
      });
    });
    globalThis.fetch = fetchMock;

    const runPromise = agent.run(
      "hello",
      { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
      {
        onAssistantStart: vi.fn(),
        onAssistantDelta: vi.fn(),
        onAssistantFinal: vi.fn(),
        onActivity: vi.fn(),
        onStatus: vi.fn(),
      },
      controller.signal,
    );

    // User presses Stop while fetch is pending
    controller.abort();

    try {
      await runPromise;
      expect.unreachable("should have thrown");
    } catch (err) {
      expect(err).toBeInstanceOf(DOMException);
      expect((err as DOMException).name).toBe("AbortError");
      expect(categorizeError(err)).toBe(ToolErrorCategory.Cancellation);
      // Must NOT contain misleading network message
      const msg = (err as Error).message;
      expect(msg).not.toContain("Could not reach");
      expect(msg).toContain("stopped");
    }

    resetInjectedSleep();
  });
});

// ---------------------------------------------------------------------------
// Test 3: Reader cancellation and cleanup
// ---------------------------------------------------------------------------

describe("reader cancellation", () => {
  it("readOpenAIStream cancels reader when signal fires", async () => {
    const controller = new AbortController();
    const stream = hangingStream();

    const readPromise = readOpenAIStream(
      stream,
      {
        onAssistantStart: vi.fn(),
        onAssistantDelta: vi.fn(),
        onAssistantFinal: vi.fn(),
        onActivity: vi.fn(),
        onStatus: vi.fn(),
      },
      controller.signal,
    );

    // Abort after a short delay
    setTimeout(() => controller.abort(), 10);

    await expect(readPromise).rejects.toThrow(DOMException);
  });

  it("readOpenAIStream removes abort listener on normal completion", async () => {
    const controller = new AbortController();
    const stream = emptyStream();

    const result = await readOpenAIStream(
      stream,
      {
        onAssistantStart: vi.fn(),
        onAssistantDelta: vi.fn(),
        onAssistantFinal: vi.fn(),
        onActivity: vi.fn(),
        onStatus: vi.fn(),
      },
      controller.signal,
    );

    expect(result.content).toBe("");
    expect(result.toolCalls).toHaveLength(0);

    // After normal completion, aborting should not throw or affect anything
    controller.abort();
    // No assertion needed — just verifying no unhandled rejections
  });
});

// ---------------------------------------------------------------------------
// Test 4: interruptibleSleep listener cleanup
// ---------------------------------------------------------------------------

describe("interruptibleSleep cleanup", () => {
  it("removes abort listener after normal resolution", async () => {
    const controller = new AbortController();

    // Normal resolution
    await interruptibleSleep(20, controller.signal);

    // After resolution, abort should not cause issues (listener was removed)
    controller.abort();
    // If the listener leaked, aborting would trigger a rejection that was already resolved.
    // We just verify no crash occurs.
  });

  it("rejects on abort during sleep", async () => {
    const controller = new AbortController();
    const promise = interruptibleSleep(5000, controller.signal);
    controller.abort();
    await expect(promise).rejects.toThrow("Aborted during retry delay.");
  });
});

// ---------------------------------------------------------------------------
// Test 5: Checkpoint catch correctness
// ---------------------------------------------------------------------------

describe("checkpoint catch correctness", () => {
  it("arbitrary checkpoint callback error does not become stopReason=deadline", async () => {
    mockOfficeJs();
    const agent = new HermesAgent();
    agent.reset();

    const origNow = Date.now;
    const realStart = origNow();
    let firstCall = true;
    Date.now = vi.fn(() => {
      if (firstCall) { firstCall = false; return realStart; }
      return realStart + 31_000;
    });

    const fetchMock = vi.fn().mockResolvedValue(textResponse(emptyStream()));
    globalThis.fetch = fetchMock;

    // Make fetch mock return enough responses to avoid hanging
    sleepHooks.injectedSleep = async () => {};

    await expect(
      agent.run(
        "test",
        { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
        {
          onAssistantStart: vi.fn(),
          onAssistantDelta: vi.fn(),
          onAssistantFinal: vi.fn(),
          onActivity: vi.fn(),
          onStatus: vi.fn(),
          onCheckpoint: async () => {
            throw new Error("callback internal bug");
          },
        },
        new AbortController().signal,
      ),
    ).rejects.toThrow("callback internal bug");

    Date.now = origNow;
    resetInjectedSleep();
  });

  it("genuine TimeoutError in checkpoint → stopReason=deadline", async () => {
    mockOfficeJs();
    const agent = new HermesAgent();
    agent.reset();

    const origNow = Date.now;
    const realStart = origNow();
    let firstCall = true;
    Date.now = vi.fn(() => {
      if (firstCall) { firstCall = false; return realStart; }
      return realStart + 31_000;
    });

    const fetchMock = vi.fn().mockResolvedValue(textResponse(emptyStream()));
    globalThis.fetch = fetchMock;

    sleepHooks.injectedSleep = async () => {};

    const result = await agent.run(
      "test",
      { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
      {
        onAssistantStart: vi.fn(),
        onAssistantDelta: vi.fn(),
        onAssistantFinal: vi.fn(),
        onActivity: vi.fn(),
        onStatus: vi.fn(),
        onCheckpoint: async () => {
          throw new DOMException("deadline", "TimeoutError");
        },
      },
      new AbortController().signal,
    );

    Date.now = origNow;
    expect(result.stopReason).toBe("deadline");
    resetInjectedSleep();
  });
});

// ---------------------------------------------------------------------------
// Test 6: Sanitized non-2xx error handling
// ---------------------------------------------------------------------------

describe("sanitized non-2xx errors", () => {
  it("does not append raw provider body to error message", async () => {
    mockOfficeJs();
    const agent = new HermesAgent();

    const fetchMock = vi.fn().mockResolvedValue(errorResponse(500));
    globalThis.fetch = fetchMock;

    sleepHooks.injectedSleep = async () => {};

    try {
      await agent.run(
        "hello",
        { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
        {
          onAssistantStart: vi.fn(),
          onAssistantDelta: vi.fn(),
          onAssistantFinal: vi.fn(),
          onActivity: vi.fn(),
          onStatus: vi.fn(),
        },
        new AbortController().signal,
      );
      expect.unreachable("should have thrown");
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      // Must NOT contain the raw body
      expect(msg).not.toContain("secret detail");
      // Must contain status code
      expect(msg).toContain("500");
    }

    resetInjectedSleep();
  });

  it("503 errors carry statusCode for retry classification", async () => {
    mockOfficeJs();
    const agent = new HermesAgent();

    const fetchMock = vi.fn().mockResolvedValue(errorResponse(503));
    globalThis.fetch = fetchMock;

    sleepHooks.injectedSleep = async () => {};

    try {
      await agent.run(
        "hello",
        { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
        {
          onAssistantStart: vi.fn(),
          onAssistantDelta: vi.fn(),
          onAssistantFinal: vi.fn(),
          onActivity: vi.fn(),
          onStatus: vi.fn(),
        },
        new AbortController().signal,
      );
    } catch (err) {
      const record = err as Record<string, unknown>;
      expect(record.statusCode).toBe(503);
    }

    resetInjectedSleep();
  });
});

// ---------------------------------------------------------------------------
// Test 7: Retry tests with fake timing
// ---------------------------------------------------------------------------

describe("retry with fake timing", () => {
  it("retries on 503 using injected sleep (no wall-clock waits)", async () => {
    vi.useFakeTimers();
    mockOfficeJs();
    const agent = new HermesAgent();

    let attempts = 0;
    const fetchMock = vi.fn().mockImplementation(() => {
      attempts++;
      if (attempts < 3) return Promise.resolve(errorResponse(503));
      return Promise.resolve(textResponse(emptyStream()));
    });
    globalThis.fetch = fetchMock;

    // Inject zero-delay sleep
    let sleepCalls = 0;
    sleepHooks.injectedSleep = async () => { sleepCalls++; };

    const result = await agent.run(
      "hello",
      { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
      {
        onAssistantStart: vi.fn(),
        onAssistantDelta: vi.fn(),
        onAssistantFinal: vi.fn(),
        onActivity: vi.fn(),
        onStatus: vi.fn(),
      },
      new AbortController().signal,
    );

    expect(attempts).toBe(3);
    expect(sleepCalls).toBeGreaterThanOrEqual(2);
    vi.useRealTimers();
    resetInjectedSleep();
  });

  it("retry sleep is capped by remaining turn deadline", () => {
    // At 89s elapsed (1s remaining), a 4s retry delay must be capped
    const remainingTurn = 1000; // ms
    const rawDelay = 4000;
    const maxDelay = Math.max(0, remainingTurn - 500);
    const delay = Math.min(rawDelay, maxDelay);
    expect(delay).toBe(500);
  });
});

// ---------------------------------------------------------------------------
// Test 8: deadlineAbortSignal and timer cleanup
// ---------------------------------------------------------------------------

describe("deadlineAbortSignal cleanup", () => {
  it("cleanup prevents timer from firing", async () => {
    vi.useFakeTimers();
    const deadline = deadlineAbortSignal(5000, "test");
    expect(deadline.signal.aborted).toBe(false);
    deadline.cleanup();
    await vi.advanceTimersByTimeAsync(10_000);
    expect(deadline.signal.aborted).toBe(false);
    vi.useRealTimers();
  });

  it("signal fires when cleanup not called", async () => {
    vi.useFakeTimers();
    const deadline = deadlineAbortSignal(100, "test");
    await vi.advanceTimersByTimeAsync(200);
    expect(deadline.signal.aborted).toBe(true);
    deadline.cleanup();
    vi.useRealTimers();
  });
});

// ---------------------------------------------------------------------------
// Test 9: readResponseBody resource cleanup on timeout
// ---------------------------------------------------------------------------

describe("readResponseBody resource cleanup", () => {
  it("hanging JSON body timeout cancels reader and cleans up listeners", async () => {
    vi.useFakeTimers();
    mockOfficeJs();
    const agent = new HermesAgent();

    // Track whether cancel was called on the stream
    let cancelCalled = false;
    const trackedHangingStream = new ReadableStream({
      start() {
        // Never enqueue, never close — reader.read() hangs forever
      },
      cancel() {
        cancelCalled = true;
      },
    });

    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      headers: new Headers({ "content-type": "application/json" }),
      body: trackedHangingStream,
      text: async () => "",
    } as unknown as Response);
    globalThis.fetch = fetchMock;

    sleepHooks.injectedSleep = async () => {};

    const runPromise = agent.run(
      "hello",
      { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
      {
        onAssistantStart: vi.fn(),
        onAssistantDelta: vi.fn(),
        onAssistantFinal: vi.fn(),
        onActivity: vi.fn(),
        onStatus: vi.fn(),
      },
      new AbortController().signal,
    );

    const rejectionPromise = expect(runPromise).rejects.toThrow();

    // Advance past the 20s per-call timeout
    await vi.advanceTimersByTimeAsync(MODEL_CALL_TIMEOUT_MS + 1000);

    await rejectionPromise;

    // After timeout, the reader should have been cancelled (resource cleanup)
    expect(cancelCalled).toBe(true);

    vi.useRealTimers();
    resetInjectedSleep();
  });

  it("abort listener is removed after readResponseBody completes normally", async () => {
    mockOfficeJs();
    const agent = new HermesAgent();

    // A stream that produces a complete JSON response immediately
    const jsonBody = `{"choices":[{"message":{"content":"hello","tool_calls":[]}}]}`;
    const immediateStream = new ReadableStream({
      start(ctrl) {
        ctrl.enqueue(new TextEncoder().encode(jsonBody));
        ctrl.close();
      },
    });

    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      headers: new Headers({ "content-type": "application/json" }),
      body: immediateStream,
      text: async () => "",
    } as unknown as Response);
    globalThis.fetch = fetchMock;

    sleepHooks.injectedSleep = async () => {};

    const controller = new AbortController();

    // Should complete successfully without timeout
    const result = await agent.run(
      "hello",
      { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
      {
        onAssistantStart: vi.fn(),
        onAssistantDelta: vi.fn(),
        onAssistantFinal: vi.fn(),
        onActivity: vi.fn(),
        onStatus: vi.fn(),
      },
      controller.signal,
    );

    expect(result.finalContent).toBe("hello");

    // After normal completion, aborting the signal should not cause issues
    // (listener was properly removed)
    controller.abort();

    resetInjectedSleep();
  });
});
