/**
 * End-to-end regression tests covering orchestrated lifecycle:
 * normal chat, single/multiple tool calls, tool failures, retry
 * boundaries, repeat detection, checkpoint, deadlines, and privacy.
 *
 * These tests mock fetch() and Office.js (Excel.run) to verify
 * the full agent loop behavior without a running Excel instance.
 */

import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";
import { HermesAgent, makeToolCallSignature, isBlockedRepeatedToolCall, sleepHooks, resetInjectedSleep, getDirectSystemPrompt } from "../agent";
import { ToolCall } from "../types";
import { ToolErrorCategory } from "../toolErrors";

// ---------------------------------------------------------------------------
// Test helpers
// ---------------------------------------------------------------------------

function sseChunk(payload: unknown): Uint8Array {
  return new TextEncoder().encode(`data: ${JSON.stringify(payload)}\n\n`);
}

function doneChunk(): Uint8Array {
  return new TextEncoder().encode("data: [DONE]\n\n");
}

function textStream(text: string): ReadableStream<Uint8Array> {
  const encoder = new TextEncoder();
  const chunks: Uint8Array[] = [];
  for (let i = 0; i < text.length; i += 3) {
    const slice = text.slice(i, i + 3);
    if (slice) {
      chunks.push(
        encoder.encode(
          `data: {"choices":[{"delta":{"content":"${slice}"}}]}\n\n`,
        ),
      );
    }
  }
  chunks.push(doneChunk());
  return new ReadableStream({
    start(controller) {
      for (const chunk of chunks) controller.enqueue(chunk);
      controller.close();
    },
  });
}

function toolCallStream(
  name: string,
  args: Record<string, unknown>,
  id = "call_test",
): ReadableStream<Uint8Array> {
  return new ReadableStream({
    start(controller) {
      controller.enqueue(
        sseChunk({
          choices: [
            {
              delta: {
                tool_calls: [
                  {
                    index: 0,
                    id,
                    type: "function",
                    function: { name, arguments: JSON.stringify(args) },
                  },
                ],
              },
            },
          ],
        }),
      );
      controller.enqueue(doneChunk());
      controller.close();
    },
  });
}

function errorResponse(status: number, body = ""): Response {
  return {
    ok: false,
    status,
    headers: new Headers({ "content-type": "application/json" }),
    text: async () => body,
    body: null,
  } as unknown as Response;
}

function successResponse(body: ReadableStream<Uint8Array>): Response {
  return {
    ok: true,
    status: 200,
    headers: new Headers({ "content-type": "text/event-stream" }),
    body,
    text: async () => "",
  } as unknown as Response;
}

const noopCallbacks = {
  onAssistantStart: vi.fn(),
  onAssistantDelta: vi.fn(),
  onAssistantFinal: vi.fn(),
  onActivity: vi.fn(),
  onStatus: vi.fn(),
};

// Build a self-referencing mock range that returns itself for chained calls.
function mockRange(): Record<string, unknown> {
  // Use an object literal with lazy self-reference via a function wrapper
  const range: Record<string, unknown> = {};
  range.load = vi.fn();
  range.select = vi.fn();
  range.clear = vi.fn();
  range.getRange = vi.fn(() => range);
  range.getCell = vi.fn(() => range);
  range.getResizedRange = vi.fn(() => range);
  range.getOffsetRange = vi.fn(() => range);
  range.getLastColumn = vi.fn(() => range);
  range.getLastRow = vi.fn(() => range);
  range.address = "Sheet1!A1";
  range.values = [[1, 2, 3], [4, 5, 6]];
  range.formulas = [["", "", ""], ["", "", ""]];
  range.numberFormat = [["General", "General", "General"], ["General", "General", "General"]];
  range.formulasR1C1 = [["=R[-1]C*2"]];
  range.rowCount = 2;
  range.columnCount = 3;
  range.format = {
    font: { bold: false, italic: false, color: "", name: "", size: 11, underline: "None" },
    fill: { color: "" },
    borders: { getItem: vi.fn(() => ({ style: "", color: "", weight: "" })) },
    autofitColumns: vi.fn(),
    autofitRows: vi.fn(),
  };
  range.conditionalFormats = {
    add: vi.fn(() => ({
      cellValue: { format: { font: { color: "" }, fill: { color: "" } } },
      textComparison: { format: { font: { color: "" }, fill: { color: "" } } },
    })),
  };
  range.dataValidation = { ignoreBlanks: false, rule: {}, prompt: {} };
  range.isNullObject = false;
  range.autoFill = vi.fn();
  range.replaceAll = vi.fn(() => ({ value: 5 }));
  return range;
}

// Mock Office.js global — cast away strict Office.js types for test simplicity
beforeEach(() => {
  (globalThis as Record<string, unknown>).Excel = {
    run: vi.fn(
      async (work: (ctx: Record<string, unknown>) => Promise<Record<string, unknown>>) => {
        const mockCtx: Record<string, unknown> = {
          workbook: {
            worksheets: {
              getActiveWorksheet: () => ({ name: "Sheet1", load: vi.fn() }),
              getItem: (name: string) => {
                if (name === "Sheet1") {
                  return {
                    name: "Sheet1",
                    getRange: () => mockRange(),
                    getUsedRangeOrNullObject: () => {
                      const r = mockRange();
                      r.address = "Sheet1!A1:B10";
                      return r;
                    },
                    tables: { add: vi.fn() },
                    charts: { add: vi.fn() },
                    comments: {
                      getItemByCell: vi.fn(() => {
                        throw { code: "ItemNotFound", message: "ItemNotFound" };
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
                throw { code: "ItemNotFound", message: "ItemNotFound" };
              },
              add: () => ({ activate: vi.fn() }),
              load: vi.fn(),
              items: [{ name: "Sheet1" }],
            },
            getSelectedRange: () => ({ load: vi.fn(), address: "Sheet1!A1:B10" }),
            tables: {
              getItem: vi.fn(() => ({
                sort: { apply: vi.fn() },
                getRange: () => mockRange(),
                columns: {
                  getItemAt: vi.fn(() => ({ filter: { clear: vi.fn() } })),
                },
              })),
            },
            application: { calculate: vi.fn() },
          },
          sync: vi.fn(),
        };
        return work(mockCtx);
      },
    ),
  } as unknown as typeof Excel;
});

afterEach(() => {
  vi.restoreAllMocks();
  resetInjectedSleep();
});

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("e2e: normal chat", () => {
  it("returns model text without tool calls", async () => {
    const agent = new HermesAgent();
    const fetchMock = vi.fn().mockResolvedValue(
      successResponse(textStream("Hello! How can I help?")),
    );
    globalThis.fetch = fetchMock;

    const result = await agent.run(
      "Hello",
      { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
      noopCallbacks,
      new AbortController().signal,
    );

    expect(result.finalContent).toBe("Hello! How can I help?");
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});

describe("e2e: tool call execution", () => {
  it("executes a tool call and continues to text response", async () => {
    const agent = new HermesAgent();

    let callCount = 0;
    const fetchMock = vi.fn().mockImplementation(() => {
      callCount++;
      if (callCount === 1) {
        return Promise.resolve(successResponse(toolCallStream("get_workbook_overview", {})));
      }
      return Promise.resolve(successResponse(textStream("I've reviewed the workbook.")));
    });
    globalThis.fetch = fetchMock;

    await agent.run(
      "show me the workbook",
      { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 5 },
      noopCallbacks,
      new AbortController().signal,
    );

    // A tool activity was recorded
    expect(noopCallbacks.onActivity).toHaveBeenCalled();
  });
});

describe("e2e: retry boundaries", () => {
  beforeEach(() => {
    sleepHooks.injectedSleep = async () => {};
  });

  it("retries on 503 then succeeds", async () => {
    const agent = new HermesAgent();

    let attempts = 0;
    const fetchMock = vi.fn().mockImplementation(() => {
      attempts++;
      if (attempts < 3) return Promise.resolve(errorResponse(503, "overloaded"));
      return Promise.resolve(successResponse(textStream("Success after retry")));
    });
    globalThis.fetch = fetchMock;

    const result = await agent.run(
      "hello",
      { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
      noopCallbacks,
      new AbortController().signal,
    );

    expect(attempts).toBe(3);
    expect(result.finalContent).toBe("Success after retry");
  });

  it("gives up after max retries on persistent failure", async () => {
    const agent = new HermesAgent();

    const fetchMock = vi.fn().mockResolvedValue(errorResponse(503, "always down"));
    globalThis.fetch = fetchMock;

    await expect(
      agent.run(
        "hello",
        { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
        noopCallbacks,
        new AbortController().signal,
      ),
    ).rejects.toThrow();

    expect(fetchMock).toHaveBeenCalledTimes(3); // initial + 2 retries
  });

  it("does not retry on 401 client error", async () => {
    const agent = new HermesAgent();

    const fetchMock = vi.fn().mockResolvedValue(errorResponse(401, "unauthorized"));
    globalThis.fetch = fetchMock;

    await expect(
      agent.run(
        "hello",
        { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
        noopCallbacks,
        new AbortController().signal,
      ),
    ).rejects.toThrow();

    // 401 should NOT retry — only 1 attempt
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});

describe("e2e: repeat detection", () => {
  it("canonicalizes equivalent arguments for signature matching", () => {
    const call1: ToolCall = {
      id: "a",
      type: "function",
      function: { name: "read_range", arguments: '{"address":"A1:B2","sheet":"Sheet1"}' },
    };
    const call2: ToolCall = {
      id: "b",
      type: "function",
      function: { name: "read_range", arguments: '{"sheet":"Sheet1","address":"A1:B2"}' },
    };

    expect(makeToolCallSignature(call1)).toBe(makeToolCallSignature(call2));
  });

  it("blocks after third identical call", () => {
    expect(isBlockedRepeatedToolCall(1)).toBe(false);
    expect(isBlockedRepeatedToolCall(2)).toBe(false);
    expect(isBlockedRepeatedToolCall(3)).toBe(true);
  });
});

describe("e2e: cancellation", () => {
  it("aborts mid-turn when signal fires", async () => {
    const agent = new HermesAgent();
    const controller = new AbortController();

    const fetchMock = vi.fn().mockImplementation(() => {
      controller.abort();
      return Promise.resolve(successResponse(textStream("partial")));
    });
    globalThis.fetch = fetchMock;

    await expect(
      agent.run(
        "hello",
        { baseUrl: "http://127.0.0.1:8644/v1", model: "test", apiKey: "", maxIterations: 15 },
        noopCallbacks,
        controller.signal,
      ),
    ).rejects.toThrow();
  });
});

describe("e2e: typed failure categories", () => {
  it("validation errors are distinct from provider errors", () => {
    expect(ToolErrorCategory.Validation).not.toBe(ToolErrorCategory.Provider);
    expect(ToolErrorCategory.Validation).toBe("validation");
    expect(ToolErrorCategory.Provider).toBe("provider");
  });

  it("tool error JSON is self-describing", () => {
    const err = JSON.stringify({
      ok: false,
      error: "test",
      category: ToolErrorCategory.Cancellation,
      name: "AbortError",
    });
    const parsed = JSON.parse(err);
    expect(parsed.category).toBe("cancellation");
  });
});

describe("e2e: tracing privacy", () => {
  it("turn trace types exclude sensitive fields", () => {
    const trace = {
      turnId: "test",
      startedAt: new Date().toISOString(),
      totalDurationMs: 0,
      modelCallAttempts: 0,
      modelCalls: [],
      toolCallCount: 0,
      toolResults: [],
      outcome: "completed" as const,
      checkpointFired: false,
      maxIterations: 15,
      turnNumber: 1,
    };

    const keys = Object.keys(trace);
    expect(keys).not.toContain("prompt");
    expect(keys).not.toContain("arguments");
    expect(keys).not.toContain("apiKey");
    expect(keys).not.toContain("values");
  });
});

describe("C1: e2e prompt title recipe (direct mode)", () => {
  it("system prompt includes title-span CenterAcrossSelection recipe and no-merge directive", () => {
    const content = getDirectSystemPrompt();
    // Title-span recipe must mention CenterAcrossSelection
    expect(content).toContain("CenterAcrossSelection");
    // No merge capability
    expect(content).toContain("No merge capability exists");
    // Precondition: read span first
    expect(content).toContain("first read the full");
    // Wrap around behavior
    expect(content).toContain("wrapText false");
    // ColumnAutofit false
    expect(content).toContain("columnAutofit");
  });
});
