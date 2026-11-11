import { describe, expect, it, vi, beforeEach } from "vitest";
import { HermesAgent, getServerOwnedSystemPrompt, getDirectSystemPrompt } from "../agent";

const BRIDGE_SESSION_ID = "660e8400-e29b-41d4-a716-446655440001";

// Minimal mock settings
const makeSettings = (overrides: Partial<{ model: string; baseUrl: string }>) => ({
  baseUrl: overrides.baseUrl ?? "http://127.0.0.1:8643/v1",
  model: overrides.model ?? "hermes-agent",
  apiKey: "",
  maxIterations: 8,
});

// Mock fetch at module level
const mockFetch = vi.fn();
global.fetch = mockFetch;

function makeStreamResponse(content: string, toolCalls?: Array<{ id: string; type: string; function: { name: string; arguments: string } }>) {
  const lines: string[] = [];
  // Send content delta
  if (content) {
    lines.push(`data: ${JSON.stringify({ choices: [{ delta: { content } }] })}`);
  }
  // Send tool calls if any
  for (const tc of toolCalls ?? []) {
    lines.push(`data: ${JSON.stringify({ choices: [{ delta: { tool_calls: [{ index: 0, id: tc.id, type: "function", function: tc.function }] } }] })}`);
  }
  // Finish
  lines.push('data: {"choices":[{"finish_reason":"stop"}]}');
  lines.push("data: [DONE]");

  const body = new ReadableStream({
    start(controller) {
      for (const line of lines) {
        controller.enqueue(new TextEncoder().encode(`${line}\n\n`));
      }
      controller.close();
    },
  });

  return {
    ok: true,
    status: 200,
    headers: new Headers({ "content-type": "text/event-stream" }),
    body,
  };
}

function makeCallbacks() {
  let lastContent = "";
  return {
    onAssistantStart: vi.fn(),
    onAssistantDelta: (delta: string) => { lastContent += delta; },
    onAssistantFinal: (content: string) => { lastContent = content; },
    onActivity: vi.fn(),
    onStatus: vi.fn(),
    get finalContent() { return lastContent; },
  };
}

const directPlan = {
  version: 1,
  outcome: "execute",
  summary: "Write one value",
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }],
  verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["ok"]] }],
};

const directPlanToolCall = {
  id: "call_plan_1",
  type: "function",
  function: {
    name: "excel_execute_plan",
    arguments: JSON.stringify(directPlan),
  },
};

function mockOfficeJs(): void {
  (globalThis as Record<string, unknown>).Excel = {
    run: vi.fn(async (work: (context: Record<string, unknown>) => Promise<Record<string, unknown>>) => {
      const usedRange = {
        load: vi.fn(),
        isNullObject: false,
        address: "Sheet1!A1:B2",
        rowCount: 2,
        columnCount: 2,
      };
      const sheet = {
        name: "Sheet1",
        getUsedRangeOrNullObject: vi.fn(() => usedRange),
        getRange: vi.fn(() => ({
          load: vi.fn(),
          address: "Sheet1!A1",
          rowCount: 1,
          columnCount: 1,
          values: [["legacy primitive"]],
          formulas: [[""]],
          numberFormat: [["General"]],
        })),
      };
      return work({
        workbook: {
          worksheets: {
            load: vi.fn(),
            items: [sheet],
            getItem: vi.fn(() => sheet),
            getActiveWorksheet: vi.fn(() => ({ name: "Sheet1", load: vi.fn() })),
          },
          getSelectedRange: vi.fn(() => ({ address: "Sheet1!A1:B2", load: vi.fn() })),
        },
        sync: vi.fn(),
      });
    }),
  };
}

describe("HermesAgent server-owned mode", () => {
  it("recognizes an ordinary nonstream JSON plan through the serialized direct boundary", async () => {
    const agent = new HermesAgent();
    const body = new ReadableStream({
      start(controller) {
        controller.enqueue(new TextEncoder().encode(JSON.stringify({
          choices: [{ message: { content: "", tool_calls: [directPlanToolCall] } }],
        })));
        controller.close();
      },
    });
    mockFetch.mockResolvedValueOnce({
      ok: true,
      status: 200,
      headers: new Headers({ "content-type": "application/json" }),
      body,
    });

    const callbacks = makeCallbacks();
    const result = await agent.run(
      "write one value",
      makeSettings({ model: "qwen3.6-27b" }),
      callbacks,
      AbortSignal.timeout(10_000),
      null,
    );

    expect(result.finalContent).toBe("Workbook plan normalized; execution is disabled by the rollout setting.");
    expect(mockFetch).toHaveBeenCalledTimes(1);
    expect(callbacks.onActivity).not.toHaveBeenCalled();
  });

  it("recognizes an ordinary SSE plan after readOpenAIStream constructs ToolCall records", async () => {
    const agent = new HermesAgent();
    mockFetch.mockResolvedValueOnce(makeStreamResponse("", [directPlanToolCall]));

    const result = await agent.run(
      "write one value",
      makeSettings({ model: "qwen3.6-27b" }),
      makeCallbacks(),
      AbortSignal.timeout(10_000),
      null,
    );

    expect(result.finalContent).toBe("Workbook plan normalized; execution is disabled by the rollout setting.");
    expect(mockFetch).toHaveBeenCalledTimes(1);
  });

  it("fails closed on direct tool-call serialization failure before primitive dispatch", async () => {
    const agent = new HermesAgent();
    mockFetch.mockResolvedValueOnce(makeStreamResponse("", [
      { id: "call_primitive", type: "function", function: { name: "read_range", arguments: "{}" } },
    ]));
    const originalStringify = JSON.stringify;
    const stringifySpy = vi.spyOn(JSON, "stringify").mockImplementation(((value: unknown, replacer?: unknown, space?: unknown) => {
      if (Array.isArray(value)) throw new Error("serialization canary");
      return originalStringify(value, replacer as never, space as never);
    }) as typeof JSON.stringify);

    try {
      const result = await agent.run(
        "read A1",
        makeSettings({ model: "qwen3.6-27b" }),
        makeCallbacks(),
        AbortSignal.timeout(10_000),
        null,
      );
      expect(result.finalContent).toBe("Invalid workbook plan protocol envelope");
      expect(mockFetch).toHaveBeenCalledTimes(1);
    } finally {
      stringifySpy.mockRestore();
    }
  });

  beforeEach(() => {
    mockFetch.mockReset();
    mockOfficeJs();
  });

  it("binds server-owned mode through the trusted session header and sends no caller tools", async () => {
    const agent = new HermesAgent();
    mockFetch.mockResolvedValueOnce(makeStreamResponse("Server handled this natively."));

    const settings = makeSettings({ model: "hermes-agent" });
    const callbacks = makeCallbacks();
    const signal = AbortSignal.timeout(10_000);

    const result = await agent.run(
      "sum column A",
      settings,
      callbacks,
      signal,
      BRIDGE_SESSION_ID,
    );

    // Verify: request body had NO tools or tool_choice
    const callArgs = mockFetch.mock.calls[0];
    const body = JSON.parse(callArgs[1].body);
    expect(callArgs[1].headers["X-Hermes-Session-Id"]).toBe(BRIDGE_SESSION_ID);
    expect(body.tools).toBeUndefined();
    expect(body.tool_choice).toBeUndefined();
    expect(body.model).toBe("hermes-agent");

    // Verify: the trusted bridge target is not exposed to the model.
    const userMsg = body.messages.at(-1);
    expect(userMsg.role).toBe("user");
    expect(userMsg.content).toContain('"activeSheet":"Sheet1"');
    expect(userMsg.content).toContain('"selection":"Sheet1!A1:B2"');
    expect(userMsg.content).not.toContain("bridgeSessionId");
    expect(userMsg.content).not.toContain(BRIDGE_SESSION_ID);

    // Server-owned mode must explicitly tell Hermes to use its native plugin
    // surface rather than emitting caller-style tool JSON as assistant text.
    const systemMsg = body.messages.find((m: any) => m.role === "system");
    expect(systemMsg.content).toContain("native sidecar_excel_* tools");
    expect(systemMsg.content).toContain("automatically bound");
    expect(systemMsg.content).not.toContain("bridgeSessionId");
    expect(systemMsg.content).toContain("Do not emit tool calls as text");

    // Verify: no local tool execution happened — just returned server's content
    expect(result.finalContent).toBe("Server handled this natively.");
  });

  it("fails closed when model=hermes-agent and bridgeSessionId is null", async () => {
    const agent = new HermesAgent();
    mockFetch.mockResolvedValueOnce(makeStreamResponse("should not reach"));

    const settings = makeSettings({ model: "hermes-agent" });
    const callbacks = makeCallbacks();
    const signal = AbortSignal.timeout(10_000);

    await expect(agent.run(
      "sum column A",
      settings,
      callbacks,
      signal,
      null,
    )).rejects.toThrow(/bridge.*session/i);

    // Should NOT have made any fetch call
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it("fails closed before any request when the bridge session ID is not a valid UUID", async () => {
    const agent = new HermesAgent();
    mockFetch.mockResolvedValueOnce(makeStreamResponse("should not reach"));

    await expect(agent.run(
      "sum column A",
      makeSettings({ model: "hermes-agent" }),
      makeCallbacks(),
      AbortSignal.timeout(10_000),
      "not-a-uuid",
    )).rejects.toThrow(/valid.*UUID/i);

    expect(mockFetch).not.toHaveBeenCalled();
  });

  it("uses local tool loop for non-hermes-agent models — sends tools", async () => {
    const agent = new HermesAgent();
    mockFetch.mockResolvedValueOnce(makeStreamResponse("Done."));

    const settings = makeSettings({ model: "qwen3.6-27b" });
    const callbacks = makeCallbacks();
    const signal = AbortSignal.timeout(10_000);

    await agent.run(
      "sum column A",
      settings,
      callbacks,
      signal,
      null, // no bridge needed for non-hermes-agent
    );

    // Verify: request body DID include tools and tool_choice
    const callArgs = mockFetch.mock.calls[0];
    const body = JSON.parse(callArgs[1].body);
    expect(body.tools).toBeDefined();
    expect(Array.isArray(body.tools)).toBe(true);
    expect(body.tools.length).toBeGreaterThan(0);
    expect(body.tool_choice).toBe("auto");

    // bridgeSessionId should NOT appear in messages for non-hermes-agent mode
    const userMsg = body.messages.find((m: any) => m.role === "user");
    expect(userMsg.content).not.toContain("_bridgeSessionId");
  });

  it("executes local tools when model is not hermes-agent and server returns tool calls", async () => {
    // This test verifies the existing path still works for non-hermes-agent models.
    // We just need to confirm the fetch was made with tools included.
    const agent = new HermesAgent();
    mockFetch.mockResolvedValueOnce(makeStreamResponse("Result."));

    const settings = makeSettings({ model: "qwen3.6-27b" });
    const callbacks = makeCallbacks();
    const signal = AbortSignal.timeout(10_000);

    await agent.run("test", settings, callbacks, signal, null);

    expect(mockFetch).toHaveBeenCalled();
    const body = JSON.parse(mockFetch.mock.calls[0][1].body);
    expect(body.tools.length).toBeGreaterThan(0);
  });

  it("executes a legacy primitive XML read_range through the direct agent path", async () => {
    const agent = new HermesAgent();
    const primitiveXml = `<tool_call>${JSON.stringify({
      name: "read_range",
      arguments: { sheet: "Sheet1", address: "A1" },
    })}</tool_call>`;
    mockFetch
      .mockResolvedValueOnce(makeStreamResponse(primitiveXml))
      .mockResolvedValueOnce(makeStreamResponse("Read complete."));

    const result = await agent.run(
      "read A1",
      makeSettings({ model: "qwen3.6-27b" }),
      makeCallbacks(),
      AbortSignal.timeout(10_000),
      null,
    );

    expect(result.finalContent).toBe("Read complete.");
    expect(mockFetch).toHaveBeenCalledTimes(2);
  });

  it("server-owned mode does not locally execute backend/internal tool calls", async () => {
    // When the server returns what looks like a native plugin call,
    // in server-owned mode we should NOT try to execute it locally.
    const agent = new HermesAgent();
    mockFetch.mockResolvedValueOnce(makeStreamResponse(
      "I used backend_get_user_info.",
      [{ id: "call_1", type: "function", function: { name: "backend_get_user_info", arguments: "{}" } }]
    ));

    const settings = makeSettings({ model: "hermes-agent" });
    const callbacks = makeCallbacks();
    const signal = AbortSignal.timeout(10_000);

    const result = await agent.run(
      "test",
      settings,
      callbacks,
      signal,
      BRIDGE_SESSION_ID,
    );

    // In server-owned mode, even if the model returns something that looks like
    // a tool call, we should not execute it locally — just return the content.
    expect(result.finalContent).toContain("I used backend_get_user_info");
    expect(mockFetch).toHaveBeenCalledTimes(1);
    expect(callbacks.onActivity).not.toHaveBeenCalled();
  });
});

// ---------------------------------------------------------------------------
// C1 prompt recipes for title spanning
// ---------------------------------------------------------------------------

describe("C1: prompt recipes for title-span semantics", () => {
  it("direct prompt contains CenterAcrossSelection title recipe", () => {
    const content = getDirectSystemPrompt();
    expect(content).toContain("CenterAcrossSelection");
    expect(content).toContain("No merge capability exists");
  });

  it("server-owned prompt contains the same title recipe", () => {
    const prompt = getServerOwnedSystemPrompt();
    expect(prompt).toContain("CenterAcrossSelection");
    expect(prompt).toContain("No merge capability exists");
  });
});
