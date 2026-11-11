// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { HermesAgent, type AgentCallbacks } from "../agent";
import { getPresentationContext, executePowerPointToolParsed } from "../powerPointTools";
import { asToolError, validationError } from "../toolErrors";
import { validateToolCall } from "../toolValidation";

vi.mock("../powerPointTools", () => ({
  getPresentationContext: vi.fn(),
  executePowerPointToolParsed: vi.fn(),
  toolSchemas: [{
    type: "function",
    function: {
      name: "office_tool",
      description: "test",
      parameters: { type: "object", properties: {}, additionalProperties: false },
    },
  }],
}));

const settings = { baseUrl: "http://test.local/v1", model: "test", maxIterations: 1 };
const callbacks = (): AgentCallbacks => ({
  onAssistantStart: vi.fn(),
  onAssistantDelta: vi.fn(),
  onAssistantFinal: vi.fn(),
  onActivity: vi.fn(),
  onStatus: vi.fn(),
});
const finalResponse = () => new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "later" } }] }), {
  status: 200,
  headers: { "content-type": "application/json" },
});
const toolResponse = () => new Response(JSON.stringify({ choices: [{ message: {
  role: "assistant",
  content: "",
  tool_calls: [{ id: "tool-1", type: "function", function: { name: "office_tool", arguments: "{}" } }],
} }] }), { status: 200, headers: { "content-type": "application/json" } });

async function flush(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
  await Promise.resolve();
}

beforeEach(() => {
  vi.mocked(getPresentationContext).mockResolvedValue("metadata");
  vi.mocked(executePowerPointToolParsed).mockResolvedValue(JSON.stringify({ ok: true }));
});

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
  vi.useRealTimers();
});

describe("fail-closed pending external operation latch", () => {
  it("blocks a second run after a timed-out non-cooperative presentation context until it settles", async () => {
    vi.useFakeTimers();
    let settleContext!: (value: string) => void;
    vi.mocked(getPresentationContext).mockReturnValue(new Promise((resolve) => { settleContext = resolve; }));
    const agent = new HermesAgent();
    const first = agent.run("first", settings, callbacks(), new AbortController().signal);
    const firstAssertion = expect(first).rejects.toMatchObject({ name: "TimeoutError", toolError: { category: "timeout" } });
    await vi.advanceTimersByTimeAsync(90_001);
    await firstAssertion;

    await expect(agent.run("second", settings, callbacks(), new AbortController().signal)).rejects.toMatchObject({
      toolError: { category: "internal" },
      message: expect.not.stringContaining("first"),
    });
    expect(getPresentationContext).toHaveBeenCalledTimes(1);
    expect(agent.history).toEqual([]);

    settleContext("metadata");
    await flush();
    vi.mocked(getPresentationContext).mockResolvedValue("metadata");
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(finalResponse()));
    await expect(agent.run("later", settings, callbacks(), new AbortController().signal)).resolves.toMatchObject({ finalContent: "later" });
  });

  it("reset invalidates stale history but cannot bypass a pending context operation", async () => {
    let settleContext!: (value: string) => void;
    vi.mocked(getPresentationContext).mockReturnValue(new Promise((resolve) => { settleContext = resolve; }));
    const agent = new HermesAgent();
    const first = agent.run("stale", settings, callbacks(), new AbortController().signal);
    agent.reset();

    await expect(agent.run("blocked", settings, callbacks(), new AbortController().signal)).rejects.toMatchObject({ toolError: { category: "internal" } });
    settleContext("metadata");
    await expect(first).rejects.toMatchObject({ name: "AbortError" });
    await flush();
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(finalResponse()));
    await expect(agent.run("later", settings, callbacks(), new AbortController().signal)).resolves.toMatchObject({ finalContent: "later" });
    expect(agent.history.some((message) => message.content?.includes("stale"))).toBe(false);
  });

  it("blocks a second run while a timed-out non-cooperative fetch remains unsettled", async () => {
    vi.useFakeTimers();
    const fetchMock = vi.fn(() => new Promise<Response>(() => {}));
    vi.stubGlobal("fetch", fetchMock);
    const agent = new HermesAgent();
    const first = agent.run("first", settings, callbacks(), new AbortController().signal);
    const firstAssertion = expect(first).rejects.toMatchObject({ name: "TimeoutError" });
    await vi.advanceTimersByTimeAsync(20_001);
    await firstAssertion;

    await expect(agent.run("second", settings, callbacks(), new AbortController().signal)).rejects.toMatchObject({ toolError: { category: "internal" } });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("blocks a second run while the original response reader.read remains unsettled", async () => {
    let settleRead!: (result: ReadableStreamReadResult<Uint8Array>) => void;
    const reader = {
      read: vi.fn(() => new Promise<ReadableStreamReadResult<Uint8Array>>((resolve) => { settleRead = resolve; })),
      cancel: vi.fn(() => Promise.resolve()),
    };
    const response = { ok: true, status: 200, body: { getReader: () => reader }, headers: { get: () => "text/event-stream" } } as unknown as Response;
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(response));
    const agent = new HermesAgent();
    const controller = new AbortController();
    const first = agent.run("first", settings, callbacks(), controller.signal);
    await vi.waitFor(() => expect(reader.read).toHaveBeenCalledTimes(1));
    controller.abort(new DOMException("stop", "AbortError"));
    await expect(first).rejects.toMatchObject({ name: "AbortError" });

    await expect(agent.run("second", settings, callbacks(), new AbortController().signal)).rejects.toMatchObject({ toolError: { category: "internal" } });
    settleRead({ done: true, value: undefined });
    await flush();
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(finalResponse()));
    await expect(agent.run("later", settings, callbacks(), new AbortController().signal)).resolves.toMatchObject({ finalContent: "later" });
  });

  it.each([false, true])("keeps a new run blocked until best-effort reader cancellation settles (reject=%s)", async (rejectCancel) => {
    let settleRead!: (result: ReadableStreamReadResult<Uint8Array>) => void;
    let settleCancel!: (reason?: unknown) => void;
    const reader = {
      read: vi.fn(() => new Promise<ReadableStreamReadResult<Uint8Array>>((resolve) => { settleRead = resolve; })),
      cancel: vi.fn(() => new Promise<void>((resolve, reject) => { settleCancel = rejectCancel ? (reason) => reject(reason) : () => resolve(); })),
    };
    const response = { ok: true, status: 200, body: { getReader: () => reader }, headers: { get: () => "text/event-stream" } } as unknown as Response;
    const fetchMock = vi.fn().mockResolvedValue(response);
    vi.stubGlobal("fetch", fetchMock);
    const agent = new HermesAgent();
    const controller = new AbortController();
    const first = agent.run("first", settings, callbacks(), controller.signal);
    await vi.waitFor(() => expect(reader.read).toHaveBeenCalledTimes(1));

    controller.abort(new DOMException("stop", "AbortError"));
    settleRead({ done: true, value: undefined });
    await expect(first).rejects.toMatchObject({ name: "AbortError" });
    await expect(agent.run("second", settings, callbacks(), new AbortController().signal)).rejects.toMatchObject({ toolError: { category: "internal" } });
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(agent.history).toEqual([]);

    if (rejectCancel) settleCancel(new Error("cancel failed"));
    else settleCancel();
    await flush();
    fetchMock.mockResolvedValue(finalResponse());
    await expect(agent.run("later", settings, callbacks(), new AbortController().signal)).resolves.toMatchObject({ finalContent: "later" });
  });

  it("blocks a second run while a non-cooperative checkpoint callback remains unsettled", async () => {
    vi.useFakeTimers({ now: 0 });
    let settleContext!: (value: string) => void;
    vi.mocked(getPresentationContext).mockReturnValue(new Promise((resolve) => { settleContext = resolve; }));
    let contextSettled = false;
    vi.spyOn(Date, "now").mockImplementation(() => contextSettled ? 30_001 : 0);
    let settleCheckpoint!: (decision: "continue" | "stop") => void;
    const cb = callbacks();
    cb.onCheckpoint = vi.fn(() => new Promise<"continue" | "stop">((resolve) => { settleCheckpoint = resolve; }));
    const agent = new HermesAgent();
    const first = agent.run("first", settings, cb, new AbortController().signal);
    await flush();
    contextSettled = true;
    settleContext("metadata");
    for (let i = 0; i < 20; i += 1) await Promise.resolve();
    expect(cb.onCheckpoint).toHaveBeenCalledTimes(1);
    const firstAssertion = expect(first).rejects.toMatchObject({ name: "TimeoutError" });
    await vi.advanceTimersByTimeAsync(90_001);
    await firstAssertion;

    await expect(agent.run("second", settings, callbacks(), new AbortController().signal)).rejects.toMatchObject({ toolError: { category: "internal" } });
    settleCheckpoint("continue");
    await flush();
    vi.setSystemTime(0);
    vi.mocked(getPresentationContext).mockResolvedValue("metadata");
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(finalResponse()));
    await expect(agent.run("later", settings, callbacks(), new AbortController().signal)).resolves.toMatchObject({ finalContent: "later" });
  });
});

describe("fail-closed error sanitization", () => {
  it("does not trust arbitrary valid:false records, while preserving branded validation messages", async () => {
    const sentinel = "RAW_SENTINEL_VALID_FALSE";
    vi.stubGlobal("Office", { onReady: vi.fn() });
    const { formatAgentError } = await import("../taskpane/taskpane");
    for (const category of ["internal", "office_host", "validation"] as const) {
      const record = { ok: false, valid: false, error: sentinel, category };
      expect(JSON.stringify(asToolError(record, category))).not.toContain(sentinel);
      expect(formatAgentError({ toolError: record })).not.toContain(sentinel);
    }

    const validation = validateToolCall("missing_tool", "{}");
    expect(validation.valid).toBe(false);
    expect(asToolError(validation).error).toBe("Unknown tool: missing_tool");
    expect(formatAgentError({ toolError: validation })).toContain("Unknown tool: missing_tool");
    expect(JSON.stringify(asToolError(validation))).toContain("Unknown tool: missing_tool");
    expect(JSON.stringify(asToolError(validationError("branded validation")))).toContain("branded validation");
  });

  it("replaces arbitrary internal, Office, protocol, network, and provider messages while retaining safe metadata", () => {
    const sentinel = "RAW_SENTINEL_SHOULD_NEVER_ESCAPE";
    for (const category of ["internal", "office_host", "protocol", "network", "provider"] as const) {
      const value = asToolError({ category, message: sentinel, statusCode: 503, code: "InvalidArgument" }, category);
      expect(JSON.stringify(value)).not.toContain(sentinel);
      expect(value.statusCode).toBe(503);
      expect(value.details?.code).toBe("InvalidArgument");
    }
    const capability = asToolError({ ok: false, category: "office_host", error: sentinel, details: { available: false, requiredApi: "PowerPointApi 1.8", reason: sentinel } });
    expect(JSON.stringify(capability)).not.toContain(sentinel);
    expect(capability.details).toMatchObject({ available: false, requiredApi: "PowerPointApi 1.8" });
  });

  it("uses the sanitized toolError message for task-pane display", async () => {
    vi.stubGlobal("Office", { onReady: vi.fn() });
    const { formatAgentError } = await import("../taskpane/taskpane");
    const sentinel = "RAW_SENTINEL_TASKPANE";
    expect(formatAgentError(new Error(sentinel))).not.toContain(sentinel);
  });

  it("does not put a raw context failure into later model messages or history", async () => {
    const sentinel = "RAW_SENTINEL_CONTEXT";
    vi.mocked(getPresentationContext).mockRejectedValueOnce(new Error(sentinel)).mockResolvedValue("metadata");
    const fetchMock = vi.fn().mockResolvedValue(finalResponse());
    vi.stubGlobal("fetch", fetchMock);
    const agent = new HermesAgent();
    await expect(agent.run("first", settings, callbacks(), new AbortController().signal)).rejects.toSatisfy((error: unknown) => {
      expect(JSON.stringify((error as { toolError?: unknown }).toolError)).not.toContain(sentinel);
      return true;
    });
    const result = await agent.run("later", settings, callbacks(), new AbortController().signal);
    expect(result.finalContent).toBe("later");
    expect(JSON.stringify(fetchMock.mock.calls)).not.toContain(sentinel);
    expect(JSON.stringify(agent.history)).not.toContain(sentinel);
  });

  it("sanitizes an arbitrary Office failure before it can escape the agent", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(toolResponse()));
    vi.mocked(executePowerPointToolParsed).mockRejectedValue(new Error("RAW_SENTINEL_OFFICE"));
    const agent = new HermesAgent();
    await expect(agent.run("edit", settings, callbacks(), new AbortController().signal)).rejects.toSatisfy((error: unknown) => {
      expect(JSON.stringify(error)).not.toContain("RAW_SENTINEL_OFFICE");
      expect(JSON.stringify(agent.history)).not.toContain("RAW_SENTINEL_OFFICE");
      return true;
    });
  });
});
