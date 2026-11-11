import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { HermesAgent, readOpenAIStream, readResponseBody, type AgentCallbacks } from "../agent";
import { getPresentationContext } from "../powerPointTools";

const executePowerPointToolParsed = vi.fn();
vi.mock("../powerPointTools", () => ({
  getPresentationContext: vi.fn().mockResolvedValue("deck metadata"),
  executePowerPointToolParsed: (...args: unknown[]) => executePowerPointToolParsed(...args),
  toolSchemas: [{
    type: "function",
    function: { name: "office_tool", description: "test", parameters: { type: "object", properties: {}, additionalProperties: false } },
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

function jsonResponse(content: string, toolCalls?: unknown[]): Response {
  return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content, ...(toolCalls ? { tool_calls: toolCalls } : {}) } }] }), {
    status: 200,
    headers: { "content-type": "application/json" },
  });
}

function officeCall(): unknown {
  return { id: "office-call", type: "function", function: { name: "office_tool", arguments: "{}" } };
}

function fakeResponse(reader: { read: () => Promise<unknown>; cancel: () => Promise<unknown> }): Response {
  return { body: { getReader: () => reader } } as unknown as Response;
}

describe("abort containment and Office overlap", () => {
  beforeEach(() => {
    vi.mocked(getPresentationContext).mockResolvedValue("deck metadata");
    executePowerPointToolParsed.mockReset();
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("rejects a hanging response read at abort without awaiting a hanging cancel", async () => {
    const cancel = vi.fn(() => new Promise<never>(() => {}));
    let rejectRead!: (reason?: unknown) => void;
    const reader = { read: vi.fn(() => new Promise<never>((_, reject) => { rejectRead = reject; })), cancel };
    const controller = new AbortController();
    const pending = readResponseBody(fakeResponse(reader), controller.signal);

    controller.abort(new DOMException("stop", "AbortError"));

    await expect(pending).rejects.toMatchObject({ name: "AbortError" });
    expect(cancel).toHaveBeenCalledTimes(1);
    rejectRead(new Error("late read failed"));
    await Promise.resolve();
  });

  it("rejects a hanging SSE read at timeout and contains a rejecting cancel", async () => {
    const cancel = vi.fn(() => Promise.reject(new Error("cancel failed")));
    const reader = { read: vi.fn(() => new Promise<never>(() => {})), cancel };
    const controller = new AbortController();
    const pending = readOpenAIStream(
      { getReader: () => reader } as unknown as ReadableStream<Uint8Array>,
      callbacks(),
      controller.signal,
    );

    controller.abort(new DOMException("deadline", "TimeoutError"));

    await expect(pending).rejects.toMatchObject({ name: "TimeoutError" });
    expect(cancel).toHaveBeenCalledTimes(1);
    await Promise.resolve();
  });

  it("rejects valid JSON left unterminated at SSE EOF", async () => {
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new TextEncoder().encode("data: {\"choices\":[{\"delta\":{\"content\":\"tail\"}}]}"));
        controller.close();
      },
    });

    await expect(readOpenAIStream(body, callbacks(), new AbortController().signal))
      .rejects.toMatchObject({ category: "protocol" });
  });

  it("marks an unterminated DONE tail partial and does not retry", async () => {
    const fetchMock = vi.fn().mockResolvedValue(new Response(new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new TextEncoder().encode("data: {\"choices\":[{\"delta\":{\"content\":\"before-tail\"}}]}\n\ndata: [DONE]"));
        controller.close();
      },
    }), { status: 200, headers: { "content-type": "text/event-stream" } }));
    vi.stubGlobal("fetch", fetchMock);

    await expect(new HermesAgent().run("tail", settings, callbacks(), new AbortController().signal))
      .rejects.toMatchObject({ partial: true, toolError: { category: "protocol" } });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("blocks new runs while a timed-out Office mutation is still settling, then permits a later run", async () => {
    let settleOffice!: (value: string) => void;
    let officeSettled = false;
    executePowerPointToolParsed.mockImplementation(() => {
      if (officeSettled) return Promise.resolve(JSON.stringify({ ok: true }));
      return new Promise<string>((resolve) => { settleOffice = resolve; });
    });
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(jsonResponse("", [officeCall()]))
      .mockResolvedValueOnce(jsonResponse("later"));
    vi.stubGlobal("fetch", fetchMock);

    const agent = new HermesAgent();
    const firstController = new AbortController();
    const first = agent.run("mutate first", settings, callbacks(), firstController.signal);
    await vi.waitFor(() => expect(executePowerPointToolParsed).toHaveBeenCalledTimes(1));
    firstController.abort(new DOMException("stop", "AbortError"));
    await expect(first).rejects.toMatchObject({ name: "AbortError" });

    const fetchesAfterAbort = fetchMock.mock.calls.length;
    const second = agent.run("mutate second", settings, callbacks(), new AbortController().signal);
    await expect(second).rejects.toMatchObject({
      message: expect.stringContaining("may still complete"),
      toolError: { category: "office_host" },
    });
    expect(fetchMock).toHaveBeenCalledTimes(fetchesAfterAbort);
    expect(agent.history).toEqual([]);

    officeSettled = true;
    settleOffice(JSON.stringify({ ok: true }));
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();
    const later = await agent.run("later request", settings, callbacks(), new AbortController().signal);
    expect(later.finalContent).toBe("later");
    expect(agent.history.filter((message) => message.role === "user")).toHaveLength(1);
    expect(agent.history[0]?.content).toContain("later request");
    expect(agent.history[0]?.content).not.toContain("mutate first");
  });
});
