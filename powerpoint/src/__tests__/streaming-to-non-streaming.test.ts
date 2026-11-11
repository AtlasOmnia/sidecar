import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";
import { HermesAgent } from "../agent";
import { getPresentationContext } from "../powerPointTools";
import { HermesSettings } from "../settings";

// Mock powerPointTools to avoid Office.js dependencies
vi.mock("../powerPointTools", () => ({
  getPresentationContext: vi.fn().mockResolvedValue("Mock presentation context"),
  executePowerPointTool: vi.fn().mockImplementation(
    (_name: string, _args: string) => Promise.resolve(JSON.stringify({ ok: true }))
  ),
  toolSchemas: [
    {
      type: "function",
      function: {
        name: "test_tool",
        description: "A fake tool for testing",
        parameters: { type: "object", properties: {} }
      }
    }
  ]
}));

const dummySettings: HermesSettings = {
  baseUrl: "http://test.local/v1",
  model: "test-model",
  maxIterations: 3
};

function noopCallbacks() {
  return {
    onAssistantStart: vi.fn(),
    onAssistantDelta: vi.fn(),
    onAssistantFinal: vi.fn(),
    onActivity: vi.fn(),
    onStatus: vi.fn()
  };
}

/** Build a standard non-stream chat completions JSON body string. */
function nonStreamResponseBody(
  content: string,
  toolCalls?: Array<{ id: string; name: string; arguments: string }>
): string {
  const msg: Record<string, unknown> = {
    role: "assistant",
    content
  };
  if (toolCalls && toolCalls.length > 0) {
    msg.tool_calls = toolCalls.map((tc) => ({
      id: tc.id,
      type: "function" as const,
      function: { name: tc.name, arguments: tc.arguments }
    }));
  }
  return JSON.stringify({
    id: "test-chatcmpl",
    object: "chat.completion",
    created: 1000000,
    model: "test-model",
    choices: [
      {
        index: 0,
        message: msg,
        finish_reason: toolCalls && toolCalls.length > 0 ? "tool_calls" : "stop"
      }
    ],
    usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 }
  });
}

/** Helper: a standard 200 OK non-stream JSON Response. */
function nonStreamResponse(content: string, toolCalls?: Array<{ id: string; name: string; arguments: string }>): Response {
  return new Response(
    nonStreamResponseBody(content, toolCalls),
    { status: 200, headers: { "Content-Type": "application/json" } }
  );
}

/** SSE stream chunk for a delta with the given content. */
function streamChunk(content: string): Uint8Array {
  const payload = `data: {"choices":[{"index":0,"delta":{"role":"assistant","content":${JSON.stringify(content)}}}]}\n\n`;
  return new TextEncoder().encode(payload);
}

function doneChunk(): Uint8Array {
  return new TextEncoder().encode("data: [DONE]\n\n");
}

function makeStreamResponse(chunks: Uint8Array[]): Response {
  return new Response(
    new ReadableStream({
      async start(controller) {
        for (const chunk of chunks) {
          controller.enqueue(chunk);
        }
        controller.close();
      }
    }),
    { status: 200, headers: { "Content-Type": "text/event-stream" } }
  );
}

describe("non-stream fallback in agent loop", () => {
  let agent: HermesAgent;
  let fetchSpy: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    vi.mocked(getPresentationContext).mockResolvedValue("Mock presentation context");
    agent = new HermesAgent();
    fetchSpy = vi.fn();
    vi.stubGlobal("fetch", fetchSpy);
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  // ---------------------------------------------------------------------------
  // Regression coverage for streaming/non-stream response negotiation.
  // ---------------------------------------------------------------------------

  it("processes parseable non-stream JSON from HTTP 404 without another fetch", async () => {
    const fallbackContent = "Hello from non-stream fallback";
    fetchSpy.mockResolvedValueOnce(
      new Response(nonStreamResponseBody(fallbackContent), {
        status: 404,
        statusText: "Not Found",
        headers: { "Content-Type": "application/json" }
      })
    );

    const callbacks = noopCallbacks();
    const result = await agent.run("test", dummySettings, callbacks, new AbortController().signal);

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    const requestBody = JSON.parse(fetchSpy.mock.calls[0][1].body as string);
    expect(requestBody.stream).toBe(true);
    expect(result.finalContent).toBe(fallbackContent);
    expect(callbacks.onAssistantDelta).toHaveBeenCalledExactlyOnceWith(fallbackContent);
  });

  it("processes parseable non-stream JSON from HTTP 500 without another fetch", async () => {
    const fallbackContent = "recovered after 500";
    fetchSpy.mockResolvedValueOnce(
      new Response(nonStreamResponseBody(fallbackContent), {
        status: 500,
        statusText: "Internal Server Error",
        headers: { "Content-Type": "application/json" }
      })
    );

    const callbacks = noopCallbacks();
    const result = await agent.run("test", dummySettings, callbacks, new AbortController().signal);

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toBe(fallbackContent);
  });

  it("retries with stream=false when stream response body is absent", async () => {
    const fallbackContent = "body was null, fell back";
    fetchSpy
      .mockResolvedValueOnce(
        new Response(null, {
          status: 200,
          headers: { "Content-Type": "text/event-stream" }
        })
      )
      .mockResolvedValueOnce(nonStreamResponse(fallbackContent));

    const callbacks = noopCallbacks();
    const result = await agent.run("test", dummySettings, callbacks, new AbortController().signal);

    expect(fetchSpy).toHaveBeenCalledTimes(2);
    const secondBody = JSON.parse(fetchSpy.mock.calls[1][1].body as string);
    expect(secondBody.stream).toBe(false);
    expect(result.finalContent).toBe(fallbackContent);
  });

  it("retries with stream=false when stream returns [DONE]-only content", async () => {
    const fallbackContent = "recovered from empty stream";
    fetchSpy
      .mockResolvedValueOnce(makeStreamResponse([doneChunk()]))
      .mockResolvedValueOnce(nonStreamResponse(fallbackContent));

    const callbacks = noopCallbacks();
    const result = await agent.run("test", dummySettings, callbacks, new AbortController().signal);

    expect(fetchSpy).toHaveBeenCalledTimes(2);
    const secondBody = JSON.parse(fetchSpy.mock.calls[1][1].body as string);
    expect(secondBody.stream).toBe(false);
    expect(result.finalContent).toBe(fallbackContent);
  });

  it("extracts tool calls from non-stream fallback response preserving id/name/arguments", async () => {
    const body = nonStreamResponseBody("", [
      { id: "call_abc123", name: "get_presentation_overview", arguments: '{"name":"current"}' },
      { id: "call_def456", name: "read_all_slides", arguments: '{"slideIndex":0}' }
    ]);

    fetchSpy
      .mockResolvedValueOnce(
        new Response(body, {
          status: 404,
          headers: { "Content-Type": "application/json" }
        })
      )
      .mockResolvedValueOnce(makeStreamResponse([streamChunk("tool calls processed"), doneChunk()]));

    const callbacks = noopCallbacks();
    const result = await agent.run("test", dummySettings, callbacks, new AbortController().signal);

    expect(fetchSpy).toHaveBeenCalledTimes(2);
    const secondBody = JSON.parse(fetchSpy.mock.calls[1][1].body as string);
    expect(secondBody.stream).toBe(true);
    expect(result.finalContent).toBe("tool calls processed");
  });

  it("raises connectivity error after exhausting network retries", async () => {
    fetchSpy.mockRejectedValue(new TypeError("fetch failed"));

    const callbacks = noopCallbacks();
    await expect(
      agent.run("test", dummySettings, callbacks, new AbortController().signal)
    ).rejects.toThrow(/Could not reach/);

    // MAX_RETRIES=2 means the initial request plus two network retries.
    expect(fetchSpy).toHaveBeenCalledTimes(3);
  });

  it("does NOT fall back on HTTP 400 with error body (validation error, not stream issue)", async () => {
    fetchSpy.mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          error: { message: "Invalid model parameter", type: "invalid_request_error" }
        }),
        {
          status: 400,
          statusText: "Bad Request",
          headers: { "Content-Type": "application/json" }
        }
      )
    );

    const callbacks = noopCallbacks();
    await expect(
      agent.run("test", dummySettings, callbacks, new AbortController().signal)
    ).rejects.toThrow(/Model endpoint returned 400/);

    // Exactly one attempt — HTTP 400 does NOT trigger non-stream fallback
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });

  it("does not produce duplicate assistant deltas from fallback response", async () => {
    const fallbackContent = "unique delta";
    const errBody = nonStreamResponseBody(fallbackContent);
    fetchSpy.mockResolvedValueOnce(
      new Response(errBody, {
        status: 404,
        headers: { "Content-Type": "application/json" }
      })
    );

    const callbacks = noopCallbacks();
    const result = await agent.run("test", dummySettings, callbacks, new AbortController().signal);

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    // onAssistantDelta should be called exactly once with the content
    const deltaCalls = callbacks.onAssistantDelta.mock.calls;
    expect(deltaCalls).toHaveLength(1);
    expect(deltaCalls[0][0]).toBe(fallbackContent);
    expect(result.finalContent).toBe(fallbackContent);
  });
});
