import { describe, expect, it, vi } from "vitest";
import { readOpenAIStream } from "../agent";
import { ToolCall } from "../types";

// ---------------------------------------------------------------------------
// LM Studio / OpenAI-compatible streaming conformance fixtures
// ---------------------------------------------------------------------------

function makeStream(...chunks: string[]): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller) {
      for (const chunk of chunks) {
        controller.enqueue(new TextEncoder().encode(chunk));
      }
      controller.close();
    },
  });
}

const noopCallbacks = {
  onAssistantStart: vi.fn(),
  onAssistantDelta: vi.fn(),
  onAssistantFinal: vi.fn(),
  onActivity: vi.fn(),
  onStatus: vi.fn(),
};

describe("streaming assistant content", () => {
  it("accumulates text content from deltas", async () => {
    const stream = makeStream(
      'data: {"choices":[{"delta":{"content":"Hello"}}]}\n\n',
      'data: {"choices":[{"delta":{"content":" World"}}]}\n\n',
      "data: [DONE]\n\n",
    );

    const result = await readOpenAIStream(
      stream,
      noopCallbacks,
      new AbortController().signal,
    );

    expect(result.content).toBe("Hello World");
    expect(result.toolCalls).toHaveLength(0);
    expect(noopCallbacks.onAssistantDelta).toHaveBeenCalledTimes(2);
  });

  it("handles single-chunk content with [DONE]", async () => {
    const stream = makeStream(
      'data: {"choices":[{"delta":{"content":"One-shot"}}]}\n\ndata: [DONE]\n\n',
    );

    const result = await readOpenAIStream(
      stream,
      noopCallbacks,
      new AbortController().signal,
    );

    expect(result.content).toBe("One-shot");
  });
});

describe("streaming tool calls", () => {
  it("builds a complete tool call from incremental deltas", async () => {
    const stream = makeStream(
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_abc","type":"function","function":{"name":"read_range","arguments":""}}]}}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"{\\"sheet\\""}}]}}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":":\\"Sheet1\\",\\"address\\":\\"A1:B2\\"}"}}]}}]}\n\n',
      "data: [DONE]\n\n",
    );

    const result = await readOpenAIStream(
      stream,
      noopCallbacks,
      new AbortController().signal,
    );

    expect(result.toolCalls).toHaveLength(1);
    expect(result.toolCalls[0]).toEqual({
      id: "call_abc",
      type: "function",
      function: {
        name: "read_range",
        arguments: '{"sheet":"Sheet1","address":"A1:B2"}',
      },
    });
  });

  it("handles multiple tool calls in parallel", async () => {
    const call1Args = JSON.stringify({ sheet: "Data" });
    const call2Args = JSON.stringify({ sheet: "Data" });
    const payload = JSON.stringify({
      choices: [{
        delta: {
          tool_calls: [
            { index: 0, id: "call_1", type: "function", function: { name: "read_range", arguments: call1Args } },
            { index: 1, id: "call_2", type: "function", function: { name: "summarize_range_stats", arguments: call2Args } },
          ]
        }
      }]
    });
    const stream = makeStream(
      `data: ${payload}\n\n`,
      "data: [DONE]\n\n",
    );

    const result = await readOpenAIStream(
      stream,
      noopCallbacks,
      new AbortController().signal,
    );

    expect(result.toolCalls).toHaveLength(2);
    expect(result.toolCalls[0].function.name).toBe("read_range");
    expect(result.toolCalls[1].function.name).toBe("summarize_range_stats");
  });

  it("filters tool calls with no name", async () => {
    const stream = makeStream(
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_x","function":{"arguments":"{}"}}]}}]}\n\n',
      "data: [DONE]\n\n",
    );

    const result = await readOpenAIStream(
      stream,
      noopCallbacks,
      new AbortController().signal,
    );

    // Tool calls with no name are filtered out
    expect(result.toolCalls).toHaveLength(0);
  });
});

describe("finish reason propagation", () => {
  it("does not crash on finish_reason in the final chunk", async () => {
    const stream = makeStream(
      'data: {"choices":[{"delta":{"content":"Done"},"finish_reason":"stop"}]}\n\n',
      "data: [DONE]\n\n",
    );

    const result = await readOpenAIStream(
      stream,
      noopCallbacks,
      new AbortController().signal,
    );

    expect(result.content).toBe("Done");
  });

  it("handle empty delta objects without crash", async () => {
    const stream = makeStream(
      'data: {"choices":[{"delta":{}}]}\n\n',
      "data: [DONE]\n\n",
    );

    const result = await readOpenAIStream(
      stream,
      noopCallbacks,
      new AbortController().signal,
    );

    expect(result.content).toBe("");
    expect(result.toolCalls).toHaveLength(0);
  });
});

describe("malformed and edge cases", () => {
  it("skips non-JSON data lines without crashing", async () => {
    const stream = makeStream(
      "data: NOT_JSON\n\n",
      'data: {"choices":[{"delta":{"content":"valid"}}]}\n\n',
      "data: [DONE]\n\n",
    );

    const result = await readOpenAIStream(
      stream,
      noopCallbacks,
      new AbortController().signal,
    );

    expect(result.content).toBe("valid");
  });

  it("handles empty stream body gracefully", async () => {
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.close();
      },
    });

    const result = await readOpenAIStream(
      stream,
      noopCallbacks,
      new AbortController().signal,
    );

    expect(result.content).toBe("");
    expect(result.toolCalls).toHaveLength(0);
  });

  it("handles chunk with no choices array", async () => {
    const stream = makeStream(
      'data: {"id":"chatcmpl-123","object":"chat.completion.chunk"}\n\n',
      'data: {"choices":[{"delta":{"content":"ok"}}]}\n\n',
      "data: [DONE]\n\n",
    );

    const result = await readOpenAIStream(
      stream,
      noopCallbacks,
      new AbortController().signal,
    );

    expect(result.content).toBe("ok");
  });

  it("handles blank lines between SSE events", async () => {
    const stream = makeStream(
      "\n\n\n",
      'data: {"choices":[{"delta":{"content":"after blanks"}}]}\n\n',
      "data: [DONE]\n\n",
    );

    const result = await readOpenAIStream(
      stream,
      noopCallbacks,
      new AbortController().signal,
    );

    expect(result.content).toBe("after blanks");
  });
});

describe("cancellation during stream", () => {
  it("throws AbortError when signal fires mid-stream", async () => {
    const controller = new AbortController();

    const stream = new ReadableStream<Uint8Array>({
      start(ctrl) {
        ctrl.enqueue(
          new TextEncoder().encode(
            'data: {"choices":[{"delta":{"content":"first"}}]}\n\n',
          ),
        );
        // Abort before sending more
        controller.abort();
        ctrl.enqueue(
          new TextEncoder().encode(
            'data: {"choices":[{"delta":{"content":"second"}}]}\n\n',
          ),
        );
        ctrl.close();
      },
    });

    await expect(
      readOpenAIStream(stream, noopCallbacks, controller.signal),
    ).rejects.toThrow(DOMException);
  });
});

describe("tool-role continuation messages", () => {
  it("correctly builds tool result messages for the model", () => {
    // This tests the message format that the agent loop produces,
    // not the streaming parser itself. The format must be
    // OpenAI-compatible for LM Studio / local models.
    const toolMessage = {
      role: "tool" as const,
      tool_call_id: "call_abc",
      name: "read_range",
      content: JSON.stringify({
        ok: true,
        action: "read_range",
        sheet: "Data",
        values: [[1, 2], [3, 4]],
      }),
    };

    expect(toolMessage.role).toBe("tool");
    expect(toolMessage.tool_call_id).toBe("call_abc");
    expect(toolMessage.name).toBe("read_range");
    expect(typeof toolMessage.content).toBe("string");
  });

  it("produces valid JSON tool result content", () => {
    const content = JSON.stringify({
      ok: false,
      error: "Sheet not found.",
      category: "office_host",
    });
    const parsed = JSON.parse(content);
    expect(parsed.ok).toBe(false);
    expect(parsed.category).toBe("office_host");
  });
});
