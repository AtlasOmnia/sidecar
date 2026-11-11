import { describe, expect, it, vi } from "vitest";
import { isBlockedRepeatedToolCall, makeToolCallSignature, readOpenAIStream } from "../agent";
import { ToolCall } from "../types";

describe("tool call loop guard helpers", () => {
  it("builds a stable signature from tool name and serialized arguments", () => {
    const toolCall: ToolCall = {
      id: "test",
      type: "function",
      function: {
        name: "summarize_range_stats",
        arguments: '{"sheet":"Sheet1","address":"A1:B2"}'
      }
    };

    expect(makeToolCallSignature(toolCall)).toBe(
      "summarize_range_stats:{\"address\":\"A1:B2\",\"sheet\":\"Sheet1\"}"
    );
  });

  it("canonicalizes equivalent argument JSON for repeat detection", () => {
    const toolCallA: ToolCall = {
      id: "test",
      type: "function",
      function: {
        name: "read_range",
        arguments: '{"sheet":"Sheet1","address":"A1:B2"}'
      }
    };

    const toolCallB: ToolCall = {
      id: "test_2",
      type: "function",
      function: {
        name: "read_range",
        arguments: '{"address":"A1:B2","sheet":"Sheet1"}'
      }
    };

    expect(makeToolCallSignature(toolCallA)).toBe(makeToolCallSignature(toolCallB));
  });

  it("treats equivalent JSON-ordered calls as repeated loop entries", () => {
    const counter = new Map<string, number>();
    const toolCallA: ToolCall = {
      id: "test",
      type: "function",
      function: {
        name: "read_range",
        arguments: '{"sheet":"Sheet1","address":"A1:B2"}'
      }
    };
    const toolCallB: ToolCall = {
      id: "test_2",
      type: "function",
      function: {
        name: "read_range",
        arguments: '{"address":"A1:B2","sheet":"Sheet1"}'
      }
    };

    const nextCount = (call: ToolCall) => {
      const signature = makeToolCallSignature(call);
      const count = (counter.get(signature) ?? 0) + 1;
      counter.set(signature, count);
      return isBlockedRepeatedToolCall(count);
    };

    expect(nextCount(toolCallA)).toBe(false);
    expect(nextCount(toolCallB)).toBe(false);
    expect(nextCount(toolCallA)).toBe(true);
  });

  it("marks calls as blocked only after the third identical appearance", () => {
    expect(isBlockedRepeatedToolCall(1)).toBe(false);
    expect(isBlockedRepeatedToolCall(2)).toBe(false);
    expect(isBlockedRepeatedToolCall(3)).toBe(true);
    expect(isBlockedRepeatedToolCall(4)).toBe(true);
  });

  it("parses a final SSE block without a trailing separator", async () => {
    const callbacks = {
      onAssistantStart: vi.fn(),
      onAssistantDelta: vi.fn(),
      onAssistantFinal: vi.fn(),
      onActivity: vi.fn(),
      onStatus: vi.fn()
    };

    const payload =
      '{"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_1","function":{"name":"read_range","arguments":"{\\"sheet\\":\\"Data\\"}"}}]}}]}';

    const chunk = `data: ${payload}\ndata: [DONE]`;

    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new TextEncoder().encode(chunk));
        controller.close();
      }
    });

    const result = await readOpenAIStream(body, callbacks, new AbortController().signal);

    expect(result.toolCalls).toEqual([
      {
        id: "call_1",
        type: "function",
        function: {
          name: "read_range",
          arguments: '{"sheet":"Data"}'
        }
      }
    ]);
    expect(result.content).toBe("");
  });

  it("concatenates repeated incremental argument deltas without guessing retransmission", async () => {
    const callbacks = {
      onAssistantStart: vi.fn(),
      onAssistantDelta: vi.fn(),
      onAssistantFinal: vi.fn(),
      onActivity: vi.fn(),
      onStatus: vi.fn()
    };

    const payload1 =
      '{"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_2","function":{"name":"read_range","arguments":"AAA"}}]}}]}';
    const payload2 =
      '{"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"AAA"}}]}}]}';

    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new TextEncoder().encode(`data: ${payload1}\n`));
        controller.enqueue(new TextEncoder().encode(`data: ${payload2}\n`));
        controller.close();
      }
    });

    const result = await readOpenAIStream(body, callbacks, new AbortController().signal);

    expect(result.toolCalls).toHaveLength(1);
    expect(result.toolCalls[0]).toEqual({
      id: "call_2",
      type: "function",
      function: {
        name: "read_range",
        arguments: "AAAAAA"
      }
    });
  });

  it("preserves legitimate 2-char overlaps at streaming boundaries", async () => {
    // Delta 1: args="AAABB" then Delta 2: args="BBCCC"
    // Overlap "BB" = 2 chars — too short for minOverlap=3 threshold.
    // Must not be stripped; expected result: "AAABBBBCCC"
    const callbacks = {
      onAssistantStart: vi.fn(),
      onAssistantDelta: vi.fn(),
      onAssistantFinal: vi.fn(),
      onActivity: vi.fn(),
      onStatus: vi.fn()
    };

    const payload1 = '{"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_o1","function":{"name":"t","arguments":"AAABB"}}]}}]}';
    const payload2 = '{"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"BBCCC"}}]}}]}';

    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new TextEncoder().encode(`data: ${payload1}\n`));
        controller.enqueue(new TextEncoder().encode(`data: ${payload2}\n`));
        controller.close();
      }
    });

    const result = await readOpenAIStream(body, callbacks, new AbortController().signal);

    expect(result.toolCalls).toHaveLength(1);
    // BB overlap of 2 chars must NOT be stripped; concatenated as-is
    expect(result.toolCalls[0].function.arguments).toBe("AAABBBBCCC");
  });

  it("preserves legitimate >=3 char overlap between incremental deltas", async () => {
    // Delta values are incremental protocol fragments. String overlap is content,
    // not evidence of retransmission: "AAABBB" + "BBBCCC" must retain all Bs.
    const callbacks = {
      onAssistantStart: vi.fn(),
      onAssistantDelta: vi.fn(),
      onAssistantFinal: vi.fn(),
      onActivity: vi.fn(),
      onStatus: vi.fn()
    };

    const payload1 = '{"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_o2","function":{"name":"t","arguments":"AAABBB"}}]}}]}';
    const payload2 = '{"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"BBBCCC"}}]}}]}';

    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new TextEncoder().encode(`data: ${payload1}\n`));
        controller.enqueue(new TextEncoder().encode(`data: ${payload2}\n`));
        controller.close();
      }
    });

    const result = await readOpenAIStream(body, callbacks, new AbortController().signal);

    expect(result.toolCalls).toHaveLength(1);
    expect(result.toolCalls[0].function.arguments).toBe("AAABBBBBBCCC");
  });
});
