import { describe, expect, it, vi } from "vitest";
import { HermesAgent, chatCompletionsUrl, readOpenAIJsonResponse } from "../agent";

describe("chatCompletionsUrl", () => {
  it("appends chat completions to a /v1 base URL", () => {
    expect(chatCompletionsUrl("http://127.0.0.1:8645/v1")).toBe(
      "http://127.0.0.1:8645/v1/chat/completions"
    );
  });

  it("does not double-append chat completions", () => {
    expect(chatCompletionsUrl("http://127.0.0.1:8645/v1/chat/completions")).toBe(
      "http://127.0.0.1:8645/v1/chat/completions"
    );
  });

  it("targets the Sidecar bridge and refuses an unconnected (empty) base", () => {
    expect(chatCompletionsUrl("http://127.0.0.1:47630/v1")).toBe("http://127.0.0.1:47630/v1/chat/completions");
    expect(chatCompletionsUrl("http://127.0.0.1:47630")).toBe("http://127.0.0.1:47630/v1/chat/completions");
    expect(() => chatCompletionsUrl("")).toThrow(/Connect to Hermes/);
  });
});

describe("non-stream compatibility", () => {
  const callbacks = {
    onAssistantStart: () => undefined,
    onAssistantDelta: (_delta: string) => undefined,
    onAssistantFinal: () => undefined,
    onActivity: () => undefined,
    onStatus: () => undefined
  };

  it("parses assistant content and function tool calls from JSON", async () => {
    const deltas: string[] = [];
    const result = await readOpenAIJsonResponse(JSON.stringify({ choices: [{ message: {
      content: "I'll check the document.",
      tool_calls: [{ id: "call_abc", type: "function", function: { name: "get_document_overview", arguments: "{}" } }]
    }}] }), { ...callbacks, onAssistantDelta: (delta) => deltas.push(delta) }, new AbortController().signal);

    expect(result.content).toBe("I'll check the document.");
    expect(deltas).toEqual(["I'll check the document."]);
    expect(result.toolCalls[0]?.function.name).toBe("get_document_overview");
  });

  it("rejects malformed JSON response envelopes", async () => {
    await expect(readOpenAIJsonResponse("{}", callbacks, new AbortController().signal))
      .rejects.toThrow("choices[0].message");
  });

  it("parses a final SSE event even when the stream closes without a blank-line delimiter", async () => {
    const payload = JSON.stringify({ choices: [{ delta: { content: "final streamed text" } }] });
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new TextEncoder().encode(`data: ${payload}`));
        controller.close();
      }
    });
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      new Response(stream, { headers: { "content-type": "text/event-stream" } })
    );

    const result = await new HermesAgent().run(
      "Summarize this document",
      { baseUrl: "http://127.0.0.1:8645/v1", model: "gpt-4.1-mini", maxIterations: 1 },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: () => undefined
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toBe("final streamed text");
  });
});
