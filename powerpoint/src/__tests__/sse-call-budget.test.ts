import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { HermesAgent, readOpenAIStream, type AgentCallbacks } from "../agent";
import { getPresentationContext } from "../powerPointTools";

vi.mock("../powerPointTools", () => ({
  getPresentationContext: vi.fn().mockResolvedValue("metadata"),
  toolSchemas: [],
  executePowerPointToolParsed: vi.fn(),
}));

const settings = { baseUrl: "http://test.local/v1", model: "test", maxIterations: 1 };
const callbacks = (): AgentCallbacks => ({
  onAssistantStart: vi.fn(),
  onAssistantDelta: vi.fn(),
  onAssistantFinal: vi.fn(),
  onActivity: vi.fn(),
  onStatus: vi.fn(),
});

function sseResponse(payload: string): Response {
  return new Response(new ReadableStream({
    start(controller) {
      controller.enqueue(new TextEncoder().encode(payload));
      controller.close();
    },
  }), { status: 200, headers: { "content-type": "text/event-stream" } });
}

function deltaEvent(content: string, lineEnding = "\n"): string {
  return `data: {"choices":[{"delta":{"content":${JSON.stringify(content)}}}]}${lineEnding}${lineEnding}`;
}

describe("logical model-call budget and incremental SSE framing", () => {
  beforeEach(() => {
    vi.mocked(getPresentationContext).mockResolvedValue("metadata");
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
    vi.useRealTimers();
  });

  it("delivers a complete CRLF event before the next chunk is released", async () => {
    let releaseSecondChunk!: () => void;
    const secondChunk = new Promise<void>((resolve) => { releaseSecondChunk = resolve; });
    const first = deltaEvent("first", "\r\n");
    const body = new ReadableStream<Uint8Array>({
      async start(controller) {
        controller.enqueue(new TextEncoder().encode(first));
        await secondChunk;
        controller.enqueue(new TextEncoder().encode(`data: [DONE]\r\n\r\n`));
        controller.close();
      },
    });
    const cb = callbacks();
    const pending = readOpenAIStream(body, cb, new AbortController().signal);
    await vi.waitFor(() => expect(cb.onAssistantDelta).toHaveBeenCalledWith("first"));
    expect(cb.onAssistantDelta).toHaveBeenCalledTimes(1);
    releaseSecondChunk();
    await expect(pending).resolves.toMatchObject({ content: "first" });
  });

  it("frames an event when CRLF delimiter bytes are split across chunks", async () => {
    const payload = deltaEvent("split", "\r\n");
    const splitAt = payload.indexOf("\r\n\r\n") + 2;
    const chunks = [payload.slice(0, splitAt), payload.slice(splitAt, splitAt + 1), payload.slice(splitAt + 1), "data: [DONE]\r\n\r\n"];
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        for (const chunk of chunks) controller.enqueue(new TextEncoder().encode(chunk));
        controller.close();
      },
    });
    await expect(readOpenAIStream(body, callbacks(), new AbortController().signal)).resolves.toMatchObject({ content: "split" });
  });

  it("joins multiline data fields before parsing JSON", async () => {
    const payload = "data: {\"choices\":[{\"delta\":\n" +
      "data: {\"content\":\"multiline\"}}]}\n\n" +
      "data: [DONE]\n\n";
    await expect(readOpenAIStream(
      new ReadableStream({ start(controller) { controller.enqueue(new TextEncoder().encode(payload)); controller.close(); } }),
      callbacks(),
      new AbortController().signal,
    )).resolves.toMatchObject({ content: "multiline" });
  });

  it("rejects malformed first-event JSON as protocol without retry or fallback", async () => {
    const fetchMock = vi.fn().mockResolvedValue(sseResponse("data: {not-json}\n\n"));
    vi.stubGlobal("fetch", fetchMock);
    await expect(new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal))
      .rejects.toMatchObject({ toolError: { category: "protocol" } });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("marks a malformed EOF tail after a valid delta partial and never retries", async () => {
    const fetchMock = vi.fn().mockResolvedValue(sseResponse(`${deltaEvent("valid")}data: {not-json}`));
    vi.stubGlobal("fetch", fetchMock);
    await expect(new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal))
      .rejects.toMatchObject({ partial: true, toolError: { category: "protocol" } });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("uses one 20-second envelope across a 503 response and retry backoff", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);
    const fetchMock = vi.fn().mockImplementationOnce(() => new Promise<Response>((resolve) => {
      setTimeout(() => resolve(new Response("busy", { status: 503 })), 19_750);
    })).mockResolvedValueOnce(sseResponse(`${deltaEvent("late")}data: [DONE]\n\n`));
    vi.stubGlobal("fetch", fetchMock);
    const pending = new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal);
    const rejection = expect(pending).rejects.toMatchObject({ name: "TimeoutError" });
    await vi.advanceTimersByTimeAsync(19_750);
    await vi.advanceTimersByTimeAsync(250);
    await rejection;
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("does not start another fetch when remaining call budget cannot cover backoff", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);
    const fetchMock = vi.fn().mockImplementationOnce(() => new Promise<Response>((resolve) => {
      setTimeout(() => resolve(new Response("busy", { status: 503 })), 19_900);
    }));
    vi.stubGlobal("fetch", fetchMock);
    const pending = new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal);
    const rejection = expect(pending).rejects.toMatchObject({ name: "TimeoutError" });
    await vi.advanceTimersByTimeAsync(19_900);
    await vi.advanceTimersByTimeAsync(250);
    await rejection;
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("cleans the shared deadline timer and abort listeners after success", async () => {
    vi.useFakeTimers();
    const removeSpy = vi.spyOn(AbortSignal.prototype, "removeEventListener");
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(sseResponse(`${deltaEvent("done")}data: [DONE]\n\n`)));
    await new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal);
    expect(vi.getTimerCount()).toBe(0);
    expect(removeSpy).toHaveBeenCalled();
  });
});
