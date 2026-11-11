import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";
import { HermesAgent, deadlineSignal, readOpenAIStream, MODEL_CALL_TIMEOUT_MS, type AgentCallbacks } from "../agent";
import { getPresentationContext } from "../powerPointTools";

vi.mock("../powerPointTools", () => ({
  getPresentationContext: vi.fn().mockResolvedValue("presentation metadata"),
  toolSchemas: [],
  executePowerPointToolParsed: vi.fn(),
}));

const settings = { baseUrl: "http://test.local/v1", model: "test", maxIterations: 1 };
const callbacks = (): AgentCallbacks => ({ onAssistantStart: vi.fn(), onAssistantDelta: vi.fn(), onAssistantFinal: vi.fn(), onActivity: vi.fn(), onStatus: vi.fn(), onCheckpoint: undefined });

beforeEach(() => { vi.mocked(getPresentationContext).mockResolvedValue("presentation metadata"); });

afterEach(() => { vi.restoreAllMocks(); vi.unstubAllGlobals(); vi.useRealTimers(); });

describe("bounded model-call deadlines", () => {
  it("aborts a deadline signal with TimeoutError and cleans its timer", async () => {
    vi.useFakeTimers();
    const deadline = deadlineSignal(100);
    await vi.advanceTimersByTimeAsync(101);
    expect(deadline.controller.signal.reason.name).toBe("TimeoutError");
    deadline.cleanup();
  });

  it("cancels a pending SSE reader when its signal aborts", async () => {
    let cancelled = false;
    const stream = new ReadableStream<Uint8Array>({ start() {}, cancel() { cancelled = true; } });
    const controller = new AbortController();
    const promise = readOpenAIStream(stream, callbacks(), controller.signal);
    controller.abort(new DOMException("deadline", "TimeoutError"));
    await expect(promise).rejects.toMatchObject({ name: "TimeoutError" });
    expect(cancelled).toBe(true);
  });

  it("invokes the 30-second checkpoint once before continuing", async () => {
    let nowCalls = 0; vi.spyOn(Date, "now").mockImplementation(() => nowCalls++ < 2 ? 0 : 30_001);
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(new Response("data: {\"choices\":[{\"delta\":{\"content\":\"done\"}}]}\n\ndata: [DONE]\n\n", { status: 200, headers: { "content-type": "text/event-stream" } })));
    const cb = callbacks();
    cb.onCheckpoint = vi.fn().mockResolvedValue("continue");
    await new HermesAgent().run("hello", settings, cb, new AbortController().signal);
    expect(cb.onCheckpoint).toHaveBeenCalledTimes(1);
  });
});
