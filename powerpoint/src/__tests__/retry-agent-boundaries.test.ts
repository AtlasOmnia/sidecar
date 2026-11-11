import { afterEach, describe, expect, it, vi } from "vitest";
import { HermesAgent, type AgentCallbacks } from "../agent";

vi.mock("../powerPointTools", () => ({
  getPresentationContext: vi.fn().mockResolvedValue("metadata"),
  toolSchemas: [],
  executePowerPointToolParsed: vi.fn(),
}));

const settings = { baseUrl: "http://test.local/v1", model: "test", maxIterations: 1 };
const callbacks = (): AgentCallbacks => ({ onAssistantStart: vi.fn(), onAssistantDelta: vi.fn(), onAssistantFinal: vi.fn(), onActivity: vi.fn(), onStatus: vi.fn() });
function stream(text: string): Response {
  const payload = `data: {"choices":[{"delta":{"content":${JSON.stringify(text)}}}]}\n\ndata: [DONE]\n\n`;
  return new Response(new ReadableStream({ start(controller) { controller.enqueue(new TextEncoder().encode(payload)); controller.close(); } }), { status: 200, headers: { "content-type": "text/event-stream" } });
}

afterEach(() => { vi.restoreAllMocks(); vi.unstubAllGlobals(); });

describe("bounded transient model retries", () => {
  it("retries an unparseable 503 at most once before succeeding", async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(new Response("upstream failure", { status: 503 })).mockResolvedValueOnce(stream("recovered"));
    vi.stubGlobal("fetch", fetchMock);
    const result = await new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal);
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(result.finalContent).toBe("recovered");
  });

  it("does not retry a non-transient 401 response", async () => {
    const fetchMock = vi.fn().mockResolvedValue(new Response("unauthorized", { status: 401 }));
    vi.stubGlobal("fetch", fetchMock);
    await expect(new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal)).rejects.toMatchObject({ toolError: { category: "provider", statusCode: 401 } });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("fails closed after streamed bytes instead of duplicating output with a retry", async () => {
    const partial = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new TextEncoder().encode('data: {"choices":[{"delta":{"content":"partial"}}]}\n\n'));
        setTimeout(() => controller.error(new TypeError("connection lost")), 0);
      },
    });
    const fetchMock = vi.fn().mockResolvedValue(new Response(partial, { status: 200, headers: { "content-type": "text/event-stream" } }));
    vi.stubGlobal("fetch", fetchMock);
    const cb = callbacks();
    await expect(new HermesAgent().run("hello", settings, cb, new AbortController().signal)).rejects.toMatchObject({ toolError: { category: "internal" } });
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(cb.onAssistantDelta).toHaveBeenCalledWith("partial");
  });
});


  it("retries a fetch TypeError and then succeeds", async () => {
    vi.useFakeTimers();
    const fetchMock = vi.fn().mockRejectedValueOnce(new TypeError("socket reset")).mockResolvedValueOnce(stream("recovered"));
    vi.stubGlobal("fetch", fetchMock);
    const promise = new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal);
    await vi.runAllTimersAsync();
    await expect(promise).resolves.toMatchObject({ finalContent: "recovered" });
    expect(fetchMock).toHaveBeenCalledTimes(2);
    vi.useRealTimers();
  });

  it("retries 429 and preserves the status on persistent 5xx", async () => {
    vi.useFakeTimers();
    const retryFetch = vi.fn().mockResolvedValueOnce(new Response("busy", { status: 429 })).mockResolvedValueOnce(stream("ok"));
    vi.stubGlobal("fetch", retryFetch);
    const retryPromise = new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal);
    await vi.runAllTimersAsync();
    await expect(retryPromise).resolves.toMatchObject({ finalContent: "ok" });
    expect(retryFetch).toHaveBeenCalledTimes(2);

    const persistent = vi.fn().mockImplementation(() => new Response("down", { status: 503 }));
    vi.stubGlobal("fetch", persistent);
    const persistentPromise = new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal);
    const persistentAssertion = expect(persistentPromise).rejects.toMatchObject({ toolError: { category: "provider", statusCode: 503 }, statusCode: 503 });
    await vi.runAllTimersAsync();
    await persistentAssertion;
    expect(persistent).toHaveBeenCalledTimes(3);
    vi.useRealTimers();
  });

  it("does not retry protocol, abort, timeout, or other 4xx failures", async () => {
    const cases: Array<[() => unknown, string, number?]> = [
      [() => new Response("bad", { status: 400 }), "provider", 400],
      [() => new Response("not-json", { status: 200, headers: { "content-type": "application/json" } }), "protocol"],
      [() => new DOMException("stop", "AbortError"), "cancellation"],
      [() => new DOMException("late", "TimeoutError"), "timeout"],
    ];
    for (const [makeFailure, category, statusCode] of cases) {
      const fetchMock = vi.fn().mockImplementation(() => { const failure = makeFailure(); return failure instanceof Response ? Promise.resolve(failure) : Promise.reject(failure); });
      vi.stubGlobal("fetch", fetchMock);
      await expect(new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal)).rejects.toMatchObject({ toolError: { category, ...(statusCode === undefined ? {} : { statusCode }) } });
      expect(fetchMock).toHaveBeenCalledTimes(1);
    }
  });

  it("uses one stream:false fallback", async () => {
    const first = new Response(null, { status: 200, headers: { "content-type": "text/event-stream" } });
    const fetchMock = vi.fn().mockResolvedValueOnce(first).mockResolvedValueOnce(new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "fallback" } }] }), { status: 200, headers: { "content-type": "application/json" } }));
    vi.stubGlobal("fetch", fetchMock);
    const result = await new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal);
    expect(result.finalContent).toBe("fallback");
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(JSON.parse(fetchMock.mock.calls[1][1].body).stream).toBe(false);
  });

  it("interruptible backoff removes listeners after resolve and reject", async () => {
    const { interruptibleSleep } = await import("../agent");
    vi.useFakeTimers();
    const controller = new AbortController();
    const add = vi.spyOn(controller.signal, "addEventListener");
    const remove = vi.spyOn(controller.signal, "removeEventListener");
    const pending = interruptibleSleep(10, controller.signal);
    await vi.advanceTimersByTimeAsync(10);
    await expect(pending).resolves.toBeUndefined();
    expect(remove).toHaveBeenCalled();
    add.mockClear(); remove.mockClear();
    const aborted = interruptibleSleep(100, controller.signal);
    controller.abort(new DOMException("stop", "AbortError"));
    await expect(aborted).rejects.toMatchObject({ name: "AbortError" });
    expect(remove).toHaveBeenCalled();
    vi.useRealTimers();
  });
