import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { HermesAgent, MAX_SSE_EVENT_CHARS, MAX_SSE_STREAM_BYTES, sleepHooks } from "../agent";
import * as agentModule from "../agent";
import * as wordTools from "../wordTools";

const snapshot = {
  overview: { charCount: 1, wordCount: 1, paragraphCount: 1, selection: { start: null, end: null, text: "" } },
  documentText: { text: "x", truncated: false, originalLength: 1, returnedLength: 1 },
};

const settings = { baseUrl: "http://127.0.0.1:8645/v1", model: "gpt-4.1-mini", maxIterations: 1 };
const callbacks = {
  onAssistantStart: () => undefined,
  onAssistantDelta: () => undefined,
  onAssistantFinal: () => undefined,
  onActivity: () => undefined,
  onStatus: () => undefined,
};

function response(payload: string, status = 200, contentType = "application/json"): Response {
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: new Headers({ "content-type": contentType }),
    body: null,
    text: () => Promise.resolve(payload),
  } as unknown as Response;
}

describe("HermesAgent production deadline boundaries", () => {
  it("selects the bounded size-aware server-owned turn budget from valid numeric counts", () => {
    const selectBudget = (agentModule as unknown as {
      getSizeAwareTurnTimeoutMs?: (charCount: unknown) => number;
    }).getSizeAwareTurnTimeoutMs;
    const cases: Array<[unknown, number]> = [
      [0, 120_000],
      [5_000, 120_000],
      [5_000.1, 150_000],
      [5_001, 150_000],
      [10_000, 150_000],
      [10_000.1, 180_000],
      [15_000, 180_000],
      [15_000.1, 210_000],
      [20_000, 210_000],
      [20_000.1, 240_000],
      [25_000, 240_000],
      [25_000.1, 270_000],
      [30_000, 270_000],
      [30_000.1, 300_000],
      [Number.MAX_VALUE, 300_000],
      [-1, 120_000],
      [Number.NEGATIVE_INFINITY, 120_000],
      [Number.NaN, 120_000],
      [Number.POSITIVE_INFINITY, 120_000],
      ["5_000.1", 120_000],
      [null, 120_000],
      [undefined, 120_000],
    ];

    for (const [charCount, expected] of cases) {
      const actual = selectBudget?.(charCount);
      expect(actual, `charCount=${String(charCount)}`).toBe(expected);
      expect(actual).toBeGreaterThanOrEqual(120_000);
      expect(actual).toBeLessThanOrEqual(300_000);
    }
  });

  beforeEach(() => {
    vi.useFakeTimers();
    vi.spyOn(wordTools, "getDocumentSnapshot").mockResolvedValue(snapshot);
    sleepHooks.injectedSleep = async () => undefined;
  });

  afterEach(() => {
    vi.useRealTimers();
    vi.restoreAllMocks();
    sleepHooks.injectedSleep = null;
  });

  it("times out a fetch that never resolves", async () => {
    vi.stubGlobal("fetch", vi.fn(() => new Promise<Response>(() => undefined)));
    const pending = new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal);
    const rejection = expect(pending).rejects.toMatchObject({ name: "TimeoutError" });
    await vi.advanceTimersByTimeAsync(0);
    await vi.advanceTimersByTimeAsync(20_000);
    await rejection;
  });

  it("times out a hung SSE reader and cancels the reader", async () => {
    let cancelled = false;
    const body = new ReadableStream<Uint8Array>({
      pull: () => new Promise<void>(() => undefined),
      cancel: () => { cancelled = true; },
    });
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      headers: new Headers({ "content-type": "text/event-stream" }),
      body,
    } as Response));
    const pending = new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal);
    const rejection = expect(pending).rejects.toMatchObject({ name: "TimeoutError" });
    await vi.advanceTimersByTimeAsync(0);
    await vi.advanceTimersByTimeAsync(20_000);
    await rejection;
    expect(cancelled).toBe(true);
  });

  it("rejects and cancels an oversized SSE chunk at the production boundary", async () => {
    let cancelled = false;
    const body = new ReadableStream<Uint8Array>({
      start: (controller) => controller.enqueue(new Uint8Array(MAX_SSE_STREAM_BYTES + 1)),
      cancel: () => { cancelled = true; },
    });
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      headers: new Headers({ "content-type": "text/event-stream" }),
      body,
    } as Response));

    await expect(new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal))
      .rejects.toMatchObject({ name: "ProtocolError", category: "protocol" });
    expect(cancelled).toBe(true);
  });

  it("rejects and cancels an oversized SSE event before JSON parsing", async () => {
    let cancelled = false;
    const oversizedEvent = new TextEncoder().encode(`data: {"choices":[{"delta":{"content":"${"x".repeat(MAX_SSE_EVENT_CHARS)}"}}]}\n\n`);
    const body = new ReadableStream<Uint8Array>({
      start: (controller) => controller.enqueue(oversizedEvent),
      cancel: () => { cancelled = true; },
    });
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      headers: new Headers({ "content-type": "text/event-stream" }),
      body,
    } as Response));

    await expect(new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal))
      .rejects.toMatchObject({ name: "ProtocolError", category: "protocol" });
    expect(cancelled).toBe(true);
  });

  it("contains repeated resolved zero-byte SSE chunks as a protocol failure", async () => {
    let pulls = 0;
    let cancelled = false;
    const body = new ReadableStream<Uint8Array>({
      pull: (controller) => {
        pulls += 1;
        if (pulls <= 1025) {
          controller.enqueue(new Uint8Array(0));
        } else if (pulls === 1026) {
          controller.enqueue(new Uint8Array([1]));
        } else {
          controller.close();
        }
      },
      cancel: () => { cancelled = true; },
    });
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      headers: new Headers({ "content-type": "text/event-stream" }),
      body,
    } as Response));

    await expect(new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal))
      .rejects.toMatchObject({ name: "ProtocolError", category: "protocol" });
    expect(pulls).toBeLessThan(2048);
    expect(cancelled).toBe(true);
  });

  it("contains repeated resolved zero-byte bounded-body chunks as a protocol failure", async () => {
    let pulls = 0;
    let cancelled = false;
    const body = new ReadableStream<Uint8Array>({
      pull: (controller) => {
        pulls += 1;
        if (pulls <= 1025) {
          controller.enqueue(new Uint8Array(0));
        } else if (pulls === 1026) {
          controller.enqueue(new Uint8Array([1]));
        } else {
          controller.close();
        }
      },
      cancel: () => { cancelled = true; },
    });
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
      ok: false,
      status: 500,
      headers: new Headers({ "content-type": "text/plain" }),
      body,
    } as Response));

    await expect(new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal))
      .rejects.toMatchObject({ name: "ProtocolError", category: "protocol" });
    expect(pulls).toBeLessThan(2048);
    expect(cancelled).toBe(true);
  });

  it("rejects malformed SSE data and cancels the reader", async () => {
    let cancelled = false;
    const body = new ReadableStream<Uint8Array>({
      start: (controller) => controller.enqueue(new TextEncoder().encode("data: not-json\n\n")),
      cancel: () => { cancelled = true; },
    });
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      headers: new Headers({ "content-type": "text/event-stream" }),
      body,
    } as Response));

    await expect(new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal))
      .rejects.toMatchObject({ name: "ProtocolError", category: "protocol" });
    expect(cancelled).toBe(true);
  });

  it("rejects SSE data without the choices/delta envelope", async () => {
    let cancelled = false;
    const body = new ReadableStream<Uint8Array>({
      start: (controller) => controller.enqueue(new TextEncoder().encode("data: {\"choices\":[]}\n\n")),
      cancel: () => { cancelled = true; },
    });
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      headers: new Headers({ "content-type": "text/event-stream" }),
      body,
    } as Response));

    await expect(new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal))
      .rejects.toMatchObject({ name: "ProtocolError", category: "protocol" });
    expect(cancelled).toBe(true);
  });

  it("ignores custom named SSE events between chat deltas", async () => {
    const deltas: string[] = [];
    const body = new ReadableStream<Uint8Array>({
      start: (controller) => {
        controller.enqueue(new TextEncoder().encode([
          'data: {"choices":[{"delta":{"content":"Hello"}}]}',
          "",
          "event: hermes.tool.progress",
          'data: {"tool":"web_search","status":"running"}',
          "",
          'data: {"choices":[{"delta":{"content":" world"}}]}',
          "",
          "data: [DONE]",
          "",
          "",
        ].join("\n")));
        controller.close();
      },
    });
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      headers: new Headers({ "content-type": "text/event-stream" }),
      body,
    } as Response));

    const result = await new HermesAgent().run(
      "summarize",
      settings,
      { ...callbacks, onAssistantDelta: (delta) => deltas.push(delta) },
      new AbortController().signal,
    );

    expect(result.finalContent).toBe("Hello world");
    expect(deltas).toEqual(["Hello", " world"]);
  });

  it("times out a hung non-streaming body read", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      headers: new Headers({ "content-type": "application/json" }),
      body: null,
      text: () => new Promise<string>(() => undefined),
    } as Response));
    const pending = new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal);
    const rejection = expect(pending).rejects.toMatchObject({ name: "TimeoutError" });
    await vi.advanceTimersByTimeAsync(0);
    await vi.advanceTimersByTimeAsync(20_000);
    await rejection;
  });

  it("keeps user Stop as AbortError when it wins before the deadline", async () => {
    vi.stubGlobal("fetch", vi.fn(() => new Promise<Response>(() => undefined)));
    const controller = new AbortController();
    const reason = new DOMException("Stopped by user", "AbortError");
    const pending = new HermesAgent().run("summarize", settings, callbacks, controller.signal);
    await vi.advanceTimersByTimeAsync(0);
    controller.abort(reason);
    await expect(pending).rejects.toBe(reason);
    await vi.advanceTimersByTimeAsync(20_000);
  });

  it("bounds a hung startup snapshot by the total turn deadline", async () => {
    vi.spyOn(wordTools, "getDocumentSnapshot").mockReturnValue(new Promise<typeof snapshot>(() => undefined));
    const pending = new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal);
    const rejection = expect(pending).rejects.toMatchObject({ name: "TimeoutError" });
    await vi.advanceTimersByTimeAsync(0);
    await vi.advanceTimersByTimeAsync(120_000);
    await rejection;
  });

  it("bounds a hung Word tool by the total turn deadline", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(response(JSON.stringify({
      choices: [{ message: { content: "", tool_calls: [{ id: "1", type: "function", function: { name: "insert_text", arguments: JSON.stringify({ text: "x" }) } }] } }],
    }))));
    vi.spyOn(wordTools, "executeWordToolParsed").mockReturnValue(new Promise<string>(() => undefined));
    const pending = new HermesAgent().run("insert text", settings, callbacks, new AbortController().signal);
    const rejection = expect(pending).rejects.toMatchObject({ name: "TimeoutError" });
    await vi.advanceTimersByTimeAsync(0);
    await vi.advanceTimersByTimeAsync(120_000);
    await rejection;
  });

  it("retries a transient network error once with injected abortible sleep", async () => {
    const fetchMock = vi.fn()
      .mockRejectedValueOnce(new TypeError("network down"))
      .mockResolvedValueOnce(response(JSON.stringify({ choices: [{ message: { content: "done" } }] })));
    vi.stubGlobal("fetch", fetchMock);

    const result = await new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal);
    expect(result.finalContent).toBe("done");
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it("does not retry ordinary provider 4xx responses", async () => {
    const fetchMock = vi.fn().mockResolvedValue(response("bad request", 400));
    vi.stubGlobal("fetch", fetchMock);

    await expect(new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal)).rejects.toThrow("Model request failed");
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("keeps network and provider failures free of URL and body secrets", async () => {
    const secretSettings = {
      ...settings,
      baseUrl: "http://user:url-password@example.invalid/v1?query-secret=do-not-show",
    };
    vi.stubGlobal("fetch", vi.fn().mockRejectedValue(new TypeError("raw-provider-exception-secret")));
    const networkFailure = await new HermesAgent().run("summarize", secretSettings, callbacks, new AbortController().signal)
      .then(() => new Error("expected network failure"), (error: unknown) => error as Error);
    expect(networkFailure.message).not.toContain("url-password");
    expect(networkFailure.message).not.toContain("query-secret");
    expect(networkFailure.message).not.toContain("raw-provider-exception-secret");
    expect(networkFailure.message).toContain("Could not reach the Hermes model service");

    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(response("provider-body-secret", 500)));
    const providerFailure = await new HermesAgent().run("summarize", secretSettings, callbacks, new AbortController().signal)
      .then(() => new Error("expected provider failure"), (error: unknown) => error as Error);
    expect(providerFailure.message).not.toContain("provider-body-secret");
    expect(providerFailure.message).toContain("HTTP 500");
  });

  it.each([429, 500])("retries transient provider status %s once", async (status) => {
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response("retry", status))
      .mockResolvedValueOnce(response(JSON.stringify({ choices: [{ message: { content: "recovered" } }] })));
    vi.stubGlobal("fetch", fetchMock);

    const result = await new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal);

    expect(result.finalContent).toBe("recovered");
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it("caps transient retries at three attempts with 1s and 2s abortible backoff", async () => {
    const fetchMock = vi.fn().mockRejectedValue(new TypeError("network down"));
    const sleeps: number[] = [];
    sleepHooks.injectedSleep = async (ms) => {
      sleeps.push(ms);
    };
    vi.stubGlobal("fetch", fetchMock);

    await expect(new HermesAgent().run("summarize", settings, callbacks, new AbortController().signal))
      .rejects.toThrow("Model request failed after 3 attempts");

    expect(fetchMock).toHaveBeenCalledTimes(3);
    expect(sleeps).toEqual([1000, 2000]);
  });

  it("stops at the checkpoint when the production callback chooses Stop", async () => {
    vi.spyOn(wordTools, "getDocumentSnapshot").mockImplementation(() => new Promise((resolve) => {
      setTimeout(() => resolve(snapshot), 30_000);
    }));
    const onCheckpoint = vi.fn().mockResolvedValue(false);
    const pending = new HermesAgent().run(
      "summarize",
      settings,
      { ...callbacks, onCheckpoint },
      new AbortController().signal,
    );

    await vi.advanceTimersByTimeAsync(30_000);
    const result = await pending;

    expect(onCheckpoint).toHaveBeenCalledTimes(1);
    expect(result.stopReason).toBe("checkpoint");
    expect(result.finalContent).toContain("30-second checkpoint");
  });

  it("preserves user AbortError while the production checkpoint is open", async () => {
    vi.spyOn(wordTools, "getDocumentSnapshot").mockImplementation(() => new Promise((resolve) => {
      setTimeout(() => resolve(snapshot), 30_000);
    }));
    const controller = new AbortController();
    const reason = new DOMException("Stopped by the user.", "AbortError");
    const onCheckpoint = vi.fn().mockImplementation(() => new Promise<boolean>(() => undefined));
    const pending = new HermesAgent().run("summarize", settings, { ...callbacks, onCheckpoint }, controller.signal);

    await vi.advanceTimersByTimeAsync(30_000);
    controller.abort(reason);

    await expect(pending).rejects.toBe(reason);
  });

  it("preserves TimeoutError while the production checkpoint is open", async () => {
    vi.spyOn(wordTools, "getDocumentSnapshot").mockImplementation(() => new Promise((resolve) => {
      setTimeout(() => resolve(snapshot), 30_000);
    }));
    const onCheckpoint = vi.fn().mockImplementation(() => new Promise<boolean>(() => undefined));
    const pending = new HermesAgent().run("summarize", settings, { ...callbacks, onCheckpoint }, new AbortController().signal);
    const rejection = expect(pending).rejects.toMatchObject({ name: "TimeoutError" });

    await vi.advanceTimersByTimeAsync(120_000);

    await rejection;
  });
});
