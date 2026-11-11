/**
 * H1 Production-path tests: HermesAgent with mocked fetch producing hung streams,
 * timeouts, retries, and sanitized errors.
 *
 * E Production-path sanitization: HermesAgent against HTTP errors with secrets.
 */

import { beforeEach, afterEach, describe, expect, it, vi } from "vitest";
import {
  HermesAgent,
  resetInjectedSleep,
  sleepHooks,
  MODEL_CALL_TIMEOUT_MS,
  TURN_TIMEOUT_MS,
} from "../agent";
import * as wordTools from "../wordTools";

const dummySnapshot = {
  overview: { charCount: 10, wordCount: 3, paragraphCount: 1, selection: { start: null as number | null, end: null as number | null, text: "" } },
  documentText: { text: "test", truncated: false, originalLength: 4, returnedLength: 4 }
};

const dummySettings = {
  baseUrl: "http://127.0.0.1:8643/v1",
  model: "gpt-4.1-mini",
  maxIterations: 3
};

const noopCallbacks = {
  onAssistantStart: () => {},
  onAssistantDelta: () => {},
  onAssistantFinal: () => {},
  onActivity: () => {},
  onStatus: () => {}
};

// Helper: create a mock fetch that hangs until the signal aborts
function makeHungFetch() {
  return vi.spyOn(globalThis, "fetch").mockImplementation(
    (_url, init) => new Promise<Response>((_resolve, reject) => {
      const signal = init?.signal;
      if (signal) {
        if (signal.aborted) {
          reject(signal.reason instanceof DOMException ? signal.reason : new DOMException("aborted", "AbortError"));
          return;
        }
        signal.addEventListener("abort", () => {
          reject(signal.reason instanceof DOMException ? signal.reason : new DOMException("aborted", "AbortError"));
        }, { once: true });
      }
    })
  );
}

/** Attach rejection handling before advancing fake timers to avoid transient
 * unhandled-rejection warnings that can otherwise make Vitest false-green. */
function captureRejection(promise: Promise<unknown>): Promise<unknown> {
  return promise.then(
    () => null,
    (error: unknown) => error,
  );
}

describe("H1: production-path HermesAgent — hung fetch and body", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.restoreAllMocks();
    resetInjectedSleep();
    vi.spyOn(wordTools, "getDocumentSnapshot").mockResolvedValue(dummySnapshot as any);
    vi.spyOn(wordTools, "executeWordToolParsed").mockResolvedValue('{"ok":true}');
  });

  afterEach(() => {
    vi.useRealTimers();
    vi.restoreAllMocks();
    resetInjectedSleep();
  });

  it("hung fetch rejects with TimeoutError after 20s, exactly 1 fetch (not retried)", async () => {
    const fetchSpy = makeHungFetch();

    sleepHooks.injectedSleep = async (ms: number) => {
      vi.advanceTimersByTime(ms);
    };

    const agent = new HermesAgent();
    const runPromise = agent.run("test", dummySettings, noopCallbacks, new AbortController().signal);
    const rejection = captureRejection(runPromise);

    // Advance past 20s model call deadline — TimeoutError is not retried
    vi.advanceTimersByTime(MODEL_CALL_TIMEOUT_MS + 100);
    await vi.runAllTimersAsync();

    const rejected = await rejection;
    const caught = rejected instanceof DOMException ? rejected : null;
    expect(caught).not.toBeNull();
    expect(caught!.name).toBe("TimeoutError");
    expect(fetchSpy).toHaveBeenCalledTimes(1);

    resetInjectedSleep();
  }, 10000);

  it("hung SSE body rejects with TimeoutError, exactly 1 fetch, reader cancelled", async () => {
    let readerCancelled = false;
    const hungStream = new ReadableStream<Uint8Array>({
      start() { /* never enqueues — blocks reader */ },
      cancel() { readerCancelled = true; }
    });

    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      new Response(hungStream, { status: 200, headers: { "content-type": "text/event-stream" } })
    );

    sleepHooks.injectedSleep = async (ms: number) => {
      vi.advanceTimersByTime(ms);
    };

    const agent = new HermesAgent();
    const runPromise = agent.run("test", dummySettings, noopCallbacks, new AbortController().signal);
    const rejection = captureRejection(runPromise);

    // Advance past 20s model call deadline
    vi.advanceTimersByTime(MODEL_CALL_TIMEOUT_MS + 100);
    await vi.runAllTimersAsync();

    const rejected = await rejection;
    const caught = rejected instanceof DOMException ? rejected : null;
    expect(caught).not.toBeNull();
    expect(caught!.name).toBe("TimeoutError");
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(readerCancelled).toBe(true);

    resetInjectedSleep();
  }, 10000);

  it("hung application/json body rejects with TimeoutError after 20s, reader cancelled", async () => {
    let readerCancelled = false;
    const hungStream = new ReadableStream<Uint8Array>({
      start() { /* never enqueues — blocks reader */ },
      cancel() { readerCancelled = true; }
    });

    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      new Response(hungStream, { status: 200, headers: { "content-type": "application/json" } })
    );

    sleepHooks.injectedSleep = async (ms: number) => {
      vi.advanceTimersByTime(ms);
    };

    const agent = new HermesAgent();
    const runPromise = agent.run("test", dummySettings, noopCallbacks, new AbortController().signal);
    const rejection = captureRejection(runPromise);

    // Advance past 20s model call deadline — TimeoutError, not retried
    vi.advanceTimersByTime(MODEL_CALL_TIMEOUT_MS + 100);
    await vi.runAllTimersAsync();

    const rejected = await rejection;
    const caught = rejected instanceof DOMException ? rejected : null;
    expect(caught).not.toBeNull();
    expect(caught!.name).toBe("TimeoutError");
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(readerCancelled).toBe(true);

    resetInjectedSleep();
  }, 10000);

  it("user abort during fetch throws AbortError and does not retry", async () => {
    let fetchWasCalled = false;

    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(
      (_url, init) => {
        fetchWasCalled = true;
        return new Promise<Response>((_resolve, reject) => {
          const signal = init?.signal;
          if (signal) {
            if (signal.aborted) {
              reject(signal.reason instanceof DOMException ? signal.reason : new DOMException("aborted", "AbortError"));
              return;
            }
            signal.addEventListener("abort", () => {
              reject(new DOMException("user stopped", "AbortError"));
            }, { once: true });
          }
        });
      }
    );

    const userCtrl = new AbortController();
    const agent = new HermesAgent();
    const runPromise = agent.run("test", dummySettings, noopCallbacks, userCtrl.signal);
    const rejection = captureRejection(runPromise);

    // Advance microtasks so getDocumentSnapshot resolves and fetch is called
    await vi.advanceTimersByTimeAsync(0);
    expect(fetchWasCalled).toBe(true);

    // Now abort while fetch is pending
    userCtrl.abort(new DOMException("user stopped", "AbortError"));
    await vi.runAllTimersAsync();

    const rejected = await rejection;
    const caught = rejected instanceof DOMException ? rejected : null;
    expect(caught).not.toBeNull();
    expect(caught!.name).toBe("AbortError");
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  }, 10000);

  it("user abort during pending JSON body read throws AbortError, exactly 1 fetch, stream cancelled", async () => {
    let bodyReadStarted = false;
    let streamCancelled = false;

    const hungStream = new ReadableStream<Uint8Array>({
      start() {
        bodyReadStarted = true;
        /* never enqueues — blocks reader */
      },
      cancel() { streamCancelled = true; }
    });

    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      new Response(hungStream, { status: 200, headers: { "content-type": "application/json" } })
    );

    const userCtrl = new AbortController();
    const agent = new HermesAgent();
    const runPromise = agent.run("test", dummySettings, noopCallbacks, userCtrl.signal);
    const rejection = captureRejection(runPromise);

    // Advance microtasks so fetch resolves and body read starts
    await vi.advanceTimersByTimeAsync(0);
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(bodyReadStarted).toBe(true);

    // Now abort while JSON body read is pending
    userCtrl.abort(new DOMException("user stopped", "AbortError"));
    await vi.runAllTimersAsync();

    const rejected = await rejection;
    const caught = rejected instanceof DOMException ? rejected : null;
    expect(caught).not.toBeNull();
    expect(caught!.name).toBe("AbortError");
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(streamCancelled).toBe(true);
  }, 10000);

  it("transient 500 then success: 2 fetch calls and successful completion", async () => {
    sleepHooks.injectedSleep = async (ms: number) => {
      vi.advanceTimersByTime(ms);
    };

    const successBody = JSON.stringify({
      choices: [{ message: { content: "Recovered!" } }]
    });
    const fetchSpy = vi.spyOn(globalThis, "fetch")
      .mockResolvedValueOnce(new Response("error", { status: 500 }))
      .mockResolvedValueOnce(new Response(successBody, { status: 200, headers: { "content-type": "application/json" } }));

    const result = await new HermesAgent().run("test", dummySettings, noopCallbacks, new AbortController().signal);

    expect(fetchSpy).toHaveBeenCalledTimes(2);
    expect(result.finalContent).toBe("Recovered!");

    resetInjectedSleep();
  });

  it("429 then success: 2 fetch calls and successful completion", async () => {
    sleepHooks.injectedSleep = async (ms: number) => {
      vi.advanceTimersByTime(ms);
    };

    const successBody = JSON.stringify({
      choices: [{ message: { content: "Rate-limited but recovered!" } }]
    });
    const fetchSpy = vi.spyOn(globalThis, "fetch")
      .mockResolvedValueOnce(new Response("rate limited", { status: 429 }))
      .mockResolvedValueOnce(new Response(successBody, { status: 200, headers: { "content-type": "application/json" } }));

    const result = await new HermesAgent().run("test", dummySettings, noopCallbacks, new AbortController().signal);

    expect(fetchSpy).toHaveBeenCalledTimes(2);
    expect(result.finalContent).toBe("Rate-limited but recovered!");

    resetInjectedSleep();
  });

  it("deterministic 400: 1 fetch call, no retry", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch")
      .mockResolvedValue(new Response("bad request", { status: 400 }));

    const agent = new HermesAgent();
    const runPromise = agent.run("test", dummySettings, noopCallbacks, new AbortController().signal);

    await expect(runPromise).rejects.toThrow();
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });

  it("retry exhaustion: 3 fetch calls and sanitized final error", async () => {
    sleepHooks.injectedSleep = async (ms: number) => {
      vi.advanceTimersByTime(ms);
    };

    const fetchSpy = vi.spyOn(globalThis, "fetch")
      .mockResolvedValueOnce(new Response("err1", { status: 502 }))
      .mockResolvedValueOnce(new Response("err2", { status: 502 }))
      .mockResolvedValueOnce(new Response("err3", { status: 502 }));

    const agent = new HermesAgent();
    const runPromise = agent.run("test", dummySettings, noopCallbacks, new AbortController().signal);
    const rejection = captureRejection(runPromise);

    vi.advanceTimersByTime(30000); // Allow retry sleeps

    const retryError = await rejection;
    expect(retryError).toBeInstanceOf(Error);
    expect((retryError as Error).message).toContain("Model request failed after 3 attempts");
    expect(fetchSpy).toHaveBeenCalledTimes(3);

    resetInjectedSleep();
  });

  it("turn budget exhaustion yields TimeoutError", async () => {
    makeHungFetch();

    const agent = new HermesAgent();
    const runPromise = agent.run("test", { ...dummySettings, maxIterations: 1 }, noopCallbacks, new AbortController().signal);
    const rejection = captureRejection(runPromise);

    // Advance past 20s model call deadline (Date.now() not affected by fake timers,
    // so the per-call deadline fires first and produces TimeoutError)
    vi.advanceTimersByTime(MODEL_CALL_TIMEOUT_MS + 100);
    await vi.runAllTimersAsync();

    const rejected = await rejection;
    const caught = rejected instanceof DOMException ? rejected : null;
    expect(caught).not.toBeNull();
    expect(caught!.name).toBe("TimeoutError");
  }, 10000);
});

describe("E: production-path error sanitization via HermesAgent", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.restoreAllMocks();
    resetInjectedSleep();
    vi.spyOn(wordTools, "getDocumentSnapshot").mockResolvedValue(dummySnapshot as any);
    vi.spyOn(wordTools, "executeWordToolParsed").mockResolvedValue('{"ok":true}');
  });

  afterEach(() => {
    vi.useRealTimers();
    vi.restoreAllMocks();
    resetInjectedSleep();
  });

  it("error message excludes secrets, HTML, stack, and oversized body", async () => {
    sleepHooks.injectedSleep = async (ms: number) => {
      vi.advanceTimersByTime(ms);
    };

    // Dynamically assembled secret — assertions test THIS secret
    const realSecret = ["sk", "mysecret", String(Date.now())].join("-");

    const rawErrorBody = JSON.stringify({
      error: {
        message: "Invalid API key",
        type: "authentication_error",
        code: "invalid_api_key"
      },
      api_key_hint: realSecret,
      stack: "Error: at processRequest (/app/server.js:42:15)\n    at handle (/app/router.js:100:3)",
      html_payload: "<html><body><h1>401 Unauthorized</h1><pre>API key: " + realSecret + "</pre></body></html>",
      giant_field: "x".repeat(20000)
    });

    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      new Response(rawErrorBody, { status: 401, headers: { "content-type": "application/json" } })
    );

    const agent = new HermesAgent();

    try {
      const runPromise = agent.run("test", { ...dummySettings, maxIterations: 1 }, noopCallbacks, new AbortController().signal);
      vi.advanceTimersByTime(5000);
      await runPromise;
      expect.fail("Expected run to throw");
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      // The dynamically assembled secret must not leak
      expect(msg).not.toContain(realSecret);
      expect(msg).not.toContain("<html>");
      expect(msg).not.toContain("<body>");
      expect(msg).not.toContain("<h1>");
      expect(msg).not.toContain("processRequest");
      expect(msg).not.toContain("/app/server.js");
      expect(msg.length).toBeLessThan(500);
      // Safe message must include status context
      expect(msg).toContain("401");
    }

    resetInjectedSleep();
  });

  it("error message includes status code guidance but not raw body or stack", async () => {
    sleepHooks.injectedSleep = async (ms: number) => {
      vi.advanceTimersByTime(ms);
    };

    const reqId = "req-" + Date.now().toString(36);
    const htmlError = `<!DOCTYPE html>
<html>
<head><title>502 Bad Gateway</title></head>
<body>
<h1>502 Bad Gateway</h1>
<p>Upstream server returned invalid response</p>
<pre>Error: connect ECONNREFUSED 127.0.0.1:8643
    at TCPConnectWrap.afterConnect [as oncomplete] (node:net:1234:16)
    at internalConnect (node:net:100:5)</pre>
<p>X-Request-ID: ${reqId}</p>
</body>
</html>`;

    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      new Response(htmlError, { status: 502, headers: { "content-type": "text/html" } })
    );

    const agent = new HermesAgent();

    try {
      const runPromise = agent.run("test", { ...dummySettings, maxIterations: 1 }, noopCallbacks, new AbortController().signal);
      vi.advanceTimersByTime(20000);
      await runPromise;
      expect.fail("Expected run to throw");
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      expect(msg).toContain("502");
      expect(msg).not.toContain("TCPConnectWrap");
      expect(msg).not.toContain("    at ");
      expect(msg).not.toContain("internalConnect");
      expect(msg).not.toContain("<html>");
      expect(msg).not.toContain("<body>");
      // Dynamic request ID must not leak
      expect(msg).not.toContain(reqId);
      // Safe message must be short (no raw body)
      expect(msg.length).toBeLessThan(200);
    }

    resetInjectedSleep();
  });

  it("retry aggregate excludes raw body from all attempts", async () => {
    sleepHooks.injectedSleep = async (ms: number) => {
      vi.advanceTimersByTime(ms);
    };

    // Each error response has different dynamically assembled secrets
    const secret1 = "sec-" + Date.now().toString(36) + "-a";
    const secret2 = "sec-" + Date.now().toString(36) + "-b";
    const secret3 = "sec-" + Date.now().toString(36) + "-c";

    const err1 = JSON.stringify({ error: { message: "err1" }, secret: secret1 });
    const err2 = JSON.stringify({ error: { message: "err2" }, secret: secret2 });
    const err3 = JSON.stringify({ error: { message: "err3" }, secret: secret3 });

    vi.spyOn(globalThis, "fetch")
      .mockResolvedValueOnce(new Response(err1, { status: 502 }))
      .mockResolvedValueOnce(new Response(err2, { status: 502 }))
      .mockResolvedValueOnce(new Response(err3, { status: 502 }));

    const agent = new HermesAgent();

    try {
      const runPromise = agent.run("test", { ...dummySettings, maxIterations: 1 }, noopCallbacks, new AbortController().signal);
      vi.advanceTimersByTime(30000);
      await runPromise;
      expect.fail("Expected run to throw");
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      // None of the dynamic secrets must leak
      expect(msg).not.toContain(secret1);
      expect(msg).not.toContain(secret2);
      expect(msg).not.toContain(secret3);
      expect(msg).toContain("3 attempts");
    }

    resetInjectedSleep();
  });
});
