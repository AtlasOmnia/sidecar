import { describe, expect, it, vi } from "vitest";
import { OfficeCompanionBridge } from "../officeBridge";

const SESSION_ID = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const SESSION_KEY = "s".repeat(64);
const PANE_ID = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
const TICKET_ID = "cccccccc-cccc-4ccc-8ccc-cccccccccccc";
const REQUEST_ID = "dddddddd-dddd-4ddd-8ddd-dddddddddddd";

function response(status: number, body?: unknown): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
  });
}

describe("OfficeCompanionBridge", () => {
  it("invokes the default fetch with the global Window-compatible receiver", async () => {
    const originalFetch = globalThis.fetch;
    const fetchWithReceiver = vi.fn(function (this: typeof globalThis) {
      if (this !== globalThis) {
        throw new TypeError("Can only call Window.fetch on instances of Window");
      }
      return Promise.resolve(response(201, {
        ok: true,
        sessionId: SESSION_ID,
        sessionKey: SESSION_KEY,
        protocolVersion: "1.0",
        expiresInSeconds: 60,
      }));
    });
    globalThis.fetch = fetchWithReceiver as typeof fetch;
    try {
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "word",
        documentId: async () => "doc-1",
        paneId: PANE_ID,
        executeTool: vi.fn(),
      });
      await bridge.connect();
      expect(fetchWithReceiver).toHaveBeenCalledTimes(1);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it("binds the default WebKit timer functions to the global Window-compatible receiver", async () => {
    const originalFetch = globalThis.fetch;
    const originalSetInterval = globalThis.setInterval;
    const originalClearInterval = globalThis.clearInterval;
    const fetchWithReceiver = vi.fn(function (this: typeof globalThis, _input: RequestInfo | URL, init?: RequestInit) {
      if (this !== globalThis) throw new TypeError("invalid fetch receiver");
      return Promise.resolve(init?.method === "DELETE" ? response(204) : response(201, {
        ok: true,
        sessionId: SESSION_ID,
        sessionKey: SESSION_KEY,
        protocolVersion: "1.0",
        expiresInSeconds: 60,
      }));
    });
    const timerHandles: number[] = [];
    const setIntervalWithReceiver = vi.fn(function (this: typeof globalThis) {
      if (this !== globalThis) throw new TypeError("Can only call Window.setInterval on instances of Window");
      const handle = timerHandles.length + 1;
      timerHandles.push(handle);
      return handle;
    });
    const clearIntervalWithReceiver = vi.fn(function (this: typeof globalThis) {
      if (this !== globalThis) throw new TypeError("invalid clearInterval receiver");
    });
    globalThis.fetch = fetchWithReceiver as typeof fetch;
    globalThis.setInterval = setIntervalWithReceiver as unknown as typeof setInterval;
    globalThis.clearInterval = clearIntervalWithReceiver as unknown as typeof clearInterval;
    try {
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "word", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(),
      });
      await bridge.start();
      expect(setIntervalWithReceiver).toHaveBeenCalledTimes(2);
      await bridge.stop();
      expect(clearIntervalWithReceiver).toHaveBeenCalledTimes(2);
    } finally {
      globalThis.fetch = originalFetch;
      globalThis.setInterval = originalSetInterval;
      globalThis.clearInterval = originalClearInterval;
    }
  });

  it("registers an exact document/pane session and keeps the key in memory", async () => {
    const fetchMock = vi.fn(async (_url: string, init?: RequestInit) =>
      response(201, {
        ok: true,
        sessionId: SESSION_ID,
        sessionKey: SESSION_KEY,
        protocolVersion: "1.0",
        expiresInSeconds: 60,
      }),
    );
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "word",
      documentId: async () => "doc-1",
      paneId: PANE_ID,
      executeTool: vi.fn(),
      fetchImpl: fetchMock as typeof fetch,
    });

    const session = await bridge.connect();
    expect(session.sessionId).toBe(SESSION_ID);
    expect(bridge.sessionId).toBe(SESSION_ID);
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe("/word/session");
    expect(JSON.parse(String(init.body))).toEqual({
      protocolVersion: "1.0",
      documentId: "doc-1",
      paneId: PANE_ID,
    });
    expect(JSON.stringify(bridge.publicState)).not.toContain(SESSION_KEY);
  });

  it("deduplicates concurrent registration attempts", async () => {
    let resolveFetch!: (value: Response) => void;
    const fetchMock = vi.fn(() => new Promise<Response>((resolve) => {
      resolveFetch = resolve;
    }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "word", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });
    const first = bridge.connect();
    const second = bridge.connect();
    await Promise.resolve();
    resolveFetch(response(201, {
      ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
      protocolVersion: "1.0", expiresInSeconds: 60,
    }));
    const [a, b] = await Promise.all([first, second]);
    expect(a.sessionId).toBe(SESSION_ID);
    expect(b.sessionId).toBe(SESSION_ID);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("does not leak a session when stop races registration", async () => {
    let resolveRegistration!: (value: Response) => void;
    const fetchMock = vi.fn()
      .mockImplementationOnce(() => new Promise<Response>((resolve) => {
        resolveRegistration = resolve;
      }))
      .mockResolvedValueOnce(response(204));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "word", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });
    const connecting = bridge.connect();
    await Promise.resolve();
    await Promise.resolve();
    const stopping = bridge.stop();
    resolveRegistration(response(201, {
      ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
      protocolVersion: "1.0", expiresInSeconds: 60,
    }));
    await connecting;
    await stopping;
    expect(bridge.sessionId).toBeNull();
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it("claims one pending ticket, executes canonical validation, and submits JSON result", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: [[1]] }));
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, {
        ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
        protocolVersion: "1.0", expiresInSeconds: 60,
      }))
      .mockResolvedValueOnce(response(200, {
        ticketId: TICKET_ID,
        requestId: REQUEST_ID,
        tool: "read_document_text",
        arguments: { maxLength: 100 },
      }))
      .mockResolvedValueOnce(response(200, { ok: true, ticketId: TICKET_ID }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "word",
      documentId: async () => "doc-1",
      paneId: PANE_ID,
      executeTool,
      fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(executeTool).toHaveBeenCalledWith("read_document_text", { maxLength: 100 });
    const [resultUrl, resultInit] = fetchMock.mock.calls[2] as [string, RequestInit];
    expect(resultUrl).toBe("/word/result");
    expect(JSON.parse(String(resultInit.body))).toEqual({
      ticketId: TICKET_ID,
      result: { ok: true, data: [[1]] },
    });
  });

  it("retries a cached result without re-executing a mutating ticket", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: { written: true } }));
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, {
        ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
        protocolVersion: "1.0", expiresInSeconds: 60,
      }))
      .mockResolvedValueOnce(response(200, {
        ticketId: TICKET_ID, requestId: REQUEST_ID,
        tool: "insert_text", arguments: { text: "hello", location: "end" },
      }))
      .mockRejectedValueOnce(new TypeError("connection reset after write"))
      .mockResolvedValueOnce(response(200, { ok: true, ticketId: TICKET_ID }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "word", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();
    await expect(bridge.pollOnce()).rejects.toThrow("connection reset");
    expect(await bridge.pollOnce()).toBe(true);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[3][0]).toBe("/word/result");
  });

  it("drops a terminally rejected cached result and resumes polling", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: { written: true } }));
    let pendingCalls = 0;
    const fetchMock = vi.fn(async (url: string) => {
      if (url.endsWith("/word/session")) {
        return response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url.endsWith("/word/pending")) {
        pendingCalls += 1;
        return pendingCalls === 1
          ? response(200, {
              ticketId: TICKET_ID,
              requestId: REQUEST_ID,
              tool: "insert_text",
              arguments: { text: "slow write", location: "end" },
            })
          : response(204);
      }
      if (url.endsWith("/word/result")) {
        return response(409, { ok: false, error: "ticket is no longer claimed" });
      }
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "word", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as typeof fetch,
    });

    await bridge.connect();
    await expect(bridge.pollOnce()).rejects.toThrow("ticket is no longer claimed");
    expect(await bridge.pollOnce()).toBe(false);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(pendingCalls).toBe(2);
  });

  it("does not execute a malformed or wrong-session ticket", async () => {
    const executeTool = vi.fn();
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, {
        ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
        protocolVersion: "1.0", expiresInSeconds: 60,
      }))
      .mockResolvedValueOnce(response(200, {
        ticketId: "not-a-uuid", requestId: REQUEST_ID,
        tool: "read_document_text", arguments: {},
      }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "word", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();
    await expect(bridge.pollOnce()).rejects.toThrow("ticketId");
    expect(executeTool).not.toHaveBeenCalled();
  });

  it("rejects companion responses by UTF-8 byte length, not JavaScript characters", async () => {
    const oversizedError = "😀".repeat(2_700_000);
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, {
        ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
        protocolVersion: "1.0", expiresInSeconds: 60,
      }))
      .mockResolvedValueOnce(response(500, { error: oversizedError }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "word", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });

    await bridge.connect();
    await expect(bridge.heartbeatOnce()).rejects.toThrow("response is too large");
  });

  it("rejects an oversized serialized result before submission or caching", async () => {
    let resultSubmissions = 0;
    const executeTool = vi.fn(async () => JSON.stringify({ data: "😀".repeat(2_700_000) }));
    const fetchMock = vi.fn(async (url: string) => {
      if (url.endsWith("/word/session")) {
        return response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url.endsWith("/word/pending")) {
        return response(200, {
          ticketId: TICKET_ID, requestId: REQUEST_ID,
          tool: "read_document_text", arguments: {},
        });
      }
      if (url.endsWith("/word/result")) {
        resultSubmissions += 1;
        return response(200, { ok: true, ticketId: TICKET_ID });
      }
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "word", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as typeof fetch,
    });

    await bridge.connect();
    await expect(bridge.pollOnce()).rejects.toThrow("result is too large");
    expect(resultSubmissions).toBe(0);
  });

  it("bounds the retry cache by exact serialized bytes and entry count", () => {
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "word", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: vi.fn() as unknown as typeof fetch,
    });
    const internals = bridge as unknown as {
      cacheCompletedResult: (ticketId: string, result: unknown) => void;
      completedResults: Map<string, unknown>;
      completedResultBytes: number;
    };

    for (let index = 0; index < 4; index += 1) {
      internals.cacheCompletedResult(`ticket-${index}`, { ok: true, index });
    }
    expect(internals.completedResults.size).toBe(4);
    expect(() => internals.cacheCompletedResult("ticket-5", { ok: true })).toThrow("capacity");

    const byteBounded = new OfficeCompanionBridge({
      hostPrefix: "word", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: vi.fn() as unknown as typeof fetch,
    }) as unknown as {
      cacheCompletedResult: (ticketId: string, result: unknown) => void;
      completedResultBytes: number;
    };
    byteBounded.cacheCompletedResult("ticket-a", { data: "x".repeat(6 * 1024 * 1024) });
    expect(() => byteBounded.cacheCompletedResult(
      "ticket-b", { data: "x".repeat(6 * 1024 * 1024) },
    )).toThrow("capacity");
    expect(byteBounded.completedResultBytes).toBeLessThanOrEqual(12 * 1024 * 1024);
  });

  it("unregisters only its authenticated session on stop", async () => {
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, {
        ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
        protocolVersion: "1.0", expiresInSeconds: 60,
      }))
      .mockResolvedValueOnce(response(204));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "word", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();
    await bridge.stop();
    const [url, init] = fetchMock.mock.calls[1] as [string, RequestInit];
    expect(url).toBe("/word/session");
    expect(init.method).toBe("DELETE");
    expect((init.headers as Record<string, string>)["X-Hermes-Session-Id"]).toBe(SESSION_ID);
    expect(bridge.sessionId).toBeNull();
  });
});
