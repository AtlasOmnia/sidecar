import { describe, expect, it, vi } from "vitest";
import { OfficeCompanionBridge } from "../officeBridge";

const SESSION_ID = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const SESSION_ID_B = "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee";
const SESSION_KEY = "s".repeat(64);
const SESSION_KEY_B = "t".repeat(64);
const PANE_ID = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
const PANE_ID_B = "ffffffff-ffff-4fff-8fff-ffffffffffff";
const TICKET_ID = "cccccccc-cccc-4ccc-8ccc-cccccccccccc";
const TICKET_ID_B = "99999999-9999-4999-8999-999999999999";
const REQUEST_ID = "dddddddd-dddd-4ddd-8ddd-dddddddddddd";
const REQUEST_ID_B = "88888888-8888-4888-8888-888888888888";

function response(status: number, body?: unknown): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
  });
}

async function flushMicrotasks(count = 30): Promise<void> {
  for (let index = 0; index < count; index += 1) await Promise.resolve();
}

type CachedTestResult = {
  requestId: string;
  ticketId: string;
  sessionId: string;
  result: unknown;
  serializedBody: string;
  resultAttempts: number;
  retryNotBefore: number | null;
};

function cachedResultsOf(bridge: OfficeCompanionBridge): Map<string, CachedTestResult> {
  return (bridge as unknown as { completedResults: Map<string, CachedTestResult> }).completedResults;
}

describe("OfficeCompanionBridge", () => {
  it("binds receiver-sensitive fetch implementations to globalThis", async () => {
    const fetchMock = vi.fn(function (this: typeof globalThis) {
      if (this !== globalThis) throw new TypeError("Illegal invocation");
      return Promise.resolve(response(201, {
        ok: true,
        sessionId: SESSION_ID,
        sessionKey: SESSION_KEY,
        protocolVersion: "1.0",
        expiresInSeconds: 60,
      }));
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "",
      documentId: async () => "doc-1",
      paneId: PANE_ID,
      executeTool: vi.fn(),
      fetchImpl: fetchMock as typeof fetch,
    });

    await expect(bridge.connect()).resolves.toMatchObject({ sessionId: SESSION_ID });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("binds receiver-sensitive timer implementations to globalThis", async () => {
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, {
        ok: true,
        sessionId: SESSION_ID,
        sessionKey: SESSION_KEY,
        protocolVersion: "1.0",
        expiresInSeconds: 60,
      }))
      .mockResolvedValueOnce(response(204));
    let nextTimer = 0;
    const setIntervalMock = vi.fn(function (
      this: typeof globalThis,
      _handler: TimerHandler,
      _timeout?: number,
    ) {
      if (this !== globalThis) throw new TypeError("Illegal invocation");
      nextTimer += 1;
      return nextTimer;
    });
    const clearIntervalMock = vi.fn(function (this: typeof globalThis, _timer?: number) {
      if (this !== globalThis) throw new TypeError("Illegal invocation");
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "",
      documentId: async () => "doc-1",
      paneId: PANE_ID,
      executeTool: vi.fn(),
      fetchImpl: fetchMock as typeof fetch,
      setIntervalImpl: setIntervalMock as unknown as typeof setInterval,
      clearIntervalImpl: clearIntervalMock as unknown as typeof clearInterval,
    });

    await expect(bridge.start()).resolves.toMatchObject({ sessionId: SESSION_ID });
    expect(setIntervalMock).toHaveBeenCalledTimes(2);
    await bridge.stop();
    expect(clearIntervalMock).toHaveBeenCalledTimes(2);
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
      hostPrefix: "powerpoint", baseUrl: "",
      documentId: async () => "doc-1",
      paneId: PANE_ID,
      executeTool: vi.fn(),
      fetchImpl: fetchMock as typeof fetch,
    });

    const session = await bridge.connect();
    expect(session.sessionId).toBe(SESSION_ID);
    expect(bridge.sessionId).toBe(SESSION_ID);
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe("/powerpoint/session");
    expect(JSON.parse(String(init.body))).toEqual({
      protocolVersion: "1.0",
      documentId: "doc-1",
      paneId: PANE_ID,
    });
    expect(JSON.stringify(bridge.publicState)).not.toContain(SESSION_KEY);
  });

  it("preserves the global receiver required by WebView fetch", async () => {
    const fetchMock = vi.fn(function (this: typeof globalThis) {
      if (this !== globalThis) {
        throw new TypeError("Illegal invocation");
      }
      return Promise.resolve(response(201, {
        ok: true,
        sessionId: SESSION_ID,
        sessionKey: SESSION_KEY,
        protocolVersion: "1.0",
        expiresInSeconds: 60,
      }));
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "",
      documentId: async () => "doc-1",
      paneId: PANE_ID,
      executeTool: vi.fn(),
      fetchImpl: fetchMock as typeof fetch,
    });

    await expect(bridge.connect()).resolves.toMatchObject({ sessionId: SESSION_ID });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("deduplicates concurrent registration attempts", async () => {
    let resolveFetch!: (value: Response) => void;
    const fetchMock = vi.fn(() => new Promise<Response>((resolve) => {
      resolveFetch = resolve;
    }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });
    const first = bridge.connect();
    const second = bridge.connect();
    for (let index = 0; index < 10; index += 1) await Promise.resolve();
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
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });
    const connecting = bridge.connect();
    for (let index = 0; index < 10; index += 1) await Promise.resolve();
    const stopping = bridge.stop();
    resolveRegistration(response(201, {
      ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
      protocolVersion: "1.0", expiresInSeconds: 60,
    }));
    await connecting.catch(() => undefined);
    await stopping;
    for (let index = 0; index < 10; index += 1) await Promise.resolve();
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
        tool: "read_slide_content",
        arguments: { sheet: "Sheet1", address: "A1" },
      }))
      .mockResolvedValueOnce(response(200, { ok: true, ticketId: TICKET_ID }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "",
      documentId: async () => "doc-1",
      paneId: PANE_ID,
      executeTool,
      fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(executeTool).toHaveBeenCalledWith("read_slide_content", {
      sheet: "Sheet1", address: "A1",
    });
    const [resultUrl, resultInit] = fetchMock.mock.calls[2] as [string, RequestInit];
    expect(resultUrl).toBe("/powerpoint/result");
    expect(JSON.parse(String(resultInit.body))).toEqual({
      ticketId: TICKET_ID,
      result: { ok: true, data: [[1]] },
    });
  });

  it("retries a transiently failed cached result without re-executing the mutation", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: { written: true } }));
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, {
        ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
        protocolVersion: "1.0", expiresInSeconds: 60,
      }))
      .mockResolvedValueOnce(response(200, {
        ticketId: TICKET_ID, requestId: REQUEST_ID,
        tool: "add_text_box", arguments: { sheet: "Sheet1", address: "A1", values: [[1]] },
      }))
      .mockRejectedValueOnce(new TypeError("connection reset after write"))
      .mockResolvedValueOnce(response(200, { ok: true, ticketId: TICKET_ID }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as unknown as typeof fetch, randomImpl: () => 0,
    });
    await bridge.connect();
    await expect(bridge.pollOnce()).resolves.toBe(true);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[3][0]).toBe("/powerpoint/result");
  });

  it("retains the completed result cache when an uncertain result submission fails", async () => {
    vi.useFakeTimers();
    try {
      const replacementId = "ffffffff-ffff-4fff-8fff-ffffffffffff";
      const replacementKey = "r".repeat(64);
      let registrationCount = 0;
      let resultCount = 0;
      const executeTool = vi.fn(async () => JSON.stringify({
        ok: true,
        data: { written: true },
      }));
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url === "/powerpoint/session" && init?.method === "POST") {
          registrationCount += 1;
          if (registrationCount === 1) {
            return response(201, {
              ok: true,
              sessionId: SESSION_ID,
              sessionKey: SESSION_KEY,
              protocolVersion: "1.0",
              expiresInSeconds: 60,
            });
          }
          if (registrationCount > 1) return response(201, {
            ok: true,
            sessionId: replacementId,
            sessionKey: replacementKey,
            protocolVersion: "1.0",
            expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/pending") {
          return response(200, {
            ticketId: TICKET_ID,
            requestId: REQUEST_ID,
            tool: "add_text_box",
            arguments: { sheet: "Sheet1", address: "A1", values: [[1]] },
          });
        }
        if (url === "/powerpoint/result") {
          resultCount += 1;
          const headers = (init?.headers ?? {}) as Record<string, string>;
          if (headers["X-Hermes-Session-Id"] === SESSION_ID) {
            return response(401, { ok: false, category: "Auth", error: "Unauthorized" });
          }
          return response(200, { ok: true, ticketId: TICKET_ID });
        }
        if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
        throw new Error(`unexpected URL ${url}`);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "",
        documentId: async () => "doc-1",
        paneId: PANE_ID,
        executeTool,
        fetchImpl: fetchMock as unknown as typeof fetch,
        setIntervalImpl: vi.fn(() => 1) as unknown as typeof setInterval,
        clearIntervalImpl: vi.fn() as unknown as typeof clearInterval,
        randomImpl: () => 0,
      });

      await bridge.start();
      const poll = bridge.pollOnce().catch(() => undefined);
      await flushMicrotasks();
      expect(resultCount).toBe(1);
      await bridge.stop();
      await poll;
      expect(cachedResultsOf(bridge).has(REQUEST_ID)).toBe(true);

      const reconnect = bridge.connect();
      await flushMicrotasks();
      await vi.runAllTimersAsync();
      await expect(reconnect).resolves.toMatchObject({ sessionId: replacementId });
      await vi.advanceTimersByTimeAsync(500);
      await expect(bridge.pollOnce()).resolves.toBe(true);
      expect(executeTool).toHaveBeenCalledTimes(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it("drops a terminally rejected cached result and resumes polling without re-execution", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: { written: true } }));
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, {
        ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
        protocolVersion: "1.0", expiresInSeconds: 60,
      }))
      .mockResolvedValueOnce(response(200, {
        ticketId: TICKET_ID, requestId: REQUEST_ID,
        tool: "add_text_box", arguments: { sheet: "Sheet1", address: "A1", values: [[1]] },
      }))
      .mockResolvedValueOnce(response(400, {
        ok: false, category: "Validation", error: "Ticket is no longer valid",
      }))
      .mockResolvedValueOnce(response(204));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();

    expect(await bridge.pollOnce()).toBe(true);
    expect(await bridge.pollOnce()).toBe(false);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[3][0]).toBe("/powerpoint/pending");
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
        tool: "read_slide_content", arguments: {},
      }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();
    await expect(bridge.pollOnce()).rejects.toThrow("ticketId");
    expect(executeTool).not.toHaveBeenCalled();
  });

  it("rejects companion responses by UTF-8 byte length, not JavaScript characters", async () => {
    const oversizedError = "😀".repeat(600_000);
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, {
        ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
        protocolVersion: "1.0", expiresInSeconds: 60,
      }))
      .mockResolvedValueOnce(response(500, { error: oversizedError }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });

    await bridge.connect();
    await expect(bridge.heartbeatOnce()).rejects.toThrow("response is too large");
  });

  it("sanitizes an oversized serialized result, caches it, and never re-executes", async () => {
    vi.useFakeTimers();
    try {
      const executeTool = vi.fn(async () => JSON.stringify({ data: "😀".repeat(600_000) }));
      let resultSubmissions = 0;
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url.endsWith("/powerpoint/session") && init?.method === "POST") {
          return response(201, {
            ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url.endsWith("/powerpoint/pending")) {
          return response(200, {
            ticketId: TICKET_ID, requestId: REQUEST_ID,
            tool: "read_slide_content", arguments: {},
          });
        }
        if (url.endsWith("/powerpoint/result")) {
          resultSubmissions += 1;
          if (resultSubmissions === 1) throw new TypeError("connection reset after write");
          return response(200, { ok: true, ticketId: TICKET_ID });
        }
        if (url.endsWith("/powerpoint/session") && init?.method === "DELETE") {
          return response(204);
        }
        throw new Error(`unexpected URL ${url}`);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool, fetchImpl: fetchMock as unknown as typeof fetch, randomImpl: () => 0,
      });

      await bridge.connect();
      const poll = bridge.pollOnce().catch(() => undefined);
      for (let tick = 0; tick < 30; tick += 1) await Promise.resolve();
      expect(executeTool).toHaveBeenCalledTimes(1);
      expect(resultSubmissions).toBe(1);
      const internals = bridge as unknown as {
        completedResults: Map<string, { result: unknown; serializedBody: string }>;
      };
      const cached = internals.completedResults.get(REQUEST_ID);
      expect(cached).toBeDefined();
      expect(cached?.result).toEqual({
        ok: false, category: "Result", error: "Companion result is too large",
      });
      expect(cached?.serializedBody).not.toContain("😀");
      expect(executeTool).toHaveBeenCalledTimes(1);

      await bridge.stop();
      await poll;
    } finally {
      vi.useRealTimers();
    }
  });

  it("never evicts protected results at the count or byte cap", () => {
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: vi.fn() as unknown as typeof fetch,
    });
    const internals = bridge as unknown as {
      cacheCompletedResult: (ticketId: string, result: unknown) => void;
      completedResults: Map<string, unknown>;
      completedResultBytes: number;
    };

    for (let index = 0; index < 8; index += 1) {
      internals.cacheCompletedResult(`ticket-${index}`, { ok: true, index });
    }
    expect(internals.completedResults.size).toBe(8);
    expect(() => internals.cacheCompletedResult("ticket-8", { ok: true })).toThrow("capacity");
    expect(internals.completedResults.has("ticket-0")).toBe(true);
    expect(internals.completedResults.has("ticket-8")).toBe(false);

    const byteBounded = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: vi.fn() as unknown as typeof fetch,
    }) as unknown as {
      cacheCompletedResult: (ticketId: string, result: unknown) => void;
      completedResults: Map<string, unknown>;
      completedResultBytes: number;
    };
    byteBounded.cacheCompletedResult("ticket-a", { data: "x".repeat(1_500_000) });
    byteBounded.cacheCompletedResult("ticket-b", { data: "x".repeat(1_500_000) });
    expect(() => byteBounded.cacheCompletedResult(
      "ticket-c", { data: "x".repeat(1_500_000) },
    )).toThrow("capacity");
    expect(byteBounded.completedResults.has("ticket-a")).toBe(true);
    expect(byteBounded.completedResults.has("ticket-c")).toBe(false);
    expect(byteBounded.completedResultBytes).toBeLessThanOrEqual(4 * 1024 * 1024);
  });

  it("fails closed before Office execution when protected count capacity is full", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true }));
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url === "/powerpoint/session" && init?.method === "POST") {
        return response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url === "/powerpoint/pending") {
        return response(200, {
          ticketId: TICKET_ID, requestId: REQUEST_ID,
          tool: "add_text_box", arguments: { text: "blocked-by-count" },
        });
      }
      if (url === "/powerpoint/result") {
        return response(400, { ok: false, category: "Validation", error: "not expected" });
      }
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as unknown as typeof fetch,
    });
    const internals = bridge as unknown as {
      cacheCompletedResult: (ticketId: string, result: unknown) => void;
      completedResults: Map<string, unknown>;
      reservedResultSlots: number;
      reservedResultBytes: number;
    };
    for (let index = 0; index < 8; index += 1) {
      internals.cacheCompletedResult(`protected-${index}`, { ok: true, index });
    }

    await bridge.connect();
    await expect(bridge.pollOnce()).rejects.toThrow("capacity");
    expect(executeTool).not.toHaveBeenCalled();
    expect(internals.completedResults.size).toBe(8);
    expect(internals.reservedResultSlots).toBe(0);
    expect(internals.reservedResultBytes).toBe(0);
  });

  it("fails closed before Office execution when protected byte capacity is full", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true }));
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url === "/powerpoint/session" && init?.method === "POST") {
        return response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url === "/powerpoint/pending") {
        return response(200, {
          ticketId: TICKET_ID, requestId: REQUEST_ID,
          tool: "add_text_box", arguments: { text: "blocked-by-bytes" },
        });
      }
      if (url === "/powerpoint/result") {
        return response(400, { ok: false, category: "Validation", error: "not expected" });
      }
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as unknown as typeof fetch,
    });
    const internals = bridge as unknown as {
      cacheCompletedResult: (ticketId: string, result: unknown) => void;
      completedResults: Map<string, unknown>;
      completedResultBytes: number;
      reservedResultSlots: number;
      reservedResultBytes: number;
    };
    internals.cacheCompletedResult("protected-a", { data: "x".repeat(1_500_000) });
    internals.cacheCompletedResult("protected-b", { data: "x".repeat(1_500_000) });

    await bridge.connect();
    await expect(bridge.pollOnce()).rejects.toThrow("capacity");
    expect(executeTool).not.toHaveBeenCalled();
    expect(internals.completedResults.size).toBe(2);
    expect(internals.completedResultBytes).toBeGreaterThan(3_000_000);
    expect(internals.reservedResultSlots).toBe(0);
    expect(internals.reservedResultBytes).toBe(0);
  });

  it("keeps a protected same-request result through pressure and reattaches it", async () => {
    vi.useFakeTimers();
    try {
      let registrationCount = 0;
      const executeTool = vi.fn(async () => JSON.stringify({ ok: true, value: "once" }));
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        const headers = (init?.headers ?? {}) as Record<string, string>;
        if (url === "/powerpoint/session" && init?.method === "POST") {
          registrationCount += 1;
          return response(201, {
            ok: true,
            sessionId: registrationCount === 1 ? SESSION_ID : SESSION_ID_B,
            sessionKey: registrationCount === 1 ? SESSION_KEY : SESSION_KEY_B,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/pending") {
          return response(200, {
            ticketId: headers["X-Hermes-Session-Id"] === SESSION_ID ? TICKET_ID : TICKET_ID_B,
            requestId: REQUEST_ID,
            tool: "add_text_box", arguments: { text: "protected-replay" },
          });
        }
        if (url === "/powerpoint/result") {
          return headers["X-Hermes-Session-Id"] === SESSION_ID
            ? response(401, { ok: false, category: "Auth", error: "uncertain" })
            : response(200, { ok: true, ticketId: TICKET_ID_B });
        }
        if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
        throw new Error(`unexpected URL ${url}`);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool, fetchImpl: fetchMock as unknown as typeof fetch, randomImpl: () => 0,
      });
      const internals = bridge as unknown as {
        cacheCompletedResult: (ticketId: string, result: unknown) => void;
        completedResults: Map<string, unknown>;
      };

      await bridge.connect();
      const oldPoll = bridge.pollOnce();
      await flushMicrotasks();
      expect(executeTool).toHaveBeenCalledTimes(1);
      await bridge.stop();
      await oldPoll;
      expect(internals.completedResults.has(REQUEST_ID)).toBe(true);

      for (let index = 0; index < 7; index += 1) {
        internals.cacheCompletedResult(`pressure-${index}`, { ok: true, index });
      }
      expect(() => internals.cacheCompletedResult("pressure-overflow", { ok: true })).toThrow("capacity");
      expect(internals.completedResults.has(REQUEST_ID)).toBe(true);

      await expect(bridge.connect()).resolves.toMatchObject({ sessionId: SESSION_ID_B });
      const newPoll = bridge.pollOnce();
      await vi.advanceTimersByTimeAsync(500);
      await expect(newPoll).resolves.toBe(true);
      expect(executeTool).toHaveBeenCalledTimes(1);
      expect(internals.completedResults.has(REQUEST_ID)).toBe(false);
    } finally {
      vi.useRealTimers();
    }
  });

  it("releases a reservation when a queued serializer operation is cancelled", async () => {
    let releaseFirst!: () => void;
    const firstStarted = new Promise<void>((resolve) => { releaseFirst = resolve; });
    const executeFirst = vi.fn(async () => {
      await firstStarted;
      return JSON.stringify({ ok: true, first: true });
    });
    const executeSecond = vi.fn(async () => JSON.stringify({ ok: true, second: true }));
    const makeBridge = (
      sessionId: string,
      sessionKey: string,
      paneId: string,
      ticketId: string,
      requestId: string,
      executeTool: (name: string, args: Record<string, unknown>) => Promise<string>,
    ) => {
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url === "/powerpoint/session" && init?.method === "POST") {
          return response(201, {
            ok: true, sessionId, sessionKey,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/pending") {
          return response(200, { ticketId, requestId, tool: "add_text_box", arguments: { text: "queued" } });
        }
        if (url === "/powerpoint/result") return response(200, { ok: true, ticketId });
        if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
        throw new Error(`unexpected URL ${url}`);
      });
      return new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "same-deck", paneId,
        executeTool, fetchImpl: fetchMock as unknown as typeof fetch,
      });
    };
    const bridgeA = makeBridge(SESSION_ID, SESSION_KEY, PANE_ID, TICKET_ID, REQUEST_ID, executeFirst);
    const bridgeB = makeBridge(SESSION_ID_B, SESSION_KEY_B, PANE_ID_B, TICKET_ID_B, REQUEST_ID_B, executeSecond);
    const internalsB = bridgeB as unknown as {
      reservedResultSlots: number;
      reservedResultBytes: number;
    };

    let firstPoll: Promise<boolean> | undefined;
    let secondPoll: Promise<boolean> | undefined;
    let firstReleased = false;
    try {
      await Promise.all([bridgeA.connect(), bridgeB.connect()]);
      firstPoll = bridgeA.pollOnce();
      for (let index = 0; index < 50 && executeFirst.mock.calls.length === 0; index += 1) {
        await Promise.resolve();
      }
      secondPoll = bridgeB.pollOnce();
      for (let index = 0; index < 50 && internalsB.reservedResultSlots === 0; index += 1) {
        await Promise.resolve();
      }
      expect(executeFirst).toHaveBeenCalledTimes(1);
      expect(executeSecond).not.toHaveBeenCalled();
      expect(internalsB.reservedResultSlots).toBe(1);

      await bridgeB.stop();
      await expect(secondPoll).resolves.toBe(false);
      expect(internalsB.reservedResultSlots).toBe(0);
      expect(internalsB.reservedResultBytes).toBe(0);

      releaseFirst();
      firstReleased = true;
      await expect(firstPoll).resolves.toBe(true);
      await bridgeA.stop();
    } finally {
      if (!firstReleased) releaseFirst();
      if (firstPoll) await firstPoll.catch(() => undefined);
      if (secondPoll) await secondPoll.catch(() => undefined);
      await bridgeB.stop();
      await bridgeA.stop();
    }
  });

  it("frees a protected capacity slot only after exact acknowledgement", async () => {
    let pendingCount = 0;
    const executeTool = vi.fn(async (_name: string, args: Record<string, unknown>) =>
      JSON.stringify({ ok: true, value: args.text }),
    );
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url === "/powerpoint/session" && init?.method === "POST") {
        return response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url === "/powerpoint/pending") {
        pendingCount += 1;
        return response(200, {
          ticketId: pendingCount === 1 ? TICKET_ID : TICKET_ID_B,
          requestId: pendingCount === 1 ? REQUEST_ID : REQUEST_ID_B,
          tool: "add_text_box", arguments: { text: pendingCount === 1 ? "first" : "second" },
        });
      }
      if (url === "/powerpoint/result") {
        const body = JSON.parse(String(init?.body)) as { ticketId: string };
        return response(200, { ok: true, ticketId: body.ticketId });
      }
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as unknown as typeof fetch,
    });
    const internals = bridge as unknown as {
      cacheCompletedResult: (ticketId: string, result: unknown) => void;
      completedResults: Map<string, unknown>;
    };
    for (let index = 0; index < 7; index += 1) {
      internals.cacheCompletedResult(`protected-${index}`, { ok: true, index });
    }

    await bridge.connect();
    await expect(bridge.pollOnce()).resolves.toBe(true);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(internals.completedResults.has(REQUEST_ID)).toBe(false);
    await expect(bridge.pollOnce()).resolves.toBe(true);
    expect(executeTool).toHaveBeenCalledTimes(2);
    expect(internals.completedResults.has(REQUEST_ID_B)).toBe(false);
  });

  it("keeps repeated capacity failures bounded without tombstone growth", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true }));
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url === "/powerpoint/session" && init?.method === "POST") {
        return response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url === "/powerpoint/pending") {
        return response(200, {
          ticketId: TICKET_ID, requestId: REQUEST_ID,
          tool: "add_text_box", arguments: { text: "bounded" },
        });
      }
      if (url === "/powerpoint/result") {
        return response(400, { ok: false, category: "Validation", error: "not expected" });
      }
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as unknown as typeof fetch,
    });
    const internals = bridge as unknown as {
      cacheCompletedResult: (ticketId: string, result: unknown) => void;
      completedResults: Map<string, unknown>;
      reservedResultSlots: number;
      reservedResultBytes: number;
    };
    for (let index = 0; index < 8; index += 1) {
      internals.cacheCompletedResult(`protected-${index}`, { ok: true, index });
    }

    await bridge.connect();
    for (let attempt = 0; attempt < 16; attempt += 1) {
      await expect(bridge.pollOnce()).rejects.toThrow("capacity");
    }
    expect(executeTool).not.toHaveBeenCalled();
    expect(internals.completedResults.size).toBe(8);
    expect(internals.reservedResultSlots).toBe(0);
    expect(internals.reservedResultBytes).toBe(0);
  });


  it("unregisters only its authenticated session on stop", async () => {
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, {
        ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
        protocolVersion: "1.0", expiresInSeconds: 60,
      }))
      .mockResolvedValueOnce(response(204));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();
    await bridge.stop();
    const [url, init] = fetchMock.mock.calls[1] as [string, RequestInit];
    expect(url).toBe("/powerpoint/session");
    expect(init.method).toBe("DELETE");
    expect((init.headers as Record<string, string>)["X-Hermes-Session-Id"]).toBe(SESSION_ID);
    expect(bridge.sessionId).toBeNull();
  });


  it("does not execute or publish a stale pending response after stop invalidates its lease", async () => {
    let resolvePending!: (value: Response) => void;
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true }));
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url === "/powerpoint/session" && init?.method === "POST") {
        return response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url === "/powerpoint/pending") {
        return new Promise<Response>((resolve) => { resolvePending = resolve; });
      }
      if (url === "/powerpoint/session" && init?.method === "DELETE") {
        return response(204);
      }
      if (url === "/powerpoint/result") {
        return response(200, { ok: true, ticketId: TICKET_ID });
      }
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as typeof fetch,
    });

    await bridge.connect();
    const pendingPoll = bridge.pollOnce();
    await Promise.resolve();
    await bridge.stop();
    resolvePending(response(200, {
      ticketId: TICKET_ID, requestId: REQUEST_ID,
      tool: "read_slide_content", arguments: { sheet: "Sheet1", address: "A1" },
    }));
    await pendingPoll;

    expect(executeTool).not.toHaveBeenCalled();
    expect(fetchMock.mock.calls.some(([url]) => url === "/powerpoint/result")).toBe(false);
  });


  it("suppresses a stale heartbeat failure after stop invalidates its lease", async () => {
    let resolveHeartbeat!: (value: Response) => void;
    let heartbeatHandler!: () => void;
    const onState = vi.fn();
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url === "/powerpoint/session" && init?.method === "POST") {
        return response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url === "/powerpoint/heartbeat") {
        return new Promise<Response>((resolve) => { resolveHeartbeat = resolve; });
      }
      if (url === "/powerpoint/session" && init?.method === "DELETE") {
        return response(204);
      }
      throw new Error(`unexpected URL ${url}`);
    });
    let nextTimer = 0;
    const setIntervalImpl = vi.fn((handler: TimerHandler, timeout?: number) => {
      if (timeout === 15_000) heartbeatHandler = handler as () => void;
      nextTimer += 1;
      return nextTimer;
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
      onState, setIntervalImpl: setIntervalImpl as unknown as typeof setInterval,
      clearIntervalImpl: vi.fn() as unknown as typeof clearInterval,
    });

    await bridge.start();
    heartbeatHandler();
    await Promise.resolve();
    await bridge.stop();
    const stateCountAfterStop = onState.mock.calls.length;
    resolveHeartbeat(response(401, { error: "expired" }));
    for (let index = 0; index < 50; index += 1) await Promise.resolve();

    expect(onState.mock.calls.length).toBe(stateCountAfterStop);
    expect(fetchMock.mock.calls.filter(([url, init]) =>
      url === "/powerpoint/session" && init?.method === "POST",
    )).toHaveLength(1);
  });


  it("returns promptly on stop and deletes a late valid registration exactly once", async () => {
    const lateSessionId = "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee";
    const lateSessionKey = "l".repeat(64);
    let resolveRegistration!: (value: Response) => void;
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url === "/powerpoint/session" && init?.method === "POST") {
        return new Promise<Response>((resolve) => { resolveRegistration = resolve; });
      }
      if (url === "/powerpoint/session" && init?.method === "DELETE") {
        return response(204);
      }
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });

    const connecting = bridge.connect();
    for (let index = 0; index < 10; index += 1) await Promise.resolve();
    const stopping = bridge.stop();
    let stopSettled = false;
    void stopping.then(() => { stopSettled = true; });
    await Promise.resolve();
    expect(stopSettled).toBe(true);

    resolveRegistration(response(201, {
      ok: true, sessionId: lateSessionId, sessionKey: lateSessionKey,
      protocolVersion: "1.0", expiresInSeconds: 60,
    }));
    await connecting.catch(() => undefined);
    for (let index = 0; index < 10; index += 1) await Promise.resolve();

    const deletes = fetchMock.mock.calls.filter(([url, init]) =>
      url === "/powerpoint/session" && init?.method === "DELETE",
    );
    expect(deletes).toHaveLength(1);
    const headers = deletes[0][1]?.headers as Record<string, string>;
    expect(headers["X-Hermes-Session-Id"]).toBe(lateSessionId);
    expect(headers["X-Hermes-Session-Key"]).toBe(lateSessionKey);
    expect(bridge.sessionId).toBeNull();
  });


  it("coalesces one five-attempt registration campaign with bounded retry delays", async () => {
    vi.useFakeTimers();
    const randomSpy = vi.spyOn(Math, "random").mockReturnValue(0);
    try {
      const fetchMock = vi.fn(async (url: string) => {
        if (url === "/powerpoint/session") return response(503);
        throw new Error(`unexpected URL ${url}`);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
      });

      const first = bridge.connect();
      const second = bridge.connect();
      const outcome = first.catch((error: unknown) => error);
      for (let index = 0; index < 10; index += 1) await Promise.resolve();
      expect(first).toBe(second);
      expect(fetchMock).toHaveBeenCalledTimes(1);

      await vi.advanceTimersByTimeAsync(499);
      expect(fetchMock).toHaveBeenCalledTimes(1);
      await vi.advanceTimersByTimeAsync(1);
      expect(fetchMock).toHaveBeenCalledTimes(2);
      await vi.advanceTimersByTimeAsync(999);
      expect(fetchMock).toHaveBeenCalledTimes(2);
      await vi.advanceTimersByTimeAsync(1);
      expect(fetchMock).toHaveBeenCalledTimes(3);
      await vi.advanceTimersByTimeAsync(1999);
      expect(fetchMock).toHaveBeenCalledTimes(3);
      await vi.advanceTimersByTimeAsync(1);
      expect(fetchMock).toHaveBeenCalledTimes(4);
      await vi.advanceTimersByTimeAsync(3999);
      expect(fetchMock).toHaveBeenCalledTimes(4);
      await vi.advanceTimersByTimeAsync(1);
      expect(fetchMock).toHaveBeenCalledTimes(5);
      await expect(outcome).resolves.toBeInstanceOf(Error);
    } finally {
      randomSpy.mockRestore();
      vi.useRealTimers();
    }
  });


  it("bounds each registration attempt to ten seconds including a hanging fetch", async () => {
    vi.useFakeTimers();
    const randomSpy = vi.spyOn(Math, "random").mockReturnValue(0);
    try {
      const fetchMock = vi.fn(() => new Promise<Response>(() => {}));
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
      });
      void bridge.connect().catch(() => undefined);
      for (let index = 0; index < 10; index += 1) await Promise.resolve();
      expect(fetchMock).toHaveBeenCalledTimes(1);

      await vi.advanceTimersByTimeAsync(9_999);
      expect(fetchMock).toHaveBeenCalledTimes(1);
      await vi.advanceTimersByTimeAsync(1);
      expect(fetchMock).toHaveBeenCalledTimes(1);
      await vi.advanceTimersByTimeAsync(500);
      expect(fetchMock).toHaveBeenCalledTimes(2);
    } finally {
      randomSpy.mockRestore();
      vi.useRealTimers();
    }
  });


  it("coalesces concurrent pending polls into one poll flight", async () => {
    let resolvePending!: (value: Response) => void;
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url === "/powerpoint/session" && init?.method === "POST") {
        return response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url === "/powerpoint/pending") {
        return new Promise<Response>((resolve) => { resolvePending = resolve; });
      }
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });

    await bridge.connect();
    const first = bridge.pollOnce();
    const second = bridge.pollOnce();
    await Promise.resolve();
    expect(first).toBe(second);
    resolvePending(response(204));
    await expect(first).resolves.toBe(false);
    expect(fetchMock.mock.calls.filter(([url]) => url === "/powerpoint/pending")).toHaveLength(1);
  });


  it("coalesces concurrent heartbeat calls into one heartbeat flight", async () => {
    let resolveHeartbeat!: (value: Response) => void;
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url === "/powerpoint/session" && init?.method === "POST") {
        return response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url === "/powerpoint/heartbeat") {
        return new Promise<Response>((resolve) => { resolveHeartbeat = resolve; });
      }
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });

    await bridge.connect();
    const first = bridge.heartbeatOnce();
    const second = bridge.heartbeatOnce();
    await Promise.resolve();
    expect(first).toBe(second);
    resolveHeartbeat(response(200, { ok: true }));
    await expect(first).resolves.toBeUndefined();
    expect(fetchMock.mock.calls.filter(([url]) => url === "/powerpoint/heartbeat")).toHaveLength(1);
  });


  it("retries every documented transient registration status then remains exhausted", async () => {
    vi.useFakeTimers();
    const randomSpy = vi.spyOn(Math, "random").mockReturnValue(0);
    try {
      const statuses = [401, 404, 408, 425, 429];
      let index = 0;
      const fetchMock = vi.fn(async (url: string) => {
        if (url !== "/powerpoint/session") throw new Error(`unexpected URL ${url}`);
        return response(statuses[index++]);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
      });
      const outcome = bridge.connect().catch((error: unknown) => error);
      for (let tick = 0; tick < 10; tick += 1) await Promise.resolve();
      expect(fetchMock).toHaveBeenCalledTimes(1);
      await vi.runAllTimersAsync();
      await expect(outcome).resolves.toBeInstanceOf(Error);
      expect(fetchMock).toHaveBeenCalledTimes(5);
      await vi.advanceTimersByTimeAsync(100_000);
      expect(fetchMock).toHaveBeenCalledTimes(5);
    } finally {
      randomSpy.mockRestore();
      vi.useRealTimers();
    }
  });


  it("reconnects once after a transient heartbeat failure without deleting the old lease", async () => {
    vi.useFakeTimers();
    const randomSpy = vi.spyOn(Math, "random").mockReturnValue(0);
    try {
      const replacementId = "ffffffff-ffff-4fff-8fff-ffffffffffff";
      const replacementKey = "r".repeat(64);
      let heartbeatHandler!: () => void;
      let registrationCount = 0;
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url === "/powerpoint/session" && init?.method === "POST") {
          registrationCount += 1;
          return response(201, {
            ok: true,
            sessionId: registrationCount === 1 ? SESSION_ID : replacementId,
            sessionKey: registrationCount === 1 ? SESSION_KEY : replacementKey,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/heartbeat") return response(401, { error: "expired" });
        if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
        throw new Error(`unexpected URL ${url}`);
      });
      let nextTimer = 0;
      const setIntervalImpl = vi.fn((handler: TimerHandler, timeout?: number) => {
        if (timeout === 15_000) heartbeatHandler = handler as () => void;
        nextTimer += 1;
        return nextTimer;
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
        setIntervalImpl: setIntervalImpl as unknown as typeof setInterval,
        clearIntervalImpl: vi.fn() as unknown as typeof clearInterval,
      });

      await bridge.start();
      heartbeatHandler();
      for (let tick = 0; tick < 50; tick += 1) await Promise.resolve();
      expect(fetchMock.mock.calls.filter(([url, init]) =>
        url === "/powerpoint/session" && init?.method === "DELETE",
      )).toHaveLength(0);
      expect(registrationCount).toBe(1);

      await vi.advanceTimersByTimeAsync(0);
      for (let tick = 0; tick < 50; tick += 1) await Promise.resolve();
      expect(registrationCount).toBe(2);
      expect(bridge.sessionId).toBe(replacementId);
      expect(fetchMock.mock.calls.filter(([url, init]) =>
        url === "/powerpoint/session" && init?.method === "DELETE",
      )).toHaveLength(0);
    } finally {
      randomSpy.mockRestore();
      vi.useRealTimers();
    }
  });


  it("allows an explicit connection reset after an exhausted registration campaign", async () => {
    vi.useFakeTimers();
    const randomSpy = vi.spyOn(Math, "random").mockReturnValue(0);
    try {
      let attempt = 0;
      const fetchMock = vi.fn(async (url: string) => {
        if (url !== "/powerpoint/session") throw new Error(`unexpected URL ${url}`);
        attempt += 1;
        if (attempt <= 5) return response(503);
        return response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
      });

      const first = bridge.connect().catch((error: unknown) => error);
      await vi.runAllTimersAsync();
      await expect(first).resolves.toBeInstanceOf(Error);
      expect(fetchMock).toHaveBeenCalledTimes(5);

      const second = bridge.connect();
      await expect(second).resolves.toMatchObject({ sessionId: SESSION_ID });
      expect(fetchMock).toHaveBeenCalledTimes(6);
    } finally {
      randomSpy.mockRestore();
      vi.useRealTimers();
    }
  });


  it("does not retry a registration failure after stop invalidates its campaign", async () => {
    vi.useFakeTimers();
    try {
      let resolveRegistration!: (value: Response) => void;
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url === "/powerpoint/session" && init?.method === "POST") {
          return new Promise<Response>((resolve) => { resolveRegistration = resolve; });
        }
        throw new Error(`unexpected URL ${url}`);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
      });

      const connecting = bridge.connect();
      for (let tick = 0; tick < 10; tick += 1) await Promise.resolve();
      expect(fetchMock).toHaveBeenCalledTimes(1);
      const stopping = bridge.stop();
      resolveRegistration(response(503));
      await connecting.catch(() => undefined);
      await stopping;
      await vi.runAllTimersAsync();
      expect(fetchMock).toHaveBeenCalledTimes(1);
    } finally {
      vi.useRealTimers();
    }
  });


  it("cancels a pending registration retry timer on stop", async () => {
    vi.useFakeTimers();
    const randomSpy = vi.spyOn(Math, "random").mockReturnValue(0);
    try {
      let attempt = 0;
      const fetchMock = vi.fn(async (url: string) => {
        if (url !== "/powerpoint/session") throw new Error(`unexpected URL ${url}`);
        attempt += 1;
        return response(503);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
      });
      void bridge.connect().catch(() => undefined);
      for (let tick = 0; tick < 10; tick += 1) await Promise.resolve();
      expect(attempt).toBe(1);
      await bridge.stop();
      await vi.runAllTimersAsync();
      expect(attempt).toBe(1);
    } finally {
      randomSpy.mockRestore();
      vi.useRealTimers();
    }
  });


  it("runs one poll flight and one heartbeat flight concurrently without duplication", async () => {
    let resolvePending!: (value: Response) => void;
    let resolveHeartbeat!: (value: Response) => void;
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url === "/powerpoint/session" && init?.method === "POST") {
        return response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url === "/powerpoint/pending") {
        return new Promise<Response>((resolve) => { resolvePending = resolve; });
      }
      if (url === "/powerpoint/heartbeat") {
        return new Promise<Response>((resolve) => { resolveHeartbeat = resolve; });
      }
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });

    await bridge.connect();
    const poll = bridge.pollOnce();
    const heartbeat = bridge.heartbeatOnce();
    await Promise.resolve();
    expect(fetchMock.mock.calls.filter(([url]) => url === "/powerpoint/pending")).toHaveLength(1);
    expect(fetchMock.mock.calls.filter(([url]) => url === "/powerpoint/heartbeat")).toHaveLength(1);
    resolvePending(response(204));
    resolveHeartbeat(response(200, { ok: true }));
    await expect(poll).resolves.toBe(false);
    await expect(heartbeat).resolves.toBeUndefined();
  });


  it("keeps an old pending response fenced after a replacement session is installed", async () => {
    const replacementId = "ffffffff-ffff-4fff-8fff-ffffffffffff";
    const replacementKey = "r".repeat(64);
    let registrationCount = 0;
    let resolvePending!: (value: Response) => void;
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true }));
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url === "/powerpoint/session" && init?.method === "POST") {
        registrationCount += 1;
        return response(201, {
          ok: true,
          sessionId: registrationCount === 1 ? SESSION_ID : replacementId,
          sessionKey: registrationCount === 1 ? SESSION_KEY : replacementKey,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url === "/powerpoint/pending") {
        return new Promise<Response>((resolve) => { resolvePending = resolve; });
      }
      if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
      if (url === "/powerpoint/result") return response(200, { ok: true, ticketId: TICKET_ID });
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as typeof fetch,
    });

    await bridge.connect();
    const oldPoll = bridge.pollOnce();
    await Promise.resolve();
    await bridge.stop();
    await expect(oldPoll).resolves.toBe(false);
    await bridge.connect();
    resolvePending(response(200, {
      ticketId: TICKET_ID, requestId: REQUEST_ID,
      tool: "read_slide_content", arguments: { sheet: "Sheet1", address: "A1" },
    }));
    for (let tick = 0; tick < 20; tick += 1) await Promise.resolve();

    expect(bridge.sessionId).toBe(replacementId);
    expect(executeTool).not.toHaveBeenCalled();
    expect(fetchMock.mock.calls.filter(([url]) => url === "/powerpoint/result")).toHaveLength(0);
  });


  it("stays disconnected after reconnect exhaustion until an explicit start reset", async () => {
    vi.useFakeTimers();
    const randomSpy = vi.spyOn(Math, "random").mockReturnValue(0);
    try {
      const replacementId = "ffffffff-ffff-4fff-8fff-ffffffffffff";
      const replacementKey = "r".repeat(64);
      let registrationCount = 0;
      let heartbeatHandler!: () => void;
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url === "/powerpoint/session" && init?.method === "POST") {
          registrationCount += 1;
          if (registrationCount === 1) {
            return response(201, {
              ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
              protocolVersion: "1.0", expiresInSeconds: 60,
            });
          }
          if (registrationCount <= 6) return response(503);
          return response(201, {
            ok: true, sessionId: replacementId, sessionKey: replacementKey,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/heartbeat") return response(401, { error: "expired" });
        if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
        throw new Error(`unexpected URL ${url}`);
      });
      let nextTimer = 0;
      const setIntervalImpl = vi.fn((handler: TimerHandler, timeout?: number) => {
        if (timeout === 15_000) heartbeatHandler = handler as () => void;
        nextTimer += 1;
        return nextTimer;
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
        setIntervalImpl: setIntervalImpl as unknown as typeof setInterval,
        clearIntervalImpl: vi.fn() as unknown as typeof clearInterval,
      });

      await bridge.start();
      heartbeatHandler();
      for (let tick = 0; tick < 50; tick += 1) await Promise.resolve();
      await vi.runAllTimersAsync();
      for (let tick = 0; tick < 50; tick += 1) await Promise.resolve();
      expect(registrationCount).toBe(6);
      expect(bridge.sessionId).toBeNull();
      await vi.advanceTimersByTimeAsync(100_000);
      expect(registrationCount).toBe(6);

      const reset = bridge.start();
      await expect(reset).resolves.toMatchObject({ sessionId: replacementId });
      expect(registrationCount).toBe(7);
    } finally {
      randomSpy.mockRestore();
      vi.useRealTimers();
    }
  });


  it("schedules reconnect for a direct heartbeat failure while running", async () => {
    vi.useFakeTimers();
    try {
      const replacementId = "ffffffff-ffff-4fff-8fff-ffffffffffff";
      const replacementKey = "r".repeat(64);
      let registrationCount = 0;
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url === "/powerpoint/session" && init?.method === "POST") {
          registrationCount += 1;
          return response(201, {
            ok: true,
            sessionId: registrationCount === 1 ? SESSION_ID : replacementId,
            sessionKey: registrationCount === 1 ? SESSION_KEY : replacementKey,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/heartbeat") return response(401, { error: "expired" });
        throw new Error(`unexpected URL ${url}`);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
        setIntervalImpl: vi.fn(() => 1) as unknown as typeof setInterval,
        clearIntervalImpl: vi.fn() as unknown as typeof clearInterval,
      });

      await bridge.start();
      await expect(bridge.heartbeatOnce()).rejects.toThrow("Heartbeat failed");
      await vi.advanceTimersByTimeAsync(0);
      expect(registrationCount).toBe(2);
      expect(bridge.sessionId).toBe(replacementId);
    } finally {
      vi.useRealTimers();
    }
  });


  it("aborts a timed-out registration attempt before starting its retry", async () => {
    vi.useFakeTimers();
    const randomSpy = vi.spyOn(Math, "random").mockReturnValue(0);
    try {
      let attempt = 0;
      let firstSignal!: AbortSignal;
      const fetchMock = vi.fn(async (_url: string, init?: RequestInit) => {
        attempt += 1;
        if (attempt === 1) {
          firstSignal = init?.signal as AbortSignal;
          return new Promise<Response>(() => {});
        }
        return response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
      });
      const connecting = bridge.connect();
      for (let tick = 0; tick < 10; tick += 1) await Promise.resolve();
      expect(attempt).toBe(1);
      await vi.advanceTimersByTimeAsync(10_000);
      expect(firstSignal.aborted).toBe(true);
      await vi.advanceTimersByTimeAsync(500);
      await expect(connecting).resolves.toMatchObject({ sessionId: SESSION_ID });
      expect(attempt).toBe(2);
    } finally {
      randomSpy.mockRestore();
      vi.useRealTimers();
    }
  });


  it("starts a fresh registration campaign immediately after stop cancels the old one", async () => {
    const replacementId = "ffffffff-ffff-4fff-8fff-ffffffffffff";
    const replacementKey = "r".repeat(64);
    let registrationCount = 0;
    let resolveFirst!: (value: Response) => void;
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url !== "/powerpoint/session" || init?.method !== "POST") {
        throw new Error(`unexpected URL ${url}`);
      }
      registrationCount += 1;
      if (registrationCount === 1) {
        return new Promise<Response>((resolve) => { resolveFirst = resolve; });
      }
      return response(201, {
        ok: true, sessionId: replacementId, sessionKey: replacementKey,
        protocolVersion: "1.0", expiresInSeconds: 60,
      });
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });

    const first = bridge.connect();
    for (let tick = 0; tick < 10; tick += 1) await Promise.resolve();
    await bridge.stop();
    const second = bridge.connect();
    expect(second).not.toBe(first);
    await expect(second).resolves.toMatchObject({ sessionId: replacementId });
    resolveFirst(response(201, {
      ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
      protocolVersion: "1.0", expiresInSeconds: 60,
    }));
    await first.catch(() => undefined);
    expect(registrationCount).toBe(2);
    expect(bridge.sessionId).toBe(replacementId);
  });

  it("holds a document execution slot until a stale Office mutation settles", async () => {
    let releaseExecution!: () => void;
    const executionStarted = new Promise<void>((resolve) => { releaseExecution = resolve; });
    const executeA = vi.fn(async () => {
      await executionStarted;
      return JSON.stringify({ ok: true, document: "stale-a" });
    });
    const executeB = vi.fn(async () => JSON.stringify({ ok: true, document: "stale-b" }));
    const makeBridge = (
      sessionId: string,
      sessionKey: string,
      paneId: string,
      ticketId: string,
      requestId: string,
      executeTool: (name: string, args: Record<string, unknown>) => Promise<string>,
    ) => {
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url === "/powerpoint/session" && init?.method === "POST") {
          return response(201, {
            ok: true, sessionId, sessionKey,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
        if (url === "/powerpoint/pending") {
          return response(200, {
            ticketId, requestId, tool: "add_text_box", arguments: { text: "stale" },
          });
        }
        if (url === "/powerpoint/result") return response(200, { ok: true, ticketId });
        throw new Error(`unexpected URL ${url}`);
      });
      return new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "stale-deck", paneId,
        executeTool, fetchImpl: fetchMock as typeof fetch,
      });
    };
    const bridgeA = makeBridge(
      SESSION_ID, SESSION_KEY, PANE_ID, TICKET_ID, REQUEST_ID, executeA,
    );
    const bridgeB = makeBridge(
      SESSION_ID_B, SESSION_KEY_B, PANE_ID_B, TICKET_ID_B, REQUEST_ID_B, executeB,
    );
    await Promise.all([bridgeA.connect(), bridgeB.connect()]);

    const oldPoll = bridgeA.pollOnce();
    for (let tick = 0; tick < 50 && executeA.mock.calls.length === 0; tick += 1) {
      await Promise.resolve();
    }
    const newPoll = bridgeB.pollOnce();
    for (let tick = 0; tick < 50; tick += 1) await Promise.resolve();
    expect(executeA).toHaveBeenCalledTimes(1);
    expect(executeB).not.toHaveBeenCalled();

    let oldPollSettled = false;
    void oldPoll.then(() => { oldPollSettled = true; });
    await bridgeA.stop();
    for (let tick = 0; tick < 20; tick += 1) await Promise.resolve();
    expect(oldPollSettled).toBe(false);
    expect(executeB).not.toHaveBeenCalled();

    releaseExecution();
    await expect(oldPoll).resolves.toBe(false);
    for (let tick = 0; tick < 50; tick += 1) await Promise.resolve();
    expect(executeB).toHaveBeenCalledTimes(1);
    await expect(newPoll).resolves.toBe(true);
  });

  it("keeps a stale result response from mutating a replacement session", async () => {
    const replacementId = "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee";
    const replacementKey = "r".repeat(64);
    let registrationCount = 0;
    let resolveResult!: (value: Response) => void;
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true }));
    const onState = vi.fn();
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url === "/powerpoint/session" && init?.method === "POST") {
        registrationCount += 1;
        return response(201, {
          ok: true,
          sessionId: registrationCount === 1 ? SESSION_ID : replacementId,
          sessionKey: registrationCount === 1 ? SESSION_KEY : replacementKey,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url === "/powerpoint/pending") {
        return response(200, {
          ticketId: TICKET_ID, requestId: REQUEST_ID,
          tool: "add_text_box", arguments: { text: "stale-result" },
        });
      }
      if (url === "/powerpoint/result") {
        return new Promise<Response>((resolve) => { resolveResult = resolve; });
      }
      if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "deck-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as typeof fetch, onState,
    });

    await bridge.connect();
    const oldPoll = bridge.pollOnce();
    for (let tick = 0; tick < 20 && !resolveResult; tick += 1) await Promise.resolve();
    expect(resolveResult).toBeTypeOf("function");
    await bridge.stop();
    await bridge.connect();
    const stateCountAfterReplacement = onState.mock.calls.length;
    resolveResult(response(200, { ok: true, ticketId: TICKET_ID }));
    await expect(oldPoll).resolves.toBe(false);

    expect(bridge.sessionId).toBe(replacementId);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(onState.mock.calls.length).toBe(stateCountAfterReplacement);
  });

  it("serializes mutations for one presentation while allowing unrelated documents to proceed", async () => {
    const createBridge = (
      sessionId: string,
      sessionKey: string,
      paneId: string,
      documentId: string,
      ticketId: string,
      requestId: string,
      executeTool: (name: string, args: Record<string, unknown>) => Promise<string>,
    ) => {
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url === "/powerpoint/session" && init?.method === "POST") {
          return response(201, {
            ok: true, sessionId, sessionKey,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/pending") {
          return response(200, {
            ticketId, requestId, tool: "add_text_box", arguments: { text: documentId },
          });
        }
        if (url === "/powerpoint/result") return response(200, { ok: true, ticketId });
        throw new Error(`unexpected URL ${url}`);
      });
      return new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => documentId, paneId,
        executeTool, fetchImpl: fetchMock as typeof fetch,
      });
    };

    let releaseSameDocument!: () => void;
    const sameDocumentStarted = new Promise<void>((resolve) => { releaseSameDocument = resolve; });
    const sameDocumentA = vi.fn(async () => {
      await sameDocumentStarted;
      return JSON.stringify({ ok: true, document: "same-a" });
    });
    const sameDocumentB = vi.fn(async () => JSON.stringify({ ok: true, document: "same-b" }));
    const bridgeA = createBridge(
      SESSION_ID, SESSION_KEY, PANE_ID, "same-deck", TICKET_ID, REQUEST_ID, sameDocumentA,
    );
    const bridgeB = createBridge(
      SESSION_ID_B, SESSION_KEY_B, PANE_ID_B, "same-deck", TICKET_ID_B, REQUEST_ID_B, sameDocumentB,
    );
    await Promise.all([bridgeA.connect(), bridgeB.connect()]);

    const first = bridgeA.pollOnce();
    for (let tick = 0; tick < 10; tick += 1) await Promise.resolve();
    const second = bridgeB.pollOnce();
    for (let tick = 0; tick < 10; tick += 1) await Promise.resolve();
    expect(sameDocumentA).toHaveBeenCalledTimes(1);
    expect(sameDocumentB).not.toHaveBeenCalled();

    releaseSameDocument();
    await expect(first).resolves.toBe(true);
    for (let tick = 0; tick < 10; tick += 1) await Promise.resolve();
    expect(sameDocumentB).toHaveBeenCalledTimes(1);
    await expect(second).resolves.toBe(true);

    let releaseDocumentA!: () => void;
    let releaseDocumentB!: () => void;
    const independentAReady = new Promise<void>((resolve) => { releaseDocumentA = resolve; });
    const independentBReady = new Promise<void>((resolve) => { releaseDocumentB = resolve; });
    const independentA = vi.fn(async () => {
      await independentAReady;
      return JSON.stringify({ ok: true, document: "independent-a" });
    });
    const independentB = vi.fn(async () => {
      await independentBReady;
      return JSON.stringify({ ok: true, document: "independent-b" });
    });
    const bridgeC = createBridge(
      SESSION_ID, SESSION_KEY, PANE_ID, "deck-c", TICKET_ID, REQUEST_ID, independentA,
    );
    const bridgeD = createBridge(
      SESSION_ID_B, SESSION_KEY_B, PANE_ID_B, "deck-d", TICKET_ID_B, REQUEST_ID_B, independentB,
    );
    await Promise.all([bridgeC.connect(), bridgeD.connect()]);
    const independentPollA = bridgeC.pollOnce();
    const independentPollB = bridgeD.pollOnce();
    for (let tick = 0; tick < 50; tick += 1) await Promise.resolve();
    expect(independentA).toHaveBeenCalledTimes(1);
    expect(independentB).toHaveBeenCalledTimes(1);
    releaseDocumentA();
    releaseDocumentB();
    await expect(independentPollA).resolves.toBe(true);
    await expect(independentPollB).resolves.toBe(true);
  });

  it("bounds result submission to five timed attempts with exact jittered backoff", async () => {
    vi.useFakeTimers();
    try {
      const resultSignals: AbortSignal[] = [];
      let resultAttempts = 0;
      const randomValues = [0, 0.5, 0.999, 0];
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url === "/powerpoint/session" && init?.method === "POST") {
          return response(201, {
            ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/pending") {
          return response(200, {
            ticketId: TICKET_ID, requestId: REQUEST_ID,
            tool: "add_text_box", arguments: { text: "bounded" },
          });
        }
        if (url === "/powerpoint/result") {
          resultAttempts += 1;
          resultSignals.push(init?.signal as AbortSignal);
          return new Promise<Response>(() => {});
        }
        if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
        throw new Error(`unexpected URL ${url}`);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(async () => JSON.stringify({ ok: true })),
        fetchImpl: fetchMock as unknown as typeof fetch,
        randomImpl: () => randomValues.shift() ?? 0,
      });

      await bridge.connect();
      const poll = bridge.pollOnce().catch((error: unknown) => error);
      await flushMicrotasks();
      expect(resultAttempts).toBe(1);

      const delays = [500, 1_125, 2_249, 4_000];
      for (let attempt = 0; attempt < 5; attempt += 1) {
        await vi.advanceTimersByTimeAsync(10_000);
        await flushMicrotasks();
        expect(resultSignals[attempt]?.aborted).toBe(true);
        if (attempt < 4) {
          await vi.advanceTimersByTimeAsync(delays[attempt] - 1);
          await flushMicrotasks();
          expect(resultAttempts).toBe(attempt + 1);
          await vi.advanceTimersByTimeAsync(1);
          await flushMicrotasks();
          expect(resultAttempts).toBe(attempt + 2);
        }
      }

      await expect(poll).resolves.toBeInstanceOf(Error);
      expect(resultAttempts).toBe(5);
      expect(cachedResultsOf(bridge).get(REQUEST_ID)?.resultAttempts).toBe(5);
      await vi.advanceTimersByTimeAsync(100_000);
      expect(resultAttempts).toBe(5);
      await bridge.stop();
    } finally {
      vi.useRealTimers();
    }
  });

  it("does not let a late timed-out result response delete or rewrite the cache", async () => {
    vi.useFakeTimers();
    try {
      let resolveFirstResult!: (value: Response) => void;
      let resultAttempts = 0;
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url === "/powerpoint/session" && init?.method === "POST") {
          return response(201, {
            ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/pending") {
          return response(200, {
            ticketId: TICKET_ID, requestId: REQUEST_ID,
            tool: "add_text_box", arguments: { text: "late" },
          });
        }
        if (url === "/powerpoint/result") {
          resultAttempts += 1;
          if (resultAttempts === 1) {
            return new Promise<Response>((resolve) => { resolveFirstResult = resolve; });
          }
          return new Promise<Response>(() => {});
        }
        if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
        throw new Error(`unexpected URL ${url}`);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(async () => JSON.stringify({ ok: true, value: "kept" })),
        fetchImpl: fetchMock as unknown as typeof fetch, randomImpl: () => 0,
      });

      await bridge.connect();
      const poll = bridge.pollOnce().catch(() => undefined);
      await flushMicrotasks();
      await vi.advanceTimersByTimeAsync(10_000);
      await flushMicrotasks();
      await vi.advanceTimersByTimeAsync(500);
      await flushMicrotasks();
      expect(resultAttempts).toBe(2);
      expect(cachedResultsOf(bridge).get(REQUEST_ID)?.resultAttempts).toBe(2);

      resolveFirstResult(response(200, { ok: true, ticketId: TICKET_ID }));
      await flushMicrotasks();
      expect(cachedResultsOf(bridge).has(REQUEST_ID)).toBe(true);
      expect(cachedResultsOf(bridge).get(REQUEST_ID)?.resultAttempts).toBe(2);

      await bridge.stop();
      await poll;
    } finally {
      vi.useRealTimers();
    }
  });

  it.each([
    { name: "ok false", acknowledgement: { ok: false, ticketId: TICKET_ID } },
    { name: "mismatched ticket", acknowledgement: { ok: true, ticketId: TICKET_ID_B } },
  ])("does not acknowledge a contradictory 200 result response ($name)", async ({ acknowledgement }) => {
    vi.useFakeTimers();
    try {
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url === "/powerpoint/session" && init?.method === "POST") {
          return response(201, {
            ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/pending") {
          return response(200, {
            ticketId: TICKET_ID, requestId: REQUEST_ID,
            tool: "add_text_box", arguments: { text: "ack" },
          });
        }
        if (url === "/powerpoint/result") return response(200, acknowledgement);
        if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
        throw new Error(`unexpected URL ${url}`);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(async () => JSON.stringify({ ok: true })),
        fetchImpl: fetchMock as unknown as typeof fetch, randomImpl: () => 0,
      });

      await bridge.connect();
      const poll = bridge.pollOnce().catch(() => undefined);
      await flushMicrotasks();
      await vi.runAllTimersAsync();
      await poll;
      expect(cachedResultsOf(bridge).has(REQUEST_ID)).toBe(true);
      await bridge.stop();
    } finally {
      vi.useRealTimers();
    }
  });

  it.each([401, 403, 404])("retains cached result on uncertain session/auth status %s", async (status) => {
    vi.useFakeTimers();
    try {
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url === "/powerpoint/session" && init?.method === "POST") {
          return response(201, {
            ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/pending") {
          return response(200, {
            ticketId: TICKET_ID, requestId: REQUEST_ID,
            tool: "add_text_box", arguments: { text: "auth" },
          });
        }
        if (url === "/powerpoint/result") {
          return response(status, {
            ok: false,
            category: status === 404 ? "NotFound" : "Auth",
            error: status === 404 ? "No ticket found" : "Unauthorized",
          });
        }
        if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
        throw new Error(`unexpected URL ${url}`);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool: vi.fn(async () => JSON.stringify({ ok: true })),
        fetchImpl: fetchMock as unknown as typeof fetch, randomImpl: () => 0,
      });

      await bridge.connect();
      const poll = bridge.pollOnce().catch(() => undefined);
      await flushMicrotasks();
      await vi.runAllTimersAsync();
      await poll;
      expect(cachedResultsOf(bridge).get(REQUEST_ID)?.resultAttempts).toBe(5);
      expect(cachedResultsOf(bridge).has(REQUEST_ID)).toBe(true);
      await bridge.stop();
    } finally {
      vi.useRealTimers();
    }
  });

  it("removes only the matching cache entry for a protocol-defined terminal rejection", async () => {
    vi.useFakeTimers();
    try {
      let pendingCount = 0;
      let resultCount = 0;
      const executeTool = vi.fn(async () => JSON.stringify({ ok: true }));
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        if (url === "/powerpoint/session" && init?.method === "POST") {
          return response(201, {
            ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/pending") {
          pendingCount += 1;
          if (pendingCount === 1) {
            return response(200, {
              ticketId: TICKET_ID, requestId: REQUEST_ID,
              tool: "add_text_box", arguments: { text: "terminal" },
            });
          }
          return response(200, {
            ticketId: TICKET_ID_B, requestId: REQUEST_ID_B,
            tool: "add_text_box", arguments: { text: "retained" },
          });
        }
        if (url === "/powerpoint/result") {
          resultCount += 1;
          if (resultCount === 1) {
            return response(400, {
              ok: false, category: "Validation", error: "Ticket is no longer valid",
            });
          }
          return response(401, { ok: false, category: "Auth", error: "Unauthorized" });
        }
        if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
        throw new Error(`unexpected URL ${url}`);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool, fetchImpl: fetchMock as unknown as typeof fetch, randomImpl: () => 0,
      });

      await bridge.connect();
      await expect(bridge.pollOnce()).resolves.toBe(true);
      const secondPoll = bridge.pollOnce().catch(() => undefined);
      await flushMicrotasks();
      await vi.runAllTimersAsync();
      await secondPoll;
      expect(executeTool).toHaveBeenCalledTimes(2);
      expect(cachedResultsOf(bridge).has(REQUEST_ID)).toBe(false);
      expect(cachedResultsOf(bridge).has(REQUEST_ID_B)).toBe(true);
      await bridge.stop();
    } finally {
      vi.useRealTimers();
    }
  });

  it("reattaches an old-session cache to the current same-request ticket", async () => {
    vi.useFakeTimers();
    try {
      const replacementId = SESSION_ID_B;
      const replacementKey = SESSION_KEY_B;
      let registrationCount = 0;
      const executeTool = vi.fn(async () => JSON.stringify({ ok: true, value: "once" }));
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        const headers = (init?.headers ?? {}) as Record<string, string>;
        if (url === "/powerpoint/session" && init?.method === "POST") {
          registrationCount += 1;
          return response(201, {
            ok: true,
            sessionId: registrationCount === 1 ? SESSION_ID : replacementId,
            sessionKey: registrationCount === 1 ? SESSION_KEY : replacementKey,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/pending") {
          return response(200, {
            ticketId: headers["X-Hermes-Session-Id"] === SESSION_ID ? TICKET_ID : TICKET_ID_B,
            requestId: REQUEST_ID,
            tool: "add_text_box", arguments: { text: "reattach" },
          });
        }
        if (url === "/powerpoint/result") {
          if (headers["X-Hermes-Session-Id"] === SESSION_ID) {
            return response(401, { ok: false, category: "Auth", error: "Unauthorized" });
          }
          return response(200, { ok: true, ticketId: TICKET_ID_B });
        }
        if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
        throw new Error(`unexpected URL ${url}`);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool, fetchImpl: fetchMock as unknown as typeof fetch, randomImpl: () => 0,
      });

      await bridge.connect();
      const oldPoll = bridge.pollOnce().catch(() => undefined);
      await flushMicrotasks();
      expect(cachedResultsOf(bridge).get(REQUEST_ID)?.sessionId).toBe(SESSION_ID);
      await bridge.stop();
      await oldPoll;

      await expect(bridge.connect()).resolves.toMatchObject({ sessionId: replacementId });
      await vi.advanceTimersByTimeAsync(500);
      await expect(bridge.pollOnce()).resolves.toBe(true);
      expect(executeTool).toHaveBeenCalledTimes(1);
      const resultCalls = fetchMock.mock.calls.filter(([url]) => url === "/powerpoint/result");
      expect(resultCalls).toHaveLength(2);
      const [, resultInit] = resultCalls[1] as [string, RequestInit];
      const resultHeaders = resultInit.headers as Record<string, string>;
      expect(resultHeaders["X-Hermes-Session-Id"]).toBe(replacementId);
      expect(JSON.parse(String(resultInit.body)).ticketId).toBe(TICKET_ID_B);
      expect(cachedResultsOf(bridge).has(REQUEST_ID)).toBe(false);
      await bridge.stop();
    } finally {
      vi.useRealTimers();
    }
  });

  it("does not let an unrelated replacement request consume an old cache", async () => {
    vi.useFakeTimers();
    try {
      let registrationCount = 0;
      let pendingCount = 0;
      const executeTool = vi.fn(async (_name: string, args: Record<string, unknown>) =>
        JSON.stringify({ ok: true, value: args.text }),
      );
      const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
        const headers = (init?.headers ?? {}) as Record<string, string>;
        if (url === "/powerpoint/session" && init?.method === "POST") {
          registrationCount += 1;
          return response(201, {
            ok: true,
            sessionId: registrationCount === 1 ? SESSION_ID : SESSION_ID_B,
            sessionKey: registrationCount === 1 ? SESSION_KEY : SESSION_KEY_B,
            protocolVersion: "1.0", expiresInSeconds: 60,
          });
        }
        if (url === "/powerpoint/pending") {
          pendingCount += 1;
          return response(200, {
            ticketId: pendingCount === 1 ? TICKET_ID : TICKET_ID_B,
            requestId: pendingCount === 1 ? REQUEST_ID : REQUEST_ID_B,
            tool: "add_text_box", arguments: { text: pendingCount === 1 ? "old" : "new" },
          });
        }
        if (url === "/powerpoint/result") {
          if (headers["X-Hermes-Session-Id"] === SESSION_ID) {
            return response(401, { ok: false, category: "Auth", error: "Unauthorized" });
          }
          return response(200, { ok: true, ticketId: TICKET_ID_B });
        }
        if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
        throw new Error(`unexpected URL ${url}`);
      });
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool, fetchImpl: fetchMock as unknown as typeof fetch, randomImpl: () => 0,
      });

      await bridge.connect();
      const oldPoll = bridge.pollOnce().catch(() => undefined);
      await flushMicrotasks();
      await bridge.stop();
      await oldPoll;
      await bridge.connect();
      await expect(bridge.pollOnce()).resolves.toBe(true);

      expect(executeTool).toHaveBeenCalledTimes(2);
      expect(cachedResultsOf(bridge).has(REQUEST_ID)).toBe(true);
      expect(cachedResultsOf(bridge).has(REQUEST_ID_B)).toBe(false);
      await bridge.stop();
    } finally {
      vi.useRealTimers();
    }
  });

  it("caches a completed stale Office operation and reattaches it exactly once", async () => {
    let releaseExecution!: (value: string) => void;
    const executeTool = vi.fn(() => new Promise<string>((resolve) => { releaseExecution = resolve; }));
    let registrationCount = 0;
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url === "/powerpoint/session" && init?.method === "POST") {
        registrationCount += 1;
        return response(201, {
          ok: true,
          sessionId: registrationCount === 1 ? SESSION_ID : SESSION_ID_B,
          sessionKey: registrationCount === 1 ? SESSION_KEY : SESSION_KEY_B,
          protocolVersion: "1.0", expiresInSeconds: 60,
        });
      }
      if (url === "/powerpoint/pending") {
        const headers = (init?.headers ?? {}) as Record<string, string>;
        return response(200, {
          ticketId: headers["X-Hermes-Session-Id"] === SESSION_ID ? TICKET_ID : TICKET_ID_B,
          requestId: REQUEST_ID,
          tool: "add_text_box", arguments: { text: "stale-complete" },
        });
      }
      if (url === "/powerpoint/result") {
        const headers = (init?.headers ?? {}) as Record<string, string>;
        const ticketId = headers["X-Hermes-Session-Id"] === SESSION_ID ? TICKET_ID : TICKET_ID_B;
        return response(200, { ok: true, ticketId });
      }
      if (url === "/powerpoint/session" && init?.method === "DELETE") return response(204);
      throw new Error(`unexpected URL ${url}`);
    });
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "powerpoint", baseUrl: "", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as unknown as typeof fetch,
    });

    await bridge.connect();
    const oldPoll = bridge.pollOnce();
    await flushMicrotasks();
    expect(executeTool).toHaveBeenCalledTimes(1);
    await bridge.stop();
    await bridge.connect();
    const newPoll = bridge.pollOnce();
    await flushMicrotasks();
    expect(executeTool).toHaveBeenCalledTimes(1);
    releaseExecution(JSON.stringify({ ok: true, value: "settled-after-stop" }));
    await expect(oldPoll).resolves.toBe(false);
    await expect(newPoll).resolves.toBe(true);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(cachedResultsOf(bridge).has(REQUEST_ID)).toBe(false);

    const resultCalls = fetchMock.mock.calls.filter(([url]) => url === "/powerpoint/result");
    const [, resultInit] = resultCalls[0] as [string, RequestInit];
    expect(JSON.parse(String(resultInit.body)).ticketId).toBe(TICKET_ID_B);
    await bridge.stop();
  });
});
