import { describe, expect, it, vi } from "vitest";
import { OfficeCompanionBridge } from "../officeBridge";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";

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
      hostPrefix: "excel",
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
      hostPrefix: "excel",
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

  it("exposes the exact document identity used for authenticated session registration", async () => {
    const fetchMock = vi.fn(async () =>
      response(201, {
        ok: true,
        sessionId: SESSION_ID,
        sessionKey: SESSION_KEY,
        protocolVersion: "1.0",
        expiresInSeconds: 60,
      }),
    );
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "excel",
      documentId: async () => "doc-registered",
      paneId: PANE_ID,
      executeTool: vi.fn(),
      fetchImpl: fetchMock as typeof fetch,
    });

    const session = await bridge.connect();
    expect((session as typeof session & { documentId?: string }).documentId).toBe("doc-registered");
    expect((bridge as typeof bridge & { documentId?: string }).documentId).toBe("doc-registered");
    expect((bridge.publicState as typeof bridge.publicState & { documentId?: string | null }).documentId).toBe("doc-registered");
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
      hostPrefix: "excel",
      documentId: async () => "doc-1",
      paneId: PANE_ID,
      executeTool: vi.fn(),
      fetchImpl: fetchMock as typeof fetch,
    });

    const session = await bridge.connect();
    expect(session.sessionId).toBe(SESSION_ID);
    expect(bridge.sessionId).toBe(SESSION_ID);
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe("/excel/session");
    expect(JSON.parse(String(init.body))).toEqual({
      protocolVersion: "1.0",
      documentId: "doc-1",
      paneId: PANE_ID,
      workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
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
      hostPrefix: "excel",
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
      hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
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
      hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
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
        tool: "read_range",
        arguments: { sheet: "Sheet1", address: "A1" },
      }))
      .mockResolvedValueOnce(response(200, { ok: true, ticketId: TICKET_ID }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "excel",
      documentId: async () => "doc-1",
      paneId: PANE_ID,
      executeTool,
      fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();
    expect(await bridge.pollOnce()).toBe(true);
    expect(executeTool).toHaveBeenCalledWith("read_range", {
      sheet: "Sheet1", address: "A1",
    });
    const [resultUrl, resultInit] = fetchMock.mock.calls[2] as [string, RequestInit];
    expect(resultUrl).toBe("/excel/result");
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
        tool: "write_range", arguments: { sheet: "Sheet1", address: "A1", values: [[1]] },
      }))
      .mockRejectedValueOnce(new TypeError("connection reset after write"))
      .mockResolvedValueOnce(response(200, { ok: true, ticketId: TICKET_ID }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();
    await expect(bridge.pollOnce()).rejects.toThrow("connection reset");
    expect(await bridge.pollOnce()).toBe(true);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[3][0]).toBe("/excel/result");
  });

  it("drops a cached result after a terminal companion rejection", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: { written: true } }));
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, {
        ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
        protocolVersion: "1.0", expiresInSeconds: 60,
      }))
      .mockResolvedValueOnce(response(200, {
        ticketId: TICKET_ID, requestId: REQUEST_ID,
        tool: "write_range", arguments: { sheet: "Sheet1", address: "A1", values: [[1]] },
      }))
      .mockResolvedValueOnce(response(409, { error: "Ticket is already complete" }))
      .mockResolvedValueOnce(response(204));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();

    expect(await bridge.pollOnce()).toBe(true);
    expect(await bridge.pollOnce()).toBe(false);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[3][0]).toBe("/excel/pending");
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
        tool: "read_range", arguments: {},
      }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();
    await expect(bridge.pollOnce()).rejects.toThrow("ticketId");
    expect(executeTool).not.toHaveBeenCalled();
  });

  it.each(["read_range", "execute_plan"])(
    "keeps the %s primitive round trip byte-for-byte: two executeTool arguments, no executed field, no execute-start",
    async (tool) => {
      const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: [[1]] }));
      const fetchMock = vi.fn()
        .mockResolvedValueOnce(response(201, {
          ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
          protocolVersion: "1.0", expiresInSeconds: 60,
          workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
        }))
        .mockResolvedValueOnce(response(200, {
          ticketId: TICKET_ID, requestId: REQUEST_ID,
          tool, arguments: { sheet: "Sheet1", address: "A1" },
        }))
        .mockResolvedValueOnce(response(200, { ok: true, ticketId: TICKET_ID }));
      const bridge = new OfficeCompanionBridge({
        hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
        executeTool, fetchImpl: fetchMock as typeof fetch,
      });
      await bridge.connect();
      expect(await bridge.pollOnce()).toBe(true);
      expect(executeTool).toHaveBeenCalledTimes(1);
      expect(executeTool.mock.calls[0]).toEqual([tool, { sheet: "Sheet1", address: "A1" }]);
      expect(executeTool.mock.calls[0]).toHaveLength(2);
      expect(fetchMock.mock.calls.map(([url]) => url)).toEqual([
        "/excel/session", "/excel/pending", "/excel/result",
      ]);
      expect(String((fetchMock.mock.calls[2] as [string, RequestInit])[1].body)).toBe(
        JSON.stringify({ ticketId: TICKET_ID, result: { ok: true, data: [[1]] } }),
      );
    },
  );

  it("never hands a trusted plan envelope to a primitive tool", async () => {
    const executeTool = vi.fn();
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, {
        ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
        protocolVersion: "1.0", expiresInSeconds: 60,
        workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
      }))
      .mockResolvedValueOnce(response(200, {
        ticketId: TICKET_ID, requestId: REQUEST_ID,
        tool: "execute_plan", arguments: { version: 1 },
        trusted: { turnOperationId: REQUEST_ID, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT, phase: "execute", turnKey: "4b4b0000-0000-4000-8000-0000000000a3" },
      }))
      .mockResolvedValueOnce(response(200, { ok: true }));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();
    await expect(bridge.pollOnce()).resolves.toBe(true);
    expect(executeTool).not.toHaveBeenCalled();
    expect(fetchMock.mock.calls.map(([url]) => url)).toEqual(["/excel/session", "/excel/pending", "/excel/result"]);
    expect(fetchMock.mock.calls[2][1].body).toBe(
      JSON.stringify({ ticketId: TICKET_ID, result: { ok: false, category: "Execution", error: "Office tool execution failed" } }),
    );
  });

  it("K-S4a bounds streamed pending bytes before any execution or post", async () => {
    const executeTool = vi.fn();
    const pending = response(200, {
      ticketId: TICKET_ID, requestId: REQUEST_ID, tool: "read_range",
      arguments: { data: "😀".repeat(2_700_000) },
    });
    const cancel = vi.spyOn(pending.body!, "cancel");
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, { ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY, protocolVersion: "1.0", expiresInSeconds: 60 }))
      .mockResolvedValueOnce(pending);
    const bridge = new OfficeCompanionBridge({ hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID, executeTool, fetchImpl: fetchMock as typeof fetch });
    await bridge.connect();
    await expect(bridge.pollOnce()).rejects.toThrow("response is too large");
    expect(executeTool).not.toHaveBeenCalled();
    expect(fetchMock.mock.calls.map(([url]) => url)).toEqual(["/excel/session", "/excel/pending"]);
    expect(pending.body!.locked).toBe(false);
    cancel.mockRestore();
  });

  it("K-S4b rejects declared oversized pending before reading", async () => {
    const pending = response(200, { ticketId: TICKET_ID, requestId: REQUEST_ID, tool: "read_range", arguments: {} });
    pending.headers.set("Content-Length", String(10 * 1024 * 1024 + 64 * 1024 + 1));
    const getReader = vi.spyOn(pending.body!, "getReader");
    const text = vi.spyOn(pending, "text");
    const executeTool = vi.fn(async () => "{}");
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, { ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY, protocolVersion: "1.0", expiresInSeconds: 60 }))
      .mockResolvedValueOnce(pending).mockResolvedValue(response(200, { ok: true }));
    const bridge = new OfficeCompanionBridge({ hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID, executeTool, fetchImpl: fetchMock as typeof fetch });
    await bridge.connect();
    const outcome = bridge.pollOnce();
    await expect(outcome).rejects.toThrow("response is too large");
    expect(getReader).not.toHaveBeenCalled();
    expect(text).not.toHaveBeenCalled();
    expect(executeTool).not.toHaveBeenCalled();
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it("K-S4c rejects non-numeric Content-Length on every JSON caller", async () => {
    for (const route of ["session", "heartbeat", "pending", "result"]) {
      for (const length of ["not-a-number", "-1", "1.5", "9007199254740992"]) {
        const invalid = response(route === "session" ? 201 : route === "result" || route === "heartbeat" ? 500 : 200, {});
        invalid.headers.set("Content-Length", length);
        const text = vi.spyOn(invalid, "text");
        const getReader = vi.spyOn(invalid.body!, "getReader");
        const fetchMock = vi.fn(async (url: string) => {
          if (url.endsWith(`/${route}`)) return invalid;
          if (url.endsWith("/session")) return response(201, { ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY, protocolVersion: "1.0", expiresInSeconds: 60 });
          return response(200, { ticketId: TICKET_ID, requestId: REQUEST_ID, tool: "read_range", arguments: {} });
        });
        const bridge = new OfficeCompanionBridge({ hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID, executeTool: vi.fn(async () => "{}"), fetchImpl: fetchMock as typeof fetch });
        if (route === "session") await expect(bridge.connect()).rejects.toThrow("response is too large");
        else {
          await bridge.connect();
          await expect(route === "heartbeat" ? bridge.heartbeatOnce() : bridge.pollOnce()).rejects.toThrow("response is too large");
        }
        expect(text).not.toHaveBeenCalled();
        expect(getReader).not.toHaveBeenCalled();
      }
    }
  });

  it("K-S4d preserves 413 downgrade for oversized and non-JSON bodies", async () => {
    for (const body of [JSON.stringify({ error: "😀".repeat(2_700_000) }), "not JSON"]) {
      const fetchMock = vi.fn()
        .mockResolvedValueOnce(response(201, { ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY, protocolVersion: "1.0", expiresInSeconds: 60 }))
        .mockResolvedValueOnce(response(200, { ticketId: TICKET_ID, requestId: REQUEST_ID, tool: "read_range", arguments: {} }))
        .mockResolvedValueOnce(new Response(body, { status: 413 }))
        .mockResolvedValueOnce(response(200, { ok: true }));
      const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: [[1]] }));
      const bridge = new OfficeCompanionBridge({ hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID, executeTool, fetchImpl: fetchMock as typeof fetch });
      await bridge.connect();
      await expect(bridge.pollOnce()).rejects.toMatchObject({ name: "BridgeHttpError", status: 413, message: "Result submission failed" });
      const cache = (bridge as unknown as { completedResults: Map<string, { oversize?: boolean }> }).completedResults;
      expect(cache.get(TICKET_ID)?.oversize).toBe(true);
      await expect(bridge.pollOnce()).resolves.toBe(true);
      expect(executeTool).toHaveBeenCalledTimes(1);
      expect(JSON.parse(String(fetchMock.mock.calls[3][1].body))).toEqual({ ticketId: TICKET_ID, result: { ok: false, category: "Execution", error: "Office tool result was too large to submit." } });
      expect(cache.size).toBe(0);
    }
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
      hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });

    await bridge.connect();
    await expect(bridge.heartbeatOnce()).rejects.toThrow("response is too large");
  });

  it("unregisters only its authenticated session on stop", async () => {
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(response(201, {
        ok: true, sessionId: SESSION_ID, sessionKey: SESSION_KEY,
        protocolVersion: "1.0", expiresInSeconds: 60,
      }))
      .mockResolvedValueOnce(response(204));
    const bridge = new OfficeCompanionBridge({
      hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool: vi.fn(), fetchImpl: fetchMock as typeof fetch,
    });
    await bridge.connect();
    await bridge.stop();
    const [url, init] = fetchMock.mock.calls[1] as [string, RequestInit];
    expect(url).toBe("/excel/session");
    expect(init.method).toBe("DELETE");
    expect((init.headers as Record<string, string>)["X-Hermes-Session-Id"]).toBe(SESSION_ID);
    expect(bridge.sessionId).toBeNull();
  });
});
