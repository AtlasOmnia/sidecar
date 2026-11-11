import { afterEach, describe, expect, it, vi } from "vitest";
import * as officeBridge from "../officeBridge";
import { OfficeCompanionBridge, type BridgePlanExecutionControl } from "../officeBridge";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";

const SESSION_ID = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const SESSION_KEY = "s".repeat(64);
const SESSION_ID_2 = "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee";
const SESSION_KEY_2 = "t".repeat(64);
const PANE_ID = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
const TICKET_ID = "cccccccc-cccc-4ccc-8ccc-cccccccccccc";
const REQUEST_ID = "dddddddd-dddd-4ddd-8ddd-dddddddddddd";
const COMPANION_SECRET = "COMPANION_SECRET_TEXT_31C";
const THROWN_SECRET = "WORKBOOK_SECRET_IN_THROWN_ERROR_5B2";
const R4G2_TURN_KEY = "4b4b0000-0000-4000-8000-0000000000a3"; // R4g-2a rule M: turnKey of migrated v2 tickets
const TRUSTED = { turnOperationId: REQUEST_ID, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT, phase: "execute", turnKey: R4G2_TURN_KEY };
const EXECUTE_ARGS = { version: 1, outcome: "execute", summary: "s", operations: [], verification: [] };

function response(status: number, body?: unknown): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
  });
}

function sessionResponse(sessionId = SESSION_ID, sessionKey = SESSION_KEY): Response {
  return response(201, {
    ok: true, sessionId, sessionKey,
    protocolVersion: "1.0", expiresInSeconds: 60,
    workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
  });
}

function planTicket(argumentsValue: unknown = EXECUTE_ARGS, envelope: { readonly trusted?: unknown } = { trusted: TRUSTED }): Response {
  return response(200, {
    ticketId: TICKET_ID, requestId: REQUEST_ID,
    tool: "excel_execute_plan", arguments: argumentsValue,
    ...(envelope.trusted === undefined ? {} : { trusted: envelope.trusted }),
  });
}

type FetchRoute = (init: RequestInit | undefined) => Promise<Response> | Response;

/** Route-keyed fetch double; unknown routes fail loudly. */
function routedFetch(routes: Record<string, FetchRoute[]>) {
  const calls: { url: string; init?: RequestInit }[] = [];
  const fetchImpl = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    calls.push({ url, init });
    const key = `${init?.method ?? "GET"} ${url}`;
    const queue = routes[key];
    if (queue === undefined || queue.length === 0) throw new Error(`unexpected fetch ${key}`);
    const next = queue.length > 1 ? queue.shift()! : queue[0];
    return next(init);
  });
  return { fetchImpl: fetchImpl as unknown as typeof fetch, calls, urls: () => calls.map((call) => call.url) };
}

function resultBodies(calls: { url: string; init?: RequestInit }[]): Record<string, unknown>[] {
  return calls.filter((call) => call.url === "/excel/result").map((call) => JSON.parse(String(call.init?.body)));
}

function bridgeWith(fetchImpl: typeof fetch, executeTool: ConstructorParameters<typeof OfficeCompanionBridge>[0]["executeTool"]) {
  return new OfficeCompanionBridge({
    hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
    executeTool, fetchImpl,
  });
}

afterEach(() => {
  vi.useRealTimers();
});

describe("OfficeCompanionBridge Task 13 plan lifecycle", () => {
  it("does not acknowledge an ask plan before invoking Office and posts executed:false with the result", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: false, action: "workbook_plan", reason: "ask" }));
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(sessionResponse())
      .mockResolvedValueOnce(planTicket({ version: 1, outcome: "ask", message: "bounded" }))
      .mockResolvedValueOnce(response(200, { ok: true, ticketId: TICKET_ID }));
    const bridge = bridgeWith(fetchMock as typeof fetch, executeTool);

    await bridge.connect();
    await expect(bridge.pollOnce()).resolves.toBe(true);
    expect(executeTool).toHaveBeenCalledWith(
      "excel_execute_plan",
      { version: 1, outcome: "ask", message: "bounded" },
      TRUSTED,
      expect.objectContaining({ acknowledgeExecuteStart: expect.any(Function) }),
    );
    expect(fetchMock.mock.calls.map(([url]) => url)).toEqual([
      "/excel/session", "/excel/pending", "/excel/result",
    ]);
    expect(JSON.parse(String(fetchMock.mock.calls[2][1].body))).toEqual({
      ticketId: TICKET_ID,
      executed: false,
      result: { ok: false, action: "workbook_plan", reason: "ask" },
    });
  });

  it("fails closed on a plan with a missing or stale schema version before Office and posts executed:false/invalid", async () => {
    const executeTool = vi.fn();
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(sessionResponse())
      .mockResolvedValueOnce(planTicket({ outcome: "ask", message: "invalid" }))
      .mockResolvedValueOnce(response(200, { ok: true, ticketId: TICKET_ID }));
    const bridge = bridgeWith(fetchMock as typeof fetch, executeTool);

    await bridge.connect();
    await expect(bridge.pollOnce()).resolves.toBe(true);
    expect(executeTool).not.toHaveBeenCalled();
    expect(fetchMock).toHaveBeenCalledTimes(3);
    expect(fetchMock.mock.calls.map(([url]) => url)).toEqual([
      "/excel/session", "/excel/pending", "/excel/result",
    ]);
    const body = JSON.parse(String(fetchMock.mock.calls[2][1].body));
    expect(body).toMatchObject({ ticketId: TICKET_ID, executed: false, result: { ok: false, reason: "invalid" } });
  });

  it.each([
    ["missing envelope", undefined],
    ["extra envelope key", { ...TRUSTED, extra: true }],
    ["foreign fingerprint", { ...TRUSTED, workbookPlanFingerprint: "0".repeat(64) }],
  ])("posts executed:false/invalid for a plan ticket with a %s instead of waiting out the lease", async (_label, trusted) => {
    const executeTool = vi.fn();
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket(EXECUTE_ARGS, { trusted })],
      "POST /excel/result": [() => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);

    await bridge.connect();
    await expect(bridge.pollOnce()).resolves.toBe(true);
    expect(executeTool).not.toHaveBeenCalled();
    expect(routes.urls()).toEqual(["/excel/session", "/excel/pending", "/excel/result"]);
    expect(resultBodies(routes.calls)).toEqual([
      expect.objectContaining({ ticketId: TICKET_ID, executed: false, result: expect.objectContaining({ ok: false, reason: "invalid" }) }),
    ]);
  });

  it("acknowledges execute-start only when the executor gate calls it, then posts executed:true", async () => {
    const order: string[] = [];
    const executeTool = vi.fn(async (_name: string, _args: Record<string, unknown>, _trusted?: unknown, control?: BridgePlanExecutionControl) => {
      order.push("route-start");
      await control!.acknowledgeExecuteStart();
      order.push("mutation");
      return JSON.stringify({ ok: true, action: "workbook_plan" });
    });
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket()],
      "POST /excel/execute-start": [() => { order.push("execute-start"); return response(200, { ok: true, ticketId: TICKET_ID }); }],
      "POST /excel/result": [() => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);

    await bridge.connect();
    await expect(bridge.pollOnce()).resolves.toBe(true);
    expect(order).toEqual(["route-start", "execute-start", "mutation"]);
    expect(routes.urls()).toEqual(["/excel/session", "/excel/pending", "/excel/execute-start", "/excel/result"]);
    expect(JSON.parse(String(routes.calls[2].init?.body))).toEqual({ ticketId: TICKET_ID });
    expect(resultBodies(routes.calls)).toEqual([
      { ticketId: TICKET_ID, executed: true, result: { ok: true, action: "workbook_plan" } },
    ]);
  });

  it("refuses a second acknowledgement locally with no network request", async () => {
    let second: unknown;
    const executeTool = vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
      await control!.acknowledgeExecuteStart();
      second = await control!.acknowledgeExecuteStart().then(() => "resolved", (error: unknown) => error);
      return JSON.stringify({ ok: true });
    });
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket()],
      "POST /excel/execute-start": [() => response(200, { ok: true, ticketId: TICKET_ID })],
      "POST /excel/result": [() => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    await bridge.pollOnce();

    expect(second).toBeInstanceOf(Error);
    expect(routes.urls().filter((url) => url === "/excel/execute-start")).toHaveLength(1);
    expect(resultBodies(routes.calls)[0]).toMatchObject({ executed: true });
  });

  it.each([
    ["409", () => response(409, { error: COMPANION_SECRET })],
    ["404", () => response(404, { error: COMPANION_SECRET })],
    ["500", () => response(500, { error: COMPANION_SECRET })],
    ["network rejection", () => Promise.reject(new TypeError(COMPANION_SECRET))],
  ])("rejects the acknowledgement on %s with a fixed message and posts executed:false", async (_label, ack) => {
    let ackError: unknown;
    const executeTool = vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
      ackError = await control!.acknowledgeExecuteStart().then(() => undefined, (error: unknown) => error);
      return JSON.stringify({ ok: false, action: "workbook_plan", reason: "not_started" });
    });
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket()],
      "POST /excel/execute-start": [ack],
      "POST /excel/result": [() => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    await expect(bridge.pollOnce()).resolves.toBe(true);

    expect(ackError).toBeInstanceOf(Error);
    expect((ackError as Error).message).not.toContain(COMPANION_SECRET);
    const [body] = resultBodies(routes.calls);
    expect(body).toMatchObject({ ticketId: TICKET_ID, executed: false });
    expect(JSON.stringify(body)).not.toContain(COMPANION_SECRET);
  });

  it("always settles: a hanging execute-start is aborted at EXECUTE_START_ACK_TIMEOUT_MS", async () => {
    const timeout = (officeBridge as { EXECUTE_START_ACK_TIMEOUT_MS?: number }).EXECUTE_START_ACK_TIMEOUT_MS;
    expect(typeof timeout).toBe("number");
    // Well below the companion's 20 s claim lease (protocol.py CLAIM_LEASE_SECONDS).
    expect(timeout!).toBeGreaterThan(0);
    expect(timeout!).toBeLessThanOrEqual(10_000);
    let ackSignal: AbortSignal | undefined;
    let ackSettled: "pending" | "resolved" | "rejected" = "pending";
    const executeTool = vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
      await control!.acknowledgeExecuteStart().then(() => { ackSettled = "resolved"; }, () => { ackSettled = "rejected"; });
      return JSON.stringify({ ok: false, action: "workbook_plan", reason: "not_started" });
    });
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket()],
      "POST /excel/execute-start": [(init) => {
        ackSignal = init?.signal ?? undefined;
        return new Promise<Response>(() => undefined);
      }],
      "POST /excel/result": [() => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    const polled = bridge.pollOnce();
    // Flush real macrotasks only: vi.waitFor would advance the faked setTimeout clock.
    for (let attempt = 0; attempt < 200 && ackSignal === undefined; attempt += 1) {
      await new Promise<void>((resolve) => setImmediate(resolve));
    }
    expect(ackSignal).toBeDefined();
    await vi.advanceTimersByTimeAsync(timeout! - 1);
    expect(ackSettled).toBe("pending");
    await vi.advanceTimersByTimeAsync(1);
    await expect(polled).resolves.toBe(true);
    expect(ackSettled).toBe("rejected");
    expect(ackSignal?.aborted).toBe(true);
    expect(resultBodies(routes.calls)[0]).toMatchObject({ executed: false });
  });

  it("races the acknowledgement against a supplied turn AbortSignal", async () => {
    const turn = new AbortController();
    let ackSignal: AbortSignal | undefined;
    let ackError: unknown;
    const executeTool = vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
      const pending = control!.acknowledgeExecuteStart(turn.signal).then(() => undefined, (error: unknown) => error);
      await vi.waitFor(() => expect(ackSignal).toBeDefined());
      turn.abort();
      ackError = await pending;
      return JSON.stringify({ ok: false });
    });
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket()],
      "POST /excel/execute-start": [(init) => {
        ackSignal = init?.signal ?? undefined;
        return new Promise<Response>(() => undefined);
      }],
      "POST /excel/result": [() => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    await expect(bridge.pollOnce()).resolves.toBe(true);
    expect(ackError).toBeInstanceOf(Error);
    expect(ackSignal?.aborted).toBe(true);
    expect(resultBodies(routes.calls)[0]).toMatchObject({ executed: false });
  });

  it("rejects an acknowledgement with an already-aborted turn signal without a network request", async () => {
    const turn = new AbortController();
    turn.abort();
    let ackError: unknown;
    const executeTool = vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
      ackError = await control!.acknowledgeExecuteStart(turn.signal).then(() => undefined, (error: unknown) => error);
      return JSON.stringify({ ok: false });
    });
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket()],
      "POST /excel/result": [() => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    await bridge.pollOnce();
    expect(ackError).toBeInstanceOf(Error);
    expect(routes.urls()).not.toContain("/excel/execute-start");
  });

  it("rejects the acknowledgement locally after bridge stop() bumps the lifecycle generation", async () => {
    let ackError: unknown;
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket()],
      "DELETE /excel/session": [() => response(200, { ok: true })],
      "POST /excel/result": [() => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    let bridge: OfficeCompanionBridge | undefined;
    const executeTool = vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
      await bridge!.stop();
      ackError = await control!.acknowledgeExecuteStart().then(() => undefined, (error: unknown) => error);
      return JSON.stringify({ ok: false });
    });
    bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    await bridge.pollOnce().catch(() => undefined);
    expect(ackError).toBeInstanceOf(Error);
    expect(routes.urls()).not.toContain("/excel/execute-start");
  });

  it("binds the acknowledgement to the captured session: a 401 auto-reconnect rejects it without a network request", async () => {
    const handlers: (() => void)[] = [];
    let ackError: unknown;
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse(), () => sessionResponse(SESSION_ID_2, SESSION_KEY_2)],
      "GET /excel/pending": [() => planTicket()],
      "POST /excel/heartbeat": [() => response(401, { error: "Unauthorized" })],
      "POST /excel/result": [() => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    let bridge: OfficeCompanionBridge | undefined;
    const executeTool = vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
      handlers[0]();
      await vi.waitFor(() => expect(bridge!.sessionId).toBe(SESSION_ID_2));
      ackError = await control!.acknowledgeExecuteStart().then(() => undefined, (error: unknown) => error);
      return JSON.stringify({ ok: false });
    });
    bridge = new OfficeCompanionBridge({
      hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
      executeTool, fetchImpl: routes.fetchImpl,
      setIntervalImpl: ((handler: () => void) => { handlers.push(handler); return handlers.length; }) as unknown as typeof setInterval,
      clearIntervalImpl: (() => undefined) as unknown as typeof clearInterval,
    });
    await bridge.start();
    await bridge.pollOnce();

    expect(ackError).toBeInstanceOf(Error);
    expect(routes.urls()).not.toContain("/excel/execute-start");
    expect(resultBodies(routes.calls)[0]).toMatchObject({ executed: false });
  });

  it("rejects a late acknowledgement after executeTool has settled", async () => {
    let captured: BridgePlanExecutionControl | undefined;
    const executeTool = vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
      captured = control;
      return JSON.stringify({ ok: false, reason: "ask" });
    });
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket()],
      "POST /excel/result": [() => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    await bridge.pollOnce();
    expect(captured).toBeDefined();
    await expect(captured!.acknowledgeExecuteStart()).rejects.toThrow();
    expect(routes.urls()).not.toContain("/excel/execute-start");
  });

  it.each([
    ["before", false],
    ["after", true],
  ])("a thrown route error %s the acknowledgement posts a bounded content-free result", async (_label, acknowledge) => {
    const executeTool = vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
      if (acknowledge) await control!.acknowledgeExecuteStart();
      throw new Error(THROWN_SECRET);
    });
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket()],
      "POST /excel/execute-start": [() => response(200, { ok: true, ticketId: TICKET_ID })],
      "POST /excel/result": [() => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    await expect(bridge.pollOnce()).resolves.toBe(true);
    const [body] = resultBodies(routes.calls);
    expect(body.executed).toBe(acknowledge);
    expect(body.result).toMatchObject({ ok: false });
    if (!acknowledge) expect(body.result).toMatchObject({ reason: "rejected" });
    expect(JSON.stringify(body)).not.toContain(THROWN_SECRET);
  });

  it("mixed version: a companion that rejects the executed field never sees a retried mutation", async () => {
    const executeTool = vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
      await control!.acknowledgeExecuteStart();
      return JSON.stringify({ ok: true, action: "workbook_plan" });
    });
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket()],
      "POST /excel/execute-start": [() => response(200, { ok: true, ticketId: TICKET_ID })],
      "POST /excel/result": [() => response(400, { error: "Bad Request: unexpected field executed" })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    await expect(bridge.pollOnce()).rejects.toThrow();
    await expect(bridge.pollOnce()).rejects.toThrow();
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(routes.urls().filter((url) => url === "/excel/execute-start")).toHaveLength(1);
    expect(routes.urls().filter((url) => url === "/excel/pending")).toHaveLength(1);
    const bodies = resultBodies(routes.calls);
    expect(bodies).toHaveLength(2);
    expect(bodies[1]).toEqual(bodies[0]);
    expect(bodies[0]).toMatchObject({ executed: true });
  });
});

/**
 * R3b A2 (R2d-1 B9): at most MAX_RESULT_POST_ATTEMPTS = 3 result POSTs per
 * ticket; after a 413 the remaining attempts post a fixed, content-free body
 * with the same `executed` value. Appended only; the cases above are unchanged.
 */
describe("OfficeCompanionBridge R3b bounded result re-post", () => {
  const RESULT_CANARY = "WORKBOOK_RESULT_CANARY_R3B_6D1";
  const OK_RESULT = () => response(200, { ok: true, ticketId: TICKET_ID });
  const TOO_LARGE = () => response(413, { error: "Payload Too Large" });

  function primitiveTicket(): Response {
    return response(200, {
      ticketId: TICKET_ID, requestId: REQUEST_ID,
      tool: "read_range", arguments: { sheet: "Sheet1", address: "A1" },
    });
  }

  function acknowledgingPlan(acknowledge: boolean) {
    return vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
      if (acknowledge) await control!.acknowledgeExecuteStart();
      return JSON.stringify({ ok: true, action: "workbook_plan", result: { values: [[RESULT_CANARY]] } });
    });
  }

  async function settle(promise: Promise<boolean>): Promise<string> {
    return promise.then(
      (value) => `resolved:${String(value)}`,
      (error: unknown) => `rejected:${error instanceof Error ? error.message : String(error)}`,
    );
  }

  it.each([
    ["plan executed:true", true, {
      ticketId: TICKET_ID, executed: true,
      result: { ok: false, action: "workbook_plan", category: "Execution", state: "uncertain", error: "Workbook plan execution ended without a confirmed result." },
    }],
    ["plan executed:false", false, {
      ticketId: TICKET_ID, executed: false,
      result: { ok: false, action: "workbook_plan", category: "Execution", reason: "rejected", error: "Workbook plan was not executed." },
    }],
  ] as const)("RB7b %s: a 413 makes the next attempt post the fixed content-free body with the same executed value", async (_label, acknowledge, expected) => {
    const executeTool = acknowledgingPlan(acknowledge);
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket(), () => response(204)],
      "POST /excel/execute-start": [() => response(200, { ok: true, ticketId: TICKET_ID })],
      "POST /excel/result": [TOO_LARGE, OK_RESULT],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();

    expect(await settle(bridge.pollOnce())).toBe("rejected:Payload Too Large");
    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    expect(await settle(bridge.pollOnce())).toBe("resolved:false");

    const bodies = resultBodies(routes.calls);
    expect(bodies).toHaveLength(2);
    expect(bodies[0]).toMatchObject({ ticketId: TICKET_ID, executed: acknowledge });
    expect(JSON.stringify(bodies[0])).toContain(RESULT_CANARY);
    expect(bodies[1]).toEqual(expected);
    expect(JSON.stringify(bodies[1])).not.toContain(RESULT_CANARY);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(routes.urls().filter((url) => url === "/excel/execute-start")).toHaveLength(acknowledge ? 1 : 0);
  });

  it("RB7b primitive: a 413 makes the next attempt post the fixed primitive body with no executed key", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: [[RESULT_CANARY]] }));
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => primitiveTicket(), () => response(204)],
      "POST /excel/result": [TOO_LARGE, OK_RESULT],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();

    expect(await settle(bridge.pollOnce())).toBe("rejected:Payload Too Large");
    expect(await settle(bridge.pollOnce())).toBe("resolved:true");

    const bodies = resultBodies(routes.calls);
    expect(bodies).toHaveLength(2);
    expect(bodies[0]).toEqual({ ticketId: TICKET_ID, result: { ok: true, data: [[RESULT_CANARY]] } });
    expect(bodies[1]).toEqual({ ticketId: TICKET_ID, result: { ok: false, category: "Execution", error: "Office tool result was too large to submit." } });
    expect(Object.keys(bodies[1])).toEqual(["ticketId", "result"]);
    expect(executeTool).toHaveBeenCalledTimes(1);
  });

  it("RB7b 413 x3: exactly 3 posts (1 original + 2 fixed), then the entry is dropped and the next poll reaches pending", async () => {
    const executeTool = acknowledgingPlan(true);
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket(), () => response(204)],
      "POST /excel/execute-start": [() => response(200, { ok: true, ticketId: TICKET_ID })],
      "POST /excel/result": [TOO_LARGE],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();

    const outcomes: string[] = [];
    for (let poll = 0; poll < 6; poll += 1) outcomes.push(await settle(bridge.pollOnce()));

    expect(outcomes).toEqual([
      "rejected:Payload Too Large", "rejected:Payload Too Large", "rejected:Payload Too Large",
      "resolved:false", "resolved:false", "resolved:false",
    ]);
    const bodies = resultBodies(routes.calls);
    expect(bodies).toHaveLength(3);
    expect(JSON.stringify(bodies[0])).toContain(RESULT_CANARY);
    expect(bodies[1]).toEqual(bodies[2]);
    expect(bodies[1]).toMatchObject({ executed: true, result: { state: "uncertain" } });
    expect(bodies.every((body) => body.executed === true)).toBe(true);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(routes.urls().filter((url) => url === "/excel/execute-start")).toHaveLength(1);
  });

  it.each([
    ["400", () => response(400, { error: "Bad Request" }), "Bad Request"],
    ["500", () => response(500, { error: "Internal Server Error" }), "Internal Server Error"],
    ["network rejection", () => Promise.reject(new TypeError("network down")), "network down"],
  ] as const)("RB7 %s: at most 3 identical posts, each failing poll rethrows the original error, then pending is reached", async (_label, failure, message) => {
    const executeTool = acknowledgingPlan(true);
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket(), () => response(204)],
      "POST /excel/execute-start": [() => response(200, { ok: true, ticketId: TICKET_ID })],
      "POST /excel/result": [failure],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();

    const outcomes: string[] = [];
    for (let poll = 0; poll < 10; poll += 1) outcomes.push(await settle(bridge.pollOnce()));

    expect(outcomes.slice(0, 3)).toEqual([`rejected:${message}`, `rejected:${message}`, `rejected:${message}`]);
    expect(outcomes.slice(3)).toEqual(Array.from({ length: 7 }, () => "resolved:false"));
    const bodies = resultBodies(routes.calls);
    expect(bodies).toHaveLength(3);
    expect(bodies[1]).toEqual(bodies[0]);
    expect(bodies[2]).toEqual(bodies[0]);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(routes.urls().filter((url) => url === "/excel/pending")).toHaveLength(8);
  });

  it.each([
    ["404", 404],
    ["409", 409],
  ] as const)("GB9a %s still drops the entry after one post and the next poll reaches pending", async (_label, status) => {
    const executeTool = acknowledgingPlan(true);
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket(), () => response(204)],
      "POST /excel/execute-start": [() => response(200, { ok: true, ticketId: TICKET_ID })],
      "POST /excel/result": [() => response(status, { error: COMPANION_SECRET })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();

    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    expect(await settle(bridge.pollOnce())).toBe("resolved:false");
    expect(resultBodies(routes.calls)).toHaveLength(1);
    expect(routes.urls()).toEqual([
      "/excel/session", "/excel/pending", "/excel/execute-start", "/excel/result", "/excel/pending",
    ]);
  });

  it("GB9a a network error then 200 succeeds on attempt 2 with byte-identical bodies and an unchanged executed value", async () => {
    const executeTool = acknowledgingPlan(true);
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket(), () => response(204)],
      "POST /excel/execute-start": [() => response(200, { ok: true, ticketId: TICKET_ID })],
      "POST /excel/result": [() => Promise.reject(new TypeError("reset")), () => response(500, { error: "busy" }), OK_RESULT],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();

    expect(await settle(bridge.pollOnce())).toBe("rejected:reset");
    expect(await settle(bridge.pollOnce())).toBe("rejected:busy");
    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    expect(await settle(bridge.pollOnce())).toBe("resolved:false");
    const raw = routes.calls.filter((call) => call.url === "/excel/result").map((call) => String(call.init?.body));
    expect(raw).toHaveLength(3);
    expect(raw[1]).toBe(raw[0]);
    expect(raw[2]).toBe(raw[0]);
    expect(resultBodies(routes.calls).every((body) => body.executed === true)).toBe(true);
    expect(executeTool).toHaveBeenCalledTimes(1);
  });

  it("GB9a a 200 on attempt 1 posts exactly once with today's body", async () => {
    const executeTool = acknowledgingPlan(true);
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket(), () => response(204)],
      "POST /excel/execute-start": [() => response(200, { ok: true, ticketId: TICKET_ID })],
      "POST /excel/result": [OK_RESULT],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();

    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    expect(await settle(bridge.pollOnce())).toBe("resolved:false");
    const raw = routes.calls.filter((call) => call.url === "/excel/result").map((call) => String(call.init?.body));
    expect(raw).toEqual([JSON.stringify({
      ticketId: TICKET_ID, executed: true,
      result: { ok: true, action: "workbook_plan", result: { values: [[RESULT_CANARY]] } },
    })]);
  });
});
