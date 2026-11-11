import { describe, expect, it, vi } from "vitest";
import { OfficeCompanionBridge, type BridgePlanExecutionControl, type BridgePublicState } from "../officeBridge";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";

/**
 * R4f bridge-unit rows (release dev3-r4f-release.md §5): A2 (a 413 downgrades
 * whatever its body; a transport failure on a >10 MiB result post downgrades
 * too), A5 (background errors from a superseded session are ignored),
 * R4-TX8b (a ticket already dispatched on this session is never re-executed or
 * re-posted) and the lane-A A3 stop() pin. Everything runs through the real
 * OfficeCompanionBridge with a route-keyed fetch double copied from
 * office-bridge-task13.test.ts and injected interval handlers. Large bodies are
 * built once; diagnostics print lengths, never bodies.
 */

const SESSION_ID = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const SESSION_KEY = "s".repeat(64);
const SESSION_ID_2 = "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee";
const SESSION_KEY_2 = "t".repeat(64);
const SESSION_ID_3 = "ffffffff-ffff-4fff-8fff-ffffffffffff";
const SESSION_KEY_3 = "u".repeat(64);
const PANE_ID = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
const TICKET_ID = "cccccccc-cccc-4ccc-8ccc-cccccccccccc";
const TICKET_ID_2 = "cccccccc-cccc-4ccc-8ccc-ccccccccccc2";
const REQUEST_ID = "dddddddd-dddd-4ddd-8ddd-dddddddddddd";
const R4G2_TURN_KEY = "4b4b0000-0000-4000-8000-0000000000a3"; // R4g-2a rule M: turnKey of migrated v2 tickets
const TRUSTED = { turnOperationId: REQUEST_ID, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT, phase: "execute", turnKey: R4G2_TURN_KEY };
const EXECUTE_ARGS = { version: 1, outcome: "execute", summary: "s", operations: [], verification: [] };
const RESULT_CANARY = "WORKBOOK_RESULT_CANARY_R4F_7E2";
/** Mirror of scripts/companion/protocol.py MAX_RESULT_BODY. */
const MAX_RESULT_BODY_BYTES = 10 * 1024 * 1024;
const MIB = 2 ** 20;
/** ≈10.5 MiB of ASCII with the canary in the middle (built once). */
const BIG_ASCII = `${"x".repeat(5.25 * MIB)}${RESULT_CANARY}${"x".repeat(5.25 * MIB)}`;
/** > 10 MiB of UTF-8 but < 10 Mi UTF-16 code units (built once). */
const BIG_UTF8 = "é".repeat(5.5 * MIB);

const FIXED_PRIMITIVE = { ok: false, category: "Execution", error: "Office tool result was too large to submit." };
const FIXED_UNCERTAIN = { ok: false, action: "workbook_plan", category: "Execution", state: "uncertain", error: "Workbook plan execution ended without a confirmed result." };
const FIXED_REJECTED = { ok: false, action: "workbook_plan", category: "Execution", reason: "rejected", error: "Workbook plan was not executed." };

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

function primitiveTicket(ticketId = TICKET_ID, extra: Record<string, unknown> = {}): Response {
  return response(200, {
    ticketId, requestId: REQUEST_ID,
    tool: "read_range", arguments: { sheet: "Sheet1", address: "A1", ticket: ticketId },
    ...extra,
  });
}

type FetchRoute = (init: RequestInit | undefined) => Promise<Response> | Response;

/** Route-keyed fetch double; unknown routes fail loudly (copied from office-bridge-task13.test.ts). */
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

function rawResults(calls: { url: string; init?: RequestInit }[]): string[] {
  return calls.filter((call) => call.url === "/excel/result").map((call) => String(call.init?.body));
}

function resultBodies(calls: { url: string; init?: RequestInit }[]): Record<string, unknown>[] {
  return rawResults(calls).map((raw) => JSON.parse(raw) as Record<string, unknown>);
}

function count(urls: string[], url: string, calls?: { url: string; init?: RequestInit }[], method?: string): number {
  if (calls !== undefined && method !== undefined) return calls.filter((call) => call.url === url && (call.init?.method ?? "GET") === method).length;
  return urls.filter((entry) => entry === url).length;
}

function bridgeWith(fetchImpl: typeof fetch, executeTool: ConstructorParameters<typeof OfficeCompanionBridge>[0]["executeTool"]) {
  return new OfficeCompanionBridge({
    hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
    executeTool, fetchImpl,
  });
}

function timedBridge(fetchImpl: typeof fetch, executeTool: ConstructorParameters<typeof OfficeCompanionBridge>[0]["executeTool"]) {
  const handlers: (() => void)[] = [];
  const states: BridgePublicState[] = [];
  const bridge = new OfficeCompanionBridge({
    hostPrefix: "excel", documentId: async () => "doc-1", paneId: PANE_ID,
    executeTool, fetchImpl,
    onState: (state) => { states.push({ ...state }); },
    setIntervalImpl: ((handler: () => void) => { handlers.push(handler); return handlers.length; }) as unknown as typeof setInterval,
    clearIntervalImpl: (() => undefined) as unknown as typeof clearInterval,
  });
  return { bridge, handlers, states };
}

async function settle(promise: Promise<boolean>): Promise<string> {
  return promise.then(
    (value) => `resolved:${String(value)}`,
    (error: unknown) => `rejected:${error instanceof Error ? error.message : String(error)}`,
  );
}

async function caught(promise: Promise<unknown>): Promise<unknown> {
  return promise.then(() => undefined, (error: unknown) => error);
}

async function flushMacrotasks(turns = 30): Promise<void> {
  for (let turn = 0; turn < turns; turn += 1) await new Promise<void>((done) => setImmediate(done));
}

async function waitUntil(condition: () => boolean, diag: () => string = () => ""): Promise<void> {
  for (let attempt = 0; attempt < 400 && !condition(); attempt += 1) {
    await new Promise<void>((done) => setImmediate(done));
  }
  expect(condition(), diag()).toBe(true);
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((res, rej) => { resolve = res; reject = rej; });
  return { promise, resolve, reject };
}

function acknowledgingPlan(acknowledge: boolean, payload: unknown = { values: [[RESULT_CANARY]] }) {
  return vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
    if (acknowledge) await control!.acknowledgeExecuteStart();
    return JSON.stringify({ ok: true, action: "workbook_plan", result: payload });
  });
}

function utf8Bytes(text: string): number {
  return new TextEncoder().encode(text).length;
}

function lengths(raw: string[]): string {
  return `posts=${raw.length} lengths=${JSON.stringify(raw.map((entry) => entry.length))} canary=${JSON.stringify(raw.map((entry) => entry.includes(RESULT_CANARY)))}`;
}

describe("R4f A2: a 413 downgrades whatever its body; a transport failure on a >10 MiB post downgrades", () => {
  it("BA2a plan executed:true, an HTML 413: poll rejects 'Result submission failed' and post 2 is the fixed uncertain body", async () => {
    const executeTool = acknowledgingPlan(true);
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket(), () => response(204)],
      "POST /excel/execute-start": [() => response(200, { ok: true, ticketId: TICKET_ID })],
      "POST /excel/result": [
        () => new Response("<html>Payload Too Large</html>", { status: 413, headers: { "Content-Type": "text/html" } }),
        () => response(200, { ok: true, ticketId: TICKET_ID }),
      ],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();

    const first = await settle(bridge.pollOnce());
    const second = await settle(bridge.pollOnce());
    const third = await settle(bridge.pollOnce());
    const raw = rawResults(routes.calls);
    const diag = `outcomes=${JSON.stringify([first, second, third])} ${lengths(raw)}`;
    expect(first, diag).toBe("rejected:Result submission failed");
    expect(raw.length, diag).toBe(2);
    expect(raw[0].includes(RESULT_CANARY), diag).toBe(true);
    expect(raw[1].includes(RESULT_CANARY), diag).toBe(false);
    expect(JSON.parse(raw[1]), diag).toEqual({ ticketId: TICKET_ID, executed: true, result: FIXED_UNCERTAIN });
    expect([second, third], diag).toEqual(["resolved:true", "resolved:false"]);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(count(routes.urls(), "/excel/execute-start")).toBe(1);
  });

  it("BA2c primitive, a 413 whose body stream rejects: post 2 is the fixed primitive body with keys [ticketId, result]", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: [[RESULT_CANARY]] }));
    const brokenStream = { status: 413, text: () => Promise.reject(new TypeError("body stream failed")) } as unknown as Response;
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => primitiveTicket(), () => response(204)],
      "POST /excel/result": [() => brokenStream, () => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();

    const outcomes = [await settle(bridge.pollOnce()), await settle(bridge.pollOnce())];
    const raw = rawResults(routes.calls);
    const diag = `outcomes=${JSON.stringify(outcomes)} ${lengths(raw)}`;
    expect(raw.length, diag).toBe(2);
    expect(raw[1].includes(RESULT_CANARY), diag).toBe(false);
    const body = JSON.parse(raw[1]) as Record<string, unknown>;
    expect(body, diag).toEqual({ ticketId: TICKET_ID, result: FIXED_PRIMITIVE });
    expect(Object.keys(body)).toEqual(["ticketId", "result"]);
    expect(outcomes, diag).toEqual(["rejected:Result submission failed", "resolved:true"]);
    expect(executeTool).toHaveBeenCalledTimes(1);
  });

  it("BA2d plan executed:false, a ~10.5 MiB result and a fetch rejection: same error object rethrown, post 2 fixed rejected", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: false, action: "workbook_plan", reason: "ask", message: BIG_ASCII }));
    const failure = new TypeError("Failed to fetch");
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket(), () => response(204)],
      "POST /excel/result": [() => Promise.reject(failure), () => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();

    const error = await caught(bridge.pollOnce());
    const second = await settle(bridge.pollOnce());
    const third = await settle(bridge.pollOnce());
    const raw = rawResults(routes.calls);
    const diag = `second=${second} third=${third} ${lengths(raw)} firstBytes=${raw[0] === undefined ? -1 : utf8Bytes(raw[0])}`;
    expect(error, diag).toBe(failure);
    expect(raw.length, diag).toBe(2);
    expect(raw.length).toBeLessThanOrEqual(3);
    expect(utf8Bytes(raw[0]), diag).toBeGreaterThan(MAX_RESULT_BODY_BYTES);
    expect(raw[1].length, diag).toBeLessThan(1024);
    expect(JSON.parse(raw[1]), diag).toEqual({ ticketId: TICKET_ID, executed: false, result: FIXED_REJECTED });
    expect([second, third], diag).toEqual(["resolved:true", "resolved:false"]);
    expect(executeTool).toHaveBeenCalledTimes(1);
  }, 60_000);

  it("BA2u primitive, > 10 MiB UTF-8 but < 10 Mi chars, a rejected post: post 2 is fixed (bytes, not text.length)", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: BIG_UTF8 }));
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => primitiveTicket(), () => response(204)],
      "POST /excel/result": [() => Promise.reject(new TypeError("connection reset")), () => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();

    const outcomes = [await settle(bridge.pollOnce()), await settle(bridge.pollOnce())];
    const raw = rawResults(routes.calls);
    const diag = `outcomes=${JSON.stringify(outcomes)} ${lengths(raw)}`;
    expect(raw.length, diag).toBe(2);
    expect(raw[0].length, diag).toBeLessThan(MAX_RESULT_BODY_BYTES);
    expect(utf8Bytes(raw[0]), diag).toBeGreaterThan(MAX_RESULT_BODY_BYTES);
    expect(raw[1].length, diag).toBeLessThan(1024);
    expect(JSON.parse(raw[1]), diag).toEqual({ ticketId: TICKET_ID, result: FIXED_PRIMITIVE });
    expect(outcomes, diag).toEqual(["rejected:connection reset", "resolved:true"]);
  }, 60_000);

  it("BA2p (parity) a JSON 413 keeps the base path: rejection 'Payload Too Large', post 2 fixed", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: [[RESULT_CANARY]] }));
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => primitiveTicket(), () => response(204)],
      "POST /excel/result": [() => response(413, { error: "Payload Too Large" }), () => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    expect(await settle(bridge.pollOnce())).toBe("rejected:Payload Too Large");
    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    const bodies = resultBodies(routes.calls);
    expect(bodies).toHaveLength(2);
    expect(bodies[1]).toEqual({ ticketId: TICKET_ID, result: FIXED_PRIMITIVE });
  });

  it("BA2e (parity) an empty-body 413 makes post 2 the fixed body", async () => {
    const executeTool = acknowledgingPlan(false);
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket(), () => response(204)],
      "POST /excel/result": [() => response(413), () => response(200, { ok: true, ticketId: TICKET_ID })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    expect(await settle(bridge.pollOnce())).toBe("rejected:Result submission failed");
    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    const bodies = resultBodies(routes.calls);
    expect(bodies).toHaveLength(2);
    expect(bodies[1]).toEqual({ ticketId: TICKET_ID, executed: false, result: FIXED_REJECTED });
  });

  it.each([
    ["1 KiB", 1024],
    ["exactly MAX_RESULT_BODY_BYTES (boundary, not >)", MAX_RESULT_BODY_BYTES],
  ] as const)("BA2s (parity) a %s body with a network rejection x3: 3 identical original posts, then dropped", async (_label, targetBytes) => {
    const overhead = JSON.stringify({ ticketId: TICKET_ID, result: "" }).length;
    const filler = "x".repeat(targetBytes - overhead);
    const executeTool = vi.fn(async () => filler);
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => primitiveTicket(), () => response(204)],
      "POST /excel/result": [() => Promise.reject(new TypeError("network down"))],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    const outcomes: string[] = [];
    for (let poll = 0; poll < 5; poll += 1) outcomes.push(await settle(bridge.pollOnce()));
    const raw = rawResults(routes.calls);
    const diag = `outcomes=${JSON.stringify(outcomes)} ${lengths(raw)}`;
    expect(raw.length, diag).toBe(3);
    expect(utf8Bytes(raw[0]), diag).toBe(targetBytes);
    expect(raw[1] === raw[0] && raw[2] === raw[0], diag).toBe(true);
    expect(outcomes, diag).toEqual([
      "rejected:network down", "rejected:network down", "rejected:network down", "resolved:false", "resolved:false",
    ]);
    expect(executeTool).toHaveBeenCalledTimes(1);
  }, 60_000);

  it("BA2n (parity, OQ-F5) a non-413 non-JSON 500 still rejects with SyntaxError and re-posts the original", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: [[RESULT_CANARY]] }));
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => primitiveTicket(), () => response(204)],
      "POST /excel/result": [
        () => new Response("<html>Internal Server Error</html>", { status: 500, headers: { "Content-Type": "text/html" } }),
        () => response(200, { ok: true, ticketId: TICKET_ID }),
      ],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    const error = await caught(bridge.pollOnce());
    expect(error).toBeInstanceOf(SyntaxError);
    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    const raw = rawResults(routes.calls);
    expect(raw).toHaveLength(2);
    expect(raw[1]).toBe(raw[0]);
    expect(raw[0]).toContain(RESULT_CANARY);
  });
});

describe("R4f A5: background errors from a superseded session are ignored", () => {
  function heldPendingRoutes(release: Promise<Response>) {
    return routedFetch({
      "POST /excel/session": [() => sessionResponse(), () => sessionResponse(SESSION_ID_2, SESSION_KEY_2), () => sessionResponse(SESSION_ID_3, SESSION_KEY_3)],
      "GET /excel/pending": [() => release, () => response(204)],
      "DELETE /excel/session": [() => response(200, { ok: true })],
      "POST /excel/heartbeat": [() => response(200, { ok: true })],
    });
  }

  async function staleS1Poll(release: Promise<Response>) {
    const routes = heldPendingRoutes(release);
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true }));
    const { bridge, handlers, states } = timedBridge(routes.fetchImpl, executeTool);
    await bridge.start();
    expect(bridge.sessionId).toBe(SESSION_ID);
    handlers[1]();
    await waitUntil(() => count(routes.urls(), "/excel/pending") === 1);
    await bridge.stop();
    await bridge.start();
    expect(bridge.sessionId).toBe(SESSION_ID_2);
    return { routes, executeTool, bridge, handlers, states, sessionPosts: count(routes.urls(), "/excel/session", routes.calls, "POST"), statesBefore: states.length };
  }

  it("BA5a an S1 poll 401 that lands after stop()+start() onto S2: no onState, no reconnect, S2 kept and still polled", async () => {
    const held = deferred<Response>();
    const h = await staleS1Poll(held.promise);
    held.resolve(response(401, { error: "Unauthorized" }));
    await flushMacrotasks();
    const diag = `session=${String(h.bridge.sessionId)} posts=${count(h.routes.urls(), "/excel/session", h.routes.calls, "POST")} late=${JSON.stringify(h.states.slice(h.statesBefore))}`;
    expect(h.bridge.sessionId, diag).toBe(SESSION_ID_2);
    expect(count(h.routes.urls(), "/excel/session", h.routes.calls, "POST"), diag).toBe(h.sessionPosts);
    expect(h.states.slice(h.statesBefore), diag).toEqual([]);
    h.handlers[3]();
    await waitUntil(() => count(h.routes.urls(), "/excel/pending") === 2, () => diag);
    await flushMacrotasks();
    const pendingCalls = h.routes.calls.filter((call) => call.url === "/excel/pending");
    const lastPending = pendingCalls[pendingCalls.length - 1];
    expect((lastPending.init?.headers as Record<string, string>)["X-Hermes-Session-Id"]).toBe(SESSION_ID_2);
    expect(h.bridge.sessionId).toBe(SESSION_ID_2);
  });

  it("BA5b an S1 poll network error that lands after the session was replaced: no onState call", async () => {
    const held = deferred<Response>();
    const h = await staleS1Poll(held.promise);
    held.reject(new TypeError("Failed to fetch"));
    await flushMacrotasks();
    const diag = `session=${String(h.bridge.sessionId)} late=${JSON.stringify(h.states.slice(h.statesBefore))}`;
    expect(h.states.slice(h.statesBefore), diag).toEqual([]);
    expect(h.bridge.sessionId, diag).toBe(SESSION_ID_2);
    expect(count(h.routes.urls(), "/excel/session", h.routes.calls, "POST")).toBe(h.sessionPosts);
  });

  it("BA5c (parity) a current-session 401 through the poll handler reconnects to S2 with the base onState sequence", async () => {
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse(), () => sessionResponse(SESSION_ID_2, SESSION_KEY_2)],
      "GET /excel/pending": [() => response(401, { error: "Unauthorized" }), () => response(204)],
      "POST /excel/heartbeat": [() => response(200, { ok: true })],
    });
    const { bridge, handlers, states } = timedBridge(routes.fetchImpl, vi.fn(async () => "{}"));
    await bridge.start();
    handlers[1]();
    await waitUntil(() => bridge.sessionId === SESSION_ID_2);
    await flushMacrotasks();
    expect(states).toEqual([
      { connected: true, sessionId: SESSION_ID, documentId: "doc-1" },
      { connected: false, sessionId: SESSION_ID, documentId: "doc-1", lastError: "Unauthorized" },
      { connected: true, sessionId: SESSION_ID_2, documentId: "doc-1" },
    ]);
    expect(count(routes.urls(), "/excel/session", routes.calls, "POST")).toBe(2);
  });
});

describe("R4f R4-TX8b: a ticket already dispatched on this bridge session is never re-executed or re-posted", () => {
  const OK_RESULT = () => response(200, { ok: true, ticketId: TICKET_ID });

  it("BTX1 a plan ticket acknowledged and posted 200, handed out again: second poll false, executeTool 1, one post, one execute-start", async () => {
    const executeTool = acknowledgingPlan(true);
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket(), () => planTicket(), () => response(204)],
      "POST /excel/execute-start": [() => response(200, { ok: true, ticketId: TICKET_ID })],
      "POST /excel/result": [OK_RESULT],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    const outcomes = [await settle(bridge.pollOnce()), await settle(bridge.pollOnce()), await settle(bridge.pollOnce())];
    const diag = `outcomes=${JSON.stringify(outcomes)} executeTool=${executeTool.mock.calls.length} urls=${JSON.stringify(routes.urls())}`;
    expect(outcomes, diag).toEqual(["resolved:true", "resolved:false", "resolved:false"]);
    expect(executeTool, diag).toHaveBeenCalledTimes(1);
    expect(resultBodies(routes.calls), diag).toEqual([
      expect.objectContaining({ ticketId: TICKET_ID, executed: true }),
    ]);
    expect(count(routes.urls(), "/excel/execute-start"), diag).toBe(1);
  });

  it("BTX2 primitive requeue path: the result post fails 500 x3 (dropped), the ticket is redelivered: no re-execution, no 4th post", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: [[1]] }));
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => primitiveTicket(), () => primitiveTicket(), () => response(204)],
      "POST /excel/result": [() => response(500, { error: "Internal Server Error" })],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    const outcomes: string[] = [];
    for (let poll = 0; poll < 5; poll += 1) outcomes.push(await settle(bridge.pollOnce()));
    const diag = `outcomes=${JSON.stringify(outcomes)} executeTool=${executeTool.mock.calls.length} posts=${rawResults(routes.calls).length}`;
    expect(outcomes.slice(0, 3), diag).toEqual(Array.from({ length: 3 }, () => "rejected:Internal Server Error"));
    expect(outcomes.slice(3), diag).toEqual(["resolved:false", "resolved:false"]);
    expect(executeTool, diag).toHaveBeenCalledTimes(1);
    expect(rawResults(routes.calls), diag).toHaveLength(3);
    expect(count(routes.urls(), "/excel/pending"), diag).toBe(3);
  });

  it("BTX3 two concurrent polls both receive the same ticket while executeTool is held: executeTool once, the second poll resolves false", async () => {
    const gate = deferred<string>();
    const executeTool = vi.fn(() => gate.promise);
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => primitiveTicket(), () => primitiveTicket(), () => response(204)],
      "POST /excel/result": [OK_RESULT],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    const first = settle(bridge.pollOnce());
    const second = settle(bridge.pollOnce());
    await waitUntil(() => count(routes.urls(), "/excel/pending") === 2);
    await flushMacrotasks();
    const callsWhileHeld = executeTool.mock.calls.length;
    gate.resolve(JSON.stringify({ ok: true, data: [[1]] }));
    const outcomes = [await first, await second];
    const diag = `callsWhileHeld=${callsWhileHeld} outcomes=${JSON.stringify(outcomes)} posts=${rawResults(routes.calls).length}`;
    expect(callsWhileHeld, diag).toBe(1);
    expect(executeTool, diag).toHaveBeenCalledTimes(1);
    expect(outcomes, diag).toEqual(["resolved:true", "resolved:false"]);
    expect(rawResults(routes.calls), diag).toHaveLength(1);
  });

  it("BTX6 the record is bounded at 32 ids per session (FIFO): after 33 tickets, #33 is dropped and #1 executes again", async () => {
    const ids = Array.from({ length: 33 }, (_unused, index) => `cccccccc-cccc-4ccc-8ccc-${String(index + 1).padStart(12, "0")}`);
    const executed: string[] = [];
    const executeTool = vi.fn(async (_name: string, args: Record<string, unknown>) => {
      executed.push(String(args.ticket));
      return JSON.stringify({ ok: true, data: [[1]] });
    });
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [...ids.map((id) => () => primitiveTicket(id)), () => primitiveTicket(ids[32]), () => primitiveTicket(ids[0]), () => response(204)],
      "POST /excel/result": [OK_RESULT],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    const firstRound: string[] = [];
    for (let poll = 0; poll < 33; poll += 1) firstRound.push(await settle(bridge.pollOnce()));
    expect(firstRound).toEqual(Array.from({ length: 33 }, () => "resolved:true"));
    const redeliver33 = await settle(bridge.pollOnce());
    const redeliver1 = await settle(bridge.pollOnce());
    const tail = await settle(bridge.pollOnce());
    const diag = `#33=${redeliver33} #1=${redeliver1} tail=${tail} executed=${executed.length} last=${JSON.stringify(executed.slice(-3))}`;
    expect(redeliver33, diag).toBe("resolved:false");
    expect(redeliver1, diag).toBe("resolved:true");
    expect(tail, diag).toBe("resolved:false");
    expect(executed, diag).toEqual([...ids, ids[0]]);
    expect(rawResults(routes.calls), diag).toHaveLength(34);
  });

  it("BTX4 (parity) after ticket T, a new ticket T' executes normally", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: [[1]] }));
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => primitiveTicket(), () => primitiveTicket(TICKET_ID_2), () => response(204)],
      "POST /excel/result": [OK_RESULT],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    expect([await settle(bridge.pollOnce()), await settle(bridge.pollOnce()), await settle(bridge.pollOnce())])
      .toEqual(["resolved:true", "resolved:true", "resolved:false"]);
    expect(executeTool).toHaveBeenCalledTimes(2);
    expect(resultBodies(routes.calls).map((body) => body.ticketId)).toEqual([TICKET_ID, TICKET_ID_2]);
  });

  it("BTX5 (parity) per-session scope: after stop() and connect() on a new session the same ticket id executes", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: [[1]] }));
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse(), () => sessionResponse(SESSION_ID_2, SESSION_KEY_2)],
      "GET /excel/pending": [() => primitiveTicket(), () => primitiveTicket(), () => response(204)],
      "DELETE /excel/session": [() => response(200, { ok: true })],
      "POST /excel/result": [OK_RESULT],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    await bridge.stop();
    await bridge.connect();
    expect(bridge.sessionId).toBe(SESSION_ID_2);
    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    expect(executeTool).toHaveBeenCalledTimes(2);
    expect(rawResults(routes.calls)).toHaveLength(2);
  });

  it("BTX7 (parity) an invalid plan ticket delivered twice posts invalid twice and never reaches executeTool (not recorded)", async () => {
    const executeTool = vi.fn();
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => planTicket([]), () => planTicket([]), () => response(204)],
      "POST /excel/result": [OK_RESULT],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    expect([await settle(bridge.pollOnce()), await settle(bridge.pollOnce()), await settle(bridge.pollOnce())])
      .toEqual(["resolved:true", "resolved:true", "resolved:false"]);
    expect(executeTool).not.toHaveBeenCalled();
    expect(resultBodies(routes.calls)).toEqual([
      expect.objectContaining({ ticketId: TICKET_ID, executed: false, result: expect.objectContaining({ reason: "invalid" }) }),
      expect.objectContaining({ ticketId: TICKET_ID, executed: false, result: expect.objectContaining({ reason: "invalid" }) }),
    ]);
  });

  it("BTX8 (parity) a primitive with a trusted envelope posts the fixed primitive error before dispatch, and its redelivery posts it again (not recorded)", async () => {
    const executeTool = vi.fn();
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "GET /excel/pending": [() => primitiveTicket(TICKET_ID, { trusted: TRUSTED }), () => primitiveTicket(TICKET_ID, { trusted: TRUSTED }), () => response(204)],
      "POST /excel/result": [OK_RESULT],
    });
    const bridge = bridgeWith(routes.fetchImpl, executeTool);
    await bridge.connect();
    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    expect(await settle(bridge.pollOnce())).toBe("resolved:true");
    expect(executeTool).not.toHaveBeenCalled();
    const fixed = JSON.stringify({ ticketId: TICKET_ID, result: { ok: false, category: "Execution", error: "Office tool execution failed" } });
    expect(rawResults(routes.calls)).toEqual([fixed, fixed]);
  });
});

describe("R4f lane-A A3 pin: stop() keeps the best-effort DELETE", () => {
  it.each([
    ["a DELETE network TypeError", () => Promise.reject(new TypeError("network down"))],
    ["a DELETE 500", () => response(500, { error: "Internal Server Error" })],
  ] as const)("BS1 stop() with %s resolves, clears the session, sends exactly one DELETE and never retries", async (_label, failure) => {
    const routes = routedFetch({
      "POST /excel/session": [() => sessionResponse()],
      "DELETE /excel/session": [failure],
    });
    const bridge = bridgeWith(routes.fetchImpl, vi.fn());
    await bridge.connect();
    await expect(bridge.stop()).resolves.toBeUndefined();
    await flushMacrotasks();
    expect(bridge.sessionId).toBeNull();
    expect(count(routes.urls(), "/excel/session", routes.calls, "DELETE")).toBe(1);
    expect(routes.calls).toHaveLength(2);
  });
});
