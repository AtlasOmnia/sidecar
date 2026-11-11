import { describe, expect, it, vi } from "vitest";
import { OfficeCompanionBridge, type BridgePlanExecutionControl, type BridgePublicState } from "../officeBridge";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";

/**
 * R4j J4/J5 (R4f lane-B B2 and B4) through the real OfficeCompanionBridge.
 * - J4: a cached result carries the session that produced it; a replaced session's entry is dropped
 *   unposted at the next current-session poll, never re-posted with the new session's credentials.
 * - J5: the poll timer guard is per (session, generation); a stuck replaced-session poll no longer
 *   blocks the new session's timer polling, while executeTool stays serialized bridge-wide.
 * RED at the base: BR1, BR1b, BR3, BP1, BP4. Parity pins (identical at the base and HEAD): BR2, BR4,
 * BR5, BP2, BP3, BP5, BP6.
 * Constants and helpers (response, sessionResponse, planTicket, primitiveTicket, routedFetch, count,
 * bridgeWith, timedBridge, settle, flushMacrotasks, waitUntil, deferred, acknowledgingPlan) are COPIED
 * as text from office-bridge-lifecycle-r4f.test.ts:16-165 at f3e3c0a (that file is neither imported
 * nor edited); sessionFetch is a local session-aware variant. timedBridge pushes 2 handlers per
 * start() (heartbeat, poll). Diagnostics print counts and session ids, never bodies.
 */

const SESSION_ID = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const SESSION_KEY = "s".repeat(64);
const SESSION_ID_2 = "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee";
const SESSION_KEY_2 = "t".repeat(64);
const PANE_ID = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
const TICKET_ID = "cccccccc-cccc-4ccc-8ccc-cccccccccccc";
const TICKET_ID_2 = "cccccccc-cccc-4ccc-8ccc-ccccccccccc2";
const REQUEST_ID = "dddddddd-dddd-4ddd-8ddd-dddddddddddd";
const R4G2_TURN_KEY = "4b4b0000-0000-4000-8000-0000000000a3"; // R4g-2a rule M: turnKey of migrated v2 tickets
const TRUSTED = { turnOperationId: REQUEST_ID, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT, phase: "execute", turnKey: R4G2_TURN_KEY };
const EXECUTE_ARGS = { version: 1, outcome: "execute", summary: "s", operations: [], verification: [] };
const RESULT_CANARY = "WORKBOOK_RESULT_CANARY_R4F_7E2";

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

// ---------------------------------------------------------------------------
// Local session-aware fetch double (R4j)
// ---------------------------------------------------------------------------
interface SeenCall {
  readonly method: string;
  readonly url: string;
  readonly session: string;
  readonly ticketId: string | null;
  readonly keys: string[] | null;
  readonly hasSignal: boolean;
}

const short = (session: string): string => (session === SESSION_ID ? "S1" : session === SESSION_ID_2 ? "S2" : session);

/** Routes are keyed `${method} ${url} ${sessionId}` first, then `${method} ${url}`; unknown routes fail loudly. */
function sessionFetch(routes: Record<string, FetchRoute[]>) {
  const calls: SeenCall[] = [];
  const fetchImpl = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    const method = init?.method ?? "GET";
    const session = (init?.headers as Record<string, string> | undefined)?.["X-Hermes-Session-Id"] ?? "-";
    let ticketId: string | null = null;
    let keys: string[] | null = null;
    if (url === "/excel/result" && typeof init?.body === "string") {
      const parsed = JSON.parse(init.body) as Record<string, unknown>;
      ticketId = typeof parsed.ticketId === "string" ? parsed.ticketId : null;
      keys = Object.keys(parsed);
    }
    calls.push({ method, url, session, ticketId, keys, hasSignal: init?.signal !== undefined });
    const queue = routes[`${method} ${url} ${session}`] ?? routes[`${method} ${url}`];
    if (queue === undefined || queue.length === 0) throw new Error(`unexpected fetch ${method} ${url} ${short(session)}`);
    const next = queue.length > 1 ? queue.shift()! : queue[0];
    return next(init);
  });
  const pendingGets = (session: string): number => calls.filter((call) => call.url === "/excel/pending" && call.session === session).length;
  const posts = (ticketId?: string, session?: string): SeenCall[] => calls.filter((call) =>
    call.url === "/excel/result" && (ticketId === undefined || call.ticketId === ticketId) && (session === undefined || call.session === session));
  const diag = (): string => `calls=${JSON.stringify(calls.map((call) => `${call.method} ${call.url} ${short(call.session)}${call.ticketId === null ? "" : ` ${call.ticketId === TICKET_ID ? "T1" : "T2"}`}`))}`;
  return { fetchImpl: fetchImpl as unknown as typeof fetch, calls, pendingGets, posts, diag };
}

const ticketBody = (init: RequestInit | undefined): string | null => {
  const parsed = JSON.parse(String(init?.body)) as Record<string, unknown>;
  return typeof parsed.ticketId === "string" ? parsed.ticketId : null;
};

/** S2 result route: an orphan T1 posted with S2 credentials is answered 404 (as the companion would); T2 is accepted. */
const S2_RESULT: FetchRoute = (init) => (ticketBody(init) === TICKET_ID ? response(404, { error: "Unknown ticket" }) : response(200, { ok: true }));

/** Fresh queues per test (the double shifts them). */
const common = (): Record<string, FetchRoute[]> => ({
  "POST /excel/session": [() => sessionResponse(), () => sessionResponse(SESSION_ID_2, SESSION_KEY_2)],
  "DELETE /excel/session": [() => response(200, { ok: true })],
  "POST /excel/heartbeat": [() => response(200, { ok: true })],
});

/** Starts S1, holds its first timer poll inside `hold`, rotates to S2 (stop + start). */
async function rotateWhileHeld(routes: Record<string, FetchRoute[]>, executeTool: ConstructorParameters<typeof OfficeCompanionBridge>[0]["executeTool"], holdReached: (f: ReturnType<typeof sessionFetch>) => boolean) {
  const f = sessionFetch({ ...common(), ...routes });
  const { bridge, handlers, states } = timedBridge(f.fetchImpl, executeTool);
  await bridge.start();
  expect(bridge.sessionId).toBe(SESSION_ID);
  handlers[1]();
  await waitUntil(() => holdReached(f), f.diag);
  await bridge.stop();
  await bridge.start();
  expect(bridge.sessionId).toBe(SESSION_ID_2);
  expect(handlers).toHaveLength(4);
  return { f, bridge, handlers, states };
}

// ---------------------------------------------------------------------------
// J4: per-session result cache
// ---------------------------------------------------------------------------
describe("R4j J4: a replaced session's cached result is dropped unposted, never re-posted with new credentials", () => {
  it.each([
    ["BR1", 401],
    ["BR1b", 500],
  ] as const)("%s primitive: S1 executeTool outlives the rotation, its S1 post fails %s; S2's poll drops T1, polls pending and serves T2 with S2 headers", async (_id, status) => {
    const held = deferred<string>();
    let toolCalls = 0;
    const executeTool = vi.fn(async () => {
      toolCalls += 1;
      return toolCalls === 1 ? held.promise : JSON.stringify({ ok: true, data: [[2]] });
    });
    const h = await rotateWhileHeld({
      [`GET /excel/pending ${SESSION_ID}`]: [() => primitiveTicket(TICKET_ID)],
      [`GET /excel/pending ${SESSION_ID_2}`]: [() => primitiveTicket(TICKET_ID_2), () => response(204)],
      [`POST /excel/result ${SESSION_ID}`]: [() => response(status, { error: "stale session" })],
      [`POST /excel/result ${SESSION_ID_2}`]: [S2_RESULT],
    }, executeTool, () => toolCalls === 1);
    held.resolve(JSON.stringify({ ok: true, data: [[1]] }));
    await waitUntil(() => h.f.posts(TICKET_ID, SESSION_ID).length === 1, h.f.diag);
    await flushMacrotasks();
    const outcome = await settle(h.bridge.pollOnce());
    await flushMacrotasks();
    const diag = `outcome=${outcome} ${h.f.diag()}`;
    expect(h.f.posts(TICKET_ID, SESSION_ID_2), diag).toHaveLength(0);
    expect(h.f.posts(TICKET_ID), diag).toHaveLength(1);
    expect(h.f.pendingGets(SESSION_ID_2), diag).toBe(1);
    expect(h.f.posts(TICKET_ID_2).map((call) => short(call.session)), diag).toEqual(["S2"]);
    expect(outcome, diag).toBe("resolved:true");
    expect(executeTool, diag).toHaveBeenCalledTimes(2);
  });

  it("BR3 plan: S1 acknowledged (executed:true), rotation during executeTool, the S1 post fails 500; no S2-header post for T1", async () => {
    const held = deferred<void>();
    let reached = false;
    const executeTool = vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
      await control!.acknowledgeExecuteStart();
      reached = true;
      await held.promise;
      return JSON.stringify({ ok: true, action: "workbook_plan", result: { values: [[RESULT_CANARY]] } });
    });
    const h = await rotateWhileHeld({
      [`GET /excel/pending ${SESSION_ID}`]: [() => planTicket()],
      [`GET /excel/pending ${SESSION_ID_2}`]: [() => response(204)],
      [`POST /excel/execute-start ${SESSION_ID}`]: [() => response(200, { ok: true, ticketId: TICKET_ID })],
      [`POST /excel/result ${SESSION_ID}`]: [() => response(500, { error: "Internal Server Error" })],
      [`POST /excel/result ${SESSION_ID_2}`]: [S2_RESULT],
    }, executeTool, () => reached);
    held.resolve();
    await waitUntil(() => h.f.posts(TICKET_ID, SESSION_ID).length === 1, h.f.diag);
    await flushMacrotasks();
    const outcome = await settle(h.bridge.pollOnce());
    const diag = `outcome=${outcome} ${h.f.diag()}`;
    expect(h.f.posts(TICKET_ID, SESSION_ID_2), diag).toHaveLength(0);
    expect(h.f.posts(TICKET_ID), diag).toHaveLength(1);
    expect(h.f.posts(TICKET_ID)[0]?.keys, diag).toEqual(["ticketId", "executed", "result"]);
    expect(h.f.pendingGets(SESSION_ID_2), diag).toBe(1);
    expect(outcome, diag).toBe("resolved:false");
    expect(executeTool).toHaveBeenCalledTimes(1);
  });

  it("BR2 (parity) an own-session entry answered 500 x3: three posts with the same session's headers, then dropped and pending polled", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: [[1]] }));
    const f = sessionFetch({
      ...common(),
      "GET /excel/pending": [() => primitiveTicket(TICKET_ID), () => response(204)],
      "POST /excel/result": [() => response(500, { error: "Internal Server Error" })],
    });
    const bridge = bridgeWith(f.fetchImpl, executeTool);
    await bridge.connect();
    const outcomes: string[] = [];
    for (let poll = 0; poll < 4; poll += 1) outcomes.push(await settle(bridge.pollOnce()));
    const diag = `outcomes=${JSON.stringify(outcomes)} ${f.diag()}`;
    expect(outcomes, diag).toEqual([
      "rejected:Internal Server Error", "rejected:Internal Server Error", "rejected:Internal Server Error", "resolved:false",
    ]);
    expect(f.posts(TICKET_ID).map((call) => short(call.session)), diag).toEqual(["S1", "S1", "S1"]);
    expect(f.pendingGets(SESSION_ID), diag).toBe(2);
    expect(executeTool).toHaveBeenCalledTimes(1);
  });

  it("BR4 (parity) a current-session 401 on the cached result's post through the poll handler clears the cache and reconnects (BA5c onState sequence)", async () => {
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: [[1]] }));
    const f = sessionFetch({
      ...common(),
      [`GET /excel/pending ${SESSION_ID}`]: [() => primitiveTicket(TICKET_ID)],
      [`GET /excel/pending ${SESSION_ID_2}`]: [() => response(204)],
      [`POST /excel/result ${SESSION_ID}`]: [() => response(401, { error: "Unauthorized" })],
      [`POST /excel/result ${SESSION_ID_2}`]: [S2_RESULT],
    });
    const { bridge, handlers, states } = timedBridge(f.fetchImpl, executeTool);
    await bridge.start();
    handlers[1]();
    await waitUntil(() => bridge.sessionId === SESSION_ID_2, f.diag);
    await flushMacrotasks();
    handlers[1]();
    await waitUntil(() => f.pendingGets(SESSION_ID_2) === 1, f.diag);
    await flushMacrotasks();
    expect(states).toEqual([
      { connected: true, sessionId: SESSION_ID, documentId: "doc-1" },
      { connected: false, sessionId: SESSION_ID, documentId: "doc-1", lastError: "Unauthorized" },
      { connected: true, sessionId: SESSION_ID_2, documentId: "doc-1" },
    ]);
    expect(f.posts(TICKET_ID).map((call) => short(call.session)), f.diag()).toEqual(["S1"]);
    expect(f.calls.filter((call) => call.url === "/excel/session" && call.method === "POST")).toHaveLength(2);
  });

  it("BR5 (parity) posted JSON keys unchanged (primitive [ticketId, result], plan [ticketId, executed, result]); pending and result carry no signal", async () => {
    const primitive = sessionFetch({ ...common(), "GET /excel/pending": [() => primitiveTicket(TICKET_ID)], "POST /excel/result": [() => response(200, { ok: true })] });
    const a = bridgeWith(primitive.fetchImpl, vi.fn(async () => JSON.stringify({ ok: true, data: [[1]] })));
    await a.connect();
    expect(await settle(a.pollOnce())).toBe("resolved:true");
    const plan = sessionFetch({
      ...common(),
      "GET /excel/pending": [() => planTicket()],
      "POST /excel/execute-start": [() => response(200, { ok: true, ticketId: TICKET_ID })],
      "POST /excel/result": [() => response(200, { ok: true })],
    });
    const b = bridgeWith(plan.fetchImpl, acknowledgingPlan(true));
    await b.connect();
    expect(await settle(b.pollOnce())).toBe("resolved:true");
    expect(primitive.posts().map((call) => call.keys)).toEqual([["ticketId", "result"]]);
    expect(plan.posts().map((call) => call.keys)).toEqual([["ticketId", "executed", "result"]]);
    for (const call of [...primitive.calls, ...plan.calls]) {
      expect(call.hasSignal, `${call.method} ${call.url}`).toBe(call.url === "/excel/execute-start");
    }
  });
});

// ---------------------------------------------------------------------------
// J5: per-(session, generation) timer poll guard, executeTool serialized bridge-wide
// ---------------------------------------------------------------------------
describe("R4j J5: a stuck replaced-session poll no longer blocks the new session's timer polling", () => {
  it("BP1 an S1 timer GET /pending never settles; after stop()+start() the S2 poll handler polls with S2 headers and serves a ticket", async () => {
    const stuck = deferred<Response>();
    const executeTool = vi.fn(async () => JSON.stringify({ ok: true, data: [[2]] }));
    const h = await rotateWhileHeld({
      [`GET /excel/pending ${SESSION_ID}`]: [() => stuck.promise],
      [`GET /excel/pending ${SESSION_ID_2}`]: [() => primitiveTicket(TICKET_ID_2), () => response(204)],
      [`POST /excel/result ${SESSION_ID_2}`]: [S2_RESULT],
    }, executeTool, (f) => f.pendingGets(SESSION_ID) === 1);
    h.handlers[3]();
    await waitUntil(() => h.f.posts(TICKET_ID_2, SESSION_ID_2).length === 1, h.f.diag);
    expect(h.f.pendingGets(SESSION_ID_2), h.f.diag()).toBe(1);
    expect(h.f.pendingGets(SESSION_ID), h.f.diag()).toBe(1);
    expect(executeTool).toHaveBeenCalledTimes(1);
  });

  it.each([
    ["204", () => response(204)],
    ["401", () => response(401, { error: "Unauthorized" })],
  ] as const)("BP4 S1 GET stuck, S2 GET held; releasing the stale S1 GET (%s) never clears S2's guard: S2 GETs stay 1 until it settles, then the next tick polls", async (_label, release) => {
    const stuck = deferred<Response>();
    const s2Held = deferred<Response>();
    const h = await rotateWhileHeld({
      [`GET /excel/pending ${SESSION_ID}`]: [() => stuck.promise],
      [`GET /excel/pending ${SESSION_ID_2}`]: [() => s2Held.promise, () => response(204)],
    }, vi.fn(async () => "{}"), (f) => f.pendingGets(SESSION_ID) === 1);
    const f = h.f;
    h.handlers[3]();
    await waitUntil(() => f.pendingGets(SESSION_ID_2) === 1, f.diag);
    stuck.resolve(release());
    await flushMacrotasks();
    for (let tick = 0; tick < 3; tick += 1) h.handlers[3]();
    await flushMacrotasks();
    expect(f.pendingGets(SESSION_ID_2), f.diag()).toBe(1);
    expect(h.bridge.sessionId).toBe(SESSION_ID_2);
    s2Held.resolve(response(204));
    await flushMacrotasks();
    h.handlers[3]();
    await waitUntil(() => f.pendingGets(SESSION_ID_2) === 2, f.diag);
    expect(h.states.filter((state) => state.lastError !== undefined), f.diag()).toEqual([]);
  });

  it("BP2 (parity) same session: S1 GET held, the poll handler fired x3 issues exactly one GET", async () => {
    const held = deferred<Response>();
    const f = sessionFetch({ ...common(), "GET /excel/pending": [() => held.promise, () => response(204)] });
    const { bridge, handlers } = timedBridge(f.fetchImpl, vi.fn(async () => "{}"));
    await bridge.start();
    for (let tick = 0; tick < 3; tick += 1) handlers[1]();
    await flushMacrotasks();
    expect(f.pendingGets(SESSION_ID), f.diag()).toBe(1);
  });

  it("BP3 (parity) S1 executeTool held, rotate to S2: S2 ticks issue no GET (bridge-wide serialization); after release the next tick polls S2", async () => {
    const held = deferred<string>();
    let toolCalls = 0;
    const executeTool = vi.fn(async () => {
      toolCalls += 1;
      return held.promise;
    });
    const h = await rotateWhileHeld({
      [`GET /excel/pending ${SESSION_ID}`]: [() => primitiveTicket(TICKET_ID)],
      [`GET /excel/pending ${SESSION_ID_2}`]: [() => response(204)],
      [`POST /excel/result ${SESSION_ID}`]: [() => response(200, { ok: true })],
    }, executeTool, () => toolCalls === 1);
    for (let tick = 0; tick < 3; tick += 1) h.handlers[3]();
    await flushMacrotasks();
    expect(h.f.pendingGets(SESSION_ID_2), h.f.diag()).toBe(0);
    held.resolve(JSON.stringify({ ok: true, data: [[1]] }));
    await waitUntil(() => h.f.posts(TICKET_ID, SESSION_ID).length === 1, h.f.diag);
    await flushMacrotasks();
    h.handlers[3]();
    await waitUntil(() => h.f.pendingGets(SESSION_ID_2) === 1, h.f.diag);
    expect(executeTool).toHaveBeenCalledTimes(1);
  });

  async function failingTick(ticket: FetchRoute, executeStart: FetchRoute | undefined, executeTool: ConstructorParameters<typeof OfficeCompanionBridge>[0]["executeTool"]): Promise<void> {
    const f = sessionFetch({
      ...common(),
      "GET /excel/pending": [ticket, () => response(204)],
      "POST /excel/result": [() => response(200, { ok: true })],
      ...(executeStart === undefined ? {} : { "POST /excel/execute-start": [executeStart] }),
    });
    const { bridge, handlers } = timedBridge(f.fetchImpl, executeTool);
    await bridge.start();
    handlers[1]();
    await waitUntil(() => f.posts(TICKET_ID).length === 1, f.diag);
    await flushMacrotasks();
    handlers[1]();
    await waitUntil(() => f.pendingGets(SESSION_ID) === 2, f.diag);
    expect(executeTool).toHaveBeenCalledTimes(1);
  }

  it("BP5a (parity) primitive executeTool throws: the failure is posted and the next tick polls again", async () => {
    await failingTick(() => primitiveTicket(TICKET_ID), undefined, vi.fn(async (): Promise<string> => { throw new Error("office down"); }));
  });

  it("BP5b (parity) plan acknowledgement rejected and executeTool throws: the failure is posted and the next tick polls again", async () => {
    await failingTick(() => planTicket(), () => response(409, { error: "Conflict" }),
      vi.fn(async (_n: string, _a: Record<string, unknown>, _t?: unknown, control?: BridgePlanExecutionControl) => {
        await control!.acknowledgeExecuteStart();
        return "{}";
      }));
  });

  it("BP6 (parity) current-session steady state: a tick while a poll is in flight is skipped; after it settles the next tick polls", async () => {
    const held = deferred<Response>();
    const f = sessionFetch({ ...common(), "GET /excel/pending": [() => held.promise, () => response(204)] });
    const { bridge, handlers } = timedBridge(f.fetchImpl, vi.fn(async () => "{}"));
    await bridge.start();
    handlers[1]();
    handlers[1]();
    await flushMacrotasks();
    expect(f.pendingGets(SESSION_ID), f.diag()).toBe(1);
    held.resolve(response(204));
    await flushMacrotasks();
    handlers[1]();
    await waitUntil(() => f.pendingGets(SESSION_ID) === 2, f.diag);
    await flushMacrotasks();
    handlers[1]();
    await waitUntil(() => f.pendingGets(SESSION_ID) === 3, f.diag);
    expect(count(f.calls.map((call) => call.url), "/excel/result")).toBe(0);
  });
});
