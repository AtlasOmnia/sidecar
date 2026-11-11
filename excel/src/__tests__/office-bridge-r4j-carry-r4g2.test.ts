import { describe, expect, it, vi } from "vitest";
import { OfficeCompanionBridge, type BridgePlanExecutionControl, type BridgePublicState } from "../officeBridge";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";

/**
 * R4g-2b (R4j lane-A A1/A2 carried into R4g-2) through the real OfficeCompanionBridge.
 * - A1 (BP3-plan): an S1 plan executeTool that acknowledged execute-start and is still held keeps
 *   the bridge-wide tool counter closed across a stop()+start() rotation: S2 timer ticks issue no
 *   GET /excel/pending until the S1 result is posted with S1 credentials. Kills a5/a5i/a5l.
 * - A2 (J5 on a 401 rotation): a stuck S1 timer GET does not block the timer poll of the session
 *   that a heartbeat 401 auto-reconnected to (same generation). Kills a4 (guard keyed by generation).
 * Both rows pass at the base by construction (test-only pins); RED = the mutants in r4g2b/.
 * Constants and helpers below (through rotateWhileHeld) are COPIED as text from
 * office-bridge-session-scope-r4j.test.ts:19-207 at eedc0ca (that file is neither imported nor
 * edited). No fake timers, no real-time waits; interval handlers are injected (timedBridge pushes
 * 2 handlers per start(): heartbeat, poll). Diagnostics print counts and session ids, never bodies.
 * Release: dev3-r4g2b-release.md (G1-G5).
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
// R4g-2b A1/A2
// ---------------------------------------------------------------------------
describe("R4g-2b (R4j lane-A A1/A2): plan-path tool counter across rotation; per-session poll guard after a same-generation 401 reconnect", () => {
  it("A1 BP3-plan: an S1 plan executeTool held after its ack keeps S2 timer ticks from polling; after it settles the next tick polls S2", async () => {
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
      [`POST /excel/result ${SESSION_ID}`]: [() => response(200, { ok: true })],
    }, executeTool, () => reached);
    for (let tick = 0; tick < 3; tick += 1) h.handlers[3]();
    await flushMacrotasks();
    expect(h.f.pendingGets(SESSION_ID_2), h.f.diag()).toBe(0);
    held.resolve();
    await waitUntil(() => h.f.posts(TICKET_ID, SESSION_ID).length === 1, h.f.diag);
    await flushMacrotasks();
    h.handlers[3]();
    await waitUntil(() => h.f.pendingGets(SESSION_ID_2) === 1, h.f.diag);
    expect(executeTool).toHaveBeenCalledTimes(1);
    expect(h.f.posts(TICKET_ID, SESSION_ID), h.f.diag()).toHaveLength(1);
    expect(h.f.posts(TICKET_ID, SESSION_ID_2), h.f.diag()).toHaveLength(0);
  });

  it("A2 401 auto-reconnect (same generation): a stuck S1 timer GET does not block the reconnected S2's timer poll", async () => {
    const stuck = deferred<Response>();
    const f = sessionFetch({
      ...common(),
      [`POST /excel/heartbeat ${SESSION_ID}`]: [() => response(401, { error: "Unauthorized" })],
      [`GET /excel/pending ${SESSION_ID}`]: [() => stuck.promise],
      [`GET /excel/pending ${SESSION_ID_2}`]: [() => response(204)],
    });
    const { bridge, handlers } = timedBridge(f.fetchImpl, vi.fn(async () => "{}"));
    await bridge.start();
    handlers[1]();
    await waitUntil(() => f.pendingGets(SESSION_ID) === 1, f.diag);
    handlers[0]();
    await waitUntil(() => bridge.sessionId === SESSION_ID_2, f.diag);
    expect(handlers).toHaveLength(2);
    await flushMacrotasks();
    handlers[1]();
    await waitUntil(() => f.pendingGets(SESSION_ID_2) === 1, f.diag);
    await flushMacrotasks();
    handlers[1]();
    await flushMacrotasks();
    expect(f.pendingGets(SESSION_ID_2), f.diag()).toBe(2);
    expect(f.pendingGets(SESSION_ID), f.diag()).toBe(1);
  });
});
