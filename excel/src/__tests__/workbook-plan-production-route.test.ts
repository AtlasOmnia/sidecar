import { resolve } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import * as officeBridge from "../officeBridge";
import type { BridgePlanExecutionControl, BridgeTrustedPlanEnvelope, BridgeTrustedPlanEnvelopeV2 } from "../officeBridge";
import { createExcelCompanionBridge } from "../officeBridgeExcel";
import { ServerOwnedTurnBindingStore, WORKBOOK_PLAN_TURN_TIMEOUT_MS } from "../workbookPlanOrchestrator";
import {
  BrowserIdempotencyStore,
  MemoryIdempotencyStorage,
  type WorkbookPlanExecutionAdapter,
} from "../workbookPlanExecutor";
import type { WorkbookPlanVerificationAdapter } from "../workbookPlanOrchestrator";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import type { WorkbookPlanRouteRuntimeContext } from "../workbookPlanRoute";
import { classifyWorkbookRequest } from "../workbookPlanRequest";
import { TaskPaneTurnCoordinator } from "../taskpane/taskPaneTurnCoordinator";

type RouteModule = typeof import("../taskpane/serverOwnedPlanRoute");

/**
 * R2d-1 production route: the REAL createExcelCompanionBridge -> extracted
 * server-owned route factory -> createWorkbookPlanTurn ->
 * executeAcceptedWorkbookPlan chain, driven by a stateful R2d-2 companion
 * double and a stateful Office adapter. The route module is loaded by path so
 * a missing extraction is an assertion failure, not a collection error.
 */
async function loadRouteModule(): Promise<RouteModule | undefined> {
  const specifier = resolve(__dirname, "../taskpane/serverOwnedPlanRoute");
  try {
    return (await import(/* @vite-ignore */ specifier)) as RouteModule;
  } catch {
    return undefined;
  }
}

const DOC = "d0c00000-0000-4000-8000-000000000001";
const PANE = "9a9e0000-0000-4000-8000-000000000002";
const SERVER_OP = "5e7e0000-0000-4000-8000-000000000003";
const TURN = "7a7a0000-0000-4000-8000-000000000004";
const TICKET = "71c70000-0000-4000-8000-000000000005";
const REQUEST = "4e4e0000-0000-4000-8000-000000000006";
const COMPANION_SECRET = "COMPANION_SECRET_TEXT_9F1";
const WORKBOOK_SECRET = "WORKBOOK_SECRET_CELL_77A";

const COMPILE_CONTEXT: VerifiedWorkbookCompileContext = {
  sheetNames: new Set(["Sheet1"]),
  tablesByName: new Map(),
  supportedPrimitiveNames: new Set(["write_range"]),
  supportedCapabilities: new Set(["excel_workbook_access", "workbook_plan_v1_operation_audit"]),
  chartSourceResolutionsBySheet: new Map(),
};
const RUNTIME: WorkbookPlanRouteRuntimeContext = {
  activeSheetName: "Sheet1",
  selection: { sheetName: "Sheet1", address: "A1" },
  compileContext: COMPILE_CONTEXT,
};
const EXECUTE_PLAN = {
  version: 1,
  outcome: "execute",
  summary: "write one cell",
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }],
  verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["ok"]] }],
} as const;
/** Values shape does not match the address: live resource preflight rejects it before any mutation. */
const RESOURCE_REJECTED_PLAN = {
  ...EXECUTE_PLAN,
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok", "extra"]] }],
} as const;
const TRUSTED: BridgeTrustedPlanEnvelope = {
  turnOperationId: SERVER_OP,
  workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
};
/** R4g-2a rule M: the v2 ticket envelope for companion-delivered (bridge) tickets; TRUSTED stays for direct route calls. */
const TRUSTED_V2: BridgeTrustedPlanEnvelopeV2 = { ...TRUSTED, phase: "execute", turnKey: "4b4b0000-0000-4000-8000-0000000000a3" };
const TRUSTED_V2_INSPECT: BridgeTrustedPlanEnvelopeV2 = { ...TRUSTED_V2, phase: "inspect" };

type AckMode = "ok" | 404 | 409 | 500 | "network" | "hang";
/**
 * r2d2: one-way executed checks. legacy_ignores_executed: the real pre-R2d-2
 * (cd5d021) companion, which ignores the `executed` key and accepts a plan
 * result from claimed or executing as completed. legacy_rejects_executed: an
 * extra strict double that refuses the unknown key.
 */
type ResultMode = "r2d2" | "legacy_ignores_executed" | "legacy_rejects_executed";

interface DoubleTicket {
  readonly ticketId: string;
  readonly tool: string;
  readonly args: Record<string, unknown>;
  readonly trusted?: unknown;
  state: "pending" | "claimed" | "executing" | "terminal";
  terminalReason?: "completed" | "not_executed";
}

function json(status: number, body?: unknown): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
  });
}

/** Stateful R2d-2 companion lifecycle double (claimed/executing/terminal + executed checks). */
class CompanionDouble {
  readonly sessions = new Map<string, string>();
  ticket: DoubleTicket | null = null;
  readonly results: Record<string, unknown>[] = [];
  readonly ackSignals: (AbortSignal | undefined)[] = [];
  executeStartRequests = 0;
  heartbeatUnauthorized = false;
  /** When set, an accepted execute-start (ticket now executing) answers 200 only after this settles. */
  ackGate: Promise<void> | undefined;
  private counter = 0;

  constructor(
    private readonly events: string[],
    public ackMode: AckMode = "ok",
    public resultMode: ResultMode = "r2d2",
  ) {}

  enqueue(tool: string, args: Record<string, unknown>, trusted?: unknown): void {
    this.ticket = { ticketId: TICKET, tool, args, ...(trusted === undefined ? {} : { trusted }), state: "pending" };
  }

  readonly fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const url = String(input);
    const route = url.replace(/^.*\/excel\//, "");
    const method = init?.method ?? "GET";
    const headers = (init?.headers ?? {}) as Record<string, string>;
    const authed = this.sessions.has(headers["X-Hermes-Session-Key"] ?? "");
    const body = typeof init?.body === "string" && init.body ? JSON.parse(init.body) as Record<string, unknown> : {};
    if (method === "POST" && route === "session") {
      this.counter += 1;
      const sessionId = `aaaaaaaa-aaaa-4aaa-8aaa-${String(this.counter).padStart(12, "0")}`;
      const sessionKey = `${"k".repeat(40)}${this.counter}`;
      this.sessions.set(sessionKey, sessionId);
      this.events.push("session");
      return json(201, {
        ok: true, sessionId, sessionKey, protocolVersion: "1.0", expiresInSeconds: 60,
        workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
      });
    }
    if (!authed) return json(401, { error: "Unauthorized" });
    if (method === "DELETE" && route === "session") {
      this.sessions.delete(headers["X-Hermes-Session-Key"]);
      return json(200, { ok: true });
    }
    if (method === "POST" && route === "heartbeat") {
      if (this.heartbeatUnauthorized) {
        this.sessions.delete(headers["X-Hermes-Session-Key"]);
        return json(401, { error: "Unauthorized" });
      }
      return json(200, { ok: true });
    }
    if (method === "GET" && route === "pending") {
      const ticket = this.ticket;
      if (ticket === null || ticket.state !== "pending") return json(204);
      ticket.state = "claimed";
      this.events.push("companion-claim");
      return json(200, {
        ticketId: ticket.ticketId, requestId: REQUEST, tool: ticket.tool, arguments: ticket.args,
        ...(ticket.trusted === undefined ? {} : { trusted: ticket.trusted }),
      });
    }
    if (method === "POST" && route === "execute-start") {
      this.executeStartRequests += 1;
      this.ackSignals.push(init?.signal ?? undefined);
      const mode = this.ackMode;
      if (mode === "network") throw new TypeError(`${COMPANION_SECRET} connection reset`);
      if (mode === "hang") {
        return new Promise<Response>((_resolve, reject) => {
          init?.signal?.addEventListener("abort", () => reject(new DOMException(COMPANION_SECRET, "AbortError")), { once: true });
        });
      }
      if (typeof mode === "number") {
        this.events.push(`execute-start-${mode}`);
        return json(mode, { error: COMPANION_SECRET });
      }
      const ticket = this.ticket;
      if (ticket === null || ticket.ticketId !== body.ticketId || ticket.state !== "claimed") {
        return json(409, { error: COMPANION_SECRET });
      }
      ticket.state = "executing";
      this.events.push("execute-start");
      if (this.ackGate !== undefined) await this.ackGate;
      return json(200, { ok: true, ticketId: ticket.ticketId });
    }
    if (method === "POST" && route === "result") {
      this.events.push("result");
      this.results.push(body);
      const ticket = this.ticket;
      if (ticket === null || ticket.ticketId !== body.ticketId) return json(404, { error: "No ticket found" });
      if (this.resultMode === "legacy_rejects_executed" && "executed" in body) {
        return json(400, { error: "Bad Request: unexpected field executed" });
      }
      if (this.resultMode === "legacy_ignores_executed") {
        if (ticket.state !== "claimed" && ticket.state !== "executing") return json(409, { error: "conflict" });
        ticket.state = "terminal";
        ticket.terminalReason = "completed";
        return json(200, { ok: true, ticketId: ticket.ticketId });
      }
      if (ticket.tool !== "excel_execute_plan") {
        if (ticket.state !== "claimed") return json(409, { error: "conflict" });
        ticket.state = "terminal";
        ticket.terminalReason = "completed";
        return json(200, { ok: true, ticketId: ticket.ticketId });
      }
      if (typeof body.executed !== "boolean") return json(400, { error: "executed required" });
      if (body.executed === true && ticket.state !== "executing") return json(409, { error: "conflict" });
      if (body.executed === false && ticket.state !== "claimed") return json(409, { error: "conflict" });
      ticket.state = "terminal";
      ticket.terminalReason = body.executed === true ? "completed" : "not_executed";
      return json(200, { ok: true, ticketId: ticket.ticketId });
    }
    return json(404, { error: "route" });
  };
}

interface AdapterHooks {
  beforeGate?: () => Promise<void> | void;
  duringMutation?: () => Promise<void> | void;
  duringVerify?: () => Promise<void> | void;
}

interface VerifyBehaviour {
  /** The post-mutation verification read throws this message. */
  readonly throws?: string;
  /** The post-mutation verification read reports this result instead of reading the cell. */
  readonly result?: boolean;
}

function statefulAdapter(events: string[], hooks: AdapterHooks = {}, verify: VerifyBehaviour = {}) {
  const cells = new Map<string, unknown>();
  const adapter: WorkbookPlanExecutionAdapter & WorkbookPlanVerificationAdapter = {
    async executeFirstGroup(gate) {
      await hooks.beforeGate?.();
      events.push("fresh");
      return gate.executeAfterFresh("fresh", async (group) => {
        events.push("mutation");
        for (const operation of group.operations) {
          cells.set(`${String(operation.args.sheet)}!${String(operation.args.address)}`, operation.args.values);
        }
        await hooks.duringMutation?.();
        return { status: "committed" };
      });
    },
    async executeGroup(group) {
      events.push("mutation");
      for (const operation of group.operations) {
        cells.set(`${String(operation.args.sheet)}!${String(operation.args.address)}`, operation.args.values);
      }
      return { status: "committed" };
    },
    async reconcileFailedGroup(group) {
      return group.operations.map(() => "unknown_within_batch" as const);
    },
    async verifyPlan() {
      events.push("verify-read");
      await hooks.duringVerify?.();
      if (verify.throws !== undefined) throw new Error(verify.throws);
      if (verify.result !== undefined) return verify.result;
      return JSON.stringify(cells.get("Sheet1!A1")) === JSON.stringify([["ok"]]);
    },
  };
  return { adapter, cells };
}

interface HarnessOptions {
  ackMode?: AckMode;
  resultMode?: ResultMode;
  rolloutEnabled?: boolean;
  hooks?: (harness: Harness) => AdapterHooks;
  verify?: VerifyBehaviour;
  routeOverrides?: Partial<Parameters<RouteModule["createServerOwnedPlanRoute"]>[0]>;
}

interface Harness {
  readonly events: string[];
  readonly companion: CompanionDouble;
  readonly bridge: officeBridge.OfficeCompanionBridge;
  readonly coordinator: TaskPaneTurnCoordinator;
  readonly route: ReturnType<RouteModule["createServerOwnedPlanRoute"]>;
  readonly cells: Map<string, unknown>;
  readonly controls: BridgePlanExecutionControl[];
  readonly intervals: (() => void)[];
  readonly primitiveCalls: string[];
}

async function harness(options: HarnessOptions = {}): Promise<Harness> {
  const mod = await loadRouteModule();
  expect(mod, "src/taskpane/serverOwnedPlanRoute.ts must export the extracted route").toBeDefined();
  expect(typeof mod?.createServerOwnedPlanRoute).toBe("function");
  const events: string[] = [];
  const companion = new CompanionDouble(events, options.ackMode ?? "ok", options.resultMode ?? "r2d2");
  const controls: BridgePlanExecutionControl[] = [];
  const intervals: (() => void)[] = [];
  const primitiveCalls: string[] = [];
  const bindings = new ServerOwnedTurnBindingStore({ now: () => Date.now(), ttlMs: 90_000, maxEntries: 32 });
  let current: Harness | undefined;
  let route!: ReturnType<RouteModule["createServerOwnedPlanRoute"]>;
  const bridge = createExcelCompanionBridge(
    undefined,
    (name, argumentsValue, trusted, control) => {
      if (control !== undefined) controls.push(control);
      return route.executeCompanionTool(name, argumentsValue, trusted, control);
    },
    {
      fetchImpl: companion.fetch as typeof fetch,
      documentId: async () => DOC,
      paneId: PANE,
      setIntervalImpl: ((handler: () => void) => { intervals.push(handler); return intervals.length; }) as unknown as typeof setInterval,
      clearIntervalImpl: (() => undefined) as unknown as typeof clearInterval,
    },
  );
  const coordinator = new TaskPaneTurnCoordinator({
    getOrCreateDocumentId: async () => DOC,
    getBridgeIdentity: () => ({ sessionId: bridge.sessionId, documentId: bridge.documentId }),
    reconnectForDocumentChange: async () => { throw new Error("unexpected document change"); },
    bindings,
    paneId: PANE,
  });
  await bridge.connect();
  const controller = new AbortController();
  const trustedRequest = classifyWorkbookRequest("write ok into A1");
  const turnStart = Date.now();
  const handle = coordinator.beginTurn(TURN, controller, { trustedRequest, turnStart, cancellationSignal: controller.signal });
  await expect(coordinator.prepareTurn(handle)).resolves.toMatchObject({ status: "active" });
  const activeSessionId = bridge.sessionId;
  const { adapter, cells } = statefulAdapter(events, {
    beforeGate: () => current !== undefined ? options.hooks?.(current).beforeGate?.() : undefined,
    duringMutation: () => current !== undefined ? options.hooks?.(current).duringMutation?.() : undefined,
    duringVerify: () => current !== undefined ? options.hooks?.(current).duringVerify?.() : undefined,
  }, options.verify);
  const store = new BrowserIdempotencyStore(new MemoryIdempotencyStorage());
  const claim = store.claim.bind(store);
  store.claim = async (...args: Parameters<typeof store.claim>) => {
    events.push("store-claim");
    return claim(...args);
  };
  route = mod!.createServerOwnedPlanRoute({
    isRolloutEnabled: () => options.rolloutEnabled ?? true,
    getBridgeSessionId: () => bridge.sessionId,
    getActiveTurn: () => ({
      turnOperationId: TURN,
      trustedRequest,
      sessionId: activeSessionId,
      turnStart,
      hasCancellation: true,
    }),
    coordinator,
    bindings,
    readRuntimeContext: async () => { events.push("runtime-read"); return RUNTIME; },
    createAdapter: () => adapter,
    createStore: () => store,
    inspect: async () => { events.push("inspect-read"); return "{}"; },
    awaitPaneConfirmation: (binding, signal) => coordinator.waitForWorkbookPlanConfirmation(binding, signal),
    executePrimitive: async (name) => {
      primitiveCalls.push(name);
      events.push(`primitive:${name}`);
      return JSON.stringify({ ok: true, data: [[1]] });
    },
    ...options.routeOverrides,
  });
  current = { events, companion, bridge, coordinator, route, cells, controls, intervals, primitiveCalls };
  return current;
}

const WORKBOOK_TOUCHES = ["runtime-read", "inspect-read", "store-claim", "fresh", "mutation", "verify-read"];

function expectNoWorkbookAccess(h: Harness): void {
  for (const touch of WORKBOOK_TOUCHES) expect(h.events).not.toContain(touch);
  expect(h.primitiveCalls).toEqual([]);
  expect(h.cells.size).toBe(0);
}

/** Wait for real async work (including WebCrypto) without advancing the simulated turn clock. */
async function flushUntil(condition: () => boolean): Promise<void> {
  const { setTimeout: realDelay } = await import("node:timers/promises");
  for (let attempt = 0; attempt < 2000 && !condition(); attempt += 1) {
    await realDelay(1);
  }
  expect(condition()).toBe(true);
}

function onlyResult(h: Harness): Record<string, unknown> {
  expect(h.companion.results).toHaveLength(1);
  return h.companion.results[0];
}

const WRITE_CANARY = "WRITE_PATH_CANARY_LOWER_LAYER_3C8";
const NOT_EXECUTED_REASONS = ["ask", "refuse", "rejected", "cancelled", "not_started", "invalid"];

/**
 * After a resolved acknowledgement the pane payload is either the executed
 * report (committed / verification_failed / partial with its execution
 * report) or the fixed uncertain shape. It never carries a non-executing
 * reason, never says rejected/cancelled, and is never ok unless committed.
 */
function expectAcknowledgedWritePayload(payload: Record<string, unknown>): "uncertain" | "executed" {
  expect(payload).not.toHaveProperty("reason");
  expect(payload.action).toBe("workbook_plan");
  if (payload.state === "uncertain") {
    expect(payload).toEqual({
      ok: false,
      action: "workbook_plan",
      state: "uncertain",
      error: "Workbook plan execution ended without a confirmed result.",
    });
    return "uncertain";
  }
  const inner = payload.result as { outcome?: string; execution?: unknown } | undefined;
  expect(inner, "acknowledged write must carry the executed report or the uncertain shape").toBeDefined();
  expect(["committed", "verification_failed", "partial"]).toContain(inner?.outcome);
  expect(inner?.execution, "the execution report must survive").toBeDefined();
  expect(payload.ok).toBe(inner?.outcome === "committed");
  return "executed";
}

afterEach(() => {
  vi.useRealTimers();
});

describe("R2d-1 production plan route (bridge -> route -> orchestrator -> executor)", () => {
  it("routes a trusted excel_execute_plan ticket to the orchestrator and acknowledges once, ordered claim -> fresh -> ack -> first mutation", async () => {
    const h = await harness();
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    expect(h.companion.executeStartRequests).toBe(1);
    expect(h.events.filter((event) => event === "execute-start")).toHaveLength(1);
    expect(h.events.filter((event) => event === "mutation")).toHaveLength(1);
    const order = ["companion-claim", "runtime-read", "store-claim", "fresh", "execute-start", "mutation", "verify-read", "result"];
    const positions = order.map((event) => h.events.indexOf(event));
    expect(positions.every((position) => position >= 0)).toBe(true);
    expect([...positions].sort((a, b) => a - b)).toEqual(positions);
    expect(h.cells.get("Sheet1!A1")).toEqual([["ok"]]);

    const posted = onlyResult(h);
    expect(Object.keys(posted).sort()).toEqual(["executed", "result", "ticketId"]);
    expect(posted.executed).toBe(true);
    expect(posted.result).toMatchObject({ ok: true, action: "workbook_plan", result: { outcome: "committed", verification: "verified" } });
    expect(posted.result).not.toHaveProperty("reason");
    expect(expectAcknowledgedWritePayload(posted.result as Record<string, unknown>)).toBe("executed");
    expect(h.companion.ticket).toMatchObject({ state: "terminal", terminalReason: "completed" });
  });

  it("rejects the legacy execute_plan name before any read and never reaches the primitive validator", async () => {
    const h = await harness();
    h.companion.enqueue("execute_plan", { ...EXECUTE_PLAN });

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    expectNoWorkbookAccess(h);
    expect(h.companion.executeStartRequests).toBe(0);
    const posted = onlyResult(h);
    expect(Object.keys(posted).sort()).toEqual(["result", "ticketId"]);
    expect(posted.result).toMatchObject({ ok: false });

    await expect(h.route.executeCompanionTool("execute_plan", { ...EXECUTE_PLAN }, TRUSTED)).rejects.toThrow();
    await expect(h.route.executeCompanionTool("excel_execute_plan", { ...EXECUTE_PLAN })).rejects.toThrow();
    await expect(h.route.executeCompanionTool("read_range", { sheet: "Sheet1", address: "A1" }, TRUSTED)).rejects.toThrow();
    expectNoWorkbookAccess(h);
  });

  it("default-off performs zero reads, writes and execute-starts and posts executed:false", async () => {
    const h = await harness({ rolloutEnabled: false });
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    expectNoWorkbookAccess(h);
    expect(h.companion.executeStartRequests).toBe(0);
    const posted = onlyResult(h);
    expect(posted.executed).toBe(false);
    expect(posted.result).toMatchObject({ ok: false, action: "workbook_plan", reason: "not_started" });
    expect(h.companion.ticket).toMatchObject({ state: "terminal", terminalReason: "not_executed" });
  });

  it.each([
    ["ask", { version: 1, outcome: "ask", message: "Which sheet?" }],
    ["refuse", { version: 1, outcome: "refuse", message: "Not permitted." }],
    ["rejected", { ...RESOURCE_REJECTED_PLAN }],
    ["rejected", { version: 1, outcome: "inspect", reason: "look", reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1" }] }],
  ] as const)("%s outcome sends no execute-start and posts executed:false with a bounded reason", async (reason, plan) => {
    const h = await harness();
    h.companion.enqueue("excel_execute_plan", { ...plan } as Record<string, unknown>, plan.outcome === "inspect" ? TRUSTED_V2_INSPECT : TRUSTED_V2);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    expect(h.companion.executeStartRequests).toBe(0);
    expect(h.events).not.toContain("mutation");
    expect(h.cells.size).toBe(0);
    const posted = onlyResult(h);
    expect(posted.executed).toBe(false);
    expect(posted.result).toMatchObject({ ok: false, action: "workbook_plan", reason });
    const payload = posted.result as Record<string, unknown>;
    expect(typeof payload.summary).toBe("string");
    if (reason === "ask" || reason === "refuse") {
      // The plan's own ask/refuse text reaches the model (built from the validated plan).
      expect(payload.result).toEqual({ outcome: reason, finalSummary: (plan as { message: string }).message });
    } else {
      // Every other non-executing outcome forwards only the fixed summary.
      expect(payload).not.toHaveProperty("result");
      expect(JSON.stringify(payload)).not.toContain("finalSummary");
    }
    expect(h.companion.ticket).toMatchObject({ state: "terminal", terminalReason: "not_executed" });
  });

  it.each([
    ["stale schema version", { ...EXECUTE_PLAN, version: 2 }, TRUSTED_V2],
    ["missing envelope", { ...EXECUTE_PLAN }, undefined],
    ["foreign fingerprint", { ...EXECUTE_PLAN }, { ...TRUSTED_V2, workbookPlanFingerprint: "0".repeat(64) }],
  ] as const)("invalid trusted ticket (%s) posts executed:false/invalid without touching the route", async (_label, plan, trusted) => {
    const h = await harness();
    h.companion.enqueue("excel_execute_plan", { ...plan } as Record<string, unknown>, trusted);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    expectNoWorkbookAccess(h);
    expect(h.controls).toEqual([]);
    expect(h.companion.executeStartRequests).toBe(0);
    const posted = onlyResult(h);
    expect(posted.executed).toBe(false);
    expect(posted.result).toMatchObject({ ok: false, reason: "invalid" });
    expect(h.companion.ticket).toMatchObject({ state: "terminal", terminalReason: "not_executed" });
  });

  it.each([409, 404, 500, "network"] as const)(
    "execute-start %s gives zero mutation, executed:false and a content-free result",
    async (ackMode) => {
      const h = await harness({ ackMode });
      h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);

      await expect(h.bridge.pollOnce()).resolves.toBe(true);

      expect(h.companion.executeStartRequests).toBe(1);
      expect(h.events).toContain("fresh");
      expect(h.events).not.toContain("mutation");
      expect(h.cells.size).toBe(0);
      const posted = onlyResult(h);
      expect(posted.executed).toBe(false);
      expect(posted.result).toMatchObject({ ok: false, action: "workbook_plan", reason: "not_started" });
      expect(JSON.stringify(posted)).not.toContain(COMPANION_SECRET);
      expect(h.companion.ticket).toMatchObject({ state: "terminal", terminalReason: "not_executed" });
    },
  );

  it("a hanging execute-start is aborted at the fixed timeout: zero mutation, executed:false, the turn finishes within bound", async () => {
    const timeout = (officeBridge as { EXECUTE_START_ACK_TIMEOUT_MS?: number }).EXECUTE_START_ACK_TIMEOUT_MS;
    expect(typeof timeout).toBe("number");
    expect(timeout!).toBeGreaterThan(0);
    expect(timeout!).toBeLessThanOrEqual(10_000);
    const h = await harness({ ackMode: "hang" });
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });

    let settled = false;
    const polled = h.bridge.pollOnce().finally(() => { settled = true; });
    await flushUntil(() => h.companion.executeStartRequests === 1);
    await vi.advanceTimersByTimeAsync(timeout! - 1);
    expect(settled).toBe(false);
    await vi.advanceTimersByTimeAsync(1);
    await expect(polled).resolves.toBe(true);

    expect(h.companion.ackSignals[0]?.aborted).toBe(true);
    expect(h.events).not.toContain("mutation");
    expect(h.cells.size).toBe(0);
    const posted = onlyResult(h);
    expect(posted.executed).toBe(false);
    expect(posted.result).toMatchObject({ ok: false, reason: "not_started" });
  });

  it("Stop before the acknowledgement gives zero execute-start and zero mutation", async () => {
    const h = await harness({ hooks: (hh) => ({ beforeGate: () => hh.coordinator.stopCurrentTurn() }) });
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    expect(h.companion.executeStartRequests).toBe(0);
    expect(h.events).not.toContain("mutation");
    expect(h.cells.size).toBe(0);
    const posted = onlyResult(h);
    expect(posted.executed).toBe(false);
    expect(posted.result).toMatchObject({ ok: false, reason: "cancelled" });
  });

  it("Stop during the acknowledgement aborts the fetch and gives zero mutation", async () => {
    const h = await harness({ ackMode: "hang" });
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);

    const polled = h.bridge.pollOnce();
    await vi.waitFor(() => expect(h.companion.executeStartRequests).toBe(1));
    h.coordinator.stopCurrentTurn();
    await expect(polled).resolves.toBe(true);

    expect(h.companion.ackSignals[0]?.aborted).toBe(true);
    expect(h.events).not.toContain("mutation");
    expect(h.cells.size).toBe(0);
    const posted = onlyResult(h);
    expect(posted.executed).toBe(false);
    expect(posted.result).toMatchObject({ ok: false, reason: "not_started" });
  });

  it("Stop during Office sync is reported partial and never claimed as undone", async () => {
    const h = await harness({ hooks: (hh) => ({ duringMutation: () => hh.coordinator.stopCurrentTurn() }) });
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    expect(h.companion.executeStartRequests).toBe(1);
    expect(h.events.filter((event) => event === "mutation")).toHaveLength(1);
    const posted = onlyResult(h);
    expect(posted.executed).toBe(true);
    expect(posted.result).toMatchObject({ ok: false, action: "workbook_plan", result: { outcome: "partial" } });
    expect(posted.result).not.toHaveProperty("reason");
    const summary = String((posted.result as { result: { finalSummary: string } }).result.finalSummary);
    expect(summary).toContain("not claimed to be undone");
    expect(summary.toLowerCase()).not.toContain("rolled back");
    expect(h.companion.ticket).toMatchObject({ state: "terminal", terminalReason: "completed" });
  });

  it("refuses a second acknowledgement locally with no network request", async () => {
    // The second call is made while the turn is still live (during the
    // mutation, before executeTool returns), so only the one-shot guard can
    // refuse it; a later call is also refused once the control is closed.
    let secondWhileLive: string | undefined;
    const h = await harness({
      hooks: (hh) => ({
        duringMutation: async () => {
          expect(hh.controls).toHaveLength(1);
          secondWhileLive = await hh.controls[0].acknowledgeExecuteStart().then(() => "resolved", () => "rejected");
        },
      }),
    });
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);
    await expect(h.bridge.pollOnce()).resolves.toBe(true);
    expect(h.controls).toHaveLength(1);
    expect(secondWhileLive).toBe("rejected");
    expect(h.companion.executeStartRequests).toBe(1);
    expect(onlyResult(h).executed).toBe(true);

    await expect(h.controls[0].acknowledgeExecuteStart()).rejects.toThrow();
    expect(h.companion.executeStartRequests).toBe(1);
  });

  it("a 401 auto-reconnect between poll and acknowledgement rejects the acknowledgement: zero mutation", async () => {
    // The reconnect lands while the execute-start request is in flight and the
    // companion answers 200 afterwards. No coordinator or assertLive check runs
    // between that 200 and the first mutation, so only the bridge's binding to
    // the captured PrivateSession can refuse it.
    let releaseAck!: () => void;
    const h = await harness();
    h.companion.ackGate = new Promise<void>((resolve) => { releaseAck = resolve; });
    await h.bridge.start();
    expect(h.intervals).toHaveLength(2);
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);
    const before = h.bridge.sessionId;

    const polled = h.bridge.pollOnce().catch(() => false);
    await flushUntil(() => h.companion.executeStartRequests === 1);
    expect(h.companion.ticket?.state).toBe("executing");
    h.companion.heartbeatUnauthorized = true;
    h.intervals[0]();
    await flushUntil(() => h.bridge.sessionId !== null && h.bridge.sessionId !== before);
    h.companion.heartbeatUnauthorized = false;
    releaseAck();
    await polled;

    expect(h.companion.executeStartRequests).toBe(1);
    expect(h.events).not.toContain("mutation");
    expect(h.cells.size).toBe(0);
    for (const posted of h.companion.results) expect(posted.executed).toBe(false);
    // The stale ticket is not re-executed under the new session.
    await h.bridge.pollOnce().catch(() => false);
    expect(h.events).not.toContain("mutation");
    expect(h.companion.executeStartRequests).toBe(1);
  });

  it("mixed version (realistic pre-R2d-2 companion that ignores `executed`): accepted plan acknowledges only via the gate, posts executed:true, never retries", async () => {
    const h = await harness({ resultMode: "legacy_ignores_executed" });
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);
    await expect(h.bridge.pollOnce()).resolves.toBe(false);

    expect(h.companion.executeStartRequests).toBe(1);
    expect(h.events.indexOf("fresh")).toBeLessThan(h.events.indexOf("execute-start"));
    expect(h.events.indexOf("execute-start")).toBeLessThan(h.events.indexOf("mutation"));
    expect(h.events.filter((event) => event === "mutation")).toHaveLength(1);
    expect(h.events.filter((event) => event === "companion-claim")).toHaveLength(1);
    const posted = onlyResult(h);
    expect(posted.executed).toBe(true);
    expect(posted.result).toMatchObject({ ok: true, result: { outcome: "committed" } });
  });

  it.each([
    ["ask", { version: 1, outcome: "ask", message: "Which sheet?" }],
    ["not_started", { ...EXECUTE_PLAN }],
  ] as const)("mixed version (realistic pre-R2d-2 companion): a non-executing %s outcome posts executed:false and never reports success in the pane payload", async (reason, plan) => {
    const h = await harness({ resultMode: "legacy_ignores_executed", rolloutEnabled: reason !== "not_started" });
    h.companion.enqueue("excel_execute_plan", { ...plan } as Record<string, unknown>, TRUSTED_V2);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);
    await expect(h.bridge.pollOnce()).resolves.toBe(false);

    expect(h.companion.executeStartRequests).toBe(0);
    expect(h.events).not.toContain("mutation");
    expect(h.cells.size).toBe(0);
    const posted = onlyResult(h);
    expect(posted.executed).toBe(false);
    expect(posted.result).toMatchObject({ ok: false, action: "workbook_plan", reason });
    // Known deployment coupling (recorded in case-state): this old companion
    // still marks the ticket "completed" for the plugin. Not fixable in the pane.
    expect(h.companion.ticket).toMatchObject({ state: "terminal", terminalReason: "completed" });
  });

  it("mixed version (extra strict double): an old companion that rejects `executed` never yields success and the mutation is not retried", async () => {
    const h = await harness({ resultMode: "legacy_rejects_executed" });
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);

    await expect(h.bridge.pollOnce()).rejects.toThrow();
    await expect(h.bridge.pollOnce()).rejects.toThrow();

    expect(h.companion.executeStartRequests).toBe(1);
    expect(h.events.filter((event) => event === "mutation")).toHaveLength(1);
    expect(h.events.filter((event) => event === "companion-claim")).toHaveLength(1);
    expect(h.companion.results).toHaveLength(2);
    expect(h.companion.results[1]).toEqual(h.companion.results[0]);
    expect(h.companion.ticket?.state).not.toBe("terminal");
  });

  it("a verification read that throws after a committed write reports uncertain with executed:true, never rejected and never the thrown text", async () => {
    const h = await harness({ verify: { throws: WRITE_CANARY } });
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    expect(h.companion.executeStartRequests).toBe(1);
    expect(h.events.filter((event) => event === "mutation")).toHaveLength(1);
    expect(h.cells.get("Sheet1!A1")).toEqual([["ok"]]);
    const posted = onlyResult(h);
    expect(posted.executed).toBe(true);
    expect(expectAcknowledgedWritePayload(posted.result as Record<string, unknown>)).toBe("uncertain");
    expect(JSON.stringify(posted)).not.toContain(WRITE_CANARY);
    expect(h.companion.ticket).toMatchObject({ state: "terminal", terminalReason: "completed" });
  });

  it("Stop during the post-mutation verification read is never reported as cancelled", async () => {
    const h = await harness({ hooks: (hh) => ({ duringVerify: () => hh.coordinator.stopCurrentTurn() }) });
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    expect(h.companion.executeStartRequests).toBe(1);
    expect(h.events.filter((event) => event === "mutation")).toHaveLength(1);
    expect(h.cells.get("Sheet1!A1")).toEqual([["ok"]]);
    const posted = onlyResult(h);
    expect(posted.executed).toBe(true);
    expectAcknowledgedWritePayload(posted.result as Record<string, unknown>);
    expect(JSON.stringify(posted.result)).not.toMatch(/"outcome":"(cancelled|rejected)"/);
  });

  it("the 90 s turn deadline expiring during the post-mutation verification read is never reported as cancelled", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    const h = await harness({
      hooks: () => ({ duringVerify: () => { vi.setSystemTime(Date.now() + WORKBOOK_PLAN_TURN_TIMEOUT_MS + 1); } }),
    });
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    expect(h.companion.executeStartRequests).toBe(1);
    expect(h.events.filter((event) => event === "mutation")).toHaveLength(1);
    expect(h.events).toContain("verify-read");
    expect(h.cells.get("Sheet1!A1")).toEqual([["ok"]]);
    const posted = onlyResult(h);
    expect(posted.executed).toBe(true);
    expectAcknowledgedWritePayload(posted.result as Record<string, unknown>);
    expect(JSON.stringify(posted.result)).not.toMatch(/"outcome":"(cancelled|rejected)"/);
  });

  it("verification_failed after an acknowledged write keeps the executed report unchanged", async () => {
    const h = await harness({ verify: { result: false } });
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    const posted = onlyResult(h);
    expect(posted.executed).toBe(true);
    expect(expectAcknowledgedWritePayload(posted.result as Record<string, unknown>)).toBe("executed");
    expect(posted.result).toMatchObject({
      ok: false,
      action: "workbook_plan",
      result: { outcome: "verification_failed", verification: "failed", execution: { outcome: "committed" } },
    });
  });

  it.each([
    ["readRuntimeContext", { readRuntimeContext: async () => { throw new Error(WRITE_CANARY); } }],
    ["store.claim", {
      createStore: () => {
        const store = new BrowserIdempotencyStore(new MemoryIdempotencyStorage());
        store.claim = async () => { throw new Error(WRITE_CANARY); };
        return store;
      },
    }],
  ] as const)("a non-executing rejection caused by a %s throw forwards no lower-layer text", async (_stage, routeOverrides) => {
    const h = await harness({ routeOverrides });
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, TRUSTED_V2);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    expect(h.companion.executeStartRequests).toBe(0);
    expect(h.events).not.toContain("mutation");
    const posted = onlyResult(h);
    expect(posted.executed).toBe(false);
    const payload = posted.result as Record<string, unknown>;
    expect(payload).toMatchObject({ ok: false, action: "workbook_plan", reason: "rejected" });
    expect(NOT_EXECUTED_REASONS).toContain(payload.reason);
    expect(JSON.stringify(posted), "model-visible payload leaks lower-layer text").not.toContain(WRITE_CANARY);
    expect(typeof payload.summary).toBe("string");
    expect(payload).not.toHaveProperty("result");
  });

  it("keeps the primitive bridge round trip byte-for-byte (no executed field, no control)", async () => {
    const h = await harness();
    h.companion.enqueue("read_range", { sheet: "Sheet1", address: "A1" });

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    expect(h.primitiveCalls).toEqual(["read_range"]);
    expect(h.controls).toEqual([]);
    expect(h.companion.executeStartRequests).toBe(0);
    const posted = onlyResult(h);
    expect(Object.keys(posted)).toEqual(["ticketId", "result"]);
    expect(posted).toEqual({ ticketId: TICKET, result: { ok: true, data: [[1]] } });
  });

  it("keeps workbook and companion text out of non-executing route errors", async () => {
    const h = await harness();
    const thrown = await h.route.executeCompanionTool("execute_plan", { secret: WORKBOOK_SECRET }).then(
      () => "resolved",
      (error: unknown) => (error instanceof Error ? error.message : String(error)),
    );
    expect(thrown).not.toContain(WORKBOOK_SECRET);
    expect(thrown).not.toBe("resolved");
  });
});

/**
 * R3c G5 (D6): the server-route executor boundary is (bridgeSessionId, documentId, serverOperationId).
 * Real chain with a test-owned durable store per harness; the persisted record exposes the boundary token
 * and the deterministic transaction id derived from it.
 */
describe("R3c server-route idempotency key binds the server operation (G5)", () => {
  const OTHER_OP = "5e7e0000-0000-4000-8000-000000000013";
  const IDEMPOTENCY_KEY = "hermes.excel.workbookPlan.idempotency.v1";

  async function committedRecord(trusted: BridgeTrustedPlanEnvelopeV2): Promise<{
    readonly sid: string;
    readonly record: { readonly boundary: string; readonly transactionId: string; readonly state: string };
  }> {
    const storage = new MemoryIdempotencyStorage();
    const h = await harness({ routeOverrides: { createStore: () => new BrowserIdempotencyStore(storage) } });
    const sid = h.bridge.sessionId;
    expect(typeof sid).toBe("string");
    h.companion.enqueue("excel_execute_plan", { ...EXECUTE_PLAN }, trusted);

    await expect(h.bridge.pollOnce()).resolves.toBe(true);

    expect(h.bridge.sessionId, "a clean committed turn must not rotate the bridge session (R3b)").toBe(sid);
    expect(onlyResult(h).result).toMatchObject({ ok: true, action: "workbook_plan", result: { outcome: "committed" } });
    const records = JSON.parse(storage.getItem(IDEMPOTENCY_KEY) ?? "[]") as { boundary: string; transactionId: string; state: string }[];
    expect(records).toHaveLength(1);
    return { sid: sid as string, record: records[0] };
  }

  it("RC10 persists the record under ${bridgeSessionId}.${documentId}.${serverOperationId}", async () => {
    const { sid, record } = await committedRecord(TRUSTED_V2);
    expect(record.state).toBe("terminal");
    expect(record.boundary).toBe(`${sid}.${DOC}.${SERVER_OP}`);
    expect(record.boundary).not.toContain(TURN);
  });

  it("RC11 distinct server operations in the same pane turn get distinct transaction ids", async () => {
    const first = await committedRecord(TRUSTED_V2);
    const second = await committedRecord({ ...TRUSTED_V2, turnOperationId: OTHER_OP });
    expect(second.sid).toBe(first.sid);
    expect(second.record.boundary).not.toBe(first.record.boundary);
    expect(second.record.transactionId).not.toBe(first.record.transactionId);
  });

  it("GC12 the same server operation derives the same transaction id", async () => {
    const first = await committedRecord(TRUSTED_V2);
    const second = await committedRecord(TRUSTED_V2);
    expect(second.sid).toBe(first.sid);
    expect(second.record.boundary).toBe(first.record.boundary);
    expect(second.record.transactionId).toBe(first.record.transactionId);
  });
});
