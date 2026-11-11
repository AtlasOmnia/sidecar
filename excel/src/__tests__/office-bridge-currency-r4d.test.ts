/**
 * R4d bridge-unit tests (dev3-r4d-release.md §5.1 B1-B5, §5.2 SP4).
 *
 * B1-B3 (R3b lane-B A1): a ticket whose `GET /pending` body is read after the
 * captured session/generation stopped being current (Stop, 401 auto-reconnect)
 * is dropped: pollOnce resolves false with no executeTool call and no
 * `POST /result` (plan, primitive and invalid-plan tickets alike).
 * B4-B5 (R3b P10): a primitive executeTool throw posts a fixed error body.
 * SP4: a primitive ticket with a trusted envelope posts the fixed primitive error (r2d-1 A3, R4g-2a).
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import { OfficeCompanionBridge } from "../officeBridge";
import { createServerOwnedPlanRoute } from "../taskpane/serverOwnedPlanRoute";

const PANE_ID = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
const DOC = "d0c00000-0000-4000-8000-0000000000d4";
const PLAN_OP = "5e7e0000-0000-4000-8000-0000000000d4";
const FIXED_PRIMITIVE_ERROR = { ok: false, category: "Execution", error: "Office tool execution failed" };
const CANARY = "CANARY-r4d-9c1";

function json(status: number, body?: unknown): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
  });
}

function uuid(prefix: string, n: number): string {
  return `${prefix}-0000-4000-8000-${String(n).padStart(12, "0")}`;
}

interface Ticket {
  readonly ticketId: string;
  readonly requestId: string;
  readonly tool: string;
  readonly args: unknown;
  readonly trusted?: unknown;
  claimed: boolean;
}

/** Minimal session-scoped companion double with a hold-able GET /pending. */
class Companion {
  readonly sessions = new Map<string, { sessionId: string; sessionKey: string; tickets: Ticket[] }>();
  readonly requests: { route: string; method: string; body: string }[] = [];
  readonly resultPosts: { sessionId: string | undefined; body: string }[] = [];
  readonly deleted: string[] = [];
  heartbeat401 = false;
  heldPending = 0;
  private hold: Promise<void> | null = null;
  private counter = 0;

  holdNextPending(): () => void {
    let release!: () => void;
    this.hold = new Promise<void>((resolve) => { release = resolve; });
    return () => release();
  }

  enqueue(sessionId: string, tool: string, args: unknown, trusted?: unknown): string {
    const session = [...this.sessions.values()].find((candidate) => candidate.sessionId === sessionId);
    if (session === undefined) throw new Error("unknown session");
    this.counter += 1;
    const ticket: Ticket = {
      ticketId: uuid("71c70000", this.counter),
      requestId: uuid("4e4e0000", this.counter),
      tool,
      args,
      ...(trusted === undefined ? {} : { trusted }),
      claimed: false,
    };
    session.tickets.push(ticket);
    return ticket.ticketId;
  }

  readonly fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const route = String(input).replace(/^.*\/excel\//, "");
    const method = init?.method ?? "GET";
    const headers = (init?.headers ?? {}) as Record<string, string>;
    const bodyText = typeof init?.body === "string" ? init.body : "";
    this.requests.push({ route, method, body: bodyText });
    if (method === "POST" && route === "session") {
      this.counter += 1;
      const sessionId = uuid("aaaaaaaa", this.counter).replace(/^aaaaaaaa-0000-4000-8000/, "aaaaaaaa-aaaa-4aaa-8aaa");
      const sessionKey = `${"k".repeat(40)}${this.counter}`;
      this.sessions.set(sessionKey, { sessionId, sessionKey, tickets: [] });
      return json(201, {
        ok: true, sessionId, sessionKey, protocolVersion: "1.0", expiresInSeconds: 60,
        workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
      });
    }
    if (method === "POST" && route === "result") {
      this.resultPosts.push({ sessionId: headers["X-Hermes-Session-Id"], body: bodyText });
    }
    const session = this.sessions.get(headers["X-Hermes-Session-Key"] ?? "");
    if (session === undefined || headers["X-Hermes-Session-Id"] !== session.sessionId) return json(401, { error: "Unauthorized" });
    if (method === "DELETE" && route === "session") {
      this.sessions.delete(session.sessionKey);
      this.deleted.push(session.sessionId);
      return json(200, { ok: true });
    }
    if (method === "POST" && route === "heartbeat") {
      return this.heartbeat401 ? json(401, { error: "Unauthorized" }) : json(200, { ok: true });
    }
    if (method === "GET" && route === "pending") {
      const ticket = session.tickets.find((candidate) => !candidate.claimed);
      if (ticket === undefined) return json(204);
      ticket.claimed = true;
      const hold = this.hold;
      this.hold = null;
      if (hold !== null) {
        this.heldPending += 1;
        await hold;
      }
      return json(200, {
        ticketId: ticket.ticketId, requestId: ticket.requestId, tool: ticket.tool, arguments: ticket.args,
        ...(ticket.trusted === undefined ? {} : { trusted: ticket.trusted }),
      });
    }
    if (method === "POST" && (route === "result" || route === "execute-start")) return json(200, { ok: true });
    return json(404, { error: "route" });
  };
}

interface Harness {
  readonly companion: Companion;
  readonly bridge: OfficeCompanionBridge;
  readonly intervals: (() => void)[];
  readonly executeTool: ReturnType<typeof vi.fn>;
}

const bridges: OfficeCompanionBridge[] = [];
afterEach(async () => {
  while (bridges.length > 0) await bridges.pop()!.stop().catch(() => undefined);
});

async function harness(executeImpl?: (...args: unknown[]) => Promise<string>): Promise<Harness> {
  const companion = new Companion();
  const intervals: (() => void)[] = [];
  const executeTool = vi.fn(executeImpl ?? (async () => JSON.stringify({ ok: true, data: [[1]] })));
  const bridge = new OfficeCompanionBridge({
    hostPrefix: "excel",
    documentId: async () => DOC,
    paneId: PANE_ID,
    executeTool: executeTool as unknown as ConstructorParameters<typeof OfficeCompanionBridge>[0]["executeTool"],
    fetchImpl: companion.fetch as typeof fetch,
    setIntervalImpl: ((handler: () => void) => { intervals.push(handler); return intervals.length; }) as unknown as typeof setInterval,
    clearIntervalImpl: (() => undefined) as unknown as typeof clearInterval,
  });
  bridges.push(bridge);
  await bridge.start();
  return { companion, bridge, intervals, executeTool };
}

async function flushUntil(condition: () => boolean): Promise<void> {
  for (let attempt = 0; attempt < 400 && !condition(); attempt += 1) {
    await new Promise<void>((done) => setImmediate(done));
  }
  expect(condition()).toBe(true);
}

const R4G2_TURN_KEY = "4b4b0000-0000-4000-8000-0000000000a3"; // R4g-2a rule M: turnKey of migrated v2 tickets
const TRUSTED = { turnOperationId: PLAN_OP, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT, phase: "execute", turnKey: R4G2_TURN_KEY };
const PLAN_ARGS = { version: 1, outcome: "ask", message: "bounded" };
const TICKETS = {
  plan: ["excel_execute_plan", PLAN_ARGS, TRUSTED],
  primitive: ["write_range", { sheet: "Sheet1", address: "A1", values: [[1]] }, undefined],
  invalidPlan: ["excel_execute_plan", [1, 2], TRUSTED],
} as const;

/** Holds the GET /pending of a ticket claimed on S1, runs `rotate`, then releases the body. */
async function staleTicket(h: Harness, kind: keyof typeof TICKETS, rotate: () => Promise<void>) {
  const s1 = h.bridge.sessionId!;
  const [tool, args, trusted] = TICKETS[kind];
  h.companion.enqueue(s1, tool, args, trusted);
  const release = h.companion.holdNextPending();
  const polled = h.bridge.pollOnce().catch((error: unknown) => `rejected:${String(error)}`);
  await flushUntil(() => h.companion.heldPending === 1);
  await rotate();
  release();
  return { s1, polled: await polled };
}

describe("B-rows: a ticket read on a stale session is dropped before execution (R3b lane-B A1)", () => {
  it.each(["plan", "primitive"] as const)("B1 Stop while the %s ticket body is held: false, no executeTool, no POST /result; S2 posts nothing for it", async (kind) => {
    const h = await harness();
    const { s1, polled } = await staleTicket(h, kind, () => h.bridge.stop());
    expect(polled).toBe(false);
    expect(h.executeTool).not.toHaveBeenCalled();
    expect(h.companion.resultPosts).toEqual([]);
    expect(h.companion.requests.filter((request) => request.route === "execute-start")).toEqual([]);
    await h.bridge.start();
    expect(h.bridge.sessionId).not.toBe(s1);
    expect(await h.bridge.pollOnce()).toBe(false);
    expect(h.executeTool).not.toHaveBeenCalled();
    expect(h.companion.resultPosts).toEqual([]);
  });

  it.each(["plan", "primitive"] as const)("B2 401 auto-reconnect while the %s ticket body is held: false, no executeTool, no POST /result", async (kind) => {
    const h = await harness();
    const before = h.bridge.sessionId;
    const { s1, polled } = await staleTicket(h, kind, async () => {
      h.companion.heartbeat401 = true;
      h.intervals[0]();
      await flushUntil(() => h.bridge.sessionId !== null && h.bridge.sessionId !== before);
    });
    expect(h.bridge.sessionId).not.toBe(s1);
    expect(polled).toBe(false);
    expect(h.executeTool).not.toHaveBeenCalled();
    expect(h.companion.resultPosts).toEqual([]);
  });

  it("B3 an invalid plan ticket (arguments array) on a stale session gets no invalid-ticket post", async () => {
    const h = await harness();
    const { polled } = await staleTicket(h, "invalidPlan", () => h.bridge.stop());
    expect(polled).toBe(false);
    expect(h.executeTool).not.toHaveBeenCalled();
    expect(h.companion.resultPosts).toEqual([]);
  });

  it("B4 a primitive executeTool throw posts the fixed error body and never the thrown text", async () => {
    const h = await harness(async () => { throw new Error(CANARY); });
    const ticketId = h.companion.enqueue(h.bridge.sessionId!, "write_range", { sheet: "Sheet1", address: "A1", values: [[1]] });
    expect(await h.bridge.pollOnce()).toBe(true);
    expect(h.executeTool).toHaveBeenCalledTimes(1);
    expect(h.companion.resultPosts.map((post) => post.body)).toEqual([JSON.stringify({ ticketId, result: FIXED_PRIMITIVE_ERROR })]);
    expect(JSON.parse(h.companion.resultPosts[0].body)).toEqual({ ticketId, result: FIXED_PRIMITIVE_ERROR });
    expect(h.companion.requests.some((request) => request.body.includes(CANARY))).toBe(false);
  });

  it("B5 the legacy execute_plan name through the real route posts the same fixed body", async () => {
    const route = createServerOwnedPlanRoute({
      isRolloutEnabled: () => true,
      getBridgeSessionId: () => null,
      getActiveTurn: () => ({ turnOperationId: null, trustedRequest: null, sessionId: null, turnStart: null, hasCancellation: false }),
      coordinator: {
        authorizeAndRun: async () => ({ ok: false, reason: "unused" }),
        consumeWorkbookPlanConfirmation: () => undefined,
        cancelPendingConfirmation: () => undefined,
      },
      bindings: { finishTerminal: () => undefined },
      readRuntimeContext: async () => { throw new Error("unused"); },
      createAdapter: () => { throw new Error("unused"); },
      createStore: () => { throw new Error("unused"); },
      inspect: async () => { throw new Error("unused"); },
      awaitPaneConfirmation: async () => { throw new Error("unused"); },
      executePrimitive: async () => { throw new Error("unused primitive"); },
    } as unknown as Parameters<typeof createServerOwnedPlanRoute>[0]);
    const h = await harness((...args: unknown[]) => route.executeCompanionTool(...(args as Parameters<typeof route.executeCompanionTool>)));
    const ticketId = h.companion.enqueue(h.bridge.sessionId!, "execute_plan", { version: 1 });
    expect(await h.bridge.pollOnce()).toBe(true);
    expect(h.companion.resultPosts.map((post) => post.body)).toEqual([JSON.stringify({ ticketId, result: FIXED_PRIMITIVE_ERROR })]);
    expect(h.companion.requests.some((request) => request.body.includes("not accepted by the pane"))).toBe(false);
  });
});

describe("SP4 trusted-envelope primitive ticket posts the fixed primitive error (r2d-1 A3, R4g-2a)", () => {
  it("SP4 a primitive ticket with a trusted envelope resolves pollOnce, executes nothing and posts the fixed primitive error once", async () => {
    const h = await harness();
    const ticketId = h.companion.enqueue(h.bridge.sessionId!, "write_range", { sheet: "Sheet1", address: "A1", values: [[1]] }, TRUSTED);
    await expect(h.bridge.pollOnce()).resolves.toBe(true);
    expect(h.executeTool).not.toHaveBeenCalled();
    expect(h.companion.resultPosts.map((post) => post.body)).toEqual([JSON.stringify({ ticketId, result: FIXED_PRIMITIVE_ERROR })]);
  });
});
