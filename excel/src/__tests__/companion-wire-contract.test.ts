/**
 * R2d-3: bridge leg of the shared plan wire contract.
 *
 * `fixtures/companion-wire/plan-wire-contract.v2.json` is the single recorded
 * wire contract. `scripts/tests/test_companion_wire_contract.py` proves its
 * companion responses against the real Python companion and plugin client.
 * This leg drives the REAL OfficeCompanionBridge with a fetchImpl that replays
 * those recorded responses, and asserts every request the bridge emits exactly
 * (method, path, header names and values, byte-exact JSON body) plus the
 * pane-side effects: execute-start count, executeTool arity, the posted
 * `executed` disposition and the pollOnce outcome.
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import { join } from "path";
import { createHash } from "crypto";
import {
  EXECUTE_START_ACK_TIMEOUT_MS,
  OfficeCompanionBridge,
  type BridgePlanExecutionControl,
  type BridgeTrustedPlanEnvelope,
} from "../officeBridge";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import { WORKBOOK_PLAN_TOOL_NAME } from "../workbookPlanProtocol";
import type { ServerOwnedPlanNotExecutedReason, ServerOwnedPlanRouteDependencies } from "../taskpane/serverOwnedPlanRoute";
import { createServerOwnedPlanRoute } from "../taskpane/serverOwnedPlanRoute";

type Json = null | boolean | number | string | Json[] | { [key: string]: Json };
type JsonObject = { [key: string]: Json };

interface WireResponse {
  readonly status: number;
  readonly body?: Json;
  readonly responseText?: string;
  readonly textRejects?: boolean;
  readonly empty?: boolean;
}

interface WireStep {
  readonly kind: string;
  readonly method?: string;
  readonly path?: string;
  readonly headers?: Record<string, string>;
  readonly body?: Json;
  readonly response?: WireResponse;
  readonly only?: "python" | "bridge";
}

interface BridgeAction {
  readonly call: "connect" | "pollOnce";
  readonly resolves?: boolean;
  readonly rejects?: { readonly name?: string; readonly message?: string };
}

interface BridgeSpec {
  readonly paneBehavior: string;
  readonly paneReport?: Json;
  readonly actions: readonly BridgeAction[];
  readonly expect: {
    readonly executeToolCalls: number;
    readonly executeToolArity: number | null;
    readonly executeStartRequests: number;
  };
  readonly routeDeps?: { readonly rolloutEnabled: boolean; readonly bridgeSessionId: string | null };
}

interface WireCase {
  readonly id: string;
  readonly tags: readonly string[];
  readonly sides: readonly string[];
  readonly companionVersion: string;
  readonly bridge?: BridgeSpec;
  readonly steps: readonly WireStep[];
}

interface WireFixture {
  readonly contractVersion: number;
  readonly protocolVersion: string;
  readonly workbookPlanFingerprint: string;
  readonly hostPrefix: string;
  readonly planTool: string;
  readonly origin: string;
  readonly constants: { readonly claimLeaseSeconds: number; readonly executeStartAckTimeoutMs: number };
  readonly messages: { readonly ackRejected: string };
  readonly requiredTags: readonly string[];
  readonly stepKinds: readonly string[];
  readonly ids: Record<string, string>;
  readonly samples: Record<string, string>;
  readonly cases: readonly WireCase[];
}

const FIXTURE_PATH = join(__dirname, "fixtures", "companion-wire", "plan-wire-contract.v2.json");
const FIXTURE_BYTES = readFileSync(FIXTURE_PATH);
const FIXTURE = JSON.parse(FIXTURE_BYTES.toString("utf8")) as WireFixture;
const FIXTURE_SHA256 = createHash("sha256").update(FIXTURE_BYTES).digest("hex");
const BRIDGE_CASES = FIXTURE.cases.filter((wireCase) => wireCase.sides.includes("bridge"));
/**
 * Pinned independently of the fixture (R2d-3 G2 / A1): the fixture's own
 * `requiredTags` and per-case `sides` are metadata, so a fixture-only edit could
 * otherwise drop a tag or a leg without failing anything. Rule: a new or
 * changed case (e.g. R3a's C23) must be added deliberately to BOTH this table
 * and `EXPECTED_CASE_SIDES` in scripts/tests/test_companion_wire_contract.py,
 * and the pinned per-leg counts updated in both files.
 */
const REQUIRED_TAGS: readonly string[] = [
  "session", "pending", "execute_start", "completed", "not_executed", "duplicate", "cancelled",
  "executing_uncertain", "expired_claimed", "expired_unclaimed", "4xx", "5xx", "truncation", "mixed_version",
  "primitive", "invalid_ticket", "ack_timeout", "bridge_throw", "route_payload",
  "phase_inspect", "phase_fence",
];
const EXPECTED_CASE_SIDES: ReadonlyMap<string, readonly string[]> = new Map<string, readonly string[]>([
  ["C01-session",                                         ["companion", "bridge"]],
  ["C02-completed",                                       ["companion", "plugin", "bridge"]],
  ["C03-not-executed",                                    ["companion", "plugin", "bridge"]],
  ["C04-duplicate",                                       ["companion"]],
  ["C05-claimed-executed-true",                           ["companion"]],
  ["C06-lost-ack-response",                               ["companion", "plugin", "bridge"]],
  ["C07-double-ack",                                      ["companion", "bridge"]],
  ["C07-double-ack/late_ack",                             ["bridge"]],
  ["C08-cancel-pending",                                  ["companion", "plugin", "bridge"]],
  ["C09-cancel-claimed",                                  ["companion", "plugin", "bridge"]],
  ["C10-cancel-executing",                                ["companion", "plugin", "bridge"]],
  ["C11-claimed-lease-expiry",                            ["companion", "plugin", "bridge"]],
  ["C12-executing-window",                                ["companion"]],
  ["C13-unclaimed-expiry",                                ["companion", "plugin"]],
  ["C14-4xx-set/result-without-executed",                 ["companion"]],
  ["C14-4xx-set/execute-start-ticket",                    ["companion"]],
  ["C14-4xx-set/pane-auth",                               ["companion"]],
  ["C14-4xx-set/execute-start-401",                       ["companion", "bridge"]],
  ["C14-4xx-set/second-invoke-same-turn",                 ["companion", "plugin"]],
  ["C14-4xx-set/invoke-unknown-session",                  ["companion", "plugin"]],
  ["C15-5xx-transport/504-no-category",                   ["plugin"]],
  ["C15-5xx-transport/503-html",                          ["plugin"]],
  ["C15-5xx-transport/500-internal-json",                 ["plugin"]],
  ["C16-truncation/L",                                    ["plugin"]],
  ["C16-truncation/L-short-of-content-length",            ["plugin"]],
  ["C16-truncation/B1-pending",                           ["bridge"]],
  ["C16-truncation/B2-execute-start-truncated",           ["bridge"]],
  ["C16-truncation/B2b-execute-start-text-rejects",       ["bridge"]],
  ["C17-primitive",                                       ["companion", "bridge"]],
  ["C17-primitive/B-trusted",                             ["bridge"]],
  ["C18-invalid-plan-ticket/version-2",                   ["bridge"]],
  ["C18-invalid-plan-ticket/arguments-array",             ["bridge"]],
  ["C18-invalid-plan-ticket/trusted-missing",             ["bridge"]],
  ["C18-invalid-plan-ticket/trusted-extra-key",           ["bridge"]],
  ["C18-invalid-plan-ticket/trusted-foreign-fingerprint", ["bridge"]],
  ["C18-invalid-plan-ticket/trusted-turn-not-uuid",       ["bridge"]],
  ["C18-invalid-plan-ticket/trusted-v1-two-keys",         ["bridge"]],
  ["C18-invalid-plan-ticket/trusted-phase-unknown",       ["bridge"]],
  ["C18-invalid-plan-ticket/trusted-turnkey-equals-op",   ["bridge"]],
  ["C19a-throw-before-ack",                               ["companion", "bridge"]],
  ["C19b-throw-after-ack",                                ["companion", "bridge"]],
  ["C20-ack-timeout",                                     ["bridge"]],
  ["C20-ack-timeout/aborted-signal",                      ["bridge"]],
  ["MV1-old-pane-no-executed",                            ["companion", "plugin"]],
  ["MV2-old-pane-no-fingerprint",                         ["companion", "plugin"]],
  ["MV2-B-non-negotiated-plan-ticket",                    ["bridge"]],
  ["MV3-new-pane-r2c-companion",                          ["bridge"]],
  ["C21a-rejected",                                       ["companion", "plugin", "bridge"]],
  ["C21b-refuse",                                         ["companion", "plugin", "bridge"]],
  ["C21c-cancelled",                                      ["companion", "plugin", "bridge"]],
  ["C21d-route-uncertain",                                ["companion", "plugin", "bridge"]],
  ["C21e-verification-failed",                            ["companion", "plugin", "bridge"]],
  ["C21-R/R1-rollout-disabled",                           ["bridge"]],
  ["C21-R/R2-no-bridge-session",                          ["bridge"]],
  ["C22-not-started",                                     ["companion", "plugin"]],
  ["C23-plan-fence-capacity",                             ["companion", "plugin"]],
  ["C24-inspect-then-execute",                            ["companion", "plugin"]],
  ["C25-inspect-twice",                                   ["companion", "plugin"]],
  ["C26-inspect-after-execute",                           ["companion", "plugin"]],
  ["C27-inspect-ticket-cannot-execute",                   ["companion"]],
  ["C28-phase-outcome-mismatch",                          ["companion"]],
  ["MV4-v1-plugin-envelope",                              ["companion"]],
  ["MV5-v2-plugin-to-v1-companion",                       ["plugin"]],
  ["C29-phase-outcome-mismatch-ticket",                   ["bridge"]],
  ["C30-inspect-ticket-gate-on",                          ["bridge"]],
]);
const EXPECTED_BRIDGE_CASE_COUNT = 40;
const STEP_KINDS = [
  "pane", "pane_transport", "pane_hang", "plugin_invoke_begin", "plugin_invoke_end",
  "plugin_cancel", "clock_advance", "invoke_transport",
];
/** Steps the bridge itself issues (plugin, clock and companion-only steps are skipped). */
const BRIDGE_STEP_KINDS = new Set(["pane", "pane_transport", "pane_hang"]);
const PANE_BEHAVIORS = new Set([
  "ack_then_report", "decline_no_ack", "ack_expect_reject", "throw_before_ack", "throw_after_ack",
  "double_ack", "late_ack", "hang_ack", "ack_aborted_signal", "route_real", "primitive_report", "never_called",
]);
const THROWN_CANARY = "R2D3_THROWN_WORKBOOK_CANARY";

function bindings(): Record<string, Json> {
  const values: Record<string, Json> = { $fingerprint: FIXTURE.workbookPlanFingerprint };
  for (const [name, value] of Object.entries(FIXTURE.ids)) values[`$${name}`] = value;
  for (const [name, value] of Object.entries(FIXTURE.samples)) values[`$${name}`] = value;
  return values;
}
const BINDINGS = bindings();

function resolve(value: Json | undefined): Json | undefined {
  if (typeof value === "string" && value.startsWith("$")) {
    if (!(value in BINDINGS)) throw new Error(`unbound token ${value}`);
    return BINDINGS[value];
  }
  if (Array.isArray(value)) return value.map((item) => resolve(item) as Json);
  if (value !== null && typeof value === "object") {
    const out: JsonObject = {};
    for (const [key, item] of Object.entries(value)) out[key] = resolve(item) as Json;
    return out;
  }
  return value;
}

function bridgeSteps(wireCase: WireCase): WireStep[] {
  return wireCase.steps.filter((step) => BRIDGE_STEP_KINDS.has(step.kind) && step.only !== "python");
}

function pendingBody(wireCase: WireCase): JsonObject | undefined {
  const step = bridgeSteps(wireCase).find((candidate) => candidate.path === "/excel/pending"
    && candidate.response?.status === 200 && candidate.response.body !== undefined);
  return step === undefined ? undefined : resolve(step.response?.body) as JsonObject;
}

interface Recorded {
  readonly method: string;
  readonly url: string;
  readonly headers: Record<string, string>;
  readonly body: string | undefined;
  readonly signal: AbortSignal | undefined;
}

function replayFetch(wireCase: WireCase) {
  const queue = bridgeSteps(wireCase);
  const expectedPaths = queue.map((step) => step.path);
  const requests: Recorded[] = [];
  const mismatches: string[] = [];
  const hang: { signal?: AbortSignal } = {};
  const fetchImpl = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const recorded: Recorded = {
      method: init?.method ?? "GET",
      url: String(input),
      headers: { ...(init?.headers as Record<string, string> | undefined) },
      body: typeof init?.body === "string" ? init.body : init?.body === undefined ? undefined : "<non-string>",
      signal: init?.signal ?? undefined,
    };
    requests.push(recorded);
    const step = queue.shift();
    const where = `${wireCase.id} request #${requests.length} ${recorded.method} ${recorded.url}`;
    if (step === undefined) {
      mismatches.push(`${where}: unexpected request (no recorded step left)`);
      throw new Error("unexpected request");
    }
    if (recorded.method !== step.method) mismatches.push(`${where}: method, expected ${step.method}`);
    if (recorded.url !== step.path) mismatches.push(`${where}: path, expected ${step.path}`);
    const expectedHeaders = resolve(step.headers as Json) as Record<string, string>;
    if (JSON.stringify(Object.keys(recorded.headers).sort()) !== JSON.stringify(Object.keys(expectedHeaders).sort())) {
      mismatches.push(`${where}: header names ${JSON.stringify(Object.keys(recorded.headers))}`);
    }
    for (const [name, value] of Object.entries(expectedHeaders)) {
      if (recorded.headers[name] !== value) mismatches.push(`${where}: header ${name}`);
    }
    if (Object.keys(recorded.headers).some((name) => name.toLowerCase() === "origin")) {
      mismatches.push(`${where}: the bridge must not set Origin (the browser does)`);
    }
    if (step.body === undefined) {
      if (recorded.body !== undefined) mismatches.push(`${where}: unexpected body ${recorded.body}`);
    } else {
      const expectedBody = JSON.stringify(resolve(step.body));
      if (recorded.body !== expectedBody) {
        mismatches.push(`${where}: body\n  expected ${expectedBody}\n  actual   ${recorded.body}`);
      }
    }
    const isExecuteStart = step.path === "/excel/execute-start";
    if (isExecuteStart !== (recorded.signal instanceof AbortSignal)) {
      mismatches.push(`${where}: abort signal presence`);
    }
    if (step.kind === "pane_transport") throw new TypeError("fetch failed");
    if (step.kind === "pane_hang") {
      hang.signal = recorded.signal;
      return new Promise<Response>((_resolve, reject) => {
        recorded.signal?.addEventListener("abort", () => reject(new DOMException("aborted", "AbortError")), { once: true });
      });
    }
    const response = step.response as WireResponse;
    if (response.textRejects === true) {
      return { status: response.status, text: () => Promise.reject(new Error("body stream failed")) } as unknown as Response;
    }
    if (response.responseText !== undefined) return new Response(response.responseText, { status: response.status });
    if (response.empty === true) return new Response(null, { status: response.status });
    return new Response(JSON.stringify(resolve(response.body)), {
      status: response.status,
      headers: { "Content-Type": "application/json" },
    });
  };
  return { fetchImpl, queue, requests, mismatches, hang, expectedPaths };
}

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await new Promise<void>((done) => setImmediate(done));
};

async function expectAckRejected(promise: Promise<void>, errors: string[], label: string): Promise<void> {
  try {
    await promise;
    errors.push(`${label}: acknowledgement resolved`);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message !== FIXTURE.messages.ackRejected) errors.push(`${label}: rejected with ${message}`);
  }
}

async function runBridgeCase(wireCase: WireCase): Promise<void> {
  const spec = wireCase.bridge;
  if (spec === undefined || !PANE_BEHAVIORS.has(spec.paneBehavior)) {
    throw new Error(`${wireCase.id}: no bridge script for ${spec?.paneBehavior}`);
  }
  const replay = replayFetch(wireCase);
  const report = JSON.stringify(spec.paneReport ?? null);
  const calls: unknown[][] = [];
  const scriptErrors: string[] = [];
  const executeStarts = () => replay.requests.filter((request) => request.url === "/excel/execute-start").length;
  const state: { lateControl?: BridgePlanExecutionControl; ack: string; route?: Promise<unknown> } = { ack: "unused" };
  const routeTouched: string[] = [];

  let routeExecute: ((...args: unknown[]) => Promise<string>) | undefined;
  if (spec.paneBehavior === "route_real") {
    const deps = spec.routeDeps;
    if (deps === undefined) throw new Error(`${wireCase.id}: routeDeps missing`);
    const { createServerOwnedPlanRoute } = await import("../taskpane/serverOwnedPlanRoute");
    const sessionId = deps.bridgeSessionId === null ? null : resolve(deps.bridgeSessionId) as string;
    const dependencies = new Proxy({}, {
      get(_target, property) {
        if (property === "isRolloutEnabled") return () => deps.rolloutEnabled;
        if (property === "getBridgeSessionId") return () => sessionId;
        routeTouched.push(String(property));
        throw new Error("unexpected dependency");
      },
    }) as ServerOwnedPlanRouteDependencies;
    const route = createServerOwnedPlanRoute(dependencies);
    routeExecute = (...args: unknown[]) => route.executeCompanionTool(
      args[0] as string,
      args[1] as Record<string, unknown>,
      args[2] as BridgeTrustedPlanEnvelope | undefined,
      args[3] as BridgePlanExecutionControl | undefined,
    );
  }

  const executeTool = async (...args: unknown[]): Promise<string> => {
    calls.push(args);
    const control = args[3] as BridgePlanExecutionControl | undefined;
    switch (spec.paneBehavior) {
      case "ack_then_report":
        await control?.acknowledgeExecuteStart();
        return report;
      case "decline_no_ack":
      case "primitive_report":
        return report;
      case "ack_expect_reject":
        await expectAckRejected(Promise.resolve().then(() => control?.acknowledgeExecuteStart()), scriptErrors, "ack");
        return report;
      case "throw_before_ack":
        throw new Error(THROWN_CANARY);
      case "throw_after_ack":
        await control?.acknowledgeExecuteStart();
        throw new Error(THROWN_CANARY);
      case "double_ack": {
        await control?.acknowledgeExecuteStart();
        const before = replay.requests.length;
        await expectAckRejected(Promise.resolve().then(() => control?.acknowledgeExecuteStart()), scriptErrors, "second ack");
        if (replay.requests.length !== before) scriptErrors.push("second ack issued a request");
        return report;
      }
      case "late_ack":
        state.lateControl = control;
        return report;
      case "hang_ack": {
        state.ack = "pending";
        const ack = Promise.resolve().then(() => control?.acknowledgeExecuteStart());
        ack.then(() => { state.ack = "resolved"; }, (error: unknown) => {
          state.ack = error instanceof Error ? error.message : String(error);
        });
        await ack.catch(() => undefined);
        return report;
      }
      case "ack_aborted_signal": {
        const turn = new AbortController();
        turn.abort();
        await expectAckRejected(Promise.resolve().then(() => control?.acknowledgeExecuteStart(turn.signal)), scriptErrors, "aborted ack");
        return report;
      }
      case "route_real":
        return (routeExecute as (...a: unknown[]) => Promise<string>)(...args);
      default:
        scriptErrors.push(`executeTool called for ${spec.paneBehavior}`);
        return report;
    }
  };

  const bridge = new OfficeCompanionBridge({
    hostPrefix: FIXTURE.hostPrefix,
    documentId: async () => FIXTURE.ids.documentId,
    paneId: FIXTURE.ids.paneId,
    executeTool: executeTool as ConstructorParameters<typeof OfficeCompanionBridge>[0]["executeTool"],
    fetchImpl: replay.fetchImpl as typeof fetch,
  });

  // Bridge outcomes are collected, not thrown, so a contract break always
  // surfaces as an assertion on the recorded wire (and not as a stray error).
  const actionFailures: string[] = [];
  const describeError = (error: unknown): string => error instanceof Error ? `${error.name}: ${error.message}` : String(error);
  for (const [index, action] of spec.actions.entries()) {
    const label = `${wireCase.id} action #${index + 1} ${action.call}`;
    if (action.call === "connect") {
      try {
        const session = await bridge.connect();
        if (session.sessionId !== FIXTURE.samples.sessionId) actionFailures.push(`${label}: sessionId ${session.sessionId}`);
      } catch (error) {
        actionFailures.push(`${label}: rejected ${describeError(error)}`);
      }
      continue;
    }
    if (spec.paneBehavior === "hang_ack") {
      expect(EXECUTE_START_ACK_TIMEOUT_MS).toBe(FIXTURE.constants.executeStartAckTimeoutMs);
      expect(EXECUTE_START_ACK_TIMEOUT_MS).toBeLessThan(FIXTURE.constants.claimLeaseSeconds * 1000);
      // Fake only the ack timer; flushing uses the real setImmediate.
      vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
      const poll = bridge.pollOnce();
      for (let i = 0; i < 20 && executeStarts() === 0; i += 1) await flush();
      expect(executeStarts()).toBe(1);
      const timeout = FIXTURE.constants.executeStartAckTimeoutMs;
      vi.advanceTimersByTime(timeout - 1);
      await flush();
      expect(state.ack).toBe("pending");
      expect(replay.hang.signal?.aborted).toBe(false);
      vi.advanceTimersByTime(1);
      await flush();
      expect(state.ack).toBe(FIXTURE.messages.ackRejected);
      expect(replay.hang.signal?.aborted).toBe(true);
      expect(await poll).toBe(action.resolves);
      vi.useRealTimers();
      continue;
    }
    let resolved: boolean | undefined;
    let failure: unknown;
    try {
      resolved = await bridge.pollOnce();
    } catch (error) {
      failure = error ?? new Error("rejected without a reason");
    }
    if (action.rejects === undefined) {
      if (failure !== undefined) actionFailures.push(`${label}: rejected ${describeError(failure)}`);
      else if (resolved !== action.resolves) actionFailures.push(`${label}: resolved ${resolved}, expected ${action.resolves}`);
    } else if (!(failure instanceof Error)) {
      actionFailures.push(`${label}: resolved ${resolved}, expected a rejection`);
    } else {
      if (action.rejects.name !== undefined && failure.name !== action.rejects.name) {
        actionFailures.push(`${label}: rejected ${describeError(failure)}, expected ${action.rejects.name}`);
      }
      if (action.rejects.message !== undefined && failure.message !== action.rejects.message) {
        actionFailures.push(`${label}: rejected ${describeError(failure)}, expected "${action.rejects.message}"`);
      }
    }
  }

  if (spec.paneBehavior === "late_ack") {
    const before = replay.requests.length;
    if (state.lateControl === undefined) scriptErrors.push("late ack: no acknowledgement control was handed to executeTool");
    await expectAckRejected(Promise.resolve().then(() => state.lateControl?.acknowledgeExecuteStart()), scriptErrors, "late ack");
    if (replay.requests.length !== before) scriptErrors.push("late ack issued a request");
  }

  expect(replay.mismatches).toEqual([]);
  expect(actionFailures).toEqual([]);
  expect(scriptErrors).toEqual([]);
  expect(routeTouched).toEqual([]);
  expect(replay.queue.map((step) => `${step.method} ${step.path}`), "recorded requests the bridge never sent").toEqual([]);
  expect(replay.requests.map((request) => request.url)).toEqual(replay.expectedPaths);
  expect(executeStarts()).toBe(spec.expect.executeStartRequests);
  expect(calls.length).toBe(spec.expect.executeToolCalls);
  const pending = pendingBody(wireCase);
  for (const args of calls) {
    expect(args.length).toBe(spec.expect.executeToolArity);
    expect(args[0]).toBe(pending?.tool);
    expect(args[1]).toEqual(pending?.arguments);
    if (spec.expect.executeToolArity === 4) {
      expect(args[2]).toEqual(pending?.trusted);
      expect(Object.isFrozen(args[3])).toBe(true);
      expect(typeof (args[3] as BridgePlanExecutionControl).acknowledgeExecuteStart).toBe("function");
    }
  }
  for (const request of replay.requests) expect(request.body ?? "").not.toContain(THROWN_CANARY);
}

afterEach(() => {
  vi.useRealTimers();
});

describe("plan wire contract v2: fixture guard", () => {
  it("pins the fixture header to the bridge constants and logs the fixture sha256", () => {
    console.log(`plan-wire-contract.v2.json sha256=${FIXTURE_SHA256} cases=${FIXTURE.cases.length}`);
    expect(FIXTURE.contractVersion).toBe(2);
    expect(FIXTURE.protocolVersion).toBe("1.0");
    expect(FIXTURE.planTool).toBe(WORKBOOK_PLAN_TOOL_NAME);
    expect(FIXTURE.workbookPlanFingerprint).toBe(WORKBOOK_PLAN_FINGERPRINT);
    expect(EXECUTE_START_ACK_TIMEOUT_MS).toBe(FIXTURE.constants.executeStartAckTimeoutMs);
    expect(FIXTURE.constants.executeStartAckTimeoutMs).toBe(5000);
    expect(FIXTURE.constants.executeStartAckTimeoutMs).toBeLessThan(FIXTURE.constants.claimLeaseSeconds * 1000);
    expect(FIXTURE.messages.ackRejected).toBe("Workbook plan execution start was not acknowledged.");
  });

  it("pins the required tags and every case's legs independently of the fixture", () => {
    expect(REQUIRED_TAGS.length).toBe(21);
    expect(new Set(REQUIRED_TAGS).size).toBe(REQUIRED_TAGS.length);
    expect(FIXTURE.requiredTags).toEqual(REQUIRED_TAGS);
    const tags = new Set(FIXTURE.cases.flatMap((wireCase) => wireCase.tags));
    expect(REQUIRED_TAGS.filter((tag) => !tags.has(tag))).toEqual([]);
    const ids = FIXTURE.cases.map((wireCase) => wireCase.id);
    expect(ids.filter((id) => !EXPECTED_CASE_SIDES.has(id)), "cases not in EXPECTED_CASE_SIDES").toEqual([]);
    expect([...EXPECTED_CASE_SIDES.keys()].filter((id) => !ids.includes(id)), "pinned cases missing").toEqual([]);
    for (const wireCase of FIXTURE.cases) {
      expect([...wireCase.sides].sort(), wireCase.id).toEqual([...(EXPECTED_CASE_SIDES.get(wireCase.id) ?? [])].sort());
    }
    const expectedBridge = [...EXPECTED_CASE_SIDES].filter(([, sides]) => sides.includes("bridge")).map(([id]) => id);
    expect(expectedBridge.length).toBe(EXPECTED_BRIDGE_CASE_COUNT);
    expect(BRIDGE_CASES.map((wireCase) => wireCase.id)).toEqual(expectedBridge);
    expect(BRIDGE_CASES.length).toBe(EXPECTED_BRIDGE_CASE_COUNT);
  });

  it("covers every required tag and only known step kinds and sides", () => {
    const tags = new Set(FIXTURE.cases.flatMap((wireCase) => wireCase.tags));
    expect(FIXTURE.requiredTags.filter((tag) => !tags.has(tag))).toEqual([]);
    expect([...FIXTURE.stepKinds].sort()).toEqual([...STEP_KINDS].sort());
    const ids = FIXTURE.cases.map((wireCase) => wireCase.id);
    expect(new Set(ids).size).toBe(ids.length);
    for (const wireCase of FIXTURE.cases) {
      for (const side of wireCase.sides) expect(["companion", "plugin", "bridge"]).toContain(side);
      for (const step of wireCase.steps) expect(STEP_KINDS, `${wireCase.id} ${step.kind}`).toContain(step.kind);
      if (wireCase.companionVersion !== "current") expect(wireCase.sides, wireCase.id).toEqual(["bridge"]);
    }
    expect(BRIDGE_CASES.length).toBe(EXPECTED_BRIDGE_CASE_COUNT);
    for (const wireCase of BRIDGE_CASES) {
      expect(PANE_BEHAVIORS.has(wireCase.bridge?.paneBehavior ?? ""), wireCase.id).toBe(true);
    }
  });
});

describe("plan wire contract v2: real OfficeCompanionBridge replays every bridge case", () => {
  for (const wireCase of BRIDGE_CASES) {
    it(wireCase.id, async () => {
      await runBridgeCase(wireCase);
    });
  }
});

// ---------------------------------------------------------------------------
// R4h (dev3-r4h-release §5.1): the bridge and route non-executing plan payload
// shapes pinned as ONE consumer contract (D-R4-7: both shapes are kept; no
// unification). Every literal below is test-local and independent of the
// fixture and of production. H1-H3 read the bridge-leg /excel/result bodies,
// whose bytes the replay above already proves equal to the real bridge's posts;
// H4 drives the real route's pre-read gates (the route `invalid` reason has no
// fixture case, gap g2); H5 pins the closed reason set at compile time.
// ---------------------------------------------------------------------------
const BRIDGE_NOT_EXECUTED_ERRORS = {
  invalid: "Workbook plan ticket was invalid and was not executed.",
  rejected: "Workbook plan was not executed.",
} as const;
const ROUTE_NOT_EXECUTED_SUMMARIES = {
  ask: "The workbook plan asked for clarification; nothing was executed.",
  refuse: "The workbook plan was refused; nothing was executed.",
  rejected: "The workbook plan was rejected before execution; nothing was executed.",
  cancelled: "The workbook plan was cancelled before execution; nothing was executed.",
  not_started: "Workbook plan execution did not start; nothing was executed.",
  invalid: "The workbook plan ticket was invalid; nothing was executed.",
  already_executed: "This workbook operation may have been applied earlier and was not executed again.",
} as const satisfies Record<ServerOwnedPlanNotExecutedReason, string>;
/** H5 (i): adding or removing a production reason fails tsc here (and at the `satisfies` above). */
type _Exact = [ServerOwnedPlanNotExecutedReason] extends [keyof typeof ROUTE_NOT_EXECUTED_SUMMARIES]
  ? ([keyof typeof ROUTE_NOT_EXECUTED_SUMMARIES] extends [ServerOwnedPlanNotExecutedReason] ? true : never)
  : never;
const _exact: _Exact = true;
const NON_EXECUTING_REASONS: readonly string[] = [
  "ask", "refuse", "rejected", "cancelled", "not_started", "invalid", "already_executed",
];
const UNCERTAIN_ERROR = "Workbook plan execution ended without a confirmed result.";
const EXPECTED_BRIDGE_SHAPE_CASES: readonly string[] = [
  "C18-invalid-plan-ticket/version-2",
  "C18-invalid-plan-ticket/arguments-array",
  "C18-invalid-plan-ticket/trusted-missing",
  "C18-invalid-plan-ticket/trusted-extra-key",
  "C18-invalid-plan-ticket/trusted-foreign-fingerprint",
  "C18-invalid-plan-ticket/trusted-turn-not-uuid",
  "C18-invalid-plan-ticket/trusted-v1-two-keys",
  "C18-invalid-plan-ticket/trusted-phase-unknown",
  "C18-invalid-plan-ticket/trusted-turnkey-equals-op",
  "C19a-throw-before-ack",
  "MV2-B-non-negotiated-plan-ticket",
  "C29-phase-outcome-mismatch-ticket",
];
const EXPECTED_ROUTE_SHAPE_CASES: readonly string[] = [
  "C03-not-executed",
  "C06-lost-ack-response",
  "C07-double-ack/late_ack",
  "C09-cancel-claimed",
  "C11-claimed-lease-expiry",
  "C14-4xx-set/execute-start-401",
  "C20-ack-timeout",
  "C20-ack-timeout/aborted-signal",
  "MV3-new-pane-r2c-companion",
  "C21a-rejected",
  "C21b-refuse",
  "C21c-cancelled",
  "C21-R/R1-rollout-disabled",
  "C21-R/R2-no-bridge-session",
];
const EXPECTED_INSPECT_SHAPE_CASES: readonly string[] = ["C30-inspect-ticket-gate-on"];

type ShapeVerdict = { readonly ok: true; readonly reason: string } | { readonly ok: false; readonly problems: string[] };

function keysOf(value: JsonObject): string {
  return JSON.stringify(Object.keys(value));
}

/** T1: the bridge non-executing shape. */
function bridgeShape(body: JsonObject): ShapeVerdict {
  const problems: string[] = [];
  if (body.ok !== false) problems.push("ok");
  if (body.action !== "workbook_plan") problems.push("action");
  if (body.reason === "invalid") {
    if (keysOf(body) !== JSON.stringify(["ok", "action", "reason", "error"])) problems.push(`keys ${keysOf(body)}`);
    if (body.error !== BRIDGE_NOT_EXECUTED_ERRORS.invalid) problems.push("error literal");
  } else if (body.reason === "rejected") {
    if (keysOf(body) !== JSON.stringify(["ok", "action", "category", "reason", "error"])) problems.push(`keys ${keysOf(body)}`);
    if (body.category !== "Execution") problems.push("category");
    if (body.error !== BRIDGE_NOT_EXECUTED_ERRORS.rejected) problems.push("error literal");
  } else {
    problems.push(`reason ${JSON.stringify(body.reason)}`);
  }
  return problems.length === 0 ? { ok: true, reason: body.reason as string } : { ok: false, problems };
}

/** T2: the route non-executing shape (seven reasons). */
function routeShape(body: JsonObject): ShapeVerdict {
  const problems: string[] = [];
  if (body.ok !== false) problems.push("ok");
  if (body.action !== "workbook_plan") problems.push("action");
  const reason = body.reason;
  if (typeof reason !== "string" || !NON_EXECUTING_REASONS.includes(reason)) {
    problems.push(`reason ${JSON.stringify(reason)}`);
    return { ok: false, problems };
  }
  const withResult = reason === "ask" || reason === "refuse" || reason === "already_executed";
  const expectedKeys = withResult ? ["ok", "action", "reason", "summary", "result"] : ["ok", "action", "reason", "summary"];
  if (keysOf(body) !== JSON.stringify(expectedKeys)) problems.push(`keys ${keysOf(body)}`);
  if (body.summary !== ROUTE_NOT_EXECUTED_SUMMARIES[reason as ServerOwnedPlanNotExecutedReason]) problems.push("summary literal");
  if (withResult) {
    const result = body.result;
    if (result === null || typeof result !== "object" || Array.isArray(result)) {
      problems.push("result");
    } else if (reason === "already_executed") {
      if (keysOf(result) !== JSON.stringify(["execution"])) problems.push(`result keys ${keysOf(result)}`);
      const execution = result.execution;
      if (execution === null || typeof execution !== "object" || Array.isArray(execution)) problems.push("result.execution");
    } else {
      if (keysOf(result) !== JSON.stringify(["outcome", "finalSummary"])) problems.push(`result keys ${keysOf(result)}`);
      if (result.outcome !== reason) problems.push("result.outcome");
      // R4j (R4h lane-B B-A1): code points, as production boundedCodePointLength and the Python leg's len().
      if (typeof result.finalSummary !== "string" || [...result.finalSummary].length > 2048) problems.push("result.finalSummary");
    }
  }
  return problems.length === 0 ? { ok: true, reason } : { ok: false, problems };
}

/** T8 (R4g-3 rule X1): the route's inspection shape (gate on). */
function inspectShape(body: JsonObject): ShapeVerdict {
  const problems: string[] = [];
  if (keysOf(body) !== JSON.stringify(["ok", "action", "phase", "result"])) problems.push(`keys ${keysOf(body)}`);
  if (body.ok !== true) problems.push("ok");
  if (body.action !== "workbook_plan") problems.push("action");
  if (body.phase !== "inspect") problems.push("phase");
  const result = body.result;
  if (result === null || typeof result !== "object" || Array.isArray(result)) {
    problems.push("result");
  } else {
    if (keysOf(result) !== JSON.stringify(["outcome", "inspectionCount", "reads"])) problems.push(`result keys ${keysOf(result)}`);
    if (result.outcome !== "inspected") problems.push("result.outcome");
    if (result.inspectionCount !== 1) problems.push("result.inspectionCount");
    const reads = result.reads;
    if (!Array.isArray(reads) || !reads.every((read) => read !== null && typeof read === "object" && !Array.isArray(read) && typeof read.ok === "boolean")) {
      problems.push("result.reads");
    }
  }
  return problems.length === 0 ? { ok: true, reason: "inspected" } : { ok: false, problems };
}

interface PostedBody {
  readonly id: string;
  readonly executed: Json | undefined;
  readonly result: JsonObject;
}

function postedResultBodies(): PostedBody[] {
  const out: PostedBody[] = [];
  for (const wireCase of BRIDGE_CASES) {
    for (const step of bridgeSteps(wireCase)) {
      if (step.path !== "/excel/result" || step.body === undefined) continue;
      const body = resolve(step.body) as JsonObject;
      out.push({ id: wireCase.id, executed: body.executed, result: body.result as JsonObject });
    }
  }
  return out;
}

/**
 * Live (not fixture) bodies: the real OfficeCompanionBridge's own /excel/result
 * posts for the bridge-shape emitters (pane never called, or executeTool throws
 * before the acknowledgement), replayed from the fixture's recorded responses.
 */
async function liveBridgeShapeBodies(): Promise<PostedBody[]> {
  const out: PostedBody[] = [];
  for (const wireCase of BRIDGE_CASES) {
    const behavior = wireCase.bridge?.paneBehavior;
    if (behavior !== "never_called" && behavior !== "throw_before_ack") continue;
    const replay = replayFetch(wireCase);
    const bridge = new OfficeCompanionBridge({
      hostPrefix: FIXTURE.hostPrefix,
      documentId: async () => FIXTURE.ids.documentId,
      paneId: FIXTURE.ids.paneId,
      executeTool: (async () => { throw new Error(THROWN_CANARY); }) as ConstructorParameters<typeof OfficeCompanionBridge>[0]["executeTool"],
      fetchImpl: replay.fetchImpl as typeof fetch,
    });
    for (const action of wireCase.bridge?.actions ?? []) {
      if (action.call === "connect") await bridge.connect().catch(() => undefined);
      else await bridge.pollOnce().catch(() => undefined);
    }
    for (const request of replay.requests) {
      if (request.url !== "/excel/result" || request.body === undefined) continue;
      const body = JSON.parse(request.body) as JsonObject;
      if (body.executed === false) out.push({ id: wireCase.id, executed: body.executed, result: body.result as JsonObject });
    }
  }
  return out;
}

function nonExecutingShapes(): { readonly bridge: PostedBody[]; readonly route: PostedBody[]; readonly inspect: PostedBody[]; readonly problems: string[] } {
  const bridge: PostedBody[] = [];
  const route: PostedBody[] = [];
  const inspect: PostedBody[] = [];
  const problems: string[] = [];
  for (const posted of postedResultBodies()) {
    if (posted.executed !== false) continue;
    const t1 = bridgeShape(posted.result);
    const t2 = routeShape(posted.result);
    const t8 = inspectShape(posted.result);
    const matches = [t1, t2, t8].filter((verdict) => verdict.ok).length;
    if (matches !== 1) {
      problems.push(`${posted.id}: matches ${matches} of T1/T2/T8 (exactly one required) ${JSON.stringify({ t1, t2, t8 })}`);
      continue;
    }
    if ("summary" in posted.result && "error" in posted.result) problems.push(`${posted.id}: summary and error co-occur`);
    (t1.ok ? bridge : t2.ok ? route : inspect).push(posted);
  }
  return { bridge, route, inspect, problems };
}

describe("R4h GA2 non-executing payload consumer contract (D-R4-7: both shapes kept)", () => {
  it("H1 every bridge-leg executed:false body is exactly one of the bridge shape (T1) or the route shape (T2) or the inspection shape (T8); summary and error never co-occur", async () => {
    const { bridge, route, inspect, problems } = nonExecutingShapes();
    expect(problems).toEqual([]);
    expect(bridge.length + route.length + inspect.length).toBe(postedResultBodies().filter((posted) => posted.executed === false).length);
    expect(bridge.length).toBeGreaterThan(0);
    expect(route.length).toBeGreaterThan(0);
    for (const posted of [...bridge, ...route]) {
      expect(posted.result.ok, posted.id).toBe(false);
      expect(posted.result.action, posted.id).toBe("workbook_plan");
      expect(NON_EXECUTING_REASONS, posted.id).toContain(posted.result.reason);
      expect("state" in posted.result, posted.id).toBe(false);
    }
    for (const posted of inspect) {
      expect(routeShape(posted.result).ok, posted.id).toBe(false);
      expect(bridgeShape(posted.result).ok, posted.id).toBe(false);
      for (const key of ["reason", "summary", "state", "error", "execution"]) {
        expect(key in posted.result, `${posted.id} ${key}`).toBe(false);
        expect(key in (posted.result.result as JsonObject), `${posted.id} result.${key}`).toBe(false);
      }
    }
    // The bridge shape as the real bridge posts it (not only as recorded).
    const live = await liveBridgeShapeBodies();
    expect(live.map((posted) => posted.id)).toEqual(EXPECTED_BRIDGE_SHAPE_CASES);
    for (const posted of live) {
      expect(bridgeShape(posted.result), posted.id).toEqual({ ok: true, reason: posted.result.reason });
      expect(routeShape(posted.result).ok, posted.id).toBe(false);
      expect(JSON.stringify(posted.result), posted.id).toBe(JSON.stringify(bridge.find((fixture) => fixture.id === posted.id)?.result));
    }
  });

  it("H2 executed:true bodies never look non-executing: ok:false is the uncertain shape or carries an execution report; no reason or summary", () => {
    const problems: string[] = [];
    let uncertain = 0;
    let reported = 0;
    for (const posted of postedResultBodies()) {
      if (posted.executed !== true) continue;
      const body = posted.result;
      if ("reason" in body || "summary" in body) problems.push(`${posted.id}: reason/summary with executed:true`);
      if (body.ok !== false) continue;
      if (body.state === "uncertain") {
        uncertain += 1;
        const bridgeKeys = JSON.stringify(["ok", "action", "category", "state", "error"]);
        const routeKeys = JSON.stringify(["ok", "action", "state", "error"]);
        if (keysOf(body) !== bridgeKeys && keysOf(body) !== routeKeys) problems.push(`${posted.id}: uncertain keys ${keysOf(body)}`);
        if (keysOf(body) === bridgeKeys && body.category !== "Execution") problems.push(`${posted.id}: category`);
        if (body.action !== "workbook_plan") problems.push(`${posted.id}: action`);
        if (body.error !== UNCERTAIN_ERROR) problems.push(`${posted.id}: uncertain error literal`);
      } else {
        const result = body.result;
        const execution = result !== null && typeof result === "object" && !Array.isArray(result) ? result.execution : undefined;
        if (execution === null || typeof execution !== "object" || Array.isArray(execution)) {
          problems.push(`${posted.id}: executed:true ok:false without state uncertain or an execution report`);
        } else {
          reported += 1;
        }
      }
    }
    expect(problems).toEqual([]);
    expect(uncertain).toBe(2);
    expect(reported).toBe(1);
  });

  it("H3 coverage pins: the T1 and T2 bridge-leg case ids and the reasons they cover, or the inspection shape (T8) case ids (route invalid and already_executed are not fixture-covered)", () => {
    const { bridge, route, inspect } = nonExecutingShapes();
    const bridgeIds = bridge.map((posted) => posted.id);
    const routeIds = route.map((posted) => posted.id);
    const inspectIds = inspect.map((posted) => posted.id);
    expect(bridgeIds).toEqual(EXPECTED_BRIDGE_SHAPE_CASES);
    expect(routeIds).toEqual(EXPECTED_ROUTE_SHAPE_CASES);
    expect(inspectIds).toEqual(EXPECTED_INSPECT_SHAPE_CASES);
    expect(bridgeIds.filter((id) => routeIds.includes(id) || inspectIds.includes(id))).toEqual([]);
    expect(routeIds.filter((id) => inspectIds.includes(id))).toEqual([]);
    for (const id of [...bridgeIds, ...routeIds, ...inspectIds]) expect(EXPECTED_CASE_SIDES.get(id) ?? [], id).toContain("bridge");
    expect([...new Set(bridge.map((posted) => posted.result.reason))].sort()).toEqual(["invalid", "rejected"]);
    const routeReasons = [...new Set(route.map((posted) => String(posted.result.reason)))].sort();
    expect(routeReasons).toEqual(["ask", "cancelled", "not_started", "refuse", "rejected"]);
    expect(NON_EXECUTING_REASONS.filter((reason) => !routeReasons.includes(reason)).sort()).toEqual(["already_executed", "invalid"]);
  });

  it("H4 the real route's pre-read gates (rollout off, no session, foreign fingerprint) emit the T2 shape and consult nothing else", async () => {
    const forbidden = () => vi.fn(() => { throw new Error("must not be called"); });
    const stubs = {
      coordinator: {
        authorizeAndRun: forbidden(),
        consumeWorkbookPlanConfirmation: forbidden(),
        cancelPendingConfirmation: forbidden(),
      },
      bindings: { finishTerminal: forbidden() },
      readRuntimeContext: forbidden(),
      createAdapter: forbidden(),
      createStore: forbidden(),
      inspect: forbidden(),
      awaitPaneConfirmation: forbidden(),
      executePrimitive: forbidden(),
      now: forbidden(),
    };
    const stubCalls = (): number => [
      stubs.coordinator.authorizeAndRun, stubs.coordinator.consumeWorkbookPlanConfirmation,
      stubs.coordinator.cancelPendingConfirmation, stubs.bindings.finishTerminal, stubs.readRuntimeContext,
      stubs.createAdapter, stubs.createStore, stubs.inspect, stubs.awaitPaneConfirmation, stubs.executePrimitive, stubs.now,
    ].reduce((sum, stub) => sum + stub.mock.calls.length, 0);
    const activeTurn = {
      turnOperationId: "t",
      trustedRequest: { classification: "explicit_mutation", selectionScopedRequest: false },
      sessionId: "S",
      turnStart: 1,
      hasCancellation: true,
    };
    const foreignFingerprint = "0".repeat(64);
    expect(foreignFingerprint).not.toBe(WORKBOOK_PLAN_FINGERPRINT);
    const drives = [
      { label: "rollout off", rollout: false, session: "S", reason: "not_started", reads: ["isRolloutEnabled"] },
      { label: "no bridge session", rollout: true, session: null, reason: "rejected", reads: ["isRolloutEnabled", "getBridgeSessionId"] },
      { label: "foreign fingerprint", rollout: true, session: "S", reason: "invalid", reads: ["isRolloutEnabled", "getBridgeSessionId", "getActiveTurn"] },
    ] as const;
    for (const drive of drives) {
      const reads: string[] = [];
      const gates = {
        isRolloutEnabled: vi.fn(() => drive.rollout),
        getBridgeSessionId: vi.fn(() => drive.session),
        getActiveTurn: vi.fn(() => activeTurn),
      };
      const target: Record<string, unknown> = { ...stubs, ...gates };
      const dependencies = new Proxy(target, {
        get(object, property) {
          reads.push(String(property));
          return Reflect.get(object, property);
        },
      }) as unknown as ServerOwnedPlanRouteDependencies;
      const route = createServerOwnedPlanRoute(dependencies);
      const payload = await route.executeServerOwnedWorkbookPlan(
        { version: 1, outcome: "ask", message: "m" },
        { turnOperationId: "t", workbookPlanFingerprint: foreignFingerprint } as BridgeTrustedPlanEnvelope,
      );
      const body = JSON.parse(payload) as JsonObject;
      const verdict = routeShape(body);
      expect(verdict, `${drive.label}: ${payload}`).toEqual({ ok: true, reason: drive.reason });
      expect(bridgeShape(body).ok, drive.label).toBe(false);
      expect(body, drive.label).toEqual({
        ok: false, action: "workbook_plan", reason: drive.reason, summary: ROUTE_NOT_EXECUTED_SUMMARIES[drive.reason],
      });
      expect(reads, drive.label).toEqual(drive.reads);
      expect(stubCalls(), drive.label).toBe(0);
    }
  });

  it("H5 the closed reason set is exactly seven (compile-time closure against ServerOwnedPlanNotExecutedReason; independent run-time literals)", () => {
    expect(_exact).toBe(true);
    const keys = Object.keys(ROUTE_NOT_EXECUTED_SUMMARIES).sort();
    expect(keys).toEqual([...NON_EXECUTING_REASONS].sort());
    expect(keys).toEqual(["already_executed", "ask", "cancelled", "invalid", "not_started", "refuse", "rejected"]);
    for (const [reason, summary] of Object.entries(ROUTE_NOT_EXECUTED_SUMMARIES)) {
      expect(summary.endsWith("executed.") || summary.endsWith("executed again."), reason).toBe(true);
      expect(summary.includes("may have been applied"), reason).toBe(reason === "already_executed");
      expect(summary.includes("already applied"), reason).toBe(false);
    }
    expect(Object.keys(BRIDGE_NOT_EXECUTED_ERRORS).sort()).toEqual(["invalid", "rejected"]);
  });

  it("H6 [R4j] routeShape bounds finalSummary in code points like production: 2048 astral characters pass, 2049 fail; the production validator agrees", async () => {
    const body = (reason: "ask" | "refuse", finalSummary: string): JsonObject => ({
      ok: false, action: "workbook_plan", reason, summary: ROUTE_NOT_EXECUTED_SUMMARIES[reason],
      result: { outcome: reason, finalSummary },
    });
    const EMOJI = "\u{1F600}";
    const rows: readonly (readonly [string, string, boolean])[] = [
      ["2048 emoji (4096 UTF-16 units)", EMOJI.repeat(2048), true],
      ["2049 emoji", EMOJI.repeat(2049), false],
      ["2048 ascii", "x".repeat(2048), true],
      ["2049 ascii", "x".repeat(2049), false],
      ["2047 ascii + 1 emoji", `${"x".repeat(2047)}${EMOJI}`, true],
    ];
    const observed: string[] = [];
    const expected: string[] = [];
    for (const reason of ["ask", "refuse"] as const) {
      for (const [label, text, ok] of rows) {
        const verdict = routeShape(body(reason, text));
        observed.push(`${reason} ${label}: ${verdict.ok ? "ok" : JSON.stringify(verdict.problems)}`);
        expected.push(`${reason} ${label}: ${ok ? "ok" : JSON.stringify(["result.finalSummary"])}`);
      }
    }
    expect(observed).toEqual(expected);
    const { validateWorkbookPlan } = await import("../workbookPlanValidation");
    expect(validateWorkbookPlan({ version: 1, outcome: "ask", message: EMOJI.repeat(2048) }).ok).toBe(true);
    expect(validateWorkbookPlan({ version: 1, outcome: "ask", message: EMOJI.repeat(2049) }).ok).toBe(false);
  });
});
