/**
 * R4d server-route integration tests (dev3-r4d-release.md §5.1 S1-S13, SX;
 * §5.2 SP1-SP3, F1, G2).
 *
 * Real chain: createExcelCompanionBridge -> OfficeBridgeLifecycleController ->
 * TaskPaneTurnCoordinator -> createServerOwnedPlanRoute, whose
 * `executePrimitive` is built exactly as taskpane.ts builds it:
 * createServerPrimitiveGate({ ... resolveAuthority: coordinator.serverPrimitiveAuthority,
 * executeWithAuthority: executeValidatedExcelBridgeToolWithAuthority,
 * executeReadOnly: executeValidatedExcelBridgeTool }). At the base, where the
 * gate is absent, the harness falls back to the base production value
 * (officeBridgeExcel.executeLegacyServerPrimitiveR4dPending), so RED is
 * measured against the real base wiring. New APIs are reached through casts so
 * a missing API is an assertion failure, never a type error.
 *
 * Doubles: a session-scoped companion double (copied from
 * task-pane-session-rotation-r3b.test.ts) with a hold-able GET /pending, and
 * the recording fake Excel.run (copied from primitive-authority-direct.test.ts).
 * Oracle: validateToolCall, then getExcelToolDefinition(name).execute on an
 * identically seeded fresh double (the base validated bridge path).
 */
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { afterEach, beforeAll, describe, expect, it, vi } from "vitest";
import { getExcelToolDefinition } from "../excelTools";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";
import type { BridgePublicState, OfficeCompanionBridge } from "../officeBridge";
import * as officeBridgeExcelModule from "../officeBridgeExcel";
import {
  createExcelCompanionBridge,
  executeValidatedExcelBridgeTool,
  executeValidatedExcelBridgeToolWithAuthority,
} from "../officeBridgeExcel";
import { OfficeBridgeLifecycleController } from "../officeBridgeLifecycle";
import type { PrimitiveMutationAuthority } from "../primitiveMutationAuthority";
import { createServerOwnedPlanRoute } from "../taskpane/serverOwnedPlanRoute";
import { TaskPaneTurnCoordinator, type TaskPaneTurnHandle, type TaskPaneTurnPreparation } from "../taskpane/taskPaneTurnCoordinator";
import { validateToolCall } from "../toolValidation";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import {
  BrowserIdempotencyStore,
  MemoryIdempotencyStorage,
  type WorkbookPlanExecutionAdapter,
} from "../workbookPlanExecutor";
import {
  ServerOwnedTurnBindingStore,
  type WorkbookPlanConfirmationBinding,
  type WorkbookPlanVerificationAdapter,
} from "../workbookPlanOrchestrator";
import { classifyWorkbookRequest, type TrustedWorkbookRequest } from "../workbookPlanRequest";
import type { WorkbookPlanRouteRuntimeContext } from "../workbookPlanRoute";

// ---------------------------------------------------------------------------
// Recording Office double (R4c §5 doubles). Every Excel.run, sync, queued
// property set, load and method call is appended to `log`. Range values are
// kept per cell so a readback sees what was written; `mismatchFormulas` makes
// every `formulas` read return a mismatching grid (the readback hook).
// ---------------------------------------------------------------------------
interface OfficeDouble {
  readonly log: string[];
  mismatchFormulas: boolean;
  failRuns: boolean;
  install(): void;
}

const MAX_ROWS = 1_048_576;

function colNumber(label: string): number {
  let n = 0;
  for (const ch of label.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function colLabel(n: number): string {
  let label = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

interface DoubleRange { readonly row: number; readonly col: number; readonly rows: number; readonly columns: number }

function doubleRangeOf(address: string): DoubleRange {
  const bare = (address.includes("!") ? address.slice(address.lastIndexOf("!") + 1) : address).replace(/\$/g, "").trim();
  const cell = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/i.exec(bare);
  if (cell !== null) {
    const c1 = colNumber(cell[1]);
    const r1 = Number(cell[2]);
    const c2 = colNumber(cell[3] ?? cell[1]);
    const r2 = Number(cell[4] ?? cell[2]);
    return { row: Math.min(r1, r2), col: Math.min(c1, c2), rows: Math.abs(r2 - r1) + 1, columns: Math.abs(c2 - c1) + 1 };
  }
  const whole = /^([A-Z]{1,3}):([A-Z]{1,3})$/i.exec(bare);
  if (whole !== null) {
    const c1 = colNumber(whole[1]);
    const c2 = colNumber(whole[2]);
    return { row: 1, col: Math.min(c1, c2), rows: MAX_ROWS, columns: Math.abs(c2 - c1) + 1 };
  }
  return { row: 1, col: 26, rows: 1, columns: 1 };
}

function stableJson(value: unknown): string {
  try {
    const text = JSON.stringify(value);
    return text === undefined ? "undefined" : text;
  } catch {
    return "<unserializable>";
  }
}

function createOfficeDouble(seed: Readonly<Record<string, unknown>> = { "Sheet1!C1": "=A1*2" }): OfficeDouble {
  const log: string[] = [];
  const cells = new Map<string, unknown>(Object.entries(seed));
  const state = { mismatchFormulas: false, failRuns: false };

  const generic = (path: string): unknown =>
    new Proxy(function genericNode() { /* recording node */ }, {
      get(_target, prop) {
        if (typeof prop === "symbol" || prop === "then" || prop === "toJSON") return undefined;
        return generic(`${path}.${prop}`);
      },
      set(_target, prop, value) {
        log.push(`set ${path}.${String(prop)}=${stableJson(value)}`);
        return true;
      },
      apply(_target, _this, args) {
        log.push(`call ${path}(${stableJson(args)})`);
        return generic(`${path}()`);
      },
    });

  const wrap = (object: Record<string, unknown>, path: string): unknown =>
    new Proxy(object, {
      get(target, prop) {
        if (Object.prototype.hasOwnProperty.call(target, prop)) return Reflect.get(target, prop);
        if (typeof prop === "symbol" || prop === "then") return undefined;
        return generic(`${path}.${prop}`);
      },
      set(target, prop, value) {
        const descriptor = Object.getOwnPropertyDescriptor(target, prop);
        if (descriptor?.set !== undefined) {
          descriptor.set.call(target, value);
        } else {
          log.push(`set ${path}.${String(prop)}=${stableJson(value)}`);
        }
        return true;
      },
    });

  const cellKey = (sheet: string, row: number, col: number): string => `${sheet}!${colLabel(col)}${row}`;

  const makeRange = (sheet: string, address: string): unknown => {
    const r = doubleRangeOf(address);
    const path = `range(${sheet}!${address})`;
    const grid = (fill: (row: number, col: number) => unknown): unknown[][] =>
      Array.from({ length: Math.min(r.rows, 6000) }, (_v, i) =>
        Array.from({ length: r.columns }, (_w, j) => fill(r.row + i, r.col + j)));
    const read = (row: number, col: number): unknown => cells.get(cellKey(sheet, row, col)) ?? "";
    const object: Record<string, unknown> = {
      toJSON: () => `<${path}>`,
      isNullObject: false,
      load: (fields: unknown) => { log.push(`load ${path} ${stableJson(fields)}`); },
      select: () => { log.push(`select ${path}`); },
      getCell: (row: number, col: number) => makeRange(sheet, `${colLabel(r.col + col)}${r.row + row}`),
      getResizedRange: (rows: number, columns: number) =>
        makeRange(sheet, `${colLabel(r.col)}${r.row}:${colLabel(r.col + columns)}${r.row + rows}`),
      getMergedAreasOrNullObject: () => wrap({ isNullObject: true, load: () => undefined }, `${path}.merged`),
      autoFill: (target: { address?: unknown }) => {
        log.push(`autoFill ${path} -> ${String(target?.address)}`);
        const source = read(r.row, r.col);
        const t = doubleRangeOf(String(target?.address ?? ""));
        for (let i = 0; i < t.rows; i += 1) {
          for (let j = 0; j < t.columns; j += 1) cells.set(cellKey(sheet, t.row + i, t.col + j), source);
        }
      },
    };
    Object.defineProperties(object, {
      address: { get: () => `${sheet}!${address}`, enumerable: false, configurable: true },
      rowCount: { get: () => r.rows, enumerable: false, configurable: true },
      columnCount: { get: () => r.columns, enumerable: false, configurable: true },
      values: { get: () => grid(read), enumerable: false, configurable: true },
      numberFormat: {
        get: () => grid(() => "General"),
        set: (value: unknown) => { log.push(`set ${path}.numberFormat=${stableJson(value)}`); },
        enumerable: false, configurable: true,
      },
      formulas: {
        get: () => (state.mismatchFormulas ? grid(() => "MISMATCH") : grid(read)),
        set: (value: unknown) => {
          log.push(`set ${path}.formulas=${stableJson(value)}`);
          if (Array.isArray(value)) {
            value.forEach((row, i) => {
              if (Array.isArray(row)) row.forEach((cell, j) => cells.set(cellKey(sheet, r.row + i, r.col + j), cell));
            });
          }
        },
        enumerable: false, configurable: true,
      },
      formulasR1C1: {
        get: () => grid(() => "=R[-1]C*2"),
        set: (value: unknown) => { log.push(`set ${path}.formulasR1C1=${stableJson(value)}`); },
        enumerable: false, configurable: true,
      },
    });
    return wrap(object, path);
  };

  const makeSheet = (name: string): unknown => wrap({
    toJSON: () => `<sheet ${name}>`,
    name,
    load: (fields: unknown) => { log.push(`load sheet(${name}) ${stableJson(fields)}`); },
    getRange: (address: string) => makeRange(name, address),
    getUsedRangeOrNullObject: () => wrap({ isNullObject: true, load: () => undefined, toJSON: () => "<used>" }, `sheet(${name}).used`),
    activate: () => { log.push(`activate sheet(${name})`); },
    delete: () => { log.push(`delete sheet(${name})`); },
    comments: wrap({
      getItemByCell: (range: unknown) => {
        log.push(`comments(${name}).getItemByCell(${stableJson(range)})`);
        return wrap({ load: () => undefined }, `comment(${name})`);
      },
      add: (range: unknown, text: unknown) => { log.push(`comments(${name}).add(${stableJson(range)},${stableJson(text)})`); },
    }, `sheet(${name}).comments`),
  }, `sheet(${name})`);

  const worksheets = wrap({
    toJSON: () => "<worksheets>",
    getItem: (name: string) => makeSheet(name),
    getItemOrNullObject: (name: string) => makeSheet(name),
    add: (name: string) => { log.push(`worksheets.add(${stableJson(name)})`); return makeSheet(name); },
    getActiveWorksheet: () => makeSheet("Sheet1"),
    load: (fields: unknown) => { log.push(`load worksheets ${stableJson(fields)}`); },
    items: [makeSheet("Sheet1")],
  }, "worksheets");

  const context = {
    workbook: wrap({
      worksheets,
      application: wrap({ calculate: (kind: unknown) => { log.push(`calculate(${stableJson(kind)})`); } }, "application"),
      getSelectedRange: () => wrap({ address: "Sheet1!A1", load: () => undefined }, "selection"),
    }, "workbook"),
    sync: async (): Promise<void> => { log.push("sync"); },
  };

  return {
    log,
    get mismatchFormulas() { return state.mismatchFormulas; },
    set mismatchFormulas(value: boolean) { state.mismatchFormulas = value; },
    get failRuns() { return state.failRuns; },
    set failRuns(value: boolean) { state.failRuns = value; },
    install() {
      (globalThis as Record<string, unknown>).Excel = {
        run: async (work: (ctx: unknown) => Promise<unknown>) => {
          log.push("Excel.run");
          if (state.failRuns) throw Object.assign(new Error("Host boom"), { code: "GeneralException", name: "RichApi.Error" });
          return work(context);
        },
      };
    },
  };
}

function excelRuns(log: readonly string[]): number {
  return log.filter((entry) => entry === "Excel.run").length;
}

function formulaWrites(log: readonly string[]): string[] {
  return log.filter((entry) => /^set range\(.*\)\.formulas=/.test(entry));
}

function runSegments(log: readonly string[]): string[][] {
  const starts = log.map((entry, index) => (entry === "Excel.run" ? index : -1)).filter((index) => index >= 0);
  return starts.map((start, i) => log.slice(start, starts[i + 1] ?? log.length));
}

// ---------------------------------------------------------------------------
// Fixed expectations (independent constants)
// ---------------------------------------------------------------------------
const deny = (error: string, category: string): string => JSON.stringify({ ok: false, error, category });
const DENIAL = {
  TURN_INACTIVE: deny("Workbook edit denied: no active task-pane turn.", "validation"),
  LEGACY_NO_PLAN_PATH: deny("Workbook edit denied: this legacy tool has no WorkbookPlan path and is disabled while WorkbookPlan rollout is on.", "validation"),
  CONFIRMATION_REQUIRED: deny("Workbook edit denied: this request needs confirmation. Use excel_execute_plan (WorkbookPlan) instead.", "validation"),
  DUPLICATE: deny("Workbook edit denied: the identical edit was already applied in this turn.", "repetition"),
  RESOURCE_LIMIT: deny("Workbook edit denied: the per-turn workbook edit limit was reached.", "validation"),
  PLAN_ENTERED: deny("Workbook edit denied: a workbook plan already ran in this turn.", "validation"),
} as const;

const ZERO_USAGE = {
  operationCount: 0, cellsRead: 0, cellsWritten: 0, formulaUtf8Bytes: 0, csvPayloadUtf8Bytes: 0,
  commentUtf8Bytes: 0, sheetsCreated: 0, chartsCreated: 0, tablesCreated: 0, verificationReads: 0,
};

const PROMPTS = {
  fill: "fill the sheet with numbers",
  borra: "borra la hoja",
  summarize: "summarize the totals",
  createTable: "create a table from this data",
  selected: "update the selected cells",
  showTable: "show me the table",
  writeOk: "write ok into A1",
  usual: "do the usual thing",
} as const;

const EXPECTED_CLASS: Readonly<Record<string, readonly [string, boolean]>> = {
  [PROMPTS.fill]: ["explicit_mutation", false],
  [PROMPTS.borra]: ["ambiguous_mutation", false],
  [PROMPTS.summarize]: ["read_only", false],
  [PROMPTS.createTable]: ["explicit_structure", false],
  [PROMPTS.selected]: ["explicit_mutation", true],
  [PROMPTS.showTable]: ["read_only", false],
  [PROMPTS.writeOk]: ["explicit_mutation", false],
  [PROMPTS.usual]: ["ambiguous_mutation", false],
};

function classified(prompt: string): TrustedWorkbookRequest {
  const request = classifyWorkbookRequest(prompt);
  expect([request.classification, request.selectionScopedRequest]).toEqual(EXPECTED_CLASS[prompt]);
  return request;
}

const DOC1 = "d0c00000-0000-4000-8000-000000000001";
const DOC2 = "d0c00000-0000-4000-8000-000000000002";
const PANE = "9a9e0000-0000-4000-8000-0000000000aa";
const T1 = "7a7a0000-0000-4000-8000-0000000000d1";
const T2 = "7a7a0000-0000-4000-8000-0000000000d2";
const H1_OP = "5e7e0000-0000-4000-8000-0000000000d1";
const T2_OP = "5e7e0000-0000-4000-8000-0000000000d2";
const NOT_FOUND_BODY = { ok: false, error: "Session not found", category: "NotFound" };

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
const ASK_PLAN = { version: 1, outcome: "ask", message: "Please clarify" } as const;

// ---------------------------------------------------------------------------
// R4d surfaces (absent at the base): reached defensively
// ---------------------------------------------------------------------------
type PrimitiveExecutor = (name: string, args: Record<string, unknown>) => Promise<string>;
type GateFactory = (deps: Record<string, unknown>) => PrimitiveExecutor;
interface ServerAuthorityApi {
  serverPrimitiveAuthority?: (input: { readonly turnId: string; readonly bridgeSessionId: string }) => PrimitiveMutationAuthority | null;
}

let gateFactory: GateFactory | undefined;
beforeAll(async () => {
  const mod = (await import("../primitiveMutationAuthority")) as unknown as Record<string, unknown>;
  gateFactory = typeof mod.createServerPrimitiveGate === "function" ? mod.createServerPrimitiveGate as GateFactory : undefined;
});

function serverApi(coordinator: TaskPaneTurnCoordinator): ServerAuthorityApi {
  return coordinator as unknown as ServerAuthorityApi;
}

/** Base production value of taskpane.ts `executePrimitive` (absent after R4d). */
function basePassThrough(): PrimitiveExecutor {
  const exported = (officeBridgeExcelModule as unknown as Record<string, unknown>).executeLegacyServerPrimitiveR4dPending;
  expect(exported).toBeTypeOf("function");
  return exported as PrimitiveExecutor;
}

// ---------------------------------------------------------------------------
// Companion double (copied from task-pane-session-rotation-r3b.test.ts:105-275,
// plus a hold-able GET /pending)
// ---------------------------------------------------------------------------
function json(status: number, body?: unknown): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
  });
}

function uuid(prefix: string, n: number): string {
  return `${prefix}-0000-4000-8000-${String(n).padStart(12, "0")}`;
}

type TicketState = "pending" | "claimed" | "executing" | "completed" | "not_executed" | "cancelled" | "uncertain";

interface SessionTicket {
  readonly ticketId: string;
  readonly requestId: string;
  readonly tool: string;
  readonly args: Record<string, unknown>;
  readonly trusted?: unknown;
  state: TicketState;
}

interface SessionRecord {
  readonly sessionId: string;
  readonly sessionKey: string;
  readonly documentId: string;
  readonly tickets: SessionTicket[];
}

class SessionCompanion {
  readonly byKey = new Map<string, SessionRecord>();
  readonly byId = new Map<string, SessionRecord>();
  readonly deleted: string[] = [];
  readonly results: { sessionId: string; body: Record<string, unknown> }[] = [];
  executeStarts = 0;
  heldPending = 0;
  private hold: Promise<void> | null = null;
  private sessionCounter = 0;
  private ticketCounter = 0;

  constructor(private readonly events: string[]) {}

  holdNextPending(): () => void {
    let release!: () => void;
    this.hold = new Promise<void>((done) => { release = done; });
    return () => release();
  }

  invoke(sessionId: string, serverOperationId: string, plan: Record<string, unknown>) {
    return this.enqueue(sessionId, "excel_execute_plan", { ...plan }, {
      turnOperationId: serverOperationId,
      workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
      phase: "execute",
      turnKey: "4b4b0000-0000-4000-8000-0000000000a3",
    });
  }

  invokePrimitive(sessionId: string, tool: string, args: Record<string, unknown>) {
    return this.enqueue(sessionId, tool, args, undefined);
  }

  private enqueue(sessionId: string, tool: string, args: Record<string, unknown>, trusted: unknown) {
    const session = this.byId.get(sessionId);
    if (session === undefined) return { status: 404, body: { ...NOT_FOUND_BODY } as Record<string, unknown> };
    this.ticketCounter += 1;
    const ticket: SessionTicket = {
      ticketId: uuid("71c70000", this.ticketCounter),
      requestId: uuid("4e4e0000", this.ticketCounter),
      tool,
      args,
      ...(trusted === undefined ? {} : { trusted }),
      state: "pending",
    };
    session.tickets.push(ticket);
    return { status: 202, body: { ok: true } as Record<string, unknown>, ticketId: ticket.ticketId };
  }

  resultsFor(ticketId: string | undefined): Record<string, unknown>[] {
    return this.results.filter((entry) => entry.body.ticketId === ticketId).map((entry) => entry.body);
  }

  readonly fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const url = String(input);
    const route = url.replace(/^.*\/excel\//, "");
    const method = init?.method ?? "GET";
    const headers = (init?.headers ?? {}) as Record<string, string>;
    const body = typeof init?.body === "string" && init.body ? JSON.parse(init.body) as Record<string, unknown> : {};
    if (method === "POST" && route === "session") {
      this.sessionCounter += 1;
      const sessionId = uuid("aaaaaaaa", this.sessionCounter).replace(/^aaaaaaaa-0000-4000-8000/, "aaaaaaaa-aaaa-4aaa-8aaa");
      const sessionKey = `${"k".repeat(40)}${this.sessionCounter}`;
      const record: SessionRecord = { sessionId, sessionKey, documentId: String(body.documentId), tickets: [] };
      this.byKey.set(sessionKey, record);
      this.byId.set(sessionId, record);
      this.events.push(`session:${sessionId}`);
      return json(201, {
        ok: true, sessionId, sessionKey, protocolVersion: "1.0", expiresInSeconds: 60,
        workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
      });
    }
    const session = this.byKey.get(headers["X-Hermes-Session-Key"] ?? "");
    if (session === undefined || headers["X-Hermes-Session-Id"] !== session.sessionId) {
      return json(401, { error: "Unauthorized" });
    }
    if (method === "DELETE" && route === "session") {
      this.byKey.delete(session.sessionKey);
      this.byId.delete(session.sessionId);
      for (const ticket of session.tickets) {
        if (ticket.state === "pending" || ticket.state === "claimed") ticket.state = "cancelled";
        else if (ticket.state === "executing") ticket.state = "uncertain";
      }
      this.deleted.push(session.sessionId);
      this.events.push(`delete:${session.sessionId}`);
      return json(200, { ok: true });
    }
    if (method === "POST" && route === "heartbeat") return json(200, { ok: true });
    if (method === "GET" && route === "pending") {
      const ticket = session.tickets.find((candidate) => candidate.state === "pending");
      if (ticket === undefined) return json(204);
      ticket.state = "claimed";
      this.events.push("companion-claim");
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
    if (method === "POST" && route === "execute-start") {
      this.executeStarts += 1;
      const ticket = session.tickets.find((candidate) => candidate.ticketId === body.ticketId);
      if (ticket === undefined || ticket.state !== "claimed") return json(409, { error: "conflict" });
      ticket.state = "executing";
      this.events.push("execute-start");
      return json(200, { ok: true, ticketId: ticket.ticketId });
    }
    if (method === "POST" && route === "result") {
      this.events.push("result");
      this.results.push({ sessionId: session.sessionId, body });
      const ticket = session.tickets.find((candidate) => candidate.ticketId === body.ticketId);
      if (ticket === undefined) return json(404, { error: "No ticket found" });
      if (ticket.tool !== "excel_execute_plan") {
        if (ticket.state !== "claimed") return json(409, { error: "conflict" });
        ticket.state = "completed";
        return json(200, { ok: true, ticketId: ticket.ticketId });
      }
      if (typeof body.executed !== "boolean") return json(400, { error: "executed required" });
      if (body.executed === true && ticket.state !== "executing") return json(409, { error: "conflict" });
      if (body.executed === false && ticket.state !== "claimed") return json(409, { error: "conflict" });
      ticket.state = body.executed === true ? "completed" : "not_executed";
      return json(200, { ok: true, ticketId: ticket.ticketId });
    }
    return json(404, { error: "route" });
  };
}

function statefulAdapter(events: string[]) {
  const cells = new Map<string, unknown>();
  const adapter: WorkbookPlanExecutionAdapter & WorkbookPlanVerificationAdapter = {
    async executeFirstGroup(gate) {
      events.push("fresh");
      return gate.executeAfterFresh("fresh", async (group) => {
        events.push("mutation");
        for (const operation of group.operations) {
          cells.set(`${String(operation.args.sheet)}!${String(operation.args.address)}`, operation.args.values);
        }
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
      return JSON.stringify(cells.get("Sheet1!A1")) === JSON.stringify([["ok"]]);
    },
  };
  return { adapter, cells };
}

// ---------------------------------------------------------------------------
// Harness: the real chain with the taskpane.ts wiring
// ---------------------------------------------------------------------------
interface ActiveTurnState {
  turnOperationId: string | null;
  sessionId: string | null;
  trustedRequest: TrustedWorkbookRequest | null;
  turnStart: number | null;
  hasCancellation: boolean;
}

interface Harness {
  readonly rollout: boolean;
  readonly events: string[];
  readonly companion: SessionCompanion;
  readonly lifecycle: OfficeBridgeLifecycleController;
  readonly bridge: () => OfficeCompanionBridge;
  readonly coordinator: TaskPaneTurnCoordinator;
  readonly double: OfficeDouble;
  readonly docs: { current: string };
  readonly active: ActiveTurnState;
  readonly counts: { executeTool: number };
  readonly awaitSpy: { readonly mock: { readonly calls: readonly (readonly unknown[])[] } };
  readonly mutations: () => number;
  beginTurn(turnId: string, prompt: string): TaskPaneTurnHandle;
  prepare(handle: TaskPaneTurnHandle): Promise<TaskPaneTurnPreparation>;
  /** One server primitive ticket on the bridge session; returns the posted `result` re-serialized. */
  serverPrimitive(name: string, args: Record<string, unknown>): Promise<string>;
}

const cleanups: (() => Promise<void> | void)[] = [];
afterEach(async () => {
  while (cleanups.length > 0) {
    try { await cleanups.pop()!(); } catch { /* best-effort test teardown */ }
  }
  vi.restoreAllMocks();
  delete (globalThis as Record<string, unknown>).Excel;
});

async function harness(rollout: boolean): Promise<Harness> {
  const events: string[] = [];
  const companion = new SessionCompanion(events);
  const intervals: (() => void)[] = [];
  const docs = { current: DOC1 };
  const counts = { executeTool: 0 };
  const active: ActiveTurnState = { turnOperationId: null, sessionId: null, trustedRequest: null, turnStart: null, hasCancellation: false };
  const bindings = new ServerOwnedTurnBindingStore({ now: () => Date.now(), ttlMs: 90_000, maxEntries: 32, maxBytes: 64 * 1024 });
  let route!: ReturnType<typeof createServerOwnedPlanRoute>;
  let bridge: OfficeCompanionBridge | undefined;
  const lifecycle = new OfficeBridgeLifecycleController((emit) => {
    bridge = createExcelCompanionBridge(
      emit,
      (name, argumentsValue, trusted, control) => {
        counts.executeTool += 1;
        return route.executeCompanionTool(name, argumentsValue, trusted, control);
      },
      {
        fetchImpl: companion.fetch as typeof fetch,
        documentId: async () => docs.current,
        paneId: PANE,
        setIntervalImpl: ((handler: () => void) => { intervals.push(handler); return intervals.length; }) as unknown as typeof setInterval,
        clearIntervalImpl: (() => undefined) as unknown as typeof clearInterval,
      },
    );
    return bridge;
  }, (_state: BridgePublicState) => undefined);
  const coordinator = new TaskPaneTurnCoordinator({
    getOrCreateDocumentId: async () => docs.current,
    getBridgeIdentity: () => ({ sessionId: lifecycle.sessionId, documentId: lifecycle.documentId }),
    reconnectForDocumentChange: (expected) => lifecycle.reconnectForDocumentChange(expected),
    bindings,
    paneId: PANE,
  });
  // taskpane.ts: const serverPrimitiveGate = createServerPrimitiveGate({...}) (base: the R4c pass-through).
  const executePrimitive: PrimitiveExecutor = gateFactory !== undefined
    ? gateFactory({
      getBridgeSessionId: () => lifecycle.sessionId ?? null,
      getActiveTurn: () => ({ turnOperationId: active.turnOperationId, sessionId: active.sessionId }),
      resolveAuthority: (input: { readonly turnId: string; readonly bridgeSessionId: string }) => serverApi(coordinator).serverPrimitiveAuthority!(input),
      executeWithAuthority: executeValidatedExcelBridgeToolWithAuthority,
      executeReadOnly: executeValidatedExcelBridgeTool,
    })
    : basePassThrough();
  const { adapter } = statefulAdapter(events);
  const storage = new MemoryIdempotencyStorage();
  const awaitSpy = vi.fn((binding: WorkbookPlanConfirmationBinding, signal?: AbortSignal) =>
    coordinator.waitForWorkbookPlanConfirmation(binding, signal));
  route = createServerOwnedPlanRoute({
    isRolloutEnabled: () => rollout,
    getBridgeSessionId: () => lifecycle.sessionId ?? null,
    getActiveTurn: () => ({
      turnOperationId: active.turnOperationId,
      trustedRequest: active.trustedRequest,
      sessionId: active.sessionId,
      turnStart: active.turnStart,
      hasCancellation: active.hasCancellation,
    }),
    coordinator,
    bindings,
    readRuntimeContext: async () => { events.push("runtime-read"); return RUNTIME; },
    createAdapter: () => adapter,
    createStore: () => new BrowserIdempotencyStore(storage),
    inspect: async () => { events.push("inspect-read"); return "{}"; },
    awaitPaneConfirmation: awaitSpy,
    executePrimitive,
  } as unknown as Parameters<typeof createServerOwnedPlanRoute>[0]);
  const double = createOfficeDouble();
  double.install();
  await lifecycle.start();
  cleanups.push(async () => {
    coordinator.stopCurrentTurn();
    await lifecycle.stop();
  });
  const h: Harness = {
    rollout, events, companion, lifecycle, bridge: () => bridge!, coordinator, double, docs, active, counts, awaitSpy,
    mutations: () => events.filter((event) => event === "mutation").length,
    beginTurn(turnId, prompt) {
      // sendPrompt order (taskpane.ts:464-481): identity first, then beginTurn.
      const controller = new AbortController();
      const trustedRequest = classified(prompt);
      const turnStart = Date.now();
      active.turnOperationId = turnId;
      active.turnStart = turnStart;
      active.hasCancellation = true;
      active.sessionId = lifecycle.sessionId;
      active.trustedRequest = trustedRequest;
      return coordinator.beginTurn(turnId, controller, {
        trustedRequest,
        turnStart,
        cancellationSignal: controller.signal,
        workbookPlanRolloutEnabled: rollout,
      });
    },
    async prepare(handle) {
      // continueTurn/prepareTurn (taskpane.ts:636-641): activeTurnSessionId = preparation.sessionId.
      const preparation = await coordinator.prepareTurn(handle);
      if (preparation.status === "active" && active.turnOperationId === handle.turnId) active.sessionId = preparation.sessionId;
      return preparation;
    },
    async serverPrimitive(name, args) {
      const invoked = companion.invokePrimitive(lifecycle.sessionId!, name, args);
      expect(invoked.status).toBe(202);
      expect(await bridge!.pollOnce()).toBe(true);
      const posted = companion.resultsFor(invoked.ticketId);
      expect(posted).toHaveLength(1);
      return JSON.stringify(posted[0].result);
    },
  };
  return h;
}

/** The base validated bridge path on a fresh seeded double, re-serialized as the bridge posts it. */
async function validatedOracle(calls: ReadonlyArray<readonly [string, Record<string, unknown>]>, configure?: (double: OfficeDouble) => void): Promise<{ payloads: string[]; log: string[] }> {
  const double = createOfficeDouble();
  configure?.(double);
  double.install();
  const payloads: string[] = [];
  for (const [name, args] of calls) {
    const validation = validateToolCall(name, JSON.stringify(args));
    const raw = !("valid" in validation) || validation.valid !== true
      ? JSON.stringify(validation)
      : await getExcelToolDefinition(name)!.execute(validation.parsedArgs);
    payloads.push(reserialize(raw));
  }
  delete (globalThis as Record<string, unknown>).Excel;
  return { payloads, log: [...double.log] };
}

function reserialize(raw: string): string {
  try {
    return JSON.stringify(JSON.parse(raw));
  } catch {
    return JSON.stringify(raw);
  }
}

async function flushUntil(condition: () => boolean): Promise<void> {
  for (let attempt = 0; attempt < 400 && !condition(); attempt += 1) {
    await new Promise<void>((done) => setImmediate(done));
  }
  expect(condition()).toBe(true);
}

async function settle<T>(promise: Promise<T>, stop: () => boolean): Promise<{ settled: boolean; value?: T | string }> {
  let state: { settled: boolean; value?: T | string } = { settled: false };
  promise.then((value) => { state = { settled: true, value }; }, (error: unknown) => { state = { settled: true, value: `rejected:${String(error)}` }; });
  for (let attempt = 0; attempt < 400 && !state.settled && !stop(); attempt += 1) {
    await new Promise<void>((done) => setImmediate(done));
  }
  return state;
}

const oneCell = (row: number): readonly [string, Record<string, unknown>] => ["write_range", { sheet: "Sheet1", address: `A${row}`, values: [[row]] }];
const WRITE = oneCell(1);
const READ: readonly [string, Record<string, unknown>] = ["read_range", { sheet: "Sheet1", address: "A1:B2" }];

async function liveTurn(h: Harness, turnId: string, prompt: string): Promise<TaskPaneTurnHandle> {
  const handle = h.beginTurn(turnId, prompt);
  if (h.rollout) expect((await h.prepare(handle)).status).toBe("active");
  return handle;
}

// ---------------------------------------------------------------------------
// R4d surface presence
// ---------------------------------------------------------------------------
describe("R4d server primitive gate surface", () => {
  it("exports createServerPrimitiveGate from the authority module and serverPrimitiveAuthority on the coordinator", () => {
    expect(gateFactory).toBeTypeOf("function");
    expect(typeof (TaskPaneTurnCoordinator.prototype as unknown as ServerAuthorityApi).serverPrimitiveAuthority).toBe("function");
    expect((officeBridgeExcelModule as unknown as Record<string, unknown>).executeLegacyServerPrimitiveR4dPending).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// §5.1 S-rows: no live turn on the current session (OQ-D1: both modes)
// ---------------------------------------------------------------------------
describe("S-rows: server primitives need a live pane turn on the current session (OQ-D1 ALWAYS)", () => {
  it.each([true, false])("S1 / S1-off no pane turn (rollout=%s): write_range is TURN_INACTIVE with zero Excel.run", async (rollout) => {
    const h = await harness(rollout);
    expect(await h.serverPrimitive(...WRITE)).toBe(DENIAL.TURN_INACTIVE);
    expect(excelRuns(h.double.log)).toBe(0);
  });

  it.each([
    [true, "activate_sheet", { name: "Sheet1" }],
    [true, "delete_sheet", { name: "Old" }],
    [true, "format_range", { sheet: "Sheet1", address: "A1:B2", format: { bold: true } }],
    [false, "activate_sheet", { name: "Sheet1" }],
    [false, "delete_sheet", { name: "Old" }],
    [false, "format_range", { sheet: "Sheet1", address: "A1:B2", format: { bold: true } }],
  ] as const)("S1v no pane turn (rollout=%s): %s is TURN_INACTIVE with zero Excel.run", async (rollout, name, args) => {
    const h = await harness(rollout);
    expect(await h.serverPrimitive(name, { ...args })).toBe(DENIAL.TURN_INACTIVE);
    expect(excelRuns(h.double.log)).toBe(0);
  });

  it.each([true, false])("S3 / S3-off turn begun on S1, bridge rotated to S2 (rollout=%s): TURN_INACTIVE", async (rollout) => {
    const h = await harness(rollout);
    const s1 = h.lifecycle.sessionId;
    await liveTurn(h, T1, PROMPTS.fill);
    expect(h.active.sessionId).toBe(s1);
    h.docs.current = DOC2;
    await h.lifecycle.reconnectForDocumentChange(DOC2);
    expect(h.lifecycle.sessionId).not.toBe(s1);
    expect(h.active.turnOperationId).toBe(T1);
    expect(await h.serverPrimitive(...WRITE)).toBe(DENIAL.TURN_INACTIVE);
    expect(excelRuns(h.double.log)).toBe(0);
  });

  it.each([
    [true, "finishTurn"], [true, "stopCurrentTurn"], [false, "finishTurn"], [false, "stopCurrentTurn"],
  ] as const)("S4 / S4-off (rollout=%s) %s, then write_range: TURN_INACTIVE", async (rollout, end) => {
    const h = await harness(rollout);
    const handle = await liveTurn(h, T1, PROMPTS.fill);
    if (end === "finishTurn") h.coordinator.finishTurn(handle);
    else h.coordinator.stopCurrentTurn();
    // The pane-side identity is left in place: the coordinator must refuse on its own.
    expect(h.active.turnOperationId).toBe(T1);
    expect(await h.serverPrimitive(...WRITE)).toBe(DENIAL.TURN_INACTIVE);
    expect(excelRuns(h.double.log)).toBe(0);
  });

  it("S5 rollout on: T1 ends unclean, T2 begun but not prepared; a write_range claimed on S1 is TURN_INACTIVE", async () => {
    const h = await harness(true);
    const s1 = h.lifecycle.sessionId!;
    const t1 = await liveTurn(h, T1, PROMPTS.fill);
    h.coordinator.beginHermesRequest(s1);
    h.coordinator.finishTurn(t1);
    h.beginTurn(T2, PROMPTS.fill);
    expect(h.active.sessionId).toBe(s1);
    expect(h.lifecycle.sessionId).toBe(s1);
    expect(await h.serverPrimitive(...WRITE)).toBe(DENIAL.TURN_INACTIVE);
    expect(excelRuns(h.double.log)).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// §5.1 S6: a ticket claimed on S1 and read after T2 rotated to S2 is dropped
// ---------------------------------------------------------------------------
describe("S6: a stale S1 ticket never binds to T2 (R3b lane-B A1)", () => {
  async function staleAfterRotation(kind: "confirm" | "explicit" | "primitive") {
    const h = await harness(true);
    const s1 = h.lifecycle.sessionId!;
    const t1 = await liveTurn(h, T1, PROMPTS.writeOk);
    h.coordinator.beginHermesRequest(s1);
    h.coordinator.finishTurn(t1);
    const ticket = kind === "primitive"
      ? h.companion.invokePrimitive(s1, ...WRITE)
      : h.companion.invoke(s1, H1_OP, EXECUTE_PLAN);
    expect(ticket.status).toBe(202);
    const release = h.companion.holdNextPending();
    const polled = h.bridge().pollOnce();
    cleanups.push(async () => { h.coordinator.stopCurrentTurn(); await polled.catch(() => undefined); });
    await flushUntil(() => h.companion.heldPending === 1);
    const t2 = h.beginTurn(T2, kind === "confirm" ? PROMPTS.usual : PROMPTS.writeOk);
    const p2 = await h.prepare(t2);
    expect(p2.status).toBe("active");
    const s2 = h.lifecycle.sessionId!;
    expect(s2).not.toBe(s1);
    expect(h.companion.deleted).toEqual([s1]);
    expect(h.active.sessionId).toBe(s2);
    const executeToolBefore = h.counts.executeTool;
    release();
    const outcome = await settle(polled, () => h.awaitSpy.mock.calls.length > 0 || h.events.includes("runtime-read"));
    expect(outcome).toEqual({ settled: true, value: false });
    expect(h.counts.executeTool - executeToolBefore).toBe(0);
    expect(h.companion.resultsFor(ticket.ticketId)).toEqual([]);
    expect(h.companion.results).toEqual([]);
    expect(h.events.filter((event) => event === "runtime-read" || event === "fresh")).toEqual([]);
    expect(h.awaitSpy.mock.calls).toHaveLength(0);
    expect(excelRuns(h.double.log)).toBe(0);
    return { h, t2, s2 };
  }

  it("S6 confirm variant: the held S1 plan ticket is dropped; T2's own confirmed plan commits once", async () => {
    const { h, s2 } = await staleAfterRotation("confirm");
    const own = h.companion.invoke(s2, T2_OP, EXECUTE_PLAN);
    const polled = h.bridge().pollOnce();
    cleanups.push(async () => { h.coordinator.stopCurrentTurn(); await polled.catch(() => undefined); });
    await flushUntil(() => h.awaitSpy.mock.calls.length > 0);
    expect(h.coordinator.continuePendingWorkbookPlan()).toBe(true);
    await expect(polled).resolves.toBe(true);
    expect(h.mutations()).toBe(1);
    expect(h.companion.resultsFor(own.ticketId)).toHaveLength(1);
    expect(h.companion.resultsFor(own.ticketId)[0]).toMatchObject({ executed: true, result: { ok: true } });
  });

  it("S6e explicit variant: the held S1 plan ticket is dropped; T2's own plan commits once", async () => {
    const { h, s2 } = await staleAfterRotation("explicit");
    const own = h.companion.invoke(s2, T2_OP, EXECUTE_PLAN);
    await expect(h.bridge().pollOnce()).resolves.toBe(true);
    expect(h.awaitSpy.mock.calls).toHaveLength(0);
    expect(h.mutations()).toBe(1);
    expect(h.companion.resultsFor(own.ticketId)[0]).toMatchObject({ executed: true, result: { ok: true } });
  });

  it("S6p primitive variant: the held S1 write_range is dropped with zero Excel.run and no post", async () => {
    const { h } = await staleAfterRotation("primitive");
    expect(formulaWrites(h.double.log)).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// §5.1 S7-S13: a prepared rollout-on turn applies the R4c rules to server tickets
// ---------------------------------------------------------------------------
describe("S-rows: rollout-on server primitives follow the turn's R4c authority", () => {
  it("S7 fill the sheet: 31 distinct one-cell writes, 1-30 byte-identical with one readback each, #31 RESOURCE_LIMIT", async () => {
    const calls = Array.from({ length: 31 }, (_v, i) => oneCell(i + 1));
    const reference = await validatedOracle(calls.slice(0, 30));
    const h = await harness(true);
    h.double.install();
    await liveTurn(h, T1, PROMPTS.fill);
    const posted: string[] = [];
    for (const [name, args] of calls) posted.push(await h.serverPrimitive(name, args));
    expect(posted.slice(0, 30)).toEqual(reference.payloads);
    expect(posted[30]).toBe(DENIAL.RESOURCE_LIMIT);
    const segments = runSegments(h.double.log);
    expect(segments).toHaveLength(60);
    expect(segments.filter((_segment, i) => i % 2 === 0).flat()).toEqual(reference.log);
    segments.filter((_segment, i) => i % 2 === 1).forEach((segment) => {
      expect(segment.some((entry) => entry.startsWith("set "))).toBe(false);
      expect(segment.some((entry) => entry.startsWith("load range("))).toBe(true);
    });
  });

  it("S8 identical write_range twice: the 2nd is DUPLICATE, one write", async () => {
    const h = await harness(true);
    await liveTurn(h, T1, PROMPTS.fill);
    expect(JSON.parse(await h.serverPrimitive(...WRITE)).ok).toBe(true);
    expect(await h.serverPrimitive(...WRITE)).toBe(DENIAL.DUPLICATE);
    expect(formulaWrites(h.double.log)).toHaveLength(1);
  });

  it("S9 borra la hoja: write_range is CONFIRMATION_REQUIRED (names excel_execute_plan), zero Excel.run", async () => {
    const h = await harness(true);
    await liveTurn(h, T1, PROMPTS.borra);
    const posted = await h.serverPrimitive(...WRITE);
    expect(posted).toBe(DENIAL.CONFIRMATION_REQUIRED);
    expect(posted).toContain("excel_execute_plan");
    expect(excelRuns(h.double.log)).toBe(0);
  });

  it("S9b show me the table: write_range CONFIRMATION_REQUIRED, run_calculation allowed", async () => {
    const reference = await validatedOracle([["run_calculation", {}]]);
    const h = await harness(true);
    h.double.install();
    await liveTurn(h, T1, PROMPTS.showTable);
    expect(await h.serverPrimitive(...WRITE)).toBe(DENIAL.CONFIRMATION_REQUIRED);
    expect(await h.serverPrimitive("run_calculation", {})).toBe(reference.payloads[0]);
    expect(h.double.log).toEqual(reference.log);
  });

  it("S10 create a table: delete_sheet, protect_sheet, create_pivot_table are each LEGACY_NO_PLAN_PATH, zero Excel.run", async () => {
    const h = await harness(true);
    await liveTurn(h, T1, PROMPTS.createTable);
    expect(await h.serverPrimitive("delete_sheet", { name: "Old" })).toBe(DENIAL.LEGACY_NO_PLAN_PATH);
    expect(await h.serverPrimitive("protect_sheet", { sheet: "Sheet1" })).toBe(DENIAL.LEGACY_NO_PLAN_PATH);
    expect(await h.serverPrimitive("create_pivot_table", {
      sheet: "Sheet1", address: "H1", sourceSheet: "Sheet1", sourceRange: "A1:C4", rowFields: ["a"], dataFields: [{ name: "b" }],
    })).toBe(DENIAL.LEGACY_NO_PLAN_PATH);
    expect(excelRuns(h.double.log)).toBe(0);
  });

  it("S11 an authorized server plan (ask) latches the turn: later write_range PLAN_ENTERED, activate_sheet allowed", async () => {
    const reference = await validatedOracle([["activate_sheet", { name: "Sheet1" }]]);
    const h = await harness(true);
    h.double.install();
    await liveTurn(h, T1, PROMPTS.fill);
    const plan = h.companion.invoke(h.lifecycle.sessionId!, H1_OP, ASK_PLAN);
    await expect(h.bridge().pollOnce()).resolves.toBe(true);
    expect(h.companion.resultsFor(plan.ticketId)).toHaveLength(1);
    expect(h.companion.resultsFor(plan.ticketId)[0]).toMatchObject({ executed: false });
    expect(await h.serverPrimitive(...WRITE)).toBe(DENIAL.PLAN_ENTERED);
    expect(await h.serverPrimitive("activate_sheet", { name: "Sheet1" })).toBe(reference.payloads[0]);
    expect(h.double.log).toEqual(reference.log);
  });

  it("S12 one budget per turn: 20 direct writes, then server writes 1-10 allowed and the 11th RESOURCE_LIMIT", async () => {
    const h = await harness(true);
    const handle = await liveTurn(h, T1, PROMPTS.fill);
    const authority = h.coordinator.primitiveAuthorityFor(handle);
    expect(authority).not.toBeNull();
    for (let row = 1; row <= 20; row += 1) {
      expect(JSON.parse(await authority!.executePrimitive(...oneCell(row))).ok).toBe(true);
    }
    const posted: string[] = [];
    for (let row = 21; row <= 31; row += 1) posted.push(await h.serverPrimitive(...oneCell(row)));
    expect(posted.slice(0, 10).every((payload) => JSON.parse(payload).ok === true)).toBe(true);
    expect(posted[10]).toBe(DENIAL.RESOURCE_LIMIT);
    expect(formulaWrites(h.double.log)).toHaveLength(30);
  });

  it("S13 A2 on the server route: a host-failed write posts the base ok:false payload, its identical retry executes and verifies; one op charged", async () => {
    const failed = await validatedOracle([WRITE], (double) => { double.failRuns = true; });
    const healthy = await validatedOracle([WRITE]);
    expect(JSON.parse(failed.payloads[0]).ok).toBe(false);
    const h = await harness(true);
    h.double.install();
    const handle = await liveTurn(h, T1, PROMPTS.fill);
    h.double.failRuns = true;
    expect(await h.serverPrimitive(...WRITE)).toBe(failed.payloads[0]);
    h.double.failRuns = false;
    const runsBefore = excelRuns(h.double.log);
    expect(await h.serverPrimitive(...WRITE)).toBe(healthy.payloads[0]);
    expect(excelRuns(h.double.log) - runsBefore).toBe(2);
    expect(h.coordinator.primitiveAuthorityFor(handle)?.usageSnapshot().operationCount).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// §5.1 SX: the explicit OQ-D1 server matrix
// ---------------------------------------------------------------------------
describe("SX OQ-D1 server matrix", () => {
  const rows = ([false, true] as const).flatMap((rollout) => (["none", "live"] as const).flatMap((turn) =>
    ([WRITE, READ] as const).map((call) => ({ rollout, turn, name: call[0], args: call[1] }))));
  it.each(rows)("SX rollout=$rollout turn=$turn $name", async ({ rollout, turn, name, args }) => {
    const reference = await validatedOracle([[name, args]]);
    const h = await harness(rollout);
    h.double.install();
    if (turn === "live") await liveTurn(h, T1, PROMPTS.fill);
    const posted = await h.serverPrimitive(name, args);
    if (turn === "none" && name === "write_range") {
      expect(posted).toBe(DENIAL.TURN_INACTIVE);
      expect(excelRuns(h.double.log)).toBe(0);
      return;
    }
    expect(posted).toBe(reference.payloads[0]);
    if (name === "write_range" && rollout) {
      const segments = runSegments(h.double.log);
      expect(segments).toHaveLength(2);
      expect(segments[0]).toEqual(reference.log);
    } else {
      expect(h.double.log).toEqual(reference.log);
    }
  });
});

// ---------------------------------------------------------------------------
// §5.2 SP-rows (GREEN at the base and at HEAD)
// ---------------------------------------------------------------------------
describe("SP-rows: parity pins", () => {
  it("SP1 rollout off, live turn on the current session: every posted body and the Office log equal the base validated path", async () => {
    const calls: (readonly [string, Record<string, unknown>])[] = [
      ...Array.from({ length: 31 }, (_v, i) => oneCell(i + 1)),
      oneCell(40), oneCell(40),
      ["delete_sheet", { name: "Old" }], ["delete_sheet", { name: "Old" }],
      ["write_range", { sheet: "Sheet1", address: "MyNamedRange", values: [[1]] }],
      ["write_range", { sheet: "Sheet1", address: "A:A", values: [[1]] }],
    ];
    const reference = await validatedOracle(calls);
    const h = await harness(false);
    h.double.install();
    await liveTurn(h, T1, PROMPTS.borra);
    const posted: string[] = [];
    for (const [name, args] of calls) posted.push(await h.serverPrimitive(name, args));
    expect(posted).toEqual(reference.payloads);
    expect(h.double.log).toEqual(reference.log);
  });

  it.each([true, false])("SP2 no turn (rollout=%s): reads are byte-identical and an unknown name gets the base validation payload", async (rollout) => {
    const calls: (readonly [string, Record<string, unknown>])[] = [
      READ,
      ["summarize_range_stats", { sheet: "Sheet1", address: "A1:B2" }],
      ["get_workbook_overview", {}],
      ["not_a_tool_r4d", { sheet: "Sheet1" }],
    ];
    const reference = await validatedOracle(calls);
    expect(JSON.parse(reference.payloads[3]).ok).toBe(false);
    const h = await harness(rollout);
    h.double.install();
    const posted: string[] = [];
    for (const [name, args] of calls) posted.push(await h.serverPrimitive(name, args));
    expect(posted).toEqual(reference.payloads);
    expect(h.double.log).toEqual(reference.log);
  });

  it.each([PROMPTS.fill, PROMPTS.borra, PROMPTS.summarize, PROMPTS.createTable, PROMPTS.selected])(
    "SP3 live rollout-on turn (%s): a server read_range is byte-identical and uncharged",
    async (prompt) => {
      if (gateFactory === undefined) return; // HEAD only (guarded)
      const reference = await validatedOracle([READ]);
      const h = await harness(true);
      h.double.install();
      const handle = await liveTurn(h, T1, prompt);
      expect(await h.serverPrimitive(...READ)).toBe(reference.payloads[0]);
      expect(h.double.log).toEqual(reference.log);
      expect(h.coordinator.primitiveAuthorityFor(handle)?.usageSnapshot()).toEqual(ZERO_USAGE);
    },
  );

  it("F1 (lane A P3-A1) the turn's authority and its plan-entered latch survive prepareTurn / continueTurn", async () => {
    const reference = await validatedOracle([["activate_sheet", { name: "Sheet1" }]]);
    const h = await harness(true);
    h.double.install();
    const handle = h.beginTurn(T1, PROMPTS.fill);
    const a = h.coordinator.primitiveAuthorityFor(handle);
    expect(a).not.toBeNull();
    h.coordinator.markWorkbookPlanEntered(handle);
    const continued = await h.coordinator.continueTurn(handle, (preparation) => {
      h.active.sessionId = preparation.sessionId;
      return h.coordinator.primitiveAuthorityFor(handle);
    });
    expect(continued.status).toBe("active");
    expect(continued.status === "active" && "value" in continued ? continued.value : undefined).toBe(a);
    expect(h.coordinator.primitiveAuthorityFor(handle)).toBe(a);
    expect(await a!.executePrimitive(...WRITE)).toBe(DENIAL.PLAN_ENTERED);
    expect(reserialize(await a!.executePrimitive("activate_sheet", { name: "Sheet1" }))).toBe(reference.payloads[0]);
    if (gateFactory !== undefined) {
      expect(serverApi(h.coordinator).serverPrimitiveAuthority!({ turnId: T1, bridgeSessionId: h.lifecycle.sessionId! })).toBe(a);
    }

    const off = await harness(false);
    const offHandle = off.beginTurn(T2, PROMPTS.fill);
    const b = off.coordinator.primitiveAuthorityFor(offHandle);
    expect(b).not.toBeNull();
    const offContinued = await off.coordinator.continueTurn(offHandle, () => off.coordinator.primitiveAuthorityFor(offHandle));
    expect(offContinued.status === "active" && "value" in offContinued ? offContinued.value : undefined).toBe(b);
    expect(off.coordinator.primitiveAuthorityFor(offHandle)).toBe(b);
  });

  it("G2 taskpane.ts builds the route's executePrimitive with createServerPrimitiveGate (guarded on gate presence)", () => {
    if (gateFactory === undefined) return; // HEAD only (guarded)
    const source = readFileSync(resolve(__dirname, "../taskpane/taskpane.ts"), "utf8");
    expect(source.split("createServerPrimitiveGate({")).toHaveLength(2);
    expect(source).toContain("executePrimitive: serverPrimitiveGate,");
    expect(source).toContain("executeWithAuthority: executeValidatedExcelBridgeToolWithAuthority,");
    expect(source).toContain("executeReadOnly: executeValidatedExcelBridgeTool,");
    expect(source).toContain("resolveAuthority: (input) => turnCoordinator.serverPrimitiveAuthority(input),");
  });
});
