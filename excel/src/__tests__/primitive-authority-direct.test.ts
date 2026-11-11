/**
 * R4c direct-route tests: the real HermesAgent.run loop with an injected
 * model double (globalThis.fetch, SSE) and the real TaskPaneTurnCoordinator,
 * whose per-turn authority is passed in exactly as taskpane.ts does
 * (dev3-r4c-release.md §5.1 D1-D12 with rollout ON, §5.2 P1-P9 rollout-off
 * parity, G1, G4).
 *
 * P-rows and G-rows hold at the base (no authority: the agent calls
 * executeExcelToolParsed) and at HEAD (rollout-off authority) with identical
 * assertions. The tool-message contents are read from the model double's last
 * request body: AgentRunResult.messages carries only user/assistant history.
 */
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import { HermesAgent, type AgentCallbacks } from "../agent";
import { getExcelToolDefinition, getWorkbookContext } from "../excelTools";
import { classifyWorkbookRequest, type TrustedWorkbookRequest } from "../workbookPlanRequest";
import {
  TaskPaneTurnCoordinator,
  type TaskPaneTurnHandle,
  type TaskPaneTurnStartContext,
} from "../taskpane/taskPaneTurnCoordinator";
import { ServerOwnedTurnBindingStore } from "../workbookPlanOrchestrator";
import type { ChatMessage } from "../types";

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

// ---------------------------------------------------------------------------
// Fixed expectations
// ---------------------------------------------------------------------------
const deny = (error: string, category: string): string => JSON.stringify({ ok: false, error, category });
const DENIAL = {
  READ_ONLY_SINK: deny("Workbook edits must go through the task-pane edit authority.", "validation"),
  TURN_INACTIVE: deny("Workbook edit denied: no active task-pane turn.", "validation"),
  LEGACY_NO_PLAN_PATH: deny("Workbook edit denied: this legacy tool has no WorkbookPlan path and is disabled while WorkbookPlan rollout is on.", "validation"),
  CONFIRMATION_REQUIRED: deny("Workbook edit denied: this request needs confirmation. Use excel_execute_plan (WorkbookPlan) instead.", "validation"),
  DUPLICATE: deny("Workbook edit denied: the identical edit was already applied in this turn.", "repetition"),
  RESOURCE_LIMIT: deny("Workbook edit denied: the per-turn workbook edit limit was reached.", "validation"),
  READBACK_FAILED: deny("Workbook edit denied: an earlier edit in this turn did not verify.", "validation"),
} as const;
const REPEAT_BLOCKED = JSON.stringify({
  ok: false,
  error: "Repeated identical tool call blocked; use prior result or ask for clarification.",
  category: "repetition",
});

const PROMPTS = {
  fill: "fill the sheet with numbers",
  borra: "borra la hoja",
  summarize: "summarize the totals",
  createTable: "create a table from this data",
  selected: "update the selected cells",
  showTable: "show me the table",
  dontChange: "don't change anything, just summarize the table",
  writeOk: "write ok into A1",
} as const;

const EXPECTED_CLASS: Readonly<Record<string, readonly [string, boolean]>> = {
  [PROMPTS.fill]: ["explicit_mutation", false],
  [PROMPTS.borra]: ["ambiguous_mutation", false],
  [PROMPTS.summarize]: ["read_only", false],
  [PROMPTS.createTable]: ["explicit_structure", false],
  [PROMPTS.selected]: ["explicit_mutation", true],
  [PROMPTS.showTable]: ["read_only", false],
  [PROMPTS.dontChange]: ["ambiguous_mutation", false],
  [PROMPTS.writeOk]: ["explicit_mutation", false],
};

function classified(prompt: string): TrustedWorkbookRequest {
  const request = classifyWorkbookRequest(prompt);
  expect([request.classification, request.selectionScopedRequest]).toEqual(EXPECTED_CLASS[prompt]);
  return request;
}

// ---------------------------------------------------------------------------
// Model double: globalThis.fetch returning SSE, <= 10 tool_calls per response
// ---------------------------------------------------------------------------
type Call = readonly [string, Record<string, unknown>];
type ModelResponse = readonly Call[] | string;

function sseResponse(response: ModelResponse, responseIndex: number): Response {
  const encoder = new TextEncoder();
  const delta = typeof response === "string"
    ? { content: response }
    : {
        tool_calls: response.map(([name, args], index) => ({
          index,
          id: `call_${responseIndex}_${index}`,
          type: "function",
          function: { name, arguments: JSON.stringify(args) },
        })),
      };
  const body = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(encoder.encode(`data: ${JSON.stringify({ choices: [{ delta }] })}\n\n`));
      controller.enqueue(encoder.encode("data: [DONE]\n\n"));
      controller.close();
    },
  });
  return { ok: true, status: 200, headers: new Headers({ "content-type": "text/event-stream" }), body, text: async () => "" } as unknown as Response;
}

// create_pivot_table names its pivot from Date.now(): pin the clock so the
// agent run and its oracle replay are deterministic.
const FIXED_NOW = 1_800_000_000_000;
const previousFetch = globalThis.fetch;
beforeEach(() => {
  vi.spyOn(Date, "now").mockReturnValue(FIXED_NOW);
});
afterEach(() => {
  vi.restoreAllMocks();
  globalThis.fetch = previousFetch;
  delete (globalThis as Record<string, unknown>).Excel;
});

interface AuthorityLike {
  readonly isTerminal: boolean;
  readonly rolloutEnabled: boolean;
  usageSnapshot(): Readonly<Record<string, number>>;
}

const AUTHORITY_MODULE_PATH = "../primitiveMutationAuthority";
let authorityModulePresent = false;
beforeAll(async () => {
  authorityModulePresent = (await import(/* @vite-ignore */ AUTHORITY_MODULE_PATH).catch(() => undefined)) !== undefined;
});

const ZERO_USAGE = {
  operationCount: 0, cellsRead: 0, cellsWritten: 0, formulaUtf8Bytes: 0, csvPayloadUtf8Bytes: 0,
  commentUtf8Bytes: 0, sheetsCreated: 0, chartsCreated: 0, tablesCreated: 0, verificationReads: 0,
};

interface DirectRun {
  readonly toolContents: string[];
  readonly double: OfficeDouble;
  /** Office log after the initial workbook overview, i.e. the primitive phase only. */
  readonly primitiveLog: string[];
  readonly requests: number;
  readonly finalContent: string;
  readonly authority: AuthorityLike | undefined;
  readonly coordinator: TaskPaneTurnCoordinator;
  readonly handle: TaskPaneTurnHandle;
}

interface DirectOptions {
  readonly prompt: string;
  readonly rollout: boolean;
  readonly responses: readonly ModelResponse[];
  readonly withAuthority?: boolean;
  readonly configure?: (double: OfficeDouble) => void;
  readonly onRequest?: (index: number, run: { coordinator: TaskPaneTurnCoordinator; handle: TaskPaneTurnHandle; double: OfficeDouble }) => void;
  readonly onWorkbookPlanTurn?: AgentCallbacks["onWorkbookPlanTurn"];
}

async function runDirect(options: DirectOptions): Promise<DirectRun> {
  for (const response of options.responses) {
    if (typeof response !== "string") expect(response.length).toBeLessThanOrEqual(10);
  }
  const double = createOfficeDouble();
  options.configure?.(double);
  double.install();
  const coordinator = new TaskPaneTurnCoordinator({
    getOrCreateDocumentId: async () => "11111111-1111-4111-8111-111111111111",
    getBridgeIdentity: () => ({ sessionId: null, documentId: null }),
    reconnectForDocumentChange: async () => { throw new Error("unused"); },
    bindings: new ServerOwnedTurnBindingStore({ now: () => Date.now(), ttlMs: 90_000, maxEntries: 32, maxBytes: 64 * 1024 }),
    paneId: "22222222-2222-4222-8222-222222222222",
  });
  const controller = new AbortController();
  const settings = { baseUrl: "http://127.0.0.1:9/v1", model: "test-direct-model", apiKey: "", maxIterations: 10, workbookPlanRolloutEnabled: options.rollout };
  const context: TaskPaneTurnStartContext & { workbookPlanRolloutEnabled?: boolean } = {
    trustedRequest: classifyWorkbookRequest(options.prompt),
    turnStart: Date.now(),
    cancellationSignal: controller.signal,
    workbookPlanRolloutEnabled: settings.workbookPlanRolloutEnabled === true,
  };
  const handle = coordinator.beginTurn("turn-direct", controller, context);
  const authority = options.withAuthority === false
    ? undefined
    : ((coordinator as unknown as { primitiveAuthorityFor?: (h: TaskPaneTurnHandle) => unknown }).primitiveAuthorityFor?.(handle) ?? undefined);
  const bodies: Array<{ messages: ChatMessage[] }> = [];
  let overviewEnd = -1;
  globalThis.fetch = (async (_url: unknown, init?: { body?: unknown }) => {
    const index = bodies.length;
    if (overviewEnd < 0) overviewEnd = double.log.length;
    bodies.push(JSON.parse(String(init?.body)) as { messages: ChatMessage[] });
    options.onRequest?.(index, { coordinator, handle, double });
    const response = options.responses[index] ?? "done";
    return sseResponse(response, index);
  }) as typeof fetch;
  const callbacks = {
    onAssistantStart: () => undefined,
    onAssistantDelta: () => undefined,
    onAssistantFinal: () => undefined,
    onActivity: () => undefined,
    onStatus: () => undefined,
    ...(options.onWorkbookPlanTurn === undefined ? {} : { onWorkbookPlanTurn: options.onWorkbookPlanTurn }),
    ...(authority === undefined ? {} : { primitiveAuthority: authority }),
  } as AgentCallbacks & Record<string, unknown>;
  const result = await new HermesAgent().run(options.prompt, settings, callbacks, controller.signal);
  const last = bodies[bodies.length - 1];
  const toolContents = (last?.messages ?? [])
    .filter((message) => message.role === "tool")
    .map((message) => String(message.content));
  return {
    toolContents,
    double,
    primitiveLog: double.log.slice(Math.max(overviewEnd, 0)),
    requests: bodies.length,
    finalContent: result.finalContent,
    authority: authority as AuthorityLike | undefined,
    coordinator,
    handle,
  };
}

function canonical(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonical);
  if (value !== null && typeof value === "object") {
    return Object.fromEntries(Object.entries(value).sort(([a], [b]) => a.localeCompare(b)).map(([k, v]) => [k, canonical(v)]));
  }
  return value;
}

/** Base-semantics oracle: overview, then each call through the raw definition, with the agent's repeat guard. */
async function directOracle(calls: readonly Call[], configure?: (double: OfficeDouble) => void): Promise<{ payloads: string[]; log: string[] }> {
  const double = createOfficeDouble();
  configure?.(double);
  double.install();
  await getWorkbookContext();
  const counts = new Map<string, number>();
  const payloads: string[] = [];
  for (const [name, args] of calls) {
    const signature = `${name}:${JSON.stringify(canonical(args))}`;
    const count = (counts.get(signature) ?? 0) + 1;
    counts.set(signature, count);
    payloads.push(count > 2 ? REPEAT_BLOCKED : await getExcelToolDefinition(name)!.execute(args));
  }
  return { payloads, log: [...double.log] };
}

/** Raw payloads of `calls` in order on a fresh double (no overview). */
async function rawPayloads(calls: readonly Call[]): Promise<string[]> {
  const double = createOfficeDouble();
  double.install();
  const payloads: string[] = [];
  for (const [name, args] of calls) payloads.push(await getExcelToolDefinition(name)!.execute(args));
  return payloads;
}

function chunks<T>(items: readonly T[], sizes: readonly number[]): T[][] {
  const out: T[][] = [];
  let at = 0;
  for (const size of sizes) {
    out.push(items.slice(at, at + size));
    at += size;
  }
  expect(at).toBe(items.length);
  return out;
}

function runSegments(log: readonly string[]): string[][] {
  const starts = log.map((entry, index) => (entry === "Excel.run" ? index : -1)).filter((index) => index >= 0);
  return starts.map((start, i) => log.slice(start, starts[i + 1] ?? log.length));
}

const oneCell = (row: number): Call => ["write_range", { sheet: "Sheet1", address: `A${row}`, values: [[row]] }];
const grid = (rows: number, columns: number, value: number): number[][] =>
  Array.from({ length: rows }, () => Array.from({ length: columns }, () => value));

// ---------------------------------------------------------------------------
// §5.1 D-rows — rollout ON (RED at the base)
// ---------------------------------------------------------------------------
describe("D-rows: direct-route primitives with rollout ON", () => {
  it("D1 31 distinct one-cell writes: calls 1-30 allowed and verified, call 31 RESOURCE_LIMIT", async () => {
    classified(PROMPTS.fill);
    const calls = Array.from({ length: 31 }, (_v, i) => oneCell(i + 1));
    const run = await runDirect({ prompt: PROMPTS.fill, rollout: true, responses: [...chunks(calls, [10, 10, 10, 1]), "done"] });
    expect(run.toolContents).toHaveLength(31);
    expect(run.toolContents.slice(0, 30)).toEqual(await rawPayloads(calls.slice(0, 30)));
    expect(run.toolContents.findIndex((content) => content.startsWith("{\"ok\":false"))).toBe(30);
    expect(run.toolContents[30]).toBe(DENIAL.RESOURCE_LIMIT);
    expect(formulaWrites(run.primitiveLog)).toHaveLength(30);
    const segments = runSegments(run.primitiveLog);
    expect(segments).toHaveLength(60);
    segments.forEach((segment, i) => {
      if (i % 2 === 0) expect(formulaWrites(segment)).toHaveLength(1);
      else {
        expect(segment.some((entry) => entry.startsWith("set "))).toBe(false);
        expect(segment.some((entry) => entry.startsWith("load range("))).toBe(true);
      }
    });
    expect(run.authority?.usageSnapshot()).toEqual({ ...ZERO_USAGE, operationCount: 30, cellsWritten: 30, cellsRead: 30, verificationReads: 30 });
  });

  it("D2 two 2600-cell writes: call 2 RESOURCE_LIMIT, one write", async () => {
    classified(PROMPTS.fill);
    const calls: Call[] = [
      ["write_range", { sheet: "Sheet1", address: "A1:Z100", values: grid(100, 26, 1) }],
      ["write_range", { sheet: "Sheet1", address: "A101:Z200", values: grid(100, 26, 2) }],
    ];
    const run = await runDirect({ prompt: PROMPTS.fill, rollout: true, responses: [calls, "done"] });
    expect(run.toolContents[0]).toBe((await rawPayloads(calls.slice(0, 1)))[0]);
    expect(run.toolContents[1]).toBe(DENIAL.RESOURCE_LIMIT);
    expect(formulaWrites(run.primitiveLog)).toHaveLength(1);
    expect(excelRuns(run.primitiveLog)).toBe(2);
  });

  it("D3 identical write twice: 2nd DUPLICATE, 3rd still hits the agent's repeat guard", async () => {
    classified(PROMPTS.fill);
    const call = oneCell(1);
    const run = await runDirect({ prompt: PROMPTS.fill, rollout: true, responses: [[call, call, call], "done"] });
    expect(run.toolContents[0]).toBe((await rawPayloads([call]))[0]);
    expect(run.toolContents[1]).toBe(DENIAL.DUPLICATE);
    expect(JSON.parse(run.toolContents[1]).category).toBe("repetition");
    expect(run.toolContents[2]).toBe(REPEAT_BLOCKED);
    expect(formulaWrites(run.primitiveLog)).toHaveLength(1);
  });

  it("D4 ambiguous request (borra la hoja): write denied, zero Excel.run", async () => {
    classified(PROMPTS.borra);
    const run = await runDirect({ prompt: PROMPTS.borra, rollout: true, responses: [[oneCell(1)], "done"] });
    expect(run.toolContents).toEqual([DENIAL.CONFIRMATION_REQUIRED]);
    expect(excelRuns(run.primitiveLog)).toBe(0);
  });

  it("D5 read-only request: write denied, run_calculation allowed", async () => {
    classified(PROMPTS.summarize);
    const run = await runDirect({ prompt: PROMPTS.summarize, rollout: true, responses: [[oneCell(1), ["run_calculation", {}]], "done"] });
    expect(run.toolContents[0]).toBe(DENIAL.CONFIRMATION_REQUIRED);
    expect(run.toolContents[1]).toBe((await rawPayloads([["run_calculation", {}]]))[0]);
    expect(JSON.parse(run.toolContents[1]).ok).toBe(true);
    expect(formulaWrites(run.primitiveLog)).toHaveLength(0);
    expect(excelRuns(run.primitiveLog)).toBe(1);
  });

  it("D6 legacy / no plan path tools: each LEGACY_NO_PLAN_PATH, zero Excel.run", async () => {
    classified(PROMPTS.createTable);
    const run = await runDirect({
      prompt: PROMPTS.createTable,
      rollout: true,
      responses: [[
        ["delete_sheet", { name: "Old" }],
        ["protect_sheet", { sheet: "Sheet1" }],
        ["create_pivot_table", { sheet: "Sheet1", address: "H1", sourceSheet: "Sheet1", sourceRange: "A1:C4", rowFields: ["a"], dataFields: [{ name: "b" }] }],
      ], "done"],
    });
    expect(run.toolContents).toEqual([DENIAL.LEGACY_NO_PLAN_PATH, DENIAL.LEGACY_NO_PLAN_PATH, DENIAL.LEGACY_NO_PLAN_PATH]);
    expect(excelRuns(run.primitiveLog)).toBe(0);
  });

  it("D7 selection-scoped request: write denied", async () => {
    classified(PROMPTS.selected);
    const run = await runDirect({ prompt: PROMPTS.selected, rollout: true, responses: [[oneCell(1)], "done"] });
    expect(run.toolContents).toEqual([DENIAL.CONFIRMATION_REQUIRED]);
    expect(excelRuns(run.primitiveLog)).toBe(0);
  });

  it("D8 readback mismatch on call 1 latches: calls 2-3 READBACK_FAILED, one write", async () => {
    classified(PROMPTS.fill);
    const calls: Call[] = [oneCell(1), oneCell(2), ["format_range", { sheet: "Sheet1", address: "A1:B2", format: { bold: true } }]];
    const run = await runDirect({ prompt: PROMPTS.fill, rollout: true, responses: [calls, "done"], configure: (d) => { d.mismatchFormulas = true; } });
    expect(run.toolContents[0]).toBe((await rawPayloads(calls.slice(0, 1)))[0]);
    expect(JSON.parse(run.toolContents[0]).ok).toBe(true);
    expect(run.toolContents.slice(1)).toEqual([DENIAL.READBACK_FAILED, DENIAL.READBACK_FAILED]);
    expect(formulaWrites(run.primitiveLog)).toHaveLength(1);
    expect(excelRuns(run.primitiveLog)).toBe(2);
    expect(run.primitiveLog.some((entry) => entry.includes(".format."))).toBe(false);
  });

  it.each([false, true])("D9 no primitiveAuthority (rollout=%s): write_range hits the read-only sink, read_range is byte-identical", async (rollout) => {
    classified(PROMPTS.fill);
    const calls: Call[] = [oneCell(1), ["read_range", { sheet: "Sheet1", address: "A1:B2" }]];
    const run = await runDirect({ prompt: PROMPTS.fill, rollout, responses: [calls, "done"], withAuthority: false });
    expect(run.toolContents[0]).toBe(DENIAL.READ_ONLY_SINK);
    expect(run.toolContents[1]).toBe((await rawPayloads(calls.slice(1)))[0]);
    expect(formulaWrites(run.primitiveLog)).toHaveLength(0);
  });

  it.each([true, false])("D10 finishTurn while the 2nd response is pending (rollout=%s): TURN_INACTIVE, zero writes", async (rollout) => {
    classified(PROMPTS.fill);
    const run = await runDirect({
      prompt: PROMPTS.fill,
      rollout,
      responses: [[["read_range", { sheet: "Sheet1", address: "A1" }]], [oneCell(1)], "done"],
      onRequest: (index, { coordinator, handle }) => { if (index === 1) coordinator.finishTurn(handle); },
    });
    expect(run.toolContents[1]).toBe(DENIAL.TURN_INACTIVE);
    expect(formulaWrites(run.primitiveLog)).toHaveLength(0);
    expect(run.authority?.isTerminal).toBe(true);
  });

  it.each([
    [PROMPTS.showTable, "confirmation_required"],
    [PROMPTS.dontChange, "ambiguous_mutation_intent"],
  ] as const)("D11 [R4b] %s: write_range CONFIRMATION_REQUIRED (unit mirror reason %s)", async (prompt: string, _reason: string) => {
    classified(prompt);
    const run = await runDirect({ prompt, rollout: true, responses: [[oneCell(1)], "done"] });
    expect(run.toolContents).toEqual([DENIAL.CONFIRMATION_REQUIRED]);
    expect(excelRuns(run.primitiveLog)).toBe(0);
  });

  it("D12 named range and whole column: both RESOURCE_LIMIT, zero Excel.run", async () => {
    classified(PROMPTS.fill);
    const run = await runDirect({
      prompt: PROMPTS.fill,
      rollout: true,
      responses: [[["write_range", { sheet: "Sheet1", address: "MyNamedRange", values: [[1]] }], ["write_range", { sheet: "Sheet1", address: "A:A", values: [[1]] }]], "done"],
    });
    expect(run.toolContents).toEqual([DENIAL.RESOURCE_LIMIT, DENIAL.RESOURCE_LIMIT]);
    expect(excelRuns(run.primitiveLog)).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// §5.2 P-rows — rollout OFF parity (GREEN at base and HEAD, same assertions)
// ---------------------------------------------------------------------------
async function expectParity(prompt: string, responses: readonly ModelResponse[], configure?: (double: OfficeDouble) => void): Promise<DirectRun> {
  const calls = responses.filter((response): response is readonly Call[] => typeof response !== "string").flat();
  const run = await runDirect({ prompt, rollout: false, responses, configure });
  const reference = await directOracle(calls, configure);
  expect(run.toolContents).toEqual(reference.payloads);
  expect(run.double.log).toEqual(reference.log);
  if (authorityModulePresent) {
    expect(run.authority).toBeDefined();
    expect(run.authority!.rolloutEnabled).toBe(false);
    expect(run.authority!.usageSnapshot()).toEqual(ZERO_USAGE);
  } else {
    expect(run.authority).toBeUndefined();
  }
  return run;
}

describe("P-rows: rollout OFF is byte-identical to the base", () => {
  it("P1 31 distinct one-cell writes: 31 writes, 31 Excel.run, no readback", async () => {
    classified(PROMPTS.fill);
    const calls = Array.from({ length: 31 }, (_v, i) => oneCell(i + 1));
    const run = await expectParity(PROMPTS.fill, [...chunks(calls, [10, 10, 10, 1]), "done"]);
    expect(formulaWrites(run.primitiveLog)).toHaveLength(31);
    expect(excelRuns(run.primitiveLog)).toBe(31);
  });

  it("P2 two 2600-cell writes: both written", async () => {
    classified(PROMPTS.fill);
    const run = await expectParity(PROMPTS.fill, [[
      ["write_range", { sheet: "Sheet1", address: "A1:Z100", values: grid(100, 26, 1) }],
      ["write_range", { sheet: "Sheet1", address: "A101:Z200", values: grid(100, 26, 2) }],
    ], "done"]);
    expect(formulaWrites(run.primitiveLog)).toHaveLength(2);
  });

  it("P3 identical write twice then a third: 1-2 execute, 3rd gets the existing repeat guard", async () => {
    classified(PROMPTS.fill);
    const call = oneCell(1);
    const run = await expectParity(PROMPTS.fill, [[call, call], [call], "done"]);
    expect(run.toolContents[2]).toBe(REPEAT_BLOCKED);
    expect(formulaWrites(run.primitiveLog)).toHaveLength(2);
  });

  it("P4 ambiguous request: the write executes", async () => {
    classified(PROMPTS.borra);
    const run = await expectParity(PROMPTS.borra, [[oneCell(1)], "done"]);
    expect(formulaWrites(run.primitiveLog)).toHaveLength(1);
  });

  it("P5 legacy tools, each twice: all execute", async () => {
    classified(PROMPTS.createTable);
    const legacy: Call[] = [
      ["delete_sheet", { name: "Old" }],
      ["clear_range", { sheet: "Sheet1", address: "A1:B2" }],
      ["protect_sheet", { sheet: "Sheet1" }],
      ["create_pivot_table", { sheet: "Sheet1", address: "H1", sourceSheet: "Sheet1", sourceRange: "A1:C4", rowFields: ["a"], dataFields: [{ name: "b" }] }],
    ];
    const run = await expectParity(PROMPTS.createTable, [legacy.flatMap((call) => [call, call]), "done"]);
    expect(excelRuns(run.primitiveLog)).toBe(8);
    expect(run.toolContents.some((content) => content.includes("Workbook edit denied"))).toBe(false);
  });

  it("P6 readback-mismatch hook armed: write and format execute, no readback Excel.run", async () => {
    classified(PROMPTS.fill);
    const run = await expectParity(PROMPTS.fill, [[oneCell(1), ["format_range", { sheet: "Sheet1", address: "A1:B2", format: { bold: true } }]], "done"], (d) => { d.mismatchFormulas = true; });
    expect(excelRuns(run.primitiveLog)).toBe(2);
    expect(formulaWrites(run.primitiveLog)).toHaveLength(1);
  });

  it("P7 named range and whole column: raw results, never RESOURCE_LIMIT", async () => {
    classified(PROMPTS.fill);
    const run = await expectParity(PROMPTS.fill, [[
      ["write_range", { sheet: "Sheet1", address: "MyNamedRange", values: [[1]] }],
      ["write_range", { sheet: "Sheet1", address: "A:A", values: [[1]] }],
    ], "done"]);
    expect(run.toolContents).not.toContain(DENIAL.RESOURCE_LIMIT);
    expect(excelRuns(run.primitiveLog)).toBe(2);
  });

  it("P8 sort x2, filter x2, run_calculation x3, 11 x create_sheet: all execute (repeat guard unchanged)", async () => {
    classified(PROMPTS.createTable);
    const calls: Call[] = [
      ["sort_table", { tableName: "Table1", columnIndex: 0, ascending: true }],
      ["sort_table", { tableName: "Table1", columnIndex: 0, ascending: true }],
      ["filter_table", { tableName: "Table1", columnIndex: 1, values: ["a"] }],
      ["filter_table", { tableName: "Table1", columnIndex: 1, values: ["a"] }],
      ["run_calculation", {}],
      ["run_calculation", {}],
      ["run_calculation", {}],
      ...Array.from({ length: 11 }, (_v, i): Call => ["create_sheet", { name: `S${i + 1}` }]),
    ];
    const run = await expectParity(PROMPTS.createTable, [...chunks(calls, [9, 9]), "done"]);
    expect(run.toolContents[6]).toBe(REPEAT_BLOCKED);
    expect(excelRuns(run.primitiveLog)).toBe(17);
  });

  it.each([PROMPTS.summarize, PROMPTS.selected])("P9 %s: the write executes (no intent or selection rule)", async (prompt) => {
    classified(prompt);
    const run = await expectParity(prompt, [[oneCell(1)], "done"]);
    expect(formulaWrites(run.primitiveLog)).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// G1 / G4
// ---------------------------------------------------------------------------
describe("G-rows", () => {
  const READS: Call[] = [
    ["read_range", { sheet: "Sheet1", address: "A1:B2" }],
    ["summarize_range_stats", { sheet: "Sheet1", address: "A1:B2" }],
    ["get_workbook_overview", {}],
  ];
  const G1_ROWS = [false, true].flatMap((rollout) => Object.values(PROMPTS).map((prompt) => ({ rollout, prompt })));

  it.each(G1_ROWS)("G1 read-only primitives are byte-identical and uncharged: rollout=$rollout $prompt", async ({ rollout, prompt }) => {
    classified(prompt);
    const run = await runDirect({ prompt, rollout, responses: [READS, "done"] });
    const reference = await directOracle(READS);
    expect(run.toolContents).toEqual(reference.payloads);
    expect(run.double.log).toEqual(reference.log);
    if (authorityModulePresent) expect(run.authority!.usageSnapshot()).toEqual(ZERO_USAGE);
  });

  it("G4 an excel_execute_plan response ends agent.run without dispatching a later primitive", async () => {
    classified(PROMPTS.fill);
    let planTurns = 0;
    const run = await runDirect({
      prompt: PROMPTS.fill,
      rollout: true,
      responses: [[["excel_execute_plan", { version: 1, outcome: "execute", summary: "s", operations: [{ kind: "calculate" }], verification: [{ kind: "calculation_completed" }] }]], [oneCell(1)], "done"],
      onWorkbookPlanTurn: async () => { planTurns += 1; return "plan handled"; },
    });
    expect(run.requests).toBe(1);
    expect(planTurns).toBe(1);
    expect(run.finalContent).toBe("plan handled");
    expect(excelRuns(run.primitiveLog)).toBe(0);
  });
});
