/**
 * R4c unit tests: the per-turn PrimitiveMutationAuthority, the read-only
 * exported primitive sinks, the coordinator lifecycle, resource/intent parity
 * and the static single-raw-path scans (dev3-r4c-release.md §5.1 U1-U10,
 * §5.2 G2/G5/G6/P10).
 *
 * The authority module is new in R4c, so it is loaded defensively: at the base
 * every authority row fails with an AssertionError, never with a type error.
 */
import { createHash } from "node:crypto";
import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import {
  executeExcelTool,
  executeExcelToolParsed,
  getExcelToolDefinition,
  getPrimitiveEffect,
  toolSchemas,
} from "../excelTools";
import * as excelToolsModule from "../excelTools";
import * as officeBridgeExcelModule from "../officeBridgeExcel";
import { executeValidatedExcelBridgeTool } from "../officeBridgeExcel";
import * as resourcesModule from "../workbookPlanResources";
import { accountWorkbookPlanResources, WORKBOOK_PLAN_RESOURCE_LIMITS } from "../workbookPlanResources";
import * as policyModule from "../workbookPlanPolicy";
import {
  createWorkbookPlanPolicyInputSnapshot,
  decideWorkbookPlanPolicy,
  WORKBOOK_PLAN_CAPABILITY_NAMES,
  WORKBOOK_PLAN_OPERATION_EFFECTS,
  type WorkbookPlanOperationEffect,
  type WorkbookPlanPolicyDecision,
  type WorkbookPlanResourceSummary,
  type WorkbookRequestClassification,
} from "../workbookPlanPolicy";
import { classifyWorkbookRequest, type TrustedWorkbookRequest } from "../workbookPlanRequest";
import {
  TaskPaneTurnCoordinator,
  type TaskPaneTurnHandle,
  type TaskPaneTurnStartContext,
} from "../taskpane/taskPaneTurnCoordinator";
import { ServerOwnedTurnBindingStore } from "../workbookPlanOrchestrator";
import { validateToolCall } from "../toolValidation";
import { createServerOwnedPlanRoute } from "../taskpane/serverOwnedPlanRoute";
import { run as runOfficeAcceptanceFixtures, type AcceptanceFixtureEvidence, type AcceptanceFixtureRuntime } from "../taskpane/officeAcceptanceHarness";
import { getDirectProviderTools, getDirectSystemPrompt, getServerOwnedSystemPrompt } from "../agent";
import type { WorkbookOperation, WorkbookPlan } from "../generated/workbookPlanSchema";

// ---------------------------------------------------------------------------
// Recording Office double (R4c §5 doubles). Every Excel.run, sync, queued
// property set, load and method call is appended to `log`. Range values are
// kept per cell so a readback sees what was written; `mismatchFormulas` makes
// every `formulas` read return a mismatching grid (the readback hook).
// ---------------------------------------------------------------------------
interface OfficeDouble {
  readonly log: string[];
  mismatchFormulas: boolean;
  readbackCell?: unknown;
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
  const state = { mismatchFormulas: false, failRuns: false, readbackCell: undefined as unknown };

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
        get: () => (state.readbackCell !== undefined ? grid(() => state.readbackCell) : state.mismatchFormulas ? grid(() => "MISMATCH") : grid(read)),
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
    get readbackCell() { return state.readbackCell; },
    set readbackCell(value: unknown) { state.readbackCell = value; },
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
// Fixed expectations (independent constants, I6)
// ---------------------------------------------------------------------------
const deny = (error: string, category: string): string => JSON.stringify({ ok: false, error, category });
const DENIAL = {
  READ_ONLY_SINK: deny("Workbook edits must go through the task-pane edit authority.", "validation"),
  TURN_INACTIVE: deny("Workbook edit denied: no active task-pane turn.", "validation"),
  UNCLASSIFIED: deny("Workbook edit denied: this tool is not classified.", "validation"),
  LEGACY_NO_PLAN_PATH: deny("Workbook edit denied: this legacy tool has no WorkbookPlan path and is disabled while WorkbookPlan rollout is on.", "validation"),
  CONFIRMATION_REQUIRED: deny("Workbook edit denied: this request needs confirmation. Use excel_execute_plan (WorkbookPlan) instead.", "validation"),
  DUPLICATE: deny("Workbook edit denied: the identical edit was already applied in this turn.", "repetition"),
  RESOURCE_LIMIT: deny("Workbook edit denied: the per-turn workbook edit limit was reached.", "validation"),
  READBACK_FAILED: deny("Workbook edit denied: an earlier edit in this turn did not verify.", "validation"),
  PLAN_ENTERED: deny("Workbook edit denied: a workbook plan already ran in this turn.", "validation"),
} as const;
type DenialCode = keyof typeof DENIAL;

const ALL_NAMES = toolSchemas.map((schema) => schema.function.name);
const READ_ONLY_NAMES = ["get_workbook_overview", "read_range", "summarize_range_stats"];
const LEGACY_NAMES = ["delete_sheet", "clear_range", "clear_comments", "find_replace", "unprotect_sheet", "protect_sheet", "import_csv_to_sheet", "create_pivot_table"];

const VALID_ARGS: Readonly<Record<string, Record<string, unknown>>> = {
  get_workbook_overview: {},
  read_range: { sheet: "Sheet1", address: "A1:B2" },
  summarize_range_stats: { sheet: "Sheet1", address: "A1:B2" },
  write_range: { sheet: "Sheet1", address: "A1:B1", values: [[1, "x"]] },
  format_range: { sheet: "Sheet1", address: "A1:B2", format: { bold: true } },
  conditional_format_range: { sheet: "Sheet1", address: "A1:A5", ruleType: "cellValue", operator: "GreaterThan", formula1: "5", fillColor: "#ff0000" },
  freeze_panes: { sheet: "Sheet1", rows: 1, columns: 0 },
  set_data_validation: { sheet: "Sheet1", address: "B1:B5", type: "list", source: "Yes,No" },
  create_sheet: { name: "Added" },
  delete_sheet: { name: "Old" },
  activate_sheet: { name: "Sheet1" },
  insert_table: { sheet: "Sheet1", address: "A1:C4", hasHeaders: true },
  sort_table: { tableName: "Table1", columnIndex: 0, ascending: true },
  filter_table: { tableName: "Table1", columnIndex: 1, values: ["a"] },
  create_chart: { sheet: "Sheet1", dataRange: "A1:B5", chartType: "ColumnClustered", title: "T" },
  find_replace: { sheet: "Sheet1", find: "a", replace: "b", matchCase: false },
  clear_range: { sheet: "Sheet1", address: "A1:B2" },
  autofill_formula: { sheet: "Sheet1", sourceAddress: "C1", targetAddress: "C2:C5" },
  rename_sheet: { oldName: "Sheet2", newName: "Data" },
  copy_sheet: { sourceName: "Sheet1", newName: "Copy" },
  protect_sheet: { sheet: "Sheet1" },
  unprotect_sheet: { sheet: "Sheet1" },
  create_pivot_table: { sheet: "Sheet1", address: "H1", sourceSheet: "Sheet1", sourceRange: "A1:C4", rowFields: ["a"], dataFields: [{ name: "b" }] },
  run_calculation: {},
  add_comment: { sheet: "Sheet1", address: "D1", text: "note" },
  clear_comments: { sheet: "Sheet1", address: "D1:D5" },
  import_csv_to_sheet: { sheet: "Sheet1", address: "A1", csv: "a,b\n1,2", hasHeaders: true },
};

const REQUEST = (classification: WorkbookRequestClassification, selectionScopedRequest = false): TrustedWorkbookRequest =>
  Object.freeze({ classification, selectionScopedRequest });

// ---------------------------------------------------------------------------
// Defensive access to R4c surfaces (absent at the base)
// ---------------------------------------------------------------------------
interface AuthorityLike {
  executePrimitive(name: string, args: Record<string, unknown>): Promise<string>;
  markWorkbookPlanEntered(): void;
  terminate(): void;
  readonly isTerminal: boolean;
  readonly rolloutEnabled: boolean;
  usageSnapshot(): Readonly<Record<string, number>>;
}
interface AuthorityModule {
  createPrimitiveMutationAuthority(context: { trustedRequest: TrustedWorkbookRequest | null; workbookPlanRolloutEnabled: boolean }): AuthorityLike;
  createServerPrimitiveGate?(deps: Record<string, unknown>): (name: string, args: Record<string, unknown>) => Promise<string>;
  PRIMITIVE_DENIAL_MESSAGES: Readonly<Record<string, string>>;
  PrimitiveMutationAuthority: unknown;
}
const AUTHORITY_MODULE_PATH = "../primitiveMutationAuthority";

async function loadAuthorityModule(): Promise<AuthorityModule | undefined> {
  const mod: unknown = await import(/* @vite-ignore */ AUTHORITY_MODULE_PATH).catch(() => undefined);
  return mod as AuthorityModule | undefined;
}

let authorityModule: AuthorityModule | undefined;
beforeAll(async () => {
  authorityModule = await loadAuthorityModule();
});

function requireAuthorityModule(): AuthorityModule {
  expect(authorityModule).toBeDefined();
  const record = authorityModule as unknown as Record<string, unknown>;
  expect(record.createPrimitiveMutationAuthority).toBeTypeOf("function");
  return authorityModule as AuthorityModule;
}

function mint(rollout: boolean, request: TrustedWorkbookRequest | null): AuthorityLike {
  return requireAuthorityModule().createPrimitiveMutationAuthority({ trustedRequest: request, workbookPlanRolloutEnabled: rollout });
}

function exportedFunction<T>(module: object, name: string): T {
  const value = (module as unknown as Record<string, unknown>)[name];
  expect(value).toBeTypeOf("function");
  return value as T;
}

function coordinatorDeps() {
  return {
    getOrCreateDocumentId: async () => "11111111-1111-4111-8111-111111111111",
    getBridgeIdentity: () => ({ sessionId: null, documentId: null }),
    reconnectForDocumentChange: async () => { throw new Error("unused"); },
    bindings: new ServerOwnedTurnBindingStore({ now: () => Date.now(), ttlMs: 90_000, maxEntries: 32, maxBytes: 64 * 1024 }),
    paneId: "22222222-2222-4222-8222-222222222222",
  };
}

function authorityFor(coordinator: TaskPaneTurnCoordinator, handle: TaskPaneTurnHandle): AuthorityLike | undefined {
  const value = (coordinator as unknown as { primitiveAuthorityFor?: (h: TaskPaneTurnHandle) => unknown }).primitiveAuthorityFor?.(handle) ?? undefined;
  return value as AuthorityLike | undefined;
}

async function oracle(calls: ReadonlyArray<readonly [string, Record<string, unknown>]>, configure?: (d: OfficeDouble) => void): Promise<{ payloads: string[]; log: string[] }> {
  const double = createOfficeDouble();
  configure?.(double);
  double.install();
  const payloads: string[] = [];
  for (const [name, args] of calls) {
    const tool = getExcelToolDefinition(name);
    payloads.push(tool === undefined ? JSON.stringify({ ok: false, error: `Unknown tool: ${name}`, name: "Error" }) : await tool.execute(args));
  }
  return { payloads, log: [...double.log] };
}

const ZERO_USAGE: WorkbookPlanResourceSummary = Object.freeze({
  operationCount: 0, cellsRead: 0, cellsWritten: 0, formulaUtf8Bytes: 0, csvPayloadUtf8Bytes: 0,
  commentUtf8Bytes: 0, sheetsCreated: 0, chartsCreated: 0, tablesCreated: 0, verificationReads: 0,
});

// create_pivot_table names its pivot from Date.now(): pin the clock so the
// authority run and its oracle replay are deterministic.
const FIXED_NOW = 1_800_000_000_000;
beforeEach(() => {
  vi.spyOn(Date, "now").mockReturnValue(FIXED_NOW);
});

afterEach(() => {
  vi.restoreAllMocks();
  delete (globalThis as Record<string, unknown>).Excel;
  vi.doUnmock("../excelTools");
});

// ---------------------------------------------------------------------------
// U1 — exported sinks are read-only (I2)
// ---------------------------------------------------------------------------
describe("U1 exported primitive sinks are read-only (I2)", () => {
  it("has 27 primitives: 3 read_only and 24 others, each with valid fixture arguments", () => {
    expect(ALL_NAMES).toHaveLength(27);
    expect(ALL_NAMES.filter((name) => getPrimitiveEffect(name)?.effectClass === "read_only").sort()).toEqual([...READ_ONLY_NAMES].sort());
    for (const name of ALL_NAMES) {
      const validation = validateToolCall(name, JSON.stringify(VALID_ARGS[name]));
      expect("valid" in validation && validation.valid === true, name).toBe(true);
    }
  });

  const nonReadOnly = Object.keys(VALID_ARGS).filter((name) => !READ_ONLY_NAMES.includes(name));
  it.each(nonReadOnly)("denies %s through all three sinks with zero Excel.run", async (name) => {
    const double = createOfficeDouble();
    double.install();
    expect(await executeExcelTool(name, JSON.stringify(VALID_ARGS[name]))).toBe(DENIAL.READ_ONLY_SINK);
    expect(await executeExcelToolParsed(name, VALID_ARGS[name])).toBe(DENIAL.READ_ONLY_SINK);
    expect(await executeValidatedExcelBridgeTool(name, VALID_ARGS[name])).toBe(DENIAL.READ_ONLY_SINK);
    expect(double.log).toEqual([]);
  });

  it.each(READ_ONLY_NAMES)("keeps %s byte-identical to the raw definition through all three sinks", async (name) => {
    const expected = await oracle([[name, VALID_ARGS[name]]]);
    for (const sink of [
      () => executeExcelTool(name, JSON.stringify(VALID_ARGS[name])),
      () => executeExcelToolParsed(name, VALID_ARGS[name]),
      () => executeValidatedExcelBridgeTool(name, VALID_ARGS[name]),
    ]) {
      const double = createOfficeDouble();
      double.install();
      expect(await sink()).toBe(expected.payloads[0]);
      expect(double.log).toEqual(expected.log);
    }
  });

  it("keeps the unknown-name result first and makes no Office call", async () => {
    const double = createOfficeDouble();
    double.install();
    expect(JSON.parse(await executeExcelTool("nope_tool", "{}")).error).toBe("Unknown tool: nope_tool");
    expect(JSON.parse(await executeExcelToolParsed("nope_tool", {})).error).toBe("Unknown tool: nope_tool");
    expect(JSON.parse(await executeExcelToolParsed("__proto__", {})).error).toBe("Unknown tool: __proto__");
    expect(double.log).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// U2 — single raw path: static scans (I1/I3) and the one-shot latch
// ---------------------------------------------------------------------------
const SRC_DIR = fileURLToPath(new URL("..", import.meta.url));

function productionFiles(): string[] {
  const out: string[] = [];
  const walk = (dir: string, rel: string): void => {
    for (const entry of readdirSync(dir).sort()) {
      const full = join(dir, entry);
      const relPath = rel === "" ? entry : `${rel}/${entry}`;
      if (statSync(full).isDirectory()) {
        if (entry === "__tests__" || entry === "generated") continue;
        walk(full, relPath);
      } else if (entry.endsWith(".ts")) {
        out.push(relPath);
      }
    }
  };
  walk(SRC_DIR, "");
  return out;
}

function stripComments(source: string): string {
  let out = "";
  let i = 0;
  while (i < source.length) {
    const ch = source[i];
    const next = source[i + 1];
    if (ch === "/" && next === "/") {
      while (i < source.length && source[i] !== "\n") i += 1;
      continue;
    }
    if (ch === "/" && next === "*") {
      const end = source.indexOf("*/", i + 2);
      const stop = end < 0 ? source.length : end + 2;
      out += "\n".repeat((source.slice(i, stop).match(/\n/g) ?? []).length);
      i = stop;
      continue;
    }
    if (ch === "\"" || ch === "'" || ch === "`") {
      let j = i + 1;
      while (j < source.length && source[j] !== ch) {
        if (source[j] === "\\") j += 1;
        j += 1;
      }
      out += source.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    out += ch;
    i += 1;
  }
  return out;
}

function countsByFile(pattern: RegExp): Record<string, number> {
  const result: Record<string, number> = {};
  for (const file of productionFiles()) {
    const text = stripComments(readFileSync(join(SRC_DIR, file), "utf8"));
    const count = (text.match(new RegExp(pattern.source, "g")) ?? []).length;
    if (count > 0) result[file] = count;
  }
  return result;
}

describe("U2 single raw path (I1/I3 static scans, latch)", () => {
  it("lists exactly 39 non-test .ts files including R7-H, theme and profile metadata modules", () => {
    const files = productionFiles();
    expect(files).toContain("primitiveMutationAuthority.ts");
    expect(files).toContain("taskpane/reconciliationNotice.ts");
    expect(files).toContain("taskpane/workbookPlanAcceptanceHarness.ts");
    expect(files).toContain("taskpane/workbookPlanAcceptanceEntry.ts");
    expect(files).toContain("taskpane/omniaTheme.ts");
    expect(files).toContain("taskpane/omniaThemeBootstrap.ts");
    expect(files).toContain("deskConnect.ts");
    expect(files).toHaveLength(39);
  });

  it("defines claimRawPrimitiveExecutor only in excelTools.ts and calls it once in the authority module", () => {
    expect(countsByFile(/function\s+claimRawPrimitiveExecutor\s*\(/)).toEqual({ "excelTools.ts": 1 });
    expect(countsByFile(/(?<!function )\bclaimRawPrimitiveExecutor\s*\(/)).toEqual({ "primitiveMutationAuthority.ts": 1 });
    expect(Object.keys(countsByFile(/\bclaimRawPrimitiveExecutor\b/)).sort()).toEqual(["excelTools.ts", "primitiveMutationAuthority.ts"]);
  });

  it("keeps .execute( in excelTools.ts only and getExcelToolDefinition( in excelTools.ts and the compiler only", () => {
    expect(countsByFile(/\.execute\(/)).toEqual({ "excelTools.ts": 2 });
    expect(Object.keys(countsByFile(/\bgetExcelToolDefinition\(/)).sort()).toEqual(["excelTools.ts", "workbookPlanCompiler.ts"]);
    const excel = stripComments(readFileSync(join(SRC_DIR, "excelTools.ts"), "utf8"));
    const raw = excel.slice(excel.indexOf("async function executeRawPrimitive("));
    expect(raw.slice(0, raw.indexOf("\n}\n"))).toContain("tool.execute(args)");
  });

  it("names no server pass-through and calls createServerPrimitiveGate only from taskpane", () => {
    expect(countsByFile(/\bexecuteLegacyServerPrimitiveR4dPending\b/)).toEqual({});
    expect(countsByFile(/\bfunction\s+createServerPrimitiveGate\(/)).toEqual({ "primitiveMutationAuthority.ts": 1 });
    expect(countsByFile(/(?<!function\s)\bcreateServerPrimitiveGate\(/)).toEqual({ "taskpane/taskpane.ts": 1 });
    expect(countsByFile(/\bserverPrimitiveAuthority\(/)).toEqual({ "taskpane/taskPaneTurnCoordinator.ts": 1, "taskpane/taskpane.ts": 1 });
    expect(countsByFile(/\.serverPrimitiveAuthority\(/)).toEqual({ "taskpane/taskpane.ts": 1 });
  });

  it("creates authorities only through createPrimitiveMutationAuthority (coordinator + harness callers)", () => {
    expect(countsByFile(/\bcreatePrimitiveMutationAuthority\(/)).toEqual({
      "primitiveMutationAuthority.ts": 1,
      "taskpane/taskPaneTurnCoordinator.ts": 1,
      "taskpane/officeAcceptanceHarness.ts": 1,
    });
    expect(countsByFile(/\bnew\s+PrimitiveMutationAuthority\b/)).toEqual({ "primitiveMutationAuthority.ts": 1 });
  });

  it("indexes PRIMITIVE_EFFECTS only inside getPrimitiveEffect, and the authority never names the table (R4a P3-3)", () => {
    expect(countsByFile(/PRIMITIVE_EFFECTS\[/)).toEqual({ "excelTools.ts": 1 });
    const excel = stripComments(readFileSync(join(SRC_DIR, "excelTools.ts"), "utf8"));
    const start = excel.indexOf("export function getPrimitiveEffect(");
    const body = excel.slice(start, excel.indexOf("\n}\n", start));
    expect(body).toContain("PRIMITIVE_EFFECTS[");
    expect(countsByFile(/\bPRIMITIVE_EFFECTS\b/)["primitiveMutationAuthority.ts"]).toBeUndefined();
  });

  it("adds no import to excelTools.ts (standalone schema generation)", () => {
    const excel = readFileSync(join(SRC_DIR, "excelTools.ts"), "utf8");
    expect(excel.match(/^import .*$/gm)).toEqual(['import { OpenAITool, ToolCall } from "./types";', 'import * as Papa from "papaparse";']);
    expect(excel).not.toMatch(/\brequire\(/);
  });

  it("hands the raw dispatcher out once: a second claim after a fresh module load throws", async () => {
    vi.resetModules();
    const freshExcel = await import("../excelTools");
    const claim = exportedFunction<() => unknown>(freshExcel, "claimRawPrimitiveExecutor");
    const freshAuthority: unknown = await import(/* @vite-ignore */ AUTHORITY_MODULE_PATH).catch(() => undefined);
    expect(freshAuthority).toBeDefined();
    expect(() => claim()).toThrow("Primitive executor already claimed.");
    vi.resetModules();
  });
});

// ---------------------------------------------------------------------------
// U3 — I4 decision-order matrix
// ---------------------------------------------------------------------------
type Outcome = { readonly kind: "raw" } | { readonly kind: "deny"; readonly code: DenialCode };

function independentIntentApproved(request: TrustedWorkbookRequest, effect: string | undefined): boolean {
  if (request.classification === "prohibited") return false;
  if (request.classification === "ambiguous_mutation") return false;
  if (effect === "structure" && request.classification !== "explicit_structure") return false;
  if (request.selectionScopedRequest) return false;
  if (request.classification === "read_only" && effect !== "calculate") return false;
  return true;
}

function expectedOutcome(name: string, rollout: boolean, request: TrustedWorkbookRequest | null, planEntered: boolean, terminal: boolean): Outcome {
  if (terminal) return { kind: "deny", code: "TURN_INACTIVE" };
  if (!rollout) return { kind: "raw" };
  if (getExcelToolDefinition(name) === undefined) return { kind: "raw" };
  const effect = getPrimitiveEffect(name);
  if (effect === undefined) return { kind: "deny", code: "UNCLASSIFIED" };
  if (effect.effectClass === "read_only") return { kind: "raw" };
  if (planEntered && effect.effectClass !== "view") return { kind: "deny", code: "PLAN_ENTERED" };
  if (effect.effectClass === "view") return { kind: "raw" };
  if (effect.effectClass === "destructive_legacy") return { kind: "deny", code: "LEGACY_NO_PLAN_PATH" };
  if (request === null || !independentIntentApproved(request, effect.planEffect)) return { kind: "deny", code: "CONFIRMATION_REQUIRED" };
  return { kind: "raw" };
}

const MATRIX_TOOLS = ["read_range", "activate_sheet", "write_range", "format_range", "insert_table", "run_calculation", "delete_sheet", "nope_tool"];
const MATRIX_REQUESTS: ReadonlyArray<readonly [string, TrustedWorkbookRequest | null]> = [
  ["explicit_mutation", REQUEST("explicit_mutation")],
  ["read_only", REQUEST("read_only")],
  ["ambiguous_mutation", REQUEST("ambiguous_mutation")],
  ["explicit_structure", REQUEST("explicit_structure")],
  ["explicit_mutation+selection", REQUEST("explicit_mutation", true)],
  ["prohibited", REQUEST("prohibited")],
  ["null", null],
];
const MATRIX_ROWS = MATRIX_TOOLS.flatMap((name) => [false, true].flatMap((rollout) => MATRIX_REQUESTS.flatMap(([label, request]) =>
  [false, true].flatMap((planEntered) => [false, true].map((terminal) => ({ name, rollout, label, request, planEntered, terminal }))))));

describe("U3 I4 decision order matrix", () => {
  it("has at least 48 rows", () => {
    expect(MATRIX_ROWS.length).toBeGreaterThanOrEqual(48);
  });

  it.each(MATRIX_ROWS)("$name rollout=$rollout request=$label planEntered=$planEntered terminal=$terminal", async (row) => {
    const authority = mint(row.rollout, row.request);
    if (row.planEntered) authority.markWorkbookPlanEntered();
    if (row.terminal) authority.terminate();
    const args = row.name === "nope_tool" ? {} : VALID_ARGS[row.name];
    const expected = expectedOutcome(row.name, row.rollout, row.request, row.planEntered, row.terminal);
    const reference = await oracle([[row.name, args]]);
    const double = createOfficeDouble();
    double.install();
    const result = await authority.executePrimitive(row.name, args);
    if (expected.kind === "deny") {
      expect(result).toBe(DENIAL[expected.code]);
      expect(double.log).toEqual([]);
      expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
    } else {
      expect(result).toBe(reference.payloads[0]);
      expect(double.log.slice(0, reference.log.length)).toEqual(reference.log);
      if (!row.rollout) expect(double.log).toEqual(reference.log);
    }
    if (!row.rollout) expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
  });

  it("serializes rollout-on calls: a second call starts only after the first settles (incl. readback)", async () => {
    const authority = mint(true, REQUEST("explicit_mutation"));
    const double = createOfficeDouble();
    double.install();
    const first = authority.executePrimitive("write_range", { sheet: "Sheet1", address: "A1", values: [[1]] });
    const second = authority.executePrimitive("write_range", { sheet: "Sheet1", address: "A2", values: [[2]] });
    await Promise.all([first, second]);
    const runs = double.log.map((entry, index) => (entry === "Excel.run" ? index : -1)).filter((index) => index >= 0);
    expect(runs).toHaveLength(4);
    const segments = runs.map((start, i) => double.log.slice(start, runs[i + 1] ?? double.log.length).join("\n"));
    expect(segments[0]).toContain("Sheet1!A1).formulas=");
    expect(segments[1]).not.toContain(".formulas=");
    expect(segments[2]).toContain("Sheet1!A2).formulas=");
    expect(segments[3]).not.toContain(".formulas=");
  });

  it("rollout off returns the raw promise without serialization or accounting", async () => {
    const authority = mint(false, REQUEST("read_only"));
    const double = createOfficeDouble();
    double.install();
    const results = await Promise.all([
      authority.executePrimitive("write_range", { sheet: "Sheet1", address: "A1", values: [[1]] }),
      authority.executePrimitive("delete_sheet", { name: "Old" }),
    ]);
    const reference = await oracle([["write_range", { sheet: "Sheet1", address: "A1", values: [[1]] }], ["delete_sheet", { name: "Old" }]]);
    expect(results).toEqual(reference.payloads);
    expect(excelRuns(double.log)).toBe(2);
    expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
  });
});

// ---------------------------------------------------------------------------
// U4 — byte identity, all 27 primitives, both modes, incl. host errors (I5)
// ---------------------------------------------------------------------------
describe("U4 byte identity (I5)", () => {
  const cases = ALL_NAMES.flatMap((name) => [false, true].flatMap((rollout) => [false, true].map((hostError) => ({ name, rollout, hostError }))))
    .filter((row) => !(row.rollout && LEGACY_NAMES.includes(row.name)));

  it("covers 27 names off and 19 on, each with and without host error", () => {
    expect(cases).toHaveLength((27 + 19) * 2);
  });

  it.each(cases)("$name rollout=$rollout hostError=$hostError", async ({ name, rollout, hostError }) => {
    const reference = await oracle([[name, VALID_ARGS[name]]], (d) => { d.failRuns = hostError; });
    const authority = mint(rollout, REQUEST("explicit_structure"));
    const double = createOfficeDouble();
    double.failRuns = hostError;
    double.install();
    const result = await authority.executePrimitive(name, VALID_ARGS[name]);
    expect(result).toBe(reference.payloads[0]);
    expect(double.log.slice(0, reference.log.length)).toEqual(reference.log);
    if (!rollout) expect(double.log).toEqual(reference.log);
    if (["write_range", "format_range", "add_comment", "create_sheet", "run_calculation", "autofill_formula"].includes(name) && !hostError) {
      expect(JSON.parse(result).ok).toBe(true);
    }
    if (hostError) expect(JSON.parse(result).ok).toBe(false);
  });

  it("legacy tools with rollout off are byte-identical even when repeated (no duplicate rule)", async () => {
    const authority = mint(false, REQUEST("explicit_structure"));
    const calls = LEGACY_NAMES.flatMap((name) => [[name, VALID_ARGS[name]], [name, VALID_ARGS[name]]] as const);
    const reference = await oracle(calls);
    const double = createOfficeDouble();
    double.install();
    const results: string[] = [];
    for (const [name, args] of calls) results.push(await authority.executePrimitive(name, args));
    expect(results).toEqual(reference.payloads);
    expect(double.log).toEqual(reference.log);
  });
});

// ---------------------------------------------------------------------------
// U5 — canary never appears in a denial (I6)
// ---------------------------------------------------------------------------
describe("U5 denial texts never echo argument content (I6)", () => {
  const CANARY = "CANARY-7f3e";

  it("fixes the 9 denial texts in PRIMITIVE_DENIAL_MESSAGES", () => {
    const mod = requireAuthorityModule();
    const table = mod.PRIMITIVE_DENIAL_MESSAGES;
    expect(Object.isFrozen(table)).toBe(true);
    expect(Object.keys(table).sort()).toEqual(Object.keys(DENIAL).sort());
    for (const code of Object.keys(DENIAL) as DenialCode[]) {
      expect(table[code]).toBe(JSON.parse(DENIAL[code]).error);
    }
  });

  it("returns canary-free denials on every rule", async () => {
    const results: string[] = [];
    const canaryWrite = { sheet: CANARY, address: "A1", values: [[CANARY]] };
    createOfficeDouble().install();
    results.push(await executeExcelToolParsed("write_range", canaryWrite));
    results.push(await executeExcelTool("write_range", JSON.stringify(canaryWrite)));
    const terminal = mint(true, REQUEST("explicit_mutation"));
    terminal.terminate();
    results.push(await terminal.executePrimitive("write_range", canaryWrite));
    const on = mint(true, REQUEST("explicit_mutation"));
    results.push(await on.executePrimitive("delete_sheet", { name: CANARY }));
    results.push(await mint(true, REQUEST("read_only")).executePrimitive("write_range", canaryWrite));
    results.push(await on.executePrimitive("write_range", { sheet: CANARY, address: CANARY, values: [[CANARY]] }));
    await on.executePrimitive("format_range", { sheet: CANARY, address: "A1", format: { bold: true } });
    results.push(await on.executePrimitive("format_range", { sheet: CANARY, address: "A1", format: { bold: true } }));
    const mismatch = createOfficeDouble();
    mismatch.mismatchFormulas = true;
    mismatch.install();
    const latched = mint(true, REQUEST("explicit_mutation"));
    await latched.executePrimitive("write_range", canaryWrite);
    results.push(await latched.executePrimitive("write_range", { ...canaryWrite, address: "A2" }));
    const planned = mint(true, REQUEST("explicit_mutation"));
    planned.markWorkbookPlanEntered();
    results.push(await planned.executePrimitive("write_range", canaryWrite));
    expect(results).toEqual([
      DENIAL.READ_ONLY_SINK, DENIAL.READ_ONLY_SINK, DENIAL.TURN_INACTIVE, DENIAL.LEGACY_NO_PLAN_PATH,
      DENIAL.CONFIRMATION_REQUIRED, DENIAL.RESOURCE_LIMIT, DENIAL.DUPLICATE, DENIAL.READBACK_FAILED, DENIAL.PLAN_ENTERED,
    ]);
    for (const result of results) {
      expect(result).not.toContain(CANARY);
      expect(Object.keys(JSON.parse(result)).sort()).toEqual(["category", "error", "ok"]);
    }
  });
});

// ---------------------------------------------------------------------------
// U6 — coordinator lifecycle (I7)
// ---------------------------------------------------------------------------
describe("R7-B1 primitive readback content equality", () => {
  it.each([
    ["123", 999], ["123", true], ["text", 123], ["text", true], ["true", false],
    ["10/8/2026", 46303], [16, "0x10"],
    ['="A B"', '="AB"'], ['="Ab"', '="AB"'], ['="A""b"', '="A""B"'],
    ["='A B'!A1", "='AB'!A1"], ["=SUM(A1 B1)", "=SUM(A1B1)"],
  ])("mismatched content %j / %j latches and prevents later host mutation", async (expected, actual) => {
    const double = createOfficeDouble(); double.readbackCell = actual; double.install();
    const authority = mint(true, REQUEST("explicit_mutation"));
    const first = JSON.parse(await authority.executePrimitive("write_range", { sheet: "Sheet1", address: "A1", values: [[expected]] }));
    expect(first.ok).toBe(true); // Raw response retained, not a verification verdict.
    const before = [...double.log];
    expect(await authority.executePrimitive("format_range", VALID_ARGS.format_range)).toBe(DENIAL.READBACK_FAILED);
    expect(double.log).toEqual(before);
  });
  it.each([
    ["123", 123], ["1.25e2", 125], ["TRUE", true], ["false", false],
    [123, "123"], [true, "true"], [null, ""],
    ['="A B"', '="A B"'], ['="A""b"', '="A""b"'], ["=sum(A1:A2)", "=SUM(A1:A2)"],
  ])("specific coercion or unchanged formula %j / %j does not latch", async (expected, actual) => {
    const double = createOfficeDouble(); double.readbackCell = actual; double.install();
    const authority = mint(true, REQUEST("explicit_mutation"));
    expect(JSON.parse(await authority.executePrimitive("write_range", { sheet: "Sheet1", address: "A1", values: [[expected]] })).ok).toBe(true);
    expect(JSON.parse(await authority.executePrimitive("format_range", VALID_ARGS.format_range)).ok).toBe(true);
  });
});

describe("R7-B1 r2 protected formula content", () => {
  // Microsoft structured-reference rules: headers are text, apostrophes
  // escape [, ], #, ' and @. Table/column lookup is case-insensitive, but this
  // readback comparator is NOT a complete Excel formula equivalence engine.
  // Case-only or formatting-only changes in protected formulas deliberately
  // latch unverified (permitted false negatives), never assert distinct Excel
  // targets from case alone. Exact same spelling remains verified.
  // https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables
  const distinctHeaders = [
    ["original proof", "=SUM(Table1[A + B])", "=SUM(Table1[A+B])"],
    ["nested operator header", "=SUM(Table1[[A + B]])", "=SUM(Table1[[A+B]])"],
    ["two header spaces", "=SUM(Table1[[A  B]])", "=SUM(Table1[[A B]])"],
    ["leading header space", "=SUM(Table1[[ A]])", "=SUM(Table1[[A]])"],
    ["trailing header space", "=SUM(Table1[[A ]])", "=SUM(Table1[[A]])"],
    ["escaped left bracket", "=SUM(Table1[[A'[ + B]])", "=SUM(Table1[[A'[+B]])"],
    ["escaped right bracket", "=SUM(Table1[[A'] + B]])", "=SUM(Table1[[A']+B]])"],
    ["escaped hash", "=SUM(Table1[[A'# + B]])", "=SUM(Table1[[A'#+B]])"],
    ["escaped apostrophe", "=SUM(Table1[[A'' + B]])", "=SUM(Table1[[A''+B]])"],
    ["string literal whitespace", '=IF(TRUE,"A  B",Table1[A])', '=IF(TRUE,"A B",Table1[A])'],
    ["string literal case", '=IF(TRUE,"Ab",Table1[A])', '=IF(TRUE,"AB",Table1[A])'],
  ];
  const equivalentButUnproven = [
    ["case-insensitive column", "=SUM(Table1[Sales])", "=SUM(Table1[sales])"],
    ["case-insensitive table", "=SUM(Table1[Sales])", "=SUM(table1[Sales])"],
    ["structured readability spacing", "=SUM(Table1[ [A]:[B] ])", "=SUM(Table1[[A]:[B]])"],
  ];
  it.each([...distinctHeaders, ...equivalentButUnproven])("%s mismatched readback blocks the subsequent real primitive", async (_label, expected, actual) => {
    const double = createOfficeDouble(); double.readbackCell = actual; double.install();
    const authority = mint(true, REQUEST("explicit_mutation"));
    expect(JSON.parse(await authority.executePrimitive("write_range", { sheet: "Sheet1", address: "A1", values: [[expected]] })).ok).toBe(true);
    const before = [...double.log];
    expect(excelRuns(before)).toBe(2); // Real write plus actual read_range.
    expect(await authority.executePrimitive("format_range", VALID_ARGS.format_range)).toBe(DENIAL.READBACK_FAILED);
    expect(double.log).toEqual(before); // No second host mutation or Office batch.
  });
  it.each([
    "=SUM(Table1[[A + B]])", "=SUM(Table1[[A  B]])", "=SUM(Table1[[ A ]])",
    "=SUM(Table1[[A'[ + B]])", "=SUM(Table1[[A'] + B]])", "=SUM(Table1[[A'# + B]])", "=SUM(Table1[[A'' + B]])",
    '=IF(TRUE,"A""b",Table1[Sales])', "=SUM(table1[sales])",
  ])("exact protected spelling %s permits the subsequent mutation", async (formula) => {
    const double = createOfficeDouble(); double.readbackCell = formula; double.install();
    const authority = mint(true, REQUEST("explicit_mutation"));
    expect(JSON.parse(await authority.executePrimitive("write_range", { sheet: "Sheet1", address: "A1", values: [[formula]] })).ok).toBe(true);
    expect(JSON.parse(await authority.executePrimitive("format_range", VALID_ARGS.format_range)).ok).toBe(true);
    expect(excelRuns(double.log)).toBe(3);
  });
});

describe("U6 coordinator lifecycle (I7)", () => {
  function begin(coordinator: TaskPaneTurnCoordinator, turnId: string, rollout: boolean, prompt = "write ok into A1") {
    const controller = new AbortController();
    const context: TaskPaneTurnStartContext & { workbookPlanRolloutEnabled?: boolean } = {
      trustedRequest: classifyWorkbookRequest(prompt),
      turnStart: Date.now(),
      cancellationSignal: controller.signal,
      workbookPlanRolloutEnabled: rollout,
    };
    return { controller, handle: coordinator.beginTurn(turnId, controller, context) };
  }

  it.each([false, true])("mints one authority per turn (rollout=%s), terminates it on finish/stop/replace", async (rollout) => {
    requireAuthorityModule();
    const coordinator = new TaskPaneTurnCoordinator(coordinatorDeps());
    const t1 = begin(coordinator, "turn-1", rollout);
    const a = authorityFor(coordinator, t1.handle);
    expect(a).toBeDefined();
    expect(a!.rolloutEnabled).toBe(rollout);
    expect(a!.isTerminal).toBe(false);
    coordinator.finishTurn(t1.handle);
    expect(a!.isTerminal).toBe(true);
    expect(authorityFor(coordinator, t1.handle)).toBeUndefined();
    createOfficeDouble().install();
    expect(await a!.executePrimitive("write_range", VALID_ARGS.write_range)).toBe(DENIAL.TURN_INACTIVE);

    const t2 = begin(coordinator, "turn-2", rollout);
    const b = authorityFor(coordinator, t2.handle);
    expect(b).toBeDefined();
    expect(b).not.toBe(a);
    expect(b!.usageSnapshot()).toEqual(ZERO_USAGE);
    coordinator.stopCurrentTurn();
    expect(b!.isTerminal).toBe(true);
    expect(t2.controller.signal.aborted).toBe(true);
    expect(authorityFor(coordinator, t2.handle)).toBeUndefined();

    const t3 = begin(coordinator, "turn-3", rollout);
    const c = authorityFor(coordinator, t3.handle);
    const t4 = begin(coordinator, "turn-4", rollout);
    const d = authorityFor(coordinator, t4.handle);
    expect(c!.isTerminal).toBe(true);
    expect(d!.isTerminal).toBe(false);
    expect(authorityFor(coordinator, t3.handle)).toBeUndefined();
  });

  it("mints with a null request when beginTurn has no context, and denies rollout-on mutations for it", async () => {
    requireAuthorityModule();
    const coordinator = new TaskPaneTurnCoordinator(coordinatorDeps());
    const handle = coordinator.beginTurn("turn-x", new AbortController());
    const authority = authorityFor(coordinator, handle);
    expect(authority).toBeDefined();
    expect(authority!.rolloutEnabled).toBe(false);
  });

  it("markWorkbookPlanEntered (rollout on) latches mutating and legacy tools; view and read-only still run", async () => {
    requireAuthorityModule();
    const coordinator = new TaskPaneTurnCoordinator(coordinatorDeps());
    const t = begin(coordinator, "turn-p", true, "create a table from this data");
    const authority = authorityFor(coordinator, t.handle)!;
    const mark = exportedFunction<(h: unknown) => void>(coordinator, "markWorkbookPlanEntered");
    mark.call(coordinator, { generation: t.handle.generation, turnId: t.handle.turnId, controller: t.handle.controller });
    const double = createOfficeDouble();
    double.install();
    expect(await authority.executePrimitive("write_range", VALID_ARGS.write_range)).toBe(DENIAL.PLAN_ENTERED);
    expect(await authority.executePrimitive("delete_sheet", VALID_ARGS.delete_sheet)).toBe(DENIAL.PLAN_ENTERED);
    expect(double.log).toEqual([]);
    const reference = await oracle([["activate_sheet", VALID_ARGS.activate_sheet], ["read_range", VALID_ARGS.read_range]]);
    const after = createOfficeDouble();
    after.install();
    expect(await authority.executePrimitive("activate_sheet", VALID_ARGS.activate_sheet)).toBe(reference.payloads[0]);
    expect(await authority.executePrimitive("read_range", VALID_ARGS.read_range)).toBe(reference.payloads[1]);
    expect(after.log).toEqual(reference.log);
  });

  it("markWorkbookPlanEntered ignores a stale handle", () => {
    requireAuthorityModule();
    const coordinator = new TaskPaneTurnCoordinator(coordinatorDeps());
    const t1 = begin(coordinator, "turn-1", true);
    const t2 = begin(coordinator, "turn-2", true);
    const mark = exportedFunction<(h: unknown) => void>(coordinator, "markWorkbookPlanEntered");
    expect(() => mark.call(coordinator, t1.handle)).not.toThrow();
    expect(authorityFor(coordinator, t2.handle)!.isTerminal).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// U7 — resource parity (13 kinds) and intent parity (50 combinations) (I8)
// ---------------------------------------------------------------------------
const EMPTY_RESOURCE_CONTEXT = {
  compileContext: {
    sheetNames: new Set<string>(),
    tablesByName: new Map(),
    supportedPrimitiveNames: new Set<string>(),
    supportedCapabilities: new Set<string>(),
    chartSourceResolutionsBySheet: new Map(),
  },
} as unknown as Parameters<typeof accountWorkbookPlanResources>[1];

const DESCRIPTOR_OPERATIONS: readonly WorkbookOperation[] = [
  { kind: "write_values", sheet: "Sheet1", address: "A1:B2", values: [[1, "=A1"], [true, "x"]] },
  { kind: "fill_formula", sheet: "Sheet1", sourceAddress: "C1", targetAddress: "C2:C9" },
  { kind: "format_range", sheet: "Sheet1", address: "A1:D4", format: { bold: true } },
  { kind: "conditional_format", sheet: "Sheet1", address: "A1:A9", ruleType: "cellValue" },
  { kind: "freeze_panes", sheet: "Sheet1", rows: 1, columns: 0 },
  { kind: "set_validation_list", sheet: "Sheet1", address: "B1:B5", source: "Yes,No" },
  { kind: "create_table", sheet: "Sheet1", address: "A1:C4", hasHeaders: true },
  { kind: "create_chart", sheet: "Sheet1", dataRange: "A1:B5", chartType: "ColumnClustered", title: "T" },
  { kind: "create_sheet", name: "Added" },
  { kind: "copy_sheet", sourceName: "Sheet1", newName: "Copy" },
  { kind: "rename_sheet", oldName: "Sheet2", newName: "Data" },
  { kind: "add_comment", sheet: "Sheet1", address: "D1", text: "note é" },
  { kind: "calculate" },
] as unknown as readonly WorkbookOperation[];

type Increment = (prior: WorkbookPlanResourceSummary, operation: WorkbookOperation, options: { readonly readbackCells: number; readonly readbackReads: number; readonly unknownTableBounds: "operation_only" }) => ReturnType<typeof accountWorkbookPlanResources>;
type Intent = (input: { readonly requestClassification: WorkbookRequestClassification; readonly selectionScopedRequest: boolean; readonly effect: WorkbookPlanOperationEffect }) => WorkbookPlanPolicyDecision;

const planOf = (operation: WorkbookOperation): WorkbookPlan =>
  ({ version: 1, outcome: "execute", summary: "s", operations: [operation], verification: [] }) as unknown as WorkbookPlan;

describe("U7 resource and intent parity (I8)", () => {
  it.each(DESCRIPTOR_OPERATIONS.map((operation) => [operation.kind, operation] as const))("increment from zero equals plan accounting: %s", (_kind, operation) => {
    const increment = exportedFunction<Increment>(resourcesModule, "accountPrimitiveResourceIncrement");
    const primitive = increment(ZERO_USAGE, operation, { readbackCells: 0, readbackReads: 0, unknownTableBounds: "operation_only" });
    expect(primitive).toEqual(accountWorkbookPlanResources(planOf(operation), EMPTY_RESOURCE_CONTEXT));
    expect(primitive.ok).toBe(true);
  });

  it("covers exactly 13 descriptor kinds", () => {
    expect(new Set(DESCRIPTOR_OPERATIONS.map((operation) => operation.kind)).size).toBe(13);
  });

  it.each([
    { kind: "write_values", sheet: "Sheet1", address: "MyNamedRange", values: [[1]] },
    { kind: "write_values", sheet: "Sheet1", address: "A:A", values: [[1]] },
    { kind: "write_values", sheet: "Sheet1", address: "A1:B1", values: [[1]] },
    { kind: "format_range", sheet: "Sheet1", address: "A1:Z300", format: { bold: true } },
  ] as unknown as WorkbookOperation[])("failures match plan accounting: %o", (operation) => {
    const increment = exportedFunction<Increment>(resourcesModule, "accountPrimitiveResourceIncrement");
    const primitive = increment(ZERO_USAGE, operation, { readbackCells: 0, readbackReads: 0, unknownTableBounds: "operation_only" });
    expect(primitive).toEqual(accountWorkbookPlanResources(planOf(operation), EMPTY_RESOURCE_CONTEXT));
    expect(primitive.ok).toBe(false);
  });

  it("adds readback cost and accumulates over a prior summary", () => {
    const increment = exportedFunction<Increment>(resourcesModule, "accountPrimitiveResourceIncrement");
    const prior = { ...ZERO_USAGE, operationCount: 29, cellsWritten: 10, cellsRead: 10, verificationReads: 29 };
    const op = { kind: "write_values", sheet: "Sheet1", address: "A1:B2", values: [[1, 2], [3, 4]] } as unknown as WorkbookOperation;
    const next = increment(prior, op, { readbackCells: 4, readbackReads: 1, unknownTableBounds: "operation_only" });
    expect(next).toEqual({ ok: true, usage: { ...prior, operationCount: 30, cellsWritten: 14, cellsRead: 14, verificationReads: 30 }, limits: WORKBOOK_PLAN_RESOURCE_LIMITS });
    const over = increment((next as { usage: WorkbookPlanResourceSummary }).usage, op, { readbackCells: 4, readbackReads: 1, unknownTableBounds: "operation_only" });
    expect(over).toEqual({ ok: false, reason: "resource_limit_exceeded" });
    expect(prior.operationCount).toBe(29);
  });

  it("D11 unit mirror [R4b]: write intent reason codes for the R4b-narrowed prompts", () => {
    const intent = exportedFunction<Intent>(policyModule, "decidePrimitiveIntent");
    const show = classifyWorkbookRequest("show me the table");
    const dont = classifyWorkbookRequest("don't change anything, just summarize the table");
    expect([show.classification, show.selectionScopedRequest]).toEqual(["read_only", false]);
    expect([dont.classification, dont.selectionScopedRequest]).toEqual(["ambiguous_mutation", false]);
    expect(intent({ requestClassification: show.classification, selectionScopedRequest: false, effect: "write" }))
      .toEqual({ decision: "requires_confirmation", reasonCode: "confirmation_required" });
    expect(intent({ requestClassification: dont.classification, selectionScopedRequest: false, effect: "write" }))
      .toEqual({ decision: "requires_confirmation", reasonCode: "ambiguous_mutation_intent" });
  });

  const CLASSES: WorkbookRequestClassification[] = ["read_only", "explicit_mutation", "ambiguous_mutation", "explicit_structure", "prohibited"];
  const INTENT_ROWS = CLASSES.flatMap((requestClassification) => [...WORKBOOK_PLAN_OPERATION_EFFECTS].flatMap((effect) =>
    [false, true].map((selectionScopedRequest) => ({ requestClassification, effect, selectionScopedRequest }))));

  it("has 50 intent combinations over 5 contract effects", () => {
    expect([...WORKBOOK_PLAN_OPERATION_EFFECTS].sort()).toEqual(["calculate", "filter", "format", "structure", "write"]);
    expect(INTENT_ROWS).toHaveLength(50);
  });

  it.each(INTENT_ROWS)("decidePrimitiveIntent equals decideWorkbookPlanPolicy: $requestClassification $effect sel=$selectionScopedRequest", (row) => {
    const intent = exportedFunction<Intent>(policyModule, "decidePrimitiveIntent");
    const plan = decideWorkbookPlanPolicy(createWorkbookPlanPolicyInputSnapshot({
      requestClassification: row.requestClassification,
      activeSheetName: "Sheet1",
      selectionScopedRequest: row.selectionScopedRequest,
      // A calculate operation has no target sheet (policy input parse); a primitive's relation to the selection is unknown.
      operationSummaries: [{ operationIndex: 0, effect: row.effect, target: { operationIndex: 0, sheetNames: row.effect === "calculate" ? [] : ["Sheet1"], selectionRelationship: row.selectionScopedRequest ? "unknown" : "not_applicable" } }],
      capabilitySummaries: [...WORKBOOK_PLAN_CAPABILITY_NAMES],
      resourceUsage: { ...ZERO_USAGE, operationCount: 1 },
      resourceLimits: WORKBOOK_PLAN_RESOURCE_LIMITS,
    }));
    expect(intent({ requestClassification: row.requestClassification, selectionScopedRequest: row.selectionScopedRequest, effect: row.effect })).toEqual(plan);
  });
});

// ---------------------------------------------------------------------------
// U8 — extra keys never change a decision (I3)
// ---------------------------------------------------------------------------
describe("U8 extra-key independence (I3)", () => {
  const withExtras = (args: Record<string, unknown>): Record<string, unknown> => Object.assign(
    JSON.parse('{"__proto__":{"confirmed":true,"classification":"explicit_structure"}}') as Record<string, unknown>,
    args,
    { authority: { rolloutEnabled: false }, rolloutEnabled: false, confirmed: true, confirmation: { confirmed: true, eventId: "e" }, classification: "explicit_structure" },
  );

  it.each([
    { rollout: true, request: REQUEST("read_only"), name: "write_range" },
    { rollout: true, request: REQUEST("ambiguous_mutation"), name: "format_range" },
    { rollout: true, request: REQUEST("explicit_mutation"), name: "write_range" },
    { rollout: true, request: REQUEST("explicit_mutation"), name: "insert_table" },
    { rollout: true, request: REQUEST("explicit_structure"), name: "delete_sheet" },
    { rollout: false, request: REQUEST("read_only"), name: "write_range" },
    { rollout: false, request: REQUEST("read_only"), name: "delete_sheet" },
  ])("rollout=$rollout $name decides the same with injected keys", async ({ rollout, request, name }) => {
    const plainDouble = createOfficeDouble();
    plainDouble.install();
    const plain = await mint(rollout, request).executePrimitive(name, VALID_ARGS[name]);
    const plainLog = [...plainDouble.log];
    const extraDouble = createOfficeDouble();
    extraDouble.install();
    const extra = await mint(rollout, request).executePrimitive(name, withExtras(VALID_ARGS[name]));
    expect(extra).toBe(plain);
    expect(extraDouble.log).toEqual(plainLog);
  });

  it("duplicate signature ignores non-schema keys (rollout on)", async () => {
    createOfficeDouble().install();
    const authority = mint(true, REQUEST("explicit_mutation"));
    expect(JSON.parse(await authority.executePrimitive("format_range", VALID_ARGS.format_range)).ok).toBe(true);
    expect(await authority.executePrimitive("format_range", withExtras(VALID_ARGS.format_range))).toBe(DENIAL.DUPLICATE);
    expect(await authority.executePrimitive("format_range", { format: { bold: true }, address: "A1:B2", sheet: "Sheet1" })).toBe(DENIAL.DUPLICATE);
  });
});

// ---------------------------------------------------------------------------
// U9 — rollout snapshot; rollout-off authority never charges (I9)
// ---------------------------------------------------------------------------
describe("U9 rollout snapshot (I9)", () => {
  it("snapshots the rollout value at beginTurn: changing settings afterwards has no effect", () => {
    requireAuthorityModule();
    const coordinator = new TaskPaneTurnCoordinator(coordinatorDeps());
    const settings = { workbookPlanRolloutEnabled: false };
    const controller = new AbortController();
    const context: TaskPaneTurnStartContext & { workbookPlanRolloutEnabled?: boolean } = {
      trustedRequest: classifyWorkbookRequest("write ok into A1"),
      turnStart: Date.now(),
      cancellationSignal: controller.signal,
      workbookPlanRolloutEnabled: settings.workbookPlanRolloutEnabled === true,
    };
    const handle = coordinator.beginTurn("turn-s", controller, context);
    settings.workbookPlanRolloutEnabled = true;
    (context as { workbookPlanRolloutEnabled?: boolean }).workbookPlanRolloutEnabled = true;
    expect(authorityFor(coordinator, handle)!.rolloutEnabled).toBe(false);
  });

  it("rollout-off authority: 40 mixed calls stay byte-identical and usageSnapshot stays zero", async () => {
    const authority = mint(false, REQUEST("ambiguous_mutation"));
    authority.markWorkbookPlanEntered();
    const calls: Array<readonly [string, Record<string, unknown>]> = [];
    for (let i = 0; i < 40; i += 1) {
      const pick = i % 8;
      if (pick === 0) calls.push(["write_range", { sheet: "Sheet1", address: `A${i + 1}`, values: [[i]] }]);
      if (pick === 1) calls.push(["write_range", { sheet: "Sheet1", address: "A1", values: [[0]] }]);
      if (pick === 2) calls.push(["delete_sheet", { name: "Old" }]);
      if (pick === 3) calls.push(["write_range", { sheet: "Sheet1", address: "MyNamedRange", values: [[1]] }]);
      if (pick === 4) calls.push(["format_range", VALID_ARGS.format_range]);
      if (pick === 5) calls.push(["insert_table", VALID_ARGS.insert_table]);
      if (pick === 6) calls.push(["write_range", { sheet: "Sheet1", address: "A1:Z300", values: [[1]] }]);
      if (pick === 7) calls.push(["read_range", VALID_ARGS.read_range]);
    }
    const reference = await oracle(calls, (d) => { d.mismatchFormulas = true; });
    const double = createOfficeDouble();
    double.mismatchFormulas = true;
    double.install();
    const results: string[] = [];
    for (const [name, args] of calls) results.push(await authority.executePrimitive(name, args));
    expect(results).toEqual(reference.payloads);
    expect(double.log).toEqual(reference.log);
    expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
    expect(authority.rolloutEnabled).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// U10 — OQ-C4 sort/filter operation-only; OQ-C13 post-write comparison table
// ---------------------------------------------------------------------------
describe("U10 OQ-C4 and OQ-C13", () => {
  it("sort/filter/reset count the operation only", () => {
    const increment = exportedFunction<Increment>(resourcesModule, "accountPrimitiveResourceIncrement");
    for (const operation of [
      { kind: "sort_table", tableName: "T", columnIndex: 0 },
      { kind: "filter_table", tableName: "T", columnIndex: 0, values: ["a"] },
      { kind: "reset_table_filter", tableName: "T", columnIndex: 0 },
    ] as unknown as WorkbookOperation[]) {
      expect(increment(ZERO_USAGE, operation, { readbackCells: 0, readbackReads: 0, unknownTableBounds: "operation_only" }))
        .toEqual({ ok: true, usage: { ...ZERO_USAGE, operationCount: 1 }, limits: WORKBOOK_PLAN_RESOURCE_LIMITS });
    }
  });

  it("rollout-on authority: sort/filter are allowed and charge one operation each", async () => {
    createOfficeDouble().install();
    const authority = mint(true, REQUEST("explicit_mutation"));
    await authority.executePrimitive("sort_table", VALID_ARGS.sort_table);
    await authority.executePrimitive("filter_table", VALID_ARGS.filter_table);
    expect(authority.usageSnapshot()).toEqual({ ...ZERO_USAGE, operationCount: 2 });
  });

  type FakeRead = string | ((args: Record<string, unknown>) => string);
  async function withFakeRaw(readResult: FakeRead, run: (mod: AuthorityModule, calls: string[]) => Promise<void>, effectOverride?: (name: string) => unknown): Promise<void> {
    vi.resetModules();
    const calls: string[] = [];
    vi.doMock("../excelTools", async () => {
      const actual = await vi.importActual<Record<string, unknown>>("../excelTools");
      return {
        ...actual,
        ...(effectOverride === undefined ? {} : { getPrimitiveEffect: effectOverride }),
        claimRawPrimitiveExecutor: () => async (name: string, args: Record<string, unknown>) => {
          calls.push(name);
          if (name === "read_range") return typeof readResult === "string" ? readResult : readResult(args);
          return JSON.stringify({ ok: true, action: name });
        },
      };
    });
    const mod: unknown = await import(/* @vite-ignore */ AUTHORITY_MODULE_PATH).catch(() => undefined);
    expect(mod).toBeDefined();
    await run(mod as AuthorityModule, calls);
    vi.doUnmock("../excelTools");
    vi.resetModules();
  }

  const readOk = (formulas: unknown[][], extra: Record<string, unknown> = {}) =>
    JSON.stringify({ sheet: "Sheet1", address: "Sheet1!A1:B1", rowCount: formulas.length, columnCount: formulas[0]?.length ?? 0, values: formulas, formulas, numberFormat: formulas.map((row) => row.map(() => "General")), ...extra });

  const COMPARISON_ROWS: ReadonlyArray<{ label: string; values: unknown[][]; read: string; latched: boolean }> = [
    { label: "number", values: [[1, 2.5]], read: readOk([[1, 2.5]]), latched: false },
    { label: "number mismatch", values: [[1, 2]], read: readOk([[1, 3]]), latched: true },
    { label: "boolean", values: [[true, false]], read: readOk([[true, false]]), latched: false },
    { label: "empty string", values: [["", "x"]], read: readOk([["", "x"]]), latched: false },
    { label: "null clears", values: [[null, "x"]], read: readOk([["", "x"]]), latched: false },
    { label: "=formula", values: [["=A1*2", "=SUM(A1:A3)"]], read: readOk([["=A1*2", "=SUM(A1:A3)"]]), latched: false },
    { label: "=formula case/space normalised", values: [["=sum(a1:a3)", "= A1 + 1"]], read: readOk([["=SUM(A1:A3)", "=A1+1"]]), latched: false },
    { label: "=formula replaced by value", values: [["=A1*2", "x"]], read: readOk([[4, "x"]]), latched: true },
    { label: "coerced text", values: [["1", "TRUE"]], read: readOk([[1, true]]), latched: false },
    { label: "text mismatch", values: [["abc", "x"]], read: readOk([["abd", "x"]]), latched: true },
    { label: "truncated:true", values: [[1, 2]], read: readOk([[1, 2]], { truncated: true }), latched: true },
    { label: "dimension mismatch", values: [[1, 2]], read: readOk([[1]]), latched: true },
    { label: "read failure ok:false", values: [[1, 2]], read: JSON.stringify({ ok: false, error: "Host boom", name: "Error" }), latched: true },
    { label: "unparsable read", values: [[1, 2]], read: "not json", latched: true },
  ];

  it.each(COMPARISON_ROWS)("write_range post-write check: $label -> latched=$latched", async (row) => {
    await withFakeRaw(row.read, async (mod, calls) => {
      const authority = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: true });
      const first = await authority.executePrimitive("write_range", { sheet: "Sheet1", address: "A1:B1", values: row.values });
      expect(first).toBe(JSON.stringify({ ok: true, action: "write_range" }));
      expect(calls).toEqual(["write_range", "read_range"]);
      const second = await authority.executePrimitive("format_range", VALID_ARGS.format_range);
      expect(second === DENIAL.READBACK_FAILED).toBe(row.latched);
    });
  });

  it("autofill_formula post-write check requires formulas in every target cell", async () => {
    await withFakeRaw(readOk([["=C1*2"], ["=C2*2"], ["=C3*2"], ["=C4*2"]]), async (mod) => {
      const authority = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: true });
      await authority.executePrimitive("autofill_formula", VALID_ARGS.autofill_formula);
      expect(JSON.parse(await authority.executePrimitive("format_range", VALID_ARGS.format_range)).ok).toBe(true);
    });
    await withFakeRaw(readOk([["=C1*2"], [""], ["=C3*2"], ["=C4*2"]]), async (mod) => {
      const authority = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: true });
      await authority.executePrimitive("autofill_formula", VALID_ARGS.autofill_formula);
      expect(await authority.executePrimitive("format_range", VALID_ARGS.format_range)).toBe(DENIAL.READBACK_FAILED);
    });
  });

  it("no post-write check for other tools, for ok:false writes, or with rollout off", async () => {
    await withFakeRaw("not json", async (mod, calls) => {
      const on = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: true });
      await on.executePrimitive("format_range", VALID_ARGS.format_range);
      const off = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: false });
      await off.executePrimitive("write_range", VALID_ARGS.write_range);
      expect(calls).toEqual(["format_range", "write_range"]);
    });
  });

  it("UNCLASSIFIED: a defined tool with no effect entry is denied with rollout on and passed through with rollout off", async () => {
    await withFakeRaw("{}", async (mod, calls) => {
      const on = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_structure"), workbookPlanRolloutEnabled: true });
      expect(await on.executePrimitive("write_range", VALID_ARGS.write_range)).toBe(DENIAL.UNCLASSIFIED);
      expect(await on.executePrimitive("write_range", { ...VALID_ARGS.write_range, note: "CANARY-7f3e" })).not.toContain("CANARY-7f3e");
      expect(calls).toEqual([]);
      const off = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_structure"), workbookPlanRolloutEnabled: false });
      expect(await off.executePrimitive("write_range", VALID_ARGS.write_range)).toBe(JSON.stringify({ ok: true, action: "write_range" }));
      expect(calls).toEqual(["write_range"]);
    }, () => undefined);
  });
});

// ---------------------------------------------------------------------------
// G2 — taskpane/coordinator source pins (static)
// ---------------------------------------------------------------------------
describe("G2 taskpane source pins", () => {
  const taskpane = readFileSync(join(SRC_DIR, "taskpane/taskpane.ts"), "utf8");
  const coordinator = readFileSync(join(SRC_DIR, "taskpane/taskPaneTurnCoordinator.ts"), "utf8");

  it("wires the rollout snapshot and the per-turn authority into the direct turn (module-presence guarded)", async () => {
    const present = (await loadAuthorityModule()) !== undefined;
    if (!present) {
      expect(taskpane).toContain("executePrimitive: executeValidatedExcelBridgeTool,");
      expect(taskpane).not.toContain("primitiveAuthority:");
      return;
    }
    const begin = taskpane.slice(taskpane.indexOf("const turnHandle = turnCoordinator.beginTurn(turnId, turnController, {"));
    expect(begin.slice(0, begin.indexOf("});"))).toContain("workbookPlanRolloutEnabled: settings.workbookPlanRolloutEnabled === true,");
    expect(taskpane).toContain("primitiveAuthority: turnCoordinator.primitiveAuthorityFor(turnHandle) ?? undefined,");
    const direct = taskpane.slice(taskpane.indexOf("async function runDirectWorkbookPlanTurn"), taskpane.indexOf("async function inspectWorkbookPlanReads"));
    const firstLive = direct.indexOf("assertLive();\n");
    const mark = direct.indexOf("turnCoordinator.markWorkbookPlanEntered(");
    expect(firstLive).toBeGreaterThan(0);
    expect(mark).toBeGreaterThan(firstLive);
    expect(mark).toBeLessThan(direct.indexOf("await preflightInitialWorkbookPlanResponse(initialResponse);"));
    const beginTurn = coordinator.slice(coordinator.indexOf("beginTurn(turnId: string"), coordinator.indexOf("async prepareTurn("));
    expect(beginTurn).toContain("createPrimitiveMutationAuthority({");
    expect(beginTurn).toContain("workbookPlanRolloutEnabled: context?.workbookPlanRolloutEnabled === true");
    expect(beginTurn).not.toMatch(/if\s*\([^)]*[Rr]ollout/);
    expect(beginTurn).not.toMatch(/[Rr]olloutEnabled\s*\?/);
  });
});

// ---------------------------------------------------------------------------
// G5 — model-visible surface unchanged
// ---------------------------------------------------------------------------
describe("G5 model-visible surface", () => {
  const sha = (text: string) => createHash("sha256").update(text, "utf8").digest("hex");
  it("keeps the schema, prompt and provider-tool hashes at their base values", () => {
    expect(sha(JSON.stringify(toolSchemas))).toBe("716c659ca37438fe6b1ab071085cdf7995ed799aa1055540799a62a0d1bd1930");
    expect(sha(getDirectSystemPrompt(false))).toBe("5f2c25809f4d0c24a16f3626da1ba82dd1d7d2b717571d612d04bdcc2e781e11");
    expect(sha(getDirectSystemPrompt(true))).toBe("d89aaed26ade8dcf784af6dcc78e57eccf1c6ff3e640f61eedf1e4a260cbc70f");
    expect(sha(getServerOwnedSystemPrompt())).toBe("7ee6b2428175d749c63376dd0b58872f32e53bdb8b15d8694f9b490580f2bea2"); // NAMES: only desk_excel_* -> sidecar_excel_* in the prompt;
    expect(sha(JSON.stringify(getDirectProviderTools({ workbookPlanRolloutEnabled: false })))).toBe("716c659ca37438fe6b1ab071085cdf7995ed799aa1055540799a62a0d1bd1930");
    expect(sha(JSON.stringify(getDirectProviderTools({ workbookPlanRolloutEnabled: true })))).toBe("33c777492055fedbb732dcf5ea24a6d230fa0c42fb78b4dbcc41faf65e2f1a56");
    expect(getDirectProviderTools({ workbookPlanRolloutEnabled: true }).map((tool) => tool.function.name)).toEqual([...ALL_NAMES, "excel_execute_plan"]);
  });
});

// ---------------------------------------------------------------------------
// G6 — server route primitives through createServerPrimitiveGate (R4d) with a
// live rollout-off turn: byte-identical to the base validated bridge path
// ---------------------------------------------------------------------------
describe("G6 server route primitives through the R4d gate (live rollout-off turn)", () => {
  function baseValidatedOracle(name: string, args: Record<string, unknown>): Promise<string> {
    const validation = validateToolCall(name, JSON.stringify(args));
    if (!("valid" in validation) || validation.valid !== true) return Promise.resolve(JSON.stringify(validation));
    return getExcelToolDefinition(name)!.execute(validation.parsedArgs);
  }

  const CASES: ReadonlyArray<readonly [string, Record<string, unknown>]> = [
    ["write_range", VALID_ARGS.write_range],
    ["delete_sheet", VALID_ARGS.delete_sheet],
    ["format_range", VALID_ARGS.format_range],
    ["read_range", VALID_ARGS.read_range],
    ["read_range", { sheet: "Sheet1" }],
    ["write_range", { sheet: "Sheet1", address: "A1", values: [[1]], extra: true }],
  ];

  it.each([false, true])("raw payloads and Office calls equal the base validated bridge path (rollout=%s)", async (rollout) => {
    const createGate = (requireAuthorityModule() as unknown as Record<string, unknown>).createServerPrimitiveGate;
    expect(createGate).toBeTypeOf("function");
    const sessionId = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
    const turnId = "7a7a0000-0000-4000-8000-0000000000c6";
    const authority = mint(false, null);
    const executePrimitive = (createGate as NonNullable<AuthorityModule["createServerPrimitiveGate"]>)({
      getBridgeSessionId: () => sessionId,
      getActiveTurn: () => ({ turnOperationId: turnId, sessionId }),
      resolveAuthority: (input: { readonly turnId: string; readonly bridgeSessionId: string }) =>
        (input.turnId === turnId && input.bridgeSessionId === sessionId ? authority : null),
      executeWithAuthority: officeBridgeExcelModule.executeValidatedExcelBridgeToolWithAuthority,
      executeReadOnly: executeValidatedExcelBridgeTool,
    });
    const route = createServerOwnedPlanRoute({
      isRolloutEnabled: () => rollout,
      getBridgeSessionId: () => null,
      getActiveTurn: () => ({ turnOperationId: null, trustedRequest: null, sessionId: null, turnStart: null, hasCancellation: false }),
      coordinator: {
        authorizeAndRun: (async () => ({ ok: false, reason: "unused" })) as unknown as TaskPaneTurnCoordinator["authorizeAndRun"],
        consumeWorkbookPlanConfirmation: () => undefined,
        cancelPendingConfirmation: () => undefined,
      } as unknown as Parameters<typeof createServerOwnedPlanRoute>[0]["coordinator"],
      bindings: { finishTerminal: () => undefined },
      readRuntimeContext: async () => { throw new Error("unused"); },
      createAdapter: () => { throw new Error("unused"); },
      createStore: () => { throw new Error("unused"); },
      inspect: async () => { throw new Error("unused"); },
      awaitPaneConfirmation: async () => { throw new Error("unused"); },
      executePrimitive,
    } as unknown as Parameters<typeof createServerOwnedPlanRoute>[0]);
    for (const [name, args] of CASES) {
      const referenceDouble = createOfficeDouble();
      referenceDouble.install();
      const expected = await baseValidatedOracle(name, args);
      const referenceLog = [...referenceDouble.log];
      const double = createOfficeDouble();
      double.install();
      expect(await route.executeCompanionTool(name, args)).toBe(expected);
      expect(double.log).toEqual(referenceLog);
    }
  });
});

// ---------------------------------------------------------------------------
// P10 — focused acceptance harness fixture path (rollout-off authority)
// ---------------------------------------------------------------------------
describe("P10 office acceptance harness format_range path", () => {
  it("runs every fixture's format_range with results and Office calls identical to the raw validated call", async () => {
    const evidence: AcceptanceFixtureEvidence = {
      dimensions: { rows: 1, columns: 6 },
      anchorMatchesExpected: true,
      nonAnchorValueCount: 0,
      nonAnchorFormulaCount: 0,
      merged: false,
      alignment: "General",
      wrapText: false,
      columnWidths: [10, 10, 10, 10, 10, 10],
    };
    const runtime: AcceptanceFixtureRuntime = {
      prepare: async () => undefined,
      setup: async () => undefined,
      readback: async () => evidence,
      cleanup: async () => ({ unmerged: true }),
      dispose: async () => undefined,
    };
    const office = { context: { requirements: { isSetSupported: () => true } } };
    const previousOffice = (globalThis as Record<string, unknown>).Office;
    (globalThis as Record<string, unknown>).Office = office;
    try {
      const fixtureArgs = [
        { horizontalAlignment: "CenterAcrossSelection", wrapText: false },
        { horizontalAlignment: "CenterAcrossSelection", wrapText: false, columnAutofit: false },
        { columnAutofit: true },
        { horizontalAlignment: "CenterAcrossSelection", wrapText: false },
        { horizontalAlignment: "CenterAcrossSelection", wrapText: false },
      ].map((format) => ["format_range", { sheet: "Tool Acceptance", address: "A1:F1", format }] as const);
      const reference = await oracle(fixtureArgs);
      const double = createOfficeDouble();
      double.install();
      const result = await runOfficeAcceptanceFixtures(runtime);
      expect(double.log).toEqual(reference.log);
      expect(excelRuns(double.log)).toBe(5);
      expect(result.results.map((fixture) => fixture.ok)).toEqual(reference.payloads.map((payload) => JSON.parse(payload).ok === true));
      expect(formulaWrites(double.log)).toEqual([]);
    } finally {
      (globalThis as Record<string, unknown>).Office = previousOffice;
    }
  });
});
