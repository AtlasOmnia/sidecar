/**
 * R4e-2 authority tests, file 3 (dev3-r4e2-release §5: P1-P9).
 *
 * Rollout ON: a failed primitive (throw or non-ok:true result) is refunded only
 * when no Office batch began during its raw call; otherwise the charge and
 * signature are kept and the turn latches READBACK_FAILED (R4d B-A1 / A P3-A5).
 * Rollout OFF and the server-route gate are unchanged (P8, P8s).
 *
 * The Office double and `withFakeRaw` are copied from
 * primitive-authority-refund-r4d.test.ts (not imported) with two additions:
 * `fault = "sync1" | "sync2"` makes the n-th `context.sync` of the current
 * `Excel.run` throw a RichApi.Error-shaped error after logging `sync-fail`,
 * and `withFakeRaw` hands the factory's `actual` excelTools module to the fake
 * (so a fake can enter a real `safelyExcel`). The batch counter is reached only
 * through a dynamic import, so this file type-checks at the base and RED is
 * assertion-only. Expected usage comes from a fresh oracle authority on a
 * healthy double; expected payloads from the raw oracle.
 */
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { getExcelToolDefinition } from "../excelTools";
import { createPrimitiveMutationAuthority, createServerPrimitiveGate, type PrimitiveMutationAuthority } from "../primitiveMutationAuthority";
import type { TrustedWorkbookRequest } from "../workbookPlanRequest";
import type { WorkbookRequestClassification } from "../workbookPlanPolicy";

// ---------------------------------------------------------------------------
// Recording Office double (copied from R4d F-rows, plus sync faults).
// ---------------------------------------------------------------------------
type Fault = "ok" | "run" | "sync1" | "sync2";

interface OfficeDouble {
  readonly log: string[];
  mismatchFormulas: boolean;
  fault: Fault;
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
  const state: { mismatchFormulas: boolean; fault: Fault; syncInRun: number } = { mismatchFormulas: false, fault: "ok", syncInRun: 0 };

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
    sync: async (): Promise<void> => {
      state.syncInRun += 1;
      const failAt = state.fault === "sync1" ? 1 : state.fault === "sync2" ? 2 : 0;
      if (state.syncInRun === failAt) {
        log.push("sync-fail");
        throw Object.assign(new Error("Sync boom r4e2"), { code: "GeneralException", name: "RichApi.Error" });
      }
      log.push("sync");
    },
  };

  return {
    log,
    get mismatchFormulas() { return state.mismatchFormulas; },
    set mismatchFormulas(value: boolean) { state.mismatchFormulas = value; },
    get fault() { return state.fault; },
    set fault(value: Fault) { state.fault = value; },
    install() {
      (globalThis as Record<string, unknown>).Excel = {
        run: async (work: (ctx: unknown) => Promise<unknown>) => {
          log.push("Excel.run");
          if (state.fault === "run") throw Object.assign(new Error("Host boom"), { code: "GeneralException", name: "RichApi.Error" });
          state.syncInRun = 0;
          return work(context);
        },
      };
    },
  };
}

function excelRuns(log: readonly string[]): number {
  return log.filter((entry) => entry === "Excel.run").length;
}

const matching = (log: readonly string[], pattern: RegExp): string[] => log.filter((entry) => pattern.test(entry));

// ---------------------------------------------------------------------------
// Fixed expectations (independent constants)
// ---------------------------------------------------------------------------
const deny = (error: string, category: string): string => JSON.stringify({ ok: false, error, category });
const DENIAL = {
  DUPLICATE: deny("Workbook edit denied: the identical edit was already applied in this turn.", "repetition"),
  READBACK_FAILED: deny("Workbook edit denied: an earlier edit in this turn did not verify.", "validation"),
  TURN_INACTIVE: deny("Workbook edit denied: no active task-pane turn.", "validation"),
} as const;

const ZERO_USAGE = {
  operationCount: 0, cellsRead: 0, cellsWritten: 0, formulaUtf8Bytes: 0, csvPayloadUtf8Bytes: 0,
  commentUtf8Bytes: 0, sheetsCreated: 0, chartsCreated: 0, tablesCreated: 0, verificationReads: 0,
};

const REQUEST = (classification: WorkbookRequestClassification, selectionScopedRequest = false): TrustedWorkbookRequest =>
  Object.freeze({ classification, selectionScopedRequest });

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

type Call = readonly [string, Record<string, unknown>];
type Step = readonly [string, Record<string, unknown>, Fault];

const BOLD: Call = ["format_range", { sheet: "Sheet1", address: "A1:B2", format: { bold: true } }];
/** Reversed corners: not a strict A1 for excelTools (load sync first), chargeable for the authority. */
const NUMBER_FORMAT_TWO_SYNC: Call = ["format_range", { sheet: "Sheet1", address: "B2:A1", format: { numberFormat: "0.00" } }];
const oneCell = (row: number): Call => ["write_range", { sheet: "Sheet1", address: `A${row}`, values: [[row]] }];

const FIXED_NOW = 1_800_000_000_000;
beforeEach(() => {
  vi.spyOn(Date, "now").mockReturnValue(FIXED_NOW);
});
afterEach(() => {
  vi.restoreAllMocks();
  delete (globalThis as Record<string, unknown>).Excel;
  vi.doUnmock("../excelTools");
});

/**
 * Raw oracle: each step through getExcelToolDefinition(name).execute on a fresh double. An
 * argument-validation throw is rendered as the raw dispatcher's documented `{ok:false,error,name}`.
 */
async function rawOracle(steps: readonly Step[]): Promise<{ payloads: string[]; log: string[] }> {
  const double = createOfficeDouble();
  double.install();
  const payloads: string[] = [];
  for (const [name, args, fault] of steps) {
    double.fault = fault;
    payloads.push(await Promise.resolve().then(() => getExcelToolDefinition(name)!.execute(args)).catch((error: Error) =>
      JSON.stringify({ ok: false, error: error.message, name: error.name })));
  }
  double.fault = "ok";
  return { payloads, log: [...double.log] };
}

const explicitAuthority = (rollout = true) =>
  createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: rollout });

/** Usage oracle: a fresh rollout-on authority runs the calls on a healthy double. */
async function oracleUsage(calls: readonly Call[]): Promise<unknown> {
  const double = createOfficeDouble();
  double.install();
  const authority = explicitAuthority();
  for (const [name, args] of calls) expect(JSON.parse(await authority.executePrimitive(name, args)).ok).toBe(true);
  delete (globalThis as Record<string, unknown>).Excel;
  return authority.usageSnapshot();
}

function gateFor(authority: PrimitiveMutationAuthority | null, live = true): (name: string, args: Record<string, unknown>) => Promise<string> {
  return createServerPrimitiveGate({
    getBridgeSessionId: () => "session-r4e2",
    getActiveTurn: () => (live ? { turnOperationId: "turn-r4e2", sessionId: "session-r4e2" } : { turnOperationId: null, sessionId: null }),
    resolveAuthority: () => authority,
    executeWithAuthority: (name, args, resolved) => resolved.executePrimitive(name, args),
    executeReadOnly: (name, args) => getExcelToolDefinition(name)!.execute(args),
  });
}

// ---------------------------------------------------------------------------
// withFakeRaw (copied R4c/R4d pattern) + the factory's actual module.
// ---------------------------------------------------------------------------
type ActualExcelTools = Record<string, unknown> & { getExcelToolDefinition: typeof getExcelToolDefinition };
type FakeRaw = (name: string, args: Record<string, unknown>, calls: string[], actual: ActualExcelTools) => Promise<string>;
interface AuthorityModule {
  createPrimitiveMutationAuthority: typeof createPrimitiveMutationAuthority;
}
const AUTHORITY_MODULE_PATH = "../primitiveMutationAuthority";

async function withFakeRaw(fake: FakeRaw, run: (mod: AuthorityModule, calls: string[]) => Promise<void>): Promise<void> {
  vi.resetModules();
  const calls: string[] = [];
  vi.doMock("../excelTools", async () => {
    const actual = await vi.importActual<ActualExcelTools>("../excelTools");
    return {
      ...actual,
      claimRawPrimitiveExecutor: () => async (name: string, args: Record<string, unknown>) => {
        calls.push(name);
        return fake(name, args, calls, actual);
      },
    };
  });
  const mod: unknown = await import(/* @vite-ignore */ AUTHORITY_MODULE_PATH).catch(() => undefined);
  expect(mod).toBeDefined();
  await run(mod as AuthorityModule, calls);
  vi.doUnmock("../excelTools");
  vi.resetModules();
}

const readOk = (formulas: unknown[][]) =>
  JSON.stringify({ sheet: "Sheet1", address: "Sheet1!A1", rowCount: formulas.length, columnCount: formulas[0]?.length ?? 0, values: formulas, formulas, numberFormat: formulas.map((row) => row.map(() => "General")) });

// ---------------------------------------------------------------------------
// §5.1 RED rows (rollout ON, explicit_mutation)
// ---------------------------------------------------------------------------
describe("R4e-2 P rows: a failed primitive that began an Office batch keeps its charge and latches", () => {
  it("P1 format_range whose apply sync fails: base ok:false payload, oracle charge kept, identical retry READBACK_FAILED with no batch, one set of format writes", async () => {
    const reference = await rawOracle([[BOLD[0], BOLD[1], "sync1"]]);
    expect(JSON.parse(reference.payloads[0]).ok).toBe(false);
    const expectedUsage = await oracleUsage([BOLD]);
    expect(expectedUsage).not.toEqual(ZERO_USAGE);
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority();
    double.fault = "sync1";
    const failed = await authority.executePrimitive(BOLD[0], BOLD[1]);
    double.fault = "ok";
    expect(failed).toBe(reference.payloads[0]);
    expect(matching(double.log, /^sync-fail$/)).toHaveLength(1);
    expect(authority.usageSnapshot()).toEqual(expectedUsage);
    const runs = excelRuns(double.log);
    expect(await authority.executePrimitive(BOLD[0], BOLD[1])).toBe(DENIAL.READBACK_FAILED);
    expect(excelRuns(double.log) - runs).toBe(0);
    expect(matching(double.log, /\.format\.font\.bold=/)).toHaveLength(1);
  });

  it("P2 B-A1 shape: format_range whose load sync succeeds and apply sync fails is charged; its retry and a different format_range are READBACK_FAILED; one apply batch", async () => {
    const [name, args] = NUMBER_FORMAT_TWO_SYNC;
    const reference = await rawOracle([[name, args, "sync2"]]);
    expect(JSON.parse(reference.payloads[0]).ok).toBe(false);
    expect(reference.log.filter((entry) => entry === "sync" || entry === "sync-fail")).toEqual(["sync", "sync-fail"]);
    const expectedUsage = await oracleUsage([NUMBER_FORMAT_TWO_SYNC]);
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority();
    double.fault = "sync2";
    expect(await authority.executePrimitive(name, args)).toBe(reference.payloads[0]);
    double.fault = "ok";
    expect(authority.usageSnapshot()).toEqual(expectedUsage);
    const runs = excelRuns(double.log);
    expect(await authority.executePrimitive(name, args)).toBe(DENIAL.READBACK_FAILED);
    expect(await authority.executePrimitive(BOLD[0], BOLD[1])).toBe(DENIAL.READBACK_FAILED);
    expect(excelRuns(double.log) - runs).toBe(0);
    expect(matching(double.log, /\.numberFormat=/)).toHaveLength(1);
    expect(authority.usageSnapshot()).toEqual(expectedUsage);
  });

  it("P3 write_range whose sync fails keeps its charge incl. the pre-charged readback; no readback batch; next distinct write READBACK_FAILED; reads and activate_sheet still run", async () => {
    const write = oneCell(1);
    const reference = await rawOracle([[write[0], write[1], "sync1"], ["read_range", VALID_ARGS.read_range, "ok"], ["activate_sheet", VALID_ARGS.activate_sheet, "ok"]]);
    const expectedUsage = await oracleUsage([write]);
    expect(expectedUsage).toMatchObject({ verificationReads: 1 });
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority();
    double.fault = "sync1";
    expect(await authority.executePrimitive(write[0], write[1])).toBe(reference.payloads[0]);
    double.fault = "ok";
    expect(excelRuns(double.log)).toBe(1);
    expect(authority.usageSnapshot()).toEqual(expectedUsage);
    const runs = excelRuns(double.log);
    expect(await authority.executePrimitive(...oneCell(2))).toBe(DENIAL.READBACK_FAILED);
    expect(excelRuns(double.log) - runs).toBe(0);
    expect(await authority.executePrimitive("read_range", VALID_ARGS.read_range)).toBe(reference.payloads[1]);
    expect(await authority.executePrimitive("activate_sheet", VALID_ARGS.activate_sheet)).toBe(reference.payloads[2]);
    expect(excelRuns(double.log) - runs).toBe(2);
  });

  it("P4 a raw call that enters a real Office batch and then throws rejects with the identical error, keeps its charge and latches", async () => {
    const write = oneCell(1);
    const expectedUsage = await oracleUsage([write]);
    const boom = new Error("raw boom after batch r4e2");
    const double = createOfficeDouble();
    double.install();
    await withFakeRaw(async (name, args, _calls, actual) => {
      if (name === "write_range") {
        await actual.getExcelToolDefinition("read_range")!.execute({ sheet: "Sheet1", address: "A1" });
        throw boom;
      }
      if (name === "read_range") return readOk([[Number(String(args.address).slice(1))]]);
      return JSON.stringify({ ok: true, action: name });
    }, async (mod, calls) => {
      const authority = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: true });
      const rejected = await authority.executePrimitive(write[0], write[1]).then(() => undefined, (error: unknown) => error);
      expect(rejected).toBe(boom);
      expect(excelRuns(double.log)).toBe(1);
      expect(authority.usageSnapshot()).toEqual(expectedUsage);
      expect(await authority.executePrimitive(...oneCell(2))).toBe(DENIAL.READBACK_FAILED);
      expect(await authority.executePrimitive(write[0], write[1])).toBe(DENIAL.READBACK_FAILED);
      expect(calls).toEqual(["write_range"]);
    });
  });

  it("P5 a raw call that enters a real Office batch and returns ok:false is returned unchanged, keeps its charge and latches (conservative)", async () => {
    const write = oneCell(1);
    const expectedUsage = await oracleUsage([write]);
    const hostFailed = JSON.stringify({ ok: false, error: "host failed", category: "Execution" });
    const double = createOfficeDouble();
    double.install();
    await withFakeRaw(async (name, args, _calls, actual) => {
      if (name === "write_range") {
        await actual.getExcelToolDefinition("read_range")!.execute({ sheet: "Sheet1", address: "A1" });
        return hostFailed;
      }
      if (name === "read_range") return readOk([[Number(String(args.address).slice(1))]]);
      return JSON.stringify({ ok: true, action: name });
    }, async (mod, calls) => {
      const authority = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: true });
      expect(await authority.executePrimitive(write[0], write[1])).toBe(hostFailed);
      expect(authority.usageSnapshot()).toEqual(expectedUsage);
      expect(await authority.executePrimitive(...oneCell(3))).toBe(DENIAL.READBACK_FAILED);
      expect(calls).toEqual(["write_range"]);
    });
  });

  it("P6 server route: through createServerPrimitiveGate (live turn, rollout on) P1's failing call returns the base payload and its retry is READBACK_FAILED", async () => {
    const reference = await rawOracle([[BOLD[0], BOLD[1], "sync1"]]);
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority();
    const gate = gateFor(authority);
    double.fault = "sync1";
    expect(await gate(BOLD[0], BOLD[1])).toBe(reference.payloads[0]);
    double.fault = "ok";
    const runs = excelRuns(double.log);
    expect(await gate(BOLD[0], BOLD[1])).toBe(DENIAL.READBACK_FAILED);
    expect(excelRuns(double.log) - runs).toBe(0);
  });

  it("P9 getOfficeBatchEntryCount counts Excel.run callback entries only (+1 per entered batch, +0 when Excel.run rejects first); excelTools.ts has one Excel.run( inside safelyExcel", async () => {
    const tools = (await import("../excelTools")) as Record<string, unknown> & { getExcelToolDefinition: typeof getExcelToolDefinition };
    const count = tools.getOfficeBatchEntryCount;
    expect(typeof count).toBe("function");
    const read = count as () => number;
    const double = createOfficeDouble();
    double.install();
    const start = read();
    await tools.getExcelToolDefinition("read_range")!.execute({ sheet: "Sheet1", address: "A1" });
    expect(read() - start).toBe(1);
    double.fault = "run";
    await tools.getExcelToolDefinition("format_range")!.execute(BOLD[1]);
    expect(read() - start).toBe(1);
    double.fault = "sync1";
    await tools.getExcelToolDefinition("format_range")!.execute(BOLD[1]);
    expect(read() - start).toBe(2);
    double.fault = "ok";
    await Promise.resolve().then(() => tools.getExcelToolDefinition("format_range")!.execute({ sheet: "Sheet1", address: "A1:B2" })).catch(() => undefined);
    expect(read() - start).toBe(2);
    const source = readFileSync(join(fileURLToPath(new URL("..", import.meta.url)), "excelTools.ts"), "utf8");
    expect(source.match(/Excel\.run\(/g)).toHaveLength(1);
    const at = source.indexOf("async function safelyExcel(");
    const body = source.slice(at, source.indexOf("\n}\n", at));
    expect(body).toContain("Excel.run(");
  });
});

// ---------------------------------------------------------------------------
// §5.2 GREEN pins (identical at base and HEAD)
// ---------------------------------------------------------------------------
describe("R4e-2 P pins: failures that never reached Office are still refunded; rollout off unchanged", () => {
  it("P7a [pin] argument-validation failures return the raw payload, stay uncharged with no batch, and the corrected call executes", async () => {
    const missingFormat: Call = ["format_range", { sheet: "Sheet1", address: "A1:B2" }];
    const badValues: Call = ["write_range", { sheet: "Sheet1", address: "A1", values: "x" }];
    const reference = await rawOracle([[missingFormat[0], missingFormat[1], "ok"], [badValues[0], badValues[1], "ok"]]);
    expect(reference.payloads.map((payload) => JSON.parse(payload).ok)).toEqual([false, false]);
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority();
    expect(await authority.executePrimitive(missingFormat[0], missingFormat[1])).toBe(reference.payloads[0]);
    expect(await authority.executePrimitive(badValues[0], badValues[1])).toBe(reference.payloads[1]);
    expect(excelRuns(double.log)).toBe(0);
    expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
    expect(JSON.parse(await authority.executePrimitive(BOLD[0], BOLD[1])).ok).toBe(true);
    expect(JSON.parse(await authority.executePrimitive(...oneCell(1))).ok).toBe(true);
    expect(excelRuns(double.log)).toBe(3);
  });

  it("P7b [pin] Excel.run rejecting before its callback (R4d F2 shape) refunds format_range; the identical retry executes", async () => {
    const reference = await rawOracle([[BOLD[0], BOLD[1], "run"], [BOLD[0], BOLD[1], "ok"]]);
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority();
    double.fault = "run";
    expect(await authority.executePrimitive(BOLD[0], BOLD[1])).toBe(reference.payloads[0]);
    double.fault = "ok";
    expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
    expect(await authority.executePrimitive(BOLD[0], BOLD[1])).toBe(reference.payloads[1]);
    expect(matching(double.log, /\.format\.font\.bold=/)).toHaveLength(1);
    expect(authority.usageSnapshot()).toEqual(await oracleUsage([BOLD]));
  });

  it("P7c [pin] a fake raw ok:false that never enters Office (R4d F5 shape) is refunded; the retry executes", async () => {
    let writes = 0;
    await withFakeRaw(async (name, args) => {
      if (name === "read_range") return readOk([[Number(String(args.address).slice(1))]]);
      writes += 1;
      return writes === 1 ? JSON.stringify({ ok: false, error: "host failed", category: "Execution" }) : JSON.stringify({ ok: true, action: name });
    }, async (mod, calls) => {
      const authority = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: true });
      const write = oneCell(1);
      expect(JSON.parse(await authority.executePrimitive(write[0], write[1])).ok).toBe(false);
      expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
      expect(JSON.parse(await authority.executePrimitive(write[0], write[1])).ok).toBe(true);
      expect(calls).toEqual(["write_range", "write_range", "read_range"]);
    });
  });

  it("P7d [pin] ok:true path unchanged: a verified write keeps its charge; a mismatched readback latches (R4d F6 shape)", async () => {
    const write = oneCell(1);
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority();
    expect(JSON.parse(await authority.executePrimitive(write[0], write[1])).ok).toBe(true);
    expect(JSON.parse(await authority.executePrimitive(...oneCell(2))).ok).toBe(true);
    expect(excelRuns(double.log)).toBe(4);
    double.mismatchFormulas = true;
    expect(JSON.parse(await authority.executePrimitive(...oneCell(3))).ok).toBe(true);
    expect(await authority.executePrimitive(BOLD[0], BOLD[1])).toBe(DENIAL.READBACK_FAILED);
    expect(authority.usageSnapshot()).toEqual({ ...ZERO_USAGE, operationCount: 3, cellsWritten: 3, cellsRead: 3, verificationReads: 3 });
  });

  it("P8 [pin] rollout OFF: 27 primitives healthy and with sync faults 1/2, repeated failing calls and retries, and the C1 spellings equal the raw oracle; no charge, no latch", async () => {
    const names = Object.keys(VALID_ARGS);
    expect(names).toHaveLength(27);
    const c1 = ["A1", "a1", "$A$1", " A1 ", "Sheet1!A1", "'Sheet1'!$a$1", "A1:A1", "$A$1:a1"].map(
      (address): Step => ["write_range", { sheet: "Sheet1", address, values: [[1]] }, "ok"]);
    const steps: Step[] = [
      ...names.map((name): Step => [name, VALID_ARGS[name], "ok"]),
      ...names.map((name): Step => [name, VALID_ARGS[name], "sync1"]),
      ...names.map((name): Step => [name, VALID_ARGS[name], "sync2"]),
      [BOLD[0], BOLD[1], "sync1"], [BOLD[0], BOLD[1], "sync1"], [BOLD[0], BOLD[1], "ok"], [BOLD[0], BOLD[1], "ok"],
      [NUMBER_FORMAT_TWO_SYNC[0], NUMBER_FORMAT_TWO_SYNC[1], "sync2"], [NUMBER_FORMAT_TWO_SYNC[0], NUMBER_FORMAT_TWO_SYNC[1], "ok"],
      [...oneCell(1), "sync1"], [...oneCell(1), "ok"], [...oneCell(2), "ok"],
      ...c1,
    ];
    const reference = await rawOracle(steps);
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority(false);
    const payloads: string[] = [];
    for (const [name, args, fault] of steps) {
      double.fault = fault;
      payloads.push(await authority.executePrimitive(name, args));
    }
    double.fault = "ok";
    expect(payloads).toEqual(reference.payloads);
    expect(double.log).toEqual(reference.log);
    expect(payloads.filter((payload) => payload === DENIAL.READBACK_FAILED || payload === DENIAL.DUPLICATE)).toEqual([]);
    expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
  });

  it("P8s [pin] server gate: a rollout-OFF authority executes P1's failing call and its retry like raw; a no-turn mutating ticket is TURN_INACTIVE with no batch in both rollout modes", async () => {
    const reference = await rawOracle([[BOLD[0], BOLD[1], "sync1"], [BOLD[0], BOLD[1], "ok"]]);
    const double = createOfficeDouble();
    double.install();
    const gate = gateFor(explicitAuthority(false));
    double.fault = "sync1";
    const failed = await gate(BOLD[0], BOLD[1]);
    double.fault = "ok";
    expect([failed, await gate(BOLD[0], BOLD[1])]).toEqual(reference.payloads);
    expect(double.log).toEqual(reference.log);
    const runs = excelRuns(double.log);
    for (const rollout of [false, true]) {
      const noTurn = gateFor(explicitAuthority(rollout), false);
      expect(await noTurn(BOLD[0], BOLD[1])).toBe(DENIAL.TURN_INACTIVE);
      expect(await noTurn(...oneCell(1))).toBe(DENIAL.TURN_INACTIVE);
    }
    expect(excelRuns(double.log)).toBe(runs);
  });
});
