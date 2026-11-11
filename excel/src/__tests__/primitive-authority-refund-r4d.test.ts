/**
 * R4d authority-unit tests (dev3-r4d-release.md §5.1 F2-F6, §5.2 DP1).
 *
 * F2-F5: with rollout ON, a primitive whose raw result is not `ok:true`, or
 * whose raw call throws, refunds its per-turn budget and signature (R4c lane B
 * A2). F6 pins that a verified-write (readback) failure is NOT refunded.
 * DP1: the direct route with rollout OFF stays byte-identical to the raw
 * (`066fe36`) behaviour.
 *
 * Kept separate from primitive-authority-server.test.ts: the copied
 * `withFakeRaw` pattern uses vi.doMock("../excelTools") with vi.resetModules,
 * which must not disturb the one-shot raw-dispatcher latch of that file.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { getExcelToolDefinition } from "../excelTools";
import { createPrimitiveMutationAuthority } from "../primitiveMutationAuthority";
import type { TrustedWorkbookRequest } from "../workbookPlanRequest";
import type { WorkbookRequestClassification } from "../workbookPlanPolicy";

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
// Fixed expectations (independent constants)
// ---------------------------------------------------------------------------
const deny = (error: string, category: string): string => JSON.stringify({ ok: false, error, category });
const DENIAL = {
  DUPLICATE: deny("Workbook edit denied: the identical edit was already applied in this turn.", "repetition"),
  RESOURCE_LIMIT: deny("Workbook edit denied: the per-turn workbook edit limit was reached.", "validation"),
  READBACK_FAILED: deny("Workbook edit denied: an earlier edit in this turn did not verify.", "validation"),
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
const LEGACY_NAMES = ["delete_sheet", "clear_range", "clear_comments", "find_replace", "unprotect_sheet", "protect_sheet", "import_csv_to_sheet", "create_pivot_table"];

type Call = readonly [string, Record<string, unknown>];
/** One step: the call and whether the host fails every Excel.run during it. */
type Step = readonly [string, Record<string, unknown>, boolean];

const oneCell = (row: number): Call => ["write_range", { sheet: "Sheet1", address: `A${row}`, values: [[row]] }];

// create_pivot_table names its pivot from Date.now(): pin the clock.
const FIXED_NOW = 1_800_000_000_000;
beforeEach(() => {
  vi.spyOn(Date, "now").mockReturnValue(FIXED_NOW);
});
afterEach(() => {
  vi.restoreAllMocks();
  delete (globalThis as Record<string, unknown>).Excel;
  vi.doUnmock("../excelTools");
});

/** Raw (`066fe36`) oracle: each step through getExcelToolDefinition(name).execute on a fresh double. */
async function rawOracle(steps: readonly Step[], configure?: (double: OfficeDouble) => void): Promise<{ payloads: string[]; log: string[] }> {
  const double = createOfficeDouble();
  configure?.(double);
  double.install();
  const payloads: string[] = [];
  for (const [name, args, fail] of steps) {
    double.failRuns = fail;
    payloads.push(await getExcelToolDefinition(name)!.execute(args));
  }
  double.failRuns = false;
  return { payloads, log: [...double.log] };
}

const explicitAuthority = (rollout = true) =>
  createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: rollout });

// ---------------------------------------------------------------------------
// Copied R4c withFakeRaw pattern (primitive-mutation-authority.test.ts), with
// a scripted raw implementation. Loads a fresh authority module whose raw
// dispatcher is the fake.
// ---------------------------------------------------------------------------
type FakeRaw = (name: string, args: Record<string, unknown>, calls: string[]) => Promise<string>;
interface AuthorityModule {
  createPrimitiveMutationAuthority: typeof createPrimitiveMutationAuthority;
}
const AUTHORITY_MODULE_PATH = "../primitiveMutationAuthority";

async function withFakeRaw(fake: FakeRaw, run: (mod: AuthorityModule, calls: string[]) => Promise<void>): Promise<void> {
  vi.resetModules();
  const calls: string[] = [];
  vi.doMock("../excelTools", async () => {
    const actual = await vi.importActual<Record<string, unknown>>("../excelTools");
    return {
      ...actual,
      claimRawPrimitiveExecutor: () => async (name: string, args: Record<string, unknown>) => {
        calls.push(name);
        return fake(name, args, calls);
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
// §5.1 F2-F6 (rollout ON, explicit_mutation)
// ---------------------------------------------------------------------------
describe("F-rows: rollout-on refund of failed primitive writes (lane B A2)", () => {
  it("F2 a host-failed write_range returns the base ok:false payload, is not charged, and its identical retry executes and verifies", async () => {
    const write = oneCell(1);
    const reference = await rawOracle([[write[0], write[1], true], [write[0], write[1], false]]);
    expect(JSON.parse(reference.payloads[0]).ok).toBe(false);
    expect(JSON.parse(reference.payloads[1]).ok).toBe(true);
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority();
    double.failRuns = true;
    const failed = await authority.executePrimitive(write[0], write[1]);
    double.failRuns = false;
    expect(failed).toBe(reference.payloads[0]);
    expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
    const runsBeforeRetry = excelRuns(double.log);
    const retry = await authority.executePrimitive(write[0], write[1]);
    expect(retry).toBe(reference.payloads[1]);
    expect(excelRuns(double.log) - runsBeforeRetry).toBe(2);
    expect(formulaWrites(double.log)).toHaveLength(1);
    expect(authority.usageSnapshot()).toEqual({ ...ZERO_USAGE, operationCount: 1, cellsWritten: 1, cellsRead: 1, verificationReads: 1 });
  });

  it("F3 a throwing raw write_range rejects with the same error, is not charged, and leaves no signature", async () => {
    const boom = new Error("raw boom r4d");
    let throwWrites = true;
    await withFakeRaw(async (name) => {
      if (name === "write_range" && throwWrites) throw boom;
      if (name === "read_range") return readOk([[1]]);
      return JSON.stringify({ ok: true, action: name });
    }, async (mod, calls) => {
      const authority = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: true });
      const write = oneCell(1);
      const rejected = await authority.executePrimitive(write[0], write[1]).then(() => undefined, (error: unknown) => error);
      expect(rejected).toBe(boom);
      expect((rejected as Error).message).toBe("raw boom r4d");
      expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
      throwWrites = false;
      expect(await authority.executePrimitive(write[0], write[1])).toBe(JSON.stringify({ ok: true, action: "write_range" }));
      expect(calls).toEqual(["write_range", "write_range", "read_range"]);
      expect(authority.usageSnapshot()).toEqual({ ...ZERO_USAGE, operationCount: 1, cellsWritten: 1, cellsRead: 1, verificationReads: 1 });
    });
  });

  it("F4 five host-failed distinct writes leave the full budget: 30 healthy writes allowed, the 31st RESOURCE_LIMIT", async () => {
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority();
    double.failRuns = true;
    for (let row = 1; row <= 5; row += 1) {
      const [name, args] = oneCell(row);
      expect(JSON.parse(await authority.executePrimitive(name, args)).ok).toBe(false);
    }
    double.failRuns = false;
    expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
    const healthy: string[] = [];
    for (let row = 6; row <= 36; row += 1) {
      const [name, args] = oneCell(row);
      healthy.push(await authority.executePrimitive(name, args));
    }
    expect(healthy.findIndex((payload) => payload === DENIAL.RESOURCE_LIMIT)).toBe(30);
    expect(healthy.slice(0, 30).every((payload) => JSON.parse(payload).ok === true)).toBe(true);
    expect(healthy[30]).toBe(DENIAL.RESOURCE_LIMIT);
    expect(formulaWrites(double.log)).toHaveLength(30);
    expect(authority.usageSnapshot()).toEqual({ ...ZERO_USAGE, operationCount: 30, cellsWritten: 30, cellsRead: 30, verificationReads: 30 });
  });

  it("F5 serialization and atomicity: concurrent W (fails), W, W -> ok:false, ok:true, DUPLICATE; 31 concurrent distinct writes -> 30 raw, one RESOURCE_LIMIT", async () => {
    let writes = 0;
    await withFakeRaw(async (name, args) => {
      if (name === "read_range") return readOk([[Number(String(args.address).slice(1))]]);
      writes += 1;
      return writes === 1 ? JSON.stringify({ ok: false, error: "host failed", category: "Execution" }) : JSON.stringify({ ok: true, action: name });
    }, async (mod, calls) => {
      const authority = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: true });
      const [name, args] = oneCell(1);
      const results = await Promise.all([
        authority.executePrimitive(name, args),
        authority.executePrimitive(name, args),
        authority.executePrimitive(name, args),
      ]);
      expect(results.map((result) => (result === DENIAL.DUPLICATE ? "DUPLICATE" : JSON.parse(result).ok))).toEqual([false, true, "DUPLICATE"]);
      expect(calls.filter((call) => call === "write_range")).toEqual(["write_range", "write_range"]);
      expect(authority.usageSnapshot()).toEqual({ ...ZERO_USAGE, operationCount: 1, cellsWritten: 1, cellsRead: 1, verificationReads: 1 });
    });
    await withFakeRaw(async (name, args) => (name === "read_range"
      ? readOk([[Number(String(args.address).slice(1))]])
      : JSON.stringify({ ok: true, action: name })), async (mod, calls) => {
      const authority = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: true });
      const results = await Promise.all(Array.from({ length: 31 }, (_v, i) => authority.executePrimitive(...oneCell(i + 1))));
      expect(calls.filter((call) => call === "write_range")).toHaveLength(30);
      expect(results.findIndex((result) => result === DENIAL.RESOURCE_LIMIT)).toBe(30);
      expect(results.filter((result) => result === DENIAL.RESOURCE_LIMIT)).toHaveLength(1);
    });
  });

  it("F6 no refund on a verified-write failure: a mismatched readback keeps the charge and latches READBACK_FAILED", async () => {
    const double = createOfficeDouble();
    double.mismatchFormulas = true;
    double.install();
    const authority = explicitAuthority();
    const [name, args] = oneCell(1);
    expect(JSON.parse(await authority.executePrimitive(name, args)).ok).toBe(true);
    expect(excelRuns(double.log)).toBe(2);
    expect(authority.usageSnapshot()).toEqual({ ...ZERO_USAGE, operationCount: 1, cellsWritten: 1, cellsRead: 1, verificationReads: 1 });
    expect(await authority.executePrimitive("format_range", VALID_ARGS.format_range)).toBe(DENIAL.READBACK_FAILED);
    expect(await authority.executePrimitive(name, args)).toBe(DENIAL.READBACK_FAILED);
    expect(formulaWrites(double.log)).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// §5.2 DP1 — direct route, rollout OFF, against the raw (066fe36) behaviour
// ---------------------------------------------------------------------------
describe("DP1 direct route with rollout OFF is byte-identical to the raw dispatcher", () => {
  it("DP1 27 primitives healthy and failing, repeats, a failed write and its retry, 31 writes, legacy x2, named range and whole column", async () => {
    const names = Object.keys(VALID_ARGS);
    expect(names).toHaveLength(27);
    const write = oneCell(1);
    const steps: Step[] = [
      ...names.map((name): Step => [name, VALID_ARGS[name], false]),
      ...names.map((name): Step => [name, VALID_ARGS[name], true]),
      [write[0], write[1], false], [write[0], write[1], false], [write[0], write[1], false],
      [write[0], write[1], true], [write[0], write[1], false],
      ...Array.from({ length: 31 }, (_v, i): Step => [...oneCell(i + 2), false]),
      ...LEGACY_NAMES.flatMap((name): Step[] => [[name, VALID_ARGS[name], false], [name, VALID_ARGS[name], false]]),
      ["write_range", { sheet: "Sheet1", address: "MyNamedRange", values: [[1]] }, false],
      ["write_range", { sheet: "Sheet1", address: "A:A", values: [[1]] }, false],
    ];
    const reference = await rawOracle(steps);
    const double = createOfficeDouble();
    double.install();
    const authority = createPrimitiveMutationAuthority({ trustedRequest: REQUEST("ambiguous_mutation"), workbookPlanRolloutEnabled: false });
    const payloads: string[] = [];
    for (const [name, args, fail] of steps) {
      double.failRuns = fail;
      payloads.push(await authority.executePrimitive(name, args));
    }
    double.failRuns = false;
    expect(payloads).toEqual(reference.payloads);
    expect(double.log).toEqual(reference.log);
    expect(excelRuns(double.log)).toBe(excelRuns(reference.log));
    expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
  });
});
