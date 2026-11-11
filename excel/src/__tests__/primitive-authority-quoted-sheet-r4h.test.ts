/**
 * R4h (r4e2 tester-report-B A2): the primitive authority's canonical address
 * unescapes a doubled quote inside a quoted sheet prefix ('It''s'!A1 is the
 * sheet It's), and folds nothing else (dev3-r4h-release section 3 T9, 5.5).
 *
 * Rollout ON, explicit_mutation. The Office double, excelRuns, the DENIAL and
 * ZERO_USAGE constants, REQUEST, the Date.now pin and explicitAuthority() are
 * COPIED as text from primitive-authority-canonical-r4e2.test.ts (lines
 * 20-242, 245-252, 254-257, 259-260, 280-288 and 302-303 at ac981da); that file
 * is neither imported nor edited.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { getExcelToolDefinition } from "../excelTools";
import { createPrimitiveMutationAuthority } from "../primitiveMutationAuthority";
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

// ---------------------------------------------------------------------------
// Fixed expectations (independent constants)
// ---------------------------------------------------------------------------
const deny = (error: string, category: string): string => JSON.stringify({ ok: false, error, category });
const DENIAL = {
  DUPLICATE: deny("Workbook edit denied: the identical edit was already applied in this turn.", "repetition"),
  READBACK_FAILED: deny("Workbook edit denied: an earlier edit in this turn did not verify.", "validation"),
} as const;

const ZERO_USAGE = {
  operationCount: 0, cellsRead: 0, cellsWritten: 0, formulaUtf8Bytes: 0, csvPayloadUtf8Bytes: 0,
  commentUtf8Bytes: 0, sheetsCreated: 0, chartsCreated: 0, tablesCreated: 0, verificationReads: 0,
};

const REQUEST = (classification: WorkbookRequestClassification, selectionScopedRequest = false): TrustedWorkbookRequest =>
  Object.freeze({ classification, selectionScopedRequest });

const FIXED_NOW = 1_800_000_000_000;
beforeEach(() => {
  vi.spyOn(Date, "now").mockReturnValue(FIXED_NOW);
});
afterEach(() => {
  vi.restoreAllMocks();
  delete (globalThis as Record<string, unknown>).Excel;
  vi.doUnmock("../excelTools");
});

const explicitAuthority = (rollout = true) =>
  createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: rollout });

// ---------------------------------------------------------------------------
// R4h A2 rows
// ---------------------------------------------------------------------------
const QUOTED_SHEET = "It's";

/** write_range of a 1x1 value on an arbitrary sheet (local helper). */
const writeOn = (sheet: string, address: string): readonly [string, Record<string, unknown>] =>
  ["write_range", { sheet, address, values: [[1]] }];

describe("R4h A2: quoted sheet prefixes with an escaped quote (primitive authority canonical address)", () => {
  it("QS1 quoted sheet prefix with an escaped quote ('It''s'!A1) is the same cell as A1 on sheet It's: DUPLICATE, no batch, no charge", async () => {
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority();
    const first = await authority.executePrimitive(...writeOn(QUOTED_SHEET, "A1"));
    expect(JSON.parse(first).ok, first).toBe(true);
    const usage = authority.usageSnapshot();
    expect(usage).not.toEqual(ZERO_USAGE);
    const runs = excelRuns(double.log);
    const spellings = ["'It''s'!A1", "'It''s'!$a$1", " 'It''s'!A1:A1 "];
    const results: Record<string, string> = {};
    for (const spelling of spellings) results[spelling] = await authority.executePrimitive(...writeOn(QUOTED_SHEET, spelling));
    expect(results).toEqual(Object.fromEntries(spellings.map((spelling) => [spelling, DENIAL.DUPLICATE])));
    expect(excelRuns(double.log) - runs).toBe(0);
    expect(authority.usageSnapshot()).toEqual(usage);
  });

  it("QS2 [pin] doubled, malformed or unquoted quote spellings are not folded: each answers as on a fresh authority", async () => {
    // 'It''''s'!A1 names the sheet It''s; 'It's'!A1 is malformed quoting; It''s!A1 is unquoted.
    const spellings = ["'It''''s'!A1", "'It's'!A1", "It''s!A1"];
    const observed: string[] = [];
    const expected: string[] = [];
    for (const spelling of spellings) {
      const freshDouble = createOfficeDouble();
      freshDouble.install();
      const fresh = await explicitAuthority().executePrimitive(...writeOn(QUOTED_SHEET, spelling));
      const double = createOfficeDouble();
      double.install();
      const authority = explicitAuthority();
      const first = await authority.executePrimitive(...writeOn(QUOTED_SHEET, "A1"));
      expect(JSON.parse(first).ok, first).toBe(true);
      const after = await authority.executePrimitive(...writeOn(QUOTED_SHEET, spelling));
      const verdict = after === DENIAL.DUPLICATE ? "DUPLICATE" : after === fresh ? "as-fresh" : `differs: after=${after} fresh=${fresh}`;
      observed.push(`${spelling}: ${verdict}`);
      expected.push(`${spelling}: as-fresh`);
    }
    expect(observed).toEqual(expected);
  });

  // R4j J1 (R4h lane-B B-A3 / mutant B5): two escaped-quote pairs. Production already unescapes every
  // pair (global replace); a first-pair-only unescape survives QS1 and is killed here (jm1).
  it("QS3 [R4j] two escaped-quote pairs ('It''s Bob''s'!A1) are the same cell as A1 on sheet It's Bob's: DUPLICATE, no batch, no charge", async () => {
    const sheet = "It's Bob's";
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority();
    const first = await authority.executePrimitive(...writeOn(sheet, "A1"));
    expect(JSON.parse(first).ok, first).toBe(true);
    const usage = authority.usageSnapshot();
    expect(usage).not.toEqual(ZERO_USAGE);
    const runs = excelRuns(double.log);
    const spellings = ["'It''s Bob''s'!A1", "'It''s Bob''s'!$a$1", " 'It''s Bob''s'!A1:A1 "];
    const results: Record<string, string> = {};
    for (const spelling of spellings) results[spelling] = await authority.executePrimitive(...writeOn(sheet, spelling));
    expect(results).toEqual(Object.fromEntries(spellings.map((spelling) => [spelling, DENIAL.DUPLICATE])));
    expect(excelRuns(double.log) - runs).toBe(0);
    expect(authority.usageSnapshot()).toEqual(usage);
  });

  it("QS3b [R4j pin] two-pair malformed, over-doubled or unquoted spellings are not folded: each answers as on a fresh authority", async () => {
    // 'It''s Bob's'!A1 is malformed quoting; 'It''''s Bob''s'!A1 names the sheet It''s Bob's; It''s Bob''s!A1 is unquoted.
    const sheet = "It's Bob's";
    const spellings = ["'It''s Bob's'!A1", "'It''''s Bob''s'!A1", "It''s Bob''s!A1"];
    const observed: string[] = [];
    const expected: string[] = [];
    for (const spelling of spellings) {
      const freshDouble = createOfficeDouble();
      freshDouble.install();
      const fresh = await explicitAuthority().executePrimitive(...writeOn(sheet, spelling));
      const double = createOfficeDouble();
      double.install();
      const authority = explicitAuthority();
      const first = await authority.executePrimitive(...writeOn(sheet, "A1"));
      expect(JSON.parse(first).ok, first).toBe(true);
      const after = await authority.executePrimitive(...writeOn(sheet, spelling));
      const verdict = after === DENIAL.DUPLICATE ? "DUPLICATE" : after === fresh ? "as-fresh" : `differs: after=${after} fresh=${fresh}`;
      observed.push(`${spelling}: ${verdict}`);
      expected.push(`${spelling}: as-fresh`);
    }
    expect(observed).toEqual(expected);
  });
});
