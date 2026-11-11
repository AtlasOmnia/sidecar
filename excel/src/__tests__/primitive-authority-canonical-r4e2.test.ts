/**
 * R4e-2 authority tests, file 4 (dev3-r4e2-release §5: C1-C8).
 *
 * Rollout ON: the duplicate signature canonicalizes address keys (address,
 * sourceAddress, targetAddress, dataRange) to an equivalent A1 spelling on the
 * same sheet (R4c lane-B A1). Values, sheet-name case, named ranges, whole
 * rows/columns and R1C1 stay literal (C3). Narrow-only: every literal repeat is
 * still DUPLICATE (C4). Rollout OFF is unchanged (C7).
 *
 * The Office double and the withFakeRaw pattern are copied from
 * primitive-authority-refund-r4d.test.ts (not imported), with the R4e-2 sync
 * fault switch. C4 uses a seeded PRNG (seed recorded below).
 */
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

type Call = readonly [string, Record<string, unknown>];
type Step = readonly [string, Record<string, unknown>, Fault];

const write = (address: unknown, values: unknown = [[1]], sheet: unknown = "Sheet1"): Call =>
  ["write_range", { sheet, address, values }];
const formatAt = (address: string): Call => ["format_range", { sheet: "Sheet1", address, format: { bold: true } }];
const autofill = (sourceAddress: string, targetAddress: string): Call => ["autofill_formula", { sheet: "Sheet1", sourceAddress, targetAddress }];
const chart = (dataRange: string): Call => ["create_chart", { sheet: "Sheet1", dataRange, chartType: "ColumnClustered", title: "T" }];
const comment = (address: string): Call => ["add_comment", { sheet: "Sheet1", address, text: "note" }];

const C1_SPELLINGS = ["a1", "$A$1", " A1 ", "Sheet1!A1", "'Sheet1'!$a$1", "A1:A1", "$A$1:a1"];
const C2_GROUPS: readonly (readonly [Call, readonly Call[]])[] = [
  [formatAt("A1:B2"), [formatAt("B2:A1"), formatAt("b1:a2"), formatAt("$A$2:$B$1"), formatAt("Sheet1!a1:b2")]],
  [autofill("C1", "C2:C5"), [autofill("$c$1", "c5:C2")]],
  [chart("A1:B5"), [chart("b5:a1")]],
  [comment("D1"), [comment("d1")]],
];

const FIXED_NOW = 1_800_000_000_000;
beforeEach(() => {
  vi.spyOn(Date, "now").mockReturnValue(FIXED_NOW);
});
afterEach(() => {
  vi.restoreAllMocks();
  delete (globalThis as Record<string, unknown>).Excel;
  vi.doUnmock("../excelTools");
});

async function rawOracle(steps: readonly Step[]): Promise<{ payloads: string[]; log: string[] }> {
  const double = createOfficeDouble();
  double.install();
  const payloads: string[] = [];
  for (const [name, args, fault] of steps) {
    double.fault = fault;
    payloads.push(await getExcelToolDefinition(name)!.execute(args));
  }
  double.fault = "ok";
  return { payloads, log: [...double.log] };
}

const explicitAuthority = (rollout = true) =>
  createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: rollout });

function gateFor(authority: PrimitiveMutationAuthority): (name: string, args: Record<string, unknown>) => Promise<string> {
  return createServerPrimitiveGate({
    getBridgeSessionId: () => "session-r4e2",
    getActiveTurn: () => ({ turnOperationId: "turn-r4e2", sessionId: "session-r4e2" }),
    resolveAuthority: () => authority,
    executeWithAuthority: (name, args, resolved) => resolved.executePrimitive(name, args),
    executeReadOnly: (name, args) => getExcelToolDefinition(name)!.execute(args),
  });
}

// ---------------------------------------------------------------------------
// withFakeRaw (copied R4c/R4d pattern): ok:true writes whose readback echoes
// the last written grid, so no call ever latches.
// ---------------------------------------------------------------------------
interface AuthorityModule {
  createPrimitiveMutationAuthority: typeof createPrimitiveMutationAuthority;
}
const AUTHORITY_MODULE_PATH = "../primitiveMutationAuthority";

async function withEchoRaw(run: (mod: AuthorityModule, calls: string[]) => Promise<void>): Promise<void> {
  vi.resetModules();
  const calls: string[] = [];
  let lastValues: unknown = [[""]];
  vi.doMock("../excelTools", async () => {
    const actual = await vi.importActual<Record<string, unknown>>("../excelTools");
    return {
      ...actual,
      claimRawPrimitiveExecutor: () => async (name: string, args: Record<string, unknown>) => {
        calls.push(name);
        if (name === "read_range") {
          const formulas = Array.isArray(lastValues) ? lastValues as unknown[][] : [[""]];
          return JSON.stringify({ sheet: "Sheet1", address: "Sheet1!A1", rowCount: formulas.length, columnCount: formulas[0]?.length ?? 0, values: formulas, formulas });
        }
        if (name === "write_range") lastValues = args.values;
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

// ---------------------------------------------------------------------------
// Independent oracles for C4. Literal: key-sorted JSON of the schema keys.
// Canonical upper bound: an independent implementation of the release's A1
// rule (trim, matching sheet prefix, $, case, single-cell collapse, corners).
// ---------------------------------------------------------------------------
function literalKey(args: Record<string, unknown>): string {
  return JSON.stringify(["address", "sheet", "values"].map((key) => [key, args[key]]));
}

function oracleColumn(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function oracleCanonical(sheet: unknown, value: unknown): unknown {
  if (typeof value !== "string") return value;
  let text = value.trim();
  const bang = text.lastIndexOf("!");
  if (bang >= 0) {
    let prefix = text.slice(0, bang);
    if (/^'(?:[^']|'')+'$/.test(prefix)) prefix = prefix.slice(1, -1).replace(/''/g, "'");
    if (typeof sheet !== "string" || prefix !== sheet) return value;
    text = text.slice(bang + 1);
  }
  const parts = text.split(":");
  if (parts.length < 1 || parts.length > 2) return value;
  const corners: [number, number][] = [];
  for (const part of parts) {
    const match = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$/.exec(part);
    if (match === null) return value;
    const column = oracleColumn(match[1]);
    const row = Number(match[2]);
    if (column < 1 || column > 16384 || row < 1 || row > 1048576) return value;
    corners.push([column, row]);
  }
  const columns = corners.map((corner) => corner[0]);
  const rows = corners.map((corner) => corner[1]);
  const label = (n: number): string => {
    let out = "";
    while (n > 0) { const rem = (n - 1) % 26; out = String.fromCharCode(65 + rem) + out; n = Math.floor((n - 1) / 26); }
    return out;
  };
  const first = `${label(Math.min(...columns))}${Math.min(...rows)}`;
  const last = `${label(Math.max(...columns))}${Math.max(...rows)}`;
  return first === last ? first : `${first}:${last}`;
}

function canonicalKey(args: Record<string, unknown>): string {
  return JSON.stringify([["address", oracleCanonical(args.sheet, args.address)], ["sheet", args.sheet], ["values", args.values]]);
}

/** mulberry32, fixed seed 0x5EED4E2 (recorded). */
function prng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6D2B79F5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const C4_SEED = 0x5EED4E2;

// ---------------------------------------------------------------------------
// §5.1 RED rows (rollout ON)
// ---------------------------------------------------------------------------
describe("R4e-2 C rows: equivalent A1 spellings on the same sheet are DUPLICATE", () => {
  it("C1 after write_range A1, each equivalent spelling (case, $, spaces, matching sheet prefix, quoted prefix, A1:A1) is DUPLICATE with no batch and no charge", async () => {
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority();
    expect(JSON.parse(await authority.executePrimitive(...write("A1"))).ok).toBe(true);
    const usage = authority.usageSnapshot();
    const runs = excelRuns(double.log);
    const results: Record<string, string> = {};
    for (const spelling of C1_SPELLINGS) results[spelling] = await authority.executePrimitive(...write(spelling));
    expect(results).toEqual(Object.fromEntries(C1_SPELLINGS.map((spelling) => [spelling, DENIAL.DUPLICATE])));
    expect(excelRuns(double.log) - runs).toBe(0);
    expect(authority.usageSnapshot()).toEqual(usage);
  });

  it("C2 corner order and case for format_range, autofill_formula (source and target), create_chart dataRange and add_comment are DUPLICATE", async () => {
    const double = createOfficeDouble();
    double.install();
    // explicit_structure: create_chart is a structure effect.
    const authority = createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_structure"), workbookPlanRolloutEnabled: true });
    const observed: string[] = [];
    const expected: string[] = [];
    for (const [first, variants] of C2_GROUPS) {
      const firstPayload = await authority.executePrimitive(first[0], first[1]);
      observed.push(`${first[0]} first -> ${JSON.parse(firstPayload).ok === true ? "ok" : firstPayload}`);
      expected.push(`${first[0]} first -> ok`);
      const runs = excelRuns(double.log);
      for (const variant of variants) {
        observed.push(`${variant[0]} ${JSON.stringify(variant[1])} -> ${await authority.executePrimitive(variant[0], variant[1])}`);
        expected.push(`${variant[0]} ${JSON.stringify(variant[1])} -> ${DENIAL.DUPLICATE}`);
      }
      observed.push(`${first[0]} variant batches ${excelRuns(double.log) - runs}`);
      expected.push(`${first[0]} variant batches 0`);
    }
    expect(observed).toEqual(expected);
  });

  it("C8 server route: through createServerPrimitiveGate (live turn, rollout on) A1 then $A$1 is DUPLICATE with no batch", async () => {
    const double = createOfficeDouble();
    double.install();
    const gate = gateFor(explicitAuthority());
    expect(JSON.parse(await gate(...write("A1"))).ok).toBe(true);
    const runs = excelRuns(double.log);
    expect(await gate(...write("$A$1"))).toBe(DENIAL.DUPLICATE);
    expect(excelRuns(double.log) - runs).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// §5.2 GREEN pins (identical at base and HEAD)
// ---------------------------------------------------------------------------
describe("R4e-2 C pins: documented scope, narrow-only property, robustness, refund interplay, rollout off", () => {
  it("C3 [pin] out-of-scope spellings are not DUPLICATE: the second call answers exactly as it would on a fresh authority", async () => {
    const pairs: readonly (readonly [Call, Call, "raw" | "resource"])[] = [
      [write("A1"), write("Sheet2!A1"), "raw"],
      [write("Sheet2!A1"), write("SHEET2!A1"), "raw"],
      [write("A1"), write("A1:A2", [[1], [1]]), "raw"],
      [write("A1", [["1"]]), write("A1", [[1]]), "raw"],
      [write("A1", [[1]], "sheet1"), write("A1", [[1]], "Sheet1"), "raw"],
      [write("MyRange"), write("MYRANGE"), "resource"],
      [write("A:A"), write("$A:$A"), "resource"],
      [write("R1C1"), write("A1"), "raw"],
      [write("XFE1"), write("xfe1"), "resource"],
      [write("A0"), write("a0"), "resource"],
    ];
    await withEchoRaw(async (mod, calls) => {
      const observed: string[] = [];
      const expected: string[] = [];
      for (const [first, second, kind] of pairs) {
        const fresh = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: true });
        const alone = await fresh.executePrimitive(second[0], second[1]);
        const authority = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: true });
        await authority.executePrimitive(first[0], first[1]);
        const before = calls.length;
        const after = await authority.executePrimitive(second[0], second[1]);
        const label = `${JSON.stringify(first[1])} then ${JSON.stringify(second[1])}`;
        observed.push(`${label}: ${after === DENIAL.DUPLICATE ? "DUPLICATE" : after === alone ? "as-fresh" : after} raw=${calls.length > before}`);
        expected.push(`${label}: as-fresh raw=${kind === "raw"}`);
      }
      expect(observed).toEqual(expected);
    });
  });

  it("C4 [pin] narrow-only (seed 0x5EED4E2, 500 sequences x <= 20 calls): every literal repeat is DUPLICATE, and every DUPLICATE is a canonical repeat", async () => {
    const corpus: readonly (readonly [string, number])[] = [
      ["A1", 1], ["a1", 1], ["$A$1", 1], [" A1 ", 1], ["Sheet1!A1", 1], ["'Sheet1'!$a$1", 1], ["A1:A1", 1], ["$A$1:a1", 1],
      ["B2", 1], ["b2", 1], ["$B2", 1], ["Sheet1!b$2", 1], ["Sheet2!A1", 1], ["'Sheet2'!A1", 1], ["SHEET1!A1", 1], ["MyRange", 1],
      ["A1:B2", 2], ["B2:A1", 2], ["b1:a2", 2], ["$A$2:$B$1", 2], ["Sheet1!a1:b2", 2], ["A1 :B2", 2],
    ];
    const grids: Record<number, readonly unknown[][][]> = { 1: [[[1]], [["x"]]], 2: [[[1, 2], [3, 4]], [[5, 6], [7, 8]]] };
    const random = prng(C4_SEED);
    let literalRepeats = 0;
    let canonicalOnly = 0;
    await withEchoRaw(async (mod) => {
      const violations: string[] = [];
      for (let sequence = 0; sequence < 500; sequence += 1) {
        const authority = mod.createPrimitiveMutationAuthority({ trustedRequest: REQUEST("explicit_mutation"), workbookPlanRolloutEnabled: true });
        const literal = new Set<string>();
        const canonical = new Set<string>();
        const length = 1 + Math.floor(random() * 20);
        for (let step = 0; step < length; step += 1) {
          const [address, shape] = corpus[Math.floor(random() * corpus.length)];
          const values = grids[shape][Math.floor(random() * 2)];
          const args = { sheet: "Sheet1", address, values };
          const result = await authority.executePrimitive("write_range", args);
          const isLiteralRepeat = literal.has(literalKey(args));
          const isCanonicalRepeat = canonical.has(canonicalKey(args));
          if (isLiteralRepeat) literalRepeats += 1;
          if (isCanonicalRepeat && !isLiteralRepeat) canonicalOnly += 1;
          if (isLiteralRepeat && result !== DENIAL.DUPLICATE) violations.push(`seq ${sequence} step ${step}: literal repeat ${address} -> ${result}`);
          if (result === DENIAL.DUPLICATE && !isCanonicalRepeat) violations.push(`seq ${sequence} step ${step}: DUPLICATE outside canonical oracle ${address}`);
          if (JSON.parse(result).ok === true) {
            literal.add(literalKey(args));
            canonical.add(canonicalKey(args));
          }
        }
      }
      expect(violations).toEqual([]);
    });
    expect(literalRepeats).toBeGreaterThan(100);
    expect(canonicalOnly).toBeGreaterThan(100);
  });

  it("C5 [pin] robustness: malformed address values never throw; an identical repeat is DUPLICATE only if the first call was ok:true", async () => {
    const addresses: unknown[] = [123, null, {}, [], "A".repeat(100_000), "!", "''!A1", "'Sheet1'!", "Sheet1!!A1", ":", "A1:B2:C3", "Ａ１"];
    const observed: string[] = [];
    const expected: string[] = [];
    for (const address of addresses) {
      const double = createOfficeDouble();
      double.install();
      const authority = explicitAuthority();
      const label = typeof address === "string" ? JSON.stringify(address.slice(0, 16)) : JSON.stringify(address);
      const first = await authority.executePrimitive(...write(address)).then((value) => value, () => "THREW");
      const second = await authority.executePrimitive(...write(address)).then((value) => value, () => "THREW");
      const firstOk = first !== "THREW" && JSON.parse(first).ok === true;
      observed.push(`${label}: first=${first === "THREW" ? "THREW" : "string"} second=${second === "THREW" ? "THREW" : second === DENIAL.DUPLICATE ? "DUPLICATE" : "other"}`);
      expected.push(`${label}: first=string second=${firstOk ? "DUPLICATE" : "other"}`);
    }
    expect(observed).toEqual(expected);
  });

  it("C6 [pin] refund interplay: A1 refunded after Excel.run rejects before its batch; then $A$1 executes", async () => {
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority();
    double.fault = "run";
    expect(JSON.parse(await authority.executePrimitive(...write("A1"))).ok).toBe(false);
    double.fault = "ok";
    expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
    expect(JSON.parse(await authority.executePrimitive(...write("$A$1"))).ok).toBe(true);
    expect(excelRuns(double.log)).toBe(3);
  });

  it("C7 [pin] rollout OFF: C1/C2 spellings all execute; payloads and the full double log equal the raw oracle", async () => {
    const calls: Call[] = [write("A1"), ...C1_SPELLINGS.map((spelling) => write(spelling)), ...C2_GROUPS.flatMap(([first, variants]) => [first, ...variants])];
    const steps = calls.map(([name, args]): Step => [name, args, "ok"]);
    const reference = await rawOracle(steps);
    const double = createOfficeDouble();
    double.install();
    const authority = explicitAuthority(false);
    const payloads: string[] = [];
    for (const [name, args] of calls) payloads.push(await authority.executePrimitive(name, args));
    expect(payloads).toEqual(reference.payloads);
    expect(double.log).toEqual(reference.log);
    expect(payloads).not.toContain(DENIAL.DUPLICATE);
    expect(authority.usageSnapshot()).toEqual(ZERO_USAGE);
  });
});
