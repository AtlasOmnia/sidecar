import type {
  InspectionOperation,
  VerificationRule,
  WorkbookOperation,
  WorkbookPlan,
} from "./generated/workbookPlanSchema";
import { workbookPlanContract } from "./workbookPlanContract";
import type { VerifiedWorkbookCompileContext } from "./workbookPlanCompiler";
import type { WorkbookPlanResourceLimits, WorkbookPlanResourceSummary } from "./workbookPlanPolicy";

const MAX_EXCEL_ROW = 1_048_576;
const MAX_EXCEL_COLUMN = 16_384;
const RANGE_RE = /^\$?([A-Z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Z]{1,3})\$?([1-9][0-9]*))?$/i;

export const WORKBOOK_PLAN_RESOURCE_LIMITS: WorkbookPlanResourceLimits = Object.freeze({
  ...workbookPlanContract.resourceLimits,
});

export interface WorkbookPlanResourceContext {
  readonly compileContext: VerifiedWorkbookCompileContext;
}

export interface WorkbookPlanResourceAccountingSuccess {
  readonly ok: true;
  readonly usage: WorkbookPlanResourceSummary;
  readonly limits: WorkbookPlanResourceLimits;
}

export interface WorkbookPlanResourceAccountingFailure {
  readonly ok: false;
  readonly reason:
    | "unknown_or_unbounded_range"
    | "unknown_table_bounds"
    | "range_shape_mismatch"
    | "resource_limit_exceeded";
}

export type WorkbookPlanResourceAccounting =
  | WorkbookPlanResourceAccountingSuccess
  | WorkbookPlanResourceAccountingFailure;

type MutableResourceSummary = {
  -readonly [Key in keyof WorkbookPlanResourceSummary]: number;
};

interface ParsedRange {
  readonly rows: number;
  readonly columns: number;
  readonly cells: number;
}

const emptyUsage = (): MutableResourceSummary => ({
  operationCount: 0,
  cellsRead: 0,
  cellsWritten: 0,
  formulaUtf8Bytes: 0,
  csvPayloadUtf8Bytes: 0,
  commentUtf8Bytes: 0,
  sheetsCreated: 0,
  chartsCreated: 0,
  tablesCreated: 0,
  verificationReads: 0,
});

function utf8Length(value: string): number {
  return new TextEncoder().encode(value).length;
}

function columnNumber(value: string): number | undefined {
  let result = 0;
  for (const character of value.toUpperCase()) {
    const code = character.charCodeAt(0) - 64;
    if (code < 1 || code > 26) return undefined;
    result = result * 26 + code;
    if (result > MAX_EXCEL_COLUMN) return undefined;
  }
  return result;
}

function parseBoundedA1(address: string): ParsedRange | undefined {
  if (typeof address !== "string") return undefined;
  const match = RANGE_RE.exec(address.trim());
  if (match === null) return undefined;
  const firstColumn = columnNumber(match[1]);
  const lastColumn = columnNumber(match[3] ?? match[1]);
  const firstRow = Number(match[2]);
  const lastRow = Number(match[4] ?? match[2]);
  if (
    firstColumn === undefined ||
    lastColumn === undefined ||
    !Number.isSafeInteger(firstRow) ||
    !Number.isSafeInteger(lastRow) ||
    firstRow > MAX_EXCEL_ROW ||
    lastRow > MAX_EXCEL_ROW
  ) {
    return undefined;
  }
  const rows = Math.abs(lastRow - firstRow) + 1;
  const columns = Math.abs(lastColumn - firstColumn) + 1;
  const cells = rows * columns;
  if (!Number.isSafeInteger(cells)) return undefined;
  return { rows, columns, cells };
}

function add(usage: MutableResourceSummary, key: keyof WorkbookPlanResourceSummary, value: number): void {
  usage[key] += value;
}

function tableCells(
  name: string,
  context: WorkbookPlanResourceContext,
): number | undefined {
  const table = context.compileContext.tablesByName.get(name);
  if (table === undefined || table.rowCount === undefined || table.columnCount <= 0) return undefined;
  if (!Number.isSafeInteger(table.rowCount) || table.rowCount <= 0) return undefined;
  const cells = table.rowCount * table.columnCount;
  return Number.isSafeInteger(cells) ? cells : undefined;
}

function rangeForOperation(operation: WorkbookOperation): ParsedRange | undefined {
  switch (operation.kind) {
    case "write_values":
    case "write_formula":
    case "format_range":
    case "create_table":
    case "set_validation_list":
    case "conditional_format":
    case "add_comment":
      return parseBoundedA1(operation.address);
    case "fill_formula":
      return parseBoundedA1(operation.targetAddress);
    case "create_chart":
      return parseBoundedA1(operation.dataRange);
    default:
      return undefined;
  }
}

function addVerificationCost(
  rule: VerificationRule,
  usage: MutableResourceSummary,
  context: WorkbookPlanResourceContext,
): WorkbookPlanResourceAccountingFailure | undefined {
  add(usage, "verificationReads", 1);
  switch (rule.kind) {
    case "range_values_equal": {
      const range = parseBoundedA1(rule.address);
      if (range === undefined) return { ok: false, reason: "unknown_or_unbounded_range" };
      const width = rule.expectedValues[0]?.length;
      if (width === undefined || rule.expectedValues.length !== range.rows || width !== range.columns) {
        return { ok: false, reason: "range_shape_mismatch" };
      }
      add(usage, "cellsRead", range.cells);
      return undefined;
    }
    case "range_formulas_equal": {
      const range = parseBoundedA1(rule.address);
      if (range === undefined) return { ok: false, reason: "unknown_or_unbounded_range" };
      const width = rule.expectedFormulas[0]?.length;
      if (width === undefined || rule.expectedFormulas.length !== range.rows || width !== range.columns) {
        return { ok: false, reason: "range_shape_mismatch" };
      }
      add(usage, "cellsRead", range.cells);
      return undefined;
    }
    case "range_formula_pattern":
    case "range_number_format_equal":
    case "sentinel_unchanged":
    case "comment_exists": {
      const range = parseBoundedA1(rule.address);
      if (range === undefined) return { ok: false, reason: "unknown_or_unbounded_range" };
      add(usage, "cellsRead", range.cells);
      return undefined;
    }
    case "table_row_order":
    case "table_filter_state": {
      const cells = tableCells(rule.tableName, context);
      if (cells === undefined) return { ok: false, reason: "unknown_table_bounds" };
      add(usage, "cellsRead", cells);
      return undefined;
    }
    case "table_exists":
    case "sheet_exists":
    case "chart_exists":
    case "calculation_completed":
      return undefined;
  }
}

function accountOperation(
  operation: WorkbookOperation,
  usage: MutableResourceSummary,
  context: WorkbookPlanResourceContext,
): WorkbookPlanResourceAccountingFailure | undefined {
  const range = rangeForOperation(operation);
  switch (operation.kind) {
    case "write_values": {
      if (range === undefined) return { ok: false, reason: "unknown_or_unbounded_range" };
      const width = operation.values[0]?.length;
      if (width === undefined || operation.values.length !== range.rows || width !== range.columns) {
        return { ok: false, reason: "range_shape_mismatch" };
      }
      add(usage, "cellsWritten", range.cells);
      for (const row of operation.values) {
        for (const value of row) {
          if (typeof value === "string" && value.startsWith("=")) add(usage, "formulaUtf8Bytes", utf8Length(value));
        }
      }
      return undefined;
    }
    case "write_formula":
      if (range === undefined) return { ok: false, reason: "unknown_or_unbounded_range" };
      add(usage, "cellsWritten", range.cells);
      add(usage, "formulaUtf8Bytes", utf8Length(operation.formula) * range.cells);
      return undefined;
    case "fill_formula": {
      const source = parseBoundedA1(operation.sourceAddress);
      const target = range;
      if (source === undefined || target === undefined) return { ok: false, reason: "unknown_or_unbounded_range" };
      add(usage, "cellsRead", source.cells);
      add(usage, "cellsWritten", target.cells);
      return undefined;
    }
    case "format_range":
      if (range === undefined) return { ok: false, reason: "unknown_or_unbounded_range" };
      add(usage, "cellsWritten", range.cells);
      return undefined;
    case "create_table":
      if (range === undefined) return { ok: false, reason: "unknown_or_unbounded_range" };
      add(usage, "cellsRead", range.cells);
      add(usage, "tablesCreated", 1);
      return undefined;
    case "sort_table":
    case "filter_table":
    case "reset_table_filter": {
      const cells = tableCells(operation.tableName, context);
      if (cells === undefined) return { ok: false, reason: "unknown_table_bounds" };
      add(usage, "cellsRead", cells);
      add(usage, "cellsWritten", cells);
      return undefined;
    }
    case "create_chart":
      if (range === undefined) return { ok: false, reason: "unknown_or_unbounded_range" };
      add(usage, "cellsRead", range.cells);
      add(usage, "chartsCreated", 1);
      return undefined;
    case "set_validation_list":
      if (range === undefined) return { ok: false, reason: "unknown_or_unbounded_range" };
      add(usage, "cellsWritten", range.cells);
      add(usage, "csvPayloadUtf8Bytes", utf8Length(operation.source));
      return undefined;
    case "conditional_format":
      if (range === undefined) return { ok: false, reason: "unknown_or_unbounded_range" };
      add(usage, "cellsWritten", range.cells);
      return undefined;
    case "freeze_panes":
      return undefined;
    case "create_sheet":
    case "copy_sheet":
      add(usage, "sheetsCreated", 1);
      return undefined;
    case "rename_sheet":
      return undefined;
    case "add_comment":
      if (range === undefined) return { ok: false, reason: "unknown_or_unbounded_range" };
      add(usage, "cellsWritten", range.cells);
      add(usage, "commentUtf8Bytes", utf8Length(operation.text));
      return undefined;
    case "calculate":
      return undefined;
  }
}

function exceedsLimits(usage: WorkbookPlanResourceSummary): boolean {
  for (const key of Object.keys(WORKBOOK_PLAN_RESOURCE_LIMITS) as (keyof WorkbookPlanResourceSummary)[]) {
    if (usage[key] > WORKBOOK_PLAN_RESOURCE_LIMITS[key]) return true;
  }
  return false;
}

/**
 * Count every bounded host range, table target, payload, and verification read
 * before route code is allowed to obtain mutation access.
 */
export function accountWorkbookPlanResources(
  plan: WorkbookPlan,
  context: WorkbookPlanResourceContext,
): WorkbookPlanResourceAccounting {
  const usage = emptyUsage();
  // An initial empty context may defer table dimensions, but never skip
  // independently known operation/verification costs before host capture.
  let unknownTableBounds = false;
  if (plan.outcome === "inspect") {
    usage.operationCount = 0;
    for (const read of plan.reads as readonly InspectionOperation[]) {
      const range = parseBoundedA1(read.address);
      if (range === undefined) return { ok: false, reason: "unknown_or_unbounded_range" };
      add(usage, "cellsRead", range.cells);
      add(usage, "verificationReads", 1);
    }
  } else if (plan.outcome === "execute") {
    usage.operationCount = plan.operations.length;
    for (const operation of plan.operations) {
      const failure = accountOperation(operation, usage, context);
      if (failure?.reason === "unknown_table_bounds") unknownTableBounds = true;
      else if (failure !== undefined) return failure;
    }
    for (const rule of plan.verification) {
      const failure = addVerificationCost(rule, usage, context);
      if (failure?.reason === "unknown_table_bounds") unknownTableBounds = true;
      else if (failure !== undefined) return failure;
    }
  }
  if (exceedsLimits(usage)) return { ok: false, reason: "resource_limit_exceeded" };
  if (unknownTableBounds) return { ok: false, reason: "unknown_table_bounds" };
  return { ok: true, usage, limits: WORKBOOK_PLAN_RESOURCE_LIMITS };
}

const EMPTY_PRIMITIVE_RESOURCE_CONTEXT: WorkbookPlanResourceContext = Object.freeze({
  compileContext: Object.freeze({
    sheetNames: new Set<string>(),
    tablesByName: new Map(),
    supportedPrimitiveNames: new Set<string>(),
    supportedCapabilities: new Set<string>(),
    chartSourceResolutionsBySheet: new Map(),
  }),
});

/**
 * R4c: per-turn accounting of one direct-route primitive (rollout on only).
 * Copies `prior`, adds one operation, charges the operation with the plan
 * accounting rules (sort/filter/reset count the operation only, OQ-C4), adds
 * the post-write readback cost, and applies the plan resource limits.
 */
export function accountPrimitiveResourceIncrement(
  prior: WorkbookPlanResourceSummary,
  operation: WorkbookOperation,
  options: {
    readonly readbackCells: number;
    readonly readbackReads: number;
    readonly unknownTableBounds: "operation_only";
  },
): WorkbookPlanResourceAccounting {
  const usage: MutableResourceSummary = { ...emptyUsage(), ...prior };
  add(usage, "operationCount", 1);
  if (
    operation.kind !== "sort_table" &&
    operation.kind !== "filter_table" &&
    operation.kind !== "reset_table_filter"
  ) {
    const failure = accountOperation(operation, usage, EMPTY_PRIMITIVE_RESOURCE_CONTEXT);
    if (failure !== undefined) return failure;
  }
  add(usage, "verificationReads", options.readbackReads);
  add(usage, "cellsRead", options.readbackCells);
  if (exceedsLimits(usage)) return { ok: false, reason: "resource_limit_exceeded" };
  return { ok: true, usage, limits: WORKBOOK_PLAN_RESOURCE_LIMITS };
}
