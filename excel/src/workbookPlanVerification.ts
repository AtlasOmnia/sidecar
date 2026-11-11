/**
 * Deterministic, content-free postcondition verification for compiled V1 plans.
 *
 * The authored boundary is deliberately a bounded JSON string. Preparation parses,
 * validates, clones, and freezes the rules, then returns a content-free capability
 * whose identity is held only in this module's private WeakMap. The readback
 * boundary is also a bounded JSON string, so verifier code never observes
 * adapter-owned objects, accessors, or proxies.
 *
 * Formula patterns are exact literals with the bounded `{{cell}}` template token;
 * model text is never compiled as a regular expression. Excel cell-text strings
 * are bounded to 32,767 UTF-16 code units at this verifier/host runtime boundary.
 */

import { isExcelCellTextWithinUtf16Limit } from "./workbookPlanContract";

export const MAX_VERIFICATION_RULES = 30;
export const MAX_FORMULA_PATTERN_LENGTH = 4096;
export const MAX_VERIFICATION_JSON_BYTES = 1_000_000;
export const MAX_READBACK_JSON_BYTES = 1_000_000;

const MAX_MATRIX_DIMENSION = 1000;
const FORMULA_TEMPLATE_TOKEN = "{{cell}}";
const MAX_EXCEL_COLUMN = 16_384;
const MAX_EXCEL_ROW = 1_048_576;

export type VerificationRuleKind =
  | "range_values_equal"
  | "range_formulas_equal"
  | "range_formula_pattern"
  | "range_number_format_equal"
  | "table_exists"
  | "table_row_order"
  | "table_filter_state"
  | "sheet_exists"
  | "chart_exists"
  | "comment_exists"
  | "sentinel_unchanged"
  | "calculation_completed";

export type ScalarValue = string | number | boolean | null;
export type ValueMatrix = readonly (readonly ScalarValue[])[];
export type FormulaMatrix = readonly (readonly string[])[];
export type StringMatrix = readonly (readonly string[])[];

export type VerificationRule =
  | { readonly kind: "range_values_equal"; readonly sheet: string; readonly address: string; readonly expectedValues: ValueMatrix }
  | { readonly kind: "range_formulas_equal"; readonly sheet: string; readonly address: string; readonly expectedFormulas: FormulaMatrix }
  | { readonly kind: "range_formula_pattern"; readonly sheet: string; readonly address: string; readonly pattern: string }
  | { readonly kind: "range_number_format_equal"; readonly sheet: string; readonly address: string; readonly expectedNumberFormat: string }
  | { readonly kind: "table_exists"; readonly tableName: string }
  | { readonly kind: "table_row_order"; readonly tableName: string; readonly expectedOrder: readonly string[] }
  | {
      readonly kind: "table_filter_state";
      readonly tableName: string;
      readonly columnIndex: number;
      readonly state: "filtered" | "unfiltered";
      readonly values?: readonly string[];
      readonly criterion1?: string;
      readonly criterion2?: string;
      readonly operator?: "And" | "Or";
    }
  | { readonly kind: "sheet_exists"; readonly name: string }
  | { readonly kind: "chart_exists"; readonly sheet: string; readonly title: string }
  | { readonly kind: "comment_exists"; readonly sheet: string; readonly address: string; readonly text: string }
  | { readonly kind: "sentinel_unchanged"; readonly sheet: string; readonly address: string; readonly expected: ScalarValue }
  | { readonly kind: "calculation_completed" };

export type FilterStateSnapshot = {
  readonly columnIndex: number;
  readonly state: "filtered" | "unfiltered";
  readonly values?: readonly string[];
  readonly criterion1?: string;
  readonly criterion2?: string;
  readonly operator?: "And" | "Or";
};

export type WorkbookSnapshot =
  | { readonly kind: "range_values"; readonly values: ValueMatrix; readonly volatile?: boolean }
  | { readonly kind: "range_formulas"; readonly formulas: FormulaMatrix; readonly volatile?: boolean }
  | { readonly kind: "range_number_format"; readonly numberFormat: StringMatrix }
  | { readonly kind: "table_exists"; readonly exists: boolean }
  | { readonly kind: "table_row_order"; readonly order: readonly string[] }
  | { readonly kind: "table_filter_state"; readonly state: FilterStateSnapshot }
  | { readonly kind: "sheet_exists"; readonly exists: boolean }
  | { readonly kind: "chart_exists"; readonly exists: boolean }
  | { readonly kind: "comment_exists"; readonly exists: boolean; readonly text?: string }
  | { readonly kind: "sentinel"; readonly value: ScalarValue }
  | { readonly kind: "calculation_completed"; readonly completed: boolean };

export type PreparedWorkbookPlanVerificationRules = Readonly<Record<never, never>>;

export type BoundWorkbookReadbackAdapter = Readonly<{
  readonly read: (rule: VerificationRule) => string | undefined;
  /** Index-aware readback prevents duplicate-rule identity collisions. */
  readonly readAt?: (ruleIndex: number, rule: VerificationRule) => string | undefined;
}>;

export type VerificationStatus = "verified" | "mismatched" | "unsupported" | "error";
export type VerificationReason =
  | "exact_match"
  | "value_mismatch"
  | "formula_mismatch"
  | "formula_pattern_mismatch"
  | "number_format_mismatch"
  | "row_order_mismatch"
  | "filter_state_mismatch"
  | "existence_mismatch"
  | "comment_text_mismatch"
  | "sentinel_mismatch"
  | "calculation_not_confirmed"
  | "dimension_mismatch"
  | "readback_kind_mismatch"
  | "pattern_too_large"
  | "invalid_readback"
  | "unsupported_host"
  | "unsupported_capability"
  | "volatile_readback"
  | "missing_readback"
  | "readback_error"
  | "rule_cap_exceeded";

type ReportRuleKind = VerificationRuleKind | "rule_cap";
export type VerificationEvidence = Readonly<{
  readonly ruleIndex: number;
  readonly kind: VerificationRuleKind;
  readonly status: VerificationStatus;
  readonly reason: VerificationReason;
}>;

export type VerificationCounts = Readonly<{
  readonly verified: number;
  readonly mismatched: number;
  readonly unsupported: number;
  readonly error: number;
}>;

export type VerificationFirstFailure = Readonly<{
  readonly ruleIndex: number;
  readonly kind: ReportRuleKind;
  readonly status: Exclude<VerificationStatus, "verified">;
  readonly reason: VerificationReason;
}>;

export type WorkbookPlanVerificationReport = Readonly<{
  readonly ok: boolean;
  readonly rules: readonly VerificationEvidence[];
  readonly counts: VerificationCounts;
  readonly firstFailure: VerificationFirstFailure | null;
}>;

export type VerificationInverseEvidence =
  | "not_captured"
  | "incomplete"
  | "complete_not_independently_tested"
  | "complete_and_independently_tested";

export type VerificationControlMetadata = Readonly<{
  readonly modelContinuation: false;
  readonly rollback: "forbidden" | "permitted";
  readonly inverseEvidence: VerificationInverseEvidence;
}>;

export const derivePostVerificationControl = (
  report: Pick<WorkbookPlanVerificationReport, "ok">,
  inverseEvidence: VerificationInverseEvidence,
): VerificationControlMetadata =>
  Object.freeze({
    modelContinuation: false as const,
    rollback: !report.ok && inverseEvidence === "complete_and_independently_tested" ? "permitted" as const : "forbidden" as const,
    inverseEvidence,
  });

const VERIFICATION_RULE_KINDS: readonly VerificationRuleKind[] = [
  "range_values_equal",
  "range_formulas_equal",
  "range_formula_pattern",
  "range_number_format_equal",
  "table_exists",
  "table_row_order",
  "table_filter_state",
  "sheet_exists",
  "chart_exists",
  "comment_exists",
  "sentinel_unchanged",
  "calculation_completed",
];
const SNAPSHOT_KINDS: readonly WorkbookSnapshot["kind"][] = [
  "range_values",
  "range_formulas",
  "range_number_format",
  "table_exists",
  "table_row_order",
  "table_filter_state",
  "sheet_exists",
  "chart_exists",
  "comment_exists",
  "sentinel",
  "calculation_completed",
];
const SAFE_FALLBACK_RULE_KIND: VerificationRuleKind = "calculation_completed";
const MAX_SHEET_NAME_LENGTH = 255;
const MAX_ADDRESS_LENGTH = 256;
const MAX_TABLE_NAME_LENGTH = 255;
const MAX_FORMULA_TEXT_LENGTH = 4096;
const MAX_NUMBER_FORMAT_LENGTH = 128;
const MAX_FILTER_VALUE_LENGTH = 512;
const MAX_FILTER_VALUES = 200;
const MAX_TABLE_ORDER_VALUE_LENGTH = 255;
const MAX_CHART_TITLE_LENGTH = 255;
const MAX_COMMENT_TEXT_LENGTH = 1024;
const FILTER_FIELDS = ["columnIndex", "state", "values", "criterion1", "criterion2", "operator"] as const;
type FilterField = (typeof FILTER_FIELDS)[number];
type MatrixDimensions = Readonly<{ readonly rows: number; readonly columns: number }>;
type JsonObject = Record<string, unknown>;

const preparedRules = new WeakMap<object, readonly VerificationRule[]>();

const byteLength = (value: string): number | null => {
  try {
    return new TextEncoder().encode(value).byteLength;
  } catch {
    return null;
  }
};

const isBoundedJson = (value: unknown, maxBytes: number): value is string =>
  typeof value === "string" && (byteLength(value) ?? maxBytes + 1) <= maxBytes;

const isPlainObject = (value: unknown): value is JsonObject =>
  value !== null && typeof value === "object" && !Array.isArray(value) && Object.getPrototypeOf(value) === Object.prototype;

const hasExactKeys = (value: unknown, required: readonly string[], optional: readonly string[] = []): JsonObject | null => {
  if (!isPlainObject(value)) return null;
  const allowed = new Set([...required, ...optional]);
  const keys = Object.keys(value);
  if (keys.some((key) => !allowed.has(key))) return null;
  if (required.some((key) => !Object.prototype.hasOwnProperty.call(value, key))) return null;
  return value;
};

const isBoundedString = (value: unknown, maxLength: number, minLength = 1): value is string =>
  typeof value === "string" && value.length >= minLength && value.length <= maxLength;

const isRuntimeScalar = (value: unknown): value is ScalarValue =>
  value === null || typeof value === "boolean" || typeof value === "string" && isExcelCellTextWithinUtf16Limit(value) || typeof value === "number" && Number.isFinite(value);

const inspectDenseArray = <T>(value: unknown, validator: (item: unknown) => item is T, allowEmpty = false, maxItems = MAX_MATRIX_DIMENSION): T[] | null => {
  if (!Array.isArray(value) || value.length > maxItems || value.length === 0 && !allowEmpty) return null;
  for (const key of Object.keys(value)) {
    const index = Number(key);
    if (!Number.isSafeInteger(index) || index < 0 || index >= value.length || String(index) !== key) return null;
  }
  const result: T[] = [];
  for (let index = 0; index < value.length; index += 1) {
    const item = value[index];
    if (!validator(item)) return null;
    result.push(item);
  }
  return result;
};

const inspectDenseMatrix = <T>(value: unknown, validator: (item: unknown) => item is T): T[][] | null => {
  const rows = inspectDenseArray(value, (row): row is unknown[] => Array.isArray(row));
  if (rows === null) return null;
  const result: T[][] = [];
  let columns: number | null = null;
  for (const row of rows) {
    const inspected = inspectDenseArray(row, validator);
    if (inspected === null) return null;
    if (columns === null) columns = inspected.length;
    if (inspected.length !== columns) return null;
    result.push(inspected);
  }
  return result;
};

const inspectStringSequence = (value: unknown, maxLength: number, allowEmpty = true, maxItems = MAX_MATRIX_DIMENSION): string[] | null =>
  inspectDenseArray(value, (item): item is string => isBoundedString(item, maxLength), allowEmpty, maxItems);

const isVerificationRuleKind = (value: unknown): value is VerificationRuleKind =>
  typeof value === "string" && VERIFICATION_RULE_KINDS.includes(value as VerificationRuleKind);

const isSnapshotKind = (value: unknown): value is WorkbookSnapshot["kind"] =>
  typeof value === "string" && SNAPSHOT_KINDS.includes(value as WorkbookSnapshot["kind"]);

const isRectangular = <T>(matrix: readonly (readonly T[])[]): MatrixDimensions | null => {
  if (matrix.length === 0 || matrix.length > MAX_MATRIX_DIMENSION) return null;
  const columns = matrix[0]?.length ?? 0;
  if (columns === 0 || columns > MAX_MATRIX_DIMENSION) return null;
  for (let row = 1; row < matrix.length; row += 1) if (matrix[row]?.length !== columns) return null;
  return { rows: matrix.length, columns };
};

const cloneMatrix = <T>(matrix: readonly (readonly T[])[]): readonly (readonly T[])[] =>
  Object.freeze(matrix.map((row) => Object.freeze([...row])));

const parseSheetPrefix = (text: string, start: number): number | null | undefined => {
  const exclamation = text.indexOf("!", start);
  if (exclamation < 0 || exclamation - start > 255) return undefined;
  if (text[start] === "'") {
    let index = start + 1;
    let contentLength = 0;
    while (index < exclamation) {
      if (text[index] === "\n" || text[index] === "\r") return null;
      if (text[index] === "'") {
        if (text[index + 1] === "'") {
          index += 2;
          contentLength += 1;
          continue;
        }
        return index === exclamation - 1 && contentLength > 0 ? exclamation + 1 : null;
      }
      index += 1;
      contentLength += 1;
    }
    return null;
  }
  const first = text[start] ?? "";
  const isLetter = (character: string): boolean => /^[A-Za-z]$/.test(character);
  const allowed = (character: string): boolean => isLetter(character) || /^[0-9]$/.test(character) || character === "_" || character === ".";
  if (exclamation === start || exclamation - start > 31 || !isLetter(first) && first !== "_") return null;
  for (let index = start + 1; index < exclamation; index += 1) if (!allowed(text[index] ?? "")) return null;
  return exclamation + 1;
};

const parseCellReferenceAt = (text: string, start: number): { readonly column: number; readonly row: number; readonly end: number } | null => {
  let index = start;
  const sheetEnd = parseSheetPrefix(text, start);
  if (sheetEnd === null) return null;
  if (sheetEnd !== undefined) index = sheetEnd;
  if (text[index] === "$") index += 1;
  const columnStart = index;
  let column = 0;
  while (index < text.length && /^[A-Za-z]$/.test(text[index] ?? "") && index - columnStart < 3) {
    const code = text.charCodeAt(index);
    column = column * 26 + (code >= 97 ? code - 96 : code - 64);
    index += 1;
  }
  if (index === columnStart || column > MAX_EXCEL_COLUMN) return null;
  if (text[index] === "$") index += 1;
  const rowStart = index;
  while (index < text.length && /^[0-9]$/.test(text[index] ?? "") && index - rowStart < 7) index += 1;
  if (index === rowStart || text[rowStart] === "0") return null;
  const row = Number(text.slice(rowStart, index));
  if (!Number.isSafeInteger(row) || row < 1 || row > MAX_EXCEL_ROW) return null;
  return { column, row, end: index };
};

const parseA1TargetDimensions = (address: string): MatrixDimensions | null => {
  if (!isBoundedString(address, MAX_ADDRESS_LENGTH)) return null;
  const first = parseCellReferenceAt(address, 0);
  if (first === null) return null;
  if (first.end === address.length) return { rows: 1, columns: 1 };
  if (address[first.end] !== ":") return null;
  const second = parseCellReferenceAt(address, first.end + 1);
  if (second === null || second.end !== address.length || second.column < first.column || second.row < first.row) return null;
  const rows = second.row - first.row + 1;
  const columns = second.column - first.column + 1;
  if (rows > MAX_MATRIX_DIMENSION || columns > MAX_MATRIX_DIMENSION) return null;
  return { rows, columns };
};

const isReferenceContinuation = (character: string): boolean => /^[A-Za-z0-9_$]$/.test(character);
const isBoundedCellReference = (formula: string, start: number): number | null => {
  const parsed = parseCellReferenceAt(formula, start);
  if (parsed === null || parsed.end < formula.length && isReferenceContinuation(formula[parsed.end] ?? "")) return null;
  return parsed.end;
};

const matchesFormulaPattern = (pattern: string, formula: string): boolean => {
  let patternIndex = 0;
  let formulaIndex = 0;
  while (patternIndex < pattern.length) {
    const tokenIndex = pattern.indexOf(FORMULA_TEMPLATE_TOKEN, patternIndex);
    if (tokenIndex < 0) return formula.slice(formulaIndex) === pattern.slice(patternIndex);
    const literal = pattern.slice(patternIndex, tokenIndex);
    if (formula.slice(formulaIndex, formulaIndex + literal.length) !== literal) return false;
    formulaIndex += literal.length;
    const referenceEnd = isBoundedCellReference(formula, formulaIndex);
    if (referenceEnd === null) return false;
    formulaIndex = referenceEnd;
    patternIndex = tokenIndex + FORMULA_TEMPLATE_TOKEN.length;
  }
  return formulaIndex === formula.length;
};

const normalizeFilterState = (value: unknown): FilterStateSnapshot | null => {
  const object = hasExactKeys(value, ["columnIndex", "state"], ["values", "criterion1", "criterion2", "operator"]);
  if (object === null) return null;
  const columnIndex = object.columnIndex;
  const state = object.state;
  if (typeof columnIndex !== "number" || !Number.isSafeInteger(columnIndex) || columnIndex < 0) return null;
  if (state !== "filtered" && state !== "unfiltered") return null;
  const result: { columnIndex: number; state: "filtered" | "unfiltered"; values?: readonly string[]; criterion1?: string; criterion2?: string; operator?: "And" | "Or" } = { columnIndex, state };
  if (Object.prototype.hasOwnProperty.call(object, "values")) {
    const values = inspectStringSequence(object.values, MAX_FILTER_VALUE_LENGTH, true, MAX_FILTER_VALUES);
    if (values === null) return null;
    result.values = Object.freeze(values);
  }
  for (const field of ["criterion1", "criterion2"] as const) {
    if (Object.prototype.hasOwnProperty.call(object, field)) {
      if (!isBoundedString(object[field], MAX_FILTER_VALUE_LENGTH, 0)) return null;
      result[field] = object[field] as string;
    }
  }
  if (Object.prototype.hasOwnProperty.call(object, "operator")) {
    if (object.operator !== "And" && object.operator !== "Or") return null;
    result.operator = object.operator;
  }
  return Object.freeze(result);
};

const normalizeSnapshot = (value: unknown): WorkbookSnapshot | null => {
  if (!isPlainObject(value)) return null;
  const kind = value.kind;
  if (!isSnapshotKind(kind)) return null;
  switch (kind) {
    case "range_values": {
      const object = hasExactKeys(value, ["kind", "values"], ["volatile"]);
      const values = object === null ? null : inspectDenseMatrix(object.values, isRuntimeScalar);
      if (object === null || values === null || Object.prototype.hasOwnProperty.call(object, "volatile") && typeof object.volatile !== "boolean") return null;
      return Object.freeze({ kind, values: cloneMatrix(values), ...(Object.prototype.hasOwnProperty.call(object, "volatile") ? { volatile: object.volatile as boolean } : {}) });
    }
    case "range_formulas": {
      const object = hasExactKeys(value, ["kind", "formulas"], ["volatile"]);
      const formulas = object === null ? null : inspectDenseMatrix(object.formulas, (item): item is string => isBoundedString(item, MAX_FORMULA_TEXT_LENGTH, 0));
      if (object === null || formulas === null || Object.prototype.hasOwnProperty.call(object, "volatile") && typeof object.volatile !== "boolean") return null;
      return Object.freeze({ kind, formulas: cloneMatrix(formulas), ...(Object.prototype.hasOwnProperty.call(object, "volatile") ? { volatile: object.volatile as boolean } : {}) });
    }
    case "range_number_format": {
      const object = hasExactKeys(value, ["kind", "numberFormat"]);
      const numberFormat = object === null ? null : inspectDenseMatrix(object.numberFormat, (item): item is string => isBoundedString(item, MAX_NUMBER_FORMAT_LENGTH));
      return object === null || numberFormat === null ? null : Object.freeze({ kind, numberFormat: cloneMatrix(numberFormat) });
    }
    case "table_exists":
    case "sheet_exists":
    case "chart_exists": {
      const object = hasExactKeys(value, ["kind", "exists"]);
      return object !== null && typeof object.exists === "boolean" ? Object.freeze({ kind, exists: object.exists }) : null;
    }
    case "table_row_order": {
      const object = hasExactKeys(value, ["kind", "order"]);
      const order = object === null ? null : inspectStringSequence(object.order, MAX_TABLE_ORDER_VALUE_LENGTH, false);
      return object === null || order === null ? null : Object.freeze({ kind, order: Object.freeze(order) });
    }
    case "table_filter_state": {
      const object = hasExactKeys(value, ["kind", "state"]);
      const state = object === null ? null : normalizeFilterState(object.state);
      return object === null || state === null ? null : Object.freeze({ kind, state });
    }
    case "comment_exists": {
      const object = hasExactKeys(value, ["kind", "exists"], ["text"]);
      if (object === null || typeof object.exists !== "boolean") return null;
      if (Object.prototype.hasOwnProperty.call(object, "text") && !isBoundedString(object.text, MAX_COMMENT_TEXT_LENGTH)) return null;
      return Object.freeze({ kind, exists: object.exists, ...(Object.prototype.hasOwnProperty.call(object, "text") ? { text: object.text as string } : {}) });
    }
    case "sentinel": {
      const object = hasExactKeys(value, ["kind", "value"]);
      return object !== null && isRuntimeScalar(object.value) ? Object.freeze({ kind, value: object.value }) : null;
    }
    case "calculation_completed": {
      const object = hasExactKeys(value, ["kind", "completed"]);
      return object !== null && typeof object.completed === "boolean" ? Object.freeze({ kind, completed: object.completed }) : null;
    }
  }
};

const normalizeReadback = (value: unknown): WorkbookReadbackResult | undefined | null => {
  if (value === undefined) return undefined;
  if (!isPlainObject(value) || typeof value.status !== "string") return null;
  if (value.status === "unsupported") {
    const object = hasExactKeys(value, ["status", "reason"]);
    return object !== null && (object.reason === "unsupported_host" || object.reason === "unsupported_capability" || object.reason === "volatile_readback") ? { status: "unsupported", reason: object.reason } : null;
  }
  if (value.status === "error") {
    const object = hasExactKeys(value, ["status", "reason"]);
    return object !== null && object.reason === "readback_error" ? { status: "error", reason: "readback_error" } : null;
  }
  if (value.status === "supported") {
    const object = hasExactKeys(value, ["status", "snapshot"]);
    const snapshot = object === null ? null : normalizeSnapshot(object.snapshot);
    return object === null || snapshot === null ? null : { status: "supported", snapshot };
  }
  return null;
};

type WorkbookReadbackResult =
  | { readonly status: "supported"; readonly snapshot: WorkbookSnapshot }
  | { readonly status: "unsupported"; readonly reason: "unsupported_host" | "unsupported_capability" | "volatile_readback" }
  | { readonly status: "error"; readonly reason: "readback_error" };

const cloneRule = (value: unknown): VerificationRule | null => {
  if (!isPlainObject(value) || !isVerificationRuleKind(value.kind)) return null;
  const kind = value.kind;
  const requiredString = (field: string, maxLength: number): string | null => isBoundedString(value[field], maxLength) ? value[field] as string : null;
  const range = (): { readonly sheet: string; readonly address: string; readonly dimensions: MatrixDimensions } | null => {
    const sheet = requiredString("sheet", MAX_SHEET_NAME_LENGTH);
    const address = requiredString("address", MAX_ADDRESS_LENGTH);
    const dimensions = address === null ? null : parseA1TargetDimensions(address);
    return sheet === null || address === null || dimensions === null ? null : { sheet, address, dimensions };
  };
  switch (kind) {
    case "range_values_equal": {
      const object = hasExactKeys(value, ["kind", "sheet", "address", "expectedValues"]);
      const target = object === null ? null : range();
      const matrix = object === null ? null : inspectDenseMatrix(object.expectedValues, isRuntimeScalar);
      const dimensions = matrix === null ? null : isRectangular(matrix);
      if (object === null || target === null || matrix === null || dimensions === null || dimensions.rows !== target.dimensions.rows || dimensions.columns !== target.dimensions.columns) return null;
      return Object.freeze({ kind, sheet: target.sheet, address: target.address, expectedValues: cloneMatrix(matrix) });
    }
    case "range_formulas_equal": {
      const object = hasExactKeys(value, ["kind", "sheet", "address", "expectedFormulas"]);
      const target = object === null ? null : range();
      const matrix = object === null ? null : inspectDenseMatrix(object.expectedFormulas, (item): item is string => isBoundedString(item, MAX_FORMULA_TEXT_LENGTH, 0));
      const dimensions = matrix === null ? null : isRectangular(matrix);
      if (object === null || target === null || matrix === null || dimensions === null || dimensions.rows !== target.dimensions.rows || dimensions.columns !== target.dimensions.columns) return null;
      return Object.freeze({ kind, sheet: target.sheet, address: target.address, expectedFormulas: cloneMatrix(matrix) });
    }
    case "range_formula_pattern": {
      const object = hasExactKeys(value, ["kind", "sheet", "address", "pattern"]);
      const target = object === null ? null : range();
      const pattern = object === null ? null : requiredString("pattern", MAX_FORMULA_TEXT_LENGTH);
      return object === null || target === null || pattern === null || pattern.length > MAX_FORMULA_PATTERN_LENGTH ? null : Object.freeze({ kind, sheet: target.sheet, address: target.address, pattern });
    }
    case "range_number_format_equal": {
      const object = hasExactKeys(value, ["kind", "sheet", "address", "expectedNumberFormat"]);
      const target = object === null ? null : range();
      const expectedNumberFormat = object === null ? null : requiredString("expectedNumberFormat", MAX_NUMBER_FORMAT_LENGTH);
      return object === null || target === null || expectedNumberFormat === null ? null : Object.freeze({ kind, sheet: target.sheet, address: target.address, expectedNumberFormat });
    }
    case "table_exists": {
      const object = hasExactKeys(value, ["kind", "tableName"]);
      const tableName = object === null ? null : requiredString("tableName", MAX_TABLE_NAME_LENGTH);
      return object === null || tableName === null ? null : Object.freeze({ kind, tableName });
    }
    case "table_row_order": {
      const object = hasExactKeys(value, ["kind", "tableName", "expectedOrder"]);
      const tableName = object === null ? null : requiredString("tableName", MAX_TABLE_NAME_LENGTH);
      const expectedOrder = object === null ? null : inspectStringSequence(object.expectedOrder, MAX_TABLE_ORDER_VALUE_LENGTH, false);
      return object === null || tableName === null || expectedOrder === null ? null : Object.freeze({ kind, tableName, expectedOrder: Object.freeze(expectedOrder) });
    }
    case "table_filter_state": {
      const object = hasExactKeys(value, ["kind", "tableName", "columnIndex", "state"], ["values", "criterion1", "criterion2", "operator"]);
      if (object === null) return null;
      const tableName = requiredString("tableName", MAX_TABLE_NAME_LENGTH);
      if (tableName === null || typeof object.columnIndex !== "number" || !Number.isSafeInteger(object.columnIndex) || object.columnIndex < 0 || object.state !== "filtered" && object.state !== "unfiltered") return null;
      const result: { kind: "table_filter_state"; tableName: string; columnIndex: number; state: "filtered" | "unfiltered"; values?: readonly string[]; criterion1?: string; criterion2?: string; operator?: "And" | "Or" } = { kind, tableName, columnIndex: object.columnIndex, state: object.state };
      if (Object.prototype.hasOwnProperty.call(object, "values")) {
        const values = inspectStringSequence(object.values, MAX_FILTER_VALUE_LENGTH, true, MAX_FILTER_VALUES);
        if (values === null) return null;
        result.values = Object.freeze(values);
      }
      for (const field of ["criterion1", "criterion2"] as const) {
        if (Object.prototype.hasOwnProperty.call(object, field)) {
          if (!isBoundedString(object[field], MAX_FILTER_VALUE_LENGTH, 0)) return null;
          result[field] = object[field] as string;
        }
      }
      if (Object.prototype.hasOwnProperty.call(object, "operator")) {
        if (object.operator !== "And" && object.operator !== "Or") return null;
        result.operator = object.operator;
      }
      return Object.freeze(result);
    }
    case "sheet_exists": {
      const object = hasExactKeys(value, ["kind", "name"]);
      const name = object === null ? null : requiredString("name", MAX_SHEET_NAME_LENGTH);
      return object === null || name === null ? null : Object.freeze({ kind, name });
    }
    case "chart_exists": {
      const object = hasExactKeys(value, ["kind", "sheet", "title"]);
      const sheet = object === null ? null : requiredString("sheet", MAX_SHEET_NAME_LENGTH);
      const title = object === null ? null : requiredString("title", MAX_CHART_TITLE_LENGTH);
      return object === null || sheet === null || title === null ? null : Object.freeze({ kind, sheet, title });
    }
    case "comment_exists": {
      const object = hasExactKeys(value, ["kind", "sheet", "address", "text"]);
      const target = object === null ? null : range();
      const text = object === null ? null : requiredString("text", MAX_COMMENT_TEXT_LENGTH);
      return object === null || target === null || target.dimensions.rows !== 1 || target.dimensions.columns !== 1 || text === null ? null : Object.freeze({ kind, sheet: target.sheet, address: target.address, text });
    }
    case "sentinel_unchanged": {
      const object = hasExactKeys(value, ["kind", "sheet", "address", "expected"]);
      const target = object === null ? null : range();
      return object === null || target === null || target.dimensions.rows !== 1 || target.dimensions.columns !== 1 || !isRuntimeScalar(object.expected) ? null : Object.freeze({ kind, sheet: target.sheet, address: target.address, expected: object.expected });
    }
    case "calculation_completed":
      return hasExactKeys(value, ["kind"]) === null ? null : Object.freeze({ kind });
  }
};

export const prepareWorkbookPlanVerificationRules = (rawJson: string): PreparedWorkbookPlanVerificationRules | null => {
  if (!isBoundedJson(rawJson, MAX_VERIFICATION_JSON_BYTES)) return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(rawJson) as unknown;
  } catch {
    return null;
  }
  const list = inspectDenseArray(parsed, (item): item is VerificationRule => item !== undefined, false, MAX_VERIFICATION_RULES);
  if (list === null || list.length < 1 || list.length > MAX_VERIFICATION_RULES) return null;
  const cloned = list.map((rule) => cloneRule(rule));
  if (cloned.some((rule): rule is null => rule === null)) return null;
  const token = Object.freeze(Object.create(null)) as PreparedWorkbookPlanVerificationRules;
  preparedRules.set(token, Object.freeze(cloned as VerificationRule[]));
  return token;
};

const hasVolatileReadback = (snapshot: WorkbookSnapshot): boolean =>
  (snapshot.kind === "range_values" || snapshot.kind === "range_formulas") && snapshot.volatile === true;

const expectedFilterFieldPresent = (rule: Extract<VerificationRule, { readonly kind: "table_filter_state" }>, field: FilterField): boolean =>
  field === "columnIndex" || field === "state" || Object.prototype.hasOwnProperty.call(rule, field);

const actualFilterFieldPresent = (state: FilterStateSnapshot, field: FilterField): boolean => Object.prototype.hasOwnProperty.call(state, field);

const filterStatesEqual = (rule: Extract<VerificationRule, { readonly kind: "table_filter_state" }>, actual: FilterStateSnapshot): boolean => {
  const normalizedRule = rule.criterion1 === undefined ? rule : { ...rule, operator: rule.operator ?? "And" as const };
  const normalizedActual = actual.criterion1 === undefined ? actual : { ...actual, operator: actual.operator ?? "And" as const };
  const expectedFields = FILTER_FIELDS.filter((field) => expectedFilterFieldPresent(normalizedRule, field));
  const actualFields = FILTER_FIELDS.filter((field) => actualFilterFieldPresent(normalizedActual, field));
  if (expectedFields.length !== actualFields.length || expectedFields.some((field) => !actualFields.includes(field))) return false;
  if (normalizedRule.columnIndex !== normalizedActual.columnIndex || normalizedRule.state !== normalizedActual.state) return false;
  if (normalizedRule.values !== undefined && !sameStringSequence(normalizedRule.values, normalizedActual.values ?? [])) return false;
  if (normalizedRule.criterion1 !== normalizedActual.criterion1 || normalizedRule.criterion2 !== normalizedActual.criterion2 || normalizedRule.operator !== normalizedActual.operator) return false;
  return true;
};

const sameStringSequence = (left: readonly string[], right: readonly string[]): boolean => {
  if (left.length !== right.length || left.length > MAX_MATRIX_DIMENSION || right.length > MAX_MATRIX_DIMENSION) return false;
  for (let index = 0; index < left.length; index += 1) if (left[index] !== right[index]) return false;
  return true;
};

const expectedSnapshotKindFor = (rule: VerificationRule): WorkbookSnapshot["kind"] => {
  switch (rule.kind) {
    case "range_values_equal": return "range_values";
    case "range_formulas_equal":
    case "range_formula_pattern": return "range_formulas";
    case "range_number_format_equal": return "range_number_format";
    case "table_exists": return "table_exists";
    case "table_row_order": return "table_row_order";
    case "table_filter_state": return "table_filter_state";
    case "sheet_exists": return "sheet_exists";
    case "chart_exists": return "chart_exists";
    case "comment_exists": return "comment_exists";
    case "sentinel_unchanged": return "sentinel";
    case "calculation_completed": return "calculation_completed";
  }
};

const ruleMatchesSnapshot = (rule: VerificationRule, snapshot: WorkbookSnapshot): VerificationReason | "matched" => {
  if (hasVolatileReadback(snapshot)) return "volatile_readback";
  if (snapshot.kind !== expectedSnapshotKindFor(rule)) return "readback_kind_mismatch";
  switch (rule.kind) {
    case "range_values_equal": {
      if (snapshot.kind !== "range_values") return "readback_kind_mismatch";
      const expected = isRectangular(rule.expectedValues);
      const actual = isRectangular(snapshot.values);
      if (expected === null || actual === null || expected.rows !== actual.rows || expected.columns !== actual.columns) return "dimension_mismatch";
      return matricesEqual(rule.expectedValues, snapshot.values) ? "matched" : "value_mismatch";
    }
    case "range_formulas_equal": {
      if (snapshot.kind !== "range_formulas") return "readback_kind_mismatch";
      const expected = isRectangular(rule.expectedFormulas);
      const actual = isRectangular(snapshot.formulas);
      if (expected === null || actual === null || expected.rows !== actual.rows || expected.columns !== actual.columns) return "dimension_mismatch";
      return matricesEqual(rule.expectedFormulas, snapshot.formulas) ? "matched" : "formula_mismatch";
    }
    case "range_formula_pattern": {
      if (snapshot.kind !== "range_formulas") return "readback_kind_mismatch";
      const target = parseA1TargetDimensions(rule.address);
      const actual = isRectangular(snapshot.formulas);
      if (target === null || actual === null || target.rows !== actual.rows || target.columns !== actual.columns) return "dimension_mismatch";
      if (rule.pattern.length > MAX_FORMULA_PATTERN_LENGTH) return "pattern_too_large";
      for (const row of snapshot.formulas) for (const formula of row) if (!matchesFormulaPattern(rule.pattern, formula)) return "formula_pattern_mismatch";
      return "matched";
    }
    case "range_number_format_equal": {
      if (snapshot.kind !== "range_number_format") return "readback_kind_mismatch";
      const target = parseA1TargetDimensions(rule.address);
      const actual = isRectangular(snapshot.numberFormat);
      if (target === null || actual === null || target.rows !== actual.rows || target.columns !== actual.columns) return "dimension_mismatch";
      for (const row of snapshot.numberFormat) for (const value of row) if (value !== rule.expectedNumberFormat) return "number_format_mismatch";
      return "matched";
    }
    case "table_exists": return snapshot.kind === "table_exists" && snapshot.exists ? "matched" : snapshot.kind === "table_exists" ? "existence_mismatch" : "readback_kind_mismatch";
    case "table_row_order": return snapshot.kind === "table_row_order" ? sameStringSequence(rule.expectedOrder, snapshot.order) ? "matched" : "row_order_mismatch" : "readback_kind_mismatch";
    case "table_filter_state": return snapshot.kind === "table_filter_state" ? filterStatesEqual(rule, snapshot.state) ? "matched" : "filter_state_mismatch" : "readback_kind_mismatch";
    case "sheet_exists": return snapshot.kind === "sheet_exists" && snapshot.exists ? "matched" : snapshot.kind === "sheet_exists" ? "existence_mismatch" : "readback_kind_mismatch";
    case "chart_exists": return snapshot.kind === "chart_exists" && snapshot.exists ? "matched" : snapshot.kind === "chart_exists" ? "existence_mismatch" : "readback_kind_mismatch";
    case "comment_exists":
      if (snapshot.kind !== "comment_exists") return "readback_kind_mismatch";
      if (!snapshot.exists) return "existence_mismatch";
      if (snapshot.text === undefined) return "invalid_readback";
      return snapshot.text === rule.text ? "matched" : "comment_text_mismatch";
    case "sentinel_unchanged": return snapshot.kind === "sentinel" && Object.is(rule.expected, snapshot.value) ? "matched" : snapshot.kind === "sentinel" ? "sentinel_mismatch" : "readback_kind_mismatch";
    case "calculation_completed": return snapshot.kind === "calculation_completed" && snapshot.completed ? "matched" : snapshot.kind === "calculation_completed" ? "calculation_not_confirmed" : "readback_kind_mismatch";
  }
};

const matricesEqual = <T>(left: readonly (readonly T[])[], right: readonly (readonly T[])[]): boolean => {
  const leftDimensions = isRectangular(left);
  const rightDimensions = isRectangular(right);
  if (leftDimensions === null || rightDimensions === null || leftDimensions.rows !== rightDimensions.rows || leftDimensions.columns !== rightDimensions.columns) return false;
  for (let row = 0; row < leftDimensions.rows; row += 1) for (let column = 0; column < leftDimensions.columns; column += 1) if (!Object.is(left[row]?.[column], right[row]?.[column])) return false;
  return true;
};

const emptyReport = (): WorkbookPlanVerificationReport => Object.freeze({
  ok: false,
  rules: Object.freeze([Object.freeze({ ruleIndex: 0, kind: SAFE_FALLBACK_RULE_KIND, status: "error" as const, reason: "invalid_readback" as const })]),
  counts: Object.freeze({ verified: 0, mismatched: 0, unsupported: 0, error: 1 }),
  firstFailure: Object.freeze({ ruleIndex: 0, kind: SAFE_FALLBACK_RULE_KIND, status: "error" as const, reason: "invalid_readback" as const }),
});

export const verifyWorkbookPlan = (
  prepared: PreparedWorkbookPlanVerificationRules,
  adapter: BoundWorkbookReadbackAdapter,
): WorkbookPlanVerificationReport => {
  const rules = prepared !== null && (typeof prepared === "object" || typeof prepared === "function") ? preparedRules.get(prepared) : undefined;
  if (rules === undefined) return emptyReport();
  const evidence: VerificationEvidence[] = [];
  const counts = { verified: 0, mismatched: 0, unsupported: 0, error: 0 };
  let firstFailure: VerificationFirstFailure | null = null;
  const append = (item: VerificationEvidence): void => {
    counts[item.status] += 1;
    evidence.push(Object.freeze({ ...item }));
    if (firstFailure === null && item.status !== "verified") firstFailure = Object.freeze({ ruleIndex: item.ruleIndex, kind: item.kind, status: item.status, reason: item.reason });
  };
  for (let ruleIndex = 0; ruleIndex < rules.length; ruleIndex += 1) {
    const rule = rules[ruleIndex]!;
    let item: VerificationEvidence;
    try {
      const rawReadback = adapter.readAt === undefined
        ? adapter.read(rule)
        : adapter.readAt(ruleIndex, rule);
      const readback = normalizeReadback(rawReadback === undefined ? undefined : isBoundedJson(rawReadback, MAX_READBACK_JSON_BYTES) ? (() => { try { return JSON.parse(rawReadback) as unknown; } catch { return null; } })() : null);
      if (readback === undefined) item = { ruleIndex, kind: rule.kind, status: "unsupported", reason: "missing_readback" };
      else if (readback === null) item = { ruleIndex, kind: rule.kind, status: "error", reason: "invalid_readback" };
      else if (readback.status === "unsupported") item = { ruleIndex, kind: rule.kind, status: "unsupported", reason: readback.reason };
      else if (readback.status === "error") item = { ruleIndex, kind: rule.kind, status: "error", reason: readback.reason };
      else {
        const match = ruleMatchesSnapshot(rule, readback.snapshot);
        if (match === "matched") item = { ruleIndex, kind: rule.kind, status: "verified", reason: "exact_match" };
        else if (match === "volatile_readback") item = { ruleIndex, kind: rule.kind, status: "unsupported", reason: match };
        else if (match === "dimension_mismatch") item = { ruleIndex, kind: rule.kind, status: "mismatched", reason: match };
        else if (match === "invalid_readback" || match === "readback_kind_mismatch" || match === "pattern_too_large") item = { ruleIndex, kind: rule.kind, status: "error", reason: match };
        else item = { ruleIndex, kind: rule.kind, status: "mismatched", reason: match };
      }
    } catch {
      item = { ruleIndex, kind: rule.kind, status: "error", reason: "readback_error" };
    }
    append(item);
  }
  return Object.freeze({ ok: firstFailure === null, rules: Object.freeze([...evidence]), counts: Object.freeze({ ...counts }), firstFailure });
};
