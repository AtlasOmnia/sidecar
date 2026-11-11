/**
 * R4c: one trusted primitive mutation authority per task-pane turn.
 *
 * TaskPaneTurnCoordinator.beginTurn mints exactly one authority for every pane
 * turn, in both WorkbookPlan rollout modes, and terminates it when the turn is
 * finished, stopped or replaced. The direct agent loop executes every Excel
 * primitive through `executePrimitive`.
 *
 * - Rollout off: a pure pass-through. Only a terminated authority refuses
 *   (TURN_INACTIVE); otherwise the raw dispatcher is called synchronously and
 *   its promise is returned unchanged. No effect table, accounting, signature,
 *   latch, serialization or readback is consulted.
 * - Rollout on: the I4 rule order (intent, legacy / no plan path tools,
 *   duplicate, per-turn resources incl. readback cost, the plan-entered and
 *   readback latches, and the write_range/autofill_formula post-write check),
 *   serialized per authority; a failed primitive is refunded only if it never
 *   reached Office; otherwise the turn latches.
 *
 * The raw name -> tool.execute dispatcher is claimed once, here, at module
 * load; the exported excelTools/officeBridgeExcel sinks are read-only.
 */
import {
  claimRawPrimitiveExecutor,
  getOfficeBatchEntryCount,
  getPrimitiveEffect,
  toolSchemas,
} from "./excelTools";
import { accountPrimitiveResourceIncrement } from "./workbookPlanResources";
import { decidePrimitiveIntent, type WorkbookPlanResourceSummary } from "./workbookPlanPolicy";
import type { TrustedWorkbookRequest } from "./workbookPlanRequest";
import type { WorkbookOperation } from "./generated/workbookPlanSchema";

type RawPrimitiveExecutor = (name: string, args: Record<string, unknown>) => Promise<string>;

const raw: RawPrimitiveExecutor = claimRawPrimitiveExecutor();

/** Fixed denial texts (I6). A denial never carries argument content. */
export const PRIMITIVE_DENIAL_MESSAGES = Object.freeze({
  READ_ONLY_SINK: "Workbook edits must go through the task-pane edit authority.",
  TURN_INACTIVE: "Workbook edit denied: no active task-pane turn.",
  UNCLASSIFIED: "Workbook edit denied: this tool is not classified.",
  LEGACY_NO_PLAN_PATH: "Workbook edit denied: this legacy tool has no WorkbookPlan path and is disabled while WorkbookPlan rollout is on.",
  CONFIRMATION_REQUIRED: "Workbook edit denied: this request needs confirmation. Use excel_execute_plan (WorkbookPlan) instead.",
  DUPLICATE: "Workbook edit denied: the identical edit was already applied in this turn.",
  RESOURCE_LIMIT: "Workbook edit denied: the per-turn workbook edit limit was reached.",
  READBACK_FAILED: "Workbook edit denied: an earlier edit in this turn did not verify.",
  PLAN_ENTERED: "Workbook edit denied: a workbook plan already ran in this turn.",
});

type PrimitiveDenialCode = keyof typeof PRIMITIVE_DENIAL_MESSAGES;

const denial = (code: PrimitiveDenialCode): string => JSON.stringify({
  ok: false,
  error: PRIMITIVE_DENIAL_MESSAGES[code],
  category: code === "DUPLICATE" ? "repetition" : "validation",
});

const DENIAL: Readonly<Record<PrimitiveDenialCode, string>> = Object.freeze({
  READ_ONLY_SINK: denial("READ_ONLY_SINK"),
  TURN_INACTIVE: denial("TURN_INACTIVE"),
  UNCLASSIFIED: denial("UNCLASSIFIED"),
  LEGACY_NO_PLAN_PATH: denial("LEGACY_NO_PLAN_PATH"),
  CONFIRMATION_REQUIRED: denial("CONFIRMATION_REQUIRED"),
  DUPLICATE: denial("DUPLICATE"),
  RESOURCE_LIMIT: denial("RESOURCE_LIMIT"),
  READBACK_FAILED: denial("READBACK_FAILED"),
  PLAN_ENTERED: denial("PLAN_ENTERED"),
});

/** Schema property names per known primitive; decisions read only these keys. */
const PRIMITIVE_SCHEMA_KEYS: ReadonlyMap<string, readonly string[]> = new Map(
  toolSchemas.map((schema) => {
    const properties = (schema.function.parameters as { properties?: Record<string, unknown> }).properties ?? {};
    return [schema.function.name, Object.freeze(Object.keys(properties).sort())] as const;
  }),
);

const ZERO_USAGE: WorkbookPlanResourceSummary = Object.freeze({
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

const POST_WRITE_CHECKED = new Set(["write_range", "autofill_formula"]);
const DUPLICATE_EXEMPT = new Set(["run_calculation"]);

function own(args: Record<string, unknown>, key: string): unknown {
  return Object.prototype.hasOwnProperty.call(args, key) ? args[key] : undefined;
}

function stripSheetPrefix(value: unknown): unknown {
  return typeof value === "string" && value.includes("!") ? value.slice(value.lastIndexOf("!") + 1) : value;
}

const ADDRESS_KEYS: ReadonlySet<string> = new Set(["address", "sourceAddress", "targetAddress", "dataRange"]);
const A1_CELL = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$/;

function columnNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function columnLetters(n: number): string {
  let letters = "";
  for (let rest = n; rest > 0; rest = Math.floor((rest - 1) / 26)) letters = String.fromCharCode(65 + ((rest - 1) % 26)) + letters;
  return letters;
}

/**
 * R4e-2 duplicate-signature spelling of an A1 cell or rectangle on `sheet`
 * (no `$`, uppercase, no matching sheet prefix, normalized corners, single
 * cell collapsed). Anything else (names, whole rows/columns, R1C1, another
 * sheet, non-strings) is returned unchanged. Never throws.
 */
function canonicalAddress(sheet: unknown, value: unknown): unknown {
  if (typeof value !== "string") return value;
  let text = value.trim();
  const bang = text.lastIndexOf("!");
  if (bang >= 0) {
    const quoted = text.slice(0, bang);
    const prefix = /^'(?:[^']|'')+'$/.test(quoted) ? quoted.slice(1, -1).replace(/''/g, "'") : quoted;
    if (typeof sheet !== "string" || prefix !== sheet) return value;
    text = text.slice(bang + 1);
  }
  const parts = text.split(":");
  if (parts.length < 1 || parts.length > 2) return value;
  const columns: number[] = [];
  const rows: number[] = [];
  for (const part of parts) {
    const match = A1_CELL.exec(part);
    if (match === null) return value;
    const column = columnNumber(match[1]);
    const row = Number(match[2]);
    if (column < 1 || column > 16384 || row < 1 || row > 1048576) return value;
    columns.push(column);
    rows.push(row);
  }
  const first = `${columnLetters(Math.min(...columns))}${Math.min(...rows)}`;
  const last = `${columnLetters(Math.max(...columns))}${Math.max(...rows)}`;
  return first === last ? first : `${first}:${last}`;
}

function canonicalJson(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonicalJson);
  if (value !== null && typeof value === "object") {
    const record = value as Record<string, unknown>;
    return Object.keys(record).sort().map((key) => [key, canonicalJson(record[key])]);
  }
  return value;
}

function primitiveSignature(name: string, args: Record<string, unknown>): string {
  const keys = PRIMITIVE_SCHEMA_KEYS.get(name) ?? [];
  const values = keys.map((key) => own(args, key));
  const sheetIndex = keys.indexOf("sheet");
  const sheet = sheetIndex < 0 ? undefined : values[sheetIndex];
  return JSON.stringify([name, keys.map((key, i) => {
    let value = values[i];
    if (ADDRESS_KEYS.has(key)) {
      try { value = canonicalAddress(sheet, value); } catch { value = values[i]; }
    }
    return [key, canonicalJson(value)];
  })]);
}

/** I8 accounting descriptor for one mutating primitive (schema keys only, sheet prefixes stripped). */
function describePrimitive(name: string, args: Record<string, unknown>): WorkbookOperation | undefined {
  const sheet = own(args, "sheet");
  switch (name) {
    case "write_range":
      return { kind: "write_values", sheet, address: stripSheetPrefix(own(args, "address")), values: own(args, "values") } as unknown as WorkbookOperation;
    case "autofill_formula":
      return {
        kind: "fill_formula",
        sheet,
        sourceAddress: stripSheetPrefix(own(args, "sourceAddress")),
        targetAddress: stripSheetPrefix(own(args, "targetAddress")),
      } as unknown as WorkbookOperation;
    case "format_range":
      return { kind: "format_range", sheet, address: stripSheetPrefix(own(args, "address")) } as unknown as WorkbookOperation;
    case "conditional_format_range":
      return { kind: "conditional_format", sheet, address: stripSheetPrefix(own(args, "address")) } as unknown as WorkbookOperation;
    case "freeze_panes":
      return { kind: "freeze_panes" } as unknown as WorkbookOperation;
    case "set_data_validation":
      return { kind: "set_validation_list", sheet, address: stripSheetPrefix(own(args, "address")), source: own(args, "source") } as unknown as WorkbookOperation;
    case "insert_table":
      return { kind: "create_table", sheet, address: stripSheetPrefix(own(args, "address")) } as unknown as WorkbookOperation;
    case "sort_table":
      return { kind: "sort_table" } as unknown as WorkbookOperation;
    case "filter_table":
      return { kind: "filter_table" } as unknown as WorkbookOperation;
    case "create_chart":
      return { kind: "create_chart", sheet, dataRange: stripSheetPrefix(own(args, "dataRange")) } as unknown as WorkbookOperation;
    case "create_sheet":
      return { kind: "create_sheet" } as unknown as WorkbookOperation;
    case "copy_sheet":
      return { kind: "copy_sheet" } as unknown as WorkbookOperation;
    case "rename_sheet":
      return { kind: "rename_sheet" } as unknown as WorkbookOperation;
    case "add_comment":
      return { kind: "add_comment", sheet, address: stripSheetPrefix(own(args, "address")), text: own(args, "text") } as unknown as WorkbookOperation;
    case "run_calculation":
      return { kind: "calculate" } as unknown as WorkbookOperation;
    default:
      return undefined;
  }
}

function parseJsonObject(text: string): Record<string, unknown> | undefined {
  try {
    const parsed: unknown = JSON.parse(text);
    return parsed !== null && typeof parsed === "object" && !Array.isArray(parsed) ? parsed as Record<string, unknown> : undefined;
  } catch {
    return undefined;
  }
}

function normalizeFormula(value: string): string {
  // Protected formula text is compared exactly, not parsed or normalized.
  // Structured-reference headers and apostrophe escapes can contain spaces,
  // operators, quotes and brackets. Preserve every character (including case)
  // in the WHOLE formula if any bracket/quote occurs. Excel table/column names
  // are case-insensitive, but a spelling-only mismatch here conservatively
  // means unverified. False negatives are preferable to a false verified edit.
  if (value.includes("[") || value.includes("]") || value.includes('"') || value.includes("'")) return value;
  // Never erase operand-separating whitespace (Excel's intersection operator).
  // Only discard optional spacing directly beside explicit scalar operators.
  // quoted string/sheet contents (and doubled quote escapes) byte-for-byte;
  // only unquoted ASCII spelling is case-insensitive.
  let quote: string | undefined;
  let normalized = "";
  const text = value.trim();
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (quote !== undefined) {
      normalized += ch;
      if (ch === quote) {
        if (text[i + 1] === quote) { normalized += text[++i]; }
        else quote = undefined;
      }
    } else if (ch === '"' || ch === "'") {
      quote = ch;
      normalized += ch;
    } else if (/\s/.test(ch)) {
      let end = i + 1;
      while (end < text.length && /\s/.test(text[end])) end += 1;
      const scalarOperator = /[=+\-*/^&<>]/;
      if (!scalarOperator.test(text[i - 1] ?? "") && !scalarOperator.test(text[end] ?? "")) normalized += text.slice(i, end);
      i = end - 1;
    } else {
      normalized += /[a-z]/.test(ch) ? ch.toUpperCase() : ch;
    }
  }
  return normalized;
}

function decimalValue(value: string): number | undefined {
  const text = value.trim();
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(text)) return undefined;
  const number = Number(text);
  return Number.isFinite(number) ? number : undefined;
}

/** OQ-C13 tolerant comparison of one written cell with its readback formula. */
function cellMatches(expected: unknown, actual: unknown): boolean {
  if (expected === null || expected === "") return actual === "" || actual === null;
  if (typeof expected === "number") {
    return (typeof actual === "number" && actual === expected) ||
      (typeof actual === "string" && decimalValue(actual) === expected);
  }
  if (typeof expected === "boolean") {
    return actual === expected || (typeof actual === "string" && actual.toUpperCase() === String(expected).toUpperCase());
  }
  if (typeof expected === "string") {
    if (expected.startsWith("=")) return typeof actual === "string" && normalizeFormula(actual) === normalizeFormula(expected);
    if (actual === expected) return true;
    // Permit only a specific reproducible numeric/Boolean conversion. Dates,
    // locale formats and arbitrary typed contents cannot prove text equality.
    if (typeof actual === "number") return decimalValue(expected) === actual;
    if (typeof actual === "boolean") return expected.toUpperCase() === String(actual).toUpperCase();
    return false;
  }
  return false;
}

function readbackMatches(name: string, args: Record<string, unknown>, readback: string): boolean {
  const parsed = parseJsonObject(readback);
  if (parsed === undefined || parsed.ok === false || parsed.truncated === true) return false;
  const formulas = parsed.formulas;
  if (!Array.isArray(formulas) || formulas.length === 0 || !formulas.every((row) => Array.isArray(row) && row.length > 0)) return false;
  const grid = formulas as unknown[][];
  if (name === "autofill_formula") {
    return grid.every((row) => row.every((cell) => typeof cell === "string" && cell.startsWith("=")));
  }
  const expected = own(args, "values");
  if (!Array.isArray(expected) || expected.length !== grid.length) return false;
  for (let row = 0; row < grid.length; row += 1) {
    const expectedRow: unknown = expected[row];
    if (!Array.isArray(expectedRow) || expectedRow.length !== grid[row].length) return false;
    for (let column = 0; column < expectedRow.length; column += 1) {
      if (!cellMatches(expectedRow[column], grid[row][column])) return false;
    }
  }
  return true;
}

export interface PrimitiveMutationAuthorityContext {
  readonly trustedRequest: TrustedWorkbookRequest | null;
  readonly workbookPlanRolloutEnabled: boolean;
}

const CONSTRUCTION_TOKEN = Symbol("primitive-mutation-authority");

export class PrimitiveMutationAuthority {
  private readonly request: TrustedWorkbookRequest | null;
  private readonly rollout: boolean;
  private terminal = false;
  private planEntered = false;
  private readbackFailed = false;
  private usage: WorkbookPlanResourceSummary = ZERO_USAGE;
  private readonly signatures = new Set<string>();
  private chain: Promise<unknown> = Promise.resolve();

  constructor(token: symbol, context: PrimitiveMutationAuthorityContext) {
    if (token !== CONSTRUCTION_TOKEN) throw new Error("Use createPrimitiveMutationAuthority.");
    this.rollout = context.workbookPlanRolloutEnabled === true;
    const request = context.trustedRequest;
    this.request = request === null
      ? null
      : Object.freeze({ classification: request.classification, selectionScopedRequest: request.selectionScopedRequest === true });
  }

  get isTerminal(): boolean {
    return this.terminal;
  }

  get rolloutEnabled(): boolean {
    return this.rollout;
  }

  usageSnapshot(): WorkbookPlanResourceSummary {
    return Object.freeze({ ...this.usage });
  }

  markWorkbookPlanEntered(): void {
    this.planEntered = true;
  }

  terminate(): void {
    this.terminal = true;
  }

  executePrimitive(name: string, args: Record<string, unknown>): Promise<string> {
    if (this.terminal) return Promise.resolve(DENIAL.TURN_INACTIVE);
    if (!this.rolloutEnabled) return raw(name, args);
    const result = this.chain.then(() => this.decideAndExecute(name, args));
    this.chain = result.catch(() => undefined);
    return result;
  }

  private async decideAndExecute(name: string, args: Record<string, unknown>): Promise<string> {
    if (this.terminal) return DENIAL.TURN_INACTIVE;
    if (!PRIMITIVE_SCHEMA_KEYS.has(name)) return raw(name, args);
    const effect = getPrimitiveEffect(name);
    if (effect === undefined) return DENIAL.UNCLASSIFIED;
    if (effect.effectClass === "read_only") return raw(name, args);
    if (effect.effectClass !== "view") {
      if (this.planEntered) return DENIAL.PLAN_ENTERED;
      if (this.readbackFailed) return DENIAL.READBACK_FAILED;
    }
    if (effect.effectClass === "view") return raw(name, args);
    // Legacy / no plan path tools are disabled while rollout is on.
    if (effect.effectClass === "destructive_legacy") return DENIAL.LEGACY_NO_PLAN_PATH;
    if (effect.effectClass !== "mutating" || effect.planEffect === undefined) return DENIAL.UNCLASSIFIED;

    const request = this.request;
    if (request === null) return DENIAL.CONFIRMATION_REQUIRED;
    const intent = decidePrimitiveIntent({
      requestClassification: request.classification,
      selectionScopedRequest: request.selectionScopedRequest,
      effect: effect.planEffect,
    });
    if (intent.decision !== "approved") return DENIAL.CONFIRMATION_REQUIRED;

    const signature = primitiveSignature(name, args);
    if (!DUPLICATE_EXEMPT.has(name) && this.signatures.has(signature)) return DENIAL.DUPLICATE;

    const next = this.chargeResources(name, args);
    if (next === undefined) return DENIAL.RESOURCE_LIMIT;

    // A2 (R4d) narrowed by R4e-2: a failed primitive (throw or non-ok:true) is
    // refunded only when no Office batch began during the raw call; otherwise
    // the charge and signature are kept and the turn latches (as after a
    // failed readback). Atomic: decisions run one at a time on this.chain.
    const previousUsage = this.usage;
    const signatureAdded = !this.signatures.has(signature);
    this.usage = next;
    this.signatures.add(signature);
    if (this.terminal) return DENIAL.TURN_INACTIVE;

    const officeBatchesBefore = getOfficeBatchEntryCount();
    let result: string;
    try {
      result = await raw(name, args);
    } catch (error) {
      this.settleFailure(officeBatchesBefore, previousUsage, signatureAdded, signature);
      throw error;
    }
    if (parseJsonObject(result)?.ok !== true) {
      this.settleFailure(officeBatchesBefore, previousUsage, signatureAdded, signature);
      return result;
    }
    if (POST_WRITE_CHECKED.has(name) && !this.terminal && parseJsonObject(result)?.ok === true) {
      await this.verifyWrite(name, args);
    }
    return result;
  }

  private settleFailure(officeBatchesBefore: number, previousUsage: WorkbookPlanResourceSummary, signatureAdded: boolean, signature: string): void {
    if (getOfficeBatchEntryCount() === officeBatchesBefore) {
      this.usage = previousUsage;
      if (signatureAdded) this.signatures.delete(signature);
    } else {
      this.readbackFailed = true;
    }
  }

  private chargeResources(name: string, args: Record<string, unknown>): WorkbookPlanResourceSummary | undefined {
    try {
      const operation = describePrimitive(name, args);
      if (operation === undefined) return undefined;
      let readbackCells = 0;
      let readbackReads = 0;
      if (POST_WRITE_CHECKED.has(name)) {
        const target = accountPrimitiveResourceIncrement(ZERO_USAGE, operation, { readbackCells: 0, readbackReads: 0, unknownTableBounds: "operation_only" });
        if (target.ok !== true) return undefined;
        readbackCells = target.usage.cellsWritten;
        readbackReads = 1;
      }
      const accounted = accountPrimitiveResourceIncrement(this.usage, operation, { readbackCells, readbackReads, unknownTableBounds: "operation_only" });
      return accounted.ok === true ? Object.freeze({ ...accounted.usage }) : undefined;
    } catch {
      return undefined;
    }
  }

  private async verifyWrite(name: string, args: Record<string, unknown>): Promise<void> {
    const address = own(args, name === "autofill_formula" ? "targetAddress" : "address");
    let readback: string;
    try {
      readback = await raw("read_range", { sheet: own(args, "sheet"), address });
    } catch {
      this.readbackFailed = true;
      return;
    }
    if (!readbackMatches(name, args, readback)) this.readbackFailed = true;
  }
}

export function createPrimitiveMutationAuthority(context: PrimitiveMutationAuthorityContext): PrimitiveMutationAuthority {
  return new PrimitiveMutationAuthority(CONSTRUCTION_TOKEN, context);
}

export interface ServerPrimitiveGateDependencies {
  readonly getBridgeSessionId: () => string | null;
  readonly getActiveTurn: () => { readonly turnOperationId: string | null; readonly sessionId: string | null };
  readonly resolveAuthority: (input: { readonly turnId: string; readonly bridgeSessionId: string }) => PrimitiveMutationAuthority | null;
  readonly executeWithAuthority: (name: string, args: Record<string, unknown>, authority: PrimitiveMutationAuthority) => Promise<string>;
  readonly executeReadOnly: (name: string, args: Record<string, unknown>) => Promise<string>;
}

/**
 * R4d: the server-owned route's primitive executor (D-R4-3; OQ-D1 "always",
 * in both rollout modes). A primitive ticket runs through the live pane turn's
 * authority only when that turn is on the current bridge session. Otherwise
 * read-only primitives and unknown names take the read-only sink and every
 * other primitive is denied (TURN_INACTIVE) without validation, argument read
 * or Office call. Rollout semantics come from the turn's authority snapshot.
 */
export function createServerPrimitiveGate(deps: ServerPrimitiveGateDependencies): (name: string, args: Record<string, unknown>) => Promise<string> {
  return (name, args) => {
    const bridgeSessionId = deps.getBridgeSessionId();
    const active = deps.getActiveTurn();
    const authority = bridgeSessionId !== null && active.turnOperationId !== null && active.sessionId === bridgeSessionId
      ? deps.resolveAuthority({ turnId: active.turnOperationId, bridgeSessionId })
      : null;
    if (authority !== null && !authority.isTerminal) return deps.executeWithAuthority(name, args, authority);
    if (!PRIMITIVE_SCHEMA_KEYS.has(name) || getPrimitiveEffect(name)?.effectClass === "read_only") {
      return deps.executeReadOnly(name, args);
    }
    return Promise.resolve(DENIAL.TURN_INACTIVE);
  };
}
