import type {
  InspectionOperation,
  WorkbookOperation,
  WorkbookPlan,
} from "./generated/workbookPlanSchema";
import { workbookPlanContract } from "./workbookPlanContract";
import {
  validateWorkbookPlan,
  type WorkbookPlanValidationError,
} from "./workbookPlanValidation";

export interface WorkbookPlanCanonicalizationError {
  readonly code:
    | "alias_conflict"
    | "symbol_key_not_supported"
    | "accessor_property"
    | "non_plain_object"
    | "sparse_array"
    | "work_limit_exceeded"
    | "duplicate_write";
  readonly path: string;
  readonly message: string;
  readonly firstPath?: string;
}

export interface CanonicalizeWorkbookPlanSyntaxFailure {
  readonly ok: false;
  readonly stage: "syntax";
  readonly errors: readonly WorkbookPlanCanonicalizationError[];
}

export interface CanonicalizeWorkbookPlanValidationFailure {
  readonly ok: false;
  readonly stage: "validation";
  readonly errors: readonly WorkbookPlanValidationError[];
}

export interface CanonicalizeWorkbookPlanSuccess {
  readonly ok: true;
  readonly stage: "success";
  readonly value: WorkbookPlan;
  readonly semanticDigest: string;
  readonly operationDigests: readonly string[];
  readonly readDigests: readonly string[];
}

export type CanonicalizeWorkbookPlanResult =
  | CanonicalizeWorkbookPlanSyntaxFailure
  | CanonicalizeWorkbookPlanValidationFailure
  | CanonicalizeWorkbookPlanSuccess;

type SyntaxNormalizationSuccess = {
  readonly ok: true;
  readonly value: unknown;
};

type SyntaxNormalizationFailure = {
  readonly ok: false;
  readonly errors: readonly WorkbookPlanCanonicalizationError[];
};

type SyntaxNormalizationResult = SyntaxNormalizationSuccess | SyntaxNormalizationFailure;

interface CanonicalizationBudget {
  workUnits: number;
  visitedNodes: number;
  exhausted: boolean;
}

const encoder = new TextEncoder();
const aliasMap: Readonly<Record<string, string>> = workbookPlanContract.validationConstraints.aliasMap;
const aliasKeys = Object.keys(aliasMap).sort(compareCodeUnit);
const mutatingOperationKinds: ReadonlySet<WorkbookOperation["kind"]> = new Set(
  workbookPlanContract.operationDefinitions
    .filter(({ effect }) => effect !== "calculate")
    .map(({ kind }) => kind),
);

function hasOwn(value: object, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(value, key);
}

function ownDataStringValue(value: object, key: string): string | undefined {
  let descriptor: PropertyDescriptor | undefined;
  try {
    descriptor = Object.getOwnPropertyDescriptor(value, key);
  } catch {
    return undefined;
  }

  if (descriptor === undefined || !hasOwn(descriptor, "value") || typeof descriptor.value !== "string") {
    return undefined;
  }

  return descriptor.value;
}

function compareCodeUnit(left: string, right: string): number {
  const limit = Math.min(left.length, right.length);
  for (let index = 0; index < limit; index += 1) {
    const difference = left.charCodeAt(index) - right.charCodeAt(index);
    if (difference !== 0) return difference;
  }
  return left.length - right.length;
}

function escapePointerSegment(segment: string): string {
  return segment.split("~").join("~0").split("/").join("~1");
}

function joinPath(base: string, segment: string | number): string {
  return `${base}/${escapePointerSegment(String(segment))}`;
}

function createBudget(): CanonicalizationBudget {
  return { workUnits: 0, visitedNodes: 0, exhausted: false };
}

function noteWork(budget: CanonicalizationBudget, workUnits = 1): boolean {
  if (budget.exhausted) return false;
  budget.workUnits += workUnits;
  if (budget.workUnits > workbookPlanContract.validationConstraints.maxWorkUnits) {
    budget.exhausted = true;
    return false;
  }
  return true;
}

function noteNode(budget: CanonicalizationBudget): boolean {
  if (budget.exhausted) return false;
  budget.visitedNodes += 1;
  if (budget.visitedNodes > workbookPlanContract.validationConstraints.maxVisitedNodes) {
    budget.exhausted = true;
    return false;
  }
  return noteWork(budget);
}

function syntaxError(
  code: WorkbookPlanCanonicalizationError["code"],
  path: string,
  message: string,
  firstPath?: string,
): SyntaxNormalizationFailure {
  return {
    ok: false,
    errors: [
      {
        code,
        path,
        message,
        ...(firstPath === undefined ? {} : { firstPath }),
      },
    ],
  };
}

function workLimitFailure(): SyntaxNormalizationFailure {
  return syntaxError("work_limit_exceeded", "", "Canonicalization work limit exceeded");
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (value === null || typeof value !== "object" || Array.isArray(value)) return false;
  try {
    return Object.getPrototypeOf(value) === Object.prototype;
  } catch {
    return false;
  }
}

function isPlainArray(value: unknown): value is readonly unknown[] {
  if (!Array.isArray(value)) return false;
  try {
    return Object.getPrototypeOf(value) === Array.prototype;
  } catch {
    return false;
  }
}

function canonicalArrayIndex(key: string, length: number): number | undefined {
  if (key.length === 0) return undefined;
  const index = Number(key);
  return Number.isInteger(index) && index >= 0 && index < length && String(index) === key
    ? index
    : undefined;
}

function normalizeSyntaxValue(
  value: unknown,
  path: string,
  depth: number,
  budget: CanonicalizationBudget,
): SyntaxNormalizationResult {
  if (depth > workbookPlanContract.validationConstraints.maxDepth) {
    return workLimitFailure();
  }

  if (value === null || typeof value !== "object") {
    return { ok: true, value };
  }

  if (Array.isArray(value)) {
    return normalizeSyntaxArray(value, path, depth, budget);
  }

  return normalizeSyntaxObject(value, path, depth, budget);
}

function normalizeSyntaxObject(
  value: object,
  path: string,
  depth: number,
  budget: CanonicalizationBudget,
): SyntaxNormalizationResult {
  if (!isPlainRecord(value)) {
    return syntaxError("non_plain_object", path, "Canonicalization requires a plain object");
  }

  if (!noteNode(budget)) {
    return workLimitFailure();
  }

  let ownKeys: readonly PropertyKey[];
  try {
    ownKeys = Reflect.ownKeys(value);
  } catch {
    return syntaxError("non_plain_object", path, "Canonicalization requires a plain object");
  }

  if (!noteWork(budget, ownKeys.length)) {
    return workLimitFailure();
  }

  const stringKeys: string[] = [];
  const keyPresence = new Set<string>();
  for (const key of ownKeys) {
    if (typeof key !== "string") {
      return syntaxError("symbol_key_not_supported", path, "Symbol keys are not supported");
    }
    stringKeys.push(key);
    keyPresence.add(key);
  }

  for (const alias of aliasKeys) {
    const canonical = ownDataStringValue(aliasMap, alias);
    if (canonical !== undefined && keyPresence.has(alias) && keyPresence.has(canonical)) {
      return syntaxError(
        "alias_conflict",
        joinPath(path, alias),
        "Canonical and alias cannot both be present",
      );
    }
  }

  stringKeys.sort(compareCodeUnit);
  const normalizedEntries: Array<{ readonly key: string; readonly value: unknown }> = [];

  for (const key of stringKeys) {
    let descriptor: PropertyDescriptor | undefined;
    try {
      descriptor = Object.getOwnPropertyDescriptor(value, key);
    } catch {
      return syntaxError("non_plain_object", path, "Canonicalization requires a plain object");
    }

    if (descriptor === undefined) continue;
    if (!hasOwn(descriptor, "value")) {
      return syntaxError("accessor_property", joinPath(path, key), "Accessor properties are not supported");
    }

    const child = normalizeSyntaxValue(descriptor.value, joinPath(path, key), depth + 1, budget);
    if (child.ok === false) return child;

    normalizedEntries.push({
      key: ownDataStringValue(aliasMap, key) ?? key,
      value: child.value,
    });
  }

  normalizedEntries.sort((left, right) => compareCodeUnit(left.key, right.key));

  const result: Record<string, unknown> = {};
  for (const entry of normalizedEntries) {
    Object.defineProperty(result, entry.key, {
      value: entry.value,
      enumerable: true,
      configurable: true,
      writable: true,
    });
  }

  return { ok: true, value: result };
}

function normalizeSyntaxArray(
  value: readonly unknown[],
  path: string,
  depth: number,
  budget: CanonicalizationBudget,
): SyntaxNormalizationResult {
  if (!isPlainArray(value)) {
    return syntaxError("non_plain_object", path, "Canonicalization requires a plain object");
  }

  if (!noteNode(budget)) {
    return workLimitFailure();
  }

  let ownKeys: readonly PropertyKey[];
  try {
    ownKeys = Reflect.ownKeys(value);
  } catch {
    return syntaxError("non_plain_object", path, "Canonicalization requires a plain object");
  }

  if (!noteWork(budget, ownKeys.length)) {
    return workLimitFailure();
  }

  const lengthDescriptor = Object.getOwnPropertyDescriptor(value, "length");
  if (
    lengthDescriptor === undefined ||
    !hasOwn(lengthDescriptor, "value") ||
    typeof lengthDescriptor.value !== "number" ||
    !Number.isInteger(lengthDescriptor.value) ||
    lengthDescriptor.value < 0
  ) {
    return syntaxError("non_plain_object", path, "Canonicalization requires a plain object");
  }

  const length = lengthDescriptor.value;
  const values = new Map<number, unknown>();
  const stringKeys: string[] = [];

  for (const key of ownKeys) {
    if (typeof key !== "string") {
      return syntaxError("symbol_key_not_supported", path, "Symbol keys are not supported");
    }
    stringKeys.push(key);
  }

  stringKeys.sort(compareCodeUnit);
  for (const key of stringKeys) {
    if (key === "length") continue;

    let descriptor: PropertyDescriptor | undefined;
    try {
      descriptor = Object.getOwnPropertyDescriptor(value, key);
    } catch {
      return syntaxError("non_plain_object", path, "Canonicalization requires a plain object");
    }

    if (descriptor === undefined) continue;
    if (!hasOwn(descriptor, "value")) {
      return syntaxError("accessor_property", joinPath(path, key), "Accessor properties are not supported");
    }

    const index = canonicalArrayIndex(key, length);
    if (index === undefined) {
      return syntaxError("sparse_array", joinPath(path, key), "Sparse arrays are not supported");
    }

    const child = normalizeSyntaxValue(descriptor.value, joinPath(path, index), depth + 1, budget);
    if (child.ok === false) return child;

    values.set(index, child.value);
  }

  const result = new Array<unknown>(length);
  for (let index = 0; index < length; index += 1) {
    if (!values.has(index)) {
      return syntaxError("sparse_array", joinPath(path, index), "Sparse arrays are not supported");
    }
    result[index] = values.get(index);
  }

  return { ok: true, value: result };
}

function serializeCanonicalJson(value: unknown): string {
  if (value === null) return "null";

  switch (typeof value) {
    case "string":
      return JSON.stringify(value);
    case "number":
      return JSON.stringify(value);
    case "boolean":
      return value ? "true" : "false";
    case "object":
      if (Array.isArray(value)) {
        const items: string[] = [];
        for (let index = 0; index < value.length; index += 1) {
          items.push(serializeCanonicalJson(value[index]));
        }
        return `[${items.join(",")}]`;
      }
      if (!isPlainRecord(value)) {
        throw new Error("Canonical serialization requires a plain object");
      }
      return `{${serializeCanonicalObjectEntries(value)}}`;
    default:
      return "null";
  }
}

function serializeCanonicalObjectEntries(value: Record<string, unknown>): string {
  const keys = Reflect.ownKeys(value)
    .filter((key): key is string => typeof key === "string")
    .sort(compareCodeUnit);
  const pieces: string[] = [];

  for (const key of keys) {
    let descriptor: PropertyDescriptor | undefined;
    try {
      descriptor = Object.getOwnPropertyDescriptor(value, key);
    } catch {
      throw new Error("Canonical serialization requires a plain object");
    }

    if (descriptor === undefined || !hasOwn(descriptor, "value") || descriptor.value === undefined) continue;
    pieces.push(`${JSON.stringify(key)}:${serializeCanonicalJson(descriptor.value)}`);
  }

  return pieces.join(",");
}

async function sha256Hex(text: string): Promise<string> {
  const digest = await globalThis.crypto.subtle.digest("SHA-256", encoder.encode(text));
  const bytes = new Uint8Array(digest);
  let hex = "";
  for (let index = 0; index < bytes.length; index += 1) {
    hex += bytes[index].toString(16).padStart(2, "0");
  }
  return hex;
}

async function digestCanonicalValue(value: unknown): Promise<string> {
  return sha256Hex(serializeCanonicalJson(value));
}

async function digestCanonicalArray(values: readonly unknown[]): Promise<readonly string[]> {
  const digests: string[] = [];
  for (const value of values) {
    digests.push(await digestCanonicalValue(value));
  }
  return digests;
}

function duplicateWriteError(path: string, firstPath: string): SyntaxNormalizationFailure {
  return syntaxError("duplicate_write", path, "Duplicate mutating operation", firstPath);
}

async function findDuplicateMutatingOperation(
  operations: readonly WorkbookOperation[],
): Promise<SyntaxNormalizationFailure | undefined> {
  const seen = new Map<string, string>();

  for (let index = 0; index < operations.length; index += 1) {
    const operation = operations[index];
    const digest = await digestCanonicalValue(operation);
    const path = `/operations/${index}`;
    const kind = operation.kind;

    if (mutatingOperationKinds.has(kind)) {
      const firstPath = seen.get(digest);
      if (firstPath !== undefined) {
        return duplicateWriteError(path, firstPath);
      }
      seen.set(digest, path);
    }
  }

  return undefined;
}

export async function canonicalizeWorkbookPlan(
  input: unknown,
): Promise<CanonicalizeWorkbookPlanResult> {
  const normalization = normalizeSyntaxValue(input, "", 0, createBudget());
  if (normalization.ok === false) {
    return {
      ok: false,
      stage: "syntax",
      errors: normalization.errors,
    };
  }

  const validation = validateWorkbookPlan(normalization.value);
  if (validation.ok === false) {
    return {
      ok: false,
      stage: "validation",
      errors: validation.errors,
    };
  }

  const plan = validation.value;
  let operations: readonly WorkbookOperation[] = [];
  let reads: readonly InspectionOperation[] = [];
  switch (plan.outcome) {
    case "execute":
      operations = plan.operations;
      break;
    case "inspect":
      reads = plan.reads;
      break;
    default:
      break;
  }

  const duplicate = await findDuplicateMutatingOperation(operations);
  if (duplicate !== undefined) {
    return {
      ok: false,
      stage: "syntax",
      errors: duplicate.errors,
    };
  }

  const semanticDigest = await digestCanonicalValue(plan);
  const operationDigests = await digestCanonicalArray(operations);
  const readDigests = await digestCanonicalArray(reads);

  return {
    ok: true,
    stage: "success",
    value: plan,
    semanticDigest,
    operationDigests,
    readDigests,
  };
}
