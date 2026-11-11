/**
 * Recursive, content-free interpreter for the WorkbookPlan V1 contract.
 *
 * The model-facing schema remains the source of truth for structure. Validation
 * metadata in workbookPlanContract owns the bounded rules that JSON Schema does
 * not express, such as byte limits, aliases, and rectangular matrices.
 */

import type { WorkbookPlan } from "./generated/workbookPlanSchema";
import {
  EXCEL_CELL_TEXT_MAX_LENGTH,
  isExcelCellTextWithinUtf16Limit,
  workbookPlanContract,
  type ValidationUtf8LimitName,
  type WorkbookPlanValidationConstraints,
} from "./workbookPlanContract";

type SchemaNode = Readonly<Record<string, unknown>> & {
  readonly type?: string;
  readonly const?: unknown;
  readonly enum?: readonly unknown[];
  readonly required?: readonly string[];
  readonly additionalProperties?: boolean;
  readonly properties?: Readonly<Record<string, SchemaNode>>;
  readonly oneOf?: readonly SchemaNode[];
  readonly items?: SchemaNode;
  readonly minItems?: number;
  readonly maxItems?: number;
  readonly minLength?: number;
  readonly maxLength?: number;
  readonly minimum?: number;
  readonly maximum?: number;
  readonly exclusiveMinimum?: number;
  readonly exclusiveMaximum?: number;
  readonly $ref?: string;
  readonly $defs?: Readonly<Record<string, SchemaNode>>;
};

type InputObject = Readonly<Record<string, unknown>>;

export interface WorkbookPlanValidationError {
  readonly code:
    | "unknown_property"
    | "missing_required"
    | "type_mismatch"
    | "bound_exceeded"
    | "matrix_jagged"
    | "utf8_limit_exceeded"
    | "alias_conflict"
    | "unsupported_version"
    | "unknown_discriminator"
    | "oneof_mismatch"
    | "symbol_key_not_supported"
    | "accessor_property"
    | "work_limit_exceeded";
  readonly path: string;
  readonly message: string;
  readonly limit?: number;
}

export interface WorkbookPlanSuccess {
  readonly ok: true;
  readonly value: WorkbookPlan;
}

export interface WorkbookPlanFailure {
  readonly ok: false;
  readonly errors: readonly WorkbookPlanValidationError[];
}

export type ValidateWorkbookPlanResult = WorkbookPlanSuccess | WorkbookPlanFailure;

interface ValidationBudget {
  workUnits: number;
  visitedNodes: number;
  exhausted: boolean;
}

interface ValidationContext {
  readonly rootSchema: SchemaNode;
  readonly constraints: WorkbookPlanValidationConstraints;
  readonly errors: WorkbookPlanValidationError[];
  readonly seenErrors: Set<string>;
  readonly budget: ValidationBudget;
  failureCount: number;
}

interface ValidationOptions {
  readonly ignoreOneOf?: boolean;
  readonly scalarValue?: boolean;
}

interface Discriminator {
  readonly property: string;
  readonly values: readonly unknown[];
}

const hasOwn = (value: object, key: string): boolean =>
  Object.prototype.hasOwnProperty.call(value, key);

const isPlainObject = (value: unknown): value is InputObject =>
  value !== null &&
  typeof value === "object" &&
  !Array.isArray(value) &&
  Object.getPrototypeOf(value) === Object.prototype;

const isSchemaNode = (value: unknown): value is SchemaNode =>
  value !== null && typeof value === "object" && !Array.isArray(value);

const asSchemaNode = (value: unknown): SchemaNode | undefined =>
  isSchemaNode(value) ? value : undefined;

const schemaProperties = (schema: SchemaNode): Readonly<Record<string, SchemaNode>> =>
  schema.properties ?? {};

const schemaNodes = (value: unknown): readonly SchemaNode[] => {
  if (!Array.isArray(value)) return [];
  return value.map(asSchemaNode).filter((node): node is SchemaNode => node !== undefined);
};

const schemaStrings = (value: unknown): readonly string[] => {
  if (!Array.isArray(value)) return [];
  return value.filter((item): item is string => typeof item === "string");
};

const escapePointerSegment = (segment: string): string =>
  segment.replaceAll("~", "~0").replaceAll("/", "~1");

const joinPath = (base: string, segment: string | number): string =>
  `${base}/${escapePointerSegment(String(segment))}`;

const ownDataValue = (value: object, key: string): unknown | undefined => {
  const descriptor = Object.getOwnPropertyDescriptor(value, key);
  if (descriptor === undefined || !Object.prototype.hasOwnProperty.call(descriptor, "value")) {
    return undefined;
  }
  return descriptor.value;
};

const ownDataDescriptor = (value: object, key: string): PropertyDescriptor | undefined => {
  const descriptor = Object.getOwnPropertyDescriptor(value, key);
  return descriptor !== undefined && Object.prototype.hasOwnProperty.call(descriptor, "value")
    ? descriptor
    : undefined;
};

const arrayLength = (value: readonly unknown[]): number | undefined => {
  const descriptor = ownDataDescriptor(value, "length");
  return descriptor !== undefined &&
    typeof descriptor.value === "number" &&
    Number.isInteger(descriptor.value) &&
    descriptor.value >= 0
    ? descriptor.value
    : undefined;
};

const canonicalArrayIndex = (key: string, length: number): number | undefined => {
  if (key.length === 0) return undefined;
  const index = Number(key);
  return Number.isInteger(index) && index >= 0 && index < length && String(index) === key
    ? index
    : undefined;
};

const compareCodeUnit = (left: string, right: string): number => {
  const limit = Math.min(left.length, right.length);
  for (let index = 0; index < limit; index += 1) {
    const difference = left.charCodeAt(index) - right.charCodeAt(index);
    if (difference !== 0) return difference;
  }
  return left.length - right.length;
};

const compareOwnPropertyKeys = (left: PropertyKey, right: PropertyKey): number => {
  const leftIsString = typeof left === "string";
  const rightIsString = typeof right === "string";
  if (leftIsString && rightIsString) return compareCodeUnit(left, right);
  if (leftIsString) return -1;
  if (rightIsString) return 1;
  return 0;
};

const noteWork = (context: ValidationContext, workUnits = 1): boolean => {
  if (context.budget.exhausted) return false;
  context.budget.workUnits += workUnits;
  if (context.budget.workUnits > context.constraints.maxWorkUnits) {
    context.budget.exhausted = true;
    return false;
  }
  return true;
};

const noteNode = (context: ValidationContext): boolean => {
  if (context.budget.exhausted) return false;
  context.budget.visitedNodes += 1;
  if (context.budget.visitedNodes > context.constraints.maxVisitedNodes) {
    context.budget.exhausted = true;
    return false;
  }
  return noteWork(context);
};

const exceedsUtf8ByteLimit = (value: string, limit: number): boolean => {
  let bytes = 0;
  for (let index = 0; index < value.length && bytes <= limit; index += 1) {
    const first = value.charCodeAt(index);
    if (first <= 0x7f) {
      bytes += 1;
    } else if (first <= 0x7ff) {
      bytes += 2;
    } else if (first >= 0xd800 && first <= 0xdbff) {
      const second = value.charCodeAt(index + 1);
      if (second >= 0xdc00 && second <= 0xdfff) {
        bytes += 4;
        index += 1;
      } else {
        bytes += 3;
      }
    } else {
      bytes += 3;
    }
  }
  return bytes > limit;
};

const boundedCodePointLength = (value: string, limit: number): number => {
  let length = 0;
  for (let index = 0; index < value.length && length <= limit; index += 1) {
    const first = value.charCodeAt(index);
    if (first >= 0xd800 && first <= 0xdbff) {
      const second = value.charCodeAt(index + 1);
      if (second >= 0xdc00 && second <= 0xdfff) index += 1;
    }
    length += 1;
  }
  return length;
};

const descriptorSafeDeepEqual = (left: unknown, right: unknown): boolean => {
  if (Object.is(left, right)) return true;
  if (Array.isArray(left) && Array.isArray(right)) {
    const leftLength = ownDataValue(left, "length");
    const rightLength = ownDataValue(right, "length");
    if (typeof leftLength !== "number" || leftLength !== rightLength) return false;
    for (let index = 0; index < leftLength; index += 1) {
      const leftItem = ownDataDescriptor(left, String(index));
      const rightItem = ownDataDescriptor(right, String(index));
      if (leftItem === undefined || rightItem === undefined) return false;
      if (!descriptorSafeDeepEqual(leftItem.value, rightItem.value)) return false;
    }
    return true;
  }
  if (!isPlainObject(left) || !isPlainObject(right)) return false;

  const leftKeys = Reflect.ownKeys(left).sort(compareOwnPropertyKeys);
  const rightKeys = Reflect.ownKeys(right).sort(compareOwnPropertyKeys);
  if (leftKeys.length !== rightKeys.length) return false;
  for (let index = 0; index < leftKeys.length; index += 1) {
    const leftKey = leftKeys[index];
    const rightKey = rightKeys[index];
    if (leftKey !== rightKey) return false;
    const leftProperty = Object.getOwnPropertyDescriptor(left, leftKey);
    const rightProperty = Object.getOwnPropertyDescriptor(right, rightKey);
    if (
      leftProperty === undefined ||
      rightProperty === undefined ||
      !Object.prototype.hasOwnProperty.call(leftProperty, "value") ||
      !Object.prototype.hasOwnProperty.call(rightProperty, "value") ||
      !descriptorSafeDeepEqual(leftProperty.value, rightProperty.value)
    ) {
      return false;
    }
  }
  return true;
};

const makeError = (
  code: WorkbookPlanValidationError["code"],
  path: string,
  message: string,
  limit?: number,
): WorkbookPlanValidationError => ({
  code,
  path,
  message,
  ...(limit === undefined ? {} : { limit }),
});

const errorKey = (error: WorkbookPlanValidationError): string =>
  `${error.path}\u0000${error.code}\u0000${error.message}\u0000${error.limit ?? ""}`;

const pushError = (
  context: ValidationContext,
  error: WorkbookPlanValidationError,
): void => {
  if (context.budget.exhausted) return;
  context.failureCount += 1;
  const key = errorKey(error);
  if (context.seenErrors.has(key)) return;
  context.seenErrors.add(key);
  context.errors.push(error);
};

const deduplicateAndSort = (
  errors: readonly WorkbookPlanValidationError[],
  maxErrors: number,
): readonly WorkbookPlanValidationError[] => {
  const seen = new Set<string>();
  const unique: WorkbookPlanValidationError[] = [];

  for (const error of errors) {
    const key = errorKey(error);
    if (!seen.has(key)) {
      seen.add(key);
      unique.push(error);
    }
  }

  return unique
    .sort(
      (left, right) =>
        compareCodeUnit(left.path, right.path) ||
        compareCodeUnit(left.code, right.code) ||
        compareCodeUnit(left.message, right.message) ||
        (left.limit ?? 0) - (right.limit ?? 0),
    )
    .slice(0, maxErrors);
};

const failure = (
  errors: readonly WorkbookPlanValidationError[],
  maxErrors: number,
): WorkbookPlanFailure => ({
  ok: false,
  errors: deduplicateAndSort(errors, maxErrors),
});

const resolveReference = (rootSchema: SchemaNode, reference: string): SchemaNode | undefined => {
  if (!reference.startsWith("#/")) return undefined;

  let current: unknown = rootSchema;
  for (const encodedSegment of reference.slice(2).split("/")) {
    if (!isSchemaNode(current)) return undefined;
    const segment = encodedSegment.replaceAll("~1", "/").replaceAll("~0", "~");
    if (!hasOwn(current, segment)) return undefined;
    current = current[segment];
  }
  return asSchemaNode(current);
};

const propertySchema = (schema: SchemaNode, property: string): SchemaNode | undefined =>
  asSchemaNode(schemaProperties(schema)[property]);

const schemaConst = (schema: SchemaNode): unknown =>
  hasOwn(schema, "const") ? schema["const"] : undefined;

const schemaEnum = (schema: SchemaNode): readonly unknown[] => {
  const values = schema["enum"];
  return Array.isArray(values) ? values : [];
};

const branchValuesForProperty = (
  branches: readonly SchemaNode[],
  property: string,
): readonly unknown[] => {
  const values: unknown[] = [];
  for (const branch of branches) {
    const branchProperty = propertySchema(branch, property);
    if (!branchProperty) continue;

    if (hasOwn(branchProperty, "const")) {
      values.push(schemaConst(branchProperty));
    } else {
      values.push(...schemaEnum(branchProperty));
    }
  }

  return values.filter(
    (value, index) => values.findIndex((candidate) => descriptorSafeDeepEqual(candidate, value)) === index,
  );
};

const findDiscriminator = (
  schema: SchemaNode,
  branches: readonly SchemaNode[],
): Discriminator | undefined => {
  for (const property of Object.keys(schemaProperties(schema))) {
    const values = branchValuesForProperty(branches, property);
    if (values.length > 1) return { property, values };
  }
  return undefined;
};

const branchMatchesDiscriminator = (
  branch: SchemaNode,
  property: string,
  value: unknown,
): boolean => {
  const branchProperty = propertySchema(branch, property);
  if (!branchProperty) return false;
  if (hasOwn(branchProperty, "const")) return descriptorSafeDeepEqual(schemaConst(branchProperty), value);
  return schemaEnum(branchProperty).some((candidate) => descriptorSafeDeepEqual(candidate, value));
};

const newBranchContext = (context: ValidationContext): ValidationContext => ({
  rootSchema: context.rootSchema,
  constraints: context.constraints,
  errors: [],
  seenErrors: new Set<string>(),
  budget: context.budget,
  failureCount: 0,
});

const validate = (
  value: unknown,
  schema: SchemaNode,
  path: string,
  context: ValidationContext,
  depth: number,
  options: ValidationOptions = {},
): boolean => {
  const failuresBefore = context.failureCount;

  if (context.budget.exhausted) return false;

  if (depth > context.constraints.maxDepth) {
    pushError(
      context,
      makeError(
        "bound_exceeded",
        path,
        "Validation depth exceeds the contract bound",
        context.constraints.maxDepth,
      ),
    );
    return false;
  }

  if (!noteNode(context)) return false;

  const reference = schema["$ref"];
  if (typeof reference === "string") {
    const resolved = resolveReference(context.rootSchema, reference);
    if (resolved === undefined) {
      pushError(context, makeError("oneof_mismatch", path, "Contract reference cannot be resolved"));
      return false;
    }
    return validate(value, resolved, path, context, depth + 1, {
      ...options,
      scalarValue: options.scalarValue === true || reference === "#/$defs/ScalarValue",
    });
  }

  const branches = options.ignoreOneOf === true ? [] : schemaNodes(schema["oneOf"]);
  if (branches.length > 0) {
    return validateOneOf(value, schema, branches, path, context, depth, options);
  }

  if (hasOwn(schema, "const") && !descriptorSafeDeepEqual(value, schemaConst(schema))) {
    pushError(context, makeError("type_mismatch", path, "Value does not satisfy the contract"));
    return false;
  }

  const enumValues = schemaEnum(schema);
  if (enumValues.length > 0 && !enumValues.some((candidate) => descriptorSafeDeepEqual(candidate, value))) {
    pushError(context, makeError("type_mismatch", path, "Value does not satisfy the contract"));
    return false;
  }

  const type = schema["type"];
  if (typeof type === "string") {
    switch (type) {
      case "object":
        return validateObject(value, schema, path, context, depth);
      case "array":
        return validateArray(value, schema, path, context, depth);
      case "string":
        return validateString(value, schema, path, context, options);
      case "number":
        return validateNumber(value, schema, path, context, false);
      case "integer":
        return validateNumber(value, schema, path, context, true);
      case "boolean":
        if (typeof value !== "boolean") {
          pushError(context, makeError("type_mismatch", path, "Expected boolean"));
          return false;
        }
        return true;
      case "null":
        if (value !== null) {
          pushError(context, makeError("type_mismatch", path, "Expected null"));
          return false;
        }
        return true;
      default:
        pushError(context, makeError("type_mismatch", path, "Unsupported contract type"));
        return false;
    }
  }

  return context.failureCount === failuresBefore;
};

const validateOneOf = (
  value: unknown,
  schema: SchemaNode,
  branches: readonly SchemaNode[],
  path: string,
  context: ValidationContext,
  depth: number,
  options: ValidationOptions,
): boolean => {
  const failuresBefore = context.failureCount;
  const discriminator = findDiscriminator(schema, branches);
  const objectValue = isPlainObject(value);
  const discriminatorProperty =
    discriminator !== undefined && objectValue
      ? ownDataDescriptor(value, discriminator.property)
      : undefined;
  const hasDiscriminatorValue = discriminatorProperty !== undefined;
  const discriminatorValue = discriminatorProperty?.value;

  if (
    discriminator !== undefined &&
    hasDiscriminatorValue &&
    !discriminator.values.some((candidate) =>
      descriptorSafeDeepEqual(candidate, discriminatorValue),
    )
  ) {
    pushError(
      context,
      makeError(
        "unknown_discriminator",
        joinPath(path, discriminator.property),
        "Unknown discriminator",
      ),
    );
    return false;
  }

  const candidateIndexes: number[] = [];
  for (let index = 0; index < branches.length; index += 1) {
    if (
      !hasDiscriminatorValue ||
      discriminator === undefined ||
      branchMatchesDiscriminator(branches[index], discriminator.property, discriminatorValue)
    ) {
      candidateIndexes.push(index);
    }
  }

  const branchResults = new Map<
    number,
    { readonly valid: boolean; readonly errors: readonly WorkbookPlanValidationError[] }
  >();
  for (const index of candidateIndexes) {
    if (!noteWork(context)) return false;
    const branchContext = newBranchContext(context);
    const valid = validate(value, branches[index], path, branchContext, depth + 1, options);
    branchResults.set(index, {
      valid: valid && branchContext.failureCount === 0,
      errors: branchContext.errors,
    });
    if (context.budget.exhausted) return false;
  }

  const matches = candidateIndexes.filter((index) => branchResults.get(index)?.valid === true);
  if (matches.length > 1) {
    pushError(context, makeError("oneof_mismatch", path, "Multiple contract branches matched"));
    return false;
  }

  if (matches.length === 0) {
    const selectedBranch = candidateIndexes[0] ?? 0;
    for (const error of branchResults.get(selectedBranch)?.errors ?? []) {
      pushError(context, error);
    }
    if (context.failureCount === failuresBefore) {
      pushError(context, makeError("oneof_mismatch", path, "No contract branch matched"));
    }
    return false;
  }

  if (!noteWork(context)) return false;
  const baseContext = newBranchContext(context);
  const baseValid = validate(value, schema, path, baseContext, depth + 1, {
    ...options,
    ignoreOneOf: true,
  });
  if (context.budget.exhausted) return false;
  if (baseValid && baseContext.failureCount === 0) return true;

  for (const error of baseContext.errors) pushError(context, error);
  if (context.failureCount === failuresBefore) {
    pushError(context, makeError("oneof_mismatch", path, "No contract branch matched"));
  }
  return false;
};

const validateObject = (
  value: unknown,
  schema: SchemaNode,
  path: string,
  context: ValidationContext,
  depth: number,
): boolean => {
  const failuresBefore = context.failureCount;
  if (!isPlainObject(value)) {
    pushError(context, makeError("type_mismatch", path, "Expected object"));
    return false;
  }

  const properties = schemaProperties(schema);
  const required = schemaStrings(schema["required"]);
  const knownValues = new Map<string, unknown>();
  const dataKeys = new Set<string>();
  const ownKeys = Reflect.ownKeys(value);
  if (!noteWork(context, ownKeys.length)) return false;

  for (const key of ownKeys.sort(compareOwnPropertyKeys)) {
    if (typeof key !== "string") {
      pushError(context, makeError("symbol_key_not_supported", path, "Symbol keys are not supported"));
      continue;
    }

    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (descriptor === undefined) continue;
    if (!Object.prototype.hasOwnProperty.call(descriptor, "value")) {
      pushError(
        context,
        makeError("accessor_property", joinPath(path, key), "Accessor properties are not supported"),
      );
      continue;
    }

    dataKeys.add(key);
    if (!hasOwn(properties, key)) {
      if (schema["additionalProperties"] === false) {
        pushError(context, makeError("unknown_property", joinPath(path, key), "Unknown property"));
      }
      continue;
    }

    knownValues.set(key, descriptor.value);
  }

  for (const alias of Object.keys(context.constraints.aliasMap).sort(compareCodeUnit)) {
    if (!noteWork(context)) return false;
    const canonical = context.constraints.aliasMap[alias];
    if (
      canonical !== undefined &&
      hasOwn(properties, canonical) &&
      dataKeys.has(alias) &&
      dataKeys.has(canonical)
    ) {
      pushError(
        context,
        makeError("alias_conflict", joinPath(path, alias), "Canonical and alias cannot both be present"),
      );
    }
  }

  for (const requiredProperty of required) {
    if (!noteWork(context)) return false;
    if (!knownValues.has(requiredProperty)) {
      pushError(
        context,
        makeError("missing_required", joinPath(path, requiredProperty), "Required field missing"),
      );
    }
  }

  for (const [key, childValue] of knownValues) {
    const childSchema = properties[key];
    if (childSchema !== undefined) {
      validate(childValue, childSchema, joinPath(path, key), context, depth + 1);
      if (context.budget.exhausted) return false;
    }
  }

  validateMetadata(value, path, context);
  return context.failureCount === failuresBefore;
};

const validateArray = (
  value: unknown,
  schema: SchemaNode,
  path: string,
  context: ValidationContext,
  depth: number,
): boolean => {
  const failuresBefore = context.failureCount;
  if (!Array.isArray(value)) {
    pushError(context, makeError("type_mismatch", path, "Expected array"));
    return false;
  }

  const length = arrayLength(value);
  if (length === undefined) {
    pushError(context, makeError("type_mismatch", path, "Expected array"));
    return false;
  }

  const itemValues = new Map<number, unknown>();
  const ownKeys = Reflect.ownKeys(value);
  if (!noteWork(context, ownKeys.length)) return false;
  for (const key of ownKeys.sort(compareOwnPropertyKeys)) {
    if (typeof key !== "string") {
      pushError(context, makeError("symbol_key_not_supported", path, "Symbol keys are not supported"));
      continue;
    }
    if (key === "length") continue;

    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (descriptor === undefined) continue;
    if (!Object.prototype.hasOwnProperty.call(descriptor, "value")) {
      pushError(
        context,
        makeError("accessor_property", joinPath(path, key), "Accessor properties are not supported"),
      );
      continue;
    }

    const index = canonicalArrayIndex(key, length);
    if (index === undefined) {
      pushError(context, makeError("unknown_property", joinPath(path, key), "Unknown property"));
      continue;
    }
    itemValues.set(index, descriptor.value);
  }

  const minItems = typeof schema["minItems"] === "number" ? schema["minItems"] : undefined;
  const maxItems = typeof schema["maxItems"] === "number" ? schema["maxItems"] : undefined;
  if (minItems !== undefined && length < minItems) {
    pushError(context, makeError("bound_exceeded", path, "Array is shorter than the contract bound", minItems));
    return false;
  }
  if (maxItems !== undefined && length > maxItems) {
    pushError(context, makeError("bound_exceeded", path, "Array exceeds the contract bound", maxItems));
    return false;
  }

  if (!noteWork(context, length)) return false;
  const itemSchema = schema.items;
  if (itemSchema !== undefined) {
    for (let index = 0; index < length; index += 1) {
      validate(itemValues.get(index), itemSchema, joinPath(path, index), context, depth + 1);
      if (context.budget.exhausted) return false;
    }
  }

  return context.failureCount === failuresBefore;
};

const validateString = (
  value: unknown,
  schema: SchemaNode,
  path: string,
  context: ValidationContext,
  options: ValidationOptions,
): boolean => {
  const failuresBefore = context.failureCount;
  if (typeof value !== "string") {
    pushError(context, makeError("type_mismatch", path, "Expected string"));
    return false;
  }

  const minLength = typeof schema["minLength"] === "number" ? schema["minLength"] : undefined;
  const maxLength = typeof schema["maxLength"] === "number" ? schema["maxLength"] : undefined;
  const countingLimit = maxLength ?? minLength;
  if (countingLimit === undefined) return true;

  const length = boundedCodePointLength(value, countingLimit);
  if (minLength !== undefined && length < minLength) {
    pushError(context, makeError("bound_exceeded", path, "String is shorter than the contract bound", minLength));
  } else if (maxLength !== undefined && length > maxLength) {
    pushError(context, makeError("bound_exceeded", path, "String exceeds the contract bound", maxLength));
  } else if (options.scalarValue === true && !isExcelCellTextWithinUtf16Limit(value)) {
    pushError(
      context,
      makeError(
        "bound_exceeded",
        path,
        "String exceeds the Excel cell text UTF-16 limit",
        EXCEL_CELL_TEXT_MAX_LENGTH,
      ),
    );
  }

  return context.failureCount === failuresBefore;
};

const validateNumber = (
  value: unknown,
  schema: SchemaNode,
  path: string,
  context: ValidationContext,
  integer: boolean,
): boolean => {
  const failuresBefore = context.failureCount;
  if (typeof value !== "number" || !Number.isFinite(value) || (integer && !Number.isInteger(value))) {
    pushError(
      context,
      makeError("type_mismatch", path, integer ? "Expected finite integer" : "Expected finite number"),
    );
    return false;
  }

  const minimum = typeof schema["minimum"] === "number" ? schema["minimum"] : undefined;
  const maximum = typeof schema["maximum"] === "number" ? schema["maximum"] : undefined;
  const exclusiveMinimum =
    typeof schema["exclusiveMinimum"] === "number" ? schema["exclusiveMinimum"] : undefined;
  const exclusiveMaximum =
    typeof schema["exclusiveMaximum"] === "number" ? schema["exclusiveMaximum"] : undefined;

  if (exclusiveMinimum !== undefined && value <= exclusiveMinimum) {
    pushError(context, makeError("bound_exceeded", path, "Number is below the contract bound", exclusiveMinimum));
  } else if (exclusiveMaximum !== undefined && value >= exclusiveMaximum) {
    pushError(context, makeError("bound_exceeded", path, "Number exceeds the contract bound", exclusiveMaximum));
  } else if (minimum !== undefined && value < minimum) {
    pushError(context, makeError("bound_exceeded", path, "Number is below the contract bound", minimum));
  } else if (maximum !== undefined && value > maximum) {
    pushError(context, makeError("bound_exceeded", path, "Number exceeds the contract bound", maximum));
  }

  return context.failureCount === failuresBefore;
};

const validateUtf8Value = (
  value: unknown,
  path: string,
  limitName: ValidationUtf8LimitName,
  context: ValidationContext,
): void => {
  if (typeof value === "string") {
    const limit = context.constraints.maxUtf8Bytes[limitName];
    if (exceedsUtf8ByteLimit(value, limit)) {
      pushError(
        context,
        makeError("utf8_limit_exceeded", path, "String exceeds the UTF-8 byte limit", limit),
      );
    }
    return;
  }

  if (!Array.isArray(value)) return;
  const length = arrayLength(value);
  if (length === undefined || !noteWork(context, length)) return;
  for (let index = 0; index < length; index += 1) {
    const descriptor = ownDataDescriptor(value, String(index));
    if (descriptor !== undefined) {
      validateUtf8Value(descriptor.value, joinPath(path, index), limitName, context);
    }
    if (context.budget.exhausted) return;
  }
};

const validateRectangularMatrix = (
  value: unknown,
  path: string,
  context: ValidationContext,
): void => {
  if (!Array.isArray(value)) return;
  const length = arrayLength(value);
  if (length === undefined || length < 2 || !noteWork(context, length)) return;

  const firstRowProperty = ownDataDescriptor(value, "0");
  const firstRow = firstRowProperty?.value;
  if (!Array.isArray(firstRow)) return;
  const width = arrayLength(firstRow);
  if (width === undefined || width === 0) return;

  for (let index = 1; index < length; index += 1) {
    const rowProperty = ownDataDescriptor(value, String(index));
    const row = rowProperty?.value;
    const rowWidth = Array.isArray(row) ? arrayLength(row) : undefined;
    if (rowWidth !== width) {
      pushError(
        context,
        makeError("matrix_jagged", joinPath(path, index), "Matrix rows must have uniform length"),
      );
    }
  }
};

const validateMetadata = (
  value: InputObject,
  path: string,
  context: ValidationContext,
): void => {
  const kindProperty = ownDataDescriptor(value, "kind");
  const discriminatorValue =
    kindProperty !== undefined && typeof kindProperty.value === "string"
      ? kindProperty.value
      : undefined;
  if (discriminatorValue === undefined) return;

  const utf8Fields = context.constraints.utf8ByteFields[discriminatorValue];
  if (utf8Fields !== undefined) {
    for (const field of Object.keys(utf8Fields).sort(compareCodeUnit)) {
      if (!noteWork(context)) return;
      const fieldProperty = ownDataDescriptor(value, field);
      if (fieldProperty !== undefined) {
        validateUtf8Value(fieldProperty.value, joinPath(path, field), utf8Fields[field], context);
      }
      if (context.budget.exhausted) return;
    }
  }

  const matrixFields = context.constraints.rectangularMatrices[discriminatorValue];
  if (matrixFields !== undefined) {
    for (const field of matrixFields) {
      if (!noteWork(context)) return;
      const fieldProperty = ownDataDescriptor(value, field);
      if (fieldProperty !== undefined) {
        validateRectangularMatrix(fieldProperty.value, joinPath(path, field), context);
      }
      if (context.budget.exhausted) return;
    }
  }

  const ruleTypeProperty = ownDataDescriptor(value, "ruleType");
  const ruleType =
    ruleTypeProperty !== undefined && typeof ruleTypeProperty.value === "string"
      ? ruleTypeProperty.value
      : undefined;
  const requiredFields =
    ruleType === undefined
      ? undefined
      : context.constraints.conditionalFormatRequires[ruleType];
  if (requiredFields !== undefined) {
    for (const field of requiredFields) {
      if (!noteWork(context)) return;
      if (ownDataDescriptor(value, field) === undefined) {
        pushError(
          context,
          makeError("missing_required", joinPath(path, field), "Required field missing"),
        );
      }
    }
  }
};

const rootSchema = (): SchemaNode => workbookPlanContract.schema;

const expectedVersion = (schema: SchemaNode): number => {
  const versionSchema = propertySchema(schema, "version");
  const value = versionSchema === undefined ? undefined : schemaConst(versionSchema);
  return typeof value === "number" && Number.isFinite(value) ? value : workbookPlanContract.version;
};

const preflightRoot = (
  input: InputObject,
  schema: SchemaNode,
  context: ValidationContext,
): void => {
  const properties = schemaProperties(schema);
  const ownKeys = Reflect.ownKeys(input);
  if (!noteWork(context, ownKeys.length)) return;

  for (const key of ownKeys.sort(compareOwnPropertyKeys)) {
    if (typeof key !== "string") {
      pushError(context, makeError("symbol_key_not_supported", "", "Symbol keys are not supported"));
      continue;
    }

    const descriptor = Object.getOwnPropertyDescriptor(input, key);
    if (descriptor === undefined) continue;
    if (!Object.prototype.hasOwnProperty.call(descriptor, "value")) {
      pushError(
        context,
        makeError("accessor_property", joinPath("", key), "Accessor properties are not supported"),
      );
      continue;
    }

    if (schema["additionalProperties"] === false && !hasOwn(properties, key)) {
      pushError(context, makeError("unknown_property", joinPath("", key), "Unknown property"));
    }
  }
};

const isWorkbookPlan = (
  input: unknown,
  schema: SchemaNode,
  context: ValidationContext,
): input is WorkbookPlan =>
  validate(input, schema, "", context, 0) &&
  context.failureCount === 0 &&
  !context.budget.exhausted;

const workLimitFailure = (): WorkbookPlanFailure => ({
  ok: false,
  errors: [makeError("work_limit_exceeded", "", "Validation work limit exceeded")],
});

export function validateWorkbookPlan(input: unknown): ValidateWorkbookPlanResult {
  const schema = rootSchema();
  const constraints = workbookPlanContract.validationConstraints;

  if (!isPlainObject(input)) {
    return failure([makeError("type_mismatch", "", "Expected object at root")], constraints.maxErrors);
  }

  const context: ValidationContext = {
    rootSchema: schema,
    constraints,
    errors: [],
    seenErrors: new Set<string>(),
    budget: { workUnits: 0, visitedNodes: 0, exhausted: false },
    failureCount: 0,
  };

  preflightRoot(input, schema, context);
  if (context.budget.exhausted) return workLimitFailure();
  if (context.failureCount > 0) return failure(context.errors, constraints.maxErrors);

  const versionProperty = ownDataDescriptor(input, "version");
  const version = versionProperty?.value;
  const expected = expectedVersion(schema);
  if (
    typeof version !== "number" ||
    !Number.isFinite(version) ||
    !Number.isInteger(version) ||
    version !== expected
  ) {
    return failure(
      [makeError("unsupported_version", "/version", "Unsupported plan version", expected)],
      constraints.maxErrors,
    );
  }

  if (!isWorkbookPlan(input, schema, context)) {
    if (context.budget.exhausted) return workLimitFailure();
    return failure(context.errors, constraints.maxErrors);
  }

  return { ok: true, value: input };
}
