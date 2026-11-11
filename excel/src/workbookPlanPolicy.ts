import { workbookPlanContract } from "./workbookPlanContract";

type PlainRecord = Readonly<Record<string, unknown>>;

export type WorkbookRequestClassification =
  | "read_only"
  | "explicit_mutation"
  | "ambiguous_mutation"
  | "explicit_structure"
  | "prohibited";

export type WorkbookPlanOperationEffect =
  (typeof workbookPlanContract.operationDefinitions)[number]["effect"];

export type WorkbookPlanCapabilityName =
  (typeof workbookPlanContract.operationDefinitions)[number]["capabilityRequirements"][number]["capability"];

export type WorkbookPlanSelectionRelationship =
  | "not_applicable"
  | "inside"
  | "outside"
  | "unknown";

export interface WorkbookPlanResourceSummary {
  readonly operationCount: number;
  readonly cellsRead: number;
  readonly cellsWritten: number;
  readonly formulaUtf8Bytes: number;
  readonly csvPayloadUtf8Bytes: number;
  readonly commentUtf8Bytes: number;
  readonly sheetsCreated: number;
  readonly chartsCreated: number;
  readonly tablesCreated: number;
  readonly verificationReads: number;
}

export type WorkbookPlanResourceLimits = WorkbookPlanResourceSummary;

export interface WorkbookPlanOperationTargetSummary {
  readonly operationIndex: number;
  readonly sheetNames: readonly string[];
  readonly selectionRelationship: WorkbookPlanSelectionRelationship;
}

export interface WorkbookPlanOperationSummary {
  readonly operationIndex: number;
  readonly effect: WorkbookPlanOperationEffect;
  readonly target: WorkbookPlanOperationTargetSummary;
}

export interface WorkbookPlanPolicyInput {
  readonly requestClassification: WorkbookRequestClassification;
  readonly activeSheetName: string;
  readonly selectionScopedRequest: boolean;
  readonly operationSummaries: readonly WorkbookPlanOperationSummary[];
  readonly capabilitySummaries: readonly WorkbookPlanCapabilityName[];
  readonly resourceUsage: WorkbookPlanResourceSummary;
  readonly resourceLimits: WorkbookPlanResourceLimits;
  readonly confirmation?: Readonly<{
    readonly confirmed: boolean;
    /** Present only for a pane-owned user confirmation event. */
    readonly eventId?: string;
  }>;
}

export type WorkbookPlanPolicyDecisionKind = "approved" | "requires_confirmation" | "rejected";

export type WorkbookPlanPolicyReasonCode =
  | "approved"
  | "invalid_trusted_input"
  | "unsupported_effect"
  | "prohibited_request_classification"
  | "missing_contract_capability"
  | "resource_limit_exceeded"
  | "empty_plan"
  | "inconsistent_operation_counts"
  | "ambiguous_mutation_intent"
  | "structural_intent_required"
  | "active_sheet_mismatch"
  | "selection_scope_mismatch"
  | "confirmation_required";

export interface WorkbookPlanPolicyDecision {
  readonly decision: WorkbookPlanPolicyDecisionKind;
  readonly reasonCode: WorkbookPlanPolicyReasonCode;
}

type RejectedReasonCode = Exclude<
  WorkbookPlanPolicyReasonCode,
  | "approved"
  | "ambiguous_mutation_intent"
  | "structural_intent_required"
  | "active_sheet_mismatch"
  | "selection_scope_mismatch"
  | "confirmation_required"
>;

interface ParseSuccess<T> {
  readonly ok: true;
  readonly value: T;
}

interface ParseFailure {
  readonly ok: false;
  readonly reasonCode: RejectedReasonCode;
}

type ParseOutcome<T> = ParseSuccess<T> | ParseFailure;

const REQUEST_CLASSIFICATIONS: readonly WorkbookRequestClassification[] = [
  "read_only",
  "explicit_mutation",
  "ambiguous_mutation",
  "explicit_structure",
  "prohibited",
];

const SELECTION_RELATIONSHIPS: readonly WorkbookPlanSelectionRelationship[] = [
  "not_applicable",
  "inside",
  "outside",
  "unknown",
];

const RESOURCE_KEYS = [
  "operationCount",
  "cellsRead",
  "cellsWritten",
  "formulaUtf8Bytes",
  "csvPayloadUtf8Bytes",
  "commentUtf8Bytes",
  "sheetsCreated",
  "chartsCreated",
  "tablesCreated",
  "verificationReads",
] as const;

type ResourceKey = (typeof RESOURCE_KEYS)[number];

const buildUniqueContractEffects = (): readonly WorkbookPlanOperationEffect[] => {
  const effects: WorkbookPlanOperationEffect[] = [];
  for (const definition of workbookPlanContract.operationDefinitions) {
    if (!effects.includes(definition.effect)) {
      effects.push(definition.effect);
    }
  }
  return effects;
};

const buildUniqueContractCapabilityNames = (): readonly WorkbookPlanCapabilityName[] => {
  const names: WorkbookPlanCapabilityName[] = [];
  for (const definition of workbookPlanContract.operationDefinitions) {
    for (const requirement of definition.capabilityRequirements) {
      if (!names.includes(requirement.capability)) {
        names.push(requirement.capability);
      }
    }
  }
  return names;
};

const CONTRACT_OPERATION_EFFECTS = Object.freeze(buildUniqueContractEffects());
const CONTRACT_CAPABILITY_NAMES = Object.freeze(buildUniqueContractCapabilityNames());

export const WORKBOOK_PLAN_OPERATION_EFFECTS = CONTRACT_OPERATION_EFFECTS;
export const WORKBOOK_PLAN_CAPABILITY_NAMES = CONTRACT_CAPABILITY_NAMES;

const TRUSTED_POLICY_INPUT_SNAPSHOTS = new WeakSet<object>();

const hasOwn = (value: object, key: string): boolean =>
  Object.prototype.hasOwnProperty.call(value, key);

const isSafeNonNegativeInteger = (value: unknown): value is number =>
  typeof value === "number" && Number.isSafeInteger(value) && value >= 0;

const isKnownString = <T extends string>(value: unknown, allowed: readonly T[]): value is T =>
  typeof value === "string" && allowed.includes(value as T);

const isExactPlainObject = (value: unknown): value is PlainRecord => {
  if (value === null || typeof value !== "object" || Array.isArray(value)) return false;
  try {
    return Object.getPrototypeOf(value) === Object.prototype;
  } catch {
    return false;
  }
};

const snapshotExactObject = (
  value: unknown,
  requiredKeys: readonly string[],
  optionalKeys: readonly string[] = [],
): PlainRecord | undefined => {
  if (!isExactPlainObject(value)) return undefined;

  try {
    if (Object.getOwnPropertySymbols(value).length !== 0) return undefined;

    const ownNames = Object.getOwnPropertyNames(value);
    const allowedKeys = [...requiredKeys, ...optionalKeys];

    if (ownNames.length < requiredKeys.length || ownNames.length > allowedKeys.length) {
      return undefined;
    }

    for (const name of ownNames) {
      let allowed = false;
      for (const key of allowedKeys) {
        if (key === name) {
          allowed = true;
          break;
        }
      }

      if (!allowed) return undefined;

      const descriptor = Object.getOwnPropertyDescriptor(value, name);
      if (descriptor === undefined || !hasOwn(descriptor, "value")) return undefined;
    }

    for (const key of requiredKeys) {
      let present = false;
      for (const name of ownNames) {
        if (name === key) {
          present = true;
          break;
        }
      }

      if (!present) return undefined;
    }

    const snapshot: Record<string, unknown> = {};
    for (const name of ownNames) {
      const descriptor = Object.getOwnPropertyDescriptor(value, name);
      if (descriptor === undefined || !hasOwn(descriptor, "value")) return undefined;
      snapshot[name] = descriptor.value;
    }

    return snapshot;
  } catch {
    return undefined;
  }
};

const snapshotDenseArray = (value: unknown): readonly unknown[] | undefined => {
  if (!Array.isArray(value)) return undefined;

  try {
    if (Object.getPrototypeOf(value) !== Array.prototype) return undefined;

    const lengthDescriptor = Object.getOwnPropertyDescriptor(value, "length");
    if (lengthDescriptor === undefined || !hasOwn(lengthDescriptor, "value")) return undefined;
    const length = lengthDescriptor.value;
    if (!isSafeNonNegativeInteger(length)) return undefined;

    if (Object.getOwnPropertySymbols(value).length !== 0) return undefined;

    const ownNames = Object.getOwnPropertyNames(value);
    if (ownNames.length !== length + 1) return undefined;

    const snapshot = new Array<unknown>(length);
    for (let index = 0; index < length; index += 1) {
      const key = String(index);
      const descriptor = Object.getOwnPropertyDescriptor(value, key);
      if (descriptor === undefined || !hasOwn(descriptor, "value")) return undefined;
      snapshot[index] = descriptor.value;
    }

    return snapshot;
  } catch {
    return undefined;
  }
};

const materializeSnapshotValue = (
  value: unknown,
  seen: Map<object, unknown> = new Map<object, unknown>(),
): unknown => {
  if (value === null || (typeof value !== "object" && typeof value !== "function")) {
    return value;
  }

  if (typeof value === "function") return value;

  const existing = seen.get(value);
  if (existing !== undefined) return existing;

  let isArray = false;
  try {
    isArray = Array.isArray(value);
  } catch {
    isArray = false;
  }

  let clone: Record<string, unknown> | unknown[];
  if (isArray) {
    let length = 0;
    try {
      const lengthDescriptor = Object.getOwnPropertyDescriptor(value, "length");
      if (lengthDescriptor !== undefined && hasOwn(lengthDescriptor, "value")) {
        length = isSafeNonNegativeInteger(lengthDescriptor.value) ? lengthDescriptor.value : 0;
      }
    } catch {
      length = 0;
    }
    clone = new Array<unknown>(length);
  } else {
    clone = {};
  }
  seen.set(value, clone);

  let ownNames: string[] = [];
  let ownSymbols: symbol[] = [];
  try {
    ownNames = Object.getOwnPropertyNames(value);
    ownSymbols = Object.getOwnPropertySymbols(value);
  } catch {
    ownNames = [];
    ownSymbols = [];
  }

  for (const name of ownNames) {
    if (isArray && name === "length") continue;

    let descriptor: PropertyDescriptor | undefined;
    try {
      descriptor = Object.getOwnPropertyDescriptor(value, name);
    } catch {
      descriptor = undefined;
    }
    if (descriptor === undefined || !hasOwn(descriptor, "value")) continue;

    try {
      Object.defineProperty(clone, name, {
        value: materializeSnapshotValue(descriptor.value, seen),
        enumerable: descriptor.enumerable,
        configurable: true,
        writable: true,
      });
    } catch {
      // A malformed caller-owned descriptor is represented by omission and rejected later.
    }
  }

  for (const symbol of ownSymbols) {
    let descriptor: PropertyDescriptor | undefined;
    try {
      descriptor = Object.getOwnPropertyDescriptor(value, symbol);
    } catch {
      descriptor = undefined;
    }
    if (descriptor === undefined || !hasOwn(descriptor, "value")) continue;

    try {
      Object.defineProperty(clone, symbol, {
        value: materializeSnapshotValue(descriptor.value, seen),
        enumerable: descriptor.enumerable,
        configurable: true,
        writable: true,
      });
    } catch {
      // A malformed caller-owned descriptor is represented by omission and rejected later.
    }
  }

  return Object.freeze(clone);
};

export function createWorkbookPlanPolicyInputSnapshot(
  input: WorkbookPlanPolicyInput,
): WorkbookPlanPolicyInput {
  const snapshot = materializeSnapshotValue(input);
  if (snapshot === null || typeof snapshot !== "object") {
    const invalidSnapshot = Object.freeze({});
    TRUSTED_POLICY_INPUT_SNAPSHOTS.add(invalidSnapshot);
    return invalidSnapshot as WorkbookPlanPolicyInput;
  }

  TRUSTED_POLICY_INPUT_SNAPSHOTS.add(snapshot);
  return snapshot as WorkbookPlanPolicyInput;
};

const rejectedDecision = (
  reasonCode: Exclude<WorkbookPlanPolicyReasonCode, "approved" | "ambiguous_mutation_intent" | "structural_intent_required" | "active_sheet_mismatch" | "selection_scope_mismatch" | "confirmation_required">,
): WorkbookPlanPolicyDecision => ({
  decision: "rejected",
  reasonCode,
});

const requiresConfirmationDecision = (
  reasonCode:
    | "ambiguous_mutation_intent"
    | "structural_intent_required"
    | "active_sheet_mismatch"
    | "selection_scope_mismatch"
    | "confirmation_required",
): WorkbookPlanPolicyDecision => ({
  decision: "requires_confirmation",
  reasonCode,
});

const approvedDecision = (): WorkbookPlanPolicyDecision => ({
  decision: "approved",
  reasonCode: "approved",
});

const confirmationOrApproval = (
  reasonCode:
    | "ambiguous_mutation_intent"
    | "structural_intent_required"
    | "active_sheet_mismatch"
    | "selection_scope_mismatch"
    | "confirmation_required",
  confirmation: Readonly<{ readonly confirmed: boolean }> | undefined,
): WorkbookPlanPolicyDecision => {
  if (confirmation?.confirmed === true) {
    return approvedDecision();
  }

  return requiresConfirmationDecision(reasonCode);
};

const parseRequestClassification = (
  value: unknown,
): WorkbookRequestClassification | undefined => {
  if (!isKnownString(value, REQUEST_CLASSIFICATIONS)) return undefined;
  return value;
};

const parseSelectionRelationship = (
  value: unknown,
): WorkbookPlanSelectionRelationship | undefined => {
  if (!isKnownString(value, SELECTION_RELATIONSHIPS)) return undefined;
  return value;
};

const parseResourceSummary = (
  value: unknown,
): WorkbookPlanResourceSummary | undefined => {
  const snapshot = snapshotExactObject(value, RESOURCE_KEYS, []);
  if (snapshot === undefined) return undefined;

  const summary: Record<ResourceKey, number> = {
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
  };

  for (const key of RESOURCE_KEYS) {
    const metric = snapshot[key];
    if (!isSafeNonNegativeInteger(metric)) return undefined;
    summary[key] = metric;
  }

  return summary;
};

const parseConfirmation = (
  value: unknown,
): Readonly<{ readonly confirmed: boolean; readonly eventId?: string }> | undefined => {
  const snapshot = snapshotExactObject(value, ["confirmed"], ["eventId"]);
  if (snapshot === undefined) return undefined;

  const confirmed = snapshot.confirmed;
  if (typeof confirmed !== "boolean") return undefined;
  if (hasOwn(snapshot, "eventId") && (typeof snapshot.eventId !== "string" || snapshot.eventId.length === 0)) {
    return undefined;
  }

  return hasOwn(snapshot, "eventId")
    ? { confirmed, eventId: snapshot.eventId as string }
    : { confirmed };
};

const parseSheetNames = (value: unknown): readonly string[] | undefined => {
  const snapshot = snapshotDenseArray(value);
  if (snapshot === undefined) return undefined;

  const sheetNames: string[] = [];
  for (const item of snapshot) {
    if (typeof item !== "string" || item.length === 0) return undefined;
    for (const existing of sheetNames) {
      if (existing === item) return undefined;
    }
    sheetNames.push(item);
  }

  return sheetNames;
};

const parseCapabilitySummaries = (
  value: unknown,
): ParseOutcome<readonly WorkbookPlanCapabilityName[]> => {
  const snapshot = snapshotDenseArray(value);
  if (snapshot === undefined) {
    return { ok: false, reasonCode: "invalid_trusted_input" };
  }

  const parsed: WorkbookPlanCapabilityName[] = [];
  const seen: Record<string, boolean> = Object.create(null);

  for (const item of snapshot) {
    if (!isKnownString(item, WORKBOOK_PLAN_CAPABILITY_NAMES)) {
      return { ok: false, reasonCode: "invalid_trusted_input" };
    }

    if (seen[item] === true) {
      return { ok: false, reasonCode: "invalid_trusted_input" };
    }

    seen[item] = true;
    parsed.push(item);
  }

  for (const requiredName of WORKBOOK_PLAN_CAPABILITY_NAMES) {
    if (seen[requiredName] !== true) {
      if (parsed.length < WORKBOOK_PLAN_CAPABILITY_NAMES.length) {
        return { ok: false, reasonCode: "missing_contract_capability" };
      }
      return { ok: false, reasonCode: "invalid_trusted_input" };
    }
  }

  return { ok: true, value: parsed };
};

const parseOperationSummaries = (
  value: unknown,
): ParseOutcome<readonly WorkbookPlanOperationSummary[]> => {
  const snapshot = snapshotDenseArray(value);
  if (snapshot === undefined) {
    return { ok: false, reasonCode: "invalid_trusted_input" };
  }

  const parsed: WorkbookPlanOperationSummary[] = [];

  for (let index = 0; index < snapshot.length; index += 1) {
    const item = snapshot[index];
    const operation = snapshotExactObject(item, ["operationIndex", "effect", "target"], []);
    if (operation === undefined) {
      return { ok: false, reasonCode: "invalid_trusted_input" };
    }

    const operationIndex = operation.operationIndex;
    if (!isSafeNonNegativeInteger(operationIndex) || operationIndex !== index) {
      return { ok: false, reasonCode: "inconsistent_operation_counts" };
    }

    const effect = operation.effect;
    if (typeof effect !== "string") {
      return { ok: false, reasonCode: "invalid_trusted_input" };
    }
    if (!isKnownString(effect, WORKBOOK_PLAN_OPERATION_EFFECTS)) {
      return { ok: false, reasonCode: "unsupported_effect" };
    }

    const target = snapshotExactObject(operation.target, ["operationIndex", "sheetNames", "selectionRelationship"], []);
    if (target === undefined) {
      return { ok: false, reasonCode: "invalid_trusted_input" };
    }

    const targetOperationIndex = target.operationIndex;
    if (!isSafeNonNegativeInteger(targetOperationIndex) || targetOperationIndex !== index) {
      return { ok: false, reasonCode: "inconsistent_operation_counts" };
    }

    const sheetNames = parseSheetNames(target.sheetNames);
    if (sheetNames === undefined) {
      return { ok: false, reasonCode: "invalid_trusted_input" };
    }

    const selectionRelationship = parseSelectionRelationship(target.selectionRelationship);
    if (selectionRelationship === undefined) {
      return { ok: false, reasonCode: "invalid_trusted_input" };
    }

    if (effect === "calculate") {
      if (sheetNames.length !== 0) {
        return { ok: false, reasonCode: "invalid_trusted_input" };
      }
    } else if (sheetNames.length === 0) {
      return { ok: false, reasonCode: "invalid_trusted_input" };
    }

    parsed.push({
      operationIndex,
      effect,
      target: {
        operationIndex: targetOperationIndex,
        sheetNames,
        selectionRelationship,
      },
    });
  }

  return { ok: true, value: parsed };
};

const parseWorkbookPlanPolicyInput = (
  input: unknown,
): ParseOutcome<WorkbookPlanPolicyInput> => {
  const snapshot = snapshotExactObject(
    input,
    [
      "requestClassification",
      "activeSheetName",
      "selectionScopedRequest",
      "operationSummaries",
      "capabilitySummaries",
      "resourceUsage",
      "resourceLimits",
    ],
    ["confirmation"],
  );

  if (snapshot === undefined) {
    return { ok: false, reasonCode: "invalid_trusted_input" };
  }

  const requestClassification = parseRequestClassification(snapshot.requestClassification);
  if (requestClassification === undefined) {
    return { ok: false, reasonCode: "invalid_trusted_input" };
  }

  const activeSheetName = snapshot.activeSheetName;
  if (typeof activeSheetName !== "string" || activeSheetName.length === 0) {
    return { ok: false, reasonCode: "invalid_trusted_input" };
  }

  const selectionScopedRequest = snapshot.selectionScopedRequest;
  if (typeof selectionScopedRequest !== "boolean") {
    return { ok: false, reasonCode: "invalid_trusted_input" };
  }

  const resourceUsage = parseResourceSummary(snapshot.resourceUsage);
  if (resourceUsage === undefined) {
    return { ok: false, reasonCode: "invalid_trusted_input" };
  }

  const resourceLimits = parseResourceSummary(snapshot.resourceLimits);
  if (resourceLimits === undefined) {
    return { ok: false, reasonCode: "invalid_trusted_input" };
  }

  const capabilities = parseCapabilitySummaries(snapshot.capabilitySummaries);
  if (capabilities.ok === false) {
    return { ok: false, reasonCode: capabilities.reasonCode };
  }

  const operations = parseOperationSummaries(snapshot.operationSummaries);
  if (operations.ok === false) {
    return { ok: false, reasonCode: operations.reasonCode };
  }

  const confirmation = hasOwn(snapshot, "confirmation") ? parseConfirmation(snapshot.confirmation) : undefined;
  if (hasOwn(snapshot, "confirmation") && confirmation === undefined) {
    return { ok: false, reasonCode: "invalid_trusted_input" };
  }

  return {
    ok: true,
    value: {
      requestClassification,
      activeSheetName,
      selectionScopedRequest,
      operationSummaries: operations.value,
      capabilitySummaries: capabilities.value,
      resourceUsage,
      resourceLimits,
      confirmation,
    },
  };
};

const exceedsResourceLimits = (
  usage: WorkbookPlanResourceSummary,
  limits: WorkbookPlanResourceLimits,
): boolean => {
  for (const key of RESOURCE_KEYS) {
    if (usage[key] > limits[key]) {
      return true;
    }
  }

  return false;
};

const hasExactContractCapabilities = (
  capabilitySummaries: readonly WorkbookPlanCapabilityName[],
): boolean => {
  if (capabilitySummaries.length !== WORKBOOK_PLAN_CAPABILITY_NAMES.length) {
    return false;
  }

  const seen: Record<string, boolean> = Object.create(null);
  for (const capability of capabilitySummaries) {
    if (!isKnownString(capability, WORKBOOK_PLAN_CAPABILITY_NAMES)) {
      return false;
    }
    if (seen[capability] === true) {
      return false;
    }
    seen[capability] = true;
  }

  for (const expected of WORKBOOK_PLAN_CAPABILITY_NAMES) {
    if (seen[expected] !== true) {
      return false;
    }
  }

  return true;
};

const hasAnyOffActiveTarget = (
  activeSheetName: string,
  operationSummaries: readonly WorkbookPlanOperationSummary[],
): boolean => {
  for (const operation of operationSummaries) {
    for (const sheetName of operation.target.sheetNames) {
      if (sheetName !== activeSheetName) {
        return true;
      }
    }
  }

  return false;
};

const hasAnySelectionScopedOutsideOrUnknownTarget = (
  operationSummaries: readonly WorkbookPlanOperationSummary[],
): boolean => {
  for (const operation of operationSummaries) {
    if (
      operation.target.selectionRelationship === "outside" ||
      operation.target.selectionRelationship === "unknown"
    ) {
      return true;
    }
  }

  return false;
};

const hasAnyNonCalculateEffect = (
  operationSummaries: readonly WorkbookPlanOperationSummary[],
): boolean => {
  for (const operation of operationSummaries) {
    if (operation.effect !== "calculate") {
      return true;
    }
  }

  return false;
};

const decideRejected = (
  reasonCode:
    | "invalid_trusted_input"
    | "unsupported_effect"
    | "prohibited_request_classification"
    | "missing_contract_capability"
    | "resource_limit_exceeded"
    | "empty_plan"
    | "inconsistent_operation_counts",
): WorkbookPlanPolicyDecision => rejectedDecision(reasonCode);

export function decideWorkbookPlanPolicy(input: unknown): WorkbookPlanPolicyDecision {
  try {
    if (
      input === null ||
      (typeof input !== "object" && typeof input !== "function") ||
      !TRUSTED_POLICY_INPUT_SNAPSHOTS.has(input)
    ) {
      return decideRejected("invalid_trusted_input");
    }

    const parsed = parseWorkbookPlanPolicyInput(input);
    if (parsed.ok === false) {
      return rejectedDecision(parsed.reasonCode);
    }

    const policyInput = parsed.value;

    if (policyInput.requestClassification === "prohibited") {
      return decideRejected("prohibited_request_classification");
    }

    if (
      policyInput.operationSummaries.length !== policyInput.resourceUsage.operationCount ||
      policyInput.resourceUsage.operationCount !== policyInput.operationSummaries.length
    ) {
      return decideRejected("inconsistent_operation_counts");
    }

    if (!hasExactContractCapabilities(policyInput.capabilitySummaries)) {
      return decideRejected("missing_contract_capability");
    }

    if (exceedsResourceLimits(policyInput.resourceUsage, policyInput.resourceLimits)) {
      return decideRejected("resource_limit_exceeded");
    }

    if (policyInput.resourceUsage.operationCount === 0) {
      return decideRejected("empty_plan");
    }

    if (policyInput.requestClassification === "ambiguous_mutation") {
      return confirmationOrApproval("ambiguous_mutation_intent", policyInput.confirmation);
    }

    const hasStructureEffect = policyInput.operationSummaries.some(
      (operation) => operation.effect === "structure",
    );
    if (hasStructureEffect && policyInput.requestClassification !== "explicit_structure") {
      return confirmationOrApproval("structural_intent_required", policyInput.confirmation);
    }

    if (hasAnyOffActiveTarget(policyInput.activeSheetName, policyInput.operationSummaries)) {
      return confirmationOrApproval("active_sheet_mismatch", policyInput.confirmation);
    }

    if (
      policyInput.selectionScopedRequest &&
      hasAnySelectionScopedOutsideOrUnknownTarget(policyInput.operationSummaries)
    ) {
      return confirmationOrApproval("selection_scope_mismatch", policyInput.confirmation);
    }

    if (policyInput.requestClassification === "read_only" && hasAnyNonCalculateEffect(policyInput.operationSummaries)) {
      return confirmationOrApproval("confirmation_required", policyInput.confirmation);
    }

    return approvedDecision();
  } catch {
    return decideRejected("invalid_trusted_input");
  }
}

/**
 * R4c: intent rules of decideWorkbookPlanPolicy (same order, same reason
 * codes) for one unconfirmed direct-route primitive operation. The
 * active-sheet rule is omitted (OQ-C5); a selection-scoped request is never
 * matched by a primitive target.
 */
export function decidePrimitiveIntent(input: {
  readonly requestClassification: WorkbookRequestClassification;
  readonly selectionScopedRequest: boolean;
  readonly effect: WorkbookPlanOperationEffect;
}): WorkbookPlanPolicyDecision {
  if (!REQUEST_CLASSIFICATIONS.includes(input.requestClassification)) {
    return decideRejected("invalid_trusted_input");
  }
  if (!WORKBOOK_PLAN_OPERATION_EFFECTS.includes(input.effect)) {
    return decideRejected("unsupported_effect");
  }
  if (input.requestClassification === "prohibited") {
    return decideRejected("prohibited_request_classification");
  }
  if (input.requestClassification === "ambiguous_mutation") {
    return confirmationOrApproval("ambiguous_mutation_intent", undefined);
  }
  if (input.effect === "structure" && input.requestClassification !== "explicit_structure") {
    return confirmationOrApproval("structural_intent_required", undefined);
  }
  if (input.selectionScopedRequest) {
    return confirmationOrApproval("selection_scope_mismatch", undefined);
  }
  if (input.requestClassification === "read_only" && input.effect !== "calculate") {
    return confirmationOrApproval("confirmation_required", undefined);
  }
  return approvedDecision();
}
