import { describe, expect, it } from "vitest";
import {
  createWorkbookPlanPolicyInputSnapshot,
  decideWorkbookPlanPolicy,
  WORKBOOK_PLAN_CAPABILITY_NAMES,
  WORKBOOK_PLAN_OPERATION_EFFECTS,
  type WorkbookPlanCapabilityName,
  type WorkbookPlanOperationEffect,
  type WorkbookPlanOperationSummary,
  type WorkbookPlanOperationTargetSummary,
  type WorkbookPlanPolicyInput,
  type WorkbookRequestClassification,
  type WorkbookPlanResourceSummary,
  type WorkbookPlanSelectionRelationship,
} from "../workbookPlanPolicy";

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

const BASE_RESOURCE: WorkbookPlanResourceSummary = {
  operationCount: 1,
  cellsRead: 2,
  cellsWritten: 3,
  formulaUtf8Bytes: 4,
  csvPayloadUtf8Bytes: 5,
  commentUtf8Bytes: 6,
  sheetsCreated: 7,
  chartsCreated: 8,
  tablesCreated: 9,
  verificationReads: 10,
};

const createResourceSummary = (
  overrides: Partial<Record<ResourceKey, number>> = {},
  operationCount = 1,
): WorkbookPlanResourceSummary => ({
  ...BASE_RESOURCE,
  operationCount,
  ...overrides,
});

const createTargetSummary = (
  operationIndex: number,
  sheetNames: readonly string[],
  selectionRelationship: WorkbookPlanSelectionRelationship = "not_applicable",
): WorkbookPlanOperationTargetSummary => ({
  operationIndex,
  sheetNames,
  selectionRelationship,
});

const createOperationSummary = (
  operationIndex: number,
  effect: WorkbookPlanOperationEffect,
  sheetNames: readonly string[],
  selectionRelationship: WorkbookPlanSelectionRelationship = "not_applicable",
): WorkbookPlanOperationSummary => ({
  operationIndex,
  effect,
  target: createTargetSummary(operationIndex, sheetNames, selectionRelationship),
});

interface MakeInputOptions {
  readonly requestClassification?: WorkbookRequestClassification;
  readonly activeSheetName?: string;
  readonly selectionScopedRequest?: boolean;
  readonly operationSummaries?: readonly WorkbookPlanOperationSummary[];
  readonly capabilitySummaries?: readonly WorkbookPlanCapabilityName[];
  readonly resourceUsage?: Partial<Record<ResourceKey, number>>;
  readonly resourceLimits?: Partial<Record<ResourceKey, number>>;
  readonly confirmation?: Readonly<{ readonly confirmed: boolean }>;
}

const hasOwn = (value: object, key: string): boolean =>
  Object.prototype.hasOwnProperty.call(value, key);

const makeInput = (options: MakeInputOptions = {}): WorkbookPlanPolicyInput => {
  const operationSummaries = hasOwn(options, "operationSummaries")
    ? options.operationSummaries ?? [createOperationSummary(0, "calculate", [], "not_applicable")]
    : [createOperationSummary(0, "calculate", [], "not_applicable")];
  const capabilitySummaries = hasOwn(options, "capabilitySummaries")
    ? options.capabilitySummaries ?? [...WORKBOOK_PLAN_CAPABILITY_NAMES]
    : [...WORKBOOK_PLAN_CAPABILITY_NAMES];
  const requestClassification = hasOwn(options, "requestClassification")
    ? options.requestClassification ?? "explicit_mutation"
    : "explicit_mutation";
  const activeSheetName = hasOwn(options, "activeSheetName") ? options.activeSheetName ?? "Sheet1" : "Sheet1";
  const selectionScopedRequest = hasOwn(options, "selectionScopedRequest")
    ? options.selectionScopedRequest ?? false
    : false;
  const resourceUsageOverrides = hasOwn(options, "resourceUsage") ? options.resourceUsage ?? {} : {};
  const resourceLimitsOverrides = hasOwn(options, "resourceLimits") ? options.resourceLimits ?? {} : {};

  return {
    requestClassification,
    activeSheetName,
    selectionScopedRequest,
    operationSummaries,
    capabilitySummaries,
    resourceUsage: createResourceSummary(resourceUsageOverrides, operationSummaries.length),
    resourceLimits: createResourceSummary(resourceLimitsOverrides, operationSummaries.length),
    ...(hasOwn(options, "confirmation") ? { confirmation: options.confirmation } : {}),
  };
};

const expectDecision = (
  input: WorkbookPlanPolicyInput | unknown,
  decision: "approved" | "requires_confirmation" | "rejected",
  reasonCode:
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
    | "confirmation_required",
): void => {
  expect(decideWorkbookPlanPolicy(createWorkbookPlanPolicyInputSnapshot(input as WorkbookPlanPolicyInput))).toEqual({
    decision,
    reasonCode,
  });
};

const withPrototypeCanary = <T>(key: string, value: unknown, run: () => T): T => {
  const previous = Object.getOwnPropertyDescriptor(Object.prototype, key);
  Object.defineProperty(Object.prototype, key, {
    value,
    enumerable: false,
    configurable: true,
    writable: true,
  });

  try {
    return run();
  } finally {
    if (previous === undefined) {
      Reflect.deleteProperty(Object.prototype, key);
    } else {
      Object.defineProperty(Object.prototype, key, previous);
    }
  }
};

const withOwnPropertyCanary = <T>(
  target: object,
  key: string,
  value: unknown,
  run: () => T,
): T => {
  const previous = Object.getOwnPropertyDescriptor(target, key);
  Object.defineProperty(target, key, {
    value,
    enumerable: false,
    configurable: true,
    writable: true,
  });

  try {
    return run();
  } finally {
    if (previous === undefined) {
      Reflect.deleteProperty(target, key);
    } else {
      Object.defineProperty(target, key, previous);
    }
  }
};

const invalidNumericValues = [
  -1,
  1.5,
  Number.NaN,
  Number.POSITIVE_INFINITY,
  Number.NEGATIVE_INFINITY,
  Number.MAX_SAFE_INTEGER + 1,
  "1" as unknown as number,
  {} as unknown as number,
] as const;

const invalidResourceShapeCases = RESOURCE_KEYS.flatMap((key) =>
  invalidNumericValues.map((invalidValue) => ({ key, invalidValue })),
);

const exactCapabilities = [...WORKBOOK_PLAN_CAPABILITY_NAMES];

const createActiveWritePlan = (): WorkbookPlanPolicyInput =>
  makeInput({
    requestClassification: "explicit_mutation",
    operationSummaries: [createOperationSummary(0, "write", ["Sheet1"], "inside")],
  });

describe("workbook plan policy", () => {
  it("derives contract-native effect and capability vocabularies", () => {
    expect(WORKBOOK_PLAN_OPERATION_EFFECTS).toEqual([
      "write",
      "format",
      "structure",
      "filter",
      "calculate",
    ]);
    expect(WORKBOOK_PLAN_CAPABILITY_NAMES).toEqual([
      "excel_workbook_access",
      "workbook_plan_v1_operation_audit",
    ]);
  });

  it("freezes exported vocabularies and prevents mutation from changing policy authority", () => {
    expect(Object.isFrozen(WORKBOOK_PLAN_OPERATION_EFFECTS)).toBe(true);
    expect(Object.isFrozen(WORKBOOK_PLAN_CAPABILITY_NAMES)).toBe(true);

    try {
      (WORKBOOK_PLAN_OPERATION_EFFECTS as unknown as string[]).push("destructive");
    } catch {
      // Frozen exports may throw in strict mode; either way, authority must not change.
    }
    try {
      (WORKBOOK_PLAN_CAPABILITY_NAMES as unknown as string[]).push("bogus_capability");
    } catch {
      // Frozen exports may throw in strict mode; either way, authority must not change.
    }

    expectDecision(
      makeInput({
        confirmation: { confirmed: true },
        operationSummaries: [
          createOperationSummary(0, "destructive" as unknown as WorkbookPlanOperationEffect, ["Sheet1"], "inside"),
        ],
      }),
      "rejected",
      "unsupported_effect",
    );
    expectDecision(
      makeInput({
        capabilitySummaries: [
          ...WORKBOOK_PLAN_CAPABILITY_NAMES,
          "bogus_capability" as unknown as WorkbookPlanCapabilityName,
        ],
      }),
      "rejected",
      "invalid_trusted_input",
    );
  });

  it("rejects a proxy around a registered snapshot before executing any proxy traps", () => {
    const snapshot = createWorkbookPlanPolicyInputSnapshot(makeInput());
    let trapCount = 0;
    const proxy = new Proxy(snapshot, {
      getPrototypeOf() {
        trapCount += 1;
        throw new Error("getPrototypeOf trap should not run");
      },
      ownKeys() {
        trapCount += 1;
        throw new Error("ownKeys trap should not run");
      },
      getOwnPropertyDescriptor() {
        trapCount += 1;
        throw new Error("getOwnPropertyDescriptor trap should not run");
      },
      get() {
        trapCount += 1;
        throw new Error("get trap should not run");
      },
    });

    expect(decideWorkbookPlanPolicy(proxy)).toEqual({
      decision: "rejected",
      reasonCode: "invalid_trusted_input",
    });
    expect(trapCount).toBe(0);
  });

  it("approves a calculate-only plan with zero sheet targets and exact contract capabilities", () => {
    const input = makeInput({
      requestClassification: "explicit_mutation",
      selectionScopedRequest: true,
      operationSummaries: [createOperationSummary(0, "calculate", [], "not_applicable")],
    });

    expectDecision(input, "approved", "approved");
  });

  it.each(["write", "format", "structure", "filter"] as const)(
    "rejects %s operations with zero sheet targets even when otherwise trusted and confirmed",
    (effect) => {
      const input = makeInput({
        confirmation: { confirmed: true },
        operationSummaries: [createOperationSummary(0, effect, [], "not_applicable")],
      });

      expectDecision(input, "rejected", "invalid_trusted_input");
    },
  );

  it.each(["destructive", "unknown"] as const)(
    "rejects %s effect strings even when confirmed",
    (effect) => {
      const input = makeInput({
        confirmation: { confirmed: true },
        operationSummaries: [
          createOperationSummary(0, effect as unknown as WorkbookPlanOperationEffect, ["Sheet1"], "inside"),
        ],
      });

      expectDecision(input, "rejected", "unsupported_effect");
    },
  );

  it("requires confirmation for ambiguous mutation intent and resolves it when confirmed", () => {
    const unconfirmed = makeInput({
      requestClassification: "ambiguous_mutation",
      operationSummaries: [createOperationSummary(0, "write", ["Sheet1"], "inside")],
    });
    expectDecision(unconfirmed, "requires_confirmation", "ambiguous_mutation_intent");

    const confirmed = makeInput({
      requestClassification: "ambiguous_mutation",
      confirmation: { confirmed: true },
      operationSummaries: [createOperationSummary(0, "write", ["Sheet1"], "inside")],
    });
    expectDecision(confirmed, "approved", "approved");
  });

  it("requires exact structural intent for structure effects", () => {
    const needsConfirmation = makeInput({
      requestClassification: "explicit_mutation",
      operationSummaries: [createOperationSummary(0, "structure", ["Sheet1"], "inside")],
    });
    expectDecision(needsConfirmation, "requires_confirmation", "structural_intent_required");

    const confirmed = makeInput({
      requestClassification: "explicit_mutation",
      confirmation: { confirmed: true },
      operationSummaries: [createOperationSummary(0, "structure", ["Sheet1"], "inside")],
    });
    expectDecision(confirmed, "approved", "approved");

    const explicitStructural = makeInput({
      requestClassification: "explicit_structure",
      operationSummaries: [createOperationSummary(0, "structure", ["Sheet1"], "inside")],
    });
    expectDecision(explicitStructural, "approved", "approved");
  });

  it("approves exact contract capabilities and rejects missing or duplicate entries", () => {
    const exact = makeInput({
      operationSummaries: [createOperationSummary(0, "calculate", [], "not_applicable")],
      capabilitySummaries: [...WORKBOOK_PLAN_CAPABILITY_NAMES],
    });
    expectDecision(exact, "approved", "approved");

    const missing = makeInput({
      operationSummaries: [createOperationSummary(0, "calculate", [], "not_applicable")],
      capabilitySummaries: [WORKBOOK_PLAN_CAPABILITY_NAMES[0]],
    });
    expectDecision(missing, "rejected", "missing_contract_capability");

    const duplicate = makeInput({
      operationSummaries: [createOperationSummary(0, "calculate", [], "not_applicable")],
      capabilitySummaries: [WORKBOOK_PLAN_CAPABILITY_NAMES[0], WORKBOOK_PLAN_CAPABILITY_NAMES[0]],
    });
    expectDecision(duplicate, "rejected", "invalid_trusted_input");
  });

  it("requires confirmation when any target is off the active sheet, including multi-sheet descriptors", () => {
    const input = makeInput({
      requestClassification: "explicit_mutation",
      operationSummaries: [
        createOperationSummary(0, "write", ["Sheet1", "Sheet2"], "inside"),
      ],
    });

    expectDecision(input, "requires_confirmation", "active_sheet_mismatch");
  });

  it("requires confirmation for selection-scoped outside and unknown targets but not for not-applicable targets", () => {
    const input = makeInput({
      requestClassification: "explicit_mutation",
      selectionScopedRequest: true,
      operationSummaries: [
        createOperationSummary(0, "calculate", [], "not_applicable"),
        createOperationSummary(1, "write", ["Sheet1"], "inside"),
        createOperationSummary(2, "write", ["Sheet1"], "outside"),
      ],
    });

    expectDecision(input, "requires_confirmation", "selection_scope_mismatch");

    const unknownInput = makeInput({
      requestClassification: "explicit_mutation",
      selectionScopedRequest: true,
      operationSummaries: [
        createOperationSummary(0, "calculate", [], "not_applicable"),
        createOperationSummary(1, "write", ["Sheet1"], "inside"),
        createOperationSummary(2, "write", ["Sheet1"], "unknown"),
      ],
    });

    expectDecision(unknownInput, "requires_confirmation", "selection_scope_mismatch");
  });

  it("requires confirmation for read-only requests that would mutate workbook state", () => {
    const input = makeInput({
      requestClassification: "read_only",
      operationSummaries: [createOperationSummary(0, "write", ["Sheet1"], "inside")],
    });

    expectDecision(input, "requires_confirmation", "confirmation_required");
  });

  it("rejects zero-operation plans", () => {
    const input = makeInput({
      operationSummaries: [],
      resourceUsage: { operationCount: 0 },
      resourceLimits: { operationCount: 0 },
    });

    expectDecision(input, "rejected", "empty_plan");
  });

  it("rejects inconsistent operation counts and duplicate or out-of-range operation indices", () => {
    const missingDescriptor = makeInput({
      resourceUsage: { operationCount: 2 },
      resourceLimits: { operationCount: 2 },
      operationSummaries: [createOperationSummary(0, "write", ["Sheet1"], "inside")],
    });
    expectDecision(missingDescriptor, "rejected", "inconsistent_operation_counts");

    const duplicateIndex = makeInput({
      resourceUsage: { operationCount: 2 },
      resourceLimits: { operationCount: 2 },
      operationSummaries: [
        createOperationSummary(0, "write", ["Sheet1"], "inside"),
        createOperationSummary(0, "format", ["Sheet1"], "inside"),
      ],
    });
    expectDecision(duplicateIndex, "rejected", "inconsistent_operation_counts");

    const outOfRangeIndex = makeInput({
      resourceUsage: { operationCount: 2 },
      resourceLimits: { operationCount: 2 },
      operationSummaries: [
        createOperationSummary(0, "write", ["Sheet1"], "inside"),
        createOperationSummary(2, "format", ["Sheet1"], "inside"),
      ],
    });
    expectDecision(outOfRangeIndex, "rejected", "inconsistent_operation_counts");
  });

  it.each(invalidResourceShapeCases)(
    "fails closed for invalid resource values on %s",
    ({ key, invalidValue }) => {
      const limitInput = makeInput({
        resourceLimits: { [key]: invalidValue } as Partial<Record<ResourceKey, number>>,
      });
      expectDecision(limitInput, "rejected", "invalid_trusted_input");

      const usageInput = makeInput({
        resourceUsage: { [key]: invalidValue } as Partial<Record<ResourceKey, number>>,
      });
      expectDecision(usageInput, "rejected", "invalid_trusted_input");
    },
  );

  it("accepts exact resource limits and rejects every over-limit metric", () => {
    const exact = makeInput({
      operationSummaries: [createOperationSummary(0, "write", ["Sheet1"], "inside")],
      resourceUsage: createResourceSummary({}, 1),
      resourceLimits: createResourceSummary({}, 1),
    });
    expectDecision(exact, "approved", "approved");

    for (const key of RESOURCE_KEYS) {
      const input = makeInput({
        operationSummaries: [createOperationSummary(0, "write", ["Sheet1"], "inside")],
        resourceUsage: createResourceSummary({ [key]: BASE_RESOURCE[key] + 1 }, 1),
      });
      expectDecision(
        input,
        key === "operationCount" ? "rejected" : "rejected",
        key === "operationCount" ? "inconsistent_operation_counts" : "resource_limit_exceeded",
      );
    }
  });

  it("rejects malformed collections, sparse arrays, shadowed methods, duplicate sheet names, and accessor arrays", () => {
    const sparseCapabilities = [...WORKBOOK_PLAN_CAPABILITY_NAMES];
    delete sparseCapabilities[1];
    expectDecision(
      makeInput({ capabilitySummaries: sparseCapabilities as readonly WorkbookPlanCapabilityName[] }),
      "rejected",
      "invalid_trusted_input",
    );

    const shadowedCapabilities = [...WORKBOOK_PLAN_CAPABILITY_NAMES];
    withOwnPropertyCanary(shadowedCapabilities, "forEach", () => undefined, () => {
      expectDecision(
        makeInput({ capabilitySummaries: shadowedCapabilities as readonly WorkbookPlanCapabilityName[] }),
        "rejected",
        "invalid_trusted_input",
      );
    });

    const duplicateSheetNamesPlan = makeInput({
      operationSummaries: [createOperationSummary(0, "write", ["Sheet1", "Sheet1"], "inside")],
    });
    expectDecision(duplicateSheetNamesPlan, "rejected", "invalid_trusted_input");

    const accessorSheetNames = ["Sheet1"];
    Object.defineProperty(accessorSheetNames, "0", {
      get() {
        throw new Error("getter should not run");
      },
      configurable: true,
      enumerable: true,
    });
    const accessorPlan = makeInput({
      operationSummaries: [
        createOperationSummary(0, "write", accessorSheetNames as readonly string[], "inside"),
      ],
    });
    expectDecision(accessorPlan, "rejected", "invalid_trusted_input");
  });

  it("wraps revoked proxies at the top level and nested collection positions", () => {
    const topLevel = Proxy.revocable(makeInput(), {});
    topLevel.revoke();
    expectDecision(topLevel.proxy, "rejected", "invalid_trusted_input");

    const nested = makeInput();
    const revokedOperationSummaries = Proxy.revocable(nested.operationSummaries as unknown as object, {});
    revokedOperationSummaries.revoke();
    expectDecision(
      {
        ...nested,
        operationSummaries: revokedOperationSummaries.proxy as unknown as readonly WorkbookPlanOperationSummary[],
      },
      "rejected",
      "invalid_trusted_input",
    );

    const revokedCapabilities = Proxy.revocable(nested.capabilitySummaries as unknown as object, {});
    revokedCapabilities.revoke();
    expectDecision(
      {
        ...nested,
        capabilitySummaries: revokedCapabilities.proxy as unknown as readonly WorkbookPlanCapabilityName[],
      },
      "rejected",
      "invalid_trusted_input",
    );
  });

  it("ignores poisoned prototype keys and leaves prototypes restored", () => {
    withPrototypeCanary("confirmation", { confirmed: true }, () => {
      const input = makeInput({
        requestClassification: "ambiguous_mutation",
        operationSummaries: [createOperationSummary(0, "write", ["Sheet1"], "inside")],
      });
      expectDecision(input, "requires_confirmation", "ambiguous_mutation_intent");
    });
  });
});
