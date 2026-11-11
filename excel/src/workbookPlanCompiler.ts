import type {
  ScalarValue,
  VerificationRule,
  WorkbookOperation,
  WorkbookPlan,
  ValueMatrix,
  FormulaMatrix,
  FormatSpec,
} from "./generated/workbookPlanSchema";
import { workbookPlanContract } from "./workbookPlanContract";
import { getExcelToolDefinition } from "./excelTools";

export type CompileWorkbookPlanErrorCode =
  | "unsupported_outcome"
  | "missing_unsupported_primitive_capability"
  | "unresolved_sheet"
  | "unresolved_table"
  | "out_of_range_trusted_table_column"
  | "unresolved_chart_source"
  | "unsupported_noncontiguous_chart"
  | "missing_mandatory_verification"
  | "incomplete_precondition_manifest"
  | "contract_definition_drift";

export interface CompileWorkbookPlanError {
  readonly code: CompileWorkbookPlanErrorCode;
  readonly path: string;
  readonly message: string;
}

export interface VerifiedWorkbookTable {
  readonly canonicalName: string;
  readonly columnCount: number;
  readonly rowCount?: number;
  readonly sheetName?: string;
}

export interface VerifiedWorkbookChartSourceResolution {
  readonly areaCount: number;
}

export interface VerifiedWorkbookCompileContext {
  readonly sheetNames: ReadonlySet<string>;
  readonly tablesByName: ReadonlyMap<string, VerifiedWorkbookTable>;
  readonly supportedPrimitiveNames: ReadonlySet<string>;
  readonly supportedCapabilities: ReadonlySet<string>;
  readonly chartSourceResolutionsBySheet: ReadonlyMap<
    string,
    ReadonlyMap<string, VerifiedWorkbookChartSourceResolution>
  >;
  /** Host-owned facts captured before preparation and rechecked before mutation. */
  readonly preconditionManifest?: WorkbookPlanPreconditionManifest;
}

export interface WorkbookPlanPreconditionSheet {
  readonly name: string;
  readonly exists: boolean;
}

export interface WorkbookPlanPreconditionTable {
  readonly name: string;
  readonly exists: boolean;
  readonly columnCount?: number;
  readonly rowCount?: number;
  readonly sheetName?: string;
  readonly address?: string;
}

export interface WorkbookPlanPreconditionRange {
  readonly sheet: string;
  readonly address: string;
  /** Host-resolved address; the lookup address above remains operation-relative. */
  readonly resolvedAddress?: string;
  readonly rowCount: number;
  readonly columnCount: number;
  /** First-cell sentinel captured from the trusted host pre-state. */
  readonly sentinel?: ScalarValue;
  /** Complete bounded values captured before a write_values operation. */
  readonly values?: ValueMatrix;
  /** Source formula pre-state used to bind a fill proof. */
  readonly sourceFormulas?: FormulaMatrix;
  /** Captured number format used when the operation does not change it. */
  readonly numberFormat?: string;
}

export interface WorkbookPlanPreconditionChartSource {
  readonly sheet: string;
  readonly address: string;
  readonly areaCount: number;
  readonly resolvedAddress?: string;
}

export interface WorkbookPlanPreconditionChart {
  readonly sheet: string;
  readonly title: string;
  readonly exists: boolean;
}

export interface WorkbookPlanPreconditionComment {
  readonly sheet: string;
  readonly address: string;
  readonly exists: boolean;
  readonly text?: string;
}

export interface WorkbookPlanPreconditionManifest {
  readonly version: 1;
  readonly sheets: readonly WorkbookPlanPreconditionSheet[];
  readonly tables: readonly WorkbookPlanPreconditionTable[];
  readonly ranges: readonly WorkbookPlanPreconditionRange[];
  readonly chartSources: readonly WorkbookPlanPreconditionChartSource[];
  readonly charts: readonly WorkbookPlanPreconditionChart[];
  readonly comments: readonly WorkbookPlanPreconditionComment[];
}

type VirtualChartSourceResolutionMap = Map<string, VerifiedWorkbookChartSourceResolution>;
type VirtualChartSourceResolutionsBySheet = Map<string, VirtualChartSourceResolutionMap>;

export type CalculationRequirement = "required_before_verification" | "not_required";
export type PrimitiveResultHandling = "passthrough";

export interface CompiledWorkbookPrimitive {
  readonly operationIndex: number;
  readonly kind: WorkbookOperation["kind"];
  readonly primitiveName: string;
  readonly args: Readonly<Record<string, unknown>>;
  readonly effect: (typeof workbookPlanContract.operationDefinitions)[number]["effect"];
  readonly requiredVerificationKinds: readonly VerificationRule["kind"][];
  readonly calculationRequirement: CalculationRequirement;
  readonly resultHandling: PrimitiveResultHandling;
}

export interface WorkbookPlanInternalProof {
  readonly operationIndex: number;
  /** Content-free operation binding; executor validates the digest against its trusted operation. */
  readonly operationKind: WorkbookOperation["kind"];
  readonly primitiveName?: string;
  readonly effect?: CompiledWorkbookPrimitive["effect"];
  readonly argsDigest?: string;
  readonly expectedFormulas?: FormulaMatrix;
}

export interface CompiledWorkbookMandatoryPostcondition {
  readonly operationIndex: number;
  readonly kind: VerificationRule["kind"];
  readonly proof?: WorkbookPlanInternalProof;
}

export interface CompiledWorkbookTransaction {
  readonly primitives: readonly CompiledWorkbookPrimitive[];
  readonly postconditions: readonly VerificationRule[];
  /** B1 compiler-owned proofs. Non-enumerable for Slice A object-shape compatibility. */
  readonly mandatoryPostconditions?: readonly CompiledWorkbookMandatoryPostcondition[];
  /** False means at least one primitive has no supported target-correlated proof. */
  readonly mandatoryProofsComplete?: boolean;
}

export interface CompileWorkbookPlanSuccess {
  readonly ok: true;
  readonly transaction: CompiledWorkbookTransaction;
}

export interface CompileWorkbookPlanFailure {
  readonly ok: false;
  readonly error: CompileWorkbookPlanError;
}

export type CompileWorkbookPlanResult =
  | CompileWorkbookPlanSuccess
  | CompileWorkbookPlanFailure;

type OperationDefinition = (typeof workbookPlanContract.operationDefinitions)[number];
type OperationKind = WorkbookOperation["kind"];
type VerificationKind = VerificationRule["kind"];
type OperationMetadata = {
  readonly effect: OperationDefinition["effect"];
  readonly compilerTarget: string;
  readonly requiredVerificationKinds: readonly VerificationKind[];
};

type CompiledOperationSuccess = { readonly ok: true; readonly primitive: CompiledWorkbookPrimitive };
type CompiledOperationFailure = { readonly ok: false; readonly error: CompileWorkbookPlanError };
type CompiledOperationOutcome = CompiledOperationSuccess | CompiledOperationFailure;

type ContractValidationSuccess = {
  readonly ok: true;
  readonly definitionsByKind: ReadonlyMap<OperationKind, OperationDefinition>;
};
type ContractValidationFailure = { readonly ok: false; readonly error: CompileWorkbookPlanError };
type ContractValidationOutcome = ContractValidationSuccess | ContractValidationFailure;

const expectedOperationMetadata = {
  write_values: {
    effect: "write",
    compilerTarget: "write_range",
    requiredVerificationKinds: ["range_values_equal"],
  },
  write_formula: {
    effect: "write",
    compilerTarget: "write_range",
    requiredVerificationKinds: ["range_formulas_equal"],
  },
  fill_formula: {
    effect: "write",
    compilerTarget: "autofill_formula",
    requiredVerificationKinds: ["range_formula_pattern"],
  },
  format_range: {
    effect: "format",
    compilerTarget: "format_range",
    requiredVerificationKinds: ["range_number_format_equal"],
  },
  create_table: {
    effect: "structure",
    compilerTarget: "insert_table",
    requiredVerificationKinds: ["table_exists"],
  },
  sort_table: {
    effect: "write",
    compilerTarget: "sort_table",
    requiredVerificationKinds: ["table_row_order"],
  },
  filter_table: {
    effect: "filter",
    compilerTarget: "filter_table",
    requiredVerificationKinds: ["table_filter_state"],
  },
  reset_table_filter: {
    effect: "filter",
    compilerTarget: "filter_table",
    requiredVerificationKinds: ["table_filter_state"],
  },
  create_chart: {
    effect: "structure",
    compilerTarget: "create_chart",
    requiredVerificationKinds: ["chart_exists"],
  },
  set_validation_list: {
    effect: "write",
    compilerTarget: "set_data_validation",
    requiredVerificationKinds: ["sentinel_unchanged"],
  },
  conditional_format: {
    effect: "format",
    compilerTarget: "conditional_format_range",
    requiredVerificationKinds: ["sentinel_unchanged"],
  },
  freeze_panes: {
    effect: "format",
    compilerTarget: "freeze_panes",
    requiredVerificationKinds: ["sentinel_unchanged"],
  },
  create_sheet: {
    effect: "structure",
    compilerTarget: "create_sheet",
    requiredVerificationKinds: ["sheet_exists"],
  },
  rename_sheet: {
    effect: "structure",
    compilerTarget: "rename_sheet",
    requiredVerificationKinds: ["sheet_exists"],
  },
  copy_sheet: {
    effect: "structure",
    compilerTarget: "copy_sheet",
    requiredVerificationKinds: ["sheet_exists"],
  },
  add_comment: {
    effect: "write",
    compilerTarget: "add_comment",
    requiredVerificationKinds: ["comment_exists"],
  },
  calculate: {
    effect: "calculate",
    compilerTarget: "run_calculation",
    requiredVerificationKinds: ["calculation_completed"],
  },
} as const satisfies Readonly<Record<OperationKind, OperationMetadata>>;

const expectedCapabilityRequirements = [
  { capability: "excel_workbook_access", status: "baseline" },
  { capability: "workbook_plan_v1_operation_audit", status: "unverified" },
] as const;

const errorMessages: Readonly<Record<CompileWorkbookPlanErrorCode, string>> = {
  unsupported_outcome: "Only execute plans can be compiled.",
  missing_unsupported_primitive_capability: "Workbook plan cannot be compiled with the current verified context.",
  unresolved_sheet: "Workbook plan references an unavailable sheet.",
  unresolved_table: "Workbook plan references an unavailable table.",
  out_of_range_trusted_table_column: "Workbook plan references an invalid trusted table column.",
  unresolved_chart_source: "Workbook plan references an unavailable chart source.",
  unsupported_noncontiguous_chart: "Workbook plan references a noncontiguous chart source.",
  missing_mandatory_verification: "Workbook plan operation has no complete supported target-correlated proof.",
  incomplete_precondition_manifest: "Workbook plan trusted precondition manifest is incomplete.",
  contract_definition_drift: "Workbook plan compiler metadata is inconsistent.",
};

function isContractValidationFailure(
  result: ContractValidationOutcome,
): result is ContractValidationFailure {
  return result.ok === false;
}

function isCompiledOperationFailure(
  result: CompiledOperationOutcome,
): result is CompiledOperationFailure {
  return result.ok === false;
}

function createFailure(
  code: CompileWorkbookPlanErrorCode,
  path: string,
): CompileWorkbookPlanFailure {
  return {
    ok: false,
    error: {
      code,
      path,
      message: errorMessages[code],
    },
  };
}

function createArgs(entries: readonly (readonly [string, unknown | undefined])[]): Record<string, unknown> {
  const args: Record<string, unknown> = {};
  for (const [key, value] of entries) {
    if (value !== undefined) {
      args[key] = value;
    }
  }
  return args;
}

function cloneValueMatrix(matrix: ValueMatrix): ValueMatrix {
  return matrix.map((row) => [...row]);
}

function cloneFormulaMatrix(matrix: FormulaMatrix): FormulaMatrix {
  return matrix.map((row) => [...row]);
}

function cloneVerificationRule(rule: VerificationRule): VerificationRule {
  switch (rule.kind) {
    case "range_values_equal":
      return {
        kind: rule.kind,
        sheet: rule.sheet,
        address: rule.address,
        expectedValues: cloneValueMatrix(rule.expectedValues),
      };
    case "range_formulas_equal":
      return {
        kind: rule.kind,
        sheet: rule.sheet,
        address: rule.address,
        expectedFormulas: cloneFormulaMatrix(rule.expectedFormulas),
      };
    case "range_formula_pattern":
      return {
        kind: rule.kind,
        sheet: rule.sheet,
        address: rule.address,
        pattern: rule.pattern,
      };
    case "range_number_format_equal":
      return {
        kind: rule.kind,
        sheet: rule.sheet,
        address: rule.address,
        expectedNumberFormat: rule.expectedNumberFormat,
      };
    case "table_exists":
      return {
        kind: rule.kind,
        tableName: rule.tableName,
      };
    case "table_row_order":
      return {
        kind: rule.kind,
        tableName: rule.tableName,
        expectedOrder: [...rule.expectedOrder],
      };
    case "table_filter_state":
      return {
        kind: rule.kind,
        tableName: rule.tableName,
        columnIndex: rule.columnIndex,
        state: rule.state,
        ...(rule.values === undefined ? {} : { values: [...rule.values] }),
        ...(rule.criterion1 === undefined ? {} : { criterion1: rule.criterion1 }),
        ...(rule.criterion2 === undefined ? {} : { criterion2: rule.criterion2 }),
        ...(rule.operator === undefined ? {} : { operator: rule.operator }),
      };
    case "sheet_exists":
      return {
        kind: rule.kind,
        name: rule.name,
      };
    case "chart_exists":
      return {
        kind: rule.kind,
        sheet: rule.sheet,
        title: rule.title,
      };
    case "comment_exists":
      return {
        kind: rule.kind,
        sheet: rule.sheet,
        address: rule.address,
        text: rule.text,
      };
    case "sentinel_unchanged":
      return {
        kind: rule.kind,
        sheet: rule.sheet,
        address: rule.address,
        expected: rule.expected,
      };
    case "calculation_completed":
      return {
        kind: rule.kind,
      };
    default: {
      const exhaustiveCheck: never = rule;
      return exhaustiveCheck;
    }
  }
}

function cloneVerificationRules(rules: readonly VerificationRule[]): readonly VerificationRule[] {
  return rules.map((rule) => cloneVerificationRule(rule));
}

function cloneChartSourceResolutionMap(
  source: ReadonlyMap<string, VerifiedWorkbookChartSourceResolution>,
): VirtualChartSourceResolutionMap {
  const clone = new Map<string, VerifiedWorkbookChartSourceResolution>();
  for (const [dataRange, resolution] of source) {
    clone.set(dataRange, { areaCount: resolution.areaCount });
  }
  return clone;
}

function cloneChartSourceResolutionsBySheet(
  source: ReadonlyMap<string, ReadonlyMap<string, VerifiedWorkbookChartSourceResolution>>,
): VirtualChartSourceResolutionsBySheet {
  const clone = new Map<string, VirtualChartSourceResolutionMap>();
  for (const [sheet, chartSourceByRange] of source) {
    clone.set(sheet, cloneChartSourceResolutionMap(chartSourceByRange));
  }
  return clone;
}

function arraysMatch(left: readonly string[], right: readonly string[]): boolean {
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}

function capabilityRequirementsMatch(
  left: readonly { readonly capability: string; readonly status: string }[],
  right: readonly { readonly capability: string; readonly status: string }[],
): boolean {
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) {
    if (left[index].capability !== right[index].capability) return false;
    if (left[index].status !== right[index].status) return false;
  }
  return true;
}

function validateContractDefinitions(): ContractValidationOutcome {
  const definitions = workbookPlanContract.operationDefinitions;
  const kinds = workbookPlanContract.operationKinds;

  if (definitions.length !== kinds.length) {
    return createFailure("contract_definition_drift", "/operationDefinitions");
  }

  const definitionsByKind = new Map<OperationKind, OperationDefinition>();

  for (let index = 0; index < kinds.length; index += 1) {
    const expectedKind = kinds[index];
    const definition = definitions[index];
    const metadata = expectedOperationMetadata[expectedKind];

    if (definition.kind !== expectedKind) {
      return createFailure("contract_definition_drift", `/operationDefinitions/${index}/kind`);
    }
    if (definitionsByKind.has(definition.kind)) {
      return createFailure("contract_definition_drift", `/operationDefinitions/${index}/kind`);
    }
    if (definition.effect !== metadata.effect) {
      return createFailure("contract_definition_drift", `/operationDefinitions/${index}/effect`);
    }
    if (definition.compilerTargets.length !== 1 || definition.compilerTargets[0] !== metadata.compilerTarget) {
      return createFailure("contract_definition_drift", `/operationDefinitions/${index}/compilerTargets`);
    }
    if (!arraysMatch(definition.requiredVerification, metadata.requiredVerificationKinds)) {
      return createFailure("contract_definition_drift", `/operationDefinitions/${index}/requiredVerification`);
    }
    if (
      !capabilityRequirementsMatch(
        definition.capabilityRequirements,
        expectedCapabilityRequirements,
      )
    ) {
      return createFailure("contract_definition_drift", `/operationDefinitions/${index}/capabilityRequirements`);
    }

    definitionsByKind.set(definition.kind, definition);
  }

  if (definitionsByKind.size !== kinds.length) {
    return createFailure("contract_definition_drift", "/operationDefinitions");
  }

  return { ok: true, definitionsByKind };
}

function hasSupportedCapabilities(
  supportedCapabilities: ReadonlySet<string>,
  definition: OperationDefinition,
): boolean {
  return definition.capabilityRequirements.every(({ capability }) => supportedCapabilities.has(capability));
}

function getCompiledPrimitiveName(definition: OperationDefinition): string | null {
  const compilerTarget = definition.compilerTargets[0];
  const toolDefinition = getExcelToolDefinition(compilerTarget);
  if (toolDefinition === undefined) return null;
  if (toolDefinition.schema.function.name !== compilerTarget) return null;
  return compilerTarget;
}

function ensureSheetExists(
  sheetNames: ReadonlySet<string>,
  sheetName: string,
  path: string,
): CompileWorkbookPlanFailure | null {
  if (!sheetNames.has(sheetName)) {
    return createFailure("unresolved_sheet", path);
  }
  return null;
}

function compileTableOperation(
  operationIndex: number,
  operation: Extract<WorkbookOperation, { readonly kind: "sort_table" | "filter_table" | "reset_table_filter" }>,
  definition: OperationDefinition,
  context: VerifiedWorkbookCompileContext,
): CompiledOperationOutcome {
  const table = context.tablesByName.get(operation.tableName);
  if (table === undefined) {
    return {
      ok: false,
      error: createFailure("unresolved_table", `/operations/${operationIndex}/tableName`).error,
    };
  }

  if (
    !Number.isInteger(table.columnCount) ||
    table.columnCount < 0 ||
    operation.columnIndex < 0 ||
    operation.columnIndex >= table.columnCount
  ) {
    return {
      ok: false,
      error: createFailure("out_of_range_trusted_table_column", `/operations/${operationIndex}/columnIndex`).error,
    };
  }

  if (operation.kind === "sort_table") {
    return {
      ok: true,
      primitive: {
        operationIndex,
        kind: operation.kind,
        primitiveName: definition.compilerTargets[0],
        args: createArgs([
          ["tableName", table.canonicalName],
          ["columnIndex", operation.columnIndex],
          ["ascending", operation.ascending !== false],
        ]),
        effect: definition.effect,
        requiredVerificationKinds: [...definition.requiredVerification],
        calculationRequirement: "not_required",
        resultHandling: "passthrough",
      },
    };
  }

  if (operation.kind === "filter_table") {
    const operationPath = `/operations/${operationIndex}`;
    const hasValues = operation.values !== undefined;
    const hasCriterion1 = operation.criterion1 !== undefined;
    const hasCriterion2 = operation.criterion2 !== undefined;
    const hasOperator = operation.operator !== undefined;
    const malformedValues = hasValues && (!Array.isArray(operation.values) || operation.values.length === 0 || operation.values.some((value) => typeof value !== "string"));
    const malformedCustom = hasCriterion1 && typeof operation.criterion1 !== "string" ||
      hasCriterion2 && typeof operation.criterion2 !== "string" ||
      hasOperator && operation.operator !== "And" && operation.operator !== "Or";
    if (
      malformedValues || malformedCustom ||
      hasValues && (hasCriterion1 || hasCriterion2 || hasOperator) ||
      !hasValues && !hasCriterion1
    ) {
      return { ok: false, error: createFailure("missing_mandatory_verification", operationPath).error };
    }
    return {
      ok: true,
      primitive: {
        operationIndex,
        kind: operation.kind,
        primitiveName: definition.compilerTargets[0],
        args: createArgs([
          ["tableName", table.canonicalName],
          ["columnIndex", operation.columnIndex],
          ["values", operation.values === undefined ? undefined : [...operation.values]],
          ["criterion1", operation.criterion1],
          ["criterion2", operation.criterion2],
          ["operator", operation.operator],
        ]),
        effect: definition.effect,
        requiredVerificationKinds: [...definition.requiredVerification],
        calculationRequirement: "not_required",
        resultHandling: "passthrough",
      },
    };
  }

  return {
    ok: true,
    primitive: {
      operationIndex,
      kind: operation.kind,
      primitiveName: definition.compilerTargets[0],
      args: createArgs([
        ["tableName", table.canonicalName],
        ["columnIndex", operation.columnIndex],
      ]),
      effect: definition.effect,
      requiredVerificationKinds: [...definition.requiredVerification],
      calculationRequirement: "not_required",
      resultHandling: "passthrough",
    },
  };
}

function compileSheetChangingOperation(
  operationIndex: number,
  operation: Extract<WorkbookOperation, { readonly kind: "create_sheet" | "rename_sheet" | "copy_sheet" }>,
  definition: OperationDefinition,
  virtualSheetNames: Set<string>,
  virtualChartSourceResolutionsBySheet: VirtualChartSourceResolutionsBySheet,
): CompiledOperationOutcome {
  if (operation.kind === "create_sheet") {
    if (virtualSheetNames.has(operation.name)) {
      return {
        ok: false,
        error: createFailure("unresolved_sheet", `/operations/${operationIndex}/name`).error,
      };
    }
    virtualSheetNames.add(operation.name);
    return {
      ok: true,
      primitive: {
        operationIndex,
        kind: operation.kind,
        primitiveName: definition.compilerTargets[0],
        args: createArgs([["name", operation.name]]),
        effect: definition.effect,
        requiredVerificationKinds: [...definition.requiredVerification],
        calculationRequirement: "not_required",
        resultHandling: "passthrough",
      },
    };
  }

  if (operation.kind === "rename_sheet") {
    if (!virtualSheetNames.has(operation.oldName)) {
      return {
        ok: false,
        error: createFailure("unresolved_sheet", `/operations/${operationIndex}/oldName`).error,
      };
    }
    if (operation.oldName === operation.newName) {
      return {
        ok: true,
        primitive: {
          operationIndex,
          kind: operation.kind,
          primitiveName: definition.compilerTargets[0],
          args: createArgs([
            ["oldName", operation.oldName],
            ["newName", operation.newName],
          ]),
          effect: definition.effect,
          requiredVerificationKinds: [...definition.requiredVerification],
          calculationRequirement: "not_required",
          resultHandling: "passthrough",
        },
      };
    }
    if (virtualSheetNames.has(operation.newName)) {
      return {
        ok: false,
        error: createFailure("unresolved_sheet", `/operations/${operationIndex}/newName`).error,
      };
    }
    virtualSheetNames.delete(operation.oldName);
    virtualSheetNames.add(operation.newName);

    const chartSourceBySheet = virtualChartSourceResolutionsBySheet.get(operation.oldName);
    if (chartSourceBySheet !== undefined) {
      virtualChartSourceResolutionsBySheet.delete(operation.oldName);
      virtualChartSourceResolutionsBySheet.set(operation.newName, chartSourceBySheet);
    }

    return {
      ok: true,
      primitive: {
        operationIndex,
        kind: operation.kind,
        primitiveName: definition.compilerTargets[0],
        args: createArgs([
          ["oldName", operation.oldName],
          ["newName", operation.newName],
        ]),
        effect: definition.effect,
        requiredVerificationKinds: [...definition.requiredVerification],
        calculationRequirement: "not_required",
        resultHandling: "passthrough",
      },
    };
  }

  if (!virtualSheetNames.has(operation.sourceName)) {
    return {
      ok: false,
      error: createFailure("unresolved_sheet", `/operations/${operationIndex}/sourceName`).error,
    };
  }
  if (virtualSheetNames.has(operation.newName)) {
    return {
      ok: false,
      error: createFailure("unresolved_sheet", `/operations/${operationIndex}/newName`).error,
    };
  }
  virtualSheetNames.add(operation.newName);

  const chartSourceBySheet = virtualChartSourceResolutionsBySheet.get(operation.sourceName);
  if (chartSourceBySheet !== undefined) {
    virtualChartSourceResolutionsBySheet.set(operation.newName, cloneChartSourceResolutionMap(chartSourceBySheet));
  }

  return {
    ok: true,
    primitive: {
      operationIndex,
      kind: operation.kind,
      primitiveName: definition.compilerTargets[0],
      args: createArgs([
        ["sourceName", operation.sourceName],
        ["newName", operation.newName],
      ]),
      effect: definition.effect,
      requiredVerificationKinds: [...definition.requiredVerification],
      calculationRequirement: "not_required",
      resultHandling: "passthrough",
    },
  };
}

function compileChartOperation(
  operationIndex: number,
  operation: Extract<WorkbookOperation, { readonly kind: "create_chart" }>,
  definition: OperationDefinition,
  virtualSheetNames: ReadonlySet<string>,
  virtualChartSourceResolutionsBySheet: ReadonlyMap<string, ReadonlyMap<string, VerifiedWorkbookChartSourceResolution>>,
): CompiledOperationOutcome {
  if (!virtualSheetNames.has(operation.sheet)) {
    return {
      ok: false,
      error: createFailure("unresolved_sheet", `/operations/${operationIndex}/sheet`).error,
    };
  }

  const chartSourceBySheet = virtualChartSourceResolutionsBySheet.get(operation.sheet);
  const resolution = chartSourceBySheet?.get(operation.dataRange);
  if (resolution === undefined) {
    return {
      ok: false,
      error: createFailure("unresolved_chart_source", `/operations/${operationIndex}/dataRange`).error,
    };
  }
  if (resolution.areaCount !== 1) {
    return {
      ok: false,
      error: createFailure("unsupported_noncontiguous_chart", `/operations/${operationIndex}/dataRange`).error,
    };
  }

  return {
    ok: true,
    primitive: {
      operationIndex,
      kind: operation.kind,
      primitiveName: definition.compilerTargets[0],
      args: createArgs([
        ["sheet", operation.sheet],
        ["dataRange", operation.dataRange],
        ["chartType", operation.chartType],
        ["title", operation.title],
      ]),
      effect: definition.effect,
      requiredVerificationKinds: [...definition.requiredVerification],
      calculationRequirement: "not_required",
      resultHandling: "passthrough",
    },
  };
}

function compileOperation(
  operationIndex: number,
  operation: WorkbookOperation,
  definition: OperationDefinition,
  context: VerifiedWorkbookCompileContext,
  virtualSheetNames: Set<string>,
  virtualChartSourceResolutionsBySheet: VirtualChartSourceResolutionsBySheet,
): CompiledOperationOutcome {
  const primitiveName = getCompiledPrimitiveName(definition);
  if (primitiveName === null) {
    return {
      ok: false,
      error: createFailure("contract_definition_drift", `/operationDefinitions/${operationIndex}/compilerTargets`).error,
    };
  }
  if (!context.supportedPrimitiveNames.has(primitiveName)) {
    return {
      ok: false,
      error: createFailure("missing_unsupported_primitive_capability", `/operations/${operationIndex}/kind`).error,
    };
  }
  if (!hasSupportedCapabilities(context.supportedCapabilities, definition)) {
    return {
      ok: false,
      error: createFailure("missing_unsupported_primitive_capability", `/operations/${operationIndex}/kind`).error,
    };
  }

  switch (operation.kind) {
    case "write_values": {
      const sheetFailure = ensureSheetExists(virtualSheetNames, operation.sheet, `/operations/${operationIndex}/sheet`);
      if (sheetFailure !== null) return { ok: false, error: sheetFailure.error };
      return {
        ok: true,
        primitive: {
          operationIndex,
          kind: operation.kind,
          primitiveName,
          args: createArgs([
            ["sheet", operation.sheet],
            ["address", operation.address],
            ["values", cloneValueMatrix(operation.values)],
          ]),
          effect: definition.effect,
          requiredVerificationKinds: [...definition.requiredVerification],
          calculationRequirement: "not_required",
          resultHandling: "passthrough",
        },
      };
    }
    case "write_formula": {
      const sheetFailure = ensureSheetExists(virtualSheetNames, operation.sheet, `/operations/${operationIndex}/sheet`);
      if (sheetFailure !== null) return { ok: false, error: sheetFailure.error };
      return {
        ok: true,
        primitive: {
          operationIndex,
          kind: operation.kind,
          primitiveName,
          args: createArgs([
            ["sheet", operation.sheet],
            ["address", operation.address],
            ["values", [[operation.formula]]],
          ]),
          effect: definition.effect,
          requiredVerificationKinds: [...definition.requiredVerification],
          calculationRequirement: "required_before_verification",
          resultHandling: "passthrough",
        },
      };
    }
    case "fill_formula": {
      const sheetFailure = ensureSheetExists(virtualSheetNames, operation.sheet, `/operations/${operationIndex}/sheet`);
      if (sheetFailure !== null) return { ok: false, error: sheetFailure.error };
      return {
        ok: true,
        primitive: {
          operationIndex,
          kind: operation.kind,
          primitiveName,
          args: createArgs([
            ["sheet", operation.sheet],
            ["sourceAddress", operation.sourceAddress],
            ["targetAddress", operation.targetAddress],
          ]),
          effect: definition.effect,
          requiredVerificationKinds: [...definition.requiredVerification],
          calculationRequirement: "required_before_verification",
          resultHandling: "passthrough",
        },
      };
    }
    case "format_range": {
      const sheetFailure = ensureSheetExists(virtualSheetNames, operation.sheet, `/operations/${operationIndex}/sheet`);
      if (sheetFailure !== null) return { ok: false, error: sheetFailure.error };
      return {
        ok: true,
        primitive: {
          operationIndex,
          kind: operation.kind,
          primitiveName,
          args: createArgs([
            ["sheet", operation.sheet],
            ["address", operation.address],
            ["format", { ...operation.format } satisfies FormatSpec],
          ]),
          effect: definition.effect,
          requiredVerificationKinds: [...definition.requiredVerification],
          calculationRequirement: "not_required",
          resultHandling: "passthrough",
        },
      };
    }
    case "create_table": {
      const sheetFailure = ensureSheetExists(virtualSheetNames, operation.sheet, `/operations/${operationIndex}/sheet`);
      if (sheetFailure !== null) return { ok: false, error: sheetFailure.error };
      return {
        ok: true,
        primitive: {
          operationIndex,
          kind: operation.kind,
          primitiveName,
          args: createArgs([
            ["sheet", operation.sheet],
            ["address", operation.address],
            ["hasHeaders", operation.hasHeaders],
          ]),
          effect: definition.effect,
          requiredVerificationKinds: [...definition.requiredVerification],
          calculationRequirement: "not_required",
          resultHandling: "passthrough",
        },
      };
    }
    case "sort_table":
    case "filter_table":
    case "reset_table_filter": {
      return compileTableOperation(operationIndex, operation, definition, context);
    }
    case "create_chart": {
      return compileChartOperation(
        operationIndex,
        operation,
        definition,
        virtualSheetNames,
        virtualChartSourceResolutionsBySheet,
      );
    }
    case "set_validation_list": {
      const sheetFailure = ensureSheetExists(virtualSheetNames, operation.sheet, `/operations/${operationIndex}/sheet`);
      if (sheetFailure !== null) return { ok: false, error: sheetFailure.error };
      return {
        ok: true,
        primitive: {
          operationIndex,
          kind: operation.kind,
          primitiveName,
          args: createArgs([
            ["sheet", operation.sheet],
            ["address", operation.address],
            ["type", "list"],
            ["source", operation.source],
            ["promptTitle", operation.promptTitle],
            ["promptMessage", operation.promptMessage],
          ]),
          effect: definition.effect,
          requiredVerificationKinds: [...definition.requiredVerification],
          calculationRequirement: "not_required",
          resultHandling: "passthrough",
        },
      };
    }
    case "conditional_format": {
      const sheetFailure = ensureSheetExists(virtualSheetNames, operation.sheet, `/operations/${operationIndex}/sheet`);
      if (sheetFailure !== null) return { ok: false, error: sheetFailure.error };
      return {
        ok: true,
        primitive: {
          operationIndex,
          kind: operation.kind,
          primitiveName,
          args: createArgs([
            ["sheet", operation.sheet],
            ["address", operation.address],
            ["ruleType", operation.ruleType],
            ["operator", operation.operator],
            ["formula1", operation.formula1],
            ["formula2", operation.formula2],
            ["text", operation.text],
            ["fillColor", operation.fillColor],
            ["fontColor", operation.fontColor],
          ]),
          effect: definition.effect,
          requiredVerificationKinds: [...definition.requiredVerification],
          calculationRequirement: "not_required",
          resultHandling: "passthrough",
        },
      };
    }
    case "freeze_panes": {
      const sheetFailure = ensureSheetExists(virtualSheetNames, operation.sheet, `/operations/${operationIndex}/sheet`);
      if (sheetFailure !== null) return { ok: false, error: sheetFailure.error };
      return {
        ok: true,
        primitive: {
          operationIndex,
          kind: operation.kind,
          primitiveName,
          args: createArgs([
            ["sheet", operation.sheet],
            ["rows", operation.rows],
            ["columns", operation.columns],
          ]),
          effect: definition.effect,
          requiredVerificationKinds: [...definition.requiredVerification],
          calculationRequirement: "not_required",
          resultHandling: "passthrough",
        },
      };
    }
    case "create_sheet":
    case "rename_sheet":
    case "copy_sheet":
      return compileSheetChangingOperation(
        operationIndex,
        operation,
        definition,
        virtualSheetNames,
        virtualChartSourceResolutionsBySheet,
      );
    case "add_comment": {
      const sheetFailure = ensureSheetExists(virtualSheetNames, operation.sheet, `/operations/${operationIndex}/sheet`);
      if (sheetFailure !== null) return { ok: false, error: sheetFailure.error };
      return {
        ok: true,
        primitive: {
          operationIndex,
          kind: operation.kind,
          primitiveName,
          args: createArgs([
            ["sheet", operation.sheet],
            ["address", operation.address],
            ["text", operation.text],
          ]),
          effect: definition.effect,
          requiredVerificationKinds: [...definition.requiredVerification],
          calculationRequirement: "not_required",
          resultHandling: "passthrough",
        },
      };
    }
    case "calculate":
      return {
        ok: true,
        primitive: {
          operationIndex,
          kind: operation.kind,
          primitiveName,
          args: createArgs([]),
          effect: definition.effect,
          requiredVerificationKinds: [...definition.requiredVerification],
          calculationRequirement: "required_before_verification",
          resultHandling: "passthrough",
        },
      };
    default: {
      const exhaustiveCheck: never = operation;
      return {
        ok: false,
        error: createFailure("contract_definition_drift", `/operations/${operationIndex}`).error,
      } satisfies CompiledOperationOutcome;
    }
  }
}

function operationRangeAddress(operation: WorkbookOperation): string | undefined {
  switch (operation.kind) {
    case "fill_formula": return operation.targetAddress;
    case "create_chart": return operation.dataRange;
    case "calculate":
    case "create_sheet":
    case "rename_sheet":
    case "copy_sheet":
    case "sort_table":
    case "filter_table":
    case "reset_table_filter":
    case "freeze_panes":
      return undefined;
    default: return operation.address;
  }
}

function operationSheetName(operation: WorkbookOperation): string | undefined {
  switch (operation.kind) {
    case "create_sheet":
    case "calculate":
      return undefined;
    case "rename_sheet": return operation.oldName;
    case "copy_sheet": return operation.sourceName;
    case "sort_table":
    case "filter_table":
    case "reset_table_filter":
      return undefined;
    default: return operation.sheet;
  }
}

function findManifestRange(
  manifest: WorkbookPlanPreconditionManifest | undefined,
  sheet: string,
  address: string,
): WorkbookPlanPreconditionRange | undefined {
  return manifest?.ranges.find((candidate) => candidate.sheet === sheet && candidate.address === address);
}

function isSingleCellAddress(address: string): boolean {
  return /^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$/.test(address);
}

function matrixMatchesManifestRange(
  matrix: readonly (readonly unknown[])[],
  range: WorkbookPlanPreconditionRange | undefined,
): boolean {
  if (range === undefined || matrix.length !== range.rowCount) return false;
  return matrix.every((row) => row.length === range.columnCount);
}

type MandatoryProofBinding = {
  readonly proof?: CompiledWorkbookMandatoryPostcondition;
  readonly complete: boolean;
};

const MAX_EXCEL_ROW = 1_048_576;
const MAX_EXCEL_COLUMN = 16_384;
const MAX_FILL_DIMENSION = 1000;
const MAX_FILL_CELLS = 10_000;
const SUPPORTED_FILL_FUNCTIONS = new Set([
  "ABS", "AND", "AVERAGE", "COUNT", "COUNTA", "IF", "INT", "MAX", "MIN", "MOD", "NOT", "OR", "PRODUCT", "ROUND", "SUM",
]);

type A1Cell = { readonly row: number; readonly column: number };
type A1Rectangle = { readonly first: A1Cell; readonly rows: number; readonly columns: number };

function columnNumber(value: string): number | undefined {
  let result = 0;
  for (const character of value.toUpperCase()) {
    const digit = character.charCodeAt(0) - 64;
    if (digit < 1 || digit > 26) return undefined;
    result = result * 26 + digit;
    if (result > MAX_EXCEL_COLUMN) return undefined;
  }
  return result;
}

function columnLetters(value: number): string | undefined {
  if (!Number.isSafeInteger(value) || value < 1 || value > MAX_EXCEL_COLUMN) return undefined;
  let result = "";
  let remaining = value;
  while (remaining > 0) {
    const digit = (remaining - 1) % 26;
    result = String.fromCharCode(65 + digit) + result;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return result;
}

function parseA1Cell(address: string): A1Cell | undefined {
  const match = /^(?:\$?)([A-Za-z]{1,3})(?:\$?)([1-9][0-9]*)$/.exec(address);
  if (match === null) return undefined;
  const column = columnNumber(match[1]!);
  const row = Number(match[2]);
  if (column === undefined || !Number.isSafeInteger(row) || row < 1 || row > MAX_EXCEL_ROW) return undefined;
  return { row, column };
}

function parseA1Rectangle(address: string): A1Rectangle | undefined {
  const [firstText, lastText = firstText] = address.split(":");
  if (address.split(":").length > 2) return undefined;
  const first = parseA1Cell(firstText!);
  const last = parseA1Cell(lastText!);
  if (first === undefined || last === undefined || last.column < first.column || last.row < first.row) return undefined;
  const rows = last.row - first.row + 1;
  const columns = last.column - first.column + 1;
  if (rows > MAX_FILL_DIMENSION || columns > MAX_FILL_DIMENSION) return undefined;
  return { first, rows, columns };
}

function translatedCellReference(
  columnAbsolute: boolean,
  rowAbsolute: boolean,
  column: number,
  row: number,
  deltaColumn: number,
  deltaRow: number,
): string | undefined {
  const translatedColumn = columnAbsolute ? column : column + deltaColumn;
  const translatedRow = rowAbsolute ? row : row + deltaRow;
  const letters = columnLetters(translatedColumn);
  if (letters === undefined || !Number.isSafeInteger(translatedRow) || translatedRow < 1 || translatedRow > MAX_EXCEL_ROW) return undefined;
  return `${columnAbsolute ? "$" : ""}${letters}${rowAbsolute ? "$" : ""}${translatedRow}`;
}

type FillFormulaTokenKind = "operand" | "a1Reference" | "function" | "operator" | "leftParen" | "rightParen" | "comma" | "colon";

function isSupportedFillFormulaGrammar(formula: string): boolean {
  if (!formula.startsWith("=") || formula.length > 4096) return false;
  const tokens: FillFormulaTokenKind[] = [];
  const operatorValues: string[] = [];
  let index = 1;
  while (index < formula.length) {
    const character = formula[index]!;
    if (/\s/.test(character)) {
      index += 1;
      continue;
    }
    if (character === '"') {
      index += 1;
      let closed = false;
      while (index < formula.length) {
        if (formula[index] !== '"') {
          index += 1;
          continue;
        }
        if (formula[index + 1] === '"') {
          index += 2;
          continue;
        }
        index += 1;
        closed = true;
        break;
      }
      if (!closed) return false;
      tokens.push("operand");
      continue;
    }
    const reference = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)/.exec(formula.slice(index));
    if (reference !== null) {
      const end = index + reference[0].length;
      if (/[A-Za-z0-9_$]/.test(formula[end] ?? "")) return false;
      if (columnNumber(reference[2]!) === undefined || !Number.isSafeInteger(Number(reference[4])) || Number(reference[4]) < 1 || Number(reference[4]) > MAX_EXCEL_ROW) return false;
      tokens.push("a1Reference");
      index = end;
      continue;
    }
    const identifier = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(formula.slice(index));
    if (identifier !== null) {
      const name = identifier[0]!;
      const end = index + name.length;
      const upper = name.toUpperCase();
      if (upper === "TRUE" || upper === "FALSE") {
        tokens.push("operand");
        index = end;
        continue;
      }
      if (formula[end] !== "(" || !SUPPORTED_FILL_FUNCTIONS.has(upper)) return false;
      tokens.push("function");
      index = end;
      continue;
    }
    const number = /^(?:\d+(?:\.\d*)?|\.\d+)/.exec(formula.slice(index));
    if (number !== null) {
      tokens.push("operand");
      index += number[0].length;
      continue;
    }
    const twoCharacterOperator = formula.slice(index, index + 2);
    if (["<>", "<=", ">="].includes(twoCharacterOperator)) {
      tokens.push("operator");
      operatorValues.push(twoCharacterOperator);
      index += 2;
      continue;
    }
    if (["+", "-", "*", "/", "^", "&", "=", "<", ">"].includes(character)) {
      tokens.push("operator");
      operatorValues.push(character);
      index += 1;
      continue;
    }
    if (character === "(") tokens.push("leftParen");
    else if (character === ")") tokens.push("rightParen");
    else if (character === ",") tokens.push("comma");
    else if (character === ":") tokens.push("colon");
    else return false;
    index += 1;
  }

  let expectsOperand = true;
  let previousFunction = false;
  let rangeColonCount = 0;
  const parenthesisFunctions: boolean[] = [];
  for (let tokenIndex = 0; tokenIndex < tokens.length; tokenIndex += 1) {
    const token = tokens[tokenIndex]!;
    switch (token) {
      case "operand":
      case "a1Reference":
        if (!expectsOperand) return false;
        expectsOperand = false;
        previousFunction = false;
        break;
      case "function":
        if (!expectsOperand) return false;
        expectsOperand = true;
        previousFunction = true;
        break;
      case "leftParen":
        if (!expectsOperand) return false;
        parenthesisFunctions.push(previousFunction);
        expectsOperand = true;
        previousFunction = false;
        rangeColonCount = 0;
        break;
      case "rightParen":
        if (expectsOperand || parenthesisFunctions.length === 0) return false;
        parenthesisFunctions.pop();
        expectsOperand = false;
        previousFunction = false;
        break;
      case "comma":
        if (expectsOperand || parenthesisFunctions.at(-1) !== true) return false;
        expectsOperand = true;
        previousFunction = false;
        rangeColonCount = 0;
        break;
      case "colon":
        if (
          expectsOperand ||
          rangeColonCount > 0 ||
          tokens[tokenIndex - 1] !== "a1Reference" ||
          tokens[tokenIndex + 1] !== "a1Reference"
        ) return false;
        expectsOperand = true;
        previousFunction = false;
        rangeColonCount = 1;
        break;
      case "operator": {
        const operator = operatorValues.shift();
        if (expectsOperand) {
          if (operator !== "+" && operator !== "-") return false;
        } else {
          expectsOperand = true;
          rangeColonCount = 0;
        }
        previousFunction = false;
        break;
      }
    }
  }
  return tokens.length > 0 && !expectsOperand && parenthesisFunctions.length === 0;
}

function translateFormula(formula: string, deltaColumn: number, deltaRow: number): string | undefined {
  if (!isSupportedFillFormulaGrammar(formula)) return undefined;
  let result = "";
  let index = 0;
  let parenthesisDepth = 0;
  let referenceCount = 0;
  while (index < formula.length) {
    const character = formula[index]!;
    if (character === '"') {
      const start = index;
      index += 1;
      let closed = false;
      while (index < formula.length) {
        if (formula[index] !== '"') { index += 1; continue; }
        if (formula[index + 1] === '"') { index += 2; continue; }
        index += 1;
        closed = true;
        break;
      }
      if (!closed) return undefined;
      result += formula.slice(start, index);
      continue;
    }
    if (character === "[") return undefined;
    if (character === "]" || character === "{" || character === "}" || character === "#" || character === "@" || character === "!") return undefined;
    if (character === "$" && (index === 0 || !/[A-Za-z]/.test(formula[index + 1] ?? ""))) return undefined;
    if (character === "(") parenthesisDepth += 1;
    if (character === ")") {
      parenthesisDepth -= 1;
      if (parenthesisDepth < 0) return undefined;
    }
    if (character === "R" || character === "r") {
      if (/^R[0-9]+C[0-9]+/i.test(formula.slice(index))) return undefined;
    }
    const reference = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)/.exec(formula.slice(index));
    if (reference !== null && (index === 0 || !/[A-Za-z0-9_$]/.test(formula[index - 1]!))) {
      const end = index + reference[0].length;
      if (formula[end] === "(") {
        result += reference[0];
        index = end;
        continue;
      }
      if (/[A-Za-z0-9_$]/.test(formula[end] ?? "")) return undefined;
      const column = columnNumber(reference[2]!);
      const row = Number(reference[4]);
      if (column === undefined || !Number.isSafeInteger(row) || row < 1 || row > MAX_EXCEL_ROW) return undefined;
      const translated = translatedCellReference(
        reference[1] === "$",
        reference[3] === "$",
        column,
        row,
        deltaColumn,
        deltaRow,
      );
      if (translated === undefined) return undefined;
      result += translated;
      referenceCount += 1;
      index = end;
      continue;
    }
    const identifier = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(formula.slice(index));
    if (identifier !== null) {
      const end = index + identifier[0].length;
      const upper = identifier[0]!.toUpperCase();
      if (upper === "TRUE" || upper === "FALSE") {
        result += identifier[0];
        index = end;
        continue;
      }
      if (formula[end] !== "(" || !SUPPORTED_FILL_FUNCTIONS.has(upper)) return undefined;
      result += identifier[0];
      index = end;
      continue;
    }
    result += character;
    index += 1;
  }
  const trimmed = result.trimEnd();
  if (parenthesisDepth !== 0 || /[+\-*/^&,:]$/.test(trimmed)) return undefined;
  return referenceCount === 0 || result.length > 1 ? result : undefined;
}

function expandFillFormula(
  sourceAddress: string,
  targetAddress: string,
  sourceFormulas: FormulaMatrix,
): FormulaMatrix | undefined {
  const source = parseA1Rectangle(sourceAddress);
  const target = parseA1Rectangle(targetAddress);
  if (source === undefined || target === undefined || source.rows !== 1 || source.columns !== 1 || sourceFormulas.length !== 1 || sourceFormulas[0]?.length !== 1) return undefined;
  const targetLastRow = target.first.row + target.rows - 1;
  const targetLastColumn = target.first.column + target.columns - 1;
  const sourceContained = target.first.row <= source.first.row && target.first.column <= source.first.column &&
    targetLastRow >= source.first.row && targetLastColumn >= source.first.column;
  if (!sourceContained || target.rows > 1 && target.columns > 1 || target.rows * target.columns > MAX_FILL_CELLS) return undefined;
  const sourceFormula = sourceFormulas[0]?.[0];
  if (typeof sourceFormula !== "string") return undefined;
  const formulas: string[][] = [];
  for (let rowIndex = 0; rowIndex < target.rows; rowIndex += 1) {
    const row: string[] = [];
    for (let columnIndex = 0; columnIndex < target.columns; columnIndex += 1) {
      const translated = translateFormula(
        sourceFormula,
        target.first.column + columnIndex - source.first.column,
        target.first.row + rowIndex - source.first.row,
      );
      if (translated === undefined) return undefined;
      row.push(translated);
    }
    formulas.push(row);
  }
  return formulas;
}

function makeMandatoryProof(
  operationIndex: number,
  rule: VerificationRule,
  operation?: WorkbookOperation,
  expectedFormulas?: FormulaMatrix,
): MandatoryProofBinding {
  return {
    complete: true,
    proof: Object.freeze({
      operationIndex,
      kind: rule.kind,
      ...(operation === undefined ? {} : {
        proof: Object.freeze({
          operationIndex,
          operationKind: operation.kind,
          ...(expectedFormulas === undefined ? {} : { expectedFormulas: cloneFormulaMatrix(expectedFormulas) }),
        }),
      }),
    }),
  };
}

function makeInternalProof(
  operationIndex: number,
  operation: WorkbookOperation,
  kind: VerificationRule["kind"],
  expectedFormulas?: FormulaMatrix,
): MandatoryProofBinding {
  return {
    complete: true,
    proof: Object.freeze({
      operationIndex,
      kind,
      proof: Object.freeze({
        operationIndex,
        operationKind: operation.kind,
        ...(expectedFormulas === undefined ? {} : { expectedFormulas: cloneFormulaMatrix(expectedFormulas) }),
      }),
    }),
  };
}

function bindMandatoryPostcondition(
  operationIndex: number,
  operation: WorkbookOperation,
  primitive: CompiledWorkbookPrimitive,
  _authoredRules: readonly VerificationRule[],
  context: VerifiedWorkbookCompileContext,
): MandatoryProofBinding {
  const manifest = context.preconditionManifest;
  const rangeAddress = operationRangeAddress(operation);
  const sheetName = operationSheetName(operation);
  const range = sheetName === undefined || rangeAddress === undefined
    ? undefined
    : findManifestRange(manifest, sheetName, rangeAddress);
  const strictManifest = manifest !== undefined;

  const binding = (() => {
    switch (operation.kind) {
    case "write_values":
      if (strictManifest && (!matrixMatchesManifestRange(operation.values, range) || range?.values === undefined)) return { complete: false };
      return strictManifest
        ? makeMandatoryProof(operationIndex, {
            kind: "range_values_equal",
            sheet: operation.sheet,
            address: operation.address,
            expectedValues: cloneValueMatrix(operation.values),
          }, operation)
        : makeMandatoryProof(operationIndex, {
            kind: "range_values_equal",
            sheet: operation.sheet,
            address: operation.address,
            expectedValues: cloneValueMatrix(operation.values),
          });
    case "write_formula":
      // B1 only accepts a dense exact proof for a single-cell write. A
      // repeated literal cannot prove the requested effect over a multi-cell
      // target, and deriving the dense matrix belongs to a later slice.
      if (!isSingleCellAddress(operation.address)) return { complete: false };
      if (strictManifest && !matrixMatchesManifestRange([[operation.formula]], range)) return { complete: false };
      return strictManifest
        ? makeMandatoryProof(operationIndex, {
            kind: "range_formulas_equal",
            sheet: operation.sheet,
            address: operation.address,
            expectedFormulas: [[operation.formula]],
          }, operation, [[operation.formula]])
        : makeMandatoryProof(operationIndex, {
            kind: "range_formulas_equal",
            sheet: operation.sheet,
            address: operation.address,
            expectedFormulas: [[operation.formula]],
          });
    case "fill_formula": {
      if (!strictManifest) return { complete: false };
      const source = findManifestRange(manifest, operation.sheet, operation.sourceAddress);
      const target = findManifestRange(manifest, operation.sheet, operation.targetAddress);
      if (source === undefined || target === undefined || source.sourceFormulas === undefined) return { complete: false };
      const expectedFormulas = expandFillFormula(operation.sourceAddress, operation.targetAddress, source.sourceFormulas);
      if (expectedFormulas === undefined || !matrixMatchesManifestRange(expectedFormulas, target)) return { complete: false };
      return makeInternalProof(operationIndex, operation, "range_formula_pattern", expectedFormulas);
    }
    case "format_range": {
      const { numberFormat, columnAutofit, rowAutofit } = operation.format;
      if (operation.format.borders === false ||
        (operation.format.borders !== true && (operation.format.borderColor !== undefined || operation.format.borderWeight !== undefined))) {
        return { complete: false };
      }
      const requestedKeys = Object.keys(operation.format);
      const supportedFormatKeys = new Set([
        "bold", "italic", "fontColor", "fillColor", "numberFormat", "borders", "columnAutofit", "rowAutofit",
        "horizontalAlignment", "verticalAlignment", "fontName", "fontSize", "underline", "wrapText", "indentLevel",
        "textOrientation", "borderColor", "borderWeight",
      ]);
      const onlySupportedFields = requestedKeys.every((key) => supportedFormatKeys.has(key));
      if (
        !onlySupportedFields ||
        (!strictManifest && (columnAutofit === true || rowAutofit === true))
      ) {
        return { complete: false };
      }
      if (strictManifest && requestedKeys.some((key) => key !== "numberFormat" && key !== "columnAutofit" && key !== "rowAutofit")) {
        return makeInternalProof(operationIndex, operation, "range_number_format_equal");
      }
      if (strictManifest && (columnAutofit === true || rowAutofit === true)) {
        return makeInternalProof(operationIndex, operation, "range_number_format_equal");
      }
      if (typeof numberFormat !== "string") return { complete: false };
      return strictManifest
        ? makeMandatoryProof(operationIndex, {
            kind: "range_number_format_equal",
            sheet: operation.sheet,
            address: operation.address,
            expectedNumberFormat: numberFormat,
          }, operation)
        : makeMandatoryProof(operationIndex, {
            kind: "range_number_format_equal",
            sheet: operation.sheet,
            address: operation.address,
            expectedNumberFormat: numberFormat,
          });
    }
    case "create_table":
      return strictManifest ? makeInternalProof(operationIndex, operation, "table_exists") : { complete: false };
    case "sort_table":
      // The requested direction is not a complete expected row order. A
      // model-authored table_row_order rule is supplemental and cannot be
      // promoted into the compiler-owned proof for this operation.
      return strictManifest ? makeInternalProof(operationIndex, operation, "table_row_order") : { complete: false };
    case "filter_table":
      if (strictManifest) return makeInternalProof(operationIndex, operation, "table_filter_state");
      return makeMandatoryProof(operationIndex, {
        kind: "table_filter_state",
        tableName: operation.tableName,
        columnIndex: operation.columnIndex,
        state: "filtered",
        ...(operation.values === undefined ? {} : { values: [...operation.values] }),
        ...(operation.criterion1 === undefined ? {} : { criterion1: operation.criterion1 }),
        ...(operation.criterion2 === undefined ? {} : { criterion2: operation.criterion2 }),
        ...(operation.operator === undefined ? {} : { operator: operation.operator }),
      });
    case "reset_table_filter":
      if (strictManifest) return makeInternalProof(operationIndex, operation, "table_filter_state");
      return makeMandatoryProof(operationIndex, {
        kind: "table_filter_state",
        tableName: operation.tableName,
        columnIndex: operation.columnIndex,
        state: "unfiltered",
      });
    case "create_chart":
      // Chart title is not the Office chart identity. Treating a title rule as
      // proof would allow a pre-existing chart to satisfy this operation.
      return strictManifest ? makeInternalProof(operationIndex, operation, "chart_exists") : { complete: false };
    case "set_validation_list":
      // V1 has no compiler-owned readback for these effects; a preserved
      // sentinel would only prove that the effect did not change the value.
      return strictManifest ? makeInternalProof(operationIndex, operation, "sentinel_unchanged") : { complete: false };
    case "conditional_format":
      return strictManifest ? makeInternalProof(operationIndex, operation, "sentinel_unchanged") : { complete: false };
    case "freeze_panes":
      return strictManifest ? makeInternalProof(operationIndex, operation, "sentinel_unchanged") : { complete: false };
    case "create_sheet":
      return strictManifest
        ? makeMandatoryProof(operationIndex, { kind: "sheet_exists", name: operation.name }, operation)
        : makeMandatoryProof(operationIndex, { kind: "sheet_exists", name: operation.name });
    case "rename_sheet":
      return strictManifest
        ? makeMandatoryProof(operationIndex, { kind: "sheet_exists", name: operation.newName }, operation)
        : makeMandatoryProof(operationIndex, { kind: "sheet_exists", name: operation.newName });
    case "copy_sheet":
      return strictManifest
        ? makeMandatoryProof(operationIndex, { kind: "sheet_exists", name: operation.newName }, operation)
        : makeMandatoryProof(operationIndex, { kind: "sheet_exists", name: operation.newName });
    case "add_comment":
      if (strictManifest) {
        const comment = manifestComment(manifest!, operation.sheet, operation.address);
        if (comment === undefined || comment.exists && typeof comment.text !== "string") return { complete: false };
      }
      return strictManifest
        ? makeMandatoryProof(operationIndex, {
            kind: "comment_exists",
            sheet: operation.sheet,
            address: operation.address,
            text: operation.text,
          }, operation)
        : makeMandatoryProof(operationIndex, {
            kind: "comment_exists",
            sheet: operation.sheet,
            address: operation.address,
            text: operation.text,
          });
    case "calculate":
      return strictManifest
        ? makeMandatoryProof(operationIndex, { kind: "calculation_completed" }, operation)
        : makeMandatoryProof(operationIndex, { kind: "calculation_completed" });
      default: {
        const exhaustive: never = operation;
        void exhaustive;
        return { complete: false };
      }
    }
  })();
  if (binding.proof?.proof === undefined) return binding;
  return {
    ...binding,
    proof: {
      ...binding.proof,
      proof: Object.freeze({
        ...binding.proof.proof,
        operationKind: primitive.kind,
        primitiveName: primitive.primitiveName,
        effect: primitive.effect,
      }),
    },
  }
}

function manifestHasSheet(manifest: WorkbookPlanPreconditionManifest, name: string): boolean {
  return manifest.sheets.some((sheet) => sheet.name === name);
}

function manifestSheet(manifest: WorkbookPlanPreconditionManifest, name: string): WorkbookPlanPreconditionSheet | undefined {
  return manifest.sheets.find((sheet) => sheet.name === name);
}

function manifestHasTable(manifest: WorkbookPlanPreconditionManifest, name: string): boolean {
  return manifest.tables.some((table) => table.name === name);
}

function manifestTable(manifest: WorkbookPlanPreconditionManifest, name: string): WorkbookPlanPreconditionTable | undefined {
  return manifest.tables.find((table) => table.name === name);
}

function manifestComment(
  manifest: WorkbookPlanPreconditionManifest,
  sheet: string,
  address: string,
): WorkbookPlanPreconditionComment | undefined {
  return manifest.comments.find((comment) => comment.sheet === sheet && comment.address === address);
}

function validDimension(value: number | undefined): boolean {
  return Number.isSafeInteger(value) && value !== undefined && value > 0;
}

function manifestIsCompleteForPlan(
  manifest: WorkbookPlanPreconditionManifest,
  plan: Extract<WorkbookPlan, { readonly outcome: "execute" }>,
): boolean {
  if (manifest.version !== 1) return false;
  for (const operation of plan.operations) {
    if (operation.kind === "create_sheet") {
      if (!manifestHasSheet(manifest, operation.name) || manifestSheet(manifest, operation.name)?.exists !== false) return false;
    } else if (operation.kind === "rename_sheet") {
      if (!manifestHasSheet(manifest, operation.oldName) || !manifestHasSheet(manifest, operation.newName) ||
        manifestSheet(manifest, operation.oldName)?.exists !== true || manifestSheet(manifest, operation.newName)?.exists !== false) return false;
    } else if (operation.kind === "copy_sheet") {
      if (!manifestHasSheet(manifest, operation.sourceName) || !manifestHasSheet(manifest, operation.newName) ||
        manifestSheet(manifest, operation.sourceName)?.exists !== true || manifestSheet(manifest, operation.newName)?.exists !== false) return false;
    } else if (operation.kind === "sort_table" || operation.kind === "filter_table" || operation.kind === "reset_table_filter") {
      const table = manifestTable(manifest, operation.tableName);
      if (!manifestHasTable(manifest, operation.tableName) || table?.exists !== true ||
        !validDimension(table?.columnCount) || !validDimension(table?.rowCount)) return false;
    } else {
      const sheetName = operationSheetName(operation);
      if (sheetName !== undefined && manifestSheet(manifest, sheetName)?.exists !== true) return false;
      const address = operationRangeAddress(operation);
      if (address !== undefined && sheetName !== undefined) {
        const range = findManifestRange(manifest, sheetName, address);
        if (range === undefined || !validDimension(range.rowCount) || !validDimension(range.columnCount)) return false;
        if (operation.kind === "write_values" && range.values === undefined) return false;
      }
    }
    if (operation.kind === "fill_formula") {
      const source = findManifestRange(manifest, operation.sheet, operation.sourceAddress);
      if (source === undefined || !validDimension(source.rowCount) || !validDimension(source.columnCount)) return false;
    }
    if (operation.kind === "create_chart") {
      const source = manifest.chartSources.find((candidate) => candidate.sheet === operation.sheet && candidate.address === operation.dataRange);
      if (source === undefined || !Number.isSafeInteger(source.areaCount) || source.areaCount < 1 || source.resolvedAddress === undefined) return false;
      const chart = manifest.charts.find((candidate) => candidate.sheet === operation.sheet && candidate.title === operation.title);
      if (chart === undefined || chart.exists !== false) return false;
    }
    if (operation.kind === "add_comment") {
      const comment = manifestComment(manifest, operation.sheet, operation.address);
      if (comment === undefined || typeof comment.exists !== "boolean" || comment.exists && typeof comment.text !== "string") return false;
    }
  }
  return true;
}

export function compileWorkbookPlan(
  plan: WorkbookPlan,
  context: VerifiedWorkbookCompileContext,
): CompileWorkbookPlanResult {
  if (plan.outcome !== "execute") {
    return createFailure("unsupported_outcome", "/outcome");
  }

  const contractValidation = validateContractDefinitions();
  if (isContractValidationFailure(contractValidation)) {
    return {
      ok: false,
      error: contractValidation.error,
    };
  }

  const virtualSheetNames = new Set(context.sheetNames);
  const virtualChartSourceResolutionsBySheet = cloneChartSourceResolutionsBySheet(
    context.chartSourceResolutionsBySheet,
  );
  const primitives: CompiledWorkbookPrimitive[] = [];

  for (let operationIndex = 0; operationIndex < plan.operations.length; operationIndex += 1) {
    const operation = plan.operations[operationIndex];
    const definition = contractValidation.definitionsByKind.get(operation.kind);
    if (definition === undefined) {
      return createFailure("contract_definition_drift", `/operationDefinitions/${operationIndex}/kind`);
    }

    const outcome = compileOperation(
      operationIndex,
      operation,
      definition,
      context,
      virtualSheetNames,
      virtualChartSourceResolutionsBySheet,
    );
    if (isCompiledOperationFailure(outcome)) {
      return {
        ok: false,
        error: outcome.error,
      };
    }

    primitives.push(outcome.primitive);
  }

  const authoredPostconditions = cloneVerificationRules(plan.verification);
  const mandatoryPostconditions: CompiledWorkbookMandatoryPostcondition[] = [];
  let mandatoryProofsComplete = true;
  let firstMissingMandatoryProof = -1;
  for (let operationIndex = 0; operationIndex < plan.operations.length; operationIndex += 1) {
    const binding = bindMandatoryPostcondition(
      operationIndex,
      plan.operations[operationIndex],
      primitives[operationIndex],
      authoredPostconditions,
      context,
    );
    const proofMatchesPrimitive = binding.proof !== undefined &&
      primitives[operationIndex].requiredVerificationKinds.includes(binding.proof.kind);
    if (binding.proof !== undefined && proofMatchesPrimitive) mandatoryPostconditions.push(binding.proof);
    if (!binding.complete || !proofMatchesPrimitive) {
      mandatoryProofsComplete = false;
      if (firstMissingMandatoryProof < 0) firstMissingMandatoryProof = operationIndex;
    }
  }

  // A host-backed full preflight is the executable boundary. Pure compiler
  // callers retain the Slice A inspection shape, but native preparation cannot
  // proceed without both a complete manifest and a proof for every operation.
  if (context.preconditionManifest !== undefined) {
    if (!manifestIsCompleteForPlan(context.preconditionManifest, plan)) {
      return createFailure("incomplete_precondition_manifest", "/operations");
    }
    if (!mandatoryProofsComplete) {
      return createFailure("missing_mandatory_verification", `/operations/${firstMissingMandatoryProof}/kind`);
    }
  }

  const transaction: CompiledWorkbookTransaction = {
    primitives,
    postconditions: authoredPostconditions,
  };
  Object.defineProperties(transaction, {
    mandatoryPostconditions: {
      value: Object.freeze(mandatoryPostconditions),
      enumerable: false,
      writable: false,
      configurable: false,
    },
    mandatoryProofsComplete: {
      value: mandatoryProofsComplete,
      enumerable: false,
      writable: false,
      configurable: false,
    },
  });
  return { ok: true, transaction };
}
