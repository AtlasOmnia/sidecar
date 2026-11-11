type SchemaNode = Record<string, unknown>;

export type ValidationUtf8LimitName =
  | "formulaText"
  | "commentText"
  | "validationListSource";

export type WorkbookPlanValidationConstraints = {
  readonly maxErrors: number;
  readonly maxDepth: number;
  readonly maxWorkUnits: number;
  readonly maxVisitedNodes: number;
  readonly maxUtf8Bytes: Readonly<Record<ValidationUtf8LimitName, number>>;
  readonly aliasMap: Readonly<Record<string, string>>;
  readonly utf8ByteFields: Readonly<
    Record<string, Readonly<Record<string, ValidationUtf8LimitName>>>
  >;
  readonly rectangularMatrices: Readonly<Record<string, readonly string[]>>;
  readonly conditionalFormatRequires: Readonly<Record<string, readonly string[]>>;
};

type CapabilityStatus = "baseline" | "unverified";
type ResourceCostMode = "none" | "range" | "payload" | "count" | "required" | "verification";

const OUTCOMES = ["inspect", "execute", "ask", "refuse"] as const;

const OPERATION_KINDS = [
  "write_values",
  "write_formula",
  "fill_formula",
  "format_range",
  "create_table",
  "sort_table",
  "filter_table",
  "reset_table_filter",
  "create_chart",
  "set_validation_list",
  "conditional_format",
  "freeze_panes",
  "create_sheet",
  "rename_sheet",
  "copy_sheet",
  "add_comment",
  "calculate",
] as const;

const VERIFICATION_KINDS = [
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
] as const;

const INSPECTION_KINDS = ["read_range", "summarize_range_stats"] as const;

const stringSchema = (maxLength: number, minLength = 1): SchemaNode => ({
  type: "string",
  minLength,
  maxLength,
});

const optionalStringSchema = (maxLength: number): SchemaNode => ({
  type: "string",
  maxLength,
});

const ref = (name: string): SchemaNode => ({ $ref: `#/$defs/${name}` });

const closedObject = (
  properties: Record<string, SchemaNode>,
  required: readonly string[],
  oneOf?: readonly SchemaNode[],
): SchemaNode => ({
  type: "object",
  required: [...required],
  additionalProperties: false,
  properties,
  ...(oneOf === undefined ? {} : { oneOf: [...oneOf] }),
});

export const EXCEL_CELL_TEXT_MAX_LENGTH = 32767;

export const isExcelCellTextWithinUtf16Limit = (value: string): boolean =>
  value.length <= EXCEL_CELL_TEXT_MAX_LENGTH;

const scalarValueSchema: SchemaNode = {
  oneOf: [
    { type: "string", maxLength: EXCEL_CELL_TEXT_MAX_LENGTH },
    { type: "number" },
    { type: "boolean" },
    { type: "null" },
  ],
};

const sheetProperty = stringSchema(255);
const addressProperty = stringSchema(256);
const tableNameProperty = stringSchema(255);
const formulaProperty = stringSchema(4096);
const operationKindProperty: SchemaNode = {
  type: "string",
  enum: [...OPERATION_KINDS],
};

const valueRowSchema: SchemaNode = {
  type: "array",
  minItems: 1,
  maxItems: 1000,
  items: ref("ScalarValue"),
};

const valueMatrixSchema: SchemaNode = {
  type: "array",
  minItems: 1,
  maxItems: 1000,
  items: ref("ValueRow"),
};

const formulaRowSchema: SchemaNode = {
  type: "array",
  minItems: 1,
  maxItems: 1000,
  items: { type: "string", maxLength: 4096 },
};

const formulaMatrixSchema: SchemaNode = {
  type: "array",
  minItems: 1,
  maxItems: 1000,
  items: ref("FormulaRow"),
};

const formatSpecSchema = closedObject(
  {
    bold: { type: "boolean" },
    italic: { type: "boolean" },
    fontColor: stringSchema(64),
    fillColor: stringSchema(64),
    numberFormat: stringSchema(128),
    borders: { type: "boolean" },
    columnAutofit: { type: "boolean" },
    rowAutofit: { type: "boolean" },
    horizontalAlignment: {
      type: "string",
      enum: [
        "General",
        "Left",
        "Center",
        "Right",
        "Fill",
        "Justify",
        "CenterAcrossSelection",
        "Distributed",
      ],
    },
    verticalAlignment: stringSchema(32),
    fontName: stringSchema(128),
    fontSize: { type: "number", exclusiveMinimum: 0 },
    underline: { type: "string", enum: ["None", "Single", "Double"] },
    wrapText: { type: "boolean" },
    indentLevel: { type: "integer", minimum: 0 },
    textOrientation: { type: "integer", minimum: -90, maximum: 90 },
    borderColor: stringSchema(64),
    borderWeight: { type: "string", enum: ["Hairline", "Thin", "Medium", "Thick"] },
  },
  [],
);

const operationBranches: readonly SchemaNode[] = [
  closedObject(
    {
      kind: { const: "write_values" },
      sheet: sheetProperty,
      address: addressProperty,
      values: ref("ValueMatrix"),
    },
    ["kind", "sheet", "address", "values"],
  ),
  closedObject(
    {
      kind: { const: "write_formula" },
      sheet: sheetProperty,
      address: addressProperty,
      formula: formulaProperty,
    },
    ["kind", "sheet", "address", "formula"],
  ),
  closedObject(
    {
      kind: { const: "fill_formula" },
      sheet: sheetProperty,
      sourceAddress: addressProperty,
      targetAddress: addressProperty,
    },
    ["kind", "sheet", "sourceAddress", "targetAddress"],
  ),
  closedObject(
    {
      kind: { const: "format_range" },
      sheet: sheetProperty,
      address: addressProperty,
      format: ref("FormatSpec"),
    },
    ["kind", "sheet", "address", "format"],
  ),
  closedObject(
    {
      kind: { const: "create_table" },
      sheet: sheetProperty,
      address: addressProperty,
      hasHeaders: { type: "boolean" },
    },
    ["kind", "sheet", "address", "hasHeaders"],
  ),
  closedObject(
    {
      kind: { const: "sort_table" },
      tableName: tableNameProperty,
      columnIndex: { type: "integer", minimum: 0 },
      ascending: { type: "boolean" },
    },
    ["kind", "tableName", "columnIndex"],
  ),
  closedObject(
    {
      kind: { const: "filter_table" },
      tableName: tableNameProperty,
      columnIndex: { type: "integer", minimum: 0 },
      values: {
        type: "array",
        minItems: 1,
        maxItems: 200,
        items: stringSchema(512),
      },
      criterion1: optionalStringSchema(512),
      criterion2: optionalStringSchema(512),
      operator: { type: "string", enum: ["And", "Or"] },
    },
    ["kind", "tableName", "columnIndex"],
  ),
  closedObject(
    {
      kind: { const: "reset_table_filter" },
      tableName: tableNameProperty,
      columnIndex: { type: "integer", minimum: 0 },
    },
    ["kind", "tableName", "columnIndex"],
  ),
  closedObject(
    {
      kind: { const: "create_chart" },
      sheet: sheetProperty,
      dataRange: addressProperty,
      chartType: stringSchema(128),
      title: stringSchema(255),
    },
    ["kind", "sheet", "dataRange", "chartType", "title"],
  ),
  closedObject(
    {
      kind: { const: "set_validation_list" },
      sheet: sheetProperty,
      address: addressProperty,
      source: stringSchema(4096),
      promptTitle: optionalStringSchema(255),
      promptMessage: optionalStringSchema(1024),
    },
    ["kind", "sheet", "address", "source"],
  ),
  closedObject(
    {
      kind: { const: "conditional_format" },
      sheet: sheetProperty,
      address: addressProperty,
      ruleType: { type: "string", enum: ["cellValue", "containsText"] },
      operator: optionalStringSchema(64),
      formula1: optionalStringSchema(4096),
      formula2: optionalStringSchema(4096),
      text: optionalStringSchema(1024),
      fillColor: stringSchema(64),
      fontColor: stringSchema(64),
    },
    ["kind", "sheet", "address", "ruleType"],
  ),
  closedObject(
    {
      kind: { const: "freeze_panes" },
      sheet: sheetProperty,
      rows: { type: "integer", minimum: 0 },
      columns: { type: "integer", minimum: 0 },
    },
    ["kind", "sheet", "rows", "columns"],
  ),
  closedObject(
    {
      kind: { const: "create_sheet" },
      name: stringSchema(255),
    },
    ["kind", "name"],
  ),
  closedObject(
    {
      kind: { const: "rename_sheet" },
      oldName: stringSchema(255),
      newName: stringSchema(255),
    },
    ["kind", "oldName", "newName"],
  ),
  closedObject(
    {
      kind: { const: "copy_sheet" },
      sourceName: stringSchema(255),
      newName: stringSchema(255),
    },
    ["kind", "sourceName", "newName"],
  ),
  closedObject(
    {
      kind: { const: "add_comment" },
      sheet: sheetProperty,
      address: addressProperty,
      text: stringSchema(1024),
    },
    ["kind", "sheet", "address", "text"],
  ),
  closedObject(
    { kind: { const: "calculate" } },
    ["kind"],
  ),
];

const workbookOperationSchema = closedObject(
  {
    kind: operationKindProperty,
    sheet: sheetProperty,
    address: addressProperty,
    values: ref("ValueMatrix"),
    formula: formulaProperty,
    sourceAddress: addressProperty,
    targetAddress: addressProperty,
    format: ref("FormatSpec"),
    hasHeaders: { type: "boolean" },
    tableName: tableNameProperty,
    columnIndex: { type: "integer", minimum: 0 },
    ascending: { type: "boolean" },
    criterion1: optionalStringSchema(512),
    criterion2: optionalStringSchema(512),
    operator: { type: "string", enum: ["And", "Or"] },
    dataRange: addressProperty,
    chartType: stringSchema(128),
    title: stringSchema(255),
    source: stringSchema(4096),
    promptTitle: optionalStringSchema(255),
    promptMessage: optionalStringSchema(1024),
    ruleType: { type: "string", enum: ["cellValue", "containsText"] },
    formula1: optionalStringSchema(4096),
    formula2: optionalStringSchema(4096),
    text: optionalStringSchema(1024),
    fillColor: stringSchema(64),
    fontColor: stringSchema(64),
    rows: { type: "integer", minimum: 0 },
    columns: { type: "integer", minimum: 0 },
    name: stringSchema(255),
    oldName: stringSchema(255),
    newName: stringSchema(255),
    sourceName: stringSchema(255),
  },
  ["kind"],
  operationBranches,
);

const inspectionBranches: readonly SchemaNode[] = [
  closedObject(
    {
      kind: { const: "read_range" },
      sheet: sheetProperty,
      address: addressProperty,
    },
    ["kind", "sheet", "address"],
  ),
  closedObject(
    {
      kind: { const: "summarize_range_stats" },
      sheet: sheetProperty,
      address: addressProperty,
    },
    ["kind", "sheet", "address"],
  ),
];

const inspectionOperationSchema = closedObject(
  {
    kind: { type: "string", enum: [...INSPECTION_KINDS] },
    sheet: sheetProperty,
    address: addressProperty,
  },
  ["kind"],
  inspectionBranches,
);

const verificationBranches: readonly SchemaNode[] = [
  closedObject(
    {
      kind: { const: "range_values_equal" },
      sheet: sheetProperty,
      address: addressProperty,
      expectedValues: ref("ValueMatrix"),
    },
    ["kind", "sheet", "address", "expectedValues"],
  ),
  closedObject(
    {
      kind: { const: "range_formulas_equal" },
      sheet: sheetProperty,
      address: addressProperty,
      expectedFormulas: ref("FormulaMatrix"),
    },
    ["kind", "sheet", "address", "expectedFormulas"],
  ),
  closedObject(
    {
      kind: { const: "range_formula_pattern" },
      sheet: sheetProperty,
      address: addressProperty,
      pattern: stringSchema(4096),
    },
    ["kind", "sheet", "address", "pattern"],
  ),
  closedObject(
    {
      kind: { const: "range_number_format_equal" },
      sheet: sheetProperty,
      address: addressProperty,
      expectedNumberFormat: stringSchema(128),
    },
    ["kind", "sheet", "address", "expectedNumberFormat"],
  ),
  closedObject(
    {
      kind: { const: "table_exists" },
      tableName: tableNameProperty,
    },
    ["kind", "tableName"],
  ),
  closedObject(
    {
      kind: { const: "table_row_order" },
      tableName: tableNameProperty,
      expectedOrder: {
        type: "array",
        minItems: 1,
        maxItems: 1000,
        items: stringSchema(255),
      },
    },
    ["kind", "tableName", "expectedOrder"],
  ),
  closedObject(
    {
      kind: { const: "table_filter_state" },
      tableName: tableNameProperty,
      columnIndex: { type: "integer", minimum: 0 },
      state: { type: "string", enum: ["filtered", "unfiltered"] },
      values: {
        type: "array",
        maxItems: 200,
        items: stringSchema(512),
      },
      criterion1: optionalStringSchema(512),
      criterion2: optionalStringSchema(512),
      operator: { type: "string", enum: ["And", "Or"] },
    },
    ["kind", "tableName", "columnIndex", "state"],
  ),
  closedObject(
    {
      kind: { const: "sheet_exists" },
      name: stringSchema(255),
    },
    ["kind", "name"],
  ),
  closedObject(
    {
      kind: { const: "chart_exists" },
      sheet: sheetProperty,
      title: stringSchema(255),
    },
    ["kind", "sheet", "title"],
  ),
  closedObject(
    {
      kind: { const: "comment_exists" },
      sheet: sheetProperty,
      address: addressProperty,
      text: stringSchema(1024),
    },
    ["kind", "sheet", "address", "text"],
  ),
  closedObject(
    {
      kind: { const: "sentinel_unchanged" },
      sheet: sheetProperty,
      address: addressProperty,
      expected: ref("ScalarValue"),
    },
    ["kind", "sheet", "address", "expected"],
  ),
  closedObject(
    { kind: { const: "calculation_completed" } },
    ["kind"],
  ),
];

const verificationRuleSchema = closedObject(
  {
    kind: { type: "string", enum: [...VERIFICATION_KINDS] },
    sheet: sheetProperty,
    address: addressProperty,
    expectedValues: ref("ValueMatrix"),
    expectedFormulas: ref("FormulaMatrix"),
    pattern: stringSchema(4096),
    expectedNumberFormat: stringSchema(128),
    tableName: tableNameProperty,
    expectedOrder: {
      type: "array",
      maxItems: 1000,
      items: stringSchema(255),
    },
    columnIndex: { type: "integer", minimum: 0 },
    state: { type: "string", enum: ["filtered", "unfiltered"] },
    values: {
      type: "array",
      maxItems: 200,
      items: stringSchema(512),
    },
    criterion1: optionalStringSchema(512),
    criterion2: optionalStringSchema(512),
    operator: { type: "string", enum: ["And", "Or"] },
    name: stringSchema(255),
    title: stringSchema(255),
    text: stringSchema(1024),
    expected: ref("ScalarValue"),
  },
  ["kind"],
  verificationBranches,
);

const workbookPlanSchema = {
  $schema: "https://json-schema.org/draft/2020-12/schema",
  title: "WorkbookPlan V1",
  description: "Static model-facing contract for bounded, non-destructive Excel plans.",
  type: "object",
  required: ["version", "outcome"],
  additionalProperties: false,
  properties: {
    version: {
      type: "integer",
      const: 1,
      description: "Plan schema version. Only version 1 is valid.",
    },
    outcome: {
      type: "string",
      enum: [...OUTCOMES],
      description: "The plan outcome discriminator.",
    },
    reason: stringSchema(1024),
    summary: stringSchema(2048),
    message: stringSchema(2048),
    reads: {
      type: "array",
      minItems: 1,
      maxItems: 50,
      items: ref("InspectionOperation"),
    },
    operations: {
      type: "array",
      minItems: 1,
      maxItems: 30,
      items: ref("WorkbookOperation"),
    },
    verification: {
      type: "array",
      minItems: 1,
      maxItems: 30,
      items: ref("VerificationRule"),
    },
  },
  oneOf: [
    closedObject(
      {
        version: { type: "integer", const: 1 },
        outcome: { const: "inspect" },
        reason: stringSchema(1024),
        reads: {
          type: "array",
          minItems: 1,
          maxItems: 50,
          items: ref("InspectionOperation"),
        },
      },
      ["version", "outcome", "reason", "reads"],
    ),
    closedObject(
      {
        version: { type: "integer", const: 1 },
        outcome: { const: "execute" },
        summary: stringSchema(2048),
        operations: {
          type: "array",
          minItems: 1,
          maxItems: 30,
          items: ref("WorkbookOperation"),
        },
        verification: {
          type: "array",
          minItems: 1,
          maxItems: 30,
          items: ref("VerificationRule"),
        },
      },
      ["version", "outcome", "summary", "operations", "verification"],
    ),
    closedObject(
      {
        version: { type: "integer", const: 1 },
        outcome: { type: "string", enum: ["ask", "refuse"] },
        message: stringSchema(2048),
      },
      ["version", "outcome", "message"],
    ),
  ],
  $defs: {
    ScalarValue: scalarValueSchema,
    ValueRow: valueRowSchema,
    ValueMatrix: valueMatrixSchema,
    FormulaRow: formulaRowSchema,
    FormulaMatrix: formulaMatrixSchema,
    FormatSpec: formatSpecSchema,
    InspectionOperation: inspectionOperationSchema,
    WorkbookOperation: workbookOperationSchema,
    VerificationRule: verificationRuleSchema,
  },
} as const;

const baselineCapabilities = [
  { capability: "excel_workbook_access", status: "baseline" as CapabilityStatus },
  { capability: "workbook_plan_v1_operation_audit", status: "unverified" as CapabilityStatus },
] as const;

const resourceCost = (
  cellsRead: ResourceCostMode,
  cellsWritten: ResourceCostMode,
  formulaBytes: ResourceCostMode,
  commentBytes: ResourceCostMode,
  sheetsCreated: ResourceCostMode,
  chartsCreated: ResourceCostMode,
  tablesCreated: ResourceCostMode,
  verificationReads: ResourceCostMode,
) => ({
  cellsRead,
  cellsWritten,
  formulaBytes,
  commentBytes,
  sheetsCreated,
  chartsCreated,
  tablesCreated,
  verificationReads,
});

const operationDefinition = (
  kind: (typeof OPERATION_KINDS)[number],
  effect: "write" | "format" | "structure" | "filter" | "calculate",
  compilerTargets: readonly string[],
  requiredVerification: readonly (typeof VERIFICATION_KINDS)[number][],
  cost: ReturnType<typeof resourceCost>,
) => ({
  kind,
  effect,
  compilerTargets,
  requiredVerification,
  capabilityRequirements: baselineCapabilities,
  resourceCost: cost,
});

const operationDefinitions = [
  operationDefinition("write_values", "write", ["write_range"], ["range_values_equal"], resourceCost("verification", "range", "none", "none", "none", "none", "none", "required")),
  operationDefinition("write_formula", "write", ["write_range"], ["range_formulas_equal"], resourceCost("verification", "range", "payload", "none", "none", "none", "none", "required")),
  operationDefinition("fill_formula", "write", ["autofill_formula"], ["range_formula_pattern"], resourceCost("verification", "range", "payload", "none", "none", "none", "none", "required")),
  operationDefinition("format_range", "format", ["format_range"], ["range_number_format_equal"], resourceCost("verification", "range", "none", "none", "none", "none", "none", "required")),
  operationDefinition("create_table", "structure", ["insert_table"], ["table_exists"], resourceCost("range", "none", "none", "none", "none", "none", "count", "required")),
  operationDefinition("sort_table", "write", ["sort_table"], ["table_row_order"], resourceCost("range", "range", "none", "none", "none", "none", "none", "required")),
  operationDefinition("filter_table", "filter", ["filter_table"], ["table_filter_state"], resourceCost("range", "range", "none", "none", "none", "none", "none", "required")),
  operationDefinition("reset_table_filter", "filter", ["filter_table"], ["table_filter_state"], resourceCost("range", "range", "none", "none", "none", "none", "none", "required")),
  operationDefinition("create_chart", "structure", ["create_chart"], ["chart_exists"], resourceCost("range", "none", "none", "none", "none", "count", "none", "required")),
  operationDefinition("set_validation_list", "write", ["set_data_validation"], ["sentinel_unchanged"], resourceCost("verification", "range", "none", "none", "none", "none", "none", "required")),
  operationDefinition("conditional_format", "format", ["conditional_format_range"], ["sentinel_unchanged"], resourceCost("verification", "range", "none", "none", "none", "none", "none", "required")),
  operationDefinition("freeze_panes", "format", ["freeze_panes"], ["sentinel_unchanged"], resourceCost("none", "none", "none", "none", "none", "none", "none", "required")),
  operationDefinition("create_sheet", "structure", ["create_sheet"], ["sheet_exists"], resourceCost("none", "none", "none", "none", "count", "none", "none", "required")),
  operationDefinition("rename_sheet", "structure", ["rename_sheet"], ["sheet_exists"], resourceCost("none", "none", "none", "none", "none", "none", "none", "required")),
  operationDefinition("copy_sheet", "structure", ["copy_sheet"], ["sheet_exists"], resourceCost("none", "none", "none", "none", "count", "none", "none", "required")),
  operationDefinition("add_comment", "write", ["add_comment"], ["comment_exists"], resourceCost("verification", "none", "none", "payload", "none", "none", "none", "required")),
  operationDefinition("calculate", "calculate", ["run_calculation"], ["calculation_completed"], resourceCost("none", "none", "none", "none", "none", "none", "none", "required")),
] as const;

export const workbookPlanContract = {
  version: 1,
  schema: workbookPlanSchema,
  operationKinds: OPERATION_KINDS,
  verificationKinds: VERIFICATION_KINDS,
  inspectionKinds: INSPECTION_KINDS,
  operationDefinitions,
  // Contract-derived practical execution caps. These are deliberately bounded
  // to the incumbent primitive read ceiling rather than permitting workbook-
  // scale work from a compact model payload.
  resourceLimits: {
    operationCount: 30,
    cellsRead: 5000,
    cellsWritten: 5000,
    formulaUtf8Bytes: 64 * 1024,
    csvPayloadUtf8Bytes: 64 * 1024,
    commentUtf8Bytes: 16 * 1024,
    sheetsCreated: 10,
    chartsCreated: 10,
    tablesCreated: 10,
    verificationReads: 30,
  },
  validationConstraints: {
    maxErrors: 50,
    maxDepth: 32,
    maxWorkUnits: 4096,
    maxVisitedNodes: 1024,
    maxUtf8Bytes: {
      formulaText: 4096,
      commentText: 1024,
      validationListSource: 4096,
    },
    aliasMap: {
      sheet_name: "sheet",
      range: "address",
      fill_color: "fillColor",
      font_color: "fontColor",
    },
    utf8ByteFields: {
      write_formula: { formula: "formulaText" },
      range_formulas_equal: { expectedFormulas: "formulaText" },
      range_formula_pattern: { pattern: "formulaText" },
      add_comment: { text: "commentText" },
      comment_exists: { text: "commentText" },
      set_validation_list: { source: "validationListSource" },
    },
    rectangularMatrices: {
      write_values: ["values"],
      range_values_equal: ["expectedValues"],
      range_formulas_equal: ["expectedFormulas"],
    },
    conditionalFormatRequires: {
      containsText: ["text"],
      cellValue: ["formula1"],
    },
  } satisfies WorkbookPlanValidationConstraints,
} as const;

export default workbookPlanContract;
