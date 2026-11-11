import { describe, expect, it } from "vitest";
import { excelTools, executeExcelTool, executeExcelToolParsed, getExcelToolDefinition } from "../excelTools";
import {
  compileWorkbookPlan,
  type CompiledWorkbookPrimitive,
  type VerifiedWorkbookCompileContext,
  type VerifiedWorkbookTable,
} from "../workbookPlanCompiler";
import type { VerificationRule, WorkbookOperation, WorkbookPlan } from "../generated/workbookPlanSchema";
import { workbookPlanContract } from "../workbookPlanContract";

const definitionByKind = new Map(
  workbookPlanContract.operationDefinitions.map((definition) => [definition.kind, definition] as const),
);

const supportedPrimitiveNames = workbookPlanContract.operationDefinitions.flatMap(({ compilerTargets }) => compilerTargets);
const supportedCapabilities = ["excel_workbook_access", "workbook_plan_v1_operation_audit"] as const;

const defaultSheetNames = [
  "Data",
  "Formula",
  "Format",
  "TableSheet",
  "ChartSheet",
  "ValidationSheet",
  "ConditionalSheet",
  "FreezeSheet",
  "RenameSource",
  "CopySource",
  "Comments",
] as const;

const trustedTable: VerifiedWorkbookTable = {
  canonicalName: "TrustedTable",
  columnCount: 3,
};

const createContext = (overrides: {
  readonly sheetNames?: readonly string[];
  readonly tablesByName?: readonly VerifiedWorkbookTable[];
  readonly chartSourceResolutionsBySheet?: readonly {
    readonly sheet: string;
    readonly resolutions: readonly {
      readonly dataRange: string;
      readonly areaCount: number;
    }[];
  }[];
  readonly supportedPrimitiveNames?: readonly string[];
  readonly supportedCapabilities?: readonly string[];
} = {}): VerifiedWorkbookCompileContext => ({
  sheetNames: new Set(overrides.sheetNames ?? defaultSheetNames),
  tablesByName: new Map(
    (overrides.tablesByName ?? [trustedTable]).map((table) => [table.canonicalName, table] as const),
  ),
  supportedPrimitiveNames: new Set(overrides.supportedPrimitiveNames ?? supportedPrimitiveNames),
  supportedCapabilities: new Set(overrides.supportedCapabilities ?? supportedCapabilities),
  chartSourceResolutionsBySheet: new Map(
    (overrides.chartSourceResolutionsBySheet ?? [
      {
        sheet: "ChartSheet",
        resolutions: [{ dataRange: "A1:C10", areaCount: 1 }],
      },
    ]).map(({ sheet, resolutions }) => [
      sheet,
      new Map(resolutions.map((resolution) => [resolution.dataRange, { areaCount: resolution.areaCount }] as const)),
    ] as const),
  ),
});

const executePlan = (
  operations: readonly WorkbookOperation[],
  verification: readonly VerificationRule[] = [],
): Extract<WorkbookPlan, { readonly outcome: "execute" }> => ({
  version: 1,
  outcome: "execute",
  summary: "test plan",
  operations,
  verification,
});

const expectedDefinition = (kind: WorkbookOperation["kind"]): (typeof workbookPlanContract.operationDefinitions)[number] => {
  const definition = definitionByKind.get(kind);
  if (!definition) {
    throw new Error(`Missing definition for ${kind}`);
  }
  return definition;
};

const expectedPrimitive = (
  operationIndex: number,
  kind: WorkbookOperation["kind"],
  primitiveName: string,
  args: Record<string, unknown>,
  calculationRequirement: "required_before_verification" | "not_required",
): CompiledWorkbookPrimitive => {
  const definition = expectedDefinition(kind);
  return {
    operationIndex,
    kind,
    primitiveName,
    args,
    effect: definition.effect,
    requiredVerificationKinds: [...definition.requiredVerification],
    calculationRequirement,
    resultHandling: "passthrough",
  };
};

describe("compileWorkbookPlan", () => {
  it("compiles all 17 operation kinds in order with exact primitive mappings and passthrough postconditions", () => {
    const operations: WorkbookOperation[] = [
      {
        kind: "write_values",
        sheet: "Data",
        address: "A1:B2",
        values: [
          [1, 2],
          [3, 4],
        ],
      },
      {
        kind: "write_formula",
        sheet: "Formula",
        address: "C1",
        formula: "=SUM(A1:A2)",
      },
      {
        kind: "fill_formula",
        sheet: "Formula",
        sourceAddress: "C1",
        targetAddress: "C1:C5",
      },
      {
        kind: "format_range",
        sheet: "Format",
        address: "A1:B2",
        format: {
          bold: true,
          numberFormat: "0.00",
        },
      },
      {
        kind: "create_table",
        sheet: "TableSheet",
        address: "A1:C10",
        hasHeaders: true,
      },
      {
        kind: "sort_table",
        tableName: "TrustedTable",
        columnIndex: 2,
      },
      {
        kind: "filter_table",
        tableName: "TrustedTable",
        columnIndex: 1,
        values: ["Active"],
      },
      {
        kind: "reset_table_filter",
        tableName: "TrustedTable",
        columnIndex: 1,
      },
      {
        kind: "create_chart",
        sheet: "ChartSheet",
        dataRange: "A1:C10",
        chartType: "ColumnClustered",
        title: "Revenue",
      },
      {
        kind: "set_validation_list",
        sheet: "ValidationSheet",
        address: "D1:D5",
        source: "Yes,No",
        promptTitle: "Pick one",
      },
      {
        kind: "conditional_format",
        sheet: "ConditionalSheet",
        address: "E1:E5",
        ruleType: "cellValue",
        operator: "GreaterThan",
        formula1: "10",
        fillColor: "#FFFF00",
      },
      {
        kind: "freeze_panes",
        sheet: "FreezeSheet",
        rows: 1,
        columns: 2,
      },
      {
        kind: "create_sheet",
        name: "CreatedSheet",
      },
      {
        kind: "rename_sheet",
        oldName: "RenameSource",
        newName: "RenamedSheet",
      },
      {
        kind: "copy_sheet",
        sourceName: "CopySource",
        newName: "CopiedSheet",
      },
      {
        kind: "add_comment",
        sheet: "Comments",
        address: "F1",
        text: "Needs review",
      },
      {
        kind: "calculate",
      },
    ];

    const verification: VerificationRule[] = [
      {
        kind: "range_values_equal",
        sheet: "Data",
        address: "A1:B2",
        expectedValues: [
          [1, 2],
          [3, 4],
        ],
      },
      {
        kind: "comment_exists",
        sheet: "Comments",
        address: "F1",
        text: "Needs review",
      },
    ];

    const plan = executePlan(operations, verification);
    const context = createContext();
    const planSnapshot = structuredClone(plan);
    const contextSnapshot = structuredClone(context);

    const result = compileWorkbookPlan(plan, context);
    expect(result.ok).toBe(true);
    if (!result.ok) throw new Error("Expected success");

    expect(result.transaction.primitives).toEqual([
      expectedPrimitive(0, "write_values", "write_range", {
        sheet: "Data",
        address: "A1:B2",
        values: [
          [1, 2],
          [3, 4],
        ],
      }, "not_required"),
      expectedPrimitive(1, "write_formula", "write_range", {
        sheet: "Formula",
        address: "C1",
        values: [["=SUM(A1:A2)"]],
      }, "required_before_verification"),
      expectedPrimitive(2, "fill_formula", "autofill_formula", {
        sheet: "Formula",
        sourceAddress: "C1",
        targetAddress: "C1:C5",
      }, "required_before_verification"),
      expectedPrimitive(3, "format_range", "format_range", {
        sheet: "Format",
        address: "A1:B2",
        format: {
          bold: true,
          numberFormat: "0.00",
        },
      }, "not_required"),
      expectedPrimitive(4, "create_table", "insert_table", {
        sheet: "TableSheet",
        address: "A1:C10",
        hasHeaders: true,
      }, "not_required"),
      expectedPrimitive(5, "sort_table", "sort_table", {
        tableName: "TrustedTable",
        columnIndex: 2,
        ascending: true,
      }, "not_required"),
      expectedPrimitive(6, "filter_table", "filter_table", {
        tableName: "TrustedTable",
        columnIndex: 1,
        values: ["Active"],
      }, "not_required"),
      expectedPrimitive(7, "reset_table_filter", "filter_table", {
        tableName: "TrustedTable",
        columnIndex: 1,
      }, "not_required"),
      expectedPrimitive(8, "create_chart", "create_chart", {
        sheet: "ChartSheet",
        dataRange: "A1:C10",
        chartType: "ColumnClustered",
        title: "Revenue",
      }, "not_required"),
      expectedPrimitive(9, "set_validation_list", "set_data_validation", {
        sheet: "ValidationSheet",
        address: "D1:D5",
        type: "list",
        source: "Yes,No",
        promptTitle: "Pick one",
      }, "not_required"),
      expectedPrimitive(10, "conditional_format", "conditional_format_range", {
        sheet: "ConditionalSheet",
        address: "E1:E5",
        ruleType: "cellValue",
        operator: "GreaterThan",
        formula1: "10",
        fillColor: "#FFFF00",
      }, "not_required"),
      expectedPrimitive(11, "freeze_panes", "freeze_panes", {
        sheet: "FreezeSheet",
        rows: 1,
        columns: 2,
      }, "not_required"),
      expectedPrimitive(12, "create_sheet", "create_sheet", {
        name: "CreatedSheet",
      }, "not_required"),
      expectedPrimitive(13, "rename_sheet", "rename_sheet", {
        oldName: "RenameSource",
        newName: "RenamedSheet",
      }, "not_required"),
      expectedPrimitive(14, "copy_sheet", "copy_sheet", {
        sourceName: "CopySource",
        newName: "CopiedSheet",
      }, "not_required"),
      expectedPrimitive(15, "add_comment", "add_comment", {
        sheet: "Comments",
        address: "F1",
        text: "Needs review",
      }, "not_required"),
      expectedPrimitive(16, "calculate", "run_calculation", {}, "required_before_verification"),
    ]);

    expect(result.transaction.postconditions).toEqual(verification);
    expect(result.transaction.primitives).not.toBe(plan.operations);
    expect(result.transaction.postconditions).not.toBe(plan.verification);
    expect(plan).toEqual(planSnapshot);
    expect(context).toEqual(contextSnapshot);
    expect(result.transaction.primitives[0].args).not.toBe(plan.operations[0]);
    expect(result.transaction.primitives[5].args).toEqual({
      tableName: "TrustedTable",
      columnIndex: 2,
      ascending: true,
    });
    expect(Object.keys(result.transaction.primitives[6].args).sort()).toEqual([
      "columnIndex",
      "tableName",
      "values",
    ]);
    expect(Object.keys(result.transaction.primitives[7].args).sort()).toEqual([
      "columnIndex",
      "tableName",
    ]);
  });

  it("resolves every compiler target through the stable excelTools seam and preserves executor references", () => {
    for (const definition of workbookPlanContract.operationDefinitions) {
      for (const target of definition.compilerTargets) {
        const seamDefinition = getExcelToolDefinition(target);
        const directDefinition = excelTools.find((entry) => entry.schema.function.name === target);
        expect(seamDefinition).toBeDefined();
        expect(directDefinition).toBeDefined();
        expect(seamDefinition).toBe(directDefinition);
        expect(seamDefinition?.execute).toBe(directDefinition?.execute);
      }
    }
  });

  it("returns undefined and preserves the unknown-tool path for prototype-sensitive names", async () => {
    const names = ["toString", "constructor", "__proto__", "missing_excel_tool"] as const;

    for (const name of names) {
      expect(getExcelToolDefinition(name)).toBeUndefined();
      await expect(executeExcelTool(name, "{}")).resolves.toContain(`Unknown tool: ${name}`);
      await expect(executeExcelToolParsed(name, {})).resolves.toContain(`Unknown tool: ${name}`);
    }
  });

  it("rejects non-execute plan outcomes with a single unsupported_outcome failure", () => {
    const plans: WorkbookPlan[] = [
      {
        version: 1,
        outcome: "inspect",
        reason: "inspect",
        reads: [],
      },
      {
        version: 1,
        outcome: "ask",
        message: "ask",
      },
      {
        version: 1,
        outcome: "refuse",
        message: "refuse",
      },
    ];

    for (const plan of plans) {
      const result = compileWorkbookPlan(plan, createContext());
      expect(result).toEqual({
        ok: false,
        error: {
          code: "unsupported_outcome",
          path: "/outcome",
          message: "Only execute plans can be compiled.",
        },
      });
    }
  });

  it("fails closed when a primitive is not supported and emits no partial transaction", () => {
    const plan = executePlan(
      [
        {
          kind: "write_values",
          sheet: "Data",
          address: "A1",
          values: [[1]],
        },
      ],
      [],
    );

    const result = compileWorkbookPlan(
      plan,
      createContext({ supportedPrimitiveNames: [] }),
    );

    expect(result).toEqual({
      ok: false,
      error: {
        code: "missing_unsupported_primitive_capability",
        path: "/operations/0/kind",
        message: "Workbook plan cannot be compiled with the current verified context.",
      },
    });
  });

  it("preserves virtual sheet state across create rename and copy operations", () => {
    const plan = executePlan([
      {
        kind: "create_sheet",
        name: "CreatedSheet",
      },
      {
        kind: "add_comment",
        sheet: "CreatedSheet",
        address: "A1",
        text: "created",
      },
      {
        kind: "rename_sheet",
        oldName: "RenameSource",
        newName: "RenamedSheet",
      },
      {
        kind: "write_values",
        sheet: "RenamedSheet",
        address: "A1",
        values: [["renamed"]],
      },
      {
        kind: "copy_sheet",
        sourceName: "CopySource",
        newName: "CopiedSheet",
      },
      {
        kind: "add_comment",
        sheet: "CopiedSheet",
        address: "B1",
        text: "copied",
      },
    ]);

    const result = compileWorkbookPlan(plan, createContext());
    expect(result.ok).toBe(true);
    if (!result.ok) throw new Error("Expected success");

    expect(result.transaction.primitives.map(({ primitiveName, args }) => ({ primitiveName, args }))).toEqual([
      { primitiveName: "create_sheet", args: { name: "CreatedSheet" } },
      { primitiveName: "add_comment", args: { sheet: "CreatedSheet", address: "A1", text: "created" } },
      { primitiveName: "rename_sheet", args: { oldName: "RenameSource", newName: "RenamedSheet" } },
      { primitiveName: "write_range", args: { sheet: "RenamedSheet", address: "A1", values: [["renamed"]] } },
      { primitiveName: "copy_sheet", args: { sourceName: "CopySource", newName: "CopiedSheet" } },
      { primitiveName: "add_comment", args: { sheet: "CopiedSheet", address: "B1", text: "copied" } },
    ]);
  });

  it("rejects create rename and copy sheet collisions and unresolved sources", () => {
    const cases = [
      {
        plan: executePlan([
          {
            kind: "create_sheet",
            name: "Data",
          },
        ]),
        error: {
          code: "unresolved_sheet" as const,
          path: "/operations/0/name",
          message: "Workbook plan references an unavailable sheet.",
        },
      },
      {
        plan: executePlan([
          {
            kind: "rename_sheet",
            oldName: "MissingSheet",
            newName: "RenamedSheet",
          },
        ]),
        error: {
          code: "unresolved_sheet" as const,
          path: "/operations/0/oldName",
          message: "Workbook plan references an unavailable sheet.",
        },
      },
      {
        plan: executePlan([
          {
            kind: "copy_sheet",
            sourceName: "CopySource",
            newName: "Data",
          },
        ]),
        error: {
          code: "unresolved_sheet" as const,
          path: "/operations/0/newName",
          message: "Workbook plan references an unavailable sheet.",
        },
      },
    ] as const;

    for (const { plan, error } of cases) {
      const result = compileWorkbookPlan(plan, createContext());
      expect(result).toEqual({ ok: false, error });
    }
  });

  it("resolves trusted tables exactly and rejects missing or out-of-range columns", () => {
    const success = compileWorkbookPlan(
      executePlan([
        {
          kind: "sort_table",
          tableName: "TrustedTable",
          columnIndex: 2,
        },
      ]),
      createContext({ tablesByName: [trustedTable] }),
    );

    expect(success).toEqual({
      ok: true,
      transaction: {
        primitives: [
          expectedPrimitive(0, "sort_table", "sort_table", {
            tableName: "TrustedTable",
            columnIndex: 2,
            ascending: true,
          }, "not_required"),
        ],
        postconditions: [],
      },
    });

    const failures = [
      {
        plan: executePlan([
          {
            kind: "sort_table",
            tableName: "MissingTable",
            columnIndex: 0,
          },
        ]),
        error: {
          code: "unresolved_table" as const,
          path: "/operations/0/tableName",
          message: "Workbook plan references an unavailable table.",
        },
      },
      {
        plan: executePlan([
          {
            kind: "filter_table",
            tableName: "TrustedTable",
            columnIndex: 3,
            values: ["Active"],
          },
        ]),
        error: {
          code: "out_of_range_trusted_table_column" as const,
          path: "/operations/0/columnIndex",
          message: "Workbook plan references an invalid trusted table column.",
        },
      },
    ] as const;

    for (const { plan, error } of failures) {
      const result = compileWorkbookPlan(
        plan,
        createContext({ tablesByName: [trustedTable] }),
      );
      expect(result).toEqual({ ok: false, error });
    }
  });

  it("resolves chart sources exactly and preserves the requested range unchanged", () => {
    const successPlan = executePlan([
      {
        kind: "create_chart",
        sheet: "ChartSheet",
        dataRange: "A1:C10",
        chartType: "Line",
        title: "Trend",
      },
    ]);

    const success = compileWorkbookPlan(successPlan, createContext());
    expect(success).toEqual({
      ok: true,
      transaction: {
        primitives: [
          expectedPrimitive(0, "create_chart", "create_chart", {
            sheet: "ChartSheet",
            dataRange: "A1:C10",
            chartType: "Line",
            title: "Trend",
          }, "not_required"),
        ],
        postconditions: [],
      },
    });

    const failures = [
      {
        plan: executePlan([
          {
            kind: "create_chart",
            sheet: "ChartSheet",
            dataRange: "Z1:Z10",
            chartType: "Line",
            title: "Missing",
          },
        ]),
        context: createContext(),
        error: {
          code: "unresolved_chart_source" as const,
          path: "/operations/0/dataRange",
          message: "Workbook plan references an unavailable chart source.",
        },
      },
      {
        plan: executePlan([
          {
            kind: "create_chart",
            sheet: "ChartSheet",
            dataRange: "A1:C10",
            chartType: "Line",
            title: "Split",
          },
        ]),
        context: createContext({
          chartSourceResolutionsBySheet: [
            {
              sheet: "ChartSheet",
              resolutions: [{ dataRange: "A1:C10", areaCount: 2 }],
            },
          ],
        }),
        error: {
          code: "unsupported_noncontiguous_chart" as const,
          path: "/operations/0/dataRange",
          message: "Workbook plan references a noncontiguous chart source.",
        },
      },
    ] as const;

    for (const { plan, context, error } of failures) {
      const result = compileWorkbookPlan(plan, context ?? createContext());
      expect(result).toEqual({ ok: false, error });
    }
  });

  it("preserves the trusted chart source across rename operations", () => {
    const plan = executePlan([
      {
        kind: "rename_sheet",
        oldName: "ChartSheet",
        newName: "RenamedChartSheet",
      },
      {
        kind: "create_chart",
        sheet: "RenamedChartSheet",
        dataRange: "A1:C10",
        chartType: "Line",
        title: "Trend",
      },
    ]);
    const context = createContext();
    const planSnapshot = structuredClone(plan);
    const contextSnapshot = structuredClone(context);

    const result = compileWorkbookPlan(plan, context);
    expect(result).toEqual({
      ok: true,
      transaction: {
        primitives: [
          expectedPrimitive(0, "rename_sheet", "rename_sheet", {
            oldName: "ChartSheet",
            newName: "RenamedChartSheet",
          }, "not_required"),
          expectedPrimitive(1, "create_chart", "create_chart", {
            sheet: "RenamedChartSheet",
            dataRange: "A1:C10",
            chartType: "Line",
            title: "Trend",
          }, "not_required"),
        ],
        postconditions: [],
      },
    });
    expect(plan).toEqual(planSnapshot);
    expect(context).toEqual(contextSnapshot);
  });

  it("compiles an exact no-op rename_sheet and keeps the trusted chart source available on the same sheet", () => {
    const plan = executePlan([
      {
        kind: "rename_sheet",
        oldName: "ChartSheet",
        newName: "ChartSheet",
      },
      {
        kind: "create_chart",
        sheet: "ChartSheet",
        dataRange: "A1:C10",
        chartType: "Line",
        title: "No-op",
      },
    ]);
    const context = createContext();
    const planSnapshot = structuredClone(plan);
    const contextSnapshot = structuredClone(context);

    const result = compileWorkbookPlan(plan, context);
    expect(result).toEqual({
      ok: true,
      transaction: {
        primitives: [
          expectedPrimitive(0, "rename_sheet", "rename_sheet", {
            oldName: "ChartSheet",
            newName: "ChartSheet",
          }, "not_required"),
          expectedPrimitive(1, "create_chart", "create_chart", {
            sheet: "ChartSheet",
            dataRange: "A1:C10",
            chartType: "Line",
            title: "No-op",
          }, "not_required"),
        ],
        postconditions: [],
      },
    });
    expect(plan).toEqual(planSnapshot);
    expect(context).toEqual(contextSnapshot);
  });

  it("rejects a same-name rename_sheet when the source sheet is unresolved", () => {
    const result = compileWorkbookPlan(
      executePlan([
        {
          kind: "rename_sheet",
          oldName: "MissingSheet",
          newName: "MissingSheet",
        },
      ]),
      createContext(),
    );

    expect(result).toEqual({
      ok: false,
      error: {
        code: "unresolved_sheet",
        path: "/operations/0/oldName",
        message: "Workbook plan references an unavailable sheet.",
      },
    });
  });

  it("treats case-different rename_sheet names as a real rename and moves the trusted chart source", () => {
    const plan = executePlan([
      {
        kind: "rename_sheet",
        oldName: "ChartSheet",
        newName: "chartsheet",
      },
      {
        kind: "create_chart",
        sheet: "chartsheet",
        dataRange: "A1:C10",
        chartType: "Line",
        title: "Case",
      },
    ]);
    const result = compileWorkbookPlan(plan, createContext());

    expect(result).toEqual({
      ok: true,
      transaction: {
        primitives: [
          expectedPrimitive(0, "rename_sheet", "rename_sheet", {
            oldName: "ChartSheet",
            newName: "chartsheet",
          }, "not_required"),
          expectedPrimitive(1, "create_chart", "create_chart", {
            sheet: "chartsheet",
            dataRange: "A1:C10",
            chartType: "Line",
            title: "Case",
          }, "not_required"),
        ],
        postconditions: [],
      },
    });
  });

  it("clones the trusted chart source across copy operations", () => {
    const plan = executePlan([
      {
        kind: "copy_sheet",
        sourceName: "CopySource",
        newName: "CopiedChartSheet",
      },
      {
        kind: "create_chart",
        sheet: "CopiedChartSheet",
        dataRange: "A1:C10",
        chartType: "Column",
        title: "Copy",
      },
    ]);
    const context = createContext({
      chartSourceResolutionsBySheet: [
        {
          sheet: "CopySource",
          resolutions: [{ dataRange: "A1:C10", areaCount: 1 }],
        },
      ],
    });

    const result = compileWorkbookPlan(plan, context);
    expect(result).toEqual({
      ok: true,
      transaction: {
        primitives: [
          expectedPrimitive(0, "copy_sheet", "copy_sheet", {
            sourceName: "CopySource",
            newName: "CopiedChartSheet",
          }, "not_required"),
          expectedPrimitive(1, "create_chart", "create_chart", {
            sheet: "CopiedChartSheet",
            dataRange: "A1:C10",
            chartType: "Column",
            title: "Copy",
          }, "not_required"),
        ],
        postconditions: [],
      },
    });
  });

  it("rejects untrusted chart ranges after rename and copy", () => {
    const cases = [
      {
        plan: executePlan([
          {
            kind: "rename_sheet",
            oldName: "ChartSheet",
            newName: "RenamedChartSheet",
          },
          {
            kind: "create_chart",
            sheet: "RenamedChartSheet",
            dataRange: "Z1:Z10",
            chartType: "Line",
            title: "Missing",
          },
        ]),
        context: createContext(),
        error: {
          code: "unresolved_chart_source" as const,
          path: "/operations/1/dataRange",
          message: "Workbook plan references an unavailable chart source.",
        },
      },
      {
        plan: executePlan([
          {
            kind: "copy_sheet",
            sourceName: "CopySource",
            newName: "CopiedChartSheet",
          },
          {
            kind: "create_chart",
            sheet: "CopiedChartSheet",
            dataRange: "Z1:Z10",
            chartType: "Line",
            title: "Missing",
          },
        ]),
        context: createContext({
          chartSourceResolutionsBySheet: [
            {
              sheet: "CopySource",
              resolutions: [{ dataRange: "A1:C10", areaCount: 1 }],
            },
          ],
        }),
        error: {
          code: "unresolved_chart_source" as const,
          path: "/operations/1/dataRange",
          message: "Workbook plan references an unavailable chart source.",
        },
      },
    ] as const;

    for (const { plan, context, error } of cases) {
      const result = compileWorkbookPlan(plan, context);
      expect(result).toEqual({ ok: false, error });
    }
  });

  it("moves chart-source mappings on rename and preserves the original mapping on copy", () => {
    const renameResult = compileWorkbookPlan(
      executePlan([
        {
          kind: "rename_sheet",
          oldName: "ChartSheet",
          newName: "RenamedChartSheet",
        },
        {
          kind: "create_sheet",
          name: "ChartSheet",
        },
        {
          kind: "create_chart",
          sheet: "ChartSheet",
          dataRange: "A1:C10",
          chartType: "Line",
          title: "Legacy",
        },
      ]),
      createContext(),
    );

    expect(renameResult).toEqual({
      ok: false,
      error: {
        code: "unresolved_chart_source",
        path: "/operations/2/dataRange",
        message: "Workbook plan references an unavailable chart source.",
      },
    });

    const copyResult = compileWorkbookPlan(
      executePlan([
        {
          kind: "copy_sheet",
          sourceName: "CopySource",
          newName: "CopiedChartSheet",
        },
        {
          kind: "create_chart",
          sheet: "CopySource",
          dataRange: "A1:C10",
          chartType: "Line",
          title: "Source still trusted",
        },
      ]),
      createContext({
        chartSourceResolutionsBySheet: [
          {
            sheet: "CopySource",
            resolutions: [{ dataRange: "A1:C10", areaCount: 1 }],
          },
        ],
      }),
    );

    expect(copyResult).toEqual({
      ok: true,
      transaction: {
        primitives: [
          expectedPrimitive(0, "copy_sheet", "copy_sheet", {
            sourceName: "CopySource",
            newName: "CopiedChartSheet",
          }, "not_required"),
          expectedPrimitive(1, "create_chart", "create_chart", {
            sheet: "CopySource",
            dataRange: "A1:C10",
            chartType: "Line",
            title: "Source still trusted",
          }, "not_required"),
        ],
        postconditions: [],
      },
    });
  });

  it("attaches contract metadata deterministically and marks formula and calculation operations as requiring calculation before verification", () => {
    const plan = executePlan([
      {
        kind: "write_values",
        sheet: "Data",
        address: "A1",
        values: [[1]],
      },
      {
        kind: "write_formula",
        sheet: "Formula",
        address: "A1",
        formula: "=A1+1",
      },
      {
        kind: "fill_formula",
        sheet: "Formula",
        sourceAddress: "A1",
        targetAddress: "A1:A3",
      },
      {
        kind: "calculate",
      },
    ]);

    const result = compileWorkbookPlan(plan, createContext());
    expect(result.ok).toBe(true);
    if (!result.ok) throw new Error("Expected success");

    const expectations = [
      { kind: "write_values", calculationRequirement: "not_required" },
      { kind: "write_formula", calculationRequirement: "required_before_verification" },
      { kind: "fill_formula", calculationRequirement: "required_before_verification" },
      { kind: "calculate", calculationRequirement: "required_before_verification" },
    ] as const;

    for (let index = 0; index < expectations.length; index += 1) {
      const expected = expectations[index];
      const primitive = result.transaction.primitives[index];
      const definition = expectedDefinition(expected.kind);
      expect(primitive.effect).toBe(definition.effect);
      expect(primitive.requiredVerificationKinds).toEqual(definition.requiredVerification);
      expect(primitive.calculationRequirement).toBe(expected.calculationRequirement);
    }
    expect(result.transaction.primitives[0].calculationRequirement).toBe("not_required");
  });

  it("fails closed when the contract target metadata drifts", () => {
    const definition = workbookPlanContract.operationDefinitions[0];
    const originalTargets = [...definition.compilerTargets];
    try {
      const driftedTargets = [...originalTargets];
      driftedTargets[0] = "__drifted_target__";
      expect(Reflect.set(definition, "compilerTargets", driftedTargets)).toBe(true);
      const result = compileWorkbookPlan(
        executePlan([
          {
            kind: "write_values",
            sheet: "Data",
            address: "A1",
            values: [[1]],
          },
        ]),
        createContext(),
      );

      expect(result).toEqual({
        ok: false,
        error: {
          code: "contract_definition_drift",
          path: "/operationDefinitions/0/compilerTargets",
          message: "Workbook plan compiler metadata is inconsistent.",
        },
      });
    } finally {
      expect(Reflect.set(definition, "compilerTargets", originalTargets)).toBe(true);
    }
  });

  it("does not serialize synthetic canaries in failure payloads", () => {
    const plan = executePlan([
      {
        kind: "write_formula",
        sheet: "__SHEET_CANARY__",
        address: "A1",
        formula: "=__FORMULA_CANARY__",
      },
      {
        kind: "add_comment",
        sheet: "Comments",
        address: "A1",
        text: "__COMMENT_CANARY__",
      },
      {
        kind: "create_sheet",
        name: "__NAME_CANARY__",
      },
    ]);

    const result = compileWorkbookPlan(plan, createContext({ supportedPrimitiveNames: [] }));
    expect(result).toEqual({
      ok: false,
      error: {
        code: "missing_unsupported_primitive_capability",
        path: "/operations/0/kind",
        message: "Workbook plan cannot be compiled with the current verified context.",
      },
    });

    expect(result.ok).toBe(false);
    if (result.ok) throw new Error("Expected failure");
    const serialized = JSON.stringify(result.error);
    expect(serialized).not.toContain("__SHEET_CANARY__");
    expect(serialized).not.toContain("__FORMULA_CANARY__");
    expect(serialized).not.toContain("__COMMENT_CANARY__");
    expect(serialized).not.toContain("__NAME_CANARY__");
  });
});
