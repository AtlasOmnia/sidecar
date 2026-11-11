import { describe, expect, it } from "vitest";
import { toolSchemas } from "../excelTools";
import { getDirectProviderTools, getDirectSystemPrompt } from "../agent";
import { workbookPlanSchema } from "../generated/workbookPlanSchema";
import {
  classifyWorkbookRequest,
  type TrustedWorkbookRequest,
} from "../workbookPlanRequest";
import {
  accountWorkbookPlanResources,
  WORKBOOK_PLAN_RESOURCE_LIMITS,
  type WorkbookPlanResourceContext,
} from "../workbookPlanResources";
import {
  buildWorkbookPlanPolicyInput,
  preflightWorkbookPlanResources,
  type WorkbookPlanRouteRuntimeContext,
} from "../workbookPlanRoute";
import {
  WorkbookPlanTurnState,
  type WorkbookPlanTurnStatePhase,
} from "../workbookPlanOrchestrator";
import type { WorkbookPlan } from "../generated/workbookPlanSchema";
import { createWorkbookPlanPolicyInputSnapshot, decideWorkbookPlanPolicy } from "../workbookPlanPolicy";

const runtime: WorkbookPlanRouteRuntimeContext = {
  activeSheetName: "Sheet1",
  selection: { sheetName: "Sheet1", address: "B2:C3" },
  compileContext: {
    sheetNames: new Set(["Sheet1", "Sheet2"]),
    tablesByName: new Map(),
    supportedPrimitiveNames: new Set(),
    supportedCapabilities: new Set(),
    chartSourceResolutionsBySheet: new Map(),
  },
};

const trustedRequest = (overrides: Partial<TrustedWorkbookRequest> = {}): TrustedWorkbookRequest => ({
  classification: "explicit_mutation",
  selectionScopedRequest: false,
  ...overrides,
});

const executePlan = (overrides: Partial<Extract<WorkbookPlan, { outcome: "execute" }>> = {}): Extract<WorkbookPlan, { outcome: "execute" }> => ({
  version: 1,
  outcome: "execute",
  summary: "write the requested value",
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }],
  verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["ok"]] }],
  ...overrides,
});

describe("Task 11 slice A route, policy, resource, and ordering contracts", () => {
  it("advertises exactly one generated plan function only for strict rollout true", () => {
    const off = getDirectProviderTools({ workbookPlanRolloutEnabled: false });
    const on = getDirectProviderTools({ workbookPlanRolloutEnabled: true });
    expect(off.map((tool) => tool.function.name)).toEqual(toolSchemas.map((tool) => tool.function.name));
    expect(off.some((tool) => tool.function.name === "excel_execute_plan")).toBe(false);
    expect(on.filter((tool) => tool.function.name === "excel_execute_plan")).toHaveLength(1);
    expect(on.find((tool) => tool.function.name === "excel_execute_plan")?.function.parameters).toBe(workbookPlanSchema);
    expect(getDirectSystemPrompt(true)).toContain("excel_execute_plan");
    expect(getDirectSystemPrompt(false)).not.toContain("excel_execute_plan");
  });

  it("classifies user intent independently of a model-authored plan and requires a real confirmation event", () => {
    expect(classifyWorkbookRequest("Please inspect the selected range")).toMatchObject({
      classification: "read_only",
      selectionScopedRequest: true,
    });
    expect(classifyWorkbookRequest("Create a table from the selected range")).toMatchObject({
      classification: "explicit_structure",
      selectionScopedRequest: true,
    });

    const offSelection = buildWorkbookPlanPolicyInput(
      executePlan({ operations: [{ kind: "write_values", sheet: "Sheet2", address: "A1", values: [["model"]] }] }),
      runtime,
      trustedRequest({ classification: "explicit_mutation" }),
    );
    expect(offSelection.confirmation).toBeUndefined();
    expect(offSelection.requestClassification).toBe("explicit_mutation");
    expect(offSelection.operationSummaries[0]?.target.sheetNames).toEqual(["Sheet2"]);

    const confirmed = buildWorkbookPlanPolicyInput(
      executePlan({ operations: [{ kind: "write_values", sheet: "Sheet2", address: "A1", values: [["model"]] }] }),
      runtime,
      trustedRequest({ confirmation: { eventId: "confirmation-1", confirmed: true, source: "user" } }),
    );
    expect(confirmed.confirmation).toEqual({ eventId: "confirmation-1", confirmed: true });
  });

  it("accounts for bounded operation and verification ranges exactly and rejects unknown scope", () => {
    const plan = executePlan({
      operations: [{ kind: "format_range", sheet: "Sheet1", address: "A1:C3", format: { bold: true } }],
      verification: [
        { kind: "range_number_format_equal", sheet: "Sheet1", address: "B2:C3", expectedNumberFormat: "General" },
        { kind: "sentinel_unchanged", sheet: "Sheet1", address: "D4", expected: null },
      ],
    });
    const accounted = accountWorkbookPlanResources(plan, runtime as WorkbookPlanResourceContext);
    expect(accounted.ok).toBe(true);
    if (accounted.ok) {
      expect(accounted.usage).toMatchObject({
        operationCount: 1,
        cellsWritten: 9,
        cellsRead: 5,
        verificationReads: 2,
      });
      expect(accounted.limits).toEqual(WORKBOOK_PLAN_RESOURCE_LIMITS);
    }

    const unbounded = accountWorkbookPlanResources(
      executePlan({
        operations: [{ kind: "write_values", sheet: "Sheet1", address: "A:A", values: [["bad"]] }],
      }),
      runtime as WorkbookPlanResourceContext,
    );
    expect(unbounded.ok).toBe(false);
  });

  it("rejects dynamic table accounting instead of fabricating a workbook-scale cap", () => {
    const tableRuntime: WorkbookPlanResourceContext = {
      ...runtime,
      compileContext: {
        ...runtime.compileContext,
        tablesByName: new Map([["Orders", { canonicalName: "Orders", columnCount: 4 }]]),
      },
    };
    const accounted = accountWorkbookPlanResources(
      executePlan({
        operations: [{ kind: "sort_table", tableName: "Orders", columnIndex: 0, ascending: true }],
        verification: [{ kind: "table_row_order", tableName: "Orders", expectedOrder: ["1"] }],
      }),
      tableRuntime,
    );
    expect(accounted.ok).toBe(false);
  });

  it("R7-B2 unknown table bounds never mask later operation/verification/payload failures or aggregate limits", () => {
    const empty: WorkbookPlanResourceContext = { compileContext: { sheetNames: new Set(), tablesByName: new Map(), supportedPrimitiveNames: new Set(), supportedCapabilities: new Set(), chartSourceResolutionsBySheet: new Map() } };
    const unknown = { kind: "sort_table", tableName: "Table1", columnIndex: 0, ascending: true } as const;
    const huge = { kind: "format_range", sheet: "Sheet1", address: "A1:XFD1048576", format: { bold: true } } as const;
    const half = { ...huge, address: `A1:A${Math.floor(WORKBOOK_PLAN_RESOURCE_LIMITS.cellsWritten / 2) + 1}` };
    for (const operations of [
      [unknown, huge], [huge, unknown],
      [{ kind: "filter_table", tableName: "Table1", columnIndex: 0, values: ["x"] } as const, huge],
      [{ kind: "reset_table_filter", tableName: "Table1", columnIndex: 0 } as const, huge],
      [unknown, half, half],
    ]) {
      expect(preflightWorkbookPlanResources(executePlan({ operations }), empty)).toEqual({ ok: false, reason: "resource_limit_exceeded" });
    }
    expect(preflightWorkbookPlanResources(executePlan({
      verification: [
        { kind: "table_row_order", tableName: "Table1", expectedOrder: ["1"] },
        { kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1:XFD1048576", expected: null },
      ],
    }), empty)).toEqual({ ok: false, reason: "resource_limit_exceeded" });
    expect(preflightWorkbookPlanResources(executePlan({ operations: [unknown, { kind: "write_values", sheet: "Sheet1", address: "A1:B2", values: [["bad shape"]] }] }), empty)).toEqual({ ok: false, reason: "range_shape_mismatch" });
    expect(preflightWorkbookPlanResources(executePlan({ operations: [unknown, { ...huge, address: "A:A" }] }), empty)).toEqual({ ok: false, reason: "unknown_or_unbounded_range" });
    const smallUnknown = executePlan({ operations: [unknown, { ...huge, address: "A1" }] });
    expect(accountWorkbookPlanResources(smallUnknown, empty)).toEqual({ ok: false, reason: "unknown_table_bounds" });
    expect(preflightWorkbookPlanResources(smallUnknown, empty)).toEqual({ ok: true }); // Only initial empty context may defer.
    expect(preflightWorkbookPlanResources(smallUnknown, runtime)).toEqual({ ok: false, reason: "unknown_table_bounds" });
  });

  it.each([
    ["Sort selected cells", { kind: "sort_table", tableName: "Table1", columnIndex: 0, ascending: true }],
    ["Filter selected cells", { kind: "filter_table", tableName: "Table1", columnIndex: 0, values: ["x"] }],
    ["Reset filter for selected cells", { kind: "reset_table_filter", tableName: "Table1", columnIndex: 0 }],
  ] as const)("R7-B3 %s cannot approve a same-sheet table outside the selection without confirmation", (prompt, operation) => {
    const tableRuntime: WorkbookPlanRouteRuntimeContext = {
      ...runtime, selection: { sheetName: "Sheet1", address: "A1:B3" },
      compileContext: {
        ...runtime.compileContext,
        tablesByName: new Map([["Table1", { canonicalName: "Table1", columnCount: 2, rowCount: 3, sheetName: "Sheet1" }]]),
        preconditionManifest: { version: 1, sheets: [{ name: "Sheet1", exists: true }], tables: [{ name: "Table1", exists: true, columnCount: 2, rowCount: 3, sheetName: "Sheet1", address: "Sheet1!A20:B22" }], ranges: [], chartSources: [], charts: [], comments: [] },
      },
    };
    const request = classifyWorkbookRequest(prompt);
    expect(request).toMatchObject({ classification: "explicit_mutation", selectionScopedRequest: true });
    const plan = executePlan({ operations: [operation] });
    const input = buildWorkbookPlanPolicyInput(plan, tableRuntime, request);
    expect(decideWorkbookPlanPolicy(createWorkbookPlanPolicyInputSnapshot(input))).toEqual({ decision: "requires_confirmation", reasonCode: "selection_scope_mismatch" });
    expect(input.operationSummaries[0].target.selectionRelationship).toBe("unknown");
    const confirmed = buildWorkbookPlanPolicyInput(plan, tableRuntime, { ...request, confirmation: { eventId: "r7b3-confirmation", confirmed: true, source: "user" } });
    expect(decideWorkbookPlanPolicy(createWorkbookPlanPolicyInputSnapshot(confirmed)).decision).toBe("approved");
    const unscoped = buildWorkbookPlanPolicyInput(plan, tableRuntime, classifyWorkbookRequest("Sort Table1"));
    expect(decideWorkbookPlanPolicy(createWorkbookPlanPolicyInputSnapshot(unscoped)).decision).toBe("approved");
  });

  it("R7-B3 non-range mutation siblings are unknown; only calculation is selection-not-applicable", () => {
    for (const operation of [
      { kind: "freeze_panes", sheet: "Sheet1", rows: 1, columns: 0 },
      { kind: "create_sheet", name: "NewSheet" },
      { kind: "rename_sheet", oldName: "Sheet1", newName: "Renamed" },
      { kind: "copy_sheet", sourceName: "Sheet1", newName: "Copied" },
    ] as const) {
      expect(buildWorkbookPlanPolicyInput(executePlan({ operations: [operation] }), runtime, classifyWorkbookRequest("Change selected cells")).operationSummaries[0].target.selectionRelationship).toBe("unknown");
    }
    const calculate = buildWorkbookPlanPolicyInput(executePlan({ operations: [{ kind: "calculate" }] }), runtime, classifyWorkbookRequest("Calculate selected cells"));
    expect(calculate.operationSummaries[0].target.selectionRelationship).toBe("not_applicable");
  });

  it("uses one shared none-inspected-executing-terminal state machine", () => {
    const state = new WorkbookPlanTurnState();
    const phases: WorkbookPlanTurnStatePhase[] = [state.phase];
    expect(state.beginExecution()).toBe(true);
    phases.push(state.phase);
    expect(state.beginInspection()).toBe(false);
    expect(state.beginExecution()).toBe(false);
    state.finishTerminal();
    phases.push(state.phase);
    expect(phases).toEqual(["none", "executing", "terminal"]);
  });
});
