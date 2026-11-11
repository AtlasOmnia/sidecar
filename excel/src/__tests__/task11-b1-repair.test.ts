import { describe, expect, it } from "vitest";
import {
  compileWorkbookPlan,
  type VerifiedWorkbookCompileContext,
  type WorkbookPlanPreconditionManifest,
} from "../workbookPlanCompiler";
import {
  BrowserIdempotencyStore,
  MemoryIdempotencyStorage,
  createTrustedOperationBoundary,
  executeAcceptedWorkbookPlan,
  prepareWorkbookPlanExecution,
} from "../workbookPlanExecutor";
import { createNativeWorkbookPlanAdapter } from "../workbookPlanOfficeAdapter";
import type { WorkbookOperation, WorkbookPlan, VerificationRule } from "../generated/workbookPlanSchema";
import type { WorkbookPlanPolicyInput } from "../workbookPlanPolicy";

const IDS = {
  sessionId: "11111111-1111-4111-8111-111111111111",
  documentId: "22222222-2222-4222-8222-222222222222",
  turnId: "33333333-3333-4333-8333-333333333333",
};
const boundary = createTrustedOperationBoundary(IDS);

const manifest = (rangeOverrides: Partial<WorkbookPlanPreconditionManifest["ranges"][number]>[] = []): WorkbookPlanPreconditionManifest => ({
  version: 1,
  sheets: [{ name: "Sheet1", exists: true }],
  tables: [],
  ranges: ["A1", "B1"].map((address, index) => ({
    sheet: "Sheet1",
    address,
    rowCount: 1,
    columnCount: 1,
    sentinel: index === 0 ? "before" : null,
    values: [[index === 0 ? "before" : null]],
    numberFormat: "General",
    ...rangeOverrides[index],
  })),
  chartSources: [],
  charts: [],
  comments: [],
});

const context = (preconditionManifest?: WorkbookPlanPreconditionManifest): VerifiedWorkbookCompileContext => ({
  sheetNames: new Set(["Sheet1"]),
  tablesByName: new Map(),
  supportedPrimitiveNames: new Set([
    "write_range",
    "format_range",
    "autofill_formula",
    "set_data_validation",
    "conditional_format_range",
    "insert_table",
  ]),
  supportedCapabilities: new Set(["excel_workbook_access", "workbook_plan_v1_operation_audit"]),
  chartSourceResolutionsBySheet: new Map(),
  ...(preconditionManifest === undefined ? {} : { preconditionManifest }),
});

const manifestForRange = (
  address: string,
  rowCount: number,
  columnCount: number,
): WorkbookPlanPreconditionManifest => {
  const base = manifest();
  return {
    ...base,
    ranges: [
      ...base.ranges,
      { sheet: "Sheet1", address, rowCount, columnCount, sentinel: null, values: Array.from({ length: rowCount }, () => Array(columnCount).fill(null)), numberFormat: "General" },
    ],
  };
};

const plan = (operations: readonly WorkbookOperation[], verification: readonly VerificationRule[] = [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "B1", expected: null }]): WorkbookPlan => ({
  version: 1,
  outcome: "execute",
  summary: "B1 proof",
  operations,
  verification,
});

const write = (address: string, value: string): WorkbookOperation => ({
  kind: "write_values",
  sheet: "Sheet1",
  address,
  values: [[value]],
});

const policy = (effects: readonly ("write" | "format")[]): WorkbookPlanPolicyInput => ({
  requestClassification: "explicit_mutation",
  activeSheetName: "Sheet1",
  selectionScopedRequest: false,
  operationSummaries: effects.map((effect, operationIndex) => ({
    operationIndex,
    effect,
    target: { operationIndex, sheetNames: ["Sheet1"], selectionRelationship: "inside" },
  })),
  capabilitySummaries: ["excel_workbook_access", "workbook_plan_v1_operation_audit"],
  resourceUsage: {
    operationCount: effects.length,
    cellsRead: 0,
    cellsWritten: effects.length,
    formulaUtf8Bytes: 0,
    csvPayloadUtf8Bytes: 0,
    commentUtf8Bytes: 0,
    sheetsCreated: 0,
    chartsCreated: 0,
    tablesCreated: 0,
    verificationReads: 1,
  },
  resourceLimits: {
    operationCount: 100,
    cellsRead: 100,
    cellsWritten: 100,
    formulaUtf8Bytes: 10000,
    csvPayloadUtf8Bytes: 10000,
    commentUtf8Bytes: 10000,
    sheetsCreated: 10,
    chartsCreated: 10,
    tablesCreated: 10,
    verificationReads: 100,
  },
});

function officeContext(options: { readonly staleA1?: boolean; readonly staleB1?: boolean; readonly missingExistenceApi?: boolean } = {}) {
  const syncs: number[] = [];
  let mutations = 0;
  const makeRange = (address: string): Record<string, unknown> => ({
    address,
    rowCount: address === "B1" && options.staleB1 === true ? 2 : 1,
    columnCount: 1,
    values: [[address === "A1" ? (options.staleA1 === true ? "changed" : "before") : null]],
    formulas: [[""]],
    numberFormat: [["General"]],
    load: () => undefined,
    format: { set: () => { mutations += 1; } },
    dataValidation: {},
    conditionalFormats: { add: () => undefined },
  });
  const sheet = {
    name: "Sheet1",
    getRange: (address: string) => makeRange(address),
    charts: { getItemOrNullObject: () => ({ isNullObject: true, load: () => undefined }) },
    comments: { getItemByCell: () => ({ isNullObject: true, content: undefined, load: () => undefined }) },
  };
  const nullTable = { isNullObject: true, load: () => undefined };
  const context = {
    workbook: {
      worksheets: {
        ...(options.missingExistenceApi ? {} : { getItemOrNullObject: () => ({ isNullObject: false, load: () => undefined }) }),
        getItem: () => sheet,
      },
      tables: { getItemOrNullObject: () => nullTable, getItem: () => nullTable },
    },
    sync: async () => { syncs.push(1); },
  };
  return { context, syncs, get mutations() { return mutations; } };
}

async function acceptedB1(options: { readonly staleB1?: boolean } = {}) {
  const preconditionManifest = manifest(options.staleB1 ? [{}, { rowCount: 1 }] : []);
  const operations: WorkbookOperation[] = [
    write("A1", "after"),
    { kind: "format_range", sheet: "Sheet1", address: "B1", format: { numberFormat: "0.00" } },
  ];
  const result = await prepareWorkbookPlanExecution({
    plan: plan(operations),
    compileContext: context(preconditionManifest),
    policyInput: policy(["write", "format"]),
    trustedOperationBoundary: boundary,
    uuidSource: (() => {
      const values = [IDS.turnId, "44444444-4444-4444-8444-444444444444", "55555555-5555-4555-8555-555555555555"];
      let index = 0;
      return () => values[index++]!;
    })(),
  });
  expect(result.accepted).toBe(true);
  if (!result.accepted) throw new Error("expected accepted B1 transaction");
  return result.transaction;
}

describe("Task 11 B1 compiler-bound proofs", () => {
  it("binds mandatory proofs to each primitive target, independent of model rules", () => {
    const result = compileWorkbookPlan(
      plan([write("A1", "x"), write("B1", "y")], [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "B1", expected: null }]),
      context(manifest()),
    );
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.transaction.mandatoryPostconditions?.map((proof) => [proof.operationIndex, proof.kind])).toEqual([
      [0, "range_values_equal"],
      [1, "range_values_equal"],
    ]);
    expect(result.transaction.postconditions).toEqual([{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "B1", expected: null }]);
  });

  it("accepts a deterministic create-table proof when the host manifest is complete", () => {
    const baseManifest = manifest();
    const result = compileWorkbookPlan(
      plan([{ kind: "create_table", sheet: "Sheet1", address: "A1:B2", hasHeaders: true }]),
      context({
        ...baseManifest,
        ranges: [...baseManifest.ranges, { sheet: "Sheet1", address: "A1:B2", rowCount: 2, columnCount: 2, sentinel: null, numberFormat: "General" }],
      }),
    );
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.transaction.mandatoryPostconditions?.[0]?.proof?.operationKind).toBe("create_table");
  });

  it("does not promote a model-authored row-order rule into the compiler proof", () => {
    const baseManifest = manifest();
    const result = compileWorkbookPlan(
      plan(
        [{ kind: "sort_table", tableName: "Table1", columnIndex: 0, ascending: true }],
        [{ kind: "table_row_order", tableName: "Table1", expectedOrder: ["a", "b"] }],
      ),
      {
        ...context(baseManifest),
        supportedPrimitiveNames: new Set(["sort_table"]),
        tablesByName: new Map([["Table1", { canonicalName: "Table1", columnCount: 1, rowCount: 2 }]]),
        preconditionManifest: { ...baseManifest, tables: [{ name: "Table1", exists: true, columnCount: 1, rowCount: 2 }] },
      },
    );
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.transaction.mandatoryPostconditions?.[0]?.proof?.operationKind).toBe("sort_table");
  });

  it("retains an unrelated model rule only as an additional check", async () => {
    const transactionResult = await prepareWorkbookPlanExecution({
      plan: plan([write("A1", "x")], [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "B1", expected: null }]),
      compileContext: context(manifest()),
      policyInput: policy(["write"]),
      trustedOperationBoundary: boundary,
      uuidSource: (() => {
        const values = [IDS.turnId, "44444444-4444-4444-8444-444444444444"];
        let index = 0;
        return () => values[index++]!;
      })(),
    });
    expect(transactionResult.accepted).toBe(true);
    if (!transactionResult.accepted) return;
    const adapter = {
      async executeFirstGroup(gate: Parameters<NonNullable<ReturnType<typeof createNativeWorkbookPlanAdapter>["executeFirstGroup"]>>[0]) {
        return gate.executeAfterFresh("fresh", async () => ({ status: "committed" as const }));
      },
      async executeGroup() { return { status: "committed" as const }; },
      async reconcileFailedGroup(group: { operations: readonly { operationIndex: number }[] }) { return group.operations.map(() => "not_committed" as const); },
      verifyMandatoryPostconditions: async () => true,
      verifyPlan: async (rules: readonly unknown[]) => {
        expect(rules.map((rule) => (rule as { readonly kind: string }).kind)).toEqual(["range_values_equal", "sentinel_unchanged"]);
        return { ok: true, rules: [], counts: { verified: 0, mismatched: 0, unsupported: 0, error: 0 }, firstFailure: null } as never;
      },
    };
    const report = await executeAcceptedWorkbookPlan(transactionResult.transaction, boundary, adapter, new BrowserIdempotencyStore(new MemoryIdempotencyStorage()));
    expect(report.outcome).toBe("committed");
  });
});

describe("Task 11 B1 whole-plan TOCTOU gate", () => {
  it("fails before mutation when a later effect group's dimension is stale", async () => {
    const transaction = await acceptedB1({ staleB1: true });
    const fake = officeContext({ staleB1: true });
    const previous = (globalThis as Record<string, unknown>).Excel;
    (globalThis as Record<string, unknown>).Excel = { run: async (fn: (ctx: typeof fake.context) => Promise<unknown>) => fn(fake.context) };
    try {
      const report = await executeAcceptedWorkbookPlan(transaction, boundary, createNativeWorkbookPlanAdapter(), new BrowserIdempotencyStore(new MemoryIdempotencyStorage()));
      expect(report.outcome).toBe("stale");
      expect(fake.mutations).toBe(0);
      expect(fake.syncs).toEqual([1]);
    } finally {
      if (previous === undefined) delete (globalThis as Record<string, unknown>).Excel;
      else (globalThis as Record<string, unknown>).Excel = previous;
    }
  });

  it("fails before mutation when a captured sentinel changes", async () => {
    const transaction = await acceptedB1();
    const fake = officeContext({ staleA1: true });
    const previous = (globalThis as Record<string, unknown>).Excel;
    (globalThis as Record<string, unknown>).Excel = { run: async (fn: (ctx: typeof fake.context) => Promise<unknown>) => fn(fake.context) };
    try {
      const report = await executeAcceptedWorkbookPlan(transaction, boundary, createNativeWorkbookPlanAdapter(), new BrowserIdempotencyStore(new MemoryIdempotencyStorage()));
      expect(report.outcome).toBe("stale");
      expect(fake.mutations).toBe(0);
    } finally {
      if (previous === undefined) delete (globalThis as Record<string, unknown>).Excel;
      else (globalThis as Record<string, unknown>).Excel = previous;
    }
  });

  it("fails closed when the host cannot provide required existence readback", async () => {
    const transaction = await acceptedB1();
    const fake = officeContext({ missingExistenceApi: true });
    const previous = (globalThis as Record<string, unknown>).Excel;
    (globalThis as Record<string, unknown>).Excel = { run: async (fn: (ctx: typeof fake.context) => Promise<unknown>) => fn(fake.context) };
    try {
      const report = await executeAcceptedWorkbookPlan(transaction, boundary, createNativeWorkbookPlanAdapter(), new BrowserIdempotencyStore(new MemoryIdempotencyStorage()));
      expect(report.outcome).toBe("partial");
      expect(fake.mutations).toBe(0);
      expect(fake.syncs).toEqual([1]);
    } finally {
      if (previous === undefined) delete (globalThis as Record<string, unknown>).Excel;
      else (globalThis as Record<string, unknown>).Excel = previous;
    }
  });
});

describe("Task 11 B1 mandatory-proof semantic invariant", () => {
  const compile = (operation: WorkbookOperation, preconditionManifest = manifest()) =>
    compileWorkbookPlan(plan([operation]), context(preconditionManifest));
  const compileWithoutManifest = (operation: WorkbookOperation) =>
    compileWorkbookPlan(plan([operation]), context());

  it("keeps pure no-manifest compilation compatible while omitting unprovable proofs", () => {
    const operations: WorkbookOperation[] = [
      { kind: "fill_formula", sheet: "Sheet1", sourceAddress: "A1", targetAddress: "A1:A2" },
      { kind: "set_validation_list", sheet: "Sheet1", address: "A1", source: "one,two" },
      {
        kind: "conditional_format",
        sheet: "Sheet1",
        address: "A1",
        ruleType: "cellValue",
        operator: "greaterThan",
        formula1: "0",
        fillColor: "#ff0000",
      },
      { kind: "format_range", sheet: "Sheet1", address: "A1", format: { bold: true } },
    ];

    for (const operation of operations) {
      const result = compileWithoutManifest(operation);
      expect(result.ok).toBe(true);
      if (!result.ok) continue;
      expect(Object.keys(result.transaction)).toEqual(["primitives", "postconditions"]);
      expect(result.transaction.mandatoryProofsComplete).toBe(false);
      expect(result.transaction.mandatoryPostconditions).toEqual([]);
    }
  });

  it("keeps no-manifest number-format-only format proof-complete", () => {
    const result = compileWithoutManifest({
      kind: "format_range",
      sheet: "Sheet1",
      address: "A1",
      format: { numberFormat: "0.00", columnAutofit: false, rowAutofit: false },
    });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.transaction.mandatoryProofsComplete).toBe(true);
    expect(result.transaction.mandatoryPostconditions).toMatchObject([{ kind: "range_number_format_equal" }]);
  });

  it("keeps no-manifest single-cell write_formula proof-complete", () => {
    const result = compileWithoutManifest({ kind: "write_formula", sheet: "Sheet1", address: "A1", formula: "=B1" });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.transaction.mandatoryProofsComplete).toBe(true);
    expect(result.transaction.mandatoryPostconditions).toMatchObject([{ kind: "range_formulas_equal" }]);
  });

  it("rejects unsupported no-manifest plans at preparation", async () => {
    const cases: readonly { readonly operation: WorkbookOperation; readonly effect: "write" | "format" }[] = [
      { operation: { kind: "fill_formula", sheet: "Sheet1", sourceAddress: "A1", targetAddress: "A1:A2" }, effect: "write" },
      { operation: { kind: "set_validation_list", sheet: "Sheet1", address: "A1", source: "one,two" }, effect: "write" },
      {
        operation: {
          kind: "conditional_format",
          sheet: "Sheet1",
          address: "A1",
          ruleType: "containsText",
          text: "needle",
          fillColor: "FF0000",
          fontColor: "FFFFFF",
        },
        effect: "format",
      },
      { operation: { kind: "format_range", sheet: "Sheet1", address: "A1", format: { bold: true } }, effect: "format" },
    ];

    for (const { operation, effect } of cases) {
      const result = await prepareWorkbookPlanExecution({
        plan: plan([operation]),
        compileContext: context(),
        policyInput: policy([effect]),
        trustedOperationBoundary: boundary,
      });
      expect(result, operation.kind).toMatchObject({ accepted: false, reason: "compile_failed" });
    }
  });

  it("binds a compiler-owned proof for bold-only format", () => {
    expect(compile({ kind: "format_range", sheet: "Sheet1", address: "A1", format: { bold: true } })).toMatchObject({ ok: true });
  });

  it("binds a compiler-owned proof when number format is combined with rich formatting", () => {
    expect(compile({ kind: "format_range", sheet: "Sheet1", address: "A1", format: { numberFormat: "0.00", bold: true } })).toMatchObject({ ok: true });
  });

  it("binds compiler-owned proofs for validation and conditional-format effects", () => {
    expect(compile({ kind: "set_validation_list", sheet: "Sheet1", address: "A1", source: "one,two" })).toMatchObject({ ok: true });
    expect(compile({
      kind: "conditional_format",
      sheet: "Sheet1",
      address: "A1",
      ruleType: "cellValue",
      operator: "greaterThan",
      formula1: "0",
      fillColor: "#ff0000",
    })).toMatchObject({ ok: true });
  });

  it("accepts fill_formula when an exact relative template proof exists", () => {
    const base = manifest([{ sourceFormulas: [["=B1"]] }]);
    const preconditions = {
      ...base,
      ranges: [...base.ranges, { sheet: "Sheet1", address: "A1:A2", rowCount: 2, columnCount: 1, sentinel: null, numberFormat: "General" }],
    };
    const result = compile(
      { kind: "fill_formula", sheet: "Sheet1", sourceAddress: "A1", targetAddress: "A1:A2" },
      preconditions,
    );
    expect(result).toMatchObject({ ok: true });
    if (result.ok) expect(result.transaction.mandatoryPostconditions?.[0]?.proof?.expectedFormulas).toEqual([["=B1"], ["=B2"]]);
  });

  it("fails closed for a multi-cell write_formula target", () => {
    expect(compile(
      { kind: "write_formula", sheet: "Sheet1", address: "A1:B2", formula: "=A1" },
      manifestForRange("A1:B2", 2, 2),
    )).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
  });

  it("fails before acceptance when write_values dimensions do not match the manifest target", () => {
    expect(compile(
      { kind: "write_values", sheet: "Sheet1", address: "A1:B2", values: [["a"]] },
      manifestForRange("A1:B2", 2, 2),
    )).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
  });

  it("accepts number-format-only format and proves the requested format", () => {
    const result = compile({ kind: "format_range", sheet: "Sheet1", address: "A1", format: { numberFormat: "0.00" } });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.transaction.mandatoryPostconditions?.[0]).toMatchObject({ kind: "range_number_format_equal" });
  });

  it("accepts a 1x1 write_formula target with an exact matrix", () => {
    const result = compile({ kind: "write_formula", sheet: "Sheet1", address: "A1", formula: "=B1" });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.transaction.mandatoryPostconditions?.[0].proof?.expectedFormulas).toEqual([["=B1"]]);
  });

  it("accepts write_values only with a complete matching matrix", () => {
    const result = compile({ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["after"]] });
    expect(result.ok).toBe(true);
  });
});
