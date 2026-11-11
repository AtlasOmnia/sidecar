import { describe, expect, it } from "vitest";
import { getExcelToolDefinition } from "../excelTools";
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
  type WorkbookPlanEffectGroup,
  type WorkbookPlanEffectOperation,
  type WorkbookPlanExecutionAdapter,
  validateTrustedOperationProofBindings,
} from "../workbookPlanExecutor";
import type { WorkbookOperation, WorkbookPlan } from "../generated/workbookPlanSchema";
import type { WorkbookPlanPolicyInput } from "../workbookPlanPolicy";
import type { VerificationRule } from "../workbookPlanVerification";

const IDS = {
  sessionId: "11111111-1111-4111-8111-111111111111",
  documentId: "22222222-2222-4222-8222-222222222222",
  turnId: "33333333-3333-4333-8333-333333333333",
};
const boundary = createTrustedOperationBoundary(IDS);
const primitiveNames = [
  "write_range", "format_range", "autofill_formula", "insert_table", "sort_table", "filter_table",
  "create_chart", "set_data_validation", "conditional_format_range", "freeze_panes", "create_sheet",
  "rename_sheet", "copy_sheet", "add_comment", "run_calculation",
];

function manifest(): WorkbookPlanPreconditionManifest {
  return {
    version: 1,
    sheets: [{ name: "Sheet1", exists: true }],
    tables: [],
    ranges: ["A1", "B1", "C1", "D1"].map((address) => ({
      sheet: "Sheet1", address, rowCount: 1, columnCount: 1, sentinel: null, values: [[null]], numberFormat: "General",
    })),
    chartSources: [],
    charts: [],
    comments: [{ sheet: "Sheet1", address: "D1", exists: false }],
  };
}

function context(preconditionManifest = manifest()): VerifiedWorkbookCompileContext {
  return {
    sheetNames: new Set(["Sheet1"]),
    tablesByName: new Map(),
    supportedPrimitiveNames: new Set(primitiveNames.filter((name) => getExcelToolDefinition(name) !== undefined)),
    supportedCapabilities: new Set(["excel_workbook_access", "workbook_plan_v1_operation_audit"]),
    chartSourceResolutionsBySheet: new Map(),
    preconditionManifest,
  };
}

function plan(operations: readonly WorkbookOperation[]): WorkbookPlan {
  return { version: 1, outcome: "execute", summary: "C2 proof repair", operations, verification: [{ kind: "calculation_completed" }] };
}

function policy(effects: readonly ("write" | "format" | "calculate")[]): WorkbookPlanPolicyInput {
  return {
    requestClassification: "explicit_mutation",
    activeSheetName: "Sheet1",
    selectionScopedRequest: false,
    operationSummaries: effects.map((effect, operationIndex) => ({
      operationIndex,
      effect,
      target: { operationIndex, sheetNames: effect === "calculate" ? [] : ["Sheet1"], selectionRelationship: "inside" },
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
      operationCount: 30,
      cellsRead: 100,
      cellsWritten: 100,
      formulaUtf8Bytes: 10000,
      csvPayloadUtf8Bytes: 10000,
      commentUtf8Bytes: 10000,
      sheetsCreated: 10,
      chartsCreated: 10,
      tablesCreated: 10,
      verificationReads: 30,
    },
  };
}

function uuidSource(): () => string {
  const values = [IDS.turnId, "44444444-4444-4444-8444-444444444444", "55555555-5555-4555-8555-555555555555", "66666666-6666-4666-8666-666666666666", "77777777-7777-4777-8777-777777777777", "88888888-8888-4888-8888-888888888888"];
  let index = 0;
  return () => values[index++]!;
}

const ruleOnlyOperations: readonly WorkbookOperation[] = [
  { kind: "write_values", sheet: "Sheet1", address: "A1", values: [[1]] },
  { kind: "write_formula", sheet: "Sheet1", address: "B1", formula: "=A1" },
  { kind: "format_range", sheet: "Sheet1", address: "C1", format: { numberFormat: "0.00" } },
  { kind: "add_comment", sheet: "Sheet1", address: "D1", text: "new" },
  { kind: "calculate" },
];

describe("C2 R1 mandatory compiler proofs", () => {
  it("binds an immutable trusted operation proof for every rule-only operation class", () => {
    const result = compileWorkbookPlan(plan(ruleOnlyOperations), context());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.transaction.mandatoryProofsComplete).toBe(true);
    expect(result.transaction.mandatoryPostconditions).toHaveLength(ruleOnlyOperations.length);
    expect(result.transaction.mandatoryPostconditions?.map((entry) => [entry.operationIndex, entry.kind, entry.proof?.operationKind])).toEqual([
      [0, "range_values_equal", "write_values"],
      [1, "range_formulas_equal", "write_formula"],
      [2, "range_number_format_equal", "format_range"],
      [3, "comment_exists", "add_comment"],
      [4, "calculation_completed", "calculate"],
    ]);
  });

  it("rejects same-kind proofs with a wrong cryptographic argument digest or effect", () => {
    const result = compileWorkbookPlan(plan(ruleOnlyOperations), context());
    expect(result.ok).toBe(true);
    if (!result.ok || result.transaction.mandatoryPostconditions === undefined) return;
    const operations = result.transaction.primitives.map((primitive, index) => ({
      operationId: `${index + 1}`.padStart(36, "0"),
      ...primitive,
    }));
    const operationDigests = operations.map((_, index) => index.toString(16).padStart(64, "0"));
    const proofs = result.transaction.mandatoryPostconditions.map((entry) => ({
      ...entry,
      ...(entry.proof === undefined ? {} : { proof: { ...entry.proof, argsDigest: operationDigests[entry.operationIndex] } }),
    }));
    expect(validateTrustedOperationProofBindings(proofs, operations, operationDigests)).toBe(true);

    const wrongArguments = proofs.map((entry, index) => index === 0 && entry.proof !== undefined
      ? { ...entry, proof: { ...entry.proof, argsDigest: "f".repeat(64) } }
      : entry);
    expect(validateTrustedOperationProofBindings(wrongArguments, operations, operationDigests)).toBe(false);

    const wrongEffect = proofs.map((entry, index) => index === 0 && entry.proof !== undefined
      ? { ...entry, proof: { ...entry.proof, effect: "format" as never } }
      : entry);
    expect(validateTrustedOperationProofBindings(wrongEffect, operations, operationDigests)).toBe(false);
  });

  it("does not report or persist committed when any operation-bound mandatory proof fails", async () => {
    const prepared = await prepareWorkbookPlanExecution({
      plan: plan(ruleOnlyOperations),
      compileContext: context(),
      policyInput: policy(["write", "write", "format", "write", "calculate"]),
      trustedOperationBoundary: boundary,
      uuidSource: uuidSource(),
    });
    expect(prepared.accepted).toBe(true);
    if (!prepared.accepted) return;
    let verificationCalls = 0;
    let verifiedProofCount = 0;
    const adapter: WorkbookPlanExecutionAdapter = {
      async executeFirstGroup(gate) {
        return gate.executeAfterFresh("fresh", async () => ({ status: "committed" }));
      },
      async executeGroup() { return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch" as const); },
      async verifyMandatoryPostconditions(proofs, operations) {
        verificationCalls += 1;
        verifiedProofCount = proofs.length;
        expect(proofs.every((proof) => proof.proof !== undefined)).toBe(true);
        expect(new Set(proofs.map((proof) => proof.operationIndex)).size).toBe(operations.length);
        return false;
      },
    };
    const report = await executeAcceptedWorkbookPlan(
      prepared.transaction,
      boundary,
      adapter,
      new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
    );
    expect(verificationCalls).toBe(1);
    expect(verifiedProofCount).toBe(ruleOnlyOperations.length);
    expect(report.outcome).toBe("partial");
    expect(report.verification).toBe("reconciliation_failed");
    expect(report.operations.every((operation) => operation.state === "unknown_within_batch")).toBe(true);
  });
});

type AnyRecord = Record<string, any>;

type ReconciliationOperation = Pick<WorkbookPlanEffectOperation, "operationId" | "operationIndex" | "kind" | "primitiveName" | "args">;

function valueReconciliationFake(initial: unknown[][]) {
  let loaded = false;
  let pendingValues: unknown[][] | undefined;
  let failMutation = true;
  let current = initial.map((row) => [...row]);
  const range: AnyRecord = {};
  Object.defineProperty(range, "values", {
    get: () => {
      if (!loaded) throw new Error("read before load+sync: values");
      return current;
    },
    set: (value: unknown) => { pendingValues = value as unknown[][]; },
  });
  range.formulas = [[""]];
  range.numberFormat = [["General"]];
  range.load = (fields: string | readonly string[]) => {
    const values = typeof fields === "string" ? fields.split(",") : [...fields];
    if (values.join(",") !== "values,formulas,numberFormat" && values.join(",") !== "values") throw new Error("unexpected value read load");
    loaded = false;
  };
  const context = {
    workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
    sync: async () => {
      if (pendingValues !== undefined) {
        const next = pendingValues;
        pendingValues = undefined;
        if (failMutation) throw new Error("forced mutation sync failure");
        current = next.map((row) => [...row]);
      }
      loaded = true;
    },
  };
  const group = (values: unknown[][]): { groupIndex: number; effect: "write"; operations: readonly ReconciliationOperation[] } => ({
    groupIndex: 0,
    effect: "write",
    operations: [{
      operationId: "11111111-1111-4111-8111-111111111111",
      operationIndex: 0,
      kind: "write_values",
      primitiveName: "write_range",
      args: { sheet: "Sheet1", address: "A1", values },
    }],
  });
  return {
    context,
    group,
    setCurrent(values: readonly (readonly unknown[])[]) { current = values.map((row) => [...row]); loaded = true; },
    async withExcel<T>(task: () => Promise<T>): Promise<T> {
      const globals = globalThis as AnyRecord;
      const previous = globals.Excel;
      globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
      try { return await task(); } finally {
        if (previous === undefined) delete globals.Excel;
        else globals.Excel = previous;
      }
    },
  };
}

function commentReconciliationFake(initialContent?: string) {
  let pendingContent: { comment: AnyRecord; value: string } | undefined;
  let pendingAdd: AnyRecord | undefined;
  let failMutation = true;
  let comments: AnyRecord[] = [];
  const makeComment = (content: string): AnyRecord => {
    let contentLoaded = false;
    const location: AnyRecord = { address: "Sheet1!A1", load: (fields: string | readonly string[]) => {
      const requested = typeof fields === "string" ? fields.split(",") : [...fields];
      if (requested.length !== 1 || requested[0] !== "address") throw new Error("unexpected comment location load");
      location.__loaded = true;
    } };
    const comment: AnyRecord = {
      getLocation: () => location,
      load: (fields: string | readonly string[]) => {
        const requested = typeof fields === "string" ? fields.split(",") : [...fields];
        if (requested.length !== 1 || requested[0] !== "content") throw new Error("unexpected comment content load");
        contentLoaded = true;
      },
    };
    Object.defineProperty(comment, "content", {
      get: () => { if (!contentLoaded) throw new Error("read before load+sync: comment content"); return comment.__content; },
      set: (value: string) => { pendingContent = { comment, value }; },
    });
    comment.__content = content;
    return comment;
  };
  if (initialContent !== undefined) comments = [makeComment(initialContent)];
  let countLoaded = false;
  let itemsLoaded = false;
  const countResult: AnyRecord = {};
  Object.defineProperty(countResult, "value", { get: () => { if (!countLoaded) throw new Error("read before sync: comment count"); return comments.length; } });
  const collection: AnyRecord = {
    getCount: () => countResult,
    load: (fields: string | readonly string[]) => {
      const requested = typeof fields === "string" ? fields.split(",") : [...fields];
      if (requested.join(",") !== "items") throw new Error("unexpected comments load");
      itemsLoaded = false;
    },
    add: (_range: unknown, content: string) => {
      const comment = makeComment(content);
      pendingAdd = comment;
      return comment;
    },
  };
  Object.defineProperty(collection, "items", { get: () => { if (!itemsLoaded) throw new Error("read before sync: comments items"); return comments; } });
  const range: AnyRecord = { load: () => undefined };
  const context = {
    workbook: { worksheets: { getItem: () => ({ getRange: () => range, comments: collection }) } },
    sync: async () => {
      if (pendingContent !== undefined || pendingAdd !== undefined) {
        const content = pendingContent;
        const add = pendingAdd;
        pendingContent = undefined;
        pendingAdd = undefined;
        if (failMutation) throw new Error("forced comment mutation sync failure");
        if (content !== undefined) content.comment.__content = content.value;
        if (add !== undefined) comments.push(add);
      }
      countLoaded = true;
      itemsLoaded = true;
    },
  };
  const group = (text: string): { groupIndex: number; effect: "write"; operations: readonly ReconciliationOperation[] } => ({
    groupIndex: 0,
    effect: "write",
    operations: [{
      operationId: "22222222-2222-4222-8222-222222222222",
      operationIndex: 0,
      kind: "add_comment",
      primitiveName: "add_comment",
      args: { sheet: "Sheet1", address: "A1", text },
    }],
  });
  return {
    context,
    group,
    setCurrent(content: string | undefined) { comments = content === undefined ? [] : [makeComment(content)]; },
    async withExcel<T>(task: () => Promise<T>): Promise<T> {
      const globals = globalThis as AnyRecord;
      const previous = globals.Excel;
      globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
      try { return await task(); } finally {
        if (previous === undefined) delete globals.Excel;
        else globals.Excel = previous;
      }
    },
  };
}

describe("C2 R2 exact operation pre-state reconciliation", () => {
  it.each([
    ["requested current state", [["requested"]], "committed"],
    ["exact prior state", [["prior"]], "not_committed"],
    ["third partial state", [["third"]], "unknown_within_batch"],
  ] as const)("classifies write_values from %s", async (_label, current, expected) => {
    const fake = valueReconciliationFake([["prior"]]);
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    const operationGroup = fake.group([["requested"]]);
    await expect(fake.withExcel(() => adapter.executeGroup(operationGroup, 0))).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
    fake.setCurrent(current);
    await expect(fake.withExcel(() => adapter.reconcileFailedGroup(operationGroup))).resolves.toEqual([expected]);
  });

  it.each([
    ["requested new comment", undefined, "requested", "committed"],
    ["prior absence", undefined, undefined, "not_committed"],
    ["third comment at bound cell", undefined, "third", "unknown_within_batch"],
    ["requested existing update", "prior", "requested", "committed"],
    ["exact prior existing content", "prior", "prior", "not_committed"],
    ["third existing content", "prior", "third", "unknown_within_batch"],
  ] as const)("classifies add_comment from %s", async (_label, initial, current, expected) => {
    const fake = commentReconciliationFake(initial ?? undefined);
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    const operationGroup = fake.group("requested");
    await expect(fake.withExcel(() => adapter.executeGroup(operationGroup, 0))).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
    fake.setCurrent(current ?? undefined);
    await expect(fake.withExcel(() => adapter.reconcileFailedGroup(operationGroup))).resolves.toEqual([expected]);
  });
});

describe("C2 R3 non-vacuous border compilation", () => {
  const compileFormat = (format: Record<string, unknown>) => compileWorkbookPlan(
    plan([{ kind: "format_range", sheet: "Sheet1", address: "A1", format } as WorkbookOperation]),
    context(),
  );

  it("rejects standalone border color/weight and borders:false instead of accepting a no-op", () => {
    expect(compileFormat({ borderColor: "#ff0000" })).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFormat({ borderWeight: "Medium" })).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFormat({ borders: false })).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
  });

  const borderSides = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"] as const;
  const borderItems = (overrides: Record<string, unknown> = {}) => borderSides.map((sideIndex) => ({
    sideIndex, style: "Continuous", color: "#111111", weight: "Medium", ...overrides,
  }));
  function strictBorderFake(items: unknown, unreadable = false, loadMode: "queued" | "noop" = "queued") {
    let borderLoadQueued = false;
    let syncCompleted = false;
    let syncCount = 0;
    const borderCollection: AnyRecord = {};
    Object.defineProperty(borderCollection, "items", { get: () => {
      if (!syncCompleted || unreadable || !borderLoadQueued) throw new Error("read before border load+sync");
      return items;
    } });
    borderCollection.load = (fields: string | readonly string[]) => {
      const requested = typeof fields === "string" ? fields.split(",") : [...fields];
      if (requested.join(",") !== "items/sideIndex,items/style,items/color,items/weight") throw new Error("unexpected border load");
      if (loadMode === "queued") borderLoadQueued = true;
    };
    const range: AnyRecord = {
      values: [[null]], formulas: [[""]], numberFormat: [["General"]], format: { borders: borderCollection },
      load: (fields: string | readonly string[]) => {
        const requested = typeof fields === "string" ? fields.split(",") : [...fields];
        if (requested.join(",") !== "values,formulas,numberFormat") throw new Error("unexpected range load");
      },
    };
    const context = {
      workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
      sync: async () => { syncCompleted = borderLoadQueued; syncCount += 1; },
    };
    const globals = globalThis as AnyRecord;
    const previous = globals.Excel;
    globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
    return {
      get syncCount() { return syncCount; },
      async run<T>(task: () => Promise<T>): Promise<T> {
        try { return await task(); } finally {
          if (previous === undefined) delete globals.Excel;
          else globals.Excel = previous;
        }
      },
    };
  }

  const borderFormatGroup = (format: AnyRecord): WorkbookPlanEffectGroup => ({
    groupIndex: 0,
    effect: "format",
    operations: [{
      operationId: "22222222-2222-4222-8222-222222222222",
      operationIndex: 0,
      kind: "format_range",
      primitiveName: "format_range",
      args: { sheet: "Sheet1", address: "A1", format },
    }],
  });

  it("loads the range and border collection before a completed sync and accepts all six exact sides", async () => {
    const fake = strictBorderFake(borderItems());
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.reconcileFailedGroup(borderFormatGroup({ borders: true, borderColor: "#111111", borderWeight: "Medium" })))).resolves.toEqual(["committed"]);
    expect(fake.syncCount).toBe(1);
  });

  it.each([
    ["missing", []],
    ["duplicate", [...borderItems().slice(0, 5), borderItems()[0]]],
    ["extra", [...borderItems(), { sideIndex: "Extra", style: "Continuous", color: "#111111", weight: "Medium" }]],
    ["malformed", [...borderItems().slice(0, 5), { sideIndex: "InsideVertical", style: "Continuous" }]],
    ["wrong style", borderItems({ style: "Dash" })],
    ["wrong color", borderItems({ color: "#222222" })],
    ["wrong weight", borderItems({ weight: "Thin" })],
  ] as const)("rejects %s six-side border readback", async (_label, items) => {
    const fake = strictBorderFake(items);
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.reconcileFailedGroup(borderFormatGroup({ borders: true, borderColor: "#111111", borderWeight: "Medium" })))).resolves.toEqual(["unknown_within_batch"]);
  });

  it("fails closed when the six-side border collection remains unreadable after sync", async () => {
    const fake = strictBorderFake(borderItems(), true);
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.reconcileFailedGroup(borderFormatGroup({ borders: true })))).resolves.toEqual(["unknown_within_batch"]);
  });

  it("does not let a completed sync unlock border items when the collection load queued no read", async () => {
    const fake = strictBorderFake(borderItems(), false, "noop");
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.reconcileFailedGroup(borderFormatGroup({ borders: true })))).resolves.toEqual(["unknown_within_batch"]);
    expect(fake.syncCount).toBe(1);
  });
});

function testRangeShape(address: string): { readonly rows: number; readonly columns: number } | undefined {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?$/.exec(address);
  if (match === null) return undefined;
  const column = (value: string): number => [...value.toUpperCase()].reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0);
  const firstColumn = column(match[1]!);
  const lastColumn = column(match[3] ?? match[1]!);
  const firstRow = Number(match[2]);
  const lastRow = Number(match[4] ?? match[2]);
  return lastColumn < firstColumn || lastRow < firstRow ? undefined : { rows: lastRow - firstRow + 1, columns: lastColumn - firstColumn + 1 };
}

function fillContext(sourceAddress: string, targetAddress: string, sourceFormulas: readonly (readonly string[])[]): VerifiedWorkbookCompileContext {
  const base = manifest();
  const ranges = base.ranges.filter((range) => range.address !== sourceAddress && range.address !== targetAddress);
  const targetShape = testRangeShape(targetAddress);
  return context({
    ...base,
    ranges: [
      ...ranges,
      { sheet: "Sheet1", address: sourceAddress, rowCount: sourceFormulas.length, columnCount: sourceFormulas[0]!.length, sourceFormulas },
      { sheet: "Sheet1", address: targetAddress, rowCount: targetShape?.rows ?? 0, columnCount: targetShape?.columns ?? 0 },
    ],
  });
}

describe("C2 R4 bounded exact fill-formula expansion", () => {
  const compileFill = (formula: string, sourceAddress = "A1", targetAddress = "A1:C1", sourceFormulas: readonly (readonly string[])[] = [[formula]]) =>
    compileWorkbookPlan(
      plan([{ kind: "fill_formula", sheet: "Sheet1", sourceAddress, targetAddress }]),
      fillContext(sourceAddress, targetAddress, sourceFormulas),
    );

  it("binds horizontally translated relative, absolute, mixed, string-literal, and range references", () => {
    const horizontal = compileFill("=A1+$B$1", "A1", "A1:C1");
    expect(horizontal.ok).toBe(true);
    if (horizontal.ok) expect(horizontal.transaction.mandatoryPostconditions?.[0]?.proof?.expectedFormulas).toEqual([["=A1+$B$1", "=B1+$B$1", "=C1+$B$1"]]);

    const vertical = compileFill("=A1+A$1+$A1+$A$1+SUM(A1:B2)", "A1", "A1:A3");
    expect(vertical.ok).toBe(true);
    if (vertical.ok) expect(vertical.transaction.mandatoryPostconditions?.[0]?.proof?.expectedFormulas).toEqual([
      ["=A1+A$1+$A1+$A$1+SUM(A1:B2)"],
      ["=A2+A$1+$A2+$A$1+SUM(A2:B3)"],
      ["=A3+A$1+$A3+$A$1+SUM(A3:B4)"],
    ]);

    const literal = compileFill('="A1 "&A1', "A1", "A1:B1");
    expect(literal.ok).toBe(true);
    if (literal.ok) expect(literal.transaction.mandatoryPostconditions?.[0]?.proof?.expectedFormulas).toEqual([["=\"A1 \"&A1", "=\"A1 \"&B1"]]);

    const comparison = compileFill("=IF(A1>=1,A1,0)", "A1", "A1:B1");
    expect(comparison.ok).toBe(true);
    if (comparison.ok) expect(comparison.transaction.mandatoryPostconditions?.[0]?.proof?.expectedFormulas).toEqual([["=IF(A1>=1,A1,0)", "=IF(B1>=1,B1,0)"]]);

    const boolean = compileFill("=IF(TRUE,A1,0)", "A1", "A1:B1");
    expect(boolean.ok).toBe(true);
    if (boolean.ok) expect(boolean.transaction.mandatoryPostconditions?.[0]?.proof?.expectedFormulas).toEqual([["=IF(TRUE,A1,0)", "=IF(TRUE,B1,0)"]]);
  });

  it("requires source containment, one-axis extension, and the conservative total-cell bound", () => {
    expect(compileFill("=A1", "B1", "A1")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=A1", "A1", "A1:B2")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=A1", "A1", "A1:K1000")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
  });

  it.each([
    "=SUM(1:2)",
    "=A1:5",
    '="left":"right"',
    "=1:A1",
    "=A1:B1:C1",
    "=(A1):B1",
  ])("rejects colon outside direct A1 reference tokens: %s", (formula) => {
    expect(compileFill(formula)).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
  });

  it("fails closed for overflow, unsupported grammar, malformed addresses, and multi-cell source patterns", () => {
    expect(compileFill("=XFD1", "A1", "B1")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=R1C1")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=Table1[Amount]")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=NamedRange")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=Sheet2!A1")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=[Book.xlsx]Sheet1!A1")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=A0")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=A1$", "A1", "A1:B1")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=A1?", "A1", "A1:B1")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=A1%1", "A1", "A1:B1")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=1.2.3", "A1", "A1:B1")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=SUM(A1,,B1)", "A1", "A1:B1")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=A1 B1", "A1", "A1:B1")).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
    expect(compileFill("=A1", "A1", "A1:A0")).toMatchObject({ ok: false });
    expect(compileFill("=A1", "A1:B1", "A1:B1", [["=A1", "=B1"]])).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification" } });
  });
});

describe("C2 R5 required Office loads", () => {
  async function withNoLoadRange<T>(task: (adapter: ReturnType<typeof import("../workbookPlanOfficeAdapter").createNativeWorkbookPlanAdapter>, range: Record<string, unknown>) => Promise<T>): Promise<T> {
    const range: Record<string, unknown> = { values: [[1]], formulas: [[""]], numberFormat: [["General"]] };
    const context = { workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } }, sync: async () => undefined };
    const globals = globalThis as Record<string, unknown>;
    const previous = globals.Excel;
    globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    return task(adapter, range).finally(() => {
      if (previous === undefined) delete globals.Excel;
      else globals.Excel = previous;
    });
  }

  it("does not reconcile a pre-populated range without the range load method", async () => {
    await withNoLoadRange(async (adapter) => {
      const group = {
        groupIndex: 0,
        effect: "write" as const,
        operations: [{ operationId: "33333333-3333-4333-8333-333333333333", operationIndex: 0, kind: "write_values" as const, primitiveName: "write_range", args: { sheet: "Sheet1", address: "A1", values: [[1]] } }],
      };
      await expect(adapter.reconcileFailedGroup(group)).resolves.toEqual(["unknown_within_batch"]);
    });
  });

  it("requires queued range loads for values, formulas, and numberFormat rather than trusting pre-populated fields", async () => {
    const makeRangeFake = (loadMode: "queued" | "noop") => {
      let rangeLoaded = false;
      let loadQueued = false;
      const range: AnyRecord = {};
      for (const [key, value] of [["values", [[1]]], ["formulas", [["=A1"]]], ["numberFormat", [["0.00"]]]] as const) {
        Object.defineProperty(range, key, { get: () => {
          if (!rangeLoaded) throw new Error(`read before range load+sync: ${key}`);
          return value;
        } });
      }
      range.load = (fields: string | readonly string[]) => {
        const requested = typeof fields === "string" ? fields.split(",") : [...fields];
        if (requested.join(",") !== "values,formulas,numberFormat") throw new Error("unexpected range load");
        if (loadMode === "queued") loadQueued = true;
      };
      const context = {
        workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
        sync: async () => { rangeLoaded = loadQueued; },
      };
      const globals = globalThis as AnyRecord;
      const previous = globals.Excel;
      globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
      const group = (operation: ReconciliationOperation): WorkbookPlanEffectGroup => ({ groupIndex: 0, effect: "write", operations: [operation] });
      return {
        range,
        group,
        async run<T>(task: () => Promise<T>): Promise<T> {
          try { return await task(); } finally {
            if (previous === undefined) delete globals.Excel;
            else globals.Excel = previous;
          }
        },
      };
    };
    const valueOperation: ReconciliationOperation = {
      operationId: "33333333-3333-4333-8333-333333333333", operationIndex: 0, kind: "write_values", primitiveName: "write_range",
      args: { sheet: "Sheet1", address: "A1", values: [[1]] },
    };
    const formulaOperation: ReconciliationOperation = {
      ...valueOperation, kind: "write_formula", primitiveName: "write_range", args: { sheet: "Sheet1", address: "A1", values: [["=A1"]] },
    };
    for (const operation of [valueOperation, formulaOperation]) {
      const positive = makeRangeFake("queued");
      const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
      await expect(positive.run(() => adapter.reconcileFailedGroup(positive.group(operation)))).resolves.toEqual(["committed"]);
      const noOp = makeRangeFake("noop");
      const noOpAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
      await expect(noOp.run(() => noOpAdapter.reconcileFailedGroup(noOp.group(operation)))).resolves.toEqual(["unknown_within_batch"]);
    }
    const publicPositive = makeRangeFake("queued");
    const publicAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(publicPositive.run(() => publicAdapter.verifyPlan([{ kind: "range_number_format_equal", sheet: "Sheet1", address: "A1", expectedNumberFormat: "0.00" }]))).resolves.toMatchObject({ ok: true });
    const publicNoOp = makeRangeFake("noop");
    const publicNoOpAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(publicNoOp.run(() => publicNoOpAdapter.verifyPlan([{ kind: "range_number_format_equal", sheet: "Sheet1", address: "A1", expectedNumberFormat: "0.00" }]))).resolves.toMatchObject({ ok: false });
  });

  it("requires the RangeFormat owner and every nested Font, Fill, and BorderCollection owner to load", async () => {
    type Owner = "range" | "format" | "font" | "fill" | "borders";
    const expectedFormat: AnyRecord = {
      bold: true, italic: true, fontColor: "#111111", fillColor: "#eeeeee", numberFormat: "0.00", borders: true,
      borderColor: "#222222", borderWeight: "Medium", columnAutofit: true, rowAutofit: true,
      horizontalAlignment: "Center", verticalAlignment: "Center", fontName: "Aptos", fontSize: 12, underline: "Single",
      wrapText: true, indentLevel: 2, textOrientation: 15,
    };
    const makeFormatFake = (missing?: Owner) => {
      const loaded = new Set<Owner>();
      const pending: Array<() => void> = [];
      const queueLoad = (owner: Owner, expected: readonly string[]) => (fields: string | readonly string[]) => {
        const requested = typeof fields === "string" ? fields.split(",") : [...fields];
        if (requested.join(",") !== expected.join(",")) throw new Error(`unexpected ${owner} load`);
        if (missing !== owner) pending.push(() => loaded.add(owner));
      };
      const property = (target: AnyRecord, key: string, owner: Owner, value: unknown): void => {
        Object.defineProperty(target, key, {
          get: () => { if (!loaded.has(owner)) throw new Error(`read before ${owner} load+sync`); return value; },
          set: () => undefined,
        });
      };
      const font: AnyRecord = {};
      for (const [key, value] of [["bold", true], ["italic", true], ["color", "#111111"], ["name", "Aptos"], ["size", 12], ["underline", "Single"]] as const) property(font, key, "font", value);
      font.load = queueLoad("font", ["bold", "italic", "color", "size", "name", "underline"]);
      const fill: AnyRecord = {};
      property(fill, "color", "fill", "#eeeeee");
      fill.load = queueLoad("fill", ["color"]);
      const sides = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"];
      const borders: AnyRecord = {};
      const borderItems = sides.map((sideIndex) => ({ sideIndex, style: "Continuous", color: "#222222", weight: "Medium" }));
      Object.defineProperty(borders, "items", { get: () => { if (!loaded.has("borders")) throw new Error("read before borders load+sync"); return borderItems; } });
      borders.load = queueLoad("borders", ["items/sideIndex", "items/style", "items/color", "items/weight"]);
      const format: AnyRecord = { font, fill, borders };
      for (const [key, value] of [["horizontalAlignment", "Center"], ["verticalAlignment", "Center"], ["wrapText", true], ["indentLevel", 2], ["textOrientation", 15], ["columnWidth", 100], ["rowHeight", 20]] as const) property(format, key, "format", value);
      format.load = queueLoad("format", ["horizontalAlignment", "verticalAlignment", "wrapText", "indentLevel", "textOrientation", "columnWidth", "rowHeight"]);
      const range: AnyRecord = { format };
      property(range, "numberFormat", "range", [["0.00"]]);
      range.load = queueLoad("range", ["values", "formulas", "numberFormat"]);
      range.values = [[1]];
      range.formulas = [["=A1"]];
      const context = {
        workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
        sync: async () => { for (const commit of pending.splice(0)) commit(); },
      };
      const globals = globalThis as AnyRecord;
      const previous = globals.Excel;
      globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
      const operation: ReconciliationOperation = {
        operationId: "44444444-4444-4444-8444-444444444444", operationIndex: 0, kind: "format_range", primitiveName: "format_range",
        args: { sheet: "Sheet1", address: "A1", format: expectedFormat },
      };
      return {
        group: { groupIndex: 0, effect: "format", operations: [operation] } as WorkbookPlanEffectGroup,
        async run<T>(task: () => Promise<T>): Promise<T> {
          try { return await task(); } finally {
            if (previous === undefined) delete globals.Excel;
            else globals.Excel = previous;
          }
        },
      };
    };
    const positive = makeFormatFake();
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(positive.run(() => adapter.reconcileFailedGroup(positive.group))).resolves.toEqual(["committed"]);
    for (const missing of ["range", "format", "font", "fill", "borders"] as const) {
      const fake = makeFormatFake(missing);
      const negativeAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
      await expect(fake.run(() => negativeAdapter.reconcileFailedGroup(fake.group))).resolves.toEqual(["unknown_within_batch"]);
    }
  });
});

describe("C2 R6 exact filter and reset criteria", () => {
  function filterFake(criteria: Record<string, unknown>, loadMode: "queued" | "noop" = "queued") {
    let loaded = false;
    let loadQueued = false;
    const filter: Record<string, unknown> = {};
    Object.defineProperty(filter, "criteria", { get: () => { if (!loaded) throw new Error("read before filter load+sync"); return criteria; } });
    filter.load = (fields: string | readonly string[]) => {
      const requested = typeof fields === "string" ? fields.split(",") : [...fields];
      if (requested.length !== 1 || requested[0] !== "criteria") throw new Error("unexpected filter load");
      if (loadMode === "queued") loadQueued = true;
    };
    const context = {
      workbook: { tables: { getItem: () => ({ columns: { getItemAt: () => ({ filter }) } }) }, worksheets: { getItem: () => ({}) } },
      sync: async () => { loaded = loadQueued; },
    };
    const globals = globalThis as AnyRecord;
    const previous = globals.Excel;
    globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
    const group = (operation: ReconciliationOperation): { groupIndex: number; effect: "filter"; operations: readonly ReconciliationOperation[] } => ({ groupIndex: 0, effect: "filter", operations: [operation] });
    return {
      group,
      async run<T>(task: () => Promise<T>): Promise<T> {
        try { return await task(); } finally {
          if (previous === undefined) delete globals.Excel;
          else globals.Excel = previous;
        }
      },
    };
  }

  const filterOperation = (kind: "filter_table" | "reset_table_filter", args: Record<string, unknown>): ReconciliationOperation => ({
    operationId: "44444444-4444-4444-8444-444444444444",
    operationIndex: 0,
    kind: kind as ReconciliationOperation["kind"],
    primitiveName: "filter_table",
    args: { tableName: "Table1", columnIndex: 0, ...args },
  });

  it("accepts exact Values, Custom default-And, Custom Or, and reset matrices", async () => {
    const cases = [
      { kind: "filter_table" as const, args: { values: ["x", "y"] }, criteria: { filterOn: "Values", values: ["x", "y"] } },
      { kind: "filter_table" as const, args: { criterion1: ">0", criterion2: "<10" }, criteria: { filterOn: "Custom", criterion1: ">0", criterion2: "<10" } },
      { kind: "filter_table" as const, args: { criterion1: ">0", operator: "Or" }, criteria: { filterOn: "Custom", criterion1: ">0", operator: "Or" } },
      { kind: "reset_table_filter" as const, args: {}, criteria: { filterOn: "None" } },
    ];
    for (const item of cases) {
      const fake = filterFake(item.criteria);
      await expect(fake.run(() => (async () => {
        const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
        return adapter.reconcileFailedGroup(fake.group(filterOperation(item.kind, item.args)));
      })())).resolves.toEqual(["committed"]);
    }
  });

  it.each([
    ["stale values", "filter_table", { values: ["x"] }, { filterOn: "Values", values: ["stale"] }],
    ["inconsistent values plus custom", "filter_table", { values: ["x"] }, { filterOn: "Custom", values: ["x"] }],
    ["malformed values array", "filter_table", { values: ["x"] }, { filterOn: "Values", values: "x" }],
    ["extra custom field", "filter_table", { criterion1: ">0" }, { filterOn: "Custom", criterion1: ">0", values: ["x"] }],
    ["stale criterion", "filter_table", { criterion1: ">0" }, { filterOn: "Custom", criterion1: "<0" }],
    ["stale operator", "filter_table", { criterion1: ">0", operator: "Or" }, { filterOn: "Custom", criterion1: ">0", operator: "And" }],
    ["unknown enum", "filter_table", { criterion1: ">0" }, { filterOn: "Unknown", criterion1: ">0" }],
    ["unknown operator", "filter_table", { criterion1: ">0", operator: "And" }, { filterOn: "Custom", criterion1: ">0", operator: "Xor" }],
    ["stale reset fields", "reset_table_filter", {}, { filterOn: "None", values: ["stale"] }],
  ] as const)("rejects %s filter matrix", async (_label, kind, args, criteria) => {
    const fake = filterFake(criteria);
    await expect(fake.run(() => (async () => {
      const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
      return adapter.reconcileFailedGroup(fake.group(filterOperation(kind, args)));
    })())).resolves.toEqual(["unknown_within_batch"]);
  });
  it("requires Values mode for value filters and rejects contradictory custom fields", async () => {
    const accepted = filterFake({ filterOn: "Values", values: ["x"] });
    await expect(accepted.run(() => (async () => {
      const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
      return adapter.reconcileFailedGroup(accepted.group(filterOperation("filter_table", { values: ["x"] })));
    })())).resolves.toEqual(["committed"]);

    const inconsistent = filterFake({ filterOn: "Custom", values: ["x"] });
    await expect(inconsistent.run(() => (async () => {
      const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
      return adapter.reconcileFailedGroup(inconsistent.group(filterOperation("filter_table", { values: ["x"] })));
    })())).resolves.toEqual(["unknown_within_batch"]);
  });

  it("requires Custom mode for exact custom criteria and an exact cleared reset state", async () => {
    const custom = filterFake({ filterOn: "Values", values: ["x"] });
    await expect(custom.run(() => (async () => {
      const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
      return adapter.reconcileFailedGroup(custom.group(filterOperation("filter_table", { criterion1: ">0", criterion2: "<10", operator: "And" })));
    })())).resolves.toEqual(["unknown_within_batch"]);

    const staleReset = filterFake({ filterOn: "None", values: ["stale"] });
    await expect(staleReset.run(() => (async () => {
      const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
      return adapter.reconcileFailedGroup(staleReset.group(filterOperation("reset_table_filter", {})));
    })())).resolves.toEqual(["unknown_within_batch"]);
  });

  it("does not let a completed sync unlock filter criteria when Filter.load queued no read", async () => {
    const fake = filterFake({ filterOn: "Values", values: ["x"] }, "noop");
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.reconcileFailedGroup(fake.group(filterOperation("filter_table", { values: ["x"] })))))
      .resolves.toEqual(["unknown_within_batch"]);
  });

  it("rejects contradictory and incomplete filter request shapes at the compiler boundary", () => {
    const compileFilter = (fields: Record<string, unknown>) => {
      const base = context();
      return compileWorkbookPlan(
        plan([{ kind: "filter_table", tableName: "Table1", columnIndex: 0, ...fields } as WorkbookOperation]),
        {
          ...base,
          tablesByName: new Map([["Table1", { canonicalName: "Table1", columnCount: 2, rowCount: 4, sheetName: "Sheet1" }]]),
          preconditionManifest: {
            ...base.preconditionManifest!,
            tables: [{ name: "Table1", exists: true, sheetName: "Sheet1", columnCount: 2, rowCount: 4 }],
          },
        },
      );
    };
    const rejected = [
      { values: ["x"], criterion1: ">0" },
      { values: ["x"], criterion2: "<10" },
      { values: ["x"], operator: "And" },
      { criterion2: "<10" },
      { operator: "And" },
      {},
    ];
    for (const fields of rejected) {
      expect(compileFilter(fields)).toMatchObject({ ok: false, error: { code: "missing_mandatory_verification", path: "/operations/0" } });
    }
    expect(compileFilter({ values: ["x"] })).toMatchObject({ ok: true });
    expect(compileFilter({ criterion1: ">0" })).toMatchObject({ ok: true });
    expect(compileFilter({ criterion1: ">0", criterion2: "<10", operator: "Or" })).toMatchObject({ ok: true });
  });
});

describe("C2 R5 strict object families", () => {
  const runWithContext = (context: unknown) => {
    return async <T>(task: () => Promise<T>): Promise<T> => {
      const globals = globalThis as AnyRecord;
      const previous = globals.Excel;
      globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
      try { return await task(); } finally {
        if (previous === undefined) delete globals.Excel;
        else globals.Excel = previous;
      }
    };
  };
  const group = (effect: "write" | "format" | "structure" | "filter" | "calculate", operation: ReconciliationOperation): WorkbookPlanEffectGroup => ({ groupIndex: 0, effect, operations: [operation] });

  it("gates DataValidation type/rule/prompt reads on the DataValidation owner load", async () => {
    const makeFake = (loadMode: "queued" | "noop") => {
      let loaded = false;
      let queued = false;
      const validation: AnyRecord = {};
      Object.defineProperty(validation, "type", { get: () => { if (!loaded) throw new Error("read before validation load+sync"); return "List"; } });
      Object.defineProperty(validation, "rule", { get: () => { if (!loaded) throw new Error("read before validation load+sync"); return { list: { inCellDropDown: true, source: "one,two" } }; } });
      Object.defineProperty(validation, "prompt", { get: () => { if (!loaded) throw new Error("read before validation load+sync"); return { showPrompt: true, title: "Pick", message: "Choose" }; } });
      validation.load = (fields: string | readonly string[]) => {
        const requested = typeof fields === "string" ? fields.split(",") : [...fields];
        if (requested.join(",") !== "type,rule,prompt") throw new Error("unexpected validation load");
        if (loadMode === "queued") queued = true;
      };
      const context = { workbook: { worksheets: { getItem: () => ({ getRange: () => ({ dataValidation: validation }) }) } }, sync: async () => { loaded = queued; } };
      const run = runWithContext(context);
      const operation: ReconciliationOperation = { operationId: IDS.turnId, operationIndex: 0, kind: "set_validation_list", primitiveName: "set_data_validation", args: { sheet: "Sheet1", address: "A1", source: "one,two", promptTitle: "Pick", promptMessage: "Choose" } };
      return { run, group: group("write", operation) };
    };
    for (const [mode, expected] of [["queued", "committed"], ["noop", "unknown_within_batch"]] as const) {
      const fake = makeFake(mode);
      const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
      await expect(fake.run(() => adapter.reconcileFailedGroup(fake.group))).resolves.toEqual([expected]);
    }
  });

  it("gates TableSort fields and the created-table object/range proof on their actual owners", async () => {
    const sortOperation: ReconciliationOperation = { operationId: IDS.turnId, operationIndex: 0, kind: "sort_table", primitiveName: "sort_table", args: { tableName: "Table1", columnIndex: 1, ascending: false } };
    const makeSortFake = (loadMode: "queued" | "noop") => {
      let loaded = false;
      let queued = false;
      const sort: AnyRecord = {};
      Object.defineProperty(sort, "fields", { get: () => { if (!loaded) throw new Error("read before TableSort load+sync"); return [{ key: 1, ascending: false }]; } });
      sort.load = (fields: string | readonly string[]) => {
        const requested = typeof fields === "string" ? fields.split(",") : [...fields];
        if (requested.join(",") !== "fields") throw new Error("unexpected TableSort load");
        if (loadMode === "queued") queued = true;
      };
      const body = { values: [[1, 2]], load: (fields: string | readonly string[]) => { if (fields !== "values") throw new Error("unexpected table body load"); } };
      const context = { workbook: { tables: { getItem: () => ({ sort, getDataBodyRange: () => body }) }, worksheets: { getItem: () => ({}) } }, sync: async () => { loaded = queued; } };
      return { run: runWithContext(context), group: group("write", sortOperation) };
    };
    const positive = makeSortFake("queued");
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(positive.run(() => adapter.executeGroup(positive.group, 0))).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
    await expect(positive.run(() => adapter.reconcileFailedGroup(positive.group))).resolves.toEqual(["committed"]);
    const negative = makeSortFake("noop");
    const negativeAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(negative.run(() => negativeAdapter.reconcileFailedGroup(negative.group))).resolves.toEqual(["unknown_within_batch"]);

    const makeTableFake = (loadMode: "queued" | "reconcile-noop") => {
      let created = false;
      let targetLoaded = false;
      let rangeLoaded = false;
      let targetLoadCount = 0;
      let rangeLoadCount = 0;
      const pending: Array<() => void> = [];
      const target: AnyRecord = {};
      const objectName = "Hermes_Table_33333333_3333_4333_8333_333333333333";
      Object.defineProperties(target, {
        isNullObject: { get: () => { if (!targetLoaded) throw new Error("read before table load+sync"); return !created; } },
        name: { get: () => { if (!targetLoaded) throw new Error("read before table load+sync"); return objectName; }, set: () => undefined },
        showHeaders: { get: () => { if (!targetLoaded) throw new Error("read before table load+sync"); return true; } },
        showTotals: { get: () => { if (!targetLoaded) throw new Error("read before table load+sync"); return false; } },
      });
      target.load = (fields: string | readonly string[]) => {
        const requested = typeof fields === "string" ? fields.split(",") : [...fields];
        const accepted = requested.join(",") === "isNullObject" || requested.join(",") === "isNullObject,name,showHeaders,showTotals";
        if (!accepted) throw new Error("unexpected table load");
        targetLoadCount += 1;
        targetLoaded = false;
        if (loadMode === "queued" || targetLoadCount === 1) pending.push(() => { targetLoaded = true; });
      };
      const range: AnyRecord = {};
      for (const [key, value] of [["address", "Sheet1!A1:C4"], ["rowCount", 4], ["columnCount", 3]] as const) Object.defineProperty(range, key, { get: () => { if (!rangeLoaded) throw new Error("read before created-table range load+sync"); return value; } });
      range.load = (fields: string | readonly string[]) => {
        const requested = typeof fields === "string" ? fields.split(",") : [...fields];
        if (requested.join(",") !== "address,rowCount,columnCount") throw new Error("unexpected created-table range load");
        rangeLoadCount += 1;
        rangeLoaded = false;
        if (loadMode === "queued") pending.push(() => { rangeLoaded = true; });
      };
      target.getRange = () => range;
      const sheet = { getRange: () => range, tables: { add: () => { created = true; return target; } } };
      const context = { workbook: { tables: { getItemOrNullObject: () => target }, worksheets: { getItem: () => sheet } }, sync: async () => { for (const commit of pending.splice(0)) commit(); } };
      const operation: ReconciliationOperation = { operationId: IDS.turnId, operationIndex: 0, kind: "create_table", primitiveName: "insert_table", args: { sheet: "Sheet1", address: "A1:C4", hasHeaders: true } };
      return { run: runWithContext(context), group: group("structure", operation), get counts() { return { targetLoadCount, rangeLoadCount }; } };
    };
    const positiveTable = makeTableFake("queued");
    const tableAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(positiveTable.run(() => tableAdapter.executeGroup(positiveTable.group, 0))).resolves.toEqual({ status: "committed" });
    await expect(positiveTable.run(() => tableAdapter.reconcileFailedGroup(positiveTable.group))).resolves.toEqual(["committed"]);
    expect(positiveTable.counts.targetLoadCount).toBe(2);
    expect(positiveTable.counts.rangeLoadCount).toBe(1);
    const negativeTable = makeTableFake("reconcile-noop");
    const negativeTableAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(negativeTable.run(() => negativeTableAdapter.executeGroup(negativeTable.group, 0))).resolves.toEqual({ status: "committed" });
    await expect(negativeTable.run(() => negativeTableAdapter.reconcileFailedGroup(negativeTable.group))).resolves.toEqual(["unknown_within_batch"]);
  });

  it("gates worksheet null-object identity for create, rename, and copy", async () => {
    const makeFake = (kind: "create_sheet" | "rename_sheet" | "copy_sheet", loadMode: "queued" | "noop") => {
      const loaded = new Map<string, boolean>();
      const queued = new Set<string>();
      const exists = new Map(kind === "rename_sheet" ? [["Old", false], ["New", true]] : [[kind === "create_sheet" ? "New" : "Copy", true]]);
      const objects = new Map<string, AnyRecord>();
      const object = (name: string) => {
        const existing = objects.get(name);
        if (existing !== undefined) return existing;
        const target: AnyRecord = {};
        Object.defineProperty(target, "isNullObject", { get: () => { if (loaded.get(name) !== true) throw new Error("read before worksheet load+sync"); return exists.get(name) !== true; } });
        target.load = (fields: string | readonly string[]) => {
          if (fields !== "isNullObject") throw new Error("unexpected worksheet load");
          loaded.set(name, false);
          if (loadMode === "queued") queued.add(name);
        };
        objects.set(name, target);
        return target;
      };
      const context = { workbook: { worksheets: { getItemOrNullObject: (name: string) => object(name), getItem: () => ({}) } }, sync: async () => { for (const name of queued) loaded.set(name, true); queued.clear(); } };
      const args = kind === "create_sheet" ? { name: "New" } : kind === "rename_sheet" ? { oldName: "Old", newName: "New" } : { sourceName: "Old", newName: "Copy" };
      const operation: ReconciliationOperation = { operationId: IDS.turnId, operationIndex: 0, kind, primitiveName: kind, args };
      return { run: runWithContext(context), group: group("structure", operation) };
    };
    for (const kind of ["create_sheet", "rename_sheet", "copy_sheet"] as const) {
      const positive = makeFake(kind, "queued");
      const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
      await expect(positive.run(() => adapter.reconcileFailedGroup(positive.group))).resolves.toEqual(["committed"]);
      const negative = makeFake(kind, "noop");
      const negativeAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
      await expect(negative.run(() => negativeAdapter.reconcileFailedGroup(negative.group))).resolves.toEqual(["unknown_within_batch"]);
    }
  });

  it("gates FreezePanes location and Application calculationState, including mismatch and unavailable reads", async () => {
    const makeFreeze = (mode: "queued" | "noop", rows = 2, columns = 3) => {
      let loaded = false;
      let queued = false;
      const location: AnyRecord = {};
      for (const [key, value] of [["isNullObject", false], ["address", "Sheet1!A1"], ["rowCount", rows], ["columnCount", columns]] as const) Object.defineProperty(location, key, { get: () => { if (!loaded) throw new Error("read before freeze location load+sync"); return value; } });
      location.load = (fields: string | readonly string[]) => { const requested = typeof fields === "string" ? fields.split(",") : [...fields]; if (requested.join(",") !== "isNullObject,address,rowCount,columnCount") throw new Error("unexpected freeze load"); if (mode === "queued") queued = true; };
      const context = { workbook: { worksheets: { getItem: () => ({ freezePanes: { getLocationOrNullObject: () => location } }) } }, sync: async () => { loaded = queued; } };
      const operation: ReconciliationOperation = { operationId: IDS.turnId, operationIndex: 0, kind: "freeze_panes", primitiveName: "freeze_panes", args: { sheet: "Sheet1", rows: 2, columns: 3 } };
      return { run: runWithContext(context), group: group("format", operation) };
    };
    const positive = makeFreeze("queued");
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(positive.run(() => adapter.reconcileFailedGroup(positive.group))).resolves.toEqual(["committed"]);
    const mismatch = makeFreeze("queued", 1, 3);
    const mismatchAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(mismatch.run(() => mismatchAdapter.reconcileFailedGroup(mismatch.group))).resolves.toEqual(["unknown_within_batch"]);
    const unavailable = makeFreeze("noop");
    const unavailableAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(unavailable.run(() => unavailableAdapter.reconcileFailedGroup(unavailable.group))).resolves.toEqual(["unknown_within_batch"]);

    const makeCalculation = (mode: "queued" | "noop", state = "Done") => {
      let loaded = false;
      let queued = false;
      const application: AnyRecord = {};
      Object.defineProperty(application, "calculationState", { get: () => { if (!loaded) throw new Error("read before calculationState load+sync"); return state; } });
      application.load = (fields: string | readonly string[]) => { if (fields !== "calculationState") throw new Error("unexpected calculation load"); if (mode === "queued") queued = true; };
      const context = { workbook: { application, worksheets: { getItem: () => ({}) } }, sync: async () => { loaded = queued; } };
      const operation: ReconciliationOperation = { operationId: IDS.turnId, operationIndex: 0, kind: "calculate", primitiveName: "run_calculation", args: {} };
      return { run: runWithContext(context), group: group("calculate", operation), application };
    };
    const calcPositive = makeCalculation("queued");
    const calcAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(calcPositive.run(() => calcAdapter.reconcileFailedGroup(calcPositive.group))).resolves.toEqual(["committed"]);
    const calcMismatch = makeCalculation("queued", "Calculating");
    const calcMismatchAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(calcMismatch.run(() => calcMismatchAdapter.verifyPlan([{ kind: "calculation_completed" }]))).resolves.toMatchObject({ ok: false });
    const calcNoOp = makeCalculation("noop");
    const calcNoOpAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(calcNoOp.run(() => calcNoOpAdapter.reconcileFailedGroup(calcNoOp.group))).resolves.toEqual(["unknown_within_batch"]);
  });

  it("uses conditional-format identity IDs before details, accepts one genuinely new same-rule object, and rejects multiple new objects", async () => {
    const makeFake = (identityMode: "queued" | "noop", detailMode: "queued" | "noop", newItemCount = 1) => {
      let identityLoaded = false;
      let detailLoaded = false;
      let addCount = 0;
      const pending: Array<() => void> = [];
      const expectedRule = { operator: "GreaterThan", formula1: "0" };
      const makeItem = (id: string): AnyRecord => {
        const ruleTarget: AnyRecord = {};
        Object.defineProperty(ruleTarget, "rule", { get: () => { if (!detailLoaded) throw new Error("read before conditional details load+sync"); return expectedRule; }, set: () => undefined });
        const fill: AnyRecord = {};
        Object.defineProperty(fill, "color", { get: () => { if (!detailLoaded) throw new Error("read before conditional fill load+sync"); return "#00ff00"; }, set: () => undefined });
        const font: AnyRecord = {};
        Object.defineProperty(font, "color", { get: () => { if (!detailLoaded) throw new Error("read before conditional font load+sync"); return "#000000"; }, set: () => undefined });
        const item: AnyRecord = { cellValue: { rule: ruleTarget, format: { fill, font } } };
        Object.defineProperties(item, {
          id: { get: () => { if (!identityLoaded && !detailLoaded) throw new Error("read before conditional identity load+sync"); return id; } },
          type: { get: () => { if (!identityLoaded && !detailLoaded) throw new Error("read before conditional identity load+sync"); return "CellValue"; } },
        });
        return item;
      };
      const items: AnyRecord[] = [makeItem("pre-existing-same-rule")];
      const formats: AnyRecord = {};
      Object.defineProperty(formats, "items", { get: () => {
        if (!identityLoaded && !detailLoaded) throw new Error("read before conditional collection load+sync");
        return items;
      } });
      formats.load = (fields: string | readonly string[]) => {
        const requested = typeof fields === "string" ? fields.split(",") : [...fields];
        const identityFields = ["items/id"];
        const detailFields = ["items/id", "items/type", "items/cellValue/rule", "items/cellValue/format/fill/color", "items/cellValue/format/font/color"];
        if (requested.join(",") === identityFields.join(",")) {
          identityLoaded = false;
          if (identityMode === "queued") pending.push(() => { identityLoaded = true; });
        } else if (requested.join(",") === detailFields.join(",")) {
          detailLoaded = false;
          if (detailMode === "queued") pending.push(() => { detailLoaded = true; });
        } else throw new Error("unexpected conditional load");
      };
      formats.add = () => {
        addCount += 1;
        const created = makeItem(`new-${addCount}`);
        items.push(created);
        for (let index = 1; index < newItemCount; index += 1) items.push(makeItem(`new-extra-${index}`));
        return created;
      };
      const range = { conditionalFormats: formats };
      const context = { workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } }, sync: async () => { for (const commit of pending.splice(0)) commit(); } };
      const operation: ReconciliationOperation = { operationId: IDS.turnId, operationIndex: 0, kind: "conditional_format", primitiveName: "conditional_format_range", args: { sheet: "Sheet1", address: "A1", ruleType: "cellValue", operator: "greaterThan", formula1: "0", fillColor: "#00ff00", fontColor: "#000000" } };
      return { run: runWithContext(context), group: group("format", operation) };
    };
    const positive = makeFake("queued", "queued");
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(positive.run(() => adapter.executeGroup(positive.group, 0))).resolves.toEqual({ status: "committed" });
    await expect(positive.run(() => adapter.reconcileFailedGroup(positive.group))).resolves.toEqual(["committed"]);
    const ambiguous = makeFake("queued", "queued", 2);
    const ambiguousAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(ambiguous.run(() => ambiguousAdapter.executeGroup(ambiguous.group, 0))).resolves.toEqual({ status: "committed" });
    await expect(ambiguous.run(() => ambiguousAdapter.reconcileFailedGroup(ambiguous.group))).resolves.toEqual(["unknown_within_batch"]);
    const missingIdentity = makeFake("noop", "queued");
    const missingIdentityAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(missingIdentity.run(() => missingIdentityAdapter.executeGroup(missingIdentity.group, 0))).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
    const missingDetails = makeFake("queued", "noop");
    const missingDetailsAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(missingDetails.run(() => missingDetailsAdapter.executeGroup(missingDetails.group, 0))).resolves.toEqual({ status: "committed" });
    await expect(missingDetails.run(() => missingDetailsAdapter.reconcileFailedGroup(missingDetails.group))).resolves.toEqual(["unknown_within_batch"]);
  });

  it("gates public table and chart verification collections on their collection loads", async () => {
    const makeTablePublic = (mode: "queued" | "noop") => {
      let loaded = false;
      let queued = false;
      const table: AnyRecord = {};
      Object.defineProperty(table, "isNullObject", { get: () => { if (!loaded) throw new Error("read before public table load+sync"); return false; } });
      table.load = (fields: string | readonly string[]) => { if (fields !== "isNullObject") throw new Error("unexpected public table load"); if (mode === "queued") queued = true; };
      const context = { workbook: { tables: { getItemOrNullObject: () => table }, worksheets: { getItem: () => ({}) } }, sync: async () => { loaded = queued; } };
      return runWithContext(context);
    };
    const tableRule = { kind: "table_exists" as const, tableName: "Table1" };
    const tablePositive = makeTablePublic("queued");
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(tablePositive(() => adapter.verifyPlan([tableRule]))).resolves.toMatchObject({ ok: true });
    const tableNoOp = makeTablePublic("noop");
    const tableNoOpAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(tableNoOp(() => tableNoOpAdapter.verifyPlan([tableRule]))).resolves.toMatchObject({ ok: false });

    const makeChartPublic = (mode: "queued" | "noop") => {
      let loaded = false;
      let queued = false;
      const title: AnyRecord = {};
      Object.defineProperty(title, "text", { get: () => { if (!loaded) throw new Error("read before public chart load+sync"); return "Revenue"; } });
      const chart = { title };
      const charts: AnyRecord = {};
      Object.defineProperty(charts, "items", { get: () => { if (!loaded) throw new Error("read before public chart load+sync"); return [chart]; } });
      charts.load = (fields: string | readonly string[]) => { if (fields !== "items/title/text") throw new Error("unexpected public chart load"); if (mode === "queued") queued = true; };
      const context = { workbook: { worksheets: { getItem: () => ({ charts }) } }, sync: async () => { loaded = queued; } };
      return runWithContext(context);
    };
    const chartRule = { kind: "chart_exists" as const, sheet: "Sheet1", title: "Revenue" };
    const chartPositive = makeChartPublic("queued");
    const chartAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(chartPositive(() => chartAdapter.verifyPlan([chartRule]))).resolves.toMatchObject({ ok: true });
    const chartNoOp = makeChartPublic("noop");
    const chartNoOpAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(chartNoOp(() => chartNoOpAdapter.verifyPlan([chartRule]))).resolves.toMatchObject({ ok: false });
  });
});

describe("C2 R7 chart source and geometry binding", () => {
  function chartFake(
    areaCount: number,
    geometry: { left: number; top: number; width: number; height: number },
    sourceRangeAddress = "Sheet1!A1:B2",
    chartOverrides: AnyRecord = {},
  ) {
    let created = false;
    let sourceRangeLoaded = false;
    let sourceRangeLoadQueued = false;
    let sourceAreasLoaded = false;
    let sourceAreasLoadQueued = false;
    const sourceRange: AnyRecord = {};
    Object.defineProperty(sourceRange, "address", { get: () => {
      if (!sourceRangeLoaded) throw new Error("read before source Range load+sync");
      return sourceRangeAddress;
    } });
    sourceRange.load = (fields: string | readonly string[]) => {
      const requested = typeof fields === "string" ? fields.split(",") : [...fields];
      if (requested.join(",") !== "address") throw new Error("unexpected source Range load");
      sourceRangeLoaded = false;
      sourceRangeLoadQueued = true;
    };
    const sourceAreas: AnyRecord = {};
    Object.defineProperty(sourceAreas, "address", { get: () => {
      if (!sourceAreasLoaded) throw new Error("read before RangeAreas load+sync");
      return "Sheet1!A1:B2";
    } });
    Object.defineProperty(sourceAreas, "areaCount", { get: () => {
      if (!sourceAreasLoaded) throw new Error("read before RangeAreas load+sync");
      return areaCount;
    } });
    sourceAreas.load = (fields: string | readonly string[]) => {
      const requested = typeof fields === "string" ? fields.split(",") : [...fields];
      if (requested.join(",") !== "address,areaCount") throw new Error("unexpected RangeAreas load");
      sourceAreasLoaded = false;
      sourceAreasLoadQueued = true;
    };
    let nullChartLoaded = false;
    let nullChartLoadQueued = false;
    const nullChart: AnyRecord = {};
    Object.defineProperty(nullChart, "isNullObject", { get: () => {
      if (!nullChartLoaded) throw new Error("read before null chart load+sync");
      return chartOverrides.isNullObject === false ? false : true;
    } });
    nullChart.load = (fields: string | readonly string[]) => {
      if ((typeof fields === "string" ? fields.split(",") : [...fields]).join(",") !== "isNullObject") throw new Error("unexpected null chart load");
      nullChartLoaded = false;
      nullChartLoadQueued = true;
    };
    let chartLoaded = false;
    let chartLoadQueued = false;
    let chartName = typeof chartOverrides.name === "string" ? chartOverrides.name : "Hermes_Chart_33333333_3333_4333_8333_333333333333";
    let chartType = typeof chartOverrides.chartType === "string" ? chartOverrides.chartType : "ColumnClustered";
    let titleText = chartOverrides.title !== undefined && typeof chartOverrides.title === "object" && typeof chartOverrides.title.text === "string" ? chartOverrides.title.text : "Revenue";
    const chartState = { ...geometry };
    for (const key of ["left", "top", "width", "height"] as const) {
      if (typeof chartOverrides[key] === "number") chartState[key] = chartOverrides[key];
    }
    const title: AnyRecord = {};
    Object.defineProperty(title, "text", {
      get: () => { if (!chartLoaded) throw new Error("read before chart title load+sync"); return titleText; },
      set: (value: unknown) => { titleText = value as string; },
    });
    const chart: AnyRecord = { title };
    Object.defineProperties(chart, {
      isNullObject: { get: () => { if (!chartLoaded) throw new Error("read before chart load+sync"); return false; } },
      name: { get: () => { if (!chartLoaded) throw new Error("read before chart name load+sync"); return chartName; }, set: (value: unknown) => { chartName = value as string; } },
      chartType: { get: () => { if (!chartLoaded) throw new Error("read before chart type load+sync"); return chartType; }, set: (value: unknown) => { chartType = value as string; } },
      left: { get: () => { if (!chartLoaded) throw new Error("read before chart geometry load+sync"); return chartState.left; } },
      top: { get: () => { if (!chartLoaded) throw new Error("read before chart geometry load+sync"); return chartState.top; } },
      width: { get: () => { if (!chartLoaded) throw new Error("read before chart geometry load+sync"); return chartState.width; } },
      height: { get: () => { if (!chartLoaded) throw new Error("read before chart geometry load+sync"); return chartState.height; } },
    });
    chart.load = (fields: string | readonly string[]) => {
      const requested = typeof fields === "string" ? fields.split(",") : [...fields];
      if (requested.join(",") !== "isNullObject,name,chartType,title/text,left,top,width,height") throw new Error("unexpected chart load");
      chartLoaded = false;
      chartLoadQueued = true;
    };
    let passedSource: unknown;
    const charts = {
      getItemOrNullObject: () => created ? chart : nullChart,
      add: (type: string, source: unknown, seriesBy: string) => {
        if (type !== "ColumnClustered" || seriesBy !== "Auto") throw new Error("unexpected chart add");
        passedSource = source;
        created = true;
        return chart;
      },
    };
    const sheet = { getRange: () => sourceRange, getRanges: () => sourceAreas, charts };
    const context = {
      workbook: { worksheets: { getItem: () => sheet } },
      sync: async () => {
        if (sourceRangeLoadQueued) {
          sourceRangeLoaded = true;
          sourceRangeLoadQueued = false;
        }
        if (sourceAreasLoadQueued) {
          sourceAreasLoaded = true;
          sourceAreasLoadQueued = false;
        }
        if (nullChartLoadQueued) {
          nullChartLoaded = true;
          nullChartLoadQueued = false;
        }
        if (chartLoadQueued) {
          chartLoaded = true;
          chartLoadQueued = false;
        }
        if (!created) return;
        if (typeof chartOverrides.postSyncName === "string") chart.name = chartOverrides.postSyncName;
        if (typeof chartOverrides.postSyncType === "string") chart.chartType = chartOverrides.postSyncType;
        if (typeof chartOverrides.postSyncTitle === "string" && chart.title !== undefined) chart.title.text = chartOverrides.postSyncTitle;
      },
    };
    const globals = globalThis as AnyRecord;
    const previous = globals.Excel;
    globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
    const group: WorkbookPlanEffectGroup = {
      groupIndex: 0,
      effect: "structure",
      operations: [{
        operationId: "33333333-3333-4333-8333-333333333333",
        operationIndex: 0,
        kind: "create_chart",
        primitiveName: "create_chart",
        args: { sheet: "Sheet1", dataRange: "A1:B2", chartType: "ColumnClustered", title: "Revenue" },
      }],
    };
    return {
      group,
      context,
      sourceRange,
      sourceAreas,
      nullChart,
      chart,
      get passedSource() { return passedSource; },
      async run<T>(task: () => Promise<T>): Promise<T> {
        try { return await task(); } finally {
          if (previous === undefined) delete globals.Excel;
          else globals.Excel = previous;
        }
      },
    };
  }

  it("keeps chart proof properties unreadable until a completed sync applies each queued load", async () => {
    const fake = chartFake(1, { left: 10, top: 10, width: 200, height: 100 });
    expect(() => fake.sourceRange.address).toThrow(/before source Range load\+sync/);
    fake.sourceRange.load("address");
    expect(() => fake.sourceRange.address).toThrow(/before source Range load\+sync/);
    await fake.context.sync();
    expect(fake.sourceRange.address).toBe("Sheet1!A1:B2");

    expect(() => fake.sourceAreas.areaCount).toThrow(/before RangeAreas load\+sync/);
    fake.sourceAreas.load("address,areaCount");
    expect(() => fake.sourceAreas.areaCount).toThrow(/before RangeAreas load\+sync/);
    await fake.context.sync();
    expect(fake.sourceAreas.areaCount).toBe(1);

    expect(() => fake.nullChart.isNullObject).toThrow(/before null chart load\+sync/);
    fake.nullChart.load("isNullObject");
    expect(() => fake.nullChart.isNullObject).toThrow(/before null chart load\+sync/);
    await fake.context.sync();
    expect(fake.nullChart.isNullObject).toBe(true);

    expect(() => fake.chart.name).toThrow(/before chart name load\+sync/);
    fake.chart.load("isNullObject,name,chartType,title/text,left,top,width,height");
    expect(() => fake.chart.name).toThrow(/before chart name load\+sync/);
    await fake.context.sync();
    expect(fake.chart.name).toBe("Hermes_Chart_33333333_3333_4333_8333_333333333333");
    fake.chart.load("isNullObject,name,chartType,title/text,left,top,width,height");
    expect(() => fake.chart.name).toThrow(/before chart name load\+sync/);
  });

  it("rejects a multi-area source before charts.add mutation", async () => {
    const fake = chartFake(2, { left: 10, top: 10, width: 200, height: 100 });
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.executeGroup(fake.group, 0))).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
    expect(fake.passedSource).toBeUndefined();
  });

  it("passes the exact loaded source Range proxy to charts.add after source identity checks", async () => {
    const fake = chartFake(1, { left: 10, top: 10, width: 200, height: 100 });
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.executeGroup(fake.group, 0))).resolves.toEqual({ status: "committed" });
    expect(fake.passedSource).toBe(fake.sourceRange);
  });

  it("rejects a source Range whose loaded address disagrees with the RangeAreas identity", async () => {
    const fake = chartFake(1, { left: 10, top: 10, width: 200, height: 100 }, "Sheet1!B1:C2");
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.executeGroup(fake.group, 0))).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
    expect(fake.passedSource).toBeUndefined();
  });

  it("rejects a pre-existing internal chart-name collision before mutation", async () => {
    const fake = chartFake(1, { left: 10, top: 10, width: 200, height: 100 }, "Sheet1!A1:B2", {
      isNullObject: false,
      name: "Hermes_Chart_33333333_3333_4333_8333_333333333333",
    });
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.executeGroup(fake.group, 0))).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
    expect(fake.passedSource).toBeUndefined();
  });

  it.each([
    ["wrong type", { chartType: "Line" }],
    ["wrong geometry", { left: -1 }],
  ] as const)("does not claim a chart with %s committed", async (_label, overrides) => {
    const fake = chartFake(1, { left: 10, top: 10, width: 200, height: 100 }, "Sheet1!A1:B2", overrides);
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.executeGroup(fake.group, 0))).resolves.toEqual({ status: "committed" });
    await expect(fake.run(() => adapter.reconcileFailedGroup(fake.group))).resolves.toEqual(["unknown_within_batch"]);
  });

  it.each([
    ["wrong post-sync name", { postSyncName: "WrongName" }],
    ["wrong post-sync type", { postSyncType: "Line" }],
    ["wrong post-sync title", { postSyncTitle: "WrongTitle" }],
  ] as const)("does not claim a chart with %s committed", async (_label, overrides) => {
    const fake = chartFake(1, { left: 10, top: 10, width: 200, height: 100 }, "Sheet1!A1:B2", overrides);
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.executeGroup(fake.group, 0))).resolves.toEqual({ status: "committed" });
    await expect(fake.run(() => adapter.reconcileFailedGroup(fake.group))).resolves.toEqual(["unknown_within_batch"]);
  });

  it("fails closed when the exact source Range load is unavailable", async () => {
    const fake = chartFake(1, { left: 10, top: 10, width: 200, height: 100 });
    delete fake.sourceRange.load;
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.executeGroup(fake.group, 0))).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
  });

  it("fails closed when the RangeAreas source load is unavailable", async () => {
    const fake = chartFake(1, { left: 10, top: 10, width: 200, height: 100 });
    delete fake.sourceAreas.load;
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.executeGroup(fake.group, 0))).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
  });

  it("does not claim a created chart with malformed geometry committed", async () => {
    const fake = chartFake(1, { left: -1, top: 10, width: 200, height: 100 });
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.run(() => adapter.executeGroup(fake.group, 0))).resolves.toEqual({ status: "committed" });
    await expect(fake.run(() => adapter.reconcileFailedGroup(fake.group))).resolves.toEqual(["unknown_within_batch"]);
  });
});

describe("C2 R8 production proof seam coverage", () => {
  it("runs preparation, native first-group execution, mandatory verification, and final report for write_values", async () => {
    const ranges = new Map<string, Record<string, unknown>>();
    let pendingValues: unknown[][] | undefined;
    const makeRange = (address: string): Record<string, unknown> => {
      const existing = ranges.get(address);
      if (existing !== undefined) return existing;
      const range: Record<string, unknown> = {
        address: `Sheet1!${address}`,
        rowCount: 1,
        columnCount: 1,
        values: [[null]],
        formulas: [[""]],
        numberFormat: [["General"]],
        load: () => undefined,
      };
      Object.defineProperty(range, "values", {
        configurable: true,
        get: () => range.__values,
        set: (value: unknown) => { pendingValues = value as unknown[][]; },
      });
      range.__values = [[null]];
      ranges.set(address, range);
      return range;
    };
    const sheet = { getRange: (address: string) => makeRange(address) };
    const officeContext = {
      workbook: {
        worksheets: {
          getItem: () => sheet,
          getItemOrNullObject: () => ({ isNullObject: false, load: () => undefined }),
        },
        tables: { getItemOrNullObject: () => ({ isNullObject: true, load: () => undefined }) },
      },
      sync: async () => {
        if (pendingValues !== undefined) {
          makeRange("A1").__values = pendingValues.map((row) => [...row]);
          pendingValues = undefined;
        }
      },
    };
    const globals = globalThis as AnyRecord;
    const previous = globals.Excel;
    globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(officeContext) };
    try {
      const prepared = await prepareWorkbookPlanExecution({
        plan: plan([{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [[1]] }]),
        compileContext: context({
          ...manifest(),
          comments: [],
          ranges: manifest().ranges.map((range) => ({ ...range, resolvedAddress: `Sheet1!${range.address}` })),
        }),
        policyInput: policy(["write"]),
        trustedOperationBoundary: boundary,
        uuidSource: uuidSource(),
      });
      expect(prepared.accepted).toBe(true);
      if (!prepared.accepted) return;
      const native = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
      const report = await executeAcceptedWorkbookPlan(
        prepared.transaction,
        boundary,
        native,
        new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      );
      expect(report.outcome).toBe("committed");
      expect(report.verification).toBe("not_attempted");
      expect(JSON.stringify(prepared.transaction)).not.toContain("preState");
    } finally {
      if (previous === undefined) delete globals.Excel;
      else globals.Excel = previous;
    }
  });
});

type SortBodyFakeOptions = {
  readonly body?: unknown[][];
  readonly bodyMode?: "normal" | "missing" | "missing_load" | "malformed" | "ragged" | "oversized" | "read_before_sync";
  readonly mutationMode?: "apply" | "fail";
  readonly ascending?: boolean;
  readonly columnIndex?: number;
};

function sortBodyFake(options: SortBodyFakeOptions = {}) {
  let body = (options.body ?? [[3, "c"], [1, "a"], [2, "b"]]).map((row) => [...row]);
  let fields: AnyRecord[] = [{ key: options.columnIndex ?? 0, ascending: options.ascending !== false }];
  let bodyLoadQueued = false;
  let bodyLoaded = false;
  let sortLoadQueued = false;
  let sortLoaded = false;
  let pendingSort: AnyRecord[] | undefined;
  let mutationCalls = 0;
  const bodyRange: AnyRecord = {};
  const bodyValues = (): unknown[][] => {
    switch (options.bodyMode) {
      case "malformed": return ["not-a-row"] as never;
      case "ragged": return [[1, 2], [3]];
      case "oversized": return Array.from({ length: 10_001 }, () => [1]);
      default: return body;
    }
  };
  Object.defineProperty(bodyRange, "values", {
    get: () => {
      if (!bodyLoaded) throw new Error("read before table data-body load+sync");
      return bodyValues();
    },
  });
  bodyRange.load = (requested: string | readonly string[]) => {
    const fieldsRequested = typeof requested === "string" ? requested.split(",") : [...requested];
    if (fieldsRequested.length !== 1 || fieldsRequested[0] !== "values") throw new Error("unexpected table body load");
    if (options.bodyMode === "missing_load") throw new Error("table data-body load unavailable");
    bodyLoadQueued = true;
    bodyLoaded = false;
  };
  const sort: AnyRecord = {
    load: (requested: string | readonly string[]) => {
      const fieldsRequested = typeof requested === "string" ? requested.split(",") : [...requested];
      if (fieldsRequested.length !== 1 || fieldsRequested[0] !== "fields") throw new Error("unexpected TableSort load");
      sortLoadQueued = true;
      sortLoaded = false;
    },
    get fields() {
      if (!sortLoaded) throw new Error("read before TableSort load+sync");
      return fields;
    },
    apply: (requested: readonly AnyRecord[]) => {
      mutationCalls += 1;
      pendingSort = requested.map((field) => ({ ...field }));
    },
  };
  const table: AnyRecord = { sort };
  if (options.bodyMode !== "missing") table.getDataBodyRange = () => bodyRange;
  const context = {
    workbook: {
      tables: { getItem: () => table },
      worksheets: { getItem: () => ({}) },
    },
    sync: async () => {
      if (options.bodyMode !== "read_before_sync") {
        if (bodyLoadQueued) bodyLoaded = true;
        if (sortLoadQueued) sortLoaded = true;
      }
      bodyLoadQueued = false;
      sortLoadQueued = false;
      if (pendingSort !== undefined) {
        const next = pendingSort;
        pendingSort = undefined;
        if (options.mutationMode === "fail") throw new Error("forced sort mutation sync failure");
        fields = next;
        const key = next[0]?.key;
        const ascending = next[0]?.ascending !== false;
        if (Number.isSafeInteger(key) && typeof key === "number") {
          body = [...body].sort((left, right) => {
            const leftKey = left[key];
            const rightKey = right[key];
            if (typeof leftKey !== "number" || typeof rightKey !== "number") return 0;
            return (leftKey - rightKey) * (ascending ? 1 : -1);
          });
        }
      }
    },
  };
  const operation: ReconciliationOperation = {
    operationId: "44444444-4444-4444-8444-444444444444",
    operationIndex: 0,
    kind: "sort_table",
    primitiveName: "sort_table",
    args: { tableName: "Table1", columnIndex: options.columnIndex ?? 0, ascending: options.ascending !== false },
  };
  const group: WorkbookPlanEffectGroup = { groupIndex: 0, effect: "write", operations: [operation] };
  const globals = globalThis as AnyRecord;
  const previous = globals.Excel;
  return {
    context,
    group,
    get body() { return bodyValues().map((row) => [...row]); },
    get mutationCalls() { return mutationCalls; },
    async withExcel<T>(task: () => Promise<T>): Promise<T> {
      globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
      try { return await task(); } finally {
        if (previous === undefined) delete globals.Excel;
        else globals.Excel = previous;
      }
    },
  };
}

describe("C2 Q1 exact native sort body proof", () => {
  it("does not commit matching sort metadata when the data body stayed unchanged", async () => {
    const fake = sortBodyFake({ mutationMode: "fail" });
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.withExcel(() => adapter.executeGroup(fake.group, 0))).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
    await expect(fake.withExcel(() => adapter.reconcileFailedGroup(fake.group))).resolves.toEqual(["not_committed"]);
  });

  it.each([
    ["missing getDataBodyRange", "missing"],
    ["missing data-body load", "missing_load"],
    ["malformed rows", "malformed"],
    ["ragged rows", "ragged"],
    ["oversized rows", "oversized"],
    ["invalid key values", "normal"],
    ["read before completed sync", "read_before_sync"],
  ] as const)("fails closed before mutation for %s", async (_label, bodyMode) => {
    const fake = sortBodyFake({ bodyMode, body: bodyMode === "normal" ? [["not-numeric", "x"], [2, "y"]] : undefined });
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.withExcel(() => adapter.executeGroup(fake.group, 0))).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
    expect(fake.mutationCalls).toBe(0);
  });

  it.each([
    [true, [[1, "a"], [2, "b"], [3, "c"]]],
    [false, [[3, "c"], [2, "b"], [1, "a"]]],
  ] as const)("proves the exact full-row %s sort body", async (ascending, expected) => {
    const fake = sortBodyFake({ mutationMode: "apply", ascending });
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.withExcel(() => adapter.executeGroup(fake.group, 0))).resolves.toEqual({ status: "committed" });
    expect(fake.body).toEqual(expected);
    await expect(fake.withExcel(() => adapter.reconcileFailedGroup(fake.group))).resolves.toEqual(["committed"]);
  });

  it("does not pass the mandatory final proof when the sort body mismatches", async () => {
    const fake = sortBodyFake({ mutationMode: "fail" });
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.withExcel(() => adapter.executeGroup(fake.group, 0))).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
    await expect(fake.withExcel(() => adapter.verifyMandatoryPostconditions!([
      { operationIndex: 0, kind: "table_row_order", proof: { operationIndex: 0, operationKind: "sort_table" } },
    ], fake.group.operations))).resolves.toBe(false);
  });
});

type PublicFilterFake = {
  readonly withExcel: <T>(task: () => Promise<T>) => Promise<T>;
};

function nativePublicFilterFake(criteria: unknown, loadMode: "queued" | "noop" = "queued"): PublicFilterFake {
  let loaded = false;
  let queued = false;
  const filter: AnyRecord = {};
  Object.defineProperty(filter, "criteria", {
    get: () => {
      if (!loaded) throw new Error("read before public filter load+sync");
      return criteria;
    },
  });
  filter.load = (fields: string | readonly string[]) => {
    const requested = typeof fields === "string" ? fields.split(",") : [...fields];
    if (requested.length !== 1 || requested[0] !== "criteria") throw new Error("unexpected public filter load");
    loaded = false;
    if (loadMode === "queued") queued = true;
  };
  const context = {
    workbook: {
      tables: { getItem: () => ({ columns: { getItemAt: () => ({ filter }) } }) },
      worksheets: { getItem: () => ({}) },
    },
    sync: async () => { if (queued) loaded = true; queued = false; },
  };
  const globals = globalThis as AnyRecord;
  const previous = globals.Excel;
  return {
    async withExcel<T>(task: () => Promise<T>): Promise<T> {
      globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
      try { return await task(); } finally {
        if (previous === undefined) delete globals.Excel;
        else globals.Excel = previous;
      }
    },
  };
}

function publicCustomFilterRule(fields: Record<string, unknown>): VerificationRule {
  return {
    kind: "table_filter_state",
    tableName: "Table1",
    columnIndex: 0,
    state: "filtered",
    criterion1: ">0",
    ...fields,
  } as VerificationRule;
}

describe("C2 Q2 exact native public filter verification", () => {
  it("verifies an omitted custom operator against native omitted default-And state", async () => {
    const fake = nativePublicFilterFake({ filterOn: "Custom", criterion1: ">0" });
    const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(fake.withExcel(() => adapter.verifyPlan([publicCustomFilterRule({})]))).resolves.toMatchObject({ ok: true });
  });

  it("normalizes explicit And and omitted default-And symmetrically only for custom filters", async () => {
    const omittedExpected = nativePublicFilterFake({ filterOn: "Custom", criterion1: ">0", operator: "And" });
    const omittedAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(omittedExpected.withExcel(() => omittedAdapter.verifyPlan([publicCustomFilterRule({})]))).resolves.toMatchObject({ ok: true });

    const explicitExpected = nativePublicFilterFake({ filterOn: "Custom", criterion1: ">0" });
    const explicitAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(explicitExpected.withExcel(() => explicitAdapter.verifyPlan([publicCustomFilterRule({ operator: "And" })]))).resolves.toMatchObject({ ok: true });
  });

  it("keeps Or exact and rejects present malformed or extra native criteria instead of sanitizing them", async () => {
    const orFake = nativePublicFilterFake({ filterOn: "Custom", criterion1: ">0", operator: "And" });
    const orAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(orFake.withExcel(() => orAdapter.verifyPlan([publicCustomFilterRule({ operator: "Or" })]))).resolves.toMatchObject({ ok: false });

    for (const criteria of [
      { filterOn: "Custom", criterion1: ">0", criterion2: 10, operator: "And" },
      { filterOn: "Custom", criterion1: ">0", operator: "Xor" },
      { filterOn: "Custom", criterion1: ">0", operator: "And", unexpected: true },
    ]) {
      const fake = nativePublicFilterFake(criteria);
      const adapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
      await expect(fake.withExcel(() => adapter.verifyPlan([publicCustomFilterRule({ operator: "And" })]))).resolves.toMatchObject({ ok: false });
    }
  });

  it("retains strict Values and reset filter semantics through native public verification", async () => {
    const valuesFake = nativePublicFilterFake({ filterOn: "Values", values: ["x"] });
    const valuesAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(valuesFake.withExcel(() => valuesAdapter.verifyPlan([{
      kind: "table_filter_state", tableName: "Table1", columnIndex: 0, state: "filtered", values: ["x"],
    }]))).resolves.toMatchObject({ ok: true });

    const resetFake = nativePublicFilterFake({ filterOn: "None", criterion1: ">0" });
    const resetAdapter = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    await expect(resetFake.withExcel(() => resetAdapter.verifyPlan([{
      kind: "table_filter_state", tableName: "Table1", columnIndex: 0, state: "unfiltered",
    }]))).resolves.toMatchObject({ ok: false });
  });
});

describe("C2 Q3 pre-state privacy and terminal cleanup", () => {
  const privacyManifest = (): WorkbookPlanPreconditionManifest => ({
    ...manifest(),
    ranges: manifest().ranges.map((range) => range.address === "A1" ? { ...range, values: [["secret-value"]] } : range),
    comments: [{ sheet: "Sheet1", address: "D1", exists: true, text: "secret-comment" }],
  });
  const privacyPlan = (): WorkbookPlan => plan([
    { kind: "write_values", sheet: "Sheet1", address: "A1", values: [[1]] },
    { kind: "add_comment", sheet: "Sheet1", address: "D1", text: "requested-comment" },
  ]);
  const privacyPolicy = policy(["write", "write"]);
  const directStrings = (value: unknown, seen = new Set<object>()): string[] => {
    if (typeof value === "string") return [value];
    if (value === null || typeof value !== "object" || seen.has(value)) return [];
    seen.add(value);
    const record = value as Record<string, unknown>;
    return Object.getOwnPropertyNames(record).flatMap((key) => [key, ...directStrings(record[key], seen)]);
  };

  it("keeps workbook values, comment text, and expectedPreState out of compiled mandatory proofs by direct inspection", () => {
    const result = compileWorkbookPlan(privacyPlan(), context(privacyManifest()));
    expect(result.ok).toBe(true);
    if (!result.ok || result.transaction.mandatoryPostconditions === undefined) return;
    const strings = directStrings(result.transaction.mandatoryPostconditions).map((value) => value.toLowerCase());
    expect(strings).not.toContain("secret-value");
    expect(strings).not.toContain("secret-comment");
    expect(strings.some((value) => value.toLowerCase().includes("expectedprestate"))).toBe(false);
    expect(Object.prototype.hasOwnProperty.call(result.transaction, "preconditionManifest")).toBe(false);
  });

  it("releases the trusted full precondition snapshot after the first-gate terminal boundary", async () => {
    const prepared = await prepareWorkbookPlanExecution({
      plan: privacyPlan(),
      compileContext: context(privacyManifest()),
      policyInput: privacyPolicy,
      trustedOperationBoundary: boundary,
      uuidSource: uuidSource(),
    });
    expect(prepared.accepted).toBe(true);
    if (!prepared.accepted) return;
    let firstGate: { readonly preconditionManifest?: unknown } | undefined;
    let manifestWasUsable = false;
    const adapter: WorkbookPlanExecutionAdapter = {
      async executeFirstGroup(gate) {
        firstGate = gate;
        manifestWasUsable = gate.preconditionManifest !== undefined;
        return gate.executeAfterFresh("fresh", async () => ({ status: "committed" }));
      },
      async executeGroup() { return { status: "committed" }; },
      async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch" as const); },
      async verifyMandatoryPostconditions() { return true; },
    };
    const report = await executeAcceptedWorkbookPlan(
      prepared.transaction,
      boundary,
      adapter,
      new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
    );
    expect(report.outcome).toBe("committed");
    expect(manifestWasUsable).toBe(true);
    expect(JSON.stringify(prepared.transaction)).not.toContain("secret-value");
    expect(JSON.stringify(report)).not.toContain("secret-comment");
    expect(firstGate?.preconditionManifest).toBeUndefined();
  });

  it("does not let native adapter reuse observe a prior turn's comment pre-state", async () => {
    const first = commentReconciliationFake("secret-comment");
    const native = (await import("../workbookPlanOfficeAdapter")).createNativeWorkbookPlanAdapter();
    const prepared = await prepareWorkbookPlanExecution({
      plan: plan([{ kind: "add_comment", sheet: "Sheet1", address: "A1", text: "requested-comment" }]),
      compileContext: { ...context(), preconditionManifest: undefined },
      policyInput: policy(["write"]),
      trustedOperationBoundary: boundary,
      uuidSource: (() => {
        const values = [IDS.turnId, "22222222-2222-4222-8222-222222222222"];
        let index = 0;
        return () => values[index++]!;
      })(),
    });
    expect(prepared.accepted).toBe(true);
    if (!prepared.accepted) return;
    const wrapper: WorkbookPlanExecutionAdapter = {
      async executeFirstGroup(gate) { return gate.executeAfterFresh("fresh", async (group) => native.executeGroup(group, 0)); },
      executeGroup: (group, index, signal) => native.executeGroup(group, index, signal),
      reconcileFailedGroup: (group) => native.reconcileFailedGroup(group),
      verifyMandatoryPostconditions: (proofs, operations, signal) => native.verifyMandatoryPostconditions!(proofs, operations, signal),
      releaseExecutionState: (operations) => native.releaseExecutionState?.(operations),
    };
    const report = await first.withExcel(() => executeAcceptedWorkbookPlan(
      prepared.transaction,
      boundary,
      wrapper,
      new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
    ));
    expect(report.outcome).toBe("partial");

    const reused = commentReconciliationFake("secret-comment");
    await expect(reused.withExcel(() => native.reconcileFailedGroup(reused.group("requested-comment")))).resolves.toEqual(["unknown_within_batch"]);
  });
});

describe("C2 Q4 collision-resistant content-free proof binding", () => {
  it("binds known 32-bit digest-collision arguments to distinct SHA-256 operation digests", async () => {
    const values = ["08Ik2cjXL4", "o9Oifq7bGv"] as const;
    const digests: string[] = [];

    for (const value of values) {
      const preconditionManifest: WorkbookPlanPreconditionManifest = {
        ...manifest(),
        ranges: manifest().ranges.map((range) => range.address === "A1" ? { ...range, values: [[null]] } : range),
      };
      const prepared = await prepareWorkbookPlanExecution({
        plan: plan([{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [[value]] }]),
        compileContext: context(preconditionManifest),
        policyInput: policy(["write"]),
        trustedOperationBoundary: boundary,
        uuidSource: uuidSource(),
      });
      expect(prepared.accepted).toBe(true);
      if (!prepared.accepted) continue;

      const adapter: WorkbookPlanExecutionAdapter = {
        async executeFirstGroup(gate) { return gate.executeAfterFresh("fresh", async () => ({ status: "committed" })); },
        async executeGroup() { return { status: "committed" }; },
        async reconcileFailedGroup(group) { return group.operations.map(() => "unknown_within_batch" as const); },
        async verifyMandatoryPostconditions(proofs) {
          const digest = proofs[0]?.proof?.argsDigest;
          if (digest !== undefined) digests.push(digest);
          return true;
        },
      };
      const report = await executeAcceptedWorkbookPlan(
        prepared.transaction,
        boundary,
        adapter,
        new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      );
      expect(report.outcome).toBe("committed");
    }

    expect(digests).toHaveLength(2);
    expect(digests).toEqual(digests.map((digest) => expect.stringMatching(/^[0-9a-f]{64}$/)));
    expect(digests[0]).not.toBe(digests[1]);
  });
});
