import { describe, expect, it } from "vitest";
import { createNativeWorkbookPlanAdapter } from "../workbookPlanOfficeAdapter";
import type { WorkbookPlanEffectGroup } from "../workbookPlanExecutor";

type AnyRecord = Record<string, any>;

function strictFormulaReadbackFake(formulas: readonly (readonly string[])[]) {
  let loaded = false;
  let syncCount = 0;
  const loadCalls: Array<string | readonly string[]> = [];
  const range: AnyRecord = {};
  Object.defineProperty(range, "formulas", {
    enumerable: true,
    get: () => {
      if (!loaded) throw new Error("read before load+sync: formulas");
      return formulas;
    },
  });
  range.load = (fields: string | readonly string[]) => {
    loadCalls.push(fields);
    const requested = typeof fields === "string" ? fields.split(",") : [...fields];
    if (requested.some((field) => !["values", "formulas", "numberFormat"].includes(field))) {
      throw new Error("strict fake rejects unknown range load field");
    }
  };
  const context = {
    workbook: {
      worksheets: {
        getItem: () => ({ getRange: () => range }),
      },
    },
    sync: async () => {
      loaded = true;
      syncCount += 1;
    },
  };
  return {
    context,
    loadCalls,
    get syncCount() { return syncCount; },
    async withExcel<T>(task: () => Promise<T>): Promise<T> {
      const globals = globalThis as AnyRecord;
      const previous = globals.Excel;
      globals.Excel = {
        run: async (callback: (value: unknown) => Promise<unknown>) => callback(context),
      };
      try {
        return await task();
      } finally {
        if (previous === undefined) delete globals.Excel;
        else globals.Excel = previous;
      }
    },
  };
}

function strictFormatReadbackFake(format: AnyRecord) {
  let loaded = false;
  let syncCount = 0;
  const loadCalls: Array<string | readonly string[]> = [];
  const range: AnyRecord = {};
  Object.defineProperty(range, "format", {
    enumerable: true,
    get: () => format,
  });
  Object.defineProperty(range, "numberFormat", {
    enumerable: true,
    get: () => {
      if (!loaded) throw new Error("read before load+sync: numberFormat");
      return [["General"]];
    },
  });
  range.load = (fields: string | readonly string[]) => {
    loadCalls.push(fields);
    const requested = typeof fields === "string" ? fields.split(",") : [...fields];
    const allowed = new Set([
      "values", "formulas", "numberFormat", "format/font/bold", "format/font/italic", "format/font/color",
      "format/font/size", "format/font/name", "format/font/underline", "format/fill/color",
      "format/borders/items/style", "format/borders/items/color", "format/borders/items/weight",
      "format/horizontalAlignment", "format/verticalAlignment", "format/wrapText", "format/textOrientation",
      "format/indentLevel", "format/columnWidth", "format/rowHeight",
    ]);
    if (requested.some((field) => !allowed.has(field))) throw new Error("strict fake rejects unknown format load field");
  };
  const borders = (format.borders ?? {}) as AnyRecord;
  format.borders = borders;
  const allowedOwner = new Set(["horizontalAlignment", "verticalAlignment", "wrapText", "indentLevel", "textOrientation", "columnWidth", "rowHeight"]);
  format.load = (fields: string | readonly string[]) => {
    const requested = typeof fields === "string" ? fields.split(",") : [...fields];
    if (requested.some((field) => !allowedOwner.has(field))) throw new Error("strict fake rejects unknown range-format field");
  };
  const font = (format.font ?? {}) as AnyRecord;
  format.font = font;
  font.load = (fields: string | readonly string[]) => {
    const requested = typeof fields === "string" ? fields.split(",") : [...fields];
    if (requested.some((field) => !["bold", "italic", "color", "size", "name", "underline"].includes(field))) throw new Error("strict fake rejects unknown font field");
  };
  const fill = (format.fill ?? {}) as AnyRecord;
  format.fill = fill;
  fill.load = (fields: string | readonly string[]) => {
    const requested = typeof fields === "string" ? fields.split(",") : [...fields];
    if (requested.length !== 1 || requested[0] !== "color") throw new Error("strict fake rejects unknown fill field");
  };
  borders.load = (fields: string | readonly string[]) => {
    const requested = typeof fields === "string" ? fields.split(",") : [...fields];
    if (requested.join(",") !== "items/sideIndex,items/style,items/color,items/weight") {
      throw new Error("strict fake rejects unknown border load field");
    }
  };

  const context = {
    workbook: {
      worksheets: {
        getItem: () => ({ getRange: () => range }),
      },
    },
    sync: async () => {
      loaded = true;
      syncCount += 1;
    },
  };
  return {
    context,
    loadCalls,
    get syncCount() { return syncCount; },
    async withExcel<T>(task: () => Promise<T>): Promise<T> {
      const globals = globalThis as AnyRecord;
      const previous = globals.Excel;
      globals.Excel = {
        run: async (callback: (value: unknown) => Promise<unknown>) => callback(context),
      };
      try {
        return await task();
      } finally {
        if (previous === undefined) delete globals.Excel;
        else globals.Excel = previous;
      }
    },
  };
}

function strictConditionalCreatedIdentityFake(ambiguous = false) {
  const preExisting = {
    id: "pre-existing",
    type: "CellValue",
    cellValue: { rule: { operator: "EqualTo", formula1: "99" }, format: { fill: { color: "#aaaaaa" }, font: { color: "#000000" } } },
  };
  const created = {
    id: "created-by-operation",
    type: "CellValue",
    cellValue: { rule: { operator: "GreaterThan", formula1: "0" }, format: { fill: { color: "#00ff00" }, font: { color: "#000000" } } },
  };
  const secondCreated = { ...created, id: "second-created-by-operation" };
  const formats: AnyRecord = {
    items: [preExisting],
    load: () => undefined,
    add: () => {
      formats.items.push(created);
      if (ambiguous) formats.items.push(secondCreated);
      return created;
    },
  };
  const range = { conditionalFormats: formats };
  const context = {
    workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
    sync: async () => undefined,
  };
  return {
    context,
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

function strictFreezeFake(rowCount: number, columnCount: number, isNullObject = false) {
  let loaded = false;
  const location: AnyRecord = {};
  for (const [key, value] of Object.entries({ isNullObject, address: "A1:C2", rowCount, columnCount })) {
    Object.defineProperty(location, key, { get: () => { if (!loaded) throw new Error(`read before load+sync: freeze.${key}`); return value; } });
  }
  location.load = (fields: string | readonly string[]) => {
    const requested = typeof fields === "string" ? fields.split(",") : [...fields];
    if (requested.join(",") !== "isNullObject,address,rowCount,columnCount") throw new Error("strict fake rejects freeze load");
  };
  const context = {
    workbook: { worksheets: { getItem: () => ({ freezePanes: { getLocationOrNullObject: () => location } }) } },
    sync: async () => { loaded = true; },
  };
  return {
    context,
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

function failedFormulaGroup(): WorkbookPlanEffectGroup {
  return {
    groupIndex: 0,
    effect: "write",
    operations: [{
      operationId: "11111111-1111-4111-8111-111111111111",
      operationIndex: 0,
      kind: "write_formula",
      primitiveName: "write_range",
      args: {
        sheet: "Sheet1",
        address: "A1",
        values: [["=A1+B1"]],
      },
    }],
  };
}

function strictConditionalFormatFake() {
  let loaded = false;
  const conditional = {
    id: "new-conditional-id",
    type: "CellValue",
    cellValue: {
      rule: { operator: "GreaterThan", formula1: "0" },
      format: { fill: { color: "#00ff00" }, font: { color: "#000000" } },
    },
  };
  const formats: AnyRecord = {};
  Object.defineProperty(formats, "items", {
    get: () => {
      if (!loaded) throw new Error("read before load+sync: conditionalFormats.items");
      return [conditional];
    },
  });
  formats.load = (fields: string | readonly string[]) => {
    const requested = typeof fields === "string" ? fields.split(",") : [...fields];
    const expected = [
      "items/id", "items/type", "items/cellValue/rule", "items/cellValue/format/fill/color", "items/cellValue/format/font/color",
    ];
    if (requested.join(",") !== expected.join(",")) throw new Error("strict fake rejects conditional format load");
  };
  const range = { conditionalFormats: formats };
  const context = {
    workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
    sync: async () => { loaded = true; },
  };
  return {
    context,
    async withExcel<T>(task: () => Promise<T>): Promise<T> {
      const globals = globalThis as AnyRecord;
      const previous = globals.Excel;
      globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
      try {
        return await task();
      } finally {
        if (previous === undefined) delete globals.Excel;
        else globals.Excel = previous;
      }
    },
  };
}

function strictValidationFake() {
  let loaded = false;
  const validation: AnyRecord = {};
  const values: AnyRecord = {
    type: "List",
    rule: { list: { inCellDropDown: true, source: "one,two" } },
    prompt: { showPrompt: true, title: "Pick", message: "Choose" },
  };
  for (const key of Object.keys(values)) {
    Object.defineProperty(validation, key, {
      get: () => {
        if (!loaded) throw new Error(`read before load+sync: dataValidation.${key}`);
        return values[key];
      },
    });
  }
  validation.load = (fields: string | readonly string[]) => {
    const requested = typeof fields === "string" ? fields.split(",") : [...fields];
    if (requested.join(",") !== "type,rule,prompt") throw new Error("strict fake rejects validation load");
  };
  const range = { dataValidation: validation };
  const context = {
    workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
    sync: async () => { loaded = true; },
  };
  return {
    context,
    async withExcel<T>(task: () => Promise<T>): Promise<T> {
      const globals = globalThis as AnyRecord;
      const previous = globals.Excel;
      globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
      try {
        return await task();
      } finally {
        if (previous === undefined) delete globals.Excel;
        else globals.Excel = previous;
      }
    },
  };
}

function strictChartIdentityFake(operationId: string) {
  let loaded = false;
  const requestedNames: string[] = [];
  const preExistingChart: AnyRecord = { isNullObject: false, title: { text: "Revenue" }, name: "PreExistingChart" };
  const nullChart: AnyRecord = { isNullObject: true };
  const charts = {
    getItemOrNullObject: (name: string) => {
      requestedNames.push(name);
      return name === "Revenue" ? preExistingChart : nullChart;
    },
  };
  const context = {
    workbook: {
      worksheets: { getItem: () => ({ charts }) },
    },
    sync: async () => { loaded = true; },
  };
  preExistingChart.load = (fields: string | readonly string[]) => {
    const requested = typeof fields === "string" ? fields.split(",") : [...fields];
    if (requested.length !== 1 || requested[0] !== "isNullObject") throw new Error("strict fake rejects chart load");
  };
  nullChart.load = preExistingChart.load;
  return {
    context,
    requestedNames,
    async withExcel<T>(task: () => Promise<T>): Promise<T> {
      const globals = globalThis as AnyRecord;
      const previous = globals.Excel;
      globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
      try {
        return await task();
      } finally {
        if (previous === undefined) delete globals.Excel;
        else globals.Excel = previous;
      }
    },
  };
}

function strictTableRowOrderFake() {
  let loadedBody = false;
  let loadedWhole = false;
  let dataBodyRangeCalls = 0;
  const bodyRange: AnyRecord = {};
  const wholeRange: AnyRecord = {};
  Object.defineProperty(bodyRange, "values", {
    get: () => {
      if (!loadedBody) throw new Error("read before load+sync: table data body values");
      return [["body-1"], ["body-2"]];
    },
  });
  Object.defineProperty(wholeRange, "values", {
    get: () => {
      if (!loadedWhole) throw new Error("read before load+sync: table whole values");
      return [["header"], ["body-1"], ["body-2"], ["totals"]];
    },
  });
  bodyRange.load = (fields: string | readonly string[]) => {
    const requested = typeof fields === "string" ? fields.split(",") : [...fields];
    if (requested.length !== 1 || requested[0] !== "values") throw new Error("strict fake rejects data-body load");
  };
  wholeRange.load = bodyRange.load;
  const table = {
    getDataBodyRange: () => {
      dataBodyRangeCalls += 1;
      return bodyRange;
    },
    getRange: () => wholeRange,
  };
  const context = {
    workbook: {
      tables: { getItem: () => table },
      worksheets: { getItem: () => ({}) },
    },
    sync: async () => {
      loadedBody = true;
      loadedWhole = true;
    },
  };
  return {
    context,
    table,
    get dataBodyRangeCalls() { return dataBodyRangeCalls; },
    async withExcel<T>(task: () => Promise<T>): Promise<T> {
      const globals = globalThis as AnyRecord;
      const previous = globals.Excel;
      globals.Excel = { run: async (callback: (value: unknown) => Promise<unknown>) => callback(context) };
      try {
        return await task();
      } finally {
        if (previous === undefined) delete globals.Excel;
        else globals.Excel = previous;
      }
    },
  };
}

function failedConditionalGroup(): WorkbookPlanEffectGroup {
  return {
    groupIndex: 0,
    effect: "format",
    operations: [{
      operationId: "55555555-5555-4555-8555-555555555555",
      operationIndex: 0,
      kind: "conditional_format",
      primitiveName: "conditional_format_range",
      args: {
        sheet: "Sheet1",
        address: "A1:B2",
        ruleType: "cellValue",
        operator: "greaterThan",
        formula1: "0",
        fillColor: "#00ff00",
        fontColor: "#000000",
      },
    }],
  };
}

function failedValidationGroup(): WorkbookPlanEffectGroup {
  return {
    groupIndex: 0,
    effect: "write",
    operations: [{
      operationId: "44444444-4444-4444-8444-444444444444",
      operationIndex: 0,
      kind: "set_validation_list",
      primitiveName: "set_data_validation",
      args: {
        sheet: "Sheet1",
        address: "A1",
        source: "one,two",
        promptTitle: "Pick",
        promptMessage: "Choose",
      },
    }],
  };
}

function failedChartGroup(operationId: string): WorkbookPlanEffectGroup {
  return {
    groupIndex: 0,
    effect: "structure",
    operations: [{
      operationId,
      operationIndex: 0,
      kind: "create_chart",
      primitiveName: "create_chart",
      args: {
        sheet: "Sheet1",
        dataRange: "A1:B2",
        chartType: "ColumnClustered",
        title: "Revenue",
      },
    }],
  };
}

function failedFormatGroup(format: AnyRecord = { bold: true }): WorkbookPlanEffectGroup {
  return {
    groupIndex: 0,
    effect: "format",
    operations: [{
      operationId: "22222222-2222-4222-8222-222222222222",
      operationIndex: 0,
      kind: "format_range",
      primitiveName: "format_range",
      args: {
        sheet: "Sheet1",
        address: "A1",
        format,
      },
    }],
  };
}

describe("Task 11 Slice C2 conservative Office reconciliation", () => {
  it("does not claim a formula write committed when the loaded target formula differs", async () => {
    const fake = strictFormulaReadbackFake([["=A1+C1"]]);
    const adapter = createNativeWorkbookPlanAdapter();

    await expect(fake.withExcel(() => adapter.reconcileFailedGroup(failedFormulaGroup()))).resolves.toEqual([
      "unknown_within_batch",
    ]);
    expect(fake.syncCount).toBe(1);
    expect(fake.loadCalls).toEqual([["values", "formulas", "numberFormat"]]);
  });

  it("reconciles a validation operation only from exact type, rule, source, and prompt readback", async () => {
    const fake = strictValidationFake();
    const adapter = createNativeWorkbookPlanAdapter();

    await expect(fake.withExcel(() => adapter.reconcileFailedGroup(failedValidationGroup()))).resolves.toEqual([
      "committed",
    ]);
  });

  it("reconciles a conditional format only from one exact loaded rule and differential format", async () => {
    const fake = strictConditionalFormatFake();
    const adapter = createNativeWorkbookPlanAdapter();

    await expect(fake.withExcel(() => adapter.reconcileFailedGroup(failedConditionalGroup()))).resolves.toEqual([
      "unknown_within_batch",
    ]);
  });

  it("does not claim a format operation committed when a requested bold value mismatches", async () => {
    const fake = strictFormatReadbackFake({ font: { bold: false } });
    const adapter = createNativeWorkbookPlanAdapter();

    await expect(fake.withExcel(() => adapter.reconcileFailedGroup(failedFormatGroup()))).resolves.toEqual([
      "unknown_within_batch",
    ]);
    expect(fake.syncCount).toBe(1);
  });

  it.each([
    ["italic", { italic: true }, { font: { italic: false } }],
    ["font color", { fontColor: "#111111" }, { font: { color: "#222222" } }],
    ["font size", { fontSize: 12 }, { font: { size: 11 } }],
    ["font name", { fontName: "Aptos" }, { font: { name: "Calibri" } }],
    ["underline", { underline: "Single" }, { font: { underline: "None" } }],
    ["fill color", { fillColor: "#111111" }, { fill: { color: "#222222" } }],
    ["horizontal alignment", { horizontalAlignment: "Center" }, { horizontalAlignment: "General" }],
    ["vertical alignment", { verticalAlignment: "Center" }, { verticalAlignment: "Top" }],
    ["wrap", { wrapText: true }, { wrapText: false }],
    ["text orientation", { textOrientation: 15 }, { textOrientation: 0 }],
    ["indent", { indentLevel: 2 }, { indentLevel: 0 }],
    ["number format", { numberFormat: "0.00" }, {}],
    ["borders", { borders: true }, { borders: { items: [] } }],
    ["column autofit observable", { columnAutofit: true }, { columnWidth: null }],
    ["row autofit observable", { rowAutofit: true }, { rowHeight: null }],
  ] as const)("does not claim a %s mismatch committed", async (_label, requested, actual) => {
    const fake = strictFormatReadbackFake(actual);
    const adapter = createNativeWorkbookPlanAdapter();

    await expect(fake.withExcel(() => adapter.reconcileFailedGroup(failedFormatGroup(requested)))).resolves.toEqual([
      "unknown_within_batch",
    ]);
  });

  it("verifies table row order from the data-body range, excluding headers and totals", async () => {
    const fake = strictTableRowOrderFake();
    const adapter = createNativeWorkbookPlanAdapter();

    await expect(fake.withExcel(() => adapter.verifyPlan([
      { kind: "table_row_order", tableName: "Table1", expectedOrder: ["body-1", "body-2"] },
    ]))).resolves.toMatchObject({ ok: true, counts: { verified: 1 } });
    expect(fake.dataBodyRangeCalls).toBe(1);
  });

  it("reconciles frozen rows and columns from Worksheet.freezePanes location readback", async () => {
    const fake = strictFreezeFake(2, 3);
    const adapter = createNativeWorkbookPlanAdapter();
    const group: WorkbookPlanEffectGroup = {
      groupIndex: 0,
      effect: "format",
      operations: [{
        operationId: "66666666-6666-4666-8666-666666666666",
        operationIndex: 0,
        kind: "freeze_panes",
        primitiveName: "freeze_panes",
        args: { sheet: "Sheet1", rows: 2, columns: 3 },
      }],
    };
    await expect(fake.withExcel(() => adapter.reconcileFailedGroup(group))).resolves.toEqual(["committed"]);
  });

  it("does not reconcile a pre-existing same-title chart as the created chart", async () => {
    const operationId = "33333333-3333-4333-8333-333333333333";
    const fake = strictChartIdentityFake(operationId);
    const adapter = createNativeWorkbookPlanAdapter();

    await expect(fake.withExcel(() => adapter.reconcileFailedGroup(failedChartGroup(operationId)))).resolves.toEqual([
      "unknown_within_batch",
    ]);
    expect(fake.requestedNames).not.toEqual(["Revenue"]);
  });

  it("binds the new conditional format rule when an unrelated pre-existing rule remains", async () => {
    const fake = strictConditionalCreatedIdentityFake();
    const adapter = createNativeWorkbookPlanAdapter();
    const group = failedConditionalGroup();

    await expect(fake.withExcel(() => adapter.executeGroup(group, 0))).resolves.toEqual({ status: "committed" });
    await expect(fake.withExcel(() => adapter.verifyMandatoryPostconditions!([
      { operationIndex: 0, kind: "sentinel_unchanged", proof: { operationIndex: 0, operationKind: "conditional_format" } },
    ], group.operations))).resolves.toBe(true);
  });

  it("rejects ambiguous conditional readback when more than one new rule appears", async () => {
    const fake = strictConditionalCreatedIdentityFake(true);
    const adapter = createNativeWorkbookPlanAdapter();
    const group = failedConditionalGroup();

    await expect(fake.withExcel(() => adapter.executeGroup(group, 0))).resolves.toEqual({ status: "committed" });
    await expect(fake.withExcel(() => adapter.verifyMandatoryPostconditions!([
      { operationIndex: 0, kind: "sentinel_unchanged", proof: { operationIndex: 0, operationKind: "conditional_format" } },
    ], group.operations))).resolves.toBe(false);
  });
});
