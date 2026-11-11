import { describe, expect, it } from "vitest";
import { createNativeWorkbookPlanAdapter } from "../workbookPlanOfficeAdapter";

describe("native Office workbook plan adapter", () => {
  it("queues one effect group in one Excel.run and does not dispatch one primitive call per operation", async () => {
    const syncs: string[] = [];
    const range = (address: string) => ({
      values: [[null]] as unknown,
      formulas: undefined as unknown,
      load: () => undefined,
      format: {},
      address,
    });
    const sheet = { getRange: range, name: "Sheet1" };
    const context = {
      workbook: { worksheets: { getItem: (name: string) => ({ ...sheet, name }) } },
      sync: async () => { syncs.push("sync"); },
    };
    const previous = (globalThis as Record<string, unknown>).Excel;
    (globalThis as Record<string, unknown>).Excel = { run: async (fn: (ctx: typeof context) => Promise<unknown>) => fn(context) };
    try {
      const adapter = createNativeWorkbookPlanAdapter();
      const result = await adapter.executeGroup({ groupIndex: 0, effect: "write", operations: [
        { operationId: "11111111-1111-4111-8111-111111111111", operationIndex: 0, kind: "write_values", primitiveName: "write_range", args: { sheet: "Sheet1", address: "A1", values: [[1]] } },
        { operationId: "22222222-2222-4222-8222-222222222222", operationIndex: 1, kind: "write_values", primitiveName: "write_range", args: { sheet: "Sheet1", address: "A2", values: [[2]] } },
      ] }, 0);
      expect(result).toEqual({ status: "committed" });
      expect(syncs).toEqual(["sync", "sync"]);
    } finally {
      if (previous === undefined) delete (globalThis as Record<string, unknown>).Excel;
      else (globalThis as Record<string, unknown>).Excel = previous;
    }
  });

  it("executes every compiler-emitted primitive in one Office.run group", async () => {
    const syncs: string[] = [];
    const makeRange = (address: string) => {
      const range: Record<string, unknown> = {
        address,
        values: [["value"]],
        formulas: [["=A1"]],
        numberFormat: [["General"]],
        format: {
          font: {},
          fill: {},
          autofitColumns: () => undefined,
          autofitRows: () => undefined,
        },
        dataValidation: {},
        conditionalFormats: {
          items: [],
          load: () => undefined,
          add: (type: string) => type === "CellValue"
            ? { id: `new-${type}`, cellValue: { rule: {}, format: { fill: {}, font: {} } } }
            : { id: `new-${type}`, textComparison: { rule: {}, format: { fill: {}, font: {} } } },
        },
        load: () => undefined,
        autoFill: () => undefined,
      };
      return range;
    };
    const table = {
      sort: { apply: () => undefined },
      getDataBodyRange: () => ({ values: [[1]], load: () => undefined }),
      columns: { getItemAt: () => ({ filter: { clear: () => undefined, applyValuesFilter: () => undefined, applyCustomFilter: () => undefined } }) },
    };
    const makeSheet = (name: string) => ({
      name,
      getRange: (address: string) => makeRange(address),
      getRanges: (address: string) => ({ address: `${name}!${address}`, areaCount: 1, load: () => undefined }),
      tables: { add: () => table },
      charts: { add: () => ({ title: { text: "Chart" } }), getItemOrNullObject: () => ({ isNullObject: true, load: () => undefined }) },
      comments: {
        items: [{ content: "note", load: () => undefined, getLocation: () => ({ address: `${name}!A1`, load: () => undefined }) }],
        getCount: () => ({ value: 1 }),
        load: () => undefined,
        add: () => ({ content: "note", load: () => undefined, getLocation: () => ({ address: `${name}!A1`, load: () => undefined }) }),
      },
      freezePanes: { freezeRows: () => undefined, freezeColumns: () => undefined },
      copy: () => makeSheet(`${name} Copy`),
    });
    const sheet = makeSheet("Sheet1");
    const context = {
      workbook: {
        worksheets: { getItem: () => sheet, add: () => sheet },
        tables: { getItem: () => table, getItemOrNullObject: () => ({ isNullObject: true, load: () => undefined }) },
        comments: { add: () => undefined },
        application: { calculate: () => undefined },
      },
      sync: async () => { syncs.push("sync"); },
    };
    const previous = (globalThis as Record<string, unknown>).Excel;
    (globalThis as Record<string, unknown>).Excel = { run: async (fn: (ctx: typeof context) => Promise<unknown>) => fn(context) };
    try {
      const adapter = createNativeWorkbookPlanAdapter();
      const operations = [
        { kind: "write_values", args: { sheet: "Sheet1", address: "A1", values: [[1]] } },
        { kind: "write_formula", args: { sheet: "Sheet1", address: "A1", values: [["=1"]] } },
        { kind: "fill_formula", args: { sheet: "Sheet1", sourceAddress: "A1", targetAddress: "A1:A2" } },
        { kind: "format_range", args: { sheet: "Sheet1", address: "A1", format: { bold: true } } },
        { kind: "create_table", args: { sheet: "Sheet1", address: "A1:B2", hasHeaders: true } },
        { kind: "sort_table", args: { tableName: "Table1", columnIndex: 0, ascending: true } },
        { kind: "filter_table", args: { tableName: "Table1", columnIndex: 0, values: ["x"] } },
        { kind: "reset_table_filter", args: { tableName: "Table1", columnIndex: 0 } },
        { kind: "create_chart", args: { sheet: "Sheet1", dataRange: "A1:B2", chartType: "ColumnClustered", title: "Chart" } },
        { kind: "set_validation_list", args: { sheet: "Sheet1", address: "A1", source: "x,y" } },
        { kind: "conditional_format", args: { sheet: "Sheet1", address: "A1", ruleType: "containsText", operator: "contains", text: "ok" } },
        { kind: "freeze_panes", args: { sheet: "Sheet1", rows: 1, columns: 1 } },
        { kind: "create_sheet", args: { name: "New" } },
        { kind: "rename_sheet", args: { oldName: "Sheet1", newName: "Renamed" } },
        { kind: "copy_sheet", args: { sourceName: "Sheet1", newName: "Copy" } },
        { kind: "add_comment", args: { sheet: "Sheet1", address: "A1", text: "note" } },
        { kind: "calculate", args: {} },
      ] as const;
      const result = await adapter.executeGroup({
        groupIndex: 0,
        effect: "write",
        operations: operations.map((operation, operationIndex) => ({
          operationId: `${String(operationIndex + 1).padStart(8, "0")}-1111-4111-8111-111111111111`,
          operationIndex,
          kind: operation.kind,
          primitiveName: "test",
          args: operation.args,
        })),
      }, 0);
      expect(result).toEqual({ status: "committed" });
      expect(syncs.length).toBe(5);
    } finally {
      if (previous === undefined) delete (globalThis as Record<string, unknown>).Excel;
      else (globalThis as Record<string, unknown>).Excel = previous;
    }
  });

  it("reads table, chart, filter, comment, and calculation verification rules", async () => {
    const tableRange = { values: [["first"], ["second"]], load: () => undefined };
    const table = {
      isNullObject: false,
      load: () => undefined,
      getRange: () => tableRange,
      getDataBodyRange: () => tableRange,
      columns: { getItemAt: () => ({ filter: { criteria: { filterOn: "Values", values: ["x"] }, load: () => undefined } }) },
    };
    const comment = { content: "note", load: () => undefined, getLocation: () => ({ address: "Sheet1!A1", load: () => undefined }) };
    const chart = { isNullObject: false, title: { text: "Chart" }, load: () => undefined };
    const sheet = {
      name: "Sheet1",
      getRange: () => ({ values: [["sentinel"]], formulas: [["=1"]], numberFormat: [["General"]], load: () => undefined }),
      charts: { load: () => undefined, items: [chart] },
      comments: { items: [comment], getCount: () => ({ value: 1 }), load: () => undefined },
    };
    const context = {
      workbook: {
        worksheets: { getItem: () => sheet, getItemOrNullObject: () => ({ isNullObject: false, load: () => undefined }) },
        tables: { getItem: () => table, getItemOrNullObject: () => table },
        comments: { items: [comment], load: () => undefined, add: () => comment },
        application: { calculationState: "Done", load: () => undefined },
      },
      sync: async () => undefined,
    };
    const previous = (globalThis as Record<string, unknown>).Excel;
    (globalThis as Record<string, unknown>).Excel = { run: async (fn: (ctx: typeof context) => Promise<unknown>) => fn(context) };
    try {
      const adapter = createNativeWorkbookPlanAdapter();
      const result = await adapter.verifyPlan([
        { kind: "table_exists", tableName: "Table1" },
        { kind: "table_row_order", tableName: "Table1", expectedOrder: ["first", "second"] },
        { kind: "table_filter_state", tableName: "Table1", columnIndex: 0, state: "filtered", values: ["x"] },
        { kind: "chart_exists", sheet: "Sheet1", title: "Chart" },
        { kind: "comment_exists", sheet: "Sheet1", address: "A1", text: "note" },
        { kind: "calculation_completed" },
      ]);
      expect(result.ok).toBe(true);
      expect(result.counts.verified).toBe(6);
    } finally {
      if (previous === undefined) delete (globalThis as Record<string, unknown>).Excel;
      else (globalThis as Record<string, unknown>).Excel = previous;
    }
  });
});
