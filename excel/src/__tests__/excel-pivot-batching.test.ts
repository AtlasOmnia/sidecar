import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("createPivotTable Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("creates a pivot from explicit input addresses in one sync", async () => {
    const events: string[] = [];
    const sync = vi.fn(async () => {
      events.push("sync");
      return undefined;
    });
    const source = {
      address: "A1:C4",
      rowCount: 4,
      columnCount: 3,
      worksheet: { name: "Data" },
      load: vi.fn()
    };
    const targetCell = { address: "G2", load: vi.fn() };
    const pivotTable = { name: "PivotTableFixture", load: vi.fn() };
    const pivotTables = { add: vi.fn(() => pivotTable) };
    const sourceSheet = { getRange: vi.fn(() => source) };
    const targetSheet = {
      getRange: vi.fn(() => targetCell),
      pivotTables,
      activate: vi.fn(() => events.push("activate"))
    };
    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn((name: string) => name === "Data" ? sourceSheet : targetSheet)
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const createPivotTable = excelTools.find(({ schema }) =>
      schema.function.name === "create_pivot_table"
    );
    const result = JSON.parse(await createPivotTable!.execute({
      sheet: "Pivot",
      address: "G2",
      sourceSheet: "Data",
      sourceRange: "A1:C4",
      rowFields: [],
      dataFields: []
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(source.load).not.toHaveBeenCalled();
    expect(targetCell.load).not.toHaveBeenCalled();
    expect(pivotTable.load).not.toHaveBeenCalled();
    expect(targetSheet.activate).toHaveBeenCalledTimes(1);
    expect(pivotTables.add).toHaveBeenCalledWith(
      expect.any(String),
      "Data!A1:C4",
      targetCell
    );
    expect(result.pivotTableName).toMatch(/^PivotTable\d+$/);
    expect(result).toEqual({
      ok: true,
      action: "create_pivot_table",
      pivotTableName: expect.any(String),
      source: "Data!A1:C4",
      targetSheet: "Pivot",
      targetCell: "G2",
      rowFields: [],
      columnFields: [],
      dataFields: []
    });
  });

  it("quotes source sheet names containing spaces and apostrophes", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const source = { address: "A1:C4", load: vi.fn() };
    const targetCell = { address: "G2", load: vi.fn() };
    const pivotTable = { name: "PivotTableFixture" };
    const pivotTables = { add: vi.fn(() => pivotTable) };
    const sourceSheet = { getRange: vi.fn(() => source) };
    const targetSheet = {
      getRange: vi.fn(() => targetCell),
      pivotTables,
      activate: vi.fn()
    };
    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn((name: string) => name === "Data Sheet's" ? sourceSheet : targetSheet)
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const createPivotTable = excelTools.find(({ schema }) =>
      schema.function.name === "create_pivot_table"
    );
    const result = JSON.parse(await createPivotTable!.execute({
      sheet: "Pivot",
      address: "G2",
      sourceSheet: "Data Sheet's",
      sourceRange: "A1:C4",
      rowFields: [],
      dataFields: []
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(pivotTables.add).toHaveBeenCalledWith(
      expect.any(String),
      "'Data Sheet''s'!A1:C4",
      targetCell
    );
    expect(result).toEqual({
      ok: true,
      action: "create_pivot_table",
      pivotTableName: expect.any(String),
      source: "'Data Sheet''s'!A1:C4",
      targetSheet: "Pivot",
      targetCell: "G2",
      rowFields: [],
      columnFields: [],
      dataFields: []
    });
  });
});
