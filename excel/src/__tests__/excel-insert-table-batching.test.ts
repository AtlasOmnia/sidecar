import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("insertTable Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("derives strict-A1 response address without loading table range metadata", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const tableLoad = vi.fn();
    const tableRangeLoad = vi.fn();
    const select = vi.fn();
    const tableRange = {
      address: "Data Sheet's!$A$1:$C$4",
      load: tableRangeLoad,
      select
    };
    const table = {
      name: "TableFixture",
      style: "",
      load: tableLoad,
      getRange: vi.fn(() => tableRange)
    };
    const sheet = {
      tables: { add: vi.fn(() => table) }
    };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const insertTable = excelTools.find(({ schema }) => schema.function.name === "insert_table");
    const result = JSON.parse(await insertTable!.execute({
      sheet: "Data Sheet's",
      address: "$A$1:$C$4",
      hasHeaders: true
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(sheet.tables.add).toHaveBeenCalledWith("$A$1:$C$4", true);
    expect(table.style).toBe("TableStyleMedium2");
    expect(tableLoad).toHaveBeenCalledWith("name");
    expect(tableRangeLoad).not.toHaveBeenCalled();
    expect(select).toHaveBeenCalledOnce();
    expect(result).toEqual({
      ok: true,
      action: "insert_table",
      tableName: "TableFixture",
      address: "'Data Sheet''s'!A1:C4"
    });
  });

  it("retains host address loading for non-A1 range expressions", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const tableRangeLoad = vi.fn();
    const tableRange = {
      address: "Data!NamedRange",
      load: tableRangeLoad,
      select: vi.fn()
    };
    const table = {
      name: "NamedTable",
      style: "",
      load: vi.fn(),
      getRange: vi.fn(() => tableRange)
    };
    const sheet = {
      tables: { add: vi.fn(() => table) }
    };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const insertTable = excelTools.find(({ schema }) => schema.function.name === "insert_table");
    const result = JSON.parse(await insertTable!.execute({
      sheet: "Data",
      address: "NamedRange",
      hasHeaders: false
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(tableRangeLoad).toHaveBeenCalledWith("address");
    expect(result).toEqual({
      ok: true,
      action: "insert_table",
      tableName: "NamedTable",
      address: "Data!NamedRange"
    });
  });
});
