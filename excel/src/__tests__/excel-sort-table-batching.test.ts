import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("sortTable input validation and batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("sorts with a valid non-negative integer column index in one sync", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const range = {
      select: vi.fn()
    };

    const sort = {
      apply: vi.fn()
    };

    const table = {
      sort,
      getRange: vi.fn().mockReturnValue(range)
    };

    const context = {
      sync,
      workbook: {
        tables: {
          getItem: vi.fn().mockReturnValue(table)
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const sortTableTool = excelTools.find(({ schema }) => schema.function.name === "sort_table");
    const result = JSON.parse(await sortTableTool!.execute({ tableName: "Sales", columnIndex: 2, ascending: false }));

    expect(sort.apply).toHaveBeenCalledWith([{ key: 2, ascending: false }]);
    expect(range.select).toHaveBeenCalledTimes(1);
    expect(context.workbook.tables.getItem).toHaveBeenCalledWith("Sales");
    expect(sync).toHaveBeenCalledTimes(1);
    expect(result).toEqual({ ok: true, action: "sort_table", tableName: "Sales", columnIndex: 2, ascending: false });
  });

  it("rejects non-integer columnIndex before Excel.run", () => {
    const excelRun = vi.fn();
    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: excelRun
    };

    const sortTableTool = excelTools.find(({ schema }) => schema.function.name === "sort_table");

    expect(() => sortTableTool!.execute({ tableName: "Sales", columnIndex: 1.5 })).toThrow(
      "Missing or invalid non-negative integer argument \"columnIndex\"."
    );
    expect(excelRun).toHaveBeenCalledTimes(0);
  });

  it("rejects negative columnIndex before Excel.run", () => {
    const excelRun = vi.fn();
    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: excelRun
    };

    const sortTableTool = excelTools.find(({ schema }) => schema.function.name === "sort_table");

    expect(() => sortTableTool!.execute({ tableName: "Sales", columnIndex: -1 })).toThrow(
      "Missing or invalid non-negative integer argument \"columnIndex\"."
    );
    expect(excelRun).toHaveBeenCalledTimes(0);
  });
});