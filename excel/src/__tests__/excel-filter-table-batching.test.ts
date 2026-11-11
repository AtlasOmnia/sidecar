import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("filterTable input validation and batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("applies a values filter in a single sync", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const select = vi.fn();

    const filter = {
      applyValuesFilter: vi.fn(),
      applyCustomFilter: vi.fn(),
      clear: vi.fn()
    };

    const range = {
      select
    };

    const table = {
      columns: {
        getItemAt: vi.fn().mockReturnValue({ filter })
      },
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

    const run = vi.fn(async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context));
    (globalThis as unknown as { Excel: unknown }).Excel = {
      run
    };

    const filterTableTool = excelTools.find(({ schema }) =>
      schema.function.name === "filter_table"
    );

    const result = JSON.parse(await filterTableTool!.execute({
      tableName: "Sales",
      columnIndex: 1,
      values: ["North", "West"]
    }));

    expect(run).toHaveBeenCalledTimes(1);
    expect(sync).toHaveBeenCalledTimes(1);
    expect(context.workbook.tables.getItem).toHaveBeenCalledWith("Sales");
    expect(table.columns.getItemAt).toHaveBeenCalledWith(1);
    expect(filter.applyValuesFilter).toHaveBeenCalledWith(["North", "West"]);
    expect(filter.applyCustomFilter).not.toHaveBeenCalled();
    expect(filter.clear).not.toHaveBeenCalled();
    expect(range.select).toHaveBeenCalledOnce();
    expect(result).toEqual({
      ok: true,
      action: "filter_table",
      tableName: "Sales",
      columnIndex: 1
    });
  });

  it("rejects a negative columnIndex before Excel.run", async () => {
    const run = vi.fn(async (callback: (requestContext: unknown) => Promise<unknown>) => callback({
      workbook: {
        tables: {
          getItem: vi.fn()
        }
      }
    }));

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run
    };

    const filterTableTool = excelTools.find(({ schema }) =>
      schema.function.name === "filter_table"
    );

    expect(() => filterTableTool!.execute({
      tableName: "Sales",
      columnIndex: -1
    })).toThrow("Missing or invalid non-negative integer argument \"columnIndex\".");
    expect(run).not.toHaveBeenCalled();
  });
});
