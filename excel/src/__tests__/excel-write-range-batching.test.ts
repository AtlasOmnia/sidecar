import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("writeRange Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("flushes the strict-A1 write mutation in one sync", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const range = {
      rowCount: 2,
      columnCount: 2,
      load: vi.fn(),
      formulas: undefined as unknown,
      select: vi.fn()
    };

    const sheet = {
      getRange: vi.fn(() => range)
    };

    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn(() => sheet)
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const writeRangeTool = excelTools.find(({
      schema
    }) => schema.function.name === "write_range");
    const result = JSON.parse(
      await writeRangeTool!.execute({
        sheet: "Data's Sheet",
        address: "$A$1:$B$2",
        values: [["A", "B"], ["C", "D"]]
      })
    );

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).toHaveBeenCalledTimes(0);
    expect(range.formulas).toEqual([["A", "B"], ["C", "D"]]);
    expect(range.select).toHaveBeenCalledOnce();
    expect(result).toEqual({
      ok: true,
      action: "write_range",
      sheet: "Data's Sheet",
      address: "'Data''s Sheet'!A1:B2",
      rows: 2,
      columns: 2
    });
  });

  it("validates payload dimensions before mutating on strict-A1 mismatch", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    let formulasSet: unknown;

    const range: {
      address: string;
      rowCount: number;
      columnCount: number;
      load: ReturnType<typeof vi.fn>;
      select: ReturnType<typeof vi.fn>;
      formulas?: unknown;
    } = {
      address: "'Data Sheet'!A1:B2",
      rowCount: 2,
      columnCount: 2,
      load: vi.fn(),
      select: vi.fn()
    };

    Object.defineProperty(range, "formulas", {
      configurable: true,
      enumerable: true,
      set(value) {
        formulasSet = value;
      }
    });

    const sheet = {
      getRange: vi.fn(() => range)
    };

    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn(() => sheet)
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const writeRangeTool = excelTools.find(({
      schema
    }) => schema.function.name === "write_range");
    const result = JSON.parse(
      await writeRangeTool!.execute({
        sheet: "Data Sheet",
        address: "A1:B2",
        values: [["A", "B"], ["C", "D"], ["E", "F"]]
      })
    );

    expect(sync).toHaveBeenCalledTimes(0);
    expect(range.load).toHaveBeenCalledTimes(0);
    expect(formulasSet).toBeUndefined();
    expect(range.select).not.toHaveBeenCalled();
    expect(result).toEqual({
      ok: false,
      name: "Error",
      error: "Range dimension mismatch. The target range is 2x2, but values are 3x2."
    });
  });
});
