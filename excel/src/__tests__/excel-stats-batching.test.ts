import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("summarize_range_stats Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("derives the response address for strict A1 ranges without loading host metadata", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const range = {
      values: [[10, "2"], ["", null]],
      address: "'Data''s Sheet'!B2:C3",
      load: vi.fn()
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

    const statsTool = excelTools.find(({ schema }) =>
      schema.function.name === "summarize_range_stats"
    );
    const result = JSON.parse(await statsTool!.execute({
      sheet: "Data's Sheet",
      address: "$B$2:$C$3"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).toHaveBeenCalledWith("values");
    expect(range.load).not.toHaveBeenCalledWith(expect.stringContaining("address"));
    expect(result).toEqual({
      action: "summarize_range_stats",
      sheet: "Data's Sheet",
      address: "'Data''s Sheet'!B2:C3",
      count: 4,
      numericCount: 2,
      sum: 12,
      average: 6,
      min: 2,
      max: 10,
      blankCount: 2
    });
  });

  it("retains host address loading for non-A1 inputs", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const range = {
      values: [[1, 2]],
      address: "Data!NamedRange",
      load: vi.fn()
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

    const statsTool = excelTools.find(({ schema }) =>
      schema.function.name === "summarize_range_stats"
    );
    const result = JSON.parse(await statsTool!.execute({
      sheet: "Data",
      address: "NamedRange"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).toHaveBeenCalledWith(["address", "values"]);
    expect(result.address).toBe("Data!NamedRange");
  });
});
