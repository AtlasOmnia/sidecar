import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("clearRange Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("returns deterministic quoted response without loading range metadata", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const range = {
      address: "'Data Sheet'!$A$1:$B$2",
      clear: vi.fn(),
      select: vi.fn(),
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

    const clearRangeTool = excelTools.find(({ schema }) => schema.function.name === "clear_range");
    const result = JSON.parse(
      await clearRangeTool!.execute({
        sheet: "Data's Sheet",
        address: "A1:B2"
      })
    );

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.clear).toHaveBeenCalledWith("All");
    expect(range.select).toHaveBeenCalledOnce();
    expect(range.load).not.toHaveBeenCalled();
    expect(result).toEqual({
      ok: true,
      action: "clear_range",
      sheet: "Data's Sheet",
      address: "'Data''s Sheet'!A1:B2"
    });
  });

  it("strips sheet prefixes from qualified addresses", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const range = {
      address: "'Data's Sheet'!$C$3:$D$4",
      clear: vi.fn(),
      select: vi.fn(),
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

    const clearRangeTool = excelTools.find(({ schema }) => schema.function.name === "clear_range");
    const result = JSON.parse(
      await clearRangeTool!.execute({
        sheet: "Data's Sheet",
        address: "'Data's Sheet'!$C$3:$D$4"
      })
    );

    expect(range.load).not.toHaveBeenCalled();
    expect(result.address).toBe("'Data''s Sheet'!C3:D4");
  });
});
