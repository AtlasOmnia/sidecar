import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("findReplace Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("loads isNullObject in a first sync then runs replaceAll in a second sync", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);

    const range = {
      address: "'Data Sheet'!B2:C4",
      isNullObject: false,
      load: vi.fn(),
      replaceAll: vi.fn(() => ({ value: 3 }))
    };
    const sheet = {
      getUsedRangeOrNullObject: vi.fn(() => range)
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

    const tool = excelTools.find(({ schema }) => schema.function.name === "find_replace");
    const result = JSON.parse(await tool!.execute({
      sheet: "Data Sheet",
      find: "foo",
      replace: "bar",
      matchCase: false
    }));

    // Two syncs: isNullObject check + replaceAll
    expect(sync).toHaveBeenCalledTimes(2);
    expect(range.load).toHaveBeenCalledWith("isNullObject");
    expect(range.load).toHaveBeenCalledWith("address");
    expect(range.replaceAll).toHaveBeenCalledWith("foo", "bar", { matchCase: false });
    expect(result).toEqual({
      ok: true,
      action: "find_replace",
      sheet: "Data Sheet",
      address: "B2:C4",
      changed: 3
    });
  });

  it("returns no-op for null-object ranges after syncing isNullObject", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);

    const range = {
      address: "",
      isNullObject: true,
      load: vi.fn(),
      replaceAll: vi.fn()
    };

    const sheet = {
      getUsedRangeOrNullObject: vi.fn(() => range)
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

    const tool = excelTools.find(({ schema }) => schema.function.name === "find_replace");
    const result = JSON.parse(await tool!.execute({
      sheet: "Empty",
      find: "foo",
      replace: "bar",
      matchCase: false
    }));

    // One sync for isNullObject load — returns before replaceAll
    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).toHaveBeenCalledWith("isNullObject");
    expect(range.replaceAll).toHaveBeenCalledTimes(0);
    expect(result).toEqual({
      ok: true,
      action: "find_replace",
      changed: 0,
      message: "Sheet has no used range."
    });
  });
});
