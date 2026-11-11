import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("copySheet Office.js reliability", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("returns a clear error when source and destination names are the same", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const getItem = vi.fn();
    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const copySheetTool = excelTools.find(({ schema }) =>
      schema.function.name === "copy_sheet"
    );
    const result = JSON.parse(await copySheetTool!.execute({
      sourceName: "Data",
      newName: "Data"
    }));

    expect(result).toMatchObject({
      ok: false,
      error: "Source and destination sheet names must be different."
    });
    expect(getItem).not.toHaveBeenCalled();
    expect(sync).not.toHaveBeenCalled();
  });

  it("copies and activates the copied sheet with one sync", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);

    let copiedName = "";
    const copied = {
      set name(value: string) {
        copiedName = value;
      },
      activate: vi.fn()
    };

    const sheet = {
      copy: vi.fn(() => copied)
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

    const copySheetTool = excelTools.find(({ schema }) =>
      schema.function.name === "copy_sheet"
    );
    const result = JSON.parse(await copySheetTool!.execute({
      sourceName: "Data",
      newName: "Archive"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(sheet.copy).toHaveBeenCalledWith("After", sheet);
    expect(copiedName).toBe("Archive");
    expect(copied.activate).toHaveBeenCalledTimes(1);
    expect(result).toEqual({
      ok: true,
      action: "copy_sheet",
      sourceName: "Data",
      newName: "Archive"
    });
  });

  it("returns a user-friendly error when destination name already exists", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);

    const sheet = {
      copy: vi.fn(() => {
        throw new Error("The name 'Archive' already exists. Please type a different name.");
      })
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

    const copySheetTool = excelTools.find(({ schema }) =>
      schema.function.name === "copy_sheet"
    );
    const result = JSON.parse(await copySheetTool!.execute({
      sourceName: "Data",
      newName: "Archive"
    }));

    expect(result).toEqual({
      ok: false,
      error: "A sheet or item with that name already exists. Please choose a different name.",
      name: "Error"
    });
    expect(sync).not.toHaveBeenCalled();
  });
});
