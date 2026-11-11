import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("renameSheet Office.js reliability", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("returns ok without Excel.run when oldName and newName are identical", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);

    let assignedName = "";
    const sheet = {
      set name(value: string) {
        assignedName = value;
      }
    };

    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn(() => sheet)
        }
      }
    };

    let runCalled = false;
    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => {
        runCalled = true;
        return callback(context);
      }
    };

    const renameSheetTool = excelTools.find(({ schema }) =>
      schema.function.name === "rename_sheet"
    );
    const result = JSON.parse(await renameSheetTool!.execute({
      oldName: "Data",
      newName: "Data"
    }));

    // With the short-circuit, this must NOT call Excel.run or context.sync.
    expect(result).toEqual({
      ok: true,
      action: "rename_sheet",
      oldName: "Data",
      newName: "Data"
    });
    expect(runCalled).toBe(false);
    expect(sync).not.toHaveBeenCalled();
  });

  it("renames the sheet with one sync when names differ", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);

    let assignedName = "";
    const sheet = {
      set name(value: string) {
        assignedName = value;
      }
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

    const renameSheetTool = excelTools.find(({ schema }) =>
      schema.function.name === "rename_sheet"
    );
    const result = JSON.parse(await renameSheetTool!.execute({
      oldName: "Data",
      newName: "Archive"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(assignedName).toBe("Archive");
    expect(result).toEqual({
      ok: true,
      action: "rename_sheet",
      oldName: "Data",
      newName: "Archive"
    });
  });
});
