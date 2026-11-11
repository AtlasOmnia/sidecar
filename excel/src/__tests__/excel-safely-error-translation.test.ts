import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("safelyExcel error translation", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("passes through domain-specific errors containing 'not found' without mistranslation", async () => {
    // When a pivot field or other domain error contains "not found",
    // safelyExcel must NOT translate it to "Sheet not found."
    const sync = vi.fn().mockResolvedValue(undefined);
    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn(() => {
            throw new Error(
              'Row field "Month" not found in source data or cannot be added. ' +
              "Available hierarchies may differ from column headers."
            );
          })
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) =>
        callback(context)
    };

    const renameSheet = excelTools.find(
      ({ schema }) => schema.function.name === "rename_sheet"
    );
    const result = JSON.parse(
      await renameSheet!.execute({ oldName: "Data", newName: "Archive" })
    );

    // The original domain error must pass through — it must NOT be
    // overwritten with "Sheet not found."
    expect(result.ok).toBe(false);
    expect(result.error).toContain("Month");
    expect(result.error).toContain("not found in source data");
    expect(result.error).not.toContain("Sheet not found");
  });

  it("still translates genuine ItemNotFound to a user-friendly sheet-not-found message", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn(() => {
            throw {
              code: "ItemNotFound",
              message: "The requested worksheet does not exist.",
              name: "OfficeExtension.Error"
            };
          })
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) =>
        callback(context)
    };

    const renameSheet = excelTools.find(
      ({ schema }) => schema.function.name === "rename_sheet"
    );
    const result = JSON.parse(
      await renameSheet!.execute({ oldName: "Data", newName: "Archive" })
    );

    expect(result.ok).toBe(false);
    expect(result.error).toContain("Sheet not found");
    expect(result.name).toBe("OfficeExtension.Error");
  });

  it("still translates name-conflict errors to a user-friendly message", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn(() => {
            throw new Error(
              "The name 'Archive' already exists. Please type a different name."
            );
          })
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) =>
        callback(context)
    };

    const renameSheet = excelTools.find(
      ({ schema }) => schema.function.name === "rename_sheet"
    );
    const result = JSON.parse(
      await renameSheet!.execute({ oldName: "Data", newName: "Archive" })
    );

    expect(result.ok).toBe(false);
    expect(result.error).toContain(
      "A sheet or item with that name already exists"
    );
    expect(sync).not.toHaveBeenCalled();
  });
});
