import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("freezePanes Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("activates and freezes a known worksheet in one Office.js batch", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const freezePanes = {
      freezeRows: vi.fn(),
      freezeColumns: vi.fn(),
      unfreeze: vi.fn()
    };
    const sheet = {
      activate: vi.fn(),
      freezePanes
    };
    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn(() => sheet),
          getActiveWorksheet: vi.fn(() => sheet)
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const freezePanesTool = excelTools.find(({ schema }) =>
      schema.function.name === "freeze_panes"
    );
    const result = JSON.parse(await freezePanesTool!.execute({
      sheet: "Data",
      rows: 1,
      columns: 2
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(sheet.activate).toHaveBeenCalledTimes(1);
    expect(freezePanes.freezeRows).toHaveBeenCalledWith(1);
    expect(freezePanes.freezeColumns).toHaveBeenCalledWith(2);
    expect(freezePanes.unfreeze).not.toHaveBeenCalled();
    expect(result).toEqual({
      ok: true,
      action: "freeze_panes",
      sheet: "Data",
      rows: 1,
      columns: 2,
      status: "frozen"
    });
  });
});
