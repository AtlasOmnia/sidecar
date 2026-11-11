import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("createSheet Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("creates and activates the requested worksheet without a redundant name readback", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const sheet = {
      activate: vi.fn(),
      load: vi.fn()
    };
    const add = vi.fn(() => sheet);
    const context = {
      sync,
      workbook: { worksheets: { add } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const createSheet = excelTools.find(({ schema }) =>
      schema.function.name === "create_sheet"
    );
    const result = JSON.parse(await createSheet!.execute({ name: "Analysis" }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(add).toHaveBeenCalledWith("Analysis");
    expect(sheet.load).not.toHaveBeenCalled();
    expect(sheet.activate).toHaveBeenCalledTimes(1);
    expect(result).toEqual({
      ok: true,
      action: "create_sheet",
      name: "Analysis"
    });
  });
});
