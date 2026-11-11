import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("setDataValidation Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("retains host address loading for non-A1 validation requests", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    let addressLoaded = false;
    const validation = {
      ignoreBlanks: false,
      rule: undefined as unknown,
      prompt: undefined as unknown
    };
    const range = {
      load: vi.fn((property: unknown) => {
        if (property === "address") {
          addressLoaded = true;
        }
      }),
      get address() {
        if (!addressLoaded) {
          throw new Error("PropertyNotLoaded: address");
        }
        return "Data!NamedRange";
      },
      dataValidation: validation,
      select: vi.fn()
    };
    const sheet = { getRange: vi.fn(() => range) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const setDataValidation = excelTools.find(({ schema }) =>
      schema.function.name === "set_data_validation"
    );
    const result = JSON.parse(await setDataValidation!.execute({
      sheet: "Data",
      address: "NamedRange",
      type: "list",
      source: "Yes,No"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).toHaveBeenCalledWith("address");
    expect(validation.ignoreBlanks).toBe(true);
    expect(validation.rule).toEqual({
      list: {
        inCellDropDown: true,
        source: "Yes,No"
      }
    });
    expect(range.select).toHaveBeenCalledOnce();
    expect(result).toEqual({
      ok: true,
      action: "set_data_validation",
      sheet: "Data",
      address: "Data!NamedRange",
      type: "list"
    });
  });

  it("derives a strict-A1 response address without loading host metadata", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const validation = {
      ignoreBlanks: false,
      rule: undefined as unknown,
      prompt: undefined as unknown
    };
    const range = {
      load: vi.fn(),
      get address() {
        throw new Error("PropertyNotLoaded: address");
      },
      dataValidation: validation,
      select: vi.fn()
    };
    const sheet = { getRange: vi.fn(() => range) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const setDataValidation = excelTools.find(({ schema }) =>
      schema.function.name === "set_data_validation"
    );
    const result = JSON.parse(await setDataValidation!.execute({
      sheet: "Data's Sheet",
      address: "$B$2:$B$20",
      type: "list",
      source: "Yes,No"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).not.toHaveBeenCalled();
    expect(validation.ignoreBlanks).toBe(true);
    expect(validation.rule).toEqual({
      list: {
        inCellDropDown: true,
        source: "Yes,No"
      }
    });
    expect(range.select).toHaveBeenCalledOnce();
    expect(result).toEqual({
      ok: true,
      action: "set_data_validation",
      sheet: "Data's Sheet",
      address: "'Data''s Sheet'!B2:B20",
      type: "list"
    });
  });
});
