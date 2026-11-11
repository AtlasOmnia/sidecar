import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("autofillFormula Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("uses one sync for strict A1 autofill requests", async () => {
    const events: string[] = [];
    const sync = vi.fn().mockImplementation(async () => {
      events.push("sync");
    });
    const source = {
      formulasR1C1: [["=RC[-1]*2"]],
      formulas: [["=A2*2"]],
      rowCount: 1,
      columnCount: 1,
      load: vi.fn(),
      autoFill: vi.fn(() => events.push("autoFill"))
    };
    const target = {
      address: "B2:B5",
      rowCount: 4,
      columnCount: 1,
      load: vi.fn(),
      select: vi.fn(() => events.push("select"))
    };
    const sheet = {
      getRange: vi.fn((address: string) => address === "B2" ? source : target)
    };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const autofillFormula = excelTools.find(({ schema }) =>
      schema.function.name === "autofill_formula"
    );
    const result = JSON.parse(await autofillFormula!.execute({
      sheet: "Data",
      sourceAddress: "B2",
      targetAddress: "B2:B5"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(source.load).not.toHaveBeenCalled();
    expect(target.load).not.toHaveBeenCalled();
    expect(source.autoFill).toHaveBeenCalledWith(target);
    expect(target.select).toHaveBeenCalledOnce();
    expect(events).toEqual(["autoFill", "select", "sync"]);
    expect(result).toEqual({
      ok: true,
      action: "autofill_formula",
      sheet: "Data",
      source: "B2",
      target: "B2:B5"
    });
  });

  it("falls back to an R1C1 formula copy when strict-A1 autofill is rejected", async () => {
    const events: string[] = [];
    let syncCount = 0;
    const sync = vi.fn().mockImplementation(async () => {
      syncCount += 1;
      events.push(`sync${syncCount}`);
      if (syncCount === 1) {
        throw new Error("autoFill unsupported");
      }
    });
    const source = {
      formulasR1C1: [["=RC[-1]*2"]],
      load: vi.fn(),
      autoFill: vi.fn(() => events.push("autoFill"))
    };
    let targetFormulas: unknown;
    const target = {
      get formulasR1C1(): unknown {
        return targetFormulas;
      },
      set formulasR1C1(value: unknown) {
        events.push("setFormulas");
        targetFormulas = value;
      },
      select: vi.fn(() => events.push("select"))
    };
    const sheet = {
      getRange: vi.fn((address: string) => address === "B2" ? source : target)
    };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const autofillFormula = excelTools.find(({ schema }) =>
      schema.function.name === "autofill_formula"
    );
    const result = JSON.parse(await autofillFormula!.execute({
      sheet: "Data",
      sourceAddress: "B2",
      targetAddress: "B2:B5"
    }));

    expect(sync).toHaveBeenCalledTimes(3);
    expect(source.load).toHaveBeenCalledWith("formulasR1C1");
    expect(target.formulasR1C1).toEqual([
      ["=RC[-1]*2"],
      ["=RC[-1]*2"],
      ["=RC[-1]*2"],
      ["=RC[-1]*2"]
    ]);
    expect(target.select).toHaveBeenCalledTimes(2);
    expect(events.slice(-3)).toEqual(["setFormulas", "select", "sync3"]);
    expect(result).toEqual({
      ok: true,
      action: "autofill_formula",
      sheet: "Data",
      source: "B2",
      target: "B2:B5"
    });
  });
});
