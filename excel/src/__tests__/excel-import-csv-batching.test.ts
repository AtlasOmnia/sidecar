import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("importCsvToSheet Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("derives the final address for strict-A1 CSV imports without host address loading", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const headerRange = { format: { font: { bold: false } } };
    const writeRange = {
      address: "HostNormalized!B2:C3",
      values: [] as unknown[][],
      load: vi.fn(),
      getCell: vi.fn(() => ({ getResizedRange: vi.fn(() => headerRange) }))
    };
    const target = {
      load: vi.fn(),
      getResizedRange: vi.fn(() => writeRange)
    };
    const sheet = { getRange: vi.fn(() => target) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const importCsvToSheet = excelTools.find(({ schema }) =>
      schema.function.name === "import_csv_to_sheet"
    );
    const result = JSON.parse(await importCsvToSheet!.execute({
      sheet: "Data's Sheet",
      address: "$B$2",
      csv: "Name,Amount\nAda,10",
      hasHeaders: true
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(target.load).not.toHaveBeenCalled();
    expect(target.getResizedRange).toHaveBeenCalledWith(1, 1);
    expect(writeRange.load).not.toHaveBeenCalled();
    expect(writeRange.values).toEqual([["Name", "Amount"], ["Ada", "10"]]);
    expect(headerRange.format.font.bold).toBe(true);
    expect(result).toEqual({
      ok: true,
      action: "import_csv_to_sheet",
      sheet: "Data's Sheet",
      address: "'Data''s Sheet'!B2:C3",
      rows: 2,
      columns: 2,
      hasHeaders: true,
      parseErrors: []
    });
  });

  it("preserves an interior blank CSV row and the following record position", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const writeRange = {
      address: "Data!A1:B3",
      values: [] as unknown[][],
      load: vi.fn()
    };
    const target = {
      getResizedRange: vi.fn(() => writeRange)
    };
    const sheet = { getRange: vi.fn(() => target) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const importCsvToSheet = excelTools.find(({ schema }) =>
      schema.function.name === "import_csv_to_sheet"
    );
    const result = JSON.parse(await importCsvToSheet!.execute({
      sheet: "Data",
      address: "A1",
      csv: "a,x\n,\nb,y\n",
      hasHeaders: false
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(target.getResizedRange).toHaveBeenCalledWith(2, 1);
    expect(writeRange.values).toEqual([["a", "x"], ["", ""], ["b", "y"]]);
    expect(result).toMatchObject({
      ok: true,
      action: "import_csv_to_sheet",
      sheet: "Data",
      address: "Data!A1:B3",
      rows: 3,
      columns: 2,
      hasHeaders: false
    });
    expect(result.parseErrors).toEqual([
      expect.stringContaining("Unable to auto-detect delimiting character")
    ]);
  });

  it("retains host address loading for non-A1 CSV targets", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const writeRange = {
      address: "Data!NamedRange",
      values: [] as unknown[][],
      load: vi.fn()
    };
    const target = {
      getResizedRange: vi.fn(() => writeRange)
    };
    const sheet = { getRange: vi.fn(() => target) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const importCsvToSheet = excelTools.find(({ schema }) =>
      schema.function.name === "import_csv_to_sheet"
    );
    const result = JSON.parse(await importCsvToSheet!.execute({
      sheet: "Data",
      address: "NamedRange",
      csv: "a,b",
      hasHeaders: false
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(writeRange.load).toHaveBeenCalledWith("address");
    expect(result).toMatchObject({
      ok: true,
      action: "import_csv_to_sheet",
      sheet: "Data",
      address: "Data!NamedRange",
      rows: 1,
      columns: 2,
      hasHeaders: false,
      parseErrors: []
    });
  });

  // ── Regression: trailing blank row wider than data ────────────────────────
  //
  // "a,b\n,,\n" produces raw PapaParse data [["a","b"],["","",""],[""]]
  // where the middle trailing row is 3 columns wide while the real CSV is
  // only 2 columns. After terminal blank-row removal, totalCols must be 2
  // so the write range does not spill into column C.
  it("discarded trailing blank row width does not inflate write range", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const writeRange = {
      address: "Data!A1:B1",
      values: [] as unknown[][],
      load: vi.fn()
    };
    const target = {
      getResizedRange: vi.fn(() => writeRange)
    };
    const sheet = { getRange: vi.fn(() => target) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const importCsvToSheet = excelTools.find(({ schema }) =>
      schema.function.name === "import_csv_to_sheet"
    );
    const result = JSON.parse(await importCsvToSheet!.execute({
      sheet: "Data",
      address: "A1",
      csv: "a,b\n,,\n",
      hasHeaders: false
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    // getResizedRange(rows-1, cols-1) = (0, 1) for 1 row × 2 cols
    // If totalCols were inflated to 3, this would be (0, 2)
    expect(target.getResizedRange).toHaveBeenCalledWith(0, 1);
    expect(writeRange.values).toEqual([["a", "b"]]);
    expect(result).toMatchObject({
      ok: true,
      action: "import_csv_to_sheet",
      sheet: "Data",
      rows: 1,
      columns: 2,
      hasHeaders: false
    });
  });
});
