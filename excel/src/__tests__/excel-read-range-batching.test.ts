import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

interface MockRange {
  address: string;
  rowCount: number;
  columnCount: number;
  values: unknown[];
  formulas: unknown[];
  numberFormat: unknown[];
  load: ReturnType<typeof vi.fn>;
  getCell?: ReturnType<typeof vi.fn>;
  getResizedRange?: ReturnType<typeof vi.fn>;
}

describe("readRange Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("derives a strict-A1 response address without loading host address metadata", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);

    const range: MockRange = {
      get address(): string {
        throw new Error("PropertyNotLoaded: address");
      },
      rowCount: 2,
      columnCount: 3,
      values: [[1, 2, 3], [4, 5, 6]],
      formulas: [[null, null, null], ["=A1+B1", null, null]],
      numberFormat: [["General", "General", "General"], ["General", "General", "General"]],
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

    const readRangeTool = excelTools.find(({ schema }) =>
      schema.function.name === "read_range"
    );
    const result = JSON.parse(await readRangeTool!.execute({
      sheet: "Data's Sheet",
      address: "$B$2:$D$3"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(sheet.getRange).toHaveBeenCalledWith("$B$2:$D$3");
    expect(range.load).toHaveBeenCalledWith(["rowCount", "columnCount", "values", "formulas", "numberFormat"]);
    expect(result).toEqual({
      sheet: "Data's Sheet",
      address: "'Data''s Sheet'!B2:D3",
      rowCount: 2,
      columnCount: 3,
      values: [[1, 2, 3], [4, 5, 6]],
      formulas: [[null, null, null], ["=A1+B1", null, null]],
      numberFormat: [["General", "General", "General"], ["General", "General", "General"]]
    });
  });

  it("retains host address loading for non-A1 expressions", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const range: MockRange = {
      address: "Data!NamedRange",
      rowCount: 1,
      columnCount: 2,
      values: [[1, 2]],
      formulas: [[null, null]],
      numberFormat: [["General", "General"]],
      load: vi.fn()
    };
    const sheet = { getRange: vi.fn(() => range) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const readRangeTool = excelTools.find(({ schema }) =>
      schema.function.name === "read_range"
    );
    const result = JSON.parse(await readRangeTool!.execute({
      sheet: "Data",
      address: "NamedRange"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).toHaveBeenCalledWith(["address", "rowCount", "columnCount", "values", "formulas", "numberFormat"]);
    expect(result).toMatchObject({
      sheet: "Data",
      address: "Data!NamedRange",
      rowCount: 1,
      columnCount: 2,
      values: [[1, 2]]
    });
  });

  it("uses a bounded preview batch for large strict A1 reads", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);

    const preview: MockRange = {
      get address(): string {
        throw new Error("PropertyNotLoaded: address");
      },
      get rowCount(): number {
        throw new Error("PropertyNotLoaded: rowCount");
      },
      get columnCount(): number {
        throw new Error("PropertyNotLoaded: columnCount");
      },
      values: [["A", "B", "C", "D"]],
      formulas: [[null, null, null, null]],
      numberFormat: [["General", "General", "General", "General"]],
      load: vi.fn()
    };

    const anchor = {
      getResizedRange: vi.fn(() => preview)
    };

    const range = {
      address: "Data!A1:D5000",
      rowCount: 5000,
      columnCount: 4,
      values: [[]],
      formulas: [[]],
      numberFormat: [[]],
      load: vi.fn(),
      getCell: vi.fn(() => anchor)
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

    const readRangeTool = excelTools.find(({ schema }) =>
      schema.function.name === "read_range"
    );
    const result = JSON.parse(await readRangeTool!.execute({
      sheet: "Data",
      address: "A1:D5000"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).not.toHaveBeenCalledWith("address");
    expect(range.getCell).toHaveBeenCalledWith(0, 0);
    expect(anchor.getResizedRange).toHaveBeenCalledWith(1250 - 1, range.columnCount - 1);
    expect(preview.load).toHaveBeenCalledWith(["values", "formulas", "numberFormat"]);
    expect(result).toMatchObject({
      sheet: "Data",
      address: "Data!A1:D5000",
      rowCount: 5000,
      columnCount: 4,
      truncated: true,
      note: "Range has 20000 cells. Returning the first 1250 rows only.",
      previewAddress: "Data!A1:D1250"
    });
  });

  it("skips redundant rowCount/columnCount on non-A1 bounded preview", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);

    const preview: MockRange = {
      address: "Data!A1:D2000",
      rowCount: 2000,
      columnCount: 4,
      values: [["X", "Y", "Z", "W"]],
      formulas: [[null, null, null, null]],
      numberFormat: [["General", "General", "General", "General"]],
      load: vi.fn()
    };

    const anchor = {
      getResizedRange: vi.fn(() => preview)
    };

    const range: MockRange = {
      address: "Data!NamedRange",
      rowCount: 8000,
      columnCount: 4,
      values: [[]],
      formulas: [[]],
      numberFormat: [[]],
      load: vi.fn(),
      getCell: vi.fn(() => anchor)
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

    const readRangeTool = excelTools.find(({ schema }) =>
      schema.function.name === "read_range"
    );
    const result = JSON.parse(await readRangeTool!.execute({
      sheet: "Data",
      address: "NamedRange"
    }));

    // Non-A1 large reads require 2 syncs: one to load the full range metadata,
    // one to load the bounded preview payload.
    expect(sync).toHaveBeenCalledTimes(2);

    // First sync: load full range metadata (address + dimensions + values for size check)
    expect(range.load).toHaveBeenCalledWith([
      "address", "rowCount", "columnCount", "values", "formulas", "numberFormat"
    ]);

    // Preview loads only address + content; rowCount/columnCount are derived locally.
    expect(preview.load).toHaveBeenCalledWith([
      "address", "values", "formulas", "numberFormat"
    ]);

    // previewRows = floor(5000 / 4) = 1250
    expect(anchor.getResizedRange).toHaveBeenCalledWith(1250 - 1, range.columnCount - 1);

    expect(result).toMatchObject({
      sheet: "Data",
      address: "Data!NamedRange",
      rowCount: 8000,
      columnCount: 4,
      truncated: true,
      note: "Range has 32000 cells. Returning the first 1250 rows only.",
      previewAddress: "Data!A1:D2000",
      values: [["X", "Y", "Z", "W"]],
      formulas: [[null, null, null, null]],
      numberFormat: [["General", "General", "General", "General"]]
    });
  });
});
