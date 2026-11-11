import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { getWorkbookContext } from "../excelTools";

interface MockSheet {
  name: string;
  load: ReturnType<typeof vi.fn>;
  getUsedRangeOrNullObject: () => MockUsedRange;
}

interface MockUsedRange {
  isNullObject: boolean;
  address: string;
  rowCount: number;
  columnCount: number;
  load: ReturnType<typeof vi.fn>;
}

function makeUsedRange(address: string, rowCount: number, columnCount: number): MockUsedRange {
  return {
    isNullObject: false,
    address,
    rowCount,
    columnCount,
    load: vi.fn()
  };
}

function makeSheet(name: string, usedRange: MockUsedRange): MockSheet {
  return {
    name,
    load: vi.fn(),
    getUsedRangeOrNullObject: () => usedRange
  };
}

describe("getWorkbookContext Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("loads sheet metadata in a first sync then used ranges in a second sync", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const sheets = [
      makeSheet("Sheet1", makeUsedRange("Sheet1!A1:B4", 4, 2)),
      makeSheet("Sheet2", makeUsedRange("Sheet2!C2:F8", 7, 4)),
      makeSheet("Summary", makeUsedRange("Summary!A1:D3", 3, 4)),
      makeSheet("Empty", {
        isNullObject: true,
        address: "",
        rowCount: 0,
        columnCount: 0,
        load: vi.fn()
      })
    ];
    const activeSheet = sheets[0];
    const selection = { address: "Sheet1!B2", load: vi.fn() };
    const worksheets = {
      items: sheets,
      load: vi.fn(),
      getActiveWorksheet: () => activeSheet
    };
    const context = {
      sync,
      workbook: {
        worksheets,
        getSelectedRange: () => selection
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const result = JSON.parse(await getWorkbookContext());

    // Two syncs: one for sheet metadata, one for used ranges
    expect(sync).toHaveBeenCalledTimes(2);
    expect(result).toEqual({
      activeSheet: "Sheet1",
      selection: "Sheet1!B2",
      sheets: [
        { name: "Sheet1", usedRange: { address: "A1:B4", rowCount: 4, columnCount: 2 } },
        { name: "Sheet2", usedRange: { address: "C2:F8", rowCount: 7, columnCount: 4 } },
        { name: "Summary", usedRange: { address: "A1:D3", rowCount: 3, columnCount: 4 } },
        { name: "Empty", usedRange: null }
      ]
    });
  });
});
