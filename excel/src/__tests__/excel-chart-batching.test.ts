import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("createChart Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("creates and positions a chart with one host sync", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const topLeft = {};
    const bottomRight = {};
    const lastColumn = {
      getOffsetRange: vi.fn()
        .mockReturnValueOnce(topLeft)
        .mockReturnValueOnce(bottomRight)
    };
    const source = {
      address: "A1:C4",
      rowIndex: 0,
      columnIndex: 0,
      columnCount: 3,
      load: vi.fn(),
      getLastColumn: vi.fn(() => lastColumn)
    };
    const chart = {
      title: { text: "" },
      legend: { position: "" },
      setPosition: vi.fn(),
      activate: vi.fn()
    };
    const sheet = {
      getRange: vi.fn(() => source),
      getCell: vi.fn()
        .mockReturnValueOnce(topLeft)
        .mockReturnValueOnce(bottomRight),
      charts: { add: vi.fn(() => chart) }
    };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const createChart = excelTools.find(({ schema }) =>
      schema.function.name === "create_chart"
    );
    const result = JSON.parse(await createChart!.execute({
      sheet: "Data",
      dataRange: "A1:C4",
      chartType: "ColumnClustered",
      title: "Revenue"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(source.load).not.toHaveBeenCalled();
    expect(source.getLastColumn).toHaveBeenCalledTimes(1);
    expect(sheet.charts.add).toHaveBeenCalledWith("ColumnClustered", source, "Auto");
    expect(lastColumn.getOffsetRange).toHaveBeenNthCalledWith(1, 0, 2);
    expect(lastColumn.getOffsetRange).toHaveBeenNthCalledWith(2, 14, 9);
    expect(chart.setPosition).toHaveBeenCalledWith(topLeft, bottomRight);
    expect(chart.activate).toHaveBeenCalledOnce();
    expect(result).toEqual({
      ok: true,
      action: "create_chart",
      sheet: "Data",
      source: "A1:C4",
      chartType: "ColumnClustered",
      title: "Revenue"
    });
  });

  it("returns a sanitized source range for absolute references without loading metadata", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const topLeft = {};
    const bottomRight = {};
    const lastColumn = {
      getOffsetRange: vi.fn()
        .mockReturnValueOnce(topLeft)
        .mockReturnValueOnce(bottomRight)
    };
    const source = {
      rowIndex: 2,
      columnIndex: 1,
      columnCount: 1,
      load: vi.fn(),
      getLastColumn: vi.fn(() => lastColumn)
    };
    const chart = {
      title: { text: "" },
      legend: { position: "" },
      setPosition: vi.fn(),
      activate: vi.fn()
    };
    const sheet = {
      getRange: vi.fn(() => source),
      charts: { add: vi.fn(() => chart) }
    };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const createChart = excelTools.find(({ schema }) =>
      schema.function.name === "create_chart"
    );
    const result = JSON.parse(await createChart!.execute({
      sheet: "Dashboard",
      dataRange: "$A$1:$C$4",
      chartType: "Line",
      title: "Trend"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(source.load).not.toHaveBeenCalled();
    expect(source.getLastColumn).toHaveBeenCalledTimes(1);
    expect(sheet.charts.add).toHaveBeenCalledWith("Line", source, "Auto");
    expect(chart.setPosition).toHaveBeenCalledTimes(1);
    expect(chart.activate).toHaveBeenCalledOnce();
    expect(result).toEqual({
      ok: true,
      action: "create_chart",
      sheet: "Dashboard",
      source: "A1:C4",
      chartType: "Line",
      title: "Trend"
    });
  });
});
