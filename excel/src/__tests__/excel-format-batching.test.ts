import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("formatRange Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("retains host address loading for non-A1 formatting requests", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const range = {
      address: "Data!NamedRange",
      rowCount: 3,
      columnCount: 2,
      load: vi.fn(),
      select: vi.fn(),
      format: {
        font: { bold: false, italic: false, color: "" },
        fill: { color: "" },
        autofitColumns: vi.fn(),
        autofitRows: vi.fn(),
        borders: { getItem: vi.fn() }
      }
    };
    const sheet = { getRange: vi.fn(() => range) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const formatRange = excelTools.find(({ schema }) =>
      schema.function.name === "format_range"
    );
    const result = JSON.parse(await formatRange!.execute({
      sheet: "Data",
      address: "NamedRange",
      format: { bold: true, fillColor: "#ffff00", columnAutofit: false }
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).toHaveBeenCalledWith("address");
    expect(range.format.font.bold).toBe(true);
    expect(range.format.fill.color).toBe("#ffff00");
    expect(range.select).toHaveBeenCalledOnce();
    expect(result).toEqual({
      ok: true,
      action: "format_range",
      sheet: "Data",
      address: "Data!NamedRange"
    });
  });

  it("derives a strict-A1 response address without loading host metadata", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const range = {
      address: "host-normalized-address",
      load: vi.fn(),
      select: vi.fn(),
      format: {
        font: { bold: false, italic: false, color: "" },
        fill: { color: "" },
        autofitColumns: vi.fn(),
        autofitRows: vi.fn(),
        borders: { getItem: vi.fn() }
      }
    };
    const sheet = { getRange: vi.fn(() => range) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const formatRange = excelTools.find(({ schema }) =>
      schema.function.name === "format_range"
    );
    const result = JSON.parse(await formatRange!.execute({
      sheet: "Data's Sheet",
      address: "$B$2:$C$4",
      format: { bold: true, fillColor: "#ffff00", columnAutofit: false }
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).not.toHaveBeenCalled();
    expect(range.format.font.bold).toBe(true);
    expect(range.format.fill.color).toBe("#ffff00");
    expect(range.select).toHaveBeenCalledOnce();
    expect(result).toEqual({
      ok: true,
      action: "format_range",
      sheet: "Data's Sheet",
      address: "'Data''s Sheet'!B2:C4"
    });
  });

  it("derives number-format dimensions from an A1 range without address metadata", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const range = {
      address: "host-normalized-address",
      rowCount: 3,
      columnCount: 2,
      numberFormat: undefined as unknown,
      load: vi.fn(),
      select: vi.fn(),
      format: {
        font: { bold: false, italic: false, color: "" },
        fill: { color: "" },
        autofitColumns: vi.fn(),
        autofitRows: vi.fn(),
        borders: { getItem: vi.fn() }
      }
    };
    const sheet = { getRange: vi.fn(() => range) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const formatRange = excelTools.find(({ schema }) =>
      schema.function.name === "format_range"
    );
    const result = JSON.parse(await formatRange!.execute({
      sheet: "Data Sheet",
      address: "$B$2:$C$4",
      format: { numberFormat: "0.00", columnAutofit: false }
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).not.toHaveBeenCalled();
    expect(range.numberFormat).toEqual([
      ["0.00", "0.00"],
      ["0.00", "0.00"],
      ["0.00", "0.00"]
    ]);
    expect(range.select).toHaveBeenCalledOnce();
    expect(result).toEqual({
      ok: true,
      action: "format_range",
      sheet: "Data Sheet",
      address: "'Data Sheet'!B2:C4"
    });
  });
});

// ---------------------------------------------------------------------------
// C1 behavior: autofit opt-in, CenterAcrossSelection precondition, alignment
// ---------------------------------------------------------------------------

describe("formatRange C1: autofit opt-in behavior", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("omitted columnAutofit does not call autofitColumns", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const autofitColumns = vi.fn();
    const autofitRows = vi.fn();
    const range = {
      address: "Sheet1!A1",
      load: vi.fn(),
      select: vi.fn(),
      format: {
        font: { bold: false, italic: false, color: "" },
        fill: { color: "" },
        autofitColumns,
        autofitRows,
        borders: { getItem: vi.fn() }
      }
    };
    const sheet = { getRange: vi.fn(() => range) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };
    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (cb: (ctx: unknown) => Promise<unknown>) => cb(context)
    };

    const formatRange = excelTools.find(({ schema }) =>
      schema.function.name === "format_range"
    );
    await formatRange!.execute({
      sheet: "Sheet1",
      address: "A1",
      format: { bold: true }
    });

    expect(autofitColumns).not.toHaveBeenCalled();
    expect(autofitRows).not.toHaveBeenCalled();
  });

  it("columnAutofit: false does not call autofitColumns", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const autofitColumns = vi.fn();
    const autofitRows = vi.fn();
    const range = {
      address: "Sheet1!A1",
      load: vi.fn(),
      select: vi.fn(),
      format: {
        font: { bold: false, italic: false, color: "" },
        fill: { color: "" },
        autofitColumns,
        autofitRows,
        borders: { getItem: vi.fn() }
      }
    };
    const sheet = { getRange: vi.fn(() => range) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };
    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (cb: (ctx: unknown) => Promise<unknown>) => cb(context)
    };

    const formatRange = excelTools.find(({ schema }) =>
      schema.function.name === "format_range"
    );
    await formatRange!.execute({
      sheet: "Sheet1",
      address: "A1",
      format: { bold: true, columnAutofit: false }
    });

    expect(autofitColumns).not.toHaveBeenCalled();
    expect(autofitRows).not.toHaveBeenCalled();
  });

  it("columnAutofit: true calls autofitColumns exactly once", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const autofitColumns = vi.fn();
    const autofitRows = vi.fn();
    const range = {
      address: "Sheet1!A1",
      load: vi.fn(),
      select: vi.fn(),
      format: {
        font: { bold: false, italic: false, color: "" },
        fill: { color: "" },
        autofitColumns,
        autofitRows,
        borders: { getItem: vi.fn() }
      }
    };
    const sheet = { getRange: vi.fn(() => range) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };
    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (cb: (ctx: unknown) => Promise<unknown>) => cb(context)
    };

    const formatRange = excelTools.find(({ schema }) =>
      schema.function.name === "format_range"
    );
    await formatRange!.execute({
      sheet: "Sheet1",
      address: "A1",
      format: { bold: true, columnAutofit: true }
    });

    expect(autofitColumns).toHaveBeenCalledTimes(1);
    expect(autofitRows).not.toHaveBeenCalled();
  });

  it("rowAutofit: true calls autofitRows exactly once", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const autofitColumns = vi.fn();
    const autofitRows = vi.fn();
    const range = {
      address: "Sheet1!A1",
      load: vi.fn(),
      select: vi.fn(),
      format: {
        font: { bold: false, italic: false, color: "" },
        fill: { color: "" },
        autofitColumns,
        autofitRows,
        borders: { getItem: vi.fn() }
      }
    };
    const sheet = { getRange: vi.fn(() => range) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };
    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (cb: (ctx: unknown) => Promise<unknown>) => cb(context)
    };

    const formatRange = excelTools.find(({ schema }) =>
      schema.function.name === "format_range"
    );
    await formatRange!.execute({
      sheet: "Sheet1",
      address: "A1",
      format: { bold: true, rowAutofit: true }
    });

    expect(autofitRows).toHaveBeenCalledTimes(1);
    expect(autofitColumns).not.toHaveBeenCalled();
  });

  it("rowAutofit: false does not call autofitRows", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const autofitColumns = vi.fn();
    const autofitRows = vi.fn();
    const range = {
      address: "Sheet1!A1",
      load: vi.fn(),
      select: vi.fn(),
      format: {
        font: { bold: false, italic: false, color: "" },
        fill: { color: "" },
        autofitColumns,
        autofitRows,
        borders: { getItem: vi.fn() }
      }
    };
    const sheet = { getRange: vi.fn(() => range) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };
    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (cb: (ctx: unknown) => Promise<unknown>) => cb(context)
    };

    const formatRange = excelTools.find(({ schema }) =>
      schema.function.name === "format_range"
    );
    await formatRange!.execute({
      sheet: "Sheet1",
      address: "A1",
      format: { bold: true, rowAutofit: false }
    });

    expect(autofitRows).not.toHaveBeenCalled();
  });
});

// ---------------------------------------------------------------------------
// C1 behavior: CenterAcrossSelection precondition validation
// ---------------------------------------------------------------------------

describe("formatRange C1: CenterAcrossSelection precondition checks", () => {
  let previousExcel: unknown;
  let previousOffice: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
    previousOffice = (globalThis as unknown as { Office?: unknown }).Office;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
    (globalThis as unknown as { Office?: unknown }).Office = previousOffice;
  });

  function helperForCenterAcrossSelection(opts: {
      address: string;
      rowCount: number;
      columnCount: number;
      values: unknown[][];
      formulas?: unknown[][];
      mergedAreas?: Array<{ address: string }>;
      officeApiSupported?: boolean;
    }) {
      const autofitColFn = vi.fn(() => {
        events.push("autofit-columns");
      });
      const autofitRowFn = vi.fn(() => {
        events.push("autofit-rows");
      });
      const selectFn = vi.fn(() => {
        events.push("select");
      });
      const events: string[] = [];
      let syncCount = 0;
      const sync = vi.fn().mockImplementation(() => {
        syncCount++;
        const token = `sync#${syncCount}`;
        events.push(token);
        return Promise.resolve(undefined);
      });

      const load = vi.fn((value: string | string[]) => {
        events.push(`load:${Array.isArray(value) ? value.join("|") : value}`);
      });

      const fontState = {
        bold: false,
        italic: false,
        color: "",
        name: "",
        size: 11,
        underline: "None"
      };
      const fillState = { color: "" };
      const numberFormatState = { value: undefined as unknown };
      const alignmentState = { value: undefined as unknown };
      const wrapState = { value: false };

      const fontObj = {} as Record<string, unknown>;
      Object.defineProperty(fontObj, "bold", {
        configurable: true,
        get: () => fontState.bold,
        set: (value) => {
          events.push("set:bold");
          fontState.bold = value === true;
        }
      });
      Object.defineProperty(fontObj, "italic", {
        configurable: true,
        get: () => fontState.italic,
        set: (value) => {
          events.push("set:italic");
          fontState.italic = value === true;
        }
      });
      Object.defineProperty(fontObj, "color", {
        configurable: true,
        get: () => fontState.color,
        set: (value) => {
          events.push("set:fontColor");
          fontState.color = typeof value === "string" ? value : String(value);
        }
      });
      Object.defineProperty(fontObj, "name", {
        configurable: true,
        get: () => fontState.name,
        set: (value) => {
          fontState.name = String(value);
        }
      });
      Object.defineProperty(fontObj, "size", {
        configurable: true,
        get: () => fontState.size,
        set: (value) => {
          fontState.size = typeof value === "number" ? value : fontState.size;
        }
      });
      Object.defineProperty(fontObj, "underline", {
        configurable: true,
        get: () => fontState.underline,
        set: (value) => {
          fontState.underline = String(value);
        }
      });

      const fillObj = {} as Record<string, unknown>;
      Object.defineProperty(fillObj, "color", {
        configurable: true,
        get: () => fillState.color,
        set: (value) => {
          events.push("set:fillColor");
          fillState.color = typeof value === "string" ? value : "";
        }
      });

      const format: Record<string, unknown> = {
        font: fontObj,
        fill: fillObj,
        autofitColumns: vi.fn(() => {
          events.push("autofit-columns");
        }),
        autofitRows: vi.fn(() => {
          events.push("autofit-rows");
        }),
        borders: { getItem: vi.fn() },
        horizontalAlignment: undefined as unknown
      };

      Object.defineProperty(format, "horizontalAlignment", {
        configurable: true,
        get: () => alignmentState.value,
        set: (value) => {
          events.push("set:horizontalAlignment");
          alignmentState.value = value;
        }
      });

      Object.defineProperty(format, "wrapText", {
        configurable: true,
        get: () => wrapState.value,
        set: (value) => {
          events.push("set:wrapText");
          wrapState.value = value as boolean;
        }
      });

      format.autofitColumns = autofitColFn;
      format.autofitRows = autofitRowFn;

      const isSetSupported = vi.fn().mockReturnValue(opts.officeApiSupported ?? true);
      (globalThis as unknown as { Office: unknown }).Office = {
        context: {
          requirements: {
            isSetSupported
          }
        }
      };

      const mergedAreasOrNullObject = {
        isNullObject: opts.mergedAreas ? opts.mergedAreas.length === 0 : true,
        load: vi.fn(() => {
          events.push("load:mergedAreas");
        })
      };

      const range = {
        address: opts.address,
        rowCount: opts.rowCount,
        columnCount: opts.columnCount,
        values: opts.values,
        formulas: opts.formulas ?? opts.values,
        load,
        select: selectFn,
        format,
        getMergedAreasOrNullObject: vi.fn(() => mergedAreasOrNullObject)
      };
      Object.defineProperty(range, "numberFormat", {
        configurable: true,
        get: () => numberFormatState.value,
        set: (value) => {
          events.push("set:numberFormat");
          numberFormatState.value = value;
        }
      });

      const sheet = { getRange: vi.fn(() => range) };
      const context = {
        sync,
        workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
      };
      (globalThis as unknown as { Excel: unknown }).Excel = {
        run: async (cb: (ctx: unknown) => Promise<unknown>) => cb(context)
      };

      return {
        range,
        sheet,
        context,
        sync,
        events,
        syncCount: () => syncCount,
        fontState,
        fillState,
        numberFormatState,
        alignmentState,
        wrapState,
        autofitColFn,
        autofitRowFn,
        selectFn,
        load,
        isSetSupported
      };
    }

    it("centerAcrossSelection preloads CAS metadata before any formatting mutations", async () => {
      const { events, isSetSupported, range } = helperForCenterAcrossSelection({
        address: "Sheet1!A1:F1",
        rowCount: 1,
        columnCount: 6,
        values: [[null, null, null, null, null, null]],
        formulas: [["", "", "", "", "", ""]]
      });

      const formatRange = excelTools.find(({ schema }) =>
        schema.function.name === "format_range"
      );
      const result = JSON.parse(await formatRange!.execute({
        sheet: "Sheet1",
        address: "A1:F1",
        format: {
          horizontalAlignment: "centerAcrossSelection",
          bold: true,
          fillColor: "#ffcc00",
          numberFormat: "General",
          wrapText: true,
          columnAutofit: true,
          rowAutofit: true
        }
      }));

      expect(result).toEqual({
        ok: true,
        action: "format_range",
        sheet: "Sheet1",
        address: "Sheet1!A1:F1"
      });
      expect(isSetSupported).toHaveBeenCalledWith("ExcelApi", "1.13");
      expect(events.indexOf("load:rowCount|columnCount|values|formulas")).toBeLessThan(events.indexOf("sync#1"));
      expect(events.indexOf("load:mergedAreas")).toBeLessThan(events.indexOf("sync#1"));
      const firstSync = events.indexOf("sync#1");
      const secondSync = events.indexOf("sync#2");
      expect(firstSync).toBeLessThan(secondSync);
      expect(events.indexOf("set:bold")).toBeGreaterThan(firstSync);
      expect(events.indexOf("set:fillColor")).toBeGreaterThan(firstSync);
      expect(events.indexOf("set:numberFormat")).toBeGreaterThan(firstSync);
      expect(events.indexOf("set:horizontalAlignment")).toBeGreaterThan(firstSync);
      expect(events.indexOf("set:wrapText")).toBeGreaterThan(firstSync);
      expect(events.indexOf("autofit-columns")).toBeGreaterThan(firstSync);
      expect(events.indexOf("autofit-rows")).toBeGreaterThan(firstSync);
      expect(events.indexOf("select")).toBeGreaterThan(firstSync);
      expect(secondSync).toBeGreaterThan(events.indexOf("select"));
      expect(range.getMergedAreasOrNullObject).toHaveBeenCalledTimes(1);
    });

    it("occupied non-anchor cell fails before any mutation", async () => {
      const { events, range } = helperForCenterAcrossSelection({
        address: "Sheet1!A1:F1",
        rowCount: 1,
        columnCount: 6,
        values: [[null, null, "data-", null, null, null]],
        formulas: [["", "", "", "", "", ""]]
      });

      const formatRange = excelTools.find(({ schema }) =>
        schema.function.name === "format_range"
      );
      const result = JSON.parse(await formatRange!.execute({
        sheet: "Sheet1",
        address: "A1:F1",
        format: {
          horizontalAlignment: "CenterAcrossSelection",
          bold: true,
          fillColor: "#ff0000"
        }
      }));

      expect(result.ok).toBe(false);
      expect(events).toContain("sync#1");
      expect(events.indexOf("set:bold")).toBe(-1);
      expect(events.indexOf("set:fillColor")).toBe(-1);
      expect(events.indexOf("select")).toBe(-1);
      expect(range.select).not.toHaveBeenCalled();
    });

    it("one-cell range fails for CenterAcrossSelection before any mutation", async () => {
      const { events, range } = helperForCenterAcrossSelection({
        address: "Sheet1!A1",
        rowCount: 1,
        columnCount: 1,
        values: [[null]],
        formulas: [[""]]
      });

      const formatRange = excelTools.find(({ schema }) =>
        schema.function.name === "format_range"
      );
      const result = JSON.parse(await formatRange!.execute({
        sheet: "Sheet1",
        address: "A1",
        format: {
          horizontalAlignment: "CenterAcrossSelection",
          bold: true
        }
      }));

      expect(result.ok).toBe(false);
      expect(events).toContain("sync#1");
      expect(events.indexOf("set:bold")).toBe(-1);
      expect(range.select).not.toHaveBeenCalled();
    });

    it("multi-row range fails for CenterAcrossSelection before any mutation", async () => {
      const { events, range } = helperForCenterAcrossSelection({
        address: "Sheet1!A1:F2",
        rowCount: 2,
        columnCount: 6,
        values: [
          [null, null, null, null, null, null],
          [null, null, null, null, null, null]
        ],
        formulas: [
          ["", "", "", "", "", ""],
          ["", "", "", "", "", ""]
        ]
      });

      const formatRange = excelTools.find(({ schema }) =>
        schema.function.name === "format_range"
      );
      const result = JSON.parse(await formatRange!.execute({
        sheet: "Sheet1",
        address: "A1:F2",
        format: {
          horizontalAlignment: "CenterAcrossSelection",
          bold: true
        }
      }));

      expect(result.ok).toBe(false);
      expect(events).toContain("sync#1");
      expect(events.indexOf("set:bold")).toBe(-1);
      expect(range.select).not.toHaveBeenCalled();
    });

    it("non-anchor formula fails before mutation even when displayed value is empty", async () => {
      const { events, range } = helperForCenterAcrossSelection({
        address: "Sheet1!A1:F1",
        rowCount: 1,
        columnCount: 6,
        values: [[null, null, null, null, null, null]],
        formulas: [["", "", "=IF(1,1)", "", "", ""]],
      });

      const formatRange = excelTools.find(({ schema }) =>
        schema.function.name === "format_range"
      );
      const result = JSON.parse(await formatRange!.execute({
        sheet: "Sheet1",
        address: "A1:F1",
        format: {
          horizontalAlignment: "CenterAcrossSelection",
          bold: true
        }
      }));

      expect(result.ok).toBe(false);
      expect(events).toContain("sync#1");
      expect(events.indexOf("set:bold")).toBe(-1);
      expect(range.select).not.toHaveBeenCalled();
    });

    it("already-merged span fails before formatting", async () => {
      const { events, range } = helperForCenterAcrossSelection({
        address: "Sheet1!A1:F1",
        rowCount: 1,
        columnCount: 6,
        values: [[null, null, null, null, null, null]],
        formulas: [["", "", "", "", "", ""]],
        mergedAreas: [{ address: "Sheet1!A1:F1" }]
      });

      const formatRange = excelTools.find(({ schema }) =>
        schema.function.name === "format_range"
      );
      const result = JSON.parse(await formatRange!.execute({
        sheet: "Sheet1",
        address: "A1:F1",
        format: {
          horizontalAlignment: "CenterAcrossSelection",
          bold: true
        }
      }));

      expect(result.ok).toBe(false);
      expect(events).toContain("sync#1");
      expect(events.indexOf("set:bold")).toBe(-1);
      expect(range.getMergedAreasOrNullObject).toHaveBeenCalledTimes(1);
    });

    it("unsupported ExcelApi 1.13 path fails with zero mutations and no host sync", async () => {
      const { events, sync, range, isSetSupported } = helperForCenterAcrossSelection({
        address: "Sheet1!A1:F1",
        rowCount: 1,
        columnCount: 6,
        values: [[null, null, null, null, null, null]],
        formulas: [["", "", "", "", "", ""]],
        officeApiSupported: false
      });

      const formatRange = excelTools.find(({ schema }) =>
        schema.function.name === "format_range"
      );
      const result = JSON.parse(await formatRange!.execute({
        sheet: "Sheet1",
        address: "A1:F1",
        format: {
          horizontalAlignment: "CenterAcrossSelection",
          bold: true
        }
      }));

      expect(result).toMatchObject({
        ok: false,
        error: expect.stringContaining("ExcelApi 1.13")
      });
      expect(isSetSupported).toHaveBeenCalledWith("ExcelApi", "1.13");
      expect(sync).toHaveBeenCalledTimes(0);
      expect(events.indexOf("set:bold")).toBe(-1);
      expect(range.select).not.toHaveBeenCalled();
    });

    it("non-strict CAS precondition failure uses loaded host address metadata", async () => {
      const { events, range } = helperForCenterAcrossSelection({
        address: "NamedRange",
        rowCount: 1,
        columnCount: 6,
        values: [[null, null, "x", null, null, null]],
        formulas: [["", "", "", "", "", ""]]
      });

      const formatRange = excelTools.find(({ schema }) =>
        schema.function.name === "format_range"
      );
      const result = JSON.parse(await formatRange!.execute({
        sheet: "Sheet1",
        address: "NamedRange",
        format: {
          horizontalAlignment: "CenterAcrossSelection",
          bold: true
        }
      }));

      expect(result).toEqual({
        ok: false,
        error: expect.stringContaining("non-anchor cell C1"),
        action: "format_range",
        sheet: "Sheet1",
        address: "NamedRange"
      });
      expect(result.address).toBe(range.address);
      expect(range.load).toHaveBeenCalled();
      expect(events).toContain("sync#1");
      expect(events.indexOf("set:bold")).toBe(-1);
    });
});