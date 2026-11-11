import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("conditionalFormatRange Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("derives a strict-A1 response address without loading host metadata", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const conditionalRule = {
      rule: undefined as unknown,
      format: {
        fill: { color: "" },
        font: { color: "" }
      }
    };
    const conditionalFormat = { cellValue: conditionalRule };
    const range = {
      address: "host-normalized-address",
      load: vi.fn(),
      conditionalFormats: { add: vi.fn(() => conditionalFormat) }
    };
    const sheet = { getRange: vi.fn(() => range) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const conditionalFormatRange = excelTools.find(({ schema }) =>
      schema.function.name === "conditional_format_range"
    );
    const result = JSON.parse(await conditionalFormatRange!.execute({
      sheet: "Data's Sheet",
      address: "$B$2:$C$5",
      ruleType: "cellValue",
      formula1: "10",
      fillColor: "#ffff00",
      fontColor: "#000000"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).not.toHaveBeenCalled();
    expect(range.conditionalFormats.add).toHaveBeenCalledWith("CellValue");
    expect(conditionalRule.rule).toEqual({ operator: "GreaterThan", formula1: "10", formula2: undefined });
    expect(conditionalRule.format.fill.color).toBe("#ffff00");
    expect(conditionalRule.format.font.color).toBe("#000000");
    expect(result).toEqual({
      ok: true,
      action: "conditional_format_range",
      sheet: "Data's Sheet",
      address: "'Data''s Sheet'!B2:C5",
      ruleType: "cellValue"
    });
  });

  it("retains host address loading for non-A1 expressions", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const textRule = {
      rule: undefined as unknown,
      format: {
        fill: { color: "" },
        font: { color: "" }
      }
    };
    const conditionalFormat = { textComparison: textRule };
    const range = {
      address: "Data!NamedRange",
      load: vi.fn(),
      conditionalFormats: { add: vi.fn(() => conditionalFormat) }
    };
    const sheet = { getRange: vi.fn(() => range) };
    const context = {
      sync,
      workbook: { worksheets: { getItem: vi.fn(() => sheet) } }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const conditionalFormatRange = excelTools.find(({ schema }) =>
      schema.function.name === "conditional_format_range"
    );
    const result = JSON.parse(await conditionalFormatRange!.execute({
      sheet: "Data",
      address: "NamedRange",
      ruleType: "containsText",
      text: "foo"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).toHaveBeenCalledWith("address");
    expect(range.conditionalFormats.add).toHaveBeenCalledWith("ContainsText");
    expect(textRule.rule).toEqual({ operator: "Contains", text: "foo" });
    expect(result).toEqual({
      ok: true,
      action: "conditional_format_range",
      sheet: "Data",
      address: "Data!NamedRange",
      ruleType: "containsText"
    });
  });

  it("advertises only implemented rule types and fails before context execution for unsupported values", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const conditionalFormatRangeTool = excelTools.find(({ schema }) =>
      schema.function.name === "conditional_format_range"
    );
    const conditionalToolParams = conditionalFormatRangeTool?.schema.function.parameters as Record<string, unknown>;
    const ruleTypeProperty = conditionalToolParams?.properties as Record<string, { enum?: unknown }>;

    expect(ruleTypeProperty?.ruleType?.enum).toEqual(["cellValue", "containsText"]);

    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn()
        }
      }
    };

    const run = vi.fn(async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context));
    (globalThis as unknown as { Excel: unknown }).Excel = {
      run
    };

    const result = JSON.parse(await conditionalFormatRangeTool!.execute({
      sheet: "Data",
      address: "B2:B5",
      ruleType: "colorScale"
    }));

    expect(result).toEqual({
      ok: false,
      name: "Error",
      error: "Unsupported conditional format ruleType \"colorScale\"."
    });
    expect(run).not.toHaveBeenCalled();
    expect(sync).not.toHaveBeenCalled();
  });
});
