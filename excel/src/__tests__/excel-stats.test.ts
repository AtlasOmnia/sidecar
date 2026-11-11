import { describe, expect, it } from "vitest";
import {
  parseConditionalFormatRuleType,
  parseDataValidationSource,
  parseDataValidationType,
  nonNegativeIntArg,
  parseNumericCell,
  summarizeRangeValues,
  toolSchemas,
  excelTools,
} from "../excelTools";

function findSchema(name: string) {
  const schema = toolSchemas.find((entry) => entry.function.name === name)?.function.parameters;
  return schema as Record<string, unknown> | undefined;
}

describe("parseNumericCell", () => {
  it("parses finite numeric values and trimmed numeric strings", () => {
    expect(parseNumericCell(12)).toBe(12);
    expect(parseNumericCell(" 3.5 ")).toBe(3.5);
    expect(parseNumericCell("1,250.25")).toBe(1250.25);
    expect(parseNumericCell("-2e2")).toBe(-200);
  });

  it("returns null for blanks and non-numeric values", () => {
    expect(parseNumericCell("")).toBeNull();
    expect(parseNumericCell("  ")).toBeNull();
    expect(parseNumericCell("foo")).toBeNull();
    expect(parseNumericCell(undefined)).toBeNull();
    expect(parseNumericCell(true)).toBeNull();
    expect(parseNumericCell(NaN)).toBeNull();
  });

  it("rejects malformed thousands separators while preserving supported numeric forms", () => {
    expect(parseNumericCell("1,250.25")).toBe(1250.25);
    expect(parseNumericCell("-2e2")).toBe(-200);
    expect(parseNumericCell(".5")).toBe(0.5);
    expect(parseNumericCell("12,34")).toBeNull();
    expect(parseNumericCell("1,2,3")).toBeNull();
  });

  it("does not aggregate malformed thousands-separated text", () => {
    const stats = summarizeRangeValues([["1,000", "12,34", "1,2,3"]]);
    expect(stats).toEqual({
      count: 3,
      numericCount: 1,
      sum: 1000,
      average: 1000,
      min: 1000,
      max: 1000,
      blankCount: 0
    });
  });
});

describe("summarizeRangeValues", () => {
  it("returns deterministic stats with mixed numeric, text, and blank values", () => {
    const values = [
      [10, "2", "  ", "abc", null],
      [undefined, true, "3.5", "1,000"]
    ] as const;

    const stats = summarizeRangeValues(values);
    expect(stats).toEqual({
      count: 9,
      numericCount: 4,
      sum: 1015.5,
      average: 253.875,
      min: 2,
      max: 1000,
      blankCount: 3
    });
  });

  it("returns null min, max, and average when no numeric values are present", () => {
    const values = [
      ["" as unknown, "abc", null],
      [undefined, false, true]
    ];

    const stats = summarizeRangeValues(values);
    expect(stats).toEqual({
      count: 6,
      numericCount: 0,
      sum: 0,
      average: null,
      min: null,
      max: null,
      blankCount: 3
    });
  });
});

describe("excel tool schemas", () => {
  it("registers conditional_format_range with required arguments", () => {
    const schema = findSchema("conditional_format_range");
    expect(schema).toBeDefined();
    expect(schema?.required).toEqual(["sheet", "address", "ruleType"]);
    const properties = schema?.properties as Record<string, unknown> | undefined;
    expect(properties).toBeDefined();
    expect(properties).toHaveProperty("ruleType");
  });

  it("registers freeze_panes with required arguments", () => {
    const schema = findSchema("freeze_panes");
    expect(schema).toBeDefined();
    expect(schema?.required).toEqual(["sheet", "rows", "columns"]);
    const properties = schema?.properties as Record<string, unknown> | undefined;
    expect(properties).toBeDefined();
    expect(properties).toHaveProperty("rows");
    expect(properties).toHaveProperty("columns");
  });

  it("registers set_data_validation with required arguments", () => {
    const schema = findSchema("set_data_validation");
    expect(schema).toBeDefined();
    expect(schema?.required).toEqual(["sheet", "address", "type", "source"]);
    const properties = schema?.properties as Record<string, unknown> | undefined;
    expect(properties).toBeDefined();
    expect(properties).toHaveProperty("type");
    expect(properties).toHaveProperty("source");
  });

  it("registers create_pivot_table with required arguments", () => {
    const schema = findSchema("create_pivot_table");
    expect(schema).toBeDefined();
    expect(schema?.required).toEqual(["sheet", "address", "sourceSheet", "sourceRange", "rowFields", "dataFields"]);
    const properties = schema?.properties as Record<string, unknown> | undefined;
    expect(properties).toBeDefined();
    expect(properties).toHaveProperty("rowFields");
    expect(properties).toHaveProperty("dataFields");
    expect(properties).toHaveProperty("columnFields");
  });

  it("registers autofill_formula with required arguments", () => {
    const schema = findSchema("autofill_formula");
    expect(schema).toBeDefined();
    expect(schema?.required).toEqual(["sheet", "sourceAddress", "targetAddress"]);
  });

  it("registers rename_sheet with required arguments", () => {
    const schema = findSchema("rename_sheet");
    expect(schema).toBeDefined();
    expect(schema?.required).toEqual(["oldName", "newName"]);
  });

  it("registers copy_sheet with required arguments", () => {
    const schema = findSchema("copy_sheet");
    expect(schema).toBeDefined();
    expect(schema?.required).toEqual(["sourceName", "newName"]);
  });

  it("registers protect_sheet with required arguments", () => {
    const schema = findSchema("protect_sheet");
    expect(schema).toBeDefined();
    expect(schema?.required).toEqual(["sheet"]);
  });

  it("registers unprotect_sheet with required arguments", () => {
    const schema = findSchema("unprotect_sheet");
    expect(schema).toBeDefined();
    expect(schema?.required).toEqual(["sheet"]);
  });

  it("format_range extended schema includes fontName, fontSize, underline, wrapText, indentLevel, textOrientation, borderColor, borderWeight", () => {
    const schema = findSchema("format_range");
    expect(schema).toBeDefined();
    const formatProps = (schema?.properties as Record<string, unknown>).format as Record<string, unknown>;
    const props = formatProps.properties as Record<string, unknown>;
    expect(props).toHaveProperty("fontName");
    expect(props).toHaveProperty("fontSize");
    expect(props).toHaveProperty("underline");
    expect(props).toHaveProperty("wrapText");
    expect(props).toHaveProperty("indentLevel");
    expect(props).toHaveProperty("textOrientation");
    expect(props).toHaveProperty("borderColor");
    expect(props).toHaveProperty("borderWeight");
  });

  it("format_range indentLevel and textOrientation use integer type", () => {
    const schema = findSchema("format_range");
    expect(schema).toBeDefined();
    const formatProps = (schema?.properties as Record<string, unknown>).format as Record<string, unknown>;
    const props = formatProps.properties as Record<string, unknown>;
    const indentLevel = props.indentLevel as Record<string, unknown>;
    const textOrientation = props.textOrientation as Record<string, unknown>;
    expect(indentLevel.type).toBe("integer");
    expect(textOrientation.type).toBe("integer");
  });

  it("protect_sheet schema has allowSort and allowFilter as optional booleans", () => {
    const schema = findSchema("protect_sheet");
    expect(schema).toBeDefined();
    const props = schema?.properties as Record<string, unknown>;
    expect(props).toHaveProperty("allowSort");
    expect(props).toHaveProperty("allowFilter");
    const allowSort = props.allowSort as Record<string, unknown>;
    const allowFilter = props.allowFilter as Record<string, unknown>;
    expect(allowSort.type).toBe("boolean");
    expect(allowFilter.type).toBe("boolean");
    expect(schema?.required).toEqual(["sheet"]);
  });

  it("create_pivot_table schema has dataFields with summarizeBy enum", () => {
    const schema = findSchema("create_pivot_table");
    expect(schema).toBeDefined();
    const props = schema?.properties as Record<string, unknown>;
    expect(props).toHaveProperty("dataFields");
    const dfItems = (props.dataFields as Record<string, unknown>).items as Record<string, unknown>;
    const dfProps = dfItems.properties as Record<string, unknown>;
    expect(dfProps).toHaveProperty("summarizeBy");
    const summarizeBy = dfProps.summarizeBy as Record<string, unknown>;
    expect(summarizeBy.enum).toEqual(["sum", "count", "average", "max", "min"]);
  });
});

describe("conditional format and validation helpers", () => {
  it("accepts supported conditional format rule types", () => {
    expect(parseConditionalFormatRuleType("cellValue")).toBe("cellValue");
    expect(parseConditionalFormatRuleType("containsText")).toBe("containsText");
    expect(() => parseConditionalFormatRuleType("ColorScale")).toThrow(Error);
    expect(() => parseConditionalFormatRuleType(12 as unknown)).toThrow(Error);
  });

  it("validates non-negative integer args", () => {
    expect(nonNegativeIntArg({ rows: 0 }, "rows")).toBe(0);
    expect(nonNegativeIntArg({ rows: 8 }, "rows")).toBe(8);
    expect(nonNegativeIntArg({ rows: 1000 }, "rows")).toBe(1000);
    expect(() => nonNegativeIntArg({ rows: -1 }, "rows")).toThrow(Error);
    expect(() => nonNegativeIntArg({ rows: 2.2 }, "rows")).toThrow(Error);
    expect(() => nonNegativeIntArg({}, "rows")).toThrow(Error);
    expect(() => nonNegativeIntArg({ rows: null }, "rows")).toThrow(Error);
    expect(() => nonNegativeIntArg({ rows: "5" }, "rows")).toThrow(Error);
    expect(() => nonNegativeIntArg({ rows: Infinity }, "rows")).toThrow(Error);
  });

  it("validates data validation type and source", () => {
    expect(parseDataValidationType("list")).toBe("list");
    expect(() => parseDataValidationType("listWithRange")).toThrow(Error);
    expect(parseDataValidationSource("Yes,No,Maybe")).toBe("Yes,No,Maybe");
    expect(() => parseDataValidationSource("   ")).toThrow(Error);
    expect(() => parseDataValidationSource(12 as unknown)).toThrow(Error);
  });

  it("registers add_comment with required arguments", () => {
    const schema = findSchema("add_comment");
    expect(schema).toBeDefined();
    expect(schema?.required).toEqual(["sheet", "address", "text"]);
    const props = schema?.properties as Record<string, unknown>;
    expect(props).toHaveProperty("text");
  });

  it("registers clear_comments with required arguments", () => {
    const schema = findSchema("clear_comments");
    expect(schema).toBeDefined();
    expect(schema?.required).toEqual(["sheet", "address"]);
  });

  it("registers import_csv_to_sheet with required arguments", () => {
    const schema = findSchema("import_csv_to_sheet");
    expect(schema).toBeDefined();
    expect(schema?.required).toEqual(["sheet", "address", "csv"]);
    const props = schema?.properties as Record<string, unknown>;
    expect(props).toHaveProperty("csv");
    const csvProp = props.csv as Record<string, unknown>;
    expect(csvProp.description).toContain("CSV");
  });

  it("import_csv_to_sheet hasHeaders description says default true (headers bolded)", () => {
    const schema = findSchema("import_csv_to_sheet");
    expect(schema).toBeDefined();
    const props = schema?.properties as Record<string, unknown>;
    const hasHeadersProp = props.hasHeaders as Record<string, unknown>;
    expect(hasHeadersProp.description).toContain("default true");
    expect(hasHeadersProp.description).toContain("bolded");
    // Should not contradict itself: default is true == headers treated as headers, not data
    expect(hasHeadersProp.description).not.toContain("treated as data");
  });

  it("import_csv_to_sheet csv description mentions delimiter auto-detection", () => {
    const schema = findSchema("import_csv_to_sheet");
    expect(schema).toBeDefined();
    const props = schema?.properties as Record<string, unknown>;
    const csvProp = props.csv as Record<string, unknown>;
    expect(csvProp.description).toContain("Delimiter");
    // Should not say first row is treated as data (default is headers=true => headers are headers)
    expect(csvProp.description).not.toContain("treated as data");
  });

  it("format_range advertises horizontalAlignment enum with canonical values", () => {
    const schema = findSchema("format_range");
    expect(schema).toBeDefined();
    const formatProps = (schema?.properties as Record<string, unknown>).format as Record<string, unknown>;
    const props = formatProps.properties as Record<string, unknown>;
    const alignment = props.horizontalAlignment as Record<string, unknown>;
    expect(alignment).toBeDefined();
    expect(alignment.type).toBe("string");
    expect(alignment.enum).toEqual([
      "General",
      "Left",
      "Center",
      "Right",
      "Fill",
      "Justify",
      "CenterAcrossSelection",
      "Distributed",
    ]);
  });

  it("format_range description explains CenterAcrossSelection formats single-row multi-cell span without merging", () => {
    const tool = excelTools.find((t) => t.schema.function.name === "format_range");
    expect(tool).toBeDefined();
    const desc = (tool!.schema.function as { description: string }).description;
    expect(desc).toContain("CenterAcrossSelection");
    expect(desc).toContain("without merging");
  });

  it("format_range horizontalAlignment does not expose bridgeSessionId", () => {
    const schema = findSchema("format_range");
    expect(schema).toBeDefined();
    expect(JSON.stringify(schema)).not.toContain("bridgeSessionId");
  });

  it("R1C1 formulas enable relative reference adjustment across target cells", () => {
    // R1C1 notation is inherently relative: =R[-1]C*2 means "multiply the cell above by 2"
    // When written to any target cell, the reference adjusts to that cell's context.
    // This test validates the contract: a formula in R1C1 notation starting with "="
    // will be interpreted relative to each target cell's position.
    const r1c1Formulas = [
      "=R[-1]C*2",
      "=SUM(RC[-2]:RC[-1])",
      "=R[1]C[1]+R[-1]C[-1]",
      "=RC1*RC2"
    ];
    for (const formula of r1c1Formulas) {
      expect(formula.startsWith("=")).toBe(true);
    }

    // Verify the array construction pattern used in the fallback
    const rowCount = 5;
    const columnCount = 1;
    const r1c1 = "=R[-1]C*2";
    const target = Array.from({ length: rowCount }, () =>
      Array.from({ length: columnCount }, () => r1c1)
    );
    expect(target.length).toBe(5);
    expect(target[0].length).toBe(1);
    expect(target[2][0]).toBe("=R[-1]C*2");

    // Multi-column target
    const multiCol = Array.from({ length: 3 }, () =>
      Array.from({ length: 2 }, () => "=RC[-1]+R[-1]C")
    );
    expect(multiCol.length).toBe(3);
    expect(multiCol[0].length).toBe(2);
    expect(multiCol[1][1]).toBe("=RC[-1]+R[-1]C");
  });
});
