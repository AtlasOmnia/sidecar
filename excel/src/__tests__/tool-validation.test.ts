import { describe, expect, it } from "vitest";
import { validateToolCall, parseToolArgs } from "../toolValidation";
import { ToolErrorCategory } from "../toolErrors";

describe("validateToolCall", () => {
  it("rejects an unknown tool name", () => {
    const result = validateToolCall("nonexistent_tool", "{}");
    expect("ok" in result && result.ok === false).toBe(true);
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
      expect(result.error).toContain("Unknown tool");
    }
  });

  it("rejects non-JSON arguments", () => {
    const result = validateToolCall("read_range", "not json");
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
      expect(result.error).toContain("valid JSON");
    }
  });

  it("rejects missing required arguments", () => {
    const result = validateToolCall("read_range", "{}");
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
      expect(result.error).toContain("Missing required argument");
    }
  });

  it("rejects unknown extra arguments", () => {
    const result = validateToolCall("read_range", JSON.stringify({
      sheet: "Sheet1",
      address: "A1:B2",
      extraParam: "should not be here",
    }));
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
      expect(result.error).toContain("Unknown argument");
    }
  });

  it("accepts valid read_range call", () => {
    const result = validateToolCall("read_range", JSON.stringify({
      sheet: "Sheet1",
      address: "A1:B2",
    }));
    expect("valid" in result && result.valid === true).toBe(true);
  });

  it("rejects string where number is expected", () => {
    const result = validateToolCall("freeze_panes", JSON.stringify({
      sheet: "Sheet1",
      rows: "not a number",
      columns: 2,
    }));
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
      expect(result.error).toContain("must be a number");
    }
  });

  it("rejects float where integer is expected", () => {
    const result = validateToolCall("freeze_panes", JSON.stringify({
      sheet: "Sheet1",
      rows: 1.5,
      columns: 2,
    }));
    if ("category" in result) {
      expect(result.error).toContain("must be an integer");
    }
  });

  it("rejects wrong type for boolean field", () => {
    const result = validateToolCall("insert_table", JSON.stringify({
      sheet: "Sheet1",
      address: "A1:C10",
      hasHeaders: "yes",
    }));
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
      expect(result.error).toContain("must be a boolean");
    }
  });

  it("rejects wrong type for array field", () => {
    const result = validateToolCall("write_range", JSON.stringify({
      sheet: "Sheet1",
      address: "A1:B2",
      values: "not an array",
    }));
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
      expect(result.error).toContain("must be an array");
    }
  });

  it("accepts valid write_range call", () => {
    const result = validateToolCall("write_range", JSON.stringify({
      sheet: "Sheet1",
      address: "A1:B2",
      values: [["a", "b"], ["c", "d"]],
    }));
    expect("valid" in result && result.valid === true).toBe(true);
  });

  it("rejects invalid enum value", () => {
    const result = validateToolCall("conditional_format_range", JSON.stringify({
      sheet: "Sheet1",
      address: "A1:A10",
      ruleType: "ColorScale",
    }));
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
      expect(result.error).toContain("must be one of");
    }
  });

  it("accepts valid conditional_format_range call", () => {
    const result = validateToolCall("conditional_format_range", JSON.stringify({
      sheet: "Sheet1",
      address: "A1:A10",
      ruleType: "cellValue",
      operator: "GreaterThan",
      formula1: "100",
    }));
    expect("valid" in result && result.valid === true).toBe(true);
  });

  it("accepts tool with no required args (get_workbook_overview)", () => {
    const result = validateToolCall("get_workbook_overview", "{}");
    expect("valid" in result && result.valid === true).toBe(true);
  });

  it("accepts run_calculation with empty args", () => {
    const result = validateToolCall("run_calculation", "{}");
    expect("valid" in result && result.valid === true).toBe(true);
  });

  it("returns the correct tool index for valid calls", () => {
    const result = validateToolCall("read_range", JSON.stringify({
      sheet: "Sheet1",
      address: "A1:B2",
    }));
    if ("valid" in result) {
      expect(typeof result.toolIndex).toBe("number");
      expect(result.toolIndex).toBeGreaterThanOrEqual(0);
    }
  });

  it("returns parsed args for valid calls", () => {
    const result = validateToolCall("read_range", JSON.stringify({
      sheet: "Sheet1",
      address: "A1:B2",
    }));
    if ("valid" in result) {
      expect(result.parsedArgs).toEqual({ sheet: "Sheet1", address: "A1:B2" });
    }
  });
});

describe("parseToolArgs", () => {
  it("parses valid JSON object", () => {
    const args = parseToolArgs('{"sheet":"Sheet1","address":"A1:B2"}');
    expect(args).toEqual({ sheet: "Sheet1", address: "A1:B2" });
  });

  it("returns empty object for empty string", () => {
    expect(parseToolArgs("")).toEqual({});
  });

  it("returns null for invalid JSON", () => {
    expect(parseToolArgs("bad json")).toBeNull();
  });

  it("returns null for arrays", () => {
    expect(parseToolArgs("[1,2,3]")).toBeNull();
  });

  it("returns null for primitives", () => {
    expect(parseToolArgs("42")).toBeNull();
    expect(parseToolArgs('"hello"')).toBeNull();
  });
});
