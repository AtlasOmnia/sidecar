import { describe, expect, it } from "vitest";
import {
  detectArithmeticRequest,
  formatArithmeticResult,
  ArithmeticOp,
} from "../deterministicArithmetic";

describe("detectArithmeticRequest", () => {
  it("detects sum request", () => {
    expect(detectArithmeticRequest("sum the selected range")).toEqual({
      op: "sum",
    });
    expect(detectArithmeticRequest("total of selection")).toEqual({
      op: "sum",
    });
    expect(detectArithmeticRequest("add up these cells")).toEqual({
      op: "sum",
    });
  });

  it("detects average request", () => {
    expect(detectArithmeticRequest("average of selected range")).toEqual({
      op: "average",
    });
    expect(detectArithmeticRequest("avg of these cells")).toEqual({
      op: "average",
    });
    expect(detectArithmeticRequest("what's the mean")).toEqual({
      op: "average",
    });
  });

  it("detects min request", () => {
    expect(detectArithmeticRequest("min of selection")).toEqual({
      op: "min",
    });
    expect(detectArithmeticRequest("lowest value in these cells")).toEqual({
      op: "min",
    });
  });

  it("detects max request", () => {
    expect(detectArithmeticRequest("max of selection")).toEqual({
      op: "max",
    });
    expect(detectArithmeticRequest("highest value")).toEqual({
      op: "max",
    });
  });

  it("detects count request", () => {
    expect(detectArithmeticRequest("count the selected cells")).toEqual({
      op: "count",
    });
    expect(detectArithmeticRequest("how many cells")).toEqual({
      op: "count",
    });
  });

  it("returns null for complex requests", () => {
    expect(
      detectArithmeticRequest("sum the selected range and format it as currency"),
    ).toBeNull();
    expect(
      detectArithmeticRequest("create a chart from the selection"),
    ).toBeNull();
    expect(
      detectArithmeticRequest("why is the sum so high?"),
    ).toBeNull();
    expect(
      detectArithmeticRequest("format the selected cells bold"),
    ).toBeNull();
    expect(
      detectArithmeticRequest("sort the table by column A"),
    ).toBeNull();
  });

  it("returns null for requests without selection context", () => {
    // Long request without selection keywords
    expect(
      detectArithmeticRequest(
        "I need to calculate the sum of all values in column B where column A equals 'Widget'",
      ),
    ).toBeNull();
  });

  it("returns null for empty or very short input", () => {
    expect(detectArithmeticRequest("")).toBeNull();
    expect(detectArithmeticRequest("ab")).toBeNull();
  });

  it("is case-insensitive", () => {
    expect(detectArithmeticRequest("SUM THE SELECTION")).toEqual({
      op: "sum",
    });
    expect(detectArithmeticRequest("Average of These Cells")).toEqual({
      op: "average",
    });
  });
});

describe("formatArithmeticResult", () => {
  it("formats sum result", () => {
    const stats = {
      count: 10,
      numericCount: 8,
      sum: 1500,
      average: 187.5,
      min: 50,
      max: 300,
      blankCount: 2,
    };
    const result = formatArithmeticResult("sum", stats, "Sheet1!A1:B5");
    expect(result).toContain("**1,500**");
    expect(result).toContain("sum");
  });

  it("formats average result", () => {
    const stats = {
      count: 5,
      numericCount: 5,
      sum: 500,
      average: 100,
      min: 10,
      max: 200,
      blankCount: 0,
    };
    const result = formatArithmeticResult("average", stats, "A1:A5");
    expect(result).toContain("**100**");
    expect(result).toContain("average");
  });

  it("handles null average (no numeric values)", () => {
    const stats = {
      count: 5,
      numericCount: 0,
      sum: 0,
      average: null,
      min: null,
      max: null,
      blankCount: 2,
    };
    const result = formatArithmeticResult("average", stats, "A1:A5");
    expect(result).toContain("no numeric values");
  });

  it("handles null min/max", () => {
    const stats = {
      count: 5,
      numericCount: 0,
      sum: 0,
      average: null,
      min: null,
      max: null,
      blankCount: 2,
    };
    const minResult = formatArithmeticResult("min", stats, "A1:A5");
    expect(minResult).toContain("no numeric values");

    const maxResult = formatArithmeticResult("max", stats, "A1:A5");
    expect(maxResult).toContain("no numeric values");
  });

  it("formats count result", () => {
    const stats = {
      count: 25,
      numericCount: 20,
      sum: 1000,
      average: 50,
      min: 1,
      max: 100,
      blankCount: 3,
    };
    const result = formatArithmeticResult("count", stats, "Sheet1!A1:C10");
    expect(result).toContain("**25**");
    expect(result).toContain("20 numeric");
    expect(result).toContain("3 blank");
  });
});
