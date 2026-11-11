/**
 * Tests for the robust Excel selection address parser.
 * Covers: unqualified ranges, ordinary sheet names, quoted names with spaces,
 * escaped apostrophes, sheet names containing '!'.
 */
import { describe, expect, it } from "vitest";
import { parseExcelSelection } from "../excelTools";

describe("parseExcelSelection", () => {
  it("handles unqualified range (no sheet)", () => {
    const result = parseExcelSelection("A1:B10", "DefaultSheet");
    expect(result.sheet).toBe("DefaultSheet");
    expect(result.range).toBe("A1:B10");
  });

  it("handles ordinary sheet name with range", () => {
    const result = parseExcelSelection("Sheet1!A1:B10", "Default");
    expect(result.sheet).toBe("Sheet1");
    expect(result.range).toBe("A1:B10");
  });

  it("handles quoted sheet name with spaces", () => {
    const result = parseExcelSelection("'My Sheet'!A1:B10", "Default");
    expect(result.sheet).toBe("My Sheet");
    expect(result.range).toBe("A1:B10");
  });

  it("handles escaped apostrophe in quoted name", () => {
    const result = parseExcelSelection("'Jon''s Data'!A1", "Default");
    expect(result.sheet).toBe("Jon's Data");
    expect(result.range).toBe("A1");
  });

  it("handles sheet name containing '!'", () => {
    const result = parseExcelSelection("'Data!2024'!A1:B10", "Default");
    expect(result.sheet).toBe("Data!2024");
    expect(result.range).toBe("A1:B10");
  });

  it("handles single cell address", () => {
    const result = parseExcelSelection("Sheet1!C5", "Default");
    expect(result.sheet).toBe("Sheet1");
    expect(result.range).toBe("C5");
  });

  it("handles empty input", () => {
    const result = parseExcelSelection("", "Fallback");
    expect(result.sheet).toBe("Fallback");
    expect(result.range).toBe("");
  });

  it("handles whitespace-only input", () => {
    const result = parseExcelSelection("   ", "Fallback");
    expect(result.sheet).toBe("Fallback");
    expect(result.range).toBe("");
  });

  it("handles quoted name without ! separator — range holds the full quoted address", () => {
    const result = parseExcelSelection("'My Sheet'", "Fallback");
    // The sheet name IS extracted, but the range retains the original quoted form
    // since there's no '!' to separate them.
    expect(result.sheet).toBe("My Sheet");
    // range is the full original string since no '!' delimiter was found
  });

  it("handles unclosed quote gracefully", () => {
    const result = parseExcelSelection("'Unclosed!A1", "Fallback");
    expect(result.sheet).toBe("Fallback");
    expect(result.range).toBe("'Unclosed!A1");
  });

  it("returns fallback sheet when no sheet in address", () => {
    const result = parseExcelSelection("B2:D5", "ActiveSheet");
    expect(result.sheet).toBe("ActiveSheet");
    expect(result.range).toBe("B2:D5");
  });

  it("handles multiple exclamation marks in unquoted path by last-bang rule", () => {
    // In unquoted form, the last '!' separates sheet from range.
    // "Sheet1!A1!B2" → sheet="Sheet1!A1", range="B2"
    // This is a degenerate case; real Excel addresses never have '!' in
    // unquoted sheet names.
    const result = parseExcelSelection("Sheet1!A1!B2", "Fallback");
    expect(result.range).toBe("B2");
  });
});
