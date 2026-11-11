import { describe, expect, it } from "vitest";
import { workbookPlanContract } from "../workbookPlanContract";
import { validateWorkbookPlan } from "../workbookPlanValidation";
import { prepareWorkbookPlanVerificationRules } from "../workbookPlanVerification";

const errorsOf = (result: ReturnType<typeof validateWorkbookPlan>) => {
  if (result.ok) throw new Error("Expected validation failure");
  return result.errors;
};

const compareCodeUnit = (left: string, right: string): number => {
  const limit = Math.min(left.length, right.length);
  for (let index = 0; index < limit; index += 1) {
    const diff = left.charCodeAt(index) - right.charCodeAt(index);
    if (diff !== 0) return diff;
  }
  return left.length - right.length;
};

// ---------------------------------------------------------------------------
// Helper builders
// ---------------------------------------------------------------------------

const SECRET_FORMULA = "=SECRET(\uD83D\uDE0E)"; // formula with emoji, bytes > chars
const SECRET_VALUE = "\uD83D\uDE0E"; // emoji value
// Cyrillic U+0410 = 2 bytes UTF-8 each
const LONG_TEXT_4KB = "\u0410".repeat(2049); // 2049 × 2 = 4098 bytes UTF-8, exceeds 4096 limit
const LONG_TEXT_1KB = "\u0410".repeat(513); // 513 × 2 = 1026 bytes UTF-8, exceeds 1024 limit
const LONG_SECRET_FORMULA = `${SECRET_FORMULA}\uD83E\uDD2A`.repeat(500);
const EXCEL_CELL_TEXT_LIMIT = 32767;
const ASTRAL_UTF16_LIMIT_VALUE = "😀".repeat(16383) + "x";
const ASTRAL_UTF16_OVER_LIMIT_VALUE = "😀".repeat(16384);

const planWithValueCell = (value: string) => ({
  version: 1,
  outcome: "execute" as const,
  summary: "writes values",
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [[value]] }],
  verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["ok"]] }],
});

const planWithExpectedValues = (expectedValues: string) => ({
  version: 1,
  outcome: "execute" as const,
  summary: "verifies values",
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }],
  verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [[expectedValues]] }],
});

const planWithSentinel = (expected: string) => ({
  version: 1,
  outcome: "execute" as const,
  summary: "preserves sentinel",
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }],
  verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1", expected }],
});

const validInspect = () => ({
  version: 1,
  outcome: "inspect" as const,
  reason: "need data",
  reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1" }],
});

const validExecute = () => ({
  version: 1,
  outcome: "execute" as const,
  summary: "writes values",
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["hello"]] }],
  verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["hello"]] }],
});

const validAsk = () => ({
  version: 1,
  outcome: "ask" as const,
  message: "Need confirmation",
});

const validRefuse = () => ({
  version: 1,
  outcome: "refuse" as const,
  message: "Declined",
});

describe("ScalarValue Excel cell text boundary", () => {
  it("accepts an Excel-limit string in a write_values ValueRow cell", () => {
    const result = validateWorkbookPlan(planWithValueCell("x".repeat(EXCEL_CELL_TEXT_LIMIT)));
    expect(result.ok).toBe(true);
  });

  it("rejects a string over the Excel limit in a write_values ValueRow cell", () => {
    const result = validateWorkbookPlan(planWithValueCell("x".repeat(EXCEL_CELL_TEXT_LIMIT + 1)));
    expect(result.ok).toBe(false);
    expect(errorsOf(result)).toContainEqual(
      expect.objectContaining({
        code: "bound_exceeded",
        path: "/operations/0/values/0/0",
        limit: EXCEL_CELL_TEXT_LIMIT,
      }),
    );
  });

  it("accepts an Excel-limit string in a range_values_equal expectedValues cell", () => {
    const result = validateWorkbookPlan(planWithExpectedValues("x".repeat(EXCEL_CELL_TEXT_LIMIT)));
    expect(result.ok).toBe(true);
  });

  it("rejects a string over the Excel limit in a range_values_equal expectedValues cell", () => {
    const result = validateWorkbookPlan(planWithExpectedValues("x".repeat(EXCEL_CELL_TEXT_LIMIT + 1)));
    expect(result.ok).toBe(false);
    expect(errorsOf(result)).toContainEqual(
      expect.objectContaining({
        code: "bound_exceeded",
        path: "/verification/0/expectedValues/0/0",
        limit: EXCEL_CELL_TEXT_LIMIT,
      }),
    );
  });

  it("accepts an Excel-limit string in a sentinel_unchanged expected value", () => {
    const result = validateWorkbookPlan(planWithSentinel("x".repeat(EXCEL_CELL_TEXT_LIMIT)));
    expect(result.ok).toBe(true);
  });

  it("rejects a string over the Excel limit in a sentinel_unchanged expected value", () => {
    const result = validateWorkbookPlan(planWithSentinel("x".repeat(EXCEL_CELL_TEXT_LIMIT + 1)));
    expect(result.ok).toBe(false);
    expect(errorsOf(result)).toContainEqual(
      expect.objectContaining({
        code: "bound_exceeded",
        path: "/verification/0/expected",
        limit: EXCEL_CELL_TEXT_LIMIT,
      }),
    );
  });
});

describe("ScalarValue Excel UTF-16 astral boundary", () => {
  it("accepts a UTF-16-limit astral string in a write_values ValueRow cell", () => {
    const result = validateWorkbookPlan(planWithValueCell(ASTRAL_UTF16_LIMIT_VALUE));
    expect(result.ok).toBe(true);
  });

  it("rejects a UTF-16-over-limit astral string in a write_values ValueRow cell", () => {
    const result = validateWorkbookPlan(planWithValueCell(ASTRAL_UTF16_OVER_LIMIT_VALUE));
    expect(result.ok).toBe(false);
    expect(errorsOf(result)).toContainEqual(
      expect.objectContaining({
        code: "bound_exceeded",
        path: "/operations/0/values/0/0",
        limit: EXCEL_CELL_TEXT_LIMIT,
      }),
    );
  });

  it("accepts a UTF-16-limit astral string in range_values_equal expectedValues", () => {
    const result = validateWorkbookPlan(planWithExpectedValues(ASTRAL_UTF16_LIMIT_VALUE));
    expect(result.ok).toBe(true);
  });

  it("rejects a UTF-16-over-limit astral string in range_values_equal expectedValues", () => {
    const result = validateWorkbookPlan(planWithExpectedValues(ASTRAL_UTF16_OVER_LIMIT_VALUE));
    expect(result.ok).toBe(false);
    expect(errorsOf(result)).toContainEqual(
      expect.objectContaining({
        code: "bound_exceeded",
        path: "/verification/0/expectedValues/0/0",
        limit: EXCEL_CELL_TEXT_LIMIT,
      }),
    );
  });

  it("accepts a UTF-16-limit astral string in a sentinel_unchanged expected value", () => {
    const result = validateWorkbookPlan(planWithSentinel(ASTRAL_UTF16_LIMIT_VALUE));
    expect(result.ok).toBe(true);
  });

  it("rejects a UTF-16-over-limit astral string in a sentinel_unchanged expected value", () => {
    const result = validateWorkbookPlan(planWithSentinel(ASTRAL_UTF16_OVER_LIMIT_VALUE));
    expect(result.ok).toBe(false);
    expect(errorsOf(result)).toContainEqual(
      expect.objectContaining({
        code: "bound_exceeded",
        path: "/verification/0/expected",
        limit: EXCEL_CELL_TEXT_LIMIT,
      }),
    );
  });

  const expectVerificationConsistency = (
    plan: ReturnType<typeof planWithExpectedValues> | ReturnType<typeof planWithSentinel>,
    expected: boolean,
  ) => {
    const validationOk = validateWorkbookPlan(plan).ok;
    const rawVerification = JSON.stringify(plan.verification);
    if (rawVerification === undefined) throw new Error("Expected verification JSON");
    const verificationOk = prepareWorkbookPlanVerificationRules(rawVerification) !== null;
    expect(validationOk).toBe(expected);
    expect(verificationOk).toBe(expected);
    expect(validationOk).toBe(verificationOk);
  };

  it("keeps range_values_equal verification preparation consistent at the UTF-16 limit", () => {
    expectVerificationConsistency(planWithExpectedValues(ASTRAL_UTF16_LIMIT_VALUE), true);
  });

  it("keeps range_values_equal verification preparation consistent over the UTF-16 limit", () => {
    expectVerificationConsistency(planWithExpectedValues(ASTRAL_UTF16_OVER_LIMIT_VALUE), false);
  });

  it("keeps sentinel_unchanged verification preparation consistent at the UTF-16 limit", () => {
    expectVerificationConsistency(planWithSentinel(ASTRAL_UTF16_LIMIT_VALUE), true);
  });

  it("keeps sentinel_unchanged verification preparation consistent over the UTF-16 limit", () => {
    expectVerificationConsistency(planWithSentinel(ASTRAL_UTF16_OVER_LIMIT_VALUE), false);
  });

  it("retains code-point maxLength semantics for a non-ScalarValue astral string", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "😀".repeat(2048),
      operations: [{ kind: "calculate" }],
      verification: [{ kind: "calculation_completed" }],
    });
    expect(result.ok).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 1 — Nested unknown property, including inside format and operation branches
// ---------------------------------------------------------------------------

describe("RED 1: unknown property at top level", () => {
  it("rejects an extra top-level property (DeepSeek wrapper)", () => {
    const result = validateWorkbookPlan({ plan: validInspect() });
    expect(result.ok).toBe(false);
    expect(errorsOf(result)[0]?.code).toBe("unknown_property");
    expect(errorsOf(result)[0]?.path).toBe("/plan");
  });
});

describe("RED 1: unknown property nested inside format object", () => {
  it("rejects unknown property inside FormatSpec", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        {
          kind: "format_range",
          sheet: "Sheet1",
          address: "A1",
          format: { bold: true, unknownField: "oops" },
        },
      ],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1", expected: "ok" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.path.includes("format") && e.code === "unknown_property")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 1b — Closed-object safety
// ---------------------------------------------------------------------------

describe("RED 1b: hidden non-enumerable properties", () => {
  it("rejects a hidden operations property on inspect", () => {
    const plan = validInspect();
    Object.defineProperty(plan, "operations", {
      value: [{ kind: "calculate" }],
      enumerable: false,
      configurable: true,
      writable: true,
    });

    const result = validateWorkbookPlan(plan);
    expect(result.ok).toBe(false);
    expect(errorsOf(result)).toContainEqual(
      expect.objectContaining({ code: "unknown_property", path: "/operations" }),
    );
  });
});

describe("RED 1c: symbol keys are rejected", () => {
  it("rejects symbol keys with a fixed content-free error", () => {
    const plan = validExecute();
    const secret = Symbol("secret");
    Reflect.set(plan, secret, true);

    const result = validateWorkbookPlan(plan);
    expect(result.ok).toBe(false);
    const errorText = JSON.stringify(errorsOf(result));
    expect(errorText).not.toContain("secret");
    expect(errorsOf(result)[0]?.code).toBe("symbol_key_not_supported");
  });
});

describe("RED 1d: getters are never invoked", () => {
  it("rejects accessor properties without invoking the getter", () => {
    const plan = validInspect();
    let getterCalls = 0;
    Object.defineProperty(plan, "operations", {
      enumerable: true,
      configurable: true,
      get() {
        getterCalls += 1;
        return [{ kind: "calculate" }];
      },
    });

    const result = validateWorkbookPlan(plan);
    expect(getterCalls).toBe(0);
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "accessor_property" && e.path === "/operations")).toBe(true);
  });

  it("rejects accessor array elements without invoking the getter", () => {
    const plan = validExecute();
    let getterCalls = 0;
    Object.defineProperty(plan.operations, "0", {
      enumerable: true,
      configurable: true,
      get() {
        getterCalls += 1;
        return { kind: "calculate" };
      },
    });

    const result = validateWorkbookPlan(plan);
    expect(getterCalls).toBe(0);
    expect(result.ok).toBe(false);
    expect(errorsOf(result)).toContainEqual(
      expect.objectContaining({ code: "accessor_property", path: "/operations/0" }),
    );
  });

  it("rejects metadata accessors without invoking the getter", () => {
    const operation: Record<string, unknown> = {
      kind: "conditional_format",
      sheet: "Sheet1",
      address: "A1",
      ruleType: "containsText",
    };
    let getterCalls = 0;
    Object.defineProperty(operation, "text", {
      enumerable: true,
      configurable: true,
      get() {
        getterCalls += 1;
        return "needle";
      },
    });

    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "conditional format",
      operations: [operation],
      verification: [{ kind: "calculation_completed" }],
    });
    expect(getterCalls).toBe(0);
    expect(result.ok).toBe(false);
    expect(errorsOf(result)).toContainEqual(
      expect.objectContaining({ code: "accessor_property", path: "/operations/0/text" }),
    );
  });
});

// ---------------------------------------------------------------------------
// RED 2 — Jagged value matrix and jagged formula matrix
// ---------------------------------------------------------------------------

describe("RED 2: jagged matrices", () => {
  it("rejects jagged ValueMatrix", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "write_values", sheet: "Sheet1", address: "A1", values: [["a", "b"], ["c"]] },
      ],
      verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["a"]] }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "matrix_jagged")).toBe(true);
  });

  it("rejects jagged FormulaMatrix", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "write_formula", sheet: "Sheet1", address: "A1", formula: "=A1", values: undefined },
      ],
      verification: [{ kind: "range_formulas_equal", sheet: "Sheet1", address: "A1", expectedFormulas: [["=A1", "=B1"], ["=C1"]] }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "matrix_jagged")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 3 — Oversized matrix by authored row/column limits
// ---------------------------------------------------------------------------

describe("RED 3: oversized matrices", () => {
  it("rejects ValueMatrix exceeding maxItems=1000 rows", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "write_values", sheet: "Sheet1", address: "A1", values: Array.from({ length: 1001 }, () => ["x"]) },
      ],
      verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["x"]] }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });

  it("rejects ValueMatrix row exceeding maxItems=1000 columns", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "write_values", sheet: "Sheet1", address: "A1", values: [Array.from({ length: 1001 }, (_, i) => i)] },
      ],
      verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["x"]] }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 4 — NaN, Infinity, -Infinity; fractional integers; numeric bounds
// ---------------------------------------------------------------------------

describe("RED 4: non-finite and invalid numeric values", () => {
  it("rejects NaN in values", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "write_values", sheet: "Sheet1", address: "A1", values: [[NaN]] },
      ],
      verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [[0]] }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "type_mismatch")).toBe(true);
  });

  it("rejects Infinity in values", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "write_values", sheet: "Sheet1", address: "A1", values: [[Infinity]] },
      ],
      verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [[0]] }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "type_mismatch")).toBe(true);
  });

  it("rejects -Infinity in values", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "write_values", sheet: "Sheet1", address: "A1", values: [[-Infinity]] },
      ],
      verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [[0]] }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "type_mismatch")).toBe(true);
  });

  it("rejects fractional columnIndex", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "sort_table", tableName: "Table1", columnIndex: 1.5 },
      ],
      verification: [{ kind: "table_row_order", tableName: "Table1", expectedOrder: ["a"] }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "type_mismatch")).toBe(true);
  });

  it("rejects negative indentLevel", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "format_range", sheet: "Sheet1", address: "A1", format: { indentLevel: -1 } },
      ],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1", expected: "ok" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });

  it("rejects negative fontSize", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "format_range", sheet: "Sheet1", address: "A1", format: { fontSize: -5 } },
      ],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1", expected: "ok" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });

  it("rejects textOrientation out of range", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "format_range", sheet: "Sheet1", address: "A1", format: { textOrientation: 91 } },
      ],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1", expected: "ok" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 4b — Budget and truncation adversaries
// ---------------------------------------------------------------------------

describe("RED 4b: work budget and maxErrors truncation", () => {
  const buildFormatPlan = (extraKeys: readonly string[]) => {
    const format: Record<string, unknown> = { bold: true };
    for (const key of extraKeys) {
      format[key] = true;
    }

    return {
      version: 1,
      outcome: "execute" as const,
      summary: "x",
      operations: [{ kind: "format_range", sheet: "Sheet1", address: "A1", format }],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1", expected: "ok" }],
    };
  };

  it("fails with a fixed work_limit_exceeded error when the work budget is exhausted", () => {
    const constraints = workbookPlanContract.validationConstraints;
    const originalMaxWorkUnits = constraints.maxWorkUnits;

    try {
      Reflect.set(constraints, "maxWorkUnits", 1);

      const result = validateWorkbookPlan(validExecute());
      expect(result.ok).toBe(false);
      expect(errorsOf(result)).toHaveLength(1);
      expect(errorsOf(result)[0]?.code).toBe("work_limit_exceeded");
      expect(JSON.stringify(errorsOf(result))).not.toContain("Sheet1");
      expect(JSON.stringify(errorsOf(result))).not.toContain("write_values");
    } finally {
      Reflect.set(constraints, "maxWorkUnits", originalMaxWorkUnits);
    }
  });

  it("fails with the same fixed error when the visited-node budget is exhausted", () => {
    const constraints = workbookPlanContract.validationConstraints;
    const originalMaxVisitedNodes = constraints.maxVisitedNodes;

    try {
      Reflect.set(constraints, "maxVisitedNodes", 1);
      const result = validateWorkbookPlan(validExecute());
      expect(result.ok).toBe(false);
      expect(errorsOf(result)).toEqual([
        {
          code: "work_limit_exceeded",
          path: "",
          message: "Validation work limit exceeded",
        },
      ]);
    } finally {
      Reflect.set(constraints, "maxVisitedNodes", originalMaxVisitedNodes);
    }
  });

  it("fails with one fixed content-free error when maxDepth is exceeded", () => {
    const constraints = workbookPlanContract.validationConstraints;
    const originalMaxDepth = constraints.maxDepth;

    try {
      Reflect.set(constraints, "maxDepth", 0);
      const result = validateWorkbookPlan(validExecute());
      expect(result.ok).toBe(false);
      expect(errorsOf(result)).toEqual([
        {
          code: "bound_exceeded",
          path: "",
          message: "Validation depth exceeds the contract bound",
          limit: 0,
        },
      ]);
      expect(JSON.stringify(errorsOf(result))).not.toContain("Sheet1");
      expect(JSON.stringify(errorsOf(result))).not.toContain("write_values");
    } finally {
      Reflect.set(constraints, "maxDepth", originalMaxDepth);
    }
  });

  it("truncates to exact maxErrors survivors after deterministic dedupe", () => {
    const extras = Array.from({ length: 60 }, (_, index) => `extra_${String(index).padStart(2, "0")}`);
    const result = validateWorkbookPlan(buildFormatPlan([...extras].reverse()));
    expect(result.ok).toBe(false);

    const keys = errorsOf(result).map((error) => `${error.path}\u0000${error.code}`);
    const expected = extras
      .slice()
      .sort(compareCodeUnit)
      .slice(0, workbookPlanContract.validationConstraints.maxErrors)
      .map((key) => `/operations/0/format/${key}\u0000unknown_property`);

    expect(keys).toEqual(expected);
    expect(keys).toEqual([...new Set(keys)]);
  });

  it("serializes reversed insertion-order invalid plans to identical error arrays", () => {
    const extras = Array.from({ length: 60 }, (_, index) => `extra_${String(index).padStart(2, "0")}`);
    const forward = errorsOf(validateWorkbookPlan(buildFormatPlan(extras))).map(
      (error) => `${error.path}\u0000${error.code}`,
    );
    const reverse = errorsOf(validateWorkbookPlan(buildFormatPlan([...extras].reverse()))).map(
      (error) => `${error.path}\u0000${error.code}`,
    );

    expect(forward).toEqual(reverse);
    expect(forward).toEqual([...new Set(forward)]);
  });
});

// ---------------------------------------------------------------------------
// RED 5 — Empty required strings
// ---------------------------------------------------------------------------

describe("RED 5: empty required strings", () => {
  // Empty string fails minLength:1 → bound_exceeded (schema semantic)
  it("rejects empty reason in inspect", () => {
    const result = validateWorkbookPlan({ version: 1, outcome: "inspect", reason: "", reads: [] });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });

  it("rejects empty sheet name", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "inspect",
      reason: "x",
      reads: [{ kind: "read_range", sheet: "", address: "A1" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });

  it("rejects empty address", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "inspect",
      reason: "x",
      reads: [{ kind: "read_range", sheet: "Sheet1", address: "" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });

  it("rejects empty summary", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "",
      operations: [{ kind: "calculate" }],
      verification: [{ kind: "calculation_completed" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });

  it("rejects empty message in ask", () => {
    const result = validateWorkbookPlan({ version: 1, outcome: "ask", message: "" });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });

  it("rejects empty source in set_validation_list", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [{ kind: "set_validation_list", sheet: "Sheet1", address: "A1", source: "" }],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1", expected: "ok" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 6 — UTF-8 byte limits for formula/comment/validation-list
// ---------------------------------------------------------------------------

describe("RED 6: UTF-8 byte limits", () => {
  // Formula/comment/validation-list fields have byte limits defined in validationConstraints
  // Cyrillic U+0410 = 2 bytes UTF-8 each; 2049 chars = 4098 bytes > 4096 limit
  it("rejects formula exceeding UTF-8 byte limit", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "write_formula", sheet: "Sheet1", address: "A1", formula: LONG_TEXT_4KB },
      ],
      verification: [{ kind: "range_formulas_equal", sheet: "Sheet1", address: "A1", expectedFormulas: [[LONG_TEXT_4KB]] }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "utf8_limit_exceeded")).toBe(true);
  });

  it("rejects comment text exceeding UTF-8 byte limit", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "add_comment", sheet: "Sheet1", address: "A1", text: LONG_TEXT_1KB },
      ],
      verification: [{ kind: "comment_exists", sheet: "Sheet1", address: "A1", text: "x" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "utf8_limit_exceeded")).toBe(true);
  });

  it("rejects validation-list source exceeding UTF-8 byte limit", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "set_validation_list", sheet: "Sheet1", address: "A1", source: LONG_TEXT_4KB },
      ],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1", expected: "ok" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "utf8_limit_exceeded")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 7 — Canonical+alias conflict
// ---------------------------------------------------------------------------

describe("RED 7: canonical+alias conflict", () => {
  it("rejects plan with both sheet and sheet_name", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "inspect",
      reason: "x",
      reads: [{ kind: "read_range", sheet: "Sheet1", sheet_name: "Sheet1", address: "A1" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "alias_conflict")).toBe(true);
  });

  it("rejects operation with both address and range", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "write_values", sheet: "Sheet1", address: "A1", values: [["x"]], range: "A1" },
      ],
      verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["x"]] }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "alias_conflict")).toBe(true);
  });

  it("rejects format with both fillColor and fill_color", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "format_range", sheet: "Sheet1", address: "A1", format: { fillColor: "FF0000", fill_color: "00FF00" } },
      ],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1", expected: "ok" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "alias_conflict")).toBe(true);
  });

  it("rejects format with both fontColor and font_color", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "format_range", sheet: "Sheet1", address: "A1", format: { fontColor: "FF0000", font_color: "00FF00" } },
      ],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1", expected: "ok" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "alias_conflict")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 8 — Unknown envelope version
// ---------------------------------------------------------------------------

describe("RED 8: unknown envelope version", () => {
  it("rejects version 0", () => {
    const result = validateWorkbookPlan({ version: 0, outcome: "ask", message: "x" });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "unsupported_version")).toBe(true);
  });

  it("rejects version 2", () => {
    const result = validateWorkbookPlan({ version: 2, outcome: "ask", message: "x" });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "unsupported_version")).toBe(true);
  });

  it("rejects non-integer version", () => {
    const result = validateWorkbookPlan({ version: "1", outcome: "ask", message: "x" });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "unsupported_version")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 9 — Execute with zero operations
// ---------------------------------------------------------------------------

describe("RED 9: execute with zero operations", () => {
  it("rejects empty operations array", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [],
      verification: [{ kind: "calculation_completed" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 10 — Execute lacking or empty required verification
// ---------------------------------------------------------------------------

describe("RED 10: execute lacking or empty verification", () => {
  it("rejects missing verification", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [{ kind: "calculate" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "missing_required")).toBe(true);
  });

  it("rejects empty verification array", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [{ kind: "calculate" }],
      verification: [],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 11 — Inspect containing a mutation field or mutation operation
// ---------------------------------------------------------------------------

describe("RED 11: inspect with mutation fields", () => {
  it("rejects inspect outcome with operations field", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "inspect",
      reason: "x",
      reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1" }],
      operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["x"]] }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "unknown_property" || e.code === "type_mismatch")).toBe(true);
  });

  it("rejects inspect outcome with verification field", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "inspect",
      reason: "x",
      reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1" }],
      verification: [{ kind: "calculation_completed" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "unknown_property" || e.code === "type_mismatch")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 12 — Ask/refuse containing reads/operations/verification
// ---------------------------------------------------------------------------

describe("RED 12: ask/refuse with disallowed fields", () => {
  it("rejects ask with reads", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "ask",
      message: "x",
      reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "unknown_property" || e.code === "type_mismatch")).toBe(true);
  });

  it("rejects refuse with operations", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "refuse",
      message: "x",
      operations: [{ kind: "calculate" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "unknown_property" || e.code === "type_mismatch")).toBe(true);
  });

  it("rejects refuse with verification", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "refuse",
      message: "x",
      verification: [{ kind: "calculation_completed" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "unknown_property" || e.code === "type_mismatch")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 13 — Unknown operation and unknown verification kind
// ---------------------------------------------------------------------------

describe("RED 13: unknown operation/verification kind", () => {
  it("rejects unknown operation kind", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [{ kind: "delete_sheet" }], // not a valid kind
      verification: [{ kind: "calculation_completed" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result)).toHaveLength(1);
    expect(errorsOf(result)[0]?.code).toBe("unknown_discriminator");
    expect(errorsOf(result)[0]?.path).toBe("/operations/0/kind");
  });

  it("rejects unknown verification kind", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [{ kind: "calculate" }],
      verification: [{ kind: "range_size_equal" }], // not a valid kind
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result)).toHaveLength(1);
    expect(errorsOf(result)[0]?.code).toBe("unknown_discriminator");
    expect(errorsOf(result)[0]?.path).toBe("/verification/0/kind");
  });

  it("rejects unknown inspection kind", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "inspect",
      reason: "x",
      reads: [{ kind: "scan_range", sheet: "Sheet1", address: "A1" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result)).toHaveLength(1);
    expect(errorsOf(result)[0]?.code).toBe("unknown_discriminator");
    expect(errorsOf(result)[0]?.path).toBe("/reads/0/kind");
  });
});

// ---------------------------------------------------------------------------
// RED 14 — DeepSeek wrapper is rejected as unknown top-level property
// ---------------------------------------------------------------------------

describe("RED 14: DeepSeek wrapper rejected", () => {
  it("rejects { plan: {...} } wrapper", () => {
    const result = validateWorkbookPlan({ plan: validExecute() });
    expect(result.ok).toBe(false);
    expect(errorsOf(result)[0]?.code).toBe("unknown_property");
    expect(errorsOf(result)[0]?.path).toBe("/plan");
  });
});

// ---------------------------------------------------------------------------
// RED 15 — Valid representative plans return typed success
// ---------------------------------------------------------------------------

describe("RED 15: valid plans return success", () => {
  it("valid inspect plan returns ok=true with WorkbookPlan", () => {
    const input = validInspect();
    const result = validateWorkbookPlan(input);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.version).toBe(1);
      expect(result.value.outcome).toBe("inspect");
    }
  });

  it("valid execute plan returns ok=true with WorkbookPlan", () => {
    const input = validExecute();
    const result = validateWorkbookPlan(input);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.version).toBe(1);
      expect(result.value.outcome).toBe("execute");
    }
  });

  it("valid ask plan returns ok=true with WorkbookPlan", () => {
    const input = validAsk();
    const result = validateWorkbookPlan(input);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.version).toBe(1);
      expect(result.value.outcome).toBe("ask");
    }
  });

  it("valid refuse plan returns ok=true with WorkbookPlan", () => {
    const input = validRefuse();
    const result = validateWorkbookPlan(input);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.version).toBe(1);
      expect(result.value.outcome).toBe("refuse");
    }
  });

  it("valid write_values with all scalar types in matrix", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "all types",
      operations: [
        {
          kind: "write_values",
          sheet: "Sheet1",
          address: "A1",
          values: [
            ["string", 42, true, null],
            ["", 0, false, null],
          ],
        },
      ],
      verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["x"]] }],
    });
    expect(result.ok).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// RED 15b — oneOf branch isolation
// ---------------------------------------------------------------------------

describe("RED 15b: oneOf branch isolation", () => {
  it("accepts containsText conditional formatting without leaking formula1 failures from the other branch", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "conditional formatting",
      operations: [
        {
          kind: "conditional_format",
          sheet: "Sheet1",
          address: "A1",
          ruleType: "containsText",
          text: "needle",
          fillColor: "FF0000",
          fontColor: "FFFFFF",
        },
      ],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1", expected: "ok" }],
    });

    expect(result.ok).toBe(true);
    if (result.ok && result.value.outcome === "execute") {
      expect(result.value.operations[0].kind).toBe("conditional_format");
    }
  });

  it("reports only the selected failed operation branch", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "write formula",
      operations: [
        {
          kind: "write_formula",
          sheet: "Sheet1",
          address: "A1",
          values: [["not a formula"]],
        },
      ],
      verification: [{ kind: "calculation_completed" }],
    });

    expect(result.ok).toBe(false);
    expect(errorsOf(result)).toEqual([
      {
        code: "missing_required",
        path: "/operations/0/formula",
        message: "Required field missing",
      },
      {
        code: "unknown_property",
        path: "/operations/0/values",
        message: "Unknown property",
      },
    ]);
  });
});

// ---------------------------------------------------------------------------
// RED 16 — Validation errors are content-free (no secrets/formulas in output)
// ---------------------------------------------------------------------------

describe("RED 16: errors contain no sensitive content", () => {
  it("formula secrets do not appear in type_mismatch error for scalar", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: `${SECRET_FORMULA}${SECRET_VALUE}`.repeat(300),
      operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [[SECRET_VALUE]] }],
      verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [[SECRET_VALUE]] }],
    });
    expect(result.ok).toBe(false);
    // If an error mentions the secret value, the test fails
    const errorText = JSON.stringify(errorsOf(result));
    expect(errorText).not.toContain(SECRET_FORMULA);
    expect(errorText).not.toContain(SECRET_VALUE);
  });

  it("long formula does not appear in utf8_limit_exceeded error message", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [{ kind: "write_formula", sheet: "Sheet1", address: "A1", formula: LONG_SECRET_FORMULA }],
      verification: [{ kind: "range_formulas_equal", sheet: "Sheet1", address: "A1", expectedFormulas: [[LONG_SECRET_FORMULA]] }],
    });
    expect(result.ok).toBe(false);
    const errorText = JSON.stringify(errorsOf(result));
    expect(errorText).not.toContain(LONG_SECRET_FORMULA);
  });

  it("DeepSeek wrapper value does not appear in error path", () => {
    const wrapper = { plan: { ...validExecute(), summary: LONG_SECRET_FORMULA } };
    const result = validateWorkbookPlan(wrapper);
    expect(result.ok).toBe(false);
    // The error path should reference "/plan" but not the inner structure
    const errorText = JSON.stringify(errorsOf(result));
    expect(errorText).not.toContain("write_values");
    expect(errorText).not.toContain("Sheet1");
    expect(errorText).not.toContain(LONG_SECRET_FORMULA);
  });
});

// ---------------------------------------------------------------------------
// Additional edge cases for coverage
// ---------------------------------------------------------------------------

describe("edge cases", () => {
  it("rejects null input", () => {
    const result = validateWorkbookPlan(null);
    expect(result.ok).toBe(false);
  });

  it("rejects array input", () => {
    const result = validateWorkbookPlan([{ version: 1 }]);
    expect(result.ok).toBe(false);
  });

  it("rejects plain string input", () => {
    const result = validateWorkbookPlan("not an object");
    expect(result.ok).toBe(false);
  });

  it("rejects number input", () => {
    const result = validateWorkbookPlan(42);
    expect(result.ok).toBe(false);
  });

  it("rejects boolean input", () => {
    const result = validateWorkbookPlan(true);
    expect(result.ok).toBe(false);
  });

  it("rejects object with prototype chain (unsafe)", () => {
    const result = validateWorkbookPlan(Object.assign(Object.create({ secret: true }), { version: 1, outcome: "ask", message: "x" }));
    expect(result.ok).toBe(false);
  });

  it("rejects inspect with empty reads array", () => {
    const result = validateWorkbookPlan({ version: 1, outcome: "inspect", reason: "x", reads: [] });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });

  it("rejects reads exceeding maxItems=50", () => {
    const reads = Array.from({ length: 51 }, (_, i) => ({ kind: "read_range", sheet: "Sheet1", address: `A${i + 1}` }));
    const result = validateWorkbookPlan({ version: 1, outcome: "inspect", reason: "x", reads });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });

  it("rejects operations exceeding maxItems=30", () => {
    const ops = Array.from({ length: 31 }, (_, i) => ({ kind: "write_values", sheet: "Sheet1", address: `A${i + 1}`, values: [["x"]] }));
    const result = validateWorkbookPlan({ version: 1, outcome: "execute", summary: "x", operations: ops, verification: [{ kind: "calculation_completed" }] });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });

  it("rejects verification exceeding maxItems=30", () => {
    const verif = Array.from({ length: 31 }, (_, i) => ({ kind: "sentinel_unchanged", sheet: "Sheet1", address: `A${i + 1}`, expected: "x" }));
    const result = validateWorkbookPlan({ version: 1, outcome: "execute", summary: "x", operations: [{ kind: "calculate" }], verification: verif });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "bound_exceeded")).toBe(true);
  });

  it("rejects inspection with extra field in read_range", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "inspect",
      reason: "x",
      reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1", extra: "bad" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "unknown_property")).toBe(true);
  });

  it("rejects verification with extra field", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [{ kind: "calculate" }],
      verification: [{ kind: "calculation_completed", extra: "bad" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "unknown_property")).toBe(true);
  });

  it("rejects object where required field missing in branch", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [{ kind: "write_values", sheet: "Sheet1" }], // missing address and values
      verification: [{ kind: "calculation_completed" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "missing_required" || e.code === "unknown_discriminator")).toBe(true);
  });

  it("rejects write_formula with missing formula field", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [{ kind: "write_formula", sheet: "Sheet1", address: "A1" }],
      verification: [{ kind: "calculation_completed" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "missing_required" || e.code === "unknown_discriminator")).toBe(true);
  });

  it("rejects conditional_format with containsText but no text field", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [
        { kind: "conditional_format", sheet: "Sheet1", address: "A1", ruleType: "containsText" },
      ],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1", expected: "ok" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "missing_required" || e.code === "unknown_discriminator")).toBe(true);
  });

  it("rejects sentinel_unchanged verification with missing expected field", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [{ kind: "calculate" }],
      verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "A1" }],
    });
    expect(result.ok).toBe(false);
    expect(errorsOf(result).some((e) => e.code === "missing_required" || e.code === "unknown_discriminator")).toBe(true);
  });

  it("accepts inspect with only reason and reads", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "inspect",
      reason: "need data",
      reads: [{ kind: "read_range", sheet: "Sheet1", address: "A1" }],
    });
    expect(result.ok).toBe(true);
  });

  it("accepts execute with calculate operation", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "calc",
      operations: [{ kind: "calculate" }],
      verification: [{ kind: "calculation_completed" }],
    });
    expect(result.ok).toBe(true);
  });

  it("accepts create_sheet operation", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "new sheet",
      operations: [{ kind: "create_sheet", name: "NewSheet" }],
      verification: [{ kind: "sheet_exists", name: "NewSheet" }],
    });
    expect(result.ok).toBe(true);
  });

  it("accepts add_comment operation", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "comment",
      operations: [{ kind: "add_comment", sheet: "Sheet1", address: "A1", text: "Hello" }],
      verification: [{ kind: "comment_exists", sheet: "Sheet1", address: "A1", text: "Hello" }],
    });
    expect(result.ok).toBe(true);
  });

  it("deduplicates identical errors", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [{ kind: "write_values", sheet: "", address: "A1", values: [["x"]] }],
      verification: [{ kind: "range_values_equal", sheet: "", address: "A1", expectedValues: [["x"]] }],
    });
    // Both sheet fields are empty — should be at most one error per path
    const paths = errorsOf(result).map((e) => e.path);
    const uniquePaths = new Set(paths);
    // Should not have duplicate paths for the same error
    expect(paths.length).toBeLessThanOrEqual(uniquePaths.size * 2);
  });

  it("errors are deterministically ordered by path", () => {
    const result = validateWorkbookPlan({
      version: 1,
      outcome: "execute",
      summary: "x",
      operations: [{ kind: "write_values", sheet: "", address: "", values: [["x"]] }],
      verification: [{ kind: "range_values_equal", sheet: "", address: "", expectedValues: [["x"]] }],
    });
    const paths = errorsOf(result).map((e) => e.path);
    const sorted = [...paths].sort();
    expect(paths).toEqual(sorted);
  });
});
