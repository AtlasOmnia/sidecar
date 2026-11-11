import { describe, expect, it } from "vitest";
import {
  MAX_VERIFICATION_RULES,
  BoundWorkbookReadbackAdapter,
  PreparedWorkbookPlanVerificationRules,
  VerificationRule,
  WorkbookSnapshot,
  VerificationInverseEvidence,
  derivePostVerificationControl,
  prepareWorkbookPlanVerificationRules,
  verifyWorkbookPlan,
} from "../workbookPlanVerification";

type SnapshotResult = { readonly status: "supported"; readonly snapshot: WorkbookSnapshot };

const prepared = (rules: readonly VerificationRule[]): PreparedWorkbookPlanVerificationRules => {
  const token = prepareWorkbookPlanVerificationRules(JSON.stringify(rules));
  expect(token).not.toBeNull();
  return token!;
};

const readback = (snapshots: readonly WorkbookSnapshot[]): BoundWorkbookReadbackAdapter => {
  let index = 0;
  return {
    read: () => {
      const snapshot = snapshots[index];
      index += 1;
      return snapshot === undefined ? undefined : JSON.stringify({ status: "supported", snapshot } satisfies SnapshotResult);
    },
  };
};

const verify = (rule: VerificationRule, snapshot: WorkbookSnapshot): ReturnType<typeof verifyWorkbookPlan> =>
  verifyWorkbookPlan(prepared([rule]), readback([snapshot]));

const rules: readonly VerificationRule[] = [
  { kind: "range_values_equal", sheet: "Data", address: "A1:B2", expectedValues: [[1, "x"], [true, null]] },
  { kind: "range_formulas_equal", sheet: "Data", address: "A1:B1", expectedFormulas: [["=A2", "=B2"]] },
  { kind: "range_formula_pattern", sheet: "Data", address: "A1:A2", pattern: "=SUM({{cell}})" },
  { kind: "range_number_format_equal", sheet: "Data", address: "A1:B1", expectedNumberFormat: "0.00" },
  { kind: "table_exists", tableName: "Sales" },
  { kind: "table_row_order", tableName: "Sales", expectedOrder: ["r1", "r2"] },
  { kind: "table_filter_state", tableName: "Sales", columnIndex: 1, state: "filtered", values: ["Open", "Closed"], criterion1: "Open", operator: "Or" },
  { kind: "sheet_exists", name: "Data" },
  { kind: "chart_exists", sheet: "Data", title: "Revenue" },
  { kind: "comment_exists", sheet: "Data", address: "A1", text: "Note" },
  { kind: "sentinel_unchanged", sheet: "Data", address: "Z1", expected: 42 },
  { kind: "calculation_completed" },
];

const snapshots: readonly WorkbookSnapshot[] = [
  { kind: "range_values", values: [[1, "x"], [true, null]] },
  { kind: "range_formulas", formulas: [["=A2", "=B2"]] },
  { kind: "range_formulas", formulas: [["=SUM(A1)"], ["=SUM(B2)"]] },
  { kind: "range_number_format", numberFormat: [["0.00", "0.00"]] },
  { kind: "table_exists", exists: true },
  { kind: "table_row_order", order: ["r1", "r2"] },
  { kind: "table_filter_state", state: { columnIndex: 1, state: "filtered", values: ["Open", "Closed"], criterion1: "Open", operator: "Or" } },
  { kind: "sheet_exists", exists: true },
  { kind: "chart_exists", exists: true },
  { kind: "comment_exists", exists: true, text: "Note" },
  { kind: "sentinel", value: 42 },
  { kind: "calculation_completed", completed: true },
];

const valuesRule = rules[0] as Extract<VerificationRule, { readonly kind: "range_values_equal" }>;
const formulasRule = rules[1] as Extract<VerificationRule, { readonly kind: "range_formulas_equal" }>;
const patternRule = rules[2] as Extract<VerificationRule, { readonly kind: "range_formula_pattern" }>;
const numberFormatRule = rules[3] as Extract<VerificationRule, { readonly kind: "range_number_format_equal" }>;
const rowOrderRule = rules[5] as Extract<VerificationRule, { readonly kind: "table_row_order" }>;
const filterRule = rules[6] as Extract<VerificationRule, { readonly kind: "table_filter_state" }>;
const sentinelRule = rules[10] as Extract<VerificationRule, { readonly kind: "sentinel_unchanged" }>;

describe("workbook plan deterministic verification", () => {
  it("verifies all twelve V1 rules through an opaque prepared capability", () => {
    const report = verifyWorkbookPlan(prepared(rules), readback(snapshots));
    expect(report.ok).toBe(true);
    expect(report.rules).toHaveLength(12);
    expect(report.counts).toEqual({ verified: 12, mismatched: 0, unsupported: 0, error: 0 });
    expect(report.firstFailure).toBeNull();
  });

  it("classifies a mismatch for every V1 rule without exposing content", () => {
    const mismatches: readonly [VerificationRule, WorkbookSnapshot][] = [
      [{ ...valuesRule, expectedValues: [[1, "different"], [true, null]] }, snapshots[0]],
      [{ ...formulasRule, expectedFormulas: [["=WRONG", "=B2"]] }, snapshots[1]],
      [{ ...patternRule, pattern: "=AVERAGE({{cell}})" }, snapshots[2]],
      [{ ...numberFormatRule, expectedNumberFormat: "0.000" }, snapshots[3]],
      [rules[4], { kind: "table_exists", exists: false }],
      [{ ...rowOrderRule, expectedOrder: ["r2", "r1"] }, snapshots[5]],
      [{ ...filterRule, state: "unfiltered", values: [] }, snapshots[6]],
      [rules[7], { kind: "sheet_exists", exists: false }],
      [rules[8], { kind: "chart_exists", exists: false }],
      [rules[9], { kind: "comment_exists", exists: false, text: "Missing" }],
      [{ ...sentinelRule, expected: "42" }, snapshots[10]],
      [rules[11], { kind: "calculation_completed", completed: false }],
    ];
    for (const [rule, snapshot] of mismatches) {
      const report = verify(rule, snapshot);
      expect(report.ok).toBe(false);
      expect(report.rules[0]?.status).toBe("mismatched");
      expect(JSON.stringify(report)).not.toContain("different");
    }
  });

  it("requires rectangular dimensions and compares cells row-major", () => {
    const dimension = verify(valuesRule, { kind: "range_values", values: [[1, "x"]] });
    expect(dimension.rules[0]?.reason).toBe("dimension_mismatch");
    const rowMajor = verify({ ...valuesRule, expectedValues: [[1, 2], [3, 4]] }, { kind: "range_values", values: [[1, 3], [2, 4]] });
    expect(rowMajor.rules[0]?.reason).toBe("value_mismatch");
  });

  it("does not coerce scalar types or fold string case", () => {
    for (const [expected, actual] of [[1, "1"], [true, 1], ["Open", "open"], [null, ""]] as const) {
      const report = verify({ kind: "sentinel_unchanged", sheet: "Data", address: "A1", expected }, { kind: "sentinel", value: actual });
      expect(report.rules[0]?.status).toBe("mismatched");
    }
  });

  it("uses bounded literal/template formula matching, not model-supplied regex", () => {
    expect(verify({ ...patternRule, address: "A1", pattern: "=SUM(A1.*)" }, { kind: "range_formulas", formulas: [["=SUM(A1.*)"]] }).ok).toBe(true);
    expect(verify({ ...patternRule, address: "A1", pattern: "=SUM(A1.*)" }, { kind: "range_formulas", formulas: [["=SUM(A100)"]] }).rules[0]?.status).toBe("mismatched");
    expect(verify({ ...patternRule, pattern: "=SUM({{cell}})" }, { kind: "range_formulas", formulas: [["=SUM($A$1)"], ["=SUM(Sheet1!B2)"]] }).ok).toBe(true);
    expect(verify({ ...patternRule, address: "A1", pattern: "=SUM({{cell}})" }, { kind: "range_formulas", formulas: [["=SUM('Sheet 1'!$B$2)"]] }).ok).toBe(true);
    for (const reference of ["Sheet1", "foo123bar", "A0", "XFE1"]) {
      expect(verify({ ...patternRule, pattern: "={{cell}}" }, { kind: "range_formulas", formulas: [[`=${reference}`]] }).rules[0]?.status).toBe("mismatched");
    }
  });

  it("allows contract-valid empty formula cells", () => {
    const rule: VerificationRule = { kind: "range_formulas_equal", sheet: "Data", address: "A1:B1", expectedFormulas: [["", "=B2"]] };
    expect(verify(rule, { kind: "range_formulas", formulas: [["", "=B2"]] }).ok).toBe(true);
    expect(verify({ ...patternRule, address: "A1", pattern: "=SUM({{cell}})" }, { kind: "range_formulas", formulas: [[""]] }).rules[0]?.status).toBe("mismatched");
  });

  it("requires formula and number-format readback dimensions to equal the authored A1 target", () => {
    expect(verify({ ...patternRule, address: "A1:B2" }, { kind: "range_formulas", formulas: [["=SUM(A1)"]] }).rules[0]?.reason).toBe("dimension_mismatch");
    expect(verify({ ...numberFormatRule, address: "A1:B2" }, { kind: "range_number_format", numberFormat: [["0.00"]] }).rules[0]?.reason).toBe("dimension_mismatch");
    expect(verify({ ...patternRule, address: "$A$1:$B$2" }, { kind: "range_formulas", formulas: [["=SUM(A1)", "=SUM(B1)"], ["=SUM(A2)", "=SUM(B2)"]] }).ok).toBe(true);
  });

  it("validates every A1 target before any adapter read", () => {
    for (const rule of [
      { ...valuesRule, address: "B2:A1" },
      { ...formulasRule, address: "A0" },
      { ...patternRule, address: "A1:B2:C3" },
      { ...numberFormatRule, address: "A1:B1001" },
      { kind: "comment_exists", sheet: "Data", address: "A1:B2", text: "Note" } as VerificationRule,
      { kind: "sentinel_unchanged", sheet: "Data", address: "not-a-range", expected: 1 } as VerificationRule,
    ]) {
      let calls = 0;
      const token = prepareWorkbookPlanVerificationRules(JSON.stringify([rule]));
      expect(token).toBeNull();
      const report = verifyWorkbookPlan(token as unknown as PreparedWorkbookPlanVerificationRules, { read: () => { calls += 1; return undefined; } });
      expect(report.ok).toBe(false);
      expect(calls).toBe(0);
    }
  });

  it("compares comments and filters exactly without reporting content", () => {
    const wrongText = verify(rules[9], { kind: "comment_exists", exists: true, text: "Different" });
    expect(wrongText.rules[0]?.status).toBe("mismatched");
    expect(JSON.stringify(wrongText)).not.toContain("Different");
    expect(verify(filterRule, { kind: "table_filter_state", state: { operator: "Or", criterion1: "Open", values: ["Open", "Closed"], state: "filtered", columnIndex: 1 } }).ok).toBe(true);
    expect(verify(filterRule, { kind: "table_filter_state", state: { columnIndex: 1, state: "filtered", values: ["Closed", "Open"], criterion1: "Open", operator: "Or" } }).rules[0]?.reason).toBe("filter_state_mismatch");
  });

  it("accepts contract-valid empty optional filter criteria through authored and readback boundaries", () => {
    const emptyCriteriaRule = {
      kind: "table_filter_state",
      tableName: "Sales",
      columnIndex: 1,
      state: "filtered",
      criterion1: "",
      criterion2: "",
    } as const;
    const token = prepareWorkbookPlanVerificationRules(JSON.stringify([emptyCriteriaRule]));
    expect(token).not.toBeNull();
    const report = verifyWorkbookPlan(token!, {
      read: () => JSON.stringify({
        status: "supported",
        snapshot: {
          kind: "table_filter_state",
          state: { columnIndex: 1, state: "filtered", criterion1: "", criterion2: "" },
        },
      }),
    });
    expect(report.ok).toBe(true);
    expect(report.rules[0]?.status).toBe("verified");
    expect(report.counts).toEqual({ verified: 1, mismatched: 0, unsupported: 0, error: 0 });
  });

  it("distinguishes empty and nonempty optional filter criteria and rejects unknown fields deterministically", () => {
    const emptyCriteriaRule = {
      kind: "table_filter_state",
      tableName: "Sales",
      columnIndex: 1,
      state: "filtered",
      criterion1: "",
      criterion2: "",
    } as const;
    const nonemptyCriteriaRule = { ...emptyCriteriaRule, criterion1: "Open", criterion2: "Closed" } as const;
    const mismatch = (rule: typeof emptyCriteriaRule | typeof nonemptyCriteriaRule, state: Record<string, unknown>) => verifyWorkbookPlan(
      prepareWorkbookPlanVerificationRules(JSON.stringify([rule]))!,
      { read: () => JSON.stringify({ status: "supported", snapshot: { kind: "table_filter_state", state } }) },
    );
    expect(mismatch(emptyCriteriaRule, { columnIndex: 1, state: "filtered", criterion1: "Open", criterion2: "" }).rules[0]?.reason).toBe("filter_state_mismatch");
    expect(mismatch(nonemptyCriteriaRule, { columnIndex: 1, state: "filtered", criterion1: "", criterion2: "Closed" }).rules[0]?.reason).toBe("filter_state_mismatch");

    let authoredReads = 0;
    expect(prepareWorkbookPlanVerificationRules(JSON.stringify([{ ...emptyCriteriaRule, extra: true }]))).toBeNull();
    expect(verifyWorkbookPlan(null as unknown as PreparedWorkbookPlanVerificationRules, { read: () => { authoredReads += 1; return undefined; } }).rules[0]?.reason).toBe("invalid_readback");
    expect(authoredReads).toBe(0);

    const readbackExtra = verifyWorkbookPlan(
      prepared([emptyCriteriaRule]),
      { read: () => JSON.stringify({ status: "supported", snapshot: { kind: "table_filter_state", state: { columnIndex: 1, state: "filtered", criterion1: "", criterion2: "", extra: true } } }) },
    );
    expect(readbackExtra.rules[0]?.reason).toBe("invalid_readback");
  });

  it("rejects malformed or oversized rule JSON before parsing and before reads", () => {
    for (const raw of ["not json", "[]", JSON.stringify([{}]), JSON.stringify(Array.from({ length: MAX_VERIFICATION_RULES + 1 }, () => rules[11])), JSON.stringify([JSON.parse(JSON.stringify({ ...rules[11], extra: "bad" }))])]) {
      let calls = 0;
      const token = prepareWorkbookPlanVerificationRules(raw);
      expect(token).toBeNull();
      verifyWorkbookPlan(token as unknown as PreparedWorkbookPlanVerificationRules, { read: () => { calls += 1; return undefined; } });
      expect(calls).toBe(0);
    }
    const oversized = `[${JSON.stringify({ kind: "range_formula_pattern", sheet: "Data", address: "A1", pattern: "x".repeat(4096) })}${" ".repeat(1_000_000)}]`;
    expect(prepareWorkbookPlanVerificationRules(oversized)).toBeNull();
  });

  it("rejects raw arrays, objects, and transparent token/rule proxies without traps or adapter calls", () => {
    let tokenTraps = 0;
    const tokenProxy = new Proxy({}, { get: () => { tokenTraps += 1; throw new Error("TOKEN_TRAP"); }, ownKeys: () => { tokenTraps += 1; throw new Error("TOKEN_KEYS_TRAP"); } });
    let calls = 0;
    const report = verifyWorkbookPlan(tokenProxy as unknown as PreparedWorkbookPlanVerificationRules, { read: () => { calls += 1; return undefined; } });
    expect(report.ok).toBe(false);
    expect(tokenTraps).toBe(0);
    expect(calls).toBe(0);

    let ruleTraps = 0;
    const rawRuleProxy = new Proxy(rules[4]!, { get: () => { ruleTraps += 1; throw new Error("RULE_TRAP"); }, ownKeys: () => { ruleTraps += 1; throw new Error("RULE_KEYS_TRAP"); } });
    expect(prepareWorkbookPlanVerificationRules(rawRuleProxy as unknown as string)).toBeNull();
    expect(ruleTraps).toBe(0);
    expect(JSON.stringify(report)).not.toMatch(/TOKEN_TRAP|RULE_TRAP/);
  });

  it("rejects malformed, oversized, and extra-field readback JSON exactly", () => {
    const cases = [
      "not json",
      JSON.stringify({ status: "unsupported", reason: "unsupported_host", extra: true }),
      JSON.stringify({ status: "supported", snapshot: { kind: "table_exists", exists: true, extra: "SECRET" } }),
      JSON.stringify({ status: "supported", snapshot: { kind: "table_filter_state", state: { columnIndex: 1, state: "filtered", extra: "SECRET" } } }),
      `${JSON.stringify({ status: "supported", snapshot: { kind: "table_exists", exists: true } })}${" ".repeat(1_000_000)}`,
    ];
    for (const raw of cases) {
      const report = verifyWorkbookPlan(prepared([rules[4]]), { read: () => raw });
      expect(report.ok).toBe(false);
      expect(report.rules[0]?.reason).toBe("invalid_readback");
      expect(JSON.stringify(report)).not.toContain("SECRET");
    }
  });

  it("rejects malformed snapshots and requires finite bounded scalars", () => {
    const cases: readonly [VerificationRule, string][] = [
      [rules[4], JSON.stringify({ kind: "table_exists", exists: "true" })],
      [rules[7], JSON.stringify({ kind: "sheet_exists", exists: 1 })],
      [rules[8], JSON.stringify({ kind: "chart_exists", exists: {} })],
      [rules[9], JSON.stringify({ kind: "comment_exists", exists: true })],
      [rules[11], JSON.stringify({ kind: "calculation_completed", completed: "complete" })],
      [sentinelRule, JSON.stringify({ kind: "sentinel", value: 1e400 })],
    ];
    for (const [rule, snapshot] of cases) {
      const raw = snapshot.includes("null") && rule.kind === "sentinel_unchanged"
        ? JSON.stringify({ status: "supported", snapshot: { kind: "sentinel", value: 1e400 } }).replace("null", "1e400")
        : JSON.stringify({ status: "supported", snapshot });
      const report = verifyWorkbookPlan(prepared([rule]), { read: () => raw });
      expect(report.rules[0]?.reason).toBe("invalid_readback");
    }
  });

  it("accepts exactly 32767 scalar text units and rejects 32768 on both boundaries", () => {
    const accepted = "x".repeat(32767);
    const acceptedRule: VerificationRule = { kind: "sentinel_unchanged", sheet: "Data", address: "A1", expected: accepted };
    expect(verify(acceptedRule, { kind: "sentinel", value: accepted }).ok).toBe(true);
    const rejectedRule = { ...acceptedRule, expected: "x".repeat(32768) };
    expect(prepareWorkbookPlanVerificationRules(JSON.stringify([rejectedRule]))).toBeNull();
    const rejectedReadback = verifyWorkbookPlanVerificationRulesForRaw(prepared([acceptedRule]), JSON.stringify({ status: "supported", snapshot: { kind: "sentinel", value: "x".repeat(32768) } }));
    expect(rejectedReadback.rules[0]?.reason).toBe("invalid_readback");
  });

  it("classifies unsupported, error, missing, thrown, and volatile readbacks", () => {
    expect(verifyWorkbookPlan(prepared([rules[4]]), { read: () => JSON.stringify({ status: "unsupported", reason: "unsupported_capability" }) }).rules[0]?.status).toBe("unsupported");
    expect(verifyWorkbookPlan(prepared([rules[4]]), { read: () => JSON.stringify({ status: "error", reason: "readback_error" }) }).rules[0]?.status).toBe("error");
    expect(verifyWorkbookPlan(prepared([rules[4]]), { read: () => undefined }).rules[0]?.reason).toBe("missing_readback");
    const thrown = verifyWorkbookPlan(prepared([rules[4]]), { read: () => { throw new Error("SECRET_OFFICE_ERROR"); } });
    expect(thrown.rules[0]?.reason).toBe("readback_error");
    expect(JSON.stringify(thrown)).not.toContain("SECRET_OFFICE_ERROR");
    const volatile = verify(valuesRule, { kind: "range_values", values: [[1, "x"], [true, null]], volatile: true });
    expect(volatile.rules[0]?.reason).toBe("volatile_readback");
  });

  it("reports deterministic counts, first failure, cap, and fresh frozen output", () => {
    const calls: string[] = [];
    const report = verifyWorkbookPlan(prepared(rules.slice(0, 4)), {
      read: (rule) => {
        calls.push(rule.kind);
        if (calls.length === 1) return JSON.stringify({ status: "supported", snapshot: { kind: "range_values", values: [[0, 0], [0, 0]] } });
        if (calls.length === 2) return JSON.stringify({ status: "unsupported", reason: "unsupported_host" });
        if (calls.length === 3) return JSON.stringify({ status: "error", reason: "readback_error" });
        return JSON.stringify({ status: "supported", snapshot: snapshots[3] });
      },
    });
    expect(calls).toEqual(rules.slice(0, 4).map((rule) => rule.kind));
    expect(report.counts).toEqual({ verified: 1, mismatched: 1, unsupported: 1, error: 1 });
    expect(report.firstFailure).toEqual({ ruleIndex: 0, kind: "range_values_equal", status: "mismatched", reason: "value_mismatch" });
    const first = verify(sentinelRule, { kind: "sentinel", value: "SECRET_VALUE" });
    const second = verify(sentinelRule, { kind: "sentinel", value: "SECRET_VALUE" });
    expect(first).not.toBe(second);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first.rules)).toBe(true);
    expect(JSON.stringify(first)).not.toContain("SECRET_VALUE");
  });

  it("exposes only one read method and emits no continuation", () => {
    let readCalls = 0;
    const report = verifyWorkbookPlan(prepared([rules[11]]), { read: () => { readCalls += 1; return JSON.stringify({ status: "supported", snapshot: { kind: "calculation_completed", completed: true } }); } });
    expect(readCalls).toBe(1);
    expect(report.ok).toBe(true);
    const failed = verify(rules[4], { kind: "table_exists", exists: false });
    const successful = verify(rules[4], { kind: "table_exists", exists: true });
    const evidence: readonly VerificationInverseEvidence[] = ["not_captured", "incomplete", "complete_not_independently_tested", "complete_and_independently_tested"];
    for (const item of evidence) {
      expect(derivePostVerificationControl(failed, item).modelContinuation).toBe(false);
      expect(derivePostVerificationControl(failed, item).rollback).toBe(item === "complete_and_independently_tested" ? "permitted" : "forbidden");
      expect(derivePostVerificationControl(successful, item).rollback).toBe("forbidden");
    }
  });
});

const verifyWorkbookPlanVerificationRulesForRaw = (token: PreparedWorkbookPlanVerificationRules, raw: string): ReturnType<typeof verifyWorkbookPlan> =>
  verifyWorkbookPlan(token, { read: () => raw });
