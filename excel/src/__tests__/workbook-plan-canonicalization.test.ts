import { describe, expect, it } from "vitest";
import { validateWorkbookPlan } from "../workbookPlanValidation";
import { canonicalizeWorkbookPlan } from "../workbookPlanCanonicalization";

async function withPrototypeCanary<T>(
  key: string,
  value: unknown,
  run: () => Promise<T>,
): Promise<T> {
  const previous = Object.getOwnPropertyDescriptor(Object.prototype, key);
  Object.defineProperty(Object.prototype, key, {
    value,
    enumerable: false,
    configurable: true,
    writable: true,
  });

  try {
    return await run();
  } finally {
    if (previous === undefined) {
      Reflect.deleteProperty(Object.prototype, key);
    } else {
      Object.defineProperty(Object.prototype, key, previous);
    }
  }
}

describe("WorkbookPlan canonicalization", () => {
  it("normalizes alias-only raw input before validation", async () => {
    const raw = {
      version: 1,
      outcome: "execute" as const,
      summary: "alias normalization",
      operations: [
        {
          kind: "write_values" as const,
          sheet_name: "Sheet1",
          range: "A1",
          values: [["x"]],
        },
      ],
      verification: [
        {
          kind: "range_values_equal" as const,
          sheet: "Sheet1",
          address: "A1",
          expectedValues: [["x"]],
        },
      ],
    };

    expect(validateWorkbookPlan(raw).ok).toBe(false);

    const result = await canonicalizeWorkbookPlan(raw);

    expect(result.stage).toBe("success");
    if (result.stage !== "success") return;

    const plan = result.value;
    if (plan.outcome !== "execute") {
      throw new Error("Expected execute plan");
    }

    expect(plan).toEqual({
      version: 1,
      outcome: "execute",
      summary: "alias normalization",
      operations: [
        {
          kind: "write_values",
          sheet: "Sheet1",
          address: "A1",
          values: [["x"]],
        },
      ],
      verification: [
        {
          kind: "range_values_equal",
          sheet: "Sheet1",
          address: "A1",
          expectedValues: [["x"]],
        },
      ],
    });
    expect(plan).not.toBe(raw);
    expect(plan.operations).not.toBe(raw.operations);
    const operation = plan.operations[0];
    if (operation.kind !== "write_values") {
      throw new Error("Expected write_values operation");
    }
    expect(operation).not.toBe(raw.operations[0]);
    expect(operation.values).not.toBe(raw.operations[0].values);
    expect(plan.verification).not.toBe(raw.verification);
    expect(Object.getPrototypeOf(plan)).toBe(Object.prototype);
    expect(Object.getPrototypeOf(plan.operations)).toBe(Array.prototype);
    expect(Object.getPrototypeOf(operation)).toBe(Object.prototype);
    expect(result.semanticDigest).toMatch(/^[0-9a-f]{64}$/);
    expect(raw).toEqual({
      version: 1,
      outcome: "execute",
      summary: "alias normalization",
      operations: [
        {
          kind: "write_values",
          sheet_name: "Sheet1",
          range: "A1",
          values: [["x"]],
        },
      ],
      verification: [
        {
          kind: "range_values_equal",
          sheet: "Sheet1",
          address: "A1",
          expectedValues: [["x"]],
        },
      ],
    });
  });

  it("normalizes fill_color and font_color aliases inside nested format objects", async () => {
    const raw = {
      version: 1,
      outcome: "execute" as const,
      summary: "format alias normalization",
      operations: [
        {
          kind: "format_range" as const,
          sheet: "Sheet1",
          address: "A1",
          format: {
            bold: true,
            fill_color: "FF0000",
            font_color: "0000FF",
          },
        },
      ],
      verification: [
        {
          kind: "sentinel_unchanged" as const,
          sheet: "Sheet1",
          address: "A1",
          expected: "ok",
        },
      ],
    };

    const result = await canonicalizeWorkbookPlan(raw);

    expect(result.stage).toBe("success");
    if (result.stage !== "success") return;

    const plan = result.value;
    if (plan.outcome !== "execute") {
      throw new Error("Expected execute plan");
    }
    const operation = plan.operations[0];
    if (operation.kind !== "format_range") {
      throw new Error("Expected format_range operation");
    }

    expect(operation.format).toEqual({
      bold: true,
      fillColor: "FF0000",
      fontColor: "0000FF",
    });
    expect(operation.format).not.toBe(raw.operations[0].format);
    expect(Object.getPrototypeOf(operation.format)).toBe(Object.prototype);
    expect(result.semanticDigest).toMatch(/^[0-9a-f]{64}$/);
    expect(raw.operations[0].format).toEqual({
      bold: true,
      fill_color: "FF0000",
      font_color: "0000FF",
    });
  });

  it("rejects same-object alias and canonical properties without leaking values", async () => {
    const raw = {
      version: 1,
      outcome: "inspect" as const,
      reason: "alias conflict",
      reads: [
        {
          kind: "read_range" as const,
          sheet: "Sheet1",
          sheet_name: "CANARY_ALIAS",
          address: "A1",
        },
      ],
    };

    const result = await canonicalizeWorkbookPlan(raw);

    expect(result.stage).toBe("syntax");
    if (result.stage !== "syntax") return;

    expect(result.errors[0]).toEqual(
      expect.objectContaining({
        code: "alias_conflict",
        path: "/reads/0/sheet_name",
      }),
    );
    expect(JSON.stringify(result.errors)).not.toContain("CANARY_ALIAS");
  });

  it("rejects polluted prototype aliases that would otherwise rewrite unknown reads into sheet", async () => {
    const canaryKey = "__hermes_alias_to_sheet_canary__";
    await withPrototypeCanary(canaryKey, "sheet", async () => {
      const result = await canonicalizeWorkbookPlan({
        version: 1,
        outcome: "inspect" as const,
        reason: "polluted alias to sheet",
        reads: [
          {
            kind: "read_range" as const,
            address: "A1",
            [canaryKey]: "CANARY_SHEET",
          },
        ],
      });

      expect(result.stage).toBe("validation");
      if (result.stage !== "validation") return;

      expect(result.errors.some((error) => error.code === "unknown_property")).toBe(true);
      expect(result.errors.some((error) => error.code === "missing_required")).toBe(true);
      expect(JSON.stringify(result.errors)).not.toContain("CANARY_SHEET");
    });
  });

  it("rejects polluted prototype aliases that would otherwise disappear into __proto__", async () => {
    const canaryKey = "__hermes_alias_to_proto_canary__";
    await withPrototypeCanary(canaryKey, "__proto__", async () => {
      const result = await canonicalizeWorkbookPlan({
        version: 1,
        outcome: "inspect" as const,
        reason: "polluted alias to proto",
        reads: [
          {
            kind: "read_range" as const,
            sheet: "Sheet1",
            address: "A1",
            [canaryKey]: "CANARY_PROTO",
          },
        ],
      });

      expect(result.stage).toBe("validation");
      if (result.stage !== "validation") return;

      expect(result.errors.some((error) => error.code === "unknown_property")).toBe(true);
      expect(result.errors.some((error) => error.path.includes(canaryKey))).toBe(true);
      expect(JSON.stringify(result.errors)).not.toContain("CANARY_PROTO");
    });
  });

  it("preserves JSON-authored __proto__ data keys and keeps canonical sheet conflicts unknown", async () => {
    const jsonProtoPlan = JSON.parse(
      '{"version":1,"outcome":"inspect","reason":"json proto","reads":[{"kind":"read_range","sheet":"Sheet1","address":"A1","__proto__":{"marker":"CANARY_JSON_PROTO"}}]}',
    );
    const jsonProtoRead = jsonProtoPlan.reads[0] as Record<string, unknown>;

    expect(Object.prototype.hasOwnProperty.call(jsonProtoRead, "__proto__")).toBe(true);
    expect(Object.getPrototypeOf(jsonProtoRead)).toBe(Object.prototype);

    const jsonProtoDescriptor = Object.getOwnPropertyDescriptor(jsonProtoRead, "__proto__");
    expect(jsonProtoDescriptor).toEqual(
      expect.objectContaining({
        enumerable: true,
        configurable: true,
        writable: true,
      }),
    );
    expect(jsonProtoDescriptor?.value).toEqual({ marker: "CANARY_JSON_PROTO" });

    const jsonProtoResult = await canonicalizeWorkbookPlan(jsonProtoPlan);
    expect(jsonProtoResult.stage).toBe("validation");
    if (jsonProtoResult.stage !== "validation") return;

    expect(
      jsonProtoResult.errors.some(
        (error) => error.code === "unknown_property" && error.path === "/reads/0/__proto__",
      ),
    ).toBe(true);
    expect(JSON.stringify(jsonProtoResult.errors)).not.toContain("CANARY_JSON_PROTO");

    const canaryKey = "__hermes_alias_overwrite_sheet_canary__";
    await withPrototypeCanary(canaryKey, "sheet", async () => {
      const conflictResult = await canonicalizeWorkbookPlan({
        version: 1,
        outcome: "inspect" as const,
        reason: "canonical plus polluted",
        reads: [
          {
            kind: "read_range" as const,
            sheet: "Sheet1",
            address: "A1",
            [canaryKey]: "CANARY_DUPLICATE",
          },
        ],
      });

      expect(conflictResult.stage).toBe("validation");
      if (conflictResult.stage !== "validation") return;

      expect(
        conflictResult.errors.some(
          (error) => error.code === "unknown_property" && error.path === `/reads/0/${canaryKey}`,
        ),
      ).toBe(true);
      expect(JSON.stringify(conflictResult.errors)).not.toContain("CANARY_DUPLICATE");
    });
  });

  it("produces the same digest for equivalent property order and JSON whitespace", async () => {
    const first = JSON.parse(
      '{"outcome":"execute","version":1,"summary":"order stable","verification":[{"kind":"sentinel_unchanged","expected":"ok","address":"A1","sheet":"Sheet1"}],"operations":[{"sheet":"Sheet1","kind":"format_range","address":"A1","format":{"fontColor":"FF0000","bold":true,"fillColor":"00FF00"}}]}',
    );
    const second = JSON.parse(
      '{\n        "version": 1,\n        "summary": "order stable",\n        "outcome": "execute",\n        "operations": [\n          {\n            "format": {\n              "fillColor": "00FF00",\n              "bold": true,\n              "fontColor": "FF0000"\n            },\n            "address": "A1",\n            "sheet": "Sheet1",\n            "kind": "format_range"\n          }\n        ],\n        "verification": [\n          {\n            "sheet": "Sheet1",\n            "expected": "ok",\n            "address": "A1",\n            "kind": "sentinel_unchanged"\n          }\n        ]\n      }',
    );

    const firstResult = await canonicalizeWorkbookPlan(first);
    const secondResult = await canonicalizeWorkbookPlan(second);

    expect(firstResult.stage).toBe("success");
    expect(secondResult.stage).toBe("success");
    if (firstResult.stage !== "success" || secondResult.stage !== "success") return;

    expect(firstResult.semanticDigest).toBe(secondResult.semanticDigest);
  });

  it("produces different digests for distinct addresses, ranges, and formulas", async () => {
    const unqualified = await canonicalizeWorkbookPlan({
      version: 1,
      outcome: "execute" as const,
      summary: "address digest",
      operations: [
        {
          kind: "write_values" as const,
          sheet: "Sheet1",
          address: "A1",
          values: [[1]],
        },
      ],
      verification: [
        {
          kind: "range_values_equal" as const,
          sheet: "Sheet1",
          address: "A1",
          expectedValues: [[1]],
        },
      ],
    });
    const qualified = await canonicalizeWorkbookPlan({
      version: 1,
      outcome: "execute" as const,
      summary: "address digest",
      operations: [
        {
          kind: "write_values" as const,
          sheet: "Sheet1",
          address: "Sheet1!A1",
          values: [[1]],
        },
      ],
      verification: [
        {
          kind: "range_values_equal" as const,
          sheet: "Sheet1",
          address: "Sheet1!A1",
          expectedValues: [[1]],
        },
      ],
    });
    const expandedRange = await canonicalizeWorkbookPlan({
      version: 1,
      outcome: "execute" as const,
      summary: "range digest",
      operations: [
        {
          kind: "write_values" as const,
          sheet: "Sheet1",
          address: "A1:A2",
          values: [[1], [2]],
        },
      ],
      verification: [
        {
          kind: "range_values_equal" as const,
          sheet: "Sheet1",
          address: "A1:A2",
          expectedValues: [[1], [2]],
        },
      ],
    });
    const separateTargets = await canonicalizeWorkbookPlan({
      version: 1,
      outcome: "execute" as const,
      summary: "range digest",
      operations: [
        {
          kind: "write_values" as const,
          sheet: "Sheet1",
          address: "A1",
          values: [[1]],
        },
        {
          kind: "write_values" as const,
          sheet: "Sheet1",
          address: "A2",
          values: [[2]],
        },
      ],
      verification: [
        {
          kind: "range_values_equal" as const,
          sheet: "Sheet1",
          address: "A1",
          expectedValues: [[1]],
        },
      ],
    });
    const formulaOne = await canonicalizeWorkbookPlan({
      version: 1,
      outcome: "execute" as const,
      summary: "formula digest",
      operations: [
        {
          kind: "write_formula" as const,
          sheet: "Sheet1",
          address: "B1",
          formula: "=SUM(A1:A2)",
        },
      ],
      verification: [
        {
          kind: "range_formulas_equal" as const,
          sheet: "Sheet1",
          address: "B1",
          expectedFormulas: [["=SUM(A1:A2)"]],
        },
      ],
    });
    const formulaTwo = await canonicalizeWorkbookPlan({
      version: 1,
      outcome: "execute" as const,
      summary: "formula digest",
      operations: [
        {
          kind: "write_formula" as const,
          sheet: "Sheet1",
          address: "B1",
          formula: "=SUM(A1:A3)",
        },
      ],
      verification: [
        {
          kind: "range_formulas_equal" as const,
          sheet: "Sheet1",
          address: "B1",
          expectedFormulas: [["=SUM(A1:A3)"]],
        },
      ],
    });

    for (const result of [unqualified, qualified, expandedRange, separateTargets, formulaOne, formulaTwo]) {
      expect(result.stage).toBe("success");
      if (result.stage !== "success") return;
      expect(result.semanticDigest).toHaveLength(64);
    }

    if (
      unqualified.stage !== "success" ||
      qualified.stage !== "success" ||
      expandedRange.stage !== "success" ||
      separateTargets.stage !== "success" ||
      formulaOne.stage !== "success" ||
      formulaTwo.stage !== "success"
    ) {
      return;
    }

    expect(unqualified.semanticDigest).not.toBe(qualified.semanticDigest);
    expect(expandedRange.semanticDigest).not.toBe(separateTargets.semanticDigest);
    expect(formulaOne.semanticDigest).not.toBe(formulaTwo.semanticDigest);
  });

  it("rejects duplicate mutating operations but allows duplicate calculate and read steps", async () => {
    const duplicateWriteValues = await canonicalizeWorkbookPlan({
      version: 1,
      outcome: "execute" as const,
      summary: "duplicate write",
      operations: [
        {
          kind: "write_values" as const,
          sheet: "Sheet1",
          address: "A1",
          values: [["CANARY_WRITE"]],
        },
        {
          kind: "write_values" as const,
          sheet: "Sheet1",
          address: "A1",
          values: [["CANARY_WRITE"]],
        },
      ],
      verification: [
        {
          kind: "range_values_equal" as const,
          sheet: "Sheet1",
          address: "A1",
          expectedValues: [["CANARY_WRITE"]],
        },
      ],
    });
    const duplicateFormatRange = await canonicalizeWorkbookPlan({
      version: 1,
      outcome: "execute" as const,
      summary: "duplicate format",
      operations: [
        {
          kind: "format_range" as const,
          sheet: "Sheet1",
          address: "A1",
          format: { bold: true, fillColor: "CANARY_FORMAT" },
        },
        {
          kind: "format_range" as const,
          sheet: "Sheet1",
          address: "A1",
          format: { bold: true, fillColor: "CANARY_FORMAT" },
        },
      ],
      verification: [
        {
          kind: "sentinel_unchanged" as const,
          sheet: "Sheet1",
          address: "A1",
          expected: "ok",
        },
      ],
    });
    const duplicateCalculate = await canonicalizeWorkbookPlan({
      version: 1,
      outcome: "execute" as const,
      summary: "duplicate calculate",
      operations: [{ kind: "calculate" as const }, { kind: "calculate" as const }],
      verification: [{ kind: "calculation_completed" as const }],
    });
    const duplicateReads = await canonicalizeWorkbookPlan({
      version: 1,
      outcome: "inspect" as const,
      reason: "duplicate reads",
      reads: [
        { kind: "read_range" as const, sheet: "Sheet1", address: "A1" },
        { kind: "read_range" as const, sheet: "Sheet1", address: "A1" },
      ],
    });

    expect(duplicateWriteValues.stage).toBe("syntax");
    if (duplicateWriteValues.stage !== "syntax") return;
    expect(duplicateWriteValues.errors[0]).toEqual(
      expect.objectContaining({
        code: "duplicate_write",
        path: "/operations/1",
        firstPath: "/operations/0",
      }),
    );
    expect(JSON.stringify(duplicateWriteValues.errors)).not.toContain("CANARY_WRITE");

    expect(duplicateFormatRange.stage).toBe("syntax");
    if (duplicateFormatRange.stage !== "syntax") return;
    expect(duplicateFormatRange.errors[0]).toEqual(
      expect.objectContaining({
        code: "duplicate_write",
        path: "/operations/1",
        firstPath: "/operations/0",
      }),
    );
    expect(JSON.stringify(duplicateFormatRange.errors)).not.toContain("CANARY_FORMAT");

    expect(duplicateCalculate.stage).toBe("success");
    expect(duplicateReads.stage).toBe("success");
    if (duplicateCalculate.stage !== "success" || duplicateReads.stage !== "success") return;
    expect(duplicateCalculate.operationDigests).toHaveLength(2);
    expect(duplicateCalculate.operationDigests[0]).toBe(duplicateCalculate.operationDigests[1]);
    expect(duplicateReads.readDigests).toHaveLength(2);
    expect(duplicateReads.readDigests[0]).toBe(duplicateReads.readDigests[1]);
  });

  it("preserves validation failures for unknown properties without healing them", async () => {
    const result = await canonicalizeWorkbookPlan({
      version: 1,
      outcome: "execute" as const,
      summary: "unknown property",
      operations: [
        {
          kind: "write_values" as const,
          sheet: "Sheet1",
          address: "A1",
          values: [["CANARY_UNKNOWN"]],
          unknownField: "CANARY_UNKNOWN",
        },
      ],
      verification: [
        {
          kind: "range_values_equal" as const,
          sheet: "Sheet1",
          address: "A1",
          expectedValues: [["CANARY_UNKNOWN"]],
        },
      ],
    });

    expect(result.stage).toBe("validation");
    if (result.stage !== "validation") return;

    expect(result.errors.some((error) => error.code === "unknown_property")).toBe(true);
    expect(JSON.stringify(result.errors)).not.toContain("CANARY_UNKNOWN");
  });

  it("preserves hidden non-enumerable unknown properties as validation failures", async () => {
    const plan = {
      version: 1,
      outcome: "execute" as const,
      summary: "hidden unknown property",
      operations: [
        {
          kind: "write_values" as const,
          sheet: "Sheet1",
          address: "A1",
          values: [["CANARY_HIDDEN"]],
        },
      ],
      verification: [
        {
          kind: "range_values_equal" as const,
          sheet: "Sheet1",
          address: "A1",
          expectedValues: [["CANARY_HIDDEN"]],
        },
      ],
    };
    Object.defineProperty(plan.operations[0] as Record<string, unknown>, "unknownField", {
      value: "CANARY_HIDDEN",
      enumerable: false,
      configurable: true,
      writable: true,
    });

    const result = await canonicalizeWorkbookPlan(plan);

    expect(result.stage).toBe("validation");
    if (result.stage !== "validation") return;

    expect(result.errors.some((error) => error.code === "unknown_property")).toBe(true);
    expect(JSON.stringify(result.errors)).not.toContain("CANARY_HIDDEN");
  });

  it("rejects symbol keys, accessors, non-plain objects, and sparse arrays without invoking getters", async () => {
    const symbolPlan = {
      version: 1,
      outcome: "inspect" as const,
      reason: "symbol key",
      reads: [{ kind: "read_range" as const, sheet: "Sheet1", address: "A1" }],
    };
    Reflect.set(symbolPlan, Symbol("secret"), "CANARY_SYMBOL");
    const symbolResult = await canonicalizeWorkbookPlan(symbolPlan);
    expect(symbolResult.stage).toBe("syntax");
    if (symbolResult.stage !== "syntax") return;
    expect(symbolResult.errors[0]).toEqual(
      expect.objectContaining({ code: "symbol_key_not_supported" }),
    );
    expect(JSON.stringify(symbolResult.errors)).not.toContain("CANARY_SYMBOL");

    const accessorPlan = {
      version: 1,
      outcome: "execute" as const,
      summary: "accessor",
      operations: [
        {
          kind: "write_values" as const,
          values: [["x"]],
          address: "A1",
          sheet: "Sheet1",
        },
      ],
      verification: [
        {
          kind: "range_values_equal" as const,
          sheet: "Sheet1",
          address: "A1",
          expectedValues: [["x"]],
        },
      ],
    };
    let getterCalls = 0;
    Object.defineProperty(accessorPlan.operations[0] as Record<string, unknown>, "sheet", {
      enumerable: true,
      configurable: true,
      get() {
        getterCalls += 1;
        return "Sheet1";
      },
    });
    const accessorResult = await canonicalizeWorkbookPlan(accessorPlan);
    expect(getterCalls).toBe(0);
    expect(accessorResult.stage).toBe("syntax");
    if (accessorResult.stage !== "syntax") return;
    expect(accessorResult.errors[0]).toEqual(
      expect.objectContaining({ code: "accessor_property", path: "/operations/0/sheet" }),
    );

    const nonPlainPlan = {
      version: 1,
      outcome: "execute" as const,
      summary: "non plain",
      operations: [
        {
          kind: "write_values" as const,
          sheet: "Sheet1",
          address: "A1",
          values: [["x"]],
        },
      ],
      verification: [
        {
          kind: "range_values_equal" as const,
          sheet: "Sheet1",
          address: "A1",
          expectedValues: [["x"]],
        },
      ],
      extra: new Date("2024-01-01T00:00:00.000Z"),
    };
    const nonPlainResult = await canonicalizeWorkbookPlan(nonPlainPlan);
    expect(nonPlainResult.stage).toBe("syntax");
    if (nonPlainResult.stage !== "syntax") return;
    expect(nonPlainResult.errors[0]).toEqual(
      expect.objectContaining({ code: "non_plain_object" }),
    );

    const sparsePlan = {
      version: 1,
      outcome: "execute" as const,
      summary: "sparse",
      operations: [
        {
          kind: "write_values" as const,
          sheet: "Sheet1",
          address: "A1",
          values: [["x"]],
        },
      ],
      verification: [
        {
          kind: "range_values_equal" as const,
          sheet: "Sheet1",
          address: "A1",
          expectedValues: [["x"]],
        },
      ],
      extra: ["x", , "z"],
    };
    const sparseResult = await canonicalizeWorkbookPlan(sparsePlan);
    expect(sparseResult.stage).toBe("syntax");
    if (sparseResult.stage !== "syntax") return;
    expect(sparseResult.errors[0]).toEqual(
      expect.objectContaining({ code: "sparse_array" }),
    );
  });

  it("fails content-free when clone work, depth, or visited-node limits are exceeded", async () => {
    const deepRoot: Record<string, unknown> = {};
    let deepCursor = deepRoot;
    for (let index = 0; index < 40; index += 1) {
      const next: Record<string, unknown> = {};
      deepCursor.nested = next;
      deepCursor = next;
    }

    const depthResult = await canonicalizeWorkbookPlan({
      version: 1,
      outcome: "execute" as const,
      summary: "depth exhaustion",
      operations: [
        {
          kind: "write_values" as const,
          sheet: "Sheet1",
          address: "A1",
          values: [["x"]],
        },
      ],
      verification: [
        {
          kind: "range_values_equal" as const,
          sheet: "Sheet1",
          address: "A1",
          expectedValues: [["x"]],
        },
      ],
      extra: deepRoot,
    });

    const wideResult = await canonicalizeWorkbookPlan({
      version: 1,
      outcome: "execute" as const,
      summary: "work exhaustion",
      operations: [
        {
          kind: "write_values" as const,
          sheet: "Sheet1",
          address: "A1",
          values: [["x"]],
        },
      ],
      verification: [
        {
          kind: "range_values_equal" as const,
          sheet: "Sheet1",
          address: "A1",
          expectedValues: [["x"]],
        },
      ],
      extra: Array.from({ length: 5000 }, (_, index) => index),
    });

    expect(depthResult.stage).toBe("syntax");
    if (depthResult.stage !== "syntax") return;
    expect(depthResult.errors[0]).toEqual(
      expect.objectContaining({ code: "work_limit_exceeded" }),
    );

    expect(wideResult.stage).toBe("syntax");
    if (wideResult.stage !== "syntax") return;
    expect(wideResult.errors[0]).toEqual(
      expect.objectContaining({ code: "work_limit_exceeded" }),
    );
  });
});
