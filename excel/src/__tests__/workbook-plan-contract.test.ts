import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import {
  copyFileSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { workbookPlanContract } from "../workbookPlanContract";
import {
  workbookPlanFingerprint,
  workbookPlanSchema,
} from "../generated/workbookPlanSchema";

const compareCodeUnit = (left: string, right: string): number => {
  const limit = Math.min(left.length, right.length);
  for (let index = 0; index < limit; index += 1) {
    const diff = left.charCodeAt(index) - right.charCodeAt(index);
    if (diff !== 0) return diff;
  }
  return left.length - right.length;
};

const canonicalizeDeterministically = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(canonicalizeDeterministically);
  if (value !== null && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value)
        .sort(([left], [right]) => compareCodeUnit(left, right))
        .map(([key, child]) => [key, canonicalizeDeterministically(child)]),
    );
  }
  return value;
};

// Keep the fingerprint assertion aligned with the existing generator. Changing
// that canonicalizer is outside this validator-only patch's authorized scope.
const canonicalizeForFingerprint = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(canonicalizeForFingerprint);
  if (value !== null && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value)
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([key, child]) => [key, canonicalizeForFingerprint(child)]),
    );
  }
  return value;
};

const sha256 = (bytes: Uint8Array): string => createHash("sha256").update(bytes).digest("hex");

describe("WorkbookPlan V1 contract", () => {
  it("defines version 1 and the canonical outcome discriminator", () => {
    expect(workbookPlanContract.schema.properties.version).toMatchObject({
      type: "integer",
      const: 1,
    });
    expect(workbookPlanContract.schema.properties.outcome.enum).toEqual([
      "inspect",
      "execute",
      "ask",
      "refuse",
    ]);
    expect(workbookPlanSchema.properties.version).toMatchObject({
      type: "integer",
      const: 1,
    });
  });

  it("does not expose the feasibility seed's non-canonical filter field", () => {
    expect(workbookPlanContract.schema.$defs.WorkbookOperation.properties).not.toHaveProperty("valuesFilter");
    expect(workbookPlanContract.schema.$defs.WorkbookOperation.properties).toHaveProperty("values");
    expect(workbookPlanSchema.$defs.WorkbookOperation.properties).not.toHaveProperty("valuesFilter");
  });

  it("preserves the required dialect ordering and fingerprints the canonical schema", () => {
    expect(workbookPlanContract.operationKinds).toEqual([
      "write_values",
      "write_formula",
      "fill_formula",
      "format_range",
      "create_table",
      "sort_table",
      "filter_table",
      "reset_table_filter",
      "create_chart",
      "set_validation_list",
      "conditional_format",
      "freeze_panes",
      "create_sheet",
      "rename_sheet",
      "copy_sheet",
      "add_comment",
      "calculate",
    ]);
    expect(workbookPlanContract.verificationKinds).toEqual([
      "range_values_equal",
      "range_formulas_equal",
      "range_formula_pattern",
      "range_number_format_equal",
      "table_exists",
      "table_row_order",
      "table_filter_state",
      "sheet_exists",
      "chart_exists",
      "comment_exists",
      "sentinel_unchanged",
      "calculation_completed",
    ]);

    expect(JSON.stringify(canonicalizeDeterministically(workbookPlanContract.schema))).toBe(
      JSON.stringify(canonicalizeDeterministically(workbookPlanSchema)),
    );

    const digest = createHash("sha256")
      .update(JSON.stringify(canonicalizeForFingerprint(workbookPlanContract.schema)), "utf8")
      .digest("hex");
    expect(workbookPlanFingerprint).toBe(digest);
    expect(workbookPlanFingerprint).toMatch(/^[0-9a-f]{64}$/);
  });

  it("encodes closed outcome branches with the required non-empty fields", () => {
    const branches = workbookPlanSchema.oneOf as readonly Record<string, unknown>[];
    expect(branches).toHaveLength(3);
    expect(branches.map((branch) => branch.required)).toEqual([
      ["version", "outcome", "reason", "reads"],
      ["version", "outcome", "summary", "operations", "verification"],
      ["version", "outcome", "message"],
    ]);
    expect((branches[0].properties as Record<string, unknown>)).not.toHaveProperty("operations");
    expect((branches[0].properties as Record<string, unknown>)).not.toHaveProperty("verification");
    expect((branches[1].properties as Record<string, unknown>)).not.toHaveProperty("reads");
    expect((branches[2].properties as Record<string, unknown>)).not.toHaveProperty("operations");
    expect((branches[2].properties as Record<string, unknown>)).not.toHaveProperty("verification");
    for (const branch of branches) {
      expect(branch.type).toBe("object");
      expect(branch.additionalProperties).toBe(false);
    }
  });

  it("keeps routing, policy, replacement, and destructive fields out of the model schema", () => {
    const modelText = JSON.stringify(workbookPlanSchema);
    for (const forbidden of [
      "bridge",
      "session",
      "policy",
      "timeout",
      "idempotency",
      "request",
      "ticket",
      "transaction",
      "replacement",
      "clear_range",
      "delete_sheet",
      "clear_comments",
      "unprotect_sheet",
      "create_or_replace_table",
    ]) {
      expect(modelText.toLowerCase()).not.toContain(forbidden);
    }
  });

  it("declares complete deterministic metadata for every ordered operation", () => {
    const primitiveTargets = new Set([
      "read_range",
      "summarize_range_stats",
      "write_range",
      "format_range",
      "conditional_format_range",
      "freeze_panes",
      "set_data_validation",
      "create_sheet",
      "insert_table",
      "sort_table",
      "filter_table",
      "create_chart",
      "autofill_formula",
      "rename_sheet",
      "copy_sheet",
      "run_calculation",
      "add_comment",
    ]);
    const costDimensions = [
      "cellsRead",
      "cellsWritten",
      "formulaBytes",
      "commentBytes",
      "sheetsCreated",
      "chartsCreated",
      "tablesCreated",
      "verificationReads",
    ];

    expect(workbookPlanContract.operationDefinitions.map(({ kind }) => kind)).toEqual(
      workbookPlanContract.operationKinds,
    );
    for (const definition of workbookPlanContract.operationDefinitions) {
      expect(definition.effect).toMatch(/^(write|format|structure|filter|calculate)$/);
      expect(definition.compilerTargets.length).toBeGreaterThan(0);
      expect(definition.compilerTargets.every((target) => primitiveTargets.has(target))).toBe(true);
      expect(definition.requiredVerification.length).toBeGreaterThan(0);
      expect(definition.requiredVerification.every((kind) =>
        workbookPlanContract.verificationKinds.includes(kind),
      )).toBe(true);
      expect(definition.capabilityRequirements.some(({ status }) => status === "baseline")).toBe(true);
      expect(definition.capabilityRequirements.some(({ status }) => status === "unverified")).toBe(true);
      expect(Object.keys(definition.resourceCost)).toEqual(costDimensions);
    }

    expect(workbookPlanContract.validationConstraints).toEqual({
      maxErrors: 50,
      maxDepth: 32,
      maxWorkUnits: 4096,
      maxVisitedNodes: 1024,
      maxUtf8Bytes: {
        formulaText: 4096,
        commentText: 1024,
        validationListSource: 4096,
      },
      aliasMap: {
        sheet_name: "sheet",
        range: "address",
        fill_color: "fillColor",
        font_color: "fontColor",
      },
      utf8ByteFields: {
        write_formula: { formula: "formulaText" },
        range_formulas_equal: { expectedFormulas: "formulaText" },
        range_formula_pattern: { pattern: "formulaText" },
        add_comment: { text: "commentText" },
        comment_exists: { text: "commentText" },
        set_validation_list: { source: "validationListSource" },
      },
      rectangularMatrices: {
        write_values: ["values"],
        range_values_equal: ["expectedValues"],
        range_formulas_equal: ["expectedFormulas"],
      },
      conditionalFormatRequires: {
        containsText: ["text"],
        cellValue: ["formula1"],
      },
    });
  });


  it("closes every object schema recursively", () => {
    const assertClosed = (node: unknown): void => {
      if (Array.isArray(node)) {
        for (const item of node) assertClosed(item);
        return;
      }
      if (node === null || typeof node !== "object") return;

      const schemaNode = node as Record<string, unknown>;
      if (schemaNode.type === "object") {
        expect(schemaNode.additionalProperties).toBe(false);
      }
      for (const child of Object.values(schemaNode)) assertClosed(child);
    };

    assertClosed(workbookPlanContract.schema);
    assertClosed(workbookPlanSchema);
  });

  it("recursively resolves every local $ref in the contract and generated schema", () => {
    const collectRefs = (node: unknown, refs: string[] = []): string[] => {
      if (Array.isArray(node)) {
        for (const item of node) collectRefs(item, refs);
        return refs;
      }
      if (node === null || typeof node !== "object") return refs;

      const schemaNode = node as Record<string, unknown>;
      if (typeof schemaNode.$ref === "string") refs.push(schemaNode.$ref);
      for (const child of Object.values(schemaNode)) collectRefs(child, refs);
      return refs;
    };

    const resolveRef = (root: unknown, reference: string): unknown => {
      if (!reference.startsWith("#/")) return undefined;
      let current: unknown = root;
      for (const segment of reference.slice(2).split("/")) {
        if (current === null || typeof current !== "object") return undefined;
        const decoded = segment.replaceAll("~1", "/").replaceAll("~0", "~");
        if (!Object.prototype.hasOwnProperty.call(current, decoded)) return undefined;
        current = (current as Record<string, unknown>)[decoded];
      }
      return current;
    };

    const assertAllRefsResolve = (schema: unknown): void => {
      const references = [...new Set(collectRefs(schema))];
      expect(references.length).toBeGreaterThan(0);
      for (const reference of references) {
        expect(reference.startsWith("#/")).toBe(true);
        expect(resolveRef(schema, reference)).not.toBeUndefined();
      }
    };

    assertAllRefsResolve(workbookPlanContract.schema);
    assertAllRefsResolve(workbookPlanSchema);
  });

  it("rejects a mutated temporary contract without touching real artifacts", () => {
    const repositoryRoot = fileURLToPath(new URL("../..", import.meta.url));
    const generatorPath = join(repositoryRoot, "scripts/generate-workbook-plan-contract.mjs");
    const realContractPath = join(repositoryRoot, "src/workbookPlanContract.ts");
    const realGeneratedPath = join(repositoryRoot, "src/generated/workbookPlanSchema.ts");
    const realContractBefore = readFileSync(realContractPath);
    const realGeneratedBefore = readFileSync(realGeneratedPath);
    const realContractHashBefore = sha256(realContractBefore);
    const realGeneratedHashBefore = sha256(realGeneratedBefore);
    const temporaryRoot = mkdtempSync(join(tmpdir(), "hermes-workbook-plan-"));

    try {
      const temporaryContractPath = join(temporaryRoot, "workbookPlanContract.ts");
      const temporaryGeneratedPath = join(temporaryRoot, "workbookPlanSchema.ts");
      copyFileSync(realContractPath, temporaryContractPath);
      copyFileSync(realGeneratedPath, temporaryGeneratedPath);

      const originalContract = readFileSync(temporaryContractPath, "utf8");
      const marker = "description: \"Static model-facing contract for bounded, non-destructive Excel plans.\"";
      expect(originalContract).toContain(marker);
      writeFileSync(
        temporaryContractPath,
        originalContract.replace(marker, "description: \"Temporary drift mutation.\""),
        "utf8",
      );

      const result = spawnSync(
        process.execPath,
        [
          generatorPath,
          "--check",
          "--contract",
          temporaryContractPath,
          "--generated",
          temporaryGeneratedPath,
        ],
        { cwd: repositoryRoot, encoding: "utf8" },
      );

      expect(result.status).not.toBe(0);
      expect(sha256(readFileSync(realContractPath))).toBe(realContractHashBefore);
      expect(sha256(readFileSync(realGeneratedPath))).toBe(realGeneratedHashBefore);
      expect(readFileSync(realContractPath)).toEqual(realContractBefore);
      expect(readFileSync(realGeneratedPath)).toEqual(realGeneratedBefore);
    } finally {
      rmSync(temporaryRoot, { recursive: true, force: true });
    }
  }, 120_000);

  it("runs the workbook drift check first in the canonical test command", () => {
    const repositoryRoot = fileURLToPath(new URL("../..", import.meta.url));
    const packageJson = JSON.parse(readFileSync(join(repositoryRoot, "package.json"), "utf8")) as {
      scripts: { test: string };
    };

    expect(packageJson.scripts.test.startsWith("npm run check:workbook-plan &&")).toBe(true);
  });

  it("checks the real candidate and writes only deterministic temporary output", () => {
    const repositoryRoot = fileURLToPath(new URL("../..", import.meta.url));
    const generatorPath = join(repositoryRoot, "scripts/generate-workbook-plan-contract.mjs");
    const realContractPath = join(repositoryRoot, "src/workbookPlanContract.ts");
    const realGeneratedPath = join(repositoryRoot, "src/generated/workbookPlanSchema.ts");
    const realContractBefore = readFileSync(realContractPath);
    const realGeneratedBefore = readFileSync(realGeneratedPath);
    const realContractHashBefore = sha256(realContractBefore);
    const realGeneratedHashBefore = sha256(realGeneratedBefore);
    const temporaryRoot = mkdtempSync(join(tmpdir(), "hermes-workbook-plan-"));

    try {
      const realCheck = spawnSync(
        process.execPath,
        [generatorPath, "--check"],
        { cwd: repositoryRoot, encoding: "utf8" },
      );
      expect(realCheck.status).toBe(0);
      expect(realCheck.stderr).toBe("");
      expect(realCheck.stdout).toBe("");

      const temporaryContractPath = join(temporaryRoot, "workbookPlanContract.ts");
      const temporaryGeneratedPath = join(temporaryRoot, "workbookPlanSchema.ts");
      copyFileSync(realContractPath, temporaryContractPath);
      copyFileSync(realGeneratedPath, temporaryGeneratedPath);

      const temporaryWrite = spawnSync(
        process.execPath,
        [
          generatorPath,
          "--contract",
          temporaryContractPath,
          "--generated",
          temporaryGeneratedPath,
        ],
        { cwd: repositoryRoot, encoding: "utf8" },
      );
      expect(temporaryWrite.status).toBe(0);
      expect(temporaryWrite.stderr).toBe("");
      expect(temporaryWrite.stdout).not.toMatch(/\/Users\/|[A-Z]:\\/);
      expect(readFileSync(temporaryGeneratedPath)).toEqual(realGeneratedBefore);

      expect(sha256(readFileSync(realContractPath))).toBe(realContractHashBefore);
      expect(sha256(readFileSync(realGeneratedPath))).toBe(realGeneratedHashBefore);
      expect(readFileSync(realContractPath)).toEqual(realContractBefore);
      expect(readFileSync(realGeneratedPath)).toEqual(realGeneratedBefore);
    } finally {
      rmSync(temporaryRoot, { recursive: true, force: true });
    }
  }, 120_000);
});
