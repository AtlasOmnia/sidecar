// @vitest-environment node
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import * as excelToolsModule from "../excelTools";
import { excelTools, getExcelToolDefinition, toolSchemas } from "../excelTools";
import { workbookPlanContract } from "../workbookPlanContract";
import {
  FULL_FUNCTIONALITY_ACCEPTANCE_CASES,
  FULL_FUNCTIONALITY_TOOL_NAMES,
} from "../taskpane/fullFunctionalityAcceptanceHarness";

/*
 * R4a: every Excel primitive carries exactly one effect class.
 * The new surface is read defensively through the module namespace so that a
 * missing export is a runtime AssertionError (RED), not a type error.
 * EXPECTED is a hand-written literal (D-R4-1 (a), D-R4-2 OMH default); it is
 * deliberately not derived from the module or from the workbook-plan contract.
 */

type LocalEffectClass = "read_only" | "mutating" | "view" | "destructive_legacy";
type LocalPlanEffect = "write" | "format" | "structure" | "filter" | "calculate";
interface LocalPrimitiveEffect {
  readonly effectClass: LocalEffectClass;
  readonly planEffect?: LocalPlanEffect;
  readonly planKinds?: readonly string[];
}

const mod = excelToolsModule as unknown as Record<string, unknown>;

const EXPECTED: Record<string, LocalPrimitiveEffect> = {
  get_workbook_overview: { effectClass: "read_only" },
  read_range: { effectClass: "read_only" },
  summarize_range_stats: { effectClass: "read_only" },
  write_range: { effectClass: "mutating", planEffect: "write", planKinds: ["write_values", "write_formula"] },
  autofill_formula: { effectClass: "mutating", planEffect: "write", planKinds: ["fill_formula"] },
  format_range: { effectClass: "mutating", planEffect: "format", planKinds: ["format_range"] },
  conditional_format_range: { effectClass: "mutating", planEffect: "format", planKinds: ["conditional_format"] },
  freeze_panes: { effectClass: "mutating", planEffect: "format", planKinds: ["freeze_panes"] },
  set_data_validation: { effectClass: "mutating", planEffect: "write", planKinds: ["set_validation_list"] },
  insert_table: { effectClass: "mutating", planEffect: "structure", planKinds: ["create_table"] },
  sort_table: { effectClass: "mutating", planEffect: "write", planKinds: ["sort_table"] },
  filter_table: { effectClass: "mutating", planEffect: "filter", planKinds: ["filter_table", "reset_table_filter"] },
  create_chart: { effectClass: "mutating", planEffect: "structure", planKinds: ["create_chart"] },
  create_sheet: { effectClass: "mutating", planEffect: "structure", planKinds: ["create_sheet"] },
  rename_sheet: { effectClass: "mutating", planEffect: "structure", planKinds: ["rename_sheet"] },
  copy_sheet: { effectClass: "mutating", planEffect: "structure", planKinds: ["copy_sheet"] },
  add_comment: { effectClass: "mutating", planEffect: "write", planKinds: ["add_comment"] },
  run_calculation: { effectClass: "mutating", planEffect: "calculate", planKinds: ["calculate"] },
  activate_sheet: { effectClass: "view" },
  delete_sheet: { effectClass: "destructive_legacy" },
  clear_range: { effectClass: "destructive_legacy" },
  clear_comments: { effectClass: "destructive_legacy" },
  find_replace: { effectClass: "destructive_legacy" },
  unprotect_sheet: { effectClass: "destructive_legacy" },
  protect_sheet: { effectClass: "destructive_legacy" },
  import_csv_to_sheet: { effectClass: "destructive_legacy" },
  create_pivot_table: { effectClass: "destructive_legacy" },
};

const EXPECTED_DESTRUCTIVE_LEGACY = [
  "clear_comments",
  "clear_range",
  "create_pivot_table",
  "delete_sheet",
  "find_replace",
  "import_csv_to_sheet",
  "protect_sheet",
  "unprotect_sheet",
];
const EXPECTED_READ_ONLY = ["get_workbook_overview", "read_range", "summarize_range_stats"];

// Recorded at base 216cb72 (r4a/base-toolschemas-sha.log): 27 schemas, 12144 bytes.
const BASE_TOOL_SCHEMAS_SHA256 = "716c659ca37438fe6b1ab071085cdf7995ed799aa1055540799a62a0d1bd1930";

const PLUGIN_DIR = resolve(__dirname, "../../plugins/atlas-sidecar-excel");
const PLUGIN_PREFIX = "sidecar_excel_";
const PLUGIN_PLAN_TOOL = "sidecar_excel_execute_plan";

const sorted = (values: Iterable<string>): string[] => [...values].sort();
const excelToolNames = (): string[] => excelTools.map((tool) => tool.schema.function.name);

function requireSurface(): {
  table: Record<string, LocalPrimitiveEffect>;
  getEffect: (name: string) => LocalPrimitiveEffect | undefined;
  missing: (names: readonly string[]) => string[];
} {
  expect(mod.PRIMITIVE_EFFECTS, "PRIMITIVE_EFFECTS export").toBeTypeOf("object");
  expect(mod.getPrimitiveEffect, "getPrimitiveEffect export").toBeTypeOf("function");
  expect(mod.missingPrimitiveEffects, "missingPrimitiveEffects export").toBeTypeOf("function");
  return {
    table: mod.PRIMITIVE_EFFECTS as Record<string, LocalPrimitiveEffect>,
    getEffect: mod.getPrimitiveEffect as (name: string) => LocalPrimitiveEffect | undefined,
    missing: mod.missingPrimitiveEffects as (names: readonly string[]) => string[],
  };
}

function namesOfClass(table: Record<string, LocalPrimitiveEffect>, effectClass: LocalEffectClass): string[] {
  return sorted(Object.keys(table).filter((name) => table[name].effectClass === effectClass));
}

function pluginJsonNames(): string[] {
  const parsed = JSON.parse(readFileSync(resolve(PLUGIN_DIR, "tool_schemas.json"), "utf8")) as Array<{ name: string }>;
  return parsed.map((entry) => entry.name);
}

function pluginYamlNames(): string[] {
  const text = readFileSync(resolve(PLUGIN_DIR, "plugin.yaml"), "utf8");
  const lines = text.split(/\r?\n/);
  const start = lines.findIndex((line) => line.trim() === "provides_tools:");
  expect(start, "plugin.yaml provides_tools key").toBeGreaterThanOrEqual(0);
  const names: string[] = [];
  for (const line of lines.slice(start + 1)) {
    const match = /^\s+-\s+(\S+)\s*$/.exec(line);
    if (!match) break;
    names.push(match[1]);
  }
  return names;
}

const stripPlugin = (names: readonly string[]): string[] =>
  sorted(names.filter((name) => name !== PLUGIN_PLAN_TOOL).map((name) => name.slice(PLUGIN_PREFIX.length)));

describe("R4a primitive effect table", () => {
  it("E1 completeness: every excelTools primitive has exactly one entry", () => {
    const { table, missing } = requireSurface();
    const names = excelToolNames();
    expect(names).toHaveLength(27);
    expect(Object.keys(table)).toHaveLength(27);
    expect(sorted(Object.keys(table))).toEqual(sorted(names));
    expect(missing(names)).toEqual([]);
    expect(missing(["write_range", "new_tool"])).toEqual(["new_tool"]);
    expect(missing(["zz_b", "read_range", "zz_a", "__proto__"])).toEqual(["zz_b", "zz_a", "__proto__"]);
  });

  it("E2 exact table matches D-R4-2 / D-R4-1 and class counts 3/15/1/8", () => {
    const { table } = requireSurface();
    const plain = JSON.parse(JSON.stringify(table)) as Record<string, LocalPrimitiveEffect>;
    expect(plain).toEqual(EXPECTED);
    expect(namesOfClass(table, "read_only")).toHaveLength(3);
    expect(namesOfClass(table, "mutating")).toHaveLength(15);
    expect(namesOfClass(table, "view")).toHaveLength(1);
    expect(namesOfClass(table, "destructive_legacy")).toHaveLength(8);
    for (const [name, entry] of Object.entries(table)) {
      const keys = sorted(Object.keys(entry));
      if (entry.effectClass === "mutating") {
        expect(keys, name).toEqual(["effectClass", "planEffect", "planKinds"]);
        expect(entry.planKinds!.length, name).toBeGreaterThan(0);
        expect(new Set(entry.planKinds).size, name).toBe(entry.planKinds!.length);
      } else {
        expect(keys, name).toEqual(["effectClass"]);
      }
    }
  });

  it("E3 table agrees with the workbook-plan contract in both directions", () => {
    const { table } = requireSurface();
    const definitions = workbookPlanContract.operationDefinitions;
    const compilerTargets = new Set<string>();
    for (const definition of definitions) {
      for (const target of definition.compilerTargets) {
        compilerTargets.add(target);
        const entry = table[target];
        expect(entry, `${definition.kind} -> ${target}`).toBeDefined();
        expect(entry.effectClass, target).toBe("mutating");
        expect(entry.planKinds, target).toContain(definition.kind);
        expect(entry.planEffect, target).toBe(definition.effect);
      }
    }
    for (const [name, entry] of Object.entries(table)) {
      for (const kind of entry.planKinds ?? []) {
        const definition = definitions.find((candidate) => candidate.kind === kind);
        expect(definition, `${name} planKind ${kind}`).toBeDefined();
        expect(definition!.compilerTargets as readonly string[], `${kind} targets`).toContain(name);
      }
    }
    expect(namesOfClass(table, "mutating")).toEqual(sorted(compilerTargets));
    expect(compilerTargets.size).toBe(15);
    const allPlanKinds = Object.values(table).flatMap((entry) => entry.planKinds ?? []);
    expect(allPlanKinds).toHaveLength(17);
    expect(sorted(allPlanKinds)).toEqual(sorted(workbookPlanContract.operationKinds));
  });

  it("E4a plugin oracle: tool_schemas.json and plugin.yaml mirror excelTools plus excel_execute_plan", () => {
    const jsonNames = pluginJsonNames();
    const yamlNames = pluginYamlNames();
    expect(jsonNames).toHaveLength(28);
    expect(yamlNames).toHaveLength(28);
    expect(jsonNames).toContain(PLUGIN_PLAN_TOOL);
    expect(yamlNames).toContain(PLUGIN_PLAN_TOOL);
    for (const name of [...jsonNames, ...yamlNames]) expect(name.startsWith(PLUGIN_PREFIX), name).toBe(true);
    expect(stripPlugin(jsonNames)).toEqual(sorted(excelToolNames()));
    expect(stripPlugin(yamlNames)).toEqual(sorted(excelToolNames()));
  });

  it("E4 plugin parity: plugin tool names equal the table keys", () => {
    const { table } = requireSurface();
    expect(stripPlugin(pluginJsonNames())).toEqual(sorted(Object.keys(table)));
    expect(stripPlugin(pluginYamlNames())).toEqual(sorted(Object.keys(table)));
  });

  it("E5 inspection kinds are read_only; read_only minus inspection kinds is get_workbook_overview", () => {
    const { table } = requireSurface();
    const inspection = workbookPlanContract.inspectionKinds as readonly string[];
    for (const kind of inspection) expect(table[kind]?.effectClass, kind).toBe("read_only");
    expect(namesOfClass(table, "read_only").filter((name) => !inspection.includes(name))).toEqual([
      "get_workbook_overview",
    ]);
  });

  it("E6 acceptance-harness mutates flags agree with the table", () => {
    const { table } = requireSurface();
    expect(FULL_FUNCTIONALITY_ACCEPTANCE_CASES.length).toBeGreaterThan(0);
    for (const testCase of FULL_FUNCTIONALITY_ACCEPTANCE_CASES) {
      const entry = table[testCase.tool];
      expect(entry, testCase.id).toBeDefined();
      expect(testCase.mutates, testCase.id).toBe(entry.effectClass !== "read_only");
    }
    expect(sorted(new Set(FULL_FUNCTIONALITY_TOOL_NAMES))).toEqual(sorted(Object.keys(table)));
  });

  it("E7 owner-decision sets: destructive_legacy, view and run_calculation", () => {
    const { table } = requireSurface();
    expect(namesOfClass(table, "destructive_legacy")).toEqual(EXPECTED_DESTRUCTIVE_LEGACY);
    expect(namesOfClass(table, "view")).toEqual(["activate_sheet"]);
    expect(namesOfClass(table, "read_only")).toEqual(EXPECTED_READ_ONLY);
    expect(table.run_calculation?.effectClass).toBe("mutating");
    expect(table.run_calculation?.planEffect).toBe("calculate");
    expect(table.sort_table?.effectClass).toBe("mutating");
    expect(table.filter_table?.effectClass).toBe("mutating");
  });

  it("E8 table is deeply frozen and lookup is own-property only", () => {
    const { table, getEffect } = requireSurface();
    expect(Object.isFrozen(table)).toBe(true);
    for (const [name, entry] of Object.entries(table)) {
      expect(Object.isFrozen(entry), name).toBe(true);
      if (entry.planKinds) expect(Object.isFrozen(entry.planKinds), `${name}.planKinds`).toBe(true);
    }
    const mutableTable = table as Record<string, unknown>;
    expect(() => {
      mutableTable.zz_probe = { effectClass: "read_only" };
    }).toThrow(TypeError);
    expect(() => {
      (table.write_range as { effectClass: string }).effectClass = "read_only";
    }).toThrow(TypeError);
    expect(() => {
      (table.write_range.planKinds as string[]).push("zz_probe");
    }).toThrow(TypeError);
    expect(table.write_range.effectClass).toBe("mutating");
    expect("zz_probe" in table).toBe(false);

    for (const name of Object.keys(EXPECTED)) expect(getEffect(name), name).toBe(table[name]);
    for (const name of ["__proto__", "constructor", "toString", "hasOwnProperty", "excel_write_range", "", "WRITE_RANGE"]) {
      expect(getEffect(name), JSON.stringify(name)).toBeUndefined();
    }
    for (const value of [undefined, null, 42, {}, ["write_range"]]) {
      expect(getEffect(value as unknown as string), String(value)).toBeUndefined();
    }
  });

  it("E9 model-visible tool schemas are unchanged", () => {
    expect(toolSchemas).toHaveLength(27);
    for (const schema of toolSchemas) expect(sorted(Object.keys(schema))).toEqual(["function", "type"]);
    const text = JSON.stringify(toolSchemas);
    for (const marker of ["effectClass", "planKinds", "destructive_legacy", "PRIMITIVE"]) {
      expect(text.includes(marker), marker).toBe(false);
    }
    expect(createHash("sha256").update(text, "utf8").digest("hex")).toBe(BASE_TOOL_SCHEMAS_SHA256);
  });

  it("E10 harness oracle alone: tools with any mutates:false case are exactly the read-only three", () => {
    const nonMutating = new Set(
      FULL_FUNCTIONALITY_ACCEPTANCE_CASES.filter((testCase) => !testCase.mutates).map((testCase) => testCase.tool as string),
    );
    expect(sorted(nonMutating)).toEqual(EXPECTED_READ_ONLY);
    for (const testCase of FULL_FUNCTIONALITY_ACCEPTANCE_CASES) {
      expect(testCase.mutates, testCase.id).toBe(!EXPECTED_READ_ONLY.includes(testCase.tool));
    }
  });

  it("E11 every EXPECTED key is a registered primitive; plugin-prefixed names are not", () => {
    expect(Object.keys(EXPECTED)).toHaveLength(27);
    for (const name of Object.keys(EXPECTED)) expect(getExcelToolDefinition(name), name).toBeDefined();
    expect(getExcelToolDefinition("excel_write_range")).toBeUndefined();
  });
});
