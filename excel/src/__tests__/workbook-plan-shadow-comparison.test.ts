import { describe, expect, it, vi } from "vitest";
import { cpSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import * as authority from "../primitiveMutationAuthority";
import * as tools from "../excelTools";
import * as compiler from "../workbookPlanCompiler";
import * as validation from "../workbookPlanValidation";
import * as policy from "../workbookPlanPolicy";
import * as request from "../workbookPlanRequest";
import * as route from "../workbookPlanRoute";
import * as protocol from "../workbookPlanProtocol";
import * as contract from "../workbookPlanContract";
import * as office from "../workbookPlanOfficeAdapter";
import * as executor from "../workbookPlanExecutor";
const kernel = { ...authority, ...tools, ...compiler, ...validation, ...policy, ...request, ...route, ...protocol, ...contract, ...office, ...executor };
const root = resolve(__dirname, "../..");
const script = join(root, "scripts/run-workbook-plan-conformance.mjs");
const corpus = join(__dirname, "fixtures/workbook-plan-conformance/shadow-v1");
const temp = () => mkdtempSync(join(process.env.TMPDIR!, "r6c-test-"));
async function core() { const url = pathToFileURL(script).href; return await import(/* @vite-ignore */ url); }
async function run(overrides: Record<string, unknown> = {}) { return (await core()).runWorkbookPlanShadowComparison({ kernel, corpus, ...overrides }); }
describe("R6c offline primary/shadow comparison", () => {
  it("S1/S5 executes primary primitives and independently reconciles eight outcomes", async () => {
    const result = await run();
    expect(result.records).toHaveLength(8);
    expect(result.summary).toMatchObject({ total: 8, pass: 8, fail: 0, error: 0 });
    const byId = (id: string) => result.records.find((r: any) => r.caseId === id);
    expect(byId("equal")).toMatchObject({ primaryOutcome: "committed", shadowOutcome: "compiled", targetSetEqual: true, effectClassesEqual: true, primaryOpCount: 1, shadowOpCount: 1 });
    expect(byId("target_mismatch")).toMatchObject({ targetSetEqual: false, effectClassesEqual: true });
    expect(byId("effect_mismatch")).toMatchObject({ targetSetEqual: true, effectClassesEqual: false });
    expect(byId("count_mismatch")).toMatchObject({ primaryOpCount: 2, shadowOpCount: 1 });
    expect(byId("policy_reject")).toMatchObject({ shadowStage: "policy", shadowOutcome: "requires_confirmation" });
    expect(byId("compile_reject")).toMatchObject({ shadowStage: "compile", shadowOutcome: "rejected" });
    expect(byId("primitive_denied")).toMatchObject({ primaryOutcome: "denied", primaryOpCount: 0 });
    expect(byId("malformed")).toMatchObject({ shadowStage: "protocol", shadowOutcome: "rejected" });
  });
  it("S2 shadow never calls Office, native adapter or executor, including a swallowed trap", async () => {
    const adapter = vi.fn(() => { throw new Error("adapter forbidden"); });
    const execute = vi.fn(() => { throw new Error("executor forbidden"); });
    const good = await run({ kernel: { ...kernel, createNativeWorkbookPlanAdapter: adapter, executeAcceptedWorkbookPlan: execute } });
    expect(good.records.every((r: any) => r.shadowExecuteCalls === 0)).toBe(true);
    expect(adapter).not.toHaveBeenCalled(); expect(execute).not.toHaveBeenCalled();
    const mutant = await run({ kernel: { ...kernel, decideWorkbookPlanPolicy: (input: unknown) => {
      try { (globalThis as any).Excel.run(() => undefined); } catch { /* counts even if swallowed */ }
      return policy.decideWorkbookPlanPolicy(input);
    } } });
    expect(mutant.records.filter((r: any) => r.shadowExecuteCalls > 0).every((r: any) => r.status === "error")).toBe(true);
    expect(mutant.summary.error).toBeGreaterThan(0);
  });
  it("S3 metadata-only output restores global state, contains no canary/targets, and is deterministic", async () => {
    const before = Object.getOwnPropertyDescriptor(globalThis, "Excel");
    const out = temp(); const fetch = vi.fn(() => { throw new Error("network forbidden"); });
    vi.stubGlobal("fetch", fetch);
    try {
      const a = await run({ out }); const b = await run(); expect(a).toEqual(b);
      expect(Object.getOwnPropertyDescriptor(globalThis, "Excel")).toEqual(before);
      const text = readFileSync(join(out, "shadow-records.jsonl"), "utf8");
      expect(text).not.toMatch(/SYNTHETIC_SHADOW_CANARY|Sheet1|A1|values|modelOutput|userPrompt/);
      expect(fetch).not.toHaveBeenCalled();
    } finally { vi.unstubAllGlobals(); }
  });
  it("S4/S6 CLI exercises compiled modules; baseline R6b records remain byte-identical", async () => {
    const out = temp();
    const cli = spawnSync(process.execPath, [script, "--shadow", "--out", out, "--candidate", "6b2eaf011a1bdb1949374747c65f82b016b61a5d"], { cwd: root, encoding: "utf8", timeout: 120000 });
    expect(cli.status, cli.stderr).toBe(0);
    const rows = readFileSync(join(out, "shadow-records.jsonl"), "utf8").trim().split("\n").map(line => JSON.parse(line));
    expect(rows.at(-1)).toMatchObject({ total: 8, pass: 8, fail: 0, error: 0 });
    const base = temp();
    const baseline = spawnSync(process.execPath, [script, "--out", base, "--candidate", "6b2eaf011a1bdb1949374747c65f82b016b61a5d"], { cwd: root, encoding: "utf8", timeout: 120000 });
    expect(baseline.status, baseline.stderr).toBe(0);
    expect(createHash("sha256").update(readFileSync(join(base, "records.jsonl"))).digest("hex")).toBe("e72a16b638c312caad3c99c0e9e492ae0178a9e4dfa6bd31ed348eb20791a9d0");
  }, 120_000);
  it("S6 CLI fails on independently observed mismatch and strict host error", () => {
    for (const mode of ["fail", "error"] as const) {
      const input = join(temp(), "corpus"); cpSync(corpus, input, { recursive: true });
      const file = join(input, "equal.json"); const fixture = JSON.parse(readFileSync(file, "utf8"));
      if (mode === "fail") fixture.primary[0].args.address = "B2";
      else fixture.workbookSeed.sheets = ["Other"];
      writeFileSync(file, JSON.stringify(fixture) + "\n");
      const manifestPath = join(input, "manifest.json"); const m = JSON.parse(readFileSync(manifestPath, "utf8"));
      m.files["equal.json"] = createHash("sha256").update(readFileSync(file)).digest("hex");
      m.corpusSha256 = createHash("sha256").update(Object.entries(m.files).map(([k, v]) => `${k}:${v}`).join("")).digest("hex");
      writeFileSync(manifestPath, JSON.stringify(m));
      const out = temp();
      const cli = spawnSync(process.execPath, [script, "--shadow", "--shadow-corpus", input, "--out", out], { cwd: root, encoding: "utf8", timeout: 120000 });
      expect(cli.status, cli.stderr).toBe(1);
      const rows = readFileSync(join(out, "shadow-records.jsonl"), "utf8").trim().split("\n").map(line => JSON.parse(line));
      expect(rows.find(r => r.caseId === "equal").status).toBe(mode);
      expect(rows.at(-1)[mode]).toBe(1);
    }
  }, 120_000);
});
