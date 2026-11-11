import { describe, expect, it, vi } from "vitest";
import { existsSync, mkdtempSync, readFileSync, readdirSync, cpSync, appendFileSync, symlinkSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { spawnSync } from "node:child_process";
import * as orchestrator from "../workbookPlanOrchestrator";
import * as office from "../workbookPlanOfficeAdapter";
import * as executor from "../workbookPlanExecutor";
import * as request from "../workbookPlanRequest";
import * as route from "../workbookPlanRoute";
import * as protocol from "../workbookPlanProtocol";
import * as resources from "../workbookPlanResources";
import * as schema from "../generated/workbookPlanSchema";
import * as serverRoute from "../taskpane/serverOwnedPlanRoute";
import * as coordinator from "../taskpane/taskPaneTurnCoordinator";
const root = resolve(__dirname, "../..");
const script = join(root, "scripts/run-workbook-plan-conformance.mjs");
const original = join(__dirname, "fixtures/model-conformance");
const synthetic = join(__dirname, "fixtures/workbook-plan-conformance/v1");
const kernel = { ...orchestrator, ...office, ...executor, ...request, ...route, ...protocol, ...resources, ...schema, ...serverRoute, ...coordinator };
const table: Record<string, [string, string, number]> = {
  write: ["committed", "verified", 1], format: ["committed", "verified", 1], table: ["committed", "verified", 2],
  inspect_execute: ["committed", "verified", 1], ask: ["ask", "not_attempted", 0], refuse: ["refuse", "not_attempted", 0],
  confirmed: ["committed", "verified", 1], declined: ["rejected", "not_attempted", 0], expired: ["rejected", "not_attempted", 0],
  resource_limit: ["rejected", "not_attempted", 0], stale: ["partial", "not_attempted", 0],
  mismatch: ["verification_failed", "failed", 1], partial: ["partial", "not_attempted", 1],
  malformed: ["rejected", "not_attempted", 0], unknown_tool: ["rejected", "not_attempted", 0],
  duplicate: ["rejected", "not_attempted", 0], extra_field: ["rejected", "not_attempted", 0],
  spanish: ["committed", "verified", 1], xml: ["committed", "verified", 1],
};
async function core() {
  expect(existsSync(script), "C1 evaluator core exists").toBe(true);
  const specifier = pathToFileURL(script).href;
  return await import(/* @vite-ignore */ specifier);
}
async function run(extra: Record<string, unknown> = {}) {
  return (await core()).runWorkbookPlanConformance({ kernel, corpora: [original, synthetic], clock: () => 1000, ...extra });
}
const scratch = () => mkdtempSync(join(process.env.TMPDIR!, "r6b-test-"));
describe("offline workbook plan conformance", () => {
  it("C1/C2/C4/C5/C10 independently reconciles all cases and outcomes", async () => {
    const result = await run();
    expect(result.summary.total).toBe(32); // 12 frozen + 20 v1 (19 direct + 1 server)
    expect(result.summary.notPlanCandidate).toBe(9);
    expect(result.summary.total).toBe(result.summary.pass + result.summary.fail + result.summary.unavailable + result.summary.error + result.summary.contractDrift);
    expect(result.summary.fail).toBe(0);
    expect(result.summary.contractDrift).toBe(2);
    expect(result.records.filter((r: any) => r.status === "contractDrift").map((r: any) => r.caseId)).toEqual(["openai-tool-call-001", "xml-fallback-001"]);
    expect(result.records.find((r: any) => r.caseId === "server_route")).toMatchObject({ route: "server-owned", outcome: "committed", verification: "verified", mutationCount: 1, status: "pass" });
    expect(result.summary.error).toBe(0);
    for (const [caseId, [outcome, verification, mutationCount]] of Object.entries(table)) {
      expect(result.records.find((r: any) => r.caseId === caseId)).toMatchObject({ outcome, verification, mutationCount, status: "pass" });
    }
    expect(result.records.find((r: any) => r.caseId === "partial")).toMatchObject({ rollbackStatus: "unavailable", opCount: 2 });
    expect(result.records.filter((r: any) => r.route === "primitive_not_plan")).toHaveLength(9);
  });
  it("C3 rejects unsynchronized reads and unknown APIs; mutations wait for sync", async () => {
    const { createStrictExcelDouble } = await core();
    const d = createStrictExcelDouble({ sheets: ["Sheet1"], cells: {} });
    await d.Excel.run(async (c: any) => {
      const r = c.workbook.worksheets.getItem("Sheet1").getRange("A1");
      expect(() => r.values).toThrow();
      r.load("values"); expect(() => r.values).toThrow();
      await c.sync(); expect(r.values).toEqual([[null]]);
      r.values = [["PRIVATE_CANARY_R6B"]]; expect(d.mutationCount).toBe(0);
      await c.sync(); expect(d.mutationCount).toBe(1);
      expect(() => r.unsupported()).toThrow();
    });
    const bad = await run({ kernel: { ...kernel, readNativeWorkbookPlanRuntimeContext: async () => {
      return (globalThis as any).Excel.run(async (c: any) => c.workbook.worksheets.getItem("Sheet1").getRange("A1").values);
    } } });
    expect(bad.records.find((r: any) => r.caseId === "write").status).not.toBe("pass");
  });
  it("C6/C9 privacy, no network, restored global, and determinism", async () => {
    const fetch = vi.fn(() => { throw new Error("network forbidden"); });
    vi.stubGlobal("fetch", fetch);
    const before = (globalThis as any).Excel;
    try {
      const out = scratch(); const a = await run({ out }); const b = await run();
      expect(a).toEqual(b); expect(fetch).not.toHaveBeenCalled();
      expect((globalThis as any).Excel).toBe(before);
      const text = readFileSync(join(out, "records.jsonl"), "utf8");
      expect(text).not.toContain("PRIVATE_CANARY_R6B");
      expect(text).not.toMatch(/userPrompt|values|formulas|finalSummary|modelOutput|Sheet1|A1/);
      const keys = ["evaluatorVersion", "candidateSha", "schemaFingerprint", "corpusId", "corpusVersion", "corpusSha256", "caseId", "transport", "locale", "modelId", "modelRevision", "budgets", "context", "route", "stageReached", "outcome", "verification", "rollbackStatus", "mutationCount", "opCount", "status", "reasonCode", "protocolAccepted"];
      for (const record of a.records) expect(Object.keys(record).sort()).toEqual(keys.sort());
    } finally { vi.unstubAllGlobals(); }
  });
  it("C7 CLI compiles the real kernel; refuses repo and symlink output", async () => {
    await core();
    const before = readdirSync(root).filter(n => n.startsWith("r6b-conformance-build-"));
    const out = scratch();
    const cli = spawnSync(process.execPath, [script, "--out", out, "--candidate", "0123456789abcdef0123456789abcdef01234567"], { cwd: root, encoding: "utf8", timeout: 120000 });
    expect(cli.status, cli.stderr).toBe(0);
    const lines = readFileSync(join(out, "records.jsonl"), "utf8").trim().split("\n").map(x => JSON.parse(x));
    expect(lines).toHaveLength(33); expect(lines.at(-1).total).toBe(32);
    expect(readdirSync(root).filter(n => n.startsWith("r6b-conformance-build-"))).toEqual(before);
    const denied = spawnSync(process.execPath, [script, "--out", root], { encoding: "utf8" });
    expect(denied.status).toBe(2);
    const alias = join(scratch(), "alias"); symlinkSync(root, alias, "dir");
    expect(spawnSync(process.execPath, [script, "--out", join(alias, "new-output")], { encoding: "utf8" }).status).toBe(2);
    const fileAlias = scratch(); symlinkSync(join(root, "package.json"), join(fileAlias, "records.jsonl"));
    expect(spawnSync(process.execPath, [script, "--out", fileAlias], { encoding: "utf8" }).status).toBe(2);
  }, 150000);
  it("R1 pins only actual frozen mismatches; unpinned drift fails and a matching pin errors", async () => {
    const evaluator = await core();
    const a = evaluator.loadCorpus(original); const b = evaluator.loadCorpus(synthetic);
    const unpinned = await run({ corpora: [a, { ...b, manifest: { ...b.manifest, contractDrift: [] } }] });
    expect(unpinned.summary.fail).toBe(2);
    expect(unpinned.summary.contractDrift).toBe(0);
    const matching = structuredClone(b);
    matching.manifest.contractDrift = [{ caseId: "partial-multi-op-001", reasonCode: "legacy_plan_shape_never_accepted", provenanceCommit: "9d3afdded0ed0a75beef22a736f5818df97c1daf" }];
    const pinnedPass = await run({ corpora: [a, matching] });
    expect(pinnedPass.records.find((r: any) => r.caseId === "partial-multi-op-001")).toMatchObject({ status: "error", reasonCode: "contract_drift_pin_matches_label" });
    expect(pinnedPass.summary.error).toBe(1);
    const out = scratch(); const corpus = join(out, "corpus"); cpSync(synthetic, corpus, { recursive: true });
    // Test manifest edits do not alter the frozen corpus or any fixture bytes.
    const { writeFileSync } = await import("node:fs");
    writeFileSync(join(corpus, "manifest.json"), JSON.stringify({ ...b.manifest, contractDrift: [] }));
    const cli = spawnSync(process.execPath, [script, "--out", join(out, "results"), "--corpus", corpus], { encoding: "utf8", timeout: 120000 });
    expect(cli.status).toBe(1);
  }, 150000);
  it("R1 server coverage invokes the exported route and inspection seam", async () => {
    const server = (await core()).loadCorpus(synthetic);
    const fixture = server.fixtures.find((f: any) => f.caseId === "server_route");
    const routeFactory = vi.fn((kernel as any).createServerOwnedPlanRoute);
    const result = await run({ kernel: { ...kernel, createServerOwnedPlanRoute: routeFactory }, corpora: [{ ...server, manifest: { ...server.manifest, contractDrift: [] }, fixtures: [fixture] }] });
    expect(routeFactory).toHaveBeenCalledTimes(1);
    expect(result.records[0]).toMatchObject({ status: "pass", stageReached: "verification", outcome: "committed", mutationCount: 1 });
  });
  it("C8 refuses a corrupt original before any records are written", async () => {
    await core(); const temp = scratch(); const copy = join(temp, "original"); cpSync(original, copy, { recursive: true });
    appendFileSync(join(copy, "openai-tool-call-001.json"), " ");
    const out = join(temp, "out");
    const cli = spawnSync(process.execPath, [script, "--out", out, "--original-corpus", copy], { encoding: "utf8" });
    expect(cli.status).toBe(3); expect(existsSync(join(out, "records.jsonl"))).toBe(false);
  }, 120_000);
});

describe("R6b-r2 independent review repairs", () => {
  it("R2 frozen labels compare protocol acceptance, even when policy rejects downstream", async () => {
    const evaluator = await core();
    const a = evaluator.loadCorpus(original); const b = evaluator.loadCorpus(synthetic);
    const write = b.fixtures.find((f: any) => f.caseId === "write");
    const fixture = { ...a.fixtures.find((f: any) => f.caseId === "partial-multi-op-001"), workbookSeed: write.workbookSeed };
    const output = write.modelOutput;
    const normalized = await protocol.normalizeWorkbookPlanResponse(output.content ?? "", JSON.stringify(output.tool_calls ?? output.toolCalls ?? []));
    if (normalized.ok !== true || normalized.kind !== "plan") throw new Error("invalid synthetic test plan");
    const downstreamReject = vi.fn(async (deps: any) => {
      await deps.getCompileContextForPlan(normalized.plan);
      deps.policyInput(normalized.plan);
      return { outcome: "rejected", verification: "not_attempted" };
    });
    const injected = { ...kernel, normalizeWorkbookPlanResponse: async () => normalized, createWorkbookPlanTurn: downstreamReject };
    const mismatch = await run({ kernel: injected, corpora: [{ ...a, fixtures: [fixture] }] });
    expect(mismatch.records[0]).toMatchObject({ protocolAccepted: true, outcome: "rejected", stageReached: "policy_compile", status: "fail" });
    const accepted = { ...fixture, expectedDisposition: "accepted" };
    const match = await run({ kernel: injected, corpora: [{ ...a, fixtures: [accepted] }] });
    expect(match.records[0]).toMatchObject({ protocolAccepted: true, outcome: "rejected", status: "pass" });
    expect(downstreamReject).toHaveBeenCalledTimes(2);
  });

  it("R3 CLI cannot substitute a named synthetic corpus for the frozen original role", async () => {
    const temp = scratch(); const copy = join(temp, "named-corpus"); cpSync(synthetic, copy, { recursive: true });
    const manifest = JSON.parse(readFileSync(join(copy, "manifest.json"), "utf8"));
    writeFileSync(join(copy, "manifest.json"), JSON.stringify({ ...manifest, contractDrift: [] }));
    const out = join(temp, "results");
    const cli = spawnSync(process.execPath, [script, "--out", out, "--original-corpus", copy, "--corpus", copy], { encoding: "utf8", timeout: 120000 });
    expect(cli.status).toBe(3);
    expect(existsSync(join(out, "records.jsonl"))).toBe(false);
  }, 150000);

  it("R4 scratch and its aliases must resolve outside the repository before any build", async () => {
    const evaluator = await core();
    // This fails before invoking a dangerous CLI at the unguarded RED base.
    expect(typeof evaluator.validateScratchDirectory).toBe("function");
    expect(() => evaluator.validateScratchDirectory(root)).toThrow("repo_scratch_forbidden");
    const alias = join(scratch(), "scratch-alias"); symlinkSync(root, alias, "dir");
    expect(() => evaluator.validateScratchDirectory(alias)).toThrow("repo_scratch_forbidden");
    const before = readdirSync(root).filter(n => n.startsWith("r6b-conformance-build-"));
    for (const bad of [root, alias]) {
      const out = scratch();
      const cli = spawnSync(process.execPath, [script, "--out", out], { encoding: "utf8", env: { ...process.env, TMPDIR: bad } });
      expect(cli.status).toBe(2);
      expect(existsSync(join(out, "records.jsonl"))).toBe(false);
    }
    expect(readdirSync(root).filter(n => n.startsWith("r6b-conformance-build-"))).toEqual(before);
  }, 120_000);

  it("R5 manifest metadata is bounded before loading or serializing any record", async () => {
    const evaluator = await core(); const a = evaluator.loadCorpus(original); const b = evaluator.loadCorpus(synthetic);
    const metadata = { userPrompt: "PRIVATE_CANARY_R6B" };
    const temp = scratch(); const copy = join(temp, "corpus"); cpSync(synthetic, copy, { recursive: true });
    writeFileSync(join(copy, "manifest.json"), JSON.stringify({ ...b.manifest, corpusVersion: metadata }));
    expect(() => evaluator.loadCorpus(copy)).toThrow("invalid_corpus_metadata");
    const out = join(temp, "results");
    await expect(run({ out, corpora: [a, { ...b, manifest: { ...b.manifest, corpusVersion: metadata } }] })).rejects.toThrow("invalid_corpus_metadata");
    expect(existsSync(join(out, "records.jsonl"))).toBe(false);
  });

  it("R6 expired context reads, retained methods, mutations and sync all fail", async () => {
    const d = (await core()).createStrictExcelDouble(); let context: any; let range: any; let load: any;
    await d.Excel.run(async (c: any) => {
      context = c; range = c.workbook.worksheets.getItem("Sheet1").getRange("A1");
      range.load("values"); await c.sync(); load = range.load;
    });
    expect(() => range.values).toThrow("ContextExpired");
    expect(() => load("values")).toThrow("ContextExpired");
    expect(() => { range.values = [[1]]; }).toThrow("ContextExpired");
    expect(() => context.workbook.getSelectedRange()).toThrow("ContextExpired");
    await expect(async () => context.sync()).rejects.toThrow("ContextExpired");
    expect(d.mutationCount).toBe(0);
  });

  it("R7 a range from another still-active context cannot be used to create a table", async () => {
    const d = (await core()).createStrictExcelDouble();
    await d.Excel.run(async (a: any) => {
      const foreign = a.workbook.worksheets.getItem("Sheet1").getRange("A1");
      foreign.load("address"); await a.sync();
      await d.Excel.run(async (b: any) => {
        expect(() => b.workbook.worksheets.getItem("Sheet1").tables.add(foreign, true)).toThrow("InvalidRequestContext");
        await b.sync();
      });
    });
    expect(d.mutationCount).toBe(0);
  });

  it("R8 swallowed expired-proxy misuse still invalidates native adapter evidence", async () => {
    const server = (await core()).loadCorpus(synthetic);
    const fixture = server.fixtures.find((f: any) => f.caseId === "write");
    const result = await run({ corpora: [{ ...server, manifest: { ...server.manifest, contractDrift: [] }, fixtures: [fixture] }], kernel: {
      ...kernel, readNativeWorkbookPlanRuntimeContext: async (plan: any) => {
        let expired: any;
        await (globalThis as any).Excel.run(async (c: any) => { expired = c.workbook.worksheets.getItem("Sheet1").getRange("A1"); });
        try { expired.values = [["PRIVATE_CANARY_R6B"]]; }
        catch { /* Deliberately swallowed adapter error; the evaluator must detect its counter. */ }
        return office.readNativeWorkbookPlanRuntimeContext(plan);
      },
    } });
    expect(result.records[0]).toMatchObject({ status: "error", reasonCode: "strict_office_api_error" });
  });
  it("R9 the context, workbook and sheet table collection cannot silently hide unknown APIs", async () => {
    const d = (await core()).createStrictExcelDouble();
    await d.Excel.run(async (c: any) => {
      expect(() => c.unsupported).toThrow("ApiNotFound");
      expect(() => c.workbook.unsupported).toThrow("ApiNotFound");
      expect(() => c.workbook.worksheets.getItem("Sheet1").tables.unsupported).toThrow("ApiNotFound");
    });
    expect(d.apiErrors).toBe(3);
  });

  it("R10 JSON-escaped secrets are caught after parsing", async () => {
    const evaluator = await core();
    const temp = scratch();
    // \u0073k- decodes to "sk-" only after JSON.parse; a raw-text scan misses it.
    const escapedFixture = { caseId: "escaped-secret", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { content: "{\"tool_calls\":[]}", tool_calls: [] } };
    const fixtureText = JSON.stringify(escapedFixture).replace('"sourceClass":"generic"', '"sourceClass":"\\u0073k-AAAAAAAAAAAAAAAAAAAA"');
    const { writeFileSync } = await import("node:fs");
    writeFileSync(join(temp, "escaped-secret.json"), fixtureText);
    const { createHash } = await import("node:crypto");
    const digest = createHash("sha256").update(fixtureText).digest("hex");
    const manifest = { version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "escaped-secret.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`escaped-secret.json:${digest}`).digest("hex") };
    writeFileSync(join(temp, "manifest.json"), JSON.stringify(manifest));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow("corpus_sensitive_pattern");
  });

  it("R11 write-then-load reads the just-written value in one sync", async () => {
    const d = (await core()).createStrictExcelDouble();
    await d.Excel.run(async (c: any) => {
      const r = c.workbook.worksheets.getItem("Sheet1").getRange("A1");
      r.values = [[1]]; r.load("values"); await c.sync();
      expect(r.values).toEqual([[1]]);
      r.load("values"); r.values = [[2]]; await c.sync();
      expect(r.values).toEqual([[2]]);
    });
    expect(d.mutationCount).toBe(2);
  });

  it("R12 a sensitive property name is rejected", async () => {
    const evaluator = await core();
    const temp = scratch();
    const fixture = { caseId: "sensitive-key", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [], token: "opaque-secret" } };
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "sensitive-key.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "sensitive-key.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`sensitive-key.json:${digest}`).digest("hex") }));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow("corpus_sensitive_pattern");
  });

  it("R13 escaped secrets inside a nested transport payload are rejected", async () => {
    const evaluator = await core();
    const temp = scratch();
    const fixture = { caseId: "nested-escape", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [{ id: "s", type: "function", function: { name: "excel_execute_plan", arguments: "{\"summary\":\"\\u0073k-AAAAAAAAAAAAAAAAAAAA\"}" } }] } };
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "nested-escape.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "nested-escape.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`nested-escape.json:${digest}`).digest("hex") }));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow("corpus_sensitive_pattern");
  });

  it("R14 a range written through another proxy still requires its own reload", async () => {
    const d = (await core()).createStrictExcelDouble();
    await d.Excel.run(async (c: any) => {
      const a = c.workbook.worksheets.getItem("Sheet1").getRange("A1");
      a.load("values"); await c.sync();
      expect(a.values).toEqual([[null]]);
      const b = c.workbook.worksheets.getItem("Sheet1").getRange("A1");
      b.values = [[5]]; await c.sync();
      expect(a.values).toEqual([[null]]);
      a.load("values"); await c.sync();
      expect(a.values).toEqual([[5]]);
    });
    expect(d.mutationCount).toBe(1);
  });

  it("R15 collection descendant loads resolve during the same sync", async () => {
    const d = (await core()).createStrictExcelDouble();
    await d.Excel.run(async (c: any) => {
      const ws = c.workbook.worksheets;
      ws.load("items/name"); await c.sync();
      expect(ws.items.map((s: any) => s.name)).toEqual(["Sheet1"]);
    });
  });

  it("R16 escaped secrets inside an XML transport payload are rejected", async () => {
    const evaluator = await core();
    const temp = scratch();
    const fixture = { caseId: "xml-escape", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [{ id: "x", type: "function", function: { name: "excel_execute_plan", arguments: "<tool_call>{\"summary\":\"\\u0073k-AAAAAAAAAAAAAAAAAAAA\"}</tool_call>" } }] } };
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "xml-escape.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "xml-escape.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`xml-escape.json:${digest}`).digest("hex") }));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow("corpus_sensitive_pattern");
  });

  it("R17 a non-SHA candidate identity is rejected before records are written", async () => {
    const evaluator = await core();
    await expect(evaluator.runWorkbookPlanConformance({ kernel, corpora: [original, synthetic], candidateSha: "sk-AAAAAAAAAAAAAAAAAAAA" })).rejects.toThrow("invalid_candidate_sha");
    await expect(evaluator.runWorkbookPlanConformance({ kernel, corpora: [original, synthetic], candidateSha: "not-a-sha" })).rejects.toThrow("invalid_candidate_sha");
  });

  it("R18 a non-range object passed as a table range is rejected", async () => {
    const d = (await core()).createStrictExcelDouble();
    await d.Excel.run(async (c: any) => {
      const ws = c.workbook.worksheets.getItem("Sheet1");
      expect(() => ws.tables.add(ws, true)).toThrow("ApiNotFound");
    });
    expect(d.mutationCount).toBe(0);
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R19 an expired confirmation rejects a valid confirm candidate with zero mutations", async () => {
    const result = await run();
    const expired = result.records.find((r: any) => r.caseId === "expired");
    const confirmed = result.records.find((r: any) => r.caseId === "confirmed");
    expect(expired).toMatchObject({ outcome: "rejected", verification: "not_attempted", mutationCount: 0, status: "pass" });
    expect(confirmed).toMatchObject({ outcome: "committed", mutationCount: 1, status: "pass" });
  });

  it("R20 malformed matrix rows are rejected before mutation", async () => {
    const d = (await core()).createStrictExcelDouble();
    await expect(d.Excel.run(async (c: any) => {
      const r = c.workbook.worksheets.getItem("Sheet1").getRange("A1");
      r.values = [{ 0: "x", length: 1 }];
      await c.sync();
    })).rejects.toThrow("double_shape_mismatch");
    expect(d.mutationCount).toBe(0);
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R21 sensitive property names inside XML-wrapped JSON are rejected", async () => {
    const evaluator = await core();
    const temp = scratch();
    const fixture = { caseId: "xml-key", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [{ id: "x", type: "function", function: { name: "excel_execute_plan", arguments: "<tool_call>{\"password\":\"opaque-secret\"}</tool_call>" } }] } };
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "xml-key.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "xml-key.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`xml-key.json:${digest}`).digest("hex") }));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow("corpus_sensitive_pattern");
  });

  it("R22 forward-slash aliases cannot bypass repo output protection", async () => {
    const evaluator = await core();
    const alias = join(scratch(), "alias"); symlinkSync(root, alias, "dir");
    const fwd = alias.replaceAll("\\", "/") + "/new-output";
    expect(() => evaluator.validateOutputDirectory(fwd)).toThrow("repo_output_forbidden");
  });

  it("R23 in-memory corpora cannot produce evidence records", async () => {
    const evaluator = await core();
    const a = evaluator.loadCorpus(original); const b = evaluator.loadCorpus(synthetic);
    const out = scratch();
    await expect(run({ out, corpora: [a, b] })).rejects.toThrow("inmemory_corpus_not_evidence");
    expect(existsSync(join(out, "records.jsonl"))).toBe(false);
    const bad = { ...b, fixtures: b.fixtures.map((f: any, i: number) => i === 0 ? { ...f, caseId: "sk-AAAAAAAAAAAAAAAAAAAA" } : f) };
    await expect(run({ corpora: [a, bad] })).rejects.toThrow("corpus_sensitive_pattern");
  });

  it("R24 multiple XML-wrapped JSON envelopes cannot hide a sensitive key", async () => {
    const evaluator = await core();
    const temp = scratch();
    const fixture = { caseId: "xml-multi", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [{ id: "x", type: "function", function: { name: "excel_execute_plan", arguments: "<tool_call>{\"password\":\"opaque-secret\"}</tool_call><tool_call>{}</tool_call>" } }] } };
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "xml-multi.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "xml-multi.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`xml-multi.json:${digest}`).digest("hex") }));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow("corpus_sensitive_pattern");
  });

  it("R25 missing-sheet errors invalidate evidence", async () => {
    const d = (await core()).createStrictExcelDouble();
    await expect(d.Excel.run(async (c: any) => {
      c.workbook.worksheets.getItem("Nope");
    })).rejects.toThrow("missing_sheet");
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R26 multiple XML envelopes with a hyphenated sensitive key are rejected", async () => {
    const evaluator = await core();
    const temp = scratch();
    const fixture = { caseId: "xml-multi-token", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [{ id: "x", type: "function", function: { name: "excel_execute_plan", arguments: "<tool_call>{\"access-token\":\"opaque-value\"}</tool_call><tool_call>{}</tool_call>" } }] } };
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "xml-multi-token.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "xml-multi-token.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`xml-multi-token.json:${digest}`).digest("hex") }));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow("corpus_sensitive_pattern");
  });

  it("R27 a swallowed missing-table lookup still invalidates evidence", async () => {
    const d = (await core()).createStrictExcelDouble();
    await d.Excel.run(async (c: any) => {
      try { c.workbook.tables.getItem("Nope"); } catch {}
    });
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R28 missing-object ranges cannot fabricate usable writes", async () => {
    const d = (await core()).createStrictExcelDouble();
    await expect(d.Excel.run(async (c: any) => {
      const nullSheet = c.workbook.worksheets.getItemOrNullObject("Nope");
      nullSheet.getRange("A1").values = [[1]];
      await c.sync();
    })).rejects.toThrow("ItemNotFound");
    expect(d.mutationCount).toBe(0);
    expect(d.apiErrors).toBeGreaterThan(0);
    const d2 = (await core()).createStrictExcelDouble();
    await expect(d2.Excel.run(async (c: any) => {
      c.workbook.tables.getItemOrNullObject("Nope").getRange();
    })).rejects.toThrow("ItemNotFound");
    expect(d2.apiErrors).toBeGreaterThan(0);
  });

  it("R29 formatting a missing-object range invalidates evidence", async () => {
    const d = (await core()).createStrictExcelDouble();
    await expect(d.Excel.run(async (c: any) => {
      c.workbook.worksheets.getItemOrNullObject("Nope").getRange("A1").format.font.bold = true;
      await c.sync();
    })).rejects.toThrow("ItemNotFound");
    expect(d.mutationCount).toBe(0);
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R30 table creation on a missing-object sheet invalidates evidence", async () => {
    const d = (await core()).createStrictExcelDouble();
    await d.Excel.run(async (c: any) => {
      const r = c.workbook.worksheets.getItem("Sheet1").getRange("A1");
      expect(() => c.workbook.worksheets.getItemOrNullObject("Nope").tables.add(r, true)).toThrow("ItemNotFound");
    });
    expect(d.mutationCount).toBe(0);
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R31 sensitive content in an ordinary property name is rejected", async () => {
    const evaluator = await core();
    const temp = scratch();
    const fixture: any = { caseId: "key-content", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [] }, "sk-AAAAAAAAAAAAAAAAAAAA": "harmless" };
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "key-content.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "key-content.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`key-content.json:${digest}`).digest("hex") }));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow("corpus_sensitive_pattern");
  });

  it("R32 escaped paths inside XML envelopes are rejected", async () => {
    const evaluator = await core();
    const temp = scratch();
    const fixture = { caseId: "xml-escaped-path", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [{ id: "x", type: "function", function: { name: "excel_execute_plan", arguments: "<tool_call>{\"note\":\"\\/home\\/someone\\/private\"}</tool_call><tool_call>{}</tool_call>" } }] } };
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "xml-escaped-path.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "xml-escaped-path.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`xml-escaped-path.json:${digest}`).digest("hex") }));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow("corpus_sensitive_pattern");
  });

  it("R33 a range from another sheet cannot create a table", async () => {
    const d = (await core()).createStrictExcelDouble({ sheets: ["Sheet1", "Sheet2"] });
    await d.Excel.run(async (c: any) => {
      const r = c.workbook.worksheets.getItem("Sheet2").getRange("A1");
      expect(() => c.workbook.worksheets.getItem("Sheet1").tables.add(r, true)).toThrow("InvalidRequestContext");
    });
    expect(d.mutationCount).toBe(0);
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R34 reading a missing-object range invalidates evidence", async () => {
    const d = (await core()).createStrictExcelDouble();
    await expect(d.Excel.run(async (c: any) => {
      const r = c.workbook.worksheets.getItemOrNullObject("Nope").getRange("A1");
      r.load("values"); await c.sync();
    })).rejects.toThrow("ItemNotFound");
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R35 invalid descendant loads are rejected", async () => {
    const d = (await core()).createStrictExcelDouble();
    await expect(d.Excel.run(async (c: any) => {
      const r = c.workbook.worksheets.getItem("Sheet1").getRange("A1");
      r.load("values/noSuchProperty");
      await c.sync();
    })).rejects.toThrow("ApiNotFound");
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R36 nested JSON with escaped quotes inside XML envelopes is rejected", async () => {
    const evaluator = await core();
    const temp = scratch();
    const fixture = { caseId: "xml-nested", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [{ id: "x", type: "function", function: { name: "excel_execute_plan", arguments: '<tool_call>{"note":"{\\"password\\":\\"opaque-value\\"}"}</tool_call><tool_call>{}</tool_call>' } }] } };
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "xml-nested.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "xml-nested.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`xml-nested.json:${digest}`).digest("hex") }));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow("corpus_sensitive_pattern");
  });

  it("R37 invalid Excel coordinates are rejected", async () => {
    const d = (await core()).createStrictExcelDouble();
    await expect(d.Excel.run(async (c: any) => {
      c.workbook.worksheets.getItem("Sheet1").getRange("A0");
    })).rejects.toThrow("double_range_limit");
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R38 empty collections still reject unsupported descendant loads", async () => {
    const d = (await core()).createStrictExcelDouble();
    await expect(d.Excel.run(async (c: any) => {
      c.workbook.tables.load("items/noSuchProperty");
      await c.sync();
    })).rejects.toThrow("ApiNotFound");
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R39 deeply nested JSON inside multiple XML envelopes is rejected", async () => {
    const evaluator = await core();
    const temp = scratch();
    const secret = JSON.stringify({ password: "opaque-value" });
    const nested = JSON.stringify({ note: secret });
    const payload = `<tool_call>${JSON.stringify({ note: nested })}</tool_call><tool_call>{}</tool_call>`;
    const fixture = { caseId: "xml-deep", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [{ id: "x", type: "function", function: { name: "excel_execute_plan", arguments: payload } }] } };
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "xml-deep.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "xml-deep.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`xml-deep.json:${digest}`).digest("hex") }));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow("corpus_sensitive_pattern");
  });

  it("R40 multi-segment descendant paths are rejected even on empty collections", async () => {
    const d = (await core()).createStrictExcelDouble();
    await expect(d.Excel.run(async (c: any) => {
      c.workbook.tables.load("items/name/noSuchProperty");
      await c.sync();
    })).rejects.toThrow("ApiNotFound");
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R41 the state oracle rejects a wrong-value mutation", async () => {
    const evaluator = await core();
    const temp = scratch();
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const fixture = { caseId: "wrong-value", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "Write values to cells", workbookContext: { activeSheet: "Sheet1", selectedRange: "A1:D10" } },
      workbookSeed: { sheets: ["Sheet1"], cells: { "Sheet1!C1": "PRIVATE_CANARY_R6B" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [{ id: "synthetic", type: "function", function: { name: "excel_execute_plan",
        arguments: JSON.stringify({ version: 1, outcome: "execute", summary: "Synthetic change", operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["WRONG_VALUE"]] }], verification: [{ kind: "sentinel_unchanged", sheet: "Sheet1", address: "B1", expected: null }] }) } }] },
      expect: { outcome: "committed", verification: "verified", mutationCount: 1, state: { cells: { "Sheet1!A1": "PRIVATE_CANARY_R6B", "Sheet1!C1": "PRIVATE_CANARY_R6B" }, numberFormats: {}, tables: [] } } };
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "wrong-value.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "wrong-value.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`wrong-value.json:${digest}`).digest("hex") }));
    const result = await evaluator.runWorkbookPlanConformance({ kernel, corpora: [{ path: temp, role: "synthetic" }], clock: () => 1000 });
    expect(result.records[0].status).toBe("fail");
    expect(result.records[0].reasonCode).toBe("expectation_mismatch");
  });

  it("R42 oversized transport payloads are rejected, not treated as clean", async () => {
    const evaluator = await core();
    const temp = scratch();
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const payload = "a".repeat(300000) + "\\u0073k-AAAAAAAAAAAAAAAAAAAA";
    const fixture = { caseId: "oversized", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [{ id: "x", type: "function", function: { name: "excel_execute_plan", arguments: payload } }] } };
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "oversized.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "oversized.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`oversized.json:${digest}`).digest("hex") }));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow("corpus_inspection_limit_exceeded");
  });

  it("R43 nonexistent formatting properties are rejected", async () => {
    const d = (await core()).createStrictExcelDouble();
    await expect(d.Excel.run(async (c: any) => {
      const r = c.workbook.worksheets.getItem("Sheet1").getRange("A1");
      (r.format.fill as any).bold = true;
      await c.sync();
    })).rejects.toThrow("ApiNotFound");
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R44 invalid load arguments are rejected with a recorded API error", async () => {
    const d = (await core()).createStrictExcelDouble();
    await expect(d.Excel.run(async (c: any) => {
      const r = c.workbook.worksheets.getItem("Sheet1").getRange("A1");
      try { (r as any).load(null); } catch { /* swallowed misuse must still record an error */ }
      await c.sync();
    })).resolves.toBeUndefined();
    expect(d.apiErrors).toBeGreaterThan(0);
  });

  it("R45 deep escape chains are rejected, not treated as clean", async () => {
    const evaluator = await core();
    const temp = scratch();
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const payload = "\\" + "u005c".repeat(33) + "u0073k-" + "A".repeat(20);
    const fixture = { caseId: "deep-escape", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [{ id: "x", type: "function", function: { name: "excel_execute_plan", arguments: payload } }] } };
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "deep-escape.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "deep-escape.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`deep-escape.json:${digest}`).digest("hex") }));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow();
  });

  it("R46 escaped property names are rejected", async () => {
    const evaluator = await core();
    const temp = scratch();
    const { writeFileSync } = await import("node:fs");
    const { createHash } = await import("node:crypto");
    const fixture: any = { caseId: "escaped-key", provenance: "synthetic", sourceClass: "generic", transport: "openai-tool-call",
      syntheticInput: { userPrompt: "seed", workbookContext: { activeSheet: "Sheet1" } }, confirmation: "none", faults: {},
      modelOutput: { tool_calls: [] } };
    fixture["\\u0073k-AAAAAAAAAAAAAAAAAAAA"] = "harmless";
    const text = JSON.stringify(fixture); writeFileSync(join(temp, "escaped-key.json"), text);
    const digest = createHash("sha256").update(text).digest("hex");
    writeFileSync(join(temp, "manifest.json"), JSON.stringify({ version: 1, corpusId: "workbook-plan-conformance", corpusVersion: "1.1.0",
      files: { "escaped-key.json": digest }, fixtureCount: 1, corpusSha256: createHash("sha256").update(`escaped-key.json:${digest}`).digest("hex") }));
    expect(() => evaluator.loadCorpus(temp, { role: "synthetic" })).toThrow("corpus_sensitive_pattern");
  });

  it("R47 swallowed setter clone errors still record an API error", async () => {
    const d = (await core()).createStrictExcelDouble();
    await expect(d.Excel.run(async (c: any) => {
      const r = c.workbook.worksheets.getItem("Sheet1").getRange("A1");
      try { (r as any).values = () => 0; } catch { /* swallowed misuse */ }
      r.values = [["ok"]];
      await c.sync();
    })).resolves.toBeUndefined();
    expect(d.apiErrors).toBeGreaterThan(0);
  });
});
