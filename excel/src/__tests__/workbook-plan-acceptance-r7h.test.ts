import { describe, it, expect, vi, afterEach } from "vitest";
import { resolve } from "node:path";
import { readFileSync, existsSync } from "node:fs";
import { createRequire } from "node:module";
import ts from "typescript";
import { StrictWorkbookPlanExcelDouble } from "./fixtures/strictWorkbookPlanExcelDouble";

function disposableRuntime(host: StrictWorkbookPlanExcelDouble) {
  let opened = false;
  let cleanupCount = 0;
  return {
    get cleanupCount() { return cleanupCount; },
    async createDisposableWorkbook(runId: string, sheetNames: readonly string[]) {
      expect(opened).toBe(false); opened = true;
      for (const sheet of sheetNames) { expect(sheet.startsWith(`R7H_${runId}_`)).toBe(true); host.sheets.set(sheet, [[null]]); }
    },
    async assertDisposableIdentity(runId: string, sheetNames: readonly string[]) {
      if (!opened || sheetNames.length !== 7 || host.sheets.size !== 7 || sheetNames.some(n => !n.startsWith(`R7H_${runId}_`) || !host.sheets.has(n))) throw new Error("wrong disposable identity");
    },
    async cleanupDisposableWorkbook() { cleanupCount++; host.sheets.clear(); opened = false; },
  };
}


async function moduleUnderTest() {
  try { return await import(/* @vite-ignore */ resolve(__dirname, "../taskpane/workbookPlanAcceptanceHarness")); }
  catch { return undefined; }
}
afterEach(() => vi.unstubAllGlobals());
describe("R7-H offline native matrix", () => {
  it("has no environment-controlled arbitrary receipt-file write", () => {
    const source = readFileSync(resolve(__dirname, "workbook-plan-acceptance-r7h.test.ts"), "utf8");
    expect(source).not.toContain("process.env." + "R7H_RECEIPT_PATH");
  });
  it("strict double requires load/sync and commits writes only at sync", async () => {
    const host = new StrictWorkbookPlanExcelDouble(); host.sheets.set("synthetic", [[null]]);
    await host.run(async value => {
      const ctx = value as { workbook: { worksheets: { getItem: (name: string) => { getRange: (address: string) => { values: unknown; load: (fields: string) => void } } } }; sync: () => Promise<void> };
      const range = ctx.workbook.worksheets.getItem("synthetic").getRange("A1");
      expect(() => range.values).toThrow("read before load+sync");
      range.load("values"); expect(() => range.values).toThrow("read before load+sync");
      await ctx.sync(); expect(range.values).toEqual([[null]]);
      range.values = [["queued"]]; expect(host.sheets.get("synthetic")).toEqual([[null]]);
      await ctx.sync(); expect(host.sheets.get("synthetic")).toEqual([["queued"]]);
      return undefined;
    });
    expect(host.writes).toBe(1);
  });
  it("runs every direct/server case through native Office readback and emits metadata only", async () => {
    vi.stubGlobal("__HERMES_OFFICE_ACCEPTANCE__", true);
    const host = new StrictWorkbookPlanExcelDouble(); vi.stubGlobal("Excel", { run: host.run });
    const runtime = disposableRuntime(host);
    const m = await moduleUnderTest(); expect(m).toBeDefined();
    const receipts = await m!.runWorkbookPlanAcceptance(runtime, "offline01");

    expect(receipts).toHaveLength(20);
    expect(receipts.every((r: { pass: boolean }) => r.pass), JSON.stringify(receipts.filter((r: { pass: boolean }) => !r.pass))).toBe(true);
    expect(new Set(receipts.map((r: { route: string; caseId: string }) => `${r.route}:${r.caseId}`)).size).toBe(20);
    expect(host.writes).toBe(8); // execute, inspect-execute, replay seed, recovery after rotation per route
    expect(host.runs).toBeGreaterThan(20);
    expect(host.violations).toEqual([]);
    expect(runtime.cleanupCount).toBe(1);
    expect(host.sheets.size).toBe(0);
    for (const receipt of receipts) expect(Object.keys(receipt).sort()).toEqual(["caseId", "correctionCount", "executionOutcome", "inspectionCount", "outcome", "pass", "reason", "route", "verification"].sort());
    expect(JSON.stringify(receipts)).not.toContain("synthetic-value");
    expect(JSON.stringify(receipts)).not.toContain("R7H_offline01");
  });
  it("cleans up after identity refusal and makes zero mutations", async () => {
    vi.stubGlobal("__HERMES_OFFICE_ACCEPTANCE__", true);
    const host = new StrictWorkbookPlanExcelDouble(); vi.stubGlobal("Excel", { run: host.run });
    const runtime = disposableRuntime(host);
    runtime.assertDisposableIdentity = async () => { throw new Error("wrong disposable identity"); };
    const m = await moduleUnderTest(); expect(m).toBeDefined();
    await expect(m!.runWorkbookPlanAcceptance(runtime, "offline02")).rejects.toThrow("wrong disposable identity");
    expect(runtime.cleanupCount).toBe(1); expect(host.writes).toBe(0);
  });
});

describe("R7-H acceptance boundary", () => {
  it("exports a real guarded matrix harness (test-first missing-module RED)", async () => {
    expect(await moduleUnderTest()).toBeDefined();
  });
  it("fails closed without the acceptance compile definition", async () => {
    const m = await moduleUnderTest();
    expect(m).toBeDefined();
    await expect(m!.runWorkbookPlanAcceptance({})).rejects.toThrow("acceptance build");
  });
  it.each([
    "http://localhost:3000/taskpane.html?workbookPlanAcceptance=1",
    "https://evil.test/taskpane.html?workbookPlanAcceptance=1",
    "https://localhost:3000/taskpane.html?acceptance=1",
    "https://localhost:3000/taskpane.html?workbookPlanAcceptance=1&workbookPlanAcceptance=1",
    "https://localhost:3000/taskpane.html?workbookPlanAcceptance=1#run",
    "https://localhost:3000/taskpane.html?workbookPlanAcceptance=1&unknown=1",
  ])("rejects non-exclusive acceptance location %s", async href => {
    const m = await moduleUnderTest(); expect(m).toBeDefined();
    expect(m!.isWorkbookPlanAcceptanceLocationAllowed({ href })).toBe(false);
  });
  it("accepts the dedicated query with one Office host parameter", async () => {
    const m = await moduleUnderTest(); expect(m).toBeDefined();
    expect(m!.isWorkbookPlanAcceptanceLocationAllowed({ href: "https://localhost:3000/taskpane.html?workbookPlanAcceptance=1&_host_Info=Excel" })).toBe(true);
  });
  it("uses a separate conditional acceptance entry without running webpack", async () => {
    const factory = createRequire(import.meta.url)("../../webpack.config.js");
    const previous = process.env.HERMES_OFFICE_ACCEPTANCE;
    try {
      delete process.env.HERMES_OFFICE_ACCEPTANCE;
      const prod = await factory({}, { mode: "production" });
      expect(prod.entry).toEqual({ omniaTheme: "./src/taskpane/omniaThemeBootstrap.ts", taskpane: "./src/taskpane/taskpane.ts" });
      process.env.HERMES_OFFICE_ACCEPTANCE = "1";
      const acceptance = await factory({}, { mode: "production" });
      expect(acceptance.entry.workbookPlanAcceptance).toBe("./src/taskpane/workbookPlanAcceptanceEntry.ts");
    } finally {
      if (previous === undefined) delete process.env.HERMES_OFFICE_ACCEPTANCE; else process.env.HERMES_OFFICE_ACCEPTANCE = previous;
    }
  });
  it("installs only with compile guard and valid dedicated location", async () => {
    const m = await moduleUnderTest(); expect(m).toBeDefined();
    const runtime = disposableRuntime(new StrictWorkbookPlanExcelDouble());
    const target: Record<string, unknown> = { __hermesExcelWorkbookPlanAcceptance: {} };
    const location = { href: "https://localhost:3000/taskpane.html?workbookPlanAcceptance=1" };
    expect(m!.installWorkbookPlanAcceptanceHarness(target, runtime, location)).toBe(false);
    expect(target).not.toHaveProperty("__hermesExcelWorkbookPlanAcceptance");
    vi.stubGlobal("__HERMES_OFFICE_ACCEPTANCE__", true);
    expect(m!.installWorkbookPlanAcceptanceHarness(target, runtime, location)).toBe(true);
    expect(target).toHaveProperty("__hermesExcelWorkbookPlanAcceptance.run");
    expect(m!.installWorkbookPlanAcceptanceHarness(target, runtime, { href: location.href + "&acceptance=1" })).toBe(false);
    expect(target).not.toHaveProperty("__hermesExcelWorkbookPlanAcceptance");
  });
  it("production entry transitive graph excludes the harness and acceptance entry", () => {
    const seen = new Set<string>();
    function visit(file: string) {
      if (seen.has(file)) return; seen.add(file);
      const source = ts.createSourceFile(file, readFileSync(file, "utf8"), ts.ScriptTarget.Latest, true);
      function scan(node: ts.Node) {
        let spec: string | undefined;
        if ((ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) && node.moduleSpecifier && ts.isStringLiteral(node.moduleSpecifier)) spec = node.moduleSpecifier.text;
        if (ts.isCallExpression(node) && node.expression.kind === ts.SyntaxKind.ImportKeyword && ts.isStringLiteral(node.arguments[0])) spec = node.arguments[0].text;
        if (spec?.startsWith(".")) {
          const target = resolve(file, "..", spec);
          for (const suffix of [".ts", "/index.ts"]) if (existsSync(target + suffix)) { visit(target + suffix); break; }
        }
        ts.forEachChild(node, scan);
      }
      scan(source);
    }
    visit(resolve(__dirname, "../taskpane/taskpane.ts"));
    expect([...seen].filter(f => /workbookPlanAcceptance/.test(f))).toEqual([]);
    const config = readFileSync(resolve(__dirname, "../../webpack.config.js"), "utf8");
    expect(config).toContain('...(acceptanceEnabled ? { workbookPlanAcceptance: "./src/taskpane/workbookPlanAcceptanceEntry.ts" } : {})');
  });
});
