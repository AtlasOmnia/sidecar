// @vitest-environment node
import { existsSync, readFileSync } from "node:fs";
import { EventEmitter } from "node:events";
import { resolve } from "node:path";
import { afterEach, beforeAll, afterAll, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";
import {
  FULL_FUNCTIONALITY_ACCEPTANCE_CASES,
  FULL_FUNCTIONALITY_ACCEPTANCE_ORIGIN,
  FULL_FUNCTIONALITY_ACCEPTANCE_PATH,
  FULL_FUNCTIONALITY_ACCEPTANCE_QUERY_KEY,
  FULL_FUNCTIONALITY_ACCEPTANCE_QUERY_VALUE,
  FULL_FUNCTIONALITY_ACCEPTANCE_URL,
  FULL_FUNCTIONALITY_EVIDENCE_MAX_BYTES,
  FULL_FUNCTIONALITY_SHEET_PREFIX,
  FULL_FUNCTIONALITY_TOOL_NAMES,
  createFullFunctionalityAcceptanceRunId,
  createFullFunctionalityAcceptanceSheetSet,
  createPendingFullFunctionalityAcceptanceEvidence,
  evaluateFullFunctionalityAcceptanceCase,
  installFullFunctionalityAcceptanceHarness,
  isFullFunctionalityAcceptanceLocationAllowed,
  runFullFunctionalityAcceptance,
  serializeFullFunctionalityAcceptanceEvidence,
  type FullFunctionalityAcceptanceCaseDefinition,
  type FullFunctionalityAcceptanceCaseEvidence,
  type FullFunctionalityAcceptancePhaseEvidence,
  type FullFunctionalityAcceptanceToolResultEvidence,
} from "../taskpane/fullFunctionalityAcceptanceHarness";
import { bootstrap } from "../taskpane/bootstrapCoordinator";
import type { BootstrapDeps } from "../taskpane/bootstrapCoordinator";
// @ts-expect-error — ESM .mjs module, no declaration file; vitest resolves at runtime
import { requestText, buildRunnerEnv } from "../../scripts/request-text.mjs";

const REPO_ROOT = resolve(__dirname, "../..");
const RUN_ID = "a1b2c3d4";
const PRIVATE_CONTENT = "PRIVATE-WORKBOOK-CONTENT";
const previousWindow = (globalThis as { window?: unknown }).window;
const previousExcel = (globalThis as { Excel?: unknown }).Excel;

function readRepositoryFile(path: string): string {
  const absolute = resolve(REPO_ROOT, path);
  return existsSync(absolute) ? readFileSync(absolute, "utf8") : "";
}

function phase(
  status: FullFunctionalityAcceptancePhaseEvidence["status"] = "passed",
  matchesExpectation = true,
): FullFunctionalityAcceptancePhaseEvidence {
  return { status, captured: status !== "pending", matchesExpectation };
}

function toolResult(
  status: FullFunctionalityAcceptanceToolResultEvidence["status"] = "passed",
  ok = true,
  outcome: FullFunctionalityAcceptanceToolResultEvidence["outcome"] = "success",
): FullFunctionalityAcceptanceToolResultEvidence {
  return { status, ok, outcome };
}

function definition(
  expectedOutcome: FullFunctionalityAcceptanceCaseDefinition["expectedOutcome"] = "success",
): FullFunctionalityAcceptanceCaseDefinition {
  return {
    id: "write_range:positive",
    tool: "write_range",
    kind: expectedOutcome === "success" ? "positive" : "negative",
    fixture: "data",
    mutates: true,
    expectedOutcome,
    sentinelRequired: true,
    resetFixture: true,
    allowUnsupported: false,
    buildArguments: () => ({}),
    buildPrerequisites: () => [],
  };
}

function evaluate(
  overrides: Partial<{
    pre: FullFunctionalityAcceptancePhaseEvidence;
    result: FullFunctionalityAcceptanceToolResultEvidence;
    post: FullFunctionalityAcceptancePhaseEvidence;
    sentinel: FullFunctionalityAcceptancePhaseEvidence;
  }> = {},
  caseDefinition = definition(),
): FullFunctionalityAcceptanceCaseEvidence {
  return evaluateFullFunctionalityAcceptanceCase(caseDefinition, {
    pre: phase(),
    result: toolResult(),
    post: phase(),
    sentinel: phase(),
    ...overrides,
  });
}

afterEach(() => {
  (globalThis as { window?: unknown }).window = previousWindow;
  (globalThis as { Excel?: unknown }).Excel = previousExcel;
  delete (globalThis as { __hermesExcelFullFunctionalityAcceptance?: unknown })
    .__hermesExcelFullFunctionalityAcceptance;
});

describe("full-functionality acceptance harness guards", () => {
  it("defines one exact HTTPS localhost task-pane opt-in distinct from the focused harness", () => {
    expect(FULL_FUNCTIONALITY_ACCEPTANCE_ORIGIN).toBe("https://localhost:3000");
    expect(FULL_FUNCTIONALITY_ACCEPTANCE_PATH).toBe("/taskpane.html");
    expect(FULL_FUNCTIONALITY_ACCEPTANCE_QUERY_KEY).toBe(
      "fullFunctionalityAcceptance",
    );
    expect(FULL_FUNCTIONALITY_ACCEPTANCE_QUERY_VALUE).toBe("1");
    expect(FULL_FUNCTIONALITY_ACCEPTANCE_URL).toBe(
      "https://localhost:3000/taskpane.html?fullFunctionalityAcceptance=1",
    );
    expect(
      isFullFunctionalityAcceptanceLocationAllowed({
        href: FULL_FUNCTIONALITY_ACCEPTANCE_URL,
      }),
    ).toBe(true);
    expect(
      isFullFunctionalityAcceptanceLocationAllowed({
        href: `${FULL_FUNCTIONALITY_ACCEPTANCE_URL}&_host_Info=Excel$Mac$16.01$en-US$$$$0`,
      }),
    ).toBe(true);
  });

  it("rejects focused-harness URLs, fragments, duplicate metadata, unknown parameters, and endpoint lookalikes", () => {
    const rejected = [
      "https://localhost:3000/taskpane.html?acceptance=1",
      "https://localhost:3000/taskpane.html",
      "https://localhost:3000/?fullFunctionalityAcceptance=1",
      "https://localhost:3000/other.html?fullFunctionalityAcceptance=1",
      "https://localhost:3000/taskpane.html?fullFunctionalityAcceptance=0",
      "https://localhost:3000/taskpane.html?fullFunctionalityAcceptance=1&fullFunctionalityAcceptance=1",
      "https://localhost:3000/taskpane.html?fullFunctionalityAcceptance=1&_host_Info=",
      "https://localhost:3000/taskpane.html?fullFunctionalityAcceptance=1&_host_Info=Excel&_host_Info=Mac",
      "https://localhost:3000/taskpane.html?fullFunctionalityAcceptance=1&extra=1",
      "https://localhost:3000/taskpane.html?fullFunctionalityAcceptance=1#fragment",
      "http://localhost:3000/taskpane.html?fullFunctionalityAcceptance=1",
      "https://127.0.0.1:3000/taskpane.html?fullFunctionalityAcceptance=1",
      "https://localhost.evil:3000/taskpane.html?fullFunctionalityAcceptance=1",
      "https://localhost:3001/taskpane.html?fullFunctionalityAcceptance=1",
      "https://user:pass@localhost:3000/taskpane.html?fullFunctionalityAcceptance=1",
    ];

    for (const href of rejected) {
      expect(
        isFullFunctionalityAcceptanceLocationAllowed({ href }),
        href,
      ).toBe(false);
    }
  });

  it("fails closed when unit tests do not define the acceptance compile flag", () => {
    (globalThis as { window?: unknown }).window = {
      location: { href: FULL_FUNCTIONALITY_ACCEPTANCE_URL },
    };

    expect(installFullFunctionalityAcceptanceHarness()).toBe(false);
    expect(
      (globalThis as {
        window?: { __hermesExcelFullFunctionalityAcceptance?: unknown };
      }).window?.__hermesExcelFullFunctionalityAcceptance,
    ).toBeUndefined();
  });
});

describe("disposable acceptance fixture contract", () => {
  it("creates opaque eight-character run IDs and seven unique valid prefixed sheet names", () => {
    const generatedRunId = createFullFunctionalityAcceptanceRunId();
    expect(generatedRunId).toMatch(/^[a-z0-9]{8}$/);

    const sheets = createFullFunctionalityAcceptanceSheetSet(RUN_ID);
    const names = Object.entries(sheets)
      .filter(([key]) => key !== "runId")
      .map(([, value]) => value);

    expect(sheets.runId).toBe(RUN_ID);
    expect(names).toHaveLength(7);
    expect(new Set(names).size).toBe(7);
    for (const name of names) {
      expect(name.startsWith(`${FULL_FUNCTIONALITY_SHEET_PREFIX}${RUN_ID}_`)).toBe(
        true,
      );
      expect(name.length).toBeLessThanOrEqual(31);
      expect(name).not.toMatch(/[\\/?*:[\]]/);
    }
  });

  it("rejects caller-chosen run IDs that could widen cleanup scope", () => {
    for (const runId of ["", "short", "TOOLONG123", "../evil!", "abcdefgh_"]) {
      expect(() => createFullFunctionalityAcceptanceSheetSet(runId), runId).toThrow(
        /run id/i,
      );
    }
  });
});

describe("27-tool deterministic coverage structure", () => {
  it("derives exactly the canonical 27-tool surface", () => {
    const canonical = excelTools.map(({ schema }) => schema.function.name).sort();
    expect(canonical).toHaveLength(27);
    expect([...FULL_FUNCTIONALITY_TOOL_NAMES].sort()).toEqual(canonical);
  });

  it("declares a positive and negative or boundary case for every canonical tool", () => {
    const ids = FULL_FUNCTIONALITY_ACCEPTANCE_CASES.map(({ id }) => id);
    expect(new Set(ids).size).toBe(ids.length);

    for (const tool of FULL_FUNCTIONALITY_TOOL_NAMES) {
      const cases = FULL_FUNCTIONALITY_ACCEPTANCE_CASES.filter(
        (candidate) => candidate.tool === tool,
      );
      expect(cases.some(({ kind }) => kind === "positive"), tool).toBe(true);
      expect(
        cases.some(({ kind }) => kind === "negative" || kind === "boundary"),
        tool,
      ).toBe(true);
    }
  });

  it("keeps every mutation inside generated fixtures with required sentinel evidence", () => {
    const sheets = createFullFunctionalityAcceptanceSheetSet(RUN_ID);
    for (const acceptanceCase of FULL_FUNCTIONALITY_ACCEPTANCE_CASES) {
      const executableCase = acceptanceCase as unknown as {
        resetFixture?: boolean;
        allowUnsupported?: boolean;
        buildPrerequisites?: (
          sheets: ReturnType<typeof createFullFunctionalityAcceptanceSheetSet>,
        ) => readonly Record<string, unknown>[];
      };
      expect(acceptanceCase.id).toMatch(
        new RegExp(`^${acceptanceCase.tool}:(positive|negative|boundary)(:[a-z0-9-]+)?$`),
      );
      expect(acceptanceCase.buildArguments(sheets)).toEqual(expect.any(Object));
      if (acceptanceCase.mutates) {
        expect(acceptanceCase.sentinelRequired, acceptanceCase.id).toBe(true);
      }
      const serializedArguments = JSON.stringify(
        acceptanceCase.buildArguments(sheets),
      );
      expect(serializedArguments).not.toContain("bridgeSessionId");
      expect(serializedArguments).not.toContain('"password"');
      expect(executableCase.resetFixture, acceptanceCase.id).toBe(true);
      expect(executableCase.buildPrerequisites, acceptanceCase.id).toEqual(
        expect.any(Function),
      );
      const serializedPrerequisites = JSON.stringify(
        executableCase.buildPrerequisites?.(sheets) ?? [],
      );
      expect(serializedPrerequisites).not.toContain("bridgeSessionId");
      expect(serializedPrerequisites).not.toContain('"password"');
    }
  });

  it("declares deterministic independent table, protection, activation, and pivot boundaries", () => {
    const sheets = createFullFunctionalityAcceptanceSheetSet(RUN_ID);
    const byId = Object.fromEntries(
      FULL_FUNCTIONALITY_ACCEPTANCE_CASES.map((acceptanceCase) => [
        acceptanceCase.id,
        acceptanceCase,
      ]),
    );

    expect(byId["insert_table:positive"].buildArguments(sheets)).toMatchObject({
      sheet: sheets.tables,
      address: "H1:K4",
    });
    expect(byId["insert_table:negative"].buildArguments(sheets)).toMatchObject({
      sheet: sheets.tables,
      address: "A1:D4",
    });
    expect(byId["sort_table:positive"].buildArguments(sheets)).toMatchObject({
      tableName: `${sheets.tables}_Table`,
    });
    expect(byId["filter_table:positive"].buildArguments(sheets)).toMatchObject({
      tableName: `${sheets.tables}_Table`,
    });
    expect(byId["activate_sheet:positive"]).toMatchObject({
      mutates: true,
      sentinelRequired: true,
    });
    expect(byId["protect_sheet:positive"].buildArguments(sheets)).toMatchObject({
      sheet: sheets.formats,
    });
    expect(byId["protect_sheet:negative"].buildArguments(sheets)).toMatchObject({
      sheet: sheets.protected,
    });
    expect(byId["unprotect_sheet:boundary"]).toMatchObject({
      expectedOutcome: "success",
    });
    expect(byId["create_pivot_table:negative"]).toMatchObject({
      expectedOutcome: "rejected-without-mutation",
    });
    expect(byId["autofill_formula:positive"].buildArguments(sheets)).toMatchObject({
      sourceAddress: "G2",
      targetAddress: "G2:G8",
    });
    expect(byId["create_chart:positive"].buildArguments(sheets)).toMatchObject({
      sheet: sheets.data,
      dataRange: "B1:C8",
    });
    expect(byId["create_chart:positive"].fixture).toBe("data");

    for (const id of [
      "create_pivot_table:positive",
      "create_pivot_table:negative",
      "add_comment:positive",
      "add_comment:negative",
      "clear_comments:positive",
      "clear_comments:negative",
    ]) {
      expect(
        (byId[id] as unknown as { allowUnsupported?: boolean }).allowUnsupported,
        id,
      ).toBe(true);
    }

    const temporaryNames = [
      "create_sheet:positive",
      "delete_sheet:positive",
      "rename_sheet:positive",
      "copy_sheet:positive",
    ].flatMap((id) =>
      Object.values(byId[id].buildArguments(sheets)).filter(
        (value): value is string =>
          typeof value === "string" && value.startsWith(FULL_FUNCTIONALITY_SHEET_PREFIX),
      ),
    );
    expect(new Set(temporaryNames).size).toBe(temporaryNames.length);

    const harnessSource = readRepositoryFile(
      "src/taskpane/fullFunctionalityAcceptanceHarness.ts",
    );
    expect(harnessSource).toContain("table.name = tableName(sheets)");
    expect(harnessSource).toContain("tableStateMatches");
    expect(harnessSource).toContain("contentMatches");
    expect(harnessSource).toContain("formatStateMatches");
  });
});

describe("content-free evidence and strict evaluation", () => {
  it("preallocates pre/result/post/sentinel evidence for every case and pending later-native slots", () => {
    const evidence = createPendingFullFunctionalityAcceptanceEvidence(RUN_ID);
    expect(evidence.schemaVersion).toBe(1);
    expect(evidence.runId).toBe(RUN_ID);
    expect(evidence.pass).toBe(false);
    expect(evidence.cases).toHaveLength(FULL_FUNCTIONALITY_ACCEPTANCE_CASES.length);
    for (const acceptanceCase of evidence.cases) {
      expect(acceptanceCase.pass).toBe(false);
      expect(acceptanceCase.pre).toEqual(phase("pending", false));
      expect(acceptanceCase.result).toEqual(
        toolResult("pending", false, "malformed"),
      );
      expect(acceptanceCase.post).toEqual(phase("pending", false));
      expect(acceptanceCase.sentinel).toEqual(phase("pending", false));
      expect(acceptanceCase.noPartialMutation).toBe(false);
    }
    expect(evidence.futureNative).toEqual({
      formula: { status: "pending", nativeRequired: true, resultCount: 0 },
      workflow: { status: "pending", nativeRequired: true, resultCount: 0 },
      isolation: { status: "pending", nativeRequired: true, resultCount: 0 },
      recovery: { status: "pending", nativeRequired: true, resultCount: 0 },
    });
  });

  it("requires result, post-state, sentinel, and no-partial-mutation agreement", () => {
    expect(evaluate().pass).toBe(true);
    expect(evaluate().noPartialMutation).toBe(true);
    expect(evaluate({ post: phase("failed", false) }).pass).toBe(false);
    expect(evaluate({ sentinel: phase("failed", false) }).pass).toBe(false);
    expect(evaluate({ pre: phase("pending", false) }).pass).toBe(false);
    expect(
      evaluate(
        {
          result: toolResult("passed", false, "expected-error"),
        },
        definition("rejected-without-mutation"),
      ).pass,
    ).toBe(true);
    expect(
      evaluate(
        {
          result: toolResult("passed", true, "success"),
          post: phase("failed", false),
        },
        definition("rejected-without-mutation"),
      ).pass,
    ).toBe(false);
  });

  it("accepts unsupported results only for cases that explicitly allow them", () => {
    const unsupportedResult = toolResult("unsupported", false, "unsupported");
    const unsupportedDefinition = definition("unsupported-without-mutation");

    expect(evaluate({ result: unsupportedResult }, unsupportedDefinition).pass).toBe(
      false,
    );
    expect(
      evaluate(
        { result: unsupportedResult },
        { ...unsupportedDefinition, allowUnsupported: true },
      ).pass,
    ).toBe(true);
  });

  it("serializes only bounded allow-listed metadata and drops injected workbook content", () => {
    const pending = createPendingFullFunctionalityAcceptanceEvidence(RUN_ID);
    const poisoned = {
      ...pending,
      rawValues: PRIVATE_CONTENT,
      cases: pending.cases.map((acceptanceCase, index) =>
        index === 0
          ? {
              ...acceptanceCase,
              values: [[PRIVATE_CONTENT]],
              formulas: [["=PRIVATE()"]],
              rawError: PRIVATE_CONTENT,
              arguments: { csv: PRIVATE_CONTENT },
            }
          : acceptanceCase,
      ),
    } as typeof pending;

    const serialized = serializeFullFunctionalityAcceptanceEvidence(poisoned);
    const decoded = JSON.parse(serialized) as Record<string, unknown>;
    expect(new TextEncoder().encode(serialized).byteLength).toBeLessThanOrEqual(
      FULL_FUNCTIONALITY_EVIDENCE_MAX_BYTES,
    );
    expect(Object.keys(decoded).sort()).toEqual(
      [
        "cases",
        "cleanup",
        "failCount",
        "futureNative",
        "pass",
        "passCount",
        "runId",
        "schemaVersion",
      ].sort(),
    );
    expect(serialized).not.toContain(PRIVATE_CONTENT);
    expect(serialized).not.toContain("rawValues");
    expect(serialized).not.toContain("values");
    expect(serialized).not.toContain("formulas");
    expect(serialized).not.toContain("rawError");
    expect(serialized).not.toContain("arguments");
    expect(serialized).not.toContain('"csv":');
    expect(serialized).toContain("import_csv_to_sheet");
  });

  it("derives a passing direct-matrix summary only from every strict case and cleanup result", () => {
    const completeCases = FULL_FUNCTIONALITY_ACCEPTANCE_CASES.map(
      (acceptanceCase) =>
        evaluateFullFunctionalityAcceptanceCase(acceptanceCase, {
          pre: phase(),
          result:
            acceptanceCase.expectedOutcome === "success"
              ? toolResult()
              : acceptanceCase.expectedOutcome === "rejected-without-mutation"
                ? toolResult("passed", false, "expected-error")
                : toolResult("unsupported", false, "unsupported"),
          post: phase(),
          sentinel: phase(),
        }),
    );
    const complete = {
      ...createPendingFullFunctionalityAcceptanceEvidence(RUN_ID),
      pass: true,
      cases: completeCases,
      passCount: completeCases.length,
      failCount: 0,
      cleanup: { attempted: true, removedCount: 7, pass: true },
    };

    const decoded = JSON.parse(
      serializeFullFunctionalityAcceptanceEvidence(complete),
    ) as { pass: boolean; passCount: number; failCount: number };
    expect(decoded).toMatchObject({
      pass: true,
      passCount: FULL_FUNCTIONALITY_ACCEPTANCE_CASES.length,
      failCount: 0,
    });
  });

  it("refuses to manufacture native results when the real Excel host is absent", async () => {
    delete (globalThis as { Excel?: unknown }).Excel;
    await expect(runFullFunctionalityAcceptance()).rejects.toThrow(
      /real Excel host required/i,
    );
  });
});

describe("acceptance-only integration artifacts", () => {
  it("loads the dedicated harness through the bootstrap coordinator and compile hook", () => {
    const taskpane = readRepositoryFile("src/taskpane/taskpane.ts");
    const coordinator = readRepositoryFile("src/taskpane/bootstrapCoordinator.ts");
    // taskpane.ts declares the compile flag and delegates to bootstrap()
    expect(taskpane).toContain("__HERMES_OFFICE_ACCEPTANCE__");
    expect(taskpane).toContain("bootstrap(");
    expect(taskpane).toContain("fullFunctionalityAcceptanceHarness");
    // The harness-specific install logic is in taskpane.ts's compile-guarded callbacks,
    // not in the generic coordinator
    expect(taskpane).toContain("installFullFunctionalityAcceptanceHarness");
    // The coordinator itself must NOT contain harness-specific names (it's generic)
    expect(coordinator).not.toContain("installFullFunctionalityAcceptanceHarness");
    expect(coordinator).not.toContain("installOfficeAcceptanceHarness");
  });

  it("provides a server-only runner for the exact URL without gateway, proxy, companion, or model startup", () => {
    const runner = readRepositoryFile("scripts/run-full-excel-acceptance.mjs");
    const reqMod = readRepositoryFile("scripts/request-text.mjs");
    expect(runner).toContain("buildRunnerEnv");
    // HERMES_OFFICE_ACCEPTANCE is now set by buildRunnerEnv in the helper
    expect(reqMod).toContain("HERMES_OFFICE_ACCEPTANCE");
    expect(runner).toContain(FULL_FUNCTIONALITY_ACCEPTANCE_URL);
    expect(runner).toContain("__hermesExcelFullFunctionalityAcceptance");
    expect(runner).not.toMatch(/npm run (proxy|start)/);
    expect(runner).not.toMatch(/excel[_-]companion/i);
    expect(runner).not.toMatch(/gateway/i);
    expect(runner).not.toMatch(/8643|8644|8650/);
    expect(runner).toContain("fileURLToPath(import.meta.url)");
    expect(runner).toContain("const childExit");
    expect(runner).toContain("child.exitCode");
    expect(runner).toContain('require.resolve("webpack-cli/bin/cli.js"');
    expect(runner).not.toContain("spawn(npmCommand");
    expect(runner).toContain('child.kill("SIGKILL")');
  });

  it("documents the static-versus-native evidence boundary and exact cleanup contract", () => {
    const documentation = readRepositoryFile(
      "docs/full-functionality-acceptance.md",
    );
    expect(documentation).toContain(FULL_FUNCTIONALITY_ACCEPTANCE_URL);
    expect(documentation).toContain("27");
    expect(documentation).toMatch(/positive/i);
    expect(documentation).toMatch(/negative|boundary/i);
    expect(documentation).toMatch(/finally/i);
    expect(documentation).toMatch(/content-free/i);
    expect(documentation).toMatch(/does not prove native Excel acceptance/i);
  });

  // -----------------------------------------------------------------------
  // B2 — taskpane bootstrap coordinator behavioral tests
  // -----------------------------------------------------------------------

  describe("B2 bootstrap coordinator behavioral", () => {
    const compileKey = "__HERMES_OFFICE_ACCEPTANCE__" as keyof typeof globalThis;

    beforeAll(() => {
      (globalThis as Record<string, unknown>)[compileKey] = true;
    });

    afterAll(() => {
      delete (globalThis as Record<string, unknown>)[compileKey];
    });

    function flushMicrotasks(): Promise<void> {
      return new Promise((r) => setImmediate(r));
    }

    it("B2-1 full-acceptance lane: install returns false → boot exactly once, focused never loaded", async () => {
      const bootTaskpane = vi.fn();
      const loadAndInstallFocusedHarness = vi.fn();

      const deps: BootstrapDeps = {
        isFullFunctionalityAcceptanceUrl: () => true,
        loadAndInstallFullHarness: () => Promise.resolve(false),
        loadAndInstallFocusedHarness,
        bootTaskpane,
      };

      bootstrap(deps);
      await flushMicrotasks();

      expect(bootTaskpane).toHaveBeenCalledTimes(1);
      expect(loadAndInstallFocusedHarness).not.toHaveBeenCalled();
    });

    it("B2-2 full-acceptance lane: import rejected → boot exactly once", async () => {
      const bootTaskpane = vi.fn();
      const loadAndInstallFocusedHarness = vi.fn();

      const deps: BootstrapDeps = {
        isFullFunctionalityAcceptanceUrl: () => true,
        loadAndInstallFullHarness: () => Promise.reject(new Error("import failed")),
        loadAndInstallFocusedHarness,
        bootTaskpane,
      };

      bootstrap(deps);
      await flushMicrotasks();

      expect(bootTaskpane).toHaveBeenCalledTimes(1);
      expect(loadAndInstallFocusedHarness).not.toHaveBeenCalled();
    });

    it("B2-3 full-acceptance lane: install throws → boot exactly once", async () => {
      const bootTaskpane = vi.fn();
      const loadAndInstallFocusedHarness = vi.fn();

      const deps: BootstrapDeps = {
        isFullFunctionalityAcceptanceUrl: () => true,
        loadAndInstallFullHarness: () => Promise.reject(new Error("install exploded")),
        loadAndInstallFocusedHarness,
        bootTaskpane,
      };

      bootstrap(deps);
      await flushMicrotasks();

      expect(bootTaskpane).toHaveBeenCalledTimes(1);
      expect(loadAndInstallFocusedHarness).not.toHaveBeenCalled();
    });

    it("B2-4 full-acceptance lane: install returns true → normal boot zero times", async () => {
      const bootTaskpane = vi.fn();
      const loadAndInstallFocusedHarness = vi.fn();

      const deps: BootstrapDeps = {
        isFullFunctionalityAcceptanceUrl: () => true,
        loadAndInstallFullHarness: () => Promise.resolve(true),
        loadAndInstallFocusedHarness,
        bootTaskpane,
      };

      bootstrap(deps);
      await flushMicrotasks();

      expect(bootTaskpane).not.toHaveBeenCalled();
      expect(loadAndInstallFocusedHarness).not.toHaveBeenCalled();
    });

    it("B2-5 normal lane: boots normal exactly once, focused install called", async () => {
      const bootTaskpane = vi.fn();
      const loadAndInstallFullHarness = vi.fn();
      let focusedInstalled = false;
      const loadAndInstallFocusedHarness = () => {
        focusedInstalled = true;
        return Promise.resolve();
      };

      const deps: BootstrapDeps = {
        isFullFunctionalityAcceptanceUrl: () => false,
        loadAndInstallFullHarness,
        loadAndInstallFocusedHarness,
        bootTaskpane,
      };

      bootstrap(deps);
      // Normal lane boots synchronously
      expect(bootTaskpane).toHaveBeenCalledTimes(1);
      expect(loadAndInstallFullHarness).not.toHaveBeenCalled();

      // Focused harness install happens asynchronously
      await flushMicrotasks();
      expect(focusedInstalled).toBe(true);
    });

    it("B2-6 focused import rejection is consumed without unhandled rejection", async () => {
      const bootTaskpane = vi.fn();
      const loadAndInstallFullHarness = vi.fn();

      // Track unhandled rejections
      const unhandled: unknown[] = [];
      const onUnhandled = (reason: unknown) => {
        unhandled.push(reason);
      };
      process.on("unhandledRejection", onUnhandled);

      const deps: BootstrapDeps = {
        isFullFunctionalityAcceptanceUrl: () => false,
        loadAndInstallFullHarness,
        loadAndInstallFocusedHarness: () => Promise.reject(new Error("focused import failed")),
        bootTaskpane,
      };

      try {
        bootstrap(deps);
        // Normal boot already happened synchronously
        expect(bootTaskpane).toHaveBeenCalledTimes(1);
        expect(loadAndInstallFullHarness).not.toHaveBeenCalled();

        // Flush microtasks so the rejected focused promise settles
        await flushMicrotasks();
        // Give Node a macrotask tick to fire unhandledRejection if any
        await new Promise((r) => setTimeout(r, 50));

        expect(unhandled).toHaveLength(0);
      } finally {
        process.off("unhandledRejection", onUnhandled);
      }
    });
  });

  // -----------------------------------------------------------------------
  // B3 — requestText / buildRunnerEnv behavioral tests
  // -----------------------------------------------------------------------

  describe("B3 requestText behavioral", () => {
    // Alias for a fake request: EventEmitter + destroy()
    type FakeRequest = EventEmitter & { destroy: () => void };
    // Alias for a fake response: EventEmitter + statusCode + resume()
    type FakeResponse = EventEmitter & { statusCode: number; resume: () => void };

    function makeFakeRequest(): FakeRequest {
      const r = new EventEmitter() as unknown as FakeRequest;
      (r as unknown as { destroy?: unknown }).destroy = vi.fn();
      return r;
    }

    function makeFakeResponse(statusCode = 200): FakeResponse {
      const r = new EventEmitter() as unknown as FakeResponse;
      r.statusCode = statusCode;
      r.resume = vi.fn();
      return r;
    }

    it("B3-1 resolves null within absolute deadline when server trickles without ending", async () => {
      const fakeRequest = makeFakeRequest();
      const fakeResponse = makeFakeResponse();

      const fakeTransport = {
        get: (
          _url: string,
          _opts: unknown,
          cb: (res: FakeResponse) => void,
        ) => {
          // Schedule response on next tick so the setTimeout in requestText
          // registers before any data arrives.
          setImmediate(() => cb(fakeResponse));
          return fakeRequest;
        },
      };

      const startedAt = Date.now();
      const result = await requestText("https://example.com/", {
        timeout: 80,
        transport: fakeTransport,
      } as never);
      const elapsed = Date.now() - startedAt;

      expect(result).toBeNull();
      // Must resolve within bounded window (absolute deadline + overhead)
      expect(elapsed).toBeLessThan(500);
      // Request must be destroyed exactly once
      expect(fakeRequest.destroy).toHaveBeenCalledTimes(1);
    });

    it("B3-2 ignores late events after absolute deadline settlement", async () => {
      let responseCb!: (res: FakeResponse) => void;

      const fakeRequest = makeFakeRequest();

      const fakeTransport = {
        get: (
          _url: string,
          _opts: unknown,
          cb: (res: FakeResponse) => void,
        ) => {
          responseCb = cb;
          return fakeRequest;
        },
      };

      const resultPromise = requestText("https://example.com/", {
        timeout: 50,
        transport: fakeTransport,
      } as never);

      // Let the deadline fire first (don't invoke the response callback)
      await new Promise((r) => setTimeout(r, 100));

      // Now deliver the response — it's late, should be ignored
      const fakeResponse = makeFakeResponse();
      responseCb(fakeResponse);

      // Emit late data and end events
      fakeResponse.emit("data", Buffer.from("late-data"));
      fakeResponse.emit("end");

      const result = await resultPromise;
      expect(result).toBeNull();
      // Destroy must have been called exactly once (by the deadline)
      expect(fakeRequest.destroy).toHaveBeenCalledTimes(1);
    });

    it("B3-3 clears cleanup timer on successful response", async () => {
      const clearTimeoutSpy = vi.spyOn(globalThis, "clearTimeout" as never);

      const fakeRequest = makeFakeRequest();
      const fakeResponse = makeFakeResponse();

      const fakeTransport = {
        get: (
          _url: string,
          _opts: unknown,
          cb: (res: FakeResponse) => void,
        ) => {
          setImmediate(() => {
            cb(fakeResponse);
            // Emit data and end promptly
            fakeResponse.emit("data", Buffer.from("hello"));
            fakeResponse.emit("end");
          });
          return fakeRequest;
        },
      };

      const result = await requestText("https://example.com/", {
        timeout: 500,
        transport: fakeTransport,
      } as never);

      expect(result).toBe("hello");
      // clearTimeout must have been called (at least once)
      const clearCalls = clearTimeoutSpy.mock.calls.filter(
        (call) => typeof call[0] === "object" || typeof call[0] === "number",
      );
      expect(clearCalls.length).toBeGreaterThanOrEqual(1);

      clearTimeoutSpy.mockRestore();
    });

    it("B3-4 buildRunnerEnv does not inject a machine-specific PYTHON key", () => {
      const env = buildRunnerEnv({ PATH: "/usr/bin", HOME: "/home/test" });
      expect(env).toHaveProperty("NODE_ENV", "development");
      expect(env).toHaveProperty("HERMES_OFFICE_ACCEPTANCE", "1");
      expect(env).toHaveProperty("PATH", "/usr/bin");
      expect(env).toHaveProperty("HOME", "/home/test");
      // Must not inject any PYTHON key
      expect(Object.keys(env)).not.toContain("PYTHON");
      // Also not a case variant
      const keys = Object.keys(env);
      expect(keys.filter((k) => k.toUpperCase() === "PYTHON")).toHaveLength(0);
    });
  });
});
