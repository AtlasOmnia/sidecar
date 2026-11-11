import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  installOfficeAcceptanceHarness,
  type AcceptanceFixture,
  type AcceptanceFixtureEvidence,
  type AcceptanceFixtureRuntime,
  type AcceptanceHarnessRunResult,
} from "../taskpane/officeAcceptanceHarness";

const SHEET = "Tool Acceptance";
const ADDRESS = "A1:F1";
const ACCEPTANCE_TITLE = "Acceptance title";
const AUTOFIT_ACCEPTANCE_ANCHOR =
  "Native autofit acceptance anchor with deterministic width change";
const PRIVATE_CONTENT = "PRIVATE-WORKBOOK-CONTENT";

interface FakeState {
  values: unknown[][];
  formulas: string[][];
  merged: boolean;
  alignment: string;
  wrapText: boolean;
  columnWidths: number[];
  autofitColumnsCalled: boolean;
  formatAssignments: FormatAssignment[];
  excelRunCount: number;
  rawEvents: string[];
}

interface FakeRuntimeOptions {
  readonly mutateWidthAfterToolFor?: readonly AcceptanceFixture["id"][];
  readonly autofitChangesWidth?: boolean;
}

interface FormatAssignment {
  fixtureId: AcceptanceFixture["id"];
  property: "horizontalAlignment" | "wrapText" | "autofitColumns";
  value?: string | boolean;
}

let previousWindow: unknown;
let previousExcel: unknown;
let previousOffice: unknown;

beforeEach(() => {
  previousWindow = (globalThis as { window?: unknown }).window;
  previousExcel = (globalThis as { Excel?: unknown }).Excel;
  previousOffice = (globalThis as { Office?: unknown }).Office;
  delete (globalThis as { window?: unknown }).window;
  delete (globalThis as { __hermesExcelAcceptance?: unknown }).__hermesExcelAcceptance;
});

afterEach(() => {
  (globalThis as { window?: unknown }).window = previousWindow;
  (globalThis as { Excel?: unknown }).Excel = previousExcel;
  (globalThis as { Office?: unknown }).Office = previousOffice;
  delete (globalThis as { __hermesExcelAcceptance?: unknown }).__hermesExcelAcceptance;
});

function makeFakeRuntime(
  options: FakeRuntimeOptions = {},
): { runtime: AcceptanceFixtureRuntime; state: FakeState } {
  const state: FakeState = {
    values: [[null, null, null, null, null, null]],
    formulas: [["", "", "", "", "", ""]],
    merged: false,
    alignment: "General",
    wrapText: true,
    columnWidths: [8, 8, 8, 8, 8, 8],
    autofitColumnsCalled: false,
    formatAssignments: [],
    excelRunCount: 0,
    rawEvents: [],
  };

  let activeFixture: AcceptanceFixture | undefined;
  let readbackCount = 0;
  const autofitChangesWidth = options.autofitChangesWidth ?? true;

  const range = {
    rowCount: 1,
    columnCount: 6,
    get values() {
      return state.values;
    },
    set values(value: unknown[][]) {
      state.values = value;
    },
    get formulas() {
      return state.formulas;
    },
    set formulas(value: string[][]) {
      state.formulas = value;
    },
    address: `${SHEET}!${ADDRESS}`,
    load: () => undefined,
    getMergedAreasOrNullObject: () => ({
      isNullObject: !state.merged,
      load: () => undefined,
    }),
    format: {
      get horizontalAlignment() {
        return state.alignment;
      },
      set horizontalAlignment(value: string) {
        state.alignment = value;
        if (activeFixture) {
          state.formatAssignments.push({
            fixtureId: activeFixture.id,
            property: "horizontalAlignment",
            value,
          });
        }
      },
      get wrapText() {
        return state.wrapText;
      },
      set wrapText(value: boolean) {
        state.wrapText = value;
        if (activeFixture) {
          state.formatAssignments.push({
            fixtureId: activeFixture.id,
            property: "wrapText",
            value,
          });
        }
      },
      autofitColumns: () => {
        state.autofitColumnsCalled = true;
        if (activeFixture) {
          state.formatAssignments.push({
            fixtureId: activeFixture.id,
            property: "autofitColumns",
          });
        }
        if (autofitChangesWidth) {
          state.columnWidths[0] = 24;
        }
      },
    },
    select: () => undefined,
  };

  const runtime: AcceptanceFixtureRuntime = {
    async prepare() {
      (globalThis as { Office?: unknown }).Office = {
        context: {
          requirements: {
            isSetSupported: () => true,
          },
        },
      };
      (globalThis as { Excel?: unknown }).Excel = {
        run: async (callback: (context: unknown) => Promise<unknown>) => {
          state.excelRunCount += 1;
          return callback({
            workbook: {
              worksheets: {
                getItem: () => ({ getRange: () => range }),
              },
            },
            sync: async () => undefined,
          });
        },
      };
    },
    async setup(fixture: AcceptanceFixture) {
      activeFixture = fixture;
      readbackCount = 0;
      state.values = [[
        fixture.id === "column-autofit-true" ? AUTOFIT_ACCEPTANCE_ANCHOR : ACCEPTANCE_TITLE,
        null,
        null,
        null,
        null,
        null,
      ]];
      state.formulas = [["", "", "", "", "", ""]];
      state.merged = fixture.id === "pre-merged-rejection";
      state.alignment = "General";
      state.wrapText = true;
      state.columnWidths = [8, 8, 8, 8, 8, 8];
      state.autofitColumnsCalled = false;
      if (fixture.id === "occupied-c1-precondition") {
        state.values[0][2] = PRIVATE_CONTENT;
      }
      state.rawEvents.push(`setup:${fixture.sheet}:${fixture.address}`);
      if (state.merged) {
        state.rawEvents.push(`merge:${fixture.sheet}:${fixture.address}`);
      }
    },
    async readback(): Promise<AcceptanceFixtureEvidence> {
      readbackCount += 1;
      if (
        readbackCount === 2 &&
        activeFixture &&
        options.mutateWidthAfterToolFor?.includes(activeFixture.id)
      ) {
        state.columnWidths[5] = 9;
      }
      const nonAnchorValues = state.values[0].slice(1).filter((value) =>
        value !== null && value !== undefined && value !== "",
      ).length;
      const nonAnchorFormulas = state.formulas[0].slice(1).filter((value) => value !== "").length;
      return {
        dimensions: { rows: 1, columns: 6 },
        anchorMatchesExpected:
          state.values[0][0] ===
          (activeFixture?.id === "column-autofit-true"
            ? AUTOFIT_ACCEPTANCE_ANCHOR
            : ACCEPTANCE_TITLE),
        nonAnchorValueCount: nonAnchorValues,
        nonAnchorFormulaCount: nonAnchorFormulas,
        merged: state.merged,
        alignment: state.alignment,
        wrapText: state.wrapText,
        columnWidths: [...state.columnWidths],
        autofitColumnsCalled: state.autofitColumnsCalled,
      };
    },
    async cleanup(fixture: AcceptanceFixture) {
      if (fixture.id === "pre-merged-rejection") {
        expect(fixture.sheet).toBe(SHEET);
        expect(fixture.address).toBe(ADDRESS);
        state.merged = false;
        state.rawEvents.push(`unmerge:${fixture.sheet}:${fixture.address}`);
        return { unmerged: true };
      }
      return { unmerged: false };
    },
    async dispose() {
      state.rawEvents.push(`dispose:${SHEET}`);
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    },
  };

  return { runtime, state };
}

function installForTests(runtime: AcceptanceFixtureRuntime, url = "https://localhost:3000/?acceptance=1") {
  (globalThis as { window?: unknown }).window = {
    location: new URL(url),
  };
  return installOfficeAcceptanceHarness(runtime);
}

async function runInstalledHarness(): Promise<AcceptanceHarnessRunResult> {
  const harness = (globalThis as {
    window?: { __hermesExcelAcceptance?: { run: () => Promise<AcceptanceHarnessRunResult> } };
  }).window?.__hermesExcelAcceptance;
  if (!harness) {
    throw new Error("acceptance harness was not installed");
  }
  return harness.run();
}

describe("guarded Office acceptance harness", () => {
  it("allows Office task-pane URLs with one acceptance flag and one host metadata value", () => {
    const { runtime } = makeFakeRuntime();

    expect(
      installForTests(
        runtime,
        "https://localhost:3000/taskpane.html?acceptance=1&_host_Info=Excel$Mac$16.01$en-US$$$$0",
      ),
    ).toBe(true);
  });

  it("rejects malformed acceptance metadata and endpoint lookalikes", () => {
    const { runtime } = makeFakeRuntime();
    const rejectedUrls = [
      "https://localhost:3000/taskpane.html",
      "https://localhost:3000/taskpane.html?acceptance=0",
      "https://localhost:3000/taskpane.html?acceptance=1&acceptance=1",
      "https://localhost:3000/taskpane.html?acceptance=1&_host_Info=Excel&_host_Info=Mac",
      "https://localhost:3000/taskpane.html?acceptance=1&_host_Info=",
      "https://localhost:3000/taskpane.html?acceptance=1&extra=1",
      "http://localhost:3000/taskpane.html?acceptance=1",
      "https://127.0.0.1:3000/taskpane.html?acceptance=1",
      "https://localhost.evil:3000/taskpane.html?acceptance=1",
      "https://localhost:3001/taskpane.html?acceptance=1",
    ];

    for (const url of rejectedUrls) {
      expect(installForTests(runtime, url), url).toBe(false);
    }
  });

  it("runs every bounded fixture through the real validated dispatcher path", async () => {
    const { runtime, state } = makeFakeRuntime();
    expect(installForTests(runtime)).toBe(true);

    const result = await runInstalledHarness();

    expect(result.sheet).toBe(SHEET);
    expect(result.results.map(({ id }) => id)).toEqual([
      "valid-center-across-selection",
      "column-autofit-false",
      "column-autofit-true",
      "occupied-c1-precondition",
      "pre-merged-rejection",
    ]);
    expect(result.passCount).toBe(5);
    expect(result.failCount).toBe(0);
    expect(state.excelRunCount).toBe(5);
  });

  it("proves omitted and false autofit preserve widths while true autofit is observed", async () => {
    const { runtime } = makeFakeRuntime();
    expect(installForTests(runtime)).toBe(true);
    const result = await runInstalledHarness();
    const byId = Object.fromEntries(result.results.map((fixture) => [fixture.id, fixture]));

    expect(byId["valid-center-across-selection"]).toMatchObject({
      ok: true,
      pass: true,
      dimensions: { rows: 1, columns: 6 },
      autofitColumnsCalled: false,
      widthsChanged: false,
      mergedBefore: false,
      mergedAfter: false,
    });
    expect(byId["column-autofit-false"]).toMatchObject({
      ok: true,
      pass: true,
      autofitColumnsCalled: false,
      widthsChanged: false,
    });
    expect(byId["column-autofit-true"]).toMatchObject({
      ok: true,
      pass: true,
      autofitColumnsCalled: true,
      widthsChanged: true,
    });
  });

  it("keeps the column-autofit-true request independent of alignment and wrapping", async () => {
    const { runtime, state } = makeFakeRuntime();
    expect(installForTests(runtime)).toBe(true);

    const result = await runInstalledHarness();
    const fixture = result.results.find(({ id }) => id === "column-autofit-true");

    expect(fixture).toMatchObject({
      ok: true,
      pass: true,
      formattingChanged: false,
      widthsChanged: true,
    });
    expect(
      state.formatAssignments.filter(({ fixtureId }) => fixtureId === "column-autofit-true"),
    ).toEqual([
      {
        fixtureId: "column-autofit-true",
        property: "autofitColumns",
      },
    ]);
  });

  it("fails column-autofit-true when native readback reports unchanged widths", async () => {
    const { runtime } = makeFakeRuntime({ autofitChangesWidth: false });
    expect(installForTests(runtime)).toBe(true);

    const result = await runInstalledHarness();
    const fixture = result.results.find(({ id }) => id === "column-autofit-true");

    expect(fixture).toMatchObject({
      ok: true,
      pass: false,
      dimensions: { rows: 1, columns: 6 },
      autofitColumnsCalled: true,
      widthsChanged: false,
      formattingChanged: false,
      mergedBefore: false,
      mergedAfter: false,
      nonAnchorValueCount: 0,
      nonAnchorFormulaCount: 0,
    });
  });

  it("requires all six widths to stay unchanged for omitted and false autofit", async () => {
    const { runtime } = makeFakeRuntime({
      mutateWidthAfterToolFor: ["valid-center-across-selection", "column-autofit-false"],
    });
    expect(installForTests(runtime)).toBe(true);

    const result = await runInstalledHarness();
    const byId = Object.fromEntries(result.results.map((fixture) => [fixture.id, fixture]));

    expect(byId["valid-center-across-selection"]).toMatchObject({
      ok: true,
      pass: false,
      widthsChanged: true,
    });
    expect(byId["column-autofit-false"]).toMatchObject({
      ok: true,
      pass: false,
      widthsChanged: true,
    });
  });

  it("records the occupied C1 precondition rejection without mutation", async () => {
    const { runtime } = makeFakeRuntime();
    expect(installForTests(runtime)).toBe(true);
    const result = await runInstalledHarness();
    const fixture = result.results.find(({ id }) => id === "occupied-c1-precondition");

    expect(fixture).toMatchObject({
      ok: false,
      pass: true,
      nonAnchorValueCount: 1,
      formattingChanged: false,
      widthsChanged: false,
      mergedBefore: false,
      mergedAfter: false,
    });
    expect(fixture?.message).toContain("C1");
  });

  it("rejects a pre-merged fixture and unmerges only that disposable span after evidence", async () => {
    const { runtime, state } = makeFakeRuntime();
    expect(installForTests(runtime)).toBe(true);
    const result = await runInstalledHarness();
    const fixture = result.results.find(({ id }) => id === "pre-merged-rejection");

    expect(fixture).toMatchObject({
      ok: false,
      pass: true,
      mergedBefore: true,
      mergedAfter: true,
      formattingChanged: false,
      widthsChanged: false,
      cleanupUnmerged: true,
    });
    expect(state.rawEvents.filter((event) => event.startsWith("merge:") || event.startsWith("unmerge:"))).toEqual([
      `merge:${SHEET}:${ADDRESS}`,
      `unmerge:${SHEET}:${ADDRESS}`,
    ]);
    expect(state.rawEvents).not.toContain("merge:Sheet1:A1:F1");
    expect(state.rawEvents).not.toContain("unmerge:Sheet1:A1:F1");
  });

  it("requires all six widths to stay unchanged for occupied and merged rejection", async () => {
    const { runtime } = makeFakeRuntime({
      mutateWidthAfterToolFor: ["occupied-c1-precondition", "pre-merged-rejection"],
    });
    expect(installForTests(runtime)).toBe(true);

    const result = await runInstalledHarness();
    const byId = Object.fromEntries(result.results.map((fixture) => [fixture.id, fixture]));

    expect(byId["occupied-c1-precondition"]).toMatchObject({
      ok: false,
      pass: false,
      widthsChanged: true,
    });
    expect(byId["pre-merged-rejection"]).toMatchObject({
      ok: false,
      pass: false,
      widthsChanged: true,
    });
  });

  it("returns only bounded fixture evidence and never workbook content", async () => {
    const { runtime } = makeFakeRuntime();
    expect(installForTests(runtime)).toBe(true);
    const result = await runInstalledHarness();
    const serialized = JSON.stringify(result);
    const allowedRunKeys = ["sheet", "results", "passCount", "failCount"];
    const allowedFixtureKeys = [
      "id",
      "ok",
      "pass",
      "message",
      "dimensions",
      "autofitColumnsCalled",
      "widthsChanged",
      "formattingChanged",
      "mergedBefore",
      "mergedAfter",
      "cleanupUnmerged",
      "nonAnchorValueCount",
      "nonAnchorFormulaCount",
    ];

    expect(Object.keys(result).sort()).toEqual([...allowedRunKeys].sort());
    for (const fixture of result.results) {
      expect(Object.keys(fixture).every((key) => allowedFixtureKeys.includes(key))).toBe(true);
    }
    expect(serialized).not.toContain(PRIVATE_CONTENT);
    expect(serialized).not.toContain("Acceptance title");
    expect(serialized).not.toContain("values");
    expect(serialized).not.toContain("formulas");
  });
});
