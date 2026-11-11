import { executeValidatedExcelBridgeToolWithAuthority } from "../officeBridgeExcel";
import { createPrimitiveMutationAuthority } from "../primitiveMutationAuthority";

declare const __HERMES_OFFICE_ACCEPTANCE__: boolean;

export const ACCEPTANCE_ORIGIN = "https://localhost:3000";
export const TOOL_ACCEPTANCE_SHEET = "Tool Acceptance";
export const TOOL_ACCEPTANCE_ADDRESS = "A1:F1";
const TOOL_ACCEPTANCE_COLUMN_COUNT = 6;

const ACCEPTANCE_TITLE = "Acceptance title";
const AUTOFIT_ACCEPTANCE_ANCHOR =
  "Native autofit acceptance anchor with deterministic width change";
const COMPILED_ACCEPTANCE_ENABLED =
  typeof __HERMES_OFFICE_ACCEPTANCE__ === "undefined"
    ? true
    : __HERMES_OFFICE_ACCEPTANCE__;

export type AcceptanceFixtureId =
  | "valid-center-across-selection"
  | "column-autofit-false"
  | "column-autofit-true"
  | "occupied-c1-precondition"
  | "pre-merged-rejection";

export interface AcceptanceFixture {
  readonly id: AcceptanceFixtureId;
  readonly sheet: typeof TOOL_ACCEPTANCE_SHEET;
  readonly address: typeof TOOL_ACCEPTANCE_ADDRESS;
  readonly columnAutofit?: boolean;
}

export interface AcceptanceFixtureEvidence {
  dimensions: { rows: number; columns: number };
  anchorMatchesExpected: boolean;
  nonAnchorValueCount: number;
  nonAnchorFormulaCount: number;
  merged: boolean;
  alignment: string;
  wrapText: boolean;
  columnWidths: readonly number[];
  autofitColumnsCalled?: boolean;
}

export interface AcceptanceFixtureRuntime {
  prepare(): Promise<void>;
  setup(fixture: AcceptanceFixture): Promise<void>;
  readback(fixture: AcceptanceFixture): Promise<AcceptanceFixtureEvidence>;
  cleanup(fixture: AcceptanceFixture): Promise<{ unmerged: boolean }>;
  dispose(): Promise<void>;
}

export interface AcceptanceFixtureResult {
  id: AcceptanceFixtureId;
  ok: boolean;
  pass: boolean;
  message: string;
  dimensions: { rows: number; columns: number };
  autofitColumnsCalled: boolean;
  widthsChanged: boolean;
  formattingChanged: boolean;
  mergedBefore: boolean;
  mergedAfter: boolean;
  cleanupUnmerged: boolean;
  nonAnchorValueCount: number;
  nonAnchorFormulaCount: number;
}

export interface AcceptanceHarnessRunResult {
  sheet: typeof TOOL_ACCEPTANCE_SHEET;
  results: AcceptanceFixtureResult[];
  passCount: number;
  failCount: number;
}

interface AcceptanceWindow extends Record<string, unknown> {
  location?: { href: string };
  __hermesExcelAcceptance?: { run: () => Promise<AcceptanceHarnessRunResult> };
}

interface ParsedToolResult {
  ok: boolean;
}

const FIXTURES: readonly AcceptanceFixture[] = [
  {
    id: "valid-center-across-selection",
    sheet: TOOL_ACCEPTANCE_SHEET,
    address: TOOL_ACCEPTANCE_ADDRESS,
  },
  {
    id: "column-autofit-false",
    sheet: TOOL_ACCEPTANCE_SHEET,
    address: TOOL_ACCEPTANCE_ADDRESS,
    columnAutofit: false,
  },
  {
    id: "column-autofit-true",
    sheet: TOOL_ACCEPTANCE_SHEET,
    address: TOOL_ACCEPTANCE_ADDRESS,
    columnAutofit: true,
  },
  {
    id: "occupied-c1-precondition",
    sheet: TOOL_ACCEPTANCE_SHEET,
    address: TOOL_ACCEPTANCE_ADDRESS,
  },
  {
    id: "pre-merged-rejection",
    sheet: TOOL_ACCEPTANCE_SHEET,
    address: TOOL_ACCEPTANCE_ADDRESS,
  },
];

export function installOfficeAcceptanceHarness(
  runtime: AcceptanceFixtureRuntime = createNativeFixtureRuntime(),
  location = getAcceptanceWindow().location,
): boolean {
  const target = getAcceptanceWindow();
  delete target.__hermesExcelAcceptance;

  if (!COMPILED_ACCEPTANCE_ENABLED || !location || !isAcceptanceLocationAllowed(location)) {
    return false;
  }

  target.__hermesExcelAcceptance = {
    run: () => run(runtime),
  };
  return true;
}

export function isAcceptanceLocationAllowed(location: { href: string }): boolean {
  try {
    const url = new URL(location.href);
    const params = new URLSearchParams(url.search);
    const acceptanceValues = params.getAll("acceptance");
    const hostInfoValues = params.getAll("_host_Info");
    let hasUnknownParameter = false;

    params.forEach((_value, key) => {
      if (key !== "acceptance" && key !== "_host_Info") {
        hasUnknownParameter = true;
      }
    });

    return (
      url.origin === ACCEPTANCE_ORIGIN &&
      url.protocol === "https:" &&
      url.hostname === "localhost" &&
      url.port === "3000" &&
      !hasUnknownParameter &&
      acceptanceValues.length === 1 &&
      acceptanceValues[0] === "1" &&
      (hostInfoValues.length === 0 ||
        (hostInfoValues.length === 1 && hostInfoValues[0] !== ""))
    );
  } catch {
    return false;
  }
}

export async function run(
  runtime: AcceptanceFixtureRuntime = createNativeFixtureRuntime(),
): Promise<AcceptanceHarnessRunResult> {
  const results: AcceptanceFixtureResult[] = [];
  let prepared = false;

  try {
    await runtime.prepare();
    prepared = true;
    for (const fixture of FIXTURES) {
      results.push(await runFixture(runtime, fixture));
    }
  } catch {
    for (const fixture of FIXTURES.slice(results.length)) {
      results.push(fixtureRuntimeFailure(fixture));
    }
  } finally {
    if (prepared) {
      try {
        await runtime.dispose();
      } catch {
        // The result is deliberately content-free; disposal failures stay local.
      }
    }
  }

  const passCount = results.filter((fixture) => fixture.pass).length;
  return {
    sheet: TOOL_ACCEPTANCE_SHEET,
    results,
    passCount,
    failCount: results.length - passCount,
  };
}

async function runFixture(
  runtime: AcceptanceFixtureRuntime,
  fixture: AcceptanceFixture,
): Promise<AcceptanceFixtureResult> {
  let before: AcceptanceFixtureEvidence | undefined;
  let after: AcceptanceFixtureEvidence | undefined;
  let toolResult: ParsedToolResult = { ok: false };
  let result = fixtureRuntimeFailure(fixture);

  try {
    await runtime.setup(fixture);
    before = await runtime.readback(fixture);
    const rawToolResult = await executeValidatedExcelBridgeToolWithAuthority(
      "format_range",
      fixture.id === "column-autofit-true"
        ? {
            sheet: fixture.sheet,
            address: fixture.address,
            format: { columnAutofit: true },
          }
        : {
            sheet: fixture.sheet,
            address: fixture.address,
            format: {
              horizontalAlignment: "CenterAcrossSelection",
              wrapText: false,
              ...(fixture.columnAutofit === undefined
                ? {}
                : { columnAutofit: fixture.columnAutofit }),
            },
          },
      createPrimitiveMutationAuthority({ trustedRequest: null, workbookPlanRolloutEnabled: false }),
    );
    toolResult = parseToolResult(rawToolResult);
    after = await runtime.readback(fixture);
    result = evaluateFixture(fixture, toolResult, before, after, false);
  } catch {
    // Keep the bounded failure result and still run cleanup below.
  }

  try {
    const cleanup = await runtime.cleanup(fixture);
    result = {
      ...result,
      cleanupUnmerged: cleanup.unmerged,
    };
    if (fixture.id === "pre-merged-rejection" && cleanup.unmerged !== true) {
      result = {
        ...result,
        pass: false,
        message: "pre-merged span cleanup was not observed",
      };
    }
  } catch {
    result = {
      ...result,
      pass: false,
      message: "fixture cleanup failed",
    };
  }

  return result;
}

function evaluateFixture(
  fixture: AcceptanceFixture,
  toolResult: ParsedToolResult,
  before: AcceptanceFixtureEvidence,
  after: AcceptanceFixtureEvidence,
  cleanupUnmerged: boolean,
): AcceptanceFixtureResult {
  const widthsChanged = haveColumnWidthsChanged(before.columnWidths, after.columnWidths);
  const formattingChanged =
    before.alignment !== after.alignment || before.wrapText !== after.wrapText;
  const dimensionsUnchanged =
    before.dimensions.rows === after.dimensions.rows &&
    before.dimensions.columns === after.dimensions.columns;
  const contentUnchanged =
    before.anchorMatchesExpected === after.anchorMatchesExpected &&
    before.nonAnchorValueCount === after.nonAnchorValueCount &&
    before.nonAnchorFormulaCount === after.nonAnchorFormulaCount;
  const mergeStateUnchanged = before.merged === after.merged;
  const autofitColumnsCalled =
    after.autofitColumnsCalled === true || widthsChanged;
  const common = {
    id: fixture.id,
    ok: toolResult.ok,
    dimensions: after.dimensions,
    autofitColumnsCalled,
    widthsChanged,
    formattingChanged,
    mergedBefore: before.merged,
    mergedAfter: after.merged,
    cleanupUnmerged,
    nonAnchorValueCount: after.nonAnchorValueCount,
    nonAnchorFormulaCount: after.nonAnchorFormulaCount,
  };

  if (fixture.id === "column-autofit-true") {
    const pass =
      toolResult.ok &&
      dimensionsUnchanged &&
      contentUnchanged &&
      mergeStateUnchanged &&
      before.merged === false &&
      after.merged === false &&
      widthsChanged &&
      !formattingChanged;
    return {
      ...common,
      pass,
      message: pass
        ? "column autofit applied without unrelated formatting mutation"
        : "column autofit evidence failed",
    };
  }

  if (fixture.id === "valid-center-across-selection" || fixture.id === "column-autofit-false") {
    const pass =
      toolResult.ok &&
      after.anchorMatchesExpected &&
      after.nonAnchorValueCount === 0 &&
      after.nonAnchorFormulaCount === 0 &&
      after.merged === false &&
      after.alignment === "CenterAcrossSelection" &&
      after.wrapText === false &&
      !widthsChanged;
    return {
      ...common,
      pass,
      message: pass
        ? "CenterAcrossSelection applied without changing column widths"
        : "valid CenterAcrossSelection evidence failed",
    };
  }

  if (fixture.id === "occupied-c1-precondition") {
    const pass =
      !toolResult.ok &&
      after.nonAnchorValueCount === 1 &&
      before.alignment === after.alignment &&
      before.wrapText === after.wrapText &&
      !widthsChanged &&
      before.merged === after.merged;
    return {
      ...common,
      pass,
      message: pass
        ? "occupied C1 precondition rejected without mutation"
        : "occupied C1 precondition evidence failed",
    };
  }

  const pass =
    !toolResult.ok &&
    before.merged &&
    after.merged &&
    before.alignment === after.alignment &&
    before.wrapText === after.wrapText &&
    !widthsChanged;
  return {
    ...common,
    pass,
    message: pass
      ? "pre-merged span rejected without mutation"
      : "pre-merged rejection evidence failed",
  };
}

function parseToolResult(raw: string): ParsedToolResult {
  try {
    const value: unknown = JSON.parse(raw);
    return {
      ok: typeof value === "object" && value !== null && (value as { ok?: unknown }).ok === true,
    };
  } catch {
    return { ok: false };
  }
}

function haveColumnWidthsChanged(
  before: readonly number[],
  after: readonly number[],
): boolean {
  if (
    before.length !== TOOL_ACCEPTANCE_COLUMN_COUNT ||
    after.length !== TOOL_ACCEPTANCE_COLUMN_COUNT
  ) {
    return true;
  }
  return after.some((width, index) => width !== before[index]);
}

function fixtureRuntimeFailure(fixture: AcceptanceFixture): AcceptanceFixtureResult {
  return {
    id: fixture.id,
    ok: false,
    pass: false,
    message: "fixture runtime unavailable",
    dimensions: { rows: 0, columns: 0 },
    autofitColumnsCalled: false,
    widthsChanged: false,
    formattingChanged: false,
    mergedBefore: false,
    mergedAfter: false,
    cleanupUnmerged: false,
    nonAnchorValueCount: 0,
    nonAnchorFormulaCount: 0,
  };
}

function getAcceptanceWindow(): AcceptanceWindow {
  const root = globalThis as unknown as { window?: AcceptanceWindow };
  return root.window ?? (globalThis as unknown as AcceptanceWindow);
}

function createNativeFixtureRuntime(): AcceptanceFixtureRuntime {
  let prepared = false;

  return {
    async prepare() {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.add(TOOL_ACCEPTANCE_SHEET);
        sheet.activate();
        await context.sync();
      });
      prepared = true;
    },
    async setup(fixture) {
      assertDisposableFixture(fixture);
      await Excel.run(async (context) => {
        const range = getNativeFixtureRange(context);
        const anchorContent =
          fixture.id === "column-autofit-true" ? AUTOFIT_ACCEPTANCE_ANCHOR : ACCEPTANCE_TITLE;
        range.values = [
          [
            anchorContent,
            null,
            fixture.id === "occupied-c1-precondition" ? "occupied" : null,
            null,
            null,
            null,
          ],
        ];
        range.format.horizontalAlignment = "General";
        range.format.wrapText = true;
        range.format.columnWidth = 8;
        if (fixture.id === "pre-merged-rejection") {
          range.merge(false);
        }
        await context.sync();
      });
    },
    async readback(fixture) {
      assertDisposableFixture(fixture);
      return Excel.run(async (context) => {
        const range = getNativeFixtureRange(context);
        const columnRanges = Array.from(
          { length: TOOL_ACCEPTANCE_COLUMN_COUNT },
          (_value, columnIndex) => range.getCell(0, columnIndex),
        );
        const merged = range.getMergedAreasOrNullObject();
        range.load(["rowCount", "columnCount", "values", "formulas"]);
        merged.load("isNullObject");
        range.format.load(["horizontalAlignment", "wrapText"]);
        for (const columnRange of columnRanges) {
          columnRange.format.load("columnWidth");
        }
        await context.sync();

        const values = (range.values ?? []) as unknown[][];
        const formulas = (range.formulas ?? []) as unknown[][];
        return {
          dimensions: { rows: range.rowCount, columns: range.columnCount },
          anchorMatchesExpected:
            values[0]?.[0] ===
            (fixture.id === "column-autofit-true" ? AUTOFIT_ACCEPTANCE_ANCHOR : ACCEPTANCE_TITLE),
          nonAnchorValueCount: countNonBlank(values[0]?.slice(1)),
          nonAnchorFormulaCount: countNonBlank(formulas[0]?.slice(1)),
          merged: !merged.isNullObject,
          alignment: String(range.format.horizontalAlignment),
          wrapText: range.format.wrapText === true,
          columnWidths: columnRanges.map((columnRange) =>
            typeof columnRange.format.columnWidth === "number" ? columnRange.format.columnWidth : 0,
          ),
        };
      });
    },
    async cleanup(fixture) {
      assertDisposableFixture(fixture);
      if (fixture.id !== "pre-merged-rejection") {
        return { unmerged: false };
      }
      await Excel.run(async (context) => {
        const range = getNativeFixtureRange(context);
        range.unmerge();
        await context.sync();
      });
      return { unmerged: true };
    },
    async dispose() {
      if (!prepared) {
        return;
      }
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(TOOL_ACCEPTANCE_SHEET);
        sheet.delete();
        await context.sync();
      });
      prepared = false;
    },
  };
}

function getNativeFixtureRange(context: Excel.RequestContext): Excel.Range {
  const sheet = context.workbook.worksheets.getItem(TOOL_ACCEPTANCE_SHEET);
  return sheet.getRange(TOOL_ACCEPTANCE_ADDRESS);
}

function assertDisposableFixture(fixture: AcceptanceFixture): void {
  if (
    fixture.sheet !== TOOL_ACCEPTANCE_SHEET ||
    fixture.address !== TOOL_ACCEPTANCE_ADDRESS
  ) {
    throw new Error("invalid acceptance fixture target");
  }
}

function countNonBlank(values: unknown[] | undefined): number {
  return (values ?? []).filter((value) => value !== null && value !== undefined && value !== "").length;
}
