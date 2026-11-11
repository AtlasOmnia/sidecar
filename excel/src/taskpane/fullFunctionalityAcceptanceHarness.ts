import { excelTools } from "../excelTools";
import type { ExcelToolDefinition } from "../excelTools";

declare const __HERMES_OFFICE_ACCEPTANCE__: boolean;

export const FULL_FUNCTIONALITY_ACCEPTANCE_ORIGIN = "https://localhost:3000";
export const FULL_FUNCTIONALITY_ACCEPTANCE_PATH = "/taskpane.html";
export const FULL_FUNCTIONALITY_ACCEPTANCE_QUERY_KEY = "fullFunctionalityAcceptance";
export const FULL_FUNCTIONALITY_ACCEPTANCE_QUERY_VALUE = "1";
export const FULL_FUNCTIONALITY_ACCEPTANCE_URL =
  "https://localhost:3000/taskpane.html?fullFunctionalityAcceptance=1";
export const FULL_FUNCTIONALITY_SHEET_PREFIX = "Hermes_Acceptance_";
export const FULL_FUNCTIONALITY_EVIDENCE_MAX_BYTES = 64 * 1024;

export type FullFunctionalityAcceptanceToolName =
  ExcelToolDefinition["schema"]["function"]["name"];
export type FullFunctionalityAcceptanceCaseKind = "positive" | "negative" | "boundary";
export type FullFunctionalityAcceptanceExpectedOutcome =
  | "success"
  | "rejected-without-mutation"
  | "unsupported-without-mutation";
export type FullFunctionalityAcceptanceEvidenceStatus =
  | "pending"
  | "passed"
  | "failed"
  | "unsupported";
export type FullFunctionalityAcceptanceFixtureRole =
  | "data"
  | "formats"
  | "tables"
  | "pivot"
  | "comments"
  | "protected"
  | "expectedFailures";

export type FullFunctionalityAcceptancePrerequisite =
  | {
      readonly kind: "create-sheet";
      readonly name: string;
    }
  | {
      readonly kind: "seed-comment";
      readonly sheet: string;
      readonly address: string;
      readonly text: string;
    };

export interface FullFunctionalityAcceptanceSheetSet {
  readonly runId: string;
  readonly data: string;
  readonly formats: string;
  readonly tables: string;
  readonly pivot: string;
  readonly comments: string;
  readonly protected: string;
  readonly expectedFailures: string;
}

export interface FullFunctionalityAcceptanceCaseDefinition {
  readonly id: string;
  readonly tool: FullFunctionalityAcceptanceToolName;
  readonly kind: FullFunctionalityAcceptanceCaseKind;
  readonly fixture: FullFunctionalityAcceptanceFixtureRole;
  readonly mutates: boolean;
  readonly expectedOutcome: FullFunctionalityAcceptanceExpectedOutcome;
  readonly sentinelRequired: boolean;
  readonly resetFixture: true;
  readonly allowUnsupported: boolean;
  readonly buildArguments: (
    sheets: FullFunctionalityAcceptanceSheetSet,
  ) => Record<string, unknown>;
  readonly buildPrerequisites: (
    sheets: FullFunctionalityAcceptanceSheetSet,
  ) => readonly FullFunctionalityAcceptancePrerequisite[];
}

export interface FullFunctionalityAcceptancePhaseEvidence {
  readonly status: FullFunctionalityAcceptanceEvidenceStatus;
  readonly captured: boolean;
  readonly matchesExpectation: boolean;
}

export interface FullFunctionalityAcceptanceToolResultEvidence {
  readonly status: FullFunctionalityAcceptanceEvidenceStatus;
  readonly ok: boolean;
  readonly outcome:
    | "success"
    | "expected-error"
    | "unsupported"
    | "unexpected-error"
    | "malformed";
}

export interface FullFunctionalityAcceptanceCaseEvidence {
  readonly id: string;
  readonly tool: FullFunctionalityAcceptanceToolName;
  readonly kind: FullFunctionalityAcceptanceCaseKind;
  readonly pass: boolean;
  readonly pre: FullFunctionalityAcceptancePhaseEvidence;
  readonly result: FullFunctionalityAcceptanceToolResultEvidence;
  readonly post: FullFunctionalityAcceptancePhaseEvidence;
  readonly sentinel: FullFunctionalityAcceptancePhaseEvidence;
  readonly noPartialMutation: boolean;
}

export interface FullFunctionalityAcceptanceDeferredSlot {
  readonly status: "pending";
  readonly nativeRequired: true;
  readonly resultCount: 0;
}

export interface FullFunctionalityAcceptanceEvidence {
  readonly schemaVersion: 1;
  readonly runId: string;
  readonly pass: boolean;
  readonly cases: readonly FullFunctionalityAcceptanceCaseEvidence[];
  readonly passCount: number;
  readonly failCount: number;
  readonly cleanup: {
    readonly attempted: boolean;
    readonly removedCount: number;
    readonly pass: boolean;
  };
  readonly futureNative: {
    readonly formula: FullFunctionalityAcceptanceDeferredSlot;
    readonly workflow: FullFunctionalityAcceptanceDeferredSlot;
    readonly isolation: FullFunctionalityAcceptanceDeferredSlot;
    readonly recovery: FullFunctionalityAcceptanceDeferredSlot;
  };
}

export interface FullFunctionalityAcceptanceWindow extends Record<string, unknown> {
  location?: { href: string };
  __hermesExcelFullFunctionalityAcceptance?: {
    run: () => Promise<FullFunctionalityAcceptanceEvidence>;
  };
}

const COMPILED_ACCEPTANCE_ENABLED =
  typeof __HERMES_OFFICE_ACCEPTANCE__ !== "undefined" &&
  __HERMES_OFFICE_ACCEPTANCE__ === true;

const FULL_FUNCTIONALITY_ROLE_SUFFIX: Readonly<
  Record<FullFunctionalityAcceptanceFixtureRole, string>
> = {
  data: "D",
  formats: "F",
  tables: "T",
  pivot: "P",
  comments: "C",
  protected: "R",
  expectedFailures: "X",
};

const CASE_TOOL_NAMES = excelTools.map(
  ({ schema }) => schema.function.name,
) as FullFunctionalityAcceptanceToolName[];

export const FULL_FUNCTIONALITY_TOOL_NAMES: readonly FullFunctionalityAcceptanceToolName[] =
  Object.freeze([...CASE_TOOL_NAMES]);

function defineCase(
  tool: FullFunctionalityAcceptanceToolName,
  kind: FullFunctionalityAcceptanceCaseKind,
  fixture: FullFunctionalityAcceptanceFixtureRole,
  mutates: boolean,
  expectedOutcome: FullFunctionalityAcceptanceExpectedOutcome,
  buildArguments: FullFunctionalityAcceptanceCaseDefinition["buildArguments"],
  options: {
    readonly suffix?: string;
    readonly allowUnsupported?: boolean;
    readonly buildPrerequisites?: FullFunctionalityAcceptanceCaseDefinition["buildPrerequisites"];
  } = {},
): FullFunctionalityAcceptanceCaseDefinition {
  return {
    id: `${tool}:${kind}${options.suffix ? `:${options.suffix}` : ""}`,
    tool,
    kind,
    fixture,
    mutates,
    expectedOutcome,
    sentinelRequired: mutates,
    resetFixture: true,
    allowUnsupported: options.allowUnsupported === true,
    buildArguments,
    buildPrerequisites: options.buildPrerequisites ?? (() => []),
  };
}

const tableName = (sheets: FullFunctionalityAcceptanceSheetSet): string =>
  `${sheets.tables}_Table`;
const temporarySheet = (
  sheets: FullFunctionalityAcceptanceSheetSet,
  suffix: string,
): string => `${sheets.expectedFailures}_${suffix}`;

/*
 * This is the reusable Phase C registry. B2/B3 deliberately do not execute
 * these cases; they only create and read back the fixtures that later native
 * lanes will use. Every argument that names a worksheet or table is derived
 * from the run-scoped sheet set.
 */
export const FULL_FUNCTIONALITY_ACCEPTANCE_CASES: readonly FullFunctionalityAcceptanceCaseDefinition[] = [
  defineCase("get_workbook_overview", "positive", "data", false, "success", () => ({})),
  defineCase("get_workbook_overview", "boundary", "expectedFailures", false, "success", () => ({})),

  defineCase("read_range", "positive", "data", false, "success", (sheets) => ({
    sheet: sheets.data,
    address: "A1:H8",
  })),
  defineCase("read_range", "boundary", "expectedFailures", false, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.expectedFailures,
    address: "not-an-address",
  })),

  defineCase("summarize_range_stats", "positive", "data", false, "success", (sheets) => ({
    sheet: sheets.data,
    address: "C2:C8",
  })),
  defineCase("summarize_range_stats", "boundary", "expectedFailures", false, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.expectedFailures,
    address: "A0:A1",
  })),

  defineCase("write_range", "positive", "data", true, "success", (sheets) => ({
    sheet: sheets.data,
    address: "F2:F3",
    values: [["B2_WRITE_VALUE"], ["B3_WRITE_VALUE"]],
  })),
  defineCase("write_range", "negative", "expectedFailures", true, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.expectedFailures,
    address: "A1:B2",
    values: [["DIMENSION_MISMATCH"]],
  })),

  defineCase("format_range", "positive", "formats", true, "success", (sheets) => ({
    sheet: sheets.formats,
    address: "A1:F2",
    format: {
      bold: true,
      fillColor: "#1F4E78",
      fontColor: "#FFFFFF",
      numberFormat: "0.00",
      horizontalAlignment: "Center",
      verticalAlignment: "Center",
      wrapText: true,
      borders: true,
      columnAutofit: false,
      rowAutofit: false,
    },
  })),
  defineCase("format_range", "negative", "expectedFailures", true, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.expectedFailures,
    address: "A1:B1",
    format: { horizontalAlignment: "NotAnExcelAlignment" },
  })),

  defineCase("conditional_format_range", "positive", "data", true, "success", (sheets) => ({
    sheet: sheets.data,
    address: "C2:C8",
    ruleType: "cellValue",
    operator: "GreaterThan",
    formula1: "1000",
    fillColor: "#FFF2CC",
  })),
  defineCase("conditional_format_range", "negative", "expectedFailures", true, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.expectedFailures,
    address: "A1:B1",
    ruleType: "cellValue",
    operator: "NotAnOperator",
    formula1: "1000",
  })),

  defineCase("freeze_panes", "positive", "formats", true, "success", (sheets) => ({
    sheet: sheets.formats,
    rows: 1,
    columns: 1,
  })),
  defineCase("freeze_panes", "boundary", "expectedFailures", true, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.expectedFailures,
    rows: -1,
    columns: 0,
  })),

  defineCase("set_data_validation", "positive", "data", true, "success", (sheets) => ({
    sheet: sheets.data,
    address: "F2:F8",
    type: "list",
    source: "Open,Closed,Pending",
    promptTitle: "Synthetic status",
    promptMessage: "Choose a fixed fixture status.",
  })),
  defineCase("set_data_validation", "negative", "expectedFailures", true, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.expectedFailures,
    address: "A1:B1",
    type: "list",
    source: "",
  })),

  defineCase("create_sheet", "positive", "expectedFailures", true, "success", (sheets) => ({
    name: temporarySheet(sheets, "CS"),
  })),
  defineCase("create_sheet", "negative", "expectedFailures", true, "rejected-without-mutation", (sheets) => ({
    name: sheets.data,
  })),

  defineCase(
    "delete_sheet",
    "positive",
    "expectedFailures",
    true,
    "success",
    (sheets) => ({ name: temporarySheet(sheets, "DS") }),
    {
      buildPrerequisites: (sheets) => [
        { kind: "create-sheet", name: temporarySheet(sheets, "DS") },
      ],
    },
  ),
  defineCase("delete_sheet", "negative", "expectedFailures", true, "rejected-without-mutation", (sheets) => ({
    name: temporarySheet(sheets, "M"),
  })),

  defineCase("activate_sheet", "positive", "formats", true, "success", (sheets) => ({
    name: sheets.formats,
  })),
  defineCase("activate_sheet", "negative", "expectedFailures", true, "rejected-without-mutation", (sheets) => ({
    name: temporarySheet(sheets, "M"),
  })),

  defineCase("insert_table", "positive", "tables", true, "success", (sheets) => ({
    sheet: sheets.tables,
    address: "H1:K4",
    hasHeaders: true,
  })),
  defineCase("insert_table", "negative", "tables", true, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.tables,
    address: "A1:D4",
    hasHeaders: true,
  })),

  defineCase("sort_table", "positive", "tables", true, "success", (sheets) => ({
    tableName: tableName(sheets),
    columnIndex: 2,
    ascending: false,
  })),
  defineCase("sort_table", "boundary", "tables", true, "rejected-without-mutation", (sheets) => ({
    tableName: tableName(sheets),
    columnIndex: 99,
    ascending: true,
  })),

  defineCase("filter_table", "positive", "tables", true, "success", (sheets) => ({
    tableName: tableName(sheets),
    columnIndex: 0,
    values: ["North", "South"],
  })),
  defineCase("filter_table", "negative", "tables", true, "rejected-without-mutation", (sheets) => ({
    tableName: tableName(sheets),
    columnIndex: 99,
    criterion1: ">1000",
  })),

  defineCase("create_chart", "positive", "data", true, "success", (sheets) => ({
    sheet: sheets.data,
    dataRange: "B1:C8",
    chartType: "ColumnClustered",
    title: "Synthetic acceptance chart",
  })),
  defineCase("create_chart", "negative", "expectedFailures", true, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.expectedFailures,
    dataRange: "A1:B2",
    chartType: "NotAnExcelChart",
    title: "Invalid synthetic chart",
  })),

  defineCase("find_replace", "positive", "data", true, "success", (sheets) => ({
    sheet: sheets.data,
    find: "North",
    replace: "North-East",
    matchCase: true,
  })),
  defineCase("find_replace", "boundary", "expectedFailures", true, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.expectedFailures,
    find: "",
    replace: "NOOP",
    matchCase: false,
  })),

  defineCase("clear_range", "positive", "formats", true, "success", (sheets) => ({
    sheet: sheets.formats,
    address: "F3:F5",
  })),
  defineCase("clear_range", "negative", "protected", true, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.protected,
    address: "A1:B2",
  })),

  defineCase("autofill_formula", "positive", "data", true, "success", (sheets) => ({
    sheet: sheets.data,
    sourceAddress: "G2",
    targetAddress: "G2:G8",
  })),
  defineCase("autofill_formula", "negative", "expectedFailures", true, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.expectedFailures,
    sourceAddress: "A1:B1",
    targetAddress: "A1:B4",
  })),

  defineCase(
    "rename_sheet",
    "positive",
    "expectedFailures",
    true,
    "success",
    (sheets) => ({
      oldName: temporarySheet(sheets, "RS"),
      newName: temporarySheet(sheets, "RD"),
    }),
    {
      buildPrerequisites: (sheets) => [
        { kind: "create-sheet", name: temporarySheet(sheets, "RS") },
      ],
    },
  ),
  defineCase("rename_sheet", "negative", "expectedFailures", true, "rejected-without-mutation", (sheets) => ({
    oldName: sheets.expectedFailures,
    newName: sheets.data,
  })),

  defineCase("copy_sheet", "positive", "data", true, "success", (sheets) => ({
    sourceName: sheets.data,
    newName: temporarySheet(sheets, "CP"),
  })),
  defineCase("copy_sheet", "negative", "data", true, "rejected-without-mutation", (sheets) => ({
    sourceName: sheets.data,
    newName: sheets.data,
  })),

  defineCase("protect_sheet", "positive", "formats", true, "success", (sheets) => ({
    sheet: sheets.formats,
    allowSort: true,
    allowFilter: true,
  })),
  defineCase("protect_sheet", "negative", "protected", true, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.protected,
  })),

  defineCase("unprotect_sheet", "positive", "protected", true, "success", (sheets) => ({
    sheet: sheets.protected,
  })),
  defineCase("unprotect_sheet", "boundary", "formats", true, "success", (sheets) => ({
    sheet: sheets.formats,
  })),

  defineCase(
    "create_pivot_table",
    "positive",
    "pivot",
    true,
    "success",
    (sheets) => ({
      sheet: sheets.pivot,
      address: "A1",
      sourceSheet: sheets.tables,
      sourceRange: "A1:D21",
      rowFields: ["Region"],
      dataFields: [{ name: "Amount", summarizeBy: "sum" }],
    }),
    { allowUnsupported: true },
  ),
  defineCase(
    "create_pivot_table",
    "negative",
    "pivot",
    true,
    "rejected-without-mutation",
    (sheets) => ({
      sheet: sheets.pivot,
      address: "A1",
      sourceSheet: sheets.expectedFailures,
      sourceRange: "A1:D4",
      rowFields: ["MissingField"],
      dataFields: [{ name: "MissingAmount", summarizeBy: "sum" }],
    }),
    { allowUnsupported: true },
  ),

  defineCase("run_calculation", "positive", "data", true, "success", () => ({})),
  defineCase("run_calculation", "boundary", "expectedFailures", true, "success", () => ({})),

  defineCase(
    "add_comment",
    "positive",
    "comments",
    true,
    "success",
    (sheets) => ({
      sheet: sheets.comments,
      address: "B2",
      text: "Synthetic review note",
    }),
    { allowUnsupported: true },
  ),
  defineCase(
    "add_comment",
    "positive",
    "comments",
    true,
    "success",
    (sheets) => ({
      sheet: sheets.comments,
      address: "C3",
      text: "Synthetic updated review note",
    }),
    {
      suffix: "update",
      allowUnsupported: true,
      buildPrerequisites: (sheets) => [
        {
          kind: "seed-comment",
          sheet: sheets.comments,
          address: "C3",
          text: "Synthetic existing review note",
        },
      ],
    },
  ),
  defineCase(
    "add_comment",
    "negative",
    "expectedFailures",
    true,
    "rejected-without-mutation",
    (sheets) => ({
      sheet: sheets.expectedFailures,
      address: "A1:B2",
      text: "Synthetic invalid comment target",
    }),
    { allowUnsupported: true },
  ),

  defineCase(
    "clear_comments",
    "positive",
    "comments",
    true,
    "success",
    (sheets) => ({ sheet: sheets.comments, address: "C3" }),
    {
      allowUnsupported: true,
      buildPrerequisites: (sheets) => [
        {
          kind: "seed-comment",
          sheet: sheets.comments,
          address: "C3",
          text: "Synthetic clear target",
        },
        {
          kind: "seed-comment",
          sheet: sheets.comments,
          address: "D4",
          text: "Synthetic outside clear target",
        },
      ],
    },
  ),
  defineCase(
    "clear_comments",
    "negative",
    "expectedFailures",
    true,
    "rejected-without-mutation",
    (sheets) => ({ sheet: sheets.expectedFailures, address: "A0:A1" }),
    { allowUnsupported: true },
  ),

  defineCase("import_csv_to_sheet", "positive", "data", true, "success", (sheets) => ({
    sheet: sheets.data,
    address: "K1",
    csv: "Name,Count\nAlpha,2\nBeta,3",
    hasHeaders: true,
  })),
  defineCase("import_csv_to_sheet", "negative", "expectedFailures", true, "rejected-without-mutation", (sheets) => ({
    sheet: sheets.expectedFailures,
    address: "A1",
    csv: "Name,\"unterminated",
    hasHeaders: true,
  })),
];

const VALID_RUN_ID = /^[a-z0-9]{8}$/;
const VALID_EVIDENCE_STATUSES: readonly FullFunctionalityAcceptanceEvidenceStatus[] = [
  "pending",
  "passed",
  "failed",
  "unsupported",
];
const VALID_RESULT_OUTCOMES: readonly FullFunctionalityAcceptanceToolResultEvidence["outcome"][] = [
  "success",
  "expected-error",
  "unsupported",
  "unexpected-error",
  "malformed",
];

export function isFullFunctionalityAcceptanceLocationAllowed(location: { href: string }): boolean {
  try {
    const url = new URL(location.href);
    if (
      url.origin !== FULL_FUNCTIONALITY_ACCEPTANCE_ORIGIN ||
      url.protocol !== "https:" ||
      url.hostname !== "localhost" ||
      url.port !== "3000" ||
      url.pathname !== FULL_FUNCTIONALITY_ACCEPTANCE_PATH ||
      url.username !== "" ||
      url.password !== "" ||
      url.hash !== ""
    ) {
      return false;
    }

    const params = new URLSearchParams(url.search);
    const acceptanceValues = params.getAll(FULL_FUNCTIONALITY_ACCEPTANCE_QUERY_KEY);
    const hostInfoValues = params.getAll("_host_Info");
    let unknownParameter = false;
    params.forEach((_value, key) => {
      if (key !== FULL_FUNCTIONALITY_ACCEPTANCE_QUERY_KEY && key !== "_host_Info") {
        unknownParameter = true;
      }
    });

    return (
      !unknownParameter &&
      acceptanceValues.length === 1 &&
      acceptanceValues[0] === FULL_FUNCTIONALITY_ACCEPTANCE_QUERY_VALUE &&
      (hostInfoValues.length === 0 ||
        (hostInfoValues.length === 1 && hostInfoValues[0].length > 0))
    );
  } catch {
    return false;
  }
}

export function createFullFunctionalityAcceptanceRunId(): string {
  const randomUuid = globalThis.crypto?.randomUUID;
  if (typeof randomUuid !== "function") {
    throw new Error("secure acceptance run ID generation is unavailable");
  }
  const runId = randomUuid.call(globalThis.crypto).replace(/-/g, "").slice(0, 8).toLowerCase();
  if (!VALID_RUN_ID.test(runId)) {
    throw new Error("generated acceptance run ID is invalid");
  }
  return runId;
}

export function createFullFunctionalityAcceptanceSheetSet(
  runId: string,
): FullFunctionalityAcceptanceSheetSet {
  if (!VALID_RUN_ID.test(runId)) {
    throw new Error("invalid acceptance run ID");
  }

  const makeName = (role: FullFunctionalityAcceptanceFixtureRole): string => {
    const name = `${FULL_FUNCTIONALITY_SHEET_PREFIX}${runId}_${FULL_FUNCTIONALITY_ROLE_SUFFIX[role]}`;
    if (name.length > 31 || /[\\/?*:[\]]/.test(name)) {
      throw new Error("generated acceptance run sheet name is invalid");
    }
    return name;
  };

  return {
    runId,
    data: makeName("data"),
    formats: makeName("formats"),
    tables: makeName("tables"),
    pivot: makeName("pivot"),
    comments: makeName("comments"),
    protected: makeName("protected"),
    expectedFailures: makeName("expectedFailures"),
  };
}

function pendingPhase(): FullFunctionalityAcceptancePhaseEvidence {
  return { status: "pending", captured: false, matchesExpectation: false };
}

function passedPhase(): FullFunctionalityAcceptancePhaseEvidence {
  return { status: "passed", captured: true, matchesExpectation: true };
}

function failedPhase(): FullFunctionalityAcceptancePhaseEvidence {
  return { status: "failed", captured: false, matchesExpectation: false };
}

function pendingToolResult(): FullFunctionalityAcceptanceToolResultEvidence {
  return { status: "pending", ok: false, outcome: "malformed" };
}

function futureNativeSlot(): FullFunctionalityAcceptanceDeferredSlot {
  return { status: "pending", nativeRequired: true, resultCount: 0 };
}

function pendingFutureNative(): FullFunctionalityAcceptanceEvidence["futureNative"] {
  return {
    formula: futureNativeSlot(),
    workflow: futureNativeSlot(),
    isolation: futureNativeSlot(),
    recovery: futureNativeSlot(),
  };
}

export function createPendingFullFunctionalityAcceptanceEvidence(
  runId: string,
): FullFunctionalityAcceptanceEvidence {
  createFullFunctionalityAcceptanceSheetSet(runId);
  const cases = FULL_FUNCTIONALITY_ACCEPTANCE_CASES.map((definition) => ({
    id: definition.id,
    tool: definition.tool,
    kind: definition.kind,
    pass: false,
    pre: pendingPhase(),
    result: pendingToolResult(),
    post: pendingPhase(),
    sentinel: pendingPhase(),
    noPartialMutation: false,
  }));

  return {
    schemaVersion: 1,
    runId,
    pass: false,
    cases,
    passCount: 0,
    failCount: cases.length,
    cleanup: { attempted: false, removedCount: 0, pass: false },
    futureNative: pendingFutureNative(),
  };
}

function phaseMatches(evidence: FullFunctionalityAcceptancePhaseEvidence): boolean {
  return evidence.status === "passed" && evidence.captured && evidence.matchesExpectation;
}

function resultMatches(
  definition: FullFunctionalityAcceptanceCaseDefinition,
  result: FullFunctionalityAcceptanceToolResultEvidence,
): boolean {
  if (definition.expectedOutcome === "success") {
    if (result.status === "passed" && result.ok && result.outcome === "success") {
      return true;
    }
  } else if (definition.expectedOutcome === "rejected-without-mutation") {
    if (result.status === "passed" && !result.ok && result.outcome === "expected-error") {
      return true;
    }
  } else if (
    definition.allowUnsupported &&
    (result.status === "passed" || result.status === "unsupported") &&
    !result.ok &&
    result.outcome === "unsupported"
  ) {
    return true;
  }
  return (
    definition.allowUnsupported &&
    (result.status === "passed" || result.status === "unsupported") &&
    !result.ok &&
    result.outcome === "unsupported"
  );
}

export function evaluateFullFunctionalityAcceptanceCase(
  definition: FullFunctionalityAcceptanceCaseDefinition,
  evidence: Omit<
    FullFunctionalityAcceptanceCaseEvidence,
    "id" | "tool" | "kind" | "pass" | "noPartialMutation"
  >,
): FullFunctionalityAcceptanceCaseEvidence {
  const preMatches = phaseMatches(evidence.pre);
  const postMatches = phaseMatches(evidence.post);
  const sentinelMatches = phaseMatches(evidence.sentinel);
  const noPartialMutation =
    preMatches &&
    postMatches &&
    sentinelMatches &&
    (definition.expectedOutcome === "success" || !evidence.result.ok);
  const pass =
    noPartialMutation &&
    resultMatches(definition, evidence.result);

  return {
    id: definition.id,
    tool: definition.tool,
    kind: definition.kind,
    pass,
    ...evidence,
    noPartialMutation,
  };
}

function safeEvidenceStatus(
  value: unknown,
): FullFunctionalityAcceptanceEvidenceStatus {
  return VALID_EVIDENCE_STATUSES.includes(value as FullFunctionalityAcceptanceEvidenceStatus)
    ? (value as FullFunctionalityAcceptanceEvidenceStatus)
    : "pending";
}

function safeOutcome(
  value: unknown,
): FullFunctionalityAcceptanceToolResultEvidence["outcome"] {
  return VALID_RESULT_OUTCOMES.includes(value as FullFunctionalityAcceptanceToolResultEvidence["outcome"])
    ? (value as FullFunctionalityAcceptanceToolResultEvidence["outcome"])
    : "malformed";
}

function safePhase(value: unknown): FullFunctionalityAcceptancePhaseEvidence {
  const phase = value as Partial<FullFunctionalityAcceptancePhaseEvidence> | null;
  const status = safeEvidenceStatus(phase?.status);
  return {
    status,
    captured: phase?.captured === true,
    matchesExpectation: phase?.matchesExpectation === true,
  };
}

function safeResult(value: unknown): FullFunctionalityAcceptanceToolResultEvidence {
  const result = value as Partial<FullFunctionalityAcceptanceToolResultEvidence> | null;
  return {
    status: safeEvidenceStatus(result?.status),
    ok: result?.ok === true,
    outcome: safeOutcome(result?.outcome),
  };
}

export function serializeFullFunctionalityAcceptanceEvidence(
  evidence: FullFunctionalityAcceptanceEvidence,
): string {
  const cases = FULL_FUNCTIONALITY_ACCEPTANCE_CASES.map((definition, index) => {
    const candidate = evidence.cases[index];
    const pre = safePhase(candidate?.pre);
    const result = safeResult(candidate?.result);
    const post = safePhase(candidate?.post);
    const sentinel = safePhase(candidate?.sentinel);
    const noPartialMutation =
      phaseMatches(pre) &&
      phaseMatches(post) &&
      phaseMatches(sentinel) &&
      (definition.expectedOutcome === "success" || !result.ok);
    return {
      id: definition.id,
      tool: definition.tool,
      kind: definition.kind,
      pass:
        candidate?.pass === true &&
        definition.id === candidate.id &&
        noPartialMutation &&
        resultMatches(definition, result),
      pre,
      result,
      post,
      sentinel,
      noPartialMutation,
    };
  });
  const passCount = cases.filter((candidate) => candidate.pass).length;
  const cleanup = {
    attempted: evidence.cleanup?.attempted === true,
    removedCount:
      typeof evidence.cleanup?.removedCount === "number" &&
      Number.isSafeInteger(evidence.cleanup.removedCount) &&
      evidence.cleanup.removedCount >= 0
        ? evidence.cleanup.removedCount
        : 0,
    pass: evidence.cleanup?.pass === true,
  };
  const bounded = {
    schemaVersion: 1 as const,
    runId: VALID_RUN_ID.test(evidence.runId) ? evidence.runId : "invalid",
    pass:
      cases.length === FULL_FUNCTIONALITY_ACCEPTANCE_CASES.length &&
      passCount === cases.length &&
      cleanup.attempted &&
      cleanup.pass,
    cases,
    passCount,
    failCount: cases.length - passCount,
    cleanup,
    futureNative: pendingFutureNative(),
  };
  const serialized = JSON.stringify(bounded);
  if (new TextEncoder().encode(serialized).byteLength > FULL_FUNCTIONALITY_EVIDENCE_MAX_BYTES) {
    throw new Error("full-functionality acceptance evidence exceeds the 64 KiB limit");
  }
  return serialized;
}

function getAcceptanceWindow(): FullFunctionalityAcceptanceWindow {
  const root = globalThis as unknown as { window?: FullFunctionalityAcceptanceWindow };
  return root.window ?? (globalThis as unknown as FullFunctionalityAcceptanceWindow);
}

export function installFullFunctionalityAcceptanceHarness(
  location?: { href: string },
): boolean {
  const target = getAcceptanceWindow();
  delete target.__hermesExcelFullFunctionalityAcceptance;
  const candidateLocation = location ?? target.location;
  if (
    !COMPILED_ACCEPTANCE_ENABLED ||
    !candidateLocation ||
    !isFullFunctionalityAcceptanceLocationAllowed(candidateLocation)
  ) {
    return false;
  }
  target.__hermesExcelFullFunctionalityAcceptance = {
    run: runFullFunctionalityAcceptance,
  };
  return true;
}

type FixtureCell = string | number | boolean | null;
type FixtureRole = FullFunctionalityAcceptanceFixtureRole;

interface FixtureSeed {
  readonly role: FixtureRole;
  readonly sheet: string;
  readonly address: string;
  readonly grid: readonly (readonly FixtureCell[])[];
  readonly sentinelAddress: string;
  readonly sentinelValues: readonly (readonly FixtureCell[])[];
  readonly requiresMergedFailure: boolean;
  readonly requiresProtection: boolean;
}

interface NativeFixtureReadback {
  readonly dimensionsMatch: boolean;
  readonly contentMatches: boolean;
  readonly sentinelMatches: boolean;
  readonly formulaPresent: boolean;
  readonly mergedStateMatches: boolean;
  readonly protectionStateMatches: boolean;
  readonly formatStateMatches: boolean;
  readonly tableStateMatches: boolean;
}

interface NativeRunState {
  readonly createdSheetNames: Set<string>;
  previousActiveSheet?: string;
}

const DATA_GRID: readonly (readonly FixtureCell[])[] = [
  ["ID", "Region", "Amount", "Rate", "Date", "Blank", "Formula", "Error"],
  [1, "North", 1250, 0.1, "2026-01-15", "", "=C2*D2", "=1/0"],
  [2, "South", 875, 0.2, "2026-02-15", "", "", "=1/0"],
  [3, "North", 1500, 0.15, "2026-03-15", "", "", "=1/0"],
  [4, "East", 0, 0.05, "2026-04-15", "", "", "=1/0"],
  [5, "West", -25, 0, "2026-05-15", "", "", "=1/0"],
  [6, "North", 2500, 0.3, "2026-06-15", "", "", "=1/0"],
  [7, "South", 125, 0.25, "2026-07-15", "", "", "=1/0"],
];

const TABLE_GRID: readonly (readonly FixtureCell[])[] = [
  ["Region", "Product", "Amount", "Date"],
  ["North", "Alpha", 100, "2026-01-01"],
  ["South", "Beta", 200, "2026-01-02"],
  ["East", "Gamma", 300, "2026-01-03"],
  ["North", "Alpha", 100, "2026-01-04"],
  ["South", "Delta", 400, "2026-01-05"],
  ["East", "Gamma", 300, "2026-01-06"],
  ["North", "Epsilon", 500, "2026-01-07"],
  ["South", "Beta", 200, "2026-01-08"],
  ["East", "Zeta", 600, "2026-01-09"],
  ["North", "Alpha", 100, "2026-01-10"],
  ["South", "Delta", 400, "2026-01-11"],
  ["East", "Gamma", 300, "2026-01-12"],
  ["North", "Epsilon", 500, "2026-01-13"],
  ["South", "Beta", 200, "2026-01-14"],
  ["East", "Zeta", 600, "2026-01-15"],
  ["North", "Alpha", 100, "2026-01-16"],
  ["South", "Delta", 400, "2026-01-17"],
  ["East", "Gamma", 300, "2026-01-18"],
  ["North", "Epsilon", 500, "2026-01-19"],
  ["South", "Beta", 200, "2026-01-20"],
];

const INSERT_TABLE_GRID: readonly (readonly FixtureCell[])[] = [
  ["Region", "Product", "Amount", "Date"],
  ["North", "Insert Alpha", 10, "2026-02-01"],
  ["South", "Insert Beta", 20, "2026-02-02"],
  ["East", "Insert Gamma", 30, "2026-02-03"],
];

const FORMAT_NUMBER_FORMATS: readonly (readonly string[])[] = [
  ["@", "@", "yyyy-mm-dd", "$#,##0.00", "0.00%", "@"],
  ["@", "@", "yyyy-mm-dd", "$#,##0.00", "0.00%", "@"],
  ["@", "@", "yyyy-mm-dd", "$#,##0.00", "0.00%", "@"],
  ["@", "@", "yyyy-mm-dd", "$#,##0.00", "0.00%", "@"],
  ["@", "@", "yyyy-mm-dd", "$#,##0.00", "0.00%", "@"],
];

function blankGrid(rows: number, columns: number): readonly (readonly FixtureCell[])[] {
  return Array.from({ length: rows }, () => Array.from({ length: columns }, () => ""));
}

function fixtureSeeds(sheets: FullFunctionalityAcceptanceSheetSet): readonly FixtureSeed[] {
  return [
    {
      role: "data",
      sheet: sheets.data,
      address: "A1:H8",
      grid: DATA_GRID,
      sentinelAddress: "J1:J2",
      sentinelValues: [["DATA_SENTINEL_A"], ["DATA_SENTINEL_B"]],
      requiresMergedFailure: false,
      requiresProtection: false,
    },
    {
      role: "formats",
      sheet: sheets.formats,
      address: "A1:F5",
      grid: [
        ["Synthetic title band", "", "", "", "", ""],
        ["Header", "Long synthetic text for wrapping", "Date", "Currency", "Percent", "Sentinel scope"],
        ["Row 1", "Fixed fixture text", "2026-08-01", 1234.5, 0.25, "FORMAT_SCOPE"],
        ["Row 2", "Another fixed fixture text", "2026-08-02", -45, 0.05, "FORMAT_SCOPE"],
        ["Row 3", "Final fixed fixture text", "2026-08-03", 0, 1, "FORMAT_SCOPE"],
      ],
      sentinelAddress: "H1:H2",
      sentinelValues: [["FORMATS_SENTINEL_A"], ["FORMATS_SENTINEL_B"]],
      requiresMergedFailure: false,
      requiresProtection: false,
    },
    {
      role: "tables",
      sheet: sheets.tables,
      address: "A1:D21",
      grid: TABLE_GRID,
      sentinelAddress: "F1:F2",
      sentinelValues: [["TABLES_SENTINEL_A"], ["TABLES_SENTINEL_B"]],
      requiresMergedFailure: false,
      requiresProtection: false,
    },
    {
      role: "pivot",
      sheet: sheets.pivot,
      address: "A1:B3",
      grid: blankGrid(3, 2),
      sentinelAddress: "F1:F2",
      sentinelValues: [["PIVOT_SENTINEL_A"], ["PIVOT_SENTINEL_B"]],
      requiresMergedFailure: false,
      requiresProtection: false,
    },
    {
      role: "comments",
      sheet: sheets.comments,
      address: "A1:D4",
      grid: [
        ["Comment target", "", "", ""],
        ["Update target", "", "", ""],
        ["Clear target", "", "", ""],
        ["Outside target", "", "", ""],
      ],
      sentinelAddress: "F1:F2",
      sentinelValues: [["COMMENTS_SENTINEL_A"], ["COMMENTS_SENTINEL_B"]],
      requiresMergedFailure: false,
      requiresProtection: false,
    },
    {
      role: "protected",
      sheet: sheets.protected,
      address: "A1:D4",
      grid: [
        ["Protected header", "", "", ""],
        ["Protected row 1", 10, "", ""],
        ["Protected row 2", 20, "", ""],
        ["Protected row 3", 30, "", ""],
      ],
      sentinelAddress: "F1:F2",
      sentinelValues: [["PROTECTED_SENTINEL_A"], ["PROTECTED_SENTINEL_B"]],
      requiresMergedFailure: false,
      requiresProtection: true,
    },
    {
      role: "expectedFailures",
      sheet: sheets.expectedFailures,
      address: "A1:F4",
      grid: [
        ["Merged failure anchor", "", "Occupied failure", "", "", ""],
        ["Expected failure row", "", "", "", "", ""],
        ["Invalid target row", "", "", "", "", ""],
        ["Protected target row", "", "", "", "", ""],
      ],
      sentinelAddress: "H1:H2",
      sentinelValues: [["FAILURES_SENTINEL_A"], ["FAILURES_SENTINEL_B"]],
      requiresMergedFailure: true,
      requiresProtection: true,
    },
  ];
}

async function createAndSeedFixtures(
  sheets: FullFunctionalityAcceptanceSheetSet,
  state: NativeRunState,
): Promise<void> {
  await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const existingSheets = worksheets;
    existingSheets.load("items/name");
    const activeSheet = worksheets.getActiveWorksheet();
    activeSheet.load("name");
    await context.sync();

    const targetNames = Object.values(sheets).filter((name): name is string => name !== sheets.runId);
    if (targetNames.some((name) => existingSheets.items.some((sheet) => sheet.name === name))) {
      throw new Error("acceptance fixture name collision");
    }
    state.previousActiveSheet = activeSheet.name;

    for (const seed of fixtureSeeds(sheets)) {
      const sheet = worksheets.add(seed.sheet);
      state.createdSheetNames.add(seed.sheet);
      const range = sheet.getRange(seed.address);
      range.formulas = seed.grid as Excel.Range["formulas"];
      sheet.getRange(seed.sentinelAddress).values = seed.sentinelValues as Excel.Range["values"];
      if (seed.role === "formats") {
        range.format.font.bold = true;
        range.format.fill.color = "#D9EAF7";
        range.numberFormat = FORMAT_NUMBER_FORMATS as Excel.Range["numberFormat"];
      }
      if (seed.role === "tables") {
        sheet.getRange("H1:K4").formulas = INSERT_TABLE_GRID as Excel.Range["formulas"];
        const table = sheet.tables.add(seed.address, true);
        table.name = tableName(sheets);
        table.style = "TableStyleMedium2";
      }
      if (seed.role === "expectedFailures") {
        sheet.getRange("A1:B1").merge(false);
      }
      if (seed.requiresProtection) {
        sheet.protection.protect({}, "");
      }
    }
    const dataSheet = worksheets.getItem(sheets.data);
    dataSheet.activate();
    await context.sync();
  });
}

async function readFixture(
  sheets: FullFunctionalityAcceptanceSheetSet,
  seed: FixtureSeed,
): Promise<NativeFixtureReadback> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(seed.sheet);
    const range = sheet.getRange(seed.address);
    const sentinel = sheet.getRange(seed.sentinelAddress);
    const merged = range.getMergedAreasOrNullObject();
    const tables = sheet.tables;
    const tableRange =
      seed.role === "tables"
        ? tables.getItem(tableName(sheets)).getRange()
        : null;
    const insertTableRange =
      seed.role === "tables" ? sheet.getRange("H1:K4") : null;
    range.load(["rowCount", "columnCount", "formulas", "numberFormat"]);
    range.format.fill.load("color");
    range.format.font.load("bold");
    sentinel.load("values");
    merged.load("isNullObject");
    tables.load("items/name");
    tableRange?.load(["address", "rowCount", "columnCount"]);
    insertTableRange?.load("formulas");
    const protection = sheet.protection as unknown as {
      load: (property: string) => void;
      protected: boolean;
    };
    protection.load("protected");
    await context.sync();

    const formulas = (range.formulas ?? []) as unknown[][];
    const sentinelValues = (sentinel.values ?? []) as unknown[][];
    const expectedRows = seed.grid.length;
    const expectedColumns = seed.grid[0]?.length ?? 0;
    const contentMatches = JSON.stringify(formulas) === JSON.stringify(seed.grid);
    const sentinelMatches = JSON.stringify(sentinelValues) === JSON.stringify(seed.sentinelValues);
    const formulaPresent = formulas.some((row) =>
      row.some((value) => typeof value === "string" && value.startsWith("=")),
    );
    return {
      dimensionsMatch: range.rowCount === expectedRows && range.columnCount === expectedColumns,
      contentMatches,
      sentinelMatches,
      formulaPresent: seed.role === "data" ? formulaPresent : true,
      mergedStateMatches: seed.requiresMergedFailure ? merged.isNullObject === false : merged.isNullObject === true,
      protectionStateMatches: seed.requiresProtection ? protection.protected === true : true,
      formatStateMatches:
        seed.role === "formats"
          ? range.format.font.bold === true &&
            range.format.fill.color.toUpperCase() === "#D9EAF7" &&
            JSON.stringify(range.numberFormat) === JSON.stringify(FORMAT_NUMBER_FORMATS)
          : true,
      tableStateMatches:
        seed.role === "tables"
          ? tables.items.length === 1 &&
            tables.items[0].name === tableName(sheets) &&
            tableRange !== null &&
            tableRange.rowCount === TABLE_GRID.length &&
            tableRange.columnCount === (TABLE_GRID[0]?.length ?? 0) &&
            tableRange.address.replace(/\$/g, "").endsWith("!A1:D21") &&
            insertTableRange !== null &&
            JSON.stringify(insertTableRange.formulas) === JSON.stringify(INSERT_TABLE_GRID)
          : tables.items.length === 0,
    };
  });
}

async function readBackFixtures(
  sheets: FullFunctionalityAcceptanceSheetSet,
): Promise<ReadonlyMap<FixtureRole, boolean>> {
  const results = new Map<FixtureRole, boolean>();
  for (const seed of fixtureSeeds(sheets)) {
    try {
      const readback = await readFixture(sheets, seed);
      results.set(
        seed.role,
        readback.dimensionsMatch &&
          readback.contentMatches &&
          readback.sentinelMatches &&
          readback.formulaPresent &&
          readback.mergedStateMatches &&
          readback.protectionStateMatches &&
          readback.formatStateMatches &&
          readback.tableStateMatches,
      );
    } catch {
      results.set(seed.role, false);
    }
  }
  return results;
}

async function cleanupFixtures(
  state: NativeRunState,
): Promise<{ removedCount: number; pass: boolean }> {
  if (state.createdSheetNames.size === 0) {
    return { removedCount: 0, pass: true };
  }

  try {
    const expectedCount = state.createdSheetNames.size;
    const removedCount = await Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load("items/name");
      await context.sync();
      const owned = worksheets.items.filter((sheet) => state.createdSheetNames.has(sheet.name));
      const protections = owned.map((sheet) => {
        const protection = sheet.protection as unknown as {
          load: (property: string) => void;
          protected: boolean;
          unprotect: (password?: string) => void;
        };
        protection.load("protected");
        return protection;
      });
      await context.sync();
      for (let index = 0; index < owned.length; index += 1) {
        if (protections[index].protected) {
          protections[index].unprotect("");
        }
        owned[index].delete();
      }
      if (
        state.previousActiveSheet &&
        worksheets.items.some((sheet) => sheet.name === state.previousActiveSheet)
      ) {
        worksheets.getItem(state.previousActiveSheet).activate();
      }
      await context.sync();
      return owned.length;
    });
    state.createdSheetNames.clear();
    return { removedCount, pass: removedCount === expectedCount };
  } catch {
    return { removedCount: 0, pass: false };
  }
}

function evidenceAfterFixtureReadback(
  pending: FullFunctionalityAcceptanceEvidence,
  readback: ReadonlyMap<FixtureRole, boolean>,
): FullFunctionalityAcceptanceEvidence {
  const cases = pending.cases.map((candidate, index) => {
    const definition = FULL_FUNCTIONALITY_ACCEPTANCE_CASES[index];
    const fixtureReady = readback.get(definition.fixture) === true;
    return {
      ...candidate,
      // This is setup/readback evidence only. The tool result, post-state, and
      // sentinel-after-mutation slots stay pending until Phase C executes.
      pre: fixtureReady ? passedPhase() : failedPhase(),
    };
  });
  return { ...pending, cases };
}

export async function runFullFunctionalityAcceptance(): Promise<FullFunctionalityAcceptanceEvidence> {
  const excel = (globalThis as { Excel?: { run?: unknown } }).Excel;
  if (!excel || typeof excel.run !== "function") {
    throw new Error("real Excel host required");
  }

  const runId = createFullFunctionalityAcceptanceRunId();
  const sheets = createFullFunctionalityAcceptanceSheetSet(runId);
  const pending = createPendingFullFunctionalityAcceptanceEvidence(runId);
  const state: NativeRunState = { createdSheetNames: new Set<string>() };
  let fixtureReadback = new Map<FixtureRole, boolean>();
  let cleanup = { attempted: false, removedCount: 0, pass: false };
  let evidence = pending;

  try {
    await createAndSeedFixtures(sheets, state);
    fixtureReadback = new Map(await readBackFixtures(sheets));
    evidence = evidenceAfterFixtureReadback(evidence, fixtureReadback);
  } catch {
    evidence = evidenceAfterFixtureReadback(evidence, fixtureReadback);
  } finally {
    cleanup = { attempted: true, ...(await cleanupFixtures(state)) };
  }

  return {
    ...evidence,
    pass: false,
    passCount: 0,
    failCount: evidence.cases.length,
    cleanup,
  };
}

export const __fullFunctionalityAcceptanceFixtureRoles = FULL_FUNCTIONALITY_ROLE_SUFFIX;
