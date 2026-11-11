import { describe, expect, it } from "vitest";
import { classifyWorkbookRequest, type TrustedWorkbookRequest } from "../workbookPlanRequest";
import {
  buildWorkbookPlanPolicyInput,
  type WorkbookPlanRouteRuntimeContext,
} from "../workbookPlanRoute";
import {
  createWorkbookPlanPolicyInputSnapshot,
  decideWorkbookPlanPolicy,
  type WorkbookRequestClassification,
} from "../workbookPlanPolicy";
import type { WorkbookPlan } from "../generated/workbookPlanSchema";

/*
 * R4b adversarial request-classification corpus. The classifier may only narrow:
 * for every prompt the new classification is at least as strict as the legacy
 * (base) classification, and selectionScopedRequest never changes.
 */

// BEGIN BASE ORACLE 8637337:src/workbookPlanRequest.ts
const STRUCTURE_WORDS = /\b(create|add|build|new|insert|make|rename|copy|table|chart|sheet)\b/i;
const MUTATION_WORDS = /\b(write|edit|update|change|format|fill|formula|sort|filter|comment|calculate|apply|set|populate|enter)\b/i;
const READ_WORDS = /\b(read|inspect|show|list|summarize|summary|count|average|stats|what|which|how many|look at)\b/i;
const PROHIBITED_WORDS = /\b(delete|remove|clear|erase|unprotect|replace|destroy)\b/i;
const SELECTION_WORDS = /\b(selected|selection|highlighted|highlight|current range|this range|these cells|active range)\b/i;

function oracleClassify(userRequest: string): {
  classification: WorkbookRequestClassification;
  selectionScopedRequest: boolean;
} {
  const request = typeof userRequest === "string" ? userRequest.trim() : "";
  const selectionScopedRequest = SELECTION_WORDS.test(request);
  let classification: WorkbookRequestClassification;

  if (request.length === 0) {
    classification = "ambiguous_mutation";
  } else if (PROHIBITED_WORDS.test(request)) {
    classification = "prohibited";
  } else if (STRUCTURE_WORDS.test(request)) {
    classification = "explicit_structure";
  } else if (MUTATION_WORDS.test(request)) {
    classification = "explicit_mutation";
  } else if (READ_WORDS.test(request)) {
    classification = "read_only";
  } else {
    classification = "ambiguous_mutation";
  }
  return { classification, selectionScopedRequest };
}
// END BASE ORACLE

const RANK: Readonly<Record<WorkbookRequestClassification, number>> = {
  explicit_structure: 0,
  explicit_mutation: 1,
  read_only: 2,
  ambiguous_mutation: 3,
  prohibited: 4,
};

const S = "explicit_structure" as const;
const M = "explicit_mutation" as const;
const R = "read_only" as const;
const A = "ambiguous_mutation" as const;
const P = "prohibited" as const;

type Category =
  | "noun"
  | "negation"
  | "selection-negation"
  | "narrow"
  | "non-ascii"
  | "mixed"
  | "non-english"
  | "prohibited"
  | "pin";

interface CorpusRow {
  readonly id: string;
  readonly lang: string;
  readonly category: Category;
  readonly prompt: string;
  readonly base: WorkbookRequestClassification;
  readonly expected: WorkbookRequestClassification;
  readonly selection: boolean;
}

const row = (
  id: string,
  lang: string,
  category: Category,
  prompt: string,
  base: WorkbookRequestClassification,
  expected: WorkbookRequestClassification,
  selection = false,
): CorpusRow => ({ id, lang, category, prompt, base, expected, selection });

// R4b release §5.1 (72 prompts). base = legacy class at 8637337/8caffd8; expected = R4b class.
const CORPUS: readonly CorpusRow[] = [
  row("en-noun-1", "en", "noun", "show me the table", S, R),
  row("en-noun-2", "en", "noun", "which sheet has the totals?", S, R),
  row("en-noun-3", "en", "noun", "what does the chart show", S, R),
  row("en-noun-4", "en", "noun", "read the formula in B2", M, R),
  row("en-noun-5", "en", "noun", "list the comment on C3", M, R),
  row("en-noun-6", "en", "noun", "show the data set in A1:D20", M, R),
  row("en-noun-7", "en", "noun", "how many rows are in the sales table", S, R),
  row("en-noun-8", "en", "noun", "describe the chart on Sheet2", S, R),
  row("en-noun-9", "en", "noun", "explain the formula in C5", M, R),
  row("en-noun-10", "en", "noun", "Could you tell me which sheet is active?", S, R),
  row("en-neg-1", "en", "negation", "don't change anything, just summarize the table", S, A),
  row("en-neg-2", "en", "negation", "do not format, only count the rows", M, A),
  row("en-neg-3", "en", "negation", "summarize the sheet without editing it", S, R),
  row("en-neg-4", "en", "negation", "never sort the table, tell me the top value", S, A),
  row("en-neg-5", "en", "negation", "No new sheets please, what is in A1?", S, A),
  row("en-neg-6", "en", "negation", "dont update the chart", S, A),
  row("en-neg-7", "en", "negation", "show the totals but don\u2019t write anything", M, R),
  row("en-sel-1", "en", "selection-negation", "not the selected range, write to D1", M, A, true),
  row("en-sel-2", "en", "selection-negation", "write the totals outside the selected cells", M, A, true),
  row("en-sel-3", "en", "selection-negation", "format everything except the highlighted cells", M, A, true),
  row("en-narrow-1", "en", "narrow", "make it bold", S, M),
  row("en-narrow-2", "en", "narrow", "write 5 in A1 on the summary sheet", S, M),
  row("en-narrow-3", "en", "narrow", "copy A1 to B1", S, M),
  row("en-narrow-4", "en", "narrow", "add a comment to B2 saying checked", S, M),
  row("en-narrow-5", "en", "narrow", "sort the table by Amount", S, M),
  row("mix-de-1", "de+en", "mixed", "bitte table erstellen", S, A),
  row("mix-es-1", "es+en", "mixed", "insertar chart en la hoja", S, A),
  row("mix-ja-1", "ja+en", "mixed", "\u8868\u3092\u4f5c\u6210\u3057\u3066 table", S, A),
  row("mix-zh-1", "zh+en", "mixed", "\u521b\u5efa\u4e00\u4e2a\u56fe\u8868 chart", S, A),
  row("mix-fr-1", "fr+en", "mixed", "cr\u00e9e une table sur la feuille", S, A),
  row("mix-es-2", "es+en", "mixed", "a\u00f1adir una fila a la table", S, A),
  row("mix-de-2", "de+en", "mixed", "Tabelle erstellen und sheet umbenennen", S, A),
  row("mix-zh-2", "zh+en", "mixed", "\u5199\u5165 A1 write ok", M, A),
  row("mix-fr-2", "fr+en", "mixed", "r\u00e9sum\u00e9 of the table: show me", S, A),
  row("en-cafe", "en", "non-ascii", "write 5 into A1 on sheet Caf\u00e9", S, A),
  row("fr-1", "fr", "non-english", "cr\u00e9er un tableau", A, A),
  row("es-1", "es", "non-english", "borra la hoja", A, A),
  row("ja-1", "ja", "non-english", "\u8868\u3092\u4f5c\u6210", A, A),
  row("es-2", "es", "non-english", "mostrar la tabla", A, A),
  row("de-1", "de", "non-english", "zeige mir die Summe", A, A),
  row("zh-1", "zh", "non-english", "\u663e\u793a\u603b\u548c", A, A),
  row("ja-2", "ja", "non-english", "A1\u306b\u66f8\u3044\u3066", A, A),
  row("fr-2", "fr", "non-english", "affiche le tableau", A, A),
  row("fr-3", "fr", "non-english", "\u00e9crire 5 dans A1", A, A),
  row("de-2", "de", "non-english", "schreibe 5 in A1", A, A),
  row("zh-2", "zh", "non-english", "\u5220\u9664\u5de5\u4f5c\u8868", A, A),
  row("mix-es-3", "es+en", "mixed", "show la tabla", R, R),
  row("mix-es-4", "es+en", "mixed", "list la sheet de ventas", S, R),
  row("pin-dash", "en", "pin", "please write ok into A1 \u2013 thanks", M, M),
  row("pin-quotes", "en", "pin", "write \u201cok\u201d into A1", M, M),
  row("pro-1", "en", "prohibited", "delete the sheet", P, P),
  row("pro-2", "en", "prohibited", "don't delete anything, show the table", P, P),
  row("pro-3", "en", "prohibited", "clear A1:B2 and create a table", P, P),
  row("pro-4", "en", "prohibited", "remove the chart", P, P),
  row("pro-5", "es+en", "prohibited", "borrar y delete hoja", P, P),
  row("pro-6", "en", "prohibited", "find and replace foo with bar", P, P),
  row("pin-1", "en", "pin", "write ok into A1", M, M),
  row("pin-2", "en", "pin", "Create a table from the selected range", S, S, true),
  row("pin-3", "en", "pin", "create a chart from A1:B5 on a new sheet", S, S),
  row("pin-4", "en", "pin", "rename the sheet to Q3", S, S),
  row("pin-5", "en", "pin", "fill the formula down to row 20", M, M),
  row("pin-6", "en", "pin", "calculate the workbook", M, M),
  row("pin-7", "en", "pin", "Please inspect the selected range", R, R, true),
  row("pin-8", "en", "pin", "do the usual thing", A, A),
  row("pin-9", "en", "pin", "update the numbers", M, M),
  row("pin-10", "en", "pin", "show me the totals", R, R),
  row("pin-11", "en", "pin", "yes", A, A),
  row("pin-12", "en", "pin", "", A, A),
  row("pin-13", "en", "pin", "set B2 to 10", M, M),
  row("pin-14", "en", "pin", "what is the average of the selection", R, R, true),
  row("pin-15", "en", "pin", "summarize this range and then format it as currency", M, M, true),
  row("pin-16", "en", "pin", "show me how to create a table", S, S),
];

// Q3 fixtures copied from task11-route-policy-resource.test.ts:24-49 (runtime, executePlan).
const runtime: WorkbookPlanRouteRuntimeContext = {
  activeSheetName: "Sheet1",
  selection: { sheetName: "Sheet1", address: "B2:C3" },
  compileContext: {
    sheetNames: new Set(["Sheet1", "Sheet2"]),
    tablesByName: new Map(),
    supportedPrimitiveNames: new Set(),
    supportedCapabilities: new Set(),
    chartSourceResolutionsBySheet: new Map(),
  },
};

const executePlan = (overrides: Partial<Extract<WorkbookPlan, { outcome: "execute" }>> = {}): Extract<WorkbookPlan, { outcome: "execute" }> => ({
  version: 1,
  outcome: "execute",
  summary: "write the requested value",
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }],
  verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["ok"]] }],
  ...overrides,
});

const PLANS = {
  p1: executePlan({
    operations: [{ kind: "write_values", sheet: "Sheet1", address: "B2", values: [["ok"]] }],
    verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "B2", expectedValues: [["ok"]] }],
  }),
  p2: executePlan({
    summary: "create a table",
    operations: [{ kind: "create_table", sheet: "Sheet1", address: "B2:C3", hasHeaders: true }],
    verification: [{ kind: "calculation_completed" }],
  }),
  p3: executePlan({
    summary: "calculate",
    operations: [{ kind: "calculate" }],
    verification: [{ kind: "calculation_completed" }],
  }),
} as const;

type PlanId = keyof typeof PLANS;
type DecisionKind = "approved" | "requires_confirmation" | "rejected";
const DECISION_STRICTNESS: Readonly<Record<DecisionKind, number>> = {
  approved: 0,
  requires_confirmation: 1,
  rejected: 2,
};

const decide = (planId: PlanId, request: TrustedWorkbookRequest): DecisionKind =>
  decideWorkbookPlanPolicy(
    createWorkbookPlanPolicyInputSnapshot(buildWorkbookPlanPolicyInput(PLANS[planId], runtime, request)),
  ).decision;

const P1_FOR_CLASS: Readonly<Record<WorkbookRequestClassification, DecisionKind>> = {
  explicit_structure: "approved",
  explicit_mutation: "approved",
  read_only: "requires_confirmation",
  ambiguous_mutation: "requires_confirmation",
  prohibited: "rejected",
};

const MUST_NOT_APPROVE_P1: ReadonlySet<Category> = new Set<Category>([
  "noun",
  "negation",
  "selection-negation",
  "non-ascii",
  "mixed",
]);

const label = (r: CorpusRow) => `${r.id} ${JSON.stringify(r.prompt)}`;
const CASES = CORPUS.map((r) => [label(r), r] as const);

describe("R4b corpus table integrity", () => {
  it("holds 72 unique prompts with the release language and RED counts", () => {
    expect(CORPUS).toHaveLength(72);
    expect(new Set(CORPUS.map((r) => r.id)).size).toBe(72);
    const nonEnglish = CORPUS.filter((r) => r.lang !== "en");
    expect(nonEnglish).toHaveLength(23);
    const langs = (prefix: string) => CORPUS.filter((r) => r.lang === prefix || r.lang === `${prefix}+en`).length;
    expect({ es: langs("es"), fr: langs("fr"), de: langs("de"), ja: langs("ja"), zh: langs("zh") }).toEqual({
      es: 7,
      fr: 5,
      de: 4,
      ja: 3,
      zh: 4,
    });
    expect(CORPUS.filter((r) => r.base !== r.expected)).toHaveLength(36);
    expect(
      CORPUS.filter((r) => MUST_NOT_APPROVE_P1.has(r.category) && P1_FOR_CLASS[r.base] === "approved"),
    ).toHaveLength(31);
    for (const r of CORPUS) expect(RANK[r.expected]).toBeGreaterThanOrEqual(RANK[r.base]);
  });
});

describe("Q10 base oracle self-check", () => {
  it.each(CASES)("oracle reproduces the base column: %s", (_name, r) => {
    expect(oracleClassify(r.prompt)).toEqual({ classification: r.base, selectionScopedRequest: r.selection });
  });
});

describe("Q1 corpus classification", () => {
  it.each(CASES)("classifies %s", (_name, r) => {
    expect(classifyWorkbookRequest(r.prompt).classification).toBe(r.expected);
  });
});

describe("Q2 corpus monotone narrowing", () => {
  it.each(CASES)("never widens and keeps selection: %s", (_name, r) => {
    const oracle = oracleClassify(r.prompt);
    const next = classifyWorkbookRequest(r.prompt);
    expect(RANK[next.classification]).toBeGreaterThanOrEqual(RANK[oracle.classification]);
    expect(next.selectionScopedRequest).toBe(oracle.selectionScopedRequest);
  });
});

describe("Q3 policy cross-check", () => {
  it.each(CASES)("Q3a policy is never more permissive than base (base oracle decision pinned): %s", (_name, r) => {
    const oracleRequest: TrustedWorkbookRequest = oracleClassify(r.prompt);
    const nextRequest = classifyWorkbookRequest(r.prompt);
    for (const planId of ["p1", "p2", "p3"] as const) {
      const before = decide(planId, oracleRequest);
      const after = decide(planId, nextRequest);
      expect(DECISION_STRICTNESS[after], `${planId} ${before} -> ${after}`).toBeGreaterThanOrEqual(
        DECISION_STRICTNESS[before],
      );
    }
    expect(decide("p1", oracleRequest)).toBe(P1_FOR_CLASS[r.base]);
  });
  it.each(CASES)("Q3b expected new policy decision for the one-op write plan: %s", (_name, r) => {
    const nextRequest = classifyWorkbookRequest(r.prompt);
    expect(decide("p1", nextRequest)).toBe(P1_FOR_CLASS[r.expected]);
    if (MUST_NOT_APPROVE_P1.has(r.category)) {
      expect(decide("p1", nextRequest)).not.toBe("approved");
    }
  });

  it("builds all three plan fixtures through accounting and policy", () => {
    const request: TrustedWorkbookRequest = { classification: "explicit_structure", selectionScopedRequest: false };
    expect(decide("p1", request)).toBe("approved");
    expect(decide("p2", request)).toBe("approved");
    expect(decide("p3", request)).toBe("approved");
    const mutation: TrustedWorkbookRequest = { classification: "explicit_mutation", selectionScopedRequest: false };
    expect(decide("p2", mutation)).toBe("requires_confirmation");
    expect(decide("p3", { classification: "read_only", selectionScopedRequest: true })).toBe("approved");
  });
});

describe("Q4 existing classifier expectations stay exact", () => {
  it.each([
    ["Please inspect the selected range", "read_only", true],
    ["Create a table from the selected range", "explicit_structure", true],
    ["write ok into A1", "explicit_mutation", false],
    ["do the usual thing", "ambiguous_mutation", false],
  ] as const)("%s -> %s", (prompt, classification, selectionScopedRequest) => {
    expect(classifyWorkbookRequest(prompt)).toEqual({ classification, selectionScopedRequest });
  });

  it.each(["yes", "proceed", "approved"])("caller event never confirms the prompt %s", (prompt) => {
    const request = classifyWorkbookRequest(prompt, {
      eventId: "11111111-1111-4111-8111-111111111111",
      confirmed: true,
    });
    expect(request.confirmation).toBeUndefined();
    expect("confirmation" in request).toBe(false);
  });
});

describe("Q5 confirmation is never derived from text", () => {
  it("no corpus prompt yields a confirmation key with a caller event", () => {
    for (const r of CORPUS) {
      const request = classifyWorkbookRequest(r.prompt, { eventId: `event-${r.id}`, confirmed: true });
      expect("confirmation" in request, r.id).toBe(false);
    }
  });
});

describe("Q6 result shape", () => {
  it("returns exactly classification and selectionScopedRequest, in order", () => {
    for (const r of CORPUS) {
      expect(Object.keys(classifyWorkbookRequest(r.prompt)), r.id).toEqual(["classification", "selectionScopedRequest"]);
    }
  });

  it("treats non-string input as an empty ambiguous request", () => {
    const result = classifyWorkbookRequest(undefined as unknown as string);
    expect(result).toEqual({ classification: "ambiguous_mutation", selectionScopedRequest: false });
    expect(Object.keys(result)).toEqual(["classification", "selectionScopedRequest"]);
  });
});

describe("Q7 generated monotone narrowing", () => {
  const POOL = [
    // read verbs and fillers
    "show", "show me", "read", "list", "what", "which", "how many", "summarize", "tell me", "describe",
    "explain", "inspect", "look at", "stats", "please", "could you", "just", "ok", "help me",
    // structure verbs and nouns
    "create", "add", "build", "new", "insert", "make", "rename", "copy", "table", "chart", "sheet",
    "worksheet", "tables", "charts", "sheets",
    // mutation verbs and nouns
    "write", "edit", "update", "change", "format", "fill", "sort", "filter", "calculate", "set", "enter",
    "formula", "formulas", "comment", "data set", "dataset",
    // prohibited
    "delete", "clear", "remove", "replace",
    // negations
    "don't", "dont", "don\u2019t", "do not", "not", "no", "never", "without", "nothing", "stop", "can't",
    // selection phrases and selection negation
    "selected", "selection", "highlighted", "this range", "these cells", "the selected range", "except",
    "outside", "other than", "instead of",
    // non-ASCII words and symbols
    "caf\u00e9", "\u8868", "r\u00e9sum\u00e9", "\u521b\u5efa", "a\u00f1adir", "\u4f5c\u6210", "\u00fcber",
    "\u2013", "\u201cok\u201d", "\u{1F44D}", "\u200b",
    // neutral tokens and punctuation
    "the", "it", "A1", "B2:C3", "to", "on", "and", ",", ".", "?", "the totals",
  ];

  it("never widens or flips selection for >= 20,000 seeded prompts", () => {
    let seed = 20261006;
    const next = () => {
      seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff;
      return seed / 0x80000000;
    };
    let generated = 0;
    const widened: string[] = [];
    const histogram: Record<string, number> = {};
    for (let i = 0; i < 20000; i += 1) {
      const length = 1 + Math.floor(next() * 6);
      const words: string[] = [];
      for (let j = 0; j < length; j += 1) words.push(POOL[Math.floor(next() * POOL.length)]);
      const prompt = words.join(" ");
      generated += 1;
      const oracle = oracleClassify(prompt);
      const result = classifyWorkbookRequest(prompt);
      histogram[`${oracle.classification}->${result.classification}`] =
        (histogram[`${oracle.classification}->${result.classification}`] ?? 0) + 1;
      if (
        RANK[result.classification] < RANK[oracle.classification] ||
        result.selectionScopedRequest !== oracle.selectionScopedRequest
      ) {
        widened.push(prompt);
      }
    }
    expect(generated).toBeGreaterThanOrEqual(20000);
    expect(widened).toEqual([]);
    // The generator must reach every legacy class, or the check proves little.
    for (const cls of Object.keys(RANK)) {
      expect(Object.keys(histogram).some((key) => key.startsWith(`${cls}->`)), cls).toBe(true);
    }
  });
});

describe("Q8 non-letter non-ASCII does not trigger the non-ASCII rule", () => {
  it.each([
    "please write ok into A1 \u2013 thanks",
    "write \u201cok\u201d into A1",
    "write ok into A1 \u{1F44D}",
  ])("%s keeps its base class", (prompt) => {
    expect(classifyWorkbookRequest(prompt)).toEqual(oracleClassify(prompt));
    expect(classifyWorkbookRequest(prompt).classification).toBe("explicit_mutation");
  });
});

describe("Q9 bounded cost", () => {
  const big = (unit: string) => unit.repeat(Math.ceil(100_000 / unit.length));
  it.each([
    ["not the", big("not the ")],
    ["show", big("show ")],
    ["mixed scripts", big("\u8868 table caf\u00e9 \u521b\u5efa write ")],
    ["except gap", `${big("except ")}selected`],
    ["filler", big("please ")],
    ["noun run", big("data set table formula ")],
  ])("classifies a 100 KB %s input", (_name, prompt) => {
    expect(prompt.length).toBeGreaterThanOrEqual(100_000);
    const result = classifyWorkbookRequest(prompt);
    expect(RANK[result.classification]).toBeGreaterThanOrEqual(RANK[oracleClassify(prompt).classification]);
  });
});

// R4h (r4b tester P3-2): witnesses that fail if the legacy if/else chain is reordered, including the
// stricter-only reorders that the 72-row corpus cannot see. base = oracle class; expected = classifier class.
const LEGACY_ORDER_ROWS: readonly CorpusRow[] = [
  row("lo-1", "en", "pin", "create a table and fill it", S, S),
  row("lo-2", "en", "pin", "rename the sheet and fill the formula", S, S),
  row("lo-3", "en", "pin", "show and update the totals", M, M),
  row("lo-4", "en", "pin", "list the sheets then create a chart", S, S),
];

describe("Q11 legacy chain order pinned in the narrowing direction", () => {
  it.each(LEGACY_ORDER_ROWS.map((r) => [label(r), r] as const))("order witness %s", (_name, r) => {
    expect(oracleClassify(r.prompt)).toEqual({ classification: r.base, selectionScopedRequest: false });
    const next = classifyWorkbookRequest(r.prompt);
    expect(next).toEqual({ classification: r.expected, selectionScopedRequest: false });
    expect(RANK[next.classification]).toBeGreaterThanOrEqual(RANK[oracleClassify(r.prompt).classification]);
    expect(CORPUS.some((c) => c.prompt === r.prompt)).toBe(false);
  });
});
