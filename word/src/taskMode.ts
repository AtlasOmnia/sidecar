export const TASK_MODES = {
  SUMMARIZE: "summarize",
  REWRITE_SELECTION: "rewrite_selection",
  FORMAT_SELECTION: "format_selection",
  PROOFREAD_REVIEW: "proofread_review",
  PROOFREAD_ANNOTATE: "proofread_annotate",
  PROOFREAD_EDIT: "proofread_edit",
  COMPLEX_AGENT: "complex_agent"
} as const;

export type TaskMode = (typeof TASK_MODES)[keyof typeof TASK_MODES];

const SUMMARY_KEYWORDS = [
  /\b(summary|summarize|summarisation|summarization|abstract|brief)\b/i,
  /\btake away the gist\b/i,
  /\btldr\b/i
];

const PROOFREAD_KEYWORDS = [
  /\bproof[\s-]*read\b/i,
  /\bgrammar\b/i,
  /\bspelling\b/i,
  /\bcheck\s+for\s+(grammar|spelling|errors?)\b/i,
  /\bcheck\b.{0,50}\b(grammar|spelling|errors?)\b/i,
  /\breview\s+for\s+(issues?|problems?)\b/i,
  /\breview\s+this\s+for\s+(issues?|problems?)\b/i,
  /\b(?:review|check)\s+(?:this|that|it)\b/i,
  /\bfind\s+(grammar|style|errors?)\b/i,
  /\b(?:review|check)\b.{0,50}\b(document|text|selection|writing|content|draft)\b/i
];

const EDIT_INTENT_KEYWORDS = [
  /\b(correct|fix|edit|revise)\b/i,
  /\bapply\b.{0,30}\b(changes?|corrections?|edits?)\b/i,
  /\bmake\b.{0,30}\b(changes?|corrections?|edits?)\b/i
];

const ANNOTATION_INTENT_KEYWORDS = [
  /\bannotat(?:e|ed|ing|ion)\b/i,
  /\b(?:add|insert|leave|place|write|post)\b.{0,30}\bcomments?\b/i,
  /\bcomments?\b.{0,20}\b(?:on|about|for)\b/i,
  /\bwith\s+comments?\b/i
];

const EXPLICIT_CORRECTION_KEYWORDS = [
  /\b(?:correct|fix|revise)\b.{0,30}\b(?:document|text|selection|writing|errors?|grammar|spelling|typos?|mistakes?|issues?|problems?)\b/i,
  /\b(?:document|text|selection|writing|errors?|grammar|spelling|typos?|mistakes?|issues?|problems?)\b.{0,30}\b(?:correct|fix|revise)\b/i,
  /\bapply\b.{0,30}\bcorrections?\b/i
];

const FORMAT_KEYWORDS = [
  /\bformat\b/i,
  /\bbold\b/i,
  /\bitalic\b/i,
  /\bunderline\b/i,
  /\bformatting\b/i,
  /\bhighlight\b/i,
  /\bfont\b/i,
  /\bfont\s+size\b/i,
  /\bchange\s+style\b/i
];

const REWRITE_KEYWORDS = [
  /\brewrite\b/i,
  /\bparaphrase\b/i,
  /\brephrase\b/i,
  /\bwhat i highlighted\b/i
];

export function classifyTaskMode(rawRequest: string): TaskMode {
  const normalized = normalize(rawRequest);
  const hasProofreadIntent = matchesAny(normalized, PROOFREAD_KEYWORDS);
  const hasEditIntent = matchesAny(normalized, EDIT_INTENT_KEYWORDS);
  const hasAnnotationIntent = matchesAny(normalized, ANNOTATION_INTENT_KEYWORDS);
  const hasCorrectionIntent = (hasProofreadIntent && hasEditIntent)
    || matchesAny(normalized, EXPLICIT_CORRECTION_KEYWORDS);

  if (matchesAny(normalized, SUMMARY_KEYWORDS)) {
    return TASK_MODES.SUMMARIZE;
  }

  if (hasProofreadIntent || hasAnnotationIntent || hasCorrectionIntent) {
    if (hasCorrectionIntent) {
      return TASK_MODES.PROOFREAD_EDIT;
    }
    if (hasAnnotationIntent) {
      return TASK_MODES.PROOFREAD_ANNOTATE;
    }
    return TASK_MODES.PROOFREAD_REVIEW;
  }

  if (matchesAny(normalized, REWRITE_KEYWORDS)) {
    return TASK_MODES.REWRITE_SELECTION;
  }

  if (matchesAny(normalized, FORMAT_KEYWORDS)) {
    return TASK_MODES.FORMAT_SELECTION;
  }

  return TASK_MODES.COMPLEX_AGENT;
}

export function getIterationBudget(taskMode: TaskMode | string, maxIterations: number): number {
  const safeMaxIterations = Math.max(1, Number.isFinite(maxIterations) ? Math.floor(maxIterations) : 1);

  if (taskMode === TASK_MODES.SUMMARIZE) {
    return Math.min(2, safeMaxIterations);
  }
  if (taskMode === TASK_MODES.REWRITE_SELECTION) {
    return Math.min(3, safeMaxIterations);
  }
  if (taskMode === TASK_MODES.FORMAT_SELECTION) {
    return Math.min(3, safeMaxIterations);
  }
  if (taskMode === TASK_MODES.PROOFREAD_REVIEW) {
    return Math.min(2, safeMaxIterations);
  }
  if (taskMode === TASK_MODES.PROOFREAD_ANNOTATE) {
    return Math.min(2, safeMaxIterations);
  }
  if (taskMode === TASK_MODES.PROOFREAD_EDIT) {
    return Math.min(4, safeMaxIterations);
  }

  return safeMaxIterations;
}

function normalize(rawRequest: string): string {
  return (rawRequest || "").toLowerCase().replace(/\s+/g, " ").trim();
}

function matchesAny(normalized: string, patterns: RegExp[]): boolean {
  return patterns.some((pattern) => pattern.test(normalized));
}
