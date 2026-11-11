import type { WorkbookRequestClassification } from "./workbookPlanPolicy";

export interface TrustedConfirmationEvent {
  readonly eventId: string;
  readonly confirmed: true;
  readonly source: "user";
}

export interface TrustedWorkbookRequest {
  readonly classification: WorkbookRequestClassification;
  readonly selectionScopedRequest: boolean;
  readonly confirmation?: TrustedConfirmationEvent;
}

const STRUCTURE_WORDS = /\b(create|add|build|new|insert|make|rename|copy|table|chart|sheet)\b/i;
const MUTATION_WORDS = /\b(write|edit|update|change|format|fill|formula|sort|filter|comment|calculate|apply|set|populate|enter)\b/i;
const READ_WORDS = /\b(read|inspect|show|list|summarize|summary|count|average|stats|what|which|how many|look at)\b/i;
const PROHIBITED_WORDS = /\b(delete|remove|clear|erase|unprotect|replace|destroy)\b/i;
const SELECTION_WORDS = /\b(selected|selection|highlighted|highlight|current range|this range|these cells|active range)\b/i;

// R4b narrowing-only refinement. The legacy chain above stays authoritative as a
// floor: the final class is the stricter of legacy and refined, so the refinement
// can never widen a request. Strictness follows plan policy (workbookPlanPolicy.ts).
const CLASSIFICATION_STRICTNESS: Readonly<Record<WorkbookRequestClassification, number>> = {
  explicit_structure: 0,
  explicit_mutation: 1,
  read_only: 2,
  ambiguous_mutation: 3,
  prohibited: 4,
};

const NON_ASCII_LETTER = /(?![\u0000-\u007f])\p{L}/u;
const NEGATION_WORDS = /\b(?:don['\u2019]?t|do not|does not|doesn['\u2019]?t|didn['\u2019]?t|never|without|not|no|nothing|avoid|stop|shouldn['\u2019]?t|mustn['\u2019]?t|must not|cannot|can['\u2019]?t)\b/i;
const SELECTION_NEGATION = /\b(?:not|outside|other than|except|excluding|instead of|rather than|apart from|besides|beyond)\b[^.;!?]{0,24}?\b(?:selected|selection|highlighted|highlight|current range|this range|these cells|active range)\b/i;
const LEADING_FILLER = /^(?:(?:please|pls|kindly|hey|hi|ok|okay|so|now|just|can you|could you|would you|will you|help me|i want to|i'd like to|i would like to)\b[\s,:]*)+/i;
const READ_LEAD = /^(?:read|inspect|show|list|summarize|summarise|count|average|what|which|how many|how much|look at|tell me|describe|explain|where|display|view|check|review)\b/i;
const NOUN_PHRASES = /\b(?:data ?sets?|tables?|charts?|sheets?|worksheets?|formulas?|comments?)\b/gi;
const STRUCTURE_VERBS = /\b(?:create|add|build|new|insert|make|rename|copy)\b/i;
const STRUCTURE_NOUNS = /\b(?:tables?|charts?|sheets?|worksheets?)\b/i;
const MUTATION_VERBS = /\b(?:write|edit|update|change|format|fill|sort|filter|calculate|apply|set|populate|enter)\b/i;
const MUTATION_NOUNS = /\b(?:formulas?|comments?)\b/i;

function legacyClassification(request: string): WorkbookRequestClassification {
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
  return classification;
}

function refineClassification(request: string): WorkbookRequestClassification {
  // Any non-ASCII letter: the English keyword lists cannot be trusted (D-R4-6).
  if (NON_ASCII_LETTER.test(request)) return "ambiguous_mutation";
  // "not the selected range, write to D1": the selection flag stays as is (never
  // true -> false); the class carries the narrowing instead.
  if (SELECTION_NEGATION.test(request)) return "ambiguous_mutation";
  const readLead = READ_LEAD.test(request.replace(LEADING_FILLER, ""));
  if (NEGATION_WORDS.test(request)) return readLead ? "read_only" : "ambiguous_mutation";
  const withoutNouns = request.replace(NOUN_PHRASES, " ");
  const structureVerb = STRUCTURE_VERBS.test(withoutNouns);
  const mutationVerb = MUTATION_VERBS.test(withoutNouns);
  if (readLead && !structureVerb && !mutationVerb) return "read_only";
  const structureNoun = STRUCTURE_NOUNS.test(request);
  if (structureVerb && structureNoun) return "explicit_structure";
  if (structureVerb || mutationVerb) return "explicit_mutation";
  if (structureNoun || MUTATION_NOUNS.test(request)) {
    return READ_WORDS.test(request) ? "read_only" : "ambiguous_mutation";
  }
  if (READ_WORDS.test(request)) return "read_only";
  return "ambiguous_mutation";
}

const stricterClassification = (
  left: WorkbookRequestClassification,
  right: WorkbookRequestClassification,
): WorkbookRequestClassification =>
  CLASSIFICATION_STRICTNESS[right] > CLASSIFICATION_STRICTNESS[left] ? right : left;

/**
 * Classify only the user-authored request. Model operations are intentionally
 * not consulted here; they are proposal evidence for a later policy check.
 */
export function classifyWorkbookRequest(
  userRequest: string,
  confirmationEvent?: { readonly eventId: string; readonly confirmed: boolean },
): TrustedWorkbookRequest {
  const request = typeof userRequest === "string" ? userRequest.trim() : "";
  const selectionScopedRequest = SELECTION_WORDS.test(request);
  const legacy = legacyClassification(request);
  const classification: WorkbookRequestClassification =
    legacy === "prohibited" || request.length === 0
      ? legacy
      : stricterClassification(legacy, refineClassification(request));

  // Confirmation is deliberately not derived from prompt text or a caller
  // supplied event. Only the active pane-owned Continue callback may mint it.
  void confirmationEvent;
  return { classification, selectionScopedRequest };
}
