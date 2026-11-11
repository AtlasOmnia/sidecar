import { describe, expect, it } from "vitest";
import { classifyTaskMode, getIterationBudget, TASK_MODES } from "../taskMode";

const PROOFREAD_ANNOTATE_MODE = "proofread_annotate";

describe("classifyTaskMode", () => {
  it('routes summarize requests to summarize mode', () => {
    expect(classifyTaskMode("Summarize this document for me")).toBe(TASK_MODES.SUMMARIZE);
    expect(classifyTaskMode("Give me a summary of the text")).toBe(TASK_MODES.SUMMARIZE);
  });

  it('routes proofreading without an edit verb to review-only mode', () => {
    expect(classifyTaskMode("Proofread this document")).toBe(TASK_MODES.PROOFREAD_REVIEW);
    expect(classifyTaskMode("proof read the document")).toBe(TASK_MODES.PROOFREAD_REVIEW);
    expect(classifyTaskMode("proof-read this document")).toBe(TASK_MODES.PROOFREAD_REVIEW);
    expect(classifyTaskMode("Check for grammar and spelling errors")).toBe(TASK_MODES.PROOFREAD_REVIEW);
    expect(classifyTaskMode("Review this for issues")).toBe(TASK_MODES.PROOFREAD_REVIEW);
  });

  it('routes ordinary review and check requests to read-only review mode', () => {
    expect(classifyTaskMode("Review this document")).toBe(TASK_MODES.PROOFREAD_REVIEW);
    expect(classifyTaskMode("Check this document")).toBe(TASK_MODES.PROOFREAD_REVIEW);
    expect(classifyTaskMode("Review the selected text for style issues")).toBe(TASK_MODES.PROOFREAD_REVIEW);
  });

  it('treats short review and check imperatives as read-only review', () => {
    expect(classifyTaskMode("Review this")).toBe(TASK_MODES.PROOFREAD_REVIEW);
    expect(classifyTaskMode("Check this")).toBe(TASK_MODES.PROOFREAD_REVIEW);
  });

  it('routes explicit annotation requests to annotation mode', () => {
    expect(classifyTaskMode("Proofread and annotate the issues")).toBe(PROOFREAD_ANNOTATE_MODE);
    expect(classifyTaskMode("Review the document and add comments for each error")).toBe(PROOFREAD_ANNOTATE_MODE);
    expect(classifyTaskMode("Annotate the spelling errors")).toBe(PROOFREAD_ANNOTATE_MODE);
  });

  it('routes proofreading with explicit change intent to edit mode', () => {
    expect(classifyTaskMode("Proofread and correct the document")).toBe(TASK_MODES.PROOFREAD_EDIT);
    expect(classifyTaskMode("Fix the grammar and spelling errors")).toBe(TASK_MODES.PROOFREAD_EDIT);
    expect(classifyTaskMode("Proof read and apply the corrections")).toBe(TASK_MODES.PROOFREAD_EDIT);
  });

  it('prioritizes explicit correction over annotation when both are requested', () => {
    expect(classifyTaskMode("Correct the errors and add comments for the changes")).toBe(TASK_MODES.PROOFREAD_EDIT);
  });

  it('routes rewrite-selection requests to rewrite_selection mode', () => {
    expect(classifyTaskMode("Rewrite the selected text")).toBe(TASK_MODES.REWRITE_SELECTION);
    expect(classifyTaskMode("Paraphrase what I highlighted")).toBe(TASK_MODES.REWRITE_SELECTION);
  });

  it('routes format-selection requests to format_selection mode', () => {
    expect(classifyTaskMode("Make the selection bold and italic")).toBe(TASK_MODES.FORMAT_SELECTION);
    expect(classifyTaskMode("Format the selected text as heading")).toBe(TASK_MODES.FORMAT_SELECTION);
  });

  it('does not treat a bare selection reference as formatting intent', () => {
    expect(classifyTaskMode("What does the selection say?")).toBe(TASK_MODES.COMPLEX_AGENT);
    expect(classifyTaskMode("Review the selected text for style issues")).toBe(TASK_MODES.PROOFREAD_REVIEW);
    expect(classifyTaskMode("Check the document selection for errors")).toBe(TASK_MODES.PROOFREAD_REVIEW);
  });

  it('falls back to complex_agent for ambiguous or complex requests', () => {
    expect(classifyTaskMode("Analyze this document deeply and suggest improvements")).toBe(TASK_MODES.COMPLEX_AGENT);
    expect(classifyTaskMode("What do you think about the structure?")).toBe(TASK_MODES.COMPLEX_AGENT);
    expect(classifyTaskMode("Help me with this project")).toBe(TASK_MODES.COMPLEX_AGENT);
  });

  it('falls back to complex_agent when no keywords match', () => {
    expect(classifyTaskMode("Insert a table here")).toBe(TASK_MODES.COMPLEX_AGENT);
    expect(classifyTaskMode("What is the weather today?")).toBe(TASK_MODES.COMPLEX_AGENT);
  });

  it('prioritizes specific modes over general terms', () => {
    // "rewrite" should hit rewrite_selection, not complex_agent
    expect(classifyTaskMode("Rewrite this section")).toBe(TASK_MODES.REWRITE_SELECTION);
    // "format" should hit format_selection
    expect(classifyTaskMode("Format the selected text")).toBe(TASK_MODES.FORMAT_SELECTION);
  });

  it('is case-insensitive', () => {
    expect(classifyTaskMode("PROOFREAD this document")).toBe(TASK_MODES.PROOFREAD_REVIEW);
    expect(classifyTaskMode("SUMMARIZE the content")).toBe(TASK_MODES.SUMMARIZE);
  });

  it('handles empty string as complex_agent', () => {
    expect(classifyTaskMode("")).toBe(TASK_MODES.COMPLEX_AGENT);
  });
});

describe("getIterationBudget", () => {
  it('returns bounded budgets for routine modes', () => {
    expect(getIterationBudget(TASK_MODES.SUMMARIZE, 15)).toBe(2);
    expect(getIterationBudget(TASK_MODES.REWRITE_SELECTION, 15)).toBe(3);
    expect(getIterationBudget(TASK_MODES.FORMAT_SELECTION, 15)).toBe(3);
    expect(getIterationBudget(TASK_MODES.PROOFREAD_REVIEW, 15)).toBe(2);
    expect(getIterationBudget(PROOFREAD_ANNOTATE_MODE, 15)).toBe(2);
    expect(getIterationBudget(TASK_MODES.PROOFREAD_EDIT, 15)).toBe(4);
  });

  it('caps routine budgets below the user maximum', () => {
    // Even with maxIterations=50, routine modes stay bounded
    expect(getIterationBudget(TASK_MODES.SUMMARIZE, 50)).toBe(2);
    expect(getIterationBudget(TASK_MODES.PROOFREAD_REVIEW, 50)).toBe(2);
    expect(getIterationBudget(TASK_MODES.PROOFREAD_EDIT, 50)).toBe(4);
  });

  it('returns configured maximum for complex_agent', () => {
    expect(getIterationBudget(TASK_MODES.COMPLEX_AGENT, 15)).toBe(15);
    expect(getIterationBudget(TASK_MODES.COMPLEX_AGENT, 8)).toBe(8);
    expect(getIterationBudget(TASK_MODES.COMPLEX_AGENT, 50)).toBe(50);
  });

  it('returns configured maximum for unknown modes', () => {
    expect((getIterationBudget("unknown_mode" as any, 10) as number)).toBe(10);
  });

  it('respects the floor of maxIterations=1', () => {
    // Routine budgets should never exceed maxIterations if it's very low
    expect(getIterationBudget(TASK_MODES.SUMMARIZE, 1)).toBe(1);
    expect(getIterationBudget(TASK_MODES.PROOFREAD_REVIEW, 1)).toBe(1);
    expect(getIterationBudget(TASK_MODES.PROOFREAD_EDIT, 2)).toBe(2);
  });
});
