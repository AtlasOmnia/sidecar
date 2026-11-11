import { describe, expect, it } from "vitest";
import { validateApplyEditPlan } from "../wordTools";

describe("apply_edit_plan validator", () => {
  it("validates a minimal valid plan", () => {
    const result = validateApplyEditPlan({
      edits: [
        {
          find: "foo",
          replace: "bar"
        }
      ]
    });

    expect(result).toEqual({
      ok: true,
      plan: {
        edits: [
          {
            find: "foo",
            replace: "bar",
            matchCase: false,
            matchWholeWord: false
          }
        ]
      }
    });
  });

  it("enforces 1-100 edits", () => {
    expect(validateApplyEditPlan({ edits: [] })).toMatchObject({
      ok: false,
      error: "apply_edit_plan must include 1-100 edits."
    });

    const tooMany = { edits: Array.from({ length: 101 }, () => ({ find: "x", replace: "y" })) };
    expect(validateApplyEditPlan(tooMany)).toMatchObject({ ok: false });
  });

  it("requires object payload", () => {
    expect(validateApplyEditPlan(null)).toMatchObject({ ok: false });
    expect(validateApplyEditPlan([{}])).toMatchObject({ ok: false });
  });

  it("requires edits array", () => {
    expect(validateApplyEditPlan({ edits: "no" as unknown as string[] })).toMatchObject({ ok: false });
  });

  it("rejects empty find and identical find/replace", () => {
    expect(validateApplyEditPlan({ edits: [{ find: "", replace: "x" }] })).toMatchObject({
      ok: false,
      error: "Edit 0 is missing a non-empty find value."
    });

    expect(validateApplyEditPlan({ edits: [{ find: "same", replace: "same" }] })).toMatchObject({
      ok: false,
      error: "Edit 0 must have different find and replace text."
    });
  });

  it("rejects non-string find/replace", () => {
    expect(validateApplyEditPlan({ edits: [{ find: 1 as unknown as string, replace: "x" }] })).toMatchObject({
      ok: false,
      error: "Edit 0 is missing a non-empty find value."
    });

    expect(
      validateApplyEditPlan({ edits: [{ find: "x", replace: 2 as unknown as string }] })
    ).toMatchObject({
      ok: false,
      error: "Edit 0 replace must be a string."
    });
  });

  it("accepts and preserves explicit boolean defaults", () => {
    const result = validateApplyEditPlan({
      edits: [
        {
          find: "foo",
          replace: "bar",
          matchCase: true,
          matchWholeWord: true,
          reason: "typo fix",
          highlightColor: "yellow"
        }
      ]
    });

    expect(result).toMatchObject({
      ok: true,
      plan: {
        edits: [
          {
            find: "foo",
            replace: "bar",
            matchCase: true,
            matchWholeWord: true,
            reason: "typo fix",
            highlightColor: "yellow"
          }
        ]
      }
    });
  });

  it("trims reason and rejects too long fields", () => {
    const result = validateApplyEditPlan({
      edits: [{
        find: "foo",
        replace: "bar",
        reason: "   short reason   "
      }]
    });

    expect(result).toMatchObject({
      ok: true,
      plan: {
        edits: [{
          reason: "short reason"
        }]
      }
    });

    expect(
      validateApplyEditPlan({
        edits: [{
          find: "foo",
          replace: "bar",
          reason: "x".repeat(300)
        }]
      })
    ).toMatchObject({ ok: false, error: "Edit 0 reason exceeds 250 characters." });
  });

  it("validates highlight colors", () => {
    expect(validateApplyEditPlan({ edits: [{ find: "foo", replace: "bar", highlightColor: "yellow" }] })).toMatchObject({
      ok: true
    });
    expect(validateApplyEditPlan({ edits: [{ find: "foo", replace: "bar", highlightColor: "#ffaa00" }] })).toMatchObject({
      ok: true
    });
    expect(validateApplyEditPlan({ edits: [{ find: "foo", replace: "bar", highlightColor: "chartreuse" }] })).toMatchObject({
      ok: false,
      error: "Edit 0 uses an unsupported highlight color."
    });
  });

  it.each([
    ["matchCase", "Edit 0 matchCase must be a boolean."],
    ["matchWholeWord", "Edit 0 matchWholeWord must be a boolean."]
  ])("rejects non-boolean %s instead of silently defaulting it", (field, error) => {
    expect(validateApplyEditPlan({
      edits: [{ find: "foo", replace: "bar", [field]: "true" }]
    })).toEqual({ ok: false, error });
  });
});
