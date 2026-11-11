import { describe, expect, it } from "vitest";
import { validateToolCall } from "../toolValidation";

describe("recursive Word tool schema validation", () => {
  it("rejects unknown nested edit fields with an indexed path", () => {
    const result = validateToolCall("apply_edit_plan", JSON.stringify({
      edits: [{ find: "a", replace: "b", overwriteAll: true }],
    }));

    expect(result).toMatchObject({ ok: false, category: "validation" });
    if ("error" in result) expect(result.error).toContain("edits[0].overwriteAll");
  });

  it("validates nested array objects and root required fields", () => {
    expect(validateToolCall("apply_edit_plan", JSON.stringify({ edits: [{ find: "", replace: "b" }] }))).toMatchObject({ ok: false });
    expect(validateToolCall("apply_edit_plan", JSON.stringify({}))).toMatchObject({ ok: false });
    expect(validateToolCall("apply_edit_plan", JSON.stringify({ edits: [{ find: "a", replace: "b", matchCase: true }] }))).toMatchObject({ valid: true });
  });

  it("enforces numeric bounds, integer distinction, arrays, and object property counts", () => {
    expect(validateToolCall("read_document_text", JSON.stringify({ maxLength: 20_001 }))).toMatchObject({ ok: false });
    expect(validateToolCall("read_document_text", JSON.stringify({ maxLength: 1.5 }))).toMatchObject({ ok: false });
    expect(validateToolCall("format_selection", JSON.stringify({}))).toMatchObject({ ok: false });
    expect(validateToolCall("format_selection", JSON.stringify({ fontSize: 0 }))).toMatchObject({ ok: false });
    expect(validateToolCall("insert_text", JSON.stringify({ text: "x", location: "overwrite_all" }))).toMatchObject({ ok: false });
  });

  it("rejects unknown root fields before Word execution", () => {
    expect(validateToolCall("insert_text", JSON.stringify({ text: "x", mode: "overwrite_all" })))
      .toMatchObject({ ok: false, category: "validation" });
  });
});
