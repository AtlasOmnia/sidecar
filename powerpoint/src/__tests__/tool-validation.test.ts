import { describe, expect, it, vi } from "vitest";
import { validateToolArguments } from "../toolValidation";

const schema = {
  type: "object",
  properties: {
    slideIndex: { type: "number", minimum: 0, maximum: 9, integer: true },
    font: { type: "object", properties: { bold: { type: "boolean" } }, additionalProperties: false },
    bullets: { type: "array", items: { type: "string" } },
  },
  required: ["slideIndex", "font"],
  additionalProperties: false,
};

describe("recursive tool argument validation", () => {
  it.each([
    ["not-json", "invalid JSON"],
    ["null", "object"],
    ["[]", "object"],
    ["1", "object"],
    ["{\"slideIndex\":0}", "font"],
    ["{\"slideIndex\":0,\"font\":{},\"secret\":1}", "unknown"],
    ["{\"slideIndex\":0,\"font\":{\"secret\":true}}", "unknown"],
    ["{\"slideIndex\":1.2,\"font\":{}}", "integer"],
    ["{\"slideIndex\":10,\"font\":{}}", "maximum"],
    ["{\"slideIndex\":0,\"font\":{},\"bullets\":[1]}", "string"],
  ])("rejects %s", (raw, message) => {
    const result = validateToolArguments(raw, schema);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.error.toLowerCase()).toContain(message.toLowerCase());
  });

  it("rejects non-finite numbers", () => {
    const result = validateToolArguments(JSON.stringify({ slideIndex: NaN, font: {} }), schema);
    expect(result.ok).toBe(false);
  });

  it("returns the parsed object used for execution", () => {
    const result = validateToolArguments(JSON.stringify({ slideIndex: 2, font: { bold: true }, bullets: ["a"] }), schema);
    expect(result).toEqual({ ok: true, parsedArgs: { slideIndex: 2, font: { bold: true }, bullets: ["a"] } });
  });

  it("reports unknown tools through the optional registry", () => {
    const result = validateToolArguments("{}", undefined, "missing_tool");
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.category).toBe("validation");
  });
});
