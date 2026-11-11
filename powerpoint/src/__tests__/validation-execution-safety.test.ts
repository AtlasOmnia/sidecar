import { afterEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { validateSchemaValue } from "../toolValidation";
import { executePowerPointTool, toolSchemas } from "../powerPointTools";

const indexSchema = { type: "integer", minimum: 0, maximum: 2, enum: [0, 1, 2] };

describe("registered PowerPoint surface", () => {
  it("retains canonical pane handlers, with only 15 advertised plugin tools", () => {
    expect(toolSchemas).toHaveLength(18);
    const gated = new Set(["add_speaker_notes", "duplicate_slide", "add_image"]);
    expect(toolSchemas.filter((tool) => gated.has(tool.function.name))).toHaveLength(3);
    expect(toolSchemas.filter((tool) => !gated.has(tool.function.name))).toHaveLength(15);
  });
});

describe("integer schema validation", () => {
  it("enforces integer minimum, maximum, and enum constraints", () => {
    const boundedInteger = { type: "integer", minimum: 0, maximum: 2 };
    expect(validateSchemaValue(-1, boundedInteger)).toContain("minimum");
    expect(validateSchemaValue(3, boundedInteger)).toContain("maximum");
    expect(validateSchemaValue(1, { ...indexSchema, enum: [0, 2] })).toContain("one of");
    expect(validateSchemaValue(1.5, indexSchema)).toContain("integer");
    expect(validateSchemaValue(Number.NaN, indexSchema)).toContain("integer");
    expect(validateSchemaValue(Number.POSITIVE_INFINITY, indexSchema)).toContain("integer");
  });
});

function installOfficeMock(count = 2) {
  const deleteMethod = vi.fn();
  const moveTo = vi.fn();
  const applyLayout = vi.fn();
  const getItemAt = vi.fn((index: number) => ({
    id: `slide-${index}`,
    index,
    load: vi.fn(),
    delete: deleteMethod,
    moveTo,
    applyLayout,
    shapes: { getItem: vi.fn() }
  }));
  const slides = {
    load: vi.fn(),
    getCount: vi.fn(() => ({ value: count })),
    getItemAt
  };
  const run = vi.fn(async (work: (context: any) => Promise<unknown>) => work({
    presentation: {
      slides,
      getSelectedSlides: vi.fn(() => ({ items: [], load: vi.fn() })),
      slideMasters: { items: [], load: vi.fn() }
    },
    sync: vi.fn(async () => undefined)
  }));
  vi.stubGlobal("PowerPoint", { run });
  return { run, slides, getItemAt, deleteMethod, moveTo, applyLayout };
}

afterEach(() => { vi.unstubAllGlobals(); });

describe("pre-Office slide-index rejection", () => {
  it.each([
    ["add_text_box", '{"text":"x","slideIndex":-1}'],
    ["read_slide_content", '{"slideIndex":1.5}'],
    ["delete_slide", '{"slideIndex":1e999}'],
    ["reorder_slide", '{"fromIndex":-1,"toIndex":0}'],
    ["set_slide_layout", '{"slideIndex":0.5,"layoutId":"layout"}']
  ])("rejects invalid %s before PowerPoint.run", async (name, rawArgs) => {
    const { run, getItemAt, deleteMethod, moveTo, applyLayout } = installOfficeMock();
    const result = JSON.parse(await executePowerPointTool(name, rawArgs));
    expect(result.ok).toBe(false);
    expect(result.category).toBe("validation");
    expect(run).not.toHaveBeenCalled();
    expect(getItemAt).not.toHaveBeenCalled();
    expect(deleteMethod).not.toHaveBeenCalled();
    expect(moveTo).not.toHaveBeenCalled();
    expect(applyLayout).not.toHaveBeenCalled();
  });
});

describe("realistic Office slide-index bounds", () => {
  it("rejects count and beyond-count indices before mutation methods", async () => {
    const cases: Array<[string, string]> = [
      ["delete_slide", '{"slideIndex":2}'],
      ["reorder_slide", '{"fromIndex":2,"toIndex":0}'],
      ["set_slide_layout", '{"slideIndex":2,"layoutId":"layout"}']
    ];
    for (const [name, rawArgs] of cases) {
      const { getItemAt, deleteMethod, moveTo, applyLayout } = installOfficeMock(2);
      const result = JSON.parse(await executePowerPointTool(name, rawArgs));
      expect(result.ok).toBe(false);
      expect(result.category).toBe("validation");
      expect(getItemAt).not.toHaveBeenCalled();
      expect(deleteMethod).not.toHaveBeenCalled();
      expect(moveTo).not.toHaveBeenCalled();
      expect(applyLayout).not.toHaveBeenCalled();
      vi.unstubAllGlobals();
    }
  });
});

describe("typed serialized tool failures", () => {
  it.each(["add_speaker_notes", "duplicate_slide", "add_image"])("adds office_host category to gated %s result and preserves requiredApi", async (name) => {
    const args = name === "add_speaker_notes" ? { text: "notes" } : name === "duplicate_slide" ? { slideIndex: 0 } : { base64Image: "dGVzdA==", slideIndex: 0 };
    const result = JSON.parse(await executePowerPointTool(name, JSON.stringify(args)));
    expect(result).toMatchObject({ ok: false, category: "office_host" });
    expect(result.details.requiredApi).toBeTruthy();
  });

  it("normalizes last-slide guard and Office code failures without dropping details", async () => {
    installOfficeMock(1);
    const lastSlide = JSON.parse(await executePowerPointTool("delete_slide", '{"slideIndex":0}'));
    expect(lastSlide).toMatchObject({ ok: false, category: "office_host" });

    vi.stubGlobal("PowerPoint", { run: vi.fn().mockRejectedValue({ code: "InvalidRequest", message: "host rejected request" }) });
    const officeFailure = JSON.parse(await executePowerPointTool("get_presentation_overview", "{}"));
    expect(officeFailure).toMatchObject({ ok: false, category: "office_host", details: { code: "InvalidRequest" } });
  });
});

describe("release metadata and marker cleanup", () => {
  it("declares Node engine in package metadata and lock root parity", () => {
    const pkg = JSON.parse(readFileSync(resolve(process.cwd(), "package.json"), "utf8"));
    const lock = JSON.parse(readFileSync(resolve(process.cwd(), "package-lock.json"), "utf8"));
    expect(pkg.engines?.node).toBe(">=20.9.0");
    expect(lock.packages[""]?.engines?.node).toBe(pkg.engines.node);
  });

  it("has no stray marker in tool validation", () => {
    const source = readFileSync(resolve(process.cwd(), "src/toolValidation.ts"), "utf8");
    expect(source).not.toContain("//MARKER_ABC");
  });
});
