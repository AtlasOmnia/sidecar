import { afterEach, describe, expect, it, vi } from "vitest";
import { executePowerPointTool, executePowerPointToolParsed, getPresentationContext, toolSchemas } from "../powerPointTools";

// Per-proxy load state changes only at sync; repeated access creates a fresh
// proxy where appropriate. These are offline Office.js contract fixtures.
function installFixture(text = "abcdef", nullTarget = false) {
  const pending: Array<() => void> = [];
  const writes: string[] = [];
  function property<T>(value: T, label: string) {
    let loaded = false;
    return {
      get value() {
        if (!loaded) throw Object.assign(new Error(`${label} not loaded`), { code: "PropertyNotLoaded" });
        return value;
      },
      ready: () => { loaded = true; }
    };
  }
  function range(value: string) {
    const content = property(value, "text");
    return {
      get text() { return content.value; },
      set text(value: string) { writes.push(value); },
      isNullObject: false,
      load: vi.fn((props: string) => { if (props.includes("text")) pending.push(content.ready); })
    };
  }
  const target = {
    get id(): string { throw new Error("target shape.id not loaded"); },
    getTextFrameOrNullObject: vi.fn(() => {
      const hasText = property(!nullTarget, "hasText");
      const isNullObject = property(nullTarget, "isNullObject");
      const textRange = range("old old");
      return {
        get hasText() { return hasText.value; },
        get isNullObject() { return isNullObject.value; },
        get textRange() {
          if (isNullObject.value) throw new Error("null text frame accessed");
          return textRange;
        },
        load: vi.fn((props: string) => {
          if (props.includes("hasText")) pending.push(hasText.ready);
          if (props.includes("isNullObject")) pending.push(isNullObject.ready);
        })
      };
    })
  };
  const shapeId = property("shape-text", "shape.id");
  const contentShape = {
    get id() { return shapeId.value; },
    getTextFrameOrNullObject: target.getTextFrameOrNullObject
  };
  const createdShapes: Array<{ readonly id: string; load: ReturnType<typeof vi.fn> }> = [];
  const shapes = {
    items: [contentShape],
    load: vi.fn((props: string) => { if (props.includes("items/id")) pending.push(shapeId.ready); }),
    getItem: vi.fn(() => target),
    addTextBox: vi.fn(() => {
      const id = property(`created-${createdShapes.length}`, "created shape.id");
      const shape = {
        get id() { return id.value; },
        load: vi.fn((props: string) => { if (props.includes("id")) pending.push(id.ready); })
      };
      createdShapes.push(shape);
      return shape;
    })
  };
  const slide = { id: "slide-0", index: 0, shapes, load: vi.fn() };
  const createdSlide = { ...slide, id: "slide-new", index: 1 };
  const slides = {
    items: [slide],
    load: vi.fn(),
    getCount: vi.fn(() => ({ value: slides.items.length })),
    getItemAt: vi.fn(() => slide),
    getItem: vi.fn(() => createdSlide),
    add: vi.fn(() => { slides.items.push(createdSlide); })
  };
  const selectedRange = range(text);
  const title = property("  ", "presentation.title");
  const context = {
    presentation: {
      get title() { return title.value; },
      load: vi.fn((props: string) => { if (props.includes("title")) pending.push(title.ready); }),
      slides,
      getSelectedSlides: vi.fn(() => ({ items: [slide], load: vi.fn() })),
      getSelectedTextRangeOrNullObject: vi.fn(() => selectedRange)
    },
    sync: vi.fn(async () => { pending.splice(0).forEach((load) => load()); })
  };
  const run = vi.fn(async (work: (ctx: typeof context) => Promise<unknown>) => work(context));
  vi.stubGlobal("PowerPoint", { run });
  return { run, writes, target, createdShapes, shapes, selectedRange };
}

afterEach(() => vi.unstubAllGlobals());

describe("dependency-neutral overview length contract", () => {
  it("publishes the same integer bounds enforced by execution", () => {
    const schema = toolSchemas.find((tool) => tool.function.name === "get_presentation_overview");
    const properties = schema?.function.parameters.properties as Record<string, unknown> | undefined;
    expect(properties?.maxLength).toMatchObject({ type: "integer", minimum: 1, maximum: 20000 });
  });

  it.each([1, 4, 20000])("honors maxLength %s with exact metadata", async (maxLength) => {
    installFixture("x".repeat(20001));
    const result = JSON.parse(await executePowerPointTool("get_presentation_overview", JSON.stringify({ maxLength })));
    expect(result).toMatchObject({
      ok: true, title: "Untitled Presentation", selectedCount: 1, slideCount: 1,
      selectedText: "x".repeat(maxLength), selectedTextOriginalLength: 20001,
      selectedTextReturnedLength: maxLength, selectedTextTruncated: true
    });
    expect(result.selectedSlides).toEqual([{ id: "slide-0", index: 0 }]);
  });

  it("keeps default 5000 characters for the tool and presentation context", async () => {
    installFixture("x".repeat(5001));
    for (const result of [
      JSON.parse(await executePowerPointTool("get_presentation_overview", "{}")),
      JSON.parse(await getPresentationContext())
    ]) {
      expect(result).toMatchObject({ ok: true, selectedTextReturnedLength: 5000, selectedTextTruncated: true });
    }
  });

  it.each(["abcdef", "", "abc"])("does not truncate short or empty selection %j", async (text) => {
    installFixture(text);
    const result = JSON.parse(await executePowerPointTool("get_presentation_overview", '{"maxLength":6}'));
    expect(result).toMatchObject({ ok: true, selectedText: text, selectedTextTruncated: false, selectedTextReturnedLength: text.length });
  });

  it("returns empty text for a null selection range", async () => {
    const fixture = installFixture();
    fixture.selectedRange.isNullObject = true;
    const result = JSON.parse(await executePowerPointTool("get_presentation_overview", '{"maxLength":1}'));
    expect(result).toMatchObject({ ok: true, selectedText: "", selectedTextReturnedLength: 0, selectedTextTruncated: false });
  });

  it.each([0, -1, 1.5, 20001, "3", null, Number.NaN, Number.POSITIVE_INFINITY])(
    "rejects invalid parsed maxLength %j before Office runs", async (maxLength) => {
      const fixture = installFixture();
      const result = JSON.parse(await executePowerPointToolParsed("get_presentation_overview", { maxLength }));
      expect(result).toMatchObject({ ok: false, category: "validation" });
      expect(fixture.run).not.toHaveBeenCalled();
    }
  );

  it("rejects fractional maxLength on the JSON entrypoint", async () => {
    const fixture = installFixture();
    const result = JSON.parse(await executePowerPointTool("get_presentation_overview", '{"maxLength":2.5}'));
    expect(result).toMatchObject({ ok: false, category: "validation" });
    expect(fixture.run).not.toHaveBeenCalled();
  });
});

describe("dependency-neutral property loading", () => {
  it.each([undefined, "Body text"])("loads each created text-shape id with optional body %j", async (body) => {
    const fixture = installFixture();
    const result = JSON.parse(await executePowerPointTool("create_text_slide", JSON.stringify({ title: "Title", body })));
    expect(result).toMatchObject({ ok: true, newSlideIndex: 1, titleShapeId: "created-0", bodyShapeId: body ? "created-1" : null });
    fixture.createdShapes.forEach((shape) => expect(shape.load).toHaveBeenCalledWith("id"));
  });

  it("returns the requested target id without an unloaded shape.id read", async () => {
    const fixture = installFixture();
    const result = JSON.parse(await executePowerPointTool("update_slide_text", '{"slideIndex":0,"shapeId":"target","text":"new"}'));
    expect(result).toMatchObject({ ok: true, shapeId: "target", beforeLength: 7, afterLength: 3 });
    expect(fixture.writes).toEqual(["new"]);
  });

  it("retains the exact loaded frame/range during all-shape replacement", async () => {
    const fixture = installFixture();
    const result = JSON.parse(await executePowerPointTool("find_and_replace_text", '{"slideIndex":0,"findText":"old","replaceText":"new"}'));
    expect(result).toMatchObject({ ok: true, totalReplacements: 2, updatedShapes: [{ shapeId: "shape-text", replacements: 2 }] });
    expect(fixture.target.getTextFrameOrNullObject).toHaveBeenCalledOnce();
    expect(fixture.writes).toEqual(["new new"]);
  });

  it("loads null state and skips a null frame without mutation", async () => {
    const fixture = installFixture("", true);
    const result = JSON.parse(await executePowerPointTool("find_and_replace_text", '{"slideIndex":0,"findText":"old","replaceText":"new"}'));
    expect(result).toMatchObject({ ok: true, totalReplacements: 0, updatedShapes: [] });
    expect(fixture.target.getTextFrameOrNullObject).toHaveBeenCalledOnce();
    expect(fixture.writes).toEqual([]);
  });

  it.each([false, true])("loads null state for a targeted frame (null=%s)", async (nullFrame) => {
    const fixture = installFixture("", nullFrame);
    const result = JSON.parse(await executePowerPointTool("find_and_replace_text", '{"slideIndex":0,"shapeId":"target","findText":"old","replaceText":"new"}'));
    expect(result.ok).toBe(!nullFrame);
    if (nullFrame) expect(result.error).toBe("Specified shape exists but has no text frame.");
    expect(fixture.writes).toEqual(nullFrame ? [] : ["new new"]);
  });
});
