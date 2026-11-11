import { afterEach, describe, expect, it, vi } from "vitest";
import { executePowerPointTool } from "../powerPointTools";

function loadedValue<T>(value: T, label: string) {
  let loaded = false;
  return {
    get value(): T {
      if (!loaded) throw Object.assign(new Error(`${label} was not loaded`), { code: "PropertyNotLoaded" });
      return value;
    },
    markLoaded: () => { loaded = true; }
  };
}

function installStrictOfficeMock() {
  const pending: Array<() => void> = [];
  const slideId = loadedValue("slide-0", "slide.id");
  const slideIndex = loadedValue(0, "slide.index");
  const shapeId = loadedValue("shape-1", "shape.id");
  const deleteSlide = vi.fn();
  const addTextBox = vi.fn();
  const addGeometricShape = vi.fn();

  const createdShape = {
    get id() { return shapeId.value; },
    load: vi.fn((props: string) => { if (props.includes("id")) pending.push(shapeId.markLoaded); })
  };
  addTextBox.mockReturnValue(createdShape);
  addGeometricShape.mockReturnValue(createdShape);

  let frameSequence = 0;
  const contentShapeId = loadedValue("content-shape", "contentShape.id");
  const contentShape = {
    get id() { return contentShapeId.value; },
    getTextFrameOrNullObject: vi.fn(() => {
      frameSequence += 1;
      const hasText = loadedValue(true, `textFrame${frameSequence}.hasText`);
      const isNullObject = loadedValue(false, `textFrame${frameSequence}.isNullObject`);
      const text = loadedValue("Meal prep", `textFrame${frameSequence}.textRange.text`);
      const textRange = {
        get text() { return text.value; },
        load: vi.fn((props: string) => { if (props.includes("text")) pending.push(text.markLoaded); })
      };
      return {
        get hasText() { return hasText.value; },
        get isNullObject() { return isNullObject.value; },
        get textRange() { return textRange; },
        load: vi.fn((props: string) => {
          if (props.includes("hasText")) pending.push(hasText.markLoaded);
          if (props.includes("isNullObject")) pending.push(isNullObject.markLoaded);
        })
      };
    })
  };

  const shapes = {
    items: [contentShape],
    load: vi.fn((props: string) => { if (props.includes("items/id")) pending.push(contentShapeId.markLoaded); }),
    addTextBox,
    addGeometricShape,
    getItem: vi.fn()
  };
  const slide = {
    get id() { return slideId.value; },
    get index() { return slideIndex.value; },
    shapes,
    delete: deleteSlide,
    load: vi.fn((props: string) => {
      if (props.includes("id")) pending.push(slideId.markLoaded);
      if (props.includes("index")) pending.push(slideIndex.markLoaded);
    })
  };
  const slides = {
    items: [slide],
    load: vi.fn(),
    getCount: vi.fn(() => ({ value: 2 })),
    getItemAt: vi.fn(() => slide)
  };
  const context = {
    presentation: {
      slides,
      getSelectedSlides: vi.fn(() => ({ items: [], load: vi.fn() })),
      slideMasters: { items: [], load: vi.fn() }
    },
    sync: vi.fn(async () => { pending.splice(0).forEach((load) => load()); })
  };
  vi.stubGlobal("PowerPoint", { run: vi.fn(async (work: (ctx: typeof context) => Promise<unknown>) => work(context)) });
  return { slide, createdShape, deleteSlide, addTextBox, addGeometricShape, contentShape };
}

afterEach(() => vi.unstubAllGlobals());

describe("live Office.js property loading", () => {
  it("loads slide and created shape identities before returning add_text_box", async () => {
    const mock = installStrictOfficeMock();
    const result = JSON.parse(await executePowerPointTool("add_text_box", '{"slideIndex":0,"text":"Breakfast"}'));
    expect(result).toMatchObject({ ok: true, slideId: "slide-0", slideIndex: 0, shapeId: "shape-1" });
    expect(mock.slide.load).toHaveBeenCalledWith("id,index");
    expect(mock.createdShape.load).toHaveBeenCalledWith("id");
  });

  it("loads created shape identity before returning add_shape", async () => {
    const mock = installStrictOfficeMock();
    const result = JSON.parse(await executePowerPointTool("add_shape", '{"slideIndex":0,"shapeType":"Rectangle"}'));
    expect(result).toMatchObject({ ok: true, slideId: "slide-0", slideIndex: 0, shapeId: "shape-1" });
    expect(mock.createdShape.load).toHaveBeenCalledWith("id");
  });

  it("loads slide identity before delete_slide reads the index", async () => {
    const mock = installStrictOfficeMock();
    const result = JSON.parse(await executePowerPointTool("delete_slide", '{"slideIndex":0}'));
    expect(result).toMatchObject({ ok: true, slideIndex: 0 });
    expect(mock.deleteSlide).toHaveBeenCalledOnce();
  });

  it("reuses the loaded text-frame proxies when reading slide content", async () => {
    const mock = installStrictOfficeMock();
    const result = JSON.parse(await executePowerPointTool("read_slide_content", '{"slideIndex":0}'));
    expect(result).toMatchObject({ ok: true, textBearingShapes: 1 });
    expect(result.shapes[0]).toMatchObject({ shapeId: "content-shape", text: "Meal prep" });
    expect(mock.contentShape.getTextFrameOrNullObject).toHaveBeenCalledOnce();
  });
});
